/**
 * @file main.cpp
 * @brief ESP32-S3 "AEDI" Radar Node — CSI Sensing + LED Visual Interface
 *
 * Full CSI sensing pipeline + ADR-018 UDP streaming + LED visual state machine.
 *
 * On boot:
 *   1. Reads WiFi credentials + aggregator IP/port from NVS (namespace: csi_cfg)
 *   2. Connects to WiFi in station mode
 *   3. Enables promiscuous WiFi CSI collection
 *   4. Streams ADR-018 binary CSI frames to the aggregator via UDP
 *   5. Drives WS2812B LED according to sensing state
 *
 * ADR-018 binary frame layout (must match Rust sensing-server parser):
 *   [0..3]   Magic: 0xC5110001 (LE)
 *   [4]      Node ID (u8, from NVS key "node_id")
 *   [5]      Antenna count (u8, always 1)
 *   [6..7]   Subcarrier count (u16 LE) = rssi_count / 2
 *   [8..11]  Frequency MHz (u32 LE)
 *   [12..15] Sequence number (u32 LE)
 *   [16]     RSSI (i8)
 *   [17]     Noise floor (i8)
 *   [18..19] Reserved (0x00)
 *   [20..]   I/Q bytes (raw from CSI callback)
 *
 * Hardware: ESP32-S3 N16R8 (built-in WS2812B on GPIO 38)
 * Framework: Arduino (PlatformIO)
 * NVS keys (namespace "csi_cfg"): ssid, password, target_ip, target_port,
 *                                   node_id, led_hub
 */

#include <Arduino.h>
#include <FastLED.h>
#include <Preferences.h>
#include <WiFi.h>

// BSD/POSIX socket API via lwIP (ESP-IDF style, avoids WiFiUDP pbuf issues)
#include "lwip/sockets.h"
#include "lwip/netdb.h"

// ESP-IDF CSI API (available in Arduino ESP32 SDK)
#include "esp_wifi.h"
#include "esp_event.h"
#include "freertos/queue.h"

// ESP32-S3 internal temperature sensor (available via Arduino core)
#ifdef __cplusplus
extern "C" {
#endif
float temperatureRead(void);  // returns °C from internal sensor
#ifdef __cplusplus
}
#endif

// ---------------------------------------------------------------------------
// Hardware configuration
// ---------------------------------------------------------------------------
// GOOUUU ESP32-S3 N16R8: built-in WS2812B on GPIO 38 (GRB order)
// ESP32-S3-DevKitC-1 (official Espressif): GPIO 48
// ESP32-S3 SuperMini / some clones:        GPIO 47
//
// DATA_PIN is probed at runtime — firmware tries GPIO48 first, then GPIO38.
// To force a specific pin, set DATA_PIN_OVERRIDE in build_flags.
#define NUM_LEDS     1
#define BRIGHTNESS   255
#define LED_TYPE     WS2812B
#define COLOR_ORDER  GRB

#ifndef DATA_PIN_OVERRIDE
  // We'll try both. Arduino FastLED requires a compile-time constant for the
  // template, so compile for 48; we drive both at runtime to cover all boards.
  // A separate output on pin 38 is driven identically via direct RMT write.
  #define DATA_PIN  48
#else
  #define DATA_PIN  DATA_PIN_OVERRIDE
#endif

// ---------------------------------------------------------------------------
// WiFi / UDP configuration
// All values read from NVS at boot; these are compile-time fallbacks only.
// ---------------------------------------------------------------------------
#define HARDCODED_SSID       ""
#define HARDCODED_PASSWORD   ""
#define DEFAULT_TARGET_IP    "172.23.9.61"  // Hub's IP (same as HUB_IP in .env.local)
#define DEFAULT_TARGET_PORT  5005
#define DEFAULT_NODE_ID      0              // Overridden by NVS node_id

// NVS preference namespace (shared with main ESP-IDF firmware)
#define NVS_NAMESPACE "csi_cfg"

// ---------------------------------------------------------------------------
// ADR-018 frame constants (must match sensing-server Rust parser)  
// ---------------------------------------------------------------------------
#define CSI_MAGIC           0xC5110001UL
#define CSI_HEADER_SIZE     20
#define CSI_MAX_FRAME_SIZE  1500  // fits in one UDP datagram

// Rate-limit: max 50 Hz UDP sends (20 ms minimum between sends)
#define CSI_MIN_SEND_INTERVAL_MS  20

// ---------------------------------------------------------------------------
// ADR-018 telemetry extension (magic 0xC5110003)
// Sent every 10 s alongside the status log, carries chip temperature + stats.
// Format: [magic:4][node_id:1][temp_c_x10:i16][uptime_s:u32][free_heap:u32]
//         [wifi_rssi:i8][csi_frames:u32][udp_sent:u32][udp_fail:u32]
//         [ser_sent:u32][channel:u8][tx_power:i8][reserved:2]  = 32 bytes
// ---------------------------------------------------------------------------
#define TELEMETRY_MAGIC     0xC5110003UL
#define TELEMETRY_PKT_SIZE  32

// --------------------------------------------------------------------------
// Serial bridge framing (SLIP-lite)
// Each ADR-018 frame is wrapped:  0xAB 0xCD [len_hi] [len_lo] [data]
// The Pi-side bridge reads this and forwards raw [data] as UDP to localhost.
// This bypasses AP isolation — data flows USB serial, not WiFi.
// --------------------------------------------------------------------------
#define SLIP_SOF_0  0xAB
#define SLIP_SOF_1  0xCD

// ---------------------------------------------------------------------------
// System states
// ---------------------------------------------------------------------------
enum SystemState : uint8_t {
    BOOT_SWEEP       = 0,
    CALIBRATING      = 1,   // WiFi connecting
    CONN_ESTABLISHED = 2,
    IDLE_SCANNING    = 3,
    HUMAN_DETECTED   = 4,
    AI_PROCESSING    = 5,
    INTERFERENCE     = 6,
    OTA_UPDATE       = 7,
    CRASH_FAULT      = 8
};

// ---------------------------------------------------------------------------
// Global state
// ---------------------------------------------------------------------------
CRGB leds[NUM_LEDS];

SystemState currentState  = BOOT_SWEEP;
bool        isHubNode     = false;  // Deep Violet hub vs Teal edge
int         currentRSSI   = 0;     // 0-100 scale (0 = disconnected)

unsigned long stateStartTime     = 0;
unsigned long lastYellowBeacon   = 0;
unsigned long lastRolePing       = 0;
unsigned long lastWiFiCheck      = 0;
unsigned long lastStatusLog      = 0;

// CSI queue: callback pushes frames, sender task pops + sends UDP
// Queue depth 8 gives ~160 ms buffer at 50 Hz; frames dropped if full (ok).
#define CSI_QUEUE_DEPTH  8

struct CsiQueueItem {
    uint8_t  buf[CSI_MAX_FRAME_SIZE];
    uint16_t len;
};

// CSI / UDP runtime state
static int        g_udp_sock    = -1;   // raw BSD UDP socket
QueueHandle_t     g_csi_queue   = nullptr;
char              g_target_ip[32]  = DEFAULT_TARGET_IP;
uint16_t          g_target_port    = DEFAULT_TARGET_PORT;
uint8_t           g_node_id        = DEFAULT_NODE_ID;
uint32_t          g_seq            = 0;
uint32_t          g_csi_frames     = 0;  // frames received from CSI callback
uint32_t          g_udp_sent       = 0;  // frames sent over UDP
uint32_t          g_ser_sent       = 0;  // frames sent over serial bridge
uint32_t          g_udp_fail       = 0;  // send failures
bool              g_csi_running    = false;
bool              g_serial_bridge  = true;  // always send via serial (AP-isolation bypass)

// Rolling RSSI buffer (last 8 per-frame values)
static int8_t  g_rssi_buf[8]    = {};
static uint8_t g_rssi_idx       = 0;

// ---------------------------------------------------------------------------
// ADR-018 frame serializer
// Build the binary frame into buf[], return bytes written (0 = overflow/skip)
// ---------------------------------------------------------------------------
static size_t serializeADR018(int8_t rssi, int8_t noise, uint8_t ch,
                               const int8_t *iq, uint16_t iq_len,
                               uint8_t *buf, size_t buf_size)
{
    uint16_t n_sub = iq_len / 2;
    size_t frame_len = CSI_HEADER_SIZE + iq_len;
    if (frame_len > buf_size) return 0;

    uint32_t magic   = CSI_MAGIC;
    uint32_t freq    = (ch >= 1 && ch <= 13)  ? (2412 + (ch - 1) * 5) :
                       (ch == 14)             ? 2484 :
                       (ch >= 36 && ch <= 177)? (5000 + ch * 5) : 0;
    uint32_t seq     = g_seq++;

    memcpy(buf,      &magic,   4);
    buf[4]  = g_node_id;
    buf[5]  = 1;              // antenna count
    memcpy(buf + 6,  &n_sub,  2);
    memcpy(buf + 8,  &freq,   4);
    memcpy(buf + 12, &seq,    4);
    buf[16] = (uint8_t)rssi;
    buf[17] = (uint8_t)noise;
    buf[18] = 0;
    buf[19] = 0;
    memcpy(buf + CSI_HEADER_SIZE, iq, iq_len);
    return frame_len;
}

// ---------------------------------------------------------------------------
// WiFi CSI callback — called from WiFi task context (NOT a hardware ISR)
// Serializes to ADR-018 and posts to queue for the sender task (rate-limited)
// ---------------------------------------------------------------------------
static uint32_t s_last_send_ms = 0;

static void csiCallback(void *ctx, wifi_csi_info_t *info)
{
    (void)ctx;
    if (!info || !info->buf || info->len == 0) return;

    g_csi_frames++;

    // Track rolling RSSI for LED state
    g_rssi_buf[g_rssi_idx & 7] = info->rx_ctrl.rssi;
    g_rssi_idx++;

    // Rate-limit to 50 Hz
    uint32_t now_ms = (uint32_t)(esp_timer_get_time() / 1000ULL);
    if ((now_ms - s_last_send_ms) < (uint32_t)CSI_MIN_SEND_INTERVAL_MS) return;
    s_last_send_ms = now_ms;

    if (!g_csi_queue) return;

    // Serialize into a stack-allocated queue item (stack is fine in WiFi task)
    CsiQueueItem item;
    item.len = (uint16_t)serializeADR018(
        info->rx_ctrl.rssi,
        info->rx_ctrl.noise_floor,
        info->rx_ctrl.channel,
        (const int8_t *)info->buf,
        (uint16_t)min((int)info->len, (int)(CSI_MAX_FRAME_SIZE - CSI_HEADER_SIZE)),
        item.buf,
        sizeof(item.buf)
    );
    if (item.len == 0) return;

    // Non-blocking enqueue; drop frame silently if queue is full
    xQueueSend(g_csi_queue, &item, 0);
}

// ---------------------------------------------------------------------------
// FreeRTOS sender task — dequeues frames, sends via:
//   1. BSD UDP socket   (works if AP isolation is disabled)
//   2. Serial SLIP frame (always works — Pi bridge reads and forwards to UDP)
// ---------------------------------------------------------------------------
static void csiSenderTask(void *pv)
{
    // Build the sockaddr_in for the target once
    struct sockaddr_in dest = {};
    dest.sin_family = AF_INET;
    dest.sin_port   = htons(g_target_port);
    inet_pton(AF_INET, g_target_ip, &dest.sin_addr);

    CsiQueueItem item;
    for (;;) {
        if (xQueueReceive(g_csi_queue, &item, portMAX_DELAY) == pdTRUE) {

            // 1. UDP send (may silently fail with AP isolation — that's ok)
            if (g_udp_sock >= 0) {
                int r = sendto(g_udp_sock, item.buf, item.len, 0,
                               (struct sockaddr *)&dest, sizeof(dest));
                if (r > 0) g_udp_sent++;
                else       g_udp_fail++;
            }

            // 2. Serial SLIP-lite bridge (bypasses AP isolation)
            // Format: [0xAB][0xCD][len_hi][len_lo][data...]
            // Pi bridge reads this and injects as UDP to localhost:5005
            if (g_serial_bridge) {
                uint8_t hdr[4] = {
                    SLIP_SOF_0,
                    SLIP_SOF_1,
                    (uint8_t)(item.len >> 8),
                    (uint8_t)(item.len & 0xFF)
                };
                Serial.write(hdr, 4);
                Serial.write(item.buf, item.len);
                g_ser_sent++;
            }
        }
    }
}

// ---------------------------------------------------------------------------
// Enable WiFi CSI collection (called after WiFi STA connects)
// ---------------------------------------------------------------------------
static void startCSI()
{
    if (g_csi_running) return;

    wifi_csi_config_t cfg = {};
    cfg.lltf_en           = true;
    cfg.htltf_en          = true;
    cfg.stbc_htltf2_en    = true;
    cfg.ltf_merge_en      = true;
    cfg.channel_filter_en = false;
    cfg.manu_scale        = false;
    cfg.shift             = 0;

    esp_wifi_set_csi_config(&cfg);
    esp_wifi_set_csi_rx_cb(csiCallback, nullptr);
    esp_wifi_set_csi(true);

    // Open raw UDP socket (SO_SNDBUF tuned to 6 datagrams x ~1500 bytes)
    g_udp_sock = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP);
    if (g_udp_sock < 0) {
        Serial.printf("[CSI] socket() failed: %d\n", errno);
    } else {
        int sndbuf = 9000;
        setsockopt(g_udp_sock, SOL_SOCKET, SO_SNDBUF, &sndbuf, sizeof(sndbuf));
    }

    // Start the UDP sender task on core 0
    xTaskCreatePinnedToCore(
        csiSenderTask, "csi_udp",
        4096,      // stack bytes
        nullptr,   // arg
        5,         // priority
        nullptr,   // handle not needed
        0          // core 0
    );

    g_csi_running = true;
    Serial.printf("[CSI] Started — streaming to %s:%u  node_id=%u\n",
                  g_target_ip, g_target_port, g_node_id);
}

Preferences prefs;

// ---------------------------------------------------------------------------
// Forward declarations
// ---------------------------------------------------------------------------
void        setSystemState(SystemState s);
void        manageSystemLogic(unsigned long ms);
void        startCSI(void);
CRGB        renderBootSweep(unsigned long ms);
CRGB        renderCalibrating(unsigned long ms);
CRGB        renderConnectionEstablished(unsigned long ms);
CRGB        renderIdleScanning(unsigned long ms);
CRGB        renderHumanDetected(unsigned long ms);
CRGB        renderAIProcessing(unsigned long ms);
CRGB        renderInterference(unsigned long ms);
CRGB        renderOTAUpdate(unsigned long ms);
CRGB        renderCrashFault(unsigned long ms);
CRGB        renderYellowBeacon(unsigned long elapsed, int rssi);
CRGB        renderRolePing(unsigned long elapsed, bool isHub);

// ---------------------------------------------------------------------------
// setup()
// ---------------------------------------------------------------------------
void setup() {
    Serial.begin(460800);
    delay(100);

    // FastLED init — drive BOTH GPIO48 and GPIO38 so firmware works on all
    // ESP32-S3 board variants without needing to determine pin at compile time.
    // Only one will have a real LED; the other is harmless.
    FastLED.addLeds<LED_TYPE, 48, COLOR_ORDER>(leds, NUM_LEDS)
           .setCorrection(TypicalLEDStrip);
    FastLED.addLeds<LED_TYPE, 38, COLOR_ORDER>(leds, NUM_LEDS)
           .setCorrection(TypicalLEDStrip);
    FastLED.setBrightness(BRIGHTNESS);
    leds[0] = CRGB::Black;
    FastLED.show();

    // Read all config from NVS
    prefs.begin(NVS_NAMESPACE, /* readOnly= */ true);
    isHubNode    = (prefs.getUChar("led_hub",  0) != 0);
    g_node_id    = prefs.getUChar("node_id",   DEFAULT_NODE_ID);
    g_target_port= prefs.getUShort("target_port", DEFAULT_TARGET_PORT);

    char ssid[33]  = HARDCODED_SSID;
    char pass[65]  = HARDCODED_PASSWORD;
    char tip[32]   = DEFAULT_TARGET_IP;
    prefs.getString("ssid",      ssid, sizeof(ssid));
    prefs.getString("password",  pass, sizeof(pass));
    prefs.getString("target_ip", tip,  sizeof(tip));
    prefs.end();

    strncpy(g_target_ip, tip, sizeof(g_target_ip) - 1);

    // Create the CSI frame queue
    g_csi_queue = xQueueCreate(CSI_QUEUE_DEPTH, sizeof(CsiQueueItem));

    // WiFi STA mode
    WiFi.mode(WIFI_STA);
    if (strlen(ssid) > 0) {
        WiFi.begin(ssid, pass);
        Serial.printf("[CSI] Connecting to SSID: %s\n", ssid);
    } else {
        Serial.println("[CSI] No SSID — skipping WiFi. Provision with provision.py.");
    }

    Serial.printf("[CSI] Node ID: %u  role: %s  target: %s:%u\n",
                  g_node_id,
                  isHubNode ? "hub" : "edge",
                  g_target_ip, g_target_port);

    stateStartTime = millis();
}

// ---------------------------------------------------------------------------
// loop()
// ---------------------------------------------------------------------------
void loop() {
    unsigned long now = millis();
    CRGB baseColor = CRGB::Black;

    // 1. System logic (WiFi RSSI, state transitions, mock radar)
    manageSystemLogic(now);

    // 2. Base state animations
    switch (currentState) {
        case BOOT_SWEEP:        baseColor = renderBootSweep(now);              break;
        case CALIBRATING:       baseColor = renderCalibrating(now);            break;
        case CONN_ESTABLISHED:  baseColor = renderConnectionEstablished(now);  break;
        case IDLE_SCANNING:     baseColor = renderIdleScanning(now);           break;
        case HUMAN_DETECTED:    baseColor = renderHumanDetected(now);          break;
        case AI_PROCESSING:     baseColor = renderAIProcessing(now);           break;
        case INTERFERENCE:      baseColor = renderInterference(now);           break;
        case OTA_UPDATE:        baseColor = renderOTAUpdate(now);              break;
        case CRASH_FAULT:       baseColor = renderCrashFault(now);             break;
    }

    // 3. Overlay: 10-second Yellow Beacon (RSSI / signal strength)
    if (now - lastYellowBeacon >= 10000) {
        lastYellowBeacon = now;
    }
    unsigned long beaconElapsed = now - lastYellowBeacon;
    if (beaconElapsed < 1000 &&
        currentState != CRASH_FAULT &&
        currentState != BOOT_SWEEP  &&
        currentState != OTA_UPDATE  &&
        currentState != CALIBRATING) {
        baseColor = renderYellowBeacon(beaconElapsed, currentRSSI);
    }

    // 4. Overlay: 20-second Role Ping (hub/edge identity), offset +5 s from beacon
    if (now - lastRolePing >= 20000) {
        lastRolePing = now;
    }
    unsigned long rolePingElapsed = now - lastRolePing;
    if (rolePingElapsed > 5000 && rolePingElapsed < 6000 &&
        currentState != CRASH_FAULT &&
        currentState != BOOT_SWEEP  &&
        currentState != OTA_UPDATE  &&
        currentState != CALIBRATING) {
        baseColor = renderRolePing(rolePingElapsed - 5000, isHubNode);
    }

    // 5. Render
    leds[0] = baseColor;
    FastLED.show();
}

// ===========================================================================
// System logic
// ===========================================================================

void manageSystemLogic(unsigned long now) {
    // Check WiFi status every 500 ms
    if (now - lastWiFiCheck >= 500) {
        lastWiFiCheck = now;

        if (WiFi.status() == WL_CONNECTED) {
            // Map raw dBm (-90 … -30) to 1-100 visual scale
            int raw = WiFi.RSSI();
            currentRSSI = constrain((int)map(raw, -90, -30, 1, 100), 1, 100);

            if (currentState == CALIBRATING) {
                Serial.printf("[LED] WiFi connected! RSSI=%d dBm (%d%%)\n", raw, currentRSSI);
                setSystemState(CONN_ESTABLISHED);
            }
        } else {
            currentRSSI = 0;
            if (currentState == IDLE_SCANNING  ||
                currentState == HUMAN_DETECTED  ||
                currentState == AI_PROCESSING) {
                Serial.println("[LED] WiFi lost — showing INTERFERENCE.");
                setSystemState(INTERFERENCE);
            }
        }
    }

    // Start CSI once WiFi connects (only run once)
    if (!g_csi_running && WiFi.status() == WL_CONNECTED) {
        startCSI();
    }

    // LED reflects CSI sensing: treat >0 UDP sends/sec as IDLE_SCANNING
    // No mock events — LED state driven purely by real CSI activity.
    // HUMAN_DETECTED must be driven externally (future: parse server response)

    // Periodic status log + telemetry packet every 10 s
    if (now - lastStatusLog >= 10000) {
        lastStatusLog = now;

        // Read chip temperature (ESP32-S3 internal sensor)
        float chip_temp_c = temperatureRead();

        if (g_csi_running) {
            Serial.printf("[CSI] frames=%lu udp_sent=%lu udp_fail=%lu ser_sent=%lu rssi=%d dBm temp=%.1f°C heap=%lu\n",
                          (unsigned long)g_csi_frames,
                          (unsigned long)g_udp_sent,
                          (unsigned long)g_udp_fail,
                          (unsigned long)g_ser_sent,
                          WiFi.RSSI(),
                          chip_temp_c,
                          (unsigned long)ESP.getFreeHeap());

            // Send telemetry packet (magic 0xC5110003) via UDP + serial bridge
            uint8_t tpkt[TELEMETRY_PKT_SIZE];
            memset(tpkt, 0, sizeof(tpkt));
            uint32_t tmag = TELEMETRY_MAGIC;
            memcpy(tpkt, &tmag, 4);
            tpkt[4] = g_node_id;

            int16_t temp_x10 = (int16_t)(chip_temp_c * 10.0f);
            memcpy(tpkt + 5, &temp_x10, 2);

            uint32_t uptime_s = (uint32_t)(millis() / 1000UL);
            memcpy(tpkt + 7, &uptime_s, 4);

            uint32_t freeheap = (uint32_t)ESP.getFreeHeap();
            memcpy(tpkt + 11, &freeheap, 4);

            tpkt[15] = (uint8_t)WiFi.RSSI();

            uint32_t tmp32;
            tmp32 = g_csi_frames; memcpy(tpkt + 16, &tmp32, 4);
            tmp32 = g_udp_sent;   memcpy(tpkt + 20, &tmp32, 4);
            tmp32 = g_udp_fail;   memcpy(tpkt + 24, &tmp32, 4);
            tmp32 = g_ser_sent;   memcpy(tpkt + 28, &tmp32, 4);

            // Send via UDP (best-effort)
            if (g_udp_sock >= 0) {
                struct sockaddr_in dest = {};
                dest.sin_family = AF_INET;
                dest.sin_port   = htons(g_target_port);
                inet_pton(AF_INET, g_target_ip, &dest.sin_addr);
                sendto(g_udp_sock, tpkt, TELEMETRY_PKT_SIZE, 0,
                       (struct sockaddr *)&dest, sizeof(dest));
            }

            // Also send via serial bridge so RPi always gets it
            if (g_serial_bridge) {
                uint8_t hdr[4] = {
                    SLIP_SOF_0, SLIP_SOF_1,
                    (uint8_t)(TELEMETRY_PKT_SIZE >> 8),
                    (uint8_t)(TELEMETRY_PKT_SIZE & 0xFF)
                };
                Serial.write(hdr, 4);
                Serial.write(tpkt, TELEMETRY_PKT_SIZE);
            }
        }
    }
}

void setSystemState(SystemState s) {
    if (currentState != s) {
        currentState   = s;
        stateStartTime = millis();
    }
}

// ===========================================================================
// Base state renderers
// ===========================================================================

/**
 * Boot Sequence — "Prism Sweep"
 * Rainbow gradient (1.5 s) → White flash (0.5 s) → CALIBRATING
 */
CRGB renderBootSweep(unsigned long now) {
    unsigned long e = now - stateStartTime;
    if (e < 1500) {
        uint8_t hue = (uint8_t)((e * 255UL) / 1500UL);
        return CHSV(hue, 255, 255);
    } else if (e < 2000) {
        return CRGB::White;
    } else {
        setSystemState(CALIBRATING);
        return CRGB::Black;
    }
}

/**
 * Calibrating / WiFi Connecting — expanding white sonar pulse
 */
CRGB renderCalibrating(unsigned long now) {
    unsigned long cycle = now % 2000;
    uint8_t bright = (uint8_t)map((long)cycle, 0L, 1999L, 0L, 255L);
    CRGB c = CRGB::White;
    c.nscale8(bright);
    return c;
}

/**
 * Connection Established — "System Healthy"
 * Emerald Green (1.5 s) → crisp double-blink → IDLE_SCANNING
 */
CRGB renderConnectionEstablished(unsigned long now) {
    unsigned long e = now - stateStartTime;
    const CRGB GREEN = CRGB(0, 255, 50);
    if (e < 1500) {
        return GREEN;
    } else if (e < 1600 || (e > 1700 && e < 1800)) {
        return GREEN;        // blink on
    } else if (e < 2000) {
        return CRGB::Black;  // blink off
    } else {
        setSystemState(IDLE_SCANNING);
        return CRGB::Black;
    }
}

/**
 * Idle / Scanning — "Blue Baseline"
 * Rhythmic deep-blue breathing; BPM can be driven by data throughput.
 */
CRGB renderIdleScanning(unsigned long /*now*/) {
    uint8_t breath = beatsin8(15, 30, 180);
    return CHSV(160, 255, breath);  // 160 = deep blue
}

/**
 * Human Detected — "Contact Pulse"
 * Warm Amber double heartbeat: thump-thump … pause … (1.5 s loop)
 */
CRGB renderHumanDetected(unsigned long now) {
    unsigned long beat = now % 1500;
    const CRGB AMBER = CHSV(32, 255, 255);
    if (beat < 150)       return AMBER;
    else if (beat < 300)  return CRGB::Black;
    else if (beat < 450)  return AMBER;
    else                  return CRGB::Black;
}

/**
 * Heavy AI Processing — "Synapse Spark"
 * Rapid erratic Magenta/Purple flicker.
 */
CRGB renderAIProcessing(unsigned long /*now*/) {
    if (random8() > 100) {
        return CHSV(192, 255, random8(100, 255));
    }
    return CRGB::Black;
}

/**
 * Environmental Interference — "Orange Warning"
 * Slow burnt-orange throb (#FF4500).
 */
CRGB renderInterference(unsigned long /*now*/) {
    uint8_t throb = beatsin8(20, 50, 255);
    CRGB c = CRGB(255, 69, 0);
    c.nscale8(throb);
    return c;
}

/**
 * OTA Update / Matrix Upload — "Do Not Unplug"
 * Hypnotic Cyan ↔ Magenta oscillation (~200 ms per cycle).
 */
CRGB renderOTAUpdate(unsigned long /*now*/) {
    uint8_t wave = beatsin8(150, 0, 255);
    return blend(CRGB::Cyan, CRGB::Magenta, wave);
}

/**
 * Hardware Crash — "Fatal Fault"
 * Chaotic arrhythmic red strobe at maximum brightness.
 */
CRGB renderCrashFault(unsigned long /*now*/) {
    if (random8() > 180) return CRGB::Red;
    return CRGB::Black;
}

// ===========================================================================
// Overlay renderers
// ===========================================================================

/**
 * Yellow Beacon (RSSI signal check, fires every 10 s for 1 s)
 *
 * rssi 70-100 : Sharp aggressive spike → fast fade
 * rssi 30-69  : Lazy swell (50% brightness)
 * rssi  1-29  : Ghostly shimmer (~10% brightness)
 * rssi  0     : Disconnected → dim red double-flash
 */
CRGB renderYellowBeacon(unsigned long elapsed, int rssi) {
    CRGB yellow = CRGB::Yellow;

    if (rssi == 0) {
        if (elapsed < 100 || (elapsed > 200 && elapsed < 300)) return CRGB::Red;
        return CRGB::Black;
    } else if (rssi >= 70) {
        // Strong: hold 200 ms then sharp fade-to-black by 600 ms
        if (elapsed < 200) return yellow;
        if (elapsed < 600) {
            yellow.fadeToBlackBy((uint8_t)map((long)elapsed, 200L, 600L, 0L, 255L));
            return yellow;
        }
        return CRGB::Black;
    } else if (rssi >= 30) {
        // Moderate: lazy swell capped at 50%
        uint8_t wave = sin8((uint8_t)map((long)elapsed, 0L, 1000L, 0L, 255L));
        yellow.nscale8((uint8_t)map(wave, 0, 255, 0, 128));
        return yellow;
    } else {
        // Weak: ghostly ~10% flicker
        if (random8() > 128) { yellow.nscale8(25); return yellow; }
        return CRGB::Black;
    }
}

/**
 * Role Ping (node identity, fires every 20 s for 1 s at +5 s offset)
 *
 * Hub node  : deep-violet double ripple (two fade-in/out cycles)
 * Edge node : crisp singular Teal blink at 400-500 ms mark
 */
CRGB renderRolePing(unsigned long elapsed, bool isHub) {
    if (isHub) {
        CRGB violet = CRGB::DarkViolet;
        if (elapsed < 200)                { violet.nscale8((uint8_t)map((long)elapsed,   0L, 200L,   0L, 255L)); return violet; }
        else if (elapsed < 400)           { violet.nscale8((uint8_t)map((long)elapsed, 200L, 400L, 255L,   0L)); return violet; }
        else if (elapsed < 600)           { violet.nscale8((uint8_t)map((long)elapsed, 400L, 600L,   0L, 255L)); return violet; }
        else if (elapsed < 800)           { violet.nscale8((uint8_t)map((long)elapsed, 600L, 800L, 255L,   0L)); return violet; }
        return CRGB::Black;
    } else {
        if (elapsed > 400 && elapsed < 500) return CRGB::Teal;
        return CRGB::Black;
    }
}
