/**
 * @file main.cpp
 * @brief ESP32-S3 "AEDI" Radar Node — LED Visual Interface
 *
 * Single WS2812B/NeoPixel RGB LED state machine implementing the complete
 * AEDI visual language spec (docs/led.md). Communicates radar, network,
 * and system state without a screen using layered animations.
 *
 * Hardware: ESP32-S3 N16R8 (built-in LED on GPIO 48, WS2812B GRB order)
 * Framework: Arduino (PlatformIO)
 * Library:   FastLED ^3.6.0
 *
 * LED pin 48 is the built-in addressable LED on most ESP32-S3 DevKitC-1
 * and compatible modules. Change DATA_PIN if your board differs.
 *
 * Node role (hub vs edge) is read from NVS key "led_hub" in namespace
 * "csi_cfg". Set with:
 *   python firmware/esp32-csi-node/provision.py --led-hub hub
 *   python firmware/esp32-csi-node/provision.py --led-hub edge
 */

#include <Arduino.h>
#include <FastLED.h>
#include <Preferences.h>
#include <WiFi.h>

// ---------------------------------------------------------------------------
// Hardware configuration
// ---------------------------------------------------------------------------
// GOOUUU ESP32-S3 N16R8: built-in WS2812B on GPIO 38
// ESP32-S3-DevKitC-1 (official Espressif): GPIO 48
// ESP32-S3 SuperMini / some clones:        GPIO 47
// Change DATA_PIN to match your board's silkscreen or schematic.
#define NUM_LEDS     1
#define DATA_PIN     38      // GOOUUU N16R8 built-in RGB LED
#define BRIGHTNESS   255
#define LED_TYPE     WS2812B
#define COLOR_ORDER  GRB

// ---------------------------------------------------------------------------
// WiFi credentials — set these or provision via NVS
// Reads from NVS "csi_cfg" namespace (keys: "ssid", "password").
// Falls back to HARDCODED_SSID / HARDCODED_PASSWORD if NVS is empty.
// ---------------------------------------------------------------------------
#define HARDCODED_SSID     ""
#define HARDCODED_PASSWORD ""

// NVS preference namespace (shared with main ESP-IDF firmware)
#define NVS_NAMESPACE "csi_cfg"

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

unsigned long stateStartTime    = 0;
unsigned long lastYellowBeacon  = 0;
unsigned long lastRolePing      = 0;
unsigned long lastWiFiCheck     = 0;
unsigned long lastRadarMockEvent = 0;

Preferences prefs;

// ---------------------------------------------------------------------------
// Forward declarations
// ---------------------------------------------------------------------------
void        setSystemState(SystemState s);
void        manageSystemLogic(unsigned long ms);
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

    // FastLED init
    FastLED.addLeds<LED_TYPE, DATA_PIN, COLOR_ORDER>(leds, NUM_LEDS)
           .setCorrection(TypicalLEDStrip);
    FastLED.setBrightness(BRIGHTNESS);
    leds[0] = CRGB::Black;
    FastLED.show();

    // Read node role from NVS ("led_hub": 1=hub, 0=edge)
    prefs.begin(NVS_NAMESPACE, /* readOnly= */ true);
    isHubNode = (prefs.getUChar("led_hub", 0) != 0);

    // Read WiFi credentials from NVS
    char ssid[33]  = HARDCODED_SSID;
    char pass[65]  = HARDCODED_PASSWORD;
    prefs.getString("ssid",     ssid, sizeof(ssid));
    prefs.getString("password", pass, sizeof(pass));
    prefs.end();

    // Non-blocking WiFi init
    WiFi.mode(WIFI_STA);
    if (strlen(ssid) > 0) {
        WiFi.begin(ssid, pass);
        Serial.printf("[LED] Connecting to SSID: %s\n", ssid);
    } else {
        Serial.println("[LED] No SSID — skipping WiFi. Set with provision.py.");
    }

    Serial.printf("[LED] Node role: %s\n", isHubNode ? "hub (Deep Violet)" : "edge (Teal)");

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

    // Mock radar: randomly trigger HUMAN_DETECTED while idle
    // Replace this block with real CSI event signals in production.
    if (currentState == IDLE_SCANNING &&
        (now - lastRadarMockEvent > 12000)) {
        if (random(10) > 6) {
            Serial.println("[LED] Mock CSI: human detected.");
            setSystemState(HUMAN_DETECTED);
            lastRadarMockEvent = now;
        }
    }

    // Return to idle after 3 s of HUMAN_DETECTED
    if (currentState == HUMAN_DETECTED && (now - stateStartTime > 3000)) {
        Serial.println("[LED] Mock CSI: clearing — return to scan.");
        setSystemState(IDLE_SCANNING);
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
