// ---------------------------------------------------------------------------
// lcd_gc9a01.h — Waveshare ESP32-S3 1.28" round LCD (GC9A01, 240x240) driver.
// Self-contained header — included only when -DENABLE_LCD_GC9A01 is set.
//
// Pin map (Waveshare ESP32-S3-Touch-LCD-1.28 / ESP32-S3-LCD-1.28):
//   SCLK=10  MOSI=11  CS=9  DC=8  RST=14  BL=2
//
// Surfaces four pieces of live state each frame (all local, no server):
//   - Node ID + hub/edge role
//   - WiFi status dot + RSSI (dBm)
//   - CSI packet rate (pkts/s)
//   - Color-coded activity ring, driven by caller-supplied 0..1 "activity"
// ---------------------------------------------------------------------------
#pragma once
#ifdef ENABLE_LCD_GC9A01

#include <Arduino_GFX_Library.h>

#define LCD_SCLK 10
#define LCD_MOSI 11
#define LCD_CS    9
#define LCD_DC    8
#define LCD_RST  14
#define LCD_BL    2

static Arduino_DataBus *lcd_bus = nullptr;
static Arduino_GFX     *lcd     = nullptr;

// Previous values (for dirty-rect updates — avoid redrawing unchanged pixels)
static int   s_prev_rssi       = -999;
static float s_prev_activity   = -1.0f;
static uint8_t s_prev_wifi     = 0xFF;
static uint32_t s_prev_pkt_rate= 0xFFFFFFFF;

// HSV-ish helper: map 0..1 activity to a red→green ring color (cold→warm).
static uint16_t activity_color(float a) {
    if (a < 0.0f) a = 0.0f;
    if (a > 1.0f) a = 1.0f;
    // Low = green, mid = yellow, high = red (classic "presence heatmap")
    uint8_t r = (uint8_t)(255.0f * a);
    uint8_t g = (uint8_t)(255.0f * (1.0f - fabsf(a - 0.5f) * 2.0f));
    uint8_t b = 30;
    return ((r & 0xF8) << 8) | ((g & 0xFC) << 3) | (b >> 3);
}

inline void lcd_init() {
    lcd_bus = new Arduino_ESP32SPI(LCD_DC, LCD_CS, LCD_SCLK, LCD_MOSI, GFX_NOT_DEFINED, HSPI);
    lcd     = new Arduino_GC9A01(lcd_bus, LCD_RST, 0 /* rotation */, true /* IPS */);
    lcd->begin();
    pinMode(LCD_BL, OUTPUT);
    digitalWrite(LCD_BL, HIGH);

    lcd->fillScreen(BLACK);

    // Static header
    lcd->setTextColor(WHITE);
    lcd->setTextSize(2);
    lcd->setCursor(60, 20);
    lcd->print("AEDI-S");
    lcd->setTextSize(1);
    lcd->setTextColor(0x7BEF);      // light grey
    lcd->setCursor(72, 42);
    lcd->print("CSI Sensor Node");
}

// Render or update the display. Call from loop() — internally rate-limited.
// activity: 0..1 local proxy for "something moving" (e.g. amplitude stdev)
// pkt_rate: CSI frames/s observed locally
inline void lcd_update(uint8_t node_id, bool is_hub, int wifi_rssi_dbm,
                       uint32_t csi_frames_per_s, float activity) {
    static uint32_t last_draw = 0;
    uint32_t now = millis();
    if (now - last_draw < 200) return;   // 5 Hz is plenty for human eyes
    last_draw = now;

    if (!lcd) return;

    // --- Activity ring (outer annulus) ---
    uint16_t ring_col = activity_color(activity);
    // Draw ring as two concentric circles filled with black in between when activity changes
    if (fabsf(activity - s_prev_activity) > 0.02f) {
        s_prev_activity = activity;
        for (int r = 112; r < 120; r++) lcd->drawCircle(120, 120, r, ring_col);
    }

    // --- Node ID (big, center) ---
    lcd->fillRect(80, 70, 80, 40, BLACK);
    lcd->setTextColor(is_hub ? 0xFD20 : 0x07FF);   // hub=orange, edge=cyan
    lcd->setTextSize(4);
    lcd->setCursor(node_id < 10 ? 108 : 96, 74);
    lcd->print(node_id);
    lcd->setTextSize(1);
    lcd->setTextColor(0x7BEF);
    lcd->setCursor(is_hub ? 105 : 103, 112);
    lcd->print(is_hub ? "HUB" : "EDGE");

    // --- WiFi status ---
    uint8_t wifi_state = (wifi_rssi_dbm == 0) ? 0 : (wifi_rssi_dbm > -65 ? 2 : 1);
    if (wifi_state != s_prev_wifi || abs(wifi_rssi_dbm - s_prev_rssi) > 2) {
        s_prev_wifi = wifi_state; s_prev_rssi = wifi_rssi_dbm;
        lcd->fillRect(60, 135, 120, 24, BLACK);
        uint16_t wc = (wifi_state == 2) ? 0x07E0 : (wifi_state == 1) ? 0xFFE0 : 0xF800;
        lcd->fillCircle(75, 147, 5, wc);
        lcd->setTextColor(WHITE);
        lcd->setTextSize(1);
        lcd->setCursor(88, 144);
        if (wifi_rssi_dbm == 0) lcd->print("WiFi: OFFLINE");
        else { lcd->print("WiFi: "); lcd->print(wifi_rssi_dbm); lcd->print(" dBm"); }
    }

    // --- Packet rate ---
    if (csi_frames_per_s != s_prev_pkt_rate) {
        s_prev_pkt_rate = csi_frames_per_s;
        lcd->fillRect(60, 165, 120, 18, BLACK);
        lcd->setTextColor(0x7BEF);
        lcd->setTextSize(1);
        lcd->setCursor(78, 168);
        lcd->print("CSI: ");
        lcd->print(csi_frames_per_s);
        lcd->print(" pkt/s");
    }
}

#endif // ENABLE_LCD_GC9A01
