// ---------------------------------------------------------------------------
// lcd_st7789_240x320.h — Waveshare ESP32-S3-Touch-LCD-2 (ST7789, 240x320 IPS).
//
// Display only (CST816T capacitive touch not used for CSI readout).
// Pin map per Waveshare schematic:
//   LCD SCLK=39  MOSI=38  DC=41  CS=42  RST=40  BL=15
//
// Self-contained; included only when -DENABLE_LCD_ST7789_240 is set.
// ---------------------------------------------------------------------------
#pragma once
#ifdef ENABLE_LCD_ST7789_240

#include <Arduino_GFX_Library.h>

// Arduino_GFX defines color name macros that collide with CRGB identifiers.
#undef GREEN
#undef RED
#undef BLUE
#undef YELLOW
#undef ORANGE
#undef MAGENTA
#undef CYAN
#undef PURPLE
#undef WHITE
#undef BLACK
#undef NAVY
#undef DARKGREEN
#undef DARKCYAN
#undef MAROON
#undef OLIVE
#undef LIGHTGREY
#undef DARKGREY
#undef PINK

#define LCD_BLACK 0x0000
#define LCD_WHITE 0xFFFF
#define LCD_GREEN 0x07E0
#define LCD_RED   0xF800
#define LCD_BLUE  0x001F
#define LCD_YELL  0xFFE0
#define LCD_CYAN  0x07FF
#define LCD_ORNG  0xFD20
#define LCD_DGREY 0x7BEF
#define LCD_LGREY 0xC618
#define LCD_NAVY  0x0841

#define LCD_SCLK 39
#define LCD_MOSI 38
#define LCD_CS   42
#define LCD_DC   41
#define LCD_RST  40
#define LCD_BL   15

#define LCD_W 240
#define LCD_H 320

static Arduino_DataBus *lcd_bus = nullptr;
static Arduino_GFX     *lcd     = nullptr;

struct NeighbourSlot { uint8_t id; int8_t rssi; uint32_t last_seen_ms; };
#define LCD_MAX_NEIGHBOURS 5
static NeighbourSlot s_neighbours[LCD_MAX_NEIGHBOURS] = {};

static uint16_t activity_color(float a) {
    if (a < 0.0f) a = 0.0f;
    if (a > 1.0f) a = 1.0f;
    uint8_t r = (uint8_t)(255.0f * a);
    uint8_t g = (uint8_t)(255.0f * (1.0f - fabsf(a - 0.5f) * 2.0f));
    uint8_t b = 30;
    return ((r & 0xF8) << 8) | ((g & 0xFC) << 3) | (b >> 3);
}

inline void lcd_init() {
    pinMode(LCD_BL, OUTPUT);
    digitalWrite(LCD_BL, LOW);

    lcd_bus = new Arduino_ESP32SPI(LCD_DC, LCD_CS, LCD_SCLK, LCD_MOSI, GFX_NOT_DEFINED);
    // 240x320 ST7789, portrait, IPS, no x/y offsets.
    lcd = new Arduino_ST7789(lcd_bus, LCD_RST, 0 /* portrait */, true /* IPS */,
                             240, 320, 0, 0, 0, 0);
    lcd->begin();

    // Visible boot sequence — proves panel is alive.
    lcd->fillScreen(LCD_RED);
    digitalWrite(LCD_BL, HIGH);
    delay(250);
    lcd->fillScreen(LCD_GREEN);
    delay(250);
    lcd->fillScreen(LCD_BLUE);
    delay(250);
    lcd->fillScreen(LCD_BLACK);

    lcd->fillRect(0, 0, LCD_W, 36, LCD_NAVY);
    lcd->setTextColor(LCD_WHITE);
    lcd->setTextSize(3);
    lcd->setCursor(24, 8);
    lcd->print("AEDI-S");

    lcd->setTextSize(1);
    lcd->setTextColor(LCD_DGREY);
    lcd->setCursor(8, 240);
    lcd->print("NEIGHBOURS");
    lcd->drawFastHLine(0, 250, LCD_W, LCD_DGREY);
}

inline void lcd_set_neighbour(uint8_t id, int8_t rssi) {
    uint32_t now = millis();
    for (int i = 0; i < LCD_MAX_NEIGHBOURS; i++) {
        if (s_neighbours[i].id == id) {
            s_neighbours[i].rssi = rssi;
            s_neighbours[i].last_seen_ms = now;
            return;
        }
    }
    for (int i = 0; i < LCD_MAX_NEIGHBOURS; i++) {
        if (s_neighbours[i].id == 0) {
            s_neighbours[i].id = id;
            s_neighbours[i].rssi = rssi;
            s_neighbours[i].last_seen_ms = now;
            return;
        }
    }
}

static void lcd_draw_main(uint8_t node_id, bool is_hub, int wifi_rssi_dbm,
                           uint32_t csi_frames_per_s, float activity) {
    static uint8_t  prev_id = 0xFF;
    static bool     prev_hub = false;
    static int      prev_rssi = -999;
    static uint32_t prev_rate = 0xFFFFFFFF;
    static float    prev_act = -1.0f;

    if (node_id != prev_id || is_hub != prev_hub) {
        prev_id = node_id; prev_hub = is_hub;
        lcd->fillRect(0, 40, LCD_W, 110, LCD_BLACK);
        lcd->setTextColor(is_hub ? LCD_ORNG : LCD_CYAN);
        lcd->setTextSize(11);
        int nx = node_id < 10 ? 90 : 38;
        lcd->setCursor(nx, 48);
        lcd->print(node_id);
        lcd->setTextSize(2);
        lcd->setTextColor(LCD_LGREY);
        lcd->setCursor(is_hub ? 96 : 88, 130);
        lcd->print(is_hub ? "HUB" : "EDGE");
    }

    if (abs(wifi_rssi_dbm - prev_rssi) > 2 || csi_frames_per_s != prev_rate) {
        prev_rssi = wifi_rssi_dbm; prev_rate = csi_frames_per_s;
        lcd->fillRect(0, 160, LCD_W, 60, LCD_BLACK);
        uint16_t wc = (wifi_rssi_dbm == 0) ? LCD_RED
                    : (wifi_rssi_dbm > -65 ? LCD_GREEN : LCD_YELL);
        lcd->fillCircle(16, 172, 6, wc);
        lcd->setTextColor(LCD_WHITE);
        lcd->setTextSize(2);
        lcd->setCursor(32, 164);
        if (wifi_rssi_dbm == 0) lcd->print("OFFLINE");
        else { lcd->print(wifi_rssi_dbm); lcd->print(" dBm"); }
        lcd->setTextColor(LCD_DGREY);
        lcd->setCursor(16, 194);
        lcd->print("CSI ");
        lcd->print(csi_frames_per_s);
        lcd->print(" pkt/s");
    }

    if (fabsf(activity - prev_act) > 0.02f) {
        prev_act = activity;
        int by = 220, bh = 16;
        lcd->drawRect(8, by, LCD_W - 16, bh, LCD_DGREY);
        lcd->fillRect(9, by + 1, LCD_W - 18, bh - 2, LCD_BLACK);
        int w = (int)((LCD_W - 18) * activity);
        if (w > 0) lcd->fillRect(9, by + 1, w, bh - 2, activity_color(activity));
    }
}

static void lcd_draw_neighbours(uint32_t now) {
    int y = 256;
    const int row_h = 14;
    for (int i = 0; i < LCD_MAX_NEIGHBOURS; i++) {
        int ry = y + i * row_h;
        if (ry + row_h > 320) break;
        lcd->fillRect(0, ry, LCD_W, row_h, LCD_BLACK);
        if (s_neighbours[i].id == 0) continue;
        bool stale = (now - s_neighbours[i].last_seen_ms) > 5000;
        uint16_t col = stale ? LCD_DGREY
                     : (s_neighbours[i].rssi > -65 ? LCD_GREEN : LCD_YELL);
        lcd->fillCircle(10, ry + 7, 4, col);
        lcd->setTextColor(stale ? LCD_DGREY : LCD_WHITE);
        lcd->setTextSize(1);
        lcd->setCursor(22, ry + 3);
        lcd->print("node ");
        lcd->print(s_neighbours[i].id);
        lcd->print("  ");
        lcd->print(s_neighbours[i].rssi);
        lcd->print(" dBm");
    }
}

inline void lcd_update(uint8_t node_id, bool is_hub, int wifi_rssi_dbm,
                       uint32_t csi_frames_per_s, float activity) {
    static uint32_t last_draw = 0;
    uint32_t now = millis();
    if (now - last_draw < 200) return;   // 5 Hz
    last_draw = now;
    if (!lcd) return;
    lcd_draw_main(node_id, is_hub, wifi_rssi_dbm, csi_frames_per_s, activity);
    lcd_draw_neighbours(now);
}

#endif // ENABLE_LCD_ST7789_240
