// ---------------------------------------------------------------------------
// lcd_co5300_164.h — Waveshare ESP32-S3-Touch-AMOLED-1.64 (CO5300, 280x456).
// Self-contained; included only when -DENABLE_LCD_CO5300_164 is set.
//
// Hardware:  QSPI AMOLED, rounded-rectangle 1.64", CST820 touch (I2C).
// Pin map (VERIFIED from Arduino_GFX dev-device table —
//   WAVESHARE_ESP32_S3_TOUCH_AMOLED_1_64):
//   QSPI  CS=9  SCK=10  D0=11  D1=12  D2=13  D3=14
//   RST=21
//   CO5300 column/row offsets: 20,0,180,24  (MANDATORY for this panel)
//
// CO5300 requires a full-frame Arduino_Canvas — partial-window writes don't
// render correctly. We draw into the canvas and flush() once per update.
//
// Exposes the same API as lcd_st7789_147.h so main.cpp is unchanged:
//   lcd_init()
//   lcd_update(node_id, is_hub, rssi_dbm, csi_frames_per_s, activity)
//   lcd_set_neighbour(id, rssi)
// ---------------------------------------------------------------------------
#pragma once
#ifdef ENABLE_LCD_CO5300_164

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

// Waveshare ESP32-S3-Touch-AMOLED-1.64 — QSPI pins (verified)
#define LCD_QSPI_CS   9
#define LCD_QSPI_SCK  10
#define LCD_QSPI_D0   11
#define LCD_QSPI_D1   12
#define LCD_QSPI_D2   13
#define LCD_QSPI_D3   14
#define LCD_RST       21

// CO5300 active-area offsets for the 1.64" 280x456 panel (MANDATORY)
#define LCD_COL_OFF1  20
#define LCD_ROW_OFF1  0
#define LCD_COL_OFF2  180
#define LCD_ROW_OFF2  24

#define LCD_W 280
#define LCD_H 456

static Arduino_DataBus *lcd_bus    = nullptr;
static Arduino_GFX     *lcd_panel  = nullptr;  // CO5300 direct driver
static Arduino_Canvas  *lcd        = nullptr;  // Full-frame canvas (flushed)

struct NeighbourSlot { uint8_t id; int8_t rssi; uint32_t last_seen_ms; };
#define LCD_MAX_NEIGHBOURS 4
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
    // QSPI bus: Arduino_ESP32QSPI(cs, sck, d0, d1, d2, d3)
    lcd_bus = new Arduino_ESP32QSPI(
        LCD_QSPI_CS, LCD_QSPI_SCK,
        LCD_QSPI_D0, LCD_QSPI_D1, LCD_QSPI_D2, LCD_QSPI_D3);

    // CO5300 driver — Arduino_GFX v1.6.5 signature (no `ips` param):
    //   Arduino_CO5300(bus, rst, r, w, h, col_off1, row_off1, col_off2, row_off2)
    // This matches the WAVESHARE_ESP32_S3_TOUCH_AMOLED_1_64 preset verbatim.
    lcd_panel = new Arduino_CO5300(
        lcd_bus, LCD_RST, 0 /* rotation */,
        LCD_W, LCD_H,
        LCD_COL_OFF1, LCD_ROW_OFF1, LCD_COL_OFF2, LCD_ROW_OFF2);

    // CO5300 does not accept partial-window writes reliably — wrap in a full
    // RGB565 framebuffer and flush() once per update.
    lcd = new Arduino_Canvas(LCD_W, LCD_H, lcd_panel, 0, 0, 0);

    Serial.println("[lcd] calling Arduino_Canvas::begin() ...");
    if (!lcd->begin()) {
        Serial.println("[lcd] Arduino_Canvas begin() FAILED — no PSRAM or wiring?");
        delete lcd;       lcd       = nullptr;
        delete lcd_panel; lcd_panel = nullptr;
        return;
    }
    Serial.println("[lcd] begin OK — running boot flash");

    // Boot flash — confirms panel + QSPI wiring are alive before any text.
    lcd->fillScreen(LCD_RED);   lcd->flush(); delay(200);
    lcd->fillScreen(LCD_GREEN); lcd->flush(); delay(200);
    lcd->fillScreen(LCD_BLUE);  lcd->flush(); delay(200);
    lcd->fillScreen(LCD_BLACK);

    // Header bar
    lcd->fillRect(0, 0, LCD_W, 40, LCD_NAVY);
    lcd->setTextColor(LCD_WHITE);
    lcd->setTextSize(3);
    lcd->setCursor(20, 8);
    lcd->print("AEDI-S");

    lcd->setTextSize(1);
    lcd->setTextColor(LCD_DGREY);
    lcd->setCursor(LCD_W - 110, 15);
    lcd->print("1.64\" AMOLED");

    // Neighbours footer label
    lcd->setCursor(8, 350);
    lcd->setTextColor(LCD_DGREY);
    lcd->print("NEIGHBOURS");
    lcd->drawFastHLine(0, 358, LCD_W, LCD_DGREY);

    lcd->flush();
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

    // Big node number block (y=56..180)
    if (node_id != prev_id || is_hub != prev_hub) {
        prev_id = node_id; prev_hub = is_hub;
        lcd->fillRect(0, 48, LCD_W, 140, LCD_BLACK);
        lcd->setTextColor(is_hub ? LCD_ORNG : LCD_CYAN);
        lcd->setTextSize(12);
        int nx = node_id < 10 ? (LCD_W / 2 - 36) : (LCD_W / 2 - 72);
        lcd->setCursor(nx, 60);
        lcd->print(node_id);
        lcd->setTextSize(3);
        lcd->setTextColor(LCD_LGREY);
        lcd->setCursor(is_hub ? (LCD_W / 2 - 30) : (LCD_W / 2 - 40), 160);
        lcd->print(is_hub ? "HUB" : "EDGE");
    }

    // RSSI + packet rate (y=200..290)
    if (abs(wifi_rssi_dbm - prev_rssi) > 2 || csi_frames_per_s != prev_rate) {
        prev_rssi = wifi_rssi_dbm; prev_rate = csi_frames_per_s;
        lcd->fillRect(0, 200, LCD_W, 90, LCD_BLACK);
        uint16_t wc = (wifi_rssi_dbm == 0) ? LCD_RED
                    : (wifi_rssi_dbm > -65 ? LCD_GREEN : LCD_YELL);
        lcd->fillCircle(16, 216, 7, wc);
        lcd->setTextColor(LCD_WHITE);
        lcd->setTextSize(2);
        lcd->setCursor(32, 208);
        if (wifi_rssi_dbm == 0) lcd->print("OFFLINE");
        else { lcd->print(wifi_rssi_dbm); lcd->print(" dBm"); }
        lcd->setTextColor(LCD_DGREY);
        lcd->setTextSize(2);
        lcd->setCursor(16, 250);
        lcd->print("CSI ");
        lcd->print(csi_frames_per_s);
        lcd->print(" pkt/s");
    }

    // Activity bar (y=300..324)
    if (fabsf(activity - prev_act) > 0.02f) {
        prev_act = activity;
        int by = 300, bh = 22;
        lcd->drawRect(8, by, LCD_W - 16, bh, LCD_DGREY);
        lcd->fillRect(9, by + 1, LCD_W - 18, bh - 2, LCD_BLACK);
        int w = (int)((LCD_W - 18) * activity);
        if (w > 0) lcd->fillRect(9, by + 1, w, bh - 2, activity_color(activity));
    }
}

static void lcd_draw_neighbours(uint32_t now) {
    const int y0 = 365;
    const int row_h = 20;
    for (int i = 0; i < LCD_MAX_NEIGHBOURS; i++) {
        int ry = y0 + i * row_h;
        if (ry + row_h > LCD_H) break;
        lcd->fillRect(0, ry, LCD_W, row_h, LCD_BLACK);
        if (s_neighbours[i].id == 0) continue;
        bool stale = (now - s_neighbours[i].last_seen_ms) > 5000;
        uint16_t col = stale ? LCD_DGREY
                     : (s_neighbours[i].rssi > -65 ? LCD_GREEN : LCD_YELL);
        lcd->fillCircle(12, ry + 10, 5, col);
        lcd->setTextColor(stale ? LCD_DGREY : LCD_WHITE);
        lcd->setTextSize(1);
        lcd->setCursor(26, ry + 6);
        lcd->print("node ");
        lcd->print(s_neighbours[i].id);
        lcd->print("   ");
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
    lcd->flush();   // push the canvas framebuffer to the CO5300 panel
}

#endif // ENABLE_LCD_CO5300_164
