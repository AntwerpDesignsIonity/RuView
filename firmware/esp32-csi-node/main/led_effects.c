/**
 * @file led_effects.c
 * @brief WS2812 LED effects + HTTP control (ADR-062).
 *
 * Drives a single onboard addressable RGB LED (WS2812) via the ESP-IDF
 * led_strip component. Animation runs in a dedicated FreeRTOS task at
 * ~50 Hz so effects like breathe/rainbow don't block the main loop.
 *
 * To build, ensure `led_strip` is registered as a component dependency
 * (already added to main/idf_component.yml).
 */

#include "led_effects.h"

#include <string.h>
#include <stdlib.h>

#include "freertos/FreeRTOS.h"
#include "freertos/task.h"
#include "freertos/semphr.h"
#include "esp_log.h"
#include "esp_http_server.h"
#include "sdkconfig.h"
#include "led_strip.h"

#ifndef CONFIG_LED_EFFECTS_GPIO
#define CONFIG_LED_EFFECTS_GPIO 48
#endif
#ifndef CONFIG_LED_EFFECTS_COUNT
#define CONFIG_LED_EFFECTS_COUNT 1
#endif

static const char *TAG = "led";

typedef enum {
    FX_OFF = 0,
    FX_SOLID,
    FX_BREATHE,
    FX_BLINK,
    FX_RAINBOW,
    FX_ALERT,
} led_fx_t;

static led_strip_handle_t s_strip      = NULL;
static SemaphoreHandle_t  s_lock       = NULL;
static TaskHandle_t       s_task       = NULL;
static bool               s_inited     = false;

static led_fx_t s_fx        = FX_SOLID;
static uint8_t  s_r         = 0;
static uint8_t  s_g         = 32;
static uint8_t  s_b         = 0;
static uint8_t  s_brightness = 40;  /* 0..255 */

static const char *fx_name(led_fx_t fx)
{
    switch (fx) {
        case FX_OFF:     return "off";
        case FX_SOLID:   return "solid";
        case FX_BREATHE: return "breathe";
        case FX_BLINK:   return "blink";
        case FX_RAINBOW: return "rainbow";
        case FX_ALERT:   return "alert";
        default:         return "solid";
    }
}

static led_fx_t fx_from_name(const char *s)
{
    if (!s)                       return FX_SOLID;
    if (!strcmp(s, "off"))        return FX_OFF;
    if (!strcmp(s, "solid"))      return FX_SOLID;
    if (!strcmp(s, "breathe"))    return FX_BREATHE;
    if (!strcmp(s, "blink"))      return FX_BLINK;
    if (!strcmp(s, "rainbow"))    return FX_RAINBOW;
    if (!strcmp(s, "alert"))      return FX_ALERT;
    return FX_SOLID;
}

/** 8-bit HSV-to-RGB for the rainbow effect. */
static void hsv_to_rgb(uint16_t h, uint8_t s, uint8_t v,
                       uint8_t *r, uint8_t *g, uint8_t *b)
{
    uint8_t region = (h / 60) % 6;
    uint16_t remainder = (h - region * 60) * 255 / 60;
    uint8_t p  = (v * (255 - s)) / 255;
    uint8_t q  = (v * (255 - (s * remainder) / 255)) / 255;
    uint8_t t  = (v * (255 - (s * (255 - remainder)) / 255)) / 255;

    switch (region) {
        case 0:  *r = v; *g = t; *b = p; break;
        case 1:  *r = q; *g = v; *b = p; break;
        case 2:  *r = p; *g = v; *b = t; break;
        case 3:  *r = p; *g = q; *b = v; break;
        case 4:  *r = t; *g = p; *b = v; break;
        default: *r = v; *g = p; *b = q; break;
    }
}

static void apply_pixel(uint8_t r, uint8_t g, uint8_t b, uint8_t brightness)
{
    if (!s_strip) return;
    uint32_t rr = ((uint32_t)r * brightness) / 255;
    uint32_t gg = ((uint32_t)g * brightness) / 255;
    uint32_t bb = ((uint32_t)b * brightness) / 255;
    led_strip_set_pixel(s_strip, 0, (uint8_t)rr, (uint8_t)gg, (uint8_t)bb);
    led_strip_refresh(s_strip);
}

static void led_task(void *arg)
{
    (void)arg;
    uint32_t phase = 0;
    const TickType_t dt = pdMS_TO_TICKS(20);   /* ~50 Hz */
    while (1) {
        led_fx_t fx;
        uint8_t r, g, b, br;
        xSemaphoreTake(s_lock, portMAX_DELAY);
        fx = s_fx; r = s_r; g = s_g; b = s_b; br = s_brightness;
        xSemaphoreGive(s_lock);

        switch (fx) {
            case FX_OFF:
                apply_pixel(0, 0, 0, 0);
                break;
            case FX_SOLID:
                apply_pixel(r, g, b, br);
                break;
            case FX_BREATHE: {
                /* 3 s sine-like cycle (triangle). */
                uint32_t t = phase % 150;  /* 150 * 20 ms = 3 s */
                uint32_t level = t < 75 ? t : 150 - t;  /* 0..75 */
                uint8_t scaled = (uint8_t)((br * level) / 75);
                apply_pixel(r, g, b, scaled);
                break;
            }
            case FX_BLINK: {
                uint32_t t = phase % 50;   /* 1 s total */
                apply_pixel(r, g, b, (t < 25) ? br : 0);
                break;
            }
            case FX_RAINBOW: {
                uint8_t rr, gg, bb;
                hsv_to_rgb((phase * 6) % 360, 255, 255, &rr, &gg, &bb);
                apply_pixel(rr, gg, bb, br);
                break;
            }
            case FX_ALERT: {
                /* Fast red blink, full brightness regardless of setting. */
                uint32_t t = phase % 20;   /* 400 ms */
                apply_pixel(255, 0, 0, (t < 10) ? 200 : 0);
                break;
            }
        }
        phase++;
        vTaskDelay(dt);
    }
}

esp_err_t led_effects_set(const char *effect, uint8_t r, uint8_t g, uint8_t b, uint8_t brightness)
{
    if (!s_inited) return ESP_ERR_INVALID_STATE;
    xSemaphoreTake(s_lock, portMAX_DELAY);
    s_fx = fx_from_name(effect);
    s_r = r; s_g = g; s_b = b;
    s_brightness = brightness;
    xSemaphoreGive(s_lock);
    ESP_LOGI(TAG, "set effect=%s rgb=(%u,%u,%u) br=%u", fx_name(s_fx), r, g, b, brightness);
    return ESP_OK;
}

/* -------- HTTP handlers -------- */

static esp_err_t led_get_handler(httpd_req_t *req)
{
    char buf[160];
    uint8_t r, g, b, br;
    led_fx_t fx;
    xSemaphoreTake(s_lock, portMAX_DELAY);
    fx = s_fx; r = s_r; g = s_g; b = s_b; br = s_brightness;
    xSemaphoreGive(s_lock);

    int n = snprintf(buf, sizeof(buf),
        "{\"effect\":\"%s\",\"r\":%u,\"g\":%u,\"b\":%u,\"brightness\":%u,\"gpio\":%d}",
        fx_name(fx), r, g, b, br, CONFIG_LED_EFFECTS_GPIO);
    httpd_resp_set_type(req, "application/json");
    return httpd_resp_send(req, buf, n);
}

/** Very small hand-rolled JSON extractor (avoids cJSON dep). */
static bool json_get_int(const char *body, const char *key, int *out)
{
    char needle[32];
    snprintf(needle, sizeof(needle), "\"%s\"", key);
    const char *p = strstr(body, needle);
    if (!p) return false;
    p = strchr(p + strlen(needle), ':');
    if (!p) return false;
    p++;
    while (*p == ' ' || *p == '\t') p++;
    *out = atoi(p);
    return true;
}

static bool json_get_str(const char *body, const char *key, char *out, size_t cap)
{
    char needle[32];
    snprintf(needle, sizeof(needle), "\"%s\"", key);
    const char *p = strstr(body, needle);
    if (!p) return false;
    p = strchr(p + strlen(needle), ':');
    if (!p) return false;
    p++;
    while (*p == ' ' || *p == '\t') p++;
    if (*p != '"') return false;
    p++;
    size_t i = 0;
    while (*p && *p != '"' && i + 1 < cap) out[i++] = *p++;
    out[i] = 0;
    return true;
}

static esp_err_t led_post_handler(httpd_req_t *req)
{
    char body[256] = {0};
    int total = 0;
    while (total < (int)sizeof(body) - 1) {
        int n = httpd_req_recv(req, body + total, sizeof(body) - 1 - total);
        if (n <= 0) break;
        total += n;
    }
    body[total] = 0;

    char effect[16] = "solid";
    int r = s_r, g = s_g, b = s_b, br = s_brightness;
    json_get_str(body, "effect", effect, sizeof(effect));
    int tmp;
    if (json_get_int(body, "r", &tmp))          r  = tmp & 0xFF;
    if (json_get_int(body, "g", &tmp))          g  = tmp & 0xFF;
    if (json_get_int(body, "b", &tmp))          b  = tmp & 0xFF;
    if (json_get_int(body, "brightness", &tmp)) br = (tmp < 0) ? 0 : (tmp > 255 ? 255 : tmp);

    led_effects_set(effect, (uint8_t)r, (uint8_t)g, (uint8_t)b, (uint8_t)br);

    char resp[128];
    int n = snprintf(resp, sizeof(resp),
        "{\"ok\":true,\"effect\":\"%s\",\"r\":%d,\"g\":%d,\"b\":%d,\"brightness\":%d}",
        effect, r, g, b, br);
    httpd_resp_set_type(req, "application/json");
    return httpd_resp_send(req, resp, n);
}

esp_err_t led_effects_register_endpoints(void *httpd_handle)
{
    if (!httpd_handle) return ESP_ERR_INVALID_ARG;
    httpd_handle_t h = (httpd_handle_t)httpd_handle;

    httpd_uri_t get_uri = {
        .uri = "/led", .method = HTTP_GET,
        .handler = led_get_handler, .user_ctx = NULL,
    };
    httpd_register_uri_handler(h, &get_uri);

    httpd_uri_t post_uri = {
        .uri = "/led", .method = HTTP_POST,
        .handler = led_post_handler, .user_ctx = NULL,
    };
    httpd_register_uri_handler(h, &post_uri);

    ESP_LOGI(TAG, "registered /led endpoints (GPIO %d)", CONFIG_LED_EFFECTS_GPIO);
    return ESP_OK;
}

esp_err_t led_effects_init(void)
{
    if (s_inited) return ESP_OK;

    s_lock = xSemaphoreCreateMutex();
    if (!s_lock) return ESP_ERR_NO_MEM;

    led_strip_config_t strip_cfg = {
        .strip_gpio_num = CONFIG_LED_EFFECTS_GPIO,
        .max_leds       = CONFIG_LED_EFFECTS_COUNT,
        .led_model      = LED_MODEL_WS2812,
        .color_component_format = LED_STRIP_COLOR_COMPONENT_FMT_GRB,
        .flags = { .invert_out = 0 },
    };
    led_strip_rmt_config_t rmt_cfg = {
        .clk_src       = RMT_CLK_SRC_DEFAULT,
        .resolution_hz = 10 * 1000 * 1000,
        .mem_block_symbols = 64,
        .flags = { .with_dma = 0 },
    };

    esp_err_t err = led_strip_new_rmt_device(&strip_cfg, &rmt_cfg, &s_strip);
    if (err != ESP_OK) {
        ESP_LOGE(TAG, "led_strip init failed: %s", esp_err_to_name(err));
        return err;
    }

    led_strip_clear(s_strip);
    s_inited = true;

    BaseType_t ok = xTaskCreate(led_task, "led_fx", 2048, NULL, 3, &s_task);
    if (ok != pdPASS) {
        ESP_LOGE(TAG, "led_task create failed");
        return ESP_FAIL;
    }

    ESP_LOGI(TAG, "LED init OK — GPIO %d (WS2812)", CONFIG_LED_EFFECTS_GPIO);
    return ESP_OK;
}
