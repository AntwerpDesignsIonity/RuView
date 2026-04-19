/**
 * @file led_effects.h
 * @brief Addressable WS2812 LED driver + HTTP control endpoints.
 *
 * ADR-062: Onboard status LED for ESP32-S3 N16R8 / SuperMini boards.
 * Target GPIO is configurable via Kconfig (default GPIO48).
 * A single WS2812 pixel is driven via the RMT peripheral.
 *
 * Exposes two HTTP endpoints (registered on the OTA HTTPD on port 8032):
 *   GET  /led        — return current effect + color + brightness
 *   POST /led        — set effect (JSON body: {"effect":"solid","r":0,"g":255,"b":0,"brightness":40})
 *
 * Effects: "off" | "solid" | "breathe" | "blink" | "rainbow" | "alert"
 */

#ifndef LED_EFFECTS_H
#define LED_EFFECTS_H

#include "esp_err.h"

#ifdef __cplusplus
extern "C" {
#endif

/** Initialize the LED driver and start the effect task.
 *  Safe to call multiple times; second call is a no-op. */
esp_err_t led_effects_init(void);

/** Register /led HTTP handlers on an existing httpd handle. */
esp_err_t led_effects_register_endpoints(void *httpd_handle);

/** Optional: directly set an effect programmatically. */
esp_err_t led_effects_set(const char *effect, uint8_t r, uint8_t g, uint8_t b, uint8_t brightness);

#ifdef __cplusplus
}
#endif

#endif /* LED_EFFECTS_H */
