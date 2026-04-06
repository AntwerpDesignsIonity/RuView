/**
 * @file autorepair.c
 * @brief Autorepair: WDT, brownout detector, selfcheck with fallback recovery.
 *
 * Implements the three-layer self-healing system described in autorepair.h.
 */

#include "autorepair.h"

#include <string.h>
#include "freertos/FreeRTOS.h"
#include "freertos/task.h"
#include "esp_log.h"
#include "esp_system.h"
#include "esp_wifi.h"
#include "esp_task_wdt.h"
#include "nvs_flash.h"
#include "nvs.h"
#include "sdkconfig.h"

/* Optional: brownout detector API (present on ESP32-S3). */
#if __has_include("soc/brownout_hal.h")
#include "soc/brownout_hal.h"
#endif

#include "csi_collector.h"
#include "stream_sender.h"

static const char *TAG = "autorepair";

/* ── Thresholds ──────────────────────────────────────────────────────── */
#define HEAP_CRITICAL_BYTES   (16 * 1024)   /* 16 KB minimum free heap */
#define FALLBACK_REBOOT_LIMIT 3             /* Consecutive reboots before fallback */
#define WIFI_RETRY_MAX        3             /* Quick reconnection attempts */
#define STABLE_UPTIME_SEC     300           /* 5 min = consider stable, clear reboot counter */

/* ── State ───────────────────────────────────────────────────────────── */
static autorepair_stats_t s_stats;
static bool s_initialized = false;
static uint32_t s_wdt_timeout_sec = 15;
static int64_t  s_boot_time_us = 0;

/* External: set by brownout ISR */
static volatile bool s_brownout_flag = false;

/* ── NVS helpers ─────────────────────────────────────────────────────── */
static esp_err_t _nvs_read_u32(const char *key, uint32_t *out)
{
    nvs_handle_t h;
    esp_err_t err = nvs_open(AR_NVS_NAMESPACE, NVS_READONLY, &h);
    if (err != ESP_OK) { *out = 0; return err; }
    err = nvs_get_u32(h, key, out);
    if (err != ESP_OK) *out = 0;
    nvs_close(h);
    return err;
}

static esp_err_t _nvs_write_u32(const char *key, uint32_t val)
{
    nvs_handle_t h;
    esp_err_t err = nvs_open(AR_NVS_NAMESPACE, NVS_READWRITE, &h);
    if (err != ESP_OK) return err;
    err = nvs_set_u32(h, key, val);
    if (err == ESP_OK) nvs_commit(h);
    nvs_close(h);
    return err;
}

static esp_err_t _nvs_read_u8(const char *key, uint8_t *out)
{
    nvs_handle_t h;
    esp_err_t err = nvs_open(AR_NVS_NAMESPACE, NVS_READONLY, &h);
    if (err != ESP_OK) { *out = 0; return err; }
    err = nvs_get_u8(h, key, out);
    if (err != ESP_OK) *out = 0;
    nvs_close(h);
    return err;
}

static esp_err_t _nvs_write_u8(const char *key, uint8_t val)
{
    nvs_handle_t h;
    esp_err_t err = nvs_open(AR_NVS_NAMESPACE, NVS_READWRITE, &h);
    if (err != ESP_OK) return err;
    err = nvs_set_u8(h, key, val);
    if (err == ESP_OK) nvs_commit(h);
    nvs_close(h);
    return err;
}

/* ── Brownout ISR handler ────────────────────────────────────────────── */
/* ESP-IDF fires this on the system event loop when CONFIG_ESP_BROWNOUT_DET=y */
static void brownout_event_handler(void *arg, esp_event_base_t base,
                                   int32_t id, void *data)
{
    (void)arg; (void)base; (void)id; (void)data;
    s_brownout_flag = true;
    s_stats.brownout_events++;
    ESP_EARLY_LOGE(TAG, "*** BROWNOUT detected — voltage dip ***");
}

/* ── WiFi state check ────────────────────────────────────────────────── */
static bool _wifi_is_connected(void)
{
    wifi_ap_record_t ap;
    return (esp_wifi_sta_get_ap_info(&ap) == ESP_OK);
}

static bool _wifi_reconnect(void)
{
    for (int i = 0; i < WIFI_RETRY_MAX; i++) {
        ESP_LOGW(TAG, "WiFi reconnect attempt %d/%d", i + 1, WIFI_RETRY_MAX);
        s_stats.wifi_reconnects++;
        esp_wifi_connect();
        vTaskDelay(pdMS_TO_TICKS(2000));
        if (_wifi_is_connected()) {
            ESP_LOGI(TAG, "WiFi reconnected on attempt %d", i + 1);
            return true;
        }
    }
    return false;
}

/* ── Force reboot with reason ────────────────────────────────────────── */
static void _force_reboot(uint8_t reason)
{
    ESP_LOGE(TAG, "Forcing reboot — reason=0x%02X", reason);

    /* Persist reason and increment reboot counter */
    _nvs_write_u8(AR_NVS_LAST_REASON, reason);
    uint32_t cnt = 0;
    _nvs_read_u32(AR_NVS_REBOOT_COUNT, &cnt);
    _nvs_write_u32(AR_NVS_REBOOT_COUNT, cnt + 1);

    s_stats.forced_reboots++;
    vTaskDelay(pdMS_TO_TICKS(100)); /* Let NVS flush */
    esp_restart();
}

/* ══════════════════════════════════════════════════════════════════════ */
/*  Public API                                                           */
/* ══════════════════════════════════════════════════════════════════════ */

esp_err_t autorepair_init(uint32_t wdt_timeout_sec)
{
    if (s_initialized) return ESP_OK;

    memset(&s_stats, 0, sizeof(s_stats));
    s_boot_time_us = esp_timer_get_time();
    s_wdt_timeout_sec = (wdt_timeout_sec > 0) ? wdt_timeout_sec : 15;

    /* ── 1. Read NVS reboot history ──────────────────────────────────── */
    uint32_t reboot_cnt = 0;
    _nvs_read_u32(AR_NVS_REBOOT_COUNT, &reboot_cnt);

    uint8_t last_reason = 0;
    _nvs_read_u8(AR_NVS_LAST_REASON, &last_reason);

    if (reboot_cnt > 0) {
        ESP_LOGW(TAG, "Previous autorepair reboots: %lu (last reason=0x%02X)",
                 (unsigned long)reboot_cnt, last_reason);
    }

    /* Activate fallback mode if too many consecutive reboots */
    if (reboot_cnt >= FALLBACK_REBOOT_LIMIT) {
        ESP_LOGW(TAG, "*** FALLBACK MODE — %lu consecutive reboots exceed limit (%d) ***",
                 (unsigned long)reboot_cnt, FALLBACK_REBOOT_LIMIT);
        s_stats.fallback_active = true;
        _nvs_write_u8(AR_NVS_FALLBACK_ACTIVE, 1);
    } else {
        uint8_t fb = 0;
        _nvs_read_u8(AR_NVS_FALLBACK_ACTIVE, &fb);
        s_stats.fallback_active = (fb != 0);
    }

    /* ── 2. Subscribe main task to Task WDT ─────────────────────────── */
#if CONFIG_ESP_TASK_WDT_EN
    /* ESP-IDF v5.x: Task WDT is already initialized by the system.
     * We just add the current (main) task to be monitored. */
    esp_err_t wdt_err = esp_task_wdt_add(xTaskGetCurrentTaskHandle());
    if (wdt_err == ESP_OK) {
        ESP_LOGI(TAG, "Main task subscribed to Task WDT (system timeout applies)");
    } else if (wdt_err == ESP_ERR_INVALID_STATE) {
        /* WDT not initialized yet — init with our timeout, then add. */
        esp_task_wdt_config_t wdt_cfg = {
            .timeout_ms = s_wdt_timeout_sec * 1000,
            .idle_core_mask = 0,  /* Don't watch idle tasks */
            .trigger_panic = true,
        };
        esp_task_wdt_reconfigure(&wdt_cfg);
        esp_task_wdt_add(xTaskGetCurrentTaskHandle());
        ESP_LOGI(TAG, "Task WDT configured: timeout=%lus, panic=true",
                 (unsigned long)s_wdt_timeout_sec);
    } else {
        ESP_LOGW(TAG, "Task WDT add failed: %s", esp_err_to_name(wdt_err));
    }
#else
    ESP_LOGW(TAG, "Task WDT disabled in sdkconfig — no WDT protection!");
#endif

    /* ── 3. Register brownout event handler ──────────────────────────── */
    /* ESP-IDF v5.x fires SYSTEM_EVENT on brownout if CONFIG_ESP_BROWNOUT_DET=y.
     * We hook into the default event loop. */
    esp_event_handler_register(ESP_EVENT_ANY_BASE, ESP_EVENT_ANY_ID,
                               brownout_event_handler, NULL);
    ESP_LOGI(TAG, "Brownout detector event handler registered");

    s_stats.heap_free_min = esp_get_free_heap_size();
    s_initialized = true;

    ESP_LOGI(TAG, "Autorepair initialized (WDT=%lus, fallback=%s, prior_reboots=%lu)",
             (unsigned long)s_wdt_timeout_sec,
             s_stats.fallback_active ? "ACTIVE" : "off",
             (unsigned long)reboot_cnt);

    return ESP_OK;
}

uint8_t autorepair_feed(void)
{
    if (!s_initialized) return AUTOREPAIR_OK;

    /* Feed the task watchdog — prevents reboot */
#if CONFIG_ESP_TASK_WDT_EN
    esp_task_wdt_reset();
#endif
    s_stats.wdt_feeds++;

    /* Quick health checks */
    uint8_t status = AUTOREPAIR_OK;

    /* Heap check */
    uint32_t free_heap = esp_get_free_heap_size();
    if (free_heap < s_stats.heap_free_min) {
        s_stats.heap_free_min = free_heap;
    }
    if (free_heap < HEAP_CRITICAL_BYTES) {
        status |= AUTOREPAIR_HEAP_LOW;
        ESP_LOGW(TAG, "Heap critically low: %lu bytes free", (unsigned long)free_heap);
    }

    /* Brownout event */
    if (s_brownout_flag) {
        status |= AUTOREPAIR_BROWNOUT_EVT;
        s_brownout_flag = false; /* Acknowledge — will log but not reboot immediately */
        ESP_LOGW(TAG, "Brownout event acknowledged — monitoring stability");
    }

    /* Auto-clear fallback after stable uptime */
    if (s_stats.fallback_active) {
        int64_t uptime_us = esp_timer_get_time() - s_boot_time_us;
        if (uptime_us > (int64_t)STABLE_UPTIME_SEC * 1000000LL) {
            ESP_LOGI(TAG, "Stable for %ds — clearing fallback mode", STABLE_UPTIME_SEC);
            autorepair_clear_fallback();
        }
    }

    s_stats.last_status = status;
    if (status == AUTOREPAIR_OK) {
        s_stats.selfcheck_pass++;
    } else {
        s_stats.selfcheck_fail++;
    }

    return status;
}

uint8_t autorepair_selfcheck(void)
{
    if (!s_initialized) return AUTOREPAIR_OK;

    uint8_t status = AUTOREPAIR_OK;

    /* ── WiFi connectivity ──────────────────────────────────────────── */
#ifndef CONFIG_CSI_MOCK_SKIP_WIFI_CONNECT
    if (!_wifi_is_connected()) {
        ESP_LOGW(TAG, "WiFi disconnected — attempting recovery");
        status |= AUTOREPAIR_WIFI_LOST;
        if (_wifi_reconnect()) {
            status &= ~AUTOREPAIR_WIFI_LOST; /* Recovered */
        }
    }
#endif

    /* ── Heap headroom ──────────────────────────────────────────────── */
    uint32_t free_heap = esp_get_free_heap_size();
    if (free_heap < HEAP_CRITICAL_BYTES) {
        status |= AUTOREPAIR_HEAP_LOW;
        ESP_LOGE(TAG, "Heap depleted: %lu bytes — no recovery possible",
                 (unsigned long)free_heap);
    }

    /* ── Brownout history ───────────────────────────────────────────── */
    if (s_stats.brownout_events > 3) {
        status |= AUTOREPAIR_BROWNOUT_EVT;
        ESP_LOGE(TAG, "Multiple brownout events (%lu) — power supply unstable",
                 (unsigned long)s_stats.brownout_events);
    }

    /* ── Final verdict ──────────────────────────────────────────────── */
    s_stats.last_status = status;
    if (status == AUTOREPAIR_OK) {
        s_stats.selfcheck_pass++;
    } else {
        s_stats.selfcheck_fail++;

        /* If critical failures persist and WiFi is lost, force reboot */
        if ((status & AUTOREPAIR_WIFI_LOST) && s_stats.selfcheck_fail > 5) {
            _force_reboot(status);
        }
    }

    return status;
}

void autorepair_get_stats(autorepair_stats_t *out)
{
    if (out) {
        *out = s_stats;
        /* Update heap minimum to current */
        uint32_t free_heap = esp_get_free_heap_size();
        if (free_heap < out->heap_free_min) {
            out->heap_free_min = free_heap;
        }
    }
}

bool autorepair_is_fallback(void)
{
    return s_stats.fallback_active;
}

esp_err_t autorepair_clear_fallback(void)
{
    s_stats.fallback_active = false;
    _nvs_write_u8(AR_NVS_FALLBACK_ACTIVE, 0);
    _nvs_write_u32(AR_NVS_REBOOT_COUNT, 0);
    _nvs_write_u8(AR_NVS_LAST_REASON, 0);
    ESP_LOGI(TAG, "Fallback cleared — reboot counter reset");
    return ESP_OK;
}
