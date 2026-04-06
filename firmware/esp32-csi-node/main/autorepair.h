/**
 * @file autorepair.h
 * @brief Autorepair: Watchdog (WDT), brownout detector, and selfcheck routines.
 *
 * Provides three layers of self-healing:
 *  1. Task watchdog — feeds the WDT from the main loop; triggers reboot on hang.
 *  2. Brownout detector — logs and optionally reboots on low-voltage events.
 *  3. Selfcheck — periodic validation of WiFi, UDP, CSI pipeline health.
 *     Falls back to safe defaults and reboots if recovery fails.
 */

#ifndef AUTOREPAIR_H
#define AUTOREPAIR_H

#include "esp_err.h"
#include <stdbool.h>
#include <stdint.h>

/* ── Selfcheck status bits ─────────────────────────────────────────────── */
#define AUTOREPAIR_OK             0x00
#define AUTOREPAIR_WIFI_LOST      0x01
#define AUTOREPAIR_UDP_STALL      0x02
#define AUTOREPAIR_CSI_STALL      0x04
#define AUTOREPAIR_HEAP_LOW       0x08
#define AUTOREPAIR_BROWNOUT_EVT   0x10

/* ── Reboot reason keys (stored in NVS "autorepair" namespace) ─────── */
#define AR_NVS_NAMESPACE          "autorepair"
#define AR_NVS_LAST_REASON        "last_reason"
#define AR_NVS_REBOOT_COUNT       "reboot_cnt"
#define AR_NVS_FALLBACK_ACTIVE    "fallback"

/**
 * Autorepair runtime statistics (read-only snapshot).
 */
typedef struct {
    uint32_t selfcheck_pass;     /**< Number of passed selfchecks */
    uint32_t selfcheck_fail;     /**< Number of failed selfchecks */
    uint32_t wifi_reconnects;    /**< WiFi reconnection attempts */
    uint32_t brownout_events;    /**< Brownout events detected */
    uint32_t wdt_feeds;          /**< Total WDT feed count */
    uint32_t forced_reboots;     /**< Reboots triggered by autorepair */
    uint8_t  last_status;        /**< Last selfcheck status bitmask */
    bool     fallback_active;    /**< True if running in fallback config */
    uint32_t heap_free_min;      /**< Minimum free heap observed (bytes) */
} autorepair_stats_t;

/**
 * Initialize the autorepair subsystem.
 *
 * - Subscribes the main task to the Task WDT.
 * - Installs brownout event handler.
 * - Reads reboot history from NVS.
 * - If reboot count exceeds threshold, activates fallback config.
 *
 * Call ONCE from app_main(), after NVS is initialized.
 *
 * @param wdt_timeout_sec  Task WDT timeout (5–120 s). 0 = use default (15 s).
 * @return ESP_OK on success.
 */
esp_err_t autorepair_init(uint32_t wdt_timeout_sec);

/**
 * Feed the watchdog timer.
 *
 * Must be called from the main loop at least once per WDT timeout period.
 * Also runs a quick selfcheck (WiFi link, heap headroom).
 *
 * @return Selfcheck status bitmask (AUTOREPAIR_OK if all healthy).
 */
uint8_t autorepair_feed(void);

/**
 * Run a full selfcheck and attempt recovery if problems are found.
 *
 * Recovery actions (in order):
 *  1. WiFi lost → esp_wifi_connect() retry (up to 3 attempts).
 *  2. Heap critically low → log warning (FreeRTOS can't free).
 *  3. CSI stall → restart CSI collector.
 *  4. UDP stall → reinitialise UDP sender.
 *  5. All recovery fails → save reason to NVS, force reboot.
 *
 * @return Selfcheck status bitmask after recovery attempts.
 */
uint8_t autorepair_selfcheck(void);

/**
 * Get a snapshot of autorepair statistics.
 */
void autorepair_get_stats(autorepair_stats_t *out);

/**
 * Check if fallback configuration is active (reduced feature set).
 * Fallback activates after 3+ consecutive autorepair reboots.
 */
bool autorepair_is_fallback(void);

/**
 * Clear reboot counter and fallback flag (call after successful 5-min uptime).
 */
esp_err_t autorepair_clear_fallback(void);

#endif /* AUTOREPAIR_H */
