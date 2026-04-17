---
name: 'ESP32 Firmware'
description: 'Conventions for ESP32 firmware: PlatformIO Arduino framework (ionity/), ESP-IDF C firmware (firmware/esp32-csi-node/), provisioning, and NVS configuration.'
applyTo: '{ionity/**,firmware/**}'
---

# ESP32 Firmware Conventions

## Two Firmware Codebases

| Directory | Framework | Language | Purpose |
|-----------|-----------|----------|---------|
| `ionity/` | PlatformIO + Arduino | C++ | Production CSI streaming firmware |
| `firmware/esp32-csi-node/` | ESP-IDF v5.4 | C | Advanced CSI firmware with NVS, OTA, TDM |

## Supported Hardware

- **ESP32-S3** (8MB flash, N16R8) — primary target, dual-core Xtensa
- **ESP32-S3 SuperMini** (4MB flash) — compact variant
- **ESP32-C6** — RISC-V, mmWave bridge only (not CSI)
- **NOT supported:** ESP32 (original), ESP32-C3 — single-core, can't run CSI DSP

## PlatformIO Build (ionity/)

```bash
cd ionity
platformio run -e esp32s3_n16r8           # build
platformio run -e esp32s3_n16r8 -t upload # flash
platformio device monitor --baud 460800   # serial monitor
```

Board config is in `ionity/platformio.ini`. Monitor baud rate is **460800** — not 115200.

## ESP-IDF Build (firmware/esp32-csi-node/)

```bash
# Requires ESP-IDF v5.4 installed
cd firmware/esp32-csi-node
idf.py build
idf.py -p /dev/ttyACM0 flash
idf.py monitor
```

### sdkconfig Variants

- `sdkconfig.defaults.template` → 8MB flash (production)
- `sdkconfig.defaults.4mb` → 4MB flash (SuperMini)
- Never commit `sdkconfig` directly — use `sdkconfig.defaults.*`

## Provisioning

WiFi credentials and hub target are stored in NVS (Non-Volatile Storage):

```bash
python firmware/esp32-csi-node/provision.py \
  --port /dev/ttyACM0 \
  --ssid "WiFiName" \
  --password "secret" \
  --target-ip 192.168.0.181
```

NVS keys: `ssid`, `password`, `target_ip`, `target_port`, `node_id`

## Code Conventions

- Use `ESP_LOGI`/`ESP_LOGW`/`ESP_LOGE` for logging — never `printf`.
- CSI callback runs on WiFi task — keep it fast (<1ms). Copy data and process on a separate task.
- LED status codes are documented in `docs/led-indication.md`.
- Pin assignments: use `#define` constants at top of file, never magic numbers.
- Firmware binary must be under 1100 KB (CI size gate).

## Flash Layout (8MB)

| Partition | Offset | Size | Purpose |
|-----------|--------|------|---------|
| bootloader | 0x0 | 32KB | Second-stage bootloader |
| partition-table | 0x8000 | 4KB | Partition map |
| ota_data | 0xF000 | 8KB | OTA state |
| app | 0x20000 | ~1MB | Application firmware |
| nvs | — | 16KB | WiFi creds, node config |

## Testing

- No unit test framework on-device. Test logic through QEMU: `.github/workflows/firmware-qemu.yml`.
- Always test on real hardware before marking firmware PRs ready — mock mode missed the Kconfig threshold bug.
- Serial monitor at 460800 baud confirms CSI streaming with `CSI_DATA:` prefix lines.

## Security

- Never hardcode WiFi credentials in source — always use NVS provisioning.
- OTA updates must be signed when enabled.
- Validate all incoming UDP packets (source IP allowlist in production).
