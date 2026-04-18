#!/usr/bin/env bash
# Flash firmware via PlatformIO, then write NVS config, for every connected ESP32-S3.
# Usage:   flash-all-nodes.sh <SSID> <PASSWORD> <TARGET_IP>
#          BOARD_ENV=<env> flash-all-nodes.sh ...    # override the PlatformIO env
#
# BOARD_ENV options (from ionity/platformio.ini):
#   esp32s3_n16r8       (default) — headless CSI node, 16MB flash / 8MB PSRAM
#   esp32s3_touch_lcd_2 — Waveshare ESP32-S3-Touch-LCD-2 (2.4" ST7789 240x320)
#   esp32s3_lcd_1_47    — Waveshare ESP32-S3-LCD-1.47 (ST7789 172x320)
#   esp32s3_amoled_1_64 — Waveshare Touch-AMOLED-1.64 (CSI only, no display yet)
set -euo pipefail

SSID="${1:-}"
PASS="${2:-}"
HUB_IP="${3:-}"
BOARD_ENV="${BOARD_ENV:-esp32s3_n16r8}"
if [[ -z "$SSID" || -z "$HUB_IP" ]]; then
  echo "Usage: $0 <SSID> <PASSWORD> <TARGET_IP>" >&2
  echo "       BOARD_ENV=<env> $0 ...   (default: esp32s3_n16r8)" >&2
  exit 2
fi

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

# Discover ports (ACM first)
PORTS=()
for p in /dev/ttyACM0 /dev/ttyACM1 /dev/ttyACM2 /dev/ttyACM3 /dev/ttyUSB0 /dev/ttyUSB1 /dev/ttyUSB2; do
  [[ -e "$p" ]] && PORTS+=("$p")
done
if [[ ${#PORTS[@]} -eq 0 ]]; then
  echo "No ESP32 serial ports found" >&2
  exit 1
fi

echo "[flash-all] Found ${#PORTS[@]} port(s): ${PORTS[*]}"
echo "[flash-all] SSID=$SSID  HUB_IP=$HUB_IP  BOARD_ENV=$BOARD_ENV"

NODE_ID=1
for PORT in "${PORTS[@]}"; do
  # Role: node 1 = hub (aggregator light), 2+ = edge
  if [[ $NODE_ID -eq 1 ]]; then LED_ROLE="hub"; else LED_ROLE="edge"; fi

  echo
  echo "=========================================="
  echo "[flash-all] Node $NODE_ID on $PORT ($LED_ROLE)"
  echo "=========================================="

  # 1. Flash firmware via PlatformIO
  ( cd ionity && pio run -e "$BOARD_ENV" -t upload --upload-port "$PORT" ) \
    || { echo "[flash-all] FAILED firmware flash on $PORT"; NODE_ID=$((NODE_ID+1)); continue; }

  # 2. Provision NVS (WiFi + target IP + node id + LED role)
  python3 firmware/esp32-csi-node/provision.py \
    --port "$PORT" \
    --no-firmware \
    --ssid "$SSID" \
    --password "$PASS" \
    --target-ip "$HUB_IP" \
    --target-port 5005 \
    --node-id "$NODE_ID" \
    --led-hub "$LED_ROLE" \
    || { echo "[flash-all] FAILED NVS provision on $PORT"; NODE_ID=$((NODE_ID+1)); continue; }

  echo "[flash-all] Node $NODE_ID ready."
  NODE_ID=$((NODE_ID+1))
done

echo
echo "[flash-all] Done. $((NODE_ID-1)) board(s) processed."
