#!/usr/bin/env bash
# ============================================================
#  AEDI-S — ESP32-S3 Node Flash + Provision Script
#
#  Run this script ON the Antwerp-Pi hub when you connect
#  each ESP32-S3 node via USB.
#
#  Usage:
#    ./scripts/flash-node.sh /dev/ttyUSB0 "YourWiFi" "password" 1
#                             ^port         ^SSID       ^pass     ^node_id
#
#  Hub IP:  192.168.0.181  (Antwerp-Pi)
#  Hub UDP: 5005
# ============================================================
set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
FW_DIR="${REPO_DIR}/firmware/esp32-csi-node/release_bins"

# ── Load local credentials if present ────────────────────────
ENV_LOCAL="${REPO_DIR}/.env.local"
if [[ -f "${ENV_LOCAL}" ]]; then
  # shellcheck disable=SC1090
  set -a; source "${ENV_LOCAL}"; set +a
fi

# ── Colours ─────────────────────────────────────────────────
RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'
CYAN='\033[0;36m'; BOLD='\033[1m'; RESET='\033[0m'
log() { echo -e "${CYAN}[flash]${RESET} $1"; }
ok()  { echo -e "  ${GREEN}✔${RESET} $1"; }
err() { echo -e "  ${RED}✖${RESET} $1"; exit 1; }

# ── Args ─────────────────────────────────────────────────────
PORT="${1:-/dev/ttyUSB0}"
SSID="${2:-${WIFI_SSID:-}}"
PASS="${3:-${WIFI_PASS:-}}"
NODE_ID="${4:-1}"
HUB_IP="${HUB_IP:-192.168.0.181}"
HUB_PORT="${HUB_PORT:-5005}"

# Detect 4MB vs 8MB flash from device
# Default to 8MB firmware (ESP32-S3 8MB); change to 4mb if needed
FLASH_SIZE="${FLASH_SIZE:-8mb}"

if [[ "${FLASH_SIZE}" == "4mb" ]]; then
  FW_BIN="${FW_DIR}/esp32-csi-node-4mb.bin"
  PT_BIN="${FW_DIR}/partition-table-4mb.bin"
else
  FW_BIN="${FW_DIR}/esp32-csi-node.bin"
  PT_BIN="${FW_DIR}/partition-table.bin"
fi

# ── Requirements check ───────────────────────────────────────
echo -e "\n${BOLD}=== AEDI-S Node Flash ===${RESET}"
echo -e "  Port     : ${PORT}"
echo -e "  SSID     : ${SSID:-<not set — will prompt>}"
echo -e "  Node ID  : ${NODE_ID}"
echo -e "  Hub      : ${HUB_IP}:${HUB_PORT}/udp"
echo -e "  Firmware : ${FLASH_SIZE} (${FW_BIN##*/})"
echo ""

if [[ ! -e "${PORT}" ]]; then
  err "Serial port ${PORT} not found. Is the ESP32 plugged in?"
fi

if ! python3 -c "import esptool" 2>/dev/null; then
  log "Installing esptool..."
  pip install esptool --quiet
fi

# ── Prompt if SSID not given ─────────────────────────────────
if [[ -z "${SSID}" ]]; then
  read -rp "  WiFi SSID: " SSID
fi
if [[ -z "${PASS}" ]]; then
  read -rsp "  WiFi Password: " PASS
  echo ""
fi

# ── Flash firmware ───────────────────────────────────────────
log "Flashing firmware to ${PORT}..."
python3 -m esptool \
  --chip esp32s3 \
  --port "${PORT}" \
  --baud 460800 \
  write_flash \
    0x0      "${FW_DIR}/bootloader.bin" \
    0x8000   "${PT_BIN}" \
    0xf000   "${FW_DIR}/ota_data_initial.bin" \
    0x20000  "${FW_BIN}"
ok "Firmware flashed"

# ── Provision WiFi + Hub IP ──────────────────────────────────
log "Provisioning: WiFi='${SSID}', hub=${HUB_IP}:${HUB_PORT}, node_id=${NODE_ID}"
python3 "${REPO_DIR}/firmware/esp32-csi-node/provision.py" \
  --port "${PORT}" \
  --ssid "${SSID}" \
  --password "${PASS}" \
  --target-ip "${HUB_IP}" \
  --target-port "${HUB_PORT}" \
  --node-id "${NODE_ID}"
ok "Provisioned"

log "Node ${NODE_ID} ready. Unplug and power up — it will connect and stream CSI to ${HUB_IP}:${HUB_PORT}/udp"
echo ""
echo -e "${GREEN}${BOLD}Done! Watch the hub: curl http://192.168.0.181:3000/api/v1/sensing/latest${RESET}"
