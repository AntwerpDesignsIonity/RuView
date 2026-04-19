#!/usr/bin/env bash
# OTA flash an ESP32-S3 node running v0.6.1+ firmware.
# Usage: ./scripts/ota-flash.sh <node_ip> [bin_path]
set -euo pipefail

IP="${1:?Usage: $0 <node_ip> [bin_path]}"
BIN="${2:-firmware/esp32-csi-node/release_bins/esp32-csi-node.bin}"
PORT=8032

[[ -f "$BIN" ]] || { echo "FATAL: binary not found: $BIN" >&2; exit 1; }

echo "=== OTA flash: $IP:$PORT ==="
echo "--- before:"
curl -sS -m 3 "http://$IP:$PORT/ota/status" || { echo "OTA endpoint unreachable"; exit 1; }
echo

echo "--- uploading $(wc -c < "$BIN") bytes..."
START=$(date +%s)
RESP=$(curl -sS -X POST \
  -H 'Content-Type: application/octet-stream' \
  --data-binary "@$BIN" \
  --max-time 180 \
  -w '\nHTTP %{http_code}\n' \
  "http://$IP:$PORT/ota")
echo "$RESP"
END=$(date +%s)
echo "--- duration: $((END-START))s"

echo "--- waiting 15s for reboot..."
sleep 15

echo "--- after:"
curl -sS -m 3 "http://$IP:$PORT/ota/status" || echo "(not yet responding — node still booting)"
echo
