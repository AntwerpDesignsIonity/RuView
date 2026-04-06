#!/usr/bin/env bash
# ============================================================
#  RuView — Antwerp-Pi Hub Startup Script
#  Hub IP : 192.168.0.181
#  ESP32 nodes send UDP CSI to port 5005 on this IP.
#  Web UI  : http://192.168.0.181:3000
#  WS      : ws://192.168.0.181:3001/ws/sensing
# ============================================================
set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
RUST_DIR="${REPO_DIR}/rust-port/wifi-densepose-rs"
BINARY="${RUST_DIR}/target/release/sensing-server"
UI_DIR="${REPO_DIR}/ui"

# Colours
RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'
CYAN='\033[0;36m'; BOLD='\033[1m'; RESET='\033[0m'

# ── Mode: docker | native | auto (default) ─────────────────────────────
MODE="${1:-auto}"

log() { echo -e "${CYAN}[RuView]${RESET} $1"; }
ok()  { echo -e "  ${GREEN}✔${RESET} $1"; }
warn(){ echo -e "  ${YELLOW}⚠${RESET} $1"; }
err() { echo -e "  ${RED}✖${RESET} $1"; }

# ── Firewall: ensure UDP 5005 is open ──────────────────────────────────
open_ports() {
  if command -v ufw &>/dev/null; then
    sudo ufw allow 5005/udp &>/dev/null && ok "UDP 5005 open (ufw)" || true
    sudo ufw allow 3000/tcp &>/dev/null && ok "TCP 3000 open (ufw)" || true
    sudo ufw allow 3001/tcp &>/dev/null && ok "TCP 3001 open (ufw)" || true
  fi
}

# ── Docker mode ────────────────────────────────────────────────────────
start_docker() {
  log "Starting via Docker (ARM64 native)..."
  if ! sudo docker images ionity/ruview:latest --format '{{.ID}}' | grep -q .; then
    log "Image not found — pulling..."
    sudo docker pull ionity/ruview:latest
  fi

  # Stop any previous container
  sudo docker rm -f ruview 2>/dev/null || true

  sudo docker run -d \
    --name ruview \
    --restart unless-stopped \
    -p 3000:3000 \
    -p 3001:3001 \
    -p 5005:5005/udp \
    -e CSI_SOURCE=auto \
    -e NODE_IP=192.168.0.181 \
    ionity/ruview:latest

  ok "Container 'ruview' started"
  log "UI  → http://192.168.0.181:3000"
  log "WS  → ws://192.168.0.181:3001/ws/sensing"
  log "ESP32 nodes should target UDP 192.168.0.181:5005"
}

# ── Native mode ────────────────────────────────────────────────────────
start_native() {
  export PATH="$HOME/.cargo/bin:$PATH"

  if [[ ! -f "${BINARY}" ]]; then
    log "Binary not found — building sensing server (~10 min first time)..."
    cd "${RUST_DIR}"
    cargo build -p wifi-densepose-sensing-server --release --no-default-features
  fi

  log "Starting native sensing server..."
  # Kill any existing instance
  pkill -f sensing-server 2>/dev/null || true
  sleep 0.5

  nohup "${BINARY}" \
    --host 0.0.0.0 \
    --port 3000 \
    --ws-port 3001 \
    --csi-port 5005 \
    > "${REPO_DIR}/logs/sensing-server.log" 2>&1 &

  ok "sensing-server started (PID $!)"
  log "Logs → ${REPO_DIR}/logs/sensing-server.log"
  log "UI  → http://192.168.0.181:3000"
  log "WS  → ws://192.168.0.181:3001/ws/sensing"
}

# ── Auto: use native if built, else Docker ─────────────────────────────
main() {
  echo -e "\n${BOLD}=====================================${RESET}"
  echo -e "${BOLD}  RuView Hub — Antwerp-Pi${RESET}"
  echo -e "${BOLD}  Hub IP: 192.168.0.181${RESET}"
  echo -e "${BOLD}=====================================${RESET}\n"

  mkdir -p "${REPO_DIR}/logs"
  open_ports

  case "${MODE}" in
    docker)  start_docker ;;
    native)  start_native ;;
    auto)
      if [[ -f "${BINARY}" ]]; then
        start_native
      else
        warn "Native binary not built yet — using Docker."
        start_docker
      fi
      ;;
    stop)
      log "Stopping all RuView services..."
      sudo docker rm -f ruview 2>/dev/null && ok "Docker container stopped" || true
      pkill -f sensing-server 2>/dev/null && ok "Native server stopped" || true
      ;;
    status)
      echo -e "\n${BOLD}── Docker status ──${RESET}"
      sudo docker ps --filter name=ruview --format "table {{.ID}}\t{{.Status}}\t{{.Ports}}" 2>/dev/null || true
      echo -e "\n${BOLD}── Native process ──${RESET}"
      pgrep -a sensing-server 2>/dev/null || echo "  not running"
      echo ""
      ;;
    *)
      err "Unknown mode: ${MODE}"
      echo "Usage: $0 [auto|docker|native|stop|status]"
      exit 1
      ;;
  esac
}

main
