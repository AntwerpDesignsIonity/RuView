#!/usr/bin/env bash
# ╔════════════════════════════════════════════════════════════════╗
# ║  AEDI-S quick launcher — real-data-only sensing stack         ║
# ║  Ionity Global (Pty) Ltd  |  www.ionity.today                 ║
# ╚════════════════════════════════════════════════════════════════╝
#
# Usage:
#   ./start.sh              start server, open dashboard in browser
#   ./start.sh no-browser   start server without launching browser
#   ./start.sh stop         stop the server
#   ./start.sh status       show status + URLs
#   ./start.sh restart      stop then start
#   ./start.sh logs         tail the sensing-server log
#   ./start.sh build        force rebuild of the sensing-server binary
#
# Ports (override via env):
#   HTTP_PORT (3000)  WS_PORT (3001)  BIND_ADDR (0.0.0.0)

set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "${REPO_DIR}"

HTTP_PORT="${HTTP_PORT:-3000}"
WS_PORT="${WS_PORT:-3001}"
BIND_ADDR="${BIND_ADDR:-0.0.0.0}"

BIN="${REPO_DIR}/rust-port/wifi-densepose-rs/target/debug/sensing-server"
LOG_DIR="${REPO_DIR}/logs"
LOG_FILE="${LOG_DIR}/sensing-server.log"
PID_FILE="${LOG_DIR}/sensing-server.pid"

# ── Colours ──────────────────────────────────────────────────────
C_B=$'\033[1m'; C_C=$'\033[36m'; C_G=$'\033[32m'; C_Y=$'\033[33m'
C_R=$'\033[31m'; C_D=$'\033[2m'; C_0=$'\033[0m'
info() { echo -e "${C_G}[✓]${C_0} $*"; }
warn() { echo -e "${C_Y}[!]${C_0} $*"; }
err()  { echo -e "${C_R}[x]${C_0} $*" >&2; }
die()  { err "$*"; exit 1; }

# ── Host IP detection ────────────────────────────────────────────
detect_ip() {
  local ip
  ip="$(hostname -I 2>/dev/null | awk '{print $1}')"
  [[ -z "$ip" ]] && ip="$(ip route get 1.1.1.1 2>/dev/null | awk '/src/ {for(i=1;i<=NF;i++) if($i=="src") print $(i+1)}' | head -1)"
  [[ -z "$ip" ]] && ip="127.0.0.1"
  echo "$ip"
}

# ── Banner / URL list ────────────────────────────────────────────
print_urls() {
  local host="$1"
  cat <<EOF

${C_B}${C_C}═══ AEDI-S Web GUI ═════════════════════════════════════════════${C_0}
  ${C_B}Dashboard:${C_0}    http://${host}:${HTTP_PORT}/ui/index.html
  Pose Fusion:  http://${host}:${HTTP_PORT}/ui/pose-fusion.html
  Nodes:        http://${host}:${HTTP_PORT}/ui/nodes.html
  Monitor:      http://${host}:${HTTP_PORT}/ui/monitor.html
  Spectrogram:  http://${host}:${HTTP_PORT}/ui/spectrogram.html
  Observatory:  http://${host}:${HTTP_PORT}/ui/observatory.html
  Diagnostics:  http://${host}:${HTTP_PORT}/ui/diagnostics.html
  Provisioning: http://${host}:${HTTP_PORT}/ui/provisioning.html
  ${C_D}Health:       http://${host}:${HTTP_PORT}/health${C_0}
  ${C_D}WebSocket:    ws://${host}:${WS_PORT}/ws/sensing${C_0}
${C_B}${C_C}═══════════════════════════════════════════════════════════════${C_0}

EOF
}

# ── Is server running? ───────────────────────────────────────────
running_pid() {
  pgrep -f "sensing-server.*--http-port ${HTTP_PORT}" | head -1 || true
}

# ── Build if missing ─────────────────────────────────────────────
ensure_binary() {
  if [[ -x "${BIN}" && "${1:-}" != "force" ]]; then
    return 0
  fi
  info "Building sensing-server (this takes ~40 s the first time)…"
  cd "${REPO_DIR}/rust-port/wifi-densepose-rs"
  cargo build -p wifi-densepose-sensing-server --no-default-features
  cd "${REPO_DIR}"
  [[ -x "${BIN}" ]] || die "Build succeeded but binary not found at ${BIN}"
  info "Build complete."
}

# ── Open a URL in the default browser ────────────────────────────
open_browser() {
  local url="$1"
  if command -v xdg-open >/dev/null 2>&1; then
    xdg-open "$url" >/dev/null 2>&1 &
  elif command -v open >/dev/null 2>&1; then
    open "$url" >/dev/null 2>&1 &
  elif [[ -n "${BROWSER:-}" ]]; then
    "$BROWSER" "$url" >/dev/null 2>&1 &
  else
    warn "No browser launcher found — open the URL manually."
    return 1
  fi
}

# ── Wait for /health to respond ──────────────────────────────────
wait_ready() {
  local tries=30
  while (( tries-- > 0 )); do
    if curl -sf "http://127.0.0.1:${HTTP_PORT}/health" >/dev/null 2>&1; then
      return 0
    fi
    sleep 0.5
  done
  return 1
}

# ── Start ────────────────────────────────────────────────────────
cmd_start() {
  local open_gui="${1:-yes}"
  local pid
  pid="$(running_pid)"
  if [[ -n "$pid" ]]; then
    info "Server already running (PID $pid)."
  else
    ensure_binary
    mkdir -p "${LOG_DIR}"
    info "Starting sensing-server on :${HTTP_PORT} (WS :${WS_PORT})…"
    nohup "${BIN}" \
      --http-port "${HTTP_PORT}" \
      --ws-port "${WS_PORT}" \
      --bind-addr "${BIND_ADDR}" \
      --ui-path "${REPO_DIR}/ui" \
      > "${LOG_FILE}" 2>&1 &
    disown
    pid=$!
    echo "$pid" > "${PID_FILE}"

    if ! wait_ready; then
      err "Server did not become healthy in 15 s."
      err "Last 20 log lines:"
      tail -20 "${LOG_FILE}" | sed 's/^/    /'
      exit 1
    fi
    info "Server healthy (PID $pid)."
  fi

  # Health summary
  local health
  health="$(curl -s "http://127.0.0.1:${HTTP_PORT}/health" 2>/dev/null || echo '{}')"
  echo -e "  ${C_D}${health}${C_0}"

  local host
  host="$(detect_ip)"
  print_urls "$host"

  if [[ "$open_gui" == "yes" ]]; then
    local url="http://${host}:${HTTP_PORT}/ui/index.html"
    info "Opening ${url}"
    open_browser "$url" || true
  fi
}

# ── Stop ─────────────────────────────────────────────────────────
cmd_stop() {
  local pid
  pid="$(running_pid)"
  if [[ -z "$pid" ]]; then
    warn "No sensing-server running on :${HTTP_PORT}."
    return 0
  fi
  info "Stopping sensing-server (PID $pid)…"
  kill "$pid" 2>/dev/null || true
  local tries=10
  while (( tries-- > 0 )); do
    kill -0 "$pid" 2>/dev/null || { info "Stopped."; rm -f "${PID_FILE}"; return 0; }
    sleep 0.3
  done
  warn "Forcing kill…"
  kill -9 "$pid" 2>/dev/null || true
  rm -f "${PID_FILE}"
}

# ── Status ───────────────────────────────────────────────────────
cmd_status() {
  local pid host
  pid="$(running_pid)"
  host="$(detect_ip)"
  if [[ -n "$pid" ]]; then
    info "Running — PID $pid"
    curl -s "http://127.0.0.1:${HTTP_PORT}/health" 2>/dev/null | sed 's/^/  /'
    echo
    print_urls "$host"
  else
    warn "Not running."
    echo "  Start with: ./start.sh"
  fi
}

# ── Dispatch ─────────────────────────────────────────────────────
case "${1:-start}" in
  ""|start)        cmd_start yes ;;
  no-browser|headless) cmd_start no ;;
  stop)            cmd_stop ;;
  restart)         cmd_stop; cmd_start yes ;;
  status)          cmd_status ;;
  logs)            mkdir -p "${LOG_DIR}"; touch "${LOG_FILE}"; tail -f "${LOG_FILE}" ;;
  build)           ensure_binary force ;;
  help|-h|--help)
    sed -n '2,20p' "$0" | sed 's/^# \{0,1\}//'
    ;;
  *)
    die "Unknown command: '$1'. Run './start.sh help' for usage."
    ;;
esac
