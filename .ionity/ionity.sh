#!/usr/bin/env bash
# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  ionity.sh — IONITY RuView CLI Entrypoint                                   ║
# ║  Ionity Global (Pty) Ltd  |  ai@ionity.today  |  www.ionity.today           ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
set -euo pipefail

# ── Resolve repo root (script lives in .ionity/) ─────────────────────────────
IONITY_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_DIR="$(cd "${IONITY_DIR}/.." && pwd)"

# ── Colour helpers ────────────────────────────────────────────────────────────
if [[ -t 1 ]]; then
  BOLD='\033[1m'; DIM='\033[2m'; RESET='\033[0m'
  RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[0;33m'
  CYAN='\033[0;36m'; BLUE='\033[0;34m'
else
  BOLD=''; DIM=''; RESET=''
  RED=''; GREEN=''; YELLOW=''; CYAN=''; BLUE=''
fi

ok()      { echo -e "  ${GREEN}✔${RESET}  $*"; }
warn()    { echo -e "  ${YELLOW}⚠${RESET}  $*"; }
err()     { echo -e "  ${RED}✖${RESET}  $*" >&2; }
step()    { echo -e "  ${CYAN}→${RESET}  $*"; }
log()     { echo -e "     $*"; }
section() { echo -e "\n${BOLD}${BLUE}══  $*  ══${RESET}"; }
die()     { err "$*"; exit 1; }

# ── Default configuration ─────────────────────────────────────────────────────
: "${HTTP_PORT:=3000}"        # UI/API port (matches existing ui/start-ui.sh convention)
: "${WS_PORT:=8765}"          # WebSocket port
: "${HUB_PORT:=5005}"         # UDP CSI listener port
: "${HUB_IP:=localhost}"      # Bind/advertise address
: "${IONITY_SOURCE:=auto}"    # Data source: auto | esp32 | simulated | mmfi
: "${IONITY_MODE:=auto}"      # Runtime mode: auto | native | docker
: "${IONITY_OPEN_GUI:=0}"     # Auto-open browser: 0 | 1
: "${IONITY_SKIP_PROVISION:=0}" # Skip ESP32 scan/provision: 0 | 1
: "${IONITY_TRACE:=info,tower_http=debug}" # RUST_LOG filter for tracing output

# ── Derived paths ─────────────────────────────────────────────────────────────
RUST_DIR="${REPO_DIR}/rust-port/wifi-densepose-rs"
BINARY="${RUST_DIR}/target/release/sensing-server"
UI_DIR="${REPO_DIR}/ui"
LOG_DIR="${REPO_DIR}/logs"
LOG_BUILD_DIR="${LOG_DIR}/build"
LOG_FLASH_DIR="${LOG_DIR}/flash"
LOG_PROVISION_DIR="${LOG_DIR}/provision"
LOG_RUNTIME_DIR="${LOG_DIR}/runtime"
LOG_HARDWARE_DIR="${LOG_DIR}/hardware"
LOG_SESSION_DIR="${LOG_DIR}/session"
FW_DIR="${REPO_DIR}/firmware/esp32-csi-node"
VENV_DIR="${REPO_DIR}/.venv"
SCRIPTS_DIR="${REPO_DIR}/scripts"

export REPO_DIR RUST_DIR BINARY UI_DIR LOG_DIR FW_DIR VENV_DIR SCRIPTS_DIR
export LOG_BUILD_DIR LOG_FLASH_DIR LOG_PROVISION_DIR LOG_RUNTIME_DIR LOG_HARDWARE_DIR LOG_SESSION_DIR
export HTTP_PORT WS_PORT HUB_PORT HUB_IP IONITY_SOURCE IONITY_MODE
export IONITY_OPEN_GUI IONITY_SKIP_PROVISION IONITY_TRACE
export BOLD DIM RESET RED GREEN YELLOW CYAN BLUE

# ── Source lib modules ────────────────────────────────────────────────────────
LIB_DIR="${IONITY_DIR}/lib"
# shellcheck source=lib/stack.sh
source "${LIB_DIR}/stack.sh"
# shellcheck source=lib/build.sh
source "${LIB_DIR}/build.sh"
# shellcheck source=lib/deps.sh
source "${LIB_DIR}/deps.sh"
# shellcheck source=lib/provisioner.sh
source "${LIB_DIR}/provisioner.sh"
# shellcheck source=lib/scripts_menu.sh
source "${LIB_DIR}/scripts_menu.sh"

# ── Parse global flags before the sub-command ────────────────────────────────
_parse_flags() {
  while [[ $# -gt 0 ]]; do
    case "$1" in
      --skip-provision)   IONITY_SKIP_PROVISION=1; shift ;;
      --open-gui)         IONITY_OPEN_GUI=1;       shift ;;
      --source=*)         IONITY_SOURCE="${1#*=}";  shift ;;
      --source)           IONITY_SOURCE="${2:-auto}"; shift 2 ;;
      --port=*)           HTTP_PORT="${1#*=}";       shift ;;
      --port)             HTTP_PORT="${2:-3000}";    shift 2 ;;
      --mode=*)           IONITY_MODE="${1#*=}";     shift ;;
      --mode)             IONITY_MODE="${2:-auto}";  shift 2 ;;
      --native)           IONITY_MODE=native;        shift ;;
      --docker)           IONITY_MODE=docker;        shift ;;
      --trace)            IONITY_TRACE=debug;         shift ;;
      --log-level=*)      IONITY_TRACE="${1#*=}";    shift ;;
      --log-level)        IONITY_TRACE="${2:-debug}"; shift 2 ;;
      *)                  break ;;
    esac
  done
  # Re-export after parse
  export HTTP_PORT WS_PORT HUB_PORT IONITY_SOURCE IONITY_MODE IONITY_OPEN_GUI IONITY_SKIP_PROVISION IONITY_TRACE
}

# ── Print usage ───────────────────────────────────────────────────────────────
_usage() {
  cat <<EOF

  ${BOLD}ionity.sh${RESET} — IONITY RuView CLI

  ${BOLD}USAGE${RESET}
    .ionity/ionity.sh <command> [flags]

  ${BOLD}COMMANDS${RESET}
    run           Start the sensing server (default)
    build         Build the Rust sensing server + install Python deps
    stop          Stop the sensing server
    restart       Stop then start
    status        Show server, ports, and board status
    gui           Open all dashboards in the browser
    services      Interactive service manager
    provision     Flash/provision ESP32 nodes
    scan          Scan USB ports for connected ESP32 boards
    scripts       Interactive script library menu
    deps          Check and install all dependencies
    logs          Tail sensing-server logs
    help          Show this help

  ${BOLD}FLAGS${RESET}
    --skip-provision    Skip ESP32 board scan on startup
    --open-gui          Auto-open browsers after server starts
    --source <src>      Data source: auto | esp32 | simulated | mmfi
    --port <port>       HTTP port (default: ${HTTP_PORT})
    --mode native|docker  Runtime mode (default: auto)
    --native            Shorthand for --mode=native
    --docker            Shorthand for --mode=docker
    --trace             Enable debug-level tracing (sets RUST_LOG=debug)
    --log-level <lvl>   Set RUST_LOG filter (e.g. trace, debug, info, warn, error)

  ${BOLD}EXAMPLES${RESET}
    .ionity/ionity.sh run
    .ionity/ionity.sh run --skip-provision --open-gui
    .ionity/ionity.sh run --source simulated --open-gui
    .ionity/ionity.sh build
    .ionity/ionity.sh stop
    .ionity/ionity.sh status
    .ionity/ionity.sh provision
    .ionity/ionity.sh scripts
    .ionity/ionity.sh logs

EOF
}

# ── Banner ────────────────────────────────────────────────────────────────────
_banner() {
  echo ""
  echo -e "${CYAN}${BOLD}  ╔═══════════════════════════════════════════════════╗${RESET}"
  echo -e "${CYAN}${BOLD}  ║   π  IONITY  RuView  —  WiFi DensePose System    ║${RESET}"
  echo -e "${CYAN}${BOLD}  ╚═══════════════════════════════════════════════════╝${RESET}"
  echo -e "     ${DIM}http://${HUB_IP}:${HTTP_PORT}/ui/index.html${RESET}"
  echo ""
}

# ── Open GUI helper (update to include monitor.html) ─────────────────────────
_open_monitor_gui() {
  local base="http://${HUB_IP}:${HTTP_PORT}"
  _ionity_open_browser "${base}/ui/monitor.html" && ok "Monitor Hub" || true
}

# ── Main dispatch ─────────────────────────────────────────────────────────────
CMD="${1:-run}"
shift 2>/dev/null || true
_parse_flags "$@"

case "${CMD}" in
  run|start)
    _banner
    section "Starting IONITY RuView"
    ionity_start_stack
    ;;

  build)
    _banner
    section "Build"
    ionity_check_deps
    ionity_build
    ;;

  stop)
    section "Stopping Services"
    ionity_stop
    ;;

  restart)
    section "Restarting"
    ionity_stop
    sleep 1
    ionity_start_stack
    ;;

  status)
    _banner
    ionity_status
    ;;

  gui|open)
    ionity_open_guis
    _open_monitor_gui
    ;;

  services|svc)
    _banner
    ionity_services
    ;;

  provision|flash|prov)
    _banner
    ionity_provisioner "$@"
    ;;

  scan)
    _banner
    section "ESP32 Board Scan"
    ionity_scan_boards
    ;;

  scripts|script|run-script)
    _banner
    ionity_scripts_menu "$@"
    ;;

  deps|check-deps|install)
    _banner
    section "Dependency Check"
    ionity_check_deps
    ;;

  logs|log)
    local log_file="${LOG_DIR}/sensing-server.log"
    if [[ -f "${log_file}" ]]; then
      step "Tailing ${log_file} (Ctrl-C to stop)…"
      tail -f "${log_file}"
    else
      warn "Log not found: ${log_file}"
      warn "Start the server first with: .ionity/ionity.sh run"
      exit 1
    fi
    ;;

  help|--help|-h)
    _usage
    ;;

  *)
    err "Unknown command: ${CMD}"
    _usage
    exit 1
    ;;
esac
