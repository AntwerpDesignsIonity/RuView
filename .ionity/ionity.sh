#!/usr/bin/env bash
# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║                                                                              ║
# ║   ██╗ ██████╗ ███╗   ██╗██╗████████╗██╗   ██╗                               ║
# ║   ██║██╔═══██╗████╗  ██║██║╚══██╔══╝╚██╗ ██╔╝                               ║
# ║   ██║██║   ██║██╔██╗ ██║██║   ██║    ╚████╔╝                                ║
# ║   ██║██║   ██║██║╚██╗██║██║   ██║     ╚██╔╝                                 ║
# ║   ██║╚██████╔╝██║ ╚████║██║   ██║      ██║                                  ║
# ║   ╚═╝ ╚═════╝ ╚═╝  ╚═══╝╚═╝   ╚═╝      ╚═╝                                  ║
# ║                                                                              ║
# ║   Ionity RuView — Unified Stack Launcher                                     ║
# ║   Ionity Global (Pty) Ltd, South Africa                                      ║
# ║   ai@ionity.today  |  www.ionity.today  |  +27 646 999 877                   ║
# ║   Version: 1.1.0  |  April 2026                                              ║
# ║                                                                              ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
#
#  USAGE
#    ./.ionity/ionity.sh [COMMAND] [OPTIONS]
#
#  COMMANDS
#    run           Full stack startup (deps → build → serve)   [DEFAULT]
#    deps          Check and install all dependencies only
#    build         Build Rust binary and Python env only
#    docker        Start via Docker (no build required)
#    provision     Ionity Provisioner — flash and provision ESP32 nodes
#    status        Show status of all running services
#    stop          Stop all Ionity services
#    scripts       List and run the included Ionity scripts
#    help          Show this message
#
#  ENVIRONMENT (.env.local overrides all)
#    IONITY_MODE   native | docker | auto   (default: auto)
#    IONITY_SOURCE esp32 | simulate | auto  (default: auto)
#    HUB_IP        Aggregator IP            (default: detected)
#    HUB_PORT      UDP CSI port             (default: 5005)
#    HTTP_PORT     HTTP port                (default: 3000)
#    WS_PORT       WebSocket port           (default: 3001)
#
# ─────────────────────────────────────────────────────────────────────────────

set -euo pipefail
IFS=$'\n\t'

# ── Absolute paths ────────────────────────────────────────────────────────────
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_DIR="$(cd "${SCRIPT_DIR}/.." && pwd)"
IONITY_DIR="${SCRIPT_DIR}"
RUST_DIR="${REPO_DIR}/rust-port/wifi-densepose-rs"
BINARY="${RUST_DIR}/target/release/sensing-server"
UI_DIR="${REPO_DIR}/ui"
FW_DIR="${REPO_DIR}/firmware/esp32-csi-node"
SCRIPTS_DIR="${REPO_DIR}/scripts"
VENV_DIR="${REPO_DIR}/.venv"
LOG_DIR="${REPO_DIR}/logs"

# ── Colours ───────────────────────────────────────────────────────────────────
RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'
CYAN='\033[0;36m'; BLUE='\033[0;34m'; MAGENTA='\033[0;35m'
BOLD='\033[1m'; DIM='\033[2m'; RESET='\033[0m'

# ── Load .env.local ───────────────────────────────────────────────────────────
ENV_LOCAL="${REPO_DIR}/.env.local"
if [[ -f "${ENV_LOCAL}" ]]; then
  # shellcheck disable=SC1090
  set -a; source "${ENV_LOCAL}"; set +a
fi

# ── Defaults ──────────────────────────────────────────────────────────────────
IONITY_MODE="${IONITY_MODE:-auto}"
IONITY_SOURCE="${IONITY_SOURCE:-auto}"
HUB_IP="${HUB_IP:-$(hostname -I 2>/dev/null | awk '{print $1}')}"
HUB_PORT="${HUB_PORT:-5005}"
HTTP_PORT="${HTTP_PORT:-3000}"
WS_PORT="${WS_PORT:-3001}"

# ── Logging helpers ───────────────────────────────────────────────────────────
banner() {
  echo ""
  echo -e "${CYAN}${BOLD}╔══════════════════════════════════════════════════════════╗${RESET}"
  echo -e "${CYAN}${BOLD}║${RESET}  ${BOLD}$1${RESET}"
  echo -e "${CYAN}${BOLD}║${RESET}  ${DIM}Ionity Global (Pty) Ltd  |  ai@ionity.today  |  www.ionity.today${RESET}"
  echo -e "${CYAN}${BOLD}╚══════════════════════════════════════════════════════════╝${RESET}"
  echo ""
}
section() { echo -e "\n${BLUE}${BOLD}── $1 ──────────────────────────────────────────${RESET}"; }
log()     { echo -e "  ${CYAN}[Ionity]${RESET} $1"; }
ok()      { echo -e "  ${GREEN}${BOLD}✔${RESET} $1"; }
warn()    { echo -e "  ${YELLOW}${BOLD}⚠${RESET}  $1"; }
err()     { echo -e "  ${RED}${BOLD}✖${RESET} $1"; }
die()     { err "$1"; exit 1; }
step()    { echo -e "  ${MAGENTA}▶${RESET} $1"; }

# ── Source sub-modules ────────────────────────────────────────────────────────
# shellcheck disable=SC1091
source "${IONITY_DIR}/lib/deps.sh"
# shellcheck disable=SC1091
source "${IONITY_DIR}/lib/build.sh"
# shellcheck disable=SC1091
source "${IONITY_DIR}/lib/stack.sh"
# shellcheck disable=SC1091
source "${IONITY_DIR}/lib/provisioner.sh"
# shellcheck disable=SC1091
source "${IONITY_DIR}/lib/scripts_menu.sh"

# ── Header ────────────────────────────────────────────────────────────────────
print_header() {
  clear 2>/dev/null || true
  echo -e "${CYAN}${BOLD}"
  cat <<'LOGO'
  ██╗ ██████╗ ███╗   ██╗██╗████████╗██╗   ██╗
  ██║██╔═══██╗████╗  ██║██║╚══██╔══╝╚██╗ ██╔╝
  ██║██║   ██║██╔██╗ ██║██║   ██║    ╚████╔╝
  ██║██║   ██║██║╚██╗██║██║   ██║     ╚██╔╝
  ██║╚██████╔╝██║ ╚████║██║   ██║      ██║
  ╚═╝ ╚═════╝ ╚═╝  ╚═══╝╚═╝   ╚═╝      ╚═╝
LOGO
  echo -e "${RESET}"
  echo -e "  ${BOLD}Ionity RuView — Unified Stack Launcher${RESET}"
  echo -e "  ${DIM}Ionity Global (Pty) Ltd  |  ai@ionity.today  |  www.ionity.today${RESET}"
  echo -e "  ${DIM}Hub: ${HUB_IP}  |  HTTP: ${HTTP_PORT}  |  WS: ${WS_PORT}  |  UDP: ${HUB_PORT}${RESET}"
  echo ""
}

# ── COMMAND: run (full stack) ─────────────────────────────────────────────────
cmd_run() {
  print_header
  banner "Full Stack Startup"
  mkdir -p "${LOG_DIR}"

  # Parse run sub-flags
  while [[ $# -gt 0 ]]; do
    case "$1" in
      --skip-provision) export IONITY_SKIP_PROVISION=1; shift ;;
      --open-gui)       export IONITY_OPEN_GUI=1; shift ;;
      --yes|-y)         export IONITY_SKIP_PROVISION=1; export IONITY_OPEN_GUI=1; shift ;;
      *) shift ;;
    esac
  done

  section "1/4  Dependency Check"
  ionity_check_deps

  section "2/4  Build"
  ionity_build

  section "3/4  Start Stack"
  ionity_start_stack

  section "4/4  Ready"
  echo ""
  ok "Ionity RuView is running"
  echo ""
  echo -e "  ${BOLD}Web UI${RESET}      →  ${CYAN}http://${HUB_IP}:${HTTP_PORT}/ui/index.html${RESET}"
  echo -e "  ${BOLD}Health${RESET}      →  ${CYAN}http://${HUB_IP}:${HTTP_PORT}/health${RESET}"
  echo -e "  ${BOLD}WebSocket${RESET}   →  ${CYAN}ws://${HUB_IP}:${WS_PORT}/ws/sensing${RESET}"
  echo -e "  ${BOLD}ESP32 UDP${RESET}   →  ${CYAN}udp://${HUB_IP}:${HUB_PORT}${RESET}"
  echo ""
  if [[ "${IONITY_SOURCE}" == "wifi" ]]; then
    ok "Neighbour-WiFi mode active: scanning all visible BSSIDs as passive RF sensors"
    echo -e "  ${DIM}Set IONITY_SOURCE=esp32 in .env.local to switch to hardware CSI mode${RESET}"
  elif [[ "${IONITY_SOURCE}" == "auto" ]] || [[ -z "${IONITY_SOURCE}" ]]; then
    log "Source: auto-detect (ESP32 CSI → neighbour-WiFi → simulate)"
    echo -e "  ${DIM}To force neighbour-WiFi scan: set IONITY_SOURCE=wifi in .env.local${RESET}"
  fi
  echo ""
  echo -e "  ${DIM}Logs: ${LOG_DIR}/sensing-server.log${RESET}"
  echo -e "  ${DIM}Manage: ./.ionity/ionity.sh services${RESET}"
  echo ""

  # Auto-open GUIs if not already opened during server start
  if [[ "${IONITY_OPEN_GUI:-0}" == "1" ]]; then
    ionity_open_guis
  fi

  # Drop into the interactive service manager (skip when no TTY or --open-gui headless)
  if [[ -t 0 ]]; then
    ionity_services
  else
    log "Non-interactive mode — server running in background. Use '.ionity/ionity.sh services' to manage."
    # Keep script alive so nohup/background callers see success
    wait
  fi
}

# ── COMMAND: wifi (neighbour-WiFi mode) ──────────────────────────────────────
cmd_wifi() {
  print_header
  export IONITY_SOURCE="wifi"
  banner "Ionity — Neighbour-WiFi Sensing Mode"
  step "Scanning all visible BSSIDs (own AP + neighbours) as passive RF sensors…"
  log "Set IONITY_SOURCE=wifi in .env.local to make this the default."
  echo ""
  ionity_check_deps
  ionity_build
  ionity_start_stack
  echo ""
  ok "Neighbour-WiFi sensing active"
  echo -e "  ${BOLD}Source${RESET} → wifi (multi-BSSID passive radar)"
  echo -e "  ${BOLD}UI${RESET}     → ${CYAN}http://${HUB_IP}:${HTTP_PORT}/ui/index.html${RESET}"
  echo ""
}

# ── COMMAND: deps ─────────────────────────────────────────────────────────────
cmd_deps() {
  print_header
  banner "Ionity — Dependency Check & Install"
  ionity_check_deps
  echo ""
  ok "All dependencies satisfied."
  echo ""
}

# ── COMMAND: build ────────────────────────────────────────────────────────────
cmd_build() {
  print_header
  banner "Ionity — Build"
  ionity_check_deps
  ionity_build
  echo ""
  ok "Build complete."
  echo ""
}

# ── COMMAND: docker ───────────────────────────────────────────────────────────
cmd_docker() {
  print_header
  banner "Ionity — Docker Stack"
  ionity_start_docker
}

# ── COMMAND: provision ────────────────────────────────────────────────────────
cmd_provision() {
  print_header
  banner "Ionity Provisioner — ESP32 Node Flash & Provision"
  ionity_provisioner "$@"
}

# ── COMMAND: auto-provision ───────────────────────────────────────────────────
cmd_auto_provision() {
  print_header
  banner "Ionity Auto-Provision — Mesh Role Assignment"
  ionity_provisioner auto "$@"
}
# ── COMMAND: scan (board detect + provision offer) ───────────────────
cmd_scan() {
  print_header
  banner "Ionity — ESP32 Board Scanner"
  ionity_scan_boards "$@"
}
# ── COMMAND: status ───────────────────────────────────────────────────────────
cmd_status() {
  print_header
  banner "Ionity — Service Status"
  ionity_status
}

# ── COMMAND: stop ─────────────────────────────────────────────────────────────
cmd_stop() {
  banner "Ionity — Stopping All Services"
  ionity_stop
}

# ── COMMAND: autorepair ───────────────────────────────────────────────────────
cmd_autorepair() {
  print_header
  banner "Ionity — Autorepair Supervisor"
  ionity_autorepair "$@"
}

# ── COMMAND: services (interactive manager) ───────────────────────────────────
cmd_services() {
  print_header
  banner "Ionity — Service Manager"
  ionity_services "$@"
}

# ── COMMAND: scripts ──────────────────────────────────────────────────────────
cmd_scripts() {
  print_header
  banner "Ionity — Script Library"
  ionity_scripts_menu "$@"
}

# ── COMMAND: help ─────────────────────────────────────────────────────────────
cmd_help() {
  print_header
  echo -e "${BOLD}USAGE${RESET}"
  echo -e "  ./.ionity/ionity.sh [COMMAND] [OPTIONS]"
  echo ""
  echo -e "${BOLD}COMMANDS${RESET}"
  echo -e "  ${GREEN}run${RESET}          Full stack: deps → build → serve  ${DIM}[default]${RESET}"
  echo -e "  ${GREEN}wifi${RESET}         Neighbour-WiFi mode: scan all BSSIDs as passive RF sensors"
  echo -e "  ${GREEN}deps${RESET}         Check and install all dependencies"
  echo -e "  ${GREEN}build${RESET}        Build Rust binary + Python venv"
  echo -e "  ${GREEN}docker${RESET}       Start via Docker (skip native build)"
  echo -e "  ${GREEN}provision${RESET}    Ionity Provisioner — flash and configure ESP32 nodes"
  echo -e "  ${GREEN}auto-provision${RESET} Auto-detect boards, assign TX/RX roles, flash mesh"
  echo -e "  ${GREEN}scan${RESET}         Detect connected ESP32 boards + offer to provision"
  echo -e "  ${GREEN}services${RESET}     Interactive service manager (start/stop/restart menu)"
  echo -e "  ${GREEN}status${RESET}       Show running services and health"
  echo -e "  ${GREEN}stop${RESET}         Stop all Ionity services"
  echo -e "  ${GREEN}autorepair${RESET}   Watchdog supervisor: monitor + auto-restart on failure"
  echo -e "  ${GREEN}scripts${RESET}      Browse and run the Ionity script library"
  echo -e "  ${GREEN}help${RESET}         Show this help"
  echo ""
  echo -e "${BOLD}ENVIRONMENT  ${DIM}(set in .env.local)${RESET}"
  echo -e "  IONITY_MODE    native | docker | auto   (default: auto)"
  echo -e "  IONITY_SOURCE  esp32 | simulate | auto  (default: auto)"
  echo -e "  HUB_IP         Aggregator LAN IP"
  echo -e "  HUB_PORT       UDP CSI port             (default: 5005)"
  echo -e "  HTTP_PORT      HTTP port                (default: 3000)"
  echo -e "  WS_PORT        WebSocket port           (default: 3001)"
  echo ""
  echo -e "${BOLD}EXAMPLES${RESET}"
  echo -e "  ./.ionity/ionity.sh                  # full auto startup"
  echo -e "  ./.ionity/ionity.sh provision        # interactive ESP32 provisioner"
  echo -e "  ./.ionity/ionity.sh provision /dev/ttyACM0 2  # flash node 2 on ACM0"
  echo -e "  ./.ionity/ionity.sh docker           # Docker-only mode"
  echo -e "  ./.ionity/ionity.sh scripts          # browse script library"
  echo -e "  ./.ionity/ionity.sh stop             # stop everything"
  echo ""
  echo -e "  ${DIM}Ionity Global (Pty) Ltd  |  ai@ionity.today  |  www.ionity.today${RESET}"
  echo ""
}

# ── Activate Python venv (always, before any command) ────────────────────────
# This ensures python3, esptool, pyserial etc. from the project venv are
# available for ALL sub-commands (provision, scripts, build, run…) without
# requiring the user to manually activate first.
if [[ -f "${VENV_DIR}/bin/activate" ]]; then
  # shellcheck disable=SC1090
  source "${VENV_DIR}/bin/activate"
elif [[ "${1:-}" != "deps" && "${1:-}" != "help" && "${1:-}" != "--help" && "${1:-}" != "-h" ]]; then
  echo -e "  ${YELLOW}⚠${RESET}  Python venv not found at ${VENV_DIR}"
  echo -e "     Run '.ionity/ionity.sh deps' to create and populate it."
  echo -e "     Continuing with system Python…"
fi

# ── Entrypoint ────────────────────────────────────────────────────────────────
COMMAND="${1:-run}"
shift 2>/dev/null || true

case "${COMMAND}" in
  run)        cmd_run "$@" ;;
  wifi)       cmd_wifi "$@" ;;
  deps)       cmd_deps "$@" ;;
  build)      cmd_build "$@" ;;
  docker)     cmd_docker "$@" ;;
  provision)  cmd_provision "$@" ;;
  auto-provision) cmd_auto_provision "$@" ;;
  scan)       cmd_scan "$@" ;;
  services)   cmd_services "$@" ;;
  status)     cmd_status "$@" ;;
  stop)       cmd_stop "$@" ;;
  autorepair) cmd_autorepair "$@" ;;
  scripts)    cmd_scripts "$@" ;;
  help|--help|-h) cmd_help ;;
  *)
    err "Unknown command: ${COMMAND}"
    echo ""
    cmd_help
    exit 1
    ;;
esac
