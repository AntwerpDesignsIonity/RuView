#!/usr/bin/env bash
# ╔══════════════════════════════════════════════════════════════════╗
# ║  RuView — single unified launcher                               ║
# ║  Ionity Global (Pty) Ltd  |  www.ionity.today                   ║
# ╚══════════════════════════════════════════════════════════════════╝
#
# Usage:
#   ./launch.sh              auto-detect source (ESP32 → WiFi → sim)
#   ./launch.sh esp32        force USB/CSI mode (ESP32 boards)
#   ./launch.sh wifi         neighbour-WiFi passive radar (no boards needed)
#   ./launch.sh sim          simulation mode (no hardware at all)
#   ./launch.sh stop         stop all running services
#   ./launch.sh status       show current service status
#   ./launch.sh logs         tail live logs (Ctrl-C to exit)
#   ./launch.sh install      run full dependency installer
#   ./launch.sh provision    interactive ESP32 flash & provision wizard
#
# Environment overrides (set in .env.local or pass inline):
#   HTTP_PORT   (default 3000)   WS_PORT (default 3001)   UDP_PORT (default 5005)
#   HUB_IP      (default: auto)  WIFI_SSID / WIFI_PASS

set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
IONITY="${REPO_DIR}/.ionity/ionity.sh"

# ── Colours ────────────────────────────────────────────────────────
BOLD=$'\033[1m'; CYAN=$'\033[36m'; GREEN=$'\033[32m'; YELLOW=$'\033[33m'
RED=$'\033[31m'; DIM=$'\033[2m'; RESET=$'\033[0m'

banner() {
  echo ""
  echo -e "${BOLD}${CYAN}╔══════════════════════════════════════════════╗${RESET}"
  echo -e "${BOLD}${CYAN}║  🛜  RuView WiFi Sensing Platform            ║${RESET}"
  echo -e "${BOLD}${CYAN}╚══════════════════════════════════════════════╝${RESET}"
  echo ""
}

die()  { echo -e "${RED}[ERROR]${RESET} $*" >&2; exit 1; }
info() { echo -e "${GREEN}[✓]${RESET} $*"; }
warn() { echo -e "${YELLOW}[!]${RESET} $*"; }

# ── Prereq check ───────────────────────────────────────────────────
[[ -f "${IONITY}" ]] || die "ionity.sh not found at ${IONITY} — is this the RuView repo?"

MODE="${1:-auto}"

# ── Dispatch ───────────────────────────────────────────────────────
case "${MODE}" in

  # ── Start modes ──────────────────────────────────────────────────
  auto)
    banner
    info "Source: auto-detect (ESP32 → WiFi → simulate)"
    IONITY_SKIP_PROVISION=1 exec "${IONITY}" run --yes
    ;;

  esp32)
    banner
    info "Source: ESP32 USB/CSI boards"
    echo -e "  ${DIM}Bridge reads /dev/ttyACM* serial at 460800 baud → UDP :5005${RESET}"
    echo ""
    IONITY_SOURCE=esp32 IONITY_SKIP_PROVISION=1 exec "${IONITY}" run --yes
    ;;

  wifi)
    banner
    info "Source: neighbour-WiFi passive radar (no ESP32 boards required)"
    echo -e "  ${DIM}Scans all visible BSSIDs as passive RF sensors${RESET}"
    echo -e "  ${DIM}Works with: any WiFi router or access point in range${RESET}"
    echo ""
    exec "${IONITY}" wifi
    ;;

  sim|simulate|simulation)
    banner
    info "Source: simulation (no hardware required)"
    echo -e "  ${DIM}Generates synthetic CSI data for UI development / demos${RESET}"
    echo ""
    IONITY_SOURCE=simulate IONITY_SKIP_PROVISION=1 exec "${IONITY}" run --yes
    ;;

  # ── Management ───────────────────────────────────────────────────
  stop)
    info "Stopping all RuView services…"
    "${IONITY}" stop
    info "All services stopped."
    ;;

  status)
    "${IONITY}" status
    ;;

  logs)
    LOG_DIR="${REPO_DIR}/logs"
    echo -e "${BOLD}Tailing logs (Ctrl-C to exit)…${RESET}"
    echo ""
    tail -f \
      "${LOG_DIR}/sensing-server.log" \
      "${LOG_DIR}/csi-bridge.log" \
      "${LOG_DIR}/autorepair.log" \
      "${LOG_DIR}/fault.log" 2>/dev/null
    ;;

  install)
    banner
    info "Running full installer…"
    exec "${REPO_DIR}/install.sh"
    ;;

  provision)
    banner
    info "Starting ESP32 provision wizard…"
    exec "${IONITY}" provision
    ;;

  help|-h|--help)
    banner
    echo -e "${BOLD}Usage:${RESET}  ./launch.sh [mode]"
    echo ""
    echo -e "  ${CYAN}(no arg)${RESET}     Auto-detect source (ESP32 → WiFi → sim)"
    echo -e "  ${CYAN}esp32${RESET}        Force USB ESP32 boards (CSI over serial)"
    echo -e "  ${CYAN}wifi${RESET}         Neighbour-WiFi passive radar — ${BOLD}no boards needed${RESET}"
    echo -e "  ${CYAN}sim${RESET}          Simulation — ${BOLD}no hardware at all${RESET}"
    echo -e "  ${CYAN}stop${RESET}         Stop all services"
    echo -e "  ${CYAN}status${RESET}       Show running services"
    echo -e "  ${CYAN}logs${RESET}         Tail live log output"
    echo -e "  ${CYAN}install${RESET}      Full dependency installer"
    echo -e "  ${CYAN}provision${RESET}    Flash & provision ESP32 boards"
    echo ""
    echo -e "${BOLD}Environment:${RESET}"
    echo -e "  HTTP_PORT   HTTP server port      (default: 3000)"
    echo -e "  WS_PORT     WebSocket port        (default: 3001)"
    echo -e "  UDP_PORT    ESP32 CSI UDP port    (default: 5005)"
    echo -e "  HUB_IP      Aggregator IP         (default: auto)"
    echo ""
    echo -e "${BOLD}Quick examples:${RESET}"
    echo -e "  # No hardware — just see the UI:"
    echo -e "  ${DIM}./launch.sh sim${RESET}"
    echo ""
    echo -e "  # Any WiFi router as sensor:"
    echo -e "  ${DIM}./launch.sh wifi${RESET}"
    echo ""
    echo -e "  # ESP32 boards via USB:"
    echo -e "  ${DIM}./launch.sh esp32${RESET}"
    echo ""
    ;;

  *)
    die "Unknown mode: '${MODE}'. Run './launch.sh help' for usage."
    ;;
esac
