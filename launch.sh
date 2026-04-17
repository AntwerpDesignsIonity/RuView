#!/usr/bin/env bash
# ╔══════════════════════════════════════════════════════════════════╗
# ║  AEDI-S — single unified launcher                               ║
# ║  Ionity Global (Pty) Ltd  |  www.ionity.today                   ║
# ╚══════════════════════════════════════════════════════════════════╝
#
# Usage:
#   ./launch.sh              auto-detect source (ESP32 → WiFi → sim)
#   ./launch.sh esp32        ESP32 mode (default transport: USB serial bridge)
#   ./launch.sh esp32 usb    ESP32 via USB serial bridge → local UDP
#   ./launch.sh esp32 wifi   ESP32 direct WiFi UDP → hub
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
  echo -e "${BOLD}${CYAN}║  🛜  AEDI-S WiFi Sensing Platform            ║${RESET}"
  echo -e "${BOLD}${CYAN}╚══════════════════════════════════════════════╝${RESET}"
  echo ""
}

die()  { echo -e "${RED}[ERROR]${RESET} $*" >&2; exit 1; }
info() { echo -e "${GREEN}[✓]${RESET} $*"; }
warn() { echo -e "${YELLOW}[!]${RESET} $*"; }

# ── Prereq check ───────────────────────────────────────────────────
[[ -f "${IONITY}" ]] || die "ionity.sh not found at ${IONITY} — is this the AEDI-S repo?"

MODE="${1:-auto}"
ESP32_LINK_MODE="${2:-usb}"

# ── Dispatch ───────────────────────────────────────────────────────
case "${MODE}" in

  # ── Start modes ──────────────────────────────────────────────────
  auto)
    banner
    info "Source: auto-detect (ESP32 → WiFi → simulate)"
    IONITY_SKIP_PROVISION_PROMPT=1 exec "${IONITY}" run --yes
    ;;

  esp32)
    case "${ESP32_LINK_MODE}" in
      usb)
        banner
        info "Source: ESP32 boards over USB serial bridge"
        echo -e "  ${DIM}Bridge reads /dev/ttyACM* serial at 460800 baud → UDP :5005${RESET}"
        echo ""
        IONITY_SOURCE=esp32 IONITY_ESP32_LINK=usb IONITY_SKIP_PROVISION_PROMPT=1 exec "${IONITY}" run --yes
        ;;
      wifi)
        banner
        info "Source: ESP32 boards over WiFi UDP"
        echo -e "  ${DIM}ESP32 nodes must be provisioned to send CSI to this hub on UDP :5005${RESET}"
        echo ""
        IONITY_SOURCE=esp32 IONITY_ESP32_LINK=wifi IONITY_SKIP_PROVISION_PROMPT=1 exec "${IONITY}" run --yes
        ;;
      *)
        die "Unknown ESP32 transport: '${ESP32_LINK_MODE}'. Use './launch.sh esp32 usb' or './launch.sh esp32 wifi'."
        ;;
    esac
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
  autorepair|watchdog)
    banner
    info "Starting autorepair watchdog (monitors sensing-server, CSI bridge, bridge-controller)"
    info "Checks every 15 s — restarts any offline service automatically"
    echo -e "  ${DIM}Logs: logs/autorepair.log  |  Faults: logs/fault.log${RESET}"
    echo ""
    exec "${IONITY}" autorepair
    ;;
  stop)
    info "Stopping all AEDI-S services…"
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
    echo -e "  ${CYAN}esp32 usb${RESET}    ESP32 boards via USB serial bridge"
    echo -e "  ${CYAN}esp32 wifi${RESET}   ESP32 boards via direct WiFi UDP"
    echo -e "  ${CYAN}wifi${RESET}         Neighbour-WiFi passive radar — ${BOLD}no boards needed${RESET}"
    echo -e "  ${CYAN}sim${RESET}          Simulation — ${BOLD}no hardware at all${RESET}"
    echo -e "  ${CYAN}stop${RESET}         Stop all services"
    echo -e "  ${CYAN}status${RESET}       Show running services"
    echo -e "  ${CYAN}logs${RESET}         Tail live log output"
    echo -e "  ${CYAN}autorepair${RESET}   Start watchdog (auto-restart on failure, 15s checks)"
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
    echo -e "  # ESP32 boards via USB serial bridge:"
    echo -e "  ${DIM}./launch.sh esp32 usb${RESET}"
    echo ""
    echo -e "  # ESP32 boards via direct WiFi UDP:"
    echo -e "  ${DIM}./launch.sh esp32 wifi${RESET}"
    echo ""
    ;;

  *)
    die "Unknown mode: '${MODE}'. Run './launch.sh help' for usage."
    ;;
esac
