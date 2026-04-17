#!/usr/bin/env bash
# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  log-monitor.sh — AEDI-S Stack Log Monitor                                  ║
# ║  Ionity (Pty) Ltd  |  ai@ionity.today  |  www.ionity.today                  ║
# ║                                                                              ║
# ║  Monitors: sensing-server · AEDI (local AI) · csi-bridge · autorepair       ║
# ║                                                                              ║
# ║  Usage: ./scripts/log-monitor.sh [--lines N] [--filter SERVICE]             ║
# ║    --lines N        Tail N historical lines per log on start (default: 30)  ║
# ║    --filter SERVICE Only show: server | aedi | bridge | autorepair | fault  ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
LOGS_DIR="${REPO_DIR}/logs"

# ── Colours ──────────────────────────────────────────────────────────────────
BOLD='\033[1m';   DIM='\033[2m';    RESET='\033[0m'
RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[0;33m'
BLUE='\033[0;34m'; MAGENTA='\033[0;35m'; CYAN='\033[0;36m'; WHITE='\033[0;37m'
RED_B='\033[1;31m'; GREEN_B='\033[1;32m'; YELLOW_B='\033[1;33m'
CYAN_B='\033[1;36m'; MAGENTA_B='\033[1;35m'

# Per-service label colours
C_SERVER="${CYAN_B}"
C_AEDI="${MAGENTA_B}"
C_BRIDGE="${YELLOW_B}"
C_AUTOREPAIR="${BLUE}"
C_FAULT="${RED_B}"

# ── Args ─────────────────────────────────────────────────────────────────────
TAIL_LINES=30
FILTER=""

while [[ $# -gt 0 ]]; do
  case "$1" in
    --lines)   TAIL_LINES="${2:-30}"; shift 2 ;;
    --filter)  FILTER="${2:-}";       shift 2 ;;
    -h|--help)
      sed -n '2,12p' "$0" | sed 's/^# //' | sed 's/^#//'
      exit 0 ;;
    *) echo "Unknown option: $1" >&2; exit 1 ;;
  esac
done

# ── Log file paths ────────────────────────────────────────────────────────────
LOG_SERVER="${LOGS_DIR}/sensing-server.log"
LOG_AEDI="${LOGS_DIR}/aedi.log"
LOG_BRIDGE="${LOGS_DIR}/csi-bridge.log"
LOG_AUTOREPAIR="${LOGS_DIR}/autorepair.log"
LOG_FAULT="${LOGS_DIR}/fault.log"

PID_SERVER="${LOGS_DIR}/sensing-server.pid"
PID_AEDI="${LOGS_DIR}/aedi.pid"
PID_BRIDGE="${LOGS_DIR}/csi-bridge.pid"

# ── Helpers ──────────────────────────────────────────────────────────────────
ok()   { echo -e "  ${GREEN}✔${RESET}  $*"; }
warn() { echo -e "  ${YELLOW}⚠${RESET}  $*"; }
err()  { echo -e "  ${RED}✖${RESET}  $*"; }
dim()  { echo -e "  ${DIM}$*${RESET}"; }

pid_status() {
  local label="$1" pid_file="$2"
  if [[ -f "${pid_file}" ]]; then
    local pid; pid=$(cat "${pid_file}" 2>/dev/null || echo "")
    if [[ -n "${pid}" ]] && kill -0 "${pid}" 2>/dev/null; then
      echo -e "  ${GREEN}●${RESET} ${label} ${DIM}(PID ${pid})${RESET}"
    else
      echo -e "  ${RED}●${RESET} ${label} ${DIM}(dead — stale PID ${pid})${RESET}"
    fi
  else
    echo -e "  ${YELLOW}●${RESET} ${label} ${DIM}(no PID file)${RESET}"
  fi
}

health_check() {
  local label="$1" url="$2" color="$3"
  local status
  if status=$(curl -sf --max-time 2 "${url}" 2>/dev/null); then
    echo -e "  ${GREEN}↑${RESET} ${color}${label}${RESET} ${DIM}${url}${RESET} ${GREEN}online${RESET}"
  else
    echo -e "  ${RED}↓${RESET} ${color}${label}${RESET} ${DIM}${url}${RESET} ${RED}unreachable${RESET}"
  fi
}

# ── Severity colorizer (applied inside each tail pipe) ───────────────────────
# Input: raw log line → colored output
colorize_line() {
  # Called with: colorize_line <PREFIX_COLOR> <LABEL>
  # Reads from stdin via awk
  local prefix_color="$1" label="$2"
  awk -v pc="${prefix_color}" -v lbl="${label}" \
      -v RST="${RESET}" -v DIM="${DIM}" \
      -v RED="${RED_B}" -v YLW="${YELLOW}" \
      -v GRN="${GREEN}" -v WHT="${WHITE}" '
  {
    prefix = pc "[" lbl "] " RST

    # Severity detection (case-insensitive)
    line = $0
    lo  = tolower(line)

    # Strip ANSI escapes for matching, but keep original for display
    gsub(/\033\[[0-9;]*m/, "", lo)

    if (lo ~ /error|fatal|crash|panic|fail|fault|killed|segfault/) {
      print prefix RED line RST
    } else if (lo ~ /warn|unhealthy|unreachable|dead|stale|silent|miss/) {
      print prefix YLW line RST
    } else if (lo ~ /ok|success|started|online|ready|restarted|pass|✔/) {
      print prefix GRN line RST
    } else if (lo ~ /debug|trace/) {
      print prefix DIM line RST
    } else {
      print prefix WHT line RST
    }
    fflush()
  }'
}
export -f colorize_line

# ── Banner ───────────────────────────────────────────────────────────────────
clear
echo ""
echo -e "${CYAN}${BOLD}  ╔══════════════════════════════════════════════════════════╗${RESET}"
echo -e "${CYAN}${BOLD}  ║  ⬡  AEDI-S — Stack Log Monitor             $(date '+%H:%M:%S')  ║${RESET}"
echo -e "${CYAN}${BOLD}  ╚══════════════════════════════════════════════════════════╝${RESET}"
echo ""

# ── Service status ────────────────────────────────────────────────────────────
echo -e "  ${BOLD}Process Status${RESET}"
pid_status "sensing-server" "${PID_SERVER}"
pid_status "AEDI (local AI)" "${PID_AEDI}"
pid_status "csi-bridge     " "${PID_BRIDGE}"
echo ""

echo -e "  ${BOLD}Health Endpoints${RESET}"
health_check "sensing-server" "http://localhost:3000/health"       "${C_SERVER}"
health_check "AEDI           " "http://localhost:3002/health"       "${C_AEDI}"
health_check "Ollama         " "http://localhost:11434/api/version" "${C_AEDI}"
echo ""

# ── Legend ────────────────────────────────────────────────────────────────────
echo -e "  ${BOLD}Legend${RESET}"
echo -e "  ${C_SERVER}[SERVER]${RESET}     sensing-server  (HTTP :3000, WS :3001, UDP :5005)"
echo -e "  ${C_AEDI}[AEDI]${RESET}       local AI        (Ollama/Gemma3 :3002)"
echo -e "  ${C_BRIDGE}[BRIDGE]${RESET}     csi-bridge      (ESP32 serial → UDP)"
echo -e "  ${C_AUTOREPAIR}[AUTOREPAIR]${RESET} watchdog        (auto-restart supervisor)"
echo -e "  ${C_FAULT}[FAULT]${RESET}      fault events    (structured JSON audit log)"
echo ""
echo -e "  ${DIM}Tailing last ${TAIL_LINES} lines · Ctrl-C to exit${RESET}"
echo -e "  ${DIM}──────────────────────────────────────────────────────────${RESET}"
echo ""

# ── Build list of (log_file, color, label) to monitor ────────────────────────
declare -a LOG_FILES=()
declare -a LOG_COLORS=()
declare -a LOG_LABELS=()

add_log() {
  local f="$1" c="$2" l="$3"
  [[ -n "${FILTER}" ]] && [[ "${l,,}" != *"${FILTER,,}"* ]] && return
  LOG_FILES+=("${f}")
  LOG_COLORS+=("${c}")
  LOG_LABELS+=("${l}")
}

add_log "${LOG_SERVER}"    "${C_SERVER}"     "SERVER    "
add_log "${LOG_AEDI}"      "${C_AEDI}"       "AEDI      "
add_log "${LOG_BRIDGE}"    "${C_BRIDGE}"     "BRIDGE    "
add_log "${LOG_AUTOREPAIR}" "${C_AUTOREPAIR}" "AUTOREPAIR"
add_log "${LOG_FAULT}"     "${C_FAULT}"      "FAULT     "

if [[ ${#LOG_FILES[@]} -eq 0 ]]; then
  err "No logs matched filter '${FILTER}'"
  exit 1
fi

# ── Kill tracker for cleanup ──────────────────────────────────────────────────
TAIL_PIDS=()

cleanup() {
  echo ""
  echo -e "  ${DIM}Stopping log monitor…${RESET}"
  for pid in "${TAIL_PIDS[@]:-}"; do
    kill "${pid}" 2>/dev/null || true
  done
  exit 0
}
trap cleanup SIGINT SIGTERM

# ── Spawn one tail per log file ───────────────────────────────────────────────
for i in "${!LOG_FILES[@]}"; do
  log_file="${LOG_FILES[$i]}"
  color="${LOG_COLORS[$i]}"
  label="${LOG_LABELS[$i]}"

  if [[ ! -f "${log_file}" ]]; then
    echo -e "  ${DIM}[${label// /}] log not found — will watch for creation: ${log_file}${RESET}"
    # Use tail on /dev/null and switch when file appears
    (
      while [[ ! -f "${log_file}" ]]; do sleep 1; done
      tail -n "${TAIL_LINES}" -f "${log_file}" 2>/dev/null | colorize_line "${color}" "${label}"
    ) &
  else
    tail -n "${TAIL_LINES}" -f "${log_file}" 2>/dev/null \
      | colorize_line "${color}" "${label}" &
  fi

  TAIL_PIDS+=($!)
done

# ── Wait (blocking until Ctrl-C) ─────────────────────────────────────────────
wait
