#!/usr/bin/env bash
# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  AEDI-S — Raspberry Pi Zero 2 W "Cognitum Seed" Setup                      ║
# ║  Installs all dependencies for the edge AI pipeline + sensing server.       ║
# ║                                                                              ║
# ║  Usage:                                                                      ║
# ║    bash scripts/setup-rpi-seed.sh              # Full setup                  ║
# ║    bash scripts/setup-rpi-seed.sh --skip-rust   # Skip Rust toolchain        ║
# ║    bash scripts/setup-rpi-seed.sh --skip-node   # Skip Node.js               ║
# ║    bash scripts/setup-rpi-seed.sh --check        # Check only (no install)    ║
# ║                                                                              ║
# ║  Ionity Global (Pty) Ltd  |  ai@ionity.today  |  www.ionity.today           ║
# ╚══════════════════════════════════════════════════════════════════════════════╝

set -euo pipefail

# ── Options ───────────────────────────────────────────────────────────────────
SKIP_RUST=false
SKIP_NODE=false
CHECK_ONLY=false

while [[ $# -gt 0 ]]; do
  case "$1" in
    --skip-rust)  SKIP_RUST=true;  shift ;;
    --skip-node)  SKIP_NODE=true;  shift ;;
    --check)      CHECK_ONLY=true; shift ;;
    -h|--help)
      head -14 "$0" | tail -9
      exit 0
      ;;
    *) echo "Unknown option: $1"; exit 1 ;;
  esac
done

# ── Paths ─────────────────────────────────────────────────────────────────────
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_DIR="$(cd "${SCRIPT_DIR}/.." && pwd)"
RUST_DIR="${REPO_DIR}/rust-port/wifi-densepose-rs"
MODELS_DIR="${REPO_DIR}/models/pretrained"
VENV_DIR="${REPO_DIR}/.venv"
LOG_FILE="${REPO_DIR}/logs/rpi-setup.log"

mkdir -p "${REPO_DIR}/logs"

# ── Colours ───────────────────────────────────────────────────────────────────
if [ -t 1 ]; then
  RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'
  CYAN='\033[0;36m'; BOLD='\033[1m'; DIM='\033[2m'; RESET='\033[0m'
else
  RED=''; GREEN=''; YELLOW=''; CYAN=''; BOLD=''; DIM=''; RESET=''
fi

ok()   { echo -e "  ${GREEN}✔${RESET} $1"; }
warn() { echo -e "  ${YELLOW}⚠${RESET}  $1"; }
fail() { echo -e "  ${RED}✖${RESET} $1"; }
step() { echo -e "\n${CYAN}${BOLD}── $1 ──${RESET}"; }
info() { echo -e "  ${DIM}$1${RESET}"; }

# ── Banner ────────────────────────────────────────────────────────────────────
echo ""
echo -e "${CYAN}${BOLD}╔══════════════════════════════════════════════════════════╗${RESET}"
echo -e "${CYAN}${BOLD}║${RESET}  ${BOLD}AEDI-S — Cognitum Seed (RPi Zero 2 W) Setup${RESET}"
echo -e "${CYAN}${BOLD}║${RESET}  ${DIM}Edge AI pipeline: Python + ONNX + Node.js + Rust${RESET}"
echo -e "${CYAN}${BOLD}╚══════════════════════════════════════════════════════════╝${RESET}"
echo ""

exec > >(tee -a "${LOG_FILE}") 2>&1
echo "--- setup-rpi-seed.sh started $(date -u +%Y-%m-%dT%H:%M:%SZ) ---" >> "${LOG_FILE}"

# ══════════════════════════════════════════════════════════════════════════════
#  STEP 1: System prerequisites
# ══════════════════════════════════════════════════════════════════════════════
step "1/7  System prerequisites"

# Detect architecture
ARCH="$(uname -m)"
info "Architecture: ${ARCH}"

if [[ "${ARCH}" != "aarch64" && "${ARCH}" != "armv7l" && "${ARCH}" != "armv6l" ]]; then
  warn "Not running on ARM — this script is designed for Raspberry Pi."
  warn "Continuing anyway (may work on other Linux systems)."
fi

# RAM check
TOTAL_RAM_MB=$(awk '/MemTotal/ {print int($2/1024)}' /proc/meminfo 2>/dev/null || echo 0)
if [[ "${TOTAL_RAM_MB}" -lt 400 ]]; then
  warn "RAM: ${TOTAL_RAM_MB} MB — Pi Zero 2 W has 512 MB. Swap is recommended."
else
  ok "RAM: ${TOTAL_RAM_MB} MB"
fi

# Disk check
DISK_FREE_MB=$(df -m "${REPO_DIR}" 2>/dev/null | awk 'NR==2 {print $4}' || echo 0)
if [[ "${DISK_FREE_MB}" -lt 2000 ]]; then
  warn "Disk: ${DISK_FREE_MB} MB free (2 GB+ recommended)"
else
  ok "Disk: ${DISK_FREE_MB} MB free"
fi

if ${CHECK_ONLY}; then
  echo ""
  info "Check-only mode — skipping installs, just verifying what's present."
fi

# Ensure apt packages are up-to-date
if ! ${CHECK_ONLY}; then
  step "1b/7  Updating apt package index"
  sudo apt-get update -qq
  ok "apt index updated"

  step "1c/7  Installing system dependencies"
  sudo apt-get install -y -qq \
    build-essential \
    pkg-config \
    libssl-dev \
    libffi-dev \
    python3-dev \
    python3-venv \
    python3-pip \
    git \
    curl \
    wget \
    cmake \
    2>&1 | tail -1
  ok "System packages installed"
fi

# ══════════════════════════════════════════════════════════════════════════════
#  STEP 2: Python 3 + venv + ONNX Runtime + pyserial
# ══════════════════════════════════════════════════════════════════════════════
step "2/7  Python 3 & ONNX Runtime"

if command -v python3 &>/dev/null; then
  PY_VER="$(python3 --version 2>&1)"
  ok "Python: ${PY_VER}"
else
  fail "Python 3 not found. Install: sudo apt-get install python3"
  exit 1
fi

# Create venv if missing
if [[ ! -f "${VENV_DIR}/bin/activate" ]]; then
  if ${CHECK_ONLY}; then
    warn "Python venv not found at ${VENV_DIR}"
  else
    info "Creating Python virtual environment…"
    python3 -m venv "${VENV_DIR}"
    ok "venv created at ${VENV_DIR}"
  fi
else
  ok "Python venv exists"
fi

# Activate venv
if [[ -f "${VENV_DIR}/bin/activate" ]]; then
  # shellcheck disable=SC1090
  source "${VENV_DIR}/bin/activate"
  ok "venv activated ($(python3 --version))"
fi

# Install Python packages for edge inference
if ! ${CHECK_ONLY}; then
  info "Installing edge inference packages (onnxruntime, pyserial, numpy, scipy)…"

  # On armv7l/aarch64, onnxruntime may need special wheels
  pip install --upgrade pip setuptools wheel -q

  # Core edge packages
  pip install -q \
    onnxruntime \
    pyserial \
    numpy \
    scipy \
    requests \
    pyyaml \
    2>&1 | tail -3

  ok "Python edge packages installed"
else
  # Check what's available
  python3 -c "import onnxruntime; print(f'  onnxruntime {onnxruntime.__version__}')" 2>/dev/null && ok "onnxruntime" || warn "onnxruntime not installed"
  python3 -c "import serial; print(f'  pyserial {serial.__version__}')" 2>/dev/null && ok "pyserial" || warn "pyserial not installed"
  python3 -c "import numpy; print(f'  numpy {numpy.__version__}')" 2>/dev/null && ok "numpy" || warn "numpy not installed"
  python3 -c "import scipy; print(f'  scipy {scipy.__version__}')" 2>/dev/null && ok "scipy" || warn "scipy not installed"
fi

# ══════════════════════════════════════════════════════════════════════════════
#  STEP 3: Node.js (for JS spatial pipelines)
# ══════════════════════════════════════════════════════════════════════════════
step "3/7  Node.js"

if ${SKIP_NODE}; then
  info "Skipped (--skip-node)"
else
  if command -v node &>/dev/null; then
    NODE_VER="$(node --version 2>&1)"
    ok "Node.js: ${NODE_VER}"
  else
    if ${CHECK_ONLY}; then
      warn "Node.js not installed"
    else
      info "Installing Node.js 20 LTS via NodeSource…"
      curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash - 2>&1 | tail -2
      sudo apt-get install -y -qq nodejs 2>&1 | tail -1
      ok "Node.js $(node --version) installed"
    fi
  fi

  if command -v npm &>/dev/null; then
    ok "npm: $(npm --version)"
  fi
fi

# ══════════════════════════════════════════════════════════════════════════════
#  STEP 4: Rust toolchain (for sensing-server)
# ══════════════════════════════════════════════════════════════════════════════
step "4/7  Rust toolchain"

if ${SKIP_RUST}; then
  info "Skipped (--skip-rust)"
else
  if command -v rustc &>/dev/null; then
    RUST_VER="$(rustc --version 2>&1)"
    ok "Rust: ${RUST_VER}"
  else
    if ${CHECK_ONLY}; then
      warn "Rust not installed"
    else
      info "Installing Rust via rustup (this takes a few minutes on Pi)…"
      curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y 2>&1 | tail -3
      # shellcheck disable=SC1090
      source "${HOME}/.cargo/env" 2>/dev/null || true
      export PATH="${HOME}/.cargo/bin:${PATH}"
      ok "Rust $(rustc --version) installed"
    fi
  fi

  if command -v cargo &>/dev/null; then
    ok "Cargo: $(cargo --version)"
  fi
fi

# ══════════════════════════════════════════════════════════════════════════════
#  STEP 5: Download ONNX models & weights
# ══════════════════════════════════════════════════════════════════════════════
step "5/7  ONNX models & weights"

MODEL_FILES=(
  "model-q4.bin"
  "presence-head.json"
  "config.json"
)

# Check which files already exist
MISSING_FILES=()
for f in "${MODEL_FILES[@]}"; do
  if [[ -f "${MODELS_DIR}/${f}" ]]; then
    ok "${f} ($(stat -c%s "${MODELS_DIR}/${f}" 2>/dev/null || echo '?') bytes)"
  else
    MISSING_FILES+=("${f}")
    warn "${f} missing"
  fi
done

# Check for pretrained-encoder.onnx (primary ONNX model)
if [[ -f "${MODELS_DIR}/pretrained-encoder.onnx" ]]; then
  ok "pretrained-encoder.onnx present"
else
  MISSING_FILES+=("pretrained-encoder.onnx")
  warn "pretrained-encoder.onnx missing"
fi

if [[ "${#MISSING_FILES[@]}" -gt 0 ]]; then
  info ""
  info "Some model files are missing. To download from HuggingFace:"
  info ""
  info "  # Option 1: Using huggingface-cli (recommended)"
  info "  pip install huggingface-hub"
  info "  huggingface-cli download ionity-global/wifi-densepose-pretrained \\"
  info "    --local-dir ${MODELS_DIR} \\"
  info "    --include 'model-q4.bin' 'presence-head.json' 'config.json' 'pretrained-encoder.onnx'"
  info ""
  info "  # Option 2: Using wget"
  info "  HF_BASE=https://huggingface.co/ionity-global/wifi-densepose-pretrained/resolve/main"
  for f in "${MISSING_FILES[@]}"; do
    info "  wget -O ${MODELS_DIR}/${f} \${HF_BASE}/${f}"
  done
  info ""

  if ! ${CHECK_ONLY}; then
    # Attempt download if huggingface-hub is available or wget
    if python3 -c "import huggingface_hub" 2>/dev/null; then
      info "Attempting download via huggingface-hub…"
      python3 -c "
from huggingface_hub import hf_hub_download
import os
repo = 'ionity-global/wifi-densepose-pretrained'
dest = '${MODELS_DIR}'
for f in ${MISSING_FILES[@]@Q}.split():
    try:
        hf_hub_download(repo_id=repo, filename=f, local_dir=dest)
        print(f'  Downloaded: {f}')
    except Exception as e:
        print(f'  Failed: {f} — {e}')
" 2>&1 || true
    else
      info "huggingface-hub not installed — install with: pip install huggingface-hub"
      info "Or manually download the files listed above."
    fi
  fi
else
  ok "All model files present"
fi

# ══════════════════════════════════════════════════════════════════════════════
#  STEP 6: Serial port permissions (dialout group)
# ══════════════════════════════════════════════════════════════════════════════
step "6/7  Serial port permissions"

CURRENT_USER="${USER:-$(whoami)}"
if id -nG "${CURRENT_USER}" | grep -qw dialout; then
  ok "${CURRENT_USER} is in dialout group"
else
  if ${CHECK_ONLY}; then
    warn "${CURRENT_USER} is NOT in dialout group — ESP32 serial will not work"
  else
    info "Adding ${CURRENT_USER} to dialout group…"
    sudo usermod -a -G dialout "${CURRENT_USER}"
    ok "${CURRENT_USER} added to dialout — log out and back in to take effect"
  fi
fi

# Check for connected ESP32 devices
ESP_PORTS=()
for p in /dev/ttyUSB* /dev/ttyACM*; do
  [[ -e "${p}" ]] && ESP_PORTS+=("${p}")
done
if [[ "${#ESP_PORTS[@]}" -gt 0 ]]; then
  ok "Serial ports detected: ${ESP_PORTS[*]}"
else
  info "No ESP32 serial ports detected (plug in via USB to use)"
fi

# ══════════════════════════════════════════════════════════════════════════════
#  STEP 7: Swap file (recommended for Pi Zero 2 W 512 MB RAM)
# ══════════════════════════════════════════════════════════════════════════════
step "7/7  Swap & performance tuning"

SWAP_TOTAL=$(free -m | awk '/Swap/ {print $2}')
if [[ "${SWAP_TOTAL}" -ge 1024 ]]; then
  ok "Swap: ${SWAP_TOTAL} MB"
elif [[ "${SWAP_TOTAL}" -gt 0 ]]; then
  warn "Swap: ${SWAP_TOTAL} MB (1024+ MB recommended for Rust build + inference)"
  if ! ${CHECK_ONLY}; then
    info "To increase swap:"
    info "  sudo dphys-swapfile swapoff"
    info "  sudo sed -i 's/CONF_SWAPSIZE=.*/CONF_SWAPSIZE=2048/' /etc/dphys-swapfile"
    info "  sudo dphys-swapfile setup && sudo dphys-swapfile swapon"
  fi
else
  warn "No swap configured — Rust compilation will likely OOM on 512 MB RAM"
  if ! ${CHECK_ONLY}; then
    info "Creating 2 GB swap file…"
    if [[ -f /etc/dphys-swapfile ]]; then
      sudo sed -i 's/CONF_SWAPSIZE=.*/CONF_SWAPSIZE=2048/' /etc/dphys-swapfile
      sudo dphys-swapfile setup 2>&1 | tail -1
      sudo dphys-swapfile swapon
      ok "2 GB swap enabled"
    else
      sudo fallocate -l 2G /swapfile
      sudo chmod 600 /swapfile
      sudo mkswap /swapfile
      sudo swapon /swapfile
      echo '/swapfile none swap sw 0 0' | sudo tee -a /etc/fstab > /dev/null
      ok "2 GB swap created at /swapfile"
    fi
  fi
fi

# ══════════════════════════════════════════════════════════════════════════════
#  Summary
# ══════════════════════════════════════════════════════════════════════════════
echo ""
echo -e "${CYAN}${BOLD}╔══════════════════════════════════════════════════════════╗${RESET}"
echo -e "${CYAN}${BOLD}║${RESET}  ${BOLD}Setup Complete — Cognitum Seed Ready${RESET}"
echo -e "${CYAN}${BOLD}╚══════════════════════════════════════════════════════════╝${RESET}"
echo ""
echo -e "  ${BOLD}Quick-start commands:${RESET}"
echo ""
echo -e "  ${DIM}# 1. Activate Python venv${RESET}"
echo -e "  source ${VENV_DIR}/bin/activate"
echo ""
echo -e "  ${DIM}# 2. Run CSI bridge (ESP32 → Cognitum Seed)${RESET}"
echo -e "  python scripts/seed_csi_bridge.py \\"
echo -e "    --seed-url https://169.254.42.1:8443 \\"
echo -e "    --token \"\$SEED_TOKEN\" --udp-port 5006"
echo ""
echo -e "  ${DIM}# 3. Run JS spatial pipelines${RESET}"
echo -e "  node scripts/snn-csi-processor.js"
echo -e "  node scripts/mincut-person-counter.js"
echo ""
if ! ${SKIP_RUST}; then
  echo -e "  ${DIM}# 4. Build & run Rust sensing server${RESET}"
  echo -e "  cd rust-port/wifi-densepose-rs"
  echo -e "  cargo build -p wifi-densepose-sensing-server --release --no-default-features"
  echo -e "  ./target/release/sensing-server"
  echo ""
fi
echo -e "  ${DIM}# 5. Launch full Ionity stack${RESET}"
echo -e "  ./.ionity/ionity.sh run"
echo ""

if ! id -nG "${CURRENT_USER}" | grep -qw dialout; then
  echo -e "  ${YELLOW}${BOLD}⚠  Log out and back in for serial port access (dialout group)${RESET}"
  echo ""
fi

echo -e "  ${DIM}Log: ${LOG_FILE}${RESET}"
echo ""
