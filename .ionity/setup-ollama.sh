#!/usr/bin/env bash
# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  setup-ollama.sh — Install Ollama + Gemma 3 4B for AEDI assistant           ║
# ║  Ionity (Pty) Ltd  |  ai@ionity.today  |  www.ionity.today                  ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
set -euo pipefail

BOLD='\033[1m'; DIM='\033[2m'; RESET='\033[0m'
RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[0;33m'; CYAN='\033[0;36m'
ok()   { echo -e "  ${GREEN}✔${RESET}  $*"; }
warn() { echo -e "  ${YELLOW}⚠${RESET}  $*"; }
err()  { echo -e "  ${RED}✖${RESET}  $*" >&2; }
step() { echo -e "  ${CYAN}→${RESET}  $*"; }

echo ""
echo -e "${CYAN}${BOLD}  ╔═══════════════════════════════════════════════════╗${RESET}"
echo -e "${CYAN}${BOLD}  ║   ⬡ AEDI — Ollama + Gemma 3 Setup                ║${RESET}"
echo -e "${CYAN}${BOLD}  ╚═══════════════════════════════════════════════════╝${RESET}"
echo ""

ARCHIVE="/tmp/ollama-linux-arm64.tar.zst"
INSTALL_DIR="/usr/local"

# ── Step 1: Check for downloaded archive ──────────────────────────────────────
if [[ -f "${ARCHIVE}" ]]; then
    SIZE=$(stat -c%s "${ARCHIVE}" 2>/dev/null || echo 0)
    SIZE_MB=$(( SIZE / 1048576 ))
    step "Found ${ARCHIVE} (${SIZE_MB} MB)"
    if [[ "${SIZE_MB}" -lt 1000 ]]; then
        warn "Archive seems incomplete (expected ~1210 MB, got ${SIZE_MB} MB)"
        echo "  Download URL: https://github.com/ollama/ollama/releases/download/v0.20.2/ollama-linux-arm64.tar.zst"
        echo "  Resume: curl -C - -L -o ${ARCHIVE} <url>"
        exit 1
    fi
else
    err "No archive found at ${ARCHIVE}"
    echo ""
    echo "  Download first (1.21 GB):"
    echo "    curl -L -o ${ARCHIVE} \\"
    echo "      https://github.com/ollama/ollama/releases/download/v0.20.2/ollama-linux-arm64.tar.zst"
    echo ""
    exit 1
fi

# ── Step 2: Extract ──────────────────────────────────────────────────────────
step "Extracting to ${INSTALL_DIR}/..."
sudo tar --zstd -xf "${ARCHIVE}" -C "${INSTALL_DIR}/" 2>&1
ok "Extracted"

# ── Step 3: Verify binary ────────────────────────────────────────────────────
if [[ -x "${INSTALL_DIR}/bin/ollama" ]]; then
    VERSION=$("${INSTALL_DIR}/bin/ollama" --version 2>/dev/null || echo "unknown")
    ok "ollama binary installed: ${VERSION}"
else
    err "ollama binary not found at ${INSTALL_DIR}/bin/ollama"
    exit 1
fi

# ── Step 4: Create systemd service ───────────────────────────────────────────
step "Creating systemd service..."
sudo tee /etc/systemd/system/ollama.service > /dev/null <<'EOF'
[Unit]
Description=Ollama LLM Service
After=network-online.target

[Service]
ExecStart=/usr/local/bin/ollama serve
User=antwerp
Group=antwerp
Restart=always
RestartSec=3
Environment="HOME=/home/antwerp"
Environment="OLLAMA_HOST=0.0.0.0"

[Install]
WantedBy=default.target
EOF

sudo systemctl daemon-reload
sudo systemctl enable ollama
sudo systemctl start ollama
sleep 3

if curl -sf http://localhost:11434/api/tags &>/dev/null; then
    ok "Ollama service running"
else
    warn "Ollama service started but API not responding yet — wait a moment"
fi

# ── Step 5: Pull Gemma 3 4B ──────────────────────────────────────────────────
step "Pulling Gemma 3 4B model (~2.5 GB)..."
echo "  This will take a few minutes on this connection."
"${INSTALL_DIR}/bin/ollama" pull gemma3:4b 2>&1
ok "Gemma 3 4B ready"

# ── Step 6: Quick test ───────────────────────────────────────────────────────
step "Testing inference..."
RESP=$("${INSTALL_DIR}/bin/ollama" run gemma3:4b "Say hello in exactly 5 words" 2>/dev/null | head -1)
ok "Model response: ${RESP}"

# ── Done ──────────────────────────────────────────────────────────────────────
echo ""
echo -e "${GREEN}${BOLD}  ✔  Ollama + Gemma 3 4B installed and ready!${RESET}"
echo ""
echo "  Now restart AEDI:"
echo "    .ionity/ionity.sh aedi restart"
echo ""
echo "  Or open the AEDI chat:"
echo "    http://$(hostname -I | awk '{print $1}'):3000/ui/aedi.html"
echo ""
