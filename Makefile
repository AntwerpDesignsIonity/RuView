# WiFi-DensePose Makefile
# ============================================================

.PHONY: verify verify-verbose verify-audit install install-verify install-python \
        install-rust install-browser install-docker install-field install-full \
        check build-rust build-wasm test-rust bench run-api run-viz clean clean-pio \
        desktop desktop-clean help

# ─── Installation ────────────────────────────────────────────
# Guided interactive installer
install:
	@./install.sh

# Profile-specific installs (non-interactive)
install-verify:
	@./install.sh --profile verify --yes

install-python:
	@./install.sh --profile python --yes

install-rust:
	@./install.sh --profile rust --yes

install-browser:
	@./install.sh --profile browser --yes

install-docker:
	@./install.sh --profile docker --yes

install-field:
	@./install.sh --profile field --yes

install-full:
	@./install.sh --profile full --yes

# Hardware and environment check only (no install)
check:
	@./install.sh --check-only

# ─── Verification ────────────────────────────────────────────
# Trust Kill Switch -- one-command proof replay
verify:
	@./verify

# Verbose mode -- show detailed feature statistics and Doppler spectrum
verify-verbose:
	@./verify --verbose

# Full audit -- verify pipeline + scan codebase for mock/random patterns
verify-audit:
	@./verify --verbose --audit

# ─── Rust Builds ─────────────────────────────────────────────
build-rust:
	cd rust-port/wifi-densepose-rs && cargo build --release

build-wasm:
	cd rust-port/wifi-densepose-rs && wasm-pack build crates/wifi-densepose-wasm --target web --release

build-wasm-mat:
	cd rust-port/wifi-densepose-rs && wasm-pack build crates/wifi-densepose-wasm --target web --release -- --features mat

test-rust:
	cd rust-port/wifi-densepose-rs && cargo test --workspace

bench:
	cd rust-port/wifi-densepose-rs && cargo bench --package wifi-densepose-signal

# ─── Run ─────────────────────────────────────────────────────
run-api:
	uvicorn v1.src.api.main:app --host 0.0.0.0 --port 8000

run-api-dev:
	uvicorn v1.src.api.main:app --host 0.0.0.0 --port 8000 --reload

run-viz:
	python3 -m http.server 3000 --directory ui

run-docker:
	docker compose up

# ─── Clean ───────────────────────────────────────────────────
clean:
	rm -f .install.log
	cd rust-port/wifi-densepose-rs && cargo clean 2>/dev/null || true

# Reclaim disk by removing PlatformIO caches, duplicate riscv32 toolchains,
# and per-env build trees. Safe to run any time; next build re-fetches.
clean-pio:
	@echo "Before:" && df -h / | awk 'NR==2 {print "  /  " $$3 " used / " $$2 " total (" $$5 " full)"}'
	@rm -rf $$HOME/.platformio/.cache/* 2>/dev/null || true
	@find $$HOME/.platformio/packages -maxdepth 1 -type d -name 'toolchain-riscv32-esp@*' -exec rm -rf {} + 2>/dev/null || true
	@rm -rf ionity/.pio/build/*/.sconsign* ionity/.pio/build/*/firmware.elf 2>/dev/null || true
	@rm -rf firmware/esp32-csi-node/build firmware/esp32-csi-node/managed_components 2>/dev/null || true
	@echo "After:"  && df -h / | awk 'NR==2 {print "  /  " $$3 " used / " $$2 " total (" $$5 " full)"}'
	@echo "Done. To free more, run 'make clean' (cargo clean) too."

# ─── Help ────────────────────────────────────────────────────
help:
	@echo "WiFi-DensePose Build Targets"
	@echo "============================================================"
	@echo ""
	@echo "  Installation:"
	@echo "    make install          Interactive guided installer"
	@echo "    make install-verify   Verification only (~5 MB)"
	@echo "    make install-python   Full Python pipeline (~500 MB)"
	@echo "    make install-rust     Rust pipeline with ~810x speedup"
	@echo "    make install-browser  WASM for browser (~10 MB)"
	@echo "    make install-docker   Docker-based deployment"
	@echo "    make install-field    WiFi-Mat disaster kit (~62 MB)"
	@echo "    make install-full     Everything available"
	@echo "    make check            Hardware/environment check only"
	@echo ""
	@echo "  Verification:"
	@echo "    make verify           Run the trust kill switch"
	@echo "    make verify-verbose   Verbose with feature details"
	@echo "    make verify-audit     Full verification + codebase audit"
	@echo ""
	@echo "  Build:"
	@echo "    make build-rust       Build Rust workspace (release)"
	@echo "    make build-wasm       Build WASM package (browser)"
	@echo "    make build-wasm-mat   Build WASM with WiFi-Mat (field)"
	@echo "    make test-rust        Run all Rust tests"
	@echo "    make bench            Run signal processing benchmarks"
	@echo ""
	@echo "  Run:"
	@echo "    make run-api          Start Python API server"
	@echo "    make run-api-dev      Start API with hot-reload"
	@echo "    make run-viz          Serve 3D visualization (port 3000)"
	@echo "    make run-docker       Start Docker dev stack"
	@echo ""
	@echo "  Utility:"
	@echo "    make clean            Remove build artifacts"
	@echo "    make help             Show this help"
	@echo ""

# ─── AEDI Desktop (.NET MAUI) ────────────────────────────────
# Builds the desktop client for the host platform. Requires .NET 10 SDK
# with the maui workload installed (`dotnet workload install maui`).
desktop:
	@echo "Building AEDI Desktop (host platform)..."
	@command -v dotnet >/dev/null 2>&1 || { echo "dotnet not found. Install .NET 10 SDK + maui workload."; exit 1; }
	cd desktop && dotnet build AEDI.Desktop.sln -c Release

desktop-clean:
	@echo "Cleaning AEDI Desktop build artifacts..."
	@find desktop -type d \( -name bin -o -name obj \) -exec rm -rf {} + 2>/dev/null || true
