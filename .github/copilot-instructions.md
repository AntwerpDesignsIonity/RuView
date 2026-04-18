# Copilot Workspace Instructions — AEDI-S

> **Ionity Global (Pty) Ltd** | ai@ionity.today | [www.ionity.today](https://www.ionity.today)
> WiFi-based human sensing platform: CSI → presence, vitals, pose estimation, activity recognition.

## Quick Start — Production Stack

```bash
# Full stack: deps → build → serve (auto-detects native vs Docker)
./.ionity/ionity.sh run --yes

# Status / stop
./.ionity/ionity.sh status
./.ionity/ionity.sh stop
```

**Services after startup:**

| Service | URL | Description |
|---------|-----|-------------|
| Web UI | `http://<HUB_IP>:3000/ui/index.html` | Main dashboard |
| Health | `http://<HUB_IP>:3000/health` | JSON health endpoint |
| WebSocket | `ws://<HUB_IP>:3001/ws/sensing` | Real-time CSI stream |
| UDP CSI | `udp://<HUB_IP>:5005` | ESP32 CSI data ingress |
| Python API | `http://localhost:8000` | FastAPI (optional) |

## Architecture Overview

```
ESP32-S3 nodes (CSI over UDP:5005)
        │
        ▼
┌─────────────────────────────────────┐
│  Rust Sensing Server (Axum)         │
│  HTTP :3000  │  WS :3001  │ UDP :5005│
│  Signal proc │  Pose est  │ Vitals  │
└─────────────────────────────────────┘
        │
        ▼
┌─────────────────┐   ┌──────────────┐
│  Static UI      │   │  Python API  │
│  (served :3000) │   │  (FastAPI)   │
└─────────────────┘   └──────────────┘
```

### Codebase Layout

| Directory | Purpose |
|-----------|---------|
| `rust-port/wifi-densepose-rs/` | Rust workspace (16 crates) — sensing server, signal, NN, MAT |
| `v1/` | Python v1 pipeline (FastAPI, training, data processing) |
| `ui/` | Static HTML/JS dashboards (served by Rust server) |
| `ionity/` | ESP32 PlatformIO firmware (Arduino framework) |
| `firmware/esp32-csi-node/` | ESP32 ESP-IDF firmware (C) + provision scripts |
| `.ionity/` | Unified stack launcher (ionity.sh + lib modules) |
| `scripts/` | Operational scripts (bridges, benchmarks, visualizers) |
| `docs/adr/` | Architecture Decision Records (43 ADRs) |
| `docker/` | Docker Compose + Dockerfiles |

## Build & Test

### Rust (primary runtime)
```bash
cd rust-port/wifi-densepose-rs
cargo build -p wifi-densepose-sensing-server --release --no-default-features  # server only
cargo test --workspace --no-default-features   # 1,463+ tests; see docs/issues/known-test-failures.md
cargo bench --package wifi-densepose-signal    # benchmarks
```

### Python
```bash
source .venv/bin/activate
python v1/data/proof/verify.py          # determinism proof — must print VERDICT: PASS
cd v1 && python -m pytest tests/ -x -q  # test suite
```

### ESP32 Firmware (PlatformIO)
```bash
cd ionity
platformio run -e esp32s3_n16r8          # build
platformio run -e esp32s3_n16r8 -t upload # flash
platformio device monitor --baud 460800   # serial monitor
```

### Verification (Trust Kill Switch)
```bash
make verify          # one-command proof replay
make verify-verbose  # detailed feature statistics
make verify-audit    # full pipeline + codebase audit
```

## Environment Configuration

Copy `example.env` to `.env.local` and override as needed:

| Variable | Default | Description |
|----------|---------|-------------|
| `IONITY_MODE` | `auto` | `native` / `docker` / `auto` |
| `IONITY_SOURCE` | `auto` | `esp32` / `wifi` / `simulate` / `auto` |
| `HUB_IP` | auto-detected | Aggregator LAN IP |
| `HUB_PORT` | `5005` | UDP CSI port |
| `HTTP_PORT` | `3000` | HTTP server port |
| `WS_PORT` | `3001` | WebSocket server port |

## Ionity CLI Commands

| Command | Description |
|---------|-------------|
| `.ionity/ionity.sh run [--yes]` | Full stack startup (deps → build → serve) |
| `.ionity/ionity.sh wifi` | Neighbour-WiFi passive sensing mode |
| `.ionity/ionity.sh deps` | Check and install all dependencies |
| `.ionity/ionity.sh build` | Build Rust binary + Python venv |
| `.ionity/ionity.sh docker` | Start via Docker |
| `.ionity/ionity.sh provision` | Interactive ESP32 flash & provision |
| `.ionity/ionity.sh scan` | Detect connected ESP32 boards |
| `.ionity/ionity.sh services` | Interactive service manager |
| `.ionity/ionity.sh status` | Show running services and health |
| `.ionity/ionity.sh stop` | Stop all services |
| `.ionity/ionity.sh autorepair` | Watchdog supervisor with auto-restart |
| `.ionity/ionity.sh scripts` | Browse script library |

## VS Code Tasks

Pre-configured tasks in [.vscode/tasks.json](.vscode/tasks.json):
- **IONITY: PlatformIO Build** — Build ESP32 firmware
- **IONITY: PlatformIO Upload** — Flash firmware to board
- **IONITY: PlatformIO Monitor** — Serial output at 460800 baud
- **IONITY: Provision ESP32 WiFi** — WiFi credential provisioning

## Process Management

- PIDs: `logs/sensing-server.pid`, `logs/aedi.pid`, `logs/csi-bridge.pid`
- Logs: `logs/sensing-server.log`
- Health: `curl http://localhost:3000/health`
- Stop: `.ionity/ionity.sh stop` or `pkill -f sensing-server`

## Conventions

- **Rust**: `cargo fmt` + `cargo clippy`, files < 500 lines, typed public APIs
- **Python**: Black formatting (line-length 88), flake8, mypy
- **Git**: Feature branches (`feature/*`, `feat/*`), PRs to `main`
- **Testing**: TDD London School (mock-first), run tests after every code change
- **Security**: No hardcoded secrets, validate at system boundaries, OWASP Top 10

## CI/CD Pipelines

| Workflow | Trigger | Purpose |
|----------|---------|---------|
| `ci.yml` | push/PR | Python tests, Black, flake8, mypy, Bandit security scan |
| `firmware-ci.yml` | firmware changes | ESP32 build, size gate (<1100 KB) |
| `verify-pipeline.yml` | core changes | Determinism proof replay |
| `firmware-qemu.yml` | firmware changes | QEMU chaos testing |
| `cd.yml` | push main/tags | Staging → production deploy |
| `security-scan.yml` | scheduled | Security audit |

## Supported Hardware

| Device | Chip | Role | Cost |
|--------|------|------|------|
| ESP32-S3 (8MB) | Xtensa dual-core | WiFi CSI sensing | ~$9 |
| ESP32-S3 SuperMini (4MB) | Xtensa dual-core | WiFi CSI (compact) | ~$6 |
| ESP32-C6 + MR60BHA2 | RISC-V + 60 GHz | mmWave HR/BR/presence | ~$15 |

**Not supported:** ESP32 (original), ESP32-C3 — single-core, can't run CSI DSP.

## Pre-Merge Checklist

1. `cargo test --workspace --no-default-features` — no NEW failures (see docs/issues/known-test-failures.md for tracked flakes)
2. `python v1/data/proof/verify.py` — VERDICT: PASS
3. Update CHANGELOG.md under `[Unreleased]`
4. Update README/CLAUDE.md if scope changed

## Detailed Documentation

- [CLAUDE.md](../CLAUDE.md) — Full crate table, ADR index, build details, swarm orchestration
- [docs/user-guide.md](../docs/user-guide.md) — End-user setup and usage guide
- [docs/adr/](../docs/adr/) — 43 Architecture Decision Records
- [docs/build-guide.md](../docs/build-guide.md) — Cross-compile and build details
- [docs/led-indication.md](../docs/led-indication.md) — ESP32 LED status codes
