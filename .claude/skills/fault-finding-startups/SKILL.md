---
name: "Fault Finding Startups"
description: "Systematic diagnosis and resolution of startup failures across the Ionity RuView stack: Rust sensing-server, ESP32 firmware boot, Python venv/packages, Docker containers, serial ports, and network services. Use when any component fails to start, hangs during boot, crashes on init, or produces startup errors."
---

# Fault Finding Startups

## What This Skill Does

Provides a structured triage workflow to diagnose and fix startup failures across the full Ionity RuView stack. Covers the 4-phase startup pipeline (deps → build → stack → ready) and every component that can fail during initialisation.

## Startup Architecture

The Ionity launcher (`ionity.sh run`) executes 4 sequential phases:

```
Phase 1: Dependency Check  →  Phase 2: Build  →  Phase 3: Stack Start  →  Phase 4: Ready
  bash, curl, git, make        Python pip install     ESP32 board scan       Health check
  python3, pip, venv           Rust cargo build       Native or Docker       GUI open
  Rust toolchain               Firmware (optional)    Port binding           WebSocket alive
  Node.js (optional)                                  PID file write
```

A failure in any phase blocks all subsequent phases.

---

## Quick Triage — Start Here

Run these 5 checks in order. Stop at the first failure — that's your fault.

### 1. Is the previous instance still running?

```bash
pgrep -fa sensing-server
lsof -i :3000 -i :3001 -i :5005
cat logs/sensing-server.pid 2>/dev/null && ps -p "$(cat logs/sensing-server.pid)" 2>/dev/null
```

**Fix:** Kill stale processes before restarting.

```bash
./.ionity/ionity.sh stop
# or manually:
pkill -f sensing-server; sleep 1
```

### 2. Are ports free?

```bash
ss -tlnp | grep -E ':(3000|3001|5005)\s'
```

**Fix:** Kill the process holding the port, or change ports in `.env.local`:

```bash
echo 'HTTP_PORT=3002' >> .env.local
```

### 3. Does the binary/venv exist?

```bash
ls -la rust-port/wifi-densepose-rs/target/release/sensing-server
ls -la .venv/bin/activate
```

**Fix:** Rebuild.

```bash
./.ionity/ionity.sh build
```

### 4. Can the binary start at all?

```bash
RUST_LOG=debug rust-port/wifi-densepose-rs/target/release/sensing-server \
  --http-port 3000 --ws-port 3001 --udp-port 5005 \
  --bind-addr 0.0.0.0 --source simulate --ui-path ui/
```

Look at the first error line. Common causes below.

### 5. Does health respond?

```bash
curl -sf http://localhost:3000/health | python3 -m json.tool
```

**If no response:** Server crashed after binding — check `logs/sensing-server.log`.

---

## Phase-by-Phase Fault Guide

### Phase 1: Dependency Failures

| Symptom | Cause | Fix |
|---------|-------|-----|
| `bash: version 4` | Bash too old | `sudo apt install bash` or use `brew install bash` on macOS |
| `python3: command not found` | No Python | `sudo apt install python3 python3-venv python3-pip` |
| `python3 -c assert sys.version_info >= (3,10)` fails | Python < 3.10 | Install Python 3.10+ via deadsnakes PPA or pyenv |
| `pip: command not found` | pip missing | `python3 -m ensurepip --upgrade` |
| `No module named venv` | venv not installed | `sudo apt install python3-venv` |
| `cargo: command not found` | Rust not installed | `curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs \| sh` |
| venv creation fails | Disk full or permissions | Check `df -h` and `ls -la .venv/` |

### Phase 2: Build Failures

#### Rust Build

| Symptom | Cause | Fix |
|---------|-------|-----|
| `error[E0463]: can't find crate` | Missing dependency or stale cache | `cargo clean -p wifi-densepose-sensing-server && cargo build ...` |
| `error: linker cc not found` | No C compiler | `sudo apt install build-essential` |
| `error: could not compile` + out of memory | OOM during link | Add swap: `sudo fallocate -l 4G /swapfile && sudo mkswap /swapfile && sudo swapon /swapfile` |
| Build succeeds but binary missing | Wrong target dir | Check `CARGO_TARGET_DIR` env var; default is `rust-port/wifi-densepose-rs/target/release/` |
| `Blocking waiting for file lock` | Another cargo process | `rm -f rust-port/wifi-densepose-rs/target/.cargo-lock` |

**Full rebuild from scratch:**

```bash
cd rust-port/wifi-densepose-rs
cargo clean
cargo build -p wifi-densepose-sensing-server --release --no-default-features
```

**Check build log:**

```bash
cat logs/build.log | grep -E '^error|^warning.*unused|FAILED'
```

#### Python Build

| Symptom | Cause | Fix |
|---------|-------|-----|
| `pip install` hangs | Network or PyPI issue | `pip install --timeout 30 -r requirements.txt` |
| `ModuleNotFoundError` at runtime | Package not in venv | Activate venv first: `source .venv/bin/activate && pip install -e .` |
| numpy/scipy build fails | Missing BLAS/LAPACK | `sudo apt install libopenblas-dev liblapack-dev` |
| Version conflict | Lock file stale | `pip install -r v1/requirements-lock.txt --force-reinstall` |

### Phase 3: Stack Start Failures

#### Native Sensing Server

| Symptom | Cause | Fix |
|---------|-------|-----|
| `Address already in use` | Port conflict | Kill previous instance or change port |
| `Permission denied` binding port | Port < 1024 without root | Use ports > 1024 (default 3000/3001/5005 are fine) |
| Server starts then immediately exits | Panic on init | Run with `RUST_LOG=trace` and check stderr |
| `UI path not found` | Wrong `--ui-path` | Verify `ui/` directory exists at repo root |
| No health response after 15s | Slow start or silent crash | Check `logs/sensing-server.log` tail; verify PID is alive |
| `source: esp32` but no data | ESP32 not sending CSI | Check serial connection and WiFi provisioning |

**Debug startup directly:**

```bash
RUST_LOG=debug,tower_http=trace \
  rust-port/wifi-densepose-rs/target/release/sensing-server \
  --http-port 3000 --ws-port 3001 --udp-port 5005 \
  --bind-addr 127.0.0.1 --source simulate --ui-path ui/ 2>&1 | head -50
```

#### Docker

| Symptom | Cause | Fix |
|---------|-------|-----|
| `docker: command not found` | Docker not installed | Install Docker Engine |
| `permission denied` | User not in docker group | `sudo usermod -aG docker $USER` then re-login |
| Container exits immediately | Missing env vars or crash | `docker logs ruview` |
| Port mapping fails | Host port taken | Change ports or stop conflicting service |
| `image not found` | Not pulled | `docker pull ionity/ruview:latest` |

#### ESP32 Board Scan

| Symptom | Cause | Fix |
|---------|-------|-----|
| `No serial ports detected` | Board not plugged in or no driver | Check USB cable (must be data cable, not charge-only); install CP2102/CH340 driver |
| Port detected but `no ESP32 detected` | Wrong baud rate or board in wrong mode | Hold BOOT button while plugging in; try `esptool.py --port /dev/ttyUSB0 chip_id` |
| `esp32` or `esp32c3` — not supported | Single-core chip | Use ESP32-S3 or ESP32-C6 only |
| `No firmware binaries` | Release bins missing | Build firmware or download from GitHub releases |
| Provision fails mid-flash | Loose USB or timeout | Use short USB cable; retry with `--baud 115200` |
| `/dev/ttyACM0: Permission denied` | No serial group access | `sudo usermod -aG dialout $USER` then re-login |

### Phase 4: Ready-State Failures

| Symptom | Cause | Fix |
|---------|-------|-----|
| Health returns but `source: simulate` when ESP32 expected | Auto-detect fell through | Set `IONITY_SOURCE=esp32` in `.env.local` |
| GUI won't open | No display manager or headless | Open URL manually in browser |
| WebSocket connects then drops | CORS or firewall | Check browser console; ensure WS port (3001) is open |
| `tick: 0` stuck | No CSI data flowing | Verify ESP32 is provisioned and on same network |

---

## Log Analysis

### Where are the logs?

| Component | Log Location |
|-----------|-------------|
| Sensing server | `logs/sensing-server.log` |
| Build output | `logs/build.log` |
| Server PID | `logs/sensing-server.pid` |
| AEDI PID | `logs/aedi.pid` |

### Quick log commands

```bash
# Last 50 lines of server log
tail -50 logs/sensing-server.log

# Watch live
tail -f logs/sensing-server.log

# Find errors
grep -i 'error\|panic\|fatal\|WARN' logs/sensing-server.log | tail -20

# Find startup sequence
grep -E 'listening|bound|started|ready|failed' logs/sensing-server.log | tail -10
```

---

## Environment Variable Overrides

All configurable via `.env.local` at repo root:

| Variable | Default | Purpose |
|----------|---------|---------|
| `IONITY_MODE` | `auto` | `native`, `docker`, or `auto` |
| `IONITY_SOURCE` | `auto` | `esp32`, `wifi`, `simulate`, or `auto` |
| `HUB_IP` | auto-detected | Bind/advertise IP |
| `HUB_PORT` | `5005` | UDP CSI port |
| `HTTP_PORT` | `3000` | HTTP API port |
| `WS_PORT` | `3001` | WebSocket port |
| `IONITY_TRACE` | `info,tower_http=debug` | Rust log level |
| `IONITY_SKIP_PROVISION` | `0` | Skip ESP32 board scan |

---

## Nuclear Reset

When all else fails — clean slate:

```bash
# 1. Stop everything
./.ionity/ionity.sh stop
pkill -f sensing-server 2>/dev/null; pkill -f aedi 2>/dev/null

# 2. Clean build artifacts
cd rust-port/wifi-densepose-rs && cargo clean && cd ../..
rm -rf .venv

# 3. Remove stale state
rm -f logs/*.pid logs/*.log

# 4. Full rebuild from scratch
./.ionity/ionity.sh run
```

---

## Troubleshooting Decision Tree

```
Start failed?
├── Phase 1 error? → Check dependency table above
├── Phase 2 error? → Check build log: logs/build.log
├── Phase 3 error?
│   ├── Port in use?     → Kill or change port
│   ├── Binary missing?  → Rebuild
│   ├── Crash on start?  → Run manually with RUST_LOG=debug
│   └── ESP32 scan fail? → Check USB/serial/driver
└── Phase 4 error?
    ├── Health 200 but wrong source? → Set IONITY_SOURCE
    ├── No health response? → Check PID alive, check log
    └── GUI won't load? → Open URL manually, check CORS
```
