# WiFi-DensePose Build and Run Guide

Covers every way to build, run, and deploy the system -- from a zero-hardware verification to a full ESP32 mesh with 3D visualization.

---

## Table of Contents

1. [Quick Start (Verification Only -- No Hardware)](#1-quick-start-verification-only----no-hardware)
2. [Python Pipeline (v1/)](#2-python-pipeline-v1)
3. [Rust Pipeline (v2)](#3-rust-pipeline-v2)
4. [Three.js Visualization](#4-threejs-visualization)
5. [Docker Deployment](#5-docker-deployment)
6. [ESP32 Hardware Setup](#6-esp32-hardware-setup)
7. [Environment-Specific Builds](#7-environment-specific-builds)

---

## 1. Quick Start (Verification Only -- No Hardware)

The fastest way to confirm the signal processing pipeline is real and deterministic. Requires only Python 3.8+, numpy, and scipy. No WiFi hardware, no GPU, no Docker.

```bash
# From the repository root:
./verify
```

This runs three phases:

1. **Environment checks** -- confirms Python, numpy, scipy, and proof files are present.
2. **Proof pipeline replay** -- feeds a published reference signal through the full signal processing chain (noise filtering, Hamming windowing, amplitude normalization, FFT-based Doppler extraction, power spectral density via scipy.fft) and computes a SHA-256 hash of the output.
3. **Production code integrity scan** -- scans `v1/src/` for `np.random.rand` / `np.random.randn` calls in production code (test helpers are excluded).

Exit codes:
- `0` PASS -- pipeline hash matches the published expected hash
- `1` FAIL -- hash mismatch or error
- `2` SKIP -- no expected hash file to compare against

Additional flags:

```bash
./verify --verbose         # Detailed feature statistics and Doppler spectrum
./verify --verbose --audit # Full verification + codebase audit

# Or via make:
make verify
make verify-verbose
make verify-audit
```

If the expected hash file is missing, regenerate it:

```bash
python3 v1/data/proof/verify.py --generate-hash
```

### Minimal dependencies for verification only

```bash
pip install numpy==1.26.4 scipy==1.14.1
```

Or install the pinned set that guarantees hash reproducibility:

```bash
pip install -r v1/requirements-lock.txt
```

The lock file pins: `numpy==1.26.4`, `scipy==1.14.1`, `pydantic==2.10.4`, `pydantic-settings==2.7.1`.

---

## 2. Python Pipeline (v1/)

The Python pipeline lives under `v1/` and provides the full API server, signal processing, sensing modules, and WebSocket streaming.

### Prerequisites

- Python 3.8+
- pip

### Install (verification-only -- lightweight)

```bash
pip install -r v1/requirements-lock.txt
```

This installs only the four packages needed for deterministic pipeline verification.

### Install (full pipeline with API server)

```bash
pip install -r requirements.txt
```

This pulls in FastAPI, uvicorn, torch, OpenCV, SQLAlchemy, Redis client, and all other runtime dependencies.

### Verify the pipeline

```bash
python3 v1/data/proof/verify.py
```

Same as `./verify` but calls the Python script directly, skipping the bash wrapper's codebase scan phase.

### Run the API server

```bash
uvicorn v1.src.api.main:app --host 0.0.0.0 --port 8000
```

The server exposes:
- REST API docs: http://localhost:8000/docs
- Health check: http://localhost:8000/health
- Latest poses: http://localhost:8000/api/v1/pose/latest
- WebSocket pose stream: ws://localhost:8000/ws/pose/stream
- WebSocket analytics: ws://localhost:8000/ws/analytics/events

For development with auto-reload:

```bash
uvicorn v1.src.api.main:app --host 0.0.0.0 --port 8000 --reload
```

### Run with commodity WiFi (RSSI sensing -- no custom hardware)

The commodity sensing module (`v1/src/sensing/`) extracts presence and motion features from standard Linux WiFi metrics (RSSI, noise floor, link quality) without any hardware modification. See [ADR-013](adr/ADR-013-feature-level-sensing-commodity-gear.md) for full design details.

Requirements:
- Any Linux machine with a WiFi interface (laptop, Raspberry Pi, etc.)
- Connected to a WiFi access point (the AP is the signal source)
- No root required for basic RSSI reading via `/proc/net/wireless`

The module provides:
- `LinuxWifiCollector` -- reads real RSSI from `/proc/net/wireless` and `iw` commands
- `RssiFeatureExtractor` -- computes rolling statistics, FFT spectral features, CUSUM change-point detection
- `PresenceClassifier` -- rule-based presence/motion classification

What it can detect:
| Capability | Single Receiver | 3+ Receivers |
|-----------|----------------|-------------|
| Binary presence | Yes (90-95%) | Yes (90-95%) |
| Coarse motion (still/moving) | Yes (85-90%) | Yes (85-90%) |
| Room-level location | No | Marginal (70-80%) |

What it cannot detect: body pose, heartbeat, reliable respiration. See ADR-013 for the honest capability matrix.

### Python project structure

```
v1/
  src/
    api/
      main.py              # FastAPI application entry point
      routers/             # REST endpoint routers (pose, stream, health)
      middleware/           # Auth, rate limiting
      websocket/           # WebSocket connection manager, pose stream
    config/                # Settings, domain configs
    sensing/
      rssi_collector.py    # LinuxWifiCollector + SimulatedCollector
      feature_extractor.py # RssiFeatureExtractor (FFT, CUSUM, spectral)
      classifier.py        # PresenceClassifier (rule-based)
      backend.py           # SensingBackend protocol
  data/
    proof/
      sample_csi_data.json       # Deterministic reference signal
      expected_features.sha256   # Published expected hash
      verify.py                  # One-command verification script
  requirements-lock.txt          # Pinned deps for hash reproducibility
```

---

## 3. Rust Pipeline (v2)

A high-performance Rust port with ~810x speedup over the Python pipeline for the full signal processing chain.

### Prerequisites

- Rust 1.70+ (install via [rustup](https://rustup.rs/))
- cargo (included with Rust)
- No additional system libraries required when building with `--no-default-features`

> **Note:** The `sensing-server` crate must be built with `--no-default-features`. The default feature set enables `eigenvalue` → `ndarray-linalg` → `lax`, which fails to compile on Rust 1.80+ (Rust 1.94+ in particular). The signal processing used by the sensing server does not require BLAS.

### Build

```bash
cd rust-port/wifi-densepose-rs

# Build the sensing server (recommended — no BLAS dependency)
cargo build -p wifi-densepose-sensing-server --release --no-default-features

# Or build the full workspace (requires libopenblas-dev on Linux, openblas via brew on macOS)
cargo build --release
```

Release profile is configured with LTO, single codegen unit, and `-O3` for maximum performance.

### Run

The compiled binary is at `target/release/sensing-server`. Default ports are HTTP 8080 and WS 8765 — always override them explicitly:

```bash
# Native run (ESP32 live CSI source, shown on Pi at 172.23.9.61)
./target/release/sensing-server \
  --http-port 3000 \
  --ws-port 3001 \
  --udp-port 5005 \
  --bind-addr 0.0.0.0 \
  --source esp32 \
  --ui-path /path/to/RuView/ui

# Simulated mode (no hardware)
./target/release/sensing-server \
  --http-port 3000 \
  --ws-port 3001 \
  --source simulate
```

**CLI flags reference:**

| Flag | Default | Description |
|------|---------|-------------|
| `--http-port` | 8080 | HTTP / REST API port |
| `--ws-port` | 8765 | WebSocket port (UI connects here) |
| `--udp-port` | 5005 | UDP port receiving ESP32 CSI frames |
| `--bind-addr` | 127.0.0.1 | Bind address — use `0.0.0.0` for LAN access |
| `--source` | auto | `auto`, `esp32`, `simulate`, `wifi` |
| `--ui-path` | (none) | Serve UI static files from this directory |

**Health check:**

```bash
curl http://localhost:3000/health
# {"status":"ok","source":"esp32","clients":0,"tick":0}
# tick > 0 means frames are arriving from ESP32 nodes
```

**End-to-end self-test (no hardware):**

```python
# Send a synthetic ESP32 CSI frame to verify the server processes UDP
import socket, struct, time
sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
for seq in range(10):
    pkt  = struct.pack('<I', 0xC5110001)   # [0..3]  magic
    pkt += struct.pack('B', 1)             # [4]     node_id
    pkt += struct.pack('B', 1)             # [5]     n_antennas
    pkt += struct.pack('B', 8)             # [6]     n_subcarriers
    pkt += struct.pack('B', 0)             # [7]     reserved
    pkt += struct.pack('<H', 2412)         # [8..9]  freq_mhz
    pkt += struct.pack('<I', seq)          # [10..13] sequence
    pkt += struct.pack('b', -50)           # [14]    rssi (dBm)
    pkt += struct.pack('b', -95)           # [15]    noise_floor
    pkt += bytes(4)                        # [16..19] reserved
    for k in range(8):
        pkt += struct.pack('bb', 30+k, 10+k)  # I/Q pairs
    sock.sendto(pkt, ('127.0.0.1', 5005))
    time.sleep(0.05)
sock.close()
# Then: curl http://localhost:3000/health → tick should equal 10
```

### Test

```bash
cd rust-port/wifi-densepose-rs
cargo test --workspace --no-default-features
```

Runs 1,031+ tests across all workspace crates.

### Benchmark

```bash
cd rust-port/wifi-densepose-rs
cargo bench --package wifi-densepose-signal
```

Expected throughput:
| Operation | Latency | Throughput |
|-----------|---------|------------|
| CSI Preprocessing (4x64) | ~5.19 us | 49-66 Melem/s |
| Phase Sanitization (4x64) | ~3.84 us | 67-85 Melem/s |
| Feature Extraction (4x64) | ~9.03 us | 7-11 Melem/s |
| Motion Detection | ~186 ns | -- |
| Full Pipeline | ~18.47 us | ~54,000 fps |

### Workspace crates

The Rust workspace contains 10 crates under `crates/`:

| Crate | Description |
|-------|-------------|
| `wifi-densepose-core` | Core types, traits, and domain models |
| `wifi-densepose-signal` | Signal processing (FFT, phase unwrapping, Doppler, correlation) |
| `wifi-densepose-nn` | Neural network inference (ONNX Runtime, candle, tch) |
| `wifi-densepose-api` | Axum-based HTTP/WebSocket API server |
| `wifi-densepose-db` | Database layer (SQLx, PostgreSQL, SQLite, Redis) |
| `wifi-densepose-config` | Configuration loading (env vars, YAML, TOML) |
| `wifi-densepose-hardware` | Hardware adapters (ESP32, Intel 5300, Atheros, UDP, PCAP) |
| `wifi-densepose-wasm` | WebAssembly bindings for browser deployment |
| `wifi-densepose-cli` | Command-line interface |
| `wifi-densepose-mat` | WiFi-Mat disaster response module (search and rescue) |

Build individual crates:

```bash
# Signal processing only
cargo build --release --package wifi-densepose-signal

# API server
cargo build --release --package wifi-densepose-api

# Disaster response module
cargo build --release --package wifi-densepose-mat

# WASM target (see Section 7 for full instructions)
cargo build --release --package wifi-densepose-wasm --target wasm32-unknown-unknown
```

---

## 4. Three.js Visualization

A browser-based 3D visualization dashboard that renders DensePose body models with 24 body parts, signal visualization, and environment rendering.

### Run

Open `ui/viz.html` directly in a browser:

```bash
# macOS
open ui/viz.html

# Linux
xdg-open ui/viz.html

# Or serve it locally
python3 -m http.server 3000 --directory ui
# Then open http://localhost:3000/viz.html
```

### WebSocket connection

The visualization connects to `ws://localhost:8000/ws/pose` for real-time pose data. If no server is running, it falls back to a demo mode with simulated data so you can still see the 3D rendering.

To see live data:

1. Start the API server (Python or Rust)
2. Open `ui/viz.html`
3. The dashboard will connect automatically

---

## 5. Docker Deployment

### Development (with hot-reload, Postgres, Redis, Prometheus, Grafana)

```bash
docker compose up
```

This starts:
- `wifi-densepose-dev` -- API server with `--reload`, debug logging, auth disabled (port 8000)
- `postgres` -- PostgreSQL 15 (port 5432)
- `redis` -- Redis 7 with AOF persistence (port 6379)
- `prometheus` -- metrics scraping (port 9090)
- `grafana` -- dashboards (port 3000, login: admin/admin)
- `nginx` -- reverse proxy (ports 80, 443)

```bash
# View logs
docker compose logs -f wifi-densepose

# Run tests inside the container
docker compose exec wifi-densepose pytest tests/ -v

# Stop everything
docker compose down

# Stop and remove volumes
docker compose down -v
```

### Production

Uses the production Dockerfile stage with 4 uvicorn workers, auth enabled, rate limiting, and resource limits.

```bash
# Build production image
docker build --target production -t wifi-densepose:latest .

# Run standalone
docker run -d \
  --name wifi-densepose \
  -p 8000:8000 \
  -e ENVIRONMENT=production \
  -e SECRET_KEY=your-secret-key \
  wifi-densepose:latest
```

For the full production stack with Docker Swarm secrets:

```bash
# Create required secrets first
echo "db_password_here" | docker secret create db_password -
echo "redis_password_here" | docker secret create redis_password -
echo "jwt_secret_here" | docker secret create jwt_secret -
echo "api_key_here" | docker secret create api_key -
echo "grafana_password_here" | docker secret create grafana_password -

# Set required environment variables
export DATABASE_URL=postgresql://wifi_user:db_password_here@postgres:5432/wifi_densepose
export REDIS_URL=redis://redis:6379/0
export SECRET_KEY=your-secret-key
export JWT_SECRET=your-jwt-secret
export ALLOWED_HOSTS=your-domain.com
export POSTGRES_DB=wifi_densepose
export POSTGRES_USER=wifi_user

# Deploy with Docker Swarm
docker stack deploy -c docker-compose.prod.yml wifi-densepose
```

Production compose includes:
- 3 API server replicas with rolling updates and rollback
- Resource limits (2 CPU, 4GB RAM per replica)
- Health checks on all services
- JSON file logging with rotation
- Separate monitoring network (overlay)
- Prometheus with alerting rules and 15-day retention
- Grafana with provisioned datasources and dashboards

### Dockerfile stages

The multi-stage `Dockerfile` provides four targets:

| Target | Use | Command |
|--------|-----|---------|
| `development` | Local dev with hot-reload | `docker build --target development .` |
| `production` | Optimized production image | `docker build --target production .` |
| `testing` | Runs pytest during build | `docker build --target testing .` |
| `security` | Runs safety + bandit scans | `docker build --target security .` |

---

## 6. ESP32 Hardware Setup

Uses ESP32-S3 boards as WiFi CSI sensor nodes. See [ADR-012](adr/ADR-012-esp32-csi-sensor-mesh.md) for the full specification.

### Tested Hardware

| Board | Flash | MAC (example) | Role | Cost |
|-------|-------|---------------|------|------|
| FireBeetle ESP32-S3 rev v0.1 | 8 MB | 34:85:18:42:08:14 | CSI node 1 | ~$9 |
| GOOUUU Tech R16N8 ESP32-S3 rev v0.2 | 16 MB | 98:a3:16:e3:88:e8 | CSI node 2 | ~$9 |
| GOOUUU Tech R16N8 ESP32-S3 rev v0.2 | 16 MB | 98:a3:16:e7:fe:30 | CSI node 3 | ~$9 |
| Raspberry Pi 5 (16 GB) | — | — | Aggregator hub | ~$80 |

**Not supported:** Original ESP32, ESP32-C3 (single-core — cannot run the CSI DSP pipeline).

### Bill of Materials (Starter Kit)

| Item | Qty | Unit Cost | Total |
|------|-----|-----------|-------|
| ESP32-S3 board (any variant with 4+ MB flash) | 3 | ~$9 | ~$27 |
| USB-A to USB-C cables | 3 | $3 | $9 |
| USB power adapter (multi-port) | 1 | $15 | $15 |
| Aggregator (RPi 5 or laptop) | 1 | $0 (existing) | $0 |
| WiFi router **without client isolation** | 1 | $0 (existing) | $0 |
| **Total** | | | **~$51** |

> **Router requirement:** The AP must allow WiFi client-to-client traffic. Many public/guest AP modes and some ISP routers enable _client isolation_, which blocks UDP between connected devices. Use a home router, mobile hotspot, or Pi as AP. Test with `ping <node-ip>` from the hub — if 100% packet loss, isolation is on.

### Prerequisites

Install ESP-IDF (Espressif's official development framework):

```bash
# Clone ESP-IDF
mkdir -p ~/esp
cd ~/esp
git clone --recursive https://github.com/espressif/esp-idf.git
cd esp-idf
git checkout v5.2  # Pin to tested version

# Install tools
./install.sh esp32s3

# Activate environment (run each session)
. ./export.sh
```

### Flash a node

There are two steps: **flash** the firmware binary, then **provision** WiFi credentials and hub address over serial.

#### Step 1: Flash firmware

```bash
# 8 MB boards (FireBeetle ESP32-S3, most DevKit variants)
idf.py -p /dev/ttyACM0 -b 115200 flash
# Note: use 115200 baud for reliability — higher speeds cause CRC errors on some cables

# Or for 4 MB boards:
cp sdkconfig.defaults.4mb sdkconfig.defaults
idf.py -p /dev/ttyACM0 -b 115200 flash
```

Pre-built binaries are in releases/ and on the GitHub Releases page. To flash from a pre-built release:

```bash
esptool.py --chip esp32s3 --port /dev/ttyACM0 --baud 115200 write_flash \
  0x0000  releases/bootloader.bin \
  0x8000  releases/partition-table.bin \
  0xd000  releases/ota_data_initial.bin \
  0x10000 releases/esp32-csi-node.bin
```

#### Step 2: Provision credentials (NVS-based)

Credentials are stored in NVS flash (survives OTA updates). The `provision.py` script handles this over serial — no menuconfig or firmware rebuild required:

```bash
# Plug in each node via USB, then:
python3 firmware/esp32-csi-node/provision.py \
  --port /dev/ttyACM0 \
  --ssid "YourWiFi" \
  --password "YourPass" \
  --target-ip 192.168.1.20 \
  --target-port 5005 \
  --node-id 1 \
  --edge-tier 0 \
  --subk-count 8
```

Repeat for each node, incrementing `--node-id` (1, 2, 3, ...).

**Provision script options:**

| Flag | Description |
|------|-------------|
| `--port` | Serial port (`/dev/ttyACM0`, `/dev/ttyUSB0`, `COM7`) |
| `--ssid` | WiFi SSID to connect to |
| `--password` | WiFi password |
| `--target-ip` | Hub/aggregator IP that receives UDP CSI frames |
| `--target-port` | UDP port on hub (default: 5005) |
| `--node-id` | Integer 1–255, unique per node |
| `--edge-tier` | 0 = full CSI, 1 = feature-only (reduced bandwidth) |
| `--subk-count` | Subcarrier count per frame (8 for compact, 56 for full) |

#### Monitor serial output

```bash
python3 -m serial.tools.miniterm /dev/ttyACM0 115200
```

You should see the node connect to WiFi, resolve the hub IP, and begin logging `csi_send ok` lines. Press `Ctrl+]` to exit.

Repeat for each node.

### Firmware project structure

```
firmware/esp32-csi-node/
  CMakeLists.txt
  sdkconfig.defaults          # Menuconfig defaults with CSI enabled
  main/
    main.c                    # Entry point, WiFi init, CSI callback
    csi_collector.c           # CSI data collection and buffering
    feature_extract.c         # On-device FFT and feature extraction
    stream_sender.c           # UDP stream to aggregator
    config.h                  # Node configuration
    Kconfig.projbuild         # Menuconfig options
  components/
    esp_dsp/                  # Espressif DSP library for FFT
```

Each node does on-device feature extraction (raw I/Q to amplitude + phase + spectral bands), reducing bandwidth from ~11 KB/frame to ~470 bytes/frame. Nodes stream features via UDP to the aggregator.

### Run the aggregator (sensing server)

The sensing server is the aggregator. It listens on a UDP port for CSI frames, fuses data from all nodes, and serves the web UI + WebSocket.

```bash
# Native binary (built above)
./rust-port/wifi-densepose-rs/target/release/sensing-server \
  --http-port 3000 \
  --ws-port 3001 \
  --udp-port 5005 \
  --bind-addr 0.0.0.0 \
  --source esp32 \
  --ui-path ./ui

# Open the UI: http://<hub-ip>:3000/ui/index.html
```

Watch the health endpoint to confirm frames are arriving:

```bash
watch -n 1 'curl -s http://localhost:3000/health'
# "tick" increments each time a CSI frame is processed
# tick > 0 means real data is flowing
```

### Verify with real hardware

After provisioning at least one node on the same network as the hub:

```bash
# Check tick is climbing
curl -s http://localhost:3000/health
# Expected with live nodes: {"status":"ok","source":"esp32","clients":0,"tick":42}

# Check a sensing frame
curl -s http://localhost:3000/api/v1/sensing/latest
```

If tick stays at 0, check for AP client isolation (see Router requirement above) and verify the hub IP in the node NVS matches the host running the sensing server.

### What the ESP32 mesh can and cannot detect

| Capability | 1 Node | 3 Nodes | 6 Nodes |
|-----------|--------|---------|---------|
| Presence detection | Good | Excellent | Excellent |
| Coarse motion | Good | Excellent | Excellent |
| Room-level location | None | Good | Excellent |
| Respiration | Marginal | Good | Good |
| Heartbeat | Poor | Poor | Marginal |
| Multi-person count | None | Marginal | Good |
| Pose estimation | None | Poor | Marginal |

---

## 7. Environment-Specific Builds

### Browser (WASM)

Compiles the Rust pipeline to WebAssembly for in-browser execution. See [ADR-009](adr/ADR-009-rvf-wasm-runtime-edge-deployment.md) for the edge deployment architecture.

Prerequisites:

```bash
# Install wasm-pack
curl https://rustwasm.github.io/wasm-pack/installer/init.sh -sSf | sh

# Or via cargo
cargo install wasm-pack

# Add the WASM target
rustup target add wasm32-unknown-unknown
```

Build:

```bash
cd rust-port/wifi-densepose-rs

# Build WASM package (outputs to pkg/)
wasm-pack build crates/wifi-densepose-wasm --target web --release

# Build with disaster response module included
wasm-pack build crates/wifi-densepose-wasm --target web --release -- --features mat
```

The output `pkg/` directory contains `.wasm`, `.js` glue, and TypeScript definitions. Import in a web project:

```javascript
import init, { WifiDensePoseWasm } from './pkg/wifi_densepose_wasm.js';

async function main() {
  await init();
  const processor = new WifiDensePoseWasm();
  const result = processor.process_frame(csiJsonString);
  console.log(JSON.parse(result));
}
main();
```

Run WASM tests:

```bash
wasm-pack test --headless --chrome crates/wifi-densepose-wasm
```

Container size targets by deployment profile:

| Profile | Size | Suitable For |
|---------|------|-------------|
| Browser (int8 quantization) | ~10 MB | Chrome/Firefox dashboard |
| IoT (int4 quantization) | ~0.7 MB | ESP32, constrained devices |
| Mobile (int8 quantization) | ~6 MB | iOS/Android WebView |
| Field (fp16 quantization) | ~62 MB | Offline disaster tablets |

### IoT (ESP32)

See [Section 6](#6-esp32-hardware-setup) for full ESP32 setup. The firmware runs on the device itself (C, compiled with ESP-IDF). The Rust aggregator runs on a host machine.

For deploying the WASM runtime to a Raspberry Pi or similar:

```bash
# Cross-compile for ARM
rustup target add aarch64-unknown-linux-gnu
cargo build --release --package wifi-densepose-cli --target aarch64-unknown-linux-gnu
```

### Server (Docker)

See [Section 5](#5-docker-deployment).

Quick reference:

```bash
# Development
docker compose up

# Production standalone
docker build --target production -t wifi-densepose:latest .
docker run -d -p 8000:8000 wifi-densepose:latest

# Production stack (Swarm)
docker stack deploy -c docker-compose.prod.yml wifi-densepose
```

### Server (Direct -- no Docker)

```bash
# 1. Install Python dependencies
pip install -r requirements.txt

# 2. Set environment variables (copy from example.env)
cp example.env .env
# Edit .env with your settings

# 3. Run with uvicorn (production)
uvicorn v1.src.api.main:app \
  --host 0.0.0.0 \
  --port 8000 \
  --workers 4

# Or run the Rust API server
cd rust-port/wifi-densepose-rs
cargo run --release --package wifi-densepose-api
```

### Development (local with hot-reload)

Python:

```bash
# Create virtual environment
python3 -m venv venv
source venv/bin/activate

# Install all dependencies including dev tools
pip install -r requirements.txt

# Run with auto-reload
uvicorn v1.src.api.main:app --host 0.0.0.0 --port 8000 --reload

# Run verification in another terminal
./verify --verbose

# Run tests
pytest tests/ -v
pytest --cov=wifi_densepose --cov-report=html
```

Rust:

```bash
cd rust-port/wifi-densepose-rs

# Build in debug mode (faster compilation)
cargo build

# Run tests with output
cargo test --workspace -- --nocapture

# Watch mode (requires cargo-watch)
cargo install cargo-watch
cargo watch -x 'test --workspace' -x 'build --release'

# Run benchmarks
cargo bench --package wifi-densepose-signal
```

Both (visualization + API):

```bash
# Terminal 1: Start API server
uvicorn v1.src.api.main:app --host 0.0.0.0 --port 8000 --reload

# Terminal 2: Serve visualization
python3 -m http.server 3000 --directory ui

# Open http://localhost:3000/viz.html -- it connects to ws://localhost:8000/ws/pose
```

---

## Appendix: Key File Locations

| File | Purpose |
|------|---------|
| `./verify` | Trust kill switch -- one-command pipeline proof |
| `Makefile` | `make verify`, `make verify-verbose`, `make verify-audit` |
| `v1/requirements-lock.txt` | Pinned Python deps for hash reproducibility |
| `requirements.txt` | Full Python deps (API server, torch, etc.) |
| `v1/data/proof/verify.py` | Python verification script |
| `v1/data/proof/sample_csi_data.json` | Deterministic reference signal |
| `v1/data/proof/expected_features.sha256` | Published expected hash |
| `v1/src/api/main.py` | FastAPI application entry point |
| `v1/src/sensing/` | Commodity WiFi sensing module (RSSI) |
| `rust-port/wifi-densepose-rs/Cargo.toml` | Rust workspace root |
| `ui/viz.html` | Three.js 3D visualization |
| `Dockerfile` | Multi-stage Docker build (dev/prod/test/security) |
| `docker-compose.yml` | Development stack (Postgres, Redis, Prometheus, Grafana) |
| `docker-compose.prod.yml` | Production stack (Swarm, secrets, resource limits) |
| `docs/adr/ADR-009-rvf-wasm-runtime-edge-deployment.md` | WASM edge deployment architecture |
| `docs/adr/ADR-012-esp32-csi-sensor-mesh.md` | ESP32 firmware and mesh specification |
| `docs/adr/ADR-013-feature-level-sensing-commodity-gear.md` | Commodity WiFi (RSSI) sensing |
