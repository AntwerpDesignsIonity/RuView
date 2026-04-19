# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed
- **Room Atlas — uncalibrated nodes no longer pile up at one point** ([ui/room-atlas.html](ui/room-atlas.html), [ui/room-atlas.js](ui/room-atlas.js)) — When the server returns the uncalibrated default position (every node at `[2.0, 0.0, 1.5]`), the atlas now detects the stack and auto-spreads nodes around the room perimeter as a *visual placeholder* (orange dashed ring + `(placeholder)` label in the node list). A new amber banner above the perspective view explains why and offers a one-click **Apply default layout** button that POSTs perimeter coordinates to `/api/v1/nodes/positions`. Uses the new shared feedback library for the busy/success/error UX. Result: the atlas is immediately useful on first boot instead of looking broken.

### Added
- **Per-node remote config (CSI rate / channel / TX power)** ([sensing-server main.rs](rust-port/wifi-densepose-rs/crates/wifi-densepose-sensing-server/src/main.rs), [ui/monitor.html](ui/monitor.html)) — New `POST /api/v1/node-config` / `GET /api/v1/node-config?node_id=N` / `GET /api/v1/node-config/all` endpoints let the operator change a node's CSI broadcast rate (1–200 Hz), radio channel (1–165) and TX power (−4..20 dBm) without reflashing. Writes are validated, persisted to `data/node-config.json`, and exposed in the `/api/v1/nodes` response (`config.revision` monotonically increases so firmware polling can short-circuit). New **Nodes** tab in monitor.html renders one row per node with editable Rate / Channel / TX Power inputs, an Apply button, and a bulk "apply rate to all active" control. Uses the shared toast layer for success/error feedback.
- **Server-side auto-calibration of node positions** ([sensing-server main.rs](rust-port/wifi-densepose-rs/crates/wifi-densepose-sensing-server/src/main.rs), [ui/room-atlas.html](ui/room-atlas.html), [ui/room-atlas.js](ui/room-atlas.js)) — New `POST /api/v1/calibration/auto` endpoint and a 30-second startup task that derive a deterministic perimeter layout from the set of currently active node IDs (any node seen within the last 10 s). The slot count is `max(active_id) + 1`, so individual nodes keep their corner across reboots even when others come or go. The Room Atlas calibration banner now exposes a one-click **Auto-calibrate** button next to the existing manual fallback. No env flags required; eliminates the manual layout step for the common single-room deployment.
- **AEDI Desktop scaffold (`desktop/`)** — Renamed from `releases/desktop/RuView.*`; 4 .NET 10 MAUI projects under [`desktop/AEDI.Desktop.sln`](desktop/AEDI.Desktop.sln). Brand assets wired in: [`desktop/AEDI.Desktop/Resources/Images/appicon.svg`](desktop/AEDI.Desktop/Resources/Images/appicon.svg), [`desktop/AEDI.Desktop/Resources/Splash/splash.svg`](desktop/AEDI.Desktop/Resources/Splash/splash.svg). Shared metadata in [`desktop/Directory.Build.props`](desktop/Directory.Build.props) including a `PREVIEW_ONLY` MSBuild constant for android/ios targets. New CI workflow [`.github/workflows/aedi-desktop-ci.yml`](.github/workflows/aedi-desktop-ci.yml) builds Windows + macOS on every PR touching `desktop/`. New `make desktop` and `make desktop-clean` targets in [`Makefile`](Makefile).
- **Shared UI feedback library `ui/shared/feedback.js`** ([feedback.js](ui/shared/feedback.js)) — Dependency-free `toast.{success,error,warn,info}`, `busy()` (long-running operation indicator with auto-resolve), `stickyError()` (persistent banner with copy-to-clipboard), and `fetchJSON()` (auto-toast on network/HTTP/JSON errors). Auto-loaded by [`ui/shared/nav.js`](ui/shared/nav.js) so every page using shared nav gets it. Self-mounting; no DOM scaffolding required.
- **`/api/v1/preprocessing/stats` endpoint** ([sensing-server main.rs](rust-port/wifi-densepose-rs/crates/wifi-densepose-sensing-server/src/main.rs)) — Reports the on/off state of five env-flag-gated CSI preprocessing stages (`hampel`, `phase_sanitizer`, `adaptive_denoise`, `csi_ratio`, `subcarrier_selection`). All default-off; verify hash unchanged. Wired into the [`ui/diagnostics.html`](ui/diagnostics.html) probe list and the new automated probe runner [`ui/tests/ui-smoke.html`](ui/tests/ui-smoke.html).
- **UI smoke test page** ([ui/tests/ui-smoke.html](ui/tests/ui-smoke.html)) — One-click probe runner that hits 17 endpoints + assets and shows pass/fail with latency. Designed for ops-style "is this server healthy?" checks at `http://<host>:3000/ui/tests/ui-smoke.html`.
- **ADR-083: CSI Preprocessing Accuracy Roadmap** ([docs/adr/ADR-083-csi-accuracy-roadmap.md](docs/adr/ADR-083-csi-accuracy-roadmap.md)) — Documents the staged rollout plan for the five preprocessing modules: D.1 (this PR, scaffolding + endpoint), D.2 (per-stage `NodeState` integration + counters + UI page), D.3 (re-baseline verify hash, promote defaults).

### Changed
- **Rebrand: RuView → AEDI in desktop client + ADRs** —
  - [`desktop/`](desktop/) (was `releases/desktop/`): all 4 csproj files, all .cs namespaces/usings, the .sln file, and the package metadata are rebranded. `git mv` preserved history. Brand README still credits "rUv (ruvnet) and the original RuView project" upstream.
  - ADRs renamed: [`ADR-031`](docs/adr/ADR-031-aedi-sensing-first-rf-mode.md), [`ADR-056`](docs/adr/ADR-056-aedi-desktop-capabilities.md). [`ADR-081`](docs/adr/ADR-081-per-node-telemetry-sqlite.md) now references `aedi_logs.sqlite3` / `aedi_telemetry.sqlite3`.
  - [`.claude-flow/daemon-state.json`](.claude-flow/daemon-state.json) — stale macOS absolute paths replaced with workspace-relative paths.
  - **One-shot SQLite auto-migration** ([sensing-server main.rs](rust-port/wifi-densepose-rs/crates/wifi-densepose-sensing-server/src/main.rs) — `migrate_legacy_log_db()`): on every startup, if `logs/ruview_logs.sqlite3` (and `-wal`/`-shm` siblings) exist and the new `logs/aedi_logs.sqlite3` does not, the files are renamed in place. Idempotent; logs each migration at INFO.

### Added
- **Monitor Hub UI button repairs** ([ui/monitor.html](ui/monitor.html)) — Two Scripts tab endpoints fixed: `model/sona` → `model/sona/profiles` (was 404), `pose/zones` → `pose/zones/summary` (was 404). All 18 scripts now return 200 OK end-to-end. `runScript()` rewritten to show HTTP status, duration in ms, content-type-aware body rendering, and explicit ✓/✗ indicators (no more silent "No response from server" with no detail). Both `runScript` and `pingNode` now explicitly attached to `window` so inline `onclick="runScript(N)"` handlers work even under bundlers that scope top-level declarations.
- **Ping button** now reports rich state (`status`, `firmware_status`, `rssi`, `last_seen_ms`, `ml_trained_pct`) by looking up the node in the latest `/api/v1/nodes` response, then auto-switches to the Logs tab so the user actually sees the result.
- **Provision Wizard buttons** (`Flash & Provision`, `Write NVS Only`) now generate the correct two-step PlatformIO + `provision.py --no-firmware` command and **auto-copy to clipboard** (instead of emitting an obsolete `.ionity/ionity.sh provision …` invocation that didn't exist).
- **Start/Stop Server buttons** now show a clear instruction (`./.ionity/ionity.sh run --yes` / `stop`) when the optional bridge controller on `:3002` isn't running, and auto-switch to the Logs tab so the message is visible.
- **WebSocket auto-reconnect** ([ui/monitor.html](ui/monitor.html)) — `connectSignalWs()` now schedules a 3 s reconnect on `onclose`, so the live data stream heals itself across server restarts and brief network blips. Status text reflects the retry state.
- **Vitals tab parsing fix** — `pollVitals()` now reads the actual server payload shape `{ vital_signs: { heart_rate_bpm, breathing_rate_bpm, heartbeat_confidence, … } }` instead of the legacy flat `{ hr, rr, … }` that no longer matches `/api/v1/vital-signs`. HR/RR cards in the Vitals tab populate correctly again.
- **ESP32 tab column header** changed from "FW Version" to "Firmware" (now reflects the runtime-derived `firmware_status`, not a static version string).

### Added (firmware status / IP / ML calibration round)
- **Firmware status, node IP, and ML calibration in `/api/v1/nodes`** ([sensing-server main.rs](rust-port/wifi-densepose-rs/crates/wifi-densepose-sensing-server/src/main.rs)) — Each node now reports `ip` (captured from UDP source address on every CSI/vitals packet, no separate registration step needed), `firmware_status` (derived live from runtime signals: `streaming`=CSI+edge_vitals in last 10s, `csi_only`=CSI but no edge_vitals, `stale`=last seen >10s, `offline`=>60s), and four `ml_*` fields surfacing the per-node TinyML aggression baseline: `ml_trained_pct` (0.0–1.0 calibration progress), `ml_anomaly`, `ml_aggression`, `ml_samples`. The aggression module's per-node Welford baseline is naturally **placement-independent** — it learns each node's own RF environment online, so accuracy holds no matter where the ESP32 is mounted.
- **Monitor Hub UI updated** ([ui/monitor.html](ui/monitor.html)) — Overview "Active Nodes" table and ESP32 tab table now show IP, colour-coded firmware status badge (green=streaming, blue=csi_only, yellow=stale, red=offline), ML training progress %, and live anomaly %. New `badge-yellow` class added to badge palette. Caption explains the placement-independent online baseline.
- **Live Nodes table on `/ui/nodes.html`** — added IP and Firmware columns (now 9 columns). Both the column headers and the empty-state colspan updated.
- **`provision.py` hardening** ([firmware/esp32-csi-node/provision.py](firmware/esp32-csi-node/provision.py)) — `flash_firmware()` now refuses to flash if `release_bins/` is empty (clear error pointing to the working PlatformIO route) and refuses if `app.bin` is older than any source file under `main/` by more than 1h (mtime check). New `--allow-stale-firmware` escape hatch for emergencies. Prevents the node-id-collision footgun that wasted hours on node 2.
- **`firmware/esp32-csi-node/release_bins/README.md`** — Documents why the directory is empty by default, recommends PlatformIO + `--no-firmware` workflow, lists ESP-IDF rebuild steps for advanced users.

### Added (earlier this cycle)
- **TinyML on-line aggression / anomaly detector** ([crates/wifi-densepose-sensing-server/src/aggression.rs](rust-port/wifi-densepose-rs/crates/wifi-densepose-sensing-server/src/aggression.rs)) — Per-node Welford rolling baseline of `motion_band_power` plus short/long EMA pair. Emits three live numbers per WS broadcast: `anomaly` (z-score → tanh, environment outlier), `aggression` (positive derivative of EMA pair, escalation lead signal), and `predicted` (~30-sample-ahead forecast). Baseline persists to `data/aggression_baseline.json` so the trained 3D environment survives restarts. Atomic save every 60 s. 5 unit tests, all passing.
- **`GET /api/v1/aggression`** — Returns current aggregate (max anomaly across nodes, max aggression, average `trained_pct`) plus per-node detail. Served by [sensing-server main.rs](rust-port/wifi-densepose-rs/crates/wifi-densepose-sensing-server/src/main.rs).
- **`aggression` field added to `pose_data` WS payload** — UI clients can now render escalation badges / predict-banners straight off the existing pose stream.
- **`scripts/diagnose-node.py`** — Reachability + state diagnoser for a registered ESP32 node. Resolves MAC from `nodes.yaml`, looks up its IP via ARP, pings, probes OTA :8032, queries the hub `/api/v1/nodes`, and prints a single VERDICT line with the right remediation command. No mutation. Confirms there is no network re-provisioning path today, so off-network nodes still need a USB reflash.
- **CSI activity sparkline on Touch-LCD-2** ([ionity/src/lcd_st7789_240x320.h](ionity/src/lcd_st7789_240x320.h)) — 60-sample / ~12 s rolling mini-waterfall replaces the static activity bar. Bottom-anchored coloured bars (green = quiet, yellow = movement, red = high motion / interference) make CSI variance trend visible at a glance without needing the web UI.

### Fixed
- **`THREE.OrbitControls is not a constructor` in mobile gaussian-splats webview** ([ui/mobile/src/assets/webview/gaussian-splats.html](ui/mobile/src/assets/webview/gaussian-splats.html)) — Pinned three.js from `r165` to `r147` (last release that ships a global-attaching legacy `examples/js/controls/OrbitControls.js`; r150+ removed it). Added a defensive console.error shim that names the root cause if a future bundler swaps in a newer Three.js without ESM. The repo's other 3D pages (`viz.html`, `xyz-segment.html`, `tests/sdk-diagnostics.html`) already use the vendored r147 and were not affected.
- **Stale Touch-LCD-2 pin-map comment** ([ionity/src/lcd_st7789_240x320.h](ionity/src/lcd_st7789_240x320.h)) — Header doc block listed `DC=41 CS=42 RST=40 BL=15` (from the wrong Waveshare schematic), while the code already used the correct CircuitPython-verified map `DC=42 CS=45 RST=0 BL=1`. Comment now matches code; explicitly notes that `BL=15` would land on the battery-ADC pin and leave the screen dark.

### Notes
- **Node 2 provisioning RESOLVED**: 98:a3:16:e7:fe:30 now streams as `node_id=2` from 192.168.124.10. Root cause: the pre-built ESP-IDF binary in `firmware/esp32-csi-node/release_bins/` is older than the source tree and ignores the NVS `node_id` override (it always reports `node_id=1` regardless of NVS) — so `provision.py` (default flow) overwrote the working PlatformIO firmware with the stale pre-built bin and the device collided with real node 1. Workflow that worked: (1) `pio run -e esp32s3_n16r8 -t upload` to install the up-to-date Arduino/PlatformIO firmware which honors NVS correctly, (2) `provision.py … --no-firmware` to write only the NVS partition. Hub `/api/v1/nodes` confirms `[1, 2, 3, 4, 5, 6]` all `active`. The stale `release_bins/` blob should be rebuilt or removed; `provision.py` should warn when the on-device firmware is newer than the bundled binary.
- Build verified: `pio run -e esp32s3_touch_lcd_2` → SUCCESS, RAM 13.6 %, Flash 28.7 % (958 KB, under the 1100 KB CI gate).
- Sensing-server build verified: `cargo build -p wifi-densepose-sensing-server --no-default-features --bin sensing-server` clean, 21 warnings (20 pre-existing dead-code + nothing new from `aggression`).
- Unit tests: `cargo test -p wifi-densepose-sensing-server --no-default-features aggression::` → 5 passed, 0 failed.

## [v0.5.4-esp32] — 2026-04-02

### Added
- **ADR-069: ESP32 CSI → Cognitum Seed RVF ingest pipeline** — Live-validated pipeline connecting ESP32-S3 CSI sensing to Cognitum Seed (Pi Zero 2 W) edge intelligence appliance. 339 vectors ingested, 100% kNN validation, SHA-256 witness chain verified.
- **Feature vector packet (magic 0xC5110003)** — New 48-byte packet with 8 normalized dimensions (presence, motion, breathing, heart rate, phase variance, person count, fall, RSSI) sent at 1 Hz alongside vitals.
- **`scripts/seed_csi_bridge.py`** — Python bridge: UDP listener → HTTPS ingest with bearer token auth, `--validate` (kNN + PIR ground truth), `--stats`, `--compact` modes, hash-based vector IDs, NaN/inf rejection, source IP filtering, retry logic.
- **Arena Physica research** — 26 research documents in `docs/research/` covering Maxwell's equations in WiFi sensing, Arena Physica Studio analysis, SOTA WiFi sensing 2025-2026, GOAP implementation plan for ESP32 + Pi Zero.
- **Cognitum Seed MCP integration** — 114-tool MCP proxy enables AI assistants to query sensing state, vectors, witness chain, and device status directly.

### Fixed
- **Compressed frame magic collision** — Reassigned compressed frame magic from `0xC5110003` to `0xC5110005` to free `0xC5110003` for feature vectors.
- **Uninitialized `s_top_k[0]` read** — Guarded variance computation against `s_top_k_count == 0` in `send_feature_vector()`.
- **Presence score normalization** — Bridge now divides by 15.0 instead of clamping, preserving dynamic range for raw values 1.41-14.92.
- **Stale magic references** — Updated ADR-039, DDD model to reflect `0xC5110005` for compressed frames.

### Security
- **Credential exposure remediation** — Removed hardcoded WiFi passwords and bearer tokens from source files. Added NVS binary/CSV patterns to `.gitignore`. Environment variable fallback for bearer token.
- **NaN/Inf injection prevention** — Bridge validates all feature dimensions are finite before Seed ingest.
- **UDP source filtering** — `--allowed-sources` argument restricts packet acceptance to known ESP32 IPs.

### Changed
- Wire format table now includes 6 magic numbers: `0xC5110001` (raw), `0xC5110002` (vitals), `0xC5110003` (features), `0xC5110004` (WASM events), `0xC5110005` (compressed), `0xC5110006` (fused vitals).

## [v0.5.3-esp32] — 2026-03-30

### Added
- **Cross-node RSSI-weighted feature fusion** — Multiple ESP32 nodes fuse CSI features using RSSI-based weighting. Closer node gets higher weight. Reduces variance noise by 29%, keypoint jitter by 72%.
- **DynamicMinCut person separation** — Uses `ruvector_mincut::DynamicMinCut` on the subcarrier temporal correlation graph to detect independent motion clusters. Replaces variance-based heuristic for multi-person counting.
- **RSSI-based position tracking** — Skeleton position driven by RSSI differential between nodes. Walk between ESP32s and the skeleton follows you.
- **Per-node state pipeline (ADR-068)** — Each ESP32 node gets independent `HashMap<u8, NodeState>` with frame history, classification, vitals, and person count. Fixes #249 (the #1 user-reported issue).
- **RuVector Phase 1-3 integration** — Subcarrier importance weighting, temporal keypoint smoothing (EMA), coherence gating, skeleton kinematic constraints (Jakobsen relaxation), compressed pose history.
- **Client-side lerp smoothing** — UI keypoints interpolate between frames (alpha=0.15) for fluid skeleton movement.
- **Multi-node mesh tests** — 8 integration tests covering 1-255 node configurations.
- **`wifi_densepose` Python package** — `from wifi_densepose import WiFiDensePose` now works (#314).

### Fixed
- **Watchdog crash on busy LANs (#321)** — Batch-limited edge_dsp to 4 frames before 20ms yield. Fixed idle-path busy-spin (`pdMS_TO_TICKS(5)==0`).
- **No detection from edge vitals (#323)** — Server now generates `sensing_update` from Tier 2+ vitals packets.
- **RSSI byte offset mismatch (#332)** — Server parsed RSSI from wrong byte (was reading sequence counter).
- **Stack overflow risk** — Moved 4KB of BPM scratch buffers from stack to static storage.
- **Stale node memory leak** — `node_states` HashMap evicts nodes inactive >60s.
- **Unsafe raw pointer removed** — Replaced with safe `.clone()` for adaptive model borrow.
- **Firmware CI** — Upgraded to IDF v5.4, replaced `xxd` with `od` (#327).
- **Person count double-counting** — Multi-node aggregation changed from `sum` to `max`.
- **Skeleton jitter** — Removed tick-based noise, dampened procedural animation, recalibrated feature scaling for real ESP32 data.

### Changed
- Motion-responsive skeleton: arm swing (0-80px) driven by CSI variance, leg kick (0-50px) by motion_band_power, vertical bob when walking.
- Person count thresholds recalibrated for real ESP32 hardware (1→2 at 0.70, EMA alpha 0.04).
- Vital sign filtering: larger median window (31), faster EMA (0.05), looser HR jump filter (15 BPM).
- Vendored ruvector updated to v2.1.0-40 (316 commits ahead).

### Benchmarks (2-node mesh, COM6 + COM9, 30s)
| Metric | Baseline | v0.5.3 | Improvement |
|--------|----------|--------|-------------|
| Variance noise | 109.4 | 77.6 | **-29%** |
| Feature stability | std=154.1 | std=105.4 | **-32%** |
| Keypoint jitter | std=4.5px | std=1.3px | **-72%** |
| Confidence | 0.643 | 0.686 | **+7%** |
| Presence accuracy | 93.4% | 94.6% | **+1.3pp** |

### Verified
- Real hardware: COM6 (node 1) + COM9 (node 2) on ruv.net WiFi
- All 284 Rust tests pass, 352 signal crate tests pass
- Firmware builds clean at 843 KB
- QEMU CI: 11/11 jobs green

## [v0.5.2-esp32] — 2026-03-28

### Fixed
- RSSI byte offset in frame parser (#332)
- Per-node state pipeline for multi-node sensing (#249)
- Firmware CI upgraded to IDF v5.4 (#327)

## [v0.5.1-esp32] — 2026-03-27

### Fixed
- Watchdog crash on busy LANs (#321)
- No detection from edge vitals (#323)
- `wifi_densepose` Python package import (#314)
- Pre-compiled firmware binaries added to release

## [v0.5.0-esp32] — 2026-03-15

### Added
- **60 GHz mmWave sensor fusion (ADR-063)** — Auto-detects Seeed MR60BHA2 (60 GHz, HR/BR/presence) and HLK-LD2410 (24 GHz, presence/distance) on UART at boot. Probes 115200 then 256000 baud, registers device capabilities, starts background parser.
- **48-byte fused vitals packet** (magic `0xC5110004`) — Kalman-style fusion: mmWave 80% + CSI 20% when both available. Automatic fallback to standard 32-byte CSI-only packet.
- **Server-side fusion bridge** (`scripts/mmwave_fusion_bridge.py`) — Reads two serial ports simultaneously for dual-sensor setups where mmWave runs on a separate ESP32.
- **Multimodal ambient intelligence roadmap (ADR-064)** — 25+ applications from fall detection to sleep monitoring to RF tomography.

### Verified
- Real hardware: ESP32-S3 (COM7) WiFi CSI + ESP32-C6/MR60BHA2 (COM4) 60 GHz mmWave running concurrently. HR=75 bpm, BR=25/min at 52 cm range. All 11 QEMU CI jobs green.

## [v0.4.3-esp32] — 2026-03-15

### Fixed
- **Fall detection false positives (#263)** — Default threshold raised from 2.0 to 15.0 rad/s²; normal walking (2-5 rad/s²) no longer triggers alerts. Added 3-consecutive-frame debounce and 5-second cooldown between alerts. Verified on real ESP32-S3 hardware: 0 false alerts in 60s / 1,300+ live WiFi CSI frames.
- **Kconfig default mismatch** — `CONFIG_EDGE_FALL_THRESH` Kconfig default was still 2000 (=2.0) while `nvs_config.c` fallback was updated to 15.0. Fixed Kconfig to 15000. Caught by real hardware testing — mock data did not reproduce.
- **provision.py NVS generator API change** — `esp_idf_nvs_partition_gen` package changed its `generate()` signature; switched to subprocess-first invocation for cross-version compatibility.
- **QEMU CI pipeline (11 jobs)** — Fixed all failures: fuzz test `esp_timer` stubs, QEMU `libgcrypt` dependency, NVS matrix generator, IDF container `pip` path, flash image padding, validation WARN handling, swarm `ip`/`cargo` missing.

### Added
- **4MB flash support (#265)** — `partitions_4mb.csv` and `sdkconfig.defaults.4mb` for ESP32-S3 boards with 4MB flash (e.g. SuperMini). Dual OTA slots, 1.856 MB each. Thanks to @sebbu for the community workaround that confirmed feasibility.
- **`--strict` flag** for `validate_qemu_output.py` — WARNs now pass by default in CI (no real WiFi in QEMU); use `--strict` to fail on warnings.

## [Unreleased]

### Added
- **QEMU ESP32-S3 testing platform (ADR-061)** — 9-layer firmware testing without hardware
  - Mock CSI generator with 10 physics-based scenarios (empty room, walking, fall, multi-person, etc.)
  - Single-node QEMU runner with 16-check UART validation
  - Multi-node TDM mesh simulation (TAP networking, 2-6 nodes)
  - GDB remote debugging with VS Code integration
  - Code coverage via gcov/lcov + apptrace
  - Fuzz testing (3 libFuzzer targets + ASAN/UBSAN)
  - NVS provisioning matrix (14 configs)
  - Snapshot-based regression testing (sub-second VM restore)
  - Chaos testing with fault injection + health monitoring
- **QEMU Swarm Configurator (ADR-062)** — YAML-driven multi-ESP32 test orchestration
  - 4 topologies: star, mesh, line, ring
  - 3 node roles: sensor, coordinator, gateway
  - 9 swarm-level assertions (boot, crashes, TDM, frame rate, fall detection, etc.)
  - 7 presets: smoke (2n/15s), standard (3n/60s), ci-matrix, large-mesh, line-relay, ring-fault, heterogeneous
  - Health oracle with cross-node validation
- **QEMU installer** (`install-qemu.sh`) — auto-detects OS, installs deps, builds Espressif QEMU fork
- **Unified QEMU CLI** (`qemu-cli.sh`) — single entry point for all 11 QEMU test commands
- CI: `firmware-qemu.yml` workflow with QEMU test matrix, fuzz testing, NVS validation, and swarm test jobs
- User guide: QEMU testing and swarm configurator section with plain-language walkthrough
- Native sensing-server startup now exports tracing defaults and supports selectable output via `SENSING_TRACE_FORMAT`

### Changed
- The Rust sensing server now emits HTTP request traces through `tower_http::trace::TraceLayer` for both HTTP and WebSocket listeners

### Fixed
- Firmware now boots in QEMU: WiFi/UDP/OTA/display guards for mock CSI mode
- 9 bugs in mock_csi.c (LFSR bias, MAC filter init, scenario loop, overflow burst timing)
- `scripts/start-aedi-s.sh` now uses the current sensing-server CLI flags (`--bind-addr`, `--http-port`, `--udp-port`, `--ui-path`) instead of stale options that prevented startup
- 23 bugs from ADR-061 deep review (inject_fault.py writes, CI cache, snapshot log corruption, etc.)
- 16 bugs from ADR-062 deep review (log filename mismatch, SLIRP port collision, heap false positives, etc.)
- All scripts: `--help` flags, prerequisite checks with install hints, standardized exit codes

- **Sensing server UI API completion (ADR-043)** — 14 fully-functional REST endpoints for model management, CSI recording, and training control
  - Model CRUD: `GET /api/v1/models`, `GET /api/v1/models/active`, `POST /api/v1/models/load`, `POST /api/v1/models/unload`, `DELETE /api/v1/models/:id`, `GET /api/v1/models/lora/profiles`, `POST /api/v1/models/lora/activate`
  - CSI recording: `GET /api/v1/recording/list`, `POST /api/v1/recording/start`, `POST /api/v1/recording/stop`, `DELETE /api/v1/recording/:id`
  - Training control: `GET /api/v1/train/status`, `POST /api/v1/train/start`, `POST /api/v1/train/stop`
  - Recording writes CSI frames to `.jsonl` files via tokio background task
  - Model/recording directories scanned at startup, state managed via `Arc<RwLock<AppStateInner>>`
- **ADR-044: Provisioning tool enhancements** — 5-phase plan for complete NVS coverage (7 missing keys), JSON config files, mesh presets, read-back/verify, and auto-detect
- **25 real mobile tests** replacing `it.todo()` placeholders — 205 assertions covering components, services, stores, hooks, screens, and utils
- **Project MERIDIAN (ADR-027)** — Cross-environment domain generalization for WiFi pose estimation (1,858 lines, 72 tests)
  - `HardwareNormalizer` — Catmull-Rom cubic interpolation resamples any hardware CSI to canonical 56 subcarriers; z-score + phase sanitization
  - `DomainFactorizer` + `GradientReversalLayer` — adversarial disentanglement of pose-relevant vs environment-specific features
  - `GeometryEncoder` + `FilmLayer` — Fourier positional encoding + DeepSets + FiLM for zero-shot deployment given AP positions
  - `VirtualDomainAugmentor` — synthetic environment diversity (room scale, wall material, scatterers, noise) for 4x training augmentation
  - `RapidAdaptation` — 10-second unsupervised calibration via contrastive test-time training + LoRA adapters
  - `CrossDomainEvaluator` — 6-metric evaluation protocol (MPJPE in-domain/cross-domain/few-shot/cross-hardware, domain gap ratio, adaptation speedup)
- ADR-027: Cross-Environment Domain Generalization — 10 SOTA citations (PerceptAlign, X-Fi ICLR 2025, AM-FM, DGSense, CVPR 2024)
- **Cross-platform RSSI adapters** — macOS CoreWLAN (`MacosCoreWlanScanner`) and Linux `iw` (`LinuxIwScanner`) Rust adapters with `#[cfg(target_os)]` gating
- macOS CoreWLAN Python sensing adapter with Swift helper (`mac_wifi.swift`)
- macOS synthetic BSSID generation (FNV-1a hash) for Sonoma 14.4+ BSSID redaction
- Linux `iw dev <iface> scan` parser with freq-to-channel conversion and `scan dump` (no-root) mode
- ADR-025: macOS CoreWLAN WiFi Sensing (ORCA)

### Fixed
- **sendto ENOMEM crash (Issue #127)** — CSI callbacks in promiscuous mode exhaust lwIP pbuf pool causing guru meditation crash. Fixed with 50 Hz rate limiter in `csi_collector.c` and 100 ms ENOMEM backoff in `stream_sender.c`. Hardware-verified on ESP32-S3 (200+ callbacks, zero crashes)
- **Provisioning script missing TDM/edge flags (Issue #130)** — Added `--tdm-slot`, `--tdm-total`, `--edge-tier`, `--pres-thresh`, `--fall-thresh`, `--vital-win`, `--vital-int`, `--subk-count` to `provision.py`
- **WebSocket "RECONNECTING" on Dashboard/Live Demo** — `sensingService.start()` now called on app init in `app.js` so WebSocket connects immediately instead of waiting for Sensing tab visit
- **Mobile WebSocket port** — `ws.service.ts` `buildWsUrl()` uses same-origin port instead of hardcoded port 3001
- **Mobile Jest config** — `testPathIgnorePatterns` no longer silently ignores the entire test directory
- Removed synthetic byte counters from Python `MacosWifiCollector` — now reports `tx_bytes=0, rx_bytes=0` instead of fake incrementing values

---

## [3.0.0] - 2026-03-01

Major release: AETHER contrastive embedding model, Docker Hub images, and comprehensive UI overhaul.

### Added — AETHER Contrastive Embedding Model (ADR-024)
- **Project AETHER** — self-supervised contrastive learning for WiFi CSI fingerprinting, similarity search, and anomaly detection (`9bbe956`)
- `embedding.rs` module: `ProjectionHead`, `InfoNceLoss`, `CsiAugmenter`, `FingerprintIndex`, `PoseEncoder`, `EmbeddingExtractor` (909 lines, zero external ML dependencies)
- SimCLR-style pretraining with 5 physically-motivated augmentations (temporal jitter, subcarrier masking, Gaussian noise, phase rotation, amplitude scaling)
- CLI flags: `--pretrain`, `--pretrain-epochs`, `--embed`, `--build-index <type>`
- Four HNSW-compatible fingerprint index types: `env_fingerprint`, `activity_pattern`, `temporal_baseline`, `person_track`
- Cross-modal `PoseEncoder` for WiFi-to-camera embedding alignment
- VICReg regularization for embedding collapse prevention
- 53K total parameters (55 KB at INT8) — fits on ESP32

### Added — Docker & Deployment
- Published Docker Hub images: `ruvnet/wifi-densepose:latest` (132 MB Rust) and `ruvnet/wifi-densepose:python` (569 MB) (`add9f19`)
- Multi-stage Dockerfile for Rust sensing server with RuVector crates
- `docker-compose.yml` orchestrating both Rust and Python services
- RVF model export via `--export-rvf` and load via `--load-rvf` CLI flags

### Added — Documentation
- 33 use cases across 4 vertical tiers: Everyday, Specialized, Robotics & Industrial, Extreme (`0afd9c5`)
- "Why WiFi Wins" comparison table (WiFi vs camera vs LIDAR vs wearable vs PIR)
- Mermaid architecture diagrams: end-to-end pipeline, signal processing detail, deployment topology (`50f0fc9`)
- Models & Training section with RuVector crate links (GitHub + crates.io), SONA component table (`965a1cc`)
- RVF container section with deployment targets table (ESP32 0.7 MB to server 50+ MB)
- Collapsible README sections for improved navigation (`478d964`, `99ec980`, `0ebd6be`)
- Installation and Quick Start moved above Table of Contents (`50acbf7`)
- CSI hardware requirement notice (`528b394`)

### Fixed
- **UI auto-detects server port from page origin** — no more hardcoded `localhost:8080`; works on any port (Docker :3000, native :8080, custom) (`3b72f35`, closes #55)
- **Docker port mismatch** — server now binds 3000/3001 inside container as documented (`44b9c30`)
- Added `/ws/sensing` WebSocket route to the HTTP server so UI only needs one port
- Fixed README API endpoint references: `/api/v1/health` → `/health`, `/api/v1/sensing` → `/api/v1/sensing/latest`
- Multi-person tracking limit corrected: configurable default 10, no hard software cap (`e2ce250`)

---

## [2.0.0] - 2026-02-28

Major release: complete Rust sensing server, full DensePose training pipeline, RuVector v2.0.4 integration, ESP32-S3 firmware, and 6 security hardening patches.

### Added — Rust Sensing Server
- **Full DensePose-compatible REST API** served by Axum (`d956c30`)
  - `GET /health` — server health
  - `GET /api/v1/sensing/latest` — live CSI sensing data
  - `GET /api/v1/vital-signs` — breathing rate (6-30 BPM) and heartbeat (40-120 BPM)
  - `GET /api/v1/pose/current` — 17 COCO keypoints derived from WiFi signal field
  - `GET /api/v1/info` — server build and feature info
  - `GET /api/v1/model/info` — RVF model container metadata
  - `ws://host/ws/sensing` — real-time WebSocket stream
- Three data sources: `--source esp32` (UDP CSI), `--source windows` (netsh RSSI), `--source simulated` (deterministic reference)
- Auto-detection: server probes ESP32 UDP and Windows WiFi, falls back to simulated
- Three.js visualization UI with 3D body skeleton, signal heatmap, phase plot, Doppler bars, vital signs panel
- Static UI serving via `--ui-path` flag
- Throughput: 9,520–11,665 frames/sec (release build)

### Added — ADR-021: Vital Sign Detection
- `VitalSignDetector` with breathing (6-30 BPM) and heartbeat (40-120 BPM) extraction from CSI fluctuations (`1192de9`)
- FFT-based spectral analysis with configurable band-pass filters
- Confidence scoring based on spectral peak prominence
- REST endpoint `/api/v1/vital-signs` with real-time JSON output

### Added — ADR-023: DensePose Training Pipeline (Phases 1-8)
- `wifi-densepose-train` crate with complete 8-phase pipeline (`fc409df`, `ec98e40`, `fce1271`)
  - Phase 1: `DataPipeline` with MM-Fi and Wi-Pose dataset loaders
  - Phase 2: `CsiToPoseTransformer` — 4-head cross-attention + 2-layer GCN on COCO skeleton
  - Phase 3: 6-term composite loss (MSE, bone length, symmetry, joint angle, temporal, confidence)
  - Phase 4: `DynamicPersonMatcher` via ruvector-mincut (O(n^1.5 log n) Hungarian assignment)
  - Phase 5: `SonaAdapter` — MicroLoRA rank-4 with EWC++ memory preservation
  - Phase 6: `SparseInference` — progressive 3-layer model loading (A: essential, B: refinement, C: full)
  - Phase 7: `RvfContainer` — single-file model packaging with segment-based binary format
  - Phase 8: End-to-end training with cosine-annealing LR, early stopping, checkpoint saving
- CLI: `--train`, `--dataset`, `--epochs`, `--save-rvf`, `--load-rvf`, `--export-rvf`
- Benchmark: ~11,665 fps inference, 229 tests passing

### Added — ADR-016: RuVector Training Integration (all 5 crates)
- `ruvector-mincut` → `DynamicPersonMatcher` in `metrics.rs` + subcarrier selection (`81ad09d`, `a7dd31c`)
- `ruvector-attn-mincut` → antenna attention in `model.rs` + noise-gated spectrogram
- `ruvector-temporal-tensor` → `CompressedCsiBuffer` in `dataset.rs` + compressed breathing/heartbeat
- `ruvector-solver` → sparse subcarrier interpolation (114→56) + Fresnel triangulation
- `ruvector-attention` → spatial attention in `model.rs` + attention-weighted BVP
- Vendored all 11 RuVector crates under `vendor/ruvector/` (`d803bfe`)

### Added — ADR-017: RuVector Signal & MAT Integration (7 integration points)
- `gate_spectrogram()` — attention-gated noise suppression (`18170d7`)
- `attention_weighted_bvp()` — sensitivity-weighted velocity profiles
- `mincut_subcarrier_partition()` — dynamic sensitive/insensitive subcarrier split
- `solve_fresnel_geometry()` — TX-body-RX distance estimation
- `CompressedBreathingBuffer` + `CompressedHeartbeatSpectrogram`
- `BreathingDetector` + `HeartbeatDetector` (MAT crate, real FFT + micro-Doppler)
- Feature-gated behind `cfg(feature = "ruvector")` (`ab2453e`)

### Added — ADR-018: ESP32-S3 Firmware & Live CSI Pipeline
- ESP32-S3 firmware with FreeRTOS CSI extraction (`92a5182`)
- ADR-018 binary frame format: `[0xAD, 0x18, len_hi, len_lo, payload]`
- Rust `Esp32Aggregator` receiving UDP frames on port 5005
- `bridge.rs` converting I/Q pairs to amplitude/phase vectors
- NVS provisioning for WiFi credentials
- Pre-built binary quick start documentation (`696a726`)

### Added — ADR-014: SOTA Signal Processing
- 6 algorithms, 83 tests (`fcb93cc`)
  - Hampel filter (median + MAD, resistant to 50% contamination)
  - Conjugate multiplication (reference-antenna ratio, cancels common-mode noise)
  - Phase sanitization (unwrap + linear detrend, removes CFO/SFO)
  - Fresnel zone geometry (TX-body-RX distance from first-principles physics)
  - Body Velocity Profile (micro-Doppler extraction, 5.7x speedup)
  - Attention-gated spectrogram (learned noise suppression)

### Added — ADR-015: Public Dataset Training Strategy
- MM-Fi and Wi-Pose dataset specifications with download links (`4babb32`, `5dc2f66`)
- Verified dataset dimensions, sampling rates, and annotation formats
- Cross-dataset evaluation protocol

### Added — WiFi-Mat Disaster Detection Module
- Multi-AP triangulation for through-wall survivor detection (`a17b630`, `6b20ff0`)
- Triage classification (breathing, heartbeat, motion)
- Domain events: `survivor_detected`, `survivor_updated`, `alert_created`
- WebSocket broadcast at `/ws/mat/stream`

### Added — Infrastructure
- Guided 7-step interactive installer with 8 hardware profiles (`8583f3e`)
- Comprehensive build guide for Linux, macOS, Windows, Docker, ESP32 (`45f8a0d`)
- 12 Architecture Decision Records (ADR-001 through ADR-012) (`337dd96`)

### Added — UI & Visualization
- Sensing-only UI mode with Gaussian splat visualization (`b7e0f07`)
- Three.js 3D body model (17 joints, 16 limbs) with signal-viz components
- Tabs: Dashboard, Hardware, Live Demo, Sensing, Architecture, Performance, Applications
- WebSocket client with automatic reconnection and exponential backoff

### Added — Rust Signal Processing Crate
- Complete Rust port of WiFi-DensePose with modular workspace (`6ed69a3`)
  - `wifi-densepose-signal` — CSI processing, phase sanitization, feature extraction
  - `wifi-densepose-core` — shared types and configuration
  - `wifi-densepose-nn` — neural network inference (DensePose head, RCNN)
  - `wifi-densepose-hardware` — ESP32 aggregator, hardware interfaces
  - `wifi-densepose-config` — configuration management
- Comprehensive benchmarks and validation tests (`3ccb301`)

### Added — Python Sensing Pipeline
- `WindowsWifiCollector` — RSSI collection via `netsh wlan show networks`
- `RssiFeatureExtractor` — variance, spectral bands (motion 0.5-4 Hz, breathing 0.1-0.5 Hz), change points
- `PresenceClassifier` — rule-based 3-state classification (ABSENT / PRESENT_STILL / ACTIVE)
- Cross-receiver agreement scoring for multi-AP confidence boosting
- WebSocket sensing server (`ws_server.py`) broadcasting JSON at 2 Hz
- Deterministic CSI proof bundles for reproducible verification (`v1/data/proof/`)
- Commodity sensing unit tests (`b391638`)

### Changed
- Rust hardware adapters now return explicit errors instead of silent empty data (`6e0e539`)

### Fixed
- Review fixes for end-to-end training pipeline (`45f0304`)
- Dockerfile paths updated from `src/` to `v1/src/` (`7872987`)
- IoT profile installer instructions updated for aggregator CLI (`f460097`)
- `process.env` reference removed from browser ES module (`e320bc9`)

### Performance
- 5.7x Doppler extraction speedup via optimized FFT windowing (`32c75c8`)
- Single 2.1 MB static binary, zero Python dependencies for Rust server

### Security
- Fix SQL injection in status command and migrations (`f9d125d`)
- Fix XSS vulnerabilities in UI components (`5db55fd`)
- Fix command injection in statusline.cjs (`4cb01fd`)
- Fix path traversal vulnerabilities (`896c4fc`)
- Fix insecure WebSocket connections — enforce wss:// on non-localhost (`ac094d4`)
- Fix GitHub Actions shell injection (`ab2e7b4`)
- Fix 10 additional vulnerabilities, remove 12 dead code instances (`7afdad0`)

---

## [1.1.0] - 2025-06-07

### Added
- Complete Python WiFi-DensePose system with CSI data extraction and router interface
- CSI processing and phase sanitization modules
- Batch processing for CSI data in `CSIProcessor` and `PhaseSanitizer`
- Hardware, pose, and stream services for WiFi-DensePose API
- Comprehensive CSS styles for UI components and dark mode support
- API and Deployment documentation

### Fixed
- Badge links for PyPI and Docker in README
- Async engine creation poolclass specification

---

## [1.0.0] - 2024-12-01

### Added
- Initial release of WiFi-DensePose
- Real-time WiFi-based human pose estimation using Channel State Information (CSI)
- DensePose neural network integration for body surface mapping
- RESTful API with comprehensive endpoint coverage
- WebSocket streaming for real-time pose data
- Multi-person tracking with configurable capacity (default 10, up to 50+)
- Fall detection and activity recognition
- Domain configurations: healthcare, fitness, smart home, security
- CLI interface for server management and configuration
- Hardware abstraction layer for multiple WiFi chipsets
- Phase sanitization and signal processing pipeline
- Authentication and rate limiting
- Background task management
- Cross-platform support (Linux, macOS, Windows)

### Documentation
- User guide and API reference
- Deployment and troubleshooting guides
- Hardware setup and calibration instructions
- Performance benchmarks
- Contributing guidelines

[Unreleased]: https://github.com/ruvnet/wifi-densepose/compare/v3.0.0...HEAD
[3.0.0]: https://github.com/ruvnet/wifi-densepose/compare/v2.0.0...v3.0.0
[2.0.0]: https://github.com/ruvnet/wifi-densepose/compare/v1.1.0...v2.0.0
[1.1.0]: https://github.com/ruvnet/wifi-densepose/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/ruvnet/wifi-densepose/releases/tag/v1.0.0
