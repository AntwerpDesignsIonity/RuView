# ESP32-S3 "AEDI" Radar Node — LED Visual Interface

Single WS2812B/NeoPixel RGB LED (GPIO 48, built-in on ESP32-S3 N16R8) used as the
complete visual status panel for the AEDI radar node. Communicates radar, network, and
system state without a screen using layered animations.

## Hardware

| Item | Value |
|------|-------|
| LED type | WS2812B (NeoPixel) |
| Data pin | GPIO 38 (GOOUUU N16R8 built-in RGB) |
| Color order | GRB |
| Max brightness | 255 (full power) |
| Framework | FastLED ≥ 3.6.0 |

> **Pin reference:** GOOUUU N16R8 = GPIO 38 · Official ESP32-S3-DevKitC-1 = GPIO 48 · SuperMini / some clones = GPIO 47

### Physical diffusion options

| Technique | Effect |
|-----------|--------|
| **Frosted "Eye"** — dome of frosted acrylic or milky PLA over the LED | Soft, ambient orb; eliminates harsh pin-point glare |
| **Edge-Lit Radar Wave** — clear acrylic slab with laser-etched concentric rings above the LED | Light travels through acrylic and illuminates only the etched lines |
| **Light Pipe Halo** — curved clear tubing wrapped around the inside of a cylindrical enclosure | Single LED appears as an illuminated ring/halo |

---

## Animation Architecture

The renderer uses a **two-layer** model evaluated every `loop()` iteration:

```
Layer 2 (overlay)  ──  10-s Yellow Beacon  /  20-s Role Ping
                                ↓ (overrides base when active)
Layer 1 (base)     ──  Boot → Calibrating → Conn. Established
                        → Idle → Human Detected → AI Processing
                        → Interference → OTA Update → Crash Fault
```

The overlay fires periodically and takes precedence over the base state for its duration.
Critical states (`CRASH_FAULT`, `BOOT_SWEEP`, `OTA_UPDATE`, `CALIBRATING`) suppress overlays.

---

## II. System Lifecycle & Critical States

### Boot Sequence — "Prism Sweep"

| Property | Value |
|----------|-------|
| Trigger | Power-on or reboot |
| Look | Fluid rainbow gradient: Red → Yellow → Green → Cyan → Blue → Magenta → White |
| Timing | 1 500 ms sweep → 500 ms maximum-brightness White flash → transitions to `CALIBRATING` |

### Calibrating / WiFi Connecting

| Property | Value |
|----------|-------|
| Trigger | Boot complete; still seeking WiFi SSID |
| Look | Slow expanding White sonar pulse (0% → 100% brightness, resets) |
| Timing | Repeats every 2 s until `WL_CONNECTED` |

### Connection Established — "System Healthy"

| Property | Value |
|----------|-------|
| Trigger | WiFi handshake complete; AI pipeline verified |
| Look | Solid bright **Emerald Green** `CRGB(0, 255, 50)` |
| Timing | Holds 1 500 ms → crisp double-blink → transitions to `IDLE_SCANNING` |

### Hardware Crash — "Fatal Fault"

| Property | Value |
|----------|-------|
| Trigger | Watchdog reset, brownout detection, or unrecoverable exception |
| Look | Chaotic arrhythmic strobe of pure **Red** `#FF0000` at maximum brightness |
| Timing | Random intervals — mimics a broken circuit; suppresses all overlays |

### OTA Update / Matrix Upload — "Do Not Unplug"

| Property | Value |
|----------|-------|
| Trigger | Firmware flashing or ONNX model download in progress |
| Look | Hypnotic **Cyan ↔ Magenta** oscillation |
| Timing | ~200 ms per crossfade (beatsin8 at 150 BPM); suppresses all overlays |

---

## III. Network Health, Roles & Signal Strength

### 10-Second Yellow Beacon

Fires once every 10 seconds, dominates the LED for exactly 1 second.

| RSSI | Look |
|------|------|
| 70–100 % (Strong) | Sharp aggressive spike to maximum **Yellow**; hold 200 ms, sharp fade-to-black by 600 ms |
| 30–69 % (Moderate) | "Lazy" Yellow swell — slow fade up to 50 % brightness, slow fade down |
| 1–29 % (Weak) | Dim, ghostly Yellow shimmer at ~10 % brightness with slight flicker |
| 0 % (Disconnected) | **Red** double-flash instead of Yellow |

> **Sonar Tail modifier** — the fade-out speed encodes latency: fast fade = low latency, long dragging fade = high latency.

### 20-Second Role Ping — "Node Identity"

Fires once every 20 seconds, offset by +5 s so it never overlaps the 10-s beacon.

| Role | Look |
|------|------|
| **Hub / Sink node** | Smooth double-ripple in **Deep Violet** (`DarkViolet`) — two fade-in/out cycles over 800 ms |
| **Edge / Transmitter node** | Crisp singular **Teal** blink at the 400–500 ms mark |

Set the role on-device with `provision.py`:

```bash
# Hub (receiving aggregator node)
python firmware/esp32-csi-node/provision.py --port /dev/ttyUSB0 --led-hub hub

# Edge sensor (transmitting node)
python firmware/esp32-csi-node/provision.py --port /dev/ttyUSB0 --led-hub edge
```

### Environmental Interference — "Orange Warning"

| Property | Value |
|----------|-------|
| Trigger | Excessive 2.4/5 GHz noise detected, or CSI multipath path blocked |
| Look | Slow continuous throb in **Burnt Orange** `#FF4500` |
| Timing | Overrides the Blue Baseline until environment clears or WiFi reconnects |

---

## IV. Radar Sensing & AI Activity

### Radar Calibration — "Baseline Sweep"

| Property | Value |
|----------|-------|
| Trigger | First boot in new room; mapping static environment |
| Look | Slow expanding **White** pulse (sonar-ping effect) |
| Timing | Same animation as WiFi connecting; distinguishable by context |

### Idle / Scanning — "Blue Baseline"

| Property | Value |
|----------|-------|
| Trigger | Default state — system online, no humans currently detected |
| Look | Dim rhythmic "breathing" in **Deep Blue** (`CHSV(160, 255, …)`) |
| Timing | Speed maps to CSI data throughput: fast/shallow pulses = high data rate; slow/deep = low rate |

### Human Detected — "Contact Pulse"

| Property | Value |
|----------|-------|
| Trigger | AI model scores a positive hit on human presence or movement |
| Look | **Warm Amber** `CHSV(32, 255, 255)` double heartbeat: thump-thump … pause |
| Timing | 150 ms on → 150 ms off → 150 ms on → 1 050 ms off (1 500 ms loop) |
| Optional | Sync to detected respiration or heart rate via mmWave / CSI vitals pipeline |

### Heavy AI Processing — "Synapse Spark"

| Property | Value |
|----------|-------|
| Trigger | Heavy matrix inference or MinCut algorithm running |
| Look | Rapid erratic flickering in **Magenta / Purple** `CHSV(192, …)` at varying brightness |
| Timing | Random pattern — resembles a tiny electrical storm |

---

## V. State Integration with Firmware

`currentState` is set externally by application logic via `setSystemState()`. Wire real events as follows:

| Event | Call |
|-------|------|
| OTA partition write begins | `setSystemState(OTA_UPDATE)` |
| OTA complete | `setSystemState(CALIBRATING)` → `CONN_ESTABLISHED` after reconnect |
| CSI model detection hit | `setSystemState(HUMAN_DETECTED)` |
| Heavy inference batch | `setSystemState(AI_PROCESSING)` |
| Interference detected | `setSystemState(INTERFERENCE)` |
| Panic handler | `setSystemState(CRASH_FAULT)` |

---

## VI. Build & Flash (IONITY PlatformIO)

```bash
# Build
cd ionity
pio run -e esp32s3_n16r8

# Flash
pio run -e esp32s3_n16r8 -t upload

# Monitor (460800 baud)
pio device monitor --baud 460800

# Set node role via NVS provisioning
python firmware/esp32-csi-node/provision.py --port /dev/ttyUSB0 --led-hub hub
python firmware/esp32-csi-node/provision.py --port /dev/ttyUSB0 --led-hub edge
```

Source: [ionity/src/main.cpp](../ionity/src/main.cpp)  
PlatformIO config: [ionity/platformio.ini](../ionity/platformio.ini)
