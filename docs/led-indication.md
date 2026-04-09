# AEDI Node — LED Quick Reference

Single WS2812B RGB LED (GPIO 38 on GOOUUU N16R8) is the complete visual status panel.
Two-layer system: **base state** (always on) + **overlay** (fires periodically, takes priority).

---

## At a Glance

| Color | Animation | Layer | Meaning |
|-------|-----------|-------|---------|
| 🌈 Rainbow → White | Sweep then flash | Base | **Boot** — power-on / reboot in progress |
| ⚪ White | Slow sonar expanding pulse (2 s loop) | Base | **Calibrating** — searching for WiFi SSID |
| 🟢 Emerald Green | Solid → double-blink | Base | **Connected** — WiFi + AI pipeline verified |
| 🔵 Deep Blue | Slow rhythmic breathing | Base | **Idle / Scanning** — online, no human detected |
| 🟠 Burnt Orange | Slow continuous throb | Base | **Interference** — RF noise or CSI path blocked |
| 🟡 Warm Amber | Double heartbeat (thump-thump … pause) | Base | **Human Detected** — positive detection |
| 🟣 Magenta / Purple | Rapid erratic flicker | Base | **AI Processing** — heavy inference running |
| 🔵↔🩷 Cyan ↔ Magenta | Hypnotic crossfade (~200 ms) | Base | **OTA Update** — flashing / model download, do not unplug |
| 🔴 Red | Chaotic arrhythmic strobe | Base | **Crash / Fatal Fault** — watchdog / brownout / exception |

---

## Periodic Overlays (appear on top of base state)

| Interval | Color | Animation | Meaning |
|----------|-------|-----------|---------|
| Every **10 s** (1 s burst) | 🟡 Yellow — bright sharp spike | Aggressive spike, hold 200 ms, fast fade | WiFi **Strong** (RSSI 70–100 %) |
| Every **10 s** (1 s burst) | 🟡 Yellow — half-bright lazy swell | Slow fade up to 50 %, slow fade down | WiFi **Moderate** (RSSI 30–69 %) |
| Every **10 s** (1 s burst) | 🟡 Yellow — dim ghostly shimmer | ~10 % brightness with slight flicker | WiFi **Weak** (RSSI 1–29 %) |
| Every **10 s** (1 s burst) | 🔴 Red | Double-flash | WiFi **Disconnected** (RSSI 0 %) |
| Every **20 s** (+5 s offset) | 🟣 Deep Violet | Smooth double-ripple over 800 ms | Node role = **Hub / Sink** |
| Every **20 s** (+5 s offset) | 🩵 Teal | Crisp single blink at 400–500 ms | Node role = **Edge / Transmitter** |

> The 10 s beacon and 20 s role ping are offset so they never collide.
> **Fade speed** on the Yellow beacon also encodes latency: fast fade = low latency · slow dragging fade = high latency.

---

## Reading the Base Blue Breathing

The breathing speed of the idle **Deep Blue** state maps to CSI data throughput:

| Breathing speed | Throughput |
|-----------------|------------|
| Fast, shallow pulses | High data rate |
| Slow, deep pulses | Low data rate |

---

## State Priority (highest → lowest)

```
CRASH_FAULT        — always wins, suppresses everything
OTA_UPDATE         — suppresses overlays
CALIBRATING        — suppresses overlays
BOOT_SWEEP         — suppresses overlays
─────── Overlays fire above this line ───────
INTERFERENCE       — overrides Idle Blue
HUMAN_DETECTED
AI_PROCESSING
IDLE_SCANNING      — default base
CONN_ESTABLISHED   — transient (1.5 s then → Idle)
```

---

## Your Current Status Explained

| What you see | What it means |
|--------------|---------------|
| **Solid / breathing Blue** | Idle — online, no human detected |
| **Yellow pulse every ~10 s** | WiFi beacon — brightness = signal strength |
| **Blue/Teal blink every ~20 s** | Role ping — Teal = Edge node · Violet = Hub node |

All three patterns together = **node healthy, scanning, no presence detected**.

---

## Firmware State Calls

```cpp
setSystemState(OTA_UPDATE);       // OTA write begins
setSystemState(CALIBRATING);      // After OTA or on boot
setSystemState(HUMAN_DETECTED);   // CSI/AI positive hit
setSystemState(AI_PROCESSING);    // Heavy inference running
setSystemState(INTERFERENCE);     // RF noise detected
setSystemState(CRASH_FAULT);      // Panic handler
```
