# Ionity AEDI-S Launcher

**Ionity Global (Pty) Ltd, South Africa**
ai@ionity.today | [www.ionity.today](https://www.ionity.today) | +27 646 999 877

Single-command startup for the full AEDI-S WiFi sensing stack — Rust server, Python env, UI, ESP32 provisioning, and all analysis scripts.

---

## Run This First

From the repo root:

```bash
bash .ionity/ionity.sh deps
```

This installs everything needed: Rust, Python venv + packages, esptool, pyserial, Node.js, gcc/g++, and opens firewall ports. Safe to re-run any time.

---

## Then Start the Stack

```bash
bash .ionity/ionity.sh
```

That's it. Runs: deps → build (if needed) → sensing server. Opens at:

- **UI** → `http://<your-pi-ip>:3000/ui/index.html`
- **Health** → `http://<your-pi-ip>:3000/health`
- **WebSocket** → `ws://<your-pi-ip>:3001/ws/sensing`
- **ESP32 UDP** → `<your-pi-ip>:5005`

---

## Flash an ESP32 Node

Plug a node in via USB, then:

```bash
bash .ionity/ionity.sh provision
```

Interactive wizard — auto-detects USB port, prompts for node ID, reads WiFi credentials from `.env.local` automatically.

Or direct (no prompts):

```bash
bash .ionity/ionity.sh provision /dev/ttyACM0 1
```

---

## All Commands

| Command | What it does |
|---------|-------------|
| `bash .ionity/ionity.sh` | **Full auto** — deps + build + serve |
| `bash .ionity/ionity.sh deps` | **Run first** — check and install all dependencies |
| `bash .ionity/ionity.sh build` | Build Rust binary + Python venv only |
| `bash .ionity/ionity.sh provision` | Flash + provision ESP32 nodes (interactive) |
| `bash .ionity/ionity.sh provision /dev/ttyACM0 1` | Flash node 1 directly |
| `bash .ionity/ionity.sh docker` | Start via Docker instead of native binary |
| `bash .ionity/ionity.sh status` | Show server health, ports, ESP32 ARP table |
| `bash .ionity/ionity.sh stop` | Stop all Ionity services |
| `bash .ionity/ionity.sh scripts` | Browse and run 40+ analysis scripts |
| `bash .ionity/ionity.sh help` | Full usage reference |

---

## Typical First-Time Setup

```bash
# 1. Clone / enter repo
cd AEDI-S

# 2. Set credentials (copy once, never committed)
cp example.env .env.local
# Edit .env.local: set WIFI_SSID, WIFI_PASS, HUB_IP

# 3. Install all dependencies
bash .ionity/ionity.sh deps

# 4. Flash each ESP32 node (one at a time, via USB)
bash .ionity/ionity.sh provision

# 5. Start the stack
bash .ionity/ionity.sh
```

---

## File Structure

```
.ionity/
  ionity.sh          ← Single entry point — run this
  README.md          ← This file
  lib/
    deps.sh          ← Dependency checker + installer
    build.sh         ← Rust + Python build logic
    stack.sh         ← Start / stop / status
    provisioner.sh   ← Ionity Provisioner (ESP32 flash + NVS)
    scripts_menu.sh  ← 40-script interactive library
```

All credentials are read from `.env.local` in the repo root — never committed.

---

*Ionity Global (Pty) Ltd  |  ai@ionity.today  |  www.ionity.today*
