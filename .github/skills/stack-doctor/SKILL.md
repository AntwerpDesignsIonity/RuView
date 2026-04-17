---
name: stack-doctor
description: 'Automated diagnosis and repair of the RuView sensing stack. Checks all services (Rust server, ESP32 nodes, Python API, UI), identifies failures, and applies fixes. Use when services fail to start, crash, or behave unexpectedly.'
---

# Stack Doctor

Automated health check and repair for the full Ionity RuView sensing platform.

## When to Use

- Any service fails to start or crashes
- Health endpoint returns errors or unexpected values
- ESP32 nodes not sending CSI data
- UI not loading or showing stale data
- WebSocket connections dropping
- After system restart or network change

## Full Stack Health Check

Run these checks in order. Report each result clearly.

### 1. Process Check

```bash
echo "=== Process Status ==="
pgrep -fa sensing-server && echo "SENSING_SERVER: RUNNING" || echo "SENSING_SERVER: STOPPED"
cat logs/sensing-server.pid 2>/dev/null && echo "" || echo "NO PID FILE"
```

### 2. Port Check

```bash
echo "=== Port Bindings ==="
ss -tlnp 2>/dev/null | grep -E ':(3000|3001|5005)\s' || echo "NO PORTS BOUND"
```

### 3. Health Endpoint

```bash
echo "=== Health Check ==="
curl -sf --max-time 5 http://localhost:3000/health | python3 -m json.tool 2>/dev/null || echo "HEALTH_ENDPOINT: UNREACHABLE"
```

### 4. UI Serving

```bash
echo "=== UI Check ==="
curl -sf --max-time 3 http://localhost:3000/ui/index.html | head -1 && echo "UI: SERVING" || echo "UI: NOT SERVING"
```

### 5. Log Analysis

```bash
echo "=== Recent Errors ==="
grep -i "error\|panic\|fatal\|failed" logs/sensing-server.log 2>/dev/null | tail -10 || echo "NO ERRORS IN LOG"
```

### 6. Binary Check

```bash
echo "=== Binary Status ==="
ls -la rust-port/wifi-densepose-rs/target/release/sensing-server 2>/dev/null && echo "BINARY: EXISTS" || echo "BINARY: MISSING"
```

### 7. Python Environment

```bash
echo "=== Python Venv ==="
test -f .venv/bin/activate && echo "VENV: EXISTS" || echo "VENV: MISSING"
source .venv/bin/activate 2>/dev/null && python3 -c "import numpy, scipy; print('PYTHON_DEPS: OK')" || echo "PYTHON_DEPS: MISSING"
```

## Repair Procedures

### Server Not Running

```bash
.ionity/ionity.sh stop    # Clean up stale processes
sleep 1
.ionity/ionity.sh run --yes
```

### Port Conflict

```bash
# Find and kill process on port 3000
lsof -ti :3000 | xargs kill -9 2>/dev/null
# Or change port:
echo 'HTTP_PORT=3002' >> .env.local
```

### Binary Missing or Outdated

```bash
cd rust-port/wifi-densepose-rs
cargo build -p wifi-densepose-sensing-server --release --no-default-features
```

### Python Venv Broken

```bash
rm -rf .venv
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### ESP32 Not Sending Data

1. Check ESP32 is powered and LED indicates WiFi connected (solid green)
2. Verify provisioned target_ip matches this machine: `hostname -I | awk '{print $1}'`
3. Check ESP32 serial output: `platformio device monitor --baud 460800`
4. Re-provision if needed: `python firmware/esp32-csi-node/provision.py --port /dev/ttyACM0 --ssid "WiFi" --password "pass" --target-ip $(hostname -I | awk '{print $1}')`

### After Network/IP Change

The ESP32 nodes target a specific hub IP. After an IP change:

1. Find new IP: `hostname -I | awk '{print $1}'`
2. Re-provision each ESP32 with the new target_ip
3. Restart the sensing server

## Verification

After any repair, run:

```bash
curl -sf http://localhost:3000/health | python3 -m json.tool
```

Expected: `{ "status": "ok", "source": "esp32", "clients": N, "tick": N }`

## Key References

- [Fault finding skill](.claude/skills/fault-finding-startups/SKILL.md) — detailed triage for each component
- [LED status codes](docs/led-indication.md) — ESP32 LED meanings
- [User guide](docs/user-guide.md) — end-user setup documentation
