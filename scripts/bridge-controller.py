#!/usr/bin/env python3
"""
Bridge Controller — HTTP API for USB/WiFi mode switching.

Controls the CSI serial bridge process (start/stop) and provides
status information for the UI connection mode toggle.

Endpoints:
    GET  /api/bridge/status  → {"mode": "usb"|"wifi", "bridge_pid": int|null,
                                 "bridge_running": bool, "wifi_reachable": bool}
    POST /api/bridge/mode    → {"mode": "usb"|"wifi"}  → switches mode
    GET  /api/bridge/test    → {"reachable": [...], "unreachable": [...]}

Runs on port 3002 by default.
"""

import argparse
import json
import os
import signal
import socket
import subprocess
import sys
import time
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path

REPO_DIR = Path(__file__).resolve().parent.parent
BRIDGE_SCRIPT = REPO_DIR / "scripts" / "csi-serial-bridge.py"
BRIDGE_PID_FILE = REPO_DIR / "logs" / "csi-bridge.pid"
BRIDGE_LOG_FILE = REPO_DIR / "logs" / "csi-bridge.log"
FAULT_LOG_FILE = REPO_DIR / "logs" / "fault.log"
MODE_FILE = REPO_DIR / "logs" / "connection-mode"
SERVER_PID_FILE = REPO_DIR / "logs" / "sensing-server.pid"
SERVER_LOG_FILE = REPO_DIR / "logs" / "sensing-server.log"
SERVER_BINARY = REPO_DIR / "rust-port" / "wifi-densepose-rs" / "target" / "release" / "sensing-server"

# Discover venv python
VENV_PYTHON = REPO_DIR / ".venv" / "bin" / "python3"
if not VENV_PYTHON.exists():
    VENV_PYTHON = Path(sys.executable)


def _get_server_pid():
    """Read sensing server PID from file and check if alive."""
    if not SERVER_PID_FILE.exists():
        return None
    try:
        pid = int(SERVER_PID_FILE.read_text().strip())
        os.kill(pid, 0)
        return pid
    except (ValueError, ProcessLookupError, PermissionError):
        return None


def _stop_server():
    """Stop the sensing server process."""
    pid = _get_server_pid()
    if pid:
        try:
            os.kill(pid, signal.SIGTERM)
            # Wait up to 3s for graceful shutdown
            for _ in range(30):
                time.sleep(0.1)
                try:
                    os.kill(pid, 0)
                except ProcessLookupError:
                    break
            else:
                try:
                    os.kill(pid, signal.SIGKILL)
                except ProcessLookupError:
                    pass
        except ProcessLookupError:
            pass
    # Also pkill stray server processes
    subprocess.run(["pkill", "-f", "sensing-server"], capture_output=True)
    time.sleep(0.3)
    if SERVER_PID_FILE.exists():
        SERVER_PID_FILE.unlink(missing_ok=True)


def _start_server(source="esp32"):
    """Start the sensing server binary."""
    _stop_server()
    os.makedirs(REPO_DIR / "logs", exist_ok=True)

    if not SERVER_BINARY.exists():
        return None, "Server binary not found. Build with: cargo build -p wifi-densepose-sensing-server --release"

    # Read config
    env = os.environ.copy()
    http_port = env.get("HTTP_PORT", "3000")
    ws_port = env.get("WS_PORT", "3001")
    # Default bind: 0.0.0.0 so ESP32 nodes and LAN clients can reach it.
    # Override via SENSING_BIND_ADDR=127.0.0.1 for localhost-only dev.
    bind_addr = env.get("SENSING_BIND_ADDR", "0.0.0.0")
    ui_path = str(REPO_DIR / "ui")

    cmd = [
        str(SERVER_BINARY),
        "--source", source,
        "--bind-addr", bind_addr,
        "--http-port", http_port,
        "--ws-port", ws_port,
        "--udp-port", "5005",
        "--ui-path", ui_path,
    ]

    with open(SERVER_LOG_FILE, "a") as log_fh:
        proc = subprocess.Popen(
            cmd,
            stdout=log_fh,
            stderr=log_fh,
            start_new_session=True,
            cwd=str(REPO_DIR),
        )

    SERVER_PID_FILE.write_text(str(proc.pid))
    time.sleep(1)
    # Verify it started
    try:
        os.kill(proc.pid, 0)
        return proc.pid, None
    except ProcessLookupError:
        return None, "Server exited immediately. Check logs/sensing-server.log"


def _get_bridge_pid():
    """Read PID from file and check if process is alive."""
    if not BRIDGE_PID_FILE.exists():
        return None
    try:
        pid = int(BRIDGE_PID_FILE.read_text().strip())
        os.kill(pid, 0)  # Check if alive (signal 0)
        return pid
    except (ValueError, ProcessLookupError, PermissionError):
        return None


def _stop_bridge():
    """Stop the serial bridge process."""
    # Kill by PID file
    pid = _get_bridge_pid()
    if pid:
        try:
            os.kill(pid, signal.SIGTERM)
            time.sleep(0.5)
            # Force kill if still alive
            try:
                os.kill(pid, 0)
                os.kill(pid, signal.SIGKILL)
            except ProcessLookupError:
                pass
        except ProcessLookupError:
            pass

    # Also pkill any stray bridge processes
    subprocess.run(
        ["pkill", "-f", "csi-serial-bridge"],
        capture_output=True,
    )
    time.sleep(0.3)

    # Clean up PID file
    if BRIDGE_PID_FILE.exists():
        BRIDGE_PID_FILE.unlink(missing_ok=True)


def _start_bridge():
    """Start the serial bridge process."""
    _stop_bridge()  # Clean slate

    os.makedirs(REPO_DIR / "logs", exist_ok=True)

    with open(BRIDGE_LOG_FILE, "a") as log_fh:
        proc = subprocess.Popen(
            [
                str(VENV_PYTHON),
                str(BRIDGE_SCRIPT),
                "--fault-log",
                str(FAULT_LOG_FILE),
            ],
            stdout=log_fh,
            stderr=log_fh,
            start_new_session=True,
        )

    BRIDGE_PID_FILE.write_text(str(proc.pid))
    return proc.pid


def _test_wifi_connectivity():
    """Ping ESP32 nodes to check WiFi connectivity (AP isolation test)."""
    # Discover ESP32 IPs from ARP table
    result = subprocess.run(
        ["arp", "-a"],
        capture_output=True,
        text=True,
    )
    esp_ips = []
    for line in result.stdout.splitlines():
        lower = line.lower()
        if "esp32" in lower or "espressif" in lower:
            # Extract IP from arp output: hostname (IP) at MAC ...
            parts = line.split()
            for p in parts:
                if p.startswith("(") and p.endswith(")"):
                    esp_ips.append(p.strip("()"))

    reachable = []
    unreachable = []
    for ip in esp_ips:
        ret = subprocess.run(
            ["ping", "-c", "1", "-W", "1", ip],
            capture_output=True,
        )
        if ret.returncode == 0:
            reachable.append(ip)
        else:
            unreachable.append(ip)

    return reachable, unreachable


def _get_current_mode():
    """Read current mode from file."""
    if MODE_FILE.exists():
        mode = MODE_FILE.read_text().strip()
        if mode in ("usb", "wifi"):
            return mode
    # Default: if bridge is running → usb, else wifi
    return "usb" if _get_bridge_pid() else "wifi"


def _set_mode(mode):
    """Set connection mode."""
    os.makedirs(MODE_FILE.parent, exist_ok=True)
    MODE_FILE.write_text(mode)


# ── ESP32 USB auto-detect & auto-provision ────────────────────────────────────

PROVISION_SCRIPT = REPO_DIR / "firmware" / "esp32-csi-node" / "provision.py"
import re as _re


def _scan_usb_boards():
    """Enumerate connected serial ports and score them as likely ESP32 boards.

    Returns a list of dicts:
      {device, label, manufacturer, hwid, chip, score, recommended_env}
    Sorted descending by score.
    """
    try:
        from serial.tools import list_ports
    except ImportError:
        return {"error": "pyserial not installed in venv (.venv/bin/python -m pip install pyserial)"}

    boards = []
    for port in list_ports.comports():
        text = " ".join(filter(None, [port.device, port.description, port.manufacturer, port.hwid]))
        score = 0
        chip = "unknown"
        env = "esp32s3_n16r8"  # default 8 MB

        if _re.search(r"VID:PID=303[Aa]:|Espressif|USB JTAG", text):
            score += 100; chip = "ESP32-S3 (native USB-CDC)"
        if _re.search(r"VID:PID=10C4:|CP210|Silicon Labs", text):
            score += 60;  chip = "ESP32 via CP210x bridge"
        if _re.search(r"VID:PID=1A86:|CH340|wch", text):
            score += 40;  chip = "ESP32 via CH340 bridge"
        if _re.search(r"USB Serial|UART", text, _re.IGNORECASE):
            score += 15

        # Heuristic: SuperMini boards often report 4 MB flash via dmesg later;
        # we can't tell from VID/PID alone, so just flag it as a candidate.
        if "supermini" in text.lower():
            env = "esp32s3_supermini"

        boards.append({
            "device": port.device,
            "label": port.description or port.device,
            "manufacturer": port.manufacturer or "",
            "hwid": port.hwid or "",
            "chip": chip,
            "score": score,
            "recommended_env": env,
            "is_esp32": score >= 40,
        })

    boards.sort(key=lambda b: (b["score"], b["device"]), reverse=True)
    return boards


def _run_provision(port, ssid, password, target_ip, node_id,
                   no_firmware=True, extra_args=None, timeout=120):
    """Invoke firmware/esp32-csi-node/provision.py for a single board.

    Returns dict {ok, returncode, stdout, stderr, command}.
    """
    if not PROVISION_SCRIPT.exists():
        return {"ok": False, "error": f"provision.py not found at {PROVISION_SCRIPT}"}

    cmd = [str(VENV_PYTHON), str(PROVISION_SCRIPT),
           "--port", port,
           "--ssid", ssid,
           "--password", password,
           "--target-ip", target_ip,
           "--node-id", str(node_id)]
    if no_firmware:
        cmd.append("--no-firmware")
    if extra_args:
        cmd.extend(extra_args)

    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        return {
            "ok": proc.returncode == 0,
            "returncode": proc.returncode,
            "stdout": proc.stdout[-4000:],   # tail to keep response small
            "stderr": proc.stderr[-4000:],
            "command": " ".join(cmd[:6]) + " … --node-id " + str(node_id),
        }
    except subprocess.TimeoutExpired:
        return {"ok": False, "error": f"timeout after {timeout}s", "command": " ".join(cmd[:6])}
    except Exception as e:
        return {"ok": False, "error": str(e), "command": " ".join(cmd[:6])}


class BridgeHandler(BaseHTTPRequestHandler):
    """HTTP request handler for bridge control API."""

    def _send_json(self, data, status=200):
        body = json.dumps(data).encode()
        try:
            self.send_response(status)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            self.send_header("Access-Control-Allow-Headers", "Content-Type")
            self.end_headers()
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            # Client disconnected before we finished writing — harmless
            pass

    def log_message(self, format, *args):
        # Silence noisy per-request logs; errors still go to stderr
        return

    def do_OPTIONS(self):
        """Handle CORS preflight."""
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        if self.path == "/api/bridge/status":
            pid = _get_bridge_pid()
            mode = _get_current_mode()
            self._send_json({
                "mode": mode,
                "bridge_pid": pid,
                "bridge_running": pid is not None,
            })

        elif self.path == "/api/bridge/test":
            reachable, unreachable = _test_wifi_connectivity()
            self._send_json({
                "reachable": reachable,
                "unreachable": unreachable,
                "ap_isolation": len(unreachable) > 0 and len(reachable) == 0,
            })

        elif self.path == "/api/server/status":
            pid = _get_server_pid()
            self._send_json({
                "running": pid is not None,
                "pid": pid,
            })

        elif self.path == "/api/usb/scan":
            # Live USB serial port enumeration with ESP32 scoring.
            result = _scan_usb_boards()
            if isinstance(result, dict) and "error" in result:
                self._send_json(result, 500)
            else:
                self._send_json({"boards": result, "count": len(result)})

        else:
            self._send_json({"error": "Not found"}, 404)

    def do_POST(self):
        if self.path == "/api/bridge/mode":
            content_len = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_len) if content_len > 0 else b""

            try:
                data = json.loads(body) if body else {}
            except json.JSONDecodeError:
                self._send_json({"error": "Invalid JSON"}, 400)
                return

            mode = data.get("mode", "").lower()
            if mode not in ("usb", "wifi"):
                self._send_json(
                    {"error": "mode must be 'usb' or 'wifi'"}, 400
                )
                return

            if mode == "usb":
                pid = _start_bridge()
                _set_mode("usb")
                self._send_json({
                    "mode": "usb",
                    "bridge_pid": pid,
                    "bridge_running": True,
                    "message": "Serial bridge started — USB mode active",
                })

            elif mode == "wifi":
                _stop_bridge()
                _set_mode("wifi")
                # Quick connectivity test
                reachable, unreachable = _test_wifi_connectivity()
                ap_isolated = len(unreachable) > 0 and len(reachable) == 0

                self._send_json({
                    "mode": "wifi",
                    "bridge_pid": None,
                    "bridge_running": False,
                    "wifi_reachable": reachable,
                    "wifi_unreachable": unreachable,
                    "ap_isolation_detected": ap_isolated,
                    "message": (
                        "AP isolation detected — ESP32 nodes cannot reach the hub over WiFi. "
                        "Disable AP/client isolation in your router settings."
                        if ap_isolated
                        else "WiFi mode active — ESP32 nodes sending CSI directly via UDP"
                    ),
                })

        elif self.path == "/api/server/start":
            content_len = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_len) if content_len > 0 else b""
            try:
                data = json.loads(body) if body else {}
            except json.JSONDecodeError:
                data = {}
            source = data.get("source", "esp32")
            pid, err = _start_server(source)
            if err:
                self._send_json({"error": err, "running": False}, 500)
            else:
                self._send_json({"running": True, "pid": pid, "source": source})

        elif self.path == "/api/server/stop":
            _stop_server()
            self._send_json({"running": False, "message": "Server stopped"})

        elif self.path == "/api/server/restart":
            content_len = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_len) if content_len > 0 else b""
            try:
                data = json.loads(body) if body else {}
            except json.JSONDecodeError:
                data = {}
            source = data.get("source", "esp32")
            _stop_server()
            time.sleep(0.5)
            pid, err = _start_server(source)
            if err:
                self._send_json({"error": err, "running": False}, 500)
            else:
                self._send_json({"running": True, "pid": pid, "source": source, "message": "Server restarted"})

        elif self.path == "/api/provision/run":
            # Provision a single board with NVS-only flash.
            # Body: { port, ssid, password, target_ip, node_id, no_firmware? }
            content_len = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_len) if content_len > 0 else b""
            try:
                data = json.loads(body) if body else {}
            except json.JSONDecodeError:
                self._send_json({"error": "Invalid JSON"}, 400); return
            for f in ("port", "ssid", "password", "target_ip", "node_id"):
                if f not in data or data[f] in (None, ""):
                    self._send_json({"error": f"missing required field: {f}"}, 400); return
            result = _run_provision(
                port=data["port"], ssid=data["ssid"], password=data["password"],
                target_ip=data["target_ip"], node_id=int(data["node_id"]),
                no_firmware=bool(data.get("no_firmware", True)),
            )
            self._send_json(result, 200 if result.get("ok") else 500)

        elif self.path == "/api/provision/auto":
            # Auto-detect every connected ESP32 and provision sequentially with
            # auto-incrementing node IDs.
            # Body: { ssid, password, target_ip, start_node_id?, no_firmware? }
            content_len = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_len) if content_len > 0 else b""
            try:
                data = json.loads(body) if body else {}
            except json.JSONDecodeError:
                self._send_json({"error": "Invalid JSON"}, 400); return
            for f in ("ssid", "password", "target_ip"):
                if f not in data or data[f] in (None, ""):
                    self._send_json({"error": f"missing required field: {f}"}, 400); return
            start_id = int(data.get("start_node_id", 1))

            boards = _scan_usb_boards()
            if isinstance(boards, dict) and "error" in boards:
                self._send_json(boards, 500); return
            esp32s = [b for b in boards if b.get("is_esp32")]
            if not esp32s:
                self._send_json({"error": "No ESP32 boards detected on USB."}, 404); return

            results = []
            for i, b in enumerate(esp32s):
                node_id = start_id + i
                r = _run_provision(
                    port=b["device"], ssid=data["ssid"], password=data["password"],
                    target_ip=data["target_ip"], node_id=node_id,
                    no_firmware=bool(data.get("no_firmware", True)),
                )
                r["port"] = b["device"]; r["node_id"] = node_id; r["chip"] = b["chip"]
                results.append(r)

            ok_count = sum(1 for r in results if r.get("ok"))
            self._send_json({
                "total": len(results),
                "succeeded": ok_count,
                "failed": len(results) - ok_count,
                "results": results,
            }, 200 if ok_count == len(results) else 207)

        else:
            self._send_json({"error": "Not found"}, 404)

    def log_message(self, format, *args):
        """Suppress default HTTP logging."""
        pass


def main():
    parser = argparse.ArgumentParser(description="Bridge Controller API")
    parser.add_argument("--port", type=int, default=3002, help="HTTP port")
    parser.add_argument("--bind", default="0.0.0.0", help="Bind address")
    args = parser.parse_args()

    server = HTTPServer((args.bind, args.port), BridgeHandler)
    print(f"Bridge controller listening on {args.bind}:{args.port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down bridge controller")
        server.shutdown()


if __name__ == "__main__":
    main()
