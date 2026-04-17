#!/usr/bin/env python3
"""
AEDI-S cross-platform launcher — Ionity Global (Pty) Ltd.

Stdlib-only Python ≥ 3.9. Works on Linux / macOS / Windows.

Subcommands:
    start [--source esp32|wifi|simulate] [--bind 0.0.0.0] [--http 3000]
    stop
    status
    logs [--tail N]
    build
    doctor
    provision  (ESP32 WiFi credentials)

Exit codes:
    0 = ok, 1 = failure, 2 = unhealthy
"""
from __future__ import annotations

import argparse
import json
import os
import platform
import shutil
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent
LOGS = ROOT / "logs"
LOGS.mkdir(exist_ok=True)

IS_WIN = platform.system() == "Windows"
EXE = ".exe" if IS_WIN else ""

SERVER_BIN = ROOT / "rust-port" / "wifi-densepose-rs" / "target" / "release" / f"sensing-server{EXE}"
SERVER_LOG = LOGS / "sensing-server.log"
SERVER_PID = LOGS / "sensing-server.pid"
UI_PATH = ROOT / "ui"

DEFAULT_HTTP = 3000
DEFAULT_WS = 3001
DEFAULT_UDP = 5005

# ── colour helpers ─────────────────────────────────────────────────────────
def _supports_color() -> bool:
    if IS_WIN:
        # Win10+ terminals support ANSI; fall back on legacy
        return os.environ.get("WT_SESSION") or os.environ.get("TERM_PROGRAM")
    return sys.stdout.isatty()

C = {"g": "\033[32m", "r": "\033[31m", "y": "\033[33m", "b": "\033[36m", "d": "\033[2m", "0": "\033[0m"} \
    if _supports_color() else {k: "" for k in "gryb d0"}

def _pr(color: str, *msg):
    print(C.get(color, "") + " ".join(str(m) for m in msg) + C["0"])

def ok(*m): _pr("g", "✓", *m)
def warn(*m): _pr("y", "!", *m)
def err(*m): _pr("r", "✗", *m)
def info(*m): _pr("b", "›", *m)

# ── process helpers (portable) ─────────────────────────────────────────────
def _pid_alive(pid: int) -> bool:
    if IS_WIN:
        try:
            out = subprocess.check_output(["tasklist", "/FI", f"PID eq {pid}"], stderr=subprocess.DEVNULL, text=True)
            return str(pid) in out
        except Exception:
            return False
    try:
        os.kill(pid, 0)
        return True
    except (ProcessLookupError, PermissionError):
        return False
    except OSError:
        return False

def _read_pid() -> int | None:
    try:
        pid = int(SERVER_PID.read_text().strip())
        return pid if _pid_alive(pid) else None
    except Exception:
        return None

def _kill(pid: int) -> None:
    if IS_WIN:
        subprocess.run(["taskkill", "/F", "/PID", str(pid)], check=False, capture_output=True)
    else:
        subprocess.run(["kill", str(pid)], check=False, capture_output=True)

def _port_free(port: int, host: str = "127.0.0.1") -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.2)
        return s.connect_ex((host, port)) != 0

def _http_json(url: str, timeout: float = 3.0) -> dict | None:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "aedi-launch"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8", "replace"))
    except (urllib.error.URLError, TimeoutError, OSError, json.JSONDecodeError):
        return None

# ── subcommands ────────────────────────────────────────────────────────────
def cmd_start(args) -> int:
    if not SERVER_BIN.exists():
        err(f"Binary missing: {SERVER_BIN}")
        info("Run: python launch.py build")
        return 1
    existing = _read_pid()
    if existing:
        warn(f"Already running as PID {existing}. Use `stop` first.")
        return 0
    if not _port_free(args.http):
        err(f"Port {args.http} busy — another process is listening.")
        return 1
    cmd = [
        str(SERVER_BIN),
        "--source", args.source,
        "--bind-addr", args.bind,
        "--http-port", str(args.http),
        "--ws-port", str(args.ws),
        "--udp-port", str(args.udp),
        "--ui-path", str(UI_PATH),
    ]
    info("Launching:", " ".join(cmd))
    log_fh = open(SERVER_LOG, "a", buffering=1, encoding="utf-8")
    kw = dict(stdout=log_fh, stderr=log_fh, cwd=str(ROOT))
    if IS_WIN:
        DETACHED = 0x00000008
        kw["creationflags"] = DETACHED  # type: ignore
    else:
        kw["start_new_session"] = True
    proc = subprocess.Popen(cmd, **kw)
    SERVER_PID.write_text(str(proc.pid))
    # probe health up to 6 s
    for _ in range(30):
        time.sleep(0.2)
        h = _http_json(f"http://127.0.0.1:{args.http}/health")
        if h and h.get("status") == "ok":
            ok(f"Server up. PID {proc.pid}  source={h.get('source')}  tick={h.get('tick')}")
            info(f"UI:     http://localhost:{args.http}/ui/index.html")
            info(f"Health: http://localhost:{args.http}/health")
            info(f"WS:     ws://localhost:{args.ws}/ws/sensing")
            return 0
    err("Server didn't respond to /health in 6s — check logs.")
    return 2

def cmd_stop(_args) -> int:
    pid = _read_pid()
    if not pid:
        warn("Not running (no live PID file).")
        SERVER_PID.unlink(missing_ok=True)
        return 0
    _kill(pid)
    for _ in range(20):
        if not _pid_alive(pid):
            break
        time.sleep(0.1)
    SERVER_PID.unlink(missing_ok=True)
    ok(f"Stopped PID {pid}")
    return 0

def cmd_status(args) -> int:
    pid = _read_pid()
    if not pid:
        warn("Server not running.")
        return 2
    h = _http_json(f"http://127.0.0.1:{args.http}/health")
    if not h:
        err(f"PID {pid} alive but /health unreachable.")
        return 2
    ok(f"Running PID {pid}  status={h.get('status')}  source={h.get('source')}  tick={h.get('tick')}  clients={h.get('clients')}")
    latest = _http_json(f"http://127.0.0.1:{args.http}/api/v1/sensing/latest")
    if latest:
        cls = latest.get("classification", {}) or {}
        nodes = latest.get("nodes", []) or []
        print(f"  presence      = {cls.get('presence')}")
        print(f"  motion        = {cls.get('motion_level')}")
        print(f"  est_persons   = {latest.get('estimated_persons')}")
        print(f"  active_nodes  = {len(nodes)}")
        for n in nodes[:5]:
            print(f"    node {n.get('node_id')}: rssi={n.get('rssi_dbm'):>5} dBm  sub={n.get('subcarrier_count')}")
    return 0

def cmd_logs(args) -> int:
    if not SERVER_LOG.exists():
        warn("No log file yet.")
        return 0
    try:
        with open(SERVER_LOG, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
        for l in lines[-args.tail:]:
            sys.stdout.write(l)
        return 0
    except Exception as e:
        err(f"Log read failed: {e}")
        return 1

def cmd_build(_args) -> int:
    cargo = shutil.which("cargo")
    if not cargo:
        err("cargo not found. Install Rust: https://rustup.rs")
        return 1
    cwd = ROOT / "rust-port" / "wifi-densepose-rs"
    info(f"cargo build -p wifi-densepose-sensing-server --release (in {cwd})")
    r = subprocess.run(
        [cargo, "build", "-p", "wifi-densepose-sensing-server", "--release", "--no-default-features"],
        cwd=str(cwd),
    )
    if r.returncode == 0:
        ok("Build succeeded.")
        return 0
    err("Build failed.")
    return 1

def cmd_doctor(args) -> int:
    findings = []
    # binary
    findings.append(("binary", SERVER_BIN.exists(), f"{SERVER_BIN.name}  ({SERVER_BIN})"))
    # cargo
    findings.append(("cargo", shutil.which("cargo") is not None, "toolchain"))
    # python venv
    findings.append(("venv", (ROOT / ".venv").exists() or (ROOT / ".venv" / "Scripts").exists(), ".venv/ present"))
    # ui
    findings.append(("ui", (UI_PATH / "index.html").exists(), "ui/index.html"))
    # running?
    pid = _read_pid()
    findings.append(("running", pid is not None, f"PID {pid}" if pid else "no server"))
    # health
    h = _http_json(f"http://127.0.0.1:{args.http}/health") if pid else None
    findings.append(("health", bool(h), str(h) if h else "n/a"))
    # ports — TCP probe only for TCP services; UDP can't be tested this way.
    for p in (args.http, DEFAULT_WS):
        findings.append((f"tcp:{p}", not _port_free(p) if pid else True, "listening" if pid else "skip"))
    findings.append((f"udp:{DEFAULT_UDP}", True, "CSI ingress (UDP — cannot probe)"))
    bad = 0
    for key, ok_, msg in findings:
        (ok if ok_ else warn)(f"{key:12s} {msg}")
        bad += 0 if ok_ else 1
    return 0 if bad == 0 else 1

def cmd_provision(args) -> int:
    script = ROOT / "firmware" / "esp32-csi-node" / "provision.py"
    if not script.exists():
        err(f"Provisioner missing: {script}")
        return 1
    cmd = [sys.executable, str(script), "--ssid", args.ssid, "--password", args.password]
    if args.port:
        cmd += ["--port", args.port]
    if args.target_ip:
        cmd += ["--target-ip", args.target_ip]
    return subprocess.run(cmd, cwd=str(ROOT)).returncode

# ── argparse wiring ────────────────────────────────────────────────────────
def main() -> int:
    p = argparse.ArgumentParser(prog="launch.py", description="AEDI-S cross-platform launcher")
    p.add_argument("--http", type=int, default=int(os.environ.get("HTTP_PORT", DEFAULT_HTTP)))
    p.add_argument("--ws", type=int, default=int(os.environ.get("WS_PORT", DEFAULT_WS)))
    p.add_argument("--udp", type=int, default=int(os.environ.get("UDP_PORT", DEFAULT_UDP)))
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("start", help="Start sensing-server")
    s.add_argument("--source", default=os.environ.get("IONITY_SOURCE", "esp32"),
                   choices=["esp32", "wifi", "simulate"])
    s.add_argument("--bind", default=os.environ.get("SENSING_BIND_ADDR", "0.0.0.0"))
    s.set_defaults(fn=cmd_start)

    sub.add_parser("stop", help="Stop sensing-server").set_defaults(fn=cmd_stop)
    sub.add_parser("status", help="Show health + latest frame").set_defaults(fn=cmd_status)

    l = sub.add_parser("logs", help="Print tail of server log")
    l.add_argument("--tail", type=int, default=80)
    l.set_defaults(fn=cmd_logs)

    sub.add_parser("build", help="cargo build --release").set_defaults(fn=cmd_build)
    sub.add_parser("doctor", help="Environment/health diagnostic").set_defaults(fn=cmd_doctor)

    pr = sub.add_parser("provision", help="Provision ESP32 WiFi credentials")
    pr.add_argument("--ssid", required=True)
    pr.add_argument("--password", required=True)
    pr.add_argument("--port", help="Serial port, e.g. /dev/ttyUSB0 or COM7")
    pr.add_argument("--target-ip", help="Hub LAN IP for UDP CSI")
    pr.set_defaults(fn=cmd_provision)

    args = p.parse_args()
    try:
        return args.fn(args)
    except KeyboardInterrupt:
        warn("Interrupted.")
        return 1

if __name__ == "__main__":
    sys.exit(main())
