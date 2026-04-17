#!/usr/bin/env python3
"""
CSI Serial Bridge — AP-isolation bypass
========================================
Reads ADR-018 CSI frames from ESP32 USB serial ports and forwards them
as UDP datagrams to the sensing-server on localhost:5005.

The ESP32 firmware wraps each ADR-018 frame with a 4-byte SLIP header:
  [0xAB][0xCD][len_hi][len_lo][...ADR-018 binary frame...]

This bridge strips the header and sends the raw ADR-018 frame to the
sensing-server as if it had arrived directly over WiFi UDP.

Usage:
    python scripts/csi-serial-bridge.py [--udp-port 5005]

Auto-detects /dev/ttyACM* and /dev/ttyUSB* ports at startup.
Reconnects automatically on disconnect.
Fault log written to logs/fault.log (JSON-lines).
"""

import json
import os
import socket
import serial
import serial.tools.list_ports
import threading
import time
import sys
import logging
import argparse

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("csi-bridge")

# SLIP header bytes (must match SLIP_SOF_0/1 in main.cpp)
SOF_0 = 0xAB
SOF_1 = 0xCD
BAUD  = 460800

# How long a port can be silent before it is flagged as faulted (seconds)
NODE_SILENCE_THRESHOLD = 30

stats = {}         # port -> {"rx": int, "fwd": int, "err": int}
last_frame = {}    # port -> float (epoch of most recent forwarded frame)
port_state = {}    # port -> "active" | "silent" | "disconnected"
_fault_log_path = None  # set in main()


def find_esp32_ports():
    """Return list of ttyACM* and ttyUSB* device paths."""
    found = []
    for p in serial.tools.list_ports.comports():
        if "ttyACM" in p.device or "ttyUSB" in p.device:
            found.append(p.device)
    found.sort()
    return found


def _write_fault(event: str, port: str, detail: str = ""):
    """Append a JSON-line fault record to the fault log."""
    if not _fault_log_path:
        return
    record = {
        "ts": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "event": event,
        "port": port,
        "detail": detail,
        "stats": stats.get(port, {}),
    }
    try:
        with open(_fault_log_path, "a") as fh:
            fh.write(json.dumps(record) + "\n")
    except OSError:
        pass


def bridge_port(port: str, udp_sock: socket.socket, dest: tuple):
    """Read SLIP-framed ADR-018 frames from serial and forward as UDP."""
    stats[port] = {"rx": 0, "fwd": 0, "err": 0}
    last_frame[port] = time.time()
    port_state[port] = "active"
    buf = bytearray()

    while True:
        try:
            s = serial.Serial(port, BAUD, timeout=1)
            log.info(f"[{port}] Connected at {BAUD} baud")
            port_state[port] = "active"
            _write_fault("port_connected", port)

            while True:
                chunk = s.read(512)
                if not chunk:
                    continue
                buf.extend(chunk)

                # Scan for complete SLIP frames
                while len(buf) >= 4:
                    # Find SOF marker
                    idx = -1
                    for i in range(len(buf) - 1):
                        if buf[i] == SOF_0 and buf[i + 1] == SOF_1:
                            idx = i
                            break

                    if idx < 0:
                        # No marker — keep last byte (may be partial SOF)
                        buf = buf[-1:]
                        break

                    if idx > 0:
                        buf = buf[idx:]  # discard garbage before SOF

                    if len(buf) < 4:
                        break  # wait for len bytes

                    frame_len = (buf[2] << 8) | buf[3]

                    if frame_len == 0 or frame_len > 1500:
                        buf = buf[2:]  # bad length — skip this SOF pair
                        stats[port]["err"] += 1
                        continue

                    if len(buf) < 4 + frame_len:
                        break  # incomplete frame — wait for more data

                    frame = bytes(buf[4: 4 + frame_len])
                    buf = buf[4 + frame_len:]

                    stats[port]["rx"] += 1

                    # Validate ADR-018 magic (0xC5110001 LE)
                    if len(frame) < 20:
                        stats[port]["err"] += 1
                        continue
                    magic = int.from_bytes(frame[:4], "little")
                    if magic != 0xC5110001:
                        stats[port]["err"] += 1
                        continue

                    # Forward as UDP to sensing-server
                    udp_sock.sendto(frame, dest)
                    stats[port]["fwd"] += 1
                    last_frame[port] = time.time()
                    if port_state[port] != "active":
                        log.info(f"[{port}] Node recovered — frames flowing again")
                        _write_fault("node_recovered", port)
                        port_state[port] = "active"

        except serial.SerialException as e:
            if port_state[port] != "disconnected":
                log.warning(f"[{port}] Serial error: {e} — reconnecting in 2s")
                _write_fault("port_disconnected", port, str(e))
                port_state[port] = "disconnected"
            time.sleep(2)
        except Exception as e:
            log.error(f"[{port}] Unexpected: {e}")
            _write_fault("port_error", port, str(e))
            time.sleep(2)


def stats_loop():
    while True:
        time.sleep(10)
        now = time.time()
        for port in sorted(stats.keys()):
            s = stats[port]
            silent_secs = now - last_frame.get(port, now)
            state = port_state.get(port, "unknown")
            log.info(f"[{port}] state={state} rx={s['rx']} fwd={s['fwd']} err={s['err']} "
                     f"silent={silent_secs:.0f}s")

            # Flag silence faults
            if (state == "active" and silent_secs > NODE_SILENCE_THRESHOLD):
                log.warning(f"[{port}] Node silent for {silent_secs:.0f}s — possible fault")
                _write_fault("node_silent", port,
                             f"No frames for {silent_secs:.0f}s")
                port_state[port] = "silent"
            elif state == "silent" and silent_secs <= NODE_SILENCE_THRESHOLD:
                log.info(f"[{port}] Node resumed after silence")
                _write_fault("node_resumed", port)
                port_state[port] = "active"


def main():
    parser = argparse.ArgumentParser(description="CSI Serial → UDP bridge")
    parser.add_argument("--udp-host", default="127.0.0.1")
    parser.add_argument("--udp-port", type=int, default=5005)
    parser.add_argument("--ports", nargs="*",
                        help="Serial ports (auto-detect if omitted)")
    parser.add_argument("--fault-log", default="",
                        help="Path to JSON-lines fault log (default: logs/fault.log next to script)")
    args = parser.parse_args()

    global _fault_log_path
    if args.fault_log:
        _fault_log_path = args.fault_log
    else:
        # Default: logs/fault.log relative to repo root (two levels up from scripts/)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        repo_root = os.path.dirname(script_dir)
        logs_dir = os.path.join(repo_root, "logs")
        os.makedirs(logs_dir, exist_ok=True)
        _fault_log_path = os.path.join(logs_dir, "fault.log")

    log.info(f"Fault log: {_fault_log_path}")

    dest = (args.udp_host, args.udp_port)

    udp_sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    log.info(f"Forwarding ADR-018 frames to udp://{dest[0]}:{dest[1]}")

    ports = args.ports if args.ports else find_esp32_ports()
    if not ports:
        log.error("No serial ports found. Plug in ESP32 boards or use --ports.")
        _write_fault("startup_no_ports", "", "No /dev/ttyACM* or /dev/ttyUSB* found")
        sys.exit(1)

    log.info(f"Bridging ports: {ports}")
    _write_fault("bridge_start", "", f"ports={ports} dest={dest[0]}:{dest[1]}")

    threads = []
    for port in ports:
        t = threading.Thread(target=bridge_port, args=(port, udp_sock, dest),
                             daemon=True, name=f"bridge-{port}")
        t.start()
        threads.append(t)

    stats_thread = threading.Thread(target=stats_loop, daemon=True)
    stats_thread.start()

    log.info("Bridge running. Ctrl-C to stop.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        _write_fault("bridge_stop", "", "KeyboardInterrupt")
        log.info("Stopped.")


if __name__ == "__main__":
    main()
