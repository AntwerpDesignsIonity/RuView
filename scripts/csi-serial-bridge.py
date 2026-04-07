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
"""

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

stats = {}  # port -> {"rx": int, "fwd": int}


def find_esp32_ports():
    """Return list of ttyACM* and ttyUSB* device paths."""
    found = []
    for p in serial.tools.list_ports.comports():
        if "ttyACM" in p.device or "ttyUSB" in p.device:
            found.append(p.device)
    found.sort()
    return found


def bridge_port(port: str, udp_sock: socket.socket, dest: tuple):
    """Read SLIP-framed ADR-018 frames from serial and forward as UDP."""
    stats[port] = {"rx": 0, "fwd": 0, "err": 0}
    buf = bytearray()

    while True:
        try:
            s = serial.Serial(port, BAUD, timeout=1)
            log.info(f"[{port}] Connected at {BAUD} baud")

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

        except serial.SerialException as e:
            log.warning(f"[{port}] Serial error: {e} — reconnecting in 2s")
            time.sleep(2)
        except Exception as e:
            log.error(f"[{port}] Unexpected: {e}")
            time.sleep(2)


def stats_loop():
    while True:
        time.sleep(10)
        for port, s in sorted(stats.items()):
            log.info(f"[{port}] rx={s['rx']} fwd={s['fwd']} err={s['err']}")


def main():
    parser = argparse.ArgumentParser(description="CSI Serial → UDP bridge")
    parser.add_argument("--udp-host", default="127.0.0.1")
    parser.add_argument("--udp-port", type=int, default=5005)
    parser.add_argument("--ports", nargs="*",
                        help="Serial ports (auto-detect if omitted)")
    args = parser.parse_args()

    dest = (args.udp_host, args.udp_port)

    udp_sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    log.info(f"Forwarding ADR-018 frames to udp://{dest[0]}:{dest[1]}")

    ports = args.ports if args.ports else find_esp32_ports()
    if not ports:
        log.error("No serial ports found. Plug in ESP32 boards or use --ports.")
        sys.exit(1)

    log.info(f"Bridging ports: {ports}")

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
        log.info("Stopped.")


if __name__ == "__main__":
    main()
