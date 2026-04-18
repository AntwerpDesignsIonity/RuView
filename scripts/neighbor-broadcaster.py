#!/usr/bin/env python3
"""
neighbor-broadcaster.py — Subscribes to the AEDI-S WebSocket, extracts node
roster, UDP-broadcasts compact packets to 255.255.255.255:5006 so the node 4
LCD can show neighbour status without every node broadcasting on its own.

Packet format (small enough for one ESP32 UDP read):
    [magic:4='RSTR'][count:u8]([node_id:u8][rssi:i8])*count
    Max 16 nodes -> 4 + 1 + 16*2 = 37 bytes.

Run:
    python3 scripts/neighbor-broadcaster.py &
"""
import asyncio
import json
import socket
import struct
import sys
import os
import base64

HUB_HOST = os.environ.get("AEDI_HUB", "127.0.0.1")
WS_PORT  = int(os.environ.get("AEDI_WS_PORT", "3001"))
BCAST_IP = os.environ.get("AEDI_BCAST_IP", "255.255.255.255")
BCAST_PORT = int(os.environ.get("AEDI_BCAST_PORT", "5006"))
INTERVAL_S = float(os.environ.get("AEDI_BCAST_INTERVAL", "2.0"))
MAGIC = b"RSTR"


def build_packet(nodes: list[dict]) -> bytes:
    count = min(len(nodes), 16)
    buf = bytearray(MAGIC + bytes([count]))
    for n in nodes[:count]:
        nid = int(n.get("node_id", 0)) & 0xFF
        rssi = int(n.get("rssi_dbm", -100))
        if rssi < -128: rssi = -128
        if rssi > 127:  rssi = 127
        buf.extend(struct.pack("<bb", nid, rssi))
    return bytes(buf)


async def ws_connect_and_stream(sock: socket.socket):
    import asyncio
    reader, writer = await asyncio.open_connection(HUB_HOST, WS_PORT)
    key = base64.b64encode(os.urandom(16)).decode()
    req = (
        f"GET /ws/sensing HTTP/1.1\r\n"
        f"Host: {HUB_HOST}:{WS_PORT}\r\n"
        f"Upgrade: websocket\r\n"
        f"Connection: Upgrade\r\n"
        f"Sec-WebSocket-Key: {key}\r\n"
        f"Sec-WebSocket-Version: 13\r\n\r\n"
    )
    writer.write(req.encode())
    await writer.drain()
    # Read handshake response up to end-of-headers
    hdr = b""
    while b"\r\n\r\n" not in hdr:
        chunk = await reader.read(4096)
        if not chunk:
            raise RuntimeError("WS handshake closed")
        hdr += chunk
    if b"101" not in hdr[:32]:
        raise RuntimeError(f"WS handshake failed: {hdr[:80]!r}")

    last_bcast = 0.0
    latest_nodes: list[dict] = []

    async def broadcaster():
        nonlocal last_bcast
        while True:
            await asyncio.sleep(INTERVAL_S)
            pkt = build_packet(latest_nodes)
            try:
                sock.sendto(pkt, (BCAST_IP, BCAST_PORT))
                last_bcast = asyncio.get_event_loop().time()
                ids = [int(n.get("node_id", 0)) for n in latest_nodes[:16]]
                print(f"[roster] broadcast {len(ids)} nodes: {ids} ({len(pkt)} B)", flush=True)
            except Exception as e:
                print(f"[roster] broadcast error: {e}", file=sys.stderr)

    async def ws_reader():
        while True:
            b = await reader.readexactly(2)
            ln = b[1] & 0x7F
            if ln == 126:
                ln = struct.unpack(">H", await reader.readexactly(2))[0]
            elif ln == 127:
                ln = struct.unpack(">Q", await reader.readexactly(8))[0]
            payload = await reader.readexactly(ln) if ln else b""
            try:
                d = json.loads(payload)
            except Exception:
                continue
            nodes = d.get("nodes", [])
            if isinstance(nodes, list):
                latest_nodes.clear()
                latest_nodes.extend(nodes)

    await asyncio.gather(broadcaster(), ws_reader())


async def main():
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
    try:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    except Exception:
        pass
    print(f"[roster] broadcasting roster to {BCAST_IP}:{BCAST_PORT} every {INTERVAL_S}s", flush=True)

    while True:
        try:
            await ws_connect_and_stream(sock)
        except Exception as e:
            print(f"[roster] ws error ({e}); retrying in 3s", file=sys.stderr)
            await asyncio.sleep(3)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        pass
