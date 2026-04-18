#!/usr/bin/env python3
"""Diagnose why a known-online ESP32 node is not streaming CSI to the hub.

Resolves a node_id from `firmware/esp32-csi-node/nodes.yaml`, finds its
current IP via the ARP table, then probes:

  1. ICMP ping  (is the device alive on the LAN?)
  2. TCP :8032  (is the OTA HTTP server up = is the firmware running?)
  3. UDP :5005 listener on the hub (have we ever seen this node's frames?)

Usage:
    python scripts/diagnose-node.py --node-id 2
    python scripts/diagnose-node.py --node-id 2 --hub http://localhost:3000

The script does NOT change anything on the device. It exits 0 on a
clean diagnosis, non-zero otherwise.

NOTE: There is no network re-provisioning endpoint in the current
firmware; if a node is alive on Wi-Fi but not streaming, the only
remediation today is a USB re-flash via `provision.py`. This tool just
helps confirm the situation before you grab a cable.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import socket
import subprocess
import sys
import urllib.request
from pathlib import Path
from typing import Optional


def run(cmd: list[str], timeout: int = 5) -> tuple[int, str]:
    try:
        out = subprocess.run(
            cmd, capture_output=True, text=True, timeout=timeout
        )
        return out.returncode, (out.stdout + out.stderr).strip()
    except (subprocess.TimeoutExpired, FileNotFoundError):
        return 1, ""


def load_registry(repo_root: Path) -> dict:
    """Tiny YAML reader — enough for our flat schema, no extra dep."""
    yaml_path = repo_root / "firmware" / "esp32-csi-node" / "nodes.yaml"
    if not yaml_path.exists():
        return {}
    nodes: dict = {}
    current: Optional[dict] = None
    for raw in yaml_path.read_text().splitlines():
        line = raw.rstrip()
        if not line or line.lstrip().startswith("#"):
            continue
        m = re.match(r"^  - node_id:\s*(\d+)\s*$", line)
        if m:
            current = {"node_id": int(m.group(1))}
            nodes[current["node_id"]] = current
            continue
        if current is None:
            continue
        m = re.match(r"^    (\w+):\s*(.+?)\s*$", line)
        if m:
            value = m.group(2)
            # Strip inline `# comment` and surrounding quotes.
            value = re.sub(r"\s+#.*$", "", value).strip().strip('"')
            current[m.group(1)] = value
    return nodes


def find_ip_for_mac(mac: str) -> Optional[str]:
    if not mac or not shutil.which("ip"):
        return None
    rc, out = run(["ip", "neigh"])
    if rc != 0:
        return None
    mac = mac.lower()
    for line in out.splitlines():
        if mac in line.lower():
            return line.split()[0]
    return None


def ping(ip: str) -> bool:
    rc, _ = run(["ping", "-c", "2", "-W", "1", ip], timeout=6)
    return rc == 0


def probe_tcp(ip: str, port: int, timeout: float = 2.0) -> bool:
    try:
        with socket.create_connection((ip, port), timeout=timeout):
            return True
    except OSError:
        return False


def hub_streaming_nodes(hub_url: str) -> list[int]:
    try:
        with urllib.request.urlopen(f"{hub_url}/api/v1/nodes", timeout=3) as r:
            data = json.loads(r.read().decode())
        return sorted({int(n["node_id"]) for n in data.get("nodes", [])})
    except Exception:
        return []


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--node-id", type=int,
                   help="Diagnose a single node ID. Omit to audit ALL registered nodes.")
    p.add_argument("--hub", default="http://localhost:3000",
                   help="Sensing-server base URL (default: http://localhost:3000)")
    p.add_argument("--repo-root", default=str(Path(__file__).resolve().parent.parent))
    args = p.parse_args()

    repo_root = Path(args.repo_root)
    registry = load_registry(repo_root)
    if not registry:
        print("nodes.yaml empty or unreadable")
        return 2

    streaming = hub_streaming_nodes(args.hub)
    if args.node_id is None:
        return audit_all(registry, streaming, args.hub)
    record = registry.get(args.node_id)
    if not record:
        print(f"node {args.node_id}: NOT FOUND in nodes.yaml")
        return 2
    return diagnose_one(args.node_id, record, streaming, args.hub)


def diagnose_one(node_id: int, record: dict, streaming: list[int], hub_url: str) -> int:
    mac = record.get("mac", "")
    print(f"node {node_id}: mac={mac or '(unset)'} "
          f"board={record.get('board', '?')!r} "
          f"role={record.get('led_role', '?')}")

    ip = find_ip_for_mac(mac) if mac else None
    if not ip:
        print(f"  arp:    no entry for MAC {mac!r} \u2014 device may be off-net "
              f"or hasn't been talked to recently")
    else:
        print(f"  arp:    {ip}")

    is_streaming = node_id in streaming
    print(f"  hub:    streaming nodes = {streaming or '(hub unreachable)'}")
    print(f"  hub:    node {node_id} streaming = {is_streaming}")

    if ip:
        print(f"  ping:   {'OK' if ping(ip) else 'FAIL'}")
        print(f"  :8032:  {'open (OTA server up = firmware running)' if probe_tcp(ip, 8032) else 'refused/filtered'}")

    print()
    if is_streaming:
        print("VERDICT: node is healthy and streaming. Nothing to do.")
        return 0
    if ip and probe_tcp(ip, 8032):
        print("VERDICT: firmware is running but not streaming to this hub.")
        print("         Most likely the node's NVS `target_ip` points at a")
        print("         different host. Re-provision over USB:")
        print(f"           python firmware/esp32-csi-node/provision.py "
              f"--no-firmware --target-ip <THIS_HUB_IP> --node-id {node_id}")
        return 1
    if ip:
        print("VERDICT: device is on Wi-Fi but firmware appears down.")
        print("         No network re-provisioning endpoint exists; reflash via USB:")
        print(f"           python firmware/esp32-csi-node/provision.py "
              f"--ssid <SSID> --password <PASS> --target-ip <HUB_IP> --node-id {node_id}")
        return 1
    print("VERDICT: device not visible on the LAN. Plug it in over USB and run:")
    print(f"           python firmware/esp32-csi-node/provision.py "
          f"--ssid <SSID> --password <PASS> --target-ip <HUB_IP> --node-id {node_id}")
    return 1


def audit_all(registry: dict, streaming: list[int], hub_url: str) -> int:
    """Walk every registered node and bucket each by current state.

    Buckets:
      OK              \u2014 streaming to the hub
      ON_WIFI_NO_CSI  \u2014 ARP entry + ping OK, but not streaming  (\u2190 the user's case)
      OTA_RESPONDS    \u2014 :8032 reachable but no CSI \u2192 wrong target_ip in NVS
      OFFLINE         \u2014 no ARP entry / not on the LAN
      UNREGISTERED    \u2014 MAC slot empty in nodes.yaml
    """
    rows: list[tuple[int, str, str, str, str]] = []
    on_wifi_no_csi: list[int] = []
    ota_wrong_target: list[int] = []

    for node_id in sorted(registry):
        rec = registry[node_id]
        mac = rec.get("mac", "")
        if not mac:
            rows.append((node_id, "?", "(empty)", "UNREGISTERED",
                         "fill `mac:` slot in nodes.yaml when you flash this ID"))
            continue
        ip = find_ip_for_mac(mac)
        if node_id in streaming:
            rows.append((node_id, ip or "?", mac, "OK", "streaming"))
            continue
        if not ip:
            rows.append((node_id, "-", mac, "OFFLINE", "not on LAN \u2014 power off or different subnet"))
            continue
        # On the LAN but not streaming. Distinguish "firmware down" from
        # "firmware up but pointing at the wrong hub".
        if probe_tcp(ip, 8032):
            ota_wrong_target.append(node_id)
            rows.append((node_id, ip, mac, "OTA_RESPONDS",
                         "firmware up, target_ip != this hub \u2014 reprovision over USB with --no-firmware"))
        else:
            on_wifi_no_csi.append(node_id)
            rows.append((node_id, ip, mac, "ON_WIFI_NO_CSI",
                         "firmware crashed/disabled \u2014 needs USB reflash"))

    # Pretty table.
    w_id, w_ip, w_mac, w_st = 4, 16, 19, 14
    print(f"{'id':>{w_id}}  {'ip':<{w_ip}}  {'mac':<{w_mac}}  {'state':<{w_st}}  remediation")
    print("-" * 110)
    for nid, ip, mac, st, hint in rows:
        print(f"{nid:>{w_id}}  {ip:<{w_ip}}  {mac:<{w_mac}}  {st:<{w_st}}  {hint}")

    print()
    print(f"hub streaming nodes: {streaming or '(hub unreachable)'}")
    if on_wifi_no_csi:
        print(f"\u26a0  on Wi-Fi but firmware down \u2192 needs USB reflash: {on_wifi_no_csi}")
    if ota_wrong_target:
        print(f"\u26a0  firmware up but wrong hub  \u2192 USB reprovision (no flash): {ota_wrong_target}")
    if not on_wifi_no_csi and not ota_wrong_target:
        print("\u2713  every Wi-Fi-visible node is either streaming or fully offline.")
    return 0 if not (on_wifi_no_csi or ota_wrong_target) else 1


if __name__ == "__main__":
    sys.exit(main())
