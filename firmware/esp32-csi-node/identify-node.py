#!/usr/bin/env python3
"""identify-node.py — detect a connected ESP32 and look it up in nodes.yaml.

Usage:
  python firmware/esp32-csi-node/identify-node.py                 # all ports
  python firmware/esp32-csi-node/identify-node.py --port /dev/ttyACM0
  python firmware/esp32-csi-node/identify-node.py --provision     # auto re-provision known nodes
  python firmware/esp32-csi-node/identify-node.py --flash         # auto flash + provision
  python firmware/esp32-csi-node/identify-node.py --provision --node-id 3   # claim slot 3 for unknown MAC

Reads the wifi MAC via esptool, looks it up in nodes.yaml, prints the assigned
node_id, board, pio_env and LED role. With --provision it sources WIFI_SSID /
WIFI_PASS / HUB_IP / HUB_PORT from .env.local and runs provision.py with the
correct node_id. With --flash it also runs `pio run -e <env> -t upload` first.
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
NODES_YAML = Path(__file__).parent / "nodes.yaml"
PROVISION = Path(__file__).parent / "provision.py"
ENV_FILE = ROOT / ".env.local"
PIO = Path.home() / ".platformio" / "penv" / "bin" / "pio"
ESPTOOL = Path.home() / ".platformio" / "packages" / "tool-esptoolpy" / "esptool.py"
PYTHON = Path.home() / ".platformio" / "penv" / "bin" / "python"
IONITY = ROOT / "ionity"


def load_nodes() -> list[dict]:
    """Load node registry from nodes.yaml. Falls back to a minimal regex
    parser if PyYAML isn't installed (e.g. running on a fresh system before
    `pip install -r requirements.txt`)."""
    try:
        import yaml  # type: ignore
    except ImportError:
        return _load_nodes_fallback()
    data = yaml.safe_load(NODES_YAML.read_text()) or {}
    nodes = data.get("nodes", []) or []
    # Normalise: ensure every node has at least node_id; coerce mac to lower str.
    out: list[dict] = []
    for n in nodes:
        if not isinstance(n, dict) or "node_id" not in n:
            continue
        n["mac"] = (n.get("mac") or "").lower()
        out.append(n)
    return out


def _load_nodes_fallback() -> list[dict]:
    """Tiny YAML reader for the simple schema in nodes.yaml. Used only when
    PyYAML is not available."""
    nodes: list[dict] = []
    cur: dict | None = None
    for raw in NODES_YAML.read_text().splitlines():
        line = raw.split("#", 1)[0].rstrip()
        if not line.strip():
            continue
        if line.startswith("nodes:"):
            continue
        if line.lstrip().startswith("- "):
            if cur:
                nodes.append(cur)
            cur = {}
            line = line.lstrip()[2:]
        if cur is None:
            continue
        m = re.match(r"\s*([a-zA-Z_]+):\s*(.*)$", line)
        if not m:
            continue
        key, val = m.group(1), m.group(2).strip()
        if val.startswith('"') and val.endswith('"'):
            val = val[1:-1]
        if key == "node_id":
            try:
                val = int(val)
            except ValueError:
                pass
        cur[key] = val
    if cur:
        nodes.append(cur)
    return nodes


def list_ports() -> list[str]:
    return sorted(
        str(p)
        for p in Path("/dev").glob("ttyACM*")
    ) + sorted(str(p) for p in Path("/dev").glob("ttyUSB*"))


def read_mac(port: str) -> str | None:
    if not ESPTOOL.exists():
        print(f"!! esptool not found at {ESPTOOL}", file=sys.stderr)
        return None
    try:
        out = subprocess.run(
            [str(PYTHON), str(ESPTOOL), "--port", port, "read_mac"],
            capture_output=True, text=True, timeout=25,
        ).stdout
    except subprocess.TimeoutExpired:
        return None
    for line in out.splitlines():
        m = re.search(r"MAC:\s*([0-9a-f:]{17})", line, re.I)
        if m:
            return m.group(1).lower()
    return None


def load_env() -> dict[str, str]:
    env: dict[str, str] = {}
    if not ENV_FILE.exists():
        return env
    for line in ENV_FILE.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        env[k.strip()] = v.strip().strip('"').strip("'")
    return env


def do_flash(node: dict, port: str) -> bool:
    env = node.get("pio_env")
    if not env:
        print("!! node has no pio_env, skipping flash")
        return False
    if not PIO.exists():
        print(f"!! pio not found at {PIO}", file=sys.stderr)
        return False
    print(f">> Flashing {env} to {port} ...")
    rc = subprocess.run(
        [str(PIO), "run", "-e", env, "-t", "upload", "--upload-port", port],
        cwd=IONITY,
    ).returncode
    return rc == 0


def do_provision(node: dict, port: str) -> bool:
    env = load_env()
    ssid = env.get("WIFI_SSID")
    pwd = env.get("WIFI_PASS")
    hub = env.get("HUB_IP")
    hub_port = env.get("HUB_PORT", "5005")
    if not (ssid and pwd and hub):
        print("!! Missing WIFI_SSID/WIFI_PASS/HUB_IP in .env.local", file=sys.stderr)
        return False
    args = [
        str(PYTHON), str(PROVISION),
        "--port", port, "--no-firmware",
        "--ssid", ssid, "--password", pwd,
        "--target-ip", hub, "--target-port", hub_port,
        "--node-id", str(node["node_id"]),
        "--led-hub", node.get("led_role", "edge"),
    ]
    print(f">> Provisioning node-id={node['node_id']} via {port} ...")
    return subprocess.run(args).returncode == 0


def find_node(nodes: list[dict], mac: str) -> dict | None:
    mac = mac.lower()
    for n in nodes:
        if (n.get("mac") or "").lower() == mac:
            return n
    return None


def export_json(nodes: list[dict], out: Path) -> None:
    import json
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps({"nodes": nodes}, indent=2))


def main() -> int:
    ap = argparse.ArgumentParser(description="Detect and identify connected AEDI-S ESP32 nodes")
    ap.add_argument("--port", help="serial port to probe (default: scan all ttyACM*/ttyUSB*)")
    ap.add_argument("--provision", action="store_true", help="re-provision NVS for matched nodes")
    ap.add_argument("--flash", action="store_true", help="flash firmware then provision (implies --provision)")
    ap.add_argument("--node-id", type=int, help="if MAC unknown, claim this node_id from nodes.yaml")
    ap.add_argument("--write-json", metavar="PATH", help="export the registry as JSON to PATH and exit (no probing)")
    ap.add_argument("--auto", action="store_true", help="auto re-provision any KNOWN node that is currently in download mode (no flash, idempotent)")
    args = ap.parse_args()

    if not NODES_YAML.exists():
        print(f"!! Registry missing: {NODES_YAML}", file=sys.stderr)
        return 1
    nodes = load_nodes()

    if args.write_json:
        export_json(nodes, Path(args.write_json))
        print(f">> Wrote registry JSON to {args.write_json}")
        return 0

    if args.auto:
        # Auto mode = provision known nodes silently, never flash, never prompt.
        args.provision = True
    ports = [args.port] if args.port else list_ports()
    if not ports:
        print("!! No ttyACM*/ttyUSB* ports found.")
        return 1

    flash = args.flash
    provision = args.provision or flash
    rc = 0
    for port in ports:
        if not Path(port).exists():
            print(f"-- {port}: missing")
            continue
        mac = read_mac(port)
        if not mac:
            print(f"-- {port}: could not read MAC (chip in run mode? try --before usb-reset)")
            rc = 1
            continue
        match = find_node(nodes, mac)
        if match:
            print(f"== {port}  mac={mac}  node-id={match['node_id']}  env={match['pio_env']}  board={match.get('board','?')}")
            if flash and not do_flash(match, port):
                rc = 1
                continue
            if provision and not do_provision(match, port):
                rc = 1
        else:
            print(f"?? {port}  mac={mac}  UNKNOWN — not in nodes.yaml")
            if args.node_id is not None:
                slot = next((n for n in nodes if n.get("node_id") == args.node_id), None)
                if not slot:
                    print(f"!! node-id {args.node_id} not defined in nodes.yaml", file=sys.stderr)
                    rc = 1
                    continue
                print(f">> Claiming slot {args.node_id} ({slot.get('board','?')}) for new MAC.")
                slot = dict(slot, mac=mac)
                if flash and not do_flash(slot, port):
                    rc = 1
                    continue
                if provision and not do_provision(slot, port):
                    rc = 1
                    continue
                # Persist MAC into nodes.yaml
                txt = NODES_YAML.read_text()
                # Replace the empty mac line under matching node_id block
                pattern = re.compile(
                    rf'(- node_id: {args.node_id}\s*\n\s+mac: ")"',
                    re.MULTILINE,
                )
                new_txt, n_sub = pattern.subn(rf'\g<1>{mac}"', txt)
                if n_sub:
                    NODES_YAML.write_text(new_txt)
                    print(f">> Wrote MAC {mac} into nodes.yaml for node-id={args.node_id}")
            else:
                print(f"   Hint: re-run with --node-id <N> --provision to register and provision.")
                rc = 2
    return rc


if __name__ == "__main__":
    sys.exit(main())
