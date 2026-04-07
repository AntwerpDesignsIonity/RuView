#!/usr/bin/env python3
"""
ESP32-S3 CSI Node Provisioning Script

Writes WiFi credentials and aggregator target to the ESP32's NVS partition
so users can configure a pre-built firmware binary without recompiling.

Usage:
    python provision.py --ssid "MyWiFi" --password "secret" --target-ip 192.168.1.20

Requirements:
    pip install esptool nvs-partition-gen
    (or use the nvs_partition_gen.py bundled with ESP-IDF)
"""

import argparse
import csv
import io
import os
import re
import struct
import subprocess
import sys
import tempfile


# NVS partition table offset — default for ESP-IDF 4MB flash with standard
# partition scheme.  The "nvs" partition starts at 0x9000 (36864) and is
# 0x6000 (24576) bytes.
NVS_PARTITION_OFFSET = 0x9000
NVS_PARTITION_SIZE = 0x6000  # 24 KiB


def build_nvs_csv(args):
    """Build an NVS CSV string for the csi_cfg namespace."""
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["key", "type", "encoding", "value"])
    writer.writerow(["csi_cfg", "namespace", "", ""])
    if args.ssid:
        writer.writerow(["ssid", "data", "string", args.ssid])
    if args.password is not None:
        writer.writerow(["password", "data", "string", args.password])
    if args.target_ip:
        writer.writerow(["target_ip", "data", "string", args.target_ip])
    if args.target_port is not None:
        writer.writerow(["target_port", "data", "u16", str(args.target_port)])
    if args.node_id is not None:
        writer.writerow(["node_id", "data", "u8", str(args.node_id)])
    # TDM mesh settings
    if args.tdm_slot is not None:
        writer.writerow(["tdm_slot", "data", "u8", str(args.tdm_slot)])
    if args.tdm_total is not None:
        writer.writerow(["tdm_nodes", "data", "u8", str(args.tdm_total)])
    # Edge intelligence settings (ADR-039)
    if args.edge_tier is not None:
        writer.writerow(["edge_tier", "data", "u8", str(args.edge_tier)])
    if args.pres_thresh is not None:
        writer.writerow(["pres_thresh", "data", "u16", str(args.pres_thresh)])
    if args.fall_thresh is not None:
        writer.writerow(["fall_thresh", "data", "u16", str(args.fall_thresh)])
    if args.vital_win is not None:
        writer.writerow(["vital_win", "data", "u16", str(args.vital_win)])
    if args.vital_int is not None:
        writer.writerow(["vital_int", "data", "u16", str(args.vital_int)])
    if args.subk_count is not None:
        writer.writerow(["subk_count", "data", "u8", str(args.subk_count)])
    # ADR-060: Channel override and MAC filter
    if args.channel is not None:
        writer.writerow(["csi_channel", "data", "u8", str(args.channel)])
    if args.filter_mac is not None:
        mac_bytes = bytes(int(b, 16) for b in args.filter_mac.split(":"))
        # NVS blob: write as hex-encoded string for CSV compatibility
        writer.writerow(["filter_mac", "data", "hex2bin", mac_bytes.hex()])
    # ADR-073: Multi-frequency channel hopping
    if args.hop_channels is not None:
        channels = [int(c.strip()) for c in args.hop_channels.split(",")]
        writer.writerow(["hop_count", "data", "u8", str(len(channels))])
        # Store as NVS blob (firmware reads "chan_list" as uint8 blob)
        chan_bytes = bytes(channels)
        writer.writerow(["chan_list", "data", "hex2bin", chan_bytes.hex()])
        writer.writerow(["dwell_ms", "data", "u32", str(args.hop_dwell)])
    # Node role: 0=txrx (default), 1=tx (transmitter), 2=rx (receiver)
    if args.node_role is not None:
        role_map = {"txrx": 0, "tx": 1, "rx": 2}
        writer.writerow(["node_role", "data", "u8", str(role_map[args.node_role])])
    # LED hub/edge role (read by ionity/src/main.cpp): 1=hub, 0=edge
    if args.led_hub is not None:
        led_hub_val = 1 if args.led_hub == "hub" else 0
        writer.writerow(["led_hub", "data", "u8", str(led_hub_val)])
    # ADR-066: Swarm bridge configuration
    if args.seed_url is not None:
        writer.writerow(["seed_url", "data", "string", args.seed_url])
    if args.seed_token is not None:
        writer.writerow(["seed_token", "data", "string", args.seed_token])
    if args.zone is not None:
        writer.writerow(["zone_name", "data", "string", args.zone])
    if args.swarm_hb is not None:
        writer.writerow(["swarm_hb", "data", "u16", str(args.swarm_hb)])
    if args.swarm_ingest is not None:
        writer.writerow(["swarm_ingest", "data", "u16", str(args.swarm_ingest)])
    return buf.getvalue()


def generate_nvs_binary(csv_content, size):
    """Generate an NVS partition binary from CSV using nvs_partition_gen.py."""
    with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f_csv:
        f_csv.write(csv_content)
        csv_path = f_csv.name

    bin_path = csv_path.replace(".csv", ".bin")

    try:
        # Method 1: subprocess invocation (most reliable across package versions)
        for module_name in ["esp_idf_nvs_partition_gen", "nvs_partition_gen"]:
            try:
                subprocess.check_call(
                    [sys.executable, "-m", module_name, "generate",
                     csv_path, bin_path, hex(size)],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                )
                with open(bin_path, "rb") as f:
                    return f.read()
            except (subprocess.CalledProcessError, FileNotFoundError):
                continue

        # Method 2: ESP-IDF bundled script
        idf_path = os.environ.get("IDF_PATH", "")
        gen_script = os.path.join(idf_path, "components", "nvs_flash",
                                  "nvs_partition_generator", "nvs_partition_gen.py")
        if os.path.isfile(gen_script):
            subprocess.check_call([
                sys.executable, gen_script, "generate",
                csv_path, bin_path, hex(size)
            ])
            with open(bin_path, "rb") as f:
                return f.read()

        raise RuntimeError(
            "NVS partition generator not available. "
            "Install: pip install esp-idf-nvs-partition-gen"
        )

    finally:
        for p in (csv_path, bin_path):
            if os.path.isfile(p):
                os.unlink(p)


def flash_nvs(port, baud, nvs_bin):
    """Flash the NVS partition binary to the ESP32."""
    with tempfile.NamedTemporaryFile(suffix=".bin", delete=False) as f:
        f.write(nvs_bin)
        bin_path = f.name

    try:
        cmd = [
            sys.executable, "-m", "esptool",
            "--chip", "esp32s3",
            "--port", port,
            "--baud", str(baud),
            "write_flash",
            hex(NVS_PARTITION_OFFSET), bin_path,
        ]
        print(f"Flashing NVS partition ({len(nvs_bin)} bytes) to {port}...")
        subprocess.check_call(cmd)
        print("NVS provisioning complete!")
    finally:
        os.unlink(bin_path)


def detect_serial_port(preferred_port=None):
    """Detect the most likely ESP32 serial port when one is not provided."""
    if preferred_port:
        return preferred_port

    try:
        from serial.tools import list_ports
    except ImportError as exc:
        raise RuntimeError(
            "pyserial is required for auto-detect. Install esptool/pyserial or pass --port explicitly."
        ) from exc

    candidates = []
    for port in list_ports.comports():
        text = " ".join(filter(None, [port.device, port.description, port.manufacturer, port.hwid]))
        score = 0

        if re.search(r"VID:PID=(303A:|303A )|ESP32|Espressif|USB JTAG", text, re.IGNORECASE):
            score += 100
        if re.search(r"VID:PID=(10C4:|10C4 )|CP210|Silicon Labs", text, re.IGNORECASE):
            score += 60
        if re.search(r"VID:PID=(1A86:|1A86 )|CH340|wch", text, re.IGNORECASE):
            score += 40
        if re.search(r"USB Serial|UART", text, re.IGNORECASE):
            score += 15

        candidates.append({
            "device": port.device,
            "label": port.description or port.device,
            "score": score,
        })

    if not candidates:
        raise RuntimeError("No serial ports found. Connect the ESP32-S3 and try again.")

    candidates.sort(key=lambda item: (item["score"], item["device"]), reverse=True)
    if len(candidates) == 1 or candidates[0]["score"] > candidates[1]["score"]:
        print(f"Auto-detected ESP32 serial port: {candidates[0]['device']} [{candidates[0]['label']}]")
        return candidates[0]["device"]

    options = "\n".join(f"  {item['device']} - {item['label']}" for item in candidates)
    raise RuntimeError(
        "Unable to choose a single ESP32 serial port automatically. "
        "Pass --port explicitly. Candidates:\n" + options
    )


def main():
    parser = argparse.ArgumentParser(
        description="Provision ESP32-S3 CSI Node with WiFi and aggregator settings",
        epilog="Example: python provision.py --ssid MyWiFi --password secret --target-ip 192.168.1.20",
    )
    parser.add_argument("--port", help="Serial port (auto-detected when omitted, e.g. COM7, /dev/ttyUSB0)")
    parser.add_argument("--baud", type=int, default=460800, help="Flash baud rate (default: 460800)")
    parser.add_argument("--ssid", help="WiFi SSID")
    parser.add_argument("--password", help="WiFi password")
    parser.add_argument("--target-ip", help="Aggregator host IP (e.g. 192.168.1.20)")
    parser.add_argument("--target-port", type=int, help="Aggregator UDP port (default: 5005)")
    parser.add_argument("--node-id", type=int, help="Node ID 0-255 (default: 1)")
    # TDM mesh settings
    parser.add_argument("--tdm-slot", type=int, help="TDM slot index for this node (0-based)")
    parser.add_argument("--tdm-total", type=int, help="Total number of TDM nodes in mesh")
    # Edge intelligence settings (ADR-039)
    parser.add_argument("--edge-tier", type=int, choices=[0, 1, 2],
                        help="Edge processing tier: 0=off, 1=stats, 2=vitals")
    parser.add_argument("--pres-thresh", type=int, help="Presence detection threshold (default: 50)")
    parser.add_argument("--fall-thresh", type=int, help="Fall detection threshold in milli-units "
                        "(value/1000 = rad/s²). Default: 15000 → 15.0 rad/s². "
                        "Raise to reduce false positives in high-traffic areas.")
    parser.add_argument("--vital-win", type=int, help="Phase history window in frames (default: 300)")
    parser.add_argument("--vital-int", type=int, help="Vitals packet interval in ms (default: 1000)")
    parser.add_argument("--subk-count", type=int, help="Top-K subcarrier count (default: 32)")
    # ADR-060: Channel override and MAC filter
    parser.add_argument("--channel", type=int, help="CSI channel (1-14 for 2.4GHz, 36-177 for 5GHz). "
                        "Overrides auto-detection from connected AP.")
    parser.add_argument("--filter-mac", type=str, help="MAC address to filter CSI frames (AA:BB:CC:DD:EE:FF)")
    # ADR-073: Multi-frequency channel hopping
    parser.add_argument("--hop-channels", type=str, help="Comma-separated channel list for hopping (e.g. '1,6,11')")
    parser.add_argument("--hop-dwell", type=int, default=200, help="Dwell time per channel in ms (default: 200)")
    # Node role assignment for multistatic mesh
    parser.add_argument("--node-role", type=str, choices=["tx", "rx", "txrx"],
                        help="Node role: tx=transmitter (NDP injection), rx=receiver (CSI capture), "
                             "txrx=both (default bistatic mode)")
    # LED visual role (stored as NVS key 'led_hub', read by ionity/src/main.cpp)
    parser.add_argument("--led-hub", type=str, choices=["hub", "edge"],
                        help="LED role: hub=Deep Violet double-ripple (receiving/aggregator node), "
                             "edge=Teal single blink (transmitting sensor node)")
    # ADR-066: Swarm bridge
    parser.add_argument("--seed-url", type=str, help="Cognitum Seed base URL (e.g. http://10.1.10.236)")
    parser.add_argument("--seed-token", type=str, help="Seed Bearer token (from pairing)")
    parser.add_argument("--zone", type=str, help="Zone name for this node (e.g. lobby, hallway)")
    parser.add_argument("--swarm-hb", type=int, help="Swarm heartbeat interval in seconds (default 30)")
    parser.add_argument("--swarm-ingest", type=int, help="Swarm vector ingest interval in seconds (default 5)")
    parser.add_argument("--dry-run", action="store_true", help="Generate NVS binary but don't flash")

    args = parser.parse_args()

    has_value = any([
        args.ssid, args.password is not None, args.target_ip,
        args.target_port, args.node_id is not None,
        args.tdm_slot is not None, args.tdm_total is not None,
        args.edge_tier is not None, args.pres_thresh is not None,
        args.fall_thresh is not None, args.vital_win is not None,
        args.vital_int is not None, args.subk_count is not None,
        args.channel is not None, args.filter_mac is not None,
        args.seed_url is not None, args.zone is not None,
        args.node_role is not None, args.led_hub is not None,
    ])
    if not has_value:
        parser.error("At least one config value must be specified")

    # Validate TDM: if one is given, both should be
    if (args.tdm_slot is not None) != (args.tdm_total is not None):
        parser.error("--tdm-slot and --tdm-total must be specified together")
    if args.tdm_slot is not None and args.tdm_slot >= args.tdm_total:
        parser.error(f"--tdm-slot ({args.tdm_slot}) must be less than --tdm-total ({args.tdm_total})")

    # ADR-060: Validate channel and MAC filter
    if args.channel is not None:
        if not ((1 <= args.channel <= 14) or (36 <= args.channel <= 177)):
            parser.error(f"--channel must be 1-14 (2.4GHz) or 36-177 (5GHz), got {args.channel}")
    if args.filter_mac is not None:
        parts = args.filter_mac.split(":")
        if len(parts) != 6:
            parser.error(f"--filter-mac must be in AA:BB:CC:DD:EE:FF format, got '{args.filter_mac}'")
        try:
            for p in parts:
                val = int(p, 16)
                if val < 0 or val > 255:
                    raise ValueError
        except ValueError:
            parser.error(f"--filter-mac contains invalid hex bytes: '{args.filter_mac}'")

    print("Building NVS configuration:")
    if args.ssid:
        print(f"  WiFi SSID:     {args.ssid}")
    if args.password is not None:
        print(f"  WiFi Password: {'*' * len(args.password)}")
    if args.target_ip:
        print(f"  Target IP:     {args.target_ip}")
    if args.target_port:
        print(f"  Target Port:   {args.target_port}")
    if args.node_id is not None:
        print(f"  Node ID:       {args.node_id}")
    if args.node_role is not None:
        role_desc = {"tx": "transmitter (NDP injection)", "rx": "receiver (CSI capture)", "txrx": "transceiver (both)"}
        print(f"  Node Role:     {args.node_role} ({role_desc.get(args.node_role, '?')})")
    if args.led_hub is not None:
        led_desc = {"hub": "Deep Violet ripple (aggregator)", "edge": "Teal blink (sensor)"}
        print(f"  LED Role:      {args.led_hub} ({led_desc[args.led_hub]})")
    if args.tdm_slot is not None:
        print(f"  TDM Slot:      {args.tdm_slot} of {args.tdm_total}")
    if args.edge_tier is not None:
        tier_desc = {0: "off (raw CSI)", 1: "stats", 2: "vitals"}
        print(f"  Edge Tier:     {args.edge_tier} ({tier_desc.get(args.edge_tier, '?')})")
    if args.pres_thresh is not None:
        print(f"  Pres Thresh:   {args.pres_thresh}")
    if args.fall_thresh is not None:
        print(f"  Fall Thresh:   {args.fall_thresh}")
    if args.vital_win is not None:
        print(f"  Vital Window:  {args.vital_win} frames")
    if args.vital_int is not None:
        print(f"  Vital Interval:{args.vital_int} ms")
    if args.subk_count is not None:
        print(f"  Top-K Subcarr: {args.subk_count}")
    if args.channel is not None:
        print(f"  CSI Channel:   {args.channel}")
    if args.filter_mac is not None:
        print(f"  Filter MAC:    {args.filter_mac}")
    if args.seed_url is not None:
        print(f"  Seed URL:      {args.seed_url}")
    if args.zone is not None:
        print(f"  Zone:          {args.zone}")
    if args.swarm_hb is not None:
        print(f"  Swarm HB:      {args.swarm_hb}s")
    if args.swarm_ingest is not None:
        print(f"  Swarm Ingest:  {args.swarm_ingest}s")

    csv_content = build_nvs_csv(args)

    try:
        nvs_bin = generate_nvs_binary(csv_content, NVS_PARTITION_SIZE)
    except Exception as e:
        print(f"\nError generating NVS binary: {e}", file=sys.stderr)
        print("\nFallback: save CSV and flash manually with ESP-IDF tools.", file=sys.stderr)
        fallback_path = "nvs_config.csv"
        with open(fallback_path, "w") as f:
            f.write(csv_content)
        print(f"Saved NVS CSV to {fallback_path}", file=sys.stderr)
        print(f"Flash with: python $IDF_PATH/components/nvs_flash/"
              f"nvs_partition_generator/nvs_partition_gen.py generate "
              f"{fallback_path} nvs.bin 0x6000", file=sys.stderr)
        sys.exit(1)

    resolved_port = detect_serial_port(args.port)

    if args.dry_run:
        out = "nvs_provision.bin"
        with open(out, "wb") as f:
            f.write(nvs_bin)
        print(f"NVS binary saved to {out} ({len(nvs_bin)} bytes)")
        print(f"Flash manually: python -m esptool --chip esp32s3 --port {resolved_port} "
              f"write_flash 0x9000 {out}")
        return

    flash_nvs(resolved_port, args.baud, nvs_bin)


if __name__ == "__main__":
    main()
