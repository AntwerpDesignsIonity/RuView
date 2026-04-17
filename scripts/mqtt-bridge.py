#!/usr/bin/env python3
"""
AEDI-S MQTT Bridge — relays sensing-server WebSocket frames to MQTT.

Purpose
-------
* Subscribe to the Rust sensing-server WebSocket (`ws://HOST:3001/ws/sensing`)
* Republish normalized frames onto MQTT topics so external consumers
  (ML trainers, home-automation, BI tools) can tap into CSI data without
  a direct WS connection.
* Optional lightweight feature mapping for the WiFi-DensePose ("PyTorch-mini")
  inference pipeline — keeps a rolling buffer of subcarrier amplitudes and
  publishes derived feature vectors on `aedis/ml/features`.

Topics (prefix configurable, default `aedis/`)
----------------------------------------------
    aedis/sensing/frame      full JSON sensing frame (QoS 0, rate-limited)
    aedis/sensing/tick       monotonic tick counter
    aedis/nodes/<id>/rssi    per-node RSSI stream
    aedis/nodes/<id>/amps    per-node CSI amplitudes (binary float32 array)
    aedis/features           mapped feature vector (heart, breath, motion, ...)
    aedis/vitals             vital-signs JSON
    aedis/ml/features        compact vector for PyTorch-mini inference
    aedis/ml/inference       inference results (if a model is loaded)
    aedis/ml/cmd             (subscribe) external commands {op: "infer"|"reload"}
    aedis/status             bridge heartbeat (every 5 s)

Usage
-----
    pip install paho-mqtt websockets
    python scripts/mqtt-bridge.py \
        --broker localhost --port 1883 \
        --ws ws://localhost:3001/ws/sensing \
        --topic-prefix aedis/ \
        --rate-limit 5

If `paho-mqtt` is not installed this script prints an install hint and exits
with code 3 so the bridge-controller can report the missing dependency in the
GUI without crashing the stack.
"""

from __future__ import annotations

import argparse
import asyncio
import json
import os
import signal
import struct
import sys
import time
from pathlib import Path
from typing import Any

REPO_DIR = Path(__file__).resolve().parent.parent
LOG_DIR = REPO_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)
PID_FILE = LOG_DIR / "mqtt-bridge.pid"
LOG_FILE = LOG_DIR / "mqtt-bridge.log"


def _log(msg: str) -> None:
    line = f"[{time.strftime('%H:%M:%S')}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a") as f:
            f.write(line + "\n")
    except OSError:
        pass


def _require(pkg: str, name: str):
    try:
        return __import__(pkg)
    except ImportError:
        sys.stderr.write(
            f"[mqtt-bridge] Missing dependency: {name}\n"
            f"  Install with:  pip install {name}\n"
        )
        sys.exit(3)


# Lazy imports so the script can emit a clean error if deps are missing.
paho = _require("paho.mqtt.client", "paho-mqtt")
websockets_mod = _require("websockets", "websockets")

import paho.mqtt.client as mqtt  # noqa: E402
import websockets  # noqa: E402


class MqttBridge:
    def __init__(self, args: argparse.Namespace) -> None:
        self.args = args
        self.prefix = args.topic_prefix.rstrip("/") + "/"
        self.ws_url = args.ws
        self.rate_limit_hz = max(1, args.rate_limit)
        self.last_pub_ts = 0.0
        self.frame_count = 0
        self.mqtt_client = mqtt.Client(client_id=f"aedis-bridge-{os.getpid()}")
        self.mqtt_client.on_connect = self._on_mqtt_connect
        self.mqtt_client.on_disconnect = self._on_mqtt_disconnect
        self.mqtt_client.on_message = self._on_mqtt_message
        if args.username:
            self.mqtt_client.username_pw_set(args.username, args.password or "")
        self._stop = asyncio.Event()

        # Rolling buffer for ML features (last N frames of amplitudes).
        self.ml_buffer: list[list[float]] = []
        self.ml_buffer_len = 32

        # Optional torch model (lazy-loaded).
        self.torch_model = None
        self.torch_device = "cpu"

    # ── MQTT callbacks ────────────────────────────────────────────────────
    def _on_mqtt_connect(
        self, client: Any, userdata: Any, flags: Any, rc: int
    ) -> None:
        if rc == 0:
            _log(f"MQTT connected to {self.args.broker}:{self.args.port}")
            client.subscribe(self.prefix + "ml/cmd", qos=1)
            client.publish(
                self.prefix + "status",
                json.dumps({"state": "online", "pid": os.getpid(), "ts": time.time()}),
                qos=1,
                retain=True,
            )
        else:
            _log(f"MQTT connect failed rc={rc}")

    def _on_mqtt_disconnect(self, client: Any, userdata: Any, rc: int) -> None:
        _log(f"MQTT disconnected rc={rc}")

    def _on_mqtt_message(self, client: Any, userdata: Any, msg: Any) -> None:
        try:
            cmd = json.loads(msg.payload.decode("utf-8"))
        except (json.JSONDecodeError, UnicodeDecodeError):
            return
        op = cmd.get("op")
        if op == "infer":
            asyncio.get_event_loop().create_task(self._run_inference())
        elif op == "reload":
            self._load_torch_model()
        elif op == "ping":
            client.publish(
                self.prefix + "status",
                json.dumps({"pong": True, "ts": time.time()}),
            )

    # ── Optional PyTorch-mini inference ──────────────────────────────────
    def _load_torch_model(self) -> None:
        """Best-effort load of references/wifi_densepose_pytorch.py."""
        try:
            import torch  # type: ignore
        except ImportError:
            _log("torch not installed — inference disabled")
            return
        try:
            sys.path.insert(0, str(REPO_DIR / "references"))
            from wifi_densepose_pytorch import (  # type: ignore
                ModalityTranslationNetwork,
            )
            self.torch_model = ModalityTranslationNetwork().to(self.torch_device)
            self.torch_model.eval()
            _log("PyTorch-mini model loaded (ModalityTranslationNetwork)")
        except Exception as exc:  # noqa: BLE001
            _log(f"torch model load failed: {exc}")
            self.torch_model = None

    async def _run_inference(self) -> None:
        if self.torch_model is None or not self.ml_buffer:
            self.mqtt_client.publish(
                self.prefix + "ml/inference",
                json.dumps({"error": "model not loaded or empty buffer"}),
            )
            return
        try:
            import torch  # type: ignore
        except ImportError:
            return
        amps = self.ml_buffer[-1]
        x = torch.tensor(amps, dtype=torch.float32).view(1, 1, 1, -1)
        with torch.no_grad():
            out = self.torch_model(x)
        result = {
            "shape": list(out.shape) if hasattr(out, "shape") else None,
            "mean": float(out.mean().item()) if hasattr(out, "mean") else None,
            "ts": time.time(),
        }
        self.mqtt_client.publish(self.prefix + "ml/inference", json.dumps(result))

    # ── WebSocket ingest → MQTT publish ──────────────────────────────────
    def _mapped_features(self, frame: dict) -> dict:
        """Compact feature vector for ML pipelines."""
        feats = frame.get("features", {}) or {}
        vs = frame.get("vital_signs", {}) or {}
        cls = frame.get("classification", {}) or {}
        return {
            "ts": time.time(),
            "tick": frame.get("tick"),
            "mean_rssi": feats.get("mean_rssi"),
            "variance": feats.get("variance"),
            "spectral_power": feats.get("spectral_power"),
            "dominant_freq_hz": feats.get("dominant_freq_hz"),
            "hr": vs.get("heart_rate_bpm"),
            "br": vs.get("breathing_rate_bpm"),
            "motion": cls.get("motion_level"),
            "presence": bool(cls.get("presence")),
            "confidence": cls.get("confidence"),
            "quality": frame.get("signal_quality_score"),
        }

    def _publish_frame(self, frame: dict) -> None:
        now = time.time()
        min_dt = 1.0 / self.rate_limit_hz
        if now - self.last_pub_ts < min_dt:
            return  # rate-limit protects MQTT broker
        self.last_pub_ts = now
        self.frame_count += 1

        c = self.mqtt_client

        # 1) Full frame (JSON, rate-limited)
        c.publish(self.prefix + "sensing/frame", json.dumps(frame), qos=0)

        # 2) Tick
        tick = frame.get("tick")
        if tick is not None:
            c.publish(self.prefix + "sensing/tick", str(tick), qos=0)

        # 3) Per-node streams
        for node in frame.get("nodes", []) or []:
            nid = node.get("node_id")
            if nid is None:
                continue
            base = self.prefix + f"nodes/{nid}/"
            if "rssi_dbm" in node:
                c.publish(base + "rssi", f"{node['rssi_dbm']:.1f}", qos=0)
            amps = node.get("amplitude") or []
            if amps:
                # Binary float32 payload — efficient for downstream ML.
                c.publish(
                    base + "amps",
                    struct.pack(f"{len(amps)}f", *amps),
                    qos=0,
                )
                self.ml_buffer.append(list(amps))
                if len(self.ml_buffer) > self.ml_buffer_len:
                    self.ml_buffer.pop(0)

        # 4) Vitals
        if frame.get("vital_signs"):
            c.publish(
                self.prefix + "vitals",
                json.dumps(frame["vital_signs"]),
                qos=0,
            )

        # 5) Compact ML features
        c.publish(
            self.prefix + "ml/features",
            json.dumps(self._mapped_features(frame)),
            qos=0,
        )

    async def _heartbeat(self) -> None:
        while not self._stop.is_set():
            try:
                self.mqtt_client.publish(
                    self.prefix + "status",
                    json.dumps(
                        {
                            "state": "running",
                            "frames": self.frame_count,
                            "ml_buffer": len(self.ml_buffer),
                            "torch_loaded": self.torch_model is not None,
                            "ts": time.time(),
                        }
                    ),
                )
            except Exception:  # noqa: BLE001
                pass
            try:
                await asyncio.wait_for(self._stop.wait(), timeout=5.0)
            except asyncio.TimeoutError:
                pass

    async def _ws_loop(self) -> None:
        backoff = 1.0
        while not self._stop.is_set():
            try:
                _log(f"Connecting WS {self.ws_url}")
                async with websockets.connect(
                    self.ws_url, max_size=2**22
                ) as ws:
                    _log("WS connected")
                    backoff = 1.0
                    async for msg in ws:
                        if self._stop.is_set():
                            break
                        try:
                            frame = json.loads(msg)
                        except json.JSONDecodeError:
                            continue
                        self._publish_frame(frame)
            except Exception as exc:  # noqa: BLE001
                _log(f"WS error: {exc} — retry in {backoff:.1f}s")
                try:
                    await asyncio.wait_for(self._stop.wait(), timeout=backoff)
                except asyncio.TimeoutError:
                    pass
                backoff = min(backoff * 1.6, 15.0)

    # ── Lifecycle ────────────────────────────────────────────────────────
    async def run(self) -> None:
        PID_FILE.write_text(str(os.getpid()))
        loop = asyncio.get_running_loop()
        for sig in (signal.SIGINT, signal.SIGTERM):
            loop.add_signal_handler(sig, self._stop.set)

        try:
            self.mqtt_client.connect(self.args.broker, self.args.port, keepalive=30)
        except Exception as exc:  # noqa: BLE001
            _log(f"MQTT connect failed: {exc}")
            return
        self.mqtt_client.loop_start()

        if self.args.load_model:
            self._load_torch_model()

        await asyncio.gather(self._ws_loop(), self._heartbeat())

        _log("shutting down")
        try:
            self.mqtt_client.publish(
                self.prefix + "status",
                json.dumps({"state": "offline", "ts": time.time()}),
                qos=1,
                retain=True,
            )
        finally:
            self.mqtt_client.loop_stop()
            self.mqtt_client.disconnect()
            if PID_FILE.exists():
                PID_FILE.unlink(missing_ok=True)


def main() -> None:
    p = argparse.ArgumentParser(description="AEDI-S MQTT bridge")
    p.add_argument("--broker", default=os.getenv("MQTT_BROKER", "localhost"))
    p.add_argument("--port", type=int, default=int(os.getenv("MQTT_PORT", "1883")))
    p.add_argument("--username", default=os.getenv("MQTT_USER"))
    p.add_argument("--password", default=os.getenv("MQTT_PASS"))
    p.add_argument(
        "--ws",
        default=os.getenv("AEDIS_WS", "ws://localhost:3001/ws/sensing"),
    )
    p.add_argument(
        "--topic-prefix", default=os.getenv("MQTT_TOPIC_PREFIX", "aedis/")
    )
    p.add_argument(
        "--rate-limit",
        type=int,
        default=int(os.getenv("MQTT_RATE_HZ", "5")),
        help="Max frames/sec published to MQTT",
    )
    p.add_argument(
        "--load-model",
        action="store_true",
        help="Load PyTorch-mini model at startup",
    )
    args = p.parse_args()

    bridge = MqttBridge(args)
    try:
        asyncio.run(bridge.run())
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
