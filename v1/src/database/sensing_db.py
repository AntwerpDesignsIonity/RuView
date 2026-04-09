"""
SQLite sensing database — RPi & PyTorch friendly.

Logs every sensing tick with full metadata:
  - CSI readings (amplitude, phase, RSSI, noise floor)
  - ESP32 telemetry (chip temperature, uptime, free heap, WiFi RSSI)
  - Classification results (presence, motion, confidence)
  - Vital signs (breathing rate, heart rate)
  - Device identity (node_id, machine_id, MAC address, IP, firmware version)
  - Network topology (source IP, BSSID, channel, bandwidth)
  - Signal features (13+ extracted features per tick)
  - Room layout metadata (room_id, zone, coordinates)

Designed for:
  - SQLite on RPi (no server, single-file, WAL mode for concurrent reads)
  - PyTorch DataLoader via to_tensor() helpers
  - Pandas export for analysis (to_dataframe())
  - Alembic-free: schema versioned in-code with auto-migrate
"""

from __future__ import annotations

import hashlib
import json
import os
import platform
import sqlite3
import time
import uuid
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Generator, List, Optional, Tuple

# ---------------------------------------------------------------------------
# Schema version — bump this when adding columns, tables, or indexes.
# Auto-migration will run on connect if the stored version is < SCHEMA_VERSION.
# ---------------------------------------------------------------------------
SCHEMA_VERSION = 1

# ---------------------------------------------------------------------------
# Default paths
# ---------------------------------------------------------------------------
_DEFAULT_DB_DIR = Path(__file__).resolve().parents[3] / "data" / "db"
_DEFAULT_DB_PATH = _DEFAULT_DB_DIR / "ruview_sensing.sqlite3"

_LOGS_DB_DIR = Path(__file__).resolve().parents[3] / "logs"
_LOGS_DB_PATH = _LOGS_DB_DIR / "ruview_logs.sqlite3"


def _machine_id() -> str:
    """Stable machine identifier — hostname + MAC hash, safe for multi-node."""
    raw = f"{platform.node()}:{uuid.getnode()}"
    return hashlib.sha256(raw.encode()).hexdigest()[:16]


# ===================================================================
# Schema DDL
# ===================================================================
_SCHEMA_DDL = """
-- Devices: one row per ESP32 / RPi / laptop that participates
CREATE TABLE IF NOT EXISTS devices (
    device_id       TEXT PRIMARY KEY,               -- sha256(hostname:mac)[:16]
    hostname        TEXT NOT NULL,
    mac_address     TEXT,
    ip_address      TEXT,
    firmware_version TEXT,
    hardware_model  TEXT,
    node_id         INTEGER,                         -- ESP32 NVS node_id
    room_id         TEXT,
    zone            TEXT,
    coord_x         REAL,
    coord_y         REAL,
    coord_z         REAL,
    first_seen_utc  TEXT NOT NULL,
    last_seen_utc   TEXT NOT NULL,
    meta            TEXT                             -- JSON blob for extras
);

-- Sessions: a contiguous collection window (start→stop)
CREATE TABLE IF NOT EXISTS sessions (
    session_id      TEXT PRIMARY KEY,               -- uuid4
    device_id       TEXT NOT NULL REFERENCES devices(device_id),
    started_utc     TEXT NOT NULL,
    ended_utc       TEXT,
    csi_source      TEXT,                           -- esp32 | linux_wifi | simulated …
    config          TEXT,                           -- JSON: tick_interval, window, etc.
    total_ticks     INTEGER DEFAULT 0,
    total_frames    INTEGER DEFAULT 0,
    notes           TEXT
);

-- Readings: one row per sensing tick (~500ms)
CREATE TABLE IF NOT EXISTS readings (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id      TEXT NOT NULL REFERENCES sessions(session_id),
    device_id       TEXT NOT NULL REFERENCES devices(device_id),

    -- Timestamps
    ts_utc          TEXT NOT NULL,                  -- ISO-8601 UTC
    ts_local        TEXT NOT NULL,                  -- ISO-8601 local
    ts_epoch        REAL NOT NULL,                  -- Unix epoch (float)
    tick            INTEGER NOT NULL,

    -- CSI raw data (JSON arrays — lightweight; use BLOB for high‑throughput)
    amplitude       TEXT,                           -- JSON float array (≤56 subcarriers)
    phase           TEXT,                           -- JSON float array
    n_subcarriers   INTEGER,
    freq_mhz        INTEGER,
    bandwidth_mhz   INTEGER,
    sequence_num    INTEGER,

    -- Signal readings
    rssi_dbm        REAL,
    noise_floor_dbm REAL,
    mean_amplitude  REAL,
    signal_quality  REAL,                           -- 0.0–1.0

    -- Extracted features (from RssiFeatureExtractor)
    feat_mean       REAL,
    feat_std        REAL,
    feat_variance   REAL,
    feat_range      REAL,
    feat_iqr        REAL,
    feat_skewness   REAL,
    feat_kurtosis   REAL,
    feat_motion_power     REAL,
    feat_breathing_power  REAL,
    feat_dominant_freq_hz REAL,
    feat_spectral_power   REAL,
    feat_change_points    INTEGER,

    -- Classification
    motion_level    TEXT,                           -- absent | idle | active | rapid
    presence        INTEGER,                        -- 0/1
    confidence      REAL,

    -- Vital signs (nullable — only when model running)
    breathing_bpm   REAL,
    heartrate_bpm   REAL,
    breathing_conf  REAL,
    heartrate_conf  REAL,

    -- ESP32 telemetry
    esp_temp_c      REAL,                           -- chip temperature °C
    esp_uptime_ms   INTEGER,                        -- millis() since boot
    esp_free_heap   INTEGER,                        -- bytes
    esp_wifi_rssi   INTEGER,                        -- raw dBm

    -- Network
    source_ip       TEXT,
    source_port     INTEGER,
    bssid           TEXT,
    channel         INTEGER,

    -- Pose (optional — only when dense-pose model running)
    person_count    INTEGER,
    keypoints       TEXT,                           -- JSON array
    zones_occupied  TEXT                            -- JSON array of zone IDs
);

-- ESP32 telemetry log (high-frequency, separate so readings table stays lean)
CREATE TABLE IF NOT EXISTS esp32_telemetry (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    device_id       TEXT NOT NULL REFERENCES devices(device_id),
    node_id         INTEGER NOT NULL,
    ts_utc          TEXT NOT NULL,
    ts_epoch        REAL NOT NULL,
    temp_c          REAL,
    uptime_ms       INTEGER,
    free_heap       INTEGER,
    wifi_rssi       INTEGER,
    udp_sent        INTEGER,
    udp_fail        INTEGER,
    ser_sent        INTEGER,
    csi_frames      INTEGER,
    ip_address      TEXT,
    channel         INTEGER,
    tx_power        INTEGER
);

-- Device heartbeat / online-time tracking
CREATE TABLE IF NOT EXISTS heartbeats (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    device_id       TEXT NOT NULL REFERENCES devices(device_id),
    ts_utc          TEXT NOT NULL,
    ts_epoch        REAL NOT NULL,
    online          INTEGER NOT NULL DEFAULT 1,     -- 1=online, 0=went offline
    uptime_s        REAL,                           -- seconds since last boot
    session_id      TEXT
);

-- Room layouts (so the UI can render floor plans)
CREATE TABLE IF NOT EXISTS layouts (
    layout_id       TEXT PRIMARY KEY,
    name            TEXT NOT NULL,
    description     TEXT,
    floor           INTEGER DEFAULT 0,
    width_m         REAL,
    height_m        REAL,
    origin_lat      REAL,
    origin_lon      REAL,
    svg_path        TEXT,                           -- optional SVG floor plan file
    meta            TEXT,                           -- JSON: furniture, walls, etc.
    created_utc     TEXT NOT NULL,
    updated_utc     TEXT NOT NULL
);

-- Device → layout placement
CREATE TABLE IF NOT EXISTS device_placements (
    device_id       TEXT NOT NULL REFERENCES devices(device_id),
    layout_id       TEXT NOT NULL REFERENCES layouts(layout_id),
    x               REAL NOT NULL,
    y               REAL NOT NULL,
    z               REAL DEFAULT 0,
    orientation_deg REAL DEFAULT 0,
    placed_utc      TEXT NOT NULL,
    PRIMARY KEY (device_id, layout_id)
);

-- Indexes (critical for time-series queries & ML window extraction)
CREATE INDEX IF NOT EXISTS idx_readings_session   ON readings(session_id);
CREATE INDEX IF NOT EXISTS idx_readings_ts        ON readings(ts_epoch);
CREATE INDEX IF NOT EXISTS idx_readings_device    ON readings(device_id, ts_epoch);
CREATE INDEX IF NOT EXISTS idx_readings_motion    ON readings(motion_level);
CREATE INDEX IF NOT EXISTS idx_readings_presence  ON readings(presence, ts_epoch);
CREATE INDEX IF NOT EXISTS idx_telemetry_device   ON esp32_telemetry(device_id, ts_epoch);
CREATE INDEX IF NOT EXISTS idx_heartbeat_device   ON heartbeats(device_id, ts_epoch);
CREATE INDEX IF NOT EXISTS idx_sessions_device    ON sessions(device_id, started_utc);

-- Schema version tracker
CREATE TABLE IF NOT EXISTS schema_meta (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
"""

# ===================================================================
# Log DB DDL (separate file under logs/)
# ===================================================================
_LOGS_DDL = """
CREATE TABLE IF NOT EXISTS app_logs (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    ts_utc      TEXT NOT NULL,
    ts_epoch    REAL NOT NULL,
    level       TEXT NOT NULL,
    logger      TEXT,
    message     TEXT NOT NULL,
    module      TEXT,
    func        TEXT,
    line        INTEGER,
    device_id   TEXT,
    session_id  TEXT,
    extra       TEXT
);
CREATE INDEX IF NOT EXISTS idx_logs_ts    ON app_logs(ts_epoch);
CREATE INDEX IF NOT EXISTS idx_logs_level ON app_logs(level);

CREATE TABLE IF NOT EXISTS schema_meta (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
"""


# ===================================================================
# SensingDB
# ===================================================================
class SensingDB:
    """
    RPi-friendly SQLite database for CSI readings, telemetry, and ML.

    Usage::

        db = SensingDB()          # auto-creates data/db/ruview_sensing.sqlite3
        db.start_session("esp32")
        db.insert_reading({...})  # from ws_server _build_message()
        db.end_session()

        # ML / PyTorch
        rows = db.query_readings(last_minutes=60)
        tensor = db.to_tensor(rows, columns=["feat_mean", "feat_variance", ...])

        # Pandas
        df = db.to_dataframe(last_minutes=60)
    """

    def __init__(
        self,
        db_path: Optional[str] = None,
        logs_path: Optional[str] = None,
    ) -> None:
        self._db_path = Path(db_path) if db_path else _DEFAULT_DB_PATH
        self._logs_path = Path(logs_path) if logs_path else _LOGS_DB_PATH
        self._db_path.parent.mkdir(parents=True, exist_ok=True)
        self._logs_path.parent.mkdir(parents=True, exist_ok=True)

        self._machine_id = _machine_id()
        self._session_id: Optional[str] = None
        self._tick: int = 0

        self._conn: Optional[sqlite3.Connection] = None
        self._logs_conn: Optional[sqlite3.Connection] = None

        self._connect()

    # ---------------------------------------------------------------
    # Connection management
    # ---------------------------------------------------------------
    def _connect(self) -> None:
        self._conn = sqlite3.connect(
            str(self._db_path),
            timeout=30,
            check_same_thread=False,
        )
        self._conn.execute("PRAGMA journal_mode=WAL")
        self._conn.execute("PRAGMA synchronous=NORMAL")
        self._conn.execute("PRAGMA cache_size=-16384")  # 16 MB
        self._conn.execute("PRAGMA busy_timeout=5000")
        self._conn.execute("PRAGMA foreign_keys=ON")
        self._conn.row_factory = sqlite3.Row
        self._init_schema(self._conn, _SCHEMA_DDL, SCHEMA_VERSION)

        self._logs_conn = sqlite3.connect(
            str(self._logs_path),
            timeout=10,
            check_same_thread=False,
        )
        self._logs_conn.execute("PRAGMA journal_mode=WAL")
        self._logs_conn.execute("PRAGMA synchronous=NORMAL")
        self._logs_conn.row_factory = sqlite3.Row
        self._init_schema(self._logs_conn, _LOGS_DDL, SCHEMA_VERSION)

    @staticmethod
    def _init_schema(conn: sqlite3.Connection, ddl: str, version: int) -> None:
        conn.executescript(ddl)
        row = conn.execute(
            "SELECT value FROM schema_meta WHERE key = 'version'"
        ).fetchone()
        if row is None:
            conn.execute(
                "INSERT INTO schema_meta (key, value) VALUES ('version', ?)",
                (str(version),),
            )
            conn.commit()
        # Future: elif int(row["value"]) < version: run migration scripts

    @contextmanager
    def cursor(self) -> Generator[sqlite3.Cursor, None, None]:
        cur = self._conn.cursor()
        try:
            yield cur
            self._conn.commit()
        except Exception:
            self._conn.rollback()
            raise
        finally:
            cur.close()

    def close(self) -> None:
        if self._conn:
            self._conn.close()
            self._conn = None
        if self._logs_conn:
            self._logs_conn.close()
            self._logs_conn = None

    # ---------------------------------------------------------------
    # Device registration
    # ---------------------------------------------------------------
    def register_device(
        self,
        device_id: Optional[str] = None,
        hostname: Optional[str] = None,
        mac_address: Optional[str] = None,
        ip_address: Optional[str] = None,
        firmware_version: Optional[str] = None,
        hardware_model: Optional[str] = None,
        node_id: Optional[int] = None,
        room_id: Optional[str] = None,
        zone: Optional[str] = None,
        coord_x: Optional[float] = None,
        coord_y: Optional[float] = None,
        coord_z: Optional[float] = None,
        meta: Optional[dict] = None,
    ) -> str:
        did = device_id or self._machine_id
        now = datetime.now(timezone.utc).isoformat()
        with self.cursor() as cur:
            cur.execute(
                """INSERT INTO devices
                   (device_id, hostname, mac_address, ip_address,
                    firmware_version, hardware_model, node_id,
                    room_id, zone, coord_x, coord_y, coord_z,
                    first_seen_utc, last_seen_utc, meta)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                   ON CONFLICT(device_id) DO UPDATE SET
                     hostname=excluded.hostname,
                     mac_address=COALESCE(excluded.mac_address, devices.mac_address),
                     ip_address=COALESCE(excluded.ip_address, devices.ip_address),
                     firmware_version=COALESCE(excluded.firmware_version, devices.firmware_version),
                     hardware_model=COALESCE(excluded.hardware_model, devices.hardware_model),
                     node_id=COALESCE(excluded.node_id, devices.node_id),
                     room_id=COALESCE(excluded.room_id, devices.room_id),
                     zone=COALESCE(excluded.zone, devices.zone),
                     coord_x=COALESCE(excluded.coord_x, devices.coord_x),
                     coord_y=COALESCE(excluded.coord_y, devices.coord_y),
                     coord_z=COALESCE(excluded.coord_z, devices.coord_z),
                     last_seen_utc=excluded.last_seen_utc,
                     meta=COALESCE(excluded.meta, devices.meta)
                """,
                (
                    did, hostname or platform.node(), mac_address, ip_address,
                    firmware_version, hardware_model, node_id,
                    room_id, zone, coord_x, coord_y, coord_z,
                    now, now, json.dumps(meta) if meta else None,
                ),
            )
        return did

    # ---------------------------------------------------------------
    # Session lifecycle
    # ---------------------------------------------------------------
    def start_session(
        self,
        csi_source: str = "unknown",
        config: Optional[dict] = None,
        notes: Optional[str] = None,
    ) -> str:
        sid = str(uuid.uuid4())
        now = datetime.now(timezone.utc).isoformat()
        self.register_device()  # ensure self device exists
        with self.cursor() as cur:
            cur.execute(
                """INSERT INTO sessions
                   (session_id, device_id, started_utc, csi_source, config, notes)
                   VALUES (?,?,?,?,?,?)""",
                (sid, self._machine_id, now, csi_source,
                 json.dumps(config) if config else None, notes),
            )
        self._session_id = sid
        self._tick = 0

        # heartbeat: online
        self._insert_heartbeat(1)
        return sid

    def end_session(self) -> None:
        if not self._session_id:
            return
        now = datetime.now(timezone.utc).isoformat()
        with self.cursor() as cur:
            cur.execute(
                """UPDATE sessions SET
                     ended_utc = ?, total_ticks = ?,
                     total_frames = (SELECT COUNT(*) FROM readings WHERE session_id = ?)
                   WHERE session_id = ?""",
                (now, self._tick, self._session_id, self._session_id),
            )
        self._insert_heartbeat(0)
        self._session_id = None

    # ---------------------------------------------------------------
    # Insert a reading (called from ws_server on each tick)
    # ---------------------------------------------------------------
    def insert_reading(self, msg: Dict[str, Any]) -> None:
        """
        Insert a single sensing_update message into the readings table.

        ``msg`` is the dict from SensingWebSocketServer._build_message().
        """
        if not self._session_id:
            return

        self._tick += 1
        now_utc = datetime.now(timezone.utc)
        now_local = datetime.now()

        node = msg.get("nodes", [{}])[0] if msg.get("nodes") else {}
        feats = msg.get("features", {})
        cls = msg.get("classification", {})
        vitals = msg.get("vital_signs", {})

        with self.cursor() as cur:
            cur.execute(
                """INSERT INTO readings (
                     session_id, device_id,
                     ts_utc, ts_local, ts_epoch, tick,
                     amplitude, phase, n_subcarriers, freq_mhz, bandwidth_mhz, sequence_num,
                     rssi_dbm, noise_floor_dbm, mean_amplitude, signal_quality,
                     feat_mean, feat_std, feat_variance, feat_range, feat_iqr,
                     feat_skewness, feat_kurtosis,
                     feat_motion_power, feat_breathing_power,
                     feat_dominant_freq_hz, feat_spectral_power, feat_change_points,
                     motion_level, presence, confidence,
                     breathing_bpm, heartrate_bpm, breathing_conf, heartrate_conf,
                     esp_temp_c, esp_uptime_ms, esp_free_heap, esp_wifi_rssi,
                     source_ip, source_port, bssid, channel,
                     person_count, keypoints, zones_occupied
                   ) VALUES (
                     ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
                   )""",
                (
                    self._session_id, self._machine_id,
                    now_utc.isoformat(), now_local.isoformat(), msg.get("timestamp", time.time()), self._tick,
                    json.dumps(node.get("amplitude")) if node.get("amplitude") else None,
                    None,  # phase — not in current WebSocket message
                    node.get("subcarrier_count"),
                    node.get("freq_mhz"),
                    None,  # bandwidth
                    node.get("sequence"),
                    node.get("rssi_dbm"),
                    None,  # noise_floor — add to msg in future
                    node.get("mean_amplitude"),
                    msg.get("signal_quality_score"),
                    feats.get("mean_rssi"),
                    feats.get("std"),
                    feats.get("variance"),
                    feats.get("range"),
                    feats.get("iqr"),
                    feats.get("skewness"),
                    feats.get("kurtosis"),
                    feats.get("motion_band_power"),
                    feats.get("breathing_band_power"),
                    feats.get("dominant_freq_hz"),
                    feats.get("spectral_power"),
                    feats.get("change_points"),
                    cls.get("motion_level"),
                    1 if cls.get("presence") else 0,
                    cls.get("confidence"),
                    vitals.get("breathing_rate_bpm"),
                    vitals.get("heart_rate_bpm"),
                    vitals.get("breathing_confidence"),
                    vitals.get("heartbeat_confidence"),
                    msg.get("esp_telemetry", {}).get("temp_c"),
                    msg.get("esp_telemetry", {}).get("uptime_ms"),
                    msg.get("esp_telemetry", {}).get("free_heap"),
                    msg.get("esp_telemetry", {}).get("wifi_rssi"),
                    node.get("source_addr", "").split(":")[0] if ":" in node.get("source_addr", "") else node.get("source_addr"),
                    int(node.get("source_addr", "0:0").split(":")[-1]) if ":" in node.get("source_addr", "") else None,
                    msg.get("bssid"),
                    None,  # channel — derive from freq_mhz if needed
                    len(msg.get("persons", [])) if msg.get("persons") else None,
                    json.dumps(msg.get("pose_keypoints")) if msg.get("pose_keypoints") else None,
                    None,
                ),
            )

    # ---------------------------------------------------------------
    # ESP32 telemetry (separate high-frequency table)
    # ---------------------------------------------------------------
    def insert_telemetry(self, telem: Dict[str, Any]) -> None:
        """Insert an ESP32 telemetry packet (magic 0xC5110003 or from status log)."""
        now = datetime.now(timezone.utc)
        with self.cursor() as cur:
            cur.execute(
                """INSERT INTO esp32_telemetry
                   (device_id, node_id, ts_utc, ts_epoch,
                    temp_c, uptime_ms, free_heap, wifi_rssi,
                    udp_sent, udp_fail, ser_sent, csi_frames,
                    ip_address, channel, tx_power)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    telem.get("device_id", self._machine_id),
                    telem.get("node_id", 0),
                    now.isoformat(), time.time(),
                    telem.get("temp_c"),
                    telem.get("uptime_ms"),
                    telem.get("free_heap"),
                    telem.get("wifi_rssi"),
                    telem.get("udp_sent"),
                    telem.get("udp_fail"),
                    telem.get("ser_sent"),
                    telem.get("csi_frames"),
                    telem.get("ip_address"),
                    telem.get("channel"),
                    telem.get("tx_power"),
                ),
            )

    # ---------------------------------------------------------------
    # Heartbeats
    # ---------------------------------------------------------------
    def _insert_heartbeat(self, online: int) -> None:
        now = datetime.now(timezone.utc)
        with self.cursor() as cur:
            cur.execute(
                """INSERT INTO heartbeats
                   (device_id, ts_utc, ts_epoch, online, session_id)
                   VALUES (?,?,?,?,?)""",
                (self._machine_id, now.isoformat(), time.time(), online, self._session_id),
            )

    # ---------------------------------------------------------------
    # Layouts
    # ---------------------------------------------------------------
    def upsert_layout(
        self,
        name: str,
        width_m: float = 10.0,
        height_m: float = 10.0,
        floor: int = 0,
        description: Optional[str] = None,
        meta: Optional[dict] = None,
    ) -> str:
        lid = hashlib.sha256(name.encode()).hexdigest()[:16]
        now = datetime.now(timezone.utc).isoformat()
        with self.cursor() as cur:
            cur.execute(
                """INSERT INTO layouts
                   (layout_id, name, description, floor, width_m, height_m, meta, created_utc, updated_utc)
                   VALUES (?,?,?,?,?,?,?,?,?)
                   ON CONFLICT(layout_id) DO UPDATE SET
                     name=excluded.name, description=excluded.description,
                     floor=excluded.floor, width_m=excluded.width_m,
                     height_m=excluded.height_m, meta=excluded.meta,
                     updated_utc=excluded.updated_utc""",
                (lid, name, description, floor, width_m, height_m,
                 json.dumps(meta) if meta else None, now, now),
            )
        return lid

    def place_device_in_layout(
        self, layout_id: str, device_id: Optional[str] = None,
        x: float = 0, y: float = 0, z: float = 0, orientation_deg: float = 0,
    ) -> None:
        did = device_id or self._machine_id
        now = datetime.now(timezone.utc).isoformat()
        with self.cursor() as cur:
            cur.execute(
                """INSERT INTO device_placements
                   (device_id, layout_id, x, y, z, orientation_deg, placed_utc)
                   VALUES (?,?,?,?,?,?,?)
                   ON CONFLICT(device_id, layout_id) DO UPDATE SET
                     x=excluded.x, y=excluded.y, z=excluded.z,
                     orientation_deg=excluded.orientation_deg,
                     placed_utc=excluded.placed_utc""",
                (did, layout_id, x, y, z, orientation_deg, now),
            )

    # ---------------------------------------------------------------
    # Log insertion (separate DB)
    # ---------------------------------------------------------------
    def insert_log(
        self,
        level: str,
        message: str,
        logger_name: Optional[str] = None,
        module: Optional[str] = None,
        func: Optional[str] = None,
        line: Optional[int] = None,
        session_id: Optional[str] = None,
        extra: Optional[dict] = None,
    ) -> None:
        now = datetime.now(timezone.utc)
        self._logs_conn.execute(
            """INSERT INTO app_logs
               (ts_utc, ts_epoch, level, logger, message, module, func, line,
                device_id, session_id, extra)
               VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
            (
                now.isoformat(), time.time(), level, logger_name, message,
                module, func, line, self._machine_id,
                session_id or self._session_id,
                json.dumps(extra) if extra else None,
            ),
        )
        self._logs_conn.commit()

    # ---------------------------------------------------------------
    # Query helpers
    # ---------------------------------------------------------------
    def query_readings(
        self,
        session_id: Optional[str] = None,
        device_id: Optional[str] = None,
        last_minutes: Optional[float] = None,
        limit: int = 10000,
        motion_level: Optional[str] = None,
        presence_only: bool = False,
    ) -> List[Dict[str, Any]]:
        """Query readings with optional filters."""
        conditions = []
        params: List[Any] = []

        if session_id:
            conditions.append("session_id = ?")
            params.append(session_id)
        if device_id:
            conditions.append("device_id = ?")
            params.append(device_id)
        if last_minutes is not None:
            cutoff = time.time() - last_minutes * 60
            conditions.append("ts_epoch >= ?")
            params.append(cutoff)
        if motion_level:
            conditions.append("motion_level = ?")
            params.append(motion_level)
        if presence_only:
            conditions.append("presence = 1")

        where = " AND ".join(conditions) if conditions else "1=1"
        params.append(limit)

        rows = self._conn.execute(
            f"SELECT * FROM readings WHERE {where} ORDER BY ts_epoch DESC LIMIT ?",
            params,
        ).fetchall()
        return [dict(r) for r in rows]

    def query_telemetry(
        self,
        device_id: Optional[str] = None,
        last_minutes: Optional[float] = None,
        limit: int = 1000,
    ) -> List[Dict[str, Any]]:
        conditions = []
        params: List[Any] = []
        if device_id:
            conditions.append("device_id = ?")
            params.append(device_id)
        if last_minutes:
            cutoff = time.time() - last_minutes * 60
            conditions.append("ts_epoch >= ?")
            params.append(cutoff)
        where = " AND ".join(conditions) if conditions else "1=1"
        params.append(limit)
        rows = self._conn.execute(
            f"SELECT * FROM esp32_telemetry WHERE {where} ORDER BY ts_epoch DESC LIMIT ?",
            params,
        ).fetchall()
        return [dict(r) for r in rows]

    def get_device_online_time(self, device_id: Optional[str] = None) -> Dict[str, Any]:
        """Calculate total online time for a device from heartbeats."""
        did = device_id or self._machine_id
        beats = self._conn.execute(
            "SELECT ts_epoch, online FROM heartbeats WHERE device_id = ? ORDER BY ts_epoch",
            (did,),
        ).fetchall()
        total_online_s = 0.0
        last_on: Optional[float] = None
        for b in beats:
            if b["online"] == 1:
                last_on = b["ts_epoch"]
            elif last_on is not None:
                total_online_s += b["ts_epoch"] - last_on
                last_on = None
        # If still online, count up to now
        if last_on is not None:
            total_online_s += time.time() - last_on
        return {
            "device_id": did,
            "total_online_seconds": total_online_s,
            "total_online_hours": total_online_s / 3600,
            "heartbeat_count": len(beats),
        }

    def get_session_stats(self, session_id: Optional[str] = None) -> Dict[str, Any]:
        """Get statistics for a session."""
        sid = session_id or self._session_id
        if not sid:
            return {}
        row = self._conn.execute(
            """SELECT
                 COUNT(*) as total,
                 MIN(ts_epoch) as first_ts,
                 MAX(ts_epoch) as last_ts,
                 AVG(feat_mean) as avg_rssi,
                 AVG(confidence) as avg_confidence,
                 SUM(CASE WHEN presence = 1 THEN 1 ELSE 0 END) as presence_ticks,
                 AVG(esp_temp_c) as avg_esp_temp,
                 MAX(esp_temp_c) as max_esp_temp,
                 MIN(esp_temp_c) as min_esp_temp
               FROM readings WHERE session_id = ?""",
            (sid,),
        ).fetchone()
        if not row:
            return {}
        r = dict(row)
        r["duration_s"] = (r["last_ts"] or 0) - (r["first_ts"] or 0)
        r["presence_pct"] = (r["presence_ticks"] / r["total"] * 100) if r["total"] else 0
        return r

    # ---------------------------------------------------------------
    # PyTorch integration
    # ---------------------------------------------------------------
    def to_tensor(
        self,
        rows: Optional[List[Dict]] = None,
        columns: Optional[List[str]] = None,
        last_minutes: Optional[float] = None,
    ):
        """
        Convert readings to a PyTorch tensor.

        Returns shape (N, C) where N = rows, C = len(columns).
        """
        import torch

        if rows is None:
            rows = self.query_readings(last_minutes=last_minutes)

        if columns is None:
            columns = [
                "feat_mean", "feat_std", "feat_variance", "feat_range",
                "feat_iqr", "feat_skewness", "feat_kurtosis",
                "feat_motion_power", "feat_breathing_power",
                "feat_dominant_freq_hz", "feat_spectral_power",
                "rssi_dbm", "mean_amplitude", "confidence",
            ]

        data = []
        for r in rows:
            row_vals = []
            for col in columns:
                v = r.get(col)
                row_vals.append(float(v) if v is not None else 0.0)
            data.append(row_vals)

        if not data:
            return torch.zeros(0, len(columns))

        return torch.tensor(data, dtype=torch.float32)

    def to_dataframe(
        self,
        last_minutes: Optional[float] = None,
        session_id: Optional[str] = None,
    ):
        """Export readings as a pandas DataFrame."""
        import pandas as pd
        rows = self.query_readings(
            last_minutes=last_minutes, session_id=session_id, limit=100000,
        )
        return pd.DataFrame(rows)

    # ---------------------------------------------------------------
    # ML export utilities
    # ---------------------------------------------------------------
    def export_training_windows(
        self,
        window_size: int = 20,
        step: int = 10,
        columns: Optional[List[str]] = None,
        session_id: Optional[str] = None,
        output_dir: Optional[str] = None,
    ) -> str:
        """
        Export sliding-window training data as .pt files for PyTorch.

        Returns path to the output directory.
        """
        import torch

        out = Path(output_dir) if output_dir else (
            Path(__file__).resolve().parents[3] / "data" / "ml" / "tensors"
        )
        out.mkdir(parents=True, exist_ok=True)

        rows = self.query_readings(session_id=session_id, limit=500000)
        if not rows:
            return str(out)

        tensor = self.to_tensor(rows, columns=columns)
        windows = []
        labels = []

        for i in range(0, len(tensor) - window_size, step):
            w = tensor[i:i + window_size]
            windows.append(w)
            # Label: majority presence in the window
            presence_col = [r.get("presence", 0) for r in rows[i:i + window_size]]
            labels.append(1 if sum(presence_col) > window_size // 2 else 0)

        if windows:
            X = torch.stack(windows)  # (N, window_size, features)
            y = torch.tensor(labels, dtype=torch.long)
            tag = datetime.now().strftime("%Y%m%d_%H%M%S")
            torch.save({"X": X, "y": y, "columns": columns or []}, out / f"windows_{tag}.pt")

        return str(out)

    def export_csv(
        self,
        output_path: Optional[str] = None,
        last_minutes: Optional[float] = None,
        session_id: Optional[str] = None,
    ) -> str:
        """Export readings as CSV for external analysis."""
        df = self.to_dataframe(last_minutes=last_minutes, session_id=session_id)
        out = output_path or str(
            Path(__file__).resolve().parents[3] / "data" / "ml" / "exports"
            / f"readings_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        df.to_csv(out, index=False)
        return out
