"""
SQLite logging handler — writes structured log records to logs/ruview_logs.sqlite3.

Plugs into Python's standard logging framework.  Every log record gets
device_id + session_id context automatically.

Usage::

    import logging
    from v1.src.database.db_log_handler import SensingDBLogHandler

    handler = SensingDBLogHandler()
    logging.getLogger().addHandler(handler)
"""

from __future__ import annotations

import json
import logging
import sqlite3
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional


_LOGS_DB_PATH = Path(__file__).resolve().parents[3] / "logs" / "ruview_logs.sqlite3"

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
"""


class SensingDBLogHandler(logging.Handler):
    """Python logging handler that writes to SQLite."""

    def __init__(
        self,
        db_path: Optional[str] = None,
        device_id: Optional[str] = None,
        level: int = logging.INFO,
    ) -> None:
        super().__init__(level=level)
        self._path = Path(db_path) if db_path else _LOGS_DB_PATH
        self._path.parent.mkdir(parents=True, exist_ok=True)
        self._device_id = device_id
        self._conn: Optional[sqlite3.Connection] = None
        self._connect()

    def _connect(self) -> None:
        self._conn = sqlite3.connect(str(self._path), timeout=10, check_same_thread=False)
        self._conn.execute("PRAGMA journal_mode=WAL")
        self._conn.execute("PRAGMA synchronous=NORMAL")
        self._conn.executescript(_LOGS_DDL)

    def emit(self, record: logging.LogRecord) -> None:
        try:
            now = datetime.now(timezone.utc)
            extra = {}
            for key in ("request_id", "ip", "user_agent", "path"):
                val = getattr(record, key, None)
                if val is not None:
                    extra[key] = val

            self._conn.execute(
                """INSERT INTO app_logs
                   (ts_utc, ts_epoch, level, logger, message,
                    module, func, line, device_id, session_id, extra)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    now.isoformat(),
                    time.time(),
                    record.levelname,
                    record.name,
                    self.format(record),
                    record.module,
                    record.funcName,
                    record.lineno,
                    self._device_id,
                    getattr(record, "session_id", None),
                    json.dumps(extra) if extra else None,
                ),
            )
            self._conn.commit()
        except Exception:
            self.handleError(record)

    def close(self) -> None:
        if self._conn:
            self._conn.close()
            self._conn = None
        super().close()
