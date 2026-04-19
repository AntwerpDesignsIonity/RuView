#!/usr/bin/env python3
"""
events-to-sqlite.py — tail data/events.jsonl into logs/aedi_logs.sqlite3.

Runs as a long-lived sidecar. Creates schema on first run, keeps a byte
offset in a sidecar state file so it survives restarts without re-importing.
"""
from __future__ import annotations

import argparse
import json
import signal
import sqlite3
import sys
import time
from pathlib import Path


SCHEMA = """
CREATE TABLE IF NOT EXISTS events (
    id      INTEGER PRIMARY KEY AUTOINCREMENT,
    ts      TEXT    NOT NULL,
    kind    TEXT    NOT NULL,
    payload TEXT    NOT NULL,
    created INTEGER NOT NULL DEFAULT (strftime('%s','now'))
);
CREATE INDEX IF NOT EXISTS idx_events_kind ON events(kind);
CREATE INDEX IF NOT EXISTS idx_events_ts   ON events(ts);
"""


def open_db(path: Path) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path), isolation_level=None)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    conn.executescript(SCHEMA)
    return conn


def read_offset(state: Path) -> int:
    try:
        return int(state.read_text().strip() or 0)
    except (FileNotFoundError, ValueError):
        return 0


def write_offset(state: Path, offset: int) -> None:
    state.parent.mkdir(parents=True, exist_ok=True)
    state.write_text(str(offset))


def run(events_path: Path, db_path: Path, state_path: Path, poll_s: float) -> None:
    conn = open_db(db_path)
    offset = read_offset(state_path)
    stop = {"v": False}

    def _handle_signal(signum, _frame):
        print(f"[events-to-sqlite] stopping on signal {signum}", flush=True)
        stop["v"] = True

    signal.signal(signal.SIGINT, _handle_signal)
    signal.signal(signal.SIGTERM, _handle_signal)

    print(f"[events-to-sqlite] tailing {events_path} -> {db_path} from offset {offset}", flush=True)

    while not stop["v"]:
        if not events_path.exists():
            time.sleep(poll_s)
            continue

        size = events_path.stat().st_size
        # File truncated/rotated — restart from 0.
        if size < offset:
            print(f"[events-to-sqlite] file shrank ({offset} -> {size}), restarting at 0", flush=True)
            offset = 0

        if size == offset:
            time.sleep(poll_s)
            continue

        with events_path.open("rb") as fh:
            fh.seek(offset)
            chunk = fh.read()

        text = chunk.decode("utf-8", errors="replace")
        lines = text.split("\n")
        # Last element is either "" (if chunk ended on \n) or a partial line.
        partial = lines.pop()
        consumed = len(chunk) - len(partial.encode("utf-8"))

        rows = []
        for line in lines:
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
            except json.JSONDecodeError:
                continue
            rows.append(
                (
                    obj.get("ts", ""),
                    obj.get("kind", ""),
                    json.dumps(obj.get("payload", {}), separators=(",", ":")),
                )
            )

        if rows:
            conn.executemany(
                "INSERT INTO events (ts, kind, payload) VALUES (?, ?, ?)",
                rows,
            )
            print(f"[events-to-sqlite] inserted {len(rows)} rows (total offset now {offset + consumed})", flush=True)

        offset += consumed
        write_offset(state_path, offset)

    conn.close()
    print("[events-to-sqlite] stopped cleanly", flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Tail events.jsonl into SQLite.")
    parser.add_argument("--events", default="data/events.jsonl", help="Source JSONL file")
    parser.add_argument("--db", default="logs/aedi_logs.sqlite3", help="Destination SQLite DB")
    parser.add_argument(
        "--state",
        default="logs/events-to-sqlite.offset",
        help="File to persist byte offset",
    )
    parser.add_argument("--poll", type=float, default=1.0, help="Poll interval (seconds)")
    args = parser.parse_args()

    run(Path(args.events), Path(args.db), Path(args.state), args.poll)
    return 0


if __name__ == "__main__":
    sys.exit(main())
