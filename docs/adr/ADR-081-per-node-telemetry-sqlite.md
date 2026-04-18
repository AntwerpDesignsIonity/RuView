# ADR-081: Per-node telemetry persistence (SQLite ring)

- **Status:** Proposed
- **Date:** 2026-04-18
- **Deciders:** Stack maintainers
- **Related:** [ADR-080](ADR-080-mac-based-node-identity.md), `logs/aedi_logs.sqlite3`

## Context

Per-node CSI metrics (RSSI, frame rate, motion level, person count) are
currently in-memory only inside `wifi-densepose-sensing-server`. After a
restart everything is gone, which:

- Breaks long-term drift analysis (RuvSense longitudinal modules).
- Hides intermittent node failures (we only see the live state).
- Forces the UI to render a flat-line graph from a 5-minute rolling buffer.

The repo already ships `logs/aedi_logs.sqlite3` for application logs, so
SQLite is an established dependency.

## Decision

Add a per-node telemetry table to a sibling SQLite file
(`logs/aedi_telemetry.sqlite3`) with a fixed-size ring buffer per node.

### Schema

```sql
CREATE TABLE IF NOT EXISTS node_telemetry (
    mac          TEXT NOT NULL,           -- canonical node identity (ADR-080)
    ts_ms        INTEGER NOT NULL,        -- unix epoch ms
    rssi_dbm     REAL,
    frame_rate   REAL,
    motion_level TEXT,                    -- 'low' | 'med' | 'high'
    person_count INTEGER,
    PRIMARY KEY (mac, ts_ms)
);
CREATE INDEX IF NOT EXISTS idx_node_telemetry_ts ON node_telemetry (ts_ms);
```

### Retention

- Ring trimmed nightly: keep last 7 days at 1 Hz, last 90 days downsampled
  to 1/min via a SQL `INSERT INTO ... SELECT AVG(...) GROUP BY` rollup.
- Hard cap: 50 MB on disk; oldest rows evicted when exceeded.

### Write path

- Server writes one batch INSERT per node per second from the existing
  `node_states` map (no new collection logic).
- Use SQLite `WAL` mode to avoid blocking the live tick loop.

### Read path

- New endpoint: `GET /api/v1/nodes/{mac}/telemetry?from=<ms>&to=<ms>`
  returns JSON array of rows, capped at 5000 rows per call.

## Consequences

**Positive:**
- Node-failure forensics: "node 5 dropped offline at 03:42 last Tuesday"
  becomes a one-query answer.
- UI can plot real historical graphs.
- Trivial backup — just copy the .sqlite3 file.

**Negative:**
- Disk wear (~ 50 MB rotation per long-running install).
- One more file to back up alongside `nodes.yaml`.

## Implementation status

Not started. Design only — track in `docs/issues/telemetry-persistence.md`
once an owner is assigned.
