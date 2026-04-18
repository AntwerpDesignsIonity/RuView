# ADR-082: REST provisioning + WebSocket registry events

- **Status:** Proposed
- **Date:** 2026-04-18
- **Deciders:** Stack maintainers
- **Related:** [ADR-080](ADR-080-mac-based-node-identity.md), `firmware/esp32-csi-node/identify-node.py`

## Context

Provisioning a new ESP32 today requires SSH'ing to the host and running
`identify-node.py` or `.ionity/ionity.sh provision`. There is no way for the
web UI to:

1. Trigger a flash/provision from a browser.
2. Observe in real time when a new node appears in `nodes.yaml`.

## Decision

Two thin layers, both stubs in this ADR:

### A. REST provisioning endpoints

```
POST /api/v1/nodes/{mac}/provision
  body: { ssid, password, target_ip }
  → 202 Accepted on enqueue, 501 in v1 (returns "not implemented")

POST /api/v1/nodes/{mac}/flash
  body: { firmware_variant: "n16r8" | "lcd_1_47" | "touch_lcd_2" }
  → 202 Accepted on enqueue, 501 in v1
```

Implementation will shell out to `identify-node.py` or `pio run -t upload`
in a background task with a job queue. Auth: bearer token in header (set
in `.env.local` as `IONITY_PROVISION_TOKEN`).

### B. WebSocket registry events

Add a typed event to the existing `/ws/sensing` channel:

```json
{ "type": "registry-update",
  "added":   [{ "mac": "...", "node_id": 7 }],
  "removed": [],
  "ts_ms":   1747000000000 }
```

Server fires on file change of `firmware/esp32-csi-node/nodes.yaml` (use
`notify` crate). UI updates the Mesh tile and registry table without polling.

## Consequences

**Positive:**
- Fully browser-driven provisioning workflow becomes possible.
- Mesh tile updates instantly when a new node joins.

**Negative:**
- Provisioning over the network is a privileged operation — auth is
  mandatory or the system can be bricked remotely.
- File-watch on the registry adds one more long-lived task.

## Implementation status

- REST stubs: not started. Should return `501 Not Implemented` initially
  with a `Link` header pointing here.
- WS event: not started.
- Auth scheme: deferred — implement only when stubs are promoted.
