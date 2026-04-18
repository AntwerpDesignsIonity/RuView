# ADR-080: MAC address as canonical node identity

- **Status:** Accepted
- **Date:** 2026-04-18
- **Deciders:** Stack maintainers
- **Related:** [ADR-028](ADR-028-esp32-capability-audit.md), `firmware/esp32-csi-node/nodes.yaml`

## Context

Node identity is currently a small integer (1..N) chosen at provisioning time
and stored in NVS on the ESP32. Operationally, the LAN-stable identifier we
actually have is the WiFi STA MAC address (assigned to the chip at fab and
visible in every UDP CSI frame and in `arp` tables).

Issues with integer-only identity:

1. Wiping/re-flashing a board loses its identity unless `identify-node.py`
   is re-run with the right `--node-id`.
2. Two physically different boards can collide on the same ID after a
   re-provisioning mistake.
3. Mesh recovery after a power cycle requires the ID to be re-asserted by
   the firmware before any frames are credited correctly.
4. Cross-installation portability (moving a node to a different deployment)
   is awkward — the ID is local but the MAC is global.

## Decision

Treat the **WiFi STA MAC** as the canonical, immutable node identity.

- `firmware/esp32-csi-node/nodes.yaml` continues to assign a small integer
  *slot number* per MAC, used only for human-readable indexing and UI display
  ordering.
- The Rust server keys all per-node state by MAC (already stored in
  `node_states`). The slot integer becomes a lookup, not a primary key.
- New nodes auto-register: an unknown MAC enumerates with the next free slot
  number (already implemented in `identify-node.py --node-id`).
- `nodes.yaml` schema gains an explicit `primary_id: mac` marker (informational).

## Consequences

**Positive:**
- Re-flashing a board preserves identity (MAC is in silicon).
- ID collisions become impossible — two boards always have different MACs.
- The registry survives `git clone` to another deployment unchanged.

**Negative:**
- Operators must remember to update `nodes.yaml` if a board is physically
  replaced (the new board has a new MAC).
- Slot numbers in the UI are no longer guaranteed to match historical logs
  if a board is replaced and its slot is re-assigned.

## Migration

No code changes required — the in-memory representation is already MAC-keyed.
This ADR formalises the existing behaviour and locks in the contract.
Future endpoints (e.g. `/api/v1/nodes/{mac}/telemetry`) MUST accept MAC, not
slot integer.
