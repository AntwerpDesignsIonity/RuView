# ADR-083: CSI Preprocessing Accuracy Roadmap

| Field | Value |
|---|---|
| Status | Proposed |
| Date | 2026-04-18 |
| Decision-makers | rUv (ruvnet), Ionity Global engineering |
| Related ADRs | ADR-014 (SOTA signal processing), ADR-022 (multi-BSSID), ADR-029 (RuvSense multistatic), ADR-031 (sensing-first RF mode), ADR-081 (per-node telemetry) |
| Supersedes | — |

## Context

The Rust signal crate ships five preprocessing modules that are **already
implemented and unit-tested** but are not yet wired into the live UDP→state
pipeline inside `wifi-densepose-sensing-server`. They are:

| Module | Source | Purpose |
|---|---|---|
| `hampel` | `wifi-densepose-signal/src/hampel.rs` | Outlier suppression (sliding median + MAD) on amplitude |
| `phase_sanitizer` | `…/src/phase_sanitizer.rs` | Per-frame phase unwrap, STO/SFO removal |
| `adaptive_denoise` | `…/src/adaptive_denoise.rs` | SNR-aware adaptive denoising |
| `csi_ratio` | `…/src/csi_ratio.rs` | Conjugate-multiply across antennas — cancels common CFO |
| `subcarrier_selection` | `…/src/subcarrier_selection.rs` | Variance/sensitivity-driven subcarrier subset |

Web research confirms each technique is standard in CSI literature:
- **Hampel**: standard outlier removal in WiFi CSI denoising stacks.
- **Phase sanitization**: required for any cross-frame phase-based feature
  (FALCON, Widar, WiPose all do this).
- **CSI ratio**: state of the art for hardware-noise cancellation since
  IndoTrack (2017) and remains the first step in modern systems.
- **Adaptive denoising**: replaces fixed-cutoff filters that destroy
  fast biomechanical signals (chest wall motion, gait swing).
- **Subcarrier selection**: 30-114 raw subcarriers contain redundant
  noise; modern pipelines pick the top-k most informative.

Wiring them today requires:
1. Mutating the per-`NodeState` CSI buffer between UDP parse and
   `MultistaticFuser` call.
2. Re-baselining the deterministic verify hash
   (`v1/data/proof/expected_features.sha256`).
3. A way to A/B compare with stages off vs on.

## Decision

Roll out the five preprocessing stages **behind individual environment
flags**, default-off, in a single deterministic order:

```
hampel  →  phase_sanitizer  →  adaptive_denoise  →  csi_ratio  →  subcarrier_selection
```

| Stage | Env flag |
|---|---|
| Hampel filter | `IONITY_HAMPEL=1` |
| Phase sanitizer | `IONITY_PHASE_SANITIZER=1` |
| Adaptive denoiser | `IONITY_ADAPTIVE_DENOISE=1` |
| CSI ratio | `IONITY_CSI_RATIO=1` |
| Subcarrier selection | `IONITY_SUBCARRIER_SELECTION=1` |

Expose the live state at `GET /api/v1/preprocessing/stats` so the UI can
display which stages are active without restarting the server.

## Status by phase

### Phase D.1 — *complete in this PR*

- ✅ `GET /api/v1/preprocessing/stats` endpoint live, returns the
  configured-vs-active matrix.
- ✅ Endpoint is included in `ui/diagnostics.html` probe list.
- ✅ Endpoint is included in `ui/tests/ui-smoke.html` probe runner.
- ✅ ADR documented (this file).
- ✅ All flags default-off → verify hash unchanged → CI green.

### Phase D.2 — *next PR*

- ⬜ Add `PreprocConfig` to `NodeState` carrying compiled configs for each
  enabled stage.
- ⬜ Insert per-stage hooks in the per-node UDP handler **after** parse,
  **before** `MultistaticFuser::ingest`.
- ⬜ Per-stage rolling counters (frames in, outliers removed, NaNs
  rejected, subcarriers retained) reported in the stats endpoint.
- ⬜ New `ui/preprocessing.html` page with toggle UI + sparkline of
  counters.

### Phase D.3 — *after acceptance*

- ⬜ Re-run `python v1/data/proof/verify.py --generate-hash` with the
  recommended preprocessing profile (likely
  `IONITY_HAMPEL=1 IONITY_CSI_RATIO=1 IONITY_PHASE_SANITIZER=1`).
- ⬜ Update `v1/data/proof/expected_features.sha256`.
- ⬜ Promote the three accepted defaults from env-flag to compiled-in.
- ⬜ Witness bundle regen (ADR-028 contract).

## Consequences

**Positive**
- Improvements ship in increments, each independently bisectable.
- Verify hash stays green until we *intentionally* re-baseline.
- Operators can A/B in production by flipping one env var.
- The UI surfaces what is actually running, eliminating the "did it
  apply?" guesswork that has bitten this stack before (see
  `docs/issues/known-test-failures.md`).

**Negative**
- The five modules remain unused in the runtime path until D.2 lands,
  giving false confidence to a casual reader.
- Adds a new env-flag namespace (`IONITY_*`) that must be documented in
  `example.env` once D.2 ships.

**Mitigations**
- This ADR is the single source of truth for the rollout schedule.
- The `/api/v1/preprocessing/stats` endpoint is *visibly disabled* by
  default — no silent-win footgun.

## Verification

Today, against a running server:

```bash
curl -s http://localhost:3000/api/v1/preprocessing/stats | jq .
# → { stages: [...], note: "..." }
```

Toggle a stage:

```bash
IONITY_HAMPEL=1 ./.ionity/ionity.sh run --yes
curl -s http://localhost:3000/api/v1/preprocessing/stats | jq '.stages[] | select(.enabled)'
# → { name: "hampel", enabled: true, ... }
```

The ui-smoke probe runner
(`http://<host>:3000/ui/tests/ui-smoke.html`) will list the endpoint as
`PASS` once the server is up.

## References

- `rust-port/wifi-densepose-rs/crates/wifi-densepose-signal/src/{hampel,phase_sanitizer,adaptive_denoise,csi_ratio,subcarrier_selection}.rs`
- `rust-port/wifi-densepose-rs/crates/wifi-densepose-sensing-server/src/main.rs` — `preprocessing_stats_endpoint`
- ADR-014 — SOTA signal processing acceptance criteria
- ADR-029 — RuvSense multistatic sensing mode (existing 13 wired modules)
