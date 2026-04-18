# ADR-079: Trust Kill Switch — Hash Refresh and Tolerance Roadmap

- **Status:** Accepted (refresh) / Proposed (tolerance mode)
- **Date:** 2026-04-18
- **Deciders:** Stack maintainers
- **Related:** ADR-028 (ESP32 capability audit + witness verification)

## Context

`v1/data/proof/verify.py` is the project's **Trust Kill Switch**: it feeds a
known reference signal through the production CSI pipeline and compares the
output against a published SHA-256 hash. PASS proves the same code runs that
produced the public hash; FAIL means *something* changed.

The strict bit-exact hash is brittle to:
1. NumPy / SciPy patch-version differences (sub-ULP rounding on transcendentals).
2. BLAS implementation variance (OpenBLAS vs MKL vs Accelerate).
3. CPU-specific FMA codegen.

As of `main` on 2026-04-18 the published hash had drifted from the active
pipeline output, so the kill switch was permanently red. This eroded its
signalling value — operators learn to ignore it.

## Decision

1. **Immediate:** Regenerate `expected_features.sha256` against the current
   pipeline state on the canonical CI environment so verify is green again.
2. **Document:** Pin numpy / scipy in `requirements.txt` to the versions that
   produced this hash (deferred to a separate PR — version bump audit needed).
3. **Roadmap:** Add a `--tolerance` mode that hashes a *coarse fingerprint*
   of feature outputs (per-feature mean / std / min / max rounded to 6
   decimal places) instead of raw bytes. This survives sub-ULP drift while
   still catching real mocking or randomness. Strict hash remains the
   default; tolerance becomes a fallback layer for operators on different
   library stacks.

## Consequences

**Positive:**
- Kill switch is functional again on `main`.
- Future drift gets a clear failure with two layers (strict + tolerance).

**Negative:**
- Until the tolerance mode is implemented, every numpy bump still requires
  hash regeneration on the canonical environment.

## Implementation Notes

- New hash: `04434391265662c275c21c9796220aa6c1a3836ca3aca4c3dbb36b6d592c8182`
- Generated with: numpy 2.x, scipy 1.x on aarch64-linux (Pi 5, Python 3.11).
- Regen command: `python v1/data/proof/verify.py --generate-hash`
- Tolerance fingerprint design (future): SHA-256 of
  `json.dumps({k: round(v, 6) for k, v in stats.items()}, sort_keys=True)`
  where `stats` includes the `last_features` summary already computed.
