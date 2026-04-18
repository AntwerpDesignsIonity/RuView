# Known test failures and proof drift

**Last reviewed:** 2026-04-18

## Summary

Two known issues that the `copilot-instructions.md` "0 failed" claim glossed
over. Both are now tracked here so the docs match reality.

| ID  | Component                                          | Severity | Status |
|-----|----------------------------------------------------|----------|--------|
| T-1 | `wifi-densepose-signal::ruvsense::field_model::test_estimate_occupancy_noise_only` | low | known-flake |
| T-2 | `v1/data/proof/verify.py` hash drift on numpy/scipy bumps | medium | mitigated by ADR-079 |

---

## T-1 — `test_estimate_occupancy_noise_only` (known flake)

**Symptom:** Test asserts that estimated occupancy on pure-noise input is < a
fixed threshold; it occasionally exceeds the threshold by < 5% on certain
ARM/x86 BLAS combos.

**Root cause:** SVD of a noise-only matrix produces eigenvalues whose ratio
sits exactly at the threshold; floating-point variance flips the verdict.

**Mitigation options:**

1. Widen the threshold by 10% (cheapest, accept some sensitivity loss).
2. Replace the single-shot assertion with a Monte-Carlo run (1000 trials,
   require P(false-positive) < 5%).
3. Mark `#[ignore]` and run nightly only.

**Owner:** unassigned. Until fixed, treat workspace test pass count as
"1,463 / 1,464 ± 1 flake" rather than absolute zero.

## T-2 — Trust Kill Switch hash drift

**Symptom:** `python v1/data/proof/verify.py` produces a different SHA-256
than the published `expected_features.sha256` after a numpy or scipy
patch-version bump.

**Root cause:** Sub-ULP rounding in transcendental functions (FFT, exp, log)
varies between library versions.

**Mitigation:** ADR-079 — regenerate the hash on the canonical CI
environment and add a `--tolerance` mode (TODO) that compares feature
fingerprints rather than raw bytes.

**Status:** Hash regenerated 2026-04-18. Verify is GREEN on `main`.

---

## How to use this doc

Before merging a PR that changes signal processing, NN inference, or
`v1/src/processing/`:

1. Run `cargo test --workspace --no-default-features` and accept up to 1
   known flake (T-1) — investigate any *new* failure.
2. Run `python v1/data/proof/verify.py`. If FAIL, decide:
   - Is this an intentional pipeline change? → regenerate hash, document
     the change in the PR.
   - Is this an environment drift? → see ADR-079.
3. If you fix T-1 or fully resolve T-2 (with a `--tolerance` mode), remove
   the entry here.
