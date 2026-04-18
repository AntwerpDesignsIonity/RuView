# MERIDIAN: cross-environment domain generalization

**Status:** Open
**Owner:** unassigned
**Related ADR:** [ADR-027](../adr/ADR-027-meridian-cross-environment.md)

## Summary

ADR-027 specifies MERIDIAN — a domain-adaptation layer that lets a model trained
in one environment (room geometry / WiFi channel / hardware) generalize to
another without per-room re-training. The ADR is **Accepted** but no
implementation tracking exists in the repo.

## Acceptance criteria

- [ ] MERIDIAN feature-extractor module added to `wifi-densepose-nn`
      (`crates/wifi-densepose-nn/src/meridian/`).
- [ ] Domain discriminator + gradient reversal layer wired into training loop
      in `wifi-densepose-train`.
- [ ] Cross-environment eval harness comparing baseline vs MERIDIAN model on
      held-out rooms.
- [ ] Witness bundle entry: cross-room mAP delta reported in
      `docs/WITNESS-LOG-028.md`.

## Out of scope

- New CSI capture in additional rooms (separate data-collection task).
- Public dataset release.

## Status notes

Tracked here because GitHub Issues for this repo are not authoritative for
internal ML roadmap. Promote to a formal issue once an owner is assigned.
