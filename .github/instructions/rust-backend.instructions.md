---
name: 'Rust Backend'
description: 'Conventions for the Rust workspace: 18 crates, Axum server, signal processing, NN inference, and ESP32 hardware integration.'
applyTo: 'rust-port/**/*.rs'
---

# Rust Backend Conventions

## Workspace Structure

The Rust workspace at `rust-port/wifi-densepose-rs/` contains 18 crates. Key ones:

- **wifi-densepose-sensing-server** — Main Axum server (HTTP :3000, WS :3001, UDP :5005). This is the production entry point.
- **wifi-densepose-signal** — Signal processing + RuvSense multistatic sensing (14 modules in `src/ruvsense/`).
- **wifi-densepose-core** — Core types, traits, error types, CSI frame primitives.
- **wifi-densepose-nn** — Neural network inference (ONNX, PyTorch, Candle backends).
- **wifi-densepose-vitals** — ESP32 CSI-grade vital sign extraction.
- **wifi-densepose-mat** — Mass Casualty Assessment Tool.
- **wifi-densepose-hardware** — ESP32 aggregator, TDM protocol.

## Build Commands

```bash
# Server only (production build):
cargo build -p wifi-densepose-sensing-server --release --no-default-features

# Full workspace test (must be 0 failed):
cargo test --workspace --no-default-features

# Single crate check (fast):
cargo check -p <crate-name> --no-default-features

# Benchmarks:
cargo bench --package wifi-densepose-signal
```

Always use `--no-default-features` for builds and tests — default features pull in optional GPU/training deps.

## Code Style

- `cargo fmt` before commit — no exceptions.
- `cargo clippy` must pass with zero warnings on changed crates.
- Files under 500 lines. Split into modules when approaching limit.
- All public APIs must have typed signatures — no `impl Into<String>` on boundaries.
- Use `thiserror` for error types, not manual `impl Display`.
- Prefer `anyhow::Result` only in binary crates (cli, sensing-server). Library crates use typed errors.

## Crate Boundaries

- Library crates must NOT depend on `tokio` runtime features directly — accept async traits.
- `wifi-densepose-core` has zero internal dependencies. Never add deps to core that pull in other workspace crates.
- The sensing-server crate orchestrates; signal/nn/vitals crates do the work. Keep business logic out of the server.

## Axum Server Patterns

- Handlers return `Result<Json<T>, AppError>` — never panic in handlers.
- Use `axum::extract::State` for shared state, wrapped in `Arc`.
- WebSocket handlers use the `ws/` route prefix.
- Health endpoint at `/health` returns JSON `{ status, source, clients, tick }`.

## Testing

- TDD London School: mock dependencies at crate boundaries.
- Use `#[cfg(test)]` modules in the same file for unit tests.
- Integration tests go in `tests/` directory of each crate.
- Run tests after every code change: `cargo test -p <changed-crate> --no-default-features`.

## Dependency Order (for publishing)

1. core → 2. vitals, wifiscan, hardware, config, db → 3. signal (depends on core) → 4. nn, ruvector → 5. train, mat → 6. api, wasm → 7. sensing-server, cli
