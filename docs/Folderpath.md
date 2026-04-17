# AEDI-S — Complete File System & Folder Path Reference

> **Version:** 0.6.0 | **Updated:** 2026-04-06 | **Repository:** [ionity.today](https://www.ionity.today)

This document maps every directory and key file in the AEDI-S monorepo. Use it to navigate the codebase, understand boundaries between Rust/Python/firmware/UI layers, and locate ML pipelines, edge computing infrastructure, and logging systems.

---

## Table of Contents

- [Root Layout](#root-layout)
- [Rust Workspace (15 Crates)](#rust-workspace)
- [RuvSense Signal Processing Modules](#ruvsense-modules)
- [Cross-Viewpoint Fusion](#cross-viewpoint-fusion)
- [Neural Network & Training Crates](#neural-network--training)
- [Mass Casualty Assessment Tool (MAT)](#mass-casualty-assessment-tool)
- [Sensing Server](#sensing-server)
- [Python v1 Codebase](#python-v1-codebase)
- [ML Models & Data](#ml-models--data)
- [Firmware (ESP32)](#firmware-esp32)
- [Web & Mobile UI](#web--mobile-ui)
- [Edge Computing Infrastructure](#edge-computing-infrastructure)
- [Monitoring & Logging](#monitoring--logging)
- [Docker & Deployment](#docker--deployment)
- [Scripts & Utilities](#scripts--utilities)
- [Documentation & ADRs](#documentation--adrs)
- [CI/CD & GitHub Workflows](#cicd--github-workflows)
- [Configuration & Tooling](#configuration--tooling)
- [Examples & Tutorials](#examples--tutorials)

---

## Root Layout

```
AEDI-S/
│
├── CLAUDE.md                     # Project conventions & architecture map
├── README.md                     # Public-facing documentation & wiki
├── CHANGELOG.md                  # Version history
├── LICENSE                       # MIT
├── Makefile                      # Build orchestration
├── pyproject.toml                # Python project metadata (v1.2.0)
├── requirements.txt              # Python dependencies (PyTorch, FastAPI, etc.)
├── example.env                   # Environment template (secrets go here, not committed)
├── deploy.sh                     # One-command production deployment
├── install.sh                    # Development environment bootstrapper
├── nvs_config.csv                # ESP32 NVS configuration matrix
├── benchmark_baseline.json       # Performance regression baseline
├── verify                        # Verification binary (ADR-028 proof)
├── VIEW-AEDI.code-workspace      # VS Code multi-root workspace
│
├── rust-port/                    # ── PRIMARY: Rust workspace ──────────────
├── v1/                           # ── Python v1 codebase ──────────────────
├── firmware/                     # ── ESP32-S3 C firmware ──────────────────
├── ui/                           # ── Web + Mobile (React, Three.js, Expo) ─
├── data/                         # ── CSI datasets, trained models ─────────
├── scripts/                      # ── 44+ utility scripts ─────────────────
├── docs/                         # ── 78 ADRs, DDD models, research ───────
├── examples/                     # ── Medical, sleep, stress, environment ──
├── docker/                       # ── Multi-arch Docker images ────────────
├── monitoring/                   # ── Prometheus, Grafana, alerting ────────
├── logging/                      # ── Fluentd log aggregation ─────────────
├── logs/                         # ── Runtime log output ──────────────────
├── plans/                        # ── Roadmap, phase specs ────────────────
├── references/                   # ── Reference implementations ───────────
├── releases/                     # ── Built binaries (desktop, firmware) ──
├── assets/                       # ── Screenshots, demo zips ──────────────
├── vendor/                       # ── Git submodules (ruvector, etc.) ─────
│
├── .github/                      # CI/CD workflows (8 pipelines)
├── .claude/                      # Claude Code agents/skills/settings
├── .claude-flow/                 # Swarm coordination state & metrics
├── .ionity/                      # Ionity platform integration CLI
├── .swarm/                       # Agent swarm state
├── .vscode/                      # VS Code settings
└── .venv/                        # Python virtual environment (not committed)
```

---

## Rust Workspace

**Location:** `rust-port/wifi-densepose-rs/`
**Workspace version:** 0.3.0 | **Rust edition:** 2021 | **Resolver:** v2

16 crates (15 in workspace + 1 excluded `wasm-edge` for `no_std` target).

```
rust-port/wifi-densepose-rs/
├── Cargo.toml                        # Workspace manifest (15 members)
├── Cargo.lock                        # Deterministic dependency lock
│
├── crates/
│   │
│   │  ── CORE LAYER ────────────────────────────────────────────────
│   ├── wifi-densepose-core/          # Core types, traits, error types
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── error.rs              # Unified error enum
│   │       ├── traits.rs             # CsiProcessor, PoseEstimator, etc.
│   │       ├── types.rs              # CsiFrame, Subcarrier, Antenna
│   │       └── utils.rs              # Shared helpers
│   │
│   ├── wifi-densepose-config/        # Configuration management
│   │   └── src/lib.rs                # TOML/YAML/env config loading
│   │
│   ├── wifi-densepose-db/            # Database layer (Postgres, SQLite, Redis)
│   │   └── src/lib.rs                # SQLx async abstractions
│   │
│   │  ── SIGNAL PROCESSING ─────────────────────────────────────────
│   ├── wifi-densepose-signal/        # SOTA signal pipeline (depends: core)
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── bvp.rs                # Blood volume pulse extraction
│   │       ├── csi_processor.rs      # Main CSI pipeline
│   │       ├── csi_ratio.rs          # Ratio-based sanitization
│   │       ├── features.rs           # Feature vector extraction
│   │       ├── fresnel.rs            # RF propagation physics
│   │       ├── hampel.rs             # Hampel outlier filter
│   │       ├── hardware_norm.rs      # Hardware normalization
│   │       ├── motion.rs             # Motion index computation
│   │       ├── phase_sanitizer.rs    # Phase unwrapping & cleanup
│   │       ├── spectrogram.rs        # Time-frequency analysis
│   │       ├── subcarrier_selection.rs  # Intelligent subcarrier pruning
│   │       └── ruvsense/             # ← 16 SOTA sensing modules (see below)
│   │
│   ├── wifi-densepose-vitals/        # Vital sign extraction (ADR-021)
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── breathing.rs          # Respiratory rate (0.1-0.5 Hz BP)
│   │       ├── heartrate.rs          # HR + HRV (0.8-2.0 Hz BP)
│   │       ├── anomaly.rs            # Apnea / arrhythmia anomaly
│   │       ├── preprocessor.rs       # Signal conditioning
│   │       ├── store.rs              # Persistent vital store
│   │       └── types.rs              # VitalSample, VitalWindow
│   │
│   ├── wifi-densepose-wifiscan/      # Multi-BSSID WiFi scanning (ADR-022)
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── error.rs
│   │       ├── adapter/              # OS network adapter abstraction
│   │       ├── domain/               # Scan result domain models
│   │       ├── pipeline/             # Scan → filter → rank pipeline
│   │       └── port/                 # Platform-specific ports (Linux/Win/Mac)
│   │
│   │  ── ML / INFERENCE ────────────────────────────────────────────
│   ├── wifi-densepose-nn/            # Neural network inference
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── densepose.rs          # DensePose head architecture
│   │       ├── inference.rs          # Batch inference engine
│   │       ├── onnx.rs               # ONNX Runtime backend
│   │       ├── tensor.rs             # Tensor utilities
│   │       ├── translator.rs         # PyTorch ↔ Rust weight translation
│   │       └── error.rs
│   │
│   ├── wifi-densepose-train/         # Training pipeline (depends: signal, nn)
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── trainer.rs            # Main training loop
│   │       ├── model.rs              # WiFlow architecture (1.8M params)
│   │       ├── dataset.rs            # CompressedCsiBuffer, MM-Fi loader
│   │       ├── losses.rs             # Contrastive + MSE + physics loss
│   │       ├── eval.rs               # PCK, AUC, MPJPE evaluation
│   │       ├── metrics.rs            # DynamicPersonMatcher (mincut)
│   │       ├── aedis_metrics.rs     # AEDI-S-specific KPIs
│   │       ├── config.rs             # Training hyperparameters
│   │       ├── domain.rs             # Training domain model
│   │       ├── geometry.rs           # Skeleton constraint validation
│   │       ├── subcarrier.rs         # 114→56 sparse interpolation
│   │       ├── rapid_adapt.rs        # Online adaptation (EWC + LoRA)
│   │       ├── virtual_aug.rs        # Synthetic data augmentation
│   │       ├── proof.rs              # ADR-028 deterministic proof
│   │       └── bin/
│   │           ├── train.rs          # `cargo run --bin train`
│   │           └── verify_training.rs
│   │
│   ├── wifi-densepose-ruvector/      # RuVector v2.0.4 integration
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── viewpoint/            # ← Cross-viewpoint fusion (see below)
│   │       ├── signal/               # Signal bridge modules
│   │       │   ├── mod.rs
│   │       │   ├── bvp.rs            # BVP via ruvector-attention
│   │       │   ├── fresnel.rs        # Fresnel via ruvector-solver
│   │       │   ├── spectrogram.rs    # Spectro via ruvector-attn-mincut
│   │       │   └── subcarrier.rs     # Sparse interp via ruvector-solver
│   │       ├── mat/                  # MAT vital sign bridge
│   │       │   ├── mod.rs
│   │       │   ├── breathing.rs      # Breathing via ruvector-temporal-tensor
│   │       │   ├── heartbeat.rs      # HR via ruvector-temporal-tensor
│   │       │   └── triangulation.rs  # Position via ruvector-solver
│   │       └── crv/                  # Cross-room variant (planned)
│   │
│   │  ── APPLICATIONS ──────────────────────────────────────────────
│   ├── wifi-densepose-mat/           # Mass Casualty Assessment Tool
│   │   └── src/                      # (see MAT section below)
│   │
│   ├── wifi-densepose-sensing-server/# Lightweight Axum sensing server
│   │   └── src/                      # (see Sensing Server section below)
│   │
│   ├── wifi-densepose-api/           # REST API (Axum)
│   │   └── src/lib.rs                # Route definitions, middleware
│   │
│   ├── wifi-densepose-cli/           # CLI tool (`wifi-densepose` binary)
│   │   └── src/
│   │       ├── main.rs               # Entry point
│   │       ├── lib.rs                # CLI modules
│   │       └── mat.rs                # MAT subcommands
│   │
│   │  ── DEPLOYMENT TARGETS ────────────────────────────────────────
│   ├── wifi-densepose-wasm/          # WebAssembly bindings (browser)
│   │   └── src/
│   │       ├── lib.rs                # wasm-bindgen exports
│   │       └── mat.rs                # MAT functions for JS
│   │
│   ├── wifi-densepose-wasm-edge/     # WASM edge (no_std, excluded)
│   │   └── Targets: wasm32-unknown-unknown
│   │
│   ├── wifi-densepose-desktop/       # Tauri v2 desktop app
│   │   └── src/
│   │       ├── main.rs               # Tauri bootstrapper
│   │       ├── lib.rs
│   │       ├── state.rs              # Application state
│   │       └── commands/             # IPC command handlers
│   │
│   │  ── HARDWARE ──────────────────────────────────────────────────
│   └── wifi-densepose-hardware/      # ESP32 aggregator, TDM, QUIC
│       └── src/
│           ├── lib.rs
│           ├── bridge.rs             # ESP32 ↔ Server bridge
│           ├── csi_frame.rs          # CSI frame parser
│           ├── esp32_parser.rs       # ESP32 binary protocol
│           ├── error.rs
│           ├── aggregator/
│           │   └── mod.rs            # Multi-node CSI aggregation
│           ├── esp32/
│           │   ├── mod.rs
│           │   ├── tdm.rs            # Time-division multiplexing
│           │   ├── secure_tdm.rs     # Encrypted TDM (ADR-032)
│           │   └── quic_transport.rs # QUIC mesh networking
│           └── bin/                  # Hardware test binaries
│
├── data/
│   ├── adaptive_model.json           # Adaptive classifier state
│   ├── models/
│   │   ├── trained-pretrain-*.rvf    # Pretrained model checkpoints
│   │   └── trained-supervised-*.rvf  # Supervised model checkpoints
│   └── recordings/
│       ├── rec_*.csi.jsonl           # Raw CSI recordings (JSONL)
│       └── rec_*.csi.meta.json       # Recording metadata
│
├── docs/
│   ├── adr/                          # 78 Architecture Decision Records
│   └── ddd/                          # 8 Domain-Driven Design models
│
└── target/                           # Build artifacts (gitignored)
    ├── debug/                        # Debug builds
    └── release/                      # Release builds
        └── sensing-server            # Production binary
```

---

## RuvSense Modules

**Location:** `rust-port/wifi-densepose-rs/crates/wifi-densepose-signal/src/ruvsense/`

16 state-of-the-art multistatic WiFi sensing modules:

```
ruvsense/
├── mod.rs                    # Module registry & public API
│
│ ── SIGNAL CONDITIONING ─────────────────────────────────
├── multiband.rs              # Multi-band CSI fusion, cross-channel coherence
├── phase_align.rs            # Iterative LO phase offset estimation (circular mean)
├── coherence.rs              # Z-score coherence scoring, DriftProfile
├── coherence_gate.rs         # Accept / PredictOnly / Reject / Recalibrate gate
│
│ ── SPATIAL PROCESSING ──────────────────────────────────
├── multistatic.rs            # Attention-weighted N×(N-1) link fusion
├── field_model.rs            # SVD room eigenstructure, perturbation extraction
├── tomography.rs             # RF tomography, ISTA L1 sparse solver, voxel grid
│
│ ── TRACKING & IDENTIFICATION ───────────────────────────
├── pose_tracker.rs           # 17-keypoint Kalman tracker + AETHER re-ID embeddings
├── longitudinal.rs           # Welford running stats, biomechanics drift detection
├── attractor_drift.rs        # Strange attractor stability monitoring
│
│ ── ACTIVITY & GESTURE ──────────────────────────────────
├── gesture.rs                # DTW template-matching gesture classifier
├── temporal_gesture.rs       # Temporal gesture sequence recognition
├── intention.rs              # Pre-movement lead signal detection (200-500ms)
│
│ ── ENVIRONMENT & SECURITY ──────────────────────────────
├── cross_room.rs             # Environment fingerprinting, room transition graph
└── adversarial.rs            # Physically impossible signal detection, spoofing guard
```

---

## Cross-Viewpoint Fusion

**Location:** `rust-port/wifi-densepose-rs/crates/wifi-densepose-ruvector/src/viewpoint/`

```
viewpoint/
├── mod.rs                    # Module exports
├── attention.rs              # CrossViewpointAttention, GeometricBias, softmax + G_bias
├── geometry.rs               # GeometricDiversityIndex, Cramer-Rao bounds, Fisher Information
├── coherence.rs              # Phase phasor coherence, hysteresis gate
└── fusion.rs                 # MultistaticArray aggregate root, domain events
```

---

## Neural Network & Training

### WiFlow Architecture (ADR-072)
```
Model Pipeline:
  Raw CSI (64 subcarriers × 20 frames)
      ↓
  TCN encoder (temporal convolutions)
      ↓
  Axial attention (spatial × temporal)
      ↓
  Pose decoder → 17 COCO keypoints (x, y, confidence)
      ↓
  Quantization: FP32 → INT4 (881 KB → 8 KB for ESP32 SRAM)

Training Flow:
  1. Self-supervised contrastive pretraining (ADR-024/070)
  2. Supervised fine-tuning on MM-Fi / Wi-Pose (ADR-015)
  3. Per-room LoRA adaptation (2,048 params per adapter)
  4. EWC continual learning (avoid catastrophic forgetting)
  5. Quantization: FP32 → 4-bit (fits ESP32 SRAM)
  6. Export: .rvf container (model + metadata + witness hash)
```

### Training Crate File Map
```
wifi-densepose-train/src/
├── trainer.rs          # Epoch loop, gradient accumulation, checkpointing
├── model.rs            # WiFlow: TCN + axial attention + pose decoder
├── dataset.rs          # CompressedCsiBuffer, streaming JSONL, MM-Fi
├── losses.rs           # ContrastiveLoss + PoseMSE + SkeletonPhysicsLoss
├── eval.rs             # PCK@20, AUC, MPJPE, skeleton violation rate
├── metrics.rs          # DynamicPersonMatcher (ruvector-mincut graph cut)
├── subcarrier.rs       # 114→56 sparse interpolation (ruvector-solver)
├── rapid_adapt.rs      # Online: EWC penalty + LoRA inject + distillation
├── virtual_aug.rs      # Synthetic augmentation (noise, rotation, occlusion)
├── proof.rs            # SHA-256 deterministic pipeline proof (ADR-028)
└── aedis_metrics.rs   # Dashboard-specific KPIs
```

---

## Mass Casualty Assessment Tool

**Location:** `rust-port/wifi-densepose-rs/crates/wifi-densepose-mat/`

```
wifi-densepose-mat/src/
├── lib.rs                    # MAT public API
│
├── detection/                # Vital sign detection pipeline
│   ├── mod.rs
│   ├── pipeline.rs           # Ingestion → processing → classification
│   ├── breathing.rs          # Respiratory rate extraction
│   ├── heartbeat.rs          # Heart rate extraction
│   ├── movement.rs           # Movement index computation
│   └── ensemble.rs           # Multi-signal ensemble voting
│
├── localization/             # Survivor positioning
│   ├── mod.rs
│   ├── triangulation.rs      # RSSI/ToF triangulation
│   ├── depth.rs              # Through-debris depth estimation
│   └── fusion.rs             # Multi-sensor position fusion
│
├── tracking/                 # Lifecycle & re-identification
│   ├── mod.rs
│   ├── tracker.rs            # Multi-target tracker
│   ├── lifecycle.rs          # Detected → Confirmed → Lost
│   ├── kalman.rs             # Kalman state estimation
│   └── fingerprint.rs        # CSI-based re-ID fingerprint
│
├── alerting/                 # Alert dispatch & triage
│   ├── mod.rs
│   ├── dispatcher.rs         # Alert routing & deduplication
│   ├── generator.rs          # Alert generation rules
│   └── triage_service.rs     # START triage classification
│
├── ml/                       # ML helpers
│   ├── mod.rs
│   ├── vital_signs_classifier.rs  # Vitals → triage category
│   └── debris_model.rs       # RF attenuation through materials
│
├── domain/                   # DDD domain models
├── api/                      # REST endpoints for MAT
└── integration/              # External system bridges
```

---

## Sensing Server

**Location:** `rust-port/wifi-densepose-rs/crates/wifi-densepose-sensing-server/`

The lightweight Axum HTTP server that bridges ESP32 hardware → UI.

```
wifi-densepose-sensing-server/src/
├── main.rs                   # Server bootstrap, Axum router
├── lib.rs                    # Library exports
│
│ ── INFERENCE & CLASSIFICATION ──────────────────────────
├── adaptive_classifier.rs    # Online learning classifier (ADR-048)
├── sparse_inference.rs       # Low-latency sparse model inference
├── embedding.rs              # 128-dim CSI embedding engine
├── graph_transformer.rs      # GNN-based pattern recognition
├── sona.rs                   # Self-Organizing Neural Architecture
│
│ ── MODEL MANAGEMENT ────────────────────────────────────
├── model_manager.rs          # Model loading, hot-swapping, versioning
├── rvf_container.rs          # RVF cognitive container (model + metadata)
├── rvf_pipeline.rs           # RVF ingestion → inference pipeline
│
│ ── SENSING BRIDGES ─────────────────────────────────────
├── vital_signs.rs            # HTTP → vitals extraction bridge
├── multistatic_bridge.rs     # Multi-node CSI fusion bridge
├── field_bridge.rs           # Room field model bridge
├── tracker_bridge.rs         # Pose tracker bridge
│
│ ── TRAINING & DATA ─────────────────────────────────────
├── trainer.rs                # In-server training loop
├── training_api.rs           # /api/v1/training/* endpoints
├── dataset.rs                # Recording → training dataset
└── recording.rs              # CSI session recording to JSONL
```

---

## Python v1 Codebase

**Location:** `v1/`

```
v1/
├── __init__.py
├── setup.py
├── requirements-lock.txt         # Pinned dependency versions
│
├── src/
│   ├── app.py                    # FastAPI application root
│   ├── main.py                   # Uvicorn entry point
│   ├── cli.py                    # Click CLI (`wifi-densepose` / `wdp`)
│   ├── config.py                 # Settings loader
│   ├── logger.py                 # Structured logging (structlog)
│   │
│   ├── api/                      # REST API layer
│   │   ├── main.py               # API router aggregation
│   │   ├── dependencies.py       # FastAPI dependency injection
│   │   ├── routers/
│   │   │   ├── health.py         # GET /health
│   │   │   ├── pose.py           # POST /pose/estimate
│   │   │   └── stream.py         # WebSocket /stream/csi
│   │   ├── middleware/
│   │   │   ├── auth.py           # JWT + API key authentication
│   │   │   ├── cors.py           # CORS policy
│   │   │   ├── error_handler.py  # Global exception handler
│   │   │   └── rate_limit.py     # Token bucket rate limiter
│   │   └── websocket/            # WebSocket handlers
│   │
│   ├── core/                     # Core processing
│   │   ├── csi_processor.py      # CSI pipeline (numpy/scipy)
│   │   ├── phase_sanitizer.py    # Phase unwrapping
│   │   └── router_interface.py   # Hardware abstraction
│   │
│   ├── models/                   # Neural network definitions
│   │   ├── densepose_head.py     # DensePose prediction head
│   │   └── modality_translation.py  # CSI → image translation net
│   │
│   ├── services/                 # Business logic
│   │   ├── orchestrator.py       # Service orchestration
│   │   ├── pose_service.py       # Pose estimation service
│   │   ├── stream_service.py     # CSI stream management
│   │   ├── hardware_service.py   # Hardware lifecycle
│   │   ├── health_check.py       # System health
│   │   ├── metrics.py            # Prometheus metrics
│   │   └── training.service.py   # Training orchestrator
│   │
│   ├── sensing/                  # Sensing backend
│   │   ├── backend.py            # Sensing pipeline coordinator
│   │   ├── classifier.py         # Activity classifier
│   │   ├── feature_extractor.py  # CSI feature extraction
│   │   ├── rssi_collector.py     # RSSI-only mode
│   │   ├── mac_wifi.swift        # macOS CoreWLAN native bridge
│   │   └── ws_server.py          # WebSocket push server
│   │
│   ├── hardware/                 # Hardware integration
│   │   ├── csi_extractor.py      # ESP32 CSI frame decoder
│   │   └── router_interface.py   # Router management
│   │
│   ├── database/                 # Data persistence
│   │   ├── connection.py         # SQLAlchemy async session
│   │   ├── models.py             # ORM models
│   │   ├── model_types.py        # Type definitions
│   │   └── migrations/           # Alembic migrations
│   │
│   ├── config/
│   │   ├── settings.py           # Pydantic Settings
│   │   └── domains.py            # Domain configuration
│   │
│   ├── commands/                 # CLI commands
│   │   ├── start.py
│   │   ├── status.py
│   │   └── stop.py
│   │
│   └── tasks/                    # Background tasks
│       ├── backup.py             # Periodic backup
│       ├── cleanup.py            # Data cleanup
│       └── monitoring.py         # Health pulse
│
├── data/
│   ├── proof/                    # ADR-028 deterministic verification
│   │   ├── verify.py             # Trust Kill Switch — feeds reference signal,
│   │   │                         #   hashes output, compares to published hash
│   │   ├── expected_features.sha256  # Published expected hash
│   │   ├── sample_csi_data.json  # 1,000 synthetic CSI frames (seed=42)
│   │   ├── sample_csi_meta.json  # Metadata for proof frames
│   │   └── generate_reference_signal.py
│   │
│   ├── test_wifi_densepose.db    # SQLite test database
│   └── wifi_densepose_fallback.db
│
├── tests/
│   ├── unit/                     # Unit tests
│   ├── integration/              # Integration tests
│   ├── e2e/                      # End-to-end tests
│   ├── performance/              # Performance benchmarks
│   ├── fixtures/                 # Test fixtures
│   └── mocks/                    # Mock objects
│
└── scripts/
    ├── start.py                  # Start server
    ├── status.py                 # Check status
    └── stop.py                   # Stop server
```

---

## ML Models & Data

### Trained Model Assets

```
data/
├── models/                       # Pre-trained model checkpoints
│   └── (downloaded from HuggingFace or trained locally)
└── recordings/                   # CSI recordings for training/replay
    ├── rec_*.csi.jsonl           # Raw CSI frames (one JSON per line)
    └── rec_*.csi.meta.json       # Recording metadata (duration, nodes, etc.)

rust-port/wifi-densepose-rs/data/
├── adaptive_model.json           # Online adaptive classifier state
├── models/
│   ├── trained-pretrain-*.rvf    # Self-supervised pretraining checkpoints
│   └── trained-supervised-*.rvf  # Supervised fine-tuning checkpoints
└── recordings/
    ├── rec_*.csi.jsonl           # Rust-collected CSI recordings
    └── rec_*.csi.meta.json       # Recording metadata
```

### HuggingFace Published Models

| File | Size | Purpose |
|------|------|---------|
| `model.safetensors` | 48 KB | Contrastive encoder (128-dim embeddings) |
| `model-q4.bin` | 8 KB | 4-bit quantized (ESP32 SRAM inference) |
| `model-q2.bin` | 4 KB | 2-bit ultra-compact edge variant |
| `presence-head.json` | 2.6 KB | Presence detection head (100% accuracy) |
| `node-1.json` / `node-2.json` | 21 KB | Per-room LoRA adapters |

### Model Format: RVF (RuVector Format)

```
.rvf container structure:
  ├── header          # Version, dimensions, metadata
  ├── weights         # Quantized model weights
  ├── config          # Hyperparameters, architecture spec
  ├── witness_hash    # SHA-256 integrity proof
  └── metadata        # Training provenance, dataset info
```

---

## Firmware (ESP32)

**Location:** `firmware/esp32-csi-node/`

```
firmware/esp32-csi-node/
├── CMakeLists.txt                # ESP-IDF project file
├── README.md
├── provision.py                  # WiFi + target IP provisioning
├── build_firmware.bat            # Windows build script
├── build_firmware.ps1            # PowerShell build script
│
├── main/                         # ESP-IDF main component
│   ├── CMakeLists.txt
│   ├── Kconfig.projbuild         # Build-time configuration
│   ├── idf_component.yml         # Component dependencies
│   ├── main.c                    # RTOS entry point
│   │
│   │ ── CSI CAPTURE ────────────────────────────────────
│   ├── csi_collector.c/h         # WiFi CSI callback + buffering
│   ├── stream_sender.c/h         # UDP/TCP CSI frame streaming
│   │
│   │ ── EDGE INTELLIGENCE ──────────────────────────────
│   ├── edge_processing.c/h       # On-device ML inference (INT4)
│   ├── wasm_runtime.c/h          # Wasm3 WASM runtime integration
│   ├── wasm_upload.c/h           # OTA WASM module deployment
│   ├── rvf_parser.c/h            # RVF model file loading
│   │
│   │ ── MESH & NETWORKING ──────────────────────────────
│   ├── swarm_bridge.c/h          # Multi-node mesh networking
│   │
│   │ ── SENSORS ────────────────────────────────────────
│   ├── mmwave_sensor.c/h         # 60 GHz mmWave (MR60BHA2 SoC)
│   │
│   │ ── DISPLAY (optional) ─────────────────────────────
│   ├── display_hal.c/h           # Hardware abstraction (SPI/I2C)
│   ├── display_task.c/h          # RTOS display task
│   ├── display_ui.c/h            # LVGL UI rendering
│   ├── lv_conf.h                 # LVGL configuration
│   │
│   │ ── SYSTEM ─────────────────────────────────────────
│   ├── nvs_config.c/h            # Non-volatile storage config
│   ├── ota_update.c/h            # Over-the-air firmware updates
│   ├── power_mgmt.c/h            # Sleep modes, battery management
│   └── mock_csi.c/h              # Test-only mock (stripped in release)
│
├── components/
│   └── wasm3/                    # Vendored Wasm3 runtime
│
├── sdkconfig.defaults            # 8MB flash default config (real CSI)
├── sdkconfig.defaults.4mb        # 4MB flash variant
├── sdkconfig.defaults.template   # Base template for releases
├── sdkconfig.coverage            # Code coverage instrumentation
├── sdkconfig.qemu                # QEMU emulation config
│
├── partitions_4mb.csv            # 4MB partition table
├── partitions_display.csv        # Display variant partitions
│
├── test/                         # Firmware unit tests
└── release_bins/                 # Signed release binaries
    ├── esp32-csi-node.bin        # 8MB firmware
    ├── esp32-csi-node-4mb.bin    # 4MB firmware
    ├── bootloader.bin
    ├── partition-table.bin
    ├── partition-table-4mb.bin
    └── ota_data_initial.bin
```

### Supported Hardware Matrix

| Device | Chip | Role | Flash | Cost |
|--------|------|------|-------|------|
| ESP32-S3 (8MB) | Xtensa dual-core 240 MHz | CSI sensing node | 8 MB | ~$9 |
| ESP32-S3 SuperMini (4MB) | Xtensa dual-core 240 MHz | CSI compact node | 4 MB | ~$6 |
| ESP32-C6 + MR60BHA2 | RISC-V + 60 GHz FMCW | mmWave HR/BR/presence | — | ~$15 |
| HLK-LD2410 | 24 GHz FMCW | Presence + distance | — | ~$3 |

**Not supported:** ESP32 (original), ESP32-C3 — single-core, insufficient for CSI DSP pipeline.

---

## Web & Mobile UI

**Location:** `ui/`

```
ui/
├── index.html                    # App shell
├── app.js                        # Application entry
├── style.css                     # Global styles
├── viz.html                      # Visualization standalone
├── start-ui.sh                   # Dev server launcher
│
├── config/
│   └── api.config.js             # API endpoint configuration
│
├── components/                   # UI components
│   ├── DashboardTab.js           # Main dashboard
│   ├── SensingTab.js             # Live sensing view
│   ├── HardwareTab.js            # Hardware management
│   ├── LiveDemoTab.js            # Live demo (no hardware)
│   ├── TrainingPanel.js          # Model training controls
│   ├── ModelPanel.js             # Model management
│   ├── SettingsPanel.js          # Settings configuration
│   ├── PoseDetectionCanvas.js    # Pose visualization canvas
│   ├── TabManager.js             # Tab navigation
│   ├── body-model.js             # 3D body visualization
│   ├── gaussian-splats.js        # 3D Gaussian splatting
│   ├── signal-viz.js             # Signal visualization
│   ├── dashboard-hud.js          # HUD overlay
│   ├── environment.js            # Environment renderer
│   └── scene.js                  # Three.js scene
│
├── services/                     # Service layer
│   ├── api.service.js            # HTTP API client
│   ├── websocket.service.js      # WebSocket management
│   ├── sensing.service.js        # Sensing data handler
│   ├── pose.service.js           # Pose data handler
│   ├── model.service.js          # Model operations
│   ├── training.service.js       # Training controls
│   ├── stream.service.js         # Stream management
│   ├── health.service.js         # Health monitoring
│   ├── data-processor.js         # Data transformation
│   └── websocket-client.js       # Low-level WS client
│
├── utils/
│   ├── backend-detector.js       # Auto-detect Rust/Python backend
│   ├── mock-server.js            # Development mock server
│   └── pose-renderer.js          # Pose skeleton rendering
│
├── pose-fusion/                  # Dual-modal pose fusion UI
│   ├── build.sh                  # WASM build script
│   ├── css/                      # Styles
│   ├── js/                       # Fusion logic
│   └── pkg/                      # WASM packages
│
├── observatory/                  # Cinematic Three.js dashboard (ADR-047)
│   └── observatory.html          # 5-panel holographic display
│
├── pose-fusion.html              # Pose fusion standalone page
│
├── mobile/                       # React Native / Expo mobile app
│   ├── App.tsx                   # Root component
│   ├── app.config.ts             # Expo configuration
│   ├── package.json              # Node dependencies
│   ├── tsconfig.json             # TypeScript config
│   ├── eas.json                  # EAS Build config
│   │
│   ├── src/
│   │   ├── screens/              # App screens (Live, MAT, Vitals, Settings)
│   │   ├── components/           # Reusable UI components
│   │   ├── services/             # API + WebSocket services
│   │   ├── stores/               # State management
│   │   ├── hooks/                # Custom React hooks
│   │   ├── navigation/           # React Navigation setup
│   │   ├── theme/                # Design tokens
│   │   ├── types/                # TypeScript types
│   │   ├── utils/                # Utility functions
│   │   ├── constants/            # App constants
│   │   └── assets/               # Images, icons
│   │
│   └── e2e/                      # Maestro E2E tests
│       ├── live_screen.yaml
│       ├── mat_screen.yaml
│       ├── vitals_screen.yaml
│       ├── settings_screen.yaml
│       ├── zones_screen.yaml
│       └── offline_fallback.yaml
│
└── tests/
    ├── test-runner.html          # Browser test runner
    ├── test-runner.js            # Test framework
    └── integration-test.html     # Integration tests
```

---

## Edge Computing Infrastructure

### Architecture Overview

```
┌─────────────────────────────────────────────────────────┐
│                    EDGE TIER (On-Device)                  │
│                                                           │
│  ESP32-S3 Node ($9)          ESP32-S3 Node ($9)          │
│  ┌─────────────────┐        ┌─────────────────┐          │
│  │ csi_collector.c  │        │ csi_collector.c  │          │
│  │ edge_processing.c│        │ edge_processing.c│          │
│  │ wasm_runtime.c   │        │ wasm_runtime.c   │          │
│  │ swarm_bridge.c   │        │ swarm_bridge.c   │          │
│  │ power_mgmt.c     │        │ power_mgmt.c     │          │
│  └────────┬─────────┘        └────────┬─────────┘          │
│           │ UDP/TCP (CSI frames)       │                    │
│           └──────────┬─────────────────┘                    │
│                      ▼                                      │
│  ┌──────────────────────────────────┐                       │
│  │   Optional: Cognitum Seed ($131) │                       │
│  │   - Vector store (RVF format)    │                       │
│  │   - kNN similarity search        │                       │
│  │   - Witness chain (Ed25519)      │                       │
│  │   - 114-tool MCP proxy           │                       │
│  └──────────────┬───────────────────┘                       │
└─────────────────┼───────────────────────────────────────────┘
                  │ HTTP/WebSocket
                  ▼
┌─────────────────────────────────────────────────────────────┐
│                    FOG TIER (Local Server)                    │
│                                                               │
│  ┌─────────────────────────────────────────────────────┐     │
│  │         Sensing Server (Axum, Rust binary)          │     │
│  │  Port 3000 — /api/v1/sensing, /api/v1/nodes, etc.  │     │
│  │                                                     │     │
│  │  adaptive_classifier.rs  │  model_manager.rs        │     │
│  │  sparse_inference.rs     │  vital_signs.rs          │     │
│  │  multistatic_bridge.rs   │  trainer.rs              │     │
│  │  graph_transformer.rs    │  recording.rs            │     │
│  └────────────┬────────────────────────────────────────┘     │
│               │ Metrics (Prometheus /metrics)                 │
│               ▼                                               │
│  ┌─────────────────────┐  ┌─────────────────────┐           │
│  │ Prometheus (:9090)  │  │ Grafana (:3001)     │           │
│  │ Scrape: 15s         │  │ Dashboard: JSON     │           │
│  │ Alerts: YAML rules  │  │ Panels: 12          │           │
│  └─────────────────────┘  └─────────────────────┘           │
│                                                               │
│  ┌─────────────────────────────────────────────────────┐     │
│  │            Fluentd Log Aggregator                    │     │
│  │  Sources: kubernetes, systemd, wifi-densepose        │     │
│  │  Sinks: Elasticsearch, S3, stdout                    │     │
│  └─────────────────────────────────────────────────────┘     │
└───────────────────────────────────────────────────────────────┘
                  │
                  ▼ (optional)
┌─────────────────────────────────────────────────────────────┐
│                    CLOUD TIER (Optional)                      │
│                                                               │
│  Docker: ruvnet/wifi-densepose:latest (amd64 + arm64)        │
│  Kubernetes: Helm chart + HPA scaling                        │
│  HuggingFace: ruv/aedi-s (model hosting)                     │
│  Google Cloud: gcloud-train.sh (GPU training)                │
└─────────────────────────────────────────────────────────────┘
```

### Edge Module Categories (ADR-041)

Documented in `docs/edge-modules/`:

| Module Category | File | Capabilities |
|----------------|------|-------------|
| **Core** | `core.md` | Presence, motion, occupancy, vital signs |
| **Medical** | `medical.md` | BP estimation, fall detection, respiratory |
| **Security** | `security.md` | Intrusion, perimeter, anomaly detection |
| **Industrial** | `industrial.md` | Machine monitoring, worker safety |
| **Retail** | `retail.md` | Customer flow, dwell time, heatmaps |
| **Spatial-Temporal** | `spatial-temporal.md` | Room mapping, path tracking |
| **Signal Intel** | `signal-intelligence.md` | Spectrum analysis, interference |
| **AI Security** | `ai-security.md` | Adversarial detection, spoofing guard |
| **Adaptive Learning** | `adaptive-learning.md` | SNN online learning, LoRA adapters |
| **Autonomous** | `autonomous.md` | Self-healing mesh, auto-calibration |
| **Exotic** | `exotic.md` | Through-wall imaging, passive radar |

### WASM Edge Deployment (ADR-040)

```
wifi-densepose-wasm-edge/     # no_std crate, targets wasm32-unknown-unknown
  ↓ Builds to .wasm module
  ↓ Uploaded to ESP32 via wasm_upload.c
  ↓ Executed by wasm3 runtime (wasm_runtime.c)
  ↓ Hot-swappable: new module replaces old without reflash
```

---

## Monitoring & Logging

### Prometheus Metrics (`monitoring/`)

```
monitoring/
├── prometheus-config.yml     # Scrape configuration
│   ├── Global: 15s scrape interval
│   ├── Targets: sensing-server:3000/metrics
│   ├── Kubernetes: API server, nodes, cadvisor, pods
│   └── External: node-exporter, alertmanager
│
├── alerting-rules.yml        # Alert definitions
│   ├── HighLatency           # CSI processing > 100ms
│   ├── NodeDown              # ESP32 node offline > 5m
│   ├── VitalSignAnomaly      # HR/BR out of range
│   ├── HighMemoryUsage       # Server memory > 80%
│   ├── CsiFrameDropRate      # Frame drops > 5%
│   └── MeshPartition         # Node unreachable
│
└── grafana-dashboard.json    # Pre-built Grafana dashboard
    ├── Panel: CSI Frame Rate (real-time)
    ├── Panel: Signal Quality Score
    ├── Panel: Vital Signs (HR, BR)
    ├── Panel: Node Health Matrix
    ├── Panel: Inference Latency Histogram
    ├── Panel: Person Count Over Time
    ├── Panel: Memory & CPU Usage
    ├── Panel: Model Version Tracker
    ├── Panel: Alert History
    ├── Panel: Mesh Topology Graph
    ├── Panel: Training Progress
    └── Panel: Edge Inference Stats
```

### Log Aggregation (`logging/`)

```
logging/
└── fluentd-config.yml
    │
    ├── Source: Kubernetes container logs (/var/log/containers/*.log)
    ├── Source: systemd journal (ESP32 bridge service)
    ├── Source: wifi-densepose application logs
    │
    ├── Filter: Kubernetes metadata enrichment
    ├── Filter: Log level classification
    ├── Filter: CSI frame parsing (structured JSON)
    │
    ├── Output: Elasticsearch (search & analytics)
    ├── Output: S3 (long-term archival)
    └── Output: stdout (development)
```

### Runtime Logs (`logs/`)

```
logs/
├── sensing-server.pid        # Server process ID file
├── sensing-server.log        # Server output (when daemonized)
├── csi-stream.log            # Raw CSI frame log
├── training.log              # Training pipeline output
└── mesh-events.log           # Mesh topology changes
```

### Application-Level Logging

**Rust (tracing crate):**
```
RUST_LOG=wifi_densepose_sensing_server=debug,tower_http=info
→ Structured JSON logs with span context, timing, and trace IDs
→ Exported via tracing-subscriber env-filter + json layer
```

**Python (structlog):**
```
v1/src/logger.py → structlog with JSON output
v1/src/services/metrics.py → Prometheus client metrics
```

---

## Docker & Deployment

**Location:** `docker/`

```
docker/
├── Dockerfile.python             # Python v1 image
│   └── Base: python:3.11-slim
│       Installs: requirements.txt + PyTorch CPU
│       Exposes: 3000
│
├── Dockerfile.rust               # Rust sensing server image
│   └── Stage 1: rust:1.85 (build)
│       Stage 2: debian:bookworm-slim (runtime)
│       Binary: sensing-server
│       Exposes: 3000
│
├── docker-compose.yml            # Full stack orchestration
│   ├── Service: sensing-server (Rust)
│   ├── Service: python-api (Python v1)
│   ├── Service: prometheus
│   ├── Service: grafana
│   ├── Service: postgres
│   ├── Service: redis
│   └── Service: fluentd
│
├── wifi-densepose-v1.rvf         # Pre-built RVF model
└── .dockerignore
```

### Deployment Quick Reference

```bash
# Docker (simulated data, no hardware)
docker pull ruvnet/wifi-densepose:latest
docker run -p 3000:3000 ruvnet/wifi-densepose:latest

# Docker Compose (full stack with monitoring)
cd docker && docker-compose up -d

# Native Rust binary
cd rust-port/wifi-densepose-rs
cargo build -p wifi-densepose-sensing-server --release
./target/release/sensing-server --port 3000

# Kubernetes (via Helm or raw manifests)
# See deploy.sh for automated deployment
```

---

## Scripts & Utilities

**Location:** `scripts/` (44+ files)

### Categories

| Category | Scripts | Purpose |
|----------|---------|---------|
| **Signal Analysis** | `rf-scan.js`, `rf-scan-multifreq.js`, `csi-spectrogram.js`, `csi-graph-visualizer.js`, `deep-scan.js` | Real-time spectrum visualization and CSI analysis |
| **Health Sensing** | `sleep-monitor.js`, `apnea-detector.js`, `stress-monitor.js`, `gait-analyzer.js` | Vital signs, sleep quality, stress, and gait |
| **Security** | `through-wall-detector.js`, `passive-radar.js`, `device-fingerprint.js`, `room-fingerprint.js` | Through-wall imaging, passive radar, device ID |
| **ML Training** | `train-ruvllm.js`, `train-wiflow.js`, `train-camera-free.js`, `collect-training-data.py`, `benchmark-model.py` | Model training, data collection, benchmarking |
| **Person Counting** | `mincut-person-counter.js`, `snn-csi-processor.js` | Graph-based person separation, SNN adaptation |
| **Materials** | `material-detector.js`, `material-classifier.js` | Object detection via subcarrier null changes |
| **RF Tomography** | `rf-tomography.js` | 2D room imaging via RF backprojection |
| **Mesh Graph** | `mesh-graph-transformer.js` | GNN-based mesh topology analysis |
| **QEMU Testing** | `qemu_swarm.py`, `qemu-esp32s3-test.sh`, `qemu-mesh-test.sh`, `qemu-chaos-test.sh` | Firmware emulation and testing |
| **Deployment** | `flash-node.sh`, `provision.py`, `deploy.sh`, `start-aedi-s.sh` | Hardware flashing, provisioning, deployment |
| **Publishing** | `publish-huggingface.py`, `publish-huggingface.sh`, `generate-witness-bundle.sh` | Model publishing, witness verification |
| **Utilities** | `check_health.py`, `swarm_health.py`, `inject_fault.py`, `seed_csi_bridge.py` | Health checks, fault injection, data bridging |

### QEMU Swarm Presets (`scripts/swarm_presets/`)

```
swarm_presets/
├── smoke.yaml              # Quick 1-node smoke test
├── standard.yaml           # 3-node baseline mesh
├── large_mesh.yaml         # 10+ node stress test
├── heterogeneous.yaml      # Mixed ESP32-S3 + C6 hardware
├── line_relay.yaml         # Linear relay topology
├── ring_fault.yaml         # Ring with fault injection
└── ci_matrix.yaml          # CI test matrix
```

---

## Documentation & ADRs

**Location:** `docs/`

### Architecture Decision Records (78 ADRs)

```
docs/adr/
├── README.md                 # ADR index and navigation guide
│
│ ── CORE ARCHITECTURE (001-013) ────────────────────────
├── ADR-001  WiFi-MAT disaster detection
├── ADR-002  RuVector RVF integration strategy
├── ADR-003  RVF cognitive containers for CSI
├── ADR-004  HNSW vector search fingerprinting
├── ADR-005  SONA self-learning pose estimation
├── ADR-006  GNN-enhanced CSI pattern recognition
├── ADR-007  Post-quantum cryptography secure sensing
├── ADR-008  Distributed consensus multi-AP
├── ADR-009  RVF WASM runtime edge deployment
├── ADR-010  Witness chains audit trail integrity
├── ADR-011  Python proof-of-reality mock elimination
├── ADR-012  ESP32 CSI sensor mesh
├── ADR-013  Feature-level sensing commodity gear
│
│ ── SIGNAL PROCESSING (014-017) ────────────────────────
├── ADR-014  SOTA signal processing ✓ ACCEPTED
├── ADR-015  MM-Fi + Wi-Pose training datasets ✓ ACCEPTED
├── ADR-016  RuVector integration ✓ COMPLETE
├── ADR-017  RuVector signal + MAT integration → PROPOSED
│
│ ── HARDWARE & FIRMWARE (018-022) ──────────────────────
├── ADR-018  ESP32 dev implementation
├── ADR-019  Sensing-only UI mode
├── ADR-020  Rust RuVector AI model migration
├── ADR-021  Vital sign detection RVDNA pipeline
├── ADR-022  Windows WiFi enhanced fidelity
│
│ ── ML PIPELINE (023-027) ──────────────────────────────
├── ADR-023  Trained DensePose model pipeline
├── ADR-024  Contrastive CSI embedding (AETHER) ✓ ACCEPTED
├── ADR-025  macOS CoreWLAN WiFi sensing
├── ADR-026  Survivor track lifecycle
├── ADR-027  Cross-environment domain generalization (MERIDIAN) ✓ ACCEPTED
│
│ ── VERIFICATION & SECURITY (028-032) ──────────────────
├── ADR-028  ESP32 capability audit ✓ ACCEPTED
├── ADR-029  RuvSense multistatic sensing mode → PROPOSED
├── ADR-030  RuvSense persistent field model → PROPOSED
├── ADR-031  AEDI-S sensing-first RF mode → PROPOSED
├── ADR-032  Multistatic mesh security hardening → PROPOSED
│
│ ── PLATFORM (033-056) ─────────────────────────────────
├── ADR-033  CRV signal-line sensing integration
├── ADR-034  Expo mobile app
├── ADR-035  Live sensing UI accuracy
├── ADR-036  RVF training pipeline UI
├── ADR-037  Multi-person pose detection
├── ADR-038  Sublinear goal-oriented action planning
├── ADR-039  ESP32 edge intelligence
├── ADR-040  WASM programmable sensing
├── ADR-041  WASM module collection
├── ADR-042  Coherent human-channel imaging (CHCI)
├── ADR-043  Sensing server UI/API completion
├── ADR-044  Provisioning tool enhancements
├── ADR-045  AMOLED display support
├── ADR-046  Android TV box Armbian deployment
├── ADR-047  Psychohistory observatory visualization
├── ADR-048  Adaptive CSI classifier
├── ADR-049  Cross-platform WiFi interface detection
├── ADR-050  Quality engineering & security hardening
├── ADR-052  DDD bounded contexts
├── ADR-052  Tauri desktop frontend (duplicate number)
├── ADR-053  UI design system
├── ADR-054  Desktop full implementation
├── ADR-055  Integrated sensing server
├── ADR-056  AEDI-S desktop capabilities
│
│ ── FIRMWARE (057-062) ─────────────────────────────────
├── ADR-057  Firmware CSI build guard
├── ADR-058  RuVector WASM browser pose example
├── ADR-059  Live ESP32 CSI pipeline
├── ADR-060  Provision channel + MAC filter
├── ADR-061  QEMU ESP32-S3 firmware testing
├── ADR-062  QEMU swarm configurator
│
│ ── MULTIMODAL (063-068) ───────────────────────────────
├── ADR-063  mmWave sensor fusion
├── ADR-064  Multimodal ambient intelligence
├── ADR-065  Happiness scoring seed bridge
├── ADR-066  ESP32 swarm seed coordinator
├── ADR-067  RuVector v2.0.5 upgrade
├── ADR-068  Per-node state pipeline
│
│ ── ADVANCED ML (069-078) ──────────────────────────────
├── ADR-069  Cognitum Seed CSI pipeline
├── ADR-070  Self-supervised pretraining
├── ADR-071  RuVLLM training pipeline
├── ADR-072  WiFlow architecture
├── ADR-073  Multi-frequency mesh scan
├── ADR-074  Spiking neural CSI sensing
├── ADR-075  MinCut person separation
├── ADR-076  CSI spectrogram embeddings
├── ADR-077  Novel RF sensing applications
└── ADR-078  Multi-freq mesh applications
```

### Domain-Driven Design Models

```
docs/ddd/
├── README.md
├── chci-domain-model.md                  # Computer-Human Channel Interaction
├── deployment-platform-domain-model.md   # Docker/K8s/edge deployment
├── hardware-platform-domain-model.md     # ESP32 + mmWave hardware
├── ruvsense-domain-model.md              # Multistatic sensing domain
├── sensing-server-domain-model.md        # Axum server domain
├── signal-processing-domain-model.md     # CSI pipeline domain
├── training-pipeline-domain-model.md     # ML training domain
└── wifi-mat-domain-model.md              # MAT disaster response domain
```

### Research & Tutorials

```
docs/research/
├── architecture/         # System architecture papers
├── arena-physica/        # Physical benchmarking arena design
├── neural-decoding/      # Brain ↔ WiFi neural decoding research
├── quantum-sensing/      # Quantum-enhanced WiFi sensing theory
├── rf-topological-sensing/  # RF topology insights
└── sota-surveys/         # State-of-the-art literature reviews

docs/tutorials/
└── cognitum-seed-pretraining.md  # Tutorial: Seed pretraining walkthrough

docs/edge-modules/
├── README.md             # Edge module overview
├── core.md               # Core sensing modules
├── medical.md            # Healthcare edge modules
├── security.md           # Security / intrusion modules
├── industrial.md         # Industrial IoT modules
├── retail.md             # Retail analytics modules
├── spatial-temporal.md   # Spatial mapping modules
├── signal-intelligence.md  # Spectrum analysis modules
├── ai-security.md        # Adversarial defense modules
├── adaptive-learning.md  # Online learning modules
├── autonomous.md         # Self-healing mesh modules
└── exotic.md             # Through-wall, passive radar
```

---

## CI/CD & GitHub Workflows

**Location:** `.github/workflows/`

```
.github/workflows/
├── ci.yml                    # Main CI: Rust tests + Python proof + lint
│   ├── Job: rust-test        # cargo test --workspace --no-default-features
│   ├── Job: python-proof     # python v1/data/proof/verify.py
│   ├── Job: lint             # cargo clippy + flake8
│   └── Trigger: push to main, PRs
│
├── cd.yml                    # Continuous deployment
│   ├── Job: docker-build     # Multi-arch image (amd64 + arm64)
│   ├── Job: docker-push      # Push to Docker Hub
│   └── Trigger: tagged releases
│
├── firmware-ci.yml           # ESP32 firmware build verification
│   ├── Job: build-8mb        # 8MB flash firmware
│   ├── Job: build-4mb        # 4MB flash firmware
│   └── Trigger: firmware/ changes
│
├── firmware-qemu.yml         # QEMU firmware emulation tests
│   ├── Job: qemu-boot        # Boot test in QEMU
│   ├── Job: qemu-mesh        # Mesh topology test
│   └── Trigger: firmware/ changes
│
├── verify-pipeline.yml       # ADR-028 witness verification
│   ├── Job: proof            # Deterministic hash verification
│   ├── Job: witness-bundle   # Full bundle generation
│   └── Trigger: manual, main push
│
├── desktop-release.yml       # Tauri desktop app release
│   ├── Job: build-macos      # macOS arm64 + x86_64
│   ├── Job: build-linux      # Linux x86_64
│   ├── Job: build-windows    # Windows x86_64
│   └── Trigger: desktop/* tags
│
├── security-scan.yml         # Security scanning
│   ├── Job: cargo-audit      # Rust dependency audit
│   ├── Job: bandit           # Python security scan
│   └── Trigger: weekly + PRs
│
└── update-submodules.yml     # Vendor submodule updates
    └── Trigger: weekly schedule
```

---

## Configuration & Tooling

### Claude Code Configuration (`.claude/`)

```
.claude/
├── settings.json             # Per-project Copilot settings
├── settings.local.json       # Local overrides (gitignored)
├── memory.db                 # Context memory database
│
├── agents/                   # 20+ custom agent definitions
│   ├── core/                 # Core development agents
│   ├── architecture/         # Architecture analysis
│   ├── testing/              # Test generation
│   ├── documentation/        # Doc generation
│   ├── security/             # Security audit
│   ├── github/               # PR/issue management
│   ├── sona/                 # Self-learning orchestration
│   ├── sparc/                # SPARC methodology
│   └── ...                   # 12+ more categories
│
├── skills/                   # 30+ domain skills
│   ├── agentdb-*/            # Vector database skills (5)
│   ├── v3-*/                 # Claude Flow v3 skills (9)
│   ├── github-*/             # GitHub integration skills (5)
│   ├── swarm-*/              # Swarm orchestration skills (2)
│   └── ...                   # Other domain skills
│
├── commands/                 # Custom CLI commands
└── helpers/                  # Utility functions
```

### Ionity Platform (`.ionity/`)

```
.ionity/
├── ionity.sh                 # Platform CLI entry point
├── README.md
└── lib/
    ├── build.sh              # Build orchestration
    ├── deps.sh               # Dependency management
    ├── provisioner.sh        # Deployment provisioner
    ├── scripts_menu.sh       # Interactive script launcher
    └── stack.sh              # Stack management
```

### Claude Flow V3 (`.claude-flow/`)

```
.claude-flow/
├── config.yaml               # Workflow configuration
├── daemon-state.json          # Daemon process state
├── CAPABILITIES.md            # Capabilities matrix
└── metrics/
    ├── codebase-map.json      # Codebase topology graph
    ├── learning.json          # Learning progress
    ├── security-audit.json    # Security audit state
    ├── swarm-activity.json    # Swarm statistics
    └── v3-progress.json       # V3 milestone tracking
```

---

## Examples & Tutorials

**Location:** `examples/`

| Example | Directory | Description |
|---------|-----------|-------------|
| **Live Streaming** | `aedis_live.py` | Real-time CSI visualization |
| **Room Monitor** | `environment/room_monitor.py` | Room-level occupancy analytics |
| **Blood Pressure** | `medical/bp_estimator.py` | Contactless BP via 60 GHz mmWave |
| **Vitals Suite** | `medical/vitals_suite.py` | Full vital signs dashboard |
| **Happiness Vector** | `happiness-vector/` | Emotion/mood sensing with swarm |
| **Sleep Quality** | `sleep/apnea_screener.py` | Overnight sleep + apnea screening |
| **Stress Monitor** | `stress/hrv_stress_monitor.py` | HRV-based stress detection |

---

## Quick Access Cheat Sheet

| I want to... | Go to |
|--------------|-------|
| Build the Rust sensing server | `rust-port/wifi-densepose-rs/` |
| Train an ML model | `rust-port/wifi-densepose-rs/crates/wifi-densepose-train/` |
| Run Python v1 API | `v1/src/app.py` |
| Flash ESP32 firmware | `firmware/esp32-csi-node/` |
| View the web UI | `ui/index.html` |
| Run signal analysis scripts | `scripts/rf-scan.js` |
| Check monitoring dashboards | `monitoring/grafana-dashboard.json` |
| Read architecture decisions | `docs/adr/` |
| See domain models | `docs/ddd/` |
| Check CI pipeline | `.github/workflows/ci.yml` |
| Deploy via Docker | `docker/docker-compose.yml` |
| View research papers | `docs/research/` |
| Download pre-trained models | `https://huggingface.co/ruv/aedi-s` |
| Run deterministic proof | `python v1/data/proof/verify.py` |
| Test firmware in QEMU | `scripts/qemu-esp32s3-test.sh` |
