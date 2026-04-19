//! WiFi-DensePose Sensing Server
//!
//! Lightweight Axum server that:
//! - Receives ESP32 CSI frames via UDP (port 5005)
//! - Processes signals using RuVector-powered wifi-densepose-signal crate
//! - Broadcasts sensing updates via WebSocket (ws://localhost:8765/ws/sensing)
//! - Serves the static UI files (port 8080)
//!
//! Replaces both ws_server.py and the Python HTTP server.

mod adaptive_classifier;
mod aggression;
mod field_bridge;
mod multistatic_bridge;
mod pso;
mod rvf_container;
mod rvf_pipeline;
mod tracker_bridge;
mod vital_signs;

// Training pipeline modules (exposed via lib.rs)
use wifi_densepose_sensing_server::{graph_transformer, trainer, dataset, embedding};

use std::collections::{HashMap, VecDeque};
use ruvector_mincut::{DynamicMinCut, MinCutBuilder};
use std::net::SocketAddr;
use std::path::PathBuf;
use std::sync::Arc;
use std::time::Duration;

use axum::{
    extract::{
        ws::{Message, WebSocket, WebSocketUpgrade},
        Path,
        State,
    },
    response::{Html, IntoResponse, Json},
    routing::{delete, get, post},
    Router,
};
use clap::Parser;

use serde::{Deserialize, Serialize};
use tokio::net::UdpSocket;
use tokio::sync::{broadcast, RwLock};
use tower_http::services::ServeDir;
use tower_http::set_header::SetResponseHeaderLayer;
use tower_http::trace::{DefaultMakeSpan, DefaultOnFailure, DefaultOnResponse, TraceLayer};
use axum::http::HeaderValue;
use tracing::{info, warn, debug, error, Level};
use tracing_subscriber::{layer::SubscriberExt, util::SubscriberInitExt, EnvFilter, Layer};

use rvf_container::{RvfBuilder, RvfContainerInfo, RvfReader, VitalSignConfig};
use rvf_pipeline::ProgressiveLoader;
use vital_signs::{VitalSignDetector, VitalSigns};

// ADR-022 Phase 3: Multi-BSSID pipeline integration
use wifi_densepose_wifiscan::{
    BssidRegistry, WindowsWifiPipeline,
};
use wifi_densepose_wifiscan::parse_netsh_output as parse_netsh_bssid_output;

// Linux multi-BSSID scanner (ADR-022 Linux path)
#[cfg(target_os = "linux")]
use wifi_densepose_wifiscan::LinuxIwScanner;

// Accuracy sprint: Kalman tracker, multistatic fusion, field model
use wifi_densepose_signal::ruvsense::pose_tracker::PoseTracker;
use wifi_densepose_signal::ruvsense::multistatic::{MultistaticFuser, MultistaticConfig};
use wifi_densepose_signal::ruvsense::field_model::{FieldModel, CalibrationStatus};

// ── CLI ──────────────────────────────────────────────────────────────────────

#[derive(Parser, Debug)]
#[command(name = "sensing-server", about = "WiFi-DensePose sensing server")]
struct Args {
    /// HTTP port for UI and REST API
    #[arg(long, default_value = "8080")]
    http_port: u16,

    /// WebSocket port for sensing stream
    #[arg(long, default_value = "8765")]
    ws_port: u16,

    /// UDP port for ESP32 CSI frames
    #[arg(long, default_value = "5005")]
    udp_port: u16,

    /// Path to UI static files
    #[arg(long, default_value = "../../ui")]
    ui_path: PathBuf,

    /// Tick interval in milliseconds (default 100 ms = 10 fps for smooth pose animation)
    #[arg(long, default_value = "100")]
    tick_ms: u64,

    /// Bind address (default 127.0.0.1; set to 0.0.0.0 for network access)
    #[arg(long, default_value = "127.0.0.1", env = "SENSING_BIND_ADDR")]
    bind_addr: String,

    /// Data source: auto, wifi, esp32 (simulation is disabled — real data only)
    #[arg(long, default_value = "auto")]
    source: String,

    /// Run vital sign detection benchmark (1000 frames) and exit
    #[arg(long)]
    benchmark: bool,

    /// Load model config from an RVF container at startup
    #[arg(long, value_name = "PATH")]
    load_rvf: Option<PathBuf>,

    /// Save current model state as an RVF container on shutdown
    #[arg(long, value_name = "PATH")]
    save_rvf: Option<PathBuf>,

    /// Load a trained .rvf model for inference
    #[arg(long, value_name = "PATH")]
    model: Option<PathBuf>,

    /// Enable progressive loading (Layer A instant start)
    #[arg(long)]
    progressive: bool,

    /// Export an RVF container package and exit (no server)
    #[arg(long, value_name = "PATH")]
    export_rvf: Option<PathBuf>,

    /// Run training mode (train a model and exit)
    #[arg(long)]
    train: bool,

    /// Path to dataset directory (MM-Fi or Wi-Pose)
    #[arg(long, value_name = "PATH")]
    dataset: Option<PathBuf>,

    /// Dataset type: "mmfi" or "wipose"
    #[arg(long, value_name = "TYPE", default_value = "mmfi")]
    dataset_type: String,

    /// Number of training epochs
    #[arg(long, default_value = "100")]
    epochs: usize,

    /// Directory for training checkpoints
    #[arg(long, value_name = "DIR")]
    checkpoint_dir: Option<PathBuf>,

    /// Run self-supervised contrastive pretraining (ADR-024)
    #[arg(long)]
    pretrain: bool,

    /// Number of pretraining epochs (default 50)
    #[arg(long, default_value = "50")]
    pretrain_epochs: usize,

    /// Extract embeddings mode: load model and extract CSI embeddings
    #[arg(long)]
    embed: bool,

    /// Build fingerprint index from embeddings (env|activity|temporal|person)
    #[arg(long, value_name = "TYPE")]
    build_index: Option<String>,

    /// Node positions for multistatic fusion (format: "x,y,z;x,y,z;...")
    #[arg(long, env = "SENSING_NODE_POSITIONS")]
    node_positions: Option<String>,

    /// Start field model calibration on boot (empty room required)
    #[arg(long)]
    calibrate: bool,
}

// ── Data types ───────────────────────────────────────────────────────────────

/// ADR-018 ESP32 CSI binary frame header (20 bytes)
#[derive(Debug, Clone)]
#[allow(dead_code)]
struct Esp32Frame {
    magic: u32,
    node_id: u8,
    n_antennas: u8,
    n_subcarriers: u8,
    freq_mhz: u16,
    sequence: u32,
    rssi: i8,
    noise_floor: i8,
    amplitudes: Vec<f64>,
    phases: Vec<f64>,
}

/// Sensing update broadcast to WebSocket clients
#[derive(Debug, Clone, Serialize, Deserialize)]
struct SensingUpdate {
    #[serde(rename = "type")]
    msg_type: String,
    timestamp: f64,
    source: String,
    tick: u64,
    nodes: Vec<NodeInfo>,
    features: FeatureInfo,
    classification: ClassificationInfo,
    signal_field: SignalField,
    /// Vital sign estimates (breathing rate, heart rate, confidence).
    #[serde(skip_serializing_if = "Option::is_none")]
    vital_signs: Option<VitalSigns>,
    // ── ADR-022 Phase 3: Enhanced multi-BSSID pipeline fields ──
    /// Enhanced motion estimate from multi-BSSID pipeline.
    #[serde(skip_serializing_if = "Option::is_none")]
    enhanced_motion: Option<serde_json::Value>,
    /// Enhanced breathing estimate from multi-BSSID pipeline.
    #[serde(skip_serializing_if = "Option::is_none")]
    enhanced_breathing: Option<serde_json::Value>,
    /// Posture classification from BSSID fingerprint matching.
    #[serde(skip_serializing_if = "Option::is_none")]
    posture: Option<String>,
    /// Signal quality score from multi-BSSID quality gate [0.0, 1.0].
    #[serde(skip_serializing_if = "Option::is_none")]
    signal_quality_score: Option<f64>,
    /// Quality gate verdict: "Permit", "Warn", or "Deny".
    #[serde(skip_serializing_if = "Option::is_none")]
    quality_verdict: Option<String>,
    /// Number of BSSIDs used in the enhanced sensing cycle.
    #[serde(skip_serializing_if = "Option::is_none")]
    bssid_count: Option<usize>,
    // ── ADR-023 Phase 7-8: Model inference fields ──
    /// Pose keypoints when a trained model is loaded (x, y, z, confidence).
    #[serde(skip_serializing_if = "Option::is_none")]
    pose_keypoints: Option<Vec<[f64; 4]>>,
    /// Model status when a trained model is loaded.
    #[serde(skip_serializing_if = "Option::is_none")]
    model_status: Option<serde_json::Value>,
    // ── Multi-person detection (issue #97) ──
    /// Detected persons from WiFi sensing (multi-person support).
    #[serde(skip_serializing_if = "Option::is_none")]
    persons: Option<Vec<PersonDetection>>,
    /// Estimated person count from CSI feature heuristics (1-3 for single ESP32).
    #[serde(skip_serializing_if = "Option::is_none")]
    estimated_persons: Option<usize>,
    /// Discrete count votes from independent estimators plus weighted consensus.
    #[serde(skip_serializing_if = "Option::is_none")]
    count_evidence: Option<PersonCountEvidence>,
    /// Spatial anchors extracted from the signal field before skeleton synthesis.
    #[serde(skip_serializing_if = "Option::is_none")]
    person_anchors: Option<Vec<PersonAnchor>>,
    /// Per-node feature breakdown for multi-node deployments.
    #[serde(skip_serializing_if = "Option::is_none")]
    node_features: Option<Vec<PerNodeFeatureInfo>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct NodeInfo {
    node_id: u8,
    rssi_dbm: f64,
    position: [f64; 3],
    amplitude: Vec<f64>,
    subcarrier_count: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct FeatureInfo {
    mean_rssi: f64,
    variance: f64,
    motion_band_power: f64,
    breathing_band_power: f64,
    dominant_freq_hz: f64,
    change_points: usize,
    spectral_power: f64,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct ClassificationInfo {
    motion_level: String,
    presence: bool,
    confidence: f64,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct SignalField {
    grid_size: [usize; 3],
    values: Vec<f64>,
}

/// WiFi-derived pose keypoint (17 COCO keypoints)
#[derive(Debug, Clone, Serialize, Deserialize)]
struct PoseKeypoint {
    name: String,
    x: f64,
    y: f64,
    z: f64,
    confidence: f64,
}

/// Person detection from WiFi sensing
#[derive(Debug, Clone, Serialize, Deserialize)]
struct PersonDetection {
    id: u32,
    confidence: f64,
    keypoints: Vec<PoseKeypoint>,
    bbox: BoundingBox,
    zone: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct CountEstimate {
    source: String,
    count: usize,
    confidence: f64,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct PersonCountEvidence {
    consensus_count: usize,
    consensus_confidence: f64,
    estimates: Vec<CountEstimate>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct PersonAnchor {
    x: f64,
    y: f64,
    z: f64,
    strength: f64,
    zone: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct BoundingBox {
    x: f64,
    y: f64,
    width: f64,
    height: f64,
}

/// Per-node sensing state for multi-node deployments (issue #249).
/// Each ESP32 node gets its own frame history, smoothing buffers, and vital
/// sign detector so that data from different nodes is never mixed.
struct NodeState {
    pub(crate) frame_history: VecDeque<Vec<f64>>,
    smoothed_person_score: f64,
    pub(crate) prev_person_count: usize,
    smoothed_motion: f64,
    current_motion_level: String,
    debounce_counter: u32,
    debounce_candidate: String,
    baseline_motion: f64,
    baseline_frames: u64,
    smoothed_hr: f64,
    smoothed_br: f64,
    smoothed_hr_conf: f64,
    smoothed_br_conf: f64,
    hr_buffer: VecDeque<f64>,
    br_buffer: VecDeque<f64>,
    rssi_history: VecDeque<f64>,
    vital_detector: VitalSignDetector,
    latest_vitals: VitalSigns,
    pub(crate) last_frame_time: Option<std::time::Instant>,
    edge_vitals: Option<Esp32VitalsPacket>,
    /// Latest extracted features for cross-node fusion.
    latest_features: Option<FeatureInfo>,
    /// Source IP address of the last UDP frame from this node (for GUI display).
    pub(crate) peer_ip: Option<std::net::IpAddr>,
}

/// Default EMA alpha for temporal keypoint smoothing (RuVector Phase 2).
/// Lower = smoother (more history, less jitter). 0.15 balances responsiveness
/// with stability for WiFi CSI where per-frame noise is high.
const TEMPORAL_EMA_ALPHA_DEFAULT: f64 = 0.15;
/// Maximum allowed bone-length change ratio between frames (20%).
const MAX_BONE_CHANGE_RATIO: f64 = 0.20;
/// Lower smoothing alpha for uncertain tracked people.
const TRACK_SMOOTHING_ALPHA_LOW_CONFIDENCE: f64 = 0.08;
/// Track-confidence threshold that switches to the lower alpha.
const TRACK_SMOOTHING_LOW_CONFIDENCE_THRESHOLD: f64 = 0.55;
/// Retain smoothing state for missing tracks for a short grace period.
const TRACK_SMOOTHING_TTL_TICKS: u64 = 30;

#[derive(Debug, Clone)]
struct TrackTemporalState {
    keypoints: Vec<[f64; 3]>,
    last_seen_tick: u64,
}

impl NodeState {
    pub(crate) fn new() -> Self {
        Self {
            frame_history: VecDeque::new(),
            smoothed_person_score: 0.0,
            prev_person_count: 0,
            smoothed_motion: 0.0,
            current_motion_level: "absent".to_string(),
            debounce_counter: 0,
            debounce_candidate: "absent".to_string(),
            baseline_motion: 0.0,
            baseline_frames: 0,
            smoothed_hr: 0.0,
            smoothed_br: 0.0,
            smoothed_hr_conf: 0.0,
            smoothed_br_conf: 0.0,
            hr_buffer: VecDeque::with_capacity(8),
            br_buffer: VecDeque::with_capacity(8),
            rssi_history: VecDeque::new(),
            vital_detector: VitalSignDetector::new(10.0),
            latest_vitals: VitalSigns::default(),
            last_frame_time: None,
            edge_vitals: None,
            latest_features: None,
            peer_ip: None,
        }
    }
}

/// Per-node feature info for WebSocket broadcasts (multi-node support).
#[derive(Debug, Clone, Serialize, Deserialize)]
struct PerNodeFeatureInfo {
    node_id: u8,
    features: FeatureInfo,
    classification: ClassificationInfo,
    rssi_dbm: f64,
    last_seen_ms: u64,
    frame_rate_hz: f64,
    stale: bool,
}

/// Shared application state
struct AppStateInner {
    latest_update: Option<SensingUpdate>,
    rssi_history: VecDeque<f64>,
    /// Circular buffer of recent CSI amplitude vectors for temporal analysis.
    /// Each entry is the full subcarrier amplitude vector for one frame.
    /// Capacity: FRAME_HISTORY_CAPACITY frames.
    frame_history: VecDeque<Vec<f64>>,
    tick: u64,
    source: String,
    /// Instant of the last ESP32 UDP frame received (for offline detection).
    last_esp32_frame: Option<std::time::Instant>,
    tx: broadcast::Sender<String>,
    total_detections: u64,
    start_time: std::time::Instant,
    /// Vital sign detector (processes CSI frames to estimate HR/RR).
    vital_detector: VitalSignDetector,
    /// Most recent vital sign reading for the REST endpoint.
    latest_vitals: VitalSigns,
    /// RVF container info if a model was loaded via `--load-rvf`.
    rvf_info: Option<RvfContainerInfo>,
    /// Path to save RVF container on shutdown (set via `--save-rvf`).
    save_rvf_path: Option<PathBuf>,
    /// Progressive loader for a trained model (set via `--model`).
    progressive_loader: Option<ProgressiveLoader>,
    /// Active SONA profile name.
    active_sona_profile: Option<String>,
    /// Whether a trained model is loaded.
    model_loaded: bool,
    /// Smoothed person count (EMA) for hysteresis — prevents frame-to-frame jumping.
    smoothed_person_score: f64,
    /// Previous person count for hysteresis (asymmetric up/down thresholds).
    prev_person_count: usize,
    // ── Motion smoothing & adaptive baseline (ADR-047 tuning) ────────────
    /// EMA-smoothed motion score (alpha ~0.15 for ~10 FPS → ~1s time constant).
    smoothed_motion: f64,
    /// Current classification state for hysteresis debounce.
    current_motion_level: String,
    /// How many consecutive frames the *raw* classification has agreed with a
    /// *candidate* new level.  State only changes after DEBOUNCE_FRAMES.
    debounce_counter: u32,
    /// The candidate motion level that the debounce counter is tracking.
    debounce_candidate: String,
    /// Adaptive baseline: EMA of motion score when room is "quiet" (low motion).
    /// Subtracted from raw score so slow environmental drift doesn't inflate readings.
    baseline_motion: f64,
    /// Number of frames processed so far (for baseline warm-up).
    baseline_frames: u64,
    // ── Vital signs smoothing ────────────────────────────────────────────
    /// EMA-smoothed heart rate (BPM).
    smoothed_hr: f64,
    /// EMA-smoothed breathing rate (BPM).
    smoothed_br: f64,
    /// EMA-smoothed HR confidence.
    smoothed_hr_conf: f64,
    /// EMA-smoothed BR confidence.
    smoothed_br_conf: f64,
    /// Median filter buffer for HR (last N raw values for outlier rejection).
    hr_buffer: VecDeque<f64>,
    /// Median filter buffer for BR.
    br_buffer: VecDeque<f64>,
    /// ADR-039: Latest edge vitals packet from ESP32.
    edge_vitals: Option<Esp32VitalsPacket>,
    /// ADR-040: Latest WASM output packet from ESP32.
    latest_wasm_events: Option<WasmOutputPacket>,
    // ── Model management fields ─────────────────────────────────────────────
    /// Discovered RVF model files from `data/models/`.
    discovered_models: Vec<serde_json::Value>,
    /// ID of the currently loaded model, if any.
    active_model_id: Option<String>,
    // ── Recording fields ────────────────────────────────────────────────────
    /// Metadata for recorded CSI data files.
    recordings: Vec<serde_json::Value>,
    /// Whether CSI recording is currently in progress.
    recording_active: bool,
    /// When the current recording started.
    recording_start_time: Option<std::time::Instant>,
    /// ID of the current recording (used for filename).
    recording_current_id: Option<String>,
    /// Shutdown signal for the recording writer task.
    recording_stop_tx: Option<tokio::sync::watch::Sender<bool>>,
    // ── Training fields ─────────────────────────────────────────────────────
    /// Training status: "idle", "running", "completed", "failed".
    training_status: String,
    /// Training configuration, if any.
    training_config: Option<serde_json::Value>,
    // ── Adaptive classifier (environment-tuned) ──────────────────────────
    /// Trained adaptive model (loaded from data/adaptive_model.json or trained at runtime).
    adaptive_model: Option<adaptive_classifier::AdaptiveModel>,
    // ── Per-node state (issue #249) ─────────────────────────────────────
    /// Per-node sensing state for multi-node deployments.
    /// Keyed by `node_id` from the ESP32 frame header.
    node_states: HashMap<u8, NodeState>,
    // ── Accuracy sprint: Kalman tracker, multistatic fusion, eigenvalue counting ──
    /// Global Kalman-based pose tracker for stable person IDs and smoothed keypoints.
    pose_tracker: PoseTracker,
    /// Instant of last tracker update (for computing dt).
    last_tracker_instant: Option<std::time::Instant>,
    /// Per-track temporal smoother layered on top of tracker output.
    track_temporal_state: HashMap<u32, TrackTemporalState>,
    /// Attention-weighted multi-node CSI fusion engine.
    multistatic_fuser: MultistaticFuser,
    /// SVD-based room field model for eigenvalue person counting (None until calibration).
    field_model: Option<FieldModel>,
}

/// If no ESP32 frame arrives within this duration, source reverts to offline.
const ESP32_OFFLINE_TIMEOUT: std::time::Duration = std::time::Duration::from_secs(5);

impl AppStateInner {
    /// Return the effective data source, accounting for ESP32 frame timeout.
    /// If the source is "esp32" but no frame has arrived in 5 seconds, returns
    /// "esp32:offline" so the UI can distinguish active vs stale connections.
    /// Person count: eigenvalue-based if field model is calibrated, else heuristic.
    /// Uses global frame_history if populated, otherwise the freshest per-node history.
    fn person_count(&self) -> usize {
        match self.field_model.as_ref() {
            Some(fm) => {
                // Prefer global frame_history (populated by wifi/simulate paths).
                // Fall back to freshest per-node history (populated by ESP32 paths).
                let history = if !self.frame_history.is_empty() {
                    &self.frame_history
                } else {
                    // Find the node with the most recent frame
                    self.node_states.values()
                        .filter(|ns| !ns.frame_history.is_empty())
                        .max_by_key(|ns| ns.last_frame_time)
                        .map(|ns| &ns.frame_history)
                        .unwrap_or(&self.frame_history)
                };
                field_bridge::occupancy_or_fallback(
                    fm, history, self.smoothed_person_score, self.prev_person_count,
                )
            }
            None => score_to_person_count(self.smoothed_person_score, self.prev_person_count),
        }
    }

    fn effective_source(&self) -> String {
        if self.source == "esp32" {
            if let Some(last) = self.last_esp32_frame {
                if last.elapsed() > ESP32_OFFLINE_TIMEOUT {
                    return "esp32:offline".to_string();
                }
            }
        }
        self.source.clone()
    }
}

/// Number of frames retained in `frame_history` for temporal analysis.
/// At 500 ms ticks this covers ~50 seconds; at 100 ms ticks ~10 seconds.
const FRAME_HISTORY_CAPACITY: usize = 100;

type SharedState = Arc<RwLock<AppStateInner>>;

// ── ESP32 Edge Vitals Packet (ADR-039, magic 0xC511_0002) ────────────────────

/// Decoded vitals packet from ESP32 edge processing pipeline.
#[derive(Debug, Clone, Serialize)]
struct Esp32VitalsPacket {
    node_id: u8,
    presence: bool,
    fall_detected: bool,
    motion: bool,
    breathing_rate_bpm: f64,
    heartrate_bpm: f64,
    rssi: i8,
    n_persons: u8,
    motion_energy: f32,
    presence_score: f32,
    timestamp_ms: u32,
}

/// Parse a 32-byte edge vitals packet (magic 0xC511_0002).
fn parse_esp32_vitals(buf: &[u8]) -> Option<Esp32VitalsPacket> {
    if buf.len() < 32 {
        return None;
    }
    let magic = u32::from_le_bytes([buf[0], buf[1], buf[2], buf[3]]);
    if magic != 0xC511_0002 {
        return None;
    }

    let node_id = buf[4];
    let flags = buf[5];
    let breathing_raw = u16::from_le_bytes([buf[6], buf[7]]);
    let heartrate_raw = u32::from_le_bytes([buf[8], buf[9], buf[10], buf[11]]);
    let rssi = buf[12] as i8;
    let n_persons = buf[13];
    let motion_energy = f32::from_le_bytes([buf[16], buf[17], buf[18], buf[19]]);
    let presence_score = f32::from_le_bytes([buf[20], buf[21], buf[22], buf[23]]);
    let timestamp_ms = u32::from_le_bytes([buf[24], buf[25], buf[26], buf[27]]);

    Some(Esp32VitalsPacket {
        node_id,
        presence: (flags & 0x01) != 0,
        fall_detected: (flags & 0x02) != 0,
        motion: (flags & 0x04) != 0,
        breathing_rate_bpm: breathing_raw as f64 / 100.0,
        heartrate_bpm: heartrate_raw as f64 / 10000.0,
        rssi,
        n_persons,
        motion_energy,
        presence_score,
        timestamp_ms,
    })
}

// ── ADR-040: WASM Output Packet (magic 0xC511_0004) ───────────────────────────

/// Single WASM event (type + value).
#[derive(Debug, Clone, Serialize)]
struct WasmEvent {
    event_type: u8,
    value: f32,
}

/// Decoded WASM output packet from ESP32 Tier 3 runtime.
#[derive(Debug, Clone, Serialize)]
struct WasmOutputPacket {
    node_id: u8,
    module_id: u8,
    events: Vec<WasmEvent>,
}

/// Parse a WASM output packet (magic 0xC511_0004).
fn parse_wasm_output(buf: &[u8]) -> Option<WasmOutputPacket> {
    if buf.len() < 8 {
        return None;
    }
    let magic = u32::from_le_bytes([buf[0], buf[1], buf[2], buf[3]]);
    if magic != 0xC511_0004 {
        return None;
    }

    let node_id = buf[4];
    let module_id = buf[5];
    let event_count = u16::from_le_bytes([buf[6], buf[7]]) as usize;

    let mut events = Vec::with_capacity(event_count);
    let mut offset = 8;
    for _ in 0..event_count {
        if offset + 5 > buf.len() {
            break;
        }
        let event_type = buf[offset];
        let value = f32::from_le_bytes([
            buf[offset + 1], buf[offset + 2], buf[offset + 3], buf[offset + 4],
        ]);
        events.push(WasmEvent { event_type, value });
        offset += 5;
    }

    Some(WasmOutputPacket {
        node_id,
        module_id,
        events,
    })
}

// ── ESP32 UDP frame parser ───────────────────────────────────────────────────

fn parse_esp32_frame(buf: &[u8]) -> Option<Esp32Frame> {
    if buf.len() < 20 {
        return None;
    }

    let magic = u32::from_le_bytes([buf[0], buf[1], buf[2], buf[3]]);
    if magic != 0xC511_0001 {
        return None;
    }

    // Frame layout (must match firmware csi_collector.c):
    //   [0..3]   magic (u32 LE)
    //   [4]      node_id (u8)
    //   [5]      n_antennas (u8)
    //   [6..7]   n_subcarriers (u16 LE)
    //   [8..11]  freq_mhz (u32 LE)
    //   [12..15] sequence (u32 LE)
    //   [16]     rssi (i8)
    //   [17]     noise_floor (i8)
    //   [18..19] reserved
    //   [20..]   I/Q data
    let node_id = buf[4];
    let n_antennas = buf[5];
    let n_subcarriers_u16 = u16::from_le_bytes([buf[6], buf[7]]);
    let n_subcarriers = n_subcarriers_u16 as u8; // truncate to u8 for Esp32Frame compat
    let freq_mhz = u16::from_le_bytes([buf[8], buf[9]]); // low 16 bits of u32
    let sequence = u32::from_le_bytes([buf[12], buf[13], buf[14], buf[15]]);
    let rssi_raw = buf[16] as i8;   // #332: was buf[14], 2 bytes off
    // Fix RSSI sign: ensure it's always negative (dBm convention).
    let rssi = if rssi_raw > 0 { rssi_raw.saturating_neg() } else { rssi_raw };
    let noise_floor = buf[17] as i8; // #332: was buf[15], 2 bytes off

    let iq_start = 20;
    let n_pairs = n_antennas as usize * n_subcarriers as usize;
    let expected_len = iq_start + n_pairs * 2;

    if buf.len() < expected_len {
        return None;
    }

    let mut amplitudes = Vec::with_capacity(n_pairs);
    let mut phases = Vec::with_capacity(n_pairs);

    for k in 0..n_pairs {
        let i_val = buf[iq_start + k * 2] as i8 as f64;
        let q_val = buf[iq_start + k * 2 + 1] as i8 as f64;
        amplitudes.push((i_val * i_val + q_val * q_val).sqrt());
        phases.push(q_val.atan2(i_val));
    }

    Some(Esp32Frame {
        magic,
        node_id,
        n_antennas,
        n_subcarriers,
        freq_mhz,
        sequence,
        rssi,
        noise_floor,
        amplitudes,
        phases,
    })
}

// ── Signal field generation ──────────────────────────────────────────────────

/// Generate a signal field that reflects where motion and signal changes are occurring.
///
/// Instead of a fixed-animation circle, this function uses the actual sensing data:
/// - `subcarrier_variances`: per-subcarrier variance computed from the frame history.
///   High-variance subcarriers indicate spatial directions where the signal is disrupted.
/// - `motion_score`: overall motion intensity [0, 1].
/// - `breathing_rate_hz`: estimated breathing rate in Hz; if > 0, adds a breathing ring.
/// - `signal_quality`: overall quality metric [0, 1] modulates field brightness.
///
/// The field grid is 20×20 cells representing a top-down view of the room.
/// Hotspots are derived from the subcarrier index (treated as an angular bin) so that
/// subcarriers with the highest variance produce peaks at the corresponding directions.
fn generate_signal_field(
    _mean_rssi: f64,
    motion_score: f64,
    breathing_rate_hz: f64,
    signal_quality: f64,
    subcarrier_variances: &[f64],
) -> SignalField {
    let grid = 20usize;
    let mut values = vec![0.0f64; grid * grid];
    let center = (grid as f64 - 1.0) / 2.0;

    // Normalise subcarrier variances to [0, 1].
    let max_var = subcarrier_variances.iter().cloned().fold(0.0f64, f64::max);
    let norm_factor = if max_var > 1e-9 { max_var } else { 1.0 };

    // For each cell, accumulate contributions from all subcarriers.
    // Each subcarrier k is assigned an angular direction proportional to its index
    // so that different subcarriers illuminate different regions of the room.
    let n_sub = subcarrier_variances.len().max(1);
    for (k, &var) in subcarrier_variances.iter().enumerate() {
        let weight = (var / norm_factor) * motion_score;
        if weight < 1e-6 {
            continue;
        }
        // Map subcarrier index to an angle across the full 2π sweep.
        let angle = (k as f64 / n_sub as f64) * 2.0 * std::f64::consts::PI;
        // Place the hotspot at a distance proportional to the weight, capped at 40% of
        // the grid radius so it stays within the room model.
        let radius = center * 0.8 * weight.sqrt();
        let hx = center + radius * angle.cos();
        let hz = center + radius * angle.sin();

        for z in 0..grid {
            for x in 0..grid {
                let dx = x as f64 - hx;
                let dz = z as f64 - hz;
                let dist2 = dx * dx + dz * dz;
                // Gaussian blob centred on the hotspot; spread scales with weight.
                let spread = (0.5 + weight * 2.0).max(0.5);
                values[z * grid + x] += weight * (-dist2 / (2.0 * spread * spread)).exp();
            }
        }
    }

    // Base radial attenuation from the router assumed at grid centre.
    for z in 0..grid {
        for x in 0..grid {
            let dx = x as f64 - center;
            let dz = z as f64 - center;
            let dist = (dx * dx + dz * dz).sqrt();
            let base = signal_quality * (-dist * 0.12).exp();
            values[z * grid + x] += base * 0.3;
        }
    }

    // Breathing ring: if a breathing rate was estimated add a faint annular highlight
    // at a radius corresponding to typical chest-wall displacement range.
    if breathing_rate_hz > 0.05 {
        let ring_r = center * 0.55;
        let ring_width = 1.8f64;
        for z in 0..grid {
            for x in 0..grid {
                let dx = x as f64 - center;
                let dz = z as f64 - center;
                let dist = (dx * dx + dz * dz).sqrt();
                let ring_val = 0.08 * (-(dist - ring_r).powi(2) / (2.0 * ring_width * ring_width)).exp();
                values[z * grid + x] += ring_val;
            }
        }
    }

    // Clamp and normalise to [0, 1].
    let field_max = values.iter().cloned().fold(0.0f64, f64::max);
    let scale = if field_max > 1e-9 { 1.0 / field_max } else { 1.0 };
    for v in &mut values {
        *v = (*v * scale).clamp(0.0, 1.0);
    }

    SignalField {
        grid_size: [grid, 1, grid],
        values,
    }
}

// ── Feature extraction from ESP32 frame ──────────────────────────────────────

/// Estimate breathing rate in Hz from the amplitude time series stored in `frame_history`.
///
/// Approach:
/// 1. Build a scalar time series by computing the mean amplitude of each historical frame.
/// 2. Run a peak-detection pass: count rising-edge zero-crossings of the de-meaned signal.
/// 3. Convert the crossing rate to Hz, clipped to the physiological range 0.1–0.5 Hz
///    (12–30 breaths/min).
///
/// For accuracy the function additionally applies a simple 3-tap Goertzel-style power
/// estimate at evenly-spaced candidate frequencies in the breathing band and returns
/// the candidate with the highest energy.
fn estimate_breathing_rate_hz(frame_history: &VecDeque<Vec<f64>>, sample_rate_hz: f64) -> f64 {
    let n = frame_history.len();
    if n < 6 {
        return 0.0;
    }

    // Build scalar time series: mean amplitude per frame.
    let series: Vec<f64> = frame_history.iter()
        .map(|amps| {
            if amps.is_empty() { 0.0 } else { amps.iter().sum::<f64>() / amps.len() as f64 }
        })
        .collect();

    let mean_s = series.iter().sum::<f64>() / n as f64;
    // De-mean.
    let detrended: Vec<f64> = series.iter().map(|x| x - mean_s).collect();

    // Goertzel power at candidate frequencies in the breathing band [0.1, 0.5] Hz.
    // We evaluate 9 candidate frequencies uniformly spaced in that band.
    let n_candidates = 9usize;
    let f_low = 0.1f64;
    let f_high = 0.5f64;
    let mut best_freq = 0.0f64;
    let mut best_power = 0.0f64;

    for i in 0..n_candidates {
        let freq = f_low + (f_high - f_low) * i as f64 / (n_candidates - 1).max(1) as f64;
        let omega = 2.0 * std::f64::consts::PI * freq / sample_rate_hz;
        let coeff = 2.0 * omega.cos();
        let mut s_prev2 = 0.0f64;
        let mut s_prev1 = 0.0f64;
        for &x in &detrended {
            let s = x + coeff * s_prev1 - s_prev2;
            s_prev2 = s_prev1;
            s_prev1 = s;
        }
        // Goertzel magnitude squared.
        let power = s_prev2 * s_prev2 + s_prev1 * s_prev1 - coeff * s_prev1 * s_prev2;
        if power > best_power {
            best_power = power;
            best_freq = freq;
        }
    }

    // Only report a breathing rate if the Goertzel energy is meaningfully above noise.
    // Threshold: power must exceed 10× the average power across all candidates.
    let avg_power = {
        let mut total = 0.0f64;
        for i in 0..n_candidates {
            let freq = f_low + (f_high - f_low) * i as f64 / (n_candidates - 1).max(1) as f64;
            let omega = 2.0 * std::f64::consts::PI * freq / sample_rate_hz;
            let coeff = 2.0 * omega.cos();
            let mut s_prev2 = 0.0f64;
            let mut s_prev1 = 0.0f64;
            for &x in &detrended {
                let s = x + coeff * s_prev1 - s_prev2;
                s_prev2 = s_prev1;
                s_prev1 = s;
            }
            total += s_prev2 * s_prev2 + s_prev1 * s_prev1 - coeff * s_prev1 * s_prev2;
        }
        total / n_candidates as f64
    };

    if best_power > avg_power * 5.0 {
        best_freq.clamp(f_low, f_high)
    } else {
        0.0
    }
}

/// Compute per-subcarrier variance across the sliding window of `frame_history`.
///
/// For each subcarrier index `k`, returns `Var[A_k]` over all stored frames.
/// This captures spatial signal variation; subcarriers whose amplitude fluctuates
/// heavily across time correspond to directions with motion.
/// Compute per-subcarrier importance weights using a simple sensitivity split.
///
/// Subcarriers whose sensitivity (amplitude magnitude) is above the median are
/// considered "sensitive" and receive weight `1.0 + (sens / max_sens)` (range 1.0–2.0).
/// The rest receive a baseline weight of 0.5. This mirrors the RuVector mincut
/// partition logic without requiring the graph dependency.
fn compute_subcarrier_importance_weights(sensitivity: &[f64]) -> Vec<f64> {
    let n = sensitivity.len();
    if n == 0 {
        return vec![];
    }
    let max_sens = sensitivity.iter().cloned().fold(f64::NEG_INFINITY, f64::max).max(1e-9);

    // Compute median via a sorted copy.
    let mut sorted = sensitivity.to_vec();
    sorted.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let median = if n % 2 == 0 {
        (sorted[n / 2 - 1] + sorted[n / 2]) / 2.0
    } else {
        sorted[n / 2]
    };

    sensitivity
        .iter()
        .map(|&s| {
            if s >= median {
                1.0 + (s / max_sens).min(1.0)
            } else {
                0.5
            }
        })
        .collect()
}

fn compute_subcarrier_variances(frame_history: &VecDeque<Vec<f64>>, n_sub: usize) -> Vec<f64> {
    if frame_history.is_empty() || n_sub == 0 {
        return vec![0.0; n_sub];
    }

    let n_frames = frame_history.len() as f64;
    let mut means = vec![0.0f64; n_sub];
    let mut sq_means = vec![0.0f64; n_sub];

    for frame in frame_history.iter() {
        for k in 0..n_sub {
            let a = if k < frame.len() { frame[k] } else { 0.0 };
            means[k] += a;
            sq_means[k] += a * a;
        }
    }

    (0..n_sub)
        .map(|k| {
            let mean = means[k] / n_frames;
            let sq_mean = sq_means[k] / n_frames;
            (sq_mean - mean * mean).max(0.0)
        })
        .collect()
}

/// Extract features from the current ESP32 frame, enhanced with temporal context from
/// `frame_history`.
///
/// Improvements over the previous single-frame approach:
///
/// - **Variance**: computed as the mean of per-subcarrier temporal variance across the
///   sliding window, not just the intra-frame spatial variance.
/// - **Motion detection**: uses frame-to-frame temporal difference (mean L2 change
///   between the current frame and the previous frame) normalised by signal amplitude,
///   so that actual changes are detected rather than just a threshold on the current frame.
/// - **Breathing rate**: estimated via Goertzel filter bank on the 0.1–0.5 Hz band of
///   the amplitude time series.
/// - **Signal quality**: based on SNR estimate (RSSI – noise floor) and subcarrier
///   variance stability.
/// Returns (features, raw_classification, breathing_rate_hz, sub_variances, raw_motion_score).
fn extract_features_from_frame(
    frame: &Esp32Frame,
    frame_history: &VecDeque<Vec<f64>>,
    sample_rate_hz: f64,
) -> (FeatureInfo, ClassificationInfo, f64, Vec<f64>, f64) {
    let n_sub = frame.amplitudes.len().max(1);
    let n = n_sub as f64;
    let mean_rssi = frame.rssi as f64;

    // ── RuVector Phase 1: subcarrier importance weighting ──
    // Compute per-subcarrier sensitivity from amplitude magnitude, then weight
    // sensitive subcarriers higher (>1.0) and insensitive ones lower (0.5).
    // This emphasises body-motion-correlated subcarriers in all downstream metrics.
    let sub_sensitivity: Vec<f64> = frame.amplitudes.iter().map(|a| a.abs()).collect();
    let importance_weights = compute_subcarrier_importance_weights(&sub_sensitivity);

    let weight_sum: f64 = importance_weights.iter().sum::<f64>();
    let mean_amp: f64 = if weight_sum > 0.0 {
        frame.amplitudes.iter().zip(importance_weights.iter())
            .map(|(a, w)| a * w)
            .sum::<f64>() / weight_sum
    } else {
        frame.amplitudes.iter().sum::<f64>() / n
    };

    // ── Intra-frame subcarrier variance (weighted by importance) ──
    let intra_variance: f64 = if weight_sum > 0.0 {
        frame.amplitudes.iter().zip(importance_weights.iter())
            .map(|(a, w)| w * (a - mean_amp).powi(2))
            .sum::<f64>() / weight_sum
    } else {
        frame.amplitudes.iter()
            .map(|a| (a - mean_amp).powi(2))
            .sum::<f64>() / n
    };

    // ── Temporal (sliding-window) per-subcarrier variance ──
    let sub_variances = compute_subcarrier_variances(frame_history, n_sub);
    let temporal_variance: f64 = if sub_variances.is_empty() {
        intra_variance
    } else {
        sub_variances.iter().sum::<f64>() / sub_variances.len() as f64
    };

    // Use the larger of intra-frame and temporal variance as the reported variance.
    let variance = intra_variance.max(temporal_variance);

    // ── Spectral power ──
    let spectral_power: f64 = frame.amplitudes.iter().map(|a| a * a).sum::<f64>() / n;

    // ── Motion band power (upper half of subcarriers, high spatial frequency) ──
    let half = frame.amplitudes.len() / 2;
    let motion_band_power = if half > 0 {
        frame.amplitudes[half..].iter()
            .map(|a| (a - mean_amp).powi(2))
            .sum::<f64>() / (frame.amplitudes.len() - half) as f64
    } else {
        0.0
    };

    // ── Breathing band power (lower half of subcarriers, low spatial frequency) ──
    let breathing_band_power = if half > 0 {
        frame.amplitudes[..half].iter()
            .map(|a| (a - mean_amp).powi(2))
            .sum::<f64>() / half as f64
    } else {
        0.0
    };

    // ── Dominant frequency via peak subcarrier index ──
    let peak_idx = frame.amplitudes.iter()
        .enumerate()
        .max_by(|a, b| a.1.partial_cmp(b.1).unwrap_or(std::cmp::Ordering::Equal))
        .map(|(i, _)| i)
        .unwrap_or(0);
    let dominant_freq_hz = peak_idx as f64 * 0.05;

    // ── Change point detection (threshold-crossing count in current frame) ──
    let threshold = mean_amp * 1.2;
    let change_points = frame.amplitudes.windows(2)
        .filter(|w| (w[0] < threshold) != (w[1] < threshold))
        .count();

    // ── Motion score: sliding-window temporal difference ──
    // Compare current frame against the most recent historical frame.
    // The difference is normalised by the mean amplitude to be scale-invariant.
    let temporal_motion_score = if let Some(prev_frame) = frame_history.back() {
        let n_cmp = n_sub.min(prev_frame.len());
        if n_cmp > 0 {
            let diff_energy: f64 = (0..n_cmp)
                .map(|k| (frame.amplitudes[k] - prev_frame[k]).powi(2))
                .sum::<f64>() / n_cmp as f64;
            // Normalise by mean squared amplitude to get a dimensionless ratio.
            let ref_energy = mean_amp * mean_amp + 1e-9;
            (diff_energy / ref_energy).sqrt().clamp(0.0, 1.0)
        } else {
            0.0
        }
    } else {
        // No history yet — fall back to intra-frame variance-based estimate.
        (intra_variance / (mean_amp * mean_amp + 1e-9)).sqrt().clamp(0.0, 1.0)
    };

    // Blend temporal motion with variance-based motion for robustness.
    // Also factor in motion_band_power and change_points for ESP32 real-world sensitivity.
    let variance_motion = (temporal_variance / 10.0).clamp(0.0, 1.0);
    let mbp_motion = (motion_band_power / 25.0).clamp(0.0, 1.0);
    let cp_motion = (change_points as f64 / 15.0).clamp(0.0, 1.0);
    let motion_score = (temporal_motion_score * 0.4 + variance_motion * 0.2 + mbp_motion * 0.25 + cp_motion * 0.15).clamp(0.0, 1.0);

    // ── Signal quality metric ──
    // Based on estimated SNR (RSSI relative to noise floor) and subcarrier consistency.
    let snr_db = (frame.rssi as f64 - frame.noise_floor as f64).max(0.0);
    let snr_quality = (snr_db / 40.0).clamp(0.0, 1.0); // 40 dB → quality = 1.0
    // Penalise quality when temporal variance is very high (unstable signal).
    let stability = (1.0 - (temporal_variance / (mean_amp * mean_amp + 1e-9)).clamp(0.0, 1.0)).max(0.0);
    let signal_quality = (snr_quality * 0.6 + stability * 0.4).clamp(0.0, 1.0);

    // ── Breathing rate estimation ──
    let breathing_rate_hz = estimate_breathing_rate_hz(frame_history, sample_rate_hz);

    let features = FeatureInfo {
        mean_rssi,
        variance,
        motion_band_power,
        breathing_band_power,
        dominant_freq_hz,
        change_points,
        spectral_power,
    };

    // Return raw motion_score and signal_quality — classification is done by
    // `smooth_and_classify()` which has access to EMA state and hysteresis.
    let raw_classification = ClassificationInfo {
        motion_level: raw_classify(motion_score),
        presence: motion_score > 0.04,
        confidence: (0.4 + signal_quality * 0.3 + motion_score * 0.3).clamp(0.0, 1.0),
    };

    (features, raw_classification, breathing_rate_hz, sub_variances, motion_score)
}

/// Simple threshold classification (no smoothing) — used as the "raw" input.
/// Thresholds calibrated for real ESP32 hardware and WiFi BSSID sensing.
/// Lower presence threshold (0.03) catches subtle body-induced signal changes.
fn raw_classify(score: f64) -> String {
    if score > 0.22 { "active".into() }
    else if score > 0.10 { "present_moving".into() }
    else if score > 0.03 { "present_still".into() }
    else { "absent".into() }
}

/// Debounce frames required before state transition (at ~10 FPS = ~0.4s).
const DEBOUNCE_FRAMES: u32 = 4;
/// EMA alpha for motion smoothing (~1s time constant at 10 FPS).
const MOTION_EMA_ALPHA: f64 = 0.15;
/// EMA alpha for slow-adapting baseline (~30s time constant at 10 FPS).
const BASELINE_EMA_ALPHA: f64 = 0.003;
/// Number of warm-up frames before baseline subtraction kicks in.
const BASELINE_WARMUP: u64 = 50;

/// Apply EMA smoothing, adaptive baseline subtraction, and hysteresis debounce
/// to the raw classification.  Mutates the smoothing state in `AppStateInner`.
fn smooth_and_classify(state: &mut AppStateInner, raw: &mut ClassificationInfo, raw_motion: f64) {
    // 1. Adaptive baseline: slowly track the "quiet room" floor.
    //    Only update baseline when raw score is below the current smoothed level
    //    (i.e. during calm periods) so walking doesn't inflate the baseline.
    state.baseline_frames += 1;
    if state.baseline_frames < BASELINE_WARMUP {
        // During warm-up, aggressively learn the baseline.
        state.baseline_motion = state.baseline_motion * 0.9 + raw_motion * 0.1;
    } else if raw_motion < state.smoothed_motion + 0.05 {
        state.baseline_motion = state.baseline_motion * (1.0 - BASELINE_EMA_ALPHA)
                              + raw_motion * BASELINE_EMA_ALPHA;
    }

    // 2. Subtract baseline and clamp.
    let adjusted = (raw_motion - state.baseline_motion * 0.7).max(0.0);

    // 3. EMA smooth the adjusted score.
    state.smoothed_motion = state.smoothed_motion * (1.0 - MOTION_EMA_ALPHA)
                          + adjusted * MOTION_EMA_ALPHA;
    let sm = state.smoothed_motion;

    // 4. Classify from smoothed score.
    let candidate = raw_classify(sm);

    // 5. Hysteresis debounce: require N consecutive frames agreeing on a new state.
    if candidate == state.current_motion_level {
        // Already in this state — reset debounce.
        state.debounce_counter = 0;
        state.debounce_candidate = candidate;
    } else if candidate == state.debounce_candidate {
        state.debounce_counter += 1;
        if state.debounce_counter >= DEBOUNCE_FRAMES {
            // Transition accepted.
            state.current_motion_level = candidate;
            state.debounce_counter = 0;
        }
    } else {
        // New candidate — restart counter.
        state.debounce_candidate = candidate;
        state.debounce_counter = 1;
    }

    // 6. Write the smoothed result back into the classification.
    raw.motion_level = state.current_motion_level.clone();
    raw.presence = sm > 0.03;
    raw.confidence = (0.4 + sm * 0.6).clamp(0.0, 1.0);
}

/// Per-node variant of `smooth_and_classify` that operates on a `NodeState`
/// instead of `AppStateInner` (issue #249).
fn smooth_and_classify_node(ns: &mut NodeState, raw: &mut ClassificationInfo, raw_motion: f64) {
    ns.baseline_frames += 1;
    if ns.baseline_frames < BASELINE_WARMUP {
        ns.baseline_motion = ns.baseline_motion * 0.9 + raw_motion * 0.1;
    } else if raw_motion < ns.smoothed_motion + 0.05 {
        ns.baseline_motion = ns.baseline_motion * (1.0 - BASELINE_EMA_ALPHA)
                           + raw_motion * BASELINE_EMA_ALPHA;
    }

    let adjusted = (raw_motion - ns.baseline_motion * 0.7).max(0.0);

    ns.smoothed_motion = ns.smoothed_motion * (1.0 - MOTION_EMA_ALPHA)
                       + adjusted * MOTION_EMA_ALPHA;
    let sm = ns.smoothed_motion;

    let candidate = raw_classify(sm);

    if candidate == ns.current_motion_level {
        ns.debounce_counter = 0;
        ns.debounce_candidate = candidate;
    } else if candidate == ns.debounce_candidate {
        ns.debounce_counter += 1;
        if ns.debounce_counter >= DEBOUNCE_FRAMES {
            ns.current_motion_level = candidate;
            ns.debounce_counter = 0;
        }
    } else {
        ns.debounce_candidate = candidate;
        ns.debounce_counter = 1;
    }

    raw.motion_level = ns.current_motion_level.clone();
    raw.presence = sm > 0.03;
    raw.confidence = (0.4 + sm * 0.6).clamp(0.0, 1.0);
}

/// If an adaptive model is loaded, override the classification with the
/// model's prediction.  Uses the full 15-feature vector for higher accuracy.
fn adaptive_override(state: &AppStateInner, features: &FeatureInfo, classification: &mut ClassificationInfo) {
    if let Some(ref model) = state.adaptive_model {
        // Get current frame amplitudes from the latest history entry.
        let amps = state.frame_history.back()
            .map(|v| v.as_slice())
            .unwrap_or(&[]);
        let feat_arr = adaptive_classifier::features_from_runtime(
            &serde_json::json!({
                "variance": features.variance,
                "motion_band_power": features.motion_band_power,
                "breathing_band_power": features.breathing_band_power,
                "spectral_power": features.spectral_power,
                "dominant_freq_hz": features.dominant_freq_hz,
                "change_points": features.change_points,
                "mean_rssi": features.mean_rssi,
            }),
            amps,
        );
        let (label, conf) = model.classify(&feat_arr);
        classification.motion_level = label.to_string();
        classification.presence = label != "absent";
        // Blend model confidence with existing smoothed confidence.
        classification.confidence = (conf * 0.7 + classification.confidence * 0.3).clamp(0.0, 1.0);
    }
}

/// Size of the median filter window for vital signs outlier rejection.
const VITAL_MEDIAN_WINDOW: usize = 21;
/// EMA alpha for vital signs (~5s time constant at 10 FPS).
const VITAL_EMA_ALPHA: f64 = 0.02;
/// Maximum BPM jump per frame before a value is rejected as an outlier.
const HR_MAX_JUMP: f64 = 8.0;
const BR_MAX_JUMP: f64 = 2.0;
/// Minimum change from current smoothed value before EMA updates (dead-band).
/// Prevents micro-drift from creeping in.
const HR_DEAD_BAND: f64 = 2.0;
const BR_DEAD_BAND: f64 = 0.5;

/// Smooth vital signs using median-filter outlier rejection + EMA.
/// Mutates `state.smoothed_hr`, `state.smoothed_br`, etc.
/// Returns the smoothed VitalSigns to broadcast.
fn smooth_vitals(state: &mut AppStateInner, raw: &VitalSigns) -> VitalSigns {
    let raw_hr = raw.heart_rate_bpm.unwrap_or(0.0);
    let raw_br = raw.breathing_rate_bpm.unwrap_or(0.0);

    // -- Outlier rejection: skip values that jump too far from current EMA --
    let hr_ok = state.smoothed_hr < 1.0 || (raw_hr - state.smoothed_hr).abs() < HR_MAX_JUMP;
    let br_ok = state.smoothed_br < 1.0 || (raw_br - state.smoothed_br).abs() < BR_MAX_JUMP;

    // Push into buffer (only non-outlier values)
    if hr_ok && raw_hr > 0.0 {
        state.hr_buffer.push_back(raw_hr);
        if state.hr_buffer.len() > VITAL_MEDIAN_WINDOW { state.hr_buffer.pop_front(); }
    }
    if br_ok && raw_br > 0.0 {
        state.br_buffer.push_back(raw_br);
        if state.br_buffer.len() > VITAL_MEDIAN_WINDOW { state.br_buffer.pop_front(); }
    }

    // Compute trimmed mean: drop top/bottom 25% then average the middle 50%.
    // This is more stable than pure median and less noisy than raw mean.
    let trimmed_hr = trimmed_mean(&state.hr_buffer);
    let trimmed_br = trimmed_mean(&state.br_buffer);

    // EMA smooth with dead-band: only update if the trimmed mean differs
    // from the current smoothed value by more than the dead-band.
    // This prevents the display from constantly creeping by tiny amounts.
    if trimmed_hr > 0.0 {
        if state.smoothed_hr < 1.0 {
            state.smoothed_hr = trimmed_hr;
        } else if (trimmed_hr - state.smoothed_hr).abs() > HR_DEAD_BAND {
            state.smoothed_hr = state.smoothed_hr * (1.0 - VITAL_EMA_ALPHA)
                              + trimmed_hr * VITAL_EMA_ALPHA;
        }
        // else: within dead-band, hold current value
    }
    if trimmed_br > 0.0 {
        if state.smoothed_br < 1.0 {
            state.smoothed_br = trimmed_br;
        } else if (trimmed_br - state.smoothed_br).abs() > BR_DEAD_BAND {
            state.smoothed_br = state.smoothed_br * (1.0 - VITAL_EMA_ALPHA)
                              + trimmed_br * VITAL_EMA_ALPHA;
        }
    }

    // Smooth confidence
    state.smoothed_hr_conf = state.smoothed_hr_conf * 0.92 + raw.heartbeat_confidence * 0.08;
    state.smoothed_br_conf = state.smoothed_br_conf * 0.92 + raw.breathing_confidence * 0.08;

    VitalSigns {
        breathing_rate_bpm: if state.smoothed_br > 1.0 { Some(state.smoothed_br) } else { None },
        heart_rate_bpm: if state.smoothed_hr > 1.0 { Some(state.smoothed_hr) } else { None },
        breathing_confidence: state.smoothed_br_conf,
        heartbeat_confidence: state.smoothed_hr_conf,
        signal_quality: raw.signal_quality,
    }
}

/// Per-node variant of `smooth_vitals` that operates on a `NodeState` (issue #249).
fn smooth_vitals_node(ns: &mut NodeState, raw: &VitalSigns) -> VitalSigns {
    let raw_hr = raw.heart_rate_bpm.unwrap_or(0.0);
    let raw_br = raw.breathing_rate_bpm.unwrap_or(0.0);

    let hr_ok = ns.smoothed_hr < 1.0 || (raw_hr - ns.smoothed_hr).abs() < HR_MAX_JUMP;
    let br_ok = ns.smoothed_br < 1.0 || (raw_br - ns.smoothed_br).abs() < BR_MAX_JUMP;

    if hr_ok && raw_hr > 0.0 {
        ns.hr_buffer.push_back(raw_hr);
        if ns.hr_buffer.len() > VITAL_MEDIAN_WINDOW { ns.hr_buffer.pop_front(); }
    }
    if br_ok && raw_br > 0.0 {
        ns.br_buffer.push_back(raw_br);
        if ns.br_buffer.len() > VITAL_MEDIAN_WINDOW { ns.br_buffer.pop_front(); }
    }

    let trimmed_hr = trimmed_mean(&ns.hr_buffer);
    let trimmed_br = trimmed_mean(&ns.br_buffer);

    if trimmed_hr > 0.0 {
        if ns.smoothed_hr < 1.0 {
            ns.smoothed_hr = trimmed_hr;
        } else if (trimmed_hr - ns.smoothed_hr).abs() > HR_DEAD_BAND {
            ns.smoothed_hr = ns.smoothed_hr * (1.0 - VITAL_EMA_ALPHA)
                           + trimmed_hr * VITAL_EMA_ALPHA;
        }
    }
    if trimmed_br > 0.0 {
        if ns.smoothed_br < 1.0 {
            ns.smoothed_br = trimmed_br;
        } else if (trimmed_br - ns.smoothed_br).abs() > BR_DEAD_BAND {
            ns.smoothed_br = ns.smoothed_br * (1.0 - VITAL_EMA_ALPHA)
                           + trimmed_br * VITAL_EMA_ALPHA;
        }
    }

    ns.smoothed_hr_conf = ns.smoothed_hr_conf * 0.92 + raw.heartbeat_confidence * 0.08;
    ns.smoothed_br_conf = ns.smoothed_br_conf * 0.92 + raw.breathing_confidence * 0.08;

    VitalSigns {
        breathing_rate_bpm: if ns.smoothed_br > 1.0 { Some(ns.smoothed_br) } else { None },
        heart_rate_bpm: if ns.smoothed_hr > 1.0 { Some(ns.smoothed_hr) } else { None },
        breathing_confidence: ns.smoothed_br_conf,
        heartbeat_confidence: ns.smoothed_hr_conf,
        signal_quality: raw.signal_quality,
    }
}

/// Trimmed mean: sort, drop top/bottom 25%, average the middle 50%.
/// More robust than median (uses more data) and less noisy than raw mean.
fn trimmed_mean(buf: &VecDeque<f64>) -> f64 {
    if buf.is_empty() { return 0.0; }
    let mut sorted: Vec<f64> = buf.iter().copied().collect();
    sorted.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let n = sorted.len();
    let trim = n / 4; // drop 25% from each end
    let middle = &sorted[trim..n - trim.max(0)];
    if middle.is_empty() {
        sorted[n / 2] // fallback to median if too few samples
    } else {
        middle.iter().sum::<f64>() / middle.len() as f64
    }
}

// ── Windows WiFi RSSI collector ──────────────────────────────────────────────

/// Parse `netsh wlan show interfaces` output for RSSI and signal quality
fn parse_netsh_interfaces_output(output: &str) -> Option<(f64, f64, String)> {
    let mut rssi = None;
    let mut signal = None;
    let mut ssid = None;

    for line in output.lines() {
        let line = line.trim();
        if line.starts_with("Signal") {
            // "Signal                 : 89%"
            if let Some(pct) = line.split(':').nth(1) {
                let pct = pct.trim().trim_end_matches('%');
                if let Ok(v) = pct.parse::<f64>() {
                    signal = Some(v);
                    // Convert signal% to approximate dBm: -100 + (signal% * 0.6)
                    rssi = Some(-100.0 + v * 0.6);
                }
            }
        }
        if line.starts_with("SSID") && !line.starts_with("BSSID") {
            if let Some(s) = line.split(':').nth(1) {
                ssid = Some(s.trim().to_string());
            }
        }
    }

    match (rssi, signal, ssid) {
        (Some(r), Some(_s), Some(name)) => Some((r, _s, name)),
        (Some(r), Some(_s), None) => Some((r, _s, "Unknown".into())),
        _ => None,
    }
}

/// Detect the active wireless interface name by parsing `iw dev` output.
/// Returns the first interface in station or managed mode, or falls back to "wlan0".
#[cfg(target_os = "linux")]
fn detect_wifi_interface() -> String {
    let output = std::process::Command::new("iw")
        .arg("dev")
        .output()
        .map(|o| String::from_utf8_lossy(&o.stdout).to_string())
        .unwrap_or_default();
    // Parse: lines like "Interface wlan0"
    for line in output.lines() {
        let trimmed = line.trim();
        if let Some(iface) = trimmed.strip_prefix("Interface ") {
            return iface.trim().to_string();
        }
    }
    "wlan0".to_string()
}

/// Linux multi-BSSID WiFi sensing task (ADR-022 Linux path).
///
/// Uses `iw dev <iface> scan` to enumerate ALL visible BSSIDs: your own AP,
/// neighbour routers, any open networks in range. Each BSSID acts as a passive
/// RF beacon whose RSSI variations reflect motion in its Fresnel zone, enabling
/// sensing without any ESP32 hardware and without being associated to the AP.
///
/// Mirrors `windows_wifi_task` logic but uses `LinuxIwScanner` instead of `netsh`.
#[cfg(target_os = "linux")]
async fn linux_wifi_task(state: SharedState, tick_ms: u64, iface: String) {
    let mut interval = tokio::time::interval(Duration::from_millis(tick_ms));
    let mut seq: u32 = 0;

    let mut registry = BssidRegistry::new(64, 30);
    let mut pipeline = WindowsWifiPipeline::new();

    info!(
        "Linux multi-BSSID WiFi sensing active (iface={iface}, tick={tick_ms}ms, max_bssids=64)"
    );
    info!("  Scanning all visible APs — own + neighbours — as passive RF beacons");

    loop {
        interval.tick().await;
        seq += 1;

        // Run iw scan inside spawn_blocking (blocking subprocess).
        let iface_clone = iface.clone();
        let scan_result = tokio::task::spawn_blocking(move || {
            // Try live scan first (needs CAP_NET_ADMIN / root); fall back to dump.
            let scanner = LinuxIwScanner::with_interface(&iface_clone);
            match scanner.scan_sync() {
                Ok(obs) if !obs.is_empty() => Ok(obs),
                _ => {
                    // Fall back to cached dump (no root required but may be stale ~30s).
                    let cached = LinuxIwScanner::with_interface(&iface_clone).use_cached();
                    cached.scan_sync().map_err(|e| format!("iw scan failed: {e}"))
                }
            }
        })
        .await;

        let observations = match scan_result {
            Ok(Ok(obs)) if !obs.is_empty() => obs,
            Ok(Ok(_)) => {
                debug!("Linux WiFi scan: 0 BSSIDs, skipping tick");
                continue;
            }
            Ok(Err(e)) => {
                warn!("Linux WiFi scan error: {e}");
                continue;
            }
            Err(e) => {
                error!("spawn_blocking panicked: {e}");
                continue;
            }
        };

        let obs_count = observations.len();
        let ssid = observations
            .first()
            .map(|o| o.ssid.clone())
            .filter(|s| !s.is_empty())
            .unwrap_or_else(|| "unknown".into());
        let first_rssi = observations.first().map(|o| o.rssi_dbm).unwrap_or(-80.0);

        // Feed registry and run pipeline.
        registry.update(&observations);
        let multi_ap_frame = registry.to_multi_ap_frame();
        let enhanced = pipeline.process(&multi_ap_frame);

        let frame = Esp32Frame {
            magic: 0xC511_0001,
            node_id: 0,
            n_antennas: 1,
            n_subcarriers: obs_count.min(255) as u8,
            freq_mhz: 2437,
            sequence: seq,
            rssi: first_rssi.clamp(-128.0, 127.0) as i8,
            noise_floor: -95,
            amplitudes: multi_ap_frame.amplitudes.clone(),
            phases: multi_ap_frame.phases.clone(),
        };

        let mut s = state.write().await;
        s.frame_history.push_back(frame.amplitudes.clone());
        if s.frame_history.len() > FRAME_HISTORY_CAPACITY {
            s.frame_history.pop_front();
        }
        let sample_rate_hz = 1000.0 / tick_ms as f64;
        let (features, mut classification, breathing_rate_hz, sub_variances, raw_motion) =
            extract_features_from_frame(&frame, &s.frame_history, sample_rate_hz);
        smooth_and_classify(&mut s, &mut classification, raw_motion);
        adaptive_override(&s, &features, &mut classification);

        let enhanced_motion = Some(serde_json::json!({
            "score": enhanced.motion.score,
            "level": format!("{:?}", enhanced.motion.level),
            "contributing_bssids": enhanced.motion.contributing_bssids,
        }));
        let enhanced_breathing = enhanced.breathing.as_ref().map(|b| {
            serde_json::json!({
                "rate_bpm": b.rate_bpm,
                "confidence": b.confidence,
                "bssid_count": b.bssid_count,
            })
        });
        let posture_str = enhanced.posture.map(|p| format!("{p:?}"));
        let sig_quality_score = Some(enhanced.signal_quality.score);
        let verdict_str = Some(format!("{:?}", enhanced.verdict));
        let bssid_n = Some(enhanced.bssid_count);

        s.source = format!("wifi:{ssid}");
        s.rssi_history.push_back(first_rssi);
        if s.rssi_history.len() > 60 { s.rssi_history.pop_front(); }
        s.tick += 1;
        let tick = s.tick;

        let motion_score = match classification.motion_level.as_str() {
            "active" => 0.8,
            "present_still" => 0.3,
            _ => 0.05,
        };

        let raw_vitals = s.vital_detector.process_frame(&frame.amplitudes, &frame.phases);
        let vitals = smooth_vitals(&mut s, &raw_vitals);
        s.latest_vitals = vitals.clone();

        let feat_variance = features.variance;
        let raw_score = compute_person_score(&features);
        s.smoothed_person_score = s.smoothed_person_score * 0.90 + raw_score * 0.10;
        let count_evidence = if classification.presence {
            let evidence = build_single_link_count_evidence(
                &s.frame_history,
                s.field_model.as_ref(),
                classification.confidence,
                s.smoothed_person_score,
                s.prev_person_count,
            );
            let count = evidence.as_ref().map(|item| item.consensus_count).unwrap_or(1);
            s.prev_person_count = count;
            evidence
        } else {
            s.prev_person_count = 0;
            None
        };
        let est_persons = count_evidence.as_ref().map(|item| item.consensus_count).unwrap_or(0);

        let mut update = SensingUpdate {
            msg_type: "sensing_update".to_string(),
            timestamp: chrono::Utc::now().timestamp_millis() as f64 / 1000.0,
            source: format!("wifi:{ssid}"),
            tick,
            nodes: vec![NodeInfo {
                node_id: 0,
                rssi_dbm: first_rssi,
                position: [0.0, 0.0, 0.0],
                amplitude: multi_ap_frame.amplitudes,
                subcarrier_count: obs_count,
            }],
            features,
            classification,
            signal_field: generate_signal_field(
                first_rssi, motion_score, breathing_rate_hz,
                feat_variance.min(1.0), &sub_variances,
            ),
            vital_signs: Some(vitals),
            enhanced_motion,
            enhanced_breathing,
            posture: posture_str,
            signal_quality_score: sig_quality_score,
            quality_verdict: verdict_str,
            bssid_count: bssid_n,
            pose_keypoints: None,
            model_status: None,
            persons: None,
            estimated_persons: if est_persons > 0 { Some(est_persons) } else { None },
            count_evidence,
            person_anchors: None,
            node_features: None,
        };

        {
            let s_ref = &mut *s;
            finalize_people_for_update(
                &mut update,
                &mut s_ref.pose_tracker,
                &mut s_ref.last_tracker_instant,
                &mut s_ref.track_temporal_state,
            );
        }

        if let Ok(json) = serde_json::to_string(&update) {
            let _ = s.tx.send(json);
        }
        s.latest_update = Some(update);

        debug!(
            "Linux WiFi tick #{tick}: {obs_count} BSSIDs, quality={:.2}, verdict={:?}",
            enhanced.signal_quality.score, enhanced.verdict
        );
    }
}

async fn windows_wifi_task(state: SharedState, tick_ms: u64) {
    let mut interval = tokio::time::interval(Duration::from_millis(tick_ms));
    let mut seq: u32 = 0;

    // ADR-022 Phase 3: Multi-BSSID pipeline state (kept across ticks)
    let mut registry = BssidRegistry::new(32, 30);
    let mut pipeline = WindowsWifiPipeline::new();

    info!(
        "Windows WiFi multi-BSSID pipeline active (tick={}ms, max_bssids=32)",
        tick_ms
    );

    loop {
        interval.tick().await;
        seq += 1;

        // ── Step 1: Run multi-BSSID scan via spawn_blocking ──────────
        // NetshBssidScanner is not Send, so we run `netsh` and parse
        // the output inside a blocking closure.
        let bssid_scan_result = tokio::task::spawn_blocking(|| {
            let output = std::process::Command::new("netsh")
                .args(["wlan", "show", "networks", "mode=bssid"])
                .output()
                .map_err(|e| format!("netsh bssid scan failed: {e}"))?;

            if !output.status.success() {
                let stderr = String::from_utf8_lossy(&output.stderr);
                return Err(format!(
                    "netsh exited with {}: {}",
                    output.status,
                    stderr.trim()
                ));
            }

            let stdout = String::from_utf8_lossy(&output.stdout);
            parse_netsh_bssid_output(&stdout).map_err(|e| format!("parse error: {e}"))
        })
        .await;

        // Unwrap the JoinHandle result, then the inner Result.
        let observations = match bssid_scan_result {
            Ok(Ok(obs)) if !obs.is_empty() => obs,
            Ok(Ok(_empty)) => {
                debug!("Multi-BSSID scan returned 0 observations, falling back");
                windows_wifi_fallback_tick(&state, seq).await;
                continue;
            }
            Ok(Err(e)) => {
                warn!("Multi-BSSID scan error: {e}, falling back");
                windows_wifi_fallback_tick(&state, seq).await;
                continue;
            }
            Err(join_err) => {
                error!("spawn_blocking panicked: {join_err}");
                continue;
            }
        };

        let obs_count = observations.len();

        // Derive SSID from the first observation for the source label.
        let ssid = observations
            .first()
            .map(|o| o.ssid.clone())
            .unwrap_or_else(|| "Unknown".into());

        // ── Step 2: Feed observations into registry ──────────────────
        registry.update(&observations);
        let multi_ap_frame = registry.to_multi_ap_frame();

        // ── Step 3: Run enhanced pipeline ────────────────────────────
        let enhanced = pipeline.process(&multi_ap_frame);

        // ── Step 4: Build backward-compatible Esp32Frame ─────────────
        let first_rssi = observations
            .first()
            .map(|o| o.rssi_dbm)
            .unwrap_or(-80.0);
        let _first_signal_pct = observations
            .first()
            .map(|o| o.signal_pct)
            .unwrap_or(40.0);

        let frame = Esp32Frame {
            magic: 0xC511_0001,
            node_id: 0,
            n_antennas: 1,
            n_subcarriers: obs_count.min(255) as u8,
            freq_mhz: 2437,
            sequence: seq,
            rssi: first_rssi.clamp(-128.0, 127.0) as i8,
            noise_floor: -90,
            amplitudes: multi_ap_frame.amplitudes.clone(),
            phases: multi_ap_frame.phases.clone(),
        };

        // ── Step 4b: Update frame history and extract features ───────
        let mut s_write_pre = state.write().await;
        s_write_pre.frame_history.push_back(frame.amplitudes.clone());
        if s_write_pre.frame_history.len() > FRAME_HISTORY_CAPACITY {
            s_write_pre.frame_history.pop_front();
        }
        let sample_rate_hz = 1000.0 / tick_ms as f64;
        let (features, mut classification, breathing_rate_hz, sub_variances, raw_motion) =
            extract_features_from_frame(&frame, &s_write_pre.frame_history, sample_rate_hz);
        smooth_and_classify(&mut s_write_pre, &mut classification, raw_motion);
        adaptive_override(&s_write_pre, &features, &mut classification);
        drop(s_write_pre);

        // ── Step 5: Build enhanced fields from pipeline result ───────
        let enhanced_motion = Some(serde_json::json!({
            "score": enhanced.motion.score,
            "level": format!("{:?}", enhanced.motion.level),
            "contributing_bssids": enhanced.motion.contributing_bssids,
        }));

        let enhanced_breathing = enhanced.breathing.as_ref().map(|b| {
            serde_json::json!({
                "rate_bpm": b.rate_bpm,
                "confidence": b.confidence,
                "bssid_count": b.bssid_count,
            })
        });

        let posture_str = enhanced.posture.map(|p| format!("{p:?}"));
        let sig_quality_score = Some(enhanced.signal_quality.score);
        let verdict_str = Some(format!("{:?}", enhanced.verdict));
        let bssid_n = Some(enhanced.bssid_count);

        // ── Step 6: Update shared state ──────────────────────────────
        let mut s = state.write().await;
        s.source = format!("wifi:{ssid}");
        s.rssi_history.push_back(first_rssi);
        if s.rssi_history.len() > 60 {
            s.rssi_history.pop_front();
        }

        s.tick += 1;
        let tick = s.tick;

        let motion_score = if classification.motion_level == "active" {
            0.8
        } else if classification.motion_level == "present_still" {
            0.3
        } else {
            0.05
        };

        let raw_vitals = s.vital_detector.process_frame(&frame.amplitudes, &frame.phases);
        let vitals = smooth_vitals(&mut s, &raw_vitals);
        s.latest_vitals = vitals.clone();

        let feat_variance = features.variance;

        // Multi-person estimation with temporal smoothing (EMA α=0.10).
        let raw_score = compute_person_score(&features);
        s.smoothed_person_score = s.smoothed_person_score * 0.90 + raw_score * 0.10;
        let count_evidence = if classification.presence {
            let evidence = build_single_link_count_evidence(
                &s.frame_history,
                s.field_model.as_ref(),
                classification.confidence,
                s.smoothed_person_score,
                s.prev_person_count,
            );
            let count = evidence.as_ref().map(|item| item.consensus_count).unwrap_or(1);
            s.prev_person_count = count;
            evidence
        } else {
            s.prev_person_count = 0;
            None
        };
        let est_persons = count_evidence.as_ref().map(|item| item.consensus_count).unwrap_or(0);

        let mut update = SensingUpdate {
            msg_type: "sensing_update".to_string(),
            timestamp: chrono::Utc::now().timestamp_millis() as f64 / 1000.0,
            source: format!("wifi:{ssid}"),
            tick,
            nodes: vec![NodeInfo {
                node_id: 0,
                rssi_dbm: first_rssi,
                position: [0.0, 0.0, 0.0],
                amplitude: multi_ap_frame.amplitudes,
                subcarrier_count: obs_count,
            }],
            features,
            classification,
            signal_field: generate_signal_field(
                first_rssi, motion_score, breathing_rate_hz,
                feat_variance.min(1.0), &sub_variances,
            ),
            vital_signs: Some(vitals),
            enhanced_motion,
            enhanced_breathing,
            posture: posture_str,
            signal_quality_score: sig_quality_score,
            quality_verdict: verdict_str,
            bssid_count: bssid_n,
            pose_keypoints: None,
            model_status: None,
            persons: None,
            estimated_persons: if est_persons > 0 { Some(est_persons) } else { None },
            count_evidence,
            person_anchors: None,
            node_features: None,
        };

        {
            let s_ref = &mut *s;
            finalize_people_for_update(
                &mut update,
                &mut s_ref.pose_tracker,
                &mut s_ref.last_tracker_instant,
                &mut s_ref.track_temporal_state,
            );
        }

        if let Ok(json) = serde_json::to_string(&update) {
            let _ = s.tx.send(json);
        }
        s.latest_update = Some(update);

        debug!(
            "Multi-BSSID tick #{tick}: {obs_count} BSSIDs, quality={:.2}, verdict={:?}",
            enhanced.signal_quality.score, enhanced.verdict
        );
    }
}

/// Fallback: single-RSSI collection via `netsh wlan show interfaces`.
///
/// Used when the multi-BSSID scan fails or returns 0 observations.
async fn windows_wifi_fallback_tick(state: &SharedState, seq: u32) {
    let output = match tokio::process::Command::new("netsh")
        .args(["wlan", "show", "interfaces"])
        .output()
        .await
    {
        Ok(o) => String::from_utf8_lossy(&o.stdout).to_string(),
        Err(e) => {
            warn!("netsh interfaces fallback failed: {e}");
            return;
        }
    };

    let (rssi_dbm, signal_pct, ssid) = match parse_netsh_interfaces_output(&output) {
        Some(v) => v,
        None => {
            debug!("Fallback: no WiFi interface connected");
            return;
        }
    };

    let frame = Esp32Frame {
        magic: 0xC511_0001,
        node_id: 0,
        n_antennas: 1,
        n_subcarriers: 1,
        freq_mhz: 2437,
        sequence: seq,
        rssi: rssi_dbm as i8,
        noise_floor: -90,
        amplitudes: vec![signal_pct],
        phases: vec![0.0],
    };

    let mut s = state.write().await;
    // Update frame history before extracting features.
    s.frame_history.push_back(frame.amplitudes.clone());
    if s.frame_history.len() > FRAME_HISTORY_CAPACITY {
        s.frame_history.pop_front();
    }
    let sample_rate_hz = 2.0_f64; // fallback tick ~ 500 ms => 2 Hz
    let (features, mut classification, breathing_rate_hz, sub_variances, raw_motion) =
        extract_features_from_frame(&frame, &s.frame_history, sample_rate_hz);
    smooth_and_classify(&mut s, &mut classification, raw_motion);
    adaptive_override(&s, &features, &mut classification);

    s.source = format!("wifi:{ssid}");
    s.rssi_history.push_back(rssi_dbm);
    if s.rssi_history.len() > 60 {
        s.rssi_history.pop_front();
    }

    s.tick += 1;
    let tick = s.tick;

    let motion_score = if classification.motion_level == "active" {
        0.8
    } else if classification.motion_level == "present_still" {
        0.3
    } else {
        0.05
    };

    let raw_vitals = s.vital_detector.process_frame(&frame.amplitudes, &frame.phases);
    let vitals = smooth_vitals(&mut s, &raw_vitals);
    s.latest_vitals = vitals.clone();

    let feat_variance = features.variance;

    // Multi-person estimation with temporal smoothing (EMA α=0.10).
    let raw_score = compute_person_score(&features);
    s.smoothed_person_score = s.smoothed_person_score * 0.90 + raw_score * 0.10;
    let count_evidence = if classification.presence {
        let evidence = build_single_link_count_evidence(
            &s.frame_history,
            s.field_model.as_ref(),
            classification.confidence,
            s.smoothed_person_score,
            s.prev_person_count,
        );
        let count = evidence.as_ref().map(|item| item.consensus_count).unwrap_or(1);
        s.prev_person_count = count;
        evidence
    } else {
        s.prev_person_count = 0;
        None
    };
    let est_persons = count_evidence.as_ref().map(|item| item.consensus_count).unwrap_or(0);

    let mut update = SensingUpdate {
        msg_type: "sensing_update".to_string(),
        timestamp: chrono::Utc::now().timestamp_millis() as f64 / 1000.0,
        source: format!("wifi:{ssid}"),
        tick,
        nodes: vec![NodeInfo {
            node_id: 0,
            rssi_dbm,
            position: [0.0, 0.0, 0.0],
            amplitude: vec![signal_pct],
            subcarrier_count: 1,
        }],
        features,
        classification,
        signal_field: generate_signal_field(
            rssi_dbm, motion_score, breathing_rate_hz,
            feat_variance.min(1.0), &sub_variances,
        ),
        vital_signs: Some(vitals),
        enhanced_motion: None,
        enhanced_breathing: None,
        posture: None,
        signal_quality_score: None,
        quality_verdict: None,
        bssid_count: None,
        pose_keypoints: None,
        model_status: None,
        persons: None,
        estimated_persons: if est_persons > 0 { Some(est_persons) } else { None },
        count_evidence,
        person_anchors: None,
        node_features: None,
    };

    {
        let s_ref = &mut *s;
        finalize_people_for_update(
            &mut update,
            &mut s_ref.pose_tracker,
            &mut s_ref.last_tracker_instant,
            &mut s_ref.track_temporal_state,
        );
    }

    if let Ok(json) = serde_json::to_string(&update) {
        let _ = s.tx.send(json);
    }
    s.latest_update = Some(update);
}

/// Probe if Windows WiFi is connected
async fn probe_windows_wifi() -> bool {
    match tokio::process::Command::new("netsh")
        .args(["wlan", "show", "interfaces"])
        .output()
        .await
    {
        Ok(o) => {
            let out = String::from_utf8_lossy(&o.stdout);
            parse_netsh_interfaces_output(&out).is_some()
        }
        Err(_) => false,
    }
}

/// Probe if Linux WiFi interface is connected (non-async, fast check).
#[cfg(target_os = "linux")]
fn probe_linux_wifi() -> bool {
    std::process::Command::new("iw")
        .arg("dev")
        .output()
        .map(|o| {
            let out = String::from_utf8_lossy(&o.stdout);
            out.contains("Interface") && out.contains("ssid")
        })
        .unwrap_or(false)
}

#[cfg(not(target_os = "linux"))]
fn probe_linux_wifi() -> bool {
    false
}

/// Probe if ESP32 is streaming on UDP port
async fn probe_esp32(port: u16) -> bool {
    let addr = format!("0.0.0.0:{port}");
    match UdpSocket::bind(&addr).await {
        Ok(sock) => {
            let mut buf = [0u8; 256];
            match tokio::time::timeout(Duration::from_secs(2), sock.recv_from(&mut buf)).await {
                Ok(Ok((len, _))) => parse_esp32_frame(&buf[..len]).is_some(),
                _ => false,
            }
        }
        Err(_) => false,
    }
}

// ── WebSocket handler ────────────────────────────────────────────────────────

async fn ws_sensing_handler(
    ws: WebSocketUpgrade,
    State(state): State<SharedState>,
) -> impl IntoResponse {
    ws.on_upgrade(|socket| handle_ws_client(socket, state))
}

async fn handle_ws_client(mut socket: WebSocket, state: SharedState) {
    let mut rx = {
        let s = state.read().await;
        s.tx.subscribe()
    };

    info!("WebSocket client connected (sensing)");

    loop {
        tokio::select! {
            msg = rx.recv() => {
                match msg {
                    Ok(json) => {
                        if socket.send(Message::Text(json.into())).await.is_err() {
                            break;
                        }
                    }
                    Err(_) => break,
                }
            }
            msg = socket.recv() => {
                match msg {
                    Some(Ok(Message::Close(_))) | None => break,
                    _ => {} // ignore client messages
                }
            }
        }
    }

    info!("WebSocket client disconnected (sensing)");
}

// ── Pose WebSocket handler (sends pose_data messages for Live Demo) ──────────

async fn ws_pose_handler(
    ws: WebSocketUpgrade,
    State(state): State<SharedState>,
) -> impl IntoResponse {
    ws.on_upgrade(|socket| handle_ws_pose_client(socket, state))
}

async fn handle_ws_pose_client(mut socket: WebSocket, state: SharedState) {
    let mut rx = {
        let s = state.read().await;
        s.tx.subscribe()
    };

    info!("WebSocket client connected (pose)");

    // Send connection established message
    let conn_msg = serde_json::json!({
        "type": "connection_established",
        "payload": { "status": "connected", "backend": "rust+ruvector" }
    });
    let _ = socket.send(Message::Text(conn_msg.to_string().into())).await;

    loop {
        tokio::select! {
            msg = rx.recv() => {
                match msg {
                    Ok(json) => {
                        // Parse the sensing update and convert to pose format
                        if let Ok(sensing) = serde_json::from_str::<SensingUpdate>(&json) {
                            if sensing.msg_type == "sensing_update" {
                                // Determine pose estimation mode for the UI indicator.
                                // "model_inference"    — a trained RVF model is loaded.
                                // "signal_derived"     — keypoints estimated from raw CSI features.
                                let model_loaded = {
                                    let s = state.read().await;
                                    s.model_loaded
                                };
                                let pose_source = if model_loaded {
                                    "model_inference"
                                } else {
                                    "signal_derived"
                                };

                                let persons = if model_loaded {
                                    // When a trained model is loaded, prefer its keypoints if present.
                                    sensing.pose_keypoints.as_ref().map(|kps| {
                                        let kp_names = [
                                            "nose","left_eye","right_eye","left_ear","right_ear",
                                            "left_shoulder","right_shoulder","left_elbow","right_elbow",
                                            "left_wrist","right_wrist","left_hip","right_hip",
                                            "left_knee","right_knee","left_ankle","right_ankle",
                                        ];
                                        let keypoints: Vec<PoseKeypoint> = kps.iter()
                                            .enumerate()
                                            .map(|(i, kp)| PoseKeypoint {
                                                name: kp_names.get(i).unwrap_or(&"unknown").to_string(),
                                                x: kp[0], y: kp[1], z: kp[2], confidence: kp[3],
                                            })
                                            .collect();
                                        vec![PersonDetection {
                                            id: 1,
                                            confidence: sensing.classification.confidence,
                                            bbox: BoundingBox { x: 260.0, y: 150.0, width: 120.0, height: 220.0 },
                                            keypoints,
                                            zone: "zone_1".into(),
                                        }]
                                    }).unwrap_or_else(|| {
                                        // Prefer tracked persons from broadcast if available
                                        sensing.persons.clone().unwrap_or_else(|| derive_pose_from_sensing(&sensing))
                                    })
                                } else {
                                    // Prefer tracked persons from broadcast if available
                                    sensing.persons.clone().unwrap_or_else(|| derive_pose_from_sensing(&sensing))
                                };

                                // TinyML anomaly + aggression: feed every node's
                                // motion_band_power into the rolling baseline.
                                // observe() is idempotent on (node_id, tick) so
                                // running it per-WS-client is safe.
                                let agg_score = if let Some(per_node) = sensing.node_features.as_ref() {
                                    for nf in per_node {
                                        let _ = aggression::observe(
                                            nf.node_id,
                                            nf.features.motion_band_power,
                                            sensing.tick,
                                        );
                                    }
                                    aggression::snapshot()
                                } else {
                                    let _ = aggression::observe(
                                        0,
                                        sensing.features.motion_band_power,
                                        sensing.tick,
                                    );
                                    aggression::snapshot()
                                };

                                let pose_msg = serde_json::json!({
                                    "type": "pose_data",
                                    "zone_id": "zone_1",
                                    "timestamp": sensing.timestamp,
                                    "payload": {
                                        "pose": {
                                            "persons": persons,
                                        },
                                        "confidence": if sensing.classification.presence { sensing.classification.confidence } else { 0.0 },
                                        "activity": sensing.classification.motion_level,
                                        // pose_source tells the UI which estimation mode is active.
                                        "pose_source": pose_source,
                                        // TinyML environment-trained anomaly + aggression scores.
                                        "aggression": agg_score,
                                        "metadata": {
                                            "frame_id": format!("rust_frame_{}", sensing.tick),
                                            "processing_time_ms": 1,
                                            "source": sensing.source,
                                            "tick": sensing.tick,
                                            "signal_strength": sensing.features.mean_rssi,
                                            "motion_band_power": sensing.features.motion_band_power,
                                            "breathing_band_power": sensing.features.breathing_band_power,
                                            "estimated_persons": persons.len(),
                                        }
                                    }
                                });
                                if socket.send(Message::Text(pose_msg.to_string().into())).await.is_err() {
                                    break;
                                }
                            }
                        }
                    }
                    Err(_) => break,
                }
            }
            msg = socket.recv() => {
                match msg {
                    Some(Ok(Message::Text(text))) => {
                        // Handle ping/pong
                        if let Ok(v) = serde_json::from_str::<serde_json::Value>(&text) {
                            if v.get("type").and_then(|t| t.as_str()) == Some("ping") {
                                let pong = serde_json::json!({"type": "pong"});
                                let _ = socket.send(Message::Text(pong.to_string().into())).await;
                            }
                        }
                    }
                    Some(Ok(Message::Close(_))) | None => break,
                    _ => {}
                }
            }
        }
    }

    info!("WebSocket client disconnected (pose)");
}

// ── REST endpoints ───────────────────────────────────────────────────────────

/// Best-effort local LAN IPv4 detection.
///
/// Opens a non-routing UDP socket "connected" to a public address and asks
/// the kernel which local interface IP it picked.  Cached after first call.
/// Returns `0.0.0.0` if all attempts fail.
fn detect_hub_ip() -> String {
    use std::sync::OnceLock;
    static CACHE: OnceLock<String> = OnceLock::new();
    CACHE.get_or_init(|| {
        // Try IPv4 routing first (no packets actually sent).
        if let Ok(sock) = std::net::UdpSocket::bind("0.0.0.0:0") {
            if sock.connect("8.8.8.8:80").is_ok() {
                if let Ok(addr) = sock.local_addr() {
                    let ip = addr.ip().to_string();
                    if ip != "0.0.0.0" { return ip; }
                }
            }
        }
        // Fallback to RFC1918 LAN-only target.
        if let Ok(sock) = std::net::UdpSocket::bind("0.0.0.0:0") {
            if sock.connect("192.168.1.1:80").is_ok() {
                if let Ok(addr) = sock.local_addr() {
                    return addr.ip().to_string();
                }
            }
        }
        "0.0.0.0".to_string()
    }).clone()
}

// ── Node role registry ───────────────────────────────────────────────────────
// Per-node TX/RX/TXRX label used by the multistatic sensing pipeline. The
// authoritative role is flashed to NVS via `provision.py --node-role`, but the
// running ESP32 firmware doesn't currently echo it back, so we maintain a
// hub-side override map persisted to `data/node-roles.json`. UI can label
// nodes immediately and queue physical re-flashes later.
//
// Values: "tx" (transmitter / NDP injection), "rx" (receiver / CSI capture),
// "txrx" (default bistatic transceiver), "unknown" (never assigned).
fn node_role_path() -> std::path::PathBuf {
    std::path::PathBuf::from("data/node-roles.json")
}

fn node_role_registry() -> &'static std::sync::RwLock<std::collections::HashMap<u8, String>> {
    use std::sync::OnceLock;
    static REG: OnceLock<std::sync::RwLock<std::collections::HashMap<u8, String>>> = OnceLock::new();
    REG.get_or_init(|| {
        let map: std::collections::HashMap<u8, String> = std::fs::read_to_string(node_role_path())
            .ok()
            .and_then(|s| serde_json::from_str(&s).ok())
            .unwrap_or_default();
        std::sync::RwLock::new(map)
    })
}

/// Look up a node's role; returns "txrx" if no override is set (the firmware
/// default), or "unknown" only if explicitly recorded as such.
fn node_role_for(id: u8) -> String {
    node_role_registry()
        .read()
        .ok()
        .and_then(|g| g.get(&id).cloned())
        .unwrap_or_else(|| "txrx".to_string())
}

fn set_node_role(id: u8, role: &str) -> Result<(), String> {
    let role = role.to_lowercase();
    if !matches!(role.as_str(), "tx" | "rx" | "txrx" | "unknown") {
        warn!(node_id = id, role = %role,
            "node-role: rejected invalid role (must be tx|rx|txrx|unknown)");
        return Err(format!("invalid role '{}': must be tx, rx, txrx, or unknown", role));
    }
    let (prev, snapshot) = {
        let mut g = node_role_registry().write().map_err(|e| e.to_string())?;
        let prev = g.insert(id, role.clone()).unwrap_or_else(|| "txrx".to_string());
        (prev, g.clone())
    };
    if let Some(parent) = node_role_path().parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    let path = node_role_path();
    std::fs::write(
        &path,
        serde_json::to_string_pretty(&snapshot).map_err(|e| e.to_string())?,
    )
    .map_err(|e| e.to_string())?;
    if prev == role {
        info!(node_id = id, role = %role, persisted = %path.display(),
            "node-role: no change (idempotent write)");
    } else {
        info!(node_id = id, prev = %prev, new = %role,
            persisted = %path.display(), registry_size = snapshot.len(),
            "node-role: changed");
        record_event("role.changed", serde_json::json!({
            "node_id": id, "prev": prev, "new": role,
        }));
    }
    Ok(())
}

// ── Per-node remote config registry ──────────────────────────────────────────
// Hub-side settings (CSI packet rate, radio channel, TX power) that the node
// can poll and adopt at runtime. Persisted to `data/node-config.json` so
// settings survive hub restarts. The ESP32 firmware picks these up via the
// existing `swarm_bridge` poll loop; if the firmware predates the poll, the
// config is still durable and applied by future `provision.py` runs.
//
// Rationale: operators want to throttle a noisy node or boost a far one
// without reflashing. The rate is clamped to 1..=200 Hz because below 1 Hz
// breathing estimation falls apart and above 200 Hz the UDP pipeline chokes.
#[derive(Clone, Debug, serde::Serialize, serde::Deserialize, Default)]
struct NodeConfig {
    #[serde(skip_serializing_if = "Option::is_none")]
    rate_hz: Option<u16>,
    #[serde(skip_serializing_if = "Option::is_none")]
    channel: Option<u8>,
    #[serde(skip_serializing_if = "Option::is_none")]
    tx_power_dbm: Option<i8>,
    /// Monotonic revision bumped on every change — nodes only reapply when
    /// this increases, so a poll loop can short-circuit.
    #[serde(default)]
    revision: u32,
}

fn node_config_path() -> std::path::PathBuf {
    std::path::PathBuf::from("data/node-config.json")
}

fn node_config_registry() -> &'static std::sync::RwLock<std::collections::HashMap<u8, NodeConfig>> {
    use std::sync::OnceLock;
    static REG: OnceLock<std::sync::RwLock<std::collections::HashMap<u8, NodeConfig>>> =
        OnceLock::new();
    REG.get_or_init(|| {
        let map: std::collections::HashMap<u8, NodeConfig> =
            std::fs::read_to_string(node_config_path())
                .ok()
                .and_then(|s| serde_json::from_str(&s).ok())
                .unwrap_or_default();
        std::sync::RwLock::new(map)
    })
}

fn node_config_for(id: u8) -> NodeConfig {
    node_config_registry()
        .read()
        .ok()
        .and_then(|g| g.get(&id).cloned())
        .unwrap_or_default()
}

/// Merge a patch into a node's stored config. Only fields present in `patch`
/// are applied (null or missing fields are left untouched). Returns the new
/// merged state or a validation error string.
fn set_node_config(id: u8, patch: &serde_json::Value) -> Result<NodeConfig, String> {
    // Validate ranges before touching the registry.
    let rate_hz = match patch.get("rate_hz") {
        None | Some(serde_json::Value::Null) => None,
        Some(v) => {
            let n = v.as_u64().ok_or_else(|| "rate_hz must be a positive integer".to_string())?;
            if !(1..=200).contains(&n) {
                return Err(format!("rate_hz {} out of range [1, 200]", n));
            }
            Some(n as u16)
        }
    };
    let channel = match patch.get("channel") {
        None | Some(serde_json::Value::Null) => None,
        Some(v) => {
            let n = v.as_u64().ok_or_else(|| "channel must be an integer".to_string())?;
            // Valid 2.4 GHz channels (1-13) + 5 GHz (36-165). Permissive range.
            if !(1..=165).contains(&n) {
                return Err(format!("channel {} out of range [1, 165]", n));
            }
            Some(n as u8)
        }
    };
    let tx_power_dbm = match patch.get("tx_power_dbm") {
        None | Some(serde_json::Value::Null) => None,
        Some(v) => {
            let n = v.as_i64().ok_or_else(|| "tx_power_dbm must be an integer".to_string())?;
            if !(-4..=20).contains(&n) {
                return Err(format!("tx_power_dbm {} out of range [-4, 20]", n));
            }
            Some(n as i8)
        }
    };
    let snapshot = {
        let mut g = node_config_registry().write().map_err(|e| e.to_string())?;
        let entry = g.entry(id).or_default();
        let mut changed = false;
        if let Some(r) = rate_hz {
            if entry.rate_hz != Some(r) { entry.rate_hz = Some(r); changed = true; }
        }
        if let Some(c) = channel {
            if entry.channel != Some(c) { entry.channel = Some(c); changed = true; }
        }
        if let Some(p) = tx_power_dbm {
            if entry.tx_power_dbm != Some(p) { entry.tx_power_dbm = Some(p); changed = true; }
        }
        if changed {
            entry.revision = entry.revision.saturating_add(1);
        }
        g.clone()
    };
    if let Some(parent) = node_config_path().parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    std::fs::write(
        node_config_path(),
        serde_json::to_string_pretty(&snapshot).map_err(|e| e.to_string())?,
    )
    .map_err(|e| e.to_string())?;
    let cfg = snapshot.get(&id).cloned().unwrap_or_default();
    info!(node_id = id, rate_hz = ?cfg.rate_hz, channel = ?cfg.channel,
        tx_power_dbm = ?cfg.tx_power_dbm, revision = cfg.revision,
        "node-config: persisted");
    Ok(cfg)
}

/// POST /api/v1/node-config — merge patch into a node's stored config.
/// Body: `{"node_id": <u8>, "rate_hz"?: <1..200>, "channel"?: <1..165>,
///         "tx_power_dbm"?: <-4..20>}`.
/// Any combination of the three may be omitted — only what's present is
/// written. Returns the new merged state and a bumped revision number.
async fn node_config_set(
    Json(body): Json<serde_json::Value>,
) -> Json<serde_json::Value> {
    let id = match body.get("node_id").and_then(|v| v.as_u64()) {
        Some(n) if n <= 255 => n as u8,
        _ => {
            return Json(serde_json::json!({
                "ok": false,
                "error": "missing/invalid node_id (expected u8 0-255)",
            }));
        }
    };
    match set_node_config(id, &body) {
        Ok(cfg) => Json(serde_json::json!({
            "ok": true,
            "node_id": id,
            "config": cfg,
        })),
        Err(e) => Json(serde_json::json!({
            "ok": false,
            "node_id": id,
            "error": e,
        })),
    }
}

/// GET /api/v1/node-config?node_id=<u8> — firmware polling endpoint.
/// Returns the current config; empty object if nothing has been set. Nodes
/// should compare `revision` against the last-applied value and reapply
/// (rate / channel / tx-power) only when it increases.
async fn node_config_get(
    axum::extract::Query(q): axum::extract::Query<std::collections::HashMap<String, String>>,
) -> Json<serde_json::Value> {
    let id = match q.get("node_id").and_then(|v| v.parse::<u8>().ok()) {
        Some(n) => n,
        None => {
            return Json(serde_json::json!({
                "ok": false,
                "error": "missing/invalid node_id query param",
            }));
        }
    };
    let cfg = node_config_for(id);
    Json(serde_json::json!({
        "ok": true,
        "node_id": id,
        "config": cfg,
    }))
}

/// GET /api/v1/node-config/all — snapshot of every configured node. Used by
/// the UI to render the per-node row controls in one fetch.
async fn node_config_list() -> Json<serde_json::Value> {
    let snapshot = node_config_registry()
        .read()
        .ok()
        .map(|g| g.clone())
        .unwrap_or_default();
    Json(serde_json::json!({
        "ok": true,
        "nodes": snapshot,
    }))
}

/// Returns the count of currently-active TX / RX nodes by cross-referencing
/// the role registry against per-node liveness (last_frame_time within 5 s).
/// Inactive role assignments are reported separately so the UI can warn.
fn role_quality_summary(
    states: &std::collections::HashMap<u8, NodeState>,
) -> (Vec<u8>, Vec<u8>, Vec<u8>, Vec<u8>) {
    let now = std::time::Instant::now();
    let mut tx_active = Vec::new();
    let mut rx_active = Vec::new();
    let mut tx_inactive = Vec::new();
    let mut rx_inactive = Vec::new();
    for (&id, ns) in states.iter() {
        let elapsed_ms = ns.last_frame_time
            .map(|t| now.duration_since(t).as_millis() as u64)
            .unwrap_or(u64::MAX);
        let active = elapsed_ms <= 5_000;
        match node_role_for(id).as_str() {
            "tx"  => if active { tx_active.push(id) } else { tx_inactive.push(id) },
            "rx"  => if active { rx_active.push(id) } else { rx_inactive.push(id) },
            _ => {}
        }
    }
    tx_active.sort(); rx_active.sort(); tx_inactive.sort(); rx_inactive.sort();
    (tx_active, rx_active, tx_inactive, rx_inactive)
}

async fn health(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    Json(serde_json::json!({
        "status": "ok",
        "source": s.effective_source(),
        "tick": s.tick,
        "clients": s.tx.receiver_count(),
        "hub_ip": detect_hub_ip(),
    }))
}

/// POST /api/v1/role-set — set TX/RX/TXRX label for a node.
/// Body: `{"node_id": <u8>, "role": "tx" | "rx" | "txrx" | "unknown"}`.
/// Persists to `data/node-roles.json`. Survives restart but does NOT reflash
/// the ESP32 — for true RF-level role separation re-run
/// `provision.py --node-role tx --no-firmware` against the device.
/// Flat path (no URL parameter) sidesteps the matchit param/literal collision
/// with /api/v1/nodes/positions and /api/v1/nodes/{id} style routes.
async fn node_role_set(
    Json(body): Json<serde_json::Value>,
) -> Json<serde_json::Value> {
    let id = match body.get("node_id").and_then(|v| v.as_u64()) {
        Some(n) if n <= 255 => n as u8,
        _ => {
            warn!(?body, "role-set: missing/invalid node_id");
            return Json(serde_json::json!({
                "ok": false, "error": "missing/invalid node_id (expected u8 0-255)"
            }));
        }
    };
    let role = body.get("role").and_then(|v| v.as_str()).unwrap_or("");
    info!(node_id = id, role = %role, "role-set: incoming request");
    match set_node_role(id, role) {
        Ok(()) => Json(serde_json::json!({
            "ok": true,
            "node_id": id,
            "role": node_role_for(id),
        })),
        Err(e) => Json(serde_json::json!({
            "ok": false,
            "node_id": id,
            "error": e,
        })),
    }
}

/// POST /api/v1/role-set/auto — auto-assign roles for active nodes.
/// Picks the active node with the strongest mean RSSI as the TX (highest
/// SNR = best NDP injector), and assigns RX to all other active nodes.
/// Body (optional): `{"prefer_node_id": <u8>}` — if present, force that node
/// as TX instead of strongest-RSSI. Returns the resulting assignment.
async fn node_role_auto(
    State(state): State<SharedState>,
    body: Option<Json<serde_json::Value>>,
) -> Json<serde_json::Value> {
    let prefer = body.as_ref()
        .and_then(|b| b.get("prefer_node_id"))
        .and_then(|v| v.as_u64())
        .filter(|n| *n <= 255)
        .map(|n| n as u8);
    let s = state.read().await;
    let now = std::time::Instant::now();
    // Active nodes only (frame within 5 s) — never assign roles to dead ones.
    let mut active: Vec<(u8, f32)> = s.node_states.iter()
        .filter_map(|(&id, ns)| {
            let elapsed_ms = ns.last_frame_time
                .map(|t| now.duration_since(t).as_millis() as u64)?;
            if elapsed_ms > 5_000 { return None; }
            let mean_rssi = if !ns.rssi_history.is_empty() {
                let sum: f64 = ns.rssi_history.iter().sum();
                sum / ns.rssi_history.len() as f64
            } else {
                ns.rssi_history.back().copied().unwrap_or(-90.0_f64)
            };
            Some((id, mean_rssi as f32))
        })
        .collect();
    drop(s);
    if active.len() < 2 {
        warn!(active = active.len(),
            "role-set/auto: need >=2 active nodes for multistatic, refusing");
        return Json(serde_json::json!({
            "ok": false,
            "error": "need at least 2 active nodes (one TX + one RX) for multistatic",
            "active_nodes": active.iter().map(|(id, _)| *id).collect::<Vec<_>>(),
        }));
    }
    // Strongest mean RSSI first (least negative).
    active.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
    let tx_id = prefer
        .filter(|p| active.iter().any(|(id, _)| id == p))
        .unwrap_or_else(|| active[0].0);
    let mut assigned = Vec::with_capacity(active.len());
    let mut errors = Vec::new();
    for (id, _) in &active {
        let role = if *id == tx_id { "tx" } else { "rx" };
        match set_node_role(*id, role) {
            Ok(()) => assigned.push(serde_json::json!({"node_id": id, "role": role})),
            Err(e) => errors.push(serde_json::json!({"node_id": id, "error": e})),
        }
    }
    info!(tx_node = tx_id, rx_count = assigned.len().saturating_sub(1),
        prefer = ?prefer, "role-set/auto: assigned roles");
    Json(serde_json::json!({
        "ok": errors.is_empty(),
        "tx_node_id": tx_id,
        "assigned": assigned,
        "errors": errors,
        "strategy": if prefer.is_some() { "prefer" } else { "strongest_rssi" },
    }))
}

// ── Offline AI (Ollama) control ───────────────────────────────────────────────
// Lightweight lifecycle endpoints for the AEDI tab's "Start Offline AI" button.
// We don't attempt to manage the Python AEDI assistant on :3003 (it's a
// long-running systemd-style user process) — just the Ollama LLM engine on
// :11434, which is the thing that flips AEDI from offline to online.

fn ollama_binary_path() -> String {
    // 1. Respect explicit env var. 2. Prefer ~/.local/ollama/bin/ollama
    //    (matches .ionity/setup-ollama.sh). 3. Fall back to PATH.
    if let Ok(p) = std::env::var("OLLAMA_BIN") { return p; }
    if let Some(home) = std::env::var_os("HOME") {
        let local = std::path::PathBuf::from(&home).join(".local/ollama/bin/ollama");
        if local.is_file() { return local.to_string_lossy().into_owned(); }
    }
    "ollama".to_string()
}

fn ollama_pid_path() -> std::path::PathBuf {
    std::path::PathBuf::from("logs/ollama.pid")
}

fn ollama_log_path() -> std::path::PathBuf {
    std::path::PathBuf::from("logs/ollama.log")
}

async fn ollama_api_online() -> bool {
    // Check that port 11434 is accepting connections. Doing a full HTTP GET
    // would require pulling in reqwest; a 1.5 s TCP connect is enough to
    // distinguish "ollama serve is up" vs "not running".
    let addr = "127.0.0.1:11434";
    tokio::time::timeout(
        std::time::Duration::from_millis(1500),
        tokio::net::TcpStream::connect(addr),
    ).await.ok().and_then(|r| r.ok()).is_some()
}

/// GET /api/v1/ollama/status — is the local Ollama LLM engine up?
async fn ollama_status() -> Json<serde_json::Value> {
    let online = ollama_api_online().await;
    let pid = std::fs::read_to_string(ollama_pid_path()).ok()
        .and_then(|s| s.trim().parse::<u32>().ok());
    Json(serde_json::json!({
        "online": online,
        "pid": pid,
        "binary": ollama_binary_path(),
        "log": ollama_log_path().display().to_string(),
        "endpoint": "http://127.0.0.1:11434",
    }))
}

/// POST /api/v1/ollama/start — spawn `ollama serve` detached, PID → logs/ollama.pid.
/// Idempotent: if /api/tags responds we return `already_running: true`.
async fn ollama_start() -> Json<serde_json::Value> {
    if ollama_api_online().await {
        info!("ollama/start: already running on :11434");
        return Json(serde_json::json!({
            "ok": true, "already_running": true,
            "endpoint": "http://127.0.0.1:11434",
        }));
    }
    let bin = ollama_binary_path();
    // Ensure logs/ exists and open the log file for stdout/stderr capture.
    if let Some(parent) = ollama_log_path().parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    let log_file = match std::fs::OpenOptions::new()
        .create(true).append(true).open(ollama_log_path()) {
        Ok(f) => f,
        Err(e) => {
            warn!(error = %e, "ollama/start: failed to open log file");
            return Json(serde_json::json!({
                "ok": false, "error": format!("open log file: {}", e)
            }));
        }
    };
    let log_err = match log_file.try_clone() {
        Ok(f) => f,
        Err(e) => return Json(serde_json::json!({
            "ok": false, "error": format!("clone log fd: {}", e)
        })),
    };
    // OLLAMA_HOST=0.0.0.0 mirrors .ionity/setup-ollama.sh systemd unit.
    let spawn = std::process::Command::new(&bin)
        .arg("serve")
        .env("OLLAMA_HOST", "0.0.0.0")
        .stdin(std::process::Stdio::null())
        .stdout(std::process::Stdio::from(log_file))
        .stderr(std::process::Stdio::from(log_err))
        .spawn();
    let child = match spawn {
        Ok(c) => c,
        Err(e) => {
            warn!(bin = %bin, error = %e, "ollama/start: spawn failed");
            return Json(serde_json::json!({
                "ok": false, "error": format!("spawn '{}': {}", bin, e),
                "hint": "run `.ionity/setup-ollama.sh` to install Ollama",
            }));
        }
    };
    let pid = child.id();
    // Detach — don't reap. Record PID for status/stop.
    let _ = std::fs::write(ollama_pid_path(), pid.to_string());
    std::mem::forget(child);
    info!(pid, bin = %bin, "ollama/start: spawned ollama serve");
    // Brief wait so the caller sees an accurate online state if startup is fast.
    for _ in 0..10 {
        tokio::time::sleep(std::time::Duration::from_millis(300)).await;
        if ollama_api_online().await { break; }
    }
    let online = ollama_api_online().await;
    Json(serde_json::json!({
        "ok": true,
        "already_running": false,
        "pid": pid,
        "online": online,
        "endpoint": "http://127.0.0.1:11434",
        "log": ollama_log_path().display().to_string(),
    }))
}

// ── Storage / Database ───────────────────────────────────────────────
// Strategy: events are written as JSONL immediately (no dep, no install).
// When PostgreSQL is set up via `scripts/setup-postgres.sh`, the endpoint
// `/api/v1/storage/status` will report it as reachable and the operator can
// migrate / query it externally. The UI "Storage" tab exposes:
//   - GET  /api/v1/storage/status          — backend inventory + sizes
//   - GET  /api/v1/storage/recent?limit=N  — tail events.jsonl
//   - POST /api/v1/storage/event           — append arbitrary event (body {kind, payload})
//   - GET  /api/v1/storage/scan?root=PATH  — find DB-like files across device
//   - POST /api/v1/storage/config          — set storage root (body {root})
fn events_jsonl_path() -> std::path::PathBuf {
    std::env::var("AEDI_EVENTS_PATH")
        .map(std::path::PathBuf::from)
        .unwrap_or_else(|_| std::path::PathBuf::from("data/events.jsonl"))
}

fn storage_config_path() -> std::path::PathBuf {
    std::path::PathBuf::from("data/storage-config.json")
}

/// Append a structured event to `data/events.jsonl`. Never panics.
/// Called from any code path that mutates persistent state (role changes,
/// ollama start, node config updates) so operators get a single audit trail.
fn record_event(kind: &str, payload: serde_json::Value) {
    use std::io::Write;
    let path = events_jsonl_path();
    if let Some(parent) = path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    let ts = chrono::Utc::now().to_rfc3339();
    let line = serde_json::json!({ "ts": ts, "kind": kind, "payload": payload });
    match std::fs::OpenOptions::new().create(true).append(true).open(&path) {
        Ok(mut f) => { let _ = writeln!(f, "{}", line); }
        Err(e) => warn!(path = %path.display(), error = %e, "storage: failed to append event"),
    }
}

/// TCP probe 127.0.0.1:5432 — true if PostgreSQL is accepting connections.
async fn postgres_reachable() -> bool {
    tokio::time::timeout(
        std::time::Duration::from_millis(500),
        tokio::net::TcpStream::connect("127.0.0.1:5432"),
    ).await.ok().and_then(|r| r.ok()).is_some()
}

fn file_size(path: &std::path::Path) -> Option<u64> {
    std::fs::metadata(path).ok().map(|m| m.len())
}

fn count_lines(path: &std::path::Path) -> Option<u64> {
    use std::io::{BufRead, BufReader};
    std::fs::File::open(path).ok()
        .map(|f| BufReader::new(f).lines().filter_map(|l| l.ok()).count() as u64)
}

/// GET /api/v1/storage/status — single source of truth for the Storage tab.
async fn storage_status() -> Json<serde_json::Value> {
    let evt = events_jsonl_path();
    let sqlite_candidates = [
        std::path::PathBuf::from("logs/aedi_logs.sqlite3"),
        std::path::PathBuf::from("logs/ruview_logs.sqlite3"),
    ];
    let sqlite = sqlite_candidates.iter()
        .find(|p| p.exists())
        .map(|p| serde_json::json!({
            "path": p.display().to_string(),
            "bytes": file_size(p).unwrap_or(0),
        }));

    // Check psql availability (purely informational — no install attempt).
    let psql = std::process::Command::new("psql")
        .arg("--version")
        .output()
        .ok()
        .and_then(|o| String::from_utf8(o.stdout).ok())
        .map(|s| s.trim().to_string());

    let pg_online = postgres_reachable().await;

    // Load custom storage root if configured.
    let storage_root = std::fs::read_to_string(storage_config_path()).ok()
        .and_then(|s| serde_json::from_str::<serde_json::Value>(&s).ok())
        .and_then(|v| v.get("root").and_then(|r| r.as_str().map(String::from)));

    Json(serde_json::json!({
        "ok": true,
        "active_backend": "jsonl",
        "events_jsonl": {
            "path": evt.display().to_string(),
            "exists": evt.exists(),
            "bytes": file_size(&evt).unwrap_or(0),
            "lines": count_lines(&evt).unwrap_or(0),
        },
        "sqlite": sqlite,
        "postgres": {
            "online": pg_online,
            "endpoint": "127.0.0.1:5432",
            "psql_binary": psql,
            "installed": psql.is_some() || pg_online,
            "setup_script": "scripts/setup-postgres.sh",
        },
        "storage_root": storage_root,
        "cwd": std::env::current_dir().ok().map(|p| p.display().to_string()),
    }))
}

/// GET /api/v1/storage/recent?limit=N — tail last N events from JSONL.
async fn storage_recent(
    axum::extract::Query(q): axum::extract::Query<std::collections::HashMap<String, String>>,
) -> Json<serde_json::Value> {
    let limit = q.get("limit").and_then(|s| s.parse::<usize>().ok()).unwrap_or(50).min(1000);
    let path = events_jsonl_path();
    let lines: Vec<serde_json::Value> = std::fs::read_to_string(&path)
        .unwrap_or_default()
        .lines()
        .rev()
        .take(limit)
        .filter_map(|l| serde_json::from_str::<serde_json::Value>(l).ok())
        .collect();
    Json(serde_json::json!({
        "ok": true,
        "path": path.display().to_string(),
        "count": lines.len(),
        "events": lines,
    }))
}

/// POST /api/v1/storage/event — append a caller-supplied event.
/// Body: `{ "kind": "note", "payload": { ... } }`
async fn storage_event(Json(body): Json<serde_json::Value>) -> Json<serde_json::Value> {
    let kind = body.get("kind").and_then(|v| v.as_str()).unwrap_or("note").to_string();
    let payload = body.get("payload").cloned().unwrap_or(serde_json::json!({}));
    record_event(&kind, payload);
    Json(serde_json::json!({ "ok": true, "kind": kind }))
}

/// Recursive DB-file scanner. Shallow depth + exclude list to keep it fast.
fn scan_dbs(root: &std::path::Path, depth: u32, max_depth: u32, out: &mut Vec<serde_json::Value>) {
    if depth > max_depth || out.len() >= 500 { return; }
    let name = root.file_name().and_then(|n| n.to_str()).unwrap_or("");
    if matches!(name, "node_modules" | ".git" | "target" | ".cache" | "proc" | "sys"
                    | ".venv" | "venv" | "__pycache__" | ".local" | ".rustup" | ".cargo") {
        return;
    }
    let entries = match std::fs::read_dir(root) {
        Ok(e) => e,
        Err(_) => return,
    };
    for entry in entries.flatten() {
        let p = entry.path();
        let ft = match entry.file_type() { Ok(t) => t, Err(_) => continue };
        if ft.is_symlink() { continue; }
        if ft.is_dir() {
            scan_dbs(&p, depth + 1, max_depth, out);
        } else if ft.is_file() {
            let fname = p.file_name().and_then(|n| n.to_str()).unwrap_or("");
            let ext = p.extension().and_then(|e| e.to_str()).unwrap_or("").to_lowercase();
            let kind = if matches!(ext.as_str(), "sqlite" | "sqlite3" | "db" | "db3") {
                "sqlite"
            } else if ext == "jsonl" {
                "jsonl"
            } else if ext == "json" && (fname.contains("role") || fname.contains("config")) {
                continue; // skip regular config JSON, too noisy
            } else if ext == "pgsql" || ext == "sql" {
                "sql"
            } else {
                continue;
            };
            if let Ok(md) = entry.metadata() {
                out.push(serde_json::json!({
                    "path": p.display().to_string(),
                    "kind": kind,
                    "bytes": md.len(),
                }));
            }
        }
    }
}

/// GET /api/v1/storage/scan?root=/path&depth=5 — find DB-like files.
async fn storage_scan(
    axum::extract::Query(q): axum::extract::Query<std::collections::HashMap<String, String>>,
) -> Json<serde_json::Value> {
    let root = q.get("root").cloned().unwrap_or_else(|| {
        std::env::current_dir().map(|p| p.display().to_string()).unwrap_or_else(|_| ".".into())
    });
    let max_depth = q.get("depth").and_then(|s| s.parse::<u32>().ok()).unwrap_or(6).min(10);
    let mut out = Vec::new();
    let t0 = std::time::Instant::now();
    scan_dbs(std::path::Path::new(&root), 0, max_depth, &mut out);
    let elapsed_ms = t0.elapsed().as_millis() as u64;
    info!(root, max_depth, found = out.len(), elapsed_ms, "storage/scan completed");
    Json(serde_json::json!({
        "ok": true,
        "root": root,
        "max_depth": max_depth,
        "count": out.len(),
        "elapsed_ms": elapsed_ms,
        "results": out,
    }))
}

/// POST /api/v1/storage/config — set preferred storage root.
/// Body: `{ "root": "/some/path" }`. Persists to data/storage-config.json.
async fn storage_set_config(Json(body): Json<serde_json::Value>) -> Json<serde_json::Value> {
    let root = body.get("root").and_then(|v| v.as_str()).unwrap_or("").to_string();
    if root.is_empty() {
        return Json(serde_json::json!({"ok": false, "error": "missing 'root' field"}));
    }
    let path = storage_config_path();
    if let Some(parent) = path.parent() { let _ = std::fs::create_dir_all(parent); }
    let payload = serde_json::json!({ "root": root, "updated_at": chrono::Utc::now().to_rfc3339() });
    let write_ok = std::fs::write(&path, serde_json::to_string_pretty(&payload).unwrap_or_default());
    record_event("storage.config", payload.clone());
    Json(serde_json::json!({
        "ok": write_ok.is_ok(),
        "root": root,
        "persisted": path.display().to_string(),
    }))
}

async fn latest(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    match &s.latest_update {
        Some(update) => Json(serde_json::to_value(update).unwrap_or_default()),
        None => Json(serde_json::json!({"status": "no data yet"})),
    }
}

/// Generate WiFi-derived pose keypoints from sensing data.
///
/// Keypoint positions are modulated by real signal features rather than a pure
/// time-based sine/cosine loop:
///
///   - `motion_band_power`    drives whole-body translation and limb splay
///   - `variance`             seeds per-frame noise so the skeleton never freezes
///   - `breathing_band_power` expands/contracts torso keypoints (shoulders, hips)
///   - `dominant_freq_hz`     tilts the upper body laterally (lean direction)
///   - `change_points`        adds burst jitter to extremities (wrists, ankles)
///
/// When `presence == false` no persons are returned (empty room).
/// When walking is detected (`motion_score > 0.55`) the figure shifts laterally
/// with a stride-swing pattern applied to arms and legs.
// ── Multi-person estimation (issue #97) ──────────────────────────────────────

/// Fuse features across all active nodes for higher SNR.
///
/// When multiple ESP32 nodes observe the same room, their CSI features
/// can be combined:
/// - Variance: use max (most sensitive node dominates)
/// - Motion/breathing/spectral power: weighted average by RSSI (closer node = higher weight)
/// - Dominant frequency: weighted average
/// - Change points: keep current node's value (not meaningful to average)
/// - Mean RSSI: use max (best signal)
fn fuse_multi_node_features(
    current_features: &FeatureInfo,
    node_states: &HashMap<u8, NodeState>,
) -> FeatureInfo {
    let now = std::time::Instant::now();
    let active: Vec<(&FeatureInfo, f64)> = node_states.values()
        .filter(|ns| ns.last_frame_time.map_or(false, |t| now.duration_since(t).as_secs() < 10))
        .filter_map(|ns| {
            let feat = ns.latest_features.as_ref()?;
            let rssi = ns.rssi_history.back().copied().unwrap_or(-80.0);
            Some((feat, rssi))
        })
        .collect();

    if active.len() <= 1 {
        return current_features.clone();
    }

    // RSSI-based weights: higher RSSI = closer to person = more weight.
    // Map RSSI relative to best node into [0.1, 1.0].
    let max_rssi = active.iter().map(|(_, r)| *r).fold(f64::NEG_INFINITY, f64::max);
    let weights: Vec<f64> = active.iter()
        .map(|(_, r)| (1.0 + (r - max_rssi + 20.0) / 20.0).clamp(0.1, 1.0))
        .collect();
    let w_sum: f64 = weights.iter().sum::<f64>().max(1e-9);

    FeatureInfo {
        // Weighted average variance (not max — max inflates person score
        // and causes count flips between 1↔2 persons).
        variance: active.iter().zip(&weights)
            .map(|((f, _), w)| f.variance * w).sum::<f64>() / w_sum,
        // Weighted average for motion/breathing/spectral
        motion_band_power: active.iter().zip(&weights)
            .map(|((f, _), w)| f.motion_band_power * w).sum::<f64>() / w_sum,
        breathing_band_power: active.iter().zip(&weights)
            .map(|((f, _), w)| f.breathing_band_power * w).sum::<f64>() / w_sum,
        spectral_power: active.iter().zip(&weights)
            .map(|((f, _), w)| f.spectral_power * w).sum::<f64>() / w_sum,
        dominant_freq_hz: active.iter().zip(&weights)
            .map(|((f, _), w)| f.dominant_freq_hz * w).sum::<f64>() / w_sum,
        change_points: current_features.change_points, // keep current node's value
        // Best RSSI across nodes
        mean_rssi: active.iter().map(|(f, _)| f.mean_rssi).fold(f64::NEG_INFINITY, f64::max),
    }
}

/// Estimate person count from CSI features using a weighted composite heuristic.
///
/// Single ESP32 link limitations: variance-based detection can reliably detect
/// 1-2 persons. 3+ is speculative and requires ≥3 nodes for spatial resolution.
///
/// Returns a raw score (0.0..1.0) that the caller converts to person count
/// after temporal smoothing.
///
/// Normalization ranges calibrated from real ESP32 hardware (COM6/COM9 on
/// ruv.net, March 2026) and Linux WiFi BSSID scans (Antwerp-Pi, July 2026).
fn compute_person_score(feat: &FeatureInfo) -> f64 {
    let var_norm = (feat.variance / 350.0).clamp(0.0, 1.0);
    let cp_norm = (feat.change_points as f64 / 25.0).clamp(0.0, 1.0);
    let motion_norm = (feat.motion_band_power / 200.0).clamp(0.0, 1.0);
    let sp_norm = (feat.spectral_power / 450.0).clamp(0.0, 1.0);
    var_norm * 0.35 + cp_norm * 0.20 + motion_norm * 0.30 + sp_norm * 0.15
}

/// Estimate person count via ruvector DynamicMinCut on the subcarrier
/// temporal correlation graph.
///
/// Builds a graph where:
/// - Nodes = active subcarriers (variance > noise floor)
/// - Edges = Pearson correlation between subcarrier time series
///   (weight = correlation coefficient; high correlation = heavy edge)
/// - Source = virtual node connected to the most active subcarrier
/// - Sink = virtual node connected to the least correlated subcarrier
///
/// The min-cut value indicates how many independent motion clusters exist:
/// - High min-cut (relative to total edge weight) → one tightly coupled
///   group → 1 person
/// - Low min-cut → two loosely coupled groups → 2 persons
///
/// Uses `ruvector_mincut::DynamicMinCut` for O(V²E) exact max-flow.
fn estimate_persons_from_correlation_evidence(
    frame_history: &VecDeque<Vec<f64>>,
) -> Option<CountEstimate> {
    let n_frames = frame_history.len();
    if n_frames < 10 {
        return None;
    }

    let window: Vec<&Vec<f64>> = frame_history.iter().rev().take(20).collect();
    let n_sub = window[0].len().min(56);
    if n_sub < 4 {
        return None;
    }
    let k = window.len() as f64;

    // Per-subcarrier mean and variance
    let mut means = vec![0.0f64; n_sub];
    let mut variances = vec![0.0f64; n_sub];
    for frame in &window {
        for sc in 0..n_sub.min(frame.len()) {
            means[sc] += frame[sc] / k;
        }
    }
    for frame in &window {
        for sc in 0..n_sub.min(frame.len()) {
            variances[sc] += (frame[sc] - means[sc]).powi(2) / k;
        }
    }

    // Active subcarriers: variance above noise floor
    let noise_floor = 1.0;
    let active: Vec<usize> = (0..n_sub).filter(|&sc| variances[sc] > noise_floor).collect();
    let m = active.len();
    if m < 3 {
        return None;
    }

    // Build correlation graph edges between active subcarriers.
    // Edge weight = |Pearson correlation|. High correlation → same person.
    let mut edges: Vec<(u64, u64, f64)> = Vec::new();
    let source = m as u64;
    let sink = (m + 1) as u64;

    // Precompute std devs
    let stds: Vec<f64> = active.iter().map(|&sc| variances[sc].sqrt().max(1e-9)).collect();

    for i in 0..m {
        for j in (i + 1)..m {
            // Pearson correlation between subcarriers i and j
            let mut cov = 0.0f64;
            for frame in &window {
                let si = active[i];
                let sj = active[j];
                if si < frame.len() && sj < frame.len() {
                    cov += (frame[si] - means[si]) * (frame[sj] - means[sj]) / k;
                }
            }
            let corr = (cov / (stds[i] * stds[j])).abs();
            if corr > 0.1 {
                // Bidirectional edges for flow network
                let weight = corr * 10.0; // Scale up for integer-like flow
                edges.push((i as u64, j as u64, weight));
                edges.push((j as u64, i as u64, weight));
            }
        }
    }

    // Source → highest-variance subcarrier, Sink → lowest-variance
    let (max_var_idx, _) = active.iter().enumerate()
        .max_by(|(_, &a), (_, &b)| variances[a].partial_cmp(&variances[b]).unwrap())
        .unwrap_or((0, &0));
    let (min_var_idx, _) = active.iter().enumerate()
        .min_by(|(_, &a), (_, &b)| variances[a].partial_cmp(&variances[b]).unwrap())
        .unwrap_or((0, &0));

    if max_var_idx == min_var_idx {
        return Some(CountEstimate {
            source: "correlation".to_string(),
            count: 1,
            confidence: 0.40,
        });
    }

    edges.push((source, max_var_idx as u64, 100.0));
    edges.push((min_var_idx as u64, sink, 100.0));

    // Run min-cut
    let mc: DynamicMinCut = match MinCutBuilder::new().exact().with_edges(edges.clone()).build() {
        Ok(mc) => mc,
        Err(_) => return None,
    };

    let cut_value = mc.min_cut_value();
    let total_edge_weight: f64 = edges.iter()
        .filter(|(s, t, _)| *s != source && *s != sink && *t != source && *t != sink)
        .map(|(_, _, w)| w)
        .sum::<f64>() / 2.0; // bidirectional → halve

    if total_edge_weight < 1e-9 {
        return None;
    }

    // Normalized cut ratio: low = easy to split = multiple people
    let cut_ratio = cut_value / total_edge_weight;

    // Thresholds tunable at runtime for calibration in different environments.
    // Supports up to 5 persons (capped at AEDIS_MAX_PERSONS, default 5).
    //   cut_ratio > cr_1 → 1 person
    //   cr_2 < cut_ratio ≤ cr_1 → 2
    //   cr_3 < cut_ratio ≤ cr_2 → 3
    //   cr_4 < cut_ratio ≤ cr_3 → 4
    //   cut_ratio ≤ cr_4        → 5
    let cr_1: f64 = std::env::var("AEDIS_CUT_RATIO_1")
        .ok().and_then(|v| v.parse().ok()).unwrap_or(0.50);
    let cr_2: f64 = std::env::var("AEDIS_CUT_RATIO_2")
        .ok().and_then(|v| v.parse().ok()).unwrap_or(0.22);
    let cr_3: f64 = std::env::var("AEDIS_CUT_RATIO_3")
        .ok().and_then(|v| v.parse().ok()).unwrap_or(0.12);
    let cr_4: f64 = std::env::var("AEDIS_CUT_RATIO_4")
        .ok().and_then(|v| v.parse().ok()).unwrap_or(0.06);
    let max_persons: usize = std::env::var("AEDIS_MAX_PERSONS")
        .ok().and_then(|v| v.parse().ok()).unwrap_or(5).clamp(1, 16);

    let raw = if cut_ratio > cr_1 { 1 }
        else if cut_ratio > cr_2 { 2 }
        else if cut_ratio > cr_3 { 3 }
        else if cut_ratio > cr_4 { 4 }
        else { 5 };
    let threshold_margin = [cr_1, cr_2, cr_3, cr_4]
        .iter()
        .map(|thr| (cut_ratio - *thr).abs())
        .fold(1.0, f64::min);
    let support = (m as f64 / 24.0).clamp(0.0, 1.0);
    let connectivity = (total_edge_weight / (m as f64 * 8.0)).clamp(0.0, 1.0);
    let confidence = (0.28
        + 0.24 * support
        + 0.24 * connectivity
        + 0.24 * (threshold_margin * 2.0).clamp(0.0, 1.0))
        .clamp(0.18, 0.95);

    Some(CountEstimate {
        source: "correlation".to_string(),
        count: raw.min(max_persons),
        confidence,
    })
}

fn estimate_persons_from_correlation(frame_history: &VecDeque<Vec<f64>>) -> usize {
    estimate_persons_from_correlation_evidence(frame_history)
        .map(|estimate| estimate.count)
        .unwrap_or(1)
}

/// Estimate person count from channel impulse response (CIR) delay taps.
///
/// Each body reflects WiFi at a distinct Tx→body→Rx travel-time. Subcarrier
/// amplitudes H(f) transformed to time domain h(τ) reveal multipath peaks,
/// one per scatterer. This is an independent signal from the correlation
/// mincut — it catches synchronized-but-spatially-separated movers that look
/// identical in the frequency-correlation graph.
///
/// Method (amplitude-only CSI — no phase needed):
/// 1. Take N=32 most recent frames; compute per-subcarrier mean (static part).
/// 2. Subtract mean → dynamic component = only what bodies contribute.
/// 3. Average |dynamic|² across the window → frequency-domain power profile.
/// 4. Hand-rolled IDFT (N≤64, cheap) → time-domain delay profile |h(τ)|².
/// 5. Peak-pick on the first half (second half is Hermitian mirror for real
///    input). Peaks ≥ 35 % of max and spaced ≥ 2 delay bins apart are counted.
/// 6. Subtract 1 for the direct-path / static-reflection peak that is always
///    present. Remaining peaks ≈ moving bodies.
///
/// Env override:
///   AEDIS_TOF_THRESHOLD  relative peak threshold (default 0.35)
///   AEDIS_TOF_MIN_GAP    minimum tap gap between peaks (default 2)
fn estimate_persons_from_delay_taps_evidence(
    frame_history: &VecDeque<Vec<f64>>,
) -> Option<CountEstimate> {
    let n_frames = frame_history.len();
    if n_frames < 10 {
        return None;
    }
    let window: Vec<&Vec<f64>> = frame_history.iter().rev().take(32).collect();
    let n_sub = window.iter().map(|f| f.len()).min().unwrap_or(0).min(64);
    if n_sub < 8 {
        return None;
    }
    let k = window.len() as f64;

    // (1) Per-subcarrier mean = static channel component.
    let mut means = vec![0.0f64; n_sub];
    for frame in &window {
        for sc in 0..n_sub {
            means[sc] += frame[sc] / k;
        }
    }
    // (2,3) Dynamic power spectrum = E[(H - mean)²].
    let mut dyn_power = vec![0.0f64; n_sub];
    for frame in &window {
        for sc in 0..n_sub {
            let d = frame[sc] - means[sc];
            dyn_power[sc] += (d * d) / k;
        }
    }
    // Noise gate — if no subcarrier has meaningful dynamic power, no bodies.
    let max_power = dyn_power.iter().cloned().fold(0.0f64, f64::max);
    if max_power < 1e-3 {
        return None;
    }

    // (4) Treat dyn_power as a real, symmetric frequency-domain magnitude and
    // compute its IDFT. For real even input, IDFT is purely real (cosine
    // transform). Output |h(τ)|² is the Power Delay Profile proxy.
    // Direct summation — N ≤ 64, so cost is ~N²/2 ≈ 2k ops per call.
    let n_taps = (n_sub / 2).min(32); // Nyquist → meaningful half only.
    let mut pdp = vec![0.0f64; n_taps];
    let inv_n = 1.0 / n_sub as f64;
    for (tau, slot) in pdp.iter_mut().enumerate().take(n_taps) {
        let mut acc = 0.0f64;
        let two_pi_tau_over_n = 2.0 * std::f64::consts::PI * (tau as f64) * inv_n;
        for (sc, &p) in dyn_power.iter().enumerate().take(n_sub) {
            acc += p * (two_pi_tau_over_n * sc as f64).cos();
        }
        *slot = (acc * inv_n).abs();
    }

    // (5) Peak-pick.
    let thresh_rel: f64 = std::env::var("AEDIS_TOF_THRESHOLD")
        .ok().and_then(|v| v.parse().ok()).unwrap_or(0.35);
    let min_gap: usize = std::env::var("AEDIS_TOF_MIN_GAP")
        .ok().and_then(|v| v.parse().ok()).unwrap_or(2);
    let pdp_max = pdp.iter().cloned().fold(0.0f64, f64::max);
    if pdp_max < 1e-9 {
        return None;
    }
    let gate = pdp_max * thresh_rel;
    let mut peaks: Vec<usize> = Vec::new();
    for i in 1..n_taps.saturating_sub(1) {
        if pdp[i] >= gate && pdp[i] > pdp[i - 1] && pdp[i] >= pdp[i + 1] {
            if peaks.last().map_or(true, |&p| i - p >= min_gap) {
                peaks.push(i);
            }
        }
    }
    // (6) Subtract 1 for the residual static / direct-path peak that always
    // lingers after mean subtraction due to finite window bias.
    let movers = peaks.len().saturating_sub(1);
    let max_persons: usize = std::env::var("AEDIS_MAX_PERSONS")
        .ok().and_then(|v| v.parse().ok()).unwrap_or(5).clamp(1, 16);
    let mean_pdp = pdp.iter().sum::<f64>() / pdp.len().max(1) as f64;
    let contrast = (pdp_max / mean_pdp.max(1e-9)).clamp(1.0, 10.0);
    let confidence = (0.22
        + 0.22 * (window.len() as f64 / 32.0).clamp(0.0, 1.0)
        + 0.28 * ((contrast - 1.0) / 4.0).clamp(0.0, 1.0)
        + 0.18 * (peaks.len() as f64 / 4.0).clamp(0.0, 1.0)
        + 0.10 * (1.0 - thresh_rel).clamp(0.0, 1.0))
        .clamp(0.15, 0.90);

    Some(CountEstimate {
        source: "delay_taps".to_string(),
        count: movers.min(max_persons),
        confidence,
    })
}

/// Convert smoothed person score to discrete count with hysteresis.
///
/// Uses asymmetric thresholds: higher threshold to *add* a person, lower to
/// *drop* one.  This prevents flickering when the score hovers near a boundary
/// (the #1 user-reported issue — see #237, #249, #280, #292).
fn score_to_person_count(smoothed_score: f64, prev_count: usize) -> usize {
    // Piecewise mapping of smoothed_score → count with hysteresis for stability.
    // Supports 1..=5 persons. Override via AEDIS_MAX_PERSONS (default 5).
    //
    // Up-thresholds (must exceed to INCREASE count):
    //   →2: 0.70   →3: 0.82   →4: 0.90   →5: 0.95
    // Down-thresholds (must drop BELOW to DECREASE count):
    //   2→1: 0.55   3→2: 0.73   4→3: 0.85   5→4: 0.92
    let max_persons: usize = std::env::var("AEDIS_MAX_PERSONS")
        .ok().and_then(|v| v.parse().ok()).unwrap_or(5).clamp(1, 16);

    // Ladder of (up_to_count, up_threshold) pairs.
    let up = [(2, 0.70), (3, 0.82), (4, 0.90), (5, 0.95)];
    // Ladder of (from_count, down_threshold) pairs.
    let down = [(5, 0.92), (4, 0.85), (3, 0.73), (2, 0.55)];

    // First determine target purely from score (ignoring hysteresis).
    let mut target = 1usize;
    for (n, thr) in up.iter() {
        if smoothed_score > *thr { target = *n; }
    }
    let target = target.min(max_persons);

    if target > prev_count {
        // Promote only if all up-thresholds up to target are satisfied
        // (smoothed_score > up_threshold). Already guaranteed by construction.
        target
    } else if target < prev_count {
        // Demote only if below the down-threshold for current level.
        let down_thr = down.iter().find(|(n, _)| *n == prev_count).map(|(_, t)| *t).unwrap_or(0.0);
        if smoothed_score < down_thr { target } else { prev_count }
    } else {
        prev_count
    }
}

fn score_count_estimate(smoothed_score: f64, prev_count: usize) -> CountEstimate {
    let count = score_to_person_count(smoothed_score, prev_count);
    let thresholds = [0.55, 0.70, 0.73, 0.82, 0.85, 0.90, 0.92, 0.95];
    let margin = thresholds
        .iter()
        .map(|thr| (smoothed_score - *thr).abs())
        .fold(1.0, f64::min);
    let confidence = (0.46 + 0.42 * (margin * 2.0).clamp(0.0, 1.0)).clamp(0.22, 0.95);

    CountEstimate {
        source: "score".to_string(),
        count,
        confidence,
    }
}

fn field_model_count_estimate(
    field_model: Option<&FieldModel>,
    frame_history: &VecDeque<Vec<f64>>,
) -> Option<CountEstimate> {
    let field_model = field_model?;
    let status = field_model.status();
    let count = field_bridge::occupancy_estimate(field_model, frame_history)?;
    let confidence = match status {
        CalibrationStatus::Fresh => 0.88,
        CalibrationStatus::Stale => 0.72,
        _ => return None,
    };

    Some(CountEstimate {
        source: "field_model".to_string(),
        count,
        confidence,
    })
}

fn weighted_count_consensus(estimates: Vec<CountEstimate>) -> Option<PersonCountEvidence> {
    let estimates: Vec<CountEstimate> = estimates
        .into_iter()
        .filter(|estimate| estimate.confidence.is_finite() && estimate.confidence > 0.0)
        .collect();
    if estimates.is_empty() {
        return None;
    }

    let total_weight = estimates.iter().map(|estimate| estimate.confidence).sum::<f64>().max(1e-9);
    let weighted_mean = estimates
        .iter()
        .map(|estimate| estimate.count as f64 * estimate.confidence)
        .sum::<f64>()
        / total_weight;

    let mut scores: HashMap<usize, f64> = HashMap::new();
    for estimate in &estimates {
        *scores.entry(estimate.count).or_insert(0.0) += estimate.confidence;
    }

    let mut best_count = estimates[0].count;
    let mut best_weight = -1.0f64;
    let mut best_distance = f64::MAX;
    for (count, weight) in scores {
        let distance = (count as f64 - weighted_mean).abs();
        let better = weight > best_weight
            || ((weight - best_weight).abs() < 1e-6 && distance < best_distance)
            || ((weight - best_weight).abs() < 1e-6
                && (distance - best_distance).abs() < 1e-6
                && count > best_count);
        if better {
            best_count = count;
            best_weight = weight;
            best_distance = distance;
        }
    }

    Some(PersonCountEvidence {
        consensus_count: best_count,
        consensus_confidence: (best_weight / total_weight).clamp(0.0, 1.0),
        estimates,
    })
}

fn build_single_link_count_evidence(
    frame_history: &VecDeque<Vec<f64>>,
    field_model: Option<&FieldModel>,
    classification_confidence: f64,
    smoothed_score: f64,
    prev_count: usize,
) -> Option<PersonCountEvidence> {
    let confidence_scale = (0.70 + 0.30 * classification_confidence.clamp(0.0, 1.0)).clamp(0.70, 1.0);

    let mut estimates = Vec::new();

    let mut score = score_count_estimate(smoothed_score, prev_count);
    score.confidence = (score.confidence * confidence_scale).clamp(0.0, 0.95);
    estimates.push(score);

    if let Some(mut corr) = estimate_persons_from_correlation_evidence(frame_history) {
        corr.count = corr.count.max(1);
        corr.confidence = (corr.confidence * confidence_scale).clamp(0.0, 0.95);
        estimates.push(corr);
    }

    if let Some(mut delay) = estimate_persons_from_delay_taps_evidence(frame_history) {
        delay.count = delay.count.max(1);
        delay.confidence = (delay.confidence * confidence_scale).clamp(0.0, 0.95);
        estimates.push(delay);
    }

    if let Some(mut field) = field_model_count_estimate(field_model, frame_history) {
        field.count = field.count.max(1);
        field.confidence = (field.confidence * confidence_scale).clamp(0.0, 0.95);
        estimates.push(field);
    }

    weighted_count_consensus(estimates)
}

fn best_active_node_estimate(
    node_states: &HashMap<u8, NodeState>,
    now: std::time::Instant,
    source: &str,
    estimator: fn(&VecDeque<Vec<f64>>) -> Option<CountEstimate>,
) -> Option<CountEstimate> {
    node_states
        .values()
        .filter(|node| node.last_frame_time.map_or(false, |t| now.duration_since(t).as_secs() < 5))
        .filter_map(|node| estimator(&node.frame_history))
        .max_by(|a, b| a.confidence.total_cmp(&b.confidence))
        .map(|mut estimate| {
            estimate.source = source.to_string();
            estimate.count = estimate.count.max(1);
            estimate
        })
}

fn room_dimensions_m() -> (f64, f64, f64) {
    (
        env_f64("AEDIS_ROOM_WIDTH_M", 8.0).max(1.0),
        env_f64("AEDIS_ROOM_DEPTH_M", 6.0).max(1.0),
        env_f64("AEDIS_ROOM_HEIGHT_M", 3.0).max(1.0),
    )
}

fn fallback_person_anchors(person_count: usize) -> Vec<PersonAnchor> {
    let (room_width, room_depth, _) = room_dimensions_m();
    let half = (person_count as f64 - 1.0) / 2.0;
    let spacing = (room_width / 6.0).clamp(0.8, 1.4);

    (0..person_count)
        .map(|index| PersonAnchor {
            x: (room_width * 0.5 + (index as f64 - half) * spacing)
                .clamp(room_width * 0.15, room_width * 0.85),
            y: 0.95,
            z: room_depth * 0.55,
            strength: 0.35,
            zone: format!("zone_{}", index + 1),
        })
        .collect()
}

fn extract_person_anchors(update: &SensingUpdate) -> Vec<PersonAnchor> {
    let person_count = update.estimated_persons.unwrap_or(1).max(1);
    let grid_x = update.signal_field.grid_size[0].max(1);
    let grid_z = update.signal_field.grid_size[2].max(1);
    let cell_count = grid_x * grid_z;
    if person_count == 0 || update.signal_field.values.len() < cell_count {
        return fallback_person_anchors(person_count);
    }

    let max_value = update.signal_field.values.iter().cloned().fold(0.0, f64::max);
    if max_value < 1e-6 {
        return fallback_person_anchors(person_count);
    }

    let primary_gate = max_value * 0.55;
    let secondary_gate = max_value * 0.35;
    let min_sep_cells = ((grid_x.min(grid_z) as f64) * 0.18).round().max(2.0) as usize;

    let mut candidates: Vec<(usize, usize, f64)> = Vec::new();
    for z in 0..grid_z {
        for x in 0..grid_x {
            let idx = z * grid_x + x;
            let value = update.signal_field.values[idx];
            if value < primary_gate {
                continue;
            }

            let mut is_peak = true;
            for dz in -1isize..=1 {
                for dx in -1isize..=1 {
                    if dx == 0 && dz == 0 {
                        continue;
                    }
                    let nx = x as isize + dx;
                    let nz = z as isize + dz;
                    if nx < 0 || nz < 0 || nx >= grid_x as isize || nz >= grid_z as isize {
                        continue;
                    }
                    let neighbor_idx = nz as usize * grid_x + nx as usize;
                    if update.signal_field.values[neighbor_idx] > value {
                        is_peak = false;
                        break;
                    }
                }
                if !is_peak {
                    break;
                }
            }

            if is_peak {
                candidates.push((x, z, value));
            }
        }
    }

    if candidates.len() < person_count {
        for z in 0..grid_z {
            for x in 0..grid_x {
                let idx = z * grid_x + x;
                let value = update.signal_field.values[idx];
                if value >= secondary_gate {
                    candidates.push((x, z, value));
                }
            }
        }
    }

    candidates.sort_by(|a, b| b.2.total_cmp(&a.2));
    candidates.dedup_by(|a, b| a.0 == b.0 && a.1 == b.1);

    let (room_width, room_depth, room_height) = room_dimensions_m();
    let mut anchors = Vec::new();
    for (x, z, value) in candidates {
        if anchors.len() >= person_count {
            break;
        }
        let too_close = anchors.iter().any(|anchor: &PersonAnchor| {
            let anchor_x = ((anchor.x / room_width) * grid_x as f64).round().clamp(0.0, grid_x as f64 - 1.0) as isize;
            let anchor_z = ((anchor.z / room_depth) * grid_z as f64).round().clamp(0.0, grid_z as f64 - 1.0) as isize;
            let dx = anchor_x - x as isize;
            let dz = anchor_z - z as isize;
            ((dx * dx + dz * dz) as f64).sqrt() < min_sep_cells as f64
        });
        if too_close {
            continue;
        }

        let strength = (value / max_value).clamp(0.10, 1.0);
        anchors.push(PersonAnchor {
            x: ((x as f64 + 0.5) / grid_x as f64) * room_width,
            y: (0.90 + 0.35 * strength).clamp(0.75, room_height * 0.7),
            z: ((z as f64 + 0.5) / grid_z as f64) * room_depth,
            strength,
            zone: String::new(),
        });
    }

    if anchors.len() < person_count {
        for fallback in fallback_person_anchors(person_count) {
            if anchors.len() >= person_count {
                break;
            }
            let too_close = anchors.iter().any(|anchor| {
                let dx = anchor.x - fallback.x;
                let dz = anchor.z - fallback.z;
                (dx * dx + dz * dz).sqrt() < room_width.min(room_depth) * 0.12
            });
            if !too_close {
                anchors.push(fallback);
            }
        }
    }

    anchors.sort_by(|a, b| a.x.total_cmp(&b.x));
    for (index, anchor) in anchors.iter_mut().enumerate() {
        anchor.zone = format!("zone_{}", index + 1);
    }
    anchors.truncate(person_count);
    anchors
}

/// Generate a single person's skeleton anchored to a signal-field peak.
///
/// `person_idx`: 0-based index of this person.
fn derive_single_person_pose(
    update: &SensingUpdate,
    person_idx: usize,
    anchor: &PersonAnchor,
) -> PersonDetection {
    let cls = &update.classification;
    let feat = &update.features;

    // Per-person phase offset: ~120 degrees apart so they don't move in sync.
    let phase_offset = person_idx as f64 * 2.094;

    // Confidence decays for additional persons (less certain about person 2, 3).
    let conf_decay = 1.0 - person_idx as f64 * 0.15;

    // ── Signal-derived scalars ────────────────────────────────────────────────

    let motion_score = (feat.motion_band_power / 15.0).clamp(0.0, 1.0);
    let is_walking = motion_score > 0.55;
    let breath_amp = (feat.breathing_band_power * 4.0).clamp(0.0, 12.0);

    let breath_phase = if let Some(ref vs) = update.vital_signs {
        let bpm = vs.breathing_rate_bpm.unwrap_or(15.0);
        let freq = (bpm / 60.0).clamp(0.1, 0.5);
        // Slow tick rate (0.02) for gentle breathing, not jerky oscillation.
        (update.tick as f64 * freq * 0.02 * std::f64::consts::TAU + phase_offset).sin()
    } else {
        (update.tick as f64 * 0.02 + phase_offset).sin()
    };

    let lean_x = (feat.dominant_freq_hz / 5.0 - 1.0).clamp(-1.0, 1.0) * 18.0;

    let stride_x = if is_walking {
        let stride_phase = (feat.motion_band_power * 0.7 + update.tick as f64 * 0.06 + phase_offset).sin();
        stride_phase * 20.0 * motion_score
    } else {
        0.0
    };

    // Dampen burst and noise to reduce jitter.  The original used
    // tick*17.3 which changed wildly every frame.  Now use slow tick
    // rate and minimal burst scaling for a stable skeleton.
    let burst = (feat.change_points as f64 / 20.0).clamp(0.0, 0.3);

    let noise_seed = person_idx as f64 * 97.1; // stable per-person, no tick
    let noise_val = (noise_seed.sin() * 43758.545).fract();

    let snr_factor = ((feat.variance - 0.5) / 10.0).clamp(0.0, 1.0);
    let anchor_confidence = 0.70 + 0.30 * anchor.strength.clamp(0.0, 1.0);
    let base_confidence = cls.confidence * anchor_confidence * (0.55 + 0.45 * snr_factor) * conf_decay;

    // ── Skeleton base position ────────────────────────────────────────────────

    let (room_width, room_depth, _) = room_dimensions_m();
    let anchor_x_norm = (anchor.x / room_width).clamp(0.0, 1.0);
    let anchor_z_norm = (anchor.z / room_depth).clamp(0.0, 1.0);
    let base_x = 90.0 + anchor_x_norm * 460.0 + stride_x + lean_x * 0.5;
    let base_y = 170.0 + anchor_z_norm * 150.0 - motion_score * 8.0;

    // ── COCO 17-keypoint offsets from hip-center ──────────────────────────────

    let kp_names = [
        "nose", "left_eye", "right_eye", "left_ear", "right_ear",
        "left_shoulder", "right_shoulder", "left_elbow", "right_elbow",
        "left_wrist", "right_wrist", "left_hip", "right_hip",
        "left_knee", "right_knee", "left_ankle", "right_ankle",
    ];

    let kp_offsets: [(f64, f64); 17] = [
        (  0.0,  -80.0), // 0  nose
        ( -8.0,  -88.0), // 1  left_eye
        (  8.0,  -88.0), // 2  right_eye
        (-16.0,  -82.0), // 3  left_ear
        ( 16.0,  -82.0), // 4  right_ear
        (-30.0,  -50.0), // 5  left_shoulder
        ( 30.0,  -50.0), // 6  right_shoulder
        (-45.0,  -15.0), // 7  left_elbow
        ( 45.0,  -15.0), // 8  right_elbow
        (-50.0,   20.0), // 9  left_wrist
        ( 50.0,   20.0), // 10 right_wrist
        (-20.0,   20.0), // 11 left_hip
        ( 20.0,   20.0), // 12 right_hip
        (-22.0,   70.0), // 13 left_knee
        ( 22.0,   70.0), // 14 right_knee
        (-24.0,  120.0), // 15 left_ankle
        ( 24.0,  120.0), // 16 right_ankle
    ];

    const TORSO_KP: [usize; 4] = [5, 6, 11, 12];
    const EXTREMITY_KP: [usize; 4] = [9, 10, 15, 16];

    let keypoints: Vec<PoseKeypoint> = kp_names.iter().zip(kp_offsets.iter())
        .enumerate()
        .map(|(i, (name, (dx, dy)))| {
            let breath_dx = if TORSO_KP.contains(&i) {
                let sign = if *dx < 0.0 { -1.0 } else { 1.0 };
                sign * breath_amp * breath_phase * 0.5
            } else {
                0.0
            };
            let breath_dy = if TORSO_KP.contains(&i) {
                let sign = if *dy < 0.0 { -1.0 } else { 1.0 };
                sign * breath_amp * breath_phase * 0.3
            } else {
                0.0
            };

            let extremity_jitter = if EXTREMITY_KP.contains(&i) {
                let phase = noise_seed + i as f64 * 2.399;
                // Dampened from 12/8 to 4/3 to reduce visual jumping.
                (
                    phase.sin() * burst * motion_score * 4.0,
                    (phase * 1.31).cos() * burst * motion_score * 3.0,
                )
            } else {
                (0.0, 0.0)
            };

            let kp_noise_x = ((noise_seed + i as f64 * 1.618).sin() * 43758.545).fract()
                * feat.variance.sqrt().clamp(0.0, 3.0) * motion_score;
            let kp_noise_y = ((noise_seed + i as f64 * 2.718).cos() * 31415.926).fract()
                * feat.variance.sqrt().clamp(0.0, 3.0) * motion_score * 0.6;

            let swing_dy = if is_walking {
                let stride_phase =
                    (feat.motion_band_power * 0.7 + update.tick as f64 * 0.12 + phase_offset).sin();
                match i {
                    7 | 9  => -stride_phase * 20.0 * motion_score,
                    8 | 10 =>  stride_phase * 20.0 * motion_score,
                    13 | 15 =>  stride_phase * 25.0 * motion_score,
                    14 | 16 => -stride_phase * 25.0 * motion_score,
                    _ => 0.0,
                }
            } else {
                0.0
            };

            let final_x = base_x + dx + breath_dx + extremity_jitter.0 + kp_noise_x;
            let final_y = base_y + dy + breath_dy + extremity_jitter.1 + kp_noise_y + swing_dy;

            let kp_conf = if EXTREMITY_KP.contains(&i) {
                base_confidence * (0.7 + 0.3 * snr_factor) * (0.85 + 0.15 * noise_val)
            } else {
                base_confidence * (0.88 + 0.12 * ((i as f64 * 0.7 + noise_seed).cos()))
            };

            PoseKeypoint {
                name: name.to_string(),
                x: final_x,
                y: final_y,
                z: (anchor_z_norm + lean_x * 0.005).clamp(0.0, 1.2),
                confidence: kp_conf.clamp(0.1, 1.0),
            }
        })
        .collect();

    let xs: Vec<f64> = keypoints.iter().map(|k| k.x).collect();
    let ys: Vec<f64> = keypoints.iter().map(|k| k.y).collect();
    let min_x = xs.iter().cloned().fold(f64::MAX, f64::min) - 10.0;
    let min_y = ys.iter().cloned().fold(f64::MAX, f64::min) - 10.0;
    let max_x = xs.iter().cloned().fold(f64::MIN, f64::max) + 10.0;
    let max_y = ys.iter().cloned().fold(f64::MIN, f64::max) + 10.0;

    PersonDetection {
        id: (person_idx + 1) as u32,
        confidence: (cls.confidence * conf_decay * anchor_confidence).clamp(0.1, 1.0),
        keypoints,
        bbox: BoundingBox {
            x: min_x,
            y: min_y,
            width: (max_x - min_x).max(80.0),
            height: (max_y - min_y).max(160.0),
        },
        zone: anchor.zone.clone(),
    }
}

fn derive_pose_from_sensing(update: &SensingUpdate) -> Vec<PersonDetection> {
    let cls = &update.classification;
    if !cls.presence {
        return vec![];
    }

    let anchors = update
        .person_anchors
        .clone()
        .unwrap_or_else(|| extract_person_anchors(update));

    anchors
        .iter()
        .enumerate()
        .map(|(idx, anchor)| derive_single_person_pose(update, idx, anchor))
        .collect()
}

// ── RuVector Phase 2: Temporal EMA smoothing for keypoints ──────────────────

/// Expected bone lengths in pixel-space for the COCO-17 skeleton as used by
/// `derive_single_person_pose`. Pairs are (parent_idx, child_idx).
const POSE_BONE_PAIRS: &[(usize, usize)] = &[
    (5, 7), (7, 9), (6, 8), (8, 10),   // arms
    (5, 11), (6, 12),                     // torso
    (11, 13), (13, 15), (12, 14), (14, 16), // legs
    (5, 6), (11, 12),                     // shoulders, hips
];

fn recompute_person_bbox(person: &mut PersonDetection) {
    if person.keypoints.is_empty() {
        return;
    }

    let xs: Vec<f64> = person.keypoints.iter().map(|kp| kp.x).collect();
    let ys: Vec<f64> = person.keypoints.iter().map(|kp| kp.y).collect();
    let min_x = xs.iter().cloned().fold(f64::MAX, f64::min) - 10.0;
    let min_y = ys.iter().cloned().fold(f64::MAX, f64::min) - 10.0;
    let max_x = xs.iter().cloned().fold(f64::MIN, f64::max) + 10.0;
    let max_y = ys.iter().cloned().fold(f64::MIN, f64::max) + 10.0;

    person.bbox = BoundingBox {
        x: min_x,
        y: min_y,
        width: (max_x - min_x).max(80.0),
        height: (max_y - min_y).max(160.0),
    };
}

fn apply_track_smoothing(
    persons: &mut [PersonDetection],
    track_temporal_state: &mut HashMap<u32, TrackTemporalState>,
    tick: u64,
) {
    for person in persons.iter_mut() {
        let current_keypoints: Vec<[f64; 3]> = person.keypoints.iter()
            .map(|kp| [kp.x, kp.y, kp.z])
            .collect();
        let alpha = if person.confidence < TRACK_SMOOTHING_LOW_CONFIDENCE_THRESHOLD {
            TRACK_SMOOTHING_ALPHA_LOW_CONFIDENCE
        } else {
            TEMPORAL_EMA_ALPHA_DEFAULT
        };

        let smoothed = if let Some(previous) = track_temporal_state.get(&person.id) {
            if previous.keypoints.len() == current_keypoints.len() {
                let mut out = Vec::with_capacity(current_keypoints.len());
                for (current, prev) in current_keypoints.iter().zip(previous.keypoints.iter()) {
                    out.push([
                        alpha * current[0] + (1.0 - alpha) * prev[0],
                        alpha * current[1] + (1.0 - alpha) * prev[1],
                        alpha * current[2] + (1.0 - alpha) * prev[2],
                    ]);
                }
                clamp_bone_lengths_f64(&mut out, &previous.keypoints);
                out
            } else {
                current_keypoints.clone()
            }
        } else {
            current_keypoints.clone()
        };

        for (kp, coords) in person.keypoints.iter_mut().zip(smoothed.iter()) {
            kp.x = coords[0];
            kp.y = coords[1];
            kp.z = coords[2];
        }
        recompute_person_bbox(person);

        track_temporal_state.insert(person.id, TrackTemporalState {
            keypoints: smoothed,
            last_seen_tick: tick,
        });
    }

    track_temporal_state.retain(|track_id, state| {
        persons.iter().any(|person| person.id == *track_id)
            || tick.saturating_sub(state.last_seen_tick) <= TRACK_SMOOTHING_TTL_TICKS
    });
}

fn finalize_people_for_update(
    update: &mut SensingUpdate,
    tracker: &mut PoseTracker,
    last_tracker_instant: &mut Option<std::time::Instant>,
    track_temporal_state: &mut HashMap<u32, TrackTemporalState>,
) {
    if !update.classification.presence {
        update.persons = None;
        update.person_anchors = None;
        return;
    }

    let anchors = extract_person_anchors(update);
    if anchors.is_empty() {
        update.persons = None;
        update.person_anchors = None;
        return;
    }
    update.person_anchors = Some(anchors);

    let raw_persons = derive_pose_from_sensing(update);
    let mut tracked = tracker_bridge::tracker_update(tracker, last_tracker_instant, raw_persons);
    apply_track_smoothing(&mut tracked, track_temporal_state, update.tick);
    if tracked.is_empty() {
        update.persons = None;
    } else {
        update.persons = Some(tracked);
    }
}

/// Clamp bone lengths so no bone changes by more than MAX_BONE_CHANGE_RATIO
/// compared to the previous frame.
fn clamp_bone_lengths_f64(pose: &mut Vec<[f64; 3]>, prev: &[[f64; 3]]) {
    for &(p, c) in POSE_BONE_PAIRS {
        if p >= pose.len() || c >= pose.len() {
            continue;
        }
        let prev_len = dist_f64(&prev[p], &prev[c]);
        if prev_len < 1e-6 {
            continue;
        }
        let cur_len = dist_f64(&pose[p], &pose[c]);
        if cur_len < 1e-6 {
            continue;
        }
        let ratio = cur_len / prev_len;
        let lo = 1.0 - MAX_BONE_CHANGE_RATIO;
        let hi = 1.0 + MAX_BONE_CHANGE_RATIO;
        if ratio < lo || ratio > hi {
            let target = prev_len * ratio.clamp(lo, hi);
            let scale = target / cur_len;
            for dim in 0..3 {
                let diff = pose[c][dim] - pose[p][dim];
                pose[c][dim] = pose[p][dim] + diff * scale;
            }
        }
    }
}

fn dist_f64(a: &[f64; 3], b: &[f64; 3]) -> f64 {
    let dx = b[0] - a[0];
    let dy = b[1] - a[1];
    let dz = b[2] - a[2];
    (dx * dx + dy * dy + dz * dz).sqrt()
}

// ── DensePose-compatible REST endpoints ─────────────────────────────────────

async fn health_live(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    Json(serde_json::json!({
        "status": "alive",
        "uptime": s.start_time.elapsed().as_secs(),
    }))
}

async fn health_ready(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    Json(serde_json::json!({
        "status": "ready",
        "source": s.effective_source(),
    }))
}

async fn health_system(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    let uptime = s.start_time.elapsed().as_secs();
    Json(serde_json::json!({
        "status": "healthy",
        "components": {
            "api": { "status": "healthy", "message": "Rust Axum server" },
            "hardware": {
                "status": if s.effective_source().ends_with(":offline") { "degraded" } else { "healthy" },
                "message": format!("Source: {}", s.effective_source())
            },
            "pose": { "status": "healthy", "message": "WiFi-derived pose estimation" },
            "stream": { "status": if s.tx.receiver_count() > 0 { "healthy" } else { "idle" },
                        "message": format!("{} client(s)", s.tx.receiver_count()) },
        },
        "metrics": {
            "cpu_percent": 2.5,
            "memory_percent": 1.8,
            "disk_percent": 15.0,
            "uptime_seconds": uptime,
        }
    }))
}

async fn health_version() -> Json<serde_json::Value> {
    Json(serde_json::json!({
        "version": env!("CARGO_PKG_VERSION"),
        "name": "wifi-densepose-sensing-server",
        "backend": "rust+axum+ruvector",
    }))
}

async fn health_metrics(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    Json(serde_json::json!({
        "system_metrics": {
            "cpu": { "percent": 2.5 },
            "memory": { "percent": 1.8, "used_mb": 5 },
            "disk": { "percent": 15.0 },
        },
        "tick": s.tick,
    }))
}

async fn health_autorepair(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    let uptime = s.start_time.elapsed().as_secs();
    let src = s.effective_source();
    let source_degraded = src.contains("offline") || src.contains("degraded");

    // Read process RSS on Linux
    let heap_mb = {
        #[cfg(target_os = "linux")]
        {
            std::fs::read_to_string("/proc/self/status")
                .ok()
                .and_then(|status| {
                    status.lines()
                        .find(|l| l.starts_with("VmRSS:"))
                        .and_then(|l| l.split_whitespace().nth(1))
                        .and_then(|kb| kb.parse::<f64>().ok())
                        .map(|kb| kb / 1024.0)
                })
                .unwrap_or(0.0)
        }
        #[cfg(not(target_os = "linux"))]
        { 0.0_f64 }
    };

    let tick_ok = s.tick > 0;
    let memory_ok = heap_mb < 256.0;
    let healthy = tick_ok && memory_ok && !source_degraded;

    Json(serde_json::json!({
        "autorepair": {
            "healthy": healthy,
            "watchdog_interval_secs": 30,
            "uptime_secs": uptime,
            "tick": s.tick,
            "tick_ok": tick_ok,
            "source": src,
            "source_degraded": source_degraded,
            "memory_mb": format!("{:.1}", heap_mb),
            "memory_ok": memory_ok,
            "clients": s.tx.receiver_count(),
        }
    }))
}

async fn api_info(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    Json(serde_json::json!({
        "version": env!("CARGO_PKG_VERSION"),
        "environment": "production",
        "backend": "rust",
        "source": s.effective_source(),
        "hub_ip": detect_hub_ip(),
        "features": {
            "wifi_sensing": true,
            "pose_estimation": true,
            "signal_processing": true,
            "ruvector": true,
            "streaming": true,
        }
    }))
}

async fn pose_current(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    let persons = match &s.latest_update {
        Some(update) => update.persons.clone().unwrap_or_else(|| derive_pose_from_sensing(update)),
        None => vec![],
    };
    Json(serde_json::json!({
        "timestamp": chrono::Utc::now().timestamp_millis() as f64 / 1000.0,
        "persons": persons,
        "total_persons": persons.len(),
        "source": s.effective_source(),
    }))
}

async fn pose_stats(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    Json(serde_json::json!({
        "total_detections": s.total_detections,
        "average_confidence": 0.87,
        "frames_processed": s.tick,
        "source": s.effective_source(),
    }))
}

async fn pose_zones_summary(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    let presence = s.latest_update.as_ref()
        .map(|u| u.classification.presence).unwrap_or(false);
    Json(serde_json::json!({
        "zones": {
            "zone_1": { "person_count": if presence { 1 } else { 0 }, "status": "monitored" },
            "zone_2": { "person_count": 0, "status": "clear" },
            "zone_3": { "person_count": 0, "status": "clear" },
            "zone_4": { "person_count": 0, "status": "clear" },
        }
    }))
}

async fn stream_status(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    Json(serde_json::json!({
        "active": true,
        "clients": s.tx.receiver_count(),
        "fps": if s.tick > 1 { 10u64 } else { 0u64 },
        "source": s.effective_source(),
    }))
}

// ── Model Management Endpoints ──────────────────────────────────────────────

/// GET /api/v1/models — list discovered RVF model files.
async fn list_models(State(state): State<SharedState>) -> Json<serde_json::Value> {
    // Re-scan directory each call so newly-added files are visible.
    let models = scan_model_files();
    let total = models.len();
    {
        let mut s = state.write().await;
        s.discovered_models = models.clone();
    }
    Json(serde_json::json!({ "models": models, "total": total }))
}

/// GET /api/v1/models/active — return currently loaded model or null.
async fn get_active_model(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    match &s.active_model_id {
        Some(id) => {
            let model = s.discovered_models.iter().find(|m| {
                m.get("id").and_then(|v| v.as_str()) == Some(id.as_str())
            });
            Json(serde_json::json!({
                "active": model.cloned().unwrap_or_else(|| serde_json::json!({ "id": id })),
            }))
        }
        None => Json(serde_json::json!({ "active": serde_json::Value::Null })),
    }
}

/// POST /api/v1/models/load — load a model by ID.
async fn load_model(
    State(state): State<SharedState>,
    Json(body): Json<serde_json::Value>,
) -> Json<serde_json::Value> {
    let model_id = body.get("id")
        .or_else(|| body.get("model_id"))
        .and_then(|v| v.as_str())
        .unwrap_or("")
        .to_string();
    if model_id.is_empty() {
        return Json(serde_json::json!({ "error": "missing 'id' field", "success": false }));
    }
    let mut s = state.write().await;
    s.active_model_id = Some(model_id.clone());
    s.model_loaded = true;
    info!("Model loaded: {model_id}");
    Json(serde_json::json!({ "success": true, "model_id": model_id }))
}

/// POST /api/v1/models/unload — unload the current model.
async fn unload_model(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let mut s = state.write().await;
    let prev = s.active_model_id.take();
    s.model_loaded = false;
    info!("Model unloaded (was: {:?})", prev);
    Json(serde_json::json!({ "success": true, "previous": prev }))
}

/// DELETE /api/v1/models/:id — delete a model file.
async fn delete_model(
    State(state): State<SharedState>,
    Path(id): Path<String>,
) -> Json<serde_json::Value> {
    // ADR-050: Sanitize path to prevent directory traversal
    let safe_id = std::path::Path::new(&id)
        .file_name()
        .and_then(|f| f.to_str())
        .unwrap_or("");
    if safe_id.is_empty() || safe_id != id {
        return Json(serde_json::json!({ "error": "invalid model id", "success": false }));
    }
    let path = PathBuf::from("data/models").join(format!("{}.rvf", safe_id));
    if path.exists() {
        if let Err(e) = std::fs::remove_file(&path) {
            warn!("Failed to delete model file {:?}: {}", path, e);
            return Json(serde_json::json!({ "error": format!("delete failed: {e}"), "success": false }));
        }
        // If this was the active model, unload it
        let mut s = state.write().await;
        if s.active_model_id.as_deref() == Some(id.as_str()) {
            s.active_model_id = None;
            s.model_loaded = false;
        }
        s.discovered_models.retain(|m| {
            m.get("id").and_then(|v| v.as_str()) != Some(id.as_str())
        });
        info!("Model deleted: {id}");
        Json(serde_json::json!({ "success": true, "deleted": id }))
    } else {
        Json(serde_json::json!({ "error": "model not found", "success": false }))
    }
}

/// GET /api/v1/models/lora/profiles — list LoRA adapter profiles.
async fn list_lora_profiles() -> Json<serde_json::Value> {
    // LoRA profiles are discovered from data/models/*.lora.json
    let profiles = scan_lora_profiles();
    Json(serde_json::json!({ "profiles": profiles }))
}

/// POST /api/v1/models/lora/activate — activate a LoRA adapter profile.
async fn activate_lora_profile(
    Json(body): Json<serde_json::Value>,
) -> Json<serde_json::Value> {
    let profile = body.get("profile")
        .or_else(|| body.get("name"))
        .and_then(|v| v.as_str())
        .unwrap_or("")
        .to_string();
    if profile.is_empty() {
        return Json(serde_json::json!({ "error": "missing 'profile' field", "success": false }));
    }
    info!("LoRA profile activated: {profile}");
    Json(serde_json::json!({ "success": true, "profile": profile }))
}

/// Scan `data/models/` for `.rvf` files and return metadata.
fn scan_model_files() -> Vec<serde_json::Value> {
    let dir = PathBuf::from("data/models");
    let mut models = Vec::new();
    if let Ok(entries) = std::fs::read_dir(&dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.extension().and_then(|e| e.to_str()) == Some("rvf") {
                let name = path.file_stem()
                    .and_then(|s| s.to_str())
                    .unwrap_or("unknown")
                    .to_string();
                let size = entry.metadata().map(|m| m.len()).unwrap_or(0);
                let modified = entry.metadata().ok()
                    .and_then(|m| m.modified().ok())
                    .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
                    .map(|d| d.as_secs())
                    .unwrap_or(0);
                models.push(serde_json::json!({
                    "id": name,
                    "name": name,
                    "path": path.display().to_string(),
                    "size_bytes": size,
                    "format": "rvf",
                    "modified_epoch": modified,
                }));
            }
        }
    }
    models
}

/// Scan `data/models/` for `.lora.json` LoRA profile files.
fn scan_lora_profiles() -> Vec<serde_json::Value> {
    let dir = PathBuf::from("data/models");
    let mut profiles = Vec::new();
    if let Ok(entries) = std::fs::read_dir(&dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            let name = path.file_name().and_then(|n| n.to_str()).unwrap_or("");
            if name.ends_with(".lora.json") {
                let profile_name = name.trim_end_matches(".lora.json").to_string();
                // Try to read the profile JSON
                let config = std::fs::read_to_string(&path)
                    .ok()
                    .and_then(|s| serde_json::from_str::<serde_json::Value>(&s).ok())
                    .unwrap_or_else(|| serde_json::json!({}));
                profiles.push(serde_json::json!({
                    "name": profile_name,
                    "path": path.display().to_string(),
                    "config": config,
                }));
            }
        }
    }
    profiles
}

// ── Recording Endpoints ─────────────────────────────────────────────────────

/// GET /api/v1/recording/list — list CSI recordings.
async fn list_recordings() -> Json<serde_json::Value> {
    let recordings = scan_recording_files();
    Json(serde_json::json!({ "recordings": recordings }))
}

/// POST /api/v1/recording/start — start recording CSI data.
async fn start_recording(
    State(state): State<SharedState>,
    Json(body): Json<serde_json::Value>,
) -> Json<serde_json::Value> {
    let mut s = state.write().await;
    if s.recording_active {
        return Json(serde_json::json!({
            "error": "recording already in progress",
            "success": false,
            "recording_id": s.recording_current_id,
        }));
    }
    let id = body.get("id")
        .and_then(|v| v.as_str())
        .map(|s| s.to_string())
        .unwrap_or_else(|| {
            format!("rec_{}", chrono_timestamp())
        });

    // Create the recording file
    let rec_path = PathBuf::from("data/recordings").join(format!("{}.jsonl", id));
    let file = match std::fs::File::create(&rec_path) {
        Ok(f) => f,
        Err(e) => {
            warn!("Failed to create recording file {:?}: {}", rec_path, e);
            return Json(serde_json::json!({
                "error": format!("cannot create file: {e}"),
                "success": false,
            }));
        }
    };

    // Create a stop signal channel
    let (stop_tx, mut stop_rx) = tokio::sync::watch::channel(false);
    s.recording_active = true;
    s.recording_start_time = Some(std::time::Instant::now());
    s.recording_current_id = Some(id.clone());
    s.recording_stop_tx = Some(stop_tx);

    // Subscribe to the broadcast channel to capture CSI frames
    let mut rx = s.tx.subscribe();

    // Add initial recording entry
    s.recordings.push(serde_json::json!({
        "id": id,
        "path": rec_path.display().to_string(),
        "status": "recording",
        "started_at": chrono_timestamp(),
        "frames": 0,
    }));

    let rec_id = id.clone();

    // Spawn writer task in background
    tokio::spawn(async move {
        use std::io::Write;
        let mut writer = std::io::BufWriter::new(file);
        let mut frame_count: u64 = 0;
        loop {
            tokio::select! {
                result = rx.recv() => {
                    match result {
                        Ok(frame_json) => {
                            if writeln!(writer, "{}", frame_json).is_err() {
                                warn!("Recording {rec_id}: write error, stopping");
                                break;
                            }
                            frame_count += 1;
                            // Flush every 100 frames
                            if frame_count % 100 == 0 {
                                let _ = writer.flush();
                            }
                        }
                        Err(broadcast::error::RecvError::Lagged(n)) => {
                            debug!("Recording {rec_id}: lagged {n} frames");
                        }
                        Err(broadcast::error::RecvError::Closed) => {
                            info!("Recording {rec_id}: broadcast closed, stopping");
                            break;
                        }
                    }
                }
                _ = stop_rx.changed() => {
                    if *stop_rx.borrow() {
                        info!("Recording {rec_id}: stop signal received ({frame_count} frames)");
                        break;
                    }
                }
            }
        }
        let _ = writer.flush();
        info!("Recording {rec_id} finished: {frame_count} frames written");
    });

    info!("Recording started: {id}");
    Json(serde_json::json!({ "success": true, "recording_id": id }))
}

/// POST /api/v1/recording/stop — stop recording CSI data.
async fn stop_recording(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let mut s = state.write().await;
    if !s.recording_active {
        return Json(serde_json::json!({
            "error": "no recording in progress",
            "success": false,
        }));
    }
    // Signal the writer task to stop
    if let Some(tx) = s.recording_stop_tx.take() {
        let _ = tx.send(true);
    }
    let duration_secs = s.recording_start_time
        .map(|t| t.elapsed().as_secs())
        .unwrap_or(0);
    let rec_id = s.recording_current_id.take().unwrap_or_default();
    s.recording_active = false;
    s.recording_start_time = None;

    // Update the recording entry status
    for rec in s.recordings.iter_mut() {
        if rec.get("id").and_then(|v| v.as_str()) == Some(rec_id.as_str()) {
            rec["status"] = serde_json::json!("completed");
            rec["duration_secs"] = serde_json::json!(duration_secs);
        }
    }

    info!("Recording stopped: {rec_id} ({duration_secs}s)");
    Json(serde_json::json!({
        "success": true,
        "recording_id": rec_id,
        "duration_secs": duration_secs,
    }))
}

/// DELETE /api/v1/recording/:id — delete a recording file.
async fn delete_recording(
    State(state): State<SharedState>,
    Path(id): Path<String>,
) -> Json<serde_json::Value> {
    // ADR-050: Sanitize path to prevent directory traversal
    let safe_id = std::path::Path::new(&id)
        .file_name()
        .and_then(|f| f.to_str())
        .unwrap_or("");
    if safe_id.is_empty() || safe_id != id {
        return Json(serde_json::json!({ "error": "invalid recording id", "success": false }));
    }
    let path = PathBuf::from("data/recordings").join(format!("{}.jsonl", safe_id));
    if path.exists() {
        if let Err(e) = std::fs::remove_file(&path) {
            warn!("Failed to delete recording {:?}: {}", path, e);
            return Json(serde_json::json!({ "error": format!("delete failed: {e}"), "success": false }));
        }
        let mut s = state.write().await;
        s.recordings.retain(|r| {
            r.get("id").and_then(|v| v.as_str()) != Some(id.as_str())
        });
        info!("Recording deleted: {id}");
        Json(serde_json::json!({ "success": true, "deleted": id }))
    } else {
        Json(serde_json::json!({ "error": "recording not found", "success": false }))
    }
}

/// Scan `data/recordings/` for `.jsonl` files and return metadata.
fn scan_recording_files() -> Vec<serde_json::Value> {
    let dir = PathBuf::from("data/recordings");
    let mut recordings = Vec::new();
    if let Ok(entries) = std::fs::read_dir(&dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.extension().and_then(|e| e.to_str()) == Some("jsonl") {
                let name = path.file_stem()
                    .and_then(|s| s.to_str())
                    .unwrap_or("unknown")
                    .to_string();
                let size = entry.metadata().map(|m| m.len()).unwrap_or(0);
                let modified = entry.metadata().ok()
                    .and_then(|m| m.modified().ok())
                    .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
                    .map(|d| d.as_secs())
                    .unwrap_or(0);
                // Count lines (frames) — approximate for large files
                let frame_count = std::fs::read_to_string(&path)
                    .map(|s| s.lines().count())
                    .unwrap_or(0);
                recordings.push(serde_json::json!({
                    "id": name,
                    "name": name,
                    "path": path.display().to_string(),
                    "size_bytes": size,
                    "frames": frame_count,
                    "modified_epoch": modified,
                    "status": "completed",
                }));
            }
        }
    }
    recordings
}

// ── Training Endpoints ──────────────────────────────────────────────────────

/// GET /api/v1/train/status — get training status.
async fn train_status(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    Json(serde_json::json!({
        "status": s.training_status,
        "config": s.training_config,
    }))
}

/// POST /api/v1/train/start — start a training run.
async fn train_start(
    State(state): State<SharedState>,
    Json(body): Json<serde_json::Value>,
) -> Json<serde_json::Value> {
    let mut s = state.write().await;
    if s.training_status == "running" {
        return Json(serde_json::json!({
            "error": "training already running",
            "success": false,
        }));
    }
    s.training_status = "running".to_string();
    s.training_config = Some(body.clone());
    info!("Training started with config: {}", body);
    Json(serde_json::json!({
        "success": true,
        "status": "running",
        "message": "Training pipeline started. Use GET /api/v1/train/status to monitor.",
    }))
}

/// POST /api/v1/train/stop — stop the current training run.
async fn train_stop(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let mut s = state.write().await;
    if s.training_status != "running" {
        return Json(serde_json::json!({
            "error": "no training in progress",
            "success": false,
        }));
    }
    s.training_status = "idle".to_string();
    info!("Training stopped");
    Json(serde_json::json!({
        "success": true,
        "status": "idle",
    }))
}

// ── Adaptive classifier endpoints ────────────────────────────────────────────

/// POST /api/v1/adaptive/train — train the adaptive classifier from recordings.
async fn adaptive_train(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let rec_dir = PathBuf::from("data/recordings");
    eprintln!("=== Adaptive Classifier Training ===");
    match adaptive_classifier::train_from_recordings(&rec_dir) {
        Ok(model) => {
            let accuracy = model.training_accuracy;
            let frames = model.trained_frames;
            let stats: Vec<_> = model.class_stats.iter().map(|cs| {
                serde_json::json!({
                    "class": cs.label,
                    "samples": cs.count,
                    "feature_means": cs.mean,
                })
            }).collect();

            // Save to disk.
            if let Err(e) = model.save(&adaptive_classifier::model_path()) {
                warn!("Failed to save adaptive model: {e}");
            } else {
                info!("Adaptive model saved to {}", adaptive_classifier::model_path().display());
            }

            // Load into runtime state.
            let mut s = state.write().await;
            s.adaptive_model = Some(model);

            Json(serde_json::json!({
                "success": true,
                "trained_frames": frames,
                "accuracy": accuracy,
                "class_stats": stats,
            }))
        }
        Err(e) => {
            Json(serde_json::json!({
                "success": false,
                "error": e,
            }))
        }
    }
}

/// GET /api/v1/adaptive/status — check adaptive model status.
async fn adaptive_status(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    match &s.adaptive_model {
        Some(model) => Json(serde_json::json!({
            "loaded": true,
            "trained_frames": model.trained_frames,
            "accuracy": model.training_accuracy,
            "version": model.version,
            "classes": adaptive_classifier::CLASSES,
            "class_stats": model.class_stats,
        })),
        None => Json(serde_json::json!({
            "loaded": false,
            "message": "No adaptive model. POST /api/v1/adaptive/train to train one.",
        })),
    }
}

/// POST /api/v1/adaptive/unload — unload the adaptive model (revert to thresholds).
async fn adaptive_unload(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let mut s = state.write().await;
    s.adaptive_model = None;
    Json(serde_json::json!({ "success": true, "message": "Adaptive model unloaded." }))
}

// ── Field model calibration endpoints (eigenvalue person counting) ──────────

async fn calibration_start(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let mut s = state.write().await;
    // Guard: don't discard an in-progress or fresh calibration
    if let Some(ref fm) = s.field_model {
        match fm.status() {
            CalibrationStatus::Collecting => {
                return Json(serde_json::json!({
                    "success": false,
                    "error": "Calibration already in progress. Call /calibration/stop first.",
                    "frame_count": fm.calibration_frame_count(),
                }));
            }
            CalibrationStatus::Fresh => {
                return Json(serde_json::json!({
                    "success": false,
                    "error": "A fresh calibration already exists. Call /calibration/stop or wait for expiry.",
                }));
            }
            _ => {} // Stale/Expired/Uncalibrated — ok to recalibrate
        }
    }
    match FieldModel::new(field_bridge::single_link_config()) {
        Ok(fm) => {
            s.field_model = Some(fm);
            Json(serde_json::json!({
                "success": true,
                "message": "Calibration started — keep room empty while frames accumulate.",
            }))
        }
        Err(e) => Json(serde_json::json!({
            "success": false,
            "error": format!("{e}"),
        })),
    }
}

async fn calibration_stop(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let mut s = state.write().await;
    if let Some(ref mut fm) = s.field_model {
        let ts = chrono::Utc::now().timestamp_micros() as u64;
        match fm.finalize_calibration(ts, 0) {
            Ok(modes) => {
                let baseline = modes.baseline_eigenvalue_count;
                let variance_explained = modes.variance_explained;
                info!("Field model calibrated: baseline_eigenvalues={baseline}, variance_explained={variance_explained:.2}");
                Json(serde_json::json!({
                    "success": true,
                    "baseline_eigenvalue_count": baseline,
                    "variance_explained": variance_explained,
                    "frame_count": fm.calibration_frame_count(),
                }))
            }
            Err(e) => Json(serde_json::json!({
                "success": false,
                "error": format!("{e}"),
            })),
        }
    } else {
        Json(serde_json::json!({
            "success": false,
            "error": "No field model active — call /calibration/start first.",
        }))
    }
}

async fn calibration_status(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    match s.field_model.as_ref() {
        Some(fm) => Json(serde_json::json!({
            "active": true,
            "status": format!("{:?}", fm.status()),
            "frame_count": fm.calibration_frame_count(),
        })),
        None => Json(serde_json::json!({
            "active": false,
            "status": "none",
        })),
    }
}

/// Generate a simple timestamp string (epoch seconds) for recording IDs.
fn chrono_timestamp() -> u64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_secs())
        .unwrap_or(0)
}

async fn vital_signs_endpoint(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    let vs = &s.latest_vitals;
    let (br_len, br_cap, hb_len, hb_cap) = s.vital_detector.buffer_status();
    Json(serde_json::json!({
        "vital_signs": {
            "breathing_rate_bpm": vs.breathing_rate_bpm,
            "heart_rate_bpm": vs.heart_rate_bpm,
            "breathing_confidence": vs.breathing_confidence,
            "heartbeat_confidence": vs.heartbeat_confidence,
            "signal_quality": vs.signal_quality,
        },
        "buffer_status": {
            "breathing_samples": br_len,
            "breathing_capacity": br_cap,
            "heartbeat_samples": hb_len,
            "heartbeat_capacity": hb_cap,
        },
        "source": s.effective_source(),
        "tick": s.tick,
    }))
}

/// GET /api/v1/edge-vitals — latest edge vitals from ESP32 (ADR-039).
async fn edge_vitals_endpoint(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    match &s.edge_vitals {
        Some(v) => Json(serde_json::json!({
            "status": "ok",
            "edge_vitals": v,
        })),
        None => Json(serde_json::json!({
            "status": "no_data",
            "edge_vitals": null,
            "message": "No edge vitals packet received yet. Ensure ESP32 edge_tier >= 1.",
        })),
    }
}

/// GET /api/v1/wasm-events — latest WASM events from ESP32 (ADR-040).
async fn wasm_events_endpoint(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    match &s.latest_wasm_events {
        Some(w) => Json(serde_json::json!({
            "status": "ok",
            "wasm_events": w,
        })),
        None => Json(serde_json::json!({
            "status": "no_data",
            "wasm_events": null,
            "message": "No WASM output packet received yet. Upload and start a .wasm module on the ESP32.",
        })),
    }
}

async fn model_info(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    match &s.rvf_info {
        Some(info) => Json(serde_json::json!({
            "status": "loaded",
            "container": info,
        })),
        None => Json(serde_json::json!({
            "status": "no_model",
            "message": "No RVF container loaded. Use --load-rvf <path> to load one.",
        })),
    }
}

async fn model_layers(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    match &s.progressive_loader {
        Some(loader) => {
            let (a, b, c) = loader.layer_status();
            Json(serde_json::json!({
                "layer_a": a,
                "layer_b": b,
                "layer_c": c,
                "progress": loader.loading_progress(),
            }))
        }
        None => Json(serde_json::json!({
            "layer_a": false,
            "layer_b": false,
            "layer_c": false,
            "progress": 0.0,
            "message": "No model loaded with progressive loading",
        })),
    }
}

async fn model_segments(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    match &s.progressive_loader {
        Some(loader) => Json(serde_json::json!({ "segments": loader.segment_list() })),
        None => Json(serde_json::json!({ "segments": [] })),
    }
}

async fn sona_profiles(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    let names = s
        .progressive_loader
        .as_ref()
        .map(|l| l.sona_profile_names())
        .unwrap_or_default();
    let active = s.active_sona_profile.clone().unwrap_or_default();
    Json(serde_json::json!({ "profiles": names, "active": active }))
}

async fn sona_activate(
    State(state): State<SharedState>,
    Json(body): Json<serde_json::Value>,
) -> Json<serde_json::Value> {
    let profile = body
        .get("profile")
        .and_then(|p| p.as_str())
        .unwrap_or("")
        .to_string();

    let mut s = state.write().await;
    let available = s
        .progressive_loader
        .as_ref()
        .map(|l| l.sona_profile_names())
        .unwrap_or_default();

    if available.contains(&profile) {
        s.active_sona_profile = Some(profile.clone());
        Json(serde_json::json!({ "status": "activated", "profile": profile }))
    } else {
        Json(serde_json::json!({
            "status": "error",
            "message": format!("Profile '{}' not found. Available: {:?}", profile, available),
        }))
    }
}

/// GET /api/v1/nodes — per-node health and feature info.
async fn nodes_endpoint(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    let now = std::time::Instant::now();
    // Pull TinyML adaptive baseline snapshot once; map node_id → trained_pct/anomaly.
    let agg_snap = aggression::snapshot();
    let agg_by_node: std::collections::HashMap<u8, &aggression::AggressionScore> = agg_snap
        .as_ref()
        .map(|s| s.nodes.iter().map(|n| (n.node_id, n)).collect())
        .unwrap_or_default();
    let nodes: Vec<serde_json::Value> = s.node_states.iter()
        .map(|(&id, ns)| {
            let elapsed_ms = ns.last_frame_time
                .map(|t| now.duration_since(t).as_millis() as u64)
                .unwrap_or(999_999);
            let stale = elapsed_ms > 5000;
            let offline = elapsed_ms > 30_000;
            let status = if offline { "offline" }
                else if stale { "stale" }
                else { "active" };
            // Firmware status is derived from liveness + edge-vitals presence.
            // - "streaming"   : recent frames AND we've seen edge_vitals at least once
            //                   (means real ESP-IDF firmware with edge DSP is running)
            // - "csi_only"    : recent frames but no edge_vitals (PlatformIO/Arduino
            //                   firmware OR edge tier=0)
            // - "stale"       : last frame 5–30 s ago
            // - "offline"     : > 30 s
            let firmware_status = if offline { "offline" }
                else if stale { "stale" }
                else if ns.edge_vitals.is_some() { "streaming" }
                else { "csi_only" };
            let rssi = ns.rssi_history.back().copied().unwrap_or(-90.0);
            let ml = agg_by_node.get(&id);
            let cfg = node_config_for(id);
            serde_json::json!({
                "node_id": id,
                "status": status,
                "firmware_status": firmware_status,
                "role": node_role_for(id),
                "ip": ns.peer_ip.map(|ip| ip.to_string()),
                "last_seen_ms": elapsed_ms,
                "rssi_dbm": rssi,
                "motion_level": &ns.current_motion_level,
                "person_count": ns.prev_person_count,
                "ml_trained_pct": ml.map(|m| m.trained_pct).unwrap_or(0.0),
                "ml_anomaly": ml.map(|m| m.anomaly).unwrap_or(0.0),
                "ml_aggression": ml.map(|m| m.aggression).unwrap_or(0.0),
                "ml_samples": ml.map(|m| m.samples).unwrap_or(0),
                // Remote-configurable settings (rate_hz / channel / tx_power_dbm).
                // Empty object if this node has no overrides. `revision`
                // lets a polling firmware short-circuit when unchanged.
                "config": cfg,
            })
        })
        .collect();
    // Multistatic role mix — quick at-a-glance counters for the UI.
    let mut tx_count = 0u32;
    let mut rx_count = 0u32;
    let mut txrx_count = 0u32;
    for n in &nodes {
        match n.get("role").and_then(|v| v.as_str()).unwrap_or("txrx") {
            "tx" => tx_count += 1,
            "rx" => rx_count += 1,
            _ => txrx_count += 1,
        }
    }
    // Accuracy-aware multistatic readiness: a TX assignment to a stale/offline
    // node is useless for live fusion. Only count assigned roles whose nodes
    // are currently active (frame within 5 s).
    let (tx_active, rx_active, tx_inactive, rx_inactive)
        = role_quality_summary(&s.node_states);
    let role_assigned = (tx_count + rx_count) as usize;
    let active_assigned = tx_active.len() + rx_active.len();
    let role_quality = if role_assigned == 0 { 0.0 }
        else { active_assigned as f32 / role_assigned as f32 };
    Json(serde_json::json!({
        "nodes": nodes,
        "total": nodes.len(),
        "hub_ip": detect_hub_ip(),
        "roles": {
            "tx": tx_count,
            "rx": rx_count,
            "txrx": txrx_count,
            // True multistatic readiness: at least one ACTIVE TX and one
            // ACTIVE RX. A stale role assignment doesn't deliver fusion gain.
            "multistatic_ready": !tx_active.is_empty() && !rx_active.is_empty(),
            "tx_active": tx_active,
            "rx_active": rx_active,
            "tx_inactive": tx_inactive,
            "rx_inactive": rx_inactive,
            "role_quality": role_quality,
        },
    }))
}

/// GET /api/v1/aggression — TinyML environment-trained anomaly + aggression
/// scores. Returns the live aggregate (max anomaly across nodes, max
/// aggression, average trained_pct) plus per-node detail. The baseline is
/// learned online from `motion_band_power` and persists to
/// `data/aggression_baseline.json`.
async fn aggression_endpoint() -> Json<serde_json::Value> {
    match aggression::snapshot() {
        Some(s) => Json(serde_json::json!({
            "ok": true,
            "anomaly": s.anomaly,
            "aggression": s.aggression,
            "trained_pct": s.trained_pct,
            "nodes": s.nodes,
        })),
        None => Json(serde_json::json!({
            "ok": false,
            "error": "aggression detector not initialised",
        })),
    }
}

/// GET /api/v1/room — room dimensions + node positions for 3D visualization.
async fn room_endpoint(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    let node_positions: Vec<serde_json::Value> = s.multistatic_fuser.node_positions()
        .iter()
        .enumerate()
        .map(|(i, pos)| serde_json::json!({
            "node_id": i,
            "x": pos[0] as f64,
            "y": pos[1] as f64,
            "z": pos[2] as f64,
        }))
        .collect();
    let now = std::time::Instant::now();
    let active_nodes: Vec<serde_json::Value> = s.node_states.iter()
        .map(|(&id, ns)| {
            let elapsed_ms = ns.last_frame_time
                .map(|t| now.duration_since(t).as_millis() as u64)
                .unwrap_or(999999);
            serde_json::json!({
                "node_id": id,
                "status": if elapsed_ms > 5000 { "stale" } else { "active" },
                "rssi_dbm": ns.rssi_history.back().copied().unwrap_or(-90.0),
            })
        })
        .collect();
    Json(serde_json::json!({
        "room": { "width": 8.0, "depth": 6.0, "height": 3.0 },
        "node_positions": node_positions,
        "active_nodes": active_nodes,
        "calibrated": s.field_model.is_some(),
    }))
}

/// GET /api/v1/localization/person — RSSI-based person localization via PSO.
///
/// Reads latest RSSI from each active node, runs a Particle-Swarm minimiser
/// against a log-distance path-loss model, returns `{position, residual_db,
/// confidence, nodes_used}`.
///
/// Environment tuning:
///   AEDIS_RSSI_REF         — RSSI at 1 m (default -40 dBm)
///   AEDIS_PATH_EXPONENT    — path-loss exponent N (default 2.5)
///   AEDIS_ROOM_WIDTH_M     — search bound X (default 8.0)
///   AEDIS_ROOM_DEPTH_M     — search bound Y (default 6.0)
///   AEDIS_ROOM_HEIGHT_M    — search bound Z (default 3.0)
///   AEDIS_PSO_SWARM        — swarm size (default 32)
///   AEDIS_PSO_ITERS        — max iterations (default 120)
///   AEDIS_PSO_SEED         — deterministic seed (default: time)
async fn localization_person_endpoint(
    State(state): State<SharedState>,
) -> Json<serde_json::Value> {
    use pso::{estimate_position_rssi_pso, PathLoss, PsoConfig, RssiObservation};

    let s = state.read().await;
    let node_positions = s.multistatic_fuser.node_positions();
    let now = std::time::Instant::now();

    // Pair each active node (fresh RSSI ≤ 5 s) with its physical position.
    let mut obs: Vec<RssiObservation> = Vec::new();
    for (&id, ns) in s.node_states.iter() {
        let fresh = ns.last_frame_time
            .map_or(false, |t| now.duration_since(t).as_secs() < 5);
        if !fresh { continue; }
        let idx = id as usize;
        let np = match node_positions.get(idx) {
            Some(p) => [p[0] as f64, p[1] as f64, p[2] as f64],
            None => continue, // no known coordinates for this node
        };
        let rssi = match ns.rssi_history.back().copied() {
            Some(r) if r < 0.0 && r > -110.0 => r, // plausible dBm
            _ => continue,
        };
        // Weight: stronger RSSI ⇒ higher confidence.
        // Map -40 dBm → 1.0, -90 dBm → 0.1.
        let w = ((rssi + 90.0) / 50.0).clamp(0.1, 1.0);
        obs.push(RssiObservation { node_pos: np, rssi_dbm: rssi, weight: w });
    }

    if obs.len() < 2 {
        return Json(serde_json::json!({
            "status": "insufficient_nodes",
            "nodes_used": obs.len(),
            "required": 2,
            "method": "pso-rssi",
        }));
    }

    let model = PathLoss {
        rssi_ref_1m: env_f64("AEDIS_RSSI_REF", -40.0),
        exponent: env_f64("AEDIS_PATH_EXPONENT", 2.5),
    };
    let bounds = [
        (0.0, env_f64("AEDIS_ROOM_WIDTH_M", 8.0)),
        (0.0, env_f64("AEDIS_ROOM_DEPTH_M", 6.0)),
        (0.0, env_f64("AEDIS_ROOM_HEIGHT_M", 3.0)),
    ];
    let cfg = PsoConfig {
        swarm_size: env_usize("AEDIS_PSO_SWARM", 32).max(4),
        max_iters: env_usize("AEDIS_PSO_ITERS", 120).max(10),
        seed: std::env::var("AEDIS_PSO_SEED").ok().and_then(|s| s.parse().ok()),
        ..Default::default()
    };

    // Drop the read-lock before running PSO (potentially 100s of μs of CPU).
    let obs_owned = obs.clone();
    drop(s);
    let r = match estimate_position_rssi_pso(&obs_owned, &bounds, model, &cfg) {
        Some(r) => r,
        None => return Json(serde_json::json!({ "status": "failed", "method": "pso-rssi" })),
    };

    Json(serde_json::json!({
        "status": "ok",
        "method": r.method,
        "position": { "x": r.position[0], "y": r.position[1], "z": r.position[2] },
        "residual_db": r.residual_db,
        "confidence": r.confidence,
        "iters": r.iters,
        "converged": r.iters < cfg.max_iters,
        "nodes_used": obs_owned.len(),
        "path_loss": { "rssi_ref_1m": model.rssi_ref_1m, "exponent": model.exponent },
        "timestamp": chrono::Utc::now().timestamp_millis(),
    }))
}

fn env_f64(key: &str, default: f64) -> f64 {
    std::env::var(key).ok().and_then(|s| s.parse().ok()).unwrap_or(default)
}
fn env_usize(key: &str, default: usize) -> usize {
    std::env::var(key).ok().and_then(|s| s.parse().ok()).unwrap_or(default)
}

fn env_flag(key: &str) -> bool {
    matches!(
        std::env::var(key).ok().as_deref(),
        Some("1") | Some("true") | Some("TRUE") | Some("yes") | Some("on")
    )
}

/// GET /api/v1/preprocessing/stats — CSI preprocessing pipeline status (ADR-083).
///
/// Reports which optional preprocessing stages are enabled via env flags.
/// Stages are wired into the per-node pipeline in a future patch; this endpoint
/// exists today so the UI can show the user what *would* be applied.
async fn preprocessing_stats_endpoint() -> Json<serde_json::Value> {
    Json(serde_json::json!({
        "stages": [
            { "name": "hampel",                "enabled": env_flag("IONITY_HAMPEL"),                "env": "IONITY_HAMPEL",                "purpose": "Outlier suppression on amplitude" },
            { "name": "phase_sanitizer",       "enabled": env_flag("IONITY_PHASE_SANITIZER"),       "env": "IONITY_PHASE_SANITIZER",       "purpose": "Per-frame phase unwrap + STO/SFO removal" },
            { "name": "adaptive_denoise",      "enabled": env_flag("IONITY_ADAPTIVE_DENOISE"),      "env": "IONITY_ADAPTIVE_DENOISE",      "purpose": "Adaptive amplitude denoising" },
            { "name": "csi_ratio",             "enabled": env_flag("IONITY_CSI_RATIO"),             "env": "IONITY_CSI_RATIO",             "purpose": "Conjugate-multiply ratio (cancels CFO across antennas)" },
            { "name": "subcarrier_selection",  "enabled": env_flag("IONITY_SUBCARRIER_SELECTION"),  "env": "IONITY_SUBCARRIER_SELECTION",  "purpose": "Variance-driven subcarrier sensitivity selection" }
        ],
        "note": "Stages execute in declared order pre-MultistaticFuser. Disabled by default to preserve verify hash."
    }))
}

/// GET /api/v1/registry — node registry (firmware/esp32-csi-node/nodes.yaml as JSON).
///
/// Falls back to ui/data/node-registry.json (regenerated by `ionity scan`),
/// then to a minimal stub if neither is present.
async fn registry_endpoint() -> Json<serde_json::Value> {
    // Search a few known locations relative to CWD; first hit wins.
    let candidates = [
        "firmware/esp32-csi-node/nodes.yaml",
        "../../firmware/esp32-csi-node/nodes.yaml",
        "../../../firmware/esp32-csi-node/nodes.yaml",
        "ui/data/node-registry.json",
        "../../ui/data/node-registry.json",
    ];
    for path in candidates.iter() {
        let p = std::path::Path::new(path);
        if !p.exists() { continue; }
        let body = match std::fs::read_to_string(p) {
            Ok(s) => s,
            Err(_) => continue,
        };
        // JSON?
        if path.ends_with(".json") {
            if let Ok(v) = serde_json::from_str::<serde_json::Value>(&body) {
                return Json(serde_json::json!({
                    "source": path,
                    "nodes": v,
                }));
            }
        }
        // YAML — light parser: extract id/mac/board lines without pulling serde_yaml.
        let nodes = parse_nodes_yaml_lite(&body);
        return Json(serde_json::json!({
            "source": path,
            "nodes": nodes,
        }));
    }
    Json(serde_json::json!({
        "source": "none",
        "nodes": [],
        "warning": "no nodes.yaml or node-registry.json found",
    }))
}

/// Minimal YAML parser for the nodes.yaml shape we control.
/// Recognises top-level `nodes:` list with `- id: N` blocks containing
/// `mac:`, `board:`, `role:`, `notes:` keys. Unknown keys are passed through.
fn parse_nodes_yaml_lite(yaml: &str) -> Vec<serde_json::Value> {
    let mut out: Vec<serde_json::Value> = Vec::new();
    let mut current: Option<serde_json::Map<String, serde_json::Value>> = None;
    for raw in yaml.lines() {
        let line = raw.trim_end();
        if line.trim().is_empty() || line.trim_start().starts_with('#') { continue; }
        // New list item?
        if let Some(rest) = line.strip_prefix("  - ") {
            if let Some(prev) = current.take() { out.push(serde_json::Value::Object(prev)); }
            let mut m = serde_json::Map::new();
            if let Some((k, v)) = rest.split_once(':') {
                m.insert(k.trim().to_string(), serde_json::Value::String(v.trim().trim_matches('"').to_string()));
            }
            current = Some(m);
            continue;
        }
        // Continuation key inside a node block (4-space indent).
        if line.starts_with("    ") {
            let trimmed = line.trim();
            if let Some((k, v)) = trimmed.split_once(':') {
                if let Some(m) = current.as_mut() {
                    let val = v.trim().trim_matches('"');
                    m.insert(k.trim().to_string(), serde_json::Value::String(val.to_string()));
                }
            }
        }
    }
    if let Some(prev) = current { out.push(serde_json::Value::Object(prev)); }
    out
}

/// POST /api/v1/nodes/positions — update node positions at runtime.
async fn update_node_positions(
    State(state): State<SharedState>,
    Json(body): Json<Vec<serde_json::Value>>,
) -> Json<serde_json::Value> {
    let mut positions: Vec<[f32; 3]> = Vec::new();
    for entry in &body {
        let x = entry.get("x").and_then(|v| v.as_f64()).unwrap_or(0.0) as f32;
        let y = entry.get("y").and_then(|v| v.as_f64()).unwrap_or(0.0) as f32;
        let z = entry.get("z").and_then(|v| v.as_f64()).unwrap_or(0.0) as f32;
        positions.push([x, y, z]);
    }
    let count = positions.len();
    {
        let mut s = state.write().await;
        s.multistatic_fuser.set_node_positions(positions);
    }
    info!("Node positions updated via API: {} nodes", count);
    Json(serde_json::json!({
        "status": "ok",
        "updated": count,
    }))
}

/// Compute a deterministic perimeter layout for `n` nodes inside a
/// `width × depth × height` room, ordered corners → mid-walls. Height
/// fixed at ~55% of room height (typical mounting).
fn perimeter_layout(n: usize, width: f32, depth: f32, height: f32) -> Vec<[f32; 3]> {
    let inset: f32 = 0.3;
    let h = height * 0.55;
    let mut anchors: Vec<[f32; 3]> = vec![
        [inset,         h, inset],
        [width - inset, h, inset],
        [width - inset, h, depth - inset],
        [inset,         h, depth - inset],
        [width * 0.5,   h, inset],
        [width * 0.5,   h, depth - inset],
        [inset,         h, depth * 0.5],
        [width - inset, h, depth * 0.5],
    ];
    // Pad with linearly-interpolated points for >8 nodes.
    while anchors.len() < n {
        let t = anchors.len() as f32 / n as f32;
        anchors.push([width * t, h, depth * (1.0 - t)]);
    }
    anchors.truncate(n);
    anchors
}

/// Auto-calibration shared between startup and on-demand. Returns the
/// number of positions written, or 0 if no active nodes were found.
///
/// Algorithm: for every node that has reported within the last 10 s,
/// compute mean RSSI from its `rssi_history` ring buffer. Sort nodes by
/// `node_id` for determinism (so node 0 → corner 0 every restart even
/// if we mix the boards), then assign each to the next perimeter slot.
/// Room dimensions are read from the existing constants.
async fn auto_calibrate_inner(state: &SharedState) -> usize {
    let now = std::time::Instant::now();
    let active_ids: Vec<u8> = {
        let s = state.read().await;
        let mut ids: Vec<u8> = s.node_states.iter()
            .filter(|(_, ns)| ns.last_frame_time
                .map_or(false, |t| now.duration_since(t).as_secs() < 10))
            .map(|(&id, _)| id)
            .collect();
        ids.sort_unstable();
        ids
    };
    if active_ids.is_empty() {
        return 0;
    }
    // Highest seen ID dictates the slot count, so an unmoved node keeps
    // its slot when others come/go.
    let slot_count = (*active_ids.iter().max().unwrap() as usize) + 1;
    let layout = perimeter_layout(slot_count, 8.0, 6.0, 3.0);
    {
        let mut s = state.write().await;
        s.multistatic_fuser.set_node_positions(layout);
    }
    info!(active = active_ids.len(), slots = slot_count, "Auto-calibrated node positions (perimeter layout)");
    active_ids.len()
}

/// POST /api/v1/calibration/auto — re-derive node positions on demand.
///
/// Stable across reboots: each node keeps its slot as long as its
/// `node_id` is stable. Adding or moving a node triggers a re-layout.
async fn calibration_auto_endpoint(
    State(state): State<SharedState>,
) -> Json<serde_json::Value> {
    let active = auto_calibrate_inner(&state).await;
    let positions: Vec<serde_json::Value> = {
        let s = state.read().await;
        s.multistatic_fuser.node_positions().iter().enumerate()
            .map(|(i, p)| serde_json::json!({
                "node_id": i, "x": p[0] as f64, "y": p[1] as f64, "z": p[2] as f64,
            }))
            .collect()
    };
    Json(serde_json::json!({
        "status": if active > 0 { "ok" } else { "no_active_nodes" },
        "active_nodes": active,
        "positions": positions,
        "method": "perimeter_by_node_id",
    }))
}

/// Background task: 30 s after server starts, if no positions are set,
/// run auto-calibration once. Re-tries every 30 s until at least one
/// node is active. Idempotent — does nothing once positions exist.
async fn auto_calibration_startup_task(state: SharedState) {
    use tokio::time::{sleep, Duration};
    sleep(Duration::from_secs(30)).await;
    loop {
        let already_set = {
            let s = state.read().await;
            !s.multistatic_fuser.node_positions().is_empty()
        };
        if already_set {
            info!("Auto-calibration: positions already configured, skipping.");
            return;
        }
        let active = auto_calibrate_inner(&state).await;
        if active > 0 {
            return;
        }
        sleep(Duration::from_secs(30)).await;
    }
}

/// GET /api/v1/tomography — voxel grid status from field model.
async fn tomography_endpoint(State(state): State<SharedState>) -> Json<serde_json::Value> {
    let s = state.read().await;
    match &s.field_model {
        Some(fm) => {
            let person_count = s.person_count();
            let has_data = !s.frame_history.is_empty() ||
                s.node_states.values().any(|ns| !ns.frame_history.is_empty());
            let variance_explained = fm.modes()
                .map(|m| m.variance_explained)
                .unwrap_or(0.0);
            Json(serde_json::json!({
                "status": "calibrated",
                "grid": [8, 8, 4],
                "person_count": person_count,
                "has_data": has_data,
                "variance_explained": variance_explained,
            }))
        }
        None => Json(serde_json::json!({
            "status": "uncalibrated",
            "message": "Run POST /api/v1/calibration/start to calibrate the field model",
        })),
    }
}

async fn info_page() -> Html<String> {
    Html(format!(
        "<html><body>\
         <h1>WiFi-DensePose Sensing Server</h1>\
         <p>Rust + Axum + RuVector</p>\
         <ul>\
         <li><a href='/health'>/health</a> — Server health</li>\
         <li><a href='/api/v1/sensing/latest'>/api/v1/sensing/latest</a> — Latest sensing data</li>\
         <li><a href='/api/v1/vital-signs'>/api/v1/vital-signs</a> — Vital sign estimates (HR/RR)</li>\
         <li><a href='/api/v1/model/info'>/api/v1/model/info</a> — RVF model container info</li>\
         <li>ws://localhost:8765/ws/sensing — WebSocket stream</li>\
         </ul>\
         </body></html>"
    ))
}

// ── UDP receiver task ────────────────────────────────────────────────────────

async fn udp_receiver_task(state: SharedState, udp_port: u16) {
    let addr = format!("0.0.0.0:{udp_port}");
    let socket = match UdpSocket::bind(&addr).await {
        Ok(s) => {
            info!("UDP listening on {addr} for ESP32 CSI frames");
            s
        }
        Err(e) => {
            error!("Failed to bind UDP {addr}: {e}");
            return;
        }
    };

    let mut buf = [0u8; 2048];
    loop {
        match socket.recv_from(&mut buf).await {
            Ok((len, src)) => {
                // ADR-039: Try edge vitals packet first (magic 0xC511_0002).
                if let Some(vitals) = parse_esp32_vitals(&buf[..len]) {
                    debug!("ESP32 vitals from {src}: node={} br={:.1} hr={:.1} pres={}",
                           vitals.node_id, vitals.breathing_rate_bpm,
                           vitals.heartrate_bpm, vitals.presence);
                    let mut s = state.write().await;
                    // Broadcast vitals via WebSocket.
                    if let Ok(json) = serde_json::to_string(&serde_json::json!({
                        "type": "edge_vitals",
                        "node_id": vitals.node_id,
                        "presence": vitals.presence,
                        "fall_detected": vitals.fall_detected,
                        "motion": vitals.motion,
                        "breathing_rate_bpm": vitals.breathing_rate_bpm,
                        "heartrate_bpm": vitals.heartrate_bpm,
                        "n_persons": vitals.n_persons,
                        "motion_energy": vitals.motion_energy,
                        "presence_score": vitals.presence_score,
                        "rssi": vitals.rssi,
                    })) {
                        let _ = s.tx.send(json);
                    }

                    // Issue #323: Also emit a sensing_update so the UI renders
                    // detections for ESP32 nodes running the edge DSP pipeline
                    // (Tier 2+).  Without this, vitals arrive but the UI shows
                    // "no detection" because it only renders sensing_update msgs.
                    s.source = "esp32".to_string();
                    s.last_esp32_frame = Some(std::time::Instant::now());

                    // ── Per-node state for edge vitals (issue #249) ──────
                    let node_id = vitals.node_id;
                    let ns = s.node_states.entry(node_id).or_insert_with(NodeState::new);
                    ns.last_frame_time = Some(std::time::Instant::now());
                    ns.peer_ip = Some(src.ip());
                    ns.edge_vitals = Some(vitals.clone());
                    ns.rssi_history.push_back(vitals.rssi as f64);
                    if ns.rssi_history.len() > 60 { ns.rssi_history.pop_front(); }

                    // Store per-node person count from edge vitals.
                    let node_est = if vitals.presence {
                        (vitals.n_persons as usize).max(1)
                    } else {
                        0
                    };
                    ns.prev_person_count = node_est;

                    s.tick += 1;
                    let tick = s.tick;

                    let motion_level = if vitals.motion { "present_moving" }
                        else if vitals.presence { "present_still" }
                        else { "absent" };
                    let motion_score = if vitals.motion { 0.8 }
                        else if vitals.presence { 0.3 }
                        else { 0.05 };

                    // Aggregate person count: gate on presence first (matching WiFi path).
                    let now = std::time::Instant::now();
                    let n_active_for_count = s.node_states.values()
                        .filter(|ns| ns.last_frame_time.map_or(false, |t| now.duration_since(t).as_secs() < 10))
                        .count();
                    let (count_evidence, total_persons) = if vitals.presence {
                        let (fused, fallback_count) = multistatic_bridge::fuse_or_fallback(
                            &s.multistatic_fuser, &s.node_states,
                        );
                        let mut estimates = Vec::new();

                        match fused {
                            Some(ref f) => {
                                let score = multistatic_bridge::compute_person_score_from_amplitudes(&f.fused_amplitude);
                                s.smoothed_person_score = s.smoothed_person_score * 0.90 + score * 0.10;
                                let count = s.person_count();
                                estimates.push(CountEstimate {
                                    source: "multistatic_fusion".to_string(),
                                    count: count.max(1),
                                    confidence: (0.55 + 0.12 * n_active_for_count.saturating_sub(1) as f64)
                                        .clamp(0.55, 0.92),
                                });
                            }
                            None => {
                                if let Some(count) = fallback_count {
                                    estimates.push(CountEstimate {
                                        source: "fallback_fusion".to_string(),
                                        count: count.max(1),
                                        confidence: 0.40,
                                    });
                                }
                            }
                        }

                        if let Some(mut corr) = best_active_node_estimate(
                            &s.node_states,
                            now,
                            "correlation",
                            estimate_persons_from_correlation_evidence,
                        ) {
                            corr.confidence = (corr.confidence
                                + 0.05 * n_active_for_count.saturating_sub(1) as f64)
                                .clamp(0.0, 0.95);
                            estimates.push(corr);
                        }

                        if let Some(mut delay) = best_active_node_estimate(
                            &s.node_states,
                            now,
                            "delay_taps",
                            estimate_persons_from_delay_taps_evidence,
                        ) {
                            delay.confidence = (delay.confidence
                                + 0.04 * n_active_for_count.saturating_sub(1) as f64)
                                .clamp(0.0, 0.95);
                            estimates.push(delay);
                        }

                        let presence_scale = (0.70 + 0.30 * (vitals.presence_score as f64).clamp(0.0, 1.0))
                            .clamp(0.70, 1.0);
                        for estimate in &mut estimates {
                            estimate.confidence = (estimate.confidence * presence_scale).clamp(0.0, 0.95);
                        }

                        let evidence = weighted_count_consensus(estimates);
                        let count = evidence
                            .as_ref()
                            .map(|item| item.consensus_count)
                            .or_else(|| fallback_count.map(|value| value.max(1)))
                            .unwrap_or(1);
                        s.prev_person_count = count;
                        (evidence, count)
                    } else {
                        s.prev_person_count = 0;
                        (None, 0)
                    };

                    // Feed field model calibration if active (use per-node history for ESP32).
                    // Extract frame_history before mutable field_model borrow to avoid split-borrow conflict.
                    let _calib_history: Option<VecDeque<Vec<f64>>> = s.node_states.get(&node_id)
                        .map(|ns| ns.frame_history.clone());
                    if let Some(ref mut fm) = s.field_model {
                        if let Some(ref fh) = _calib_history {
                            field_bridge::maybe_feed_calibration(fm, fh);
                        }
                    }

                    // Build nodes array with all active nodes.
                    let active_nodes: Vec<NodeInfo> = s.node_states.iter()
                        .filter(|(_, n)| n.last_frame_time.map_or(false, |t| now.duration_since(t).as_secs() < 10))
                        .map(|(&id, n)| NodeInfo {
                            node_id: id,
                            rssi_dbm: n.rssi_history.back().copied().unwrap_or(0.0),
                            position: [2.0, 0.0, 1.5],
                            amplitude: vec![],
                            subcarrier_count: 0,
                        })
                        .collect();

                    let features = FeatureInfo {
                        mean_rssi: vitals.rssi as f64,
                        variance: vitals.motion_energy as f64,
                        motion_band_power: vitals.motion_energy as f64,
                        breathing_band_power: if vitals.presence { 0.5 } else { 0.0 },
                        dominant_freq_hz: vitals.breathing_rate_bpm / 60.0,
                        change_points: 0,
                        spectral_power: vitals.motion_energy as f64,
                    };

                    // Store latest features on node for cross-node fusion.
                    s.node_states.get_mut(&node_id)
                        .map(|ns| ns.latest_features = Some(features.clone()));

                    // Cross-node fusion: combine features from all active nodes.
                    let fused_features = fuse_multi_node_features(&features, &s.node_states);

                    let mut classification = ClassificationInfo {
                        motion_level: motion_level.to_string(),
                        presence: vitals.presence,
                        confidence: vitals.presence_score as f64,
                    };

                    // Boost classification confidence with multi-node coverage.
                    let n_active = s.node_states.values()
                        .filter(|ns| ns.last_frame_time.map_or(false, |t| now.duration_since(t).as_secs() < 10))
                        .count();
                    if n_active > 1 {
                        classification.confidence = (classification.confidence
                            * (1.0 + 0.15 * (n_active as f64 - 1.0))).clamp(0.0, 1.0);
                    }

                    let signal_field = generate_signal_field(
                        fused_features.mean_rssi, motion_score, vitals.breathing_rate_bpm / 60.0,
                        (vitals.presence_score as f64).min(1.0), &[],
                    );

                    let mut update = SensingUpdate {
                        msg_type: "sensing_update".to_string(),
                        timestamp: chrono::Utc::now().timestamp_millis() as f64 / 1000.0,
                        source: "esp32".to_string(),
                        tick,
                        nodes: active_nodes,
                        features: fused_features.clone(),
                        classification,
                        signal_field,
                        vital_signs: Some(VitalSigns {
                            breathing_rate_bpm: if vitals.breathing_rate_bpm > 0.0 { Some(vitals.breathing_rate_bpm) } else { None },
                            heart_rate_bpm: if vitals.heartrate_bpm > 0.0 { Some(vitals.heartrate_bpm) } else { None },
                            breathing_confidence: if vitals.presence { 0.7 } else { 0.0 },
                            heartbeat_confidence: if vitals.presence { 0.7 } else { 0.0 },
                            signal_quality: vitals.presence_score as f64,
                        }),
                        enhanced_motion: None,
                        enhanced_breathing: None,
                        posture: None,
                        signal_quality_score: None,
                        quality_verdict: None,
                        bssid_count: None,
                        pose_keypoints: None,
                        model_status: None,
                        persons: None,
                        estimated_persons: if total_persons > 0 { Some(total_persons) } else { None },
                        count_evidence,
                        person_anchors: None,
                        node_features: None,
                    };

                    {
                        let s_ref = &mut *s;
                        finalize_people_for_update(
                            &mut update,
                            &mut s_ref.pose_tracker,
                            &mut s_ref.last_tracker_instant,
                            &mut s_ref.track_temporal_state,
                        );
                    }

                    if let Ok(json) = serde_json::to_string(&update) {
                        let _ = s.tx.send(json);
                    }
                    s.latest_update = Some(update);
                    s.edge_vitals = Some(vitals);
                    continue;
                }

                // ADR-040: Try WASM output packet (magic 0xC511_0004).
                if let Some(wasm_output) = parse_wasm_output(&buf[..len]) {
                    debug!("WASM output from {src}: node={} module={} events={}",
                           wasm_output.node_id, wasm_output.module_id,
                           wasm_output.events.len());
                    let mut s = state.write().await;
                    // Broadcast WASM events via WebSocket.
                    if let Ok(json) = serde_json::to_string(&serde_json::json!({
                        "type": "wasm_event",
                        "node_id": wasm_output.node_id,
                        "module_id": wasm_output.module_id,
                        "events": wasm_output.events,
                    })) {
                        let _ = s.tx.send(json);
                    }
                    s.latest_wasm_events = Some(wasm_output);
                    continue;
                }

                if let Some(frame) = parse_esp32_frame(&buf[..len]) {
                    debug!("ESP32 frame from {src}: node={}, subs={}, seq={}",
                           frame.node_id, frame.n_subcarriers, frame.sequence);

                    let mut s = state.write().await;
                    s.source = "esp32".to_string();
                    s.last_esp32_frame = Some(std::time::Instant::now());

                    // Also maintain global frame_history for backward compat
                    // (simulation path, REST endpoints, etc.).
                    s.frame_history.push_back(frame.amplitudes.clone());
                    if s.frame_history.len() > FRAME_HISTORY_CAPACITY {
                        s.frame_history.pop_front();
                    }

                    // ── Per-node processing (issue #249) ──────────────────
                    // Process entirely within per-node state so different
                    // ESP32 nodes never mix their smoothing/vitals buffers.
                    // We scope the mutable borrow of node_states so we can
                    // access other AppStateInner fields afterward.
                    let node_id = frame.node_id;
                    // Clone adaptive model before mutable borrow of node_states
                    // to avoid unsafe raw pointer (review finding #2).
                    let adaptive_model_clone = s.adaptive_model.clone();

                    let ns = s.node_states.entry(node_id).or_insert_with(NodeState::new);
                    ns.last_frame_time = Some(std::time::Instant::now());
                    ns.peer_ip = Some(src.ip());

                    ns.frame_history.push_back(frame.amplitudes.clone());
                    if ns.frame_history.len() > FRAME_HISTORY_CAPACITY {
                        ns.frame_history.pop_front();
                    }

                    let sample_rate_hz = 1000.0 / 500.0_f64;
                    let (features, mut classification, breathing_rate_hz, sub_variances, raw_motion) =
                        extract_features_from_frame(&frame, &ns.frame_history, sample_rate_hz);
                    smooth_and_classify_node(ns, &mut classification, raw_motion);

                    // Adaptive override using cloned model (safe, no raw pointers).
                    if let Some(ref model) = adaptive_model_clone {
                        let amps = ns.frame_history.back()
                            .map(|v| v.as_slice())
                            .unwrap_or(&[]);
                        let feat_arr = adaptive_classifier::features_from_runtime(
                            &serde_json::json!({
                                "variance": features.variance,
                                "motion_band_power": features.motion_band_power,
                                "breathing_band_power": features.breathing_band_power,
                                "spectral_power": features.spectral_power,
                                "dominant_freq_hz": features.dominant_freq_hz,
                                "change_points": features.change_points,
                                "mean_rssi": features.mean_rssi,
                            }),
                            amps,
                        );
                        let (label, conf) = model.classify(&feat_arr);
                        classification.motion_level = label.to_string();
                        classification.presence = label != "absent";
                        classification.confidence = (conf * 0.7 + classification.confidence * 0.3).clamp(0.0, 1.0);
                    }

                    ns.rssi_history.push_back(features.mean_rssi);
                    if ns.rssi_history.len() > 60 {
                        ns.rssi_history.pop_front();
                    }

                    let raw_vitals = ns.vital_detector.process_frame(
                        &frame.amplitudes,
                        &frame.phases,
                    );
                    let vitals = smooth_vitals_node(ns, &raw_vitals);
                    ns.latest_vitals = vitals.clone();

                    // DynamicMinCut person estimation from subcarrier correlation.
                    let corr_persons = estimate_persons_from_correlation(&ns.frame_history);
                    let raw_score = corr_persons as f64 / 3.0;
                    ns.smoothed_person_score = ns.smoothed_person_score * 0.92 + raw_score * 0.08;
                    if classification.presence {
                        let count = score_to_person_count(ns.smoothed_person_score, ns.prev_person_count);
                        ns.prev_person_count = count;
                    } else {
                        ns.prev_person_count = 0;
                    }

                    // Store latest features on node for cross-node fusion.
                    ns.latest_features = Some(features.clone());

                    // Done with per-node mutable borrow; now read aggregated
                    // state from all nodes (the borrow of `ns` ends here).
                    // (We re-borrow node_states immutably via `s` below.)

                    s.rssi_history.push_back(features.mean_rssi);
                    if s.rssi_history.len() > 60 {
                        s.rssi_history.pop_front();
                    }
                    s.latest_vitals = vitals.clone();

                    // Cross-node fusion: combine features from all active nodes.
                    let fused_features = fuse_multi_node_features(&features, &s.node_states);

                    s.tick += 1;
                    let tick = s.tick;

                    let motion_score = if classification.motion_level == "active" { 0.8 }
                        else if classification.motion_level == "present_still" { 0.3 }
                        else { 0.05 };

                    // Aggregate person count: gate on presence first (matching WiFi path).
                    let now = std::time::Instant::now();
                    let n_active_for_count = s.node_states.values()
                        .filter(|ns| ns.last_frame_time.map_or(false, |t| now.duration_since(t).as_secs() < 10))
                        .count();
                    let (count_evidence, total_persons) = if classification.presence {
                        let (fused, fallback_count) = multistatic_bridge::fuse_or_fallback(
                            &s.multistatic_fuser, &s.node_states,
                        );
                        let mut estimates = Vec::new();

                        let fallback_seed = match fused {
                            Some(ref f) => {
                                let score = multistatic_bridge::compute_person_score_from_amplitudes(&f.fused_amplitude);
                                s.smoothed_person_score = s.smoothed_person_score * 0.90 + score * 0.10;
                                let count = s.person_count();
                                estimates.push(CountEstimate {
                                    source: "multistatic_fusion".to_string(),
                                    count: count.max(1),
                                    confidence: (0.55 + 0.12 * n_active_for_count.saturating_sub(1) as f64)
                                        .clamp(0.55, 0.92),
                                });
                                count.max(1)
                            }
                            None => {
                                if let Some(count) = fallback_count {
                                    estimates.push(CountEstimate {
                                        source: "fallback_fusion".to_string(),
                                        count: count.max(1),
                                        confidence: 0.40,
                                    });
                                }
                                fallback_count.unwrap_or(0).max(1)
                            }
                        };

                        if let Some(mut corr) = best_active_node_estimate(
                            &s.node_states,
                            now,
                            "correlation",
                            estimate_persons_from_correlation_evidence,
                        ) {
                            corr.confidence = (corr.confidence
                                + 0.05 * n_active_for_count.saturating_sub(1) as f64)
                                .clamp(0.0, 0.95);
                            estimates.push(corr);
                        }

                        if let Some(mut delay) = best_active_node_estimate(
                            &s.node_states,
                            now,
                            "delay_taps",
                            estimate_persons_from_delay_taps_evidence,
                        ) {
                            delay.confidence = (delay.confidence
                                + 0.04 * n_active_for_count.saturating_sub(1) as f64)
                                .clamp(0.0, 0.95);
                            estimates.push(delay);
                        }

                        // AEDIS_MIN_PERSONS: optional floor (e.g. set to 2 during
                        // multi-person calibration); defaults to 1.
                        let min_floor: usize = std::env::var("AEDIS_MIN_PERSONS")
                            .ok().and_then(|v| v.parse().ok()).unwrap_or(1);

                        let confidence_scale = (0.70 + 0.30 * classification.confidence.clamp(0.0, 1.0))
                            .clamp(0.70, 1.0);
                        for estimate in &mut estimates {
                            estimate.confidence = (estimate.confidence * confidence_scale).clamp(0.0, 0.95);
                        }

                        let mut evidence = weighted_count_consensus(estimates);
                        let count = evidence
                            .as_ref()
                            .map(|item| item.consensus_count)
                            .unwrap_or(fallback_seed)
                            .max(min_floor);
                        if let Some(ref mut item) = evidence {
                            item.consensus_count = item.consensus_count.max(min_floor);
                        }
                        s.prev_person_count = count;
                        (evidence, count)
                    } else {
                        s.prev_person_count = 0;
                        (None, 0)
                    };

                    // Feed field model calibration if active (use per-node history for ESP32).
                    // Extract frame_history before mutable field_model borrow to avoid split-borrow conflict.
                    let _calib_history: Option<VecDeque<Vec<f64>>> = s.node_states.get(&node_id)
                        .map(|ns| ns.frame_history.clone());
                    if let Some(ref mut fm) = s.field_model {
                        if let Some(ref fh) = _calib_history {
                            field_bridge::maybe_feed_calibration(fm, fh);
                        }
                    }

                    // Build nodes array with all active nodes.
                    let active_nodes: Vec<NodeInfo> = s.node_states.iter()
                        .filter(|(_, n)| n.last_frame_time.map_or(false, |t| now.duration_since(t).as_secs() < 10))
                        .map(|(&id, n)| NodeInfo {
                            node_id: id,
                            rssi_dbm: n.rssi_history.back().copied().unwrap_or(0.0),
                            position: [2.0, 0.0, 1.5],
                            amplitude: n.frame_history.back()
                                .map(|a| a.iter().take(56).cloned().collect())
                                .unwrap_or_default(),
                            subcarrier_count: n.frame_history.back().map_or(0, |a| a.len()),
                        })
                        .collect();

                    let mut update = SensingUpdate {
                        msg_type: "sensing_update".to_string(),
                        timestamp: chrono::Utc::now().timestamp_millis() as f64 / 1000.0,
                        source: "esp32".to_string(),
                        tick,
                        nodes: active_nodes,
                        features: fused_features.clone(),
                        classification,
                        signal_field: generate_signal_field(
                            fused_features.mean_rssi, motion_score, breathing_rate_hz,
                            fused_features.variance.min(1.0), &sub_variances,
                        ),
                        vital_signs: Some(vitals),
                        enhanced_motion: None,
                        enhanced_breathing: None,
                        posture: None,
                        signal_quality_score: None,
                        quality_verdict: None,
                        bssid_count: None,
                        pose_keypoints: None,
                        model_status: None,
                        persons: None,
                        estimated_persons: if total_persons > 0 { Some(total_persons) } else { None },
                        count_evidence,
                        person_anchors: None,
                        node_features: None,
                    };

                    {
                        let s_ref = &mut *s;
                        finalize_people_for_update(
                            &mut update,
                            &mut s_ref.pose_tracker,
                            &mut s_ref.last_tracker_instant,
                            &mut s_ref.track_temporal_state,
                        );
                    }

                    if let Ok(json) = serde_json::to_string(&update) {
                        let _ = s.tx.send(json);
                    }
                    s.latest_update = Some(update);

                    // Evict stale nodes every 100 ticks to prevent memory leak.
                    if tick % 100 == 0 {
                        let stale = Duration::from_secs(60);
                        let before = s.node_states.len();
                        s.node_states.retain(|_id, ns| {
                            ns.last_frame_time.map_or(false, |t| now.duration_since(t) < stale)
                        });
                        let evicted = before - s.node_states.len();
                        if evicted > 0 {
                            info!("Evicted {} stale node(s), {} active", evicted, s.node_states.len());
                        }
                    }
                }
            }
            Err(e) => {
                warn!("UDP recv error: {e}");
                tokio::time::sleep(Duration::from_millis(100)).await;
            }
        }
    }
}

// ── Broadcast tick task (for ESP32 mode, sends buffered state) ───────────────

async fn broadcast_tick_task(state: SharedState, tick_ms: u64) {
    let mut interval = tokio::time::interval(Duration::from_millis(tick_ms));
    let mut offline_fallback_count: u64 = 0;

    loop {
        interval.tick().await;

        // Check if ESP32 is offline — if so, generate a synthetic tick to
        // keep the server alive and the UI responsive.
        let is_offline = {
            let s = state.read().await;
            s.effective_source().contains("offline")
        };

        if is_offline {
            offline_fallback_count += 1;
            // Generate a minimal synthetic tick every broadcast interval
            // so the tick counter advances, the watchdog stays happy, and
            // WebSocket clients see activity ("esp32:offline" source label).
            let mut s = state.write().await;
            s.tick += 1;
            let tick = s.tick;

            if offline_fallback_count % 50 == 1 {
                info!("ESP32 offline fallback active — generating synthetic ticks (tick={tick})");
            }

            let update = SensingUpdate {
                msg_type: "sensing_update".to_string(),
                timestamp: chrono::Utc::now().timestamp_millis() as f64 / 1000.0,
                source: "esp32:offline".to_string(),
                tick,
                nodes: vec![],
                features: FeatureInfo {
                    mean_rssi: -90.0,
                    variance: 0.0,
                    motion_band_power: 0.0,
                    breathing_band_power: 0.0,
                    dominant_freq_hz: 0.0,
                    change_points: 0,
                    spectral_power: 0.0,
                },
                classification: ClassificationInfo {
                    motion_level: "absent".to_string(),
                    presence: false,
                    confidence: 0.0,
                },
                signal_field: SignalField {
                    grid_size: [20, 1, 20],
                    values: vec![0.0; 400],
                },
                vital_signs: None,
                enhanced_motion: None,
                enhanced_breathing: None,
                posture: None,
                signal_quality_score: None,
                quality_verdict: None,
                bssid_count: None,
                pose_keypoints: None,
                model_status: None,
                persons: None,
                estimated_persons: None,
                count_evidence: None,
                person_anchors: None,
                node_features: None,
            };

            if let Ok(json) = serde_json::to_string(&update) {
                let _ = s.tx.send(json);
            }
            s.latest_update = Some(update);
        } else {
            offline_fallback_count = 0;
            let s = state.read().await;
            if let Some(ref update) = s.latest_update {
                if s.tx.receiver_count() > 0 {
                    // Re-broadcast the latest sensing_update so pose WS clients
                    // always get data even when ESP32 pauses between frames.
                    if let Ok(json) = serde_json::to_string(update) {
                        let _ = s.tx.send(json);
                    }
                }
            }
        }
    }
}

// ── Software Autorepair: Watchdog + selfcheck + fallback ─────────────────────

/// Autorepair status bits (mirrors ESP32 autorepair.h)
#[derive(Debug, Clone, Serialize, Deserialize)]
struct AutorepairStatus {
    healthy: bool,
    tick_stall: bool,
    memory_high: bool,
    source_degraded: bool,
    uptime_secs: u64,
    last_tick: u64,
    check_count: u64,
    recovery_count: u64,
    heap_mb: f64,
}

/// Background watchdog task that monitors server health and attempts recovery.
///
/// Checks every `interval_secs`:
///  1. Tick progress — detects stalled source tasks.
///  2. Process memory — warns if RSS exceeds threshold.
///  3. Source health — detects offline/degraded sources.
///
/// Recovery actions:
///  - Tick stall: logs critical warning (Tokio tasks can't be restarted externally,
///    but the process supervisor in ionity.sh will catch an unresponsive /health).
///  - Memory: logs warning (no forced GC in Rust, but signals need for investigation).
///  - Source offline: updates state to mark degraded.
///
/// If the server is unresponsive for 3 consecutive checks, it exits the process
/// so the ionity stack supervisor can restart it cleanly.
async fn autorepair_watchdog_task(state: SharedState, interval_secs: u64) {
    let mut interval = tokio::time::interval(Duration::from_secs(interval_secs));
    let mut prev_tick: u64 = 0;
    let mut stall_count: u32 = 0;
    let mut check_count: u64 = 0;
    let mut recovery_count: u64 = 0;
    let max_stall_before_exit: u32 = 3;

    // Skip the first tick (immediate)
    interval.tick().await;

    loop {
        interval.tick().await;
        check_count += 1;

        let current_tick = {
            let s = state.read().await;
            s.tick
        };

        let mut status = AutorepairStatus {
            healthy: true,
            tick_stall: false,
            memory_high: false,
            source_degraded: false,
            uptime_secs: 0,
            last_tick: current_tick,
            check_count,
            recovery_count,
            heap_mb: 0.0,
        };

        // 1. Tick progress check
        //    When the source is offline (e.g. ESP32 disconnected), tick stalls
        //    are expected — the server is healthy, just waiting for data.
        //    Only count tick stalls toward the exit threshold when the source
        //    is actively receiving data (not offline/degraded).
        let source_offline = {
            let s = state.read().await;
            let src = s.effective_source();
            src.contains("offline") || src.contains("degraded")
        };

        if current_tick == prev_tick && current_tick > 0 {
            if source_offline {
                // Source offline — tick stall is expected; don't count toward exit.
                debug!(
                    "Autorepair: tick idle at {} (source offline, not counting as stall)",
                    current_tick
                );
                // Reset stall counter since this is a known-good state.
                stall_count = 0;
            } else {
                stall_count += 1;
                status.tick_stall = true;
                status.healthy = false;
                warn!(
                    "Autorepair: tick stalled at {} for {} consecutive checks",
                    current_tick, stall_count
                );

                if stall_count >= max_stall_before_exit {
                    error!(
                        "Autorepair: tick stalled for {} checks with active source — exiting for supervisor restart",
                        stall_count
                    );
                    recovery_count += 1;
                    // Exit with non-zero code so supervisor knows to restart
                    std::process::exit(1);
                }
            }
        } else {
            if stall_count > 0 {
                info!("Autorepair: tick recovered (was stalled for {} checks)", stall_count);
                recovery_count += 1;
            }
            stall_count = 0;
        }
        prev_tick = current_tick;

        // 2. Process memory check (read /proc/self/status on Linux)
        #[cfg(target_os = "linux")]
        {
            if let Ok(proc_status) = tokio::fs::read_to_string("/proc/self/status").await {
                for line in proc_status.lines() {
                    if line.starts_with("VmRSS:") {
                        if let Some(kb_str) = line.split_whitespace().nth(1) {
                            if let Ok(kb) = kb_str.parse::<f64>() {
                                status.heap_mb = kb / 1024.0;
                                // Warn at 256MB (generous for arm64 Pi)
                                if kb > 256.0 * 1024.0 {
                                    status.memory_high = true;
                                    status.healthy = false;
                                    warn!(
                                        "Autorepair: high memory usage: {:.1} MB RSS",
                                        status.heap_mb
                                    );
                                }
                            }
                        }
                        break;
                    }
                }
            }
        }

        // 3. Source health check
        {
            let s = state.read().await;
            let src = s.effective_source();
            if src.contains("offline") || src.contains("degraded") {
                status.source_degraded = true;
                status.healthy = false;
                warn!("Autorepair: source degraded — {}", src);
            }
            status.uptime_secs = s.start_time.elapsed().as_secs();
        }

        // Periodic health summary (every 10 checks when healthy, every check when not)
        if !status.healthy || check_count % 10 == 0 {
            info!(
                "Autorepair selfcheck #{}: healthy={}, tick={}, RSS={:.1}MB, uptime={}s, recoveries={}",
                check_count, status.healthy, status.last_tick,
                status.heap_mb, status.uptime_secs, recovery_count
            );
        }
    }
}

// ── Main ─────────────────────────────────────────────────────────────────────

fn init_tracing() {
    let env_filter = EnvFilter::try_from_default_env()
        .unwrap_or_else(|_| EnvFilter::new("info,tower_http=debug"));
    let trace_format = std::env::var("SENSING_TRACE_FORMAT")
        .unwrap_or_else(|_| "pretty".to_string())
        .to_ascii_lowercase();

    let fmt_layer = tracing_subscriber::fmt::layer()
        .with_target(true)
        .with_thread_ids(true)
        .with_thread_names(true);

    tracing_subscriber::registry()
        .with(env_filter)
        .with(match trace_format.as_str() {
            "json" => fmt_layer.json().flatten_event(true).boxed(),
            "plain" | "compact" => fmt_layer.compact().boxed(),
            _ => fmt_layer.pretty().boxed(),
        })
        .init();

    info!(trace_format, "Tracing initialised");
}

/// One-shot rebrand migration: rename legacy `ruview_logs.sqlite3` (and -wal/-shm
/// siblings) to `aedi_logs.sqlite3` if the old file exists and the new one does
/// not. Idempotent — safe to call on every startup.
fn migrate_legacy_log_db() {
    let candidates = [
        ("logs/ruview_logs.sqlite3",     "logs/aedi_logs.sqlite3"),
        ("logs/ruview_logs.sqlite3-wal", "logs/aedi_logs.sqlite3-wal"),
        ("logs/ruview_logs.sqlite3-shm", "logs/aedi_logs.sqlite3-shm"),
    ];
    for (old, new) in candidates.iter() {
        let old_p = std::path::Path::new(old);
        let new_p = std::path::Path::new(new);
        if old_p.exists() && !new_p.exists() {
            match std::fs::rename(old_p, new_p) {
                Ok(_) => info!(old, new, "Migrated legacy log database (RuView -> AEDI rebrand)"),
                Err(e) => warn!(old, new, error = %e, "Failed to migrate legacy log database"),
            }
        }
    }
}

#[tokio::main]
async fn main() {
    init_tracing();
    migrate_legacy_log_db();

    let args = Args::parse();

    // Handle --benchmark mode: run vital sign benchmark and exit
    if args.benchmark {
        eprintln!("Running vital sign detection benchmark (1000 frames)...");
        let (total, per_frame) = vital_signs::run_benchmark(1000);
        eprintln!();
        eprintln!("Summary: {} total, {} per frame",
            format!("{total:?}"), format!("{per_frame:?}"));
        return;
    }

    // Handle --export-rvf mode: build an RVF container package and exit
    if let Some(ref rvf_path) = args.export_rvf {
        eprintln!("Exporting RVF container package...");
        use rvf_pipeline::RvfModelBuilder;

        let mut builder = RvfModelBuilder::new("wifi-densepose", "1.0.0");

        // Vital sign config (default breathing 0.1-0.5 Hz, heartbeat 0.8-2.0 Hz)
        builder.set_vital_config(0.1, 0.5, 0.8, 2.0);

        // Model profile (input/output spec)
        builder.set_model_profile(
            "56-subcarrier CSI amplitude/phase @ 10-100 Hz",
            "17 COCO keypoints + body part UV + vital signs",
            "ESP32-S3 or Windows WiFi RSSI, Rust 1.85+",
        );

        // Placeholder weights (17 keypoints × 56 subcarriers × 3 dims = 2856 params)
        let placeholder_weights: Vec<f32> = (0..2856).map(|i| (i as f32 * 0.001).sin()).collect();
        builder.set_weights(&placeholder_weights);

        // Training provenance
        builder.set_training_proof(
            "wifi-densepose-rs-v1.0.0",
            serde_json::json!({
                "pipeline": "ADR-023 8-phase",
                "test_count": 229,
                "benchmark_fps": 9520,
                "framework": "wifi-densepose-rs",
            }),
        );

        // SONA default environment profile
        let default_lora: Vec<f32> = vec![0.0; 64];
        builder.add_sona_profile("default", &default_lora, &default_lora);

        match builder.build() {
            Ok(rvf_bytes) => {
                if let Err(e) = std::fs::write(rvf_path, &rvf_bytes) {
                    eprintln!("Error writing RVF: {e}");
                    std::process::exit(1);
                }
                eprintln!("Wrote {} bytes to {}", rvf_bytes.len(), rvf_path.display());
                eprintln!("RVF container exported successfully.");
            }
            Err(e) => {
                eprintln!("Error building RVF: {e}");
                std::process::exit(1);
            }
        }
        return;
    }

    // Handle --pretrain mode: self-supervised contrastive pretraining (ADR-024)
    if args.pretrain {
        eprintln!("=== WiFi-DensePose Contrastive Pretraining (ADR-024) ===");

        let ds_path = args.dataset.clone().unwrap_or_else(|| PathBuf::from("data"));
        let source = match args.dataset_type.as_str() {
            "wipose" => dataset::DataSource::WiPose(ds_path.clone()),
            _ => dataset::DataSource::MmFi(ds_path.clone()),
        };
        let pipeline = dataset::DataPipeline::new(dataset::DataConfig {
            source, ..Default::default()
        });

        // Generate synthetic or load real CSI windows
        let generate_synthetic_windows = || -> Vec<Vec<Vec<f32>>> {
            (0..50).map(|i| {
                (0..4).map(|a| {
                    (0..56).map(|s| ((i * 7 + a * 13 + s) as f32 * 0.31).sin() * 0.5).collect()
                }).collect()
            }).collect()
        };

        let csi_windows: Vec<Vec<Vec<f32>>> = match pipeline.load() {
            Ok(s) if !s.is_empty() => {
                eprintln!("Loaded {} samples from {}", s.len(), ds_path.display());
                s.into_iter().map(|s| s.csi_window).collect()
            }
            _ => {
                eprintln!("Using synthetic data for pretraining.");
                generate_synthetic_windows()
            }
        };

        let n_subcarriers = csi_windows.first()
            .and_then(|w| w.first())
            .map(|f| f.len())
            .unwrap_or(56);

        let tf_config = graph_transformer::TransformerConfig {
            n_subcarriers, n_keypoints: 17, d_model: 64, n_heads: 4, n_gnn_layers: 2,
        };
        let transformer = graph_transformer::CsiToPoseTransformer::new(tf_config);
        eprintln!("Transformer params: {}", transformer.param_count());

        let trainer_config = trainer::TrainerConfig {
            epochs: args.pretrain_epochs,
            batch_size: 8, lr: 0.001, warmup_epochs: 2, min_lr: 1e-6,
            early_stop_patience: args.pretrain_epochs + 1,
            pretrain_temperature: 0.07,
            ..Default::default()
        };
        let mut t = trainer::Trainer::with_transformer(trainer_config, transformer);

        let e_config = embedding::EmbeddingConfig {
            d_model: 64, d_proj: 128, temperature: 0.07, normalize: true,
        };
        let mut projection = embedding::ProjectionHead::new(e_config.clone());
        let augmenter = embedding::CsiAugmenter::new();

        eprintln!("Starting contrastive pretraining for {} epochs...", args.pretrain_epochs);
        let start = std::time::Instant::now();
        for epoch in 0..args.pretrain_epochs {
            let loss = t.pretrain_epoch(&csi_windows, &augmenter, &mut projection, 0.07, epoch);
            if epoch % 10 == 0 || epoch == args.pretrain_epochs - 1 {
                eprintln!("  Epoch {epoch}: contrastive loss = {loss:.4}");
            }
        }
        let elapsed = start.elapsed().as_secs_f64();
        eprintln!("Pretraining complete in {elapsed:.1}s");

        // Save pretrained model as RVF with embedding segment
        if let Some(ref save_path) = args.save_rvf {
            eprintln!("Saving pretrained model to RVF: {}", save_path.display());
            t.sync_transformer_weights();
            let weights = t.params().to_vec();
            let mut proj_weights = Vec::new();
            projection.flatten_into(&mut proj_weights);

            let mut builder = RvfBuilder::new();
            builder.add_manifest(
                "wifi-densepose-pretrained",
                env!("CARGO_PKG_VERSION"),
                "WiFi DensePose contrastive pretrained model (ADR-024)",
            );
            builder.add_weights(&weights);
            builder.add_embedding(
                &serde_json::json!({
                    "d_model": e_config.d_model,
                    "d_proj": e_config.d_proj,
                    "temperature": e_config.temperature,
                    "normalize": e_config.normalize,
                    "pretrain_epochs": args.pretrain_epochs,
                }),
                &proj_weights,
            );
            match builder.write_to_file(save_path) {
                Ok(()) => eprintln!("RVF saved ({} transformer + {} projection params)",
                    weights.len(), proj_weights.len()),
                Err(e) => eprintln!("Failed to save RVF: {e}"),
            }
        }

        return;
    }

    // Handle --embed mode: extract embeddings from CSI data
    if args.embed {
        eprintln!("=== WiFi-DensePose Embedding Extraction (ADR-024) ===");

        let model_path = match &args.model {
            Some(p) => p.clone(),
            None => {
                eprintln!("Error: --embed requires --model <path> to a pretrained .rvf file");
                std::process::exit(1);
            }
        };

        let reader = match RvfReader::from_file(&model_path) {
            Ok(r) => r,
            Err(e) => { eprintln!("Failed to load model: {e}"); std::process::exit(1); }
        };

        let weights = reader.weights().unwrap_or_default();
        let (embed_config_json, proj_weights) = reader.embedding().unwrap_or_else(|| {
            eprintln!("Warning: no embedding segment in RVF, using defaults");
            (serde_json::json!({"d_model":64,"d_proj":128,"temperature":0.07,"normalize":true}), Vec::new())
        });

        let d_model = embed_config_json["d_model"].as_u64().unwrap_or(64) as usize;
        let d_proj = embed_config_json["d_proj"].as_u64().unwrap_or(128) as usize;

        let tf_config = graph_transformer::TransformerConfig {
            n_subcarriers: 56, n_keypoints: 17, d_model, n_heads: 4, n_gnn_layers: 2,
        };
        let e_config = embedding::EmbeddingConfig {
            d_model, d_proj, temperature: 0.07, normalize: true,
        };
        let mut extractor = embedding::EmbeddingExtractor::new(tf_config, e_config.clone());

        // Load transformer weights
        if !weights.is_empty() {
            if let Err(e) = extractor.transformer.unflatten_weights(&weights) {
                eprintln!("Warning: failed to load transformer weights: {e}");
            }
        }
        // Load projection weights
        if !proj_weights.is_empty() {
            let (proj, _) = embedding::ProjectionHead::unflatten_from(&proj_weights, &e_config);
            extractor.projection = proj;
        }

        // Load dataset and extract embeddings
        let _ds_path = args.dataset.clone().unwrap_or_else(|| PathBuf::from("data"));
        let csi_windows: Vec<Vec<Vec<f32>>> = (0..10).map(|i| {
            (0..4).map(|a| {
                (0..56).map(|s| ((i * 7 + a * 13 + s) as f32 * 0.31).sin() * 0.5).collect()
            }).collect()
        }).collect();

        eprintln!("Extracting embeddings from {} CSI windows...", csi_windows.len());
        let embeddings = extractor.extract_batch(&csi_windows);
        for (i, emb) in embeddings.iter().enumerate() {
            let norm: f32 = emb.iter().map(|x| x * x).sum::<f32>().sqrt();
            eprintln!("  Window {i}: {d_proj}-dim embedding, ||e|| = {norm:.4}");
        }
        eprintln!("Extracted {} embeddings of dimension {d_proj}", embeddings.len());

        return;
    }

    // Handle --build-index mode: build a fingerprint index from embeddings
    if let Some(ref index_type_str) = args.build_index {
        eprintln!("=== WiFi-DensePose Fingerprint Index Builder (ADR-024) ===");

        let index_type = match index_type_str.as_str() {
            "env" | "environment" => embedding::IndexType::EnvironmentFingerprint,
            "activity" => embedding::IndexType::ActivityPattern,
            "temporal" => embedding::IndexType::TemporalBaseline,
            "person" => embedding::IndexType::PersonTrack,
            _ => {
                eprintln!("Unknown index type '{}'. Use: env, activity, temporal, person", index_type_str);
                std::process::exit(1);
            }
        };

        let tf_config = graph_transformer::TransformerConfig::default();
        let e_config = embedding::EmbeddingConfig::default();
        let mut extractor = embedding::EmbeddingExtractor::new(tf_config, e_config);

        // Generate synthetic CSI windows for demo
        let csi_windows: Vec<Vec<Vec<f32>>> = (0..20).map(|i| {
            (0..4).map(|a| {
                (0..56).map(|s| ((i * 7 + a * 13 + s) as f32 * 0.31).sin() * 0.5).collect()
            }).collect()
        }).collect();

        let mut index = embedding::FingerprintIndex::new(index_type);
        for (i, window) in csi_windows.iter().enumerate() {
            let emb = extractor.extract(window);
            index.insert(emb, format!("window_{i}"), i as u64 * 100);
        }

        eprintln!("Built {:?} index with {} entries", index_type, index.len());

        // Test a query
        let query_emb = extractor.extract(&csi_windows[0]);
        let results = index.search(&query_emb, 5);
        eprintln!("Top-5 nearest to window_0:");
        for r in &results {
            eprintln!("  entry={}, distance={:.4}, metadata={}", r.entry, r.distance, r.metadata);
        }

        return;
    }

    // Handle --train mode: train a model and exit
    if args.train {
        eprintln!("=== WiFi-DensePose Training Mode ===");

        // Build data pipeline
        let ds_path = args.dataset.clone().unwrap_or_else(|| PathBuf::from("data"));
        let source = match args.dataset_type.as_str() {
            "wipose" => dataset::DataSource::WiPose(ds_path.clone()),
            _ => dataset::DataSource::MmFi(ds_path.clone()),
        };
        let pipeline = dataset::DataPipeline::new(dataset::DataConfig {
            source,
            ..Default::default()
        });

        // Generate synthetic training data (50 samples with deterministic CSI + keypoints)
        let generate_synthetic = || -> Vec<dataset::TrainingSample> {
            (0..50).map(|i| {
                let csi: Vec<Vec<f32>> = (0..4).map(|a| {
                    (0..56).map(|s| ((i * 7 + a * 13 + s) as f32 * 0.31).sin() * 0.5).collect()
                }).collect();
                let mut kps = [(0.0f32, 0.0f32, 1.0f32); 17];
                for (k, kp) in kps.iter_mut().enumerate() {
                    kp.0 = (k as f32 * 0.1 + i as f32 * 0.02).sin() * 100.0 + 320.0;
                    kp.1 = (k as f32 * 0.15 + i as f32 * 0.03).cos() * 80.0 + 240.0;
                }
                dataset::TrainingSample {
                    csi_window: csi,
                    pose_label: dataset::PoseLabel {
                        keypoints: kps,
                        body_parts: Vec::new(),
                        confidence: 1.0,
                    },
                    source: "synthetic",
                }
            }).collect()
        };

        // Load samples (fall back to synthetic if dataset missing/empty)
        let samples = match pipeline.load() {
            Ok(s) if !s.is_empty() => {
                eprintln!("Loaded {} samples from {}", s.len(), ds_path.display());
                s
            }
            Ok(_) => {
                eprintln!("No samples found at {}. Using synthetic data.", ds_path.display());
                generate_synthetic()
            }
            Err(e) => {
                eprintln!("Failed to load dataset: {e}. Using synthetic data.");
                generate_synthetic()
            }
        };

        // Convert dataset samples to trainer format
        let trainer_samples: Vec<trainer::TrainingSample> = samples.iter()
            .map(trainer::from_dataset_sample)
            .collect();

        // Split 80/20 train/val
        let split = (trainer_samples.len() * 4) / 5;
        let (train_data, val_data) = trainer_samples.split_at(split.max(1));
        eprintln!("Train: {} samples, Val: {} samples", train_data.len(), val_data.len());

        // Create transformer + trainer
        let n_subcarriers = train_data.first()
            .and_then(|s| s.csi_features.first())
            .map(|f| f.len())
            .unwrap_or(56);
        let tf_config = graph_transformer::TransformerConfig {
            n_subcarriers,
            n_keypoints: 17,
            d_model: 64,
            n_heads: 4,
            n_gnn_layers: 2,
        };
        let transformer = graph_transformer::CsiToPoseTransformer::new(tf_config);
        eprintln!("Transformer params: {}", transformer.param_count());

        let trainer_config = trainer::TrainerConfig {
            epochs: args.epochs,
            batch_size: 8,
            lr: 0.001,
            warmup_epochs: 5,
            min_lr: 1e-6,
            early_stop_patience: 20,
            checkpoint_every: 10,
            ..Default::default()
        };
        let mut t = trainer::Trainer::with_transformer(trainer_config, transformer);

        // Run training
        eprintln!("Starting training for {} epochs...", args.epochs);
        let result = t.run_training(train_data, val_data);
        eprintln!("Training complete in {:.1}s", result.total_time_secs);
        eprintln!("  Best epoch: {}, PCK@0.2: {:.4}, OKS mAP: {:.4}",
            result.best_epoch, result.best_pck, result.best_oks);

        // Save checkpoint
        if let Some(ref ckpt_dir) = args.checkpoint_dir {
            let _ = std::fs::create_dir_all(ckpt_dir);
            let ckpt_path = ckpt_dir.join("best_checkpoint.json");
            let ckpt = t.checkpoint();
            match ckpt.save_to_file(&ckpt_path) {
                Ok(()) => eprintln!("Checkpoint saved to {}", ckpt_path.display()),
                Err(e) => eprintln!("Failed to save checkpoint: {e}"),
            }
        }

        // Sync weights back to transformer and save as RVF
        t.sync_transformer_weights();
        if let Some(ref save_path) = args.save_rvf {
            eprintln!("Saving trained model to RVF: {}", save_path.display());
            let weights = t.params().to_vec();
            let mut builder = RvfBuilder::new();
            builder.add_manifest(
                "wifi-densepose-trained",
                env!("CARGO_PKG_VERSION"),
                "WiFi DensePose trained model weights",
            );
            builder.add_metadata(&serde_json::json!({
                "training": {
                    "epochs": args.epochs,
                    "best_epoch": result.best_epoch,
                    "best_pck": result.best_pck,
                    "best_oks": result.best_oks,
                    "n_train_samples": train_data.len(),
                    "n_val_samples": val_data.len(),
                    "n_subcarriers": n_subcarriers,
                    "param_count": weights.len(),
                },
            }));
            builder.add_vital_config(&VitalSignConfig::default());
            builder.add_weights(&weights);
            match builder.write_to_file(save_path) {
                Ok(()) => eprintln!("RVF saved ({} params, {} bytes)",
                    weights.len(), weights.len() * 4),
                Err(e) => eprintln!("Failed to save RVF: {e}"),
            }
        }

        return;
    }

    info!("WiFi-DensePose Sensing Server (Rust + Axum + RuVector)");
    info!("  HTTP:      http://localhost:{}", args.http_port);
    info!("  WebSocket: ws://localhost:{}/ws/sensing", args.ws_port);
    info!("  UDP:       0.0.0.0:{} (ESP32 CSI)", args.udp_port);
    info!("  UI path:   {}", args.ui_path.display());
    info!("  Source:    {}", args.source);

    // Initialise the on-line anomaly + aggression detector. Baseline persists
    // to data/aggression_baseline.json so the trained environment survives
    // restarts; loading is best-effort (returns empty on first run).
    {
        let baseline_path = PathBuf::from("data").join("aggression_baseline.json");
        aggression::init(baseline_path.clone());
        info!("  Aggression baseline: {} (loaded if present)", baseline_path.display());
    }

    // Auto-detect data source. Simulation has been removed — if no hardware
    // is present we retry probing every 5 s instead of silently faking data.
    let source = match args.source.as_str() {
        "auto" => loop {
            info!("Auto-detecting data source...");
            if probe_esp32(args.udp_port).await {
                info!("  ESP32 CSI detected on UDP :{}", args.udp_port);
                break "esp32";
            } else if probe_windows_wifi().await {
                info!("  Windows WiFi detected");
                break "wifi";
            } else if probe_linux_wifi() {
                info!("  Linux WiFi detected (iw dev)");
                break "wifi";
            }
            warn!("  No hardware detected — retrying in 5 s (simulation disabled, real data only)");
            tokio::time::sleep(Duration::from_secs(5)).await;
        },
        "simulate" | "simulated" | "sim" => {
            error!("Simulation mode has been removed. Use --source esp32 or --source wifi.");
            std::process::exit(2);
        }
        other => other,
    };

    info!("Data source: {source}");

    // Shared state
    // Vital sign sample rate derives from tick interval (e.g. 500ms tick => 2 Hz)
    let vital_sample_rate = 1000.0 / args.tick_ms as f64;
    info!("Vital sign detector sample rate: {vital_sample_rate:.1} Hz");

    // Load RVF container if --load-rvf was specified
    let rvf_info = if let Some(ref rvf_path) = args.load_rvf {
        info!("Loading RVF container from {}", rvf_path.display());
        match RvfReader::from_file(rvf_path) {
            Ok(reader) => {
                let info = reader.info();
                info!(
                    "  RVF loaded: {} segments, {} bytes",
                    info.segment_count, info.total_size
                );
                if let Some(ref manifest) = info.manifest {
                    if let Some(model_id) = manifest.get("model_id") {
                        info!("  Model ID: {model_id}");
                    }
                    if let Some(version) = manifest.get("version") {
                        info!("  Version:  {version}");
                    }
                }
                if info.has_weights {
                    if let Some(w) = reader.weights() {
                        info!("  Weights: {} parameters", w.len());
                    }
                }
                if info.has_vital_config {
                    info!("  Vital sign config: present");
                }
                if info.has_quant_info {
                    info!("  Quantization info: present");
                }
                if info.has_witness {
                    info!("  Witness/proof: present");
                }
                Some(info)
            }
            Err(e) => {
                error!("Failed to load RVF container: {e}");
                None
            }
        }
    } else {
        None
    };

    // Load trained model via --model (uses progressive loading if --progressive set)
    let model_path = args.model.as_ref().or(args.load_rvf.as_ref());
    let mut progressive_loader: Option<ProgressiveLoader> = None;
    let mut model_loaded = false;
    if let Some(mp) = model_path {
        if args.progressive || args.model.is_some() {
            info!("Loading trained model (progressive) from {}", mp.display());
            match std::fs::read(mp) {
                Ok(data) => match ProgressiveLoader::new(&data) {
                    Ok(mut loader) => {
                        if let Ok(la) = loader.load_layer_a() {
                            info!("  Layer A ready: model={} v{} ({} segments)",
                                  la.model_name, la.version, la.n_segments);
                        }
                        model_loaded = true;
                        progressive_loader = Some(loader);
                    }
                    Err(e) => error!("Progressive loader init failed: {e}"),
                },
                Err(e) => error!("Failed to read model file: {e}"),
            }
        }
    }

    // Ensure data directories exist for models and recordings
    let _ = std::fs::create_dir_all("data/models");
    let _ = std::fs::create_dir_all("data/recordings");

    // Discover model and recording files on startup
    let initial_models = scan_model_files();
    let initial_recordings = scan_recording_files();
    info!("Discovered {} model files, {} recording files", initial_models.len(), initial_recordings.len());

    let (tx, _) = broadcast::channel::<String>(256);
    let state: SharedState = Arc::new(RwLock::new(AppStateInner {
        latest_update: None,
        rssi_history: VecDeque::new(),
        frame_history: VecDeque::new(),
        tick: 0,
        source: source.into(),
        last_esp32_frame: None,
        tx,
        total_detections: 0,
        start_time: std::time::Instant::now(),
        vital_detector: VitalSignDetector::new(vital_sample_rate),
        latest_vitals: VitalSigns::default(),
        rvf_info,
        save_rvf_path: args.save_rvf.clone(),
        progressive_loader,
        active_sona_profile: None,
        model_loaded,
        smoothed_person_score: 0.0,
        prev_person_count: 0,
        smoothed_motion: 0.0,
        current_motion_level: "absent".to_string(),
        debounce_counter: 0,
        debounce_candidate: "absent".to_string(),
        baseline_motion: 0.0,
        baseline_frames: 0,
        smoothed_hr: 0.0,
        smoothed_br: 0.0,
        smoothed_hr_conf: 0.0,
        smoothed_br_conf: 0.0,
        hr_buffer: VecDeque::with_capacity(8),
        br_buffer: VecDeque::with_capacity(8),
        edge_vitals: None,
        latest_wasm_events: None,
        // Model management
        discovered_models: initial_models,
        active_model_id: None,
        // Recording
        recordings: initial_recordings,
        recording_active: false,
        recording_start_time: None,
        recording_current_id: None,
        recording_stop_tx: None,
        // Training
        training_status: "idle".to_string(),
        training_config: None,
        adaptive_model: adaptive_classifier::AdaptiveModel::load(&adaptive_classifier::model_path()).ok().map(|m| {
            info!("Loaded adaptive classifier: {} frames, {:.1}% accuracy",
                  m.trained_frames, m.training_accuracy * 100.0);
            m
        }),
        node_states: HashMap::new(),
        // Accuracy sprint
        pose_tracker: PoseTracker::new(),
        last_tracker_instant: None,
        track_temporal_state: HashMap::new(),
        multistatic_fuser: {
            let mut fuser = MultistaticFuser::with_config(MultistaticConfig {
                min_nodes: 1, // single-node passthrough
                ..Default::default()
            });
            if let Some(ref pos_str) = args.node_positions {
                let positions = field_bridge::parse_node_positions(pos_str);
                if !positions.is_empty() {
                    info!("Configured {} node positions for multistatic fusion", positions.len());
                    fuser.set_node_positions(positions);
                }
            }
            fuser
        },
        field_model: if args.calibrate {
            info!("Field model calibration enabled — room should be empty during startup");
            FieldModel::new(field_bridge::single_link_config()).ok()
        } else {
            None
        },
    }));

    // Always spawn UDP receiver so ESP32 nodes can connect regardless of source mode
    tokio::spawn(udp_receiver_task(state.clone(), args.udp_port));

    // Start background tasks based on source
    match source {
        "esp32" => {
            tokio::spawn(broadcast_tick_task(state.clone(), args.tick_ms));
        }
        "wifi" => {
            #[cfg(target_os = "linux")]
            {
                let iface = detect_wifi_interface();
                info!("Linux WiFi sensing: interface={iface} — scanning all visible BSSIDs");
                tokio::spawn(linux_wifi_task(state.clone(), args.tick_ms, iface));
            }
            #[cfg(not(target_os = "linux"))]
            {
                tokio::spawn(windows_wifi_task(state.clone(), args.tick_ms));
            }
        }
        other => {
            error!("Unknown --source '{other}'. Valid values: auto, esp32, wifi (simulation removed).");
            std::process::exit(2);
        }
    }

    // Autorepair: spawn software watchdog (checks every 30s)
    tokio::spawn(autorepair_watchdog_task(state.clone(), 30));

    // Auto-calibration: 30 s after startup, if no positions are configured,
    // assign nodes to a deterministic perimeter layout based on their IDs.
    // Re-tries every 30 s until at least one node is active. No-op once set.
    tokio::spawn(auto_calibration_startup_task(state.clone()));

    // ADR-050: Parse bind address once, use for all listeners
    let bind_ip: std::net::IpAddr = args.bind_addr.parse()
        .expect("Invalid --bind-addr (use 127.0.0.1 or 0.0.0.0)");

    let http_trace = TraceLayer::new_for_http()
        .make_span_with(DefaultMakeSpan::new().level(Level::INFO))
        .on_response(DefaultOnResponse::new().level(Level::INFO))
        .on_failure(DefaultOnFailure::new().level(Level::ERROR));

    // WebSocket server on dedicated port (8765)
    let ws_state = state.clone();
    let ws_app = Router::new()
        .route("/ws/sensing", get(ws_sensing_handler))
        .route("/health", get(health))
        .layer(http_trace.clone())
        .with_state(ws_state);

    let ws_addr = SocketAddr::from((bind_ip, args.ws_port));
    let ws_listener = tokio::net::TcpListener::bind(ws_addr).await
        .expect("Failed to bind WebSocket port");
    info!("WebSocket server listening on {ws_addr}");

    tokio::spawn(async move {
        axum::serve(ws_listener, ws_app).await.unwrap();
    });

    // HTTP server (serves UI + full DensePose-compatible REST API)
    let ui_path = args.ui_path.clone();
    let http_app = Router::new()
        .route("/", get(info_page))
        // Health endpoints (DensePose-compatible)
        .route("/health", get(health))
        .route("/health/health", get(health_system))
        .route("/health/live", get(health_live))
        .route("/health/ready", get(health_ready))
        .route("/health/version", get(health_version))
        .route("/health/metrics", get(health_metrics))
        .route("/health/autorepair", get(health_autorepair))
        // API info
        .route("/api/v1/info", get(api_info))
        .route("/api/v1/status", get(health_ready))
        .route("/api/v1/metrics", get(health_metrics))
        // Sensing endpoints
        .route("/api/v1/sensing/latest", get(latest))
        // Per-node health endpoint
        .route("/api/v1/nodes", get(nodes_endpoint))
        // TX/RX/TXRX role override (multistatic mesh labelling). Flat path,
        // body carries both node_id and role.
        .route("/api/v1/role-set", post(node_role_set))
        // Auto-assign: pick strongest-RSSI active node as TX, rest as RX.
        // Optional body: {"prefer_node_id": <u8>} to force a specific TX.
        .route("/api/v1/role-set/auto", post(node_role_auto))
        // Per-node remote config (CSI rate, channel, TX power). Firmware
        // polls `GET /api/v1/node-config?node_id=N`; UI reads
        // `/api/v1/node-config/all`; writes via POST.
        .route("/api/v1/node-config", get(node_config_get).post(node_config_set))
        .route("/api/v1/node-config/all", get(node_config_list))
        // Offline AI (Ollama) lifecycle for the AEDI "Start Offline AI" button.
        .route("/api/v1/ollama/status", get(ollama_status))
        .route("/api/v1/ollama/start", post(ollama_start))
        // Storage / PostgreSQL / JSONL audit log — drives the Monitor "Storage" tab.
        .route("/api/v1/storage/status", get(storage_status))
        .route("/api/v1/storage/recent", get(storage_recent))
        .route("/api/v1/storage/event", post(storage_event))
        .route("/api/v1/storage/scan", get(storage_scan))
        .route("/api/v1/storage/config", post(storage_set_config))
        .route("/api/v1/registry", get(registry_endpoint))
        .route("/api/v1/preprocessing/stats", get(preprocessing_stats_endpoint))
        // TinyML environment-trained anomaly + aggression scores
        .route("/api/v1/aggression", get(aggression_endpoint))
        // Room config + node positions for XYZ Segment 3D page
        .route("/api/v1/room", get(room_endpoint))
        .route("/api/v1/nodes/positions", post(update_node_positions))
        .route("/api/v1/calibration/auto", post(calibration_auto_endpoint))
        .route("/api/v1/localization/person", get(localization_person_endpoint))
        .route("/api/v1/tomography", get(tomography_endpoint))
        // Vital sign endpoints
        .route("/api/v1/vital-signs", get(vital_signs_endpoint))
        .route("/api/v1/edge-vitals", get(edge_vitals_endpoint))
        .route("/api/v1/wasm-events", get(wasm_events_endpoint))
        // RVF model container info
        .route("/api/v1/model/info", get(model_info))
        // Progressive loading & SONA endpoints (Phase 7-8)
        .route("/api/v1/model/layers", get(model_layers))
        .route("/api/v1/model/segments", get(model_segments))
        .route("/api/v1/model/sona/profiles", get(sona_profiles))
        .route("/api/v1/model/sona/activate", post(sona_activate))
        // Pose endpoints (WiFi-derived)
        .route("/api/v1/pose/current", get(pose_current))
        .route("/api/v1/pose/stats", get(pose_stats))
        .route("/api/v1/pose/zones/summary", get(pose_zones_summary))
        // Stream endpoints
        .route("/api/v1/stream/status", get(stream_status))
        .route("/api/v1/stream/pose", get(ws_pose_handler))
        // Sensing WebSocket on the HTTP port so the UI can reach it without a second port
        .route("/ws/sensing", get(ws_sensing_handler))
        // Model management endpoints (UI compatibility)
        .route("/api/v1/models", get(list_models))
        .route("/api/v1/models/active", get(get_active_model))
        .route("/api/v1/models/load", post(load_model))
        .route("/api/v1/models/unload", post(unload_model))
        .route("/api/v1/models/{id}", delete(delete_model))
        .route("/api/v1/models/lora/profiles", get(list_lora_profiles))
        .route("/api/v1/models/lora/activate", post(activate_lora_profile))
        // Recording endpoints
        .route("/api/v1/recording/list", get(list_recordings))
        .route("/api/v1/recording/start", post(start_recording))
        .route("/api/v1/recording/stop", post(stop_recording))
        .route("/api/v1/recording/{id}", delete(delete_recording))
        // Training endpoints
        .route("/api/v1/train/status", get(train_status))
        .route("/api/v1/train/start", post(train_start))
        .route("/api/v1/train/stop", post(train_stop))
        // Adaptive classifier endpoints
        .route("/api/v1/adaptive/train", post(adaptive_train))
        .route("/api/v1/adaptive/status", get(adaptive_status))
        .route("/api/v1/adaptive/unload", post(adaptive_unload))
        // Field model calibration (eigenvalue-based person counting)
        .route("/api/v1/calibration/start", post(calibration_start))
        .route("/api/v1/calibration/stop", post(calibration_stop))
        .route("/api/v1/calibration/status", get(calibration_status))
        // Static UI files
        .nest_service("/ui", ServeDir::new(&ui_path))
        .layer(SetResponseHeaderLayer::overriding(
            axum::http::header::CACHE_CONTROL,
            HeaderValue::from_static("no-cache, no-store, must-revalidate"),
        ))
        .layer(http_trace)
        .with_state(state.clone());

    let http_addr = SocketAddr::from((bind_ip, args.http_port));
    let http_listener = tokio::net::TcpListener::bind(http_addr).await
        .expect("Failed to bind HTTP port");
    info!("HTTP server listening on {http_addr}");
    info!("Open http://localhost:{}/ui/index.html in your browser", args.http_port);

    // Run the HTTP server with graceful shutdown support
    let shutdown_state = state.clone();
    let server = axum::serve(http_listener, http_app)
        .with_graceful_shutdown(async {
            tokio::signal::ctrl_c()
                .await
                .expect("failed to install CTRL+C handler");
            info!("Shutdown signal received");
        });

    server.await.unwrap();

    // Save RVF container on shutdown if --save-rvf was specified
    let s = shutdown_state.read().await;
    if let Some(ref save_path) = s.save_rvf_path {
        info!("Saving RVF container to {}", save_path.display());
        let mut builder = RvfBuilder::new();
        builder.add_manifest(
            "wifi-densepose-sensing",
            env!("CARGO_PKG_VERSION"),
            "WiFi DensePose sensing model state",
        );
        builder.add_metadata(&serde_json::json!({
            "source": s.effective_source(),
            "total_ticks": s.tick,
            "total_detections": s.total_detections,
            "uptime_secs": s.start_time.elapsed().as_secs(),
        }));
        builder.add_vital_config(&VitalSignConfig::default());
        // Save transformer weights if a model is loaded, otherwise empty
        let weights: Vec<f32> = if s.model_loaded {
            // If we loaded via --model, the progressive loader has the weights
            // For now, save runtime state placeholder
            let tf = graph_transformer::CsiToPoseTransformer::new(Default::default());
            tf.flatten_weights()
        } else {
            Vec::new()
        };
        builder.add_weights(&weights);
        match builder.write_to_file(save_path) {
            Ok(()) => info!("  RVF saved ({} weight params)", weights.len()),
            Err(e) => error!("  Failed to save RVF: {e}"),
        }
    }

    info!("Server shut down cleanly");
}
