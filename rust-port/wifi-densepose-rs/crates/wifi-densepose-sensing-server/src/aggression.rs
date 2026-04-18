//! Lightweight TinyML-style anomaly + aggression detector.
//!
//! Runs entirely in the sensing-server process (single-threaded behind a
//! `Mutex`), no external model dependencies. Trains an *online* per-node
//! baseline of `motion_band_power` using Welford's algorithm, then derives
//! three live numbers per WS broadcast tick:
//!
//! * **anomaly**   — clamped tanh(z-score / 3) where z = (x - mean) / std.
//!                   0.0 = matches baseline, 1.0 = far outside trained range.
//! * **aggression**— positive normalised derivative of the short-horizon EMA
//!                   relative to the long-horizon EMA. Catches *escalation*
//!                   (slow walk → sudden lunge) which a pure z-score misses.
//! * **predicted** — naive 30-sample-ahead linear extrapolation of the EMA
//!                   pair, useful as a pre-event lead signal for the UI.
//!
//! State persists to `data/aggression_baseline.json` so the "trained
//! environment" survives restarts. Saved once a minute, atomically.
//!
//! Deliberately NOT a published-API crate: this is internal to the binary.
//! Kept small and dependency-free (only `serde` + `std`).

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::OnceLock;
use std::sync::Mutex;
use std::time::{SystemTime, UNIX_EPOCH};

use serde::{Deserialize, Serialize};

const TRAIN_TARGET_SAMPLES: f64 = 1000.0; // ~5-6 min @ ~3 Hz per node
const EMA_SHORT_ALPHA: f64 = 0.30;
const EMA_LONG_ALPHA: f64 = 0.05;
const FORECAST_STEPS: f64 = 30.0;
const SAVE_INTERVAL_SECS: u64 = 60;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct NodeBaseline {
    /// Welford running mean.
    mean: f64,
    /// Welford running M2 (sum of squared deviations).
    m2: f64,
    /// Number of observations included in the baseline.
    count: u64,
    ema_short: f64,
    ema_long: f64,
    /// Last observed motion_band_power (for derivative).
    last_value: f64,
    /// Last sensing tick observed (idempotency). `None` means never seen,
    /// which is distinct from tick=0 — we want to accept tick 0 once.
    #[serde(default)]
    last_tick: Option<u64>,
}

impl NodeBaseline {
    fn observe(&mut self, x: f64) {
        // Welford online mean / variance.
        self.count += 1;
        let delta = x - self.mean;
        self.mean += delta / (self.count as f64);
        let delta2 = x - self.mean;
        self.m2 += delta * delta2;

        if self.count == 1 {
            self.ema_short = x;
            self.ema_long = x;
        } else {
            self.ema_short = EMA_SHORT_ALPHA * x + (1.0 - EMA_SHORT_ALPHA) * self.ema_short;
            self.ema_long = EMA_LONG_ALPHA * x + (1.0 - EMA_LONG_ALPHA) * self.ema_long;
        }
        self.last_value = x;
    }

    fn variance(&self) -> f64 {
        if self.count < 2 {
            0.0
        } else {
            self.m2 / (self.count as f64 - 1.0)
        }
    }

    fn std_dev(&self) -> f64 {
        self.variance().sqrt()
    }

    /// Stabilised z-score: returns 0 until we have enough samples.
    fn z_score(&self, x: f64) -> f64 {
        let s = self.std_dev();
        if self.count < 30 || s < 1e-6 {
            0.0
        } else {
            (x - self.mean) / s
        }
    }

    fn trained_pct(&self) -> f64 {
        ((self.count as f64) / TRAIN_TARGET_SAMPLES).min(1.0)
    }
}

#[derive(Debug, Clone, Serialize)]
pub struct AggressionScore {
    pub node_id: u8,
    /// 0..1 — how unusual the current value is vs trained baseline.
    pub anomaly: f64,
    /// 0..1 — rate-of-escalation signal (short-EMA pulling away from long-EMA).
    pub aggression: f64,
    /// Predicted motion_band_power ~30 samples ahead.
    pub predicted: f64,
    pub baseline_mean: f64,
    pub baseline_std: f64,
    /// 0..1 — how trained this node's environment baseline is.
    pub trained_pct: f64,
    pub samples: u64,
}

#[derive(Debug, Clone, Serialize)]
pub struct AggregateScore {
    /// max anomaly across nodes — primary "something is wrong" gauge.
    pub anomaly: f64,
    /// max aggression across nodes — primary "something is escalating" gauge.
    pub aggression: f64,
    /// average trained_pct across nodes — environment learning progress.
    pub trained_pct: f64,
    /// Per-node detail.
    pub nodes: Vec<AggressionScore>,
}

#[derive(Debug, Default, Serialize, Deserialize)]
pub struct Detector {
    nodes: HashMap<u8, NodeBaseline>,
    /// Unix seconds of last successful save_to_disk().
    #[serde(default)]
    last_save_unix: u64,
}

impl Detector {
    pub fn new() -> Self {
        Self::default()
    }

    /// Try to load from disk; on any error return a fresh detector.
    pub fn load_or_new(path: &Path) -> Self {
        std::fs::read_to_string(path)
            .ok()
            .and_then(|s| serde_json::from_str::<Self>(&s).ok())
            .unwrap_or_default()
    }

    /// Atomic save: write to `<path>.tmp` then rename.
    pub fn save(&mut self, path: &Path) -> std::io::Result<()> {
        if let Some(parent) = path.parent() {
            std::fs::create_dir_all(parent)?;
        }
        let tmp = path.with_extension("json.tmp");
        let json = serde_json::to_string_pretty(self).map_err(|e| {
            std::io::Error::new(std::io::ErrorKind::InvalidData, e)
        })?;
        std::fs::write(&tmp, json)?;
        std::fs::rename(&tmp, path)?;
        self.last_save_unix = unix_now();
        Ok(())
    }

    pub fn maybe_save(&mut self, path: &Path) {
        if unix_now().saturating_sub(self.last_save_unix) >= SAVE_INTERVAL_SECS {
            // Best-effort, never panic the broadcast loop on disk errors.
            let _ = self.save(path);
        }
    }

    /// Observe a single (node_id, motion_band_power, tick) sample.
    /// Idempotent on tick — repeated calls with the same tick for a node only
    /// update the baseline once (callers may run per-WS-client).
    pub fn observe(&mut self, node_id: u8, x: f64, tick: u64) -> AggressionScore {
        let entry = self.nodes.entry(node_id).or_default();
        let new_tick = entry.last_tick.map_or(true, |last| tick > last);
        if x.is_finite() && new_tick {
            entry.observe(x);
            entry.last_tick = Some(tick);
        }
        let z = entry.z_score(x);
        let anomaly = (z.abs() / 3.0).tanh().clamp(0.0, 1.0);
        let mean = entry.mean;
        let scale = entry.std_dev().max(1e-3).max(mean.abs() * 0.05);
        let agg_raw = (entry.ema_short - entry.ema_long).max(0.0) / scale;
        let aggression = (agg_raw / 2.0).tanh().clamp(0.0, 1.0);
        let predicted = entry.ema_short + (entry.ema_short - entry.ema_long) * FORECAST_STEPS / 10.0;
        AggressionScore {
            node_id,
            anomaly,
            aggression,
            predicted,
            baseline_mean: mean,
            baseline_std: entry.std_dev(),
            trained_pct: entry.trained_pct(),
            samples: entry.count,
        }
    }

    /// Aggregate snapshot across all known nodes (no observation).
    pub fn aggregate(&self) -> AggregateScore {
        let mut node_scores: Vec<AggressionScore> = self
            .nodes
            .iter()
            .map(|(&id, b)| {
                let z = b.z_score(b.last_value);
                let anomaly = (z.abs() / 3.0).tanh().clamp(0.0, 1.0);
                let scale = b.std_dev().max(1e-3).max(b.mean.abs() * 0.05);
                let aggression = (((b.ema_short - b.ema_long).max(0.0) / scale) / 2.0)
                    .tanh()
                    .clamp(0.0, 1.0);
                let predicted = b.ema_short + (b.ema_short - b.ema_long) * FORECAST_STEPS / 10.0;
                AggressionScore {
                    node_id: id,
                    anomaly,
                    aggression,
                    predicted,
                    baseline_mean: b.mean,
                    baseline_std: b.std_dev(),
                    trained_pct: b.trained_pct(),
                    samples: b.count,
                }
            })
            .collect();
        node_scores.sort_by_key(|s| s.node_id);

        let max_anomaly = node_scores.iter().map(|s| s.anomaly).fold(0.0_f64, f64::max);
        let max_aggression = node_scores.iter().map(|s| s.aggression).fold(0.0_f64, f64::max);
        let avg_trained = if node_scores.is_empty() {
            0.0
        } else {
            node_scores.iter().map(|s| s.trained_pct).sum::<f64>() / (node_scores.len() as f64)
        };

        AggregateScore {
            anomaly: max_anomaly,
            aggression: max_aggression,
            trained_pct: avg_trained,
            nodes: node_scores,
        }
    }
}

fn unix_now() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_secs())
        .unwrap_or(0)
}

// ── Process-wide singleton ───────────────────────────────────────────────────

static GLOBAL: OnceLock<Mutex<Detector>> = OnceLock::new();
static GLOBAL_PATH: OnceLock<PathBuf> = OnceLock::new();

/// Initialise the singleton once. `path` is where the JSON baseline is
/// loaded from / persisted to. Safe to call multiple times — only the
/// first call wins.
pub fn init(path: PathBuf) {
    let _ = GLOBAL_PATH.set(path.clone());
    let _ = GLOBAL.set(Mutex::new(Detector::load_or_new(&path)));
}

/// Convenience wrapper used by the WS broadcast hook.
/// Observes one (node_id, value, tick), returns the score, and triggers
/// a periodic atomic save. Returns `None` if `init()` was never called.
pub fn observe(node_id: u8, motion_band_power: f64, tick: u64) -> Option<AggressionScore> {
    let cell = GLOBAL.get()?;
    let path = GLOBAL_PATH.get()?.clone();
    let mut guard = cell.lock().ok()?;
    let score = guard.observe(node_id, motion_band_power, tick);
    guard.maybe_save(&path);
    Some(score)
}

/// Snapshot of the environment-wide aggregate score.
pub fn snapshot() -> Option<AggregateScore> {
    let cell = GLOBAL.get()?;
    let guard = cell.lock().ok()?;
    Some(guard.aggregate())
}

// ── Tests ────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn baseline_z_score_is_zero_until_warm() {
        let mut d = Detector::new();
        // First handful of samples → should not produce wild scores.
        for _ in 0..5 {
            let s = d.observe(1, 0.10, 1);  // same tick, idempotent
            assert!(s.anomaly < 1e-6, "anomaly should be 0 while warming up");
        }
    }

    #[test]
    fn anomaly_rises_on_outlier_after_training() {
        let mut d = Detector::new();
        // Stationary noisy baseline centred on 0.10 with std ~0.05.
        for t in 0..200u64 {
            let noise = ((t as f64 * 0.731).sin() + (t as f64 * 1.913).cos()) * 0.025;
            d.observe(7, 0.10 + noise, t);
        }
        let calm = d.observe(7, 0.10, 201);
        let spike = d.observe(7, 5.00, 202);
        assert!(calm.anomaly < 0.3, "calm score too high: {}", calm.anomaly);
        assert!(spike.anomaly > 0.8, "spike score too low: {}", spike.anomaly);
    }

    #[test]
    fn aggression_responds_to_escalation() {
        let mut d = Detector::new();
        for t in 0..100u64 {
            d.observe(3, 0.10, t);
        }
        // Steep escalation over a few ticks.
        for t in 100..120u64 {
            d.observe(3, 0.10 + (t as f64 - 100.0) * 0.05, t);
        }
        let s = d.observe(3, 1.10, 121);
        assert!(s.aggression > 0.3, "aggression too low for ramp: {}", s.aggression);
    }

    #[test]
    fn save_and_load_roundtrip() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("baseline.json");
        let mut d = Detector::new();
        for t in 0..50u64 {
            d.observe(2, 0.20, t);
        }
        d.save(&path).unwrap();
        let d2 = Detector::load_or_new(&path);
        let agg = d2.aggregate();
        assert_eq!(agg.nodes.len(), 1);
        assert_eq!(agg.nodes[0].samples, 50);
    }

    #[test]
    fn aggregate_max_across_nodes() {
        let mut d = Detector::new();
        for t in 0..100u64 {
            d.observe(1, 0.10, t);
            d.observe(2, 0.10, t);
        }
        d.observe(2, 9.0, 200);
        let agg = d.aggregate();
        assert!(agg.anomaly > 0.5, "aggregate anomaly should track worst node");
        assert_eq!(agg.nodes.len(), 2);
    }
}
