//! Particle Swarm Optimization (PSO) — generic bounded minimizer.
//!
//! Canonical constriction-factor PSO (Clerc & Kennedy 2002):
//!   v_{t+1} = χ · [ v_t + c1·r1·(pbest − x_t) + c2·r2·(gbest − x_t) ]
//!   x_{t+1} = x_t + v_{t+1}
//! where χ = 2 / |2 − φ − √(φ² − 4φ)|, φ = c1 + c2 > 4.
//!
//! The default preset (c1=c2=2.05, χ≈0.7298) guarantees convergence on
//! convex problems and behaves well on multimodal CSI cost surfaces.
//!
//! Dependencies: none beyond `std` — deterministic when `seed` is set, so
//! the localization endpoint can be unit-tested.
//!
//! Primary use case in AEDI-S:
//!   RSSI → position trilateration with log-distance path-loss
//!   (see [`estimate_position_rssi_pso`]). PSO is preferred over
//!   Gauss-Newton because the RSSI residual surface is non-convex
//!   (multipath nulls, NLOS) and PSO is derivative-free.

#![allow(clippy::needless_range_loop)]

use std::f64::consts::PI;

/// Tunable PSO hyperparameters.
#[derive(Debug, Clone)]
pub struct PsoConfig {
    pub swarm_size: usize,
    pub max_iters: usize,
    /// Cognitive coefficient (pull toward personal best).
    pub c1: f64,
    /// Social coefficient (pull toward global best).
    pub c2: f64,
    /// Early-stop if global-best cost stalls for this many iters.
    pub stall_iters: usize,
    /// Early-stop if cost < this threshold.
    pub cost_target: f64,
    /// Deterministic PRNG seed. `None` ⇒ nondeterministic (time-based).
    pub seed: Option<u64>,
}

impl Default for PsoConfig {
    fn default() -> Self {
        Self {
            swarm_size: 32,
            max_iters: 120,
            c1: 2.05,
            c2: 2.05,
            stall_iters: 20,
            cost_target: 1e-9,
            seed: None,
        }
    }
}

/// Outcome of a PSO run.
#[derive(Debug, Clone)]
pub struct PsoResult {
    pub best_x: Vec<f64>,
    pub best_cost: f64,
    pub iters_run: usize,
    pub converged: bool,
}

/// Minimise `objective(x)` subject to `bounds[i].0 ≤ x_i ≤ bounds[i].1`.
///
/// - `objective` must be deterministic w.r.t. `x` (no side effects).
/// - Returns the best position found across all particles.
///
/// # Example
/// ```ignore
/// use crate::pso::{pso_minimize, PsoConfig};
/// let sphere = |x: &[f64]| x.iter().map(|v| v*v).sum();
/// let bounds = vec![(-5.0, 5.0); 3];
/// let cfg = PsoConfig { seed: Some(42), ..Default::default() };
/// let r = pso_minimize(&sphere, &bounds, &cfg);
/// assert!(r.best_cost < 1e-4);
/// ```
pub fn pso_minimize<F>(objective: &F, bounds: &[(f64, f64)], cfg: &PsoConfig) -> PsoResult
where
    F: Fn(&[f64]) -> f64,
{
    assert!(!bounds.is_empty(), "pso_minimize: empty bounds");
    assert!(cfg.swarm_size >= 4, "pso_minimize: swarm_size must be ≥ 4");

    // Constriction factor χ (Clerc 1999). φ must exceed 4 for convergence.
    let phi = cfg.c1 + cfg.c2;
    let chi = if phi > 4.0 {
        2.0 / (phi - 2.0 + (phi * phi - 4.0 * phi).sqrt()).abs()
    } else {
        0.7298 // safe default
    };

    let dims = bounds.len();
    let mut rng = XorShift64::new(cfg.seed.unwrap_or_else(|| seed_from_time()));

    // Initialise particles uniformly in bounds; velocities in ±range.
    let mut positions: Vec<Vec<f64>> = Vec::with_capacity(cfg.swarm_size);
    let mut velocities: Vec<Vec<f64>> = Vec::with_capacity(cfg.swarm_size);
    let mut pbest: Vec<Vec<f64>> = Vec::with_capacity(cfg.swarm_size);
    let mut pbest_cost: Vec<f64> = Vec::with_capacity(cfg.swarm_size);
    let mut gbest: Vec<f64> = vec![0.0; dims];
    let mut gbest_cost = f64::INFINITY;

    for _ in 0..cfg.swarm_size {
        let mut x = vec![0.0f64; dims];
        let mut v = vec![0.0f64; dims];
        for d in 0..dims {
            let (lo, hi) = bounds[d];
            let range = hi - lo;
            x[d] = lo + rng.uniform() * range;
            v[d] = (rng.uniform() * 2.0 - 1.0) * range * 0.1;
        }
        let cost = objective(&x);
        pbest_cost.push(cost);
        pbest.push(x.clone());
        if cost < gbest_cost {
            gbest_cost = cost;
            gbest = x.clone();
        }
        positions.push(x);
        velocities.push(v);
    }

    // Main loop.
    let mut stall = 0usize;
    let mut iters_run = 0usize;
    let mut converged = false;
    for it in 0..cfg.max_iters {
        iters_run = it + 1;
        let prev_gbest_cost = gbest_cost;

        for i in 0..cfg.swarm_size {
            for d in 0..dims {
                let r1 = rng.uniform();
                let r2 = rng.uniform();
                let cog = cfg.c1 * r1 * (pbest[i][d] - positions[i][d]);
                let soc = cfg.c2 * r2 * (gbest[d] - positions[i][d]);
                velocities[i][d] = chi * (velocities[i][d] + cog + soc);

                // Velocity clamp to prevent divergence (0.5·range).
                let (lo, hi) = bounds[d];
                let vmax = 0.5 * (hi - lo);
                if velocities[i][d] > vmax { velocities[i][d] = vmax; }
                if velocities[i][d] < -vmax { velocities[i][d] = -vmax; }

                positions[i][d] += velocities[i][d];

                // Reflecting boundary: bounce off walls.
                if positions[i][d] < lo {
                    positions[i][d] = lo;
                    velocities[i][d] = -velocities[i][d] * 0.5;
                } else if positions[i][d] > hi {
                    positions[i][d] = hi;
                    velocities[i][d] = -velocities[i][d] * 0.5;
                }
            }
            let cost = objective(&positions[i]);
            if cost < pbest_cost[i] {
                pbest_cost[i] = cost;
                pbest[i].copy_from_slice(&positions[i]);
                if cost < gbest_cost {
                    gbest_cost = cost;
                    gbest.copy_from_slice(&positions[i]);
                }
            }
        }

        if gbest_cost < cfg.cost_target {
            converged = true;
            break;
        }
        if (prev_gbest_cost - gbest_cost).abs() < 1e-12 {
            stall += 1;
            if stall >= cfg.stall_iters {
                converged = true;
                break;
            }
        } else {
            stall = 0;
        }
    }

    PsoResult {
        best_x: gbest,
        best_cost: gbest_cost,
        iters_run,
        converged,
    }
}

// ── RSSI → position trilateration via PSO ──────────────────────────────────

/// Log-distance path-loss model parameters.
/// RSSI(d) = RSSI0 − 10·N·log10(d) + noise
#[derive(Debug, Clone, Copy)]
pub struct PathLoss {
    /// Reference RSSI at 1 m (dBm). Typical: -40 dBm.
    pub rssi_ref_1m: f64,
    /// Path-loss exponent. Free-space=2.0, indoor office=2.5–3.5, heavy walls=4.0.
    pub exponent: f64,
}

impl Default for PathLoss {
    fn default() -> Self {
        Self { rssi_ref_1m: -40.0, exponent: 2.5 }
    }
}

impl PathLoss {
    /// Expected RSSI in dBm at Euclidean distance `d_m` (metres).
    #[inline]
    pub fn predict_rssi(&self, d_m: f64) -> f64 {
        let d = d_m.max(0.25); // avoid log10(0); within 25 cm antenna coupling dominates
        self.rssi_ref_1m - 10.0 * self.exponent * d.log10()
    }

    /// Convert an observed RSSI to a distance estimate (metres).
    #[inline]
    pub fn estimate_distance(&self, rssi_dbm: f64) -> f64 {
        10f64.powf((self.rssi_ref_1m - rssi_dbm) / (10.0 * self.exponent))
    }
}

/// A single node observation for trilateration.
#[derive(Debug, Clone, Copy)]
pub struct RssiObservation {
    pub node_pos: [f64; 3],
    pub rssi_dbm: f64,
    /// Confidence weight in [0,1]. Low-RSSI nodes should be down-weighted.
    pub weight: f64,
}

/// Result of an RSSI PSO localization run.
#[derive(Debug, Clone)]
pub struct LocalizationResult {
    pub position: [f64; 3],
    /// RMS residual in dB across contributing nodes.
    pub residual_db: f64,
    /// Normalised confidence ∈ [0,1]. 1.0 ⇒ residual ≤ 1 dB.
    pub confidence: f64,
    pub iters: usize,
    pub method: &'static str,
}

/// Estimate target position from ≥3 RSSI observations using PSO.
///
/// Cost = Σ w_i · (RSSI_obs_i − RSSI_pred_i(x))².
///
/// Returns `None` if fewer than 2 nodes supplied (under-determined).
pub fn estimate_position_rssi_pso(
    observations: &[RssiObservation],
    bounds: &[(f64, f64); 3],
    model: PathLoss,
    cfg: &PsoConfig,
) -> Option<LocalizationResult> {
    if observations.len() < 2 { return None; }
    let obs: Vec<RssiObservation> = observations.iter().copied().collect();

    let objective = |x: &[f64]| -> f64 {
        let mut total = 0.0f64;
        let mut wsum = 0.0f64;
        for o in &obs {
            let dx = x[0] - o.node_pos[0];
            let dy = x[1] - o.node_pos[1];
            let dz = x[2] - o.node_pos[2];
            let d = (dx * dx + dy * dy + dz * dz).sqrt();
            let predicted = model.predict_rssi(d);
            let err = predicted - o.rssi_dbm;
            total += o.weight * err * err;
            wsum += o.weight;
        }
        if wsum > 0.0 { total / wsum } else { total }
    };

    let bs = [bounds[0], bounds[1], bounds[2]];
    let r = pso_minimize(&objective, &bs, cfg);
    let residual_db = r.best_cost.max(0.0).sqrt();
    // Confidence: 1 dB RMSE ⇒ 1.0, 10 dB ⇒ ~0.0 (smooth exp).
    let confidence = (-residual_db / 4.0).exp().clamp(0.0, 1.0);
    Some(LocalizationResult {
        position: [r.best_x[0], r.best_x[1], r.best_x[2]],
        residual_db,
        confidence,
        iters: r.iters_run,
        method: "pso-rssi",
    })
}

// ── deterministic PRNG (xorshift64*) ───────────────────────────────────────

struct XorShift64 { state: u64 }

impl XorShift64 {
    fn new(seed: u64) -> Self { Self { state: if seed == 0 { 0x9E3779B97F4A7C15 } else { seed } } }
    #[inline]
    fn next_u64(&mut self) -> u64 {
        let mut x = self.state;
        x ^= x << 13;
        x ^= x >> 7;
        x ^= x << 17;
        self.state = x;
        x.wrapping_mul(0x2545_F491_4F6C_DD1D)
    }
    /// Uniform in [0,1).
    #[inline]
    fn uniform(&mut self) -> f64 {
        // Use top 53 bits for IEEE-754 double mantissa.
        (self.next_u64() >> 11) as f64 / ((1u64 << 53) as f64)
    }
}

fn seed_from_time() -> u64 {
    use std::time::{SystemTime, UNIX_EPOCH};
    let ns = SystemTime::now().duration_since(UNIX_EPOCH)
        .map(|d| d.as_nanos() as u64).unwrap_or(0xA5A5_5A5A_1234_CAFE);
    ns ^ (PI.to_bits())
}

// ── tests ──────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    fn seeded() -> PsoConfig {
        PsoConfig { seed: Some(1729), swarm_size: 24, max_iters: 200, ..Default::default() }
    }

    #[test]
    fn sphere_minimum_at_origin() {
        let f = |x: &[f64]| x.iter().map(|v| v * v).sum::<f64>();
        let bounds = vec![(-5.0, 5.0); 4];
        let r = pso_minimize(&f, &bounds, &seeded());
        assert!(r.best_cost < 1e-3, "sphere cost too high: {}", r.best_cost);
    }

    #[test]
    fn rastrigin_gets_close() {
        // Highly multimodal — PSO should beat random search.
        let f = |x: &[f64]| {
            let a = 10.0;
            let n = x.len() as f64;
            a * n + x.iter().map(|v| v * v - a * (2.0 * PI * v).cos()).sum::<f64>()
        };
        let bounds = vec![(-5.12, 5.12); 3];
        let r = pso_minimize(&f, &bounds, &seeded());
        assert!(r.best_cost < 5.0, "rastrigin cost too high: {}", r.best_cost);
    }

    #[test]
    fn rssi_localization_recovers_3d_position() {
        // Ground truth target at (2.5, 1.5, 1.2) m. 4 nodes at room corners.
        let model = PathLoss { rssi_ref_1m: -40.0, exponent: 2.5 };
        let gt = [2.5, 1.5, 1.2];
        let nodes = [[0.0, 0.0, 2.5], [5.0, 0.0, 2.5], [5.0, 4.0, 2.5], [0.0, 4.0, 2.5]];
        let obs: Vec<RssiObservation> = nodes.iter().map(|p| {
            let d = ((gt[0]-p[0]).powi(2) + (gt[1]-p[1]).powi(2) + (gt[2]-p[2]).powi(2)).sqrt();
            RssiObservation { node_pos: *p, rssi_dbm: model.predict_rssi(d), weight: 1.0 }
        }).collect();
        let bounds = [(0.0, 8.0), (0.0, 6.0), (0.0, 3.0)];
        let cfg = seeded();
        let r = estimate_position_rssi_pso(&obs, &bounds, model, &cfg).unwrap();
        let err = ((r.position[0]-gt[0]).powi(2) + (r.position[1]-gt[1]).powi(2) + (r.position[2]-gt[2]).powi(2)).sqrt();
        assert!(err < 0.5, "recovered {:?} err={:.3} m residual={:.3} dB", r.position, err, r.residual_db);
        assert!(r.confidence > 0.5, "confidence too low: {}", r.confidence);
    }

    #[test]
    fn path_loss_roundtrip() {
        let m = PathLoss::default();
        let d = 3.2;
        let rssi = m.predict_rssi(d);
        let back = m.estimate_distance(rssi);
        assert!((back - d).abs() < 1e-9);
    }

    #[test]
    fn rssi_returns_none_with_insufficient_nodes() {
        let model = PathLoss::default();
        let obs = [RssiObservation { node_pos: [0.0;3], rssi_dbm: -55.0, weight: 1.0 }];
        let bounds = [(0.0, 5.0); 3];
        assert!(estimate_position_rssi_pso(&obs, &bounds, model, &seeded()).is_none());
    }

    #[test]
    fn deterministic_with_seed() {
        let f = |x: &[f64]| x.iter().map(|v| v * v).sum::<f64>();
        let bounds = vec![(-3.0, 3.0); 3];
        let a = pso_minimize(&f, &bounds, &seeded());
        let b = pso_minimize(&f, &bounds, &seeded());
        assert_eq!(a.best_x, b.best_x);
        assert_eq!(a.best_cost, b.best_cost);
    }
}
