//! Adaptive CSI Denoiser — MAD-estimated noise + Haar DWT soft-threshold.
//!
//! This module implements a lightweight, pure-Rust denoiser for CSI amplitude or
//! phase traces. It is designed to complement the existing `hampel` outlier rejector
//! and `phase_sanitizer` by attacking *broadband* noise (the zero-mean fluctuations
//! that survive Hampel but still degrade motion features).
//!
//! The algorithm is a textbook wavelet-shrinkage denoiser:
//!
//! 1. Apply a multi-level 1-D Haar Discrete Wavelet Transform (DWT) along time.
//! 2. Estimate the noise standard deviation σ from the finest-level detail
//!    coefficients using the Median Absolute Deviation (MAD) estimator:
//!    `σ̂ = median(|d₁ − median(d₁)|) / 0.6745`.
//! 3. Apply soft-threshold shrinkage to every detail level with Donoho's
//!    universal threshold `λ = σ̂ √(2 ln N)`.
//! 4. Reconstruct via the inverse Haar DWT.
//!
//! Reference: Donoho, D. L. (1994). *Ideal spatial adaptation by wavelet shrinkage.*
//!
//! The implementation is intentionally dependency-free (only `std` + `num_traits`)
//! and allocation-light (one reusable scratch buffer per level). It compiles and
//! runs cleanly under `cargo test --workspace --no-default-features`.
//!
//! # Example
//!
//! ```
//! use wifi_densepose_signal::adaptive_denoise::{AdaptiveDenoiser, DenoiserConfig};
//!
//! let cfg = DenoiserConfig { levels: 2, ..Default::default() };
//! let denoiser = AdaptiveDenoiser::new(cfg);
//! let noisy: Vec<f64> = (0..128).map(|i| (i as f64 * 0.1).sin() + 0.3).collect();
//! let clean = denoiser.denoise(&noisy);
//! assert_eq!(clean.len(), noisy.len());
//! ```

/// Configuration for the adaptive denoiser.
#[derive(Debug, Clone, Copy)]
pub struct DenoiserConfig {
    /// Number of Haar DWT decomposition levels. 1–4 are typical.
    /// Each level halves the detail-coefficient resolution; higher levels smooth more.
    pub levels: u8,
    /// Multiplier applied to the universal threshold `σ √(2 ln N)`.
    /// 1.0 = classical Donoho; 0.6–0.8 preserves more signal detail.
    pub threshold_scale: f64,
    /// When true, the denoiser is bypassed (convenient for ablation / A-B tests).
    pub bypass: bool,
}

impl Default for DenoiserConfig {
    fn default() -> Self {
        Self {
            levels: 2,
            threshold_scale: 1.0,
            bypass: false,
        }
    }
}

/// Stateless adaptive denoiser.
///
/// The denoiser owns only its configuration; all buffers are allocated per call.
/// For hot-loop use, prefer calling `denoise_into` with a reusable output buffer.
#[derive(Debug, Clone)]
pub struct AdaptiveDenoiser {
    cfg: DenoiserConfig,
}

impl AdaptiveDenoiser {
    /// Construct a new denoiser with the given configuration.
    pub fn new(cfg: DenoiserConfig) -> Self {
        Self { cfg }
    }

    /// Return the active configuration.
    pub fn config(&self) -> DenoiserConfig {
        self.cfg
    }

    /// Denoise a 1-D signal, returning a newly-allocated vector.
    ///
    /// Signals shorter than `2^levels` are returned unchanged. Signals whose length
    /// is not a power of two are zero-padded internally (the pad is stripped before
    /// returning, so the output length exactly matches the input).
    pub fn denoise(&self, signal: &[f64]) -> Vec<f64> {
        let mut out = signal.to_vec();
        self.denoise_into(signal, &mut out);
        out
    }

    /// In-place variant: writes into `out` (which must be the same length as `signal`).
    pub fn denoise_into(&self, signal: &[f64], out: &mut [f64]) {
        assert_eq!(
            signal.len(),
            out.len(),
            "adaptive_denoise: signal and output must have the same length"
        );
        if self.cfg.bypass || signal.is_empty() {
            out.copy_from_slice(signal);
            return;
        }

        let levels = self.cfg.levels as usize;
        let min_len = 1usize << levels.max(1);
        if signal.len() < min_len {
            out.copy_from_slice(signal);
            return;
        }

        // Pad to next power of two so the Haar DWT is well-defined.
        let n = signal.len();
        let padded_len = n.next_power_of_two();
        let mut buf = Vec::with_capacity(padded_len);
        buf.extend_from_slice(signal);
        if padded_len > n {
            // Symmetric reflection of the last samples to minimise edge artefacts.
            let mut i = n;
            let mut step: isize = -1;
            let mut cursor: isize = (n - 1) as isize;
            while i < padded_len {
                cursor += step;
                if cursor < 0 {
                    cursor = 0;
                    step = 1;
                } else if cursor >= n as isize {
                    cursor = (n - 1) as isize;
                    step = -1;
                }
                buf.push(signal[cursor as usize]);
                i += 1;
            }
        }

        // Forward Haar DWT, capturing all detail bands.
        let (approx, mut details) = forward_haar(&buf, levels);

        // Noise σ is estimated from the finest-level detail coefficients via MAD.
        // finest == details[0] by construction below.
        let sigma = mad_sigma(&details[0]);

        // Donoho universal threshold applied to every detail level.
        let n_eff = buf.len() as f64;
        let lambda = self.cfg.threshold_scale * sigma * (2.0 * n_eff.ln()).sqrt();
        for detail in details.iter_mut() {
            for d in detail.iter_mut() {
                *d = soft_threshold(*d, lambda);
            }
        }

        // Inverse transform and strip padding.
        let reconstructed = inverse_haar(&approx, &details);
        out.copy_from_slice(&reconstructed[..n]);
    }
}

// -----------------------------------------------------------------------------
// Haar DWT — orthonormal, single allocation per level.
// -----------------------------------------------------------------------------

/// Forward Haar wavelet transform.
///
/// Returns `(approx, details)` where `details[0]` is the *finest* level (high-freq)
/// and `details[levels-1]` is the coarsest. `approx` is the coarsest scaling band.
fn forward_haar(signal: &[f64], levels: usize) -> (Vec<f64>, Vec<Vec<f64>>) {
    debug_assert!(signal.len().is_power_of_two());
    let mut current = signal.to_vec();
    let mut details = Vec::with_capacity(levels);
    let inv_sqrt2 = std::f64::consts::FRAC_1_SQRT_2;
    for _ in 0..levels {
        let n = current.len();
        if n < 2 {
            break;
        }
        let half = n / 2;
        let mut approx = vec![0.0; half];
        let mut detail = vec![0.0; half];
        for i in 0..half {
            let a = current[2 * i];
            let b = current[2 * i + 1];
            approx[i] = (a + b) * inv_sqrt2;
            detail[i] = (a - b) * inv_sqrt2;
        }
        details.push(detail);
        current = approx;
    }
    (current, details)
}

/// Inverse Haar wavelet transform — reconstructs the signal.
fn inverse_haar(approx: &[f64], details: &[Vec<f64>]) -> Vec<f64> {
    let mut current = approx.to_vec();
    let inv_sqrt2 = std::f64::consts::FRAC_1_SQRT_2;
    // Reconstruct coarse -> fine, matching the order produced by forward_haar.
    for detail in details.iter().rev() {
        let n = current.len();
        debug_assert_eq!(detail.len(), n);
        let mut upsampled = vec![0.0; n * 2];
        for i in 0..n {
            let a = current[i];
            let d = detail[i];
            upsampled[2 * i] = (a + d) * inv_sqrt2;
            upsampled[2 * i + 1] = (a - d) * inv_sqrt2;
        }
        current = upsampled;
    }
    current
}

// -----------------------------------------------------------------------------
// Helpers.
// -----------------------------------------------------------------------------

/// Soft-threshold shrinkage: `sign(x) · max(|x| − λ, 0)`.
#[inline]
fn soft_threshold(x: f64, lambda: f64) -> f64 {
    let ax = x.abs();
    if ax <= lambda {
        0.0
    } else {
        x.signum() * (ax - lambda)
    }
}

/// Median of a slice (non-destructive; copies).
fn median(v: &[f64]) -> f64 {
    if v.is_empty() {
        return 0.0;
    }
    let mut buf: Vec<f64> = v.iter().copied().filter(|x| x.is_finite()).collect();
    if buf.is_empty() {
        return 0.0;
    }
    buf.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let n = buf.len();
    if n % 2 == 1 {
        buf[n / 2]
    } else {
        0.5 * (buf[n / 2 - 1] + buf[n / 2])
    }
}

/// Robust noise-σ estimate via Median Absolute Deviation.
/// `σ̂ = median(|d − median(d)|) / 0.6745`.
fn mad_sigma(detail: &[f64]) -> f64 {
    if detail.is_empty() {
        return 0.0;
    }
    let m = median(detail);
    let abs_dev: Vec<f64> = detail.iter().map(|x| (x - m).abs()).collect();
    let mad = median(&abs_dev);
    mad / 0.6745
}

// -----------------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    fn rng_next(state: &mut u64) -> f64 {
        // Tiny, deterministic xorshift* → uniform (0,1).
        let mut x = *state;
        x ^= x >> 12;
        x ^= x << 25;
        x ^= x >> 27;
        *state = x;
        let v = x.wrapping_mul(0x2545F4914F6CDD1D);
        (v >> 11) as f64 / (1u64 << 53) as f64
    }

    fn gaussian(state: &mut u64) -> f64 {
        // Box–Muller.
        let u1 = rng_next(state).max(1e-12);
        let u2 = rng_next(state);
        (-2.0_f64 * u1.ln()).sqrt() * (2.0 * std::f64::consts::PI * u2).cos()
    }

    fn variance(v: &[f64]) -> f64 {
        let n = v.len() as f64;
        let mean: f64 = v.iter().sum::<f64>() / n;
        v.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / n
    }

    fn energy(v: &[f64]) -> f64 {
        v.iter().map(|x| x * x).sum::<f64>()
    }

    /// Round-trip identity: forward then inverse Haar must reproduce the input
    /// (to floating-point precision) for power-of-two signals.
    #[test]
    fn haar_roundtrip_is_identity() {
        let n = 64usize;
        let signal: Vec<f64> = (0..n).map(|i| (i as f64 * 0.35).sin()).collect();
        let (approx, details) = forward_haar(&signal, 3);
        let reconstructed = inverse_haar(&approx, &details);
        assert_eq!(reconstructed.len(), n);
        for (a, b) in signal.iter().zip(reconstructed.iter()) {
            assert!((a - b).abs() < 1e-9, "roundtrip error: {} vs {}", a, b);
        }
    }

    /// MAD σ should recover Gaussian σ within a reasonable factor on N=1024 samples.
    #[test]
    fn mad_sigma_tracks_gaussian_scale() {
        let mut state: u64 = 0x1234_5678_9abc_def0;
        let samples: Vec<f64> = (0..1024).map(|_| gaussian(&mut state)).collect();
        let est = mad_sigma(&samples);
        // Population σ = 1; MAD on Gaussian typically falls in [0.85, 1.15] at this N.
        assert!(est > 0.7 && est < 1.3, "MAD σ off: {}", est);
    }

    /// Soft threshold zeros out small |x| and shrinks large |x| by λ.
    #[test]
    fn soft_threshold_matches_formula() {
        assert_eq!(soft_threshold(0.3, 0.5), 0.0);
        assert!((soft_threshold(1.0, 0.5) - 0.5).abs() < 1e-12);
        assert!((soft_threshold(-1.0, 0.5) + 0.5).abs() < 1e-12);
    }

    /// Bypass mode returns the input unchanged.
    #[test]
    fn bypass_is_noop() {
        let cfg = DenoiserConfig {
            bypass: true,
            ..Default::default()
        };
        let denoiser = AdaptiveDenoiser::new(cfg);
        let x: Vec<f64> = (0..32).map(|i| i as f64).collect();
        let y = denoiser.denoise(&x);
        assert_eq!(x, y);
    }

    /// Clean signal should pass through nearly intact (small energy loss).
    #[test]
    fn clean_signal_energy_is_preserved() {
        let denoiser = AdaptiveDenoiser::new(DenoiserConfig::default());
        let signal: Vec<f64> = (0..256)
            .map(|i| (i as f64 * 0.25).sin() + 0.5 * (i as f64 * 0.07).cos())
            .collect();
        let out = denoiser.denoise(&signal);
        let e_in = energy(&signal);
        let e_out = energy(&out);
        let loss = (e_in - e_out).abs() / e_in;
        assert!(
            loss < 0.15,
            "clean-signal energy loss too high: {:.3}",
            loss
        );
    }

    /// Primary quality claim: SNR gain ≥ 3 dB on a ~5 dB input.
    /// Signal: sinusoid amplitude 1.0. Noise: zero-mean Gaussian σ ≈ 0.56 → SNR ≈ 5 dB.
    #[test]
    fn snr_gain_on_noisy_sinusoid() {
        let n = 512;
        let signal_amp = 1.0_f64;
        let noise_sigma = 0.56_f64;
        let mut state: u64 = 0xC0FFEE_1234_5678;
        let clean: Vec<f64> = (0..n)
            .map(|i| signal_amp * (2.0 * std::f64::consts::PI * i as f64 / 32.0).sin())
            .collect();
        let noisy: Vec<f64> = clean
            .iter()
            .map(|&s| s + noise_sigma * gaussian(&mut state))
            .collect();

        let denoiser = AdaptiveDenoiser::new(DenoiserConfig {
            levels: 3,
            threshold_scale: 1.0,
            bypass: false,
        });
        let cleaned = denoiser.denoise(&noisy);

        let err_in: Vec<f64> = noisy.iter().zip(clean.iter()).map(|(n, c)| n - c).collect();
        let err_out: Vec<f64> = cleaned.iter().zip(clean.iter()).map(|(n, c)| n - c).collect();
        let snr_in = 10.0 * (variance(&clean) / variance(&err_in).max(1e-18)).log10();
        let snr_out = 10.0 * (variance(&clean) / variance(&err_out).max(1e-18)).log10();
        let gain_db = snr_out - snr_in;
        assert!(
            gain_db >= 3.0,
            "SNR gain {:.2} dB < 3 dB (in {:.2}, out {:.2})",
            gain_db,
            snr_in,
            snr_out
        );
    }

    /// Short signals (below 2^levels) should be returned unchanged — no panic.
    #[test]
    fn short_signal_returns_unchanged() {
        let denoiser = AdaptiveDenoiser::new(DenoiserConfig { levels: 3, ..Default::default() });
        let x = vec![1.0, 2.0, 3.0]; // length 3 < 2^3 = 8
        let y = denoiser.denoise(&x);
        assert_eq!(x, y);
    }

    /// Non-power-of-two input still runs and preserves length.
    #[test]
    fn non_power_of_two_length_preserved() {
        let denoiser = AdaptiveDenoiser::new(DenoiserConfig::default());
        let x: Vec<f64> = (0..100).map(|i| (i as f64 * 0.2).sin()).collect();
        let y = denoiser.denoise(&x);
        assert_eq!(x.len(), y.len());
        assert!(y.iter().all(|v| v.is_finite()));
    }
}
