/**
 * AEDI-S SDK — Sensing Pipeline Integration
 *
 * Connects the GPU compute, WebGL rendering, and data services
 * into a unified real-time sensing pipeline. This module is
 * the bridge between raw CSI data and visual output.
 */

import { sdk } from './core.js';
import { GPUCompute } from './gpu-compute.js';
import {
  InstancedCSIField,
  CSIParticleSystem,
  PostProcessingComposer,
  BloomPass,
  FXAAPass,
} from './webgl-pipeline.js';

// ---- Sensing Pipeline ---------------------------------------------------

export class SensingPipeline {
  constructor(options = {}) {
    this._compute = null;
    this._csiField = null;
    this._particles = null;
    this._postProcess = null;
    this._scene = null;
    this._renderer = null;
    this._camera = null;
    this._options = options;

    // Processing state
    this._csiBuffer = [];
    this._bufferSize = options.bufferSize || 64;
    this._lastFeatures = null;
    this._frameCount = 0;

    // Callbacks
    this._onFeatures = options.onFeatures || null;
  }

  /** Initialize all pipeline components. */
  async init(threeScene, renderer, camera) {
    this._scene = threeScene;
    this._renderer = renderer;
    this._camera = camera;

    // GPU Compute
    this._compute = new GPUCompute();
    await this._compute.init();
    sdk.registerModule('gpu-compute', this._compute);

    // Instanced CSI field (floor heatmap)
    this._csiField = new InstancedCSIField(threeScene, {
      gridX: this._options.gridX || 24,
      gridZ: this._options.gridZ || 18,
      cellSize: this._options.cellSize || 0.33,
    });
    sdk.registerModule('csi-field', this._csiField);

    // Motion particles
    this._particles = new CSIParticleSystem(threeScene, {
      maxParticles: this._options.maxParticles || 1500,
      emitRate: 40,
      lifetime: 2.5,
    });
    sdk.registerModule('particles', this._particles);

    // Post-processing (only on standard+ tier)
    const tier = sdk.tier;
    if (tier === 'standard' || tier === 'advanced' || tier === 'compute') {
      this._postProcess = new PostProcessingComposer(renderer, threeScene, camera);
      this._postProcess.addPass(new BloomPass({ threshold: 0.7, strength: 0.8, radius: 0.3 }));
      this._postProcess.addPass(new FXAAPass());
      sdk.registerModule('post-processing', this._postProcess);
    }

    console.log(`[Pipeline] Initialized | Compute: ${this._compute.backend} | Tier: ${tier} | PostFX: ${!!this._postProcess}`);
    return this;
  }

  /** Feed a raw sensing WebSocket message. */
  async processSensingData(message) {
    this._frameCount++;

    // Extract signal metrics
    const metrics = message.signal_metrics || {};
    const classification = message.classification || {};
    const vitalSigns = message.vital_signs || {};
    const nodes = message.nodes || [];

    // Buffer amplitude for FFT
    if (typeof metrics.variance === 'number') {
      this._csiBuffer.push(metrics.variance);
      if (this._csiBuffer.length > this._bufferSize) {
        this._csiBuffer.shift();
      }
    }

    // Compute motion features periodically
    if (this._csiBuffer.length >= 16 && this._frameCount % 5 === 0) {
      const features = await this._compute.extractMotionFeatures(
        new Float32Array(this._csiBuffer)
      );
      this._lastFeatures = features;
      if (this._onFeatures) this._onFeatures(features);
    }

    // Update CSI field heatmap
    this._updateCSIField(nodes, classification, metrics);

    // Update particle system
    const motionLevel = classification.motion_level || 0;
    this._particles.setMotionIntensity(motionLevel);

    if (nodes.length > 0) {
      // Center particles on detected activity
      const cx = nodes.reduce((s, n) => s + (n.x || 0), 0) / nodes.length;
      const cz = nodes.reduce((s, n) => s + (n.z || 0), 0) / nodes.length;
      this._particles.setEmitOrigin(cx, 0.5, cz);
    }

    // Burst particles on motion events
    if (motionLevel > 0.7) {
      this._particles.burst(0, 1.0, 0, Math.floor(motionLevel * 15));
    }

    sdk.events.emit('pipeline:data', {
      classification, vitalSigns, metrics, nodes,
      features: this._lastFeatures,
      frameCount: this._frameCount,
    });
  }

  /** Per-frame update — call from animation loop. */
  update(delta) {
    if (this._particles) this._particles.update(delta);
  }

  /** Render with post-processing if available. */
  render(delta) {
    if (this._postProcess) {
      this._postProcess.render(delta);
    }
    // If no post-processing, the caller handles normal rendering
    return !!this._postProcess;
  }

  /** Get last computed features. */
  get features() { return this._lastFeatures; }

  /** Get compute backend name. */
  get computeBackend() { return this._compute?.backend || 'none'; }

  _updateCSIField(nodes, classification, metrics) {
    const gridX = this._options.gridX || 24;
    const gridZ = this._options.gridZ || 18;
    const data = new Float32Array(gridX * gridZ);

    // Generate field from node positions and signal metrics
    const presence = classification.presence ? 1 : 0;
    const motion = classification.motion_level || 0;
    const baseLevel = (metrics.variance || 0) * 0.3;

    for (const node of nodes) {
      const nx = ((node.x || 0) + 4) / 8; // Normalize to 0-1 for 8m room
      const nz = ((node.z || 0) + 3) / 6; // Normalize to 0-1 for 6m room

      const gx = Math.floor(nx * gridX);
      const gz = Math.floor(nz * gridZ);
      const rssiNorm = Math.max(0, Math.min(1, 1 + (node.rssi || -80) / 80));

      // Radial falloff from node position
      for (let iz = 0; iz < gridZ; iz++) {
        for (let ix = 0; ix < gridX; ix++) {
          const dx = (ix - gx) / gridX;
          const dz = (iz - gz) / gridZ;
          const dist = Math.sqrt(dx * dx + dz * dz);
          const falloff = Math.exp(-dist * dist * 20);
          const idx = iz * gridX + ix;
          data[idx] = Math.min(1, data[idx] + falloff * rssiNorm * (0.3 + motion * 0.7));
        }
      }
    }

    // Global base level
    if (presence) {
      for (let i = 0; i < data.length; i++) {
        data[i] = Math.min(1, data[i] + baseLevel);
      }
    }

    this._csiField.update(data);
  }

  dispose() {
    if (this._compute) this._compute.dispose();
    if (this._csiField) this._csiField.dispose();
    if (this._particles) this._particles.dispose();
    if (this._postProcess) this._postProcess.dispose();
  }
}

// ---- Signal Analyzer (high-level API) -----------------------------------

export class SignalAnalyzer {
  constructor() {
    this._compute = new GPUCompute();
    this._history = [];
    this._maxHistory = 600; // 60 seconds at 10 Hz
  }

  async init() {
    await this._compute.init();
    return this;
  }

  /** Push a new CSI sample. */
  push(sample) {
    this._history.push({
      timestamp: Date.now(),
      amplitude: sample.amplitude || 0,
      phase: sample.phase || 0,
      rssi: sample.rssi || -80,
      variance: sample.variance || 0,
    });
    if (this._history.length > this._maxHistory) {
      this._history.shift();
    }
  }

  /** Get real-time spectrum of recent samples. */
  async getSpectrum() {
    if (this._history.length < 8) return null;
    const amps = new Float32Array(this._history.map(s => s.amplitude));
    return this._compute.fft(amps);
  }

  /** Get breathing rate estimate. */
  async getBreathingRate() {
    if (this._history.length < 30) return null;

    const recent = this._history.slice(-100);
    const signal = new Float32Array(recent.map(s => s.variance));
    const features = await this._compute.extractMotionFeatures(signal);

    // Breathing typically 12-20 BPM = 0.2-0.33 Hz
    return {
      rate: features.dominantFreq > 0.1 && features.dominantFreq < 0.5
        ? features.dominantFreq * 60
        : null,
      confidence: features.breathingBand / (features.motionBand + 0.001),
      raw: features,
    };
  }

  /** Get motion classification. */
  getMotionLevel() {
    if (this._history.length < 5) return 0;
    const recent = this._history.slice(-10);
    let sum = 0;
    for (const s of recent) sum += s.variance;
    return Math.min(1, sum / recent.length * 5);
  }

  /** Clear history buffer. */
  clear() {
    this._history = [];
  }

  dispose() {
    this._compute.dispose();
  }
}
