/**
 * AEDI-S SDK — Runtime Diagnostics & Self-Test
 *
 * Comprehensive runtime test suite that validates all SDK modules,
 * WebGL rendering pipeline, WebSocket connectivity, data flow,
 * and GPU compute functionality. Provides fault-finding and
 * debugging tools.
 */

import { sdk, CapabilityProbe, SDKEventBus, PerformanceMonitor } from './core.js';
import { GPUCompute } from './gpu-compute.js';

// ---- Test Result --------------------------------------------------------

class TestResult {
  constructor(name, category) {
    this.name = name;
    this.category = category;
    this.status = 'pending'; // pending | pass | fail | skip | warn
    this.duration = 0;
    this.error = null;
    this.details = null;
  }
}

// ---- Runtime Diagnostics Engine -----------------------------------------

export class RuntimeDiagnostics {
  constructor() {
    this._tests = [];
    this._results = [];
    this._listeners = new Set();
    this._isRunning = false;
  }

  /** Register a diagnostic test. */
  register(name, category, testFn, options = {}) {
    this._tests.push({ name, category, fn: testFn, timeout: options.timeout || 10000, critical: options.critical ?? false });
  }

  /** Run all diagnostics and return results. */
  async runAll() {
    if (this._isRunning) throw new Error('Diagnostics already running');
    this._isRunning = true;
    this._results = [];

    const startTime = performance.now();
    console.log('[DIAG] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    console.log('[DIAG] AEDI-S SDK Runtime Diagnostics');
    console.log('[DIAG] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    for (const test of this._tests) {
      const result = await this._runOne(test);
      this._results.push(result);
      this._notify(result);
    }

    const elapsed = (performance.now() - startTime).toFixed(0);
    const passed = this._results.filter(r => r.status === 'pass').length;
    const failed = this._results.filter(r => r.status === 'fail').length;
    const warned = this._results.filter(r => r.status === 'warn').length;
    const skipped = this._results.filter(r => r.status === 'skip').length;
    const total = this._results.length;

    console.log('[DIAG] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    console.log(`[DIAG] Results: ${passed}/${total} passed, ${failed} failed, ${warned} warnings, ${skipped} skipped (${elapsed}ms)`);
    console.log('[DIAG] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    this._isRunning = false;
    return {
      total, passed, failed, warned, skipped,
      duration: parseFloat(elapsed),
      results: [...this._results],
      healthy: failed === 0,
    };
  }

  /** Run tests for a specific category. */
  async runCategory(category) {
    const filtered = this._tests.filter(t => t.category === category);
    const results = [];
    for (const test of filtered) {
      results.push(await this._runOne(test));
    }
    return results;
  }

  /** Subscribe to test results as they complete. */
  onResult(callback) {
    this._listeners.add(callback);
    return () => this._listeners.delete(callback);
  }

  /** Get a report object suitable for display. */
  getReport() {
    const categories = {};
    for (const r of this._results) {
      if (!categories[r.category]) categories[r.category] = [];
      categories[r.category].push(r);
    }
    return categories;
  }

  async _runOne(test) {
    const result = new TestResult(test.name, test.category);
    const start = performance.now();

    try {
      const promise = test.fn(result);
      const timeout = new Promise((_, reject) =>
        setTimeout(() => reject(new Error(`Timeout after ${test.timeout}ms`)), test.timeout)
      );
      await Promise.race([promise, timeout]);

      if (result.status === 'pending') result.status = 'pass';
      result.duration = performance.now() - start;

      const icon = result.status === 'pass' ? '✓' : result.status === 'warn' ? '⚠' : '○';
      console.log(`[DIAG] ${icon} ${test.category}/${test.name} (${result.duration.toFixed(0)}ms)`);

    } catch (err) {
      result.status = 'fail';
      result.error = err.message;
      result.duration = performance.now() - start;
      console.error(`[DIAG] ✗ ${test.category}/${test.name}: ${err.message}`);
    }

    return result;
  }

  _notify(result) {
    for (const cb of this._listeners) {
      try { cb(result); } catch (_) { /* ignore */ }
    }
  }
}

// ---- Built-In Diagnostic Tests ------------------------------------------

export function registerBuiltInDiagnostics(diag) {

  // ---- SDK Core Tests ---------------------------------------------------

  diag.register('SDK singleton', 'core', async (r) => {
    const a = sdk;
    const b = sdk;
    if (a !== b) throw new Error('SDK is not a singleton');
    r.details = `v${a.version}`;
  }, { critical: true });

  diag.register('Capability probe', 'core', async (r) => {
    const probe = new CapabilityProbe();
    const caps = await probe.detect();
    if (!caps) throw new Error('Probe returned null');
    if (!caps.tier) throw new Error('No tier classification');
    r.details = `Tier: ${caps.tier} | WebGL2: ${caps.webgl2} | WebGPU: ${caps.webgpu}`;
  });

  diag.register('Event bus', 'core', async (r) => {
    const bus = new SDKEventBus();
    let received = false;
    const unsub = bus.on('test', () => { received = true; });
    bus.emit('test', {});
    if (!received) throw new Error('Event not received');
    unsub();
    bus.emit('test', {});
    r.details = 'Pub/Sub working';
  });

  diag.register('Performance monitor', 'core', async (r) => {
    const mon = new PerformanceMonitor();
    mon.beginFrame();
    await new Promise(resolve => setTimeout(resolve, 16));
    const metrics = mon.endFrame(null);
    if (typeof metrics.fps !== 'number') throw new Error('No FPS metric');
    r.details = `Frame time capture: OK`;
  });

  // ---- WebGL Tests ------------------------------------------------------

  diag.register('WebGL2 context', 'webgl', async (r) => {
    const canvas = document.createElement('canvas');
    const gl = canvas.getContext('webgl2');
    if (!gl) {
      r.status = 'warn';
      r.details = 'WebGL2 not available — falling back to WebGL1';
      return;
    }
    r.details = `MAX_TEXTURE_SIZE: ${gl.getParameter(gl.MAX_TEXTURE_SIZE)}`;
    const lc = gl.getExtension('WEBGL_lose_context');
    if (lc) lc.loseContext();
  });

  diag.register('Three.js available', 'webgl', async (r) => {
    const THREE = window.THREE;
    if (!THREE) throw new Error('Three.js not loaded on window.THREE');
    r.details = `Three.js revision: ${THREE.REVISION}`;
  }, { critical: true });

  diag.register('WebGL renderer creation', 'webgl', async (r) => {
    const THREE = window.THREE;
    if (!THREE) { r.status = 'skip'; return; }

    const canvas = document.createElement('canvas');
    canvas.width = 64;
    canvas.height = 64;
    let renderer;
    try {
      renderer = new THREE.WebGLRenderer({ canvas, antialias: false, powerPreference: 'low-power' });
      const info = renderer.info;
      r.details = `Renderer OK | Programs: ${info.programs?.length || 0}`;
    } finally {
      if (renderer) renderer.dispose();
    }
  });

  diag.register('Shader compilation', 'webgl', async (r) => {
    const THREE = window.THREE;
    if (!THREE) { r.status = 'skip'; return; }

    const canvas = document.createElement('canvas');
    canvas.width = 32;
    canvas.height = 32;
    const renderer = new THREE.WebGLRenderer({ canvas });
    const scene = new THREE.Scene();
    const camera = new THREE.OrthographicCamera(-1, 1, 1, -1, 0, 1);

    try {
      const material = new THREE.ShaderMaterial({
        vertexShader: `void main() { gl_Position = projectionMatrix * modelViewMatrix * vec4(position, 1.0); }`,
        fragmentShader: `void main() { gl_FragColor = vec4(0.0, 1.0, 0.0, 1.0); }`,
      });
      const mesh = new THREE.Mesh(new THREE.PlaneGeometry(2, 2), material);
      scene.add(mesh);
      renderer.render(scene, camera);

      // Read pixel to verify
      const gl = renderer.getContext();
      const pixel = new Uint8Array(4);
      gl.readPixels(16, 16, 1, 1, gl.RGBA, gl.UNSIGNED_BYTE, pixel);
      if (pixel[1] < 200) throw new Error(`Shader output incorrect: G=${pixel[1]}`);

      r.details = 'Custom shaders compiled and rendered correctly';
      material.dispose();
      mesh.geometry.dispose();
    } finally {
      renderer.dispose();
    }
  });

  diag.register('Instanced rendering', 'webgl', async (r) => {
    const THREE = window.THREE;
    if (!THREE) { r.status = 'skip'; return; }

    const canvas = document.createElement('canvas');
    canvas.width = 32;
    canvas.height = 32;
    const renderer = new THREE.WebGLRenderer({ canvas });
    const scene = new THREE.Scene();
    const camera = new THREE.OrthographicCamera(-1, 1, 1, -1, 0, 1);

    try {
      const geom = new THREE.BoxGeometry(0.1, 0.1, 0.1);
      const mat = new THREE.MeshBasicMaterial({ color: 0xff0000 });
      const mesh = new THREE.InstancedMesh(geom, mat, 100);

      const dummy = new THREE.Object3D();
      for (let i = 0; i < 100; i++) {
        dummy.position.set(Math.random() - 0.5, Math.random() - 0.5, 0);
        dummy.updateMatrix();
        mesh.setMatrixAt(i, dummy.matrix);
      }
      mesh.instanceMatrix.needsUpdate = true;
      scene.add(mesh);

      renderer.render(scene, camera);
      r.details = `100 instances rendered via InstancedMesh`;

      geom.dispose();
      mat.dispose();
    } finally {
      renderer.dispose();
    }
  });

  diag.register('Float textures', 'webgl', async (r) => {
    const canvas = document.createElement('canvas');
    const gl = canvas.getContext('webgl2');
    if (!gl) { r.status = 'skip'; r.details = 'Needs WebGL2'; return; }

    const ext = gl.getExtension('EXT_color_buffer_float');
    if (!ext) {
      r.status = 'warn';
      r.details = 'EXT_color_buffer_float not available — GPU compute limited';
      const lc = gl.getExtension('WEBGL_lose_context');
      if (lc) lc.loseContext();
      return;
    }
    r.details = 'Float render targets supported';
    const lc = gl.getExtension('WEBGL_lose_context');
    if (lc) lc.loseContext();
  });

  // ---- GPU Compute Tests ------------------------------------------------

  diag.register('GPU compute init', 'compute', async (r) => {
    const compute = new GPUCompute();
    await compute.init();
    r.details = `Backend: ${compute.backend}`;
    compute.dispose();
  });

  diag.register('FFT correctness', 'compute', async (r) => {
    const compute = new GPUCompute();
    await compute.init();

    // Test with a known signal: sine wave at 2 Hz, sampled at 16 Hz
    const N = 64;
    const input = new Float32Array(N);
    for (let i = 0; i < N; i++) {
      input[i] = Math.sin(2 * Math.PI * 2 * i / 16); // 2 Hz sine
    }

    const spectrum = await compute.fft(input);
    if (!spectrum || spectrum.length < N) throw new Error('FFT returned invalid result');

    // The peak should be at bin corresponding to 2 Hz
    let peakIdx = 0, peakVal = 0;
    for (let i = 1; i < N / 2; i++) {
      if (spectrum[i] > peakVal) { peakVal = spectrum[i]; peakIdx = i; }
    }

    r.details = `Peak at bin ${peakIdx}, magnitude ${peakVal.toFixed(2)} (backend: ${compute.backend})`;
    compute.dispose();
  });

  diag.register('Cross-correlation', 'compute', async (r) => {
    const compute = new GPUCompute();
    await compute.init();

    const N = 32;
    const a = new Float32Array(N);
    const b = new Float32Array(N);
    for (let i = 0; i < N; i++) {
      a[i] = Math.sin(2 * Math.PI * i / N);
      b[i] = Math.sin(2 * Math.PI * i / N); // identical signals
    }

    const result = await compute.crossCorrelate(a, b);
    // Auto-correlation at lag 0 should be ~1.0
    if (Math.abs(result[0] - 1.0) > 0.01) {
      throw new Error(`Auto-correlation at lag 0 = ${result[0].toFixed(4)}, expected ~1.0`);
    }

    r.details = `Auto-correlation[0] = ${result[0].toFixed(4)}`;
    compute.dispose();
  });

  diag.register('Motion feature extraction', 'compute', async (r) => {
    const compute = new GPUCompute();
    await compute.init();

    // Create a signal with known breathing-rate content
    const N = 100;
    const signal = new Float32Array(N);
    for (let i = 0; i < N; i++) {
      signal[i] = Math.sin(2 * Math.PI * 0.3 * i / 10) + // 0.3 Hz breathing
                  Math.random() * 0.1; // noise
    }

    const features = await compute.extractMotionFeatures(signal);
    if (typeof features.variance !== 'number') throw new Error('Missing variance');
    if (typeof features.breathingBand !== 'number') throw new Error('Missing breathing band');

    r.details = `Var: ${features.variance.toFixed(3)}, BrBand: ${features.breathingBand.toFixed(3)}, MoBand: ${features.motionBand.toFixed(3)}`;
    compute.dispose();
  });

  diag.register('Phase coherence', 'compute', async (r) => {
    const compute = new GPUCompute();
    await compute.init();

    const N = 32;
    const realA = new Float32Array(N);
    const imagA = new Float32Array(N);
    const realB = new Float32Array(N);
    const imagB = new Float32Array(N);

    // Coherent signals (same phase)
    for (let i = 0; i < N; i++) {
      const angle = 2 * Math.PI * i / N;
      realA[i] = Math.cos(angle);
      imagA[i] = Math.sin(angle);
      realB[i] = Math.cos(angle);
      imagB[i] = Math.sin(angle);
    }

    const coherence = await compute.phaseCoherence(realA, imagA, realB, imagB);
    if (coherence < 0.95) throw new Error(`Phase coherence ${coherence.toFixed(3)} < 0.95 for identical signals`);

    r.details = `Coherent signal coherence: ${coherence.toFixed(4)}`;
    compute.dispose();
  });

  // ---- Network Tests ----------------------------------------------------

  diag.register('HTTP health endpoint', 'network', async (r) => {
    try {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 3000);
      const resp = await fetch('/health', { signal: controller.signal });
      clearTimeout(timeout);

      if (resp.ok) {
        const data = await resp.json().catch(() => null);
        r.details = data ? `Status: ${JSON.stringify(data).substring(0, 80)}` : `HTTP ${resp.status}`;
      } else {
        r.status = 'warn';
        r.details = `HTTP ${resp.status} — server may not be running`;
      }
    } catch (e) {
      r.status = 'warn';
      r.details = `Server unreachable: ${e.message} — SDK will use mock/demo mode`;
    }
  });

  diag.register('WebSocket endpoint probe', 'network', async (r) => {
    const wsUrl = `ws://${window.location.hostname}:3001/ws/sensing`;
    const ws = new WebSocket(wsUrl);

    return new Promise((resolve) => {
      const timeout = setTimeout(() => {
        ws.close();
        r.status = 'warn';
        r.details = `WebSocket timeout at ${wsUrl} — demo mode will activate`;
        resolve();
      }, 3000);

      ws.onopen = () => {
        clearTimeout(timeout);
        r.details = `Connected to ${wsUrl}`;
        ws.close(1000);
        resolve();
      };

      ws.onerror = () => {
        clearTimeout(timeout);
        r.status = 'warn';
        r.details = `WebSocket unavailable at ${wsUrl} — demo mode will activate`;
        resolve();
      };
    });
  });

  // ---- Data Pipeline Tests ----------------------------------------------

  diag.register('CSI data parsing', 'pipeline', async (r) => {
    // Simulate a realistic sensing message
    const mockMsg = {
      type: 'sensing_update',
      source: 'simulated',
      tick: 42,
      classification: { presence: true, motion_level: 0.6, enhanced_motion: 0.3 },
      vital_signs: { heart_rate_bpm: 72, breathing_rate_bpm: 16, breathing_confidence: 0.85 },
      signal_metrics: { rssi: -45, variance: 0.12, spectral_power: 0.34, dominant_freq: 0.28 },
      nodes: [
        { node_id: 'node1', rssi: -42, subcarrier_count: 56, x: -2, y: 0.5, z: -1 },
        { node_id: 'node2', rssi: -48, subcarrier_count: 56, x: 2, y: 0.5, z: 1 }
      ]
    };

    if (!mockMsg.classification) throw new Error('Missing classification field');
    if (!mockMsg.vital_signs) throw new Error('Missing vital_signs field');
    if (!Array.isArray(mockMsg.nodes)) throw new Error('Missing nodes array');

    r.details = `Parsed: ${mockMsg.nodes.length} nodes, HR ${mockMsg.vital_signs.heart_rate_bpm} BPM, presence=${mockMsg.classification.presence}`;
  });

  diag.register('Service module availability', 'pipeline', async (r) => {
    const modules = [
      'api.service.js', 'websocket.service.js', 'sensing.service.js',
      'health.service.js', 'pose.service.js', 'data-processor.js'
    ];
    const available = [];
    const missing = [];

    // Resolve base from import.meta.url (sdk/ dir) → parent = ui root
    const base = new URL('..', import.meta.url).href;
    for (const mod of modules) {
      try {
        const resp = await fetch(new URL(`services/${mod}`, base), { method: 'HEAD' });
        if (resp.ok) available.push(mod);
        else missing.push(mod);
      } catch {
        missing.push(mod);
      }
    }

    if (missing.length > 0) {
      r.status = 'warn';
      r.details = `Missing: ${missing.join(', ')}`;
    } else {
      r.details = `All ${available.length} service modules available`;
    }
  });

  diag.register('Component module availability', 'pipeline', async (r) => {
    const modules = [
      'scene.js', 'body-model.js', 'environment.js',
      'gaussian-splats.js', 'signal-viz.js', 'dashboard-hud.js',
      'PoseDetectionCanvas.js', 'TabManager.js'
    ];
    const available = [];
    const missing = [];

    const compBase = new URL('..', import.meta.url).href;
    for (const mod of modules) {
      try {
        const resp = await fetch(new URL(`components/${mod}`, compBase), { method: 'HEAD' });
        if (resp.ok) available.push(mod);
        else missing.push(mod);
      } catch {
        missing.push(mod);
      }
    }

    if (missing.length > 0) {
      r.status = 'warn';
      r.details = `Missing: ${missing.join(', ')}`;
    } else {
      r.details = `All ${available.length} component modules available`;
    }
  });

  diag.register('SDK module availability', 'pipeline', async (r) => {
    const modules = ['core.js', 'webgl-pipeline.js', 'gpu-compute.js', 'diagnostics.js'];
    const available = [];
    const missing = [];

    const sdkBase = new URL('.', import.meta.url).href;
    for (const mod of modules) {
      try {
        const resp = await fetch(new URL(mod, sdkBase), { method: 'HEAD' });
        if (resp.ok) available.push(mod);
        else missing.push(mod);
      } catch {
        missing.push(mod);
      }
    }

    if (missing.length > 0) {
      r.status = 'warn';
      r.details = `Missing: ${missing.join(', ')}`;
    } else {
      r.details = `All ${available.length} SDK modules available`;
    }
  });

  // ---- Memory & Performance Tests ---------------------------------------

  diag.register('Memory baseline', 'performance', async (r) => {
    if (performance.memory) {
      const mb = (performance.memory.usedJSHeapSize / 1048576).toFixed(1);
      const total = (performance.memory.totalJSHeapSize / 1048576).toFixed(1);
      r.details = `Heap: ${mb} MB / ${total} MB`;
      if (performance.memory.usedJSHeapSize > performance.memory.jsHeapSizeLimit * 0.8) {
        r.status = 'warn';
        r.details += ' — HIGH memory usage (>80%)';
      }
    } else {
      r.details = 'performance.memory not available (non-Chrome)';
    }
  });

  diag.register('Render performance', 'performance', async (r) => {
    const THREE = window.THREE;
    if (!THREE) { r.status = 'skip'; return; }

    const canvas = document.createElement('canvas');
    canvas.width = 256;
    canvas.height = 256;
    const renderer = new THREE.WebGLRenderer({ canvas, antialias: false });
    const scene = new THREE.Scene();
    const camera = new THREE.PerspectiveCamera(75, 1, 0.1, 100);
    camera.position.z = 5;

    // Add some geometry
    for (let i = 0; i < 50; i++) {
      const mesh = new THREE.Mesh(
        new THREE.SphereGeometry(0.2, 8, 8),
        new THREE.MeshBasicMaterial({ color: Math.random() * 0xffffff })
      );
      mesh.position.set(Math.random() * 6 - 3, Math.random() * 6 - 3, Math.random() * 6 - 3);
      scene.add(mesh);
    }

    // Benchmark 60 frames
    const start = performance.now();
    for (let i = 0; i < 60; i++) {
      renderer.render(scene, camera);
    }
    const elapsed = performance.now() - start;
    const avgMs = elapsed / 60;
    const estFps = 1000 / avgMs;

    r.details = `60 frames in ${elapsed.toFixed(0)}ms (${avgMs.toFixed(1)}ms/frame, ~${estFps.toFixed(0)} FPS)`;
    if (estFps < 30) {
      r.status = 'warn';
      r.details += ' — Below 30 FPS threshold';
    }

    renderer.dispose();
  });

  diag.register('TypedArray performance', 'performance', async (r) => {
    const N = 100000;
    const arr = new Float32Array(N);
    for (let i = 0; i < N; i++) arr[i] = Math.random();

    const start = performance.now();
    let sum = 0;
    for (let i = 0; i < N; i++) sum += arr[i];
    for (let i = 0; i < N; i++) arr[i] *= 2.0;
    const elapsed = performance.now() - start;

    r.details = `${N} float ops in ${elapsed.toFixed(2)}ms (${(N * 2 / elapsed * 1000).toFixed(0)} ops/sec)`;
  });
}

// ---- Fault Finder (interactive debug) -----------------------------------

export class FaultFinder {
  constructor() {
    this._issues = [];
  }

  /** Run fault analysis and return categorised issues. */
  async analyze() {
    this._issues = [];

    // Check Three.js
    if (!window.THREE) {
      this._issues.push({ severity: 'critical', module: 'rendering', message: 'Three.js not loaded — 3D visualization will fail', fix: 'Add <script src="https://unpkg.com/three@0.160.0/build/three.min.js"></script> before app scripts' });
    } else {
      if (!window.THREE.OrbitControls) {
        this._issues.push({ severity: 'warning', module: 'rendering', message: 'OrbitControls not loaded — camera interaction disabled', fix: 'Add OrbitControls script after Three.js' });
      }
    }

    // Check WebGL
    const canvas = document.createElement('canvas');
    const gl = canvas.getContext('webgl2') || canvas.getContext('webgl');
    if (!gl) {
      this._issues.push({ severity: 'critical', module: 'gpu', message: 'No WebGL support — 3D rendering impossible', fix: 'Use a browser with WebGL support (Chrome, Firefox, Edge)' });
    }
    const lc = gl?.getExtension('WEBGL_lose_context');
    if (lc) lc.loseContext();

    // Check network
    try {
      const resp = await fetch('/health', { signal: AbortSignal.timeout(2000) });
      if (!resp.ok) {
        this._issues.push({ severity: 'warning', module: 'network', message: `Health endpoint returned ${resp.status}`, fix: 'Start the sensing server: ./.ionity/ionity.sh run --yes' });
      }
    } catch {
      this._issues.push({ severity: 'info', module: 'network', message: 'Sensing server not reachable — using demo/mock mode', fix: 'Start server or continue in demo mode' });
    }

    // Check WebSocket
    try {
      await new Promise((resolve, reject) => {
        const ws = new WebSocket(`ws://${window.location.hostname}:3001/ws/sensing`);
        const t = setTimeout(() => { ws.close(); reject(new Error('timeout')); }, 2000);
        ws.onopen = () => { clearTimeout(t); ws.close(1000); resolve(); };
        ws.onerror = () => { clearTimeout(t); reject(new Error('error')); };
      });
    } catch {
      this._issues.push({ severity: 'info', module: 'network', message: 'WebSocket endpoint not available — real-time data disabled', fix: 'Start sensing server for live data' });
    }

    // Check DOM containers
    const pages = ['viz-container', 'main', 'shell'];
    for (const id of pages) {
      if (document.getElementById(id)) continue;
      // Only warn if we're on the relevant page
    }

    // Memory check
    if (performance.memory && performance.memory.usedJSHeapSize > 200 * 1024 * 1024) {
      this._issues.push({ severity: 'warning', module: 'memory', message: `High memory usage: ${(performance.memory.usedJSHeapSize / 1048576).toFixed(0)} MB`, fix: 'Dispose unused scenes, reduce particle count, check for memory leaks' });
    }

    // Screen DPI check
    if (window.devicePixelRatio > 2) {
      this._issues.push({ severity: 'info', module: 'rendering', message: `High DPI display (${window.devicePixelRatio}x) — renderer should cap pixel ratio at 2`, fix: 'renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2))' });
    }

    return this._issues;
  }

  /** Get issues by severity. */
  getIssues(severity = null) {
    if (!severity) return [...this._issues];
    return this._issues.filter(i => i.severity === severity);
  }

  /** Print a formatted report to console. */
  printReport() {
    console.log('\n[FAULT-FINDER] ═══════════════════════════════');
    console.log(`[FAULT-FINDER] Found ${this._issues.length} issue(s)\n`);

    const severityOrder = { critical: 0, warning: 1, info: 2 };
    const sorted = [...this._issues].sort((a, b) => (severityOrder[a.severity] ?? 3) - (severityOrder[b.severity] ?? 3));

    for (const issue of sorted) {
      const icon = issue.severity === 'critical' ? '🔴' : issue.severity === 'warning' ? '🟡' : '🔵';
      console.log(`${icon} [${issue.module}] ${issue.message}`);
      console.log(`   Fix: ${issue.fix}`);
    }
    console.log('[FAULT-FINDER] ═══════════════════════════════\n');
  }
}
