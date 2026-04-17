/**
 * AEDI-S SDK Core — Sensing Platform Runtime
 *
 * Central SDK entry point that bootstraps capabilities detection, WebGL/WebGPU
 * rendering pipeline, GPU compute for CSI signal processing, and provides
 * a unified API surface for all sensing features.
 *
 * @version 1.0.0
 */

// ---- Capability Detection ------------------------------------------------

export class CapabilityProbe {
  constructor() {
    this._results = null;
  }

  /** Run full capability probe and cache results. */
  async detect() {
    if (this._results) return this._results;

    const canvas = document.createElement('canvas');
    const gl2 = canvas.getContext('webgl2');
    const gl1 = gl2 || canvas.getContext('webgl');

    const webgl2 = !!gl2;
    const webgl1 = !!gl1;

    let webgpu = false;
    let gpuAdapter = null;
    try {
      if (navigator.gpu) {
        gpuAdapter = await navigator.gpu.requestAdapter();
        webgpu = !!gpuAdapter;
      }
    } catch (_) { /* WebGPU not available */ }

    // WebGL extensions
    const ext = {};
    if (gl2 || gl1) {
      const ctx = gl2 || gl1;
      const available = ctx.getSupportedExtensions() || [];
      ext.floatTextures = available.includes('EXT_color_buffer_float') || available.includes('OES_texture_float');
      ext.floatBlend = available.includes('EXT_float_blend');
      ext.instancedArrays = webgl2 || available.includes('ANGLE_instanced_arrays');
      ext.drawBuffers = webgl2 || available.includes('WEBGL_draw_buffers');
      ext.depthTexture = webgl2 || available.includes('WEBGL_depth_texture');
      ext.anisotropic = available.includes('EXT_texture_filter_anisotropic');
      ext.compressedTexS3tc = available.includes('WEBGL_compressed_texture_s3tc');
      ext.compressedTexAstc = available.includes('WEBGL_compressed_texture_astc');
      ext.parallelShaderCompile = available.includes('KHR_parallel_shader_compile');
      ext.maxTextureSize = ctx.getParameter(ctx.MAX_TEXTURE_SIZE);
      ext.maxRenderbufferSize = ctx.getParameter(ctx.MAX_RENDERBUFFER_SIZE);
      ext.maxViewportDims = ctx.getParameter(ctx.MAX_VIEWPORT_DIMS);
      ext.vendor = ctx.getParameter(ctx.VENDOR);
      ext.renderer = ctx.getParameter(ctx.RENDERER);
      const dbgExt = ctx.getExtension('WEBGL_debug_renderer_info');
      if (dbgExt) {
        ext.unmaskedVendor = ctx.getParameter(dbgExt.UNMASKED_VENDOR_WEBGL);
        ext.unmaskedRenderer = ctx.getParameter(dbgExt.UNMASKED_RENDERER_WEBGL);
      }
      // Clean up
      const loseCtx = ctx.getExtension('WEBGL_lose_context');
      if (loseCtx) loseCtx.loseContext();
    }

    // WebGPU device limits
    let gpuLimits = null;
    if (gpuAdapter) {
      gpuLimits = {
        maxBufferSize: gpuAdapter.limits.maxBufferSize,
        maxStorageBufferBindingSize: gpuAdapter.limits.maxStorageBufferBindingSize,
        maxComputeWorkgroupSizeX: gpuAdapter.limits.maxComputeWorkgroupSizeX,
        maxComputeWorkgroupSizeY: gpuAdapter.limits.maxComputeWorkgroupSizeY,
        maxComputeWorkgroupSizeZ: gpuAdapter.limits.maxComputeWorkgroupSizeZ,
      };
    }

    // Platform info
    const platform = {
      userAgent: navigator.userAgent,
      hardwareConcurrency: navigator.hardwareConcurrency || 1,
      deviceMemory: navigator.deviceMemory || null,
      maxTouchPoints: navigator.maxTouchPoints || 0,
      devicePixelRatio: window.devicePixelRatio || 1,
      screenWidth: screen.width,
      screenHeight: screen.height,
      colorDepth: screen.colorDepth,
    };

    // Feature tier classification
    let tier = 'basic'; // WebGL1 only
    if (webgl2 && ext.instancedArrays && ext.floatTextures) tier = 'standard';
    if (webgl2 && ext.floatTextures && ext.drawBuffers && platform.hardwareConcurrency >= 4) tier = 'advanced';
    if (webgpu) tier = 'compute';

    this._results = {
      webgl1, webgl2, webgpu,
      extensions: ext,
      gpuLimits,
      platform,
      tier,
      timestamp: Date.now()
    };

    return this._results;
  }

  /** Quick read of cached results. */
  get results() { return this._results; }
}

// ---- SDK Event Bus -------------------------------------------------------

export class SDKEventBus {
  constructor() {
    this._handlers = new Map();
  }

  on(event, handler) {
    if (!this._handlers.has(event)) this._handlers.set(event, new Set());
    this._handlers.get(event).add(handler);
    return () => this._handlers.get(event)?.delete(handler);
  }

  emit(event, data) {
    const handlers = this._handlers.get(event);
    if (!handlers) return;
    for (const h of handlers) {
      try { h(data); } catch (e) { console.error(`[SDK] Event handler error (${event}):`, e); }
    }
  }

  off(event, handler) {
    this._handlers.get(event)?.delete(handler);
  }

  clear() {
    this._handlers.clear();
  }
}

// ---- Performance Monitor -------------------------------------------------

export class PerformanceMonitor {
  constructor() {
    this._frameTimes = new Float64Array(120);
    this._frameIndex = 0;
    this._frameCount = 0;
    this._lastFrameTime = 0;
    this._gpuTimers = [];
    this.metrics = {
      fps: 0,
      avgFrameTime: 0,
      minFrameTime: Infinity,
      maxFrameTime: 0,
      drawCalls: 0,
      triangles: 0,
      points: 0,
      memoryMB: 0,
      gpuTime: 0,
    };
  }

  beginFrame() {
    this._lastFrameTime = performance.now();
  }

  endFrame(renderer) {
    const now = performance.now();
    const dt = now - this._lastFrameTime;

    this._frameTimes[this._frameIndex] = dt;
    this._frameIndex = (this._frameIndex + 1) % this._frameTimes.length;
    this._frameCount++;

    // Compute metrics every 30 frames
    if (this._frameCount % 30 === 0) {
      const count = Math.min(this._frameCount, this._frameTimes.length);
      let sum = 0, min = Infinity, max = 0;
      for (let i = 0; i < count; i++) {
        const t = this._frameTimes[i];
        sum += t;
        if (t < min) min = t;
        if (t > max) max = t;
      }
      this.metrics.avgFrameTime = sum / count;
      this.metrics.minFrameTime = min;
      this.metrics.maxFrameTime = max;
      this.metrics.fps = 1000 / this.metrics.avgFrameTime;

      // Renderer info
      if (renderer?.info) {
        this.metrics.drawCalls = renderer.info.render.calls;
        this.metrics.triangles = renderer.info.render.triangles;
        this.metrics.points = renderer.info.render.points;
        this.metrics.memoryMB = (renderer.info.memory.geometries * 0.1 + renderer.info.memory.textures * 2).toFixed(1);
      }
    }

    return this.metrics;
  }
}

// ---- SDK Singleton -------------------------------------------------------

export class AEDISDK {
  constructor() {
    if (AEDISDK._instance) return AEDISDK._instance;
    AEDISDK._instance = this;

    this.version = '1.0.0';
    this.capabilities = new CapabilityProbe();
    this.events = new SDKEventBus();
    this.performance = new PerformanceMonitor();
    this._modules = new Map();
    this._initialized = false;
  }

  static getInstance() {
    if (!AEDISDK._instance) new AEDISDK();
    return AEDISDK._instance;
  }

  /** Initialize the SDK — call once before using any features. */
  async init(options = {}) {
    if (this._initialized) return this;

    console.log(`[AEDI-SDK] Initializing v${this.version}...`);

    // Detect capabilities
    const caps = await this.capabilities.detect();
    console.log(`[AEDI-SDK] GPU Tier: ${caps.tier} | WebGL2: ${caps.webgl2} | WebGPU: ${caps.webgpu}`);
    if (caps.extensions.unmaskedRenderer) {
      console.log(`[AEDI-SDK] GPU: ${caps.extensions.unmaskedRenderer}`);
    }

    this.events.emit('sdk:init', { version: this.version, capabilities: caps });
    this._initialized = true;

    return this;
  }

  /** Register a module into the SDK. */
  registerModule(name, module) {
    this._modules.set(name, module);
    this.events.emit('sdk:module:registered', { name });
  }

  /** Get a registered module. */
  getModule(name) {
    return this._modules.get(name);
  }

  /** Get the detected capability tier. */
  get tier() {
    return this.capabilities.results?.tier || 'unknown';
  }

  /** Dispose all modules and reset. */
  dispose() {
    for (const [name, mod] of this._modules) {
      if (typeof mod.dispose === 'function') mod.dispose();
    }
    this._modules.clear();
    this.events.clear();
    this._initialized = false;
    AEDISDK._instance = null;
  }
}

// Export singleton
export const sdk = AEDISDK.getInstance();
