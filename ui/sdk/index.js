/**
 * AEDI-S SDK — Unified Module Index
 *
 * Single import point for all SDK capabilities:
 *   import { sdk, SensingPipeline, RuntimeDiagnostics, ... } from './sdk/index.js';
 *
 * @version 1.0.0
 * @license MIT OR Apache-2.0
 */

// Core SDK
export { AEDISDK, sdk, CapabilityProbe, SDKEventBus, PerformanceMonitor } from './core.js';

// WebGL Rendering Pipeline
export {
  BloomPass,
  FXAAPass,
  PostProcessingComposer,
  InstancedCSIField,
  CSIParticleSystem,
  VolumeCSIRenderer,
} from './webgl-pipeline.js';

// GPU Compute
export { GPUCompute } from './gpu-compute.js';

// Sensing Pipeline
export { SensingPipeline, SignalAnalyzer } from './sensing-pipeline.js';

// Diagnostics & Debugging
export {
  RuntimeDiagnostics,
  registerBuiltInDiagnostics,
  FaultFinder,
} from './diagnostics.js';
