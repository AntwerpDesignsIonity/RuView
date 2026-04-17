/**
 * AEDI-S SDK — WebGL2 Advanced Rendering Pipeline
 *
 * Post-processing stack, instanced rendering, particle systems,
 * volume rendering, and screen-space effects. Builds on Three.js
 * with feature-gated capabilities from the SDK core probe.
 */

const getThree = () => window.THREE;

// ---- Post-Processing Pass Base -------------------------------------------

class RenderPass {
  constructor(name) {
    this.name = name;
    this.enabled = true;
    this._uniforms = {};
  }

  setUniform(key, value) { this._uniforms[key] = value; }
  render(renderer, readBuffer, writeBuffer, delta) { /* override */ }
  dispose() { /* override */ }
}

// ---- Bloom Pass (screen-space glow) --------------------------------------

export class BloomPass extends RenderPass {
  constructor(options = {}) {
    super('bloom');
    const THREE = getThree();

    this.threshold = options.threshold ?? 0.6;
    this.strength = options.strength ?? 1.2;
    this.radius = options.radius ?? 0.4;

    // Luminance extraction shader
    this._luminanceMaterial = new THREE.ShaderMaterial({
      uniforms: {
        tDiffuse: { value: null },
        luminosityThreshold: { value: this.threshold },
        smoothWidth: { value: 0.01 },
      },
      vertexShader: `
        varying vec2 vUv;
        void main() {
          vUv = uv;
          gl_Position = projectionMatrix * modelViewMatrix * vec4(position, 1.0);
        }
      `,
      fragmentShader: `
        uniform sampler2D tDiffuse;
        uniform float luminosityThreshold;
        uniform float smoothWidth;
        varying vec2 vUv;

        void main() {
          vec4 texel = texture2D(tDiffuse, vUv);
          float luma = dot(texel.rgb, vec3(0.299, 0.587, 0.114));
          float alpha = smoothstep(luminosityThreshold, luminosityThreshold + smoothWidth, luma);
          gl_FragColor = mix(vec4(0.0), texel, alpha);
        }
      `,
    });

    // Gaussian blur shader (separable)
    this._blurMaterial = new THREE.ShaderMaterial({
      uniforms: {
        tDiffuse: { value: null },
        direction: { value: new THREE.Vector2(1.0, 0.0) },
        resolution: { value: new THREE.Vector2(1, 1) },
      },
      vertexShader: `
        varying vec2 vUv;
        void main() {
          vUv = uv;
          gl_Position = projectionMatrix * modelViewMatrix * vec4(position, 1.0);
        }
      `,
      fragmentShader: `
        uniform sampler2D tDiffuse;
        uniform vec2 direction;
        uniform vec2 resolution;
        varying vec2 vUv;

        void main() {
          vec2 texSize = 1.0 / resolution;
          vec4 result = vec4(0.0);
          float weights[5];
          weights[0] = 0.227027; weights[1] = 0.1945946;
          weights[2] = 0.1216216; weights[3] = 0.054054;
          weights[4] = 0.016216;

          result += texture2D(tDiffuse, vUv) * weights[0];
          for (int i = 1; i < 5; i++) {
            vec2 offset = direction * texSize * float(i) * 2.0;
            result += texture2D(tDiffuse, vUv + offset) * weights[i];
            result += texture2D(tDiffuse, vUv - offset) * weights[i];
          }
          gl_FragColor = result;
        }
      `,
    });

    // Composite shader
    this._compositeMaterial = new THREE.ShaderMaterial({
      uniforms: {
        tDiffuse: { value: null },
        tBloom: { value: null },
        bloomStrength: { value: this.strength },
      },
      vertexShader: `
        varying vec2 vUv;
        void main() {
          vUv = uv;
          gl_Position = projectionMatrix * modelViewMatrix * vec4(position, 1.0);
        }
      `,
      fragmentShader: `
        uniform sampler2D tDiffuse;
        uniform sampler2D tBloom;
        uniform float bloomStrength;
        varying vec2 vUv;

        void main() {
          vec4 base = texture2D(tDiffuse, vUv);
          vec4 bloom = texture2D(tBloom, vUv);
          gl_FragColor = base + bloom * bloomStrength;
        }
      `,
    });

    this._quad = new THREE.Mesh(new THREE.PlaneGeometry(2, 2), null);
    this._orthoScene = new THREE.Scene();
    this._orthoCamera = new THREE.OrthographicCamera(-1, 1, 1, -1, 0, 1);
    this._orthoScene.add(this._quad);

    this._rt1 = null;
    this._rt2 = null;
  }

  _ensureTargets(w, h) {
    const THREE = getThree();
    if (this._rt1 && this._rt1.width === w && this._rt1.height === h) return;

    if (this._rt1) this._rt1.dispose();
    if (this._rt2) this._rt2.dispose();

    const halfW = Math.max(1, w >> 1);
    const halfH = Math.max(1, h >> 1);

    const opts = { minFilter: THREE.LinearFilter, magFilter: THREE.LinearFilter, format: THREE.RGBAFormat };

    this._rt1 = new THREE.WebGLRenderTarget(halfW, halfH, opts);
    this._rt2 = new THREE.WebGLRenderTarget(halfW, halfH, opts);
  }

  render(renderer, sceneRT, writeBuffer) {
    const size = renderer.getSize(new getThree().Vector2());
    this._ensureTargets(size.x, size.y);

    // 1. Extract luminance
    this._quad.material = this._luminanceMaterial;
    this._luminanceMaterial.uniforms.tDiffuse.value = sceneRT.texture;
    renderer.setRenderTarget(this._rt1);
    renderer.render(this._orthoScene, this._orthoCamera);

    // 2. Blur H
    this._quad.material = this._blurMaterial;
    this._blurMaterial.uniforms.tDiffuse.value = this._rt1.texture;
    this._blurMaterial.uniforms.direction.value.set(1.0, 0.0);
    this._blurMaterial.uniforms.resolution.value.set(this._rt1.width, this._rt1.height);
    renderer.setRenderTarget(this._rt2);
    renderer.render(this._orthoScene, this._orthoCamera);

    // 3. Blur V
    this._blurMaterial.uniforms.tDiffuse.value = this._rt2.texture;
    this._blurMaterial.uniforms.direction.value.set(0.0, 1.0);
    renderer.setRenderTarget(this._rt1);
    renderer.render(this._orthoScene, this._orthoCamera);

    // 4. Composite
    this._quad.material = this._compositeMaterial;
    this._compositeMaterial.uniforms.tDiffuse.value = sceneRT.texture;
    this._compositeMaterial.uniforms.tBloom.value = this._rt1.texture;
    this._compositeMaterial.uniforms.bloomStrength.value = this.strength;
    renderer.setRenderTarget(writeBuffer);
    renderer.render(this._orthoScene, this._orthoCamera);
  }

  dispose() {
    this._luminanceMaterial.dispose();
    this._blurMaterial.dispose();
    this._compositeMaterial.dispose();
    if (this._rt1) this._rt1.dispose();
    if (this._rt2) this._rt2.dispose();
    this._quad.geometry.dispose();
  }
}

// ---- FXAA Anti-Aliasing Pass --------------------------------------------

export class FXAAPass extends RenderPass {
  constructor() {
    super('fxaa');
    const THREE = getThree();

    this._material = new THREE.ShaderMaterial({
      uniforms: {
        tDiffuse: { value: null },
        resolution: { value: new THREE.Vector2(1, 1) },
      },
      vertexShader: `
        varying vec2 vUv;
        void main() {
          vUv = uv;
          gl_Position = projectionMatrix * modelViewMatrix * vec4(position, 1.0);
        }
      `,
      fragmentShader: `
        uniform sampler2D tDiffuse;
        uniform vec2 resolution;
        varying vec2 vUv;

        #define FXAA_REDUCE_MIN (1.0 / 128.0)
        #define FXAA_REDUCE_MUL (1.0 / 8.0)
        #define FXAA_SPAN_MAX 8.0

        void main() {
          vec2 texSize = 1.0 / resolution;
          vec3 rgbNW = texture2D(tDiffuse, vUv + vec2(-1.0, -1.0) * texSize).xyz;
          vec3 rgbNE = texture2D(tDiffuse, vUv + vec2( 1.0, -1.0) * texSize).xyz;
          vec3 rgbSW = texture2D(tDiffuse, vUv + vec2(-1.0,  1.0) * texSize).xyz;
          vec3 rgbSE = texture2D(tDiffuse, vUv + vec2( 1.0,  1.0) * texSize).xyz;
          vec3 rgbM  = texture2D(tDiffuse, vUv).xyz;

          vec3 luma = vec3(0.299, 0.587, 0.114);
          float lumaNW = dot(rgbNW, luma);
          float lumaNE = dot(rgbNE, luma);
          float lumaSW = dot(rgbSW, luma);
          float lumaSE = dot(rgbSE, luma);
          float lumaM  = dot(rgbM,  luma);

          float lumaMin = min(lumaM, min(min(lumaNW, lumaNE), min(lumaSW, lumaSE)));
          float lumaMax = max(lumaM, max(max(lumaNW, lumaNE), max(lumaSW, lumaSE)));

          vec2 dir;
          dir.x = -((lumaNW + lumaNE) - (lumaSW + lumaSE));
          dir.y =  ((lumaNW + lumaSW) - (lumaNE + lumaSE));

          float dirReduce = max((lumaNW + lumaNE + lumaSW + lumaSE) * (0.25 * FXAA_REDUCE_MUL), FXAA_REDUCE_MIN);
          float rcpDirMin = 1.0 / (min(abs(dir.x), abs(dir.y)) + dirReduce);
          dir = min(vec2(FXAA_SPAN_MAX), max(vec2(-FXAA_SPAN_MAX), dir * rcpDirMin)) * texSize;

          vec3 rgbA = 0.5 * (
            texture2D(tDiffuse, vUv + dir * (1.0 / 3.0 - 0.5)).xyz +
            texture2D(tDiffuse, vUv + dir * (2.0 / 3.0 - 0.5)).xyz);
          vec3 rgbB = rgbA * 0.5 + 0.25 * (
            texture2D(tDiffuse, vUv + dir * -0.5).xyz +
            texture2D(tDiffuse, vUv + dir *  0.5).xyz);

          float lumaB = dot(rgbB, luma);
          if (lumaB < lumaMin || lumaB > lumaMax) {
            gl_FragColor = vec4(rgbA, 1.0);
          } else {
            gl_FragColor = vec4(rgbB, 1.0);
          }
        }
      `,
    });

    this._quad = new THREE.Mesh(new THREE.PlaneGeometry(2, 2), this._material);
    this._scene = new THREE.Scene();
    this._camera = new THREE.OrthographicCamera(-1, 1, 1, -1, 0, 1);
    this._scene.add(this._quad);
  }

  render(renderer, readBuffer, writeBuffer) {
    const size = renderer.getSize(new getThree().Vector2());
    this._material.uniforms.tDiffuse.value = readBuffer.texture;
    this._material.uniforms.resolution.value.set(size.x, size.y);
    renderer.setRenderTarget(writeBuffer);
    renderer.render(this._scene, this._camera);
  }

  dispose() {
    this._material.dispose();
    this._quad.geometry.dispose();
  }
}

// ---- Post Processing Composer -------------------------------------------

export class PostProcessingComposer {
  constructor(renderer, scene, camera) {
    const THREE = getThree();
    this._renderer = renderer;
    this._scene = scene;
    this._camera = camera;
    this._passes = [];

    const size = renderer.getSize(new THREE.Vector2());
    const opts = { minFilter: THREE.LinearFilter, magFilter: THREE.LinearFilter, format: THREE.RGBAFormat };
    this._rtA = new THREE.WebGLRenderTarget(size.x, size.y, opts);
    this._rtB = new THREE.WebGLRenderTarget(size.x, size.y, opts);
  }

  addPass(pass) {
    this._passes.push(pass);
    return this;
  }

  removePass(name) {
    this._passes = this._passes.filter(p => p.name !== name);
    return this;
  }

  render(delta) {
    const renderer = this._renderer;
    const THREE = getThree();

    const size = renderer.getSize(new THREE.Vector2());
    if (this._rtA.width !== size.x || this._rtA.height !== size.y) {
      this._rtA.setSize(size.x, size.y);
      this._rtB.setSize(size.x, size.y);
    }

    // Render scene to first RT
    renderer.setRenderTarget(this._rtA);
    renderer.render(this._scene, this._camera);

    let read = this._rtA;
    let write = this._rtB;

    const activePasses = this._passes.filter(p => p.enabled);

    for (let i = 0; i < activePasses.length; i++) {
      const isLast = i === activePasses.length - 1;
      const target = isLast ? null : write; // Last pass renders to screen
      activePasses[i].render(renderer, read, target, delta);

      // Swap buffers
      if (!isLast) {
        const tmp = read;
        read = write;
        write = tmp;
      }
    }

    // If no passes, copy to screen
    if (activePasses.length === 0) {
      renderer.setRenderTarget(null);
      renderer.render(this._scene, this._camera);
    }
  }

  dispose() {
    for (const pass of this._passes) pass.dispose();
    this._passes = [];
    this._rtA.dispose();
    this._rtB.dispose();
  }
}

// ---- Instanced CSI Field Renderer (GPU-efficient) -----------------------

export class InstancedCSIField {
  constructor(scene, options = {}) {
    const THREE = getThree();

    this._scene = scene;
    this._gridX = options.gridX || 32;
    this._gridZ = options.gridZ || 32;
    this._cellSize = options.cellSize || 0.25;
    this._count = this._gridX * this._gridZ;

    // Instanced disc geometry
    const baseGeom = new THREE.CircleGeometry(this._cellSize * 0.45, 8);
    baseGeom.rotateX(-Math.PI / 2);

    this._mesh = new THREE.InstancedMesh(
      baseGeom,
      new THREE.MeshBasicMaterial({ vertexColors: true, transparent: true, opacity: 0.85, side: THREE.DoubleSide }),
      this._count
    );
    this._mesh.instanceMatrix.setUsage(THREE.DynamicDrawUsage);

    // Per-instance color
    this._colors = new Float32Array(this._count * 3);
    this._mesh.instanceColor = new THREE.InstancedBufferAttribute(this._colors, 3);
    this._mesh.instanceColor.setUsage(THREE.DynamicDrawUsage);

    // Data buffer for signal values
    this._data = new Float32Array(this._count);

    // Initialize positions
    const dummy = new THREE.Object3D();
    const halfX = (this._gridX * this._cellSize) / 2;
    const halfZ = (this._gridZ * this._cellSize) / 2;

    for (let iz = 0; iz < this._gridZ; iz++) {
      for (let ix = 0; ix < this._gridX; ix++) {
        const idx = iz * this._gridX + ix;
        dummy.position.set(
          ix * this._cellSize - halfX + this._cellSize * 0.5,
          0.01,
          iz * this._cellSize - halfZ + this._cellSize * 0.5
        );
        dummy.updateMatrix();
        this._mesh.setMatrixAt(idx, dummy.matrix);
      }
    }
    this._mesh.instanceMatrix.needsUpdate = true;

    scene.add(this._mesh);
  }

  /** Update with a flat Float32Array of signal values [0..1]. */
  update(data) {
    const len = Math.min(data.length, this._count);
    for (let i = 0; i < len; i++) {
      const v = Math.max(0, Math.min(1, data[i]));
      this._data[i] = v;

      // Blue → Cyan → Green → Yellow → Red
      let r, g, b;
      if (v < 0.25) {
        const t = v * 4;
        r = 0; g = t * 0.6; b = 0.4 + t * 0.4;
      } else if (v < 0.5) {
        const t = (v - 0.25) * 4;
        r = 0; g = 0.6 + t * 0.4; b = 0.8 - t * 0.8;
      } else if (v < 0.75) {
        const t = (v - 0.5) * 4;
        r = t; g = 1.0; b = 0;
      } else {
        const t = (v - 0.75) * 4;
        r = 1.0; g = 1.0 - t * 0.7; b = 0;
      }

      this._colors[i * 3] = r;
      this._colors[i * 3 + 1] = g;
      this._colors[i * 3 + 2] = b;
    }

    this._mesh.instanceColor.needsUpdate = true;
  }

  dispose() {
    this._scene.remove(this._mesh);
    this._mesh.geometry.dispose();
    this._mesh.material.dispose();
  }
}

// ---- CSI Particle System (motion visualization) -------------------------

export class CSIParticleSystem {
  constructor(scene, options = {}) {
    const THREE = getThree();

    this._scene = scene;
    this._maxParticles = options.maxParticles || 2000;
    this._emitRate = options.emitRate || 50; // particles/sec
    this._lifetime = options.lifetime || 3.0; // seconds
    this._activeCount = 0;

    // Particle data arrays
    this._positions = new Float32Array(this._maxParticles * 3);
    this._velocities = new Float32Array(this._maxParticles * 3);
    this._colors = new Float32Array(this._maxParticles * 3);
    this._sizes = new Float32Array(this._maxParticles);
    this._ages = new Float32Array(this._maxParticles);
    this._lifetimes = new Float32Array(this._maxParticles);

    // Geometry
    const geom = new THREE.BufferGeometry();
    geom.setAttribute('position', new THREE.BufferAttribute(this._positions, 3).setUsage(THREE.DynamicDrawUsage));
    geom.setAttribute('color', new THREE.BufferAttribute(this._colors, 3).setUsage(THREE.DynamicDrawUsage));
    geom.setAttribute('size', new THREE.BufferAttribute(this._sizes, 1).setUsage(THREE.DynamicDrawUsage));

    // Custom particle material
    const material = new THREE.ShaderMaterial({
      vertexShader: `
        attribute float size;
        attribute vec3 color;
        varying vec3 vColor;
        varying float vAlpha;

        void main() {
          vColor = color;
          vec4 mvPosition = modelViewMatrix * vec4(position, 1.0);
          gl_PointSize = size * (200.0 / -mvPosition.z);
          gl_Position = projectionMatrix * mvPosition;
          vAlpha = clamp(1.0 - length(position.xz) * 0.05, 0.1, 1.0);
        }
      `,
      fragmentShader: `
        varying vec3 vColor;
        varying float vAlpha;

        void main() {
          float dist = length(gl_PointCoord - vec2(0.5));
          if (dist > 0.5) discard;
          float alpha = smoothstep(0.5, 0.1, dist) * vAlpha;
          gl_FragColor = vec4(vColor, alpha * 0.6);
        }
      `,
      transparent: true,
      depthWrite: false,
      blending: THREE.AdditiveBlending,
    });

    this._points = new THREE.Points(geom, material);
    this._points.frustumCulled = false;
    scene.add(this._points);

    this._emitAccumulator = 0;
    this._motionIntensity = 0;
  }

  /** Set the desired motion intensity [0..1] which controls emit rate and color. */
  setMotionIntensity(intensity) {
    this._motionIntensity = Math.max(0, Math.min(1, intensity));
  }

  /** Set an emission origin point. */
  setEmitOrigin(x, y, z) {
    this._emitOrigin = { x, y, z };
  }

  /** Emit a burst of particles at a specific location. */
  burst(x, y, z, count = 20) {
    for (let i = 0; i < count; i++) {
      this._emitParticle(x, y, z, 2.0);
    }
  }

  _emitParticle(ox, oy, oz, speedMul = 1.0) {
    if (this._activeCount >= this._maxParticles) return;

    const idx = this._activeCount;
    const i3 = idx * 3;

    this._positions[i3] = ox + (Math.random() - 0.5) * 0.3;
    this._positions[i3 + 1] = oy + Math.random() * 0.2;
    this._positions[i3 + 2] = oz + (Math.random() - 0.5) * 0.3;

    const angle = Math.random() * Math.PI * 2;
    const speed = (0.3 + Math.random() * 0.7) * speedMul;
    this._velocities[i3] = Math.cos(angle) * speed;
    this._velocities[i3 + 1] = 0.5 + Math.random() * 1.5;
    this._velocities[i3 + 2] = Math.sin(angle) * speed;

    // Color based on motion intensity
    const m = this._motionIntensity;
    this._colors[i3] = 0.1 + m * 0.9;
    this._colors[i3 + 1] = 0.5 + (1 - m) * 0.5;
    this._colors[i3 + 2] = 1.0 - m * 0.8;

    this._sizes[idx] = 3 + Math.random() * 6;
    this._ages[idx] = 0;
    this._lifetimes[idx] = this._lifetime * (0.5 + Math.random() * 0.5);

    this._activeCount++;
  }

  update(delta) {
    // Emit new particles based on motion intensity
    const rate = this._emitRate * this._motionIntensity;
    this._emitAccumulator += rate * delta;

    while (this._emitAccumulator >= 1 && this._activeCount < this._maxParticles) {
      const o = this._emitOrigin || { x: 0, y: 0.5, z: 0 };
      this._emitParticle(o.x, o.y, o.z);
      this._emitAccumulator -= 1;
    }

    // Update existing particles
    for (let i = this._activeCount - 1; i >= 0; i--) {
      this._ages[i] += delta;

      if (this._ages[i] >= this._lifetimes[i]) {
        // Swap-remove
        this._swapRemove(i);
        continue;
      }

      const i3 = i * 3;
      const life = this._ages[i] / this._lifetimes[i];

      // Physics: gravity + drift
      this._velocities[i3 + 1] -= 1.5 * delta; // gravity
      this._positions[i3] += this._velocities[i3] * delta;
      this._positions[i3 + 1] += this._velocities[i3 + 1] * delta;
      this._positions[i3 + 2] += this._velocities[i3 + 2] * delta;

      // Fade out
      this._sizes[i] *= (1 - life * 0.3 * delta);

      // Floor bounce
      if (this._positions[i3 + 1] < 0.02) {
        this._positions[i3 + 1] = 0.02;
        this._velocities[i3 + 1] *= -0.3;
      }
    }

    // Update geometry
    const geom = this._points.geometry;
    geom.attributes.position.needsUpdate = true;
    geom.attributes.color.needsUpdate = true;
    geom.attributes.size.needsUpdate = true;
    geom.setDrawRange(0, this._activeCount);
  }

  _swapRemove(i) {
    const last = this._activeCount - 1;
    if (i !== last) {
      const i3 = i * 3, l3 = last * 3;
      this._positions[i3] = this._positions[l3];
      this._positions[i3 + 1] = this._positions[l3 + 1];
      this._positions[i3 + 2] = this._positions[l3 + 2];
      this._velocities[i3] = this._velocities[l3];
      this._velocities[i3 + 1] = this._velocities[l3 + 1];
      this._velocities[i3 + 2] = this._velocities[l3 + 2];
      this._colors[i3] = this._colors[l3];
      this._colors[i3 + 1] = this._colors[l3 + 1];
      this._colors[i3 + 2] = this._colors[l3 + 2];
      this._sizes[i] = this._sizes[last];
      this._ages[i] = this._ages[last];
      this._lifetimes[i] = this._lifetimes[last];
    }
    this._activeCount--;
  }

  dispose() {
    this._scene.remove(this._points);
    this._points.geometry.dispose();
    this._points.material.dispose();
  }
}

// ---- Volume Slicer (CSI field volume rendering) -------------------------

export class VolumeCSIRenderer {
  constructor(scene, options = {}) {
    const THREE = getThree();

    this._scene = scene;
    this._slices = options.slices || 16;
    this._width = options.width || 8;
    this._depth = options.depth || 6;
    this._height = options.height || 3;

    // Create stack of transparent planes
    this._group = new THREE.Group();
    this._group.name = 'volume-csi';
    this._planes = [];

    const sliceGeom = new THREE.PlaneGeometry(this._width, this._depth);
    sliceGeom.rotateX(-Math.PI / 2);

    for (let i = 0; i < this._slices; i++) {
      const y = (i / (this._slices - 1)) * this._height;
      const canvas = document.createElement('canvas');
      canvas.width = 64;
      canvas.height = 64;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = 'rgba(0, 40, 80, 0.02)';
      ctx.fillRect(0, 0, 64, 64);

      const texture = new THREE.CanvasTexture(canvas);
      texture.minFilter = THREE.LinearFilter;
      texture.magFilter = THREE.LinearFilter;

      const mat = new THREE.MeshBasicMaterial({
        map: texture,
        transparent: true,
        opacity: 0.15,
        side: THREE.DoubleSide,
        depthWrite: false,
        blending: THREE.AdditiveBlending,
      });

      const mesh = new THREE.Mesh(sliceGeom.clone(), mat);
      mesh.position.y = y;
      this._group.add(mesh);
      this._planes.push({ mesh, canvas, ctx, texture });
    }

    scene.add(this._group);
  }

  /** Update volume with 3D data: signalField[sliceIdx] = Float32Array(64*64). */
  updateSlice(sliceIdx, data) {
    if (sliceIdx < 0 || sliceIdx >= this._slices) return;
    const { canvas, ctx, texture } = this._planes[sliceIdx];
    const imgData = ctx.createImageData(64, 64);

    for (let i = 0; i < data.length && i < 64 * 64; i++) {
      const v = Math.max(0, Math.min(1, data[i]));
      const px = i * 4;
      // Blue-to-green-to-red gradient
      if (v < 0.5) {
        imgData.data[px] = 0;
        imgData.data[px + 1] = Math.floor(v * 2 * 200);
        imgData.data[px + 2] = Math.floor((1 - v * 2) * 200);
      } else {
        imgData.data[px] = Math.floor((v - 0.5) * 2 * 255);
        imgData.data[px + 1] = Math.floor((1 - (v - 0.5) * 2) * 200);
        imgData.data[px + 2] = 0;
      }
      imgData.data[px + 3] = Math.floor(v * 180);
    }

    ctx.putImageData(imgData, 0, 0);
    texture.needsUpdate = true;
  }

  dispose() {
    for (const p of this._planes) {
      p.mesh.geometry.dispose();
      p.mesh.material.dispose();
      p.texture.dispose();
    }
    this._scene.remove(this._group);
  }
}
