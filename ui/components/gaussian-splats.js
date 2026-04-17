/**
 * Gaussian Splat Renderer for WiFi Sensing Visualization
 *
 * Renders a 3D signal field using Three.js Points with custom ShaderMaterial.
 * Each "splat" is a screen-space disc whose size, color and opacity are driven
 * by the sensing data:
 *   - Size  : signal variance / disruption magnitude
 *   - Color : blue (quiet) -> green (presence) -> red (active motion)
 *   - Opacity: classification confidence
 */

// Use global THREE from CDN (loaded in SensingTab)
const getThree = () => window.THREE;

// ---- Custom Splat Shaders ------------------------------------------------

const SPLAT_VERTEX = `
  attribute float splatSize;
  attribute vec3  splatColor;
  attribute float splatOpacity;

  varying vec3  vColor;
  varying float vOpacity;

  void main() {
    vColor   = splatColor;
    vOpacity = splatOpacity;

    vec4 mvPosition = modelViewMatrix * vec4(position, 1.0);
    gl_PointSize = splatSize * (300.0 / -mvPosition.z);
    gl_Position  = projectionMatrix * mvPosition;
  }
`;

const SPLAT_FRAGMENT = `
  varying vec3  vColor;
  varying float vOpacity;

  void main() {
    // Circular soft-edge disc
    float dist = length(gl_PointCoord - vec2(0.5));
    if (dist > 0.5) discard;
    float alpha = smoothstep(0.5, 0.2, dist) * vOpacity;
    gl_FragColor = vec4(vColor, alpha);
  }
`;

// ---- Color helpers -------------------------------------------------------

/** Map a scalar 0-1 to blue -> green -> red gradient */
function valueToColor(v) {
  const clamped = Math.max(0, Math.min(1, v));
  // blue(0) -> cyan(0.25) -> green(0.5) -> yellow(0.75) -> red(1)
  let r, g, b;
  if (clamped < 0.5) {
    const t = clamped * 2;
    r = 0;
    g = t;
    b = 1 - t;
  } else {
    const t = (clamped - 0.5) * 2;
    r = t;
    g = 1 - t;
    b = 0;
  }
  return [r, g, b];
}

// ---- Node marker color palette -------------------------------------------

const NODE_MARKER_COLORS = [0x00ccff, 0xff6600, 0x00ff88, 0xff00cc, 0xffcc00, 0x8800ff, 0x00ffcc, 0xff0044];

// ---- COCO-17 skeleton connections ----------------------------------------

const COCO_SKELETON = [
  [0,1],[0,2],[1,3],[2,4],           // head
  [5,6],                              // shoulder bar
  [5,7],[7,9],[6,8],[8,10],          // arms
  [5,11],[6,12],[11,12],             // torso
  [11,13],[13,15],[12,14],[14,16],   // legs
];

const COCO_KP_NAMES = [
  'nose','left_eye','right_eye','left_ear','right_ear',
  'left_shoulder','right_shoulder',
  'left_elbow','right_elbow',
  'left_wrist','right_wrist',
  'left_hip','right_hip',
  'left_knee','right_knee',
  'left_ankle','right_ankle',
];

// Default positions to spread nodes in a triangle when server reports identical coords
const DEFAULT_NODE_POSITIONS = [
  [-7, 0.5, -7],   // node 1 — back-left corner
  [ 7, 0.5, -7],   // node 2 — back-right corner
  [ 0, 0.5,  7],   // node 3 — front-centre
  [-7, 0.5,  7],   // node 4 — front-left
  [ 7, 0.5,  7],   // node 5 — front-right
];

// ---- GaussianSplatRenderer -----------------------------------------------

export class GaussianSplatRenderer {
  /**
   * @param {HTMLElement} container - DOM element to attach the renderer to
   * @param {object}      [opts]
   * @param {number}      [opts.width]  - canvas width  (default container width)
   * @param {number}      [opts.height] - canvas height (default 500)
   */
  constructor(container, opts = {}) {
    const THREE = getThree();
    if (!THREE) throw new Error('Three.js not loaded');

    this.container = container;
    this.width  = opts.width  || container.clientWidth || 800;
    this.height = opts.height || 500;

    // Scene
    this.scene = new THREE.Scene();
    this.scene.background = new THREE.Color(0x0a0a12);

    // Camera — perspective looking down at the room
    this.camera = new THREE.PerspectiveCamera(55, this.width / this.height, 0.1, 200);
    this.camera.position.set(0, 14, 14);
    this.camera.lookAt(0, 0, 0);

    // Renderer
    this.renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
    this.renderer.setSize(this.width, this.height);
    this.renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    container.appendChild(this.renderer.domElement);

    // Grid & room
    this._createRoom(THREE);

    // Signal field splats (20x20 = 400 points on the floor plane)
    this.gridSize = 20;
    this._createFieldSplats(THREE);

    // Node markers (ESP32 / router positions)
    this._createNodeMarkers(THREE);

    // Dynamic per-node markers (multi-node support)
    this.nodeMarkers = new Map(); // nodeId -> THREE.Mesh
    this._THREE = THREE;

    // Body disruption blob
    this._createBodyBlob(THREE);

    // 3-D skeleton (joints + bones)
    this._createSkeleton(THREE);

    // Simple orbit-like mouse rotation
    this._setupMouseControls();

    // Animation state
    this._animFrame = null;
    this._lastData = null;

    // Start render loop
    this._animate();
  }

  // ---- Scene setup -------------------------------------------------------

  _createRoom(THREE) {
    // Floor grid
    const grid = new THREE.GridHelper(20, 20, 0x1a3a4a, 0x0d1f28);
    this.scene.add(grid);

    // Room boundary wireframe
    const boxGeo = new THREE.BoxGeometry(20, 6, 20);
    const edges  = new THREE.EdgesGeometry(boxGeo);
    const line   = new THREE.LineSegments(
      edges,
      new THREE.LineBasicMaterial({ color: 0x1a4a5a, opacity: 0.3, transparent: true })
    );
    line.position.y = 3;
    this.scene.add(line);
  }

  _createFieldSplats(THREE) {
    const count = this.gridSize * this.gridSize;

    const positions = new Float32Array(count * 3);
    const sizes     = new Float32Array(count);
    const colors    = new Float32Array(count * 3);
    const opacities = new Float32Array(count);

    // Lay splats on the floor plane (y = 0.05 to sit just above grid)
    for (let iz = 0; iz < this.gridSize; iz++) {
      for (let ix = 0; ix < this.gridSize; ix++) {
        const idx = iz * this.gridSize + ix;
        positions[idx * 3 + 0] = (ix - this.gridSize / 2) + 0.5; // x
        positions[idx * 3 + 1] = 0.05;                            // y
        positions[idx * 3 + 2] = (iz - this.gridSize / 2) + 0.5; // z

        sizes[idx]     = 1.5;
        colors[idx * 3]     = 0.1;
        colors[idx * 3 + 1] = 0.2;
        colors[idx * 3 + 2] = 0.6;
        opacities[idx] = 0.15;
      }
    }

    const geo = new THREE.BufferGeometry();
    geo.setAttribute('position',    new THREE.BufferAttribute(positions, 3));
    geo.setAttribute('splatSize',   new THREE.BufferAttribute(sizes, 1));
    geo.setAttribute('splatColor',  new THREE.BufferAttribute(colors, 3));
    geo.setAttribute('splatOpacity',new THREE.BufferAttribute(opacities, 1));

    const mat = new THREE.ShaderMaterial({
      vertexShader:   SPLAT_VERTEX,
      fragmentShader: SPLAT_FRAGMENT,
      transparent: true,
      depthWrite: false,
      blending: THREE.AdditiveBlending,
    });

    this.fieldPoints = new THREE.Points(geo, mat);
    this.scene.add(this.fieldPoints);
  }

  _createNodeMarkers(THREE) {
    // Router at center — green sphere
    const routerGeo = new THREE.SphereGeometry(0.3, 16, 16);
    const routerMat = new THREE.MeshBasicMaterial({ color: 0x00ff88, transparent: true, opacity: 0.8 });
    this.routerMarker = new THREE.Mesh(routerGeo, routerMat);
    this.routerMarker.position.set(0, 0.5, 0);
    this.scene.add(this.routerMarker);

    // ESP32 node — cyan sphere (default position, updated from data)
    const nodeGeo = new THREE.SphereGeometry(0.25, 16, 16);
    const nodeMat = new THREE.MeshBasicMaterial({ color: 0x00ccff, transparent: true, opacity: 0.8 });
    this.nodeMarker = new THREE.Mesh(nodeGeo, nodeMat);
    this.nodeMarker.position.set(2, 0.5, 1.5);
    this.scene.add(this.nodeMarker);
  }

  _createBodyBlob(THREE) {
    // A cluster of splats representing body disruption
    const count = 64;
    const positions = new Float32Array(count * 3);
    const sizes     = new Float32Array(count);
    const colors    = new Float32Array(count * 3);
    const opacities = new Float32Array(count);

    for (let i = 0; i < count; i++) {
      // Random sphere distribution
      const theta = Math.random() * Math.PI * 2;
      const phi   = Math.acos(2 * Math.random() - 1);
      const r     = Math.random() * 1.5;
      positions[i * 3]     = r * Math.sin(phi) * Math.cos(theta);
      positions[i * 3 + 1] = r * Math.cos(phi) + 2;
      positions[i * 3 + 2] = r * Math.sin(phi) * Math.sin(theta);

      sizes[i] = 2 + Math.random() * 3;
      colors[i * 3]     = 0.2;
      colors[i * 3 + 1] = 0.8;
      colors[i * 3 + 2] = 0.3;
      opacities[i] = 0.0; // hidden until presence detected
    }

    const geo = new THREE.BufferGeometry();
    geo.setAttribute('position',    new THREE.BufferAttribute(positions, 3));
    geo.setAttribute('splatSize',   new THREE.BufferAttribute(sizes, 1));
    geo.setAttribute('splatColor',  new THREE.BufferAttribute(colors, 3));
    geo.setAttribute('splatOpacity',new THREE.BufferAttribute(opacities, 1));

    const mat = new THREE.ShaderMaterial({
      vertexShader:   SPLAT_VERTEX,
      fragmentShader: SPLAT_FRAGMENT,
      transparent: true,
      depthWrite: false,
      blending: THREE.AdditiveBlending,
    });

    this.bodyBlob = new THREE.Points(geo, mat);
    this.scene.add(this.bodyBlob);
  }

  // ---- 3-D Skeleton ------------------------------------------------------

  _createSkeleton(THREE) {
    const BONE_COUNT  = COCO_SKELETON.length;  // 16
    const JOINT_COUNT = COCO_KP_NAMES.length;  // 17

    // -- Bone lines --
    const bonePositions = new Float32Array(BONE_COUNT * 6); // 2 verts × 3 floats
    const boneGeo = new THREE.BufferGeometry();
    boneGeo.setAttribute('position', new THREE.BufferAttribute(bonePositions, 3));
    const boneMat = new THREE.LineBasicMaterial({
      color: 0x00e5ff, transparent: true, opacity: 0.0, linewidth: 2,
    });
    this.skeletonBones = new THREE.LineSegments(boneGeo, boneMat);
    this.scene.add(this.skeletonBones);

    // -- Joint splats (re-use custom shader for consistency) --
    const jPositions = new Float32Array(JOINT_COUNT * 3);
    const jSizes     = new Float32Array(JOINT_COUNT).fill(14);
    const jColors    = new Float32Array(JOINT_COUNT * 3);
    const jOpacities = new Float32Array(JOINT_COUNT).fill(0.0);

    // head joints cyan, torso lime-green, legs/arms amber
    const HEAD_JOINTS  = new Set([0,1,2,3,4]);
    const TORSO_JOINTS = new Set([5,6,11,12]);
    for (let i = 0; i < JOINT_COUNT; i++) {
      if (HEAD_JOINTS.has(i)) {
        jColors[i*3]=0.0; jColors[i*3+1]=0.9; jColors[i*3+2]=1.0;
      } else if (TORSO_JOINTS.has(i)) {
        jColors[i*3]=0.2; jColors[i*3+1]=1.0; jColors[i*3+2]=0.3;
      } else {
        jColors[i*3]=1.0; jColors[i*3+1]=0.7; jColors[i*3+2]=0.0;
      }
    }

    const jGeo = new THREE.BufferGeometry();
    jGeo.setAttribute('position',    new THREE.BufferAttribute(jPositions, 3));
    jGeo.setAttribute('splatSize',   new THREE.BufferAttribute(jSizes, 1));
    jGeo.setAttribute('splatColor',  new THREE.BufferAttribute(jColors, 3));
    jGeo.setAttribute('splatOpacity',new THREE.BufferAttribute(jOpacities, 1));

    this.skeletonJoints = new THREE.Points(jGeo, new THREE.ShaderMaterial({
      vertexShader: SPLAT_VERTEX, fragmentShader: SPLAT_FRAGMENT,
      transparent: true, depthWrite: false, blending: THREE.AdditiveBlending,
    }));
    this.scene.add(this.skeletonJoints);
  }

  /** Convert image-space keypoint to 3-D world coordinates.
   *  cx, bottomY — centroid-x and ankle-level-y of the person in image px.
   *  scale — px→m conversion. */
  _kpToWorld(kp, cx, bottomY, scale) {
    return [
      (kp.x - cx) * scale,         // world X (left/right)
      (bottomY - kp.y) * scale,    // world Y (height — image y is inverted)
      0,                            // world Z (stick person on centre plane)
    ];
  }

  /** Update 3-D skeleton from persons array. */
  _updateSkeleton(persons) {
    if (!this.skeletonBones || !this.skeletonJoints) return;

    const jGeo  = this.skeletonJoints.geometry;
    const jPos  = jGeo.attributes.position.array;
    const jOpac = jGeo.attributes.splatOpacity.array;
    const bGeo  = this.skeletonBones.geometry;
    const bPos  = bGeo.attributes.position.array;

    const hide = () => {
      this.skeletonBones.material.opacity = 0;
      jOpac.fill(0);
      jGeo.attributes.splatOpacity.needsUpdate = true;
    };

    if (!persons || persons.length === 0) { hide(); return; }

    const person = persons[0];
    const kps    = person.keypoints || [];
    if (kps.length === 0) { hide(); return; }

    // Build name→kp map
    const kpMap = {};
    kps.forEach(k => { kpMap[k.name] = k; });

    // Collect all available coords to compute centroid + scale
    const allX = kps.map(k => k.x);
    const allY = kps.map(k => k.y);
    const cx      = allX.reduce((a, b) => a + b, 0) / allX.length;
    const bottomY = Math.max(...allY);  // ankle level (image y increases downward)
    const topY    = Math.min(...allY);  // head level
    const heightPx = bottomY - topY;

    // Scale so that the rendered skeleton is ~1.7 m tall in world space
    const scale = heightPx > 20 ? 1.7 / heightPx : 0.009;

    // Build ordered world positions
    const worldKps = COCO_KP_NAMES.map(name => {
      const kp = kpMap[name];
      return kp ? this._kpToWorld(kp, cx, bottomY, scale) : null;
    });

    // Move body blob to hip midpoint
    const lh = kpMap['left_hip'], rh = kpMap['right_hip'];
    if (lh && rh && this.bodyBlob) {
      const hipX = ((lh.x + rh.x) / 2 - cx) * scale;
      const hipY = (bottomY - (lh.y + rh.y) / 2) * scale;
      this.bodyBlob.position.set(hipX, hipY + 0.1, 0);
    }

    // Update joint positions
    worldKps.forEach((pos, i) => {
      if (pos) {
        jPos[i*3] = pos[0]; jPos[i*3+1] = pos[1]; jPos[i*3+2] = pos[2];
        jOpac[i]  = 0.95;
      } else {
        jOpac[i] = 0;
      }
    });

    // Update bone positions
    COCO_SKELETON.forEach(([ai, bi], b) => {
      const a = worldKps[ai], bv = worldKps[bi];
      if (a && bv) {
        bPos[b*6+0]=a[0];  bPos[b*6+1]=a[1];  bPos[b*6+2]=a[2];
        bPos[b*6+3]=bv[0]; bPos[b*6+4]=bv[1]; bPos[b*6+5]=bv[2];
      } else {
        bPos.fill(0, b*6, b*6+6);
      }
    });

    this.skeletonBones.material.opacity = 0.85;
    bGeo.attributes.position.needsUpdate  = true;
    jGeo.attributes.position.needsUpdate  = true;
    jGeo.attributes.splatOpacity.needsUpdate = true;
  }

  // ---- Mouse controls (simple orbit) -------------------------------------

  _setupMouseControls() {
    let isDragging = false;
    let prevX = 0, prevY = 0;
    let azimuth = 0, elevation = 55;
    const radius = 20;

    const updateCamera = () => {
      const phi   = (elevation * Math.PI) / 180;
      const theta = (azimuth * Math.PI) / 180;
      this.camera.position.set(
        radius * Math.sin(phi) * Math.sin(theta),
        radius * Math.cos(phi),
        radius * Math.sin(phi) * Math.cos(theta)
      );
      this.camera.lookAt(0, 0, 0);
    };

    const canvas = this.renderer.domElement;
    canvas.addEventListener('mousedown', (e) => {
      isDragging = true;
      prevX = e.clientX;
      prevY = e.clientY;
    });
    canvas.addEventListener('mousemove', (e) => {
      if (!isDragging) return;
      azimuth   += (e.clientX - prevX) * 0.4;
      elevation -= (e.clientY - prevY) * 0.4;
      elevation  = Math.max(15, Math.min(85, elevation));
      prevX = e.clientX;
      prevY = e.clientY;
      updateCamera();
    });
    canvas.addEventListener('mouseup',   () => { isDragging = false; });
    canvas.addEventListener('mouseleave',() => { isDragging = false; });

    // Scroll to zoom
    canvas.addEventListener('wheel', (e) => {
      e.preventDefault();
      const delta = e.deltaY > 0 ? 1.05 : 0.95;
      this.camera.position.multiplyScalar(delta);
      this.camera.position.clampLength(8, 40);
    }, { passive: false });

    updateCamera();
  }

  // ---- Data update -------------------------------------------------------

  /**
   * Update the visualization with new sensing data.
   * @param {object} data - sensing_update JSON from ws_server
   */
  update(data) {
    this._lastData = data;
    if (!data) return;

    const features       = data.features        || {};
    const classification = data.classification  || {};
    const signalField    = data.signal_field     || {};
    const nodes          = data.nodes            || [];
    const persons        = data.persons          || [];

    // -- Update signal field splats ----------------------------------------
    if (signalField.values && this.fieldPoints) {
      const geo    = this.fieldPoints.geometry;
      const clr    = geo.attributes.splatColor.array;
      const sizes  = geo.attributes.splatSize.array;
      const opac   = geo.attributes.splatOpacity.array;
      const vals   = signalField.values;
      const count  = Math.min(vals.length, this.gridSize * this.gridSize);

      for (let i = 0; i < count; i++) {
        const v = vals[i];
        const [r, g, b] = valueToColor(v);
        clr[i * 3]     = r;
        clr[i * 3 + 1] = g;
        clr[i * 3 + 2] = b;
        sizes[i] = 1.0 + v * 4.0;
        opac[i]  = 0.1 + v * 0.6;
      }

      geo.attributes.splatColor.needsUpdate  = true;
      geo.attributes.splatSize.needsUpdate   = true;
      geo.attributes.splatOpacity.needsUpdate = true;
    }

    // -- Update body blob --------------------------------------------------
    if (this.bodyBlob) {
      const bGeo  = this.bodyBlob.geometry;
      const bOpac = bGeo.attributes.splatOpacity.array;
      const bClr  = bGeo.attributes.splatColor.array;
      const bSize = bGeo.attributes.splatSize.array;
      const bPos  = bGeo.attributes.position.array;

      const presence   = classification.presence     || false;
      const motionLvl  = classification.motion_level || 'absent';
      const confidence = classification.confidence   || 0;
      const breathing  = features.breathing_band_power || 0;

      // Breathing pulsation
      const breathPulse = 1.0 + Math.sin(Date.now() * 0.004) * Math.min(breathing * 3, 0.4);

      for (let i = 0; i < bOpac.length; i++) {
        if (presence) {
          // Always show blob when presence=true — use max(confidence, 0.45) so it's
          // visible even while confidence calibrates up from 0
          bOpac[i] = Math.max(confidence, 0.45) * 0.5;

          // Color by motion level
          if (motionLvl === 'active') {
            bClr[i * 3]     = 1.0;
            bClr[i * 3 + 1] = 0.2;
            bClr[i * 3 + 2] = 0.1;
          } else {
            bClr[i * 3]     = 0.1;
            bClr[i * 3 + 1] = 0.8;
            bClr[i * 3 + 2] = 0.4;
          }

          bSize[i] = (2 + Math.random() * 2) * breathPulse;
        } else {
          bOpac[i] = 0.0;
        }
      }

      bGeo.attributes.splatOpacity.needsUpdate = true;
      bGeo.attributes.splatColor.needsUpdate   = true;
      bGeo.attributes.splatSize.needsUpdate    = true;
    }

    // -- Update node positions (legacy single-node) ------------------------
    if (nodes.length > 0 && nodes[0].position) {
      const pos = nodes[0].position;
      this.nodeMarker.position.set(pos[0], 0.5, pos[2]);
    }

    // -- Update dynamic per-node markers (multi-node support) --------------
    if (nodes && nodes.length > 0 && this.scene) {
      const THREE = this._THREE || window.THREE;
      if (THREE) {
        // Detect whether all nodes share identical positions (server default)
        const positions = nodes.map(n => n.position || [0, 0, 0]);
        const allSame = positions.every(p =>
          p[0] === positions[0][0] && p[1] === positions[0][1] && p[2] === positions[0][2]
        );

        const activeIds = new Set();
        for (let ni = 0; ni < nodes.length; ni++) {
          const node = nodes[ni];
          activeIds.add(node.node_id);
          if (!this.nodeMarkers.has(node.node_id)) {
            const geo = new THREE.SphereGeometry(0.25, 16, 16);
            const mat = new THREE.MeshBasicMaterial({
              color: NODE_MARKER_COLORS[node.node_id % NODE_MARKER_COLORS.length],
              transparent: true,
              opacity: 0.8,
            });
            const marker = new THREE.Mesh(geo, mat);
            this.scene.add(marker);
            this.nodeMarkers.set(node.node_id, marker);
          }
          const marker = this.nodeMarkers.get(node.node_id);
          // Spread nodes around room if all positions are identical (not provisioned)
          if (allSame) {
            const fallback = DEFAULT_NODE_POSITIONS[ni % DEFAULT_NODE_POSITIONS.length];
            marker.position.set(fallback[0], fallback[1], fallback[2]);
          } else {
            const pos = node.position || [0, 0, 0];
            marker.position.set(pos[0], 0.5, pos[2]);
          }
        }
        // Remove stale markers
        for (const [id, marker] of this.nodeMarkers) {
          if (!activeIds.has(id)) {
            this.scene.remove(marker);
            this.nodeMarkers.delete(id);
          }
        }
      }
    }

    // -- Update 3-D skeleton -----------------------------------------------
    this._updateSkeleton(persons);
  }

  // ---- Render loop -------------------------------------------------------

  _animate() {
    this._animFrame = requestAnimationFrame(() => this._animate());

    // Gentle router glow pulse
    if (this.routerMarker) {
      const pulse = 0.6 + 0.3 * Math.sin(Date.now() * 0.003);
      this.routerMarker.material.opacity = pulse;
    }

    this.renderer.render(this.scene, this.camera);
  }

  // ---- Resize / cleanup --------------------------------------------------

  resize(width, height) {
    this.width  = width;
    this.height = height;
    this.camera.aspect = width / height;
    this.camera.updateProjectionMatrix();
    this.renderer.setSize(width, height);
  }

  dispose() {
    if (this._animFrame) {
      cancelAnimationFrame(this._animFrame);
    }
    if (this.skeletonBones)  this.scene.remove(this.skeletonBones);
    if (this.skeletonJoints) this.scene.remove(this.skeletonJoints);
    this.renderer.dispose();
    if (this.renderer.domElement.parentNode) {
      this.renderer.domElement.parentNode.removeChild(this.renderer.domElement);
    }
  }
}
