import { injectNav, NAV_CSS } from './shared/nav.js';
import { toast, busy, stickyError, fetchJSON } from './shared/feedback.js';

const navStyle = document.createElement('style');
navStyle.textContent = NAV_CSS;
document.head.appendChild(navStyle);
injectNav();

const WS_URL = `ws://${location.hostname}:3001/ws/sensing`;
const ROOM_URL = '/api/v1/room';
const LATEST_URL = '/api/v1/sensing/latest';
const POSITIONS_URL = '/api/v1/nodes/positions';
const AUTO_CAL_URL = '/api/v1/calibration/auto';
const STATE = {
  room: { width: 8, depth: 6, height: 3, calibrated: false, nodePositions: [], activeNodes: [] },
  data: null,
  hotspots: [],
  connection: 'connecting',
  lastUpdateMs: 0,
  ws: null,
};
const $ = (id) => document.getElementById(id);

/**
 * Distribute N nodes evenly around the room perimeter at 1.6 m height.
 * Used both as a visual fallback when nodes are stacked at the
 * uncalibrated default, AND as the payload for the "Apply default
 * layout" button which POSTs to /api/v1/nodes/positions.
 */
function defaultPerimeterLayout(count, room) {
  if (count <= 0) return [];
  const W = room.width, D = room.depth, H = room.height;
  // Anchor positions: corners first, then mid-walls. Always inside by 0.3 m.
  const inset = 0.3;
  const candidates = [
    [inset,        H * 0.55, inset],
    [W - inset,    H * 0.55, inset],
    [W - inset,    H * 0.55, D - inset],
    [inset,        H * 0.55, D - inset],
    [W * 0.5,      H * 0.55, inset],
    [W * 0.5,      H * 0.55, D - inset],
    [inset,        H * 0.55, D * 0.5],
    [W - inset,    H * 0.55, D * 0.5],
  ];
  // Pad with interpolated points if more than 8 nodes.
  while (candidates.length < count) {
    const t = candidates.length / count;
    candidates.push([W * t, H * 0.55, D * (1 - t)]);
  }
  return candidates.slice(0, count);
}

/** Return true if all positions in the array sit within `eps` of each other. */
function nodesAreStacked(nodes, eps = 0.05) {
  if (nodes.length < 2) return false;
  const [x0, y0, z0] = [nodes[0].x, nodes[0].y, nodes[0].z];
  return nodes.every(n =>
    Math.abs(n.x - x0) < eps &&
    Math.abs(n.y - y0) < eps &&
    Math.abs(n.z - z0) < eps
  );
}

function fitCanvas(canvas) {
  const ratio = Math.min(window.devicePixelRatio || 1, 2);
  const rect = canvas.getBoundingClientRect();
  const width = Math.max(180, Math.floor(rect.width * ratio));
  const height = Math.max(140, Math.floor(rect.height * ratio));
  if (canvas.width !== width || canvas.height !== height) {
    canvas.width = width;
    canvas.height = height;
  }
  const ctx = canvas.getContext('2d');
  ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
  return { ctx, width: rect.width, height: rect.height };
}

function roomVolume() {
  return STATE.room.width * STATE.room.depth * STATE.room.height;
}

function formatAgo(ms) {
  if (!ms) return '--';
  const delta = Math.max(0, Date.now() - ms);
  if (delta < 1000) return 'now';
  if (delta < 60000) return `${(delta / 1000).toFixed(1)} s ago`;
  return `${Math.round(delta / 60000)} min ago`;
}

function mergedNodes() {
  const roomPositions = new Map((STATE.room.nodePositions || []).map((node) => [node.node_id, node]));
  const activeNodes = new Map((STATE.room.activeNodes || []).map((node) => [node.node_id, node]));
  const frameNodes = new Map(((STATE.data && STATE.data.nodes) || []).map((node) => [node.node_id, node]));
  const ids = new Set([...roomPositions.keys(), ...activeNodes.keys(), ...frameNodes.keys()]);
  const merged = [...ids].sort((a, b) => a - b).map((id) => {
    const roomNode = roomPositions.get(id) || {};
    const activeNode = activeNodes.get(id) || {};
    const frameNode = frameNodes.get(id) || {};
    const pos = frameNode.position || [roomNode.x, roomNode.y, roomNode.z];
    return {
      node_id: id,
      x: Number.isFinite(pos && pos[0]) ? pos[0] : STATE.room.width * 0.5,
      y: Number.isFinite(pos && pos[1]) ? pos[1] : STATE.room.height * 0.5,
      z: Number.isFinite(pos && pos[2]) ? pos[2] : STATE.room.depth * 0.5,
      status: activeNode.status || (frameNodes.has(id) ? 'active' : 'offline'),
      rssi_dbm: activeNode.rssi_dbm ?? frameNode.rssi_dbm ?? null,
      subcarrier_count: frameNode.subcarrier_count || 0,
      placeholder: false,
    };
  });

  // VISUAL FALLBACK: if every node sits at the uncalibrated default
  // (typically [2.0, 0.0, 1.5]), spread them around the room perimeter
  // so the atlas isn't a single overlapping dot. The placeholder flag
  // is exposed so the UI can show this is *not* a real measurement.
  if (!STATE.room.calibrated && nodesAreStacked(merged)) {
    const layout = defaultPerimeterLayout(merged.length, STATE.room);
    merged.forEach((node, i) => {
      const [x, y, z] = layout[i] || [node.x, node.y, node.z];
      node.x = x; node.y = y; node.z = z;
      node.placeholder = true;
    });
  }
  return merged;
}

function extractHotspots(signalField, limit = 6) {
  if (!signalField || !Array.isArray(signalField.values)) return [];
  const [gx = 20, , gz = 20] = signalField.grid_size || [20, 1, 20];
  const values = signalField.values;
  const maxValue = values.reduce((max, value) => Math.max(max, Math.abs(value || 0)), 0);
  if (!maxValue) return [];
  const peaks = [];
  for (let z = 0; z < gz; z++) {
    for (let x = 0; x < gx; x++) {
      const idx = z * gx + x;
      const value = Math.abs(values[idx] || 0);
      if (value < maxValue * 0.35) continue;
      let peak = true;
      for (let dz = -1; dz <= 1 && peak; dz++) {
        for (let dx = -1; dx <= 1; dx++) {
          if (dx === 0 && dz === 0) continue;
          const nx = x + dx;
          const nz = z + dz;
          if (nx < 0 || nz < 0 || nx >= gx || nz >= gz) continue;
          if (Math.abs(values[nz * gx + nx] || 0) > value) { peak = false; break; }
        }
      }
      if (peak) peaks.push({ x, z, value });
    }
  }
  peaks.sort((a, b) => b.value - a.value);
  return peaks.slice(0, limit).map((peak, index) => ({
    label: `Hotspot ${index + 1}`,
    intensity: peak.value / maxValue,
    x: ((peak.x + 0.5) / gx) * STATE.room.width,
    y: 0.65 + (peak.value / maxValue) * Math.min(STATE.room.height * 0.5, 1.6),
    z: ((peak.z + 0.5) / gz) * STATE.room.depth,
  }));
}

function roomSubjects() {
  const anchors = (STATE.data && STATE.data.person_anchors) || [];
  if (anchors.length) {
    return anchors.map((anchor, index) => ({
      label: `Subject ${index + 1}`,
      x: anchor.x,
      y: anchor.y || 0.95,
      z: anchor.z,
      intensity: anchor.strength || 0.5,
      confidence: STATE.data.persons && STATE.data.persons[index] ? STATE.data.persons[index].confidence : null,
    }));
  }
  const fallbackCount = Math.max(0, STATE.data && STATE.data.estimated_persons || 0);
  return STATE.hotspots.slice(0, fallbackCount).map((spot, index) => ({
    label: `Subject ${index + 1}`,
    x: spot.x,
    y: 0.95,
    z: spot.z,
    intensity: spot.intensity,
    confidence: null,
  }));
}

function renderHero() {
  const data = STATE.data || {};
  const cls = data.classification || {};
  $('hero-connection').textContent = STATE.connection === 'live' ? 'Live' : STATE.connection === 'snapshot' ? 'Snapshot' : STATE.connection === 'stale' ? 'Stale' : 'Connecting';
  $('hero-source').textContent = data.source || '--';
  $('hero-room').textContent = `${STATE.room.width.toFixed(1)} x ${STATE.room.depth.toFixed(1)} x ${STATE.room.height.toFixed(1)} m`;
  $('hero-update').textContent = formatAgo(STATE.lastUpdateMs);
  $('pill-connection').className = `pill ${STATE.connection === 'live' ? 'ok' : STATE.connection === 'stale' ? 'warn' : ''}`;
  $('pill-source').className = `pill ${cls.presence ? 'ok' : ''}`;
  $('pill-room').className = `pill ${STATE.room.calibrated ? 'ok' : 'warn'}`;
  $('pill-update').className = `pill ${STATE.lastUpdateMs ? '' : 'warn'}`;
  $('atlas-note').textContent = data.timestamp ? `Frame ${Math.round(data.tick || 0)} from ${data.source || 'unknown'} - ${formatAgo(STATE.lastUpdateMs)}` : 'Waiting for sensing data...';
}

function renderMetrics() {
  const data = STATE.data || {};
  const cls = data.classification || {};
  const vitals = data.vital_signs || {};
  const evidence = data.count_evidence || {};
  const nodes = mergedNodes();
  const liveNodes = nodes.filter((node) => node.status !== 'offline').length;
  $('metric-presence').textContent = cls.presence ? 'Present' : 'Absent';
  $('metric-motion').textContent = `Motion ${String((data.enhanced_motion && data.enhanced_motion.level) || cls.motion_level || '--').toUpperCase()}`;
  $('metric-persons').textContent = String(data.estimated_persons || 0);
  $('metric-consensus').textContent = evidence.consensus_confidence != null ? `Consensus ${(evidence.consensus_confidence * 100).toFixed(0)}%` : 'Consensus --';
  $('metric-quality').textContent = data.signal_quality_score != null ? `${Math.round(data.signal_quality_score * 100)}%` : '--';
  $('metric-quality-verdict').textContent = `Verdict ${data.quality_verdict || '--'}`;
  $('metric-vitals').textContent = vitals.breathing_rate_bpm ? `${vitals.breathing_rate_bpm.toFixed(1)} RPM` : '--';
  $('metric-vitals-sub').textContent = vitals.heart_rate_bpm ? `Heart ${vitals.heart_rate_bpm.toFixed(0)} BPM` : 'Heart --';
  $('metric-volume').textContent = `${roomVolume().toFixed(0)} m3`;
  $('metric-room-shape').textContent = `${STATE.room.width.toFixed(1)} x ${STATE.room.depth.toFixed(1)} x ${STATE.room.height.toFixed(1)}`;
  $('metric-coverage').textContent = String(roomSubjects().length + STATE.hotspots.length);
  $('metric-nodes').textContent = `${liveNodes} live / ${nodes.length} known`;
}

function renderEvidence() {
  const host = $('evidence-list');
  const evidence = STATE.data && STATE.data.count_evidence;
  if (!evidence || !Array.isArray(evidence.estimates) || !evidence.estimates.length) {
    host.innerHTML = '<div class="empty">No evidence packet yet. This panel fills when sensing frames carry weighted occupancy votes.</div>';
    return;
  }
  host.innerHTML = evidence.estimates.map((estimate) => `
    <div class="evidence-row">
      <div class="row-top">
        <div>
          <div class="row-label">${estimate.source.replaceAll('_', ' ')}</div>
          <div class="row-meta">count ${estimate.count}</div>
        </div>
        <div class="row-meta">${Math.round((estimate.confidence || 0) * 100)}%</div>
      </div>
      <div class="bar"><span style="width:${Math.max(6, Math.round((estimate.confidence || 0) * 100))}%"></span></div>
    </div>
  `).join('') + `
    <div class="evidence-row">
      <div class="row-top">
        <div>
          <div class="row-label">Consensus result</div>
          <div class="row-meta">count ${evidence.consensus_count}</div>
        </div>
        <div class="row-meta">${Math.round((evidence.consensus_confidence || 0) * 100)}%</div>
      </div>
      <div class="bar person"><span style="width:${Math.max(6, Math.round((evidence.consensus_confidence || 0) * 100))}%"></span></div>
    </div>`;
}

function renderNodes() {
  const host = $('node-list');
  const nodes = mergedNodes();
  if (!nodes.length) {
    host.innerHTML = '<div class="empty">No nodes reported yet. Once /api/v1/room or the sensing stream reports nodes, they will appear here with positions.</div>';
    return;
  }
  host.innerHTML = nodes.map((node) => `
    <div class="node-row">
      <div class="row-top">
        <div class="row-label">Node ${node.node_id}${node.placeholder ? ' <span style="color:#ff9f43;font-size:10px;letter-spacing:1px">(placeholder)</span>' : ''}</div>
        <div class="row-meta">${node.status}</div>
      </div>
      <div class="meta-grid">
        <div>Position <strong>${node.x.toFixed(1)}, ${node.y.toFixed(1)}, ${node.z.toFixed(1)} m</strong></div>
        <div>RSSI <strong>${node.rssi_dbm != null ? node.rssi_dbm.toFixed(0) + ' dBm' : '--'}</strong></div>
        <div>Subcarriers <strong>${node.subcarrier_count || '--'}</strong></div>
        <div>Role <strong>${node.placeholder ? 'Layout fallback' : (node.status === 'active' ? 'Streaming' : 'Known point')}</strong></div>
      </div>
    </div>
  `).join('');
}

function renderObjects() {
  const host = $('object-list');
  const subjects = roomSubjects();
  const people = (STATE.data && STATE.data.persons) || [];
  const sections = [];
  if (subjects.length) {
    sections.push(...subjects.map((subject, index) => `
      <div class="object-row">
        <div class="row-top">
          <div class="row-label">${subject.label}</div>
          <div class="row-meta">anchor</div>
        </div>
        <div class="meta-grid">
          <div>Location <strong>${subject.x.toFixed(1)}, ${subject.z.toFixed(1)} m</strong></div>
          <div>Height <strong>${subject.y.toFixed(2)} m</strong></div>
          <div>Signal <strong>${Math.round((subject.intensity || 0) * 100)}%</strong></div>
          <div>Track <strong>${people[index] ? '#' + people[index].id : '--'}</strong></div>
        </div>
        <div class="bar person"><span style="width:${Math.max(6, Math.round((subject.intensity || 0) * 100))}%"></span></div>
      </div>`));
  }
  if (STATE.hotspots.length) {
    sections.push(...STATE.hotspots.map((spot) => `
      <div class="object-row">
        <div class="row-top">
          <div class="row-label">${spot.label}</div>
          <div class="row-meta">field peak</div>
        </div>
        <div class="meta-grid">
          <div>Location <strong>${spot.x.toFixed(1)}, ${spot.z.toFixed(1)} m</strong></div>
          <div>Elevation <strong>${spot.y.toFixed(2)} m</strong></div>
          <div>Intensity <strong>${Math.round(spot.intensity * 100)}%</strong></div>
          <div>Purpose <strong>signal hotspot</strong></div>
        </div>
        <div class="bar hot"><span style="width:${Math.max(6, Math.round(spot.intensity * 100))}%"></span></div>
      </div>`));
  }
  host.innerHTML = sections.length ? sections.join('') : '<div class="empty">No room objects yet. Once live frames arrive, anchors and hotspots will populate this atlas.</div>';
}

function drawRoomRect(ctx, width, height) {
  ctx.strokeStyle = 'rgba(0,136,255,0.22)';
  ctx.lineWidth = 1;
  ctx.strokeRect(18, 18, width - 36, height - 36);
}

function drawPlan() {
  const { ctx, width, height } = fitCanvas($('canvas-top'));
  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = '#030510';
  ctx.fillRect(0, 0, width, height);
  drawRoomRect(ctx, width, height);
  const mapX = (x) => 18 + (x / STATE.room.width) * (width - 36);
  const mapZ = (z) => 18 + (z / STATE.room.depth) * (height - 36);
  STATE.hotspots.forEach((spot) => {
    const radius = 12 + spot.intensity * 28;
    const gradient = ctx.createRadialGradient(mapX(spot.x), mapZ(spot.z), 2, mapX(spot.x), mapZ(spot.z), radius);
    gradient.addColorStop(0, `rgba(255,159,67,${0.45 + spot.intensity * 0.25})`);
    gradient.addColorStop(1, 'rgba(255,159,67,0)');
    ctx.fillStyle = gradient;
    ctx.beginPath(); ctx.arc(mapX(spot.x), mapZ(spot.z), radius, 0, Math.PI * 2); ctx.fill();
  });
  mergedNodes().forEach((node) => {
    const x = mapX(node.x);
    const y = mapZ(node.z);
    ctx.fillStyle = node.placeholder ? '#ff9f43' : (node.status === 'active' ? '#00ccff' : '#5577aa');
    ctx.beginPath(); ctx.moveTo(x, y - 8); ctx.lineTo(x + 8, y); ctx.lineTo(x, y + 8); ctx.lineTo(x - 8, y); ctx.closePath(); ctx.fill();
    if (node.placeholder) {
      ctx.strokeStyle = 'rgba(255,159,67,0.55)'; ctx.lineWidth = 1; ctx.setLineDash([3, 3]);
      ctx.beginPath(); ctx.arc(x, y, 12, 0, Math.PI * 2); ctx.stroke(); ctx.setLineDash([]);
    }
  });
  roomSubjects().forEach((subject, index) => {
    const x = mapX(subject.x);
    const y = mapZ(subject.z);
    ctx.strokeStyle = '#00e676';
    ctx.lineWidth = 2;
    ctx.beginPath(); ctx.arc(x, y, 10 + subject.intensity * 6, 0, Math.PI * 2); ctx.stroke();
    ctx.fillStyle = '#c8d8f0';
    ctx.fillText(String(index + 1), x + 12, y - 10);
  });
}

function drawElevation(canvasId, axis) {
  const { ctx, width, height } = fitCanvas($(canvasId));
  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = '#030510';
  ctx.fillRect(0, 0, width, height);
  drawRoomRect(ctx, width, height);
  const horizontal = axis === 'front' ? STATE.room.width : STATE.room.depth;
  const mapX = (value) => 18 + (value / horizontal) * (width - 36);
  const mapY = (value) => height - 18 - (value / STATE.room.height) * (height - 36);
  const bins = new Array(20).fill(0);
  STATE.hotspots.forEach((spot) => {
    const pos = axis === 'front' ? spot.x : spot.z;
    const bin = Math.max(0, Math.min(bins.length - 1, Math.floor((pos / horizontal) * bins.length)));
    bins[bin] = Math.max(bins[bin], spot.intensity);
  });
  bins.forEach((value, index) => {
    if (!value) return;
    const x0 = 18 + (index / bins.length) * (width - 36);
    const barW = (width - 36) / bins.length;
    const y0 = mapY(value * Math.min(STATE.room.height * 0.7, 1.8));
    ctx.fillStyle = `rgba(0,136,255,${0.12 + value * 0.34})`;
    ctx.fillRect(x0, y0, barW - 1, height - 18 - y0);
  });
  mergedNodes().forEach((node) => {
    const px = mapX(axis === 'front' ? node.x : node.z);
    const py = mapY(node.y);
    ctx.fillStyle = node.placeholder ? '#ff9f43' : (node.status === 'active' ? '#00ccff' : '#5577aa');
    ctx.fillRect(px - 5, py - 5, 10, 10);
  });
  roomSubjects().forEach((subject, index) => {
    const px = mapX(axis === 'front' ? subject.x : subject.z);
    const py = mapY(subject.y);
    ctx.strokeStyle = '#00e676';
    ctx.lineWidth = 2;
    ctx.beginPath(); ctx.arc(px, py, 8 + subject.intensity * 5, 0, Math.PI * 2); ctx.stroke();
    ctx.fillStyle = '#c8d8f0';
    ctx.fillText(String(index + 1), px + 10, py - 10);
  });
}

function drawIso() {
  const { ctx, width, height } = fitCanvas($('canvas-iso'));
  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = '#030510';
  ctx.fillRect(0, 0, width, height);
  const scale = Math.min(width / (STATE.room.width + STATE.room.depth), height / (STATE.room.height + STATE.room.depth * 0.8)) * 0.85;
  const cx = width * 0.48;
  const cy = height * 0.78;
  const project = (x, y, z) => ([
    cx + (x - STATE.room.width / 2) * scale - (z - STATE.room.depth / 2) * scale * 0.65,
    cy - y * scale * 0.92 + (z - STATE.room.depth / 2) * scale * 0.36,
  ]);
  const corners = [
    [0, 0, 0], [STATE.room.width, 0, 0], [STATE.room.width, 0, STATE.room.depth], [0, 0, STATE.room.depth],
    [0, STATE.room.height, 0], [STATE.room.width, STATE.room.height, 0], [STATE.room.width, STATE.room.height, STATE.room.depth], [0, STATE.room.height, STATE.room.depth],
  ].map(([x, y, z]) => project(x, y, z));
  ctx.strokeStyle = 'rgba(0,136,255,0.24)';
  ctx.lineWidth = 1.2;
  [[0,1],[1,2],[2,3],[3,0],[4,5],[5,6],[6,7],[7,4],[0,4],[1,5],[2,6],[3,7]].forEach(([a, b]) => {
    ctx.beginPath(); ctx.moveTo(...corners[a]); ctx.lineTo(...corners[b]); ctx.stroke();
  });
  STATE.hotspots.forEach((spot) => {
    const base = project(spot.x, 0.02, spot.z);
    const top = project(spot.x, spot.y, spot.z);
    ctx.strokeStyle = `rgba(255,159,67,${0.22 + spot.intensity * 0.45})`;
    ctx.lineWidth = 4 + spot.intensity * 4;
    ctx.beginPath(); ctx.moveTo(...base); ctx.lineTo(...top); ctx.stroke();
  });
  mergedNodes().forEach((node) => {
    const [x, y] = project(node.x, node.y, node.z);
    ctx.fillStyle = node.placeholder ? '#ff9f43' : (node.status === 'active' ? '#00ccff' : '#5577aa');
    ctx.beginPath(); ctx.arc(x, y, 7, 0, Math.PI * 2); ctx.fill();
    ctx.fillStyle = '#c8d8f0';
    ctx.fillText(`N${node.node_id}`, x + 10, y - 10);
  });
  roomSubjects().forEach((subject, index) => {
    const [x, y] = project(subject.x, subject.y, subject.z);
    ctx.strokeStyle = '#00e676';
    ctx.lineWidth = 2;
    ctx.beginPath(); ctx.arc(x, y, 9 + subject.intensity * 4, 0, Math.PI * 2); ctx.stroke();
    ctx.fillStyle = '#00e676';
    ctx.fillRect(x - 1, y - 18, 2, 36);
    ctx.fillStyle = '#c8d8f0';
    ctx.fillText(String(index + 1), x + 12, y - 10);
  });
}

function renderAtlas() {
  drawPlan();
  drawElevation('canvas-front', 'front');
  drawElevation('canvas-side', 'side');
  drawIso();
  $('iso-caption').textContent = `${mergedNodes().length} nodes · ${roomSubjects().length} anchors · ${STATE.hotspots.length} hotspots`;
}

function renderAll() {
  renderHero();
  renderMetrics();
  renderEvidence();
  renderNodes();
  renderObjects();
  renderAtlas();
  renderCalibrationBanner();
}

/**
 * Show a "needs calibration" banner above the atlas the first time the
 * UI sees stacked nodes. Includes a one-click "Apply default layout"
 * button that POSTs perimeter coordinates to /api/v1/nodes/positions.
 */
function renderCalibrationBanner() {
  const banner = $('atlas-calibration');
  if (!banner) return;
  const nodes = mergedNodes();
  const stacked = nodes.some(n => n.placeholder);
  if (!stacked && STATE.room.calibrated) {
    banner.style.display = 'none';
    return;
  }
  banner.style.display = '';
  const live = nodes.filter(n => n.status !== 'offline').length;
  $('atlas-calibration-msg').textContent = stacked
    ? `${live} node${live === 1 ? '' : 's'} are reporting the same default position. The atlas is showing a perimeter fallback layout — apply it (or open Provisioning to set positions per-node).`
    : 'Room is not calibrated yet. Apply the default perimeter layout to get a usable atlas in one click.';
}

async function applyDefaultLayout() {
  const nodes = mergedNodes();
  const live = nodes.filter(n => n.status !== 'offline');
  const target = live.length || nodes.length;
  if (target === 0) {
    toast.warn('No nodes to position', { detail: 'Wait for at least one node to come online.' });
    return;
  }
  const layout = defaultPerimeterLayout(target, STATE.room);
  const payload = layout.map(([x, y, z]) => ({ x, y, z }));
  const stop = busy(`Posting ${payload.length} node positions…`);
  try {
    const data = await fetchJSON(POSITIONS_URL, {
      silent: true,
      fetchOpts: {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      },
    });
    stop({ ok: `Updated ${data.updated ?? payload.length} positions` });
    refreshRoom();
  } catch (e) {
    stop({ err: 'Layout update failed', detail: e.message });
    stickyError('Could not POST /api/v1/nodes/positions', { detail: e.message });
  }
}
window.applyDefaultLayout = applyDefaultLayout;

/**
 * Ask the server to auto-derive node positions from its own telemetry
 * (active node IDs → deterministic perimeter layout). Stable across
 * reboots as long as node IDs don't change. Mix-tolerant: nodes can be
 * added or removed and the next call will re-layout.
 */
async function autoCalibrate() {
  const stop = busy('Auto-calibrating from live telemetry…');
  try {
    const data = await fetchJSON(AUTO_CAL_URL, {
      silent: true,
      fetchOpts: { method: 'POST', headers: { 'Content-Type': 'application/json' } },
    });
    if (data.status === 'no_active_nodes') {
      stop({ err: 'No active nodes', detail: 'Wait for nodes to start streaming.' });
      return;
    }
    stop({ ok: `Auto-calibrated ${data.active_nodes} active node${data.active_nodes === 1 ? '' : 's'}` });
    refreshRoom();
  } catch (e) {
    stop({ err: 'Auto-calibration failed', detail: e.message });
    stickyError('Could not POST /api/v1/calibration/auto', { detail: e.message });
  }
}
window.autoCalibrate = autoCalibrate;

function applyRoomPayload(payload) {
  if (!payload) return;
  if (payload.room) {
    STATE.room.width = payload.room.width || STATE.room.width;
    STATE.room.depth = payload.room.depth || STATE.room.depth;
    STATE.room.height = payload.room.height || STATE.room.height;
  }
  STATE.room.nodePositions = payload.node_positions || [];
  STATE.room.activeNodes = payload.active_nodes || [];
  STATE.room.calibrated = !!payload.calibrated;
}

function applySensingPayload(payload, connection = 'live') {
  if (!payload || (payload.type && payload.type !== 'sensing_update')) return;
  STATE.data = payload;
  STATE.hotspots = extractHotspots(payload.signal_field, 6);
  STATE.lastUpdateMs = Date.now();
  STATE.connection = connection;
  renderAll();
}

async function refreshRoom() {
  try {
    const response = await fetch(ROOM_URL);
    if (!response.ok) return;
    applyRoomPayload(await response.json());
    renderAll();
  } catch (_) {}
}

async function fetchLatest() {
  try {
    const response = await fetch(LATEST_URL);
    if (!response.ok) return;
    applySensingPayload(await response.json(), 'snapshot');
  } catch (_) {}
}

function connect() {
  if (STATE.ws && STATE.ws.readyState <= 1) return;
  STATE.connection = 'connecting';
  renderHero();
  try {
    STATE.ws = new WebSocket(WS_URL);
  } catch (_) {
    setTimeout(connect, 2000);
    return;
  }
  STATE.ws.onmessage = (event) => {
    try { applySensingPayload(JSON.parse(event.data), 'live'); } catch (_) {}
  };
  STATE.ws.onclose = () => {
    STATE.connection = STATE.lastUpdateMs ? 'stale' : 'connecting';
    renderHero();
    setTimeout(connect, 2000);
  };
  STATE.ws.onerror = () => {
    try { STATE.ws.close(); } catch (_) {}
  };
}

window.addEventListener('resize', renderAtlas);
setInterval(() => {
  if (STATE.lastUpdateMs && Date.now() - STATE.lastUpdateMs > 6000 && STATE.connection === 'live') {
    STATE.connection = 'stale';
  }
  renderHero();
}, 1000);
setInterval(refreshRoom, 10000);

fetchLatest();
refreshRoom();
connect();
renderAll();