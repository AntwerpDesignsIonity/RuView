/**
 * Sensing WebSocket Service
 *
 * Manages the connection to the Python sensing WebSocket server
 * (ws://localhost:8765) and provides a callback-based API for the UI.
 *
 * Falls back to simulated data only after MAX_RECONNECT_ATTEMPTS exhausted.
 * While reconnecting the service stays in "reconnecting" state and does NOT
 * emit simulated frames so the UI can clearly distinguish live vs. fallback data.
 */

// Derive WebSocket URL from the page origin so it works on any host.
// The Python sensing WebSocket server listens on port 8765.
const _wsProto = (typeof window !== 'undefined' && window.location.protocol === 'https:') ? 'wss:' : 'ws:';
const _wsHostname = (typeof window !== 'undefined' && window.location.hostname) ? window.location.hostname : 'localhost';
const SENSING_WS_URL = `${_wsProto}//${_wsHostname}:3001/ws/sensing`;
const RECONNECT_DELAYS = [1000, 2000, 4000, 8000, 16000];
const MAX_RECONNECT_ATTEMPTS = 20;
// Number of failed attempts that must occur before simulation starts.
// This prevents the UI from flashing "SIMULATED" on a brief hiccup.
const SIM_FALLBACK_AFTER_ATTEMPTS = 5;
const SIMULATION_INTERVAL = 500; // ms
const HEALTH_CHECK_INTERVAL = 15000; // ms
const WATCHDOG_INTERVAL = 1000; // ms
const STALE_FRAME_WARN_MS = 7000;
const STALE_FRAME_RECOVER_MS = 20000;
const AUTO_RECOVER_COOLDOWN_MS = 20000;

class SensingService {
  constructor() {
    /** @type {WebSocket|null} */
    this._ws = null;
    this._listeners = new Set();
    this._stateListeners = new Set();
    this._diagnosticsListeners = new Set();
    this._reconnectAttempt = 0;
    this._reconnectTimer = null;
    this._simTimer = null;
    this._healthTimer = null;
    this._watchdogTimer = null;
    this._healthInFlight = false;
    this._isRunning = false;
    this._isStopping = false;
    this._lastAutoRecoverAt = 0;
    // Connection state: disconnected | connecting | connected | reconnecting | simulated
    this._state = 'disconnected';
    // Data-source label exposed to the UI:
    //   "live"              — real ESP32 hardware connected
    //   "server-simulated"  — server is running but using synthetic data (no hardware)
    //   "reconnecting"      — WebSocket disconnected, retrying
    //   "simulated"         — client-side fallback simulation (server unreachable)
    this._dataSource = 'reconnecting';
    // The raw source string from the server (e.g. "esp32", "simulated", "simulate")
    this._serverSource = null;
    this._lastMessage = null;
    this._lastFrameAt = 0;

    this._boundOnlineHandler = null;
    this._boundOfflineHandler = null;
    this._boundVisibilityHandler = null;

    this._diagnostics = {
      wsUrl: SENSING_WS_URL,
      state: this._state,
      dataSource: this._dataSource,
      serverSource: this._serverSource,
      reconnectAttempt: this._reconnectAttempt,
      maxReconnectAttempts: MAX_RECONNECT_ATTEMPTS,
      lastFrameAt: null,
      lastFrameAgeMs: null,
      networkOnline: typeof navigator !== 'undefined' ? navigator.onLine : true,
      visibility: typeof document !== 'undefined' ? document.visibilityState : 'visible',
      health: {
        status: 'unknown', // unknown | healthy | degraded | down | stale
        latencyMs: null,
        httpStatus: null,
        error: null,
        lastCheckedAt: null,
      },
    };

    // Ring buffer of recent RSSI values for sparkline
    this._rssiHistory = [];
    this._maxHistory = 60;
  }

  // ---- Public API --------------------------------------------------------

  /** Start the service (connect or simulate). */
  start() {
    if (this._isRunning) return;
    this._isRunning = true;
    this._isStopping = false;
    this._attachBrowserListeners();
    this._startRunners();
    this._connect();
  }

  /** Stop the service entirely. */
  stop() {
    this._isRunning = false;
    this._isStopping = true;
    this._detachBrowserListeners();
    this._clearTimers();
    if (this._ws) {
      this._ws.close(1000, 'client stop');
      this._ws = null;
    }
    this._setState('disconnected');
    this._setDataSource('reconnecting');
    this._emitDiagnostics();
  }

  /** Force a reconnect cycle immediately. */
  forceReconnect() {
    if (!this._isRunning) return;
    this._clearReconnectTimer();
    this._stopSimulation();
    if (this._ws) {
      try {
        this._ws.close(4001, 'force reconnect');
      } catch {
        this._ws = null;
      }
    }
    this._setState('reconnecting');
    this._setDataSource('reconnecting');
    this._connect();
  }

  /** Register a callback for sensing data updates. Returns unsubscribe fn. */
  onData(callback) {
    this._listeners.add(callback);
    // Immediately push last known data if available
    if (this._lastMessage) callback(this._lastMessage);
    return () => this._listeners.delete(callback);
  }

  /** Register a callback for connection state changes. Returns unsubscribe fn. */
  onStateChange(callback) {
    this._stateListeners.add(callback);
    callback(this._state);
    return () => this._stateListeners.delete(callback);
  }

  /** Register a callback for diagnostics updates. Returns unsubscribe fn. */
  onDiagnostics(callback) {
    this._diagnosticsListeners.add(callback);
    callback(this.diagnostics);
    return () => this._diagnosticsListeners.delete(callback);
  }

  /** Get the RSSI sparkline history (array of floats). */
  getRssiHistory() {
    return [...this._rssiHistory];
  }

  /** Get per-node RSSI history (object keyed by node_id). */
  getPerNodeRssiHistory() {
    return { ...(this._perNodeRssiHistory || {}) };
  }

  /** Current connection state. */
  get state() {
    return this._state;
  }

  /**
   * Current data source label.
    * "live"              — frames are arriving from real ESP32 hardware
    * "hardware-offline"  — server reports ESP32 source but no CSI frames yet
    * "server-simulated"  — server is online but synthesizing data
    * "reconnecting"      — transport degraded, retrying
    * "simulated"         — local client-side fallback simulation
   */
  get dataSource() {
    return this._dataSource;
  }

  /** Snapshot of connection diagnostics for UI cards and debug panels. */
  get diagnostics() {
    return {
      ...this._diagnostics,
      health: { ...this._diagnostics.health },
    };
  }

  // ---- Connection --------------------------------------------------------

  _connect() {
    if (!this._isRunning) return;
    if (this._ws && this._ws.readyState <= WebSocket.OPEN) return;

    this._setState('connecting');

    try {
      this._ws = new WebSocket(SENSING_WS_URL);
    } catch (err) {
      console.warn('[Sensing] WebSocket constructor failed:', err.message);
      this._fallbackToSimulation();
      return;
    }

    this._ws.onopen = () => {
      console.info('[Sensing] Connected to', SENSING_WS_URL);
      this._reconnectAttempt = 0;
      this._stopSimulation();
      this._setState('connected');
      // Don't assume "live" yet — wait for first frame's source field.
      // Fetch server status to determine actual data source immediately.
      this._detectServerSource('ws-open');
      this._emitDiagnostics();
    };

    this._ws.onmessage = (evt) => {
      try {
        const data = JSON.parse(evt.data);
        this._handleData(data);
      } catch (e) {
        console.warn('[Sensing] Invalid message:', e.message);
      }
    };

    this._ws.onerror = () => {
      // onerror is always followed by onclose, so we handle reconnect there
    };

    this._ws.onclose = (evt) => {
      console.info('[Sensing] Connection closed (code=%d)', evt.code);
      this._ws = null;
      if (this._isStopping || !this._isRunning || evt.code === 1000) {
        this._setState('disconnected');
        this._emitDiagnostics();
        return;
      }
      if (evt.code !== 1000) {
        this._scheduleReconnect();
      } else {
        this._setState('disconnected');
        this._setDataSource('reconnecting');
      }
    };
  }

  _scheduleReconnect() {
    if (this._reconnectAttempt >= MAX_RECONNECT_ATTEMPTS) {
      console.warn('[Sensing] Max reconnect attempts (%d) reached, switching to simulation', MAX_RECONNECT_ATTEMPTS);
      this._fallbackToSimulation();
      return;
    }

    const delay = RECONNECT_DELAYS[Math.min(this._reconnectAttempt, RECONNECT_DELAYS.length - 1)];
    this._reconnectAttempt++;
    console.info('[Sensing] Reconnecting in %dms (attempt %d/%d)', delay, this._reconnectAttempt, MAX_RECONNECT_ATTEMPTS);

    this._setState('reconnecting');
    this._setDataSource('reconnecting');

    this._clearReconnectTimer();
    this._reconnectTimer = setTimeout(() => {
      this._reconnectTimer = null;
      this._connect();
    }, delay);

    // Only start simulation after several failed attempts so a brief hiccup
    // does not immediately switch the UI to "SIMULATED DATA".
    if (this._reconnectAttempt >= SIM_FALLBACK_AFTER_ATTEMPTS && this._state !== 'simulated') {
      this._fallbackToSimulation();
    }

    this._emitDiagnostics();
  }

  // ---- Simulation fallback -----------------------------------------------

  _fallbackToSimulation() {
    this._setState('simulated');
    this._setDataSource('simulated');
    if (this._simTimer) return; // already running
    console.info('[Sensing] Running in simulation mode');

    this._simTimer = setInterval(() => {
      const data = this._generateSimulatedData();
      this._handleData(data);
    }, SIMULATION_INTERVAL);

    this._emitDiagnostics();
  }

  _stopSimulation() {
    if (this._simTimer) {
      clearInterval(this._simTimer);
      this._simTimer = null;
    }
  }

  _generateSimulatedData() {
    const t = Date.now() / 1000;
    const baseRssi = -45;
    const variance = 1.5 + Math.sin(t * 0.1) * 1.0;
    const motionBand = 0.05 + Math.abs(Math.sin(t * 0.3)) * 0.15;
    const breathBand = 0.03 + Math.abs(Math.sin(t * 0.05)) * 0.08;
    const isPresent = variance > 0.8;
    const isActive = motionBand > 0.12;

    // Generate signal field
    const gridSize = 20;
    const values = [];
    for (let iz = 0; iz < gridSize; iz++) {
      for (let ix = 0; ix < gridSize; ix++) {
        const cx = gridSize / 2, cy = gridSize / 2;
        const dist = Math.sqrt((ix - cx) ** 2 + (iz - cy) ** 2);
        let v = Math.max(0, 1 - dist / (gridSize * 0.7)) * 0.3;
        // Body blob
        const bx = cx + 3 * Math.sin(t * 0.2);
        const by = cy + 2 * Math.cos(t * 0.15);
        const bodyDist = Math.sqrt((ix - bx) ** 2 + (iz - by) ** 2);
        if (isPresent) {
          v += Math.exp(-bodyDist * bodyDist / 8) * (0.3 + motionBand * 3);
        }
        values.push(Math.min(1, Math.max(0, v + Math.random() * 0.05)));
      }
    }

    return {
      type: 'sensing_update',
      timestamp: t,
      source: 'simulated',
      // Explicit machine-readable marker so the UI can always detect simulated
      // frames regardless of which code path produced them.
      _simulated: true,
      nodes: [{
        node_id: 1,
        rssi_dbm: baseRssi + Math.sin(t * 0.5) * 3,
        position: [2, 0, 1.5],
        amplitude: [],
        subcarrier_count: 0,
      }],
      features: {
        mean_rssi: baseRssi + Math.sin(t * 0.5) * 3,
        variance,
        std: Math.sqrt(variance),
        motion_band_power: motionBand,
        breathing_band_power: breathBand,
        dominant_freq_hz: 0.3 + Math.sin(t * 0.02) * 0.1,
        change_points: Math.floor(Math.random() * 3),
        spectral_power: motionBand + breathBand + Math.random() * 0.1,
        range: variance * 3,
        iqr: variance * 1.5,
        skewness: (Math.random() - 0.5) * 0.5,
        kurtosis: Math.random() * 2,
      },
      classification: {
        motion_level: isActive ? 'active' : (isPresent ? 'present_still' : 'absent'),
        presence: isPresent,
        confidence: isPresent ? 0.75 + Math.random() * 0.2 : 0.5 + Math.random() * 0.3,
      },
      signal_field: {
        grid_size: [gridSize, 1, gridSize],
        values,
      },
    };
  }

  // ---- Server source detection -------------------------------------------

  /**
    * Fetch `/health` to find out if the server is using real
   * hardware or simulation. Called once on WebSocket open.
   */
  async _detectServerSource(reason = 'manual') {
    if (this._healthInFlight) return;
    this._healthInFlight = true;
    const start = Date.now();
    try {
      const resp = await fetch('/health', { cache: 'no-store' });
      const latencyMs = Date.now() - start;
      if (resp.ok) {
        const json = await resp.json();
        this._applyServerSource(json.source);
        this._diagnostics.health = {
          status: reason === 'watchdog' ? 'degraded' : 'healthy',
          latencyMs,
          httpStatus: resp.status,
          error: null,
          lastCheckedAt: Date.now(),
        };
      } else {
        // Can't reach health endpoint — assume live until first frame tells us
        this._diagnostics.health = {
          status: 'degraded',
          latencyMs,
          httpStatus: resp.status,
          error: `HTTP ${resp.status}`,
          lastCheckedAt: Date.now(),
        };
      }
    } catch (err) {
      this._diagnostics.health = {
        status: 'down',
        latencyMs: null,
        httpStatus: null,
        error: err?.message || 'health-check-failed',
        lastCheckedAt: Date.now(),
      };
      if (this._state === 'connected') {
        this._setDataSource('reconnecting');
      }
    } finally {
      this._healthInFlight = false;
      this._emitDiagnostics();
    }
  }

  /**
   * Map a raw server source string to the UI data-source label.
   */
  _applyServerSource(rawSource) {
    this._serverSource = rawSource;
    if (rawSource === 'esp32' || rawSource === 'wifi' || rawSource === 'live') {
      this._setDataSource('live');
    } else if (rawSource === 'esp32:offline') {
      // Hardware provisioned but no UDP frames yet — waiting for ESP32 WiFi or bridge
      this._setDataSource('hardware-offline');
    } else if (rawSource === 'simulated' || rawSource === 'simulate') {
      this._setDataSource('server-simulated');
    } else {
      // Unknown source — show as server-simulated to be safe
      this._setDataSource('server-simulated');
    }
    this._emitDiagnostics();
  }

  /** @return {string|null} Raw server source (e.g. "esp32", "simulated") */
  get serverSource() {
    return this._serverSource;
  }

  // ---- Data handling -----------------------------------------------------

  _handleData(data) {
    this._lastMessage = data;
    this._lastFrameAt = Date.now();

    // Track the server's source field from each frame so the UI
    // can react if the server switches between esp32 ↔ simulated at runtime.
    if (data.source && this._state === 'connected') {
      const raw = data.source;
      if (raw !== this._serverSource) {
        this._applyServerSource(raw);
      }
    }

    // Update RSSI history for sparkline
    if (data.features && data.features.mean_rssi != null) {
      this._rssiHistory.push(data.features.mean_rssi);
      if (this._rssiHistory.length > this._maxHistory) {
        this._rssiHistory.shift();
      }
    }

    // Per-node RSSI tracking
    if (!this._perNodeRssiHistory) this._perNodeRssiHistory = {};
    if (data.nodes) {
      for (const nf of data.nodes) {
        if (!this._perNodeRssiHistory[nf.node_id]) {
          this._perNodeRssiHistory[nf.node_id] = [];
        }
        this._perNodeRssiHistory[nf.node_id].push(nf.rssi_dbm);
        if (this._perNodeRssiHistory[nf.node_id].length > this._maxHistory) {
          this._perNodeRssiHistory[nf.node_id].shift();
        }
      }
    }

    // Notify all listeners
    for (const cb of this._listeners) {
      try {
        cb(data);
      } catch (e) {
        console.error('[Sensing] Listener error:', e);
      }
    }

    this._emitDiagnostics();
  }

  // ---- State management --------------------------------------------------

  _setState(newState) {
    if (newState === this._state) return;
    this._state = newState;
    for (const cb of this._stateListeners) {
      try { cb(newState); } catch (e) { /* ignore */ }
    }
    this._emitDiagnostics();
  }

  /**
   * Update the dataSource label and notify state listeners so the UI can
   * react without needing a separate subscription.
   * @param {'live'|'hardware-offline'|'server-simulated'|'reconnecting'|'simulated'} source
   */
  _setDataSource(source) {
    if (source === this._dataSource) return;
    this._dataSource = source;
    for (const cb of this._stateListeners) {
      try { cb(this._state); } catch (e) { /* ignore */ }
    }
    this._emitDiagnostics();
  }

  _emitDiagnostics() {
    const lastFrameAgeMs = this._lastFrameAt > 0 ? Date.now() - this._lastFrameAt : null;
    const health = { ...this._diagnostics.health };

    if (this._state === 'connected' && Number.isFinite(lastFrameAgeMs) && lastFrameAgeMs > STALE_FRAME_WARN_MS) {
      if (health.status === 'healthy' || health.status === 'unknown') {
        health.status = 'stale';
      }
    }

    this._diagnostics = {
      ...this._diagnostics,
      wsUrl: SENSING_WS_URL,
      state: this._state,
      dataSource: this._dataSource,
      serverSource: this._serverSource,
      reconnectAttempt: this._reconnectAttempt,
      maxReconnectAttempts: MAX_RECONNECT_ATTEMPTS,
      lastFrameAt: this._lastFrameAt || null,
      lastFrameAgeMs,
      networkOnline: typeof navigator !== 'undefined' ? navigator.onLine : true,
      visibility: typeof document !== 'undefined' ? document.visibilityState : 'visible',
      health,
    };

    for (const cb of this._diagnosticsListeners) {
      try { cb(this.diagnostics); } catch (e) { /* ignore */ }
    }
  }

  _startRunners() {
    if (!this._healthTimer) {
      this._healthTimer = setInterval(() => {
        if (!this._isRunning) return;
        this._detectServerSource('interval');
      }, HEALTH_CHECK_INTERVAL);
    }

    if (!this._watchdogTimer) {
      this._watchdogTimer = setInterval(() => {
        if (!this._isRunning) return;

        const lastFrameAgeMs = this._lastFrameAt > 0 ? Date.now() - this._lastFrameAt : null;
        if (this._state === 'connected' && Number.isFinite(lastFrameAgeMs) && lastFrameAgeMs > STALE_FRAME_WARN_MS) {
          this._emitDiagnostics();
        }

        if (
          this._state === 'connected' &&
          Number.isFinite(lastFrameAgeMs) &&
          lastFrameAgeMs > STALE_FRAME_RECOVER_MS &&
          Date.now() - this._lastAutoRecoverAt > AUTO_RECOVER_COOLDOWN_MS
        ) {
          this._lastAutoRecoverAt = Date.now();
          this._detectServerSource('watchdog');
          this.forceReconnect();
        }
      }, WATCHDOG_INTERVAL);
    }

    this._emitDiagnostics();
  }

  _attachBrowserListeners() {
    if (typeof window !== 'undefined' && !this._boundOnlineHandler) {
      this._boundOnlineHandler = () => {
        this._emitDiagnostics();
        if (this._isRunning && this._state !== 'connected') {
          this.forceReconnect();
        }
      };
      this._boundOfflineHandler = () => {
        this._setDataSource('reconnecting');
        this._emitDiagnostics();
      };
      window.addEventListener('online', this._boundOnlineHandler);
      window.addEventListener('offline', this._boundOfflineHandler);
    }

    if (typeof document !== 'undefined' && !this._boundVisibilityHandler) {
      this._boundVisibilityHandler = () => {
        this._emitDiagnostics();
        if (document.visibilityState === 'visible' && this._isRunning && this._state === 'connected') {
          this._detectServerSource('visibility');
        }
      };
      document.addEventListener('visibilitychange', this._boundVisibilityHandler);
    }
  }

  _detachBrowserListeners() {
    if (typeof window !== 'undefined' && this._boundOnlineHandler) {
      window.removeEventListener('online', this._boundOnlineHandler);
      window.removeEventListener('offline', this._boundOfflineHandler);
      this._boundOnlineHandler = null;
      this._boundOfflineHandler = null;
    }

    if (typeof document !== 'undefined' && this._boundVisibilityHandler) {
      document.removeEventListener('visibilitychange', this._boundVisibilityHandler);
      this._boundVisibilityHandler = null;
    }
  }

  _clearReconnectTimer() {
    if (this._reconnectTimer) {
      clearTimeout(this._reconnectTimer);
      this._reconnectTimer = null;
    }
  }

  _clearTimers() {
    this._stopSimulation();
    this._clearReconnectTimer();

    if (this._healthTimer) {
      clearInterval(this._healthTimer);
      this._healthTimer = null;
    }

    if (this._watchdogTimer) {
      clearInterval(this._watchdogTimer);
      this._watchdogTimer = null;
    }
  }
}

export const sensingService = new SensingService();