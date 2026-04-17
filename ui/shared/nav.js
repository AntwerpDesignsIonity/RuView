/**
 * Ionity AEDI-S — Shared Navigation Bar
 * Auto-detects current page and highlights active link.
 * Auto-connects to health endpoint and shows live status.
 */
export function injectNav(containerId = 'global-nav') {
  const page = location.pathname.split('/').pop() || 'index.html';
  const links = [
    { href: 'index.html',       label: 'Dashboard',  icon: '◉' },
    { href: 'monitor.html',     label: 'Monitor',    icon: '◎' },
    { href: 'observatory.html', label: 'Observatory', icon: '◈' },
    { href: 'pose-fusion.html', label: 'Pose Fusion', icon: '⬡' },
    { href: 'xyz-segment.html', label: 'XYZ Segment', icon: '◆' },
    { href: 'aedi.html',        label: 'AEDI',       icon: '⬢' },
  ];

  const nav = document.getElementById(containerId);
  if (!nav) return;

  nav.innerHTML = `
    <a class="gnav-brand" href="index.html"><img src="aedi-logo.svg" alt="AEDI-S" style="height:22px;vertical-align:middle;margin-right:6px">AEDI-S</a>
    <span class="gnav-sep">│</span>
    ${links.map(l =>
      `<a class="gnav-link${page === l.href ? ' active' : ''}" href="${l.href}">${l.icon} ${l.label}</a>`
    ).join('')}
    <div class="gnav-status">
      <span class="dot dot-dim" id="nav-health-dot"></span>
      <span id="nav-source-label">--</span>
      <span id="nav-tick-label">t:--</span>
    </div>
  `;

  // Poll health
  async function poll() {
    try {
      const r = await fetch('/health');
      const d = await r.json();
      const dot = document.getElementById('nav-health-dot');
      const src = document.getElementById('nav-source-label');
      const tick = document.getElementById('nav-tick-label');
      if (dot) { dot.className = 'dot ' + (d.status === 'ok' ? 'dot-green' : 'dot-red'); }
      if (src) src.textContent = d.source || '--';
      if (tick) tick.textContent = `t:${d.tick || 0}`;
    } catch { /* offline */ }
  }
  poll();
  setInterval(poll, 3000);
}

/** Shared nav bar CSS (inject into <style> or <head>) */
export const NAV_CSS = `
#global-nav {
  display: flex; align-items: center; gap: 0;
  background: #070718; border-bottom: 1px solid #1a2040;
  height: 42px; padding: 0 12px; flex-shrink: 0;
  overflow-x: auto; overflow-y: hidden;
}
.gnav-brand {
  font-size: 15px; font-weight: bold; color: #00ccff;
  letter-spacing: 2px; margin-right: 16px; white-space: nowrap;
  text-decoration: none;
}
.gnav-sep { color: #334466; margin: 0 4px; }
.gnav-link {
  display: inline-block; padding: 6px 12px; color: #5577aa;
  text-decoration: none; border-radius: 4px; white-space: nowrap;
  transition: color 0.15s, background 0.15s; font-size: 12px;
}
.gnav-link:hover { color: #c8d8f0; background: #111230; }
.gnav-link.active { color: #00ccff; background: #0d1535; border-bottom: 2px solid #00ccff; }
.gnav-status {
  margin-left: auto; display: flex; align-items: center; gap: 10px;
  font-size: 11px; color: #5577aa; flex-shrink: 0;
}
.dot { width: 8px; height: 8px; border-radius: 50%; display: inline-block; }
.dot-green { background: #00e676; box-shadow: 0 0 6px #00e676; }
.dot-red   { background: #ff3d57; box-shadow: 0 0 6px #ff3d57; }
.dot-dim   { background: #334466; }
`;
