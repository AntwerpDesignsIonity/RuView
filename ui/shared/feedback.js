/**
 * Ionity AEDI-S — Shared Feedback / Toast Library
 *
 * Lightweight, dependency-free toast + busy + sticky-error system.
 * Use from any page after `injectFeedback()` mounts the host element.
 *
 *   import { toast, busy, stickyError, fetchJSON } from './shared/feedback.js';
 *
 *   toast.success('Saved');
 *   toast.error('Could not reach node');
 *   const stop = busy('Connecting…');
 *   stop();
 *   stickyError('UDP listener failed', { detail: 'EADDRINUSE :5005' });
 *   const data = await fetchJSON('/api/v1/registry');   // auto error toast
 */

let host = null;
let stickyHost = null;

const STYLE_ID = 'aedi-feedback-style';
const FEEDBACK_CSS = `
#aedi-toast-host {
  position: fixed; top: 56px; right: 16px; z-index: 99999;
  display: flex; flex-direction: column; gap: 8px;
  pointer-events: none; max-width: 380px;
}
.aedi-toast {
  background: #0d1535; color: #c8d8f0;
  border: 1px solid #1a2040; border-left-width: 4px;
  padding: 10px 14px; border-radius: 4px;
  font: 12px/1.4 system-ui, -apple-system, "Segoe UI", sans-serif;
  box-shadow: 0 4px 12px rgba(0,0,0,0.5);
  pointer-events: auto;
  animation: aedi-toast-in 0.2s ease-out;
  cursor: pointer;
}
.aedi-toast.success { border-left-color: #00e676; }
.aedi-toast.error   { border-left-color: #ff5252; }
.aedi-toast.warn    { border-left-color: #ffb300; }
.aedi-toast.info    { border-left-color: #00ccff; }
.aedi-toast.busy    { border-left-color: #00ccff; opacity: 0.85; }
.aedi-toast .aedi-toast-title { font-weight: 600; margin-bottom: 2px; color: #fff; }
.aedi-toast .aedi-toast-detail { font-size: 11px; color: #8899bb; margin-top: 4px; word-break: break-word; }
.aedi-toast.fadeout { animation: aedi-toast-out 0.25s ease-in forwards; }
@keyframes aedi-toast-in  { from {opacity:0; transform: translateX(20px);} to {opacity:1; transform:none;} }
@keyframes aedi-toast-out { to   {opacity:0; transform: translateX(20px);} }

#aedi-sticky-host {
  position: fixed; bottom: 0; left: 0; right: 0; z-index: 99998;
  display: flex; flex-direction: column; gap: 1px;
  background: transparent;
}
.aedi-sticky {
  background: #2a0a0a; color: #ffd0d0;
  border-top: 1px solid #ff5252;
  padding: 8px 16px;
  font: 12px/1.4 system-ui, -apple-system, "Segoe UI", sans-serif;
  display: flex; align-items: center; gap: 12px;
}
.aedi-sticky .aedi-sticky-msg { flex: 1; }
.aedi-sticky .aedi-sticky-detail { font-size: 11px; color: #ff8888; margin-left: 8px; opacity: 0.85; }
.aedi-sticky button {
  background: #441515; color: #ffd0d0; border: 1px solid #ff5252;
  padding: 3px 10px; border-radius: 3px; cursor: pointer; font: inherit;
}
.aedi-sticky button:hover { background: #ff5252; color: #fff; }
`;

function ensureMounted() {
  if (host && stickyHost) return;
  if (!document.getElementById(STYLE_ID)) {
    const s = document.createElement('style');
    s.id = STYLE_ID;
    s.textContent = FEEDBACK_CSS;
    document.head.appendChild(s);
  }
  if (!host) {
    host = document.createElement('div');
    host.id = 'aedi-toast-host';
    document.body.appendChild(host);
  }
  if (!stickyHost) {
    stickyHost = document.createElement('div');
    stickyHost.id = 'aedi-sticky-host';
    document.body.appendChild(stickyHost);
  }
}

/** Mount the feedback hosts. Idempotent. Call once on page load. */
export function injectFeedback() {
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', ensureMounted, { once: true });
  } else {
    ensureMounted();
  }
}

function buildToast(kind, title, opts = {}) {
  ensureMounted();
  const el = document.createElement('div');
  el.className = `aedi-toast ${kind}`;
  el.innerHTML = `<div class="aedi-toast-title">${escapeHtml(title)}</div>` +
    (opts.detail ? `<div class="aedi-toast-detail">${escapeHtml(opts.detail)}</div>` : '');
  el.addEventListener('click', () => dismiss(el));
  host.appendChild(el);

  const ttl = opts.ttl ?? (kind === 'error' ? 8000 : kind === 'busy' ? 0 : 3500);
  let timer = null;
  if (ttl > 0) timer = setTimeout(() => dismiss(el), ttl);

  return {
    dismiss: () => { if (timer) clearTimeout(timer); dismiss(el); },
    update: (newTitle, newDetail) => {
      const t = el.querySelector('.aedi-toast-title');
      const d = el.querySelector('.aedi-toast-detail');
      if (t && newTitle != null) t.textContent = newTitle;
      if (newDetail != null) {
        if (d) d.textContent = newDetail;
        else {
          const nd = document.createElement('div');
          nd.className = 'aedi-toast-detail';
          nd.textContent = newDetail;
          el.appendChild(nd);
        }
      }
    },
  };
}

function dismiss(el) {
  if (!el || !el.parentNode) return;
  el.classList.add('fadeout');
  setTimeout(() => { if (el.parentNode) el.parentNode.removeChild(el); }, 250);
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[c]));
}

/** Transient toast notifications. */
export const toast = {
  success: (title, opts) => buildToast('success', title, opts),
  error:   (title, opts) => buildToast('error',   title, opts),
  warn:    (title, opts) => buildToast('warn',    title, opts),
  info:    (title, opts) => buildToast('info',    title, opts),
};

/**
 * Show a busy toast that stays until you call its returned `stop()` fn.
 * Returns a function that dismisses + optionally shows a result toast.
 *
 *   const stop = busy('Probing /health…');
 *   try { await fetch('/health'); stop({ ok: 'Healthy' }); }
 *   catch (e) { stop({ err: 'Probe failed', detail: e.message }); }
 */
export function busy(title, opts = {}) {
  const t = buildToast('busy', title, { ...opts, ttl: 0 });
  return (result) => {
    t.dismiss();
    if (!result) return;
    if (result.ok)  toast.success(result.ok,  result.detail ? { detail: result.detail } : {});
    if (result.err) toast.error(result.err,   result.detail ? { detail: result.detail } : {});
  };
}

/**
 * Persistent sticky banner for system-level errors. Includes a
 * "Copy details" button (clipboard) and a dismiss button.
 */
export function stickyError(message, opts = {}) {
  ensureMounted();
  const detail = opts.detail || '';
  const el = document.createElement('div');
  el.className = 'aedi-sticky';
  el.innerHTML = `
    <span>⚠</span>
    <span class="aedi-sticky-msg">${escapeHtml(message)}</span>
    ${detail ? `<span class="aedi-sticky-detail">${escapeHtml(detail)}</span>` : ''}
    <button data-act="copy">Copy details</button>
    <button data-act="dismiss">Dismiss</button>
  `;
  el.querySelector('[data-act="copy"]').addEventListener('click', async () => {
    const text = `[AEDI sticky-error] ${message}\n${detail}\n${new Date().toISOString()}\n${location.href}`;
    try { await navigator.clipboard.writeText(text); toast.success('Copied'); }
    catch { toast.error('Copy failed', { detail: 'Clipboard blocked' }); }
  });
  el.querySelector('[data-act="dismiss"]').addEventListener('click', () => {
    if (el.parentNode) el.parentNode.removeChild(el);
  });
  stickyHost.appendChild(el);
  return { dismiss: () => { if (el.parentNode) el.parentNode.removeChild(el); } };
}

/**
 * Wrapped fetch that returns parsed JSON and emits a toast on failure.
 * `silent: true` to suppress the error toast (caller handles it).
 *
 *   const cfg = await fetchJSON('/api/v1/registry');
 */
export async function fetchJSON(url, { silent = false, fetchOpts = {} } = {}) {
  let res;
  try {
    res = await fetch(url, fetchOpts);
  } catch (e) {
    if (!silent) toast.error(`Network error: ${url}`, { detail: e.message });
    throw e;
  }
  if (!res.ok) {
    const detail = `HTTP ${res.status} ${res.statusText}`;
    if (!silent) toast.error(`Request failed: ${url}`, { detail });
    throw new Error(detail);
  }
  try {
    return await res.json();
  } catch (e) {
    if (!silent) toast.error(`Invalid JSON: ${url}`, { detail: e.message });
    throw e;
  }
}

// Auto-mount when imported as a module
injectFeedback();
