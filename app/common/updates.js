/* BTS Office — update notifications (renderer side)
 * Shows a small card when an update has downloaded, and provides the "Updates" section
 * used on File ▸ Options & about. Talks to updater.js in the main process via window.bts.
 */
(function () {
  const desktop = !!(window.bts && window.bts.isDesktop && window.bts.updateState);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const ic = (n, s) => (window.icon ? window.icon(n, s) : '');
  let state = null;
  const listeners = new Set();

  const css = `
  #bts-update-card { position: fixed; right: 18px; bottom: 40px; z-index: 3500; width: 372px; background: var(--surface); color: var(--ink);
    border: 1px solid var(--line); border-radius: 14px; box-shadow: var(--shadow-pop); padding: 14px 16px 14px; display: flex; gap: 12px;
    animation: bts-up-in .2s ease-out; user-select: none; }
  @keyframes bts-up-in { from { opacity: 0; transform: translateY(10px); } }
  @media (prefers-reduced-motion: reduce) { #bts-update-card { animation: none; } }
  #bts-update-card .u-ic { width: 34px; height: 34px; border-radius: 10px; background: var(--accent-soft); color: var(--accent); display: grid; place-items: center; flex: none; }
  html[data-theme="dark"] #bts-update-card .u-ic { color: color-mix(in srgb, var(--accent) 55%, white); }
  #bts-update-card b { display: block; font-size: 14px; margin-bottom: 2px; }
  #bts-update-card p { margin: 0 0 10px; color: var(--ink-2); line-height: 1.45; font-size: 12.5px; }
  #bts-update-card .u-row { display: flex; gap: 8px; }
  #bts-update-card .btn { height: 30px; padding: 0 11px; white-space: nowrap; }
  #bts-update-card .u-x { position: absolute; right: 8px; top: 8px; }
  .u-notes { white-space: pre-wrap; max-height: 120px; overflow: auto; background: var(--surface-2); border: 1px solid var(--line); border-radius: 8px; padding: 8px 10px; font-size: 12.5px; color: var(--ink-2); margin: 8px 0 0; }
  .u-panel { max-width: 560px; border: 1px solid var(--line); background: var(--surface); border-radius: 12px; padding: 16px 18px; }
  .u-panel .u-line { display: flex; align-items: center; gap: 12px; }
  .u-panel .u-status { color: var(--ink-2); margin-top: 2px; }
  .u-bar { height: 6px; background: var(--surface-3); border-radius: 4px; overflow: hidden; margin-top: 10px; }
  .u-bar i { display: block; height: 100%; background: var(--accent); transition: width .3s; }
  `;
  function injectCSS() { if (document.getElementById('bts-update-css')) return; const st = document.createElement('style'); st.id = 'bts-update-css'; st.textContent = css; document.head.appendChild(st); }

  function statusText(s) {
    if (!s) return '';
    if (!s.supported) return 'Automatic updates aren’t available for this copy (developer build or package-managed install).';
    switch (s.status) {
      case 'checking': return 'Checking for updates…';
      case 'downloading': return `Downloading version ${esc(s.version)}… ${s.percent || 0}%`;
      case 'ready': return `Version ${esc(s.version)} is ready. It installs when you close BTS Office, or restart now.`;
      case 'uptodate': return 'You have the latest version.';
      case 'error': return esc(s.error || 'The update check failed.');
      default: return 'BTS Office checks for updates automatically.';
    }
  }

  function renderCard() {
    injectCSS();
    let card = document.getElementById('bts-update-card');
    const show = state && state.status === 'ready' && state.dismissed !== state.version;
    if (!show) { if (card) card.remove(); return; }
    if (card && card.dataset.v === state.version) return;
    if (card) card.remove();
    card = document.createElement('div'); card.id = 'bts-update-card'; card.dataset.v = state.version; card.setAttribute('role', 'status');
    card.innerHTML = `<div class="u-ic">${ic('sparkles', 18)}</div><div style="flex:1;min-width:0">
      <b>Update ready: version ${esc(state.version)}</b>
      <p>Restart to get the latest fixes and features. If you choose Later, it installs the next time you close BTS Office.</p>
      <div class="u-row"><button class="btn primary" data-u="install">${ic('rotate-cw', 15)} Restart now</button><button class="btn" data-u="later">Later</button>${state.notes ? '<button class="btn" data-u="notes" style="margin-left:auto">What’s new</button>' : ''}</div>
      <div class="u-notes hidden">${esc(state.notes)}</div></div>`;
    card.addEventListener('click', async (e) => {
      const b = e.target.closest('[data-u]'); if (!b) return;
      if (b.dataset.u === 'install') { b.disabled = true; const ok = await bts.installUpdate(); if (!ok) b.disabled = false; }
      if (b.dataset.u === 'later') { bts.dismissUpdate(); card.remove(); }
      if (b.dataset.u === 'notes') card.querySelector('.u-notes').classList.toggle('hidden');
    });
    document.body.appendChild(card);
  }

  /* Panel for the About page. Returns HTML; call BTSUpdates.wirePanel(root) after inserting. */
  function panelHTML() {
    if (!desktop) return '<div class="u-panel">Updates are managed by the desktop app.</div>';
    const s = state || {};
    const bar = s.status === 'downloading' ? `<div class="u-bar"><i style="width:${s.percent || 0}%"></i></div>` : '';
    const btn = s.status === 'ready'
      ? `<button class="btn primary" data-u="install">${ic('rotate-cw', 15)} Restart now</button>`
      : `<button class="btn" data-u="check" ${!s.supported || s.status === 'checking' || s.status === 'downloading' ? 'disabled' : ''}>${ic('refresh-cw', 15)} Check for updates</button>`;
    return `<div class="u-panel" id="u-panel"><div class="u-line"><div style="flex:1"><b>Version ${esc(s.current || '')}</b><div class="u-status">${statusText(s)}</div></div>${btn}</div>${bar}
      ${s.status === 'ready' && s.notes ? `<div class="u-notes">${esc(s.notes)}</div>` : ''}
      <label class="check" style="margin-top:12px"><input type="checkbox" id="u-auto" ${s.supported ? '' : 'disabled'}> Download and install updates automatically</label></div>`;
  }
  async function wirePanel(root) {
    const p = (root || document).querySelector('#u-panel'); if (!p) return;
    const settings = await bts.getSettings();
    const cb = p.querySelector('#u-auto'); if (cb) { cb.checked = settings.autoUpdate !== false; cb.onchange = () => bts.setSettings({ autoUpdate: cb.checked }); }
    p.addEventListener('click', async (e) => {
      const b = e.target.closest('[data-u]'); if (!b) return;
      if (b.dataset.u === 'check') { b.disabled = true; await bts.checkForUpdates(); }
      if (b.dataset.u === 'install') { b.disabled = true; const ok = await bts.installUpdate(); if (!ok) b.disabled = false; }
    });
    const refresh = () => { const cur = document.getElementById('u-panel'); if (!cur) { listeners.delete(refresh); return; } const tmp = document.createElement('div'); tmp.innerHTML = panelHTML(); cur.replaceWith(tmp.firstElementChild); wirePanel(); };
    listeners.clear(); listeners.add(refresh);
  }

  async function start() {
    if (!desktop) return;
    bts.onUpdateState((s) => { state = s; renderCard(); listeners.forEach(f => f()); });
    state = await bts.updateState(); renderCard();
  }
  window.BTSUpdates = { panelHTML, wirePanel, get state() { return state; } };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
