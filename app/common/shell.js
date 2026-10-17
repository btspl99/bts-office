/* BTS Office — shared application shell
 *
 * Every editor page (docs/sheets/slides/pdf) uses this file. It builds the title bar,
 * wires the ribbon, provides the File (backstage) view, file open/save plumbing, dialogs,
 * menus, color picker, toasts and keyboard shortcuts.
 *
 * Usage in an app:
 *   BTS.init({
 *     app: 'docs', name: 'BTS Docs', letter: 'D', defaultTitle: 'Document',
 *     saveFormats: [{ ext:'docx', label:'Word document (.docx)' }, ...],   // first = default
 *     openAccept: '.docx,.txt',             // browser fallback only
 *     onCommand(cmd, el, ev) {},            // any [data-cmd] click inside ribbon/statusbar/panes
 *     async load({ name, data:Uint8Array, path }) {},  // open a file into the editor
 *     async save(ext) { return Uint8Array|string },     // serialise to a format
 *     newDocument() {},                     // reset to blank
 *     exportPDF: async () => ({ pageSize, landscape }) // optional: prep before printToPDF; return opts
 *     afterExportPDF: () => {},             // optional cleanup
 *     onZoom(pct) {}, zoom: 100,            // optional status bar zoom
 *     templates: [{ name, thumbHTML, apply() }], // optional extra "New" cards
 *     search(text) {}                        // title bar search box (Enter)
 *   });
 *   BTS.markDirty() / BTS.markClean()  · BTS.title · BTS.filePath
 *   BTS.toast(msg, {error}) · BTS.dialog({title, body, buttons:[{label, value, primary}], onOpen}) -> Promise(value, root)
 *   BTS.prompt(title, label, value) -> Promise(string|null)
 *   BTS.menu(anchorEl|{x,y}, items:[{label, icon, kbd, action, disabled, sep, head}])
 *   BTS.colorPicker(anchorEl, { auto:'Automatic', none:'No color' }) -> Promise(color|'auto'|'none'|null)
 *   BTS.gridPicker(anchorEl) -> Promise({rows, cols}|null)
 *   BTS.setStatus(key, html)  (status bar slot created on demand, left side)
 *   BTS.pickFile(accept) -> Promise({name, data}|null)  (images etc.)
 *   BTS.showTab(name) · BTS.esc(str)
 *   BTS.prefs (File > Options; see PREF_DEFAULTS) · BTS.setPref(key, value) · BTS.onPrefs(cb(prefs, key)) · BTS.defaultExt()
 *   BTS.openBackstage(view)  views: home, new, open, info, saveas, export, print, about, options
 */
(function () {
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const desktop = !!(window.bts && window.bts.isDesktop);

  const APPS = {
    docs: { name: 'BTS Docs', letter: 'D', color: '#3C4FE0', ext: 'docx' },
    sheets: { name: 'BTS Sheets', letter: 'S', color: '#0F8A6A', ext: 'xlsx' },
    slides: { name: 'BTS Slides', letter: 'P', color: '#D9661A', ext: 'pptx' },
    pdf: { name: 'BTS PDF', letter: 'R', color: '#C7354A', ext: 'pdf' },
  };
  const FILE_COLOR = { docx: '#3C4FE0', doc: '#3C4FE0', txt: '#47526a', md: '#47526a', html: '#47526a', odt: '#3C4FE0', rtf: '#3C4FE0', xlsx: '#0F8A6A', xls: '#0F8A6A', csv: '#0F8A6A', ods: '#0F8A6A', pptx: '#D9661A', odp: '#D9661A', pdf: '#C7354A' };

  const BTS = window.BTS = {
    cfg: null, title: 'Untitled', filePath: null, dirty: false, esc, $, $$, desktop, APPS, FILE_COLOR,
    /* Document properties. Apps read these when saving (docx/xlsx/pptx core properties) and set them when opening a file. */
    props: { title: '', subject: '', author: '', company: '', keywords: '', comments: '', category: '' },
    /* The person using BTS Office (File > Options). Used as default author, and for comments and tracked changes. */
    userName: (() => { try { return localStorage.getItem('bts-user-name') || ''; } catch { return ''; } })(),
  };
  /* ---------- options (File > Options) ----------
   * BTS.prefs holds the person's options. It is read synchronously from localStorage ('bts-prefs') when this file
   * loads, so apps can use it while they build their first document; BTS.init then merges settings.json (desktop).
   * BTS.setPref(key, value) saves to both; BTS.onPrefs(cb) → cb(prefs, changedKey) after every change. */
  const PREF_DEFAULTS = { startScreen: false, zoom: 100, spellCheck: true, spellLanguages: null, autosaveSeconds: 0,
    defaultFormat_docs: 'docx', defaultFormat_sheets: 'xlsx', defaultFormat_slides: 'pptx',
    docsFont: '', docsSize: 0, sheetsFont: '', sheetsSize: 0, pageSize: 'A4', units: 'cm' };
  BTS.PREF_DEFAULTS = PREF_DEFAULTS;
  BTS.prefs = Object.assign({}, PREF_DEFAULTS, (() => { try { return JSON.parse(localStorage.getItem('bts-prefs') || '{}') || {}; } catch { return {}; } })());
  const prefListeners = [];
  const mirrorPrefs = () => { try { localStorage.setItem('bts-prefs', JSON.stringify(BTS.prefs)); } catch {} };
  BTS.onPrefs = (cb) => { prefListeners.push(cb); return () => { const i = prefListeners.indexOf(cb); if (i >= 0) prefListeners.splice(i, 1); }; };
  const firePrefs = (k) => prefListeners.forEach(cb => { try { cb(BTS.prefs, k); } catch (e) { console.error(e); } });
  BTS.setPref = (k, v) => {
    BTS.prefs[k] = v; mirrorPrefs();
    const done = desktop ? bts.setSettings({ [k]: v }) : Promise.resolve(true);
    firePrefs(k); return done;
  };
  // other windows: pick up changes made in another window's Options page
  window.addEventListener('storage', (e) => { if (e.key !== 'bts-prefs' || !e.newValue) return; try { Object.assign(BTS.prefs, JSON.parse(e.newValue)); firePrefs(null); } catch {} });
  BTS.setProps = (p) => { Object.assign(BTS.props, p || {}); };
  BTS.resetProps = () => { BTS.props = { title: '', subject: '', author: BTS.userName || '', company: '', keywords: '', comments: '', category: '' }; };

  /* ---------- theme ---------- */
  function applyTheme(t) {
    const dark = t === 'dark' || (t === 'system' && matchMedia('(prefers-color-scheme: dark)').matches);
    document.documentElement.dataset.theme = dark ? 'dark' : 'light';
    BTS.theme = t;
  }
  BTS.setTheme = (t) => { applyTheme(t); try { localStorage.setItem('bts-theme', t); } catch {} if (desktop) bts.setSettings({ theme: t }); };
  applyTheme((() => { try { return localStorage.getItem('bts-theme') || 'light'; } catch { return 'light'; } })());

  /* ---------- toasts ---------- */
  BTS.toast = (msg, opts = {}) => {
    let host = $('#toasts'); if (!host) { host = document.createElement('div'); host.id = 'toasts'; document.body.appendChild(host); }
    const t = document.createElement('div'); t.className = 'toast' + (opts.error ? ' err' : '');
    t.innerHTML = (opts.error ? icon('circle-alert', 16) : icon('check', 16)) + '<span>' + esc(msg) + '</span>';
    host.appendChild(t); setTimeout(() => t.remove(), opts.ms || 2600);
  };

  /* ---------- popovers ---------- */
  let openPop = null;
  function closePop() { if (openPop) { const p = openPop; openPop = null; p.el.remove(); p.resolve && p.resolve(null); document.removeEventListener('mousedown', onDocDown, true); document.removeEventListener('keydown', onPopKey, true); } }
  function onDocDown(e) { if (openPop && !openPop.el.contains(e.target)) { if (openPop.anchor && openPop.anchor.contains && openPop.anchor.contains(e.target)) { e.stopPropagation(); e.preventDefault(); } closePop(); } }
  function onPopKey(e) {
    if (!openPop) return;
    if (e.key === 'Escape') { e.preventDefault(); closePop(); return; }
    const items = $$('.menu-item:not(.disabled)', openPop.el); if (!items.length) return;
    const cur = items.findIndex(i => i.classList.contains('focus'));
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault(); items.forEach(i => i.classList.remove('focus'));
      const n = e.key === 'ArrowDown' ? (cur + 1) % items.length : (cur - 1 + items.length) % items.length; items[n].classList.add('focus');
    } else if (e.key === 'Enter' && cur >= 0) { e.preventDefault(); items[cur].click(); }
  }
  BTS.closePopover = closePop;
  BTS.popover = (anchor, el, opts = {}) => {
    closePop();
    el.classList.add('pop'); document.body.appendChild(el);
    let x, y;
    if (anchor && anchor.getBoundingClientRect) { const r = anchor.getBoundingClientRect(); x = r.left; y = r.bottom + 3; }
    else { x = anchor.x; y = anchor.y; }
    const w = el.offsetWidth, h = el.offsetHeight;
    if (x + w > innerWidth - 6) x = Math.max(6, innerWidth - w - 6);
    if (y + h > innerHeight - 6) y = Math.max(6, (anchor && anchor.getBoundingClientRect ? anchor.getBoundingClientRect().top - h - 3 : innerHeight - h - 6));
    el.style.left = x + 'px'; el.style.top = y + 'px';
    return new Promise(resolve => {
      openPop = { el, resolve, anchor };
      setTimeout(() => { document.addEventListener('mousedown', onDocDown, true); document.addEventListener('keydown', onPopKey, true); }, 0);
    });
  };
  BTS.resolvePopover = (v) => { if (openPop) { const p = openPop; openPop = null; p.resolve && p.resolve(v); p.el.remove(); document.removeEventListener('mousedown', onDocDown, true); document.removeEventListener('keydown', onPopKey, true); } };

  BTS.menu = (anchor, items) => {
    const el = document.createElement('div');
    items.forEach(it => {
      if (it.sep) { el.insertAdjacentHTML('beforeend', '<div class="menu-sep"></div>'); return; }
      if (it.head) { el.insertAdjacentHTML('beforeend', '<div class="menu-head">' + esc(it.head) + '</div>'); return; }
      const d = document.createElement('div'); d.className = 'menu-item' + (it.disabled ? ' disabled' : '');
      d.innerHTML = (it.icon ? icon(it.icon, 16) : (it.checked !== undefined ? (it.checked ? icon('check', 16) : '<span style="width:16px"></span>') : '')) + (it.html || '<span>' + esc(it.label) + '</span>') + (it.kbd ? '<span class="kbd">' + esc(it.kbd) + '</span>' : '');
      d.addEventListener('mousedown', e => e.preventDefault());
      d.addEventListener('click', () => { BTS.resolvePopover(it.value ?? it.label); it.action && it.action(); });
      el.appendChild(d);
    });
    return BTS.popover(anchor, el);
  };

  const PALETTE = [
    ['#000000', '#262626', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#d9d9d9', '#efefef', '#ffffff'],
    ['#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff', '#9900ff', '#ff00ff'],
    ['#e6b8af', '#f4cccc', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#cfe2f3', '#d9d2e9', '#ead1dc'],
    ['#dd7e6b', '#ea9999', '#f9cb9c', '#ffe599', '#b6d7a8', '#a2c4c9', '#a4c2f4', '#9fc5e8', '#b4a7d6', '#d5a6bd'],
    ['#cc4125', '#e06666', '#f6b26b', '#ffd966', '#93c47d', '#76a5af', '#6d9eeb', '#6fa8dc', '#8e7cc3', '#c27ba0'],
    ['#a61c00', '#cc0000', '#e69138', '#f1c232', '#6aa84f', '#45818e', '#3c78d8', '#3d85c6', '#674ea7', '#a64d79'],
    ['#5b0f00', '#660000', '#783f04', '#7f6000', '#274e13', '#0c343d', '#1c4587', '#073763', '#20124d', '#4c1130'],
  ];
  BTS.PALETTE = PALETTE;
  BTS.colorPicker = (anchor, opts = {}) => {
    const el = document.createElement('div'); el.className = 'color-pop';
    el.innerHTML = '<div class="color-grid">' + PALETTE.flat().map(c => `<button data-c="${c}" title="${c}" style="background:${c}"></button>`).join('') + '</div>' +
      '<div class="row">' + (opts.auto ? `<button data-c="auto">${esc(opts.auto)}</button>` : '') + (opts.none ? `<button data-c="none">${esc(opts.none)}</button>` : '') +
      `<input type="color" title="Custom color" value="${opts.value && opts.value[0] === '#' ? opts.value : '#3c4fe0'}"></div>`;
    el.addEventListener('mousedown', e => { if (e.target.tagName !== 'INPUT') e.preventDefault(); });
    el.addEventListener('click', e => { const b = e.target.closest('button[data-c]'); if (b) BTS.resolvePopover(b.dataset.c); });
    $('input[type=color]', el).addEventListener('change', e => BTS.resolvePopover(e.target.value));
    return BTS.popover(anchor, el);
  };


  /* ---------- font picker (shared by Docs, Sheets, Slides) ----------
   * BTS.fontPicker(anchor, { current, theme: [{label, name}] }) -> Promise(fontName|null)
   * Searchable list grouped by category, each name previewed in its own font, with recent fonts on top. */
  BTS.fontPicker = (anchor, opts = {}) => {
    const fonts = (window.BTS_FONTS || []).slice();
    const cats = window.BTS_FONT_CATEGORIES || {};
    const el = document.createElement('div'); el.className = 'font-pop';
    el.innerHTML = `<div class="fp-search">${icon('search', 15)}<input placeholder="Search fonts" spellcheck="false"></div><div class="fp-list" role="listbox"></div><div class="fp-foot"><span>${fonts.filter(f => f.bundled).length} fonts included</span><span>Free to share and embed</span></div>`;
    const list = el.querySelector('.fp-list'); const inp = el.querySelector('input');
    const stack = (n) => window.btsFontStack ? btsFontStack(n) : `'${n}'`;
    const row = (f, label) => { const avail = f.bundled || (window.btsFontAvailable ? btsFontAvailable(f.name) : true); return `<div class="menu-item fp-item${f.name === opts.current ? ' on' : ''}${avail ? '' : ' fp-missing'}" data-font="${esc(f.name)}" role="option"><span class="fp-name" style="font-family:${stack(f.name)}">${esc(label || f.name)}</span>${f.cat === 'indian' ? `<span class="kbd" style="font-family:${stack(f.name)}">हिन्दी मराठी</span>` : ''}${f.alias ? `<span class="kbd">for ${esc(f.alias)}</span>` : ''}${!avail ? '<span class="kbd">not installed</span>' : ''}</div>`; };
    const render = (q) => {
      q = (q || '').trim().toLowerCase(); let h = '';
      if (!q) {
        if (opts.theme && opts.theme.length) h += '<div class="menu-head">Theme fonts</div>' + opts.theme.map(t => row(fonts.find(f => f.name === t.name) || { name: t.name, cat: 'sans', bundled: true }, t.name).replace('</div>', `<span class="kbd">${esc(t.label)}</span></div>`)).join('');
        const rec = (window.btsRecentFonts ? btsRecentFonts() : []).map(n => fonts.find(f => f.name === n)).filter(Boolean);
        if (rec.length) h += '<div class="menu-head">Recently used</div>' + rec.map(f => row(f)).join('');
      }
      for (const [c, label] of Object.entries(cats)) {
        const items = fonts.filter(f => f.cat === c && (!q || f.name.toLowerCase().includes(q) || (f.alias || '').toLowerCase().includes(q))).sort((a, b) => a.name.localeCompare(b.name));
        if (items.length) h += `<div class="menu-head">${esc(label)}</div>` + items.map(f => row(f)).join('');
      }
      list.innerHTML = h || '<div class="empty-note" style="padding:12px">No fonts match.</div>';
      const on = list.querySelector('.on'); if (on && !q) on.scrollIntoView({ block: 'center' });
    };
    render('');
    inp.addEventListener('input', () => render(inp.value));
    inp.addEventListener('keydown', e => { if (e.key === 'Enter') { const first = list.querySelector('.menu-item.focus') || list.querySelector('.fp-item'); if (first) { e.preventDefault(); first.click(); } } });
    el.addEventListener('mousedown', e => { if (e.target !== inp) e.preventDefault(); });
    list.addEventListener('click', e => { const it = e.target.closest('[data-font]'); if (it) { if (window.btsRecentFonts) btsRecentFonts(it.dataset.font); BTS.resolvePopover(it.dataset.font); } });
    const pr = BTS.popover(anchor, el);
    setTimeout(() => inp.focus(), 0);
    return pr;
  };

  /* ---------- command search ("Tell me what you want to do") ----------
   * Indexes every ribbon command (button label, tooltip and tab) so people can find any feature by name.
   * Opens with Alt+Q or Ctrl+Shift+P, or the lightbulb button in the title bar. */
  BTS.commandSearch = () => {
    const tabs = Object.fromEntries($$('#ribbon-tabs [data-tab]').map(b => [b.dataset.tab, b.textContent.trim()]));
    const seen = new Set(); const cmds = [];
    $$('#ribbon .ribbon-panel').forEach(panel => {
      $$('[data-cmd]', panel).forEach(b => {
        if (b.tagName === 'SELECT' || b.tagName === 'INPUT' || b.disabled) return;
        const label = (b.getAttribute('title') || b.getAttribute('aria-label') || b.textContent || '').replace(/\s+/g, ' ').trim();
        if (!label) return;
        const group = b.closest('.rgroup')?.querySelector('.rgroup-label')?.textContent.trim() || '';
        const key = b.dataset.cmd + '|' + label; if (seen.has(key)) return; seen.add(key);
        cmds.push({ label, group, tab: tabs[panel.dataset.panel] || '', panel: panel.dataset.panel, el: b, icon: b.dataset.icon });
      });
    });
    $$('#titlebar [data-search]').forEach(b => cmds.push({ label: b.dataset.search, group: '', tab: 'Title bar', el: b, icon: b.dataset.searchIcon }));
    if (window.BTSLegacyFonts) cmds.push({ label: 'Convert old Hindi/Marathi fonts (Kruti Dev, Shivaji) to Unicode', group: 'Indian languages', tab: '', run: () => BTSLegacyFonts.dialog(), icon: 'languages' });
    const back = document.createElement('div'); back.className = 'modal-back cmd-back';
    back.innerHTML = `<div class="cmd-pal" role="dialog" aria-label="Search commands"><div class="cmd-in">${icon('lightbulb', 18)}<input placeholder="Tell me what you want to do, for example “insert table” or “page numbers”" spellcheck="false"></div><div class="cmd-list"></div><div class="cmd-foot"><span>↑ ↓ to choose · Enter to run · Esc to close</span><span>${cmds.length} commands</span></div></div>`;
    document.body.appendChild(back);
    const inp = back.querySelector('input'); const list = back.querySelector('.cmd-list'); let shown = []; let sel = 0;
    const score = (c, q) => { const l = (c.label + ' ' + c.group + ' ' + c.tab).toLowerCase(); if (!q) return 1; let s = 0; for (const w of q.split(/\s+/)) { if (!w) continue; const i = l.indexOf(w); if (i < 0) return 0; s += i === 0 ? 3 : (l[i - 1] === ' ' ? 2 : 1); } return s; };
    const render = () => { const q = inp.value.trim().toLowerCase(); shown = cmds.map(c => [score(c, q), c]).filter(x => x[0] > 0).sort((a, b) => b[0] - a[0]).slice(0, 40).map(x => x[1]); sel = Math.min(sel, Math.max(0, shown.length - 1));
      list.innerHTML = shown.length ? shown.map((c, i) => `<div class="cmd-item${i === sel ? ' sel' : ''}" data-i="${i}">${icon(c.icon || 'square-dot', 16)}<span>${esc(c.label)}</span><small>${esc([c.tab, c.group].filter(Boolean).join(' › '))}</small></div>`).join('') : '<div class="empty-note" style="padding:14px 16px">No matching command. Try another word.</div>';
      const s = list.querySelector('.sel'); if (s) s.scrollIntoView({ block: 'nearest' }); };
    const close = () => { back.remove(); document.removeEventListener('keydown', key, true); };
    const run = (c) => { close(); if (!c) return; if (c.panel) BTS.showTab(c.panel); setTimeout(() => (c.run ? c.run() : c.el.click()), 30); };
    const key = (e) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); close(); } else if (e.key === 'ArrowDown') { e.preventDefault(); sel = Math.min(sel + 1, shown.length - 1); render(); } else if (e.key === 'ArrowUp') { e.preventDefault(); sel = Math.max(sel - 1, 0); render(); } else if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); run(shown[sel]); } };
    document.addEventListener('keydown', key, true);
    inp.addEventListener('input', () => { sel = 0; render(); });
    list.addEventListener('mousemove', e => { const it = e.target.closest('[data-i]'); if (it && +it.dataset.i !== sel) { sel = +it.dataset.i; list.querySelectorAll('.cmd-item').forEach((x, i) => x.classList.toggle('sel', i === sel)); } });
    list.addEventListener('click', e => { const it = e.target.closest('[data-i]'); if (it) run(shown[+it.dataset.i]); });
    back.addEventListener('mousedown', e => { if (e.target === back) close(); });
    render(); inp.focus();
  };

  BTS.gridPicker = (anchor, maxR = 8, maxC = 10) => {
    const el = document.createElement('div');
    el.innerHTML = '<div class="grid-picker-label">Insert table</div><div class="grid-picker"></div>';
    const g = $('.grid-picker', el); const lab = $('.grid-picker-label', el);
    for (let r = 0; r < maxR; r++) for (let c = 0; c < maxC; c++) { const d = document.createElement('div'); d.dataset.r = r; d.dataset.c = c; g.appendChild(d); }
    g.addEventListener('mouseover', e => { const d = e.target.closest('[data-r]'); if (!d) return; const R = +d.dataset.r, C = +d.dataset.c; $$('div', g).forEach(x => x.classList.toggle('on', +x.dataset.r <= R && +x.dataset.c <= C)); lab.textContent = `${C + 1} × ${R + 1} table`; });
    g.addEventListener('mousedown', e => e.preventDefault());
    g.addEventListener('click', e => { const d = e.target.closest('[data-r]'); if (d) BTS.resolvePopover({ rows: +d.dataset.r + 1, cols: +d.dataset.c + 1 }); });
    return BTS.popover(anchor, el);
  };

  /* ---------- modal dialogs ---------- */
  BTS.dialog = ({ title, body, buttons, onOpen, width }) => new Promise(resolve => {
    const back = document.createElement('div'); back.className = 'modal-back';
    const btns = buttons || [{ label: 'OK', value: true, primary: true }];
    back.innerHTML = `<div class="modal" role="dialog" aria-modal="true" style="${width ? 'width:' + width + 'px' : ''}"><header>${esc(title)}<button class="rb x" data-x title="Close">${icon('x', 16)}</button></header><div class="body"></div><footer>${btns.map((b, i) => `<button class="btn ${b.primary ? 'primary' : ''} ${b.danger ? 'danger' : ''}" data-i="${i}">${esc(b.label)}</button>`).join('')}</footer></div>`;
    const bodyEl = $('.body', back); if (typeof body === 'string') bodyEl.innerHTML = body; else if (body) bodyEl.appendChild(body);
    const done = (v) => { back.remove(); document.removeEventListener('keydown', key, true); resolve(v); };
    const key = (e) => {
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); done(null); }
      if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA' && !e.target.closest('.modal-noenter')) { const p = btns.findIndex(b => b.primary); if (p >= 0) { e.preventDefault(); e.stopPropagation(); finish(p); } }
    };
    const finish = (i) => { const b = btns[i]; if (b.validate && !b.validate(back)) return; done(typeof b.value === 'function' ? b.value(back) : b.value); };
    back.addEventListener('click', e => { if (e.target.closest('[data-x]')) done(null); const b = e.target.closest('footer [data-i]'); if (b) finish(+b.dataset.i); });
    document.addEventListener('keydown', key, true);
    document.body.appendChild(back);
    onOpen && onOpen(back);
    const f = $('input,select,textarea', bodyEl); if (f) { f.focus(); f.select && f.select(); } else { const p = $('footer .primary', back); p && p.focus(); }
  });
  BTS.prompt = (title, label, value = '') => BTS.dialog({
    title, body: `<div class="field"><label>${esc(label)}</label><input id="bts-prompt" value="${esc(value)}"></div>`,
    buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => $('#bts-prompt', r).value }],
  });
  BTS.confirm = (title, text, ok = 'OK') => BTS.dialog({ title, body: `<p style="margin:4px 0 8px;max-width:440px;line-height:1.5">${esc(text)}</p>`, buttons: [{ label: 'Cancel', value: false }, { label: ok, primary: true, value: true }] });

  /* ---------- file helpers ---------- */
  BTS.pickFile = (accept) => new Promise(resolve => {
    const inp = document.createElement('input'); inp.type = 'file'; inp.accept = accept || '';
    inp.onchange = async () => { const f = inp.files[0]; if (!f) return resolve(null); resolve({ name: f.name, type: f.type, data: new Uint8Array(await f.arrayBuffer()) }); };
    inp.click();
  });
  BTS.bytesToDataURL = (bytes, mime) => new Promise(res => { const r = new FileReader(); r.onload = () => res(r.result); r.readAsDataURL(new Blob([bytes], { type: mime })); });
  BTS.dataURLToBytes = (url) => { const b = atob(url.split(',')[1]); const a = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) a[i] = b.charCodeAt(i); return a; };
  BTS.download = (name, data, mime) => { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([data], { type: mime || 'application/octet-stream' })); a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 4000); };
  const extOf = (n) => (n.match(/\.([^.]+)$/) || [, ''])[1].toLowerCase();
  BTS.extOf = extOf;
  const baseOf = (n) => n.replace(/\.[^.]+$/, '');

  /* ---------- dirty / title ---------- */
  function updateTitle() {
    const inp = $('#tb-docname'); if (inp && document.activeElement !== inp) inp.value = BTS.title;
    const st = $('#tb-state'); if (st) st.textContent = BTS.dirty ? '• Edited' : (BTS.filePath ? 'Saved' : '');
    const t = `${BTS.dirty ? '● ' : ''}${BTS.title} — ${BTS.cfg.name}`; document.title = t; if (desktop) bts.setTitle(t);
  }
  BTS.markDirty = () => { if (!BTS.dirty) { BTS.dirty = true; if (desktop) bts.setDirty(true); updateTitle(); } };
  BTS.markClean = () => { BTS.dirty = false; if (desktop) bts.setDirty(false); updateTitle(); };
  BTS.setTitle = (t) => { BTS.title = t || BTS.cfg.defaultTitle; updateTitle(); };

  async function loadFile(f) {
    if (!f) return;
    if (f.error) { BTS.toast(f.error, { error: true }); return; }
    if (f.handledElsewhere) return;
    try {
      document.body.style.cursor = 'progress';
      BTS.resetProps();
      await BTS.cfg.load(f);
      const ext = extOf(f.name);
      BTS.title = baseOf(f.name);
      BTS.filePath = (BTS.cfg.saveFormats.some(s => s.ext === ext) && nativeExt(ext)) ? f.path || null : null;
      BTS.loadedExt = ext;
      BTS.markClean();
    } catch (err) { console.error(err); BTS.toast('Couldn’t open ' + f.name + ': ' + (err.message || err), { error: true, ms: 5000 }); }
    finally { document.body.style.cursor = ''; }
  }
  BTS.loadFile = loadFile;

  BTS.open = async () => {
    closeBackstage();
    if (desktop) { const f = await bts.openDialog(BTS.cfg.app); if (!f) return; if (BTS.dirty || BTS.hasContent?.()) { if (!f.handledElsewhere) { bts.openApp(null, f.path); } return; } loadFile(f); }
    else { const f = await BTS.pickFile(BTS.cfg.openAccept); loadFile(f); }
  };

  async function writeOut(ext, path) {
    const data = await BTS.cfg.save(ext);
    if (desktop) {
      const r = await bts.writeFile(path, data);
      if (r.error) { BTS.toast('Couldn’t save: ' + r.error, { error: true, ms: 5000 }); return false; }
      return r;
    } else { BTS.download(BTS.title + '.' + ext, data); return { name: BTS.title + '.' + ext }; }
  }
  /* The format Save uses for a new file: File > Options > Save ("defaultFormat_<app>"), else the app's first format.
   * Files in the app's main format or in the chosen default format are saved in place; others go through Save As. */
  BTS.defaultExt = () => { const fm = BTS.cfg.saveFormats; const p = BTS.prefs['defaultFormat_' + BTS.cfg.app]; return (p && fm.some(f => f.ext === p)) ? p : fm[0].ext; };
  const nativeExt = (e) => e === BTS.cfg.saveFormats[0].ext || e === BTS.defaultExt();
  BTS.save = async () => {
    if (!BTS.filePath) return BTS.saveAs();
    const fe = extOf(BTS.filePath);
    const r = await writeOut(nativeExt(fe) ? fe : BTS.cfg.saveFormats[0].ext, BTS.filePath);
    if (r) { BTS.markClean(); BTS.toast('Saved'); }
    return !!r;
  };
  BTS.saveAs = async (ext) => {
    closeBackstage();
    ext = ext || BTS.defaultExt();
    const fmt = BTS.cfg.saveFormats.find(f => f.ext === ext) || BTS.cfg.saveFormats[0];
    if (desktop) {
      const p = await bts.saveDialog(BTS.title + '.' + fmt.ext, [{ name: fmt.label, extensions: [fmt.ext] }]);
      if (!p) return false;
      const r = await writeOut(fmt.ext, p); if (!r) return false;
      if (nativeExt(fmt.ext)) { BTS.filePath = p; BTS.title = baseOf(r.name); BTS.markClean(); }
      BTS.toast('Saved as ' + r.name);
      return true;
    } else { await writeOut(fmt.ext); BTS.markClean(); return true; }
  };
  BTS.exportPDF = async () => {
    closeBackstage();
    const opts = (BTS.cfg.exportPDF && await BTS.cfg.exportPDF()) || {};
    try {
      if (desktop) {
        const p = await bts.saveDialog(BTS.title + '.pdf', [{ name: 'PDF document', extensions: ['pdf'] }]);
        if (!p) return;
        await new Promise(r => setTimeout(r, 60));
        const data = await bts.printToPDF(opts);
        const r = await bts.writeFile(p, data);
        if (r.error) BTS.toast(r.error, { error: true }); else BTS.toast('Exported ' + r.name);
      } else window.print();
    } finally { BTS.cfg.afterExportPDF && BTS.cfg.afterExportPDF(); }
  };
  BTS.print = async () => {
    closeBackstage();
    const opts = (BTS.cfg.exportPDF && await BTS.cfg.exportPDF()) || {};
    try { if (desktop) await bts.print(); else window.print(); }
    finally { BTS.cfg.afterExportPDF && BTS.cfg.afterExportPDF(); }
  };
  BTS.newWindow = (app) => { if (desktop) bts.openApp(app || BTS.cfg.app); else window.open((app || BTS.cfg.app) + '.html'); };

  /* ---------- backstage ---------- */
  function closeBackstage() { const b = $('#backstage'); if (b) b.classList.remove('open'); }
  BTS.closeBackstage = closeBackstage;
  async function renderBackstage(view) {
    const b = $('#backstage'); const main = $('.bs-main', b);
    $$('nav button[data-view]', b).forEach(x => x.classList.toggle('active', x.dataset.view === view));
    const cfg = BTS.cfg;
    if (view === 'home' || view === 'new') {
      const other = Object.entries(APPS).filter(([k]) => k !== cfg.app && k !== 'pdf');
      main.innerHTML = `<h1>${view === 'home' ? greet() : 'New'}</h1>
        <div class="bs-cards">
          <button class="bs-card" data-act="blank"><div class="thumb">${cfg.blankThumb || icon('plus', 34)}</div><div class="lbl">Blank ${esc(cfg.defaultTitle.toLowerCase())}</div></button>
          ${(cfg.templates || []).map((t, i) => `<button class="bs-card" data-tpl="${i}"><div class="thumb">${t.thumbHTML || ''}</div><div class="lbl">${esc(t.name)}</div></button>`).join('')}
        </div>
        <h2>Start something else</h2>
        <div class="bs-cards">${other.map(([k, a]) => `<button class="bs-card wide" data-app="${k}"><div class="thumb" style="color:${a.color};gap:10px;font-weight:600">${appBadge(k, 30)} ${a.name}</div></button>`).join('')}</div>
        ${view === 'home' ? '<h2>Recent</h2><div class="bs-list" id="bs-recent"></div>' : ''}`;
      if (view === 'home') renderRecent($('#bs-recent', main), 8);
    } else if (view === 'open') {
      main.innerHTML = `<h1>Open</h1><div class="bs-actions"><button class="bs-action" data-act="browse">${icon('folder-open')}<div><b>Browse this PC</b><small>Open a file stored on your computer</small></div></button></div><h2>Recent</h2><div class="bs-list" id="bs-recent"></div>`;
      renderRecent($('#bs-recent', main), 30);
    } else if (view === 'saveas' || view === 'export') {
      const fm = view === 'export' ? [{ ext: 'pdf', label: 'PDF document (.pdf)', pdf: true }, ...cfg.saveFormats.slice(1)] : cfg.saveFormats;
      main.innerHTML = `<h1>${view === 'export' ? 'Export' : 'Save a copy'}</h1><div class="bs-actions">${fm.map(f => `<button class="bs-action" data-fmt="${f.ext}" ${f.pdf ? 'data-pdf' : ''}>${icon(f.pdf ? 'file-down' : (f.icon || 'file'))}<div><b>${esc(f.label)}</b><small>${esc(f.hint || (f.pdf ? 'Fixed layout, good for sharing and printing' : 'Save in this format'))}</small></div></button>`).join('')}</div>`;
    } else if (view === 'print') {
      main.innerHTML = `<h1>Print</h1><div class="bs-actions"><button class="bs-action" data-act="print">${icon('printer')}<div><b>Print</b><small>Choose a printer and settings</small></div></button><button class="bs-action" data-act="pdf">${icon('file-down')}<div><b>Save as PDF</b><small>Create a PDF you can share</small></div></button></div>`;
    } else if (view === 'about') {
      const info = desktop ? await bts.info() : { version: 'web' };
      main.innerHTML = `<h1>About</h1>
        <div class="about-head">${appBadge(cfg.app, 54)}<div><div class="about-name">${esc(cfg.name)}</div><div class="about-ver">BTS Office · Version ${esc(info.version)}</div></div></div>
        <p class="about-p">BTS Office is a free, open-source office suite for students — documents, spreadsheets, presentations and PDFs, with support for the file formats your school and college already use (.docx, .xlsx, .pptx, .pdf).</p>
        <div class="about-box"><span>Developer</span><b>Bhoomiputra Tech Solutions Private Limited</b><span>Website</span><a href="#" data-ext="https://btspl.tech">btspl.tech</a></div>
        <p class="about-p">Released under the GNU General Public License v3. You can use, share and improve it freely. Built with Electron, HyperFormula, ExcelJS, SheetJS, docx, Mammoth, PptxGenJS, PDF.js, Chart.js and Lucide icons.</p>
        <h2>Updates</h2>
        ${window.BTSUpdates ? BTSUpdates.panelHTML() : ''}
        <h2>Keyboard shortcuts</h2>
        <div class="about-keys">${[['Ctrl+N', 'New window'], ['Ctrl+O', 'Open'], ['Ctrl+S', 'Save'], ['Ctrl+Shift+S', 'Save a copy'], ['Ctrl+P', 'Print'], ['Ctrl+Z / Ctrl+Y', 'Undo / Redo'], ['Ctrl+F', 'Find'], ['Ctrl+H', 'Replace'], ['Ctrl+F1', 'Show or hide the ribbon'], ['Alt+Q', 'Search commands'], ...(window.BTSVoice ? [['Alt+Shift+D', 'Voice typing']] : []), ...(window.BTSIndic ? [['Ctrl+Alt+H / Ctrl+Alt+M', 'Type in Hindi / Marathi']] : []), ...(cfg.shortcuts || [])].map(([k, v]) => `<b>${k}</b><span>${v}</span>`).join('')}</div>`;
      if (window.BTSUpdates) BTSUpdates.wirePanel(main);
    } else if (view === 'options') {
      await renderOptions(main);
    } else if (view === 'info') {
      const P = BTS.props; const fld = (k, label, ph) => `<div class="field"><label>${label}</label><input data-prop="${k}" value="${esc(P[k] || '')}" placeholder="${esc(ph || '')}"></div>`;
      main.innerHTML = `<h1>Info</h1>${cfg.infoHTML ? cfg.infoHTML() : ''}
        <h2>Properties</h2><div class="props-grid">${fld('title', 'Title', BTS.title)}${fld('subject', 'Subject')}${fld('author', 'Author', BTS.userName || 'Your name')}${fld('company', 'Company or school')}${fld('keywords', 'Tags', 'Separate with commas')}${fld('category', 'Category')}<div class="field" style="grid-column:1/-1"><label>Comments</label><textarea data-prop="comments">${esc(P.comments || '')}</textarea></div></div>
        <p class="empty-note" style="padding:0">Properties are saved inside the file and show up in Windows Explorer and other office apps.</p>
        <h2>Version history</h2><div class="bs-list" id="bs-versions"><div class="empty-note">Loading…</div></div>`;
      main.querySelectorAll('[data-prop]').forEach(i => i.addEventListener('change', () => { BTS.props[i.dataset.prop] = i.value; BTS.markDirty(); }));
      const vl = main.querySelector('#bs-versions');
      if (!desktop || !BTS.filePath) vl.innerHTML = '<div class="empty-note">Save the file to start keeping versions. Every time you save, BTS Office keeps a copy (the last 30) that you can go back to.</div>';
      else { const vs = await bts.versions(BTS.filePath); vl.innerHTML = vs.length ? vs.map((v, i) => `<button class="bs-row" data-version="${esc(v.id)}"><div class="fi" style="background:var(--accent)">${icon('history', 16)}</div><div><div>${new Date(v.time).toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' })}${i === 0 ? ' (current)' : ''}</div><div class="meta">${Math.max(1, Math.round(v.size / 1024))} KB · opens as a copy in a new window</div></div></button>`).join('') : '<div class="empty-note">No saved versions yet. A version is kept each time you save.</div>'; }
    }
  }
  /* ---------- File > Options ---------- */
  const OPT_SECTIONS = [['general', 'General', 'sliders-horizontal'], ['proofing', 'Proofing', 'spell-check'], ['save', 'Save', 'save'], ['new', 'New documents', 'file-plus'], ['language', 'Language', 'languages']];
  async function renderOptions(main) {
    const P = BTS.prefs; const st = desktop ? await bts.getSettings().catch(() => ({})) : {};
    const sel = (key, opts, val) => `<select class="opt-sel" data-pref="${key}">${opts.map(([v, l]) => `<option value="${esc(v)}"${String(v) === String(val) ? ' selected' : ''}>${esc(l)}</option>`).join('')}</select>`;
    const sw = (id, on, attrs = '') => `<label class="opt-sw"><input type="checkbox" id="${id}"${on ? ' checked' : ''} ${attrs}><span></span></label>`;
    const row = (label, hint, ctl, extra = '') => `<div class="opt-row" ${extra}><div class="opt-l"><b>${label}</b>${hint ? `<small>${hint}</small>` : ''}</div><div class="opt-c">${ctl}</div></div>`;
    const fmtOpts = (app) => (({ docs: [['docx', 'Word document (.docx)'], ['odt', 'OpenDocument Text (.odt)']], sheets: [['xlsx', 'Excel workbook (.xlsx)'], ['ods', 'OpenDocument spreadsheet (.ods)']], slides: [['pptx', 'PowerPoint (.pptx)'], ['odp', 'OpenDocument presentation (.odp)']] })[app]);
    const sizes = [['0', 'Default'], ...[8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20].map(s => [String(s), String(s)])];
    const fontBtn = (key) => `<button class="btn opt-font" data-font-pref="${key}" style="font-family:${P[key] ? (window.btsFontStack ? btsFontStack(P[key]) : `'${P[key]}'`).replace(/"/g, "'") : 'inherit'}">${esc(P[key] || 'Default (Calibri)')}${icon('chevron-down', 14)}</button>`;
    const indicL = window.BTSIndic && typeof BTSIndic.setLayout === 'function' && Array.isArray(BTSIndic.layouts) && BTSIndic.layouts.length;
    let indicDef = ''; try { indicDef = localStorage.getItem('bts-indic-default') || ''; } catch {}
    main.innerHTML = `<h1>Options</h1>
      <div class="opt-page">
        <div class="opt-nav" role="navigation" aria-label="Options sections">${OPT_SECTIONS.map(([id, l, ic], i) => `<button data-opt-sec="${id}"${i === 0 ? ' class="on"' : ''}>${icon(ic, 16)}<span>${l}</span></button>`).join('')}</div>
        <div class="opt-body">
          <section class="opt-sec" id="opt-general"><h2>General</h2><div class="opt-card">
            ${row('Your name', 'Shown as the author of your files, comments and changes.', `<input class="opt-in" id="bts-user-name" value="${esc(BTS.userName)}" placeholder="Your name" spellcheck="false">`)}
            ${row('Appearance', 'System follows your Windows light or dark setting.', `<div class="opt-seg">${['light', 'dark', 'system'].map(t => `<button data-set-theme="${t}"${BTS.theme === t ? ' class="on"' : ''}>${icon(t === 'light' ? 'sun' : t === 'dark' ? 'moon' : 'monitor', 14)}${t[0].toUpperCase() + t.slice(1)}</button>`).join('')}</div>`)}
            ${row('When an app opens', 'What you see when you start Docs, Sheets or Slides.', sel('startScreen', [['false', 'Open a blank document'], ['true', 'Show the start screen']], !!P.startScreen))}
            ${row('Zoom', 'Starting zoom for Docs and Sheets.', sel('zoom', [75, 90, 100, 110, 125, 150, 175, 200].map(z => [z, z + '%' + (z === 100 ? ' (default)' : '')]), P.zoom || 100))}
          </div></section>
          <section class="opt-sec" id="opt-proofing"><h2>Proofing</h2><div class="opt-card" id="opt-spell">
            ${row('Check spelling as I type', 'Underlines words that may be misspelt. Right-click a word for suggestions.', sw('opt-spell-on', P.spellCheck !== false, desktop ? '' : 'disabled'))}
            ${row('Spelling languages', desktop ? 'Words are accepted if they are right in any chosen language.' : 'Available in the desktop app.', `<div class="opt-langs" id="opt-spell-langs">${desktop ? '<span class="opt-note">Loading…</span>' : ''}</div>`)}
          </div></section>
          <section class="opt-sec" id="opt-save"><h2>Save</h2><div class="opt-card">
            ${['docs', 'sheets', 'slides'].map(a => row(`Save new ${a === 'docs' ? 'documents' : a === 'sheets' ? 'workbooks' : 'presentations'} as`, APPS[a].name + (a === 'docs' ? ' · Save a copy can still use any format.' : ''), sel('defaultFormat_' + a, fmtOpts(a), P['defaultFormat_' + a] || fmtOpts(a)[0][0]))).join('')}
            ${row('Save recovery info', 'If BTS Office closes by accident, you can get your work back the next time you open the app.', sel('autosaveSeconds', [[0, 'Soon after each change (recommended)'], [10, 'Every 10 seconds'], [30, 'Every 30 seconds'], [60, 'Every minute'], [120, 'Every 2 minutes'], [300, 'Every 5 minutes'], [600, 'Every 10 minutes']], P.autosaveSeconds || 0))}
            ${row('Keep version history', 'Keeps a copy of .docx, .xlsx and .pptx files every time you save (the last 30). See File &gt; Info.', sw('bts-versions-on', st.versionHistory !== false, desktop ? '' : 'disabled'))}
          </div></section>
          <section class="opt-sec" id="opt-new"><h2>New documents</h2><div class="opt-card">
            ${row('BTS Docs font', 'Body text font and size for new documents. Normally Calibri 11.', `<div class="opt-pair">${fontBtn('docsFont')}${sel('docsSize', sizes, P.docsSize || 0)}</div>`)}
            ${row('BTS Sheets font', 'Cell font and size for new workbooks. Normally Calibri 11.', `<div class="opt-pair">${fontBtn('sheetsFont')}${sel('sheetsSize', sizes, P.sheetsSize || 0)}</div>`)}
            ${row('Paper size', 'Page size for new documents in BTS Docs.', sel('pageSize', [['A4', 'A4 (21 × 29.7 cm)'], ['Letter', 'Letter (8.5 × 11 in)']], P.pageSize || 'A4'))}
            ${row('Measure in', 'Units for the ruler and margins in BTS Docs.', `<div class="opt-seg">${[['cm', 'Centimetres'], ['in', 'Inches']].map(([v, l]) => `<button data-set-units="${v}"${(P.units || 'cm') === v ? ' class="on"' : ''}>${l}</button>`).join('')}</div>`)}
          </div></section>
          <section class="opt-sec" id="opt-language"><h2>Language</h2><div class="opt-card">
            ${window.BTSIndic ? row('Default typing language', 'Language each app starts typing in. Ctrl+Alt+H / Ctrl+Alt+M switch any time.', `<select class="opt-sel" id="opt-indic-def">${[['', 'English'], ['hi', 'हिंदी — Hindi'], ['mr', 'मराठी — Marathi']].map(([v, l]) => `<option value="${v}"${v === indicDef ? ' selected' : ''}>${l}</option>`).join('')}</select>`) : ''}
            ${indicL ? row('Hindi/Marathi keyboard', 'How keys turn into Devanagari letters.', `<div class="opt-pair"><select class="opt-sel" id="opt-indic-layout">${BTSIndic.layouts.map(l => `<option value="${esc(l.id)}"${l.id === BTSIndic.layout ? ' selected' : ''} title="${esc(l.desc || '')}">${esc(l.label)}</option>`).join('')}</select>${typeof BTSIndic.showKeyboard === 'function' ? `<button class="btn icon-only" id="opt-indic-kb" title="Show the on-screen keyboard">${icon('keyboard', 15)}</button>` : ''}</div>`, '') + `<div class="opt-sub" id="opt-indic-desc">${esc((BTSIndic.layouts.find(l => l.id === BTSIndic.layout) || BTSIndic.layouts[0]).desc || '')}</div>` : ''}
          </div>
          ${window.BTSIndic || window.BTSVoice || window.BTSLegacyFonts ? `<h2>Indian languages &amp; voice typing</h2>
          <div class="bs-actions" style="flex-direction:row;flex-wrap:wrap">${window.BTSIndic ? `<button class="btn" data-act="indic-help">${icon('languages', 15)} Hindi &amp; Marathi typing</button>` : ''}${window.BTSLegacyFonts ? `<button class="btn" data-act="legacy">${icon('replace', 15)} Convert old fonts (Kruti Dev, Shivaji)</button>` : ''}${window.BTSVoice ? `<button class="btn" data-act="voice-try">${icon('mic', 15)} Voice typing</button>` : ''}</div>
          ${window.BTSVoice && desktop ? `<p class="empty-note" style="padding:8px 0 0;max-width:620px" id="bts-voice-models">Voice typing works offline after a one-time download of a speech model for each language.</p>` : ''}` : ''}
          </section>
        </div>
      </div>`;
    const body = main.querySelector('.opt-body');
    // mini-nav: jump to a section, highlight the one in view
    main.querySelectorAll('[data-opt-sec]').forEach(b => b.addEventListener('click', () => { const s = main.querySelector('#opt-' + b.dataset.optSec); main.scrollTo({ top: s.offsetTop - 24, behavior: 'smooth' }); }));
    const spy = () => { let cur = OPT_SECTIONS[0][0]; for (const [id] of OPT_SECTIONS) { const s = main.querySelector('#opt-' + id); if (s && s.getBoundingClientRect().top - main.getBoundingClientRect().top < 120) cur = id; } if (main.scrollTop + main.clientHeight >= main.scrollHeight - 4) cur = OPT_SECTIONS[OPT_SECTIONS.length - 1][0]; main.querySelectorAll('[data-opt-sec]').forEach(b => b.classList.toggle('on', b.dataset.optSec === cur)); };
    main.onscroll = () => { if (main.querySelector(".opt-page")) spy(); };
    const saved = () => BTS.toast('Saved', { ms: 1200 });
    // simple selects bound to BTS.prefs
    body.querySelectorAll('select[data-pref]').forEach(s => s.addEventListener('change', () => {
      const k = s.dataset.pref; let v = s.value;
      if (k === 'startScreen') v = v === 'true'; else if (/^(zoom|autosaveSeconds|docsSize|sheetsSize)$/.test(k)) v = +v;
      BTS.setPref(k, v); saved();
    }));
    const un = main.querySelector('#bts-user-name'); un.addEventListener('change', () => { BTS.userName = un.value.trim(); try { localStorage.setItem('bts-user-name', BTS.userName); } catch {} if (desktop) bts.setSettings({ userName: BTS.userName }); saved(); });
    body.querySelectorAll('[data-set-theme]').forEach(b => b.addEventListener('click', () => { BTS.setTheme(b.dataset.setTheme); body.querySelectorAll('[data-set-theme]').forEach(x => x.classList.toggle('on', x === b)); }));
    body.querySelectorAll('[data-set-units]').forEach(b => b.addEventListener('click', () => { BTS.setPref('units', b.dataset.setUnits); body.querySelectorAll('[data-set-units]').forEach(x => x.classList.toggle('on', x === b)); saved(); }));
    body.querySelectorAll('[data-font-pref]').forEach(b => b.addEventListener('click', async () => {
      const k = b.dataset.fontPref; let f;
      if (window.BTS_FONTS) f = await BTS.fontPicker(b, { current: BTS.prefs[k] || 'Calibri' }); else f = await BTS.prompt('Font', 'Font name', BTS.prefs[k] || 'Calibri');
      if (!f) return; f = f.trim(); const v = f === 'Calibri' ? '' : f;
      BTS.setPref(k, v); b.innerHTML = esc(v || 'Default (Calibri)') + icon('chevron-down', 14); b.style.fontFamily = v ? (window.btsFontStack ? btsFontStack(v) : `'${v}'`).replace(/"/g, "'") : 'inherit'; saved();
    }));
    const vo = main.querySelector('#bts-versions-on'); if (desktop) vo.onchange = () => { bts.setSettings({ versionHistory: vo.checked }); saved(); };
    // proofing (main process: main/prefs.js)
    const spOn = main.querySelector('#opt-spell-on'); const langsEl = main.querySelector('#opt-spell-langs');
    const applySpell = async () => { try { return await bts.ext('prefs-apply'); } catch (e) { console.error(e); return null; } };
    if (desktop) {
      spOn.onchange = async () => { await BTS.setPref('spellCheck', spOn.checked); await applySpell(); saved(); };
      let sp = null; try { sp = await bts.ext('prefs-spell'); } catch {}
      if (!sp) langsEl.innerHTML = '<span class="opt-note">Not available.</span>';
      else if (!sp.choose) langsEl.innerHTML = '<span class="opt-note">Your Mac picks spelling languages from System Settings.</span>';
      else {
        const cur = Array.isArray(P.spellLanguages) ? P.spellLanguages : sp.languages;
        langsEl.innerHTML = sp.offer.map(l => `<label class="opt-chip"><input type="checkbox" value="${esc(l.id)}"${cur.includes(l.id) ? ' checked' : ''}><span>${esc(l.label)}</span></label>`).join('') || '<span class="opt-note">No spelling dictionaries found.</span>';
        langsEl.addEventListener('change', async (e) => {
          const list = [...langsEl.querySelectorAll('input:checked')].map(i => i.value);
          if (!list.length) { e.target.checked = true; BTS.toast('Keep at least one language.', { error: true }); return; }
          await BTS.setPref('spellLanguages', list); await applySpell(); saved();
        });
      }
    }
    // language
    const idf = main.querySelector('#opt-indic-def'); if (idf) idf.onchange = () => { try { if (idf.value) localStorage.setItem('bts-indic-default', idf.value); else localStorage.removeItem('bts-indic-default'); } catch {} if (window.BTSIndic) BTSIndic.enable(idf.value || false); saved(); };
    const il = main.querySelector('#opt-indic-layout'); if (il) il.onchange = () => { try { BTSIndic.setLayout(il.value); } catch (e) { console.error(e); } const d = main.querySelector('#opt-indic-desc'); const l = BTSIndic.layouts.find(x => x.id === il.value); if (d) d.textContent = (l && l.desc) || ''; saved(); };
    main.querySelector('#opt-indic-kb')?.addEventListener('click', () => { closeBackstage(); BTSIndic.showKeyboard(); });
    main.querySelector('[data-act="indic-help"]')?.addEventListener('click', indicHelp);
    main.querySelector('[data-act="legacy"]')?.addEventListener('click', () => BTSLegacyFonts.dialog());
    main.querySelector('[data-act="voice-try"]')?.addEventListener('click', () => { closeBackstage(); BTSVoice.toggle(); });
    const vm = main.querySelector('#bts-voice-models');
    if (vm) { const showModels = async () => { try { const r = await BTSVoice.models(); const got = (r.models || []).filter(m => m.downloaded); vm.innerHTML = got.length ? `Speech models on this computer: ${got.map(m => esc(m.label)).join(', ')} (${Math.max(1, Math.round(r.bytes / 1048576))} MB). <a href="#" data-act="voice-remove" style="color:var(--accent);font-weight:600">Remove</a>` : 'Voice typing works offline after a one-time download of a speech model for each language (about 40 MB).'; const rm = vm.querySelector('[data-act="voice-remove"]'); if (rm) rm.onclick = async (ev) => { ev.preventDefault(); const freed = await BTSVoice.removeModels(); BTS.toast(`Removed speech models (${Math.round(freed / 1048576)} MB freed)`); showModels(); }; } catch { } }; showModels(); }
  }
  function greet() { const h = new Date().getHours(); return h < 12 ? 'Good morning' : h < 17 ? 'Good afternoon' : 'Good evening'; }
  async function renderRecent(el, n) {
    if (!el) return;
    if (!desktop) { el.innerHTML = '<div class="empty-note">Recent files appear here in the desktop app.</div>'; return; }
    const list = (await bts.recent()).slice(0, n);
    if (!list.length) { el.innerHTML = '<div class="empty-note">Files you open or save will show up here.</div>'; return; }
    el.innerHTML = list.map(r => { const e = extOf(r.name); return `<button class="bs-row" data-path="${esc(r.path)}"><div class="fi" style="background:${FILE_COLOR[e] || '#47526a'}">${esc(e.toUpperCase().slice(0, 4))}</div><div><div>${esc(r.name)}</div><div class="meta">${esc(r.path)}</div></div><div class="meta" style="margin-left:auto">${new Date(r.time).toLocaleDateString()}</div></button>`; }).join('');
  }
  function appBadge(app, s) { const a = APPS[app] || { color: '#172033', letter: 'B' }; return `<span style="width:${s}px;height:${s}px;border-radius:${s * .26}px;background:${a.color};color:#fff;display:inline-grid;place-items:center;font-weight:800;font-size:${s * .46}px;flex:none">${a.letter}</span>`; }
  BTS.appBadge = appBadge;

  function buildBackstage() {
    const b = document.createElement('div'); b.id = 'backstage';
    b.innerHTML = `<nav>
      <button class="bs-back" data-act="close">${icon('arrow-left')} Back</button>
      <button data-view="home">${icon('house')} Home</button>
      <button data-view="new">${icon('file-plus')} New</button>
      <button data-view="open">${icon('folder-open')} Open</button>
      <button data-view="info">${icon('info')} Info</button>
      <button data-act="save">${icon('save')} Save</button>
      <button data-view="saveas">${icon('save-all')} Save a copy</button>
      <button data-view="export">${icon('file-output')} Export</button>
      <button data-view="print">${icon('printer')} Print</button>
      <div class="bs-gap"></div>
      <button data-view="about">${icon('badge-info')} About</button>
      <button data-view="options">${icon('settings')} Options</button>
      <button data-act="closewin">${icon('x')} Close</button></nav><div class="bs-main"></div>`;
    document.body.appendChild(b);
    b.addEventListener('click', async e => {
      const v = e.target.closest('[data-view]'); if (v) return renderBackstage(v.dataset.view);
      const a = e.target.closest('[data-act]');
      if (a) {
        const act = a.dataset.act;
        if (act === 'close') closeBackstage();
        else if (act === 'save') { closeBackstage(); BTS.save(); }
        else if (act === 'browse') BTS.open();
        else if (act === 'blank') { closeBackstage(); BTS.hasContent?.() || BTS.dirty ? BTS.newWindow() : (BTS.resetProps(), BTS.cfg.newDocument()); }
        else if (act === 'print') BTS.print();
        else if (act === 'pdf') BTS.exportPDF();
        else if (act === 'closewin') { closeBackstage(); window.close(); }
        return;
      }
      const t = e.target.closest('[data-tpl]'); if (t) { closeBackstage(); if (BTS.dirty && !(await BTS.confirm('Replace current file?', 'Your unsaved changes will be lost.', 'Replace'))) return; BTS.cfg.templates[+t.dataset.tpl].apply(); BTS.title = BTS.cfg.templates[+t.dataset.tpl].name; BTS.filePath = null; BTS.markDirty(); return; }
      const ap = e.target.closest('[data-app]'); if (ap) { closeBackstage(); BTS.newWindow(ap.dataset.app); return; }
      const f = e.target.closest('[data-fmt]'); if (f) { if (f.hasAttribute('data-pdf')) BTS.exportPDF(); else BTS.saveAs(f.dataset.fmt); return; }
      const r = e.target.closest('[data-path]'); if (r) { closeBackstage(); const file = await bts.readFile(r.dataset.path); if (extOf(file.name || '') && appOfExt(extOf(file.name)) === BTS.cfg.app && !BTS.dirty && !BTS.hasContent?.()) loadFile(file); else bts.openApp(null, r.dataset.path); return; }
      const ex = e.target.closest('[data-ext]'); if (ex) { e.preventDefault(); if (desktop) bts.openExternal(ex.dataset.ext); else window.open(ex.dataset.ext); return; }
      const vv = e.target.closest('[data-version]'); if (vv) { bts.openVersion(BTS.filePath, vv.dataset.version); return; }
    });
  }
  function appOfExt(e) { return ({ docx: 'docs', txt: 'docs', html: 'docs', htm: 'docs', md: 'docs', odt: 'docs', rtf: 'docs', xlsx: 'sheets', xls: 'sheets', csv: 'sheets', tsv: 'sheets', ods: 'sheets', pptx: 'slides', pdf: 'pdf' })[e]; }
  BTS.openBackstage = (view = 'home') => { BTS.closePopover(); $('#backstage').classList.add('open'); renderBackstage(view); };

  /* ---------- title bar ---------- */
  function buildTitlebar() {
    const tb = document.createElement('header'); tb.id = 'titlebar';
    const cfg = BTS.cfg;
    tb.innerHTML = `<div class="tb-app"><span class="tb-logo">${cfg.letter}</span><span>${esc(cfg.name)}</span></div>
      <span class="tb-sep"></span>
      <button class="tb-btn" data-cmd="save" title="Save (Ctrl+S)">${icon('save', 16)}</button>
      <button class="tb-btn" data-cmd="undo" title="Undo (Ctrl+Z)">${icon('undo-2', 16)}</button>
      <button class="tb-btn" data-cmd="redo" title="Redo (Ctrl+Y)">${icon('redo-2', 16)}</button>
      <div class="tb-name"><input id="tb-docname" spellcheck="false" title="Rename"><span class="tb-state" id="tb-state"></span></div>
      <span style="margin-left:auto"></span>
      ${window.BTSIndic && cfg.app !== 'docs' ? `<button class="tb-btn tb-indic" data-cmd="indic-typing" data-search="Type in Hindi or Marathi (phonetic typing)" data-search-icon="languages" title="Type in Hindi or Marathi (Ctrl+Alt+H / Ctrl+Alt+M)"><span class="tb-indic-l">अ</span></button>` : ''}
      ${window.BTSVoice ? `<button class="tb-btn tb-voice" data-cmd="voice-typing" data-search="Voice typing (dictate)" data-search-icon="mic" title="Voice typing (Alt+Shift+D)">${icon('mic', 16)}</button>` : ''}
      <button class="tb-btn tb-tell" data-cmd="command-search" title="Search commands (Alt+Q)">${icon('lightbulb', 16)}</button>
      ${cfg.search ? `<label class="tb-search" style="margin-left:6px">${icon('search', 15)}<input id="tb-search" placeholder="Search in ${esc(cfg.defaultTitle.toLowerCase())}"></label>` : ''}`;
    document.body.prepend(tb);
    const inp = $('#tb-docname', tb);
    inp.addEventListener('change', () => { const v = inp.value.trim(); if (v && v !== BTS.title) { BTS.title = v.replace(/[\\/:*?"<>|]/g, ''); BTS.filePath = null; BTS.markDirty(); } updateTitle(); });
    inp.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === 'Escape') { if (e.key === 'Escape') inp.value = BTS.title; inp.blur(); BTS.cfg.focus && BTS.cfg.focus(); } });
    const s = $('#tb-search', tb);
    if (s) s.addEventListener('keydown', e => { if (e.key === 'Enter') { cfg.search(s.value); } if (e.key === 'Escape') { s.value = ''; s.blur(); BTS.cfg.focus && BTS.cfg.focus(); } });
    tb.addEventListener('click', e => { const c = e.target.closest('[data-cmd]'); if (c) dispatch(c.dataset.cmd, c, e); });
    if (window.BTSIndic && BTSIndic.onChange) { const upd = () => { const b = $('.tb-indic', tb); if (!b) return; const on = BTSIndic.isOn(); b.classList.toggle('on', on); b.querySelector('.tb-indic-l').textContent = on ? (BTSIndic.lang === 'mr' ? 'म' : 'हि') : 'अ'; b.title = on ? (BTSIndic.lang === 'mr' ? 'Typing in Marathi' : 'Typing in Hindi') + ' — click to change (Ctrl+Alt+H / Ctrl+Alt+M)' : 'Type in Hindi or Marathi (Ctrl+Alt+H / Ctrl+Alt+M)'; }; BTSIndic.onChange(upd); upd(); }
    if (navigator.platform.startsWith('Mac')) document.documentElement.classList.add('mac');
    if (desktop) bts.titlebarColor(getComputedStyle(document.body).getPropertyValue('--accent').trim() || '#3C4FE0', '#ffffff');
  }

  /* ---------- ribbon ---------- */
  BTS.showTab = (name) => {
    if (name === 'file') return BTS.openBackstage();
    $$('#ribbon-tabs [data-tab]').forEach(b => b.classList.toggle('active', b.dataset.tab === name));
    $$('#ribbon .ribbon-panel').forEach(p => p.classList.toggle('active', p.dataset.panel === name));
    document.body.classList.remove('ribbon-collapsed');
    BTS.activeTab = name;
  };
  function wireRibbon() {
    const tabs = $('#ribbon-tabs');
    if (tabs) {
      tabs.addEventListener('mousedown', e => { if (e.target.closest('[data-tab]')) e.preventDefault(); });
      tabs.addEventListener('click', e => { const t = e.target.closest('[data-tab]'); if (t) BTS.showTab(t.dataset.tab); const c = e.target.closest('[data-cmd]'); if (c) dispatch(c.dataset.cmd, c, e); });
      tabs.addEventListener('dblclick', e => { if (e.target.closest('[data-tab]') && !e.target.closest('.file-tab')) document.body.classList.toggle('ribbon-collapsed'); });
      const first = $('[data-tab]:not(.file-tab)', tabs); if (first) BTS.showTab(first.dataset.tab);
    }
    // Ribbon buttons must not steal focus from the editor (keeps text selection).
    document.addEventListener('mousedown', e => { const b = e.target.closest('#ribbon .rb, #statusbar .sb-btn, .keepfocus'); if (b) e.preventDefault(); }, true);
    document.addEventListener('click', e => {
      const c = e.target.closest('[data-cmd]');
      if (!c || c.closest('#titlebar') || c.closest('#ribbon-tabs') || c.closest('#backstage')) return;
      if (c.tagName === 'SELECT' || c.tagName === 'INPUT') return;
      dispatch(c.dataset.cmd, c, e);
    });
    document.addEventListener('change', e => { const c = e.target.closest('select[data-cmd], input[data-cmd]'); if (c) dispatch(c.dataset.cmd, c, e); });
  }

  function dispatch(cmd, el, ev) {
    switch (cmd) {
      case 'file': return BTS.openBackstage();
      case 'save': return BTS.save();
      case 'saveas': return BTS.saveAs();
      case 'open': return BTS.open();
      case 'new': return BTS.newWindow();
      case 'print': return BTS.print();
      case 'export-pdf': return BTS.exportPDF();
      case 'command-search': return BTS.commandSearch();
      case 'voice-typing': if (window.BTSVoice) BTSVoice.toggle(); return;
      case 'indic-typing': return indicMenu(el);
      case 'legacy-fonts': if (window.BTSLegacyFonts) BTSLegacyFonts.dialog(); return;
      case 'toggle-ribbon': document.body.classList.toggle('ribbon-collapsed'); return;
      case 'theme-toggle': BTS.setTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark'); return;
      case 'zoom-in': return setZoom(BTS.zoom + 10);
      case 'zoom-out': return setZoom(BTS.zoom - 10);
      case 'zoom-100': return setZoom(100);
    }
    try { BTS.cfg.onCommand && BTS.cfg.onCommand(cmd, el, ev); }
    catch (err) { console.error(err); BTS.toast(String(err.message || err), { error: true }); }
  }
  BTS.dispatch = dispatch;
  function indicMenu(el) {
    if (!window.BTSIndic) return;
    const cur = BTSIndic.isOn() ? BTSIndic.lang : '';
    const tick = (v) => (cur === v ? 'check' : '');
    BTS.menu(el, [
      { head: 'Type with English letters' },
      { label: 'English (normal typing)', icon: tick('') || 'type', action: () => BTSIndic.enable(false) },
      { label: 'हिंदी — Hindi', icon: tick('hi') || 'languages', kbd: 'Ctrl+Alt+H', action: () => BTSIndic.enable('hi') },
      { label: 'मराठी — Marathi', icon: tick('mr') || 'languages', kbd: 'Ctrl+Alt+M', action: () => BTSIndic.enable('mr') },
      { sep: true },
      ...(window.BTSLegacyFonts ? [{ label: 'Convert old fonts (Kruti Dev, Shivaji)…', icon: 'replace', action: () => BTSLegacyFonts.dialog() }] : []),
      { label: 'How phonetic typing works', icon: 'circle-help', action: indicHelp },
    ]);
  }
  function indicHelp() {
    const rows = [['namaste', 'नमस्ते'], ['bhaarat / bharat', 'भारत'], ['mahaaraashTra', 'महाराष्ट्र'], ['kRShNa', 'कृष्ण'], ['shrI', 'श्री'], ['dnyaan', 'ज्ञान'], ['Laaj (Marathi)', 'ळाज']];
    BTS.dialog({ title: 'Typing in Hindi and Marathi', body: `<p style="margin-top:0;line-height:1.55">Type words the way they sound, using English letters. Each word turns into Devanagari as you type and is finished when you press Space, Enter or a punctuation mark. Backspace works letter by letter.</p>
      <div style="display:grid;grid-template-columns:auto 1fr;gap:6px 18px;margin:10px 0 14px">${rows.map(([a, b]) => `<code>${esc(a)}</code><b style="font-size:16px">${b}</b>`).join('')}</div>
      <p style="line-height:1.55;color:var(--ink-2);margin:0">Capital letters make the retroflex and long sounds (T, D, N, I, U, L). <b>Ctrl+Alt+H</b> turns Hindi on or off, <b>Ctrl+Alt+M</b> does the same for Marathi. Works offline in every BTS Office app.</p>`, buttons: [{ label: 'OK', value: 1, primary: true }] });
  }

  /* ---------- status bar ---------- */
  function setZoom(z) {
    z = Math.max(BTS.cfg.minZoom || 10, Math.min(BTS.cfg.maxZoom || 500, Math.round(z)));
    BTS.zoom = z; const r = $('#sb-zoom'); if (r) r.value = z; const v = $('#sb-zoom-val'); if (v) v.textContent = z + '%';
    BTS.cfg.onZoom && BTS.cfg.onZoom(z);
  }
  BTS.setZoom = setZoom;
  function buildStatusbar() {
    let sb = $('#statusbar');
    if (!sb) { sb = document.createElement('footer'); sb.id = 'statusbar'; document.body.appendChild(sb); }
    sb.insertAdjacentHTML('beforeend', `<span class="spacer"></span><span id="sb-right" class="sb-item"></span>` +
      (BTS.cfg.onZoom ? `<button class="sb-btn" data-cmd="zoom-out" title="Zoom out">${icon('minus', 14)}</button><input type="range" id="sb-zoom" min="${BTS.cfg.minZoom || 10}" max="${BTS.cfg.maxZoom || 300}" value="100" title="Zoom"><button class="sb-btn" data-cmd="zoom-in" title="Zoom in">${icon('plus', 14)}</button><button class="sb-btn zoom-val" id="sb-zoom-val" data-cmd="zoom-100" title="Reset zoom">100%</button>` : ''));
    const r = $('#sb-zoom'); if (r) r.addEventListener('input', () => setZoom(+r.value));
  }
  BTS.setStatus = (key, html) => {
    const sb = $('#statusbar'); let el = $('[data-sb="' + key + '"]', sb);
    if (!el) { el = document.createElement('span'); el.className = 'sb-item'; el.dataset.sb = key; sb.insertBefore(el, $('.spacer', sb)); }
    el.innerHTML = html;
  };
  BTS.setStatusRight = (html) => { const e = $('#sb-right'); if (e) e.innerHTML = html; };

  /* ---------- keyboard ---------- */
  function keys(e) {
    if (e.altKey && e.shiftKey && !e.ctrlKey && e.code === 'KeyD' && window.BTSVoice) { e.preventDefault(); BTSVoice.toggle(); return; }
    if ((e.altKey && !e.ctrlKey && e.key.toLowerCase() === 'q') || ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'p')) { e.preventDefault(); BTS.commandSearch(); return; }
    const mod = e.ctrlKey || e.metaKey; if (!mod) return;
    const k = e.key.toLowerCase();
    if (k === 's' && e.shiftKey) { e.preventDefault(); BTS.saveAs(); }
    else if (k === 's') { e.preventDefault(); BTS.save(); }
    else if (k === 'o') { e.preventDefault(); BTS.open(); }
    else if (k === 'n' && !e.shiftKey) { e.preventDefault(); BTS.newWindow(); }
    else if (k === 'p' && !e.shiftKey) { e.preventDefault(); BTS.print(); }
    else if (e.key === 'F1') { e.preventDefault(); document.body.classList.toggle('ribbon-collapsed'); }
    else if ((k === '=' || k === '+') && e.altKey) { e.preventDefault(); setZoom(BTS.zoom + 10); }
    else if (k === '-' && e.altKey) { e.preventDefault(); setZoom(BTS.zoom - 10); }
  }

  /* ---------- init ---------- */
  BTS.init = async (cfg) => {
    BTS.cfg = cfg = Object.assign({ defaultTitle: 'Untitled', saveFormats: [{ ext: 'txt', label: 'Text' }] }, cfg);
    // options: settings.json (desktop) is the source of truth; localStorage is the fast copy read at load
    // (merged after the shell is built: apps run code right after calling BTS.init without awaiting it)
    const settingsP = desktop ? bts.getSettings().catch(() => ({})) : Promise.resolve({});
    const mergePrefs = (settings) => { let ch = false; for (const k of Object.keys(PREF_DEFAULTS)) if (settings[k] !== undefined && JSON.stringify(settings[k]) !== JSON.stringify(BTS.prefs[k])) { BTS.prefs[k] = settings[k]; ch = true; } if (ch) { mirrorPrefs(); firePrefs(null); } };
    document.body.dataset.app = cfg.app;
    if (cfg.hasContent && !BTS.hasContent) BTS.hasContent = cfg.hasContent;
    BTS.title = cfg.defaultTitle;
    BTS.resetProps();
    buildTitlebar(); buildBackstage(); wireRibbon(); buildStatusbar();
    BTS.zoom = cfg.zoom || 100;
    if ((cfg.app === 'docs' || cfg.app === 'sheets') && +BTS.prefs.zoom && +BTS.prefs.zoom !== 100) BTS.zoom = +BTS.prefs.zoom;
    if (cfg.onZoom) setZoom(BTS.zoom);
    document.addEventListener('keydown', keys);
    $$('#ribbon-tabs .file-tab').forEach(b => b.addEventListener('click', () => BTS.openBackstage()));
    updateTitle();
    if (desktop) {
      bts.onSaveThenClose(async () => { const ok = await BTS.save(); if (ok) bts.closeNow(); });
      const s = await settingsP; mergePrefs(s); if (s.theme && !localStorage.getItem('bts-theme')) applyTheme(s.theme);
      if (s.userName && !BTS.userName) BTS.userName = s.userName;
      const f = await bts.initialFile(); if (f) await loadFile(f);
    }
    // File > Options > General: show the start screen instead of a blank document (not when a file was opened)
    if (BTS.prefs.startScreen && cfg.app !== 'pdf' && !BTS.loadedExt && !BTS.filePath) BTS.openBackstage('home');
    // File > Options > Language: default typing language
    { let d = ''; try { d = localStorage.getItem('bts-indic-default') || ''; } catch {} if (d && window.BTSIndic && BTSIndic.enable && !BTSIndic.isOn?.()) { try { BTSIndic.enable(d); } catch (e) { console.error(e); } } }
    // drag & drop files onto the window
    document.addEventListener('dragover', e => { if (e.dataTransfer.types.includes('Files') && !e.target.closest('[data-own-drop]')) e.preventDefault(); });
    document.addEventListener('drop', async e => {
      if (e.target.closest('[data-own-drop]')) return;
      const f = e.dataTransfer.files[0]; if (!f) return; e.preventDefault();
      const ext = extOf(f.name);
      const fp = desktop ? bts.pathFor(f) : '';
      if (appOfExt(ext) && appOfExt(ext) !== cfg.app && fp) { bts.openApp(null, fp); return; }
      if (cfg.onDropFile && cfg.onDropFile(f, e)) return;
      if (appOfExt(ext) === cfg.app) { if ((BTS.dirty || BTS.hasContent?.()) && fp) { bts.openApp(null, fp); return; } loadFile({ name: f.name, path: fp || undefined, data: new Uint8Array(await f.arrayBuffer()) }); }
    });
    cfg.ready && cfg.ready();
  };
})();
