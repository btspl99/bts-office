/* BTS PDF — Organize pages: page grid (drag to reorder, rotate, delete, duplicate, insert), quick page
 * actions, split and extract. Changes are applied to the PDF bytes with pdf-lib, annotations are moved
 * along with their pages, then the document is reloaded. */
import { S, $, $$, esc, clamp, newId, loadPdfLib, pdfjsLib, APP, baseViewport } from './state.js';
import * as IO from './io.js';
import { buildImagesPdf, pickImageFiles } from './images.js';

/* ======================= page ranges ======================= */
/** "1-3, 5, 8-" -> [[0,1,2],[4],[7..n-1]] (0-based). Throws a friendly error. */
export function parseRanges(str, n) {
  const groups = [];
  for (const part of String(str).split(/[,;]+/).map(s => s.trim()).filter(Boolean)) {
    const m = /^(\d*)\s*(?:[-–—]|to)\s*(\d*|end)$/i.exec(part) || /^(\d+)$/.exec(part);
    if (!m) throw new Error(`“${part}” isn’t a page range. Use numbers like 1-3, 5`);
    let a, b;
    if (m.length === 2) a = b = +m[1];
    else { a = m[1] ? +m[1] : 1; b = !m[2] || /end/i.test(m[2]) ? n : +m[2]; }
    if (a < 1 || b < 1 || a > n || b > n) throw new Error(`This PDF has ${n} page${n === 1 ? '' : 's'} — “${part}” is outside that`);
    const g = []; if (a <= b) for (let i = a; i <= b; i++) g.push(i - 1); else for (let i = a; i >= b; i--) g.push(i - 1);
    groups.push(g);
  }
  if (!groups.length) throw new Error('Type the pages you want, for example 1-3, 5');
  return groups;
}
const rangeLabel = (g) => { const a = g[0] + 1, b = g[g.length - 1] + 1; return g.length === 1 ? `page ${a}` : `pages ${a}-${b}`; };

/* ======================= applying a page list ======================= */
// item: { key, src: 0 (this PDF) | n (inserted PDF), idx, rot, blank?: {w,h} }
let sources = [];   // index 0 = current document; others: { bytes, doc (pdf.js), name }

const itemsFromDoc = () => S.pages.map(p => ({ key: newId(), src: 0, idx: p.i, rot: 0 }));

/** Build the new PDF from items, remap annotations and reload. */
export async function applyItems(items, { message } = {}) {
  if (!items.length) throw new Error('A PDF needs at least one page');
  const L = await loadPdfLib();
  const base = await IO.buildBaseBytes();
  let doc;
  try { doc = await L.PDFDocument.load(base, { updateMetadata: false }); }
  catch { throw new Error('This PDF is password-protected, so its pages can’t be changed.'); }
  const orig = doc.getPages();
  const ext = new Map();
  const used = new Set();
  const out = [];
  for (const it of items) {
    let page;
    if (it.blank) { page = L.PDFPage.create(doc); page.setSize(it.blank.w, it.blank.h); }
    else if (it.src === 0) {
      if (!used.has(it.idx)) { page = orig[it.idx]; used.add(it.idx); }
      else [page] = await doc.copyPages(doc, [it.idx]);
    } else {
      if (!ext.has(it.src)) {
        try { ext.set(it.src, await L.PDFDocument.load(sources[it.src].bytes)); }
        catch { throw new Error(`“${sources[it.src].name}” is password-protected or damaged.`); }
      }
      [page] = await doc.copyPages(ext.get(it.src), [it.idx]);
    }
    if (it.rot) page.setRotation(L.degrees(((page.getRotation().angle + it.rot) % 360 + 360) % 360));
    out.push(page);
  }
  for (let i = doc.getPageCount() - 1; i >= 0; i--) doc.removePage(i);
  out.forEach(pg => doc.addPage(pg));
  const bytes = await doc.save({ useObjectStreams: false });

  // move annotations with their pages
  const anns = []; const seen = new Set();
  items.forEach((it, j) => {
    if (it.blank || it.src !== 0) return;
    for (const a of S.anns) {
      if (a.p !== it.idx) continue;
      const b = seen.has(a.id) ? { ...structuredClone(a), id: newId() } : structuredClone(a);
      seen.add(a.id); b.p = j;
      if (b.type === 'image' && (((it.rot % 180) + 180) % 180) === 90) {   // keep pictures upright and unstretched
        const [x1, y1, x2, y2] = b.r; const cx = (x1 + x2) / 2, cy = (y1 + y2) / 2, hw = (y2 - y1) / 2, hh = (x2 - x1) / 2;
        b.r = [cx - hw, cy - hh, cx + hw, cy + hh];
      }
      anns.push(b);
    }
  });
  S.anns = anns; S.sel = null; S.undo = []; S.redo = [];
  S.modified = true;
  await APP.openBytes(bytes, { keepAnns: true, keepView: true });
  BTS.markDirty();
  if (message) BTS.toast(message);
}

/* ======================= quick actions (ribbon) ======================= */
const run = (fn) => fn().catch(e => { console.warn(e); BTS.toast(e.message || String(e), { error: true, ms: 4500 }); });
export const rotatePage = (deg) => run(async () => { const it = itemsFromDoc(); it[S.cur].rot = deg; const keep = S.cur; await applyItems(it); APP.goTo(keep); });
export const deletePage = () => run(async () => {
  if (S.pages.length < 2) throw new Error('A PDF needs at least one page');
  const keep = S.cur;
  if (S.anns.some(a => a.p === keep) && !(await BTS.confirm('Delete page', `Page ${keep + 1} has annotations. Delete the page and its annotations?`, 'Delete'))) return;
  const it = itemsFromDoc(); it.splice(keep, 1); await applyItems(it, { message: `Page ${keep + 1} deleted` }); APP.goTo(Math.min(keep, S.pages.length - 1));
});
export const duplicatePage = () => run(async () => { const it = itemsFromDoc(); const keep = S.cur; it.splice(keep + 1, 0, { ...it[keep], key: newId() }); await applyItems(it, { message: 'Page duplicated' }); APP.goTo(keep + 1); });
export const insertBlank = () => run(async () => {
  const vp = S.pages[S.cur].proxy.getViewport({ scale: 1, rotation: S.pages[S.cur].proxy.rotate });
  const it = itemsFromDoc(); const keep = S.cur; it.splice(keep + 1, 0, { key: newId(), blank: { w: vp.width, h: vp.height }, rot: 0 });
  await applyItems(it, { message: 'Blank page added' }); APP.goTo(keep + 1);
});

/** Ribbon "Insert pictures": add pictures as new pages after the current page. */
export const insertImages = () => run(async () => {
  const files = await pickImageFiles(); if (!files.length) return;
  const vp = S.pages[S.cur].proxy.getViewport({ scale: 1 });
  const bytes = await buildImagesPdf(files, { size: Math.abs(vp.width - 612) < 3 || Math.abs(vp.height - 612) < 3 ? 'letter' : 'a4', orient: 'auto', margin: 'small', shrink: true });
  const L = await loadPdfLib(); const n = (await L.PDFDocument.load(bytes)).getPageCount();
  sources = [{ doc: S.doc, name: BTS.title }, { bytes, name: 'pictures' }];
  const it = itemsFromDoc(); const keep = S.cur;
  it.splice(keep + 1, 0, ...Array.from({ length: n }, (_, i) => ({ key: newId(), src: 1, idx: i, rot: 0 })));
  try { await applyItems(it, { message: `Added ${n} picture page${n === 1 ? '' : 's'}` }); } finally { sources = []; }
  APP.goTo(keep + 1);
});

/* ======================= picking files ======================= */
const pickPdf = () => BTS.pickFile('.pdf,application/pdf');
async function addSource(bytes, name) {
  const doc = await pdfjsLib.getDocument({ data: bytes.slice(), isEvalSupported: false }).promise;
  sources.push({ bytes, doc, name });
  return sources.length - 1;
}

/* ======================= page grid ======================= */
let G = null;   // { items, sel:Set(key), anchor, root, io, thumbs: Map }

export async function openGrid() {
  if (!S.doc) return;
  if (G) return;
  sources = [{ doc: S.doc, name: BTS.title }];
  G = { items: itemsFromDoc(), sel: new Set(), anchor: null, thumbs: new Map(), dirty: false };
  const root = document.createElement('div'); root.id = 'organize'; root.tabIndex = 0;
  root.innerHTML = `
    <header class="org-bar">
      <div class="org-title"><b>Organize pages</b><span class="org-count"></span></div>
      <div class="org-tools">
        <button class="rb md" data-org="rotl" data-icon="rotate-ccw" title="Rotate selected pages left">Left</button>
        <button class="rb md" data-org="rotr" data-icon="rotate-cw" title="Rotate selected pages right">Right</button>
        <button class="rb md" data-org="dup" data-icon="copy" title="Duplicate selected pages">Duplicate</button>
        <button class="rb md" data-org="del" data-icon="trash-2" title="Delete selected pages (Delete)">Delete</button>
        <span class="org-sep"></span>
        <button class="rb md" data-org="blank" data-icon="file-plus-2" title="Insert a blank page after the selection">Blank page</button>
        <button class="rb md" data-org="pdf" data-icon="file-input" title="Insert the pages of another PDF after the selection">From PDF</button>
        <button class="rb md" data-org="img" data-icon="image-plus" title="Insert pictures as new pages after the selection">From images</button>
        <span class="org-sep"></span>
        <button class="rb md" data-org="extract" data-icon="file-output" title="Save the selected pages as a new PDF">Extract</button>
        <button class="rb md" data-org="split" data-icon="scissors" title="Split this PDF into several files">Split</button>
        <button class="rb md" data-org="all" data-icon="square-check-big" title="Select all pages (Ctrl+A)">Select all</button>
      </div>
      <div class="org-end"><button class="btn" data-org="cancel">Cancel</button><button class="btn primary" data-org="apply">Apply changes</button></div>
    </header>
    <div class="org-grid" role="listbox" aria-multiselectable="true"></div>
    <div class="org-hint">Drag pages to change their order · click to select, Ctrl or Shift + click to select several</div>`;
  $('#viewer-wrap').appendChild(root);
  G.root = root;
  root.querySelectorAll('[data-icon]').forEach(el => { if (!el.querySelector('svg')) el.insertAdjacentHTML('afterbegin', icon(el.dataset.icon, 16)); });
  G.io = new IntersectionObserver(onThumbIO, { root: $('.org-grid', root), rootMargin: '300px' });
  root.addEventListener('click', onClick);
  root.addEventListener('dblclick', e => { const c = e.target.closest('.org-card'); if (c) { const it = G.items.find(i => i.key === c.dataset.key); if (it?.src === 0 && !it.blank) { const target = it.idx; closeGrid(); APP.goTo(target); } } });
  root.addEventListener('keydown', onKey);
  initDrag();
  document.body.classList.add('organizing');
  renderGrid();
  root.focus();
}

function closeGrid() {
  if (!G) return;
  G.io.disconnect(); G.root.remove();
  for (let i = 1; i < sources.length; i++) sources[i].doc?.destroy?.();
  sources = []; G = null;
  document.body.classList.remove('organizing');
}
export const gridOpen = () => !!G;

function renderGrid() {
  const grid = $('.org-grid', G.root);
  G.io.disconnect();
  grid.innerHTML = '';
  G.items.forEach((it, i) => {
    const c = document.createElement('div');
    c.className = 'org-card' + (G.sel.has(it.key) ? ' sel' : '');
    c.dataset.key = it.key; c.draggable = true; c.setAttribute('role', 'option'); c.setAttribute('aria-selected', G.sel.has(it.key));
    const { w, h } = dims(it);
    const rot = it.rot % 360;
    const sw = rot % 180 ? h : w, sh = rot % 180 ? w : h;
    const k = Math.min(150 / sw, 190 / sh);
    c.innerHTML = `<div class="org-thumb" style="width:${Math.round(sw * k)}px;height:${Math.round(sh * k)}px"><div class="org-img" style="width:${Math.round(w * k)}px;height:${Math.round(h * k)}px;transform:translate(-50%,-50%) rotate(${rot}deg)">${it.blank ? '' : '<div class="page-ph"></div>'}</div></div>
      <div class="org-meta"><span class="org-num">${i + 1}</span>${it.src ? `<span class="org-tag" title="${esc(sources[it.src].name)}">new</span>` : ''}${it.blank ? '<span class="org-tag">blank</span>' : ''}</div>
      <div class="org-quick"><button data-q="rotl" title="Rotate left">${icon('rotate-ccw', 14)}</button><button data-q="rotr" title="Rotate right">${icon('rotate-cw', 14)}</button><button data-q="del" title="Delete">${icon('trash-2', 14)}</button></div>`;
    grid.appendChild(c);
    if (!it.blank) { const t = G.thumbs.get(thumbKey(it)); if (t) $('.org-img', c).replaceChildren(t.cloneNode ? cloneCanvas(t) : t); else G.io.observe(c); }
  });
  updateCount();
}
const thumbKey = (it) => it.src + ':' + it.idx;
function dims(it) {
  if (it.blank) return { w: it.blank.w, h: it.blank.h };
  const proxyVp = it._vp; if (proxyVp) return proxyVp;
  if (it.src === 0) { const p = S.pages[it.idx]; const vp = p.proxy.getViewport({ scale: 1, rotation: p.proxy.rotate }); return (it._vp = { w: vp.width, h: vp.height }); }
  return it._vp || { w: 595, h: 842 };
}
function cloneCanvas(c) { const o = document.createElement('canvas'); o.width = c.width; o.height = c.height; o.getContext('2d').drawImage(c, 0, 0); return o; }
function updateCount() {
  const n = G.items.length, s = G.sel.size;
  $('.org-count', G.root).textContent = `${n} page${n === 1 ? '' : 's'}${s ? ` · ${s} selected` : ''}`;
  $$('[data-org="rotl"],[data-org="rotr"],[data-org="dup"],[data-org="del"],[data-org="extract"]', G.root).forEach(b => b.disabled = !s);
}

const tq = []; let tActive = 0;
function onThumbIO(entries) {
  for (const e of entries) if (e.isIntersecting) { G.io.unobserve(e.target); tq.push(e.target); }
  pump();
}
function pump() {
  while (G && tActive < 3 && tq.length) {
    const card = tq.shift(); tActive++;
    drawThumb(card).catch(e => console.warn(e)).finally(() => { tActive--; pump(); });
  }
}
async function drawThumb(card) {
  if (!G) return;
  const it = G.items.find(i => i.key === card.dataset.key); if (!it || it.blank) return;
  const key = thumbKey(it);
  let c = G.thumbs.get(key);
  if (!c) {
    const page = await sources[it.src].doc.getPage(it.idx + 1);
    const vp1 = page.getViewport({ scale: 1, rotation: page.rotate });
    const k = Math.min(300 / vp1.width, 380 / vp1.height);
    const vp = page.getViewport({ scale: k, rotation: page.rotate });
    c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
    await page.render({ canvasContext: c.getContext('2d', { alpha: false }), viewport: vp }).promise;
    if (!G) return;
    G.thumbs.set(key, c);
  }
  const box = $('.org-img', card); if (box) box.replaceChildren(cloneCanvas(c));
}

/* ---------------- selection & commands ---------------- */
function selectedIdx() { return G.items.map((it, i) => G.sel.has(it.key) ? i : -1).filter(i => i >= 0); }
function insertAt() { const s = selectedIdx(); return s.length ? s[s.length - 1] + 1 : G.items.length; }
function changedGrid() { G.dirty = true; renderGrid(); }

async function onClick(e) {
  const q = e.target.closest('[data-q]');
  if (q) {
    const card = q.closest('.org-card'); const it = G.items.find(i => i.key === card.dataset.key);
    if (q.dataset.q === 'rotl') it.rot = (it.rot + 270) % 360;
    if (q.dataset.q === 'rotr') it.rot = (it.rot + 90) % 360;
    if (q.dataset.q === 'del') { G.items = G.items.filter(x => x !== it); G.sel.delete(it.key); }
    return changedGrid();
  }
  const card = e.target.closest('.org-card');
  if (card) {
    const k = card.dataset.key;
    if (e.shiftKey && G.anchor) {
      const a = G.items.findIndex(i => i.key === G.anchor), b = G.items.findIndex(i => i.key === k);
      if (!(e.ctrlKey || e.metaKey)) G.sel.clear();
      for (let i = Math.min(a, b); i <= Math.max(a, b); i++) G.sel.add(G.items[i].key);
    } else if (e.ctrlKey || e.metaKey) { G.sel.has(k) ? G.sel.delete(k) : G.sel.add(k); G.anchor = k; }
    else { G.sel.clear(); G.sel.add(k); G.anchor = k; }
    $$('.org-card', G.root).forEach(c => { c.classList.toggle('sel', G.sel.has(c.dataset.key)); c.setAttribute('aria-selected', G.sel.has(c.dataset.key)); });
    return updateCount();
  }
  if (e.target.closest('.org-grid') && !e.target.closest('.org-card')) { G.sel.clear(); $$('.org-card.sel', G.root).forEach(c => c.classList.remove('sel')); updateCount(); }
  const b = e.target.closest('[data-org]'); if (!b || b.disabled) return;
  const sel = selectedIdx();
  switch (b.dataset.org) {
    case 'rotl': case 'rotr': sel.forEach(i => G.items[i].rot = (G.items[i].rot + (b.dataset.org === 'rotl' ? 270 : 90)) % 360); return changedGrid();
    case 'dup': { const at = insertAt(); const copies = sel.map(i => ({ ...G.items[i], key: newId() })); G.items.splice(at, 0, ...copies); G.sel = new Set(copies.map(c => c.key)); return changedGrid(); }
    case 'del':
      if (sel.length >= G.items.length) return BTS.toast('A PDF needs at least one page');
      G.items = G.items.filter((_, i) => !sel.includes(i)); G.sel.clear(); return changedGrid();
    case 'all': G.items.forEach(i => G.sel.add(i.key)); return renderGrid();
    case 'blank': {
      const ref = G.items[Math.max(0, insertAt() - 1)]; const d = ref ? dims(ref) : { w: 595.28, h: 841.89 };
      const it = { key: newId(), blank: { w: d.w, h: d.h }, rot: 0 }; G.items.splice(insertAt(), 0, it); G.sel = new Set([it.key]); return changedGrid();
    }
    case 'pdf': {
      const f = await pickPdf(); if (!f) return;
      try {
        const src = await addSource(f.data, f.name);
        const n = sources[src].doc.numPages; const at = insertAt();
        const add = []; for (let i = 0; i < n; i++) { const pg = await sources[src].doc.getPage(i + 1); const vp = pg.getViewport({ scale: 1, rotation: pg.rotate }); add.push({ key: newId(), src, idx: i, rot: 0, _vp: { w: vp.width, h: vp.height } }); }
        G.items.splice(at, 0, ...add); G.sel = new Set(add.map(a => a.key)); changedGrid();
        BTS.toast(`Added ${n} page${n === 1 ? '' : 's'} from ${f.name}`);
      } catch (err) { BTS.toast('Couldn’t read that PDF: ' + (err.message || err), { error: true }); }
      return;
    }
    case 'img': {
      const files = await pickImageFiles(); if (!files.length) return;
      try {
        const ref = G.items[Math.max(0, insertAt() - 1)]; const d = ref ? dims(ref) : null;
        const bytes = await buildImagesPdf(files, { size: d && Math.abs(d.w - 612) < 3 ? 'letter' : 'a4', orient: 'auto', margin: 'small', shrink: true });
        const src = await addSource(bytes, files.length === 1 ? files[0].name : files.length + ' pictures');
        const n = sources[src].doc.numPages; const at = insertAt(); const add = [];
        for (let i = 0; i < n; i++) { const pg = await sources[src].doc.getPage(i + 1); const vp = pg.getViewport({ scale: 1 }); add.push({ key: newId(), src, idx: i, rot: 0, _vp: { w: vp.width, h: vp.height } }); }
        G.items.splice(at, 0, ...add); G.sel = new Set(add.map(a => a.key)); changedGrid();
      } catch (err) { BTS.toast(err.message || String(err), { error: true }); }
      return;
    }
    case 'extract': return extractItems(sel.map(i => G.items[i]));
    case 'split': return splitDialog();
    case 'cancel':
      if (G.dirty && !(await BTS.confirm('Discard changes?', 'Close without applying your page changes?', 'Discard'))) return;
      return closeGrid();
    case 'apply': {
      if (!G.dirty) return closeGrid();
      const items = G.items;
      b.disabled = true;
      try { BTS.toast('Applying page changes…', { ms: 1200 }); await applyGridItems(items); closeGrid(); }
      catch (err) { b.disabled = false; BTS.toast(err.message || String(err), { error: true, ms: 4500 }); }
      return;
    }
  }
}
async function applyGridItems(items) {
  const n = items.length;
  await applyItems(items, { message: `Pages updated · ${n} page${n === 1 ? '' : 's'}` });
}

function onKey(e) {
  e.stopPropagation();
  if (isTypingTarget(e.target)) return;
  const mod = e.ctrlKey || e.metaKey;
  if (e.key === 'Escape') { e.preventDefault(); G.root.querySelector('[data-org="cancel"]').click(); }
  else if ((e.key === 'Delete' || e.key === 'Backspace') && G.sel.size) { e.preventDefault(); G.root.querySelector('[data-org="del"]').click(); }
  else if (mod && e.key.toLowerCase() === 'a') { e.preventDefault(); G.root.querySelector('[data-org="all"]').click(); }
  else if (mod && e.key === 'Enter') { e.preventDefault(); G.root.querySelector('[data-org="apply"]').click(); }
}
const isTypingTarget = (t) => t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA');

/* ---------------- drag to reorder ---------------- */
function initDrag() {
  const grid = $('.org-grid', G.root);
  const marker = document.createElement('div'); marker.className = 'org-marker hidden';
  let dragKeys = null, dropAt = -1;
  grid.addEventListener('dragstart', e => {
    const c = e.target.closest('.org-card'); if (!c) return;
    if (!G.sel.has(c.dataset.key)) { G.sel.clear(); G.sel.add(c.dataset.key); $$('.org-card', G.root).forEach(x => x.classList.toggle('sel', G.sel.has(x.dataset.key))); updateCount(); }
    dragKeys = new Set(G.sel);
    e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', 'pages');
    requestAnimationFrame(() => $$('.org-card.sel', G.root).forEach(x => x.classList.add('dragging')));
  });
  grid.addEventListener('dragover', e => {
    if (!dragKeys) return;
    e.preventDefault(); e.dataTransfer.dropEffect = 'move';
    const cards = $$('.org-card', grid);
    let best = null, bd = Infinity, after = false;
    for (const c of cards) {
      const r = c.getBoundingClientRect();
      const cx = r.left + r.width / 2, cy = r.top + r.height / 2;
      const d = Math.hypot((e.clientX - cx) / 2, e.clientY - cy);
      if (d < bd) { bd = d; best = c; after = e.clientX > cx; }
    }
    if (!best) return;
    const i = G.items.findIndex(it => it.key === best.dataset.key);
    dropAt = after ? i + 1 : i;
    const r = best.getBoundingClientRect(), gr = grid.getBoundingClientRect();
    if (!marker.parentNode) grid.appendChild(marker);
    marker.classList.remove('hidden');
    marker.style.left = (after ? r.right + 6 : r.left - 8) - gr.left + grid.scrollLeft + 'px';
    marker.style.top = r.top - gr.top + grid.scrollTop + 'px'; marker.style.height = r.height + 'px';
    // auto-scroll near the edges
    if (e.clientY < gr.top + 40) grid.scrollTop -= 14; else if (e.clientY > gr.bottom - 40) grid.scrollTop += 14;
  });
  grid.addEventListener('drop', e => {
    if (!dragKeys) return;
    e.preventDefault();
    const moving = G.items.filter(it => dragKeys.has(it.key));
    let at = dropAt; for (let i = 0; i < dropAt; i++) if (dragKeys.has(G.items[i].key)) at--;
    const rest = G.items.filter(it => !dragKeys.has(it.key));
    rest.splice(clamp(at, 0, rest.length), 0, ...moving);
    const changed = rest.some((it, i) => it !== G.items[i]);
    G.items = rest; dragKeys = null; marker.remove();
    if (changed) changedGrid(); else renderGrid();
  });
  grid.addEventListener('dragend', () => { dragKeys = null; marker.remove(); $$('.org-card.dragging', G.root).forEach(x => x.classList.remove('dragging')); });
}

/* ======================= extract & split ======================= */
async function fullBytes() { return IO.buildAnnotatedPdf(); }

async function buildSubset(L, srcDoc, idxs, rots) {
  const out = await L.PDFDocument.create();
  const pages = await out.copyPages(srcDoc, idxs);
  pages.forEach((pg, k) => { if (rots?.[k]) pg.setRotation(L.degrees((pg.getRotation().angle + rots[k]) % 360)); out.addPage(pg); });
  return out.save();
}

/** Extract grid items (only pages of this PDF; inserted pages are included too). */
async function extractItems(items) {
  try {
    const L = await loadPdfLib();
    const full = await L.PDFDocument.load(await fullBytes(), { ignoreEncryption: true });
    const out = await L.PDFDocument.create();
    const ext = new Map();
    for (const it of items) {
      let pg;
      if (it.blank) { pg = out.addPage([it.blank.w, it.blank.h]); }
      else {
        const from = it.src === 0 ? full : (ext.get(it.src) || ext.set(it.src, await L.PDFDocument.load(sources[it.src].bytes)).get(it.src));
        [pg] = await out.copyPages(from, [it.idx]); out.addPage(pg);
      }
      if (it.rot) pg.setRotation(L.degrees((pg.getRotation().angle + it.rot) % 360));
    }
    const label = items.length === 1 ? 'page' : `${items.length} pages`;
    await IO.saveBytes(`${BTS.title} - ${label}.pdf`, await out.save(), 'PDF document', 'pdf');
  } catch (e) { BTS.toast(e.message || String(e), { error: true }); }
}

/** Ribbon "Extract pages": ask which pages, save them as one new PDF. */
export async function extractDialog() {
  if (!S.doc) return;
  const n = S.pages.length;
  const v = await BTS.dialog({
    title: 'Extract pages',
    body: `<p class="dlg-p">Save some pages of this PDF as a new file. Your annotations and form answers are included.</p>
      <div class="field"><label for="ex-range">Pages (this PDF has ${n})</label><input id="ex-range" value="${S.cur + 1}" placeholder="e.g. 1-3, 5"></div>`,
    buttons: [{ label: 'Cancel', value: null }, { label: 'Save pages…', primary: true, value: (r) => $('#ex-range', r).value }],
  });
  if (v == null) return;
  try {
    const idx = parseRanges(v, n).flat();
    const L = await loadPdfLib();
    const full = await L.PDFDocument.load(await fullBytes(), { ignoreEncryption: true });
    const bytes = await buildSubset(L, full, idx);
    await IO.saveBytes(`${BTS.title} - ${idx.length === 1 ? 'page ' + (idx[0] + 1) : 'pages ' + v.replace(/\s+/g, '')}.pdf`, bytes, 'PDF document', 'pdf');
  } catch (e) { BTS.toast(e.message || String(e), { error: true, ms: 4500 }); }
}

export async function splitDialog() {
  if (!S.doc) return;
  const n = G ? G.items.length : S.pages.length;
  if (n < 2) return BTS.toast('This PDF has only one page');
  let mode = 'ranges';
  const v = await BTS.dialog({
    title: 'Split PDF', width: 470,
    body: `<p class="dlg-p">Make several smaller PDFs from this one (${n} pages).</p>
      <div class="opt-list">
        <label class="check"><input type="radio" name="sp" value="ranges" checked> By page ranges</label>
        <div class="field sub"><input id="sp-ranges" placeholder="e.g. 1-3, 4-6, 7" value="${n > 2 ? `1-${Math.ceil(n / 2)}, ${Math.ceil(n / 2) + 1}-${n}` : '1, 2'}"><small>Each range becomes one file.</small></div>
        <label class="check"><input type="radio" name="sp" value="every"> Every <input id="sp-n" type="number" min="1" max="${n}" value="${Math.min(2, n)}" class="num"> pages</label>
        <label class="check"><input type="radio" name="sp" value="each"> Every page as its own file</label>
      </div>`,
    onOpen: (r) => {
      r.addEventListener('change', e => { if (e.target.name === 'sp') mode = e.target.value; });
      $('#sp-ranges', r).addEventListener('focus', () => { $('input[value="ranges"]', r).checked = true; mode = 'ranges'; });
      $('#sp-n', r).addEventListener('focus', () => { $('input[value="every"]', r).checked = true; mode = 'every'; });
    },
    buttons: [{ label: 'Cancel', value: null }, { label: 'Split…', primary: true, value: (r) => ({ mode, ranges: $('#sp-ranges', r).value, every: +$('#sp-n', r).value }) }],
  });
  if (!v) return;
  try {
    let groups;
    if (v.mode === 'ranges') groups = parseRanges(v.ranges, n);
    else { const k = v.mode === 'each' ? 1 : clamp(Math.round(v.every) || 1, 1, n); groups = []; for (let i = 0; i < n; i += k) groups.push(Array.from({ length: Math.min(k, n - i) }, (_, j) => i + j)); }
    if (groups.length > 300) throw new Error('That would make too many files');
    const L = await loadPdfLib();
    if (G && G.dirty) throw new Error('Apply your page changes first, then split');
    const src = await L.PDFDocument.load(await fullBytes(), { ignoreEncryption: true });
    const files = [];
    for (const g of groups) files.push({ name: `${BTS.title} - ${rangeLabel(g)}.pdf`, bytes: await buildSubset(L, src, g) });
    await saveMany(files);
  } catch (e) { BTS.toast(e.message || String(e), { error: true, ms: 4500 }); }
}

/** Save several files: desktop asks once for a name/folder, then writes them side by side. */
async function saveMany(files) {
  if (!window.bts?.isDesktop) { for (const f of files) BTS.download(f.name, f.bytes, 'application/pdf'); BTS.toast(`Downloaded ${files.length} files`); return; }
  const first = await bts.saveDialog(`${BTS.title}.pdf`, [{ name: 'PDF document', extensions: ['pdf'] }]);
  if (!first) return;
  const dir = first.replace(/[^/\\]*$/, '');
  const base = first.slice(dir.length).replace(/\.pdf$/i, '');
  let ok = 0, lastPath = '';
  for (const f of files) {
    const name = f.name.replace(BTS.title, base).replace(/[/\\:*?"<>|]+/g, '-');
    const r = await bts.writeFile(dir + name, f.bytes);
    if (r.error) { BTS.toast('Couldn’t save ' + name + ': ' + r.error, { error: true }); break; }
    ok++; lastPath = r.path || dir + name;
  }
  if (ok) {
    BTS.toast(`Saved ${ok} file${ok === 1 ? '' : 's'} in ${dir.replace(/[/\\]$/, '').split(/[/\\]/).pop() || dir}`, { ms: 3500 });
  }
}
