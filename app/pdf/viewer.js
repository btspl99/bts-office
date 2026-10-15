/* BTS PDF — page layout, lazy rendering, zoom, rotation, navigation */
import { S, pdfjsLib, CSS_UNITS, $, $$, clamp, baseViewport } from './state.js';

export const hooks = {
  pageRendered: [],   // (p) after canvas render — used by find highlights
  textReady: [],      // (p) after text layer built
  layout: [],         // () after pages are (re)laid out
  current: [],        // (i) current page changed
};
const emit = (name, ...a) => hooks[name].forEach(f => { try { f(...a); } catch (e) { console.error(e); } });

const viewer = () => $('#viewer');
const pagesEl = () => $('#pages');
const MAX_CANVAS_PIXELS = 16777216;   // 4096², keeps memory sane at high zoom
const GAP = 14;

let renderIO = null, visIO = null;
const visible = new Set();       // pages intersecting the viewport (for current-page detection)
const queue = new Set();          // pages waiting to render
let active = 0;

/* ---------------- building the page skeletons ---------------- */
export async function buildPages() {
  const host = pagesEl(); host.innerHTML = '';
  S.pages = [];
  renderIO?.disconnect(); visIO?.disconnect(); visible.clear(); queue.clear();
  // Fetch page proxies in order; getPage is cheap (no rendering).
  const proxies = [];
  for (let i = 1; i <= S.doc.numPages; i++) proxies.push(S.doc.getPage(i));
  const all = await Promise.all(proxies);
  all.forEach((proxy, i) => {
    const el = document.createElement('div');
    el.className = 'page'; el.dataset.i = i;
    el.innerHTML = '<div class="page-ph"></div>';
    S.pages.push({ i, proxy, el, vp1: baseViewport(proxy), canvas: null, renderedScale: 0, renderedRot: -1, task: null, textRot: -1, linkRot: -1, near: false });
  });
  renderIO = new IntersectionObserver(onRenderIO, { root: viewer(), rootMargin: '900px 300px' });
  visIO = new IntersectionObserver(onVisIO, { root: viewer(), threshold: [0, .25, .5, .75, 1] });
  layout();
  S.pages.forEach(p => { renderIO.observe(p.el); visIO.observe(p.el); });
}

/** Arrange pages in rows according to the view mode and size them for the current scale. */
export function layout() {
  const host = pagesEl(); if (!S.doc) return;
  host.className = 'mode-' + S.mode;
  const per = S.mode === 'spread' ? 2 : 1;
  const rows = [];
  for (let i = 0; i < S.pages.length; i += per) rows.push(S.pages.slice(i, i + per));
  const frag = document.createDocumentFragment();
  rows.forEach((r, ri) => {
    const row = document.createElement('div'); row.className = 'row'; row.dataset.row = ri;
    r.forEach(p => row.appendChild(p.el));
    frag.appendChild(row);
  });
  host.replaceChildren(frag);
  S.pages.forEach(p => sizePage(p));
  applySingle();
  emit('layout');
}

function sizePage(p) {
  p.vp1 = baseViewport(p.proxy);
  const w = p.vp1.width * S.scale, h = p.vp1.height * S.scale;
  p.el.style.width = w + 'px'; p.el.style.height = h + 'px';
  p.el.style.setProperty('--scale-factor', S.scale);
}

function applySingle() {
  const rowOf = (i) => S.mode === 'spread' ? Math.floor(i / 2) : i;
  const cr = rowOf(S.cur);
  $$('#pages > .row').forEach((r, ri) => r.classList.toggle('off', S.mode === 'single' && ri !== cr));
}

/* ---------------- lazy rendering ---------------- */
function onRenderIO(entries) {
  for (const e of entries) {
    const p = S.pages[+e.target.dataset.i]; if (!p) continue;
    p.near = e.isIntersecting;
    if (e.isIntersecting) schedule(p);
    else release(p);
  }
}
function onVisIO(entries) {
  for (const e of entries) {
    const i = +e.target.dataset.i;
    if (e.isIntersecting) visible.add(i); else visible.delete(i);
  }
  updateCurrent();
}

function needsRender(p) { return p.renderedScale !== S.scale || p.renderedRot !== S.rot || !p.canvas; }

export function schedule(p, front = false) {
  if (!needsRender(p) && p.textRot === S.rot) return;
  if (front) { queue.delete(p); const rest = [...queue]; queue.clear(); queue.add(p); rest.forEach(x => queue.add(x)); }
  else queue.add(p);
  pump();
}
function pump() {
  while (active < 2 && queue.size) {
    // prefer pages that are actually visible, nearest to current page
    let best = null, bd = Infinity;
    for (const p of queue) {
      const d = (visible.has(p.i) ? 0 : 1000) + Math.abs(p.i - S.cur);
      if (d < bd) { bd = d; best = p; }
    }
    queue.delete(best);
    if (!best.near && !visible.has(best.i) && best.i !== S.cur) continue;
    active++;
    renderPage(best).catch(err => { if (!['RenderingCancelledException', 'AbortException'].includes(err?.name)) console.error(err); })
      .finally(() => { active--; pump(); });
  }
}

/** Drop the bitmap of a page far from view to keep memory low. */
function release(p) {
  queue.delete(p);
  if (p.task) { try { p.task.cancel(); } catch {} p.task = null; }
  if (p.canvas && S.pages.length > 6) { p.canvas.width = p.canvas.height = 0; p.canvas.remove(); p.canvas = null; p.renderedScale = 0; p.el.classList.remove('rendered'); }
}

export async function renderPage(p) {
  const scale = S.scale, rot = S.rot;
  if (needsRender(p)) {
    if (p.task) { try { p.task.cancel(); } catch {} }
    const vp = p.proxy.getViewport({ scale, rotation: (p.proxy.rotate + rot) % 360 });
    let dpr = window.devicePixelRatio || 1;
    const px = vp.width * vp.height * dpr * dpr;
    if (px > MAX_CANVAS_PIXELS) dpr = Math.sqrt(MAX_CANVAS_PIXELS / (vp.width * vp.height));
    const c = document.createElement('canvas');
    c.className = 'pg-canvas';
    c.width = Math.floor(vp.width * dpr); c.height = Math.floor(vp.height * dpr);
    const ctx = c.getContext('2d', { alpha: false });
    const task = p.proxy.render({ canvasContext: ctx, viewport: vp, transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : null, annotationMode: S.form.has ? pdfjsLib.AnnotationMode.ENABLE_FORMS : pdfjsLib.AnnotationMode.ENABLE });   // form boxes are drawn by forms.js
    p.task = task;
    await task.promise;
    p.task = null;
    if (scale !== S.scale || rot !== S.rot) { c.width = c.height = 0; schedule(p); return; }
    if (p.canvas) { p.canvas.replaceWith(c); p.canvas.width = p.canvas.height = 0; } else p.el.prepend(c);
    p.canvas = c; p.renderedScale = scale; p.renderedRot = rot;
    p.el.classList.add('rendered');
    $('.page-ph', p.el)?.remove();
    emit('pageRendered', p);
  }
  if (p.textRot !== S.rot) await buildTextLayer(p);
  else if (p.text && p.textScale !== S.scale) {
    try { p.text.layer.update({ viewport: p.proxy.getViewport({ scale: S.scale, rotation: (p.proxy.rotate + S.rot) % 360 }) }); } catch {}
  }
  if (p.text) p.textScale = S.scale;
  if (p.linkRot !== S.rot) buildLinks(p);
}

/* ---------------- text layer ---------------- */
export async function getText(i) {
  if (!S.textCache.has(i)) S.textCache.set(i, S.pages[i].proxy.getTextContent());
  return S.textCache.get(i);
}

async function buildTextLayer(p) {
  const rot = S.rot;
  const token = p.textToken = (p.textToken || 0) + 1;
  const tc = await getText(p.i);
  if (rot !== S.rot || token !== p.textToken) return;
  // Note: TextLayer.cancel() in pdf.js 4.10 can throw inside its stream pump, so stale builds are
  // simply discarded via the token instead of being cancelled.
  const div = document.createElement('div'); div.className = 'textLayer';
  const vp = p.proxy.getViewport({ scale: S.scale, rotation: (p.proxy.rotate + rot) % 360 });
  const layer = new pdfjsLib.TextLayer({ textContentSource: tc, container: div, viewport: vp });
  try { await layer.render(); } catch (e) { if (e?.name !== 'AbortException') console.warn(e); return; }
  if (rot !== S.rot || token !== p.textToken) return;
  $('.textLayer', p.el)?.remove();
  const ann = $('.annLayer', p.el);
  if (ann) ann.after(div); else p.el.appendChild(div);
  const ll = $('.linkLayer', p.el); if (ll) p.el.appendChild(ll);
  const end = document.createElement('div'); end.className = 'endOfContent'; div.appendChild(end);
  div.addEventListener('mousedown', () => div.classList.add('selecting'));
  p.text = { layer, div };
  p.textRot = rot;
  emit('textReady', p);
}
document.addEventListener('mouseup', () => $$('.textLayer.selecting').forEach(d => d.classList.remove('selecting')));

/* ---------------- link layer (internal + external links) ---------------- */
async function buildLinks(p) {
  p.linkRot = S.rot;
  let annots;
  try { annots = p.annotsCache ||= await p.proxy.getAnnotations({ intent: 'display' }); } catch { return; }
  $('.linkLayer', p.el)?.remove();
  const links = annots.filter(a => a.subtype === 'Link' && a.rect);
  if (!links.length) return;
  const layer = document.createElement('div'); layer.className = 'linkLayer';
  for (const a of links) {
    const [x1, y1, x2, y2] = a.rect;
    const pa = p.vp1.convertToViewportPoint(x1, y1), pb = p.vp1.convertToViewportPoint(x2, y2);
    const l = Math.min(pa[0], pb[0]), t = Math.min(pa[1], pb[1]), w = Math.abs(pa[0] - pb[0]), h = Math.abs(pa[1] - pb[1]);
    const el = document.createElement('a');
    el.style.cssText = `left:${l / p.vp1.width * 100}%;top:${t / p.vp1.height * 100}%;width:${w / p.vp1.width * 100}%;height:${h / p.vp1.height * 100}%`;
    if (a.url || a.unsafeUrl) {
      const url = a.url || a.unsafeUrl; el.href = '#'; el.title = url;
      el.addEventListener('click', e => { e.preventDefault(); if (window.bts?.openExternal) bts.openExternal(url); else window.open(url, '_blank', 'noopener'); });
    } else if (a.dest) {
      el.href = '#'; el.title = 'Go to destination';
      el.addEventListener('click', e => { e.preventDefault(); goToDest(a.dest); });
    } else continue;
    layer.appendChild(el);
  }
  p.el.appendChild(layer);
}

/** Navigate to a PDF destination (named string or explicit array). */
export async function goToDest(dest) {
  try {
    let d = dest;
    if (typeof d === 'string') d = await S.doc.getDestination(d);
    if (!Array.isArray(d)) return;
    const ref = d[0];
    const idx = typeof ref === 'object' && ref ? await S.doc.getPageIndex(ref) : (Number.isInteger(ref) ? ref : 0);
    const kind = d[1]?.name;
    let top = null, left = null;
    if (kind === 'XYZ') { left = d[2]; top = d[3]; }
    else if (kind === 'FitH' || kind === 'FitBH') top = d[2];
    else if (kind === 'FitR') { left = d[2]; top = d[5]; }
    goTo(idx, { top, left });
  } catch (e) { console.warn('Bad destination', e); }
}

/* ---------------- current page ---------------- */
let curRaf = 0;
export function updateCurrent() {
  if (curRaf) return;
  curRaf = requestAnimationFrame(() => {
    curRaf = 0;
    if (!S.doc || S.mode === 'single') return;
    const v = viewer().getBoundingClientRect();
    const probe = v.top + Math.min(v.height * 0.35, 260);
    let best = S.cur, bestScore = -Infinity;
    for (const i of visible) {
      const r = S.pages[i].el.getBoundingClientRect();
      const visH = Math.min(r.bottom, v.bottom) - Math.max(r.top, v.top);
      const contains = r.top <= probe && r.bottom >= probe;
      const score = (contains ? 1e6 : 0) + visH - i * 0.001;
      if (score > bestScore) { bestScore = score; best = i; }
    }
    setCurrent(best);
  });
}
export function setCurrent(i) {
  i = clamp(i, 0, S.pages.length - 1);
  if (i === S.cur && S._curInit) return;
  S._curInit = true; S.cur = i;
  emit('current', i);
}

/* ---------------- navigation ---------------- */
export function goTo(i, { top = null, left = null, smooth = false } = {}) {
  if (!S.doc) return;
  i = clamp(i, 0, S.pages.length - 1);
  const p = S.pages[i];
  if (S.mode === 'single') { S.cur = i; applySingle(); emit('current', i); }
  const v = viewer();
  const vr = v.getBoundingClientRect(), pr = p.el.getBoundingClientRect();
  let y = v.scrollTop + (pr.top - vr.top) - GAP;
  let x = v.scrollLeft;
  if (top != null || left != null) {
    const [vx, vy] = p.vp1.convertToViewportPoint(left ?? 0, top ?? 0);
    if (top != null) y = v.scrollTop + (pr.top - vr.top) + vy * S.scale - 8;
    if (left != null && v.scrollWidth > v.clientWidth) x = v.scrollLeft + (pr.left - vr.left) + vx * S.scale - 8;
  }
  v.scrollTo({ top: y, left: x, behavior: smooth ? 'smooth' : 'auto' });
  if (S.mode !== 'single') setCurrent(i);
  schedule(p, true);
}
export const next = () => goTo(S.cur + (S.mode === 'spread' ? 2 - (S.cur % 2) : 1));
export const prev = () => goTo(S.mode === 'spread' ? S.cur - 2 - (S.cur % 2 ? -1 : 0) : S.cur - 1);

/* ---------------- zoom / fit / rotate ---------------- */
/** Reflect S.scale in the shell's status-bar zoom controls without re-entering cfg.onZoom. */
function pushZoom() {
  const z = Math.round(S.scale / CSS_UNITS * 100); BTS.zoom = z;
  const r = $('#sb-zoom'); if (r) r.value = z; const v = $('#sb-zoom-val'); if (v) v.textContent = z + '%';
}

/** Change scale keeping the point under (cx, cy) — or the viewport centre — fixed. */
export function setScale(scale, { cx, cy, fit = null } = {}) {
  if (!S.doc) return;
  scale = clamp(scale, 0.1 * CSS_UNITS, 5 * CSS_UNITS);
  S.fit = fit;
  updateFitButtons();
  if (Math.abs(scale - S.scale) < 1e-4) { pushZoom(); return; }
  const v = viewer(); const vr = v.getBoundingClientRect();
  cx ??= vr.left + vr.width / 2; cy ??= vr.top + Math.min(vr.height / 2, 200);
  const p = S.pages[S.cur]; const r0 = p.el.getBoundingClientRect();
  const ok = r0.width > 0 && r0.height > 0;
  const fx = ok ? (cx - r0.left) / r0.width : 0, fy = ok ? (cy - r0.top) / r0.height : 0;
  S.scale = scale;
  S.pages.forEach(sizePage);
  if (ok) {
    const r1 = p.el.getBoundingClientRect();
    v.scrollLeft += (r1.left + fx * r1.width) - cx;
    v.scrollTop += (r1.top + fy * r1.height) - cy;
  }
  pushZoom();
  S.pages.forEach(q => { if (q.near || visible.has(q.i)) schedule(q); });
  emit('layout');
}

export function fitScale(kind) {
  const v = viewer(); if (!S.doc) return S.scale;
  const cs = getComputedStyle(pagesEl());
  const padX = parseFloat(cs.paddingLeft) + parseFloat(cs.paddingRight) + 4;
  const padY = parseFloat(cs.paddingTop) + parseFloat(cs.paddingBottom);
  const scrollbar = 14;
  const availW = v.clientWidth - padX - (kind === 'page' ? 0 : scrollbar);
  const availH = v.clientHeight - padY;
  const per = S.mode === 'spread' ? 2 : 1;
  let maxW = 0, curH = 0;
  for (let i = 0; i < S.pages.length; i += per) {
    let w = 0; for (let k = 0; k < per && i + k < S.pages.length; k++) w += baseViewport(S.pages[i + k].proxy).width;
    maxW = Math.max(maxW, w);
    if (S.pages.length > 400 && i > 60) break; // huge docs: sample the first pages
  }
  const gaps = (per - 1) * GAP;
  const cur = baseViewport(S.pages[S.cur].proxy);
  curH = cur.height;
  const sw = (availW - gaps) / maxW;
  if (kind === 'width') return sw;
  if (kind === 'page') { const curW = per === 2 ? maxW : cur.width; return Math.min((availW - gaps) / curW, availH / curH); }
  return Math.min(sw, 1.25 * CSS_UNITS);  // auto: fit width, but no larger than 125%
}

export function applyFit(kind) {
  if (!S.doc) return;
  const s = fitScale(kind);
  setScale(s, { fit: kind });
  if (kind === 'page') goTo(S.cur);
}

function updateFitButtons() {
  $$('[data-fit]').forEach(b => b.classList.toggle('on', b.dataset.fit === S.fit));
}

export function rotate(delta) {
  if (!S.doc) return;
  const keep = S.cur;
  S.rot = (S.rot + delta + 360) % 360;
  S.pages.forEach(p => { sizePage(p); p.textRot = -1; p.linkRot = -1; });
  emit('layout');
  if (S.fit) setScale(fitScale(S.fit), { fit: S.fit });
  goTo(keep);
  S.pages.forEach(q => { if (q.near || visible.has(q.i)) schedule(q); });
}

export function setMode(mode) {
  if (!S.doc) { S.mode = mode; return; }
  const keep = S.cur;
  S.mode = mode;
  layout();
  if (S.fit) setScale(fitScale(S.fit), { fit: S.fit });
  requestAnimationFrame(() => { goTo(keep); S.pages.forEach(q => { if (q.near || visible.has(q.i)) schedule(q); }); });
}

/** Re-render everything (after theme / dark-read changes nothing needed; after doc changes yes). */
export function invalidateAll() {
  S.pages.forEach(p => { p.renderedScale = 0; p.textRot = -1; p.linkRot = -1; p.formRot = -1; p.annotsCache = null; });
  S.pages.forEach(q => { if (q.near || visible.has(q.i)) schedule(q); });
}

export function teardown() {
  renderIO?.disconnect(); visIO?.disconnect(); visible.clear(); queue.clear();
  S.pages.forEach(p => { try { p.task?.cancel(); } catch {} if (p.canvas) p.canvas.width = p.canvas.height = 0; });
  S.pages = [];
  pagesEl().innerHTML = '';
}

/* ---------------- single-page mode: wheel flips pages at the edges ---------------- */
let flipLock = 0;
export function singleWheel(e) {
  if (S.mode !== 'single' || e.ctrlKey) return;
  const v = viewer(); const now = Date.now();
  if (now < flipLock) return;
  const atBottom = v.scrollTop + v.clientHeight >= v.scrollHeight - 2;
  const atTop = v.scrollTop <= 1;
  if (e.deltaY > 0 && atBottom && S.cur < S.pages.length - 1) { flipLock = now + 350; e.preventDefault(); goTo(S.cur + 1); v.scrollTop = 0; }
  else if (e.deltaY < 0 && atTop && S.cur > 0) { flipLock = now + 350; e.preventDefault(); goTo(S.cur - 1); v.scrollTop = v.scrollHeight; }
}
