/* BTS PDF — left sidebar: page thumbnails and document outline (bookmarks) */
import { S, $, $$, esc, baseViewport } from './state.js';
import { goTo, goToDest, hooks } from './viewer.js';

const SB = { pane: 'thumbs', open: false, io: null, outlineLoaded: false, token: 0 };
const THUMB_W = 132;

export function sidebarOpen() { return SB.open; }
export function showSidebar(pane, toggle = false) {
  const el = $('#sidebar');
  if (toggle && SB.open && (!pane || pane === SB.pane)) { hideSidebar(); return; }
  if (pane) SB.pane = pane;
  SB.open = true; el.classList.remove('hidden');
  $$('#sidebar .sb-tabs [data-pane]').forEach(b => b.classList.toggle('active', b.dataset.pane === SB.pane));
  $$('#sidebar .sb-pane').forEach(p => p.classList.toggle('active', p.dataset.pane === SB.pane));
  syncButtons();
  if (SB.pane === 'thumbs') buildThumbs(); else buildOutline();
  try { localStorage.setItem('bts-pdf-sidebar', SB.pane); } catch {}
  window.dispatchEvent(new Event('bts-pdf-resize'));
}
export function hideSidebar() {
  SB.open = false; $('#sidebar').classList.add('hidden'); syncButtons();
  try { localStorage.setItem('bts-pdf-sidebar', ''); } catch {}
  window.dispatchEvent(new Event('bts-pdf-resize'));
}
function syncButtons() { $$('#ribbon [data-cmd="sidebar"]').forEach(b => b.classList.toggle('on', SB.open && b.dataset.pane === SB.pane)); }

/* ---------------- thumbnails ---------------- */
let built = false;
function buildThumbs() {
  if (built || !S.doc) return;
  built = true;
  const host = $('#thumbs'); host.innerHTML = '';
  SB.io?.disconnect();
  SB.io = new IntersectionObserver(onIO, { root: host, rootMargin: '400px 0px' });
  const frag = document.createDocumentFragment();
  S.pages.forEach(p => {
    const vp = baseViewport(p.proxy);
    const h = Math.round(THUMB_W * vp.height / vp.width);
    const b = document.createElement('button');
    b.className = 'thumb'; b.dataset.i = p.i; b.title = 'Page ' + (p.i + 1);
    b.innerHTML = `<div class="tbox" style="width:${THUMB_W}px;height:${h}px"></div><span>${p.i + 1}</span>`;
    frag.appendChild(b);
  });
  host.appendChild(frag);
  $$('.thumb', host).forEach(b => SB.io.observe(b));
  markCurrent(S.cur, true);
}
const tq = []; let tActive = 0;
function onIO(entries) {
  for (const e of entries) if (e.isIntersecting && !e.target.dataset.done) { e.target.dataset.done = '1'; tq.push(e.target); }
  pumpThumbs();
}
function pumpThumbs() {
  const token = SB.token;
  while (tActive < 2 && tq.length) {
    const b = tq.shift(); tActive++;
    drawThumb(b, token).catch(e => { if (e?.name !== 'RenderingCancelledException') console.warn(e); }).finally(() => { tActive--; pumpThumbs(); });
  }
}
async function drawThumb(b, token) {
  const p = S.pages[+b.dataset.i]; if (!p || token !== SB.token) return;
  const vp0 = baseViewport(p.proxy);
  const dpr = Math.min(2, window.devicePixelRatio || 1);
  const scale = THUMB_W / vp0.width;
  const vp = p.proxy.getViewport({ scale: scale * dpr, rotation: (p.proxy.rotate + S.rot) % 360 });
  const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
  await p.proxy.render({ canvasContext: c.getContext('2d', { alpha: false }), viewport: vp }).promise;
  if (token !== SB.token) return;
  const box = $('.tbox', b); box.replaceChildren(c);
}
export function rebuildThumbs() {
  SB.token++; built = false; tq.length = 0;
  if (SB.open && SB.pane === 'thumbs') buildThumbs();
}

function markCurrent(i, scroll) {
  const host = $('#thumbs');
  $$('.thumb.cur', host).forEach(b => b.classList.remove('cur'));
  const b = $(`.thumb[data-i="${i}"]`, host); if (!b) return;
  b.classList.add('cur');
  if (scroll && SB.open && SB.pane === 'thumbs') {
    const hr = host.getBoundingClientRect(), br = b.getBoundingClientRect();
    if (br.top < hr.top || br.bottom > hr.bottom) host.scrollTop += br.top - hr.top - hr.height / 2 + br.height / 2;
  }
}
hooks.current.push(i => markCurrent(i, true));

/* ---------------- outline ---------------- */
async function buildOutline() {
  if (SB.outlineLoaded || !S.doc) return;
  SB.outlineLoaded = true;
  const host = $('#outline');
  host.innerHTML = '<div class="sb-empty">Loading…</div>';
  let ol = null;
  try { ol = await S.doc.getOutline(); } catch {}
  if (!ol || !ol.length) {
    host.innerHTML = `<div class="sb-empty">${icon('list-tree', 28)}<b>No bookmarks</b><span>This PDF doesn’t have a table of contents.</span></div>`;
    return;
  }
  host.innerHTML = '';
  host.appendChild(tree(ol, 0));
}
function tree(items, depth) {
  const ul = document.createElement('ul'); ul.className = 'ol-list'; ul.setAttribute('role', depth ? 'group' : 'tree');
  for (const it of items) {
    const li = document.createElement('li'); li.setAttribute('role', 'treeitem');
    const kids = it.items && it.items.length;
    const open = depth < 1;
    li.innerHTML = `<div class="ol-row" style="padding-left:${6 + depth * 14}px">${kids ? `<button class="ol-tw" aria-label="Expand">${icon('chevron-right', 14)}</button>` : '<span class="ol-tw sp"></span>'}<button class="ol-t" title="${esc(it.title)}" style="${it.bold ? 'font-weight:700;' : ''}${it.italic ? 'font-style:italic;' : ''}">${esc(it.title || 'Untitled')}</button></div>`;
    if (kids) { li.appendChild(tree(it.items, depth + 1)); if (open) li.classList.add('open'); li.setAttribute('aria-expanded', open); }
    $('.ol-t', li).addEventListener('click', () => {
      $$('#outline .ol-row.cur').forEach(r => r.classList.remove('cur'));
      $('.ol-row', li).classList.add('cur');
      if (it.dest) goToDest(it.dest);
      else if (it.url) window.bts?.openExternal ? bts.openExternal(it.url) : window.open(it.url);
    });
    if (kids) $('.ol-tw', li).addEventListener('click', () => { li.classList.toggle('open'); li.setAttribute('aria-expanded', li.classList.contains('open')); });
    ul.appendChild(li);
  }
  return ul;
}

export function resetSidebar() {
  SB.token++; built = false; SB.outlineLoaded = false; tq.length = 0;
  SB.io?.disconnect();
  $('#thumbs').innerHTML = ''; $('#outline').innerHTML = '';
}
export function afterLoad() {
  if (SB.open) { if (SB.pane === 'thumbs') buildThumbs(); else buildOutline(); }
}
export function initSidebar() {
  $('#thumbs').addEventListener('click', e => { const b = e.target.closest('.thumb'); if (b) goTo(+b.dataset.i); });
  let pref = null; try { pref = localStorage.getItem('bts-pdf-sidebar'); } catch {}
  return pref;
}
