/* BTS PDF — annotations: data model, SVG overlays, drawing tools, selection, undo/redo.
 * Every annotation is stored in PDF user-space coordinates (points, y up) of its page, so it is
 * independent of zoom and view rotation and can be written straight into the PDF with pdf-lib.
 *   highlight|underline|strike: { rects:[[x1,y1,x2,y2]], color }
 *   pen:     { pts:[[x,y]...], color, w }
 *   rect|ellipse: { r:[x1,y1,x2,y2], color, w }
 *   arrow:   { a:[x,y], b:[x,y], color, w }
 *   note:    { x, y, text, color }            (x,y = top-left of the note icon)
 *   textbox: { x, y, text, color, size }      (x,y = top-left of the first line)
 * "Display frame" = the page as it normally appears (its own /Rotate applied, no view rotation).
 */
import { S, $, $$, esc, clamp, newId, clientToPdf, MARKUP, NOTE, wrapText, measureHelv, isTyping } from './state.js';

const FONT = 'Helvetica, Arial, "Liberation Sans", sans-serif';
export const TB = { lh: 1.25, base: 0.9 };   // textbox line height & first-baseline factors

/* ---------------- geometry helpers ---------------- */
export const nativeVp = (p) => (p.vpN ||= p.proxy.getViewport({ scale: 1, rotation: p.proxy.rotate }));
const V = (vp, x, y) => vp.convertToViewportPoint(x, y);
const pts = (vp, list) => list.map(([x, y]) => V(vp, x, y).map(n => +n.toFixed(2)).join(',')).join(' ');
const corners = ([x1, y1, x2, y2]) => [[x1, y1], [x2, y1], [x2, y2], [x1, y2]];
const normRect = ([a, b, c, d]) => [Math.min(a, c), Math.min(b, d), Math.max(a, c), Math.max(a, c) === a ? Math.max(b, d) : Math.max(b, d)];
/** rect (PDF) -> bbox in a viewport frame */
function vpBox(vp, r) { const c = corners(r).map(([x, y]) => V(vp, x, y)); const xs = c.map(q => q[0]), ys = c.map(q => q[1]); return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]; }
/** rotation of the view relative to the display frame (0/90/180/270) */
const deltaRot = (p, vp) => ((vp.rotation - p.proxy.rotate) % 360 + 360) % 360;

/** Lines for a markup rect, computed in the display frame so underline is under the text even on rotated pages. */
export function markupLine(p, a, r) {
  const vpN = nativeVp(p); const [l, t, rr, b] = vpBox(vpN, r);
  const h = b - t;
  const y = a.type === 'underline' ? b - Math.max(0.6, h * 0.06) : t + h * 0.55;
  return { from: vpN.convertToPdfPoint(l, y), to: vpN.convertToPdfPoint(rr, y), w: Math.max(0.8, h * 0.075) };
}

export function textLayout(a, measure) {
  const size = a.size || 14; const m = measure || measureHelv(size);
  const lines = String(a.text || '').split('\n');
  const w = Math.max(size * 0.6, ...lines.map(l => m(l)));
  return { size, lines, w, h: lines.length * size * TB.lh, lh: size * TB.lh, base: size * TB.base };
}
export function noteLayout(a, p) {
  const m = measureHelv(NOTE.fs);
  const lines = a.text ? wrapText(a.text, NOTE.maxW, m) : [];
  const w = lines.length ? Math.max(...lines.map(l => m(l))) + NOTE.pad * 2 : 0;
  const h = lines.length * NOTE.lh + NOTE.pad * 2;
  const vpN = nativeVp(p); const [dx] = V(vpN, a.x, a.y);
  const right = dx + NOTE.icon + 4 + w <= vpN.width;
  return { lines, w, h, ox: right ? NOTE.icon + 4 : -(w + 4) };
}

/** Bounding box of an annotation in the display frame [l,t,r,b] (points). */
function dispBox(p, a) {
  const vpN = nativeVp(p);
  switch (a.type) {
    case 'highlight': case 'underline': case 'strike': {
      const bs = a.rects.map(r => vpBox(vpN, r));
      return [Math.min(...bs.map(b => b[0])), Math.min(...bs.map(b => b[1])), Math.max(...bs.map(b => b[2])), Math.max(...bs.map(b => b[3]))];
    }
    case 'pen': { const q = a.pts.map(([x, y]) => V(vpN, x, y)); const hw = a.w / 2; return [Math.min(...q.map(v => v[0])) - hw, Math.min(...q.map(v => v[1])) - hw, Math.max(...q.map(v => v[0])) + hw, Math.max(...q.map(v => v[1])) + hw]; }
    case 'rect': case 'ellipse': { const b = vpBox(vpN, a.r); const hw = a.w / 2; return [b[0] - hw, b[1] - hw, b[2] + hw, b[3] + hw]; }
    case 'arrow': { const A = V(vpN, ...a.a), B = V(vpN, ...a.b); const hs = headSize(a); return [Math.min(A[0], B[0]) - hs, Math.min(A[1], B[1]) - hs, Math.max(A[0], B[0]) + hs, Math.max(A[1], B[1]) + hs]; }
    case 'note': { const [x, y] = V(vpN, a.x, a.y); return [x, y, x + NOTE.icon, y + NOTE.icon]; }
    case 'textbox': { const [x, y] = V(vpN, a.x, a.y); const L = textLayout(a); return [x - 2, y - 2, x + L.w + 2, y + L.h + 2]; }
  }
  return [0, 0, 0, 0];
}
const headSize = (a) => Math.max(7, a.w * 3.2);

function distSeg(px, py, ax, ay, bx, by) {
  const dx = bx - ax, dy = by - ay; const l2 = dx * dx + dy * dy;
  let t = l2 ? ((px - ax) * dx + (py - ay) * dy) / l2 : 0; t = clamp(t, 0, 1);
  return Math.hypot(px - (ax + t * dx), py - (ay + t * dy));
}
/** Is display-frame point (x,y) on annotation a? tol in points. */
function hit(p, a, x, y, tol) {
  const vpN = nativeVp(p);
  const inBox = (b, t = tol) => x >= b[0] - t && x <= b[2] + t && y >= b[1] - t && y <= b[3] + t;
  switch (a.type) {
    case 'highlight': case 'underline': case 'strike': return a.rects.some(r => inBox(vpBox(vpN, r), 1));
    case 'pen': { const q = a.pts.map(([px, py]) => V(vpN, px, py)); if (q.length === 1) return Math.hypot(x - q[0][0], y - q[0][1]) < a.w / 2 + tol; for (let i = 1; i < q.length; i++) if (distSeg(x, y, ...q[i - 1], ...q[i]) < a.w / 2 + tol) return true; return false; }
    case 'rect': { const b = vpBox(vpN, a.r); if (!inBox(b)) return false; return Math.min(Math.abs(x - b[0]), Math.abs(x - b[2]), Math.abs(y - b[1]), Math.abs(y - b[3])) < a.w / 2 + tol; }
    case 'ellipse': { const b = vpBox(vpN, a.r); const cx = (b[0] + b[2]) / 2, cy = (b[1] + b[3]) / 2, rx = Math.max(1, (b[2] - b[0]) / 2), ry = Math.max(1, (b[3] - b[1]) / 2); const d = Math.hypot((x - cx) / rx, (y - cy) / ry); return Math.abs(d - 1) * Math.min(rx, ry) < a.w / 2 + tol; }
    case 'arrow': { const A = V(vpN, ...a.a), B = V(vpN, ...a.b); return distSeg(x, y, ...A, ...B) < a.w / 2 + tol + 2; }
    case 'note': case 'textbox': return inBox(dispBox(p, a), 1);
  }
  return false;
}
export function hitTest(p, cx, cy) {
  const [px, py] = clientToPdf(p, cx, cy);
  const [x, y] = V(nativeVp(p), px, py);
  const tol = 5 / S.scale;
  const list = S.anns.filter(a => a.p === p.i);
  // shapes before markup (markup sits underneath), last drawn first
  for (const a of [...list].reverse()) if (!MARKUP.has(a.type) && hit(p, a, x, y, tol)) return a;
  for (const a of [...list].reverse()) if (MARKUP.has(a.type) && hit(p, a, x, y, tol)) return a;
  return null;
}

/* ---------------- SVG building (shared by screen + print) ---------------- */
/** Build SVG inner markup for page p in viewport vp (scale 1 = points).
 *  part: 'hl' (highlights, drawn with multiply) or 'ann' (everything else).
 *  opts.print: notes show their text callouts; no selection box. */
export function svgContent(p, vp, part, opts = {}) {
  const out = [];
  const dr = deltaRot(p, vp);
  for (const a of S.anns) {
    if (a.p !== p.i) continue;
    const isHl = a.type === 'highlight';
    if ((part === 'hl') !== isHl) continue;
    const id = `data-id="${a.id}"`;
    const c = esc(a.color);
    switch (a.type) {
      case 'highlight':
        out.push(`<g ${id} fill="${c}" fill-opacity="0.42">${a.rects.map(r => `<polygon points="${pts(vp, corners(r))}"/>`).join('')}</g>`); break;
      case 'underline': case 'strike':
        out.push(`<g ${id} stroke="${c}" stroke-linecap="butt">${a.rects.map(r => { const L = markupLine(p, a, r); const A = V(vp, ...L.from), B = V(vp, ...L.to); return `<line x1="${A[0]}" y1="${A[1]}" x2="${B[0]}" y2="${B[1]}" stroke-width="${L.w}"/>`; }).join('')}</g>`); break;
      case 'pen': {
        const d = a.pts.map(([x, y], i) => { const v = V(vp, x, y); return (i ? 'L' : 'M') + v[0].toFixed(2) + ' ' + v[1].toFixed(2); }).join('') + (a.pts.length === 1 ? 'l0.01 0' : '');
        out.push(`<path ${id} d="${d}" fill="none" stroke="${c}" stroke-width="${a.w}" stroke-linecap="round" stroke-linejoin="round"/>`); break;
      }
      case 'rect': out.push(`<polygon ${id} points="${pts(vp, corners(a.r))}" fill="none" stroke="${c}" stroke-width="${a.w}" stroke-linejoin="miter"/>`); break;
      case 'ellipse': { const b = vpBox(vp, a.r); out.push(`<ellipse ${id} cx="${(b[0] + b[2]) / 2}" cy="${(b[1] + b[3]) / 2}" rx="${(b[2] - b[0]) / 2}" ry="${(b[3] - b[1]) / 2}" fill="none" stroke="${c}" stroke-width="${a.w}"/>`); break; }
      case 'arrow': {
        const A = V(vp, ...a.a), B = V(vp, ...a.b); const g = arrowGeom(A, B, a);
        out.push(`<g ${id}><line x1="${A[0]}" y1="${A[1]}" x2="${g.base[0]}" y2="${g.base[1]}" stroke="${c}" stroke-width="${a.w}" stroke-linecap="round"/><polygon points="${g.head.map(q => q.join(',')).join(' ')}" fill="${c}"/></g>`); break;
      }
      case 'textbox': {
        const L = textLayout(a); const [x, y] = V(vp, a.x, a.y);
        out.push(`<g ${id} transform="translate(${x} ${y}) rotate(${dr})"><rect class="tb-frame" x="-2" y="-2" width="${L.w + 4}" height="${L.h + 4}" fill="transparent"/><text font-family='${FONT}' font-size="${L.size}" fill="${c}" xml:space="preserve">${L.lines.map((l, i) => `<tspan x="0" y="${L.base + i * L.lh}">${esc(l) || ' '}</tspan>`).join('')}</text></g>`); break;
      }
      case 'note': {
        const [x, y] = V(vp, a.x, a.y); const s = NOTE.icon;
        let call = '';
        if (opts.print && a.text) {
          const N = noteLayout(a, p);
          call = `<g transform="translate(${N.ox} 0)"><rect width="${N.w}" height="${N.h}" rx="2" fill="#fffef5" stroke="${c}" stroke-width="0.8"/><text font-family='${FONT}' font-size="${NOTE.fs}" fill="#1b1b1b" xml:space="preserve">${N.lines.map((l, i) => `<tspan x="${NOTE.pad}" y="${NOTE.pad + NOTE.fs * 0.85 + i * NOTE.lh}">${esc(l) || ' '}</tspan>`).join('')}</text></g>`;
        }
        out.push(`<g ${id} class="note" transform="translate(${x} ${y}) rotate(${dr})">${call}<rect width="${s}" height="${s}" rx="3" fill="${c}" stroke="rgba(0,0,0,.35)" stroke-width="0.8"/><path d="M4.5 6h9M4.5 9h9M4.5 12h6" stroke="rgba(0,0,0,.55)" stroke-width="1.2" stroke-linecap="round"/></g>`); break;
      }
    }
  }
  if (part === 'ann' && !opts.print && S.sel) {
    const a = S.anns.find(x => x.id === S.sel && x.p === p.i);
    if (a) {
      const b = dispBox(p, a); const vpN = nativeVp(p);
      const P = [[b[0], b[1]], [b[2], b[1]], [b[2], b[3]], [b[0], b[3]]].map(([x, y]) => V(vp, ...vpN.convertToPdfPoint(x, y)));
      const pad = 3 / S.scale;
      const xs = P.map(q => q[0]), ys = P.map(q => q[1]);
      out.push(`<rect class="selbox" x="${Math.min(...xs) - pad}" y="${Math.min(...ys) - pad}" width="${Math.max(...xs) - Math.min(...xs) + 2 * pad}" height="${Math.max(...ys) - Math.min(...ys) + 2 * pad}" fill="none" stroke-width="${1.25 / S.scale}" stroke-dasharray="${4 / S.scale} ${3 / S.scale}"/>`);
    }
  }
  return out.join('');
}
export function arrowGeom(A, B, a) {
  const hs = headSize(a); const ang = Math.atan2(B[1] - A[1], B[0] - A[0]);
  const len = Math.hypot(B[0] - A[0], B[1] - A[1]); const h = Math.min(hs, len * 0.6);
  const base = [B[0] - Math.cos(ang) * h * 0.8, B[1] - Math.sin(ang) * h * 0.8];
  const l = [B[0] - Math.cos(ang - 0.42) * h, B[1] - Math.sin(ang - 0.42) * h];
  const r = [B[0] - Math.cos(ang + 0.42) * h, B[1] - Math.sin(ang + 0.42) * h];
  return { base, head: [B, l, r].map(q => q.map(n => +n.toFixed(2))) };
}

/* ---------------- on-screen layers ---------------- */
function ensureLayers(p) {
  let hl = $('.hlLayer', p.el), an = $('.annLayer', p.el);
  if (!hl) {
    p.el.insertAdjacentHTML('beforeend', '<svg class="hlLayer" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="none"></svg><svg class="annLayer" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="none"></svg>');
    hl = $('.hlLayer', p.el); an = $('.annLayer', p.el);
    const tl = $('.textLayer', p.el); if (tl) p.el.appendChild(tl);   // keep text layer above overlays
    const ll = $('.linkLayer', p.el); if (ll) p.el.appendChild(ll);
  }
  return [hl, an];
}
export function renderPage(i) {
  const p = S.pages[i]; if (!p) return;
  const has = S.anns.some(a => a.p === i) || $('.annLayer', p.el);
  if (!has) return;
  const [hl, an] = ensureLayers(p);
  const vb = `0 0 ${p.vp1.width} ${p.vp1.height}`;
  hl.setAttribute('viewBox', vb); an.setAttribute('viewBox', vb);
  hl.innerHTML = svgContent(p, p.vp1, 'hl');
  an.innerHTML = svgContent(p, p.vp1, 'ann') + (draft && draft.p === i ? svgContentFor(p, draft) : '');
}
function svgContentFor(p, a) { const keep = S.anns; S.anns = [a]; try { return svgContent(p, p.vp1, a.type === 'highlight' ? 'hl' : 'ann'); } finally { S.anns = keep; } }
export function renderAll() { S.pages.forEach(p => renderPage(p.i)); }

/* ---------------- history ---------------- */
const snap = () => JSON.stringify(S.anns);
export function pushUndo() { S.undo.push(snap()); if (S.undo.length > 200) S.undo.shift(); S.redo = []; }
function restore(json) {
  const before = new Set(S.anns.map(a => a.p));
  S.anns = JSON.parse(json); S.sel = null;
  S.anns.forEach(a => before.add(a.p));
  before.forEach(i => renderPage(i));
  changed(false);
}
export function undo() { if (!S.undo.length) return BTS.toast('Nothing to undo'); S.redo.push(snap()); restore(S.undo.pop()); }
export function redo() { if (!S.redo.length) return BTS.toast('Nothing to redo'); S.undo.push(snap()); restore(S.redo.pop()); }

export const listeners = { change: [] };
function changed(mark = true) {
  if (mark) BTS.markDirty();
  else if (!S.anns.length && !S.modified && S.undo.length === 0) BTS.markClean();
  else BTS.markDirty();
  listeners.change.forEach(f => f());
}

export function addAnn(a) { pushUndo(); a.id ||= newId(); S.anns.push(a); renderPage(a.p); changed(); return a; }
export function deleteAnn(id, { record = true } = {}) {
  const a = S.anns.find(x => x.id === id); if (!a) return;
  if (record) pushUndo();
  S.anns = S.anns.filter(x => x.id !== id);
  if (S.sel === id) S.sel = null;
  closeNotePopup();
  renderPage(a.p); changed();
}
export function deleteSelected() {
  if (!S.sel) { BTS.toast('Select an annotation first (Select tool, then click it)'); return; }
  deleteAnn(S.sel);
}
export function clearPage(i) {
  if (!S.anns.some(a => a.p === i)) return BTS.toast('No annotations on this page');
  pushUndo(); S.anns = S.anns.filter(a => a.p !== i); S.sel = null; renderPage(i); changed();
}
export function select(id) {
  const old = S.anns.find(a => a.id === S.sel);
  S.sel = id;
  if (old) renderPage(old.p);
  const a = S.anns.find(x => x.id === id); if (a) renderPage(a.p);
  listeners.change.forEach(f => f());
}
export function setColorOfSelected(color) {
  const a = S.anns.find(x => x.id === S.sel); if (!a) return false;
  pushUndo(); a.color = color; renderPage(a.p); changed(); return true;
}
export function setWidthOfSelected(w) {
  const a = S.anns.find(x => x.id === S.sel); if (!a || !('w' in a)) return false;
  pushUndo(); a.w = w; renderPage(a.p); changed(); return true;
}
function translate(a, dx, dy) {
  const mv = ([x, y]) => [x + dx, y + dy];
  if (a.rects) a.rects = a.rects.map(([x1, y1, x2, y2]) => [x1 + dx, y1 + dy, x2 + dx, y2 + dy]);
  if (a.pts) a.pts = a.pts.map(mv);
  if (a.r) a.r = [a.r[0] + dx, a.r[1] + dy, a.r[2] + dx, a.r[3] + dy];
  if (a.a) { a.a = mv(a.a); a.b = mv(a.b); }
  if ('x' in a) { a.x += dx; a.y += dy; }
}

/* ---------------- text markup from selection ---------------- */
/** Collect selected text rectangles per page, in PDF coordinates. */
export function selectionRects() {
  const sel = getSelection(); if (!sel || sel.isCollapsed || !sel.rangeCount) return null;
  const range = sel.getRangeAt(0);
  const byPage = new Map(); let text = '';
  for (const p of S.pages) {
    const tl = $('.textLayer', p.el); if (!tl || !range.intersectsNode(tl)) continue;
    const walker = document.createTreeWalker(tl, NodeFilter.SHOW_TEXT);
    let n; const rects = [];
    while ((n = walker.nextNode())) {
      if (!range.intersectsNode(n) || !n.nodeValue.trim()) continue;
      const r = document.createRange();
      r.selectNodeContents(n);
      if (n === range.startContainer) r.setStart(n, range.startOffset);
      if (n === range.endContainer) r.setEnd(n, range.endOffset);
      if (r.collapsed) continue;
      text += r.toString();
      for (const cr of r.getClientRects()) {
        if (cr.width < 0.5 || cr.height < 0.5) continue;
        const c = [[cr.left, cr.top], [cr.right, cr.top], [cr.right, cr.bottom], [cr.left, cr.bottom]].map(([x, y]) => clientToPdf(p, x, y));
        const xs = c.map(q => q[0]), ys = c.map(q => q[1]);
        rects.push([Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]);
      }
    }
    if (rects.length) byPage.set(p.i, mergeRects(p, rects));
  }
  return byPage.size ? { byPage, text } : null;
}
/** Merge rectangles that sit on the same line (in the display frame). */
function mergeRects(p, rects) {
  const vpN = nativeVp(p);
  const items = rects.map(r => vpBox(vpN, r)).sort((a, b) => a[1] - b[1] || a[0] - b[0]);
  const lines = [];
  for (const b of items) {
    const h = b[3] - b[1];
    const line = lines.find(L => { const ov = Math.min(L[3], b[3]) - Math.max(L[1], b[1]); return ov > 0.5 * Math.min(h, L[3] - L[1]) && b[0] <= L[2] + Math.max(h, L[3] - L[1]) * 0.8 && b[2] >= L[0] - h * 0.8; });
    if (line) { line[0] = Math.min(line[0], b[0]); line[1] = Math.min(line[1], b[1]); line[2] = Math.max(line[2], b[2]); line[3] = Math.max(line[3], b[3]); }
    else lines.push([...b]);
  }
  return lines.map(([l, t, r, b]) => { const A = vpN.convertToPdfPoint(l, t), B = vpN.convertToPdfPoint(r, b); return [Math.min(A[0], B[0]), Math.min(A[1], B[1]), Math.max(A[0], B[0]), Math.max(A[1], B[1])].map(n => +n.toFixed(2)); });
}
export function markupFromSelection(type) {
  const got = selectionRects();
  if (!got) { BTS.toast('Select some text first'); return false; }
  pushUndo();
  for (const [i, rects] of got.byPage) {
    S.anns.push({ id: newId(), p: i, type, rects, color: S.colors[type], text: got.text.slice(0, 500) });
    renderPage(i);
  }
  getSelection().removeAllRanges();
  hideSelPop();
  changed();
  return true;
}

/* ---------------- floating selection toolbar ---------------- */
let selPop = null;
function showSelPop(x, y) {
  hideSelPop();
  selPop = document.createElement('div'); selPop.className = 'sel-pop';
  selPop.innerHTML = [['highlight', 'highlighter', 'Highlight'], ['underline', 'underline', 'Underline'], ['strike', 'strikethrough', 'Strikeout'], ['copy', 'copy', 'Copy']]
    .map(([k, ic, t]) => `<button data-k="${k}" title="${t}">${icon(ic, 16)}</button>`).join('');
  selPop.addEventListener('mousedown', e => e.preventDefault());
  selPop.addEventListener('click', e => {
    const b = e.target.closest('[data-k]'); if (!b) return;
    if (b.dataset.k === 'copy') { navigator.clipboard?.writeText(getSelection().toString()).then(() => BTS.toast('Copied')); hideSelPop(); }
    else markupFromSelection(b.dataset.k);
  });
  document.body.appendChild(selPop);
  const w = selPop.offsetWidth;
  selPop.style.left = clamp(x - w / 2, 8, innerWidth - w - 8) + 'px';
  selPop.style.top = clamp(y + 10, 8, innerHeight - 50) + 'px';
}
export function hideSelPop() { selPop?.remove(); selPop = null; }

/* ---------------- note popup & text box editor ---------------- */
let notePop = null;
export function closeNotePopup() { if (notePop) { const n = notePop; notePop = null; n.commit(); n.el.remove(); } }
function openNotePopup(a, focus = true) {
  closeNotePopup();
  const p = S.pages[a.p]; if (!p) return;
  const el = document.createElement('div'); el.className = 'note-pop';
  el.innerHTML = `<header><span class="dot" style="background:${esc(a.color)}"></span>Note<button class="rb" data-x title="Delete note">${icon('trash-2', 15)}</button></header><textarea placeholder="Type your note…" spellcheck="true"></textarea>`;
  const ta = $('textarea', el); ta.value = a.text || '';
  const [vx, vy] = p.vp1.convertToViewportPoint(a.x, a.y);
  el.style.left = `calc(${vx / p.vp1.width * 100}% + ${NOTE.icon * S.scale + 6}px)`;
  el.style.top = `${vy / p.vp1.height * 100}%`;
  p.el.appendChild(el);
  const before = a.text || '';
  const commit = () => {
    const v = ta.value;
    if (v !== before) { S.undo.push(JSON.stringify(S.anns.map(x => x.id === a.id ? { ...x, text: before } : x))); S.redo = []; a.text = v; renderPage(a.p); changed(); }
  };
  notePop = { el, commit, id: a.id };
  el.addEventListener('mousedown', e => e.stopPropagation());
  el.addEventListener('pointerdown', e => e.stopPropagation());
  $('[data-x]', el).addEventListener('click', () => { const id = a.id; notePop = null; el.remove(); deleteAnn(id); });
  ta.addEventListener('keydown', e => { if (e.key === 'Escape') { e.preventDefault(); closeNotePopup(); } e.stopPropagation(); });
  // keep the popup inside the page horizontally
  requestAnimationFrame(() => { const pr = p.el.getBoundingClientRect(), r = el.getBoundingClientRect(); if (r.right > pr.right - 4) { el.style.left = `calc(${vx / p.vp1.width * 100}% - ${r.width + 6}px)`; } });
  if (focus) setTimeout(() => ta.focus(), 0);
}

let tbEdit = null;
export function closeTextEditor() { if (tbEdit) { const t = tbEdit; tbEdit = null; t.commit(); } }
function editTextbox(a, isNew) {
  closeTextEditor(); closeNotePopup();
  const p = S.pages[a.p];
  const ta = document.createElement('textarea'); ta.className = 'tb-editor';
  ta.value = a.text || ''; ta.spellcheck = true;
  const [vx, vy] = p.vp1.convertToViewportPoint(a.x, a.y);
  const fs = (a.size || 14) * S.scale;
  Object.assign(ta.style, { left: vx / p.vp1.width * 100 + '%', top: vy / p.vp1.height * 100 + '%', fontSize: fs + 'px', lineHeight: TB.lh, color: a.color, transform: `rotate(${S.rot}deg)` });
  p.el.appendChild(ta);
  const fit = () => { ta.style.height = 'auto'; ta.style.width = 'auto'; ta.style.width = Math.max(fs * 4, ta.scrollWidth + fs) + 'px'; ta.style.height = ta.scrollHeight + 'px'; };
  ta.addEventListener('input', fit); fit();
  const hiddenId = a.id; a._editing = true; renderPage(a.p);
  const before = a.text || '';
  const commit = () => {
    delete a._editing;
    const v = ta.value.replace(/\s+$/, '');
    ta.remove();
    if (!v) { if (isNew) { S.anns = S.anns.filter(x => x.id !== hiddenId); S.undo.pop(); renderPage(a.p); listeners.change.forEach(f => f()); } else deleteAnn(hiddenId); return; }
    if (v !== before) { if (!isNew) pushUndo(); a.text = v; changed(); }
    renderPage(a.p);
  };
  tbEdit = { commit };
  ta.addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Escape' || (e.key === 'Enter' && (e.ctrlKey || e.metaKey))) { e.preventDefault(); closeTextEditor(); } });
  ta.addEventListener('blur', () => setTimeout(() => { if (tbEdit && tbEdit.commit === commit) closeTextEditor(); }, 0));
  ta.addEventListener('mousedown', e => e.stopPropagation());
  ta.addEventListener('pointerdown', e => e.stopPropagation());
  setTimeout(() => { ta.focus(); ta.setSelectionRange(ta.value.length, ta.value.length); }, 0);
}

/* ---------------- pointer interaction ---------------- */
let draft = null;
const pageOf = (t) => { const el = t.closest?.('.page'); return el ? S.pages[+el.dataset.i] : null; };

export function initInteractions() {
  const host = $('#pages'); const viewer = $('#viewer');

  // Select / markup tools: mouse-based so we can block text selection when grabbing a shape.
  host.addEventListener('mousedown', e => {
    if (e.button !== 0) return;
    const tool = S.tool;
    if (!['text', 'highlight', 'underline', 'strike'].includes(tool)) return;
    if (e.target.closest('.note-pop, .tb-editor, .linkLayer a')) return;
    hideSelPop();
    const p = pageOf(e.target); if (!p) return;
    const a = hitTest(p, e.clientX, e.clientY);
    if (!a) { if (S.sel) select(null); closeNotePopup(); return; }
    if (MARKUP.has(a.type)) {   // let text selection start; select highlight on click without drag
      const sx = e.clientX, sy = e.clientY;
      const up = (ev) => { window.removeEventListener('mouseup', up); if (Math.hypot(ev.clientX - sx, ev.clientY - sy) < 3 && getSelection().isCollapsed) select(a.id); };
      window.addEventListener('mouseup', up);
      return;
    }
    e.preventDefault(); getSelection().removeAllRanges();
    select(a.id);
    if (a.type === 'note') openNotePopup(a, false); else closeNotePopup();
    const start = clientToPdf(p, e.clientX, e.clientY); let moved = false, last = start;
    const mv = (ev) => {
      const cur = clientToPdf(p, ev.clientX, ev.clientY);
      if (!moved && Math.hypot(ev.clientX - e.clientX, ev.clientY - e.clientY) < 3) return;
      if (!moved) { pushUndo(); moved = true; closeNotePopup(); }
      translate(a, cur[0] - last[0], cur[1] - last[1]); last = cur; renderPage(p.i);
    };
    const up = () => { window.removeEventListener('mousemove', mv); window.removeEventListener('mouseup', up); if (moved) { changed(); if (a.type === 'note') openNotePopup(a, false); } };
    window.addEventListener('mousemove', mv); window.addEventListener('mouseup', up);
  });
  host.addEventListener('dblclick', e => {
    const p = pageOf(e.target); if (!p || S.tool !== 'text') return;
    const a = hitTest(p, e.clientX, e.clientY);
    if (a?.type === 'textbox') { e.preventDefault(); getSelection().removeAllRanges(); editTextbox(a, false); }
    else if (a?.type === 'note') { e.preventDefault(); openNotePopup(a, true); }
  });

  // After a text selection: markup tools apply immediately; select tool shows the mini toolbar.
  document.addEventListener('mouseup', e => {
    if (!S.doc || e.button !== 0) return;
    setTimeout(() => {
      const sel = getSelection();
      if (!sel || sel.isCollapsed || !sel.anchorNode || !$('#pages').contains(sel.anchorNode)) return;
      if (MARKUP.has(S.tool)) markupFromSelection(S.tool);
      else if (S.tool === 'text' && sel.toString().trim()) showSelPop(e.clientX, e.clientY);
    }, 0);
  });

  // Drawing tools: pointer events with capture.
  host.addEventListener('pointerdown', e => {
    if (e.button !== 0) return;
    const tool = S.tool;
    if (tool === 'hand') return handDrag(e, viewer);
    if (!['pen', 'rect', 'ellipse', 'arrow', 'note', 'textbox', 'eraser'].includes(tool)) return;
    if (e.target.closest('.note-pop, .tb-editor')) return;
    const p = pageOf(e.target); if (!p) return;
    e.preventDefault();
    closeTextEditor(); hideSelPop();
    const P0 = clientToPdf(p, e.clientX, e.clientY).map(n => +n.toFixed(2));
    if (tool === 'note') {
      const a = addAnn({ p: p.i, type: 'note', x: P0[0], y: P0[1], text: '', color: S.colors.note });
      // anchor so the icon's top-left is where the user clicked, in the display frame
      select(a.id); openNotePopup(a, true); return;
    }
    if (tool === 'textbox') {
      // existing textbox under cursor? edit it instead
      const hitA = hitTest(p, e.clientX, e.clientY);
      if (hitA?.type === 'textbox') { editTextbox(hitA, false); return; }
      pushUndo();
      const size = Math.max(8, Math.round(S.width * 2 + 10));
      const a = { id: newId(), p: p.i, type: 'textbox', x: P0[0], y: P0[1], text: '', color: S.colors.textbox, size };
      S.anns.push(a); editTextbox(a, true); return;
    }
    try { p.el.setPointerCapture(e.pointerId); } catch {}
    if (tool === 'eraser') {
      let rec = false;
      const erase = (ev) => {
        const q = pageOf(document.elementFromPoint(ev.clientX, ev.clientY) || p.el) || p;
        const a = hitTest(q, ev.clientX, ev.clientY);
        if (a) { if (!rec) { pushUndo(); rec = true; } deleteAnn(a.id, { record: false }); }
      };
      erase(e);
      const mv = (ev) => erase(ev);
      const up = () => { p.el.removeEventListener('pointermove', mv); p.el.removeEventListener('pointerup', up); p.el.removeEventListener('pointercancel', up); };
      p.el.addEventListener('pointermove', mv); p.el.addEventListener('pointerup', up); p.el.addEventListener('pointercancel', up);
      return;
    }
    const color = S.colors[tool], w = S.width;
    if (tool === 'pen') draft = { id: 'draft', p: p.i, type: 'pen', pts: [P0], color, w };
    else if (tool === 'arrow') draft = { id: 'draft', p: p.i, type: 'arrow', a: P0, b: P0, color, w };
    else draft = { id: 'draft', p: p.i, type: tool, r: [...P0, ...P0], color, w };
    renderPage(p.i);
    const mv = (ev) => {
      const evs = ev.getCoalescedEvents ? ev.getCoalescedEvents() : [ev];
      for (const ce of (evs.length ? evs : [ev])) {
        let P = clientToPdf(p, ce.clientX, ce.clientY).map(n => +n.toFixed(2));
        if (draft.type === 'pen') {
          const L = draft.pts[draft.pts.length - 1];
          if (Math.hypot(P[0] - L[0], P[1] - L[1]) * S.scale >= 1.5) draft.pts.push(P);
        } else if (draft.type === 'arrow') {
          if (ce.shiftKey) P = snap45(draft.a, P);
          draft.b = P;
        } else {
          if (ce.shiftKey) { const dx = P[0] - P0[0], dy = P[1] - P0[1]; const m = Math.max(Math.abs(dx), Math.abs(dy)); P = [P0[0] + Math.sign(dx || 1) * m, P0[1] + Math.sign(dy || 1) * m]; }
          draft.r = [P0[0], P0[1], P[0], P[1]];
        }
      }
      renderPage(p.i);
    };
    const up = () => {
      p.el.removeEventListener('pointermove', mv); p.el.removeEventListener('pointerup', up); p.el.removeEventListener('pointercancel', up);
      const d = draft; draft = null;
      if (!d) return;
      let ok = true;
      if (d.type === 'pen') { if (d.pts.length > 2) d.pts = simplify(d.pts, 0.35 / Math.max(1, S.scale / 2)); }
      else if (d.type === 'arrow') ok = Math.hypot(d.b[0] - d.a[0], d.b[1] - d.a[1]) * S.scale > 6;
      else { d.r = [Math.min(d.r[0], d.r[2]), Math.min(d.r[1], d.r[3]), Math.max(d.r[0], d.r[2]), Math.max(d.r[1], d.r[3])]; ok = (d.r[2] - d.r[0]) * S.scale > 4 && (d.r[3] - d.r[1]) * S.scale > 4; }
      if (ok) { delete d.id; addAnn(d); } else renderPage(p.i);
    };
    p.el.addEventListener('pointermove', mv); p.el.addEventListener('pointerup', up); p.el.addEventListener('pointercancel', up);
  });

  document.addEventListener('keydown', e => {
    if (!S.doc || isTyping(e.target)) return;
    if ((e.key === 'Delete' || e.key === 'Backspace') && S.sel) { e.preventDefault(); deleteAnn(S.sel); }
    if (e.key === 'Escape') { if (S.sel) select(null); hideSelPop(); }
  });
  viewer.addEventListener('scroll', hideSelPop, { passive: true });
}

function snap45(a, b) {
  const dx = b[0] - a[0], dy = b[1] - a[1]; const ang = Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * (Math.PI / 4); const l = Math.hypot(dx, dy);
  return [a[0] + Math.cos(ang) * l, a[1] + Math.sin(ang) * l];
}
/** Ramer–Douglas–Peucker line simplification (keeps pen strokes small in the saved PDF). */
function simplify(pts, eps) {
  if (pts.length < 3) return pts;
  const keep = new Uint8Array(pts.length); keep[0] = keep[pts.length - 1] = 1;
  const stack = [[0, pts.length - 1]];
  while (stack.length) {
    const [s, e] = stack.pop(); let md = 0, mi = -1;
    for (let i = s + 1; i < e; i++) { const d = distSeg(...pts[i], ...pts[s], ...pts[e]); if (d > md) { md = d; mi = i; } }
    if (md > eps) { keep[mi] = 1; stack.push([s, mi], [mi, e]); }
  }
  return pts.filter((_, i) => keep[i]);
}

function handDrag(e, viewer) {
  if (e.target.closest('.note-pop, .tb-editor, .linkLayer a')) return;
  e.preventDefault();
  const sx = e.clientX, sy = e.clientY, sl = viewer.scrollLeft, st = viewer.scrollTop;
  document.body.classList.add('grabbing');
  const mv = (ev) => { viewer.scrollLeft = sl - (ev.clientX - sx); viewer.scrollTop = st - (ev.clientY - sy); };
  const up = () => { document.body.classList.remove('grabbing'); window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up); };
  window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up);
}

export function resetAnnotations() {
  S.anns = []; S.sel = null; S.undo = []; S.redo = []; draft = null;
  closeNotePopup(); closeTextEditor(); hideSelPop();
}
export const annCount = () => S.anns.length;
