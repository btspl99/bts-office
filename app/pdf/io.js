/* BTS PDF — saving (annotations burned in with pdf-lib), printing, and page tools */
import { S, $, esc, loadPdfLib, hexToRgb01, NOTE, wrapText, measureHelv, baseViewport } from './state.js';
import { nativeVp, markupLine, textLayout, noteLayout, arrowGeom, svgContent, TB } from './annotations.js';

/* ======================= burn-in ======================= */
const f2 = (n) => +(+n).toFixed(3);

/** Build the pdf-lib operators that draw all annotations of page p (in PDF user space). */
function pageOps(L, page, p, font) {
  const ops = [];
  const { pushGraphicsState: q, popGraphicsState: Q, setStrokingRgbColor: SC, setFillingRgbColor: FC, setLineWidth: W, moveTo: M, lineTo: Lt, closePath: h, stroke: S_, fill: f, appendBezierCurve: C, setLineCap, setLineJoin, LineCapStyle, LineJoinStyle, concatTransformationMatrix: cm } = L;
  const col = (hex) => hexToRgb01(hex);
  const vpN = nativeVp(p);
  const D = (x, y) => vpN.convertToPdfPoint(x, y);            // display frame -> user space
  const Vp = (x, y) => vpN.convertToViewportPoint(x, y);      // user space -> display frame
  const poly = (list) => { list.forEach(([x, y], i) => ops.push(i ? Lt(f2(x), f2(y)) : M(f2(x), f2(y)))); ops.push(h()); };
  const gs = (o) => page.maybeEmbedGraphicsState(o);
  // text in the display frame: baseline origin at display (dx,dy), font size s
  const textAt = (str, dx, dy, s) => {
    const o = D(dx, dy), ux = D(dx + 1, dy), up = D(dx, dy - 1);
    const a = ux[0] - o[0], b = ux[1] - o[1], c = up[0] - o[0], d = up[1] - o[1];
    ops.push(L.beginText(), L.setFontAndSize(font.key, s), L.setTextMatrix(f2(a), f2(b), f2(c), f2(d), f2(o[0]), f2(o[1])), L.showText(font.enc(str)), L.endText());
  };
  const ellipse = (cx, cy, rx, ry) => {      // in display frame, 4 béziers
    const k = 0.5522847498;
    const P = (x, y) => D(x, y).map(f2);
    const seg = [[cx + rx, cy], [cx, cy + ry], [cx - rx, cy], [cx, cy - ry], [cx + rx, cy]];
    ops.push(M(...P(...seg[0])));
    const ctrl = [
      [[cx + rx, cy + ry * k], [cx + rx * k, cy + ry]], [[cx - rx * k, cy + ry], [cx - rx, cy + ry * k]],
      [[cx - rx, cy - ry * k], [cx - rx * k, cy - ry]], [[cx + rx * k, cy - ry], [cx + rx, cy - ry * k]],
    ];
    for (let i = 0; i < 4; i++) ops.push(C(...P(...ctrl[i][0]), ...P(...ctrl[i][1]), ...P(...seg[i + 1])));
    ops.push(h());
  };

  const anns = S.anns.filter(a => a.p === p.i);
  // highlights first (under everything), multiply so text stays readable
  for (const a of anns.filter(a => a.type === 'highlight')) {
    const g = gs({ opacity: 0.42, blendMode: L.BlendMode.Multiply });
    ops.push(q(), L.setGraphicsState(g), FC(...col(a.color)));
    for (const [x1, y1, x2, y2] of a.rects) { poly([[x1, y1], [x2, y1], [x2, y2], [x1, y2]]); ops.push(f()); }
    ops.push(Q());
  }
  for (const a of anns) {
    const c = col(a.color || '#000000');
    switch (a.type) {
      case 'underline': case 'strike':
        ops.push(q(), SC(...c), setLineCap(LineCapStyle.Butt));
        for (const r of a.rects) { const m = markupLine(p, a, r); ops.push(W(f2(m.w)), M(f2(m.from[0]), f2(m.from[1])), Lt(f2(m.to[0]), f2(m.to[1])), S_()); }
        ops.push(Q()); break;
      case 'pen':
        ops.push(q(), SC(...c), W(f2(a.w)), setLineCap(LineCapStyle.Round), setLineJoin(LineJoinStyle.Round));
        a.pts.forEach(([x, y], i) => ops.push(i ? Lt(f2(x), f2(y)) : M(f2(x), f2(y))));
        if (a.pts.length === 1) ops.push(Lt(f2(a.pts[0][0] + 0.01), f2(a.pts[0][1])));
        ops.push(S_(), Q()); break;
      case 'rect': {
        const [x1, y1, x2, y2] = a.r;
        ops.push(q(), SC(...c), W(f2(a.w)), setLineJoin(LineJoinStyle.Miter)); poly([[x1, y1], [x2, y1], [x2, y2], [x1, y2]]); ops.push(S_(), Q()); break;
      }
      case 'ellipse': {
        const A = Vp(a.r[0], a.r[1]), B = Vp(a.r[2], a.r[3]);
        const cx = (A[0] + B[0]) / 2, cy = (A[1] + B[1]) / 2, rx = Math.abs(A[0] - B[0]) / 2, ry = Math.abs(A[1] - B[1]) / 2;
        ops.push(q(), SC(...c), W(f2(a.w))); ellipse(cx, cy, rx, ry); ops.push(S_(), Q()); break;
      }
      case 'arrow': {
        const g = arrowGeom(a.a, a.b, a);   // computed directly in user space (shape is rotation-invariant)
        ops.push(q(), SC(...c), FC(...c), W(f2(a.w)), setLineCap(LineCapStyle.Round), M(f2(a.a[0]), f2(a.a[1])), Lt(f2(g.base[0]), f2(g.base[1])), S_());
        poly(g.head); ops.push(f(), Q()); break;
      }
      case 'textbox': {
        const Lay = textLayout(a);
        const [dx, dy] = Vp(a.x, a.y);
        ops.push(q(), FC(...c));
        Lay.lines.forEach((line, i) => { if (line) textAt(line, dx, dy + Lay.base + i * Lay.lh, Lay.size); });
        ops.push(Q()); break;
      }
      case 'note': {
        const [dx, dy] = Vp(a.x, a.y); const s = NOTE.icon;
        // icon
        ops.push(q(), FC(...c), SC(0.35, 0.35, 0.35), W(0.8)); poly([D(dx, dy), D(dx + s, dy), D(dx + s, dy + s), D(dx, dy + s)]); ops.push(L.fillAndStroke());
        ops.push(SC(0.3, 0.3, 0.3), W(1.2), setLineCap(LineCapStyle.Round));
        for (const [y, x2] of [[6, 13.5], [9, 13.5], [12, 10.5]]) { const A = D(dx + 4.5, dy + y), B = D(dx + x2, dy + y); ops.push(M(f2(A[0]), f2(A[1])), Lt(f2(B[0]), f2(B[1])), S_()); }
        ops.push(Q());
        // callout with the note text
        if (a.text) {
          const N = noteLayout(a, p); const bx = dx + N.ox;
          ops.push(q(), FC(1, 0.996, 0.96), SC(...c), W(0.8)); poly([D(bx, dy), D(bx + N.w, dy), D(bx + N.w, dy + N.h), D(bx, dy + N.h)]); ops.push(L.fillAndStroke());
          ops.push(FC(0.106, 0.106, 0.106));
          N.lines.forEach((line, i) => { if (line) textAt(line, bx + NOTE.pad, dy + NOTE.pad + NOTE.fs * 0.85 + i * NOTE.lh, NOTE.fs); });
          ops.push(Q());
        }
        break;
      }
    }
  }
  return ops;
}

/** Helvetica (WinAnsi) with a safe encoder: unsupported characters become '?'. */
const SUBST = { '\u2018': "'", '\u2019': "'", '\u201C': '"', '\u201D': '"', '\u2013': '-', '\u2014': '-', '\u20B9': 'Rs.', '\u2026': '...', '\u00A0': ' ' };
function makeFont(doc, L) {
  const font = doc.embedStandardFont(L.StandardFonts.Helvetica);
  const charset = new Set(font.getCharacterSet());
  const enc = (str) => font.encodeText([...String(str).normalize('NFC')].map(ch => charset.has(ch.codePointAt(0)) ? ch : (SUBST[ch] ?? '?')).join(''));
  /** per-page handle: registers the font in that page's resources on first use */
  return (page) => { let k = null; return { enc, get key() { return (k ||= page.node.newFontDictionary('Helv', font.ref)); } }; };
}

/** Bytes of the document with annotations burned in. */
export async function buildAnnotatedPdf() {
  if (!S.anns.length && !S.modified) return S.bytes;
  const L = await loadPdfLib();
  let doc;
  try { doc = await L.PDFDocument.load(S.bytes, { updateMetadata: false }); }
  catch (e) {
    if (e instanceof L.EncryptedPDFError || /encrypt/i.test(e.message)) return rasterPdf(L);
    throw e;
  }
  if (!S.anns.length) return await doc.save({ useObjectStreams: false });
  const pages = doc.getPages();
  const fontFor = makeFont(doc, L);
  for (const p of S.pages) {
    if (!S.anns.some(a => a.p === p.i)) continue;
    const page = pages[p.i]; if (!page) continue;
    const ops = pageOps(L, page, p, fontFor(page));
    page.pushOperators(L.pushGraphicsState(), ...ops, L.popGraphicsState());
  }
  return await doc.save({ useObjectStreams: false });
}

/** Fallback for encrypted PDFs: rebuild from page images + vector annotations. */
async function rasterPdf(L) {
  const doc = await L.PDFDocument.create();
  const fontFor = makeFont(doc, L);
  for (const p of S.pages) {
    const vpN = nativeVp(p);
    const scale = 2;
    const vp = p.proxy.getViewport({ scale, rotation: p.proxy.rotate });
    const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
    await p.proxy.render({ canvasContext: c.getContext('2d', { alpha: false }), viewport: vp, intent: 'print' }).promise;
    const jpg = await (await new Promise(r => c.toBlob(r, 'image/jpeg', 0.9))).arrayBuffer();
    c.width = c.height = 0;
    const W = vpN.width, H = vpN.height;
    const page = doc.addPage([W, H]);
    const img = await doc.embedJpg(jpg);
    page.drawImage(img, { x: 0, y: 0, width: W, height: H });
    if (S.anns.some(a => a.p === p.i)) {
      const [a, b, cc, d, e, f] = vpN.transform;   // user -> display (y down)
      page.pushOperators(L.pushGraphicsState(), L.concatTransformationMatrix(a, -b, cc, -d, e, H - f), ...pageOps(L, page, p, fontFor(page)), L.popGraphicsState());
    }
  }
  return await doc.save();
}

/* ======================= printing ======================= */
let printUrls = [];
/** Fill #print-root with one image per page (+ vector annotation overlay). Returns printToPDF options. */
export async function preparePrint() {
  const root = $('#print-root');
  cleanupPrint();
  if (!S.doc) return {};
  BTS.toast('Preparing pages for printing…', { ms: 1500 });
  const sizes = new Map(); let css = '';
  const scale = S.pages.length > 150 ? 1.6 : 2.2;   // ~115–160 dpi
  const frag = document.createDocumentFragment();
  for (const p of S.pages) {
    const vpN = nativeVp(p);
    const key = `${Math.round(vpN.width)}x${Math.round(vpN.height)}`;
    if (!sizes.has(key)) { const n = 'ps' + sizes.size; sizes.set(key, n); css += `@page ${n} { size: ${vpN.width}pt ${vpN.height}pt; margin: 0; }\n.pp-${n} { page: ${n}; width: ${vpN.width}pt; height: ${vpN.height}pt; }\n`; }
    const vp = p.proxy.getViewport({ scale, rotation: p.proxy.rotate });
    const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
    await p.proxy.render({ canvasContext: c.getContext('2d', { alpha: false }), viewport: vp, intent: 'print', annotationMode: 2 }).promise;
    const blob = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.92));
    c.width = c.height = 0;
    const url = URL.createObjectURL(blob); printUrls.push(url);
    const div = document.createElement('div'); div.className = 'print-page pp-' + sizes.get(key);
    const vb = `0 0 ${vpN.width} ${vpN.height}`;
    div.innerHTML = `<img src="${url}" alt="">` +
      (S.anns.some(a => a.p === p.i) ? `<svg viewBox="${vb}" preserveAspectRatio="none" style="mix-blend-mode:multiply">${svgContent(p, vpN, 'hl', { print: true })}</svg><svg viewBox="${vb}" preserveAspectRatio="none">${svgContent(p, vpN, 'ann', { print: true })}</svg>` : '');
    frag.appendChild(div);
  }
  const style = document.createElement('style'); style.id = 'print-css'; style.textContent = css;
  document.head.appendChild(style);
  root.replaceChildren(frag);
  await Promise.all([...root.querySelectorAll('img')].map(i => i.decode().catch(() => {})));
  document.body.classList.add('printing');
  const first = nativeVp(S.pages[0]);
  return { pageSize: { width: first.width / 72, height: first.height / 72 }, landscape: false };
}
export function cleanupPrint() {
  document.body.classList.remove('printing');
  $('#print-css')?.remove();
  $('#print-root').replaceChildren();
  printUrls.forEach(u => URL.revokeObjectURL(u)); printUrls = [];
}

/* ======================= tools ======================= */
async function saveBytes(name, bytes, label, ext) {
  if (window.bts?.isDesktop) {
    const path = await bts.saveDialog(name, [{ name: label, extensions: [ext] }]);
    if (!path) return false;
    const r = await bts.writeFile(path, bytes);
    if (r.error) { BTS.toast('Couldn’t save: ' + r.error, { error: true }); return false; }
    BTS.toast('Saved ' + r.name); return true;
  }
  BTS.download(name, bytes, ext === 'pdf' ? 'application/pdf' : 'image/png'); return true;
}

export async function extractPage(i) {
  const L = await loadPdfLib();
  const src = await L.PDFDocument.load(await buildAnnotatedPdf(), { ignoreEncryption: true });
  const out = await L.PDFDocument.create();
  const [pg] = await out.copyPages(src, [i]);
  out.addPage(pg);
  return saveBytes(`${BTS.title} - page ${i + 1}.pdf`, await out.save(), 'PDF document', 'pdf');
}

export async function pageImage(i) {
  const p = S.pages[i]; const vpN = nativeVp(p);
  const scale = 2.5;
  const vp = p.proxy.getViewport({ scale, rotation: p.proxy.rotate });
  const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
  const ctx = c.getContext('2d');
  await p.proxy.render({ canvasContext: ctx, viewport: vp, intent: 'print' }).promise;
  if (S.anns.some(a => a.p === i)) {
    const svg = (part) => `<svg xmlns="http://www.w3.org/2000/svg" width="${c.width}" height="${c.height}" viewBox="0 0 ${vpN.width} ${vpN.height}">${svgContent(p, vpN, part, { print: true })}</svg>`;
    for (const part of ['hl', 'ann']) {
      const img = new Image(); img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg(part));
      await img.decode();
      ctx.globalCompositeOperation = part === 'hl' ? 'multiply' : 'source-over';
      ctx.drawImage(img, 0, 0);
    }
  }
  const blob = await new Promise(r => c.toBlob(r, 'image/png'));
  c.width = c.height = 0;
  return saveBytes(`${BTS.title} - page ${i + 1}.png`, new Uint8Array(await blob.arrayBuffer()), 'PNG image', 'png');
}

/** Append the pages of another PDF. Returns the merged bytes (caller reloads). */
export async function combineWith(otherBytes) {
  const L = await loadPdfLib();
  let a;
  try { a = await L.PDFDocument.load(S.bytes); }
  catch (e) { throw new Error('This PDF is password-protected, so pages can’t be added to it.'); }
  let b;
  try { b = await L.PDFDocument.load(otherBytes); }
  catch (e) { throw new Error('The other PDF is password-protected or damaged.'); }
  const pages = await a.copyPages(b, b.getPageIndices());
  pages.forEach(pg => a.addPage(pg));
  return { bytes: await a.save({ useObjectStreams: false }), added: pages.length };
}
