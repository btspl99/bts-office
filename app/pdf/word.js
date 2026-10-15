/* BTS PDF — PDF → Word (.docx) with docx.js: rebuilds lines and paragraphs from pdf.js text items,
 * detects headings by font size, bold/italic from font names, simple two-column pages, page breaks,
 * and pictures (embedded images, or the whole page when it has no text — e.g. a scan). */
import { S, $, esc, loadScript, pdfjsLib, fmtSize } from './state.js';
import * as IO from './io.js';
import { progress, Cancelled, nextFrame } from './ui.js';

const TW = 20;   // twips per point
const DEVA = /[ऀ-ॿ]/;

/* ---------------- fonts ---------------- */
function fontInfo(page, fontName, style) {
  let name = '', bold = false, italic = false;
  try { const f = page.commonObjs.get(fontName); name = f?.name || ''; bold = !!f?.bold || !!f?.black; italic = !!f?.italic; } catch {}
  const raw = name.replace(/^[A-Z]{6}\+/, '');
  if (/bold|black|heavy|semibold|demi/i.test(raw)) bold = true;
  if (/italic|oblique/i.test(raw)) italic = true;
  let fam = raw.split(/[-,]/)[0].replace(/(PS)?MT$/, '').replace(/([a-z])([A-Z])/g, '$1 $2').trim();
  const map = [[/^arial|helvetica|liberation sans|arimo/i, 'Arial'], [/^times|liberation serif|tinos/i, 'Times New Roman'], [/^courier|liberation mono|cousine/i, 'Courier New'], [/^calibri|carlito/i, 'Calibri'], [/^cambria|caladea/i, 'Cambria'], [/^symbol/i, 'Symbol']];
  for (const [re, n] of map) if (re.test(fam)) { fam = n; break; }
  if (!fam || /^(g_d|f\d|t\d|ff\d|font)/i.test(fam)) fam = style?.fontFamily === 'serif' ? 'Times New Roman' : style?.fontFamily === 'monospace' ? 'Courier New' : 'Calibri';
  return { fam, bold, italic };
}

/* ---------------- layout analysis ---------------- */
/** pdf.js text items -> lines in reading order (display frame, y down). */
function buildLines(page, tc, vp) {
  const items = [];
  for (const it of tc.items) {
    if (!it.str || !it.str.length) continue;
    const t = pdfjsLib.Util.transform(vp.transform, it.transform);
    const size = Math.hypot(t[2], t[3]);
    if (size < 1) continue;
    if (Math.abs(t[1]) > Math.abs(t[0]) * 0.2) continue;      // rotated text: skip
    const x = t[4], y = t[5];                                  // baseline start
    const w = it.width * vp.scale; const fi = fontInfo(page, it.fontName, tc.styles[it.fontName]);
    items.push({ str: it.str, x, y, w, x2: x + w, size, ...fi });
  }
  // group into lines by baseline
  items.sort((a, b) => a.y - b.y || a.x - b.x);
  const lines = [];
  for (const it of items) {
    const ln = lines.find(l => Math.abs(l.y - it.y) < Math.max(l.size, it.size) * 0.45 && it.x < l.x2 + Math.max(l.size, it.size) * 3 && it.x2 > l.x - it.size * 2);
    if (ln) { ln.items.push(it); ln.x = Math.min(ln.x, it.x); ln.x2 = Math.max(ln.x2, it.x2); ln.size = Math.max(ln.size, it.size); }
    else lines.push({ y: it.y, x: it.x, x2: it.x2, size: it.size, items: [it] });
  }
  for (const ln of lines) {
    ln.items.sort((a, b) => a.x - b.x);
    // merge into runs, inserting spaces where there are gaps
    const runs = []; let prev = null;
    for (const it of ln.items) {
      let s = it.str;
      if (prev) {
        const gap = it.x - prev.x2;
        if (gap > it.size * 0.18 && !/\s$/.test(prev.str) && !/^\s/.test(s)) s = ' ' + s;
      }
      const last = runs[runs.length - 1];
      if (last && last.bold === it.bold && last.italic === it.italic && Math.abs(last.size - it.size) < 0.6 && last.fam === it.fam) last.text += s;
      else runs.push({ text: s, bold: it.bold, italic: it.italic, size: it.size, fam: it.fam });
      prev = it;
    }
    ln.runs = runs;
    ln.text = runs.map(r => r.text).join('');
    ln.bold = runs.every(r => r.bold || !r.text.trim());
  }
  return lines.filter(l => l.text.trim());
}

/** Order lines: detect a two-column band layout, otherwise top to bottom. */
function readingOrder(lines, pageW) {
  if (lines.length < 8) return lines.sort((a, b) => a.y - b.y);
  // find a vertical gutter in the middle third that few lines cross
  let best = null;
  for (let gx = pageW * 0.35; gx <= pageW * 0.65; gx += 4) {
    let cross = 0, left = 0, right = 0;
    for (const l of lines) { if (l.x < gx - 2 && l.x2 > gx + 2) cross++; else if (l.x2 <= gx + 2) left++; else right++; }
    const score = Math.min(left, right) - cross * 3;
    if (!best || score > best.score) best = { gx, score, cross, left, right };
  }
  if (!best || best.left < 4 || best.right < 4 || best.cross > lines.length * 0.25) return lines.sort((a, b) => a.y - b.y);
  const gx = best.gx;
  const sorted = lines.slice().sort((a, b) => a.y - b.y);
  const out = []; let band = [];
  const flush = () => { out.push(...band.filter(l => l.x2 <= gx + 2), ...band.filter(l => l.x2 > gx + 2)); band = []; };
  for (const l of sorted) {
    if (l.x < gx - 2 && l.x2 > gx + 2) { flush(); out.push(l); } else band.push(l);
  }
  flush();
  out.twoCol = true;
  return out;
}

/** Lines -> paragraphs. */
function buildParagraphs(lines, pageW) {
  const paras = []; let cur = null;
  for (const l of lines) {
    const gap = cur ? l.y - cur.lastY : 0;
    const sameStyle = cur && Math.abs(cur.size - l.size) < 1 && cur.bold === l.bold;
    const leftOk = cur && (Math.abs(l.x - cur.x) < l.size * 1.5 || Math.abs(l.x - cur.lastX) < l.size * 0.8);
    const join = cur && sameStyle && leftOk && gap > 0 && gap < l.size * 1.75 && !/^\s*([••\-–*]|\d+[.)]|[a-z][.)])\s/.test(l.text);
    if (join) {
      const lastRun = cur.runs[cur.runs.length - 1];
      if (/[-‐]$/.test(lastRun.text) && /^[a-z]/.test(l.text)) lastRun.text = lastRun.text.replace(/[-‐]$/, '');
      else if (!/\s$/.test(lastRun.text)) lastRun.text += ' ';
      cur.runs.push(...l.runs.map(r => ({ ...r }))); cur.lastY = l.y; cur.lastX = l.x; cur.x2 = Math.max(cur.x2, l.x2); cur.lines++;
    } else {
      cur = { runs: l.runs.map(r => ({ ...r })), size: l.size, bold: l.bold, x: l.x, x2: l.x2, lastY: l.y, lastX: l.x, firstY: l.y, lines: 1, gapBefore: gap };
      paras.push(cur);
    }
  }
  for (const p of paras) {
    const mid = (p.x + p.x2) / 2;
    p.center = p.lines <= 3 && Math.abs(mid - pageW / 2) < pageW * 0.05 && p.x > pageW * 0.18;
    p.text = p.runs.map(r => r.text).join('');
  }
  return paras;
}

/* ---------------- pictures ---------------- */
async function findImages(page, vp) {
  const ops = await page.getOperatorList();
  const O = pdfjsLib.OPS; const out = [];
  let ctm = [1, 0, 0, 1, 0, 0]; const stack = [];
  const mul = pdfjsLib.Util.transform;
  for (let i = 0; i < ops.fnArray.length; i++) {
    const fn = ops.fnArray[i], args = ops.argsArray[i];
    if (fn === O.save) stack.push(ctm);
    else if (fn === O.restore) ctm = stack.pop() || [1, 0, 0, 1, 0, 0];
    else if (fn === O.transform) ctm = mul(ctm, args);
    else if (fn === O.paintFormXObjectBegin) { stack.push(ctm); if (args?.[0]) ctm = mul(ctm, args[0]); }
    else if (fn === O.paintFormXObjectEnd) ctm = stack.pop() || ctm;
    else if (fn === O.paintImageXObject || fn === O.paintInlineImageXObject || fn === O.paintImageXObjectRepeat) {
      const m = mul(vp.transform, ctm);
      const pts = [[0, 0], [1, 0], [0, 1], [1, 1]].map(([x, y]) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]]);
      const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
      const box = { x: Math.max(0, Math.min(...xs)), y: Math.max(0, Math.min(...ys)), x2: Math.min(vp.width, Math.max(...xs)), y2: Math.min(vp.height, Math.max(...ys)) };
      if (box.x2 - box.x > 30 && box.y2 - box.y > 30) out.push(box);
    }
  }
  return out;
}
const pageCanvases = new Map();
async function pageCanvas(page, k) {
  const key = page._pageIndex + ':' + k;
  if (pageCanvases.has(key)) return pageCanvases.get(key);
  const v = page.getViewport({ scale: k, rotation: page.rotate });
  const c = document.createElement('canvas'); c.width = Math.round(v.width); c.height = Math.round(v.height);
  const ctx = c.getContext('2d', { alpha: false }); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
  await page.render({ canvasContext: ctx, viewport: v, intent: 'print' }).promise;
  pageCanvases.set(key, c);
  return c;
}
function freeCanvases() { pageCanvases.forEach(c => { c.width = c.height = 0; }); pageCanvases.clear(); }
async function renderCrop(page, vp, box, dpi = 150) {
  const k = dpi / 72;
  const c = await pageCanvas(page, k);
  const b = box || { x: 0, y: 0, x2: vp.width, y2: vp.height };
  const o = document.createElement('canvas'); o.width = Math.max(1, Math.round((b.x2 - b.x) * k)); o.height = Math.max(1, Math.round((b.y2 - b.y) * k));
  o.getContext('2d').drawImage(c, Math.round(b.x * k), Math.round(b.y * k), o.width, o.height, 0, 0, o.width, o.height);
  const blob = await new Promise(r => o.toBlob(r, box ? 'image/png' : 'image/jpeg', 0.85));
  return { data: new Uint8Array(await blob.arrayBuffer()), type: box ? 'png' : 'jpg' };
}

/* ---------------- docx building ---------------- */
function runOf(D, r, bodySize, heading) {
  const opts = { text: r.text, bold: r.bold || undefined, italics: r.italic || undefined, font: DEVA.test(r.text) ? { ascii: r.fam, hAnsi: r.fam, cs: 'Nirmala UI' } : r.fam };
  if (!heading || Math.abs(r.size - bodySize) > 0.6) opts.size = Math.max(8, Math.round(r.size * 2));
  return new D.TextRun(opts);
}

export async function pdfToWord() {
  if (!S.doc) return;
  const h = progress('Converting to Word', 'Reading the pages…');
  let bytes;
  try { bytes = await buildDocx(h); }
  catch (e) {
    h.close(); freeCanvases();
    if (e instanceof Cancelled) return BTS.toast('Conversion stopped');
    console.warn(e); return BTS.toast('Couldn’t convert: ' + (e.message || e), { error: true, ms: 5000 });
  }
  h.close();
  const path = await IO.saveBytes(`${BTS.title}.docx`, bytes, 'Word document', 'docx', { quiet: true });
  if (typeof path === 'string') {
    const go = await BTS.dialog({ title: 'Word file ready', body: `<p class="dlg-p">Saved <b>${esc(path.split(/[/\\]/).pop())}</b> (${fmtSize(bytes.length)}). Layouts with tables or many columns may need a little tidying.</p>`, buttons: [{ label: 'Close', value: false }, { label: 'Open in BTS Docs', primary: true, value: true }] });
    if (go) bts.openApp('docs', path);
  }
  return bytes;
}

/** Convert the open PDF to .docx bytes. h: progress handle (optional). */
export async function buildDocx(h = { set() {}, cancelled: false }) {
  const D = await loadScript('docx.iife.js', 'docx');
  const n = S.pages.length;
  let doc;
  try {
    // body text size = the most common size (by characters)
    const pagesData = [];
    const sizes = new Map();
    for (let i = 0; i < n; i++) {
      if (h.cancelled) throw new Cancelled();
      h.set(i / n * 0.5, `Reading page ${i + 1} of ${n}…`);
      const page = S.pages[i].proxy;
      let vp = page.getViewport({ scale: 1, rotation: page.rotate });
      await page.getOperatorList();      // loads fonts so their real names (bold/italic) are known
      const tc = await page.getTextContent();
      let lines = buildLines(page, tc, vp);
      if (page.rotate && tc.items.length && lines.length < tc.items.filter(i => i.str?.trim()).length / 3) {
        // text runs sideways on a turned page: read it the way it was written
        const vp0 = page.getViewport({ scale: 1, rotation: 0 }); const l0 = buildLines(page, tc, vp0);
        if (l0.length > lines.length) { vp = vp0; lines = l0; }
      }
      lines.forEach(l => l.runs.forEach(r => { const k = Math.round(r.size * 2) / 2; sizes.set(k, (sizes.get(k) || 0) + r.text.length); }));
      pagesData.push({ page, vp, lines });
      await nextFrame();
    }
    const body = [...sizes.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || 11;
    const sections = [];
    for (let i = 0; i < n; i++) {
      if (h.cancelled) throw new Cancelled();
      h.set(0.5 + i / n * 0.45, `Building page ${i + 1} of ${n}…`);
      const { page, vp, lines } = pagesData[i];
      const W = vp.width, H = vp.height;
      const chars = lines.reduce((s, l) => s + l.text.trim().length, 0);
      const children = [];
      let margin = 72;
      if (chars < 15) {
        // no real text (a scan or a drawing): keep the page as a picture
        const img = await renderCrop(page, vp, null, 150);
        margin = 36;
        const maxW = W - 2 * margin, maxH = H - 2 * margin - 10; const k = Math.min(maxW / W, maxH / H);
        children.push(new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.ImageRun({ type: img.type, data: img.data, transformation: { width: Math.round(W * k * 96 / 72), height: Math.round(H * k * 96 / 72) } })] }));
      } else {
        const ordered = readingOrder(lines, W);
        const minX = Math.min(...lines.map(l => l.x)), maxX = Math.max(...lines.map(l => l.x2));
        margin = Math.max(36, Math.min(90, minX, W - maxX));
        const paras = buildParagraphs(ordered, W);
        // pictures: place them before the first paragraph that starts below them
        let pics = [];
        try { pics = (await findImages(page, vp)).filter(b => (b.x2 - b.x) * (b.y2 - b.y) < W * H * 0.85); } catch {}
        pics.sort((a, b) => a.y - b.y);
        const usable = W - 2 * margin;
        const flushPics = async (beforeY) => {
          while (pics.length && pics[0].y < beforeY) {
            const b = pics.shift(); const img = await renderCrop(page, vp, b, 150);
            let w = b.x2 - b.x, hh = b.y2 - b.y; const k = Math.min(1, usable / w); w *= k; hh *= k;
            children.push(new D.Paragraph({ alignment: Math.abs((b.x + b.x2) / 2 - W / 2) < W * 0.1 ? D.AlignmentType.CENTER : D.AlignmentType.LEFT, children: [new D.ImageRun({ type: img.type, data: img.data, transformation: { width: Math.round(w * 96 / 72), height: Math.round(hh * 96 / 72) } })] }));
          }
        };
        for (const p of paras) {
          if (!ordered.twoCol) await flushPics(p.firstY - p.size);
          const big = p.size / body;
          const short = p.text.length < 140 && p.lines <= 3;
          const heading = short && (big >= 1.55 ? D.HeadingLevel.HEADING_1 : big >= 1.25 ? D.HeadingLevel.HEADING_2 : (big >= 1.08 && p.bold) ? D.HeadingLevel.HEADING_3 : null);
          children.push(new D.Paragraph({
            heading: heading || undefined,
            alignment: p.center ? D.AlignmentType.CENTER : undefined,
            spacing: { after: Math.round(Math.min(18, p.size * 0.6) * TW), before: heading ? Math.round(p.size * 0.6 * TW) : undefined },
            indent: !p.center && p.x - margin > 12 && !ordered.twoCol ? { left: Math.round(Math.min(usable / 2, p.x - margin) * TW) } : undefined,
            children: p.runs.map(r => runOf(D, r, body, !!heading)),
          }));
        }
        await flushPics(Infinity);
      }
      freeCanvases();
      const m = Math.round(margin * TW);
      sections.push({
        properties: { page: { size: { width: Math.round(W * TW), height: Math.round(H * TW), orientation: W > H ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT }, margin: { top: m, bottom: m, left: m, right: m } } },
        children: children.length ? children : [new D.Paragraph('')],
      });
      await nextFrame();
    }
    h.set(0.97, 'Saving…');
    doc = new D.Document({
      creator: BTS.userName || 'BTS PDF', title: S.meta?.info?.Title || BTS.title,
      styles: { default: { document: { run: { font: 'Calibri', size: Math.round(body * 2) } } } },
      sections,
    });
    const blob = await D.Packer.toBlob(doc);
    return new Uint8Array(await blob.arrayBuffer());
  } finally { freeCanvases(); }
}

/** Plain text (one string per page) -> .docx bytes (used by OCR "Save Word"). */
export async function textToDocx(pages, title) {
  const D = await loadScript('docx.iife.js', 'docx');
  const sections = pages.map(t => ({
    properties: {},
    children: (t.trim() ? t.trim().split(/\n{2,}/) : ['']).map(par => new D.Paragraph({
      spacing: { after: 160 },
      children: par.split('\n').flatMap((ln, i) => [...(i ? [new D.TextRun({ break: 1 })] : []), new D.TextRun({ text: ln, font: DEVA.test(ln) ? { ascii: 'Calibri', hAnsi: 'Calibri', cs: 'Nirmala UI' } : 'Calibri' })]),
    })),
  }));
  const doc = new D.Document({ creator: 'BTS PDF', title, sections });
  return new Uint8Array(await (await D.Packer.toBlob(doc)).arrayBuffer());
}
