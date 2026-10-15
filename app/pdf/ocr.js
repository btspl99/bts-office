/* BTS PDF — Recognize text (OCR) in scanned pages with tesseract.js, fully offline.
 * Engine, language data (English, Hindi, Marathi) and core live in app/vendor/tesseract/ and are
 * loaded only when this tool is first used. Results: an invisible text layer written into the PDF
 * (so the scan becomes searchable / selectable) and a side panel with the text. */
import { S, $, $$, esc, loadScript, loadPdfLib, APP } from './state.js';
import * as IO from './io.js';
import { parseRanges } from './organize.js';
import { progress, Cancelled } from './ui.js';
import { textToDocx } from './word.js';

const LANGS = [['eng', 'English'], ['hin', 'Hindi (हिन्दी)'], ['mar', 'Marathi (मराठी)']];
const BASE = new URL('../vendor/tesseract/', import.meta.url).href;
let worker = null, workerLangs = '';

/* ---------------- engine ---------------- */
const SIMD_TEST = new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0, 1, 5, 1, 96, 0, 1, 123, 3, 2, 1, 0, 10, 10, 1, 8, 0, 65, 0, 253, 15, 253, 98, 11]);
async function getWorker(langs, log) {
  if (worker && workerLangs === langs) return worker;
  await dropWorker();
  const T = await loadScript('tesseract/tesseract.min.js', 'Tesseract');
  const simd = (() => { try { return WebAssembly.validate(SIMD_TEST); } catch { return false; } })();
  worker = await T.createWorker(langs.split('+'), 1, {
    workerPath: BASE + 'worker.min.js',
    corePath: BASE + (simd ? 'tesseract-core-simd-lstm.wasm.js' : 'tesseract-core-lstm.wasm.js'),
    langPath: BASE.replace(/\/$/, ''), gzip: true, cacheMethod: 'none', workerBlobURL: false,
    logger: (m) => log?.(m),
  });
  workerLangs = langs;
  return worker;
}
async function dropWorker() { const w = worker; worker = null; workerLangs = ''; if (w) { try { await w.terminate(); } catch {} } }

const DEVA = /[ऀ-ॿ]/;
/** Tesseract often reads the Devanagari full stop (।) as a bar. */
function tidy(s, deva) {
  if (!deva) return s;
  return s.replace(/(^|[ऀ-ॿ\s])[|]/g, '$1।').replace(/[|](?=\s|$)/g, (m, i, str) => DEVA.test(str.slice(Math.max(0, i - 3), i)) ? '।' : m);
}

/* ---------------- dialog & run ---------------- */
export async function ocrDialog() {
  if (!S.doc) return;
  let pref = {}; try { pref = JSON.parse(localStorage.getItem('bts-pdf-ocr') || '{}'); } catch {}
  const langs = pref.langs || ['eng'];
  const n = S.pages.length;
  const v = await BTS.dialog({
    title: 'Recognize text (OCR)', width: 500,
    body: `<p class="dlg-p">Reads the words in scanned pages or photos, so you can search, select and copy them. Works offline.</p>
      <div class="field"><label>Pages</label>
        <div class="opt-list tight">
          <label class="check"><input type="radio" name="op" value="cur" ${n === 1 || pref.scope !== 'all' ? 'checked' : ''}> This page (${S.cur + 1})</label>
          ${n > 1 ? `<label class="check"><input type="radio" name="op" value="all" ${pref.scope === 'all' ? 'checked' : ''}> All ${n} pages</label>
          <label class="check"><input type="radio" name="op" value="range"> Pages <input id="ocr-range" class="num wide" placeholder="e.g. 1-3, 5"></label>` : ''}
        </div></div>
      <div class="field"><label>Language of the text</label>
        <div class="ocr-langs">${LANGS.map(([k, l]) => `<label class="chip"><input type="checkbox" value="${k}" ${langs.includes(k) ? 'checked' : ''}><span>${l}</span></label>`).join('')}</div></div>
      <label class="check"><input type="checkbox" id="ocr-layer" ${pref.layer === false ? '' : 'checked'}> Make the PDF searchable (adds invisible text to the pages)</label>
      <label class="check"><input type="checkbox" id="ocr-skip" ${pref.skip === false ? '' : 'checked'}> Skip pages that already have text</label>`,
    onOpen: (r) => { const ri = $('#ocr-range', r); ri?.addEventListener('focus', () => r.querySelector('input[value="range"]').checked = true); },
    buttons: [{ label: 'Cancel', value: null }, {
      label: 'Recognize text', primary: true,
      validate: (r) => { if (!$$('.ocr-langs input:checked', r).length) { BTS.toast('Choose at least one language'); return false; } return true; },
      value: (r) => ({ scope: r.querySelector('input[name="op"]:checked').value, range: $('#ocr-range', r)?.value || '', langs: $$('.ocr-langs input:checked', r).map(i => i.value), layer: $('#ocr-layer', r).checked, skip: $('#ocr-skip', r).checked }),
    }],
  });
  if (!v) return;
  try { localStorage.setItem('bts-pdf-ocr', JSON.stringify({ langs: v.langs, layer: v.layer, skip: v.skip, scope: v.scope === 'all' ? 'all' : 'cur' })); } catch {}
  let pages;
  try { pages = v.scope === 'all' ? S.pages.map(p => p.i) : v.scope === 'range' ? [...new Set(parseRanges(v.range, n).flat())] : [S.cur]; }
  catch (e) { return BTS.toast(e.message, { error: true, ms: 4000 }); }
  return runOcr(pages, v);
}

export async function runOcr(pages, { langs = ['eng'], layer = true, skip = true } = {}) {
  const h = progress('Recognizing text', 'Starting the text reader…');
  h.set(-1);
  h.onCancel = () => dropWorker();
  const results = [];   // { i, text, lines: [{ words:[{text,bbox}], bbox }], k, existing }
  const deva = langs.some(l => l !== 'eng');
  let skipped = 0;
  try {
    const order = langs.slice().sort((a, b) => (a === 'eng') - (b === 'eng'));   // Indian script first helps mixed pages
    let stage = '';
    const w = await getWorker(order.join('+'), (m) => {
      if (h.cancelled) return;
      if (m.status === 'loading language traineddata') { stage = 'Loading languages…'; h.set(m.progress, stage); }
      else if (m.status === 'recognizing text') h.set((h.done + m.progress) / h.total);
    });
    h.total = pages.length; h.done = 0;
    for (const i of pages) {
      if (h.cancelled) throw new Cancelled();
      const p = S.pages[i];
      if (skip) {
        const tc = await p.proxy.getTextContent();
        const existing = tc.items.map(it => (it.str || '') + (it.hasEOL ? '\n' : '')).join('').trim();
        if (existing.replace(/\s+/g, '').length > 25) { results.push({ i, text: existing, existing: true }); skipped++; h.done++; continue; }
      }
      h.set(h.done / h.total, `Reading page ${i + 1}${pages.length > 1 ? ` (${h.done + 1} of ${pages.length})` : ''}…`);
      const { canvas, k } = await renderForOcr(p);
      if (h.cancelled) throw new Cancelled();
      const r = await w.recognize(canvas, {}, { text: true, blocks: true });
      canvas.width = canvas.height = 0;
      results.push({ i, k, ...collect(r.data, deva) });
      h.done++;
    }
  } catch (e) {
    h.close();
    if (h.cancelled || e instanceof Cancelled) { BTS.toast('Text recognition stopped'); return; }
    console.warn(e); BTS.toast('Text recognition failed: ' + (e.message || e), { error: true, ms: 5000 }); await dropWorker(); return;
  }
  const found = results.filter(r => !r.existing && r.text.trim());
  if (layer && found.length) {
    h.set(-1, 'Adding the text to the PDF…');
    try { await addTextLayer(found); }
    catch (e) { console.warn(e); BTS.toast(e.message || String(e), { error: true, ms: 5000 }); }
  }
  h.close();
  showPanel(results);
  const words = found.reduce((s, r) => s + r.lines.reduce((a, l) => a + l.words.length, 0), 0);
  if (!found.length) BTS.toast(skipped ? 'These pages already have text — nothing new to recognize' : 'No text was found on these pages', { ms: 3500 });
  else BTS.toast(`Found ${words} word${words === 1 ? '' : 's'}${layer ? ' · the PDF is now searchable — save to keep it' : ''}`, { ms: 3800 });
  return results;
}

async function renderForOcr(p) {
  const vp1 = p.proxy.getViewport({ scale: 1, rotation: p.proxy.rotate });
  let k = 300 / 72;
  const maxSide = 4200; if (Math.max(vp1.width, vp1.height) * k > maxSide) k = maxSide / Math.max(vp1.width, vp1.height);
  const vp = p.proxy.getViewport({ scale: k, rotation: p.proxy.rotate });
  const canvas = document.createElement('canvas'); canvas.width = Math.round(vp.width); canvas.height = Math.round(vp.height);
  const ctx = canvas.getContext('2d', { alpha: false, willReadFrequently: true });
  ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
  await p.proxy.render({ canvasContext: ctx, viewport: vp, intent: 'print' }).promise;
  return { canvas, k };
}

/** Flatten tesseract blocks into lines/words and plain text (paragraphs separated by blank lines). */
function collect(data, deva) {
  const lines = []; const paras = [];
  for (const b of data.blocks || []) for (const par of b.paragraphs || []) {
    const pl = [];
    for (const ln of par.lines || []) {
      const words = (ln.words || []).filter(w => w.text && w.text.trim() && w.confidence > 15).map(w => ({ text: tidy(w.text.trim(), deva), bbox: w.bbox }));
      if (!words.length) continue;
      lines.push({ words, bbox: ln.bbox });
      pl.push(words.map(w => w.text).join(' '));
    }
    if (pl.length) paras.push(pl.join('\n'));
  }
  return { lines, text: paras.join('\n\n') };
}

/* ---------------- invisible text layer ---------------- */
/** A tiny TrueType font with two empty glyphs, so viewers never need a real font for the hidden text. */
function glyphlessTTF() {
  const T = {};
  const buf = (n, fill) => { const b = new DataView(new ArrayBuffer(n)); fill(b); return new Uint8Array(b.buffer); };
  T.head = buf(54, b => { b.setUint32(0, 0x10000); b.setUint32(4, 0x10000); b.setUint32(12, 0x5F0F3CF5); b.setUint16(16, 0x000B); b.setUint16(18, 1000); b.setInt16(40, 500); b.setInt16(42, 1000); b.setUint16(46, 3); b.setInt16(48, 2); });
  T.hhea = buf(36, b => { b.setUint32(0, 0x10000); b.setInt16(4, 1000); b.setUint16(10, 500); b.setInt16(18, 1); b.setUint16(34, 2); });
  T.maxp = buf(32, b => { b.setUint32(0, 0x10000); b.setUint16(4, 2); b.setUint16(14, 2); });
  T.hmtx = buf(8, b => { b.setUint16(0, 500); b.setUint16(4, 500); });
  T.loca = buf(6, () => {});
  T.glyf = new Uint8Array(0);
  T.cmap = buf(36, b => { b.setUint16(2, 1); b.setUint16(4, 3); b.setUint16(6, 1); b.setUint32(8, 12); b.setUint16(12, 4); b.setUint16(14, 24); b.setUint16(18, 2); b.setUint16(20, 2); b.setUint16(26, 0xFFFF); b.setUint16(30, 0xFFFF); b.setInt16(32, 1); });
  T.post = buf(32, b => { b.setUint32(0, 0x30000); b.setInt16(8, -100); b.setInt16(10, 50); });
  const tags = Object.keys(T).sort();
  const pad4 = (n) => (n + 3) & ~3;
  let off = 12 + 16 * tags.length;
  const total = tags.reduce((s, t) => s + pad4(T[t].length), off);
  const out = new Uint8Array(total); const dv = new DataView(out.buffer);
  dv.setUint32(0, 0x10000); dv.setUint16(4, tags.length); dv.setUint16(6, 128); dv.setUint16(8, 3); dv.setUint16(10, tags.length * 16 - 128);
  const sum = (bytes) => { const p = new Uint8Array(pad4(bytes.length)); p.set(bytes); const v = new DataView(p.buffer); let s = 0; for (let i = 0; i < p.length; i += 4) s = (s + v.getUint32(i)) >>> 0; return s; };
  let headOff = 0;
  tags.forEach((t, i) => {
    const r = 12 + i * 16;
    for (let c = 0; c < 4; c++) out[r + c] = t.charCodeAt(c);
    dv.setUint32(r + 4, sum(T[t])); dv.setUint32(r + 8, off); dv.setUint32(r + 12, T[t].length);
    out.set(T[t], off); if (t === 'head') headOff = off;
    off += pad4(T[t].length);
  });
  dv.setUint32(headOff + 8, (0xB1B0AFBA - sum(out)) >>> 0);
  return out;
}

function makeOcrFont(doc, L) {
  const ctx = doc.context;
  const ttf = glyphlessTTF();
  const fontFile = ctx.register(ctx.flateStream(ttf, { Length1: ttf.length }));
  const fd = ctx.register(ctx.obj({ Type: 'FontDescriptor', FontName: 'GlyphLessFont', Flags: 5, FontBBox: [0, 0, 500, 1000], ItalicAngle: 0, Ascent: 1000, Descent: 0, CapHeight: 1000, StemV: 80, FontFile2: fontFile }));
  const map = new Uint8Array(131072); for (let i = 1; i < map.length; i += 2) map[i] = 1;   // every CID -> glyph 1
  const cidToGid = ctx.register(ctx.flateStream(map));
  const cid = ctx.register(ctx.obj({ Type: 'Font', Subtype: 'CIDFontType2', BaseFont: 'GlyphLessFont', CIDSystemInfo: { Registry: L.PDFString.of('Adobe'), Ordering: L.PDFString.of('Identity'), Supplement: 0 }, FontDescriptor: fd, DW: 500, CIDToGIDMap: cidToGid }));
  const ranges = []; for (let hi = 0; hi < 256; hi++) { const h = hi.toString(16).padStart(2, '0').toUpperCase(); ranges.push(`<${h}00> <${h}FF> <${h}00>`); }
  let cm = '/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n';
  for (let i = 0; i < ranges.length; i += 100) { const chunk = ranges.slice(i, i + 100); cm += `${chunk.length} beginbfrange\n${chunk.join('\n')}\nendbfrange\n`; }
  cm += 'endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n';
  const toUni = ctx.register(ctx.flateStream(new TextEncoder().encode(cm)));
  return ctx.register(ctx.obj({ Type: 'Font', Subtype: 'Type0', BaseFont: 'GlyphLessFont', Encoding: 'Identity-H', DescendantFonts: [cid], ToUnicode: toUni }));
}
const hex16 = (s) => { let o = ''; for (const ch of s) { const c = ch.codePointAt(0); if (c <= 0xFFFF) o += c.toString(16).padStart(4, '0'); } return o.toUpperCase(); };

async function addTextLayer(found) {
  const L = await loadPdfLib();
  const base = await IO.buildBaseBytes();
  let doc;
  try { doc = await L.PDFDocument.load(base, { updateMetadata: false }); }
  catch { throw new Error('This PDF is password-protected, so the text can’t be added to it. You can still copy the text from the panel.'); }
  const fontRef = makeOcrFont(doc, L);
  const pages = doc.getPages();
  const f = (n) => +n.toFixed(3);
  for (const r of found) {
    const page = pages[r.i]; const p = S.pages[r.i]; if (!page || !p) continue;
    const vpN = p.proxy.getViewport({ scale: 1, rotation: p.proxy.rotate });
    const D = (x, y) => vpN.convertToPdfPoint(x, y);
    const name = page.node.newFontDictionary('OCR', fontRef);
    const ops = [L.pushGraphicsState(), L.beginText(), L.setTextRenderingMode(L.TextRenderingMode.Invisible)];
    for (const ln of r.lines) {
      const size = Math.max(2, (ln.bbox.y1 - ln.bbox.y0) / r.k);
      const base = ln.bbox.y1 / r.k;
      ln.words.forEach((w, j) => {
        const next = ln.words[j + 1];
        const txt = next ? w.text + ' ' : w.text;
        const x0 = w.bbox.x0 / r.k, x1 = (next ? next.bbox.x0 : w.bbox.x1) / r.k;
        const glyphs = [...txt].filter(ch => ch.codePointAt(0) <= 0xFFFF).length || 1;
        const tz = Math.max(1, (x1 - x0) / (glyphs * 0.5 * size) * 100);
        const o = D(x0, base), ux = D(x0 + 1, base), up = D(x0, base - 1);
        ops.push(L.setFontAndSize(name, f(size)), L.setCharacterSqueeze(f(tz)),
          L.setTextMatrix(f(ux[0] - o[0]), f(ux[1] - o[1]), f(up[0] - o[0]), f(up[1] - o[1]), f(o[0]), f(o[1])),
          L.showText(L.PDFHexString.of(hex16(txt))));
      });
    }
    ops.push(L.endText(), L.popGraphicsState());
    page.pushOperators(...ops);
  }
  const bytes = await doc.save({ useObjectStreams: false });
  S.modified = true;
  await APP.openBytes(bytes, { keepAnns: true, keepView: true });
  BTS.markDirty();
}

/* ---------------- results panel ---------------- */
let lastResults = null;
function showPanel(results) {
  lastResults = results;
  let panel = $('#ocr-panel');
  if (!panel) {
    panel = document.createElement('aside'); panel.id = 'ocr-panel';
    panel.innerHTML = `<header><b>${icon('scan-text', 16)}Recognized text</b><button class="rb" data-ocr="close" title="Close">${icon('x', 16)}</button></header>
      <textarea spellcheck="false" aria-label="Recognized text"></textarea>
      <footer><button class="btn" data-ocr="copy">${icon('copy', 15)}Copy</button><button class="btn" data-ocr="txt" title="Save as a text file">${icon('file-text', 15)}.txt</button><button class="btn" data-ocr="docx" title="Save as a Word document">${icon('file-type', 15)}Word</button></footer>`;
    $('#workspace').appendChild(panel);
    panel.addEventListener('click', onPanel);
    panel.querySelector('textarea').addEventListener('keydown', e => e.stopPropagation());
  }
  const multi = results.length > 1;
  panel.querySelector('textarea').value = results.map(r => (multi ? `— Page ${r.i + 1}${r.existing ? ' (already had text)' : ''} —\n` : '') + (r.text || '')).join('\n\n').trim();
  panel.classList.remove('hidden');
  document.body.classList.add('ocr-open');
  window.dispatchEvent(new Event('bts-pdf-resize'));
}
export function closePanel() { $('#ocr-panel')?.classList.add('hidden'); document.body.classList.remove('ocr-open'); window.dispatchEvent(new Event('bts-pdf-resize')); }
async function onPanel(e) {
  const b = e.target.closest('[data-ocr]'); if (!b) return;
  const text = $('#ocr-panel textarea').value;
  switch (b.dataset.ocr) {
    case 'close': return closePanel();
    case 'copy': await navigator.clipboard.writeText(text); return BTS.toast('Text copied');
    case 'txt': return IO.saveBytes(`${BTS.title} (text).txt`, '﻿' + text.replace(/\r?\n/g, '\r\n'), 'Text file', 'txt');
    case 'docx': {
      const blocks = text.split(/\n(?=— Page \d+)/).map(chunk => chunk.replace(/^— Page \d+[^\n]*—\n?/, ''));
      const bytes = await textToDocx(blocks, BTS.title);
      const path = await IO.saveBytes(`${BTS.title} (text).docx`, bytes, 'Word document', 'docx');
      if (typeof path === 'string') offerDocs(path);
    }
  }
}
async function offerDocs(path) {
  const go = await BTS.dialog({ title: 'Word file saved', body: `<p class="dlg-p">Open it in BTS Docs to edit it?</p>`, buttons: [{ label: 'Not now', value: false }, { label: 'Open in BTS Docs', primary: true, value: true }] });
  if (go) bts.openApp('docs', path);
}
export const ocrResults = () => lastResults;
