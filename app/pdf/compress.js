/* BTS PDF — Compress PDF: lossless clean-up ("Basic"), page pictures ("Strong") and "smaller than" targets. */
import { S, $, esc, loadPdfLib, pdfjsLib, fmtSize } from './state.js';
import * as IO from './io.js';
import { progress, Cancelled, nextFrame } from './ui.js';

const TARGETS = [[200, '200 KB'], [500, '500 KB'], [1024, '1 MB'], [2048, '2 MB']];
const QUALITY = { high: { dpi: 150, q: 0.8 }, medium: { dpi: 110, q: 0.65 }, low: { dpi: 80, q: 0.5 } };

/* ---------------- lossless ---------------- */
/** Re-save with unused objects removed, plain streams deflated and object streams. */
export async function basicCompress(bytes) {
  const L = await loadPdfLib();
  const doc = await L.PDFDocument.load(bytes, { updateMetadata: false, ignoreEncryption: false });
  const ctx = doc.context;
  // 1. drop objects nothing refers to (left over after edits, old revisions…)
  const seen = new Set(); const stack = [ctx.trailerInfo.Root, ctx.trailerInfo.Info, ctx.trailerInfo.Encrypt].filter(Boolean);
  const visit = (o) => {
    if (!o) return;
    if (o instanceof L.PDFRef) { if (!seen.has(o.tag)) { seen.add(o.tag); stack.push(o); } return; }
    if (o instanceof L.PDFDict) { for (const [, v] of o.entries()) visit(v); return; }
    if (o instanceof L.PDFArray) { for (let i = 0; i < o.size(); i++) visit(o.get(i)); return; }
    if (o instanceof L.PDFStream) visit(o.dict);
  };
  stack.forEach(r => r instanceof L.PDFRef && seen.add(r.tag));
  while (stack.length) { const r = stack.pop(); visit(r instanceof L.PDFRef ? ctx.lookup(r) : r); }
  for (const [ref] of ctx.enumerateIndirectObjects()) if (!seen.has(ref.tag)) ctx.delete(ref);
  // 2. deflate streams that are stored uncompressed
  const Filter = L.PDFName.of('Filter'), Length = L.PDFName.of('Length');
  for (const [ref, obj] of ctx.enumerateIndirectObjects()) {
    if (!(obj instanceof L.PDFRawStream) || obj.dict.get(Filter)) continue;
    const raw = obj.getContents(); if (raw.length < 64) continue;
    const dict = {}; for (const [k, v] of obj.dict.entries()) if (k !== Length) dict[k.asString().slice(1)] = v;
    const z = ctx.flateStream(raw, dict);
    if (z.getContentsSize() < raw.length) ctx.assign(ref, z);
  }
  return doc.save({ useObjectStreams: true, addDefaultPage: false });
}

/* ---------------- page pictures ---------------- */
async function rasterCompress(bytes, { dpi, q, gray = false }, h, label) {
  const L = await loadPdfLib();
  const src = await pdfjsLib.getDocument({ data: bytes.slice(), isEvalSupported: false }).promise;
  try {
    const out = await L.PDFDocument.create();
    const info = S.meta?.info || {};
    if (info.Title) out.setTitle(info.Title); if (info.Author) out.setAuthor(info.Author);
    out.setProducer('BTS PDF');
    for (let i = 1; i <= src.numPages; i++) {
      if (h?.cancelled) throw new Cancelled();
      h?.set((i - 1) / src.numPages * (h.span || 1) + (h.base || 0), null, `${label || ''}Page ${i} of ${src.numPages}`);
      const page = await src.getPage(i);
      const vp1 = page.getViewport({ scale: 1, rotation: page.rotate });
      let k = dpi / 72; const maxPx = 5000;
      if (Math.max(vp1.width, vp1.height) * k > maxPx) k = maxPx / Math.max(vp1.width, vp1.height);
      const vp = page.getViewport({ scale: k, rotation: page.rotate });
      const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(vp.width)); c.height = Math.max(1, Math.round(vp.height));
      const ctx = c.getContext('2d', { alpha: false });
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
      await page.render({ canvasContext: ctx, viewport: vp, intent: 'print' }).promise;
      if (gray) { const d = ctx.getImageData(0, 0, c.width, c.height); const a = d.data; for (let j = 0; j < a.length; j += 4) { const y = a[j] * 0.299 + a[j + 1] * 0.587 + a[j + 2] * 0.114; a[j] = a[j + 1] = a[j + 2] = y; } ctx.putImageData(d, 0, 0); }
      const blob = await new Promise(r => c.toBlob(r, 'image/jpeg', q));
      c.width = c.height = 0;
      const img = await out.embedJpg(new Uint8Array(await blob.arrayBuffer()));
      const pg = out.addPage([vp1.width, vp1.height]);
      pg.drawImage(img, { x: 0, y: 0, width: vp1.width, height: vp1.height });
      page.cleanup();
      await nextFrame();
    }
    return await out.save({ useObjectStreams: true });
  } finally { src.destroy(); }
}

/** Search picture quality/resolution for the best result under `limit` bytes. */
async function toTarget(bytes, limit, h) {
  const basic = await basicCompress(bytes);
  if (basic.length <= limit) return { bytes: basic, lossless: true };
  const lvl = (t) => ({ dpi: Math.round(170 - 115 * t), q: +(0.82 - 0.45 * t).toFixed(2) });
  let best = null, smallest = null;
  const tryLevel = async (t, n, gray = false) => {
    h.base = 0; h.span = 1;
    const out = await rasterCompress(bytes, { ...lvl(t), gray }, h, `Try ${n} · `);
    if (!smallest || out.length < smallest.length) smallest = out;
    return out;
  };
  let n = 1;
  let r = await tryLevel(0, n++);
  if (r.length <= limit) return { bytes: r, level: lvl(0) };
  let lo = 0, hi = 1;
  r = await tryLevel(1, n++);
  if (r.length > limit) {
    r = await tryLevel(1, n++, true);    // last resort: grey pages
    if (r.length > limit) return { bytes: smallest, missed: true, level: lvl(1) };
    return { bytes: r, level: lvl(1), gray: true };
  }
  best = { bytes: r, level: lvl(1) };
  for (let i = 0; i < 4; i++) {
    const mid = (lo + hi) / 2;
    r = await tryLevel(mid, n++);
    if (r.length <= limit) { best = { bytes: r, level: lvl(mid) }; hi = mid; } else lo = mid;
  }
  return best;
}

/* ---------------- dialog ---------------- */
export async function compressDialog() {
  if (!S.doc) return;
  let src;
  try { src = await IO.buildAnnotatedPdf(); } catch (e) { return BTS.toast(e.message || String(e), { error: true }); }
  const before = src.length;
  let pref = {}; try { pref = JSON.parse(localStorage.getItem('bts-pdf-compress') || '{}'); } catch {}
  const v = await BTS.dialog({
    title: 'Compress PDF', width: 500,
    body: `<p class="dlg-p">Make this PDF smaller — handy for email and upload limits. Now: <b>${fmtSize(before)}</b>. Your PDF stays as it is; the smaller copy is saved as a new file.</p>
      <div class="opt-list cmp">
        <label class="opt"><input type="radio" name="cm" value="basic"><span><b>Basic</b><small>Keeps everything sharp and selectable. Saves a little.</small></span></label>
        <label class="opt"><input type="radio" name="cm" value="strong"><span><b>Strong</b><small>Turns every page into a picture. Much smaller, but <b>text can’t be selected or searched</b> any more.</small>
          <select id="cm-q" class="rsel"><option value="high">High quality</option><option value="medium">Medium quality</option><option value="low">Low quality (smallest)</option></select></span></label>
        <label class="opt"><input type="radio" name="cm" value="target"><span><b>Make it smaller than</b><small>Tries the best quality that fits — for portals with a size limit.</small>
          <select id="cm-t" class="rsel">${TARGETS.map(([k, l]) => `<option value="${k}">${l}</option>`).join('')}</select></span></label>
      </div>`,
    onOpen: (r) => {
      const m = pref.mode || 'basic'; r.querySelector(`input[value="${m}"]`).checked = true;
      if (pref.q) $('#cm-q', r).value = pref.q; if (pref.t) $('#cm-t', r).value = pref.t;
      $('#cm-q', r).addEventListener('focus', () => r.querySelector('input[value="strong"]').checked = true);
      $('#cm-t', r).addEventListener('focus', () => r.querySelector('input[value="target"]').checked = true);
    },
    buttons: [{ label: 'Cancel', value: null }, { label: 'Compress', primary: true, value: (r) => ({ mode: r.querySelector('input[name="cm"]:checked').value, q: $('#cm-q', r).value, t: +$('#cm-t', r).value }) }],
  });
  if (!v) return;
  try { localStorage.setItem('bts-pdf-compress', JSON.stringify(v)); } catch {}
  const h = progress('Compressing PDF', v.mode === 'basic' ? 'Cleaning up the file…' : v.mode === 'strong' ? 'Turning pages into pictures…' : `Finding the best quality under ${TARGETS.find(t => t[0] === v.t)[1]}…`);
  h.set(v.mode === 'basic' ? -1 : 0);
  let res;
  try {
    if (v.mode === 'basic') { await nextFrame(); res = { bytes: await basicCompress(src), lossless: true }; }
    else if (v.mode === 'strong') res = { bytes: await rasterCompress(src, QUALITY[v.q], h) };
    else res = await toTarget(src, v.t * 1024, h);
  } catch (e) {
    h.close();
    if (e instanceof Cancelled) return BTS.toast('Compression stopped');
    if (/encrypt/i.test(e.message)) return BTS.toast('This PDF is password-protected, so it can’t be compressed.', { error: true });
    console.warn(e); return BTS.toast(e.message || String(e), { error: true });
  }
  h.close();
  const after = res.bytes.length;
  const pctLess = Math.round((1 - after / before) * 100);
  let note = '';
  if (after >= before) note = 'This PDF is already well compressed — the copy isn’t smaller.';
  else if (res.missed) note = `It couldn’t get under ${TARGETS.find(t => t[0] === v.t)[1]} without becoming unreadable. This is the smallest it can go.`;
  else if (res.gray) note = 'Pages were made grey to fit the limit.';
  else if (!res.lossless) note = 'Text in the copy can’t be selected or searched.';
  const save = await BTS.dialog({
    title: after < before ? 'Smaller PDF ready' : 'No savings',
    body: `<div class="cmp-res"><div><small>Before</small><b>${fmtSize(before)}</b></div>${icon('arrow-right', 22)}<div class="${after < before ? 'good' : ''}"><small>After</small><b>${fmtSize(after)}</b></div>${after < before ? `<span class="cmp-pct">${pctLess}% smaller</span>` : ''}</div>${note ? `<p class="dlg-p">${esc(note)}</p>` : ''}`,
    buttons: [{ label: 'Close', value: false }, ...(after < before || res.missed ? [{ label: 'Save as new file…', primary: true, value: true }] : [])],
  });
  if (save) await IO.saveBytes(`${BTS.title} (compressed).pdf`, res.bytes, 'PDF document', 'pdf');
}
