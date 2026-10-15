/* BTS PDF — Images → PDF: pick photos / scans, put them in order, choose page size and margins. */
import { S, $, $$, esc, loadPdfLib, fmtSize, APP } from './state.js';
import * as IO from './io.js';

const MM = 72 / 25.4;
const SIZES = { a4: [595.28, 841.89], letter: [612, 792] };
const MARGINS = { none: 0, small: 10 * MM, big: 20 * MM };

/** Let the user choose several pictures. Resolves to [{ name, type, data }] (possibly empty). */
export function pickImageFiles() {
  return new Promise(resolve => {
    const inp = document.createElement('input'); inp.type = 'file'; inp.multiple = true;
    inp.accept = 'image/jpeg,image/png,image/webp,image/gif,image/bmp,.jpg,.jpeg,.png,.webp,.gif,.bmp';
    inp.onchange = async () => resolve(await Promise.all([...inp.files].map(async f => ({ name: f.name, type: f.type, data: new Uint8Array(await f.arrayBuffer()) }))));
    inp.addEventListener('cancel', () => resolve([]));
    inp.click();
  });
}

/** EXIF orientation of a JPEG (1 = normal). */
function jpegOrientation(b) {
  if (b[0] !== 0xFF || b[1] !== 0xD8) return 1;
  let i = 2;
  while (i + 4 < b.length) {
    if (b[i] !== 0xFF) return 1;
    const m = b[i + 1], len = (b[i + 2] << 8) | b[i + 3];
    if (m === 0xE1 && b[i + 4] === 0x45 && b[i + 5] === 0x78) {   // "Exif"
      const t = i + 10, le = b[t] === 0x49;
      const u16 = (o) => le ? b[o] | (b[o + 1] << 8) : (b[o] << 8) | b[o + 1];
      const u32 = (o) => le ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0 : ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0;
      const ifd = t + u32(t + 4); const n = u16(ifd);
      for (let k = 0; k < n; k++) { const e = ifd + 2 + k * 12; if (u16(e) === 0x0112) return u16(e + 8) || 1; }
      return 1;
    }
    if (m === 0xDA) return 1;
    i += 2 + len;
  }
  return 1;
}

/** Turn a picked file into something pdf-lib can embed (JPEG or PNG), upright and (optionally) not huge. */
export async function prepareImage(f, shrink = true) {
  const isJpg = /jpe?g/i.test(f.type) || /\.jpe?g$/i.test(f.name) || (f.data[0] === 0xFF && f.data[1] === 0xD8);
  const isPng = /png/i.test(f.type) || /\.png$/i.test(f.name) || (f.data[0] === 0x89 && f.data[1] === 0x50);
  let bmp;
  try { bmp = await createImageBitmap(new Blob([f.data], { type: f.type || '' }), { imageOrientation: 'from-image' }); }
  catch { throw new Error(`“${f.name}” isn’t a picture BTS PDF can read`); }
  const w = bmp.width, h = bmp.height;
  const MAX = 2400;
  const big = shrink && Math.max(w, h) > MAX;
  const rotated = isJpg && jpegOrientation(f.data) !== 1;
  if (isJpg && !big && !rotated) { bmp.close(); return { bytes: f.data, kind: 'jpg', w, h }; }
  if (isPng && !big && !(shrink && f.data.length > 1.5e6)) { bmp.close(); return { bytes: f.data, kind: 'png', w, h }; }
  const k = big ? MAX / Math.max(w, h) : 1;
  const c = document.createElement('canvas'); c.width = Math.round(w * k); c.height = Math.round(h * k);
  const x = c.getContext('2d');
  const keepPng = isPng && hasAlpha(bmp) || /gif|bmp/i.test(f.type);
  if (!keepPng) { x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); }
  x.drawImage(bmp, 0, 0, c.width, c.height); bmp.close();
  const blob = await new Promise(r => c.toBlob(r, keepPng ? 'image/png' : 'image/jpeg', 0.86));
  c.width = c.height = 0;
  return { bytes: new Uint8Array(await blob.arrayBuffer()), kind: keepPng ? 'png' : 'jpg', w: Math.round(w * k), h: Math.round(h * k) };
}
function hasAlpha(bmp) {
  const c = document.createElement('canvas'); const s = 64; c.width = c.height = s;
  const x = c.getContext('2d'); x.drawImage(bmp, 0, 0, s, s);
  const d = x.getImageData(0, 0, s, s).data; for (let i = 3; i < d.length; i += 4) if (d[i] < 250) return true;
  return false;
}

/** Build a PDF from pictures. opts: { size: 'a4'|'letter'|'fit', orient: 'auto'|'portrait'|'landscape', margin: 'none'|'small'|'big', shrink } */
export async function buildImagesPdf(files, opts = {}, onProgress) {
  const L = await loadPdfLib();
  const doc = await L.PDFDocument.create();
  const m = MARGINS[opts.margin || 'small'] ?? 0;
  let i = 0;
  for (const f of files) {
    onProgress?.(i++ / files.length, f.name);
    const im = await prepareImage(f, opts.shrink !== false);
    const img = im.kind === 'jpg' ? await doc.embedJpg(im.bytes) : await doc.embedPng(im.bytes);
    let pw, ph;
    if ((opts.size || 'a4') === 'fit') {
      let w = im.w * 0.75, h = im.h * 0.75;                 // 96 dpi
      const k = Math.min(1, 842 / Math.max(w, h)); w *= k; h *= k;
      pw = w + 2 * m; ph = h + 2 * m;
    } else {
      [pw, ph] = SIZES[opts.size] || SIZES.a4;
      const land = opts.orient === 'landscape' || (opts.orient !== 'portrait' && im.w > im.h);
      if (land) [pw, ph] = [ph, pw];
    }
    const page = doc.addPage([pw, ph]);
    const aw = pw - 2 * m, ah = ph - 2 * m;
    const k = Math.min(aw / im.w, ah / im.h);
    const w = im.w * k, h = im.h * k;
    page.drawImage(img, { x: (pw - w) / 2, y: (ph - h) / 2, width: w, height: h });
  }
  doc.setTitle(files.length === 1 ? files[0].name.replace(/\.[^.]+$/, '') : 'Pictures');
  doc.setCreator('BTS PDF');
  onProgress?.(1);
  return doc.save();
}

/* ---------------- dialog ---------------- */
export async function imagesToPdfDialog(initial) {
  let files = initial || await pickImageFiles();
  if (!files.length) return;
  const urls = new Map();
  const urlOf = (f) => { if (!urls.has(f)) urls.set(f, URL.createObjectURL(new Blob([f.data], { type: f.type || 'image/jpeg' }))); return urls.get(f); };
  let pref = {}; try { pref = JSON.parse(localStorage.getItem('bts-pdf-img2pdf') || '{}'); } catch {}
  const body = document.createElement('div'); body.className = 'i2p';
  body.innerHTML = `
    <div class="i2p-list" aria-label="Pictures in page order"></div>
    <div class="i2p-opts">
      <div class="field"><label for="i2p-size">Page size</label><select id="i2p-size"><option value="a4">A4</option><option value="letter">Letter</option><option value="fit">Same as the picture</option></select></div>
      <div class="field"><label for="i2p-orient">Orientation</label><select id="i2p-orient"><option value="auto">Automatic</option><option value="portrait">Portrait</option><option value="landscape">Landscape</option></select></div>
      <div class="field"><label for="i2p-margin">Margins</label><select id="i2p-margin"><option value="none">None</option><option value="small">Small (10 mm)</option><option value="big">Big (20 mm)</option></select></div>
      <label class="check"><input type="checkbox" id="i2p-shrink" checked> Make big photos smaller (smaller file)</label>
      <p class="hint i2p-total"></p>
    </div>`;
  const list = $('.i2p-list', body);
  const render = () => {
    list.innerHTML = files.map((f, i) => `<div class="i2p-item" draggable="true" data-i="${i}" title="${esc(f.name)}"><img src="${urlOf(f)}" alt=""><span class="i2p-n">${i + 1}</span><button class="i2p-x" data-x="${i}" title="Remove">${icon('x', 13)}</button><small>${esc(f.name)}</small></div>`).join('') +
      `<button class="i2p-add" data-add>${icon('plus', 22)}<span>Add pictures</span></button>`;
    $('.i2p-total', body).textContent = `${files.length} picture${files.length === 1 ? '' : 's'} · ${fmtSize(files.reduce((s, f) => s + f.data.length, 0))} · drag to change the order`;
  };
  let dragI = -1;
  list.addEventListener('dragstart', e => { const it = e.target.closest('.i2p-item'); if (!it) return; dragI = +it.dataset.i; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', 'img'); it.classList.add('dragging'); });
  list.addEventListener('dragover', e => { if (dragI < 0) return; e.preventDefault(); $$('.i2p-item', list).forEach(x => x.classList.remove('drop-before', 'drop-after')); const it = e.target.closest('.i2p-item'); if (!it) return; const r = it.getBoundingClientRect(); it.classList.add(e.clientX > r.left + r.width / 2 ? 'drop-after' : 'drop-before'); });
  list.addEventListener('drop', e => {
    e.preventDefault(); const it = e.target.closest('.i2p-item'); if (dragI < 0 || !it) { dragI = -1; return render(); }
    const r = it.getBoundingClientRect(); let to = +it.dataset.i + (e.clientX > r.left + r.width / 2 ? 1 : 0);
    const [f] = files.splice(dragI, 1); if (to > dragI) to--; files.splice(to, 0, f); dragI = -1; render();
  });
  list.addEventListener('dragend', () => { dragI = -1; render(); });
  list.addEventListener('click', async e => {
    const x = e.target.closest('[data-x]'); if (x) { files.splice(+x.dataset.x, 1); return render(); }
    if (e.target.closest('[data-add]')) { const more = await pickImageFiles(); if (more.length) { files = files.concat(more); render(); } }
  });
  render();
  const v = await BTS.dialog({
    title: 'Create PDF from pictures', body, width: 720,
    onOpen: (r) => { $('#i2p-size', r).value = pref.size || 'a4'; $('#i2p-orient', r).value = pref.orient || 'auto'; $('#i2p-margin', r).value = pref.margin || 'small'; },
    buttons: [{ label: 'Cancel', value: null }, { label: 'Create PDF', primary: true, validate: () => { if (!files.length) { BTS.toast('Add at least one picture'); return false; } return true; }, value: (r) => ({ size: $('#i2p-size', r).value, orient: $('#i2p-orient', r).value, margin: $('#i2p-margin', r).value, shrink: $('#i2p-shrink', r).checked }) }],
  });
  urls.forEach(u => URL.revokeObjectURL(u));
  if (!v) return;
  try { localStorage.setItem('bts-pdf-img2pdf', JSON.stringify({ size: v.size, orient: v.orient, margin: v.margin })); } catch {}
  let bytes;
  try {
    BTS.toast(`Making a PDF from ${files.length} picture${files.length === 1 ? '' : 's'}…`, { ms: 1500 });
    bytes = await buildImagesPdf(files, v);
  } catch (e) { return BTS.toast(e.message || String(e), { error: true, ms: 4500 }); }
  const name = (files.length === 1 ? files[0].name.replace(/\.[^.]+$/, '') : 'Pictures') + '.pdf';
  if (!S.doc && !BTS.dirty) {
    await BTS.loadFile({ name, data: bytes });
    BTS.markDirty();
    BTS.toast(`PDF ready · ${bytes.length > 0 ? fmtSize(bytes.length) : ''} · save it with Ctrl+S`, { ms: 3500 });
  } else {
    const path = await IO.saveBytes(name, bytes, 'PDF document', 'pdf');
    if (path && typeof path === 'string') bts.openApp(null, path);
  }
}
