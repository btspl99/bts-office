/* BTS PDF — Fill & Sign: signatures (draw / type / photo), pictures, ticks, crosses and dates.
 * Pictures become 'image' annotations (see annotations.js) that can be moved, resized and deleted,
 * and are written into the PDF by io.js. Placement: pick → click on the page. */
import { S, $, $$, esc, clamp, newId, clientToPdf, APP } from './state.js';
import * as A from './annotations.js';
import { nativeVp } from './annotations.js';

const SIG_KEY = 'bts-pdf-signature';
const loadSaved = () => { try { return JSON.parse(localStorage.getItem(SIG_KEY) || 'null'); } catch { return null; } };

/* ---------------- picture helpers ---------------- */
function addPicture(url, w, h, mime) {
  const id = 'img' + newId();
  S.images.set(id, { url, w, h, mime });
  return id;
}
const canvasToURL = (c, mime = 'image/png', q = 0.9) => c.toDataURL(mime, q);
function loadImg(src) { return new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('This picture couldn’t be read')); i.src = src; }); }

/** Crop a canvas to its non-transparent pixels (+ padding). */
function trimCanvas(c, pad = 6) {
  const x = c.getContext('2d'); const { width: W, height: H } = c;
  const d = x.getImageData(0, 0, W, H).data;
  let l = W, t = H, r = -1, b = -1;
  for (let y = 0; y < H; y++) for (let i = 0; i < W; i++) if (d[(y * W + i) * 4 + 3] > 8) { if (i < l) l = i; if (i > r) r = i; if (y < t) t = y; if (y > b) b = y; }
  if (r < 0) return null;
  l = Math.max(0, l - pad); t = Math.max(0, t - pad); r = Math.min(W - 1, r + pad); b = Math.min(H - 1, b + pad);
  const o = document.createElement('canvas'); o.width = r - l + 1; o.height = b - t + 1;
  o.getContext('2d').drawImage(c, l, t, o.width, o.height, 0, 0, o.width, o.height);
  return o;
}
/** Make white / light paper transparent (photo of a signature on paper). */
function removeWhite(c, strength = 0.5) {
  const x = c.getContext('2d'); const img = x.getImageData(0, 0, c.width, c.height); const d = img.data;
  // estimate paper brightness from the brightest 10 % of pixels
  const hist = new Uint32Array(256); for (let i = 0; i < d.length; i += 4) hist[(d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114) | 0]++;
  let acc = 0, paper = 255; const n = d.length / 4; for (let v = 255; v >= 0; v--) { acc += hist[v]; if (acc > n * 0.1) { paper = v; break; } }
  const hi = paper - 18 + strength * 10, lo = hi - 50 - strength * 40;
  for (let i = 0; i < d.length; i += 4) {
    const L = d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114;
    const a = L >= hi ? 0 : L <= lo ? 1 : (hi - L) / (hi - lo);
    d[i + 3] = Math.round(d[i + 3] * a);
    if (a > 0 && a < 1) { const k = 0.35; d[i] *= k + (1 - k) * a; d[i + 1] *= k + (1 - k) * a; d[i + 2] *= k + (1 - k) * a; }   // darken soft edges
  }
  x.putImageData(img, 0, 0);
  return c;
}
async function fileToCanvas(bytes, mime, maxSide = 1800) {
  const blob = new Blob([bytes], { type: mime || 'image/*' });
  let bmp;
  try { bmp = await createImageBitmap(blob, { imageOrientation: 'from-image' }); }
  catch { const u = URL.createObjectURL(blob); try { bmp = await loadImg(u); } finally { setTimeout(() => URL.revokeObjectURL(u), 1000); } }
  const k = Math.min(1, maxSide / Math.max(bmp.width, bmp.height));
  const c = document.createElement('canvas'); c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  bmp.close?.();
  return c;
}

/* ---------------- signature dialog ---------------- */
const INKS = [['#14171f', 'Black'], ['#1a3fb0', 'Blue']];
const TYPE_FONTS = [['BTS Kalam', 400, 'Handwriting'], ['BTS Kalam', 700, 'Bold handwriting'], ['BTS Deva', 400, 'Plain']];

export async function signatureDialog() {
  const saved = loadSaved();
  let ink = INKS[0][0], tab = saved ? 'saved' : 'draw', upload = null, typeFont = 0;
  const body = document.createElement('div'); body.className = 'sig-dlg';
  body.innerHTML = `
    <div class="seg sig-tabs" role="tablist">
      ${saved ? '<button data-tab="saved">Saved</button>' : ''}
      <button data-tab="draw">${icon('pen-line', 15)}Draw</button>
      <button data-tab="type">${icon('type', 15)}Type</button>
      <button data-tab="upload">${icon('image-up', 15)}Photo</button>
    </div>
    <div class="sig-pane" data-pane="saved">${saved ? `<div class="sig-box sig-saved"><img src="${saved.url}" alt="Saved signature"></div><button class="linkbtn" data-act="forget">${icon('trash-2', 14)}Forget saved signature</button>` : ''}</div>
    <div class="sig-pane" data-pane="draw">
      <div class="sig-box"><canvas class="sig-pad" width="1040" height="360"></canvas><span class="sig-line"></span><span class="sig-ph">Sign here with your mouse, pen or finger</span></div>
      <div class="sig-row"><div class="sig-inks">${INKS.map(([c, n]) => `<button class="sig-ink" data-ink="${c}" style="--c:${c}" title="${n}"></button>`).join('')}</div><button class="btn" data-act="clear">${icon('eraser', 15)}Clear</button></div>
    </div>
    <div class="sig-pane" data-pane="type">
      <div class="field"><label for="sig-name">Your name</label><input id="sig-name" maxlength="60" autocomplete="name" placeholder="Type your name" value="${esc(BTS.userName || '')}"></div>
      <div class="sig-fonts">${TYPE_FONTS.map(([f, w, n], i) => `<button class="sig-font" data-font="${i}" title="${n}" style="font-family:'${f}';font-weight:${w}"><span>${esc(BTS.userName || 'Your name')}</span></button>`).join('')}</div>
      <div class="sig-row"><div class="sig-inks">${INKS.map(([c, n]) => `<button class="sig-ink" data-ink="${c}" style="--c:${c}" title="${n}"></button>`).join('')}</div></div>
    </div>
    <div class="sig-pane" data-pane="upload">
      <div class="sig-box sig-up checker"><button class="btn" data-act="pick">${icon('upload', 15)}Choose a photo of your signature</button><canvas class="sig-prev hidden"></canvas></div>
      <label class="check"><input type="checkbox" id="sig-white" checked> Remove white background</label>
      <p class="hint">Tip: sign on plain white paper with a dark pen, take a sharp photo, and crop it close.</p>
    </div>
    <label class="check sig-remember"><input type="checkbox" id="sig-remember" checked> Remember this signature on this computer</label>`;

  // drawing pad
  const pad = $('.sig-pad', body); const px = pad.getContext('2d');
  let strokes = [], cur = null;
  const redraw = () => {
    px.clearRect(0, 0, pad.width, pad.height);
    for (const s of strokes) drawStroke(px, s);
    $('.sig-ph', body).classList.toggle('hidden', strokes.length > 0);
  };
  const ptOf = (e) => { const r = pad.getBoundingClientRect(); return [(e.clientX - r.left) / r.width * pad.width, (e.clientY - r.top) / r.height * pad.height, e.pointerType === 'pen' ? (e.pressure || 0.5) : 0.5]; };
  pad.addEventListener('pointerdown', e => { e.preventDefault(); try { pad.setPointerCapture(e.pointerId); } catch {} cur = { color: ink, pts: [ptOf(e)] }; strokes.push(cur); redraw(); });
  pad.addEventListener('pointermove', e => { if (!cur) return; const evs = e.getCoalescedEvents?.() || []; for (const ce of (evs.length ? evs : [e])) cur.pts.push(ptOf(ce)); redraw(); });
  const end = () => { cur = null; };
  pad.addEventListener('pointerup', end); pad.addEventListener('pointercancel', end);

  const nameInp = $('#sig-name', body);
  const syncType = () => $$('.sig-font span', body).forEach(s => s.textContent = nameInp.value.trim() || 'Your name');
  nameInp.addEventListener('input', syncType);

  const setTab = (t) => {
    tab = t;
    $$('.sig-tabs [data-tab]', body).forEach(b => b.classList.toggle('active', b.dataset.tab === t));
    $$('.sig-pane', body).forEach(p => p.classList.toggle('active', p.dataset.pane === t));
    $('.sig-remember', body).classList.toggle('hidden', t === 'saved');
    if (t === 'type') setTimeout(() => nameInp.focus(), 0);
  };
  const setInk = (c) => {
    ink = c;
    $$('.sig-ink', body).forEach(b => b.classList.toggle('on', b.dataset.ink === c));
    strokes.forEach(s => s.color = c); redraw();
    $$('.sig-font', body).forEach(b => b.style.color = c);
  };
  const setFont = (i) => { typeFont = i; $$('.sig-font', body).forEach(b => b.classList.toggle('on', +b.dataset.font === i)); };
  const renderUpload = () => {
    const prev = $('.sig-prev', body); if (!upload) return;
    prev.width = upload.width; prev.height = upload.height;
    const x = prev.getContext('2d'); x.clearRect(0, 0, prev.width, prev.height); x.drawImage(upload, 0, 0);
    if ($('#sig-white', body).checked) removeWhite(prev);
    prev.classList.remove('hidden'); $('[data-act="pick"]', body).classList.add('small');
  };
  body.addEventListener('click', async (e) => {
    const t = e.target.closest('[data-tab]'); if (t) return setTab(t.dataset.tab);
    const k = e.target.closest('[data-ink]'); if (k) return setInk(k.dataset.ink);
    const f = e.target.closest('[data-font]'); if (f) return setFont(+f.dataset.font);
    const a = e.target.closest('[data-act]'); if (!a) return;
    if (a.dataset.act === 'clear') { strokes = []; redraw(); }
    if (a.dataset.act === 'forget') { try { localStorage.removeItem(SIG_KEY); } catch {} $('[data-tab="saved"]', body)?.remove(); setTab('draw'); BTS.toast('Saved signature removed'); }
    if (a.dataset.act === 'pick') {
      const file = await BTS.pickFile('image/*'); if (!file) return;
      try { upload = await fileToCanvas(file.data, file.type, 1400); renderUpload(); } catch (err) { BTS.toast(err.message, { error: true }); }
    }
  });
  $('#sig-white', body).addEventListener('change', renderUpload);

  const produce = async () => {
    if (tab === 'saved') return saved ? { ...saved, saved: true } : null;
    let c = null;
    if (tab === 'draw') {
      if (!strokes.length) { BTS.toast('Draw your signature first'); return undefined; }
      const o = document.createElement('canvas'); o.width = pad.width; o.height = pad.height;
      const x = o.getContext('2d'); strokes.forEach(s => drawStroke(x, s)); c = trimCanvas(o, 8);
    } else if (tab === 'type') {
      const name = nameInp.value.trim(); if (!name) { BTS.toast('Type your name first'); nameInp.focus(); return undefined; }
      const [fam, w] = TYPE_FONTS[typeFont];
      try { await document.fonts.load(`${w} 120px "${fam}"`, name); } catch {}
      const o = document.createElement('canvas'); const x = o.getContext('2d');
      x.font = `${w} 120px "${fam}", cursive`; o.width = Math.ceil(x.measureText(name).width + 60); o.height = 220;
      x.font = `${w} 120px "${fam}", cursive`; x.fillStyle = ink; x.textBaseline = 'middle'; x.fillText(name, 30, 112);
      c = trimCanvas(o, 6);
    } else if (tab === 'upload') {
      if (!upload) { BTS.toast('Choose a photo first'); return undefined; }
      const o = document.createElement('canvas'); o.width = upload.width; o.height = upload.height;
      o.getContext('2d').drawImage(upload, 0, 0);
      if ($('#sig-white', body).checked) { removeWhite(o); c = trimCanvas(o, 6) || o; } else c = o;
    }
    if (!c) { BTS.toast('Nothing to use yet'); return undefined; }
    // keep the saved picture a sensible size
    const k = Math.min(1, 900 / c.width);
    if (k < 1) { const o = document.createElement('canvas'); o.width = Math.round(c.width * k); o.height = Math.round(c.height * k); o.getContext('2d').drawImage(c, 0, 0, o.width, o.height); c = o; }
    const sig = { url: canvasToURL(c), w: c.width, h: c.height };
    if ($('#sig-remember', body).checked) { try { localStorage.setItem(SIG_KEY, JSON.stringify(sig)); } catch {} }
    return sig;
  };

  let result;
  const r = await BTS.dialog({
    title: 'Your signature', body, width: 600,
    buttons: [{ label: 'Cancel', value: null }, { label: 'Use signature', primary: true, value: 'ok', validate: () => true }],
    onOpen: (root) => {
      root.querySelector('.modal').classList.add('modal-noenter');
      setTab(tab); setInk(ink); setFont(0); syncType();
      // validate asynchronously: intercept the primary button
      const okBtn = root.querySelector('footer .primary');
      okBtn.addEventListener('click', async (e) => {
        if (result !== undefined) return;
        e.stopImmediatePropagation();
        const v = await produce();
        if (v === undefined) return;
        result = v; okBtn.click();
      }, true);
    },
  });
  return r === 'ok' ? result : null;
}

function drawStroke(x, s) {
  const p = s.pts; if (!p.length) return;
  x.strokeStyle = x.fillStyle = s.color; x.lineCap = 'round'; x.lineJoin = 'round';
  const wOf = (pr) => 3.2 + pr * 4.5;
  if (p.length === 1) { x.beginPath(); x.arc(p[0][0], p[0][1], wOf(p[0][2]) / 2, 0, Math.PI * 2); x.fill(); return; }
  for (let i = 1; i < p.length; i++) {
    const a = p[i - 1], b = p[i];
    const m0 = i > 1 ? [(p[i - 2][0] + a[0]) / 2, (p[i - 2][1] + a[1]) / 2] : a;
    const m1 = [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2];
    x.lineWidth = wOf((a[2] + b[2]) / 2);
    x.beginPath(); x.moveTo(m0[0], m0[1]); x.quadraticCurveTo(a[0], a[1], m1[0], m1[1]); x.stroke();
  }
  const l = p[p.length - 1], m = [(p[p.length - 2][0] + l[0]) / 2, (p[p.length - 2][1] + l[1]) / 2];
  x.beginPath(); x.moveTo(m[0], m[1]); x.lineTo(l[0], l[1]); x.stroke();
}

/* ---------------- placing ---------------- */
let ghost = null, tipShown = false;
function startPlacing(stamp, hint) {
  S.stamp = stamp;
  APP.setTool('stamp', hint);
  ghost?.remove(); ghost = null;
  if (stamp.kind === 'image') {
    ghost = document.createElement('div'); ghost.className = 'stamp-ghost hidden';
    ghost.innerHTML = `<img src="${S.images.get(stamp.img).url}" alt="">`;
    document.body.appendChild(ghost);
  }
}
export function stopPlacing() { S.stamp = null; ghost?.remove(); ghost = null; }

/** Put picture `img` on page p, centred at display-frame point (dx, dy) with size w×h points. */
function placeImage(p, img, dx, dy, w, h) {
  const vpN = nativeVp(p);
  dx = clamp(dx, w / 2, vpN.width - w / 2); dy = clamp(dy, h / 2, vpN.height - h / 2);
  const P = vpN.convertToPdfPoint(dx - w / 2, dy - h / 2), Q = vpN.convertToPdfPoint(dx + w / 2, dy + h / 2);
  const a = A.addAnn({ p: p.i, type: 'image', img, r: [Math.min(P[0], Q[0]), Math.min(P[1], Q[1]), Math.max(P[0], Q[0]), Math.max(P[1], Q[1])].map(n => +n.toFixed(2)) });
  A.select(a.id);
  return a;
}

export function initPlacing() {
  const host = $('#pages');
  host.addEventListener('pointerdown', (e) => {
    if (S.tool !== 'stamp' || !S.stamp || e.button !== 0) return;
    const el = e.target.closest('.page'); if (!el) return;
    const p = S.pages[+el.dataset.i]; if (!p) return;
    e.preventDefault(); e.stopPropagation();
    const [dx, dy] = nativeVp(p).convertToViewportPoint(...clientToPdf(p, e.clientX, e.clientY));
    const st = S.stamp;
    if (st.kind === 'image') placeImage(p, st.img, dx, dy, st.w, st.h);
    else if (st.kind === 'text') {
      const P = clientToPdf(p, e.clientX, e.clientY).map(n => +n.toFixed(2));
      const size = 11;
      const vpN = nativeVp(p); const [ux, uy] = vpN.convertToPdfPoint(dx, dy - size * 0.75);   // baseline near the click
      const a = A.addAnn({ p: p.i, type: 'textbox', x: +ux.toFixed(2), y: +uy.toFixed(2), text: st.text, color: S.colors.textbox, size });
      A.select(a.id); void P;
    }
    if (!st.repeat) {
      stopPlacing(); APP.setTool('text');
      if (st.kind === 'image' && !tipShown) { tipShown = true; BTS.toast('Drag to move it · drag a corner to resize · Delete removes it', { ms: 3200 }); }
    }
  }, true);
  host.addEventListener('pointermove', (e) => {
    if (!ghost || S.tool !== 'stamp') return;
    const el = e.target.closest('.page');
    if (!el) { ghost.classList.add('hidden'); return; }
    const w = S.stamp.w * S.scale, h = S.stamp.h * S.scale;
    Object.assign(ghost.style, { left: e.clientX - w / 2 + 'px', top: e.clientY - h / 2 + 'px', width: w + 'px', height: h + 'px' });
    ghost.classList.remove('hidden');
  });
  host.addEventListener('pointerleave', () => ghost?.classList.add('hidden'));
  window.addEventListener('bts-pdf-sign-field', (e) => signField(e.detail));
}

/* ---------------- commands ---------------- */
export async function sign() {
  if (!S.doc) return;
  const sig = await signatureDialog(); if (!sig) return;
  const img = addPicture(sig.url, sig.w, sig.h, 'image/png');
  let w = 150, h = w * sig.h / sig.w;
  if (h > 60) { h = 60; w = h * sig.w / sig.h; }
  startPlacing({ kind: 'image', img, w, h }, 'Click where you want to sign · Esc to cancel');
}

async function signField({ page, rect }) {
  const p = S.pages[page]; if (!p) return;
  const sig = await signatureDialog(); if (!sig) return;
  const img = addPicture(sig.url, sig.w, sig.h, 'image/png');
  const vpN = nativeVp(p);
  const a = vpN.convertToViewportPoint(rect[0], rect[1]), b = vpN.convertToViewportPoint(rect[2], rect[3]);
  const bw = Math.abs(a[0] - b[0]) - 4, bh = Math.abs(a[1] - b[1]) - 4;
  let w = bw, h = w * sig.h / sig.w; if (h > bh) { h = bh; w = h * sig.w / sig.h; }
  placeImage(p, img, (a[0] + b[0]) / 2, (a[1] + b[1]) / 2, w, h);
}

export async function addImage() {
  if (!S.doc) return;
  const f = await BTS.pickFile('image/png,image/jpeg,image/webp,image/gif,image/bmp');
  if (!f) return;
  let c;
  try { c = await fileToCanvas(f.data, f.type, 1600); } catch (e) { return BTS.toast(e.message, { error: true }); }
  const isPng = /png|gif/i.test(f.type) || /\.(png|gif)$/i.test(f.name);
  const mime = isPng ? 'image/png' : 'image/jpeg';
  const img = addPicture(canvasToURL(c, mime, 0.9), c.width, c.height, mime);
  const ratio = c.width / c.height;
  let w, h;
  if (ratio > 0.66 && ratio < 0.9) { w = 35 / 25.4 * 72; h = w / ratio; }      // looks like a passport photo: 35 mm wide
  else { w = Math.min(200, c.width * 0.75); h = w / ratio; if (h > 220) { h = 220; w = h * ratio; } }
  startPlacing({ kind: 'image', img, w, h }, 'Click where you want the picture · Esc to cancel');
}

function markPicture(kind) {
  const c = document.createElement('canvas'); c.width = c.height = 96;
  const x = c.getContext('2d'); x.strokeStyle = S.colors.textbox || '#172033'; x.lineWidth = 13; x.lineCap = 'round'; x.lineJoin = 'round';
  x.beginPath();
  if (kind === 'tick') { x.moveTo(14, 52); x.lineTo(38, 78); x.lineTo(84, 18); }
  else { x.moveTo(18, 18); x.lineTo(78, 78); x.moveTo(78, 18); x.lineTo(18, 78); }
  x.stroke();
  return c;
}
export function addMark(kind) {
  if (!S.doc) return;
  const c = markPicture(kind);
  const img = addPicture(canvasToURL(c), c.width, c.height, 'image/png');
  startPlacing({ kind: 'image', img, w: 11, h: 11, repeat: true }, `Click on each box to ${kind === 'tick' ? 'tick it' : 'cross it'} · Esc when done`);
}
export function addDate() {
  if (!S.doc) return;
  const d = new Date(); const text = `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
  startPlacing({ kind: 'text', text }, `Click where you want today’s date (${text}) · Esc to cancel`);
}
