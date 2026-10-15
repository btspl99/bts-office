// BTS Office — offline voice typing: downloads the Vosk speech models once (on first use per language)
// and keeps them in userData/voice-models/<name>/model.tar.gz (the format vosk-browser loads).
// Recognition itself runs fully offline in the page (app/common/voice.js + app/vendor/vosk).
//
// Handlers (renderer: await bts.ext(name, payload)):
//   voice-models            -> { models:[{ lang, name, label, size, downloaded, bytes, partial }], bytes }
//   voice-download  lang    -> { ok:true } | { ok:false, error, cancelled? }   (progress: bts.onExt('voice-progress', {lang, phase, received, total}))
//   voice-cancel    lang    -> { ok:true }
//   voice-model     lang    -> Uint8Array (model.tar.gz) | null
//   voice-delete    lang?   -> { ok:true, freed }   (no lang = remove all voice models)
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const { promisify } = require('util');
const inflateRaw = promisify(zlib.inflateRaw), gzip = promisify(zlib.gzip);

const MODELS = {
  'en-in': { name: 'vosk-model-small-en-in-0.4', label: 'English (India)', size: 37573330 },
  hi: { name: 'vosk-model-small-hi-0.22', label: 'हिंदी (Hindi)', size: 44458845 },
};
const BASE_URL = 'https://alphacephei.com/vosk/models/';
let DIR = null;
const jobs = new Map();   // lang -> { promise, req, cancelled, sender }

const modelFile = (m) => path.join(DIR, m.name, 'model.tar.gz');
const partFile = (m) => path.join(DIR, m.name + '.zip.part');
const fsize = (f) => { try { return fs.statSync(f).size; } catch { return 0; } };
function dirSize(d) {
  let n = 0;
  try { for (const f of fs.readdirSync(d, { withFileTypes: true })) n += f.isDirectory() ? dirSize(path.join(d, f.name)) : fsize(path.join(d, f.name)); } catch { }
  return n;
}
const friendly = (msg) => /INTERNET_DISCONNECTED|NAME_NOT_RESOLVED|ADDRESS_UNREACHABLE|NETWORK_CHANGED|CONNECTION_(RESET|REFUSED|CLOSED|ABORTED|TIMED_OUT|FAILED)|TIMED_OUT|ENOTFOUND|EAI_AGAIN|ECONN|ENETUNREACH|PROXY/i.test(msg) ? 'No internet connection. Connect to the internet and try again — the part already downloaded is kept.' : msg;

/* ---------- download (resumable, via Chromium's network stack so system proxy settings apply) ---------- */
function download(lang, job) {
  const { net } = require('electron');
  const m = MODELS[lang], part = partFile(m);
  fs.mkdirSync(DIR, { recursive: true });
  return new Promise((resolve, reject) => {
    let start = fsize(part);
    const req = net.request({ url: BASE_URL + m.name + '.zip', redirect: 'follow' });
    job.req = req;
    if (start) req.setHeader('Range', 'bytes=' + start + '-');
    let stall = null, out = null, done = false;
    const fail = (err) => { if (done) return; done = true; clearTimeout(stall); try { req.abort(); } catch { } if (out) out.end(); reject(err); };
    const kick = () => { clearTimeout(stall); stall = setTimeout(() => fail(new Error('The download stopped responding (ERR_TIMED_OUT).')), 30000); };
    kick();
    req.on('response', (res) => {
      const code = res.statusCode;
      if (code === 416 && start) { done = true; clearTimeout(stall); res.on('data', () => { }); return resolve(); } // already complete
      if (code !== 200 && code !== 206) return fail(new Error('The download server answered ' + code + '. Please try again later.'));
      if (code === 200) start = 0;          // server ignored Range: start over
      const len = +[].concat(res.headers['content-length'] || [0])[0];
      const cr = String([].concat(res.headers['content-range'] || [''])[0]).match(/\/(\d+)/);
      job.total = cr ? +cr[1] : start + len || m.size;
      let received = start, last = 0;
      out = fs.createWriteStream(part, { flags: start ? 'a' : 'w' });
      out.on('error', fail);
      const send = (force) => { const t = Date.now(); if (!force && t - last < 200) return; last = t; job.emit({ phase: 'download', received, total: job.total }); };
      send(true);
      res.on('data', (c) => { kick(); received += c.length; out.write(c); send(); });
      res.on('end', () => { if (done) return; clearTimeout(stall); send(true); out.end(() => { done = true; resolve(); }); });
      res.on('aborted', () => fail(new Error(job.cancelled ? 'cancelled' : 'The download was interrupted (ERR_CONNECTION_ABORTED).')));
      res.on('error', (e) => fail(e));
    });
    req.on('error', (e) => fail(e));
    req.on('abort', () => fail(new Error('cancelled')));
    req.end();
  });
}

/* ---------- zip -> tar.gz (Node has no unzip; read the central directory and inflate each entry) ---------- */
async function unzip(buf, onEntry) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 70000); i--) if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('The downloaded file is damaged (not a zip).');
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  if (p === 0xffffffff) throw new Error('Unsupported zip file.');
  for (let i = 0; i < count; i++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error('The downloaded file is damaged.');
    const method = buf.readUInt16LE(p + 10), csize = buf.readUInt32LE(p + 20), usize = buf.readUInt32LE(p + 24);
    const nl = buf.readUInt16LE(p + 28), el = buf.readUInt16LE(p + 30), cl = buf.readUInt16LE(p + 32), lo = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, p + 46 + nl);
    p += 46 + nl + el + cl;
    const ds = lo + 30 + buf.readUInt16LE(lo + 26) + buf.readUInt16LE(lo + 28);
    const raw = buf.subarray(ds, ds + csize);
    if (name.endsWith('/')) { await onEntry(name, null); continue; }
    const data = method === 0 ? raw : method === 8 ? await inflateRaw(raw) : null;
    if (!data) throw new Error('Unsupported compression in zip.');
    if (data.length !== usize) throw new Error('The downloaded file is damaged.');
    await onEntry(name, data);
  }
}
function tarHeader(name, size, dir) {
  const h = Buffer.alloc(512);
  let prefix = '';
  if (Buffer.byteLength(name) > 100) { const i = name.lastIndexOf('/', name.length - 2); prefix = name.slice(0, i); name = name.slice(i + 1); }
  h.write(name, 0, 100); h.write(dir ? '0000755\0' : '0000644\0', 100); h.write('0000000\0', 108); h.write('0000000\0', 116);
  h.write(size.toString(8).padStart(11, '0') + '\0', 124); h.write(Math.floor(Date.now() / 1000).toString(8).padStart(11, '0') + '\0', 136);
  h.write('        ', 148); h.write(dir ? '5' : '0', 156); h.write('ustar\0', 257); h.write('00', 263); h.write(prefix, 345, 155);
  let sum = 0; for (let i = 0; i < 512; i++) sum += h[i];
  h.write(sum.toString(8).padStart(6, '0') + '\0 ', 148);
  return h;
}
async function zipToTarGz(zipPath) {
  const parts = [];
  await unzip(fs.readFileSync(zipPath), async (name, data) => {
    parts.push(tarHeader(name, data ? data.length : 0, !data));
    if (data) { parts.push(data); if (data.length % 512) parts.push(Buffer.alloc(512 - data.length % 512)); }
  });
  parts.push(Buffer.alloc(1024));
  return gzip(Buffer.concat(parts), { level: 6 });
}

async function install(lang, job) {
  const m = MODELS[lang];
  await download(lang, job);
  if (job.cancelled) throw new Error('cancelled');
  const part = partFile(m), got = fsize(part);
  if (job.total && got !== job.total) { if (got > job.total) fs.rmSync(part, { force: true }); throw new Error('The download is incomplete. Please try again.'); }
  job.emit({ phase: 'extract', received: got, total: got });
  let tgz;
  try { tgz = await zipToTarGz(part); }
  catch (e) { fs.rmSync(part, { force: true }); throw e; }   // damaged: start fresh next time
  const dir = path.join(DIR, m.name);
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(modelFile(m) + '.tmp', tgz);
  fs.renameSync(modelFile(m) + '.tmp', modelFile(m));
  fs.writeFileSync(path.join(dir, 'info.json'), JSON.stringify({ name: m.name, lang, zipBytes: got, installed: new Date().toISOString() }));
  fs.rmSync(part, { force: true });
  job.emit({ phase: 'done', received: got, total: got });
}

function list() {
  const models = Object.entries(MODELS).map(([lang, m]) => {
    const downloaded = fs.existsSync(modelFile(m));
    return { lang, name: m.name, label: m.label, size: m.size, downloaded, bytes: dirSize(path.join(DIR, m.name)) + fsize(partFile(m)), partial: fsize(partFile(m)), downloading: jobs.has(lang) };
  });
  return { models, bytes: models.reduce((a, b) => a + b.bytes, 0) };
}

function init({ app, ipcMain }) {
  DIR = path.join(app.getPath('userData'), 'voice-models');
  ipcMain.handle('bts:ext:voice-models', () => list());
  ipcMain.handle('bts:ext:voice-download', async (e, lang) => {
    const m = MODELS[lang]; if (!m) return { ok: false, error: 'Unknown language' };
    if (fs.existsSync(modelFile(m))) return { ok: true };
    let job = jobs.get(lang);
    if (job) { job.senders.add(e.sender); return job.promise; }
    job = { cancelled: false, senders: new Set([e.sender]), total: 0 };
    job.emit = (d) => { for (const s of job.senders) if (!s.isDestroyed()) s.send('bts:ext-ev:voice-progress', { lang, ...d }); };
    job.promise = install(lang, job)
      .then(() => ({ ok: true }))
      .catch(err => job.cancelled || err.message === 'cancelled' ? { ok: false, cancelled: true, error: 'Download cancelled' } : { ok: false, error: friendly(err.message || String(err)) })
      .finally(() => jobs.delete(lang));
    jobs.set(lang, job);
    return job.promise;
  });
  ipcMain.handle('bts:ext:voice-cancel', (e, lang) => {
    const job = jobs.get(lang);
    if (job) { job.cancelled = true; try { job.req && job.req.abort(); } catch { } }
    return { ok: true };
  });
  ipcMain.handle('bts:ext:voice-model', (e, lang) => {
    const m = MODELS[lang]; if (!m) return null;
    try { return new Uint8Array(fs.readFileSync(modelFile(m))); } catch { return null; }
  });
  ipcMain.handle('bts:ext:voice-delete', (e, lang) => {
    const before = list().bytes;
    for (const [l, m] of Object.entries(MODELS)) {
      if (lang && l !== lang) continue;
      const job = jobs.get(l); if (job) { job.cancelled = true; try { job.req && job.req.abort(); } catch { } }
      fs.rmSync(path.join(DIR, m.name), { recursive: true, force: true });
      fs.rmSync(partFile(m), { force: true });
    }
    return { ok: true, freed: before - list().bytes };
  });
}
module.exports = { init, MODELS, zipToTarGz };
