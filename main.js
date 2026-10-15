// BTS Office — main process
// Open-source office suite for students. License: GPL-3.0-or-later.
const { app, BrowserWindow, ipcMain, dialog, shell, Menu, nativeTheme, session } = require('electron');
const path = require('path');
const fs = require('fs');
const updater = require('./updater');
const crypto = require('crypto');
const os = require('os');

const APPS = {
  start:  { page: 'index.html',  title: 'BTS Office',  color: '#172033' },
  docs:   { page: 'docs.html',   title: 'BTS Docs',    color: '#3C4FE0' },
  sheets: { page: 'sheets.html', title: 'BTS Sheets',  color: '#0F8A6A' },
  slides: { page: 'slides.html', title: 'BTS Slides',  color: '#D9661A' },
  pdf:    { page: 'pdf.html',    title: 'BTS PDF',     color: '#C7354A' },
};
const EXT_APP = {
  '.docx': 'docs', '.doc': 'docs', '.txt': 'docs', '.html': 'docs', '.htm': 'docs', '.md': 'docs', '.rtf': 'docs', '.odt': 'docs',
  '.xlsx': 'sheets', '.xls': 'sheets', '.csv': 'sheets', '.ods': 'sheets', '.tsv': 'sheets',
  '.pptx': 'slides', '.pdf': 'pdf',
};

const windows = new Map(); // webContents.id -> { win, app, file, dirty }
const recentPath = () => path.join(app.getPath('userData'), 'recent.json');
const settingsPath = () => path.join(app.getPath('userData'), 'settings.json');

function readJSON(p, def) { try { return JSON.parse(fs.readFileSync(p, 'utf8')); } catch { return def; } }
function writeJSON(p, v) { try { fs.mkdirSync(path.dirname(p), { recursive: true }); fs.writeFileSync(p, JSON.stringify(v, null, 2)); } catch {} }

function addRecent(file) {
  if (!file) return;
  let list = readJSON(recentPath(), []).filter(r => r.path !== file);
  list.unshift({ path: file, name: path.basename(file), app: EXT_APP[path.extname(file).toLowerCase()] || 'docs', time: Date.now() });
  writeJSON(recentPath(), list.slice(0, 30));
  try { app.addRecentDocument(file); } catch {}
}

function appForFile(file) { return EXT_APP[path.extname(file).toLowerCase()] || null; }

function createWindow(appName = 'start', file = null) {
  const def = APPS[appName] || APPS.start;
  const isMac = process.platform === 'darwin';
  const win = new BrowserWindow({
    width: appName === 'start' ? 1180 : 1400,
    height: appName === 'start' ? 760 : 900,
    minWidth: 820, minHeight: 560,
    show: false,
    backgroundColor: nativeTheme.shouldUseDarkColors ? '#151a23' : '#f3f5f8',
    title: def.title,
    icon: path.join(__dirname, 'build', 'icon.png'),
    titleBarStyle: 'hidden',
    titleBarOverlay: isMac ? undefined : { color: '#00000000', symbolColor: '#ffffff', height: 40 },
    trafficLightPosition: { x: 14, y: 13 },
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: true,
    },
  });
  const rec = { win, app: appName, file, dirty: false, forceClose: false };
  windows.set(win.webContents.id, rec);
  win.loadFile(path.join(__dirname, 'app', def.page));
  win.once('ready-to-show', () => { win.show(); });
  win.webContents.setWindowOpenHandler(({ url }) => { shell.openExternal(url); return { action: 'deny' }; });
  win.webContents.on('will-navigate', (e, url) => { if (!url.startsWith('file:')) { e.preventDefault(); shell.openExternal(url); } });
  const wcId = win.webContents.id;
  win.on('close', async (e) => {
    if (rec.forceClose || !rec.dirty) return;
    e.preventDefault();
    const r = await dialog.showMessageBox(win, {
      type: 'question', buttons: ['Save', "Don't save", 'Cancel'], defaultId: 0, cancelId: 2,
      title: def.title, message: 'Save changes before closing?',
      detail: 'Your changes will be lost if you close without saving.',
    });
    if (r.response === 0) win.webContents.send('bts:save-then-close');
    else if (r.response === 1) { rec.forceClose = true; win.close(); }
  });
  win.on('closed', () => windows.delete(wcId));
  return win;
}

function openPath(file, fromWin) {
  const a = appForFile(file);
  if (!a) { dialog.showErrorBox('BTS Office', `This file type isn't supported:\n${file}`); return; }
  addRecent(file);
  // reuse an empty start window if the request came from one
  createWindow(a, file);
  if (fromWin) { const r = windows.get(fromWin.webContents.id); if (r && r.app === 'start' && r.closeOnOpen) fromWin.close(); }
}

function fileArgs(argv) {
  return argv.slice(app.isPackaged ? 1 : 2).filter(a => !a.startsWith('-') && fs.existsSync(a) && fs.statSync(a).isFile());
}
function appArg(argv) {
  const a = argv.find(x => x.startsWith('--app='));
  return a ? a.slice(6) : null;
}

// ---------- single instance ----------
const gotLock = app.requestSingleInstanceLock();
if (!gotLock) { app.quit(); }
else {
  app.on('second-instance', (e, argv) => {
    const files = fileArgs(argv);
    if (files.length) files.forEach(f => openPath(f));
    else createWindow(appArg(argv) || 'start');
  });
}

app.on('open-file', (e, file) => { e.preventDefault(); if (app.isReady()) openPath(file); else pendingFiles.push(file); });
const pendingFiles = [];

app.whenReady().then(() => {
  if (process.platform !== 'darwin') Menu.setApplicationMenu(null);
  else Menu.setApplicationMenu(Menu.buildFromTemplate([{ role: 'appMenu' }, { role: 'editMenu' }, { role: 'windowMenu' }]));
  try { session.defaultSession.setSpellCheckerLanguages(['en-US', 'en-GB', 'hi']); } catch {}
  const files = [...pendingFiles, ...fileArgs(process.argv)];
  if (files.length) files.forEach(f => openPath(f));
  else createWindow(appArg(process.argv) || 'start');
  updater.init({ getSettings: () => readJSON(settingsPath(), {}), getWindows: () => [...windows.values()] });
  // feature modules in ./main (spelling menu, online lookups, voice models, ...)
  try {
    for (const f of fs.readdirSync(path.join(__dirname, 'main')).filter(f => f.endsWith('.js')).sort()) {
      try { require('./main/' + f).init({ app, ipcMain, BrowserWindow, dialog, shell, readJSON, writeJSON, settingsPath }); }
      catch (err) { console.error('module', f, 'failed', err); }
    }
  } catch (err) { console.error('modules', err); }
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow('start'); });
});
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });

// ---------- IPC ----------
const recOf = (e) => windows.get(e.sender.id);
const winOf = (e) => BrowserWindow.fromWebContents(e.sender);

const FILTERS = {
  docs: [{ name: 'Documents', extensions: ['docx', 'txt', 'html', 'htm', 'md', 'odt', 'rtf'] }],
  sheets: [{ name: 'Spreadsheets', extensions: ['xlsx', 'xls', 'csv', 'tsv', 'ods'] }],
  slides: [{ name: 'Presentations', extensions: ['pptx'] }],
  pdf: [{ name: 'PDF documents', extensions: ['pdf'] }],
  all: [{ name: 'All supported files', extensions: ['docx', 'txt', 'html', 'md', 'odt', 'xlsx', 'xls', 'csv', 'ods', 'pptx', 'pdf'] }],
};

ipcMain.handle('bts:info', (e) => {
  const r = recOf(e);
  return { app: r?.app, file: r?.file || null, version: app.getVersion(), platform: process.platform, dark: nativeTheme.shouldUseDarkColors };
});

ipcMain.handle('bts:initial-file', async (e) => {
  const r = recOf(e);
  if (!r || !r.file) return null;
  try { const data = fs.readFileSync(r.file); return { path: r.file, name: path.basename(r.file), data: new Uint8Array(data) }; }
  catch (err) { return { error: String(err.message || err) }; }
});

ipcMain.handle('bts:open-dialog', async (e, appName) => {
  const res = await dialog.showOpenDialog(winOf(e), { properties: ['openFile'], filters: [...(FILTERS[appName] || FILTERS.all), ...FILTERS.all.filter(() => appName !== 'all'), { name: 'All files', extensions: ['*'] }] });
  if (res.canceled || !res.filePaths.length) return null;
  const file = res.filePaths[0];
  const target = appForFile(file);
  const r = recOf(e);
  if (!target) return { error: 'This file type isn\'t supported.' };
  if (target !== appName || (r && r.app === 'start')) { openPath(file); return { handledElsewhere: true }; }
  addRecent(file);
  return { path: file, name: path.basename(file), data: new Uint8Array(fs.readFileSync(file)) };
});

ipcMain.handle('bts:read-file', async (e, file) => {
  try { return { path: file, name: path.basename(file), data: new Uint8Array(fs.readFileSync(file)) }; }
  catch (err) { return { error: String(err.message || err) }; }
});

ipcMain.handle('bts:save-dialog', async (e, { defaultName, filters }) => {
  const res = await dialog.showSaveDialog(winOf(e), { defaultPath: defaultName, filters });
  return res.canceled ? null : res.filePath;
});

ipcMain.handle('bts:write-file', async (e, { file, data }) => {
  try {
    const buf = typeof data === 'string' ? Buffer.from(data, 'utf8') : Buffer.from(data);
    fs.writeFileSync(file, buf);
    const r = recOf(e); if (r && !file.toLowerCase().endsWith('.pdf')) r.file = file;
    addRecent(file);
    snapshotVersion(file, buf);
    return { ok: true, path: file, name: path.basename(file) };
  } catch (err) { return { error: String(err.message || err) }; }
});

// ---------- version history: every save of a .docx/.xlsx/.pptx keeps a copy (last 30 per file) ----------
const VERSION_EXT = ['.docx', '.xlsx', '.pptx'];
const versionsDir = (file) => path.join(app.getPath('userData'), 'versions', crypto.createHash('sha1').update(path.resolve(file).toLowerCase()).digest('hex').slice(0, 16));
function snapshotVersion(file, buf) {
  try {
    const ext = path.extname(file).toLowerCase(); if (!VERSION_EXT.includes(ext)) return;
    if (readJSON(settingsPath(), {}).versionHistory === false) return;
    const dir = versionsDir(file); fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, Date.now() + ext), buf);
    writeJSON(path.join(dir, 'meta.json'), { path: file });
    const all = fs.readdirSync(dir).filter(f => f.endsWith(ext)).sort();
    for (const old of all.slice(0, Math.max(0, all.length - 30))) fs.unlinkSync(path.join(dir, old));
  } catch (err) { console.error('version snapshot failed', err); }
}
ipcMain.handle('bts:versions', (e, file) => {
  try {
    if (!file) return []; const dir = versionsDir(file); if (!fs.existsSync(dir)) return [];
    return fs.readdirSync(dir).filter(f => /^\d+\./.test(f)).map(f => ({ id: f, time: +f.split('.')[0], size: fs.statSync(path.join(dir, f)).size })).sort((a, b) => b.time - a.time);
  } catch { return []; }
});
ipcMain.handle('bts:version-open', (e, { file, id }) => {
  const src = path.join(versionsDir(file), path.basename(id)); if (!fs.existsSync(src)) return false;
  const d = new Date(+id.split('.')[0]); const stamp = d.toISOString().slice(0, 16).replace('T', ' ').replace(':', '');
  const outDir = path.join(os.tmpdir(), 'bts-office-versions'); fs.mkdirSync(outDir, { recursive: true });
  const out = path.join(outDir, `${path.basename(file, path.extname(file))} (version ${stamp})${path.extname(file)}`);
  fs.copyFileSync(src, out); createWindow(appForFile(out), out); return true;
});

ipcMain.handle('bts:print-pdf', async (e, opts = {}) => {
  const data = await e.sender.printToPDF({
    printBackground: true,
    pageSize: opts.pageSize || 'A4',
    landscape: !!opts.landscape,
    margins: opts.margins || { marginType: 'none' },
    preferCSSPageSize: opts.preferCSSPageSize !== false,
  });
  return new Uint8Array(data);
});

ipcMain.handle('bts:print', async (e) => new Promise(res => e.sender.print({ printBackground: true }, (ok, reason) => res({ ok, reason }))));

ipcMain.handle('bts:open-app', (e, { app: a, file }) => { if (file) openPath(file); else createWindow(a || 'start'); return true; });
ipcMain.handle('bts:open-external', (e, url) => { if (/^https?:|^mailto:/.test(url)) shell.openExternal(url); return true; });
ipcMain.handle('bts:show-in-folder', (e, file) => { shell.showItemInFolder(file); return true; });
ipcMain.handle('bts:recent', () => readJSON(recentPath(), []).filter(r => fs.existsSync(r.path)));
ipcMain.handle('bts:clear-recent', () => { writeJSON(recentPath(), []); return true; });
ipcMain.handle('bts:settings-get', () => readJSON(settingsPath(), {}));
ipcMain.handle('bts:settings-set', (e, s) => { const cur = readJSON(settingsPath(), {}); writeJSON(settingsPath(), { ...cur, ...s }); return true; });

ipcMain.on('bts:dirty', (e, dirty) => { const r = recOf(e); if (r) r.dirty = !!dirty; const w = winOf(e); if (w && process.platform === 'darwin') w.setDocumentEdited(!!dirty); });
ipcMain.on('bts:set-title', (e, t) => { const w = winOf(e); if (w) w.setTitle(t); });
ipcMain.on('bts:titlebar-color', (e, { color, symbol }) => { const w = winOf(e); try { if (w && process.platform !== 'darwin') w.setTitleBarOverlay({ color, symbolColor: symbol, height: 40 }); } catch {} });
ipcMain.on('bts:close-now', (e) => { const r = recOf(e); if (r) { r.forceClose = true; r.win.close(); } });
ipcMain.on('bts:fullscreen', (e, on) => { const w = winOf(e); if (w) w.setFullScreen(!!on); });
ipcMain.on('bts:close-on-open', (e, v) => { const r = recOf(e); if (r) r.closeOnOpen = !!v; });

// ---------- developer test hook: BTS_SHOT=out.png [BTS_EVAL=script.js] [BTS_WAIT=ms] electron . --app=docs ----------
if (process.env.BTS_SHOT && (!app.isPackaged || process.env.BTS_ALLOW_TEST_HOOK === '1')) {
  app.whenReady().then(() => setTimeout(async () => {
    const w = BrowserWindow.getAllWindows()[0];
    try {
      if (process.env.BTS_EVAL) {
        const code = fs.readFileSync(process.env.BTS_EVAL, 'utf8');
        const out = await w.webContents.executeJavaScript(`(async()=>{try{${code}\n}catch(e){return 'EVAL ERROR: '+(e.stack||e)}})()`);
        if (out !== undefined) console.log('EVAL:', typeof out === 'string' ? out : JSON.stringify(out));
        await new Promise(r => setTimeout(r, +(process.env.BTS_WAIT2 || 600)));
      }
      const img = await w.webContents.capturePage();
      fs.writeFileSync(process.env.BTS_SHOT, img.toPNG());
    } catch (e) { console.error('SHOT ERROR', e); }
    app.exit(0);
  }, +(process.env.BTS_WAIT || 2500)));
  app.on('web-contents-created', (e, wc) => wc.on('console-message', (ev, level, msg, line, src) => { if (level >= 2) console.log('[renderer]', msg, (src || '').split('/').pop() + ':' + line); }));
}
