// BTS Office — automatic updates
//
// Installed copies check https://office.btspl.tech/updates/latest.yml (published by
// tools/release.sh or the GitHub "Release" workflow). When a newer version exists it
// downloads in the background (only changed blocks, thanks to the .blockmap), verifies the
// SHA-512 checksum, and installs silently when the app closes — or immediately if the
// user clicks "Restart now".
//
// Overrides (for testing or a school/lab mirror):
//   env BTS_UPDATE_URL=https://mirror.example/bts/      or  settings.json { "updateUrl": "…" }
//   settings.json { "autoUpdate": false } turns automatic checks off (manual check still works).
const { app, BrowserWindow, dialog, ipcMain } = require('electron');

const DEFAULT_FEED = 'https://office.btspl.tech/updates/';
const FIRST_CHECK_MS = 8 * 1000;
const RECHECK_MS = 4 * 60 * 60 * 1000;

let autoUpdater = null;
let state = { status: 'idle', version: null, percent: 0, notes: '', error: null, current: app.getVersion(), supported: false };
let getSettings = () => ({});
let getWindows = () => [];
let timer = null;

function broadcast() {
  for (const w of BrowserWindow.getAllWindows()) { try { w.webContents.send('bts:update-state', state); } catch {} }
}
function set(patch) { state = { ...state, ...patch }; broadcast(); }

function notesToText(n) {
  if (!n) return '';
  if (Array.isArray(n)) return n.map(x => (x.version ? x.version + ': ' : '') + (x.note || '')).join('\n');
  return String(n).replace(/<[^>]+>/g, '').trim();
}

function supportedPlatform() {
  if (!app.isPackaged && !process.env.BTS_UPDATE_URL) return false;     // dev runs never update
  if (process.platform === 'win32') return true;                         // NSIS installer
  if (process.platform === 'linux') return !!process.env.APPIMAGE;       // AppImage only (.deb uses apt)
  if (process.platform === 'darwin') return true;
  return false;
}

async function check(manual = false) {
  if (!autoUpdater) return state;
  if (['checking', 'downloading'].includes(state.status)) return state;
  if (state.status === 'ready') return state;
  try {
    set({ status: 'checking', error: null, manual });
    await autoUpdater.checkForUpdates();
  } catch (err) {
    set({ status: manual ? 'error' : 'idle', error: friendlyError(err) });
  }
  return state;
}

function friendlyError(err) {
  const m = String((err && err.message) || err || '');
  if (/ENOTFOUND|ECONNREFUSED|ETIMEDOUT|EAI_AGAIN|net::|ENETUNREACH|socket/i.test(m)) return "Couldn't reach the update server. Check your internet connection and try again.";
  if (/404/.test(m)) return 'No update information was found on the update server.';
  if (/sha512|checksum/i.test(m)) return 'The downloaded update was damaged, so it was discarded. It will be downloaded again.';
  return 'The update check failed: ' + m.split('\n')[0].slice(0, 160);
}

async function installNow() {
  if (state.status !== 'ready') return false;
  // Never lose work: if any window has unsaved changes, send the user there first.
  const dirty = getWindows().filter(r => r.dirty && !r.win.isDestroyed());
  if (dirty.length) {
    const w = dirty[0].win; w.show(); w.focus();
    await dialog.showMessageBox(w, {
      type: 'info', buttons: ['OK'], title: 'BTS Office',
      message: 'Save your work before updating',
      detail: `${dirty.length === 1 ? 'A file has' : dirty.length + ' files have'} unsaved changes. Save or close ${dirty.length === 1 ? 'it' : 'them'}, then choose Restart now again. The update also installs automatically the next time you close BTS Office.`,
    });
    return false;
  }
  for (const r of getWindows()) r.forceClose = true;
  setImmediate(() => autoUpdater.quitAndInstall(true, true)); // silent install, reopen afterwards
  return true;
}

function init(opts) {
  getSettings = opts.getSettings || getSettings;
  getWindows = opts.getWindows || getWindows;

  ipcMain.handle('bts:update-state', () => state);
  ipcMain.handle('bts:update-check', () => check(true));
  ipcMain.handle('bts:update-install', () => installNow());
  ipcMain.handle('bts:update-dismiss', () => { set({ dismissed: state.version }); return true; });

  if (!supportedPlatform()) { state.supported = false; return; }

  try { ({ autoUpdater } = require('electron-updater')); }
  catch (err) { console.error('electron-updater missing', err); return; }

  const s = getSettings();
  const feed = process.env.BTS_UPDATE_URL || s.updateUrl || DEFAULT_FEED;
  try { autoUpdater.logger = require('electron-log'); autoUpdater.logger.transports.file.level = 'info'; } catch {}
  autoUpdater.setFeedURL({ provider: 'generic', url: feed, channel: 'latest' });
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;
  autoUpdater.allowDowngrade = false;
  autoUpdater.allowPrerelease = false;
  if (!app.isPackaged) autoUpdater.forceDevUpdateConfig = true;
  state.supported = true;
  state.feed = feed;

  autoUpdater.on('checking-for-update', () => set({ status: 'checking' }));
  autoUpdater.on('update-not-available', () => set({ status: 'uptodate', checkedAt: Date.now() }));
  autoUpdater.on('update-available', (info) => set({ status: 'downloading', version: info.version, percent: 0, notes: notesToText(info.releaseNotes), size: (info.files && info.files[0] && info.files[0].size) || 0 }));
  autoUpdater.on('download-progress', (p) => set({ status: 'downloading', percent: Math.round(p.percent || 0), bytesPerSecond: p.bytesPerSecond }));
  autoUpdater.on('update-downloaded', (info) => set({ status: 'ready', version: info.version, percent: 100, notes: notesToText(info.releaseNotes) || state.notes }));
  autoUpdater.on('error', (err) => set({ status: state.manual ? 'error' : (state.status === 'ready' ? 'ready' : 'idle'), error: friendlyError(err) }));

  const auto = () => { if (getSettings().autoUpdate !== false) check(false); };
  setTimeout(auto, FIRST_CHECK_MS);
  timer = setInterval(auto, RECHECK_MS);
}

module.exports = { init, check, state: () => state };
