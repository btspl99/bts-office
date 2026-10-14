// BTS Office — secure bridge between the UI and the system (no Node access in pages)
const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('bts', {
  isDesktop: true,
  pathFor: (file) => { try { return webUtils.getPathForFile(file); } catch { return ''; } },
  info: () => ipcRenderer.invoke('bts:info'),
  initialFile: () => ipcRenderer.invoke('bts:initial-file'),
  openDialog: (appName) => ipcRenderer.invoke('bts:open-dialog', appName),
  readFile: (file) => ipcRenderer.invoke('bts:read-file', file),
  saveDialog: (defaultName, filters) => ipcRenderer.invoke('bts:save-dialog', { defaultName, filters }),
  writeFile: (file, data) => ipcRenderer.invoke('bts:write-file', { file, data }),
  printToPDF: (opts) => ipcRenderer.invoke('bts:print-pdf', opts),
  print: () => ipcRenderer.invoke('bts:print'),
  openApp: (app, file) => ipcRenderer.invoke('bts:open-app', { app, file }),
  openExternal: (url) => ipcRenderer.invoke('bts:open-external', url),
  showInFolder: (file) => ipcRenderer.invoke('bts:show-in-folder', file),
  recent: () => ipcRenderer.invoke('bts:recent'),
  clearRecent: () => ipcRenderer.invoke('bts:clear-recent'),
  getSettings: () => ipcRenderer.invoke('bts:settings-get'),
  setSettings: (s) => ipcRenderer.invoke('bts:settings-set', s),
  setDirty: (d) => ipcRenderer.send('bts:dirty', d),
  setTitle: (t) => ipcRenderer.send('bts:set-title', t),
  titlebarColor: (color, symbol) => ipcRenderer.send('bts:titlebar-color', { color, symbol }),
  closeNow: () => ipcRenderer.send('bts:close-now'),
  fullscreen: (on) => ipcRenderer.send('bts:fullscreen', on),
  closeOnOpen: (v) => ipcRenderer.send('bts:close-on-open', v),
  onSaveThenClose: (cb) => ipcRenderer.on('bts:save-then-close', () => cb()),
});
