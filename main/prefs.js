// BTS Office — applies File > Options settings that live in the main process (spelling).
// settings.json: spellCheck (false = off), spellLanguages (e.g. ['en-IN', 'en-US', 'en-GB', 'hi']).
// Pages call bts.ext('prefs-apply') after changing them (no restart needed) and bts.ext('prefs-spell') to read the state.
const { session } = require('electron');

// languages the Options page offers, in this order (only those this Electron build can check are shown)
const OFFER = [['en-IN', 'English (India)'], ['en-US', 'English (US)'], ['en-GB', 'English (UK)'], ['hi', 'Hindi']];

function info() {
  const ses = session.defaultSession;
  const avail = ses.availableSpellCheckerLanguages || [];
  let languages = []; try { languages = ses.getSpellCheckerLanguages(); } catch { }
  let enabled = true; try { enabled = ses.isSpellCheckerEnabled(); } catch { }
  return {
    enabled, languages, platform: process.platform,
    // macOS uses the system spell checker and picks languages itself
    choose: process.platform !== 'darwin',
    offer: OFFER.filter(([id]) => avail.includes(id)).map(([id, label]) => ({ id, label })),
  };
}

function apply(s) {
  const ses = session.defaultSession;
  try { ses.setSpellCheckerEnabled(s.spellCheck !== false); } catch (err) { console.error('spellcheck', err); }
  if (process.platform !== 'darwin' && Array.isArray(s.spellLanguages)) {
    const avail = ses.availableSpellCheckerLanguages || [];
    const langs = s.spellLanguages.filter(l => avail.includes(l));
    if (langs.length) { try { ses.setSpellCheckerLanguages(langs); } catch (err) { console.error('spell languages', err); } }
  }
  return info();
}

function init({ app, ipcMain, readJSON, settingsPath }) {
  const run = () => apply(readJSON(settingsPath(), {}));
  if (app.isReady()) run(); else app.whenReady().then(run);
  ipcMain.handle('bts:ext:prefs-apply', () => run());
  ipcMain.handle('bts:ext:prefs-spell', () => info());
}
module.exports = { init, apply, info };
