// BTS Office — right-click menu for text: spelling suggestions, Add to dictionary, cut/copy/paste.
// Pages that draw their own context menu call preventDefault(), so this only appears on plain text fields
// and the Docs editor.
const { Menu } = require('electron');

function init({ app }) {
  app.on('web-contents-created', (e, wc) => {
    wc.on('context-menu', (ev, p) => {
      const items = [];
      if (p.misspelledWord) {
        const sug = (p.dictionarySuggestions || []).slice(0, 6);
        if (sug.length) sug.forEach(s => items.push({ label: s, click: () => wc.replaceMisspelling(s) }));
        else items.push({ label: 'No spelling suggestions', enabled: false });
        items.push({ type: 'separator' });
        items.push({ label: `Add "${p.misspelledWord}" to dictionary`, click: () => wc.session.addWordToSpellCheckerDictionary(p.misspelledWord) });
        items.push({ label: 'Ignore', click: () => wc.replaceMisspelling(p.misspelledWord) });
        items.push({ type: 'separator' });
      }
      const f = p.editFlags || {};
      if (p.isEditable) {
        items.push({ label: 'Cut', role: 'cut', accelerator: 'CmdOrCtrl+X', enabled: f.canCut });
        items.push({ label: 'Copy', role: 'copy', accelerator: 'CmdOrCtrl+C', enabled: f.canCopy });
        items.push({ label: 'Paste', role: 'paste', accelerator: 'CmdOrCtrl+V', enabled: f.canPaste });
        items.push({ label: 'Paste as plain text', role: 'pasteAndMatchStyle', accelerator: 'CmdOrCtrl+Shift+V', enabled: f.canPaste });
        items.push({ type: 'separator' });
        items.push({ label: 'Select all', role: 'selectAll', accelerator: 'CmdOrCtrl+A' });
      } else if (p.selectionText && p.selectionText.trim()) {
        items.push({ label: 'Copy', role: 'copy', accelerator: 'CmdOrCtrl+C' });
      }
      if (!items.length) return;
      Menu.buildFromTemplate(items).popup({ window: require('electron').BrowserWindow.fromWebContents(wc) || undefined });
    });
  });
}
module.exports = { init };
