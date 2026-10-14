/* BTS PDF — modern PDF reader & light annotator. Entry module: shell config, loading, commands, keys. */
import { S, pdfjsLib, CSS_UNITS, $, $$, esc, isTyping, MARKUP } from './state.js';
import * as V from './viewer.js';
import * as A from './annotations.js';
import * as Find from './find.js';
import * as SB from './sidebar.js';
import * as IO from './io.js';

const ASSETS = {
  standardFontDataUrl: new URL('../vendor/pdfjs/standard_fonts/', import.meta.url).href,
  cMapUrl: new URL('../vendor/pdfjs/cmaps/', import.meta.url).href,
  cMapPacked: true,
  isEvalSupported: false,
  enableXfa: false,
};

/* ======================= loading ======================= */
function askPassword(wrong) {
  return BTS.dialog({
    title: wrong ? 'Incorrect password' : 'Password required',
    body: `<p style="margin:2px 0 12px;max-width:380px;line-height:1.5;color:var(--ink-2)">${wrong ? 'That password didn’t work. Please try again.' : 'This PDF is protected. Enter its password to open it.'}</p>
           <div class="field"><label for="pdf-pw">Password</label><input id="pdf-pw" type="password" autocomplete="off"></div>`,
    buttons: [{ label: 'Cancel', value: null }, { label: 'Open', primary: true, value: (r) => $('#pdf-pw', r).value }],
  });
}

async function openBytes(bytes, { keepAnns = false, keepView = false } = {}) {
  const prevCur = S.cur, prevScale = S.scale, prevFit = S.fit;
  const task = pdfjsLib.getDocument({ data: bytes.slice(), ...ASSETS });   // pdf.js transfers the buffer — keep our copy
  let cancelled = false;
  S.encrypted = false;
  task.onPassword = async (update, reason) => {
    S.encrypted = true;
    const pw = await askPassword(reason === pdfjsLib.PasswordResponses.INCORRECT_PASSWORD);
    if (pw == null) { cancelled = true; task.destroy(); } else update(pw);
  };
  let doc;
  try { doc = await task.promise; }
  catch (err) {
    if (cancelled) throw new Error('a password is needed to open it');
    if (err?.name === 'InvalidPDFException') throw new Error('it isn’t a valid PDF file, or it is damaged');
    throw err;
  }
  closeDoc({ keepAnns });
  S.doc = doc; S.loadingTask = task; S.bytes = bytes;
  try { const md = await doc.getMetadata(); S.meta = md; S.encrypted ||= !!md?.info?.IsEncrypted; } catch { S.meta = null; }
  document.body.classList.remove('nodoc');
  await V.buildPages();
  if (keepView) { S.fit = prevFit; V.setScale(prevScale, { fit: prevFit }); }
  else { S.fit = 'auto'; S.scale = 0; V.setScale(V.fitScale('auto'), { fit: 'auto' }); }
  S._curInit = false; V.setCurrent(keepView ? Math.min(prevCur, S.pages.length - 1) : 0);
  requestAnimationFrame(() => V.goTo(S.cur));
  $('#pg-total').textContent = '/ ' + S.pages.length;
  $('#pg-input').style.width = (String(S.pages.length).length * 8 + 22) + 'px';
  SB.afterLoad();
  A.renderAll();
  updateStatus();
  updateToolUI();
}

function closeDoc({ keepAnns = false } = {}) {
  V.teardown();
  Find.resetFind();
  SB.resetSidebar();
  if (!keepAnns) A.resetAnnotations();
  S.textCache.clear();
  if (S.doc) { const d = S.doc; S.doc = null; d.destroy().catch(() => {}); }
  S.meta = null; S.cur = 0; S._curInit = false;
}

function newDocument() {
  closeDoc();
  S.bytes = null; S.modified = false; S.rot = 0; S.fileName = '';
  document.body.classList.add('nodoc');
  BTS.setTitle(BTS.cfg.defaultTitle); BTS.filePath = null; BTS.markClean();
  $('#pg-input').value = ''; $('#pg-total').textContent = '/ 0';
  BTS.setStatus('page', ''); BTS.setStatus('size', ''); BTS.setStatus('anns', '');
  BTS.zoom = 100; const zv = $('#sb-zoom-val'); if (zv) zv.textContent = '100%';
  fillRecent();
}

/* ======================= status & page box ======================= */
const PAPER = [['A3', 297, 420], ['A4', 210, 297], ['A5', 148, 210], ['Letter', 215.9, 279.4], ['Legal', 215.9, 355.6], ['Tabloid', 279.4, 431.8]];
export function pageSizeLabel(p) {
  const vp = p.proxy.getViewport({ scale: 1, rotation: p.proxy.rotate });
  const w = vp.width * 25.4 / 72, h = vp.height * 25.4 / 72;
  const lo = Math.min(w, h), hi = Math.max(w, h);
  const named = PAPER.find(([, a, b]) => Math.abs(a - lo) < 2.5 && Math.abs(b - hi) < 2.5);
  const mm = `${Math.round(w)} × ${Math.round(h)} mm`;
  return named ? `${named[0]}${w > h ? ' landscape' : ''} · ${mm}` : mm;
}
function updateStatus() {
  if (!S.doc) return;
  const p = S.pages[S.cur]; if (!p) return;
  BTS.setStatus('page', `${icon('file-text', 14)}<span>Page ${S.cur + 1} of ${S.pages.length}</span>`);
  BTS.setStatus('size', `<span>${esc(pageSizeLabel(p))}</span>`);
  const n = S.anns.length;
  BTS.setStatus('anns', n ? `${icon('message-square-text', 14)}<span>${n} annotation${n === 1 ? '' : 's'}</span>` : '');
  const inp = $('#pg-input'); if (document.activeElement !== inp) inp.value = S.cur + 1;
  $('#hud-page').textContent = `${S.cur + 1} / ${S.pages.length}`;
}
V.hooks.current.push(updateStatus);
V.hooks.layout.push(() => A.renderAll());
A.listeners.change.push(() => { updateStatus(); updateToolUI(); });


/* ======================= tools ======================= */
function setTool(t) {
  if (S.tool === t && t !== 'text') t = 'text';         // clicking an active tool turns it off
  A.closeTextEditor();
  S.tool = t; document.body.dataset.tool = t;
  if (MARKUP.has(t) && getSelection().toString().trim()) A.markupFromSelection(t);
  updateToolUI();
  const hint = { highlight: 'Select text to highlight it', underline: 'Select text to underline it', strike: 'Select text to strike it out', note: 'Click on the page to add a sticky note', textbox: 'Click on the page to type', pen: 'Draw on the page · hold Shift for straight shapes', rect: 'Drag to draw a rectangle · Shift for a square', ellipse: 'Drag to draw an ellipse · Shift for a circle', arrow: 'Drag to draw an arrow · Shift snaps to 45°', eraser: 'Click or drag over annotations to remove them', hand: 'Drag to move around the page' }[t];
  BTS.setStatus('tool', hint ? `<span class="tool-hint">${icon('info', 14)}${esc(hint)}</span>` : '');
}
function currentColorKey() {
  const sel = S.anns.find(a => a.id === S.sel);
  if (sel) return { sel, color: sel.color };
  const k = S.colors[S.tool] ? S.tool : 'highlight';
  return { key: k, color: S.colors[k] };
}
function updateToolUI() {
  $$('#ribbon [data-cmd="tool"]').forEach(b => b.classList.toggle('on', b.dataset.tool === S.tool));
  const { color } = currentColorKey();
  $('#color-dot').style.background = color;
  $$('#ribbon .sw').forEach(b => b.classList.toggle('on', b.dataset.c.toLowerCase() === String(color).toLowerCase()));
  const sel = S.anns.find(a => a.id === S.sel);
  const w = sel && 'w' in sel ? sel.w : S.width;
  const ws = $('#width-sel'); if (![...ws.options].some(o => +o.value === w)) ws.insertAdjacentHTML('beforeend', `<option value="${w}">${w} pt</option>`);
  ws.value = String(w);
  $$('#ribbon [data-cmd="delete-annot"]').forEach(b => b.disabled = !S.sel);
}
function applyColor(c) {
  if (!c) return;
  const { sel, key } = currentColorKey();
  if (sel) { A.setColorOfSelected(c); S.colors[sel.type] = c; }
  else S.colors[key] = c;
  updateToolUI();
}

/* ======================= view helpers ======================= */
function setMode(m) {
  V.setMode(m);
  $$('#ribbon [data-cmd="mode"]').forEach(b => b.classList.toggle('on', b.dataset.mode === m));
  try { localStorage.setItem('bts-pdf-mode', m); } catch {}
}
function toggleDarkRead(on = !S.darkRead) {
  S.darkRead = on; document.body.classList.toggle('dark-read', on);
  $$('#ribbon [data-cmd="dark-read"]').forEach(b => b.classList.toggle('on', on));
  try { localStorage.setItem('bts-pdf-darkread', on ? '1' : ''); } catch {}
}
let fullscreen = false;
function setFullscreen(on = !fullscreen) {
  fullscreen = on;
  document.body.classList.toggle('reading', on);
  if (window.bts?.fullscreen) bts.fullscreen(on);
  else if (on) document.documentElement.requestFullscreen?.().catch(() => {});
  else if (document.fullscreenElement) document.exitFullscreen?.();
  if (on) BTS.toast('Full screen · press Esc to exit', { ms: 1800 });
  setTimeout(refit, 250);
}
function refit() { if (S.doc && S.fit) V.setScale(V.fitScale(S.fit), { fit: S.fit }); }

/* ======================= tools tab ======================= */
async function combine() {
  let f = null;
  if (window.bts?.isDesktop) {
    f = await bts.openDialog('pdf');
    if (f?.handledElsewhere || f?.error) { if (f.error) BTS.toast(f.error, { error: true }); return; }
  } else f = await BTS.pickFile('.pdf,application/pdf');
  if (!f) return;
  try {
    BTS.toast('Combining…', { ms: 1200 });
    const { bytes, added } = await IO.combineWith(f.data);
    S.modified = true;
    await openBytes(bytes, { keepAnns: true, keepView: true });
    BTS.markDirty();
    BTS.toast(`Added ${added} page${added === 1 ? '' : 's'} from ${f.name}`);
  } catch (e) { console.warn(e); BTS.toast(e.message || String(e), { error: true, ms: 4500 }); }
}
async function copyPageText() {
  const tc = await V.getText(S.cur);
  let s = ''; for (const it of tc.items) { if (it.str === undefined) continue; s += it.str + (it.hasEOL ? '\n' : ''); }
  s = s.trim();
  if (!s) return BTS.toast('This page has no selectable text');
  await navigator.clipboard.writeText(s); BTS.toast('Page text copied');
}

/* ======================= document info ======================= */
function pdfDate(d) {
  const m = /^D:(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?/.exec(d || ''); if (!m) return '';
  const dt = new Date(+m[1], (+m[2] || 1) - 1, +m[3] || 1, +m[4] || 0, +m[5] || 0);
  return dt.toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' });
}
function infoRows() {
  const i = S.meta?.info || {};
  const p = S.pages[S.cur] || S.pages[0];
  return [
    ['Title', i.Title || '—'], ['Author', i.Author || '—'], ['Subject', i.Subject || ''], ['Keywords', i.Keywords || ''],
    ['Pages', S.pages.length], ['Page size', p ? pageSizeLabel(p) : ''], ['PDF version', i.PDFFormatVersion || '—'],
    ['Created', pdfDate(i.CreationDate)], ['Modified', pdfDate(i.ModDate)], ['Application', i.Creator || ''], ['PDF producer', i.Producer || ''],
    ['File size', S.bytes ? fmtSize(S.bytes.length) : ''], ['Protected', S.encrypted ? 'Yes (password / encryption)' : 'No'],
    ['Tagged / outline', ''], ['Annotations', S.anns.length ? String(S.anns.length) + ' added in BTS PDF' : 'None added'],
  ].filter(([k, v]) => v !== '' && k !== 'Tagged / outline');
}
const fmtSize = (n) => n > 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';
function infoHTML() {
  if (!S.doc) return `<p class="empty-note">No PDF is open.</p>`;
  return `<div class="pdf-info">${infoRows().map(([k, v]) => `<div class="k">${esc(k)}</div><div class="v">${esc(v)}</div>`).join('')}</div>${BTS.filePath ? `<p class="pdf-info-path">${esc(BTS.filePath)}</p>` : ''}`;
}

/* ======================= empty state ======================= */
async function fillRecent() {
  const wrap = $('#recent-wrap'), list = $('#recent-list');
  if (!window.bts?.recent) return;
  let r = [];
  try { r = (await bts.recent()).filter(x => /\.pdf$/i.test(x.name)).slice(0, 6); } catch {}
  wrap.classList.toggle('hidden', !r.length);
  list.innerHTML = r.map(x => `<button class="rec" data-path="${esc(x.path)}" title="${esc(x.path)}"><span class="fi">PDF</span><span class="nm"><b>${esc(x.name)}</b><small>${esc(x.path.replace(/[\\/][^\\/]*$/, ''))}</small></span></button>`).join('');
}

/* ======================= commands ======================= */
function onCommand(cmd, el) {
  switch (cmd) {
    case 'first-page': return V.goTo(0);
    case 'last-page': return V.goTo(S.pages.length - 1);
    case 'prev-page': return V.prev();
    case 'next-page': return V.next();
    case 'goto-page': { const n = parseInt(el.value, 10); if (n >= 1 && n <= S.pages.length) V.goTo(n - 1); else el.value = S.cur + 1; $('#viewer').focus({ preventScroll: true }); return; }
    case 'fit-width': return V.applyFit('width');
    case 'fit-page': return V.applyFit('page');
    case 'actual-size': return V.setScale(CSS_UNITS, { fit: null });
    case 'rotate-left': V.rotate(-90); SB.rebuildThumbs(); return;
    case 'rotate-right': V.rotate(90); SB.rebuildThumbs(); return;
    case 'find': return Find.openFind();
    case 'find-next': return Find.step(1);
    case 'find-prev': return Find.step(-1);
    case 'find-close': return Find.closeFind();
    case 'find-case': return Find.toggleCase();
    case 'tool': return setTool(el.dataset.tool);
    case 'color': return BTS.colorPicker(el, { value: currentColorKey().color }).then(c => { if (c && c !== 'auto' && c !== 'none') applyColor(c); });
    case 'swatch': return applyColor(el.dataset.c);
    case 'width': { const w = +el.value; S.width = w; A.setWidthOfSelected(w); return; }
    case 'delete-annot': return A.deleteSelected();
    case 'clear-page': return A.clearPage(S.cur);
    case 'undo': return S.doc && A.undo();
    case 'redo': return S.doc && A.redo();
    case 'mode': return setMode(el.dataset.mode);
    case 'sidebar': return S.doc || el.dataset.pane ? SB.showSidebar(el.dataset.pane, true) : null;
    case 'sidebar-pane': return SB.showSidebar(el.dataset.pane);
    case 'sidebar-close': return SB.hideSidebar();
    case 'dark-read': return toggleDarkRead();
    case 'fullscreen': return setFullscreen();
    case 'combine': return combine();
    case 'extract-page': return IO.extractPage(S.cur).catch(e => BTS.toast(e.message || String(e), { error: true }));
    case 'page-image': return IO.pageImage(S.cur).catch(e => BTS.toast(e.message || String(e), { error: true }));
    case 'copy-page-text': return copyPageText();
    case 'properties': return BTS.dialog({ title: 'Document properties', body: infoHTML(), buttons: [{ label: 'Close', primary: true, value: true }], width: 520 });
  }
}

/* ======================= keyboard ======================= */
function onKey(e) {
  const mod = e.ctrlKey || e.metaKey;
  const k = e.key;
  if (mod && !e.altKey && k.toLowerCase() === 'f') { e.preventDefault(); Find.openFind(); return; }
  if (!S.doc) return;
  if (mod && (k === '=' || k === '+')) { e.preventDefault(); BTS.setZoom(BTS.zoom + zoomStep(1)); return; }
  if (mod && k === '-') { e.preventDefault(); BTS.setZoom(BTS.zoom - zoomStep(-1)); return; }
  if (mod && k === '0') { e.preventDefault(); V.applyFit('page'); return; }
  if (mod && k === '1') { e.preventDefault(); V.setScale(CSS_UNITS, { fit: null }); return; }
  if (mod && k === '2') { e.preventDefault(); V.applyFit('width'); return; }
  if (mod && e.shiftKey && k.toLowerCase() === '+') { e.preventDefault(); V.rotate(90); return; }
  if (isTyping(e.target) || $('.modal-back') || $('#backstage.open')) return;
  if (mod && !e.shiftKey && k.toLowerCase() === 'z') { e.preventDefault(); A.undo(); return; }
  if (mod && (k.toLowerCase() === 'y' || (e.shiftKey && k.toLowerCase() === 'z'))) { e.preventDefault(); A.redo(); return; }
  if (mod && k.toLowerCase() === 'g') { e.preventDefault(); Find.step(e.shiftKey ? -1 : 1); return; }
  if (mod || e.altKey) return;
  const v = $('#viewer');
  const hScroll = v.scrollWidth > v.clientWidth + 2;
  switch (k) {
    case 'PageDown': e.preventDefault(); V.next(); break;
    case 'PageUp': e.preventDefault(); V.prev(); break;
    case 'Home': e.preventDefault(); V.goTo(0); break;
    case 'End': e.preventDefault(); V.goTo(S.pages.length - 1); break;
    case 'ArrowRight': if (!hScroll) { e.preventDefault(); V.next(); } break;
    case 'ArrowLeft': if (!hScroll) { e.preventDefault(); V.prev(); } break;
    case 'ArrowDown': if (S.mode === 'single' && v.scrollTop + v.clientHeight >= v.scrollHeight - 2) { e.preventDefault(); V.next(); } else if (document.activeElement !== v) { e.preventDefault(); v.scrollTop += 60; } break;
    case 'ArrowUp': if (S.mode === 'single' && v.scrollTop <= 1) { e.preventDefault(); V.prev(); } else if (document.activeElement !== v) { e.preventDefault(); v.scrollTop -= 60; } break;
    case 'F11': e.preventDefault(); setFullscreen(); break;
    case 'Escape': if (fullscreen) { e.preventDefault(); setFullscreen(false); } else if (S.tool !== 'text' && !S.sel) setTool('text'); break;
    case 'h': setTool('highlight'); break;
    case 'v': setTool('text'); break;
    case 'p': setTool('pen'); break;
    case 'n': setTool('note'); break;
    case 't': setTool('textbox'); break;
    case 'e': setTool('eraser'); break;
    case ' ': if (document.activeElement !== v) { e.preventDefault(); v.scrollTop += (e.shiftKey ? -1 : 1) * v.clientHeight * 0.9; } break;
  }
}
const ZOOMS = [10, 25, 33, 50, 67, 75, 80, 90, 100, 110, 125, 150, 175, 200, 250, 300, 400, 500];
function zoomStep(dir) {
  const z = BTS.zoom;
  const nxt = dir > 0 ? ZOOMS.find(x => x > z + 0.5) : [...ZOOMS].reverse().find(x => x < z - 0.5);
  return Math.abs((nxt ?? z) - z) || 10;
}

/* ======================= init ======================= */
BTS.init({
  app: 'pdf', name: 'BTS PDF', letter: 'R', defaultTitle: 'Document',
  saveFormats: [{ ext: 'pdf', label: 'PDF with annotations', icon: 'file-text', hint: 'A copy of the PDF with your highlights, drawings and notes added' }],
  openAccept: '.pdf,application/pdf',
  minZoom: 10, maxZoom: 500, zoom: 100,
  shortcuts: [['PgUp / PgDn', 'Previous / next page'], ['Home / End', 'First / last page'], ['Ctrl+wheel', 'Zoom'], ['Ctrl+0 / Ctrl+2', 'Fit page / fit width'], ['H · P · N · T · E · V', 'Highlight · pen · note · text · eraser · select'], ['Delete', 'Delete selected annotation'], ['F11', 'Full screen reading']],
  blankThumb: `<div style="display:grid;place-items:center;gap:6px;color:var(--ink-3)">${icon('book-open', 34)}<small>Close PDF</small></div>`,
  onCommand,
  async load(f) {
    const bytes = f.data instanceof Uint8Array ? f.data : new Uint8Array(f.data);
    S.modified = false; S.rot = 0; S.fileName = f.name;
    await openBytes(bytes);
  },
  async save() { if (!S.doc) throw new Error('No PDF is open'); A.closeTextEditor(); A.closeNotePopup(); return IO.buildAnnotatedPdf(); },
  newDocument,
  hasContent: () => !!S.doc,
  exportPDF: async () => { A.closeTextEditor(); A.closeNotePopup(); return IO.preparePrint(); },
  afterExportPDF: () => IO.cleanupPrint(),
  onZoom(pct) { if (!S.doc) return; V.setScale(pct / 100 * CSS_UNITS, { fit: null }); },
  search(text) { if (!S.doc) return; Find.openFind(text); Find.runFind(text); },
  focus() { $('#viewer').focus({ preventScroll: true }); },
  infoHTML,
});
BTS.hasContent = () => !!S.doc;   // the shell reads BTS.hasContent (not cfg.hasContent)

/* "Export → PDF": a viewer shouldn't re-print a PDF; save the annotated copy instead. */
BTS.exportPDF = async () => {
  BTS.closeBackstage();
  if (!S.doc) return BTS.toast('Open a PDF first');
  const name = BTS.title + (S.anns.length ? ' (annotated)' : '') + '.pdf';
  try {
    const data = await BTS.cfg.save('pdf');
    if (window.bts?.isDesktop) {
      const p = await bts.saveDialog(name, [{ name: 'PDF document', extensions: ['pdf'] }]); if (!p) return;
      const r = await bts.writeFile(p, data);
      if (r.error) BTS.toast(r.error, { error: true }); else BTS.toast('Exported ' + r.name);
    } else BTS.download(name, data, 'application/pdf');
  } catch (e) { BTS.toast(e.message || String(e), { error: true }); }
};
/* Print: without a document there is nothing to print. */
const shellPrint = BTS.print;
BTS.print = async () => { if (!S.doc) { BTS.closeBackstage(); return BTS.toast('Open a PDF first'); } return shellPrint(); };

// wiring
A.initInteractions();
Find.initFind();
const sidePref = SB.initSidebar();
document.addEventListener('keydown', onKey);
document.addEventListener('keydown', (e) => { if ((e.ctrlKey || e.metaKey) && ['=', '+', '-', '0'].includes(e.key) && !e.altKey) e.preventDefault(); }, true);

const viewer = $('#viewer');
viewer.addEventListener('scroll', () => V.updateCurrent(), { passive: true });
viewer.addEventListener('wheel', (e) => {
  if (!S.doc) return;
  if (e.ctrlKey || e.metaKey) {
    e.preventDefault();
    const f = Math.exp(-e.deltaY * (e.deltaMode === 1 ? 0.05 : 0.0022));
    V.setScale(S.scale * f, { cx: e.clientX, cy: e.clientY, fit: null });
    return;
  }
  V.singleWheel(e);
}, { passive: false });
new ResizeObserver(() => { clearTimeout(refit.t); refit.t = setTimeout(refit, 60); }).observe(viewer);

$('#recent-list').addEventListener('click', async (e) => {
  const b = e.target.closest('[data-path]'); if (!b) return;
  const f = await bts.readFile(b.dataset.path);
  if (f.error) return BTS.toast(f.error, { error: true });
  BTS.loadFile(f);
});
$('#pg-input').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.target.dispatchEvent(new Event('change', { bubbles: true })); } if (e.key === 'Escape') { e.target.value = S.cur + 1; e.target.blur(); } });
$('#pg-input').addEventListener('focus', (e) => e.target.select());

// drag-over feedback on the empty drop zone (the shell performs the actual drop)
let dragDepth = 0;
document.addEventListener('dragenter', (e) => { if (e.dataTransfer?.types?.includes('Files')) { dragDepth++; document.body.classList.add('drag-over'); } });
document.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; document.body.classList.remove('drag-over'); } });
document.addEventListener('drop', () => { dragDepth = 0; document.body.classList.remove('drag-over'); }, true);

// restore view preferences
let pref = {};
try { pref = { mode: localStorage.getItem('bts-pdf-mode'), dark: localStorage.getItem('bts-pdf-darkread') }; } catch {}
if (pref.mode && ['single', 'continuous', 'spread'].includes(pref.mode)) S.mode = pref.mode;
$$('#ribbon [data-cmd="mode"]').forEach(b => b.classList.toggle('on', b.dataset.mode === S.mode));
if (pref.dark) toggleDarkRead(true);
if (sidePref === null || sidePref === 'thumbs' || sidePref === 'outline') SB.showSidebar(sidePref || 'thumbs');
setTool('text');
fillRecent();

// testing / debugging handle
window.BTSPDF = { S, V, A, Find, SB, IO, setTool, setMode, openBytes, toggleDarkRead, setFullscreen };
