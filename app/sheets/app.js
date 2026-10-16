/* BTS Sheets — application: shell integration, ribbon commands, sheet tabs, formula bar, status bar, context menus */
(function () {
  const App = window.App = {};
  const $ = (s, r = document) => r.querySelector(s);
  const { key } = SU;
  let wb;

  function newWorkbook() { wb.reset(); wb._addSheetRaw('Sheet1'); wb.activeIndex = 0; }

  function init() {
    wb = new SheetModel.Workbook(); App.wb = wb; window.WB = wb;
    newWorkbook();
    A.init(wb); Dlg.init(wb); Spark.init(wb); Pivot.init(wb); WhatIf.init(wb); DT.init(wb); Prot.init(wb); Threads.init(wb);
    Grid.init(wb, $('#grid-host'));
    Ed.init(wb); Charts.init(wb);
    wb.on((what) => { Grid.render(); App.refreshFormulaBar(); App.refreshStatus(); App.refreshRibbon(); if (what !== 'change') { Grid.updateSpacer(); renderTabs(); Charts.renderSheet(); } refreshUndo(); DT.afterChange(); Threads.render(); App.refreshContext(); Pivot.refreshPane(); });
    Grid.onSelect = () => { hideNote(); App.refreshFormulaBar(); App.refreshStatus(); App.refreshRibbon(); App.refreshValidationUI(); if (A.painter && !Grid.dragging) A.applyPainter(); Charts.selected && Charts.select(null); App.refreshContext(); Pivot.refreshPane(); Threads.render(); };
    Grid.onOutline = (o) => DT.onOutline(o);
    Grid.onFontsLoaded = () => { Charts.onFonts(); App.refreshFormulaBar(); };
    Grid.onSelectionEnd = () => { if (A.painter) A.applyPainter(); };
    Grid.onScroll = () => { Ed.reposition(); App.refreshValidationUI(); hideNote(); persistNotes(); };
    Grid.afterRender = () => { };
    Grid.positionObjects = () => Charts.position();
    Grid.onResize = (axis, idx, size) => A.setSize(axis, idx, size);
    Grid.onAutofit = (axis, idx) => A.autofit(axis, idx);
    Grid.onFill = (src, t) => A.fill(src, t, { ctrl: false });
    Grid.onFillDouble = () => A.fillDouble();
    Grid.onMoveCells = (src, t, copy) => { if (A.protectedCheck()) return; const s = wb.active; if (copy) { A.copy(false); s.sel = t; s.active = { r: t.r0, c: t.c0 }; A.pasteInternal('all'); A.clip = null; Grid.copyRange = null; } else { A.moveCells(s, src, s, t.r0, t.c0); s.sel = t; s.active = { r: t.r0, c: t.c0 }; } Grid.notify(); };
    Grid.onCellDblClick = (r, c, e) => { if (A.protectedCheck()) return; Ed.start(null, 'edit'); };
    Grid.onContextMenu = (h, e) => App.contextMenu(h, e);
    Grid.onFilterButton = (h) => Dlg.filterPopover(h);
    Grid.onHoverCell = onHoverCell;
    Grid.onMouseDownHook = (e, h) => {
      // ctrl+click on links, click on links
      if (h.zone === 'cell' && !Grid.editing && (e.ctrlKey || e.metaKey) === false) {
        const s = wb.active; const lk = s.links.get(key(h.r, h.c)); const hl = !lk && s.isFormula(h.r, h.c) ? (() => { try { return wb.hf.getCellHyperlink(s.a(h.r, h.c)); } catch { return null; } })() : null;
        if ((lk || hl) && e.detail === 1) { const p = Grid.local(e); const rc = Grid.cellRect(h.r, h.c); const tw = Grid.measure(Grid.fontStr(s.style(h.r, h.c), Grid.z), s.displayText(h.r, h.c)); if (p.x - rc.x < tw + 8) { Grid.select(h.r, h.c); setTimeout(() => Dlg.followLink(lk || hl), 0); return true; } }
      }
      return false;
    };
    buildRibbonControls();
    wireFormulaBar(); wireTabs();
    BTS.init({
      app: 'sheets', name: 'BTS Sheets', letter: 'S', defaultTitle: 'Book',
      saveFormats: [
        { ext: 'xlsx', label: 'Excel workbook (.xlsx)', icon: 'file-spreadsheet', hint: 'Best for sharing with spreadsheet apps — keeps formulas, formatting and sheets' },
        { ext: 'csv', label: 'CSV — comma separated (.csv)', icon: 'file-text', hint: 'Plain values of the current sheet only' },
        { ext: 'ods', label: 'OpenDocument spreadsheet (.ods)', icon: 'file-spreadsheet', hint: 'For LibreOffice — keeps formulas, number formats, formatting, merged cells and sheets (not charts or pictures)' },
        { ext: 'html', label: 'Web page (.html)', icon: 'file-code', hint: 'All sheets as formatted tables' },
      ],
      openAccept: '.xlsx,.xls,.csv,.tsv,.ods,.txt',
      onCommand: (cmd, el, ev) => App.cmd(cmd, el, ev),
      load: async (f) => { hideRecover(); await IO.load(wb, f); await Promise.race([STheme.preloadWorkbook(wb), new Promise(r => setTimeout(r, 1500))]); afterLoad(); },
      save: async (ext) => {
        if (Ed.isEditing()) Ed.commit();
        if (ext === 'csv') return IO.toCSV(wb.active);
        if (ext === 'ods') { const out = await IO.toODS(wb); odsNotice(); return out; }
        if (ext === 'html') return IO.toHTML(wb);
        return IO.toXLSX(wb);
      },
      newDocument: () => { newWorkbook(); afterLoad(); BTS.setTitle('Book'); BTS.markClean(); },
      exportPDF: async () => { if (Ed.isEditing()) Ed.commit(); return IO.preparePrint(wb, printSheets()); },
      afterExportPDF: () => IO.endPrint(),
      onZoom: (z) => { Grid.setZoom(z); Ed.reposition(); },
      zoom: 100, minZoom: 25, maxZoom: 400,
      templates: STemplates.list.map(t => ({ name: t.name, thumbHTML: t.thumbHTML, apply: () => { hideRecover(); BTS.resetProps(); wb.resetBook(); const d = t.build(); IO.buildWorkbook(wb, d.sheets, d.names || []); if (d.theme) wb.theme = d.theme; if (d.after) d.after(wb); wb.undoStack = []; wb.redoStack = []; STheme.preloadWorkbook(wb).then(() => { Grid.render(); Charts.onFonts(); }); afterLoad(); } })),
      blankThumb: `<div style="width:100%;height:100%;background:#fff;display:grid;grid-template-columns:18px repeat(4,1fr);grid-auto-rows:16px;font:9px Carlito,sans-serif;color:#7a8499">${Array.from({ length: 55 }, (_, i) => { const r = Math.floor(i / 5), c = i % 5; return `<div style="border-right:1px solid #e1e4e8;border-bottom:1px solid #e1e4e8;${r === 0 || c === 0 ? 'background:#f4f6f8;text-align:center' : ''}${r === 1 && c === 1 ? ';outline:2px solid #0F8A6A;outline-offset:-2px' : ''}">${r === 0 && c ? 'ABCD'[c - 1] : c === 0 && r ? r : ''}</div>`; }).join('')}</div>`,
      search: (q) => Dlg.searchFromTitlebar(q),
      infoHTML: () => infoHTML(),
      focus: () => Grid.focus(),
      ready: () => { if (!BTS.loadedExt) { checkRecovery(); /* a window that crashed moments ago still looks alive: look again once its heartbeat is stale */ setTimeout(() => { if (!BTS.loadedExt && $('#recover-bar').classList.contains('hidden')) checkRecovery(); }, 16500); } SAuto.sweep(); },
      shortcuts: [['Ctrl+Alt+M', 'New comment'], ['Ctrl+Shift+F', 'Font picker'], ['Alt+Shift+→ / ←', 'Group / ungroup'], ['Ctrl+Arrow', 'Jump to edge of data'], ['Shift+Arrow', 'Extend selection'], ['F2', 'Edit cell'], ['F4', 'Toggle $ in formula / repeat'], ['Ctrl+D / Ctrl+R', 'Fill down / right'], ['Ctrl+;', 'Insert today’s date'], ['Ctrl+`', 'Show formulas'], ['Alt+=', 'AutoSum'], ['Ctrl+Shift+L', 'Filter'], ['Ctrl+G', 'Go to'], ['Ctrl+PgUp / PgDn', 'Previous / next sheet'], ['Shift+F11', 'New sheet'], ['Ctrl+E', 'Flash fill'], ['Alt+Enter', 'New line in cell']],
    });
    /* autosave: every change marks the file dirty → schedule a snapshot; a save / new file / clean close removes it */
    SAuto.init(wb);
    const md = BTS.markDirty; BTS.markDirty = () => { md(); SAuto.schedule(); };
    const mc = BTS.markClean; BTS.markClean = () => { mc(); SAuto.clear(); };
    window.addEventListener('beforeunload', () => { if (!BTS.dirty) SAuto.clear(); });
    BTS.hasContent = () => wb.sheets.length > 1 || wb.sheets.some(s => s.usedRange().r1 >= 0 || s.charts.length || s.images.length) || wb.undoStack.length > 0;
    A.setMode('Ready');
    afterLoad(true);
    window.addEventListener('resize', () => Ed.reposition());
    document.addEventListener('keydown', (e) => { if ((e.ctrlKey || e.metaKey) && !e.altKey && document.activeElement && document.activeElement.closest && document.activeElement.closest('#fbar,#cell-editor') && (e.key === 'z' || e.key === 'y')) { /* handled by editor */ } }, true);
    // global undo/redo when focus is on ribbon/body
    document.addEventListener('keydown', (e) => {
      const a = document.activeElement; if (a && (a.tagName === 'INPUT' || a.tagName === 'TEXTAREA' || a.isContentEditable) && a.id !== 'gkeys') return;
      if (a && a.id === 'gkeys') return;
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); App.cmd('undo'); }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') { e.preventDefault(); App.cmd('redo'); }
    });
    document.addEventListener('mousedown', (e) => { if (Charts.selected && !e.target.closest('.fobj') && !e.target.closest('#ribbon') && !e.target.closest('#ribbon-tabs') && !e.target.closest('.pop') && !e.target.closest('.modal-back')) Charts.select(null); }, true);
    window.addEventListener('beforeprint', () => { if (!document.body.classList.contains('printing')) IO.preparePrint(wb, printSheets()); });
    $('#statusbar').addEventListener('contextmenu', (e) => { e.preventDefault(); statusMenu(e); });
    $('#tbl-gallery').addEventListener('mousedown', (e) => e.preventDefault());
    $('#tbl-gallery').addEventListener('click', (e) => { const b = e.target.closest('[data-id]'); if (b) A.formatAsTable(b.dataset.id); });
    $('#tbl-name').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.target.blur(); Grid.focus(); } if (e.key === 'Escape') { App.refreshContext(); Grid.focus(); } });
  }
  function printSheets() { return [wb.active]; }

  /* ---------- crash recovery bar ---------- */
  function hideRecover() { const b = $('#recover-bar'); if (b) { b.classList.add('hidden'); b.innerHTML = ''; } }
  function checkRecovery() {
    const list = SAuto.orphans(); if (!list.length) return;
    const best = list[0]; const bar = $('#recover-bar');
    const when = new Date(best.t).toLocaleString('en-IN', { day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit' });
    bar.innerHTML = `${icon('history', 16)}<span>Recover unsaved workbook “${SU.esc(best.title || 'Book')}” from ${SU.esc(when)}?${list.length > 1 ? ` <span class="rec-more">(${list.length - 1} more after this)</span>` : ''}</span><span style="flex:1"></span><button class="btn primary" id="rec-yes">Recover</button><button class="btn" id="rec-no">Dismiss</button>`;
    bar.classList.remove('hidden'); App.layoutPanes && App.layoutPanes();
    $('#rec-yes').onclick = async () => {
      if ((BTS.dirty || BTS.hasContent()) && !(await BTS.confirm('Replace current workbook?', 'The recovered workbook will replace what is open now.', 'Recover'))) return;
      let rec = null; try { rec = await SAuto.load(best.sid); } catch (e) { console.warn(e); }
      if (!rec) { SAuto.discard(best.sid); hideRecover(); BTS.toast('Sorry, that copy couldn’t be read.', { error: true }); return; }
      try {
        if (Ed.isEditing()) Ed.cancel();
        const snap = SAuto.restore(wb, rec.json);
        BTS.setTitle(snap.title || 'Book'); BTS.filePath = null;
        await Promise.race([STheme.preloadWorkbook(wb), new Promise(r => setTimeout(r, 1500))]);
        afterLoad(); Grid.render(); Charts.onFonts && Charts.onFonts();
        await SAuto.discard(best.sid); BTS.markDirty(); SAuto.flush();
        hideRecover(); BTS.toast('Workbook recovered — save it to keep it');
        if (list.length > 1) setTimeout(checkRecovery, 400);
      } catch (e) { console.error(e); BTS.toast('Couldn’t recover: ' + (e.message || e), { error: true, ms: 5000 }); }
    };
    $('#rec-no').onclick = () => { SAuto.discard(best.sid); hideRecover(); if (list.length > 1) setTimeout(checkRecovery, 200); Grid.focus(); };
  }
  App.checkRecovery = checkRecovery;

  /* ---------- .ods: tell once per window what an .ods file can't hold ---------- */
  let odsWarned = false;
  function odsNotice() {
    if (odsWarned) return; const lost = SODS.lost(wb); if (!lost.length) return; odsWarned = true;
    setTimeout(() => BTS.dialog({ title: 'Saved as OpenDocument (.ods)', body: `<p style="margin:4px 0 8px;max-width:460px;line-height:1.5">Values, formulas, number formats, cell formatting, merged cells and column widths are saved. These parts of your workbook aren’t kept in .ods:</p><ul style="margin:0 0 8px;padding-left:20px;line-height:1.6">${lost.map(l => `<li>${SU.esc(l)}</li>`).join('')}</ul><p style="margin:0;max-width:460px;line-height:1.5;color:var(--ink-2)">Keep an .xlsx copy (File ▸ Save) if you need them.</p>`, buttons: [{ label: 'OK', primary: true, value: true }] }), 300);
  }

  function afterLoad(first) {
    if (Ed.isEditing()) Ed.cancel();
    Charts.select(null); Grid.copyRange = null; A.clip = null; Threads.close(true); Pivot.paneClosed = false;
    const s = wb.active; Grid.updateSpacer(); Grid.restoreScroll(); renderTabs(); Charts.renderSheet();
    App.refreshFormulaBar(); App.refreshStatus(); App.refreshRibbon(); refreshUndo(); App.refreshValidationUI(); App.refreshContext(); Pivot.refreshPane(); App.layoutPanes();
    Grid.focus();
  }
  App.afterLoad = afterLoad;

  /* ---------- sheet tabs ---------- */
  App.switchSheet = (i) => {
    if (i < 0 || i >= wb.sheets.length || i === wb.activeIndex) return;
    if (Ed.isEditing()) { if (Ed.pointToSheet(wb.sheets[i])) { wb.activeIndex = i; Grid.restoreScroll(); renderTabs(); Charts.renderSheet(); return; } if (!Ed.commit()) return; }
    Charts.select(null);
    wb.active.scroll = { x: Grid.scroller.scrollLeft, y: Grid.scroller.scrollTop };
    wb.activeIndex = i; Grid.restoreScroll(); renderTabs(); Charts.renderSheet();
    App.refreshFormulaBar(); App.refreshStatus(); App.refreshRibbon(); App.refreshValidationUI(); App.refreshContext(); Pivot.refreshPane(); Threads.render();
    const t = $('#tabs .stab.active'); t && t.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  };
  function renderTabs() {
    const el = $('#tabs');
    el.innerHTML = wb.sheets.map((s, i) => `<button class="stab ${i === wb.activeIndex ? 'active' : ''}" data-i="${i}" draggable="true" title="${SU.esc(s.name)}">${s.protected ? icon('lock', 12) : ''}<span>${SU.esc(s.name)}</span>${s.tabColor ? `<i class="tc" style="background:${s.tabColor}"></i>` : ''}</button>`).join('');
  }
  App.renderTabs = renderTabs;
  function wireTabs() {
    const el = $('#tabs');
    $('#add-sheet').innerHTML = icon('plus', 16);
    $$('.sb-nav button').forEach(b => b.innerHTML = icon(b.dataset.sbnav === 'prev' ? 'chevron-left' : 'chevron-right', 16));
    $('#add-sheet').addEventListener('click', () => App.cmd('sheet-add'));
    $('.sb-nav').addEventListener('click', (e) => { const b = e.target.closest('[data-sbnav]'); if (b) App.switchSheet(SU.clamp(wb.activeIndex + (b.dataset.sbnav === 'next' ? 1 : -1), 0, wb.sheets.length - 1)); });
    el.addEventListener('mousedown', (e) => { const t = e.target.closest('.stab'); if (!t || e.target.tagName === 'INPUT') return; if (Ed.isEditing()) e.preventDefault(); App.switchSheet(+t.dataset.i); });
    el.addEventListener('dblclick', (e) => { const t = e.target.closest('.stab'); if (t && !Prot.structureCheck()) renameTab(+t.dataset.i); });
    el.addEventListener('contextmenu', (e) => { e.preventDefault(); const t = e.target.closest('.stab'); if (!t) return; App.switchSheet(+t.dataset.i); tabMenu(e); });
    el.addEventListener('wheel', (e) => { el.scrollLeft += e.deltaY; }, { passive: true });
    let dragI = null;
    el.addEventListener('dragstart', (e) => { const t = e.target.closest('.stab'); if (!t) return; if (wb.protectStructure) { e.preventDefault(); return; } dragI = +t.dataset.i; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/x-bts-sheet', String(dragI)); });
    el.addEventListener('dragover', (e) => { if (dragI === null) return; e.preventDefault(); const t = e.target.closest('.stab'); $$('.stab', el).forEach(x => x.classList.remove('drop-before', 'drop-after')); if (t) { const r = t.getBoundingClientRect(); t.classList.add(e.clientX < r.left + r.width / 2 ? 'drop-before' : 'drop-after'); } });
    el.addEventListener('drop', (e) => { if (dragI === null) return; e.preventDefault(); e.stopPropagation(); const t = e.target.closest('.stab'); if (t) { const r = t.getBoundingClientRect(); let to = +t.dataset.i + (e.clientX < r.left + r.width / 2 ? 0 : 1); if (to > dragI) to--; if (to !== dragI) { wb.run('Move sheet', () => wb.moveSheet(wb.sheets[dragI], to)); renderTabs(); } } dragI = null; });
    el.addEventListener('dragend', () => { dragI = null; $$('.stab', el).forEach(x => x.classList.remove('drop-before', 'drop-after')); });
  }
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  function renameTab(i) {
    const t = $(`#tabs .stab[data-i="${i}"]`); if (!t) return; const s = wb.sheets[i];
    t.innerHTML = `<input value="${SU.esc(s.name)}" maxlength="31">`; const inp = t.querySelector('input'); inp.focus(); inp.select();
    let done = false;
    const finish = (ok) => { if (done) return; done = true; if (ok && inp.value.trim() && inp.value.trim() !== s.name) { try { wb.run('Rename sheet', () => wb.renameSheet(s, inp.value.trim())); } catch (e) { BTS.toast(e.message, { error: true, ms: 4000 }); } } renderTabs(); Grid.focus(); };
    inp.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Enter') finish(true); if (e.key === 'Escape') finish(false); });
    inp.addEventListener('blur', () => finish(true));
  }
  function tabMenu(e) {
    const s = wb.active; const i = wb.activeIndex; const lockS = !!wb.protectStructure;
    BTS.menu({ x: e.clientX, y: e.clientY - 10 }, lockS ? [{ head: 'Workbook structure is protected' }, { label: s.protected ? 'Unprotect sheet' : 'Protect sheet', icon: 'lock', action: () => App.cmd('protect') }, { label: 'Unprotect workbook…', icon: 'book-lock', action: () => App.cmd('protect-workbook') }] : [
      { label: 'Insert sheet', icon: 'plus', action: () => App.cmd('sheet-add') },
      { label: 'Rename', icon: 'pencil', action: () => renameTab(i) },
      { label: 'Duplicate', icon: 'copy', action: () => App.cmd('sheet-dup') },
      { label: 'Delete', icon: 'trash-2', disabled: wb.sheets.length < 2, action: () => App.cmd('sheet-del') },
      { sep: true },
      { label: 'Move left', icon: 'arrow-left', disabled: i === 0, action: () => { wb.run('Move sheet', () => wb.moveSheet(s, i - 1)); renderTabs(); } },
      { label: 'Move right', icon: 'arrow-right', disabled: i === wb.sheets.length - 1, action: () => { wb.run('Move sheet', () => wb.moveSheet(s, i + 1)); renderTabs(); } },
      { label: 'Tab color…', icon: 'palette', action: async () => { const c = await BTS.colorPicker({ x: e.clientX, y: e.clientY - 260 }, { none: 'No color' }); if (c === null) return; wb.run('Tab color', () => wb.setProp(s, 'tabColor', c === 'none' || c === 'auto' ? null : c)); renderTabs(); } },
      { sep: true },
      { label: s.protected ? 'Unprotect sheet' : 'Protect sheet', icon: s.protected ? 'lock-open' : 'lock', action: () => App.cmd('protect') },
    ]);
  }

  /* ---------- formula bar ---------- */
  function wireFormulaBar() {
    const nb = $('#namebox');
    nb.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); const v = nb.value.trim(); if (!v) return; const ok = Dlg.gotoRef(v); if (!ok && /^[A-Za-z_][A-Za-z0-9_.]*$/.test(v) && !SU.parseRange(v)) { try { const s = wb.active; wb.run('Define name', () => wb.addName(v, '=' + SU.quoteSheet(s.name) + '!' + SU.absRangeStr(A.eff(s.sel)))); BTS.toast('Defined name ' + v); } catch (er) { BTS.toast(er.message, { error: true }); } } Grid.focus(); }
      if (e.key === 'Escape') { App.refreshFormulaBar(); Grid.focus(); }
    });
    nb.addEventListener('focus', () => nb.select());
    $('#namebox-drop').addEventListener('click', (e) => {
      const names = wb.listNames(); const tables = wb.sheets.flatMap(s => s.tables.map(t => t.name));
      if (!names.length && !tables.length) { BTS.menu($('#namebox'), [{ head: 'No defined names yet' }, { label: 'Define name…', icon: 'tag', action: () => Dlg.defineName() }]); return; }
      BTS.menu($('#namebox'), [...names.map(n => ({ label: n.name, icon: 'tag', action: () => Dlg.gotoRef(n.name) })), ...tables.map(t => ({ label: t, icon: 'table', action: () => Dlg.gotoRef(t) })), { sep: true }, { label: 'Name manager…', icon: 'tags', action: () => Dlg.nameManager() }]);
    });
  }
  App.refreshFormulaBar = () => {
    const s = wb.active; if (!s) return; const g = s.sel; const a = s.active;
    if (document.activeElement !== $('#namebox')) {
      let label;
      if (Grid.isWholeCols(g) && !Grid.isWholeRows(g)) label = (g.c1 - g.c0 + 1) + 'C';
      else if (Grid.isWholeRows(g) && !Grid.isWholeCols(g)) label = (g.r1 - g.r0 + 1) + 'R';
      else if (Grid.dragging && Grid.dragging.kind === 'sel' && (g.r1 > g.r0 || g.c1 > g.c0)) label = (g.r1 - g.r0 + 1) + 'R × ' + (g.c1 - g.c0 + 1) + 'C';
      else { label = SU.addr(a.r, a.c); const nm = wb.listNames().find(n => { const r = SU.parseRef(SU.fromHF(n.expr)); return r && (!r.sheet || r.sheet === s.name) && r.range.r0 === g.r0 && r.range.c0 === g.c0 && r.range.r1 === g.r1 && r.range.c1 === g.c1; }); if (nm) label = nm.name; }
      $('#namebox').value = label;
    }
    if (!Ed.isEditing()) { const fx = $('#fx'); fx.value = s.protected && s.style(a.r, a.c).fh ? '' : s.editText(a.r, a.c); fx.style.height = '26px'; }
  };
  App.refreshStatus = () => {
    const s = wb.active; const g = s.sel; const single = A.sels().length === 1 && g.r0 === g.r1 && g.c0 === g.c1 || (s.mergeAt(g.r0, g.c0) && s.expandMerges({ r0: g.r0, c0: g.c0, r1: g.r0, c1: g.c0 }).r1 === g.r1 && s.expandMerges({ r0: g.r0, c0: g.c0, r1: g.r0, c1: g.c0 }).c1 === g.c1);
    let html = '';
    if (!single) {
      const st = A.selectionStats(); const f = (v) => SU.format(v, s.style(s.active.r, s.active.c).nf && !/[@]/.test(s.style(s.active.r, s.active.c).nf) ? s.style(s.active.r, s.active.c).nf : 'General').text;
      const fmt = (v) => { const t = f(+v.toPrecision(12)); return t; };
      const on = App.statusOpts(); const parts = [];
      const stat = (k, label, v) => { if (on[k]) parts.push(`<span class="sb-stat" data-stat="${k}" title="Right-click the status bar to choose what to show">${label}: <b>${SU.esc(v)}</b></span>`); };
      if (st.ncount) { stat('average', 'Average', fmt(st.avg)); stat('count', 'Count', String(st.count)); stat('ncount', 'Numerical Count', String(st.ncount)); stat('min', 'Min', fmt(st.min)); stat('max', 'Max', fmt(st.max)); stat('sum', 'Sum', fmt(st.sum)); }
      else if (st.count) stat('count', 'Count', String(st.count));
      html = parts.join('');
    }
    BTS.setStatus('stats', html);
    const fl = s.filter && Object.keys(s.filter.cols || {}).length ? `<span>${icon('funnel', 13)} Filter mode: ${s.filter.range.r1 - s.filter.range.r0 - s.filterHidden.size} of ${s.filter.range.r1 - s.filter.range.r0} records</span>` : '';
    BTS.setStatus('filter', fl);
    BTS.setStatusRight([wb.protectStructure ? icon('book-lock', 13) + ' Structure locked' : '', s.protected ? icon('lock', 13) + ' Protected' : '', STheme.cur().name !== 'BTS' ? 'Theme: ' + SU.esc(STheme.cur().name) : '', Grid.showFormulas ? 'Showing formulas' : '', wb.sheets.length > 1 ? `Sheet ${wb.activeIndex + 1} of ${wb.sheets.length}` : ''].filter(Boolean).join(' · '));
  };
  function refreshUndo() { const tb = $('#titlebar'); if (!tb) return; const u = tb.querySelector('[data-cmd="undo"]'), r = tb.querySelector('[data-cmd="redo"]'); if (u) { u.style.opacity = wb.canUndo() ? 1 : .45; u.title = wb.canUndo() ? 'Undo ' + wb.undoStack[wb.undoStack.length - 1].label + ' (Ctrl+Z)' : 'Undo (Ctrl+Z)'; } if (r) { r.style.opacity = wb.canRedo() ? 1 : .45; r.title = wb.canRedo() ? 'Redo ' + wb.redoStack[wb.redoStack.length - 1].label + ' (Ctrl+Y)' : 'Redo (Ctrl+Y)'; } }

  /* ---------- ribbon controls state ---------- */
  function buildRibbonControls() {
    /* the font box opens BTS.fontPicker (search, categories, previews, recent, theme fonts) */
    $('#sel-size').innerHTML = SU.SIZES.map(s => `<option>${s}</option>`).join('');
    $('#sel-numfmt').innerHTML = SU.FORMATS.map(f => `<option value="${f.id}">${f.label}</option>`).join('') + '<option value="custom" disabled>Custom</option><option value="more">More formats…</option>';
  }
  App.refreshRibbon = () => {
    const s = wb.active; if (!s) return; const a = s.active; const st = s.style(a.r, a.c);
    const tog = (id, on) => { const e = document.getElementById(id); if (e) e.classList.toggle('on', !!on); };
    tog('btn-bold', st.b); tog('btn-italic', st.i); tog('btn-underline', st.u); tog('btn-strike', st.s); tog('btn-wrap', st.wrap);
    const v = s.value(a.r, a.c); const h = st.h || '';
    tog('btn-left', h === 'left'); tog('btn-center', h === 'center'); tog('btn-right', h === 'right');
    tog('btn-vtop', st.v === 'top'); tog('btn-vmiddle', st.v === 'middle'); tog('btn-vbottom', !st.v || st.v === 'bottom');
    tog('btn-merge', !!s.mergeAt(a.r, a.c)); tog('btn-showf', Grid.showFormulas); tog('btn-filter', !!s.filter); tog('btn-protect', s.protected);
    tog('btn-freeze', s.freeze.r || s.freeze.c); tog('btn-showcomments', App.notesShown); tog('btn-protectwb', !!wb.protectStructure); tog('btn-cmpane', Threads.isOpen());
    const sf = $('#sel-font span'); const fname = st.font || SU.defFont; if (sf.textContent !== fname) { sf.textContent = fname; sf.style.fontFamily = SU.cssFont(fname).replace(/"/g, "'"); }
    const sz = $('#sel-size'); if (document.activeElement !== sz) { const val = String(st.sz || 11); if (![...sz.options].some(o => o.value === val)) sz.insertAdjacentHTML('beforeend', `<option>${val}</option>`); sz.value = val; }
    const nf = $('#sel-numfmt'); if (document.activeElement !== nf) nf.value = SU.formatId(st.nf);
    $('#chk-vgrid').checked = s.showGrid; $('#chk-vgrid2').checked = s.showGrid; $('#chk-vhead').checked = s.showHeadings; $('#chk-vfbar').checked = !document.body.classList.contains('no-fbar');
    $('#chk-pgrid').checked = !!s.page.gridlines; $('#chk-phead').checked = !!s.page.headings; $('#chk-fitw').checked = s.page.fit === 'width'; $('#chk-fitp').checked = s.page.fit === 'page';
    document.body.classList.toggle('protected', !!s.protected);
  };

  /* ---------- validation dropdown button ---------- */
  let dvBtn = null, tlBtn = null;
  App.totalMenu = () => {
    const s = wb.active; const a = s.active; const t = s.tableAt(a.r, a.c); if (!t || !t.total) return; const cur = (t.totals || {})[a.c] || 'none';
    const rc = Grid.cellRect(a.r, a.c); const hr = $('#grid-host').getBoundingClientRect();
    BTS.menu({ x: hr.left + rc.x, y: hr.top + rc.y + rc.h + 1 }, A.TOTAL_FNS.map(([k, l]) => ({ label: l, checked: k === cur, action: () => A.setTotalFn(t, a.c, k) })).concat([{ sep: true }, { label: 'More functions…', icon: 'square-function', action: () => Dlg.insertFunction('SUBTOTAL') }]));
  };
  App.refreshValidationUI = () => {
    const s = wb.active; const a = s.active; const dv = s.validationAt(a.r, a.c);
    if (!dvBtn) { dvBtn = document.createElement('button'); dvBtn.className = 'dv-btn'; dvBtn.innerHTML = icon('chevron-down', 12); dvBtn.addEventListener('mousedown', (e) => { e.preventDefault(); e.stopPropagation(); App.openValidationList(); }); Grid.objLayer.parentElement.appendChild(dvBtn); }
    if (!dv || dv.type !== 'list' || Ed.isEditing()) { dvBtn.style.display = 'none'; } else {
      const rc = Grid.cellRect(a.r, a.c); dvBtn.style.display = ''; dvBtn.style.left = rc.x + rc.w + 1 + 'px'; dvBtn.style.top = rc.y + rc.h - 18 + 'px';
      if (rc.x + rc.w < Grid.hw() || rc.y < Grid.hh()) dvBtn.style.display = 'none';
    }
    // input prompt
    if (dv && dv.prompt && !Ed.isEditing()) showNote(a.r, a.c, dv.prompt, dv.promptTitle || '');
    // table total row: function picker button
    const t = s.tableAt(a.r, a.c);
    if (!tlBtn) { tlBtn = document.createElement('button'); tlBtn.className = 'dv-btn'; tlBtn.title = 'Choose the total function'; tlBtn.innerHTML = icon('chevron-down', 12); tlBtn.addEventListener('mousedown', (e) => { e.preventDefault(); e.stopPropagation(); App.totalMenu(); }); Grid.objLayer.parentElement.appendChild(tlBtn); }
    if (t && t.total && a.r === t.range.r1 && a.c > t.range.c0 && !Ed.isEditing() && !(dv && dv.type === 'list')) { const rc = Grid.cellRect(a.r, a.c); tlBtn.style.display = ''; tlBtn.style.left = rc.x + rc.w + 1 + 'px'; tlBtn.style.top = rc.y + rc.h - 18 + 'px'; } else tlBtn.style.display = 'none';
  };
  App.openValidationList = () => {
    const s = wb.active; const a = s.active; const dv = s.validationAt(a.r, a.c); if (!dv || dv.type !== 'list') return;
    const items = A.listItems(s, dv); const rc = Grid.cellRect(a.r, a.c); const hr = $('#grid-host').getBoundingClientRect();
    const cur = s.displayText(a.r, a.c);
    BTS.menu({ x: hr.left + rc.x, y: hr.top + rc.y + rc.h + 1 }, items.length ? items.map(t => ({ label: t, checked: t === cur, action: () => { if (A.protectedCheck()) return; A.enterValue(s, a.r, a.c, t); Grid.focus(); } })) : [{ head: 'No items in the list' }]);
    const pop = document.querySelector('.pop'); if (pop) pop.style.minWidth = Math.max(140, rc.w) + 'px';
  };

  /* ---------- notes hover ---------- */
  let noteEl = null, noteKey = null;
  function showNote(r, c, text, title, persist) {
    if (!noteEl) { noteEl = document.createElement('div'); noteEl.className = 'note-pop'; $('#grid-host').appendChild(noteEl); }
    const rc = Grid.cellRect(r, c); noteEl.innerHTML = (title ? `<b>${SU.esc(title)}</b>` : '') + SU.esc(text);
    noteEl.style.left = rc.x + rc.w + 8 + 'px'; noteEl.style.top = Math.max(Grid.hh(), rc.y - 2) + 'px'; noteEl.style.display = 'block'; noteKey = key(r, c);
  }
  function hideNote() { if (noteEl) noteEl.style.display = 'none'; noteKey = null; }
  App.notesShown = false;
  function persistNotes() {
    $$('.note-pop.persist').forEach(e => e.remove());
    if (!App.notesShown) return; const s = wb.active; const host = $('#grid-host');
    for (const [k, t] of s.comments) { const { r, c } = SU.unkey(k); const v = Grid.visible; if (!v || r < v.r0 - 1 || r > v.r1 + 1 || c < v.c0 - 1 || c > v.c1 + 1) continue; const rc = Grid.cellRect(r, c); const el = document.createElement('div'); el.className = 'note-pop persist'; el.textContent = t; el.style.left = rc.x + rc.w + 8 + 'px'; el.style.top = rc.y + 'px'; host.appendChild(el); }
  }
  App.persistNotes = persistNotes;
  function onHoverCell(cell) {
    const s = wb.active;
    if (!cell) { if (noteKey !== null && !(s.validationAt(s.active.r, s.active.c) || {}).prompt) hideNote(); return; }
    const k = key(cell.r, cell.c);
    if (s.threads.has(k)) { if (noteKey !== k) { const p = Threads.preview(s, cell.r, cell.c); showNote(cell.r, cell.c, p.text, p.title); noteEl.classList.add('thread'); } }
    else if (s.comments.has(k) && !App.notesShown) { if (noteKey !== k) { showNote(cell.r, cell.c, s.comments.get(k), 'Note'); noteEl.classList.remove('thread'); } }
    else if (s.links.has(k)) { if (noteKey !== k) showNote(cell.r, cell.c, s.links.get(k) + '\nClick to follow the link.', ''); }
    else if (noteKey !== null && noteKey !== key(s.active.r, s.active.c)) hideNote();
  }

  /* ---------- context menu ---------- */
  App.contextMenu = (h, e) => {
    const s = wb.active; const hdrCol = h.zone === 'colhdr', hdrRow = h.zone === 'rowhdr';
    const k = key(s.active.r, s.active.c); const hasNote = s.comments.has(k), hasLink = s.links.has(k);
    const t = s.tableAt(s.active.r, s.active.c);
    const items = [
      { label: 'Cut', icon: 'scissors', kbd: 'Ctrl+X', action: () => A.cutCopyCmd(true) },
      { label: 'Copy', icon: 'copy', kbd: 'Ctrl+C', action: () => A.cutCopyCmd(false) },
      { label: 'Paste', icon: 'clipboard-paste', kbd: 'Ctrl+V', action: () => A.pasteCmd() },
      { label: 'Paste values', icon: 'clipboard-type', action: () => A.clip ? A.pasteInternal('values') : A.pasteCmd('values') },
      { label: 'Paste special…', icon: 'clipboard-list', action: () => Dlg.pasteSpecial() },
      { sep: true },
    ];
    if (hdrCol) items.push({ label: 'Insert columns', icon: 'between-vertical-start', action: () => A.insertCols() }, { label: 'Delete columns', icon: 'trash-2', action: () => A.deleteCols() }, { label: 'Clear contents', icon: 'eraser', action: () => A.clear('contents') }, { sep: true }, { label: 'Column width…', icon: 'move-horizontal', action: () => Dlg.size('col') }, { label: 'AutoFit width', icon: 'arrow-left-right', action: () => A.autofit('col') }, { label: 'Hide', icon: 'eye-off', action: () => A.setHidden('c', true) }, { label: 'Unhide', icon: 'eye', action: () => A.setHidden('c', false) });
    else if (hdrRow) items.push({ label: 'Insert rows', icon: 'between-horizontal-start', action: () => A.insertRows() }, { label: 'Delete rows', icon: 'trash-2', action: () => A.deleteRows() }, { label: 'Clear contents', icon: 'eraser', action: () => A.clear('contents') }, { sep: true }, { label: 'Row height…', icon: 'move-vertical', action: () => Dlg.size('row') }, { label: 'AutoFit height', icon: 'arrow-up-down', action: () => A.autofit('row') }, { label: 'Hide', icon: 'eye-off', action: () => A.setHidden('r', true) }, { label: 'Unhide', icon: 'eye', action: () => A.setHidden('r', false) });
    else {
      items.push({ label: 'Insert…', icon: 'square-plus', action: () => Dlg.insertCells() }, { label: 'Delete…', icon: 'square-minus', action: () => Dlg.deleteCells() }, { label: 'Clear contents', icon: 'eraser', kbd: 'Del', action: () => A.clear('contents') }, { sep: true },
        { label: 'Sort A to Z', icon: 'arrow-down-a-z', action: () => A.quickSort(false) }, { label: 'Sort Z to A', icon: 'arrow-down-z-a', action: () => A.quickSort(true) }, { label: s.filter ? 'Remove filter' : 'Filter', icon: 'funnel', action: () => A.toggleFilter() }, { sep: true },
        { label: s.threads.has(k) ? 'Reply to comment' : 'New comment', icon: 'message-square-plus', kbd: 'Ctrl+Alt+M', action: () => Threads.newComment() }, ...(s.threads.has(k) ? [{ label: 'Delete comment', icon: 'message-square-x', action: () => Threads.deleteAtActive() }] : []),
        { label: hasNote ? 'Edit note' : 'New note', icon: 'sticky-note', action: () => Dlg.note() }, ...(hasNote ? [{ label: 'Delete note', icon: 'eraser', action: () => App.cmd('comment-delete') }] : []),
        { label: 'Format cells…', icon: 'paintbrush', kbd: 'Ctrl+1', action: () => Dlg.formatCells() },
        ...(s.validationAt(s.active.r, s.active.c)?.type === 'list' ? [{ label: 'Pick from drop-down list', icon: 'list', kbd: 'Alt+↓', action: () => App.openValidationList() }] : []),
        { label: 'Define name…', icon: 'tag', action: () => Dlg.defineName() },
        { label: hasLink ? 'Edit link…' : 'Link…', icon: 'link', kbd: 'Ctrl+K', action: () => Dlg.hyperlink() }, ...(hasLink ? [{ label: 'Open link', icon: 'external-link', action: () => Dlg.followLink(s.links.get(k)) }] : []));
      if (t) items.push({ sep: true }, { head: 'Table: ' + t.name }, { label: t.total ? 'Remove total row' : 'Add total row', icon: 'sigma', action: () => A.toggleTotalRow(t) }, { label: 'Summarize with PivotTable…', icon: 'table-properties', action: () => Pivot.createDialog(t) }, { label: 'Convert to range', icon: 'table-2', action: () => A.convertTableToRange(t) });
      const pv = Pivot.at(s, s.active.r, s.active.c);
      if (pv) items.splice(0, items.length, { head: 'PivotTable: ' + pv.name }, { label: 'Copy', icon: 'copy', kbd: 'Ctrl+C', action: () => A.cutCopyCmd(false) }, { label: 'Refresh', icon: 'refresh-cw', action: () => Pivot.refresh(s, pv) }, { label: 'Show field list', icon: 'panel-right', action: () => { Pivot.paneClosed = false; Pivot.refreshPane(); } }, { label: 'Value field settings…', icon: 'settings-2', disabled: !pv.values.length, action: () => Pivot.valueSettings(s, pv, Math.max(0, Math.min(pv.values.length - 1, (s.active.c - pv.out.c0 - Math.max(pv.rows.length, 1)) % Math.max(1, pv.values.length)))) }, { label: 'PivotTable options…', icon: 'settings-2', action: () => Pivot.options() }, { label: 'PivotChart', icon: 'chart-column-big', action: () => Pivot.createChart() }, { sep: true }, { label: 'Delete PivotTable', icon: 'trash-2', action: () => Pivot.remove(pv) });
      if (Grid.sparkAt(s, s.active.r, s.active.c)) items.push({ sep: true }, { label: 'Clear sparklines', icon: 'eraser', action: () => Spark.clear(false) }, { label: 'Edit sparkline data…', icon: 'table-2', action: () => Spark.editData() });
    }
    BTS.menu({ x: e.clientX, y: e.clientY }, items);
  };

  /* ---------- File › Info ---------- */
  function infoHTML() {
    let cells = 0, formulas = 0; for (const s of wb.sheets) { const u = s.usedRange(); for (let r = 0; r <= u.r1; r++) for (let c = 0; c <= u.c1; c++) if (!s.isEmpty(r, c)) { cells++; if (s.isFormula(r, c)) formulas++; } }
    return `<div class="bs-actions" style="max-width:640px"><div class="bs-action" style="cursor:default">${icon('file-spreadsheet')}<div><b>${SU.esc(BTS.title)}</b><small>${SU.esc(BTS.filePath || 'Not saved yet')}</small></div></div></div>
      <h2>Properties</h2><div style="display:grid;grid-template-columns:180px 1fr;gap:8px 16px;max-width:560px;color:var(--ink-2)">
      <span>Sheets</span><b style="color:var(--ink)">${wb.sheets.length} (${wb.sheets.map(s => SU.esc(s.name)).join(', ')})</b>
      <span>Cells with data</span><b style="color:var(--ink)">${cells}</b><span>Formulas</span><b style="color:var(--ink)">${formulas}</b>
      <span>Charts</span><b style="color:var(--ink)">${wb.sheets.reduce((a, s) => a + s.charts.length, 0)}</b><span>Defined names</span><b style="color:var(--ink)">${wb.listNames().length}</b>
      <span>PivotTables</span><b style="color:var(--ink)">${wb.sheets.reduce((a, s) => a + s.pivots.length, 0)}</b><span>Comments</span><b style="color:var(--ink)">${wb.sheets.reduce((a, s) => a + s.threads.size, 0)}</b>
      <span>Theme</span><b style="color:var(--ink)">${SU.esc(STheme.cur().name)} (${SU.esc(STheme.cur().head)} / ${SU.esc(STheme.cur().body)})</b>
      <span>Calculation</span><b style="color:var(--ink)">Automatic</b></div>
      <h2>Protect</h2><div class="bs-actions"><button class="bs-action" onclick="BTS.closeBackstage();App.cmd('protect')">${icon('lock')}<div><b>${wb.active.protected ? 'Unprotect' : 'Protect'} current sheet</b><small>Control what people can change, with an optional password</small></div></button><button class="bs-action" onclick="BTS.closeBackstage();App.cmd('protect-workbook')">${icon('book-lock')}<div><b>${wb.protectStructure ? 'Unprotect' : 'Protect'} workbook structure</b><small>Stop sheets being added, deleted, renamed or moved</small></div></button></div>`;
  }

  /* ---------- menus ---------- */
  const borderIcon = (k) => {
    const L = (x1, y1, x2, y2, w = 1.6, dash) => `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="currentColor" stroke-width="${w}" ${dash ? 'stroke-dasharray="1.5 2"' : ''}/>`;
    const grid = L(8, 2, 8, 14, .8, 1) + L(2, 8, 14, 8, .8, 1) + L(2, 2, 14, 2, .8, 1) + L(2, 14, 14, 14, .8, 1) + L(2, 2, 2, 14, .8, 1) + L(14, 2, 14, 14, .8, 1);
    const parts = { all: L(2, 2, 14, 2) + L(2, 14, 14, 14) + L(2, 2, 2, 14) + L(14, 2, 14, 14) + L(8, 2, 8, 14) + L(2, 8, 14, 8), outside: L(2, 2, 14, 2) + L(2, 14, 14, 14) + L(2, 2, 2, 14) + L(14, 2, 14, 14), 'thick-outside': L(2, 2, 14, 2, 3) + L(2, 14, 14, 14, 3) + L(2, 2, 2, 14, 3) + L(14, 2, 14, 14, 3), inside: L(8, 2, 8, 14) + L(2, 8, 14, 8), top: L(2, 2, 14, 2, 2), bottom: L(2, 14, 14, 14, 2), left: L(2, 2, 2, 14, 2), right: L(14, 2, 14, 14, 2), 'thick-bottom': L(2, 14, 14, 14, 3), 'double-bottom': L(2, 12.5, 14, 12.5, 1) + L(2, 15, 14, 15, 1), 'top-bottom': L(2, 2, 14, 2, 2) + L(2, 14, 14, 14, 2), 'top-thick-bottom': L(2, 2, 14, 2, 1.5) + L(2, 14, 14, 14, 3), none: '' };
    return `<svg class="bicon" viewBox="0 0 16 16" width="16" height="16" fill="none">${grid}${parts[k] || ''}</svg>`;
  };
  function bordersMenu(el) {
    const opt = (k, l) => ({ label: l, html: borderIcon(k) + `<span>${l}</span>`, action: () => { A.applyBorders(k); lastBorder = k; } });
    BTS.menu(el, [{ head: 'Borders' }, opt('bottom', 'Bottom border'), opt('top', 'Top border'), opt('left', 'Left border'), opt('right', 'Right border'), { sep: true }, opt('none', 'No border'), opt('all', 'All borders'), opt('outside', 'Outside borders'), opt('thick-outside', 'Thick outside borders'), opt('inside', 'Inside borders'), { sep: true }, opt('double-bottom', 'Bottom double border'), opt('thick-bottom', 'Thick bottom border'), opt('top-bottom', 'Top and bottom border'), opt('top-thick-bottom', 'Top and thick bottom border'), { sep: true },
      { label: 'Line color…', icon: 'palette', html: `<span style="width:16px;height:16px;border-radius:3px;background:${A.borderLine.color};border:1px solid var(--line)"></span><span>Line color…</span>`, action: async () => { const c = await BTS.colorPicker(el, { auto: 'Automatic' }); if (c) A.borderLine.color = c === 'auto' ? '#000000' : c; } },
      { label: 'Line style', icon: 'minus', action: () => BTS.menu(el, [['thin', 'Thin'], ['medium', 'Medium'], ['thick', 'Thick'], ['dashed', 'Dashed'], ['dotted', 'Dotted'], ['double', 'Double']].map(([k, l]) => ({ label: l, checked: A.borderLine.style === k, action: () => { A.borderLine.style = k; } }))) }]);
  }
  let lastBorder = 'all';
  let lastFill = '#ffff00', lastFont = '#d9262f';
  function cfMenu(el) {
    BTS.menu(el, [
      { head: 'Highlight cells rules' },
      { label: 'Greater than…', icon: 'chevron-right', action: () => Dlg.cfRule('gt') }, { label: 'Less than…', icon: 'chevron-left', action: () => Dlg.cfRule('lt') }, { label: 'Between…', icon: 'chevrons-left-right', action: () => Dlg.cfRule('between') }, { label: 'Equal to…', icon: 'equal', action: () => Dlg.cfRule('eq') }, { label: 'Text that contains…', icon: 'type', action: () => Dlg.cfRule('text') }, { label: 'Duplicate values', icon: 'copy', action: () => Dlg.cfRule('dup') },
      { head: 'Top / bottom rules' },
      { label: 'Top 10 items…', icon: 'arrow-up-narrow-wide', action: () => Dlg.cfRule('top') }, { label: 'Bottom 10 items…', icon: 'arrow-down-narrow-wide', action: () => Dlg.cfRule('bottom') }, { label: 'Above average', icon: 'trending-up', action: () => Dlg.cfRule('aboveavg') }, { label: 'Below average', icon: 'trending-down', action: () => Dlg.cfRule('belowavg') },
      { sep: true },
      { label: 'Data bars', icon: 'chart-bar-big', action: () => BTS.menu(el, [['#5B9BD5', 'Blue'], ['#63BE7B', 'Green'], ['#F8696B', 'Red'], ['#FFB628', 'Orange'], ['#0F8A6A', 'Teal'], ['#8E7CC3', 'Purple']].map(([c, l]) => ({ label: l + ' data bar', html: `<span style="width:30px;height:12px;background:linear-gradient(90deg,${c},#fff);border:1px solid ${c}"></span><span>${l} data bar</span>`, action: () => A.addCF({ type: 'bar', color: c }) }))) },
      { label: 'Color scales', icon: 'palette', action: () => BTS.menu(el, [[['#F8696B', '#FFEB84', '#63BE7B'], 'Red – Yellow – Green'], [['#63BE7B', '#FFEB84', '#F8696B'], 'Green – Yellow – Red'], [['#F8696B', '#FFFFFF', '#5A8AC6'], 'Red – White – Blue'], [['#FFFFFF', '#63BE7B'], 'White – Green'], [['#FFFFFF', '#F8696B'], 'White – Red'], [['#FFEF9C', '#63BE7B'], 'Yellow – Green']].map(([cols, l]) => ({ label: l, html: `<span style="width:30px;height:12px;background:linear-gradient(90deg,${cols.join(',')})"></span><span>${l}</span>`, action: () => A.addCF({ type: 'scale', colors: cols }) }))) },
      { label: 'New rule with a formula…', icon: 'square-function', action: () => Dlg.cfRule('formula') },
      { sep: true },
      { label: 'Clear rules from selected cells', icon: 'eraser', action: () => A.clearCF('sel') }, { label: 'Clear rules from entire sheet', icon: 'trash-2', action: () => A.clearCF('sheet') }, { label: 'Manage rules…', icon: 'list-checks', action: () => Dlg.cfManager() },
    ]);
  }
  function tableMenu(el) {
    const pop = document.createElement('div'); const th = STheme.cur(); const t = wb.active.tableAt(wb.active.active.r, wb.active.active.c);
    pop.innerHTML = ['Light', 'Medium', 'Dark'].map(fam => `<div class="gal-head">${fam}</div><div class="gal tables">${STheme.TABLE_STYLES.filter(x => x.startsWith(fam)).map(id => `<button title="Table style ${id}" data-id="${id}" class="${t && t.styleId === id ? 'on' : ''}">${STheme.swatchHTML(id, th)}</button>`).join('')}</div>`).join('') + `<div class="gal-foot">${t ? 'Click a style to restyle ' + SU.esc(t.name) : 'Select a cell in your data, then pick a style'}</div>`;
    pop.addEventListener('mousedown', e => e.preventDefault());
    pop.addEventListener('click', e => { const b = e.target.closest('[data-id]'); if (b) { BTS.closePopover(); A.formatAsTable(b.dataset.id); } });
    BTS.popover(el, pop);
  }
  function stylesMenu(el) {
    const pop = document.createElement('div'); pop.className = 'styles-pop';
    const groups = STheme.cellStyleGroups(STheme.cur()).map(g => ({ ...g, items: g.items.map(([n]) => [n, STheme.cellStyle(n)]) }));
    if (wb.customStyles.length) groups.unshift({ group: 'Custom', items: wb.customStyles.map(c => [c.name, c.st]) });
    const css = (st, num) => { if (num) return 'background:#fff'; const c = A.cssOfStyle(st).replace(/font-size:[^;]+;?/, '').replace(/font-family:([^;]+)/, (m, f) => 'font-family:' + SU.cssFont(f.trim()).replace(/"/g, "'")); return (c || 'background:#fff') + (st.fill ? '' : ';background:#fff'); };
    pop.innerHTML = groups.map((g, gi) => `<div class="gal-head">${g.group}</div><div class="gal styles ${g.cols === 6 ? 'six' : ''}">${g.items.map(([n, st], i) => `<button data-g="${gi}" data-i="${i}" title="${SU.esc(n)} — right-click to modify" style="${css(st || {}, g.numOnly)};color:${(st && st.color) || '#000'}">${SU.esc(n)}</button>`).join('')}</div>`).join('') +
      `<div class="menu-sep"></div><div class="menu-item" data-a="new">${icon('plus', 16)}<span>New cell style…</span></div><div class="menu-item" data-a="modify">${icon('pencil', 16)}<span>Modify a style…</span></div>`;
    pop.addEventListener('mousedown', e => e.preventDefault());
    pop.addEventListener('click', e => { const b = e.target.closest('[data-i]'); if (b) { BTS.closePopover(); A.applyCellStyle(groups[+b.dataset.g].items[+b.dataset.i][0]); return; } const a = e.target.closest('[data-a]'); if (a) { BTS.closePopover(); if (a.dataset.a === 'new') App.newCellStyle(); else App.modifyCellStyle(); } });
    pop.addEventListener('contextmenu', e => { const b = e.target.closest('[data-i]'); if (!b) return; e.preventDefault(); const name = groups[+b.dataset.g].items[+b.dataset.i][0]; BTS.closePopover(); App.modifyCellStyle(name); });
    BTS.popover(el, pop);
  }
  /** style editor shared by New and Modify */
  async function styleEditor(title, name, st, nameEditable) {
    const body = document.createElement('div'); let cur = { ...st };
    const fonts = ['', ...STheme.fontNames()];
    body.innerHTML = `<div class="dlg-grid" style="grid-template-columns:110px 1fr">
      <label>Style name</label><input id="cs-name" value="${SU.esc(name)}" ${nameEditable ? '' : 'disabled'}>
      <label>Font</label><select id="cs-font">${fonts.map(f => `<option value="${SU.esc(f)}" ${(cur.font || '') === f ? 'selected' : ''}>${f || '(Body font)'}</option>`).join('')}</select>
      <label>Size</label><input id="cs-sz" type="number" min="6" max="96" value="${cur.sz || 11}">
      <label>Font style</label><div><label class="check" style="display:inline-flex;margin-right:10px"><input type="checkbox" id="cs-b" ${cur.b ? 'checked' : ''}> Bold</label><label class="check" style="display:inline-flex;margin-right:10px"><input type="checkbox" id="cs-i" ${cur.i ? 'checked' : ''}> Italic</label><label class="check" style="display:inline-flex"><input type="checkbox" id="cs-u" ${cur.u ? 'checked' : ''}> Underline</label></div>
      <label>Font colour</label><input id="cs-color" type="color" value="${cur.color || '#000000'}">
      <label>Fill</label><div style="display:flex;gap:8px;align-items:center"><input id="cs-fill" type="color" value="${cur.fill || '#ffffff'}"><label class="check"><input type="checkbox" id="cs-nofill" ${cur.fill ? '' : 'checked'}> No fill</label></div>
      <label>Number format</label><input id="cs-nf" value="${SU.esc(cur.nf || 'General')}">
      <label>Preview</label><div id="cs-prev" class="cs-prev">AaBbCc 1,234</div></div>`;
    const read = () => { const q = (x) => body.querySelector(x); const o = { ...cur }; o.font = q('#cs-font').value || undefined; o.sz = +q('#cs-sz').value === 11 ? undefined : +q('#cs-sz').value || undefined; o.b = q('#cs-b').checked || undefined; o.i = q('#cs-i').checked || undefined; o.u = q('#cs-u').checked || undefined; o.color = q('#cs-color').value.toUpperCase() === '#000000' ? undefined : q('#cs-color').value; o.fill = q('#cs-nofill').checked ? undefined : q('#cs-fill').value; o.nf = q('#cs-nf').value === 'General' ? undefined : q('#cs-nf').value; Object.keys(o).forEach(k => o[k] === undefined && delete o[k]); return o; };
    const prev = () => { const o = read(); const el = body.querySelector('#cs-prev'); el.style.cssText = A.cssOfStyle(o).replace(/font-family:([^;]+)/, (m, f) => 'font-family:' + SU.cssFont(f.trim()).replace(/"/g, "'")); };
    body.addEventListener('input', prev); body.addEventListener('change', prev); setTimeout(prev, 0);
    const res = await BTS.dialog({ title, width: 440, body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: () => ({ name: body.querySelector('#cs-name').value.trim(), st: read() }) }] });
    return res;
  }
  App.newCellStyle = async () => {
    const s = wb.active; const base = { ...s.style(s.active.r, s.active.c) }; ['h', 'v', 'wrap', 'indent', 'locked', 'fh', 'cs', 'pv'].forEach(k => delete base[k]);
    let n = 1; while (wb.customStyles.some(c => c.name === 'Style ' + n)) n++;
    const res = await styleEditor('New cell style', 'Style ' + n, base, true); if (!res) return;
    if (!res.name) { BTS.toast('Enter a style name.', { error: true }); return; }
    if (STheme.cellStyle(res.name) && !wb.customStyles.some(c => c.name === res.name)) { BTS.toast('A built-in style already has that name.', { error: true }); return; }
    wb.run('New cell style', () => wb.setProp(wb, 'customStyles', [...wb.customStyles.filter(c => c.name !== res.name), { name: res.name, st: res.st }]));
    A.applyCellStyle(res.name); BTS.toast('Created cell style “' + res.name + '”');
  };
  App.modifyCellStyle = async (name) => {
    if (!name) { const all = [...wb.customStyles.map(c => c.name), ...STheme.cellStyleGroups().flatMap(g => g.items.map(i => i[0]))].filter((x, i, a) => a.indexOf(x) === i); name = await BTS.dialog({ title: 'Modify cell style', width: 340, body: `<div class="field"><label>Style</label><select id="ms-n">${all.map(n => `<option>${SU.esc(n)}</option>`).join('')}</select></div>`, buttons: [{ label: 'Cancel', value: null }, { label: 'Next', primary: true, value: (r) => r.querySelector('#ms-n').value }] }); if (!name) return; }
    const res = await styleEditor('Modify style: ' + name, name, STheme.cellStyle(name) || {}, false); if (!res) return;
    const custom = wb.customStyles.some(c => c.name === name);
    wb.run('Modify cell style', () => { if (custom) wb.setProp(wb, 'customStyles', wb.customStyles.map(c => c.name === name ? { name, st: res.st } : c)); else wb.setProp(wb, 'styleMods', { ...wb.styleMods, [name]: res.st }); A.restyleTagged(name); });
    BTS.toast('Updated “' + name + '” and every cell that uses it');
  };
  function themeMenu(el, mode) {
    const pop = document.createElement('div'); pop.className = 'theme-pop';
    const cur = STheme.cur();
    pop.innerHTML = `<div class="gal-head">${mode === 'colors' ? 'Theme colours' : mode === 'fonts' ? 'Theme fonts' : 'Workbook themes'}</div><div class="theme-grid ${mode === 'fonts' ? 'fonts' : ''}">${STheme.THEMES.map(t => {
      const sw = `<div class="th-sw">${t.accents.map(a => `<i style="background:${a}"></i>`).join('')}</div>`;
      const on = mode === 'colors' ? t.accents.join() === cur.accents.join() : mode === 'fonts' ? (t.head === cur.head && t.body === cur.body) : t.id === wb.theme && !wb.themeOver;
      if (mode === 'fonts') return `<button data-id="${t.id}" class="${on ? 'on' : ''}" title="${SU.esc(t.head)} / ${SU.esc(t.body)}"><b style="font-family:${SU.cssFont(t.head).replace(/"/g, "'")}">${SU.esc(t.head)}</b><span style="font-family:${SU.cssFont(t.body).replace(/"/g, "'")}">${SU.esc(t.body)}${t.num ? ' · ' + SU.esc(t.num) : ''}</span></button>`;
      return `<button data-id="${t.id}" class="${on ? 'on' : ''}" title="${SU.esc(t.name)} — ${SU.esc(t.desc)}"><div class="th-prev" style="font-family:${SU.cssFont(t.head).replace(/"/g, "'")}"><span style="color:${t.dk2}">Aa</span>${sw}</div><small>${SU.esc(t.name)}</small></button>`;
    }).join('')}</div><div class="gal-foot">Themes set the fonts and colours used by cell styles, tables, charts and PivotTables.</div>`;
    pop.addEventListener('mousedown', e => e.preventDefault());
    pop.addEventListener('click', e => { const b = e.target.closest('[data-id]'); if (!b) return; BTS.closePopover(); const t = STheme.get(b.dataset.id);
      if (!mode) return A.setTheme(t.id);
      const over = { ...(wb.themeOver || {}) }; if (mode === 'colors') Object.assign(over, { accents: t.accents, dk2: t.dk2, lt2: t.lt2 }); else Object.assign(over, { head: t.head, body: t.body, num: t.num || null });
      [over.head, over.body].filter(Boolean).forEach(f => STheme.loadFont(f, true));
      wb.run(mode === 'colors' ? 'Theme colours' : 'Theme fonts', () => { wb.setProp(wb, 'themeOver', over); A.restyleTagged(); Pivot.restyleAll(); }); Grid.render(); Charts.renderSheet(); App.refreshStatus();
    });
    pop.addEventListener('mouseover', e => { const b = e.target.closest('[data-id]'); if (b && !mode) { /* live preview via status */ } });
    BTS.popover(el, pop);
  }
  function chartMenu(el, list) { BTS.menu(el, list.map(k => { const t = Charts.TYPES.find(x => x[0] === k); return { label: t[1], icon: t[2], action: () => App.cmd('chart-' + k) }; })); }
  function chartTypeMenu(el) { const f = Charts.find(Charts.selected); BTS.menu(el, Charts.TYPES.map(([k, l, ic]) => ({ label: l, icon: ic, checked: f && f.obj.type === k ? true : undefined, action: () => Charts.update({ type: k }, 'Chart type') }))); }

  /* ---------- command dispatcher ---------- */
  App.cmd = async (cmd, el, ev) => {
    const s = wb.active; const editing = Ed.isEditing();
    const needsCommit = !['undo', 'redo', 'insert-function', 'fn-cat', 'symbol', 'use-name', 'zoom-dialog', 'zoom-sel', 'new-window', 'show-formulas', 'view-gridlines', 'view-headings', 'view-fbar', 'font', 'size', 'font-pick'].includes(cmd);
    if (editing && needsCommit && !cmd.startsWith('chart-')) { if (!Ed.commit()) return; }
    if (editing && (cmd === 'font' || cmd === 'size' || cmd === 'font-pick')) Ed.commit();
    if (['sheet-add', 'sheet-dup', 'sheet-del'].includes(cmd) && Prot.structureCheck()) return;
    const anchor = el && el.closest ? el.closest('button,select,label') || el : el;
    switch (cmd) {
      /* clipboard & history */
      case 'undo': if (Ed.isEditing()) { Ed.cancel(); return; } wb.undo(); Grid.updateSpacer(); renderTabs(); Charts.renderSheet(); Grid.notify(); return;
      case 'redo': if (Ed.isEditing()) return; wb.redo(); Grid.updateSpacer(); renderTabs(); Charts.renderSheet(); Grid.notify(); return;
      case 'cut': return A.cutCopyCmd(true);
      case 'copy': return A.cutCopyCmd(false);
      case 'paste': return A.pasteCmd();
      case 'paste-menu': return BTS.menu(anchor, [{ label: 'Paste', icon: 'clipboard-paste', kbd: 'Ctrl+V', action: () => A.pasteCmd() }, { label: 'Paste values', icon: 'clipboard-type', action: () => A.clip ? A.pasteInternal('values') : A.pasteCmd('values') }, { label: 'Paste formulas', icon: 'square-function', action: () => A.pasteInternal('formulas') }, { label: 'Paste formatting', icon: 'paintbrush', action: () => A.pasteInternal('formats') }, { label: 'Transpose', icon: 'arrow-down-up', action: () => A.clip ? A.pasteInternal('transpose') : A.pasteCmd('transpose') }, { label: 'Keep column widths', icon: 'move-horizontal', action: () => { A.pasteInternal('all'); A.pasteInternal('colwidths'); } }, { sep: true }, { label: 'Paste special…', icon: 'clipboard-list', action: () => Dlg.pasteSpecial() }]);
      case 'format-painter': if (A.painter) A.cancelPainter(); else A.startPainter(ev && ev.detail > 1); return;
      /* font */
      case 'font': return App.applyFont(el.value);
      case 'font-pick': {
        const st = s.style(s.active.r, s.active.c); const th = STheme.cur();
        const f = await BTS.fontPicker($('#sel-font'), { current: st.font || SU.defFont, theme: [{ label: 'Headings', name: th.head }, { label: 'Body', name: th.body }, ...(th.num ? [{ label: 'Numbers', name: th.num }] : [])] });
        if (f) await App.applyFont(f); Grid.focus(); return;
      }
      case 'size': { const n = parseFloat(el.value); if (n > 0) A.setFontSize(n); Grid.focus(); return; }
      case 'font-grow': return A.growFont(1);
      case 'font-shrink': return A.growFont(-1);
      case 'bold': return A.toggleStyle('b');
      case 'italic': return A.toggleStyle('i');
      case 'underline': return A.toggleStyle('u');
      case 'strike': return A.toggleStyle('s');
      case 'borders': return bordersMenu(anchor);
      case 'border-outside': return A.applyBorders('outside');
      case 'border-none': return A.applyBorders('none');
      case 'fill-color-apply': return A.applyStyle({ fill: lastFill === 'none' ? undefined : lastFill }, 'Fill color');
      case 'fill-color': { const c = await BTS.colorPicker(anchor, { none: 'No fill', value: lastFill }); if (c === null) return; lastFill = c; $('#btn-fill .swatch').style.background = c === 'none' ? 'transparent' : c; return A.applyStyle({ fill: c === 'none' ? undefined : c }, 'Fill color'); }
      case 'font-color-apply': return A.applyStyle({ color: lastFont === 'auto' ? undefined : lastFont }, 'Font color');
      case 'font-color': { const c = await BTS.colorPicker(anchor, { auto: 'Automatic', value: lastFont }); if (c === null) return; lastFont = c; $('#btn-fcolor .swatch').style.background = c === 'auto' ? '#000' : c; return A.applyStyle({ color: c === 'auto' ? undefined : c }, 'Font color'); }
      /* alignment */
      case 'align-left': case 'align-center': case 'align-right': { const h = cmd.slice(6); const cur = s.style(s.active.r, s.active.c).h; return A.applyStyle({ h: cur === h ? undefined : h }, 'Align'); }
      case 'valign-top': return A.applyStyle({ v: 'top' }, 'Align');
      case 'valign-middle': return A.applyStyle({ v: 'middle' }, 'Align');
      case 'valign-bottom': return A.applyStyle({ v: undefined }, 'Align');
      case 'wrap': { A.toggleStyle('wrap'); return; }
      case 'indent-inc': return A.applyStyle((st) => ({ indent: Math.min(15, (st.indent || 0) + 1), h: st.h === 'center' ? 'left' : st.h }), 'Indent');
      case 'indent-dec': return A.applyStyle((st) => ({ indent: Math.max(0, (st.indent || 0) - 1) || undefined }), 'Indent');
      case 'merge-menu': return BTS.menu(anchor, [{ label: 'Merge & center', icon: 'table-cells-merge', action: () => A.merge('center') }, { label: 'Merge across', icon: 'rows-3', action: () => A.merge('across') }, { label: 'Merge cells', icon: 'square', action: () => A.merge('cells') }, { label: 'Unmerge cells', icon: 'table-cells-split', action: () => A.merge('unmerge') }]);
      case 'merge-center': return A.merge('center');
      /* number */
      case 'numfmt': { const id = el.value; if (id === 'more') { Dlg.formatCells(); return; } const f = SU.FORMATS.find(x => x.id === id); if (f) A.setNumberFormat(f.nf); Grid.focus(); return; }
      case 'currency-menu': return BTS.menu(anchor, [{ label: '₹ Indian rupee', action: () => A.setNumberFormat(SU.currencyFormat('₹')) }, { label: '$ US dollar', action: () => A.setNumberFormat(SU.currencyFormat('$')) }, { label: '€ Euro', action: () => A.setNumberFormat(SU.currencyFormat('€')) }, { label: '£ Pound', action: () => A.setNumberFormat(SU.currencyFormat('£')) }, { sep: true }, { label: '₹ Accounting', action: () => A.setNumberFormat(SU.accountingFormat('₹')) }, { label: '₹ Indian grouping (12,34,567)', action: () => A.setNumberFormat('"₹"#,##,##0.00') }, { label: 'More formats…', icon: 'settings-2', action: () => Dlg.formatCells() }]);
      case 'fmt-percent': return A.setNumberFormat('0%');
      case 'fmt-comma': return A.setNumberFormat(SU.accountingFormat('', 2));
      case 'dec-inc': return A.changeDecimals(1);
      case 'dec-dec': return A.changeDecimals(-1);
      case 'format-cells': return Dlg.formatCells();
      /* styles */
      case 'cf-menu': return cfMenu(anchor);
      case 'table-menu': return tableMenu(anchor);
      case 'format-table': return A.formatAsTable('Medium 2');
      case 'styles-menu': return stylesMenu(anchor);
      /* cells */
      case 'insert-menu': return BTS.menu(anchor, [{ label: 'Insert cells…', icon: 'square-plus', kbd: 'Ctrl++', action: () => Dlg.insertCells() }, { label: 'Insert sheet rows', icon: 'between-horizontal-start', action: () => A.insertRows() }, { label: 'Insert sheet columns', icon: 'between-vertical-start', action: () => A.insertCols() }, { label: 'Insert sheet', icon: 'file-plus', action: () => App.cmd('sheet-add') }]);
      case 'delete-menu': return BTS.menu(anchor, [{ label: 'Delete cells…', icon: 'square-minus', kbd: 'Ctrl+-', action: () => Dlg.deleteCells() }, { label: 'Delete sheet rows', icon: 'trash-2', action: () => A.deleteRows() }, { label: 'Delete sheet columns', icon: 'trash-2', action: () => A.deleteCols() }, { label: 'Delete sheet', icon: 'file-x', disabled: wb.sheets.length < 2, action: () => App.cmd('sheet-del') }]);
      case 'insert-cells-dialog': return Dlg.insertCells();
      case 'delete-cells-dialog': return Dlg.deleteCells();
      case 'format-menu': return BTS.menu(anchor, [{ head: 'Cell size' }, { label: 'Row height…', icon: 'move-vertical', action: () => Dlg.size('row') }, { label: 'AutoFit row height', icon: 'arrow-up-down', action: () => A.autofit('row') }, { label: 'Column width…', icon: 'move-horizontal', action: () => Dlg.size('col') }, { label: 'AutoFit column width', icon: 'arrow-left-right', action: () => A.autofit('col') }, { head: 'Visibility' }, { label: 'Hide rows', icon: 'eye-off', kbd: 'Ctrl+9', action: () => A.setHidden('r', true) }, { label: 'Hide columns', icon: 'eye-off', kbd: 'Ctrl+0', action: () => A.setHidden('c', true) }, { label: 'Unhide rows', icon: 'eye', action: () => A.setHidden('r', false) }, { label: 'Unhide columns', icon: 'eye', action: () => A.setHidden('c', false) }, { head: 'Organize sheets' }, { label: 'Rename sheet', icon: 'pencil', action: () => renameTab(wb.activeIndex) }, { label: 'Tab color…', icon: 'palette', action: async () => { const c = await BTS.colorPicker(anchor, { none: 'No color' }); if (c !== null) { wb.run('Tab color', () => wb.setProp(s, 'tabColor', c === 'none' || c === 'auto' ? null : c)); renderTabs(); } } }, { head: 'Protection' }, { label: s.protected ? 'Unprotect sheet' : 'Protect sheet', icon: 'lock', action: () => App.cmd('protect') }, { label: 'Lock / unlock cells', icon: 'lock-open', action: () => A.applyStyle((st) => ({ locked: st.locked === false ? undefined : false }), 'Lock cells') }, { sep: true }, { label: 'Format cells…', icon: 'paintbrush', kbd: 'Ctrl+1', action: () => Dlg.formatCells() }]);
      case 'hide-rows': return A.setHidden('r', true);
      case 'hide-cols': return A.setHidden('c', true);
      /* editing */
      case 'autosum': return A.autoSum('SUM');
      case 'autosum-menu': return BTS.menu(anchor, [['SUM', 'Sum'], ['AVERAGE', 'Average'], ['COUNT', 'Count numbers'], ['MAX', 'Max'], ['MIN', 'Min']].map(([f, l]) => ({ label: l, icon: 'sigma', action: () => A.autoSum(f) })).concat([{ sep: true }, { label: 'More functions…', icon: 'square-function', action: () => Dlg.insertFunction() }]));
      case 'fill-menu': return BTS.menu(anchor, [{ label: 'Down', icon: 'arrow-down', kbd: 'Ctrl+D', action: () => A.fillDir('down') }, { label: 'Right', icon: 'arrow-right', kbd: 'Ctrl+R', action: () => A.fillDir('right') }, { label: 'Up', icon: 'arrow-up', action: () => A.fillDir('up') }, { label: 'Left', icon: 'arrow-left', action: () => A.fillDir('left') }, { sep: true }, { label: 'Series…', icon: 'list-ordered', action: () => Dlg.series() }, { label: 'Flash Fill', icon: 'zap', kbd: 'Ctrl+E', action: () => A.flashFill() }]);
      case 'fill-down': return A.fillDir('down');
      case 'fill-right': return A.fillDir('right');
      case 'clear-menu': return BTS.menu(anchor, [{ label: 'Clear all', icon: 'eraser', action: () => A.clear('all') }, { label: 'Clear formats', icon: 'remove-formatting', action: () => A.clear('formats') }, { label: 'Clear contents', icon: 'delete', kbd: 'Del', action: () => A.clear('contents') }, { label: 'Clear notes', icon: 'message-square-x', action: () => A.clear('comments') }, { label: 'Clear hyperlinks', icon: 'unlink', action: () => A.clear('links') }]);
      case 'sort-menu': return BTS.menu(anchor, [{ label: 'Sort A to Z', icon: 'arrow-down-a-z', action: () => A.quickSort(false) }, { label: 'Sort Z to A', icon: 'arrow-down-z-a', action: () => A.quickSort(true) }, { label: 'Custom sort…', icon: 'arrow-up-down', action: () => Dlg.sort() }, { sep: true }, { label: 'Filter', icon: 'funnel', checked: !!s.filter, kbd: 'Ctrl+Shift+L', action: () => A.toggleFilter() }, { label: 'Clear filter', icon: 'funnel-x', disabled: !s.filter, action: () => A.clearFilters() }, { label: 'Reapply', icon: 'refresh-ccw', disabled: !s.filter, action: () => A.reapplyFilter() }]);
      case 'find-menu': return BTS.menu(anchor, [{ label: 'Find…', icon: 'search', kbd: 'Ctrl+F', action: () => Dlg.find(false) }, { label: 'Replace…', icon: 'replace', kbd: 'Ctrl+H', action: () => Dlg.find(true) }, { label: 'Go to…', icon: 'locate', kbd: 'Ctrl+G', action: () => Dlg.goto() }, { sep: true }, { label: 'Select formulas', icon: 'square-function', action: () => { s.sel = { r0: 0, c0: 0, ...s.usedRange() }; Grid.notify(); Dlg.goto; App.cmd('goto-formulas'); } }, { label: 'Select notes', icon: 'message-square', action: () => App.cmd('goto-notes') }]);
      case 'goto-formulas': case 'goto-notes': { const hits = []; const u = s.usedRange(); for (let r = 0; r <= u.r1; r++) for (let c = 0; c <= u.c1; c++) if (cmd === 'goto-formulas' ? s.isFormula(r, c) : s.comments.has(key(r, c))) hits.push({ r0: r, c0: c, r1: r, c1: c }); if (!hits.length) { BTS.toast('No cells were found.', { error: true }); return; } s.extraSel = hits.slice(0, -1); s.sel = hits[hits.length - 1]; s.active = { r: s.sel.r0, c: s.sel.c0 }; Grid.notify(); return; }
      case 'find': return Dlg.find(false);
      case 'replace': return Dlg.find(true);
      case 'goto': return Dlg.goto();
      /* insert */
      case 'insert-image': return Charts.insertImage();
      case 'chart-column-menu': return chartMenu(anchor, ['column', 'stackedColumn', 'stackedColumn100']);
      case 'chart-bar-menu': return chartMenu(anchor, ['bar', 'stackedBar', 'stackedBar100']);
      case 'chart-pie-menu': return chartMenu(anchor, ['pie', 'doughnut']);
      case 'chart-column': case 'chart-bar': case 'chart-line': case 'chart-pie': case 'chart-doughnut': case 'chart-area': case 'chart-scatter':
      case 'chart-bubble': case 'chart-combo': case 'chart-radar': case 'chart-waterfall': case 'chart-histogram':
      case 'chart-stackedColumn': case 'chart-stackedColumn100': case 'chart-stackedBar': case 'chart-stackedBar100':
        if (Charts.selected && Charts.find(Charts.selected)?.kind === 'chart' && BTS.activeTab === 'chart') return Charts.update({ type: cmd.slice(6) }); return Charts.create(cmd.slice(6));
      case 'spark-line': case 'spark-column': case 'spark-winloss': return Spark.dialog(cmd.slice(6));
      case 'pivot-insert': return Pivot.createDialog();
      case 'comment-new': return Threads.newComment();
      case 'comments-pane': return Threads.toggle();
      case 'comments-pane-close': return Threads.close();
      case 'thread-delete': return Threads.deleteAtActive();
      case 'thread-next': return Threads.go(1);
      case 'thread-prev': return Threads.go(-1);
      /* themes */
      case 'theme-menu': return themeMenu(anchor);
      case 'theme-colors': return themeMenu(anchor, 'colors');
      case 'theme-fonts': return themeMenu(anchor, 'fonts');
      case 'theme-reset': if (wb.theme === 'bts' && !wb.themeOver) return; wb.run('Reset theme', () => { wb.setProp(wb, 'themeOver', null); wb.setProp(wb, 'theme', 'bts'); A.restyleTagged(); Pivot.restyleAll(); }); Grid.render(); Charts.renderSheet(); App.refreshStatus(); return;
      /* data tools */
      case 'consolidate': return DT.consolidateDialog();
      case 'circle-invalid': return DT.circleInvalid();
      case 'clear-circles': return DT.clearCircles();
      case 'goal-seek': return WhatIf.goalSeek();
      case 'scenarios': return WhatIf.scenarioManager();
      case 'group': return DT.group();
      case 'ungroup': return DT.ungroup();
      case 'clear-outline': return DT.clearOutline();
      case 'show-detail': return DT.detail(true);
      case 'hide-detail': return DT.detail(false);
      /* protection */
      case 'protect-workbook': return Prot.toggleWorkbook();
      case 'allow-ranges': return Prot.allowRangesDialog();
      /* pivot */
      case 'pv-refresh': { const pv = Pivot.atActive() || (document.getElementById('pv-pane') && Pivot.find(document.getElementById('pv-pane').dataset.id)?.def); if (!pv) { BTS.toast('Select a cell in a PivotTable first.', { error: true }); return; } return Pivot.refresh(Pivot.find(pv.id).s, pv); }
      case 'pivot-refresh-all': return Pivot.refreshAll();
      case 'pivot-source': return Pivot.changeSource();
      case 'pivot-options': return Pivot.options();
      case 'pivot-delete': { const pv = Pivot.atActive(); if (pv) Pivot.remove(pv); return; }
      case 'pivot-select': { const pv = Pivot.atActive(); if (pv) Grid.selectRange(pv.out); return; }
      case 'pivot-chart': return Pivot.createChart('column');
      case 'pivot-fields': { const p = document.getElementById('pv-pane'); Pivot.paneClosed = !!p && !p.classList.contains('hidden'); Pivot.refreshPane(); App.refreshContext(); return; }
      case 'pv-pane-close': Pivot.paneClosed = true; Pivot.refreshPane(); App.refreshContext(); return;
      case 'pivot-subtotals': { const pv = Pivot.atActive(); if (!pv) return; return BTS.menu(anchor, [{ label: 'Do not show subtotals', checked: pv.subtotals === false, action: () => Pivot.update(s, pv.id, { subtotals: false }, 'Subtotals') }, { label: 'Show all subtotals at bottom of group', checked: pv.subtotals !== false, action: () => Pivot.update(s, pv.id, { subtotals: true }, 'Subtotals') }]); }
      case 'pivot-grand': { const pv = Pivot.atActive(); if (!pv) return; const gr = pv.grandRows !== false, gc = pv.grandCols !== false; return BTS.menu(anchor, [{ label: 'Off for rows and columns', checked: !gr && !gc, action: () => Pivot.update(s, pv.id, { grandRows: false, grandCols: false }, 'Grand totals') }, { label: 'On for rows and columns', checked: gr && gc, action: () => Pivot.update(s, pv.id, { grandRows: true, grandCols: true }, 'Grand totals') }, { label: 'On for rows only', checked: gr && !gc, action: () => Pivot.update(s, pv.id, { grandRows: true, grandCols: false }, 'Grand totals') }, { label: 'On for columns only', checked: !gr && gc, action: () => Pivot.update(s, pv.id, { grandRows: false, grandCols: true }, 'Grand totals') }]); }
      case 'pivot-styles': return Pivot.styleGallery(anchor);
      /* table design */
      case 'table-name': { const t = s.tableAt(s.active.r, s.active.c); if (t && el.value.trim() !== t.name) { if (!A.renameTable(t, el.value)) el.value = t.name; } return; }
      case 'table-resize': { const t = s.tableAt(s.active.r, s.active.c); if (t) A.resizeTable(t); return; }
      case 'table-convert': { const t = s.tableAt(s.active.r, s.active.c); if (t) A.convertTableToRange(t); return; }
      case 'table-pivot': { const t = s.tableAt(s.active.r, s.active.c); return Pivot.createDialog(t); }
      case 'tbl-header': case 'tbl-total': case 'tbl-banded': case 'tbl-first': case 'tbl-last': case 'tbl-bandcols': case 'tbl-filter': {
        const t = s.tableAt(s.active.r, s.active.c); if (!t) return; const prop = { 'tbl-header': 'header', 'tbl-total': 'total', 'tbl-banded': 'banded', 'tbl-first': 'firstCol', 'tbl-last': 'lastCol', 'tbl-bandcols': 'bandedCols', 'tbl-filter': 'filterBtn' }[cmd];
        if (prop === 'header' && !el.checked) { BTS.toast('Tables in BTS Sheets always keep their header row.', { error: true }); el.checked = true; return; }
        return A.setTableProps(t, { [prop]: el.checked }, 'Table style options');
      }
      /* sparkline tab */
      case 'spark-type-line': case 'spark-type-column': case 'spark-type-winloss': return Spark.update({ type: cmd.slice(11) }, 'Sparkline type');
      case 'spk-high': case 'spk-low': case 'spk-neg': case 'spk-first': case 'spk-last': case 'spk-markers': return Spark.update({ show: { [{ 'spk-high': 'high', 'spk-low': 'low', 'spk-neg': 'neg', 'spk-first': 'first', 'spk-last': 'last', 'spk-markers': 'markers' }[cmd]]: el.checked } }, 'Sparkline points');
      case 'spark-color': { const c = await BTS.colorPicker(anchor, {}); if (c && c[0] === '#') Spark.update({ color: c }, 'Sparkline color'); return; }
      case 'spark-marker-color': { const c = await BTS.colorPicker(anchor, {}); if (c && c[0] === '#') Spark.update({ markerColor: c, highColor: c, lowColor: c }, 'Marker color'); return; }
      case 'spark-clear': return Spark.clear(false);
      case 'spark-clear-group': return Spark.clear(true);
      case 'spark-edit': return Spark.editData();
      /* chart extras */
      case 'chart-styles': { const f = Charts.find(Charts.selected); if (!f) return; return BTS.menu(anchor, Charts.STYLES.map(([n, l]) => ({ label: l, checked: (f.obj.style || 1) === n, action: () => Charts.update({ style: n }, 'Chart style') }))); }
      case 'chart-axis-titles': return Charts.update({ axisTitles: el.checked }, 'Axis titles');
      case 'chart-axis-edit': { const f = Charts.find(Charts.selected); if (!f) return; const d = Charts.config(f.obj).options.scales || {}; const res = await BTS.dialog({ title: 'Axis titles', width: 380, body: `<div class="field"><label>Horizontal axis title</label><input id="ax-x" value="${SU.esc((d.x && d.x.title && d.x.title.text) || '')}"></div><div class="field"><label>Vertical axis title</label><input id="ax-y" value="${SU.esc((d.y && d.y.title && d.y.title.text) || '')}"></div>`, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ x: r.querySelector('#ax-x').value, y: r.querySelector('#ax-y').value }) }] }); if (res) { const horiz = /^(bar|stackedBar|stackedBar100)$/.test(f.obj.type); Charts.update(horiz ? { xTitle: res.x, yTitle: res.y, axisTitles: true } : { xTitle: res.x, yTitle: res.y, axisTitles: true }, 'Axis titles'); } return; }
      case 'hyperlink': return Dlg.hyperlink();
      case 'comment': return Dlg.note();
      case 'insert-function': return Dlg.insertFunction();
      case 'symbol': return Dlg.symbols();
      /* page layout */
      case 'margins-menu': return BTS.menu(anchor, [['normal', 'Normal', 'Top 0.75" · Left 0.7"'], ['narrow', 'Narrow', 'Top 0.75" · Left 0.25"'], ['wide', 'Wide', 'Top 1" · Left 1"']].map(([k, l, d]) => ({ label: l, checked: s.page.margins === k, html: `<span style="min-width:70px">${l}</span><span style="color:var(--ink-3);font-size:11.5px">${d}</span>`, action: () => setPage({ margins: k }) })));
      case 'orientation-menu': return BTS.menu(anchor, [{ label: 'Portrait', icon: 'rectangle-vertical', checked: s.page.orientation === 'portrait', action: () => setPage({ orientation: 'portrait' }) }, { label: 'Landscape', icon: 'rectangle-horizontal', checked: s.page.orientation === 'landscape', action: () => setPage({ orientation: 'landscape' }) }]);
      case 'size-menu': return BTS.menu(anchor, [['A4', '21 × 29.7 cm'], ['Letter', '8.5 × 11 in'], ['Legal', '8.5 × 14 in'], ['A3', '29.7 × 42 cm'], ['A5', '14.8 × 21 cm']].map(([k, d]) => ({ label: k, checked: s.page.paper === k, html: `<span style="min-width:60px">${k}</span><span style="color:var(--ink-3);font-size:11.5px">${d}</span>`, action: () => setPage({ paper: k }) })));
      case 'printarea-menu': return BTS.menu(anchor, [{ label: 'Set print area', icon: 'scan', action: () => setPage({ printArea: A.eff(s.sel) }) }, { label: 'Clear print area', icon: 'x', disabled: !s.page.printArea, action: () => setPage({ printArea: null }) }, ...(s.page.printArea ? [{ head: 'Current: ' + SU.rangeStr(s.page.printArea) }] : [])]);
      case 'fit-width': return setPage({ fit: el.checked ? 'width' : 'none' });
      case 'fit-page': return setPage({ fit: el.checked ? 'page' : 'none' });
      case 'print-gridlines': return setPage({ gridlines: el.checked });
      case 'print-headings': return setPage({ headings: el.checked });
      case 'print-preview': return BTS.openBackstage('print');
      /* formulas */
      case 'fn-cat': return Dlg.fnMenu(anchor, el.dataset.cat);
      case 'name-manager': return Dlg.nameManager();
      case 'define-name': return Dlg.defineName();
      case 'use-name': { const names = wb.listNames(); return BTS.menu(anchor, names.length ? names.map(n => ({ label: n.name, icon: 'tag', action: () => { if (Ed.isEditing()) Ed.insertAtCaret(n.name); else Ed.start('=' + n.name, 'edit'); } })) : [{ head: 'No names defined' }]); }
      case 'names-from-sel': return Dlg.namesFromSelection();
      case 'show-formulas': Grid.showFormulas = !Grid.showFormulas; Grid.render(); App.refreshRibbon(); App.refreshStatus(); return;
      case 'recalc': wb.hf.rebuildAndRecalculate(); wb.emit('recalc'); BTS.toast('Recalculated'); return;
      case 'trace-precedents': case 'trace-dependents': {
        const a = s.active; let list = [];
        try { list = cmd === 'trace-precedents' ? wb.hf.getCellPrecedents(s.a(a.r, a.c)) : wb.hf.getCellDependents(s.a(a.r, a.c)); } catch { }
        const hl = list.filter(x => x.sheet === s.id || (x.start && x.start.sheet === s.id)).map(x => x.start ? { range: { r0: x.start.row, c0: x.start.col, r1: x.end.row, c1: x.end.col }, color: cmd === 'trace-precedents' ? '#2f6fde' : '#d6336c' } : { range: { r0: x.row, c0: x.col, r1: x.row, c1: x.col }, color: cmd === 'trace-precedents' ? '#2f6fde' : '#d6336c' });
        if (!hl.length) { BTS.toast(cmd === 'trace-precedents' ? 'This cell doesn’t refer to other cells.' : 'No formulas refer to this cell.', { error: true }); return; }
        Grid.refHL = hl; Grid.render(); return;
      }
      case 'trace-clear': Grid.refHL = []; Grid.render(); return;
      /* data */
      case 'import-text': return Dlg.importText();
      case 'sort-asc': return A.quickSort(false);
      case 'sort-desc': return A.quickSort(true);
      case 'sort-custom': return Dlg.sort();
      case 'filter': A.toggleFilter(); App.refreshRibbon(); return;
      case 'filter-clear': return A.clearFilters();
      case 'filter-reapply': return A.reapplyFilter();
      case 'text-to-columns': return Dlg.textToColumns();
      case 'flash-fill': return A.flashFill();
      case 'remove-duplicates': return Dlg.removeDuplicates();
      case 'validation': return Dlg.validation();
      case 'group-rows': return DT.group('r');
      case 'ungroup-rows': return DT.ungroup('r');
      case 'subtotal': { const g = A.dataRange(); const r = g.r1 + 1; const cells = [{ r, c: g.c0, v: 'Total' }]; for (let c = g.c0 + 1; c <= g.c1; c++) { let num = false; for (let rr = g.r0; rr <= g.r1; rr++) if (typeof s.value(rr, c) === 'number') { num = true; break; } if (num) cells.push({ r, c, v: `=SUBTOTAL(9,${SU.rangeStr({ r0: g.r0 + (A.hasHeader(s, g) ? 1 : 0), c0: c, r1: g.r1, c1: c })})` }); } wb.run('Totals row', () => { wb.setContents(s, cells); cells.forEach(x => wb.setStyle(s, x.r, x.c, SU.mergeStyle(s.style(x.r, x.c), { b: true, bT: { style: 'thin', color: '#000000' } }))); }); return; }
      /* review */
      case 'comment-delete': { const k = key(s.active.r, s.active.c); if (!s.comments.has(k)) return; wb.run('Delete note', () => { const m = new Map(s.comments); m.delete(k); wb.setProp(s, 'comments', m); }); hideNote(); return; }
      case 'comment-next': case 'comment-prev': { const keys = [...s.comments.keys()].map(SU.unkey).sort((a, b) => a.r - b.r || a.c - b.c); if (!keys.length) { BTS.toast('This sheet has no notes.', { error: true }); return; } const a = s.active; let i = cmd === 'comment-next' ? keys.findIndex(k => k.r > a.r || (k.r === a.r && k.c > a.c)) : keys.map((k, i) => [k, i]).filter(([k]) => k.r < a.r || (k.r === a.r && k.c < a.c)).map(x => x[1]).pop(); if (i === undefined || i < 0) i = cmd === 'comment-next' ? 0 : keys.length - 1; Grid.select(keys[i].r, keys[i].c); showNote(keys[i].r, keys[i].c, s.comments.get(key(keys[i].r, keys[i].c)), 'Note'); return; }
      case 'comments-show': App.notesShown = !App.notesShown; persistNotes(); App.refreshRibbon(); return;
      case 'protect': return Prot.toggleSheet();
      case 'stats': return Dlg.stats();
      /* view */
      case 'view-gridlines': s.showGrid = el ? el.checked : !s.showGrid; Grid.render(); App.refreshRibbon(); BTS.markDirty(); return;
      case 'view-headings': s.showHeadings = el ? el.checked : !s.showHeadings; Grid.updateSpacer(); Grid.renderNow(); Ed.reposition(); return;
      case 'view-fbar': document.body.classList.toggle('no-fbar', !(el ? el.checked : document.body.classList.contains('no-fbar'))); return;
      case 'zoom-dialog': return Dlg.zoom();
      case 'zoom-sel': { const g = A.eff(s.sel); const w = (s.colLeft(g.c1 + 1) - s.colLeft(g.c0)), h = (s.rowTop(g.r1 + 1) - s.rowTop(g.r0)); const z = Math.floor(Math.min((Grid.scroller.clientWidth - 60) / w, (Grid.scroller.clientHeight - 40) / h) * 100); BTS.setZoom(SU.clamp(z, 25, 400)); Grid.scroller.scrollLeft = s.colLeft(g.c0) * Grid.z; Grid.scroller.scrollTop = s.rowTop(g.r0) * Grid.z; return; }
      case 'new-window': return BTS.newWindow();
      case 'freeze-menu': return BTS.menu(anchor, [{ label: s.freeze.r || s.freeze.c ? 'Unfreeze panes' : 'Freeze panes (at selection)', icon: 'snowflake', action: () => A.freeze(s.freeze.r || s.freeze.c ? 'none' : 'cell') }, { label: 'Freeze top row', icon: 'panel-top', action: () => A.freeze('row') }, { label: 'Freeze first column', icon: 'panel-left', action: () => A.freeze('col') }]);
      /* sheets */
      case 'sheet-add': { const ns = wb.run('Insert sheet', () => wb.addSheet(null, wb.activeIndex + 1)); renderTabs(); App.switchSheet(wb.sheets.indexOf(ns)); return; }
      case 'sheet-dup': { const ns = wb.run('Duplicate sheet', () => wb.duplicateSheet(s)); renderTabs(); App.switchSheet(wb.sheets.indexOf(ns)); return; }
      case 'sheet-del': { if (wb.sheets.length < 2) return; const has = s.usedRange().r1 >= 0 || s.charts.length; if (has && !(await BTS.confirm('Delete sheet', `“${s.name}” will be deleted. You can undo this with Ctrl+Z.`, 'Delete'))) return; wb.run('Delete sheet', () => wb.deleteSheet(s)); Charts.select(null); Grid.restoreScroll(); renderTabs(); Charts.renderSheet(); App.refreshFormulaBar(); return; }
      /* chart contextual */
      case 'chart-type': return chartTypeMenu(anchor);
      case 'chart-title': return Charts.editTitle();
      case 'chart-data': { const f = Charts.find(Charts.selected); if (!f) return; const v = await BTS.prompt('Select data', 'Chart data range (e.g. A1:C8 or Sheet2!A1:B5)', (f.obj.sheet ? SU.quoteSheet(f.obj.sheet) + '!' : '') + SU.rangeStr(f.obj.range)); if (!v) return; const r = SU.parseRef(v); if (!r) { BTS.toast('That reference isn’t valid.', { error: true }); return; } Charts.update({ range: r.range, sheet: r.sheet && r.sheet !== f.s.name ? (wb.sheetByName(r.sheet) || {}).name || null : null, pick: undefined }, 'Chart data'); return; }
      case 'chart-switch': { const f = Charts.find(Charts.selected); if (!f) return; Charts.update({ switchRC: !f.obj.switchRC }); return; }
      case 'chart-legend': return Charts.update({ legend: el.checked });
      case 'chart-labels': return Charts.update({ labels: el.checked });
      case 'chart-grid': return Charts.update({ grid: el.checked });
      case 'chart-colors': return BTS.menu(anchor, Object.keys(Charts.PALETTES).map((n) => { const cols = Charts.PALETTES[n]; return { label: n, html: `<span style="display:flex;gap:1px">${cols.slice(0, 6).map(c => `<i style="width:10px;height:12px;background:${c}"></i>`).join('')}</span><span>${n === 'Theme' ? 'Theme colours (' + SU.esc(STheme.cur().name) + ')' : n}</span>`, action: () => Charts.update({ palette: n }, 'Chart colors') }; }));
      case 'chart-delete': return Charts.deleteSelected();
    }
  };
  /* font: load the face before the first paint, then apply */
  App.applyFont = async (name) => {
    const st = wb.active.style(wb.active.active.r, wb.active.active.c);
    await Promise.race([Promise.all([STheme.loadFont(name, false, false), STheme.loadFont(name, !!st.b, !!st.i)]), new Promise(r => setTimeout(r, 2500))]);
    A.applyStyle({ font: name === SU.defFont ? undefined : name }, 'Font');
    Grid.render();
  };
  /* contextual tabs: Table Design, PivotTable Analyze, Sparkline (Chart is handled by Charts.select) */
  App.refreshContext = () => {
    const s = wb.active; if (!s) return; const a = s.active; const chartSel = !!Charts.selected;
    const t = !chartSel && s.tableAt(a.r, a.c), pv = !chartSel && Pivot.at(s, a.r, a.c), sp = !chartSel && Grid.sparkAt(s, a.r, a.c);
    const show = (id, on, tab) => { const b = document.getElementById(id); const was = !b.classList.contains('hidden'); b.classList.toggle('hidden', !on); if (!on && was && BTS.activeTab === tab) BTS.showTab('home'); };
    show('tab-table', !!t, 'table'); show('tab-pivot', !!pv, 'pivot'); show('tab-spark', !!sp, 'spark');
    if (t) {
      const ni = $('#tbl-name'); if (document.activeElement !== ni) ni.value = t.name;
      const chk = (id, v) => { $(id).checked = !!v; };
      chk('#tbl-header', t.header !== false); chk('#tbl-total', t.total); chk('#tbl-banded', t.banded !== false); chk('#tbl-first', t.firstCol); chk('#tbl-last', t.lastCol); chk('#tbl-bandcols', t.bandedCols); chk('#tbl-filter', t.filterBtn !== false);
      const gal = $('#tbl-gallery'); const th = STheme.cur(); const ids = ['Light 1', 'Light 9', 'Medium 2', 'Medium 3', 'Medium 4', 'Medium 7', 'Dark 2'];
      const key = ids.join() + th.accents.join() + t.styleId; if (gal._k !== key) { gal._k = key; gal.innerHTML = ids.map(id => `<button class="gi tsw ${t.styleId === id ? 'on' : ''}" data-id="${id}" title="Table style ${id}">${STheme.swatchHTML(id, th)}</button>`).join(''); }
    }
    if (pv) { $('#pv-name-label').textContent = pv.name; const p = document.getElementById('pv-pane'); $('#btn-pvfields').classList.toggle('on', !!p && !p.classList.contains('hidden')); }
    if (sp) { const g = sp.g; ['line', 'column', 'winloss'].forEach(k => $('#spk-' + k).classList.toggle('on', g.type === k)); const sh = g.show || {}; ['high', 'low', 'neg', 'first', 'last', 'markers'].forEach(k => { $('#spk-' + k).checked = !!sh[k]; }); }
  };
  /** side panes (PivotTable Fields, Comments) take room from the grid */
  App.layoutPanes = () => {
    const open = [...document.querySelectorAll('#workspace > .side-pane')].some(p => !p.classList.contains('hidden'));
    document.getElementById('workspace').classList.toggle('with-pane', open);
    Grid.resize(); Ed.reposition(); App.refreshValidationUI();
  };
  /* status bar: right-click to choose statistics */
  const SB_KEYS = [['average', 'Average'], ['count', 'Count'], ['ncount', 'Numerical Count'], ['min', 'Minimum'], ['max', 'Maximum'], ['sum', 'Sum']];
  App.statusOpts = () => { let o = null; try { o = JSON.parse(localStorage.getItem('bts-sheets-status') || 'null'); } catch { } return o || { average: true, count: true, sum: true, ncount: false, min: false, max: false }; };
  function statusMenu(e) {
    const o = App.statusOpts();
    BTS.menu({ x: e.clientX, y: e.clientY - 230 }, [{ head: 'Show on the status bar' }, ...SB_KEYS.map(([k, l]) => ({ label: l, checked: !!o[k], action: () => { o[k] = !o[k]; try { localStorage.setItem('bts-sheets-status', JSON.stringify(o)); } catch { } App.refreshStatus(); } }))]);
  }
  App.statusMenu = statusMenu;
  function setPage(p) { const s = wb.active; wb.run('Page setup', () => wb.setProp(s, 'page', { ...s.page, ...p })); App.refreshRibbon(); BTS.toast('Page setup updated'); }

  document.addEventListener('DOMContentLoaded', init);
})();
