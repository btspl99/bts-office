/* BTS Slides — application bootstrap: shell configuration, command routing, keyboard, files, PDF. */
(function () {
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const P = () => SL.pres, S = () => SL.slide();
  const f2 = (n) => Math.round(n * 100) / 100;
  const C = window.CMD;

  /* ---------- load a presentation into the editor ---------- */
  function setPresentation(p, opts = {}) {
    if (E.editing) { E.editing = null; }
    FX.show.on && FX.end();
    if (E.view === 'master' && SLX.masterAbort) SLX.masterAbort();
    SL.pres = p; SL.cur = 0; E.sel = []; E.railSel = new Set([0]); E.fit = true;
    if (!opts.keepAssets) { /* assets are additive; unused ones are dropped at save time */ }
    SL.resetHistory(); C.buildThemeGallery(); E.renderAll(); requestAnimationFrame(E.layout);
    // thumbnails and the canvas are redrawn once the bundled fonts this presentation uses are loaded
    SL.ensureFonts(p).then(() => { if (SL.pres === p) { E.renderAll(); C.buildThemeGallery(); } });
  }
  window.SLApp = { setPresentation };

  /* ---------- PDF export: render every visible slide as a page ---------- */
  async function exportPDF() {
    if (E.editing) E.endEdit();
    if (E.view === 'master') SLX.masterClose();
    await SL.ensureFonts(P());
    const root = $('#print-root'); root.innerHTML = '';
    const W = SL.W(P()), H = SL.H(P()); const win = W / 96, hin = H / 96;
    root.style.width = win + 'in';
    const st = document.createElement('style'); st.id = 'print-page-style';
    st.textContent = `@page { size: ${win}in ${hin}in; margin: 0; }`;
    document.head.appendChild(st);
    P().slides.forEach((s, i) => {
      if (s.hidden) return;
      const page = document.createElement('div'); page.className = 'print-page'; page.style.width = win + 'in'; page.style.height = hin + 'in';
      const node = SL.renderSlide(s, P(), { index: i }); E.fitAllText(node); node.style.transform = `scale(${(win * 96) / W})`; node.style.transformOrigin = '0 0';
      page.appendChild(node); root.appendChild(page);
    });
    document.body.classList.add('printing');
    await document.fonts.ready; await new Promise(r => setTimeout(r, 120));
    return { pageSize: { width: win, height: hin }, landscape: false, preferCSSPageSize: true };
  }
  function afterExportPDF() { document.body.classList.remove('printing'); $('#print-root').innerHTML = ''; $('#print-root').style.width = ''; $('#print-page-style')?.remove(); }

  /* ---------- save / load ---------- */
  async function save(ext) {
    if (E.editing) E.endEdit();
    if (E.view === 'master') SLX.masterClose();
    if (ext === 'pptx') return await PPTXExport.build(P());
    if (ext === 'json') { const used = {}; JSON.stringify(P()).replace(/"(a[\w]+)"/g, (m, id) => { if (SL.assets[id]) used[id] = SL.assets[id]; }); return JSON.stringify({ bts: 'slides', version: 1, pres: P(), assets: used }); }
    throw new Error('Unsupported format: ' + ext);
  }
  async function load(f) {
    const ext = BTS.extOf(f.name);
    if (ext === 'json') { const o = JSON.parse(new TextDecoder().decode(f.data)); Object.assign(SL.assets, o.assets || {}); setPresentation(o.pres); return; }
    if (ext !== 'pptx') throw new Error('BTS Slides opens .pptx presentations');
    const p = await PPTXImport.parse(f.data);
    if (p.props) { BTS.setProps(p.props); delete p.props; }
    SL._lastRestored = p._restored || 0; delete p._restored;
    if (SLX.normalizeComments) SLX.normalizeComments(p);
    setPresentation(p);
  }
  function newDocument() { setPresentation(SL.newPresentation(P().theme.id === 'custom' ? 'clean' : 'clean')); BTS.title = BTS.cfg.defaultTitle; BTS.filePath = null; BTS.markClean(); }

  /* ---------- search (title bar) ---------- */
  function search(q) { if (!q) return; C._case = false; if (!C.findNext(q)) BTS.toast('No matches for “' + q + '”'); }

  function infoHTML() {
    let words = 0, chars = 0; P().slides.forEach(s => s.elements.forEach(e => { const t = e.type === 'table' ? e.rows.flat().map(c => SL.plainText(c.html)).join(' ') : SL.plainText(e.html || ''); words += (t.match(/\S+/g) || []).length; chars += t.length; }));
    const media = P().slides.reduce((a, s) => a + s.elements.filter(e => e.type === 'image').length, 0);
    const th = SL.themeOf(P());
    return `<div class="info-grid"><div class="info-thumb">${SL.thumbHTML(P().slides[0], P(), 320, 0)}</div><dl>
      <dt>Slides</dt><dd>${P().slides.length}${P().slides.some(s => s.hidden) ? ` (${P().slides.filter(s => s.hidden).length} hidden)` : ''}</dd>
      <dt>Words</dt><dd>${words.toLocaleString()}</dd><dt>Characters</dt><dd>${chars.toLocaleString()}</dd><dt>Pictures</dt><dd>${media}</dd>
      <dt>Slide size</dt><dd>${f2(SL.W(P()) / 96)} × ${f2(SL.H(P()) / 96)} in (${Math.abs(SL.W(P()) / SL.H(P()) - 16 / 9) < .02 ? '16:9' : Math.abs(SL.W(P()) / SL.H(P()) - 4 / 3) < .02 ? '4:3' : 'custom'})</dd>
      <dt>Theme</dt><dd>${BTS.esc(th.name)}</dd><dt>Location</dt><dd>${BTS.esc(BTS.filePath || 'Not saved yet')}</dd></dl></div>`;
  }

  /* ---------- zoom ---------- */
  function onZoom(pct) { if (E._fitting || !E._booted) return; if (E.view === 'sorter') { E.sorterZoom = pct / 100; E.renderSorter(); return; } E.fit = false; E.zoom = pct / 100; E.layout(); }
  function zoomFit() { E.fit = true; E.layout(); }

  /* ====================================================================== */
  /*  Command router                                                          */
  /* ====================================================================== */
  const needSel = (msg = 'Select an object on the slide first') => { if (!E.selected().length) { BTS.toast(msg); return false; } return true; };
  async function onCommand(cmd, el, ev) {
    if (window.SLX && SLX.cmd[cmd]) return SLX.cmd[cmd](el, ev);
    if (E.view === 'master' && SLX.masterBlocks && SLX.masterBlocks(cmd)) return;
    const sel = E.selected(); const one = sel[0];
    switch (cmd) {
      /* ----- edit ----- */
      case 'undo': if (E.editing) { E.endEdit(); } if (SL.undo()) { E.sel = E.sel.filter(id => SL.findEl(id)); E.renderAll(); C.buildThemeGallery(); } return;
      case 'redo': if (E.editing) E.endEdit(); if (SL.redo()) { E.renderAll(); C.buildThemeGallery(); } return;
      case 'paste': if (E.editing) { try { const t = await navigator.clipboard.readText(); document.execCommand('insertText', false, t); } catch { document.execCommand('paste'); } return; } return E.pasteFromSystem(null);
      case 'cut': if (E.editing) return document.execCommand('cut'); if (railFocused()) { C.copySlides(); return E.deleteSlides(); } return needSel() && E.copy(true);
      case 'copy': if (E.editing) return document.execCommand('copy'); if (railFocused()) return C.copySlides(); return needSel() && E.copy();
      case 'format-painter': return E.painter ? (E.painter = null, $('#stage').classList.remove('painting'), el.classList.remove('on')) : C.startPainter();
      case 'find': return E.openPane('find');
      case 'replace': return E.openPane('find');
      case 'select-menu': return BTS.menu(el, [{ label: 'Select all', icon: 'box-select', kbd: 'Ctrl+A', action: () => E.selectAll() }, { label: 'Selection pane…', icon: 'layers-2', action: () => E.openPane('selection') }]);
      /* ----- slides ----- */
      case 'new-slide': return E.newSlide();
      case 'new-slide-menu': return C.layoutPicker(el, (l) => E.newSlide(l));
      case 'layout-menu': return C.layoutPicker(el, (l) => { [...E.railSel].forEach(i => E.applyLayout(l, P().slides[i])); E.sel = []; SL.commit(); E.renderAll(); }, S().layout);
      case 'reset-slide': [...E.railSel].forEach(i => E.resetSlide(P().slides[i])); SL.commit(); E.renderAll(); return;
      case 'duplicate-slide': return E.duplicateSlides();
      case 'hide-slide': { const on = !S().hidden; [...E.railSel].forEach(i => P().slides[i].hidden = on); SL.commit(); E.renderAll(); return; }
      case 'section-menu': return BTS.menu(el, [
        { label: 'Add section', icon: 'plus', action: () => C.addSection() },
        { label: 'Rename section', icon: 'pencil', disabled: !E.sectionOf(SL.cur), action: async () => { const s = E.sectionOf(SL.cur); const n = await BTS.prompt('Rename section', 'Section name', s.name); if (n) { s.name = n; SL.commit(); E.renderAll(); } } },
        { label: 'Remove section', icon: 'trash-2', disabled: !E.sectionOf(SL.cur), action: () => { const s = E.sectionOf(SL.cur); P().sections = P().sections.filter(x => x !== s); SL.commit(); E.renderAll(); } },
        { label: 'Remove all sections', icon: 'x', disabled: !(P().sections || []).length, action: () => { P().sections = []; SL.commit(); E.renderAll(); } },
      ]);
      /* ----- font ----- */
      case 'font-size': C.setFontSize(parseFloat(el.value)); return refocus();
      case 'grow-font': return C.stepFont(1);
      case 'shrink-font': return C.stepFont(-1);
      case 'clear-format': return C.clearFormat();
      case 'bold': case 'italic': case 'underline': case 'strike': return C.toggleStyle(cmd);
      case 'shadow-text': if (!C.textTargets().length) return BTS.toast('Select a text box first'); return C.setElProp('effect', E.primary().effect === 'shadow' ? undefined : 'shadow');
      case 'char-spacing': return BTS.menu(el, [['Very tight', -1.5], ['Tight', -.75], ['Normal', 0], ['Loose', 1.5], ['Very loose', 3]].map(([n, v]) => ({ label: n, checked: (E.primary()?.spacing || 0) === v, action: () => C.setElProp('spacing', v || undefined) })));
      case 'change-case': return BTS.menu(el, [['Sentence case.', 'sentence'], ['lowercase', 'lower'], ['UPPERCASE', 'upper'], ['Capitalize Each Word', 'title'], ['tOGGLE cASE', 'toggle']].map(([n, v]) => ({ label: n, action: () => C.changeCase(v) })));
      case 'font-color': { const c = await BTS.colorPicker(el, { auto: 'Automatic', value: '#d0384e' }); if (!c) return; if (c !== 'auto') { const sw = $('#sw-font'); if (sw) sw.style.background = c; } return C.setColor(c); }
      case 'highlight': { const c = await BTS.colorPicker(el, { none: 'No color' }); if (!c) return; if (c !== 'none') $('#sw-hl').style.background = c; return C.setHighlight(c); }
      /* ----- paragraph ----- */
      case 'bullets': return C.setList('ul');
      case 'numbering': return C.setList('ol');
      case 'indent': return C.indent(1);
      case 'outdent': return C.indent(-1);
      case 'align-left': return C.setAlign('left');
      case 'align-center': return C.setAlign('center');
      case 'align-right': return C.setAlign('right');
      case 'align-justify': return C.setAlign('justify');
      case 'line-spacing': if (!C.textTargets().length) return BTS.toast('Select a text box first'); return BTS.menu(el, [...[1, 1.15, 1.5, 2, 2.5, 3].map(v => ({ label: String(v), checked: (E.primary().lh || 1) === v, action: () => C.setElProp('lh', v) })), { sep: 1 }, { label: 'Line spacing options…', action: () => E.openPane('shape') }]);
      case 'valign': if (!C.textTargets().length) return BTS.toast('Select a text box first'); return BTS.menu(el, [['Top', 'top', 'align-vertical-justify-start'], ['Middle', 'middle', 'align-vertical-justify-center'], ['Bottom', 'bottom', 'align-vertical-justify-end']].map(([n, v, ic]) => ({ label: n, icon: ic, action: () => C.setElProp('valign', v) })));
      case 'text-dir': if (!C.textTargets().length) return BTS.toast('Select a text box first'); return BTS.menu(el, [['Horizontal', 'horz'], ['Rotate all text 90°', 'vert'], ['Rotate all text 270°', 'vert270']].map(([n, v]) => ({ label: n, checked: (E.primary().dir || 'horz') === v, action: () => C.setElProp('dir', v === 'horz' ? undefined : v) })));
      /* ----- drawing ----- */
      case 'draw-shape': return E.setDrawMode({ type: 'shape', kind: el.dataset.kind });
      case 'shapes-menu': return C.shapePicker(el, (k) => E.setDrawMode({ type: 'shape', kind: k }));
      case 'change-shape': if (!needSel()) return; return C.shapePicker(el, (k) => { E.selected().forEach(e => { if (e.type === 'text' || e.type === 'shape') { e.type = 'shape'; e.kind = k; if (k !== 'line' && k !== 'arrowLine' && (!e.fill || e.fill === 'none')) e.fill = 'accent1'; } }); SL.commit(); E.renderCanvas(); E.refreshThumb(true); });
      case 'insert-textbox': return E.setDrawMode({ type: 'text' });
      case 'arrange-menu': return C.arrangeMenu(el);
      case 'quick-styles': {
        const els = E.selected().filter(e => e.type === 'shape' || e.type === 'text'); if (!els.length) return BTS.toast('Select a shape or text box first');
        const qs = C.QUICK(P()); const box = document.createElement('div');
        return C.pop(el, `<div class="menu-head">Theme styles</div><div class="qs-grid">${qs.map((q, i) => { const demo = { id: 'q' + i, type: 'shape', kind: 'roundRect', fill: q.fill, stroke: q.stroke, strokeWidth: q.strokeWidth || 0 }; return `<button data-i="${i}" class="${q.shadow ? 'sh' : ''}" title="${q.label}">${SL.shapeSVG(demo, P(), 54, 34)}<span style="color:${SL.color(q.color, P())}">Abc</span></button>`; }).join('')}</div>`, 'qs-pop', (e) => { const b = e.target.closest('[data-i]'); if (b) { BTS.closePopover(); C.applyQuick(qs[+b.dataset.i]); } }), void box;
      }
      case 'shape-fill': { if (!needSel()) return; const c = await BTS.colorPicker(el, { none: 'No fill' }); if (!c) return; return C.setFill(c); }
      case 'shape-outline': {
        if (!needSel()) return;
        return BTS.menu(el, [
          { label: 'Outline color…', icon: 'palette', action: async () => { const c = await BTS.colorPicker(el, { none: 'No outline' }); if (!c) return; E.selected().forEach(e => { e.stroke = c; if (c !== 'none' && !e.strokeWidth) e.strokeWidth = 1.5; }); SL.commit(); E.renderCanvas(); E.refreshThumb(true); } },
          { label: 'No outline', icon: 'square-dashed', action: () => { E.selected().forEach(e => { if (!SL.isLine(e)) e.stroke = 'none'; }); SL.commit(); E.renderCanvas(); E.refreshThumb(true); } },
          { head: 'Weight' }, ...[.75, 1.5, 2.25, 3, 4.5, 6].map(w => ({ label: w + ' pt', checked: one.strokeWidth === w, action: () => { E.selected().forEach(e => { e.strokeWidth = w; if (!e.stroke || e.stroke === 'none') e.stroke = 'text'; }); SL.commit(); E.renderCanvas(); E.refreshThumb(true); } })),
          { head: 'Dashes' }, ...SL.DASHES.map(([k, n]) => ({ label: n, checked: (one.dash || 'solid') === k, action: () => { E.selected().forEach(e => { e.dash = k; if (!e.stroke || e.stroke === 'none') { e.stroke = 'text'; e.strokeWidth = e.strokeWidth || 1.5; } }); SL.commit(); E.renderCanvas(); E.refreshThumb(true); } })),
          { head: 'Arrows' }, { label: 'Arrow at end', checked: one.kind === 'arrowLine', disabled: !SL.isLine(one), action: () => { E.selected().forEach(e => { if (SL.isLine(e)) e.kind = e.kind === 'arrowLine' ? 'line' : 'arrowLine'; }); SL.commit(); E.renderCanvas(); E.refreshThumb(true); } },
        ]);
      }
      case 'shape-effects': {
        if (!needSel()) return;
        return BTS.menu(el, [
          { label: 'Shadow', checked: !!one.shadow, action: () => C.setElProp('shadow', !one.shadow) },
          { head: 'Transparency' }, ...[0, 25, 50, 75].map(t => ({ label: t + '%', checked: Math.round((1 - (one.opacity ?? 1)) * 100) === t, action: () => C.setElProp('opacity', 1 - t / 100) })),
          { sep: 1 }, { label: 'More options…', icon: 'settings-2', action: () => E.openPane('shape') },
        ]);
      }
      case 'text-effects': case 'wordart-style': {
        if (!C.textTargets().length) return BTS.toast('Select a text box first');
        const t = C.textTargets()[0];
        return C.pop(el, `<div class="menu-head">Text effects</div><div class="fx-text-grid">${SL.EFFECTS.map(([k, n]) => `<button data-k="${k}" class="${(t.effect || 'none') === k ? 'on' : ''}" title="${n}"><span style="font-family:'${SL.textDefaults(t, P()).font}';color:${SL.color('accent1', P())};${SL.effectCSS(k, SL.color('accent1', P()), P())}">A</span><small>${n}</small></button>`).join('')}</div>`, 'fx-pop', (e) => { const b = e.target.closest('[data-k]'); if (b) { BTS.closePopover(); C.setElProp('effect', b.dataset.k === 'none' ? undefined : b.dataset.k); } });
      }
      case 'autofit-menu': if (!C.textTargets().length) return BTS.toast('Select a text box first'); return BTS.menu(el, [['Do not autofit', ''], ['Shrink text on overflow', 'shrink'], ['Resize shape to fit text', 'resize']].map(([n, v]) => ({ label: n, checked: (one.autofit || '') === v, action: () => { E.selected().forEach(e => { if (v) e.autofit = v; else delete e.autofit; e.fontScale = 1; }); SL.commit(); E.renderCanvas(); E.refreshThumb(true); } })));
      /* ----- insert ----- */
      case 'insert-table': {
        const g = await BTS.gridPicker(el); if (!g) return;
        const W = SL.W(P()), H = SL.H(P()); const w = Math.min(W * .8, g.cols * 180), h = Math.min(H * .7, (g.rows) * 46);
        const target = S().elements.find(e => e.ph === 'body' && SL.isEmptyText(e));
        const t = SL.newTable(g.rows, g.cols, target ? { x: target.x, y: target.y, w: target.w, h } : { x: (W - w) / 2, y: (H - h) / 2, w, h, size: 16 });
        t.size = 16; if (target) S().elements = S().elements.filter(e => e !== target);
        E.addElement(t); E.startEdit(t.id); return;
      }
      case 'insert-picture': case 'pic-replace': {
        const f = await BTS.pickFile('image/*'); if (!f) return;
        const mime = f.type || ('image/' + (BTS.extOf(f.name) === 'jpg' ? 'jpeg' : BTS.extOf(f.name) === 'svg' ? 'svg+xml' : BTS.extOf(f.name) || 'png'));
        const url = await BTS.bytesToDataURL(f.data, mime);
        if (cmd === 'pic-replace' && one && one.type === 'image') { one.src = SL.addAsset(url); one.crop = null; SL.commit(); E.renderCanvas(); E.refreshThumb(true); return; }
        return E.insertImage(url, { name: f.name });
      }
      case 'insert-icon': {
        const n = await C.iconDialog(); if (!n) return;
        const W = SL.W(P()), H = SL.H(P());
        return E.addElement(SL.newElement('icon', { name: n, x: W / 2 - 60, y: H / 2 - 60, w: 120, h: 120, color: 'accent1' }));
      }
      case 'insert-chart': {
        const ch = await C.chartDialog(null); if (!ch) return;
        const W = SL.W(P()), H = SL.H(P());
        const target = S().elements.find(e => e.ph === 'body' && SL.isEmptyText(e));
        const g = target ? { x: target.x, y: target.y, w: target.w, h: target.h } : { x: W * .15, y: H * .18, w: W * .7, h: H * .7 };
        if (target) S().elements = S().elements.filter(e => e !== target);
        return E.addElement(SL.newElement('chart', Object.assign(g, { chart: ch })));
      }
      case 'chart-edit': { if (!one || one.type !== 'chart') return BTS.toast('Select a chart first'); const ch = await C.chartDialog(one.chart); if (!ch) return; one.chart = ch; SL.commit(); E.renderCanvas(); E.refreshThumb(true); return; }
      case 'chart-type': if (!one || one.type !== 'chart') return BTS.toast('Select a chart first'); return BTS.menu(el, [['Column', 'bar', 'chart-column'], ['Line', 'line', 'chart-line'], ['Pie', 'pie', 'chart-pie'], ['Doughnut', 'doughnut', 'circle']].map(([n, k, ic]) => ({ label: n, icon: ic, action: () => { one.chart.type = k; SL.commit(); E.renderCanvas(); E.refreshThumb(true); } })));
      case 'header-footer': return headerFooterDialog();
      case 'wordart': {
        const st = await new Promise(res => { const pp = C.pop(el, `<div class="menu-head">WordArt</div><div class="fx-text-grid">${SL.EFFECTS.filter(e => e[0] !== 'none').map(([k, n]) => `<button data-k="${k}" title="${n}"><span style="font-family:'${SL.fontOf('title', P())}';color:${SL.color('accent1', P())};${SL.effectCSS(k, SL.color('accent1', P()), P())}">A</span><small>${n}</small></button>`).join('')}</div>`, 'fx-pop', (e) => { const b = e.target.closest('[data-k]'); if (b) { res(b.dataset.k); BTS.closePopover(); } }); pp.closed.then(() => res(null)); });
        if (!st) return;
        const W = SL.W(P()), H = SL.H(P());
        return E.addElement(SL.newElement('text', { x: W * .15, y: H * .38, w: W * .7, h: 130, html: '<p>Your text here</p>', size: 66, bold: true, role: 'title', align: 'center', valign: 'middle', color: 'accent1', effect: st }), { edit: true });
      }
      case 'insert-date': case 'insert-slidenum': {
        const field = cmd === 'insert-date' ? 'date' : 'num';
        const html = `<span data-field="${field}">${field === 'num' ? SL.cur + 1 : new Date().toLocaleDateString()}</span>`;
        if (E.editing) { document.execCommand('insertHTML', false, html); E.syncEditing(); return; }
        const W = SL.W(P()), H = SL.H(P());
        return E.addElement(SL.newElement('text', { x: field === 'num' ? W - 200 : 60, y: H - 60, w: 140, h: 36, html: `<p>${html}</p>`, size: 14, align: field === 'num' ? 'right' : 'left', autofit: 'resize' }));
      }
      case 'insert-symbol': {
        const sym = await C.symbolDialog(); if (!sym) return;
        if (E.editing) { E.editNode().focus(); document.execCommand('insertText', false, sym); E.syncEditing(); return; }
        const t = C.textTargets()[0];
        if (t && t.type !== 'table') { t.html = (t.html || '').replace(/(<\/(p|li)>)(?![\s\S]*<\/(p|li)>)/, sym + '$1') || `<p>${sym}</p>`; SL.commit(); E.renderCanvas(); E.refreshThumb(true); return; }
        const W = SL.W(P()), H = SL.H(P());
        return E.addElement(SL.newElement('text', { x: W / 2 - 60, y: H / 2 - 40, w: 120, h: 80, html: `<p>${BTS.esc(sym)}</p>`, size: 48, align: 'center', autofit: 'resize' }));
      }
      case 'new-comment': {
        const t = await BTS.prompt('New comment', 'Comment', ''); if (!t) return;
        const at = one ? { x: one.x + one.w, y: one.y } : { x: 20 + (S().comments || []).length * 30, y: 20 };
        (S().comments = S().comments || []).push({ author: (await authorName()), text: t, time: Date.now(), x: at.x, y: at.y });
        SL.commit(); E.renderAll(); if (E.pane !== 'comments') E.openPane('comments'); return;
      }
      case 'comments-pane': return E.openPane('comments');
      case 'prev-comment': case 'next-comment': {
        const n = P().slides.length; const d = cmd === 'next-comment' ? 1 : -1;
        for (let k = 1; k <= n; k++) { const i = ((SL.cur + d * k) % n + n) % n; if ((P().slides[i].comments || []).length) { E.goto(i); if (E.pane !== 'comments') E.openPane('comments'); return; } }
        return BTS.toast('No other comments');
      }
      case 'delete-comments': S().comments = []; SL.commit(); E.renderAll(); return;
      /* ----- design ----- */
      case 'apply-theme': {
        if (E.editing) E.endEdit();
        const id = el.dataset.themeId; const was = P().theme.id; const before = { theme: SL.clone(P().theme), master: P().master ? SL.clone(P().master) : null };
        P().theme = { id, variant: 0 };
        if (P().master) { const m = P().master; delete m.bg; delete m.layoutBg; ['title', 'body'].forEach(k => { if (m[k]) { delete m[k].color; delete m[k].font; } }); }
        if (was === 'custom') P().slides.forEach(s => { s.layer = []; if (s.background && s.background.type !== 'image') s.background = null; });
        SL.syncPlaceholders(P(), before);
        SL.commit(); C.buildThemeGallery(); E.renderAll(); SL.ensureFonts(P()).then(() => E.renderAll()); return;
      }
      case 'apply-variant': P().theme = Object.assign({}, P().theme, { variant: +el.dataset.v }); delete P().theme.colors; SL.commit(); C.buildThemeGallery(); E.renderAll(); return;
      case 'themes-more': return C.pop(el, `<div class="menu-head">All themes (${SL.THEMES.length})</div><div class="theme-all">${SL.THEMES.map(t => `<button class="th-card ${P().theme.id === t.id ? 'on' : ''}" data-tid="${t.id}" title="${t.name}">${C.themeSample(t, 0, 150)}<span>${t.name}</span></button>`).join('')}</div>`, 'theme-pop', (e) => { const b = e.target.closest('[data-tid]'); if (b) { BTS.closePopover(); onCommand('apply-theme', { dataset: { themeId: b.dataset.tid } }); } });
      case 'slide-size': return BTS.menu(el, [
        { label: 'Widescreen (16:9)', checked: Math.abs(SL.W(P()) / SL.H(P()) - 16 / 9) < .02, action: () => setSize(1280, 720) },
        { label: 'Standard (4:3)', checked: Math.abs(SL.W(P()) / SL.H(P()) - 4 / 3) < .02, action: () => setSize(960, 720) },
        { label: 'A4 paper (landscape)', checked: Math.abs(SL.W(P()) / SL.H(P()) - 1.414) < .02, action: () => setSize(1123, 794) },
        { label: 'Custom slide size…', icon: 'ratio', action: customSize },
      ]);
      case 'format-bg': return E.openPane('bg');
      /* ----- transitions ----- */
      case 'set-transition': { const t = el.dataset.fx; [...E.railSel].forEach(i => { const s = P().slides[i]; s.transition = Object.assign({ dur: 700 }, s.transition, { type: t }); if (t === 'none') delete s.transition.dir; }); SL.commit(); E.renderRail(); E.updateRibbon(); if (t !== 'none') FX.previewTransition(); return; }
      case 'trans-options': { const t = S().transition || {}; const dirs = ['push', 'wipe', 'cover', 'reveal', 'flip', 'cube', 'split'].includes(t.type) ? (t.type === 'split' ? [['Vertical out', 'left'], ['Horizontal out', 'up']] : t.type === 'flip' || t.type === 'cube' ? [['From right', 'left'], ['From left', 'right']] : [['From right', 'left'], ['From left', 'right'], ['From bottom', 'up'], ['From top', 'down']]) : null; if (!dirs) return BTS.toast('This transition has no options'); return BTS.menu(el, dirs.map(([n, d]) => ({ label: n, checked: (t.dir || 'left') === d, action: () => { [...E.railSel].forEach(i => P().slides[i].transition.dir = d); SL.commit(); FX.previewTransition(); } }))); }
      case 'trans-dur': { const v = Math.max(.05, +el.value || .7); [...E.railSel].forEach(i => { const s = P().slides[i]; s.transition = Object.assign({ type: 'none' }, s.transition, { dur: Math.round(v * 1000) }); }); SL.commit(); return refocus(); }
      case 'trans-apply-all': { const t = SL.clone(S().transition || { type: 'none', dur: 700 }); P().slides.forEach(s => s.transition = SL.clone(t)); SL.commit(); E.renderRail(); BTS.toast('Transition applied to all slides'); return; }
      case 'trans-click': [...E.railSel].forEach(i => { const s = P().slides[i]; s.transition = Object.assign({ type: 'none', dur: 700 }, s.transition); if (el.checked) delete s.transition.click; else s.transition.click = false; }); SL.commit(); return;
      case 'trans-auto-on': case 'trans-auto': { const on = $('#trans-auto-on').checked; const v = Math.max(0, +$('#trans-auto').value || 0); if (cmd === 'trans-auto' && !on) $('#trans-auto-on').checked = true; [...E.railSel].forEach(i => { const s = P().slides[i]; s.transition = Object.assign({ type: 'none', dur: 700 }, s.transition); if (on || cmd === 'trans-auto') s.transition.auto = v; else delete s.transition.auto; }); SL.commit(); E.renderRail(); return refocus(); }
      case 'preview-transition': return FX.previewTransition();
      /* ----- animations ----- */
      case 'set-anim': {
        if (!needSel('Select an object to animate')) return; const t = el.dataset.fx;
        const list = C.animOrder(S()); let next = list.length ? Math.max(...list.map(e => e.anim.order ?? 0)) + 1 : 0;
        E.selected().forEach(e => { if (t === 'none') delete e.anim; else { const had = !!e.anim; e.anim = Object.assign({ trigger: 'click', dur: 500, order: had ? e.anim.order : next++ }, e.anim, { type: t }); if (t === 'appear') e.anim.dur = 1; else if (e.anim.dur <= 1) e.anim.dur = 500; } });
        C.renumberAnims(); SL.commit(); E.renderAll(); if (t !== 'none') FX.previewAnimations(); return;
      }
      case 'anim-options': { if (!one || !one.anim) return BTS.toast('Select an animated object first'); if (!['fly', 'wipe', 'float'].includes(one.anim.type)) return BTS.toast('This effect has no options'); return BTS.menu(el, [['From bottom', 'bottom'], ['From top', 'top'], ['From left', 'left'], ['From right', 'right']].filter(d => one.anim.type !== 'float' || ['bottom', 'top'].includes(d[1])).map(([n, d]) => ({ label: n, checked: (one.anim.dir || 'bottom') === d, action: () => { E.selected().forEach(e => e.anim && (e.anim.dir = d)); SL.commit(); FX.previewAnimations(); } }))); }
      case 'anim-trigger': E.selected().forEach(e => e.anim && (e.anim.trigger = el.value)); SL.commit(); E.refreshPane(); return refocus();
      case 'anim-dur': E.selected().forEach(e => e.anim && (e.anim.dur = Math.max(50, Math.round((+el.value || .5) * 1000)))); SL.commit(); E.refreshPane(); return refocus();
      case 'anim-delay': E.selected().forEach(e => e.anim && (e.anim.delay = Math.max(0, Math.round((+el.value || 0) * 1000)))); SL.commit(); E.refreshPane(); return refocus();
      case 'anim-earlier': return C.moveAnim(-1);
      case 'anim-later': return C.moveAnim(1);
      case 'anim-pane': return E.openPane('anim');
      case 'preview-anim': return FX.previewAnimations();
      /* ----- slide show ----- */
      case 'present-start': return FX.start(0);
      case 'present-current': return FX.start(SL.cur, { allowHidden: true });
      case 'present-window': return FX.start(SL.cur, { windowed: true, allowHidden: true });
      case 'show-loop': case 'show-timings': case 'show-anims': return;
      /* ----- review ----- */
      case 'spelling': SL.spell = SL.spell === false; $('#notes').spellcheck = SL.spell !== false; E.renderCanvas(); E.updateRibbon(); BTS.toast(SL.spell !== false ? 'Spell check is on (misspelled words are underlined while you type)' : 'Spell check is off'); return;
      case 'word-count': { let w = 0, c = 0, nw = 0; P().slides.forEach(s => { s.elements.forEach(e => { const t = e.type === 'table' ? e.rows.flat().map(x => SL.plainText(x.html)).join(' ') : SL.plainText(e.html || ''); w += (t.match(/\S+/g) || []).length; c += t.replace(/\s/g, '').length; }); nw += ((s.notes || '').match(/\S+/g) || []).length; }); return BTS.dialog({ title: 'Word count', body: `<div class="wc"><div><b>${P().slides.length}</b><span>Slides</span></div><div><b>${w.toLocaleString()}</b><span>Words on slides</span></div><div><b>${c.toLocaleString()}</b><span>Characters (no spaces)</span></div><div><b>${nw.toLocaleString()}</b><span>Words in notes</span></div></div>` }); }
      /* ----- view ----- */
      case 'view-normal': return E.setView('normal');
      case 'view-sorter': return E.setView('sorter');
      case 'view-notes': return E.setView(E.view === 'notes' ? 'normal' : 'notes');
      case 'toggle-ruler': E.ruler = !E.ruler; $('#v-ruler').checked = E.ruler; requestAnimationFrame(E.layout); return;
      case 'toggle-grid': E.grid = !E.grid; $('#v-grid').checked = E.grid; E.renderCanvas(); return;
      case 'toggle-guides': E.guides = !E.guides; $('#v-guides').checked = E.guides; E.renderOverlay(); return;
      case 'toggle-snap': E.snap = el.checked; return;
      case 'toggle-notes': case 'toggle-notes-check': E.notesOn = cmd === 'toggle-notes' ? !E.notesOn : el.checked; $('#v-notes').checked = E.notesOn; E.renderNotes(); E.updateStatus(); requestAnimationFrame(E.layout); return;
      case 'zoom-fit': return zoomFit();
      case 'zoom-dialog': { const v = await BTS.dialog({ title: 'Zoom', body: `<div class="zoom-opts">${['Fit', 400, 200, 150, 100, 75, 66, 50, 33].map(z => `<label class="check"><input type="radio" name="zm" value="${z}" ${(z === 'Fit' && E.fit) || (!E.fit && Math.round(E.zoom * 100) === z) ? 'checked' : ''}> ${z === 'Fit' ? 'Fit to window' : z + '%'}</label>`).join('')}</div><div class="field"><label>Percent</label><input type="number" id="zm-pct" min="10" max="400" value="${Math.round(E.zoom * 100)}"></div>`, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => { const ch = $('input[name=zm]:checked', r); return ch && ch.value === 'Fit' ? 'fit' : +$('#zm-pct', r).value; } }], onOpen: (r) => r.addEventListener('change', e => { if (e.target.name === 'zm' && e.target.value !== 'Fit') $('#zm-pct', r).value = e.target.value; }) }); if (v === 'fit') zoomFit(); else if (v) { E.fit = false; BTS.setZoom(v); } return; }
      case 'close-pane': return E.closePane();
      /* ----- format tab ----- */
      case 'crop': if (E.cropping) return E.endCrop(); if (!one || one.type !== 'image' || !one.src) return BTS.toast('Select a picture first'); return E.startCrop(one.id);
      case 'pic-corners': if (!one || one.type !== 'image') return BTS.toast('Select a picture first'); return BTS.menu(el, [['Square', 0], ['Slightly rounded', 6], ['Rounded', 14], ['Very rounded', 28], ['Circle / oval', 50]].map(([n, v]) => ({ label: n, checked: (one.radius || 0) === v, action: () => C.setElProp('radius', v, ['image']) })));
      case 'pic-reset': if (!one || one.type !== 'image') return; one.crop = null; if (one.natW) { const k = Math.min(1, SL.W(P()) * .8 / one.natW, SL.H(P()) * .8 / one.natH); one.w = one.natW * k; one.h = one.natH * k; } one.rot = 0; one.radius = 0; SL.commit(); E.renderCanvas(); E.refreshThumb(true); return;
      case 'tbl-row-above': case 'tbl-row-below': case 'tbl-del-row': case 'tbl-col-left': case 'tbl-col-right': case 'tbl-del-col': return tableOp(cmd);
      case 'tbl-header': case 'tbl-banded': if (!one || one.type !== 'table') return; one[cmd === 'tbl-header' ? 'header' : 'banded'] = el.checked; E.syncEditing(); SL.commit(); E.renderCanvas(); E.refreshThumb(true); return;
      case 'tbl-color': { if (!one || one.type !== 'table') return; const c = await BTS.colorPicker(el, {}); if (!c) return; E.syncEditing(); one.accent = c; SL.commit(); E.renderCanvas(); E.refreshThumb(true); return; }
      case 'tbl-cell-fill': { if (!one || one.type !== 'table') return BTS.toast('Select a table first'); const tc = E.tableCell || { r: 0, c: 0 }; const c = await BTS.colorPicker(el, { auto: 'Table style', none: 'No fill' }); if (!c) return; E.syncEditing(); const cell = one.rows[tc.r][tc.c]; if (c === 'auto') delete cell.fill; else cell.fill = c; SL.commit(); E.renderCanvas(); E.refreshThumb(true); return; }
      case 'bring-forward-menu': return BTS.menu(el, [{ label: 'Bring Forward', icon: 'arrow-up-from-line', action: () => E.zorder('forward') }, { label: 'Bring to Front', icon: 'bring-to-front', action: () => E.zorder('front') }]);
      case 'send-backward-menu': return BTS.menu(el, [{ label: 'Send Backward', icon: 'arrow-down-from-line', action: () => E.zorder('backward') }, { label: 'Send to Back', icon: 'send-to-back', action: () => E.zorder('back') }]);
      case 'selection-pane': return E.openPane('selection');
      case 'align-menu': return C.alignMenu(el);
      case 'group-menu': return BTS.menu(el, [{ label: 'Group', icon: 'group', kbd: 'Ctrl+G', disabled: E.sel.length < 2, action: () => E.group() }, { label: 'Ungroup', icon: 'ungroup', kbd: 'Ctrl+Shift+G', disabled: !E.selected().some(e => e.g), action: () => E.ungroup() }]);
      case 'rotate-menu': return C.rotateMenu(el);
      case 'size-w': case 'size-h': case 'size-lock': {
        if (!one) return; if (cmd === 'size-lock') return C.setElProp('lockAspect', el.checked);
        const v = Math.max(.05, +el.value || .05) * 96; const lock = $('#size-lock').checked;
        E.selected().forEach(e => { const r = e.h / e.w; if (cmd === 'size-w') { e.w = v; if (lock) e.h = v * r; } else { e.h = v; if (lock) e.w = v / r; } });
        SL.commit(); E.renderCanvas(); E.refreshThumb(true); E.updateRibbon(); return refocus();
      }
    }
  }
  function refocus() { if (E.editing) { const n = E.editNode(); n && n.focus(); } else $('#stage').focus({ preventScroll: true }); }
  const railFocused = () => document.activeElement && (document.activeElement.id === 'rail' || document.activeElement.id === 'sorter');
  async function authorName() { try { const s = BTS.desktop ? await bts.getSettings() : {}; return s.author || 'You'; } catch { return 'You'; } }

  function setSize(w, h) {
    const W = SL.W(P()), H = SL.H(P()); if (W === w && H === h) return;
    const sx = w / W, sy = h / H, s = Math.min(sx, sy), ox = (w - W * s) / 2, oy = (h - H * s) / 2;
    const fix = (e) => { e.x = e.x * s + ox; e.y = e.y * s + oy; e.w *= s; e.h *= s; if (e.size) e.size = Math.round(e.size * s * 2) / 2; };
    P().slides.forEach(sl => { sl.elements.forEach(fix); (sl.layer || []).forEach(fix); (sl.comments || []).forEach(c => { c.x = (c.x || 0) * s + ox; c.y = (c.y || 0) * s + oy; }); });
    P().width = w; P().height = h; SL.commit(); C.buildThemeGallery(); E.renderAll(); requestAnimationFrame(E.layout);
  }
  async function customSize() {
    const r = await BTS.dialog({ title: 'Slide size', body: `<div class="field-row"><div class="field"><label>Width (in)</label><input type="number" id="ss-w" step="0.01" min="1" max="56" value="${f2(SL.W(P()) / 96)}"></div><div class="field"><label>Height (in)</label><input type="number" id="ss-h" step="0.01" min="1" max="56" value="${f2(SL.H(P()) / 96)}"></div></div><p class="pf-note">Content is scaled to fit the new size.</p>`, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (root) => [+$('#ss-w', root).value, +$('#ss-h', root).value] }] });
    if (r && r[0] > 0 && r[1] > 0) setSize(Math.round(r[0] * 96), Math.round(r[1] * 96));
  }
  async function headerFooterDialog() {
    const f = P().footer = Object.assign({ date: false, number: false, text: '', skipTitle: true, dateText: '' }, P().footer);
    const r = await BTS.dialog({ title: 'Header & footer', width: 460, body: `
      <label class="check"><input type="checkbox" id="hf-date" ${f.date ? 'checked' : ''}> Date and time</label>
      <div class="field" style="margin-left:24px"><input id="hf-datet" placeholder="Leave empty for today's date" value="${BTS.esc(f.dateText || '')}"></div>
      <label class="check"><input type="checkbox" id="hf-num" ${f.number ? 'checked' : ''}> Slide number</label>
      <label class="check"><input type="checkbox" id="hf-ftr" ${f.text ? 'checked' : ''}> Footer</label>
      <div class="field" style="margin-left:24px"><input id="hf-text" value="${BTS.esc(f.text || '')}" placeholder="Footer text, e.g. your name or school"></div>
      <label class="check"><input type="checkbox" id="hf-skip" ${f.skipTitle ? 'checked' : ''}> Don't show on title slide</label>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'Apply to all', primary: true, value: (root) => ({ date: $('#hf-date', root).checked, dateText: $('#hf-datet', root).value, number: $('#hf-num', root).checked, text: $('#hf-ftr', root).checked ? $('#hf-text', root).value : '', skipTitle: $('#hf-skip', root).checked }) }] });
    if (!r) return; P().footer = r; SL.commit(); E.renderAll();
  }
  function tableOp(cmd) {
    const t = E.primary(); if (!t || t.type !== 'table') return BTS.toast('Select a table first');
    E.syncEditing();
    const cur = E.tableCell || { r: t.rows.length - 1, c: t.rows[0].length - 1 };
    const nc = t.rows[0].length;
    if (cmd === 'tbl-row-above' || cmd === 'tbl-row-below') { const at = cur.r + (cmd === 'tbl-row-below' ? 1 : 0); t.rows.splice(at, 0, Array.from({ length: nc }, () => ({ html: '' }))); t.h += t.h / (t.rows.length - 1); if (t.rowH) { t.rowH.splice(at, 0, 1 / t.rows.length); norm(t.rowH); } E.tableCell = { r: at, c: cur.c }; }
    if (cmd === 'tbl-del-row') { if (t.rows.length <= 1) { E.deleteSelected(); return; } t.h -= t.h / t.rows.length; t.rows.splice(cur.r, 1); if (t.rowH) { t.rowH.splice(cur.r, 1); norm(t.rowH); } E.tableCell = { r: Math.min(cur.r, t.rows.length - 1), c: cur.c }; }
    if (cmd === 'tbl-col-left' || cmd === 'tbl-col-right') { const at = cur.c + (cmd === 'tbl-col-right' ? 1 : 0); t.rows.forEach(r => r.splice(at, 0, { html: '' })); t.colW.splice(at, 0, 1 / (nc + 1)); norm(t.colW); E.tableCell = { r: cur.r, c: at }; }
    if (cmd === 'tbl-del-col') { if (nc <= 1) { E.deleteSelected(); return; } t.rows.forEach(r => r.splice(cur.c, 1)); t.colW.splice(cur.c, 1); norm(t.colW); E.tableCell = { r: cur.r, c: Math.min(cur.c, nc - 2) }; }
    SL.commit(); E.renderCanvas(); E.refreshThumb(true); E.updateRibbon();
    function norm(a) { const s = a.reduce((x, y) => x + y, 0); a.forEach((v, i) => a[i] = v / s); }
  }

  /* ====================================================================== */
  /*  Keyboard                                                               */
  /* ====================================================================== */
  function onKey(e) {
    if (FX.show.on) return;
    if (document.querySelector('.modal-back')) return;
    const mod = e.ctrlKey || e.metaKey; const k = e.key;
    const tgt = e.target; const inField = tgt.matches && tgt.matches('input, textarea, select, #notes, .notes-page-text, [contenteditable]') && !tgt.closest('#slide-host');
    if (k === 'F5') { e.preventDefault(); FX.start(e.shiftKey ? SL.cur : 0, { allowHidden: e.shiftKey }); return; }
    if (inField) { if (mod && (k === 'z' || k === 'y') && (tgt.id === 'notes' || tgt.classList.contains('notes-page-text'))) return; if (!mod || !['m'].includes(k.toLowerCase())) return; }
    const editingText = !!E.editing;
    if (mod && !e.altKey) {
      const lk = k.toLowerCase();
      if (lk === 'z' && !e.shiftKey) { e.preventDefault(); return onCommand('undo'); }
      if (lk === 'y' || (lk === 'z' && e.shiftKey)) { e.preventDefault(); return onCommand('redo'); }
      if (lk === 'm') { e.preventDefault(); return E.newSlide(); }
      if (lk === 'b' || lk === 'i' || lk === 'u') { if (editingText || C.textTargets().length) { e.preventDefault(); return C.toggleStyle({ b: 'bold', i: 'italic', u: 'underline' }[lk]); } }
      if (lk === 'l' || lk === 'e' || lk === 'r' || lk === 'j') { if (editingText || C.textTargets().length) { e.preventDefault(); return C.setAlign({ l: 'left', e: 'center', r: 'right', j: 'justify' }[lk]); } }
      if (k === ']' && !e.shiftKey) { e.preventDefault(); return editingText || C.textTargets().length && !e.shiftKey ? C.stepFont(1) : E.zorder('forward'); }
      if (k === '[' && !e.shiftKey) { e.preventDefault(); return editingText || C.textTargets().length ? C.stepFont(-1) : E.zorder('backward'); }
      if (k === '}' || (k === ']' && e.shiftKey)) { e.preventDefault(); return E.zorder('front'); }
      if (k === '{' || (k === '[' && e.shiftKey)) { e.preventDefault(); return E.zorder('back'); }
      if (lk === 'f') { e.preventDefault(); return E.openPane('find'); }
      if (lk === 'h') { e.preventDefault(); return E.openPane('find'); }
      if (editingText) return; // let the browser handle copy/paste/select-all inside text
      if (lk === 'a') { e.preventDefault(); if (railFocused()) { E.railSel = new Set(P().slides.map((_, i) => i)); E.renderRail(); if (E.view === 'sorter') E.renderSorter(); return; } return E.selectAll(); }
      if (lk === 'd') { e.preventDefault(); return railFocused() || !E.sel.length ? E.duplicateSlides() : E.duplicateSelected(); }
      if (lk === 'g') { e.preventDefault(); return e.shiftKey ? E.ungroup() : E.group(); }
      if (lk === 'c') { e.preventDefault(); return onCommand('copy'); }
      if (lk === 'x') { e.preventDefault(); return onCommand('cut'); }
      if (k === 'ArrowUp' || k === 'ArrowDown') { if (!E.sel.length || railFocused() || e.shiftKey) { e.preventDefault(); return E.moveSlideBy(k === 'ArrowUp' ? -1 : 1); } }
      return;
    }
    if (editingText) {
      if (k === 'Escape') { e.preventDefault(); const id = E.editing; E.endEdit(); if (SL.findEl(id)) E.select([id]); return; }
      if (k === 'Tab') {
        const el = SL.findEl(E.editing);
        if (el && el.type === 'table') { e.preventDefault(); moveCell(el, e.shiftKey ? -1 : 1); return; }
        e.preventDefault(); return C.indent(e.shiftKey ? -1 : 1);
      }
      return;
    }
    if (E.cropping && (k === 'Escape' || k === 'Enter')) { e.preventDefault(); return E.endCrop(); }
    if (E.drawMode && k === 'Escape') { E.setDrawMode(null); return; }
    if (E.painter && k === 'Escape') { E.painter = null; $('#stage').classList.remove('painting'); return; }
    const railMode = railFocused() || (!E.sel.length && E.view === 'sorter');
    if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); if (E.sel.length && !railFocused()) return E.deleteSelected(); return E.deleteSlides(); }
    if (k === 'Escape') { if (E.sel.length) E.select([]); return; }
    if (k.startsWith('Arrow')) {
      e.preventDefault();
      if (E.sel.length && !railMode) { const d = e.shiftKey ? 10 : (e.altKey ? .5 : 1); return E.nudge(k === 'ArrowLeft' ? -d : k === 'ArrowRight' ? d : 0, k === 'ArrowUp' ? -d : k === 'ArrowDown' ? d : 0); }
      if (k === 'ArrowUp' || k === 'ArrowLeft') return E.goto(SL.cur - 1); return E.goto(SL.cur + 1);
    }
    if (k === 'PageDown') { e.preventDefault(); return E.goto(SL.cur + 1); }
    if (k === 'PageUp') { e.preventDefault(); return E.goto(SL.cur - 1); }
    if (k === 'Home' && !E.sel.length) { e.preventDefault(); return E.goto(0); }
    if (k === 'End' && !E.sel.length) { e.preventDefault(); return E.goto(P().slides.length - 1); }
    if (k === 'Enter' && railMode) { e.preventDefault(); return E.newSlide(); }
    if (k === 'Enter' && E.sel.length === 1) { e.preventDefault(); return E.startEdit(E.sel[0], { selectAll: false }); }
    if (k === 'Tab') { e.preventDefault(); const els = S().elements; if (!els.length) return; const i = els.findIndex(x => x.id === E.sel[0]); const n = els[((i + (e.shiftKey ? -1 : 1)) % els.length + els.length) % els.length]; return E.select([n.id]); }
    if (k === 'F2' && E.sel.length === 1) { e.preventDefault(); return E.startEdit(E.sel[0]); }
    // typing on a selected text object starts editing (replacing placeholder text)
    if (k.length === 1 && !e.altKey && E.sel.length === 1) {
      const el = E.primary(); if (el && (el.type === 'text' || (el.type === 'shape' && !SL.isLine(el)))) {
        e.preventDefault(); el.html = ''; E.startEdit(el.id); document.execCommand('insertText', false, k); return;
      }
    }
  }
  function moveCell(t, d) {
    const cur = E.tableCell || { r: 0, c: 0 }; const nc = t.rows[0].length; let idx = cur.r * nc + cur.c + d;
    if (idx >= t.rows.length * nc) { E.syncEditing(); t.rows.push(Array.from({ length: nc }, () => ({ html: '' }))); t.h += t.h / (t.rows.length - 1); if (t.rowH) t.rowH = null; E.tableCell = { r: t.rows.length - 1, c: 0 }; E.renderCanvas(); return; }
    idx = Math.max(0, idx); E.tableCell = { r: Math.floor(idx / nc), c: idx % nc };
    const cell = document.querySelector(`#slide-host .el[data-id="${t.id}"] td[data-r="${E.tableCell.r}"][data-c="${E.tableCell.c}"] .cell`);
    if (cell) { cell.focus(); const r = document.createRange(); r.selectNodeContents(cell); const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
  }

  /* ====================================================================== */
  /*  Init                                                                   */
  /* ====================================================================== */
  function blankThumb() {
    const p = SL.newPresentation('clean'); const s = p.slides[0];
    s.elements[0].html = ''; s.elements[1].html = '';
    return `<div class="tpl-thumb">${SL.thumbHTML(s, p, 146, 0)}</div>`;
  }
  function boot() {
    // datalists
    $('#size-list').innerHTML = SL.SIZES.map(s => `<option value="${s}">`).join('');
    BTS.hasContent = () => P().slides.length > 1 || P().slides.some(s => s.elements.some(e => (e.html && !SL.isEmptyText(e)) || !e.ph));
    BTS.init({
      app: 'slides', name: 'BTS Slides', letter: 'P', defaultTitle: 'Presentation',
      saveFormats: [{ ext: 'pptx', label: 'PowerPoint presentation (.pptx)', icon: 'presentation' }],
      openAccept: '.pptx',
      onCommand, load, save, newDocument, search, onZoom, zoom: 100, minZoom: 10, maxZoom: 400,
      exportPDF, afterExportPDF, infoHTML,
      blankThumb: blankThumb(),
      templates: SLTemplates.list.map(t => ({ name: t.name, thumbHTML: `<div class="tpl-thumb">${(() => { const p = t.build(); return SL.thumbHTML(p.slides[0], p, 146, 0) + `<div class="tpl-strip">${p.slides.slice(1, 4).map((s, i) => SL.thumbHTML(s, p, 46, i + 1)).join('')}</div>`; })()}</div>`, apply: () => setPresentation(t.build()) })),
      shortcuts: [['Ctrl+M', 'New slide'], ['F5 / Shift+F5', 'Slide show from start / current'], ['Ctrl+D', 'Duplicate object or slide'], ['Ctrl+G / Ctrl+Shift+G', 'Group / ungroup'], ['Arrows (Shift)', 'Nudge 1 px (10 px)'], ['Ctrl+↑ / Ctrl+↓', 'Move slide up / down'], ['Ctrl+] / Ctrl+[', 'Bigger / smaller text']],
      focus: () => $('#stage').focus(),
      onDropFile: (f, e) => { if (/^image\//.test(f.type) && e.target.closest('#stage')) { f.arrayBuffer().then(async b => { const url = await BTS.bytesToDataURL(new Uint8Array(b), f.type); const p = E.toSlide(e); E.insertImage(url, { at: p, name: f.name, auto: false }); }); return true; } return false; },
    });
    E._booted = true;
    C.buildGalleries();
    E.wire();
    SL.resetHistory();
    E.renderAll();
    document.addEventListener('keydown', onKey);
    document.addEventListener('selectionchange', () => { if (E.editing) { clearTimeout(boot._t); boot._t = setTimeout(E.updateRibbon, 60); } });
    document.addEventListener('paste', (e) => {
      if (FX.show.on || E.editing) return;
      const t = e.target; if (t.matches && t.matches('input, textarea, #notes, .notes-page-text')) return;
      e.preventDefault(); if (railFocused() && E.slideClip) return C.pasteSlides(); E.pasteFromSystem(e.clipboardData);
    });
    document.addEventListener('copy', (e) => { if (E.editing || FX.show.on) return; const t = e.target; if (t.matches && t.matches('input, textarea, #notes')) return; if (E.sel.length) { e.preventDefault(); E.copy(); e.clipboardData.setData('text/plain', 'BTS-SLIDES-CLIP:' + E.selected().map(x => SL.plainText(x.html || '')).join('\n')); } });
    // refresh rail etc. on every committed change made outside the editor paths
    SL.onChange = () => { };
    addEventListener('beforeunload', () => { });
    document.fonts && document.fonts.ready.then(() => { E.renderAll(); });
    // load every theme font once so the theme gallery previews are drawn in the right typefaces
    SL.ensureFonts([...new Set(SL.THEMES.flatMap(t => [t.titleFont, t.bodyFont]))]).then(() => { C.buildThemeGallery(); E.renderAll(); });
    requestAnimationFrame(E.layout);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot); else boot();
})();
