/* BTS Docs — application wiring: ribbon commands, menus, keyboard, clipboard, context menu,
 * ribbon state sync, galleries, BTS.init configuration. */
(function () {
  const D = window.Docs; const { $, $$, PX } = D; const esc = BTS.esc;
  const CARET = '';

  /* ================= galleries ================= */
  function renderStyles() {
    $('#styles-gallery').innerHTML = D.STYLES.map(s => `<button class="gi keepfocus" data-cmd="style" data-style="${s.id}" title="${s.name}"><span class="pv doc-content"><span class="pv-${s.id}" style="margin:0;padding:0;border:0;background:none;${s.id === 'title' ? 'font-size:17px' : s.id === 'h1' ? 'font-size:15px' : s.id === 'quote' ? 'font-size:12px' : s.id === 'code' ? 'font-size:11px' : 'font-size:12.5px'};display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${s.id === 'code' ? 'code()' : 'AaBbCc'}</span></span><small>${s.name}</small></button>`).join('');
    const g = $('#styles-gallery');
    g.addEventListener('mouseover', e => { const b = e.target.closest('.gi'); if (b && b.dataset.style !== hoverStyle) { hoverStyle = b.dataset.style; D.previewStyle(hoverStyle); } });
    g.addEventListener('mouseleave', () => { hoverStyle = null; D.endPreview(); });
  }
  let hoverStyle = null;
  function renderThemes() {
    $('#themes-gallery').innerHTML = D.THEMES.map(t => `<button class="gi keepfocus" data-cmd="theme" data-th="${t.id}" title="${t.name} — headings in ${t.head}, body in ${t.body}"><span class="th-aa" style="font-family:${D.fontStack(t.head)};color:${t.accent}">Aa</span><span class="th-bar">${t.colors.map(c => `<i style="background:${c}"></i>`).join('')}</span><small>${t.name}</small></button>`).join('');
  }

  /* ================= ribbon sync ================= */
  let syncQueued = false;
  D.onSelection = () => { if (syncQueued) return; syncQueued = true; requestAnimationFrame(() => { syncQueued = false; syncRibbon(); }); };
  function firstFamily(ff) { return (ff || '').split(',')[0].trim().replace(/^["']|["']$/g, ''); }
  function syncRibbon() {
    $$('#themes-gallery .gi').forEach(g => g.classList.toggle('on', g.dataset.th === D.design.theme));
    const r = D.range();
    D.updateTableTools(); D.markNav(); D.updateStatus();
    if (!r) return;
    ['bold', 'italic', 'underline', 'strikeThrough', 'subscript', 'superscript'].forEach(c => { const b = $(`[data-state="${c}"]`); let on = false; try { on = document.queryCommandState(c); } catch { } b.classList.toggle('on', on); });
    let el = r.startContainer; if (el.nodeType === 3) el = el.parentElement;
    if (r.startContainer.nodeType === 1 && r.startContainer.childNodes[r.startOffset] && r.startContainer.childNodes[r.startOffset].nodeType === 1 && r.collapsed === false) el = r.startContainer.childNodes[r.startOffset];
    const cs = getComputedStyle(el);
    const fn = $('#font-name'); if (document.activeElement !== fn) fn.value = firstFamily(cs.fontFamily);
    const fs = $('#font-size'); if (document.activeElement !== fs) fs.value = D.currentFontSize();
    const block = D.blockOf(r.startContainer);
    const al = block ? getComputedStyle(block).textAlign.replace('start', 'left').replace('end', 'right') : 'left';
    $$('[data-align]').forEach(b => b.classList.toggle('on', b.dataset.align === al));
    const st = D.styleOf(block); $$('#styles-gallery .gi').forEach(g => g.classList.toggle('on', g.dataset.style === st));
    const on = $('#styles-gallery .gi.on'); if (on && !hoverStyle) { const gal = $('#styles-gallery'); const row = Math.round((on.offsetTop - gal.firstElementChild.offsetTop) / 62); if (Math.floor(gal.scrollTop / 62) !== row) gal.scrollTop = row * 62; }
    const li = D.closest(r.startContainer, 'li'); const lt = li && li.parentElement.tagName;
    $('#btn-bullets').classList.toggle('on', lt === 'UL'); $('#btn-numbering').classList.toggle('on', lt === 'OL');
    if (block) {
      const bcs = getComputedStyle(block);
      const setv = (id, v) => { const i = $(id); if (document.activeElement !== i) i.value = v; };
      setv('#in-indent-left', +(parseFloat(block.style.marginLeft ? bcs.marginLeft : 0) / PX * 2.54).toFixed(2));
      setv('#in-indent-right', +(parseFloat(block.style.marginRight ? bcs.marginRight : 0) / PX * 2.54).toFixed(2));
      setv('#in-space-before', Math.round(parseFloat(bcs.marginTop) * 0.75));
      setv('#in-space-after', Math.round(parseFloat(bcs.marginBottom) * 0.75));
      const lh = parseFloat(bcs.lineHeight) / parseFloat(bcs.fontSize); const opts = [1, 1.15, 1.5, 2, 2.5, 3]; const near = opts.reduce((a, b) => Math.abs(b - lh) < Math.abs(a - lh) ? b : a, 1.15);
      setv('#sel-line', String(near));
    }
    const cmt = D.closest(r.startContainer, '.cmt'); if (cmt) D.activateComment(cmt.dataset.cid);
    if (D.selImg && !(r.startContainer.nodeType === 1 && r.startContainer.childNodes[r.startOffset] === D.selImg)) { /* keep picture selected until click elsewhere */ }
  }
  D.syncRibbon = syncRibbon;

  /* ================= menus ================= */
  const check = (on) => on;
  function fontMenu(anchor) {
    const cur = $('#font-name').value; const th = D.theme();
    const items = [{ head: 'Theme fonts' }, { html: `<span class="fprev" style="font-family:${D.fontStack(D.design.headFont || th.head)}">${esc(D.design.headFont || th.head)}</span><span class="kbd">Headings</span>`, action: () => D.setFontName(D.design.headFont || th.head) }, { html: `<span class="fprev" style="font-family:${D.fontStack(D.design.bodyFont || th.body)}">${esc(D.design.bodyFont || th.body)}</span><span class="kbd">Body</span>`, action: () => D.setFontName(D.design.bodyFont || th.body) }, { head: 'All fonts' }];
    D.FONTS.forEach(f => items.push({ checked: f === cur, html: `<span class="fprev" style="font-family:${D.fontStack(f)}">${esc(f)}</span>${f === 'Nirmala UI' || f === 'Mangal' ? '<span class="kbd">हिन्दी मराठी</span>' : ''}`, action: () => D.setFontName(f) }));
    BTS.menu(anchor, items);
  }
  function sizeMenu(anchor) { const cur = +$('#font-size').value; BTS.menu(anchor, D.SIZES.map(s => ({ label: String(s), checked: s === cur, action: () => D.setFontSize(s) }))); }
  function caseMenu(anchor) { BTS.menu(anchor, [{ label: 'Sentence case.', action: () => D.changeCase('sentence') }, { label: 'lowercase', action: () => D.changeCase('lower') }, { label: 'UPPERCASE', action: () => D.changeCase('upper') }, { label: 'Capitalize Each Word', action: () => D.changeCase('title') }, { label: 'tOGGLE cASE', action: () => D.changeCase('toggle') }]); }
  function bulletsMenu(anchor) {
    const lib = [['disc', '●'], ['circle', '○'], ['square', '■'], ['"–  "', '–'], ['"➢  "', '➢'], ['"✓  "', '✓']];
    const el = document.createElement('div'); el.innerHTML = '<div class="mprev-head">Bullet library</div><div class="mprev-list">' + `<button class="none" data-v="none">None</button>` + lib.map(([v, c]) => `<button data-v='${v}'>${c}  ─────\n${c}  ─────\n${c}  ─────</button>`).join('') + '</div>';
    el.addEventListener('mousedown', e => e.preventDefault()); el.addEventListener('click', e => { const b = e.target.closest('[data-v]'); if (b) BTS.resolvePopover(b.dataset.v); });
    BTS.popover(anchor, el).then(v => { if (!v) return; if (v === 'none') { const r = D.ensureSel(); const li = D.closest(r.startContainer, 'li'); if (li) D.toggleList(li.parentElement.tagName === 'OL'); return; } D.toggleList(false, v); });
  }
  function numberingMenu(anchor) {
    const lib = [['decimal', ['1.', '2.', '3.']], ['lower-alpha', ['a.', 'b.', 'c.']], ['upper-alpha', ['A.', 'B.', 'C.']], ['lower-roman', ['i.', 'ii.', 'iii.']], ['upper-roman', ['I.', 'II.', 'III.']], ['decimal-leading-zero', ['01.', '02.', '03.']], ['devanagari', ['१.', '२.', '३.']]];
    const el = document.createElement('div'); el.innerHTML = '<div class="mprev-head">Numbering library</div><div class="mprev-list">' + `<button class="none" data-v="none">None</button>` + lib.map(([v, n]) => `<button data-v="${v}">${n.map(x => x + ' ─────').join('\n')}</button>`).join('') + '</div><div class="menu-sep"></div><div class="menu-item" data-act="restart"><span>Restart numbering at…</span></div>';
    el.addEventListener('mousedown', e => e.preventDefault());
    el.addEventListener('click', e => { const b = e.target.closest('[data-v]'); if (b) BTS.resolvePopover(b.dataset.v); if (e.target.closest('[data-act=restart]')) BTS.resolvePopover('restart'); });
    BTS.popover(anchor, el).then(async v => {
      if (!v) return;
      if (v === 'none') { const r = D.ensureSel(); const li = D.closest(r.startContainer, 'li'); if (li) D.toggleList(li.parentElement.tagName === 'OL'); return; }
      if (v === 'restart') { const r = D.ensureSel(); const ol = D.closest(r.startContainer, 'ol'); if (!ol) { BTS.toast('Put the cursor in a numbered list first.'); return; } const sv = D.saveSel(); const n = await BTS.prompt('Set numbering value', 'Start at', ol.getAttribute('start') || '1'); D.restoreSel(sv); if (n && +n > 0) D.change(() => ol.setAttribute('start', String(Math.round(+n)))); return; }
      D.toggleList(true, v);
    });
  }
  function lineSpacingMenu(anchor) {
    const cur = +$('#sel-line').value;
    BTS.menu(anchor, [...[1, 1.15, 1.5, 2, 2.5, 3].map(v => ({ label: v.toFixed(v === 1.15 ? 2 : 1), checked: v === cur, action: () => D.lineSpacing(v) })), { sep: true },
    { label: 'Add space before paragraph', icon: 'arrow-up-to-line', action: () => D.paraSpacing('before', 12) }, { label: 'Remove space after paragraph', icon: 'arrow-down-to-line', action: () => D.paraSpacing('after', 0) }, { label: 'Add space after paragraph', icon: 'arrow-down-from-line', action: () => D.paraSpacing('after', 8) }, { sep: true }, { label: 'Line spacing options…', icon: 'settings-2', action: paragraphDialog }]);
  }
  function bordersMenu(anchor) { BTS.menu(anchor, [{ label: 'Bottom border', icon: 'panel-bottom', action: () => D.paraBorder('bottom') }, { label: 'Top border', icon: 'panel-top', action: () => D.paraBorder('top') }, { label: 'Left border (bar)', icon: 'panel-left', action: () => D.paraBorder('left') }, { label: 'Right border', icon: 'panel-right', action: () => D.paraBorder('right') }, { label: 'Box border', icon: 'square', action: () => D.paraBorder('box') }, { label: 'No border', icon: 'square-dashed', action: () => D.paraBorder('none') }, { sep: true }, { label: 'Horizontal line', icon: 'minus', action: D.insertHR }]); }
  function stylesMenu(anchor) { const cur = D.styleOf(D.blockOf((D.range() || D.ensureSel()).startContainer)); BTS.menu(anchor, [...D.STYLES.map(s => ({ label: s.name, checked: s.id === cur, action: () => D.setBlockStyle(s.id) })), { sep: true }, { label: 'Clear formatting', icon: 'remove-formatting', action: D.clearFormatting }]); }
  async function paragraphDialog() {
    const sv = D.saveSel(); const r = D.ensureSel(); const b = D.blockOf(r.startContainer); if (!b) return;
    const cs = getComputedStyle(b); const cm = (px) => +(parseFloat(px || 0) / PX * 2.54).toFixed(2); const pt = (px) => Math.round(parseFloat(px || 0) * 0.75);
    const ti = parseFloat(cs.textIndent) || 0;
    const v = await BTS.dialog({
      title: 'Paragraph', width: 480,
      body: `<div class="field"><label>Alignment</label><select id="pd-al">${['left', 'center', 'right', 'justify'].map(a => `<option value="${a}" ${cs.textAlign.replace('start', 'left') === a ? 'selected' : ''}>${a[0].toUpperCase() + a.slice(1)}</option>`).join('')}</select></div>
      <div class="dlg-grid"><div class="field"><label>Indent left (cm)</label><input id="pd-l" type="number" step="0.1" value="${b.style.marginLeft ? cm(cs.marginLeft) : 0}"></div><div class="field"><label>Indent right (cm)</label><input id="pd-r" type="number" step="0.1" value="${b.style.marginRight ? cm(cs.marginRight) : 0}"></div>
      <div class="field"><label>Special</label><select id="pd-sp"><option value="none">(none)</option><option value="first" ${ti > 0 ? 'selected' : ''}>First line</option><option value="hanging" ${ti < 0 ? 'selected' : ''}>Hanging</option></select></div><div class="field"><label>By (cm)</label><input id="pd-by" type="number" step="0.1" min="0" value="${Math.abs(cm(ti)) || 1.27}"></div>
      <div class="field"><label>Spacing before (pt)</label><input id="pd-b" type="number" step="1" min="0" value="${pt(cs.marginTop)}"></div><div class="field"><label>Spacing after (pt)</label><input id="pd-a" type="number" step="1" min="0" value="${pt(cs.marginBottom)}"></div>
      <div class="field"><label>Line spacing</label><select id="pd-ls">${[1, 1.15, 1.5, 2, 2.5, 3].map(x => `<option value="${x}" ${Math.abs(parseFloat(cs.lineHeight) / parseFloat(cs.fontSize) - x) < 0.05 ? 'selected' : ''}>${x}</option>`).join('')}</select></div></div>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (rt) => ({ al: $('#pd-al', rt).value, l: +$('#pd-l', rt).value, r: +$('#pd-r', rt).value, sp: $('#pd-sp', rt).value, by: +$('#pd-by', rt).value, b: +$('#pd-b', rt).value, a: +$('#pd-a', rt).value, ls: +$('#pd-ls', rt).value }) }],
    });
    D.restoreSel(sv); if (!v) return;
    D.change(() => D.blocks().forEach(bl => {
      bl.style.textAlign = v.al === 'left' ? '' : v.al; bl.style.marginLeft = v.l ? v.l + 'cm' : ''; bl.style.marginRight = v.r ? v.r + 'cm' : '';
      bl.style.textIndent = v.sp === 'first' ? v.by + 'cm' : v.sp === 'hanging' ? -v.by + 'cm' : ''; if (v.sp === 'hanging' && !v.l) bl.style.marginLeft = v.by + 'cm';
      bl.style.marginTop = v.b + 'pt'; bl.style.marginBottom = v.a + 'pt'; bl.style.lineHeight = String(v.ls);
    }));
  }
  function tableMenu(anchor) {
    const el = document.createElement('div');
    BTS.gridPicker(anchor).then(v => { if (v) D.insertTable(v.rows, v.cols, 't-plain'); });
    // add extra actions below the grid
    const pop = $('.pop'); if (pop) { pop.insertAdjacentHTML('beforeend', `<div class="menu-sep"></div><div class="menu-item" data-a="dlg">${icon('table', 16)}<span>Insert table…</span></div><div class="menu-item" data-a="quick">${icon('table-2', 16)}<span>Quick table (header row)</span></div>`); pop.addEventListener('mousedown', e => e.preventDefault()); pop.addEventListener('click', e => { const a = e.target.closest('[data-a]'); if (!a) return; BTS.closePopover(); if (a.dataset.a === 'dlg') D.insertTableDialog(); else D.insertTable(4, 3, 't-header'); }); }
  }
  function headerMenu(anchor, which) {
    const H = which === 'header';
    BTS.menu(anchor, [{ label: `Edit ${which}`, icon: 'pencil', action: () => D.editHF(which) }, { sep: true }, { head: 'Built-in' },
    { label: 'Blank', action: () => D.hfPreset(which, 'blank') }, { label: 'Document title (right)', action: () => D.hfPreset(which, 'title') }, { label: 'Centered title', action: () => D.hfPreset(which, 'centered') }, { label: 'Title with line', action: () => D.hfPreset(which, 'line') }, { label: 'Student details (name, roll no.)', action: () => D.hfPreset(which, 'student') }, { label: 'Today’s date', action: () => D.hfPreset(which, 'date') },
    ...(H ? [] : [{ label: 'Page X of Y (centered)', action: () => { D.hfPreset(which, 'remove'); D.insertPageNumber('bottom-center', 'nofm'); } }]),
    { sep: true }, { label: 'Different first page', checked: !!D.design.hdrFirst, action: () => D.setDesign({ hdrFirst: !D.design.hdrFirst }) },
    { label: `Remove ${which}`, icon: 'trash-2', action: () => D.hfPreset(which, 'remove') }]);
  }
  function pageNumMenu(anchor) { BTS.menu(anchor, [{ head: 'Top of page' }, { label: 'Left', action: () => D.insertPageNumber('top-left') }, { label: 'Center', action: () => D.insertPageNumber('top-center') }, { label: 'Right', action: () => D.insertPageNumber('top-right') }, { head: 'Bottom of page' }, { label: 'Left', action: () => D.insertPageNumber('bottom-left') }, { label: 'Center', action: () => D.insertPageNumber('bottom-center') }, { label: 'Right', action: () => D.insertPageNumber('bottom-right') }, { label: '“Page X of Y” (bottom center)', action: () => D.insertPageNumber('bottom-center', 'nofm') }, { label: '“Page X” (bottom right)', action: () => D.insertPageNumber('bottom-right', 'page') }, { sep: true }, { label: 'Remove page numbers', icon: 'trash-2', action: D.removePageNumbers }]); }
  function watermarkMenu(anchor) {
    const pre = ['CONFIDENTIAL', 'DRAFT', 'DO NOT COPY', 'SAMPLE', 'URGENT', 'ASAP'];
    BTS.menu(anchor, [{ head: 'Watermarks' }, ...pre.map(t => ({ label: t, checked: D.design.watermark && D.design.watermark.text === t, action: () => D.setDesign({ watermark: { text: t, color: '#c0c4cc', opacity: 0.4, layout: 'diagonal', font: 'Calibri' } }) })), { sep: true }, { label: 'Custom watermark…', icon: 'stamp', action: watermarkDialog }, { label: 'Remove watermark', icon: 'trash-2', disabled: !D.design.watermark, action: () => D.setDesign({ watermark: null }) }]);
  }
  async function watermarkDialog() {
    const w = D.design.watermark || { text: 'DRAFT', color: '#c0c4cc', opacity: 0.4, layout: 'diagonal', font: 'Calibri' };
    const body = document.createElement('div');
    body.innerHTML = `<div class="field"><label>Text</label><input id="wm-t" value="${esc(w.text)}" maxlength="40"></div>
      <div class="dlg-grid"><div class="field"><label>Font</label><select id="wm-f">${D.FONTS.map(f => `<option ${f === w.font ? 'selected' : ''}>${f}</option>`).join('')}</select></div><div class="field"><label>Layout</label><select id="wm-l"><option value="diagonal">Diagonal</option><option value="horizontal" ${w.layout === 'horizontal' ? 'selected' : ''}>Horizontal</option></select></div>
      <div class="field"><label>Colour</label><input id="wm-c" type="color" value="${w.color}" style="padding:2px"></div><div class="field"><label>Transparency</label><input id="wm-o" type="range" min="0.1" max="1" step="0.05" value="${w.opacity}" style="height:32px;accent-color:var(--accent)"></div></div>
      <div class="wm-prev" id="wm-p"><span></span></div>`;
    const upd = () => { const s = $('#wm-p span', body); s.textContent = $('#wm-t', body).value || ' '; s.style.color = $('#wm-c', body).value; s.style.opacity = $('#wm-o', body).value; s.style.fontFamily = D.fontStack($('#wm-f', body).value); s.style.transform = $('#wm-l', body).value === 'horizontal' ? 'none' : 'rotate(-30deg)'; };
    body.addEventListener('input', upd); body.addEventListener('change', upd); setTimeout(upd);
    const v = await BTS.dialog({ title: 'Custom watermark', body, width: 460, buttons: [{ label: 'Cancel', value: null }, { label: 'Apply', primary: true, value: (r) => ({ text: $('#wm-t', r).value.trim(), font: $('#wm-f', r).value, layout: $('#wm-l', r).value, color: $('#wm-c', r).value, opacity: +$('#wm-o', r).value }) }] });
    if (v) D.setDesign({ watermark: v.text ? v : null });
  }
  function pageBorderMenu(anchor) {
    const cur = D.design.border; const col = cur ? cur.color : '#1f2533';
    const set = (style, width) => D.setDesign({ border: style ? { style, width, color: col } : null });
    BTS.menu(anchor, [{ label: 'None', checked: !cur, action: () => set(null) }, { label: 'Thin line', checked: cur && cur.style === 'solid' && cur.width < 2, action: () => set('solid', 1) }, { label: 'Thick line', checked: cur && cur.style === 'solid' && cur.width >= 2, action: () => set('solid', 3) }, { label: 'Double line', checked: cur && cur.style === 'double', action: () => set('double', 4) }, { label: 'Dashed', checked: cur && cur.style === 'dashed', action: () => set('dashed', 1.5) }, { label: 'Dotted', checked: cur && cur.style === 'dotted', action: () => set('dotted', 2) }, { label: 'Framed (thick + thin)', checked: cur && cur.style === 'art', action: () => set('art', 3) }, { sep: true },
    { label: 'Border colour…', icon: 'palette', action: async () => { const c = await BTS.colorPicker(anchor, { value: col }); if (c && c[0] === '#') D.setDesign({ border: Object.assign({ style: 'solid', width: 1.5 }, D.design.border || {}, { color: c }) }); } }]);
  }
  function themeFontsMenu(anchor) {
    const pairs = [['Calibri Light', 'Calibri'], ['Cambria', 'Calibri'], ['Georgia', 'Verdana'], ['Times New Roman', 'Times New Roman'], ['Arial', 'Arial'], ['Garamond', 'Garamond'], ['Segoe UI', 'Segoe UI'], ['Trebuchet MS', 'Trebuchet MS'], ['Tahoma', 'Verdana'], ['Nirmala UI', 'Nirmala UI']];
    const th = D.theme(); const ch = D.design.headFont || th.head, cb = D.design.bodyFont || th.body;
    BTS.menu(anchor, [{ head: 'Heading / body fonts' }, { label: `Theme default (${th.head} / ${th.body})`, checked: !D.design.headFont, action: () => D.setDesign({ headFont: null, bodyFont: null }) }, ...pairs.map(([h, b]) => ({ checked: !!D.design.headFont && h === ch && b === cb, html: `<span style="font-family:${D.fontStack(h)};font-size:15px">${h}</span><span class="kbd" style="font-family:${D.fontStack(b)}">${b}</span>`, action: () => D.setDesign({ headFont: h, bodyFont: b }) }))]);
  }
  function spacingMenu(anchor) {
    const cur = D.design.pspacing; const opts = [['Default (8 pt, 1.15)', null], ['No paragraph space', { after: 0, line: 1, before: 0 }], ['Compact', { after: 4, line: 1, before: 0 }], ['Tight', { after: 6, line: 1.15, before: 0 }], ['Open', { after: 10, line: 1.15, before: 0 }], ['Relaxed', { after: 6, line: 1.5, before: 0 }], ['Double', { after: 8, line: 2, before: 0 }]];
    BTS.menu(anchor, opts.map(([l, v]) => ({ label: l, checked: JSON.stringify(cur || null) === JSON.stringify(v), action: () => D.setDesign({ pspacing: v }) })));
  }
  function marginsMenu(anchor) {
    const m = D.page.m; const cm = (x) => (x * 2.54).toFixed(2).replace(/\.?0+$/, '');
    BTS.menu(anchor, [...Object.entries(D.MARGINS).map(([k, v]) => ({ checked: v.t === m.t && v.b === m.b && v.l === m.l && v.r === m.r, html: `<span style="min-width:120px"><b style="font-weight:600">${v.label}</b><br><small style="color:var(--ink-3)">T ${cm(v.t)} cm · B ${cm(v.b)} · L ${cm(v.l)} · R ${cm(v.r)}</small></span>`, action: () => D.setMargins(v) })), { sep: true }, { label: 'Custom margins…', icon: 'settings-2', action: customMargins }]);
    const p = $('.pop'); if (p) $$('.menu-item', p).forEach(i => i.style.height = '42px');
  }
  async function customMargins() {
    const m = D.page.m; const cm = (x) => +(x * 2.54).toFixed(2);
    const v = await BTS.dialog({ title: 'Page setup', width: 440, body: `<div class="dlg-grid"><div class="field"><label>Top (cm)</label><input id="cm-t" type="number" step="0.1" min="0" value="${cm(m.t)}"></div><div class="field"><label>Bottom (cm)</label><input id="cm-b" type="number" step="0.1" min="0" value="${cm(m.b)}"></div><div class="field"><label>Left (cm)</label><input id="cm-l" type="number" step="0.1" min="0" value="${cm(m.l)}"></div><div class="field"><label>Right (cm)</label><input id="cm-r" type="number" step="0.1" min="0" value="${cm(m.r)}"></div><div class="field"><label>Header from top (cm)</label><input id="cm-h" type="number" step="0.1" min="0" value="${cm(D.page.hdrDist)}"></div><div class="field"><label>Footer from bottom (cm)</label><input id="cm-f" type="number" step="0.1" min="0" value="${cm(D.page.ftrDist)}"></div></div>`, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ['t', 'b', 'l', 'r', 'h', 'f'].map(k => Math.max(0, +$('#cm-' + k, r).value / 2.54)) }] });
    if (!v) return; const ps = D.paperSize();
    if (v[2] + v[3] > ps.w - 1 || v[0] + v[1] > ps.h - 1) { BTS.toast('Those margins leave no room for text.', { error: true }); return; }
    D.setPage({ m: { t: v[0], b: v[1], l: v[2], r: v[3] }, hdrDist: v[4], ftrDist: v[5] });
  }
  function orientationMenu(anchor) { BTS.menu(anchor, [{ label: 'Portrait', icon: 'rectangle-vertical', checked: D.page.orient === 'portrait', action: () => D.setPage({ orient: 'portrait' }) }, { label: 'Landscape', icon: 'rectangle-horizontal', checked: D.page.orient === 'landscape', action: () => D.setPage({ orient: 'landscape' }) }]); }
  function sizeMenuPage(anchor) { BTS.menu(anchor, Object.entries(D.PAPER).map(([k, p]) => ({ checked: D.page.size === k, html: `<span><b style="font-weight:600">${p.label}</b><br><small style="color:var(--ink-3)">${p.note}</small></span>`, action: () => D.setPage({ size: k }) }))); const p = $('.pop'); if (p) $$('.menu-item', p).forEach(i => i.style.height = '42px'); }
  function columnsMenu(anchor) { BTS.menu(anchor, [{ label: 'One', icon: 'rectangle-vertical', action: () => setColumns(1) }, { label: 'Two', icon: 'columns-2', action: () => setColumns(2) }, { label: 'Three', icon: 'columns-3', action: () => setColumns(3) }, { sep: true }, { label: 'Selected text only: select paragraphs first', disabled: true }]); }
  function setColumns(n) {
    D.ensureSel();
    D.change(() => {
      const r = D.range(); let tops;
      if (r && !r.collapsed && D.doc.contains(r.commonAncestorContainer)) { const a = D.topBlock(r.startContainer), b = D.topBlock(r.endContainer); tops = []; let x = a; while (x) { tops.push(x); if (x === b) break; x = x.nextSibling; } }
      else tops = [...D.doc.childNodes];
      // flatten existing column wrappers among targets
      const flat = []; tops.forEach(t => { if (t.nodeType === 1 && t.classList.contains('cols')) { flat.push(...t.childNodes); } else flat.push(t); });
      tops.forEach(t => { if (t.nodeType === 1 && t.classList.contains('cols')) { while (t.firstChild) t.parentNode.insertBefore(t.firstChild, t); t.remove(); } });
      if (n === 1) return;
      const nodes = flat.filter(x => x.isConnected && x.parentNode === D.doc); if (!nodes.length) return;
      const w = document.createElement('div'); w.className = 'cols cols-' + n; nodes[0].before(w); nodes.forEach(x => w.appendChild(x));
      if (!w.nextSibling) { const p = document.createElement('p'); p.innerHTML = '<br>'; w.after(p); }
    });
  }
  function breaksMenu(anchor) { BTS.menu(anchor, [{ label: 'Page break', icon: 'square-split-vertical', kbd: 'Ctrl+Enter', action: D.pageBreak }, { label: 'Blank page', icon: 'file-plus', action: D.blankPage }, { label: 'Line break', icon: 'corner-down-left', kbd: 'Shift+Enter', action: () => D.change(() => { D.ensureSel(); document.execCommand('insertLineBreak'); }) }]); }
  function tocMenu(anchor) { BTS.menu(anchor, [{ label: 'Automatic table — “Contents”', icon: 'table-of-contents', action: () => D.insertTOC('auto') }, { label: 'Automatic table — “Index of topics”', icon: 'list-tree', action: () => D.insertTOC('toc-index') }, { sep: true }, { label: 'Update table', icon: 'refresh-cw', action: () => D.updateTOC() }, { label: 'Remove table of contents', icon: 'trash-2', action: D.removeTOC }]); }
  function addTextMenu(anchor) { const cur = D.styleOf(D.blockOf((D.range() || D.ensureSel()).startContainer)); BTS.menu(anchor, [{ label: 'Do not show in contents', checked: !/^h[1-3]$/.test(cur), action: () => D.setBlockStyle('normal') }, { label: 'Level 1', checked: cur === 'h1', action: () => D.setBlockStyle('h1') }, { label: 'Level 2', checked: cur === 'h2', action: () => D.setBlockStyle('h2') }, { label: 'Level 3', checked: cur === 'h3', action: () => D.setBlockStyle('h3') }]); }
  function zoomMenu(anchor) { BTS.menu(anchor, [...[50, 75, 100, 125, 150, 200].map(z => ({ label: z + '%', checked: BTS.zoom === z, action: () => BTS.setZoom(z) })), { sep: true }, { label: 'Page width', icon: 'move-horizontal', action: () => D.fitZoom('width') }, { label: 'One page', icon: 'file', action: () => D.fitZoom('page') }]); }
  function tblDeleteMenu(anchor) { BTS.menu(anchor, [{ label: 'Delete cells (clear)', icon: 'eraser', action: () => D.tableOp('del-cells') }, { label: 'Delete columns', icon: 'columns-3', action: () => D.tableOp('del-col') }, { label: 'Delete rows', icon: 'rows-3', action: () => D.tableOp('del-row') }, { label: 'Delete table', icon: 'trash-2', action: () => D.tableOp('del-table') }]); }
  function tblBordersMenu(anchor) { BTS.menu(anchor, [{ label: 'All borders', icon: 'grid-3x3', action: () => D.tableOp('borders', 'all') }, { label: 'Outside borders only', icon: 'square', action: () => D.tableOp('borders', 'outside') }, { label: 'Thick borders', icon: 'grid-2x2', action: () => D.tableOp('borders', 'thick') }, { label: 'Box around selected cells', icon: 'square-dashed', action: () => D.tableOp('borders', 'cell-box') }, { label: 'No borders', icon: 'square-dashed', action: () => D.tableOp('borders', 'none') }]); }

  /* ================= clipboard ================= */
  const KEEP_CLASS = /^(title|subtitle|nospace|caption|tbl|tbl-auto|t-[a-z-]+|no-borders|img-(left|right|center|border|rounded|shadow)|page-break|toc|toc-[a-z0-9-]+|fld|eq|frac|fn-ref|textbox|tb-(left|inline)|cols|cols-[23]|dropcap|tab)$/;
  const KEEP_STYLE = /^(color|background-color|font-weight|font-style|text-decoration|text-decoration-line|font-size|font-family|text-align|vertical-align|margin-left|margin-right|text-indent|list-style-type|width|height|border(-[a-z]+)*|padding(-[a-z]+)*|line-height|margin-top|margin-bottom|text-transform)$/;
  function cleanPasted(html, internal) {
    let s = D.sanitize(html);
    const box = document.createElement('div'); box.innerHTML = s;
    $$('*', box).forEach(el => {
      if (el.className && typeof el.className === 'string') { const k = el.className.split(/\s+/).filter(c => KEEP_CLASS.test(c)); if (k.length) el.className = k.join(' '); else el.removeAttribute('class'); }
      const st = el.getAttribute('style');
      if (st) {
        const keep = st.split(';').map(x => x.trim()).filter(Boolean).filter(d => { const [p, v] = d.split(':').map(x => (x || '').trim().toLowerCase()); if (!KEEP_STYLE.test(p)) return false; if (!internal && /^(width|height)$/.test(p) && !/^(IMG|COL|TABLE|TD|TH)$/.test(el.tagName)) return false; if (!internal && p === 'font-family' && /var\(|inherit/.test(v)) return false; if (/var\(/.test(v)) return false; if (!internal && (p === 'margin-top' || p === 'margin-bottom' || p === 'line-height')) return false; if (!internal && p === 'background-color' && /rgba\(0, 0, 0, 0\)|transparent|^#?fff(fff)?$|rgb\(255, 255, 255\)/.test(v)) return false; if (!internal && p === 'color' && /rgb\(0, 0, 0\)|^#000(000)?$|windowtext/.test(v)) return false; return true; });
        if (keep.length) el.setAttribute('style', keep.join(';')); else el.removeAttribute('style');
      }
      if (el.id && !internal) el.removeAttribute('id');
      if (el.classList.contains('cmt')) D.unwrap(el);
    });
    $$('span:not([class]):not([style])', box).forEach(D.unwrap);
    $$('o\\:p', box).forEach(x => x.remove());
    return box.innerHTML;
  }
  D.pasteHTML = (html, internal) => {
    const clean = cleanPasted(html, internal);
    D.change(() => { D.ensureSel(); document.execCommand('insertHTML', false, clean); $$('p:empty,li:empty', D.doc).forEach(p => p.innerHTML = '<br>'); $$('#doc td', D.doc).forEach(c => { if (!c.firstChild) c.innerHTML = '<p><br></p>'; }); });
    D.renderFootnotes();
  };
  D.pasteText = (text) => {
    if (!text) return;
    D.change(() => { D.ensureSel(); if (/\n/.test(text.trim())) document.execCommand('insertHTML', false, D.textToHTML(text.replace(/\n$/, ''))); else document.execCommand('insertText', false, text.replace(/\r?\n/g, '')); });
  };
  async function pasteFromClipboard(plain) {
    D.ensureSel();
    try {
      if (plain) { D.pasteText(await navigator.clipboard.readText()); return; }
      const items = await navigator.clipboard.read();
      for (const it of items) {
        const img = it.types.find(t => t.startsWith('image/'));
        if (img) { const b = new Uint8Array(await (await it.getType(img)).arrayBuffer()); D.insertImageURL(await BTS.bytesToDataURL(b, img)); return; }
      }
      for (const it of items) if (it.types.includes('text/html')) { const h = await (await it.getType('text/html')).text(); D.pasteHTML(h, /bts-docs|class="(tbl|title|subtitle)/.test(h)); return; }
      for (const it of items) if (it.types.includes('text/plain')) { D.pasteText(await (await it.getType('text/plain')).text()); return; }
    } catch (e) {
      if (!document.execCommand('paste')) BTS.toast('Use Ctrl+V to paste (clipboard access was blocked).');
    }
  }
  function onPaste(e) {
    const cd = e.clipboardData; if (!cd) return;
    e.preventDefault();
    const files = [...cd.files].filter(f => f.type.startsWith('image/'));
    if (files.length) { files.forEach(async f => D.insertImageURL(await BTS.bytesToDataURL(new Uint8Array(await f.arrayBuffer()), f.type), f.name.replace(/\.[^.]+$/, ''))); return; }
    if (D._plainPaste) { D._plainPaste = false; D.pasteText(cd.getData('text/plain')); return; }
    const html = cd.getData('text/html');
    if (html) D.pasteHTML(html, /data-bts|class="(tbl|title|subtitle|toc)/.test(html) || html.includes('doc-content'));
    else D.pasteText(cd.getData('text/plain'));
  }
  function onCopy(e) {
    // add a marker so our own content pastes back with full fidelity
    const r = D.range(); if (!r || r.collapsed) return;
    const frag = r.cloneContents(); const div = document.createElement('div'); div.appendChild(frag);
    $$('.cell-sel', div).forEach(x => x.classList.remove('cell-sel'));
    // if selection is inside a single block, keep inline formatting from ancestors
    e.clipboardData.setData('text/html', '<meta name="bts-docs-clip" content="1">' + '<div data-bts="1">' + div.innerHTML + '</div>');
    e.clipboardData.setData('text/plain', r.toString());
    e.preventDefault();
    if (e.type === 'cut') D.change(() => { r.deleteContents(); });
  }

  /* ================= keyboard ================= */
  function atEndOf(block) { const r = D.range(); if (!r || !r.collapsed) return false; const t = document.createRange(); t.setStart(r.startContainer, r.startOffset); t.setEnd(block, block.childNodes.length); return !t.toString().replace(/​/g, '').length && !t.cloneContents().querySelector('img'); }
  function docKey(e) {
    const mod = e.ctrlKey || e.metaKey; const k = e.key;
    if (k === 'Escape') { if (D.selImg) { D.selectImage(null); } D.endPainter(); if (D.focusMode) toggleFocus(false); return; }
    if (D.selImg && (k === 'Delete' || k === 'Backspace') && !mod) { e.preventDefault(); D.imgOp('delete'); return; }
    if (k === 'Tab' && !mod && !e.altKey) {
      if (D.tableKey(e)) return;
      const r = D.range(); const li = r && D.closest(r.startContainer, 'li');
      e.preventDefault();
      if (li) { D.indent(e.shiftKey ? -1 : 1); return; }
      if (e.shiftKey) { D.indent(-1); return; }
      D.change(() => { const s = document.createElement('span'); s.className = 'tab'; s.textContent = '\t'; const rr = D.range(); rr.deleteContents(); rr.insertNode(s); const t = document.createTextNode('​'); s.after(t); D.setCaret(t, 1); });
      return;
    }
    if (k === 'Enter' && mod && !e.shiftKey) { e.preventDefault(); D.pageBreak(); return; }
    if (k === 'Enter' && !mod && !e.shiftKey) {
      const r = D.range(); const b = r && D.blockOf(r.startContainer);
      if (b && (/^H[1-6]$/.test(b.tagName) || b.classList.contains('title') || b.classList.contains('subtitle') || b.classList.contains('caption')) && atEndOf(b) && b.textContent.trim()) {
        e.preventDefault(); D.change(() => { const p = document.createElement('p'); p.innerHTML = '<br>'; b.after(p); D.setCaret(p, 0); }); D.scrollCaret(); return;
      }
      if (b && b.tagName === 'PRE') { e.preventDefault(); if (atEndOf(b) && /\n\n?$/.test(b.textContent) && b.textContent.endsWith('\n\n')) { D.change(() => { b.textContent = b.textContent.replace(/\n+$/, ''); const p = document.createElement('p'); p.innerHTML = '<br>'; b.after(p); D.setCaret(p, 0); }); return; } D.change(() => { document.execCommand('insertText', false, '\n'); if (atEndOf(b) && !b.textContent.endsWith('\n\n')) { /* need a trailing newline to show the caret line */ document.execCommand('insertText', false, '\n'); const rr = D.range(); rr.setStart(rr.startContainer, rr.startOffset - 1); rr.collapse(true); } }); return; }
      setTimeout(() => D.scrollCaret(), 0);
    }
    if (!mod) return;
    const lk = k.toLowerCase();
    if (e.altKey && ['1', '2', '3'].includes(k)) { e.preventDefault(); D.setBlockStyle('h' + k); return; }
    if (e.shiftKey && (k === '>' || k === '.')) { e.preventDefault(); D.growFont(1); return; }
    if (e.shiftKey && (k === '<' || k === ',')) { e.preventDefault(); D.growFont(-1); return; }
    if (e.shiftKey && lk === 'l') { e.preventDefault(); D.toggleList(false); return; }
    if (e.shiftKey && lk === 'n') { e.preventDefault(); D.setBlockStyle('normal'); return; }
    if (e.shiftKey && (k === '+' || k === '=')) { e.preventDefault(); D.inline('superscript'); return; }
    if (e.shiftKey && lk === 'v') { e.preventDefault(); pasteFromClipboard(true); return; }
    if (e.shiftKey) { if (lk === 'z') { e.preventDefault(); D.redo(); } return; }
    switch (lk) {
      case 'b': e.preventDefault(); D.inline('bold'); return;
      case 'i': e.preventDefault(); D.inline('italic'); return;
      case 'u': e.preventDefault(); D.inline('underline'); return;
      case '=': e.preventDefault(); D.inline('subscript'); return;
      case 'e': e.preventDefault(); D.align('center'); return;
      case 'l': e.preventDefault(); D.align('left'); return;
      case 'r': e.preventDefault(); D.align('right'); return;
      case 'j': e.preventDefault(); D.align('justify'); return;
      case 'k': e.preventDefault(); D.linkDialog(); return;
      case '1': e.preventDefault(); D.lineSpacing(1); return;
      case '2': e.preventDefault(); D.lineSpacing(2); return;
      case '5': e.preventDefault(); D.lineSpacing(1.5); return;
      case 'z': e.preventDefault(); D.undo(); return;
      case 'y': e.preventDefault(); D.redo(); return;
      case ']': e.preventDefault(); D.setFontSize(D.currentFontSize() + 1); return;
      case '[': e.preventDefault(); D.setFontSize(D.currentFontSize() - 1); return;
      case ' ': e.preventDefault(); D.inline('removeFormat'); return;
      case 'm': e.preventDefault(); D.indent(e.shiftKey ? -1 : 1); return;
      case 'd': e.preventDefault(); D.inline('strikeThrough'); return;
    }
  }
  function globalKey(e) {
    const mod = e.ctrlKey || e.metaKey; if (!mod) { if (e.key === 'F7') { e.preventDefault(); toggleSpelling(); } if (e.key === 'Escape' && D.focusMode && !document.querySelector('.modal-back')) toggleFocus(false); return; }
    const lk = e.key.toLowerCase();
    if (document.querySelector('.modal-back')) return;
    if (lk === 'f' && !e.shiftKey) { e.preventDefault(); D.openFind(false); }
    else if (lk === 'h') { e.preventDefault(); D.openFind(true); }
    else if (lk === 'g') { e.preventDefault(); gotoPage(); }
    else if (lk === 'a' && !D.rootOf(document.activeElement) && !/INPUT|TEXTAREA/.test(document.activeElement.tagName)) { e.preventDefault(); selectAll(); }
    else if ((lk === 'z' || lk === 'y') && !D.rootOf(document.activeElement) && !/INPUT|TEXTAREA/.test(document.activeElement.tagName)) { e.preventDefault(); lk === 'z' && !e.shiftKey ? D.undo() : D.redo(); }
  }
  function selectAll() { D.doc.focus(); const r = document.createRange(); r.selectNodeContents(D.doc); const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
  async function gotoPage() {
    const v = await BTS.prompt('Go to page', `Page number (1–${D.pageCount})`, String(D.caretPage())); if (!v) return;
    const n = Math.max(1, Math.min(D.pageCount, Math.round(+v) || 1)); const y = D.pages[n - 1];
    const dr = D.doc.getBoundingClientRect(); const s = D.zoom / 100;
    D.canvas.scrollTop += dr.top + y * s - D.canvas.getBoundingClientRect().top - 60;
    // caret to first block at/after y
    const blk = [...D.doc.querySelectorAll('p,h1,h2,h3,li,td')].find(b => (b.getBoundingClientRect().top - dr.top) / s >= y - 2); if (blk) D.setCaret(blk, 0);
  }

  /* ================= view toggles ================= */
  function toggleSpelling(on) {
    const now = on ?? !$('#btn-spell').classList.contains('on');
    $('#btn-spell').classList.toggle('on', now);
    [D.doc, D.hdr, D.ftr].forEach(el => el.spellcheck = now);
    document.body.classList.toggle('no-spell', !now);
    BTS.toast(now ? 'Spelling check is on' : 'Spelling check is off');
  }
  function toggleMarks(on) { const now = on ?? !document.body.classList.contains('show-marks'); document.body.classList.toggle('show-marks', now); $('#btn-marks').classList.toggle('on', now); $('#chk-marks').checked = now; }
  function toggleFocus(on) {
    D.focusMode = on ?? !D.focusMode; document.body.classList.toggle('focus-mode', D.focusMode); $('#btn-focus').classList.toggle('on', D.focusMode);
    let x = $('#focus-exit'); if (D.focusMode && !x) { document.body.insertAdjacentHTML('beforeend', `<button class="btn" id="focus-exit" data-cmd="focus-mode" title="Exit focus mode (Esc)">${icon('minimize-2', 15)} Exit focus</button>`); } else if (!D.focusMode && x) x.remove();
    D.scheduleLayout(30); D.updateStatus();
  }
  function toggleDarkCanvas() { const on = !document.body.classList.contains('dark-canvas'); document.body.classList.toggle('dark-canvas', on); $('#btn-dark-canvas').classList.toggle('on', on); }

  /* ================= context menu ================= */
  function contextMenu(e) {
    const root = D.rootOf(e.target); if (!root) return;
    e.preventDefault();
    const r = D.range(); const hasSel = r && !r.collapsed;
    const a = e.target.closest('a'); const img = e.target.tagName === 'IMG' ? e.target : null;
    const cell = e.target.closest('td,th'); const cmt = e.target.closest('.cmt'); const eq = e.target.closest('.eq'); const fn = e.target.closest('.fn-ref');
    const items = [{ label: 'Cut', icon: 'scissors', kbd: 'Ctrl+X', disabled: !hasSel && !img, action: () => document.execCommand('cut') }, { label: 'Copy', icon: 'copy', kbd: 'Ctrl+C', disabled: !hasSel && !img, action: () => document.execCommand('copy') }, { label: 'Paste', icon: 'clipboard-paste', kbd: 'Ctrl+V', action: () => pasteFromClipboard(false) }, { label: 'Paste as plain text', icon: 'clipboard-type', kbd: 'Ctrl+Shift+V', action: () => pasteFromClipboard(true) }, { sep: true }];
    if (img) items.push({ head: 'Picture' }, { label: 'In line with text', icon: 'text-cursor', action: () => D.imgOp('wrap', 'inline') }, { label: 'Wrap text — left', icon: 'wrap-text', action: () => D.imgOp('wrap', 'left') }, { label: 'Wrap text — right', icon: 'text-wrap', action: () => D.imgOp('wrap', 'right') }, { label: 'Centered', icon: 'align-center-horizontal', action: () => D.imgOp('wrap', 'center') }, { label: 'Alt text…', icon: 'message-square-text', action: D.imgAlt }, { label: 'Reset size', icon: 'rotate-ccw', action: () => D.imgOp('reset') }, { label: 'Delete picture', icon: 'trash-2', action: () => D.imgOp('delete') }, { sep: true });
    if (cell) items.push({ head: 'Table' }, { label: 'Insert row above', icon: 'between-horizontal-start', action: () => D.tableOp('row-above') }, { label: 'Insert row below', icon: 'between-horizontal-end', action: () => D.tableOp('row-below') }, { label: 'Insert column left', icon: 'between-vertical-start', action: () => D.tableOp('col-left') }, { label: 'Insert column right', icon: 'between-vertical-end', action: () => D.tableOp('col-right') }, { label: 'Merge cells', icon: 'table-cells-merge', disabled: $$('#doc .cell-sel').length < 2, action: () => D.tableOp('merge') }, { label: 'Split cell…', icon: 'table-cells-split', action: () => D.tableOp('split') }, { label: 'Delete row', icon: 'rows-3', action: () => D.tableOp('del-row') }, { label: 'Delete column', icon: 'columns-3', action: () => D.tableOp('del-col') }, { label: 'Delete table', icon: 'trash-2', action: () => D.tableOp('del-table') }, { sep: true });
    if (a) items.push({ label: 'Edit link…', icon: 'link', action: () => { D.selectNode(a, true); D.linkDialog(); } }, { label: 'Open link', icon: 'external-link', action: () => D.followLink(a) }, { label: 'Remove link', icon: 'unlink', action: () => D.change(() => D.unwrap(a)) }, { sep: true });
    else if (!img) items.push({ label: 'Link…', icon: 'link', kbd: 'Ctrl+K', action: D.linkDialog });
    if (eq) items.push({ label: 'Edit equation…', icon: 'sigma', action: () => D.equationDialog(eq) });
    if (fn) items.push({ label: 'Edit footnote…', icon: 'footprints', action: () => D.editFootnote(fn) });
    if (root === D.doc && !img) items.push({ label: 'New comment', icon: 'message-square-plus', action: D.addComment });
    if (cmt) items.push({ label: 'Delete comment', icon: 'message-square-x', action: () => D.deleteComment(cmt.dataset.cid) });
    if (!img) items.push({ sep: true }, { label: 'Paragraph…', icon: 'settings-2', action: paragraphDialog }, { label: 'Clear formatting', icon: 'remove-formatting', action: D.clearFormatting });
    BTS.menu({ x: e.clientX, y: e.clientY }, items);
  }

  /* ================= command router ================= */
  let prevTab = 'home';
  async function onCommand(cmd, el, ev) {
    const z = (k, ...a) => k(...a);
    switch (cmd) {
      case 'undo': return D.undo();
      case 'redo': return D.redo();
      // clipboard
      case 'paste': if (ev && ev.target.closest('.caret, span')) return BTS.menu(el, [{ label: 'Paste', icon: 'clipboard-paste', kbd: 'Ctrl+V', action: () => pasteFromClipboard(false) }, { label: 'Paste as plain text', icon: 'clipboard-type', kbd: 'Ctrl+Shift+V', action: () => pasteFromClipboard(true) }]); return pasteFromClipboard(false);
      case 'cut': D.ensureSel(); return document.execCommand('cut');
      case 'copy': D.ensureSel(); return document.execCommand('copy');
      case 'format-painter': return D.formatPainter(ev && ev.detail >= 2);
      // font
      case 'font-menu': return fontMenu($('#font-combo'));
      case 'size-menu': return sizeMenu($('#size-combo'));
      case 'grow': return D.growFont(1);
      case 'shrink': return D.growFont(-1);
      case 'case-menu': return caseMenu(el);
      case 'clear-format': return D.clearFormatting();
      case 'bold': return D.inline('bold');
      case 'italic': return D.inline('italic');
      case 'underline': return D.inline('underline');
      case 'strike': return D.inline('strikeThrough');
      case 'sub': return D.inline('subscript');
      case 'sup': return D.inline('superscript');
      case 'fore-color': return D.inline('foreColor', D.lastFore || '#c7354a');
      case 'fore-color-menu': { D.ensureSel(); const sv = D.saveSel(); const c = await BTS.colorPicker(el.previousElementSibling, { auto: 'Automatic' }); D.restoreSel(sv); if (!c) return; const v = c === 'auto' ? '#111111' : c; D.lastFore = v; $('#sw-fore').style.background = v; return D.inline('foreColor', v); }
      case 'hilite': return D.inline('hiliteColor', D.lastHilite || '#ffff00');
      case 'hilite-menu': { D.ensureSel(); const sv = D.saveSel(); const c = await BTS.colorPicker(el.previousElementSibling, { none: 'No colour' }); D.restoreSel(sv); if (!c) return; const v = c === 'none' ? 'transparent' : c; if (c !== 'none') { D.lastHilite = v; $('#sw-hilite').style.background = v; } return D.inline('hiliteColor', v); }
      // paragraph
      case 'bullets': return D.toggleList(false);
      case 'numbering': return D.toggleList(true);
      case 'bullets-menu': return bulletsMenu(el.previousElementSibling);
      case 'numbering-menu': return numberingMenu(el.previousElementSibling);
      case 'indent': return D.indent(1);
      case 'outdent': return D.indent(-1);
      case 'marks': case 'marks-check': return toggleMarks(cmd === 'marks-check' ? el.checked : undefined);
      case 'align-left': case 'align-center': case 'align-right': case 'align-justify': return D.align(cmd.slice(6));
      case 'line-spacing-menu': return lineSpacingMenu(el);
      case 'shading': { D.ensureSel(); const sv = D.saveSel(); const c = await BTS.colorPicker(el, { none: 'No colour' }); D.restoreSel(sv); if (c) D.paraShading(c); return; }
      case 'borders-menu': return bordersMenu(el);
      case 'paragraph-dialog': return paragraphDialog();
      // styles
      case 'style': hoverStyle = null; D.endPreview(); return D.setBlockStyle(el.dataset.style);
      case 'gal-up': $('#styles-gallery').scrollTop -= 62; return;
      case 'gal-down': $('#styles-gallery').scrollTop += 62; return;
      case 'thgal-up': $('#themes-gallery').scrollTop -= 62; return;
      case 'thgal-down': $('#themes-gallery').scrollTop += 62; return;
      case 'styles-menu': return stylesMenu(el);
      // editing
      case 'find': return D.openFind(false);
      case 'replace': return D.openFind(true);
      case 'select-all': return selectAll();
      case 'find-close': return D.closeFind();
      case 'find-next': return D.findStep(1);
      case 'find-prev': return D.findStep(-1);
      case 'replace-one': return D.replaceOne();
      case 'replace-all': return D.replaceAll();
      case 'find-toggle-replace': return D.openFind(!$('#find-pane').classList.contains('replace-mode'), $('#find-input').value);
      // insert
      case 'blank-page': return D.blankPage();
      case 'page-break': return D.pageBreak();
      case 'table-menu': return tableMenu(el);
      case 'picture': return D.pickPicture();
      case 'shapes-menu': return D.shapesMenu(el);
      case 'link': return D.linkDialog();
      case 'comment-new': return D.addComment();
      case 'header-menu': return headerMenu(el, 'header');
      case 'footer-menu': return headerMenu(el, 'footer');
      case 'pagenum-menu': return pageNumMenu(el);
      case 'textbox': return D.insertTextBox();
      case 'datetime': return D.dateTimeDialog();
      case 'hr': return D.insertHR();
      case 'dropcap': return D.dropCap();
      case 'equation': return D.equationDialog();
      case 'symbol': return D.symbolDialog();
      // design
      case 'theme': return D.setDesign({ theme: el.dataset.th, headFont: null, bodyFont: null });
      case 'spacing-menu': return spacingMenu(el);
      case 'theme-fonts-menu': return themeFontsMenu(el);
      case 'watermark-menu': return watermarkMenu(el);
      case 'page-color': { const c = await BTS.colorPicker(el, { none: 'No colour', value: D.design.pageColor || '#ffffff' }); if (c) D.setDesign({ pageColor: c === 'none' ? null : c }); return; }
      case 'page-border-menu': return pageBorderMenu(el);
      // layout
      case 'margins-menu': return marginsMenu(el);
      case 'orientation-menu': return orientationMenu(el);
      case 'size-page-menu': return sizeMenuPage(el);
      case 'columns-menu': return columnsMenu(el);
      case 'breaks-menu': return breaksMenu(el);
      case 'indent-left': return D.indentPara('left', Math.max(0, +el.value || 0));
      case 'indent-right': return D.indentPara('right', Math.max(0, +el.value || 0));
      case 'space-before': return D.paraSpacing('before', Math.max(0, +el.value || 0));
      case 'space-after': return D.paraSpacing('after', Math.max(0, +el.value || 0));
      case 'line-spacing-select': return D.lineSpacing(+el.value);
      // references
      case 'toc-menu': return tocMenu(el);
      case 'toc-update': return D.updateTOC();
      case 'addtext-menu': return addTextMenu(el);
      case 'footnote': return D.insertFootnote();
      case 'footnote-next': return D.nextFootnote();
      case 'caption': return D.caption();
      // review
      case 'spelling': return toggleSpelling();
      case 'word-count': return D.wordCountDialog();
      case 'read-aloud': return D.readAloud();
      case 'comment-delete': return D.deleteComment();
      case 'comment-delete-id': return D.deleteComment(el.dataset.cid);
      case 'comment-delete-all': return D.deleteAllComments();
      case 'comment-prev': return D.stepComment(-1);
      case 'comment-next': return D.stepComment(1);
      case 'comments-toggle': D._commentsHiddenByUser = !$('#crail').classList.contains('hidden'); D.showComments(); return D.scheduleLayout(20);
      // view
      case 'print-layout': return toggleFocus(false);
      case 'focus-mode': return toggleFocus();
      case 'dark-canvas': return toggleDarkCanvas();
      case 'ruler': document.body.classList.toggle('no-ruler', !el.checked); return D.scheduleLayout(10);
      case 'nav-pane': return D.toggleNav(el.checked);
      case 'nav-close': return D.toggleNav(false);
      case 'zoom-menu': return zoomMenu(el);
      case 'zoom-page': return D.fitZoom('page');
      case 'zoom-width': return D.fitZoom('width');
      case 'goto-page': return gotoPage();
      // table tools
      case 'tbl-delete-menu': return tblDeleteMenu(el);
      case 'tbl-row-above': return D.tableOp('row-above');
      case 'tbl-row-below': return D.tableOp('row-below');
      case 'tbl-col-left': return D.tableOp('col-left');
      case 'tbl-col-right': return D.tableOp('col-right');
      case 'tbl-merge': return D.tableOp('merge');
      case 'tbl-split': return D.tableOp('split');
      case 'tbl-style': return D.tableOp('style', el.dataset.style);
      case 'tbl-header-row': return D.tableOp('header-row', el.checked);
      case 'tbl-shading': { const sv = D.saveSel(); const c = await BTS.colorPicker(el, { none: 'No colour' }); D.restoreSel(sv); if (c) D.tableOp('shading', c); return; }
      case 'tbl-borders-menu': return tblBordersMenu(el);
      case 'cell-align-left': case 'cell-align-center': case 'cell-align-right': return D.tableOp('align', cmd.slice(11));
      case 'cell-valign-top': case 'cell-valign-middle': case 'cell-valign-bottom': return D.tableOp('valign', cmd.slice(12));
      case 'tbl-autofit-content': return D.tableOp('autofit-content');
      case 'tbl-autofit-window': return D.tableOp('autofit-window');
      case 'tbl-distribute': return D.tableOp('distribute');
      // picture tools
      case 'img-wrap-inline': case 'img-wrap-left': case 'img-wrap-right': case 'img-wrap-center': return D.imgOp('wrap', el.dataset.wrap);
      case 'img-width': return D.imgOp('w', +el.value);
      case 'img-height': return D.imgOp('h', +el.value);
      case 'img-reset': return D.imgOp('reset');
      case 'img-border': return D.imgOp('toggle', 'img-border');
      case 'img-rounded': return D.imgOp('toggle', 'img-rounded');
      case 'img-shadow': return D.imgOp('toggle', 'img-shadow');
      case 'img-alt': return D.imgAlt();
      case 'img-replace': return D.pickPicture(true);
      case 'img-delete': return D.imgOp('delete');
    }
  }

  /* ================= doc-level mouse handlers ================= */
  function wireDocMouse() {
    D.doc.addEventListener('click', e => {
      const a = e.target.closest('a'); if (a && (e.ctrlKey || e.metaKey)) { e.preventDefault(); D.followLink(a); return; }
      const tp = e.target.closest('.toc p[data-target]'); if (tp) { const t = document.getElementById(tp.dataset.target); if (t) D.scrollToEl(t, true); return; }
      const fn = e.target.closest('.fn-ref'); if (fn) { D.editFootnote(fn); return; }
    });
    D.doc.addEventListener('dblclick', e => { const eq = e.target.closest('.eq'); if (eq) D.equationDialog(eq); });
    $('#fn-area').addEventListener('click', e => { const li = e.target.closest('li[data-fn]'); if (li) { const s = $(`sup.fn-ref[data-fn="${li.dataset.fn}"]`, D.doc); if (s) { D.scrollToEl(s); D.editFootnote(s); } } });
    D.doc.addEventListener('mouseup', () => { setTimeout(() => D.applyPainter(), 0); });
    D.doc.addEventListener('mouseover', e => { const a = e.target.closest('a'); if (a && !a.title) a.title = (a.getAttribute('href') || '') + '\nCtrl+Click to follow link'; });
    // clicking the empty grey canvas below the text puts the caret at the end
    D.canvas.addEventListener('mousedown', e => {
      if (e.target === D.canvas || e.target.id === 'zoomer' || e.target.id === 'sheet-row' || e.target.id === 'body-wrap' || e.target.id === 'sheet') {
        if (e.target.id === 'body-wrap' || e.target.id === 'sheet') {
          const dr = D.doc.getBoundingClientRect(); if (e.clientY > dr.bottom) { e.preventDefault(); const last = D.doc.lastElementChild; if (last && /^(P|H\d|LI)$/.test(last.tagName)) { const r = document.createRange(); r.selectNodeContents(last); r.collapse(false); D.doc.focus(); const s = getSelection(); s.removeAllRanges(); s.addRange(r); } }
        }
      }
    });
    [D.doc, D.hdr, D.ftr].forEach(el => { el.addEventListener('paste', onPaste); el.addEventListener('copy', onCopy); el.addEventListener('cut', onCopy); el.addEventListener('keydown', docKey); el.addEventListener('contextmenu', contextMenu); });
    document.addEventListener('keydown', globalKey);
    // font name / size inputs
    const fn = $('#font-name'), fs = $('#font-size');
    fn.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); D.setFontName(fn.value.trim()); D.restoreFocus(); } if (e.key === 'Escape') D.restoreFocus(); });
    fs.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); const v = parseFloat(fs.value); if (v > 0 && v <= 1638) D.setFontSize(v); else BTS.toast('Font size must be between 1 and 1638.', { error: true }); D.restoreFocus(); } if (e.key === 'Escape') D.restoreFocus(); });
    [fn, fs].forEach(i => i.addEventListener('focus', () => i.select()));
    // number inputs in the ribbon: apply on Enter without losing the selection
    $$('#ribbon input.num').forEach(i => i.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); i.dispatchEvent(new Event('change', { bubbles: true })); D.restoreFocus(); } }));
    window.addEventListener('resize', () => D.scheduleLayout(80));
    window.addEventListener('beforeunload', () => { if (!BTS.dirty) D.clearAutosave(); });
  }
  D.restoreFocus = () => { const r = D.lastRange(); const root = r && D.rootOf(r.commonAncestorContainer); if (root) { root.focus({ preventScroll: true }); const s = getSelection(); s.removeAllRanges(); s.addRange(r); } else D.doc.focus(); };

  D.onChanged = () => { D.autosave(); clearTimeout(D._navT); D._navT = setTimeout(() => { D.renderNav(); D.refreshFind(); if ($$('sup.fn-ref', D.doc).length !== $$('#fn-area li').length) D.renderFootnotes(); D.renderComments(); }, 300); };

  /* ================= serialize & init ================= */
  D.hasContent = () => !!(D.doc.textContent.replace(/​/g, '').trim() || D.doc.querySelector('img,table,hr,.toc') || D.hdr.textContent.trim() || D.ftr.textContent.trim());
  function newDocument() {
    D.page = D.defaultPage(); D.design = D.defaultDesign(); D.comments = {};
    D.setContent('<p><br></p>'); D.renderComments(); BTS.setTitle('Document'); BTS.filePath = null; BTS.markClean(); D.setCaret(D.doc.firstChild, 0);
  }
  function infoHTML() {
    const s = D.stats(); const ps = D.PAPER[D.page.size];
    const row = (k, v) => `<div style="display:contents"><span style="color:var(--ink-3)">${k}</span><span>${v}</span></div>`;
    return `<div style="display:grid;grid-template-columns:180px 1fr;gap:10px 18px;max-width:640px;font-size:13.5px">
      ${row('Title', esc(BTS.title))}${row('Location', esc(BTS.filePath || 'Not saved yet'))}
      ${row('Pages', s.pages)}${row('Words', s.words.toLocaleString('en-IN'))}${row('Characters', s.chars.toLocaleString('en-IN'))}${row('Paragraphs', s.paras)}
      ${row('Paper', `${ps.label} · ${D.page.orient}`)}${row('Margins', `T ${(D.page.m.t * 2.54).toFixed(2)} · B ${(D.page.m.b * 2.54).toFixed(2)} · L ${(D.page.m.l * 2.54).toFixed(2)} · R ${(D.page.m.r * 2.54).toFixed(2)} cm`)}
      ${row('Comments', Object.keys(D.comments).length)}${row('Pictures', $$('img', D.doc).length)}${row('Author', esc(D.author()))}
    </div><h2>Author name</h2><div class="field" style="max-width:320px"><input id="info-author" value="${esc(D.author())}" onchange="try{localStorage.setItem('docs-author',this.value)}catch(e){}"></div>`;
  }

  function init() {
    D.initEditor(); D.initTables(); D.initInsert(); D.initPanes();
    renderStyles(); renderThemes();
    wireDocMouse();
    const cfg = {
      app: 'docs', name: 'BTS Docs', letter: 'D', defaultTitle: 'Document',
      saveFormats: [
        { ext: 'docx', label: 'Word document (.docx)', icon: 'file-text', hint: 'Opens in Microsoft Word, LibreOffice and Google Docs' },
        { ext: 'html', label: 'Web page (.html)', icon: 'globe', hint: 'Keeps all BTS Docs formatting; opens in any browser' },
        { ext: 'md', label: 'Markdown (.md)', icon: 'file-code', hint: 'Plain text with headings, lists, links and tables' },
        { ext: 'txt', label: 'Plain text (.txt)', icon: 'file-type', hint: 'Text only, no formatting' },
      ],
      openAccept: '.docx,.odt,.txt,.md,.markdown,.html,.htm,.rtf',
      onCommand,
      load: (f) => D.loadFile(f),
      async save(ext) {
        let out;
        if (ext === 'docx') out = await D.exportDocx();
        else if (ext === 'html') out = D.serializeNative();
        else if (ext === 'md') out = D.toMarkdown();
        else out = D.toText();
        return out;
      },
      newDocument,
      exportPDF: async () => D.buildPrint(),
      afterExportPDF: () => D.afterPrint(),
      onZoom: (z) => D.setZoom(z), zoom: 100, minZoom: 20, maxZoom: 300,
      templates: D.TEMPLATES.map(t => ({ name: t.name, thumbHTML: t.thumbHTML, apply: () => { t.apply(); D.renderComments(); D.canvas.scrollTop = 0; $('#recover-bar').classList.add('hidden'); } })),
      blankThumb: D.blankThumb,
      search: (text) => D.openFind(false, text),
      hasContent: D.hasContent,
      infoHTML,
      focus: () => D.restoreFocus(),
      onDropFile(f, e) {
        if (!f.type.startsWith('image/')) return false;
        const pos = document.caretRangeFromPoint && document.caretRangeFromPoint(e.clientX, e.clientY);
        if (pos && D.rootOf(pos.startContainer)) { const s = getSelection(); s.removeAllRanges(); s.addRange(pos); }
        f.arrayBuffer().then(async b => D.insertImageURL(await BTS.bytesToDataURL(new Uint8Array(b), f.type), f.name.replace(/\.[^.]+$/, '')));
        return true;
      },
      shortcuts: [['Ctrl+B / I / U', 'Bold / italic / underline'], ['Ctrl+E / L / R / J', 'Center / left / right / justify'], ['Ctrl+K', 'Insert link'], ['Ctrl+Enter', 'Page break'], ['Ctrl+Shift+> / <', 'Bigger / smaller font'], ['Ctrl+1 / 2 / 5', 'Line spacing 1 / 2 / 1.5'], ['Ctrl+Alt+1…3', 'Heading 1–3'], ['Ctrl+Shift+N', 'Normal style'], ['Ctrl+Shift+L', 'Bullets'], ['Ctrl+Shift+V', 'Paste as plain text'], ['Ctrl+G', 'Go to page'], ['Tab / Shift+Tab', 'Indent list / next table cell'], ['F7', 'Spelling on/off']],
      ready() {
        D.updateStatus();
        if (!BTS.loadedExt) D.checkRecovery();
        D.setCaret(D.doc.firstElementChild || D.doc, 0);
      },
    };
    BTS.hasContent = D.hasContent; // the shell reads BTS.hasContent (not cfg.hasContent)
    BTS.init(cfg);
    // a successful save (or load) makes the recovery copy unnecessary
    const mc = BTS.markClean; BTS.markClean = () => { mc(); D.clearAutosave(); };
    D.layout();
  }
  init();
})();
