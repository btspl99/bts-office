/* BTS Slides 1.2 — styles: font box (BTS.fontPicker), Design ▸ Fonts / Colors (+ customise dialogs),
 * shape styles gallery (subtle / moderate / intense × theme colours) and table styles gallery.
 * New ribbon commands register themselves in window.SLX.cmd (routed by app.js before the built-in switch).
 */
(function () {
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const esc = (s) => BTS.esc(s);
  const P = () => SL.pres;
  const X = window.SLX = window.SLX || { cmd: {}, ribbon: [] };
  const C = () => window.CMD;

  /* ---------------- font box ---------------- */
  // Keep the text selection while the picker (which has its own search box) is open.
  function saveRange() { const s = getSelection(); return E.editing && s.rangeCount ? s.getRangeAt(0).cloneRange() : null; }
  function restoreRange(r) {
    if (!r || !E.editing) return;
    let n = r.startContainer; if (n.nodeType === 3) n = n.parentElement;
    const ed = n && n.closest && n.closest('[contenteditable="true"]'); if (!ed || !ed.isConnected) return;
    ed.focus({ preventScroll: true }); const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  }
  X.themeFontChoices = () => { const f = SL.themeFonts(P()); return [{ label: 'Headings', name: f.title }, { label: 'Body', name: f.body }]; };
  X.cmd['font-picker'] = async (el) => {
    const r = saveRange();
    const cur = $('#font-name-label') ? $('#font-name-label').textContent : '';
    const name = await BTS.fontPicker(el, { current: cur, theme: X.themeFontChoices() });
    if (!name) { restoreRange(r); return; }
    await SL.ensureFonts([name]);
    restoreRange(r);
    if (!E.editing && !C().textTargets().length) return BTS.toast('Select a text box first');
    C().setFont(name);
  };

  /* ---------------- Design ▸ Fonts ---------------- */
  const customPairs = () => { try { return JSON.parse(localStorage.getItem('bts-slides-font-pairs') || '[]'); } catch { return []; } };
  const customSchemes = () => { try { return JSON.parse(localStorage.getItem('bts-slides-color-schemes') || '[]'); } catch { return []; } };
  function setThemeFonts(f) {
    if (E.editing) E.endEdit();
    P().theme = Object.assign({}, P().theme); if (f) P().theme.fonts = { title: f.title, body: f.body }; else delete P().theme.fonts;
    SL.commit(); C().buildThemeGallery(); SL.ensureFonts(f ? [f.title, f.body] : SL.usedFonts(P())).then(() => E.renderAll()); E.renderAll();
  }
  X.setThemeFonts = setThemeFonts;
  X.cmd['theme-fonts'] = async (el) => {
    const all = [...customPairs().map(p => Object.assign({ custom: true }, p)), ...SL.FONT_PAIRS];
    await SL.ensureFonts([...new Set(all.flatMap(p => [p.title, p.body]))]);
    const cur = SL.themeFonts(P()); const th = SL.themeOf(P());
    const row = (p, i) => `<button class="fp-pair ${p.title === cur.title && p.body === cur.body ? 'on' : ''}" data-i="${i}" title="${esc(p.title)} headings, ${esc(p.body)} body text"><span class="fpp-aa" style="font-family:${SL.fontStack(p.title)}">Aa</span><span class="fpp-n"><b style="font-family:${SL.fontStack(p.title)}">${esc(p.title)}</b><small style="font-family:${SL.fontStack(p.body)}">${esc(p.body)}</small></span><span class="kbd">${esc(p.name)}</span></button>`;
    C().pop(el, `<div class="menu-head">Theme fonts (${esc(th.name)}: ${esc(th.titleFont)} / ${esc(th.bodyFont)})</div><div class="fp-pairs">${all.map(row).join('')}</div><div class="menu-sep"></div><div class="menu-item" data-act="custom">${icon('type', 16)}<span>Customize fonts…</span></div><div class="menu-item ${P().theme.fonts ? '' : 'disabled'}" data-act="reset">${icon('rotate-ccw', 16)}<span>Reset to the theme’s fonts</span></div>`, 'pairs-pop', (e) => {
      const b = e.target.closest('[data-i]'); if (b) { BTS.closePopover(); setThemeFonts(all[+b.dataset.i]); return; }
      const a = e.target.closest('[data-act]'); if (!a) return; BTS.closePopover();
      if (a.dataset.act === 'custom') customizeFonts(); else if (a.dataset.act === 'reset') setThemeFonts(null);
    });
  };
  async function customizeFonts() {
    const cur = SL.themeFonts(P()); const st = { title: cur.title, body: cur.body };
    const r = await BTS.dialog({
      title: 'Create new theme fonts', width: 520,
      body: `<div class="field"><label>Heading font</label><button class="btn font-pick" data-k="title" title="Choose the heading font"></button></div><div class="field"><label>Body font</label><button class="btn font-pick" data-k="body" title="Choose the body font"></button></div>
        <div class="fp-sample"><div class="fps-h"></div><div class="fps-b"></div></div><div class="field"><label>Name</label><input id="cf-name" value="Custom fonts"></div>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'Save', primary: true, value: (root) => ({ name: $('#cf-name', root).value || 'Custom fonts', title: st.title, body: st.body }) }],
      onOpen: (root) => {
        const draw = () => { $$('.font-pick', root).forEach(b => { b.textContent = st[b.dataset.k]; b.style.fontFamily = SL.fontStack(st[b.dataset.k]); }); const h = $('.fps-h', root), bd = $('.fps-b', root); h.textContent = 'Heading'; h.style.fontFamily = SL.fontStack(st.title); bd.textContent = 'Body text: the quick brown fox jumps over the lazy dog. हिन्दी में भी लिखें।'; bd.style.fontFamily = SL.fontStack(st.body); };
        draw();
        root.addEventListener('click', async e => { const b = e.target.closest('.font-pick'); if (!b) return; const n = await BTS.fontPicker(b, { current: st[b.dataset.k] }); if (n) { await SL.ensureFonts([n]); st[b.dataset.k] = n; draw(); } });
      },
    });
    if (!r) return;
    const list = customPairs().filter(p => p.name !== r.name); list.unshift(r); try { localStorage.setItem('bts-slides-font-pairs', JSON.stringify(list.slice(0, 12))); } catch { }
    setThemeFonts(r);
  }

  /* ---------------- Design ▸ Colors ---------------- */
  function setThemeColors(c) {
    if (E.editing) E.endEdit();
    P().theme = Object.assign({}, P().theme); if (c) P().theme.colors = SL.clone(c); else delete P().theme.colors;
    SL.commit(); C().buildThemeGallery(); E.renderAll();
  }
  X.setThemeColors = setThemeColors;
  X.cmd['theme-colors'] = (el) => {
    const base = SL.palette(Object.assign({}, P(), { theme: { id: P().theme.id, variant: P().theme.variant, custom: P().theme.custom } }));
    const all = [...customSchemes().map(s => Object.assign({ custom: true }, s)), ...SL.COLOR_SCHEMES];
    const cur = SL.palette(P());
    const strip = (s) => [s.text || base.text, s.bg || base.bg, ...s.accents].map(c => `<i style="background:${c}"></i>`).join('');
    C().pop(el, `<div class="menu-head">Theme colors</div><div class="cs-list"><button class="cs-row ${!P().theme.colors ? 'on' : ''}" data-i="-1" title="The colors of the ${esc(SL.themeOf(P()).name)} theme"><span class="cs-strip">${strip(base)}</span><span>${esc(SL.themeOf(P()).name)} (theme)</span></button>${all.map((s, i) => `<button class="cs-row ${P().theme.colors && JSON.stringify(cur.accents) === JSON.stringify(s.accents) ? 'on' : ''}" data-i="${i}" title="${esc(s.name)}"><span class="cs-strip">${strip(s)}</span><span>${esc(s.name)}${s.custom ? ' <small>(custom)</small>' : ''}</span></button>`).join('')}</div><div class="menu-sep"></div><div class="menu-item" data-act="custom">${icon('palette', 16)}<span>Customize colors…</span></div><div class="menu-item ${P().theme.colors ? '' : 'disabled'}" data-act="reset">${icon('rotate-ccw', 16)}<span>Reset slide theme colors</span></div>`, 'cs-pop', (e) => {
      const b = e.target.closest('[data-i]'); if (b) { BTS.closePopover(); const i = +b.dataset.i; if (i < 0) return setThemeColors(null); const s = all[i]; setThemeColors(Object.assign({ name: s.name, accents: s.accents.slice() }, s.bg ? { bg: s.bg, bg2: s.bg2, text: s.text, title: s.title } : {})); return; }
      const a = e.target.closest('[data-act]'); if (!a) return; BTS.closePopover();
      if (a.dataset.act === 'custom') X.customizeColors(); else if (a.dataset.act === 'reset') setThemeColors(null);
    });
  };
  X.customizeColors = async () => {
    const pal = SL.palette(P());
    const st = { name: (P().theme.colors && P().theme.colors.name) || 'Custom colors', bg: pal.bg, bg2: pal.bg2, text: pal.text, title: pal.title, accents: pal.accents.slice() };
    const rows = [['text', 'Text / background – dark'], ['bg', 'Text / background – light'], ['title', 'Headings'], ['bg2', 'Background 2'], ...[0, 1, 2, 3, 4, 5].map(i => ['a' + i, 'Accent ' + (i + 1)])];
    const get = (k) => k[0] === 'a' && k.length === 2 ? st.accents[+k[1]] : st[k];
    const set = (k, v) => { if (k[0] === 'a' && k.length === 2) st.accents[+k[1]] = v; else st[k] = v; };
    const r = await BTS.dialog({
      title: 'Create new theme colors', width: 620,
      body: `<div class="cc-wrap"><div class="cc-rows">${rows.map(([k, n]) => `<div class="cc-row"><label>${n}</label><button class="pf-swatch" data-k="${k}" title="Choose ${n.toLowerCase()}"><i></i>${icon('chevron-down', 12)}</button><input class="cc-hex" data-k="${k}" maxlength="7" spellcheck="false" title="Hex color, e.g. #1F5AA6"></div>`).join('')}</div><div class="cc-prev"><div class="cc-thumb"></div><div class="cc-thumb2"></div></div></div><div class="field"><label>Name</label><input id="cc-name" value="${esc(st.name)}"></div>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'Save', primary: true, value: (root) => { st.name = $('#cc-name', root).value || 'Custom colors'; return st; } }],
      onOpen: (root) => {
        const draw = () => {
          $$('.pf-swatch', root).forEach(b => b.querySelector('i').style.background = get(b.dataset.k));
          $$('.cc-hex', root).forEach(i => { if (document.activeElement !== i) i.value = get(i.dataset.k); });
          const p = { width: SL.W(P()), height: SL.H(P()), theme: Object.assign({}, P().theme, { colors: SL.clone(st) }), master: P().master, slides: [] };
          const s1 = SL.newSlide(p, 'title'); s1.elements[0].html = '<p>Title</p>'; s1.elements[1].html = '<p>Subtitle text</p>';
          const s2 = SL.newSlide(p, 'titleContent'); s2.elements[0].html = '<p>Heading</p>'; s2.elements[1].html = '<ul><li>Body text</li><li>More text</li></ul>';
          s2.elements.push(SL.newElement('chart', { x: 700, y: 330, w: 480, h: 320, chart: { type: 'bar', title: '', legend: false, categories: ['A', 'B', 'C'], series: [0, 1, 2, 3, 4, 5].map(i => ({ name: 'S' + i, values: [3 + i, 5 + i % 3, 4 + (i % 2)] })) } }));
          $('.cc-thumb', root).innerHTML = SL.thumbHTML(s1, p, 220, 0); $('.cc-thumb2', root).innerHTML = SL.thumbHTML(s2, p, 220, 1);
        };
        draw();
        root.addEventListener('click', async e => { const b = e.target.closest('.pf-swatch'); if (!b) return; const c = await BTS.colorPicker(b, {}); if (c && c[0] === '#') { set(b.dataset.k, c); draw(); } });
        root.addEventListener('input', e => { const i = e.target.closest('.cc-hex'); if (!i) return; if (/^#[0-9a-f]{6}$/i.test(i.value)) { set(i.dataset.k, i.value.toUpperCase()); draw(); } });
      },
    });
    if (!r) return;
    const list = customSchemes().filter(s => s.name !== r.name); list.unshift(SL.clone(r)); try { localStorage.setItem('bts-slides-color-schemes', JSON.stringify(list.slice(0, 12))); } catch { }
    setThemeColors(r);
  };

  /* ---------------- Shape styles ---------------- */
  X.SHAPE_STYLE_ROWS = ['Colored outline', 'Colored fill', 'Light outline, colored fill', 'Subtle effect', 'Moderate effect', 'Intense effect'];
  X.shapeStyles = () => {
    const cols = ['text', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
    const out = [];
    X.SHAPE_STYLE_ROWS.forEach((label, ri) => cols.forEach((a, ci) => {
      const n = ci ? 'Accent ' + ci : 'Dark';
      let s;
      switch (ri) {
        case 0: s = { fill: 'bg', stroke: a, strokeWidth: 1.5, color: a, shadow: false }; break;
        case 1: s = { fill: a, stroke: a + '/d30', strokeWidth: 1, color: '#ffffff', shadow: false }; break;
        case 2: s = { fill: a, stroke: '#ffffff', strokeWidth: 1.5, color: '#ffffff', shadow: false }; break;
        case 3: s = { fill: a + '/l80', stroke: a, strokeWidth: 1, color: a + '/d40', shadow: false }; break;
        case 4: s = { fill: { type: 'gradient', color: a + '/l25', color2: a, angle: 90 }, stroke: 'none', strokeWidth: 0, color: '#ffffff', shadow: true }; break;
        default: s = { fill: { type: 'gradient', color: a, color2: a + '/d35', angle: 90 }, stroke: a + '/d20', strokeWidth: 1, color: '#ffffff', shadow: true, bold: true };
      }
      s.label = `${label}, ${n}`; s.row = ri; out.push(s);
    }));
    return out;
  };
  X.applyShapeStyle = (st) => {
    const els = E.selected().filter(e => e.type === 'shape' || e.type === 'text'); if (!els.length) return BTS.toast('Select a shape or text box first');
    if (E.editing) E.syncEditing();
    els.forEach(e => {
      if (SL.isLine(e)) { e.stroke = typeof st.fill === 'string' && st.fill !== 'bg' ? st.fill : (typeof st.fill === 'object' ? st.fill.color2 : st.stroke); e.strokeWidth = Math.max(2, st.strokeWidth || 2); e.shadow = !!st.shadow; return; }
      e.fill = SL.clone(st.fill); e.stroke = st.stroke; e.strokeWidth = st.strokeWidth || 0; e.shadow = !!st.shadow; e.color = st.color;
      if (st.bold) e.bold = true;
    });
    E.renderCanvas(); SL.commit(); E.refreshThumb(true); E.updateRibbon();
  };
  const styleSwatch = (s, w = 34, h = 26) => { const demo = { id: 'ss' + Math.random().toString(36).slice(2, 7), type: 'shape', kind: 'roundRect', adj: .12, fill: s.fill, stroke: s.stroke, strokeWidth: s.strokeWidth }; return `${SL.shapeSVG(demo, P(), w, h)}<span style="color:${SL.color(s.color, P())}">Abc</span>`; };
  X.cmd['shape-styles'] = (el) => {
    const ss = X.shapeStyles();
    C().pop(el, `<div class="menu-head">Theme styles</div><div class="ss-grid">${ss.map((s, i) => `<button data-i="${i}" class="${s.shadow ? 'sh' : ''}" title="${esc(s.label)}">${styleSwatch(s)}</button>`).join('')}</div><div class="menu-head">Presets</div><div class="ss-grid">${C().QUICK(P()).filter((q, i) => i % 5 === 1 || i % 5 === 2).slice(0, 7).map((q, i) => `<button data-q="${i}" title="${esc(q.label)}">${styleSwatch(q)}</button>`).join('')}</div>`, 'ss-pop', (e) => {
      const b = e.target.closest('[data-i]'); if (b) { BTS.closePopover(); X.applyShapeStyle(ss[+b.dataset.i]); return; }
      const q = e.target.closest('[data-q]'); if (q) { BTS.closePopover(); C().applyQuick(C().QUICK(P()).filter((x, i) => i % 5 === 1 || i % 5 === 2)[+q.dataset.q]); }
    });
  };
  X.buildStyleGallery = () => {
    const sg = $('#style-gal'); if (!sg) return;
    const ss = X.shapeStyles(); const pick = [1, 2, 3, 4, 5, 6, 22, 23, 24, 36, 37, 38];
    sg.innerHTML = pick.map(i => { const s = ss[i]; return `<button data-ss="${i}" class="${s.shadow ? 'sh' : ''}" title="${esc(s.label)}">${styleSwatch(s, 26, 26).replace('Abc', 'Ab')}</button>`; }).join('');
    sg.onclick = (e) => { const b = e.target.closest('[data-ss]'); if (b) X.applyShapeStyle(X.shapeStyles()[+b.dataset.ss]); };
    sg.onmousedown = (e) => e.preventDefault();
  };

  /* ---------------- Table styles ---------------- */
  X.tableStyles = () => { const out = []; SL.TABLE_STYLES.forEach(([id, n]) => [1, 2, 3, 4, 5, 6].forEach(a => out.push({ tstyle: id, accent: 'accent' + a, label: `${n} style, Accent ${a}` }))); return out; };
  const tablePreview = (st, w = 64, h = 44) => {
    const t = { type: 'table', rows: Array.from({ length: 4 }, () => Array.from({ length: 3 }, () => ({ html: '' }))), header: true, banded: true, tstyle: st.tstyle, accent: st.accent, colW: [1 / 3, 1 / 3, 1 / 3], w, h, size: 4 };
    let s = `<svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">`; const cw = w / 3, rh = h / 4;
    for (let r = 0; r < 4; r++) for (let c = 0; c < 3; c++) {
      const ts = SL.tableCellStyle(t, r, c, P()); const x = c * cw, y = r * rh;
      if (ts.bg) s += `<rect x="${x}" y="${y}" width="${cw}" height="${rh}" fill="${ts.bg}"/>`;
      if (ts.b[0]) s += `<line x1="${x}" y1="${y}" x2="${x + cw}" y2="${y}" stroke="${ts.b[0].c}" stroke-width="${Math.min(1.5, ts.b[0].w)}"/>`;
      if (ts.b[2]) s += `<line x1="${x}" y1="${y + rh}" x2="${x + cw}" y2="${y + rh}" stroke="${ts.b[2].c}" stroke-width="${Math.min(1.5, ts.b[2].w)}"/>`;
      if (ts.b[3]) s += `<line x1="${x}" y1="${y}" x2="${x}" y2="${y + rh}" stroke="${ts.b[3].c}" stroke-width="${Math.min(1.5, ts.b[3].w)}"/>`;
      if (ts.b[1] && c === 2) s += `<line x1="${x + cw}" y1="${y}" x2="${x + cw}" y2="${y + rh}" stroke="${ts.b[1].c}" stroke-width="${Math.min(1.5, ts.b[1].w)}"/>`;
      s += `<rect x="${x + 3}" y="${y + rh / 2 - 1}" width="${cw * .5}" height="2" fill="${ts.color}" opacity=".7"/>`;
    }
    return s + '</svg>';
  };
  X.applyTableStyle = (st) => {
    const t = E.primary(); if (!t || t.type !== 'table') return BTS.toast('Select a table first');
    E.syncEditing(); t.tstyle = st.tstyle; t.accent = st.accent; t.rows.forEach(r => r.forEach(c => { delete c.fill; delete c.color; }));
    SL.commit(); E.renderCanvas(); E.refreshThumb(true); E.updateRibbon();
  };
  X.cmd['table-styles'] = (el) => {
    const ts = X.tableStyles(); const t = E.primary();
    C().pop(el, SL.TABLE_STYLES.map(([id, n]) => `<div class="menu-head">${esc(n)}</div><div class="ts-grid">${ts.map((s, i) => s.tstyle !== id ? '' : `<button data-i="${i}" class="${t && (t.tstyle || 'medium') === s.tstyle && (t.accent || 'accent1') === s.accent ? 'on' : ''}" title="${esc(s.label)}">${tablePreview(s)}</button>`).join('')}</div>`).join(''), 'ts-pop', (e) => { const b = e.target.closest('[data-i]'); if (b) { BTS.closePopover(); X.applyTableStyle(ts[+b.dataset.i]); } });
  };
  X.buildTableGallery = () => {
    const g = $('#tstyle-gal'); if (!g) return; const ts = X.tableStyles();
    const pick = [0, 1, 6, 7, 12, 18]; const t = E.primary();
    g.innerHTML = pick.map(i => `<button data-ts="${i}" class="${t && t.type === 'table' && (t.tstyle || 'medium') === ts[i].tstyle && (t.accent || 'accent1') === ts[i].accent ? 'on' : ''}" title="${esc(ts[i].label)}">${tablePreview(ts[i], 46, 32)}</button>`).join('');
    g.onclick = (e) => { const b = e.target.closest('[data-ts]'); if (b) X.applyTableStyle(X.tableStyles()[+b.dataset.ts]); };
    g.onmousedown = (e) => e.preventDefault();
  };
  X.cmd['tbl-firstcol'] = (el) => { const t = E.primary(); if (!t || t.type !== 'table') return; E.syncEditing(); t.firstCol = el.checked; SL.commit(); E.renderCanvas(); E.refreshThumb(true); };
  X.cmd['tbl-lastrow'] = (el) => { const t = E.primary(); if (!t || t.type !== 'table') return; E.syncEditing(); t.lastRow = el.checked; SL.commit(); E.renderCanvas(); E.refreshThumb(true); };

  /* ---------------- ribbon state ---------------- */
  X.ribbon.push(() => {
    const els = E.selected(); const t = C().textTargets()[0];
    const lab = $('#font-name-label');
    if (lab) { let f = ''; if (E.editing) { const s = getSelection(); let n = s.anchorNode; if (n && n.nodeType === 3) n = n.parentElement; if (n) f = getComputedStyle(n).fontFamily.split(',')[0].replace(/["']/g, '').trim(); } else if (t) f = SL.textDefaults(t, P()).font; lab.textContent = f; lab.style.fontFamily = f ? SL.fontStack(f) : ''; }
    const tb = els[0] && els[0].type === 'table' ? els[0] : null;
    if (tb) { const a = $('#tbl-firstcol'), b = $('#tbl-lastrow'); if (a) a.checked = !!tb.firstCol; if (b) b.checked = !!tb.lastRow; X.buildTableGallery(); }
  });
})();
