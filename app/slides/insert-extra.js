/* BTS Slides 1.3 — Insert ▸ SmartArt and Insert ▸ Equation.
 * SmartArt: the shared BTSSmartArt module (common/smartart.js) lays the diagram out; it is inserted as ordinary,
 * individually editable shapes that share one group id (el.g). Every member keeps the source as
 * el.smart = { layout, items, box, bb } so "Layout & Text" can lay it out again in the same place.
 * Equations: el = { type:'equation', eq (linear format, same as BTS Docs), size (pt), color }; exported to .pptx/.odp
 * as a sharp PNG with the equation text as alt text ("Equation: …"), which BTS Slides turns back into an equation.
 */
(function () {
  const $ = (s, r = document) => r.querySelector(s);
  const esc = (s) => BTS.esc(s);
  const X = window.SLX = window.SLX || { cmd: {}, ribbon: [] };
  const P = () => SL.pres, S = () => SL.slide();
  const SA = () => window.BTSSmartArt;
  const f2 = (n) => Math.round(n * 100) / 100;

  /* ====================================================================== */
  /*  SmartArt                                                               */
  /* ====================================================================== */
  const KIND = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', chevron: 'chevron', pentagon: 'homePlate', arrowRight: 'rightArrow', triangle: 'triangle', trapezoid: 'trapezoid' };
  const pal = () => SL.palette(P()).accents.slice(0, 6).map(c => SL.color(c, P()));
  const textHex = () => SL.color(SL.palette(P()).text || 'text', P());
  function tok(hex, list) {
    if (!hex || hex === 'none') return 'none';
    const i = list.findIndex(p => String(p).toLowerCase() === String(hex).toLowerCase()); if (i >= 0) return 'accent' + (i + 1);
    if (String(hex).toLowerCase() === textHex().toLowerCase()) return 'text';
    return hex;
  }
  function textHTML(t) {
    let h = '', ul = '';
    String(t || '').split('\n').forEach(l => {
      if (l.startsWith('• ')) { ul += '<li>' + esc(l.slice(2)) + '</li>'; return; }
      if (ul) { h += '<ul>' + ul + '</ul>'; ul = ''; }
      h += '<p>' + (esc(l) || '<br>') + '</p>';
    });
    if (ul) h += '<ul>' + ul + '</ul>';
    return h;
  }
  /** model (from BTSSmartArt.layout) → slide elements placed at (ox, oy) */
  X.smartToElements = (model, ox, oy, src) => {
    const list = pal(); const gid = SL.uid('g'); const nm = 'SmartArt ' + ((SA().LAYOUTS.find(l => l.id === src.layout) || {}).name || '');
    return model.shapes.map((s, k) => {
      const common = { g: gid, smart: SL.clone(src), name: nm + ' ' + (k + 1), opacity: s.opacity ?? 1, rot: s.rot || 0 };
      if (s.kind === 'line' || s.kind === 'connector') {
        return SL.newElement('shape', Object.assign(common, { kind: s.kind === 'connector' ? 'arrowLine' : 'line', x: ox + Math.min(s.x, s.x + s.w), y: oy + Math.min(s.y, s.y + s.h), w: Math.abs(s.w), h: Math.abs(s.h), flipH: s.w < 0, flipV: s.h < 0, fill: 'none', stroke: tok(s.stroke, list), strokeWidth: f2((s.strokeW || 1) * .75), html: '' }));
      }
      const custom = s.kind === 'circleArrow' || (s.kind === 'trapezoid' && s.flipV);
      const el = SL.newElement('shape', Object.assign(common, {
        kind: custom ? 'custom' : (KIND[s.kind] || 'rect'), x: ox + s.x, y: oy + s.y, w: s.w, h: s.h,
        fill: tok(s.fill, list), stroke: s.stroke && s.stroke !== 'none' && s.strokeW ? tok(s.stroke, list) : 'none', strokeWidth: f2((s.strokeW || 0) * .75),
        html: s.text ? textHTML(s.text) : '', color: s.textColor ? tok(s.textColor, list) : '#ffffff', size: s.fontSize ? Math.max(4, Math.round(s.fontSize * .75 * 2) / 2) : 18,
        bold: !!s.bold, align: s.align || 'center', valign: s.valign || 'middle',
      }));
      if (s.kind === 'circleArrow') el.path = { d: s.path, w: s.w, h: s.h };
      else if (s.kind === 'trapezoid' && s.flipV) { const o = s.inset ?? Math.min(s.w, s.h) * .25; el.path = { d: `M0,0 L${f2(s.w)},0 L${f2(s.w - o)},${f2(s.h)} L${f2(o)},${f2(s.h)} Z`, w: s.w, h: s.h }; }
      else if (s.kind === 'trapezoid') el.adj = Math.round((s.inset ?? Math.min(s.w, s.h) * .25) / Math.min(s.w, s.h) * 10000) / 10000;
      if (s.pad) { const gi = SL.geom(el.kind, el.w, el.h, el).inset; el.inset = [s.pad[0] - gi[0] * el.w, s.pad[1] - gi[1] * el.h, s.pad[2] - gi[2] * el.w, s.pad[3] - gi[3] * el.h].map(v => f2(Math.max(0, v))); }
      if (s.item != null) el.smartItem = s.item;
      return el;
    });
  };
  const bboxOf = (els) => { const x = Math.min(...els.map(e => e.x)), y = Math.min(...els.map(e => e.y)); return { x, y, w: Math.max(...els.map(e => e.x + e.w)) - x, h: Math.max(...els.map(e => e.y + e.h)) - y }; };
  X.buildSmartArt = (layout, items, box) => {
    const model = SA().layout(layout, items, { w: box.w, h: box.h, palette: pal(), font: SL.fontOf('body', P()), textColor: textHex() });
    const src = { layout, items: SL.clone(items), box: { x: box.x, y: box.y, w: box.w, h: box.h } };
    const els = X.smartToElements(model, box.x, box.y, src);
    if (els.length) { const bb = bboxOf(els); els.forEach(e => { e.smart.bb = bb; }); }
    return els;
  };
  X.cmd['insert-smartart'] = async () => {
    if (!SA()) return;
    if (E.editing) E.endEdit();
    const r = await SA().dialog({ palette: pal(), font: SL.fontOf('body', P()), textColor: textHex() }); if (!r || !r.items.length) return;
    const W = SL.W(P()), H = SL.H(P());
    const target = S().elements.find(e => e.ph === 'body' && SL.isEmptyText(e));
    const box = target ? { x: target.x, y: target.y, w: target.w, h: target.h } : { x: W * .08, y: H * .2, w: W * .84, h: H * .7 };
    if (target) S().elements = S().elements.filter(e => e !== target);
    const els = X.buildSmartArt(r.id, r.items, box); if (!els.length) return;
    S().elements.push(...els); E.sel = els.map(e => e.id);
    SL.commit(); E.renderCanvas(); E.refreshThumb(true); E.updateRibbon();
  };
  const smartMembers = (el) => el && el.smart ? (el.g ? S().elements.filter(e => e.g === el.g && e.smart) : [el]) : [];
  X.cmd['smartart-edit'] = async () => {
    const el = E.selected().find(e => e.smart); if (!el) return BTS.toast('Select a SmartArt diagram first');
    if (E.editing) E.endEdit();
    const mem = smartMembers(el); const src = el.smart; const items = SL.clone(src.items || []);
    mem.forEach(m => { if (m.smartItem != null && items[m.smartItem]) { const t = SL.plainText(m.html || '').replace(/\s+/g, ' ').trim(); if (t) items[m.smartItem].text = t; } }); // keep text typed on the slide
    const r = await SA().dialog({ palette: pal(), font: SL.fontOf('body', P()), textColor: textHex(), initial: { id: src.layout, items } }); if (!r || !r.items.length) return;
    const cur = bboxOf(mem), bb = src.bb || cur, box = src.box || cur;
    const sx = bb.w ? cur.w / bb.w : 1, sy = bb.h ? cur.h / bb.h : 1;
    const nb = { x: cur.x + (box.x - bb.x) * sx, y: cur.y + (box.y - bb.y) * sy, w: box.w * sx, h: box.h * sy };
    const els = X.buildSmartArt(r.id, r.items, nb); if (!els.length) return;
    const ids = new Set(mem.map(m => m.id)); const at = S().elements.findIndex(e => ids.has(e.id));
    S().elements = S().elements.filter(e => !ids.has(e.id)); S().elements.splice(Math.max(0, at), 0, ...els);
    E.sel = els.map(e => e.id); SL.commit(); E.renderCanvas(); E.refreshThumb(true); E.updateRibbon();
  };
  X.cmd['smartart-convert'] = () => {
    const el = E.selected().find(e => e.smart); if (!el) return;
    smartMembers(el).forEach(m => { delete m.smart; delete m.smartItem; });
    SL.commit(); E.updateRibbon(); BTS.toast('Converted to shapes (still grouped)');
  };

  /* ====================================================================== */
  /*  Equations                                                              */
  /* ====================================================================== */
  const TPLS = [['Quadratic', 'x = (-b \\pm sqrt(b^2 - 4ac))/(2a)'], ['Pythagoras', 'a^2 + b^2 = c^2'], ['Area of circle', 'A = \\pi r^2'], ['Binomial', '(a + b)^2 = a^2 + 2ab + b^2'], ['Sum', '\\sum_{i=1}^{n} i = (n(n+1))/(2)'], ['Einstein', 'E = mc^2'], ['Water', 'H_2O'], ['Speed', 'v = (d)/(t)'], ['Fraction', '(a)/(b)']];
  X.equationDialog = async (el) => {
    const body = document.createElement('div'); body.className = 'eq-dlg';
    const sizes = [16, 20, 24, 28, 32, 36, 44, 54, 66, 80];
    body.innerHTML = `<div class="eq-tpls">${TPLS.map(([n, e]) => `<button data-e="${esc(e)}" title="${esc(e)}">${n}</button>`).join('')}</div>
      <div class="field"><label>Equation (linear format)</label><input id="eq-in" value="${esc(el ? el.eq : 'a^2 + b^2 = c^2')}"></div>
      <div class="eq-prev" id="eq-prev"></div>
      <div class="field-row"><div class="field"><label>Size</label><select id="eq-size">${sizes.map(s => `<option value="${s}" ${(el ? el.size : 32) === s ? 'selected' : ''}>${s} pt</option>`).join('')}${el && !sizes.includes(el.size) ? `<option selected value="${el.size}">${el.size} pt</option>` : ''}</select></div></div>
      <div class="hint">Use ^ for superscript, _ for subscript, (a)/(b) for fractions, sqrt(x) for roots and \\alpha, \\pi, \\sum, \\int, \\pm, \\le for symbols.</div>`;
    const inp = $('#eq-in', body), prev = $('#eq-prev', body); const upd = () => { prev.innerHTML = '<span>' + SL.eqToHTML(inp.value) + '</span>'; };
    upd(); inp.addEventListener('input', upd);
    body.addEventListener('click', e => { const b = e.target.closest('[data-e]'); if (b) { inp.value = b.dataset.e; upd(); inp.focus(); } });
    return BTS.dialog({ title: el ? 'Edit equation' : 'Insert equation', body, width: 560, buttons: [{ label: 'Cancel', value: null }, { label: el ? 'Update' : 'Insert', primary: true, value: () => ({ eq: inp.value.trim(), size: +$('#eq-size', body).value || 32 }), validate: () => !!inp.value.trim() }] });
  };
  /** natural size (px) of an equation element */
  X.measureEquation = (el) => {
    const d = document.createElement('div'); d.className = 'eq-measure'; d.style.fontSize = (el.size || 28) + 'pt'; d.innerHTML = '<span>' + SL.eqToHTML(el.eq || '') + '</span>';
    document.body.appendChild(d); const r = { w: d.offsetWidth, h: d.offsetHeight }; d.remove();
    return { w: Math.ceil(r.w + (el.size || 28) * .9), h: Math.ceil(r.h + (el.size || 28) * .5) };
  };
  X.cmd['insert-equation'] = async () => {
    if (E.editing) E.endEdit();
    const r = await X.equationDialog(null); if (!r) return;
    await SL.ensureFonts(['Cambria']);
    const el = SL.newElement('equation', { eq: r.eq, size: r.size, color: 'text', name: 'Equation' });
    const m = X.measureEquation(el); const W = SL.W(P()), H = SL.H(P());
    Object.assign(el, { w: m.w, h: m.h, x: (W - m.w) / 2, y: (H - m.h) / 2 });
    E.addElement(el);
  };
  X.cmd['equation-edit'] = async () => {
    const el = E.primary(); if (!el || el.type !== 'equation') return BTS.toast('Select an equation first');
    const r = await X.equationDialog(el); if (!r) return;
    const old = X.measureEquation(el), cx = el.x + el.w / 2, cy = el.y + el.h / 2, kx = el.w / old.w, ky = el.h / old.h;
    el.eq = r.eq; el.size = r.size; const m = X.measureEquation(el);
    el.w = m.w * Math.max(1, kx); el.h = m.h * Math.max(1, ky); el.x = cx - el.w / 2; el.y = cy - el.h / 2;
    SL.commit(); E.renderCanvas(); E.refreshThumb(true); E.updateRibbon();
  };
  X.cmd['equation-color'] = async (btn) => {
    const els = E.selected().filter(e => e.type === 'equation'); if (!els.length) return BTS.toast('Select an equation first');
    const c = await BTS.colorPicker(btn, { auto: 'Automatic' }); if (!c) return;
    els.forEach(e => { e.color = c === 'auto' ? 'text' : c; }); SL.commit(); E.renderCanvas(); E.refreshThumb(true);
  };

  /* Equation → PNG (data URL) for .pptx / .odp. Drawn through an SVG foreignObject with the bundled Caladea face
   * embedded, so it looks like the slide. scale = output pixels per slide px. */
  let fontCSS = null;
  async function eqFontCSS() {
    if (fontCSS != null) return fontCSS;
    fontCSS = '';
    for (const [st, file] of [['normal', 'caladea-latin-400-normal'], ['italic', 'caladea-latin-400-italic']]) {
      try { const b = await (await fetch(`fonts/lib/caladea/${file}.woff2`)).arrayBuffer(); fontCSS += `@font-face{font-family:'EqFace';font-style:${st};font-weight:400;src:url(${await BTS.bytesToDataURL(new Uint8Array(b), 'font/woff2')}) format('woff2');}`; } catch { }
    }
    return fontCSS;
  }
  SL.equationPNG = async (el, pres = P(), scale = 4) => {
    const css = await eqFontCSS();
    const w = Math.max(1, Math.round(el.w)), h = Math.max(1, Math.round(el.h));
    const html = `<div xmlns="http://www.w3.org/1999/xhtml" style="width:${w}px;height:${h}px;display:flex;align-items:center;justify-content:center;font-family:EqFace,'Cambria Math',Cambria,'Times New Roman',serif;font-style:italic;white-space:nowrap;line-height:1.25;font-size:${el.size || 28}pt;color:${SL.color(el.color || 'text', pres)}"><style>${css}.frac{display:inline-flex;flex-direction:column;vertical-align:middle;text-align:center;font-size:90%;line-height:1.1}.frac>span:first-child{border-bottom:1px solid currentColor;padding:0 2px}</style><span>${SL.eqToHTML(el.eq || '').replace(/<br>/g, '<br/>')}</span></div>`;
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}"><foreignObject x="0" y="0" width="${w}" height="${h}">${html}</foreignObject></svg>`;
    const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg); });
    const c = document.createElement('canvas'); c.width = Math.round(w * scale); c.height = Math.round(h * scale);
    const g = c.getContext('2d'); g.scale(scale, scale); g.drawImage(img, 0, 0, w, h);
    return c.toDataURL('image/png');
  };
})();
