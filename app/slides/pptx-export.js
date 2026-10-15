/* BTS Slides — PowerPoint (.pptx) export via PptxGenJS, plus post-processing with JSZip for
 * features PptxGenJS lacks: slide transitions + auto-advance timings, gradient backgrounds/fills, hidden slides,
 * alt text / decorative flags on every object, Devanagari (cs) fonts, document properties (keywords, category,
 * description), comments (legacy threaded comments), custom shows + slide show settings, and BTS metadata
 * (customXml part: master, theme overrides, custom shows and an exact copy of every slide that is restored
 * on reopen while the slide is unchanged — keeps animations, groups, table styles, comment anchors…).
 * Units: model px (96/in) → inches; font sizes stay in pt.
 */
(function () {
  const X = window.PPTXExport = {};
  const f3 = (n) => Math.round(n * 1000) / 1000;
  const IN = (px) => f3(px / 96);
  const P = () => SL.pres;

  function hex(c, pres) {
    let v = SL.color(c, pres); if (!v || v === 'none' || v === 'transparent') return null;
    if (v.startsWith('rgb')) { const m = v.match(/[\d.]+/g).map(Number); return { c: m.slice(0, 3).map(x => Math.round(x).toString(16).padStart(2, '0')).join('').toUpperCase(), a: m[3] ?? 1 }; }
    if (v.startsWith('color-mix') || v.startsWith('var(')) return null;
    v = v.replace('#', ''); if (v.length === 3) v = v.split('').map(x => x + x).join('');
    const a = v.length === 8 ? parseInt(v.slice(6), 16) / 255 : 1;
    return { c: v.slice(0, 6).toUpperCase(), a };
  }
  const hx = (c, pres) => (hex(c, pres) || {}).c;
  function cssColorToHex(css) { if (!css) return null; const d = document.createElement('i'); d.style.color = css; document.body.appendChild(d); const v = getComputedStyle(d).color; d.remove(); const m = v.match(/[\d.]+/g); if (!m) return null; if (m[3] !== undefined && +m[3] === 0) return null; return m.slice(0, 3).map(x => (+x).toString(16).padStart(2, '0')).join('').toUpperCase(); }

  /* ---------- raster helpers ---------- */
  function loadImg(src) { return new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = src; }); }
  async function svgToPNG(svg, w, h, scale = 3) {
    const url = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    const img = await loadImg(url);
    const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(w * scale)); c.height = Math.max(1, Math.round(h * scale));
    c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
    return c.toDataURL('image/png');
  }
  X.svgToPNG = svgToPNG;
  X.processImage = (el) => processImage(el);
  /* custom shape path (el.path, box-relative px) → PptxGenJS custGeom points (inches) */
  X.pathPoints = (el) => {
    const g = SL.geom('custom', el.w, el.h, el); const out = [];
    SL.pathSegs(g.d).forEach(s => {
      if (s.c === 'M') out.push({ x: IN(s.p[0]), y: IN(s.p[1]), moveTo: true });
      else if (s.c === 'L') out.push({ x: IN(s.p[0]), y: IN(s.p[1]) });
      else if (s.c === 'C') out.push({ x: IN(s.p[4]), y: IN(s.p[5]), curve: { type: 'cubic', x1: IN(s.p[0]), y1: IN(s.p[1]), x2: IN(s.p[2]), y2: IN(s.p[3]) } });
      else if (s.c === 'Z') out.push({ close: true });
    });
    return out;
  };
  async function processImage(el) {
    const url = SL.assetURL(el.src); if (!url) return null;
    const crop = el.crop || { l: 0, t: 0, r: 0, b: 0 };
    const needCrop = crop.l > .001 || crop.t > .001 || crop.r > .001 || crop.b > .001;
    const isSvg = url.startsWith('data:image/svg');
    if (!needCrop && !el.radius && !isSvg && /^data:image\/(png|jpe?g|gif)/.test(url)) return url;
    const img = await loadImg(url);
    let nw = img.naturalWidth || el.w * 3, nh = img.naturalHeight || el.h * 3;
    if (isSvg) { const k = Math.max(1, 1600 / Math.max(nw, nh)); nw *= k; nh *= k; }
    const sx = crop.l * nw, sy = crop.t * nh, sw = nw * (1 - crop.l - crop.r), sh = nh * (1 - crop.t - crop.b);
    const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(sw)); c.height = Math.max(1, Math.round(sh));
    const g = c.getContext('2d');
    if (el.radius) { const r = Math.min(c.width, c.height) * el.radius / 100; g.beginPath(); g.roundRect(0, 0, c.width, c.height, r); g.clip(); }
    g.drawImage(img, sx, sy, sw, sh, 0, 0, c.width, c.height);
    return c.toDataURL(/jpe?g/.test(url.slice(0, 20)) && !el.radius ? 'image/jpeg' : 'image/png', .92);
  }
  X.bgImage = (bg, W, H) => bgImage(bg, W, H);
  async function bgImage(bg, W, H) {
    const url = SL.assetURL(bg.src); if (!url) return null;
    if (bg.fit === 'stretch') return url;
    const img = await loadImg(url); const c = document.createElement('canvas'); const k = Math.min(3, Math.max(1, img.naturalWidth / W)); c.width = Math.round(W * k); c.height = Math.round(H * k);
    const g = c.getContext('2d'); g.fillStyle = SL.color(bg.color || '#ffffff', P()); g.fillRect(0, 0, c.width, c.height);
    const s = bg.fit === 'contain' ? Math.min(c.width / img.naturalWidth, c.height / img.naturalHeight) : Math.max(c.width / img.naturalWidth, c.height / img.naturalHeight);
    const dw = img.naturalWidth * s, dh = img.naturalHeight * s; g.drawImage(img, (c.width - dw) / 2, (c.height - dh) / 2, dw, dh);
    return c.toDataURL('image/jpeg', .9);
  }

  /* ---------- rich text → PptxGenJS runs ---------- */
  X.runs = (el, pres, base = {}) => {
    const d = SL.textDefaults(el, pres);
    const def = Object.assign({ fontFace: d.font, fontSize: el.size || 18, color: hx(d.color, pres) || '000000', bold: !!d.bold, italic: !!el.italic, underline: el.underline ? { style: 'sng' } : undefined, strike: el.strike ? 'sngStrike' : undefined }, base);
    if (el.effect === 'outline') def.outline = { size: 1, color: def.color };
    if (el.effect === 'glow' || el.effect === 'neon') def.glow = { size: 8, opacity: .6, color: hx('accent1', pres) };
    if (el.spacing) def.charSpacing = el.spacing;
    const scale = el.autofit === 'shrink' && el.fontScale && el.fontScale < 1 ? el.fontScale : 1;
    if (scale !== 1) def.fontSize = Math.round(def.fontSize * scale * 2) / 2;
    const box = document.createElement('div'); box.innerHTML = el.html || '';
    const out = [];
    const blocks = [];
    for (const n of box.childNodes) {
      if (n.nodeType === 1 && (n.tagName === 'UL' || n.tagName === 'OL')) [...n.children].forEach((li, i) => blocks.push({ node: li, list: n.tagName, num: i }));
      else if (n.nodeType === 1) blocks.push({ node: n });
      else if (n.nodeType === 3 && n.nodeValue.trim()) { const p = document.createElement('p'); p.textContent = n.nodeValue; blocks.push({ node: p }); }
    }
    if (!blocks.length) return [{ text: '', options: Object.assign({}, def, { align: el.align || 'left' }) }];
    blocks.forEach((b, bi) => {
      const para = { align: (b.node.style && b.node.style.textAlign) || el.align || 'left' };
      if (para.align === 'start') para.align = 'left'; if (para.align === 'end') para.align = 'right';
      const lvl = +(b.node.dataset && b.node.dataset.lvl) || 0;
      if (b.list === 'UL') para.bullet = lvl ? { indent: 18 } : true;
      if (b.list === 'OL') para.bullet = { type: 'number' };
      if (lvl) para.indentLevel = lvl;
      if (el.lh && el.lh !== 1) para.lineSpacingMultiple = f3(el.lh * 1.0);
      if (el.pspace) para.paraSpaceAfter = el.pspace;
      const runs = [];
      const walk = (node, st, softBreak) => {
        for (const ch of node.childNodes) {
          if (ch.nodeType === 3) { if (ch.nodeValue) runs.push({ text: ch.nodeValue, st: Object.assign({}, st), soft: softBreak.v }); softBreak.v = false; continue; }
          if (ch.nodeType !== 1) continue;
          if (ch.tagName === 'BR') { softBreak.v = true; continue; }
          const s = Object.assign({}, st); const cs = ch.style || {};
          const t = ch.tagName;
          if (t === 'B' || t === 'STRONG') s.bold = true; if (t === 'I' || t === 'EM') s.italic = true; if (t === 'U') s.underline = { style: 'sng' }; if (t === 'S' || t === 'STRIKE' || t === 'DEL') s.strike = 'sngStrike';
          if (t === 'SUB') s.subscript = true; if (t === 'SUP') s.superscript = true;
          if (t === 'A' && ch.getAttribute('href')) s.hyperlink = { url: ch.getAttribute('href') };
          if (cs.fontWeight) s.bold = cs.fontWeight === 'bold' || +cs.fontWeight >= 600; if (cs.fontStyle) s.italic = cs.fontStyle === 'italic';
          const dec = cs.textDecorationLine || cs.textDecoration || ''; if (dec.includes('underline')) s.underline = { style: 'sng' }; if (dec.includes('line-through')) s.strike = 'sngStrike';
          if (cs.color) { const c = cssColorToHex(cs.color); if (c) s.color = c; }
          if (cs.backgroundColor) { const c = cssColorToHex(cs.backgroundColor); if (c) s.highlight = c; }
          if (cs.fontSize) { const m = /([\d.]+)(pt|px)/.exec(cs.fontSize); if (m) s.fontSize = Math.round((m[2] === 'px' ? +m[1] * .75 : +m[1]) * scale * 2) / 2; }
          if (cs.fontFamily) s.fontFace = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
          walk(ch, s, softBreak);
        }
      };
      const sb = { v: false }; walk(b.node, {}, sb);
      if (!runs.length) runs.push({ text: '', st: {} });
      runs.forEach((r, ri) => {
        const o = Object.assign({}, def, r.st, para);
        if (r.soft) o.softBreakBefore = true;
        if (ri === runs.length - 1 && bi < blocks.length - 1) o.breakLine = true;
        out.push({ text: r.text, options: o });
      });
    });
    return out;
  };

  const SHAPE_MAP = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', rtTriangle: 'rtTriangle', diamond: 'diamond', parallelogram: 'parallelogram', trapezoid: 'trapezoid', pentagon: 'pentagon', hexagon: 'hexagon', octagon: 'octagon', star4: 'star4', star5: 'star5', star6: 'star6', plus: 'plus', donut: 'donut', heart: 'heart', cloud: 'cloud', can: 'can', rightArrow: 'rightArrow', leftArrow: 'leftArrow', upArrow: 'upArrow', downArrow: 'downArrow', leftRightArrow: 'leftRightArrow', chevron: 'chevron', homePlate: 'homePlate', wedgeRectCallout: 'wedgeRectCallout', wedgeRoundRectCallout: 'wedgeRoundRectCallout', wedgeEllipseCallout: 'wedgeEllipseCallout', line: 'line', arrowLine: 'line' };
  const DASH_MAP = { solid: 'solid', dash: 'dash', lgDash: 'lgDash', dot: 'sysDot', dashDot: 'dashDot', sysDash: 'sysDash' };

  function lineOpts(el, pres) {
    const c = hex(el.stroke, pres); if (!c || !el.strokeWidth) return undefined;
    const o = { color: c.c, width: el.strokeWidth, dashType: DASH_MAP[el.dash] || 'solid' }; if (c.a < 1) o.transparency = Math.round((1 - c.a) * 100);
    return o;
  }
  function fillOpts(fill, el, pres, post, name) {
    if (!fill || fill === 'none') return undefined;
    let c;
    if (typeof fill === 'object' && fill.type === 'gradient') { post.grads[name] = { c1: hx(fill.color, pres), c2: hx(fill.color2, pres), angle: fill.angle ?? 90, alpha: el.opacity ?? 1 }; c = hex(fill.color, pres); }
    else c = hex(fill, pres);
    if (!c) return undefined;
    const o = { color: c.c }; const op = (el.opacity ?? 1) * c.a; if (op < 1) o.transparency = Math.round((1 - op) * 100);
    return o;
  }
  const shadowOpt = (el) => el.shadow ? { type: 'outer', blur: 8, offset: 4, angle: 45, color: '000000', opacity: .3 } : undefined;

  /* ---------- slide masters / layouts (one PptxGenJS master per BTS layout in use) ---------- */
  async function defineMasters(pptx, pres) {
    const W = SL.W(pres), H = SL.H(pres); const out = {};
    for (const L of [...new Set(pres.slides.map(s => s.layout))]) {
      const name = 'BTS ' + SL.layoutName(L);
      const bg = SL.themeBg(pres, L); let background;
      if (bg.type === 'image' && bg.src) { const d = await bgImage(bg, W, H); if (d) background = { data: d }; }
      if (!background) background = { color: hx(bg.color || 'bg', pres) || 'FFFFFF' };
      const objects = [];
      SL.layoutPlaceholders(L, W, H, pres).forEach((sp, i) => {
        if (sp.type !== 'text') return;
        const d = SL.textDefaults(Object.assign({ ph: sp.ph }, sp), pres);
        objects.push({ placeholder: { options: { name: (sp.ph === 'title' ? 'title' : 'body') + i, type: sp.ph === 'title' ? 'title' : 'body', x: IN(sp.x), y: IN(sp.y), w: IN(sp.w), h: IN(sp.h), fontFace: d.font, fontSize: sp.size || 18, color: hx(d.color, pres) || '000000', bold: !!d.bold, italic: !!d.italic, align: sp.align || 'left', valign: sp.valign === 'middle' ? 'middle' : sp.valign === 'bottom' ? 'bottom' : 'top' }, text: '' } });
      });
      try { pptx.defineSlideMaster({ title: name, background, objects }); out[L] = name; } catch (err) { console.warn('slide master', L, err); }
    }
    return out;
  }

  /* ---------- main ---------- */
  X.build = async (pres = P()) => {
    if (!window.PptxGenJS) throw new Error('PptxGenJS is not loaded');
    await SL.ensureFonts(pres);
    const pptx = new window.PptxGenJS();
    const W = SL.W(pres), H = SL.H(pres);
    pptx.defineLayout({ name: 'BTS', width: IN(W), height: IN(H) }); pptx.layout = 'BTS';
    const props = Object.assign({}, BTS.props || {});
    pptx.author = props.author || BTS.userName || 'BTS Slides'; pptx.company = props.company || ''; pptx.subject = props.subject || ''; pptx.title = props.title || BTS.title || 'Presentation';
    const tf = { title: SL.fontOf('title', pres), body: SL.fontOf('body', pres) };
    try { pptx.theme = { headFontFace: tf.title, bodyFontFace: tf.body }; } catch { }
    const post = { trans: {}, grads: {}, bgGrads: {}, hidden: {}, descr: {}, names: [], comments: [], props, fonts: tf, pres, masterNames: {} };
    const masters = await defineMasters(pptx, pres);
    const decoCache = {};
    const secs = pres.sections || []; const usedSec = {};
    for (let i = 0; i < pres.slides.length; i++) {
      const s = pres.slides[i];
      post.names[i] = {};
      const sec = secs.find(x => x.start === s.id);
      if (sec) { let t = sec.name || 'Section'; while (usedSec[t]) t += ' '; usedSec[t] = 1; pptx.addSection({ title: t }); }
      const sopt = {}; if (Object.keys(usedSec).length) sopt.sectionTitle = Object.keys(usedSec).pop(); if (masters[s.layout]) sopt.masterName = masters[s.layout];
      const slide = pptx.addSlide(sopt);
      // background
      const bg = s.background || SL.themeBg(pres, s.layout);
      if (bg.type === 'image' && bg.src) { const d = await bgImage(bg, W, H); if (d) slide.background = { data: d }; }
      else if (bg.type === 'gradient') { slide.background = { color: hx(bg.color, pres) || 'FFFFFF' }; post.bgGrads[i] = { c1: hx(bg.color, pres), c2: hx(bg.color2, pres), angle: bg.angle ?? 135 }; }
      else slide.background = { color: hx(bg.color || 'bg', pres) || 'FFFFFF' };
      if (s.hidden) { slide.hidden = true; post.hidden[i] = true; }
      if (s.transition && ((s.transition.type && s.transition.type !== 'none') || s.transition.auto != null || s.transition.click === false)) post.trans[i] = s.transition;
      // theme graphics as a picture at the back (re-created from the theme on import)
      const kind = SL.decoKind(s.layout);
      const tag = `bts:theme=${pres.theme.id};v=${pres.theme.variant || 0};layout=${s.layout}`;
      if (!s.hideGraphics && !s.noDeco) {
        const key = kind + (s.layout === 'pictureCaption' ? ':pc' : ''); if (!decoCache[key]) { const svg = SL.decoSVG(Object.assign({}, s, { hideGraphics: false }), pres); decoCache[key] = svg ? await svgToPNG(svg.replace('<svg ', '<svg xmlns="http://www.w3.org/2000/svg" '), W, H, 1) : null; }
        if (decoCache[key]) slide.addImage({ data: decoCache[key], x: 0, y: 0, w: IN(W), h: IN(H), altText: tag, objectName: 'Theme graphics' });
      } else slide.addText('', { x: 0, y: 0, w: .01, h: .01, objectName: 'Theme graphics', altText: tag });
      // slide master / layout graphics (also kept in the BTS metadata)
      if (!s.hideGraphics && pres.master) for (const el of [...(pres.master.graphics || []), ...((pres.master.layoutGraphics || {})[s.layout] || [])]) { const nm = await addEl(slide, Object.assign({}, el, { name: 'Master graphic' }), pres, post, i, true); if (nm && !post.masterNames[el.id]) post.masterNames[el.id] = { name: nm, slide: i }; }
      // imported layout/master graphics
      for (const el of s.layer || []) await addEl(slide, el, pres, post, i, true);
      for (const el of s.elements) await addEl(slide, el, pres, post, i);
      for (const el of SL.footerEls(s, pres, i)) await addEl(slide, el, pres, post, i);
      if (s.notes) slide.addNotes(s.notes);
      post.comments[i] = (s.comments || []).filter(c => c && c.text);
    }
    let data = await pptx.write({ outputType: 'uint8array' });
    data = await X.postProcess(data, post);
    return data;
  };

  let objN = 0;
  async function addEl(slide, el, pres, post, si, isLayer) {
    if (el.hidden) return;
    if (el.ph && ((el.type === 'text' && SL.isEmptyText(el)) || (el.type === 'image' && !el.src))) return; // empty placeholders are editor-only
    const g = { x: IN(el.x), y: IN(el.y), w: Math.max(.01, IN(el.w)), h: Math.max(.01, IN(el.h)) };
    if (el.rot) g.rotate = Math.round(el.rot);
    if (el.flipH) g.flipH = true; if (el.flipV) g.flipV = true;
    const name = (el.ph === 'title' ? 'Title ' : el.ph === 'subtitle' ? 'Subtitle ' : el.ph === 'body' ? 'Text Placeholder ' : el.footer ? 'Footer ' : (el.name ? el.name.replace(/[^\w .-]/g, '').slice(0, 40) + ' ' : '')) + 'O' + (++objN);
    if (post.names[si] && el.id) post.names[si][el.id] = name;
    const alt = el.decorative ? '' : (el.alt || (el.type === 'icon' ? 'Icon: ' + el.name : ''));
    if (alt || el.decorative) post.descr[name] = { alt, decorative: !!el.decorative };
    await addEl2(slide, el, pres, post, si, isLayer, g, name);
    return name;
  }
  async function addEl2(slide, el, pres, post, si, isLayer, g, name) {
    if (el.type === 'text' || el.type === 'shape') {
      if (SL.isLine(el)) {
        const lo = lineOpts(el, pres) || { color: hx('text', pres), width: 2 };
        if (el.kind === 'arrowLine') lo.endArrowType = 'triangle';
        slide.addShape('line', Object.assign(g, { line: lo, objectName: name, shadow: shadowOpt(el) }));
        return;
      }
      const runs = X.runs(el, pres);
      const ins = el.inset || [9.6, 4.8, 9.6, 4.8];
      const o = Object.assign(g, {
        shape: el.type === 'shape' ? (SHAPE_MAP[el.kind] || 'rect') : 'rect',
        valign: el.valign === 'middle' ? 'middle' : el.valign === 'bottom' ? 'bottom' : 'top', align: el.align || 'left',
        margin: [ins[0] * .75, ins[2] * .75, ins[3] * .75, ins[1] * .75], objectName: name, wrap: el.wrap !== false,
        fill: fillOpts(el.fill, el, pres, post, name), line: lineOpts(el, pres), shadow: shadowOpt(el),
      });
      if (el.autofit === 'shrink') o.fit = 'shrink'; else if (el.autofit === 'resize') o.fit = 'resize';
      if (el.kind === 'roundRect') o.rectRadius = Math.min(el.w, el.h) * (el.adj ?? .1667) / 96;
      if (el.kind === 'trapezoid' && el.adj != null) o.rectRadius = Math.min(el.w, el.h) * el.adj / 96; // PptxGenJS writes it as the shape's "adj" guide
      if (el.type === 'shape' && el.kind === 'custom' && el.path) { o.shape = 'custGeom'; o.points = X.pathPoints(el); }
      if (el.dir === 'vert') o.vert = 'vert'; else if (el.dir === 'vert270') o.vert = 'vert270';
      if (el.lh && el.lh !== 1) o.lineSpacingMultiple = el.lh;
      if (!o.fill) delete o.fill; if (!o.line) delete o.line; if (!o.shadow) delete o.shadow;
      slide.addText(runs, o);
    } else if (el.type === 'image') {
      const data = await processImage(el); if (!data) return;
      const o = Object.assign(g, { data, objectName: name, altText: el.alt || '' });
      if (el.opacity != null && el.opacity < 1) o.transparency = Math.round((1 - el.opacity) * 100);
      slide.addImage(o);
    } else if (el.type === 'equation') {
      if (!SL.equationPNG) return;
      const data = await SL.equationPNG(el, pres, 4);
      const o = Object.assign(g, { data, objectName: name, altText: 'Equation: ' + (el.eq || '') });
      if (el.opacity != null && el.opacity < 1) o.transparency = Math.round((1 - el.opacity) * 100);
      slide.addImage(o);
    } else if (el.type === 'icon') {
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${el.w}" height="${el.h}" viewBox="0 0 24 24" preserveAspectRatio="none" fill="${el.fillIcon ? SL.color(el.color || 'accent1', pres) : 'none'}" stroke="${SL.color(el.color || 'accent1', pres)}" stroke-width="${el.strokeWidth || 1.75}" stroke-linecap="round" stroke-linejoin="round">${(window.BTS_ICONS || {})[el.name] || ''}</svg>`;
      const data = await svgToPNG(svg, el.w, el.h, 4);
      const o = Object.assign(g, { data, objectName: 'Icon ' + el.name + ' O' + objN, altText: el.alt || ('Icon: ' + el.name) });
      if (el.opacity != null && el.opacity < 1) o.transparency = Math.round((1 - el.opacity) * 100);
      slide.addImage(o);
    } else if (el.type === 'table') {
      const rows = el.rows.map((row, r) => {
        const out = [];
        row.forEach((c, ci) => {
          if (c.merged) return;
          const ts = SL.tableCellStyle(el, r, ci, pres);
          const fill = c.fill === 'none' ? null : c.fill ? hx(c.fill, pres) : (ts.bg ? hx(ts.bg, pres) : null);
          const color = c.color ? hx(c.color, pres) : (hx(ts.color, pres) || '1F2430');
          const cellEl = { html: c.html || '', size: el.size || 16, font: el.font || SL.fontOf('body', pres), color: '#' + color, bold: ts.bold || c.bold, align: c.align || 'left', role: 'body', type: 'text' };
          const runs = X.runs(cellEl, pres);
          const border = ts.b.map(b => b ? { type: 'solid', pt: Math.round(b.w * .75 * 100) / 100, color: hx(b.c, pres) || 'FFFFFF' } : { type: 'none' });
          const opt = { fill: fill ? { color: fill } : undefined, valign: c.valign || 'middle', align: c.align || 'left', border, margin: [4, 6, 4, 6] };
          if (c.cs > 1) opt.colspan = c.cs; if (c.rs > 1) opt.rowspan = c.rs;
          out.push({ text: runs, options: opt });
        });
        return out;
      });
      const colW = (el.colW || []).map(c => f3(c * el.w / 96));
      slide.addTable(rows, Object.assign(g, { colW, rowH: el.rows.map((_, r) => f3((el.rowH?.[r] ? el.rowH[r] * el.h : el.h / el.rows.length) / 96)), fontFace: el.font || SL.fontOf('body', pres), fontSize: el.size || 16, objectName: name }));
    } else if (el.type === 'chart') {
      const ch = el.chart; if (!ch) return;
      const types = { bar: 'bar', line: 'line', pie: 'pie', doughnut: 'doughnut' };
      const pal = (ch.colors && ch.colors.length ? ch.colors : SL.palette(pres).accents).map(c => hx(c, pres));
      const series = (ch.type === 'pie' || ch.type === 'doughnut') ? ch.series.slice(0, 1) : ch.series;
      const data = series.map(s => ({ name: s.name, labels: ch.categories.slice(), values: s.values.map(v => +v || 0) }));
      const txt = hx('text', pres);
      slide.addChart(types[ch.type] || 'bar', data, Object.assign(g, {
        barDir: 'col', barGrouping: 'clustered', chartColors: pal, showLegend: ch.legend !== false, legendPos: 'b', legendFontSize: 11, legendColor: txt,
        showTitle: !!ch.title, title: ch.title || '', titleFontSize: 16, titleColor: txt, catAxisLabelColor: txt, valAxisLabelColor: txt,
        showPercent: ch.type === 'pie' || ch.type === 'doughnut', dataLabelColor: 'FFFFFF', lineSize: 3, lineDataSymbolSize: 7, objectName: name,
        valGridLine: { color: 'D9D9D9', size: .5 }, catGridLine: { style: 'none' }, holeSize: 55,
      }));
    }
  }

  /* ---------- post processing (JSZip) ---------- */
  const TRANS_XML = (t) => {
    const spd = (t.dur || 700) < 500 ? 'fast' : (t.dur || 700) < 1000 ? 'med' : 'slow';
    const dir = t.dir || 'left'; const d4 = { left: 'l', right: 'r', up: 'u', down: 'd' }[dir] || 'l';
    // PowerPoint 2010+ duration attribute lives in the p14 namespace; we use the classic element + speed for compatibility.
    let inner;
    switch (t.type) {
      case 'fade': inner = '<p:fade/>'; break;
      case 'push': inner = `<p:push dir="${{ left: 'r', right: 'l', up: 'd', down: 'u' }[dir] || 'r'}"/>`; break; // pptx dir = where it comes FROM
      case 'wipe': inner = `<p:wipe dir="${{ left: 'r', right: 'l', up: 'd', down: 'u' }[dir] || 'r'}"/>`; break;
      case 'split': inner = `<p:split orient="${dir === 'up' || dir === 'down' ? 'horz' : 'vert'}" dir="out"/>`; break;
      case 'cover': inner = `<p:cover dir="${{ left: 'r', right: 'l', up: 'd', down: 'u' }[dir] || 'r'}"/>`; break;
      case 'reveal': inner = `<p:pull dir="${d4}"/>`; break;
      case 'zoom': inner = '<p:zoom/>'; break;
      case 'flip': case 'cube': inner = `<p:cover dir="${{ left: 'r', right: 'l', up: 'd', down: 'u' }[dir] || 'r'}"/>`; break;
      default: inner = '<p:fade/>';
    }
    const adv = t.auto != null && t.auto !== false ? ` advTm="${Math.round(t.auto * 1000)}"` : '';
    const click = t.click === false ? ' advClick="0"' : '';
    if (!t.type || t.type === 'none') return `<p:transition${click}${adv}/>`;
    // mc:AlternateContent so PowerPoint picks up p14 effects (flip/zoom duration) while older readers use the fallback
    const p14 = { flip: `<p14:flip dir="${dir === 'right' ? 'r' : 'l'}"/>`, cube: `<p14:prism dir="${d4}"/>`, zoom: '<p14:ripple/>' }[t.type];
    if (p14 && t.type !== 'zoom') return `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14"><p:transition spd="${spd}" p14:dur="${t.dur || 700}"${click}${adv}>${p14}</p:transition></mc:Choice><mc:Fallback><p:transition spd="${spd}"${click}${adv}>${inner}</p:transition></mc:Fallback></mc:AlternateContent>`;
    return `<p:transition spd="${spd}"${click}${adv}>${inner}</p:transition>`;
  };
  const gradXML = (g, alpha = 1) => {
    const a = alpha < 1 ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : '';
    // CSS angle: 0deg = to top, 90 = to right. DrawingML lin ang: 0 = left→right, clockwise. → ang = css - 90
    const ang = ((((g.angle ?? 90) - 90) % 360) + 360) % 360;
    return `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="${g.c1 || 'FFFFFF'}">${a}</a:srgbClr></a:gs><a:gs pos="100000"><a:srgbClr val="${g.c2 || 'FFFFFF'}">${a}</a:srgbClr></a:gs></a:gsLst><a:lin ang="${Math.round(ang * 60000)}" scaled="0"/></a:gradFill>`;
  };
  // PptxGenJS emits <a:pPr> before every run; only one (the first child of <a:p>) is valid.
  X.fixParagraphs = (xml) => xml.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (m, body) => {
    const first = body.search(/<a:(r|br|fld)\b/); if (first < 0) return m;
    const head = body.slice(0, first), tail = body.slice(first).replace(/<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/g, '');
    return '<a:p>' + head + tail + '</a:p>';
  });
  const xesc = (t) => String(t ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  X.hash = (str) => { let h = 0x811c9dc5; for (let i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; } return h.toString(16).padStart(8, '0') + ':' + str.length; };
  const DEVA = /[ऀ-ॿ꣠-ꣿ]/;
  X.devaFont = (latinFont, serif) => { const f = (window.BTS_FONTS || []).find(x => x.name === latinFont); if (f && f.cat === 'indian') return latinFont; return serif ? 'Noto Serif Devanagari' : 'Noto Sans Devanagari'; };
  const isSerif = (f) => { const r = (window.BTS_FONTS || []).find(x => x.name === f); return r ? r.cat === 'serif' : /serif|garamond|baskerville|times|georgia|cambria/i.test(f) && !/sans/i.test(f); };
  // Runs with Devanagari text: complex-script font = a Devanagari face, language = Hindi.
  X.fixDevanagari = (xml) => xml.replace(/<a:r>(<a:rPr\b[^>]*?)(\/>|>([\s\S]*?)<\/a:rPr>)<a:t>([^<]*)<\/a:t><\/a:r>/g, (m, open, close, inner, text) => {
    if (!DEVA.test(text)) return m;
    const lat = /<a:latin typeface="([^"]+)"/.exec(inner || ''); const dev = X.devaFont(lat ? lat[1] : '', lat && isSerif(lat[1]));
    let o = open.replace(/ lang="[^"]*"/, '') + ' lang="hi-IN"';
    let body = inner || '';
    if (/<a:cs\b[^>]*\/>/.test(body)) body = body.replace(/<a:cs\b[^>]*\/>/, `<a:cs typeface="${xesc(dev)}" pitchFamily="2" charset="0"/>`);
    else body += `<a:cs typeface="${xesc(dev)}" pitchFamily="2" charset="0"/>`;
    return `<a:r>${o}>${body}</a:rPr><a:t>${text}</a:t></a:r>`;
  });
  // alt text (descr) + decorative flag on the object's cNvPr, matched by object name
  X.setDescr = (xml, descr) => xml.replace(/<p:cNvPr\b([^>]*?)(\/>|>)/g, (m, attrs, end) => {
    const nm = /\bname="([^"]*)"/.exec(attrs); if (!nm) return m; const d = descr[nm[1].replace(/&amp;/g, '&')]; if (!d) return m;
    let a = attrs.replace(/\s+descr="[^"]*"/, ''); if (d.alt) a += ` descr="${xesc(d.alt)}"`;
    if (!d.decorative || d.keepAlt) return `<p:cNvPr${a}${end}`;
    const ext = '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>';
    return end === '/>' ? `<p:cNvPr${a}>${ext}</p:cNvPr>` : `<p:cNvPr${a}>${ext}`;
  });
  const NS_P = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  async function addRel(zip, relsPath, type, target) {
    let x = zip.file(relsPath) ? await zip.file(relsPath).async('string') : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const ids = [...x.matchAll(/Id="rId(\d+)"/g)].map(m => +m[1]); const id = 'rId' + ((ids.length ? Math.max(...ids) : 0) + 1);
    x = x.replace('</Relationships>', `<Relationship Id="${id}" Type="${REL}${type}" Target="${target}"/></Relationships>`); zip.file(relsPath, x); return id;
  }
  async function addOverride(zip, part, ct) { let x = await zip.file('[Content_Types].xml').async('string'); if (!x.includes(`PartName="${part}"`)) { x = x.replace('</Types>', `<Override PartName="${part}" ContentType="${ct}"/></Types>`); zip.file('[Content_Types].xml', x); } }

  X.postProcess = async (data, post) => {
    const zip = await window.JSZip.loadAsync(data);
    const n = Object.keys(zip.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f)).length;
    const pres = post.pres || P();
    for (let i = 0; i < n; i++) {
      const path = `ppt/slides/slide${i + 1}.xml`; let xml = await zip.file(path).async('string'); const orig = xml;
      xml = X.fixParagraphs(xml);
      if (post.bgGrads[i]) xml = xml.replace(/<p:bg>[\s\S]*?<\/p:bg>/, `<p:bg><p:bgPr>${gradXML(post.bgGrads[i])}<a:effectLst/></p:bgPr></p:bg>`);
      for (const [name, g] of Object.entries(post.grads)) {
        const at = xml.indexOf(`name="${name}"`); if (at < 0) continue;
        const spStart = xml.lastIndexOf('<p:sp>', at); const spEnd = xml.indexOf('</p:sp>', at);
        const seg = xml.slice(spStart, spEnd); const pr = seg.indexOf('<p:spPr'); const prEnd = seg.indexOf('</p:spPr>');
        if (pr < 0 || prEnd < 0) continue;
        const spPr = seg.slice(pr, prEnd).replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, gradXML(g, g.alpha));
        xml = xml.slice(0, spStart) + seg.slice(0, pr) + spPr + seg.slice(prEnd) + xml.slice(spEnd);
      }
      if (post.hidden[i] && !/<p:sld[^>]*show="0"/.test(xml)) xml = xml.replace(/<p:sld /, '<p:sld show="0" ');
      if (post.trans[i]) xml = xml.replace(/<p:transition[\s\S]*?(<\/p:transition>|\/>)/, '').replace('</p:cSld>', '</p:cSld>__CLR__').replace(/__CLR__(<p:clrMapOvr>[\s\S]*?<\/p:clrMapOvr>)?/, (m, clr) => (clr || '') + TRANS_XML(post.trans[i]));
      xml = X.fixDevanagari(xml);
      xml = X.setDescr(xml, post.descr || {});
      if (xml !== orig) zip.file(path, xml);
    }
    // theme fonts: Devanagari complex-script faces in the font scheme
    for (const tp of Object.keys(zip.files).filter(f => /^ppt\/theme\/theme\d+\.xml$/.test(f))) {
      let t = await zip.file(tp).async('string'); const f = post.fonts || {};
      const fix = (block, latin) => block.replace(/<a:cs typeface="[^"]*"\s*\/>/, `<a:cs typeface="${xesc(X.devaFont(latin, isSerif(latin)))}"/>`).replace(/<a:font script="Deva" typeface="[^"]*"\s*\/>/, `<a:font script="Deva" typeface="${xesc(X.devaFont(latin, isSerif(latin)))}"/>`);
      t = t.replace(/<a:majorFont>[\s\S]*?<\/a:majorFont>/, (b) => fix(b, f.title || '')).replace(/<a:minorFont>[\s\S]*?<\/a:minorFont>/, (b) => fix(b, f.body || ''));
      zip.file(tp, t);
    }
    // notes text with Devanagari
    for (const np of Object.keys(zip.files).filter(f => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(f))) { const t = await zip.file(np).async('string'); const t2 = X.fixDevanagari(t); if (t2 !== t) zip.file(np, t2); }
    // document properties
    const pr = post.props || {};
    if (zip.file('docProps/core.xml')) {
      let c = await zip.file('docProps/core.xml').async('string');
      const by = BTS.userName || pr.author || 'BTS Slides';
      c = c.replace(/<cp:lastModifiedBy>[\s\S]*?<\/cp:lastModifiedBy>/, `<cp:lastModifiedBy>${xesc(by)}</cp:lastModifiedBy>`);
      const add = [['cp:keywords', pr.keywords], ['cp:category', pr.category], ['dc:description', pr.comments]].filter(([, v]) => v).map(([k, v]) => `<${k}>${xesc(v)}</${k}>`).join('');
      c = c.replace(/<(cp:keywords|cp:category|dc:description)>[\s\S]*?<\/\1>|<(cp:keywords|cp:category|dc:description)\/>/g, '');
      if (add) c = c.replace('</cp:coreProperties>', add + '</cp:coreProperties>');
      zip.file('docProps/core.xml', c);
    }
    if (zip.file('docProps/app.xml')) { let a = await zip.file('docProps/app.xml').async('string'); a = a.replace(/<Application>[\s\S]*?<\/Application>/, '<Application>BTS Slides (BTS Office)</Application>'); zip.file('docProps/app.xml', a); }

    // comments (legacy comments with p15 threading so replies stay attached)
    const authors = []; const authorIdx = (name) => { let a = authors.find(x => x.name === name); if (!a) { a = { id: authors.length, name, last: 0 }; authors.push(a); } return a; };
    for (let i = 0; i < n; i++) {
      const list = (post.comments || [])[i] || []; if (!list.length) continue;
      let cm = '';
      const one = (c, x, y, parent) => {
        const a = authorIdx(c.author || 'Reviewer'); const idx = ++a.last;
        const ext = parent ? `<p:extLst><p:ext uri="{C676402C-5697-4E1C-873F-D02D1690AC5C}"><p15:threadingInfo xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" timeZoneBias="0"><p15:parentCm authorId="${parent.a}" idx="${parent.idx}"/></p15:threadingInfo></p:ext></p:extLst>` : '';
        cm += `<p:cm authorId="${a.id}" dt="${new Date(c.time || Date.now()).toISOString().replace('Z', '')}" idx="${idx}"><p:pos x="${Math.round((x || 0) * 6)}" y="${Math.round((y || 0) * 6)}"/><p:text>${xesc((c.resolved && !parent ? '[Resolved] ' : '') + c.text)}</p:text>${ext}</p:cm>`;
        return { a: a.id, idx };
      };
      list.forEach(c => { const el = c.elId && pres.slides[i].elements.find(e => e.id === c.elId); const x = el ? el.x + el.w - 10 : c.x, y = el ? el.y : c.y; const p0 = one(c, x, y, null); (c.replies || []).forEach(r => one(r, x, y, p0)); });
      const part = `ppt/comments/comment${i + 1}.xml`;
      zip.file(part, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:cmLst ${NS_P}>${cm}</p:cmLst>`);
      await addRel(zip, `ppt/slides/_rels/slide${i + 1}.xml.rels`, 'comments', `../comments/comment${i + 1}.xml`);
      await addOverride(zip, '/' + part, 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml');
    }
    if (authors.length) {
      zip.file('ppt/commentAuthors.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:cmAuthorLst ${NS_P}>${authors.map(a => `<p:cmAuthor id="${a.id}" name="${xesc(a.name)}" initials="${xesc(a.name.split(/\s+/).map(w => w[0]).join('').slice(0, 3).toUpperCase())}" lastIdx="${a.last}" clrIdx="${a.id % 8}"/>`).join('')}</p:cmAuthorLst>`);
      await addRel(zip, 'ppt/_rels/presentation.xml.rels', 'commentAuthors', 'commentAuthors.xml');
      await addOverride(zip, '/ppt/commentAuthors.xml', 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml');
    }

    // custom shows + slide show settings
    let px = await zip.file('ppt/presentation.xml').async('string');
    const sldRids = [...px.matchAll(/<p:sldId\b[^>]*r:id="(rId\d+)"/g)].map(m => m[1]);
    const shows = (pres.customShows || []).map((c, k) => ({ c, k, rids: c.slides.map(id => pres.slides.findIndex(s => s.id === id)).filter(i => i >= 0).map(i => sldRids[i]).filter(Boolean) }));
    if (shows.length) {
      const lst = `<p:custShowLst>${shows.map(s => `<p:custShow name="${xesc(s.c.name)}" id="${s.k}"><p:sldLst>${s.rids.map(r => `<p:sld r:id="${r}"/>`).join('')}</p:sldLst></p:custShow>`).join('')}</p:custShowLst>`;
      px = px.replace(/<p:custShowLst>[\s\S]*?<\/p:custShowLst>/, '');
      if (/<p:notesSz\b[^>]*\/>/.test(px)) px = px.replace(/(<p:notesSz\b[^>]*\/>)/, '$1' + lst); else px = px.replace(/(<p:sldSz\b[^>]*\/>)/, '$1' + lst);
      zip.file('ppt/presentation.xml', px);
    }
    const st = pres.showSetup;
    if (st && zip.file('ppt/presProps.xml')) {
      let pp = await zip.file('ppt/presProps.xml').async('string');
      const ci = st.custom ? (pres.customShows || []).findIndex(c => c.id === st.custom) : -1;
      const sp = `<p:showPr${st.loop || st.kiosk ? ' loop="1"' : ''}${st.useTimings === false ? ' useTimings="0"' : ''}${st.noAnims ? ' showAnimation="0"' : ''} showNarration="1">${st.kiosk ? '<p:kiosk/>' : st.type === 'window' ? '<p:browse/>' : '<p:present/>'}${ci >= 0 ? `<p:custShow id="${ci}"/>` : st.range ? `<p:sldRg st="${st.range.from}" end="${st.range.to}"/>` : '<p:sldAll/>'}<p:penClr><a:prstClr val="red"/></p:penClr></p:showPr>`;
      pp = pp.replace(/<p:showPr[\s\S]*?<\/p:showPr>/, '');
      if (/<p:presentationPr\b[^>]*\/>/.test(pp)) pp = pp.replace(/<p:presentationPr\b([^>]*?)\s*\/>/, (m, at) => `<p:presentationPr${at}>${sp}</p:presentationPr>`);
      else pp = pp.replace(/(<p:presentationPr\b[^>]*>)/, '$1' + sp);
      if (!/xmlns:a=/.test(pp.slice(0, pp.indexOf('>', pp.indexOf('<p:presentationPr'))))) pp = pp.replace('<p:presentationPr ', '<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ');
      zip.file('ppt/presProps.xml', pp);
    }

    // BTS metadata (customXml): exact copy of each slide + presentation-level settings
    const meta = { app: 'BTS Slides', v: 2, theme: SL.clone(pres.theme), master: pres.master ? SL.clone(pres.master) : null, customShows: (pres.customShows || []).map(c => ({ name: c.name, id: c.id, slides: c.slides.map(id => pres.slides.findIndex(s => s.id === id)).filter(i => i >= 0) })), showSetup: pres.showSetup || null, footer: pres.footer || null, slides: [] };
    if (meta.master) ['graphics'].forEach(k => (meta.master[k] || []).forEach(el => { if (el.type === 'image') { const m = post.masterNames[el.id]; if (m) { el._n = m.name; el._s = m.slide; } delete el.src; } }));
    if (meta.master && meta.master.layoutGraphics) Object.values(meta.master.layoutGraphics).forEach(l => l.forEach(el => { if (el.type === 'image') { const m = post.masterNames[el.id]; if (m) { el._n = m.name; el._s = m.slide; } delete el.src; } }));
    for (let i = 0; i < n && i < pres.slides.length; i++) {
      const s = SL.clone(pres.slides[i]); const names = post.names[i] || {};
      [...s.elements, ...(s.layer || [])].forEach(el => { if (el.type === 'image') { if (el.src) el._n = names[el.id] || null; delete el.src; } });
      if (s.background && s.background.type === 'image') { delete s.background.src; s.background._img = true; }
      meta.slides.push({ h: await X.slideHash(zip, i + 1), s });
    }
    const json = JSON.stringify(meta);
    zip.file('customXml/btsSlides.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><bts:slides xmlns:bts="urn:bts-office:slides" version="2">${xesc(json)}</bts:slides>`);
    await addRel(zip, 'ppt/_rels/presentation.xml.rels', 'customXml', '../customXml/btsSlides.xml');
    return await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  };
  // A slide's fingerprint: slide XML + its notes and comments. Used on reopen to tell whether the slide was edited elsewhere.
  X.slideHash = async (zip, num) => {
    const sp = `ppt/slides/slide${num}.xml`; const f = zip.file(sp); if (!f) return '';
    let str = await f.async('string');
    const rf = zip.file(`ppt/slides/_rels/slide${num}.xml.rels`);
    if (rf) { const rels = await rf.async('string'); for (const m of rels.matchAll(/Target="\.\.\/(notesSlides\/notesSlide\d+\.xml|comments\/comment\d+\.xml)"/g)) { const g = zip.file('ppt/' + m[1]); if (g) str += '\n' + await g.async('string'); } }
    return X.hash(str);
  };
})();
