/* BTS Slides — PowerPoint (.pptx) export via PptxGenJS, plus post-processing with JSZip for
 * features PptxGenJS lacks (slide transitions, gradient backgrounds/fills, hidden slides).
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

  /* ---------- main ---------- */
  X.build = async (pres = P()) => {
    if (!window.PptxGenJS) throw new Error('PptxGenJS is not loaded');
    const pptx = new window.PptxGenJS();
    const W = SL.W(pres), H = SL.H(pres);
    pptx.defineLayout({ name: 'BTS', width: IN(W), height: IN(H) }); pptx.layout = 'BTS';
    pptx.author = 'BTS Slides'; pptx.company = 'BTS Office'; pptx.title = BTS.title || 'Presentation';
    try { pptx.theme = { headFontFace: SL.fontOf('title', pres), bodyFontFace: SL.fontOf('body', pres) }; } catch { }
    const post = { trans: {}, grads: {}, bgGrads: {}, hidden: {} };
    const decoCache = {};
    const secs = pres.sections || []; const usedSec = {};
    for (let i = 0; i < pres.slides.length; i++) {
      const s = pres.slides[i];
      const sec = secs.find(x => x.start === s.id);
      if (sec) { let t = sec.name || 'Section'; while (usedSec[t]) t += ' '; usedSec[t] = 1; pptx.addSection({ title: t }); }
      const slide = Object.keys(usedSec).length ? pptx.addSlide({ sectionTitle: Object.keys(usedSec).pop() }) : pptx.addSlide();
      // background
      const bg = s.background || SL.themeBg(pres);
      if (bg.type === 'image' && bg.src) { const d = await bgImage(bg, W, H); if (d) slide.background = { data: d }; }
      else if (bg.type === 'gradient') { slide.background = { color: hx(bg.color, pres) || 'FFFFFF' }; post.bgGrads[i] = { c1: hx(bg.color, pres), c2: hx(bg.color2, pres), angle: bg.angle ?? 135 }; }
      else slide.background = { color: hx(bg.color || 'bg', pres) || 'FFFFFF' };
      if (s.hidden) { slide.hidden = true; post.hidden[i] = true; }
      if (s.transition && s.transition.type && s.transition.type !== 'none') post.trans[i] = s.transition;
      // theme graphics as a picture at the back (re-created from the theme on import)
      const kind = SL.decoKind(s.layout);
      const tag = `bts:theme=${pres.theme.id};v=${pres.theme.variant || 0};layout=${s.layout}`;
      if (!s.hideGraphics) {
        const key = kind; if (!decoCache[key]) { const svg = SL.decoSVG(Object.assign({}, s, { hideGraphics: false }), pres); decoCache[key] = svg ? await svgToPNG(svg.replace('<svg ', '<svg xmlns="http://www.w3.org/2000/svg" '), W, H, 1) : null; }
        if (decoCache[key]) slide.addImage({ data: decoCache[key], x: 0, y: 0, w: IN(W), h: IN(H), altText: tag, objectName: 'Theme graphics' });
      } else slide.addText('', { x: 0, y: 0, w: .01, h: .01, objectName: 'Theme graphics', altText: tag });
      // imported layout/master graphics
      for (const el of s.layer || []) await addEl(slide, el, pres, post, i, true);
      for (const el of s.elements) await addEl(slide, el, pres, post, i);
      for (const el of SL.footerEls(s, pres, i)) await addEl(slide, el, pres, post, i);
      if (s.notes) slide.addNotes(s.notes);
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
      if (el.dir === 'vert') o.vert = 'vert'; else if (el.dir === 'vert270') o.vert = 'vert270';
      if (el.lh && el.lh !== 1) o.lineSpacingMultiple = el.lh;
      if (!o.fill) delete o.fill; if (!o.line) delete o.line; if (!o.shadow) delete o.shadow;
      slide.addText(runs, o);
    } else if (el.type === 'image') {
      const data = await processImage(el); if (!data) return;
      const o = Object.assign(g, { data, objectName: name, altText: el.alt || '' });
      if (el.opacity != null && el.opacity < 1) o.transparency = Math.round((1 - el.opacity) * 100);
      slide.addImage(o);
    } else if (el.type === 'icon') {
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${el.w}" height="${el.h}" viewBox="0 0 24 24" preserveAspectRatio="none" fill="${el.fillIcon ? SL.color(el.color || 'accent1', pres) : 'none'}" stroke="${SL.color(el.color || 'accent1', pres)}" stroke-width="${el.strokeWidth || 1.75}" stroke-linecap="round" stroke-linejoin="round">${(window.BTS_ICONS || {})[el.name] || ''}</svg>`;
      const data = await svgToPNG(svg, el.w, el.h, 4);
      const o = Object.assign(g, { data, objectName: 'Icon ' + el.name + ' O' + objN, altText: el.alt || ('Icon: ' + el.name) });
      if (el.opacity != null && el.opacity < 1) o.transparency = Math.round((1 - el.opacity) * 100);
      slide.addImage(o);
    } else if (el.type === 'table') {
      const acc = hx(el.accent || 'accent1', pres) || '4472C4';
      const mixHex = (h, a) => { const r = C2(h).map(v => Math.round(v * a + 255 * (1 - a))); return r.map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase(); };
      function C2(h) { return [0, 2, 4].map(i => parseInt(h.slice(i, i + 2), 16)); }
      const rows = el.rows.map((row, r) => {
        const out = [];
        row.forEach((c) => {
          if (c.merged) return;
          const isHead = el.header && r === 0;
          let fill = c.fill ? hx(c.fill, pres) : isHead ? acc : el.banded ? ((el.header ? r - 1 : r) % 2 === 0 ? mixHex(acc, .14) : mixHex(acc, .06)) : 'FFFFFF';
          const color = c.color ? hx(c.color, pres) : isHead ? 'FFFFFF' : (el.color ? hx(el.color, pres) : '1F2430');
          const cellEl = { html: c.html || '', size: el.size || 16, font: el.font || SL.fontOf('body', pres), color: '#' + color, bold: isHead || c.bold, align: c.align || 'left', role: 'body', type: 'text' };
          const runs = X.runs(cellEl, pres);
          const opt = { fill: fill ? { color: fill } : undefined, valign: c.valign || 'middle', align: c.align || 'left', border: { type: 'solid', pt: 1, color: 'FFFFFF' }, margin: [4, 6, 4, 6] };
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
      const pal = SL.palette(pres).accents.map(c => hx(c, pres));
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
  X.postProcess = async (data, post) => {
    const zip = await window.JSZip.loadAsync(data);
    const n = Object.keys(zip.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f)).length;
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
      if (xml !== orig) zip.file(path, xml);
    }
    return await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  };
})();
