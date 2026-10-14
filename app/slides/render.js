/* BTS Slides — renderer. Turns model slides into DOM (used by the editor canvas, thumbnails,
 * slide show, PDF export and template previews). Geometry is in logical px (96/in), fonts in pt.
 */
(function () {
  const SL = window.SL = window.SL || {};
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const f2 = (n) => Math.round(n * 100) / 100;

  /* ---------- shape geometry (paths in px for a w×h box) ---------- */
  SL.SHAPES = [
    ['rect', 'Rectangle'], ['roundRect', 'Rounded rectangle'], ['ellipse', 'Oval'], ['triangle', 'Triangle'], ['rtTriangle', 'Right triangle'],
    ['diamond', 'Diamond'], ['parallelogram', 'Parallelogram'], ['trapezoid', 'Trapezoid'], ['pentagon', 'Pentagon'], ['hexagon', 'Hexagon'],
    ['octagon', 'Octagon'], ['star4', '4-point star'], ['star5', '5-point star'], ['star6', '6-point star'], ['plus', 'Plus'], ['donut', 'Donut'],
    ['heart', 'Heart'], ['cloud', 'Cloud'], ['can', 'Cylinder'], ['rightArrow', 'Right arrow'], ['leftArrow', 'Left arrow'], ['upArrow', 'Up arrow'],
    ['downArrow', 'Down arrow'], ['leftRightArrow', 'Left-right arrow'], ['chevron', 'Chevron'], ['homePlate', 'Pentagon arrow'],
    ['wedgeRectCallout', 'Rectangle callout'], ['wedgeRoundRectCallout', 'Rounded callout'], ['wedgeEllipseCallout', 'Oval callout'],
    ['line', 'Line'], ['arrowLine', 'Arrow'],
  ];
  SL.isLine = (el) => el.type === 'shape' && (el.kind === 'line' || el.kind === 'arrowLine');

  function starPts(n, w, h, inner) {
    const pts = []; const cx = w / 2, cy = h / 2;
    for (let i = 0; i < n * 2; i++) {
      const a = -Math.PI / 2 + i * Math.PI / n; const r = i % 2 ? inner : 1;
      pts.push([cx + Math.cos(a) * cx * r, cy + Math.sin(a) * cy * r]);
    }
    return 'M' + pts.map(p => f2(p[0]) + ' ' + f2(p[1])).join(' L') + ' Z';
  }
  const poly = (pts) => 'M' + pts.map(p => f2(p[0]) + ' ' + f2(p[1])).join(' L') + ' Z';
  function rrect(x, y, w, h, r) {
    r = Math.max(0, Math.min(r, w / 2, h / 2));
    return `M${f2(x + r)} ${f2(y)} H${f2(x + w - r)} A${f2(r)} ${f2(r)} 0 0 1 ${f2(x + w)} ${f2(y + r)} V${f2(y + h - r)} A${f2(r)} ${f2(r)} 0 0 1 ${f2(x + w - r)} ${f2(y + h)} H${f2(x + r)} A${f2(r)} ${f2(r)} 0 0 1 ${f2(x)} ${f2(y + h - r)} V${f2(y + r)} A${f2(r)} ${f2(r)} 0 0 1 ${f2(x + r)} ${f2(y)} Z`;
  }
  const ell = (cx, cy, rx, ry) => `M${f2(cx - rx)} ${f2(cy)} A${f2(rx)} ${f2(ry)} 0 1 0 ${f2(cx + rx)} ${f2(cy)} A${f2(rx)} ${f2(ry)} 0 1 0 ${f2(cx - rx)} ${f2(cy)} Z`;

  // → { d, d2 (stroke-only extra), inset:[l,t,r,b] fractions for text, evenodd }
  SL.geom = (kind, w, h, el = {}) => {
    const m = Math.min(w, h);
    switch (kind) {
      case 'roundRect': return { d: rrect(0, 0, w, h, m * (el.adj ?? .1667)), inset: [.05, .05, .05, .05] };
      case 'ellipse': return { d: ell(w / 2, h / 2, w / 2, h / 2), inset: [.146, .146, .146, .146] };
      case 'triangle': return { d: poly([[w / 2, 0], [w, h], [0, h]]), inset: [.25, .5, .25, 0] };
      case 'rtTriangle': return { d: poly([[0, 0], [w, h], [0, h]]), inset: [.08, .5, .5, .08] };
      case 'diamond': return { d: poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]), inset: [.25, .25, .25, .25] };
      case 'parallelogram': { const o = m * .25; return { d: poly([[o, 0], [w, 0], [w - o, h], [0, h]]), inset: [o / w, 0, o / w, 0] }; }
      case 'trapezoid': { const o = m * .25; return { d: poly([[o, 0], [w - o, 0], [w, h], [0, h]]), inset: [o / w, 0, o / w, 0] }; }
      case 'pentagon': return { d: poly([[w / 2, 0], [w, h * .382], [w * .809, h], [w * .191, h], [0, h * .382]]), inset: [.19, .3, .19, 0] };
      case 'hexagon': { const o = m * .25; return { d: poly([[o, 0], [w - o, 0], [w, h / 2], [w - o, h], [o, h], [0, h / 2]]), inset: [o / w, .1, o / w, .1] }; }
      case 'octagon': { const o = m * .2929; return { d: poly([[o, 0], [w - o, 0], [w, o], [w, h - o], [w - o, h], [o, h], [0, h - o], [0, o]]), inset: [.15, .15, .15, .15] }; }
      case 'star4': return { d: starPts(4, w, h, .38), inset: [.3, .3, .3, .3] };
      case 'star5': return { d: starPts(5, w, h, .45), inset: [.3, .35, .3, .2] };
      case 'star6': return { d: starPts(6, w, h, .58), inset: [.25, .25, .25, .25] };
      case 'plus': { const t = m * .25; return { d: poly([[t, 0], [w - t, 0], [w - t, t], [w, t], [w, h - t], [w - t, h - t], [w - t, h], [t, h], [t, h - t], [0, h - t], [0, t], [t, t]]), inset: [t / w, t / h, t / w, t / h] }; }
      case 'donut': { const t = m * .25; return { d: ell(w / 2, h / 2, w / 2, h / 2) + ' ' + ell(w / 2, h / 2, w / 2 - t, h / 2 - t), evenodd: true, inset: [.2, .2, .2, .2] }; }
      case 'heart': { const sx = w / 100, sy = h / 100, P = (x, y) => f2(x * sx) + ' ' + f2(y * sy); return { d: `M${P(50, 96)} L${P(10, 55)} C${P(-6, 38)} ${P(2, 8)} ${P(26, 6)} C${P(38, 5)} ${P(46, 12)} ${P(50, 22)} C${P(54, 12)} ${P(62, 5)} ${P(74, 6)} C${P(98, 8)} ${P(106, 38)} ${P(90, 55)} Z`, inset: [.2, .15, .2, .3] }; }
      case 'cloud': {
        const n = 9, cx = w / 2, cy = h / 2, rx = w * .42, ry = h * .38; const pts = [];
        for (let i = 0; i < n; i++) { const a = -Math.PI / 2 + i * 2 * Math.PI / n + .2; pts.push([cx + Math.cos(a) * rx, cy + Math.sin(a) * ry]); }
        let d = `M${f2(pts[0][0])} ${f2(pts[0][1])}`;
        for (let i = 0; i < n; i++) { const a = pts[i], b = pts[(i + 1) % n]; const ch = Math.hypot(b[0] - a[0], b[1] - a[1]); const r = ch * .62; d += ` A${f2(r)} ${f2(r)} 0 0 1 ${f2(b[0])} ${f2(b[1])}`; }
        return { d: d + ' Z', inset: [.2, .22, .2, .22] };
      }
      case 'can': { const ry = Math.min(h * .25, m * .125); return { d: `M0 ${f2(ry)} A${f2(w / 2)} ${f2(ry)} 0 0 1 ${f2(w)} ${f2(ry)} V${f2(h - ry)} A${f2(w / 2)} ${f2(ry)} 0 0 1 0 ${f2(h - ry)} Z`, d2: `M0 ${f2(ry)} A${f2(w / 2)} ${f2(ry)} 0 0 0 ${f2(w)} ${f2(ry)}`, inset: [0, 2 * ry / h, 0, ry / h] }; }
      case 'rightArrow': { const hx = w - Math.min(w, h * .5 * 2) * .5; return { d: poly([[0, h * .25], [hx, h * .25], [hx, 0], [w, h / 2], [hx, h], [hx, h * .75], [0, h * .75]]), inset: [0, .25, (w - hx) / w * .5, .25] }; }
      case 'leftArrow': { const hx = Math.min(w, h) * .5; return { d: poly([[w, h * .25], [hx, h * .25], [hx, 0], [0, h / 2], [hx, h], [hx, h * .75], [w, h * .75]]), inset: [hx / w * .5, .25, 0, .25] }; }
      case 'upArrow': { const hy = Math.min(w, h) * .5; return { d: poly([[w * .25, h], [w * .25, hy], [0, hy], [w / 2, 0], [w, hy], [w * .75, hy], [w * .75, h]]), inset: [.25, hy / h * .5, .25, 0] }; }
      case 'downArrow': { const hy = h - Math.min(w, h) * .5; return { d: poly([[w * .25, 0], [w * .75, 0], [w * .75, hy], [w, hy], [w / 2, h], [0, hy], [w * .25, hy]]), inset: [.25, 0, .25, (h - hy) / h * .5] }; }
      case 'leftRightArrow': { const hx = Math.min(w / 2, h * .5); return { d: poly([[0, h / 2], [hx, 0], [hx, h * .25], [w - hx, h * .25], [w - hx, 0], [w, h / 2], [w - hx, h], [w - hx, h * .75], [hx, h * .75], [hx, h]]), inset: [hx / w * .5, .25, hx / w * .5, .25] }; }
      case 'chevron': { const o = Math.min(w / 2, h * .5); return { d: poly([[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h], [o, h / 2]]), inset: [o / w, 0, o / w, 0] }; }
      case 'homePlate': { const o = Math.min(w / 2, h * .5); return { d: poly([[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h]]), inset: [0, 0, o / w * .7, 0] }; }
      case 'wedgeRectCallout': return { d: `M0 0 H${f2(w)} V${f2(h)} H${f2(w * .42)} L${f2(w * .2)} ${f2(h * 1.28)} L${f2(w * .17)} ${f2(h)} H0 Z`, inset: [0, 0, 0, 0] };
      case 'wedgeRoundRectCallout': { const r = m * .1667; return { d: `M${f2(r)} 0 H${f2(w - r)} A${f2(r)} ${f2(r)} 0 0 1 ${f2(w)} ${f2(r)} V${f2(h - r)} A${f2(r)} ${f2(r)} 0 0 1 ${f2(w - r)} ${f2(h)} H${f2(w * .42)} L${f2(w * .2)} ${f2(h * 1.28)} L${f2(w * .17)} ${f2(h)} H${f2(r)} A${f2(r)} ${f2(r)} 0 0 1 0 ${f2(h - r)} V${f2(r)} A${f2(r)} ${f2(r)} 0 0 1 ${f2(r)} 0 Z`, inset: [.03, .03, .03, .03] }; }
      case 'wedgeEllipseCallout': {
        const cx = w / 2, cy = h / 2, P = (deg) => { const a = deg * Math.PI / 180; return f2(cx + Math.cos(a) * cx) + ' ' + f2(cy + Math.sin(a) * cy); };
        return { d: `M${P(132)} A${f2(cx)} ${f2(cy)} 0 1 1 ${P(105)} L${f2(w * .18)} ${f2(h * 1.25)} Z`, inset: [.146, .146, .146, .146] };
      }
      case 'line': case 'arrowLine': return { d: `M0 0 L${f2(w)} ${f2(h)}`, line: true, inset: [0, 0, 0, 0] };
      case 'custom': if (el.path) { const sx = w / (el.path.w || 1), sy = h / (el.path.h || 1); return { d: el.path.d.replace(/(-?\d*\.?\d+(?:e-?\d+)?),(-?\d*\.?\d+(?:e-?\d+)?)/g, (_, a, b) => f2(a * sx) + ' ' + f2(b * sy)), inset: [0, 0, 0, 0], noFillPaths: el.path.noFill } } // falls through
      default: return { d: `M0 0 H${f2(w)} V${f2(h)} H0 Z`, inset: [0, 0, 0, 0] };
    }
  };

  // Make SVG ids unique per render: the same slide is drawn many times (rail, gallery, canvas) and
  // url(#id) resolves to the first match in the document, which may be inside a hidden panel.
  let uq = 0;
  SL.uniqIds = (svg) => { if (!/ id="/.test(svg)) return svg; const n = (++uq).toString(36); return svg.replace(/ id="([^"]+)"/g, ` id="$1-${n}"`).replace(/url\(#([^)]+)\)/g, `url(#$1-${n})`); };
  const DASH = { solid: '', dash: '4 3', lgDash: '8 3', dot: '1 2', dashDot: '4 2 1 2', sysDash: '3 1' };
  SL.DASHES = [['solid', 'Solid'], ['dash', 'Dash'], ['lgDash', 'Long dash'], ['dot', 'Dot'], ['dashDot', 'Dash dot']];

  SL.shapeSVG = (el, pres, w = el.w, h = el.h) => {
    const g = SL.geom(el.type === 'text' ? 'rect' : (el.kind || 'rect'), w, h, el);
    const fill = g.line ? 'none' : (SL.color(el.fill, pres) || 'none');
    const stroke = SL.color(el.stroke, pres) || 'none';
    const sw = stroke === 'none' ? 0 : (el.strokeWidth ?? 1) * 4 / 3; // pt → px
    if (fill === 'none' && stroke === 'none' && !g.line) return '';
    const dash = DASH[el.dash] ? `stroke-dasharray="${DASH[el.dash].split(' ').map(x => f2(x * Math.max(sw, 1))).join(' ')}"` : '';
    const mid = el.kind === 'arrowLine' ? `marker-end="url(#ah-${el.id})"` : '';
    const lineStroke = g.line && stroke === 'none' ? (SL.color('text', pres)) : stroke;
    const lsw = g.line && !sw ? 2 : sw;
    let defs = '';
    if (el.kind === 'arrowLine') defs = `<defs><marker id="ah-${el.id}" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="${f2(Math.max(3, 9 / Math.max(1, lsw / 2.5)))}" markerHeight="${f2(Math.max(3, 9 / Math.max(1, lsw / 2.5)))}" orient="auto-start-reverse" markerUnits="strokeWidth"><path d="M0 0 L10 5 L0 10 z" fill="${lineStroke}"/></marker></defs>`;
    let fillAttr = fill;
    if (el.fill && typeof el.fill === 'object' && el.fill.type === 'gradient') {
      defs += `<defs><linearGradient id="gf-${el.id}" gradientTransform="rotate(${el.fill.angle || 90} .5 .5)"><stop offset="0" stop-color="${SL.color(el.fill.color, pres)}"/><stop offset="1" stop-color="${SL.color(el.fill.color2, pres)}"/></linearGradient></defs>`;
      fillAttr = `url(#gf-${el.id})`;
    }
    const hit = g.line ? `<path d="${g.d}" stroke="transparent" stroke-width="14" fill="none" class="hit"/>` : '';
    // a zero-height/width viewBox disables rendering (horizontal/vertical lines) → size ≥1 and no viewBox (px coords)
    return SL.uniqIds(`<svg class="geo" width="${f2(Math.max(1, w))}" height="${f2(Math.max(1, h))}" overflow="visible">${defs}<path d="${g.d}" fill="${fillAttr}" ${g.evenodd ? 'fill-rule="evenodd"' : ''} stroke="${g.line ? lineStroke : stroke}" stroke-width="${f2(g.line ? lsw : sw)}" ${dash} ${mid} stroke-linejoin="round"/>${g.d2 && stroke !== 'none' ? `<path d="${g.d2}" fill="none" stroke="${stroke}" stroke-width="${f2(sw)}"/>` : g.d2 ? `<path d="${g.d2}" fill="none" stroke="rgba(0,0,0,.18)" stroke-width="1.2"/>` : ''}${hit}</svg>`);
  };

  /* ---------- background ---------- */
  SL.bgCSS = (bg, pres) => {
    if (!bg) bg = SL.themeBg(pres);
    if (bg.type === 'gradient') return `linear-gradient(${bg.angle ?? 135}deg, ${SL.color(bg.color, pres)}, ${SL.color(bg.color2, pres)})`;
    if (bg.type === 'image') return `${SL.color(bg.color || '#ffffff', pres)} url("${SL.assetURL(bg.src)}") center/${bg.fit === 'contain' ? 'contain' : bg.fit === 'stretch' ? '100% 100%' : 'cover'} no-repeat`;
    return SL.color(bg.color || 'bg', pres);
  };
  SL.decoSVG = (slide, pres) => {
    if (slide.hideGraphics) return '';
    const th = SL.themeOf(pres); if (!th.deco) return '';
    const W = SL.W(pres), H = SL.H(pres);
    return SL.uniqIds(`<svg class="deco" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">${th.deco(SL.decoKind(slide.layout), W, H, SL.palette(pres))}</svg>`);
  };

  /* ---------- text styling ---------- */
  SL.EFFECTS = [
    ['none', 'No effect'], ['shadow', 'Soft shadow'], ['outline', 'Outline'], ['glow', 'Glow'], ['gradient', 'Gradient fill'],
    ['lifted', 'Lifted 3-D'], ['neon', 'Neon'], ['retro', 'Retro offset'],
  ];
  SL.effectCSS = (eff, color, pres) => {
    const c = color, a1 = SL.color('accent1', pres), a2 = SL.color('accent2', pres);
    switch (eff) {
      case 'shadow': return 'text-shadow:2px 3px 6px rgba(0,0,0,.35);';
      case 'outline': return `-webkit-text-stroke:1.5px ${c};color:transparent !important;`;
      case 'glow': return `text-shadow:0 0 8px ${a1},0 0 16px ${a1};`;
      case 'gradient': return `background:linear-gradient(90deg,${a1},${a2});-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;`;
      case 'lifted': return 'text-shadow:1px 1px 0 rgba(0,0,0,.25),2px 2px 0 rgba(0,0,0,.2),3px 3px 0 rgba(0,0,0,.15),4px 6px 10px rgba(0,0,0,.3);';
      case 'neon': return `text-shadow:0 0 4px #fff,0 0 10px ${a1},0 0 22px ${a1},0 0 40px ${a1};`;
      case 'retro': return `text-shadow:4px 4px 0 ${a2};`;
      default: return '';
    }
  };
  SL.textDefaults = (el, pres) => {
    const role = el.role || 'body';
    const pal = SL.palette(pres);
    let color = el.color ? SL.color(el.color, pres) : null;
    if (!color) {
      if (el.type === 'shape') color = '#ffffff';
      else color = role === 'title' ? pal.title : pal.text;
    }
    return { font: el.font || SL.fontOf(role === 'title' ? 'title' : 'body', pres), size: el.size || 18, color, bold: el.bold ?? (role === 'title' ? SL.themeOf(pres).titleBold : false) };
  };
  const SERIF = /serif|times|georgia|cambria|garamond|book|roman|lora|playfair|caladea|bodoni|palatino|minion|baskerville|didot|merriweather|charter|mincho|song/i;
  const MONO = /mono|courier|consol|code/i;
  SL.fontStack = (f) => { f = String(f || 'Inter').replace(/['"]/g, ''); const gen = MONO.test(f) ? 'monospace' : (SERIF.test(f) && !/sans/i.test(f)) ? 'serif' : 'sans-serif'; return `'${f}', ${gen}`; };
  SL.textStyle = (el, pres) => {
    const d = SL.textDefaults(el, pres);
    let s = `font-family:${SL.fontStack(d.font)};font-size:${el.size || 18}pt;color:${d.color};line-height:${f2(1.2 * (el.lh || 1))};text-align:${el.align || 'left'};`;
    if (d.bold) s += 'font-weight:700;';
    if (el.italic) s += 'font-style:italic;';
    const deco = [el.underline && 'underline', el.strike && 'line-through'].filter(Boolean).join(' ');
    if (deco) s += `text-decoration:${deco};`;
    if (el.highlight) s += `--hl:${el.highlight};`;
    if (el.caps) s += 'text-transform:uppercase;';
    if (el.spacing) s += `letter-spacing:${el.spacing}pt;`;
    if (el.pspace != null) s += `--psp:${el.pspace}pt;`;
    if (el.effect) s += SL.effectCSS(el.effect, d.color, pres);
    return s;
  };

  /* ---------- charts (SVG) ---------- */
  SL.chartSVG = (ch, w, h, pres) => {
    if (!ch) return '';
    const pal = SL.palette(pres); const cols = pal.accents; const txt = pal.text;
    const font = SL.fontOf('body', pres);
    const cats = ch.categories || [], ser = ch.series || [];
    let out = `<svg class="chart" width="${f2(w)}" height="${f2(h)}" viewBox="0 0 ${f2(w)} ${f2(h)}" font-family="${esc(font)}" fill="${txt}">`;
    let top = 10; const fs = Math.max(10, Math.min(16, h / 22));
    if (ch.title) { out += `<text x="${w / 2}" y="${fs * 1.6}" text-anchor="middle" font-size="${fs * 1.3}" font-weight="700">${esc(ch.title)}</text>`; top = fs * 2.6; }
    const legendH = ch.legend === false ? 0 : fs * 2;
    const type = ch.type || 'bar';
    if (type === 'pie' || type === 'doughnut') {
      const vals = (ser[0]?.values || []).map(v => Math.max(0, +v || 0)); const tot = vals.reduce((a, b) => a + b, 0) || 1;
      const cx = w / 2, cy = top + (h - top - legendH) / 2, r = Math.max(5, Math.min(w, h - top - legendH) / 2 - 8);
      let a = -Math.PI / 2;
      vals.forEach((v, i) => {
        const da = v / tot * Math.PI * 2; const a2 = a + da;
        const x1 = cx + r * Math.cos(a), y1 = cy + r * Math.sin(a), x2 = cx + r * Math.cos(a2), y2 = cy + r * Math.sin(a2);
        const c = cols[i % cols.length];
        out += da >= Math.PI * 2 - 1e-6 ? `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${c}"/>` : `<path d="M${f2(cx)} ${f2(cy)} L${f2(x1)} ${f2(y1)} A${f2(r)} ${f2(r)} 0 ${da > Math.PI ? 1 : 0} 1 ${f2(x2)} ${f2(y2)} Z" fill="${c}" stroke="#fff" stroke-width="1.5"/>`;
        const am = a + da / 2; if (da > .25) out += `<text x="${f2(cx + r * .68 * Math.cos(am))}" y="${f2(cy + r * .68 * Math.sin(am) + fs * .35)}" text-anchor="middle" font-size="${fs}" fill="#fff" font-weight="700">${Math.round(v / tot * 100)}%</text>`;
        a = a2;
      });
      if (type === 'doughnut') out += `<circle cx="${cx}" cy="${cy}" r="${r * .55}" fill="${SL.color('bg', pres)}"/>`;
      if (legendH) out += legend(cats, cols, w, h - legendH * .45, fs);
      return out + '</svg>';
    }
    const all = ser.flatMap(s => s.values.map(v => +v || 0));
    let max = Math.max(0, ...all), min = Math.min(0, ...all); if (max === min) max = min + 1;
    const step = niceStep((max - min) / 5); max = Math.ceil(max / step) * step; min = Math.floor(min / step) * step;
    const L = fs * 3.2, R = 10, T = top + 4, B = fs * 2 + legendH;
    const pw = Math.max(10, w - L - R), ph = Math.max(10, h - T - B);
    const Y = (v) => T + ph - (v - min) / (max - min) * ph;
    for (let v = min; v <= max + 1e-9; v += step) {
      out += `<line x1="${L}" x2="${f2(L + pw)}" y1="${f2(Y(v))}" y2="${f2(Y(v))}" stroke="${txt}" stroke-opacity=".15"/>`;
      out += `<text x="${L - 6}" y="${f2(Y(v) + fs * .35)}" text-anchor="end" font-size="${fs}" fill-opacity=".8">${fmtNum(v)}</text>`;
    }
    const n = Math.max(1, cats.length), bw = pw / n;
    cats.forEach((c, i) => { out += `<text x="${f2(L + bw * (i + .5))}" y="${f2(T + ph + fs * 1.4)}" text-anchor="middle" font-size="${fs}" fill-opacity=".85">${esc(c)}</text>`; });
    if (type === 'line') {
      ser.forEach((s, si) => {
        const c = cols[si % cols.length];
        const pts = s.values.map((v, i) => [L + bw * (i + .5), Y(+v || 0)]);
        out += `<polyline points="${pts.map(p => f2(p[0]) + ',' + f2(p[1])).join(' ')}" fill="none" stroke="${c}" stroke-width="3" stroke-linejoin="round"/>`;
        pts.forEach(p => { out += `<circle cx="${f2(p[0])}" cy="${f2(p[1])}" r="4" fill="${c}" stroke="#fff" stroke-width="1.5"/>`; });
      });
    } else {
      const ns = Math.max(1, ser.length), gw = bw * .72, sw = gw / ns;
      ser.forEach((s, si) => s.values.forEach((v, i) => {
        v = +v || 0; const x = L + bw * i + (bw - gw) / 2 + sw * si; const y0 = Y(Math.max(0, v)), y1 = Y(Math.min(0, v));
        out += `<rect x="${f2(x + 1)}" y="${f2(y0)}" width="${f2(Math.max(1, sw - 2))}" height="${f2(Math.max(0, y1 - y0))}" fill="${cols[si % cols.length]}" rx="2"/>`;
      }));
    }
    out += `<line x1="${L}" x2="${f2(L + pw)}" y1="${f2(Y(0))}" y2="${f2(Y(0))}" stroke="${txt}" stroke-opacity=".5"/>`;
    if (legendH) out += legend(ser.map(s => s.name), cols, w, h - legendH * .45, fs);
    return out + '</svg>';
  };
  function legend(names, cols, w, y, fs) {
    const items = names.map((n, i) => ({ n: String(n ?? ''), c: cols[i % cols.length] }));
    const widths = items.map(it => fs * 1.4 + it.n.length * fs * .55 + fs);
    let x = (w - widths.reduce((a, b) => a + b, 0)) / 2; let s = '';
    items.forEach((it, i) => { s += `<rect x="${f2(x)}" y="${f2(y - fs * .75)}" width="${fs * .9}" height="${fs * .9}" rx="2" fill="${it.c}"/><text x="${f2(x + fs * 1.3)}" y="${f2(y + fs * .1)}" font-size="${fs}">${esc(it.n)}</text>`; x += widths[i]; });
    return s;
  }
  function niceStep(r) { const p = Math.pow(10, Math.floor(Math.log10(r || 1))); const m = r / p; return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 5 ? 5 : 10) * p; }
  function fmtNum(v) { return Math.abs(v) >= 1000 ? (v / 1000).toFixed(Math.abs(v) % 1000 ? 1 : 0) + 'k' : String(Math.round(v * 100) / 100); }

  /* ---------- table ---------- */
  SL.tableHTML = (el, pres, editing) => {
    const pal = SL.palette(pres); const acc = SL.color(el.accent || 'accent1', pres);
    const font = el.font || SL.fontOf('body', pres); const n = el.rows.length;
    const bodyColor = el.color ? SL.color(el.color, pres) : '#1f2430';
    let h = `<table class="tbl" style="font-family:${SL.fontStack(font)};font-size:${el.size || 16}pt;--acc:${acc};--band:${mix(acc, 0.14)};--band2:${mix(acc, 0.06)};color:${bodyColor}"><colgroup>${(el.colW || []).map(c => `<col style="width:${f2(c * 100)}%">`).join('')}</colgroup>`;
    el.rows.forEach((row, r) => {
      const isHead = el.header && r === 0;
      h += `<tr style="height:${f2(el.rowH?.[r] ? el.rowH[r] * el.h : el.h / n)}px">`;
      row.forEach((c, ci) => {
        if (c.merged) return;
        let bg = c.fill ? SL.color(c.fill, pres) : isHead ? acc : (el.banded ? ((el.header ? r - 1 : r) % 2 === 0 ? 'var(--band)' : 'var(--band2)') : '#ffffff');
        if (c.fill === 'none') bg = 'transparent';
        const col = c.color ? SL.color(c.color, pres) : isHead ? '#ffffff' : '';
        const st = `background:${bg};${col ? 'color:' + col + ';' : ''}${isHead || c.bold ? 'font-weight:700;' : ''}text-align:${c.align || 'left'};vertical-align:${c.valign || 'middle'}`;
        h += `<td style="${st}" data-r="${r}" data-c="${ci}"${c.cs > 1 ? ` colspan="${c.cs}"` : ''}${c.rs > 1 ? ` rowspan="${c.rs}"` : ''}><div class="cell"${editing ? ' contenteditable="true" spellcheck="' + (SL.spell !== false) + '"' : ''}>${c.html || ''}</div></td>`;
      });
      h += '</tr>';
    });
    return h + '</table>';
    function mix(hex, a) { return `color-mix(in srgb, ${hex} ${Math.round(a * 100)}%, white)`; }
  };

  SL.iconSVG = (el, pres, size) => {
    const body = (window.BTS_ICONS || {})[el.name] || (window.BTS_ICONS || {}).star || '';
    return `<svg class="icon-svg" width="100%" height="100%" viewBox="0 0 24 24" preserveAspectRatio="none" fill="${el.fillIcon ? SL.color(el.color || 'accent1', pres) : 'none'}" stroke="${SL.color(el.color || 'accent1', pres)}" stroke-width="${el.strokeWidth || 1.75}" stroke-linecap="round" stroke-linejoin="round" vector-effect="non-scaling-stroke">${body}</svg>`;
  };

  /* ---------- elements ---------- */
  SL.renderElement = (el, pres, opts = {}) => {
    const d = document.createElement('div');
    d.className = 'el el-' + el.type + (SL.isLine(el) ? ' el-line' : '');
    d.dataset.id = el.id;
    const tr = [];
    if (el.rot) tr.push(`rotate(${el.rot}deg)`);
    d.style.cssText = `left:${f2(el.x)}px;top:${f2(el.y)}px;width:${f2(el.w)}px;height:${f2(el.h)}px;${tr.length ? 'transform:' + tr.join(' ') + ';' : ''}${el.opacity != null && el.opacity < 1 ? 'opacity:' + el.opacity + ';' : ''}`;
    if (el.hidden && opts.editor) d.classList.add('el-hidden');
    const flip = (el.flipH || el.flipV) ? `transform:scale(${el.flipH ? -1 : 1},${el.flipV ? -1 : 1});` : '';
    if (el.shadow) d.classList.add('shadow');
    if (el.type === 'text' || el.type === 'shape') {
      const svg = SL.shapeSVG(el, pres);
      if (svg) d.insertAdjacentHTML('beforeend', `<div class="geo-wrap" style="${flip}">${svg}</div>`);
      if (!SL.isLine(el)) {
        const g = SL.geom(el.type === 'text' ? 'rect' : el.kind, el.w, el.h, el);
        const ins = el.inset || [9.6, 4.8, 9.6, 4.8];
        const pad = [g.inset[1] * el.h + ins[1], g.inset[2] * el.w + ins[2], g.inset[3] * el.h + ins[3], g.inset[0] * el.w + ins[0]];
        const tx = document.createElement('div'); tx.className = 'tx v-' + (el.valign || 'top') + (el.dir && el.dir !== 'horz' ? ' dir-' + el.dir : '');
        tx.style.padding = pad.map(v => f2(v) + 'px').join(' ');
        const tc = document.createElement('div'); tc.className = 'tc'; tc.style.cssText = SL.textStyle(el, pres);
        if (el.autofit === 'shrink' && el.fontScale && el.fontScale < 1) tc.style.fontSize = f2((el.size || 18) * el.fontScale) + 'pt';
        tc.innerHTML = el.html || '';
        if (opts.editor && el.placeholder && SL.isEmptyText(el) && opts.editingId !== el.id) { tc.dataset.ph = el.placeholder; tc.classList.add('ph-empty'); }
        if (el.wrap === false) tc.style.whiteSpace = 'nowrap';
        tx.appendChild(tc); d.appendChild(tx);
        if (el.ph && opts.editor) d.classList.add('is-ph');
      }
      if (opts.index != null) fillFields(d, opts);
    } else if (el.type === 'image') {
      const url = SL.assetURL(el.src);
      if (!url) {
        d.classList.add('img-ph');
        if (opts.editor) d.innerHTML = `<button class="img-ph-btn" data-act="fill-pic" title="Insert picture">${window.icon ? icon('image-plus', 34) : ''}<span>Insert picture</span></button>`;
      } else {
        const c = el.crop || { l: 0, t: 0, r: 0, b: 0 };
        const iw = el.w / Math.max(.01, 1 - c.l - c.r), ih = el.h / Math.max(.01, 1 - c.t - c.b);
        const st = el.stroke && el.stroke !== 'none' ? `outline:${f2((el.strokeWidth || 1) * 4 / 3)}px solid ${SL.color(el.stroke, pres)};outline-offset:-${f2((el.strokeWidth || 1) * 4 / 3 / 2)}px;` : '';
        d.innerHTML = `<div class="img-clip" style="border-radius:${el.radius ? f2(Math.min(el.w, el.h) * el.radius / 100) : 0}px;${st}${flip}"><img draggable="false" src="${url}" style="width:${f2(iw)}px;height:${f2(ih)}px;left:${f2(-c.l * iw)}px;top:${f2(-c.t * ih)}px" alt=""></div>`;
      }
    } else if (el.type === 'table') {
      d.innerHTML = SL.tableHTML(el, pres, opts.editingId === el.id);
    } else if (el.type === 'chart') {
      d.innerHTML = SL.chartSVG(el.chart, el.w, el.h, pres);
    } else if (el.type === 'icon') {
      d.innerHTML = `<div class="icon-wrap" style="${flip}">${SL.iconSVG(el, pres)}</div>`;
    }
    return d;
  };
  function fillFields(d, opts) {
    d.querySelectorAll('[data-field]').forEach(f => { if (f.dataset.field === 'num') f.textContent = String(opts.index + 1); if (f.dataset.field === 'date') f.textContent = new Date().toLocaleDateString(); });
  }

  SL.footerEls = (slide, pres, index) => {
    const f = pres.footer || {}; if (!f.date && !f.number && !f.text) return [];
    if (f.skipTitle && slide.layout === 'title') return [];
    const W = SL.W(pres), H = SL.H(pres), out = [], y = H - 50, base = { type: 'text', role: 'body', size: 12, h: 30, y, valign: 'middle', color: 'text', opacity: .75, footer: true };
    if (f.date) out.push(Object.assign({ id: 'ft-date', x: 60, w: 260, align: 'left', html: '<p>' + esc(f.dateText || new Date().toLocaleDateString('en-IN', { day: 'numeric', month: 'long', year: 'numeric' })) + '</p>' }, base));
    if (f.text) out.push(Object.assign({ id: 'ft-text', x: W / 2 - 300, w: 600, align: 'center', html: '<p>' + esc(f.text) + '</p>' }, base));
    if (f.number) out.push(Object.assign({ id: 'ft-num', x: W - 180, w: 120, align: 'right', html: '<p>' + (index + 1) + '</p>' }, base));
    return out;
  };

  /* Render a whole slide → .sl-slide element (W×H logical px). */
  SL.renderSlide = (slide, pres, opts = {}) => {
    const W = SL.W(pres), H = SL.H(pres);
    const s = document.createElement('div'); s.className = 'sl-slide' + (opts.editor ? ' editor' : '');
    s.style.cssText = `width:${W}px;height:${H}px;background:${SL.bgCSS(slide.background, pres)};`;
    s.dataset.sid = slide.id;
    s.insertAdjacentHTML('beforeend', SL.decoSVG(slide, pres));
    if (!slide.hideGraphics && slide.layer && slide.layer.length) {
      const lay = document.createElement('div'); lay.className = 'layer';
      slide.layer.forEach(el => lay.appendChild(SL.renderElement(el, pres, { index: opts.index })));
      s.appendChild(lay);
    }
    const idx = opts.index ?? pres.slides.indexOf(slide);
    for (const el of slide.elements) {
      if (!opts.editor && el.ph && ((el.type === 'text' && SL.isEmptyText(el)) || (el.type === 'image' && !el.src))) continue;
      if (!opts.editor && el.hidden) continue;
      const d = SL.renderElement(el, pres, Object.assign({}, opts, { index: idx }));
      if (opts.animHidden && opts.animHidden.has(el.id)) d.style.visibility = 'hidden';
      s.appendChild(d);
    }
    SL.footerEls(slide, pres, idx).forEach(el => { const d = SL.renderElement(el, pres, {}); d.classList.add('el-footer'); s.appendChild(d); });
    return s;
  };

  /* A scaled, non-interactive copy of a slide inside a box of width `w`. */
  SL.thumb = (slide, pres, w, index) => {
    const W = SL.W(pres), H = SL.H(pres), k = w / W;
    const box = document.createElement('div'); box.className = 'sl-thumb'; box.style.width = w + 'px'; box.style.height = f2(H * k) + 'px';
    const s = SL.renderSlide(slide, pres, { index });
    s.style.transform = `scale(${k})`; s.style.transformOrigin = '0 0';
    box.appendChild(s);
    return box;
  };
  SL.thumbHTML = (slide, pres, w, index) => SL.thumb(slide, pres, w, index).outerHTML;
})();
