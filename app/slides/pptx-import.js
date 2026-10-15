/* BTS Slides — PowerPoint (.pptx) import. Own OOXML parser on top of JSZip.
 * Handles: slide size & order, theme colours (schemeClr + clrMap + lumMod/lumOff/tint/shade/alpha), theme fonts,
 * placeholder inheritance (slide → layout → master) for position and text styles, master/layout graphics,
 * backgrounds (solid/gradient/picture/bgRef), shapes (preset + custom geometry), text runs, bullets, pictures with crop,
 * groups, connectors, tables, charts (bar/line/pie), notes, transitions, hidden slides.
 * Round-trips files written by BTS Slides (theme tag on the "Theme graphics" picture).
 */
(function () {
  const I = window.PPTXImport = {};
  const EMU = 9525; // EMU per px (96 dpi)
  const px = (v) => Math.round((+v || 0) / EMU * 100) / 100;
  const ln = (n) => n && (n.localName || n.nodeName.replace(/^.*:/, ''));
  const kids = (n, name) => n ? [...n.children].filter(c => !name || ln(c) === name) : [];
  const kid = (n, name) => n ? [...n.children].find(c => ln(c) === name) || null : null;
  const path = (n, ...names) => { for (const nm of names) { n = kid(n, nm); if (!n) return null; } return n; };
  const desc = (n, name) => { if (!n) return null; const all = n.getElementsByTagName('*'); for (const c of all) if (ln(c) === name) return c; return null; };
  const descAll = (n, name) => n ? [...n.getElementsByTagName('*')].filter(c => ln(c) === name) : [];
  const attr = (n, a) => n ? n.getAttribute(a) : null;
  const rid = (n, a = 'r:embed') => n ? (n.getAttribute(a) || n.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', a.split(':')[1])) : null;
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/emf', wmf: 'image/wmf' };

  /* ---------- zip & xml helpers ---------- */
  function ctx(zip) {
    const cache = {};
    const xml = async (p) => { if (!p) return null; if (p in cache) return cache[p]; const f = zip.file(p); if (!f) return (cache[p] = null); const t = await f.async('string'); return (cache[p] = new DOMParser().parseFromString(t, 'application/xml').documentElement); };
    const rels = async (p) => {
      const dir = p.replace(/[^/]*$/, ''), base = p.slice(dir.length);
      const r = await xml(dir + '_rels/' + base + '.rels'); const out = {};
      if (r) for (const e of kids(r, 'Relationship')) { const t = attr(e, 'Target'); out[attr(e, 'Id')] = { type: attr(e, 'Type').replace(/^.*\//, ''), target: attr(e, 'TargetMode') === 'External' ? t : resolve(dir, t), external: attr(e, 'TargetMode') === 'External' }; }
      return out;
    };
    const media = {};
    const dataURL = async (p) => {
      if (!p) return null; if (media[p] !== undefined) return media[p];
      const f = zip.file(p); if (!f) return (media[p] = null);
      const ext = p.split('.').pop().toLowerCase(); const mime = MIME[ext] || 'application/octet-stream';
      if (ext === 'emf' || ext === 'wmf' || ext === 'tif' || ext === 'tiff') return (media[p] = null); // not renderable in browser
      const b64 = await f.async('base64'); return (media[p] = `data:${mime};base64,${b64}`);
    };
    return { zip, xml, rels, dataURL };
  }
  function resolve(dir, t) {
    if (t.startsWith('/')) return t.slice(1);
    const parts = (dir + t).split('/'); const out = [];
    for (const s of parts) { if (s === '..') out.pop(); else if (s !== '.' && s !== '') out.push(s); }
    return out.join('/');
  }

  /* ---------- colours ---------- */
  function hex2rgb(h) { h = (h || '000000').replace('#', ''); return [parseInt(h.slice(0, 2), 16) || 0, parseInt(h.slice(2, 4), 16) || 0, parseInt(h.slice(4, 6), 16) || 0]; }
  function rgb2hex(r) { return '#' + r.map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join(''); }
  function rgb2hsl([r, g, b]) { r /= 255; g /= 255; b /= 255; const mx = Math.max(r, g, b), mn = Math.min(r, g, b); let h = 0, s = 0; const l = (mx + mn) / 2; if (mx !== mn) { const d = mx - mn; s = l > .5 ? d / (2 - mx - mn) : d / (mx + mn); h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; } return [h, s, l]; }
  function hsl2rgb([h, s, l]) { if (!s) return [l * 255, l * 255, l * 255]; const q = l < .5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; const f = (t) => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < .5 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; }; return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255]; }
  const PRST = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080', orange: 'FFA500', purple: '800080', darkBlue: '00008B', darkRed: '8B0000', darkGreen: '006400', lightGray: 'D3D3D3', cyan: '00FFFF', magenta: 'FF00FF' };

  // Returns { hex:'#rrggbb', a:0..1, scheme?:name } or null for a colour choice element (srgbClr, schemeClr…)
  function colorOf(n, T, phClr) {
    if (!n) return null;
    let base = null, scheme = null;
    switch (ln(n)) {
      case 'srgbClr': base = attr(n, 'val'); break;
      case 'sysClr': base = attr(n, 'lastClr') || (attr(n, 'val') === 'window' ? 'FFFFFF' : '000000'); break;
      case 'prstClr': base = PRST[attr(n, 'val')] || '000000'; break;
      case 'scrgbClr': base = rgb2hex([attr(n, 'r'), attr(n, 'g'), attr(n, 'b')].map(v => (+v / 100000) * 255)).slice(1); break;
      case 'hslClr': base = rgb2hex(hsl2rgb([+attr(n, 'hue') / 21600000, +attr(n, 'sat') / 100000, +attr(n, 'lum') / 100000])).slice(1); break;
      case 'schemeClr': {
        scheme = attr(n, 'val');
        if (scheme === 'phClr') { if (phClr) { base = phClr.hex.slice(1); scheme = phClr.scheme; } else base = '000000'; }
        else base = T.scheme(scheme);
        break;
      }
      default: return null;
    }
    let rgb = hex2rgb(base); let a = 1;
    for (const m of kids(n)) {
      const v = +attr(m, 'val') / 100000;
      switch (ln(m)) {
        case 'lumMod': { const h = rgb2hsl(rgb); h[2] *= v; rgb = hsl2rgb(h); break; }
        case 'lumOff': { const h = rgb2hsl(rgb); h[2] = Math.min(1, h[2] + v); rgb = hsl2rgb(h); break; }
        case 'satMod': { const h = rgb2hsl(rgb); h[1] = Math.min(1, h[1] * v); rgb = hsl2rgb(h); break; }
        case 'tint': rgb = rgb.map(c => c * v + 255 * (1 - v)); break;
        case 'shade': rgb = rgb.map(c => c * v); break;
        case 'alpha': a = v; break;
        case 'alphaMod': a *= v; break;
      }
    }
    return { hex: rgb2hex(rgb), a, scheme };
  }
  const colChild = (n) => n && kids(n).find(c => /Clr$/.test(ln(c)));
  const cssColor = (c) => !c ? null : c.a < 1 ? `rgba(${hex2rgb(c.hex).join(',')},${Math.round(c.a * 100) / 100})` : c.hex;

  /* ---------- theme ---------- */
  async function loadTheme(C, masterPath) {
    const mrels = await C.rels(masterPath);
    const tRel = Object.values(mrels).find(r => r.type === 'theme');
    const th = tRel ? await C.xml(tRel.target) : null;
    const master = await C.xml(masterPath);
    const clrMap = {}; const cm = kid(master, 'clrMap'); if (cm) for (const a of cm.attributes) clrMap[a.name] = a.value;
    const scheme = {};
    const cs = desc(th, 'clrScheme'); if (cs) for (const c of kids(cs)) { const ch = colChild(c); scheme[ln(c)] = ch ? (ln(ch) === 'sysClr' ? (attr(ch, 'lastClr') || '000000') : attr(ch, 'val')) : '000000'; }
    const fs = desc(th, 'fontScheme');
    const font = (k) => { const n = fs && desc(kid(fs, k), 'latin'); return n ? attr(n, 'typeface') : null; };
    const fmt = desc(th, 'fmtScheme');
    const T = {
      clrMap, raw: scheme, name: attr(desc(th, 'clrScheme'), 'name') || 'Imported', major: font('majorFont') || 'Calibri', minor: font('minorFont') || 'Calibri',
      fillStyles: fmt ? kids(kid(fmt, 'fillStyleLst')) : [], lnStyles: fmt ? kids(kid(fmt, 'lnStyleLst')) : [], bgStyles: fmt ? kids(kid(fmt, 'bgFillStyleLst')) : [],
      scheme(name) {
        const mapped = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
        let k = name; if (clrMap[name]) k = clrMap[name]; else if (mapped[name]) k = mapped[name];
        return scheme[k] || scheme[name] || '000000';
      },
      fontName(tf) { if (!tf) return null; if (tf.startsWith('+mj')) return T.major; if (tf.startsWith('+mn')) return T.minor; return tf; },
    };
    return T;
  }

  /* ---------- fills ---------- */
  function fillOf(spPr, T, phClr) {
    if (!spPr) return undefined;
    for (const c of kids(spPr)) {
      const n = ln(c);
      if (n === 'noFill') return 'none';
      if (n === 'solidFill') { const col = colorOf(colChild(c), T, phClr); return col ? { css: cssColor(col), col } : undefined; }
      if (n === 'gradFill') {
        const stops = descAll(c, 'gs').map(g => ({ pos: +attr(g, 'pos'), col: colorOf(colChild(g), T, phClr) })).filter(s => s.col).sort((a, b) => a.pos - b.pos);
        if (!stops.length) return undefined;
        const lin = kid(c, 'lin'); const ang = lin ? +attr(lin, 'ang') / 60000 : 90;
        return { grad: { type: 'gradient', color: cssColor(stops[0].col), color2: cssColor(stops[stops.length - 1].col), angle: Math.round((ang + 90) % 360) }, col: stops[0].col };
      }
      if (n === 'blipFill') return { blip: c };
      if (n === 'grpFill') return { grp: true };
      if (n === 'pattFill') { const fg = colorOf(colChild(kid(c, 'fgClr')), T, phClr); return fg ? { css: cssColor(fg), col: fg } : undefined; }
    }
    return undefined;
  }
  function styleRef(sp, name, T) { // p:style/a:fillRef etc → { idx, col }
    const st = kid(sp, 'style'); const r = st && kid(st, name); if (!r) return null;
    return { idx: +attr(r, 'idx'), col: colorOf(colChild(r), T) };
  }

  /* ---------- text styles ---------- */
  // Collect lvlNpPr nodes from a list-style container (lstStyle / titleStyle / bodyStyle / otherStyle)
  const lvlNode = (lst, lvl) => lst ? kid(lst, `lvl${lvl + 1}pPr`) : null;
  function mergedPPr(chain, lvl) { // chain: [lstStyle containers…] highest priority first; returns resolved props
    const out = {};
    for (let i = chain.length - 1; i >= 0; i--) { // lowest priority first, overwrite upwards
      const n = chain[i] && (chain[i].__lvl ? chain[i] : lvlNode(chain[i], lvl)); if (!n) continue;
      applyPPr(out, n);
    }
    return out;
  }
  function applyPPr(o, n) {
    if (attr(n, 'algn')) o.algn = attr(n, 'algn');
    if (attr(n, 'marL') != null) o.marL = +attr(n, 'marL');
    if (attr(n, 'indent') != null) o.indent = +attr(n, 'indent');
    for (const c of kids(n)) {
      const k = ln(c);
      if (k === 'buNone') o.bu = 'none'; if (k === 'buChar') { o.bu = 'char'; o.buChar = attr(c, 'char'); } if (k === 'buAutoNum') { o.bu = 'num'; o.buNum = attr(c, 'type'); } if (k === 'buBlip') o.bu = 'char';
      if (k === 'lnSpc') { const p = kid(c, 'spcPct'); if (p) o.lnSpc = +attr(p, 'val') / 100000; const pt = kid(c, 'spcPts'); if (pt) o.lnSpcPts = +attr(pt, 'val') / 100; }
      if (k === 'spcAft') { const pt = kid(c, 'spcPts'); if (pt) o.spcAft = +attr(pt, 'val') / 100; }
      if (k === 'spcBef') { const pt = kid(c, 'spcPts'); if (pt) o.spcBef = +attr(pt, 'val') / 100; }
      if (k === 'defRPr') { o.rpr = o.rpr || {}; applyRPr(o.rpr, c); }
    }
    return o;
  }
  function applyRPr(o, n) {
    if (!n) return o;
    if (attr(n, 'sz')) o.sz = +attr(n, 'sz') / 100;
    if (attr(n, 'b') != null) o.b = attr(n, 'b') === '1' || attr(n, 'b') === 'true';
    if (attr(n, 'i') != null) o.i = attr(n, 'i') === '1' || attr(n, 'i') === 'true';
    if (attr(n, 'u') != null) o.u = attr(n, 'u') !== 'none';
    if (attr(n, 'strike') != null) o.strike = attr(n, 'strike') !== 'noStrike';
    if (attr(n, 'baseline')) o.baseline = +attr(n, 'baseline');
    if (attr(n, 'cap')) o.cap = attr(n, 'cap');
    if (attr(n, 'spc')) o.spc = +attr(n, 'spc') / 100;
    for (const c of kids(n)) {
      const k = ln(c);
      if (k === 'solidFill') o.fill = colChild(c);
      if (k === 'gradFill') { const g = desc(c, 'gs'); if (g) o.fill = colChild(g); }
      if (k === 'noFill') o.fill = null;
      if (k === 'latin') o.font = attr(c, 'typeface');
      if (k === 'highlight') o.hl = colChild(c);
      if (k === 'hlinkClick') o.link = c;
    }
    return o;
  }

  /* ---------- placeholder lookup ---------- */
  function phInfo(sp) {
    const nv = kids(sp).find(c => /^nv/.test(ln(c))); const ph = nv && desc(nv, 'ph');
    if (!ph) return null;
    return { type: attr(ph, 'type') || 'body', idx: attr(ph, 'idx'), node: ph };
  }
  const TITLE_TYPES = new Set(['title', 'ctrTitle']);
  function findPh(tree, info) {
    if (!tree || !info) return null;
    const sps = descAll(tree, 'sp');
    let byIdx = null, byType = null;
    for (const s of sps) {
      const p = phInfo(s); if (!p) continue;
      if (info.idx != null && p.idx === info.idx) byIdx = byIdx || s;
      const t1 = info.type, t2 = p.type;
      if (t1 === t2 || (TITLE_TYPES.has(t1) && TITLE_TYPES.has(t2)) || (t1 === 'subTitle' && t2 === 'body') || (t1 === 'obj' && t2 === 'body') || (t1 === 'body' && t2 === 'obj')) byType = byType || s;
    }
    if (byIdx) return byIdx;
    if (byType) return byType;
    // object/body placeholders with only idx fall back to master body
    if (['body', 'obj', 'subTitle', 'pic', 'tbl', 'chart'].includes(info.type)) return sps.find(s => { const p = phInfo(s); return p && (p.type === 'body' || p.type === 'obj'); }) || null;
    return null;
  }
  function xfrmOf(spPr) {
    const x = spPr && kid(spPr, 'xfrm'); if (!x) return null;
    const off = kid(x, 'off'), ext = kid(x, 'ext');
    if (!off || !ext) return null;
    return { x: px(attr(off, 'x')), y: px(attr(off, 'y')), w: px(attr(ext, 'cx')), h: px(attr(ext, 'cy')), rot: (+attr(x, 'rot') || 0) / 60000, flipH: attr(x, 'flipH') === '1', flipV: attr(x, 'flipV') === '1' };
  }

  /* ---------- custom geometry → path ---------- */
  function custPath(cg) {
    const list = desc(cg, 'pathLst'); if (!list) return null;
    let d = ''; let W = 0, H = 0; const noFill = [];
    for (const p of kids(list, 'path')) {
      const pw = +attr(p, 'w') || 0, ph = +attr(p, 'h') || 0; W = Math.max(W, pw); H = Math.max(H, ph);
      let cur = [0, 0];
      for (const c of kids(p)) {
        const pts = kids(c, 'pt').map(t => [+attr(t, 'x'), +attr(t, 'y')]);
        switch (ln(c)) {
          case 'moveTo': d += ` M${pts[0][0]},${pts[0][1]}`; cur = pts[0]; break;
          case 'lnTo': d += ` L${pts[0][0]},${pts[0][1]}`; cur = pts[0]; break;
          case 'cubicBezTo': d += ` C${pts.map(q => q[0] + ',' + q[1]).join(' ')}`; cur = pts[2]; break;
          case 'quadBezTo': d += ` Q${pts.map(q => q[0] + ',' + q[1]).join(' ')}`; cur = pts[1]; break;
          case 'arcTo': { // approximate with elliptical arc
            const wR = +attr(c, 'wR'), hR = +attr(c, 'hR'), st = +attr(c, 'stAng') / 60000 * Math.PI / 180, sw = +attr(c, 'swAng') / 60000 * Math.PI / 180;
            const cx = cur[0] - wR * Math.cos(st), cy = cur[1] - hR * Math.sin(st); const ex = cx + wR * Math.cos(st + sw), ey = cy + hR * Math.sin(st + sw);
            d += ` A${wR},${hR} 0 ${Math.abs(sw) > Math.PI ? 1 : 0},${sw > 0 ? 1 : 0} ${Math.round(ex)},${Math.round(ey)}`; cur = [ex, ey]; break;
          }
          case 'close': d += ' Z'; break;
        }
      }
      if (attr(p, 'fill') === 'none') noFill.push(true);
    }
    if (!d) return null;
    // arc flags written as "0,1" would be scaled by the x,y regex — rewrite arc flags with spaces
    d = d.replace(/A(\d+(?:\.\d+)?),(\d+(?:\.\d+)?) 0 (\d),(\d)/g, (m, a, b, f1, f2) => `A${a},${b} 0 ${f1} ${f2}`);
    return { d: d.trim(), w: W || 1, h: H || 1, noFill: noFill.length > 0 };
  }

  /* ====================================================================== */
  /*  Main                                                                    */
  /* ====================================================================== */
  I.parse = async (data) => {
    const zip = await window.JSZip.loadAsync(data);
    const C = ctx(zip);
    const pres = await C.xml('ppt/presentation.xml'); if (!pres) throw new Error('Not a PowerPoint presentation (ppt/presentation.xml missing)');
    const prels = await C.rels('ppt/presentation.xml');
    const sz = kid(pres, 'sldSz');
    const W = sz ? px(attr(sz, 'cx')) : 1280, H = sz ? px(attr(sz, 'cy')) : 720;
    const out = { width: Math.round(W), height: Math.round(H), theme: { id: 'custom', variant: 0 }, footer: { date: false, number: false, text: '', skipTitle: true }, slides: [] };
    const ids = kids(kid(pres, 'sldIdLst'), 'sldId');
    const slidePaths = ids.map(n => prels[rid(n, 'r:id')]).filter(Boolean).map(r => r.target);
    // sections (PowerPoint 2010 extension)
    const secLst = desc(pres, 'sectionLst');
    const sectionDefs = secLst ? kids(secLst, 'section').map(s => ({ name: attr(s, 'name'), ids: descAll(s, 'sldId').map(x => attr(x, 'id')) })) : [];
    const idOrder = ids.map(n => attr(n, 'id'));
    const defaultTx = kid(pres, 'defaultTextStyle');

    let firstTheme = null; let btsTheme = null;
    const masterCache = {};
    // comment authors (legacy comments)
    const authors = {}; const caRel = Object.values(prels).find(r => r.type === 'commentAuthors');
    if (caRel) { const ca = await C.xml(caRel.target); if (ca) kids(ca, 'cmAuthor').forEach(a => { authors[attr(a, 'id')] = attr(a, 'name') || 'Reviewer'; }); }
    for (let si = 0; si < slidePaths.length; si++) {
      const sp = slidePaths[si];
      try {
        const slide = await parseSlide(C, sp, { W, H, defaultTx, masterCache, idx: si, authors });
        if (!firstTheme) firstTheme = slide._T;
        if (slide._bts) btsTheme = btsTheme || slide._bts;
        delete slide._T; delete slide._bts;
        out.slides.push(slide);
      } catch (err) { console.warn('Slide', si + 1, 'could not be fully read:', err); out.slides.push({ id: SL.uid('s'), layout: 'blank', background: null, transition: { type: 'none', dur: 700 }, notes: '', hidden: false, comments: [], layer: [], elements: [SL.newElement('text', { x: 80, y: 80, w: 800, h: 60, html: '<p>This slide could not be imported.</p>' })] }); }
    }
    if (!out.slides.length) out.slides.push(SL.newSlide(out, 'title'));
    if (btsTheme) out.theme = { id: btsTheme.theme, variant: +btsTheme.v || 0 };
    else if (firstTheme) {
      const T = firstTheme; const r = (k) => '#' + T.scheme(k);
      out.theme.custom = { name: T.name, titleFont: T.major, bodyFont: T.minor, titleBold: false, colors: { bg: r('bg1'), bg2: r('bg2'), text: r('tx1'), title: r('tx1'), accents: [1, 2, 3, 4, 5, 6].map(i => '#' + (T.raw['accent' + i] || '777777')) } };
    }
    if (sectionDefs.length) {
      out.sections = [];
      for (const s of sectionDefs) { const first = s.ids.map(id => idOrder.indexOf(id)).filter(i => i >= 0).sort((a, b) => a - b)[0]; if (first != null && out.slides[first]) out.sections.push({ start: out.slides[first].id, name: s.name || 'Section' }); }
    }
    // custom shows and slide show settings written by PowerPoint (or BTS)
    const csl = desc(pres, 'custShowLst');
    if (csl) { const ridToIdx = {}; ids.forEach((n, i) => { ridToIdx[rid(n, 'r:id')] = i; }); out.customShows = kids(csl, 'custShow').map(c => ({ id: SL.uid('cs'), _pid: attr(c, 'id'), name: attr(c, 'name') || 'Custom Show', slides: descAll(c, 'sld').map(x => ridToIdx[rid(x, 'r:id')]).filter(i => i != null && out.slides[i]).map(i => out.slides[i].id) })); }
    const ppRel = Object.values(prels).find(r => r.type === 'presProps');
    const pp = ppRel ? await C.xml(ppRel.target) : null; const shp = pp && desc(pp, 'showPr');
    if (shp) {
      const st = { type: desc(shp, 'kiosk') ? 'kiosk' : desc(shp, 'browse') ? 'window' : 'speaker', loop: attr(shp, 'loop') === '1', useTimings: attr(shp, 'useTimings') !== '0', noAnims: attr(shp, 'showAnimation') === '0' };
      st.kiosk = st.type === 'kiosk'; const rg = desc(shp, 'sldRg'); if (rg) st.range = { from: +attr(rg, 'st') || 1, to: +attr(rg, 'end') || out.slides.length };
      const cs = desc(shp, 'custShow'); if (cs && out.customShows) { const m = out.customShows.find(c => c._pid === attr(cs, 'id')); if (m) st.custom = m.id; }
      if (st.loop || st.kiosk || st.range || st.custom || st.type !== 'speaker' || !st.useTimings || st.noAnims) out.showSetup = st;
    }
    (out.customShows || []).forEach(c => delete c._pid);
    // document properties
    out.props = await readProps(C);
    // BTS metadata: restore unchanged slides exactly + presentation-level settings
    try { await applyMeta(zip, C, out, slidePaths); } catch (err) { console.warn('BTS metadata ignored:', err); }
    const clean = (e) => { delete e._oname; };
    out.slides.forEach(s => { (s.elements || []).forEach(clean); (s.layer || []).forEach(clean); });
    return out;
  };

  async function readProps(C) {
    const core = await C.xml('docProps/core.xml'), app = await C.xml('docProps/app.xml'); const p = {};
    const t = (n, name) => { const e = desc(n, name); return e ? e.textContent.trim() : ''; };
    if (core) { p.title = t(core, 'title'); p.subject = t(core, 'subject'); p.author = t(core, 'creator'); p.keywords = t(core, 'keywords'); p.category = t(core, 'category'); p.comments = t(core, 'description'); }
    if (app) p.company = t(app, 'Company');
    Object.keys(p).forEach(k => { if (!p[k]) delete p[k]; });
    return p;
  }
  async function rawHash(zip, slidePath) {
    const f = zip.file(slidePath); if (!f) return '';
    let str = await f.async('string');
    const relsPath = slidePath.replace(/[^/]*$/, '_rels/') + slidePath.split('/').pop() + '.rels';
    const rf = zip.file(relsPath);
    if (rf) { const rels = await rf.async('string'); for (const m of rels.matchAll(/Target="\.\.\/(notesSlides\/notesSlide\d+\.xml|comments\/comment\d+\.xml)"/g)) { const g = zip.file('ppt/' + m[1]); if (g) str += '\n' + await g.async('string'); } }
    return window.PPTXExport.hash(str);
  }
  async function applyMeta(zip, C, out, slidePaths) {
    const part = Object.keys(zip.files).find(f => /^customXml\/[^/]+\.xml$/.test(f) && !/itemProps/i.test(f) && !/_rels/.test(f));
    let metaFile = null;
    for (const f of Object.keys(zip.files).filter(f => /^customXml\/[^/]+\.xml$/.test(f))) { const t = await zip.file(f).async('string'); if (t.includes('urn:bts-office:slides')) { metaFile = t; break; } }
    void part; if (!metaFile) return;
    const doc = new DOMParser().parseFromString(metaFile, 'application/xml'); const meta = JSON.parse(doc.documentElement.textContent || '{}');
    if (!meta || meta.app !== 'BTS Slides') return;
    const restored = new Set();
    for (let i = 0; i < out.slides.length && i < (meta.slides || []).length; i++) {
      const m = meta.slides[i]; if (!m || !m.s) continue;
      if (m.h !== await rawHash(zip, slidePaths[i])) continue;
      const parsed = out.slides[i]; const s = m.s; const byName = {};
      [...parsed.elements, ...(parsed.layer || [])].forEach(e => { if (e._oname) byName[e._oname] = e; });
      const fixImages = (list) => list.filter(el => {
        if (el.type !== 'image') return true;
        if (!el._n) { delete el._n; return !!el.ph; }
        const p = byName[el._n]; delete el._n;
        if (!p || !p.src) return !!el.ph;
        const baked = el.radius || (el.crop && (el.crop.l || el.crop.t || el.crop.r || el.crop.b)) || /^data:image\/svg/.test(SL.assetURL(el.src || ''));
        el.src = p.src; if (baked) { el.crop = null; el.radius = 0; }
        return true;
      });
      s.elements = fixImages(s.elements || []); s.layer = fixImages(s.layer || []);
      if (s.background && s.background._img) { s.background = parsed.background && parsed.background.type === 'image' ? Object.assign({}, parsed.background, { fit: 'stretch' }) : null; }
      out.slides[i] = s; restored.add(i);
    }
    // theme overrides (fonts / colours / imported custom theme) when the theme matches
    if (meta.theme && (meta.theme.id === out.theme.id || (out.theme.id === 'custom' && meta.theme.id !== 'custom' && restored.size))) out.theme = meta.theme;
    if (meta.master) {
      const resolve = (el) => { if (el.type !== 'image') return true; const sl = out.slides[el._s]; const p = sl && [...sl.elements, ...(sl.layer || [])].find(e => e._oname === el._n); delete el._n; delete el._s; if (p && p.src) { el.src = p.src; el.crop = null; el.radius = 0; return true; } return false; };
      // master graphics live in the master, not on each slide: drop the exported copies from slides read from XML
      out.slides.forEach((s, i) => { if (!restored.has(i)) s.elements = s.elements.filter(e => !/^Master graphic O\d+$/.test(e._oname || '')); });
      if (meta.master.graphics) meta.master.graphics = meta.master.graphics.filter(resolve);
      if (meta.master.layoutGraphics) Object.keys(meta.master.layoutGraphics).forEach(k => { meta.master.layoutGraphics[k] = meta.master.layoutGraphics[k].filter(resolve); });
      out.master = meta.master;
    }
    if (meta.customShows && meta.customShows.length) out.customShows = meta.customShows.map(c => ({ id: c.id || SL.uid('cs'), name: c.name, slides: c.slides.map(i => out.slides[i] && out.slides[i].id).filter(Boolean) }));
    if (meta.showSetup) out.showSetup = meta.showSetup;
    if (meta.footer) out.footer = meta.footer;
    out._restored = restored.size;
  }

  const LAYOUT_MAP = { title: 'title', obj: 'titleContent', tx: 'titleContent', secHead: 'section', twoObj: 'twoContent', twoTxTwoObj: 'comparison', titleOnly: 'titleOnly', blank: 'blank', picTx: 'pictureCaption', objTx: 'twoContent', txAndObj: 'twoContent', cust: 'titleContent' };

  async function parseSlide(C, slidePath, G) {
    const sld = await C.xml(slidePath); const rels = await C.rels(slidePath);
    const layoutRel = Object.values(rels).find(r => r.type === 'slideLayout');
    const layoutPath = layoutRel && layoutRel.target;
    const layout = layoutPath ? await C.xml(layoutPath) : null;
    const lrels = layoutPath ? await C.rels(layoutPath) : {};
    const masterRel = Object.values(lrels).find(r => r.type === 'slideMaster');
    const masterPath = masterRel && masterRel.target;
    const master = masterPath ? await C.xml(masterPath) : null;
    const mrels = masterPath ? await C.rels(masterPath) : {};
    if (masterPath && !G.masterCache[masterPath]) G.masterCache[masterPath] = await loadTheme(C, masterPath);
    const T = masterPath ? G.masterCache[masterPath] : { scheme: () => '000000', fontName: (f) => f, major: 'Calibri', minor: 'Calibri', clrMap: {}, raw: {}, fillStyles: [], lnStyles: [], bgStyles: [] };
    // clrMapOvr on slide / layout
    const ovr = desc(kid(sld, 'clrMapOvr'), 'overrideClrMapping') || desc(kid(layout, 'clrMapOvr'), 'overrideClrMapping');
    const Tl = ovr ? Object.assign(Object.create(T), { scheme(name) { const m = attr(ovr, name); return m ? T.raw[m] || T.scheme(name) : T.scheme(name); } }) : T;

    const txStyles = master && kid(master, 'txStyles');
    const S = { id: SL.uid('s'), layout: LAYOUT_MAP[attr(layout, 'type')] || 'titleContent', background: null, hideGraphics: false, transition: { type: 'none', dur: 700 }, notes: '', hidden: attr(sld, 'show') === '0', comments: [], layer: [], elements: [] };
    S._T = T;
    const env = { C, T: Tl, rels, layout, master, lrels, mrels, txStyles, G, slide: S, layoutPath, masterPath };

    // background: slide → layout → master
    S.background = await bgOf(kid(kid(sld, 'cSld'), 'bg'), env, rels) || await bgOf(kid(kid(layout, 'cSld'), 'bg'), env, lrels) || await bgOf(kid(kid(master, 'cSld'), 'bg'), env, mrels) || { type: 'solid', color: '#' + Tl.scheme('bg1') };

    // master + layout graphics (non-placeholder shapes)
    const showMaster = attr(sld, 'showMasterSp') !== '0' && attr(layout, 'showMasterSp') !== '0';
    if (showMaster && master) await walkTree(kid(kid(master, 'cSld'), 'spTree'), Object.assign({}, env, { rels: mrels, level: 'master' }), S.layer, null, true);
    if (layout && attr(sld, 'showMasterSp') !== '0') await walkTree(kid(kid(layout, 'cSld'), 'spTree'), Object.assign({}, env, { rels: lrels, level: 'layout' }), S.layer, null, true);

    // slide shapes
    await walkTree(kid(kid(sld, 'cSld'), 'spTree'), Object.assign({}, env, { level: 'slide' }), S.elements, null, false);

    // BTS round-trip: theme graphics picture
    const tg = S.elements.findIndex(e => e.alt && /^bts:theme=/.test(e.alt));
    if (tg >= 0) {
      const m = /bts:theme=([\w-]+);v=(\d+);layout=(\w+)/.exec(S.elements[tg].alt);
      if (m) { S._bts = { theme: m[1], v: m[2] }; S.layout = m[3]; S.layer = []; if (S.elements[tg].type !== 'image') S.hideGraphics = true; }
      S.elements.splice(tg, 1);
      // restore placeholders by object name
      S.elements.forEach(e => { if (/^Title /.test(e.name || '')) { e.ph = 'title'; e.role = 'title'; e.placeholder = 'Click to add title'; delete e.name; } else if (/^Subtitle /.test(e.name || '')) { e.ph = 'subtitle'; e.placeholder = 'Click to add subtitle'; delete e.name; } else if (/^Text Placeholder /.test(e.name || '')) { e.ph = 'body'; e.placeholder = 'Click to add text'; delete e.name; } else if (/ O\d+$/.test(e.name || '')) { e.name = e.name.replace(/ ?O\d+$/, '') || undefined; if (!e.name) delete e.name; } });
      // background equal to theme background → inherit
      if (m) { const fake = { theme: { id: m[1], variant: +m[2] } }; const tb = SL.themeBg(fake); const b = S.background; if (b && ((b.type === 'solid' && tb.type === 'solid' && sameColor(b.color, SL.color(tb.color, fake))) || (b.type === 'gradient' && tb.type === 'gradient' && sameColor(b.color, SL.color(tb.color, fake))))) S.background = null; }
      // empty placeholders re-created by layout
      const specs = SL.layoutPlaceholders(S.layout, G.W, G.H, S._bts ? { theme: { id: S._bts.theme, variant: +S._bts.v } } : null);
      const have = (k) => S.elements.filter(e => e.ph === k).length;
      const need = { title: 0, subtitle: 0, body: 0, pic: 0 }; specs.forEach(s => need[s.ph]++);
      const add = [];
      specs.forEach(sp => { const k = sp.ph; if (have(k) + add.filter(a => a.ph === k).length < need[k]) { const cnt = specs.filter(s => s.ph === k).indexOf(sp); if (cnt >= have(k)) add.push(SL.newElement(sp.type, sp)); } });
      S.elements.unshift(...add);
    }

    // notes
    const nRel = Object.values(rels).find(r => r.type === 'notesSlide');
    if (nRel) {
      const nx = await C.xml(nRel.target);
      if (nx) { for (const sp of descAll(nx, 'sp')) { const p = phInfo(sp); if (p && p.type === 'body') { S.notes = descAll(sp, 'p').map(pp => descAll(pp, 't').map(t => t.textContent).join('')).join('\n').trim(); break; } } }
    }
    // transition
    const tr = findTransition(sld); if (tr) S.transition = tr;
    // comments (legacy)
    const cRel = Object.values(rels).find(r => r.type === 'comments');
    if (cRel) {
      const cx = await C.xml(cRel.target);
      if (cx) {
        const authors = G.authors || {}; const byKey = {}; const list = [];
        kids(cx, 'cm').forEach(c => {
          let text = (desc(c, 'text') || {}).textContent || ''; const resolved = /^\[Resolved\] /.test(text); text = text.replace(/^\[Resolved\] /, '');
          const o = { id: SL.uid('c'), author: authors[attr(c, 'authorId')] || 'Reviewer', text, time: Date.parse(attr(c, 'dt')) || Date.now(), x: px((+attr(desc(c, 'pos'), 'x') || 0) * 1587.5), y: px((+attr(desc(c, 'pos'), 'y') || 0) * 1587.5), replies: [] };
          if (resolved) o.resolved = true;
          const par = desc(c, 'parentCm'); const key = attr(c, 'authorId') + ':' + attr(c, 'idx');
          if (par && byKey[attr(par, 'authorId') + ':' + attr(par, 'idx')]) byKey[attr(par, 'authorId') + ':' + attr(par, 'idx')].replies.push({ id: SL.uid('r'), author: o.author, text: o.text, time: o.time });
          else { byKey[key] = o; list.push(o); }
        });
        S.comments = list;
      }
    }
    return S;
  }
  function sameColor(a, b) { return String(a || '').toLowerCase() === String(b || '').toLowerCase(); }

  function findTransition(sld) {
    let tr = kid(sld, 'transition');
    const alt = kids(sld, 'AlternateContent')[0];
    if (!tr && alt) { const ch = kid(alt, 'Choice'), fb = kid(alt, 'Fallback'); tr = (ch && kid(ch, 'transition')) || (fb && kid(fb, 'transition')); if (ch && fb && kid(fb, 'transition') && kid(ch, 'transition')) { const c = kid(ch, 'transition'); const f = kid(fb, 'transition'); tr = c; if (!kids(c).length) tr = f; tr.__fb = f; } }
    if (!tr) return null;
    const eff = kids(tr).find(k => ln(k) !== 'extLst' && ln(k) !== 'sndAc');
    const name = eff ? ln(eff) : 'cut';
    const MAP = { fade: 'fade', push: 'push', wipe: 'wipe', split: 'split', cover: 'cover', pull: 'reveal', zoom: 'zoom', flip: 'flip', prism: 'cube', cube: 'cube', dissolve: 'fade', cut: 'none', randomBar: 'wipe', blinds: 'wipe', checker: 'fade', circle: 'zoom', diamond: 'zoom', plus: 'zoom', wedge: 'wipe', wheel: 'wipe', newsflash: 'zoom', strips: 'wipe', comb: 'split', ripple: 'zoom', vortex: 'zoom', shred: 'fade', switch: 'flip', gallery: 'push', conveyor: 'push', pan: 'push', ferris: 'push', flythrough: 'zoom', warp: 'zoom', glitter: 'fade', honeycomb: 'fade', flash: 'fade', doors: 'split', window: 'split', reveal: 'reveal', morph: 'fade', curtains: 'split', wind: 'push', prestige: 'fade', fracture: 'fade', crush: 'fade', peelOff: 'reveal', pageCurlDouble: 'reveal', pageCurlSingle: 'reveal', airplane: 'fade', origami: 'flip', vortexs: 'zoom' };
    const type = MAP[name] || 'fade';
    const dirA = eff && attr(eff, 'dir');
    const DIR = { l: 'right', r: 'left', u: 'down', d: 'up' }; // pptx dir = direction it comes FROM
    let dir = dirA ? (type === 'reveal' ? ({ l: 'left', r: 'right', u: 'up', d: 'down' }[dirA]) : DIR[dirA]) : undefined;
    if (type === 'flip') dir = dirA === 'r' ? 'right' : 'left';
    if (type === 'split') dir = attr(eff, 'orient') === 'horz' ? 'up' : 'left';
    let dur = +(tr.getAttribute('p14:dur') || 0);
    if (!dur) { const spd = attr(tr, 'spd') || (tr.__fb && attr(tr.__fb, 'spd')); dur = spd === 'fast' ? 400 : spd === 'slow' ? 1200 : 700; }
    const o = { type, dur }; if (dir) o.dir = dir;
    const adv = attr(tr, 'advTm'); if (adv != null) o.auto = +adv / 1000;
    if (attr(tr, 'advClick') === '0') o.click = false;
    return o;
  }

  async function bgOf(bg, env, rels) {
    if (!bg) return null;
    const pr = kid(bg, 'bgPr');
    if (pr) {
      const f = fillOf(pr, env.T);
      if (f && f.css) return { type: 'solid', color: f.css };
      if (f && f.grad) return f.grad;
      if (f && f.blip) { const b = kid(f.blip, 'blip'); const r = rels[rid(b)]; const url = r && await env.C.dataURL(r.target); if (url) return { type: 'image', src: SL.addAsset(url), fit: 'stretch', color: '#ffffff' }; }
      return null;
    }
    const ref = kid(bg, 'bgRef');
    if (ref) {
      const col = colorOf(colChild(ref), env.T); const idx = +attr(ref, 'idx');
      const sty = idx >= 1001 ? env.T.bgStyles[idx - 1001] : env.T.fillStyles[idx - 1];
      if (sty) { const wrap = { children: [sty] }; void wrap; const f = fillOf({ children: [sty] }, env.T, col); if (f && f.grad) return f.grad; if (f && f.css) return { type: 'solid', color: f.css }; if (f && f.blip) { /* theme pictures rarely used */ } }
      if (col) return { type: 'solid', color: cssColor(col) };
    }
    return null;
  }

  /* ---------- tree walking ---------- */
  // grp: { tx: fn(x,y,w,h)->{x,y,w,h} } for group transforms; layerMode: skip placeholders
  async function walkTree(tree, env, out, grp, layerMode) {
    if (!tree) return;
    for (const n of kids(tree)) {
      const name = ln(n);
      try {
        if (name === 'sp') { if (layerMode && phInfo(n)) continue; const el = await shapeEl(n, env, grp); if (el) out.push(objMeta(n, el)); }
        else if (name === 'pic') { if (layerMode && phInfo(n)) continue; const el = await picEl(n, env, grp); if (el) out.push(objMeta(n, el)); }
        else if (name === 'cxnSp') { const el = await shapeEl(n, env, grp, true); if (el) out.push(objMeta(n, el)); }
        else if (name === 'graphicFrame') { if (layerMode && phInfo(n)) continue; const el = await frameEl(n, env, grp); if (el) out.push(objMeta(n, el)); }
        else if (name === 'grpSp') {
          const gx = kid(kid(n, 'grpSpPr'), 'xfrm'); const sub = []; const gid = SL.uid('g');
          let g2 = grp;
          if (gx) {
            const off = kid(gx, 'off'), ext = kid(gx, 'ext'), co = kid(gx, 'chOff'), ce = kid(gx, 'chExt');
            const o = { x: px(attr(off, 'x')), y: px(attr(off, 'y')), w: px(attr(ext, 'cx')), h: px(attr(ext, 'cy')) };
            const c = co && ce ? { x: px(attr(co, 'x')), y: px(attr(co, 'y')), w: px(attr(ce, 'cx')) || o.w, h: px(attr(ce, 'cy')) || o.h } : o;
            const sx = c.w ? o.w / c.w : 1, sy = c.h ? o.h / c.h : 1;
            const local = (b) => ({ x: o.x + (b.x - c.x) * sx, y: o.y + (b.y - c.y) * sy, w: b.w * sx, h: b.h * sy, rot: b.rot, flipH: b.flipH, flipV: b.flipV });
            const parent = grp;
            g2 = { tx: (b) => parent ? parent.tx(local(b)) : local(b), rot: ((+attr(gx, 'rot') || 0) / 60000) + (parent ? parent.rot : 0), fill: kid(n, 'grpSpPr') };
          }
          await walkTree(n, env, sub, g2, layerMode);
          if (!layerMode && sub.length > 1) sub.forEach(e => { if (!e.g) e.g = gid; });
          out.push(...sub);
        } else if (name === 'AlternateContent') {
          const ch = kid(n, 'Choice'), fb = kid(n, 'Fallback');
          await walkTree(fb || ch, env, out, grp, layerMode);
        }
      } catch (err) { console.warn('Skipped an object:', err); }
    }
  }
  // raw object name (BTS round trip), alt text and the Office "decorative" flag for any object
  function objMeta(node, el) {
    const nv = desc(node, 'cNvPr'); if (!nv) return el;
    const n = attr(nv, 'name'); if (n) el._oname = n;
    const d = attr(nv, 'descr'); if (d && !el.alt) el.alt = d;
    if (descAll(nv, 'decorative').some(x => attr(x, 'val') === '1')) { el.decorative = true; delete el.alt; }
    return el;
  }
  function applyGroup(b, grp) {
    if (!grp || !b) return b;
    const r = grp.tx(b); r.rot = ((b.rot || 0) + (grp.rot || 0)) % 360; return r;
  }

  function placeholderChain(sp, env) {
    const info = phInfo(sp); if (!info || env.level !== 'slide') return { info, lay: null, mas: null };
    const lay = findPh(env.layout, info);
    const mInfo = lay ? phInfo(lay) || info : info;
    const mas = findPh(env.master, { type: mInfo.type, idx: null });
    return { info, lay, mas };
  }

  async function shapeEl(sp, env, grp, isCxn) {
    const T = env.T;
    const spPr = kid(sp, 'spPr');
    const { info, lay, mas } = placeholderChain(sp, env);
    let xf = xfrmOf(spPr) || (lay && xfrmOf(kid(lay, 'spPr'))) || (mas && xfrmOf(kid(mas, 'spPr')));
    if (!xf) { if (info) xf = defaultPhBox(info, env.G); else return null; }
    xf = applyGroup(xf, grp);
    const prst = desc(spPr, 'prstGeom'); const cust = desc(spPr, 'custGeom');
    let kind = prst ? attr(prst, 'prst') : cust ? 'custom' : (lay && desc(kid(lay, 'spPr'), 'prstGeom') ? attr(desc(kid(lay, 'spPr'), 'prstGeom'), 'prst') : 'rect');
    const KIND = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', rtTriangle: 'rtTriangle', diamond: 'diamond', parallelogram: 'parallelogram', trapezoid: 'trapezoid', pentagon: 'pentagon', hexagon: 'hexagon', octagon: 'octagon', star4: 'star4', star5: 'star5', star6: 'star6', star7: 'star6', star8: 'star6', plus: 'plus', mathPlus: 'plus', donut: 'donut', heart: 'heart', cloud: 'cloud', can: 'can', rightArrow: 'rightArrow', leftArrow: 'leftArrow', upArrow: 'upArrow', downArrow: 'downArrow', leftRightArrow: 'leftRightArrow', chevron: 'chevron', homePlate: 'homePlate', wedgeRectCallout: 'wedgeRectCallout', wedgeRoundRectCallout: 'wedgeRoundRectCallout', wedgeEllipseCallout: 'wedgeEllipseCallout', line: 'line', straightConnector1: 'line', bentConnector2: 'line', bentConnector3: 'line', curvedConnector3: 'line', flowChartProcess: 'rect', flowChartAlternateProcess: 'roundRect', flowChartDecision: 'diamond', flowChartTerminator: 'roundRect', flowChartConnector: 'ellipse', snip1Rect: 'rect', snip2SameRect: 'rect', round1Rect: 'roundRect', round2SameRect: 'roundRect', frame: 'rect', flowChartInputOutput: 'parallelogram', ellipseRibbon: 'rect', notchedRightArrow: 'rightArrow', stripedRightArrow: 'rightArrow', cube: 'rect', bevel: 'rect', plaque: 'roundRect', teardrop: 'ellipse', pie: 'ellipse', chord: 'ellipse', blockArc: 'donut', sun: 'star6', moon: 'ellipse', smileyFace: 'ellipse', lightningBolt: 'star4', irregularSeal1: 'star6', irregularSeal2: 'star6', flowChartManualInput: 'trapezoid', flowChartDocument: 'rect', flowChartPredefinedProcess: 'rect', flowChartPreparation: 'hexagon', borderCallout1: 'rect', cloudCallout: 'cloud' };
    let path = null;
    if (kind === 'custom') { path = custPath(cust); if (!path) kind = 'rect'; }
    else kind = KIND[kind] || 'rect';
    if (isCxn) kind = 'line';
    const isLine = kind === 'line';

    // fill / line with style refs and placeholder inheritance
    const fRef = styleRef(sp, 'fillRef', T), lRef = styleRef(sp, 'lnRef', T), fontRef = styleRef(sp, 'fontRef', T);
    let fill = fillOf(spPr, T);
    if (fill === undefined && lay) fill = fillOf(kid(lay, 'spPr'), T);
    if (fill === undefined && fRef && fRef.idx > 0) { const sty = T.fillStyles[fRef.idx - 1]; fill = sty ? fillOf({ children: [sty] }, T, fRef.col) : (fRef.col ? { css: cssColor(fRef.col), col: fRef.col } : undefined); }
    if (fill && fill.grp && grp && grp.fill) fill = fillOf(grp.fill, T);
    const lnN = kid(spPr, 'ln') || (lay && kid(kid(lay, 'spPr'), 'ln'));
    let stroke = 'none', strokeWidth = 0, dash, arrow = false;
    const lnStyle = lRef && lRef.idx > 0 ? T.lnStyles[lRef.idx - 1] : null;
    const lnSrc = [lnN, lnStyle].filter(Boolean);
    let lnNone = false;
    for (const L of lnSrc) {
      if (kid(L, 'noFill')) { lnNone = true; break; }
      const sf = kid(L, 'solidFill'); if (sf) { const c = colorOf(colChild(sf), T, lRef && lRef.col); if (c) { stroke = cssColor(c); } }
      if (stroke !== 'none') { strokeWidth = attr(L, 'w') ? +attr(L, 'w') / 12700 : (strokeWidth || .75); const pd = desc(L, 'prstDash'); if (pd) dash = { dash: 'dash', sysDash: 'dash', lgDash: 'lgDash', dot: 'dot', sysDot: 'dot', dashDot: 'dashDot', lgDashDot: 'dashDot', sysDashDot: 'dashDot' }[attr(pd, 'val')]; break; }
    }
    if (!lnNone && stroke === 'none' && lRef && lRef.idx > 0 && lRef.col) { stroke = cssColor(lRef.col); strokeWidth = strokeWidth || (lnStyle && attr(lnStyle, 'w') ? +attr(lnStyle, 'w') / 12700 : .75); }
    if (lnNone) { stroke = 'none'; strokeWidth = 0; }
    if (lnN) { const te = kid(lnN, 'tailEnd'), he = kid(lnN, 'headEnd'); if ((te && attr(te, 'type') && attr(te, 'type') !== 'none') || (he && attr(he, 'type') && attr(he, 'type') !== 'none')) arrow = true; if (lnN && attr(lnN, 'w') && stroke !== 'none') strokeWidth = +attr(lnN, 'w') / 12700; }

    const txBody = kid(sp, 'txBody');
    const isTitle = info && TITLE_TYPES.has(info.type);
    const el = SL.newElement(isLine ? 'shape' : 'text', { x: xf.x, y: xf.y, w: xf.w, h: xf.h, rot: Math.round(((xf.rot || 0) % 360 + 360) % 360 * 10) / 10 });
    if (xf.flipH) el.flipH = true; if (xf.flipV) el.flipV = true;
    const nm = desc(sp, 'cNvPr'); if (nm) { const n = attr(nm, 'name'); if (n) el.name = n; const d = attr(nm, 'descr'); if (d) el.alt = d; }
    if (isLine) {
      el.kind = arrow ? 'arrowLine' : 'line'; el.fill = 'none'; el.stroke = stroke === 'none' ? (lRef && lRef.col ? cssColor(lRef.col) : '#000000') : stroke; el.strokeWidth = strokeWidth || 1; if (dash) el.dash = dash;
      if (lnN) { const he = kid(lnN, 'headEnd'); if (he && attr(he, 'type') && attr(he, 'type') !== 'none' && !(kid(lnN, 'tailEnd') && attr(kid(lnN, 'tailEnd'), 'type') !== 'none')) { el.flipH = !el.flipH; el.flipV = !el.flipV; } }
      return el;
    }
    const isPlainBox = kind === 'rect' && (fill === undefined || fill === 'none') && stroke === 'none';
    el.type = isPlainBox || info ? 'text' : 'shape';
    if (el.type === 'shape' || kind !== 'rect') { el.type = 'shape'; el.kind = kind; if (path) el.path = path; }
    if (kind === 'roundRect') { const gd = desc(prst, 'gd'); if (gd) { const m = /val (\d+)/.exec(attr(gd, 'fmla') || ''); if (m) el.adj = +m[1] / 100000; } }
    el.fill = fill === undefined || fill === 'none' ? 'none' : fill.grad ? fill.grad : fill.css || 'none';
    if (fill && fill.blip) { // picture-filled shape → image element
      const b = kid(fill.blip, 'blip'); const r = env.rels[rid(b)]; const url = r && await env.C.dataURL(r.target);
      if (url) return Object.assign(SL.newElement('image', { x: el.x, y: el.y, w: el.w, h: el.h, rot: el.rot, src: SL.addAsset(url), radius: kind === 'ellipse' ? 50 : kind === 'roundRect' ? 16 : 0 }), crop(fill.blip));
      el.fill = 'none';
    }
    el.stroke = stroke; el.strokeWidth = strokeWidth; if (dash) el.dash = dash;
    if (desc(spPr, 'outerShdw') || (desc(kid(sp, 'style'), 'effectRef') && +attr(desc(kid(sp, 'style'), 'effectRef'), 'idx') > 1 && false)) el.shadow = true;
    if (info) {
      el.ph = isTitle ? 'title' : info.type === 'subTitle' ? 'subtitle' : (info.type === 'body' || info.type === 'obj') ? 'body' : null;
      if (['dt', 'ftr', 'sldNum', 'hdr'].includes(info.type)) el.ph = null;
      if (!el.ph) delete el.ph;
      el.role = isTitle ? 'title' : 'body';
      if (el.ph && env.level === 'slide') el.placeholder = isTitle ? 'Click to add title' : el.ph === 'subtitle' ? 'Click to add subtitle' : 'Click to add text';
      if (['pic', 'tbl', 'chart', 'media', 'clipArt', 'dgm'].includes(info.type) && !txBody) return null;
    }
    // text
    if (txBody) readText(el, txBody, { sp, info, lay, mas, env, fontRefCol: fontRef && fontRef.col, isTitle, kind: el.type });
    if (info && ['dt', 'ftr', 'sldNum'].includes(info.type) && SL.isEmptyText(el)) return null;
    if (el.type === 'text' && SL.isEmptyText(el) && !el.ph && el.fill === 'none' && el.stroke === 'none') return null;
    return el;
  }
  function defaultPhBox(info, G) {
    const W = G.W, H = G.H;
    if (TITLE_TYPES.has(info.type)) return { x: W * .07, y: H * .05, w: W * .86, h: H * .16 };
    if (info.type === 'subTitle') return { x: W * .15, y: H * .55, w: W * .7, h: H * .2 };
    return { x: W * .07, y: H * .25, w: W * .86, h: H * .65 };
  }
  function crop(blipFill) {
    const sr = kid(blipFill, 'srcRect'); if (!sr) return {};
    const g = (k) => Math.max(0, (+attr(sr, k) || 0) / 100000);
    const c = { l: g('l'), t: g('t'), r: g('r'), b: g('b') };
    return (c.l || c.t || c.r || c.b) && c.l + c.r < .98 && c.t + c.b < .98 ? { crop: c } : {};
  }

  function readText(el, txBody, o) {
    const { env, info, lay, mas, isTitle } = o; const T = env.T;
    const bodyPr = kid(txBody, 'bodyPr');
    const layBody = lay && kid(lay, 'txBody'), masBody = mas && kid(mas, 'txBody');
    const bp = (k) => { for (const b of [bodyPr, layBody && kid(layBody, 'bodyPr'), masBody && kid(masBody, 'bodyPr')]) { if (b && attr(b, k) != null) return attr(b, k); } return null; };
    const anchor = bp('anchor'); el.valign = anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top';
    if (!anchor && o.kind === 'shape') el.valign = 'middle';
    const ins = ['lIns', 'tIns', 'rIns', 'bIns'].map((k, i) => { const v = bp(k); return v != null ? px(v) : [9.6, 4.8, 9.6, 4.8][i]; });
    el.inset = ins;
    const vert = bp('vert'); if (vert === 'vert' || vert === 'eaVert') el.dir = 'vert'; else if (vert === 'vert270') el.dir = 'vert270';
    if (bp('wrap') === 'none') el.wrap = false;
    let fontScale = 1;
    const fitN = [bodyPr, layBody && kid(layBody, 'bodyPr'), masBody && kid(masBody, 'bodyPr')].map(b => b && (kid(b, 'normAutofit') || kid(b, 'spAutoFit') || kid(b, 'noAutofit'))).find(Boolean);
    if (fitN && ln(fitN) === 'normAutofit') { el.autofit = 'shrink'; if (attr(fitN, 'fontScale')) fontScale = +attr(fitN, 'fontScale') / 100000; }
    if (fitN && ln(fitN) === 'spAutoFit') el.autofit = 'resize';
    // style chains (highest priority first)
    const master = env.txStyles;
    const masterStyle = !info ? (master && kid(master, 'otherStyle')) : isTitle ? (master && kid(master, 'titleStyle')) : (['body', 'obj', 'subTitle'].includes(info.type) ? master && kid(master, 'bodyStyle') : master && kid(master, 'otherStyle'));
    const chain = o.table ? [kid(txBody, 'lstStyle')].filter(Boolean) : [kid(txBody, 'lstStyle'), layBody && kid(layBody, 'lstStyle'), masBody && kid(masBody, 'lstStyle'), masterStyle, !info ? env.G.defaultTx : null].filter(Boolean);
    const ownLst = kid(txBody, 'lstStyle');
    const paras = kids(txBody, 'p');
    const outBlocks = []; let first = null;
    const fieldsSeen = [];
    paras.forEach(p => {
      const pPr = kid(p, 'pPr'); const lvl = +(attr(pPr, 'lvl') || 0);
      const props = mergedPPr(chain, lvl); if (pPr) applyPPr(props, pPr);
      const baseR = props.rpr || {};
      // a shape's style fontRef colour beats inherited master/default list styles (but not its own lstStyle or run props)
      if (o.fontRefCol) { const own = ownLst && lvlNode(ownLst, lvl); const ownR = own && kid(own, 'defRPr'); if (!(ownR && (kid(ownR, 'solidFill') || kid(ownR, 'noFill')))) baseR.fill = o.fontRefCol; }
      const runs = [];
      for (const r of kids(p)) {
        const k = ln(r);
        if (k === 'r' || k === 'fld') {
          const rp = Object.assign({}, baseR); applyRPr(rp, kid(r, 'rPr'));
          let text = (kid(r, 't') || {}).textContent || '';
          let field = null; if (k === 'fld') { const ft = attr(r, 'type') || ''; if (/slidenum/i.test(ft)) field = 'num'; else if (/datetime/i.test(ft)) field = 'date'; fieldsSeen.push(field); }
          runs.push({ text, rp, field });
        } else if (k === 'br') runs.push({ br: true, rp: Object.assign({}, baseR, applyRPr({}, kid(r, 'rPr'))) });
      }
      const endRp = Object.assign({}, baseR); applyRPr(endRp, kid(p, 'endParaRPr'));
      if (!first) first = runs.find(r => !r.br && r.text) ? runs.find(r => !r.br && r.text).rp : (runs[0] ? runs[0].rp : endRp);
      outBlocks.push({ props, lvl, runs, endRp });
    });
    // element defaults from first run
    const resolveColor = (rp) => { const fc = rp.fill === undefined ? (o.fontRefCol ? o.fontRefCol : null) : rp.fill; if (!fc) return null; return fc.hex ? fc : colorOf(fc, T); };
    const fr = first || {};
    const size0 = (fr.sz || (isTitle ? 44 : 18)) * fontScale;
    el.size = Math.round(size0 * 2) / 2;
    const font0Raw = fr.font; const font0 = T.fontName(font0Raw);
    const fromTheme = !font0Raw || /^\+m[jn]/.test(font0Raw);
    if (!fromTheme || !font0) el.font = font0 || (isTitle ? T.major : T.minor);
    else el.font = font0; // keep explicit: theme fonts of imported decks live in custom theme too
    const c0 = resolveColor(fr); if (c0) el.color = cssColor(c0);
    if (fr.b) el.bold = true; else if (isTitle) el.bold = false;
    if (fr.i) el.italic = true;
    const al0 = outBlocks[0] && outBlocks[0].props.algn;
    el.align = al0 === 'ctr' ? 'center' : al0 === 'r' ? 'right' : al0 === 'just' || al0 === 'dist' ? 'justify' : 'left';
    const ls = outBlocks[0] && outBlocks[0].props.lnSpc; if (ls && Math.abs(ls - 1) > .02) el.lh = Math.round(ls * 100) / 100 / (ls > 1.3 ? 1 : 1);
    const sa = outBlocks[0] && outBlocks[0].props.spcAft; if (sa) el.pspace = sa;
    else { const sb = outBlocks[1] && outBlocks[1].props.spcBef; if (sb) el.pspace = sb; }
    // build html
    let html = ''; let listOpen = null;
    const closeList = () => { if (listOpen) { html += `</${listOpen}>`; listOpen = null; } };
    for (const b of outBlocks) {
      const isBullet = b.props.bu === 'char' || b.props.bu === 'num';
      const al = b.props.algn === 'ctr' ? 'center' : b.props.algn === 'r' ? 'right' : b.props.algn === 'just' ? 'justify' : 'left';
      const st = al !== el.align ? ` style="text-align:${al}"` : '';
      const lv = b.lvl ? ` data-lvl="${b.lvl}"` : '';
      let inner = '';
      for (const r of b.runs) {
        if (r.br) { inner += '<br>'; continue; }
        let t = esc(r.text); if (r.field) t = `<span data-field="${r.field}">${t}</span>`;
        if (r.rp.cap === 'all') t = t.toUpperCase();
        const css = [];
        const sz = r.rp.sz ? Math.round(r.rp.sz * fontScale * 2) / 2 : null; if (sz && Math.abs(sz - el.size) > .01) css.push(`font-size:${sz}pt`);
        const f = T.fontName(r.rp.font); if (f && f !== el.font) css.push(`font-family:${SL.fontStack(f)}`);
        const c = resolveColor(r.rp); if (c && cssColor(c) !== el.color) css.push(`color:${cssColor(c)}`);
        if (r.rp.hl) { const h = colorOf(r.rp.hl, T); if (h) css.push(`background-color:${h.hex}`); }
        if (r.rp.spc) css.push(`letter-spacing:${r.rp.spc}pt`);
        if (css.length) t = `<span style="${css.join(';')}">${t}</span>`;
        if (r.rp.b && !el.bold) t = `<b>${t}</b>`; else if (!r.rp.b && el.bold) t = `<span style="font-weight:normal">${t}</span>`;
        if (r.rp.i && !el.italic) t = `<i>${t}</i>`; else if (!r.rp.i && el.italic) t = `<span style="font-style:normal">${t}</span>`;
        if (r.rp.u) t = `<u>${t}</u>`; if (r.rp.strike) t = `<s>${t}</s>`;
        if (r.rp.baseline > 0) t = `<sup>${t}</sup>`; else if (r.rp.baseline < 0) t = `<sub>${t}</sub>`;
        if (r.rp.link) { const id = rid(r.rp.link, 'r:id'); const rel = id && env.rels[id]; if (rel && rel.external) t = `<a href="${esc(rel.target)}">${t}</a>`; }
        inner += t;
      }
      if (!inner) inner = '<br>';
      if (isBullet) {
        const kind = b.props.bu === 'num' ? 'ol' : 'ul';
        if (listOpen !== kind) { closeList(); html += `<${kind}>`; listOpen = kind; }
        html += `<li${lv}${st}>${inner}</li>`;
      } else { closeList(); html += `<p${lv}${st}>${inner}</p>`; }
    }
    closeList();
    // drop a single empty paragraph
    if (outBlocks.length === 1 && !outBlocks[0].runs.some(r => r.text)) html = '';
    if (outBlocks.length && outBlocks.every(b => !b.runs.some(r => r.text))) html = '';
    el.html = html;
    if (!html && el.ph === 'body') { const b0 = outBlocks[0]; if (b0 && (b0.props.bu === 'char' || b0.props.bu === 'num')) el.list = b0.props.bu === 'num' ? 'ol' : 'ul'; }
    if (!html && el.ph) { const bs = (outBlocks[0] && outBlocks[0].endRp.sz) ? Math.round(outBlocks[0].endRp.sz * fontScale * 2) / 2 : null; if (bs) el.size = bs; }
  }

  async function picEl(pic, env, grp) {
    const spPr = kid(pic, 'spPr');
    const { info, lay, mas } = placeholderChain(pic, env);
    let xf = xfrmOf(spPr) || (lay && xfrmOf(kid(lay, 'spPr'))) || (mas && xfrmOf(kid(mas, 'spPr')));
    if (!xf) return null; xf = applyGroup(xf, grp);
    const bf = kid(pic, 'blipFill'); const blip = bf && kid(bf, 'blip');
    const r = blip && env.rels[rid(blip)];
    let url = r && await env.C.dataURL(r.target);
    // prefer SVG (asvg:svgBlip extension) when present
    const svgB = blip && descAll(blip, 'svgBlip')[0]; if (svgB) { const r2 = env.rels[rid(svgB)]; const u2 = r2 && await env.C.dataURL(r2.target); if (u2) url = u2; }
    if (!url) return null;
    const el = SL.newElement('image', Object.assign({ x: xf.x, y: xf.y, w: xf.w, h: xf.h, rot: Math.round(((xf.rot || 0) % 360 + 360) % 360), src: SL.addAsset(url) }, crop(bf)));
    if (xf.flipH) el.flipH = true; if (xf.flipV) el.flipV = true;
    const nm = desc(pic, 'cNvPr'); if (nm) { if (attr(nm, 'descr')) el.alt = attr(nm, 'descr'); if (attr(nm, 'name') && !/^Picture \d+$|^Image \d+$/.test(attr(nm, 'name'))) el.name = attr(nm, 'name'); }
    const prst = desc(spPr, 'prstGeom'); if (prst) { const k = attr(prst, 'prst'); if (k === 'ellipse') el.radius = 50; if (k === 'roundRect') el.radius = 16; }
    const lnN = kid(spPr, 'ln'); if (lnN && !kid(lnN, 'noFill')) { const c = colorOf(colChild(kid(lnN, 'solidFill')), env.T); if (c) { el.stroke = cssColor(c); el.strokeWidth = attr(lnN, 'w') ? +attr(lnN, 'w') / 12700 : 1; } }
    const alpha = blip && desc(blip, 'alphaModFix'); if (alpha) el.opacity = +attr(alpha, 'amt') / 100000;
    if (desc(spPr, 'outerShdw')) el.shadow = true;
    if (info && env.level === 'slide' && info.type === 'pic') el.ph = 'pic';
    return el;
  }

  async function frameEl(fr, env, grp) {
    const xfN = kid(fr, 'xfrm'); if (!xfN) return null;
    const off = kid(xfN, 'off'), ext = kid(xfN, 'ext');
    let xf = applyGroup({ x: px(attr(off, 'x')), y: px(attr(off, 'y')), w: px(attr(ext, 'cx')), h: px(attr(ext, 'cy')) }, grp);
    const gd = desc(fr, 'graphicData'); if (!gd) return null;
    const uri = attr(gd, 'uri') || '';
    if (uri.endsWith('/table')) return tableEl(desc(gd, 'tbl'), xf, env);
    if (uri.endsWith('/chart')) {
      const c = desc(gd, 'chart'); const r = c && env.rels[rid(c, 'r:id')]; if (!r) return null;
      const cx = await env.C.xml(r.target); if (!cx) return null;
      const ch = chartFrom(cx); if (!ch) return null;
      return SL.newElement('chart', { x: xf.x, y: xf.y, w: xf.w, h: xf.h, chart: ch });
    }
    // OLE objects/diagrams: use the fallback picture if any
    const pic = desc(fr, 'pic'); if (pic) { const el = await picEl(pic, env, grp); if (el) { Object.assign(el, { x: xf.x, y: xf.y, w: xf.w, h: xf.h }); return el; } }
    return null;
  }
  function tableEl(tbl, xf, env) {
    if (!tbl) return null;
    const T = env.T;
    const cols = kids(kid(tbl, 'tblGrid'), 'gridCol').map(g => +attr(g, 'w') || 1);
    const totW = cols.reduce((a, b) => a + b, 0) || 1;
    const trs = kids(tbl, 'tr'); const rowsH = trs.map(r => +attr(r, 'h') || 0); const totH = rowsH.reduce((a, b) => a + b, 0);
    const tp = kid(tbl, 'tblPr');
    let anyFill = false; let size0 = null;
    const rows = trs.map((tr) => kids(tr, 'tc').map(tc => {
      const cell = { html: '' };
      if (attr(tc, 'hMerge') === '1' || attr(tc, 'vMerge') === '1') cell.merged = true;
      if (+attr(tc, 'gridSpan') > 1) cell.cs = +attr(tc, 'gridSpan'); if (+attr(tc, 'rowSpan') > 1) cell.rs = +attr(tc, 'rowSpan');
      const tcPr = kid(tc, 'tcPr');
      if (tcPr) { const f = fillOf(tcPr, T); if (f && f.css) { cell.fill = f.css; anyFill = true; } else if (f === 'none') { cell.fill = 'none'; } const an = attr(tcPr, 'anchor'); if (an) cell.valign = an === 'ctr' ? 'middle' : an === 'b' ? 'bottom' : 'top'; }
      const tx = kid(tc, 'txBody');
      if (tx) {
        const tmp = SL.newElement('text', {}); readText(tmp, tx, { env, info: null, lay: null, mas: null, isTitle: false, kind: 'text', table: true });
        cell.html = tmp.html; if (tmp.align && tmp.align !== 'left') cell.align = tmp.align;
        if (tmp.color) cell.color = tmp.color; if (tmp.bold) cell.bold = true;
        if (tmp.size && size0 == null && tmp.html) size0 = tmp.size;
        if (tmp.size && size0 && Math.abs(tmp.size - size0) > .1 && tmp.html) cell.html = cell.html.replace(/<(p|li)([^>]*)>/g, `<$1$2><span style="font-size:${tmp.size}pt">`).replace(/<\/(p|li)>/g, '</span></$1>');
      }
      return cell;
    }));
    const el = SL.newElement('table', { x: xf.x, y: xf.y, w: xf.w, h: xf.h, rows, colW: cols.map(c => c / totW), size: size0 || 14, header: attr(tp, 'firstRow') === '1' && !anyFill, banded: attr(tp, 'bandRow') === '1' && !anyFill });
    if (totH) el.rowH = rowsH.map(h => h / totH);
    // without explicit fills and without a style flag, use a neutral look
    if (!anyFill && !el.header && !el.banded) el.rows.forEach(r => r.forEach(c => { if (!c.fill) c.fill = '#ffffff'; }));
    // normalise width of rows (pad missing cells)
    const nc = cols.length; el.rows.forEach(r => { while (r.length < nc) r.push({ html: '' }); r.length = nc; });
    return el;
  }
  function chartFrom(cx) {
    const plot = desc(cx, 'plotArea'); if (!plot) return null;
    const kinds = [['barChart', 'bar'], ['bar3DChart', 'bar'], ['lineChart', 'line'], ['line3DChart', 'line'], ['pieChart', 'pie'], ['pie3DChart', 'pie'], ['doughnutChart', 'doughnut'], ['areaChart', 'line'], ['scatterChart', 'line']];
    let node = null, type = 'bar';
    for (const [k, t] of kinds) { node = kid(plot, k); if (node) { type = t; break; } }
    if (!node) return null;
    const pts = (n) => n ? kids(desc(n, 'ptCount') ? desc(n, 'ptCount').parentNode : n, 'pt').sort((a, b) => +attr(a, 'idx') - +attr(b, 'idx')).map(p => (kid(p, 'v') || {}).textContent) : [];
    const series = kids(node, 'ser').map((s, i) => {
      const tx = kid(s, 'tx'); const name = tx ? (desc(tx, 'v') || {}).textContent || ('Series ' + (i + 1)) : 'Series ' + (i + 1);
      const cat = kid(s, 'cat') || kid(s, 'xVal'); const val = kid(s, 'val') || kid(s, 'yVal');
      const cache = (n) => n && (desc(n, 'strCache') || desc(n, 'numCache') || desc(n, 'strLit') || desc(n, 'numLit'));
      return { name, cats: pts(cache(cat)), values: pts(cache(val)).map(v => +v || 0) };
    });
    if (!series.length) return null;
    const title = desc(kid(desc(cx, 'chart'), 'title'), 't');
    return { type, title: title ? descAll(kid(desc(cx, 'chart'), 'title'), 't').map(t => t.textContent).join('') : '', categories: series[0].cats.length ? series[0].cats : series[0].values.map((_, i) => String(i + 1)), series: series.map(s => ({ name: s.name, values: s.values })), legend: !!desc(cx, 'legend') };
  }
})();
