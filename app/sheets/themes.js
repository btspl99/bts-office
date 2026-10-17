/* BTS Sheets — workbook themes (fonts + 6-accent palettes), font loading for the canvas grid,
 * the cell style gallery and the table style catalogue. Everything visual that follows the theme reads from here. */
(function () {
  const T = window.STheme = {};

  /* ---------- themes ---------- */
  const mk = (id, name, head, body, accents, dk2, lt2, extra = {}) => ({ id, name, head, body, accents, dk1: '#000000', lt1: '#FFFFFF', dk2, lt2, hlink: '#0563C1', folHlink: '#954F72', ...extra });
  T.THEMES = [
    mk('bts', 'BTS', 'Calibri', 'Calibri', ['#0F8A6A', '#2F5597', '#E07A1F', '#7A4FB5', '#D64545', '#2BA3B5'], '#1F3A4D', '#E8EEF0', { desc: 'Calibri · teal (default)' }),
    mk('office', 'Office', 'Calibri', 'Calibri', ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'], '#44546A', '#E7E6E6', { desc: 'Calibri · classic' }),
    mk('corporate', 'Corporate', 'Inter', 'Inter', ['#1D4ED8', '#0F766E', '#F59E0B', '#7C3AED', '#DC2626', '#0891B2'], '#0F172A', '#E2E8F0', { desc: 'Inter · navy' }),
    mk('slate', 'Slate', 'IBM Plex Sans', 'IBM Plex Sans', ['#475569', '#2563EB', '#0D9488', '#D97706', '#9333EA', '#E11D48'], '#1E293B', '#E5E7EB', { desc: 'IBM Plex Sans · slate' }),
    mk('academic', 'Academic', 'Source Serif 4', 'Source Sans 3', ['#7F1D1D', '#1E3A8A', '#166534', '#92400E', '#4C1D95', '#0F766E'], '#3B1F1F', '#F1ECE4', { desc: 'Source Serif 4 / Source Sans 3' }),
    mk('modern', 'Modern', 'Montserrat', 'Open Sans', ['#111827', '#F43F5E', '#06B6D4', '#8B5CF6', '#F59E0B', '#10B981'], '#111827', '#F3F4F6', { desc: 'Montserrat / Open Sans' }),
    mk('fresh', 'Fresh', 'Nunito', 'Nunito', ['#16A34A', '#3B82F6', '#F97316', '#EC4899', '#14B8A6', '#EAB308'], '#14532D', '#ECFDF5', { desc: 'Nunito · bright' }),
    mk('mono', 'Mono', 'Roboto', 'Roboto', ['#374151', '#2563EB', '#059669', '#DC2626', '#7C3AED', '#D97706'], '#111827', '#F3F4F6', { num: 'Roboto Mono', desc: 'Roboto · Roboto Mono numbers' }),
    mk('earth', 'Earth', 'Lora', 'Lato', ['#8B5E34', '#6B8E23', '#C2703D', '#556B2F', '#A0522D', '#B08D57'], '#3E2C1C', '#F3EDE3', { desc: 'Lora / Lato · warm' }),
    mk('ocean', 'Ocean', 'Poppins', 'DM Sans', ['#03396C', '#0077B6', '#2A9D8F', '#00B4D8', '#E76F51', '#6497B1'], '#012A4A', '#E6F2F8', { desc: 'Poppins / DM Sans' }),
    mk('saffron', 'Saffron', 'Poppins', 'Noto Sans', ['#FF9933', '#138808', '#000080', '#C2185B', '#D4A017', '#6D4C41'], '#3A2A12', '#FFF5E6', { desc: 'Indian palette · Noto Sans' }),
    mk('elegant', 'Elegant', 'Playfair Display', 'Lato', ['#1F2937', '#B45309', '#7C2D12', '#365314', '#1E3A8A', '#6B21A8'], '#1F2937', '#F5F0E6', { desc: 'Playfair Display / Lato' }),
  ];
  T.get = (id) => T.THEMES.find(t => t.id === id) || T.THEMES[0];
  let curCache = { id: null, over: null, th: null };
  /** the workbook's theme: a built-in theme plus optional colour/font overrides (wb.themeOver) */
  T.cur = () => {
    const id = window.WB ? WB.theme : 'bts', over = window.WB ? WB.themeOver : null;
    if (curCache.id === id && curCache.over === over && curCache.th) return curCache.th;
    const b = T.get(id); const th = over ? { ...b, ...over, name: b.name + ' (custom)' } : b;
    curCache = { id, over, th }; return th;
  };
  T.byFonts = (head, body, accents) => T.THEMES.find(t => t.head === head && t.body === body && (!accents || accents.every((a, i) => a.toUpperCase() === t.accents[i].toUpperCase())));

  /* ---------- colour helpers ---------- */
  const hex = (c) => { c = String(c || '#000').replace('#', ''); if (c.length === 3) c = [...c].map(x => x + x).join(''); return [0, 2, 4].map(i => parseInt(c.slice(i, i + 2), 16) || 0); };
  const toHex = (a) => '#' + a.map(x => Math.max(0, Math.min(255, Math.round(x))).toString(16).padStart(2, '0')).join('').toUpperCase();
  /** t > 0 lightens toward white (Excel tint), t < 0 darkens */
  T.tint = (c, t) => { const a = hex(c); return toHex(a.map(x => t < 0 ? x * (1 + t) : x + (255 - x) * t)); };
  T.lum = (c) => { const [r, g, b] = hex(c); return (0.299 * r + 0.587 * g + 0.114 * b) / 255; };
  T.textOn = (c) => T.lum(c) > 0.6 ? '#000000' : '#FFFFFF';

  /* ---------- fonts ---------- */
  const FONTS = window.BTS_FONTS || [];
  const ALIAS = {}; FONTS.forEach(f => { if (f.alias) ALIAS[f.alias] = f.name; });
  T.ALIAS = ALIAS;
  T.fontNames = () => FONTS.map(f => f.name);
  const cssCache = new Map();
  /** CSS font-family list for a cell font: the font itself, its bundled metric-compatible twin, Carlito, then a generic family */
  SU.cssFont = (name) => {
    name = name || SU.defFont; let v = cssCache.get(name); if (v) return v;
    const f = FONTS.find(x => x.name === name) || (ALIAS[name] && FONTS.find(x => x.name === ALIAS[name]));
    const gen = f ? (window.BTS_FONT_FALLBACK || {})[f.cat] || 'sans-serif' : 'sans-serif';
    const parts = ['"' + name.replace(/"/g, '') + '"']; if (ALIAS[name]) parts.push('"' + ALIAS[name] + '"');
    parts.push('Carlito', gen); v = parts.join(', '); cssCache.set(name, v); return v;
  };
  SU.FONTS = T.fontNames();
  Object.defineProperty(SU, 'defFont', { get: () => T.cur().body, configurable: true });
  /** default cell font size (11; a new workbook can start bigger or smaller: File > Options > New documents) */
  Object.defineProperty(SU, 'defSize', { get: () => (window.WB && +WB.defSize) || 11, configurable: true });
  Object.defineProperty(SU, 'numFont', { get: () => T.cur().num || null, configurable: true });
  /* Canvas text only uses a web font once it is loaded. We load each face on first use, then redraw. */
  const ready = new Set(), pending = new Map(); let redrawT = 0;
  const faceKey = (n, b, i) => n + '|' + (b ? 1 : 0) + (i ? 1 : 0);
  T.loadFont = (name, b, i) => {
    name = name || SU.defFont; const k = faceKey(name, b, i);
    if (ready.has(k)) return Promise.resolve(true);
    if (pending.has(k)) return pending.get(k);
    const fam = ['"' + name + '"']; if (ALIAS[name]) fam.push('"' + ALIAS[name] + '"');
    const spec = (i ? 'italic ' : '') + (b ? 'bold ' : '') + '16px ' + fam.join(', ');
    let p; try { p = document.fonts.load(spec, 'Aa1₹').catch(() => []); } catch { p = Promise.resolve([]); }
    p = p.then(() => { ready.add(k); pending.delete(k); clearTimeout(redrawT); redrawT = setTimeout(() => { T.onFontsLoaded && T.onFontsLoaded(); }, 16); return true; });
    pending.set(k, p); return p;
  };
  /** true when the face is ready to paint; otherwise starts loading it */
  T.fontReady = (name, b, i) => { const k = faceKey(name || SU.defFont, b, i); if (ready.has(k)) return true; T.loadFont(name, b, i); return false; };
  /** preload every font used by a workbook (styles + theme), resolves when all are ready */
  T.preloadWorkbook = (wb) => {
    const set = new Set(); const th = T.get(wb.theme); [th.head, th.body, th.num].filter(Boolean).forEach(n => { set.add(faceKey(n, 0, 0)); set.add(faceKey(n, 1, 0)); });
    for (const s of wb.sheets) for (const st of s.styles.values()) set.add(faceKey(st.font || th.body, st.b, st.i));
    return Promise.all([...set].slice(0, 200).map(k => { const [n, f] = k.split('|'); return T.loadFont(n, f[0] === '1', f[1] === '1'); }));
  };

  /* ---------- cell styles (themed) ---------- */
  const B = (style, color) => ({ style, color });
  const box = (b) => ({ bT: b, bB: b, bL: b, bR: b });
  /** The gallery for a theme. Each entry: [name, style]. Number-format styles only set `nf`. */
  T.cellStyleGroups = (th = T.cur()) => {
    const a = th.accents;
    const themed = [];
    for (const [pct, t] of [[20, 0.8], [40, 0.6], [60, 0.4]]) a.forEach((c, i) => themed.push([`${pct}% - Accent${i + 1}`, { fill: T.tint(c, t), color: pct === 60 ? T.textOn(T.tint(c, t)) : '#000000' }]));
    a.forEach((c, i) => themed.push([`Accent${i + 1}`, { fill: c, color: T.textOn(c) }]));
    return [
      { group: 'Good, Bad and Neutral', items: [['Normal', {}], ['Bad', { fill: '#FFC7CE', color: '#9C0006' }], ['Good', { fill: '#C6EFCE', color: '#006100' }], ['Neutral', { fill: '#FFEB9C', color: '#9C5700' }]] },
      { group: 'Data and Model', items: [
        ['Calculation', { fill: '#F2F2F2', color: '#FA7D00', b: true, ...box(B('thin', '#7F7F7F')) }],
        ['Check Cell', { fill: '#A5A5A5', color: '#FFFFFF', b: true, ...box(B('double', '#3F3F3F')) }],
        ['Explanatory Text', { color: '#7F7F7F', i: true }],
        ['Input', { fill: '#FFCC99', color: '#3F3F76', ...box(B('thin', '#7F7F7F')) }],
        ['Linked Cell', { color: '#FA7D00', bB: B('double', '#FF8001') }],
        ['Note', { fill: '#FFFFCC', ...box(B('thin', '#B2B2B2')) }],
        ['Output', { fill: '#F2F2F2', color: '#3F3F3F', b: true, ...box(B('thin', '#3F3F3F')) }],
        ['Warning Text', { color: '#FF0000' }]] },
      { group: 'Titles and Headings', items: [
        ['Title', { font: th.head, sz: 18, color: th.dk2 }],
        ['Heading 1', { font: th.head, sz: 15, b: true, color: th.dk2, bB: B('thick', a[0]) }],
        ['Heading 2', { font: th.head, sz: 13, b: true, color: th.dk2, bB: B('thick', T.tint(a[0], 0.5)) }],
        ['Heading 3', { b: true, color: th.dk2, bB: B('medium', T.tint(a[0], 0.6)) }],
        ['Heading 4', { b: true, color: th.dk2 }],
        ['Total', { b: true, bT: B('thin', a[0]), bB: B('double', a[0]) }]] },
      { group: 'Themed Cell Styles', items: themed, cols: 6 },
      { group: 'Number Format', items: [['Comma', { nf: SU.accountingFormat('', 2) }], ['Comma [0]', { nf: SU.accountingFormat('', 0) }], ['Currency', { nf: SU.accountingFormat('₹', 2) }], ['Currency [0]', { nf: SU.accountingFormat('₹', 0) }], ['Percent', { nf: '0%' }]], numOnly: true },
    ];
  };
  T.NUMBER_STYLES = ['Comma', 'Comma [0]', 'Currency', 'Currency [0]', 'Percent'];
  /** style for a named cell style in the current workbook (custom styles and modifications win) */
  T.cellStyle = (name, wb = window.WB, th) => {
    if (wb && wb.styleMods && wb.styleMods[name]) return { ...wb.styleMods[name] };
    if (wb && wb.customStyles) { const c = wb.customStyles.find(x => x.name === name); if (c) return { ...c.st }; }
    for (const g of T.cellStyleGroups(th || T.cur())) { const it = g.items.find(x => x[0] === name); if (it) return { ...it[1] }; }
    return null;
  };

  /* ---------- table styles ---------- */
  /* Light 1–14, Medium 1–7, Dark 1–7: 28 styles, all computed from the theme accents (1 = neutral, 2–7 = Accent1–6). */
  T.TABLE_STYLES = [];
  for (const fam of ['Light', 'Medium', 'Dark']) { const n = fam === 'Light' ? 14 : 7; for (let i = 1; i <= n; i++) T.TABLE_STYLES.push(fam + ' ' + i); }
  const thin = (c) => ({ style: 'thin', color: c });
  /** formatting roles for a table style: head, band1, band2, body, total, first, last, lines */
  const fmtCache = new WeakMap();
  T.tableFmt = (styleId, th = T.cur()) => {
    let m = fmtCache.get(th); if (!m) { m = new Map(); fmtCache.set(th, m); }
    let f = m.get(styleId); if (!f) { f = tableFmtRaw(styleId, th); m.set(styleId, f); }
    return f;
  };
  function tableFmtRaw(styleId, th) {
    const [fam, nStr] = String(styleId || 'Medium 2').split(' '); let n = +nStr || 2;
    const variant = fam === 'Light' && n > 7 ? 2 : 1; if (n > 7) n -= 7;
    const acc = n === 1 ? (fam === 'Dark' ? '#595959' : '#262626') : th.accents[n - 2];
    if (fam === 'Light' && variant === 1) {
      const ink = n === 1 ? '#000000' : T.tint(acc, -0.25);
      return { head: { b: true, color: ink, bB: thin(acc) }, band1: { fill: n === 1 ? '#D9D9D9' : T.tint(acc, 0.8) }, band2: {}, body: { color: ink }, total: { b: true, color: ink, bT: { style: 'double', color: acc } }, first: { b: true }, last: { b: true }, outline: { top: thin(acc), bottom: thin(acc) }, swatch: [null, n === 1 ? '#D9D9D9' : T.tint(acc, 0.8), '#FFFFFF', acc] };
    }
    if (fam === 'Light') {
      return { head: { b: true, fill: acc, color: '#FFFFFF' }, band1: { fill: n === 1 ? '#EDEDED' : T.tint(acc, 0.85) }, band2: {}, body: {}, total: { b: true, bT: { style: 'double', color: acc } }, first: { b: true }, last: { b: true }, outline: { top: thin(acc), bottom: thin(acc), left: thin(acc), right: thin(acc) }, swatch: [acc, n === 1 ? '#EDEDED' : T.tint(acc, 0.85), '#FFFFFF', acc] };
    }
    if (fam === 'Medium') {
      return { head: { b: true, fill: acc, color: '#FFFFFF' }, band1: { fill: n === 1 ? '#D9D9D9' : T.tint(acc, 0.8) }, band2: { fill: n === 1 ? '#F2F2F2' : T.tint(acc, 0.9) }, body: {}, total: { b: true, fill: n === 1 ? '#D9D9D9' : T.tint(acc, 0.8), bT: { style: 'double', color: acc } }, first: { b: true }, last: { b: true }, outline: null, swatch: [acc, n === 1 ? '#D9D9D9' : T.tint(acc, 0.8), n === 1 ? '#F2F2F2' : T.tint(acc, 0.9), acc] };
    }
    const dark = T.tint(acc, -0.5);
    return { head: { b: true, fill: n === 1 ? '#000000' : dark, color: '#FFFFFF', bB: thin('#FFFFFF') }, band1: { fill: T.tint(acc, -0.25), color: '#FFFFFF' }, band2: { fill: acc, color: '#FFFFFF' }, body: { color: '#FFFFFF' }, total: { b: true, fill: n === 1 ? '#000000' : dark, color: '#FFFFFF', bT: thin('#FFFFFF') }, first: { b: true, fill: n === 1 ? '#262626' : dark }, last: { b: true, fill: n === 1 ? '#262626' : dark }, outline: null, swatch: [n === 1 ? '#000000' : dark, T.tint(acc, -0.25), acc, dark] };
  }
  /** Excel built-in name for our style id (themes are written into the file so the colours match) */
  T.excelTableStyle = (id) => { const [fam, n] = String(id || 'Medium 2').split(' '); return 'TableStyle' + fam + n; };
  T.fromExcelTableStyle = (name) => { const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name || ''); if (!m) return 'Medium 2'; const max = m[1] === 'Light' ? 14 : 7; let n = +m[2]; if (n > max) n = ((n - 1) % 7) + 1; return m[1] + ' ' + n; };
  /** legacy (1.1) tables stored concrete colours in t.theme */
  function legacyFmt(th) { return { head: { b: true, fill: th.head, color: th.headText || '#FFFFFF' }, band1: { fill: th.band }, band2: th.band2 ? { fill: th.band2 } : {}, body: {}, total: { b: true, bT: thin(th.head) }, first: { b: true }, last: { b: true }, outline: null }; }
  const legacyCache = new WeakMap();
  T.fmtOfTable = (t, th) => { if (t.styleId || !t.theme) return T.tableFmt(t.styleId || 'Medium 2', th || T.cur()); let f = legacyCache.get(t.theme); if (!f) { f = legacyFmt(t.theme); legacyCache.set(t.theme, f); } return f; };
  /** formatting for cell (r,c) inside table t: {fill, color, b, bT, bB, bL, bR} (only the keys that apply) */
  T.tableCell = (t, r, c, th) => {
    const f = T.fmtOfTable(t, th); const g = t.range; const out = {};
    const put = (o) => { if (o) for (const k in o) if (o[k] !== undefined) out[k] = o[k]; };
    const isHead = t.header !== false && r === g.r0, isTotal = t.total && r === g.r1;
    if (isHead) put(f.head);
    else if (isTotal) put(f.total);
    else {
      put(f.body);
      const i = r - g.r0 - (t.header !== false ? 1 : 0), j = c - g.c0;
      if (t.banded !== false) put(i % 2 === 0 ? f.band1 : f.band2);
      else if (t.bandedCols) put(j % 2 === 0 ? f.band1 : f.band2);
      if (t.firstCol && c === g.c0) put(f.first);
      if (t.lastCol && c === g.c1) put(f.last);
    }
    if (f.outline) { if (r === g.r0 && f.outline.top) out.bT = out.bT || f.outline.top; if (r === g.r1 && f.outline.bottom) out.bB = out.bB || f.outline.bottom; if (c === g.c0 && f.outline.left) out.bL = f.outline.left; if (c === g.c1 && f.outline.right) out.bR = f.outline.right; }
    return out;
  };
  T.swatchHTML = (id, th = T.cur()) => { const f = T.tableFmt(id, th); const [h, b1, b2, line] = f.swatch; return `<i style="background:${h || '#fff'};border-bottom:1.5px solid ${line}"></i><i style="background:${b1}"></i><i style="background:${b2 || '#fff'}"></i><i style="background:${b1}"></i>`; };

  /* ---------- chart palette ---------- */
  T.chartPalette = (th = T.cur()) => [...th.accents, T.tint(th.accents[0], -0.35), T.tint(th.accents[1], -0.35), T.tint(th.accents[2], 0.35), T.tint(th.accents[3], 0.35)];

  /* ---------- theme XML for xlsx (Excel reads fonts + colours from xl/theme/theme1.xml) ---------- */
  T.patchThemeXML = (xml, th) => {
    const cs = { dk1: th.dk1, lt1: th.lt1, dk2: th.dk2, lt2: th.lt2, accent1: th.accents[0], accent2: th.accents[1], accent3: th.accents[2], accent4: th.accents[3], accent5: th.accents[4], accent6: th.accents[5], hlink: th.hlink, folHlink: th.folHlink };
    for (const [k, v] of Object.entries(cs)) {
      const re = new RegExp(`<a:${k}>[\\s\\S]*?</a:${k}>`);
      xml = xml.replace(re, `<a:${k}><a:srgbClr val="${v.replace('#', '').toUpperCase()}"/></a:${k}>`);
    }
    xml = xml.replace(/<a:clrScheme name="[^"]*"/, `<a:clrScheme name="BTS ${th.name}"`).replace(/<a:fontScheme name="[^"]*"/, `<a:fontScheme name="BTS ${th.name}"`);
    xml = xml.replace(/(<a:majorFont>\s*<a:latin typeface=")[^"]*"/, `$1${th.head}"`).replace(/(<a:minorFont>\s*<a:latin typeface=")[^"]*"/, `$1${th.body}"`);
    return xml;
  };
  T.readThemeXML = (xml) => {
    const get = (k) => { const m = new RegExp(`<a:${k}>\\s*<a:(?:srgbClr val|sysClr[^>]*lastClr)="([0-9A-Fa-f]{6})"`).exec(xml); return m ? '#' + m[1].toUpperCase() : null; };
    const head = (/<a:majorFont>\s*<a:latin typeface="([^"]*)"/.exec(xml) || [])[1], body = (/<a:minorFont>\s*<a:latin typeface="([^"]*)"/.exec(xml) || [])[1];
    return { head, body, accents: [1, 2, 3, 4, 5, 6].map(i => get('accent' + i)), dk2: get('dk2'), lt2: get('lt2') };
  };
})();
