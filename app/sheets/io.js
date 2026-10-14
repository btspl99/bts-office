/* BTS Sheets — file I/O: xlsx (ExcelJS, SheetJS fallback), xls/ods/csv/tsv (SheetJS), export xlsx/csv/ods/html, print layout */
(function () {
  const { key, unkey } = SU;
  const IO = window.IO = {};
  const HF = window.HyperFormula;

  /* ---------- colour helpers ---------- */
  const THEME = ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#0563C1', '#954F72'];
  const INDEXED = ['#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF', '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF', '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080', '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF', '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF', '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99', '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696', '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333'];
  function tint(hex, t) {
    if (!t) return hex; let [r, g, b] = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16));
    const f = (x) => Math.round(t < 0 ? x * (1 + t) : x + (255 - x) * t);
    return '#' + [r, g, b].map(f).map(x => SU.clamp(x, 0, 255).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  function fromColor(c, fallback) {
    if (!c) return fallback;
    if (c.argb) { const a = c.argb.length === 8 ? c.argb.slice(2) : c.argb; return '#' + a.toUpperCase(); }
    if (c.theme !== undefined) return tint(THEME[c.theme] || '#000000', c.tint || 0);
    if (c.indexed !== undefined) { if (c.indexed === 64 || c.indexed === 65) return fallback; return INDEXED[c.indexed] || fallback; }
    return fallback;
  }
  const toARGB = (hex) => 'FF' + String(hex || '#000000').replace('#', '').toUpperCase().padEnd(6, '0').slice(0, 6);
  IO.fromColor = fromColor;

  /* ---------- size conversions ---------- */
  const widthToPx = (w) => Math.round(w * 7 + 5);            // Excel character width (Calibri 11)
  const pxToWidth = (px) => Math.max(0, +((px - 5) / 7).toFixed(2));
  const ptToPx = (pt) => Math.round(pt / 0.75);
  const pxToPt = (px) => +(px * 0.75).toFixed(2);

  /* ---------- workbook (re)build from data ---------- */
  /** sheets: [{name, data:[[raw]] , meta:{...}}], names:[{name, expression}] */
  IO.buildWorkbook = (wb, sheets, names = []) => {
    wb.hf.destroy();
    const obj = {}; sheets.forEach(s => { obj[s.name] = Array.from(s.data || [], row => row ? Array.from(row, v => wb._toHFContent(v === undefined ? null : v)) : []); });
    const okNames = names.filter(n => /^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(n.name) && !/^[A-Za-z]{1,3}\d+$/.test(n.name));
    try { wb.hf = HF.buildFromSheets(obj, SheetModel.HF_CONFIG, okNames.map(n => ({ name: n.name, expression: SU.toHF(n.expression[0] === '=' ? n.expression : '=' + n.expression) }))); }
    catch (e) { console.warn('names rejected', e); wb.hf = HF.buildFromSheets(obj, SheetModel.HF_CONFIG); for (const n of okNames) { try { wb.hf.addNamedExpression(n.name, SU.toHF(n.expression[0] === '=' ? n.expression : '=' + n.expression)); } catch { } } }
    wb.sheets = sheets.map(sd => { const s = new SheetModel.Sheet(wb, sd.name); s.id = wb.hf.getSheetId(sd.name); Object.assign(s, sd.meta || {}); const d = s.dims(); s.rowCount = Math.max(1000, d.height + 200, s.rowCount || 0); s.colCount = Math.max(100, d.width + 10, s.colCount || 0); s.invalidateGeom(); return s; });
    wb.activeIndex = 0; wb.undoStack = []; wb.redoStack = [];
  };

  /* ================= IMPORT ================= */
  IO.load = async (wb, file) => {
    const ext = BTS.extOf(file.name);
    if (ext === 'csv' || ext === 'tsv' || ext === 'txt') return IO.loadCSV(wb, file, ext);
    if (ext === 'xlsx' || ext === 'xlsm') {
      try { await IO.loadXLSX(wb, file.data); return; }
      catch (e) {
        try { const cleaned = await stripComments(file.data); await IO.loadXLSX(wb, cleaned, file.data); return; }
        catch (e2) { console.warn('Detailed import unavailable, using basic import', e2); }
      }
    }
    IO.loadSheetJS(wb, file.data);
  };
  function decodeText(bytes) {
    if (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) return new TextDecoder('utf-8').decode(bytes.subarray(3));
    if (bytes[0] === 0xFF && bytes[1] === 0xFE) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
    try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes); } catch { return new TextDecoder('windows-1252').decode(bytes); }
  }
  IO.decodeText = decodeText;
  IO.loadCSV = (wb, file, ext) => {
    const text = decodeText(file.data);
    const rows = SU.parseDelimited(text, ext === 'tsv' ? '\t' : null);
    const { data, styles } = IO.typedRows(rows);
    IO.buildWorkbook(wb, [{ name: (file.name.replace(/\.[^.]+$/, '') || 'Sheet1').slice(0, 31).replace(/[\\\/?*\[\]:]/g, '_'), data, meta: { styles } }]);
    autofitImported(wb.sheets[0]);
  };
  IO.typedRows = (rows, r0 = 0, c0 = 0) => {
    const data = [], styles = new Map();
    rows.forEach((row, r) => { data[r + r0] = []; row.forEach((t, c) => { const p = SU.parseInput(t); let v; if (p.formula) v = p.formula; else if (p.value === null) v = null; else v = p.text ? t.replace(/^'/, '') : p.value; data[r + r0][c + c0] = v; if (p.nf) styles.set(key(r + r0, c + c0), { nf: p.nf }); }); });
    return { data, styles };
  };
  function autofitImported(s) {
    const u = s.usedRange(); if (u.c1 < 0) return; const m = new Map();
    for (let c = 0; c <= Math.min(u.c1, 200); c++) { let w = 0; for (let r = 0; r <= Math.min(u.r1, 300); r++) { const t = s.displayText(r, c); if (t) w = Math.max(w, Grid.measure(Grid.fontStr(s.style(r, c), 1), t) + 10); } if (w > SU.DEF_W) m.set(c, Math.min(320, Math.ceil(w))); }
    s.colW = m; s.invalidateGeom();
  }

  function excelFormula(f) { return '=' + String(f).replace(/^=/, ''); }
  function cellRaw(cell) {
    const v = cell.value;
    if (v === null || v === undefined) return { v: null };
    if (typeof v === 'number' || typeof v === 'boolean') return { v };
    if (typeof v === 'string') return { v };
    if (v instanceof Date) return { v: SU.jsDateToSerial(v), date: true };
    if (typeof v === 'object') {
      if (v.formula !== undefined || v.sharedFormula !== undefined) { const f = cell.formula || v.formula; if (f) return { v: excelFormula(f), result: v.result }; return { v: v.result instanceof Date ? SU.jsDateToSerial(v.result) : (v.result ?? null) }; }
      if (v.richText) return { v: v.richText.map(x => x.text).join(''), rich: v.richText };
      if (v.hyperlink !== undefined) { const t = typeof v.text === 'object' && v.text && v.text.richText ? v.text.richText.map(x => x.text).join('') : v.text; return { v: t ?? v.hyperlink, link: v.hyperlink }; }
      if (v.error) return { v: v.error, err: true };
    }
    return { v: String(v) };
  }
  function styleFromExcel(st, rich) {
    const o = {}; if (!st) return o;
    const f = st.font;
    if (f) {
      if (f.bold) o.b = true; if (f.italic) o.i = true; if (f.underline) o.u = f.underline === 'double' ? 'double' : true; if (f.strike) o.s = true;
      if (f.name && f.name !== 'Calibri') o.font = f.name; if (f.size && f.size !== 11) o.sz = f.size;
      const col = fromColor(f.color, null); if (col && col !== '#000000') o.color = col;
    }
    if (rich && rich.length && !o.b && rich.every(x => x.font && x.font.bold)) o.b = true;
    const fi = st.fill;
    if (fi && fi.type === 'pattern' && fi.pattern && fi.pattern !== 'none') { const c = fromColor(fi.fgColor, null) || fromColor(fi.bgColor, null); if (c) o.fill = c; }
    else if (fi && fi.type === 'gradient' && fi.stops && fi.stops.length) { o.fill = fromColor(fi.stops[0].color, null) || undefined; if (!o.fill) delete o.fill; }
    const bd = st.border;
    if (bd) { const one = (b) => b && b.style ? { style: b.style, color: fromColor(b.color, '#000000') } : undefined; const t = one(bd.top), bo = one(bd.bottom), l = one(bd.left), r = one(bd.right); if (t) o.bT = t; if (bo) o.bB = bo; if (l) o.bL = l; if (r) o.bR = r; }
    const al = st.alignment;
    if (al) {
      if (al.horizontal && al.horizontal !== 'general') o.h = al.horizontal === 'centerContinuous' || al.horizontal === 'distributed' || al.horizontal === 'justify' ? 'center' : al.horizontal;
      if (al.vertical) o.v = al.vertical === 'center' || al.vertical === 'middle' ? 'middle' : al.vertical === 'top' ? 'top' : al.vertical === 'distributed' || al.vertical === 'justify' ? 'middle' : 'bottom';
      if (o.v === 'bottom') delete o.v;
      if (al.wrapText) o.wrap = true; if (al.indent) o.indent = al.indent;
    }
    if (st.numFmt && st.numFmt !== 'General') o.nf = st.numFmt;
    if (st.protection && st.protection.locked === false) o.locked = false;
    return o;
  }
  /** ExcelJS can't reconcile some legacy comment parts (e.g. files from openpyxl); remove them — notes are read with SheetJS. */
  async function stripComments(data) {
    const zip = await JSZip.loadAsync(data);
    for (const name of Object.keys(zip.files)) {
      if (/^xl\/comments\d*\.xml$|^xl\/comments\/.*\.xml$|^xl\/drawings\/.*\.vml$/i.test(name)) zip.remove(name);
      else if (/^xl\/worksheets\/_rels\/.*\.rels$/.test(name)) { let x = await zip.file(name).async('string'); x = x.replace(/<Relationship [^>]*Type="[^"]*\/(comments|vmlDrawing)"[^>]*\/>/g, ''); zip.file(name, x); }
      else if (/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) { let x = await zip.file(name).async('string'); x = x.replace(/<legacyDrawing [^>]*\/>/g, ''); zip.file(name, x); }
      else if (name === '[Content_Types].xml') { let x = await zip.file(name).async('string'); x = x.replace(/<Override [^>]*comments[^>]*\/>/gi, ''); zip.file(name, x); }
    }
    return zip.generateAsync({ type: 'uint8array' });
  }
  IO.loadXLSX = async (wb, data, original) => {
    const X = new ExcelJS.Workbook();
    await X.xlsx.load(data);
    if (original) data = original;
    let meta = null;
    const metaWs = X.worksheets.find(w => w.name === '_btsmeta');
    if (metaWs) { try { let txt = ''; metaWs.eachRow(row => { const v = row.getCell(1).value; if (typeof v === 'string') txt += v; }); meta = JSON.parse(txt); } catch (e) { console.warn('meta', e); } }
    // comments via SheetJS (ExcelJS doesn't read notes)
    let notesBy = {};
    try { const xw = XLSX.read(data, { type: 'array', cellFormula: false, cellHTML: false, cellText: false, sheetStubs: false, bookFiles: false, dense: false }); for (const name of xw.SheetNames) { const ws = xw.Sheets[name]; const m = {}; for (const a in ws) { if (a[0] === '!') continue; const c = ws[a]; if (c.c && c.c.length) m[a] = c.c.map(x => x.t).join('\n').replace(/^[^:\n]{1,40}:\n/, ''); } notesBy[name] = m; } } catch (e) { console.warn(e); }
    const sheets = [];
    for (const ws of X.worksheets) {
      if (ws.name === '_btsmeta') continue;
      const name = ws.name.slice(0, 31);
      const dataArr = [], styles = new Map(), links = new Map(), comments = new Map();
      const sm = meta && meta.sheets && meta.sheets[ws.name] || {};
      ws.eachRow({ includeEmpty: false }, (row, rn) => {
        const r = rn - 1;
        row.eachCell({ includeEmpty: true }, (cell, cn) => {
          const c = cn - 1;
          if (cell.isMerged && cell.master && cell.master !== cell) { /* style only */ }
          const rv = (cell.isMerged && cell.master !== cell) ? { v: null } : cellRaw(cell);
          if (rv.v !== null && rv.v !== '') { (dataArr[r] = dataArr[r] || [])[c] = rv.v; }
          const st = styleFromExcel(cell.style, rv.rich);
          if (rv.date && !st.nf) st.nf = 'dd/mm/yyyy';
          if (Object.keys(st).length) styles.set(key(r, c), st);
          if (rv.link || (cell.hyperlink && typeof cell.hyperlink === 'string')) links.set(key(r, c), rv.link || cell.hyperlink);
          if (cell.note) comments.set(key(r, c), typeof cell.note === 'string' ? cell.note : (cell.note.texts || []).map(t => t.text).join(''));
        });
      });
      const nb = notesBy[ws.name] || {}; for (const a in nb) { const p = SU.parseAddr(a); if (p) comments.set(key(p.r, p.c), nb[a]); }
      const colW = new Map(), hiddenC = new Set(), rowH = new Map(), hiddenR = new Set();
      (ws.columns || []).forEach((col, i) => { if (!col) return; if (col.width && Math.abs(widthToPx(col.width) - SU.DEF_W) > 1) colW.set(i, widthToPx(col.width)); if (col.hidden) hiddenC.add(i); });
      ws.eachRow({ includeEmpty: true }, (row, rn) => { if (row.height && Math.abs(ptToPx(row.height) - SU.DEF_H) > 1) rowH.set(rn - 1, ptToPx(row.height)); if (row.hidden) hiddenR.add(rn - 1); });
      // default col widths
      if (ws.properties && ws.properties.defaultColWidth) { /* ignore: keep our default */ }
      const merges = (ws.model.merges || []).map(m => SU.parseRange(m)).filter(Boolean);
      const v = (ws.views || [])[0] || {};
      const freeze = v.state === 'frozen' ? { r: v.ySplit || 0, c: v.xSplit || 0 } : { r: 0, c: 0 };
      const tabColor = ws.properties && ws.properties.tabColor ? fromColor(ws.properties.tabColor, null) : null;
      // data validation
      const validations = [];
      try {
        const groups = new Map();
        const dvm = ws.dataValidations && ws.dataValidations.model || {};
        for (const addr in dvm) {
          const d = dvm[addr]; const k = JSON.stringify(d);
          const rg = SU.parseRange(addr.replace(/^range:/, '')) || (() => { const ps = addr.split(/\s+/).map(a => SU.parseRange(a)).filter(Boolean); return ps[0]; })(); if (!rg) continue;
          let g = groups.get(k); if (!g) { g = { d, cells: [] }; groups.set(k, g); } g.cells.push(rg);
        }
        for (const { d, cells } of groups.values()) {
          // merge into bounding rectangles per column run
          const bb = cells.reduce((a, b) => ({ r0: Math.min(a.r0, b.r0), c0: Math.min(a.c0, b.c0), r1: Math.max(a.r1, b.r1), c1: Math.max(a.c1, b.c1) }));
          const opMap = { between: 'between', notBetween: 'notBetween', equal: 'eq', notEqual: 'ne', greaterThan: 'gt', lessThan: 'lt', greaterThanOrEqual: 'ge', lessThanOrEqual: 'le' };
          const f0 = d.formulae && d.formulae[0] !== undefined ? String(d.formulae[0]) : '';
          const dv = { range: bb, type: d.type === 'textLength' ? 'textLength' : d.type, op: opMap[d.operator] || 'between', ignoreBlank: d.allowBlank !== false, errorMsg: d.error, errorTitle: d.errorTitle, errorStyle: d.errorStyle || 'stop', prompt: d.prompt };
          if (d.type === 'list') dv.source = f0[0] === '"' ? f0.slice(1, -1) : '=' + f0.replace(/^=/, '');
          else { dv.v1 = d.formulae && d.formulae[0] instanceof Date ? SU.jsDateToSerial(d.formulae[0]) : d.formulae ? d.formulae[0] : ''; dv.v2 = d.formulae && d.formulae[1] instanceof Date ? SU.jsDateToSerial(d.formulae[1]) : d.formulae ? d.formulae[1] : ''; }
          if (['list', 'whole', 'decimal', 'date', 'textLength'].includes(dv.type)) validations.push(dv);
        }
      } catch (e) { console.warn('dv', e); }
      // conditional formatting
      let cf = [];
      if (sm.cf) cf = sm.cf;
      else try { cf = importCF(ws); } catch (e) { console.warn('cf', e); }
      // images
      const images = [];
      try {
        const imgs = ws.getImages(); const skip = sm.chartPics || 0; const keep = imgs.slice(0, imgs.length - skip);
        for (const im of keep) {
          const media = X.getImage(+im.imageId); if (!media) continue;
          const bytes = media.buffer ? new Uint8Array(media.buffer) : media.base64 ? BTS.dataURLToBytes('data:;base64,' + media.base64) : null; if (!bytes) continue;
          const mime = 'image/' + (media.extension === 'jpg' ? 'jpeg' : media.extension || 'png');
          const src = await BTS.bytesToDataURL(bytes, mime);
          const tl = im.range.tl, br = im.range.br, ext = im.range.ext;
          const colPx = (c) => { let x = 0; for (let i = 0; i < Math.floor(c); i++) x += colW.has(i) ? colW.get(i) : SU.DEF_W; return x + (c % 1) * (colW.has(Math.floor(c)) ? colW.get(Math.floor(c)) : SU.DEF_W); };
          const rowPx = (r) => { let y = 0; for (let i = 0; i < Math.floor(r); i++) y += rowH.has(i) ? rowH.get(i) : SU.DEF_H; return y + (r % 1) * (rowH.has(Math.floor(r)) ? rowH.get(Math.floor(r)) : SU.DEF_H); };
          const tc = tl.nativeCol !== undefined ? tl.nativeCol + (tl.nativeColOff || 0) / 9525 / (colW.get(tl.nativeCol) || SU.DEF_W) : tl.col;
          const tr = tl.nativeRow !== undefined ? tl.nativeRow + (tl.nativeRowOff || 0) / 9525 / (rowH.get(tl.nativeRow) || SU.DEF_H) : tl.row;
          const x = colPx(tc), y = rowPx(tr); let w, h;
          if (ext && ext.width) { w = ext.width; h = ext.height; }
          else if (br) { const bc = br.nativeCol !== undefined ? br.nativeCol + (br.nativeColOff || 0) / 9525 / (colW.get(br.nativeCol) || SU.DEF_W) : br.col; const brr = br.nativeRow !== undefined ? br.nativeRow + (br.nativeRowOff || 0) / 9525 / (rowH.get(br.nativeRow) || SU.DEF_H) : br.row; w = colPx(bc) - x; h = rowPx(brr) - y; }
          else { w = 200; h = 150; }
          images.push({ id: 'im' + Math.random().toString(36).slice(2, 9), src, x, y, w: Math.max(10, w), h: Math.max(10, h), z: 1 });
        }
      } catch (e) { console.warn('images', e); }
      // page setup
      const ps = ws.pageSetup || {};
      const paperMap = { 9: 'A4', 1: 'Letter', 5: 'Legal', 8: 'A3', 11: 'A5' };
      const page = { orientation: ps.orientation === 'landscape' ? 'landscape' : 'portrait', paper: paperMap[ps.paperSize] || 'A4', margins: 'normal', printArea: ps.printArea ? SU.parseRange(String(ps.printArea).split(',')[0].replace(/^.*!/, '')) : null, gridlines: !!ps.showGridLines, headings: !!ps.showRowColHeaders, fit: ps.fitToPage ? (ps.fitToHeight === 0 ? 'width' : 'page') : 'none' };
      if (sm.page) Object.assign(page, sm.page);
      let filter = null; if (ws.autoFilter) { const g = SU.parseRange(typeof ws.autoFilter === 'string' ? ws.autoFilter : (ws.autoFilter.from && ws.autoFilter.to ? SU.addr(ws.autoFilter.from.row - 1, ws.autoFilter.from.column - 1) + ':' + SU.addr(ws.autoFilter.to.row - 1, ws.autoFilter.to.column - 1) : '')); if (g) filter = { range: g, cols: {} }; }
      const tables = sm.tables || [];
      if (sm.tables) { // undo baked table styling
        for (const t of tables) { const g = t.range; for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) { const st = styles.get(key(r, c)); if (!st) continue; const ns = { ...st }; delete ns.fill; if (r === g.r0) { delete ns.color; delete ns.b; } if (Object.keys(ns).length) styles.set(key(r, c), ns); else styles.delete(key(r, c)); } }
        if (sm.filter !== undefined) filter = sm.filter;
      }
      if (sm.filterCols && filter) filter.cols = sm.filterCols;
      const metaObj = { styles, links, comments, colW, hiddenC, rowH, hiddenR, merges, freeze, tabColor, validations, cf, images, page, filter, tables, charts: (sm.charts || []).map(c => ({ ...c })), showGrid: v.showGridLines !== false };
      if (sm.showGrid !== undefined) metaObj.showGrid = sm.showGrid;
      if (sm.protected) metaObj.protected = true;
      if (sm.hiddenR) metaObj.hiddenR = new Set(sm.hiddenR);
      sheets.push({ name, data: dataArr, meta: metaObj });
    }
    if (!sheets.length) sheets.push({ name: 'Sheet1', data: [], meta: {} });
    const names = [];
    try { for (const n of (X.definedNames.model || [])) { if (!n.name || n.name.startsWith('_xlnm') || n.localSheetId !== undefined) continue; if (n.ranges && n.ranges.length === 1) names.push({ name: n.name, expression: '=' + n.ranges[0] }); } } catch (e) { console.warn(e); }
    IO.buildWorkbook(wb, sheets, names);
    // re-apply filter visibility
    wb.sheets.forEach(s => { if (s.filter && Object.keys(s.filter.cols || {}).length) { wb.begin('x'); A.applyFilter(s); wb.tx = null; wb.txDepth = 0; } });
    const act = X.views && X.views[0] && X.views[0].activeTab; if (act !== undefined && act < wb.sheets.length) wb.activeIndex = act;
  };
  function importCF(ws) {
    const out = []; const list = ws.conditionalFormattings || [];
    for (const cfm of list) {
      const rg = SU.parseRange(String(cfm.ref).split(' ')[0]); if (!rg) continue;
      for (const rule of cfm.rules || []) {
        const st = rule.style || {}; const fmt = {};
        const fill = st.fill && (fromColor(st.fill.bgColor, null) || fromColor(st.fill.fgColor, null)); if (fill) fmt.fill = fill;
        const col = st.font && fromColor(st.font.color, null); if (col) fmt.color = col; if (st.font && st.font.bold) fmt.b = true;
        const base = { id: 'cf' + Math.random().toString(36).slice(2, 8), range: rg, fmt };
        const f = rule.formulae || [];
        if (rule.type === 'cellIs') { const op = { greaterThan: 'gt', lessThan: 'lt', between: 'between', equal: 'eq' }[rule.operator]; if (!op) continue; out.push({ ...base, type: op, v1: String(f[0]).replace(/^"|"$/g, ''), v2: f[1] }); }
        else if (rule.type === 'containsText') out.push({ ...base, type: 'text', v1: rule.text });
        else if (rule.type === 'duplicateValues') out.push({ ...base, type: 'dup' });
        else if (rule.type === 'uniqueValues') out.push({ ...base, type: 'unique' });
        else if (rule.type === 'top10') out.push({ ...base, type: rule.bottom ? 'bottom' : 'top', n: rule.rank || 10, pct: !!rule.percent });
        else if (rule.type === 'aboveAverage') out.push({ ...base, type: rule.aboveAverage === false ? 'belowavg' : 'aboveavg' });
        else if (rule.type === 'dataBar') out.push({ ...base, type: 'bar', color: fromColor(rule.color, '#5B9BD5') });
        else if (rule.type === 'colorScale') out.push({ ...base, type: 'scale', colors: (rule.color || []).map(c => fromColor(c, '#FFFFFF')) });
        else if (rule.type === 'expression' && f[0]) out.push({ ...base, type: 'formula', formula: '=' + f[0] });
      }
    }
    return out;
  }
  IO.loadSheetJS = (wb, data) => {
    const X = XLSX.read(data, { type: 'array', cellFormula: true, cellNF: true, cellDates: false, cellStyles: true });
    const sheets = [];
    for (const name of X.SheetNames) {
      const ws = X.Sheets[name]; const dataArr = [], styles = new Map(), links = new Map(), comments = new Map();
      for (const a in ws) {
        if (a[0] === '!') continue; const p = SU.parseAddr(a); if (!p) continue; const c = ws[a];
        let v = null;
        if (c.f) v = '=' + c.f; else if (c.t === 'n') v = c.v; else if (c.t === 'b') v = !!c.v; else if (c.t === 'e') v = c.w || '#N/A'; else if (c.t === 's' || c.t === 'str') v = String(c.v); else if (c.t === 'd') v = SU.jsDateToSerial(new Date(c.v));
        if (v !== null && v !== '') (dataArr[p.r] = dataArr[p.r] || [])[p.c] = v;
        const st = {}; if (c.z && c.z !== 'General') st.nf = c.z;
        if (c.s && c.s.fgColor && c.s.fgColor.rgb && c.s.patternType === 'solid') st.fill = '#' + c.s.fgColor.rgb.slice(-6);
        if (Object.keys(st).length) styles.set(key(p.r, p.c), st);
        if (c.l && c.l.Target) links.set(key(p.r, p.c), c.l.Target);
        if (c.c && c.c.length) comments.set(key(p.r, p.c), c.c.map(x => x.t).join('\n'));
      }
      const merges = (ws['!merges'] || []).map(m => ({ r0: m.s.r, c0: m.s.c, r1: m.e.r, c1: m.e.c }));
      const colW = new Map(), hiddenC = new Set(); (ws['!cols'] || []).forEach((col, i) => { if (!col) return; const px = col.wpx || (col.wch ? widthToPx(col.wch) : col.width ? widthToPx(col.width) : null); if (px && Math.abs(px - SU.DEF_W) > 1) colW.set(i, Math.round(px)); if (col.hidden) hiddenC.add(i); });
      const rowH = new Map(), hiddenR = new Set(); (ws['!rows'] || []).forEach((row, i) => { if (!row) return; const px = row.hpx || (row.hpt ? ptToPx(row.hpt) : null); if (px && Math.abs(px - SU.DEF_H) > 1) rowH.set(i, Math.round(px)); if (row.hidden) hiddenR.add(i); });
      sheets.push({ name: name.slice(0, 31), data: dataArr, meta: { styles, links, comments, merges, colW, hiddenC, rowH, hiddenR } });
    }
    const names = []; try { (X.Workbook && X.Workbook.Names || []).forEach(n => { if (n.Sheet === undefined && !n.Name.startsWith('_xlnm')) names.push({ name: n.Name, expression: '=' + n.Ref }); }); } catch { }
    if (!sheets.length) sheets.push({ name: 'Sheet1', data: [], meta: {} });
    IO.buildWorkbook(wb, sheets, names);
    wb.sheets.forEach(s => { if (!s.colW.size) autofitImported(s); });
  };

  /* ================= EXPORT ================= */
  function excelStyle(st, tbl) {
    const o = {};
    const font = {}; if (st.b || (tbl && tbl.b)) font.bold = true; if (st.i) font.italic = true; if (st.u) font.underline = st.u === 'double' ? 'double' : true; if (st.s) font.strike = true;
    font.name = st.font || 'Calibri'; font.size = st.sz || 11; const fc = st.color || (tbl && tbl.color); if (fc) font.color = { argb: toARGB(fc) };
    o.font = font;
    const fill = st.fill || (tbl && tbl.fill); if (fill) o.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: toARGB(fill) }, bgColor: { indexed: 64 } };
    const b = (x) => x ? { style: x.style || 'thin', color: { argb: toARGB(x.color || '#000000') } } : undefined;
    if (st.bT || st.bB || st.bL || st.bR) { o.border = {}; if (st.bT) o.border.top = b(st.bT); if (st.bB) o.border.bottom = b(st.bB); if (st.bL) o.border.left = b(st.bL); if (st.bR) o.border.right = b(st.bR); }
    if (st.h || st.v || st.wrap || st.indent) { o.alignment = {}; if (st.h) o.alignment.horizontal = st.h; if (st.v) o.alignment.vertical = st.v === 'middle' ? 'middle' : st.v; if (st.wrap) o.alignment.wrapText = true; if (st.indent) o.alignment.indent = st.indent; }
    if (st.nf) o.numFmt = st.nf;
    if (st.locked === false) o.protection = { locked: false };
    return o;
  }
  function tableCellStyle(s, r, c) {
    for (const t of s.tables) { const g = t.range; if (!SU.inRange(g, r, c)) continue; if (r === g.r0 && t.header !== false) return { fill: t.theme.head, color: t.theme.headText || '#FFFFFF', b: true }; const i = r - g.r0 - 1; if (t.total && r === g.r1) return { b: true }; return i % 2 === 0 ? { fill: t.theme.band } : null; }
    return null;
  }
  const paperCode = { A4: 9, Letter: 1, Legal: 5, A3: 8, A5: 11 };
  IO.toXLSX = async (wb) => {
    const X = new ExcelJS.Workbook(); X.creator = 'BTS Sheets'; X.created = new Date(); X.calcProperties = { fullCalcOnLoad: true };
    const meta = { v: 1, sheets: {} };
    X.views = [{ activeTab: wb.activeIndex, x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, visibility: 'visible' }];
    for (const s of wb.sheets) {
      const views = [{ state: s.freeze.r || s.freeze.c ? 'frozen' : 'normal', xSplit: s.freeze.c || undefined, ySplit: s.freeze.r || undefined, showGridLines: s.showGrid, zoomScale: 100 }];
      if (!(s.freeze.r || s.freeze.c)) { delete views[0].xSplit; delete views[0].ySplit; }
      const ws = X.addWorksheet(s.name, { properties: s.tabColor ? { tabColor: { argb: toARGB(s.tabColor) } } : {}, views });
      const u = s.usedRangeAll();
      // columns
      const maxC = Math.max(u.c1, ...[...s.colW.keys(), ...s.hiddenC].filter(x => x < 1000));
      for (let c = 0; c <= maxC; c++) { const col = ws.getColumn(c + 1); if (s.colW.has(c)) col.width = pxToWidth(s.colW.get(c)); else col.width = pxToWidth(SU.DEF_W); if (s.hiddenC.has(c)) col.hidden = true; }
      // cells
      const used = s.usedRange();
      const tblStyles = s.tables.length;
      for (let r = 0; r <= used.r1; r++) for (let c = 0; c <= used.c1; c++) {
        if (s.isEmpty(r, c)) continue;
        const raw = s.raw(r, c); const cell = ws.getCell(r + 1, c + 1);
        if (typeof raw === 'string' && raw[0] === '=') {
          let res = s.value(r, c); if (SU.isErr(res)) res = { error: res.value === '#CYCLE!' ? '#REF!' : res.value }; else if (res && typeof res === 'object') res = undefined;
          cell.value = { formula: raw.slice(1).replace(/^=/, ''), result: res === null ? undefined : res };
        } else if (typeof raw === 'string' && /^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!)$/.test(raw)) cell.value = { error: raw };
        else cell.value = raw;
        const link = s.links.get(key(r, c)); if (link && !(typeof raw === 'string' && raw[0] === '=')) cell.value = { text: String(SU.valueText(s.value(r, c))), hyperlink: link };
      }
      for (const [k, st] of s.styles) { const { r, c } = unkey(k); if (r > 1048575) continue; ws.getCell(r + 1, c + 1).style = excelStyle(st, tblStyles ? tableCellStyle(s, r, c) : null); }
      if (tblStyles) for (const t of s.tables) { const g = t.range; for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) if (!s.styles.has(key(r, c))) { const ts = tableCellStyle(s, r, c); if (ts) ws.getCell(r + 1, c + 1).style = excelStyle({}, ts); } }
      for (const [k, txt] of s.comments) { const { r, c } = unkey(k); ws.getCell(r + 1, c + 1).note = String(txt); }
      // rows
      const maxR = Math.max(used.r1, ...[...s.rowH.keys(), ...s.hiddenR].filter(x => x < 100000));
      for (let r = 0; r <= maxR; r++) { if (s.rowH.has(r) || s.hiddenR.has(r) || s.filterHidden.has(r)) { const row = ws.getRow(r + 1); if (s.rowH.has(r)) row.height = pxToPt(s.rowH.get(r)); if (s.hiddenR.has(r) || s.filterHidden.has(r)) row.hidden = true; } }
      for (const m of s.merges) { try { ws.mergeCells(m.r0 + 1, m.c0 + 1, m.r1 + 1, m.c1 + 1); } catch (e) { console.warn(e); } }
      // data validation
      for (const dv of s.validations) {
        const opMap = { between: 'between', notBetween: 'notBetween', eq: 'equal', ne: 'notEqual', gt: 'greaterThan', lt: 'lessThan', ge: 'greaterThanOrEqual', le: 'lessThanOrEqual' };
        const g = { ...dv.range, r1: Math.min(dv.range.r1, 1048575), c1: Math.min(dv.range.c1, 16383) };
        const rule = { type: dv.type, allowBlank: dv.ignoreBlank !== false, showErrorMessage: true, errorStyle: dv.errorStyle || 'stop', errorTitle: dv.errorTitle || '', error: dv.errorMsg || '' };
        if (dv.prompt) { rule.showInputMessage = true; rule.prompt = dv.prompt; }
        if (dv.type === 'list') { const src = String(dv.source || ''); rule.formulae = [src[0] === '=' ? src.slice(1) : '"' + src + '"']; }
        else { rule.operator = opMap[dv.op || 'between']; rule.formulae = [dv.v1 === '' || dv.v1 === undefined ? 0 : +dv.v1]; if (dv.op === 'between' || dv.op === 'notBetween' || !dv.op) rule.formulae.push(dv.v2 === '' || dv.v2 === undefined ? 0 : +dv.v2); }
        try { ws.dataValidations.add(SU.rangeStr(g), rule); } catch (e) { console.warn(e); }
      }
      // conditional formatting
      let prio = 1;
      for (const rule of s.cf) {
        const ref = SU.rangeStr({ ...rule.range, r1: Math.min(rule.range.r1, 1048575), c1: Math.min(rule.range.c1, 16383) });
        const style = {}; const f = rule.fmt || {};
        if (f.fill) style.fill = { type: 'pattern', pattern: 'solid', bgColor: { argb: toARGB(f.fill) } }; if (f.color || f.b) style.font = { color: f.color ? { argb: toARGB(f.color) } : undefined, bold: f.b || undefined };
        let r = null; const num = (x) => isNaN(+x) ? '"' + String(x).replace(/"/g, '""') + '"' : +x;
        switch (rule.type) {
          case 'gt': r = { type: 'cellIs', operator: 'greaterThan', formulae: [num(rule.v1)], style }; break;
          case 'lt': r = { type: 'cellIs', operator: 'lessThan', formulae: [num(rule.v1)], style }; break;
          case 'between': r = { type: 'cellIs', operator: 'between', formulae: [num(rule.v1), num(rule.v2)], style }; break;
          case 'eq': r = { type: 'cellIs', operator: 'equal', formulae: [num(rule.v1)], style }; break;
          case 'text': r = { type: 'containsText', operator: 'containsText', text: String(rule.v1), style }; break;
          case 'dup': r = { type: 'expression', formulae: [`COUNTIF(${SU.absRangeStr(rule.range)},${SU.addr(rule.range.r0, rule.range.c0)})>1`], style }; break;
          case 'unique': r = { type: 'expression', formulae: [`COUNTIF(${SU.absRangeStr(rule.range)},${SU.addr(rule.range.r0, rule.range.c0)})=1`], style }; break;
          case 'top': case 'bottom': r = { type: 'top10', rank: +rule.n || 10, percent: !!rule.pct, bottom: rule.type === 'bottom', style }; break;
          case 'aboveavg': case 'belowavg': r = { type: 'aboveAverage', aboveAverage: rule.type === 'aboveavg', style }; break;
          case 'formula': r = { type: 'expression', formulae: [String(rule.formula).replace(/^=/, '')], style }; break;
          case 'bar': r = { type: 'dataBar', minLength: 0, maxLength: 100, gradient: true, cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: toARGB(rule.color || '#5B9BD5') } }; break;
          case 'scale': { const cols = rule.colors || ['#F8696B', '#FFEB84', '#63BE7B']; r = { type: 'colorScale', cfvo: cols.length === 3 ? [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }] : [{ type: 'min' }, { type: 'max' }], color: cols.map(c => ({ argb: toARGB(c) })) }; break; }
        }
        if (r) { r.priority = prio++; try { ws.addConditionalFormatting({ ref, rules: [r] }); } catch (e) { console.warn(e); } }
      }
      // filter, page setup
      if (s.filter) ws.autoFilter = SU.rangeStr(s.filter.range);
      const pg = s.page; const marg = { normal: { left: .7, right: .7, top: .75, bottom: .75, header: .3, footer: .3 }, narrow: { left: .25, right: .25, top: .75, bottom: .75, header: .3, footer: .3 }, wide: { left: 1, right: 1, top: 1, bottom: 1, header: .5, footer: .5 } }[pg.margins] || undefined;
      ws.pageSetup = { orientation: pg.orientation, paperSize: paperCode[pg.paper] || 9, margins: marg, showGridLines: !!pg.gridlines, showRowColHeaders: !!pg.headings, fitToPage: pg.fit !== 'none', fitToWidth: pg.fit !== 'none' ? 1 : undefined, fitToHeight: pg.fit === 'page' ? 1 : pg.fit === 'width' ? 0 : undefined };
      if (pg.printArea) ws.pageSetup.printArea = SU.rangeStr(pg.printArea);
      // pictures: images, then chart snapshots
      const colPos = (x) => { let c = 0, acc = 0; while (c < 16000) { const w = s.colWidth(c); if (acc + w > x) return c + (w ? (x - acc) / w : 0); acc += w; c++; } return c; };
      const rowPos = (y) => { let r = 0, acc = 0; while (r < 1048576) { const h = s.rowHeight(r); if (acc + h > y) return r + (h ? (y - acc) / h : 0); acc += h; r++; } return r; };
      const pic = (src, o) => { const m = /^data:image\/(png|jpe?g|gif);base64,(.*)$/.exec(src); if (!m) return; const id = X.addImage({ base64: m[2], extension: m[1] === 'jpg' ? 'jpeg' : m[1] }); ws.addImage(id, { tl: { col: colPos(o.x), row: rowPos(o.y) }, ext: { width: o.w, height: o.h }, editAs: 'oneCell' }); };
      for (const im of s.images) { try { pic(im.src, im); } catch (e) { console.warn(e); } }
      let chartPics = 0;
      for (const ch of s.charts) { try { pic(Charts.toImage(ch, wb.sheetByName(ch.sheet) || s, 2), ch); chartPics++; } catch (e) { console.warn(e); } }
      const sm = { charts: s.charts.map(c => { const o = { ...c }; delete o._sheet; return o; }), tables: s.tables, cf: s.cf, chartPics, page: s.page, showGrid: s.showGrid, protected: s.protected || undefined };
      if (s.filter) { sm.filter = s.filter; sm.filterCols = s.filter.cols; }
      if (s.tables.length || s.charts.length || s.cf.length || s.filter || s.protected) meta.sheets[s.name] = sm;
      if (s.protected) { try { await ws.protect('', { selectLockedCells: true, selectUnlockedCells: true }); } catch { } }
    }
    for (const n of wb.listNames()) { try { X.definedNames.add(SU.fromHF(n.expr).replace(/^=/, ''), n.name); } catch (e) { console.warn('name', n, e); } }
    if (Object.keys(meta.sheets).length) {
      const mw = X.addWorksheet('_btsmeta', { state: 'veryHidden' }); const json = JSON.stringify(meta);
      for (let i = 0, r = 1; i < json.length; i += 30000, r++) mw.getCell(r, 1).value = json.slice(i, i + 30000);
    }
    const buf = await X.xlsx.writeBuffer();
    return new Uint8Array(buf);
  };
  IO.toCSV = (s, delim = ',') => {
    const u = s.usedRange(); const lines = [];
    for (let r = 0; r <= u.r1; r++) { const row = []; for (let c = 0; c <= u.c1; c++) { let t = s.displayText(r, c); if (t.includes(delim) || t.includes('"') || t.includes('\n')) t = '"' + t.replace(/"/g, '""') + '"'; row.push(t); } lines.push(row.join(delim)); }
    return '﻿' + lines.join('\r\n') + '\r\n';
  };
  IO.toSheetJSBook = (wb) => {
    const book = XLSX.utils.book_new();
    for (const s of wb.sheets) {
      const u = s.usedRange(); const ws = {};
      for (let r = 0; r <= u.r1; r++) for (let c = 0; c <= u.c1; c++) {
        if (s.isEmpty(r, c)) continue; const raw = s.raw(r, c), v = s.value(r, c); const st = s.style(r, c); const a = XLSX.utils.encode_cell({ r, c });
        const cell = {};
        if (SU.isErr(v)) { cell.t = 'e'; cell.v = 0x07; cell.w = v.value; } else if (typeof v === 'number') { cell.t = 'n'; cell.v = v; } else if (typeof v === 'boolean') { cell.t = 'b'; cell.v = v; } else { cell.t = 's'; cell.v = v === null ? '' : String(v); }
        if (typeof raw === 'string' && raw[0] === '=') cell.f = raw.slice(1);
        if (st.nf) cell.z = st.nf;
        ws[a] = cell;
      }
      ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(0, u.r1), c: Math.max(0, u.c1) } });
      if (s.merges.length) ws['!merges'] = s.merges.map(m => ({ s: { r: m.r0, c: m.c0 }, e: { r: m.r1, c: m.c1 } }));
      ws['!cols'] = Array.from({ length: Math.max(0, u.c1) + 1 }, (_, c) => ({ wpx: s.rawColWidth(c) }));
      XLSX.utils.book_append_sheet(book, ws, s.name);
    }
    return book;
  };
  IO.toODS = (wb) => new Uint8Array(XLSX.write(IO.toSheetJSBook(wb), { bookType: 'ods', type: 'array' }));
  IO.toHTML = (wb) => {
    const parts = wb.sheets.map(s => `<h2>${SU.esc(s.name)}</h2>` + IO.sheetTableHTML(s, s.usedRangeAll(), { gridlines: true }).html);
    return `<!doctype html><html><head><meta charset="utf-8"><title>${SU.esc(BTS.title)}</title><style>body{font-family:Calibri,Carlito,Arial,sans-serif;margin:24px}h2{font:600 16px sans-serif;margin:24px 0 8px}table{border-collapse:collapse;table-layout:fixed}td{padding:0 3px;overflow:hidden;white-space:nowrap;font-size:11pt}td.wrap{white-space:pre-wrap}</style></head><body>${parts.join('\n')}</body></html>`;
  };

  /* ---------- HTML table of a range (print/html export) ---------- */
  IO.sheetTableHTML = (s, g, opt = {}) => {
    const cols = []; for (let c = g.c0; c <= g.c1; c++) if (s.colWidth(c)) cols.push(c);
    const rows = []; for (let r = g.r0; r <= g.r1; r++) if (s.rowHeight(r)) rows.push(r);
    const skip = new Set(); const mergeAt = new Map();
    for (const m of s.merges) { if (!SU.intersects(m, g)) continue; mergeAt.set(key(m.r0, m.c0), m); for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) if (r !== m.r0 || c !== m.c0) skip.add(key(r, c)); }
    const gl = opt.gridlines ? '1px solid #c8c8c8' : '';
    let totalW = cols.reduce((a, c) => a + s.colWidth(c), 0) + (opt.headings ? 36 : 0);
    let h = `<table style="width:${totalW}px">`;
    h += '<colgroup>' + (opt.headings ? '<col style="width:36px">' : '') + cols.map(c => `<col style="width:${s.colWidth(c)}px">`).join('') + '</colgroup>';
    if (opt.headings) h += '<thead><tr><th class="hd"></th>' + cols.map(c => `<th class="hd">${SU.colName(c)}</th>`).join('') + '</tr></thead>';
    h += '<tbody>';
    const ver = s.wb.version;
    for (const r of rows) {
      h += `<tr style="height:${s.rowHeight(r)}px">` + (opt.headings ? `<th class="hd">${r + 1}</th>` : '');
      for (const c of cols) {
        if (skip.has(key(r, c))) continue;
        const st = s.style(r, c); const m = mergeAt.get(key(r, c)); const v = s.value(r, c);
        const cf = s.cf.length ? SU.evalCF(s, r, c, ver) : null; const tb = s.tables.length ? tableCellStyle(s, r, c) : null;
        const f = Grid.showFormulas && s.isFormula(r, c) ? { text: s.raw(r, c) } : SU.format(v, st.nf);
        const css = [];
        const fill = (cf && cf.fill) || st.fill || (tb && tb.fill); if (fill) css.push('background:' + fill);
        const color = (cf && cf.color) || f.color || st.color || (tb && tb.color); if (color) css.push('color:' + color);
        if (st.b || (cf && cf.b) || (tb && tb.b)) css.push('font-weight:bold'); if (st.i) css.push('font-style:italic');
        const deco = [st.u ? 'underline' : '', st.s ? 'line-through' : ''].filter(Boolean).join(' '); if (deco) css.push('text-decoration:' + deco);
        if (st.font) css.push('font-family:' + SU.cssFont(st.font).replace(/"/g, "'")); if (st.sz) css.push('font-size:' + st.sz + 'pt');
        const hal = st.h || (typeof v === 'number' ? 'right' : (typeof v === 'boolean' || SU.isErr(v)) ? 'center' : 'left'); css.push('text-align:' + (hal === 'fill' ? 'left' : hal));
        css.push('vertical-align:' + (st.v === 'middle' ? 'middle' : st.v === 'top' ? 'top' : 'bottom'));
        if (st.indent) css.push('padding-left:' + (3 + st.indent * 9) + 'px');
        const bd = (b) => b ? `${b.style === 'medium' ? 2 : b.style === 'thick' ? 3 : b.style === 'double' ? 3 : 1}px ${b.style === 'double' ? 'double' : b.style === 'dashed' || b.style === 'mediumDashed' ? 'dashed' : b.style === 'dotted' || b.style === 'hair' ? 'dotted' : 'solid'} ${b.color || '#000'}` : gl;
        const sB = m ? s.style(m.r1, c) : st, sR = m ? s.style(r, m.c1) : st;
        const bT = bd(st.bT), bB = bd(sB.bB), bL = bd(st.bL), bR = bd(sR.bR);
        if (bT) css.push('border-top:' + bT); if (bB) css.push('border-bottom:' + bB); if (bL) css.push('border-left:' + bL); if (bR) css.push('border-right:' + bR);
        let txt = SU.esc(f.text);
        if (f.acct) txt = `<span style="float:left">${SU.esc(f.acct.sym)}</span>${SU.esc(f.acct.num.trim())}`;
        if (cf && cf.bar && typeof v === 'number') css.push(`background-image:linear-gradient(90deg, ${cf.bar.color} 0, ${SU.mixColor(cf.bar.color, '#ffffff', .7)} ${Math.round(cf.bar.pct * 100)}%, transparent ${Math.round(cf.bar.pct * 100)}%);background-size:100% 70%;background-repeat:no-repeat;background-position:0 50%`);
        // allow overflow of left-aligned text into empty neighbour cells
        const overflow = !st.wrap && !m && typeof v === 'string' && hal === 'left' && c + 1 <= g.c1 && s.isEmpty(r, c + 1);
        if (overflow) css.push('overflow:visible;position:relative;z-index:1');
        h += `<td${m ? ` rowspan="${countVis(s, m.r0, m.r1, 'r')}" colspan="${countVis(s, m.c0, m.c1, 'c')}"` : ''}${st.wrap ? ' class="wrap"' : ''} style="${css.join(';')}">${txt}</td>`;
      }
      h += '</tr>';
    }
    h += '</tbody></table>';
    return { html: h, width: totalW, height: rows.reduce((a, r) => a + s.rowHeight(r), 0) + (opt.headings ? 20 : 0), cols, rows };
  };
  function countVis(s, a, b, axis) { let n = 0; for (let i = a; i <= b; i++) if (axis === 'r' ? s.rowHeight(i) : s.colWidth(i)) n++; return Math.max(1, n); }

  /* ---------- print / PDF ---------- */
  const PAPER = { A4: [8.27, 11.69], Letter: [8.5, 11], Legal: [8.5, 14], A3: [11.69, 16.54], A5: [5.83, 8.27] };
  const MARGINS = { normal: [0.75, 0.7], narrow: [0.75, 0.25], wide: [1, 1] };
  IO.preparePrint = (wb, sheetsToPrint) => {
    let root = document.getElementById('print-root'); if (root) root.remove();
    root = document.createElement('div'); root.id = 'print-root'; document.body.appendChild(root);
    const s0 = sheetsToPrint[0]; const pg = s0.page;
    const paper = PAPER[pg.paper] || PAPER.A4; const land = pg.orientation === 'landscape';
    const [pw, ph] = land ? [paper[1], paper[0]] : paper; const [mv, mh] = MARGINS[pg.margins] || MARGINS.normal;
    const printW = (pw - 2 * mh) * 96, printH = (ph - 2 * mv) * 96;
    // margins are drawn inside the page (page margin areas get the window background colour in Electron's printToPDF)
    let css = `@page { size: ${pw}in ${ph}in; margin: 0; } html, body { background: #fff !important; }`;
    let html = '';
    for (const s of sheetsToPrint) {
      const g = s.page.printArea || s.usedRangeAll();
      if (g.r1 < 0 || g.c1 < 0) { if (!s.charts.length && !s.images.length) continue; }
      const gg = { r0: g.r0, c0: g.c0, r1: Math.max(g.r1, 0), c1: Math.max(g.c1, 0) };
      // extend to include charts/images when no explicit print area
      if (!s.page.printArea) for (const o of [...s.charts, ...s.images]) { gg.c1 = Math.max(gg.c1, s.colAt(o.x + o.w)); gg.r1 = Math.max(gg.r1, s.rowAt(o.y + o.h)); }
      const t = IO.sheetTableHTML(s, gg, { gridlines: s.page.gridlines, headings: s.page.headings });
      let zoom = 1;
      if (s.page.fit === 'width' || s.page.fit === 'page') zoom = Math.min(1, printW / t.width);
      if (s.page.fit === 'page') zoom = Math.min(zoom, printH / t.height);
      // objects
      const ox = s.colLeft(gg.c0) - (s.page.headings ? 36 : 0), oy = s.rowTop(gg.r0) - (s.page.headings ? 20 : 0);
      let objs = '';
      for (const ch of s.charts) { try { const src = Charts.toImage(ch, wb.sheetByName(ch.sheet) || s, 3); objs += `<div class="p-chart" style="left:${ch.x - ox}px;top:${ch.y - oy}px;width:${ch.w}px;height:${ch.h}px;border:1px solid #d6dae1;background:#fff;padding:8px;box-sizing:border-box"><img src="${src}"></div>`; } catch (e) { console.warn(e); } }
      for (const im of s.images) objs += `<div class="p-chart" style="left:${im.x - ox}px;top:${im.y - oy}px;width:${im.w}px;height:${im.h}px"><img src="${im.src}"></div>`;
      html += `<table class="pg pgwrap" style="width:100%;border-collapse:collapse"><thead><tr><td style="height:${mv}in;padding:0"></td></tr></thead><tfoot><tr><td style="height:${mv}in;padding:0"></td></tr></tfoot><tbody><tr><td style="padding:0 ${mh}in;vertical-align:top;white-space:normal"><div style="zoom:${zoom}"><div style="position:relative">${t.html}${objs}</div></div></td></tr></tbody></table>`;
    }
    if (!html) html = '<div class="pg"></div>';
    root.innerHTML = `<style>${css}</style>` + html;
    document.body.classList.add('printing');
    return { pageSize: { width: pw, height: ph }, landscape: false };
  };
  IO.endPrint = () => { document.body.classList.remove('printing'); const r = document.getElementById('print-root'); if (r) r.remove(); };
})();
