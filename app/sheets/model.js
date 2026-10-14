/* BTS Sheets — workbook model: HyperFormula engine, per-cell styles, sheet metadata, transactions (undo/redo), structure ops */
(function () {
  const { key, unkey } = SU;
  const HF = window.HyperFormula;
  const DEF_W = 64, DEF_H = 20;
  const HF_CONFIG = {
    licenseKey: 'gpl-v3', undoLimit: 0, evaluateNullToZero: true, useArrayArithmetic: true,
    maxRows: 1048576, maxColumns: 16384, dateFormats: ['DD/MM/YYYY', 'DD-MM-YYYY', 'YYYY-MM-DD', 'DD/MM/YY', 'MM/DD/YYYY'],
    timeFormats: ['hh:mm', 'hh:mm:ss.sss'], currencySymbol: ['₹', '$', '€', '£'], functionArgSeparator: ',', decimalSeparator: '.', thousandSeparator: '',
    precisionRounding: 10,
  };

  const META_PROPS = ['styles', 'colW', 'rowH', 'hiddenR', 'hiddenC', 'merges', 'comments', 'links', 'validations', 'cf', 'tables', 'filter', 'charts', 'images', 'page', 'freeze', 'rowCount', 'colCount', 'filterHidden'];

  class Sheet {
    constructor(wb, name) {
      this.wb = wb; this.name = name; this.id = null; this.tabColor = null;
      this.styles = new Map(); this.colW = new Map(); this.rowH = new Map();
      this.hiddenR = new Set(); this.hiddenC = new Set(); this.filterHidden = new Set();
      this.merges = []; this.comments = new Map(); this.links = new Map();
      this.validations = []; this.cf = []; this.tables = []; this.filter = null;
      this.charts = []; this.images = []; this.freeze = { r: 0, c: 0 };
      this.page = { orientation: 'portrait', paper: 'A4', margins: 'normal', printArea: null, gridlines: false, headings: false, fit: 'none' };
      this.showGrid = true; this.showHeadings = true; this.protected = false;
      this.rowCount = 1000; this.colCount = 100;
      this.sel = { r0: 0, c0: 0, r1: 0, c1: 0 }; this.active = { r: 0, c: 0 }; this.scroll = { x: 0, y: 0 };
      this._geomDirty = true;
    }
    get index() { return this.wb.sheets.indexOf(this); }
    a(r, c) { return { sheet: this.id, row: r, col: c }; }
    value(r, c) { try { return this.wb.hf.getCellValue(this.a(r, c)); } catch { return null; } }
    /** Raw content as stored: formula string (=...), number, boolean, string (text) or null */
    raw(r, c) {
      const hf = this.wb.hf; const ad = this.a(r, c);
      let v; try { v = hf.isEvaluationSuspended() ? hf._serialization.getCellSerialized(ad) : hf.getCellSerialized(ad); } catch { return null; }
      if (v === undefined || v === '') return null;
      if (typeof v === 'string') { if (v[0] === '=') return SU.fromHF(v); if (v[0] === "'") return v.slice(1); }
      if (v && typeof v === 'object' && 'value' in v) return v.value;
      return v;
    }
    /** text shown in the formula bar / editor */
    editText(r, c) {
      const v = this.raw(r, c); if (v === null) return '';
      if (typeof v === 'string') { if (v[0] === '=') return v; const p = SU.parseInput(v); return p.text || v[0] === "'" ? v : "'" + v; }
      if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
      if (typeof v === 'number') {
        const nf = (this.styles.get(key(r, c)) || {}).nf;
        if (nf && SU.isDateFormat(nf)) { const d = SU.serialToDate(v); const p = (n) => String(n).padStart(2, '0'); const t = v % 1 ? ' ' + p(d.getUTCHours()) + ':' + p(d.getUTCMinutes()) + (d.getUTCSeconds() ? ':' + p(d.getUTCSeconds()) : '') : ''; if (v >= 1) return p(d.getUTCDate()) + '/' + p(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear() + t; return t.trim(); }
        if (nf && /%/.test(nf)) return SU.generalNum(+(v * 100).toPrecision(15)) + '%';
        return SU.generalNum(v);
      }
      return String(v);
    }
    isFormula(r, c) { try { return this.wb.hf.doesCellHaveFormula(this.a(r, c)); } catch { return false; } }
    isEmpty(r, c) { try { return this.wb.hf.isCellEmpty(this.a(r, c)); } catch { return true; } }
    style(r, c) { return this.styles.get(key(r, c)) || EMPTY; }
    colWidth(c) { return this.hiddenC.has(c) ? 0 : (this.colW.has(c) ? this.colW.get(c) : DEF_W); }
    rowHeight(r) { return (this.hiddenR.has(r) || this.filterHidden.has(r)) ? 0 : (this.rowH.has(r) ? this.rowH.get(r) : DEF_H); }
    rawColWidth(c) { return this.colW.has(c) ? this.colW.get(c) : DEF_W; }
    rawRowHeight(r) { return this.rowH.has(r) ? this.rowH.get(r) : DEF_H; }
    isRowHidden(r) { return this.hiddenR.has(r) || this.filterHidden.has(r); }
    /* ---------- geometry (prefix sums) ---------- */
    geom() {
      if (!this._geomDirty && this._rp && this._rp.length === this.rowCount + 1 && this._cp.length === this.colCount + 1) return;
      const rp = new Float64Array(this.rowCount + 1); let y = 0;
      const anyR = this.rowH.size || this.hiddenR.size || this.filterHidden.size;
      for (let r = 0; r < this.rowCount; r++) { rp[r] = y; y += anyR ? this.rowHeight(r) : DEF_H; } rp[this.rowCount] = y;
      const cp = new Float64Array(this.colCount + 1); let x = 0;
      for (let c = 0; c < this.colCount; c++) { cp[c] = x; x += this.colWidth(c); } cp[this.colCount] = x;
      this._rp = rp; this._cp = cp; this._geomDirty = false;
    }
    invalidateGeom() { this._geomDirty = true; }
    rowTop(r) { this.geom(); return r >= this.rowCount ? this._rp[this.rowCount] + (r - this.rowCount) * DEF_H : this._rp[Math.max(0, r)]; }
    colLeft(c) { this.geom(); return c >= this.colCount ? this._cp[this.colCount] + (c - this.colCount) * DEF_W : this._cp[Math.max(0, c)]; }
    totalH() { this.geom(); return this._rp[this.rowCount]; }
    totalW() { this.geom(); return this._cp[this.colCount]; }
    rowAt(y) { this.geom(); const a = this._rp; let lo = 0, hi = this.rowCount - 1; if (y >= a[this.rowCount]) return this.rowCount - 1; while (lo < hi) { const m = (lo + hi + 1) >> 1; if (a[m] <= y) lo = m; else hi = m - 1; } return lo; }
    colAt(x) { this.geom(); const a = this._cp; let lo = 0, hi = this.colCount - 1; if (x >= a[this.colCount]) return this.colCount - 1; while (lo < hi) { const m = (lo + hi + 1) >> 1; if (a[m] <= x) lo = m; else hi = m - 1; } return lo; }
    ensureSize(r, c) { let ch = false; if (r >= this.rowCount - 1) { this.rowCount = Math.min(SU.MAXR, Math.max(this.rowCount + 1000, r + 200)); ch = true; } if (c >= this.colCount - 1) { this.colCount = Math.min(SU.MAXC, Math.max(this.colCount + 26, c + 10)); ch = true; } if (ch) this._geomDirty = true; return ch; }
    /* ---------- extents ---------- */
    dims() { try { return this.wb.hf.getSheetDimensions(this.id); } catch { return { width: 0, height: 0 }; } }
    /** last used row/col with content (not styles); -1 when empty */
    usedRange() {
      const d = this.dims(); let h = d.height, w = d.width;
      const rowEmpty = (r) => { for (let c = 0; c < w; c++) if (!this.isEmpty(r, c)) return false; return true; };
      while (h > 0 && rowEmpty(h - 1)) h--;
      const colEmpty = (c) => { for (let r = 0; r < h; r++) if (!this.isEmpty(r, c)) return false; return true; };
      while (w > 0 && colEmpty(w - 1)) w--;
      return { r1: h - 1, c1: w - 1 };
    }
    /** used range including styled/merged cells, for printing/export */
    usedRangeAll() {
      let { r1, c1 } = this.usedRange();
      for (const [k, s] of this.styles) { if (s.fill || s.bT || s.bB || s.bL || s.bR) { const { r, c } = unkey(k); if (r > r1) r1 = r; if (c > c1) c1 = c; } }
      for (const m of this.merges) { r1 = Math.max(r1, m.r1); c1 = Math.max(c1, m.c1); }
      return { r0: 0, c0: 0, r1, c1 };
    }
    mergeAt(r, c) { for (const m of this.merges) if (r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1) return m; return null; }
    /** expand range so that it fully contains any merge it touches */
    expandMerges(g) {
      g = { ...g }; let changed = true;
      while (changed) { changed = false; for (const m of this.merges) if (SU.intersects(g, m) && !(m.r0 >= g.r0 && m.r1 <= g.r1 && m.c0 >= g.c0 && m.c1 <= g.c1)) { g.r0 = Math.min(g.r0, m.r0); g.c0 = Math.min(g.c0, m.c0); g.r1 = Math.max(g.r1, m.r1); g.c1 = Math.max(g.c1, m.c1); changed = true; } }
      return g;
    }
    validationAt(r, c) { for (let i = this.validations.length - 1; i >= 0; i--) if (SU.inRange(this.validations[i].range, r, c)) return this.validations[i]; return null; }
    tableAt(r, c) { return this.tables.find(t => SU.inRange(t.range, r, c)) || null; }
    displayText(r, c) { const v = this.value(r, c); return SU.format(v, this.style(r, c).nf).text; }
  }
  const EMPTY = Object.freeze({});
  SU.EMPTY_STYLE = EMPTY;
  SU.DEF_W = DEF_W; SU.DEF_H = DEF_H;

  /* ---------- range shifting for structure changes ---------- */
  /** axis 'r'|'c'; count>0 insert at `at`; count<0 delete [at, at-count) */
  function shiftIdx(i, at, count) { if (count > 0) return i >= at ? i + count : i; const n = -count; if (i < at) return i; if (i >= at + n) return i - n; return null; }
  function shiftRange(g, axis, at, count) {
    if (!g) return g;
    const a0 = axis === 'r' ? 'r0' : 'c0', a1 = axis === 'r' ? 'r1' : 'c1';
    let s = g[a0], e = g[a1];
    if (count > 0) { if (s >= at) s += count; if (e >= at) e += count; }
    else { const n = -count; const de = at + n - 1; if (s >= at && e <= de) return null; if (s > de) s -= n; else if (s >= at) s = at; if (e > de) e -= n; else if (e >= at) e = at - 1; }
    const out = { ...g }; out[a0] = s; out[a1] = e; return out;
  }
  SU.shiftRange = shiftRange;
  function shiftKeyMap(map, axis, at, count) {
    const out = new Map();
    for (const [k, v] of map) { const { r, c } = unkey(k); if (axis === 'r') { const nr = shiftIdx(r, at, count); if (nr !== null) out.set(key(nr, c), v); } else { const nc = shiftIdx(c, at, count); if (nc !== null) out.set(key(r, nc), v); } }
    return out;
  }
  function shiftIdxMap(map, at, count) { const out = new Map(); for (const [i, v] of map) { const n = shiftIdx(i, at, count); if (n !== null) out.set(n, v); } return out; }
  function shiftIdxSet(set, at, count) { const out = new Set(); for (const i of set) { const n = shiftIdx(i, at, count); if (n !== null) out.add(n); } return out; }

  /* ---------- workbook ---------- */
  class Workbook {
    constructor() {
      this.hf = HF.buildEmpty(HF_CONFIG);
      this.sheets = []; this.activeIndex = 0;
      this.undoStack = []; this.redoStack = []; this.tx = null; this.txDepth = 0;
      this.listeners = new Set(); this.version = 0;
      this.props = { author: '', created: new Date() };
    }
    static get config() { return HF_CONFIG; }
    get active() { return this.sheets[this.activeIndex]; }
    on(fn) { this.listeners.add(fn); }
    emit(what) { this.version++; this.listeners.forEach(fn => fn(what)); }
    sheetByName(n) { const l = String(n).toLowerCase(); return this.sheets.find(s => s.name.toLowerCase() === l); }
    sheetById(id) { return this.sheets.find(s => s.id === id); }
    uniqueName(base) { let i = 1, n = base; while (this.sheetByName(n)) n = base + ' (' + (++i) + ')'; return n; }
    nextSheetName() { let i = this.sheets.length + 1; while (this.sheetByName('Sheet' + i)) i++; return 'Sheet' + i; }

    /** Create sheet without recording (used at load/new) */
    _addSheetRaw(name, index) {
      const s = new Sheet(this, name);
      const hfName = this.hf.addSheet(name); s.id = this.hf.getSheetId(hfName);
      if (index === undefined) this.sheets.push(s); else this.sheets.splice(index, 0, s);
      return s;
    }
    reset() {
      this.hf.destroy(); this.hf = HF.buildEmpty(HF_CONFIG);
      this.sheets = []; this.activeIndex = 0; this.undoStack = []; this.redoStack = [];
    }

    /* ---------- transactions ---------- */
    begin(label) {
      if (this.txDepth++ === 0) { const s = this.active; this.tx = { label, ops: [], sheet: s, selBefore: s ? { sel: { ...s.sel }, active: { ...s.active } } : null }; }
    }
    commit() {
      if (--this.txDepth > 0) return;
      const tx = this.tx; this.tx = null;
      if (tx && tx.ops.length) {
        const s = this.active; tx.selAfter = s ? { sel: { ...s.sel }, active: { ...s.active } } : null; tx.sheetAfter = s;
        this.undoStack.push(tx); if (this.undoStack.length > 300) this.undoStack.shift(); this.redoStack = [];
        BTS.markDirty();
      }
      this.emit('change');
    }
    abort() { if (--this.txDepth > 0) return; const tx = this.tx; this.tx = null; if (tx) this._revert(tx); this.emit('change'); }
    /** run fn inside a transaction */
    run(label, fn) { this.begin(label); try { const r = fn(); this.commit(); return r; } catch (e) { this.txDepth = 1; this.commit(); throw e; } }
    _rec(op) { if (this.tx) this.tx.ops.push(op); }

    /* primitive: content */
    setContents(sheet, cells) { // cells: [{r,c,v}] where v is raw (formula string, number, bool, string, null)
      if (!cells.length) return;
      const hf = this.hf;
      const recs = cells.map(({ r, c, v }) => ({ t: 'cell', s: sheet, r, c, b: sheet.raw(r, c), a: v === '' ? null : v }));
      this._applyCells(recs, 'a');
      recs.forEach(o => this._rec(o));
      let maxR = 0, maxC = 0; cells.forEach(({ r, c }) => { if (r > maxR) maxR = r; if (c > maxC) maxC = c; }); sheet.ensureSize(maxR, maxC);
    }
    _toHFContent(v) {
      if (v === null || v === undefined || v === '') return null;
      if (typeof v === 'string') { if (v[0] === '=' && v.length > 1) return SU.toHF(v); return "'" + v; }
      return v;
    }
    _applyCells(recs, which) {
      const hf = this.hf;
      const go = () => { for (const o of recs) { try { hf.setCellContents(o.s.a(o.r, o.c), [[this._toHFContent(o[which])]]); } catch (e) { try { hf.setCellContents(o.s.a(o.r, o.c), [["'" + String(o[which])]]); } catch { } } } };
      if (recs.length > 1) hf.batch(go); else go();
    }
    /* primitive: style */
    setStyle(sheet, r, c, st) {
      const k = key(r, c); const b = sheet.styles.get(k);
      const a = st && Object.keys(st).length ? st : undefined;
      if (a) sheet.styles.set(k, a); else sheet.styles.delete(k);
      this._rec({ t: 'style', s: sheet, k, b, a });
    }
    /* primitive: replace property value (objects treated immutably) */
    setProp(obj, prop, val) {
      const b = obj[prop]; obj[prop] = val; this._rec({ t: 'prop', o: obj, p: prop, b, a: val });
      if (obj instanceof Sheet && ['colW', 'rowH', 'hiddenR', 'hiddenC', 'filterHidden', 'rowCount', 'colCount'].includes(prop)) obj.invalidateGeom();
    }
    /* primitive: custom */
    custom(apply, revert) { apply(); this._rec({ t: 'fn', apply, revert }); }

    _revert(tx) {
      const ops = tx.ops; let cells = [];
      const flush = () => { if (cells.length) { this._applyCells(cells, 'b'); cells = []; } };
      for (let i = ops.length - 1; i >= 0; i--) {
        const o = ops[i];
        if (o.t === 'cell') { cells.push(o); continue; }
        flush();
        if (o.t === 'style') { if (o.b) o.s.styles.set(o.k, o.b); else o.s.styles.delete(o.k); }
        else if (o.t === 'prop') { o.o[o.p] = o.b; if (o.o.invalidateGeom) o.o.invalidateGeom(); }
        else if (o.t === 'fn') o.revert();
      }
      flush();
    }
    _apply(tx) {
      let cells = [];
      const flush = () => { if (cells.length) { this._applyCells(cells, 'a'); cells = []; } };
      for (const o of tx.ops) {
        if (o.t === 'cell') { cells.push(o); continue; }
        flush();
        if (o.t === 'style') { if (o.a) o.s.styles.set(o.k, o.a); else o.s.styles.delete(o.k); }
        else if (o.t === 'prop') { o.o[o.p] = o.a; if (o.o.invalidateGeom) o.o.invalidateGeom(); }
        else if (o.t === 'fn') o.apply();
      }
      flush();
    }
    canUndo() { return this.undoStack.length > 0; }
    canRedo() { return this.redoStack.length > 0; }
    undo() {
      const tx = this.undoStack.pop(); if (!tx) return null;
      this._revert(tx); this.redoStack.push(tx);
      this._restoreSel(tx.sheet, tx.selBefore); BTS.markDirty(); this.emit('undo'); return tx;
    }
    redo() {
      const tx = this.redoStack.pop(); if (!tx) return null;
      this._apply(tx); this.undoStack.push(tx);
      this._restoreSel(tx.sheetAfter || tx.sheet, tx.selAfter); BTS.markDirty(); this.emit('redo'); return tx;
    }
    _restoreSel(sheet, s) {
      const i = this.sheets.indexOf(sheet); if (i >= 0) this.activeIndex = i;
      if (sheet && s) { sheet.sel = { ...s.sel }; sheet.active = { ...s.active }; }
      if (this.activeIndex >= this.sheets.length) this.activeIndex = this.sheets.length - 1;
    }

    /* ---------- workbook-wide snapshots (for ops whose effect on other formulas can't be inverted exactly) ---------- */
    snapshotFormulas() {
      const snap = new Map();
      for (const s of this.sheets) {
        let f; try { f = this.hf.getSheetFormulas(s.id); } catch { continue; }
        const m = new Map(); f.forEach((row, r) => row && row.forEach((x, c) => { if (x !== undefined && x !== null) m.set(key(r, c), x); }));
        snap.set(s, m);
      }
      return { snap, names: this.hf.getAllNamedExpressionsSerialized() };
    }
    restoreFormulas(S) {
      const hf = this.hf;
      hf.batch(() => {
        for (const [s, m] of S.snap) {
          if (!this.sheets.includes(s)) continue;
          for (const [k, f] of m) { const { r, c } = unkey(k); let cur; try { cur = hf.getCellFormula(s.a(r, c)); } catch { } if (cur !== f) try { hf.setCellContents(s.a(r, c), [[f]]); } catch { } }
        }
        for (const n of S.names) { try { if (hf.getNamedExpression(n.name, n.scope)) hf.changeNamedExpression(n.name, n.expression, n.scope); else hf.addNamedExpression(n.name, n.expression, n.scope); } catch { } }
      });
    }
    captureRange(sheet, g) { const out = []; for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) { const v = sheet.raw(r, c); if (v !== null) out.push({ r, c, v }); } return out; }
    metaSnap(sheet) { const o = {}; META_PROPS.forEach(p => o[p] = sheet[p]); return o; }
    metaRestore(sheet, o) { META_PROPS.forEach(p => sheet[p] = o[p]); sheet.invalidateGeom(); }

    /* ---------- structure: insert/delete rows/cols ---------- */
    insertRows(sheet, at, n) { this._insert(sheet, 'r', at, n); }
    insertCols(sheet, at, n) { this._insert(sheet, 'c', at, n); }
    deleteRows(sheet, at, n) { this._delete(sheet, 'r', at, n); }
    deleteCols(sheet, at, n) { this._delete(sheet, 'c', at, n); }
    _shiftedMeta(sheet, axis, at, count) {
      const m = {};
      m.styles = shiftKeyMap(sheet.styles, axis, at, count);
      m.comments = shiftKeyMap(sheet.comments, axis, at, count);
      m.links = shiftKeyMap(sheet.links, axis, at, count);
      if (axis === 'r') { m.rowH = shiftIdxMap(sheet.rowH, at, count); m.hiddenR = shiftIdxSet(sheet.hiddenR, at, count); m.filterHidden = shiftIdxSet(sheet.filterHidden, at, count); m.rowCount = Math.max(100, sheet.rowCount + count); }
      else { m.colW = shiftIdxMap(sheet.colW, at, count); m.hiddenC = shiftIdxSet(sheet.hiddenC, at, count); m.colCount = Math.max(26, sheet.colCount + count); }
      // when inserting, copy formatting from the row/col before (Excel behaviour)
      if (count > 0 && at > 0) {
        for (const [k, st] of sheet.styles) { const { r, c } = unkey(k); if ((axis === 'r' ? r : c) === at - 1) for (let i = 0; i < count; i++) m.styles.set(axis === 'r' ? key(at + i, c) : key(r, at + i), st); }
        if (axis === 'r' && sheet.rowH.has(at - 1)) for (let i = 0; i < count; i++) m.rowH.set(at + i, sheet.rowH.get(at - 1));
        if (axis === 'c' && sheet.colW.has(at - 1)) for (let i = 0; i < count; i++) m.colW.set(at + i, sheet.colW.get(at - 1));
      }
      const sr = (g) => shiftRange(g, axis, at, count);
      m.merges = sheet.merges.map(sr).filter(g => g && (g.r1 > g.r0 || g.c1 > g.c0));
      m.validations = sheet.validations.map(v => ({ ...v, range: sr(v.range) })).filter(v => v.range);
      m.cf = sheet.cf.map(v => ({ ...v, range: sr(v.range) })).filter(v => v.range);
      m.tables = sheet.tables.map(v => ({ ...v, range: sr(v.range) })).filter(v => v.range);
      m.filter = sheet.filter ? (sr(sheet.filter.range) ? { ...sheet.filter, range: sr(sheet.filter.range), cols: shiftFilterCols(sheet.filter.cols, axis, at, count) } : null) : null;
      m.charts = sheet.charts.map(ch => ({ ...ch, range: ch.sheet && ch.sheet !== sheet.name ? ch.range : (sr(ch.range) || ch.range) }));
      m.page = { ...sheet.page, printArea: sheet.page.printArea ? sr(sheet.page.printArea) : null };
      const fz = { ...sheet.freeze }; if (axis === 'r' && at < fz.r) fz.r = Math.max(0, count > 0 ? fz.r + count : fz.r - Math.min(-count, fz.r - at)); if (axis === 'c' && at < fz.c) fz.c = Math.max(0, count > 0 ? fz.c + count : fz.c - Math.min(-count, fz.c - at)); m.freeze = fz;
      return m;
    }
    _otherSheetCharts(sheet, axis, at, count) { // charts on other sheets whose source is this sheet
      const list = [];
      for (const s of this.sheets) if (s !== sheet) s.charts.forEach((ch, i) => { if (ch.sheet === sheet.name) list.push([s, i, shiftRange(ch.range, axis, at, count) || ch.range]); });
      return list;
    }
    _insert(sheet, axis, at, n) {
      const before = this.metaSnap(sheet), after = { ...before, ...this._shiftedMeta(sheet, axis, at, n) };
      const hf = this.hf;
      this.custom(() => { axis === 'r' ? hf.addRows(sheet.id, [at, n]) : hf.addColumns(sheet.id, [at, n]); this.metaRestore(sheet, after); },
        () => { axis === 'r' ? hf.removeRows(sheet.id, [at, n]) : hf.removeColumns(sheet.id, [at, n]); this.metaRestore(sheet, before); });
    }
    _delete(sheet, axis, at, n) {
      const before = this.metaSnap(sheet), after = { ...before, ...this._shiftedMeta(sheet, axis, at, -n) };
      const hf = this.hf; let lost, fsnap;
      this.custom(() => {
        const d = sheet.dims();
        lost = this.captureRange(sheet, axis === 'r' ? { r0: at, r1: Math.min(at + n, d.height) - 1, c0: 0, c1: d.width - 1 } : { r0: 0, r1: d.height - 1, c0: at, c1: Math.min(at + n, d.width) - 1 });
        fsnap = this.snapshotFormulas();
        axis === 'r' ? hf.removeRows(sheet.id, [at, n]) : hf.removeColumns(sheet.id, [at, n]); this.metaRestore(sheet, after);
      }, () => {
        axis === 'r' ? hf.addRows(sheet.id, [at, n]) : hf.addColumns(sheet.id, [at, n]);
        this._applyCells(lost.map(x => ({ s: sheet, r: x.r, c: x.c, b: x.v })), 'b');
        this.restoreFormulas(fsnap); this.metaRestore(sheet, before);
      });
    }
    /** Insert cells shifting right/down within a range — implemented via move of the block */
    shiftCells(sheet, g, dir) { // dir: 'right'|'down'|'left'|'up'
      const d = sheet.dims(); const fsnap = this.snapshotFormulas(); const hf = this.hf;
      let src, dst, clear;
      const H = g.r1 - g.r0 + 1, W = g.c1 - g.c0 + 1;
      if (dir === 'down') { src = { r0: g.r0, c0: g.c0, r1: Math.max(g.r0, d.height - 1), c1: g.c1 }; dst = { r: g.r0 + H, c: g.c0 }; clear = g; }
      else if (dir === 'right') { src = { r0: g.r0, c0: g.c0, r1: g.r1, c1: Math.max(g.c0, d.width - 1) }; dst = { r: g.r0, c: g.c0 + W }; clear = g; }
      else if (dir === 'up') { src = { r0: g.r1 + 1, c0: g.c0, r1: Math.max(g.r1 + 1, d.height - 1), c1: g.c1 }; dst = { r: g.r0, c: g.c0 }; }
      else { src = { r0: g.r0, c0: g.c1 + 1, r1: g.r1, c1: Math.max(g.c1 + 1, d.width - 1) }; dst = { r: g.r0, c: g.c0 }; }
      const area = { r0: Math.min(src.r0, g.r0), c0: Math.min(src.c0, g.c0), r1: Math.max(src.r1, dst.r + (src.r1 - src.r0)), c1: Math.max(src.c1, dst.c + (src.c1 - src.c0)) };
      const beforeCells = this.captureRange(sheet, area);
      const stylesBefore = sheet.styles;
      // styles move with cells
      const st = new Map(sheet.styles);
      const moved = []; for (let r = src.r0; r <= src.r1; r++) for (let c = src.c0; c <= src.c1; c++) { const k = key(r, c); if (st.has(k)) { moved.push([r - src.r0 + dst.r, c - src.c0 + dst.c, st.get(k)]); st.delete(k); } }
      if (dir === 'up' || dir === 'left') for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) st.delete(key(r, c));
      moved.forEach(([r, c, s]) => st.set(key(r, c), s));
      let afterCells;
      this.custom(() => {
        if (dir === 'up' || dir === 'left') hf.batch(() => { for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) hf.setCellContents(sheet.a(r, c), [[null]]); });
        if (src.r1 >= src.r0 && src.c1 >= src.c0 && (src.r0 < d.height && src.c0 < d.width)) {
          try { hf.moveCells({ start: sheet.a(src.r0, src.c0), end: sheet.a(src.r1, src.c1) }, sheet.a(dst.r, dst.c)); } catch (e) { console.warn(e); }
        }
        sheet.styles = st; afterCells = afterCells || this.captureRange(sheet, area);
      }, () => {
        hf.batch(() => { for (let r = area.r0; r <= area.r1; r++) for (let c = area.c0; c <= area.c1; c++) hf.setCellContents(sheet.a(r, c), [[null]]); });
        this._applyCells(beforeCells.map(x => ({ s: sheet, r: x.r, c: x.c, b: x.v })), 'b');
        this.restoreFormulas(fsnap); sheet.styles = stylesBefore;
      });
    }

    /* ---------- sheets ---------- */
    addSheet(name, index) {
      name = name || this.nextSheetName(); index = index === undefined ? this.sheets.length : index;
      let s;
      this.custom(() => {
        if (!s) { s = this._addSheetRaw(name, index); return; }
        const n = this.hf.addSheet(s.name); s.id = this.hf.getSheetId(n); this.sheets.splice(index, 0, s);
        if (s._restore) { this._applyCells(s._restore.map(x => ({ s, r: x.r, c: x.c, b: x.v })), 'b'); }
      }, () => { this._removeSheetRaw(s); });
      return s;
    }
    _removeSheetRaw(s) {
      const i = this.sheets.indexOf(s); this.hf.removeSheet(s.id); this.sheets.splice(i, 1);
      if (this.activeIndex >= this.sheets.length) this.activeIndex = this.sheets.length - 1;
    }
    deleteSheet(s) {
      if (this.sheets.length <= 1) throw new Error('A workbook must contain at least one visible sheet.');
      const idx = this.sheets.indexOf(s); let fsnap, contents;
      this.custom(() => { const d = s.dims(); contents = this.captureRange(s, { r0: 0, c0: 0, r1: d.height - 1, c1: d.width - 1 }); fsnap = this.snapshotFormulas(); this._removeSheetRaw(s); },
        () => { const n = this.hf.addSheet(s.name); s.id = this.hf.getSheetId(n); this.sheets.splice(idx, 0, s); this._applyCells(contents.map(x => ({ s, r: x.r, c: x.c, b: x.v })), 'b'); this.restoreFormulas(fsnap); });
    }
    renameSheet(s, name) {
      name = String(name).trim();
      if (!name || /[\\\/\?\*\[\]:]/.test(name) || name.length > 31) throw new Error('Sheet names can’t be blank, longer than 31 characters, or contain \\ / ? * [ ] :');
      const ex = this.sheetByName(name); if (ex && ex !== s) throw new Error('That name is already taken. Try a different one.');
      const old = s.name; if (old === name) return;
      const fix = (from, to) => { this.sheets.forEach(x => x.charts.forEach(ch => { if (ch.sheet === from) ch.sheet = to; })); };
      this.custom(() => { this.hf.renameSheet(s.id, name); s.name = name; fix(old, name); }, () => { this.hf.renameSheet(s.id, old); s.name = old; fix(name, old); });
    }
    moveSheet(s, to) {
      const from = this.sheets.indexOf(s); to = SU.clamp(to, 0, this.sheets.length - 1); if (from === to) return;
      const mv = (a, b) => { this.sheets.splice(a, 1); this.sheets.splice(b, 0, s); this.activeIndex = b; };
      this.custom(() => mv(from, to), () => mv(to, from));
    }
    duplicateSheet(s) {
      const name = this.uniqueName(s.name.replace(/ \(\d+\)$/, '')); const idx = this.sheets.indexOf(s) + 1;
      const d = s.dims(); const contents = this.captureRange(s, { r0: 0, c0: 0, r1: d.height - 1, c1: d.width - 1 });
      const ns = this.addSheet(name, idx);
      const m = this.metaSnap(s);
      META_PROPS.forEach(p => { const v = m[p]; this.setProp(ns, p, v instanceof Map ? new Map(v) : v instanceof Set ? new Set(v) : Array.isArray(v) ? v.map(x => ({ ...x, id: x.id ? x.id + '_' + Date.now().toString(36) : undefined })) : (v && typeof v === 'object') ? { ...v } : v); });
      this.setProp(ns, 'tabColor', s.tabColor); ns.showGrid = s.showGrid;
      // keep formulas relative to the new sheet (same text)
      this.setContents(ns, contents.map(x => ({ r: x.r, c: x.c, v: x.v })));
      return ns;
    }

    /* ---------- named ranges ---------- */
    listNames() { return this.hf.getAllNamedExpressionsSerialized().filter(n => n.scope === undefined).map(n => ({ name: n.name, expr: n.expression, comment: n.options && n.options.comment })); }
    addName(name, expr) {
      if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(name) || /^[A-Za-z]{1,3}\d+$/.test(name) || /^(r|c|rc)\d*$/i.test(name)) throw new Error('The name “' + name + '” isn’t valid. Use letters, numbers, underscore and dots, start with a letter, and don’t use a cell address.');
      expr = expr[0] === '=' ? expr : '=' + expr; expr = SU.toHF(expr);
      if (!this.hf.isItPossibleToAddNamedExpression(name, expr)) throw new Error('That name already exists or the reference isn’t valid.');
      this.custom(() => this.hf.addNamedExpression(name, expr), () => this.hf.removeNamedExpression(name));
    }
    removeName(name) { const n = this.hf.getNamedExpression(name); if (!n) return; const ex = n.expression; this.custom(() => this.hf.removeNamedExpression(name), () => this.hf.addNamedExpression(name, ex)); }
    changeName(name, expr) { const n = this.hf.getNamedExpression(name); const old = n.expression; expr = SU.toHF(expr[0] === '=' ? expr : '=' + expr); this.custom(() => this.hf.changeNamedExpression(name, expr), () => this.hf.changeNamedExpression(name, old)); }
  }
  function shiftFilterCols(cols, axis, at, count) {
    if (!cols || axis === 'r') return cols; const out = {};
    for (const k of Object.keys(cols)) { const n = shiftIdx(+k, at, count); if (n !== null) out[n] = cols[k]; }
    return out;
  }

  /* ---------- style helpers ---------- */
  SU.mergeStyle = (base, patch) => { const o = { ...base }; for (const k in patch) { if (patch[k] === undefined || patch[k] === null || patch[k] === false && k !== 'locked') delete o[k]; else o[k] = patch[k]; } return o; };

  /* ---------- conditional formatting evaluation ---------- */
  /** returns {fill, color, b, bar:{pct,color}, scale} for a cell or null */
  SU.cfStats = new WeakMap();
  SU.evalCF = (sheet, r, c, version) => {
    if (!sheet.cf.length) return null;
    let out = null;
    for (const rule of sheet.cf) {
      if (!SU.inRange(rule.range, r, c)) continue;
      const v = sheet.value(r, c);
      const st = statsFor(sheet, rule, version);
      let hit = false;
      const num = typeof v === 'number' ? v : null; const txt = SU.valueText(v);
      switch (rule.type) {
        case 'gt': hit = num !== null && num > +rule.v1; break;
        case 'lt': hit = num !== null && num < +rule.v1; break;
        case 'between': hit = num !== null && num >= Math.min(+rule.v1, +rule.v2) && num <= Math.max(+rule.v1, +rule.v2); break;
        case 'eq': hit = num !== null && !isNaN(+rule.v1) ? num === +rule.v1 : txt.toLowerCase() === String(rule.v1).toLowerCase() && txt !== ''; break;
        case 'text': hit = txt !== '' && txt.toLowerCase().includes(String(rule.v1).toLowerCase()); break;
        case 'dup': hit = txt !== '' && st.counts.get(txt.toLowerCase()) > 1; break;
        case 'unique': hit = txt !== '' && st.counts.get(txt.toLowerCase()) === 1; break;
        case 'top': hit = num !== null && st.topT !== undefined && num >= st.topT; break;
        case 'bottom': hit = num !== null && st.botT !== undefined && num <= st.botT; break;
        case 'aboveavg': hit = num !== null && num > st.avg; break;
        case 'belowavg': hit = num !== null && num < st.avg; break;
        case 'formula': { try { const f = SU.shiftFormula(rule.formula, r - rule.range.r0, c - rule.range.c0); const res = sheet.wb.hf.calculateFormula(SU.toHF(f), sheet.id); hit = res === true || (typeof res === 'number' && res !== 0); } catch { } break; }
        case 'bar': if (num !== null && st.max > st.min0) { out = out || {}; out.bar = { pct: (num - st.min0) / (st.max - st.min0), color: rule.color || '#5B9BD5' }; } else if (num !== null && st.max === st.min0 && st.max > 0) { out = out || {}; out.bar = { pct: 1, color: rule.color || '#5B9BD5' }; } continue;
        case 'scale': if (num !== null) { out = out || {}; const cols = rule.colors || ['#F8696B', '#FFEB84', '#63BE7B']; out.fill = scaleColor(num, st.min, st.max, cols); } continue;
      }
      if (hit) { out = out || {}; const f = rule.fmt || {}; if (f.fill) out.fill = f.fill; if (f.color) out.color = f.color; if (f.b) out.b = true; if (rule.stop) break; }
    }
    return out;
  };
  function statsFor(sheet, rule, version) {
    let st = SU.cfStats.get(rule); if (st && st.v === version) return st;
    st = { v: version, counts: new Map() };
    const u = sheet.usedRange(); const g = { r0: rule.range.r0, c0: rule.range.c0, r1: Math.min(rule.range.r1, u.r1), c1: Math.min(rule.range.c1, u.c1) };
    const nums = [];
    for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) { const v = sheet.value(r, c); if (typeof v === 'number') nums.push(v); const t = SU.valueText(v).toLowerCase(); if (t !== '') st.counts.set(t, (st.counts.get(t) || 0) + 1); }
    if (nums.length) {
      st.min = Math.min(...nums); st.max = Math.max(...nums); st.min0 = Math.min(0, st.min); st.avg = nums.reduce((a, b) => a + b, 0) / nums.length;
      const sorted = nums.slice().sort((a, b) => b - a); let n = +rule.n || 10; if (rule.pct) n = Math.max(1, Math.floor(nums.length * n / 100));
      st.topT = sorted[Math.min(n, sorted.length) - 1]; st.botT = sorted[Math.max(0, sorted.length - n)];
    } else { st.min = st.max = st.min0 = st.avg = 0; }
    SU.cfStats.set(rule, st); return st;
  }
  function hex(c) { c = c.replace('#', ''); return [parseInt(c.slice(0, 2), 16), parseInt(c.slice(2, 4), 16), parseInt(c.slice(4, 6), 16)]; }
  function mix(a, b, t) { const A = hex(a), B = hex(b); return '#' + A.map((x, i) => Math.round(x + (B[i] - x) * t).toString(16).padStart(2, '0')).join(''); }
  function scaleColor(v, min, max, cols) {
    if (max === min) return cols[cols.length > 2 ? 1 : 0];
    const t = (v - min) / (max - min);
    if (cols.length === 2) return mix(cols[0], cols[1], t);
    return t < .5 ? mix(cols[0], cols[1], t * 2) : mix(cols[1], cols[2], (t - .5) * 2);
  }
  SU.mixColor = mix;

  window.SheetModel = { Workbook, Sheet, shiftRange, HF_CONFIG };
})();
