/* BTS Sheets — PivotTables.
 * A definition lives on the sheet that shows its output: s.pivots = [def]
 *   def = { id, name, src:{sheet, range} | {table}, dest:{r,c}, rows:[field], cols:[field], filters:[field], sel:{field:[labels]},
 *           values:[{field, agg, show, nf, name}], subtotals, grandRows, grandCols, styleId:'Light|1', out:{r0,c0,r1,c1} }
 * Output cells are plain values (saved as values in xlsx); definitions travel in the workbook metadata so Refresh works after reopening. */
(function () {
  const Pivot = window.Pivot = {};
  const { key } = SU;
  let wb;
  Pivot.init = (w) => { wb = w; };
  const S = () => wb.active;
  const esc = SU.esc;
  const SEP = '\u0001', SEP2 = '\u0002';
  Pivot.AGG = { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', countDistinct: 'Distinct Count' };
  Pivot.SHOW = { normal: 'No calculation', pctGrand: '% of Grand Total', pctCol: '% of Column Total', pctRow: '% of Row Total' };
  Pivot.NFS = [['', 'Same as source'], ['General', 'General'], ['#,##0', 'Number 1,234'], ['#,##0.00', 'Number 1,234.50'], [SU.currencyFormat('₹', 0), 'Currency ₹1,234'], [SU.currencyFormat('₹', 2), 'Currency ₹1,234.50'], ['0%', 'Percent 12%'], ['0.0%', 'Percent 12.5%']];
  const BLANK = '(blank)';

  /* ---------- source ---------- */
  Pivot.source = (def) => {
    let sh = null, g = null;
    if (def.src.table) { for (const s of wb.sheets) { const t = s.tables.find(x => x.name === def.src.table); if (t) { sh = s; g = { ...t.range }; if (t.total) g.r1--; break; } } }
    if (!sh) { sh = wb.sheetByName(def.src.sheet); g = def.src.range; }
    if (!sh || !g) return null;
    const u = sh.usedRange(); g = { ...g, r1: Math.min(g.r1, Math.max(u.r1, g.r0)), c1: Math.min(g.c1, Math.max(u.c1, g.c0)) };
    const fields = []; const seen = new Set();
    for (let c = g.c0; c <= g.c1; c++) { let n = String(sh.displayText(g.r0, c)).trim() || 'Column' + (c - g.c0 + 1); let k = n, i = 2; while (seen.has(k.toLowerCase())) k = n + i++; seen.add(k.toLowerCase()); fields.push(k); }
    const recs = []; const nums = fields.map(() => 0), filled = fields.map(() => 0); const nfs = fields.map(() => null), ints = fields.map(() => true);
    for (let r = g.r0 + 1; r <= g.r1; r++) {
      const row = []; let any = false;
      for (let c = g.c0; c <= g.c1; c++) {
        let v = sh.value(r, c); if (SU.isErr(v)) v = v.value; if (v === '') v = null;
        const t = v === null ? '' : sh.displayText(r, c); row.push({ v, t });
        if (v !== null) { any = true; filled[c - g.c0]++; if (typeof v === 'number') { nums[c - g.c0]++; if (!Number.isInteger(v)) ints[c - g.c0] = false; if (!nfs[c - g.c0]) nfs[c - g.c0] = sh.style(r, c).nf || 'General'; } }
      }
      if (any) recs.push(row);
    }
    return { sheet: sh, range: g, fields, recs, numeric: fields.map((_, i) => filled[i] > 0 && nums[i] >= filled[i] * 0.6), nfs, ints };
  };
  Pivot.srcLabel = (def) => def.src.table ? def.src.table : SU.quoteSheet(def.src.sheet) + '!' + SU.absRangeStr(def.src.range);
  Pivot.valueName = (v) => v.name || (Pivot.AGG[v.agg] || 'Sum') + ' of ' + v.field;

  /* ---------- compute ---------- */
  function newAcc() { return { n: 0, nn: 0, sum: 0, min: Infinity, max: -Infinity, set: null }; }
  function agg(a, type) {
    if (!a) return null;
    switch (type) {
      case 'count': return a.n;
      case 'countDistinct': return a.set ? a.set.size : 0;
      case 'average': return a.nn ? a.sum / a.nn : null;
      case 'max': return a.nn ? a.max : 0;
      case 'min': return a.nn ? a.min : 0;
      default: return a.nn ? +a.sum.toPrecision(15) : (a.n ? 0 : null);
    }
  }
  Pivot.compute = (def) => {
    const src = Pivot.source(def); if (!src) return { error: 'The source data for this PivotTable can’t be found.' };
    const fi = (n) => src.fields.indexOf(n);
    const R = def.rows.map(fi).filter(i => i >= 0), C = def.cols.map(fi).filter(i => i >= 0);
    const V = def.values.filter(v => fi(v.field) >= 0), Fl = (def.filters || []).filter(n => fi(n) >= 0);
    const label = (rec, i) => rec[i].t === '' ? BLANK : rec[i].t;
    const sel = def.sel || {};
    const recs = src.recs.filter(rec => Object.keys(sel).every(n => { const i = fi(n); return i < 0 || !sel[n] || sel[n].includes(label(rec, i)); }));
    const accs = new Map(); const rowTree = new Map(), colTree = new Map();
    const addTree = (tree, rec, idx) => { let node = tree; for (const i of idx) { const l = label(rec, i); let ch = node.get(l); if (!ch) { ch = { sort: rec[i].v === null ? null : rec[i].v, kids: new Map() }; node.set(l, ch); } node = ch.kids; } };
    for (const rec of recs) {
      addTree(rowTree, rec, R); addTree(colTree, rec, C);
      const rl = R.map(i => label(rec, i)), cl = C.map(i => label(rec, i));
      for (let a = 0; a <= R.length; a++) for (let b = 0; b <= C.length; b++) {
        const k = rl.slice(0, a).join(SEP) + SEP2 + cl.slice(0, b).join(SEP);
        let arr = accs.get(k); if (!arr) { arr = V.map(newAcc); accs.set(k, arr); }
        V.forEach((v, vi) => { const x = rec[fi(v.field)].v; const ac = arr[vi]; if (x === null) return; ac.n++; if (typeof x === 'number') { ac.nn++; ac.sum += x; if (x < ac.min) ac.min = x; if (x > ac.max) ac.max = x; } if (v.agg === 'countDistinct') { (ac.set = ac.set || new Set()).add(label(rec, fi(v.field))); } });
      }
    }
    const sorted = (m) => [...m.entries()].sort((a, b) => { if (a[0] === BLANK) return 1; if (b[0] === BLANK) return -1; return SU.cmp(a[1].sort, b[1].sort) || a[0].localeCompare(b[0]); });
    const slots = (tree, depth, grand) => {
      if (!depth) return [{ path: [], kind: 'all' }];
      const out = [];
      const walk = (node, path) => { for (const [l, ch] of sorted(node)) { const p = [...path, l]; if (p.length === depth) out.push({ path: p, kind: 'leaf' }); else { walk(ch.kids, p); if (def.subtotals !== false) out.push({ path: p, kind: 'sub' }); } } };
      walk(tree, []); if (grand) out.push({ path: [], kind: 'grand' });
      return out;
    };
    const rowSlots = slots(rowTree, R.length, def.grandRows !== false), colSlots = slots(colTree, C.length, def.grandCols !== false);
    const val = (rp, cp, vi) => { const v = V[vi]; const a = accs.get(rp.join(SEP) + SEP2 + cp.join(SEP)); let x = agg(a && a[vi], v.agg); if (x === null || v.show === 'normal' || !v.show) return x;
      const base = v.show === 'pctGrand' ? agg((accs.get(SEP2) || [])[vi], v.agg) : v.show === 'pctRow' ? agg((accs.get(rp.join(SEP) + SEP2) || [])[vi], v.agg) : agg((accs.get(SEP2 + cp.join(SEP)) || [])[vi], v.agg);
      return base ? x / base : null; };
    const nfOf = (v) => v.nf ? v.nf : (v.show && v.show !== 'normal') ? '0.0%' : (v.agg === 'count' || v.agg === 'countDistinct') ? '#,##0' : (src.nfs[fi(v.field)] && src.nfs[fi(v.field)] !== 'General' ? src.nfs[fi(v.field)] : v.agg === 'average' || !src.ints[fi(v.field)] ? '#,##0.00' : '#,##0');
    // ---- layout ----
    const cells = []; const put = (r, c, v, role, nf) => { (cells[r] = cells[r] || [])[c] = { v, role, nf }; };
    let row = 0;
    for (const n of Fl) { const s2 = sel[n]; put(row, 0, n, 'filterLabel'); put(row, 1, !s2 ? '(All)' : s2.length === 1 ? s2[0] : '(Multiple Items)', 'filterValue'); row++; }
    if (Fl.length) row++;
    const nR = Math.max(R.length, 1), nV = Math.max(V.length, 1);
    const top = row;
    const colLabel = (slot, level) => slot.kind === 'grand' ? (level === 0 ? 'Grand Total' : null) : slot.kind === 'sub' ? (level === slot.path.length - 1 ? slot.path[level] + ' Total' : level < slot.path.length - 1 ? slot.path[level] : null) : slot.path[level];
    if (C.length) {
      put(row, 0, V.length === 1 ? Pivot.valueName(V[0]) : V.length ? 'Values' : '', 'caption');
      put(row, nR, C.map(i => src.fields[i]).join(' / '), 'caption'); row++;
      for (let lv = 0; lv < C.length; lv++) {
        colSlots.forEach((sl, k) => { const prev = colSlots[k - 1]; const same = prev && prev.kind === 'leaf' && sl.kind !== 'grand' && prev.path.slice(0, lv + 1).join(SEP) === sl.path.slice(0, lv + 1).join(SEP) && !(sl.kind === 'sub' && lv === sl.path.length - 1); const t = same ? null : colLabel(sl, lv); for (let vi = 0; vi < nV; vi++) put(row, nR + k * nV + vi, vi === 0 ? t : null, 'head'); });
        for (let j = 0; j < nR; j++) put(row, j, null, 'head');
        row++;
      }
      if (V.length > 1) { colSlots.forEach((sl, k) => V.forEach((v, vi) => put(row, nR + k * nV + vi, Pivot.valueName(v), 'head'))); for (let j = 0; j < nR; j++) put(row, j, null, 'head'); row++; }
      R.forEach((i, j) => put(row - 1, j, src.fields[i], 'head')); if (!R.length) put(row - 1, 0, V.length > 1 ? '' : 'Total', 'head');
    } else {
      R.forEach((i, j) => put(row, j, src.fields[i], 'head')); if (!R.length) put(row, 0, '', 'head');
      V.forEach((v, vi) => put(row, nR + vi, Pivot.valueName(v), 'head'));
      row++;
    }
    const headerEnd = row; const leafRows = [];
    rowSlots.forEach((sl, k) => {
      const kind = sl.kind; const prev = rowSlots[k - 1];
      const roleL = kind === 'grand' ? 'grandLabel' : kind === 'sub' ? 'subLabel' : 'rowLabel', roleV = kind === 'grand' ? 'grand' : kind === 'sub' ? 'sub' : 'value';
      for (let j = 0; j < nR; j++) {
        let t = null;
        if (kind === 'grand') t = j === 0 ? 'Grand Total' : null;
        else if (kind === 'all') t = j === 0 ? (R.length ? null : (C.length ? (V.length === 1 ? Pivot.valueName(V[0]) : '') : 'Total')) : null;
        else if (kind === 'sub') t = j === sl.path.length - 1 ? sl.path[j] + ' Total' : null;
        else { const same = prev && prev.kind === 'leaf' && prev.path.slice(0, j + 1).join(SEP) === sl.path.slice(0, j + 1).join(SEP); t = same ? null : sl.path[j]; }
        put(row, j, t, roleL);
      }
      if (kind === 'leaf') leafRows.push({ row, path: sl.path });
      colSlots.forEach((cs, ck) => { const role = cs.kind === 'grand' && roleV === 'value' ? 'grandCol' : roleV; V.forEach((v, vi) => put(row, nR + ck * nV + vi, val(sl.path, cs.path, vi), role, nfOf(v))); if (!V.length) put(row, nR + ck, null, role); });
      row++;
    });
    const W = nR + colSlots.length * nV;
    return { cells, H: row, W, top, headerEnd, nR, nV, V, rowSlots, colSlots, leafRows, fields: src.fields, src, R, C };
  };

  /* ---------- styles (follow the workbook theme) ---------- */
  Pivot.STYLES = []; for (const fam of ['Light', 'Medium', 'Dark']) for (let i = 1; i <= 6; i++) Pivot.STYLES.push(fam + '|' + i);
  Pivot.roleStyle = (role, styleId, th = STheme.cur()) => {
    const [fam, n] = String(styleId || 'Light|1').split('|'); const acc = th.accents[(+n || 1) - 1]; const tint = STheme.tint;
    const B = (c) => ({ style: 'thin', color: c });
    if (role === 'caption') return { b: true, color: th.dk2 };
    if (role === 'filterLabel') return { b: true };
    if (role === 'filterValue') return { i: true, h: 'left' };
    if (role === 'placeholder') return { fill: tint(acc, 0.9), color: tint(acc, -0.45), i: true };
    if (fam === 'Medium') return { head: { b: true, fill: acc, color: STheme.textOn(acc) }, sub: { b: true, fill: tint(acc, 0.8) }, subLabel: { b: true, fill: tint(acc, 0.8) }, grand: { b: true, fill: acc, color: STheme.textOn(acc) }, grandLabel: { b: true, fill: acc, color: STheme.textOn(acc) }, grandCol: { b: true, fill: tint(acc, 0.8) }, value: {}, rowLabel: {} }[role] || {};
    if (fam === 'Dark') { const d = tint(acc, -0.5); return { head: { b: true, fill: d, color: '#FFFFFF' }, sub: { b: true, fill: tint(acc, 0.6) }, subLabel: { b: true, fill: tint(acc, 0.6) }, grand: { b: true, fill: d, color: '#FFFFFF' }, grandLabel: { b: true, fill: d, color: '#FFFFFF' }, grandCol: { b: true, fill: tint(acc, 0.75) }, value: { fill: tint(acc, 0.9) }, rowLabel: { fill: tint(acc, 0.9), b: true } }[role] || {}; }
    return { head: { b: true, fill: tint(acc, 0.8), color: '#000000', bB: B(acc) }, sub: { b: true, fill: tint(acc, 0.92) }, subLabel: { b: true, fill: tint(acc, 0.92) }, grand: { b: true, fill: tint(acc, 0.8), bT: B(acc) }, grandLabel: { b: true, fill: tint(acc, 0.8), bT: B(acc) }, grandCol: { b: true }, value: {}, rowLabel: {} }[role] || {};
  };
  Pivot.swatch = (id) => { const h = Pivot.roleStyle('head', id), sb = Pivot.roleStyle('sub', id), g = Pivot.roleStyle('grand', id), v = Pivot.roleStyle('value', id); return `<i style="background:${h.fill || '#fff'}"></i><i style="background:${v.fill || '#fff'}"></i><i style="background:${sb.fill || '#fff'}"></i><i style="background:${g.fill || '#fff'}"></i>`; };

  /* ---------- write output into the sheet (call inside a transaction) ---------- */
  Pivot.render = (sheet, def, opts = {}) => {
    const res = def.rows.length || def.cols.length || def.values.length ? Pivot.compute(def) : null;
    if (res && res.error) { if (!opts.silent) BTS.toast(res.error, { error: true }); return def; }
    const d = def.dest; const vals = new Map(), sts = new Map();
    if (def.out) for (let r = def.out.r0; r <= def.out.r1; r++) for (let c = def.out.c0; c <= def.out.c1; c++) { if (!sheet.isEmpty(r, c)) vals.set(key(r, c), null); if (sheet.styles.has(key(r, c))) sts.set(key(r, c), null); }
    let out;
    if (!res) { // empty report: placeholder
      out = { r0: d.r, c0: d.c, r1: d.r + 2, c1: d.c + 2 };
      const ph = Pivot.roleStyle('placeholder', def.styleId);
      for (let r = out.r0; r <= out.r1; r++) for (let c = out.c0; c <= out.c1; c++) sts.set(key(r, c), { ...ph, pv: def.id });
      vals.set(key(d.r, d.c), def.name); vals.set(key(d.r + 1, d.c), 'Choose fields from the'); vals.set(key(d.r + 2, d.c), 'PivotTable Fields list');
      sts.set(key(d.r, d.c), { ...ph, b: true, i: false, pv: def.id });
    } else {
      out = { r0: d.r, c0: d.c, r1: d.r + res.H - 1, c1: d.c + res.W - 1 };
      // refuse to overwrite other data
      if (!opts.force) for (let r = out.r0; r <= out.r1; r++) for (let c = out.c0; c <= out.c1; c++) {
        const inOld = def.out && SU.inRange(def.out, r, c);
        if (!inOld && !sheet.isEmpty(r, c)) { if (!opts.silent) BTS.toast('There isn’t enough room for the PivotTable — it would overwrite data at ' + SU.addr(r, c) + '. Move the data or the PivotTable.', { error: true, ms: 5000 }); return def; }
        if (sheet.pivots.some(p => p.id !== def.id && p.out && SU.inRange(p.out, r, c))) { if (!opts.silent) BTS.toast('A PivotTable can’t overlap another PivotTable.', { error: true }); return def; }
      }
      for (let i = 0; i < res.H; i++) for (let j = 0; j < res.W; j++) {
        const cell = res.cells[i] && res.cells[i][j]; const r = d.r + i, c = d.c + j; const k = key(r, c);
        const role = cell ? cell.role : (i < res.headerEnd && i >= res.top ? 'head' : null);
        if (cell && cell.v !== null && cell.v !== undefined) vals.set(k, cell.v);
        const st = role ? { ...Pivot.roleStyle(role, def.styleId) } : {};
        if (cell && cell.nf && typeof cell.v === 'number') st.nf = cell.nf;
        if (role && Object.keys(st).length) sts.set(k, { ...st, pv: def.id }); else if (role) sts.set(k, { pv: def.id });
      }
    }
    wb.setContents(sheet, [...vals].map(([k, v]) => { const p = SU.unkey(k); return { r: p.r, c: p.c, v }; }));
    for (const [k, st] of sts) { const p = SU.unkey(k); wb.setStyle(sheet, p.r, p.c, st); }
    const nd = { ...def, out };
    if (def.autofit !== false && res) {
      const m = new Map(sheet.colW);
      for (let c = out.c0; c <= out.c1; c++) { let w = 0; for (let r = out.r0; r <= out.r1; r++) { const t = r === out.r0 && c === out.c0 + res.nR && res.C.length ? '' : (vals.has(key(r, c)) ? sheet.displayText(r, c) : ''); if (t) w = Math.max(w, Grid.measure(Grid.fontStr(sheet.style(r, c), 1, typeof sheet.value(r, c) === 'number'), t) + 18); } if (w) m.set(c, Math.min(360, Math.max(SU.DEF_W, sheet.colW.get(c) || 0, Math.ceil(w)))); }
      wb.setProp(sheet, 'colW', m);
    }
    return nd;
  };
  /** apply a change to a definition and re-render (undoable) */
  Pivot.update = (sheet, id, patch, label = 'PivotTable') => {
    const old = sheet.pivots.find(p => p.id === id); if (!old) return;
    wb.run(label, () => { const def = Pivot.render(sheet, { ...old, ...patch }); wb.setProp(sheet, 'pivots', sheet.pivots.map(p => p.id === id ? def : p)); Pivot.syncCharts(def); });
    Grid.updateSpacer(); Grid.render(); Pivot.refreshPane();
  };
  Pivot.refresh = (sheet, def) => Pivot.update(sheet, def.id, {}, 'Refresh PivotTable');
  Pivot.refreshAll = (silent) => {
    let n = 0;
    wb.run('Refresh all', () => { for (const s of wb.sheets) { const list = s.pivots.map(p => { n++; return Pivot.render(s, p, { silent: true }); }); if (s.pivots.length) wb.setProp(s, 'pivots', list); list.forEach(Pivot.syncCharts); } });
    Grid.render(); Pivot.refreshPane(); if (!silent) BTS.toast(n ? `Refreshed ${n} PivotTable${n > 1 ? 's' : ''}` : 'This workbook has no PivotTables', { error: !n });
  };
  /** used by theme changes (inside the caller's transaction) */
  Pivot.restyleAll = () => { for (const s of wb.sheets) if (s.pivots.length) wb.setProp(s, 'pivots', s.pivots.map(p => Pivot.render(s, p, { silent: true, force: true }))); };
  Pivot.at = (s, r, c) => s.pivots.find(p => p.out && SU.inRange(p.out, r, c)) || null;
  Pivot.atActive = () => { const s = S(); return Pivot.at(s, s.active.r, s.active.c); };
  Pivot.find = (id) => { for (const s of wb.sheets) { const p = s.pivots.find(x => x.id === id); if (p) return { s, def: p }; } return null; };
  Pivot.nextName = () => { let n = 1; while (wb.sheets.some(s => s.pivots.some(p => p.name === 'PivotTable' + n))) n++; return 'PivotTable' + n; };

  /* ---------- create ---------- */
  Pivot.createDialog = async (fromTable) => {
    const s = S(); let srcStr;
    const t = fromTable || s.tableAt(s.active.r, s.active.c);
    if (t) srcStr = t.name; else { const g = A.dataRange(); if (g.r1 <= g.r0 && s.isEmpty(g.r0, g.c0)) { BTS.toast('Select a cell in your data (with a header row) first.', { error: true }); return; } srcStr = SU.quoteSheet(s.name) + '!' + SU.absRangeStr(g); }
    const res = await BTS.dialog({ title: 'Create PivotTable', width: 440, body: `<p class="dlg-note" style="margin:0 0 6px">Choose the data that you want to analyze</p>
      <div class="field"><label>Table/Range</label><input id="pv-src" value="${esc(srcStr)}"></div>
      <p class="dlg-note" style="margin:10px 0 4px">Choose where you want the PivotTable report to be placed</p>
      <label class="check"><input type="radio" name="pv-dest" value="new" checked> New worksheet</label>
      <label class="check"><input type="radio" name="pv-dest" value="here"> Existing worksheet</label>
      <div class="field"><label>Location</label><input id="pv-loc" value="${esc(SU.quoteSheet(s.name) + '!' + SU.addr(s.active.r, Math.max(0, (A.dataRange().c1 || 0) + 2)))}"></div>`,
      onOpen: (r) => { const loc = r.querySelector('#pv-loc'); const sync = () => { loc.disabled = r.querySelector('[name=pv-dest]:checked').value === 'new'; }; r.addEventListener('change', sync); sync(); },
      buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ src: r.querySelector('#pv-src').value.trim(), dest: r.querySelector('[name=pv-dest]:checked').value, loc: r.querySelector('#pv-loc').value }) }] });
    if (!res) return;
    let src;
    const tbl = wb.sheets.flatMap(x => x.tables).find(x => x.name.toLowerCase() === res.src.toLowerCase());
    if (tbl) src = { table: tbl.name };
    else { const ref = SU.parseRef(res.src); if (!ref || ref.range.r1 <= ref.range.r0) { BTS.toast('The data range must include a header row and at least one row of data.', { error: true, ms: 4000 }); return; } src = { sheet: ref.sheet ? (wb.sheetByName(ref.sheet) || {}).name : s.name, range: ref.range }; if (!src.sheet) { BTS.toast('That sheet doesn’t exist.', { error: true }); return; } }
    let dest = null;
    if (res.dest === 'here') { const ref = SU.parseRef(res.loc); if (!ref) { BTS.toast('The location reference isn’t valid.', { error: true }); return; } dest = { sheet: ref.sheet ? wb.sheetByName(ref.sheet) : s, r: ref.range.r0, c: ref.range.c0 }; if (!dest.sheet) { BTS.toast('That sheet doesn’t exist.', { error: true }); return; } }
    return Pivot.create(src, dest);
  };
  /** create a pivot; dest null → new sheet. Returns def */
  Pivot.create = (src, dest, fields = {}, label = 'Insert PivotTable') => {
    let def;
    wb.run(label, () => {
      let sheet = dest && dest.sheet;
      if (!sheet) { sheet = wb.addSheet(wb.uniqueName('Pivot'), wb.activeIndex + 1); }
      const d = dest ? { r: dest.r, c: dest.c } : { r: 2, c: 0 };
      def = { id: 'pv' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5), name: fields.name || Pivot.nextName(), src, dest: d, rows: [], cols: [], filters: [], values: [], sel: {}, subtotals: true, grandRows: true, grandCols: true, styleId: 'Light|1', ...fields };
      def = Pivot.render(sheet, def, { silent: !!fields.silent });
      wb.setProp(sheet, 'pivots', [...sheet.pivots, def]);
      if (!dest) wb.activeIndex = wb.sheets.indexOf(sheet);
      sheet.sel = { r0: d.r, c0: d.c, r1: d.r, c1: d.c }; sheet.active = { r: d.r, c: d.c };
    });
    App.renderTabs(); Grid.restoreScroll(); Charts.renderSheet(); Grid.notify();
    Pivot.paneClosed = false; Pivot.refreshPane();
    return def;
  };
  Pivot.remove = (def) => {
    const f = Pivot.find(def.id); if (!f) return; const s = f.s;
    wb.run('Delete PivotTable', () => {
      const cells = []; const o = def.out; if (o) for (let r = o.r0; r <= o.r1; r++) for (let c = o.c0; c <= o.c1; c++) { if (!s.isEmpty(r, c)) cells.push({ r, c, v: null }); if (s.styles.has(key(r, c))) wb.setStyle(s, r, c, null); }
      wb.setContents(s, cells); wb.setProp(s, 'pivots', s.pivots.filter(p => p.id !== def.id));
      for (const sh of wb.sheets) if (sh.charts.some(ch => ch.pivot === def.id)) wb.setProp(sh, 'charts', sh.charts.map(ch => ch.pivot === def.id ? { ...ch, pivot: undefined } : ch));
    });
    Grid.render(); Pivot.refreshPane(); Charts.renderSheet();
  };

  /* ---------- PivotChart ---------- */
  Pivot.chartData = (id) => {
    const f = Pivot.find(id); if (!f) return null; const def = f.def; if (!def.values.length) return null;
    const res = Pivot.compute(def); if (res.error) return null;
    const leaves = res.rowSlots.filter(x => x.kind === 'leaf' || x.kind === 'all'); const colLeaves = res.colSlots.map((x, i) => ({ x, i })).filter(o => o.x.kind === 'leaf' || o.x.kind === 'all');
    const labels = leaves.map(x => x.path.join(' / ') || 'Total');
    const series = [];
    const rowIndex = new Map(res.rowSlots.map((sl, i) => [sl, i]));
    for (const { x: cs, i: ck } of colLeaves) res.V.forEach((v, vi) => {
      const name = (cs.path.length ? cs.path.join(' / ') : '') + (res.V.length > 1 || !cs.path.length ? (cs.path.length ? ' – ' : '') + Pivot.valueName(v) : '');
      const data = leaves.map(sl => { const rr = res.headerEnd + rowIndex.get(sl); const cell = res.cells[rr] && res.cells[rr][res.nR + ck * res.nV + vi]; return cell && typeof cell.v === 'number' ? cell.v : null; });
      series.push({ name, data });
    });
    return { labels, series, autoTitle: res.V.length === 1 ? Pivot.valueName(res.V[0]) : '' };
  };
  Pivot.syncCharts = (def) => { for (const s of wb.sheets) for (const ch of s.charts) if (ch.pivot === def.id && def.out) ch.range = { ...def.out }; };
  Pivot.createChart = (type = 'column') => {
    const def = Pivot.atActive(); if (!def) { BTS.toast('Select a cell in a PivotTable first.', { error: true }); return; }
    if (!def.values.length) { BTS.toast('Add a field to Values first.', { error: true }); return; }
    const s = S(); const o = def.out;
    const ch = { id: 'ch' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5), type, range: { ...o }, sheet: null, pivot: def.id, x: s.colLeft(o.c1 + 1) + 24, y: s.rowTop(o.r0), w: 500, h: 300, z: 1, title: Pivot.valueName(def.values[0]) + (def.rows.length ? ' by ' + def.rows.join(', ') : '') };
    wb.run('Insert PivotChart', () => wb.setProp(s, 'charts', [...s.charts, ch]));
    Charts.renderSheet(); Charts.select(ch.id);
  };

  /* ---------- field list pane ---------- */
  Pivot.paneClosed = false;
  const AREAS = [['filters', 'Filters', 'funnel'], ['cols', 'Columns', 'columns-3'], ['rows', 'Rows', 'rows-3'], ['values', 'Values', 'sigma']];
  function paneEl() {
    let el = document.getElementById('pv-pane'); if (el) return el;
    el = document.createElement('aside'); el.className = 'pane side-pane hidden'; el.id = 'pv-pane';
    el.innerHTML = `<header><span>PivotTable Fields</span><button class="rb" data-cmd="pv-pane-close" data-icon="x" title="Close the field list">${icon('x', 16)}</button></header>
      <div class="pane-body pv-body"><div class="pv-hint">Choose fields to add to the report:</div><div class="pv-search">${icon('search', 14)}<input placeholder="Search fields" spellcheck="false"></div><div class="pv-fields"></div>
      <div class="pv-hint">Drag fields between areas below:</div><div class="pv-areas">${AREAS.map(([k, l, ic]) => `<div class="pv-area" data-area="${k}"><div class="pv-area-h">${icon(ic, 14)}<span>${l}</span></div><div class="pv-drop" data-area="${k}"></div></div>`).join('')}</div>
      <div class="pv-foot"><span class="pv-src"></span><button class="btn" data-cmd="pv-refresh" title="Refresh the PivotTable from its source data">${icon('refresh-cw', 14)} Refresh</button></div></div>`;
    document.getElementById('workspace').appendChild(el);
    wirePane(el); return el;
  }
  Pivot.refreshPane = () => {
    const s = S(); const def = Pivot.atActive();
    const el = document.getElementById('pv-pane');
    if (!def || Pivot.paneClosed) { if (el && !el.classList.contains('hidden')) { el.classList.add('hidden'); App.layoutPanes(); } return; }
    const p = paneEl(); const wasHidden = p.classList.contains('hidden');
    if (wasHidden && window.Threads) Threads.close(true);
    p.classList.remove('hidden'); p.dataset.id = def.id; p.dataset.sheet = s.name;
    const src = Pivot.source(def); const fields = src ? src.fields : [];
    const used = new Set([...def.rows, ...def.cols, ...def.filters, ...def.values.map(v => v.field)]);
    const q = (p.querySelector('.pv-search input').value || '').toLowerCase();
    p.querySelector('.pv-fields').innerHTML = fields.filter(f => !q || f.toLowerCase().includes(q)).map(f => `<label class="pv-field" draggable="true" data-field="${esc(f)}" title="Drag to an area, or tick to add"><input type="checkbox" ${used.has(f) ? 'checked' : ''}><span>${esc(f)}</span>${src && src.numeric[fields.indexOf(f)] ? '<small>123</small>' : ''}${def.sel && def.sel[f] ? icon('funnel', 12) : ''}</label>`).join('') || '<div class="empty-note">No fields. Check the source data has a header row.</div>';
    for (const [k] of AREAS) {
      const list = k === 'values' ? def.values.map(v => ({ f: v.field, label: Pivot.valueName(v) })) : def[k].map(f => ({ f, label: f }));
      p.querySelector(`.pv-drop[data-area="${k}"]`).innerHTML = list.map((it, i) => `<div class="pv-chip" draggable="true" data-area="${k}" data-i="${i}" title="Click for options, drag to move"><span>${esc(it.label)}</span>${def.sel && def.sel[it.f] && k !== 'values' ? icon('funnel', 12) : ''}${icon('chevron-down', 13)}</div>`).join('');
    }
    p.querySelector('.pv-src').textContent = 'Source: ' + Pivot.srcLabel(def);
    if (wasHidden) App.layoutPanes();
  };
  function curDef() { const p = document.getElementById('pv-pane'); if (!p) return null; const f = Pivot.find(p.dataset.id); return f; }
  function moveField(def, field, fromArea, fromIdx, toArea, toIdx) {
    const d = { rows: def.rows.slice(), cols: def.cols.slice(), filters: def.filters.slice(), values: def.values.map(v => ({ ...v })) };
    if (fromArea && fromArea === toArea) { const list = d[toArea]; const [it] = list.splice(fromIdx, 1); const t = toIdx === undefined ? list.length : Math.max(0, Math.min(toIdx, list.length)); list.splice(t, 0, it); return d; }
    let moving = null;
    if (fromArea) { if (fromArea === 'values') { moving = d.values.splice(fromIdx, 1)[0]; } else d[fromArea].splice(fromIdx, 1); }
    if (toArea === 'values') { const src = Pivot.source(def); const num = src && src.numeric[src.fields.indexOf(field)]; const v = moving && fromArea === 'values' ? moving : { field, agg: num ? 'sum' : 'count', show: 'normal' }; d.values.splice(toIdx === undefined ? d.values.length : toIdx, 0, v); }
    else if (toArea) { for (const a of ['rows', 'cols', 'filters']) { const i = d[a].indexOf(field); if (i >= 0) { d[a].splice(i, 1); if (a === toArea && toIdx !== undefined && i < toIdx) toIdx--; } } d[toArea].splice(toIdx === undefined ? d[toArea].length : toIdx, 0, field); }
    return d;
  }
  function wirePane(el) {
    el.querySelector('.pv-search input').addEventListener('input', () => Pivot.refreshPane());
    el.addEventListener('mousedown', (e) => { if (!e.target.closest('input')) e.preventDefault(); });
    el.querySelector('.pv-fields').addEventListener('change', (e) => {
      const lab = e.target.closest('[data-field]'); if (!lab) return; const f = curDef(); if (!f) return; const field = lab.dataset.field; const def = f.def;
      if (e.target.checked) { const src = Pivot.source(def); const num = src && src.numeric[src.fields.indexOf(field)]; Pivot.update(f.s, def.id, moveField(def, field, null, null, num ? 'values' : 'rows'), 'Add field'); }
      else Pivot.update(f.s, def.id, { rows: def.rows.filter(x => x !== field), cols: def.cols.filter(x => x !== field), filters: def.filters.filter(x => x !== field), values: def.values.filter(v => v.field !== field) }, 'Remove field');
    });
    el.addEventListener('dragstart', (e) => {
      const fl = e.target.closest('[data-field]'), ch = e.target.closest('.pv-chip'); const f = curDef(); if (!f) return;
      let data;
      if (ch) { const area = ch.dataset.area, i = +ch.dataset.i; data = { field: area === 'values' ? f.def.values[i].field : f.def[area][i], from: area, idx: i }; }
      else if (fl) data = { field: fl.dataset.field, from: null };
      if (!data) return; e.dataTransfer.setData('text/x-bts-field', JSON.stringify(data)); e.dataTransfer.effectAllowed = 'move';
    });
    el.addEventListener('dragover', (e) => { const d = e.target.closest('.pv-drop,.pv-fields'); if (!d) return; e.preventDefault(); el.querySelectorAll('.pv-drop.over').forEach(x => x.classList.remove('over')); if (d.classList.contains('pv-drop')) d.classList.add('over'); });
    el.addEventListener('dragleave', (e) => { const d = e.target.closest('.pv-drop'); if (d && !d.contains(e.relatedTarget)) d.classList.remove('over'); });
    el.addEventListener('drop', (e) => {
      el.querySelectorAll('.pv-drop.over').forEach(x => x.classList.remove('over'));
      const raw = e.dataTransfer.getData('text/x-bts-field'); if (!raw) return; e.preventDefault(); const data = JSON.parse(raw); const f = curDef(); if (!f) return; const def = f.def;
      const zone = e.target.closest('.pv-drop'); const toFields = e.target.closest('.pv-fields');
      if (toFields && data.from) { // drag out = remove
        const d = moveField(def, data.field, data.from, data.idx, null); Pivot.update(f.s, def.id, d, 'Remove field'); return;
      }
      if (!zone) return; const to = zone.dataset.area; const chip = e.target.closest('.pv-chip'); const idx = chip ? +chip.dataset.i : undefined;
      if (data.from === to && idx === data.idx) return;
      Pivot.update(f.s, def.id, moveField(def, data.field, data.from, data.idx, to, idx), 'Move field');
    });
    el.addEventListener('click', (e) => {
      const ch = e.target.closest('.pv-chip'); if (!ch) return; const f = curDef(); if (!f) return; const def = f.def; const area = ch.dataset.area, i = +ch.dataset.i;
      const field = area === 'values' ? def.values[i].field : def[area][i]; const list = area === 'values' ? def.values : def[area];
      const mv = (to) => Pivot.update(f.s, def.id, moveField(def, field, area, i, to), 'Move field');
      BTS.menu(ch, [
        { label: 'Move up', icon: 'arrow-up', disabled: i === 0, action: () => Pivot.update(f.s, def.id, moveField(def, field, area, i, area, i - 1), 'Move field') },
        { label: 'Move down', icon: 'arrow-down', disabled: i === list.length - 1, action: () => Pivot.update(f.s, def.id, moveField(def, field, area, i, area, i + 1), 'Move field') },
        { sep: true },
        ...AREAS.filter(a => a[0] !== area).map(([k, l, ic]) => ({ label: 'Move to ' + l, icon: ic, action: () => mv(k) })),
        { sep: true },
        ...(area !== 'values' ? [{ label: 'Filter “' + field + '”…', icon: 'funnel', action: () => Pivot.filterPopover(ch, f.s, def, field) }] : []),
        ...(area === 'values' ? [{ label: 'Value field settings…', icon: 'settings-2', action: () => Pivot.valueSettings(f.s, def, i) }] : []),
        { label: 'Remove field', icon: 'x', action: () => Pivot.update(f.s, def.id, moveField(def, field, area, i, null), 'Remove field') },
      ]);
    });
  }
  Pivot.filterPopover = (anchor, sheet, def, field) => {
    const src = Pivot.source(def); if (!src) return; const i = src.fields.indexOf(field);
    const labels = [...new Map(src.recs.map(r => [r[i].t === '' ? BLANK : r[i].t, r[i].v])).entries()].sort((a, b) => a[0] === BLANK ? 1 : b[0] === BLANK ? -1 : SU.cmp(a[1], b[1])).map(x => x[0]);
    const cur = new Set((def.sel && def.sel[field]) || labels);
    const el = document.createElement('div'); el.className = 'filter-pop';
    el.innerHTML = `<div class="gal-head" style="padding:4px 6px 6px">Show items: ${esc(field)}</div><div class="fp-list"><label><input type="checkbox" data-all> (Select all)</label>${labels.map((t, k) => `<label><input type="checkbox" data-v="${k}" ${cur.has(t) ? 'checked' : ''}> ${esc(t)}</label>`).join('')}</div><div class="fp-foot"><button class="btn" data-a="cancel">Cancel</button><button class="btn primary" data-a="ok">OK</button></div>`;
    const boxes = [...el.querySelectorAll('[data-v]')]; const all = el.querySelector('[data-all]');
    const sync = () => { all.checked = boxes.every(b => b.checked); all.indeterminate = !all.checked && boxes.some(b => b.checked); }; sync();
    all.addEventListener('change', () => boxes.forEach(b => b.checked = all.checked)); el.addEventListener('change', (e) => { if (e.target.dataset.v !== undefined) sync(); });
    el.addEventListener('mousedown', (e) => { if (e.target.tagName !== 'INPUT') e.preventDefault(); });
    el.addEventListener('click', (e) => { const a = e.target.closest('[data-a]'); if (!a) return; if (a.dataset.a === 'cancel') { BTS.closePopover(); return; } const chosen = boxes.filter(b => b.checked).map(b => labels[+b.dataset.v]); BTS.closePopover(); if (!chosen.length) { BTS.toast('Select at least one item.', { error: true }); return; } const sel = { ...(def.sel || {}) }; if (chosen.length === labels.length) delete sel[field]; else sel[field] = chosen; Pivot.update(sheet, def.id, { sel }, 'Filter PivotTable'); });
    BTS.popover(anchor, el);
  };
  Pivot.valueSettings = async (sheet, def, i) => {
    const v = def.values[i];
    const res = await BTS.dialog({ title: 'Value field settings', width: 420, body: `<div class="dlg-grid">
      <label>Source field</label><b>${esc(v.field)}</b>
      <label>Custom name</label><input id="vf-name" value="${esc(v.name || '')}" placeholder="${esc(Pivot.valueName({ ...v, name: '' }))}">
      <label>Summarize by</label><select id="vf-agg">${Object.entries(Pivot.AGG).map(([k, l]) => `<option value="${k}" ${v.agg === k ? 'selected' : ''}>${l}</option>`).join('')}</select>
      <label>Show values as</label><select id="vf-show">${Object.entries(Pivot.SHOW).map(([k, l]) => `<option value="${k}" ${(v.show || 'normal') === k ? 'selected' : ''}>${l}</option>`).join('')}</select>
      <label>Number format</label><select id="vf-nf">${Pivot.NFS.map(([k, l]) => `<option value="${esc(k)}" ${(v.nf || '') === k ? 'selected' : ''}>${l}</option>`).join('')}</select></div>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ name: r.querySelector('#vf-name').value.trim(), agg: r.querySelector('#vf-agg').value, show: r.querySelector('#vf-show').value, nf: r.querySelector('#vf-nf').value }) }] });
    if (!res) return;
    const values = def.values.map((x, k) => k === i ? { ...x, name: res.name || undefined, agg: res.agg, show: res.show, nf: res.nf || undefined } : x);
    Pivot.update(sheet, def.id, { values }, 'Value field settings');
  };
  Pivot.changeSource = async () => {
    const def = Pivot.atActive(); if (!def) return; const s = S();
    const v = await BTS.prompt('Change PivotTable data source', 'Table name or range (e.g. Sales!A1:F200)', Pivot.srcLabel(def)); if (!v) return;
    const tbl = wb.sheets.flatMap(x => x.tables).find(x => x.name.toLowerCase() === v.trim().toLowerCase()); let src;
    if (tbl) src = { table: tbl.name }; else { const ref = SU.parseRef(v); if (!ref) { BTS.toast('That reference isn’t valid.', { error: true }); return; } src = { sheet: ref.sheet ? (wb.sheetByName(ref.sheet) || {}).name : s.name, range: ref.range }; if (!src.sheet) { BTS.toast('That sheet doesn’t exist.', { error: true }); return; } }
    Pivot.update(s, def.id, { src }, 'Change data source');
  };
  Pivot.options = async () => {
    const def = Pivot.atActive(); if (!def) return; const s = S();
    const res = await BTS.dialog({ title: 'PivotTable options', width: 400, body: `<div class="field"><label>PivotTable name</label><input id="po-name" value="${esc(def.name)}"></div>
      <label class="check"><input type="checkbox" id="po-sub" ${def.subtotals !== false ? 'checked' : ''}> Show subtotals</label>
      <label class="check"><input type="checkbox" id="po-gr" ${def.grandRows !== false ? 'checked' : ''}> Grand totals for rows</label>
      <label class="check"><input type="checkbox" id="po-gc" ${def.grandCols !== false ? 'checked' : ''}> Grand totals for columns</label>
      <label class="check"><input type="checkbox" id="po-fit" ${def.autofit !== false ? 'checked' : ''}> Autofit column widths on update</label>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ name: r.querySelector('#po-name').value.trim() || def.name, subtotals: r.querySelector('#po-sub').checked, grandRows: r.querySelector('#po-gr').checked, grandCols: r.querySelector('#po-gc').checked, autofit: r.querySelector('#po-fit').checked }) }] });
    if (res) Pivot.update(s, def.id, res, 'PivotTable options');
  };
  Pivot.styleGallery = (anchor) => {
    const def = Pivot.atActive(); if (!def) return; const s = S();
    const pop = document.createElement('div');
    pop.innerHTML = ['Light', 'Medium', 'Dark'].map(fam => `<div class="gal-head">${fam}</div><div class="gal tables">${Pivot.STYLES.filter(x => x.startsWith(fam)).map(id => `<button title="PivotTable style ${fam} ${id.split('|')[1]}" data-id="${id}" class="${def.styleId === id ? 'on' : ''}">${Pivot.swatch(id)}</button>`).join('')}</div>`).join('');
    pop.addEventListener('mousedown', e => e.preventDefault());
    pop.addEventListener('click', e => { const b = e.target.closest('[data-id]'); if (b) { BTS.closePopover(); Pivot.update(s, def.id, { styleId: b.dataset.id }, 'PivotTable style'); } });
    BTS.popover(anchor, pop);
  };
})();
