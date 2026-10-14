/* BTS Sheets — actions: entering values, navigation, formatting, clipboard, fill, sort, filter, merge, structure, CF, tables */
(function () {
  const { key, unkey } = SU;
  const A = window.A = {};
  let wb;
  A.init = (workbook) => { wb = workbook; };
  const S = () => wb.active;
  A.setMode = (m) => BTS.setStatus('mode', `<span class="sb-mode">${m}</span>`);

  /* ---------- protection ---------- */
  A.protectedCheck = (silent, g) => {
    const s = S(); if (!s.protected) return false;
    if (g === undefined) { const st = s.style(s.active.r, s.active.c); if (st.locked === false && !g) return false; }
    if (!silent) BTS.toast('This sheet is protected. Unprotect it on the Review tab to make changes.', { error: true, ms: 3200 });
    return true;
  };
  function guard() { return A.protectedCheck(false, true); }

  /* ---------- effective ranges ---------- */
  A.eff = (g, s = S()) => {
    const u = s.usedRangeAll(); const vis = Grid.visible || { r1: 50, c1: 20 };
    let { r0, c0, r1, c1 } = g;
    if (r1 >= SU.MAXR - 1) r1 = Math.max(u.r1, vis.r1, c1 - c0 < 30 ? 999 : 0, r0);
    if (c1 >= SU.MAXC - 1) c1 = Math.max(u.c1, vis.c1, r1 - r0 < 30 ? 99 : 0, c0);
    return { r0, c0, r1, c1 };
  };
  A.effData = (g, s = S()) => { const u = s.usedRange(); return { r0: g.r0, c0: g.c0, r1: Math.min(g.r1, Math.max(u.r1, g.r0)), c1: Math.min(g.c1, Math.max(u.c1, g.c0)) }; };
  A.sels = () => Grid.selections();
  A.forEachCell = (fn, ranges = A.sels().map(g => A.eff(g))) => { for (const g of ranges) for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) fn(r, c); };

  /* ---------- entering values ---------- */
  function balanceParens(f) {
    let depth = 0, q = false; for (const ch of f) { if (ch === '"') q = !q; else if (!q) { if (ch === '(') depth++; else if (ch === ')') depth--; } }
    if (q) f += '"'; if (depth > 0) f += ')'.repeat(depth); return f;
  }
  function normalizeFormula(f) {
    f = balanceParens(f.trim());
    // uppercase function names and refs outside strings
    f = SU.outsideStrings(f, p => p.replace(/([A-Za-z_][A-Za-z0-9_.]*)(?=\()/g, m => m.toUpperCase()));
    f = SU.mapRefs(f, (ref) => SU.refToString(ref).replace(/^(.*!)?/, (m) => m));
    return f;
  }
  A.normalizeFormula = normalizeFormula;
  /** turns typed text into raw content + optional nf; returns {v, nf, error} */
  A.interpret = (s, r, c, text) => {
    const st = s.style(r, c);
    if (st.nf === '@' && text !== '' && text[0] !== '=') return { v: text };
    const p = SU.parseInput(text);
    if (p.formula) {
      const f = normalizeFormula(p.formula);
      if (!wb.hf.validateFormula(SU.toHF(f))) return { error: f };
      return { v: f };
    }
    if (p.value === null) return { v: null };
    if (typeof p.value === 'number') {
      let v = p.value, nf;
      const cur = st.nf;
      if (cur && /%/.test(cur) && !/%/.test(text) && SU.formatId(cur) === 'percent') v = v / 100;
      if (p.nf && (!cur || cur === 'General' || (SU.isDateFormat(p.nf) !== SU.isDateFormat(cur)))) nf = p.nf;
      if (p.nf && /%/.test(p.nf) && cur && /%/.test(cur)) nf = undefined;
      return { v, nf };
    }
    if (typeof p.value === 'boolean') return { v: p.value };
    return { v: p.text && text[0] === "'" ? text.slice(1) : String(p.value) };
  };
  A.enterValue = (s, r, c, text, opts = {}) => {
    const res = A.interpret(s, r, c, text);
    if (res.error) {
      BTS.dialog({ title: 'There’s a problem with this formula', body: `<p style="max-width:440px;line-height:1.5;margin:4px 0">Check the formula for typos, unmatched brackets or missing arguments.<br><br><code style="background:var(--surface-3);padding:2px 6px;border-radius:4px">${SU.esc(res.error)}</code></p><p class="dlg-note">Not trying to type a formula? Start the entry with an apostrophe (') to enter it as text.</p>`, buttons: [{ label: 'OK', primary: true, value: true }] }).then(() => { const el = Ed.state && (Ed.state.src === 'bar' ? document.getElementById('fx') : document.getElementById('cell-editor')); el && el.focus(); });
      return false;
    }
    // validation
    if (res.v !== null) {
      const dv = s.validationAt(r, c);
      if (dv && !A.checkValidation(s, dv, res.v)) { A.validationError(dv); if ((dv.errorStyle || 'stop') === 'stop') return false; }
    }
    const targets = opts.ctrlEnter ? A.sels().map(g => A.eff(g)) : [{ r0: r, c0: c, r1: r, c1: c }];
    wb.run('Typing', () => {
      const cells = [];
      for (const g of targets) for (let rr = g.r0; rr <= g.r1; rr++) for (let cc = g.c0; cc <= g.c1; cc++) {
        if (opts.ctrlEnter && s.isRowHidden(rr)) continue;
        let v = res.v; if (typeof v === 'string' && v[0] === '=' && (rr !== r || cc !== c)) v = SU.shiftFormula(v, rr - r, cc - c);
        cells.push({ r: rr, c: cc, v });
        if (res.nf) wb.setStyle(s, rr, cc, SU.mergeStyle(s.style(rr, cc), { nf: res.nf }));
        // wrap automatically when text contains a newline
        if (typeof v === 'string' && v.includes('\n') && !s.style(rr, cc).wrap) wb.setStyle(s, rr, cc, SU.mergeStyle(s.style(rr, cc), { wrap: true }));
      }
      wb.setContents(s, cells);
      // hyperlink detection
      if (typeof res.v === 'string' && /^(https?:\/\/|www\.)\S+$/i.test(res.v) && !s.links.has(key(r, c))) { const m = new Map(s.links); m.set(key(r, c), res.v.startsWith('www.') ? 'https://' + res.v : res.v); wb.setProp(s, 'links', m); }
      // auto-grow row height for multi-line text
      if (typeof res.v === 'string' && res.v.includes('\n') && !s.rowH.has(r)) { const lines = res.v.split('\n').length; const m = new Map(s.rowH); m.set(r, Math.max(SU.DEF_H, Math.ceil(lines * (s.style(r, c).sz || 11) * 96 / 72 * 1.22 + 6))); wb.setProp(s, 'rowH', m); }
    });
    A.extendTablesAfterEntry(s, r, c);
    autoWiden(s, r, c);
    return true;
  };
  function autoWiden(s, r, c) {
    if (s.colW.has(c) || s.mergeAt(r, c)) return; const v = s.value(r, c); if (typeof v !== 'number' || s.isFormula(r, c)) return;
    const st = s.style(r, c); if (!st.nf || st.nf === 'General') return;
    const t = s.displayText(r, c); const w = Math.ceil(Grid.measure(Grid.fontStr(st, 1), t) + 10);
    if (w > s.colWidth(c)) { const m = new Map(s.colW); m.set(c, w); wb.run('Widen', () => wb.setProp(s, 'colW', m)); mergeLastTwo(); Grid.updateSpacer(); }
  }
  A.extendTablesAfterEntry = (s, r, c) => {
    for (const t of s.tables) {
      const g = t.range;
      if (r === g.r1 + 1 && c >= g.c0 && c <= g.c1 && !t.total) { wb.run('Extend table', () => { const nt = s.tables.map(x => x === t ? { ...x, range: { ...g, r1: r } } : x); wb.setProp(s, 'tables', nt); if (s.filter && s.filter.tableId === t.id) wb.setProp(s, 'filter', { ...s.filter, range: { ...s.filter.range, r1: r } }); }); wb.undoStack.length > 1 && mergeLastTwo(); return; }
    }
  };
  function mergeLastTwo() { const b = wb.undoStack.pop(), a = wb.undoStack.pop(); a.ops.push(...b.ops); a.selAfter = b.selAfter; wb.undoStack.push(a); }

  /* ---------- validation ---------- */
  A.checkValidation = (s, dv, v) => {
    if (v === null || v === '') return dv.ignoreBlank !== false;
    const cmp = (x) => {
      const a = +dv.v1, b = +dv.v2;
      switch (dv.op || 'between') { case 'between': return x >= Math.min(a, b) && x <= Math.max(a, b); case 'notBetween': return x < Math.min(a, b) || x > Math.max(a, b); case 'eq': return x === a; case 'ne': return x !== a; case 'gt': return x > a; case 'lt': return x < a; case 'ge': return x >= a; case 'le': return x <= a; }
      return true;
    };
    switch (dv.type) {
      case 'whole': return typeof v === 'number' && Number.isInteger(v) && cmp(v);
      case 'decimal': return typeof v === 'number' && cmp(v);
      case 'date': return typeof v === 'number' && cmp(v);
      case 'textLength': return cmp(String(v).length);
      case 'list': { const items = A.listItems(s, dv).map(x => String(x).toLowerCase()); return items.includes(String(SU.valueText(v)).toLowerCase()); }
      default: return true;
    }
  };
  A.listItems = (s, dv) => {
    const src = String(dv.source || '').trim();
    if (src[0] === '=') {
      const ref = SU.parseRef(src.slice(1));
      if (ref) { const sh = ref.sheet ? wb.sheetByName(ref.sheet) : s; if (!sh) return []; const g = A.effData(ref.range, sh); const out = []; for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) { const t = sh.displayText(r, c); if (t !== '') out.push(t); } return out; }
      try { const n = wb.hf.getNamedExpressionValue(src.slice(1)); if (n && n.data) return n.data.flat().filter(x => x !== null && x !== '').map(String); } catch { }
      return [];
    }
    return src.split(',').map(x => x.trim()).filter(Boolean);
  };
  A.validationError = (dv) => {
    BTS.dialog({ title: dv.errorTitle || 'Invalid entry', body: `<p style="max-width:420px;line-height:1.5;margin:4px 0">${SU.esc(dv.errorMsg || 'This value doesn’t match the data validation restrictions defined for this cell.')}</p>`, buttons: [{ label: (dv.errorStyle || 'stop') === 'stop' ? 'Retry' : 'OK', primary: true, value: true }] }).then(() => { const el = Ed.state && document.getElementById(Ed.state.src === 'bar' ? 'fx' : 'cell-editor'); el && el.focus(); });
  };

  /* ---------- navigation ---------- */
  A.selEnd = () => { const s = S(); const g = s.sel, a = s.active; return { r: g.r0 === a.r || (g.r0 <= a.r && g.r1 > a.r && s.mergeAt(a.r, a.c) && s.mergeAt(a.r, a.c).r0 === g.r0) ? g.r1 : g.r0, c: g.c0 === a.c ? g.c1 : g.c0, }; };
  A.jumpEdge = (s, cur, dr, dc) => {
    const used = s.usedRange(); const maxR = Math.max(s.rowCount - 1, used.r1), maxC = Math.max(s.colCount - 1, used.c1);
    const empty = (r, c) => s.isEmpty(r, c);
    let r = cur.r, c = cur.c;
    const inb = (r, c) => r >= 0 && c >= 0 && r <= maxR && c <= maxC;
    const next = () => { r += dr; c += dc; while (inb(r, c) && (dr ? s.isRowHidden(r) : !s.colWidth(c))) { r += dr; c += dc; } };
    const nr = r + dr, nc = c + dc;
    if (!inb(nr, nc)) return { r, c };
    if (!empty(r, c) && !empty(nr, nc)) { let pr = r, pc = c; next(); while (inb(r, c) && !empty(r, c)) { pr = r; pc = c; next(); } return { r: pr, c: pc }; }
    next(); while (inb(r, c) && empty(r, c)) { if ((dr > 0 && r > used.r1) || (dc > 0 && c > used.c1)) { return dr ? { r: dr > 0 ? maxR : 0, c } : { r, c: dc > 0 ? maxC : 0 }; } next(); }
    if (!inb(r, c)) return { r: SU.clamp(r, 0, maxR), c: SU.clamp(c, 0, maxC) };
    return { r, c };
  };
  function stepVisible(s, r, c, dr, dc) {
    let nr = r, nc = c;
    const m = s.mergeAt(r, c); if (m) { if (dr > 0) nr = m.r1; if (dc > 0) nc = m.c1; if (dr < 0) nr = m.r0; if (dc < 0) nc = m.c0; }
    do { nr += dr; nc += dc; } while (nr > 0 && nr < SU.MAXR - 1 && dr && s.isRowHidden(nr) || nc > 0 && nc < SU.MAXC - 1 && dc && !s.colWidth(nc));
    return { r: SU.clamp(nr, 0, SU.MAXR - 1), c: SU.clamp(nc, 0, SU.MAXC - 1) };
  }
  A.moveSel = (dr, dc, shift, ctrl) => {
    const s = S();
    if (shift) { const end = A.selEnd(); const n = ctrl ? A.jumpEdge(s, end, dr, dc) : stepVisible(s, end.r, end.c, dr, dc); Grid.setSel(s.active, n); return; }
    const n = ctrl ? A.jumpEdge(s, s.active, dr, dc) : stepVisible(s, s.active.r, s.active.c, dr, dc);
    Grid.select(n.r, n.c);
  };
  A.moveTo = (r, c, shift) => { const s = S(); if (shift) Grid.setSel(s.active, { r, c }); else Grid.select(r, c); };
  A.tabStart = null;
  A.moveAfterEnter = (dr, dc) => {
    const s = S(); const g = s.sel;
    if (g.r0 === g.r1 && g.c0 === g.c1) {
      if (dc) { if (!A.tabStart || A.tabStart.r !== s.active.r) A.tabStart = { r: s.active.r, c: s.active.c }; A.tabExpect = null; A.moveSel(0, dc); A.tabExpect = { ...s.active }; return; }
      if (dr === 1 && A.tabStart && A.tabExpect && A.tabExpect.r === s.active.r && A.tabExpect.c === s.active.c) { const c = A.tabStart.c; A.tabStart = null; Grid.select(s.active.r + 1, c); return; }
      A.tabStart = null;
    } const single = (g.r0 === g.r1 && g.c0 === g.c1) || (s.mergeAt(s.active.r, s.active.c) && SU.intersects(s.mergeAt(s.active.r, s.active.c), g) && s.expandMerges({ r0: s.active.r, c0: s.active.c, r1: s.active.r, c1: s.active.c }).r1 === g.r1 && s.expandMerges({ r0: s.active.r, c0: s.active.c, r1: s.active.r, c1: s.active.c }).c1 === g.c1);
    if (single || Grid.isWholeCols(g) || Grid.isWholeRows(g)) { A.moveSel(dr, dc); return; }
    let { r, c } = s.active; const e = A.eff(g);
    const adv = () => {
      if (dr) { r += dr; if (r > e.r1) { r = e.r0; c++; if (c > e.c1) c = e.c0; } if (r < e.r0) { r = e.r1; c--; if (c < e.c0) c = e.c1; } }
      else { c += dc; if (c > e.c1) { c = e.c0; r++; if (r > e.r1) r = e.r0; } if (c < e.c0) { c = e.c1; r--; if (r < e.r0) r = e.r1; } }
    };
    adv(); let guard = 0; while (guard++ < 1000) { const m = s.mergeAt(r, c); if ((m && (m.r0 !== r || m.c0 !== c)) || s.isRowHidden(r)) adv(); else break; }
    s.active = { r, c }; Grid.scrollIntoView(r, c); Grid.notify();
  };
  A.currentRegion = (s, r, c) => {
    let g = { r0: r, c0: c, r1: r, c1: c };
    const has = (r, c) => r >= 0 && c >= 0 && !s.isEmpty(r, c);
    let changed = true, n = 0;
    while (changed && n++ < 200) {
      changed = false;
      for (let cc = g.c0 - 1; cc <= g.c1 + 1; cc++) { if (has(g.r0 - 1, cc)) { g.r0--; changed = true; break; } }
      for (let cc = g.c0 - 1; cc <= g.c1 + 1; cc++) { if (has(g.r1 + 1, cc)) { g.r1++; changed = true; break; } }
      for (let rr = g.r0 - 1; rr <= g.r1 + 1; rr++) { if (has(rr, g.c0 - 1)) { g.c0--; changed = true; break; } }
      for (let rr = g.r0 - 1; rr <= g.r1 + 1; rr++) { if (has(rr, g.c1 + 1)) { g.c1++; changed = true; break; } }
      g.r0 = Math.max(0, g.r0); g.c0 = Math.max(0, g.c0);
    }
    return s.expandMerges(g);
  };
  A.selectAllSmart = () => {
    const s = S(); const g = s.sel;
    if (!s.isEmpty(s.active.r, s.active.c) || (s.active.r > 0 && !s.isEmpty(s.active.r - 1, s.active.c))) {
      const reg = A.currentRegion(s, s.active.r, s.active.c);
      if (!(reg.r0 === g.r0 && reg.r1 === g.r1 && reg.c0 === g.c0 && reg.c1 === g.c1) && (reg.r1 > reg.r0 || reg.c1 > reg.c0)) { s.sel = reg; Grid.notify(); return; }
    }
    Grid.selectAll();
  };
  /** range for data operations: the selection, or the current region when a single cell is selected */
  A.dataRange = () => {
    const s = S(); const g = s.sel;
    if (g.r0 === g.r1 && g.c0 === g.c1 || (s.mergeAt(g.r0, g.c0) && s.expandMerges({ r0: g.r0, c0: g.c0, r1: g.r0, c1: g.c0 }).r1 === g.r1)) {
      const t = s.tableAt(g.r0, g.c0); if (t) return { ...t.range };
      return A.currentRegion(s, s.active.r, s.active.c);
    }
    return A.effData(g);
  };

  /* ---------- styles ---------- */
  let lastAction = null;
  A.repeatLast = () => { if (lastAction) lastAction(); };
  A.applyStyle = (patchOrFn, label = 'Format', ranges) => {
    if (guard()) return;
    const s = S(); ranges = ranges || A.sels().map(g => A.eff(g));
    const doit = (rs) => wb.run(label, () => {
      for (const g of rs) for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) {
        const cur = s.style(r, c); const patch = typeof patchOrFn === 'function' ? patchOrFn(cur, r, c, g) : patchOrFn; if (!patch) continue;
        const ns = SU.mergeStyle(cur, patch); if (JSON.stringify(ns) !== JSON.stringify(cur)) wb.setStyle(s, r, c, ns);
      }
    });
    doit(ranges);
    lastAction = () => doit(A.sels().map(g => A.eff(g)));
    App.refreshRibbon();
  };
  A.toggleStyle = (prop, val = true) => {
    const s = S(); const cur = s.style(s.active.r, s.active.c)[prop];
    A.applyStyle({ [prop]: cur ? undefined : val }, 'Format');
  };
  A.setNumberFormat = (nf) => A.applyStyle({ nf: nf === 'General' ? undefined : nf }, 'Number format');
  A.changeDecimals = (d) => {
    const s = S(); const a = s.active; const base = s.style(a.r, a.c).nf; const sample = s.value(a.r, a.c);
    const nf = SU.changeDecimals(base, d, sample);
    A.applyStyle(() => ({ nf }), 'Decimals');
  };
  A.setFontSize = (sz) => A.applyStyle({ sz: sz === 11 ? undefined : sz }, 'Font size');
  A.growFont = (d) => { const s = S(); const cur = s.style(s.active.r, s.active.c).sz || 11; const i = SU.SIZES.findIndex(x => x >= cur); const n = d > 0 ? SU.SIZES[Math.min(SU.SIZES.length - 1, (SU.SIZES[i] === cur ? i + 1 : i))] : SU.SIZES[Math.max(0, (i < 0 ? SU.SIZES.length : i) - 1)]; A.setFontSize(n); };

  /* ---------- borders ---------- */
  A.borderLine = { style: 'thin', color: '#000000' };
  A.applyBorders = (kind) => {
    if (guard()) return;
    const s = S(); const line = { ...A.borderLine }; const thick = { style: 'medium', color: line.color }, dbl = { style: 'double', color: line.color };
    const ranges = A.sels().map(g => A.eff(g));
    wb.run('Borders', () => {
      for (const g of ranges) {
        const set = (r, c, p) => { const cur = s.style(r, c); const ns = SU.mergeStyle(cur, p); wb.setStyle(s, r, c, ns); };
        for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) {
          const top = r === g.r0, bot = r === g.r1, left = c === g.c0, right = c === g.c1; const p = {};
          switch (kind) {
            case 'all': Object.assign(p, { bT: line, bB: line, bL: line, bR: line }); break;
            case 'outside': if (top) p.bT = line; if (bot) p.bB = line; if (left) p.bL = line; if (right) p.bR = line; break;
            case 'thick-outside': if (top) p.bT = thick; if (bot) p.bB = thick; if (left) p.bL = thick; if (right) p.bR = thick; break;
            case 'inside': if (!top) p.bT = line; if (!bot) p.bB = line; if (!left) p.bL = line; if (!right) p.bR = line; break;
            case 'top': if (top) p.bT = line; break;
            case 'bottom': if (bot) p.bB = line; break;
            case 'left': if (left) p.bL = line; break;
            case 'right': if (right) p.bR = line; break;
            case 'thick-bottom': if (bot) p.bB = thick; break;
            case 'double-bottom': if (bot) p.bB = dbl; break;
            case 'top-bottom': if (top) p.bT = line; if (bot) p.bB = line; break;
            case 'top-thick-bottom': if (top) p.bT = line; if (bot) p.bB = thick; break;
            case 'none': Object.assign(p, { bT: undefined, bB: undefined, bL: undefined, bR: undefined }); break;
          }
          if (Object.keys(p).length) set(r, c, p);
        }
        if (kind === 'none') { // clear facing sides of neighbours
          for (let c = g.c0; c <= g.c1; c++) { if (g.r0 > 0 && s.style(g.r0 - 1, c).bB) set(g.r0 - 1, c, { bB: undefined }); if (s.style(g.r1 + 1, c).bT) set(g.r1 + 1, c, { bT: undefined }); }
          for (let r = g.r0; r <= g.r1; r++) { if (g.c0 > 0 && s.style(r, g.c0 - 1).bR) set(r, g.c0 - 1, { bR: undefined }); if (s.style(r, g.c1 + 1).bL) set(r, g.c1 + 1, { bL: undefined }); }
        }
      }
    });
    lastAction = () => A.applyBorders(kind);
  };

  /* ---------- clear ---------- */
  A.clear = (what = 'contents') => {
    if (guard()) return;
    const s = S(); const ranges = A.sels().map(g => what === 'contents' ? A.effData(g) : A.eff(g));
    wb.run('Clear', () => {
      for (const g of ranges) {
        if (what === 'contents' || what === 'all') { const cells = []; for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) if (!s.isEmpty(r, c)) cells.push({ r, c, v: null }); wb.setContents(s, cells); }
        if (what === 'formats' || what === 'all') { for (const k of [...s.styles.keys()]) { const { r, c } = unkey(k); if (SU.inRange(g, r, c)) wb.setStyle(s, r, c, null); } }
        if (what === 'comments' || what === 'all') { const m = new Map(s.comments); let ch = false; for (const k of [...m.keys()]) { const { r, c } = unkey(k); if (SU.inRange(g, r, c)) { m.delete(k); ch = true; } } if (ch) wb.setProp(s, 'comments', m); }
        if (what === 'links' || what === 'all' || what === 'contents') { const m = new Map(s.links); let ch = false; for (const k of [...m.keys()]) { const { r, c } = unkey(k); if (SU.inRange(g, r, c)) { m.delete(k); ch = true; } } if (ch) wb.setProp(s, 'links', m); }
        if (what === 'all') { const mg = s.merges.filter(m => !(m.r0 >= g.r0 && m.r1 <= g.r1 && m.c0 >= g.c0 && m.c1 <= g.c1)); if (mg.length !== s.merges.length) wb.setProp(s, 'merges', mg); }
      }
    });
  };

  /* ---------- merge ---------- */
  A.merge = async (kind) => {
    if (guard()) return;
    const s = S(); let g = A.eff(s.sel);
    if (kind === 'unmerge') { wb.run('Unmerge', () => wb.setProp(s, 'merges', s.merges.filter(m => !SU.intersects(m, g)))); Grid.notify(); return; }
    if (Grid.isWholeCols(s.sel) || Grid.isWholeRows(s.sel)) { BTS.toast('Select a smaller range to merge.', { error: true }); return; }
    const areas = kind === 'across' ? Array.from({ length: g.r1 - g.r0 + 1 }, (_, i) => ({ r0: g.r0 + i, r1: g.r0 + i, c0: g.c0, c1: g.c1 })) : [g];
    if (areas.every(a => a.r0 === a.r1 && a.c0 === a.c1)) return;
    // existing merge exactly equals -> toggle off (Merge & Center behaviour)
    if (kind === 'center' && s.merges.some(m => m.r0 === g.r0 && m.c0 === g.c0 && m.r1 === g.r1 && m.c1 === g.c1)) { A.merge('unmerge'); A.applyStyle({ h: undefined }, 'Unmerge'); mergeLastTwo(); return; }
    let lose = false; for (const a of areas) { let n = 0; for (let r = a.r0; r <= a.r1; r++) for (let c = a.c0; c <= a.c1; c++) if (!s.isEmpty(r, c)) n++; if (n > 1 || (n === 1 && s.isEmpty(a.r0, a.c0))) lose = true; }
    if (lose && !(await BTS.confirm('Merge cells', 'Merging cells only keeps the upper-left value and discards other values.', 'Merge'))) return;
    wb.run('Merge', () => {
      let mg = s.merges.filter(m => !areas.some(a => SU.intersects(m, a)));
      const cells = [];
      for (const a of areas) {
        if (a.r0 === a.r1 && a.c0 === a.c1) continue;
        // keep first non-empty value in the top-left cell
        let first = null; for (let r = a.r0; r <= a.r1 && first === null; r++) for (let c = a.c0; c <= a.c1; c++) if (!s.isEmpty(r, c)) { first = { r, c }; break; }
        if (first && (first.r !== a.r0 || first.c !== a.c0)) { cells.push({ r: a.r0, c: a.c0, v: s.raw(first.r, first.c) }); }
        for (let r = a.r0; r <= a.r1; r++) for (let c = a.c0; c <= a.c1; c++) if ((r !== a.r0 || c !== a.c0) && !s.isEmpty(r, c)) cells.push({ r, c, v: null });
        mg.push({ ...a });
      }
      wb.setContents(s, cells);
      wb.setProp(s, 'merges', mg);
      if (kind === 'center' || kind === 'across') for (const a of areas) { const st = s.style(a.r0, a.c0); wb.setStyle(s, a.r0, a.c0, SU.mergeStyle(st, { h: 'center', v: kind === 'center' && a.r1 > a.r0 ? (st.v || 'middle') : st.v })); }
    });
    s.sel = s.expandMerges(s.sel); Grid.notify();
  };

  /* ---------- clipboard ---------- */
  A.clip = null;
  A.copy = (cut, cd) => {
    const s = S(); const sels = A.sels();
    if (sels.length > 1) { BTS.toast('This action won’t work on multiple selections.', { error: true }); return; }
    const g = A.effData(sels[0]); const full = A.eff(sels[0]);
    const range = { r0: g.r0, c0: g.c0, r1: Math.max(g.r1, Grid.isWholeCols(sels[0]) ? g.r1 : full.r1), c1: Math.max(g.c1, Grid.isWholeRows(sels[0]) ? g.c1 : full.c1) };
    const cells = [], rows = [], html = [];
    for (let r = range.r0; r <= range.r1; r++) {
      const row = [], hrow = [];
      if (s.isRowHidden(r) && !cut) { continue; }
      for (let c = range.c0; c <= range.c1; c++) {
        const raw = s.raw(r, c); const st = s.style(r, c);
        cells.push({ r: r - range.r0, c: c - range.c0, raw, st, val: s.value(r, c), note: s.comments.get(key(r, c)), link: s.links.get(key(r, c)) });
        const txt = s.displayText(r, c); row.push(txt);
        hrow.push(`<td style="${cssOfStyle(st)}">${SU.esc(txt)}</td>`);
      }
      rows.push(row); html.push('<tr>' + hrow.join('') + '</tr>');
    }
    const id = 'bts' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    A.clip = { id, sheet: s, range, cells, cut, merges: s.merges.filter(m => m.r0 >= range.r0 && m.r1 <= range.r1 && m.c0 >= range.c0 && m.c1 <= range.c1).map(m => ({ r0: m.r0 - range.r0, r1: m.r1 - range.r0, c0: m.c0 - range.c0, c1: m.c1 - range.c0 })), colW: Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => s.rawColWidth(range.c0 + i)) };
    const tsv = SU.toTSV(rows) + '\r\n';
    const htmlStr = `<meta name="generator" content="BTS Sheets"><table data-bts-clip="${id}" style="border-collapse:collapse">${html.join('')}</table>`;
    if (cd) { cd.setData('text/plain', tsv); cd.setData('text/html', htmlStr); }
    else { try { navigator.clipboard.write([new ClipboardItem({ 'text/plain': new Blob([tsv], { type: 'text/plain' }), 'text/html': new Blob([htmlStr], { type: 'text/html' }) })]).catch(() => navigator.clipboard.writeText(tsv).catch(() => { })); } catch { } }
    A.clipText = tsv;
    Grid.copyRange = { sheet: s, range }; Grid.render();
  };
  A.cutCopyCmd = (cut) => { if (!document.execCommand(cut ? 'cut' : 'copy')) A.copy(cut, null); };
  function cssOfStyle(st) {
    const a = [];
    if (st.b) a.push('font-weight:bold'); if (st.i) a.push('font-style:italic'); if (st.u) a.push('text-decoration:underline'); if (st.color) a.push('color:' + st.color); if (st.fill) a.push('background:' + st.fill);
    if (st.font) a.push('font-family:' + st.font); if (st.sz) a.push('font-size:' + st.sz + 'pt'); if (st.h) a.push('text-align:' + st.h);
    const b = (x) => x ? `${x.style === 'medium' ? 2 : x.style === 'thick' ? 3 : 1}px ${x.style === 'double' ? 'double' : x.style === 'dashed' ? 'dashed' : x.style === 'dotted' ? 'dotted' : 'solid'} ${x.color || '#000'}` : '';
    if (st.bT) a.push('border-top:' + b(st.bT)); if (st.bB) a.push('border-bottom:' + b(st.bB)); if (st.bL) a.push('border-left:' + b(st.bL)); if (st.bR) a.push('border-right:' + b(st.bR));
    return a.join(';');
  }
  A.cssOfStyle = cssOfStyle;
  A.pasteCmd = async (mode) => {
    if (mode === undefined || mode === 'all') { if (document.execCommand('paste')) return; }
    // read system clipboard; if it matches our clip use internal
    let text = null, html = null;
    try { const items = await navigator.clipboard.read(); for (const it of items) { if (it.types.includes('text/html')) html = await (await it.getType('text/html')).text(); if (it.types.includes('text/plain')) text = await (await it.getType('text/plain')).text(); } } catch { try { text = await navigator.clipboard.readText(); } catch { } }
    const m = html && /data-bts-clip="([^"]+)"/.exec(html);
    if (A.clip && (!text && !html || (m && m[1] === A.clip.id) || (text !== null && text === A.clipText))) return A.pasteInternal(mode || 'all');
    if (mode && mode !== 'all' && mode !== 'values' && mode !== 'transpose') { if (A.clip) return A.pasteInternal(mode); }
    if (html || text) A.pasteExternal(html, text, mode);
    else if (A.clip) A.pasteInternal(mode || 'all');
  };
  A.pasteFromEvent = (cd) => {
    const html = cd.getData('text/html'), text = cd.getData('text/plain');
    const m = html && /data-bts-clip="([^"]+)"/.exec(html);
    if (A.clip && ((m && m[1] === A.clip.id) || (!html && text === A.clipText))) { A.pasteInternal('all'); return; }
    // image paste
    for (const it of cd.items || []) if (it.type.startsWith('image/')) { const f = it.getAsFile(); if (f) { f.arrayBuffer().then(b => Charts.insertImageBytes(new Uint8Array(b), f.type)); return; } }
    if (html || text) A.pasteExternal(html, text, 'all');
  };
  /** opts: mode all|values|formulas|formats|transpose|nocopyborders|colwidths|valuesnf ; op: add|sub|mul|div ; skipBlanks */
  A.pasteInternal = (mode = 'all', opts = {}) => {
    if (guard()) return;
    const clip = A.clip; if (!clip) return;
    const s = S(); const transpose = mode === 'transpose' || opts.transpose;
    const H = clip.range.r1 - clip.range.r0 + 1, W = clip.range.c1 - clip.range.c0 + 1;
    const tH = transpose ? W : H, tW = transpose ? H : W;
    let tg = A.eff(s.sel);
    // tile when selection is a multiple of the clip size
    let reps = { r: 1, c: 1 };
    const sh = tg.r1 - tg.r0 + 1, sw = tg.c1 - tg.c0 + 1;
    if (!clip.cut && sh % tH === 0 && sw % tW === 0 && (sh > tH || sw > tW) && !(Grid.isWholeCols(s.sel) || Grid.isWholeRows(s.sel))) reps = { r: sh / tH, c: sw / tW };
    const dst = { r0: tg.r0, c0: tg.c0, r1: tg.r0 + tH * reps.r - 1, c1: tg.c0 + tW * reps.c - 1 };
    // merged cells in destination
    if (s.merges.some(m => SU.intersects(m, dst) && !(m.r0 >= dst.r0 && m.r1 <= dst.r1 && m.c0 >= dst.c0 && m.c1 <= dst.c1))) { BTS.toast('We can’t do that to a merged cell.', { error: true }); return; }
    if (clip.cut && mode === 'all' && !transpose) {
      // move: formulas pointing to the moved cells follow
      A.moveCells(clip.sheet, clip.range, s, dst.r0, dst.c0);
      A.clip = null; Grid.copyRange = null; s.sel = { ...dst, r1: dst.r0 + H - 1, c1: dst.c0 + W - 1 }; s.active = { r: dst.r0, c: dst.c0 }; Grid.notify(); return;
    }
    wb.run('Paste', () => {
      const cells = [];
      for (let rr = 0; rr < reps.r; rr++) for (let cr = 0; cr < reps.c; cr++) {
        for (const cl of clip.cells) {
          const r = dst.r0 + rr * tH + (transpose ? cl.c : cl.r), c = dst.c0 + cr * tW + (transpose ? cl.r : cl.c);
          const srcR = clip.range.r0 + cl.r, srcC = clip.range.c0 + cl.c;
          if (opts.skipBlanks && (cl.raw === null || cl.raw === '')) continue;
          const wantVal = mode === 'all' || mode === 'values' || mode === 'formulas' || mode === 'transpose' || mode === 'valuesnf' || mode === 'nocopyborders';
          if (wantVal) {
            let v;
            if (mode === 'values' || mode === 'valuesnf') v = SU.isErr(cl.val) ? cl.val.value : cl.val;
            else { v = cl.raw; if (typeof v === 'string' && v[0] === '=') v = SU.shiftFormula(v, r - srcR, c - srcC); }
            if (opts.op && typeof v !== 'string') { const cur = s.value(r, c); const a = typeof cur === 'number' ? cur : 0, b = typeof v === 'number' ? v : 0; v = opts.op === 'add' ? a + b : opts.op === 'sub' ? a - b : opts.op === 'mul' ? a * b : b !== 0 ? a / b : '#DIV/0!'; }
            cells.push({ r, c, v: v === undefined ? null : v });
          }
          if (mode === 'all' || mode === 'formats' || mode === 'transpose' || mode === 'nocopyborders') {
            let st = cl.st && Object.keys(cl.st).length ? { ...cl.st } : null;
            if (st && transpose) { const { bT, bB, bL, bR } = st; st.bT = bL; st.bL = bT; st.bB = bR; st.bR = bB; Object.keys(st).forEach(k => st[k] === undefined && delete st[k]); }
            if (st && mode === 'nocopyborders') { delete st.bT; delete st.bB; delete st.bL; delete st.bR; }
            wb.setStyle(s, r, c, st);
          } else if (mode === 'valuesnf') { wb.setStyle(s, r, c, SU.mergeStyle(s.style(r, c), { nf: cl.st.nf })); }
          if (mode === 'all' || mode === 'transpose') {
            if (cl.note) { const m = new Map(s.comments); m.set(key(r, c), cl.note); wb.setProp(s, 'comments', m); }
            if (cl.link) { const m = new Map(s.links); m.set(key(r, c), cl.link); wb.setProp(s, 'links', m); }
          }
        }
        if ((mode === 'all' || mode === 'formats') && !transpose && clip.merges.length) {
          const base = { r: dst.r0 + rr * tH, c: dst.c0 + cr * tW };
          const add = clip.merges.map(m => ({ r0: m.r0 + base.r, r1: m.r1 + base.r, c0: m.c0 + base.c, c1: m.c1 + base.c }));
          wb.setProp(s, 'merges', [...s.merges.filter(m => !add.some(a => SU.intersects(a, m))), ...add]);
        }
      }
      wb.setContents(s, cells);
      if (mode === 'colwidths') { const m = new Map(s.colW); clip.colW.forEach((w, i) => m.set(dst.c0 + i, w)); wb.setProp(s, 'colW', m); }
    });
    s.sel = dst; s.active = { r: dst.r0, c: dst.c0 }; Grid.notify();
    if (clip.cut) { A.clip = null; Grid.copyRange = null; }
  };
  A.moveCells = (fromSheet, src, toSheet, r, c) => {
    if (fromSheet === toSheet && src.r0 === r && src.c0 === c) return;
    const H = src.r1 - src.r0 + 1, W = src.c1 - src.c0 + 1; const dst = { r0: r, c0: c, r1: r + H - 1, c1: c + W - 1 };
    const hf = wb.hf; let fsnap, before;
    const stylesFrom = fromSheet.styles, stylesTo = toSheet.styles;
    const notesFrom = fromSheet.comments, notesTo = toSheet.comments, linksFrom = fromSheet.links, linksTo = toSheet.links;
    const moveMap = (mapF, mapT) => {
      const nf = new Map(mapF), moved = [];
      for (const [k, v] of mapF) { const p = unkey(k); if (SU.inRange(src, p.r, p.c)) { nf.delete(k); moved.push([p.r - src.r0 + r, p.c - src.c0 + c, v]); } }
      const nt = fromSheet === toSheet ? nf : new Map(mapT);
      for (const k of [...nt.keys()]) { const p = unkey(k); if (SU.inRange(dst, p.r, p.c)) nt.delete(k); }
      moved.forEach(([rr, cc, v]) => nt.set(key(rr, cc), v));
      return [nf, nt];
    };
    wb.run('Move', () => wb.custom(() => {
      fsnap = wb.snapshotFormulas();
      before = { from: wb.captureRange(fromSheet, src), to: wb.captureRange(toSheet, dst) };
      try { hf.moveCells({ start: fromSheet.a(src.r0, src.c0), end: fromSheet.a(src.r1, src.c1) }, toSheet.a(r, c)); }
      catch (e) { // fallback: copy + clear
        const cells = wb.captureRange(fromSheet, src); hf.batch(() => { cells.forEach(x => hf.setCellContents(fromSheet.a(x.r, x.c), [[null]])); cells.forEach(x => hf.setCellContents(toSheet.a(x.r - src.r0 + r, x.c - src.c0 + c), [[wb._toHFContent(x.v)]])); });
      }
      const [sf, st] = moveMap(stylesFrom, stylesTo); fromSheet.styles = sf; toSheet.styles = st;
      const [nf, nt] = moveMap(notesFrom, notesTo); fromSheet.comments = nf; toSheet.comments = nt;
      const [lf, lt] = moveMap(linksFrom, linksTo); fromSheet.links = lf; toSheet.links = lt;
    }, () => {
      hf.batch(() => { for (const [sh, g] of [[fromSheet, src], [toSheet, dst]]) for (let rr = g.r0; rr <= g.r1; rr++) for (let cc = g.c0; cc <= g.c1; cc++) hf.setCellContents(sh.a(rr, cc), [[null]]); });
      wb._applyCells([...before.from.map(x => ({ s: fromSheet, r: x.r, c: x.c, b: x.v })), ...before.to.map(x => ({ s: toSheet, r: x.r, c: x.c, b: x.v }))], 'b');
      wb.restoreFormulas(fsnap);
      fromSheet.styles = stylesFrom; toSheet.styles = stylesTo; fromSheet.comments = notesFrom; toSheet.comments = notesTo; fromSheet.links = linksFrom; toSheet.links = linksTo;
    }));
  };
  A.pasteExternal = (html, text, mode = 'all') => {
    if (guard()) return;
    const s = S(); let grid = null, styles = null;
    if (html && /<table/i.test(html)) {
      try {
        const doc = new DOMParser().parseFromString(html, 'text/html'); const tbl = doc.querySelector('table');
        if (tbl) {
          grid = []; styles = []; const occupied = {};
          [...tbl.rows].forEach((tr, ri) => {
            let ci = 0; grid[ri] = grid[ri] || []; styles[ri] = styles[ri] || [];
            for (const td of tr.cells) {
              while (occupied[ri + ',' + ci]) ci++;
              const rs = +td.rowSpan || 1, cs = +td.colSpan || 1;
              let t = td.innerText !== undefined ? td.innerText : td.textContent; t = t.replace(/ /g, ' ').replace(/\n$/, '');
              grid[ri][ci] = t;
              const st = {}; const cssS = td.getAttribute('style') || '';
              const cs2 = (p) => { const m = new RegExp('(?:^|;)\\s*' + p + '\\s*:\\s*([^;]+)', 'i').exec(cssS); return m ? m[1].trim() : ''; };
              if (/bold|[6-9]00/.test(cs2('font-weight')) || td.querySelector('b,strong')) st.b = true;
              if (/italic/.test(cs2('font-style')) || td.querySelector('i,em')) st.i = true;
              const bg = cs2('background-color') || cs2('background'); if (bg && !/transparent|none|initial/.test(bg)) st.fill = toHex(bg);
              const col = cs2('color'); if (col) { const h = toHex(col); if (h && h !== '#000000') st.color = h; }
              const al = cs2('text-align'); if (/center|right|left/.test(al)) st.h = al.match(/center|right|left/)[0];
              styles[ri][ci] = Object.keys(st).filter(k => st[k]).length ? st : null;
              for (let a = 0; a < rs; a++) for (let b = 0; b < cs; b++) occupied[(ri + a) + ',' + (ci + b)] = true;
              ci += cs;
            }
          });
        }
      } catch (e) { console.warn(e); grid = null; }
    }
    if (!grid) { if (text == null) return; text = text.replace(/\r\n$/, '').replace(/\n$/, ''); grid = text.includes('\t') || text.includes('\n') ? SU.parseDelimited(text, '\t') : [[text]]; }
    if (mode === 'transpose') { const t = []; grid.forEach((row, r) => row.forEach((v, c) => { (t[c] = t[c] || [])[r] = v; })); grid = t; if (styles) { const ts = []; styles.forEach((row, r) => row.forEach((v, c) => { (ts[c] = ts[c] || [])[r] = v; })); styles = ts; } }
    const r0 = s.sel.r0, c0 = s.sel.c0;
    wb.run('Paste', () => {
      const cells = [];
      grid.forEach((row, ri) => row && row.forEach((t, ci) => {
        if (t === undefined) return; const r = r0 + ri, c = c0 + ci;
        const res = A.interpret(s, r, c, t);
        cells.push({ r, c, v: res.error ? t : res.v });
        if (res.nf && mode !== 'values') wb.setStyle(s, r, c, SU.mergeStyle(s.style(r, c), { nf: res.nf }));
        if (styles && styles[ri] && styles[ri][ci] && mode === 'all') wb.setStyle(s, r, c, SU.mergeStyle(s.style(r, c), styles[ri][ci]));
      }));
      wb.setContents(s, cells);
    });
    const h = grid.length, w = Math.max(...grid.map(r => r ? r.length : 0));
    s.sel = { r0, c0, r1: r0 + h - 1, c1: c0 + Math.max(1, w) - 1 }; s.active = { r: r0, c: c0 }; Grid.notify();
  };
  function toHex(c) { if (!c) return null; if (c[0] === '#') return c.length === 4 ? '#' + [...c.slice(1)].map(x => x + x).join('') : c.slice(0, 7).toLowerCase(); const m = /rgba?\((\d+)\D+(\d+)\D+(\d+)(?:\D+([\d.]+))?/.exec(c); if (!m) return null; if (m[4] !== undefined && +m[4] === 0) return null; return '#' + [m[1], m[2], m[3]].map(x => (+x).toString(16).padStart(2, '0')).join(''); }

  /* ---------- format painter ---------- */
  A.painter = null;
  A.startPainter = (sticky) => {
    const s = S(); const g = A.eff(s.sel); const H = g.r1 - g.r0 + 1, W = g.c1 - g.c0 + 1;
    const pat = []; for (let r = 0; r < Math.min(H, 200); r++) { pat[r] = []; for (let c = 0; c < Math.min(W, 100); c++) pat[r][c] = s.style(g.r0 + r, g.c0 + c); }
    A.painter = { pat, H: pat.length, W: pat[0].length, sticky };
    document.getElementById('btn-painter').classList.add('on'); document.body.classList.add('painting');
    Grid.scroller.style.cursor = 'copy';
  };
  A.cancelPainter = () => { A.painter = null; const b = document.getElementById('btn-painter'); b && b.classList.remove('on'); document.body.classList.remove('painting'); };
  A.applyPainter = () => {
    const p = A.painter; if (!p) return; const s = S(); let g = A.eff(s.sel);
    if (g.r0 === g.r1 && g.c0 === g.c1) g = { r0: g.r0, c0: g.c0, r1: g.r0 + p.H - 1, c1: g.c0 + p.W - 1 };
    wb.run('Format painter', () => { for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) { const st = p.pat[(r - g.r0) % p.H][(c - g.c0) % p.W]; wb.setStyle(s, r, c, st && Object.keys(st).length ? { ...st } : null); } });
    if (!p.sticky) A.cancelPainter();
    A.painter = null; s.sel = g; Grid.notify(); if (p.sticky) A.painter = p;
  };

  /* ---------- fill (handle / Ctrl+D / Ctrl+R) ---------- */
  const lists = [SU.DAYS_FULL, SU.DAYS_SHORT, SU.MONTHS_FULL, SU.MONTHS_SHORT];
  function listOf(v) {
    if (typeof v !== 'string') return null; const lv = v.toLowerCase();
    for (const L of lists) { const i = L.findIndex(x => x.toLowerCase() === lv); if (i >= 0) { const caseFn = v === v.toUpperCase() ? (x) => x.toUpperCase() : v === v.toLowerCase() ? (x) => x.toLowerCase() : (x) => x; return { L, i, caseFn }; } }
    return null;
  }
  /** given source values (raw) along an axis produce generator f(k) for k >= n (k index from start) */
  function seriesFor(vals, styles, ctrl) {
    const n = vals.length;
    const nums = vals.map(v => typeof v === 'number' ? v : null);
    const allNum = nums.every(v => v !== null);
    const isDate = allNum && styles.some(st => st.nf && SU.isDateFormat(st.nf) && /[dy]/i.test(st.nf.replace(/"[^"]*"/g, '')));
    if (allNum && n >= 1) {
      if (n === 1) { if (isDate) return (k) => vals[0] + k; return ctrl ? (k) => vals[0] + k : null; }
      if (isDate) {
        const ds = vals.map(v => SU.serialToDate(v)); const sameDay = ds.every(d => d.getUTCDate() === ds[0].getUTCDate());
        const months = (ds[1].getUTCFullYear() - ds[0].getUTCFullYear()) * 12 + ds[1].getUTCMonth() - ds[0].getUTCMonth();
        if (sameDay && months !== 0 && ds.every((d, i) => i === 0 || (d.getUTCFullYear() - ds[i - 1].getUTCFullYear()) * 12 + d.getUTCMonth() - ds[i - 1].getUTCMonth() === months)) {
          return (k) => { const d = new Date(Date.UTC(ds[0].getUTCFullYear(), ds[0].getUTCMonth() + months * k, ds[0].getUTCDate())); return SU.jsDateToSerial(d); };
        }
      }
      // least squares linear trend (Excel's AutoFill behaviour)
      let sx = 0, sy = 0, sxx = 0, sxy = 0; nums.forEach((y, x) => { sx += x; sy += y; sxx += x * x; sxy += x * y; });
      const b = (n * sxy - sx * sy) / (n * sxx - sx * sx), a = (sy - b * sx) / n;
      return (k) => +(a + b * k).toPrecision(15);
    }
    // text with trailing number
    const tn = vals.map(v => typeof v === 'string' && v[0] !== '=' ? /^(.*?)(\d+)(\D*)$/.exec(v) : null);
    if (tn.every(m => m) && tn.every(m => m[1] === tn[0][1] && m[3] === tn[0][3])) {
      const ns = tn.map(m => +m[2]); const step = n > 1 ? (ns[n - 1] - ns[0]) / (n - 1) : 1; const width = tn[0][2].length;
      return (k) => { const v = Math.round(ns[0] + step * k); return tn[0][1] + (tn[0][2][0] === '0' ? String(v).padStart(width, '0') : String(v)) + tn[0][3]; };
    }
    const ls = vals.map(listOf);
    if (ls.every(x => x) && ls.every(x => x.L === ls[0].L)) {
      const L = ls[0].L; const step = n > 1 ? ((ls[1].i - ls[0].i + L.length) % L.length) : 1;
      return (k) => ls[0].caseFn(L[(((ls[0].i + step * k) % L.length) + L.length) % L.length]);
    }
    return null;
  }
  A.fill = (src, target, opts = {}) => {
    if (guard()) return;
    const s = S();
    if (target.shrink) { const cl = { r0: target.r1 + 1, c0: src.c0, r1: src.r1, c1: src.c1 }; wb.run('Clear', () => { const cells = []; for (let r = cl.r0; r <= cl.r1; r++) for (let c = cl.c0; c <= cl.c1; c++) cells.push({ r, c, v: null }); wb.setContents(s, cells); }); s.sel = { ...src, r1: target.r1 }; Grid.notify(); return; }
    const vertical = target.r0 !== src.r0 || target.r1 !== src.r1;
    const forward = vertical ? target.r1 > src.r1 : target.c1 > src.c1;
    wb.run('Fill', () => {
      const cells = [];
      const lines = vertical ? [src.c0, src.c1] : [src.r0, src.r1];
      for (let li = lines[0]; li <= lines[1]; li++) {
        const srcIdx = []; const n = vertical ? src.r1 - src.r0 + 1 : src.c1 - src.c0 + 1;
        for (let k = 0; k < n; k++) srcIdx.push(vertical ? { r: src.r0 + k, c: li } : { r: li, c: src.c0 + k });
        const vals = srcIdx.map(p => s.raw(p.r, p.c)); const sts = srcIdx.map(p => s.style(p.r, p.c));
        const gen = opts.copy ? null : seriesFor(forward ? vals : vals.slice().reverse(), sts, opts.ctrl);
        const tStart = vertical ? (forward ? src.r1 + 1 : target.r0) : (forward ? src.c1 + 1 : target.c0);
        const tEnd = vertical ? (forward ? target.r1 : src.r0 - 1) : (forward ? target.c1 : src.c0 - 1);
        for (let t = tStart; t <= tEnd; t++) {
          const r = vertical ? t : li, c = vertical ? li : t;
          const off = forward ? t - (vertical ? src.r0 : src.c0) : (vertical ? src.r1 : src.c1) - t; // distance from first source in fill direction
          const si = ((forward ? t - (vertical ? src.r1 + 1 : src.c1 + 1) : (vertical ? src.r0 - 1 : src.c0 - 1) - t) % n + n) % n;
          const sIdx = forward ? si : n - 1 - si;
          const sp = srcIdx[sIdx]; let v = vals[sIdx];
          if (!opts.formatsOnly) {
            if (typeof v === 'string' && v[0] === '=') v = SU.shiftFormula(v, r - sp.r, c - sp.c);
            else if (gen && v !== null) v = gen(off);
            cells.push({ r, c, v });
          }
          if (!opts.valuesOnly) wb.setStyle(s, r, c, sts[sIdx] && Object.keys(sts[sIdx]).length ? { ...sts[sIdx] } : null);
        }
      }
      wb.setContents(s, cells);
    });
    s.sel = SU.norm({ r0: Math.min(src.r0, target.r0), c0: Math.min(src.c0, target.c0), r1: Math.max(src.r1, target.r1), c1: Math.max(src.c1, target.c1) });
    Grid.notify();
  };
  A.fillDouble = () => {
    const s = S(); const src = { ...s.sel }; let len = -1;
    const probe = (c) => { if (c < 0) return -1; let r = src.r1 + 1; if (s.isEmpty(r, c)) return -1; while (!s.isEmpty(r + 1, c)) r++; return r; };
    len = Math.max(probe(src.c0 - 1), probe(src.c1 + 1));
    if (len <= src.r1) return;
    A.fill(src, { ...src, r1: len });
  };
  A.fillDir = (dir) => {
    const s = S(); const g = A.eff(s.sel);
    if (dir === 'down') { if (g.r0 === g.r1) { if (g.r0 === 0) return; A.fill({ ...g, r0: g.r0 - 1, r1: g.r0 - 1 }, { ...g, r0: g.r0 - 1 }, { copy: true }); } else A.fill({ ...g, r1: g.r0 }, g, { copy: true }); }
    if (dir === 'right') { if (g.c0 === g.c1) { if (g.c0 === 0) return; A.fill({ ...g, c0: g.c0 - 1, c1: g.c0 - 1 }, { ...g, c0: g.c0 - 1 }, { copy: true }); } else A.fill({ ...g, c1: g.c0 }, g, { copy: true }); }
    if (dir === 'up') { if (g.r0 === g.r1) return; A.fill({ ...g, r0: g.r1 }, g, { copy: true }); }
    if (dir === 'left') { if (g.c0 === g.c1) return; A.fill({ ...g, c0: g.c1 }, g, { copy: true }); }
    s.sel = g; Grid.notify();
  };
  A.fillSeries = (opts) => { // {dir:'cols'|'rows', type:'linear'|'growth'|'date', step, stop, unit}
    const s = S(); const g = A.eff(s.sel); const start = s.value(g.r0, g.c0); if (typeof start !== 'number') { BTS.toast('Select a cell with a starting number.', { error: true }); return; }
    const cells = []; const vertical = opts.dir === 'cols';
    const len = vertical ? (g.r1 > g.r0 ? g.r1 - g.r0 + 1 : null) : (g.c1 > g.c0 ? g.c1 - g.c0 + 1 : null);
    let v = start, k = 0;
    const next = () => { if (opts.type === 'growth') return v * opts.step; if (opts.type === 'date') { const d = SU.serialToDate(v); if (opts.unit === 'month') return SU.jsDateToSerial(new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + opts.step, d.getUTCDate()))); if (opts.unit === 'year') return SU.jsDateToSerial(new Date(Date.UTC(d.getUTCFullYear() + opts.step, d.getUTCMonth(), d.getUTCDate()))); if (opts.unit === 'weekday') { let x = v; for (let i = 0; i < opts.step; i++) { x++; while ([0, 6].includes(SU.serialToDate(x).getUTCDay())) x++; } return x; } return v + opts.step; } return v + opts.step; };
    const maxN = len || 10000;
    while (k < maxN) { if (opts.stop !== '' && opts.stop !== undefined && !isNaN(+opts.stop) && ((opts.step > 0 && v > +opts.stop) || (opts.step < 0 && v < +opts.stop))) break; cells.push({ r: vertical ? g.r0 + k : g.r0, c: vertical ? g.c0 : g.c0 + k, v: +v.toPrecision(15) }); v = next(); k++; if (!len && (opts.stop === '' || opts.stop === undefined)) break; }
    wb.run('Series', () => { wb.setContents(s, cells); const st = s.style(g.r0, g.c0); cells.forEach(x => wb.setStyle(s, x.r, x.c, Object.keys(st).length ? { ...st } : null)); });
  };

  /* ---------- AutoSum ---------- */
  A.autoSum = (fn = 'SUM') => {
    if (guard()) return;
    const s = S(); const g = s.sel;
    const isNum = (r, c) => typeof s.value(r, c) === 'number' && !(s.raw(r, c) + '').match(/^=.*(SUM|SUBTOTAL)\(/i);
    if (g.r0 === g.r1 && g.c0 === g.c1) {
      const { r, c } = s.active; let rg = null;
      // look up
      let rr = r - 1; while (rr >= 0 && s.isEmpty(rr, c)) { rr--; if (r - rr > 1) break; }
      if (r > 0 && isNum(r - 1, c) || (r > 1 && s.isEmpty(r - 1, c) === false && isNum(r - 1, c))) { let top = r - 1; while (top > 0 && (isNum(top - 1, c) || (s.isEmpty(top - 1, c) === false && typeof s.value(top - 1, c) === 'number'))) top--; rg = { r0: top, c0: c, r1: r - 1, c1: c }; }
      else if (c > 0 && isNum(r, c - 1)) { let left = c - 1; while (left > 0 && isNum(r, left - 1)) left--; rg = { r0: r, c0: left, r1: r, c1: c - 1 }; }
      else if (r > 0 && !s.isEmpty(r - 1, c)) { let top = r - 1; while (top > 0 && !s.isEmpty(top - 1, c) && typeof s.value(top - 1, c) !== 'string') top--; rg = { r0: top, c0: c, r1: r - 1, c1: c }; }
      const text = '=' + fn + '(' + (rg ? SU.rangeStr(rg) : '') + ')';
      Ed.start(text, 'enter');
      if (rg) { Ed.state.point = { start: fn.length + 2, end: fn.length + 2 + SU.rangeStr(rg).length, g: rg, anchor: { r: rg.r0, c: rg.c0 }, cur: { r: rg.r1, c: rg.c1 } }; }
      else { const el = document.getElementById('cell-editor'); const sel = getSelection(); const t = el.firstChild; /* caret before ) */ try { const range = document.createRange(); const last = el.lastChild; range.setStart(last, last.textContent.length - 1); range.collapse(true); sel.removeAllRanges(); sel.addRange(range); } catch { } }
      return;
    }
    // multi-cell: totals below each column (and right of rows if selection includes an empty last column)
    const e = A.effData(g); const lastRowEmpty = (() => { for (let c = e.c0; c <= e.c1; c++) if (!s.isEmpty(g.r1, c)) return false; return true; })();
    const lastColEmpty = (() => { for (let r = e.r0; r <= g.r1; r++) if (!s.isEmpty(r, g.c1)) return false; return true; })();
    wb.run('AutoSum', () => {
      const cells = [];
      if (lastColEmpty && !lastRowEmpty && g.c1 > g.c0) { for (let r = g.r0; r <= g.r1; r++) cells.push({ r, c: g.c1, v: `=${fn}(${SU.rangeStr({ r0: r, c0: g.c0, r1: r, c1: g.c1 - 1 })})` }); }
      else { const tr = lastRowEmpty ? g.r1 : g.r1 + 1; const dr1 = lastRowEmpty ? g.r1 - 1 : g.r1; for (let c = g.c0; c <= g.c1; c++) { let any = false; for (let r = g.r0; r <= dr1; r++) if (typeof s.value(r, c) === 'number') { any = true; break; } if (any || g.c0 === g.c1) cells.push({ r: tr, c, v: `=${fn}(${SU.rangeStr({ r0: g.r0, c0: c, r1: dr1, c1: c })})` }); } }
      wb.setContents(s, cells);
    });
  };

  /* ---------- sort ---------- */
  function hasHeader(s, g) {
    if (g.r1 <= g.r0) return false;
    let hdrText = 0, bodyNum = 0, n = 0;
    for (let c = g.c0; c <= g.c1; c++) { const v = s.value(g.r0, c), w = s.value(g.r0 + 1, c); if (typeof v === 'string' && v !== '') hdrText++; if (typeof w === 'number') bodyNum++; if (v !== null) n++; }
    if (s.style(g.r0, g.c0).b && !s.style(g.r0 + 1, g.c0).b) return true;
    return hdrText === n && n > 0 && (bodyNum > 0 || g.c1 === g.c0 && false);
  }
  A.hasHeader = hasHeader;
  A.sortRange = (s, g, keys, header, byRows) => {
    // keys: [{c, desc, by:'value'|'color'|'font', custom:[list]}]
    const body = { ...g, r0: g.r0 + (header ? 1 : 0) };
    if (body.r1 <= body.r0 && !byRows) return;
    const n = body.r1 - body.r0 + 1;
    const rows = [];
    for (let i = 0; i < n; i++) { const r = body.r0 + i; if (s.isRowHidden(r) && s.filterHidden.has(r)) continue; rows.push(r); }
    const order = rows.slice().sort((a, b) => {
      for (const k of keys) {
        let va = s.value(a, k.c), vb = s.value(b, k.c);
        if (SU.isErr(va)) va = { e: 1 }; if (SU.isErr(vb)) vb = { e: 1 };
        let d;
        if (k.custom) { const ia = k.custom.indexOf(SU.valueText(va)), ib = k.custom.indexOf(SU.valueText(vb)); d = (ia < 0 ? 1e9 : ia) - (ib < 0 ? 1e9 : ib); }
        else if (k.by === 'color') { const fa = s.style(a, k.c).fill || '', fb = s.style(b, k.c).fill || ''; d = (fa === k.color ? 0 : 1) - (fb === k.color ? 0 : 1); }
        else { const blank = (v) => v === null || v === '' ; if (blank(va) && !blank(vb)) return 1; if (blank(vb) && !blank(va)) return -1; d = SU.cmp(va, vb); if (k.desc) d = -d; }
        if (d) return d;
      }
      return a - b;
    });
    if (order.every((r, i) => r === rows[i])) return;
    const W = { c0: g.c0, c1: g.c1 };
    const snap = order.map(r => { const row = []; for (let c = W.c0; c <= W.c1; c++) row.push({ v: s.raw(r, c), st: s.style(r, c), note: s.comments.get(key(r, c)), link: s.links.get(key(r, c)) }); return { r, row, h: s.rowH.get(r) }; });
    wb.run('Sort', () => {
      const cells = []; const notes = new Map(s.comments), links = new Map(s.links);
      snap.forEach((src, i) => {
        const r = rows[i];
        src.row.forEach((cell, j) => {
          const c = W.c0 + j; let v = cell.v;
          if (typeof v === 'string' && v[0] === '=') v = shiftRowRelativeWithin(v, r - src.r, g);
          cells.push({ r, c, v }); wb.setStyle(s, r, c, Object.keys(cell.st).length ? cell.st : null);
          if (cell.note) notes.set(key(r, c), cell.note); else notes.delete(key(r, c));
          if (cell.link) links.set(key(r, c), cell.link); else links.delete(key(r, c));
        });
      });
      wb.setContents(s, cells); wb.setProp(s, 'comments', notes); wb.setProp(s, 'links', links);
    });
  };
  function shiftRowRelativeWithin(f, dr) { return SU.mapRefs(f, (ref) => { if (ref.sheet) return undefined; if (ref.kind === 'cell') { if (!ref.a.ar) ref.a.r += dr; if (ref.b && !ref.b.ar) ref.b.r += dr; if (ref.a.r < 0 || (ref.b && ref.b.r < 0)) return '#REF!'; return SU.refToString(ref); } return undefined; }); }
  A.quickSort = (desc) => {
    if (guard()) return;
    const s = S(); let g = A.dataRange(); let header;
    const t = s.tableAt(s.active.r, s.active.c); const fl = s.filter && SU.inRange(s.filter.range, s.active.r, s.active.c) ? s.filter : null;
    if (t) { g = t.range; header = t.header !== false; if (t.total) g = { ...g, r1: g.r1 - 1 }; }
    else if (fl) { g = fl.range; header = true; }
    else header = hasHeader(s, g) && (s.sel.r0 === s.sel.r1);
    if (s.sel.r1 > s.sel.r0 && !t && !fl) header = false;
    A.sortRange(s, g, [{ c: s.active.c, desc }], header);
    if (fl || t) A.applyFilter(s);
  };

  /* ---------- filter ---------- */
  A.toggleFilter = () => {
    const s = S();
    if (s.filter) { wb.run('Filter', () => { wb.setProp(s, 'filter', null); wb.setProp(s, 'filterHidden', new Set()); }); Grid.updateSpacer(); return; }
    const t = s.tableAt(s.active.r, s.active.c);
    let g = t ? t.range : A.dataRange();
    if (g.r1 === g.r0 && g.c0 === g.c1 && s.isEmpty(g.r0, g.c0)) { BTS.toast('Select a cell in your data first.', { error: true }); return; }
    if (!t && (s.sel.r1 > s.sel.r0 || s.sel.c1 > s.sel.c0) && !(Grid.isWholeCols(s.sel))) g = A.effData(s.sel);
    if (Grid.isWholeCols(s.sel)) { const u = s.usedRange(); g = { r0: 0, c0: s.sel.c0, r1: Math.max(u.r1, 0), c1: s.sel.c1 }; }
    wb.run('Filter', () => wb.setProp(s, 'filter', { range: g, cols: {}, tableId: t ? t.id : undefined }));
  };
  A.applyFilter = (s = S()) => {
    const f = s.filter; const hidden = new Set();
    if (f) {
      for (let r = f.range.r0 + 1; r <= f.range.r1; r++) {
        for (const ck of Object.keys(f.cols)) {
          const cond = f.cols[ck]; const c = +ck; const v = s.value(r, c); const t = SU.valueText(v) === '' ? '' : s.displayText(r, c);
          if (cond.values && !cond.values.includes(t)) { hidden.add(r); break; }
          if (cond.cond && !A.matchCond(cond.cond, v, t, s, f, c)) { hidden.add(r); break; }
        }
      }
    }
    wb.setProp(s, 'filterHidden', hidden); Grid.updateSpacer();
  };
  A.matchCond = (cd, v, t, s, f, c) => {
    const lt = t.toLowerCase(), q = String(cd.v1 ?? '').toLowerCase(); const num = typeof v === 'number' ? v : null;
    switch (cd.op) {
      case 'eq': return num !== null && cd.v1 !== '' && !isNaN(+cd.v1) ? num === +cd.v1 : lt === q;
      case 'ne': return num !== null && !isNaN(+cd.v1) ? num !== +cd.v1 : lt !== q;
      case 'contains': return lt.includes(q); case 'notcontains': return !lt.includes(q);
      case 'begins': return lt.startsWith(q); case 'ends': return lt.endsWith(q);
      case 'gt': return num !== null && num > +cd.v1; case 'lt': return num !== null && num < +cd.v1;
      case 'ge': return num !== null && num >= +cd.v1; case 'le': return num !== null && num <= +cd.v1;
      case 'between': return num !== null && num >= +cd.v1 && num <= +cd.v2;
      case 'top': case 'bottom': case 'above': case 'below': {
        const vals = []; for (let r = f.range.r0 + 1; r <= f.range.r1; r++) { const x = s.value(r, c); if (typeof x === 'number') vals.push(x); }
        if (num === null || !vals.length) return false; vals.sort((a, b) => b - a);
        if (cd.op === 'top') return num >= vals[Math.min(vals.length, +cd.v1 || 10) - 1];
        if (cd.op === 'bottom') return num <= vals[Math.max(0, vals.length - (+cd.v1 || 10))];
        const avg = vals.reduce((a, b) => a + b, 0) / vals.length; return cd.op === 'above' ? num > avg : num < avg;
      }
    }
    return true;
  };
  A.setFilterCol = (c, cond) => {
    const s = S(); const f = s.filter; if (!f) return;
    const cols = { ...f.cols }; if (cond) cols[c] = cond; else delete cols[c];
    wb.run('Filter', () => { wb.setProp(s, 'filter', { ...f, cols }); A.applyFilter(s); });
  };
  A.clearFilters = () => { const s = S(); if (!s.filter) return; wb.run('Clear filter', () => { wb.setProp(s, 'filter', { ...s.filter, cols: {} }); A.applyFilter(s); }); };
  A.reapplyFilter = () => { const s = S(); if (!s.filter) return; wb.run('Reapply filter', () => A.applyFilter(s)); };

  /* ---------- structure ---------- */
  A.insertRows = (above = true) => {
    if (guard()) return; const s = S(); const g = s.sel; const n = Grid.isWholeRows(g) || g.r1 - g.r0 < 1000 ? g.r1 - g.r0 + 1 : 1;
    wb.run('Insert rows', () => wb.insertRows(s, above ? g.r0 : g.r1 + 1, Math.min(n, 5000)));
    Grid.updateSpacer(); Grid.notify();
  };
  A.insertCols = (left = true) => {
    if (guard()) return; const s = S(); const g = s.sel; const n = Grid.isWholeCols(g) || g.c1 - g.c0 < 200 ? Math.min(g.c1 - g.c0 + 1, 200) : 1;
    wb.run('Insert columns', () => wb.insertCols(s, left ? g.c0 : g.c1 + 1, n));
    Grid.updateSpacer(); Grid.notify();
  };
  A.deleteRows = () => {
    if (guard()) return; const s = S(); const ranges = A.sels().map(g => ({ a: g.r0, b: Math.min(g.r1, Math.max(s.rowCount, s.usedRangeAll().r1 + 1)) })).sort((x, y) => y.a - x.a);
    wb.run('Delete rows', () => ranges.forEach(x => wb.deleteRows(s, x.a, x.b - x.a + 1)));
    s.sel = { r0: ranges[ranges.length - 1].a, r1: ranges[ranges.length - 1].a, c0: s.sel.c0, c1: s.sel.c1 }; if (Grid.isWholeRows(s.sel)) {} s.active = { r: s.sel.r0, c: s.active.c };
    if (!Grid.isWholeRows(s.sel)) s.sel = { r0: s.active.r, c0: s.active.c, r1: s.active.r, c1: s.active.c };
    Grid.updateSpacer(); Grid.notify();
  };
  A.deleteCols = () => {
    if (guard()) return; const s = S(); const ranges = A.sels().map(g => ({ a: g.c0, b: Math.min(g.c1, Math.max(s.colCount, s.usedRangeAll().c1 + 1)) })).sort((x, y) => y.a - x.a);
    wb.run('Delete columns', () => ranges.forEach(x => wb.deleteCols(s, x.a, x.b - x.a + 1)));
    s.active = { r: s.active.r, c: ranges[ranges.length - 1].a }; s.sel = { r0: s.active.r, c0: s.active.c, r1: s.active.r, c1: s.active.c };
    Grid.updateSpacer(); Grid.notify();
  };
  A.insertCells = (dir) => { if (guard()) return; const s = S(); const g = A.eff(s.sel); wb.run('Insert cells', () => wb.shiftCells(s, g, dir)); Grid.notify(); };
  A.deleteCells = (dir) => { if (guard()) return; const s = S(); const g = A.eff(s.sel); wb.run('Delete cells', () => wb.shiftCells(s, g, dir)); Grid.notify(); };
  A.setHidden = (axis, hide) => {
    if (guard()) return; const s = S(); const prop = axis === 'r' ? 'hiddenR' : 'hiddenC';
    const set = new Set(s[prop]);
    for (const g of A.sels()) { const a = axis === 'r' ? g.r0 : g.c0, b = axis === 'r' ? Math.min(g.r1, s.rowCount - 1) : Math.min(g.c1, s.colCount - 1); for (let i = (hide ? a : Math.max(0, a - (a === b ? 1 : 0))); i <= (hide ? b : b + (a === b ? 1 : 0)); i++) hide ? set.add(i) : set.delete(i); if (!hide && a === 0) for (let i = 0; i <= b; i++) set.delete(i); }
    if (hide && axis === 'r' && set.size >= s.rowCount) return;
    wb.run(hide ? 'Hide' : 'Unhide', () => wb.setProp(s, prop, set));
    if (hide) { const a = s.active; if (axis === 'r' && set.has(a.r)) { let r = a.r; while (set.has(r)) r++; Grid.select(r, a.c); } if (axis === 'c' && set.has(a.c)) { let c = a.c; while (set.has(c)) c++; Grid.select(a.r, c); } }
    Grid.updateSpacer(); Grid.notify();
  };
  A.setSize = (axis, idx, size, label) => {
    const s = S(); const prop = axis === 'col' ? 'colW' : 'rowH';
    const g = A.sels().find(x => axis === 'col' ? (Grid.isWholeCols(x) && idx >= x.c0 && idx <= x.c1) : (Grid.isWholeRows(x) && idx >= x.r0 && idx <= x.r1));
    const list = g ? (axis === 'col' ? range(g.c0, g.c1) : range(g.r0, Math.min(g.r1, s.rowCount - 1))) : [idx];
    const m = new Map(s[prop]); const hid = new Set(axis === 'col' ? s.hiddenC : s.hiddenR);
    list.forEach(i => { if (size <= 0) hid.add(i); else { m.set(i, size); hid.delete(i); } });
    wb.run(label || 'Resize', () => { wb.setProp(s, prop, m); wb.setProp(s, axis === 'col' ? 'hiddenC' : 'hiddenR', hid); });
    Grid.updateSpacer(); Ed.reposition();
  };
  const range = (a, b) => { const o = []; for (let i = a; i <= b; i++) o.push(i); return o; };
  A.autofit = (axis, idx) => {
    const s = S(); const g = A.sels().find(x => axis === 'col' ? (Grid.isWholeCols(x) && idx >= x.c0 && idx <= x.c1) : (Grid.isWholeRows(x) && idx >= x.r0 && idx <= x.r1));
    const list = idx === undefined ? (axis === 'col' ? range(A.eff(s.sel).c0, A.eff(s.sel).c1) : range(A.eff(s.sel).r0, A.eff(s.sel).r1)) : g ? (axis === 'col' ? range(g.c0, Math.min(g.c1, s.colCount - 1)) : range(g.r0, Math.min(g.r1, s.rowCount - 1))) : [idx];
    const prop = axis === 'col' ? 'colW' : 'rowH'; const m = new Map(s[prop]); const u = s.usedRange();
    for (const i of list) {
      if (axis === 'col') {
        let w = 0; for (let r = 0; r <= u.r1; r++) { if (s.isEmpty(r, i) || s.isRowHidden(r)) continue; const mg = s.mergeAt(r, i); if (mg && (mg.c0 !== mg.c1)) continue; const st = s.style(r, i); if (st.wrap) continue; const t = G_text(s, r, i); w = Math.max(w, Grid.measure(Grid.fontStr(st, 1), t) + 10 + (st.indent || 0) * 9 + (s.filter && r === s.filter.range.r0 ? 16 : 0)); }
        if (w) m.set(i, Math.min(800, Math.ceil(w))); else m.delete(i);
      } else {
        let h = 0; for (let c = 0; c <= Math.max(u.c1, 0); c++) { const st = s.style(i, c); const fs = (st.sz || 11) * 96 / 72; let lines = 1; if (!s.isEmpty(i, c)) { const t = G_text(s, i, c); if (st.wrap) lines = Grid.wrapLines(Grid.fontStr(st, 1), t, Math.max(4, s.colWidth(c) - 6)).length; else lines = 1; } h = Math.max(h, Math.ceil(lines * fs * 1.22 + 5)); }
        h = Math.max(h, SU.DEF_H); if (h === SU.DEF_H) m.delete(i); else m.set(i, h);
      }
    }
    wb.run('AutoFit', () => wb.setProp(s, prop, m)); Grid.updateSpacer();
  };
  function G_text(s, r, c) { return Grid.showFormulas && s.isFormula(r, c) ? s.raw(r, c) : s.displayText(r, c); }
  A.autofitRowsFor = (s, rows) => { /* used after wrap toggle */ };

  /* ---------- freeze ---------- */
  A.freeze = (kind) => {
    const s = S(); let fz;
    if (kind === 'none') fz = { r: 0, c: 0 };
    else if (kind === 'row') fz = { r: Grid.firstVisibleRow() + 1, c: 0 };
    else if (kind === 'col') fz = { r: 0, c: Grid.firstVisibleCol() + 1 };
    else { const a = s.active; fz = { r: a.r, c: a.c }; if (Grid.isWholeRows(s.sel)) fz = { r: s.sel.r0, c: 0 }; if (Grid.isWholeCols(s.sel)) fz = { r: 0, c: s.sel.c0 }; if (s.freeze.r || s.freeze.c) fz = { r: 0, c: 0 }; }
    if (kind === 'row') fz.r = 1, Grid.scroller.scrollTop = 0; if (kind === 'col') fz.c = 1, Grid.scroller.scrollLeft = 0;
    wb.run('Freeze panes', () => wb.setProp(s, 'freeze', fz));
    if (kind === 'cell' && (fz.r || fz.c)) { Grid.scroller.scrollTop = 0; Grid.scroller.scrollLeft = 0; }
    Grid.renderNow(); App.refreshRibbon();
  };

  /* ---------- conditional formatting & tables & styles ---------- */
  A.CF_FORMATS = [
    { label: 'Light red fill with dark red text', fmt: { fill: '#FFC7CE', color: '#9C0006' } },
    { label: 'Yellow fill with dark yellow text', fmt: { fill: '#FFEB9C', color: '#9C5700' } },
    { label: 'Green fill with dark green text', fmt: { fill: '#C6EFCE', color: '#006100' } },
    { label: 'Light red fill', fmt: { fill: '#FFC7CE' } },
    { label: 'Red text', fmt: { color: '#C00000' } },
    { label: 'Bold blue text', fmt: { color: '#1F4E9E', b: true } },
  ];
  A.addCF = (rule) => {
    const s = S(); const g = A.eff(s.sel);
    rule = { id: 'cf' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5), range: g, ...rule };
    wb.run('Conditional formatting', () => wb.setProp(s, 'cf', [...s.cf, rule]));
  };
  A.clearCF = (scope) => { const s = S(); const g = A.eff(s.sel); wb.run('Clear rules', () => wb.setProp(s, 'cf', scope === 'sheet' ? [] : s.cf.filter(r => !SU.intersects(r.range, g)))); };
  A.TABLE_THEMES = [
    { name: 'Teal', head: '#0F8A6A', band: '#E3F3EE', band2: null, line: '#9CD3C3' },
    { name: 'Blue', head: '#2F5597', band: '#DCE6F4', band2: null },
    { name: 'Orange', head: '#C55A11', band: '#FBE4D5', band2: null },
    { name: 'Grey', head: '#404040', band: '#EDEDED', band2: null },
    { name: 'Gold', head: '#BF8F00', band: '#FFF2CC', band2: null },
    { name: 'Green', head: '#548235', band: '#E2EFDA', band2: null },
    { name: 'Purple', head: '#7030A0', band: '#EADCF4', band2: null },
    { name: 'Red', head: '#C00000', band: '#F8D7D7', band2: null },
    { name: 'Navy', head: '#1F3864', band: '#D9E1F2', band2: null },
    { name: 'Light teal', head: '#7CC4AF', headText: '#06392c', band: '#F1F9F6', band2: null },
    { name: 'Light blue', head: '#9DC3E6', headText: '#10305a', band: '#EEF4FB', band2: null },
    { name: 'Light grey', head: '#D0D0D0', headText: '#222', band: '#F5F5F5', band2: null },
  ];
  A.formatAsTable = async (theme) => {
    if (guard()) return;
    const s = S(); let g = A.dataRange();
    if (s.tables.some(t => SU.intersects(t.range, g))) { const t = s.tables.find(t => SU.intersects(t.range, g)); wb.run('Table style', () => wb.setProp(s, 'tables', s.tables.map(x => x === t ? { ...x, theme } : x))); return; }
    if (s.merges.some(m => SU.intersects(m, g))) { BTS.toast('Tables can’t overlap merged cells.', { error: true }); return; }
    const ok = await BTS.dialog({ title: 'Create table', width: 380, body: `<div class="field"><label>Where is the data for your table?</label><input id="tb-range" value="${SU.absRangeStr(g)}"></div><label class="check"><input type="checkbox" id="tb-head" ${hasHeader(s, g) || typeof s.value(g.r0, g.c0) === 'string' ? 'checked' : ''}> My table has headers</label>`, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ range: r.querySelector('#tb-range').value, head: r.querySelector('#tb-head').checked }) }] });
    if (!ok) return; const rg = SU.parseRange(ok.range); if (!rg) { BTS.toast('That reference isn’t valid.', { error: true }); return; } g = rg;
    let n = 1; while (s.wb.sheets.some(sh => sh.tables.some(t => t.name === 'Table' + n))) n++;
    wb.run('Format as table', () => {
      if (!ok.head) { wb.insertRows(s, g.r0, 1); const cells = []; for (let c = g.c0; c <= g.c1; c++) cells.push({ r: g.r0, c, v: 'Column' + (c - g.c0 + 1) }); wb.setContents(s, cells); g = { ...g, r1: g.r1 + 1 }; }
      const t = { id: 't' + Date.now().toString(36), name: 'Table' + n, range: g, theme, header: true, banded: true, total: false, filterBtn: true };
      wb.setProp(s, 'tables', [...s.tables, t]);
      if (!s.filter) wb.setProp(s, 'filter', { range: g, cols: {}, tableId: t.id });
      // clear explicit fills so the banding shows
      for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) { const st = s.style(r, c); if (st.fill) wb.setStyle(s, r, c, SU.mergeStyle(st, { fill: undefined })); }
    });
    A.autofit('col', undefined);
  };
  A.convertTableToRange = (t) => { const s = S(); wb.run('Convert to range', () => { // bake styles
    const g = t.range; for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) { const i = r - g.r0 - 1; const p = r === g.r0 ? { fill: t.theme.head, color: t.theme.headText || '#FFFFFF', b: true } : (i % 2 === 0 ? { fill: t.theme.band } : {}); if (Object.keys(p).length) wb.setStyle(s, r, c, SU.mergeStyle(s.style(r, c), p)); }
    wb.setProp(s, 'tables', s.tables.filter(x => x !== t)); if (s.filter && s.filter.tableId === t.id) { wb.setProp(s, 'filter', null); wb.setProp(s, 'filterHidden', new Set()); } }); };
  A.toggleTotalRow = (t) => {
    const s = S(); const g = t.range;
    wb.run('Total row', () => {
      if (!t.total) {
        const r = g.r1 + 1; const cells = [{ r, c: g.c0, v: 'Total' }];
        for (let c = g.c0 + 1; c <= g.c1; c++) { let num = false; for (let rr = g.r0 + 1; rr <= g.r1; rr++) if (typeof s.value(rr, c) === 'number') { num = true; break; } if (num) cells.push({ r, c, v: `=SUBTOTAL(109,${SU.rangeStr({ r0: g.r0 + 1, c0: c, r1: g.r1, c1: c })})` }); }
        wb.setContents(s, cells);
        wb.setProp(s, 'tables', s.tables.map(x => x === t ? { ...x, total: true, range: { ...g, r1: r } } : x));
      } else {
        const cells = []; for (let c = g.c0; c <= g.c1; c++) cells.push({ r: g.r1, c, v: null }); wb.setContents(s, cells);
        wb.setProp(s, 'tables', s.tables.map(x => x === t ? { ...x, total: false, range: { ...g, r1: g.r1 - 1 } } : x));
      }
    });
  };
  A.CELL_STYLES = [
    { group: 'Good, Bad and Neutral', items: [['Normal', null], ['Bad', { fill: '#FFC7CE', color: '#9C0006' }], ['Good', { fill: '#C6EFCE', color: '#006100' }], ['Neutral', { fill: '#FFEB9C', color: '#9C5700' }]] },
    { group: 'Data and Model', items: [['Calculation', { fill: '#F2F2F2', color: '#FA7D00', b: true, bT: { style: 'thin', color: '#7F7F7F' }, bB: { style: 'thin', color: '#7F7F7F' }, bL: { style: 'thin', color: '#7F7F7F' }, bR: { style: 'thin', color: '#7F7F7F' } }], ['Check Cell', { fill: '#A5A5A5', color: '#FFFFFF', b: true, bT: { style: 'double', color: '#3F3F3F' }, bB: { style: 'double', color: '#3F3F3F' }, bL: { style: 'double', color: '#3F3F3F' }, bR: { style: 'double', color: '#3F3F3F' } }], ['Input', { fill: '#FFCC99', color: '#3F3F76', bT: { style: 'thin', color: '#7F7F7F' }, bB: { style: 'thin', color: '#7F7F7F' }, bL: { style: 'thin', color: '#7F7F7F' }, bR: { style: 'thin', color: '#7F7F7F' } }], ['Output', { fill: '#F2F2F2', color: '#3F3F3F', b: true, bT: { style: 'thin', color: '#3F3F3F' }, bB: { style: 'thin', color: '#3F3F3F' }, bL: { style: 'thin', color: '#3F3F3F' }, bR: { style: 'thin', color: '#3F3F3F' } }], ['Note', { fill: '#FFFFCC', bT: { style: 'thin', color: '#B2B2B2' }, bB: { style: 'thin', color: '#B2B2B2' }, bL: { style: 'thin', color: '#B2B2B2' }, bR: { style: 'thin', color: '#B2B2B2' } }], ['Warning Text', { color: '#FF0000' }], ['Explanatory', { color: '#7F7F7F', i: true }], ['Linked Cell', { color: '#FA7D00', bB: { style: 'double', color: '#FF8001' } }]] },
    { group: 'Titles and Headings', items: [['Title', { sz: 18, color: '#0B4F3E', font: 'Cambria' }], ['Heading 1', { sz: 15, b: true, color: '#0B4F3E', bB: { style: 'thick', color: '#0F8A6A' } }], ['Heading 2', { sz: 13, b: true, color: '#0B4F3E', bB: { style: 'thick', color: '#87C5B4' } }], ['Heading 3', { b: true, color: '#0B4F3E', bB: { style: 'medium', color: '#A9D6C9' } }], ['Heading 4', { b: true, color: '#0B4F3E' }], ['Total', { b: true, bT: { style: 'thin', color: '#0F8A6A' }, bB: { style: 'double', color: '#0F8A6A' } }]] },
    { group: 'Themed Cell Styles', items: [['Accent 1', { fill: '#0F8A6A', color: '#FFFFFF' }], ['Accent 2', { fill: '#2F5597', color: '#FFFFFF' }], ['Accent 3', { fill: '#C55A11', color: '#FFFFFF' }], ['Accent 4', { fill: '#7030A0', color: '#FFFFFF' }], ['20% Accent 1', { fill: '#CFE8E1' }], ['20% Accent 2', { fill: '#D6DCE4' }], ['20% Accent 3', { fill: '#FBE4D5' }], ['20% Accent 4', { fill: '#E4D4F0' }]] },
  ];
  A.applyCellStyle = (st) => {
    const keep = ['nf', 'h', 'v', 'wrap', 'indent', 'locked'];
    A.applyStyle((cur) => { const o = {}; Object.keys(cur).forEach(k => { if (!keep.includes(k)) o[k] = undefined; }); return Object.assign(o, st || {}); }, 'Cell style');
  };

  /* ---------- misc ---------- */
  A.insertNow = (time) => { const d = new Date(); const v = time ? (d.getHours() * 3600 + d.getMinutes() * 60) / 86400 : SU.dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); const s = S(); wb.run('Insert date', () => { wb.setContents(s, [{ r: s.active.r, c: s.active.c, v }]); wb.setStyle(s, s.active.r, s.active.c, SU.mergeStyle(s.style(s.active.r, s.active.c), { nf: time ? 'h:mm AM/PM' : 'dd/mm/yyyy' })); }); };
  A.openDropdownAtActive = () => App.openValidationList && App.openValidationList();
  A.removeDuplicates = (g, cols, header) => {
    const s = S(); const seen = new Set(); const keepRows = [], dropRows = [];
    for (let r = g.r0 + (header ? 1 : 0); r <= g.r1; r++) { const k = cols.map(c => SU.valueText(s.value(r, c)).toLowerCase()).join('\u0001'); if (seen.has(k)) dropRows.push(r); else { seen.add(k); keepRows.push(r); } }
    if (!dropRows.length) return { removed: 0, kept: keepRows.length };
    const snap = keepRows.map(r => { const row = []; for (let c = g.c0; c <= g.c1; c++) row.push({ v: s.raw(r, c), st: s.style(r, c) }); return { r, row }; });
    wb.run('Remove duplicates', () => {
      const cells = []; let r = g.r0 + (header ? 1 : 0);
      for (const src of snap) { src.row.forEach((x, j) => { let v = x.v; if (typeof v === 'string' && v[0] === '=') v = shiftRowRelativeWithin(v, r - src.r); cells.push({ r, c: g.c0 + j, v }); wb.setStyle(s, r, g.c0 + j, Object.keys(x.st).length ? x.st : null); }); r++; }
      for (; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) { cells.push({ r, c, v: null }); }
      wb.setContents(s, cells);
    });
    return { removed: dropRows.length, kept: keepRows.length };
  };
  A.textToColumns = (g, opts) => {
    const s = S(); const c = g.c0; const rows = [];
    for (let r = g.r0; r <= g.r1; r++) { const v = s.value(r, c); const t = SU.valueText(v); rows.push(splitText(t, opts)); }
    const maxW = Math.max(...rows.map(x => x.length));
    wb.run('Text to columns', () => {
      const cells = [];
      rows.forEach((parts, i) => parts.forEach((p, j) => { const r = g.r0 + i, cc = (opts.dest !== undefined ? opts.dest.c : c) + j; const rr = (opts.dest !== undefined ? opts.dest.r : g.r0) + i; const res = A.interpret(s, rr, cc, p); cells.push({ r: rr, c: cc, v: res.error ? p : res.v }); if (res.nf) wb.setStyle(s, rr, cc, SU.mergeStyle(s.style(rr, cc), { nf: res.nf })); }));
      wb.setContents(s, cells);
    });
    return maxW;
  };
  function splitText(t, o) {
    if (o.fixed) { const out = []; let p = 0; for (const w of o.widths) { out.push(t.slice(p, w)); p = w; } out.push(t.slice(p)); return out.map(x => x.trim()); }
    const d = []; if (o.tab) d.push('\t'); if (o.comma) d.push(','); if (o.semi) d.push(';'); if (o.space) d.push(' '); if (o.other) d.push(o.other);
    if (!d.length) return [t];
    const out = []; let cur = '', q = false;
    for (let i = 0; i < t.length; i++) { const ch = t[i]; if (o.quote && ch === '"') { q = !q; continue; } if (!q && d.includes(ch)) { out.push(cur); cur = ''; if (o.consecutive) while (i + 1 < t.length && d.includes(t[i + 1])) i++; continue; } cur += ch; }
    out.push(cur); return out;
  }
  A.splitText = splitText;
  A.flashFill = () => {
    const s = S(); const c = s.active.c; let r0 = s.active.r;
    // find examples above in this column (contiguous non-empty) and source columns left
    let top = r0; while (top > 0 && !s.isEmpty(top - 1, c)) top--;
    if (s.isEmpty(top, c)) { if (top > 0 && !s.isEmpty(top - 1, c)) top--; }
    while (top > 0 && !s.isEmpty(top - 1, c)) top--;
    const examples = []; let r = top; while (!s.isEmpty(r, c)) { examples.push({ r, t: SU.valueText(s.value(r, c)) }); r++; }
    // skip header row if it's text and the next row example doesn't come from it
    const srcCols = []; for (let cc = c - 1; cc >= 0 && srcCols.length < 4; cc--) { if (!s.isEmpty(examples.length ? examples[examples.length - 1].r : r0, cc)) srcCols.push(cc); else if (srcCols.length) break; }
    for (let cc = c + 1; srcCols.length < 4 && cc < c + 4; cc++) if (!s.isEmpty(examples.length ? examples[examples.length - 1].r : r0, cc)) srcCols.push(cc);
    if (!examples.length || !srcCols.length) { BTS.toast('Type an example of the result in the first cell, then try Flash Fill again.', { error: true, ms: 3500 }); return; }
    const cands = [];
    const cases = [(x) => x, (x) => x.toUpperCase(), (x) => x.toLowerCase(), (x) => x.replace(/\w\S*/g, w => w[0].toUpperCase() + w.slice(1).toLowerCase())];
    const tokens = (x) => x.split(/[\s,@.\-_/]+/).filter(Boolean);
    for (const sc of srcCols) {
      for (const cf of cases) {
        cands.push((row) => cf(row[sc]));
        for (let i = 0; i < 4; i++) { cands.push((row) => { const t = tokens(row[sc]); return t[i] !== undefined ? cf(t[i]) : null; }); cands.push((row) => { const t = tokens(row[sc]); return t.length > i ? cf(t[t.length - 1 - i]) : null; }); cands.push((row) => { const t = tokens(row[sc]); return t[i] !== undefined ? cf(t[i][0]) : null; }); }
      }
    }
    // two-part combos from different/same columns with separator from example
    const ex0 = examples.find(e => e.t !== '');
    const sepMatch = ex0 && /([ ,.\-_@/]+)/.exec(ex0.t);
    if (sepMatch) { const sep = sepMatch[1]; const singles = cands.slice(); for (let i = 0; i < Math.min(singles.length, 120); i++) for (let j = 0; j < Math.min(singles.length, 120); j++) if (i !== j) cands.push((row) => { const a = singles[i](row), b = singles[j](row); return a && b ? a + sep + b : null; }); }
    const rowOf = (r) => { const o = {}; for (const sc of srcCols) o[sc] = SU.valueText(s.value(r, sc)); return o; };
    let headerSkip = 0;
    const fits = (fn, exs) => exs.every(e => { try { return fn(rowOf(e.r)) === e.t; } catch { return false; } });
    let best = cands.find(fn => fits(fn, examples));
    if (!best && examples.length > 1) { best = cands.find(fn => fits(fn, examples.slice(1))); if (best) headerSkip = 1; }
    if (!best) { BTS.toast('Flash Fill couldn’t find a pattern. Add another example.', { error: true }); return; }
    const start = examples[examples.length - 1].r + 1; const cells = [];
    for (let rr = start; rr < start + 100000; rr++) { if (srcCols.every(sc => s.isEmpty(rr, sc))) break; if (!s.isEmpty(rr, c)) break; const v = best(rowOf(rr)); cells.push({ r: rr, c, v: v === null ? null : v }); }
    if (!cells.length) return;
    wb.run('Flash Fill', () => wb.setContents(s, cells));
    s.sel = { r0: start, c0: c, r1: start + cells.length - 1, c1: c }; s.active = { r: start, c }; Grid.notify();
    BTS.toast(`Flash Fill changed ${cells.length} cells`);
  };
  A.selectionStats = () => {
    const s = S(); const sels = A.sels();
    let sum = 0, count = 0, ncount = 0, min = Infinity, max = -Infinity, cells = 0;
    for (const g0 of sels) { const g = A.effData(g0); for (let r = g.r0; r <= g.r1; r++) { if (s.isRowHidden(r)) continue; for (let c = g.c0; c <= g.c1; c++) { cells++; if (cells > 400000) break; const v = s.value(r, c); if (v === null || v === '' || v === undefined) continue; count++; if (typeof v === 'number') { ncount++; sum += v; if (v < min) min = v; if (v > max) max = v; } } } }
    return { sum, count, ncount, min, max, avg: ncount ? sum / ncount : 0 };
  };
})();
