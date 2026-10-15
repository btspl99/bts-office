/* BTS Sheets — data tools: Consolidate, Group/Ungroup (outline), Circle invalid data */
(function () {
  const DT = window.DT = {};
  let wb;
  DT.init = (w) => { wb = w; };
  const S = () => wb.active;
  const esc = SU.esc;

  /* ---------- Consolidate ---------- */
  const FNS = { sum: ['Sum', 'SUM'], count: ['Count', 'COUNTA'], average: ['Average', 'AVERAGE'], max: ['Max', 'MAX'], min: ['Min', 'MIN'], product: ['Product', 'PRODUCT'] };
  DT.lastRefs = [];
  DT.consolidateDialog = async () => {
    const s = S(); let refs = DT.lastRefs.slice();
    const body = document.createElement('div');
    const render = () => {
      body.innerHTML = `<div class="dlg-grid"><label>Function</label><select id="cn-fn">${Object.entries(FNS).map(([k, [l]]) => `<option value="${k}" ${DT.lastFn === k ? 'selected' : ''}>${l}</option>`).join('')}</select>
        <label>Reference</label><div style="display:flex;gap:6px"><input id="cn-ref" style="flex:1" placeholder="e.g. 'Term 1'!A1:D10" value="${esc(s.sel.r1 > s.sel.r0 || s.sel.c1 > s.sel.c0 ? SU.quoteSheet(s.name) + '!' + SU.absRangeStr(A.effData(s.sel)) : '')}"><button class="btn" data-a="add">Add</button></div></div>
        <div class="dlg-note">All references</div><div class="cn-list">${refs.map((r, i) => `<div data-i="${i}"><span>${esc(r)}</span><button class="rb" data-a="del" data-i="${i}" title="Delete reference">${icon('x', 14)}</button></div>`).join('') || '<div class="empty-hint">Add the ranges you want to combine (from any sheet).</div>'}</div>
        <div style="display:flex;gap:16px;margin-top:8px"><label class="check"><input type="checkbox" id="cn-top" ${DT.lastTop ? 'checked' : ''}> Use labels in top row</label><label class="check"><input type="checkbox" id="cn-left" ${DT.lastLeft ? 'checked' : ''}> Left column</label><label class="check"><input type="checkbox" id="cn-links" ${DT.lastLinks ? 'checked' : ''}> Create links to source data</label></div>
        <p class="dlg-note">Results are placed starting at ${SU.addr(s.active.r, s.active.c)} on this sheet.</p>`;
    };
    body.addEventListener('click', (e) => {
      const b = e.target.closest('[data-a]'); if (!b) return;
      if (b.dataset.a === 'add') { const v = body.querySelector('#cn-ref').value.trim(); if (!SU.parseRef(v)) { BTS.toast('That reference isn’t valid.', { error: true }); return; } if (!refs.includes(v)) refs.push(v); const keep = { fn: body.querySelector('#cn-fn').value, top: body.querySelector('#cn-top').checked, left: body.querySelector('#cn-left').checked, links: body.querySelector('#cn-links').checked }; DT.lastFn = keep.fn; DT.lastTop = keep.top; DT.lastLeft = keep.left; DT.lastLinks = keep.links; render(); }
      if (b.dataset.a === 'del') { refs.splice(+b.dataset.i, 1); render(); }
    });
    render();
    const res = await BTS.dialog({ title: 'Consolidate', width: 540, body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ fn: r.querySelector('#cn-fn').value, top: r.querySelector('#cn-top').checked, left: r.querySelector('#cn-left').checked, links: r.querySelector('#cn-links').checked, pending: r.querySelector('#cn-ref').value.trim() }) }] });
    if (!res) return;
    if (!refs.length && res.pending && SU.parseRef(res.pending)) refs.push(res.pending);
    DT.lastRefs = refs; DT.lastFn = res.fn; DT.lastTop = res.top; DT.lastLeft = res.left; DT.lastLinks = res.links;
    if (!refs.length) { BTS.toast('Add at least one reference.', { error: true }); return; }
    DT.consolidate(refs, res);
  };
  /** refs: ["Sheet1!A1:C5", ...]; opts {fn, top, left, links}. Writes the result at the active cell. */
  DT.consolidate = (refs, opts) => {
    const s = S(); const fn = opts.fn || 'sum';
    const srcs = refs.map(t => { const r = SU.parseRef(t); const sh = r && (r.sheet ? wb.sheetByName(r.sheet) : s); return r && sh ? { sh, g: A.effData(r.range, sh) } : null; }).filter(Boolean);
    if (!srcs.length) { BTS.toast('None of the references could be read.', { error: true }); return; }
    const rowL = [], colL = []; const cellsBy = new Map(); // "ri|ci" -> [{sh,r,c}]
    const idx = (list, label) => { let i = list.findIndex(x => x.toLowerCase() === label.toLowerCase()); if (i < 0) { list.push(label); i = list.length - 1; } return i; };
    for (const { sh, g } of srcs) {
      const r0 = g.r0 + (opts.top ? 1 : 0), c0 = g.c0 + (opts.left ? 1 : 0);
      for (let r = r0; r <= g.r1; r++) for (let c = c0; c <= g.c1; c++) {
        const ri = opts.left ? idx(rowL, String(sh.displayText(r, g.c0)).trim() || '(blank)') : r - r0;
        const ci = opts.top ? idx(colL, String(sh.displayText(g.r0, c)).trim() || '(blank)') : c - c0;
        const k = ri + '|' + ci; let a = cellsBy.get(k); if (!a) { a = []; cellsBy.set(k, a); } a.push({ sh, r, c });
      }
    }
    let H = 0, W = 0; for (const k of cellsBy.keys()) { const [ri, ci] = k.split('|').map(Number); H = Math.max(H, ri + 1); W = Math.max(W, ci + 1); }
    const R0 = s.active.r, C0 = s.active.c; const dr = opts.top ? 1 : 0, dc = opts.left ? 1 : 0;
    const cells = [];
    if (opts.top) colL.forEach((l, i) => cells.push({ r: R0, c: C0 + dc + i, v: l }));
    if (opts.left) rowL.forEach((l, i) => cells.push({ r: R0 + dr + i, c: C0, v: l }));
    const ref = (p) => (p.sh === s ? '' : SU.quoteSheet(p.sh.name) + '!') + SU.addr(p.r, p.c);
    for (const [k, list] of cellsBy) {
      const [ri, ci] = k.split('|').map(Number); const r = R0 + dr + ri, c = C0 + dc + ci;
      if (opts.links) { cells.push({ r, c, v: '=' + FNS[fn][1] + '(' + list.map(ref).join(',') + ')' }); continue; }
      const vals = list.map(p => p.sh.value(p.r, p.c)); const nums = vals.filter(v => typeof v === 'number');
      let v = null;
      if (fn === 'count') v = vals.filter(x => x !== null && x !== '').length;
      else if (!nums.length) v = null;
      else if (fn === 'sum') v = nums.reduce((a, b) => a + b, 0); else if (fn === 'average') v = nums.reduce((a, b) => a + b, 0) / nums.length;
      else if (fn === 'max') v = Math.max(...nums); else if (fn === 'min') v = Math.min(...nums); else if (fn === 'product') v = nums.reduce((a, b) => a * b, 1);
      cells.push({ r, c, v: v === null ? null : +v.toPrecision(15) });
    }
    if (A.protectedCheck(false, { r0: R0, c0: C0, r1: R0 + H + dr - 1, c1: C0 + W + dc - 1 })) return;
    const nfSrc = srcs[0].sh.style(srcs[0].g.r0 + dr, srcs[0].g.c0 + dc).nf;
    wb.run('Consolidate', () => { wb.setContents(s, cells); if (opts.top) colL.forEach((_, i) => wb.setStyle(s, R0, C0 + dc + i, SU.mergeStyle(s.style(R0, C0 + dc + i), { b: true }))); if (nfSrc) for (const x of cells) if (x.r >= R0 + dr && x.c >= C0 + dc) wb.setStyle(s, x.r, x.c, SU.mergeStyle(s.style(x.r, x.c), { nf: nfSrc })); });
    s.sel = { r0: R0, c0: C0, r1: R0 + H + dr - 1, c1: C0 + W + dc - 1 }; Grid.notify();
    BTS.toast(`Consolidated ${srcs.length} range${srcs.length > 1 ? 's' : ''}`);
  };

  /* ---------- Group / Ungroup (outline) ---------- */
  const propOf = (axis) => axis === 'r' ? 'outlineR' : 'outlineC', hidOf = (axis) => axis === 'r' ? 'hiddenR' : 'hiddenC';
  async function pickAxis(title) {
    const s = S(); const g = s.sel;
    if (Grid.isWholeRows(g)) return 'r'; if (Grid.isWholeCols(g)) return 'c';
    return BTS.dialog({ title, width: 260, body: `<label class="check"><input type="radio" name="gx" value="r" checked> Rows</label><label class="check"><input type="radio" name="gx" value="c"> Columns</label>`, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => r.querySelector('[name=gx]:checked').value }] });
  }
  const span = (g, axis) => axis === 'r' ? [g.r0, Math.min(g.r1, S().rowCount - 1)] : [g.c0, Math.min(g.c1, S().colCount - 1)];
  DT.group = async (axis) => {
    const s = S(); axis = axis || await pickAxis('Group'); if (!axis) return;
    if (A.protectedCheck(false, true, axis === 'r' ? 'formatRows' : 'formatColumns')) return;
    const [a, b] = span(A.eff(s.sel), axis); const m = new Map(s[propOf(axis)]); let capped = false;
    for (let i = a; i <= b; i++) { const L = (m.get(i) || 0) + 1; if (L > 7) { capped = true; continue; } m.set(i, L); }
    if (capped) BTS.toast('Outlines can be at most 7 levels deep.', { error: true });
    wb.run('Group', () => wb.setProp(s, propOf(axis), m)); Grid.updateSpacer(); Grid.notify(); Ed.reposition();
  };
  DT.ungroup = async (axis) => {
    const s = S(); axis = axis || await pickAxis('Ungroup'); if (!axis) return;
    const [a, b] = span(A.eff(s.sel), axis); const m = new Map(s[propOf(axis)]); const hid = new Set(s[hidOf(axis)]); let any = false;
    for (let i = a; i <= b; i++) { const L = m.get(i); if (!L) continue; any = true; if (L <= 1) { m.delete(i); hid.delete(i); } else m.set(i, L - 1); }
    if (!any) { BTS.toast('The selection isn’t grouped.', { error: true }); return; }
    wb.run('Ungroup', () => { wb.setProp(s, propOf(axis), m); wb.setProp(s, hidOf(axis), hid); }); Grid.updateSpacer(); Grid.notify(); Ed.reposition();
  };
  DT.clearOutline = () => {
    const s = S(); if (!s.outlineR.size && !s.outlineC.size) { BTS.toast('This sheet has no outline.', { error: true }); return; }
    const hr = new Set(s.hiddenR), hc = new Set(s.hiddenC); for (const k of s.outlineR.keys()) hr.delete(k); for (const k of s.outlineC.keys()) hc.delete(k);
    wb.run('Clear outline', () => { wb.setProp(s, 'outlineR', new Map()); wb.setProp(s, 'outlineC', new Map()); wb.setProp(s, 'hiddenR', hr); wb.setProp(s, 'hiddenC', hc); }); Grid.updateSpacer(); Grid.notify();
  };
  function setCollapsed(s, axis, g, collapse) {
    const hid = new Set(s[hidOf(axis)]); const map = s[propOf(axis)];
    for (let i = g.a; i <= g.b; i++) { if (collapse) hid.add(i); else if ((map.get(i) || 0) >= g.L) hid.delete(i); }
    wb.run(collapse ? 'Hide detail' : 'Show detail', () => wb.setProp(s, hidOf(axis), hid));
    Grid.updateSpacer(); Grid.notify(); Ed.reposition();
  }
  /** outline button clicked on the grid gutter */
  DT.onOutline = (o) => {
    const s = S();
    if (o.kind === 'toggle') return setCollapsed(s, o.axis, o.g, !o.collapsed);
    const hid = new Set(s[hidOf(o.axis)]); for (const [i, L] of s[propOf(o.axis)]) { if (L >= o.L) hid.add(i); else hid.delete(i); }
    wb.run('Outline level ' + o.L, () => wb.setProp(s, hidOf(o.axis), hid)); Grid.updateSpacer(); Grid.notify(); Ed.reposition();
  };
  /** innermost group touching the active row/column (including its summary line) */
  function groupAtActive(axis) {
    const s = S(); const map = s[propOf(axis)]; if (!map.size) return null; const idx = axis === 'r' ? s.active.r : s.active.c;
    const info = Grid.outlineInfo(map); let best = null;
    for (const g of info.groups) if (idx >= g.a && idx <= g.b + 1 && (!best || g.L > best.L)) best = g;
    return best;
  }
  DT.detail = (show) => {
    const s = S(); for (const axis of Grid.isWholeCols(s.sel) ? ['c', 'r'] : ['r', 'c']) { const g = groupAtActive(axis); if (g) { setCollapsed(s, axis, g, !show); return; } }
    BTS.toast('Select a cell in (or just below) a group first.', { error: true });
  };

  /* ---------- Circle invalid data ---------- */
  DT.circleInvalid = (silent) => {
    const s = S(); const set = new Set(); const u = s.usedRange();
    for (const dv of s.validations) { const g = { r0: dv.range.r0, c0: dv.range.c0, r1: Math.min(dv.range.r1, u.r1), c1: Math.min(dv.range.c1, u.c1) }; for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) { if (set.size >= 2000) break; if (s.validationAt(r, c) !== dv) continue; const v = s.value(r, c); if (v === null || v === '') continue; if (!A.checkValidation(s, dv, SU.isErr(v) ? v.value : v)) set.add(SU.key(r, c)); } }
    s._circles = set; Grid.render();
    if (!silent) BTS.toast(set.size ? `${set.size} cell${set.size > 1 ? 's' : ''} with invalid data circled` : (s.validations.length ? 'All data is valid' : 'This sheet has no data validation rules'), { error: !set.size && !s.validations.length });
  };
  DT.clearCircles = () => { const s = S(); s._circles = null; Grid.render(); };
  DT.afterChange = () => { const s = wb.active; if (s && s._circles) DT.circleInvalid(true); };
})();
