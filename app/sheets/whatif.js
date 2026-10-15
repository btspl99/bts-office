/* BTS Sheets — What-If analysis: Goal Seek (numeric solver) and Scenario Manager.
 * Scenarios are stored on the workbook: wb.scenarios = [{ id, name, sheet, cells:'B2:B4', values:[raw], comment, by, time }] */
(function () {
  const WhatIf = window.WhatIf = {};
  let wb;
  WhatIf.init = (w) => { wb = w; };
  const S = () => wb.active;
  const esc = SU.esc;

  /* ---------- Goal Seek ---------- */
  /** Solve f(x) = target for the changing cell; returns {x, value, ok, iterations}. Cell contents are restored afterwards. */
  WhatIf.solve = (setSheet, setCell, target, chSheet, chCell) => {
    const hf = wb.hf; const addr = chSheet.a(chCell.r, chCell.c); const orig = chSheet.raw(chCell.r, chCell.c);
    let calls = 0;
    const f = (x) => { calls++; hf.setCellContents(addr, [[x]]); const v = setSheet.value(setCell.r, setCell.c); return typeof v === 'number' && isFinite(v) ? v - target : NaN; };
    const tol = 1e-9 * Math.max(1, Math.abs(target)); let best = null;
    const consider = (x, fx) => { if (!isNaN(fx) && (!best || Math.abs(fx) < Math.abs(best.fx))) best = { x, fx }; };
    try {
      let x0 = typeof orig === 'number' ? orig : 0, f0 = f(x0); consider(x0, f0);
      if (Math.abs(f0) <= tol) return { x: x0, ok: true, iterations: calls, value: target + f0 };
      // secant
      let x1 = x0 !== 0 ? x0 * 1.01 + 0.001 : 0.01, f1 = f(x1); consider(x1, f1);
      for (let i = 0; i < 100 && !isNaN(f0) && !isNaN(f1); i++) {
        if (f1 === f0) break; const x2 = x1 - f1 * (x1 - x0) / (f1 - f0); if (!isFinite(x2)) break;
        x0 = x1; f0 = f1; x1 = x2; f1 = f(x1); consider(x1, f1); if (Math.abs(f1) <= tol) break;
      }
      if (best && Math.abs(best.fx) <= tol * 10) return { x: best.x, ok: true, iterations: calls, value: target + best.fx };
      // bracket + bisection
      const start = typeof orig === 'number' ? orig : 0; const fs = f(start); let lo = null, hi = null;
      for (let k = 0; k < 80 && lo === null; k++) {
        const step = Math.max(1, Math.abs(start)) * 0.01 * Math.pow(2, k);
        for (const x of [start + step, start - step]) { const fx = f(x); consider(x, fx); if (!isNaN(fx) && !isNaN(fs) && Math.sign(fx) !== Math.sign(fs)) { lo = start; hi = x; break; } }
      }
      if (lo !== null) {
        let flo = f(lo);
        for (let i = 0; i < 200; i++) { const mid = (lo + hi) / 2, fm = f(mid); consider(mid, fm); if (Math.abs(fm) <= tol || Math.abs(hi - lo) < 1e-12 * Math.max(1, Math.abs(mid))) break; if (Math.sign(fm) === Math.sign(flo)) { lo = mid; flo = fm; } else hi = mid; }
      }
      if (!best) return { ok: false, iterations: calls };
      return { x: +best.x.toPrecision(15), ok: Math.abs(best.fx) <= Math.max(tol * 1000, 1e-6 * Math.max(1, Math.abs(target))), iterations: calls, value: target + best.fx };
    } finally { hf.setCellContents(addr, [[wb._toHFContent(orig)]]); }
  };
  WhatIf.goalSeek = async (preset = {}) => {
    const s = S();
    const res = await BTS.dialog({ title: 'Goal Seek', width: 380, body: `<div class="dlg-grid">
      <label>Set cell</label><input id="gs-set" value="${esc(preset.set || SU.addr(s.active.r, s.active.c))}">
      <label>To value</label><input id="gs-to" value="${esc(preset.to ?? '')}" placeholder="e.g. 50000">
      <label>By changing cell</label><input id="gs-ch" value="${esc(preset.ch || '')}" placeholder="e.g. B3"></div>
      <p class="dlg-note">Finds the input value that makes a formula return the result you want.</p>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ set: r.querySelector('#gs-set').value, to: r.querySelector('#gs-to').value, ch: r.querySelector('#gs-ch').value }) }] });
    if (!res) return;
    return WhatIf.runGoalSeek(res.set, res.to, res.ch);
  };
  WhatIf.runGoalSeek = async (setRef, toVal, chRef, auto) => {
    const s = S();
    const sr = SU.parseRef(setRef), cr = SU.parseRef(chRef); const target = parseFloat(String(toVal).replace(/,/g, ''));
    const err = (m) => { BTS.toast(m, { error: true, ms: 4000 }); return null; };
    if (!sr || sr.range.r0 !== sr.range.r1 || sr.range.c0 !== sr.range.c1) return err('“Set cell” must be a single cell reference.');
    if (!cr || cr.range.r0 !== cr.range.r1 || cr.range.c0 !== cr.range.c1) return err('“By changing cell” must be a single cell reference.');
    if (isNaN(target)) return err('“To value” must be a number.');
    const ss = sr.sheet ? wb.sheetByName(sr.sheet) : s, cs = cr.sheet ? wb.sheetByName(cr.sheet) : s; if (!ss || !cs) return err('That sheet doesn’t exist.');
    const set = { r: sr.range.r0, c: sr.range.c0 }, ch = { r: cr.range.r0, c: cr.range.c0 };
    if (!ss.isFormula(set.r, set.c)) return err('The “Set cell” must contain a formula.');
    if (cs.isFormula(ch.r, ch.c)) return err('The “By changing cell” must contain a value, not a formula.');
    const raw = cs.raw(ch.r, ch.c); if (raw !== null && typeof raw !== 'number') return err('The “By changing cell” must contain a number.');
    if (cs.protected && cs.style(ch.r, ch.c).locked !== false) return err('The changing cell is on a protected sheet.');
    const t0 = performance.now(); const r = WhatIf.solve(ss, set, target, cs, ch); const ms = performance.now() - t0;
    if (!r || r.x === undefined) return err('Goal Seek couldn’t find a solution.');
    WhatIf.last = { ...r, ms };
    const apply = () => wb.run('Goal Seek', () => wb.setContents(cs, [{ r: ch.r, c: ch.c, v: r.x }]));
    if (auto) { apply(); return r; }
    apply(); // show the result live like Excel; Cancel undoes it
    const ok = await BTS.dialog({ title: 'Goal Seek status', width: 380, body: `<p style="margin:4px 0 10px;line-height:1.5">Goal Seeking with cell ${SU.addr(set.r, set.c)} ${r.ok ? 'found a solution.' : 'may not have found a solution.'}</p><div class="dlg-grid"><label>Target value</label><b>${esc(SU.generalNum(target))}</b><label>Current value</label><b>${esc(ss.displayText(set.r, set.c))}</b><label>${SU.addr(ch.r, ch.c)} set to</label><b>${esc(SU.generalNum(+r.x.toPrecision(12)))}</b></div><p class="dlg-note">${r.iterations} iterations · ${Math.round(ms)} ms</p>`, buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, value: true }] });
    if (!ok) { wb.undo(); wb.redoStack.pop(); Grid.notify(); }
    return r;
  };

  /* ---------- Scenario Manager ---------- */
  const cellsOf = (g) => { const out = []; for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) out.push({ r, c }); return out; };
  WhatIf.applyScenario = (sc) => {
    const s = wb.sheetByName(sc.sheet); if (!s) return; const g = SU.parseRange(sc.cells); if (!g) return;
    const cells = cellsOf(g).map((p, i) => ({ ...p, v: sc.values[i] ?? null }));
    wb.run('Show scenario', () => wb.setContents(s, cells));
    Grid.notify(); BTS.toast('Showing scenario “' + sc.name + '”');
  };
  async function editScenario(sc) {
    const s = S(); const isNew = !sc;
    const res = await BTS.dialog({ title: isNew ? 'Add scenario' : 'Edit scenario', width: 420, body: `<div class="field"><label>Scenario name</label><input id="sc-name" value="${esc(sc ? sc.name : '')}"></div><div class="field"><label>Changing cells</label><input id="sc-cells" value="${esc(sc ? sc.cells : SU.rangeStr(A.effData(s.sel)))}"></div><div class="field"><label>Comment</label><textarea id="sc-cm" style="height:56px">${esc(sc ? sc.comment || '' : 'Created by ' + (BTS.userName || 'BTS Sheets') + ' on ' + new Date().toLocaleDateString())}</textarea></div>`, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ name: r.querySelector('#sc-name').value.trim(), cells: r.querySelector('#sc-cells').value.trim(), comment: r.querySelector('#sc-cm').value }) }] });
    if (!res) return null;
    if (!res.name) { BTS.toast('Enter a scenario name.', { error: true }); return null; }
    const g = SU.parseRange(res.cells.replace(/^.*!/, '')); if (!g) { BTS.toast('The changing cells reference isn’t valid.', { error: true }); return null; }
    if ((g.r1 - g.r0 + 1) * (g.c1 - g.c0 + 1) > 32) { BTS.toast('A scenario can change up to 32 cells.', { error: true }); return null; }
    const cells = cellsOf(g); const prev = sc && sc.cells === SU.rangeStr(g) ? sc.values : cells.map(p => s.raw(p.r, p.c));
    const vals = await BTS.dialog({ title: 'Scenario values', width: 380, body: `<p class="dlg-note" style="margin:0 0 8px">Enter values for each of the changing cells.</p><div class="dlg-grid">${cells.map((p, i) => `<label>${SU.addr(p.r, p.c)}</label><input data-i="${i}" value="${esc(prev[i] === null || prev[i] === undefined ? '' : String(prev[i]))}">`).join('')}</div>`, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => [...r.querySelectorAll('[data-i]')].map(x => x.value) }] });
    if (!vals) return null;
    const values = vals.map((t, i) => { const p = A.interpret(s, cells[i].r, cells[i].c, t); return p.error ? t : p.v; });
    return { id: sc ? sc.id : 'sc' + Date.now().toString(36), name: res.name, sheet: s.name, cells: SU.rangeStr(g), values, comment: res.comment, by: BTS.userName || '', time: new Date().toISOString() };
  }
  WhatIf.scenarioManager = async () => {
    const body = document.createElement('div'); let sel = 0;
    const list = () => wb.scenarios.filter(sc => sc.sheet === S().name);
    const render = () => {
      const L = list(); sel = Math.min(sel, Math.max(0, L.length - 1));
      const cur = L[sel];
      body.innerHTML = `<div class="sc-wrap"><div class="sc-list">${L.map((sc, i) => `<div data-i="${i}" class="${i === sel ? 'on' : ''}">${esc(sc.name)}</div>`).join('') || '<div class="empty-hint">No scenarios defined. Choose Add to add scenarios.</div>'}</div>
        <div class="sc-btns"><button class="btn" data-a="add">${icon('plus', 15)} Add…</button><button class="btn" data-a="del" ${cur ? '' : 'disabled'}>${icon('trash-2', 15)} Delete</button><button class="btn" data-a="edit" ${cur ? '' : 'disabled'}>${icon('pencil', 15)} Edit…</button><button class="btn" data-a="sum" ${L.length ? '' : 'disabled'}>${icon('table', 15)} Summary…</button></div></div>
        <div class="dlg-grid" style="margin-top:10px"><label>Changing cells</label><span>${cur ? esc(cur.cells) : '—'}</span><label>Comment</label><span style="white-space:pre-wrap">${cur ? esc(cur.comment || '') : '—'}</span></div>`;
    };
    body.addEventListener('click', async (e) => {
      const it = e.target.closest('[data-i]'); if (it) { sel = +it.dataset.i; render(); return; }
      const b = e.target.closest('[data-a]'); if (!b) return; const L = list(); const cur = L[sel];
      if (b.dataset.a === 'add') { const sc = await editScenario(null); if (sc) { wb.run('Add scenario', () => wb.setProp(wb, 'scenarios', [...wb.scenarios, sc])); sel = list().length - 1; } render(); }
      if (b.dataset.a === 'edit' && cur) { const sc = await editScenario(cur); if (sc) wb.run('Edit scenario', () => wb.setProp(wb, 'scenarios', wb.scenarios.map(x => x.id === cur.id ? sc : x))); render(); }
      if (b.dataset.a === 'del' && cur) { wb.run('Delete scenario', () => wb.setProp(wb, 'scenarios', wb.scenarios.filter(x => x.id !== cur.id))); render(); }
      if (b.dataset.a === 'sum') { const x = body.closest('.modal-back'); x && x.querySelector('[data-x]').click(); setTimeout(() => WhatIf.summary(), 50); }
    });
    body.addEventListener('dblclick', (e) => { const it = e.target.closest('[data-i]'); if (it) { const sc = list()[+it.dataset.i]; if (sc) WhatIf.applyScenario(sc); } });
    render();
    const res = await BTS.dialog({ title: 'Scenario Manager', width: 520, body, buttons: [{ label: 'Close', value: 'close' }, { label: 'Show', primary: true, value: () => 'show' }] });
    if (res === 'show') { const sc = list()[sel]; if (sc) WhatIf.applyScenario(sc); }
  };
  /** compute a result cell for each scenario without changing the workbook */
  WhatIf.evaluate = (sc, resultCells) => {
    const s = wb.sheetByName(sc.sheet); const g = SU.parseRange(sc.cells); const cells = cellsOf(g); const hf = wb.hf;
    const orig = cells.map(p => s.raw(p.r, p.c));
    try { hf.batch(() => cells.forEach((p, i) => hf.setCellContents(s.a(p.r, p.c), [[wb._toHFContent(sc.values[i] ?? null)]]))); return resultCells.map(rc => rc.sheet.value(rc.r, rc.c)); }
    finally { hf.batch(() => cells.forEach((p, i) => hf.setCellContents(s.a(p.r, p.c), [[wb._toHFContent(orig[i])]]))); }
  };
  WhatIf.summary = async (resultStr) => {
    const s = S(); const L = wb.scenarios.filter(sc => sc.sheet === s.name); if (!L.length) { BTS.toast('Add a scenario first.', { error: true }); return; }
    if (resultStr === undefined) resultStr = await BTS.prompt('Scenario summary', 'Result cells (e.g. B10 or B10:B12)', SU.rangeStr(A.effData(s.sel)));
    if (!resultStr) return; const rg = SU.parseRange(resultStr.replace(/^.*!/, '')); if (!rg) { BTS.toast('That reference isn’t valid.', { error: true }); return; }
    const results = cellsOf(rg).map(p => ({ ...p, sheet: s }));
    const allChanging = [...new Set(L.flatMap(sc => cellsOf(SU.parseRange(sc.cells)).map(p => SU.addr(p.r, p.c))))].map(SU.parseAddr);
    const rows = [];
    rows.push([{ v: 'Scenario Summary', s: { font: STheme.cur().head, sz: 15, b: true, color: STheme.cur().dk2 } }]);
    rows.push([null, { v: 'Current Values:', s: { b: true } }, ...L.map(sc => ({ v: sc.name, s: { b: true } }))]);
    rows.push([{ v: 'Changing Cells:', s: { b: true } }]);
    for (const p of allChanging) rows.push([null, { v: '$' + SU.colName(p.c) + '$' + (p.r + 1) }, ...[s.raw(p.r, p.c)].map(v => ({ v: typeof v === 'string' && v[0] === '=' ? s.value(p.r, p.c) : v })), ...L.map(sc => { const g = SU.parseRange(sc.cells); const i = cellsOf(g).findIndex(q => q.r === p.r && q.c === p.c); return { v: i >= 0 ? sc.values[i] : s.value(p.r, p.c) }; })]);
    rows.push([{ v: 'Result Cells:', s: { b: true } }]);
    const evals = L.map(sc => WhatIf.evaluate(sc, results));
    results.forEach((p, ri) => rows.push([null, { v: '$' + SU.colName(p.c) + '$' + (p.r + 1) }, { v: s.value(p.r, p.c), s: { nf: s.style(p.r, p.c).nf } }, ...evals.map(ev => ({ v: SU.isErr(ev[ri]) ? ev[ri].value : ev[ri], s: { nf: s.style(p.r, p.c).nf } }))]));
    rows.push([]); rows.push([{ v: 'Notes: Current Values column represents values of changing cells at the time the Scenario Summary Report was created.', s: { i: true, color: '#5c6577' } }]);
    wb.run('Scenario summary', () => {
      const ns = wb.addSheet(wb.uniqueName('Scenario Summary'), wb.activeIndex + 1);
      const cells = []; rows.forEach((row, r) => row.forEach((cell, c) => { if (!cell) return; if (cell.v !== null && cell.v !== undefined) cells.push({ r, c, v: cell.v }); if (cell.s) wb.setStyle(ns, r, c, Object.fromEntries(Object.entries(cell.s).filter(([, v]) => v !== undefined))); }));
      wb.setContents(ns, cells);
      const acc = STheme.cur().accents[0];
      for (let c = 1; c <= L.length + 2; c++) wb.setStyle(ns, 1, c, { ...ns.style(1, c), b: true, fill: STheme.tint(acc, 0.8), bB: { style: 'thin', color: acc } });
      wb.setProp(ns, 'colW', new Map([[0, 110], [1, 110], [2, 110], ...L.map((_, i) => [3 + i, 110])]));
      wb.activeIndex = wb.sheets.indexOf(ns);
    });
    App.renderTabs(); Grid.restoreScroll(); Charts.renderSheet(); Grid.notify();
  };
})();
