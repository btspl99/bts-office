/* BTS Sheets — dialogs & popovers: functions, find/replace, go to, sort, filter, validation, CF, names, text to columns, symbols, links, notes, sizes, paste special, series, format cells */
(function () {
  const Dlg = window.Dlg = {};
  let wb;
  Dlg.init = (w) => { wb = w; };
  const S = () => wb.active;
  const $ = (s, r = document) => r.querySelector(s);
  const esc = SU.esc;

  /* ---------- Insert function ---------- */
  Dlg.insertFunction = async (preset) => {
    let cat = preset ? (SFUNCS.byName.get(preset) || {}).cat || 'All' : 'Most used', sel = preset || 'SUM';
    const MOST = ['SUM', 'AVERAGE', 'IF', 'COUNT', 'MAX', 'MIN', 'VLOOKUP', 'XLOOKUP', 'COUNTIF', 'SUMIF', 'ROUND', 'IFERROR', 'CONCATENATE', 'TODAY', 'INDEX', 'MATCH'];
    const body = document.createElement('div');
    body.innerHTML = `<div class="field"><input id="fn-search" placeholder="Search for a function — e.g. “add numbers”, “lookup”, “average”"></div><div class="fn-dlg"><div class="fn-cats"></div><div class="fn-right"><div class="fn-list"></div><div class="fn-desc"></div></div></div>`;
    const cats = ['Most used', 'All', ...SFUNCS.cats];
    const render = () => {
      const q = $('#fn-search', body).value.trim().toLowerCase();
      $('.fn-cats', body).innerHTML = cats.map(c => `<div data-cat="${esc(c)}" class="${c === cat && !q ? 'on' : ''}">${esc(c)}</div>`).join('');
      let list = q ? SFUNCS.list.filter(f => f.name.toLowerCase().includes(q) || f.desc.toLowerCase().includes(q)) : cat === 'All' ? SFUNCS.list.slice().sort((a, b) => a.name.localeCompare(b.name)) : cat === 'Most used' ? MOST.map(n => SFUNCS.byName.get(n)).filter(Boolean) : SFUNCS.list.filter(f => f.cat === cat);
      if (!list.find(f => f.name === sel) && list.length) sel = list[0].name;
      $('.fn-list', body).innerHTML = list.map(f => `<div data-fn="${f.name}" class="${f.name === sel ? 'on' : ''}">${f.name}</div>`).join('') || '<div class="empty-hint">No functions match.</div>';
      const f = SFUNCS.byName.get(sel); $('.fn-desc', body).innerHTML = f ? `<b>${f.name}(${esc(f.args)})</b><br>${esc(f.desc)}` : '';
      const on = $('.fn-list .on', body); on && on.scrollIntoView({ block: 'nearest' });
    };
    body.addEventListener('click', e => { const c = e.target.closest('[data-cat]'); if (c) { cat = c.dataset.cat; $('#fn-search', body).value = ''; render(); } const f = e.target.closest('[data-fn]'); if (f) { sel = f.dataset.fn; render(); } });
    body.addEventListener('dblclick', e => { if (e.target.closest('[data-fn]')) $('.modal footer .primary').click(); });
    $('#fn-search', body).addEventListener('input', render);
    render();
    const name = await BTS.dialog({ title: 'Insert function', body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: () => sel }] });
    if (!name) return;
    Dlg.functionArgs(name);
  };
  Dlg.functionArgs = async (name) => {
    const f = SFUNCS.byName.get(name);
    if (!f || !f.args) { insertCall(name + '()'); return; }
    const args = f.args.split(', ').filter(a => a !== '...');
    const extra = f.args.includes('...');
    const body = document.createElement('div'); body.className = 'args-dlg';
    const s = S();
    // guess first range argument
    let guess = '';
    if (/^(SUM|AVERAGE|COUNT|MAX|MIN|MEDIAN|PRODUCT|COUNTA|STDEV)/.test(name)) { const r = guessRange(s); if (r) guess = SU.rangeStr(r); }
    const rows = args.map((a, i) => `<div class="arg"><label class="${a.startsWith('[') ? 'opt' : ''}">${esc(a.replace(/[\[\]]/g, ''))}</label><input data-i="${i}" value="${i === 0 ? esc(guess) : ''}"><span class="v" data-v="${i}"></span></div>`).join('');
    body.innerHTML = `<p style="margin:0 0 10px;color:var(--ink-2)">${esc(f.desc)}</p>${rows}${extra ? '<p class="dlg-note">You can add more arguments in the formula bar after inserting.</p>' : ''}<div class="dlg-note" id="fa-result" style="font-size:13px;margin-top:10px"></div><p class="dlg-note">Tip: type a cell reference like A1 or a range like B2:B10. Text needs "quotes".</p>`;
    const build = () => { const vals = [...body.querySelectorAll('input')].map(i => i.value.trim()); while (vals.length && vals[vals.length - 1] === '') vals.pop(); return '=' + name + '(' + vals.join(',') + ')'; };
    const upd = () => {
      body.querySelectorAll('input').forEach(inp => { const v = inp.value.trim(); const sp = body.querySelector(`[data-v="${inp.dataset.i}"]`); if (!v) { sp.textContent = ''; return; } try { const r = wb.hf.calculateFormula(SU.toHF('=' + v), s.id); sp.textContent = '= ' + (r && r.data ? '{…}' : SU.valueText(r)); } catch { sp.textContent = ''; } });
      try { const r = wb.hf.calculateFormula(SU.toHF(build()), s.id); $('#fa-result', body).innerHTML = 'Formula result = <b>' + esc(r && r.data ? '{array}' : SU.valueText(r)) + '</b>'; } catch { $('#fa-result', body).textContent = ''; }
    };
    body.addEventListener('input', upd); setTimeout(upd, 0);
    const res = await BTS.dialog({ title: 'Function arguments — ' + name, body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: build }] });
    if (res) insertCall(res);
  };
  function guessRange(s) {
    const { r, c } = s.active; let top = r - 1; if (top < 0 || typeof s.value(top, c) !== 'number') { let left = c - 1; if (left >= 0 && typeof s.value(r, left) === 'number') { while (left > 0 && typeof s.value(r, left - 1) === 'number') left--; return { r0: r, c0: left, r1: r, c1: c - 1 }; } return null; }
    while (top > 0 && typeof s.value(top - 1, c) === 'number') top--; return { r0: top, c0: c, r1: r - 1, c1: c };
  }
  function insertCall(text) {
    if (Ed.isEditing()) { const cur = document.getElementById('fx').value; if (cur.startsWith('=')) { Ed.insertAtCaret(text.replace(/^=/, '')); return; } }
    if (A.protectedCheck()) return;
    Ed.start(text, 'edit');
  }
  Dlg.insertCall = insertCall;
  Dlg.fnMenu = (anchor, cat) => {
    const list = SFUNCS.list.filter(f => f.cat === cat);
    BTS.menu(anchor, [...list.map(f => ({ label: f.name, html: `<span style="min-width:110px;font-weight:600">${f.name}</span><span style="color:var(--ink-3);font-size:11.5px;max-width:300px;overflow:hidden;text-overflow:ellipsis">${esc(f.desc)}</span>`, action: () => Dlg.functionArgs(f.name) })), { sep: true }, { label: 'Insert function…', icon: 'square-function', action: () => Dlg.insertFunction() }]);
  };

  /* ---------- Find & replace ---------- */
  let findState = { q: '', rep: '', matchCase: false, whole: false, within: 'sheet', lookIn: 'values', idx: -1, results: [] };
  Dlg.find = async (replace) => {
    const body = document.createElement('div'); body.className = 'find-pane modal-noenter';
    const s = S();
    body.innerHTML = `<div class="field"><label>Find what</label><input id="fd-q" value="${esc(findState.q)}"></div>
      <div class="field ${replace ? '' : 'hidden'}" id="fd-rep-f"><label>Replace with</label><input id="fd-rep" value="${esc(findState.rep)}"></div>
      <div class="row"><label class="check"><input type="checkbox" id="fd-case" ${findState.matchCase ? 'checked' : ''}> Match case</label><label class="check" style="margin-left:12px"><input type="checkbox" id="fd-whole" ${findState.whole ? 'checked' : ''}> Match entire cell</label></div>
      <div class="row"><label style="font-size:12px;color:var(--ink-2)">Within</label><select id="fd-within" class="rsel"><option value="sheet">Sheet</option><option value="book" ${findState.within === 'book' ? 'selected' : ''}>Workbook</option></select>
      <label style="font-size:12px;color:var(--ink-2);margin-left:10px">Look in</label><select id="fd-in" class="rsel"><option value="values">Values</option><option value="formulas" ${findState.lookIn === 'formulas' ? 'selected' : ''}>Formulas</option><option value="notes" ${findState.lookIn === 'notes' ? 'selected' : ''}>Notes</option></select></div>
      <div class="row" style="flex-wrap:wrap"><button class="btn" data-a="all">Find All</button><button class="btn primary" data-a="next">Find Next</button><button class="btn ${replace ? '' : 'hidden'}" data-a="rep">Replace</button><button class="btn ${replace ? '' : 'hidden'}" data-a="repall">Replace All</button><button class="btn ${replace ? 'hidden' : ''}" data-a="showrep">Replace…</button></div>
      <div class="find-results hidden"></div><div class="dlg-note" id="fd-msg"></div>`;
    const read = () => { findState.q = $('#fd-q', body).value; findState.rep = $('#fd-rep', body).value; findState.matchCase = $('#fd-case', body).checked; findState.whole = $('#fd-whole', body).checked; findState.within = $('#fd-within', body).value; findState.lookIn = $('#fd-in', body).value; };
    const msg = (t) => $('#fd-msg', body).textContent = t;
    body.addEventListener('click', (e) => {
      const b = e.target.closest('[data-a]'); if (!b) { const it = e.target.closest('[data-res]'); if (it) { const r = findState.results[+it.dataset.res]; go(r); body.querySelectorAll('[data-res]').forEach(x => x.classList.toggle('on', x === it)); } return; }
      read(); const a = b.dataset.a;
      if (a === 'showrep') { $('#fd-rep-f', body).classList.remove('hidden'); body.querySelectorAll('[data-a="rep"],[data-a="repall"]').forEach(x => x.classList.remove('hidden')); b.classList.add('hidden'); return; }
      if (!findState.q && a !== 'repall') { msg('Type something to find.'); return; }
      if (a === 'next') { const r = Dlg.findNext(); msg(r ? '' : 'We couldn’t find what you were looking for.'); }
      if (a === 'all') { const res = Dlg.findAll(); const box = $('.find-results', body); box.classList.remove('hidden'); box.innerHTML = res.map((r, i) => `<div data-res="${i}"><span>${esc(r.s.name)}</span><span>${SU.addr(r.r, r.c)}</span><span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(r.text)}</span></div>`).join(''); msg(res.length + ' cell(s) found'); }
      if (a === 'rep') { const cur = findState.results[findState.idx]; const s2 = S(); if (cur && cur.s === s2 && cur.r === s2.active.r && cur.c === s2.active.c) replaceOne(cur); const r = Dlg.findNext(); msg(r ? '' : 'No more matches.'); }
      if (a === 'repall') { const res = Dlg.findAll(); let n = 0; wb.run('Replace all', () => { for (const r of res) if (replaceOne(r, true)) n++; }); msg(`All done. We made ${n} replacement(s).`); }
    });
    body.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); body.querySelector('[data-a="next"]').click(); } });
    await BTS.dialog({ title: replace ? 'Find and replace' : 'Find', body, buttons: [{ label: 'Close', value: null }] });
    read();
  };
  function matcher() {
    const q = findState.matchCase ? findState.q : findState.q.toLowerCase();
    const wild = /[*?]/.test(q); let re = null;
    if (wild) re = new RegExp((findState.whole ? '^' : '') + q.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + (findState.whole ? '$' : ''), findState.matchCase ? '' : 'i');
    return (t) => { if (re) return re.test(t); const x = findState.matchCase ? t : t.toLowerCase(); return findState.whole ? x === q : x.includes(q); };
  }
  function textFor(s, r, c) { if (findState.lookIn === 'formulas') return s.editText(r, c); if (findState.lookIn === 'notes') return s.comments.get(SU.key(r, c)) || ''; return s.displayText(r, c); }
  Dlg.findAll = () => {
    const m = matcher(); const out = []; const sheets = findState.within === 'book' ? wb.sheets : [S()];
    for (const s of sheets) {
      if (findState.lookIn === 'notes') { for (const [k, t] of s.comments) { const { r, c } = SU.unkey(k); if (m(t)) out.push({ s, r, c, text: t }); } continue; }
      const u = s.usedRange(); for (let r = 0; r <= u.r1; r++) for (let c = 0; c <= u.c1; c++) { if (s.isEmpty(r, c)) continue; const t = textFor(s, r, c); if (t && m(t)) out.push({ s, r, c, text: t }); }
    }
    findState.results = out; findState.idx = -1; return out;
  };
  Dlg.findNext = () => {
    const res = Dlg.findAll(); if (!res.length) return null;
    const s = S(); const a = s.active; const si = wb.activeIndex;
    let i = res.findIndex(r => wb.sheets.indexOf(r.s) > si || (r.s === s && (r.r > a.r || (r.r === a.r && r.c > a.c))));
    if (i < 0) i = 0; findState.idx = i; go(res[i]); return res[i];
  };
  function go(r) { const i = wb.sheets.indexOf(r.s); if (i !== wb.activeIndex) App.switchSheet(i); Grid.select(r.r, r.c); }
  function replaceOne(r, inTx) {
    const s = r.s; const raw = s.editText(r.r, r.c);
    const flags = findState.matchCase ? 'g' : 'gi'; const q = findState.q.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\\\*/g, '.*').replace(/\\\?/g, '.').replace(/\*/g, '.*').replace(/\?/g, '.');
    const re = new RegExp(findState.whole ? '^' + q + '$' : q, flags);
    const nt = raw.replace(re, findState.rep); if (nt === raw) return false;
    if (s.isFormula(r.r, r.c) && findState.lookIn !== 'formulas' && !nt.startsWith('=')) return false;
    const res = A.interpret(s, r.r, r.c, nt); if (res.error) return false;
    const fn = () => wb.setContents(s, [{ r: r.r, c: r.c, v: res.v }]);
    inTx ? fn() : wb.run('Replace', fn); return true;
  }
  Dlg.searchFromTitlebar = (q) => { if (!q) return; findState.q = q; findState.within = 'book'; findState.lookIn = 'values'; const r = Dlg.findNext(); if (!r) BTS.toast('No matches for “' + q + '”', { error: true }); };

  /* ---------- Go to ---------- */
  Dlg.goto = async () => {
    const names = wb.listNames();
    const body = `<div class="field"><label>Reference (e.g. B12, D4:F20, Sheet2!A1 or a defined name)</label><input id="gt-ref" list="gt-names"><datalist id="gt-names">${names.map(n => `<option value="${esc(n.name)}">`).join('')}</datalist></div>
      <div class="dlg-note">Go To Special:</div><div class="field-row" style="flex-wrap:wrap;gap:6px;margin-top:6px">${[['blanks', 'Blanks'], ['constants', 'Constants'], ['formulas', 'Formulas'], ['notes', 'Notes'], ['errors', 'Errors'], ['last', 'Last cell']].map(([k, l]) => `<button class="btn" data-sp="${k}" style="height:28px">${l}</button>`).join('')}</div>`;
    let special = null;
    const res = await BTS.dialog({ title: 'Go to', body, onOpen: (r) => r.addEventListener('click', e => { const b = e.target.closest('[data-sp]'); if (b) { special = b.dataset.sp; r.querySelector('footer .primary').click(); } }), buttons: [{ label: 'Cancel', value: null }, { label: 'Go', primary: true, value: (r) => r.querySelector('#gt-ref').value }] });
    if (special) return goSpecial(special);
    if (res) Dlg.gotoRef(res);
  };
  Dlg.gotoRef = (text) => {
    text = String(text).trim(); if (!text) return false;
    const ref = SU.parseRef(text);
    if (ref) { if (ref.sheet) { const sh = wb.sheetByName(ref.sheet); if (!sh) { BTS.toast('No sheet named ' + ref.sheet, { error: true }); return false; } App.switchSheet(wb.sheets.indexOf(sh)); } const g = ref.range; S().ensureSize(Math.min(g.r1, 1e6), Math.min(g.c1, 16000)); Grid.selectRange(g); return true; }
    const n = wb.listNames().find(x => x.name.toLowerCase() === text.toLowerCase());
    if (n) { const r2 = SU.parseRef(SU.fromHF(n.expr)); if (r2) { if (r2.sheet) App.switchSheet(wb.sheets.indexOf(wb.sheetByName(r2.sheet))); Grid.selectRange(r2.range); return true; } }
    const t = wb.sheets.flatMap(s => s.tables.map(t => ({ s, t }))).find(x => x.t.name.toLowerCase() === text.toLowerCase());
    if (t) { App.switchSheet(wb.sheets.indexOf(t.s)); Grid.selectRange(t.t.range); return true; }
    BTS.toast('Reference isn’t valid.', { error: true }); return false;
  };
  function goSpecial(kind) {
    const s = S(); const g = A.sels()[0]; const scope = (g.r0 === g.r1 && g.c0 === g.c1) ? { r0: 0, c0: 0, ...s.usedRange() } : A.effData(g);
    if (kind === 'last') { const u = s.usedRange(); Grid.select(Math.max(0, u.r1), Math.max(0, u.c1)); return; }
    const hits = [];
    for (let r = scope.r0; r <= scope.r1; r++) for (let c = scope.c0; c <= scope.c1; c++) {
      const e = s.isEmpty(r, c); let ok = false;
      if (kind === 'blanks') ok = e; else if (kind === 'constants') ok = !e && !s.isFormula(r, c); else if (kind === 'formulas') ok = s.isFormula(r, c); else if (kind === 'notes') ok = s.comments.has(SU.key(r, c)); else if (kind === 'errors') ok = SU.isErr(s.value(r, c));
      if (ok) hits.push({ r0: r, c0: c, r1: r, c1: c });
    }
    if (!hits.length) { BTS.toast('No cells were found.', { error: true }); return; }
    if (hits.length > 2000) hits.length = 2000;
    s.extraSel = hits.slice(0, -1); s.sel = hits[hits.length - 1]; s.active = { r: s.sel.r0, c: s.sel.c0 }; Grid.scrollIntoView(s.active.r, s.active.c); Grid.notify();
  }

  /* ---------- Custom sort ---------- */
  Dlg.sort = async () => {
    const s = S(); let g = A.dataRange(); const t = s.tableAt(s.active.r, s.active.c); if (t) g = t.total ? { ...t.range, r1: t.range.r1 - 1 } : t.range; else if (s.filter && SU.inRange(s.filter.range, s.active.r, s.active.c)) g = s.filter.range;
    let header = t ? true : A.hasHeader(s, g) || (s.filter && SU.inRange(s.filter.range, s.active.r, s.active.c));
    const colName = (c) => header && s.displayText(g.r0, c) ? s.displayText(g.r0, c) : 'Column ' + SU.colName(c);
    let levels = [{ c: s.active.c >= g.c0 && s.active.c <= g.c1 ? s.active.c : g.c0, desc: false, by: 'value' }];
    const body = document.createElement('div');
    const render = () => {
      body.innerHTML = `<div style="display:flex;gap:8px;align-items:center"><button class="btn" data-a="add">${icon('plus', 15)} Add level</button><button class="btn" data-a="del">${icon('trash-2', 15)} Delete level</button><button class="btn" data-a="copy">${icon('copy', 15)} Copy level</button><label class="check" style="margin-left:auto"><input type="checkbox" id="st-head" ${header ? 'checked' : ''}> My data has headers</label></div>
      <div class="sort-levels">${levels.map((l, i) => `<div class="lvl" data-i="${i}"><span style="font-size:12px;color:var(--ink-2)">${i ? 'Then by' : 'Sort by'}</span><select data-k="c">${range(g.c0, g.c1).map(c => `<option value="${c}" ${c === l.c ? 'selected' : ''}>${esc(colName(c))}</option>`).join('')}</select><select data-k="by"><option value="value">Cell values</option><option value="color" ${l.by === 'color' ? 'selected' : ''}>Cell color</option></select><select data-k="desc">${l.by === 'color' ? `<option value="0">On top</option>` : `<option value="0">${isNumCol(l.c) ? 'Smallest to largest' : 'A to Z'}</option><option value="1" ${l.desc ? 'selected' : ''}>${isNumCol(l.c) ? 'Largest to smallest' : 'Z to A'}</option><option value="custom" ${l.custom ? 'selected' : ''}>Custom list…</option>`}</select><span></span></div>`).join('')}</div>
      <p class="dlg-note">Sorting range: ${SU.rangeStr(g)}${header ? ' (first row kept as header)' : ''}</p>`;
    };
    const isNumCol = (c) => { for (let r = g.r0 + 1; r <= Math.min(g.r1, g.r0 + 30); r++) if (typeof s.value(r, c) === 'number') return true; return false; };
    const range = (a, b) => Array.from({ length: b - a + 1 }, (_, i) => a + i);
    let cur = 0;
    body.addEventListener('change', async (e) => { const lv = e.target.closest('.lvl'); if (e.target.id === 'st-head') { header = e.target.checked; render(); return; } if (!lv) return; const i = +lv.dataset.i; cur = i; const k = e.target.dataset.k; if (k === 'c') levels[i].c = +e.target.value; if (k === 'by') { levels[i].by = e.target.value; if (levels[i].by === 'color') { const cols = new Set(); for (let r = g.r0 + 1; r <= g.r1; r++) { const f = s.style(r, levels[i].c).fill; if (f) cols.add(f); } levels[i].color = [...cols][0] || null; } } if (k === 'desc') { if (e.target.value === 'custom') { const t2 = await BTS.prompt('Custom list', 'Enter the order, separated by commas', 'Low, Medium, High'); if (t2) levels[i].custom = t2.split(',').map(x => x.trim()); levels[i].desc = false; } else { levels[i].desc = e.target.value === '1'; delete levels[i].custom; } } render(); });
    body.addEventListener('click', (e) => { const lv = e.target.closest('.lvl'); if (lv) cur = +lv.dataset.i; const b = e.target.closest('[data-a]'); if (!b) return; if (b.dataset.a === 'add') levels.push({ c: g.c0, desc: false, by: 'value' }); if (b.dataset.a === 'del' && levels.length > 1) levels.splice(cur, 1); if (b.dataset.a === 'copy') levels.splice(cur + 1, 0, { ...levels[cur] }); render(); });
    render();
    const ok = await BTS.dialog({ title: 'Sort', body, width: 720, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: true }] });
    if (!ok) return;
    A.sortRange(s, g, levels, header);
    if (s.filter) wb.run('Filter', () => A.applyFilter(s));
  };

  /* ---------- Filter popover ---------- */
  Dlg.filterPopover = (hit) => {
    const s = S(); let f = s.filter;
    if (hit.table && (!f || f.tableId !== hit.table.id)) { wb.run('Filter', () => wb.setProp(s, 'filter', { range: hit.table.range, cols: {}, tableId: hit.table.id })); f = s.filter; }
    const c = hit.c; const cond = (f.cols && f.cols[c]) || {};
    const vals = new Map(); for (let r = f.range.r0 + 1; r <= f.range.r1; r++) { if (s.hiddenR.has(r)) continue; const v = s.value(r, c); const t = SU.valueText(v) === '' ? '' : s.displayText(r, c); if (!vals.has(t)) vals.set(t, v); }
    const items = [...vals.entries()].sort((a, b) => SU.cmp(a[1], b[1])).map(x => x[0]);
    const isNum = [...vals.values()].some(v => typeof v === 'number');
    const checked = new Set(cond.values || items);
    const el = document.createElement('div'); el.className = 'filter-pop';
    el.innerHTML = `<div class="menu-item" data-a="asc">${icon('arrow-down-a-z', 16)}<span>Sort ${isNum ? 'smallest to largest' : 'A to Z'}</span></div><div class="menu-item" data-a="desc">${icon('arrow-down-z-a', 16)}<span>Sort ${isNum ? 'largest to smallest' : 'Z to A'}</span></div>
      <div class="menu-sep"></div><div class="menu-item ${cond.values || cond.cond ? '' : 'disabled'}" data-a="clear">${icon('funnel-x', 16)}<span>Clear filter</span></div>
      <div class="menu-item" data-a="cond">${icon('list-filter', 16)}<span>${isNum ? 'Number' : 'Text'} filters…</span></div>
      <input class="fp-search" placeholder="Search">
      <div class="fp-list"><label><input type="checkbox" data-all> (Select all)</label>${items.map((t, i) => `<label title="${esc(t)}"><input type="checkbox" data-v="${i}" ${checked.has(t) ? 'checked' : ''}> ${esc(t === '' ? '(Blanks)' : t)}</label>`).join('')}</div>
      <div class="fp-foot"><button class="btn" data-a="cancel">Cancel</button><button class="btn primary" data-a="ok">OK</button></div>`;
    const all = $('[data-all]', el); const boxes = [...el.querySelectorAll('[data-v]')];
    const syncAll = () => { const vis = boxes.filter(b => b.parentElement.style.display !== 'none'); all.checked = vis.every(b => b.checked); all.indeterminate = !all.checked && vis.some(b => b.checked); };
    syncAll();
    all.addEventListener('change', () => boxes.forEach(b => { if (b.parentElement.style.display !== 'none') b.checked = all.checked; }));
    el.addEventListener('change', (e) => { if (e.target.dataset.v !== undefined) syncAll(); });
    $('.fp-search', el).addEventListener('input', (e) => { const q = e.target.value.toLowerCase(); boxes.forEach(b => { const t = items[+b.dataset.v].toLowerCase(); b.parentElement.style.display = t.includes(q) ? '' : 'none'; if (q) b.checked = t.includes(q); }); syncAll(); });
    el.addEventListener('mousedown', (e) => { if (e.target.tagName !== 'INPUT') e.preventDefault(); });
    el.addEventListener('click', async (e) => {
      const a = e.target.closest('[data-a]'); if (!a) return; const act = a.dataset.a;
      if (act === 'cancel') { BTS.closePopover(); return; }
      if (act === 'asc' || act === 'desc') { BTS.closePopover(); A.sortRange(s, f.range, [{ c, desc: act === 'desc' }], true); wb.run('Filter', () => A.applyFilter(s)); return; }
      if (act === 'clear') { BTS.closePopover(); A.setFilterCol(c, null); return; }
      if (act === 'cond') { BTS.closePopover(); const cd = await condDialog(isNum, cond.cond); if (cd) A.setFilterCol(c, { cond: cd }); return; }
      if (act === 'ok') { const sel = boxes.filter(b => b.checked).map(b => items[+b.dataset.v]); BTS.closePopover(); A.setFilterCol(c, sel.length === items.length ? null : { values: sel }); }
    });
    const rc = Grid.cellRect(f.range.r0, c); const hr = document.getElementById('grid-host').getBoundingClientRect();
    BTS.popover({ x: hr.left + rc.x + rc.w - 270 > hr.left ? hr.left + rc.x + rc.w - 270 : hr.left + rc.x, y: hr.top + rc.y + rc.h + 2 }, el);
  };
  async function condDialog(isNum, cur = {}) {
    const ops = isNum ? [['eq', 'equals'], ['ne', 'does not equal'], ['gt', 'is greater than'], ['ge', 'is greater than or equal to'], ['lt', 'is less than'], ['le', 'is less than or equal to'], ['between', 'is between'], ['top', 'Top N'], ['bottom', 'Bottom N'], ['above', 'Above average'], ['below', 'Below average']]
      : [['eq', 'equals'], ['ne', 'does not equal'], ['begins', 'begins with'], ['ends', 'ends with'], ['contains', 'contains'], ['notcontains', 'does not contain']];
    const body = `<div class="dlg-grid"><label>Show rows where</label><select id="cd-op">${ops.map(([k, l]) => `<option value="${k}" ${cur.op === k ? 'selected' : ''}>${l}</option>`).join('')}</select><label>Value</label><input id="cd-v1" value="${esc(cur.v1 ?? '')}"><label>And (for between)</label><input id="cd-v2" value="${esc(cur.v2 ?? '')}"></div>`;
    return BTS.dialog({ title: 'Custom filter', body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ op: r.querySelector('#cd-op').value, v1: r.querySelector('#cd-v1').value, v2: r.querySelector('#cd-v2').value }) }] });
  }

  /* ---------- Data validation ---------- */
  Dlg.validation = async () => {
    const s = S(); const g = A.eff(s.sel); const cur = s.validationAt(s.active.r, s.active.c) || { type: 'any', op: 'between', ignoreBlank: true, errorStyle: 'stop' };
    const body = document.createElement('div');
    body.innerHTML = `<div class="dlg-grid">
      <label>Allow</label><select id="dv-type"><option value="any">Any value</option><option value="whole">Whole number</option><option value="decimal">Decimal</option><option value="list">List</option><option value="date">Date</option><option value="textLength">Text length</option></select>
      <label class="dv-num">Data</label><select id="dv-op" class="dv-num"><option value="between">between</option><option value="notBetween">not between</option><option value="eq">equal to</option><option value="ne">not equal to</option><option value="gt">greater than</option><option value="lt">less than</option><option value="ge">greater than or equal to</option><option value="le">less than or equal to</option></select>
      <label class="dv-num">Minimum / value</label><input id="dv-v1" class="dv-num" value="${esc(cur.v1 ?? '')}">
      <label class="dv-num dv-max">Maximum</label><input id="dv-v2" class="dv-num dv-max" value="${esc(cur.v2 ?? '')}">
      <label class="dv-list">Source</label><input id="dv-src" class="dv-list" placeholder="Yes, No, Maybe   or   =$A$1:$A$5" value="${esc(cur.source ?? '')}">
      <label>Input message</label><input id="dv-prompt" value="${esc(cur.prompt ?? '')}" placeholder="Optional hint shown when the cell is selected">
      <label>Error style</label><select id="dv-style"><option value="stop">Stop</option><option value="warning">Warning</option><option value="information">Information</option></select>
      <label>Error title</label><input id="dv-etitle" value="${esc(cur.errorTitle ?? '')}">
      <label>Error message</label><input id="dv-emsg" value="${esc(cur.errorMsg ?? '')}" placeholder="e.g. Enter a number between 0 and 100">
      </div><label class="check"><input type="checkbox" id="dv-blank" ${cur.ignoreBlank !== false ? 'checked' : ''}> Ignore blank</label><p class="dlg-note">Applies to ${SU.rangeStr(g)}.</p>`;
    $('#dv-type', body).value = cur.type; $('#dv-op', body).value = cur.op || 'between'; $('#dv-style', body).value = cur.errorStyle || 'stop';
    const sync = () => { const t = $('#dv-type', body).value; const op = $('#dv-op', body).value; body.querySelectorAll('.dv-num').forEach(x => x.style.display = ['whole', 'decimal', 'date', 'textLength'].includes(t) ? '' : 'none'); body.querySelectorAll('.dv-max').forEach(x => x.style.display = ['whole', 'decimal', 'date', 'textLength'].includes(t) && (op === 'between' || op === 'notBetween') ? '' : 'none'); body.querySelectorAll('.dv-list').forEach(x => x.style.display = t === 'list' ? '' : 'none'); };
    body.addEventListener('change', sync); sync();
    const res = await BTS.dialog({ title: 'Data validation', body, buttons: [{ label: 'Clear all', value: 'clear' }, { label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => { const q = (x) => r.querySelector(x).value; const pd = (x) => { const p = SU.parseInput(x); return typeof p.value === 'number' ? p.value : x; }; return { type: q('#dv-type'), op: q('#dv-op'), v1: pd(q('#dv-v1')), v2: pd(q('#dv-v2')), source: q('#dv-src'), prompt: q('#dv-prompt'), errorStyle: q('#dv-style'), errorTitle: q('#dv-etitle'), errorMsg: q('#dv-emsg'), ignoreBlank: r.querySelector('#dv-blank').checked }; } }] });
    if (!res) return;
    const rest = s.validations.filter(v => !SU.intersects(v.range, g) || !(v.range.r0 >= g.r0 && v.range.r1 <= g.r1 && v.range.c0 >= g.c0 && v.range.c1 <= g.c1));
    // trim overlapping validations
    const trimmed = []; for (const v of rest) { if (!SU.intersects(v.range, g)) { trimmed.push(v); continue; } /* split by rows above/below */ const r = v.range; if (r.r0 < g.r0) trimmed.push({ ...v, range: { ...r, r1: g.r0 - 1 } }); if (r.r1 > g.r1) trimmed.push({ ...v, range: { ...r, r0: g.r1 + 1 } }); if (r.c0 < g.c0) trimmed.push({ ...v, range: { r0: Math.max(r.r0, g.r0), r1: Math.min(r.r1, g.r1), c0: r.c0, c1: g.c0 - 1 } }); if (r.c1 > g.c1) trimmed.push({ ...v, range: { r0: Math.max(r.r0, g.r0), r1: Math.min(r.r1, g.r1), c0: g.c1 + 1, c1: r.c1 } }); }
    const list = res === 'clear' || res.type === 'any' ? trimmed : [...trimmed, { ...res, range: g }];
    if (res !== 'clear' && res.type === 'list' && !String(res.source).trim()) { BTS.toast('Enter the list items for the source.', { error: true }); return; }
    wb.run('Data validation', () => wb.setProp(s, 'validations', list));
    App.refreshValidationUI();
  };

  /* ---------- Conditional formatting ---------- */
  Dlg.cfRule = async (type) => {
    const titles = { gt: 'Greater than', lt: 'Less than', between: 'Between', eq: 'Equal to', text: 'Text that contains', dup: 'Duplicate values', top: 'Top 10 items', bottom: 'Bottom 10 items', aboveavg: 'Above average', belowavg: 'Below average', formula: 'Use a formula to format cells' };
    const s = S(); const a = s.active; const sample = s.value(a.r, a.c);
    const needs1 = ['gt', 'lt', 'between', 'eq', 'text', 'top', 'bottom', 'formula'].includes(type);
    const label = { gt: 'Format cells that are GREATER THAN:', lt: 'Format cells that are LESS THAN:', between: 'Format cells that are BETWEEN:', eq: 'Format cells that are EQUAL TO:', text: 'Format cells that contain the text:', top: 'Format cells that rank in the TOP:', bottom: 'Format cells that rank in the BOTTOM:', formula: 'Format values where this formula is true (relative to the top-left cell):' }[type] || 'Format cells with:';
    const def = type === 'top' || type === 'bottom' ? 10 : type === 'formula' ? '=' + SU.addr(A.eff(s.sel).r0, A.eff(s.sel).c0) + '>0' : typeof sample === 'number' ? sample : '';
    const body = `<p style="margin:0 0 8px">${label}</p><div class="field-row">${needs1 ? `<div class="field"><input id="cf-v1" value="${esc(def)}"></div>` : ''}${type === 'between' ? `<div class="field" style="flex:0 0 auto;align-self:center">and</div><div class="field"><input id="cf-v2" value=""></div>` : ''}${type === 'top' || type === 'bottom' ? `<label class="check"><input type="checkbox" id="cf-pct"> %</label>` : ''}</div>
      <div class="field"><label>with</label><select id="cf-fmt">${A.CF_FORMATS.map((f, i) => `<option value="${i}">${f.label}</option>`).join('')}<option value="custom">Custom fill colour…</option></select></div>`;
    const res = await BTS.dialog({ title: titles[type], body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ v1: r.querySelector('#cf-v1')?.value, v2: r.querySelector('#cf-v2')?.value, pct: r.querySelector('#cf-pct')?.checked, fmt: r.querySelector('#cf-fmt').value }) }] });
    if (!res) return;
    let fmt = A.CF_FORMATS[+res.fmt]?.fmt;
    if (res.fmt === 'custom') { const c = await BTS.colorPicker(document.querySelector('[data-cmd="cf-menu"]')); if (!c || c === 'none' || c === 'auto') return; fmt = { fill: c }; }
    const rule = { type, fmt };
    if (type === 'formula') rule.formula = res.v1.startsWith('=') ? res.v1 : '=' + res.v1; else { rule.v1 = res.v1; rule.v2 = res.v2; }
    if (type === 'top' || type === 'bottom') { rule.n = +res.v1 || 10; rule.pct = !!res.pct; }
    A.addCF(rule);
  };
  Dlg.cfManager = async () => {
    const s = S();
    const body = document.createElement('div');
    const desc = (r) => ({ gt: 'Cell value > ' + r.v1, lt: 'Cell value < ' + r.v1, between: `Cell value between ${r.v1} and ${r.v2}`, eq: 'Cell value = ' + r.v1, text: 'Contains “' + r.v1 + '”', dup: 'Duplicate values', unique: 'Unique values', top: `Top ${r.n}${r.pct ? '%' : ''}`, bottom: `Bottom ${r.n}${r.pct ? '%' : ''}`, aboveavg: 'Above average', belowavg: 'Below average', bar: 'Data bar', scale: 'Color scale', formula: 'Formula: ' + r.formula }[r.type] || r.type);
    let list = s.cf.slice(); let sel = 0;
    const render = () => { body.innerHTML = `<div class="nm-wrap"><table class="nm-table"><thead><tr><th>Rule (applied in order shown)</th><th>Format</th><th>Applies to</th></tr></thead><tbody>${list.map((r, i) => `<tr data-i="${i}" class="${i === sel ? 'on' : ''}"><td>${esc(desc(r))}</td><td><span style="display:inline-block;padding:1px 10px;border-radius:3px;background:${(r.fmt && r.fmt.fill) || (r.type === 'bar' ? r.color || '#5B9BD5' : r.type === 'scale' ? 'linear-gradient(90deg,#F8696B,#FFEB84,#63BE7B)' : '#fff')};color:${(r.fmt && r.fmt.color) || '#000'}">AaBbCc</span></td><td>${SU.rangeStr(r.range)}</td></tr>`).join('') || '<tr><td colspan="3" class="empty-hint">No rules on this sheet.</td></tr>'}</tbody></table></div><div style="display:flex;gap:6px;margin-top:8px"><button class="btn" data-a="del">${icon('trash-2', 15)} Delete rule</button><button class="btn" data-a="up">${icon('arrow-up', 15)}</button><button class="btn" data-a="down">${icon('arrow-down', 15)}</button></div>`; };
    body.addEventListener('click', e => { const tr = e.target.closest('tr[data-i]'); if (tr) { sel = +tr.dataset.i; render(); } const b = e.target.closest('[data-a]'); if (!b) return; if (b.dataset.a === 'del') list.splice(sel, 1); if (b.dataset.a === 'up' && sel > 0) { [list[sel - 1], list[sel]] = [list[sel], list[sel - 1]]; sel--; } if (b.dataset.a === 'down' && sel < list.length - 1) { [list[sel + 1], list[sel]] = [list[sel], list[sel + 1]]; sel++; } sel = Math.max(0, Math.min(sel, list.length - 1)); render(); });
    render();
    const ok = await BTS.dialog({ title: 'Conditional formatting rules', body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: true }] });
    if (ok) wb.run('Manage rules', () => wb.setProp(s, 'cf', list));
  };

  /* ---------- Names ---------- */
  Dlg.defineName = async (initial) => {
    const s = S(); const g = A.eff(s.sel);
    let guess = ''; const above = g.r0 > 0 ? s.value(g.r0 - 1, g.c0) : null; const top = s.value(g.r0, g.c0);
    if (typeof top === 'string' && g.r1 > g.r0) guess = top; else if (typeof above === 'string') guess = above; else { const left = g.c0 > 0 ? s.value(g.r0, g.c0 - 1) : null; if (typeof left === 'string') guess = left; }
    guess = String(guess).trim().replace(/[^A-Za-z0-9_.]+/g, '_').replace(/^[^A-Za-z_]+/, '');
    const res = await BTS.dialog({ title: 'New name', body: `<div class="field"><label>Name</label><input id="dn-name" value="${esc(initial?.name || guess)}"></div><div class="field"><label>Refers to</label><input id="dn-ref" value="${esc(initial?.expr || ('=' + SU.quoteSheet(s.name) + '!' + SU.absRangeStr(g)))}"></div><p class="dlg-note">Use the name in formulas, e.g. =SUM(${esc(guess || 'Marks')}).</p>`, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ name: r.querySelector('#dn-name').value.trim(), expr: r.querySelector('#dn-ref').value.trim() }) }] });
    if (!res || !res.name) return;
    try { wb.run('Define name', () => initial ? (initial.name === res.name ? wb.changeName(res.name, res.expr) : (wb.removeName(initial.name), wb.addName(res.name, res.expr))) : wb.addName(res.name, res.expr)); App.refreshFormulaBar(); }
    catch (e) { BTS.toast(e.message, { error: true, ms: 4000 }); }
  };
  Dlg.nameManager = async () => {
    const body = document.createElement('div'); let sel = 0;
    const render = () => { const names = wb.listNames(); body.innerHTML = `<div style="display:flex;gap:6px;margin-bottom:8px"><button class="btn" data-a="new">${icon('plus', 15)} New…</button><button class="btn" data-a="edit">${icon('pencil', 15)} Edit…</button><button class="btn" data-a="del">${icon('trash-2', 15)} Delete</button></div><div class="nm-wrap"><table class="nm-table"><thead><tr><th>Name</th><th>Value</th><th>Refers to</th></tr></thead><tbody>${names.map((n, i) => { let v = ''; try { const x = wb.hf.getNamedExpressionValue(n.name); v = x && x.data ? '{' + x.data.slice(0, 3).map(r => r.slice(0, 3).map(SU.valueText).join(', ')).join('; ') + (x.data.length > 3 ? '…' : '') + '}' : SU.valueText(x); } catch { } return `<tr data-i="${i}" class="${i === sel ? 'on' : ''}"><td><b>${esc(n.name)}</b></td><td>${esc(v)}</td><td>${esc(SU.fromHF(n.expr))}</td></tr>`; }).join('') || '<tr><td colspan="3" class="empty-hint">No names defined yet. Select cells and click New… to name them.</td></tr>'}</tbody></table></div>`; };
    body.addEventListener('click', async e => { const tr = e.target.closest('tr[data-i]'); if (tr) { sel = +tr.dataset.i; render(); } const b = e.target.closest('[data-a]'); if (!b) return; const names = wb.listNames(); if (b.dataset.a === 'new') { await Dlg.defineName(); render(); } if (b.dataset.a === 'edit' && names[sel]) { await Dlg.defineName({ name: names[sel].name, expr: SU.fromHF(names[sel].expr) }); render(); } if (b.dataset.a === 'del' && names[sel]) { wb.run('Delete name', () => wb.removeName(names[sel].name)); sel = 0; render(); } });
    render();
    await BTS.dialog({ title: 'Name manager', body, buttons: [{ label: 'Close', primary: true, value: true }] });
  };
  Dlg.namesFromSelection = () => {
    const s = S(); const g = A.effData(s.sel); if (g.r1 <= g.r0) { BTS.toast('Select a range with labels in the top row.', { error: true }); return; }
    let n = 0; const errs = [];
    wb.run('Create names', () => { for (let c = g.c0; c <= g.c1; c++) { const t = String(s.displayText(g.r0, c)).trim().replace(/[^A-Za-z0-9_.]+/g, '_').replace(/^[^A-Za-z_]+/, ''); if (!t) continue; try { wb.addName(t, '=' + SU.quoteSheet(s.name) + '!' + SU.absRangeStr({ r0: g.r0 + 1, r1: g.r1, c0: c, c1: c })); n++; } catch (e) { errs.push(t); } } });
    BTS.toast(`Created ${n} name(s)` + (errs.length ? `; skipped ${errs.join(', ')}` : ''));
  };

  /* ---------- Text to columns ---------- */
  Dlg.textToColumns = async () => {
    const s = S(); const g = A.effData(s.sel);
    if (g.c1 !== g.c0) { BTS.toast('Select cells in one column only.', { error: true }); return; }
    const body = document.createElement('div');
    body.innerHTML = `<p style="margin:0 0 6px">Split the text in ${SU.rangeStr(g)} into columns using:</p>
      <div style="display:flex;gap:14px;flex-wrap:wrap"><label class="check"><input type="checkbox" id="tc-comma" checked> Comma</label><label class="check"><input type="checkbox" id="tc-tab"> Tab</label><label class="check"><input type="checkbox" id="tc-semi"> Semicolon</label><label class="check"><input type="checkbox" id="tc-space"> Space</label><label class="check">Other <input id="tc-other" maxlength="1" style="width:34px;height:26px;border:1px solid var(--line-2);border-radius:5px;padding:0 6px"></label></div>
      <div style="display:flex;gap:14px"><label class="check"><input type="checkbox" id="tc-cons"> Treat consecutive delimiters as one</label><label class="check"><input type="checkbox" id="tc-quote" checked> Text qualifier "</label></div>
      <div class="field" style="margin-top:6px"><label>Destination</label><input id="tc-dest" value="${SU.addr(g.r0, g.c0)}"></div>
      <div class="dlg-note">Preview</div><div class="preview-wrap"><table class="preview-tbl"></table></div>`;
    const opts = () => ({ comma: $('#tc-comma', body).checked, tab: $('#tc-tab', body).checked, semi: $('#tc-semi', body).checked, space: $('#tc-space', body).checked, other: $('#tc-other', body).value, consecutive: $('#tc-cons', body).checked, quote: $('#tc-quote', body).checked });
    const prev = () => { const o = opts(); const rows = []; for (let r = g.r0; r <= Math.min(g.r1, g.r0 + 8); r++) rows.push(A.splitText(SU.valueText(s.value(r, g.c0)), o)); $('.preview-tbl', body).innerHTML = rows.map(r => '<tr>' + r.map(x => `<td>${esc(x)}</td>`).join('') + '</tr>').join(''); };
    body.addEventListener('input', prev); body.addEventListener('change', prev); prev();
    const res = await BTS.dialog({ title: 'Text to columns', body, buttons: [{ label: 'Cancel', value: null }, { label: 'Finish', primary: true, value: () => ({ ...opts(), destStr: $('#tc-dest', body).value }) }] });
    if (!res) return; const d = SU.parseAddr(res.destStr) || { r: g.r0, c: g.c0 };
    // warn about overwrite
    const n = A.textToColumns(g, { ...res, dest: d });
  };

  /* ---------- Remove duplicates ---------- */
  Dlg.removeDuplicates = async () => {
    const s = S(); const g = A.dataRange(); let header = A.hasHeader(s, g) || !!s.tableAt(g.r0, g.c0);
    const body = document.createElement('div');
    const render = () => { body.innerHTML = `<p style="margin:0 0 6px">Select one or more columns that contain duplicates (range ${SU.rangeStr(g)}).</p><label class="check"><input type="checkbox" id="rd-head" ${header ? 'checked' : ''}> My data has headers</label><div class="preview-wrap" style="max-height:220px">${Array.from({ length: g.c1 - g.c0 + 1 }, (_, i) => g.c0 + i).map(c => `<label class="check"><input type="checkbox" data-c="${c}" checked> ${esc(header ? s.displayText(g.r0, c) || 'Column ' + SU.colName(c) : 'Column ' + SU.colName(c))}</label>`).join('')}</div>`; };
    body.addEventListener('change', e => { if (e.target.id === 'rd-head') { header = e.target.checked; render(); } }); render();
    const res = await BTS.dialog({ title: 'Remove duplicates', body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => [...r.querySelectorAll('[data-c]')].filter(x => x.checked).map(x => +x.dataset.c) }] });
    if (!res || !res.length) return;
    const r = A.removeDuplicates(g, res, header);
    BTS.dialog({ title: 'Remove duplicates', body: `<p style="margin:4px 0">${r.removed ? `${r.removed} duplicate value(s) found and removed; ${r.kept} unique value(s) remain.` : 'No duplicate values found.'}</p>`, buttons: [{ label: 'OK', primary: true, value: true }] });
  };

  /* ---------- Symbols ---------- */
  Dlg.symbols = async () => {
    const sets = { 'Common': '₹$€£¥¢©®™°±×÷≠≈≤≥∞√∑∏∆π‰§¶•…–—«»“”‘’', 'Arrows': '←↑→↓↔↕⇐⇒⇔↩↪⤴⤵▲▼◀▶', 'Math': '∀∂∃∅∇∈∉∋∏∑−∓∗∘∙√∛∝∞∠∧∨∩∪∫∴∵≅≈≠≡≤≥⊂⊃⊆⊇⊕⊗⊥⋅½⅓⅔¼¾⅛', 'Greek': 'αβγδεζηθικλμνξοπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ', 'Shapes & ticks': '✓✔✗✘☐☑☒★☆●○◆◇■□▪▫◉◎♠♣♥♦☺☹☀☁☂☎✉✈', 'Indian': '₹ॐ।॥०१२३४५६७८९अआइईउऊएऐओऔकखगघचछजझटठडढणतथदधनपफबभमयरलवशषसह' };
    let cat = 'Common', pick = null;
    const body = document.createElement('div');
    const render = () => { body.innerHTML = `<div style="display:flex;gap:6px;margin-bottom:8px;flex-wrap:wrap">${Object.keys(sets).map(k => `<button class="btn ${k === cat ? 'primary' : ''}" data-cat="${k}" style="height:28px">${k}</button>`).join('')}</div><div class="sym-grid">${[...sets[cat]].filter(x => x.trim()).map(ch => `<button data-ch="${ch}" class="${ch === pick ? 'on' : ''}">${ch}</button>`).join('')}</div>`; };
    body.addEventListener('click', e => { const c = e.target.closest('[data-cat]'); if (c) { cat = c.dataset.cat; render(); } const b = e.target.closest('[data-ch]'); if (b) { pick = b.dataset.ch; render(); } });
    body.addEventListener('dblclick', e => { const b = e.target.closest('[data-ch]'); if (b) { pick = b.dataset.ch; document.querySelector('.modal footer .primary').click(); } });
    render();
    const ch = await BTS.dialog({ title: 'Symbol', body, buttons: [{ label: 'Cancel', value: null }, { label: 'Insert', primary: true, value: () => pick }] });
    if (!ch) return;
    if (Ed.isEditing()) Ed.insertAtCaret(ch); else { if (A.protectedCheck()) return; const s = S(); Ed.start(s.editText(s.active.r, s.active.c) + ch, 'edit'); }
  };

  /* ---------- Hyperlink & notes ---------- */
  Dlg.hyperlink = async () => {
    const s = S(); const { r, c } = s.active; const k = SU.key(r, c); const cur = s.links.get(k) || '';
    const text = s.displayText(r, c);
    const res = await BTS.dialog({ title: cur ? 'Edit link' : 'Insert link', body: `<div class="field"><label>Text to display</label><input id="hl-text" value="${esc(text)}"></div><div class="field"><label>Address (web page, email, or a place in this workbook like Sheet2!A1)</label><input id="hl-url" value="${esc(cur)}" placeholder="https://"></div>`, buttons: [...(cur ? [{ label: 'Remove link', value: 'remove' }] : []), { label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (rt) => ({ text: rt.querySelector('#hl-text').value, url: rt.querySelector('#hl-url').value.trim() }) }] });
    if (!res) return;
    wb.run('Link', () => {
      const m = new Map(s.links);
      if (res === 'remove' || !res.url) m.delete(k);
      else { let url = res.url; if (/^www\./i.test(url)) url = 'https://' + url; else if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(url)) url = 'mailto:' + url; m.set(k, url); if (res.text !== text || s.isEmpty(r, c)) wb.setContents(s, [{ r, c, v: res.text || url }]); }
      wb.setProp(s, 'links', m);
    });
  };
  Dlg.followLink = (url) => {
    if (!url) return;
    if (/^#?'?[^!]+'?!/.test(url) || SU.parseRef(url.replace(/^#/, ''))) { Dlg.gotoRef(url.replace(/^#/, '')); return; }
    if (window.bts) bts.openExternal(url); else window.open(url);
  };
  Dlg.note = async () => {
    const s = S(); const { r, c } = s.active; const k = SU.key(r, c); const cur = s.comments.get(k) || '';
    const res = await BTS.dialog({ title: cur ? 'Edit note' : 'New note', body: `<div class="field"><label>Note for ${SU.addr(r, c)}</label><textarea id="nt-text" style="height:120px;width:400px">${esc(cur)}</textarea></div>`, buttons: [...(cur ? [{ label: 'Delete', value: 'del', danger: true }] : []), { label: 'Cancel', value: null }, { label: 'Save', primary: true, value: (rt) => rt.querySelector('#nt-text').value }] });
    if (res === null) return;
    wb.run('Note', () => { const m = new Map(s.comments); if (res === 'del' || !res.trim()) m.delete(k); else m.set(k, res); wb.setProp(s, 'comments', m); });
  };

  /* ---------- sizes ---------- */
  Dlg.size = async (axis) => {
    const s = S(); const a = s.active; const cur = axis === 'col' ? s.rawColWidth(a.c) : s.rawRowHeight(a.r);
    const res = await BTS.prompt(axis === 'col' ? 'Column width' : 'Row height', axis === 'col' ? 'Column width (characters)' : 'Row height (points)', axis === 'col' ? ((cur - 5) / 7).toFixed(2) : (cur * 0.75).toFixed(2));
    if (res === null) return; const n = parseFloat(res); if (isNaN(n) || n < 0 || n > (axis === 'col' ? 255 : 409)) { BTS.toast('Enter a number between 0 and ' + (axis === 'col' ? 255 : 409) + '.', { error: true }); return; }
    const px = axis === 'col' ? Math.round(n * 7 + 5) : Math.round(n / 0.75);
    const g = A.eff(s.sel); const m = new Map(axis === 'col' ? s.colW : s.rowH); const hid = new Set(axis === 'col' ? s.hiddenC : s.hiddenR);
    const [lo, hi] = axis === 'col' ? [g.c0, g.c1] : [g.r0, g.r1];
    for (let i = lo; i <= hi; i++) { if (n === 0) hid.add(i); else { m.set(i, px); hid.delete(i); } }
    wb.run('Size', () => { wb.setProp(s, axis === 'col' ? 'colW' : 'rowH', m); wb.setProp(s, axis === 'col' ? 'hiddenC' : 'hiddenR', hid); });
    Grid.updateSpacer();
  };

  /* ---------- paste special ---------- */
  Dlg.pasteSpecial = async () => {
    const body = `<div style="display:grid;grid-template-columns:1fr 1fr;gap:4px 24px"><div><div class="dlg-note">Paste</div>${[['all', 'All'], ['formulas', 'Formulas'], ['values', 'Values'], ['formats', 'Formats'], ['valuesnf', 'Values and number formats'], ['nocopyborders', 'All except borders'], ['colwidths', 'Column widths']].map(([k, l], i) => `<label class="check"><input type="radio" name="ps" value="${k}" ${i === 0 ? 'checked' : ''}> ${l}</label>`).join('')}</div>
      <div><div class="dlg-note">Operation</div>${[['', 'None'], ['add', 'Add'], ['sub', 'Subtract'], ['mul', 'Multiply'], ['div', 'Divide']].map(([k, l], i) => `<label class="check"><input type="radio" name="op" value="${k}" ${i === 0 ? 'checked' : ''}> ${l}</label>`).join('')}</div></div>
      <div style="display:flex;gap:18px;margin-top:8px"><label class="check"><input type="checkbox" id="ps-skip"> Skip blanks</label><label class="check"><input type="checkbox" id="ps-tr"> Transpose</label></div>`;
    const res = await BTS.dialog({ title: 'Paste special', body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ mode: r.querySelector('[name=ps]:checked').value, op: r.querySelector('[name=op]:checked').value, skipBlanks: r.querySelector('#ps-skip').checked, transpose: r.querySelector('#ps-tr').checked }) }] });
    if (!res) return;
    if (!A.clip) { if (res.mode === 'values' || res.transpose) { A.pasteCmd(res.transpose ? 'transpose' : 'values'); return; } BTS.toast('Copy some cells in BTS Sheets first.', { error: true }); return; }
    A.pasteInternal(res.transpose && res.mode === 'all' ? 'transpose' : res.mode, { op: res.op || null, skipBlanks: res.skipBlanks, transpose: res.transpose });
  };

  /* ---------- series ---------- */
  Dlg.series = async () => {
    const s = S(); const g = s.sel; const dir = g.c1 - g.c0 > g.r1 - g.r0 ? 'rows' : 'cols';
    const body = `<div class="dlg-grid"><label>Series in</label><select id="se-dir"><option value="cols" ${dir === 'cols' ? 'selected' : ''}>Columns</option><option value="rows" ${dir === 'rows' ? 'selected' : ''}>Rows</option></select><label>Type</label><select id="se-type"><option value="linear">Linear</option><option value="growth">Growth</option><option value="date">Date</option></select><label>Date unit</label><select id="se-unit"><option value="day">Day</option><option value="weekday">Weekday</option><option value="month">Month</option><option value="year">Year</option></select><label>Step value</label><input id="se-step" value="1"><label>Stop value</label><input id="se-stop" placeholder="(optional when a range is selected)"></div>`;
    const res = await BTS.dialog({ title: 'Series', body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ dir: r.querySelector('#se-dir').value, type: r.querySelector('#se-type').value, unit: r.querySelector('#se-unit').value, step: +r.querySelector('#se-step').value || 1, stop: r.querySelector('#se-stop').value }) }] });
    if (res) A.fillSeries(res);
  };

  /* ---------- Insert / delete cells ---------- */
  Dlg.insertCells = async () => {
    const res = await BTS.dialog({ title: 'Insert', body: ['right', 'down', 'row', 'col'].map((k, i) => `<label class="check"><input type="radio" name="ic" value="${k}" ${i === 1 ? 'checked' : ''}> ${['Shift cells right', 'Shift cells down', 'Entire row', 'Entire column'][i]}</label>`).join(''), buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => r.querySelector('[name=ic]:checked').value }] });
    if (res === 'row') A.insertRows(); else if (res === 'col') A.insertCols(); else if (res) A.insertCells(res);
  };
  Dlg.deleteCells = async () => {
    const s = S(); if (Grid.isWholeRows(s.sel)) return A.deleteRows(); if (Grid.isWholeCols(s.sel)) return A.deleteCols();
    const res = await BTS.dialog({ title: 'Delete', body: ['left', 'up', 'row', 'col'].map((k, i) => `<label class="check"><input type="radio" name="dc" value="${k}" ${i === 1 ? 'checked' : ''}> ${['Shift cells left', 'Shift cells up', 'Entire row', 'Entire column'][i]}</label>`).join(''), buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => r.querySelector('[name=dc]:checked').value }] });
    if (res === 'row') A.deleteRows(); else if (res === 'col') A.deleteCols(); else if (res) A.deleteCells(res);
  };

  /* ---------- Format cells (number tab + alignment + protection) ---------- */
  Dlg.formatCells = async () => {
    const s = S(); const a = s.active; const st = s.style(a.r, a.c); const v = s.value(a.r, a.c);
    const body = document.createElement('div');
    body.innerHTML = `<div class="dlg-grid" style="grid-template-columns:130px 1fr">
      <label>Category</label><select id="fc-cat">${SU.FORMATS.map(f => `<option value="${f.nf.replace(/"/g, '&quot;')}">${f.label}</option>`).join('')}<option value="__custom">Custom</option></select>
      <label>Format code</label><input id="fc-nf" value="${esc(st.nf || 'General')}">
      <label>Sample</label><div id="fc-sample" style="font-family:Carlito,Calibri,sans-serif;font-size:14px;padding:4px 0"></div>
      <label>Horizontal</label><select id="fc-h"><option value="">General</option><option value="left">Left</option><option value="center">Center</option><option value="right">Right</option><option value="fill">Fill</option></select>
      <label>Vertical</label><select id="fc-v"><option value="bottom">Bottom</option><option value="middle">Middle</option><option value="top">Top</option></select>
      <label>Indent</label><input id="fc-ind" type="number" min="0" max="15" value="${st.indent || 0}">
      </div><label class="check"><input type="checkbox" id="fc-wrap" ${st.wrap ? 'checked' : ''}> Wrap text</label><label class="check"><input type="checkbox" id="fc-lock" ${st.locked === false ? '' : 'checked'}> Locked (takes effect when the sheet is protected)</label>
      <p class="dlg-note">Custom codes follow spreadsheet conventions, e.g. <code>#,##0.00</code>, <code>"₹"#,##,##0</code>, <code>dd-mmm-yyyy</code>, <code>0.0%</code>.</p>`;
    const nfIn = $('#fc-nf', body), cat = $('#fc-cat', body);
    const sync = () => { const f = SU.format(typeof v === 'number' ? v : 1234.5678, nfIn.value); $('#fc-sample', body).textContent = f.text; $('#fc-sample', body).style.color = f.color || ''; };
    const found = SU.FORMATS.find(f => f.nf === (st.nf || 'General')); cat.value = found ? found.nf : '__custom';
    cat.addEventListener('change', () => { if (cat.value !== '__custom') nfIn.value = cat.value; sync(); });
    nfIn.addEventListener('input', () => { cat.value = SU.FORMATS.find(f => f.nf === nfIn.value)?.nf || '__custom'; sync(); });
    $('#fc-h', body).value = st.h || ''; $('#fc-v', body).value = st.v || 'bottom'; sync();
    const res = await BTS.dialog({ title: 'Format cells', body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ nf: nfIn.value.trim(), h: $('#fc-h', r).value, v: $('#fc-v', r).value, indent: +$('#fc-ind', r).value || 0, wrap: $('#fc-wrap', r).checked, locked: $('#fc-lock', r).checked }) }] });
    if (!res) return;
    A.applyStyle({ nf: !res.nf || res.nf === 'General' ? undefined : res.nf, h: res.h || undefined, v: res.v === 'bottom' ? undefined : res.v, indent: res.indent || undefined, wrap: res.wrap || undefined, locked: res.locked ? undefined : false }, 'Format cells');
  };

  /* ---------- Zoom ---------- */
  Dlg.zoom = async () => {
    const res = await BTS.dialog({ title: 'Zoom', body: `<div style="display:flex;flex-direction:column">${[200, 150, 125, 100, 90, 75, 50, 25].map(z => `<label class="check"><input type="radio" name="zm" value="${z}" ${BTS.zoom === z ? 'checked' : ''}> ${z}%</label>`).join('')}<label class="check"><input type="radio" name="zm" value="custom"> Custom: <input id="zm-c" type="number" min="10" max="400" value="${BTS.zoom}" style="width:70px;height:26px;border:1px solid var(--line-2);border-radius:5px;padding:0 6px">%</label></div>`, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => { const v = r.querySelector('[name=zm]:checked'); return v ? (v.value === 'custom' ? +r.querySelector('#zm-c').value : +v.value) : null; } }] });
    if (res) BTS.setZoom(res);
  };

  /* ---------- statistics ---------- */
  Dlg.stats = () => {
    let cells = 0, formulas = 0, notes = 0, charts = 0, images = 0, tables = 0;
    for (const s of wb.sheets) { const u = s.usedRange(); for (let r = 0; r <= u.r1; r++) for (let c = 0; c <= u.c1; c++) if (!s.isEmpty(r, c)) { cells++; if (s.isFormula(r, c)) formulas++; } notes += s.comments.size; charts += s.charts.length; images += s.images.length; tables += s.tables.length; }
    const s = S(); const u = s.usedRange();
    BTS.dialog({ title: 'Workbook statistics', body: `<table class="nm-table" style="min-width:360px"><tr><th colspan="2">Current sheet</th></tr><tr><td>End of sheet</td><td>${u.r1 >= 0 ? SU.addr(u.r1, u.c1) : '—'}</td></tr><tr><th colspan="2">Workbook</th></tr><tr><td>Sheets</td><td>${wb.sheets.length}</td></tr><tr><td>Cells with data</td><td>${cells}</td></tr><tr><td>Formulas</td><td>${formulas}</td></tr><tr><td>Tables</td><td>${tables}</td></tr><tr><td>Charts</td><td>${charts}</td></tr><tr><td>Pictures</td><td>${images}</td></tr><tr><td>Notes</td><td>${notes}</td></tr><tr><td>Defined names</td><td>${wb.listNames().length}</td></tr></table>`, buttons: [{ label: 'OK', primary: true, value: true }] });
  };

  /* ---------- import text into sheet ---------- */
  Dlg.importText = async () => {
    const f = await BTS.pickFile('.csv,.tsv,.txt'); if (!f) return;
    const text = IO.decodeText(f.data); const rows = SU.parseDelimited(text, /\.tsv$/i.test(f.name) ? '\t' : null);
    const s = S(); const r0 = s.active.r, c0 = s.active.c;
    wb.run('Import text', () => { const cells = []; rows.forEach((row, i) => row.forEach((t, j) => { const res = A.interpret(s, r0 + i, c0 + j, t); cells.push({ r: r0 + i, c: c0 + j, v: res.error ? t : res.v }); if (res.nf) wb.setStyle(s, r0 + i, c0 + j, { nf: res.nf }); })); wb.setContents(s, cells); });
    BTS.toast(`Imported ${rows.length} rows`);
  };
})();
