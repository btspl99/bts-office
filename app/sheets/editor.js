/* BTS Sheets — editing: in-cell editor, formula bar, reference point mode, autocomplete, signature hints, keyboard navigation */
(function () {
  const Ed = window.Ed = {};
  let wb, wrap, ed, fx, keys, acEl, hintEl;
  const $ = (s) => document.querySelector(s);

  /* ---------- text/caret helpers for contenteditable ---------- */
  function getText(el) {
    if (el.tagName === 'TEXTAREA') return el.value;
    let out = ''; const walk = (n, first) => { for (const ch of n.childNodes) { if (ch.nodeType === 3) out += ch.data; else if (ch.nodeName === 'BR') out += '\n'; else { if ((ch.nodeName === 'DIV' || ch.nodeName === 'P') && out.length && !out.endsWith('\n')) out += '\n'; walk(ch); } } };
    walk(el); return out;
  }
  function nodeOffsets(el) { const list = []; let pos = 0; const walk = (n) => { for (const ch of n.childNodes) { if (ch.nodeType === 3) { list.push([ch, pos]); pos += ch.data.length; } else if (ch.nodeName === 'BR') { list.push([ch, pos]); pos += 1; } else walk(ch); } }; walk(el); return list; }
  function getSel(el) {
    if (el.tagName === 'TEXTAREA') return [el.selectionStart, el.selectionEnd];
    const s = getSelection(); if (!s.rangeCount || !el.contains(s.anchorNode)) { const n = getText(el).length; return [n, n]; }
    const r = s.getRangeAt(0); const off = (node, o) => { const pre = document.createRange(); pre.selectNodeContents(el); try { pre.setEnd(node, o); } catch { return getText(el).length; } const frag = pre.cloneContents(); const d = document.createElement('div'); d.appendChild(frag); return getText(d).length; };
    return [off(r.startContainer, r.startOffset), off(r.endContainer, r.endOffset)];
  }
  function setSel(el, a, b = a) {
    if (el.tagName === 'TEXTAREA') { el.setSelectionRange(a, b); return; }
    const list = nodeOffsets(el); const find = (p) => { for (let i = list.length - 1; i >= 0; i--) { const [n, s] = list[i]; if (p >= s) { if (n.nodeType === 3) return [n, Math.min(p - s, n.data.length)]; return [n.parentNode, [...n.parentNode.childNodes].indexOf(n) + (p > s ? 1 : 0)]; } } return [el, 0]; };
    const [n1, o1] = find(a), [n2, o2] = find(b); const r = document.createRange(); r.setStart(n1, o1); r.setEnd(n2, o2); const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  }

  /* ---------- state ---------- */
  // Ed.state: { sheet, r, c, original, mode:'enter'|'edit', src:'cell'|'bar', point:{start,end,anchor,cur}|null, pointSheet, hist:[] }
  Ed.state = null;
  Ed.isEditing = () => !!Ed.state;
  function input() { return Ed.state && Ed.state.src === 'bar' ? fx : ed; }

  Ed.init = (workbook) => {
    wb = workbook; wrap = $('#cell-editor-wrap'); ed = $('#cell-editor'); fx = $('#fx'); keys = $('#gkeys'); acEl = $('#ac-list'); hintEl = $('#fn-hint');
    keys.addEventListener('keydown', onGridKey);
    const resetKeys = () => { keys.value = ' '; keys.select(); }; Ed.resetKeys = resetKeys;
    keys.addEventListener('focus', resetKeys);
    keys.addEventListener('input', (e) => { if (e.isComposing) return; const v = keys.value; resetKeys(); if (v && !Ed.state && !A.protectedCheck()) Ed.start(v, 'enter'); });
    keys.addEventListener('compositionend', () => { const v = keys.value; resetKeys(); if (!v) return; if (Ed.state) Ed.insertAtCaret(v); else if (!A.protectedCheck()) Ed.start(v, 'enter'); });
    keys.addEventListener('compositionstart', () => { if (!Ed.state && !A.protectedCheck(true)) { keys.value = ''; } });
    ed.addEventListener('keydown', onEditKey);
    ed.addEventListener('input', () => onEditInput(ed));
    ed.addEventListener('paste', (e) => { e.preventDefault(); const t = e.clipboardData.getData('text/plain').replace(/\r\n/g, '\n').replace(/\n$/, ''); insertText(ed, t); });
    ed.addEventListener('mouseup', () => updateAssist());
    ed.addEventListener('blur', () => setTimeout(() => { if (Ed.state && Ed.state.src === 'cell' && !document.activeElement.closest('#cell-editor,#fx,.modal-back,.pop,#ac-list,#tabs,#namebox') && document.activeElement !== keys) {/* keep editing (ribbon clicks) */ } }, 0));
    fx.addEventListener('focus', () => { if (A.protectedCheck(true)) { fx.blur(); return; } if (!Ed.state) Ed.start(null, 'edit', 'bar'); else if (Ed.state.src !== 'bar') { Ed.state.src = 'bar'; } });
    fx.addEventListener('keydown', onEditKey);
    fx.addEventListener('input', () => { autosizeFx(); onEditInput(fx); });
    fx.addEventListener('click', () => updateAssist());
    $('#fb-cancel').addEventListener('mousedown', (e) => { e.preventDefault(); Ed.cancel(); });
    $('#fb-ok').addEventListener('mousedown', (e) => { e.preventDefault(); Ed.commit(); });
    $('#fb-fx').addEventListener('mousedown', (e) => { e.preventDefault(); Dlg.insertFunction(); });
    $('#fb-expand').addEventListener('click', () => { $('#fbar').classList.toggle('expanded'); $('#fb-expand').innerHTML = icon($('#fbar').classList.contains('expanded') ? 'chevron-up' : 'chevron-down', 16); });
    $('#fb-expand').innerHTML = icon('chevron-down', 16); $('#fb-cancel').innerHTML = icon('x', 16); $('#fb-ok').innerHTML = icon('check', 16);
    acEl.addEventListener('mousedown', (e) => { e.preventDefault(); const it = e.target.closest('.ac'); if (it) acceptAC(+it.dataset.i); });
    // clipboard events
    document.addEventListener('copy', (e) => { if (Ed.state || !gridFocused()) return; e.preventDefault(); A.copy(false, e.clipboardData); });
    document.addEventListener('cut', (e) => { if (Ed.state || !gridFocused()) return; e.preventDefault(); A.copy(true, e.clipboardData); });
    document.addEventListener('paste', (e) => { if (Ed.state || !gridFocused()) return; e.preventDefault(); A.pasteFromEvent(e.clipboardData); });
    Grid.pointMode = pointClick; Grid.pointDrag = pointDrag; Grid.commitEdit = () => Ed.commit();
  };
  function gridFocused() { const a = document.activeElement; return a === keys || a === document.body || (a && a.closest && a.closest('#ribbon')); }

  /* ---------- start / commit / cancel ---------- */
  Ed.start = (initial, mode = 'edit', src = 'cell') => {
    const s = wb.active; const { r, c } = s.active;
    if (A.protectedCheck()) return;
    const original = s.editText(r, c);
    Ed.state = { sheet: s, r, c, original, mode, src, point: null, pointSheet: s, hist: [] };
    Grid.editing = true;
    document.body.classList.add('editing'); A.setMode(mode === 'enter' ? 'Enter' : 'Edit');
    $('#fb-cancel').classList.remove('hidden'); $('#fb-ok').classList.remove('hidden');
    let text = initial === null || initial === undefined ? original : initial;
    // percent cells: typing a number keeps percent semantics
    positionEditor();
    renderEditor(text); fx.value = text; autosizeFx();
    if (src === 'cell') { ed.focus({ preventScroll: true }); setSel(ed, text.length); }
    else { fx.focus({ preventScroll: true }); }
    Ed.state.hist.push(text);
    onEditInput(input(), true);
    Grid.render();
  };
  function positionEditor() {
    const st = Ed.state; if (!st) return;
    if (st.pointSheet !== st.sheet) { wrap.classList.add('hidden'); return; }
    const s = st.sheet; const rc = Grid.cellRect(st.r, st.c); const z = Grid.z;
    const style = s.style(st.r, st.c);
    wrap.classList.remove('hidden');
    const host = $('#grid-host').getBoundingClientRect();
    wrap.style.left = Math.max(Grid.hw(), rc.x) + 'px'; wrap.style.top = Math.max(Grid.hh(), rc.y) + 'px';
    wrap.style.minWidth = (rc.w - 1) + 'px'; wrap.style.minHeight = (rc.h - 1) + 'px'; wrap.style.width = 'auto'; wrap.style.height = 'auto';
    wrap.style.maxWidth = Math.max(rc.w, host.width - rc.x - 20) + 'px';
    wrap.style.background = style.fill || '#fff';
    const fs = (style.sz || 11) * 96 / 72 * z;
    ed.style.font = `${style.i ? 'italic ' : ''}${style.b ? 'bold ' : ''}${fs}px ${SU.cssFont(style.font || SU.defFont)}`;
    ed.style.color = style.color || '#000'; ed.style.textAlign = style.h === 'center' ? 'center' : style.h === 'right' ? 'right' : 'left';
    ed.style.paddingTop = Math.max(0, (rc.h - fs * 1.22) - 3 * z) + 'px';
    if (style.wrap) { ed.style.paddingTop = '1px'; }
    ed.style.whiteSpace = style.wrap ? 'pre-wrap' : 'pre';
    ed.style.width = style.wrap ? (rc.w - 1) + 'px' : '';
    ed.style.textDecoration = [style.u ? 'underline' : '', style.s ? 'line-through' : ''].join(' ').trim() || 'none';
  }
  Ed.reposition = () => { if (Ed.state) { positionEditor(); growEditor(); } };
  function growEditor() {
    if (wrap.classList.contains('hidden')) return;
    const max = parseFloat(wrap.style.maxWidth) || 600;
    if (ed.scrollWidth > max - 2 && ed.style.whiteSpace === 'pre') { ed.style.whiteSpace = 'pre-wrap'; ed.style.width = max - 2 + 'px'; }
  }
  Ed.commit = (move) => {
    const st = Ed.state; if (!st) return true;
    hideAssist();
    const text = getText(input());
    if (text === st.original && !move?.force) { finish(); return true; }
    const ok = A.enterValue(st.sheet, st.r, st.c, text, { ctrlEnter: move && move.ctrlEnter });
    if (ok === false) return false;
    finish();
    return true;
  };
  Ed.cancel = () => { if (!Ed.state) return; hideAssist(); finish(); };
  function finish() {
    const st = Ed.state; Ed.state = null; Grid.editing = false; Grid.refHL = [];
    document.body.classList.remove('editing'); A.setMode('Ready');
    wrap.classList.add('hidden'); ed.textContent = '';
    $('#fb-cancel').classList.add('hidden'); $('#fb-ok').classList.add('hidden');
    if (st && st.pointSheet !== st.sheet) App.switchSheet(wb.sheets.indexOf(st.sheet));
    keys.focus({ preventScroll: true }); App.refreshFormulaBar(); Grid.render();
  }

  /* ---------- rendering editor text (colored refs) ---------- */
  function refColorMap(text) {
    const refs = SU.formulaRefs(text); const map = new Map(); const out = [];
    for (const rf of refs) { const k = (rf.sheet || '').toLowerCase() + '!' + SU.rangeStr(rf.range); if (!map.has(k)) map.set(k, Grid.REF_COLORS[map.size % Grid.REF_COLORS.length]); out.push({ ...rf, color: map.get(k) }); }
    return out;
  }
  function renderEditor(text, caret) {
    if (text && text[0] === '=') {
      const refs = refColorMap(text); let html = '', i = 0;
      for (const rf of refs) { html += SU.esc(text.slice(i, rf.start)) + `<span class="ref-tok" style="color:${rf.color}">${SU.esc(text.slice(rf.start, rf.end))}</span>`; i = rf.end; }
      html += SU.esc(text.slice(i)); ed.innerHTML = html;
    } else ed.textContent = text;
    if (caret !== undefined && document.activeElement === ed) setSel(ed, caret);
    growEditor();
  }
  function autosizeFx() { if ($('#fbar').classList.contains('expanded')) return; fx.style.height = '26px'; if (fx.scrollHeight > 28 && document.activeElement === fx) fx.style.height = Math.min(120, fx.scrollHeight + 2) + 'px'; }
  function updateHighlights(text) {
    const st = Ed.state; if (!st) return;
    if (!text || text[0] !== '=') { Grid.refHL = []; Grid.render(); return; }
    const refs = refColorMap(text); const cur = st.pointSheet;
    Grid.refHL = refs.filter(rf => { if (!rf.sheet) return cur === st.sheet; const s = wb.sheetByName(rf.sheet); return s === cur; }).map(rf => ({ range: rf.range, color: rf.color }));
    Grid.render();
  }
  function onEditInput(el, initial) {
    const st = Ed.state; if (!st) return;
    const text = getText(el);
    if (!initial) { st.point = null; if (st.hist[st.hist.length - 1] !== text) st.hist.push(text); if (st.hist.length > 200) st.hist.shift(); }
    if (el === ed) { if (text[0] === '=') { const [a] = getSel(ed); renderEditor(text, a); } fx.value = text; autosizeFx(); }
    else { renderEditor(text); }
    growEditor();
    updateHighlights(text); updateAssist();
  }
  function setEditText(text, caret) {
    const el = input();
    if (el === fx) { fx.value = text; setSel(fx, caret); renderEditor(text); autosizeFx(); }
    else { renderEditor(text, caret); fx.value = text; }
    const st = Ed.state; if (st && st.hist[st.hist.length - 1] !== text) st.hist.push(text);
    updateHighlights(text); updateAssist();
  }
  function insertText(el, t) { const text = getText(el); const [a, b] = getSel(el); const nt = text.slice(0, a) + t + text.slice(b); if (el === ed) { renderEditor(nt, a + t.length); fx.value = nt; } else { fx.value = nt; setSel(fx, a + t.length); renderEditor(nt); } Ed.state.point = null; onEditInput(el); setSel(el, a + t.length); }
  Ed.insertAtCaret = (t, selectBack = 0) => {
    if (!Ed.state) { Ed.start('', 'edit'); }
    const el = input(); el.focus(); const text = getText(el); const [a, b] = getSel(el); const nt = text.slice(0, a) + t + text.slice(b);
    setEditText(nt, a + t.length - selectBack); setSel(el, a + t.length - selectBack);
  };
  Ed.setText = (t) => { if (!Ed.state) Ed.start(t, 'edit'); else setEditText(t, t.length); input().focus(); setSel(input(), t.length); };

  /* ---------- point mode ---------- */
  function pointable(text, caret) {
    if (!text || text[0] !== '=') return false;
    const st = Ed.state; if (st.point && st.point.end === caret) return true;
    let i = caret - 1; while (i >= 0 && text[i] === ' ') i--;
    if (i < 0) return false;
    return '=(,+-*/^&<>:;%{'.includes(text[i]) || (text[i] === ' ' && false);
  }
  Ed.pointable = () => { const st = Ed.state; if (!st) return false; const el = input(); const [a] = getSel(el); return pointable(getText(el), a); };
  function refText(g, sheetOfRef) {
    const st = Ed.state; let t = SU.rangeStr(g);
    if (sheetOfRef !== st.sheet) t = SU.quoteSheet(sheetOfRef.name) + '!' + t;
    return t;
  }
  function placeRef(g) {
    const st = Ed.state; const el = input(); const text = getText(el); let a, b;
    if (st.point) { a = st.point.start; b = st.point.end; } else { [a, b] = getSel(el); }
    const t = refText(g, st.pointSheet);
    const nt = text.slice(0, a) + t + text.slice(b);
    setEditText(nt, a + t.length);
    if (el === fx) setSel(fx, a + t.length);
    st.point = { ...(st.point || {}), start: a, end: a + t.length, g };
  }
  function pointClick(h, e) {
    const st = Ed.state; if (!st) return false; if (h.zone !== 'cell' && h.zone !== 'colhdr' && h.zone !== 'rowhdr') return false;
    const el = input(); const text = getText(el); const [a] = getSel(el);
    if (!pointable(text, a) && !(st.point && st.point.end === a)) return false;
    if (st.mode === 'edit' && !(st.point) && !pointable(text, a)) return false;
    let g;
    if (h.zone === 'colhdr') g = { r0: 0, r1: SU.MAXR - 1, c0: h.c, c1: h.c };
    else if (h.zone === 'rowhdr') g = { r0: h.r, r1: h.r, c0: 0, c1: SU.MAXC - 1 };
    else { const m = st.pointSheet.mergeAt(h.r, h.c); g = m ? { ...m } : { r0: h.r, c0: h.c, r1: h.r, c1: h.c }; }
    if (e.shiftKey && st.point && st.point.g) { const p = st.point; g = SU.norm({ r0: p.anchor ? p.anchor.r : p.g.r0, c0: p.anchor ? p.anchor.c : p.g.c0, r1: h.r, c1: h.c }); }
    placeRef(g.r1 >= SU.MAXR - 1 ? g : g);
    if (h.zone === 'colhdr') { const s = refTextCols(g); replacePoint(s); }
    if (h.zone === 'rowhdr') { const s = (g.r0 + 1) + ':' + (g.r1 + 1); replacePoint(s); }
    st.point.anchor = { r: g.r0, c: g.c0 }; st.point.cur = { r: g.r1, c: g.c1 };
    st.mode = 'enter'; A.setMode('Point');
    setTimeout(() => input().focus({ preventScroll: true }), 0);
    return true;
  }
  function refTextCols(g) { return SU.colName(g.c0) + ':' + SU.colName(g.c1); }
  function replacePoint(s) { const st = Ed.state; const el = input(); const text = getText(el); const p = st.point; const pre = st.pointSheet !== st.sheet ? SU.quoteSheet(st.pointSheet.name) + '!' : ''; const nt = text.slice(0, p.start) + pre + s + text.slice(p.end); setEditText(nt, p.start + pre.length + s.length); p.end = p.start + pre.length + s.length; }
  function pointDrag(anchor, cur) {
    const st = Ed.state; if (!st || !st.point) return;
    const g = st.pointSheet.expandMerges(SU.norm({ r0: anchor.r, c0: anchor.c, r1: cur.r, c1: cur.c }));
    placeRef(g); st.point.anchor = anchor; st.point.cur = cur;
  }
  function pointArrow(dr, dc, shift, ctrl) {
    const st = Ed.state; const s = st.pointSheet;
    let cur = st.point && st.point.cur ? { ...st.point.cur } : { r: st.r, c: st.c };
    let anchor = st.point && st.point.anchor ? { ...st.point.anchor } : { ...cur };
    if (!st.point) { cur = { r: st.r, c: st.c }; anchor = { ...cur }; }
    const n = ctrl ? A.jumpEdge(s, cur, dr, dc) : { r: SU.clamp(cur.r + dr, 0, SU.MAXR - 1), c: SU.clamp(cur.c + dc, 0, SU.MAXC - 1) };
    if (!shift) anchor = { ...n };
    const g = s.expandMerges(SU.norm({ r0: anchor.r, c0: anchor.c, r1: n.r, c1: n.c }));
    placeRef(g); st.point.anchor = anchor; st.point.cur = n;
    A.setMode('Point'); Grid.scrollIntoView(n.r, n.c);
  }
  /** called when the user clicks a sheet tab while editing a formula */
  Ed.pointToSheet = (sheet) => {
    const st = Ed.state; if (!st) return false;
    const el = input(); const text = getText(el); const [a] = getSel(el);
    if (!(text[0] === '=' && (pointable(text, a) || st.point))) return false;
    st.pointSheet = sheet; st.point = st.point && st.point.end === a ? st.point : null;
    if (st.src === 'cell') { st.src = 'bar'; fx.value = text; fx.focus(); setSel(fx, a); }
    positionEditor(); updateHighlights(text); return true;
  };

  /* ---------- F4: cycle absolute/relative ---------- */
  function toggleAbs() {
    const el = input(); const text = getText(el); if (text[0] !== '=') return; const [a] = getSel(el);
    const refs = SU.formulaRefs(text); const rf = refs.find(x => a >= x.start && a <= x.end); if (!rf || rf.ref.kind !== 'cell') return;
    const cyc = (p) => { const st = (p.ac ? 2 : 0) + (p.ar ? 1 : 0); const nx = { 0: 3, 3: 1, 1: 2, 2: 0 }[st]; p.ac = !!(nx & 2); p.ar = !!(nx & 1); };
    cyc(rf.ref.a); if (rf.ref.b) { rf.ref.b.ac = rf.ref.a.ac; rf.ref.b.ar = rf.ref.a.ar; }
    const t = SU.refToString(rf.ref); const nt = text.slice(0, rf.start) + t + text.slice(rf.end);
    setEditText(nt, rf.start + t.length); setSel(el, rf.start + t.length);
  }

  /* ---------- autocomplete & signature hint ---------- */
  let ac = null; // {items, i, start, end}
  let fnNames = null;
  function allFunctionNames() {
    if (fnNames) return fnNames;
    const set = new Set(SFUNCS.list.map(f => f.name));
    try { wb.hf.getRegisteredFunctionNames().forEach(n => { if (!/^HF\./.test(n)) set.add(n); }); } catch { }
    fnNames = [...set].sort(); return fnNames;
  }
  function updateAssist() {
    const st = Ed.state; if (!st) return hideAssist();
    const el = input(); const text = getText(el); const [a, b] = getSel(el);
    if (text[0] !== '=' || a !== b) { hideAssist(); return; }
    // autocomplete token
    const before = text.slice(0, a);
    const inStr = (before.split('"').length - 1) % 2 === 1;
    const m = !inStr && /(^|[=(,+\-*\/^&<>:;\s{])([A-Za-z_][A-Za-z0-9_.]*)$/.exec(before);
    const after = text.slice(a);
    if (m && !/^[A-Za-z0-9_.(]/.test(after) && !/^[A-Za-z]{1,3}\d+$/.test(m[2]) ) {
      const tok = m[2].toUpperCase();
      const names = wb.listNames().map(n => n.name);
      const POP = ['SUM', 'AVERAGE', 'IF', 'COUNT', 'COUNTIF', 'COUNTA', 'MAX', 'MIN', 'VLOOKUP', 'XLOOKUP', 'SUMIF', 'ROUND', 'IFERROR', 'INDEX', 'MATCH', 'TODAY', 'LEFT', 'RIGHT', 'LEN', 'TEXT', 'AND', 'OR', 'DATE', 'NOW', 'CONCATENATE', 'TRIM', 'UPPER', 'PMT', 'RANK', 'MEDIAN'];
      const rank = (n) => n === tok ? -1 : POP.includes(n) ? POP.indexOf(n) : 1000;
      const items = [...allFunctionNames().filter(n => n.startsWith(tok)).sort((a, b) => rank(a) - rank(b) || a.localeCompare(b)).map(n => ({ name: n, fn: true })), ...names.filter(n => n.toUpperCase().startsWith(tok)).map(n => ({ name: n, fn: false }))].slice(0, 60);
      if (items.length && !(items.length === 1 && items[0].name === tok && !items[0].fn)) { showAC(items, a - m[2].length, a); }
      else hideAC();
    } else hideAC();
    showHint(text, a);
  }
  function anchorPos() {
    const st = Ed.state; const el = input();
    if (el === fx) { const r = fx.getBoundingClientRect(); const ws = $('#workspace').getBoundingClientRect(); return { x: r.left - ws.left + 8, y: 2 }; }
    const r = wrap.getBoundingClientRect(); const ws = $('#workspace').getBoundingClientRect(); return { x: r.left - ws.left, y: r.bottom - ws.top + 2 };
  }
  function showAC(items, start, end) {
    ac = { items, i: 0, start, end };
    acEl.innerHTML = items.map((it, i) => { const f = SFUNCS.byName.get(it.name); return `<div class="ac ${i === 0 ? 'on' : ''}" data-i="${i}"><b>${it.fn ? '' : icon('tag', 13) + ' '}${SU.esc(it.name)}</b><small>${SU.esc(f ? f.desc : it.fn ? 'Function' : 'Defined name')}</small></div>`; }).join('');
    const p = anchorPos(); acEl.style.left = p.x + 'px'; acEl.style.top = p.y + 'px'; acEl.classList.remove('hidden');
    hintEl.classList.add('hidden');
  }
  function hideAC() { ac = null; acEl.classList.add('hidden'); }
  function hideAssist() { hideAC(); hintEl.classList.add('hidden'); }
  function moveAC(d) { if (!ac) return; ac.i = (ac.i + d + ac.items.length) % ac.items.length; [...acEl.children].forEach((x, i) => x.classList.toggle('on', i === ac.i)); acEl.children[ac.i].scrollIntoView({ block: 'nearest' }); }
  function acceptAC(i) {
    if (!ac) return; const it = ac.items[i === undefined ? ac.i : i]; const el = input(); const text = getText(el);
    const ins = it.fn ? it.name + '(' : it.name; const nt = text.slice(0, ac.start) + ins + text.slice(ac.end);
    const caret = ac.start + ins.length; hideAC();
    setEditText(nt, caret); setSel(el, caret); el.focus(); updateAssist();
  }
  function currentCall(text, caret) {
    let depth = 0, inStr = false; const stack = [];
    for (let i = 0; i < caret; i++) {
      const ch = text[i];
      if (ch === '"') { inStr = !inStr; continue; } if (inStr) continue;
      if (ch === '(') { const m = /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(text.slice(0, i)); stack.push({ name: m ? m[1].toUpperCase() : '', arg: 0 }); }
      else if (ch === ')') stack.pop();
      else if (ch === ',' && stack.length) stack[stack.length - 1].arg++;
      else if (ch === '{') depth++; else if (ch === '}') depth--;
    }
    for (let i = stack.length - 1; i >= 0; i--) if (stack[i].name) return stack[i];
    return null;
  }
  function showHint(text, caret) {
    const call = currentCall(text, caret); const f = call && SFUNCS.byName.get(call.name);
    if (!f || ac) { hintEl.classList.add('hidden'); return; }
    const args = f.args ? f.args.split(', ') : [];
    let idx = call.arg; const rep = args.findIndex(a => a === '...');
    if (rep >= 0 && idx >= rep - 1) idx = rep - 1 >= 0 ? Math.min(idx, rep - 1) : idx;
    hintEl.innerHTML = `<b>${f.name}</b>(${args.map((a, i) => i === idx ? `<span class="cur">${SU.esc(a)}</span>` : SU.esc(a)).join(', ')})`;
    hintEl.title = f.desc;
    const p = anchorPos(); hintEl.style.left = p.x + 'px'; hintEl.style.top = p.y + 'px'; hintEl.classList.remove('hidden');
  }

  /* ---------- editor keys ---------- */
  function onEditKey(e) {
    const st = Ed.state; if (!st) return;
    const el = e.currentTarget;
    if (ac) {
      if (e.key === 'ArrowDown') { e.preventDefault(); moveAC(1); return; }
      if (e.key === 'ArrowUp') { e.preventDefault(); moveAC(-1); return; }
      if (e.key === 'Tab' || (e.key === 'Enter' && !e.altKey && !e.ctrlKey && !e.shiftKey && ac.items[ac.i].name !== getText(el).slice(ac.start, ac.end).toUpperCase())) { e.preventDefault(); acceptAC(); return; }
      if (e.key === 'Escape') { e.preventDefault(); hideAC(); return; }
    }
    if (e.key === 'Enter' && e.altKey) { e.preventDefault(); insertText(el, '\n'); return; }
    if (e.key === 'Enter') {
      e.preventDefault();
      if (e.ctrlKey) { if (Ed.commit({ ctrlEnter: true, force: true })) {} return; }
      const ok = Ed.commit(); if (ok) A.moveAfterEnter(e.shiftKey ? -1 : 1, 0); return;
    }
    if (e.key === 'Tab') { e.preventDefault(); const ok = Ed.commit(); if (ok) A.moveAfterEnter(0, e.shiftKey ? -1 : 1); return; }
    if (e.key === 'Escape') { e.preventDefault(); Ed.cancel(); return; }
    if (e.key === 'F2') { e.preventDefault(); st.mode = st.mode === 'edit' ? 'enter' : 'edit'; st.point = null; A.setMode(st.mode === 'edit' ? 'Edit' : 'Enter'); return; }
    if (e.key === 'F4') { e.preventDefault(); toggleAbs(); return; }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); if (st.hist.length > 1) { st.hist.pop(); const t = st.hist[st.hist.length - 1]; setEditText(t, t.length); st.hist.length && st.hist.pop(); st.hist.push(t); setSel(el, t.length); } else Ed.cancel(); return; }
    if ((e.ctrlKey || e.metaKey) && e.key === ';') { e.preventDefault(); const d = new Date(); insertText(el, e.shiftKey ? d.toTimeString().slice(0, 5) : `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`); return; }
    if ((e.ctrlKey || e.metaKey) && (e.key === 'b' || e.key === 'i' || e.key === 'u')) { e.preventDefault(); return; }
    if (e.key.startsWith('Arrow')) {
      const dr = e.key === 'ArrowDown' ? 1 : e.key === 'ArrowUp' ? -1 : 0, dc = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0;
      const text = getText(el); const [a] = getSel(el);
      if (st.mode === 'enter' && text[0] === '=' && (pointable(text, a) || (st.point && st.point.end === a))) { e.preventDefault(); pointArrow(dr, dc, e.shiftKey, e.ctrlKey); return; }
      if (st.mode === 'enter' && st.src === 'cell' && text[0] !== '=') { e.preventDefault(); if (Ed.commit()) A.moveSel(dr, dc, e.shiftKey, e.ctrlKey); return; }
      if (st.mode === 'enter' && st.src === 'cell' && text[0] === '=' && !pointable(text, a)) { e.preventDefault(); if (Ed.commit()) A.moveSel(dr, dc, false, false); return; }
      st.point = null; setTimeout(updateAssist, 0); return;
    }
    if (e.key === 'Home' || e.key === 'End' || e.key === 'PageUp' || e.key === 'PageDown') { st.point = null; setTimeout(updateAssist, 0); }
  }

  /* ---------- grid keys (not editing) ---------- */
  function onGridKey(e) {
    const s = wb.active; const mod = e.ctrlKey || e.metaKey; const k = e.key;
    if (Charts.selected && (k === 'Delete' || k === 'Backspace')) { e.preventDefault(); Charts.deleteSelected(); return; }
    if (Charts.selected && k === 'Escape') { Charts.select(null); return; }
    if (e.altKey && k === 'ArrowDown') { e.preventDefault(); A.openDropdownAtActive(); return; }
    if (e.altKey && (k === '=' || k === '+')) { e.preventDefault(); e.stopPropagation(); A.autoSum('SUM'); return; }
    if (e.altKey && k === 'Enter') { e.preventDefault(); Ed.start(null, 'edit'); return; }
    if (e.altKey && e.shiftKey && !mod && (k === 'ArrowRight' || k === 'ArrowLeft')) { e.preventDefault(); App.cmd(k === 'ArrowRight' ? 'group' : 'ungroup'); return; }
    if (mod && e.altKey && k.toLowerCase() === 'm') { e.preventDefault(); App.cmd('comment-new'); return; }
    if (e.altKey && k === 'F5') { e.preventDefault(); App.cmd(mod ? 'pivot-refresh-all' : 'pv-refresh'); return; }
    if (mod && e.shiftKey && k.toLowerCase() === 'f') { e.preventDefault(); App.cmd('font-pick'); return; }
    if (k.startsWith('Arrow')) {
      e.preventDefault();
      const dr = k === 'ArrowDown' ? 1 : k === 'ArrowUp' ? -1 : 0, dc = k === 'ArrowRight' ? 1 : k === 'ArrowLeft' ? -1 : 0;
      A.moveSel(dr, dc, e.shiftKey, mod); return;
    }
    switch (k) {
      case 'Enter': e.preventDefault(); if (mod) return; A.moveAfterEnter(e.shiftKey ? -1 : 1, 0); return;
      case 'Tab': e.preventDefault(); A.moveAfterEnter(0, e.shiftKey ? -1 : 1); return;
      case 'F2': e.preventDefault(); if (e.shiftKey) App.cmd('comment'); else Ed.start(null, 'edit'); return;
      case 'F3': if (e.shiftKey) { e.preventDefault(); Dlg.insertFunction(); } return;
      case 'F4': e.preventDefault(); A.repeatLast(); return;
      case 'F9': e.preventDefault(); App.cmd('recalc'); return;
      case 'F11': if (e.shiftKey) { e.preventDefault(); App.cmd('sheet-add'); } return;
      case 'Delete': e.preventDefault(); A.clear('contents'); return;
      case 'Backspace': e.preventDefault(); if (!A.protectedCheck()) Ed.start('', 'enter'); return;
      case 'Escape': if (Grid.copyRange) { Grid.copyRange = null; Grid.render(); } A.cancelPainter(); return;
      case 'Home': e.preventDefault(); if (mod) { const fr = s.freeze; A.moveTo(fr.r, fr.c, e.shiftKey); } else A.moveTo(s.active.r, s.freeze.c && s.active.c >= s.freeze.c ? s.freeze.c : 0, e.shiftKey); return;
      case 'End': if (mod) { e.preventDefault(); const u = s.usedRange(); A.moveTo(Math.max(0, u.r1), Math.max(0, u.c1), e.shiftKey); } return;
      case 'PageDown': case 'PageUp': {
        e.preventDefault(); const d = k === 'PageDown' ? 1 : -1;
        if (mod) { App.switchSheet(SU.clamp(wb.activeIndex + d, 0, wb.sheets.length - 1)); return; }
        if (e.altKey) { const cur = s.active; A.moveTo(cur.r, SU.clamp(cur.c + d * 10, 0, SU.MAXC - 1), e.shiftKey); return; }
        const n = Grid.pageRows(); const cur = e.shiftKey ? A.selEnd() : s.active; const r = SU.clamp(cur.r + d * n, 0, SU.MAXR - 1);
        Grid.scroller.scrollTop += d * n * SU.DEF_H * Grid.z; A.moveTo(r, cur.c, e.shiftKey); return;
      }
      case ' ': if (mod && e.shiftKey) { e.preventDefault(); Grid.selectAll(); return; } if (mod) { e.preventDefault(); Grid.selectCols(s.sel.c0, s.sel.c1, s.active.c); return; } if (e.shiftKey) { e.preventDefault(); Grid.selectRows(s.sel.r0, s.sel.r1, s.active.r); return; } break;
      case 'ContextMenu': e.preventDefault(); { const rc = Grid.cellRect(s.active.r, s.active.c); const hr = $('#grid-host').getBoundingClientRect(); App.contextMenu({ zone: 'cell', r: s.active.r, c: s.active.c }, { clientX: hr.left + rc.x + rc.w / 2, clientY: hr.top + rc.y + rc.h / 2 }); } return;
    }
    if (mod) {
      const lk = k.toLowerCase();
      const map = {
        z: 'undo', y: 'redo', b: 'bold', i: 'italic', u: 'underline', '5': 'strike', f: 'find', h: 'replace', g: 'goto', k: 'hyperlink',
        d: 'fill-down', r: 'fill-right', '`': 'show-formulas', '~': 'show-formulas', e: 'flash-fill', '1': 'format-cells', '9': 'hide-rows', '0': 'hide-cols', l: null, t: 'format-table',
      };
      if (e.shiftKey) {
        if (lk === 'l') { e.preventDefault(); App.cmd('filter'); return; }
        if (k === '$' || k === '4') { e.preventDefault(); A.setNumberFormat(SU.currencyFormat('₹')); return; }
        if (k === '%') { e.preventDefault(); App.cmd('fmt-percent'); return; }
        if (k === '!' || k === '1') { e.preventDefault(); A.setNumberFormat('#,##0.00'); return; }
        if (k === '#' || k === '3') { e.preventDefault(); A.setNumberFormat('dd-mmm-yy'); return; }
        if (k === '@' || k === '2') { e.preventDefault(); A.setNumberFormat('h:mm AM/PM'); return; }
        if (k === '~' || k === '`') { e.preventDefault(); A.setNumberFormat('General'); return; }
        if (k === '+' || k === '=') { e.preventDefault(); App.cmd('insert-cells-dialog'); return; }
        if (k === '&' || k === '7') { e.preventDefault(); App.cmd('border-outside'); return; }
        if (k === '_' || k === '-') { e.preventDefault(); App.cmd('border-none'); return; }
        if (k === ':' || k === ';') { e.preventDefault(); A.insertNow(true); return; }
        if (lk === 'z') { e.preventDefault(); App.cmd('redo'); return; }
        if (lk === 'u') { e.preventDefault(); $('#fbar').classList.toggle('expanded'); return; }
      }
      if (k === ';') { e.preventDefault(); A.insertNow(false); return; }
      if (k === '-') { e.preventDefault(); App.cmd('delete-cells-dialog'); return; }
      if (k === '+' || k === '=') { e.preventDefault(); App.cmd('insert-cells-dialog'); return; }
      if (lk === 'a') { e.preventDefault(); A.selectAllSmart(); return; }
      if (lk === 'c' || lk === 'x' || lk === 'v') return; // native clipboard events
      if (map[lk]) { e.preventDefault(); App.cmd(map[lk]); return; }
      return;
    }
    if (k === 'F5') { e.preventDefault(); App.cmd('goto'); return; }
  }
})();
