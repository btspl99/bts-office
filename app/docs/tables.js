/* BTS Docs — tables: insert, grid model (merged cells), row/col ops, merge/split,
 * styles, shading, borders, alignment, autofit, Tab navigation, drag cell selection, column resize. */
(function () {
  const D = window.Docs; const { $, $$ } = D;

  D.TABLE_STYLES = [
    { id: 't-plain', name: 'Grid' },
    { id: 't-header', name: 'Header row' },
    { id: 't-banded', name: 'Banded rows' },
    { id: 't-grid-accent', name: 'Accent grid' },
    { id: 't-light', name: 'Light lines' },
    { id: 't-dark', name: 'Dark header' },
    { id: 't-minimal', name: 'Minimal' },
  ];
  const STYLE_IDS = D.TABLE_STYLES.map(s => s.id);

  D.makeTable = (rows, cols, style = 't-plain') => {
    const t = document.createElement('table'); t.className = 'tbl ' + style;
    const cg = document.createElement('colgroup');
    for (let c = 0; c < cols; c++) { const col = document.createElement('col'); col.style.width = (100 / cols).toFixed(3) + '%'; cg.appendChild(col); }
    t.appendChild(cg);
    const tb = document.createElement('tbody');
    for (let r = 0; r < rows; r++) { const tr = document.createElement('tr'); for (let c = 0; c < cols; c++) { const td = document.createElement('td'); td.innerHTML = '<p><br></p>'; tr.appendChild(td); } tb.appendChild(tr); }
    t.appendChild(tb); return t;
  };
  D.insertTable = (rows, cols, style) => D.change(() => {
    const t = D.makeTable(rows, cols, style);
    const r = D.ensureSel(); const cell = D.closest(r.startContainer, 'td,th');
    if (cell) { // nested tables: insert inside the cell
      cell.appendChild(t); const first = t.querySelector('td p'); D.setCaret(first, 0); return;
    }
    if (!D.doc.contains(r.startContainer)) D.doc.focus();
    D.insertBlocks([t], t.querySelector('td p'));
    D.scrollCaret();
  });
  D.insertTableDialog = async () => {
    D.ensureSel(); const sv = D.saveSel();
    const v = await BTS.dialog({
      title: 'Insert table', width: 420,
      body: `<div class="field-row"><div class="field"><label>Number of columns</label><input id="it-c" type="number" min="1" max="63" value="3"></div><div class="field"><label>Number of rows</label><input id="it-r" type="number" min="1" max="500" value="3"></div></div>
        <div class="field"><label>Style</label><select id="it-s">${D.TABLE_STYLES.map(s => `<option value="${s.id}">${s.name}</option>`).join('')}</select></div>
        <label class="check"><input type="checkbox" id="it-fit"> AutoFit to contents</label>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'Insert', primary: true, value: (r) => ({ c: +$('#it-c', r).value, r: +$('#it-r', r).value, s: $('#it-s', r).value, fit: $('#it-fit', r).checked }) }],
    });
    D.restoreSel(sv);
    if (!v) return;
    const c = Math.max(1, Math.min(63, v.c || 3)), rr = Math.max(1, Math.min(500, v.r || 3));
    D.insertTable(rr, c, v.s);
    if (v.fit) D.tableOp('autofit-content');
  };

  /* ---------- grid model ---------- */
  /** Build a matrix: grid[r][c] = cell element (merged cells repeated). */
  function grid(t) {
    const g = []; const rows = [...t.rows];
    rows.forEach((tr, r) => {
      g[r] = g[r] || [];
      let c = 0;
      [...tr.cells].forEach(cell => {
        while (g[r][c]) c++;
        const rs = cell.rowSpan || 1, cs = cell.colSpan || 1;
        for (let i = 0; i < rs; i++) for (let j = 0; j < cs; j++) { g[r + i] = g[r + i] || []; g[r + i][c + j] = cell; }
        c += cs;
      });
    });
    const w = Math.max(0, ...g.map(r => r.length));
    return { g, rows: rows.length, cols: w };
  }
  D.tableGrid = grid;
  function pos(g, cell) { for (let r = 0; r < g.g.length; r++) { const c = g.g[r].indexOf(cell); if (c >= 0) return { r, c }; } return null; }
  function ensureColgroup(t) {
    const { cols } = grid(t);
    let cg = t.querySelector(':scope > colgroup');
    if (!cg) { cg = document.createElement('colgroup'); t.prepend(cg); }
    while (cg.children.length < cols) { const c = document.createElement('col'); cg.appendChild(c); }
    while (cg.children.length > cols) cg.lastElementChild.remove();
    return cg;
  }
  function evenCols(t) { const cg = ensureColgroup(t); const n = cg.children.length; [...cg.children].forEach(c => { c.style.width = (100 / n).toFixed(3) + '%'; c.removeAttribute('width'); }); }
  function newCell(like) { const td = document.createElement(like && like.tagName === 'TH' ? 'th' : 'td'); td.innerHTML = '<p><br></p>'; if (like) { if (like.style.backgroundColor) td.style.backgroundColor = like.style.backgroundColor; if (like.style.verticalAlign) td.style.verticalAlign = like.style.verticalAlign; } return td; }

  D.currentCell = () => { const r = D.range() || D.lastRange(); return r ? D.closest(r.startContainer, 'td,th') : null; };
  D.currentTable = () => { const c = D.currentCell() || $('#doc td.cell-sel'); return c ? c.closest('table') : null; };
  D.selectedCells = () => {
    const sel = $$('#doc .cell-sel');
    if (sel.length) return sel;
    const c = D.currentCell(); return c ? [c] : [];
  };
  function clearCellSel() { $$('#doc .cell-sel').forEach(c => c.classList.remove('cell-sel')); }
  D.clearCellSel = clearCellSel;

  function rowOps(t, cell, where) {
    const g = grid(t); const p = pos(g, cell); const r = where === 'above' ? p.r : p.r + (cell.rowSpan || 1) - 1;
    const refRow = t.rows[r]; const tr = document.createElement('tr');
    for (let c = 0; c < g.cols; c++) {
      const x = g.g[r][c];
      const spansAcross = where === 'above' ? (x && pos(g, x).r < r) : (x && pos(g, x).r + (x.rowSpan || 1) - 1 > r);
      if (spansAcross) { if (pos(g, x).c === c) x.rowSpan = (x.rowSpan || 1) + 1; c += (x.colSpan || 1) - 1; continue; }
      const nc = newCell(x); if (x && x.colSpan > 1) { nc.colSpan = x.colSpan; c += x.colSpan - 1; }
      tr.appendChild(nc);
    }
    if (where === 'above') refRow.before(tr); else refRow.after(tr);
    return tr;
  }
  function colOps(t, cell, where) {
    const g = grid(t); const p = pos(g, cell); const c = where === 'left' ? p.c : p.c + (cell.colSpan || 1) - 1;
    const done = new Set();
    for (let r = 0; r < g.rows; r++) {
      const x = g.g[r][c]; if (!x) continue;
      const xp = pos(g, x);
      const spansAcross = where === 'left' ? xp.c < c : xp.c + (x.colSpan || 1) - 1 > c;
      if (spansAcross) { if (!done.has(x)) { x.colSpan = (x.colSpan || 1) + 1; done.add(x); } continue; }
      if (xp.r !== r) continue; // covered by rowspan from above: the new cell belongs to the origin row only
      const nc = newCell(x); if (x.rowSpan > 1) nc.rowSpan = x.rowSpan;
      if (where === 'left') x.before(nc); else x.after(nc);
    }
    const cg = ensureColgroup(t); evenCols(t);
  }
  function deleteRows(t, cells) {
    const g = grid(t); const rows = new Set();
    cells.forEach(cell => { const p = pos(g, cell); for (let i = 0; i < (cell.rowSpan || 1); i++) rows.add(p.r + i); });
    if (rows.size >= g.rows) { deleteTable(t); return; }
    [...rows].sort((a, b) => b - a).forEach(r => {
      const g2 = grid(t); const tr = t.rows[r];
      for (let c = 0; c < g2.cols; c++) {
        const x = g2.g[r][c]; if (!x) continue; const xp = pos(g2, x);
        if ((x.rowSpan || 1) > 1) {
          if (xp.r === r) { // origin row deleted: move cell down
            const below = t.rows[r + 1]; if (below) { x.rowSpan = x.rowSpan - 1; const ref = [...below.cells].find(bc => pos(g2, bc).c > c); below.insertBefore(x, ref || null); }
          } else if (xp.c === c) x.rowSpan = x.rowSpan - 1;
        }
        c += (x.colSpan || 1) - 1;
      }
      tr.remove();
    });
    caretTo(t.rows[Math.min([...rows][0], t.rows.length - 1)]?.cells[0]);
  }
  function deleteCols(t, cells) {
    const g = grid(t); const cols = new Set();
    cells.forEach(cell => { const p = pos(g, cell); for (let i = 0; i < (cell.colSpan || 1); i++) cols.add(p.c + i); });
    if (cols.size >= g.cols) { deleteTable(t); return; }
    [...cols].sort((a, b) => b - a).forEach(c => {
      const g2 = grid(t); const seen = new Set();
      for (let r = 0; r < g2.rows; r++) { const x = g2.g[r][c]; if (!x || seen.has(x)) continue; seen.add(x); if ((x.colSpan || 1) > 1) x.colSpan -= 1; else x.remove(); }
    });
    $$(':scope > tbody > tr, :scope > thead > tr, :scope > tr', t).forEach(tr => { if (!tr.cells.length) tr.remove(); });
    evenCols(t);
    caretTo(t.rows[0]?.cells[0]);
  }
  function deleteTable(t) {
    const p = document.createElement('p'); p.innerHTML = '<br>';
    const next = t.nextElementSibling;
    if (next && /^(P|H\d)$/.test(next.tagName)) { t.remove(); D.setCaret(next, 0); }
    else { t.replaceWith(p); D.setCaret(p, 0); }
  }
  function caretTo(cell) { if (!cell) return; const p = cell.querySelector('p,h1,h2,h3,li') || cell; D.setCaret(p, 0); }
  D.caretToCell = caretTo;

  function merge(t, cells) {
    if (cells.length < 2) { BTS.toast('Select two or more cells to merge — drag across them.'); return false; }
    const g = grid(t); let r0 = 1e9, c0 = 1e9, r1 = -1, c1 = -1;
    cells.forEach(cell => { const p = pos(g, cell); r0 = Math.min(r0, p.r); c0 = Math.min(c0, p.c); r1 = Math.max(r1, p.r + (cell.rowSpan || 1) - 1); c1 = Math.max(c1, p.c + (cell.colSpan || 1) - 1); });
    // rectangle must not cut through merged cells
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) { const x = g.g[r][c]; const p = pos(g, x); if (p.r < r0 || p.c < c0 || p.r + (x.rowSpan || 1) - 1 > r1 || p.c + (x.colSpan || 1) - 1 > c1) { BTS.toast('The selection must form a rectangle.', { error: true }); return false; } }
    const target = g.g[r0][c0]; const seen = new Set([target]);
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) {
      const x = g.g[r][c]; if (seen.has(x)) continue; seen.add(x);
      const hasText = x.textContent.trim() || x.querySelector('img,table');
      if (hasText) { [...x.childNodes].forEach(n => target.appendChild(n)); }
      x.remove();
    }
    target.rowSpan = r1 - r0 + 1; target.colSpan = c1 - c0 + 1;
    // drop empty paragraphs created by merge if there is content
    const ps = [...target.children]; if (ps.length > 1) ps.forEach(p => { if (p.tagName === 'P' && !p.textContent.trim() && !p.querySelector('img') && target.children.length > 1) p.remove(); });
    $$('tr', t).forEach(tr => { if (!tr.cells.length) tr.remove(); });
    clearCellSel(); caretTo(target);
    return true;
  }
  async function split(t, cell) {
    const cs = cell.colSpan || 1, rs = cell.rowSpan || 1;
    let cols = cs, rows = rs;
    if (cs === 1 && rs === 1) {
      const sv = D.saveSel();
      const v = await BTS.dialog({ title: 'Split cell', width: 360, body: `<div class="field-row"><div class="field"><label>Number of columns</label><input id="sp-c" type="number" min="1" max="20" value="2"></div><div class="field"><label>Number of rows</label><input id="sp-r" type="number" min="1" max="50" value="1"></div></div>`, buttons: [{ label: 'Cancel', value: null }, { label: 'Split', primary: true, value: (r) => ({ c: +$('#sp-c', r).value || 1, r: +$('#sp-r', r).value || 1 }) }] });
      D.restoreSel(sv);
      if (!v) return;
      D.change(() => splitPlain(t, cell, Math.max(1, v.c), Math.max(1, v.r)));
      return;
    }
    D.change(() => {
      const g = grid(t); const p = pos(g, cell);
      cell.colSpan = 1; cell.rowSpan = 1;
      for (let j = 1; j < cs; j++) cell.after(newCell(cell));
      for (let i = 1; i < rs; i++) {
        const tr = t.rows[p.r + i];
        const g2 = grid(t); let ref = null;
        for (const c of tr.cells) { if (pos(g2, c).c > p.c) { ref = c; break; } }
        for (let j = 0; j < cs; j++) tr.insertBefore(newCell(cell), ref);
      }
    });
  }
  function splitPlain(t, cell, nc, nr) {
    const g = grid(t); const p = pos(g, cell);
    if (nc > 1) {
      // add nc-1 columns after p.c: other rows' cells in that column get colSpan += nc-1
      const seen = new Set();
      for (let r = 0; r < g.rows; r++) { const x = g.g[r][p.c]; if (x === cell || seen.has(x)) continue; seen.add(x); x.colSpan = (x.colSpan || 1) + nc - 1; }
      for (let j = 1; j < nc; j++) cell.after(newCell(cell));
      const cg = ensureColgroup(t); evenCols(t);
    }
    if (nr > 1) {
      const rowCells = [cell]; let n = cell.nextElementSibling; for (let j = 1; j < nc; j++) { rowCells.push(n); n = n && n.nextElementSibling; }
      const g2 = grid(t); const tr = cell.parentElement;
      for (let i = 1; i < nr; i++) {
        const ntr = document.createElement('tr');
        for (let c = 0; c < g2.cols; c++) {
          const x = g2.g[p.r][c];
          if (rowCells.includes(x)) { ntr.appendChild(newCell(x)); continue; }
          if (pos(g2, x).c === c && pos(g2, x).r === p.r && i === 1) x.rowSpan = (x.rowSpan || 1) + nr - 1;
          else if (pos(g2, x).r < p.r && pos(g2, x).c === c && i === 1) x.rowSpan = (x.rowSpan || 1) + nr - 1;
          c += (x.colSpan || 1) - 1;
        }
        tr.after(ntr);
      }
    }
  }

  D.tableOp = async (op, arg) => {
    const t = D.currentTable(); if (!t) { BTS.toast('Put the cursor in a table first.'); return; }
    const cells = D.selectedCells(); const cell = cells[0] || D.currentCell() || t.rows[0].cells[0];
    if (op === 'split') return split(t, cell);
    D.change(() => {
      switch (op) {
        case 'row-above': caretTo(rowOps(t, cells[0], 'above').cells[0]); break;
        case 'row-below': caretTo(rowOps(t, cells[cells.length - 1], 'below').cells[0]); break;
        case 'col-left': colOps(t, cells[0], 'left'); break;
        case 'col-right': colOps(t, cells[cells.length - 1], 'right'); break;
        case 'del-row': deleteRows(t, cells); break;
        case 'del-col': deleteCols(t, cells); break;
        case 'del-table': deleteTable(t); break;
        case 'del-cells': cells.forEach(c => { c.innerHTML = '<p><br></p>'; }); break;
        case 'merge': merge(t, cells); break;
        case 'style': STYLE_IDS.forEach(s => t.classList.remove(s)); t.classList.add(arg); t.classList.remove('no-borders'); break;
        case 'header-row': t.classList.toggle('t-noheader', !arg); break;
        case 'shading': cells.forEach(c => { c.style.backgroundColor = arg === 'none' ? '' : arg; if (!c.getAttribute('style')) c.removeAttribute('style'); }); break;
        case 'borders': applyBorders(t, cells, arg); break;
        case 'align': cells.forEach(c => { $$('p,h1,h2,h3,li', c).forEach(p => p.style.textAlign = arg === 'left' ? '' : arg); if (!c.querySelector('p,h1,h2,h3,li')) c.style.textAlign = arg; }); break;
        case 'valign': cells.forEach(c => c.style.verticalAlign = arg); break;
        case 'autofit-content': { t.style.width = 'auto'; t.style.tableLayout = 'auto'; t.classList.remove('tbl'); t.classList.add('tbl-auto'); const cg = t.querySelector(':scope > colgroup'); if (cg) [...cg.children].forEach(c => c.style.width = ''); requestAnimationFrame(() => { freezeWidths(t, false); }); break; }
        case 'autofit-window': t.style.width = ''; t.style.tableLayout = ''; t.classList.add('tbl'); t.classList.remove('tbl-auto'); evenColsKeepRatio(t); break;
        case 'distribute': evenCols(t); t.style.tableLayout = ''; break;
      }
    });
    D.updateTableTools();
  };
  function freezeWidths(t) {
    // after auto layout, record measured px widths so later edits keep them and export can use them
    const cg = ensureColgroup(t); const g = grid(t); const widths = new Array(g.cols).fill(0);
    for (let c = 0; c < g.cols; c++) { for (let r = 0; r < g.rows; r++) { const x = g.g[r][c]; if (x && (x.colSpan || 1) === 1) { widths[c] = Math.max(widths[c], x.getBoundingClientRect().width / (D.zoom / 100)); } } }
    [...cg.children].forEach((col, i) => { if (widths[i]) col.style.width = Math.round(widths[i]) + 'px'; });
    t.style.width = Math.round(widths.reduce((a, b) => a + b, 0)) + 'px'; t.style.tableLayout = 'fixed';
  }
  function evenColsKeepRatio(t) {
    const cg = ensureColgroup(t); const ws = [...cg.children].map(c => D.toPx(c.style.width) || 0); const sum = ws.reduce((a, b) => a + b, 0);
    if (!sum || cg.children[0].style.width.endsWith('%')) { if (!cg.children[0].style.width) evenCols(t); return; }
    [...cg.children].forEach((c, i) => c.style.width = (ws[i] / sum * 100).toFixed(3) + '%');
  }
  function applyBorders(t, cells, kind) {
    if (kind === 'none') { t.classList.add('no-borders'); $$('td,th', t).forEach(c => { c.style.border = ''; }); return; }
    if (kind === 'all') { t.classList.remove('no-borders'); $$('td,th', t).forEach(c => { c.style.border = ''; }); return; }
    if (kind === 'outside') { t.classList.add('no-borders'); const g = grid(t); const line = '1px solid #3c4250'; $$('td,th', t).forEach(c => { c.style.border = ''; }); for (let r = 0; r < g.rows; r++) for (let c = 0; c < g.cols; c++) { const x = g.g[r][c]; if (r === 0) x.style.borderTop = line; if (r === g.rows - 1) x.style.borderBottom = line; if (c === 0) x.style.borderLeft = line; if (c === g.cols - 1) x.style.borderRight = line; } return; }
    if (kind === 'thick') { t.classList.remove('no-borders'); $$('td,th', t).forEach(c => { c.style.border = '2px solid #1f2533'; }); return; }
    if (kind === 'cell-box') { cells.forEach(c => c.style.border = '1.5px solid #1f2533'); return; }
  }

  /* ---------- contextual tab ---------- */
  D.updateTableTools = () => {
    const t = D.currentTable(); const tab = $('#tab-table');
    const show = !!t;
    if (tab.classList.contains('hidden') === show) { tab.classList.toggle('hidden', !show); if (!show && BTS.activeTab === 'table') BTS.showTab('home'); }
    if (t) {
      $('#chk-header-row').checked = !t.classList.contains('t-noheader') && /t-(header|banded|light|grid-accent|dark|minimal)/.test(t.className);
      $$('#tstyles-gallery .gi').forEach(g => g.classList.toggle('on', t.classList.contains(g.dataset.style)));
    }
  };
  D.renderTableStyles = () => {
    const prev = {
      't-plain': ['#fff', '#fff', '#fff', '#555'], 't-header': ['#3C4FE0', '#fff', '#fff', '#888'], 't-banded': ['#3C4FE0', '#fff', '#e3e6fb', '#c5cbf5'],
      't-grid-accent': ['#dfe3fb', '#fff', '#fff', '#3C4FE0'], 't-light': ['#fff', '#fff', '#fff', '#bbb'], 't-dark': ['#1f2533', '#fff', '#eef0f4', '#b7bdc8'], 't-minimal': ['#fff', '#fff', '#f2f3f5', 'transparent'],
    };
    $('#tstyles-gallery').innerHTML = D.TABLE_STYLES.map(s => { const [h, a, b, bd] = prev[s.id]; const rows = [h, a, b, a].map((bg, i) => `<tr>${'<td></td>'.repeat(4).replace(/<td>/g, `<td style="background:${bg};border:1px solid ${bd}${s.id === 't-minimal' && i === 0 ? ';border-bottom:2px solid #3C4FE0' : ''}">`)}</tr>`).join(''); return `<button class="gi keepfocus" data-cmd="tbl-style" data-style="${s.id}" title="${s.name}"><table>${rows}</table></button>`; }).join('');
  };

  /* ---------- keyboard: Tab navigation ---------- */
  D.tableKey = (e) => {
    if (e.key !== 'Tab') return false;
    const cell = D.currentCell(); if (!cell) return false;
    const r = D.range(); if (D.closest(r.startContainer, 'li')) return false; // lists inside cells indent instead
    e.preventDefault();
    const t = cell.closest('table'); const all = $$('td,th', t).filter(c => c.closest('table') === t);
    const i = all.indexOf(cell);
    if (e.shiftKey) { if (i > 0) selectCell(all[i - 1]); return true; }
    if (i < all.length - 1) selectCell(all[i + 1]);
    else D.change(() => { const tr = rowOps(t, cell, 'below'); caretTo(tr.cells[0]); });
    return true;
  };
  function selectCell(c) { const r = document.createRange(); r.selectNodeContents(c); const s = getSelection(); s.removeAllRanges(); s.addRange(r); if (!c.textContent.trim()) caretTo(c); }

  /* ---------- mouse: drag to select cells, column resizing ---------- */
  function wireMouse() {
    let anchor = null, active = false;
    D.doc.addEventListener('mousedown', e => {
      if (e.button !== 0) return;
      clearCellSel(); D.doc.classList.remove('cell-selecting');
      anchor = e.target.closest('td,th'); active = !!anchor;
      if (e.shiftKey && anchor) { const cur = D.currentCell(); if (cur && cur !== anchor && cur.closest('table') === anchor.closest('table')) { e.preventDefault(); markRect(cur, anchor); active = false; } }
    });
    D.doc.addEventListener('mouseover', e => {
      if (!active || !(e.buttons & 1)) return;
      const c = e.target.closest('td,th'); if (!c || c === anchor) { if (c === anchor) { clearCellSel(); D.doc.classList.remove('cell-selecting'); } return; }
      if (c.closest('table') !== anchor.closest('table')) return;
      markRect(anchor, c);
    });
    addEventListener('mouseup', () => { if (active && $('#doc .cell-sel')) { D.updateTableTools(); } active = false; });
    // column resizing via hover on cell edges
    let rz = null;
    D.doc.addEventListener('mousemove', e => {
      if (rz) return;
      const cell = e.target.closest && e.target.closest('td,th'); if (!cell) { D.doc.style.cursor = ''; return; }
      const r = cell.getBoundingClientRect(); const near = Math.abs(e.clientX - r.right) < 4;
      D.doc.style.cursor = near ? 'col-resize' : '';
    });
    D.doc.addEventListener('mousedown', e => {
      if (D.doc.style.cursor !== 'col-resize') return;
      const cell = e.target.closest('td,th'); if (!cell) return;
      e.preventDefault(); e.stopPropagation(); clearCellSel(); active = false;
      const t = cell.closest('table'); const g = grid(t); const p = pos(g, cell); const ci = p.c + (cell.colSpan || 1) - 1;
      const cg = ensureColgroup(t); const tw = t.getBoundingClientRect().width / (D.zoom / 100);
      const widths = []; for (let c = 0; c < g.cols; c++) { let w = 0; for (let r = 0; r < g.rows; r++) { const x = g.g[r][c]; if (x && (x.colSpan || 1) === 1) { w = x.getBoundingClientRect().width / (D.zoom / 100); break; } } widths.push(w || tw / g.cols); }
      D.checkpoint();
      rz = { t, cg, ci, x0: e.clientX, widths, tw };
    }, true);
    addEventListener('mousemove', e => {
      if (!rz) return;
      const dx = (e.clientX - rz.x0) / (D.zoom / 100); const w = [...rz.widths];
      if (rz.ci < w.length - 1) { const a = Math.max(20, w[rz.ci] + dx); const b = Math.max(20, w[rz.ci + 1] - (a - w[rz.ci])); if (b > 20 || dx < 0) { w[rz.ci + 1] = b; w[rz.ci] = a; } }
      else { w[rz.ci] = Math.max(20, w[rz.ci] + dx); }
      const sum = w.reduce((a, b) => a + b, 0);
      if (rz.ci < w.length - 1 && !rz.t.style.width) [...rz.cg.children].forEach((c, i) => c.style.width = (w[i] / sum * 100).toFixed(3) + '%');
      else { [...rz.cg.children].forEach((c, i) => c.style.width = Math.round(w[i]) + 'px'); rz.t.style.width = Math.round(sum) + 'px'; }
    });
    addEventListener('mouseup', () => { if (rz) { rz = null; D.afterChange(); } });
  }
  function markRect(a, b) {
    const t = a.closest('table'); const g = grid(t); const pa = pos(g, a), pb = pos(g, b);
    let r0 = Math.min(pa.r, pb.r), r1 = Math.max(pa.r + a.rowSpan - 1, pb.r + b.rowSpan - 1), c0 = Math.min(pa.c, pb.c), c1 = Math.max(pa.c + a.colSpan - 1, pb.c + b.colSpan - 1);
    // expand to include merged cells fully
    let changed = true; while (changed) { changed = false; for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) { const x = g.g[r][c]; if (!x) continue; const p = pos(g, x); if (p.r < r0) { r0 = p.r; changed = true; } if (p.c < c0) { c0 = p.c; changed = true; } if (p.r + x.rowSpan - 1 > r1) { r1 = p.r + x.rowSpan - 1; changed = true; } if (p.c + x.colSpan - 1 > c1) { c1 = p.c + x.colSpan - 1; changed = true; } } }
    clearCellSel();
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) g.g[r][c] && g.g[r][c].classList.add('cell-sel');
    D.doc.classList.add('cell-selecting');
  }
  D.markCellRect = markRect;

  D.initTables = () => { wireMouse(); D.renderTableStyles(); };
})();
