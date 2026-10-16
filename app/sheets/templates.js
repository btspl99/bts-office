/* BTS Sheets — templates for File › New (real formulas & formatting) */
(function () {
  const T = window.STemplates = {};
  const { key } = SU;
  const ACC = '#0F8A6A';
  const thin = { style: 'thin', color: '#BFC7CF' };
  const box = { bT: thin, bB: thin, bL: thin, bR: thin };
  const H = { b: true, color: '#FFFFFF', fill: ACC, h: 'center', v: 'middle', ...box };
  const INR = SU.currencyFormat('₹');

  /** spec: {name, rows:[[cell...]], start:{r,c}, styles:{'A1:B2': style}, colW:{A:px}, rowH:{1:px}, merges:['A1:D1'], freeze:{r,c}, cf:[], validations:[], charts:[], tabColor} — cell may be value or {v, s} */
  function sheet(spec) {
    const data = []; const styles = new Map();
    const put = (r, c, cell) => { if (cell === undefined || cell === null) return; let v = cell, s = null; if (typeof cell === 'object' && !Array.isArray(cell)) { v = cell.v; s = cell.s; } if (v !== undefined && v !== null && v !== '') (data[r] = data[r] || [])[c] = v; if (s) styles.set(key(r, c), { ...(styles.get(key(r, c)) || {}), ...s }); };
    const s0 = spec.start || { r: 0, c: 0 };
    (spec.rows || []).forEach((row, i) => row && row.forEach((cell, j) => put(s0.r + i, s0.c + j, cell)));
    for (const [ref, st] of Object.entries(spec.styles || {})) { const g = SU.parseRange(ref); for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) styles.set(key(r, c), { ...(styles.get(key(r, c)) || {}), ...st }); }
    const colW = new Map(Object.entries(spec.colW || {}).map(([k, v]) => [SU.colIndex(k), v]));
    const rowH = new Map(Object.entries(spec.rowH || {}).map(([k, v]) => [+k - 1, v]));
    const merges = (spec.merges || []).map(SU.parseRange);
    return { name: spec.name, data, meta: { styles, colW, rowH, merges, freeze: spec.freeze || { r: 0, c: 0 }, cf: (spec.cf || []).map((x, i) => ({ id: 'cf' + i, ...x, range: SU.parseRange(x.range) })), validations: (spec.validations || []).map(x => ({ ...x, range: SU.parseRange(x.range) })), charts: (spec.charts || []).map((x, i) => ({ id: 'tch' + i + Date.now().toString(36), ...x, range: SU.parseRange(x.range) })), tabColor: spec.tabColor || null, showGrid: spec.showGrid !== false, tables: (spec.tables || []).map((t, i) => ({ id: 'tt' + i, name: 'Table' + (i + 1), header: true, banded: true, total: false, filterBtn: true, ...t, range: SU.parseRange(t.range) })) } };
  }
  T.sheet = sheet;

  /* ---- Monthly budget ---- */
  function budget() {
    const inc = [['Salary / stipend', 25000, 25000], ['Part-time work', 6000, 5400], ['Pocket money', 3000, 3000], ['Other', 0, 750]];
    const exp = [['Rent / hostel', 9000, 9000], ['Food & groceries', 6000, 6450], ['Transport', 1500, 1320], ['Mobile & internet', 600, 599], ['Books & stationery', 1200, 860], ['Fees & courses', 2500, 2500], ['Entertainment', 1500, 2100], ['Savings', 5000, 4500], ['Miscellaneous', 1000, 1240]];
    const rows = [];
    rows.push([{ v: 'Monthly Budget', s: { sz: 20, b: true, color: '#0B4F3E' } }]);
    rows.push([{ v: 'Month:', s: { color: '#5c6577' } }, { v: SU.dateToSerial(new Date().getFullYear(), new Date().getMonth() + 1, 1), s: { nf: 'mmmm yyyy', b: true, h: 'left' } }]);
    rows.push([]);
    rows.push([{ v: 'Income', s: H }, { v: 'Planned', s: H }, { v: 'Actual', s: H }, { v: 'Difference', s: H }]);
    const i0 = rows.length; inc.forEach((x, i) => { const r = i0 + i + 1; rows.push([{ v: x[0], s: box }, { v: x[1], s: { nf: INR, ...box } }, { v: x[2], s: { nf: INR, ...box } }, { v: `=C${r}-B${r}`, s: { nf: INR, ...box } }]); });
    const it = rows.length + 1; rows.push([{ v: 'Total income', s: { b: true, fill: '#E3F3EE', ...box } }, { v: `=SUM(B${i0 + 1}:B${it - 1})`, s: { b: true, nf: INR, fill: '#E3F3EE', ...box } }, { v: `=SUM(C${i0 + 1}:C${it - 1})`, s: { b: true, nf: INR, fill: '#E3F3EE', ...box } }, { v: `=C${it}-B${it}`, s: { b: true, nf: INR, fill: '#E3F3EE', ...box } }]);
    rows.push([]);
    rows.push([{ v: 'Expenses', s: H }, { v: 'Planned', s: H }, { v: 'Actual', s: H }, { v: 'Difference', s: H }, { v: '% of income', s: H }]);
    const e0 = rows.length; exp.forEach((x, i) => { const r = e0 + i + 1; rows.push([{ v: x[0], s: box }, { v: x[1], s: { nf: INR, ...box } }, { v: x[2], s: { nf: INR, ...box } }, { v: `=B${r}-C${r}`, s: { nf: INR, ...box } }, { v: `=IFERROR(C${r}/$C$${it},0)`, s: { nf: '0.0%', ...box } }]); });
    const et = rows.length + 1; rows.push([{ v: 'Total expenses', s: { b: true, fill: '#E3F3EE', ...box } }, { v: `=SUM(B${e0 + 1}:B${et - 1})`, s: { b: true, nf: INR, fill: '#E3F3EE', ...box } }, { v: `=SUM(C${e0 + 1}:C${et - 1})`, s: { b: true, nf: INR, fill: '#E3F3EE', ...box } }, { v: `=B${et}-C${et}`, s: { b: true, nf: INR, fill: '#E3F3EE', ...box } }, { v: `=IFERROR(C${et}/$C$${it},0)`, s: { b: true, nf: '0.0%', fill: '#E3F3EE', ...box } }]);
    rows.push([]);
    rows.push([{ v: 'Balance (income − expenses)', s: { b: true, sz: 13 } }, { v: `=B${it}-B${et}`, s: { b: true, sz: 13, nf: INR } }, { v: `=C${it}-C${et}`, s: { b: true, sz: 13, nf: INR } }]);
    const bal = rows.length;
    return { sheets: [sheet({ name: 'Budget', rows, colW: { A: 210, B: 110, C: 110, D: 110, E: 100 }, rowH: { 1: 36 }, freeze: { r: 0, c: 0 }, cf: [{ type: 'lt', v1: 0, range: `D${e0 + 1}:D${et}`, fmt: { fill: '#FFC7CE', color: '#9C0006' } }, { type: 'bar', range: `C${e0 + 1}:C${et - 1}`, color: '#63BE7B' }, { type: 'lt', v1: 0, range: `B${bal}:C${bal}`, fmt: { color: '#C00000' } }], charts: [{ type: 'doughnut', range: `A${e0}:A${et - 1}`.replace(/A(\d+):A(\d+)/, `A$1:A$2`), x: 660, y: 60, w: 430, h: 300, title: 'Where the money went', sheet: null }] })] };
  }
  function fixBudgetChart(b) { const s = b.sheets[0]; const ch = s.meta.charts[0]; ch.range = { r0: ch.range.r0, c0: 0, r1: ch.range.r1, c1: 2 }; ch.pick = [1]; return b; }

  /* ---- Student marks sheet ---- */
  function marks() {
    const names = ['Aarav Sharma', 'Diya Patel', 'Ishaan Kumar', 'Ananya Iyer', 'Rohan Deshmukh', 'Sneha Reddy', 'Kabir Singh', 'Meera Nair', 'Arjun Mehta', 'Priya Joshi', 'Vihaan Gupta', 'Saanvi Rao'];
    const subj = ['Maths', 'Science', 'English', 'Hindi', 'Social Sc.'];
    const seed = (i, j) => 45 + ((i * 37 + j * 53 + i * j * 11) % 54);
    const rows = [];
    rows.push([{ v: 'Class 10-A — Term 1 Results', s: { sz: 18, b: true, color: '#0B4F3E' } }]);
    rows.push([{ v: 'Max marks per subject:', s: { color: '#5c6577' } }, null, { v: 100, s: { b: true, h: 'left' } }, { v: 'Pass marks:', s: { color: '#5c6577' } }, null, { v: 33, s: { b: true, h: 'left' } }]);
    rows.push([]);
    const hdr = ['Roll', 'Student name', ...subj, 'Total', 'Percentage', 'Grade', 'Result', 'Rank'];
    rows.push(hdr.map(h => ({ v: h, s: { ...H, wrap: true } })));
    const r0 = rows.length + 1; const rN = r0 + names.length - 1;
    names.forEach((n, i) => {
      const r = r0 + i; const row = [{ v: i + 1, s: { h: 'center', ...box } }, { v: n, s: box }];
      subj.forEach((_, j) => row.push({ v: seed(i, j), s: { h: 'center', ...box } }));
      row.push({ v: `=SUM(C${r}:G${r})`, s: { b: true, h: 'center', ...box } });
      row.push({ v: `=H${r}/(COUNT(C${r}:G${r})*$C$2)`, s: { nf: '0.0%', h: 'center', ...box } });
      row.push({ v: `=IF(I${r}>=0.9,"A+",IF(I${r}>=0.75,"A",IF(I${r}>=0.6,"B",IF(I${r}>=0.45,"C",IF(I${r}>=0.33,"D","E")))))`, s: { b: true, h: 'center', ...box } });
      row.push({ v: `=IF(COUNTIF(C${r}:G${r},"<"&$F$2)=0,"PASS","FAIL")`, s: { h: 'center', ...box } });
      row.push({ v: `=RANK(H${r},$H$${r0}:$H$${rN})`, s: { h: 'center', ...box } });
      rows.push(row);
    });
    const stat = (label, fn) => [null, { v: label, s: { b: true, h: 'right', fill: '#E3F3EE', ...box } }, ...subj.map((_, j) => { const L = SU.colName(2 + j); return { v: `=${fn}(${L}${r0}:${L}${rN})`, s: { nf: fn === 'AVERAGE' ? '0.0' : 'General', h: 'center', fill: '#E3F3EE', ...box } }; }), { v: `=${fn}(H${r0}:H${rN})`, s: { nf: fn === 'AVERAGE' ? '0.0' : 'General', h: 'center', b: true, fill: '#E3F3EE', ...box } }, { v: `=${fn}(I${r0}:I${rN})`, s: { nf: '0.0%', h: 'center', fill: '#E3F3EE', ...box } }];
    rows.push(stat('Class average', 'AVERAGE')); rows.push(stat('Highest', 'MAX')); rows.push(stat('Lowest', 'MIN'));
    rows.push([]);
    rows.push([null, { v: 'Students passed', s: { b: true } }, { v: `=COUNTIF(K${r0}:K${rN},"PASS")`, s: { b: true, h: 'center' } }, { v: 'of', s: { h: 'center', color: '#5c6577' } }, { v: `=COUNTA(B${r0}:B${rN})`, s: { b: true, h: 'center' } }]);
    rows.push([null, { v: 'Topper', s: { b: true } }, { v: `=INDEX(B${r0}:B${rN},MATCH(MAX(H${r0}:H${rN}),H${r0}:H${rN},0))`, s: { b: true, color: ACC } }]);
    return { sheets: [sheet({ name: 'Marks', rows, colW: { A: 46, B: 150, C: 64, D: 64, E: 64, F: 64, G: 70, H: 62, I: 86, J: 56, K: 64, L: 50 }, rowH: { 1: 34, 4: 36 }, freeze: { r: 4, c: 2 }, merges: ['A1:L1'],
      cf: [{ type: 'lt', v1: 33, range: `C${r0}:G${rN}`, fmt: { fill: '#FFC7CE', color: '#9C0006' } }, { type: 'text', v1: 'FAIL', range: `K${r0}:K${rN}`, fmt: { fill: '#FFC7CE', color: '#9C0006', b: true } }, { type: 'text', v1: 'PASS', range: `K${r0}:K${rN}`, fmt: { color: '#006100' } }, { type: 'scale', range: `I${r0}:I${rN}`, colors: ['#F8696B', '#FFEB84', '#63BE7B'] }, { type: 'eq', v1: 1, range: `L${r0}:L${rN}`, fmt: { fill: '#FFEB9C', color: '#9C5700', b: true } }],
      validations: [{ range: `C${r0}:G${rN}`, type: 'whole', op: 'between', v1: 0, v2: 100, errorMsg: 'Marks must be a whole number from 0 to 100.', errorTitle: 'Invalid marks', ignoreBlank: true }],
      charts: [{ type: 'column', range: `B${r0 - 1}:B${rN}`, x: 0, y: 0, w: 520, h: 280, title: 'Total marks by student', sheet: null }] })] };
  }
  function fixMarksChart(m) {
    const s = m.sheets[0]; const ch = s.meta.charts[0]; const r0 = ch.range.r0, r1 = ch.range.r1;
    ch.range = { r0, c0: 1, r1, c1: 7 }; ch.pick = [5]; // names as labels, Total series only
    let y = 0; for (let r = 0; r < r1 + 8; r++) y += s.meta.rowH.get(r) || SU.DEF_H; ch.y = y + 10; ch.x = 56; ch.w = 700; ch.legend = false; return m;
  }
  /* ---- Attendance tracker ---- */
  function attendance() {
    const names = ['Aarav Sharma', 'Diya Patel', 'Ishaan Kumar', 'Ananya Iyer', 'Rohan Deshmukh', 'Sneha Reddy', 'Kabir Singh', 'Meera Nair', 'Arjun Mehta', 'Priya Joshi'];
    const days = 22; const rows = [];
    rows.push([{ v: 'Attendance Register', s: { sz: 18, b: true, color: '#0B4F3E' } }]);
    rows.push([{ v: 'Month starting:', s: { color: '#5c6577' } }, { v: SU.dateToSerial(new Date().getFullYear(), new Date().getMonth() + 1, 1), s: { nf: 'dd-mmm-yyyy', b: true, h: 'left' } }, null, null, null, null, { v: 'P = Present   A = Absent   L = Leave', s: { color: '#5c6577', i: true } }]);
    rows.push([]);
    const hdr = [{ v: 'Student', s: H }]; for (let d = 1; d <= days; d++) hdr.push({ v: d, s: { ...H, sz: 10 } });
    hdr.push({ v: 'Present', s: H }, { v: 'Absent', s: H }, { v: 'Leave', s: H }, { v: 'Attendance %', s: { ...H, wrap: true } });
    rows.push(hdr);
    const r0 = rows.length + 1;
    names.forEach((n, i) => {
      const r = r0 + i; const row = [{ v: n, s: box }];
      for (let d = 0; d < days; d++) { const x = (i * 7 + d * 13 + i * d) % 23; row.push({ v: x === 0 ? 'A' : x === 5 ? 'L' : 'P', s: { h: 'center', sz: 10, ...box } }); }
      const L1 = SU.colName(1), L2 = SU.colName(days);
      row.push({ v: `=COUNTIF(${L1}${r}:${L2}${r},"P")`, s: { h: 'center', b: true, ...box } }, { v: `=COUNTIF(${L1}${r}:${L2}${r},"A")`, s: { h: 'center', ...box } }, { v: `=COUNTIF(${L1}${r}:${L2}${r},"L")`, s: { h: 'center', ...box } }, { v: `=IFERROR(${SU.colName(days + 1)}${r}/COUNTA(${L1}${r}:${L2}${r}),0)`, s: { nf: '0%', h: 'center', b: true, ...box } });
      rows.push(row);
    });
    const rN = r0 + names.length - 1; const pc = SU.colName(days + 4);
    rows.push([]); rows.push([{ v: 'Class average attendance', s: { b: true } }, ...Array(days - 1).fill(null), null, null, null, { v: `=AVERAGE(${pc}${r0}:${pc}${rN})`, s: { nf: '0.0%', b: true, h: 'center' } }]);
    const colW = { A: 150 }; for (let d = 1; d <= days; d++) colW[SU.colName(d)] = 28; colW[SU.colName(days + 1)] = 62; colW[SU.colName(days + 2)] = 58; colW[SU.colName(days + 3)] = 52; colW[SU.colName(days + 4)] = 88;
    return { sheets: [sheet({ name: 'Attendance', rows, colW, rowH: { 1: 34, 4: 34 }, freeze: { r: 4, c: 1 }, merges: ['B2:E2'], cf: [{ type: 'eq', v1: 'A', range: `B${r0}:${SU.colName(days)}${rN}`, fmt: { fill: '#FFC7CE', color: '#9C0006', b: true } }, { type: 'eq', v1: 'L', range: `B${r0}:${SU.colName(days)}${rN}`, fmt: { fill: '#FFEB9C', color: '#9C5700' } }, { type: 'lt', v1: 0.75, range: `${pc}${r0}:${pc}${rN}`, fmt: { color: '#C00000' } }, { type: 'bar', range: `${pc}${r0}:${pc}${rN}`, color: '#0F8A6A' }], validations: [{ range: `B${r0}:${SU.colName(days)}${rN}`, type: 'list', source: 'P,A,L', errorMsg: 'Use P, A or L.', ignoreBlank: true }] })] };
  }

  /* ---- To-do list ---- */
  function todo() {
    const tasks = [['Finish maths assignment', 'Study', 'High', 2, 'In progress'], ['Submit scholarship form', 'Admin', 'High', 0, 'Not started'], ['Read chapter 5 – Physics', 'Study', 'Medium', 3, 'Not started'], ['Group project meeting', 'College', 'Medium', 1, 'Done'], ['Pay hostel fees', 'Admin', 'High', 5, 'Not started'], ['Practice coding problems', 'Skills', 'Low', 7, 'In progress'], ['Call home', 'Personal', 'Low', 0, 'Done']];
    const today = (() => { const d = new Date(); return SU.dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); })();
    const rows = [];
    rows.push([{ v: 'My To-Do List', s: { sz: 20, b: true, color: '#0B4F3E' } }]);
    rows.push([{ v: '=COUNTIF(F4:F40,"Done")&" of "&COUNTA(B4:B40)&" tasks done"', s: { color: '#5c6577' } }, null, null, null, { v: 'Progress', s: { h: 'right', color: '#5c6577' } }, { v: '=IFERROR(COUNTIF(F4:F40,"Done")/COUNTA(B4:B40),0)', s: { nf: '0%', b: true } }]);
    rows.push(['#', 'Task', 'Category', 'Priority', 'Due date', 'Status', 'Days left'].map(h => ({ v: h, s: H })));
    tasks.forEach((t, i) => { const r = 4 + i; rows.push([{ v: i + 1, s: { h: 'center', ...box } }, { v: t[0], s: box }, { v: t[1], s: box }, { v: t[2], s: { h: 'center', ...box } }, { v: today + t[3], s: { nf: 'dd-mmm', h: 'center', ...box } }, { v: t[4], s: { h: 'center', ...box } }, { v: `=IF(F${r}="Done","",E${r}-TODAY())`, s: { h: 'center', ...box } }]); });
    for (let i = tasks.length; i < 20; i++) { const r = 4 + i; rows.push([{ v: i + 1, s: { h: 'center', ...box, color: '#9aa3b2' } }, { s: box }, { s: box }, { s: { h: 'center', ...box } }, { s: { nf: 'dd-mmm', h: 'center', ...box } }, { s: { h: 'center', ...box } }, { v: `=IF(OR(F${r}="Done",E${r}=""),"",E${r}-TODAY())`, s: { h: 'center', ...box } }]); }
    return { sheets: [sheet({ name: 'To-do', rows, colW: { A: 36, B: 260, C: 96, D: 80, E: 84, F: 100, G: 76 }, rowH: { 1: 38 }, freeze: { r: 3, c: 0 }, showGrid: false,
      cf: [{ type: 'eq', v1: 'Done', range: 'F4:F23', fmt: { fill: '#C6EFCE', color: '#006100' } }, { type: 'eq', v1: 'In progress', range: 'F4:F23', fmt: { fill: '#FFEB9C', color: '#9C5700' } }, { type: 'eq', v1: 'High', range: 'D4:D23', fmt: { color: '#C00000', b: true } }, { type: 'formula', formula: '=AND(G4<>"",G4<0)', range: 'G4:G23', fmt: { fill: '#FFC7CE', color: '#9C0006' } }],
      validations: [{ range: 'D4:D23', type: 'list', source: 'High,Medium,Low', ignoreBlank: true }, { range: 'F4:F23', type: 'list', source: 'Not started,In progress,Done', ignoreBlank: true }, { range: 'C4:C23', type: 'list', source: 'Study,College,Admin,Skills,Personal', ignoreBlank: true, errorStyle: 'warning' }] })] };
  }

  /* ---- Invoice ---- */
  function invoice() {
    const items = [['Web design – landing page', 1, 12000], ['Logo design', 1, 4500], ['Content writing (per page)', 5, 800], ['Hosting (12 months)', 1, 3600]];
    const rows = [];
    rows.push([{ v: 'INVOICE', s: { sz: 26, b: true, color: ACC } }]);
    rows.push([{ v: 'Your Name / Business', s: { b: true, sz: 12 } }, null, null, { v: 'Invoice #', s: { h: 'right', color: '#5c6577' } }, { v: 'INV-0042', s: { b: true, h: 'right' } }]);
    rows.push([{ v: 'Address line, City, State – PIN', s: { color: '#5c6577' } }, null, null, { v: 'Date', s: { h: 'right', color: '#5c6577' } }, { v: '=TODAY()', s: { nf: 'dd-mmm-yyyy', h: 'right' } }]);
    rows.push([{ v: 'GSTIN: 27ABCDE1234F1Z5   ·   +91 98765 43210', s: { color: '#5c6577' } }, null, null, { v: 'Due date', s: { h: 'right', color: '#5c6577' } }, { v: '=E3+15', s: { nf: 'dd-mmm-yyyy', h: 'right', b: true } }]);
    rows.push([]);
    rows.push([{ v: 'Bill to', s: { b: true, color: ACC } }]);
    rows.push([{ v: 'Client name' }]); rows.push([{ v: 'Client address, City', s: { color: '#5c6577' } }]);
    rows.push([]);
    rows.push(['Description', 'Qty', 'Rate', 'Amount'].map((h, i) => ({ v: h, s: { ...H, h: i ? 'right' : 'left' } })).concat([]));
    const r0 = rows.length + 1;
    items.forEach((it, i) => { const r = r0 + i; rows.push([{ v: it[0], s: { bB: thin } }, { v: it[1], s: { h: 'right', bB: thin } }, { v: it[2], s: { nf: INR, bB: thin } }, { v: `=B${r}*C${r}`, s: { nf: INR, bB: thin } }]); });
    for (let i = 0; i < 3; i++) { const r = r0 + items.length + i; rows.push([{ s: { bB: thin } }, { s: { h: 'right', bB: thin } }, { s: { nf: INR, bB: thin } }, { v: `=IF(B${r}="","",B${r}*C${r})`, s: { nf: INR, bB: thin } }]); }
    const rN = r0 + items.length + 2;
    rows.push([]);
    rows.push([null, null, { v: 'Subtotal', s: { h: 'right' } }, { v: `=SUM(D${r0}:D${rN})`, s: { nf: INR } }]);
    const sub = rows.length;
    rows.push([null, { v: 'GST rate', s: { h: 'right', color: '#5c6577' } }, { v: 0.18, s: { nf: '0%', h: 'right' } }, { v: `=D${sub}*C${sub + 1}`, s: { nf: INR } }]);
    rows.push([null, null, { v: 'Discount', s: { h: 'right' } }, { v: 0, s: { nf: INR } }]);
    rows.push([null, null, { v: 'Total due', s: { h: 'right', b: true, sz: 13, color: '#FFFFFF', fill: ACC } }, { v: `=D${sub}+D${sub + 1}-D${sub + 2}`, s: { nf: INR, b: true, sz: 13, color: '#FFFFFF', fill: ACC } }]);
    rows.push([]);
    rows.push([{ v: 'Payment: UPI yourname@bank  ·  A/C 1234567890, IFSC ABCD0001234', s: { color: '#5c6577', i: true } }]);
    rows.push([{ v: 'Thank you for your business!', s: { b: true, color: ACC } }]);
    return { sheets: [sheet({ name: 'Invoice', rows, colW: { A: 300, B: 70, C: 110, D: 130, E: 110 }, rowH: { 1: 44 }, showGrid: false, page: { orientation: 'portrait' } })] };
  }

  /* ================= 1.2 templates (themes, tables, PivotTables, sparklines) ================= */
  const rng = (seed) => () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
  const TH = (id) => STheme.get(id);
  const hdr = (th, extra = {}) => ({ b: true, color: '#FFFFFF', fill: th.accents[0], h: 'center', v: 'middle', ...extra });
  const title = (th, t, sz = 20) => ({ v: t, s: { font: th.head, sz, b: true, color: th.dk2 } });
  const sub = (t) => ({ v: t, s: { color: '#5c6577', i: true } });
  const kpi = (th, label, f, nf, i = 0) => [{ v: label, s: { color: '#5c6577', sz: 10, fill: STheme.tint(th.accents[i], 0.9), bT: { style: 'thick', color: th.accents[i] } } }, { v: f, s: { b: true, sz: 16, nf, fill: STheme.tint(th.accents[i], 0.9), font: th.head, bT: { style: 'thick', color: th.accents[i] } } }];

  /* ---- Annual budget with PivotTable summary ---- */
  function annualBudget() {
    const th = TH('corporate'); const R = rng(7); const Y = new Date().getFullYear();
    const cats = [['Salary', 'Income', 52000, 0], ['Freelance', 'Income', 9000, 0.6], ['Rent', 'Expense', 15000, 0], ['Groceries', 'Expense', 7500, 0.15], ['Transport', 'Expense', 2600, 0.25], ['Utilities', 'Expense', 2400, 0.2], ['Education', 'Expense', 4000, 0.5], ['Health', 'Expense', 1500, 0.6], ['Entertainment', 'Expense', 2500, 0.4], ['Savings & SIP', 'Expense', 8000, 0]];
    const rows = [['Date', 'Month', 'Quarter', 'Category', 'Type', 'Amount', 'Notes'].map(h => ({ v: h, s: {} }))];
    for (let m = 1; m <= 12; m++) for (const [cat, type, base, vary] of cats) {
      if (cat === 'Freelance' && R() < 0.35) continue;
      const amt = Math.round(base * (1 + (R() - 0.5) * 2 * vary) / 10) * 10; const r = rows.length + 1;
      rows.push([{ v: SU.dateToSerial(Y, m, Math.min(28, 1 + Math.floor(R() * 27))), s: { nf: 'dd-mmm-yy' } }, { v: `=CHOOSE(MONTH(A${r}),"01 Jan","02 Feb","03 Mar","04 Apr","05 May","06 Jun","07 Jul","08 Aug","09 Sep","10 Oct","11 Nov","12 Dec")` }, { v: `="Q"&ROUNDUP(MONTH(A${r})/3,0)`, s: { h: 'center' } }, cat, type, { v: amt, s: { nf: SU.currencyFormat('₹', 0) } }, cat === 'Education' ? 'Course / books' : '']);
    }
    const n = rows.length; const tr = n + 1;
    rows.push([{ v: 'Total' }, null, null, null, null, { v: `=SUBTOTAL(109,F2:F${n})`, s: { nf: SU.currencyFormat('₹', 0) } }, null]);
    const tx = sheet({ name: 'Transactions', rows, colW: { A: 92, B: 64, C: 70, D: 130, E: 84, F: 110, G: 150 }, freeze: { r: 1, c: 0 }, tables: [{ name: 'Budget', range: `A1:G${tr}`, styleId: 'Medium 2', total: true, totals: { 5: 'sum' } }], tabColor: th.accents[1] });
    const sum = sheet({ name: 'Summary', rows: [
      [title(th, `Annual Budget ${Y}`)], [sub('Totals update automatically from the Transactions table. Refresh the PivotTable after adding rows (Data ▸ Refresh All).')], [],
      [...kpi(th, 'Total income', '=SUMIFS(Budget!F:F,Budget!E:E,"Income")', SU.currencyFormat('₹', 0), 0), null, ...kpi(th, 'Total expenses', '=SUMIFS(Budget!F:F,Budget!E:E,"Expense")', SU.currencyFormat('₹', 0), 4), null, ...kpi(th, 'Net savings', '=B4-E4', SU.currencyFormat('₹', 0), 1), null, ...kpi(th, 'Savings rate', '=IFERROR(H4/B4,0)', '0.0%', 2)],
    ].map(r => r), colW: { A: 130, B: 120, C: 18, D: 130, E: 120, F: 18, G: 130, H: 120, I: 18, J: 130, K: 120 }, rowH: { 1: 36, 4: 34 }, showGrid: false, tabColor: th.accents[0] });
    // fix references: KPIs reference the Transactions sheet by name
    sum.data[3][1] = "=SUMIFS(Transactions!F2:F" + n + ",Transactions!E2:E" + n + ",\"Income\")"; sum.data[3][4] = "=SUMIFS(Transactions!F2:F" + n + ",Transactions!E2:E" + n + ",\"Expense\")"; sum.data[3][7] = '=B4-E4'; sum.data[3][10] = '=IFERROR(H4/B4,0)';
    return { theme: 'corporate', sheets: [sum, tx], after: (wb) => {
      const s = wb.sheetByName('Summary');
      const def = Pivot.create({ table: 'Budget' }, { sheet: s, r: 6, c: 0 }, { name: 'BudgetPivot', rows: ['Type', 'Category'], cols: ['Quarter'], values: [{ field: 'Amount', agg: 'sum', show: 'normal', nf: SU.currencyFormat('₹', 0) }], styleId: 'Light|1', grandRows: false, silent: true }, 'Template');
      const f = Pivot.find(def.id); const o = f.def.out;
      s.charts.push({ id: 'tplc1', type: 'stackedBar', range: { ...o }, sheet: null, pivot: def.id, x: s.colLeft(o.c1 + 1) + 24, y: s.rowTop(6), w: 540, h: 400, z: 1, title: 'Income and spending by category and quarter', style: 1, legend: true });
      wb.activeIndex = 0; s.sel = { r0: 0, c0: 0, r1: 0, c1: 0 }; s.active = { r: 0, c: 0 };
    } };
  }

  /* ---- Gradebook with sparklines ---- */
  function gradebook() {
    const th = TH('academic'); const R = rng(11);
    const names = ['Aarav Sharma', 'Diya Patel', 'Ishaan Kumar', 'Ananya Iyer', 'Rohan Deshmukh', 'Sneha Reddy', 'Kabir Singh', 'Meera Nair', 'Arjun Mehta', 'Priya Joshi', 'Vihaan Gupta', 'Saanvi Rao', 'Aditya Kulkarni', 'Kavya Menon'];
    const tests = ['Quiz 1', 'Quiz 2', 'Assign. 1', 'Mid-term', 'Quiz 3', 'Assign. 2', 'Quiz 4', 'Final'];
    const weights = [0.05, 0.05, 0.1, 0.2, 0.05, 0.1, 0.05, 0.4];
    const rows = [];
    rows.push([title(th, 'Gradebook — B.Sc. Computer Science, Semester 3', 18)]);
    rows.push([sub('Enter marks out of 100. Weighted score, grade and trend update automatically.')]);
    rows.push([{ v: 'Weight', s: { b: true, h: 'right', color: th.dk2 } }, null, ...weights.map(w => ({ v: w, s: { nf: '0%', h: 'center', color: th.dk2, fill: STheme.tint(th.accents[0], 0.9) } })), { v: '=SUM(C3:J3)', s: { nf: '0%', h: 'center', b: true } }]);
    const H2 = hdr(th, { wrap: true });
    rows.push([{ v: 'Roll', s: H2 }, { v: 'Student', s: H2 }, ...tests.map(t => ({ v: t, s: H2 })), { v: 'Weighted %', s: H2 }, { v: 'Grade', s: H2 }, { v: 'Trend', s: H2 }, { v: 'Status', s: H2 }]);
    const r0 = rows.length + 1;
    names.forEach((nm, i) => {
      const r = r0 + i; const base = 52 + R() * 38; const drift = (R() - 0.45) * 4;
      rows.push([{ v: 2301 + i, s: { h: 'center' } }, nm, ...tests.map((_, j) => ({ v: Math.max(18, Math.min(100, Math.round(base + drift * j + (R() - 0.5) * 18))), s: { h: 'center' } })),
        { v: `=SUMPRODUCT(C${r}:J${r},$C$3:$J$3)/100`, s: { nf: '0.0%', h: 'center', b: true } },
        { v: `=VLOOKUP(K${r},'Grade scale'!$A$4:$B$10,2,TRUE)`, s: { h: 'center', b: true } }, { v: null },
        { v: `=IF(K${r}<0.4,"At risk",IF(J${r}<D${r}-10,"Slipping","On track"))`, s: { h: 'center' } }]);
    });
    const rN = r0 + names.length - 1;
    rows.push([null, { v: 'Class average', s: { b: true, h: 'right' } }, ...tests.map((_, j) => { const L = SU.colName(2 + j); return { v: `=AVERAGE(${L}${r0}:${L}${rN})`, s: { nf: '0.0', h: 'center', b: true, fill: STheme.tint(th.accents[0], 0.88) } }; }), { v: `=AVERAGE(K${r0}:K${rN})`, s: { nf: '0.0%', h: 'center', b: true, fill: STheme.tint(th.accents[0], 0.88) } }, { v: `=COUNTIF(L${r0}:L${rN},"A*")&" A grades"`, s: { h: 'center', color: '#5c6577' } }]);
    const avgRow = rN + 1;
    const book = sheet({ name: 'Gradebook', rows, colW: { A: 50, B: 150, C: 60, D: 60, E: 66, F: 66, G: 60, H: 66, I: 60, J: 60, K: 84, L: 56, M: 110, N: 84 }, rowH: { 1: 34, 4: 34 }, freeze: { r: 4, c: 2 },
      cf: [{ type: 'scale', range: `C${r0}:J${rN}`, colors: ['#F8696B', '#FFEB84', '#63BE7B'] }, { type: 'bar', range: `K${r0}:K${rN}`, color: STheme.tint(th.accents[1], 0.45) }, { type: 'eq', v1: 'At risk', range: `N${r0}:N${rN}`, fmt: { fill: '#FFC7CE', color: '#9C0006', b: true } }, { type: 'eq', v1: 'Slipping', range: `N${r0}:N${rN}`, fmt: { fill: '#FFEB9C', color: '#9C5700' } }, { type: 'eq', v1: 'On track', range: `N${r0}:N${rN}`, fmt: { color: '#006100' } }],
      validations: [{ range: `C${r0}:J${rN}`, type: 'whole', op: 'between', v1: 0, v2: 100, errorTitle: 'Invalid marks', errorMsg: 'Marks must be a whole number from 0 to 100.', prompt: 'Enter marks out of 100', promptTitle: 'Marks', ignoreBlank: true }], tabColor: th.accents[0] });
    const scale = sheet({ name: 'Grade scale', rows: [[title(th, 'Grade scale', 16)], [sub('Lowest weighted score for each grade (used by VLOOKUP).')], [], ...[[0, 'F'], [0.4, 'D'], [0.5, 'C'], [0.6, 'B'], [0.7, 'B+'], [0.8, 'A'], [0.9, 'A+']].map(([v, g]) => [{ v, s: { nf: '0%', h: 'center' } }, { v: g, s: { b: true, h: 'center' } }])], colW: { A: 90, B: 70 } });
    scale.data[2] = ['From', 'Grade']; scale.meta.styles.set(SU.key(2, 0), hdr(th)); scale.meta.styles.set(SU.key(2, 1), hdr(th));
    return { theme: 'academic', sheets: [book, scale], after: (wb) => {
      const s = wb.sheetByName('Gradebook'); const th2 = STheme.get('academic');
      const items = names.map((_, i) => ({ r: r0 - 1 + i, c: 12, sheet: null, src: { r0: r0 - 1 + i, c0: 2, r1: r0 - 1 + i, c1: 9 } }));
      s.sparklines = [{ id: 'tsp1', type: 'line', color: th2.accents[1], negColor: '#D64545', markerColor: th2.accents[0], highColor: '#178a3e', lowColor: '#C00000', show: { high: true, low: true, last: true }, weight: 1.6, items },
        { id: 'tsp2', type: 'column', color: th2.accents[0], negColor: '#D64545', markerColor: th2.accents[2], highColor: th2.accents[2], lowColor: '#C00000', show: { high: true, low: true }, items: [{ r: avgRow - 1, c: 12, sheet: null, src: { r0: avgRow - 1, c0: 2, r1: avgRow - 1, c1: 9 } }] }];
      const m = new Map(s.rowH); for (let r = r0 - 1; r <= avgRow - 1; r++) m.set(r, 22); s.rowH = m; s.invalidateGeom();
    } };
  }

  /* ---- Project tracker (Gantt) ---- */
  function projectTracker() {
    const th = TH('modern'); const d0 = (() => { const t = new Date(); const day = (t.getDay() + 6) % 7; const m = new Date(t.getFullYear(), t.getMonth(), t.getDate() - day - 7); return SU.dateToSerial(m.getFullYear(), m.getMonth() + 1, m.getDate()); })();
    const tasks = [['Plan', 'Define scope & goals', 'Ananya', 0, 3, 1], ['Plan', 'Literature review', 'Rohan', 2, 6, 1], ['Plan', 'Project proposal', 'Ananya', 7, 3, 1], ['Build', 'Data collection', 'Kabir', 9, 8, 0.8], ['Build', 'Prototype app', 'Meera', 12, 10, 0.5], ['Build', 'Database design', 'Kabir', 14, 5, 0.6], ['Build', 'Testing round 1', 'Rohan', 22, 5, 0.1], ['Deliver', 'Report draft', 'Ananya', 24, 8, 0], ['Deliver', 'Presentation slides', 'Meera', 30, 4, 0], ['Deliver', 'Final submission', 'All', 35, 2, 0]];
    const days = 42; const C0 = 8; // timeline starts at column I
    const rows = [];
    rows.push([title(th, 'Project Tracker — Final Year Project', 18)]);
    rows.push([{ v: 'Start date', s: { color: '#5c6577' } }, { v: d0, s: { nf: 'dd-mmm-yyyy', b: true, h: 'left' } }, { v: 'Today', s: { color: '#5c6577', h: 'right' } }, { v: '=TODAY()', s: { nf: 'dd-mmm', b: true, h: 'center' } }, null, { v: 'Progress', s: { color: '#5c6577', h: 'right' } }, { v: `=SUMPRODUCT(E5:E${4 + tasks.length},G5:G${4 + tasks.length})/SUM(E5:E${4 + tasks.length})`, s: { nf: '0%', b: true, color: th.accents[1], h: 'center' } }]);
    const H = hdr(th, { sz: 10 });
    const top = [null, null, null, null, null, null, null, null]; for (let k = 0; k < days; k++) top.push(k % 7 === 0 ? { v: `=$B$2+${k}`, s: { nf: '"Week of "d mmm', sz: 9, b: true, color: th.dk2, h: 'center', fill: STheme.tint(th.accents[2], 0.85) } } : { s: { fill: STheme.tint(th.accents[2], 0.85) } }); rows.push(top);
    const hr = [{ v: 'Phase', s: H }, { v: 'Task', s: H }, { v: 'Owner', s: H }, { v: 'Start', s: H }, { v: 'Days', s: H }, { v: 'End', s: H }, { v: '% Done', s: H }, { v: 'Status', s: H }];
    for (let k = 0; k < days; k++) hr.push({ v: `=$B$2+${k}`, s: { ...H, nf: 'd', sz: 8 } });
    rows.push(hr);
    const r0 = rows.length + 1;
    tasks.forEach((t, i) => { const r = r0 + i; rows.push([{ v: t[0], s: { b: true, color: th.accents[1] } }, t[1], t[2], { v: `=$B$2+${t[3]}`, s: { nf: 'dd-mmm', h: 'center' } }, { v: t[4], s: { h: 'center' } }, { v: `=D${r}+E${r}-1`, s: { nf: 'dd-mmm', h: 'center' } }, { v: t[5], s: { nf: '0%', h: 'center' } }, { v: `=IF(G${r}>=1,"Done",IF(F${r}<TODAY(),"Late",IF(D${r}<=TODAY(),"In progress","Not started")))`, s: { h: 'center' } }]); });
    const rN = r0 + tasks.length - 1; const L0 = SU.colName(C0), L1 = SU.colName(C0 + days - 1);
    const colW = { A: 70, B: 170, C: 70, D: 62, E: 44, F: 62, G: 58, H: 86 }; for (let k = 0; k < days; k++) colW[SU.colName(C0 + k)] = 18;
    const acc = th.accents;
    const weekMerges = []; for (let k = 0; k < days; k += 7) weekMerges.push(`${SU.colName(C0 + k)}3:${SU.colName(C0 + Math.min(days - 1, k + 6))}3`);
    const sh = sheet({ name: 'Tracker', rows, colW, rowH: { 1: 34, 4: 26 }, freeze: { r: 4, c: 2 }, showGrid: false, merges: weekMerges,
      cf: [
        { type: 'formula', formula: `=${L0}$4=TODAY()`, range: `${L0}${r0}:${L1}${rN}`, fmt: { fill: '#FDE2E4' } },
        { type: 'formula', formula: `=AND(${L0}$4>=$D${r0},${L0}$4<$D${r0}+ROUND($E${r0}*$G${r0},0))`, range: `${L0}${r0}:${L1}${rN}`, fmt: { fill: STheme.tint(acc[2], -0.15) }, stop: true },
        { type: 'formula', formula: `=AND(${L0}$4>=$D${r0},${L0}$4<=$F${r0})`, range: `${L0}${r0}:${L1}${rN}`, fmt: { fill: STheme.tint(acc[2], 0.55) } },
        { type: 'eq', v1: 'Done', range: `H${r0}:H${rN}`, fmt: { fill: '#C6EFCE', color: '#006100' } }, { type: 'eq', v1: 'Late', range: `H${r0}:H${rN}`, fmt: { fill: '#FFC7CE', color: '#9C0006', b: true } }, { type: 'eq', v1: 'In progress', range: `H${r0}:H${rN}`, fmt: { fill: '#FFEB9C', color: '#9C5700' } },
        { type: 'bar', range: `G${r0}:G${rN}`, color: acc[1] }],
      validations: [{ range: `G${r0}:G${rN}`, type: 'decimal', op: 'between', v1: 0, v2: 1, errorMsg: 'Enter a percentage from 0% to 100%.', prompt: 'Type 50% for half done', promptTitle: '% Done', ignoreBlank: true }, { range: `C${r0}:C${rN}`, type: 'list', source: 'Ananya,Rohan,Kabir,Meera,All', ignoreBlank: true, errorStyle: 'warning' }], tabColor: acc[2] });
    for (let r = r0 - 1; r <= rN - 1; r++) for (let c = 0; c < C0 + days; c++) { const k = SU.key(r, c); const st = sh.meta.styles.get(k) || {}; sh.meta.styles.set(k, { ...st, bB: { style: 'thin', color: '#E5E7EB' } }); }
    return { theme: 'modern', sheets: [sh] };
  }

  /* ---- Sales dashboard ---- */
  function salesDashboard() {
    const th = TH('ocean'); const R = rng(23); const Y = new Date().getFullYear();
    const regions = ['North', 'South', 'East', 'West'], reps = { North: ['Aditi', 'Rahul'], South: ['Lakshmi', 'Karthik'], East: ['Sourav', 'Ritika'], West: ['Neha', 'Vikram'] };
    const products = [['Laptop', 52000], ['Tablet', 21000], ['Phone', 18000], ['Accessories', 1800]];
    const rows = [['Date', 'Month', 'Region', 'Rep', 'Product', 'Units', 'Price', 'Revenue'].map(v => ({ v }))];
    for (let i = 0; i < 96; i++) {
      const m = 1 + (i % 6); const reg = regions[Math.floor(R() * 4)]; const rep = reps[reg][Math.floor(R() * 2)]; const [p, price] = products[Math.floor(R() * 4)];
      const units = p === 'Accessories' ? 5 + Math.floor(R() * 40) : 1 + Math.floor(R() * 12); const r = rows.length + 1;
      rows.push([{ v: SU.dateToSerial(Y, m, 1 + Math.floor(R() * 27)), s: { nf: 'dd-mmm' } }, { v: `=CHOOSE(MONTH(A${r}),"01 Jan","02 Feb","03 Mar","04 Apr","05 May","06 Jun","07 Jul","08 Aug","09 Sep","10 Oct","11 Nov","12 Dec")` }, reg, rep, p, units, { v: price, s: { nf: SU.currencyFormat('₹', 0) } }, { v: `=F${r}*G${r}`, s: { nf: SU.currencyFormat('₹', 0) } }]);
    }
    const n = rows.length;
    const data = sheet({ name: 'Sales data', rows, colW: { A: 70, B: 74, C: 70, D: 80, E: 100, F: 56, G: 86, H: 104 }, freeze: { r: 1, c: 0 }, tables: [{ name: 'Sales', range: `A1:H${n}`, styleId: 'Light 9' }], tabColor: th.accents[1] });
    const money = SU.currencyFormat('₹', 0);
    const dash = sheet({ name: 'Dashboard', rows: [
      [title(th, `Sales Dashboard — H1 ${Y}`)], [sub('Built on the Sales table with PivotTables, PivotCharts and sparklines. Add rows to the table, then Data ▸ Refresh All.')], [],
      [...kpi(th, 'Revenue', `=SUM('Sales data'!H2:H${n})`, money, 0), null, ...kpi(th, 'Units sold', `=SUM('Sales data'!F2:F${n})`, '#,##0', 1), null, ...kpi(th, 'Orders', `=COUNTA('Sales data'!A2:A${n})`, '#,##0', 2), null, ...kpi(th, 'Avg order', '=IFERROR(B4/H4,0)', money, 4)],
    ], colW: { A: 110, B: 112, C: 18, D: 110, E: 112, F: 18, G: 110, H: 112, I: 18, J: 110, K: 112 }, rowH: { 1: 36, 4: 34 }, showGrid: false, tabColor: th.accents[0] });
    return { theme: 'ocean', sheets: [dash, data], after: (wb) => {
      const s = wb.sheetByName('Dashboard');
      const p1 = Pivot.create({ table: 'Sales' }, { sheet: s, r: 6, c: 0 }, { name: 'RegionProduct', rows: ['Region'], cols: ['Product'], values: [{ field: 'Revenue', agg: 'sum', show: 'normal', nf: money }], styleId: 'Medium|1', silent: true }, 'Template');
      const o1 = Pivot.find(p1.id).def.out;
      const p2 = Pivot.create({ table: 'Sales' }, { sheet: s, r: o1.r1 + 3, c: 0 }, { name: 'Monthly', rows: ['Month'], values: [{ field: 'Revenue', agg: 'sum', show: 'normal', nf: money }, { field: 'Units', agg: 'sum', show: 'normal', nf: '#,##0' }], styleId: 'Light|2', silent: true }, 'Template');
      const o2 = Pivot.find(p2.id).def.out;
      const p3 = Pivot.create({ table: 'Sales' }, { sheet: s, r: o2.r1 + 3, c: 0 }, { name: 'RegionTrend', rows: ['Region'], cols: ['Month'], values: [{ field: 'Revenue', agg: 'sum', show: 'normal', nf: money }], grandCols: false, subtotals: false, styleId: 'Light|3', silent: true }, 'Template');
      const o3 = Pivot.find(p3.id).def.out;
      const x = Math.max(s.colLeft(o1.c1 + 1), s.colLeft(o3.c1 + 2)) + 24;
      s.charts.push({ id: 'tpld1', type: 'stackedColumn', range: { ...o1 }, sheet: null, pivot: p1.id, x, y: s.rowTop(6), w: 500, h: 290, z: 1, title: 'Revenue by region and product' });
      s.charts.push({ id: 'tpld2', type: 'combo', range: { ...o2 }, sheet: null, pivot: p2.id, x, y: s.rowTop(6) + 306, w: 500, h: 290, z: 1, title: 'Monthly revenue (columns) and units (line)', legend: true });
      // trend sparklines next to the region × month pivot (body rows: header rows = 2)
      const items = []; for (let r = o3.r0 + 2; r <= o3.r1 - 1; r++) items.push({ r, c: o3.c1 + 1, sheet: null, src: { r0: r, c0: o3.c0 + 1, r1: r, c1: o3.c1 } });
      const th2 = STheme.get('ocean');
      s.sparklines = [{ id: 'tspd', type: 'line', color: th2.accents[1], negColor: '#D64545', markerColor: th2.accents[4], highColor: th2.accents[2], lowColor: th2.accents[4], show: { high: true, low: true, markers: true }, weight: 1.8, items }];
      const st = new Map(s.styles); st.set(SU.key(o3.r0 + 1, o3.c1 + 1), { b: true, fill: STheme.tint(th2.accents[0], 0.8), color: '#000000', h: 'center' }); s.styles = st;
      const cells = [{ r: o3.r0 + 1, c: o3.c1 + 1, v: 'Trend' }]; wb.hf.setCellContents(s.a(cells[0].r, cells[0].c), [['Trend']]);
      const cw = new Map(s.colW); cw.set(o3.c1 + 1, 120); s.colW = cw; s.invalidateGeom();
      wb.activeIndex = 0; s.sel = { r0: 0, c0: 0, r1: 0, c1: 0 }; s.active = { r: 0, c: 0 };
    } };
  }

  /* ================= 1.3 student templates ================= */
  const L = SU.colName;
  const todaySerial = () => { const d = new Date(); return SU.dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); };
  const INPUT = { fill: '#FFF7D6' }; // cells the student types into
  const note = (t) => ({ v: t, s: { color: '#5c6577', i: true, sz: 10 } });

  /* ---- Project plan (Gantt) ---- */
  function projectPlan() {
    const th = TH('fresh'); const acc = th.accents;
    const d0 = (() => { const t = new Date(); const day = (t.getDay() + 6) % 7; const m = new Date(t.getFullYear(), t.getMonth(), t.getDate() - day - 7); return SU.dateToSerial(m.getFullYear(), m.getMonth() + 1, m.getDate()); })();
    // [task, owner, start (offset or 'after'), days, % done]
    const tasks = [['Choose topic & get guide’s approval', 'Asha', 0, 2, 1], ['Research & literature survey', 'Rahul', 'after', 4, 1], ['Buy parts (sensor, pump, board)', 'Imran', 4, 3, 1], ['Circuit design & wiring', 'Imran', 'after', 5, 0.8], ['Write the code', 'Neha', 9, 6, 0.5], ['Assemble the model', 'Imran', 'after', 4, 0], ['Testing & fixing bugs', 'Neha', 'after', 4, 0], ['Project report', 'Rahul', 15, 10, 0.2], ['Poster & presentation slides', 'Asha', 'after', 3, 0], ['Practice the demo', 'All', 'after', 2, 0], ['Final submission & viva', 'All', 'after', 1, 0]];
    const days = 42, C0 = 9; const R0 = 8; // first task row (1-based); timeline starts at column J
    const NT = 14; const RN = R0 + NT - 1; // room for 14 tasks
    const H = hdr(th, { sz: 10, wrap: true });
    const rows = [];
    rows.push([title(th, 'Project Plan — Smart Plant Watering System', 18)]);
    rows.push([null, { v: 'Project start', s: { color: '#5c6577', h: 'right' } }, { v: d0, s: { ...INPUT, nf: 'ddd, d mmm yyyy', b: true, h: 'left', bB: { style: 'thin', color: '#E3C55A' } } }, null, null, { v: 'Today', s: { color: '#5c6577', h: 'right' } }, { v: '=TODAY()', s: { nf: 'd mmm', b: true, color: '#C2410C', h: 'center' } }, { v: `=COUNTIF(H${R0}:H${RN},"Done")&" of "&COUNTA(B${R0}:B${RN})&" tasks done"`, s: { color: '#5c6577' } }]);
    rows.push([null, { v: 'Overall progress', s: { color: '#5c6577', h: 'right' } }, { v: `=IFERROR(SUMPRODUCT(E${R0}:E${RN},G${R0}:G${RN})/SUM(E${R0}:E${RN}),0)`, s: { nf: '0%', b: true, sz: 13, color: th.dk2, h: 'left' } }, null, null, { v: 'Finish', s: { color: '#5c6577', h: 'right' } }, { v: `=MAX(F${R0}:F${RN})`, s: { nf: 'd mmm', b: true, h: 'center' } }, { v: `=IF(G3="","",IF(G3<TODAY(),"Past the finish date",(G3-TODAY())&" days to go"))`, s: { color: '#5c6577' } }]);
    rows.push([note('Type your tasks, owners, start dates and days. A start can also be a formula, like =F8+1 (the day after the task above ends). The bars and ▼ today marker move by themselves.')]);
    // row 5: week labels, row 6: day numbers, row 7: weekday letters
    const wk = Array(C0).fill(null); for (let k = 0; k < days; k++) wk.push(k % 7 === 0 ? { v: `=$C$2+${k}`, s: { nf: '"Week of "d mmm', sz: 9, b: true, color: th.dk2, h: 'center', fill: STheme.tint(acc[0], 0.82) } } : { s: { fill: STheme.tint(acc[0], 0.82) } }); rows.push(wk);
    const r6 = ['#', 'Task', 'Owner', 'Start', 'Days', 'End', '% done', 'Status', null].map(v => v === null ? null : { v, s: H });
    for (let k = 0; k < days; k++) r6.push({ v: `=$C$2+${k}`, s: { nf: 'd', sz: 8, b: true, h: 'center', color: th.dk2, fill: STheme.tint(acc[0], 0.92), ...(k % 7 === 0 ? { bL: { style: 'thin', color: '#B8C4BC' } } : {}) } });
    rows.push(r6);
    const r7 = [{ s: H }, { s: H }, { s: H }, { s: H }, { s: H }, { s: H }, { s: H }, { s: H }, null];
    for (let k = 0; k < days; k++) r7.push({ v: `=CHOOSE(WEEKDAY(${L(C0 + k)}6),"S","M","T","W","T","F","S")`, s: { sz: 8, h: 'center', color: '#7a8496', fill: STheme.tint(acc[0], 0.92), bB: { style: 'thin', color: '#B8C4BC' }, ...(k % 7 === 0 ? { bL: { style: 'thin', color: '#B8C4BC' } } : {}) } });
    rows.push(r7);
    const line = { bB: { style: 'thin', color: '#E5E7EB' } };
    for (let i = 0; i < NT; i++) {
      const r = R0 + i; const t = tasks[i];
      const start = !t ? { s: { nf: 'dd-mmm', h: 'center', ...line } } : t[2] === 'after' ? { v: `=F${r - 1}+1`, s: { nf: 'dd-mmm', h: 'center', ...line } } : { v: d0 + t[2], s: { nf: 'dd-mmm', h: 'center', ...line } };
      const row = [{ v: t ? i + 1 : null, s: { h: 'center', color: '#7a8496', ...line } }, { v: t ? t[0] : null, s: line }, { v: t ? t[1] : null, s: { color: '#47526a', ...line } }, start,
        { v: t ? t[3] : null, s: { h: 'center', ...line } }, { v: `=IF(OR(D${r}="",E${r}=""),"",D${r}+E${r}-1)`, s: { nf: 'dd-mmm', h: 'center', ...line } }, { v: t ? t[4] : null, s: { nf: '0%', h: 'center', ...line } },
        { v: `=IF(B${r}="","",IF(G${r}>=1,"Done",IF(F${r}="","",IF(F${r}<TODAY(),"Late",IF(D${r}<=TODAY(),"In progress","Not started")))))`, s: { h: 'center', sz: 10, ...line } }, null];
      for (let k = 0; k < days; k++) row.push({ s: { ...line, ...(k % 7 === 0 ? { bL: { style: 'thin', color: '#D5DBD7' } } : {}) } });
      rows.push(row);
    }
    const TL = `${L(C0)}${R0}:${L(C0 + days - 1)}${RN}`; const J = L(C0);
    const colW = { A: 30, B: 252, C: 64, D: 62, E: 44, F: 62, G: 58, H: 86, I: 10 }; for (let k = 0; k < days; k++) colW[L(C0 + k)] = 21;
    const merges = ['A6:A7', 'B6:B7', 'C6:C7', 'D6:D7', 'E6:E7', 'F6:F7', 'G6:G7', 'H6:H7', 'C2:D2']; for (let k = 0; k < days; k += 7) merges.push(`${L(C0 + k)}5:${L(C0 + Math.min(days - 1, k + 6))}5`);
    const rowH = { 1: 34, 2: 22, 3: 24, 4: 30, 5: 18, 6: 18, 7: 16 }; for (let r = R0; r <= RN; r++) rowH[r] = 24;
    const sh = sheet({ name: 'Plan', rows, colW, rowH, freeze: { r: 7, c: 2 }, showGrid: false, merges, tabColor: acc[0],
      cf: [
        { type: 'formula', formula: `=AND($E${R0}>0,${J}$6>=$D${R0},${J}$6<$D${R0}+ROUND($E${R0}*$G${R0},0))`, range: TL, fmt: { fill: STheme.tint(acc[0], -0.1) }, stop: true },
        { type: 'formula', formula: `=AND($E${R0}>0,$H${R0}="Late",${J}$6>=$D${R0},${J}$6<=$F${R0})`, range: TL, fmt: { fill: '#F5A3A3' }, stop: true },
        { type: 'formula', formula: `=AND($E${R0}>0,${J}$6>=$D${R0},${J}$6<=$F${R0})`, range: TL, fmt: { fill: STheme.tint(acc[0], 0.6) }, stop: true },
        { type: 'formula', formula: `=${J}$6=TODAY()`, range: TL, fmt: { fill: '#FDE2E4' }, stop: true },
        { type: 'formula', formula: `=WEEKDAY(${J}$6,2)>5`, range: TL, fmt: { fill: '#F4F5F7' } },
        { type: 'formula', formula: `=${J}6=TODAY()`, range: `${J}6:${L(C0 + days - 1)}7`, fmt: { fill: '#E11D48', color: '#FFFFFF', b: true } },
        { type: 'eq', v1: 'Done', range: `H${R0}:H${RN}`, fmt: { fill: '#C6EFCE', color: '#006100' } }, { type: 'eq', v1: 'Late', range: `H${R0}:H${RN}`, fmt: { fill: '#FFC7CE', color: '#9C0006', b: true } }, { type: 'eq', v1: 'In progress', range: `H${R0}:H${RN}`, fmt: { fill: '#FFEB9C', color: '#9C5700' } },
        { type: 'bar', range: `G${R0}:G${RN}`, color: STheme.tint(acc[0], 0.35) }],
      validations: [{ range: `G${R0}:G${RN}`, type: 'decimal', op: 'between', v1: 0, v2: 1, errorTitle: '% done', errorMsg: 'Type a percentage from 0% to 100%, for example 50%.', prompt: 'How much is finished? Type 50% for half.', promptTitle: '% done', ignoreBlank: true }, { range: `E${R0}:E${RN}`, type: 'whole', op: 'between', v1: 1, v2: 365, errorTitle: 'Days', errorMsg: 'Type the number of days, from 1 to 365.', ignoreBlank: true }] });
    return { theme: 'fresh', sheets: [sh] };
  }

  /* ---- Timeline (events by date) ---- */
  function timeline() {
    const th = TH('academic'); const acc = th.accents;
    const cats = [['Uprising', '#B42318'], ['Political', '#1E3A8A'], ['Movement', '#B45309'], ['Turning point', '#6B21A8'], ['Milestone', '#166534']];
    const ev = [[1857, 5, 10, 'First War of Independence begins at Meerut', 'Uprising', 'Sepoys rise against East India Company rule; it spreads to Delhi, Kanpur, Lucknow and Jhansi.'],
      [1885, 12, 28, 'Indian National Congress founded', 'Political', 'First session held in Bombay, presided over by W. C. Bonnerjee.'],
      [1905, 10, 16, 'Partition of Bengal', 'Political', 'Lord Curzon’s partition sparks the Swadeshi movement; it is reversed in 1911.'],
      [1919, 4, 13, 'Jallianwala Bagh massacre', 'Turning point', 'Troops fire on a peaceful crowd in Amritsar; public anger grows across India.'],
      [1920, 9, null, 'Non-Cooperation Movement adopted', 'Movement', 'Approved at the Calcutta session of Congress; called off in 1922 after Chauri Chaura.'],
      [1930, 3, 12, 'Dandi (Salt) March begins', 'Movement', 'Gandhi walks from Sabarmati Ashram to Dandi and breaks the salt law on 6 April.'],
      [1942, 8, 8, 'Quit India resolution passed', 'Movement', 'Passed at the Bombay session; the call is “Do or Die”.'],
      [1947, 8, 15, 'India becomes independent', 'Milestone', 'Jawaharlal Nehru becomes the first Prime Minister.'],
      [1950, 1, 26, 'Constitution comes into force', 'Milestone', 'India becomes a republic; 26 January is celebrated as Republic Day.']];
    const N = 15, R0 = 7, RN = R0 + N - 1; const NB = 24, C0 = 7; // strip H..AE
    const CN = C0 + NB - 1; const cNotes = CN + 2, cGap = CN + 3, cPos = CN + 4; const P = L(cPos), G0 = L(C0), GN = L(CN);
    const H = hdr(th, { sz: 10, wrap: true });
    const rows = [];
    const r1 = [title(th, 'Timeline — India’s Freedom Struggle', 18)]; r1[cPos] = { v: '=YEAR(TODAY())+(MONTH(TODAY())-1)/12+(DAY(TODAY())-1)/365' }; rows.push(r1);
    const r2 = [note('Type one event per row, in date order. Month and day are optional. Works for history (any year) and for project milestones (▼ marks today).')]; r2[cPos] = { v: `=IFERROR(INT(MIN(${P}${R0}:${P}${RN})),0)` }; rows.push(r2);
    const r3 = [null, null, null, null, { v: `=IF(COUNT(${P}${R0}:${P}${RN})=0,"",${P}2&" to "&(${P}2+${NB}*${P}3)&" · each column ≈ "&IF(${P}3>=1,ROUND(${P}3,1)&" years",ROUND(${P}3*12,1)&" months"))`, s: { color: th.dk2, b: true, sz: 10 } }];
    cats.forEach(([c, col], i) => { r3[C0 + i * 5] = { v: c, s: { fill: col, color: '#FFFFFF', b: true, sz: 9, h: 'center' } }; });
    r3[cPos] = { v: `=IFERROR((INT(MAX(${P}${R0}:${P}${RN}))+1-${P}2)/${NB},1)` }; rows.push(r3);
    const r4 = []; for (let k = 0; k < NB; k++) r4[C0 + k] = { v: `=$${P}$2+${k}*$${P}$3` }; rows.push(r4); // hidden: start of each column
    const r5 = [null, null, null, null, null, null, null]; for (let k = 0; k < NB; k++) r5.push(k % 4 === 0 ? { v: `=IF($${P}$3*${NB}<=3,CHOOSE(INT(MOD(${L(C0 + k)}4,1)*12+0.0001)+1,"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")&" "&INT(${L(C0 + k)}4),INT(${L(C0 + k)}4))`, s: { sz: 9, b: true, color: th.dk2, h: 'left', bL: { style: 'thin', color: '#B9AE9C' } } } : null); rows.push(r5);
    const r6 = ['#', 'Year', 'Month', 'Day', 'Event', 'Category', null].map(v => v === null ? null : { v, s: H });
    for (let k = 0; k < NB; k++) r6.push({ v: `=IF(AND($${P}$1>=${L(C0 + k)}$4,$${P}$1<${L(C0 + k)}$4+$${P}$3),"▼","")`, s: { color: '#E11D48', sz: 9, h: 'center', v: 'bottom', fill: STheme.tint(acc[0], 0.9), ...(k % 4 === 0 ? { bL: { style: 'thin', color: '#B9AE9C' } } : {}) } });
    r6.push(null, { v: 'What happened', s: H }, { v: 'Gap from previous', s: H }, { v: 'Position', s: { color: '#9aa3b2', sz: 9 } }); rows.push(r6);
    const lane = STheme.tint(acc[0], 0.95); const line = { bB: { style: 'thin', color: '#E7E1D8' }, v: 'middle' };
    for (let i = 0; i < N; i++) {
      const r = R0 + i; const e = ev[i] || [];
      const row = [{ v: e[0] ? i + 1 : null, s: { h: 'center', color: '#7a8496', ...line } }, { v: e[0] ?? null, s: { h: 'center', b: true, ...line } }, { v: e[1] ?? null, s: { h: 'center', ...line } }, { v: e[2] ?? null, s: { h: 'center', ...line } }, { v: e[3] ?? null, s: { b: true, ...line } }, { v: e[4] ?? null, s: { h: 'center', sz: 10, ...line } }, null];
      for (let k = 0; k < NB; k++) row.push({ v: `=IF($${P}${r}="","",IF(AND($${P}${r}>=${L(C0 + k)}$4,$${P}${r}<${L(C0 + k)}$4+$${P}$3),"●",""))`, s: { h: 'center', v: 'middle', color: '#FFFFFF', fill: lane, bB: { style: 'thin', color: '#FFFFFF' }, ...(k % 4 === 0 ? { bL: { style: 'thin', color: '#D9D0C2' } } : {}) } });
      row.push(null, { v: e[5] ?? null, s: { wrap: true, sz: 10, color: '#3f4656', ...line } });
      row.push({ v: i === 0 ? `=IF(${P}${r}="","","Start")` : `=IF(OR(${P}${r}="",${P}${r - 1}=""),"",IF(${P}${r}-${P}${r - 1}<1,"+"&ROUND((${P}${r}-${P}${r - 1})*365,0)&" days","+"&ROUND(${P}${r}-${P}${r - 1},1)&" yrs"))`, s: { h: 'center', color: '#5c6577', sz: 10, ...line } });
      row.push({ v: `=IF(B${r}="","",B${r}+(IF(C${r}="",1,C${r})-1)/12+(IF(D${r}="",1,D${r})-1)/365)`, s: { nf: '0.00', color: '#9aa3b2', sz: 9 } });
      rows.push(row);
    }
    const colW = { A: 30, B: 54, C: 48, D: 40, E: 300, F: 100, G: 10 }; for (let k = 0; k < NB; k++) colW[L(C0 + k)] = 20; colW[L(CN + 1)] = 10; colW[L(cNotes)] = 380; colW[L(cGap)] = 88; colW[P] = 60;
    const merges = []; for (let k = 0; k < NB; k += 4) merges.push(`${L(C0 + k)}5:${L(C0 + k + 3)}5`); cats.forEach((_, i) => merges.push(`${L(C0 + i * 5)}3:${L(C0 + i * 5 + 3)}3`)); merges.push('E3:F3');
    const rowH = { 1: 34, 2: 20, 3: 20, 5: 18, 6: 30 }; for (let r = R0; r <= RN; r++) rowH[r] = 36;
    const cf = cats.map(([c, col]) => ({ type: 'formula', formula: `=AND(${G0}${R0}="●",$F${R0}="${c}")`, range: `${G0}${R0}:${GN}${RN}`, fmt: { fill: col, color: '#FFFFFF', b: true }, stop: true }));
    cf.push({ type: 'formula', formula: `=AND($${P}$1>=${G0}$4,$${P}$1<${G0}$4+$${P}$3)`, range: `${G0}${R0}:${GN}${RN}`, fmt: { fill: '#FCE7EA' } });
    cf.push({ type: 'formula', formula: `=AND(${G0}${R0}="●",$F${R0}="")`, range: `${G0}${R0}:${GN}${RN}`, fmt: { fill: '#475569', color: '#FFFFFF' } });
    cats.forEach(([c, col]) => cf.push({ type: 'eq', v1: c, range: `F${R0}:F${RN}`, fmt: { color: col, b: true } }));
    const sh = sheet({ name: 'Timeline', rows, colW, rowH, freeze: { r: 6, c: 6 }, showGrid: false, merges, cf, tabColor: acc[0],
      validations: [{ range: `F${R0}:F${RN}`, type: 'list', source: cats.map(c => c[0]).join(','), ignoreBlank: true, errorStyle: 'warning', errorMsg: 'Pick a category from the list (or keep your own).' }, { range: `B${R0}:B${RN}`, type: 'whole', op: 'between', v1: -5000, v2: 3000, errorTitle: 'Year', errorMsg: 'Type a year as a number, like 1947.', ignoreBlank: true }, { range: `C${R0}:C${RN}`, type: 'whole', op: 'between', v1: 1, v2: 12, errorTitle: 'Month', errorMsg: 'Type the month as a number from 1 to 12 (or leave it blank).', ignoreBlank: true }, { range: `D${R0}:D${RN}`, type: 'whole', op: 'between', v1: 1, v2: 31, errorTitle: 'Day', errorMsg: 'Type the day from 1 to 31 (or leave it blank).', ignoreBlank: true }] });
    sh.meta.hiddenR = new Set([3]); sh.meta.hiddenC = new Set([cPos]);
    return { theme: 'academic', sheets: [sh] };
  }

  /* ---- Grade calculator ---- */
  function gradeCalc() {
    const th = TH('slate'); const acc = th.accents;
    const subj = [['Engineering Mathematics III', 4, 17, 24, null, 71], ['Data Structures', 4, 18, 26, null, 78], ['Digital Electronics', 3, 15, 21, 38, 64], ['Object Oriented Programming', 3, 19, 27, 44, 82], ['Computer Organisation', 3, 14, 18, null, 55], ['Data Structures Lab', 1, null, null, 42, null], ['Environmental Studies', 2, 16, 22, null, 48], ['Soft Skills', 1, 8, 9, null, 30]];
    const R0 = 10, N = 12, RN = R0 + N - 1;
    const H = hdr(th, { wrap: true, sz: 10 }); const soft = STheme.tint(acc[1], 0.9);
    const GT = "'Grade table'!$A$5:$C$12";
    const rows = [];
    rows.push([title(th, 'Grade Calculator — Semester 3', 18)]);
    rows.push([note('Type your marks in the yellow cells. Leave a part blank if your subject doesn’t have it — the weights adjust. Grades come from the Grade table sheet, which you can edit.')]);
    rows.push([]);
    rows.push([null, { v: 'Marks scheme', s: { b: true, color: th.dk2 } }, null, { v: 'Assignments', s: H }, { v: 'Internal / mid-term', s: H }, { v: 'Practical / lab', s: H }, { v: 'Final exam', s: H }, { v: 'Total', s: H }]);
    rows.push([null, { v: 'Marks out of', s: { h: 'right', color: '#5c6577' } }, null, { v: 20, s: { ...INPUT, h: 'center', b: true } }, { v: 30, s: { ...INPUT, h: 'center', b: true } }, { v: 50, s: { ...INPUT, h: 'center', b: true } }, { v: 100, s: { ...INPUT, h: 'center', b: true } }, null]);
    rows.push([null, { v: 'Weight', s: { h: 'right', color: '#5c6577' } }, null, { v: 0.15, s: { ...INPUT, nf: '0%', h: 'center', b: true } }, { v: 0.2, s: { ...INPUT, nf: '0%', h: 'center', b: true } }, { v: 0.15, s: { ...INPUT, nf: '0%', h: 'center', b: true } }, { v: 0.5, s: { ...INPUT, nf: '0%', h: 'center', b: true } }, { v: '=SUM(D6:G6)', s: { nf: '0%', h: 'center', b: true } }]);
    rows.push([null, null, null, { v: '=IF(ROUND(H6,4)<>1,"Weights should add up to 100%","")', s: { color: '#C00000', sz: 10, b: true } }]);
    rows.push([]);
    rows.push(['#', 'Subject', 'Credits', 'Assign.', 'Internal', 'Practical', 'Final exam', 'Weighted %', 'Grade', 'Grade point', 'Credit points', 'Result'].map(v => ({ v, s: H })));
    const line = { bB: { style: 'thin', color: '#E2E8F0' } };
    for (let i = 0; i < N; i++) {
      const r = R0 + i; const s = subj[i] || [];
      const mk = (v) => ({ v: v ?? null, s: { ...INPUT, h: 'center', ...line } });
      rows.push([{ v: s[0] ? i + 1 : null, s: { h: 'center', color: '#7a8496', ...line } }, { v: s[0] ?? null, s: { ...line } }, { v: s[1] ?? null, s: { h: 'center', ...line } }, mk(s[2]), mk(s[3]), mk(s[4]), mk(s[5]),
        { v: `=IF(COUNT(D${r}:G${r})=0,"",SUMPRODUCT(ISNUMBER(D${r}:G${r})*D${r}:G${r}/$D$5:$G$5*$D$6:$G$6)/SUMPRODUCT(ISNUMBER(D${r}:G${r})*$D$6:$G$6))`, s: { nf: '0.0%', h: 'center', b: true, ...line } },
        { v: `=IF(H${r}="","",VLOOKUP(H${r},${GT},2,TRUE))`, s: { h: 'center', b: true, ...line } },
        { v: `=IF(H${r}="","",VLOOKUP(H${r},${GT},3,TRUE))`, s: { h: 'center', ...line } },
        { v: `=IF(OR(H${r}="",C${r}=""),"",C${r}*J${r})`, s: { h: 'center', ...line } },
        { v: `=IF(H${r}="","",IF(J${r}>0,"Pass","Fail"))`, s: { h: 'center', ...line } }]);
    }
    const sumL = (t) => ({ v: t, s: { b: true, h: 'right', color: th.dk2 } });
    const big = (f, nf, extra = {}) => ({ v: f, s: { b: true, sz: 14, nf, h: 'center', fill: soft, bT: { style: 'thick', color: acc[1] }, ...extra } });
    rows.push([]);
    const S0 = RN + 2;
    rows.push([null, sumL('Credits counted'), { v: `=SUMPRODUCT(ISNUMBER(H${R0}:H${RN})*C${R0}:C${RN})`, s: { b: true, h: 'center' } }, null, null, null, sumL('Overall percentage'), big(`=IFERROR(SUMPRODUCT(C${R0}:C${RN},H${R0}:H${RN})/C${S0},"")`, '0.00%')]);
    rows.push([null, sumL('Credits earned (passed)'), { v: `=SUMPRODUCT(ISNUMBER(H${R0}:H${RN})*(L${R0}:L${RN}="Pass")*C${R0}:C${RN})`, s: { b: true, h: 'center' } }, null, null, null, sumL('SGPA (out of 10)'), big(`=IFERROR(ROUND(SUMPRODUCT(C${R0}:C${RN},J${R0}:J${RN})/C${S0},2),"")`, '0.00', { color: acc[1] })]);
    rows.push([null, sumL('Subjects failed'), { v: `=COUNTIF(L${R0}:L${RN},"Fail")`, s: { b: true, h: 'center' } }, null, null, null, null, note('SGPA = Σ(credits × grade point) ÷ Σ credits')]);
    rows.push([]);
    rows.push([null, note('Need this as a percentage for a form? Use File ▸ New ▸ CGPA ↔ percentage converter — and always confirm the rule with your university.')]);
    const rowH = { 1: 34, 4: 32, 9: 30 }; for (let r = R0; r <= RN; r++) rowH[r] = 22; rowH[S0] = 28; rowH[S0 + 1] = 28;
    const book = sheet({ name: 'Grades', rows, colW: { A: 30, B: 220, C: 58, D: 88, E: 82, F: 82, G: 80, H: 90, I: 56, J: 62, K: 66, L: 64 }, rowH, freeze: { r: 9, c: 2 }, showGrid: false, tabColor: acc[1],
      cf: [{ type: 'eq', v1: 'Fail', range: `L${R0}:L${RN}`, fmt: { fill: '#FFC7CE', color: '#9C0006', b: true } }, { type: 'eq', v1: 'Pass', range: `L${R0}:L${RN}`, fmt: { color: '#006100' } }, { type: 'eq', v1: 'F', range: `I${R0}:I${RN}`, fmt: { color: '#C00000' } }, { type: 'bar', range: `H${R0}:H${RN}`, color: STheme.tint(acc[1], 0.55) }, { type: 'formula', formula: '=ROUND($H$6,4)<>1', range: 'H6', fmt: { fill: '#FFC7CE', color: '#9C0006' } }],
      validations: [{ range: `C${R0}:C${RN}`, type: 'decimal', op: 'between', v1: 0, v2: 40, errorTitle: 'Credits', errorMsg: 'Type the credits for this subject, for example 4.', ignoreBlank: true },
        { range: `D${R0}:G${RN}`, type: 'decimal', op: 'ge', v1: 0, errorTitle: 'Marks', errorMsg: 'Type the marks you got as a number (or leave it blank).', ignoreBlank: true }] });
    book.meta.cf.push({ id: 'cfover', type: 'formula', formula: `=AND(ISNUMBER(D${R0}),D${R0}>D$5)`, range: SU.parseRange(`D${R0}:G${RN}`), fmt: { fill: '#FFC7CE', color: '#9C0006', b: true } });
    const scale = [[0, 'F', 0, 'Fail'], [0.4, 'P', 4, 'Pass'], [0.5, 'C', 5, 'Average'], [0.55, 'B', 6, 'Above average'], [0.6, 'B+', 7, 'Good'], [0.7, 'A', 8, 'Very good'], [0.8, 'A+', 9, 'Excellent'], [0.9, 'O', 10, 'Outstanding']];
    const gt = sheet({ name: 'Grade table', rows: [[title(th, 'Grade table', 16)], [note('10-point letter grades as in the UGC CBCS guidelines. The % bands differ between universities — change the “From” column to match yours.')], [], [{ v: 'From (weighted %)', s: H }, { v: 'Grade', s: H }, { v: 'Grade point', s: H }, { v: 'Meaning', s: H }],
      ...scale.map(([f, g, p, m]) => [{ v: f, s: { ...INPUT, nf: '0%', h: 'center', ...line } }, { v: g, s: { b: true, h: 'center', ...line } }, { v: p, s: { ...INPUT, h: 'center', ...line } }, { v: m, s: line }]), [], [note('Keep the rows in order from lowest to highest — the lookup needs that.')]],
      colW: { A: 120, B: 70, C: 90, D: 150 }, rowH: { 1: 30, 2: 30, 4: 30 }, showGrid: false, merges: ['A2:D2'], styles: { A2: { wrap: true } }, tabColor: acc[2] });
    return { theme: 'slate', sheets: [book, gt] };
  }

  /* ---- CGPA ↔ percentage converter ---- */
  const CGPA_RULES = [
    ['CBSE (CGPA × 9.5)', 'Percentage = CGPA × 9.5', 9.5, 0, 'linear', 'CBSE’s rule for the old Class 10 CGPA system (CCE). Gives an indicative percentage.'],
    ['Delhi University (CGPA × 9.5)', 'Percentage = CGPA × 9.5', 9.5, 0, 'linear', 'Used for DU CBCS courses. Check with your university.'],
    ['AICTE recommended ((CGPA − 0.75) × 10)', 'Percentage = (CGPA − 0.75) × 10', 10, -7.5, 'linear', 'Conversion recommended by AICTE for technical programmes on the 10-point scale. Check with your university.'],
    ['VTU ((CGPA − 0.75) × 10)', 'Percentage = (CGPA − 0.75) × 10', 10, -7.5, 'linear', 'Visvesvaraya Technological University, Karnataka. Check with your university.'],
    ['Anna University (CGPA × 10)', 'Percentage = CGPA × 10', 10, 0, 'linear', 'Anna University, Tamil Nadu. Check with your university.'],
    ['GTU ((CGPA − 0.5) × 10)', 'Percentage = (CGPA − 0.5) × 10', 10, -5, 'linear', 'Gujarat Technological University. Check with your university.'],
    ['Mumbai University (7.1 × CGPI + 11)', 'Percentage = 7.1 × CGPI + 11', 7.1, 11, 'linear', 'Mumbai University formula for the 10-point CGPI scheme (engineering). Meant for passing CGPIs. Older 7-point grade schemes used a different conversion, so this doesn’t apply to them — check with your university.'],
    ['Mumbai University, newer schemes (7.25 × CGPI + 11)', 'Percentage = 7.25 × CGPI + 11', 7.25, 11, 'linear', 'Seen in newer Mumbai University schemes. Not confirmed here — check with your university before using it.'],
    ['SPPU Pune (slab table)', '≥ 9.5: 20×CGPA − 100 · 8.25–9.49: 12×CGPA − 25 · 6.75–8.24: 10×CGPA − 7.5 · 5.75–6.74: 5×CGPA + 26.25 · 4.75–5.74: 10×CGPA − 2.5 · 4.0–4.74: 6.6×CGPA + 13.6', 0, 0, 'sppu', 'Savitribai Phule Pune University engineering conversion table (2015 / 2019 pattern) as commonly published. Check with your university.'],
    ['Simple estimate (CGPA × 10)', 'Percentage = CGPA × 10', 10, 0, 'linear', 'A rough estimate only — not an official rule.'],
    ['Custom (a × CGPA + b)', 'Percentage = a × CGPA + b (type a and b)', 9.5, 0, 'linear', 'Type your university’s own multiplier (a) and number to add (b) in the yellow cells.'],
  ];
  function cgpaConverter() {
    const th = TH('saffron'); const navy = th.accents[2], green = th.accents[1], orange = th.accents[0];
    const NR = CGPA_RULES.length; const RR = `Rules!$A$5:$A$${4 + NR}`; const col = (c) => `Rules!$${c}$5:$${c}$${4 + NR}`;
    const H = { b: true, color: '#FFFFFF', fill: navy, h: 'center', v: 'middle', wrap: true, sz: 10 };
    const sec = (t) => ({ v: t, s: { b: true, sz: 13, color: navy, font: th.head, bB: { style: 'medium', color: orange } } });
    const lab = (t) => ({ v: t, s: { h: 'right', color: '#47526a' } });
    const sppu = (x) => `IF(${x}>=9.5,20*${x}-100,IF(${x}>=8.25,12*${x}-25,IF(${x}>=6.75,10*${x}-7.5,IF(${x}>=5.75,5*${x}+26.25,IF(${x}>=4.75,10*${x}-2.5,IF(${x}>=4,6.6*${x}+13.6,NA()))))))`;
    const sppuInv = (p) => `IF(${p}>=90,(${p}+100)/20,IF(${p}>=74,(${p}+25)/12,IF(${p}>=60,(${p}+7.5)/10,IF(${p}>=55,(${p}-26.25)/5,IF(${p}>=45,(${p}+2.5)/10,IF(${p}>=40,(${p}-13.6)/6.6,NA()))))))`;
    const toPct = (x) => `IF(ISNUMBER(${x}),IFERROR(IF($H$2="sppu",${sppu(x)},$H$3*${x}+$H$4),"—"),"")`;
    const rows = [];
    const r1 = [title(th, 'CGPA ↔ Percentage Converter', 18)]; r1[7] = { v: `=INDEX(${col('E')},MATCH($C$7,${RR},0))` }; rows.push(r1); // H1 unused label slot
    rows.push([{ v: '⚠  Always confirm with your university. Rules differ between universities, courses and schemes, and they change — use this as a guide, not as an official result.', s: { b: true, color: '#8A4B00', fill: '#FFF1DB', wrap: true, v: 'middle', bL: { style: 'thick', color: orange } } }]);
    rows.push([]);
    rows.push([null, sec('CGPA → percentage')]);
    rows.push([]);
    rows.push([null, lab('Your CGPA / CGPI / SGPA'), { v: 8.2, s: { ...INPUT, b: true, sz: 14, h: 'center', nf: '0.00', bB: { style: 'thin', color: '#E3C55A' } } }, note('out of 10')]);
    rows.push([null, lab('Conversion rule'), { v: CGPA_RULES[0][0], s: { ...INPUT, b: true, bB: { style: 'thin', color: '#E3C55A' } } }]);
    rows.push([null, lab('Formula'), { v: `=INDEX(${col('B')},MATCH(C7,${RR},0))`, s: { color: navy, wrap: true } }]);
    rows.push([null, lab('Percentage'), { v: `=${toPct('C6')}`, s: { b: true, sz: 20, nf: '0.00" %"', h: 'center', color: '#FFFFFF', fill: green } }, { v: `=IF(C6="","",IF(OR(C6<0,C6>10),"CGPA should be between 0 and 10.",IF(NOT(ISNUMBER(C9)),"This rule has no value below CGPA 4 — check your university’s rules.",IF(OR(C9<0,C9>100),"That’s outside 0–100% — this rule may not fit your CGPA.",""))))`, s: { color: '#C00000', b: true, sz: 10, wrap: true } }]);
    rows.push([null, lab('About this rule'), { v: `=INDEX(${col('F')},MATCH(C7,${RR},0))`, s: { color: '#5c6577', i: true, wrap: true, sz: 10 } }]);
    rows.push([]);
    rows.push([null, sec('Percentage → CGPA')]);
    rows.push([]);
    rows.push([null, lab('Your percentage'), { v: 75, s: { ...INPUT, b: true, sz: 14, h: 'center', nf: '0.00" %"', bB: { style: 'thin', color: '#E3C55A' } } }, note('uses the rule chosen above')]);
    rows.push([null, lab('CGPA'), { v: `=IF(ISNUMBER(C14),IFERROR(IF($H$2="sppu",${sppuInv('C14')},(C14-$H$4)/$H$3),"—"),"")`, s: { b: true, sz: 20, nf: '0.00', h: 'center', color: '#FFFFFF', fill: navy } }, { v: '=IF(AND(ISNUMBER(C15),OR(C15<0,C15>10)),"That’s outside 0–10 — this rule may not fit.","")', s: { color: '#C00000', b: true, sz: 10 } }]);
    rows.push([]);
    rows.push([null, sec('Semester-wise SGPA → CGPA')]);
    rows.push([]);
    const T0 = rows.length + 1; rows.push([null, { v: 'Semester', s: H }, { v: 'SGPA', s: H }, { v: 'Credits', s: H }, { v: 'Percentage (rule above)', s: H }]);
    const sg = [[7.8, 20], [8.1, 22], [8.4, 24], [7.9, 22], [8.6, 23], [8.8, 21], [null, 24], [null, 20]];
    const line = { bB: { style: 'thin', color: '#EADFCF' } };
    sg.forEach(([g, c], i) => { const r = T0 + 1 + i; rows.push([null, { v: 'Semester ' + (i + 1), s: { h: 'center', ...line } }, { v: g, s: { ...INPUT, h: 'center', nf: '0.00', ...line } }, { v: c, s: { ...INPUT, h: 'center', ...line } }, { v: `=${toPct('C' + r)}`, s: { h: 'center', nf: '0.00" %"', ...line } }]); });
    const TN = T0 + sg.length;
    rows.push([null, { v: 'CGPA (credit-weighted)', s: { b: true, h: 'right', color: navy } }, { v: `=IFERROR(ROUND(SUMPRODUCT(C${T0 + 1}:C${TN},D${T0 + 1}:D${TN})/SUMPRODUCT(ISNUMBER(C${T0 + 1}:C${TN})*D${T0 + 1}:D${TN}),2),"")`, s: { b: true, sz: 14, h: 'center', nf: '0.00', fill: STheme.tint(navy, 0.88) } }, null, { v: `=${toPct('C' + (TN + 1))}`, s: { b: true, sz: 14, h: 'center', nf: '0.00" %"', fill: STheme.tint(green, 0.85) } }]);
    rows.push([null, { v: 'Simple average of SGPAs', s: { h: 'right', color: '#5c6577' } }, { v: `=IFERROR(ROUND(AVERAGE(C${T0 + 1}:C${TN}),2),"")`, s: { h: 'center', nf: '0.00' } }, null, note('Most universities use the credit-weighted CGPA.')]);
    const rowH = { 1: 34, 2: 40, 4: 24, 6: 30, 7: 24, 8: 34, 9: 40, 10: 46, 12: 24, 14: 30, 15: 40, 17: 24 }; rowH[T0] = 30; rowH[TN + 1] = 28;
    const conv = sheet({ name: 'Converter', rows, colW: { A: 24, B: 190, C: 250, D: 90, E: 170, F: 40, G: 20, H: 60 }, rowH, showGrid: false, merges: ['A2:E2', 'D9:E9', 'C10:E10', 'C8:E8'], tabColor: orange,
      validations: [{ range: 'C7', type: 'list', source: '=' + RR, errorTitle: 'Rule', errorMsg: 'Pick a rule from the list. To add your own, use “Custom” and edit it on the Rules sheet.', prompt: 'Pick your university’s rule', promptTitle: 'Rule' }, { range: 'C6', type: 'decimal', op: 'between', v1: 0, v2: 10, errorStyle: 'warning', errorMsg: 'A CGPA is usually between 0 and 10.', ignoreBlank: true }, { range: `C${T0 + 1}:C${TN}`, type: 'decimal', op: 'between', v1: 0, v2: 10, errorStyle: 'warning', errorMsg: 'An SGPA is usually between 0 and 10.', ignoreBlank: true }] });
    // hidden helpers: H2 type, H3 multiplier a, H4 add b
    conv.data[1] = conv.data[1] || []; conv.data[1][7] = `=INDEX(${col('E')},MATCH($C$7,${RR},0))`; (conv.data[2] = conv.data[2] || [])[7] = `=INDEX(${col('C')},MATCH($C$7,${RR},0))`; (conv.data[3] = conv.data[3] || [])[7] = `=INDEX(${col('D')},MATCH($C$7,${RR},0))`;
    delete conv.data[0][7]; conv.meta.hiddenC = new Set([7]);
    for (const [k, st] of conv.meta.styles) if (!st.v) conv.meta.styles.set(k, { ...st, v: 'middle' });
    const rl = [[title(th, 'Conversion rules', 16)], [note('Common published formulas. Rules differ by university, course and year — always confirm with your university. Edit the yellow “Custom” row for your own rule.')], [], ['Rule', 'Formula', 'Multiply by (a)', 'Add (b)', 'Type', 'Notes'].map(v => ({ v, s: H }))];
    CGPA_RULES.forEach((r, i) => { const custom = i === NR - 1; rl.push([{ v: r[0], s: { b: true, wrap: true, v: 'top', ...line } }, { v: r[1], s: { wrap: true, v: 'top', color: navy, ...line } }, { v: r[4] === 'sppu' ? null : r[2], s: { h: 'center', v: 'top', ...(custom ? INPUT : {}), ...line } }, { v: r[4] === 'sppu' ? null : r[3], s: { h: 'center', v: 'top', ...(custom ? INPUT : {}), ...line } }, { v: r[4], s: { h: 'center', v: 'top', color: '#9aa3b2', sz: 9, ...line } }, { v: r[5], s: { wrap: true, v: 'top', sz: 10, color: '#3f4656', ...line } }]); });
    const rrH = { 1: 30, 2: 34, 4: 40 }; CGPA_RULES.forEach((r, i) => { const lines = Math.max(2, Math.ceil(r[0].length / 24), Math.ceil(r[1].length / 34), Math.ceil(r[5].length / 52)); rrH[5 + i] = lines * 18 + 10; });
    const rules = sheet({ name: 'Rules', rows: rl, colW: { A: 220, B: 260, C: 92, D: 70, E: 50, F: 330 }, rowH: rrH, showGrid: false, merges: ['A2:F2'], styles: { A2: { wrap: true } }, freeze: { r: 4, c: 1 }, tabColor: navy });
    return { theme: 'saffron', sheets: [conv, rules] };
  }
  T.CGPA_RULES = CGPA_RULES;

  const mini = (title, lines, accent = ACC) => `<div style="width:100%;height:100%;padding:12px 10px;background:#fff;font:9px Carlito,Calibri,sans-serif;color:#333;text-align:left"><div style="font-weight:700;font-size:11px;color:#0B4F3E;margin-bottom:6px">${title}</div>${lines}</div>`;
  const grid = (rows, cols, hdr = true, color = ACC) => `<div style="display:grid;grid-template-columns:repeat(${cols},1fr);gap:1px;background:#d7dde3;border:1px solid #d7dde3">${Array.from({ length: rows * cols }, (_, i) => `<i style="display:block;height:9px;background:${hdr && i < cols ? color : (Math.floor(i / cols) % 2 ? '#f3f7f6' : '#fff')}"></i>`).join('')}</div>`;
  T.list = [
    { name: 'Monthly budget', build: () => fixBudgetChart(budget()), thumbHTML: mini('Monthly Budget', grid(5, 4) + '<div style="height:6px"></div>' + grid(6, 5) + `<div style="margin:6px auto 0;width:40px;height:40px;border-radius:50%;background:conic-gradient(${ACC} 0 35%,#2F5597 0 55%,#E07A1F 0 70%,#7A4FB5 0 85%,#D64545 0);-webkit-mask:radial-gradient(circle,transparent 11px,#000 12px)"></div>`) },
    { name: 'Student marks sheet', build: () => fixMarksChart(marks()), thumbHTML: mini('Class 10-A Results', grid(9, 8) + '<div style="display:flex;align-items:flex-end;gap:3px;height:36px;margin-top:6px">' + [60, 80, 45, 90, 70, 55, 85].map(h => `<i style="flex:1;height:${h}%;background:${ACC}"></i>`).join('') + '</div>') },
    { name: 'Attendance tracker', build: attendance, thumbHTML: mini('Attendance Register', `<div style="display:grid;grid-template-columns:2.2fr repeat(12,1fr);gap:1px;background:#d7dde3;border:1px solid #d7dde3">${Array.from({ length: 13 * 13 }, (_, i) => { const r = Math.floor(i / 13), c = i % 13; const bg = r === 0 ? ACC : c === 0 ? '#fff' : ((r * 7 + c * 5) % 11 === 0 ? '#FFC7CE' : (r * 3 + c) % 13 === 0 ? '#FFEB9C' : '#fff'); return `<i style="display:block;height:9px;background:${bg}"></i>`; }).join('')}</div>`) },
    { name: 'To-do list', build: todo, thumbHTML: mini('My To-Do List', `<div style="height:5px;background:#e3e6ea;border-radius:3px;margin-bottom:6px"><div style="width:30%;height:100%;background:${ACC};border-radius:3px"></div></div>` + Array.from({ length: 10 }, (_, i) => `<div style="display:flex;gap:4px;align-items:center;margin:3px 0"><i style="width:7px;height:7px;border:1px solid #9aa3b2;border-radius:2px;background:${i % 3 === 0 ? ACC : '#fff'}"></i><i style="flex:1;height:5px;background:#e3e6ea;border-radius:2px"></i><i style="width:22px;height:7px;border-radius:2px;background:${['#C6EFCE', '#FFEB9C', '#eceff3'][i % 3]}"></i></div>`).join('')) },
    { name: 'Annual budget + PivotTable', build: annualBudget, thumbHTML: mini('<span style="font-family:Inter,sans-serif;color:#0F172A">Annual Budget</span>', `<div style="display:flex;gap:3px;margin-bottom:5px">${['#1D4ED8', '#DC2626', '#0F766E', '#F59E0B'].map(c => `<i style="flex:1;height:14px;background:${c}22;border-top:2px solid ${c}"></i>`).join('')}</div><div style="display:flex;gap:6px"><div style="flex:1.2">${grid(9, 4, true, '#C7D2FE')}</div><div style="flex:1;display:flex;align-items:flex-end;gap:3px;height:80px">${[[30, 20], [40, 25], [35, 30], [50, 28]].map(([a, b]) => `<div style="flex:1;display:flex;flex-direction:column;justify-content:flex-end"><i style="height:${b}px;background:#0F766E"></i><i style="height:${a}px;background:#1D4ED8"></i></div>`).join('')}</div></div>`, '#1D4ED8') },
    { name: 'Gradebook with sparklines', build: gradebook, thumbHTML: mini('<span style="font-family:\'Source Serif 4\',serif;color:#3B1F1F">Gradebook</span>', Array.from({ length: 9 }, (_, i) => `<div style="display:flex;gap:2px;align-items:center;margin:2px 0"><i style="width:34px;height:7px;background:#eee"></i>${Array.from({ length: 6 }, (_, j) => `<i style="flex:1;height:7px;background:${['#F8696B', '#FFEB84', '#63BE7B', '#b5d99c', '#FFEB84', '#63BE7B'][(i + j * 2) % 6]}"></i>`).join('')}<svg width="30" height="9" viewBox="0 0 30 9"><polyline points="0,7 6,5 12,6 18,3 24,4 30,1" fill="none" stroke="#1E3A8A" stroke-width="1.3"/></svg></div>`).join('')) },
    { name: 'Project tracker (Gantt)', build: projectTracker, thumbHTML: mini('<span style="font-family:Montserrat,sans-serif;color:#111827">Project Tracker</span>', Array.from({ length: 10 }, (_, i) => `<div style="display:flex;gap:4px;align-items:center;margin:3px 0"><i style="width:38px;height:6px;background:#e3e6ea;border-radius:2px"></i><div style="flex:1;position:relative;height:7px"><i style="position:absolute;left:${[0, 5, 16, 21, 28, 33, 52, 57, 71, 83][i]}%;width:${[7, 14, 7, 19, 24, 12, 12, 19, 10, 6][i]}%;height:7px;background:${i < 3 ? '#0891B2' : '#7dd3e1'};border-radius:2px"></i></div></div>`).join('')) },
    { name: 'Sales dashboard', build: salesDashboard, thumbHTML: mini('<span style="font-family:Poppins,sans-serif;color:#012A4A">Sales Dashboard</span>', `<div style="display:flex;gap:3px;margin-bottom:5px">${['#03396C', '#0077B6', '#2A9D8F', '#E76F51'].map(c => `<i style="flex:1;height:14px;background:${c}22;border-top:2px solid ${c}"></i>`).join('')}</div><div style="display:flex;gap:6px"><div style="flex:1">${grid(6, 4, true, '#03396C')}<div style="height:4px"></div>${grid(5, 3, true, '#BFE0F2')}</div><div style="flex:1;display:flex;flex-direction:column;gap:4px"><div style="display:flex;align-items:flex-end;gap:2px;height:38px">${[60, 80, 45, 90, 70, 55].map(h => `<i style="flex:1;height:${h}%;background:#0077B6"></i>`).join('')}</div><svg viewBox="0 0 60 20" height="24"><polyline points="0,15 10,12 20,14 30,6 40,9 50,4 60,7" fill="none" stroke="#2A9D8F" stroke-width="2"/></svg></div></div>`, '#0077B6') },
    { name: 'Project plan (Gantt)', build: projectPlan, thumbHTML: mini('<span style="font-family:Nunito,sans-serif;color:#14532D">Project Plan</span>', `<div style="display:flex;gap:4px;margin-bottom:4px"><i style="width:44px"></i><div style="flex:1;position:relative;height:8px;background:#dcfce7"><i style="position:absolute;left:37%;top:-2px;width:2px;height:120px;background:#E11D48;opacity:.7"></i></div></div>` + Array.from({ length: 10 }, (_, i) => `<div style="display:flex;gap:4px;align-items:center;margin:3px 0"><i style="width:44px;height:6px;background:#e3e6ea;border-radius:2px"></i><div style="flex:1;position:relative;height:8px"><i style="position:absolute;left:${[0, 5, 10, 17, 22, 30, 40, 36, 60, 68][i]}%;width:${[5, 10, 7, 12, 14, 10, 10, 24, 8, 5][i]}%;height:8px;background:${i < 3 ? '#118A3E' : i < 5 ? 'linear-gradient(90deg,#118A3E 55%,#8bd3a5 55%)' : '#8bd3a5'};border-radius:2px"></i></div></div>`).join(''), '#16A34A') },
    { name: 'Timeline', build: timeline, thumbHTML: mini('<span style="font-family:\'Source Serif 4\',serif;color:#3B1F1F">Timeline</span>', Array.from({ length: 9 }, (_, i) => `<div style="display:flex;gap:4px;align-items:center;margin:3px 0"><i style="width:16px;height:6px;background:#e3e6ea"></i><i style="width:40px;height:6px;background:#efe9df"></i><div style="flex:1;position:relative;height:9px;background:#f6f2ec"><i style="position:absolute;left:${[0, 30, 52, 66, 68, 78, 90, 95, 97][i]}%;width:6px;height:9px;border-radius:2px;background:${['#B42318', '#1E3A8A', '#1E3A8A', '#6B21A8', '#B45309', '#B45309', '#B45309', '#166534', '#166534'][i]}"></i></div></div>`).join('')) },
    { name: 'Grade calculator (SGPA)', build: gradeCalc, thumbHTML: mini('<span style="font-family:\'IBM Plex Sans\',sans-serif;color:#1E293B">Grade Calculator</span>', `<div style="display:flex;gap:2px;margin-bottom:5px">${[1, 1, 1, 1].map(() => '<i style="flex:1;height:8px;background:#FFF7D6;border:1px solid #eadb9a"></i>').join('')}</div>` + grid(8, 6, true, '#475569') + `<div style="display:flex;gap:6px;margin-top:6px"><div style="flex:1;height:18px;background:#dbe5fb;border-top:2px solid #2563EB;font:700 9px sans-serif;color:#2563EB;display:flex;align-items:center;justify-content:center">SGPA 7.9</div><div style="flex:1;height:18px;background:#dbe5fb;border-top:2px solid #2563EB;font:700 9px sans-serif;color:#333;display:flex;align-items:center;justify-content:center">71.6%</div></div>`, '#2563EB') },
    { name: 'CGPA ↔ percentage converter', build: cgpaConverter, thumbHTML: mini('<span style="font-family:Poppins,sans-serif;color:#3A2A12">CGPA <span style="font-family:DejaVu Sans,sans-serif">↔</span> %</span>', `<div style="height:12px;background:#FFF1DB;border-left:3px solid #FF9933;margin-bottom:6px"></div><div style="display:flex;gap:4px;align-items:center;margin:4px 0"><i style="width:40px;height:6px;background:#e3e6ea"></i><i style="flex:1;height:12px;background:#FFF7D6;border:1px solid #eadb9a"></i></div><div style="display:flex;gap:4px;align-items:center;margin:4px 0"><i style="width:40px;height:6px;background:#e3e6ea"></i><i style="flex:1;height:12px;background:#FFF7D6;border:1px solid #eadb9a"></i></div><div style="display:flex;gap:4px;align-items:center;margin:4px 0"><i style="width:40px;height:6px;background:#e3e6ea"></i><div style="flex:1;height:20px;background:#138808;color:#fff;font:700 11px sans-serif;display:flex;align-items:center;justify-content:center">77.90 %</div></div><div style="height:6px"></div>` + grid(5, 3, true, '#000080'), '#FF9933') },
    { name: 'Invoice', build: invoice, thumbHTML: mini(`<span style="color:${ACC};font-size:14px">INVOICE</span>`, `<div style="height:4px;width:60%;background:#e3e6ea;margin:3px 0"></div><div style="height:4px;width:45%;background:#e3e6ea;margin:3px 0 10px"></div>` + grid(7, 4) + `<div style="margin:8px 0 0 auto;width:55%;height:12px;background:${ACC}"></div>`) },
  ];
})();
