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

  const mini = (title, lines, accent = ACC) => `<div style="width:100%;height:100%;padding:12px 10px;background:#fff;font:9px Carlito,Calibri,sans-serif;color:#333;text-align:left"><div style="font-weight:700;font-size:11px;color:#0B4F3E;margin-bottom:6px">${title}</div>${lines}</div>`;
  const grid = (rows, cols, hdr = true) => `<div style="display:grid;grid-template-columns:repeat(${cols},1fr);gap:1px;background:#d7dde3;border:1px solid #d7dde3">${Array.from({ length: rows * cols }, (_, i) => `<i style="display:block;height:9px;background:${hdr && i < cols ? ACC : (Math.floor(i / cols) % 2 ? '#f3f7f6' : '#fff')}"></i>`).join('')}</div>`;
  T.list = [
    { name: 'Monthly budget', build: () => fixBudgetChart(budget()), thumbHTML: mini('Monthly Budget', grid(5, 4) + '<div style="height:6px"></div>' + grid(6, 5) + `<div style="margin:6px auto 0;width:40px;height:40px;border-radius:50%;background:conic-gradient(${ACC} 0 35%,#2F5597 0 55%,#E07A1F 0 70%,#7A4FB5 0 85%,#D64545 0);-webkit-mask:radial-gradient(circle,transparent 11px,#000 12px)"></div>`) },
    { name: 'Student marks sheet', build: () => fixMarksChart(marks()), thumbHTML: mini('Class 10-A Results', grid(9, 8) + '<div style="display:flex;align-items:flex-end;gap:3px;height:36px;margin-top:6px">' + [60, 80, 45, 90, 70, 55, 85].map(h => `<i style="flex:1;height:${h}%;background:${ACC}"></i>`).join('') + '</div>') },
    { name: 'Attendance tracker', build: attendance, thumbHTML: mini('Attendance Register', `<div style="display:grid;grid-template-columns:2.2fr repeat(12,1fr);gap:1px;background:#d7dde3;border:1px solid #d7dde3">${Array.from({ length: 13 * 13 }, (_, i) => { const r = Math.floor(i / 13), c = i % 13; const bg = r === 0 ? ACC : c === 0 ? '#fff' : ((r * 7 + c * 5) % 11 === 0 ? '#FFC7CE' : (r * 3 + c) % 13 === 0 ? '#FFEB9C' : '#fff'); return `<i style="display:block;height:9px;background:${bg}"></i>`; }).join('')}</div>`) },
    { name: 'To-do list', build: todo, thumbHTML: mini('My To-Do List', `<div style="height:5px;background:#e3e6ea;border-radius:3px;margin-bottom:6px"><div style="width:30%;height:100%;background:${ACC};border-radius:3px"></div></div>` + Array.from({ length: 10 }, (_, i) => `<div style="display:flex;gap:4px;align-items:center;margin:3px 0"><i style="width:7px;height:7px;border:1px solid #9aa3b2;border-radius:2px;background:${i % 3 === 0 ? ACC : '#fff'}"></i><i style="flex:1;height:5px;background:#e3e6ea;border-radius:2px"></i><i style="width:22px;height:7px;border-radius:2px;background:${['#C6EFCE', '#FFEB9C', '#eceff3'][i % 3]}"></i></div>`).join('')) },
    { name: 'Invoice', build: invoice, thumbHTML: mini(`<span style="color:${ACC};font-size:14px">INVOICE</span>`, `<div style="height:4px;width:60%;background:#e3e6ea;margin:3px 0"></div><div style="height:4px;width:45%;background:#e3e6ea;margin:3px 0 10px"></div>` + grid(7, 4) + `<div style="margin:8px 0 0 auto;width:55%;height:12px;background:${ACC}"></div>`) },
  ];
})();
