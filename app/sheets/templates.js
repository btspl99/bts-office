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
    { name: 'Invoice', build: invoice, thumbHTML: mini(`<span style="color:${ACC};font-size:14px">INVOICE</span>`, `<div style="height:4px;width:60%;background:#e3e6ea;margin:3px 0"></div><div style="height:4px;width:45%;background:#e3e6ea;margin:3px 0 10px"></div>` + grid(7, 4) + `<div style="margin:8px 0 0 auto;width:55%;height:12px;background:${ACC}"></div>`) },
  ];
})();
