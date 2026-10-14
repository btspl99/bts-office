/* BTS Sheets — utilities: addresses, formula reference tokenizer, input parsing, number formatting */
(function () {
  const SU = window.SU = {};
  const MAXC = 16384, MAXR = 1048576;
  SU.MAXC = MAXC; SU.MAXR = MAXR;
  SU.key = (r, c) => r * MAXC + c;
  SU.unkey = (k) => ({ r: Math.floor(k / MAXC), c: k % MAXC });

  /* ---------- addresses ---------- */
  SU.colName = (c) => { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
  SU.colIndex = (s) => { let n = 0; s = s.toUpperCase(); for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n - 1; };
  SU.addr = (r, c, absR, absC) => (absC ? '$' : '') + SU.colName(c) + (absR ? '$' : '') + (r + 1);
  SU.rangeStr = (g) => (g.r0 === g.r1 && g.c0 === g.c1) ? SU.addr(g.r0, g.c0) : SU.addr(g.r0, g.c0) + ':' + SU.addr(g.r1, g.c1);
  SU.absRangeStr = (g) => (g.r0 === g.r1 && g.c0 === g.c1) ? SU.addr(g.r0, g.c0, 1, 1) : SU.addr(g.r0, g.c0, 1, 1) + ':' + SU.addr(g.r1, g.c1, 1, 1);
  SU.norm = (g) => ({ r0: Math.min(g.r0, g.r1), c0: Math.min(g.c0, g.c1), r1: Math.max(g.r0, g.r1), c1: Math.max(g.c0, g.c1) });
  SU.inRange = (g, r, c) => r >= g.r0 && r <= g.r1 && c >= g.c0 && c <= g.c1;
  SU.intersects = (a, b) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
  SU.parseAddr = (s) => { const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(s.trim()); if (!m) return null; const c = SU.colIndex(m[1]), r = +m[2] - 1; if (c >= MAXC || r < 0 || r >= MAXR) return null; return { r, c }; };
  /** "A1", "A1:B3", "A:C", "2:5" (optionally with $), returns {r0,c0,r1,c1} */
  SU.parseRange = (s) => {
    s = String(s || '').trim().replace(/\$/g, ''); if (!s) return null;
    let m;
    if ((m = /^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/.exec(s))) {
      const a = { r: +m[2] - 1, c: SU.colIndex(m[1]) }, b = m[3] ? { r: +m[4] - 1, c: SU.colIndex(m[3]) } : a;
      if (a.r < 0 || b.r < 0 || a.c >= MAXC || b.c >= MAXC) return null;
      return SU.norm({ r0: a.r, c0: a.c, r1: b.r, c1: b.c });
    }
    if ((m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(s))) return SU.norm({ r0: 0, c0: SU.colIndex(m[1]), r1: MAXR - 1, c1: SU.colIndex(m[2]) });
    if ((m = /^(\d+):(\d+)$/.exec(s))) return SU.norm({ r0: +m[1] - 1, c0: 0, r1: +m[2] - 1, c1: MAXC - 1 });
    return null;
  };
  SU.quoteSheet = (n) => /^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'";
  SU.unquoteSheet = (s) => s && s[0] === "'" ? s.slice(1, -1).replace(/''/g, "'") : s;
  /** "Sheet 2!A1:B3" or "'Sheet 2'!A1" → {sheet, range} */
  SU.parseRef = (s) => {
    s = String(s || '').trim(); if (s[0] === '=') s = s.slice(1);
    const i = s.lastIndexOf('!');
    if (i < 0) { const g = SU.parseRange(s); return g ? { sheet: null, range: g } : null; }
    const g = SU.parseRange(s.slice(i + 1)); if (!g) return null;
    return { sheet: SU.unquoteSheet(s.slice(0, i)), range: g };
  };

  /* ---------- formula reference tokenizer ---------- */
  const REF_Y = /((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(?:(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?|(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})|(\$?)(\d+):(\$?)(\d+))(?![A-Za-z0-9_(!])/y;
  /** Calls fn(ref) for every reference outside string literals; fn returns replacement string or undefined (keep). */
  SU.mapRefs = (f, fn) => {
    let out = '', i = 0; const n = f.length;
    while (i < n) {
      const ch = f[i];
      if (ch === '"') { let j = i + 1; while (j < n) { if (f[j] === '"') { if (f[j + 1] === '"') { j += 2; continue; } break; } j++; } out += f.slice(i, j + 1); i = j + 1; continue; }
      const prev = i > 0 ? f[i - 1] : '';
      if (/[A-Za-z0-9_$'.]/.test(ch) && !/[A-Za-z0-9_.]/.test(prev)) {
        REF_Y.lastIndex = i; const m = REF_Y.exec(f);
        if (m) {
          let ref = null;
          if (m[3] !== undefined) {
            const a = { c: SU.colIndex(m[3]), r: +m[5] - 1, ac: !!m[2], ar: !!m[4] };
            const b = m[7] !== undefined ? { c: SU.colIndex(m[7]), r: +m[9] - 1, ac: !!m[6], ar: !!m[8] } : null;
            if (a.c < MAXC && a.r >= 0 && (!b || (b.c < MAXC && b.r >= 0))) ref = { kind: 'cell', a, b };
          } else if (m[11] !== undefined) {
            const a = { c: SU.colIndex(m[11]), ac: !!m[10] }, b = { c: SU.colIndex(m[13]), ac: !!m[12] };
            if (a.c < MAXC && b.c < MAXC) ref = { kind: 'col', a, b };
          } else if (m[15] !== undefined) {
            const a = { r: +m[15] - 1, ar: !!m[14] }, b = { r: +m[17] - 1, ar: !!m[16] };
            if (a.r >= 0 && b.r >= 0) ref = { kind: 'row', a, b };
          }
          if (ref) {
            ref.sheetRaw = m[1] ? m[1].slice(0, -1) : null; ref.sheet = ref.sheetRaw ? SU.unquoteSheet(ref.sheetRaw) : null;
            ref.start = i; ref.end = i + m[0].length; ref.text = m[0];
            const rep = fn(ref);
            out += rep === undefined ? m[0] : rep; i += m[0].length; continue;
          }
        }
        // skip the whole identifier / number so we don't match inside it
        let j = i; while (j < n && /[A-Za-z0-9_.$]/.test(f[j])) j++;
        if (j === i) j = i + 1;
        out += f.slice(i, j); i = j; continue;
      }
      out += ch; i++;
    }
    return out;
  };
  SU.refToString = (ref) => {
    const pre = ref.sheetRaw ? ref.sheetRaw + '!' : '';
    if (ref.kind === 'cell') return pre + SU.addr(ref.a.r, ref.a.c, ref.a.ar, ref.a.ac) + (ref.b ? ':' + SU.addr(ref.b.r, ref.b.c, ref.b.ar, ref.b.ac) : '');
    if (ref.kind === 'col') return pre + (ref.a.ac ? '$' : '') + SU.colName(ref.a.c) + ':' + (ref.b.ac ? '$' : '') + SU.colName(ref.b.c);
    return pre + (ref.a.ar ? '$' : '') + (ref.a.r + 1) + ':' + (ref.b.ar ? '$' : '') + (ref.b.r + 1);
  };
  SU.refRange = (ref) => {
    if (ref.kind === 'cell') { const b = ref.b || ref.a; return SU.norm({ r0: ref.a.r, c0: ref.a.c, r1: b.r, c1: b.c }); }
    if (ref.kind === 'col') return SU.norm({ r0: 0, c0: ref.a.c, r1: MAXR - 1, c1: ref.b.c });
    return SU.norm({ r0: ref.a.r, c0: 0, r1: ref.b.r, c1: MAXC - 1 });
  };
  /** Shift the relative parts of all references (copy/paste, fill). */
  SU.shiftFormula = (f, dr, dc) => {
    if (!dr && !dc) return f;
    return SU.mapRefs(f, (ref) => {
      const sh = (p) => { if ('r' in p && !p.ar) p.r += dr; if ('c' in p && !p.ac) p.c += dc; return ('r' in p ? p.r >= 0 && p.r < MAXR : true) && ('c' in p ? p.c >= 0 && p.c < MAXC : true); };
      if (ref.kind === 'col') { if (!ref.a.ac) ref.a.c += dc; if (!ref.b.ac) ref.b.c += dc; if (ref.a.c < 0 || ref.b.c < 0) return '#REF!'; }
      else if (ref.kind === 'row') { if (!ref.a.ar) ref.a.r += dr; if (!ref.b.ar) ref.b.r += dr; if (ref.a.r < 0 || ref.b.r < 0) return '#REF!'; }
      else { if (!sh(ref.a) || (ref.b && !sh(ref.b))) return '#REF!'; }
      return SU.refToString(ref);
    });
  };
  /** References inside a formula (for highlighting): [{sheet, range, start, end}] */
  SU.formulaRefs = (f) => { const list = []; SU.mapRefs(f, (ref) => { list.push({ sheet: ref.sheet, range: SU.refRange(ref), start: ref.start, end: ref.end, ref }); }); return list; };

  /* ---------- dates ---------- */
  const EPOCH = Date.UTC(1899, 11, 30);
  SU.dateToSerial = (y, m, d) => (Date.UTC(y, m - 1, d) - EPOCH) / 86400000;
  SU.jsDateToSerial = (dt) => (Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate(), dt.getUTCHours(), dt.getUTCMinutes(), dt.getUTCSeconds()) - EPOCH) / 86400000;
  SU.localDateToSerial = (dt) => (Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate(), dt.getHours(), dt.getMinutes(), dt.getSeconds()) - EPOCH) / 86400000;
  SU.serialToDate = (s) => new Date(EPOCH + Math.round(s * 86400000)); // UTC fields
  const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  SU.MONTHS_FULL = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  SU.MONTHS_SHORT = SU.MONTHS_FULL.map(m => m.slice(0, 3));
  SU.DAYS_FULL = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  SU.DAYS_SHORT = SU.DAYS_FULL.map(d => d.slice(0, 3));
  const validDate = (y, m, d) => m >= 1 && m <= 12 && d >= 1 && d <= new Date(Date.UTC(y, m, 0)).getUTCDate();
  const fixYear = (y) => y < 100 ? (y < 50 ? 2000 + y : 1900 + y) : y;

  /* ---------- input parsing ---------- */
  /** Parses typed text into {value} | {formula} plus optional suggested number format nf. */
  SU.parseInput = (s) => {
    if (s == null || s === '') return { value: null };
    s = String(s);
    if (s[0] === '=' && s.length > 1) return { formula: s };
    if (s[0] === '+' && s.length > 1 && /^\+[A-Za-z(]/.test(s)) return { formula: '=' + s.slice(1) };
    if (s[0] === "'") return { value: s, text: true };
    const t = s.trim();
    if (/^(true|false)$/i.test(t)) return { value: t.toUpperCase() === 'TRUE' };
    let m;
    const numRe = /^([-+]?)((?:\d{1,3}(?:,\d{3})+|\d{1,2}(?:,\d{2})+,\d{3}|\d+)?(?:\.\d+)?)(?:[eE]([-+]?\d+))?$/;
    const toNum = (str) => parseFloat(str.replace(/,/g, ''));
    const decs = (str) => { const i = str.indexOf('.'); return i < 0 ? 0 : str.length - i - 1; };
    if ((m = numRe.exec(t)) && /\d/.test(m[2])) {
      let v = toNum(m[2]); if (m[3]) v *= Math.pow(10, +m[3]); if (m[1] === '-') v = -v;
      const r = { value: v };
      if (m[3]) r.nf = '0.00E+00';
      else if (m[2].includes(',')) r.nf = decs(m[2]) ? '#,##0.00' : '#,##0';
      return r;
    }
    if ((m = /^([-+]?)(\d[\d,]*\.?\d*|\.\d+)\s?%$/.exec(t))) { const v = toNum(m[2]) / 100 * (m[1] === '-' ? -1 : 1); return { value: +v.toPrecision(15), nf: decs(m[2]) ? '0.' + '0'.repeat(Math.min(decs(m[2]), 6)) + '%' : '0%' }; }
    if ((m = /^(-?)\s?(₹|Rs\.?|INR|\$|€|£)\s?(\d[\d,]*\.?\d*|\.\d+)$/i.exec(t)) || (m = /^(-?)(₹|\$|€|£)-?(\d[\d,]*\.?\d*)$/.exec(t))) {
      const sym = /^(rs|inr)/i.test(m[2]) ? '₹' : m[2];
      return { value: toNum(m[3]) * (m[1] === '-' ? -1 : 1), nf: SU.currencyFormat(sym, decs(m[3]) ? 2 : 0) };
    }
    // dates: dd/mm/yyyy dd-mm-yyyy dd.mm.yyyy
    if ((m = /^(\d{1,2})([\/\-.])(\d{1,2})\2(\d{2}|\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([ap]m)?)?$/i.exec(t))) {
      const d = +m[1], mo = +m[3], y = fixYear(+m[4]);
      if (validDate(y, mo, d)) {
        let v = SU.dateToSerial(y, mo, d);
        if (m[5] !== undefined) { let h = +m[5]; if (m[8]) h = (h % 12) + (/p/i.test(m[8]) ? 12 : 0); v += (h * 3600 + +m[6] * 60 + +(m[7] || 0)) / 86400; return { value: v, nf: 'dd/mm/yyyy h:mm' }; }
        return { value: v, nf: m[2] === '-' ? 'dd-mm-yyyy' : 'dd/mm/yyyy' };
      }
    }
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t))) {
      const y = +m[1], mo = +m[2], d = +m[3];
      if (validDate(y, mo, d)) { let v = SU.dateToSerial(y, mo, d); if (m[4] !== undefined) { v += (+m[4] * 3600 + +m[5] * 60 + +(m[6] || 0)) / 86400; return { value: v, nf: 'yyyy-mm-dd h:mm' }; } return { value: v, nf: 'yyyy-mm-dd' }; }
    }
    // 15-Jan-2024, 15 Jan 2024, 15 January 2024, Jan 15, 2024, 15-Jan
    if ((m = /^(\d{1,2})[\s\-]([A-Za-z]{3,9})\.?(?:[\s\-,]+(\d{2}|\d{4}))?$/.exec(t))) {
      const mo = MONTHS.indexOf(m[2].slice(0, 3).toLowerCase()) + 1; const y = m[3] ? fixYear(+m[3]) : new Date().getFullYear();
      if (mo && validDate(y, mo, +m[1]) && (m[2].length === 3 || SU.MONTHS_FULL[mo - 1].toLowerCase().startsWith(m[2].toLowerCase()))) return { value: SU.dateToSerial(y, mo, +m[1]), nf: m[3] ? 'dd-mmm-yyyy' : 'dd-mmm' };
    }
    if ((m = /^([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})$/.exec(t))) {
      const mo = MONTHS.indexOf(m[1].slice(0, 3).toLowerCase()) + 1;
      if (mo && validDate(+m[3], mo, +m[2])) return { value: SU.dateToSerial(+m[3], mo, +m[2]), nf: 'mmm d, yyyy' };
    }
    if ((m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([ap]\.?m\.?)?$/i.exec(t))) {
      let h = +m[1]; const mi = +m[2], se = +(m[3] || 0);
      if (mi < 60 && se < 60 && (m[4] ? h >= 1 && h <= 12 : h < 24)) {
        if (m[4]) h = (h % 12) + (/p/i.test(m[4]) ? 12 : 0);
        return { value: (h * 3600 + mi * 60 + se) / 86400, nf: m[4] ? (m[3] ? 'h:mm:ss AM/PM' : 'h:mm AM/PM') : (m[3] ? 'h:mm:ss' : 'h:mm') };
      }
    }
    return { value: s, text: true };
  };

  /* ---------- number formats ---------- */
  SU.currencyFormat = (sym, d = 2) => `"${sym}"#,##0${d ? '.' + '0'.repeat(d) : ''}`;
  SU.accountingFormat = (sym, d = 2) => { const z = d ? '.' + '0'.repeat(d) : ''; const s = sym ? `"${sym}"` : ''; return `_(${s}* #,##0${z}_);_(${s}* (#,##0${z});_(${s}* "-"${d ? '??' : ''}_);_(@_)`; };
  SU.FORMATS = [
    { id: 'general', label: 'General', nf: 'General', sample: 'abc 123' },
    { id: 'number', label: 'Number', nf: '0.00', sample: '1234.50' },
    { id: 'currency', label: 'Currency', nf: SU.currencyFormat('₹'), sample: '₹1,234.50' },
    { id: 'accounting', label: 'Accounting', nf: SU.accountingFormat('₹'), sample: '₹ 1,234.50' },
    { id: 'shortdate', label: 'Short Date', nf: 'dd/mm/yyyy', sample: '15/01/2024' },
    { id: 'longdate', label: 'Long Date', nf: 'dddd, d mmmm yyyy', sample: 'Monday, 15 January 2024' },
    { id: 'time', label: 'Time', nf: 'h:mm:ss AM/PM', sample: '6:30:00 PM' },
    { id: 'percent', label: 'Percentage', nf: '0.00%', sample: '12.50%' },
    { id: 'fraction', label: 'Fraction', nf: '# ?/?', sample: '1 1/2' },
    { id: 'scientific', label: 'Scientific', nf: '0.00E+00', sample: '1.23E+03' },
    { id: 'text', label: 'Text', nf: '@', sample: 'abc' },
  ];
  SU.formatId = (nf) => {
    if (!nf || nf === 'General') return 'general';
    const f = SU.FORMATS.find(x => x.nf === nf); if (f) return f.id;
    if (nf === '@') return 'text';
    if (/^_\(/.test(nf)) return 'accounting';
    if (/%/.test(nf)) return 'percent';
    if (/E\+/i.test(nf)) return 'scientific';
    if (/\?\/\?/.test(nf)) return 'fraction';
    if (/[₹$€£]/.test(nf)) return 'currency';
    if (SU.isDateFormat(nf)) return /[hs]/i.test(nf.replace(/"[^"]*"/g, '')) && !/[dy]/i.test(nf.replace(/"[^"]*"/g, '')) ? 'time' : (/dddd|mmmm/.test(nf) ? 'longdate' : 'shortdate');
    if (/[0#]/.test(nf)) return 'number';
    return 'custom';
  };
  SU.isDateFormat = (nf) => { if (!nf || nf === 'General') return false; const s = nf.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, ''); return /[dmyhs]/i.test(s) && !/[0#?]/.test(s.replace(/s+\.0+/g, '')); };

  function splitSections(nf) {
    const out = []; let cur = '', q = false;
    for (let i = 0; i < nf.length; i++) { const ch = nf[i]; if (ch === '"') q = !q; if (ch === '\\' && !q) { cur += ch + (nf[i + 1] || ''); i++; continue; } if (ch === ';' && !q) { out.push(cur); cur = ''; } else cur += ch; }
    out.push(cur); return out;
  }
  const COLORS = { red: '#d9262f', blue: '#1f5fd1', green: '#178a3e', black: '#000', white: '#fff', magenta: '#c026d3', cyan: '#0891b2', yellow: '#ca8a04' };
  const fmtCache = new Map();
  function sectionFor(nf, v) {
    let c = fmtCache.get(nf); if (!c) { c = splitSections(nf); fmtCache.set(nf, c); if (fmtCache.size > 2000) fmtCache.clear(); }
    let sec = c[0], idx = 0;
    if (typeof v === 'number') { if (v < 0 && c.length > 1) { sec = c[1]; idx = 1; } else if (v === 0 && c.length > 2) { sec = c[2]; idx = 2; } }
    else if (typeof v === 'string' && c.length > 3) { sec = c[3]; idx = 3; }
    return { sec, idx, sections: c };
  }
  function ssf(nf, v) { try { return XLSX.SSF.format(nf, v); } catch { try { return XLSX.SSF.format('General', v); } catch { return String(v); } } }

  /** Formats a value. Returns {text, color?, acct?:{sym, num}} */
  SU.format = (v, nf) => {
    if (v === null || v === undefined || v === '') return { text: '' };
    if (typeof v === 'object' && v.value !== undefined) return { text: String(v.value), err: true }; // DetailedCellError
    if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE' };
    if (typeof v === 'string') {
      if (nf && nf.includes(';')) { const { sections } = sectionFor(nf, v); if (sections.length > 3 && sections[3].includes('@')) return { text: sections[3].replace(/_\(|_\)/g, '').replace(/@/, v).replace(/"/g, '') }; }
      return { text: v };
    }
    if (typeof v !== 'number' || !isFinite(v)) return { text: String(v) };
    if (!nf || nf === 'General') return { text: generalNum(v) };
    if (nf === '@') return { text: generalNum(v) };
    const { sec, sections } = sectionFor(nf, v);
    let color; const cm = /\[(red|blue|green|black|white|magenta|cyan|yellow)\]/i.exec(sec); if (cm) color = COLORS[cm[1].toLowerCase()];
    if (/\*/.test(nf) && /_\(/.test(nf)) { // accounting: symbol left, number right
      const sym = (/"([^"]*)"\*/.exec(sections[0]) || [, ''])[1];
      const d = (/0\.(0+)/.exec(sections[0]) || [, ''])[1].length;
      let num;
      if (Math.abs(v) < Math.pow(10, -d) / 2) num = '-' + (d ? '  ' : '') + ' ';
      else { num = ssf('#,##0' + (d ? '.' + '0'.repeat(d) : ''), Math.abs(v)); num = v < 0 ? '(' + num + ')' : num + ' '; }
      return { text: (sym ? sym + ' ' : '') + num.trim(), acct: { sym, num }, color };
    }
    if (SU.isDateFormat(nf) && (v < 0 || v > 2958465)) return { text: '#'.repeat(8) };
    return { text: ssf(nf, v), color };
  };
  function generalNum(v) {
    if (Number.isInteger(v) && Math.abs(v) < 1e11) return String(v);
    const a = Math.abs(v);
    if (a !== 0 && (a >= 1e11 || a < 1e-9)) { let s = v.toExponential(5).replace(/\.?0+e/, 'e'); return s.replace('e', 'E').replace(/E([+-])(\d)$/, 'E$10$2'); }
    let s = String(+v.toPrecision(11));
    if (s.includes('e')) s = v.toFixed(10).replace(/\.?0+$/, '');
    return s;
  }
  SU.generalNum = generalNum;
  /** change number of decimals in a format code */
  SU.changeDecimals = (nf, delta, sample) => {
    if (!nf || nf === 'General') {
      let d = 0; if (typeof sample === 'number') { const s = generalNum(sample); const i = s.indexOf('.'); d = i < 0 ? 0 : s.length - i - 1; }
      d = Math.max(0, d + delta); return d ? '0.' + '0'.repeat(d) : '0';
    }
    const secs = splitSections(nf);
    return secs.map(sec => {
      if (!/[0#?]/.test(sec)) return sec;
      if (delta > 0) {
        if (/0\.0*/.test(sec) || /[0#]\.[0#?]*/.test(sec)) return sec.replace(/([0#?])\.([0#?]*)/, (m, a, b) => a + '.' + b + '0');
        return sec.replace(/([0#?])(?![0#?,.])/, '$1.0');
      } else {
        return sec.replace(/([0#?])\.([0#?]+)/, (m, a, b) => b.length > 1 ? a + '.' + b.slice(0, -1) : a);
      }
    }).join(';');
  };

  /* ---------- misc ---------- */
  SU.clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  SU.esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  SU.debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };
  SU.isErr = (v) => v && typeof v === 'object' && 'value' in v && 'type' in v;
  SU.valueText = (v) => SU.isErr(v) ? v.value : v === null || v === undefined ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
  /** comparable sort key: numbers < text < booleans < errors < blanks (Excel order) */
  SU.cmp = (a, b) => {
    const rank = (v) => v === null || v === undefined || v === '' ? 5 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3;
    const ra = rank(a), rb = rank(b); if (ra !== rb) return ra - rb;
    if (ra === 0) return a - b;
    if (ra === 1) return a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true });
    if (ra === 2) return (a ? 1 : 0) - (b ? 1 : 0);
    return 0;
  };
  SU.toTSV = (rows) => rows.map(r => r.map(v => { v = v == null ? '' : String(v); return /[\t\n"]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }).join('\t')).join('\r\n');
  SU.parseDelimited = (text, delim) => {
    const rows = []; let row = [], cur = '', q = false, i = 0;
    if (!delim) { const first = text.split(/\r?\n/)[0] || ''; delim = (first.match(/\t/g) || []).length ? '\t' : ((first.match(/;/g) || []).length > (first.match(/,/g) || []).length ? ';' : ','); }
    while (i < text.length) {
      const ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i += 2; continue; } q = false; i++; continue; } cur += ch; i++; continue; }
      if (ch === '"' && cur === '') { q = true; i++; continue; }
      if (ch === delim) { row.push(cur); cur = ''; i++; continue; }
      if (ch === '\r' && text[i + 1] === '\n') i++;
      if (ch === '\n' || ch === '\r') { row.push(cur); rows.push(row); row = []; cur = ''; i++; continue; }
      cur += ch; i++;
    }
    if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
    return rows;
  };
  SU.FONTS = ['Calibri', 'Arial', 'Cambria', 'Candara', 'Consolas', 'Courier New', 'Georgia', 'Inter', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'Figtree'];
  const FALLBACK = { Calibri: 'Calibri, Carlito', Cambria: 'Cambria, Caladea, serif', Arial: 'Arial, "Liberation Sans", Arimo, sans-serif', 'Times New Roman': '"Times New Roman", "Liberation Serif", Tinos, serif', 'Courier New': '"Courier New", "Liberation Mono", Cousine, monospace', Consolas: 'Consolas, "DejaVu Sans Mono", monospace', Georgia: 'Georgia, serif', Verdana: 'Verdana, "DejaVu Sans", sans-serif' };
  SU.cssFont = (name) => (FALLBACK[name] || ('"' + name + '"')) + ', Carlito, "Segoe UI", "DejaVu Sans", sans-serif';
  SU.SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
})();
/* formula dialect bridge: Excel <-> HyperFormula */
(function () {
  const SU = window.SU;
  // transform outside string literals
  function outsideStrings(f, fn) { return f.split(/("(?:[^"]|"")*")/).map((p, i) => i % 2 ? p : fn(p)).join(''); }
  SU.outsideStrings = outsideStrings;
  SU.toHF = (f) => {
    if (typeof f !== 'string' || f[0] !== '=') return f;
    return outsideStrings(f, p => p.replace(/_xlfn\.|_xlws\.|_xludf\./gi, '').replace(/(^|[^A-Za-z0-9_.!$'])(TRUE|FALSE)(?![A-Za-z0-9_.(!])/gi, (m, a, b) => a + b.toUpperCase() + '()').replace(/(^|[(,=+\-*\/&^<> ])@(?=[A-Za-z$(])/g, '$1'));
  };
  SU.fromHF = (f) => typeof f === 'string' ? outsideStrings(f, p => p.replace(/\b(TRUE|FALSE)\(\)/g, '$1')) : f;
})();
