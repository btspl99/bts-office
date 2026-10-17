/* BTS Sheets — OpenDocument Spreadsheet (.ods) writer.
 * Writes values, formulas (OpenFormula), number formats, cell formatting, merges, column widths / row heights,
 * hidden rows/columns, comments, links, data validation, named ranges, freeze panes, tab colours and page orientation.
 * Not written (ODF has no portable equivalent we can produce): charts, pictures, conditional formatting, sparklines, PivotTables. */
(function () {
  const ODS = window.SODS = {};
  const { key } = SU;
  const x = (s) => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const inch = (px) => (px / 96).toFixed(4) + 'in';

  /* ================= formulas: Excel A1 syntax → OpenFormula ================= */
  const MAXROW = 1048576;
  /* Excel 2010+ functions that ODF stores with a COM.MICROSOFT. prefix (LibreOffice convention) */
  const MS_FN = new Set(['IFS', 'SWITCH', 'MAXIFS', 'MINIFS', 'TEXTJOIN', 'CONCAT', 'XLOOKUP', 'XMATCH', 'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'SEQUENCE', 'RANDARRAY', 'LET', 'AGGREGATE', 'CEILING', 'FLOOR']);
  const OF_DOTTED = new Set(['ERROR.TYPE']);
  const fnName = (n) => { const u = n.toUpperCase().replace(/^_XLFN\.|^_XLWS\./, ''); if (MS_FN.has(u) || (u.includes('.') && !OF_DOTTED.has(u) && !u.startsWith('COM.') && !u.startsWith('ORG.'))) return 'COM.MICROSOFT.' + u; return u; };
  const ofSheet = (n) => '$' + (/^[A-Za-z_][A-Za-z0-9_]*$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");
  const REF = /^((?:'(?:[^']|'')+'|[A-Za-z_À-￿][\w.À-￿]*)!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(?![\w(.!À-￿])/;
  /** one Excel reference (optionally sheet-qualified) → ODF address without brackets, e.g. $Sheet1.A1:.B2 */
  function refAddr(sheetPart, area) {
    const sh = sheetPart ? ofSheet(SU.unquoteSheet(sheetPart.slice(0, -1))) : '';
    let [a, b] = area.split(':');
    if (b !== undefined) {
      if (/^\$?[A-Za-z]+$/.test(a)) { a = a + '1'; b = b + MAXROW; } // whole columns
      else if (/^\$?\d+$/.test(a)) { a = '$A' + a; b = '$AMJ' + b; } // whole rows
      return sh + '.' + a.toUpperCase() + ':.' + b.toUpperCase();
    }
    return sh + '.' + a.toUpperCase();
  }
  ODS.formula = (f) => {
    const s = String(f).replace(/^=/, ''); let out = '', i = 0, brace = 0; const n = s.length;
    while (i < n) {
      const ch = s[i];
      if (ch === '"') { let j = i + 1; while (j < n) { if (s[j] === '"') { if (s[j + 1] === '"') { j += 2; continue; } break; } j++; } out += s.slice(i, j + 1); i = j + 1; continue; }
      if (ch === '{') { brace++; out += ch; i++; continue; }
      if (ch === '}') { brace = Math.max(0, brace - 1); out += ch; i++; continue; }
      if (ch === ',') { out += ';'; i++; continue; }
      if (ch === ';') { out += brace ? '|' : ';'; i++; continue; }
      if (ch === '#') { const m = /^#(?:N\/A|DIV\/0!|REF!|NAME\?|VALUE!|NUM!|NULL!|SPILL!|CALC!|CYCLE!|ERROR!)/i.exec(s.slice(i)); if (m) { out += m[0].toUpperCase(); i += m[0].length; continue; } }
      if (/[A-Za-z_$'\dÀ-￿\\]/.test(ch) && !/[\w.]/.test(s[i - 1] || '')) {
        const rest = s.slice(i);
        const m = REF.exec(rest);
        if (m && !(/^\d/.test(rest) && !/^\$?\d+:\$?\d+/.test(rest))) { out += '[' + refAddr(m[1], m[2]) + ']'; i += m[0].length; continue; }
        if (/\d/.test(ch)) { const nm = /^\d*\.?\d+(?:[eE][+-]?\d+)?%?/.exec(rest) || /^\d+\.?/.exec(rest); out += nm[0]; i += nm[0].length; continue; }
        const id = /^[A-Za-z_\\À-￿][\w.À-￿]*/.exec(rest);
        if (id) {
          const w = id[0]; i += w.length;
          if (s[i] === '(') out += fnName(w);
          else if (/^(TRUE|FALSE)$/i.test(w)) out += w.toUpperCase() + '()';
          else out += w; // defined name
          continue;
        }
      }
      out += ch; i++;
    }
    return 'of:=' + out;
  };

  /* ================= number formats: Excel code → ODF data style ================= */
  const CUR = { '₹': ['INR', 'en', 'IN'], '$': ['USD', 'en', 'US'], '€': ['EUR', 'en', 'IE'], '£': ['GBP', 'en', 'GB'], '¥': ['JPY', 'ja', 'JP'], 'Rs': ['INR', 'en', 'IN'], 'Rs.': ['INR', 'en', 'IN'] };
  const NCOL = { red: '#FF0000', blue: '#0000FF', green: '#00FF00', black: '#000000', white: '#FFFFFF', magenta: '#FF00FF', cyan: '#00FFFF', yellow: '#FFFF00' };
  function splitSections(nf) { const out = []; let cur = '', q = false, br = false; for (let i = 0; i < nf.length; i++) { const c = nf[i]; if (c === '\\' && !q) { cur += c + (nf[i + 1] || ''); i++; continue; } if (c === '"') q = !q; else if (!q && c === '[') br = true; else if (!q && c === ']') br = false; if (c === ';' && !q && !br) { out.push(cur); cur = ''; } else cur += c; } out.push(cur); return out; }
  function tokenize(sec) {
    const t = []; let color = null, i = 0;
    while (i < sec.length) {
      const c = sec[i], rest = sec.slice(i);
      if (/^General/i.test(rest)) { t.push({ k: 'gen' }); i += 7; continue; }
      if (c === '"') { const j = sec.indexOf('"', i + 1); t.push({ k: 'lit', v: sec.slice(i + 1, j < 0 ? undefined : j) }); i = j < 0 ? sec.length : j + 1; continue; }
      if (c === '\\') { t.push({ k: 'lit', v: sec[i + 1] || '' }); i += 2; continue; }
      if (c === '_') { t.push({ k: 'lit', v: ' ' }); i += 2; continue; }
      if (c === '*') { i += 2; continue; }
      if (c === '[') {
        const j = sec.indexOf(']', i); const inner = sec.slice(i + 1, j < 0 ? undefined : j); i = j < 0 ? sec.length : j + 1;
        if (inner[0] === '$') { const sym = inner.slice(1).split('-')[0]; if (sym) t.push({ k: 'cur', v: sym }); }
        else if (/^(h+|m+|s+)$/i.test(inner)) t.push({ k: 'dt', v: inner.toLowerCase(), elapsed: true });
        else if (NCOL[inner.toLowerCase()]) color = NCOL[inner.toLowerCase()];
        continue;
      }
      if (/^AM\/PM/i.test(rest)) { t.push({ k: 'ampm' }); i += 5; continue; }
      if (/^A\/P/i.test(rest)) { t.push({ k: 'ampm' }); i += 3; continue; }
      const nm = /^[0#?][0#?,.]*(?:[eE][+-][0#]+)?|^\.[0#?]+/.exec(rest);
      if (nm) {
        // fraction "# ?/?" or "?/8"
        const fr = /^([0#?,]+ )?([0#?]+)\/([0-9#?]+)/.exec(rest);
        if (fr) { t.push({ k: 'frac', whole: fr[1] || '', num: fr[2], den: fr[3] }); i += fr[0].length; continue; }
        t.push({ k: 'num', v: nm[0] }); i += nm[0].length; continue;
      }
      if (/[dmyhse]/i.test(c)) { let j = i; while (j < sec.length && sec[j].toLowerCase() === c.toLowerCase()) j++; t.push({ k: 'dt', v: sec.slice(i, j).toLowerCase() }); i = j; continue; }
      if (c === '%') { t.push({ k: 'pct' }); i++; continue; }
      if (c === '@') { t.push({ k: 'text' }); i++; continue; }
      if (c === '$') { t.push({ k: 'cur', v: '$' }); i++; continue; }
      t.push({ k: 'lit', v: c }); i++;
    }
    // quoted currency symbols ("₹"#,##0) become real currency symbols
    t.forEach(o => { if (o.k === 'lit' && CUR[o.v.trim()] && t.some(z => z.k === 'num')) { o.k = 'cur'; o.v = o.v.trim(); } });
    return { t, color };
  }
  function numberEl(v, opt = {}) {
    let s = v; let factor = 1; while (/,$/.test(s)) { factor *= 1000; s = s.slice(0, -1); }
    const sci = /[eE][+-]([0#]+)$/.exec(s); if (sci) s = s.slice(0, sci.index);
    const [ip, dp = ''] = s.split('.');
    const decs = (dp.match(/[0#?]/g) || []).length, minDecs = (dp.match(/0/g) || []).length;
    const minInt = (ip.match(/0/g) || []).length; const grouping = /[0#?],[0#?]/.test(ip);
    if (sci) return `<number:scientific-number number:decimal-places="${decs}" number:min-decimal-places="${minDecs}" number:min-integer-digits="${Math.max(1, minInt)}" number:min-exponent-digits="${sci[1].length}"/>`;
    return `<number:number number:decimal-places="${decs}" number:min-decimal-places="${minDecs}" number:min-integer-digits="${minInt}"${grouping ? ' number:grouping="true"' : ''}${factor > 1 ? ` number:display-factor="${factor}"` : ''}${opt.pct ? '' : ''}/>`;
  }
  /** returns {kind, body(xml children), color} for one section, or null for General */
  function sectionStyle(sec) {
    const { t, color } = tokenize(sec);
    if (!t.length || (t.length === 1 && t[0].k === 'gen')) return null;
    const hasNum = t.some(o => o.k === 'num' || o.k === 'frac' || o.k === 'gen');
    const dts = t.filter(o => o.k === 'dt' || o.k === 'ampm');
    let kind = 'number';
    if (!hasNum && dts.length) kind = dts.some(o => o.k === 'dt' && /^[dy]/.test(o.v) || (o.k === 'dt' && /^m{3,}$/.test(o.v))) ? 'date' : 'time';
    else if (t.some(o => o.k === 'pct')) kind = 'percentage';
    else if (t.some(o => o.k === 'cur')) kind = 'currency';
    else if (!hasNum && t.some(o => o.k === 'text')) kind = 'text';
    if (kind === 'date' || kind === 'time') {
      // m / mm = minutes when right after hours or right before seconds
      const seq = t.filter(o => o.k === 'dt');
      seq.forEach((o, i) => { if (/^m{1,2}$/.test(o.v)) { const p = seq[i - 1], q = seq[i + 1]; if ((p && /^h/.test(p.v)) || (q && /^s/.test(q.v))) o.min = true; } });
      if (kind === 'date' && seq.every(o => o.min || /^[hs]/.test(o.v) || o.elapsed)) kind = 'time';
    }
    let body = ''; let lit = '';
    const flush = () => { if (lit) { body += `<number:text>${x(lit)}</number:text>`; lit = ''; } };
    let numDone = false;
    for (let i = 0; i < t.length; i++) {
      const o = t[i];
      if (o.k === 'lit') { lit += o.v; continue; }
      if (o.k === 'pct') { lit += '%'; continue; }
      if (o.k === 'gen') { flush(); body += '<number:number number:min-integer-digits="1"/>'; numDone = true; continue; }
      if (o.k === 'num') {
        if ((kind === 'date' || kind === 'time')) continue;
        if (numDone) { lit += o.v.replace(/[#?]/g, '').replace(/0/g, '0'); continue; }
        flush(); body += numberEl(o.v); numDone = true; continue;
      }
      if (o.k === 'frac') { flush(); const den = /^\d+$/.test(o.den) ? ` number:denominator-value="${o.den}"` : ''; body += `<number:fraction number:min-integer-digits="${(o.whole.match(/0/g) || []).length}" number:min-numerator-digits="${o.num.length}" number:min-denominator-digits="${o.den.length}"${den}/>`; numDone = true; continue; }
      if (o.k === 'cur') { flush(); const c = CUR[o.v] || ['', '', '']; body += `<number:currency-symbol${c[1] ? ` number:language="${c[1]}" number:country="${c[2]}"` : ''}>${x(o.v)}</number:currency-symbol>`; continue; }
      if (o.k === 'text') { flush(); body += '<number:text-content/>'; continue; }
      if (o.k === 'ampm') { flush(); body += '<number:am-pm/>'; continue; }
      if (o.k === 'dt') {
        if (kind !== 'date' && kind !== 'time') { lit += o.v; continue; }
        flush(); const v = o.v; const long = ' number:style="long"';
        if (o.elapsed || /^h/.test(v)) body += `<number:hours${v.length > 1 ? long : ''}/>`;
        else if (o.min) body += `<number:minutes${v.length > 1 ? long : ''}/>`;
        else if (/^s/.test(v)) { const nx = t[i + 1]; let dec = 0; if (nx && nx.k === 'num' && nx.v[0] === '.') { dec = nx.v.length - 1; i++; } body += `<number:seconds${v.length > 1 ? long : ''}${dec ? ` number:decimal-places="${dec}"` : ''}/>`; }
        else if (/^d/.test(v)) body += v.length >= 4 ? `<number:day-of-week${long}/>` : v.length === 3 ? '<number:day-of-week/>' : `<number:day${v.length > 1 ? long : ''}/>`;
        else if (/^m/.test(v)) body += v.length >= 4 ? `<number:month number:textual="true"${long}/>` : v.length === 3 ? '<number:month number:textual="true"/>' : `<number:month${v.length > 1 ? long : ''}/>`;
        else if (/^[ye]/.test(v)) body += `<number:year${v.length > 2 || v === 'e' ? long : ''}/>`;
        else lit += v;
      }
    }
    flush();
    if (!body) return null;
    return { kind, body, color, elapsed: t.some(o => o.elapsed) };
  }
  const EL = { number: 'number:number-style', percentage: 'number:percentage-style', currency: 'number:currency-style', date: 'number:date-style', time: 'number:time-style', text: 'number:text-style' };
  /** Excel number format → {kind, xml} with style name `name` (null for General / unknown) */
  ODS.numberStyle = (nf, name) => {
    if (!nf || nf === 'General') return null;
    let secs = splitSections(nf); if (secs.length > 1 && secs[0].replace(/\[[^\]]*\]/g, '').trim() === '') secs = [secs[1]];
    const main = sectionStyle(secs[0]); if (!main) return null;
    const one = (nm, st, extra = '') => `<${EL[st.kind]} style:name="${nm}"${st.kind === 'time' && st.elapsed ? ' number:truncate-on-overflow="false"' : ''}${st.kind === 'date' ? ' number:automatic-order="false"' : ''}>${st.color ? `<style:text-properties fo:color="${st.color}"/>` : ''}${st.body}${extra}</${EL[st.kind]}>`;
    if (secs.length > 1 && (main.kind === 'number' || main.kind === 'currency' || main.kind === 'percentage')) {
      const neg = sectionStyle(secs[1]);
      if (neg && neg.kind === main.kind) return { kind: main.kind, xml: one(name + 'P0', main) + one(name, neg, `<style:map style:condition="value()&gt;=0" style:apply-style-name="${name}P0"/>`) };
    }
    return { kind: main.kind, xml: one(name, main) };
  };

  /* ================= cell styles ================= */
  const BW = { thin: '0.75pt solid', hair: '0.5pt solid', medium: '1.75pt solid', thick: '2.5pt solid', double: '2.25pt double', dashed: '0.75pt dashed', dotted: '0.75pt dotted', mediumDashed: '1.75pt dashed', dashDot: '0.75pt dashed', mediumDashDot: '1.75pt dashed', dashDotDot: '0.75pt dotted', slantDashDot: '1.75pt dashed' };
  const border = (b) => `${BW[b.style] || BW.thin} ${b.color || '#000000'}`;
  const HAL = { left: 'start', center: 'center', right: 'end', justify: 'justify', fill: 'start', centerContinuous: 'center', distributed: 'justify' };
  function cellStyleXML(name, st, dataStyle, font) {
    const tp = [], cp = [], pp = [];
    if (font) tp.push(`style:font-name="${x(font)}"`);
    if (st.sz) tp.push(`fo:font-size="${st.sz}pt"`);
    if (st.b) tp.push('fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"');
    if (st.i) tp.push('fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"');
    if (st.u) tp.push(`style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"${st.u === 'double' ? ' style:text-underline-type="double"' : ''}`);
    if (st.s) tp.push('style:text-line-through-style="solid" style:text-line-through-type="single"');
    if (st.color) tp.push(`fo:color="${st.color}"`);
    if (st.fill) cp.push(`fo:background-color="${st.fill}"`);
    if (st.bT) cp.push(`fo:border-top="${border(st.bT)}"`); if (st.bB) cp.push(`fo:border-bottom="${border(st.bB)}"`);
    if (st.bL) cp.push(`fo:border-left="${border(st.bL)}"`); if (st.bR) cp.push(`fo:border-right="${border(st.bR)}"`);
    if (st.wrap) cp.push('fo:wrap-option="wrap"');
    if (st.v) cp.push(`style:vertical-align="${st.v === 'center' ? 'middle' : st.v}"`);
    if (st.h) cp.push('style:text-align-source="fix" style:repeat-content="false"');
    if (st.locked === false || st.fh) cp.push(`style:cell-protect="${st.locked === false ? (st.fh ? 'formula-hidden' : 'none') : 'protected formula-hidden'}"`);
    if (st.h) pp.push(`fo:text-align="${HAL[st.h] || 'start'}"`);
    if (st.indent) pp.push(`fo:margin-left="${(st.indent * 0.0938).toFixed(4)}in"`);
    return `<style:style style:name="${name}" style:family="table-cell" style:parent-style-name="Default"${dataStyle ? ` style:data-style-name="${dataStyle}"` : ''}>${cp.length ? `<style:table-cell-properties ${cp.join(' ')}/>` : ''}${pp.length ? `<style:paragraph-properties ${pp.join(' ')}/>` : ''}${tp.length ? `<style:text-properties ${tp.join(' ')}/>` : ''}</style:style>`;
  }

  /* ================= values ================= */
  const p2 = (n) => String(n).padStart(2, '0');
  function dateISO(v) { const d = SU.serialToDate(v); return `${String(d.getUTCFullYear()).padStart(4, '0')}-${p2(d.getUTCMonth() + 1)}-${p2(d.getUTCDate())}` + (v % 1 ? `T${p2(d.getUTCHours())}:${p2(d.getUTCMinutes())}:${p2(d.getUTCSeconds())}` : ''); }
  function durISO(v) { const neg = v < 0; v = Math.abs(v); const secs = Math.round(v * 86400); const h = Math.floor(secs / 3600), m = Math.floor(secs % 3600 / 60), s = secs % 60; return (neg ? '-' : '') + `PT${h}H${p2(m)}M${p2(s)}S`; }
  function textP(t) {
    return String(t).split(/\r?\n/).map(line => '<text:p>' + x(line).replace(/ {2,}/g, (m) => ' <text:s text:c="' + (m.length - 1) + '"/>').replace(/\t/g, '<text:tab/>') + '</text:p>').join('');
  }
  function linkP(t, url) {
    let href = String(url);
    const m = /^#?(?:(.+)!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$/.exec(href);
    if (href[0] === '#' || (m && !/^[a-z]+:/i.test(href))) { if (m) href = '#' + (m[1] ? SU.unquoteSheet(m[1]) + '.' : '') + m[2].replace(/\$/g, ''); }
    return `<text:p><text:a xlink:href="${x(href)}" xlink:type="simple">${x(t)}</text:a></text:p>`;
  }

  /* ================= validations ================= */
  function validationXML(v, name, sheet) {
    const g = v.range; const base = (/^[A-Za-z_][A-Za-z0-9_]*$/.test(sheet.name) ? sheet.name : "'" + sheet.name.replace(/'/g, "''") + "'") + '.' + SU.addr(g.r0, g.c0);
    const num = (a) => { if (a === undefined || a === null || a === '') return '0'; const s = String(a); if (s[0] === '=') return ODS.formula(s).slice(4); return isNaN(+s) ? '"' + s.replace(/"/g, '""') + '"' : String(+s); };
    /* LibreOffice writes and expects ',' between the two limits here */
    const opCond = (fn) => { const a = num(v.v1), b = num(v.v2); switch (v.op || 'between') { case 'between': return `${fn}-is-between(${a},${b})`; case 'notBetween': return `${fn}-is-not-between(${a},${b})`; case 'eq': return `${fn}()=${a}`; case 'ne': return `${fn}()!=${a}`; case 'gt': return `${fn}()>${a}`; case 'lt': return `${fn}()<${a}`; case 'ge': return `${fn}()>=${a}`; case 'le': return `${fn}()<=${a}`; } return `${fn}-is-between(${a},${b})`; };
    let cond = null, extra = '';
    if (v.type === 'list') {
      const src = String(v.source || '');
      if (src[0] === '=') cond = `of:cell-content-is-in-list(${ODS.formula(src).slice(4)})`;
      else cond = `of:cell-content-is-in-list(${src.split(',').map(s => '"' + s.trim().replace(/"/g, '""') + '"').join(';')})`;
      extra = ` table:display-list="${v.dropdown === false ? 'none' : 'unsorted'}"`;
    } else if (v.type === 'whole') cond = 'of:cell-content-is-whole-number() and ' + opCond('cell-content');
    else if (v.type === 'decimal') cond = 'of:cell-content-is-decimal-number() and ' + opCond('cell-content');
    else if (v.type === 'textLength') cond = 'of:' + opCond('cell-content-text-length');
    else if (v.type === 'custom' && v.formula) cond = `of:is-true-formula(${ODS.formula(v.formula).slice(4)})`;
    if (!cond) return null;
    const mt = { stop: 'stop', warning: 'warning', information: 'information', info: 'information' }[v.errorStyle || 'stop'] || 'stop';
    return `<table:content-validation table:name="${name}" table:condition="${x(cond)}" table:allow-empty-cell="${v.ignoreBlank !== false}" table:base-cell-address="${x(base)}"${extra}>` +
      (v.prompt || v.promptTitle ? `<table:help-message table:title="${x(v.promptTitle || '')}" table:display="true">${textP(v.prompt || '')}</table:help-message>` : '') +
      `<table:error-message table:title="${x(v.errorTitle || '')}" table:display="true" table:message-type="${mt}">${textP(v.errorMsg || 'That value isn’t allowed in this cell.')}</table:error-message></table:content-validation>`;
  }

  /* ================= page ================= */
  const PAPER = { A4: [8.2681, 11.6929], Letter: [8.5, 11], Legal: [8.5, 14], A3: [11.6929, 16.5354], A5: [5.8268, 8.2681] };
  const MARG = { normal: [0.75, 0.7], narrow: [0.75, 0.25], wide: [1, 1] };

  /* ================= write ================= */
  ODS.lost = (wb) => {
    const L = []; const any = (f) => wb.sheets.some(f);
    if (any(s => s.charts.length)) L.push('charts'); if (any(s => s.images.length)) L.push('pictures'); if (any(s => s.cf.length)) L.push('conditional formatting');
    if (any(s => s.sparklines.length)) L.push('sparklines'); if (any(s => s.pivots.length)) L.push('PivotTable settings (the numbers are kept)'); if (any(s => s.tables.length)) L.push('table filters (table colours are kept)');
    return L;
  };
  ODS.write = async (wb) => {
    const fonts = new Set(); const defFont = SU.defFont || 'Calibri'; fonts.add(defFont);
    const dataStyles = new Map(); // nf -> {name, kind, xml}
    const cellStyles = new Map(); // json -> name
    const colStyles = new Map(), rowStyles = new Map(), tableStyles = new Map(), masters = new Map();
    const valXML = []; let valN = 0;
    const dsFor = (nf) => { if (!nf || nf === 'General') return null; if (dataStyles.has(nf)) return dataStyles.get(nf); const name = 'N' + (dataStyles.size + 100); let r = null; try { r = ODS.numberStyle(nf, name); } catch (e) { console.warn('number format', nf, e); } const o = r ? { name, kind: r.kind, xml: r.xml } : null; dataStyles.set(nf, o); return o; };
    const csFor = (st) => {
      const font = st.font && st.font !== defFont ? st.font : null; if (font) fonts.add(font);
      const ds = dsFor(st.nf);
      const k = JSON.stringify([st.b, st.i, st.u, st.s, st.sz, st.color, st.fill, st.bT, st.bB, st.bL, st.bR, st.wrap, st.v, st.h, st.indent, st.locked, st.fh, font, ds && ds.name]);
      if (k === '[null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null]') return { name: null, ds };
      let name = cellStyles.get(k); if (!name) { name = 'ce' + (cellStyles.size + 1); cellStyles.set(k, name); cellStyles.set(name + '#xml', cellStyleXML(name, st, ds && ds.name, font)); }
      return { name, ds };
    };
    const colStyle = (px) => { const k = Math.round(px); if (!colStyles.has(k)) colStyles.set(k, 'co' + (colStyles.size + 1)); return colStyles.get(k); };
    const rowStyle = (px, custom) => { const k = Math.round(px) + (custom ? 'c' : 'o'); if (!rowStyles.has(k)) rowStyles.set(k, { name: 'ro' + (rowStyles.size + 1), px: Math.round(px), custom }); return rowStyles.get(k).name; };
    const masterFor = (pg) => { const k = (PAPER[pg.paper] ? pg.paper : 'A4') + '|' + (pg.orientation === 'landscape' ? 'landscape' : 'portrait') + '|' + (MARG[pg.margins] ? pg.margins : 'normal'); if (!masters.has(k)) masters.set(k, 'BTS' + (masters.size + 1)); return masters.get(k); };
    const tableStyle = (s) => { const k = (s.tabColor || '') + '|' + masterFor(s.page || {}); if (!tableStyles.has(k)) tableStyles.set(k, { name: 'ta' + (tableStyles.size + 1), tab: s.tabColor, master: masterFor(s.page || {}) }); return tableStyles.get(k).name; };

    const tables = [];
    for (const s of wb.sheets) {
      const u = s.usedRangeAll();
      let maxR = u.r1, maxC = u.c1;
      for (const k of [...s.comments.keys(), ...s.threads.keys(), ...s.links.keys()]) { if (typeof k !== 'number') continue; const p = SU.unkey(k); maxR = Math.max(maxR, p.r); maxC = Math.max(maxC, p.c); }
      for (const c of [...s.colW.keys(), ...s.hiddenC]) if (c < 1024) maxC = Math.max(maxC, c);
      for (const r of [...s.rowH.keys(), ...s.hiddenR]) if (r < 100000) maxR = Math.max(maxR, r);
      // validations (cells inside get table:content-validation-name)
      const vals = [];
      for (const v of s.validations) { const name = 'val' + (++valN); const xml = validationXML(v, name, s); if (xml) { valXML.push(xml); vals.push({ g: v.range, name }); maxR = Math.max(maxR, Math.min(v.range.r1, maxR + 200, 5000)); maxC = Math.max(maxC, Math.min(v.range.c1, 200)); } }
      const valAt = (r, c) => { for (let i = vals.length - 1; i >= 0; i--) if (SU.inRange(vals[i].g, r, c)) return vals[i].name; return null; };
      // merges
      const mTop = new Map(), covered = new Set();
      for (const m of s.merges) { mTop.set(key(m.r0, m.c0), m); for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) if (r !== m.r0 || c !== m.c0) covered.add(key(r, c)); }
      let out = `<table:table table:name="${x(s.name)}" table:style-name="${tableStyle(s)}"${s.protected ? ' table:protected="true"' : ''}${s.page && s.page.printArea ? ` table:print-ranges="${x(refAddr(SU.quoteSheet(s.name) + '!', SU.addr(s.page.printArea.r0, s.page.printArea.c0) + ':' + SU.addr(s.page.printArea.r1, s.page.printArea.c1)))}"` : ''}>`;
      // columns
      let c = 0; const nCols = Math.max(0, maxC) + 1;
      while (c < nCols) { const w = s.rawColWidth(c), hid = s.hiddenC.has(c); let n = 1; while (c + n < nCols && s.rawColWidth(c + n) === w && s.hiddenC.has(c + n) === hid) n++; out += `<table:table-column table:style-name="${colStyle(w)}"${n > 1 ? ` table:number-columns-repeated="${n}"` : ''}${hid ? ' table:visibility="collapse"' : ''} table:default-cell-style-name="Default"/>`; c += n; }
      // rows
      let prevRow = null, prevN = 0; const rowsXML = [];
      const pushRow = (xml) => { if (xml === prevRow) { prevN++; return; } if (prevRow !== null) rowsXML.push(prevN > 1 ? prevRow.replace('<table:table-row ', `<table:table-row table:number-rows-repeated="${prevN}" `) : prevRow); prevRow = xml; prevN = 1; };
      for (let r = 0; r <= Math.max(0, maxR); r++) {
        const rs = rowStyle(s.rawRowHeight(r), s.rowH.has(r));
        let cells = ''; let pend = null, pendN = 0;
        const flushEmpty = () => { if (pend !== null) { cells += pendN > 1 ? pend.replace(/^<(table:(?:covered-)?table-cell)/, `<$1 table:number-columns-repeated="${pendN}"`) : pend; pend = null; pendN = 0; } };
        for (let c = 0; c < nCols; c++) {
          const k = key(r, c); const st0 = s.style(r, c); const tbl = IO.tableCellStyle ? IO.tableCellStyle(s, r, c) : null;
          const st = tbl ? { ...tbl, ...Object.fromEntries(Object.entries(st0).filter(([, v]) => v !== undefined)) } : st0;
          const { name: sn, ds } = csFor(st); const vn = valAt(r, c);
          const attrs = (sn ? ` table:style-name="${sn}"` : '') + (vn ? ` table:content-validation-name="${vn}"` : '');
          if (covered.has(k)) { const e = `<table:covered-table-cell${attrs}/>`; if (e === pend) pendN++; else { flushEmpty(); pend = e; pendN = 1; } continue; }
          const empty = s.isEmpty(r, c); const note = s.comments.get(k), thr = s.threads.get(k), link = s.links.get(k); const m = mTop.get(k);
          if (empty && !note && !thr && !m) { const e = `<table:table-cell${attrs}/>`; if (e === pend) pendN++; else { flushEmpty(); pend = e; pendN = 1; } continue; }
          flushEmpty();
          let a = attrs + (m ? ` table:number-columns-spanned="${m.c1 - m.c0 + 1}" table:number-rows-spanned="${m.r1 - m.r0 + 1}"` : '');
          let inner = '';
          if (note || thr) {
            const txt = thr && thr.items && thr.items.length ? thr.items.map(it => (it.author ? it.author + ': ' : '') + it.text).join('\n') : String(note || '');
            const author = thr && thr.items && thr.items[0] && thr.items[0].author || '';
            inner += `<office:annotation>${author ? `<dc:creator>${x(author)}</dc:creator>` : ''}${textP(txt)}</office:annotation>`;
          }
          if (!empty) {
            const raw = s.raw(r, c), v = s.value(r, c); const disp = s.displayText(r, c);
            if (typeof raw === 'string' && raw[0] === '=') { try { a += ` table:formula="${x(ODS.formula(raw))}"`; } catch (e) { console.warn('formula', raw, e); } }
            if (SU.isErr(v)) a += ' office:value-type="string" office:string-value=""';
            else if (typeof v === 'number' && isFinite(v)) {
              const kind = ds ? ds.kind : 'float';
              if (kind === 'date' && v >= 0) a += ` office:value-type="date" office:date-value="${dateISO(v)}"`;
              else if (kind === 'time') a += ` office:value-type="time" office:time-value="${durISO(v)}"`;
              else if (kind === 'percentage') a += ` office:value-type="percentage" office:value="${v}"`;
              else if (kind === 'currency') { const sym = (/"([^"]+)"|\[\$([^\]-]+)|(\$)/.exec(st.nf) || []).slice(1).find(Boolean); const code = CUR[(sym || '').trim()]; a += code ? ` office:value-type="currency" office:currency="${code[0]}" office:value="${v}"` : ` office:value-type="float" office:value="${v}"`; }
              else a += ` office:value-type="float" office:value="${v}"`;
            } else if (typeof v === 'boolean') a += ` office:value-type="boolean" office:boolean-value="${v}"`;
            else a += ' office:value-type="string"' + (typeof raw === 'string' && raw[0] === '=' ? ` office:string-value="${x(v ?? '')}"` : '');
            inner += link ? linkP(disp, link) : textP(disp);
          }
          cells += `<table:table-cell${a}>${inner}</table:table-cell>`;
        }
        // trailing empty unstyled cells can be dropped
        if (pend !== null && pend.startsWith('<table:table-cell') && !/ table:(style|content-validation)-name=/.test(pend)) { pend = null; pendN = 0; } else flushEmpty();
        pushRow(`<table:table-row table:style-name="${rs}"${s.hiddenR.has(r) ? ' table:visibility="collapse"' : ''}>${cells || `<table:table-cell table:number-columns-repeated="${nCols}"/>`}</table:table-row>`);
      }
      pushRow(null); out += rowsXML.join('') + '</table:table>';
      tables.push(out);
    }

    // named ranges
    let names = '';
    try {
      const list = wb.hf.getAllNamedExpressionsSerialized().filter(n => n.scope === undefined);
      const items = list.map(n => { const f = SU.fromHF(n.expression); const conv = ODS.formula(f).slice(4); const single = /^\[([^\[\]]+)\]$/.exec(conv); const base = wb.sheets[0] ? ofSheet(wb.sheets[0].name) + '.$A$1' : '$Sheet1.$A$1'; return single ? `<table:named-range table:name="${x(n.name)}" table:base-cell-address="${x(base)}" table:cell-range-address="${x(single[1])}"/>` : `<table:named-expression table:name="${x(n.name)}" table:base-cell-address="${x(base)}" table:expression="${x('of:=' + conv)}"/>`; });
      if (items.length) names = `<table:named-expressions>${items.join('')}</table:named-expressions>`;
    } catch (e) { console.warn('names', e); }

    const fontDecls = `<office:font-face-decls>${[...fonts].map(f => `<style:font-face style:name="${x(f)}" svg:font-family="${x(/[\s,]/.test(f) ? "'" + f + "'" : f)}"/>`).join('')}</office:font-face-decls>`;
    const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:tableooo="http://openoffice.org/2009/table" office:version="1.3"';
    const auto = [
      ...[...colStyles].map(([px, n]) => `<style:style style:name="${n}" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="${inch(px)}"/></style:style>`),
      ...[...rowStyles.values()].map(o => `<style:style style:name="${o.name}" style:family="table-row"><style:table-row-properties style:row-height="${inch(o.px)}" fo:break-before="auto" style:use-optimal-row-height="${o.custom ? 'false' : 'true'}"/></style:style>`),
      ...[...tableStyles.values()].map(o => `<style:style style:name="${o.name}" style:family="table" style:master-page-name="${o.master}"><style:table-properties table:display="true" style:writing-mode="lr-tb"${o.tab ? ` table:tab-color="${o.tab}" tableooo:tab-color="${o.tab}"` : ''}/></style:style>`),
      ...[...dataStyles.values()].filter(Boolean).map(o => o.xml),
      ...[...cellStyles].filter(([k]) => k.endsWith('#xml')).map(([, v]) => v),
    ].join('');
    /* calculation settings = Excel behaviour: MATCH/COUNTIF etc. are case-insensitive, whole-cell, wildcards not regular expressions (ODF defaults differ) */
    const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS}><office:scripts/>${fontDecls}<office:automatic-styles>${auto}</office:automatic-styles><office:body><office:spreadsheet><table:calculation-settings table:case-sensitive="false" table:search-criteria-must-apply-to-whole-cell="true" table:use-wildcards="true" table:use-regular-expressions="false" table:automatic-find-labels="false"/>${valXML.length ? `<table:content-validations>${valXML.join('')}</table:content-validations>` : ''}${tables.join('')}${names}</office:spreadsheet></office:body></office:document-content>`;
    const layouts = [...masters].map(([k, n]) => { const [paper, orient, marg] = k.split('|'); let [w, h] = PAPER[paper]; if (orient === 'landscape') [w, h] = [h, w]; const [mv, mh] = MARG[marg]; return { n, xml: `<style:page-layout style:name="pl_${n}"><style:page-layout-properties fo:page-width="${w}in" fo:page-height="${h}in" style:print-orientation="${orient}" fo:margin-top="${mv}in" fo:margin-bottom="${mv}in" fo:margin-left="${mh}in" fo:margin-right="${mh}in"/></style:page-layout>` }; });
    const styles = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS}>${fontDecls}<office:styles><style:default-style style:family="table-cell"><style:paragraph-properties style:tab-stop-distance="0.5in"/><style:text-properties style:font-name="${x(defFont)}" fo:font-size="${SU.defSize}pt" fo:language="en" fo:country="IN"/></style:default-style><style:style style:name="Default" style:family="table-cell"><style:text-properties style:font-name="${x(defFont)}" fo:font-size="${SU.defSize}pt"/></style:style></office:styles><office:automatic-styles>${layouts.map(l => l.xml).join('')}</office:automatic-styles><office:master-styles>${layouts.map(l => `<style:master-page style:name="${l.n}" style:page-layout-name="pl_${l.n}"><style:header style:display="false"/><style:footer style:display="false"/></style:master-page>`).join('')}</office:master-styles></office:document-styles>`;
    const P = BTS.props || {}; const now = new Date().toISOString().replace(/\.\d+Z$/, '');
    const meta = `<?xml version="1.0" encoding="UTF-8"?><office:document-meta ${NS}><office:meta><meta:generator>BTS Sheets</meta:generator>${P.title ? `<dc:title>${x(P.title)}</dc:title>` : ''}${P.subject ? `<dc:subject>${x(P.subject)}</dc:subject>` : ''}${P.comments ? `<dc:description>${x(P.comments)}</dc:description>` : ''}${P.keywords ? `<meta:keyword>${x(P.keywords)}</meta:keyword>` : ''}${P.author || BTS.userName ? `<meta:initial-creator>${x(P.author || BTS.userName)}</meta:initial-creator><dc:creator>${x(BTS.userName || P.author)}</dc:creator>` : ''}<meta:creation-date>${now}</meta:creation-date><dc:date>${now}</dc:date></office:meta></office:document-meta>`;
    const item = (n, t, v) => `<config:config-item config:name="${n}" config:type="${t}">${x(v)}</config:config-item>`;
    const sheetsCfg = wb.sheets.map(s => { const fr = s.freeze || { r: 0, c: 0 }; return `<config:config-item-map-entry config:name="${x(s.name)}">${item('CursorPositionX', 'int', s.active ? s.active.c : 0)}${item('CursorPositionY', 'int', s.active ? s.active.r : 0)}${item('HorizontalSplitMode', 'short', fr.c ? 2 : 0)}${item('VerticalSplitMode', 'short', fr.r ? 2 : 0)}${item('HorizontalSplitPosition', 'int', fr.c || 0)}${item('VerticalSplitPosition', 'int', fr.r || 0)}${item('ActiveSplitRange', 'short', fr.r ? 2 : fr.c ? 3 : 2)}${item('PositionLeft', 'int', 0)}${item('PositionRight', 'int', fr.c || 0)}${item('PositionTop', 'int', 0)}${item('PositionBottom', 'int', fr.r || 0)}${item('ShowGrid', 'boolean', s.showGrid !== false)}</config:config-item-map-entry>`; }).join('');
    const settings = `<?xml version="1.0" encoding="UTF-8"?><office:document-settings ${NS}><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>${item('ViewId', 'string', 'view1')}<config:config-item-map-named config:name="Tables">${sheetsCfg}</config:config-item-map-named>${item('ActiveTable', 'string', wb.active ? wb.active.name : '')}${item('ShowGrid', 'boolean', wb.active ? wb.active.showGrid !== false : true)}${item('HasColumnRowHeaders', 'boolean', true)}${item('HasSheetTabs', 'boolean', true)}</config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set></office:settings></office:document-settings>`;
    const manifest = `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/></manifest:manifest>`;
    const zip = new JSZip();
    zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet', { compression: 'STORE' });
    zip.file('content.xml', content); zip.file('styles.xml', styles); zip.file('meta.xml', meta); zip.file('settings.xml', settings);
    zip.file('META-INF/manifest.xml', manifest);
    return await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  };

  /* ================= reading: extra detail SheetJS doesn't give us for .ods ================= */
  /** OpenFormula → Excel A1 syntax */
  ODS.fromOF = (f) => {
    let s = String(f).replace(/^(?:of|oooc|msoxl):/i, ''); if (s[0] === '=') s = s.slice(1); if (s[0] === '=') s = s.slice(1);
    let out = '', i = 0, brace = 0; const n = s.length;
    const cellPart = (p) => { const m = /^(\$?(?:'(?:[^']|'')+'|[^.']*))?\.(\$?[A-Za-z]+\$?\d+|\$?[A-Za-z]+|\$?\d+)$/.exec(p); if (!m) return null; let sh = m[1] || ''; if (sh[0] === '$') sh = sh.slice(1); return { sh: sh ? SU.unquoteSheet(sh) : '', cell: m[2].toUpperCase() }; };
    const ref = (inner) => {
      if (/^#/.test(inner) || /#REF!/i.test(inner)) return '#REF!';
      const parts = []; let cur = '', q = false; for (const ch of inner) { if (ch === "'") q = !q; if (ch === ':' && !q) { parts.push(cur); cur = ''; } else cur += ch; } parts.push(cur);
      const a = cellPart(parts[0]); if (!a) return null; const b = parts[1] !== undefined ? cellPart(parts[1]) : null; if (parts[1] !== undefined && !b) return null;
      const pre = a.sh ? SU.quoteSheet(a.sh) + '!' : '';
      if (!b) return pre + a.cell;
      // whole columns / rows written out in full
      const ca = /^(\$?[A-Z]+)\$?1$/.exec(a.cell), cb = /^(\$?[A-Z]+)\$?(1048576|65536)$/.exec(b.cell);
      if (ca && cb) return pre + ca[1] + ':' + cb[1];
      const ra = /^\$?A(\$?\d+)$/.exec(a.cell), rb = /^\$?(?:AMJ|XFD)(\$?\d+)$/.exec(b.cell);
      if (ra && rb) return pre + ra[1] + ':' + rb[1];
      return pre + a.cell + ':' + b.cell;
    };
    while (i < n) {
      const ch = s[i];
      if (ch === '"') { let j = i + 1; while (j < n) { if (s[j] === '"') { if (s[j + 1] === '"') { j += 2; continue; } break; } j++; } out += s.slice(i, j + 1); i = j + 1; continue; }
      if (ch === '[') { let j = i + 1, q = false; while (j < n && (s[j] !== ']' || q)) { if (s[j] === "'") q = !q; j++; } const r = ref(s.slice(i + 1, j)); out += r === null ? s.slice(i, j + 1) : r; i = j + 1; continue; }
      if (ch === '{') { brace++; out += ch; i++; continue; }
      if (ch === '}') { brace = Math.max(0, brace - 1); out += ch; i++; continue; }
      if (ch === ';') { out += ','; i++; continue; }
      if (ch === '|' && brace) { out += ';'; i++; continue; }
      if (ch === '~') { out += ','; i++; continue; }
      if (/[A-Za-z_]/.test(ch) && !/[\w.]/.test(s[i - 1] || '')) { const m = /^[A-Za-z_][\w.]*/.exec(s.slice(i)); let w = m[0]; i += w.length; if (s[i] === '(') w = w.replace(/^COM\.MICROSOFT\./i, '').replace(/^_xlfn\./i, ''); out += w; continue; }
      out += ch; i++;
    }
    return '=' + out;
  };

  const NSO = { office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0', table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0', style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0', fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', number: 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0', svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0', config: 'urn:oasis:names:tc:opendocument:xmlns:config:1.0' };
  const at = (el, ns, n) => el.getAttributeNS(NSO[ns], n);
  const kids = (el) => Array.from(el.children || []);
  const lenPx = (v) => { const m = /^(-?[\d.]+)(in|cm|mm|pt|pc|px)$/.exec(String(v || '').trim()); if (!m) return null; return +m[1] * { in: 96, cm: 96 / 2.54, mm: 96 / 25.4, pt: 96 / 72, pc: 16, px: 1 }[m[2]]; };
  /** ODF data style element → Excel number format code */
  function dataStyleCode(el, all, depth = 0) {
    const q = (t) => /^[\s\-+()/:]+$/.test(t) ? t : el.localName === 'percentage-style' && t.trim() === '%' ? t : '"' + t.replace(/"/g, '') + '"';
    let code = '', color = '';
    for (const c of kids(el)) {
      const ln = c.localName; const long = at(c, 'number', 'style') === 'long';
      if (ln === 'text-properties') { const fc = (at(c, 'fo', 'color') || '').toUpperCase(); color = { '#FF0000': '[Red]', '#0000FF': '[Blue]', '#00FF00': '[Green]' }[fc] || ''; }
      else if (ln === 'number') { const d = +(at(c, 'number', 'decimal-places') || 0); const mind = at(c, 'number', 'min-decimal-places'); const md = mind === null || mind === '' ? d : +mind; const mi = +(at(c, 'number', 'min-integer-digits') || 0); const g = at(c, 'number', 'grouping') === 'true'; const ip = g ? '#,##' + '0'.repeat(Math.max(1, mi)).slice(0, Math.max(1, mi)) : (mi ? '0'.repeat(mi) : '#'); const dp = d ? '.' + '0'.repeat(md) + '#'.repeat(Math.max(0, d - md)) : ''; const fac = +(at(c, 'number', 'display-factor') || 1); code += ip + dp + ','.repeat(fac >= 1000 ? Math.round(Math.log10(fac) / 3) : 0); }
      else if (ln === 'scientific-number') { const d = +(at(c, 'number', 'decimal-places') || 0); code += '0' + (d ? '.' + '0'.repeat(d) : '') + 'E+' + '0'.repeat(+(at(c, 'number', 'min-exponent-digits') || 2)); }
      else if (ln === 'fraction') { code += '# ' + '?'.repeat(+(at(c, 'number', 'min-numerator-digits') || 1)) + '/' + (at(c, 'number', 'denominator-value') || '?'.repeat(+(at(c, 'number', 'min-denominator-digits') || 1))); }
      else if (ln === 'currency-symbol') code += '"' + c.textContent + '"';
      else if (ln === 'text') code += q(c.textContent);
      else if (ln === 'text-content') code += '@';
      else if (ln === 'day') code += long ? 'dd' : 'd';
      else if (ln === 'day-of-week') code += long ? 'dddd' : 'ddd';
      else if (ln === 'month') code += at(c, 'number', 'textual') === 'true' ? (long ? 'mmmm' : 'mmm') : (long ? 'mm' : 'm');
      else if (ln === 'year') code += long ? 'yyyy' : 'yy';
      else if (ln === 'hours') code += at(el, 'number', 'truncate-on-overflow') === 'false' ? '[h]' : (long ? 'hh' : 'h');
      else if (ln === 'minutes') code += long ? 'mm' : 'm';
      else if (ln === 'seconds') { const d = +(at(c, 'number', 'decimal-places') || 0); code += (long ? 'ss' : 's') + (d ? '.' + '0'.repeat(d) : ''); }
      else if (ln === 'am-pm') code += 'AM/PM';
    }
    code = color + code;
    if (depth < 2) for (const m of kids(el).filter(c => c.localName === 'map')) { const tgt = all.get(at(m, 'style', 'apply-style-name')); const cond = at(m, 'style', 'condition') || ''; if (tgt && /value\(\)\s*>=?\s*0/.test(cond)) return dataStyleCode(tgt, all, depth + 1) + ';' + code; }
    return code;
  }
  const HAL_IN = { start: 'left', left: 'left', center: 'center', end: 'right', right: 'right', justify: 'justify' };
  const borderIn = (v) => { if (!v || v === 'none') return null; const m = /([\d.]+)(pt|in|cm|mm|px)?\s+(solid|double|dashed|dotted|fine-dashed|dash-dot|dash-dot-dot)?\s*(#[0-9a-fA-F]{6})?/.exec(v); if (!m) return null; const w = lenPx(m[1] + (m[2] || 'pt')) || 1; const st = m[3] === 'double' ? 'double' : /dash/.test(m[3] || '') && m[3] !== 'dash-dot-dot' ? 'dashed' : /dot/.test(m[3] || '') ? 'dotted' : w >= 3 ? 'thick' : w >= 1.8 ? 'medium' : 'thin'; return { style: st, color: (m[4] || '#000000').toUpperCase() }; };
  /* dates / durations as serial numbers, read in UTC (SheetJS shifts them by the local time zone) */
  const dateVal = (cell) => {
    const t = at(cell, 'office', 'value-type');
    if (t === 'date') { const m = /^(-?\d{4,})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d)(?::(\d\d(?:\.\d+)?))?)?/.exec(at(cell, 'office', 'date-value') || ''); if (!m) return null; return (Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), 0) + Math.round(+(m[6] || 0) * 1000) - Date.UTC(1899, 11, 30)) / 864e5; }
    if (t === 'time') { const m = /^(-)?P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?$/.exec(at(cell, 'office', 'time-value') || ''); if (!m) return null; const v = ((+(m[2] || 0) * 24 + +(m[3] || 0)) * 3600 + +(m[4] || 0) * 60 + +(m[5] || 0)) / 86400; return m[1] ? -v : v; }
    return null;
  };
  /** parse an .ods for number formats, cell formatting, sizes, formulas, names and view settings */
  ODS.read = async (data) => {
    const zip = await JSZip.loadAsync(data); const P = new DOMParser();
    const doc = async (n) => { const f = zip.file(n); return f ? P.parseFromString(await f.async('string'), 'application/xml') : null; };
    const content = await doc('content.xml'); if (!content) return null; const stylesDoc = await doc('styles.xml'); const settings = await doc('settings.xml');
    const docs = [stylesDoc, content].filter(Boolean);
    const fontFam = new Map(), dataStyles = new Map(), cellSt = new Map(), colSt = new Map(), rowSt = new Map(), tabSt = new Map();
    for (const d of docs) {
      for (const ff of d.getElementsByTagNameNS(NSO.style, 'font-face')) fontFam.set(at(ff, 'style', 'name'), (at(ff, 'svg', 'font-family') || '').replace(/^'|'$/g, ''));
      for (const el of d.getElementsByTagNameNS(NSO.number, '*')) if (/-style$/.test(el.localName)) dataStyles.set(at(el, 'style', 'name'), el);
    }
    for (const d of docs) for (const el of d.getElementsByTagNameNS(NSO.style, 'style')) {
      const fam = at(el, 'style', 'family'), name = at(el, 'style', 'name'); const prop = (ln) => kids(el).find(c => c.localName === ln);
      if (fam === 'table-column') { const p = prop('table-column-properties'); const w = p && lenPx(at(p, 'style', 'column-width')); if (w) colSt.set(name, Math.round(w)); }
      else if (fam === 'table-row') { const p = prop('table-row-properties'); if (p) rowSt.set(name, { h: lenPx(at(p, 'style', 'row-height')), custom: at(p, 'style', 'use-optimal-row-height') === 'false' }); }
      else if (fam === 'table') { const p = prop('table-properties'); if (p) { const tc = at(p, 'table', 'tab-color') || p.getAttributeNS('http://openoffice.org/2009/table', 'tab-color'); tabSt.set(name, tc ? tc.toUpperCase() : null); } }
      else if (fam === 'table-cell') {
        const o = {}; const parent = at(el, 'style', 'parent-style-name'); if (parent && parent !== 'Default') o._parent = parent;
        const ds = at(el, 'style', 'data-style-name'); if (ds && dataStyles.has(ds)) { try { const code = dataStyleCode(dataStyles.get(ds), dataStyles); if (code && !/^[#0]$/.test(code) && code !== 'General') o.nf = code; } catch { } }
        for (const c of kids(el)) {
          if (c.localName === 'text-properties') {
            const fw = at(c, 'fo', 'font-weight'); if (fw === 'bold' || +fw >= 600) o.b = true; if (at(c, 'fo', 'font-style') === 'italic') o.i = true;
            const ul = at(c, 'style', 'text-underline-style'); if (ul && ul !== 'none') o.u = at(c, 'style', 'text-underline-type') === 'double' ? 'double' : true;
            const lt = at(c, 'style', 'text-line-through-style'); if (lt && lt !== 'none') o.s = true;
            const col = at(c, 'fo', 'color'); if (col && /^#/.test(col) && col.toLowerCase() !== '#000000') o.color = col.toUpperCase();
            const fs = at(c, 'fo', 'font-size'); if (fs && /pt$/.test(fs)) { const v = +fs.slice(0, -2); if (v && v !== 11 && v !== 10) o.sz = v; }
            const fn = at(c, 'style', 'font-name'); const ffam = fn ? fontFam.get(fn) || fn : at(c, 'fo', 'font-family'); if (ffam && !/^(Liberation Sans|Arial|Calibri|Carlito)$/i.test(ffam)) o.font = ffam.replace(/^'|'$/g, '');
          } else if (c.localName === 'table-cell-properties') {
            const bg = at(c, 'fo', 'background-color'); if (bg && /^#/.test(bg)) o.fill = bg.toUpperCase();
            const all = borderIn(at(c, 'fo', 'border')); if (all) { o.bT = o.bB = o.bL = o.bR = all; }
            [['top', 'bT'], ['bottom', 'bB'], ['left', 'bL'], ['right', 'bR']].forEach(([s2, k]) => { const b = borderIn(at(c, 'fo', 'border-' + s2)); if (b) o[k] = b; });
            if (at(c, 'fo', 'wrap-option') === 'wrap') o.wrap = true; const va = at(c, 'style', 'vertical-align'); if (va && va !== 'automatic') o.v = va;
            const cp = at(c, 'style', 'cell-protect'); if (cp === 'none') o.locked = false; else if (cp === 'formula-hidden') { o.locked = false; o.fh = true; } else if (cp && /formula-hidden/.test(cp)) o.fh = true;
          } else if (c.localName === 'paragraph-properties') { const ta = at(c, 'fo', 'text-align'); if (ta && HAL_IN[ta]) o.h = HAL_IN[ta]; const ml = lenPx(at(c, 'fo', 'margin-left')); if (ml && ml > 4) o.indent = Math.round(ml / 9); }
        }
        cellSt.set(name, o);
      }
    }
    const resolve = (name, d = 0) => { const o = cellSt.get(name); if (!o) return null; const { _parent, ...r } = o; if (!_parent || d > 5) return r; return { ...(resolve(_parent, d + 1) || {}), ...r }; };
    const cache = new Map(); const styleOf = (n) => { if (!n || n === 'Default') return null; if (!cache.has(n)) { const r = resolve(n); cache.set(n, r && Object.keys(r).length ? r : null); } return cache.get(n); };
    // view settings (freeze panes, grid lines)
    const views = new Map(); let activeTable = null;
    if (settings) {
      for (const ci of settings.getElementsByTagNameNS(NSO.config, 'config-item')) if (at(ci, 'config', 'name') === 'ActiveTable') activeTable = ci.textContent;
      for (const named of settings.getElementsByTagNameNS(NSO.config, 'config-item-map-named')) {
        if (at(named, 'config', 'name') !== 'Tables') continue;
        for (const ent of kids(named)) { const v = {}; for (const ci of kids(ent)) if (ci.localName === 'config-item') v[at(ci, 'config', 'name')] = ci.textContent; views.set(at(ent, 'config', 'name'), v); }
      }
    }
    const sheets = new Map();
    for (const t of content.getElementsByTagNameNS(NSO.table, 'table')) {
      const name = at(t, 'table', 'name'); const S = { styles: new Map(), formulas: new Map(), values: new Map(), colW: new Map(), rowH: new Map(), hiddenC: new Set(), hiddenR: new Set() }; const colDef = [];
      const colEls = [], rowEls = [];
      const walk = (el) => { for (const k of kids(el)) { if (k.localName === 'table-column') colEls.push(k); else if (k.localName === 'table-row') rowEls.push(k); else if (/^table-(column-group|header-columns|columns|row-group|header-rows|rows)$/.test(k.localName)) walk(k); } };
      walk(t);
      let c = 0;
      for (const k of colEls) { const rep = +(at(k, 'table', 'number-columns-repeated') || 1); const w = colSt.get(at(k, 'table', 'style-name')); const hid = at(k, 'table', 'visibility') === 'collapse'; const dcs = at(k, 'table', 'default-cell-style-name'); for (let i = 0; i < rep && c < 1024; i++, c++) { if (w && Math.abs(w - SU.DEF_W) > 1) S.colW.set(c, w); if (hid) S.hiddenC.add(c); colDef[c] = dcs; } if (c >= 1024) break; }
      let r = 0;
      for (const row of rowEls) {
        const rep = +(at(row, 'table', 'number-rows-repeated') || 1); const rs = rowSt.get(at(row, 'table', 'style-name')); const hid = at(row, 'table', 'visibility') === 'collapse';
        const cells = kids(row).filter(k => k.localName === 'table-cell' || k.localName === 'covered-table-cell');
        const content2 = cells.some(k => at(k, 'table', 'formula') || k.firstElementChild);
        const reps = rep > 100 && !content2 ? 0 : Math.min(rep, 5000);
        for (let i = 0; i < reps; i++) {
          if (rs && rs.custom && rs.h && Math.abs(rs.h - SU.DEF_H) > 1) S.rowH.set(r + i, Math.round(rs.h)); if (hid) S.hiddenR.add(r + i);
          let cc = 0;
          for (const cell of cells) {
            const crep = +(at(cell, 'table', 'number-columns-repeated') || 1); const st = styleOf(at(cell, 'table', 'style-name') || colDef[cc]); const f = at(cell, 'table', 'formula');
            const n2 = crep > 256 && !f ? 0 : Math.min(crep, 1024);
            const dv = !f && n2 ? dateVal(cell) : null;
            for (let j = 0; j < n2; j++) { const k = SU.key(r + i, cc + j); if (st) S.styles.set(k, st); if (f) { try { S.formulas.set(k, ODS.fromOF(f)); } catch { } } else if (dv !== null) S.values.set(k, dv); }
            cc += crep;
          }
        }
        r += rep; if (r > 1048576) break;
      }
      const v = views.get(name) || {};
      S.freeze = { r: v.VerticalSplitMode === '2' ? +v.VerticalSplitPosition || 0 : 0, c: v.HorizontalSplitMode === '2' ? +v.HorizontalSplitPosition || 0 : 0 };
      if (v.ShowGrid === 'false') S.showGrid = false;
      S.tabColor = tabSt.get(at(t, 'table', 'style-name')) || null;
      sheets.set(name, S);
    }
    const names = [];
    for (const nr of content.getElementsByTagNameNS(NSO.table, 'named-range')) { try { names.push({ name: at(nr, 'table', 'name'), expression: ODS.fromOF('[' + at(nr, 'table', 'cell-range-address') + ']') }); } catch { } }
    for (const ne of content.getElementsByTagNameNS(NSO.table, 'named-expression')) { try { names.push({ name: at(ne, 'table', 'name'), expression: ODS.fromOF(at(ne, 'table', 'expression')) }); } catch { } }
    return { sheets, names, activeTable };
  };
})();
