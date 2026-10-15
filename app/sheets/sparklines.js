/* BTS Sheets — sparklines: tiny line / column / win-loss charts drawn inside cells on the grid canvas.
 * Model (per sheet): s.sparklines = [{ id, type:'line'|'column'|'winloss', color, negColor, markerColor, highColor, lowColor,
 *   show:{high,low,first,last,neg,markers}, weight, items:[{ r, c, sheet:null|name, src:{r0,c0,r1,c1} }] }] */
(function () {
  const Spark = window.Spark = {};
  const { key } = SU;
  let wb;
  Spark.init = (w) => { wb = w; };
  const S = () => wb.active;

  Spark.values = (s, it) => {
    const src = it.sheet ? wb.sheetByName(it.sheet) : s; if (!src) return [];
    const g = it.src; const out = [];
    if (g.r0 === g.r1) for (let c = g.c0; c <= g.c1; c++) { const v = src.value(g.r0, c); out.push(typeof v === 'number' ? v : null); }
    else for (let r = g.r0; r <= g.r1; r++) { const v = src.value(r, g.c0); out.push(typeof v === 'number' ? v : null); }
    return out;
  };

  /** draw one sparkline into the cell rectangle */
  Spark.draw = (ctx, s, g, it, x, y, w, h, dpr = 1) => {
    const vals = Spark.values(s, it); const nums = vals.filter(v => v !== null); if (!nums.length || w < 6 || h < 6) return;
    const pad = Math.min(4, h * 0.18), px = x + 3, pw = w - 6, py = y + pad, ph = h - pad * 2 - 1;
    const n = vals.length; const show = g.show || {};
    const color = g.color || STheme.tint(STheme.cur().accents[0], -0.2), neg = g.negColor || '#D64545';
    ctx.save(); ctx.beginPath(); ctx.rect(x, y, w - 1, h - 1); ctx.clip();
    if (g.type === 'column' || g.type === 'winloss') {
      const win = g.type === 'winloss';
      let min = Math.min(0, ...nums), max = Math.max(0, ...nums); if (win) { min = -1; max = 1; } if (max === min) max = min + 1;
      const zeroY = py + ph * (max / (max - min)); const slot = pw / n; const bw = Math.max(1, slot * 0.72);
      const hi = Math.max(...nums), lo = Math.min(...nums); const firstI = vals.findIndex(v => v !== null); let lastI = -1; vals.forEach((v, i) => { if (v !== null) lastI = i; });
      vals.forEach((v, i) => {
        if (v === null) return; let top, bot;
        if (win) { if (v === 0) return; top = v > 0 ? py : zeroY + 1; bot = v > 0 ? zeroY - 1 : py + ph; }
        else { const vy = py + ph * ((max - v) / (max - min)); top = Math.min(vy, zeroY); bot = Math.max(vy, zeroY); if (bot - top < 1) bot = top + 1; }
        let fill = color;
        if (v < 0 && (win || show.neg)) fill = neg;
        if (show.high && v === hi) fill = g.highColor || '#2F5597'; if (show.low && v === lo) fill = g.lowColor || '#E07A1F';
        if (show.first && i === firstI) fill = g.markerColor || '#7A4FB5'; if (show.last && i === lastI) fill = g.markerColor || '#7A4FB5';
        ctx.fillStyle = fill; ctx.fillRect(px + slot * i + (slot - bw) / 2, top, bw, bot - top);
      });
      if (!win && min < 0) { ctx.fillStyle = '#9aa3b2'; ctx.fillRect(px, Math.round(zeroY * dpr) / dpr, pw, 1 / dpr); }
    } else {
      let min = Math.min(...nums), max = Math.max(...nums); if (max === min) { max += 1; min -= 1; }
      const pt = (i, v) => [px + (n === 1 ? pw / 2 : pw * i / (n - 1)), py + ph * ((max - v) / (max - min))];
      ctx.strokeStyle = color; ctx.lineWidth = g.weight || 1.5; ctx.lineJoin = 'round'; ctx.lineCap = 'round'; ctx.beginPath(); let started = false;
      vals.forEach((v, i) => { if (v === null) { started = false; return; } const [a, b] = pt(i, v); if (!started) { ctx.moveTo(a, b); started = true; } else ctx.lineTo(a, b); });
      ctx.stroke();
      const hi = Math.max(...nums), lo = Math.min(...nums); const firstI = vals.findIndex(v => v !== null); let lastI = -1; vals.forEach((v, i) => { if (v !== null) lastI = i; });
      const dot = (i, v, c, r = 2.2) => { const [a, b] = pt(i, v); ctx.fillStyle = c; ctx.beginPath(); ctx.arc(a, b, r, 0, Math.PI * 2); ctx.fill(); };
      vals.forEach((v, i) => {
        if (v === null) return;
        if (show.markers) dot(i, v, g.markerColor || color, 1.7);
        if (show.neg && v < 0) dot(i, v, neg);
        if (show.first && i === firstI) dot(i, v, g.markerColor || '#7A4FB5');
        if (show.last && i === lastI) dot(i, v, g.markerColor || '#7A4FB5');
        if (show.high && v === hi) dot(i, v, g.highColor || '#2F5597', 2.5);
        if (show.low && v === lo) dot(i, v, g.lowColor || '#E07A1F', 2.5);
      });
    }
    ctx.restore();
  };

  /* ---------- create ---------- */
  Spark.dialog = async (type) => {
    if (A.protectedCheck(false, true, 'formatCells')) return;
    const s = S(); const g = A.effData(s.sel);
    const single = g.r0 === g.r1 && g.c0 === g.c1;
    const dataStr = single ? '' : SU.rangeStr(g);
    let locStr = '';
    if (!single) { if (g.r1 > g.r0 && g.c1 > g.c0) locStr = SU.rangeStr({ r0: g.r0, c0: g.c1 + 1, r1: g.r1, c1: g.c1 + 1 }); else if (g.r0 === g.r1) locStr = SU.addr(g.r0, g.c1 + 1); else locStr = SU.addr(g.r1 + 1, g.c0); }
    else locStr = SU.addr(g.r0, g.c0);
    const names = { line: 'Line', column: 'Column', winloss: 'Win/Loss' };
    const res = await BTS.dialog({ title: 'Create ' + names[type] + ' sparklines', width: 420, body: `<p class="dlg-note" style="margin:0 0 8px">Choose the data that you want, then where to place the sparklines.</p>
      <div class="field"><label>Data range</label><input id="sp-data" value="${SU.esc(dataStr)}" placeholder="e.g. B2:G10"></div>
      <div class="field"><label>Location range</label><input id="sp-loc" value="${SU.esc(locStr)}" placeholder="e.g. H2:H10"></div>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ data: r.querySelector('#sp-data').value, loc: r.querySelector('#sp-loc').value }) }] });
    if (!res) return;
    const dref = SU.parseRef(res.data), lref = SU.parseRef(res.loc);
    if (!dref || !lref) { BTS.toast('Enter a valid data range and location range.', { error: true }); return; }
    if (lref.sheet && lref.sheet !== s.name) { BTS.toast('The location range must be on the current sheet.', { error: true }); return; }
    const items = Spark.layout(dref.range, lref.range, dref.sheet && dref.sheet !== s.name ? dref.sheet : null);
    if (!items) { BTS.toast('The location range must be a single row or column that matches the data.', { error: true, ms: 4000 }); return; }
    Spark.add(s, type, items);
  };
  /** pair each location cell with a row (or column) of the data */
  Spark.layout = (d, l, sheet) => {
    const locCells = []; if (l.r0 === l.r1) for (let c = l.c0; c <= l.c1; c++) locCells.push({ r: l.r0, c }); else if (l.c0 === l.c1) for (let r = l.r0; r <= l.r1; r++) locCells.push({ r, c: l.c0 }); else return null;
    const H = d.r1 - d.r0 + 1, W = d.c1 - d.c0 + 1;
    if (locCells.length === 1) return H === 1 || W === 1 ? [{ ...locCells[0], sheet, src: d }] : null;
    if (H === locCells.length) return locCells.map((p, i) => ({ ...p, sheet, src: { r0: d.r0 + i, r1: d.r0 + i, c0: d.c0, c1: d.c1 } }));
    if (W === locCells.length) return locCells.map((p, i) => ({ ...p, sheet, src: { r0: d.r0, r1: d.r1, c0: d.c0 + i, c1: d.c0 + i } }));
    return null;
  };
  Spark.add = (s, type, items, opts = {}) => {
    const th = STheme.cur();
    const grp = { id: 'sp' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5), type, color: opts.color || STheme.tint(th.accents[0], -0.15), negColor: '#D64545', markerColor: STheme.tint(th.accents[1] || th.accents[0], -0.1), highColor: th.accents[1] || '#2F5597', lowColor: th.accents[2] || '#E07A1F', show: { high: type !== 'winloss', low: type !== 'winloss', neg: true, ...(opts.show || {}) }, items };
    const keys = new Set(items.map(i => key(i.r, i.c)));
    wb.run('Insert sparklines', () => {
      const rest = s.sparklines.map(g => ({ ...g, items: g.items.filter(i => !keys.has(key(i.r, i.c))) })).filter(g => g.items.length);
      wb.setProp(s, 'sparklines', [...rest, grp]);
    });
    Grid.render(); App.refreshContext && App.refreshContext();
    return grp;
  };
  /** the sparkline group at the active cell */
  Spark.atActive = () => { const s = S(); return Grid.sparkAt(s, s.active.r, s.active.c); };
  Spark.update = (patch, label = 'Sparkline') => {
    const s = S(); const hit = Spark.atActive(); if (!hit) return;
    wb.run(label, () => wb.setProp(s, 'sparklines', s.sparklines.map(g => g === hit.g ? { ...g, ...patch, show: { ...g.show, ...(patch.show || {}) } } : g)));
    Grid.render(); App.refreshContext && App.refreshContext();
  };
  Spark.clear = (whole) => {
    const s = S(); const hit = Spark.atActive(); if (!hit) return;
    const sel = A.sels().map(g => A.eff(g));
    wb.run('Clear sparklines', () => wb.setProp(s, 'sparklines', s.sparklines.map(g => whole ? (g === hit.g ? null : g) : { ...g, items: g.items.filter(i => !sel.some(x => SU.inRange(x, i.r, i.c))) }).filter(g => g && g.items.length)));
    Grid.render(); App.refreshContext && App.refreshContext();
  };
  Spark.editData = async () => {
    const s = S(); const hit = Spark.atActive(); if (!hit) return; const g = hit.g;
    const first = g.items[0], last = g.items[g.items.length - 1];
    const d = { r0: first.src.r0, c0: first.src.c0, r1: last.src.r1, c1: last.src.c1 }, l = { r0: first.r, c0: first.c, r1: last.r, c1: last.c };
    const res = await BTS.dialog({ title: 'Edit sparkline data', width: 420, body: `<div class="field"><label>Data range</label><input id="sp-data" value="${(first.sheet ? SU.quoteSheet(first.sheet) + '!' : '') + SU.rangeStr(d)}"></div><div class="field"><label>Location range</label><input id="sp-loc" value="${SU.rangeStr(l)}"></div>`, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ data: r.querySelector('#sp-data').value, loc: r.querySelector('#sp-loc').value }) }] });
    if (!res) return; const dref = SU.parseRef(res.data), lref = SU.parseRef(res.loc); if (!dref || !lref) { BTS.toast('That reference isn’t valid.', { error: true }); return; }
    const items = Spark.layout(dref.range, lref.range, dref.sheet && dref.sheet !== s.name ? dref.sheet : null); if (!items) { BTS.toast('The location range must match the data.', { error: true }); return; }
    wb.run('Edit sparklines', () => wb.setProp(s, 'sparklines', s.sparklines.map(x => x === g ? { ...g, items } : x))); Grid.render();
  };

  /* ---------- xlsx: x14 sparkline extension ---------- */
  const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', XM = 'http://schemas.microsoft.com/office/excel/2006/main';
  const argb = (c) => 'FF' + String(c || '#000000').replace('#', '').toUpperCase();
  Spark.toXML = (s) => {
    if (!s.sparklines.length) return '';
    const grp = s.sparklines.map(g => {
      const sh = g.show || {};
      const attrs = [g.type === 'line' ? '' : `type="${g.type === 'winloss' ? 'stacked' : 'column'}"`, sh.markers ? 'markers="1"' : '', sh.high ? 'high="1"' : '', sh.low ? 'low="1"' : '', sh.first ? 'first="1"' : '', sh.last ? 'last="1"' : '', sh.neg ? 'negative="1"' : '', 'displayEmptyCellsAs="gap"', g.weight ? `lineWeight="${(g.weight * 0.75).toFixed(2)}"` : ''].filter(Boolean).join(' ');
      const items = g.items.map(it => `<x14:sparkline><xm:f>${SU.esc(SU.quoteSheet(it.sheet || s.name))}!${SU.absRangeStr(it.src).replace(/\$/g, '')}</xm:f><xm:sqref>${SU.addr(it.r, it.c)}</xm:sqref></x14:sparkline>`).join('');
      return `<x14:sparklineGroup ${attrs}><x14:colorSeries rgb="${argb(g.color)}"/><x14:colorNegative rgb="${argb(g.negColor || '#D64545')}"/><x14:colorAxis rgb="FF000000"/><x14:colorMarkers rgb="${argb(g.markerColor || g.color)}"/><x14:colorFirst rgb="${argb(g.markerColor || g.color)}"/><x14:colorLast rgb="${argb(g.markerColor || g.color)}"/><x14:colorHigh rgb="${argb(g.highColor || g.color)}"/><x14:colorLow rgb="${argb(g.lowColor || g.color)}"/><x14:sparklines>${items}</x14:sparklines></x14:sparklineGroup>`;
    }).join('');
    return `<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14}"><x14:sparklineGroups xmlns:xm="${XM}">${grp}</x14:sparklineGroups></ext>`;
  };
  /** parse sparkline groups from a worksheet XML part (files made by Excel/LibreOffice) */
  Spark.fromXML = (xml, sheetName) => {
    const out = []; const re = /<x14:sparklineGroup\b([^>]*)>([\s\S]*?)<\/x14:sparklineGroup>/g; let m;
    const color = (body, tag) => { const mm = new RegExp(`<x14:${tag}[^>]*rgb="([0-9A-Fa-f]{8})"`).exec(body); return mm ? '#' + mm[1].slice(2).toUpperCase() : null; };
    while ((m = re.exec(xml))) {
      const a = m[1], body = m[2]; const at = (n) => (new RegExp(`\\b${n}="([^"]*)"`).exec(a) || [])[1];
      const type = at('type') === 'column' ? 'column' : at('type') === 'stacked' ? 'winloss' : 'line';
      const items = []; const ire = /<x14:sparkline>\s*<xm:f>([^<]*)<\/xm:f>\s*<xm:sqref>([^<]*)<\/xm:sqref>/g; let im;
      while ((im = ire.exec(body))) { const ref = SU.parseRef(im[1].replace(/&apos;/g, "'").replace(/&amp;/g, '&')); const loc = SU.parseAddr(im[2]); if (ref && loc) items.push({ r: loc.r, c: loc.c, sheet: ref.sheet && ref.sheet !== sheetName ? ref.sheet : null, src: ref.range }); }
      if (items.length) out.push({ id: 'sp' + Math.random().toString(36).slice(2, 9), type, color: color(body, 'colorSeries') || '#376092', negColor: color(body, 'colorNegative') || '#D64545', markerColor: color(body, 'colorMarkers'), highColor: color(body, 'colorHigh'), lowColor: color(body, 'colorLow'), show: { markers: at('markers') === '1', high: at('high') === '1', low: at('low') === '1', first: at('first') === '1', last: at('last') === '1', neg: at('negative') === '1' }, items });
    }
    return out;
  };
})();
