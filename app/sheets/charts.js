/* BTS Sheets — floating objects: Chart.js charts (live-updating) and pictures; drag, resize, select, delete */
(function () {
  const Charts = window.Charts = { selected: null, instances: new Map() };
  let wb;
  const PALETTES = {
    get 'Theme'() { return STheme.chartPalette(); },
    'BTS': ['#0F8A6A', '#2F5597', '#E07A1F', '#7A4FB5', '#D64545', '#2BA3B5', '#9AAB2E', '#C2569B'],
    'Office': ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E'],
    'Ocean': ['#03396C', '#005B96', '#6497B1', '#B3CDE0', '#011F4B', '#2A9D8F', '#48CAE4', '#90E0EF'],
    'Sunset': ['#F94144', '#F3722C', '#F8961E', '#F9C74F', '#90BE6D', '#43AA8B', '#577590', '#277DA1'],
    'Mono': ['#1B4332', '#2D6A4F', '#40916C', '#52B788', '#74C69D', '#95D5B2', '#B7E4C7', '#D8F3DC'],
  };
  Charts.PALETTES = PALETTES;
  Charts.init = (workbook) => {
    wb = workbook;
    if (window.Chart) { Chart.defaults.font.family = 'Carlito, Calibri, "Segoe UI", sans-serif'; Chart.defaults.animation = false; Chart.defaults.color = '#404040'; }
    wb.on(SU.debounce(() => Charts.refreshAll(), 60));
  };
  const S = () => wb.active;

  /* ---------- data extraction ---------- */
  Charts.TYPES = [
    ['column', 'Clustered column', 'chart-column'], ['stackedColumn', 'Stacked column', 'chart-column-stacked'], ['stackedColumn100', '100% stacked column', 'chart-column-stacked'],
    ['bar', 'Clustered bar', 'chart-bar'], ['stackedBar', 'Stacked bar', 'chart-bar-stacked'], ['stackedBar100', '100% stacked bar', 'chart-bar-stacked'],
    ['line', 'Line', 'chart-line'], ['area', 'Area', 'chart-area'], ['pie', 'Pie', 'chart-pie'], ['doughnut', 'Doughnut', 'circle-dot'],
    ['scatter', 'Scatter (XY)', 'chart-scatter'], ['bubble', 'Bubble', 'circle-ellipsis'], ['combo', 'Combo: column + line (secondary axis)', 'chart-column-increasing'],
    ['radar', 'Radar', 'radar'], ['waterfall', 'Waterfall', 'chart-candlestick'], ['histogram', 'Histogram', 'chart-no-axes-column'],
  ];
  Charts.STYLES = [[1, 'Style 1 — clean'], [2, 'Style 2 — bold with labels'], [3, 'Style 3 — outlined'], [4, 'Style 4 — dark'], [5, 'Style 5 — soft background'], [6, 'Style 6 — monochrome']];
  Charts.extract = (ch) => {
    if (ch.pivot && window.Pivot) { const d = Pivot.chartData(ch.pivot); if (d) return d; }
    const src = ch.sheet ? wb.sheetByName(ch.sheet) : ch._sheet; if (!src) return null;
    const g = ch.range; const H = g.r1 - g.r0 + 1, W = g.c1 - g.c0 + 1;
    const val = (r, c) => src.value(r, c);
    const txt = (r, c) => src.displayText(r, c);
    // header detection: first row has text (or blank top-left) and other rows numbers
    const firstRowText = (() => { let t = 0, n = 0; for (let c = g.c0; c <= g.c1; c++) { const v = val(g.r0, c); if (typeof v === 'string') t++; if (typeof v === 'number') n++; } return t > 0 && n <= (val(g.r0, g.c0) === null ? 0 : 0) + (W > 1 ? 0 : 0) || (t > 0 && n === 0); })();
    const firstColText = (() => { let t = 0, n = 0; for (let r = g.r0 + (firstRowText ? 1 : 0); r <= g.r1; r++) { const v = val(r, g.c0); if (typeof v === 'string' || (v !== null && src.style(r, g.c0).nf && SU.isDateFormat(src.style(r, g.c0).nf))) t++; else if (typeof v === 'number') n++; } return t > 0 && t >= n; })();
    const hr = firstRowText && H > 1 ? 1 : 0, hc = firstColText && W > 1 ? 1 : 0;
    let labels = [], series = [];
    const byCols = ch.switchRC ? false : (H - hr >= W - hc || ch.type === 'pie' || ch.type === 'doughnut' ? true : false);
    const seriesInCols = ch.switchRC === undefined ? (H - hr >= W - hc) : !ch.switchRC;
    if (ch.type === 'bubble') {
      const series = []; const name = hr ? txt(g.r0, g.c0 + Math.min(1, W - 1)) : 'Series 1'; const pts = [];
      for (let r = g.r0 + hr; r <= g.r1; r++) { const x = val(r, g.c0), y = W > 1 ? val(r, g.c0 + 1) : null, z = W > 2 ? val(r, g.c0 + 2) : 1; if (typeof x === 'number' && typeof y === 'number') pts.push({ x, y, z: typeof z === 'number' ? z : 1 }); }
      const mz = Math.max(1e-9, ...pts.map(p => Math.abs(p.z))); pts.forEach(p => { p.r = 4 + 22 * Math.sqrt(Math.abs(p.z) / mz); });
      series.push({ name, data: pts }); return { labels: [], series, scatter: true, xTitle: hr ? txt(g.r0, g.c0) : '' };
    }
    if (ch.type === 'histogram') {
      const nums = []; for (let r = g.r0 + hr; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) { const v = val(r, c); if (typeof v === 'number') nums.push(v); }
      if (!nums.length) return { labels: [], series: [] };
      const n = nums.length, min = Math.min(...nums), max = Math.max(...nums); const mean = nums.reduce((a, b) => a + b, 0) / n; const sd = Math.sqrt(nums.reduce((a, b) => a + (b - mean) ** 2, 0) / Math.max(1, n - 1));
      let w = ch.binWidth || (sd ? 3.5 * sd / Math.cbrt(n) : 1); let k = Math.max(1, Math.min(40, Math.ceil((max - min) / w) || 1)); if (!ch.binWidth) w = (max - min) / k || 1;
      const nice = (x) => +x.toPrecision(4); const counts = Array(k).fill(0); nums.forEach(v => { let i = Math.floor((v - min) / w); if (i >= k) i = k - 1; counts[Math.max(0, i)]++; });
      const labels = counts.map((_, i) => (i === 0 ? '[' : '(') + SU.generalNum(nice(min + i * w)) + ', ' + SU.generalNum(nice(min + (i + 1) * w)) + ']');
      return { labels, series: [{ name: 'Frequency', data: counts }], autoTitle: hr ? txt(g.r0, g.c0) : 'Histogram', xTitle: hr ? txt(g.r0, g.c0) : 'Bins' };
    }
    if (ch.type === 'scatter') {
      // first column = x, others = y series
      const xs = []; for (let r = g.r0 + hr; r <= g.r1; r++) xs.push(val(r, g.c0));
      for (let c = g.c0 + 1; c <= g.c1; c++) { const pts = []; for (let r = g.r0 + hr; r <= g.r1; r++) { const x = xs[r - g.r0 - hr], y = val(r, c); if (typeof x === 'number' && typeof y === 'number') pts.push({ x, y }); } series.push({ name: hr ? txt(g.r0, c) : 'Series ' + (c - g.c0), data: pts }); }
      if (W === 1) { const pts = xs.map((y, i) => typeof y === 'number' ? { x: i + 1, y } : null).filter(Boolean); series.push({ name: hr ? txt(g.r0, g.c0) : 'Series 1', data: pts }); }
      return { labels: [], series, scatter: true };
    }
    if (seriesInCols) {
      const rowsVis = []; for (let r = g.r0 + hr; r <= g.r1; r++) if (!src.isRowHidden(r)) rowsVis.push(r);
      for (const r of rowsVis) labels.push(hc ? txt(r, g.c0) : String(r - g.r0 - hr + 1));
      for (let c = g.c0 + hc; c <= g.c1; c++) { if (!src.colWidth(c)) continue; const d = []; for (const r of rowsVis) { const v = val(r, c); d.push(typeof v === 'number' ? v : null); } series.push({ name: hr ? txt(g.r0, c) : 'Series ' + (c - g.c0 - hc + 1), data: d }); }
    } else if (false) {
      for (let r = g.r0 + hr; r <= g.r1; r++) labels.push(hc ? txt(r, g.c0) : String(r - g.r0 - hr + 1));
      for (let c = g.c0 + hc; c <= g.c1; c++) { const d = []; for (let r = g.r0 + hr; r <= g.r1; r++) { const v = val(r, c); d.push(typeof v === 'number' ? v : null); } series.push({ name: hr ? txt(g.r0, c) : 'Series ' + (c - g.c0 - hc + 1), data: d }); }
    } else {
      for (let c = g.c0 + hc; c <= g.c1; c++) labels.push(hr ? txt(g.r0, c) : String(c - g.c0 - hc + 1));
      for (let r = g.r0 + hr; r <= g.r1; r++) { const d = []; for (let c = g.c0 + hc; c <= g.c1; c++) { const v = val(r, c); d.push(typeof v === 'number' ? v : null); } series.push({ name: hc ? txt(r, g.c0) : 'Series ' + (r - g.r0 - hr + 1), data: d }); }
    }
    if (ch.pick) series = series.filter((x, i) => ch.pick.includes(i));
    return { labels, series, autoTitle: series.length === 1 ? series[0].name : '' };
  };
  function hexA(c, a) { const n = c.replace('#', ''); const r = parseInt(n.slice(0, 2), 16), g = parseInt(n.slice(2, 4), 16), b = parseInt(n.slice(4, 6), 16); return `rgba(${r},${g},${b},${a})`; }
  Charts.config = (ch, forPrint) => {
    const d = Charts.extract(ch) || { labels: [], series: [] };
    const th = STheme.cur(); const style = ch.style || 1;
    let pal = ch.palette && ch.palette !== 'Theme' ? (PALETTES[ch.palette] || STheme.chartPalette()) : STheme.chartPalette();
    if (style === 6) pal = [0, .25, .45, .6, -.25, .7, -.4, .8].map(t => STheme.tint(th.accents[0], t));
    const t = ch.type; const pie = t === 'pie' || t === 'doughnut';
    const typeMap = { column: 'bar', bar: 'bar', line: 'line', area: 'line', pie: 'pie', doughnut: 'doughnut', scatter: 'scatter', combo: 'bar', stackedColumn: 'bar', stackedColumn100: 'bar', stackedBar: 'bar', stackedBar100: 'bar', radar: 'radar', bubble: 'bubble', waterfall: 'bar', histogram: 'bar' };
    const horiz = t === 'bar' || t === 'stackedBar' || t === 'stackedBar100'; const stacked = /^stacked/.test(t); const pct = /100$/.test(t);
    const dark = style === 4; const ink = dark ? '#E8ECF2' : '#303540', ink2 = dark ? '#B9C2D0' : '#595f6b', gridColor = dark ? 'rgba(255,255,255,.14)' : '#e3e6ea';
    const bg = dark ? STheme.tint(th.dk2, -0.2) : style === 5 ? STheme.tint(th.lt2, 0.4) : '#FFFFFF';
    const fontB = SU.cssFont(th.body), fontH = SU.cssFont(th.head);
    const outlined = style === 3; const bold = style === 2;
    const fillOf = (c, a) => outlined ? hexA(c, .25) : a !== undefined ? hexA(c, a) : c;
    let datasets, labels = d.labels;
    if (pie) { const s0 = d.series[0] || { name: '', data: [] }; datasets = [{ label: s0.name, data: s0.data, backgroundColor: d.labels.map((_, i) => fillOf(pal[i % pal.length])), borderColor: outlined ? d.labels.map((_, i) => pal[i % pal.length]) : bg, borderWidth: outlined ? 2 : 1.5 }]; }
    else if (t === 'waterfall') {
      const s0 = d.series[0] || { name: '', data: [] }; let run = 0; const data = [], colors = [], deltas = [];
      s0.data.forEach(v => { const x = typeof v === 'number' ? v : 0; data.push([run, run + x]); deltas.push(x); colors.push(x >= 0 ? pal[0] : pal[1]); run += x; });
      if (ch.total !== false) { data.push([0, run]); deltas.push(run); colors.push(dark ? '#9aa3b2' : '#8C8C8C'); labels = [...d.labels, 'Total']; }
      datasets = [{ label: s0.name, data, _deltas: deltas, backgroundColor: colors.map(c => fillOf(c)), borderColor: colors, borderWidth: outlined ? 2 : 0, borderSkipped: false, maxBarThickness: 64, barPercentage: .8 }];
    } else if (t === 'histogram') {
      const s0 = d.series[0] || { name: '', data: [] }; datasets = [{ label: s0.name, data: s0.data, backgroundColor: fillOf(pal[0]), borderColor: outlined ? pal[0] : bg, borderWidth: 1, barPercentage: 1, categoryPercentage: 1 }];
    } else {
      let series = d.series;
      if (pct) { const n = Math.max(0, ...series.map(x => x.data.length)); const tot = Array.from({ length: n }, (_, i) => series.reduce((a, x) => a + Math.abs(x.data[i] || 0), 0)); series = series.map(x => ({ ...x, data: x.data.map((v, i) => v === null ? null : tot[i] ? v / tot[i] * 100 : 0) })); }
      datasets = series.map((sr, i) => {
        const c = pal[i % pal.length]; const isLine = t === 'line' || t === 'area' || (t === 'combo' && i === series.length - 1 && series.length > 1) || t === 'radar';
        const ds = { label: sr.name, data: sr.data, backgroundColor: t === 'area' ? hexA(c, dark ? .45 : .35) : t === 'radar' ? hexA(c, .15) : fillOf(c), borderColor: c, borderWidth: isLine ? (bold ? 3 : 2.25) : outlined ? 2 : 0, fill: t === 'area' ? 'origin' : t === 'radar', tension: t === 'radar' ? 0 : 0.25, pointRadius: t === 'scatter' ? 4 : t === 'bubble' ? undefined : isLine ? (bold ? 4 : 3) : 0, pointBackgroundColor: c, borderRadius: style === 5 && !isLine ? 5 : (t === 'column' || t === 'bar' ? 2 : 0), showLine: false, maxBarThickness: 64 };
        if (t === 'bubble') { ds.backgroundColor = hexA(c, .55); ds.borderWidth = 1.5; }
        if (t === 'combo' && isLine) { ds.type = 'line'; ds.yAxisID = 'y1'; ds.order = 0; ds.showLine = true; ds.fill = false; } else if (t === 'combo') { ds.order = 1; }
        if (t === 'line' || t === 'area') ds.showLine = true;
        if (bold && !isLine) { ds.barPercentage = .92; ds.categoryPercentage = .85; }
        return ds;
      });
    }
    const title = ch.title !== undefined && ch.title !== null ? ch.title : (d.autoTitle || 'Chart Title');
    const at = ch.axisTitles; const xTitle = ch.xTitle ?? (d.xTitle || (horiz ? 'Values' : 'Categories')), yTitle = ch.yTitle ?? (horiz ? 'Categories' : (pct ? 'Percent' : t === 'histogram' ? 'Frequency' : d.series.length === 1 ? d.series[0].name : 'Values'));
    const tick = { color: ink2, font: { family: fontB } };
    const axTitle = (text) => ({ display: !!at && !!text, text, color: ink, font: { family: fontB, size: 12, weight: '600' } });
    let scales = {};
    if (t === 'radar') scales = { r: { grid: { color: gridColor }, angleLines: { color: gridColor }, pointLabels: { color: ink2, font: { family: fontB } }, ticks: { color: ink2, backdropColor: 'transparent', font: { family: fontB } }, beginAtZero: true } };
    else if (!pie) {
      const valAxis = { grid: { display: ch.grid !== false && style !== 2, color: gridColor }, ticks: { ...tick, ...(pct ? { callback: (v) => v + '%' } : {}) }, border: { display: false }, beginAtZero: t !== 'scatter' && t !== 'line' && t !== 'bubble' ? true : undefined, stacked: stacked || undefined, max: pct ? 100 : undefined };
      const catAxis = { ...(t === 'scatter' || t === 'bubble' ? { type: 'linear', grid: { display: ch.grid !== false, color: gridColor } } : { grid: { display: false } }), ticks: tick, border: { color: dark ? '#7d879b' : '#bfc4cc' }, stacked: stacked || undefined, offset: t === 'histogram' ? true : undefined };
      scales = horiz ? { x: { ...valAxis, title: axTitle(xTitle) }, y: { ...catAxis, title: axTitle(yTitle) } } : { x: { ...catAxis, title: axTitle(xTitle) }, y: { ...valAxis, title: axTitle(yTitle) } };
      if (t === 'combo' && datasets.length > 1) scales.y1 = { position: 'right', grid: { drawOnChartArea: false }, ticks: tick, border: { display: false }, beginAtZero: true, title: axTitle(datasets[datasets.length - 1].label) };
    }
    const cfg = {
      type: typeMap[t] || 'bar',
      data: { labels, datasets },
      options: {
        responsive: false, maintainAspectRatio: false, animation: false, devicePixelRatio: forPrint ? 3 : (window.devicePixelRatio || 1) * Math.max(1, Grid.z),
        indexAxis: horiz ? 'y' : 'x',
        layout: { padding: { top: 4, right: 10, bottom: 4, left: 4 } },
        plugins: {
          title: { display: !!title, text: title, font: { size: 16, weight: '600', family: fontH }, color: ink, padding: { top: 4, bottom: 8 } },
          legend: { display: ch.legend !== false && t !== 'waterfall' && t !== 'histogram' && (pie || d.series.length > 1 || ch.legend === true), position: pie ? 'right' : 'bottom', labels: { boxWidth: 12, boxHeight: 12, padding: 12, color: ink, font: { family: fontB } } },
          tooltip: { enabled: true, callbacks: t === 'waterfall' ? { label: (c) => ' ' + SU.generalNum(+(c.dataset._deltas[c.dataIndex]).toPrecision(10)) } : pct ? { label: (c) => ' ' + c.dataset.label + ': ' + (+c.raw).toFixed(1) + '%' } : undefined },
        },
        scales,
      },
      plugins: [bgPlugin(bg), labelsPlugin(ch, ink, fontB, bold)],
    };
    cfg._bg = bg;
    return cfg;
  };
  function bgPlugin(bg) { return { id: 'btsBg', beforeDraw(chart) { const { ctx, width, height } = chart; ctx.save(); ctx.fillStyle = bg; ctx.fillRect(0, 0, width, height); ctx.restore(); } }; }
  function labelsPlugin(ch, ink = '#303540', family = 'Carlito, Calibri, sans-serif', force) {
    return { id: 'btsLabels', afterDatasetsDraw(chart) { if (!ch.labels && !force) return; const { ctx } = chart; ctx.save(); ctx.font = '600 11px ' + family;
      const stacked = /^stacked/.test(ch.type), pct = /100$/.test(ch.type);
      chart.data.datasets.forEach((ds, di) => { const meta = chart.getDatasetMeta(di); if (meta.hidden) return; meta.data.forEach((el, i) => {
        const v = ds.data[i]; if (v === null || v === undefined) return;
        let val = Array.isArray(v) ? (ds._deltas ? ds._deltas[i] : v[1] - v[0]) : typeof v === 'object' ? v.y : v; if (ch.type === 'histogram' && !val) return;
        const txt = pct ? (+val).toFixed(0) + '%' : SU.generalNum(+(+val).toPrecision(6));
        let { x, y } = el.tooltipPosition ? el.tooltipPosition() : el; ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
        const type = meta.type || chart.config.type;
        if (type === 'pie' || type === 'doughnut') { ctx.fillStyle = '#fff'; ctx.textBaseline = 'middle'; }
        else if (stacked && type === 'bar') { ctx.fillStyle = '#fff'; ctx.textBaseline = 'middle'; if (chart.options.indexAxis === 'y') x = (el.x + el.base) / 2; else y = (el.y + el.base) / 2; }
        else if (chart.options.indexAxis === 'y') { ctx.fillStyle = ink; ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; x += 4; }
        else { ctx.fillStyle = ink; y -= 3; }
        ctx.fillText(txt, x, y); }); }); ctx.restore(); } };
  }

  /* ---------- DOM objects ---------- */
  Charts.layer = () => Grid.objLayer;
  Charts.renderSheet = () => {
    const s = S(); const layer = Grid.objLayer; if (!layer) return;
    const want = new Set([...s.charts.map(c => c.id), ...s.images.map(i => i.id)]);
    [...layer.children].forEach(el => { if (!want.has(el.dataset.id)) { const inst = Charts.instances.get(el.dataset.id); if (inst) { inst.destroy(); Charts.instances.delete(el.dataset.id); } el.remove(); } });
    for (const ch of s.charts) ensureEl(ch, 'chart');
    for (const im of s.images) ensureEl(im, 'img');
    Charts.position();
  };
  function ensureEl(obj, kind) {
    const layer = Grid.objLayer; let el = layer.querySelector(`[data-id="${obj.id}"]`);
    if (!el) {
      el = document.createElement('div'); el.className = 'fobj ' + kind; el.dataset.id = obj.id;
      el.innerHTML = (kind === 'chart' ? '<canvas></canvas>' : `<img alt="" draggable="false">`) + ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].map(h => `<div class="h ${h}" data-h="${h}"></div>`).join('') + `<div class="obj-tools">${kind === 'chart' ? `<button data-act="edit" title="Chart options">${icon('settings-2', 15)}</button>` : ''}<button data-act="del" title="Delete">${icon('trash-2', 15)}</button></div>`;
      layer.appendChild(el);
      el.addEventListener('mousedown', (e) => onObjDown(e, obj.id, kind));
      el.addEventListener('dblclick', (e) => { if (kind === 'chart') { e.preventDefault(); Charts.editTitle(obj.id); } });
      el.addEventListener('contextmenu', (e) => { e.preventDefault(); e.stopPropagation(); Charts.select(obj.id); objMenu(e, kind); });
    }
    if (kind === 'img') { const img = el.querySelector('img'); if (img.getAttribute('src') !== obj.src) img.src = obj.src; }
    el.classList.toggle('sel', Charts.selected === obj.id);
    return el;
  }
  function objMenu(e, kind) {
    BTS.menu({ x: e.clientX, y: e.clientY }, [
      ...(kind === 'chart' ? [{ label: 'Edit chart title', icon: 'heading', action: () => Charts.editTitle(Charts.selected) }, { label: 'Change chart type', icon: 'chart-column-big', action: () => App.cmd('chart-type') }, { label: 'Select data…', icon: 'table-2', action: () => App.cmd('chart-data') }, { sep: true }] : []),
      { label: 'Bring to front', icon: 'bring-to-front', action: () => Charts.bringFront() },
      { label: 'Copy as picture', icon: 'copy', action: () => Charts.copyPicture() },
      { sep: true }, { label: 'Delete', icon: 'trash-2', kbd: 'Del', action: () => Charts.deleteSelected() },
    ]);
  }
  Charts.find = (id) => { for (const s of wb.sheets) { const c = s.charts.find(x => x.id === id); if (c) return { s, obj: c, kind: 'chart' }; const i = s.images.find(x => x.id === id); if (i) return { s, obj: i, kind: 'img' }; } return null; };
  Charts.position = () => {
    const s = S(); const layer = Grid.objLayer; if (!layer) return; const z = Grid.z;
    const hw = Grid.hw(), hh = Grid.hh(); const sl = Grid.scroller.scrollLeft, st = Grid.scroller.scrollTop;
    const f = Grid.frozen();
    layer.style.clipPath = `inset(${hh}px 0 0 ${hw}px)`;
    for (const obj of [...s.charts, ...s.images]) {
      const el = layer.querySelector(`[data-id="${obj.id}"]`); if (!el) continue;
      const x = hw + obj.x * z - (obj.x * z >= f.fw ? sl : 0), y = hh + obj.y * z - (obj.y * z >= f.fh ? st : 0);
      el.style.left = x + 'px'; el.style.top = y + 'px'; el.style.width = obj.w * z + 'px'; el.style.height = obj.h * z + 'px';
      el.style.zIndex = obj.z || 1;
      if (el.classList.contains('chart')) drawChart(obj, el);
    }
  };
  function drawChart(ch, el) {
    const z = Grid.z; const w = Math.max(40, Math.round(ch.w * z - 16)), h = Math.max(40, Math.round(ch.h * z - 16));
    const key = JSON.stringify([w, h, wb.version, ch.type, ch.title, ch.legend, ch.labels, ch.grid, ch.palette, ch.switchRC, ch.range, ch.sheet, z, ch.style, ch.axisTitles, ch.xTitle, ch.yTitle, ch.pivot, wb.theme, Charts.fontGen]);
    if (el._key === key) return; el._key = key;
    const canvas = el.querySelector('canvas'); ch._sheet = Charts.find(ch.id)?.s;
    let inst = Charts.instances.get(ch.id);
    const cfg = Charts.config(ch);
    cfg.options.plugins.title.font.size = Math.round(15 * Math.max(.6, Math.min(z, 2)));
    el.style.background = cfg._bg;
    const th = STheme.cur(); STheme.fontReady(th.body, false, false); STheme.fontReady(th.head, true, false);
    if (inst && inst.config.type === cfg.type) {
      inst.data = cfg.data; inst.options = cfg.options; inst.config.plugins = cfg.plugins;
      inst.resize(w, h); inst.update('none');
    } else {
      if (inst) inst.destroy();
      canvas.style.width = w + 'px'; canvas.style.height = h + 'px'; canvas.width = w; canvas.height = h;
      inst = new Chart(canvas, cfg); Charts.instances.set(ch.id, inst);
    }
  }
  Charts.refreshAll = () => { Charts.renderSheet(); };
  Charts.fontGen = 0;
  Charts.onFonts = () => { Charts.fontGen++; Charts.position(); };

  /* ---------- interaction ---------- */
  function onObjDown(e, id, kind) {
    if (e.button !== 0) return;
    if (e.target.closest('.obj-tools')) { const act = e.target.closest('[data-act]'); if (act) { e.preventDefault(); e.stopPropagation(); if (act.dataset.act === 'del') Charts.deleteSelected(); else App.cmd('chart-type'); } return; }
    e.preventDefault(); e.stopPropagation();
    if (Ed.isEditing()) Ed.commit();
    Charts.select(id);
    const f = Charts.find(id); if (!f) return; const obj = f.obj; const z = Grid.z;
    const handle = e.target.dataset.h; const sx = e.clientX, sy = e.clientY; const o = { x: obj.x, y: obj.y, w: obj.w, h: obj.h };
    let moved = false;
    const mv = (ev) => {
      const dx = (ev.clientX - sx) / z, dy = (ev.clientY - sy) / z; if (Math.abs(dx) + Math.abs(dy) > 2) moved = true; if (!moved) return;
      let { x, y, w, h } = o;
      if (!handle) { x = Math.max(0, o.x + dx); y = Math.max(0, o.y + dy); }
      else {
        if (handle.includes('e')) w = Math.max(30, o.w + dx); if (handle.includes('s')) h = Math.max(30, o.h + dy);
        if (handle.includes('w')) { w = Math.max(30, o.w - dx); x = o.x + o.w - w; } if (handle.includes('n')) { h = Math.max(30, o.h - dy); y = o.y + o.h - h; }
        if (kind === 'img' && ev.shiftKey === false && handle.length === 2) { const ar = o.w / o.h; if (w / h > ar) w = h * ar; else h = w / ar; if (handle.includes('w')) x = o.x + o.w - w; if (handle.includes('n')) y = o.y + o.h - h; }
      }
      Object.assign(obj, { x, y, w, h }); Charts.position();
    };
    const up = () => {
      document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up);
      if (moved) { const after = { x: obj.x, y: obj.y, w: obj.w, h: obj.h }; Object.assign(obj, o); wb.run(handle ? 'Resize' : 'Move', () => wb.custom(() => Object.assign(obj, after), () => Object.assign(obj, o))); Charts.position(); }
    };
    document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
  }
  Charts.select = (id) => {
    Charts.selected = id; const f = id && Charts.find(id);
    [...Grid.objLayer.children].forEach(el => el.classList.toggle('sel', el.dataset.id === id));
    const isChart = f && f.kind === 'chart';
    document.getElementById('tab-chart').classList.toggle('hidden', !isChart);
    if (isChart) { BTS.showTab('chart'); document.getElementById('chk-legend').checked = f.obj.legend !== false; document.getElementById('chk-labels').checked = !!f.obj.labels; document.getElementById('chk-cgrid').checked = f.obj.grid !== false; }
    else if (BTS.activeTab === 'chart') BTS.showTab('insert');
    Grid.hideSelection = !!id; Grid.render();
    if (id) document.getElementById('gkeys').focus({ preventScroll: true });
    if (isChart) { document.getElementById('chk-axist').checked = !!f.obj.axisTitles; }
    App.refreshContext && App.refreshContext();
  };
  Charts.deleteSelected = () => {
    const f = Charts.find(Charts.selected); if (!f) return; const { s, obj, kind } = f; const prop = kind === 'chart' ? 'charts' : 'images';
    wb.run('Delete object', () => wb.setProp(s, prop, s[prop].filter(x => x !== obj))); Charts.select(null); Charts.renderSheet();
  };
  Charts.bringFront = () => { const f = Charts.find(Charts.selected); if (!f) return; const mz = Math.max(1, ...[...f.s.charts, ...f.s.images].map(o => o.z || 1)); const old = f.obj.z; wb.run('Bring to front', () => wb.custom(() => { f.obj.z = mz + 1; }, () => { f.obj.z = old; })); Charts.position(); };
  Charts.copyPicture = async () => {
    const f = Charts.find(Charts.selected); if (!f) return;
    let blob; if (f.kind === 'chart') { const inst = Charts.instances.get(f.obj.id); blob = await new Promise(r => inst.canvas.toBlob(r)); } else blob = new Blob([BTS.dataURLToBytes(f.obj.src)], { type: 'image/png' });
    try { await navigator.clipboard.write([new ClipboardItem({ [blob.type]: blob })]); BTS.toast('Copied'); } catch { BTS.toast('Couldn’t copy the picture.', { error: true }); }
  };
  Charts.update = (patch, label = 'Chart') => {
    const f = Charts.find(Charts.selected); if (!f || f.kind !== 'chart') return; const obj = f.obj; const before = {}; Object.keys(patch).forEach(k => before[k] = obj[k]);
    wb.run(label, () => wb.custom(() => Object.assign(obj, patch), () => Object.assign(obj, before)));
    Charts.position();
  };
  Charts.editTitle = async (id) => {
    if (id) Charts.select(id); const f = Charts.find(Charts.selected); if (!f || f.kind !== 'chart') return;
    const cur = f.obj.title !== undefined && f.obj.title !== null ? f.obj.title : (Charts.extract(f.obj)?.autoTitle || 'Chart Title');
    const t = await BTS.prompt('Chart title', 'Title (leave empty for no title)', cur);
    if (t === null) return; Charts.update({ title: t }, 'Chart title');
  };

  /* ---------- create ---------- */
  Charts.create = (type) => {
    const s = S(); let g = A.dataRange();
    if (s.sel.r1 > s.sel.r0 || s.sel.c1 > s.sel.c0) g = A.effData(s.sel);
    if (g.r0 === g.r1 && g.c0 === g.c1 && s.isEmpty(g.r0, g.c0)) { BTS.toast('Select the data you want to chart first.', { error: true }); return; }
    let any = false; for (let r = g.r0; r <= g.r1 && !any; r++) for (let c = g.c0; c <= g.c1; c++) if (typeof s.value(r, c) === 'number') { any = true; break; }
    if (!any) { BTS.toast('The selected range doesn’t contain numbers to chart.', { error: true }); return; }
    const x = s.colLeft(g.c1 + 1) + 16, y = Math.max(0, s.rowTop(g.r0));
    const ch = { id: 'ch' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5), type, range: g, sheet: null, x, y, w: 480, h: 290, legend: undefined, z: 1 };
    // keep chart visible: place near the top-left of the viewport if data is off-screen
    wb.run('Insert chart', () => wb.setProp(s, 'charts', [...s.charts, ch]));
    Charts.renderSheet(); Charts.select(ch.id);
    return ch;
  };
  Charts.insertImage = async () => {
    const f = await BTS.pickFile('image/*'); if (!f) return;
    Charts.insertImageBytes(f.data, f.type || ('image/' + (BTS.extOf(f.name) || 'png').replace('jpg', 'jpeg')));
  };
  Charts.insertImageBytes = async (bytes, mime) => {
    const src = await BTS.bytesToDataURL(bytes, mime || 'image/png');
    const img = new Image(); await new Promise((res) => { img.onload = res; img.onerror = res; img.src = src; });
    let w = img.naturalWidth || 300, h = img.naturalHeight || 200; const max = 520; if (w > max) { h = h * max / w; w = max; } if (h > 400) { w = w * 400 / h; h = 400; }
    const s = S(); const a = s.active;
    const im = { id: 'im' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5), src, x: s.colLeft(a.c), y: s.rowTop(a.r), w, h, z: 1 };
    wb.run('Insert picture', () => wb.setProp(s, 'images', [...s.images, im]));
    Charts.renderSheet(); Charts.select(im.id);
  };
  /** PNG data URL of a chart for printing/export */
  Charts.toImage = (ch, s, scale = 2) => {
    const c = document.createElement('canvas'); const w = Math.round(ch.w - 16), h = Math.round(ch.h - 16);
    c.width = w; c.height = h; c.style.width = w + 'px'; c.style.height = h + 'px';
    ch._sheet = s; const cfg = Charts.config(ch, true); cfg.options.devicePixelRatio = scale; cfg.options.responsive = false;
    const holder = document.createElement('div'); holder.style.cssText = `position:fixed;left:-5000px;top:0;width:${w}px;height:${h}px`; holder.appendChild(c); document.body.appendChild(holder);
    const inst = new Chart(c, cfg); const url = c.toDataURL('image/png'); inst.destroy(); holder.remove();
    return url;
  };
})();
