/* BTS Sheets — floating objects: Chart.js charts (live-updating) and pictures; drag, resize, select, delete */
(function () {
  const Charts = window.Charts = { selected: null, instances: new Map() };
  let wb;
  const PALETTES = {
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
  Charts.extract = (ch) => {
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
    const pal = PALETTES[ch.palette || 'BTS'] || PALETTES.BTS;
    const t = ch.type; const pie = t === 'pie' || t === 'doughnut';
    const typeMap = { column: 'bar', bar: 'bar', line: 'line', area: 'line', pie: 'pie', doughnut: 'doughnut', scatter: 'scatter' };
    let datasets;
    if (pie) { const s0 = d.series[0] || { name: '', data: [] }; datasets = [{ label: s0.name, data: s0.data, backgroundColor: d.labels.map((_, i) => pal[i % pal.length]), borderColor: '#fff', borderWidth: 1.5 }]; }
    else datasets = d.series.map((s, i) => ({ label: s.name, data: s.data, backgroundColor: t === 'area' ? hexA(pal[i % pal.length], .35) : pal[i % pal.length], borderColor: pal[i % pal.length], borderWidth: t === 'line' || t === 'area' ? 2.25 : 0, fill: t === 'area' ? 'origin' : false, tension: 0.25, pointRadius: t === 'scatter' ? 4 : t === 'line' ? 3 : 0, pointBackgroundColor: pal[i % pal.length], borderRadius: t === 'column' || t === 'bar' ? 2 : 0, showLine: false, maxBarThickness: 64 }));
    const title = ch.title !== undefined && ch.title !== null ? ch.title : (d.autoTitle || 'Chart Title');
    const gridColor = '#e3e6ea';
    const cfg = {
      type: typeMap[t] || 'bar',
      data: { labels: d.labels, datasets },
      options: {
        responsive: false, maintainAspectRatio: false, animation: false, devicePixelRatio: forPrint ? 3 : (window.devicePixelRatio || 1) * Math.max(1, Grid.z),
        indexAxis: t === 'bar' ? 'y' : 'x',
        layout: { padding: { top: 4, right: 10, bottom: 4, left: 4 } },
        plugins: {
          title: { display: !!title, text: title, font: { size: 16, weight: '600' }, color: '#303540', padding: { top: 4, bottom: 8 } },
          legend: { display: ch.legend !== false && (pie || d.series.length > 1 || ch.legend === true), position: pie ? 'right' : 'bottom', labels: { boxWidth: 12, boxHeight: 12, padding: 12 } },
          tooltip: { enabled: true },
        },
        scales: pie ? {} : {
          x: { ...(t === 'scatter' ? { type: 'linear' } : {}), grid: { display: t === 'bar' ? ch.grid !== false : false, color: gridColor }, ticks: { color: '#595f6b' }, border: { color: '#bfc4cc' }, beginAtZero: t === 'bar' },
          y: { grid: { display: t !== 'bar' && ch.grid !== false, color: gridColor }, ticks: { color: '#595f6b' }, border: { display: false }, beginAtZero: t !== 'scatter' && t !== 'line' ? true : undefined },
        },
      },
      plugins: [labelsPlugin(ch)],
    };
    return cfg;
  };
  function labelsPlugin(ch) {
    return { id: 'btsLabels', afterDatasetsDraw(chart) { if (!ch.labels) return; const { ctx } = chart; ctx.save(); ctx.font = '600 11px Carlito, Calibri, sans-serif'; ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
      chart.data.datasets.forEach((ds, di) => { const meta = chart.getDatasetMeta(di); if (meta.hidden) return; meta.data.forEach((el, i) => { const v = ds.data[i]; if (v === null || v === undefined) return; const val = typeof v === 'object' ? v.y : v; const txt = SU.generalNum(+(+val).toPrecision(6)); let { x, y } = el.tooltipPosition ? el.tooltipPosition() : el; if (chart.config.type === 'pie' || chart.config.type === 'doughnut') { ctx.fillStyle = '#fff'; ctx.textBaseline = 'middle'; } else if (chart.options.indexAxis === 'y') { ctx.fillStyle = '#303540'; ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; x += 4; } else { ctx.fillStyle = '#303540'; y -= 3; } ctx.fillText(txt, x, y); }); }); ctx.restore(); } };
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
    const key = JSON.stringify([w, h, wb.version, ch.type, ch.title, ch.legend, ch.labels, ch.grid, ch.palette, ch.switchRC, ch.range, ch.sheet, z]);
    if (el._key === key) return; el._key = key;
    const canvas = el.querySelector('canvas'); ch._sheet = Charts.find(ch.id)?.s;
    let inst = Charts.instances.get(ch.id);
    const cfg = Charts.config(ch);
    cfg.options.plugins.title.font.size = Math.round(15 * Math.max(.6, Math.min(z, 2)));
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
