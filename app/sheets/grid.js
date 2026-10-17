/* BTS Sheets — grid: canvas rendering with virtualization, frozen panes, selection, mouse interaction, floating objects layer */
(function () {
  const { key } = SU;
  const G = window.Grid = {
    z: 1, showFormulas: false, refHL: [], copyRange: null, fillPreview: null, dragging: null, headerH: 22,
  };
  let wb, host, scroller, spacer, canvas, ctx, objLayer, W = 0, H = 0, dpr = 1, raf = 0;
  const REF_COLORS = ['#2f6fde', '#d6336c', '#7c3aed', '#0d9488', '#ea580c', '#16a34a', '#b45309', '#0891b2'];
  G.REF_COLORS = REF_COLORS;
  let theme = {};
  function readTheme() {
    const cs = getComputedStyle(document.body); const dark = document.documentElement.dataset.theme === 'dark';
    theme = {
      dark, accent: cs.getPropertyValue('--accent').trim() || '#0F8A6A',
      hdrBg: dark ? '#202736' : '#f4f6f8', hdrText: dark ? '#aab3c5' : '#5c6577', hdrLine: dark ? '#323b4d' : '#d9dde3',
      hdrSelBg: dark ? '#1d4a40' : '#dcefe8', hdrSelText: dark ? '#8ef0cf' : '#0b6b52', hdrFullBg: dark ? '#0F8A6A' : '#0F8A6A',
      grid: '#e1e4e8', cellBg: '#ffffff', selFill: 'rgba(15,138,106,0.10)',
    };
  }
  G.readTheme = () => { readTheme(); G.render(); };

  G.init = (workbook, el) => {
    wb = workbook; host = el;
    host.innerHTML = `<canvas id="gcanvas"></canvas><div id="gscroll" tabindex="-1"><div id="gspace"></div></div><div id="gobjs"></div>`;
    canvas = host.querySelector('#gcanvas'); ctx = canvas.getContext('2d', { alpha: false });
    scroller = host.querySelector('#gscroll'); spacer = host.querySelector('#gspace'); objLayer = host.querySelector('#gobjs');
    G.scroller = scroller; G.objLayer = objLayer; G.canvas = canvas;
    readTheme();
    new ResizeObserver(() => { resize(); }).observe(host);
    scroller.addEventListener('scroll', onScroll, { passive: true });
    scroller.addEventListener('mousedown', onMouseDown);
    scroller.addEventListener('mousemove', onHover);
    scroller.addEventListener('mouseleave', () => G.onHoverCell && G.onHoverCell(null));
    scroller.addEventListener('dblclick', onDblClick);
    scroller.addEventListener('contextmenu', onContext);
    scroller.addEventListener('wheel', (e) => { if (e.ctrlKey) { e.preventDefault(); BTS.setZoom(BTS.zoom + (e.deltaY < 0 ? 10 : -10)); } }, { passive: false });
    new MutationObserver(() => readTheme() || G.render()).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });
    resize();
  };
  const S = () => wb.active;
  G.sheet = S;

  function resize() {
    const r = host.getBoundingClientRect(); dpr = window.devicePixelRatio || 1;
    W = Math.max(1, Math.floor(r.width)); H = Math.max(1, Math.floor(r.height));
    canvas.width = Math.round(W * dpr); canvas.height = Math.round(H * dpr); canvas.style.width = W + 'px'; canvas.style.height = H + 'px';
    updateSpacer(); G.renderNow();
  }
  G.resize = resize;

  /* ---------- geometry ---------- */
  G.hw = () => { const s = S(); const z = G.z; if (!s.showHeadings) return 0; const lastRow = s.rowAt((scroller.scrollTop + H) / z) + 1; ctx.font = `${Math.round(11 * Math.max(.8, z))}px ${getFontUI()}`; return G.ogW() + (Math.max(34, Math.ceil(ctx.measureText(String(lastRow)).width + 16)) * Math.max(.75, Math.min(z, 1.6)) | 0); };
  G.hh = () => S().showHeadings ? G.ogH() + Math.round(22 * Math.max(.75, Math.min(G.z, 1.6))) : 0;
  /* ---------- outline (group) gutters ---------- */
  const OG = 16;
  function outlineInfo(map) {
    if (map._info) return map._info;
    const keys = [...map.keys()].sort((a, b) => a - b); let max = 0; for (const k of keys) max = Math.max(max, map.get(k));
    const groups = [];
    for (let L = 1; L <= max; L++) { let a = null, prev = null; for (const k of keys) { if (map.get(k) < L) continue; if (a === null) { a = k; prev = k; continue; } if (k === prev + 1) { prev = k; continue; } groups.push({ L, a, b: prev }); a = k; prev = k; } if (a !== null) groups.push({ L, a, b: prev }); }
    const info = { max, groups }; try { Object.defineProperty(map, '_info', { value: info, configurable: true }); } catch { } return info;
  }
  G.outlineInfo = outlineInfo;
  G.ogW = () => { const s = S(); if (!s || !s.showHeadings || !s.outlineR || !s.outlineR.size) return 0; return (outlineInfo(s.outlineR).max + 1) * OG + 4; };
  G.ogH = () => { const s = S(); if (!s || !s.showHeadings || !s.outlineC || !s.outlineC.size) return 0; return (outlineInfo(s.outlineC).max + 1) * OG + 4; };
  let fontUI; function getFontUI() { return fontUI || (fontUI = getComputedStyle(document.body).fontFamily); }
  function frozen() { const s = S(), z = G.z; return { fr: s.freeze.r, fc: s.freeze.c, fw: s.colLeft(s.freeze.c) * z, fh: s.rowTop(s.freeze.r) * z }; }
  G.frozen = frozen;
  /** screen x (relative to host) of sheet column left edge */
  G.colX = (c) => { const s = S(), z = G.z, hw = curHW; return c < s.freeze.c ? hw + s.colLeft(c) * z : hw + s.colLeft(c) * z - scroller.scrollLeft; };
  G.rowY = (r) => { const s = S(), z = G.z, hh = curHH; return r < s.freeze.r ? hh + s.rowTop(r) * z : hh + s.rowTop(r) * z - scroller.scrollTop; };
  G.colAtX = (x) => { const s = S(), z = G.z, f = frozen(); const lx = x - curHW; if (lx < f.fw) return s.colAt(Math.max(0, lx) / z); return s.colAt((lx + scroller.scrollLeft) / z); };
  G.rowAtY = (y) => { const s = S(), z = G.z, f = frozen(); const ly = y - curHH; if (ly < f.fh) return s.rowAt(Math.max(0, ly) / z); return s.rowAt((ly + scroller.scrollTop) / z); };
  G.cellRect = (r, c, noMerge) => {
    const s = S(), z = G.z; const m = noMerge ? null : s.mergeAt(r, c);
    if (m) { const x = G.colX(m.c0), y = G.rowY(m.r0); return { x, y, w: (s.colLeft(m.c1 + 1) - s.colLeft(m.c0)) * z, h: (s.rowTop(m.r1 + 1) - s.rowTop(m.r0)) * z, m }; }
    return { x: G.colX(c), y: G.rowY(r), w: s.colWidth(c) * z, h: s.rowHeight(r) * z };
  };
  G.rangeRect = (g) => { const s = S(), z = G.z; const x = G.colX(g.c0), y = G.rowY(g.r0); return { x, y, w: (s.colLeft(g.c1 + 1) - s.colLeft(g.c0)) * z, h: (s.rowTop(g.r1 + 1) - s.rowTop(g.r0)) * z }; };
  let curHW = 40, curHH = 22;

  function updateSpacer() {
    const s = S(); if (!s) return; const z = G.z;
    curHW = G.hw(); curHH = G.hh();
    spacer.style.width = Math.ceil(curHW + s.totalW() * z) + 'px';
    spacer.style.height = Math.ceil(curHH + s.totalH() * z) + 'px';
  }
  G.updateSpacer = updateSpacer;
  function onScroll() {
    const s = S(); const z = G.z;
    // grow on demand
    if (scroller.scrollTop + scroller.clientHeight > curHH + s.totalH() * z - 300 && s.rowCount < SU.MAXR) { s.rowCount = Math.min(SU.MAXR, s.rowCount + 1000); s.invalidateGeom(); updateSpacer(); }
    if (scroller.scrollLeft + scroller.clientWidth > curHW + s.totalW() * z - 200 && s.colCount < SU.MAXC) { s.colCount = Math.min(SU.MAXC, s.colCount + 26); s.invalidateGeom(); updateSpacer(); }
    s.scroll = { x: scroller.scrollLeft, y: scroller.scrollTop };
    G.renderNow();
    G.onScroll && G.onScroll();
  }
  G.render = () => { if (!raf) raf = requestAnimationFrame(() => { raf = 0; G.renderNow(); }); };
  G.setZoom = (z) => {
    const s = S(); const ar = s.active; G.z = z / 100; updateSpacer();
    G.renderNow(); G.onScroll && G.onScroll();
  };
  G.restoreScroll = () => { const s = S(); s.sel = s.expandMerges(s.sel); const am = s.mergeAt(s.active.r, s.active.c); if (am) s.active = { r: am.r0, c: am.c0 }; updateSpacer(); scroller.scrollLeft = s.scroll.x; scroller.scrollTop = s.scroll.y; G.renderNow(); };

  G.scrollIntoView = (r, c) => {
    const s = S(), z = G.z, f = frozen();
    s.ensureSize(r, c); updateSpacer();
    const viewW = scroller.clientWidth - curHW - f.fw, viewH = scroller.clientHeight - curHH - f.fh;
    if (c >= s.freeze.c && c !== undefined) {
      const x0 = s.colLeft(c) * z - f.fw, x1 = s.colLeft(c + 1) * z - f.fw;
      if (x0 < scroller.scrollLeft) scroller.scrollLeft = x0; else if (x1 > scroller.scrollLeft + viewW) scroller.scrollLeft = Math.min(x0, x1 - viewW);
    }
    if (r >= s.freeze.r && r !== undefined) {
      const y0 = s.rowTop(r) * z - f.fh, y1 = s.rowTop(r + 1) * z - f.fh;
      if (y0 < scroller.scrollTop) scroller.scrollTop = y0; else if (y1 > scroller.scrollTop + viewH) scroller.scrollTop = Math.min(y0, y1 - viewH);
    }
    G.render();
  };

  /* ---------- rendering ---------- */
  const al = (v) => Math.round(v * dpr) / dpr;              // align to device pixel
  const ln = (v) => (Math.round(v * dpr) + 0.5) / dpr;        // 1 device px line center
  function fontStr(st, z, num) { const name = st.font || (num && SU.numFont) || SU.defFont; STheme.fontReady(name, st.b, st.i); return `${st.i ? 'italic ' : ''}${st.b ? 'bold ' : ''}${((st.sz || SU.defSize) * 96 / 72 * z).toFixed(2)}px ${SU.cssFont(name)}`; }
  STheme.onFontsLoaded = () => { measureCache.clear(); G.render(); G.onFontsLoaded && G.onFontsLoaded(); };
  G.fontStr = fontStr;

  G.renderNow = () => {
    const s = S(); if (!s || !ctx) return;
    const z = G.z; curHW = G.hw(); curHH = G.hh(); const hw = curHW, hh = curHH; const f = frozen();
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.fillStyle = theme.cellBg; ctx.fillRect(0, 0, W, H);
    const sl = scroller.scrollLeft, st = scroller.scrollTop;
    // quadrant definitions
    const quads = [];
    const colRange = (fromX, toX) => [s.colAt(fromX / z), s.colAt(toX / z)];
    const rowRange = (fromY, toY) => [s.rowAt(fromY / z), s.rowAt(toY / z)];
    const mainCols = colRange(f.fw + sl, W - hw + sl), mainRows = rowRange(f.fh + st, H - hh + st);
    const frCols = f.fc ? [0, f.fc - 1] : null, frRows = f.fr ? [0, f.fr - 1] : null;
    quads.push({ rows: mainRows, cols: mainCols, clip: [hw + f.fw, hh + f.fh, W - hw - f.fw, H - hh - f.fh] });
    if (frRows) quads.push({ rows: frRows, cols: mainCols, clip: [hw + f.fw, hh, W - hw - f.fw, f.fh] });
    if (frCols) quads.push({ rows: mainRows, cols: frCols, clip: [hw, hh + f.fh, f.fw, H - hh - f.fh] });
    if (frRows && frCols) quads.push({ rows: frRows, cols: frCols, clip: [hw, hh, f.fw, f.fh] });
    G.visible = { r0: mainRows[0], r1: mainRows[1], c0: mainCols[0], c1: mainCols[1] };
    for (const q of quads) { ctx.save(); ctx.beginPath(); ctx.rect(...q.clip); ctx.clip(); drawCells(s, q.rows[0], q.rows[1], q.cols[0], q.cols[1]); drawSelection(s); ctx.restore(); }
    drawHeaders(s, hw, hh, f, mainCols, mainRows, frCols, frRows);
    if (f.fr || f.fc) { ctx.strokeStyle = '#9aa3b2'; ctx.lineWidth = 1.5; ctx.beginPath(); if (f.fr) { ctx.moveTo(0, al(hh + f.fh)); ctx.lineTo(W, al(hh + f.fh)); } if (f.fc) { ctx.moveTo(al(hw + f.fw), 0); ctx.lineTo(al(hw + f.fw), H); } ctx.stroke(); }
    G.positionObjects && G.positionObjects();
    G.afterRender && G.afterRender();
  };

  function cellValueForDisplay(s, r, c) {
    if (G.showFormulas && s.isFormula(r, c)) return { text: s.raw(r, c), formula: true };
    return null;
  }

  let renderTheme = null;
  function tableStyleAt(s, r, c) {
    for (const t of s.tables) { if (SU.inRange(t.range, r, c)) return STheme.tableCell(t, r, c, renderTheme); }
    return null;
  }
  G.tableStyleAt = (s, r, c) => tableStyleAt(s, r, c);
  /* sparkline lookup per sheet (rebuilt when the sparklines array changes) */
  function sparkMap(s) {
    if (s._spk && s._spk.src === s.sparklines) return s._spk.map;
    const map = new Map(); for (const g of s.sparklines) for (const it of g.items) map.set(key(it.r, it.c), { g, it });
    s._spk = { src: s.sparklines, map }; return map;
  }
  G.sparkAt = (s, r, c) => s.sparklines.length ? sparkMap(s).get(key(r, c)) || null : null;

  function drawCells(s, r0, r1, c0, c1) {
    const z = G.z;
    // extend columns a bit for text overflow sources
    const cExt0 = Math.max(0, c0 - 8), cExt1 = Math.min(s.colCount - 1, c1 + 8);
    const merges = s.merges.filter(m => m.r1 >= r0 && m.r0 <= r1 && m.c1 >= c0 && m.c0 <= c1);
    const inMerge = new Set(); for (const m of merges) for (let r = Math.max(m.r0, r0); r <= Math.min(m.r1, r1); r++) for (let c = m.c0; c <= m.c1; c++) inMerge.add(key(r, c));
    const ver = wb.version; renderTheme = STheme.cur();
    const texts = []; const spk = s.sparklines.length ? sparkMap(s) : null;
    // pass 1: fills
    for (let r = r0; r <= r1; r++) {
      const rh = s.rowHeight(r) * z; if (!rh) continue; const y = G.rowY(r);
      for (let c = c0; c <= c1; c++) {
        const cw = s.colWidth(c) * z; if (!cw) continue;
        if (inMerge.has(key(r, c))) continue;
        const st = s.style(r, c); const cf = s.cf.length ? SU.evalCF(s, r, c, ver) : null; const tb = s.tables.length ? tableStyleAt(s, r, c) : null;
        const fill = (cf && cf.fill) || st.fill || (tb && tb.fill);
        if (fill) { ctx.fillStyle = fill; ctx.fillRect(al(G.colX(c)), al(y), al(cw) + 0.5, al(rh) + 0.5); }
        if (cf && cf.bar) drawBar(G.colX(c), y, cw, rh, cf.bar);
        if (spk) { const sp = spk.get(key(r, c)); if (sp && window.Spark) Spark.draw(ctx, s, sp.g, sp.it, G.colX(c), y, cw, rh, dpr); }
      }
    }
    // pass 2: gridlines
    if (s.showGrid) {
      ctx.strokeStyle = theme.grid; ctx.lineWidth = 1 / dpr; ctx.beginPath();
      const xa = G.colX(c0), xb = G.colX(c1 + 1), ya = G.rowY(r0), yb = G.rowY(r1 + 1);
      for (let r = r0; r <= r1 + 1; r++) { if (r <= r1 && !s.rowHeight(r) && r > r0) continue; const y = ln(G.rowY(r) - 1 / dpr); ctx.moveTo(xa, y); ctx.lineTo(xb, y); }
      for (let c = c0; c <= c1 + 1; c++) { const x = ln(G.colX(c) - 1 / dpr); ctx.moveTo(x, ya); ctx.lineTo(x, yb); }
      ctx.stroke();
    }
    // pass 3: text (with overflow)
    for (let r = r0; r <= r1; r++) {
      const rh = s.rowHeight(r) * z; if (!rh) continue; const y = G.rowY(r);
      for (let c = cExt0; c <= cExt1; c++) {
        const cw = s.colWidth(c) * z; if (!cw) continue;
        if (inMerge.has(key(r, c)) || (s.merges.length && s.mergeAt(r, c))) continue;
        if (s.isEmpty(r, c)) continue;
        const inView = c >= c0 && c <= c1;
        drawCellText(s, r, c, G.colX(c), y, cw, rh, inView, c0, c1, ver);
      }
    }
    // merges
    for (const m of merges) {
      const x = G.colX(m.c0), y = G.rowY(m.r0), w = (s.colLeft(m.c1 + 1) - s.colLeft(m.c0)) * z, h = (s.rowTop(m.r1 + 1) - s.rowTop(m.r0)) * z;
      const st = s.style(m.r0, m.c0); const cf = s.cf.length ? SU.evalCF(s, m.r0, m.c0, ver) : null; const tb = s.tables.length ? tableStyleAt(s, m.r0, m.c0) : null;
      ctx.fillStyle = (cf && cf.fill) || st.fill || (tb && tb.fill) || theme.cellBg; ctx.fillRect(al(x), al(y), al(w) - 1 / dpr, al(h) - 1 / dpr);
      if (!s.isEmpty(m.r0, m.c0)) drawCellText(s, m.r0, m.c0, x, y, w, h, true, m.c0, m.c1, ver, true);
    }
    // pass 4: borders, comments, table lines
    for (let r = r0; r <= r1; r++) {
      const rh = s.rowHeight(r) * z; if (!rh) continue; const y = G.rowY(r);
      for (let c = c0; c <= c1; c++) {
        const cw = s.colWidth(c) * z; if (!cw) continue;
        let st = s.style(r, c); const x = G.colX(c);
        if (s.tables.length) { const tb = tableStyleAt(s, r, c); if (tb && (tb.bT || tb.bB || tb.bL || tb.bR)) st = { ...st, bT: st.bT || tb.bT, bB: st.bB || tb.bB, bL: st.bL || tb.bL, bR: st.bR || tb.bR }; }
        if (st.bT || st.bB || st.bL || st.bR) drawBorders(x, y, cw, rh, st, s, r, c);
        if (s.comments.size && s.comments.has(key(r, c))) { ctx.fillStyle = '#d32f2f'; ctx.beginPath(); const rx = x + cw - 1, ry = y; ctx.moveTo(rx - 6 * Math.max(z, .8), ry); ctx.lineTo(rx, ry); ctx.lineTo(rx, ry + 6 * Math.max(z, .8)); ctx.fill(); }
        if (s.threads.size && s.threads.has(key(r, c))) { const th = s.threads.get(key(r, c)); const sz = 8 * Math.max(z, .8); ctx.fillStyle = th.resolved ? '#9aa3b2' : '#7A4FB5'; ctx.beginPath(); const rx = x + cw - 1, ry = y; if (s.comments.has(key(r, c))) { ctx.moveTo(rx - sz, ry + rh - 1); ctx.lineTo(rx, ry + rh - 1); ctx.lineTo(rx, ry + rh - 1 - sz); } else { ctx.moveTo(rx - sz, ry); ctx.lineTo(rx, ry); ctx.lineTo(rx, ry + sz); } ctx.fill(); }
      }
    }
    // table outline header separation
    for (const t of s.tables) {
      if (!SU.intersects(t.range, { r0, r1, c0, c1 })) continue;
      void t;
    }
    // filter buttons
    const fl = s.filter;
    if (fl && fl.range.r0 >= r0 && fl.range.r0 <= r1) { for (let c = Math.max(c0, fl.range.c0); c <= Math.min(c1, fl.range.c1); c++) drawFilterBtn(s, fl, c); }
    for (const t of s.tables) if (t.filterBtn !== false && t.header !== false && t.range.r0 >= r0 && t.range.r0 <= r1 && !(fl && fl.tableId === t.id)) for (let c = Math.max(c0, t.range.c0); c <= Math.min(c1, t.range.c1); c++) drawFilterBtn(s, { range: t.range, cols: {} }, c);
    // "Circle invalid data"
    if (s._circles && s._circles.size) {
      ctx.strokeStyle = '#E0261B'; ctx.lineWidth = 1.6;
      for (const k of s._circles) { const { r, c } = SU.unkey(k); if (r < r0 || r > r1 || c < c0 || c > c1) continue; const rc = G.cellRect(r, c); if (!rc.w || !rc.h) continue; ctx.beginPath(); ctx.ellipse(rc.x + rc.w / 2, rc.y + rc.h / 2, rc.w / 2 + 3, rc.h / 2 + 2, 0, 0, Math.PI * 2); ctx.stroke(); }
    }
  }
  function drawFilterBtn(s, fl, c) {
    const r = fl.range.r0; const z = G.z; if (!s.colWidth(c)) return;
    const x = G.colX(c), y = G.rowY(r), w = s.colWidth(c) * z, h = s.rowHeight(r) * z;
    const sz = Math.min(h - 4, 16 * Math.max(.8, Math.min(z, 1.4))); const bx = x + w - sz - 2, by = y + h - sz - 2;
    const active = fl.cols && fl.cols[c] && (fl.cols[c].values || fl.cols[c].cond);
    ctx.fillStyle = '#fff'; ctx.strokeStyle = active ? theme.accent : '#a6adba'; ctx.lineWidth = 1;
    ctx.beginPath(); ctx.roundRect ? ctx.roundRect(al(bx) + .5, al(by) + .5, sz - 1, sz - 1, 2) : ctx.rect(al(bx) + .5, al(by) + .5, sz - 1, sz - 1); ctx.fill(); ctx.stroke();
    ctx.fillStyle = active ? theme.accent : '#4b5363';
    if (active) { // funnel
      const cx = bx + sz / 2, cy = by + sz / 2; ctx.beginPath(); ctx.moveTo(cx - sz * .28, cy - sz * .24); ctx.lineTo(cx + sz * .28, cy - sz * .24); ctx.lineTo(cx + sz * .06, cy + .02 * sz); ctx.lineTo(cx + sz * .06, cy + sz * .26); ctx.lineTo(cx - sz * .06, cy + sz * .3); ctx.lineTo(cx - sz * .06, cy + .02 * sz); ctx.closePath(); ctx.fill();
    } else { const cx = bx + sz / 2, cy = by + sz / 2 + 1; ctx.beginPath(); ctx.moveTo(cx - sz * .22, cy - sz * .12); ctx.lineTo(cx + sz * .22, cy - sz * .12); ctx.lineTo(cx, cy + sz * .14); ctx.closePath(); ctx.fill(); }
  }
  G.filterBtnHit = (x, y) => {
    const s = S(); const z = G.z;
    const test = (range, isTable, t) => {
      const r = range.r0; const c = G.colAtX(x); if (c < range.c0 || c > range.c1 || G.rowAtY(y) !== r) return null;
      const cx = G.colX(c), cy = G.rowY(r), w = s.colWidth(c) * z, h = s.rowHeight(r) * z; const sz = Math.min(h - 4, 16 * Math.max(.8, Math.min(z, 1.4)));
      if (x >= cx + w - sz - 2 && x <= cx + w - 2 && y >= cy + h - sz - 2 && y <= cy + h - 2) return { c, r, table: isTable ? t : null };
      return null;
    };
    if (s.filter) { const h = test(s.filter.range); if (h) return h; }
    for (const t of s.tables) if (t.filterBtn !== false && t.header !== false && !(s.filter && s.filter.tableId === t.id)) { const h = test(t.range, true, t); if (h) return h; }
    return null;
  };
  function drawBar(x, y, w, h, bar) {
    const pad = 2, bw = Math.max(0, (w - pad * 2) * bar.pct);
    const g = ctx.createLinearGradient(x, 0, x + Math.max(bw, 1), 0); g.addColorStop(0, bar.color); g.addColorStop(1, SU.mixColor(bar.color, '#ffffff', .75));
    ctx.fillStyle = g; ctx.fillRect(x + pad, y + pad + 1, bw, h - pad * 2 - 2);
    ctx.strokeStyle = bar.color; ctx.lineWidth = 1 / dpr; ctx.strokeRect(ln(x + pad), ln(y + pad + 1), al(bw), al(h - pad * 2 - 3));
  }
  function drawBorders(x, y, w, h, st, s, r, c) {
    const one = (b, x1, y1, x2, y2, horiz) => {
      if (!b) return; const lw = b.style === 'thick' ? 3 : b.style === 'medium' || b.style === 'mediumDashed' ? 2 : 1;
      ctx.strokeStyle = b.color || '#000'; ctx.lineWidth = lw / (lw === 1 ? dpr : 1);
      ctx.setLineDash(b.style === 'dashed' || b.style === 'mediumDashed' ? [4, 2] : b.style === 'dotted' || b.style === 'hair' ? [1, 1.5] : []);
      const off = lw === 1 ? 0 : 0;
      if (b.style === 'double') { ctx.lineWidth = 1 / dpr; ctx.beginPath(); if (horiz) { ctx.moveTo(x1, ln(y1 - 1)); ctx.lineTo(x2, ln(y2 - 1)); ctx.moveTo(x1, ln(y1 + 1)); ctx.lineTo(x2, ln(y2 + 1)); } else { ctx.moveTo(ln(x1 - 1), y1); ctx.lineTo(ln(x2 - 1), y2); ctx.moveTo(ln(x1 + 1), y1); ctx.lineTo(ln(x2 + 1), y2); } ctx.stroke(); ctx.setLineDash([]); return; }
      ctx.beginPath(); if (horiz) { const yy = lw === 1 ? ln(y1) : al(y1); ctx.moveTo(x1 - (lw > 1 ? lw / 2 : 0), yy); ctx.lineTo(x2 + (lw > 1 ? lw / 2 : 0), yy); } else { const xx = lw === 1 ? ln(x1) : al(x1); ctx.moveTo(xx, y1 - (lw > 1 ? lw / 2 : 0)); ctx.lineTo(xx, y2 + (lw > 1 ? lw / 2 : 0)); }
      ctx.stroke(); ctx.setLineDash([]);
    };
    const m = s.mergeAt(r, c);
    if (m) { // for merged areas only draw borders on the outer edges using the cell styles
      if (r !== m.r0) st = { ...st, bT: null }; if (r !== m.r1) st = { ...st, bB: null }; if (c !== m.c0) st = { ...st, bL: null }; if (c !== m.c1) st = { ...st, bR: null };
    }
    one(st.bT, x, y - 1 / dpr, x + w, y - 1 / dpr, true); one(st.bB, x, y + h - 1 / dpr, x + w, y + h - 1 / dpr, true);
    one(st.bL, x - 1 / dpr, y, x - 1 / dpr, y + h, false); one(st.bR, x + w - 1 / dpr, y, x + w - 1 / dpr, y + h, false);
  }

  const measureCache = new Map();
  function measure(font, text) {
    const k = font + '\u0000' + text; let w = measureCache.get(k);
    if (w === undefined) { ctx.font = font; w = ctx.measureText(text).width; if (measureCache.size > 20000) measureCache.clear(); measureCache.set(k, w); }
    return w;
  }
  G.measure = measure;
  function wrapLines(font, text, maxW) {
    ctx.font = font; const out = [];
    for (const para of String(text).split('\n')) {
      const words = para.split(/(\s+)/); let line = '';
      for (const w of words) {
        const t = line + w;
        if (ctx.measureText(t).width > maxW && line.trim()) { out.push(line.trimEnd()); line = w.trimStart(); }
        else line = t;
        while (ctx.measureText(line).width > maxW && line.length > 1) { let i = line.length - 1; while (i > 1 && ctx.measureText(line.slice(0, i)).width > maxW) i--; out.push(line.slice(0, i)); line = line.slice(i); }
      }
      out.push(line);
    }
    return out;
  }
  G.wrapLines = wrapLines;

  function drawCellText(s, r, c, x, y, w, h, inView, vc0, vc1, ver, isMerge) {
    const z = G.z;
    const st = s.style(r, c);
    const fd = cellValueForDisplay(s, r, c);
    const v = fd ? fd.text : s.value(r, c);
    const cf = s.cf.length ? SU.evalCF(s, r, c, ver) : null; const tb = s.tables.length ? tableStyleAt(s, r, c) : null;
    let fmt = fd ? { text: fd.text } : SU.format(v, st.nf);
    if (fmt.text === '' ) return;
    const isNum = !fd && typeof v === 'number', isErr = !fd && SU.isErr(v), isBool = !fd && typeof v === 'boolean';
    let h_ = st.h || (isNum ? 'right' : (isErr || isBool) ? 'center' : 'left');
    if (st.h === 'fill') h_ = 'left';
    const effSt = (cf && cf.b) || (tb && tb.b) ? { ...st, b: true } : st;
    const font = fontStr(effSt, z, isNum);
    let color = (cf && cf.color) || fmt.color || st.color || (tb && tb.color) || (s.links.has(key(r, c)) ? '#1155cc' : '#000000');
    const padX = 3 * z + (st.indent || 0) * 9 * z;
    const fsz = (st.sz || SU.defSize) * 96 / 72 * z;
    // filter button room
    let btnRoom = 0; if (s.filter && r === s.filter.range.r0 && c >= s.filter.range.c0 && c <= s.filter.range.c1) btnRoom = 16 * z;
    for (const t of s.tables) if (t.header !== false && t.filterBtn !== false && r === t.range.r0 && SU.inRange(t.range, r, c)) btnRoom = 16 * z;
    ctx.font = font; ctx.fillStyle = color; ctx.textBaseline = 'alphabetic';
    const ascent = fsz * 0.78, lineH = fsz * 1.22;
    const vAlign = st.v || 'bottom';
    if (st.wrap && !fd) {
      const lines = wrapLines(font, fmt.text, Math.max(4, w - padX * 2 - btnRoom));
      const totalH = lines.length * lineH;
      let ty = vAlign === 'top' ? y + 2 * z + ascent : vAlign === 'middle' ? y + (h - totalH) / 2 + ascent : y + h - totalH - 2 * z + ascent + (lineH - fsz) / 2;
      if (!inView) return;
      ctx.save(); ctx.beginPath(); ctx.rect(x, y, w - 1, h - 1); ctx.clip();
      for (const L of lines) { const lw = measure(font, L); const tx = h_ === 'center' ? x + (w - lw) / 2 : h_ === 'right' ? x + w - padX - lw - btnRoom : x + padX; ctx.fillText(L, tx, ty); if (st.u || st.s) decorate(tx, ty, lw, fsz, st); ty += lineH; }
      ctx.restore(); return;
    }
    let text = fmt.text.replace(/\n/g, ' ');
    let tw = measure(font, text);
    const avail = w - padX * 2 - btnRoom;
    if (isNum && !fd && tw > w - 2 * z && !isMerge) { // numbers that don't fit -> ###
      if (!st.nf || st.nf === 'General') { // try fewer digits
        let prec = 10; while (tw > w - 4 * z && prec > 1) { const t2 = String(+v.toPrecision(prec)); text = t2.length < text.length ? t2 : text; if (Math.abs(v) >= 1e11 || /e/.test(t2)) { text = v.toExponential(Math.max(0, prec - 5)).replace('e', 'E').replace(/E([+-])(\d)$/, 'E$10$2'); } tw = measure(font, text); prec--; }
      }
      if (tw > w - 2 * z) { const hw_ = measure(font, '#'); text = '#'.repeat(Math.max(1, Math.floor((w - 2 * z) / hw_))); tw = measure(font, text); h_ = 'center'; }
    }
    let ty = vAlign === 'top' ? y + 2 * z + ascent : vAlign === 'middle' ? y + (h - fsz) / 2 + ascent - 1 * z : y + h - 4 * z - (fsz * 0.22) + 1;
    if (vAlign === 'bottom') ty = y + h - Math.max(3, fsz * 0.25);
    // determine clip region with overflow into empty neighbours
    let clipL = x, clipR = x + w;
    if (!isMerge && tw > avail && !isNum) {
      const free = (cc) => s.isEmpty(r, cc) && !s.mergeAt(r, cc) && !(s.sparklines.length && G.sparkAt(s, r, cc));
      if (h_ === 'left' || h_ === 'center') { let cc = c + 1, need = (h_ === 'center' ? (tw - avail) / 2 : tw - avail); while (need > 0 && cc < s.colCount && free(cc)) { const cw = s.colWidth(cc) * z; clipR += cw; need -= cw; cc++; } }
      if (h_ === 'right' || h_ === 'center') { let cc = c - 1, need = (h_ === 'center' ? (tw - avail) / 2 : tw - avail); while (need > 0 && cc >= 0 && free(cc)) { const cw = s.colWidth(cc) * z; clipL -= cw; need -= cw; cc--; } }
    }
    if (!inView) { const vx0 = G.colX(vc0), vx1 = G.colX(vc1 + 1); if (clipR <= vx0 || clipL >= vx1) return; }
    let tx;
    if (fmt.acct && !fd) { // accounting: symbol left, number right
      ctx.save(); ctx.beginPath(); ctx.rect(clipL, y, clipR - clipL - 1, h - 1); ctx.clip();
      if (clipR - clipL > w) { ctx.fillStyle = theme.cellBg; }
      ctx.fillStyle = color;
      if (fmt.acct.sym) ctx.fillText(fmt.acct.sym, x + padX, ty);
      const nw = measure(font, fmt.acct.num); ctx.fillText(fmt.acct.num, x + w - padX - nw, ty); ctx.restore(); return;
    }
    if (st.h === 'fill' && tw > 0) { const n = Math.max(1, Math.floor(avail / tw)); text = text.repeat(n); tw = measure(font, text); }
    tx = h_ === 'center' ? x + (w - tw) / 2 : h_ === 'right' ? x + w - padX - tw - btnRoom : x + padX;
    if (h_ === 'center' && btnRoom) tx = x + (w - btnRoom - tw) / 2;
    // hide gridlines beneath overflowing text
    if (clipR > x + w || clipL < x) {
      ctx.fillStyle = theme.cellBg;
      if (clipR > x + w) ctx.fillRect(x + w - 1.5 / dpr, y + 0.5, clipR - x - w, h - 1.5);
      if (clipL < x) ctx.fillRect(clipL, y + 0.5, x - clipL + 0.5, h - 1.5);
      ctx.fillStyle = color;
    }
    ctx.save(); ctx.beginPath(); ctx.rect(clipL, y, clipR - clipL - 1 / dpr, h - 1 / dpr); ctx.clip();
    ctx.fillText(text, tx, ty);
    if (st.u || st.s || s.links.has(key(r, c))) decorate(tx, ty, tw, fsz, s.links.has(key(r, c)) ? { ...st, u: true } : st);
    ctx.restore();
  }
  function decorate(tx, ty, tw, fsz, st) {
    ctx.strokeStyle = ctx.fillStyle; ctx.lineWidth = Math.max(1, fsz / 14);
    if (st.u) { ctx.beginPath(); const uy = ty + Math.max(1.5, fsz * 0.1); ctx.moveTo(tx, uy); ctx.lineTo(tx + tw, uy); ctx.stroke(); if (st.u === 'double') { ctx.beginPath(); ctx.moveTo(tx, uy + 2); ctx.lineTo(tx + tw, uy + 2); ctx.stroke(); } }
    if (st.s) { ctx.beginPath(); const sy = ty - fsz * 0.28; ctx.moveTo(tx, sy); ctx.lineTo(tx + tw, sy); ctx.stroke(); }
  }

  function drawSelection(s) {
    const z = G.z;
    // formula reference highlights
    for (const h of G.refHL) {
      const rr = G.rangeRect(clipRange(h.range)); ctx.fillStyle = h.color + '1f'; ctx.fillRect(rr.x, rr.y, rr.w, rr.h);
      ctx.strokeStyle = h.color; ctx.lineWidth = 2; ctx.strokeRect(al(rr.x) + .5, al(rr.y) + .5, al(rr.w) - 1.5, al(rr.h) - 1.5);
      ctx.fillStyle = h.color; for (const [px, py] of [[rr.x, rr.y], [rr.x + rr.w, rr.y], [rr.x, rr.y + rr.h], [rr.x + rr.w, rr.y + rr.h]]) ctx.fillRect(px - 2.5, py - 2.5, 5, 5);
    }
    if (G.hideSelection) return;
    const sels = G.selections();
    const a = s.active;
    const ar = G.cellRect(a.r, a.c);
    const main = clipRange(s.sel); const mr = G.rangeRect(main);
    for (const g0 of sels) {
      const rr = G.rangeRect(clipRange(g0));
      ctx.fillStyle = theme.selFill; ctx.beginPath(); ctx.rect(rr.x, rr.y, rr.w, rr.h);
      if (g0 === s.sel && SU.inRange(g0, a.r, a.c)) ctx.rect(ar.x, ar.y, ar.w, ar.h);
      ctx.fill('evenodd');
    }
    ctx.strokeStyle = theme.accent; ctx.lineWidth = 2;
    ctx.strokeRect(al(mr.x) - 0.0, al(mr.y) - 0.0, al(mr.w) - 1, al(mr.h) - 1);
    if (sels.length > 1) { for (const g of sels.slice(0, -1)) { const rr = G.rangeRect(clipRange(g)); ctx.lineWidth = 1; ctx.strokeRect(ln(rr.x), ln(rr.y), al(rr.w) - 1, al(rr.h) - 1); } }
    // active cell inner frame
    if (!(main.r0 === main.r1 && main.c0 === main.c1) && !(ar.m && ar.m.r0 === main.r0 && ar.m.r1 === main.r1 && ar.m.c0 === main.c0 && ar.m.c1 === main.c1)) {
      ctx.lineWidth = 1; ctx.strokeStyle = theme.accent; ctx.strokeRect(ln(ar.x), ln(ar.y), al(ar.w) - 2 / dpr, al(ar.h) - 2 / dpr);
    }
    // fill handle
    if (!G.editing && sels.length === 1) { const hx = mr.x + mr.w - 1, hy = mr.y + mr.h - 1; ctx.fillStyle = '#fff'; ctx.fillRect(hx - 4.5, hy - 4.5, 9, 9); ctx.fillStyle = theme.accent; ctx.fillRect(hx - 3.5, hy - 3.5, 7, 7); }
    // fill preview
    if (G.fillPreview) { const rr = G.rangeRect(G.fillPreview); ctx.setLineDash([3, 2]); ctx.strokeStyle = '#596274'; ctx.lineWidth = 1; ctx.strokeRect(ln(rr.x), ln(rr.y), al(rr.w) - 1, al(rr.h) - 1); ctx.setLineDash([]); }
    // copy marquee
    if (G.copyRange && G.copyRange.sheet === s) { const rr = G.rangeRect(G.copyRange.range); ctx.setLineDash([4, 3]); ctx.lineDashOffset = -G.antsOffset || 0; ctx.strokeStyle = theme.accent; ctx.lineWidth = 2; ctx.strokeRect(al(rr.x) + 1, al(rr.y) + 1, al(rr.w) - 3, al(rr.h) - 3); ctx.setLineDash([]); }
    // move/drag preview
    if (G.movePreview) { const rr = G.rangeRect(G.movePreview); ctx.strokeStyle = '#596274'; ctx.lineWidth = 2; ctx.strokeRect(al(rr.x), al(rr.y), al(rr.w) - 1, al(rr.h) - 1); }
  }
  function clipRange(g) { const s = S(); return { r0: g.r0, c0: g.c0, r1: Math.min(g.r1, s.rowCount - 1), c1: Math.min(g.c1, s.colCount - 1) }; }

  G.selections = () => { const s = S(); return [...(s.extraSel || []), s.sel]; };

  function drawHeaders(s, hw, hh, f, mainCols, mainRows, frCols, frRows) {
    if (!s.showHeadings) return;
    const z = G.z; const sels = G.selections(); const ogw = G.ogW(), ogh = G.ogH();
    const fsz = Math.round(11 * Math.max(.8, Math.min(z, 1.6)));
    ctx.font = `${fsz}px ${getFontUI()}`; ctx.textBaseline = 'middle'; ctx.textAlign = 'center';
    const colSel = (c) => sels.some(g => c >= g.c0 && c <= g.c1), rowSel = (r) => sels.some(g => r >= g.r0 && r <= g.r1);
    const colFull = (c) => sels.some(g => c >= g.c0 && c <= g.c1 && g.r0 === 0 && g.r1 >= s.rowCount - 1), rowFull = (r) => sels.some(g => r >= g.r0 && r <= g.r1 && g.c0 === 0 && g.c1 >= s.colCount - 1);
    // column header
    const drawColHdr = (c0, c1, clipX, clipW) => {
      ctx.save(); ctx.beginPath(); ctx.rect(clipX, 0, clipW, hh); ctx.clip();
      const hy = ogh, hH = hh - ogh;
      for (let c = c0; c <= c1; c++) {
        const w = s.colWidth(c) * z; const x = G.colX(c);
        if (!w) { ctx.strokeStyle = theme.hdrLine; ctx.lineWidth = 1; ctx.beginPath(); ctx.moveTo(ln(x - 2), hy + 3); ctx.lineTo(ln(x - 2), hh - 3); ctx.moveTo(ln(x + 1), hy + 3); ctx.lineTo(ln(x + 1), hh - 3); ctx.stroke(); continue; }
        const sel = colSel(c), full = colFull(c);
        ctx.fillStyle = full ? theme.hdrFullBg : sel ? theme.hdrSelBg : theme.hdrBg; ctx.fillRect(x, hy, w, hH);
        ctx.fillStyle = full ? '#fff' : sel ? theme.hdrSelText : theme.hdrText; ctx.font = `${sel ? '600 ' : ''}${fsz}px ${getFontUI()}`;
        ctx.fillText(SU.colName(c), x + w / 2, hy + hH / 2 + 0.5);
        ctx.strokeStyle = theme.hdrLine; ctx.lineWidth = 1 / dpr; ctx.beginPath(); ctx.moveTo(ln(x + w - 1 / dpr), hy); ctx.lineTo(ln(x + w - 1 / dpr), hh); ctx.stroke();
        if (sel && !full) { ctx.fillStyle = theme.accent; ctx.fillRect(x, hh - 2, w, 2); }
      }
      ctx.restore();
    };
    const drawRowHdr = (r0, r1, clipY, clipH) => {
      ctx.save(); ctx.beginPath(); ctx.rect(0, clipY, hw, clipH); ctx.clip();
      const hx = ogw, hW = hw - ogw;
      for (let r = r0; r <= r1; r++) {
        const h = s.rowHeight(r) * z; const y = G.rowY(r);
        if (!h) { ctx.strokeStyle = theme.hdrLine; ctx.lineWidth = 1; ctx.beginPath(); ctx.moveTo(hx + 3, ln(y - 2)); ctx.lineTo(hw - 3, ln(y - 2)); ctx.moveTo(hx + 3, ln(y + 1)); ctx.lineTo(hw - 3, ln(y + 1)); ctx.stroke(); continue; }
        const sel = rowSel(r), full = rowFull(r);
        ctx.fillStyle = full ? theme.hdrFullBg : sel ? theme.hdrSelBg : theme.hdrBg; ctx.fillRect(hx, y, hW, h);
        ctx.fillStyle = full ? '#fff' : sel ? theme.hdrSelText : theme.hdrText; ctx.font = `${sel ? '600 ' : ''}${fsz}px ${getFontUI()}`;
        ctx.fillText(String(r + 1), hx + hW / 2, y + h / 2 + 0.5);
        ctx.strokeStyle = theme.hdrLine; ctx.lineWidth = 1 / dpr; ctx.beginPath(); ctx.moveTo(hx, ln(y + h - 1 / dpr)); ctx.lineTo(hw, ln(y + h - 1 / dpr)); ctx.stroke();
        if (sel && !full) { ctx.fillStyle = theme.accent; ctx.fillRect(hw - 2, y, 2, h); }
      }
      ctx.restore();
    };
    drawColHdr(mainCols[0], mainCols[1], hw + f.fw, W - hw - f.fw); if (frCols) drawColHdr(frCols[0], frCols[1], hw, f.fw);
    drawRowHdr(mainRows[0], mainRows[1], hh + f.fh, H - hh - f.fh); if (frRows) drawRowHdr(frRows[0], frRows[1], hh, f.fh);
    // header borders + corner
    ctx.fillStyle = theme.hdrBg; ctx.fillRect(0, 0, hw, hh);
    const ch = hh - ogh; ctx.fillStyle = theme.hdrLine; ctx.beginPath(); ctx.moveTo(hw - 4, ogh + 5); ctx.lineTo(hw - 4, hh - 4); ctx.lineTo(hw - ch + 5, hh - 4); ctx.closePath(); ctx.fill();
    if (ogw || ogh) drawOutline(s, hw, hh, ogw, ogh, mainCols, mainRows, frCols, frRows);
    ctx.strokeStyle = theme.hdrLine; ctx.lineWidth = 1 / dpr; ctx.beginPath(); ctx.moveTo(0, ln(hh - 1 / dpr)); ctx.lineTo(W, ln(hh - 1 / dpr)); ctx.moveTo(ln(hw - 1 / dpr), 0); ctx.lineTo(ln(hw - 1 / dpr), H); ctx.stroke();
    ctx.textAlign = 'start'; ctx.textBaseline = 'alphabetic';
  }
  /* outline bars, +/- buttons and level buttons. Group of rows a..b has its summary row at b+1 (Excel default: summary below). */
  const outlineHits = [];
  function obtn(x, y, label, hot) {
    const sz = 13; ctx.fillStyle = hot ? theme.hdrSelBg : (theme.dark ? '#2a3243' : '#fff'); ctx.strokeStyle = theme.dark ? '#7d879b' : '#8a93a3'; ctx.lineWidth = 1;
    ctx.fillRect(al(x - sz / 2) + .5, al(y - sz / 2) + .5, sz - 1, sz - 1); ctx.strokeRect(al(x - sz / 2) + .5, al(y - sz / 2) + .5, sz - 1, sz - 1);
    ctx.fillStyle = theme.hdrText; ctx.font = `600 10px ${getFontUI()}`; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    if (label === '+' || label === '-') { ctx.fillRect(al(x - 3.5), al(y) - .5, 7, 1.2); if (label === '+') ctx.fillRect(al(x) - .5, al(y - 3.5), 1.2, 7); }
    else ctx.fillText(label, x, y + .5);
  }
  function drawOutline(s, hw, hh, ogw, ogh, mainCols, mainRows) {
    outlineHits.length = 0; const z = G.z;
    ctx.fillStyle = theme.hdrBg; if (ogw) ctx.fillRect(0, 0, ogw, H); if (ogh) ctx.fillRect(0, 0, W, ogh);
    ctx.strokeStyle = theme.hdrLine; ctx.lineWidth = 1 / dpr; ctx.beginPath(); if (ogw) { ctx.moveTo(ln(ogw - 1 / dpr), 0); ctx.lineTo(ln(ogw - 1 / dpr), H); } if (ogh) { ctx.moveTo(0, ln(ogh - 1 / dpr)); ctx.lineTo(W, ln(ogh - 1 / dpr)); } ctx.stroke();
    if (ogw) {
      const info = outlineInfo(s.outlineR);
      for (let L = 1; L <= info.max + 1; L++) { const x = 4 + (L - 1) * OG + OG / 2 - 2, y = hh - 11; obtn(x, y, String(L)); outlineHits.push({ x0: x - 7, x1: x + 7, y0: y - 7, y1: y + 7, kind: 'level', axis: 'r', L }); }
      ctx.save(); ctx.beginPath(); ctx.rect(0, hh, ogw, H - hh); ctx.clip();
      ctx.strokeStyle = theme.dark ? '#7d879b' : '#7a8496'; ctx.lineWidth = 1.3;
      for (const g of info.groups) {
        if (g.b + 1 < mainRows[0] - 1 && g.b + 1 < (s.freeze.r || 0)) continue;
        if (g.a > mainRows[1] + 1) continue;
        const x = 4 + (g.L - 1) * OG + OG / 2 - 2; let collapsed = true; for (let r = g.a; r <= g.b; r++) if (s.rowHeight(r)) { collapsed = false; break; }
        const sy = G.rowY(g.b + 1), sh = s.rowHeight(g.b + 1) * z; const by = sy + Math.max(sh, 12) / 2;
        if (!collapsed) { const ty = G.rowY(g.a); ctx.beginPath(); ctx.moveTo(x, ty + 2); ctx.lineTo(x, by - 7); ctx.moveTo(x, ty + 2); ctx.lineTo(x + 4, ty + 2); ctx.stroke(); }
        obtn(x, by, collapsed ? '+' : '-'); outlineHits.push({ x0: x - 7, x1: x + 7, y0: by - 7, y1: by + 7, kind: 'toggle', axis: 'r', g, collapsed });
      }
      ctx.restore();
    }
    if (ogh) {
      const info = outlineInfo(s.outlineC);
      for (let L = 1; L <= info.max + 1; L++) { const x = hw - 11, y = 4 + (L - 1) * OG + OG / 2 - 2; obtn(x, y, String(L)); outlineHits.push({ x0: x - 7, x1: x + 7, y0: y - 7, y1: y + 7, kind: 'level', axis: 'c', L }); }
      ctx.save(); ctx.beginPath(); ctx.rect(hw, 0, W - hw, ogh); ctx.clip();
      ctx.strokeStyle = theme.dark ? '#7d879b' : '#7a8496'; ctx.lineWidth = 1.3;
      for (const g of info.groups) {
        if (g.a > mainCols[1] + 1) continue;
        const y = 4 + (g.L - 1) * OG + OG / 2 - 2; let collapsed = true; for (let c = g.a; c <= g.b; c++) if (s.colWidth(c)) { collapsed = false; break; }
        const sx = G.colX(g.b + 1), sw = s.colWidth(g.b + 1) * z; const bx = sx + Math.min(Math.max(sw, 12), 24) / 2;
        if (!collapsed) { const tx = G.colX(g.a); ctx.beginPath(); ctx.moveTo(tx + 2, y); ctx.lineTo(bx - 7, y); ctx.moveTo(tx + 2, y); ctx.lineTo(tx + 2, y + 4); ctx.stroke(); }
        obtn(bx, y, collapsed ? '+' : '-'); outlineHits.push({ x0: bx - 7, x1: bx + 7, y0: y - 7, y1: y + 7, kind: 'toggle', axis: 'c', g, collapsed });
      }
      ctx.restore();
    }
  }

  /* ---------- mouse ---------- */
  function local(e) { const r = host.getBoundingClientRect(); return { x: e.clientX - r.left, y: e.clientY - r.top }; }
  G.local = local;
  function hitZone(p) {
    const hw = curHW, hh = curHH; const s = S(); const z = G.z;
    if (p.x >= scroller.clientWidth || p.y >= scroller.clientHeight) return { zone: 'scrollbar' };
    if ((p.x < G.ogW() || p.y < G.ogH()) && outlineHits.length) { const o = outlineHits.find(o => p.x >= o.x0 && p.x <= o.x1 && p.y >= o.y0 && p.y <= o.y1); return o ? { zone: 'outline', o } : { zone: 'gutter' }; }
    if (p.x < hw && p.y < hh) return { zone: 'corner' };
    if (p.y < hh) {
      const c = G.colAtX(p.x); const x = G.colX(c), w = s.colWidth(c) * z;
      // resize handle near right edge of a column (or left edge of next)
      if (p.x > x + w - 4 && p.x < x + w + 4) return { zone: 'colresize', c };
      if (p.x < x + 3 && c > 0) { let pc = c - 1; while (pc > 0 && !s.colWidth(pc)) pc--; return { zone: 'colresize', c: s.colWidth(c - 1) ? c - 1 : pc }; }
      return { zone: 'colhdr', c };
    }
    if (p.x < hw) {
      const r = G.rowAtY(p.y); const y = G.rowY(r), h = s.rowHeight(r) * z;
      if (p.y > y + h - 3 && p.y < y + h + 3) return { zone: 'rowresize', r };
      if (p.y < y + 2 && r > 0) return { zone: 'rowresize', r: r - 1 };
      return { zone: 'rowhdr', r };
    }
    const r = G.rowAtY(p.y), c = G.colAtX(p.x);
    // fill handle
    const mr = G.rangeRect(clipRange(s.sel)); const hx = mr.x + mr.w - 1, hy = mr.y + mr.h - 1;
    if (!G.editing && Math.abs(p.x - hx) <= 5 && Math.abs(p.y - hy) <= 5 && G.selections().length === 1) return { zone: 'fill', r, c };
    const fb = G.filterBtnHit(p.x, p.y); if (fb) return { zone: 'filterbtn', ...fb };
    // selection border (move)
    const onEdge = (Math.abs(p.x - mr.x) <= 3 || Math.abs(p.x - (mr.x + mr.w)) <= 3) && p.y >= mr.y && p.y <= mr.y + mr.h || (Math.abs(p.y - mr.y) <= 3 || Math.abs(p.y - (mr.y + mr.h)) <= 3) && p.x >= mr.x && p.x <= mr.x + mr.w;
    if (onEdge && !G.editing) return { zone: 'move', r, c };
    return { zone: 'cell', r, c };
  }
  G.hitZone = hitZone;
  function onHover(e) {
    if (G.dragging) return;
    const p = local(e); const h = hitZone(p);
    scroller.style.cursor = h.zone === 'outline' ? 'pointer' : h.zone === 'colresize' ? 'col-resize' : h.zone === 'rowresize' ? 'row-resize' : h.zone === 'fill' ? 'crosshair' : h.zone === 'colhdr' ? 's-resize' : h.zone === 'rowhdr' ? 'e-resize' : h.zone === 'move' ? 'move' : h.zone === 'filterbtn' ? 'default' : h.zone === 'cell' ? 'cell' : 'default';
    if (h.zone === 'cell' || h.zone === 'filterbtn') { const s = S(); const m = s.mergeAt(h.r, h.c); G.onHoverCell && G.onHoverCell(m ? { r: m.r0, c: m.c0 } : { r: h.r, c: h.c }, e); }
    else G.onHoverCell && G.onHoverCell(null);
  }

  let autoScrollTimer = null, lastMouse = null;
  function startAutoScroll(onTick) {
    stopAutoScroll();
    autoScrollTimer = setInterval(() => {
      if (!lastMouse) return; const p = lastMouse; const f = frozen();
      let dx = 0, dy = 0;
      if (p.x > scroller.clientWidth - 4) dx = Math.min(60, (p.x - scroller.clientWidth) / 2 + 10); else if (p.x < curHW + f.fw && scroller.scrollLeft > 0 && G.dragging && G.dragging.kind !== 'row') dx = -Math.min(60, (curHW + f.fw - p.x) / 2 + 10);
      if (p.y > scroller.clientHeight - 4) dy = Math.min(80, (p.y - scroller.clientHeight) / 2 + 10); else if (p.y < curHH + f.fh && scroller.scrollTop > 0 && G.dragging && G.dragging.kind !== 'col') dy = -Math.min(80, (curHH + f.fh - p.y) / 2 + 10);
      if (dx || dy) { scroller.scrollLeft += dx; scroller.scrollTop += dy; onTick(p); }
    }, 40);
  }
  function stopAutoScroll() { clearInterval(autoScrollTimer); autoScrollTimer = null; }

  function onMouseDown(e) {
    if (e.button === 1) return;
    const p = local(e); const h = hitZone(p); if (h.zone === 'scrollbar') return;
    const s = S();
    if (e.button === 2) { // right click: keep selection if inside
      if (h.zone === 'cell' && !G.selections().some(g => SU.inRange(g, h.r, h.c))) G.select(h.r, h.c);
      else if (h.zone === 'colhdr' && !(s.sel.r0 === 0 && s.sel.r1 >= s.rowCount - 1 && h.c >= s.sel.c0 && h.c <= s.sel.c1)) G.selectCols(h.c, h.c);
      else if (h.zone === 'rowhdr' && !(s.sel.c0 === 0 && s.sel.c1 >= s.colCount - 1 && h.r >= s.sel.r0 && h.r <= s.sel.r1)) G.selectRows(h.r, h.r);
      return;
    }
    e.preventDefault();
    if (G.onMouseDownHook && G.onMouseDownHook(e, h, p)) return;
    G.focus();
    const up = (fn) => { const mv = (ev) => { lastMouse = local(ev); fn(lastMouse, ev); }; const u = (ev) => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', u); stopAutoScroll(); const d = G.dragging; G.dragging = null; d && d.end && d.end(local(ev), ev); G.render(); }; document.addEventListener('mousemove', mv); document.addEventListener('mouseup', u); };
    if (h.zone === 'corner') { G.selectAll(); return; }
    if (h.zone === 'outline') { G.onOutline && G.onOutline(h.o); return; }
    if (h.zone === 'gutter') return;
    if (h.zone === 'filterbtn') { G.onFilterButton && G.onFilterButton(h); return; }
    if (h.zone === 'colresize' || h.zone === 'rowresize') {
      const isCol = h.zone === 'colresize'; const idx = isCol ? h.c : h.r; const start = isCol ? p.x : p.y; const orig = isCol ? s.rawColWidth(idx) : s.rawRowHeight(idx); const z = G.z;
      const guide = document.createElement('div'); guide.className = 'resize-guide ' + (isCol ? 'v' : 'h'); host.appendChild(guide);
      const tip = document.createElement('div'); tip.className = 'resize-tip'; host.appendChild(tip);
      let size = orig;
      const upd = (pp) => {
        size = Math.max(0, Math.round(orig + ((isCol ? pp.x : pp.y) - start) / z));
        const pos = (isCol ? G.colX(idx) : G.rowY(idx)) + size * z;
        if (isCol) { guide.style.left = pos + 'px'; tip.style.left = pos + 6 + 'px'; tip.style.top = '2px'; tip.textContent = `Width: ${(size / 7.5).toFixed(2)} (${size} px)`; }
        else { guide.style.top = pos + 'px'; tip.style.top = pos + 4 + 'px'; tip.style.left = curHW + 4 + 'px'; tip.textContent = `Height: ${(size * 0.75).toFixed(2)} (${size} px)`; }
      };
      upd(p);
      G.dragging = { kind: 'resize', end: () => { guide.remove(); tip.remove(); if (size !== orig) G.onResize && G.onResize(isCol ? 'col' : 'row', idx, size); } };
      up((pp) => upd(pp)); return;
    }
    if (h.zone === 'fill') {
      const src = { ...s.sel };
      G.dragging = { kind: 'fill', end: () => { const t = G.fillPreview; G.fillPreview = null; if (t) G.onFill && G.onFill(src, t); } };
      const upd = (pp) => {
        const r = G.rowAtY(Math.min(pp.y, scroller.clientHeight - 1)), c = G.colAtX(Math.min(pp.x, scroller.clientWidth - 1));
        let t = null;
        const dr = r > src.r1 ? r - src.r1 : r < src.r0 ? r - src.r0 : 0, dc = c > src.c1 ? c - src.c1 : c < src.c0 ? c - src.c0 : 0;
        if (Math.abs(dr) * 1.0 >= Math.abs(dc) && dr) t = dr > 0 ? { ...src, r1: r } : { ...src, r0: r };
        else if (dc) t = dc > 0 ? { ...src, c1: c } : { ...src, c0: c };
        else if (r >= src.r0 && r < src.r1 && c >= src.c0 && c <= src.c1 && src.r1 > src.r0) t = { ...src, r1: r, shrink: true };
        G.fillPreview = t; G.render();
      };
      startAutoScroll(upd); up(upd); return;
    }
    if (h.zone === 'move') {
      const src = G.selections()[0]; const off = { r: h.r - src.r0, c: h.c - src.c0 };
      G.dragging = { kind: 'move', end: (pp, ev) => { const t = G.movePreview; G.movePreview = null; if (t && (t.r0 !== src.r0 || t.c0 !== src.c0)) G.onMoveCells && G.onMoveCells(src, t, ev.ctrlKey); } };
      const upd = (pp) => { const r = Math.max(0, G.rowAtY(pp.y) - off.r), c = Math.max(0, G.colAtX(pp.x) - off.c); G.movePreview = { r0: r, c0: c, r1: r + src.r1 - src.r0, c1: c + src.c1 - src.c0 }; G.render(); };
      startAutoScroll(upd); up(upd); return;
    }
    // point mode (formula editing)
    if (G.editing && G.pointMode && G.pointMode(h, e)) {
      const anchor = { r: h.r, c: h.c };
      G.dragging = { kind: 'point' };
      const upd = (pp) => { const r = G.rowAtY(pp.y), c = G.colAtX(pp.x); G.pointDrag && G.pointDrag(anchor, { r, c }); };
      startAutoScroll(upd); up(upd); return;
    }
    if (G.editing) { G.commitEdit && G.commitEdit(); if (G.editing) return; }
    if (h.zone === 'colhdr') {
      const a = e.shiftKey ? s.active.c : h.c;
      if (e.ctrlKey) s.extraSel = [...(s.extraSel || []), s.sel]; else if (!e.shiftKey) s.extraSel = [];
      G.selectCols(a, h.c, e.shiftKey ? null : h.c);
      G.dragging = { kind: 'col' };
      const upd = (pp) => { G.selectCols(a, G.colAtX(pp.x), null, true); };
      startAutoScroll(upd); up(upd); return;
    }
    if (h.zone === 'rowhdr') {
      const a = e.shiftKey ? s.active.r : h.r;
      if (e.ctrlKey) s.extraSel = [...(s.extraSel || []), s.sel]; else if (!e.shiftKey) s.extraSel = [];
      G.selectRows(a, h.r, e.shiftKey ? null : h.r);
      G.dragging = { kind: 'row' };
      const upd = (pp) => { G.selectRows(a, G.rowAtY(pp.y), null, true); };
      startAutoScroll(upd); up(upd); return;
    }
    // cell
    if (e.shiftKey) { G.extendTo(h.r, h.c); }
    else if (e.ctrlKey) { s.extraSel = [...(s.extraSel || []), s.sel]; G.select(h.r, h.c, true); }
    else { s.extraSel = []; G.select(h.r, h.c); }
    const anchor = { ...s.active };
    G.dragging = { kind: 'sel', end: () => G.onSelectionEnd && G.onSelectionEnd() };
    const upd = (pp) => { const r = G.rowAtY(Math.min(Math.max(pp.y, curHH), scroller.clientHeight - 1)), c = G.colAtX(Math.min(Math.max(pp.x, curHW), scroller.clientWidth - 1)); G.setSel(anchor, { r, c }, true); };
    startAutoScroll(upd); up(upd);
  }
  function onDblClick(e) {
    const p = local(e); const h = hitZone(p);
    if (h.zone === 'colresize') { G.onAutofit && G.onAutofit('col', h.c); return; }
    if (h.zone === 'rowresize') { G.onAutofit && G.onAutofit('row', h.r); return; }
    if (h.zone === 'fill') { G.onFillDouble && G.onFillDouble(); return; }
    if (h.zone === 'cell' && !G.editing) { G.onCellDblClick && G.onCellDblClick(h.r, h.c, e); }
  }
  function onContext(e) { e.preventDefault(); const p = local(e); const h = hitZone(p); G.onContextMenu && G.onContextMenu(h, e); }

  /* ---------- selection API ---------- */
  G.notify = () => { G.render(); G.onSelect && G.onSelect(); };
  G.select = (r, c, keepExtra) => { const s = S(); if (!keepExtra) s.extraSel = []; r = SU.clamp(r, 0, SU.MAXR - 1); c = SU.clamp(c, 0, SU.MAXC - 1); const m = s.mergeAt(r, c); if (m) { r = m.r0; c = m.c0; } s.active = { r, c }; s.sel = m ? { ...m } : { r0: r, c0: c, r1: r, c1: c }; G.scrollIntoView(r, c); G.notify(); };
  G.setSel = (a, b, fromDrag) => { const s = S(); s.ensureSize(Math.max(a.r, b.r), Math.max(a.c, b.c)); s.active = { r: a.r, c: a.c }; s.sel = s.expandMerges(SU.norm({ r0: a.r, c0: a.c, r1: b.r, c1: b.c })); if (!fromDrag) G.scrollIntoView(b.r, b.c); else G.updateSpacer(); G.notify(); };
  G.selectRange = (g, active) => { const s = S(); s.extraSel = []; s.sel = s.expandMerges(SU.norm(g)); s.active = active || { r: s.sel.r0, c: s.sel.c0 }; G.scrollIntoView(s.active.r, s.active.c); G.notify(); };
  G.extendTo = (r, c) => { const s = S(); G.setSel(s.active, { r, c }); };
  G.selectCols = (a, b, act, drag) => { const s = S(); const c0 = Math.min(a, b), c1 = Math.max(a, b); s.sel = { r0: 0, c0, r1: SU.MAXR - 1, c1 }; s.active = { r: drag ? s.active.r : G.firstVisibleRow(), c: act !== null && act !== undefined ? act : (drag ? s.active.c : a) }; G.notify(); };
  G.selectRows = (a, b, act, drag) => { const s = S(); const r0 = Math.min(a, b), r1 = Math.max(a, b); s.sel = { r0, c0: 0, r1, c1: SU.MAXC - 1 }; s.active = { r: act !== null && act !== undefined ? act : (drag ? s.active.r : a), c: drag ? s.active.c : G.firstVisibleCol() }; G.notify(); };
  G.selectAll = () => { const s = S(); s.extraSel = []; s.sel = { r0: 0, c0: 0, r1: SU.MAXR - 1, c1: SU.MAXC - 1 }; G.notify(); };
  G.firstVisibleRow = () => { const s = S(); return s.freeze.r ? 0 : s.rowAt(scroller.scrollTop / G.z); };
  G.firstVisibleCol = () => { const s = S(); return s.freeze.c ? 0 : s.colAt(scroller.scrollLeft / G.z); };
  G.isWholeCols = (g) => g.r0 === 0 && g.r1 >= SU.MAXR - 1;
  G.isWholeRows = (g) => g.c0 === 0 && g.c1 >= SU.MAXC - 1;
  G.focus = () => { const k = document.getElementById('gkeys'); if (k && document.activeElement !== k) k.focus({ preventScroll: true }); };
  G.pageRows = () => Math.max(1, Math.floor((scroller.clientHeight - curHH) / (SU.DEF_H * G.z)) - 1);

  /* marching ants */
  setInterval(() => { if (G.copyRange) { G.antsOffset = ((G.antsOffset || 0) + 1) % 14; G.render(); } }, 120);
})();
