/* BTS Slides 1.2 — Export tab: handouts PDF (1/2/3/4/6/9 per page, notes lines), notes pages PDF,
 * slides as PNG pictures (current slide, or all slides in one .zip), text outline, document properties.
 * PDFs are made with the window's printToPDF (same renderer as the editor, bundled fonts embedded);
 * pictures are rendered from that PDF with pdf.js, so text looks exactly like the slide.
 */
(function () {
  const $ = (s, r = document) => r.querySelector(s);
  const esc = (s) => BTS.esc(s);
  const X = window.SLX = window.SLX || { cmd: {}, ribbon: [] };
  const P = () => SL.pres;
  const MM = 96 / 25.4;
  const PAPER = { A4: [210, 297], Letter: [215.9, 279.4] };

  /* Lay pages out in #print-root and print them to PDF bytes. pages: [HTMLElement], size in mm. */
  async function printPages(pages, wmm, hmm) {
    if (!BTS.desktop) { BTS.toast('PDF export needs the BTS Office desktop app', { error: true }); return null; }
    if (E.editing) E.endEdit();
    if (E.view === 'master') X.masterClose();
    await SL.ensureFonts(P());
    const root = $('#print-root'); root.innerHTML = ''; root.style.width = wmm + 'mm'; // without this Chromium shrinks pages narrower than the window
    const st = document.createElement('style'); st.id = 'print-page-style'; st.textContent = `@page { size: ${wmm}mm ${hmm}mm; margin: 0; }`; document.head.appendChild(st);
    pages.forEach(p => { p.classList.add('print-page'); p.style.width = wmm + 'mm'; p.style.height = hmm + 'mm'; root.appendChild(p); });
    root.querySelectorAll('.sl-slide').forEach(n => E.fitAllText(n));
    // print thumbnails with CSS zoom instead of transform: Chromium mis-paginates large transformed boxes (pages shrink after page 2)
    root.querySelectorAll('.sl-thumb > .sl-slide').forEach(n => { const m = /scale\(([\d.]+)\)/.exec(n.style.transform || ''); if (m) { n.style.transform = 'none'; n.style.zoom = m[1]; } });
    document.body.classList.add('printing');
    try {
      await document.fonts.ready; await new Promise(r => setTimeout(r, 150));
      return await bts.printToPDF({ pageSize: { width: wmm / 25.4, height: hmm / 25.4 }, landscape: false, preferCSSPageSize: true });
    } finally { document.body.classList.remove('printing'); root.innerHTML = ''; root.style.width = ''; st.remove(); }
  }
  X.printPages = printPages;
  async function saveBytes(name, data, filterName, ext) {
    if (!data) return false;
    if (!BTS.desktop) { BTS.download(name, data); return true; }
    const p = await bts.saveDialog(name, [{ name: filterName, extensions: [ext] }]); if (!p) return false;
    const r = await bts.writeFile(p, data); if (r.error) { BTS.toast(r.error, { error: true }); return false; }
    BTS.toast('Saved ' + r.name); return true;
  }
  X.saveBytes = saveBytes;
  const slideList = (hidden) => P().slides.map((s, i) => ({ s, i })).filter(x => hidden || !x.s.hidden);
  const thumb = (s, i, wpx) => { const t = SL.thumb(s, P(), wpx, i); t.classList.add('ho-slide'); return t; };
  const header = (pg, n, total, o) => { pg.insertAdjacentHTML('beforeend', `<div class="ho-head"><span>${esc(o.header || BTS.title)}</span><span>${esc(new Date().toLocaleDateString())}</span></div>${o.pageNums ? `<div class="ho-foot">${n} / ${total}</div>` : ''}`); };

  /* ---------------- handouts ---------------- */
  X.buildHandoutPages = (o) => {
    const [pw, ph] = o.landscape ? PAPER[o.paper].slice().reverse() : PAPER[o.paper];
    const list = slideList(o.hidden); const per = o.per; const pages = [];
    const ar = SL.H(P()) / SL.W(P());
    const M = 14, top = 22, bot = 16; const cw = pw - 2 * M, ch = ph - top - bot;
    const grid = { 1: [1, 1], 2: [1, 2], 3: [1, 3], 4: [2, 2], 6: [2, 3], 9: [3, 3] }[per] || [2, 3];
    let [cols, rows] = grid; if (o.landscape && (per === 6 || per === 2)) [cols, rows] = per === 6 ? [3, 2] : [2, 1];
    const lines = o.lines && per <= 3;
    const total = Math.ceil(list.length / per) || 1;
    for (let p = 0; p < total; p++) {
      const pg = document.createElement('div'); pg.className = 'ho-page';
      const items = list.slice(p * per, p * per + per);
      const cellW = cw / cols, cellH = ch / rows, gap = 6;
      items.forEach((x, k) => {
        const c = k % cols, r = Math.floor(k / cols);
        let sw = (lines && per === 3 ? cellW * .5 : cellW) - gap, sh = sw * ar; if (sh > cellH - gap) { sh = cellH - gap; sw = sh / ar; }
        const left = M + c * cellW + (lines && per === 3 ? 0 : (cellW - sw) / 2), tp = top + r * cellH + (cellH - sh) / 2 - (lines && per < 3 ? (cellH - sh) / 2 - gap / 2 : 0);
        const box = document.createElement('div'); box.className = 'ho-cell'; box.style.cssText = `left:${left}mm;top:${tp}mm;width:${sw}mm;height:${sh}mm`;
        box.appendChild(thumb(x.s, x.i, sw * MM)); pg.appendChild(box);
        if (o.numbers) pg.insertAdjacentHTML('beforeend', `<div class="ho-num" style="left:${left}mm;top:${tp + sh + .8}mm">${x.i + 1}</div>`);
        if (lines) {
          const lx = per === 3 ? left + sw + 8 : M, lw = per === 3 ? cw - sw - 8 : cw, ly0 = per === 3 ? tp + 2 : tp + sh + 8, ly1 = per === 3 ? tp + sh : top + (r + 1) * cellH - 4;
          let h = ''; for (let y = ly0 + 7; y <= ly1; y += 8) h += `<div class="ho-line" style="left:${lx}mm;top:${y}mm;width:${lw}mm"></div>`; pg.insertAdjacentHTML('beforeend', h);
        }
      });
      header(pg, p + 1, total, o); pages.push(pg);
    }
    return { pages, pw, ph };
  };
  X.cmd['export-handouts'] = async () => {
    let o0 = {}; try { o0 = JSON.parse(localStorage.getItem('bts-slides-handouts') || '{}'); } catch { }
    const o = await BTS.dialog({
      title: 'Export handouts (PDF)', width: 560,
      body: `<div class="ho-dlg"><div class="ho-opts">
        <div class="field"><label>Slides per page</label><div class="ho-per">${[1, 2, 3, 4, 6, 9].map(n => `<label class="ho-per-it" title="${n} slide${n > 1 ? 's' : ''} per page"><input type="radio" name="hop" value="${n}" ${(o0.per || 3) === n ? 'checked' : ''}><span>${n}</span></label>`).join('')}</div></div>
        <label class="check" title="Lines next to (3 per page) or under (1–2 per page) each slide for writing notes"><input type="checkbox" id="ho-lines" ${o0.lines !== false ? 'checked' : ''}> Lines for notes (1, 2 or 3 per page)</label>
        <div class="field-row"><div class="field"><label>Paper</label><select id="ho-paper"><option ${o0.paper !== 'Letter' ? 'selected' : ''}>A4</option><option ${o0.paper === 'Letter' ? 'selected' : ''}>Letter</option></select></div><div class="field"><label>Orientation</label><select id="ho-orient"><option value="p">Portrait</option><option value="l" ${o0.landscape ? 'selected' : ''}>Landscape</option></select></div></div>
        <div class="field"><label>Header</label><input id="ho-header" value="${esc(o0.header || BTS.title)}"></div>
        <label class="check"><input type="checkbox" id="ho-nums" ${o0.numbers !== false ? 'checked' : ''}> Slide numbers</label>
        <label class="check"><input type="checkbox" id="ho-pages" ${o0.pageNums !== false ? 'checked' : ''}> Page numbers</label>
        <label class="check"><input type="checkbox" id="ho-hidden" ${o0.hidden ? 'checked' : ''}> Include hidden slides</label>
        </div><div class="ho-prev" id="ho-prev"></div></div>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'Export…', primary: true, value: (r) => read(r) }],
      onOpen: (root) => { const draw = () => { const oo = read(root); const { pages, pw, ph } = X.buildHandoutPages(Object.assign({}, oo, { hidden: oo.hidden })); const pv = $('#ho-prev', root); const k = 170 / (pw * MM); pv.innerHTML = ''; const pg = pages[0]; pg.style.cssText = `width:${pw}mm;height:${ph}mm;transform:scale(${k});transform-origin:0 0;position:absolute;left:0;top:0`; const wrap = document.createElement('div'); wrap.style.cssText = `position:relative;width:${pw * MM * k}px;height:${ph * MM * k}px`; wrap.appendChild(pg); pv.appendChild(wrap); pv.insertAdjacentHTML('beforeend', `<div class="pf-note">${pages.length} page${pages.length > 1 ? 's' : ''}</div>`); }; draw(); root.addEventListener('change', draw); root.addEventListener('input', e => { if (e.target.id === 'ho-header') draw(); }); },
    });
    function read(root) { return { per: +$('input[name=hop]:checked', root).value, lines: $('#ho-lines', root).checked, paper: $('#ho-paper', root).value, landscape: $('#ho-orient', root).value === 'l', header: $('#ho-header', root).value, numbers: $('#ho-nums', root).checked, pageNums: $('#ho-pages', root).checked, hidden: $('#ho-hidden', root).checked }; }
    if (!o) return;
    try { localStorage.setItem('bts-slides-handouts', JSON.stringify(o)); } catch { }
    const data = await X.handoutsPDF(o);
    await saveBytes(BTS.title + ' - Handouts.pdf', data, 'PDF document', 'pdf');
  };
  X.handoutsPDF = async (o) => { await SL.ensureFonts(P()); const { pages, pw, ph } = X.buildHandoutPages(o); return printPages(pages, pw, ph); };

  /* ---------------- notes pages ---------------- */
  X.buildNotesPages = (o = {}) => {
    const [pw, ph] = PAPER[o.paper || 'A4']; const list = slideList(o.hidden); const ar = SL.H(P()) / SL.W(P());
    const sw = pw - 40, sh = sw * ar;
    return { pw, ph, pages: list.map((x, k) => {
      const pg = document.createElement('div'); pg.className = 'ho-page notes-pg';
      const box = document.createElement('div'); box.className = 'ho-cell'; box.style.cssText = `left:20mm;top:20mm;width:${sw}mm;height:${sh}mm`; box.appendChild(thumb(x.s, x.i, sw * MM)); pg.appendChild(box);
      const notes = (x.s.notes || '').trim();
      pg.insertAdjacentHTML('beforeend', `<div class="np-title" style="top:${20 + sh + 8}mm">Slide ${x.i + 1}${X.slideTitle ? ' · ' + esc(X.slideTitle(x.s, x.i)) : ''}</div><div class="np-text" style="top:${20 + sh + 16}mm;height:${ph - (20 + sh + 16) - 22}mm">${notes ? esc(notes).replace(/\n/g, '<br>') : '<span class="np-empty">No notes</span>'}</div><div class="ho-foot">${k + 1} / ${list.length}</div><div class="ho-head"><span>${esc(BTS.title)}</span><span></span></div>`);
      return pg;
    }) };
  };
  X.notesPDF = async (o) => { await SL.ensureFonts(P()); const { pages, pw, ph } = X.buildNotesPages(o); return printPages(pages, pw, ph); };
  X.cmd['export-notes'] = async () => {
    const o = await BTS.dialog({ title: 'Export notes pages (PDF)', width: 420, body: `<p class="pf-note" style="margin-top:0">Each slide is printed at the top of a page with its speaker notes below.</p><div class="field"><label>Paper</label><select id="np-paper"><option>A4</option><option>Letter</option></select></div><label class="check"><input type="checkbox" id="np-hidden"> Include hidden slides</label>`, buttons: [{ label: 'Cancel', value: null }, { label: 'Export…', primary: true, value: (r) => ({ paper: $('#np-paper', r).value, hidden: $('#np-hidden', r).checked }) }] });
    if (!o) return;
    const data = await X.notesPDF(o);
    await saveBytes(BTS.title + ' - Notes.pdf', data, 'PDF document', 'pdf');
  };

  /* ---------------- slides as PNG ---------------- */
  let pdfjs = null;
  async function pdfLib() {
    if (pdfjs) return pdfjs;
    pdfjs = await import(new URL('vendor/pdf.min.mjs', location.href).href);
    pdfjs.GlobalWorkerOptions.workerSrc = new URL('vendor/pdf.worker.min.mjs', location.href).href;
    return pdfjs;
  }
  X.slidesPDF = async (indices) => {
    const W = SL.W(P()), H = SL.H(P()); const wmm = W / MM, hmm = H / MM;
    const pages = indices.map(i => { const pg = document.createElement('div'); const node = SL.renderSlide(P().slides[i], P(), { index: i }); node.style.transformOrigin = '0 0'; pg.appendChild(node); return pg; });
    return printPages(pages, wmm, hmm);
  };
  X.renderPNGs = async (indices, width = 1920) => {
    const pdf = await X.slidesPDF(indices); if (!pdf) return null;
    const lib = await pdfLib(); const doc = await lib.getDocument({ data: pdf.slice(), isEvalSupported: false }).promise; const out = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p); const vp0 = page.getViewport({ scale: 1 }); const vp = page.getViewport({ scale: width / vp0.width });
      const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
      await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
      const blob = await new Promise(r => c.toBlob(r, 'image/png'));
      out.push(new Uint8Array(await blob.arrayBuffer()));
    }
    try { doc.destroy(); } catch { }
    return out;
  };
  X.cmd['export-png-current'] = async () => {
    BTS.toast('Rendering the slide…', { ms: 1200 });
    const pngs = await X.renderPNGs([SL.cur]); if (!pngs) return;
    await saveBytes(`${BTS.title} - Slide ${SL.cur + 1}.png`, pngs[0], 'PNG picture', 'png');
  };
  X.pngZip = async (hidden = false) => {
    const idx = slideList(hidden).map(x => x.i); const pngs = await X.renderPNGs(idx); if (!pngs) return null;
    const zip = new JSZip(); const pad = String(P().slides.length).length;
    pngs.forEach((d, k) => zip.file(`Slide ${String(idx[k] + 1).padStart(Math.max(2, pad), '0')}.png`, d));
    return zip.generateAsync({ type: 'uint8array' });
  };
  X.cmd['export-png-all'] = async () => {
    BTS.toast(`Rendering ${P().slides.filter(s => !s.hidden).length} slides…`, { ms: 1600 });
    const z = await X.pngZip(); if (!z) return;
    await saveBytes(`${BTS.title} - Slides.zip`, z, 'ZIP archive of PNG pictures', 'zip');
  };

  /* ---------------- outline & properties ---------------- */
  X.outlineText = () => P().slides.map((s, i) => {
    const t = s.elements.find(e => e.ph === 'title'); const lines = [`Slide ${i + 1}: ${t ? SL.plainText(t.html).trim() : ''}`];
    s.elements.filter(e => e !== t).forEach(e => { const txt = e.type === 'table' ? e.rows.map(r => r.map(c => SL.plainText(c.html).trim()).join('\t')).join('\n') : SL.plainText(e.html || '').trim(); if (txt) txt.split('\n').forEach(l => l.trim() && lines.push('  • ' + l.trim())); });
    if ((s.notes || '').trim()) lines.push('  Notes: ' + s.notes.trim().replace(/\n/g, '\n         '));
    return lines.join('\n');
  }).join('\n\n') + '\n';
  X.cmd['export-outline'] = () => saveBytes(BTS.title + ' - Outline.txt', new TextEncoder().encode(X.outlineText()), 'Text file', 'txt');
  X.cmd['doc-props'] = () => BTS.openBackstage('info');
  X.cmd['customize-colors'] = () => X.customizeColors();
})();
