/* BTS Docs — find & replace pane, navigation pane, word count, read aloud, format painter,
 * print/PDF layout builder, autosave & recovery. */
(function () {
  const D = window.Docs; const { $, $$, PX } = D; const esc = BTS.esc;
  const HL = typeof Highlight !== 'undefined' && CSS.highlights;

  /* ================= find & replace ================= */
  const F = { hits: [], cur: -1 };
  function findOpts() { return { q: $('#find-input').value, mc: $('#find-case').checked, ww: $('#find-word').checked }; }
  function textMap(root) {
    // Concatenate text nodes per block so matches don't span blocks.
    const nodes = []; const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, { acceptNode: n => n.parentElement.closest('.toc, .eq') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT }); let n;
    while ((n = w.nextNode())) nodes.push(n);
    const groups = []; let g = null, lastBlock = null;
    nodes.forEach(n => { const b = D.blockOf(n) || n.parentElement; if (b !== lastBlock) { g = { text: '', parts: [] }; groups.push(g); lastBlock = b; } g.parts.push({ n, start: g.text.length }); g.text += n.data; });
    return groups;
  }
  function locate(g, idx) { let p = g.parts[0]; for (const x of g.parts) { if (x.start <= idx) p = x; else break; } return { node: p.n, offset: idx - p.start }; }
  function buildRegex({ q, mc, ww }) {
    const e = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    return new RegExp(ww ? `(?<![\\p{L}\\p{N}_])${e}(?![\\p{L}\\p{N}_])` : e, 'gu' + (mc ? '' : 'i'));
  }
  D.runFind = (keepIndex) => {
    const o = findOpts(); F.hits = [];
    if (o.q) {
      const re = buildRegex(o);
      [D.doc, D.hdr, D.ftr].forEach(root => textMap(root).forEach(g => {
        let m; re.lastIndex = 0;
        while ((m = re.exec(g.text))) { if (!m[0].length) { re.lastIndex++; continue; } const a = locate(g, m.index), b = locate(g, m.index + m[0].length); const r = document.createRange(); try { r.setStart(a.node, a.offset); r.setEnd(b.node, b.offset); } catch { continue; } F.hits.push({ r, ctx: g.text, at: m.index, len: m[0].length }); }
      }));
    }
    if (!keepIndex) {
      // choose first hit after caret
      const cr = D.range() || D.lastRange(); F.cur = F.hits.length ? 0 : -1;
      if (cr && F.hits.length) { const i = F.hits.findIndex(h => h.r.compareBoundaryPoints(Range.START_TO_START, cr) >= 0); F.cur = i >= 0 ? i : 0; }
    } else F.cur = Math.min(F.cur, F.hits.length - 1);
    paintHits(); renderResults();
  };
  function paintHits() {
    if (!HL) return;
    CSS.highlights.delete('find-hit'); CSS.highlights.delete('find-cur');
    if (!F.hits.length || $('#find-pane').classList.contains('hidden')) return;
    CSS.highlights.set('find-hit', new Highlight(...F.hits.map(h => h.r)));
    if (F.cur >= 0) CSS.highlights.set('find-cur', new Highlight(F.hits[F.cur].r));
  }
  function renderResults() {
    const o = findOpts();
    $('#find-count').textContent = o.q ? (F.hits.length ? `${F.cur + 1} of ${F.hits.length}` : 'No results') : '';
    $('#find-results').innerHTML = F.hits.slice(0, 200).map((h, i) => { const a = Math.max(0, h.at - 30); return `<button class="fr ${i === F.cur ? 'on' : ''}" data-i="${i}">${a > 0 ? '…' : ''}${esc(h.ctx.slice(a, h.at))}<b>${esc(h.ctx.substr(h.at, h.len))}</b>${esc(h.ctx.slice(h.at + h.len, h.at + h.len + 50))}</button>`; }).join('');
  }
  function gotoHit(i, focusDoc) {
    if (!F.hits.length) return; F.cur = (i + F.hits.length) % F.hits.length;
    const h = F.hits[F.cur]; const s = getSelection(); s.removeAllRanges(); s.addRange(h.r.cloneRange());
    const el = h.r.startContainer.parentElement; const c = D.canvas.getBoundingClientRect(); const rr = h.r.getBoundingClientRect();
    if (rr.top < c.top + 40 || rr.bottom > c.bottom - 40) D.canvas.scrollTop += rr.top - c.top - c.height / 3;
    if (focusDoc) (D.rootOf(el) || D.doc).focus({ preventScroll: true });
    paintHits(); renderResults();
    const on = $('#find-results .fr.on'); if (on) on.scrollIntoView({ block: 'nearest' });
  }
  D.openFind = (replace, text) => {
    const p = $('#find-pane'); p.classList.remove('hidden'); p.classList.toggle('replace-mode', !!replace);
    $('#find-title').textContent = replace ? 'Find and replace' : 'Find'; $('#find-toggle-replace').textContent = replace ? 'Hide replace' : 'Show replace';
    const sel = D.selectedText(); const inp = $('#find-input');
    if (text != null) inp.value = text; else if (sel && sel.length < 80 && !sel.includes('\n')) inp.value = sel;
    inp.focus(); inp.select(); D.runFind(); if (F.hits.length) gotoHit(F.cur);
    D.scheduleLayout(30);
  };
  D.closeFind = () => { $('#find-pane').classList.add('hidden'); paintHits(); D.doc.focus({ preventScroll: true }); D.scheduleLayout(30); };
  D.findStep = (dir) => { if (!F.hits.length) D.runFind(); if (F.hits.length) gotoHit(F.cur + dir); };
  D.replaceOne = () => {
    if (!F.hits.length) { D.runFind(); return; }
    const h = F.hits[F.cur]; const rep = $('#replace-input').value;
    D.change(() => { const s = getSelection(); s.removeAllRanges(); s.addRange(h.r); (D.rootOf(h.r.startContainer) || D.doc).focus({ preventScroll: true }); document.execCommand('insertText', false, rep); });
    D.runFind(true); if (F.hits.length) gotoHit(F.cur);
  };
  D.replaceAll = (q, rep, opts = {}) => {
    if (q != null) { $('#find-input').value = q; $('#replace-input').value = rep; $('#find-case').checked = !!opts.mc; $('#find-word').checked = !!opts.ww; }
    D.runFind(); const n = F.hits.length; if (!n) { BTS.toast('No matches found'); return 0; }
    const r2 = $('#replace-input').value;
    D.change(() => { [...F.hits].reverse().forEach(h => { h.r.deleteContents(); if (r2) h.r.insertNode(document.createTextNode(r2)); }); [D.doc, D.hdr, D.ftr].forEach(r => r.normalize()); });
    D.runFind(); BTS.toast(`Replaced ${n} occurrence${n === 1 ? '' : 's'}`); return n;
  };
  function wireFind() {
    const inp = $('#find-input'); let t = 0;
    inp.addEventListener('input', () => { clearTimeout(t); t = setTimeout(() => { D.runFind(); if (F.hits.length) gotoHit(F.cur); }, 120); });
    inp.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); D.findStep(e.shiftKey ? -1 : 1); } if (e.key === 'Escape') D.closeFind(); });
    $('#replace-input').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); D.replaceOne(); } if (e.key === 'Escape') D.closeFind(); });
    ['#find-case', '#find-word'].forEach(s => $(s).addEventListener('change', () => D.runFind()));
    $('#find-results').addEventListener('click', e => { const b = e.target.closest('.fr'); if (b) gotoHit(+b.dataset.i, true); });
  }
  D.refreshFind = () => { if (!$('#find-pane').classList.contains('hidden') && $('#find-input').value) D.runFind(true); };

  /* ================= navigation pane ================= */
  D.toggleNav = (on) => {
    const p = $('#nav-pane'); const show = on ?? p.classList.contains('hidden');
    p.classList.toggle('hidden', !show); $('#chk-nav').checked = show; if (show) D.renderNav(); D.scheduleLayout(30);
  };
  D.renderNav = () => {
    const p = $('#nav-pane'); if (p.classList.contains('hidden')) return;
    const heads = $$('h1,h2,h3', D.doc).filter(h => h.textContent.trim() && !h.closest('.toc'));
    $('#nav-list').innerHTML = heads.length ? heads.map((h, i) => `<button class="nv l${h.tagName[1]}" data-i="${i}" title="${esc(h.textContent.trim())}">${esc(h.textContent.trim())}</button>`).join('') : '<div class="empty-note">Create an interactive outline of your document.<br><br>Use the Heading styles on the Home tab to give your titles structure — they’ll appear here.</div>';
    D._navHeads = heads; markNav();
  };
  function markNav() {
    const heads = D._navHeads || []; const r = D.range(); if (!r) return; let cur = -1;
    heads.forEach((h, i) => { const hr = document.createRange(); hr.selectNode(h); if (hr.compareBoundaryPoints(Range.START_TO_START, r) <= 0) cur = i; });
    $$('#nav-list .nv').forEach((b, i) => b.classList.toggle('on', i === cur));
  }
  D.markNav = markNav;
  function wireNav() { $('#nav-list').addEventListener('click', e => { const b = e.target.closest('.nv'); if (b) D.scrollToEl(D._navHeads[+b.dataset.i], true); }); }

  /* ================= word count ================= */
  D.stats = () => {
    const text = D.doc.innerText.replace(/​/g, '') + '\n' + ($('#fn-area').innerText || '');
    const paras = $$('p,h1,h2,h3,h4,h5,h6,li,pre,blockquote', D.doc).filter(p => !D.hasBlockChild(p) && p.textContent.trim()).length;
    return { pages: D.pageCount, words: D.countWords(text), chars: text.replace(/\n/g, '').length, charsNoSpace: text.replace(/\s/g, '').length, paras, lines: text.split('\n').filter(l => l.trim()).length };
  };
  D.wordCountDialog = () => {
    const s = D.stats(); const sel = D.selectedText();
    const row = (k, v) => `<tr><td>${k}</td><td>${v.toLocaleString('en-IN')}</td></tr>`;
    BTS.dialog({ title: 'Word count', body: `<table class="wc-table">${sel ? row('Words in selection', D.countWords(sel)) : ''}${row('Pages', s.pages)}${row('Words', s.words)}${row('Characters (no spaces)', s.charsNoSpace)}${row('Characters (with spaces)', s.chars)}${row('Paragraphs', s.paras)}${row('Lines', s.lines)}</table>`, buttons: [{ label: 'Close', primary: true, value: true }] });
  };

  /* ================= read aloud ================= */
  let speaking = false;
  D.readAloud = () => {
    if (!('speechSynthesis' in window)) { BTS.toast('Read aloud isn’t available on this computer.', { error: true }); return; }
    if (speaking) { stopReading(); return; }
    let text = D.selectedText(); let base = null;
    const r = D.range();
    if (!text.trim()) { text = D.doc.innerText; base = D.doc; }
    if (!text.trim()) { BTS.toast('There is nothing to read.'); return; }
    const u = new SpeechSynthesisUtterance(text.slice(0, 32000)); u.rate = 1; u.lang = /[ऀ-ॿ]/.test(text) ? 'hi-IN' : 'en-IN';
    const v = speechSynthesis.getVoices().find(v => v.lang === u.lang) || speechSynthesis.getVoices().find(v => /^en/.test(v.lang)); if (v) u.voice = v;
    u.onend = u.onerror = () => stopReading();
    speaking = true; $('#btn-read').classList.add('on'); $('#btn-read span').textContent = 'Stop reading';
    speechSynthesis.cancel(); speechSynthesis.speak(u);
    if (!speechSynthesis.getVoices().length) setTimeout(() => { if (speaking && !speechSynthesis.speaking) { /* no TTS engine */ } }, 1500);
  };
  function stopReading() { speaking = false; try { speechSynthesis.cancel(); } catch { } $('#btn-read').classList.remove('on'); $('#btn-read span').textContent = 'Read aloud'; if (HL) CSS.highlights.delete('read-word'); }
  D.stopReading = stopReading;

  /* ================= format painter ================= */
  const P = { fmt: null, sticky: false };
  D.formatPainter = (sticky) => {
    if (P.fmt && !sticky) { endPainter(); return; }
    const r = D.ensureSel(); let el = r.startContainer; if (el.nodeType === 3) el = el.parentElement;
    const cs = getComputedStyle(el); const block = D.blockOf(r.startContainer);
    P.fmt = {
      inline: { fontFamily: cs.fontFamily, fontSize: cs.fontSize, fontWeight: cs.fontWeight, fontStyle: cs.fontStyle, textDecorationLine: cs.textDecorationLine, color: cs.color, backgroundColor: inlineBg(el) },
      style: D.styleOf(block), block: block ? { textAlign: block.style.textAlign, lineHeight: block.style.lineHeight, marginLeft: block.style.marginLeft } : null,
    };
    P.sticky = sticky; $('#btn-painter').classList.add('on'); D.doc.style.cursor = 'copy';
    BTS.toast(sticky ? 'Format painter on — select text to apply. Press Esc to stop.' : 'Select the text to apply this formatting.');
  };
  function inlineBg(el) { while (el && !D.isRoot(el) && !D.BLOCK_TAGS.has(el.tagName)) { const b = el.style.backgroundColor; if (b) return b; el = el.parentElement; } return ''; }
  function endPainter() { P.fmt = null; $('#btn-painter').classList.remove('on'); D.doc.style.cursor = ''; }
  D.endPainter = endPainter;
  D.applyPainter = () => {
    if (!P.fmt) return false;
    const r = D.range(); if (!r) return false;
    if (r.collapsed) D.selectWordIfCollapsed();
    const rr = D.range(); if (!rr || rr.collapsed) return false;
    const f = P.fmt;
    D.change(() => {
      const blocks = D.blocks(rr);
      // paragraph style when whole paragraph(s) selected
      if (f.style && f.style !== 'normal' && blocks.length > 1) blocks.forEach(b => D._setBlockStyle(b, f.style));
      const r2 = D.range() || rr; D.exec('removeFormat');
      const r3 = D.range() || r2;
      D.wrapText(r3, () => { const s = document.createElement('span'); const i = f.inline; s.style.fontFamily = i.fontFamily; s.style.fontSize = (parseFloat(i.fontSize) * 0.75).toFixed(1) + 'pt'; if (+i.fontWeight >= 600) s.style.fontWeight = 'bold'; if (i.fontStyle === 'italic') s.style.fontStyle = 'italic'; if (i.textDecorationLine !== 'none') s.style.textDecoration = i.textDecorationLine; s.style.color = i.color; if (i.backgroundColor) s.style.backgroundColor = i.backgroundColor; return s; });
      if (f.block) blocks.forEach(b => { if (f.block.textAlign !== undefined) b.style.textAlign = f.block.textAlign; if (f.block.lineHeight) b.style.lineHeight = f.block.lineHeight; });
    });
    if (!P.sticky) endPainter();
    return true;
  };

  /* ================= print / PDF layout ================= */
  /** Build #print-root: a table with thead/tfoot spacers (= margins on every printed page) +
   *  absolutely positioned header/footer/watermark/border per page. */
  D.buildPrint = () => {
    D.layout();
    const ps = D.paperSize(); const eff = D.eff; const W = ps.w * PX, H = ps.h * PX;
    const root = $('#print-root');
    const N = D.pageCount;
    let css = `@page { size: ${ps.w}in ${ps.h}in; margin: 0; }
      #print-root { width: ${ps.w}in; background: ${D.design.pageColor || 'transparent'}; }
      body.printing { background: ${D.design.pageColor || '#fff'} !important; }`;
    const docClone = D.doc.cloneNode(true);
    docClone.removeAttribute('contenteditable'); docClone.removeAttribute('id'); docClone.removeAttribute('role');
    $$('[contenteditable]', docClone).forEach(e => e.removeAttribute('contenteditable'));
    $$('.cell-sel', docClone).forEach(e => e.classList.remove('cell-sel'));
    $$('hr.page-break', docClone).forEach(e => e.removeAttribute('style'));
    $$('.cmt.active', docClone).forEach(e => e.classList.remove('active'));
    const fn = $('#fn-area').cloneNode(true); fn.removeAttribute('id');
    const styleVars = D.doc.getAttribute('style') || '';
    let deco = '';
    const hdrHas = D.hdr.textContent.trim() || D.hdr.querySelector('img'), ftrHas = D.ftr.textContent.trim() || D.ftr.querySelector('img');
    const m = D.page.m;
    for (let i = 0; i < N; i++) {
      const top = i * H;
      let inner = '';
      if (D.design.watermark) inner += `<div class="print-wm" style="inset:0">${D.watermarkInner(D.design.watermark, W)}</div>`;
      if (D.design.border) inner += `<div class="print-pborder" style="${D.borderCSS(D.design.border)};top:${0.35 * PX}px;bottom:${0.35 * PX}px;left:${0.35 * PX}px;right:${0.35 * PX}px"></div>`;
      const hfStyle = `class="p-hf doc-content" style="${esc(styleVars)};left:${m.l * PX}px;right:${m.r * PX}px;`;
      if (hdrHas && !(D.design.hdrFirst && i === 0)) inner += `<div ${hfStyle}top:${D.page.hdrDist * PX}px">${D.hfHTML(D.hdr, i + 1, N)}</div>`;
      if (ftrHas && !(D.design.hdrFirst && i === 0)) inner += `<div ${hfStyle}bottom:${D.page.ftrDist * PX}px">${D.hfHTML(D.ftr, i + 1, N)}</div>`;
      if (inner) deco += `<div class="p-deco" style="top:${top}px;height:${H}px">${inner}</div>`;
    }
    // One table per manual page-break segment: thead/tfoot spacers repeat the top/bottom margin on every
    // printed page (Chromium repeats table header/footer groups across page fragments).
    const segs = [[]]; [...docClone.childNodes].forEach(n => { if (n.nodeType === 1 && n.matches('hr.page-break')) segs.push([]); else segs[segs.length - 1].push(n); });
    const tbl = (i) => `<table class="print-tbl" style="${i ? 'break-before:page' : ''}"><thead><tr><td style="height:${eff.t}px"></td></tr></thead><tfoot><tr><td style="height:${eff.b}px"></td></tr></tfoot><tbody><tr><td class="print-cell" style="padding:0 ${eff.r}px 0 ${eff.l}px"></td></tr></tbody></table>`;
    root.innerHTML = `<style>${css}</style>${deco}` + segs.map((_, i) => tbl(i)).join('');
    const cells = $$('.print-cell', root);
    segs.forEach((nodes, i) => { const d = document.createElement('div'); d.className = 'doc-content'; d.setAttribute('style', styleVars); nodes.forEach(n => d.appendChild(n)); if (!nodes.length) d.innerHTML = '<p><br></p>'; cells[i].appendChild(d); });
    if (fn.firstChild) { fn.setAttribute('style', styleVars); fn.className = 'doc-content'; cells[cells.length - 1].appendChild(fn); }
    document.body.classList.add('printing');
    return { pageSize: ps.w === 8.27 ? 'A4' : { width: ps.w, height: ps.h }, landscape: false };
  };
  D.afterPrint = () => { document.body.classList.remove('printing'); $('#print-root').innerHTML = ''; };

  /* ================= autosave & recovery =================
   * Each window writes its own slot (bts-docs-autosave:<sid>) plus a heartbeat. On startup we only offer
   * slots whose window stopped beating (crashed / closed without saving), never a live window's work. */
  const AK = 'bts-docs-autosave:'; const SID = Math.random().toString(36).slice(2, 10);
  const ls = { get: (k) => { try { return localStorage.getItem(k); } catch { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); return true; } catch { return false; } }, del: (k) => { try { localStorage.removeItem(k); } catch { } }, keys: () => { try { return Object.keys(localStorage); } catch { return []; } } };
  let autoTimer = 0;
  D.autosave = () => {
    clearTimeout(autoTimer);
    autoTimer = setTimeout(() => {
      if (!BTS.dirty) return;
      const rec = { t: Date.now(), hb: Date.now(), title: BTS.title, html: D.serializeNative() };
      if (!ls.set(AK + SID, JSON.stringify(rec))) { rec.html = rec.html.replace(/src="data:[^"]{20000,}"/g, 'src=""'); ls.set(AK + SID, JSON.stringify(rec)); }
    }, 1500);
  };
  setInterval(() => { const raw = ls.get(AK + SID); if (!raw) return; try { const r = JSON.parse(raw); r.hb = Date.now(); ls.set(AK + SID, JSON.stringify(r)); } catch { } }, 5000);
  D.clearAutosave = () => ls.del(AK + SID);
  D.checkRecovery = () => {
    const now = Date.now(); let best = null, bestKey = null; ls.del('bts-docs-autosave');
    ls.keys().filter(k => k.startsWith(AK) && k !== AK + SID).forEach(k => {
      let r = null; try { r = JSON.parse(ls.get(k)); } catch { }
      if (!r || !r.html || now - r.t > 14 * 864e5) { if (r && now - r.t > 14 * 864e5) ls.del(k); return; }
      if (now - (r.hb || r.t) < 15000) return; // that window is still open
      if (!best || r.t > best.t) { best = r; bestKey = k; }
    });
    if (!best) return;
    const bar = $('#recover-bar');
    bar.innerHTML = `${icon('history', 16)}<span>Recover unsaved document “${esc(best.title || 'Document')}” from ${esc(new Date(best.t).toLocaleString('en-IN', { day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit' }))}?</span><span style="flex:1"></span><button class="btn primary" id="rec-yes">Recover</button><button class="btn" id="rec-no">Dismiss</button>`;
    bar.classList.remove('hidden');
    $('#rec-yes').onclick = () => { ls.del(bestKey); D.loadNative(best.html); BTS.setTitle(best.title || 'Document'); BTS.filePath = null; BTS.markDirty(); D.autosave(); bar.classList.add('hidden'); BTS.toast('Document recovered'); };
    $('#rec-no').onclick = () => { ls.del(bestKey); bar.classList.add('hidden'); D.scheduleLayout(10); };
  };

  D.initPanes = () => { wireFind(); wireNav(); };
})();
