/* BTS Docs 1.2 — Track changes.
 * Insertions are <ins class="tc-ins">, deletions <del class="tc-del"> (kept, struck through), each with
 * data-author / data-time / data-tcid. Typing, Backspace/Delete (char, word, line), cut, paste and inserted
 * symbols are tracked while D.design.track is on. Paragraph splits/merges and formatting are not tracked. */
(function () {
  const D = window.Docs; const { $, $$ } = D; const esc = BTS.esc;
  const PALETTE = ['#c7354a', '#2f6fd6', '#0f8a6a', '#b45309', '#7c3aed', '#be185d', '#0e7490', '#4d7c0f'];
  D.TC_PALETTE = PALETTE;
  const who = () => BTS.userName || (BTS.props && BTS.props.author) || 'Student';
  D.tcAuthor = who;
  D.tcColorIndex = (name) => { let h = 0; for (const c of String(name || '')) h = (h * 31 + c.codePointAt(0)) >>> 0; return h % PALETTE.length; };
  D.tracking = () => !!(D.design && D.design.track) && !D.isProtected?.();
  const isTC = (el, kind) => el && el.nodeType === 1 && el.classList.contains('tc-' + kind);
  const closestTC = (node, kind) => { let el = node && (node.nodeType === 1 ? node : node.parentElement); while (el && !D.isRoot(el)) { if (isTC(el, kind)) return el; el = el.parentElement; } return null; };
  D.closestTC = closestTC;
  function makeTC(kind, tcid, author, time) {
    const el = document.createElement(kind === 'ins' ? 'ins' : 'del');
    author = author || who();
    el.className = 'tc-' + kind + ' tc-c' + D.tcColorIndex(author);
    el.dataset.author = author; el.dataset.time = time || new Date().toISOString(); el.dataset.tcid = tcid || D.uid('t');
    el.title = (kind === 'ins' ? 'Inserted' : 'Deleted') + ' by ' + author + ' · ' + new Date(el.dataset.time).toLocaleString('en-IN', { day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit' });
    return el;
  }
  D.makeTC = makeTC;
  const mine = (el) => el && el.dataset.author === who();
  const atomic = (el) => el && el.nodeType === 1 && (el.tagName === 'IMG' || el.getAttribute('contenteditable') === 'false');

  /* ---------- tracked deletion of a range ---------- */
  /** Mark everything in the range as deleted (own insertions are really removed). Returns {first, last} del elements. */
  D.trackDeleteRange = (r) => {
    if (!r || r.collapsed) return {};
    const root = D.rootOf(r.commonAncestorContainer) || D.doc;
    const tcid = D.uid('t'); let first = null, last = null;
    const top = r.commonAncestorContainer.nodeType === 3 ? r.commonAncestorContainer.parentNode : r.commonAncestorContainer;
    // atomic inline items (pictures, citations, fields …) wholly inside the range
    const atoms = [];
    const w = document.createTreeWalker(top, NodeFilter.SHOW_ELEMENT, { acceptNode: (n) => atomic(n) ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP });
    let a; while ((a = w.nextNode())) {
      const outer = a.parentElement && a.parentElement.closest('[contenteditable="false"]'); if (outer && D.rootOf(outer)) continue;
      if (/^(DIV|TABLE|HR)$/.test(a.tagName)) continue;
      if (r.intersectsNode(a) && !closestTC(a, 'del')) atoms.push(a);
    }
    const segs = D.textNodesIn(r).filter(s => !(s.node.parentElement && s.node.parentElement.closest('[contenteditable="false"]')) && !closestTC(s.node, 'del'));
    const wrapDel = (node) => {
      const prev = node.previousSibling;
      if (isTC(prev, 'del') && mine(prev)) { prev.appendChild(node); return prev; }
      const d = makeTC('del', tcid); node.parentNode.insertBefore(d, node); d.appendChild(node);
      const next = d.nextSibling; if (isTC(next, 'del') && mine(next)) { while (next.firstChild) d.appendChild(next.firstChild); next.remove(); }
      return d;
    };
    const removed = [];
    segs.forEach(({ node, start, end }) => {
      let n = node; if (start > 0) n = n.splitText(start); if (end - start < n.length) n.splitText(end - start);
      if (!n.data.length) return;
      const ins = closestTC(n, 'ins');
      if (ins && mine(ins)) { removed.push(n); return; }
      if (!/\S| /.test(n.data) && n.parentNode && /^(UL|OL|TABLE|TBODY|TR)$/.test(n.parentNode.tagName)) return;
      const d = wrapDel(n); if (!first) first = d; last = d;
    });
    atoms.forEach(el => { if (!el.isConnected) return; const ins = closestTC(el, 'ins'); if (ins && mine(ins)) { removed.push(el); return; } const d = wrapDel(el); if (!first) first = d; last = d; });
    // own insertions are removed for real; remember a caret anchor
    let anchor = null;
    removed.forEach(n => { const ins = closestTC(n, 'ins'); anchor = { parent: n.parentNode, next: n.nextSibling }; n.remove(); if (ins && !ins.textContent && !ins.querySelector('img,[contenteditable="false"]')) { anchor = { parent: ins.parentNode, next: ins.nextSibling }; ins.remove(); } });
    root.normalize();
    return { first, last, anchor };
  };

  /* ---------- tracked insertion ---------- */
  function caretPoint() {
    const r = D.range(); if (!r) return null;
    let node = r.startContainer, off = r.startOffset;
    // never insert inside a deletion: move after it
    let del = closestTC(node, 'del');
    if (!del && node.nodeType === 1 && off > 0 && isTC(node.childNodes[off - 1], 'del') && !node.childNodes[off]) { /* after a del at end: fine */ }
    while (del) { node = del.parentNode; off = [...node.childNodes].indexOf(del) + 1; del = closestTC(node, 'del'); }
    return { node, off };
  }
  /** Insert plain text at the caret as a tracked insertion. */
  D.trackInsertText = (text) => {
    if (!text) return;
    let p = caretPoint(); if (!p) return;
    let { node, off } = p;
    const ins = closestTC(node, 'ins');
    if (node.nodeType === 3 && ins && mine(ins)) { node.insertData(off, text); D.setCaret(node, off + text.length); return; }
    if (node.nodeType === 3) {
      // split the text node and insert a new <ins> between (inside any formatting spans)
      const after = node.splitText(off); node = after.parentNode; off = [...node.childNodes].indexOf(after);
      if (!node.contains(after)) return;
    }
    if (node.nodeType === 1 && /^(UL|OL|TABLE|TBODY|TR)$/.test(node.tagName)) return;
    const before = node.childNodes[off - 1];
    if (isTC(before, 'ins') && mine(before)) {
      let t = before.lastChild; if (!t || t.nodeType !== 3) { t = document.createTextNode(''); before.appendChild(t); }
      const o = t.length; t.insertData(o, text); D.setCaret(t, o + text.length); return;
    }
    const el = makeTC('ins'); const t = document.createTextNode(text); el.appendChild(t);
    node.insertBefore(el, node.childNodes[off] || null);
    // drop a lone placeholder <br> in a now non-empty block
    const blk = D.blockOf(el); if (blk) { const brs = [...blk.childNodes].filter(c => c.nodeName === 'BR'); if (brs.length === 1 && blk.lastChild === brs[0] && blk.childNodes.length === 2) brs[0].remove(); }
    D.setCaret(t, t.length);
  };
  /** Live-range based capture of whatever an operation inserts (paste, symbols…). */
  D.trackBefore = () => {
    if (!D.tracking()) return null;
    const r = D.range(); if (!r || !D.doc.contains(r.startContainer)) return null;
    if (!r.collapsed) { const res = D.trackDeleteRange(r); placeAfter(res); }
    const p = caretPoint(); if (!p) return null;
    D.setCaret(p.node, p.off);
    const live = document.createRange(); live.setStart(p.node, p.off); live.collapse(true);
    return live;
  };
  D.trackAfter = (live) => {
    if (!live) return;
    const r = D.range(); if (!r) return;
    try { live.setEnd(r.endContainer, r.endOffset); } catch { return; }
    if (live.collapsed) return;
    const tcid = D.uid('t'); let lastText = null;
    const segs = D.textNodesIn(live).filter(s => !(s.node.parentElement && s.node.parentElement.closest('[contenteditable="false"]')));
    segs.forEach(({ node, start, end }) => {
      let n = node; if (start > 0) n = n.splitText(start); if (end - start < n.length) n.splitText(end - start);
      if (!n.data.length) return; lastText = n;
      const ins = closestTC(n, 'ins'); if (ins && mine(ins)) return;
      const prev = n.previousSibling; if (isTC(prev, 'ins') && mine(prev)) { prev.appendChild(n); return; }
      const el = makeTC('ins', tcid); n.parentNode.insertBefore(el, n); el.appendChild(n);
    });
    $$('img', live.commonAncestorContainer.nodeType === 1 ? live.commonAncestorContainer : live.commonAncestorContainer.parentNode).forEach(img => { if (live.intersectsNode(img) && !closestTC(img, 'ins')) { const el = makeTC('ins', tcid); img.before(el); el.appendChild(img); } });
    if (lastText && lastText.isConnected) D.setCaret(lastText, lastText.length);
  };
  function placeAfter(res) {
    if (res.last && res.last.isConnected) { const p = res.last.parentNode; D.setCaret(p, [...p.childNodes].indexOf(res.last) + 1); }
    else if (res.anchor && res.anchor.parent.isConnected) { const p = res.anchor.parent; const i = res.anchor.next && res.anchor.next.parentNode === p ? [...p.childNodes].indexOf(res.anchor.next) : p.childNodes.length; D.setCaret(p, i); }
  }
  function placeBefore(res) {
    if (res.first && res.first.isConnected) { const p = res.first.parentNode; D.setCaret(p, [...p.childNodes].indexOf(res.first)); }
    else placeAfter(res);
  }
  /** Delete the current selection (tracked when tracking is on). Used by cut. */
  D.deleteRange = (r) => { if (D.tracking() && D.doc.contains(r.commonAncestorContainer)) { const res = D.trackDeleteRange(r); placeBefore(res); } else r.deleteContents(); };

  /* ---------- keyboard / input interception ---------- */
  function extendFor(type) {
    const s = getSelection();
    const back = /Backward/.test(type);
    const unit = /Word/.test(type) ? 'word' : /Line|Hard/.test(type) ? 'lineboundary' : 'character';
    s.modify('extend', back ? 'backward' : 'forward', unit);
    return back;
  }
  function onBeforeInput(e) {
    if (!D.tracking() || D.loading || composing) return;
    const t = e.inputType || '';
    const r = D.range(); if (!r || !D.doc.contains(r.commonAncestorContainer)) return;
    if (t === 'insertText' || t === 'insertReplacementText' || t === 'insertFromDrop' || t === 'insertFromYank') {
      let text = e.data; if (text == null && e.dataTransfer) text = e.dataTransfer.getData('text/plain');
      if (t === 'insertReplacementText' && e.getTargetRanges && e.getTargetRanges()[0]) { const sr = e.getTargetRanges()[0]; const nr = document.createRange(); nr.setStart(sr.startContainer, sr.startOffset); nr.setEnd(sr.endContainer, sr.endOffset); const s = getSelection(); s.removeAllRanges(); s.addRange(nr); }
      if (!text) return;
      e.preventDefault();
      const rr = D.range(); if (rr && !rr.collapsed) { const res = D.trackDeleteRange(rr); placeAfter(res); }
      D.trackInsertText(text.replace(/\r?\n/g, ' '));
      D.afterChange();
      return;
    }
    if (t.startsWith('delete') && t !== 'deleteByDrag') {
      const rr = D.range();
      if (!rr.collapsed) { e.preventDefault(); const res = D.trackDeleteRange(rr); placeBefore(res); D.afterChange(); return; }
      // collapsed: extend over the unit to delete, skipping text that is already deleted
      const s = getSelection(); const anchor = { n: rr.startContainer, o: rr.startOffset };
      let back = true, guard = 0, ext = null;
      while (guard++ < 400) {
        back = extendFor(t); ext = s.rangeCount ? s.getRangeAt(0).cloneRange() : null;
        if (!ext || ext.collapsed) { ext = null; break; }
        const b1 = D.blockOf(ext.startContainer), b2 = D.blockOf(ext.endContainer);
        if (b1 !== b2) { ext = 'boundary'; break; }
        const segs = D.textNodesIn(ext); const imgs = ext.cloneContents().querySelector('img,[contenteditable="false"]');
        const live = segs.some(sg => !closestTC(sg.node, 'del') && sg.node.data.slice(sg.start, sg.end).length) || (imgs && ![...ext.cloneContents().querySelectorAll('img,[contenteditable="false"]')].every(x => x.closest('del')));
        if (live) break;
        // only deleted text in the way: jump over it and try again
        const c = document.createRange(); if (back) c.setStart(ext.startContainer, ext.startOffset); else c.setStart(ext.endContainer, ext.endOffset); c.collapse(true); s.removeAllRanges(); s.addRange(c);
        ext = null;
      }
      if (ext === 'boundary') { // paragraph merge: allowed, not tracked
        const c = document.createRange(); c.setStart(anchor.n, anchor.o); c.collapse(true); s.removeAllRanges(); s.addRange(c);
        return;
      }
      e.preventDefault();
      if (!ext) { D.afterChange(); return; }
      const res = D.trackDeleteRange(ext);
      if (back) placeBefore(res); else placeAfter(res);
      D.afterChange();
    }
  }
  // IME / composition can't be cancelled: wrap the composed text afterwards
  let composing = false, compLive = null;
  function onCompStart() { if (!D.tracking()) return; composing = true; const r = D.range(); if (r && !r.collapsed) { const res = D.trackDeleteRange(r); placeAfter(res); } const p = caretPoint(); if (p) { D.setCaret(p.node, p.off); compLive = document.createRange(); compLive.setStart(p.node, p.off); compLive.collapse(true); } }
  function onCompEnd() { if (!composing) return; composing = false; const l = compLive; compLive = null; if (l) setTimeout(() => { D.trackAfter(l); D.afterChange(); }, 0); }

  /* ---------- review: list, navigate, accept / reject ---------- */
  D.tcChanges = () => {
    const els = $$('ins.tc-ins, del.tc-del', D.doc); const map = new Map(); const out = [];
    els.forEach(el => { const id = el.dataset.tcid || D.uid('t'); el.dataset.tcid = id; const key = id + (el.tagName); let g = map.get(key); if (!g) { g = { id, kind: el.tagName === 'INS' ? 'ins' : 'del', author: el.dataset.author || 'Unknown', time: el.dataset.time, els: [] }; map.set(key, g); out.push(g); } g.els.push(el); });
    return out;
  };
  function changeAtCaret(dir = 0) {
    const all = D.tcChanges(); if (!all.length) return null;
    const r = D.range() || D.lastRange();
    if (r) { const at = closestTC(r.startContainer, 'ins') || closestTC(r.startContainer, 'del'); if (at && dir === 0) return all.find(g => g.els.includes(at)); }
    if (!r) return dir < 0 ? all[all.length - 1] : all[0];
    const rangeOf = (n) => { const x = document.createRange(); x.selectNode(n); return x; };
    if (dir >= 0) { const after = all.find(g => rangeOf(g.els[0]).compareBoundaryPoints(Range.START_TO_END, r) > 0 && !(r.collapsed === false && g.els.some(e => r.intersectsNode(e)) && dir > 0)); return after || all[0]; }
    const before = [...all].reverse().find(g => rangeOf(g.els[0]).compareBoundaryPoints(Range.END_TO_START, r) < 0 || rangeOf(g.els[0]).compareBoundaryPoints(Range.START_TO_START, r) < 0); return before || all[all.length - 1];
  }
  function selectChange(g) {
    if (!g) return;
    const r = document.createRange(); r.setStartBefore(g.els[0]); r.setEndAfter(g.els[g.els.length - 1]);
    D.doc.focus({ preventScroll: true }); const s = getSelection(); s.removeAllRanges(); s.addRange(r);
    D.scrollToEl(g.els[0]);
    BTS.setStatus('tcinfo', `<span title="Change">${g.kind === 'ins' ? 'Inserted' : 'Deleted'} by ${esc(g.author)}${g.time ? ' · ' + esc(new Date(g.time).toLocaleString('en-IN', { day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit' })) : ''}</span>`);
  }
  D.tcStep = (dir) => {
    const all = D.tcChanges(); if (!all.length) { BTS.toast('There are no tracked changes in this document.'); return; }
    const r = D.range(); let g;
    if (r && !r.collapsed) {
      const cur = all.findIndex(x => x.els.some(e => r.intersectsNode(e)));
      g = cur >= 0 ? all[(cur + dir + all.length) % all.length] : changeAtCaret(dir);
    } else g = changeAtCaret(dir || 1);
    selectChange(g);
  };
  function resolve(g, accept) {
    g.els.forEach(el => {
      if (!el.isConnected) return;
      const keep = (g.kind === 'ins') === accept;
      if (keep) D.unwrap(el); else { const blk = D.blockOf(el); el.remove(); if (blk && !blk.textContent && !blk.querySelector('img,br,[contenteditable="false"]')) blk.innerHTML = '<br>'; }
    });
  }
  D.tcResolve = (accept, scope) => {
    const all = D.tcChanges(); if (!all.length) { BTS.toast('There are no tracked changes in this document.'); return; }
    if (scope === 'all') { D.change(() => all.forEach(g => resolve(g, accept))); D.doc.normalize(); BTS.toast(`${accept ? 'Accepted' : 'Rejected'} all ${all.length} change${all.length === 1 ? '' : 's'}`); refreshStatus(); return; }
    const r = D.range();
    let targets = [];
    if (r && !r.collapsed) targets = all.filter(g => g.els.some(e => r.intersectsNode(e)));
    if (!targets.length) { const g = changeAtCaret(0); if (g) targets = [g]; }
    if (!targets.length) return;
    const lastEl = targets[targets.length - 1].els.slice(-1)[0];
    const nextAfter = all[all.indexOf(targets[targets.length - 1]) + 1];
    D.change(() => targets.forEach(g => resolve(g, accept)));
    if (scope === 'next' && nextAfter && nextAfter.els[0].isConnected) selectChange(D.tcChanges().find(x => x.id === nextAfter.id) || nextAfter);
    else if (lastEl && !lastEl.isConnected) { /* caret stays */ }
    refreshStatus();
  };
  D.setTrack = (on) => {
    D.design.track = on ?? !D.design.track; BTS.markDirty();
    syncTrackUI(); BTS.toast(D.design.track ? 'Track changes is on — your edits are marked' : 'Track changes is off');
  };
  D.setMarkup = (mode) => { D.design.markup = mode; syncTrackUI(); BTS.markDirty(); };
  function syncTrackUI() {
    const b = $('#btn-track'); if (b) b.classList.toggle('on', !!D.design.track);
    document.body.classList.toggle('tc-simple', D.design.markup === 'simple');
    document.body.classList.toggle('tc-original', D.design.markup === 'original');
    refreshStatus();
  }
  D.syncTrackUI = syncTrackUI;
  function refreshStatus() {
    const n = $$('ins.tc-ins, del.tc-del', D.doc).length;
    BTS.setStatus('track', D.design.track ? `<button class="sb-btn on" data-cmd="track-toggle" title="Track changes is on — click to turn off">${icon('file-pen-line', 14)} Tracking</button>` : (n ? `<button class="sb-btn" data-cmd="tc-next" title="This document has tracked changes">${icon('file-pen-line', 14)} Changes</button>` : ''));
    if (!n) BTS.setStatus('tcinfo', '');
  }
  D.markupMenu = (anchor) => BTS.menu(anchor, [{ head: 'Display for review' }, { label: 'Simple markup (final, with change bars)', checked: D.design.markup === 'simple', action: () => D.setMarkup('simple') }, { label: 'All markup', checked: (D.design.markup || 'all') === 'all', action: () => D.setMarkup('all') }, { label: 'Original (before changes)', checked: D.design.markup === 'original', action: () => D.setMarkup('original') }]);
  D.acceptMenu = (anchor, accept) => BTS.menu(anchor, accept
    ? [{ label: 'Accept and move to next', icon: 'check', action: () => D.tcResolve(true, 'next') }, { label: 'Accept this change', icon: 'check', action: () => D.tcResolve(true) }, { label: 'Accept all changes', icon: 'check-check', action: () => D.tcResolve(true, 'all') }, { label: 'Accept all changes and stop tracking', icon: 'check-check', action: () => { D.tcResolve(true, 'all'); if (D.design.track) D.setTrack(false); } }]
    : [{ label: 'Reject and move to next', icon: 'x', action: () => D.tcResolve(false, 'next') }, { label: 'Reject this change', icon: 'x', action: () => D.tcResolve(false) }, { label: 'Reject all changes', icon: 'circle-x', action: () => D.tcResolve(false, 'all') }, { label: 'Reject all changes and stop tracking', icon: 'circle-x', action: () => { D.tcResolve(false, 'all'); if (D.design.track) D.setTrack(false); } }]);

  D.initTrack = () => {
    D.doc.addEventListener('beforeinput', onBeforeInput);
    D.doc.addEventListener('compositionstart', onCompStart);
    D.doc.addEventListener('compositionend', onCompEnd);
    syncTrackUI();
  };
  D.refreshTrackStatus = refreshStatus;
})();
