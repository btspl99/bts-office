/* BTS Slides — editor: canvas, selection, move/resize/rotate with smart guides, text editing,
 * slide rail (thumbnails, drag to reorder), slide sorter, notes, zoom and view state.
 * Everything lives on window.E (editor state) and SL (model helpers).
 */
(function () {
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const f2 = (n) => Math.round(n * 100) / 100;

  const E = window.E = {
    zoom: 1, fit: true, sel: [], editing: null, view: 'normal', grid: false, guides: false, ruler: false, snap: true,
    notesOn: true, railSel: new Set(), clip: null, slideClip: null, drawMode: null, cropping: null, tableCell: null,
  };
  const P = () => SL.pres;
  const S = () => SL.slide();
  E.selected = () => E.sel.map(id => SL.findEl(id)).filter(Boolean);
  E.primary = () => E.selected()[0] || null;

  /* ====================================================================== */
  /*  Rendering                                                              */
  /* ====================================================================== */
  E.renderAll = () => {
    if (SL.cur >= P().slides.length) SL.cur = P().slides.length - 1;
    if (SL.cur < 0) SL.cur = 0;
    E.sel = E.sel.filter(id => SL.findEl(id));
    E.railSel = new Set([...E.railSel].filter(i => i < P().slides.length)); if (!E.railSel.size) E.railSel.add(SL.cur);
    E.renderRail(); E.renderCanvas(); E.renderNotes(); E.updateStatus(); E.updateRibbon();
    if (E.view === 'sorter') E.renderSorter();
    if (E.view === 'notes') E.renderNotesView();
    E.refreshPane && E.refreshPane();
  };

  E.layout = () => {
    const host = $('#slide-host'), st = $('#stage'); if (!host) return;
    const W = SL.W(P()), H = SL.H(P());
    if (!st.clientWidth || !st.clientHeight) return; // stage hidden (sorter / notes view)
    if (E.fit) {
      const aw = st.clientWidth - 56, ah = st.clientHeight - 56;
      const z = Math.max(.1, Math.min(aw / W, ah / H));
      E.zoom = z; E._fitting = true; BTS.setZoom(Math.round(z * 100)); E._fitting = false; E.zoom = z;
    }
    host.style.width = f2(W * E.zoom) + 'px'; host.style.height = f2(H * E.zoom) + 'px';
    host.style.setProperty('--z', E.zoom);
    const sl = $('.sl-slide', host); if (sl) sl.style.transform = `scale(${E.zoom})`;
    E.renderOverlay(); E.renderRulers();
  };

  E.renderCanvas = () => {
    const host = $('#slide-host'); if (!host) return;
    const slide = S(); host.innerHTML = '';
    const sl = SL.renderSlide(slide, P(), { editor: true, editingId: E.editing, index: SL.cur });
    sl.style.transformOrigin = '0 0';
    host.appendChild(sl);
    const ov = document.createElement('div'); ov.id = 'overlay'; host.appendChild(ov);
    if (E.grid) host.classList.add('show-grid'); else host.classList.remove('show-grid');
    $('#stage').classList.toggle('hidden-slide', !!slide.hidden);
    E.fitAllText(sl);
    E.layout();
    if (E.editing) E.resumeEdit();
  };

  /* Text autofit (resize shape to fit text / shrink text on overflow) + table height sync. */
  E.fitAllText = (root) => {
    for (const el of S().elements) {
      const node = root.querySelector(`.el[data-id="${el.id}"]`); if (!node) continue;
      if (el.type === 'table') { const t = node.querySelector('table'); if (t && t.offsetHeight > el.h + 1) el.h = t.offsetHeight; continue; }
      if (el.type !== 'text' && el.type !== 'shape') continue;
      E.fitText(el, node);
    }
  };
  E.fitText = (el, node) => {
    const tx = node.querySelector('.tx'), tc = node.querySelector('.tc'); if (!tx || !tc) return false;
    if (el.autofit === 'resize' && !el.rot && el.type === 'text' && (!el.dir || el.dir === 'horz')) {
      const cs = getComputedStyle(tx); const need = Math.ceil(tc.offsetHeight + parseFloat(cs.paddingTop) + parseFloat(cs.paddingBottom));
      if (Math.abs(need - el.h) > 1) { el.h = need; node.style.height = need + 'px'; return true; }
    } else if (el.autofit === 'shrink' && !SL.isEmptyText(el)) {
      const cs = getComputedStyle(tx); const avail = el.h - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom);
      tc.style.zoom = 1;
      if (tc.offsetHeight <= avail + 1) { if (el.fontScale && el.fontScale !== 1) el.fontScale = 1; return false; }
      let lo = .25, hi = 1;
      for (let i = 0; i < 7; i++) { const m = (lo + hi) / 2; tc.style.zoom = m; if (tc.offsetHeight * m <= avail + 1) lo = m; else hi = m; }
      el.fontScale = Math.floor(lo * 40) / 40; tc.style.zoom = el.fontScale;
    }
    return false;
  };

  /* ---------- overlay: selection boxes, handles, guides ---------- */
  E.renderOverlay = (guides) => {
    const ov = $('#overlay'); if (!ov) return;
    const z = E.zoom; let h = '';
    const els = E.selected();
    if (E.grid) { /* grid drawn by CSS background on host */ }
    if (E.guides) h += `<div class="guide static v" style="left:${f2(SL.W(P()) / 2 * z)}px"></div><div class="guide static h" style="top:${f2(SL.H(P()) / 2 * z)}px"></div>`;
    for (const el of els) {
      const line = SL.isLine(el);
      h += `<div class="sel-box ${E.editing === el.id ? 'editing' : ''} ${line ? 'line' : ''}" style="left:${f2(el.x * z)}px;top:${f2(el.y * z)}px;width:${f2(el.w * z)}px;height:${f2(el.h * z)}px;transform:rotate(${el.rot || 0}deg)"></div>`;
    }
    if (els.length === 1 && !E.cropping) {
      const el = els[0], z2 = z;
      if (SL.isLine(el)) {
        const [s, e] = lineEnds(el);
        h += `<div class="hdl end" data-h="start" style="left:${f2(s.x * z2)}px;top:${f2(s.y * z2)}px"></div><div class="hdl end" data-h="end" style="left:${f2(e.x * z2)}px;top:${f2(e.y * z2)}px"></div>`;
      } else {
        h += `<div class="hdl-frame" style="left:${f2(el.x * z)}px;top:${f2(el.y * z)}px;width:${f2(el.w * z)}px;height:${f2(el.h * z)}px;transform:rotate(${el.rot || 0}deg)">` +
          ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].map(d => `<div class="hdl ${d}" data-h="${d}"></div>`).join('') +
          `<div class="hdl rot" data-h="rot" title="Rotate">${icon('rotate-cw', 12)}</div></div>`;
      }
    } else if (els.length > 1) {
      const b = E.bbox(els);
      h += `<div class="hdl-frame multi" style="left:${f2(b.x * z)}px;top:${f2(b.y * z)}px;width:${f2(b.w * z)}px;height:${f2(b.h * z)}px">` +
        ['nw', 'ne', 'se', 'sw'].map(d => `<div class="hdl ${d}" data-h="${d}"></div>`).join('') + '</div>';
    }
    if (E.cropping) {
      const el = SL.findEl(E.cropping); if (el) {
        const c = el.crop || { l: 0, t: 0, r: 0, b: 0 }; const iw = el.w / (1 - c.l - c.r), ih = el.h / (1 - c.t - c.b);
        const FW = iw * z, FH = ih * z, L = c.l * iw * z, T = c.t * ih * z, R = L + el.w * z, B = T + el.h * z;
        const hole = `path(evenodd, 'M0 0 H${f2(FW)} V${f2(FH)} H0 Z M${f2(L)} ${f2(T)} H${f2(R)} V${f2(B)} H${f2(L)} Z')`;
        h += `<div class="crop-full" style="left:${f2((el.x - c.l * iw) * z)}px;top:${f2((el.y - c.t * ih) * z)}px;width:${f2(FW)}px;height:${f2(FH)}px;background-image:url('${SL.assetURL(el.src)}');clip-path:${hole}"></div>`;
        h += `<div class="hdl-frame crop" style="left:${f2(el.x * z)}px;top:${f2(el.y * z)}px;width:${f2(el.w * z)}px;height:${f2(el.h * z)}px">` + ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].map(d => `<div class="hdl crop-h ${d}" data-h="${d}"></div>`).join('') + '</div>';
      }
    }
    if (guides) for (const g of guides) h += g.v != null ? `<div class="guide v" style="left:${f2(g.v * z)}px;top:${f2(g.a * z)}px;height:${f2((g.b - g.a) * z)}px"></div>` : `<div class="guide h" style="top:${f2(g.h * z)}px;left:${f2(g.a * z)}px;width:${f2((g.b - g.a) * z)}px"></div>`;
    if (E._marquee) { const m = E._marquee; h += `<div class="marquee" style="left:${f2(m.x * z)}px;top:${f2(m.y * z)}px;width:${f2(m.w * z)}px;height:${f2(m.h * z)}px"></div>`; }
    // comment markers
    (S().comments || []).forEach((c, i) => { h += `<div class="cmt-pin" data-cmt="${i}" style="left:${f2((c.x ?? 10 + i * 30) * z)}px;top:${f2((c.y ?? 10) * z)}px" title="${BTS.esc(c.author + ': ' + c.text)}">${icon('message-square', 14)}</div>`; });
    ov.innerHTML = h;
  };
  function lineEnds(el) {
    const s = { x: el.flipH ? el.x + el.w : el.x, y: el.flipV ? el.y + el.h : el.y };
    const e = { x: el.flipH ? el.x : el.x + el.w, y: el.flipV ? el.y : el.y + el.h };
    return [s, e];
  }
  E.bbox = (els) => {
    let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
    for (const el of els) { const b = rotBox(el); x1 = Math.min(x1, b.x); y1 = Math.min(y1, b.y); x2 = Math.max(x2, b.x + b.w); y2 = Math.max(y2, b.y + b.h); }
    return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
  };
  function rotBox(el) {
    if (!el.rot) return { x: el.x, y: el.y, w: el.w, h: el.h };
    const a = el.rot * Math.PI / 180, c = Math.abs(Math.cos(a)), s = Math.abs(Math.sin(a));
    const w = el.w * c + el.h * s, h = el.w * s + el.h * c, cx = el.x + el.w / 2, cy = el.y + el.h / 2;
    return { x: cx - w / 2, y: cy - h / 2, w, h };
  }
  E.rotBox = rotBox;

  E.renderRulers = () => {
    const rh = $('#ruler-h'), rv = $('#ruler-v'); if (!rh) return;
    rh.classList.toggle('hidden', !E.ruler || E.view !== 'normal'); rv.classList.toggle('hidden', !E.ruler || E.view !== 'normal');
    $('#stage-area').classList.toggle('with-ruler', E.ruler);
    if (!E.ruler) return;
    const host = $('#slide-host'), area = $('#stage-area'); const hr = host.getBoundingClientRect(), ar = area.getBoundingClientRect();
    const z = E.zoom, inch = 96 * z; const W = SL.W(P()) / 96, H = SL.H(P()) / 96;
    const ticks = (n, horiz) => { let s = ''; for (let i = 0; i <= n * 8; i++) { const p = f2(i / 8 * inch); const big = i % 8 === 0, mid = i % 4 === 0; s += `<i class="${big ? 'b' : mid ? 'm' : ''}" style="${horiz ? 'left' : 'top'}:${p}px">${big ? (i / 8) : ''}</i>`; } return s; };
    rh.innerHTML = `<div class="rt" style="left:${f2(hr.left - ar.left)}px;width:${f2(hr.width)}px">${ticks(Math.floor(W), true)}</div>`;
    rv.innerHTML = `<div class="rt" style="top:${f2(hr.top - ar.top)}px;height:${f2(hr.height)}px">${ticks(Math.floor(H), false)}</div>`;
  };

  /* ---------- rail ---------- */
  E.renderRail = () => {
    const rail = $('#rail'); if (!rail) return;
    const st = rail.scrollTop; rail.innerHTML = '';
    const tw = Math.max(120, rail.clientWidth - 46);
    P().slides.forEach((s, i) => {
      const sec = (P().sections || []).find(x => x.start === s.id);
      if (sec) { const h = document.createElement('div'); h.className = 'rail-sec'; h.dataset.sec = s.id; h.innerHTML = `${icon('chevron-down', 13)}<span>${BTS.esc(sec.name)}</span>`; rail.appendChild(h); }
      const it = document.createElement('div');
      it.className = 'rail-item' + (i === SL.cur ? ' cur' : '') + (E.railSel.has(i) ? ' sel' : '') + (s.hidden ? ' hidden-s' : '');
      it.dataset.i = i;
      const meta = `<div class="rail-num">${i + 1}${s.transition && s.transition.type !== 'none' ? `<span title="Has a transition">${icon('sparkles', 11)}</span>` : ''}${s.elements.some(e => e.anim) ? `<span title="Has animations">${icon('star', 11)}</span>` : ''}${(s.comments || []).length ? `<span title="Comments">${icon('message-square', 11)}</span>` : ''}</div>`;
      it.innerHTML = meta;
      const th = SL.thumb(s, P(), tw, i); it.appendChild(th);
      rail.appendChild(it);
    });
    rail.scrollTop = st;
    const cur = $('.rail-item.cur', rail); if (cur && !E._noRailScroll) { const r = cur.getBoundingClientRect(), rr = rail.getBoundingClientRect(); if (r.top < rr.top || r.bottom > rr.bottom) cur.scrollIntoView({ block: 'nearest' }); }
  };
  let thumbTimer = null;
  E.refreshThumb = (immediate) => {
    clearTimeout(thumbTimer);
    const go = () => {
      const it = $(`#rail .rail-item[data-i="${SL.cur}"]`); if (!it) return;
      const old = $('.sl-thumb', it); if (!old) return;
      old.replaceWith(SL.thumb(S(), P(), parseFloat(old.style.width), SL.cur));
    };
    if (immediate) go(); else thumbTimer = setTimeout(go, 220);
  };

  /* ---------- notes ---------- */
  E.renderNotes = () => {
    const n = $('#notes'); if (!n) return;
    if (document.activeElement !== n) n.textContent = S().notes || '';
    $('#center').classList.toggle('no-notes', !E.notesOn);
  };

  /* ---------- status bar & ribbon state ---------- */
  E.updateStatus = () => {
    const n = P().slides.length;
    BTS.setStatus('slide', `${icon('presentation', 14)} Slide ${SL.cur + 1} of ${n}${S().hidden ? ' <span class="sb-tag">Hidden</span>' : ''}`);
    const sec = E.sectionOf(SL.cur); BTS.setStatus('section', sec ? `${icon('rows-3', 13)} ${BTS.esc(sec.name)}` : '');
    const cm = (S().comments || []).length;
    BTS.setStatusRight(`<button class="sb-btn ${E.notesOn ? 'on' : ''}" data-cmd="toggle-notes" title="Show or hide notes">${icon('notebook-pen', 14)} Notes</button>` +
      `<button class="sb-btn" data-cmd="comments-pane" title="Comments">${icon('message-square', 14)} ${cm || ''}</button><span class="sb-div"></span>` +
      `<button class="sb-btn ${E.view === 'normal' ? 'on' : ''}" data-cmd="view-normal" title="Normal">${icon('panel-left', 15)}</button>` +
      `<button class="sb-btn ${E.view === 'sorter' ? 'on' : ''}" data-cmd="view-sorter" title="Slide sorter">${icon('layout-grid', 15)}</button>` +
      `<button class="sb-btn" data-cmd="present-window" title="Reading view">${icon('book-open', 15)}</button>` +
      `<button class="sb-btn" data-cmd="present-current" title="Slide show (Shift+F5)">${icon('presentation', 15)}</button><span class="sb-div"></span>`);
  };
  E.sectionOf = (idx) => { let cur = null; const secs = P().sections || []; for (let i = 0; i <= idx; i++) { const s = secs.find(x => x.start === P().slides[i].id); if (s) cur = s; } return cur; };

  E.updateRibbon = () => { E.onRibbonState && E.onRibbonState(); };

  /* ====================================================================== */
  /*  Selection                                                              */
  /* ====================================================================== */
  E.select = (ids, opts = {}) => {
    if (E.editing && !opts.keepEdit) E.endEdit();
    E.sel = [...new Set(ids)];
    E.renderOverlay(); E.updateRibbon(); E.refreshPane && E.refreshPane('sel');
  };
  E.groupIds = (el) => el.g ? S().elements.filter(e => e.g === el.g).map(e => e.id) : [el.id];
  E.selectAll = () => E.select(S().elements.map(e => e.id));

  /* ====================================================================== */
  /*  Pointer interaction on the canvas                                     */
  /* ====================================================================== */
  E.toSlide = (ev) => { const r = $('#slide-host').getBoundingClientRect(); return { x: (ev.clientX - r.left) / E.zoom, y: (ev.clientY - r.top) / E.zoom }; };

  function onStageDown(ev) {
    if (ev.button === 2) return onStageContext(ev);
    if (ev.button !== 0) return;
    const host = $('#slide-host');
    const hdl = ev.target.closest('.hdl');
    if (hdl) { ev.preventDefault(); return startHandle(ev, hdl.dataset.h, hdl.classList.contains('crop-h')); }
    if (ev.target.closest('.cmt-pin')) { E.openPane('comments'); return; }
    const phBtn = ev.target.closest('[data-act="fill-pic"]');
    if (phBtn) { ev.preventDefault(); const id = phBtn.closest('.el').dataset.id; E.select([id]); BTS.dispatch('insert-picture'); return; }
    const elNode = ev.target.closest('.sl-slide.editor > .el:not(.el-footer)');
    const p = E.toSlide(ev);
    if (E.drawMode) { ev.preventDefault(); return startDraw(ev, p); }
    // clicks inside the element being edited are left to the browser (caret placement)
    if (E.editing) {
      const ed = SL.findEl(E.editing);
      if (elNode && elNode.dataset.id === E.editing) {
        if (!ev.target.closest('.tc,.cell')) { ev.preventDefault(); if (ed.type !== 'table') placeCaretEnd($('.tc', elNode)); }
        return;
      }
      E.endEdit();
    }
    if (E.cropping && !(elNode && elNode.dataset.id === E.cropping)) E.endCrop();
    if (E.painter) { if (elNode) E.applyPainter(elNode.dataset.id); return; }
    if (!elNode) {
      if (!host.contains(ev.target) && !ev.target.closest('#stage-inner')) return;
      ev.preventDefault(); $('#stage').focus({ preventScroll: true });
      return startMarquee(ev, p, ev.shiftKey || ev.ctrlKey || ev.metaKey);
    }
    ev.preventDefault(); $('#stage').focus({ preventScroll: true });
    const id = elNode.dataset.id; const el = SL.findEl(id); if (!el) return;
    const ids = E.groupIds(el);
    const multi = ev.shiftKey || ev.ctrlKey || ev.metaKey;
    const wasSelected = E.sel.includes(id);
    if (multi) { if (wasSelected) { E.select(E.sel.filter(x => !ids.includes(x))); return; } E.select([...E.sel, ...ids]); }
    else if (!wasSelected) E.select(ids);
    if (E.cropping === id) return startCropPan(ev, el);
    startMove(ev, p, { clickedId: id, wasSelected, multi });
  }

  function placeCaretEnd(node) { if (!node) return; node.focus(); const r = document.createRange(); r.selectNodeContents(node); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
  E.placeCaretEnd = placeCaretEnd;

  /* ---------- move ---------- */
  function startMove(ev, p0, info) {
    const els = E.selected(); if (!els.length) return;
    const start = els.map(e => ({ e, x: e.x, y: e.y }));
    const bb0 = E.bbox(els);
    const cands = snapCandidates(els);
    let moved = false, dupDone = false;
    const nodes = () => els.map(e => $(`#slide-host .el[data-id="${e.id}"]`));
    let ns = nodes();
    const onMove = (e2) => {
      const p = E.toSlide(e2); let dx = p.x - p0.x, dy = p.y - p0.y;
      if (!moved && Math.hypot(dx, dy) * E.zoom < 3) return;
      if (!moved && e2.ctrlKey && !dupDone) { // Ctrl+drag duplicates
        dupDone = true; const copies = els.map(x => SL.reId(SL.clone(x))); regroup(copies); S().elements.push(...copies);
        E.sel = copies.map(c => c.id); E.renderCanvas(); els.splice(0, els.length, ...copies); start.forEach((s, i) => s.e = copies[i]); ns = nodes();
      }
      moved = true;
      if (e2.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
      let guides = [];
      if (E.snap && !e2.altKey) { const r = snap({ x: bb0.x + dx, y: bb0.y + dy, w: bb0.w, h: bb0.h }, cands, 'move'); dx += r.dx; dy += r.dy; guides = r.guides; }
      start.forEach((s, i) => { s.e.x = f2(s.x + dx); s.e.y = f2(s.y + dy); if (ns[i]) { ns[i].style.left = s.e.x + 'px'; ns[i].style.top = s.e.y + 'px'; } });
      E.renderOverlay(guides);
    };
    const onUp = () => {
      removeEventListener('pointermove', onMove); removeEventListener('pointerup', onUp);
      if (moved) { E.renderOverlay(); SL.commit(); E.refreshThumb(); E.updateRibbon(); }
      else if (info.wasSelected && !info.multi) {
        const el = SL.findEl(info.clickedId);
        if (E.sel.length > 1 && el.g) { E.select([info.clickedId]); } // click again inside group → select single member
        else if (el && (el.type === 'text' || el.type === 'shape' && !SL.isLine(el) || el.type === 'table') && E.sel.length === 1) E.startEdit(el.id, { at: ev });
      }
    };
    addEventListener('pointermove', onMove); addEventListener('pointerup', onUp);
  }
  function regroup(copies) { const map = {}; copies.forEach(c => { if (c.g) { map[c.g] = map[c.g] || SL.uid('g'); c.g = map[c.g]; } }); }
  E.regroup = regroup;

  function snapCandidates(moving) {
    const W = SL.W(P()), H = SL.H(P());
    const xs = [{ v: 0 }, { v: W / 2, c: 1 }, { v: W }], ys = [{ v: 0 }, { v: H / 2, c: 1 }, { v: H }];
    const ids = new Set(moving.map(m => m.id));
    for (const o of S().elements) {
      if (ids.has(o.id)) continue; const b = rotBox(o);
      xs.push({ v: b.x, o: b }, { v: b.x + b.w / 2, o: b, c: 1 }, { v: b.x + b.w, o: b });
      ys.push({ v: b.y, o: b }, { v: b.y + b.h / 2, o: b, c: 1 }, { v: b.y + b.h, o: b });
    }
    return { xs, ys, W, H };
  }
  function snap(b, cands, mode, edges) {
    const thr = 6 / E.zoom; let dx = 0, dy = 0; const guides = [];
    const ex = edges ? edges.x : [b.x, b.x + b.w / 2, b.x + b.w], ey = edges ? edges.y : [b.y, b.y + b.h / 2, b.y + b.h];
    let best = null;
    for (const c of cands.xs) for (const e of ex) { const d = c.v - e; if (Math.abs(d) < thr && (!best || Math.abs(d) < Math.abs(best.d))) best = { d, c }; }
    if (best) { dx = best.d; const o = best.c.o; const a = Math.min(b.y, o ? o.y : 0), z = Math.max(b.y + b.h, o ? o.y + o.h : cands.H); guides.push({ v: best.c.v, a: o ? a : 0, b: o ? z : cands.H }); }
    best = null;
    for (const c of cands.ys) for (const e of ey) { const d = c.v - e; if (Math.abs(d) < thr && (!best || Math.abs(d) < Math.abs(best.d))) best = { d, c }; }
    if (best) { dy = best.d; const o = best.c.o; const a = Math.min(b.x, o ? o.x : 0), z = Math.max(b.x + b.w, o ? o.x + o.w : cands.W); guides.push({ h: best.c.v, a: o ? a : 0, b: o ? z : cands.W }); }
    return { dx, dy, guides };
  }

  /* ---------- resize / rotate ---------- */
  function startHandle(ev, h, isCrop) {
    const els = E.selected(); if (!els.length) return;
    if (isCrop) return startCropHandle(ev, h);
    if (h === 'rot') return startRotate(ev, els[0]);
    if (h === 'start' || h === 'end') return startLineEnd(ev, els[0], h);
    if (els.length > 1) return startMultiResize(ev, els, h);
    const el = els[0]; const o = { x: el.x, y: el.y, w: el.w, h: el.h };
    const hx = h.includes('e') ? 1 : h.includes('w') ? -1 : 0, hy = h.includes('s') ? 1 : h.includes('n') ? -1 : 0;
    const a = (el.rot || 0) * Math.PI / 180, cos = Math.cos(a), sin = Math.sin(a);
    const c0 = { x: o.x + o.w / 2, y: o.y + o.h / 2 };
    const fixedL = { x: -hx * o.w / 2, y: -hy * o.h / 2 };
    const cands = snapCandidates([el]);
    const lockAspect = el.type === 'image' || el.type === 'icon' || el.lockAspect;
    let moved = false; let node = $(`#slide-host .el[data-id="${el.id}"]`);
    const onMove = (e2) => {
      moved = true;
      const p = E.toSlide(e2);
      const q = { x: cos * (p.x - c0.x) + sin * (p.y - c0.y), y: -sin * (p.x - c0.x) + cos * (p.y - c0.y) };
      let nw = hx ? (q.x - fixedL.x) * hx : o.w, nh = hy ? (q.y - fixedL.y) * hy : o.h;
      const aspect = (e2.shiftKey !== lockAspect) && hx && hy;
      if (aspect) { const s = Math.max(nw / o.w, nh / o.h); nw = o.w * s; nh = o.h * s; }
      nw = Math.max(4, nw); nh = Math.max(el.type === 'text' && el.autofit === 'resize' ? 4 : 4, nh);
      let guides = [];
      if (E.snap && !el.rot && !e2.altKey && !aspect) { // snap moving edges
        const bx = hx < 0 ? o.x + o.w - nw : o.x, by = hy < 0 ? o.y + o.h - nh : o.y;
        const r = snap({ x: bx, y: by, w: nw, h: nh }, cands, 'resize', { x: hx ? [hx > 0 ? bx + nw : bx] : [], y: hy ? [hy > 0 ? by + nh : by] : [] });
        if (hx > 0) nw += r.dx; else if (hx < 0) nw -= r.dx;
        if (hy > 0) nh += r.dy; else if (hy < 0) nh -= r.dy;
        guides = r.guides;
      }
      const lc = { x: fixedL.x + hx * nw / 2, y: fixedL.y + hy * nh / 2 };
      if (!hx) lc.x = 0; if (!hy) lc.y = 0;
      const wc = { x: c0.x + cos * lc.x - sin * lc.y, y: c0.y + sin * lc.x + cos * lc.y };
      el.w = f2(nw); el.h = f2(nh); el.x = f2(wc.x - nw / 2); el.y = f2(wc.y - nh / 2);
      if (el.type === 'table') el.userH = el.h;
      const nn = SL.renderElement(el, P(), { editor: true, editingId: E.editing, index: SL.cur }); node.replaceWith(nn); node = nn;
      if (el.type === 'text' || el.type === 'shape') E.fitText(el, nn);
      E.renderOverlay(guides);
    };
    const onUp = () => { removeEventListener('pointermove', onMove); removeEventListener('pointerup', onUp); if (moved) { E.renderCanvas(); SL.commit(); E.refreshThumb(); E.updateRibbon(); } };
    addEventListener('pointermove', onMove); addEventListener('pointerup', onUp);
  }
  function startMultiResize(ev, els, h) {
    const b0 = E.bbox(els); const st = els.map(e => ({ e, x: e.x, y: e.y, w: e.w, h: e.h, size: e.size }));
    const fx = h.includes('w') ? b0.x + b0.w : b0.x, fy = h.includes('n') ? b0.y + b0.h : b0.y;
    let moved = false;
    const onMove = (e2) => {
      moved = true; const p = E.toSlide(e2);
      let sx = Math.abs(p.x - fx) / b0.w, sy = Math.abs(p.y - fy) / b0.h;
      if (e2.shiftKey) sx = sy = Math.max(sx, sy);
      sx = Math.max(.05, sx); sy = Math.max(.05, sy);
      st.forEach(s => {
        const nx = fx + (s.x - fx) * sx * (h.includes('w') ? 1 : 1), ny = fy + (s.y - fy) * sy;
        s.e.x = f2(h.includes('w') ? fx - (fx - s.x) * sx : fx + (s.x - fx) * sx); s.e.y = f2(h.includes('n') ? fy - (fy - s.y) * sy : ny);
        s.e.w = f2(s.w * sx); s.e.h = f2(s.h * sy); void nx;
      });
      E.renderCanvas();
    };
    const onUp = () => { removeEventListener('pointermove', onMove); removeEventListener('pointerup', onUp); if (moved) { SL.commit(); E.refreshThumb(); } };
    addEventListener('pointermove', onMove); addEventListener('pointerup', onUp);
  }
  function startRotate(ev, el) {
    const c = { x: el.x + el.w / 2, y: el.y + el.h / 2 }; let moved = false; let node = $(`#slide-host .el[data-id="${el.id}"]`);
    const onMove = (e2) => {
      moved = true; const p = E.toSlide(e2);
      let a = Math.atan2(p.y - c.y, p.x - c.x) * 180 / Math.PI + 90;
      if (e2.shiftKey) a = Math.round(a / 15) * 15; else for (const s of [0, 90, 180, 270, 360]) if (Math.abs(a - s) < 3 || Math.abs(a + 360 - s) < 3) a = s;
      a = ((Math.round(a * 10) / 10) % 360 + 360) % 360;
      el.rot = a; node.style.transform = `rotate(${a}deg)`; E.renderOverlay();
      showTip(e2, Math.round(a) + '°');
    };
    const onUp = () => { removeEventListener('pointermove', onMove); removeEventListener('pointerup', onUp); hideTip(); if (moved) { SL.commit(); E.refreshThumb(); E.renderCanvas(); } };
    addEventListener('pointermove', onMove); addEventListener('pointerup', onUp);
  }
  function startLineEnd(ev, el, which) {
    const [s0, e0] = lineEnds(el); const fixed = which === 'start' ? e0 : s0; let moved = false;
    const cands = snapCandidates([el]);
    const onMove = (e2) => {
      moved = true; let p = E.toSlide(e2);
      if (e2.shiftKey) { const dx = p.x - fixed.x, dy = p.y - fixed.y, a = Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * Math.PI / 4, r = Math.hypot(dx, dy); p = { x: fixed.x + Math.cos(a) * r, y: fixed.y + Math.sin(a) * r }; }
      else if (E.snap && !e2.altKey) { const r = snap({ x: p.x, y: p.y, w: 0, h: 0 }, cands, 'pt', { x: [p.x], y: [p.y] }); p = { x: p.x + r.dx, y: p.y + r.dy }; E.renderOverlay(r.guides); }
      const S0 = which === 'start' ? p : fixed, E0 = which === 'start' ? fixed : p;
      el.x = f2(Math.min(S0.x, E0.x)); el.y = f2(Math.min(S0.y, E0.y)); el.w = f2(Math.abs(E0.x - S0.x)); el.h = f2(Math.abs(E0.y - S0.y));
      el.flipH = E0.x < S0.x; el.flipV = E0.y < S0.y; el.rot = 0;
      const node = $(`#slide-host .el[data-id="${el.id}"]`); node.replaceWith(SL.renderElement(el, P(), { editor: true }));
      if (!(E.snap && !e2.altKey)) E.renderOverlay();
    };
    const onUp = () => { removeEventListener('pointermove', onMove); removeEventListener('pointerup', onUp); if (moved) { E.renderOverlay(); SL.commit(); E.refreshThumb(); } };
    addEventListener('pointermove', onMove); addEventListener('pointerup', onUp);
  }

  /* ---------- crop ---------- */
  E.startCrop = (id) => { const el = SL.findEl(id); if (!el || el.type !== 'image' || !el.src) return; if (E.editing) E.endEdit(); E.cropping = id; E.sel = [id]; el.crop = el.crop || { l: 0, t: 0, r: 0, b: 0 }; E.renderOverlay(); E.updateRibbon(); $('#slide-host').classList.add('cropping'); };
  E.endCrop = () => { if (!E.cropping) return; E.cropping = null; $('#slide-host').classList.remove('cropping'); SL.commit(); E.renderCanvas(); E.refreshThumb(); E.updateRibbon(); };
  function startCropHandle(ev, h) {
    const el = SL.findEl(E.cropping); const c0 = Object.assign({}, el.crop); const o = { x: el.x, y: el.y, w: el.w, h: el.h };
    const iw = o.w / (1 - c0.l - c0.r), ih = o.h / (1 - c0.t - c0.b); const ix = o.x - c0.l * iw, iy = o.y - c0.t * ih;
    let moved = false;
    const onMove = (e2) => {
      moved = true; const p = E.toSlide(e2);
      if (h.includes('w')) { const nx = Math.max(ix, Math.min(p.x, o.x + o.w - 8)); el.crop.l = (nx - ix) / iw; el.x = nx; el.w = o.x + o.w - nx; }
      if (h.includes('e')) { const nr = Math.min(ix + iw, Math.max(p.x, o.x + 8)); el.crop.r = (ix + iw - nr) / iw; el.w = nr - el.x; }
      if (h.includes('n')) { const ny = Math.max(iy, Math.min(p.y, o.y + o.h - 8)); el.crop.t = (ny - iy) / ih; el.y = ny; el.h = o.y + o.h - ny; }
      if (h.includes('s')) { const nb = Math.min(iy + ih, Math.max(p.y, o.y + 8)); el.crop.b = (iy + ih - nb) / ih; el.h = nb - el.y; }
      const node = $(`#slide-host .el[data-id="${el.id}"]`); node.replaceWith(SL.renderElement(el, P(), { editor: true })); E.renderOverlay();
    };
    const onUp = () => { removeEventListener('pointermove', onMove); removeEventListener('pointerup', onUp); if (moved) SL.commit({ silent: false }); };
    addEventListener('pointermove', onMove); addEventListener('pointerup', onUp);
  }
  function startCropPan(ev, el) {
    const p0 = E.toSlide(ev); const c0 = Object.assign({}, el.crop); const iw = el.w / (1 - c0.l - c0.r), ih = el.h / (1 - c0.t - c0.b);
    const onMove = (e2) => {
      const p = E.toSlide(e2); let dl = -(p.x - p0.x) / iw, dt = -(p.y - p0.y) / ih;
      dl = Math.max(-c0.l, Math.min(c0.r, dl)); dt = Math.max(-c0.t, Math.min(c0.b, dt));
      el.crop = { l: c0.l + dl, r: c0.r - dl, t: c0.t + dt, b: c0.b - dt };
      const node = $(`#slide-host .el[data-id="${el.id}"]`); node.replaceWith(SL.renderElement(el, P(), { editor: true })); E.renderOverlay();
    };
    const onUp = () => { removeEventListener('pointermove', onMove); removeEventListener('pointerup', onUp); SL.commit(); };
    addEventListener('pointermove', onMove); addEventListener('pointerup', onUp);
  }

  /* ---------- marquee ---------- */
  function startMarquee(ev, p0, add) {
    const base = add ? E.sel.slice() : [];
    if (!add) E.select([]);
    let moved = false;
    const onMove = (e2) => {
      const p = E.toSlide(e2); if (!moved && Math.hypot(p.x - p0.x, p.y - p0.y) * E.zoom < 4) return; moved = true;
      const m = { x: Math.min(p.x, p0.x), y: Math.min(p.y, p0.y), w: Math.abs(p.x - p0.x), h: Math.abs(p.y - p0.y) }; E._marquee = m;
      const hit = S().elements.filter(el => { const b = rotBox(el); return b.x >= m.x && b.y >= m.y && b.x + b.w <= m.x + m.w && b.y + b.h <= m.y + m.h; });
      const ids = new Set(base); hit.forEach(el => E.groupIds(el).forEach(i => ids.add(i)));
      E.sel = [...ids]; E.renderOverlay();
    };
    const onUp = () => { removeEventListener('pointermove', onMove); removeEventListener('pointerup', onUp); E._marquee = null; E.renderOverlay(); E.updateRibbon(); };
    addEventListener('pointermove', onMove); addEventListener('pointerup', onUp);
  }

  /* ---------- draw new shapes / text boxes ---------- */
  E.setDrawMode = (mode) => { E.drawMode = mode; $('#stage').classList.toggle('drawing', !!mode); if (mode) BTS.toast(mode.type === 'text' ? 'Click or drag on the slide to add a text box' : 'Click or drag on the slide to draw the shape', { ms: 1800 }); };
  function startDraw(ev, p0) {
    const mode = E.drawMode; E.setDrawMode(null);
    const line = mode.type === 'shape' && (mode.kind === 'line' || mode.kind === 'arrowLine');
    const el = mode.type === 'text'
      ? SL.newElement('text', { x: p0.x, y: p0.y, w: 10, h: 10, html: '', autofit: 'resize', size: 18 })
      : SL.newElement('shape', Object.assign({ kind: mode.kind, x: p0.x, y: p0.y, w: 1, h: 1 }, line ? { stroke: 'accent1', strokeWidth: 2.25, fill: 'none' } : {}));
    S().elements.push(el); E.sel = [el.id]; E.renderCanvas();
    let moved = false;
    const onMove = (e2) => {
      const p = E.toSlide(e2); if (!moved && Math.hypot(p.x - p0.x, p.y - p0.y) * E.zoom < 4) return; moved = true;
      let w = p.x - p0.x, h = p.y - p0.y;
      if (e2.shiftKey) { if (line) { const a = Math.round(Math.atan2(h, w) / (Math.PI / 4)) * Math.PI / 4, r = Math.hypot(w, h); w = Math.cos(a) * r; h = Math.sin(a) * r; } else { const m = Math.max(Math.abs(w), Math.abs(h)); w = Math.sign(w || 1) * m; h = Math.sign(h || 1) * m; } }
      el.x = f2(Math.min(p0.x, p0.x + w)); el.y = f2(Math.min(p0.y, p0.y + h)); el.w = f2(Math.abs(w)); el.h = f2(Math.abs(h));
      if (line) { el.flipH = w < 0; el.flipV = h < 0; }
      const node = $(`#slide-host .el[data-id="${el.id}"]`); if (node) node.replaceWith(SL.renderElement(el, P(), { editor: true }));
      E.renderOverlay();
    };
    const onUp = () => {
      removeEventListener('pointermove', onMove); removeEventListener('pointerup', onUp);
      if (!moved) { if (mode.type === 'text') { el.w = 360; el.h = 40; } else if (line) { el.w = 240; el.h = 0; } else { el.w = 180; el.h = 180; el.x -= 90; el.y -= 90; } }
      if (mode.type === 'text') { if (el.w < 40) el.w = 360; if (moved) el.autofit = 'resize'; }
      E.renderCanvas(); SL.commit(); E.refreshThumb(true);
      if (mode.type === 'text') E.startEdit(el.id);
    };
    addEventListener('pointermove', onMove); addEventListener('pointerup', onUp);
  }

  let tipEl = null;
  function showTip(ev, text) { if (!tipEl) { tipEl = document.createElement('div'); tipEl.className = 'drag-tip'; document.body.appendChild(tipEl); } tipEl.textContent = text; tipEl.style.left = ev.clientX + 14 + 'px'; tipEl.style.top = ev.clientY + 14 + 'px'; }
  function hideTip() { if (tipEl) { tipEl.remove(); tipEl = null; } }

  /* ====================================================================== */
  /*  Text editing                                                           */
  /* ====================================================================== */
  E.startEdit = (id, opts = {}) => {
    const el = SL.findEl(id); if (!el) return;
    if (!(el.type === 'text' || el.type === 'table' || (el.type === 'shape' && !SL.isLine(el)))) return;
    if (E.editing && E.editing !== id) E.endEdit();
    E.sel = [id]; E.editing = id; E._editOpts = opts;
    E.renderCanvas();
  };
  // called by renderCanvas when E.editing is set: make the node editable and place the caret
  E.resumeEdit = () => {
    const el = SL.findEl(E.editing); if (!el) { E.editing = null; return; }
    const node = $(`#slide-host .el[data-id="${el.id}"]`); if (!node) return;
    node.classList.add('editing');
    const opts = E._editOpts || {}; E._editOpts = null;
    if (el.type === 'table') {
      const cells = $$('.cell', node);
      cells.forEach(c => c.addEventListener('input', () => syncTable(el, node)));
      cells.forEach(c => c.addEventListener('focus', () => { const td = c.closest('td'); E.tableCell = { r: +td.dataset.r, c: +td.dataset.c }; E.updateRibbon(); }));
      let target = null;
      if (opts.at) { const t = document.elementFromPoint(opts.at.clientX, opts.at.clientY); target = t && t.closest('.cell'); }
      if (!target && E.tableCell) target = $(`td[data-r="${E.tableCell.r}"][data-c="${E.tableCell.c}"] .cell`, node);
      target = target || cells[0];
      if (target) { if (opts.at) { target.focus(); caretAt(opts.at); } else placeCaretEnd(target); const td = target.closest('td'); E.tableCell = { r: +td.dataset.r, c: +td.dataset.c }; }
      E.renderOverlay(); E.updateRibbon(); return;
    }
    const tc = $('.tc', node);
    tc.contentEditable = 'true'; tc.spellcheck = SL.spell !== false;
    if (SL.isEmptyText(el)) tc.innerHTML = el.list === 'ul' ? '<ul><li><br></li></ul>' : el.list === 'ol' ? '<ol><li><br></li></ol>' : '<p><br></p>';
    tc.addEventListener('input', () => {
      el.html = tc.innerHTML; E.fitText(el, node);
      E.renderOverlay(); E.refreshThumb(); BTS.markDirty();
    });
    tc.addEventListener('paste', onTextPaste);
    tc.focus({ preventScroll: true });
    if (opts.at && !SL.isEmptyText(el)) caretAt(opts.at);
    else if (opts.selectAll) { const r = document.createRange(); r.selectNodeContents(tc); const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
    else placeCaretEnd(tc.querySelector('p,li') || tc);
    if (SL.isEmptyText(el)) { const b = tc.querySelector('p,li'); if (b) { const r = document.createRange(); r.setStart(b, 0); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); } }
    E.renderOverlay(); E.updateRibbon();
  };
  function caretAt(ev) { const r = document.caretRangeFromPoint(ev.clientX, ev.clientY); if (r) { const s = getSelection(); s.removeAllRanges(); s.addRange(r); } }
  function syncTable(el, node) {
    $$('.cell', node).forEach(c => { const td = c.closest('td'); const r = +td.dataset.r, ci = +td.dataset.c; if (el.rows[r] && el.rows[r][ci]) el.rows[r][ci].html = c.innerHTML; });
    const t = $('table', node); if (t && t.offsetHeight > el.h + 1) { el.h = t.offsetHeight; node.style.height = el.h + 'px'; E.renderOverlay(); }
    E.refreshThumb(); BTS.markDirty();
  }
  E.syncEditing = () => {
    if (!E.editing) return;
    const el = SL.findEl(E.editing); const node = $(`#slide-host .el[data-id="${E.editing}"]`); if (!el || !node) return;
    if (el.type === 'table') syncTable(el, node); else { const tc = $('.tc', node); el.html = tc.innerHTML; E.fitText(el, node); }
  };
  function onTextPaste(e) {
    const dt = e.clipboardData; if (!dt) return;
    e.preventDefault();
    const html = dt.getData('text/html');
    if (html && !html.includes('data-bts-slides')) {
      const clean = SL.sanitize(html); const d = document.createElement('div'); d.innerHTML = clean;
      const blocks = $$(':scope > p, :scope > ul > li, :scope > ol > li', d);
      if (blocks.length <= 1) document.execCommand('insertHTML', false, (blocks[0] || d).innerHTML);
      else document.execCommand('insertText', false, SL.plainText(clean));
    } else document.execCommand('insertText', false, dt.getData('text/plain'));
  }
  E.endEdit = () => {
    if (!E.editing) return;
    const id = E.editing; const el = SL.findEl(id);
    E.syncEditing(); E.editing = null;
    if (el) {
      if (el.type === 'table') el.rows.forEach(r => r.forEach(c => { c.html = SL.sanitize(c.html || ''); if (!SL.plainText(c.html).trim() && !/<br>/.test(c.html)) c.html = ''; if (SL.isEmptyText({ html: c.html })) c.html = ''; }));
      else {
        el.html = SL.sanitize(el.html);
        if (SL.isEmptyText(el)) {
          el.html = '';
          if (el.type === 'text' && !el.ph && !el.placeholder) { S().elements = S().elements.filter(x => x !== el); E.sel = []; }
        }
      }
    }
    getSelection().removeAllRanges();
    E.renderCanvas(); SL.commit(); E.refreshThumb(true); E.updateRibbon();
  };
  E.editNode = () => { if (!E.editing) return null; const n = $(`#slide-host .el[data-id="${E.editing}"]`); return n && (SL.findEl(E.editing).type === 'table' ? (document.activeElement.closest && document.activeElement.closest('.cell')) : $('.tc', n)); };

  /* ====================================================================== */
  /*  Element operations                                                     */
  /* ====================================================================== */
  E.addElement = (el, opts = {}) => {
    if (E.editing) E.endEdit();
    S().elements.push(el); E.sel = [el.id];
    E.renderCanvas(); SL.commit(); E.refreshThumb(true); E.updateRibbon();
    if (opts.edit) E.startEdit(el.id, { selectAll: true });
    return el;
  };
  E.deleteSelected = () => {
    const ids = new Set(E.sel); if (!ids.size) return;
    const s = S();
    s.elements = s.elements.filter(e => {
      if (!ids.has(e.id)) return true;
      if (e.ph && ((e.type === 'text' && !SL.isEmptyText(e)) || (e.type === 'image' && e.src))) { if (e.type === 'text') e.html = ''; else { e.src = null; e.crop = null; } return true; } // first delete empties a placeholder
      return false;
    });
    E.sel = []; E.renderCanvas(); SL.commit(); E.refreshThumb(true); E.updateRibbon();
  };
  E.nudge = (dx, dy) => {
    for (const el of E.selected()) { el.x = f2(el.x + dx); el.y = f2(el.y + dy); }
    E.renderCanvas(); clearTimeout(E._nudgeT); E._nudgeT = setTimeout(() => { SL.commit(); E.refreshThumb(); }, 350);
  };
  E.duplicateSelected = () => {
    const els = E.selected(); if (!els.length) return;
    const copies = els.map(e => { const c = SL.reId(SL.clone(e)); c.x += 20; c.y += 20; delete c.ph; return c; });
    regroup(copies); S().elements.push(...copies); E.sel = copies.map(c => c.id);
    E.renderCanvas(); SL.commit(); E.refreshThumb(true); E.updateRibbon();
  };
  E.zorder = (how) => {
    const s = S(); const ids = new Set(E.sel); if (!ids.size) return;
    const arr = s.elements; const selEls = arr.filter(e => ids.has(e.id)), rest = arr.filter(e => !ids.has(e.id));
    if (how === 'front') s.elements = [...rest, ...selEls];
    else if (how === 'back') s.elements = [...selEls, ...rest];
    else if (how === 'forward') { for (let i = arr.length - 2; i >= 0; i--) if (ids.has(arr[i].id) && !ids.has(arr[i + 1].id)) [arr[i], arr[i + 1]] = [arr[i + 1], arr[i]]; }
    else if (how === 'backward') { for (let i = 1; i < arr.length; i++) if (ids.has(arr[i].id) && !ids.has(arr[i - 1].id)) [arr[i], arr[i - 1]] = [arr[i - 1], arr[i]]; }
    E.renderCanvas(); SL.commit(); E.refreshThumb(true);
  };
  E.align = (how, toSlide) => {
    const els = E.selected(); if (!els.length) return;
    const rel = (toSlide || els.length === 1) ? { x: 0, y: 0, w: SL.W(P()), h: SL.H(P()) } : E.bbox(els);
    // groups move as a unit
    const units = []; const seen = new Set();
    for (const el of els) { if (seen.has(el.id)) continue; const mem = el.g ? els.filter(e => e.g === el.g) : [el]; mem.forEach(m => seen.add(m.id)); units.push(mem); }
    if (how === 'dist-h' || how === 'dist-v') {
      if (units.length < 3 && !toSlide) { BTS.toast('Select three or more objects to distribute'); return; }
      const key = how === 'dist-h' ? 'x' : 'y', dim = how === 'dist-h' ? 'w' : 'h';
      const ub = units.map(u => ({ u, b: E.bbox(u) })).sort((a, b) => a.b[key] - b.b[key]);
      const total = ub.reduce((a, x) => a + x.b[dim], 0);
      const start = toSlide ? 0 : ub[0].b[key], span = toSlide ? (key === 'x' ? SL.W(P()) : SL.H(P())) : (ub[ub.length - 1].b[key] + ub[ub.length - 1].b[dim] - start);
      const gap = (span - total) / (ub.length - (toSlide ? -1 : 1)); let pos = start + (toSlide ? gap : 0);
      ub.forEach(({ u, b }) => { const d = pos - b[key]; u.forEach(e => e[key] = f2(e[key] + d)); pos += b[dim] + gap; });
    } else {
      for (const u of units) {
        const b = E.bbox(u); let dx = 0, dy = 0;
        if (how === 'left') dx = rel.x - b.x; if (how === 'center') dx = rel.x + rel.w / 2 - (b.x + b.w / 2); if (how === 'right') dx = rel.x + rel.w - (b.x + b.w);
        if (how === 'top') dy = rel.y - b.y; if (how === 'middle') dy = rel.y + rel.h / 2 - (b.y + b.h / 2); if (how === 'bottom') dy = rel.y + rel.h - (b.y + b.h);
        u.forEach(e => { e.x = f2(e.x + dx); e.y = f2(e.y + dy); });
      }
    }
    E.renderCanvas(); SL.commit(); E.refreshThumb(true);
  };
  E.group = () => { const els = E.selected(); if (els.length < 2) return BTS.toast('Select two or more objects to group'); const g = SL.uid('g'); els.forEach(e => e.g = g); SL.commit(); E.renderOverlay(); E.updateRibbon(); };
  E.ungroup = () => { const els = E.selected(); els.forEach(e => delete e.g); SL.commit(); E.renderOverlay(); E.updateRibbon(); };

  /* ---------- clipboard ---------- */
  const MARK = 'BTS-SLIDES-CLIP:';
  E.copy = (cut) => {
    const els = E.selected(); if (!els.length) return false;
    E.clip = { els: SL.clone(els), assets: Object.fromEntries(els.filter(e => e.src).map(e => [e.src, SL.assets[e.src]])), n: 0 };
    try { navigator.clipboard.writeText(MARK + els.map(e => SL.plainText(e.html || '')).join('\n')).catch(() => {}); } catch {}
    if (cut) E.deleteSelected();
    return true;
  };
  E.pasteInternal = () => {
    if (!E.clip) return false;
    Object.assign(SL.assets, E.clip.assets || {});
    E.clip.n++;
    const copies = E.clip.els.map(e => { const c = SL.reId(SL.clone(e)); delete c.ph; c.x += 20 * E.clip.n * (E.clip.slide === S().id ? 1 : 0); c.y += 20 * E.clip.n * (E.clip.slide === S().id ? 1 : 0); return c; });
    regroup(copies); S().elements.push(...copies); E.sel = copies.map(c => c.id);
    E.renderCanvas(); SL.commit(); E.refreshThumb(true); E.updateRibbon();
    return true;
  };
  E.pasteFromSystem = async (dt) => {
    // dt: DataTransfer (paste event) or null → async clipboard
    let text = '', html = '', img = null;
    if (dt) {
      text = dt.getData('text/plain'); html = dt.getData('text/html');
      for (const it of dt.items || []) if (it.kind === 'file' && it.type.startsWith('image/')) img = it.getAsFile();
    } else {
      try { const items = await navigator.clipboard.read(); for (const it of items) { const t = it.types.find(x => x.startsWith('image/')); if (t) img = await it.getType(t); if (it.types.includes('text/html')) html = await (await it.getType('text/html')).text(); if (it.types.includes('text/plain')) text = await (await it.getType('text/plain')).text(); } }
      catch { try { text = await navigator.clipboard.readText(); } catch {} }
    }
    if (text.startsWith(MARK) && E.clip) return E.pasteInternal();
    if (img) { const url = await BTS.bytesToDataURL(new Uint8Array(await img.arrayBuffer()), img.type); return E.insertImage(url); }
    if (!text && !html && E.clip) return E.pasteInternal();
    if (!text && !html) return BTS.toast('The clipboard is empty');
    const content = html ? SL.sanitize(html) : SL.textToHTML(text);
    const W = SL.W(P());
    const el = SL.newElement('text', { x: W * .15, y: SL.H(P()) * .3, w: W * .7, h: 60, html: content, autofit: 'resize' });
    E.addElement(el);
  };

  E.insertImage = (url, opts = {}) => new Promise(res => {
    const img = new Image();
    img.onload = () => {
      const W = SL.W(P()), H = SL.H(P());
      const id = SL.addAsset(url);
      const target = opts.into || E.selected().find(e => e.type === 'image' && !e.src) || (opts.auto !== false && S().elements.find(e => e.type === 'image' && e.ph && !e.src));
      if (target) { // fill picture placeholder with a "cover" crop
        target.src = id; const ar = img.naturalWidth / img.naturalHeight, br = target.w / target.h;
        target.crop = ar > br ? { l: (1 - br / ar) / 2, r: (1 - br / ar) / 2, t: 0, b: 0 } : { t: (1 - ar / br) / 2, b: (1 - ar / br) / 2, l: 0, r: 0 };
        target.natW = img.naturalWidth; target.natH = img.naturalHeight;
        E.sel = [target.id]; E.renderCanvas(); SL.commit(); E.refreshThumb(true); E.updateRibbon(); return res(target);
      }
      let w = img.naturalWidth, h = img.naturalHeight; const k = Math.min(1, W * .6 / w, H * .6 / h); w *= k; h *= k;
      const at = opts.at || { x: (W - w) / 2, y: (H - h) / 2 };
      res(E.addElement(SL.newElement('image', { src: id, x: f2(at.x), y: f2(at.y), w: f2(w), h: f2(h), natW: img.naturalWidth, natH: img.naturalHeight, name: opts.name })));
    };
    img.onerror = () => { BTS.toast('That picture could not be read', { error: true }); res(null); };
    img.src = url;
  });

  /* ====================================================================== */
  /*  Slides                                                                 */
  /* ====================================================================== */
  E.goto = (i, opts = {}) => {
    if (E.editing) E.endEdit();
    if (E.cropping) E.endCrop();
    i = Math.max(0, Math.min(P().slides.length - 1, i));
    const changed = i !== SL.cur; SL.cur = i;
    if (!opts.keepRailSel) E.railSel = new Set([i]);
    if (changed) E.sel = [];
    E.renderAll();
  };
  E.newSlide = (layout) => {
    if (E.editing) E.endEdit();
    const cur = S(); layout = layout || (cur.layout === 'title' ? 'titleContent' : (cur.layout === 'blank' ? 'blank' : cur.layout === 'section' ? 'titleContent' : cur.layout));
    const s = SL.newSlide(P(), layout);
    s.transition = Object.assign({}, cur.transition || { type: 'none', dur: 700 });
    if (cur.background && P().slides.every(x => JSON.stringify(x.background) === JSON.stringify(cur.background))) s.background = SL.clone(cur.background);
    if (cur.layer && cur.layer.length && cur.layerShared) { s.layer = SL.clone(cur.layer); s.layerShared = true; }
    P().slides.splice(SL.cur + 1, 0, s); SL.cur++; E.railSel = new Set([SL.cur]); E.sel = [];
    SL.commit(); E.renderAll();
  };
  E.duplicateSlides = () => {
    const idx = [...E.railSel].sort((a, b) => a - b); const last = idx[idx.length - 1];
    const copies = idx.map(i => { const s = SL.clone(P().slides[i]); s.id = SL.uid('s'); s.elements.forEach(SL.reId); regroup(s.elements); return s; });
    P().slides.splice(last + 1, 0, ...copies); SL.cur = last + copies.length; E.railSel = new Set(copies.map((_, k) => last + 1 + k));
    SL.commit(); E.renderAll();
  };
  E.deleteSlides = () => {
    const idx = [...E.railSel].sort((a, b) => b - a);
    if (idx.length >= P().slides.length) { // keep one slide
      P().slides = [SL.newSlide(P(), 'title')]; SL.cur = 0;
    } else {
      const removedIds = idx.map(i => P().slides[i].id);
      idx.forEach(i => P().slides.splice(i, 1));
      fixSections(removedIds);
      SL.cur = Math.min(idx[idx.length - 1], P().slides.length - 1);
    }
    E.railSel = new Set([SL.cur]); E.sel = []; SL.commit(); E.renderAll();
  };
  function fixSections(removedIds) {
    if (!P().sections) return;
    // a section whose first slide was removed moves to the next surviving slide (or disappears)
    P().sections = P().sections.filter(sec => P().slides.some(s => s.id === sec.start) || !removedIds.includes(sec.start));
  }
  E.moveSlides = (indices, to) => { // to = insertion index in the original array
    indices = [...indices].sort((a, b) => a - b);
    const moving = indices.map(i => P().slides[i]); const curId = S().id;
    const before = indices.filter(i => i < to).length;
    P().slides = P().slides.filter((_, i) => !indices.includes(i));
    const at = Math.max(0, Math.min(P().slides.length, to - before));
    P().slides.splice(at, 0, ...moving);
    SL.cur = P().slides.findIndex(s => s.id === curId);
    E.railSel = new Set(moving.map(m => P().slides.indexOf(m)));
    SL.commit(); E.renderAll();
  };
  E.moveSlideBy = (d) => {
    const idx = [...E.railSel].sort((a, b) => a - b); if (!idx.length) return;
    if (d < 0 && idx[0] === 0) return; if (d > 0 && idx[idx.length - 1] === P().slides.length - 1) return;
    E.moveSlides(idx, d < 0 ? idx[0] - 1 : idx[idx.length - 1] + 2);
  };
  E.applyLayout = (layout, slide = S()) => {
    // keep content of placeholders: map existing placeholder content by role into new layout's placeholders
    const specs = SL.layoutPlaceholders(layout, SL.W(P()), SL.H(P()), P());
    const oldPh = slide.elements.filter(e => e.ph); const others = slide.elements.filter(e => !e.ph);
    const pool = { title: oldPh.filter(e => e.ph === 'title'), body: oldPh.filter(e => e.ph === 'body' || e.ph === 'subtitle'), pic: oldPh.filter(e => e.ph === 'pic') };
    const out = specs.map(sp => {
      const key = sp.ph === 'subtitle' ? 'body' : sp.ph; const prev = (pool[key] || []).shift();
      const n = SL.newElement(sp.type, sp);
      if (prev) { if (prev.type === 'text') { n.html = prev.html; } else if (prev.type === 'image') { n.src = prev.src; n.crop = prev.crop; } }
      return n;
    });
    // leftover placeholders with content become plain objects
    Object.values(pool).flat().forEach(e => { if ((e.type === 'text' && !SL.isEmptyText(e)) || (e.type === 'image' && e.src)) { delete e.ph; delete e.placeholder; others.push(e); } });
    slide.layout = layout; slide.elements = [...out, ...others];
  };
  E.resetSlide = (slide = S()) => {
    const specs = SL.layoutPlaceholders(slide.layout, SL.W(P()), SL.H(P()), P()); const used = new Set();
    for (const el of slide.elements) {
      if (!el.ph) continue;
      const sp = specs.find((s, i) => !used.has(i) && (s.ph === el.ph || (s.ph === 'subtitle' && el.ph === 'body')) && (used.add(i) || true));
      if (sp) Object.assign(el, { x: sp.x, y: sp.y, w: sp.w, h: sp.h, rot: 0 });
      ['font', 'size', 'color', 'bold', 'italic', 'underline', 'align', 'fill', 'stroke', 'effect', 'lh'].forEach(k => delete el[k]);
      if (sp) Object.assign(el, { size: sp.size, align: sp.align || 'left', valign: sp.valign, bold: sp.bold });
      if (el.type === 'text') el.fill = 'none', el.stroke = 'none';
    }
  };

  /* ---------- rail interactions ---------- */
  function railDown(ev) {
    const it = ev.target.closest('.rail-item, .sorter-item'); const secH = ev.target.closest('.rail-sec, .sorter-sec');
    if (secH && ev.button === 0) { const i = P().slides.findIndex(s => s.id === secH.dataset.sec); E.selectSection(i); return; }
    if (!it) return;
    const i = +it.dataset.i;
    if (ev.button === 2) { if (!E.railSel.has(i)) { E.railSel = new Set([i]); E.goto(i, { keepRailSel: true }); } return; }
    if (ev.button !== 0) return;
    if (ev.shiftKey) { const a = Math.min(SL.cur, i), b = Math.max(SL.cur, i); E.railSel = new Set(Array.from({ length: b - a + 1 }, (_, k) => a + k)); E.renderRail(); if (E.view === 'sorter') E.renderSorter(); return; }
    if (ev.ctrlKey || ev.metaKey) { if (E.railSel.has(i) && E.railSel.size > 1) E.railSel.delete(i); else E.railSel.add(i); E.renderRail(); if (E.view === 'sorter') E.renderSorter(); return; }
    if (!E.railSel.has(i)) E.railSel = new Set([i]);
    const container = ev.currentTarget; const grid = container.id === 'sorter';
    const x0 = ev.clientX, y0 = ev.clientY; let dragging = false; let marker = null; let drop = -1;
    const onMove = (e2) => {
      if (!dragging && Math.hypot(e2.clientX - x0, e2.clientY - y0) < 6) return;
      if (!dragging) { dragging = true; marker = document.createElement('div'); marker.className = 'drop-marker' + (grid ? ' v' : ''); container.appendChild(marker); container.classList.add('dragging'); }
      const items = $$(grid ? '.sorter-item' : '.rail-item', container); const cr = container.getBoundingClientRect();
      let best = null;
      for (const n of items) {
        const r = n.getBoundingClientRect();
        if (grid) { if (e2.clientY >= r.top - 12 && e2.clientY <= r.bottom + 12) { const after = e2.clientX > r.left + r.width / 2; const d = Math.abs(e2.clientX - (after ? r.right : r.left)); if (!best || d < best.d) best = { d, n, after, r }; } }
        else { const after = e2.clientY > r.top + r.height / 2; const d = Math.abs(e2.clientY - (r.top + r.height / 2)); if (!best || d < best.d) best = { d, n, after, r }; }
      }
      if (!best) return;
      drop = +best.n.dataset.i + (best.after ? 1 : 0);
      if (grid) { marker.style.left = (best.after ? best.r.right + 6 : best.r.left - 8) - cr.left + container.scrollLeft + 'px'; marker.style.top = best.r.top - cr.top + container.scrollTop + 'px'; marker.style.height = best.r.height + 'px'; }
      else { marker.style.top = (best.after ? best.r.bottom + 3 : best.r.top - 4) - cr.top + container.scrollTop + 'px'; }
      // auto-scroll
      if (e2.clientY < cr.top + 30) container.scrollTop -= 12; else if (e2.clientY > cr.bottom - 30) container.scrollTop += 12;
    };
    const onUp = () => {
      removeEventListener('pointermove', onMove); removeEventListener('pointerup', onUp);
      if (marker) marker.remove(); container.classList.remove('dragging');
      if (dragging && drop >= 0) { E.moveSlides([...E.railSel], drop); return; }
      if (!dragging) { E.railSel = new Set([i]); E.goto(i); }
    };
    addEventListener('pointermove', onMove); addEventListener('pointerup', onUp);
  }
  E.selectSection = (i) => {
    const secs = P().sections || []; let end = P().slides.length;
    for (let k = i + 1; k < P().slides.length; k++) if (secs.some(s => s.start === P().slides[k].id)) { end = k; break; }
    E.railSel = new Set(Array.from({ length: end - i }, (_, k) => i + k)); SL.cur = i; E.sel = []; E.renderAll();
  };

  /* ---------- sorter ---------- */
  E.renderSorter = () => {
    const so = $('#sorter'); const st = so.scrollTop; so.innerHTML = '';
    const tw = Math.round(220 * (E.sorterZoom || 1));
    let grid = document.createElement('div'); grid.className = 'sorter-grid'; so.appendChild(grid);
    P().slides.forEach((s, i) => {
      const sec = (P().sections || []).find(x => x.start === s.id);
      if (sec) { const h = document.createElement('div'); h.className = 'sorter-sec'; h.dataset.sec = s.id; h.innerHTML = `${icon('rows-3', 14)} <span>${BTS.esc(sec.name)}</span> <small>${countSection(i)} slides</small>`; so.appendChild(h); grid = document.createElement('div'); grid.className = 'sorter-grid'; so.appendChild(grid); }
      const it = document.createElement('div'); it.className = 'sorter-item' + (i === SL.cur ? ' cur' : '') + (E.railSel.has(i) ? ' sel' : '') + (s.hidden ? ' hidden-s' : ''); it.dataset.i = i;
      it.appendChild(SL.thumb(s, P(), tw, i));
      it.insertAdjacentHTML('beforeend', `<div class="sorter-meta"><span class="n">${i + 1}</span>${s.hidden ? icon('eye-off', 13) : ''}${s.transition && s.transition.type !== 'none' ? `<span title="Transition: ${s.transition.type}">${icon('sparkles', 13)}</span>` : ''}${s.transition?.auto ? `<span class="t">${s.transition.auto}s</span>` : ''}</div>`);
      grid.appendChild(it);
    });
    so.querySelectorAll('.sorter-grid:empty').forEach(g => g.remove());
    so.scrollTop = st;
  };
  function countSection(i) { const secs = P().sections || []; let n = 1; for (let k = i + 1; k < P().slides.length && !secs.some(s => s.start === P().slides[k].id); k++) n++; return n; }

  /* ---------- notes page view ---------- */
  E.renderNotesView = () => {
    const nv = $('#notes-view'); nv.innerHTML = '';
    const page = document.createElement('div'); page.className = 'notes-page';
    const s = S(); const tw = 560;
    page.appendChild(SL.thumb(s, P(), tw, SL.cur));
    const ta = document.createElement('div'); ta.className = 'notes-page-text'; ta.contentEditable = 'plaintext-only'; ta.spellcheck = SL.spell !== false; ta.dataset.placeholder = 'Click to add notes'; ta.textContent = s.notes || '';
    ta.addEventListener('input', () => { S().notes = ta.innerText.replace(/\n$/, ''); BTS.markDirty(); });
    ta.addEventListener('blur', () => SL.commit());
    page.appendChild(ta);
    page.insertAdjacentHTML('beforeend', `<div class="notes-page-num">${SL.cur + 1}</div>`);
    nv.appendChild(page);
  };

  E.setView = (v) => {
    if (E.editing) E.endEdit();
    E.view = v;
    $('#center').classList.toggle('hidden', v !== 'normal');
    $('#rail').classList.toggle('hidden', v === 'sorter');
    $('#sorter').classList.toggle('hidden', v !== 'sorter');
    $('#notes-view').classList.toggle('hidden', v !== 'notes');
    if (v === 'sorter') { E._fitting = true; BTS.setZoom(Math.round((E.sorterZoom || 1) * 100)); E._fitting = false; E.renderSorter(); $('#sorter').focus(); }
    if (v === 'notes') E.renderNotesView();
    if (v === 'normal') { E.renderAll(); requestAnimationFrame(() => { E.layout(); if (!E.fit) { E._fitting = true; BTS.setZoom(Math.round(E.zoom * 100)); E._fitting = false; } }); }
    E.updateStatus(); E.updateRibbon();
  };

  /* ====================================================================== */
  /*  Wiring                                                                 */
  /* ====================================================================== */
  E.wire = () => {
    const stage = $('#stage');
    stage.addEventListener('pointerdown', onStageDown);
    stage.addEventListener('dblclick', (ev) => {
      if (E.editing) return;
      const elNode = ev.target.closest('.sl-slide.editor > .el'); if (!elNode) return;
      const el = SL.findEl(elNode.dataset.id); if (!el) return;
      if (el.type === 'image') { if (!el.src) BTS.dispatch('insert-picture'); else E.startCrop(el.id); return; }
      if (el.type === 'chart') { BTS.dispatch('chart-edit'); return; }
      E.startEdit(el.id, { at: ev });
    });
    stage.addEventListener('contextmenu', e => e.preventDefault());
    const rail = $('#rail'), sorter = $('#sorter');
    [rail, sorter].forEach(c => {
      c.addEventListener('pointerdown', railDown);
      c.addEventListener('contextmenu', ev => { ev.preventDefault(); E.railContext && E.railContext(ev); });
    });
    sorter.addEventListener('dblclick', ev => { const it = ev.target.closest('.sorter-item'); if (it) { SL.cur = +it.dataset.i; E.setView('normal'); } });
    rail.addEventListener('dblclick', ev => { if (!ev.target.closest('.rail-item')) E.newSlide(); });
    const notes = $('#notes');
    try { notes.contentEditable = 'plaintext-only'; } catch { }
    notes.addEventListener('input', () => { S().notes = notes.innerText.replace(/\n$/, ''); BTS.markDirty(); });
    notes.addEventListener('blur', () => SL.commit({ silent: true }));
    // notes splitter
    $('#notes-split').addEventListener('pointerdown', ev => {
      const wrap = $('#notes-wrap'); const h0 = wrap.offsetHeight, y0 = ev.clientY;
      const mv = (e2) => { wrap.style.height = Math.max(40, Math.min(innerHeight * .6, h0 - (e2.clientY - y0))) + 'px'; E.layout(); };
      const up = () => { removeEventListener('pointermove', mv); removeEventListener('pointerup', up); };
      addEventListener('pointermove', mv); addEventListener('pointerup', up);
    });
    new ResizeObserver(() => { E.layout(); }).observe(stage);
    let rw = 0; new ResizeObserver(() => { const w = rail.clientWidth; if (Math.abs(w - rw) > 4) { rw = w; E.renderRail(); } }).observe(rail);
    stage.addEventListener('scroll', () => E.renderRulers());
    stage.addEventListener('wheel', (e) => { if (e.ctrlKey) { e.preventDefault(); E.fit = false; BTS.setZoom(BTS.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1)); } }, { passive: false });
    // drag-over images
    stage.addEventListener('dragover', e => { if (e.dataTransfer.types.includes('Files')) e.preventDefault(); });
  };
  function onStageContext(ev) { ev.preventDefault(); E.canvasContext && E.canvasContext(ev); }
})();
