/* BTS Slides — slide show: full-screen presenting, transitions, entrance animations,
 * navigation, black/white screen, laser pointer, pen annotations; plus editor previews.
 */
(function () {
  const $ = (s, r = document) => r.querySelector(s);
  const FX = window.FX = {};
  const P = () => SL.pres;
  const reduce = () => matchMedia('(prefers-reduced-motion: reduce)').matches;

  /* ---------- transitions ---------- */
  // container: positioned box holding both slides (same size). from may be null (black).
  FX.transition = (container, fromNode, toNode, tr) => {
    const type = (tr && tr.type) || 'none'; const dur = reduce() ? 1 : Math.max(50, (tr && tr.dur) || 700);
    const dir = (tr && tr.dir) || 'left';
    container.appendChild(toNode);
    if (type === 'none' || !fromNode) { if (fromNode && fromNode !== toNode) fromNode.remove(); if (type !== 'none' && !fromNode) return anim(toNode, [{ opacity: 0 }, { opacity: 1 }], dur).then(() => {}); return Promise.resolve(); }
    const ease = 'cubic-bezier(.4,0,.2,1)';
    const vec = { left: [-1, 0], right: [1, 0], up: [0, -1], down: [0, 1] }[dir] || [-1, 0];
    const off = (k) => `translate(${-vec[0] * 100 * k}%, ${-vec[1] * 100 * k}%)`;
    let ps = [];
    container.style.perspective = '';
    switch (type) {
      case 'fade': ps = [anim(toNode, [{ opacity: 0 }, { opacity: 1 }], dur)]; break;
      case 'push': ps = [anim(fromNode, [{ transform: 'translate(0,0)' }, { transform: off(-1) }], dur, ease), anim(toNode, [{ transform: off(1) }, { transform: 'translate(0,0)' }], dur, ease)]; break;
      case 'cover': ps = [anim(toNode, [{ transform: off(1) }, { transform: 'translate(0,0)' }], dur, ease)]; break;
      case 'reveal': container.insertBefore(toNode, fromNode); ps = [anim(fromNode, [{ transform: 'translate(0,0)' }, { transform: off(-1) }], dur, ease)]; break;
      case 'wipe': { const from = { left: 'inset(0 0 0 100%)', right: 'inset(0 100% 0 0)', up: 'inset(100% 0 0 0)', down: 'inset(0 0 100% 0)' }[dir]; ps = [anim(toNode, [{ clipPath: from }, { clipPath: 'inset(0 0 0 0)' }], dur, ease)]; break; }
      case 'split': ps = [anim(toNode, [{ clipPath: dir === 'up' || dir === 'down' ? 'inset(50% 0 50% 0)' : 'inset(0 50% 0 50%)' }, { clipPath: 'inset(0 0% 0 0%)' }], dur, ease)]; break;
      case 'zoom': ps = [anim(toNode, [{ transform: 'scale(.25)', opacity: 0 }, { transform: 'scale(1)', opacity: 1 }], dur, ease), anim(fromNode, [{ opacity: 1 }, { opacity: 0 }], dur)]; break;
      case 'flip': {
        container.style.perspective = '2000px'; toNode.style.visibility = 'hidden';
        ps = [anim(fromNode, [{ transform: 'rotateY(0)' }, { transform: `rotateY(${vec[0] > 0 ? 90 : -90}deg)` }], dur / 2, 'ease-in').then(() => { fromNode.style.visibility = 'hidden'; toNode.style.visibility = ''; return anim(toNode, [{ transform: `rotateY(${vec[0] > 0 ? -90 : 90}deg)` }, { transform: 'rotateY(0)' }], dur / 2, 'ease-out'); })];
        break;
      }
      case 'cube': {
        container.style.perspective = '2400px'; const s = vec[0] >= 0 ? 1 : -1;
        fromNode.style.transformOrigin = s < 0 ? '100% 50%' : '0% 50%'; toNode.style.transformOrigin = s < 0 ? '0% 50%' : '100% 50%';
        ps = [anim(fromNode, [{ transform: 'translateX(0) rotateY(0)' }, { transform: `translateX(${-s * -100}%) rotateY(${s * 90}deg)`.replace('--', '') }], dur, ease), anim(toNode, [{ transform: `translateX(${s * -100}%) rotateY(${-s * 90}deg)` }, { transform: 'translateX(0) rotateY(0)' }], dur, ease)];
        break;
      }
      case 'blur': ps = [anim(fromNode, [{ filter: 'blur(0)', opacity: 1 }, { filter: 'blur(24px)', opacity: 0 }], dur), anim(toNode, [{ filter: 'blur(24px)', opacity: 0 }, { filter: 'blur(0)', opacity: 1 }], dur)]; break;
      default: ps = [anim(toNode, [{ opacity: 0 }, { opacity: 1 }], dur)];
    }
    return Promise.all(ps).then(() => { if (fromNode.parentNode) fromNode.remove(); toNode.style.transform = ''; toNode.style.transformOrigin = ''; container.style.perspective = ''; });
  };
  function anim(node, kf, dur, easing = 'ease') {
    return new Promise(res => { const a = node.animate(kf, { duration: dur, easing, fill: 'forwards' }); a.onfinish = () => { try { if (node.isConnected) a.commitStyles(); } catch { } a.cancel(); res(); }; a.oncancel = () => res(); FX._running.push(a); });
  }
  FX._running = [];
  FX.finishAll = () => { FX._running.forEach(a => { try { a.finish(); } catch { } }); FX._running = []; };

  /* ---------- element animations ---------- */
  FX.EMPHASIS = new Set(['pulse']);
  // steps: [[{el, start(ms)}...], ...] — each step plays on one click (or automatically for after/with first)
  FX.buildSteps = (slide) => {
    const list = CMD.animOrder(slide); const steps = [];
    let cur = null, prevEnd = 0, prevStart = 0;
    list.forEach((el, i) => {
      const a = el.anim; const trig = a.trigger || 'click'; const dur = a.dur || 500, delay = a.delay || 0;
      if (trig === 'click' || !cur) { cur = { auto: i === 0 && trig !== 'click', items: [] }; steps.push(cur); prevEnd = 0; prevStart = 0; }
      let start = trig === 'with' ? prevStart : trig === 'after' ? prevEnd : 0;
      start += delay; cur.items.push({ el, start, dur });
      prevStart = start; prevEnd = Math.max(prevEnd, start + dur);
    });
    return steps;
  };
  FX.hiddenAtStart = (slide) => new Set(CMD.animOrder(slide).filter(e => !FX.EMPHASIS.has(e.anim.type)).map(e => e.id));
  FX.playStep = (slideNode, step, k = 1) => Promise.all(step.items.map(it => new Promise(res => {
    const node = slideNode.querySelector(`.el[data-id="${it.el.id}"]`); if (!node) return res();
    setTimeout(() => { animateEl(node, it.el, it.dur * k).then(res); }, it.start * k);
  })));
  function animateEl(node, el, dur) {
    const a = el.anim; const base = el.rot ? ` rotate(${el.rot}deg)` : ''; const dir = a.dir || 'bottom';
    node.style.visibility = '';
    if (reduce()) dur = 1;
    const W = SL.W(P()), H = SL.H(P());
    const flyFrom = { bottom: `translate(0, ${H - el.y + 20}px)`, top: `translate(0, ${-(el.y + el.h + 20)}px)`, left: `translate(${-(el.x + el.w + 20)}px, 0)`, right: `translate(${W - el.x + 20}px, 0)` }[dir] || `translate(0, ${H}px)`;
    const T = (t) => ({ transform: t + base });
    switch (a.type) {
      case 'appear': return anim(node, [{ opacity: 1 }, { opacity: 1 }], 1);
      case 'fade': return anim(node, [{ opacity: 0 }, { opacity: 1 }], dur);
      case 'fly': return anim(node, [T(flyFrom), T('translate(0,0)')], dur, 'cubic-bezier(.2,.7,.2,1)');
      case 'float': return anim(node, [Object.assign(T(`translate(0, ${dir === 'top' ? -60 : 60}px)`), { opacity: 0 }), Object.assign(T('translate(0,0)'), { opacity: 1 })], dur, 'cubic-bezier(.2,.7,.2,1)');
      case 'zoom': return anim(node, [Object.assign(T('scale(.3)'), { opacity: 0 }), Object.assign(T('scale(1)'), { opacity: 1 })], dur, 'cubic-bezier(.2,.7,.3,1.2)');
      case 'wipe': { const from = { left: 'inset(0 100% 0 0)', right: 'inset(0 0 0 100%)', top: 'inset(0 0 100% 0)', bottom: 'inset(100% 0 0 0)' }[dir] || 'inset(0 100% 0 0)'; return anim(node, [{ clipPath: from }, { clipPath: 'inset(0 0 0 0)' }], dur); }
      case 'split': return anim(node, [{ clipPath: 'inset(0 50% 0 50%)' }, { clipPath: 'inset(0 0% 0 0%)' }], dur);
      case 'bounce': return anim(node, [Object.assign(T(`translate(0, ${-(el.y + el.h)}px)`), { offset: 0 }), Object.assign(T('translate(0,0)'), { offset: .45, easing: 'ease-out' }), Object.assign(T('translate(0,-60px)'), { offset: .62, easing: 'ease-in' }), Object.assign(T('translate(0,0)'), { offset: .78, easing: 'ease-out' }), Object.assign(T('translate(0,-18px)'), { offset: .89, easing: 'ease-in' }), Object.assign(T('translate(0,0)'), { offset: 1 })], dur * 1.4, 'ease-in');
      case 'grow': return anim(node, [{ transform: `scale(.1) rotate(${(el.rot || 0) - 90}deg)`, opacity: 0 }, { transform: `scale(1) rotate(${el.rot || 0}deg)`, opacity: 1 }], dur, 'cubic-bezier(.2,.7,.2,1)');
      case 'swivel': return anim(node, [Object.assign(T('perspective(800px) rotateY(90deg)'), { opacity: 0 }), Object.assign(T('perspective(800px) rotateY(0)'), { opacity: 1 })], dur);
      case 'pulse': return anim(node, [T('scale(1)'), T('scale(1.12)'), T('scale(1)')], dur);
      default: return anim(node, [{ opacity: 0 }, { opacity: 1 }], dur);
    }
  }

  /* ====================================================================== */
  /*  Slide show                                                             */
  /*  Options: { windowed, allowHidden, presenter, rehearse, order:[slide indices], useTimings, loop, kiosk }  */
  /*  Presenter view is drawn inside the window (current slide, next slide, notes, timer, thumbnails): the   */
  /*  main process denies window.open, so a separate audience window/display needs a main.js change.        */
  /* ====================================================================== */
  const show = FX.show = { on: false };
  const fmt = (ms) => { const t = Math.max(0, Math.round(ms / 1000)); const h = Math.floor(t / 3600), m = Math.floor(t / 60) % 60, s = t % 60; return (h ? h + ':' : '') + String(m).padStart(2, '0') + ':' + String(s).padStart(2, '0'); };
  const setupOf = () => P().showSetup || {};
  FX.start = async (from = 0, opts = {}) => {
    if (show.on) return;
    if (E.editing) E.endEdit();
    if (E.view === 'master') { BTS.toast('Close Slide Master view first'); return; }
    const setup = setupOf();
    const vis = P().slides.map((s, i) => i).filter(i => !P().slides[i].hidden);
    let order = opts.order ? opts.order.filter(i => i >= 0 && i < P().slides.length) : vis.slice();
    if (!order.length) { BTS.toast(opts.order ? 'This custom show has no slides' : 'All slides are hidden'); return; }
    if (!opts.order && P().slides[from] && P().slides[from].hidden) { if (opts.allowHidden) { order.push(from); order.sort((a, b) => a - b); } else from = vis.find(i => i >= from) ?? vis[vis.length - 1]; }
    if (!order.includes(from)) from = order[0];
    const kiosk = !!(opts.kiosk ?? setup.kiosk);
    Object.assign(show, {
      on: true, idx: from, step: 0, steps: [], windowed: !!opts.windowed, ended: false, num: '', pen: null, strokes: {}, order,
      loop: kiosk || !!(opts.loop ?? $('#show-loop')?.checked), timings: !opts.rehearse && (kiosk || (opts.useTimings ?? $('#show-timings')?.checked !== false)),
      anims: !setup.noAnims && $('#show-anims')?.checked !== false, kiosk, rehearse: !!opts.rehearse, narrate: !!opts.narrate, presenter: !!(opts.presenter ?? (setup.presenter && !opts.windowed)),
      elapsed: 0, paused: false, rec: {}, visits: [], last: performance.now(), cover: null, gridOpen: false,
    });
    await SL.ensureFonts(P());
    const root = document.createElement('div'); root.id = 'show'; root.tabIndex = 0;
    if (opts.windowed) root.classList.add('windowed'); if (show.presenter) root.classList.add('presenter'); if (kiosk) root.classList.add('kiosk');
    const viewport = `<div class="show-viewport"><div class="show-box"><div class="show-stage"></div><canvas class="show-ink"></canvas></div><div class="show-cover"></div></div>`;
    const tools = `<button data-a="prev" title="Previous (←)">${icon('chevron-left', 18)}</button><span class="show-count"></span><button data-a="next" title="Next (→ or Space)">${icon('chevron-right', 18)}</button>
      <span class="show-sep"></span><button data-a="laser" title="Laser pointer (L)">${icon('crosshair', 17)}</button><button data-a="pen" title="Pen (P)">${icon('pen-line', 17)}</button><button data-a="hl" title="Highlighter (H)">${icon('highlighter', 17)}</button><button data-a="erase" title="Erase all ink (E)">${icon('eraser', 17)}</button>
      <span class="show-sep"></span><button data-a="grid" title="See all slides (G)">${icon('layout-grid', 17)}</button><button data-a="black" title="Black screen (B)">${icon('square', 17)}</button><button data-a="white" title="White screen (W)">${icon('square-dashed', 17)}</button>`;
    if (show.presenter) {
      root.innerHTML = `<div class="pv-top"><div class="pv-timer" title="Time since the show started"><b class="pv-elapsed">00:00</b><button data-a="pause" title="Pause the timer">${icon('pause', 16)}</button><button data-a="reset" title="Restart the timer">${icon('rotate-ccw', 16)}</button></div><span class="pv-clock" title="Current time"></span><span class="pv-count"></span><span class="pv-flex"></span><button class="pv-btn" data-a="audience" title="Show only the slide (audience view). Press S to come back to presenter view.">${icon('monitor', 16)} Audience view</button><button class="pv-btn end" data-a="end" title="End slide show (Esc)">${icon('x', 16)} End show</button></div>
        <div class="pv-main">${viewport}<div class="pv-tools show-bar">${tools}</div></div>
        <div class="pv-side"><div class="pv-h">Next slide <span class="pv-nextinfo"></span></div><div class="pv-next"></div><div class="pv-h">Notes <span class="pv-flex"></span><button data-a="notes-smaller" title="Make the notes text smaller">${icon('a-arrow-down', 15)}</button><button data-a="notes-bigger" title="Make the notes text larger">${icon('a-arrow-up', 15)}</button></div><div class="pv-notes" tabindex="-1"></div></div>
        <div class="pv-strip"></div><div class="show-laser"></div><div class="show-num hidden"></div>`;
    } else {
      root.innerHTML = `${viewport}<div class="show-laser"></div><div class="show-bar">${tools}<button data-a="end" title="End show (Esc)">${icon('x', 18)}</button><span class="show-time"></span></div><div class="show-num hidden"></div>`;
    }
    if (show.rehearse) root.insertAdjacentHTML('beforeend', `<div class="rh-bar" title="Rehearsal: the time you spend on each slide is recorded"><button data-a="next" title="Next slide">${icon('arrow-right', 16)}</button><button data-a="pause" title="Pause recording">${icon('pause', 16)}</button><b class="rh-slide">00:00</b><button data-a="repeat" title="Restart the time for this slide">${icon('rotate-ccw', 16)}</button><span class="rh-total">00:00</span></div>`);
    document.body.appendChild(root); show.root = root;
    if (show.presenter) buildStrip();
    if (!opts.windowed && BTS.desktop) bts.fullscreen(true);
    show.onResize = () => fitShow(); addEventListener('resize', show.onResize);
    root.addEventListener('keydown', onKey); root.addEventListener('pointerdown', onPointerDown); root.addEventListener('pointermove', onPointerMove); root.addEventListener('wheel', onWheel, { passive: true });
    root.addEventListener('contextmenu', onContext);
    root.addEventListener('click', onBar);
    root.focus();
    requestAnimationFrame(() => fitShow());
    fitShow(); goSlide(from, { transition: false, initial: true });
    show.clock = setInterval(tick, 250); tick();
    setIdle();
  };
  function tick() {
    if (!show.on) return; const now = performance.now(); const dt = now - show.last; show.last = now;
    if (!show.paused) { show.elapsed += dt; if (show.rehearse && !show.ended) { const id = P().slides[show.idx]?.id; if (id) show.rec[id] = (show.rec[id] || 0) + dt; } }
    const r = show.root;
    const t = r.querySelector('.show-time'); if (t) t.textContent = fmt(show.elapsed);
    const pe = r.querySelector('.pv-elapsed'); if (pe) pe.textContent = fmt(show.elapsed);
    const pc = r.querySelector('.pv-clock'); if (pc) pc.textContent = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    const rs = r.querySelector('.rh-slide'); if (rs) rs.textContent = fmt(show.rec[P().slides[show.idx]?.id] || 0);
    const rt = r.querySelector('.rh-total'); if (rt) rt.textContent = fmt(show.elapsed);
    r.classList.toggle('timer-paused', !!show.paused);
    if (FX.onShowTick) FX.onShowTick(show);
  }
  FX.end = () => {
    if (!show.on) return;
    FX.finishAll(); clearTimeout(show.autoT); clearInterval(show.clock); clearTimeout(show.idleT);
    removeEventListener('resize', show.onResize);
    show.root.remove(); show.on = false;
    if (!show.windowed && BTS.desktop) bts.fullscreen(false);
    if (show.idx != null && show.idx < P().slides.length && !show.ended) E.goto(show.idx);
    $('#stage')?.focus();
    if (show.narrate && FX.onNarrateEnd) setTimeout(() => FX.onNarrateEnd(Object.assign({}, show.rec), show.elapsed, show.order.slice()), 30);
    else if (show.rehearse && FX.onRehearseEnd) setTimeout(() => FX.onRehearseEnd(Object.assign({}, show.rec), show.elapsed, show.order.slice()), 30);
  };
  function fitShow() {
    if (!show.on) return;
    const vp = show.root.querySelector('.show-viewport'); const box = show.root.querySelector('.show-box'); const W = SL.W(P()), H = SL.H(P());
    const vw = vp.clientWidth, vh = vp.clientHeight; if (!vw || !vh) return; const k = Math.min(vw / W, vh / H);
    show.k = k; box.style.width = W + 'px'; box.style.height = H + 'px';
    box.style.transform = `translate(${(vw - W * k) / 2}px, ${(vh - H * k) / 2}px) scale(${k})`;
    const cv = show.root.querySelector('.show-ink'); cv.width = Math.round(W * Math.min(2, k * devicePixelRatio)); cv.height = Math.round(H * Math.min(2, k * devicePixelRatio)); cv.style.width = W + 'px'; cv.style.height = H + 'px';
    redrawInk();
  }
  const pos = () => show.order.indexOf(show.idx);
  function counter() {
    const p = pos(); const n = show.order.length;
    const txt = p >= 0 ? `Slide ${p + 1} of ${n}` : `Slide ${show.idx + 1} (hidden)`;
    const c = show.root.querySelector('.show-count'); if (c) c.textContent = txt;
    const pc = show.root.querySelector('.pv-count'); if (pc) pc.textContent = txt;
    if (show.presenter) updatePresenter();
  }
  const visible = () => show.order;

  /* ---------- presenter view ---------- */
  function notesSize(d) { let v = 20; try { v = +localStorage.getItem('bts-pv-notes') || 20; } catch { } if (d) { v = Math.max(12, Math.min(48, v + d)); try { localStorage.setItem('bts-pv-notes', v); } catch { } } return v; }
  function buildStrip() {
    const st = show.root.querySelector('.pv-strip'); if (!st) return; st.innerHTML = '';
    P().slides.forEach((s, i) => { const b = document.createElement('button'); b.className = 'pv-th' + (s.hidden ? ' hid' : '') + (show.order.includes(i) ? '' : ' out'); b.dataset.go = i; b.title = `Go to slide ${i + 1}`; b.appendChild(SL.thumb(s, P(), 128, i)); b.insertAdjacentHTML('beforeend', `<span>${i + 1}</span>`); st.appendChild(b); });
  }
  function updatePresenter() {
    const r = show.root; const p = pos(); const nextIdx = p >= 0 && p + 1 < show.order.length ? show.order[p + 1] : (show.loop && show.order.length ? show.order[0] : null);
    const nx = r.querySelector('.pv-next'); nx.innerHTML = '';
    const w = Math.max(160, nx.clientWidth || 340);
    if (show.ended) nx.innerHTML = '<div class="pv-end">End of slide show</div>';
    else if (nextIdx != null) nx.appendChild(SL.thumb(P().slides[nextIdx], P(), w, nextIdx)); else nx.innerHTML = '<div class="pv-end">End of slide show</div>';
    const left = show.steps.length - show.step;
    r.querySelector('.pv-nextinfo').textContent = !show.ended && left > 0 ? `· ${left} animation${left > 1 ? 's' : ''} left on this slide` : '';
    const notes = r.querySelector('.pv-notes'); const txt = (P().slides[show.idx]?.notes || '').trim();
    notes.style.fontSize = notesSize() + 'px'; notes.innerHTML = txt ? BTS.esc(txt).replace(/\n/g, '<br>') : '<span class="pv-nonotes">No notes for this slide.</span>'; notes.scrollTop = 0;
    r.querySelectorAll('.pv-th').forEach(b => b.classList.toggle('cur', +b.dataset.go === show.idx));
    const cur = r.querySelector('.pv-th.cur'); if (cur) cur.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  function goSlide(i, opts = {}) {
    FX.finishAll(); clearTimeout(show.autoT);
    const stage = show.root.querySelector('.show-stage');
    const slide = P().slides[i]; show.idx = i; show.ended = false;
    if (!show.order.includes(i) && !show.customOrder) { /* jumped to a hidden slide: keep navigating from here */ }
    show.visits.push(i);
    if (FX.onShowSlide) FX.onShowSlide(i, show);
    show.steps = show.anims ? FX.buildSteps(slide) : []; show.step = 0;
    const built = opts.built || !show.anims;
    const hidden = built ? new Set() : FX.hiddenAtStart(slide);
    const node = SL.renderSlide(slide, P(), { index: i, animHidden: hidden }); node.classList.add('show-slide');
    if (built) show.step = show.steps.length;
    const old = stage.querySelector('.show-slide, .show-end');
    show.busy = true;
    FX.transition(stage, old, node, opts.transition === false ? { type: 'none' } : slide.transition).then(() => {
      show.busy = false;
      stage.querySelectorAll('.show-slide, .show-end').forEach(n => { if (n !== node) n.remove(); });
      if (!built && show.steps[0] && show.steps[0].auto) advance();
      else scheduleAuto();
    });
    counter(); redrawInk();
  }
  function scheduleAuto() {
    clearTimeout(show.autoT);
    const tr = P().slides[show.idx]?.transition;
    let auto = tr && tr.auto != null && tr.auto !== false ? +tr.auto : null;
    if (show.kiosk && auto == null) auto = 5;
    if (!show.timings || auto == null) return;
    if (show.step < show.steps.length) { show.autoT = setTimeout(() => advance(), Math.max(300, auto * 1000 / Math.max(1, show.steps.length + 1))); return; }
    show.autoT = setTimeout(() => FX.next(), Math.max(0, auto * 1000));
  }
  function advance() {
    const stage = show.root.querySelector('.show-slide'); const st = show.steps[show.step++];
    show.busy = true;
    if (show.presenter) updatePresenter();
    FX.playStep(stage, st).then(() => { show.busy = false; scheduleAuto(); });
  }
  FX.next = () => {
    if (!show.on) return;
    if (show.busy) { FX.finishAll(); return; }
    if (show.ended) return FX.end();
    if (show.step < show.steps.length) return advance();
    const p = pos(); let nx;
    if (p >= 0) nx = show.order[p + 1]; else nx = show.order.find(v => v > show.idx);
    if (nx == null) { if (show.loop) return goSlide(show.order[0]); return endScreen(); }
    goSlide(nx);
  };
  FX.prev = () => {
    if (!show.on) return; FX.finishAll();
    if (show.ended) { show.ended = false; return goSlide(show.idx, { transition: false, built: true }); }
    const p = pos(); let pv;
    if (p >= 0) pv = show.order[p - 1]; else pv = [...show.order].reverse().find(v => v < show.idx);
    if (pv == null) return;
    goSlide(pv, { transition: false, built: true });
  };
  FX.goto = (i) => { if (show.on) goSlide(i, { transition: false }); };
  function endScreen() {
    show.ended = true; clearTimeout(show.autoT);
    if (FX.onShowSlide) FX.onShowSlide(-1, show);
    const stage = show.root.querySelector('.show-stage'); const old = stage.querySelector('.show-slide');
    const n = document.createElement('div'); n.className = 'show-end'; n.style.cssText = `width:${SL.W(P())}px;height:${SL.H(P())}px`; n.innerHTML = `<div>${show.narrate ? 'End of narration. Click to finish.' : show.rehearse ? 'End of rehearsal. Click to see your timings.' : 'End of slide show. Click to exit.'}</div>`;
    FX.transition(stage, old, n, { type: 'fade', dur: 300 });
    const c = show.root.querySelector('.show-count'); if (c) c.textContent = 'End';
    const pc = show.root.querySelector('.pv-count'); if (pc) pc.textContent = 'End of show';
    if (show.presenter) updatePresenter();
  }

  /* ---------- input ---------- */
  function onKey(e) {
    const k = e.key;
    if (show.kiosk && k !== 'Escape') { e.preventDefault(); return; }
    if (show.root.classList.contains('audience-only') && (k === 's' || k === 'S')) { show.root.classList.remove('audience-only'); requestAnimationFrame(fitShow); e.preventDefault(); return; }
    if (show.cover && !['b', 'B', 'w', 'W', '.', ','].includes(k)) { setCover(null); e.preventDefault(); return; }
    if (/^[0-9]$/.test(k)) { show.num += k; const n = show.root.querySelector('.show-num'); n.textContent = 'Go to slide ' + show.num; n.classList.remove('hidden'); e.preventDefault(); return; }
    if (k === 'Enter' && show.num) { const i = Math.max(1, Math.min(P().slides.length, +show.num)) - 1; show.num = ''; show.root.querySelector('.show-num').classList.add('hidden'); goSlide(i, { transition: false }); e.preventDefault(); return; }
    if (k === 'Backspace' && show.num) { show.num = show.num.slice(0, -1); e.preventDefault(); return; }
    switch (k) {
      case 'Escape': if (show.pen) { setPen(null); break; } if (show.gridOpen) { closeGrid(); break; } FX.end(); break;
      case 'ArrowRight': case 'ArrowDown': case 'PageDown': case ' ': case 'Enter': case 'n': case 'N': FX.next(); break;
      case 'ArrowLeft': case 'ArrowUp': case 'PageUp': case 'Backspace': case 'p': case 'P':
        if ((k === 'p' || k === 'P') && !e.ctrlKey) { if (k === 'P' || e.shiftKey) { FX.prev(); break; } setPen(show.pen === 'pen' ? null : 'pen'); break; }
        if (e.ctrlKey && (k === 'p' || k === 'P')) { setPen('pen'); break; }
        FX.prev(); break;
      case 'Home': goSlide(show.order[0] ?? 0, { transition: false }); break;
      case 'End': goSlide(show.order[show.order.length - 1], { transition: false }); break;
      case 'b': case 'B': case '.': setCover(show.cover === 'black' ? null : 'black'); break;
      case 'w': case 'W': case ',': setCover(show.cover === 'white' ? null : 'white'); break;
      case 'l': case 'L': setPen(show.pen === 'laser' ? null : 'laser'); break;
      case 'h': case 'H': setPen(show.pen === 'hl' ? null : 'hl'); break;
      case 'e': case 'E': show.strokes[show.idx] = []; redrawInk(); break;
      case 'g': case 'G': case '-': openGrid(); break;
      case 'a': case 'A': show.root.classList.toggle('cursor-hidden'); break;
      case 'F5': break;
      default: return;
    }
    e.preventDefault(); e.stopPropagation();
  }
  function onWheel(e) { if (show.gridOpen || show.kiosk || e.target.closest('.pv-side, .pv-strip')) return; clearTimeout(show.wheelT); if (show.wheelLock) return; show.wheelLock = true; show.wheelT = setTimeout(() => show.wheelLock = false, 350); if (e.deltaY > 0) FX.next(); else if (e.deltaY < 0) FX.prev(); }
  function onContext(e) {
    e.preventDefault(); if (show.kiosk) return;
    BTS.menu({ x: e.clientX, y: e.clientY }, [
      { label: 'Next', icon: 'chevron-right', action: () => FX.next() }, { label: 'Previous', icon: 'chevron-left', action: () => FX.prev() },
      { label: 'See all slides', icon: 'layout-grid', action: () => openGrid() }, { sep: 1 },
      { label: 'Laser pointer', checked: show.pen === 'laser', action: () => setPen(show.pen === 'laser' ? null : 'laser') }, { label: 'Pen', checked: show.pen === 'pen', action: () => setPen(show.pen === 'pen' ? null : 'pen') }, { label: 'Highlighter', checked: show.pen === 'hl', action: () => setPen(show.pen === 'hl' ? null : 'hl') }, { label: 'Erase all ink on slide', icon: 'eraser', action: () => { show.strokes[show.idx] = []; redrawInk(); } }, { sep: 1 },
      { label: 'Black screen', icon: 'square', action: () => setCover('black') }, { label: 'White screen', icon: 'square', action: () => setCover('white') }, { sep: 1 },
      { label: show.paused ? 'Resume timer' : 'Pause timer', icon: show.paused ? 'play' : 'pause', action: () => { show.paused = !show.paused; tick(); } }, { sep: 1 },
      { label: 'End show', icon: 'x', action: () => FX.end() },
    ]);
  }
  function onBar(e) {
    const g = e.target.closest('[data-go]'); if (g) { e.stopPropagation(); goSlide(+g.dataset.go, { transition: false }); show.root.focus(); return; }
    const b = e.target.closest('[data-a]'); if (!b) return; e.stopPropagation();
    const a = b.dataset.a;
    if (a === 'prev') FX.prev(); if (a === 'next') FX.next(); if (a === 'end') FX.end();
    if (a === 'laser' || a === 'pen' || a === 'hl') setPen(show.pen === a ? null : a);
    if (a === 'erase') { show.strokes[show.idx] = []; redrawInk(); }
    if (a === 'black') setCover(show.cover === 'black' ? null : 'black'); if (a === 'white') setCover(show.cover === 'white' ? null : 'white'); if (a === 'grid') openGrid();
    if (a === 'pause') { show.paused = !show.paused; b.innerHTML = icon(show.paused ? 'play' : 'pause', 16); b.title = show.paused ? 'Resume' : 'Pause'; tick(); }
    if (a === 'reset') { show.elapsed = 0; tick(); }
    if (a === 'repeat') { const id = P().slides[show.idx]?.id; if (id) { show.elapsed -= show.rec[id] || 0; show.rec[id] = 0; } tick(); }
    if (a === 'notes-bigger' || a === 'notes-smaller') { notesSize(a === 'notes-bigger' ? 2 : -2); show.root.querySelector('.pv-notes').style.fontSize = notesSize() + 'px'; }
    if (a === 'audience') { show.root.classList.add('audience-only'); requestAnimationFrame(fitShow); BTS.toast('Audience view: press S to return to presenter view', { ms: 2200 }); }
    show.root.focus();
  }
  function setCover(c) { show.cover = c; const cv = show.root.querySelector('.show-cover'); cv.className = 'show-cover' + (c ? ' on ' + c : ''); show.root.querySelectorAll('[data-a="black"],[data-a="white"]').forEach(b => b.classList.toggle('on', b.dataset.a === c)); }
  function setPen(p) {
    show.pen = p; show.root.classList.toggle('pen-mode', p === 'pen' || p === 'hl'); show.root.classList.toggle('laser-mode', p === 'laser');
    show.root.querySelectorAll('.show-bar [data-a]').forEach(b => { if (['pen', 'hl', 'laser'].includes(b.dataset.a)) b.classList.toggle('on', b.dataset.a === p); });
    if (p !== 'laser') show.root.querySelector('.show-laser').style.display = 'none';
  }
  function toSlidePt(e) { const r = show.root.querySelector('.show-box').getBoundingClientRect(); return { x: (e.clientX - r.left) / show.k, y: (e.clientY - r.top) / show.k }; }
  function onPointerDown(e) {
    if (e.target.closest('.show-bar, .pv-top, .pv-side, .pv-strip, .rh-bar')) return;
    if (show.gridOpen) return;
    if (show.kiosk) return;
    if (show.presenter && !show.root.classList.contains('audience-only') && !e.target.closest('.show-viewport')) return;
    if (show.cover) { setCover(null); return; }
    if (e.button === 2) return;
    if (show.pen === 'pen' || show.pen === 'hl') {
      const stroke = { kind: show.pen, pts: [toSlidePt(e)] }; (show.strokes[show.idx] = show.strokes[show.idx] || []).push(stroke);
      const mv = (e2) => { stroke.pts.push(toSlidePt(e2)); redrawInk(); };
      const up = () => { removeEventListener('pointermove', mv); removeEventListener('pointerup', up); };
      addEventListener('pointermove', mv); addEventListener('pointerup', up); return;
    }
    if (show.pen === 'laser') return;
    if (e.button === 0) FX.next();
  }
  function onPointerMove(e) {
    setIdle();
    if (show.pen === 'laser') { const l = show.root.querySelector('.show-laser'); const inVp = !show.presenter || e.target.closest('.show-viewport'); l.style.display = inVp ? 'block' : 'none'; l.style.left = e.clientX + 'px'; l.style.top = e.clientY + 'px'; }
  }
  function setIdle() { show.root.classList.remove('idle'); clearTimeout(show.idleT); show.idleT = setTimeout(() => show.on && !show.presenter && show.root.classList.add('idle'), 2500); }
  function redrawInk() {
    const cv = show.root && show.root.querySelector('.show-ink'); if (!cv) return; const ctx = cv.getContext('2d'); const k = cv.width / SL.W(P());
    ctx.setTransform(1, 0, 0, 1, 0, 0); ctx.clearRect(0, 0, cv.width, cv.height); ctx.setTransform(k, 0, 0, k, 0, 0);
    for (const s of show.strokes[show.idx] || []) {
      ctx.lineCap = 'round'; ctx.lineJoin = 'round';
      ctx.strokeStyle = s.kind === 'hl' ? 'rgba(255,230,0,.45)' : '#E5243B'; ctx.lineWidth = s.kind === 'hl' ? 22 : 4;
      ctx.beginPath(); s.pts.forEach((p, i) => i ? ctx.lineTo(p.x, p.y) : ctx.moveTo(p.x, p.y)); ctx.stroke();
    }
  }
  function openGrid() {
    if (show.gridOpen) return; show.gridOpen = true;
    const g = document.createElement('div'); g.className = 'show-grid';
    g.innerHTML = `<div class="show-grid-h">All slides <span>Click a slide to go to it · Esc to return</span></div><div class="show-grid-b"></div>`;
    const b = g.querySelector('.show-grid-b');
    P().slides.forEach((s, i) => { const it = document.createElement('button'); it.className = 'sg-item' + (i === show.idx ? ' cur' : '') + (s.hidden ? ' hid' : ''); it.dataset.i = i; it.appendChild(SL.thumb(s, P(), 240, i)); it.insertAdjacentHTML('beforeend', `<span>${i + 1}</span>`); b.appendChild(it); });
    g.addEventListener('click', e => { e.stopPropagation(); const it = e.target.closest('[data-i]'); if (it) { closeGrid(); goSlide(+it.dataset.i, { transition: false }); } });
    show.root.appendChild(g);
  }
  function closeGrid() { show.gridOpen = false; show.root.querySelector('.show-grid')?.remove(); show.root.focus(); }
  FX._fmt = fmt;

  /* ====================================================================== */
  /*  Editor previews                                                        */
  /* ====================================================================== */
  FX.previewTransition = () => {
    const host = $('#slide-host'); if (!host) return;
    if (E.editing) E.endEdit();
    const box = document.createElement('div'); box.className = 'preview-box';
    const W = SL.W(P()), H = SL.H(P()); box.style.cssText = `width:${W}px;height:${H}px;transform:scale(${E.zoom});`;
    const prev = SL.cur > 0 ? SL.renderSlide(P().slides[SL.cur - 1], P(), { index: SL.cur - 1 }) : (() => { const d = document.createElement('div'); d.className = 'sl-slide'; d.style.cssText = `width:${W}px;height:${H}px;background:#000`; return d; })();
    const cur = SL.renderSlide(S(), P(), { index: SL.cur });
    box.appendChild(prev); host.appendChild(box);
    const tr = S().transition && S().transition.type !== 'none' ? S().transition : { type: 'fade', dur: 600 };
    setTimeout(() => FX.transition(box, prev, cur, tr).then(() => setTimeout(() => box.remove(), 250)), 80);
  };
  FX.previewAnimations = async () => {
    const host = $('#slide-host'); if (!host) return;
    if (E.editing) E.endEdit();
    const slide = S(); const steps = FX.buildSteps(slide); if (!steps.length) { BTS.toast('This slide has no animations'); return; }
    const box = document.createElement('div'); box.className = 'preview-box'; const W = SL.W(P()), H = SL.H(P());
    box.style.cssText = `width:${W}px;height:${H}px;transform:scale(${E.zoom});`;
    const node = SL.renderSlide(slide, P(), { index: SL.cur, animHidden: FX.hiddenAtStart(slide) }); box.appendChild(node); host.appendChild(box);
    box.addEventListener('pointerdown', (e) => { e.stopPropagation(); FX.finishAll(); box.remove(); });
    for (const st of steps) { if (!box.isConnected) return; await new Promise(r => setTimeout(r, 350)); await FX.playStep(node, st); }
    setTimeout(() => box.remove(), 700);
  };
  const S = () => SL.slide();
})();
