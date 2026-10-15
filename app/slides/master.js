/* BTS Slides 1.2 — Slide Master view (View ▸ Slide Master).
 *
 * While the view is open the editor works on a stand-in presentation: slide 1 is the master (title + body text styles,
 * background, graphics shown on every slide), the other slides are the 8 layouts (placeholder positions, per-layout
 * background and graphics). All normal editing tools work on it, with their own undo history. "Close Master View"
 * harvests the result into pres.master, moves/resizes the placeholders of every slide that still sits at its layout
 * position, and commits it as one undo step. pres.master is saved in the .pptx (BTS metadata + PptxGenJS slide masters).
 */
(function () {
  const $ = (s, r = document) => r.querySelector(s);
  const X = window.SLX = window.SLX || { cmd: {}, ribbon: [] };
  const C = () => window.CMD;
  const P = () => SL.pres;
  const near = (a, b, t = 2) => Math.abs((a || 0) - (b || 0)) <= t;

  /* Move placeholders that still sit at their old layout position to the new one (theme change, master edits). */
  SL.syncPlaceholders = (pres, before) => {
    const W = SL.W(pres), H = SL.H(pres);
    const oldP = Object.assign({}, pres, { theme: before.theme, master: before.master || null });
    for (const s of pres.slides) {
      const os = SL.layoutPlaceholders(s.layout, W, H, oldP), ns = SL.layoutPlaceholders(s.layout, W, H, pres);
      const ok = SL.phKeys(os), nk = SL.phKeys(ns); const cnt = {};
      for (const e of s.elements) {
        if (!e.ph) continue;
        const ph = e.ph === 'subtitle' ? 'subtitle' : e.ph; const k = ph + (cnt[ph] = (cnt[ph] ?? -1) + 1);
        const o = os[ok.indexOf(k)], n = ns[nk.indexOf(k)]; if (!o || !n) continue;
        if (near(e.x, o.x) && near(e.y, o.y) && near(e.w, o.w) && near(e.h, o.h)) Object.assign(e, { x: n.x, y: n.y, w: n.w, h: n.h });
        if (e.size != null && o.size != null && near(e.size, o.size, .01)) e.size = n.size;
        if ((e.align || 'left') === (o.align || 'left') && n.align) e.align = n.align;
      }
    }
  };

  const MASTER_SAMPLE = {
    title: '<p>Click to edit Master title style</p>', subtitle: '<p>Click to edit Master subtitle style</p>',
    list: '<ul><li>Edit Master text styles</li><li data-lvl="1">Second level</li><li data-lvl="2">Third level</li><li data-lvl="3">Fourth level</li></ul>', text: '<p>Edit Master text styles</p>',
  };
  function sampleFill(el) {
    if (el.type !== 'text') return;
    el.html = el.ph === 'title' ? MASTER_SAMPLE.title : el.ph === 'subtitle' ? MASTER_SAMPLE.subtitle : el.list ? MASTER_SAMPLE.list : MASTER_SAMPLE.text;
  }
  function buildMasterPres(real) {
    const W = SL.W(real), H = SL.H(real);
    const m = SL.clone(real.master || {});
    const mp = { width: W, height: H, theme: SL.clone(real.theme), footer: { date: false, number: false, text: '' }, slides: [], _masterView: true, master: { title: m.title || {}, body: m.body || {}, bg: m.bg || null, graphics: [] } };
    // the master slide: title + body placeholders at the "Title and Content" positions
    const ms = SL.newSlide(Object.assign({}, real, { master: null }), 'titleContent');
    ms.isMaster = true; ms.mname = 'Slide Master'; ms.background = m.bg ? SL.clone(m.bg) : null;
    ms.elements.forEach(el => { sampleFill(el); const r = el.ph === 'title' ? m.title : m.body; if (r) { ['font', 'color', 'bold', 'italic'].forEach(k => { if (r[k] != null) el[k] = r[k]; }); if (r.size) el.size = el.ph === 'title' ? r.size : r.size; } delete el.placeholder; });
    (m.graphics || []).forEach(g => ms.elements.push(SL.reId(SL.clone(g))));
    mp.slides.push(ms);
    for (const L of SL.LAYOUTS) {
      const s = SL.newSlide(Object.assign({}, real, { master: m }), L.id); // positions include the current layout overrides
      s.mname = L.name; s.background = (m.layoutBg || {})[L.id] ? SL.clone(m.layoutBg[L.id]) : null;
      s.elements.forEach(el => { sampleFill(el); delete el.placeholder; });
      ((m.layoutGraphics || {})[L.id] || []).forEach(g => s.elements.push(SL.reId(SL.clone(g))));
      mp.slides.push(s);
    }
    return mp;
  }
  // Read the master stand-in back into a pres.master object.
  function harvest(mp, real) {
    const W = SL.W(mp), H = SL.H(mp); const ms = mp.slides[0];
    const out = { title: {}, body: {}, layouts: {}, layoutBg: {}, layoutGraphics: {}, graphics: [] };
    const te = ms.elements.find(e => e.ph === 'title'), be = ms.elements.find(e => e.ph === 'body');
    const grab = (el, dst, defSize) => { if (!el) return; ['font', 'color', 'bold', 'italic'].forEach(k => { if (el[k] != null) dst[k] = el[k]; }); if (el.size && !near(el.size, defSize, .01)) dst.size = el.size; };
    grab(te, out.title, 40); grab(be, out.body, 24);
    if (ms.background) out.bg = SL.clone(ms.background);
    out.graphics = ms.elements.filter(e => !e.ph).map(e => SL.clone(e));
    const sizeOnly = Object.assign({}, real, { master: { title: out.title, body: out.body } });
    mp.slides.slice(1).forEach((s, i) => {
      const L = SL.LAYOUTS[i].id;
      const def = SL.layoutPlaceholders(L, W, H, sizeOnly); const keys = SL.phKeys(def); const cnt = {};
      const lo = {};
      s.elements.filter(e => e.ph).forEach(e => {
        const k = e.ph + (cnt[e.ph] = (cnt[e.ph] ?? -1) + 1); const d = def[keys.indexOf(k)]; if (!d) return;
        const o = {};
        if (!(near(e.x, d.x, 1) && near(e.y, d.y, 1) && near(e.w, d.w, 1) && near(e.h, d.h, 1))) Object.assign(o, { x: e.x / W, y: e.y / H, w: e.w / W, h: e.h / H });
        if (e.type === 'text' && e.size && d.size && !near(e.size, d.size, .01)) o.size = e.size;
        if (e.type === 'text' && e.align && d.align && e.align !== d.align) o.align = e.align;
        if (Object.keys(o).length) lo[k] = o;
      });
      if (Object.keys(lo).length) out.layouts[L] = lo;
      if (s.background) out.layoutBg[L] = SL.clone(s.background);
      const g = s.elements.filter(e => !e.ph).map(e => SL.clone(e)); if (g.length) out.layoutGraphics[L] = g;
    });
    ['title', 'body'].forEach(k => { if (!Object.keys(out[k]).length) delete out[k]; });
    ['layouts', 'layoutBg', 'layoutGraphics'].forEach(k => { if (!Object.keys(out[k]).length) delete out[k]; });
    if (!out.graphics.length) delete out.graphics;
    return Object.keys(out).length ? out : null;
  }
  // Live preview while editing: master text styles / background follow the master slide; master title & body
  // geometry changes flow down to layouts whose placeholders were still at the previous master position.
  let lastMasterGeom = null;
  function liveUpdate() {
    const mp = SL.pres; if (!mp || !mp._masterView) return;
    const ms = mp.slides[0]; const te = ms.elements.find(e => e.ph === 'title'), be = ms.elements.find(e => e.ph === 'body');
    const st = (el) => { const o = {}; if (el) ['font', 'color', 'bold', 'italic'].forEach(k => { if (el[k] != null) o[k] = el[k]; }); return o; };
    mp.master.title = st(te); mp.master.body = st(be); mp.master.bg = ms.background || null;
    const g = { t: te && { x: te.x, y: te.y, w: te.w, h: te.h, size: te.size }, b: be && { x: be.x, y: be.y, w: be.w, h: be.h, size: be.size } };
    if (lastMasterGeom) {
      const flow = (prev, cur, layouts, ph, idx) => {
        if (!prev || !cur || (near(prev.x, cur.x, .5) && near(prev.y, cur.y, .5) && near(prev.w, cur.w, .5) && near(prev.h, cur.h, .5))) return;
        mp.slides.slice(1).forEach((s, i) => {
          if (!layouts.includes(SL.LAYOUTS[i].id)) return;
          const e = s.elements.filter(x => x.ph === ph)[idx]; if (!e) return;
          if (near(e.x, prev.x) && near(e.y, prev.y) && near(e.w, prev.w) && near(e.h, prev.h)) Object.assign(e, { x: cur.x, y: cur.y, w: cur.w, h: cur.h });
        });
      };
      flow(lastMasterGeom.t, g.t, ['titleContent', 'twoContent', 'titleOnly', 'comparison'], 'title', 0);
      flow(lastMasterGeom.b, g.b, ['titleContent'], 'body', 0);
      // master text sizes flow to every layout placeholder still at the size the previous master gave it
      const sizeFlow = (role, ps, cs) => {
        if (!ps || !cs || near(ps, cs, .01)) return;
        const W = SL.W(mp), H = SL.H(mp);
        mp.slides.slice(1).forEach((s, i) => {
          const L = SL.LAYOUTS[i].id; const mk = (sz) => SL.layoutPlaceholders(L, W, H, Object.assign({}, mp, { master: { [role]: { size: sz } } }));
          const a = mk(ps), b = mk(cs), ka = SL.phKeys(a); const cnt = {};
          s.elements.forEach(e => { if (!e.ph) return; const k = e.ph + (cnt[e.ph] = (cnt[e.ph] ?? -1) + 1); const j = ka.indexOf(k); if (j < 0 || (role === 'title') !== (e.ph === 'title')) return; if (near(e.size, a[j].size, .01)) e.size = b[j].size; });
        });
      };
      sizeFlow('title', lastMasterGeom.t && lastMasterGeom.t.size, g.t && g.t.size);
      sizeFlow('body', lastMasterGeom.b && lastMasterGeom.b.size, g.b && g.b.size);
    }
    lastMasterGeom = SL.clone(g);
  }

  /* ---------------- enter / leave ---------------- */
  let saved = null;
  X.masterOpen = () => {
    if (E.view === 'master') return;
    if (FX.show.on) return;
    if (E.editing) E.endEdit();
    if (E.view !== 'normal') E.setView('normal');
    const real = SL.pres;
    saved = { pres: real, cur: SL.cur, stack: SL.history.stack, idx: SL.history.idx, notesOn: E.notesOn, pane: E.pane };
    const mp = buildMasterPres(real);
    SL.pres = mp; SL.cur = Math.max(0, SL.LAYOUTS.findIndex(l => l.id === real.slides[saved.cur]?.layout) + 1);
    E.sel = []; E.railSel = new Set([SL.cur]); SL.resetHistory();
    E.view = 'master'; document.body.classList.add('master-view'); E.notesOn = false;
    $('#tab-master').classList.remove('hidden');
    lastMasterGeom = null; liveUpdate();
    if (E.pane && !['bg', 'shape', 'selection'].includes(E.pane)) E.closePane();
    BTS.showTab('master'); E.renderAll(); requestAnimationFrame(E.layout);
  };
  function leave() {
    document.body.classList.remove('master-view'); $('#tab-master').classList.add('hidden');
    E.view = 'normal'; E.notesOn = saved ? saved.notesOn : true; E.sel = [];
  }
  X.masterClose = () => {
    if (E.view !== 'master' || !saved) return;
    if (E.editing) E.endEdit();
    liveUpdate();
    const mp = SL.pres, real = saved.pres;
    const before = { theme: SL.clone(real.theme), master: real.master ? SL.clone(real.master) : null };
    const nm = harvest(mp, real);
    SL.pres = real; SL.history.stack = saved.stack; SL.history.idx = saved.idx; SL.cur = Math.min(saved.cur, real.slides.length - 1);
    real.theme = SL.clone(mp.theme);
    if (nm) real.master = nm; else delete real.master;
    SL.syncPlaceholders(real, before);
    leave(); saved = null; E.railSel = new Set([SL.cur]);
    SL.commit(); C().buildThemeGallery(); BTS.showTab('view');
    E.renderAll(); requestAnimationFrame(E.layout);
    SL.ensureFonts(real).then(() => E.renderAll());
  };
  X.masterAbort = () => { if (E.view !== 'master' || !saved) return; SL.pres = saved.pres; SL.history.stack = saved.stack; SL.history.idx = saved.idx; leave(); saved = null; };
  X.masterBlocks = (cmd) => {
    const blocked = ['new-slide', 'new-slide-menu', 'duplicate-slide', 'layout-menu', 'reset-slide', 'hide-slide', 'section-menu', 'reuse-slides', 'apply-theme', 'apply-variant', 'themes-more',
      'present-start', 'present-current', 'present-window', 'present-presenter', 'rehearse', 'custom-shows', 'custom-shows-menu', 'view-sorter', 'view-notes', 'slide-size', 'header-footer', 'insert-slidenum', 'trans-apply-all', 'set-transition', 'set-anim', 'new-comment'];
    if (blocked.includes(cmd)) { BTS.toast('Close Slide Master view first (Slide Master ▸ Close Master View)'); return true; }
    if (cmd === 'view-normal') { X.masterClose(); return true; }
    if (cmd === 'reset-slide') { X.cmd['master-reset-layout'](); return true; }
    return false;
  };
  E.masterGuard = () => { if (E.view === 'master') { BTS.toast('Not available in Slide Master view'); return true; } return false; };

  /* ---------------- commands ---------------- */
  X.cmd['view-master'] = () => E.view === 'master' ? X.masterClose() : X.masterOpen();
  X.cmd['master-close'] = () => X.masterClose();
  X.cmd['master-reset-layout'] = () => {
    if (E.view !== 'master') return;
    const s = SL.slide(); if (s.isMaster) { const fresh = SL.newSlide(Object.assign({}, saved.pres, { master: null }), 'titleContent'); s.elements.forEach(e => { if (!e.ph) return; const f = fresh.elements.find(x => x.ph === e.ph); if (f) Object.assign(e, { x: f.x, y: f.y, w: f.w, h: f.h, size: f.size }); ['font', 'color', 'bold', 'italic'].forEach(k => delete e[k]); }); s.background = null; }
    else { const L = SL.LAYOUTS[SL.cur - 1].id; const sizeOnly = Object.assign({}, saved.pres, { master: SL.pres.master }); const def = SL.layoutPlaceholders(L, SL.W(SL.pres), SL.H(SL.pres), sizeOnly, {}); const keys = SL.phKeys(def); const cnt = {}; s.elements.forEach(e => { if (!e.ph) return; const k = e.ph + (cnt[e.ph] = (cnt[e.ph] ?? -1) + 1); const d = def[keys.indexOf(k)]; if (d) Object.assign(e, { x: d.x, y: d.y, w: d.w, h: d.h, size: d.size, align: d.align }); }); s.background = null; }
    SL.commit(); E.renderAll();
  };
  X.cmd['master-reset-all'] = async () => {
    if (E.view !== 'master') return;
    if (!(await BTS.confirm('Reset slide master', 'Reset the master background, text styles, graphics and every layout to the theme defaults?', 'Reset'))) return;
    const real = Object.assign({}, saved.pres, { master: null });
    const mp = buildMasterPres(real); mp.theme = SL.pres.theme; SL.pres = mp; SL.cur = Math.min(SL.cur, mp.slides.length - 1); E.sel = []; lastMasterGeom = null; liveUpdate();
    SL.commit(); E.renderAll();
  };
  const masterEl = (ph) => { if (E.view !== 'master') { BTS.toast('Open View ▸ Slide Master first'); return null; } return SL.pres.slides[0].elements.find(e => e.ph === ph); };
  const afterMaster = () => { liveUpdate(); SL.commit(); E.renderAll(); };
  X.cmd['master-title-font'] = async (el) => { const e = masterEl('title'); if (!e) return; const n = await BTS.fontPicker(el, { current: SL.textDefaults(e, SL.pres).font, theme: X.themeFontChoices() }); if (!n) return; await SL.ensureFonts([n]); e.font = n; afterMaster(); };
  X.cmd['master-body-font'] = async (el) => { const e = masterEl('body'); if (!e) return; const n = await BTS.fontPicker(el, { current: SL.textDefaults(e, SL.pres).font, theme: X.themeFontChoices() }); if (!n) return; await SL.ensureFonts([n]); e.font = n; afterMaster(); };
  X.cmd['master-title-color'] = async (el) => { const e = masterEl('title'); if (!e) return; const c = await BTS.colorPicker(el, { auto: 'Theme color' }); if (!c) return; if (c === 'auto') delete e.color; else e.color = c; afterMaster(); };
  X.cmd['master-body-color'] = async (el) => { const e = masterEl('body'); if (!e) return; const c = await BTS.colorPicker(el, { auto: 'Theme color' }); if (!c) return; if (c === 'auto') delete e.color; else e.color = c; afterMaster(); };
  X.cmd['master-title-size'] = (el) => { const e = masterEl('title'); if (!e) return; const v = +el.value; if (v > 0) { e.size = v; afterMaster(); } };
  X.cmd['master-body-size'] = (el) => { const e = masterEl('body'); if (!e) return; const v = +el.value; if (v > 0) { e.size = v; afterMaster(); } };

  /* keep master styles live on every change made inside the master view */
  const origCommit = SL.commit;
  SL.commit = (opts) => { if (SL.pres && SL.pres._masterView) liveUpdate(); return origCommit(opts); };

  X.ribbon.push(() => {
    const on = E.view === 'master';
    const b = $('#b-v-master'); if (b) b.classList.toggle('on', on);
    if (on) {
      const ms = SL.pres.slides[0]; const te = ms.elements.find(e => e.ph === 'title'), be = ms.elements.find(e => e.ph === 'body');
      const ts = $('#master-title-size'), bs = $('#master-body-size');
      if (ts && document.activeElement !== ts && te) ts.value = te.size || 40; if (bs && document.activeElement !== bs && be) bs.value = be.size || 24;
      const s = SL.slide(); BTS.setStatus('slide', `${icon('layout-dashboard', 14)} Slide Master view · ${BTS.esc(s.mname || '')}${s.isMaster ? '' : ' layout'}`);
    }
  });
})();
