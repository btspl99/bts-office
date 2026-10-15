/* BTS Slides 1.2 — slide show tools: Set Up Slide Show (speaker / window / kiosk, loop, range, custom show, timings,
 * presenter view), custom slide shows, rehearse timings (+ summary), and Insert ▸ Reuse Slides.
 * Saved with the presentation: pres.showSetup = { type, loop, noAnims, presenter, useTimings, range:{from,to}|null, custom:id|null, kiosk }
 * and pres.customShows = [{ id, name, slides:[slide ids] }].
 */
(function () {
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const esc = (s) => BTS.esc(s);
  const X = window.SLX = window.SLX || { cmd: {}, ribbon: [] };
  const C = () => window.CMD;
  const P = () => SL.pres;
  const setup = () => P().showSetup || {};
  const slideTitle = (s, i) => { const t = s.elements.find(e => e.ph === 'title' && !SL.isEmptyText(e)); return t ? SL.plainText(t.html).trim().replace(/\s+/g, ' ').slice(0, 60) : 'Slide ' + (i + 1); };
  X.slideTitle = slideTitle;

  /* ---------------- order of slides for "From Beginning" ---------------- */
  X.showOrder = (st = setup()) => {
    const all = P().slides;
    if (st.custom) { const cs = (P().customShows || []).find(c => c.id === st.custom); if (cs) return cs.slides.map(id => all.findIndex(s => s.id === id)).filter(i => i >= 0); }
    let idx = all.map((s, i) => i).filter(i => !all[i].hidden);
    if (st.range) { const a = Math.max(1, st.range.from || 1) - 1, b = Math.min(all.length, st.range.to || all.length) - 1; idx = idx.filter(i => i >= a && i <= b); }
    return idx;
  };
  const start = (opts = {}) => { const order = opts.order || X.showOrder(); if (!order.length) return BTS.toast('There are no slides to show'); return FX.start(order[0], Object.assign({ order, windowed: setup().type === 'window' }, opts)); };
  X.cmd['present-start'] = () => start();
  X.cmd['present-current'] = () => FX.start(SL.cur, { allowHidden: true, windowed: setup().type === 'window' });
  X.cmd['present-presenter'] = () => FX.start(SL.cur, { allowHidden: true, presenter: true });
  X.cmd['present-window'] = () => FX.start(SL.cur, { windowed: true, allowHidden: true, presenter: false });

  /* ---------------- ribbon toggles ---------------- */
  const setSetup = (patch) => { P().showSetup = Object.assign({}, setup(), patch); SL.commit(); };
  X.cmd['show-loop'] = (el) => setSetup({ loop: el.checked });
  X.cmd['show-timings'] = (el) => setSetup({ useTimings: el.checked });
  X.cmd['show-anims'] = (el) => setSetup({ noAnims: !el.checked });
  X.cmd['show-presenter'] = (el) => setSetup({ presenter: el.checked });
  X.cmd['clear-timings'] = async () => {
    const n = P().slides.filter(s => s.transition && s.transition.auto != null).length;
    if (!n) return BTS.toast('No slide has a recorded timing');
    if (!(await BTS.confirm('Clear timings', `Remove the timings from ${n} slide(s)? Slides will advance on click.`, 'Clear'))) return;
    P().slides.forEach(s => { if (s.transition) delete s.transition.auto; }); SL.commit(); E.renderAll(); BTS.toast('Timings cleared');
  };
  X.ribbon.push(() => {
    const st = setup();
    const set = (id, v) => { const e = $('#' + id); if (e) e.checked = !!v; };
    set('show-loop', st.loop || st.type === 'kiosk'); set('show-timings', st.useTimings !== false); set('show-anims', !st.noAnims); set('show-presenter', st.presenter);
  });

  /* ---------------- Set Up Slide Show ---------------- */
  X.cmd['setup-show'] = async () => {
    const st = Object.assign({ type: 'speaker', useTimings: true }, setup()); const n = P().slides.length; const shows = P().customShows || [];
    const r = await BTS.dialog({
      title: 'Set Up Show', width: 620,
      body: `<div class="ss-cols"><div>
          <div class="pf-sec">Show type</div>
          <label class="check"><input type="radio" name="sst" value="speaker" ${st.type === 'speaker' ? 'checked' : ''}> Presented by a speaker (full screen)</label>
          <label class="check"><input type="radio" name="sst" value="window" ${st.type === 'window' ? 'checked' : ''}> Browsed by an individual (window)</label>
          <label class="check"><input type="radio" name="sst" value="kiosk" ${st.type === 'kiosk' ? 'checked' : ''}> Browsed at a kiosk (full screen)</label>
          <div class="pf-sec">Show options</div>
          <label class="check"><input type="checkbox" id="ss-loop" ${st.loop || st.type === 'kiosk' ? 'checked' : ''} ${st.type === 'kiosk' ? 'disabled' : ''}> Loop continuously until ‘Esc’</label>
          <label class="check"><input type="checkbox" id="ss-noanim" ${st.noAnims ? 'checked' : ''}> Show without animation</label>
          <label class="check"><input type="checkbox" id="ss-pv" ${st.presenter ? 'checked' : ''}> Use presenter view</label>
        </div><div>
          <div class="pf-sec">Show slides</div>
          <label class="check"><input type="radio" name="ssr" value="all" ${!st.range && !st.custom ? 'checked' : ''}> All</label>
          <label class="check"><input type="radio" name="ssr" value="range" ${st.range ? 'checked' : ''}> From <input type="number" id="ss-from" min="1" max="${n}" value="${st.range ? st.range.from : 1}" style="width:58px"> to <input type="number" id="ss-to" min="1" max="${n}" value="${st.range ? st.range.to : n}" style="width:58px"></label>
          <label class="check"><input type="radio" name="ssr" value="custom" ${st.custom ? 'checked' : ''} ${shows.length ? '' : 'disabled'}> Custom show: <select id="ss-custom" ${shows.length ? '' : 'disabled'}>${shows.map(c => `<option value="${c.id}" ${st.custom === c.id ? 'selected' : ''}>${esc(c.name)}</option>`).join('') || '<option>(none yet)</option>'}</select></label>
          <div class="pf-sec">Advance slides</div>
          <label class="check"><input type="radio" name="ssa" value="manual" ${st.useTimings === false ? 'checked' : ''}> Manually</label>
          <label class="check"><input type="radio" name="ssa" value="timings" ${st.useTimings !== false ? 'checked' : ''}> Using timings, if present</label>
          <p class="pf-note">Presenter view opens beside the slide in this window. Kiosk mode advances using timings (5 s when a slide has none) and ignores clicks; press Esc to stop.</p>
        </div></div>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (root) => {
        const type = $('input[name=sst]:checked', root).value; const rng = $('input[name=ssr]:checked', root).value;
        const from = Math.max(1, Math.min(n, +$('#ss-from', root).value || 1)), to = Math.max(from, Math.min(n, +$('#ss-to', root).value || n));
        return { type, kiosk: type === 'kiosk', loop: type === 'kiosk' || $('#ss-loop', root).checked, noAnims: $('#ss-noanim', root).checked, presenter: $('#ss-pv', root).checked,
          range: rng === 'range' ? { from, to } : null, custom: rng === 'custom' && shows.length ? $('#ss-custom', root).value : null, useTimings: $('input[name=ssa]:checked', root).value === 'timings' };
      } }],
      onOpen: (root) => root.addEventListener('change', e => { if (e.target.name === 'sst') { const k = e.target.value === 'kiosk'; const l = $('#ss-loop', root); l.disabled = k; if (k) l.checked = true; } if (e.target.id === 'ss-from' || e.target.id === 'ss-to') $('input[name=ssr][value=range]', root).checked = true; if (e.target.id === 'ss-custom') $('input[name=ssr][value=custom]', root).checked = true; }),
    });
    if (!r) return;
    P().showSetup = r; SL.commit(); E.updateRibbon(); BTS.toast('Slide show settings saved with this presentation');
  };

  /* ---------------- Custom slide shows ---------------- */
  X.cmd['custom-shows-menu'] = (el) => {
    const shows = P().customShows || [];
    BTS.menu(el, [...(shows.length ? [{ head: 'Play a custom show' }, ...shows.map(c => ({ label: `${c.name} (${c.slides.length})`, icon: 'play', action: () => X.playCustom(c.id) }))] : [{ label: 'No custom shows yet', disabled: true }]), { sep: 1 }, { label: 'Custom Shows…', icon: 'list-ordered', action: () => X.cmd['custom-shows']() }]);
  };
  X.playCustom = (id) => { const cs = (P().customShows || []).find(c => c.id === id); if (!cs) return; const order = cs.slides.map(sid => P().slides.findIndex(s => s.id === sid)).filter(i => i >= 0); if (!order.length) return BTS.toast('This custom show has no slides'); FX.start(order[0], { order, windowed: setup().type === 'window' }); };
  X.cmd['custom-shows'] = async () => {
    let sel = (P().customShows || [])[0]?.id || null;
    const res = await BTS.dialog({
      title: 'Custom Shows', width: 520,
      body: `<div class="cs-man"><div class="cs-shows" id="cs-shows"></div><div class="cs-btns"><button class="btn" data-a="new" title="Create a new custom show">New…</button><button class="btn" data-a="edit" title="Edit the selected show">Edit…</button><button class="btn" data-a="copy" title="Copy the selected show">Copy</button><button class="btn" data-a="del" title="Remove the selected show">Remove</button></div></div><p class="pf-note">A custom show is a named set of slides, in any order — for example a short version of your talk. Play it from Slide Show ▸ Custom Show, or choose it in Set Up Slide Show.</p>`,
      buttons: [{ label: 'Close', value: null }, { label: 'Show', primary: true, value: () => sel ? 'play' : null }],
      onOpen: (root) => {
        const draw = () => { const shows = P().customShows || []; $('#cs-shows', root).innerHTML = shows.length ? shows.map(c => `<div class="cs-item ${c.id === sel ? 'sel' : ''}" data-id="${c.id}">${icon('list-ordered', 14)}<span>${esc(c.name)}</span><small>${c.slides.length} slide${c.slides.length === 1 ? '' : 's'}</small></div>`).join('') : '<p class="pf-note" style="padding:10px">No custom shows yet. Click New… to create one.</p>'; };
        draw();
        root.addEventListener('click', async e => {
          const it = e.target.closest('.cs-item'); if (it) { sel = it.dataset.id; draw(); return; }
          const a = e.target.closest('[data-a]'); if (!a) return; const shows = P().customShows = P().customShows || []; const cur = shows.find(c => c.id === sel);
          if (a.dataset.a === 'new') { const c = await editShow({ id: SL.uid('cs'), name: 'Custom Show ' + (shows.length + 1), slides: [] }); if (c) { shows.push(c); sel = c.id; SL.commit(); } draw(); }
          if (a.dataset.a === 'edit' && cur) { const c = await editShow(SL.clone(cur)); if (c) { Object.assign(cur, c); SL.commit(); } draw(); }
          if (a.dataset.a === 'copy' && cur) { const c = Object.assign(SL.clone(cur), { id: SL.uid('cs'), name: 'Copy of ' + cur.name }); shows.push(c); sel = c.id; SL.commit(); draw(); }
          if (a.dataset.a === 'del' && cur) { P().customShows = shows.filter(c => c !== cur); if (setup().custom === cur.id) P().showSetup = Object.assign({}, setup(), { custom: null }); sel = P().customShows[0]?.id || null; SL.commit(); draw(); }
        });
        root.addEventListener('dblclick', e => { const it = e.target.closest('.cs-item'); if (it) { sel = it.dataset.id; $('footer .primary', root).click(); } });
      },
    });
    if (res === 'play' && sel) X.playCustom(sel);
  };
  async function editShow(cs) {
    const slides = P().slides;
    return BTS.dialog({
      title: 'Define Custom Show', width: 760,
      body: `<div class="field"><label>Slide show name</label><input id="cse-name" value="${esc(cs.name)}"></div>
        <div class="cse"><div><div class="pf-sec">Slides in presentation</div><div class="cse-list" id="cse-src">${slides.map((s, i) => `<label class="cse-it"><input type="checkbox" data-i="${i}"><span class="cse-th">${SL.thumbHTML(s, P(), 72, i)}</span><span>${i + 1}. ${esc(slideTitle(s, i))}${s.hidden ? ' <small>(hidden)</small>' : ''}</span></label>`).join('')}</div></div>
        <div class="cse-mid"><button class="btn" id="cse-add" title="Add the checked slides to the custom show">Add ${icon('chevron-right', 14)}</button></div>
        <div><div class="pf-sec">Slides in custom show</div><div class="cse-list" id="cse-dst"></div></div></div>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (root) => Object.assign(cs, { name: $('#cse-name', root).value.trim() || cs.name }) }],
      onOpen: (root) => {
        const draw = () => { $('#cse-dst', root).innerHTML = cs.slides.map((id, k) => { const i = slides.findIndex(s => s.id === id); if (i < 0) return ''; return `<div class="cse-it" data-k="${k}"><span class="cse-th">${SL.thumbHTML(slides[i], P(), 72, i)}</span><span>${k + 1}. ${esc(slideTitle(slides[i], i))}</span><span class="cse-acts"><button class="rb" data-m="up" title="Move up">${icon('arrow-up', 13)}</button><button class="rb" data-m="down" title="Move down">${icon('arrow-down', 13)}</button><button class="rb" data-m="del" title="Remove from the custom show">${icon('x', 13)}</button></span></div>`; }).join('') || '<p class="pf-note" style="padding:10px">Check slides on the left and click Add.</p>'; };
        draw();
        $('#cse-add', root).addEventListener('click', () => { $$('#cse-src input:checked', root).forEach(c => { cs.slides.push(slides[+c.dataset.i].id); c.checked = false; }); draw(); });
        $('#cse-dst', root).addEventListener('click', e => { const b = e.target.closest('[data-m]'); if (!b) return; const k = +b.closest('[data-k]').dataset.k; const a = cs.slides; if (b.dataset.m === 'del') a.splice(k, 1); if (b.dataset.m === 'up' && k > 0) [a[k - 1], a[k]] = [a[k], a[k - 1]]; if (b.dataset.m === 'down' && k < a.length - 1) [a[k + 1], a[k]] = [a[k], a[k + 1]]; draw(); });
      },
    });
  }

  /* ---------------- Rehearse timings ---------------- */
  X.cmd['rehearse'] = () => start({ rehearse: true, presenter: false });
  FX.onRehearseEnd = async (rec, total, order) => {
    const ids = Object.keys(rec).filter(id => rec[id] > 300); if (!ids.length) return;
    const rows = P().slides.map((s, i) => ({ s, i, ms: rec[s.id] })).filter(r => r.ms > 300);
    const r = await BTS.dialog({
      title: 'Rehearsal timings', width: 520,
      body: `<div class="rh-sum"><div class="rh-total-big">${icon('timer', 22)} <b>${FX._fmt(total)}</b><span>total time for ${rows.length} slide${rows.length === 1 ? '' : 's'}</span></div>
        <table class="rh-table"><thead><tr><th>Slide</th><th>Title</th><th>Time</th><th>Previous</th></tr></thead><tbody>${rows.map(x => `<tr><td>${x.i + 1}</td><td>${esc(slideTitle(x.s, x.i))}</td><td><b>${FX._fmt(x.ms)}</b></td><td>${x.s.transition && x.s.transition.auto != null ? FX._fmt(x.s.transition.auto * 1000) : '—'}</td></tr>`).join('')}</tbody></table>
        <p class="pf-note">Keep the new timings to make the slides advance automatically when you play the show with timings.</p></div>`,
      buttons: [{ label: 'Discard', value: null }, { label: 'Keep and play', value: 'play' }, { label: 'Keep timings', primary: true, value: 'keep' }],
    });
    if (!r) return;
    rows.forEach(x => { x.s.transition = Object.assign({ type: 'none', dur: 700 }, x.s.transition, { auto: Math.round(x.ms / 100) / 10 }); });
    P().showSetup = Object.assign({}, setup(), { useTimings: true });
    SL.commit(); E.renderAll(); BTS.toast(`Timings saved for ${rows.length} slide(s)`);
    void order;
    if (r === 'play') start({ useTimings: true });
  };

  /* ---------------- Insert ▸ Reuse Slides ---------------- */
  // Bake theme tokens into explicit values so a slide keeps its look in another theme ("keep source formatting").
  function bakeSlide(s, src) {
    const col = (v) => { if (!v || v === 'none' || typeof v !== 'string') return v; const c = SL.color(v, src); return c && c[0] === '#' ? c : v; };
    const fixFill = (f) => f && typeof f === 'object' && f.type === 'gradient' ? Object.assign({}, f, { color: col(f.color), color2: col(f.color2) }) : col(f);
    s.background = SL.clone(s.background || SL.themeBg(src, s.layout));
    if (s.background.color) s.background.color = col(s.background.color); if (s.background.color2) s.background.color2 = col(s.background.color2);
    const th = SL.themeOf(src);
    if (th.deco && !s.hideGraphics) {
      const svg = SL.decoSVG(s, src); if (svg) { const url = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg.replace('<svg ', '<svg xmlns="http://www.w3.org/2000/svg" ')); s.layer = [SL.newElement('image', { x: 0, y: 0, w: SL.W(src), h: SL.H(src), src: SL.addAsset(url), name: 'Source theme graphics', decorative: true }), ...(s.layer || [])]; s.noDeco = true; }
    }
    const els = [...s.elements, ...(s.layer || [])];
    els.forEach(el => {
      if (el.type === 'text' || el.type === 'shape') { const d = SL.textDefaults(el, src); if (!el.font) el.font = d.font; if (!el.color && el.type === 'text') el.color = d.color; if (el.bold == null && el.role === 'title') el.bold = d.bold; }
      if (el.color) el.color = col(el.color); if (el.fill) el.fill = fixFill(el.fill); if (el.stroke) el.stroke = col(el.stroke);
      if (el.type === 'table') { el.accent = col(el.accent || 'accent1'); el.font = el.font || SL.fontOf('body', src); el.rows.forEach(r => r.forEach(c => { if (c.fill) c.fill = col(c.fill); if (c.color) c.color = col(c.color); })); }
      if (el.type === 'icon') el.color = col(el.color || 'accent1');
      if (el.type === 'chart' && el.chart) el.chart.colors = SL.palette(src).accents.slice();
    });
  }
  function adoptSlide(s, src) { // use the destination theme: drop the source look
    const pal = SL.palette(src); const tf = SL.themeFonts(src); const same = (a, b) => String(a || '').toLowerCase() === String(b || '').toLowerCase();
    if (src.theme.id === 'custom') { s.layer = []; s.background = null; }
    s.elements.forEach(el => { if (el.type === 'text' && (same(el.color, pal.text) || same(el.color, pal.title))) delete el.color; if (el.font && (el.font === tf.title || el.font === tf.body)) delete el.font; });
  }
  X.reuseSlides = async () => {
    if (E.masterGuard && E.masterGuard()) return;
    let src = null, name = '';
    const chosen = new Set();
    const res = await BTS.dialog({
      title: 'Reuse Slides', width: 760,
      body: `<div class="ru-top"><button class="btn primary" id="ru-browse" title="Choose a PowerPoint file">${icon('folder-open', 15)} Browse…</button><span id="ru-name" class="pf-note">Choose a .pptx file to see its slides.</span></div>
        <div class="ru-grid" id="ru-grid"></div>
        <div class="ru-foot"><label class="check" title="Keep the colors, fonts and background of the source presentation"><input type="checkbox" id="ru-keep"> Keep source formatting</label><span class="pf-note" id="ru-count"></span><button class="btn sm" id="ru-all" title="Select every slide">Select all</button></div>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'Insert', primary: true, value: (root) => ({ keep: $('#ru-keep', root).checked }), validate: () => { if (!src) { BTS.toast('Choose a file first'); return false; } if (!chosen.size) { BTS.toast('Click the slides you want to insert'); return false; } return true; } }],
      onOpen: (root) => {
        const grid = $('#ru-grid', root);
        const draw = () => { $('#ru-count', root).textContent = src ? `${chosen.size} of ${src.slides.length} selected` : ''; grid.querySelectorAll('.ru-it').forEach(b => b.classList.toggle('sel', chosen.has(+b.dataset.i))); };
        $('#ru-browse', root).addEventListener('click', async () => {
          const f = await BTS.pickFile('.pptx'); if (!f) return;
          try { src = await window.PPTXImport.parse(f.data); delete src.props; name = f.name; } catch (err) { BTS.toast('Couldn’t read that presentation: ' + err.message, { error: true }); return; }
          await SL.ensureFonts(src);
          $('#ru-name', root).textContent = `${name} · ${src.slides.length} slides`; chosen.clear();
          grid.innerHTML = src.slides.map((s, i) => `<button class="ru-it" data-i="${i}" title="Click to select slide ${i + 1}">${SL.thumbHTML(s, src, 160, i)}<span>${i + 1}. ${esc(slideTitle(s, i))}</span></button>`).join('');
          draw();
        });
        grid.addEventListener('click', e => { const b = e.target.closest('.ru-it'); if (!b) return; const i = +b.dataset.i; if (chosen.has(i)) chosen.delete(i); else chosen.add(i); draw(); });
        $('#ru-all', root).addEventListener('click', () => { if (!src) return; if (chosen.size === src.slides.length) chosen.clear(); else src.slides.forEach((_, i) => chosen.add(i)); draw(); });
        X._reuseHook = { set: async (data, fname, picks) => { src = await window.PPTXImport.parse(data); delete src.props; name = fname; grid.innerHTML = src.slides.map((s, i) => `<button class="ru-it" data-i="${i}">${SL.thumbHTML(s, src, 160, i)}<span>${i + 1}. ${esc(slideTitle(s, i))}</span></button>`).join(''); chosen.clear(); (picks || src.slides.map((_, i) => i)).forEach(i => chosen.add(i)); draw(); }, root };
      },
    });
    X._reuseHook = null;
    if (!res || !src) return;
    X.insertReused(src, [...chosen].sort((a, b) => a - b), res.keep);
  };
  X.insertReused = (src, picks, keep) => {
    const kx = SL.W(P()) / SL.W(src), ky = SL.H(P()) / SL.H(src);
    const add = picks.map(i => {
      const s = SL.clone(src.slides[i]);
      if (keep) bakeSlide(s, src); else adoptSlide(s, src);
      s.id = SL.uid('s'); s.comments = s.comments || [];
      [...s.elements, ...(s.layer || [])].forEach(el => { SL.reId(el); el.x *= kx; el.y *= ky; el.w *= kx; el.h *= ky; });
      E.regroup(s.elements);
      return s;
    });
    const at = SL.cur + 1; P().slides.splice(at, 0, ...add); SL.cur = at; E.railSel = new Set(add.map((_, k) => at + k)); E.sel = [];
    SL.commit(); E.renderAll(); BTS.toast(`Inserted ${add.length} slide(s)${keep ? ' with source formatting' : ''}`);
  };
  X.cmd['reuse-slides'] = () => X.reuseSlides();
})();
