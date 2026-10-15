/* BTS Slides 1.3 — Record narration + Export ▸ Video.
 * Narration: Slide Show ▸ Record Narration plays the show in rehearsal mode and records the microphone separately for
 * every slide (getUserMedia + MediaRecorder), together with the time spent on it. Kept in memory with the open
 * presentation (SL.narration = { slideId: { blob, ms } }) and used by the video export; the timings are written to the
 * slides (transition.auto) like Rehearse Timings.
 * Video: every slide is rendered to a picture (same renderer as PDF / PNG export), played on a canvas with the slide
 * transitions (fade, push, wipe, cover, split, zoom → drawn on the canvas), captured with canvas.captureStream + the
 * narration audio and recorded by MediaRecorder as MP4 (H.264) when Chromium supports it, otherwise WebM.
 * Recording runs in real time (a 10-slide deck at 5 s per slide takes about 50 s).
 */
(function () {
  const $ = (s, r = document) => r.querySelector(s);
  const esc = (s) => BTS.esc(s);
  const X = window.SLX = window.SLX || { cmd: {}, ribbon: [] };
  const P = () => SL.pres;
  SL.narration = SL.narration || {};
  const narrationFor = () => SL.narration; // keyed by slide id (ids are unique), so undo/redo keep it
  const fmt = (ms) => { const t = Math.max(0, Math.round(ms / 1000)); return Math.floor(t / 60) + ':' + String(t % 60).padStart(2, '0'); };

  /* ====================================================================== */
  /*  Record narration                                                       */
  /* ====================================================================== */
  let nar = null; // { stream, rec, cur:{ id, t0, paused, pauseAt }, chunks, done:{id:{blob,ms}}, pending:[promises] }
  function audioMime() { for (const m of ['audio/webm;codecs=opus', 'audio/webm', 'audio/ogg;codecs=opus', 'audio/mp4']) if (window.MediaRecorder && MediaRecorder.isTypeSupported(m)) return m; return ''; }
  function stopCurrent() {
    if (!nar || !nar.rec) return;
    const rec = nar.rec, cur = nar.cur; nar.rec = null; nar.cur = null;
    const now = performance.now(); const ms = now - cur.t0 - cur.paused - (cur.pauseAt ? now - cur.pauseAt : 0);
    nar.pending.push(new Promise(res => {
      const chunks = [];
      rec.ondataavailable = (e) => { if (e.data && e.data.size) chunks.push(e.data); };
      rec.onstop = () => { if (ms > 250) nar.done[cur.id] = { blob: new Blob(chunks, { type: rec.mimeType || 'audio/webm' }), ms: Math.round(ms) }; res(); };
      try { rec.stop(); } catch { res(); }
    }));
  }
  FX.onShowSlide = (i, show) => {
    if (!nar || !show.narrate) return;
    stopCurrent();
    if (i < 0) return; // end screen
    const id = P().slides[i] && P().slides[i].id; if (!id) return;
    const rec = new MediaRecorder(nar.stream, nar.mime ? { mimeType: nar.mime } : {});
    const chunks = []; rec.ondataavailable = (e) => { if (e.data && e.data.size) chunks.push(e.data); };
    rec.start(); nar.rec = rec; nar.cur = { id, t0: performance.now(), paused: 0, pauseAt: show.paused ? performance.now() : 0 };
    if (show.paused) try { rec.pause(); } catch { }
  };
  FX.onShowTick = (show) => {
    if (!nar || !show.narrate || !nar.rec) return;
    const c = nar.cur;
    if (show.paused && !c.pauseAt) { c.pauseAt = performance.now(); try { nar.rec.pause(); } catch { } }
    else if (!show.paused && c.pauseAt) { c.paused += performance.now() - c.pauseAt; c.pauseAt = 0; try { nar.rec.resume(); } catch { } }
  };
  X.startNarration = async () => {
    if (!navigator.mediaDevices || !window.MediaRecorder) { BTS.toast('Recording is not available here', { error: true }); return false; }
    let stream;
    try { stream = await navigator.mediaDevices.getUserMedia({ audio: { echoCancellation: true, noiseSuppression: true } }); }
    catch (e) { BTS.toast('Couldn’t use the microphone: ' + (e.message || e.name), { error: true, ms: 5000 }); return false; }
    nar = { stream, mime: audioMime(), rec: null, cur: null, done: {}, pending: [] };
    const order = X.showOrder ? X.showOrder() : P().slides.map((s, i) => i).filter(i => !P().slides[i].hidden);
    if (!order.length) { stream.getTracks().forEach(t => t.stop()); nar = null; BTS.toast('There are no slides to show'); return false; }
    await FX.start(order[0], { order, rehearse: true, narrate: true, presenter: false });
    const root = $('#show'); if (root) root.insertAdjacentHTML('beforeend', '<div class="nar-bar"><i></i>Recording narration — press → for the next slide, Esc to stop</div>');
    return true;
  };
  FX.onNarrateEnd = async () => {
    if (!nar) return;
    stopCurrent(); await Promise.all(nar.pending);
    nar.stream.getTracks().forEach(t => t.stop());
    const done = nar.done; nar = null;
    const ids = Object.keys(done); if (!ids.length) { BTS.toast('No narration was recorded'); return; }
    ids.forEach(id => { SL.narration[id] = done[id]; const s = P().slides.find(x => x.id === id); if (s) s.transition = Object.assign({ type: 'none', dur: 700 }, s.transition, { auto: Math.round(done[id].ms / 100) / 10 }); });
    P().showSetup = Object.assign({}, P().showSetup || {}, { useTimings: true });
    SL.commit(); E.renderAll();
    BTS.toast(`Narration and timings recorded for ${ids.length} slide${ids.length === 1 ? '' : 's'}. Export ▸ Video uses them.`, { ms: 4500 });
  };
  X.cmd['record-narration'] = async () => {
    const have = Object.keys(narrationFor(P())).filter(id => P().slides.some(s => s.id === id)).length;
    const r = await BTS.dialog({
      title: 'Record narration', width: 470,
      body: `<p style="margin:0 0 10px;line-height:1.5">The slide show starts and your microphone records what you say on each slide. Move to the next slide with → or Space; press Esc to stop.</p>
        <p class="pf-note" style="margin:0">The time you spend on each slide is saved as its timing. The narration is kept while this presentation is open and is added when you use <b>Export ▸ Video</b>.</p>
        ${have ? `<p class="pf-note">${have} slide${have === 1 ? ' has' : 's have'} narration already. Recording again replaces it for the slides you visit.</p>` : ''}`,
      buttons: [...(have ? [{ label: 'Delete narration', value: 'clear' }] : []), { label: 'Cancel', value: null }, { label: 'Start recording', primary: true, value: 'go' }],
    });
    if (r === 'clear') { SL.narration = {}; BTS.toast('Narration deleted'); return; }
    if (r === 'go') X.startNarration();
  };

  /* ====================================================================== */
  /*  Video export                                                           */
  /* ====================================================================== */
  X.videoMime = (height, audio) => {
    const ok = (m) => window.MediaRecorder && MediaRecorder.isTypeSupported(m);
    const avc = height > 720 ? ['avc1.640028', 'avc1.4D0028', 'avc1.42E028', 'avc1.42E01E'] : ['avc1.42E01F', 'avc1.4D401F', 'avc1.42E01E'];
    const list = [];
    avc.forEach(v => { if (audio) list.push(`video/mp4;codecs=${v},mp4a.40.2`, `video/mp4;codecs=${v},opus`); else list.push(`video/mp4;codecs=${v}`); });
    list.push(audio ? 'video/webm;codecs=vp9,opus' : 'video/webm;codecs=vp9', audio ? 'video/webm;codecs=vp8,opus' : 'video/webm;codecs=vp8', 'video/webm');
    return list.find(ok) || '';
  };
  const TR = { fade: 1, push: 1, wipe: 1, cover: 1, reveal: 1, split: 1, zoom: 1 };
  function drawTransition(g, a, b, p, tr, W, H) {
    const e = p < .5 ? 2 * p * p : 1 - Math.pow(-2 * p + 2, 2) / 2; // ease in-out
    const type = TR[tr.type] ? tr.type : 'fade', dir = tr.dir || 'left';
    const v = { left: [-1, 0], right: [1, 0], up: [0, -1], down: [0, 1] }[dir] || [-1, 0];
    g.globalAlpha = 1;
    switch (type) {
      case 'push': g.drawImage(a, v[0] * W * e, v[1] * H * e, W, H); g.drawImage(b, -v[0] * W * (1 - e), -v[1] * H * (1 - e), W, H); break;
      case 'cover': g.drawImage(a, 0, 0, W, H); g.drawImage(b, -v[0] * W * (1 - e), -v[1] * H * (1 - e), W, H); break;
      case 'reveal': g.drawImage(b, 0, 0, W, H); g.drawImage(a, v[0] * W * e, v[1] * H * e, W, H); break;
      case 'wipe': { g.drawImage(a, 0, 0, W, H); g.save(); g.beginPath(); const r = { left: [W * (1 - e), 0, W * e, H], right: [0, 0, W * e, H], up: [0, H * (1 - e), W, H * e], down: [0, 0, W, H * e] }[dir] || [W * (1 - e), 0, W * e, H]; g.rect(...r); g.clip(); g.drawImage(b, 0, 0, W, H); g.restore(); break; }
      case 'split': { g.drawImage(a, 0, 0, W, H); g.save(); g.beginPath(); if (dir === 'up' || dir === 'down') g.rect(0, H / 2 * (1 - e), W, H * e); else g.rect(W / 2 * (1 - e), 0, W * e, H); g.clip(); g.drawImage(b, 0, 0, W, H); g.restore(); break; }
      case 'zoom': { g.drawImage(a, 0, 0, W, H); g.globalAlpha = e; const k = .25 + .75 * e; g.drawImage(b, W * (1 - k) / 2, H * (1 - k) / 2, W * k, H * k); g.globalAlpha = 1; break; }
      default: g.drawImage(a, 0, 0, W, H); g.globalAlpha = e; g.drawImage(b, 0, 0, W, H); g.globalAlpha = 1;
    }
  }
  /* o = { height: 720|1080, secs: 5, useTimings, narration, transitions: 'slides'|'fade'|'none', hidden, onProgress(frac, text), canvas, signal:{cancel} }
   * → { data: Uint8Array, mime, ext, width, height, duration } | null (cancelled) */
  X.exportVideo = async (o = {}) => {
    const pres = P(); const H0 = o.height || 720; const W0 = Math.round(H0 * 16 / 9 / 2) * 2;
    const idx = pres.slides.map((s, i) => i).filter(i => o.hidden || !pres.slides[i].hidden);
    if (!idx.length) throw new Error('There are no slides to export');
    const prog = o.onProgress || (() => { }); const sig = o.signal || {};
    prog(0, 'Drawing the slides…');
    const pngs = await X.renderPNGs(idx, W0); if (!pngs) return null;
    if (sig.cancel) return null;
    // fit each slide (any aspect ratio) into the video frame
    const frames = [];
    for (const d of pngs) {
      const bmp = await createImageBitmap(new Blob([d], { type: 'image/png' }));
      const c = document.createElement('canvas'); c.width = W0; c.height = H0; const g = c.getContext('2d'); g.fillStyle = '#000'; g.fillRect(0, 0, W0, H0);
      const k = Math.min(W0 / bmp.width, H0 / bmp.height); const w = bmp.width * k, h = bmp.height * k; g.drawImage(bmp, (W0 - w) / 2, (H0 - h) / 2, w, h); bmp.close && bmp.close();
      frames.push(c);
    }
    // timeline
    const narr = o.narration ? narrationFor(pres) : {};
    let actx = null, dest = null; const buffers = {};
    const narrIds = idx.map(i => pres.slides[i].id).filter(id => narr[id]);
    if (narrIds.length) {
      actx = new AudioContext(); dest = actx.createMediaStreamDestination();
      for (const id of narrIds) { try { buffers[id] = await actx.decodeAudioData(await narr[id].blob.arrayBuffer()); } catch (e) { console.warn('narration decode', e); } }
    }
    let t = 0; const seg = idx.map((i, k) => {
      const s = pres.slides[i]; const tr = s.transition || {};
      let dur = o.useTimings && tr.auto != null && tr.auto !== false ? +tr.auto : (o.secs || 5);
      const buf = buffers[s.id]; if (buf) dur = Math.max(dur, buf.duration + .4);
      dur = Math.max(.5, dur);
      let trans = null;
      if (k > 0 && o.transitions !== 'none') {
        if (o.transitions === 'fade') trans = { type: 'fade', dur: Math.min(.8, dur / 2) };
        else if (tr.type && tr.type !== 'none') trans = { type: tr.type, dir: tr.dir, dur: Math.min((tr.dur || 700) / 1000, dur / 2) };
      }
      const r = { start: t, dur, trans, id: s.id }; t += dur; return r;
    });
    const total = t;
    // canvas + recorder
    const cv = o.canvas || document.createElement('canvas'); cv.width = W0; cv.height = H0; const g = cv.getContext('2d');
    g.drawImage(frames[0], 0, 0);
    const mime = X.videoMime(H0, !!dest);
    const vstream = cv.captureStream(30);
    const tracks = [...vstream.getVideoTracks(), ...(dest ? dest.stream.getAudioTracks() : [])];
    const stream = new MediaStream(tracks);
    const rec = new MediaRecorder(stream, Object.assign({ videoBitsPerSecond: H0 > 720 ? 8e6 : 4e6, audioBitsPerSecond: 128000 }, mime ? { mimeType: mime } : {}));
    const chunks = []; rec.ondataavailable = (e) => { if (e.data && e.data.size) chunks.push(e.data); };
    const stopped = new Promise(res => { rec.onstop = res; });
    rec.start(1000);
    const t0 = performance.now();
    if (actx) { await actx.resume(); const a0 = actx.currentTime + .05; seg.forEach(sg => { const b = buffers[sg.id]; if (!b) return; const src = actx.createBufferSource(); src.buffer = b; src.connect(dest); src.start(a0 + sg.start + (sg.trans ? sg.trans.dur * .5 : 0)); }); }
    await new Promise(res => {
      let k = 0;
      const step = () => {
        const now = (performance.now() - t0) / 1000;
        if (sig.cancel || now >= total) { res(); return; }
        while (k < seg.length - 1 && now >= seg[k + 1].start) k++;
        const sg = seg[k];
        if (sg.trans && now - sg.start < sg.trans.dur) drawTransition(g, frames[k - 1], frames[k], (now - sg.start) / sg.trans.dur, sg.trans, W0, H0);
        else { g.globalAlpha = 1; g.drawImage(frames[k], 0, 0); }
        prog(Math.min(1, now / total), `Recording slide ${k + 1} of ${seg.length} · ${fmt(now * 1000)} / ${fmt(total * 1000)}`);
        setTimeout(step, 1000 / 30);
      };
      step();
    });
    g.drawImage(frames[frames.length - 1], 0, 0);
    await new Promise(r => setTimeout(r, 120));
    rec.stop(); await stopped;
    tracks.forEach(tk => tk.stop()); if (actx) actx.close();
    if (sig.cancel) return null;
    const type = rec.mimeType || mime || 'video/webm';
    const data = new Uint8Array(await new Blob(chunks, { type }).arrayBuffer());
    return { data, mime: type, ext: /mp4/.test(type) ? 'mp4' : 'webm', width: W0, height: H0, duration: total, slides: seg.length };
  };

  X.cmd['export-video'] = async () => {
    if (!BTS.desktop) return BTS.toast('Video export needs the BTS Office desktop app', { error: true });
    if (!window.MediaRecorder || !HTMLCanvasElement.prototype.captureStream) return BTS.toast('Video recording is not available here', { error: true });
    const pres = P();
    const timed = pres.slides.some(s => !s.hidden && s.transition && s.transition.auto != null && s.transition.auto !== false);
    const nN = Object.keys(narrationFor(pres)).filter(id => pres.slides.some(s => s.id === id)).length;
    let o0 = {}; try { o0 = JSON.parse(localStorage.getItem('bts-slides-video') || '{}'); } catch { }
    const mp4 = !!X.videoMime(1080, false).startsWith('video/mp4');
    const o = await BTS.dialog({
      title: 'Export video', width: 470,
      body: `<div class="vid-dlg">
        <div class="field-row"><div class="field"><label>Quality</label><select id="vd-h"><option value="720">HD 720p (1280 × 720)</option><option value="1080" ${o0.height === 1080 ? 'selected' : ''}>Full HD 1080p (1920 × 1080)</option></select></div>
        <div class="field"><label>Transitions</label><select id="vd-tr"><option value="slides">As set on the slides</option><option value="fade" ${o0.transitions === 'fade' ? 'selected' : ''}>Crossfade between all slides</option><option value="none" ${o0.transitions === 'none' ? 'selected' : ''}>None</option></select></div></div>
        <label class="check"><input type="radio" name="vd-t" value="secs" ${!timed ? 'checked' : ''}> Seconds on each slide <input type="number" id="vd-secs" min="1" max="120" step="0.5" value="${o0.secs || 5}" style="width:64px;margin-left:6px"></label>
        <label class="check" title="${timed ? 'Use the timings saved with Rehearse Timings or Record Narration' : 'No slide has a recorded timing yet (Slide Show ▸ Rehearse Timings)'}"><input type="radio" name="vd-t" value="timings" ${timed ? 'checked' : 'disabled'}> Use recorded timings${timed ? '' : ' (none yet)'}</label>
        <label class="check"><input type="checkbox" id="vd-nar" ${nN ? 'checked' : 'disabled'}> Include narration${nN ? ` (${nN} slide${nN === 1 ? '' : 's'})` : ' — record it with Slide Show ▸ Record Narration'}</label>
        <p class="pf-note">The video is recorded in real time: about ${fmt(pres.slides.filter(s => !s.hidden).length * (o0.secs || 5) * 1000)} for this presentation. Animations are shown in their final state.${mp4 ? '' : ' This computer can’t make MP4 here, so the video will be a WebM file (plays in VLC, Chrome and most players).'}</p></div>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'Export…', primary: true, value: (r) => ({ height: +$('#vd-h', r).value, transitions: $('#vd-tr', r).value, secs: Math.max(1, +$('#vd-secs', r).value || 5), useTimings: $('input[name=vd-t]:checked', r).value === 'timings', narration: $('#vd-nar', r).checked }) }],
    });
    if (!o) return;
    try { localStorage.setItem('bts-slides-video', JSON.stringify({ height: o.height, transitions: o.transitions, secs: o.secs })); } catch { }
    if (E.editing) E.endEdit();
    const sig = { cancel: false }; let res = null, err = null, finished = false;
    const body = document.createElement('div'); body.className = 'vid-prog';
    body.innerHTML = `<canvas></canvas><div class="bar"><i></i></div><div class="st">Getting ready…</div>`;
    const cv = $('canvas', body);
    let root = null;
    const dlg = BTS.dialog({ title: 'Exporting video', width: 520, body, buttons: [{ label: 'Cancel', value: 'cancel' }], onOpen: (r) => { root = r; } });
    dlg.then(() => { if (!finished) sig.cancel = true; });
    const job = X.exportVideo(Object.assign({}, o, { canvas: cv, signal: sig, onProgress: (f, txt) => { $('.bar > i', body).style.width = (f * 100).toFixed(1) + '%'; $('.st', body).textContent = txt; } })).then(r => { res = r; }, e => { err = e; });
    await job; finished = true;
    if (root && root.isConnected) root.querySelector('[data-x]').click();
    if (err) { console.error(err); return BTS.toast('Video export failed: ' + (err.message || err), { error: true, ms: 6000 }); }
    if (!res) return BTS.toast('Video export cancelled');
    const ok = await X.saveBytes(`${BTS.title}.${res.ext}`, res.data, res.ext === 'mp4' ? 'MP4 video' : 'WebM video', res.ext);
    if (ok && res.ext !== 'mp4') BTS.toast('Saved as WebM: MP4 recording is not available on this computer', { ms: 5000 });
  };
})();
