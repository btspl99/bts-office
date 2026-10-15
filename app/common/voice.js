/* BTS Office — offline voice typing (all apps)
 *
 *   BTSVoice.toggle()          open the panel and start listening / stop
 *   BTSVoice.start(lang)       lang: 'en-in' (English, India) | 'hi' (हिंदी); default = last used
 *   BTSVoice.stop()            stop listening (releases the microphone) — panel stays open
 *   BTSVoice.close()           stop and hide the panel
 *   BTSVoice.isOn()            true while listening
 *   BTSVoice.lang              current language
 *   BTSVoice.models()          -> { models:[...], bytes }   (downloaded speech models, desktop only)
 *   BTSVoice.removeModels()    -> bytes freed (all languages)
 *
 * Recognised text goes to the caret of the last focused text box / editable area, through
 * document.execCommand('insertText') so the app's undo and change tracking see it.
 * Spoken commands: "full stop"/"period", "comma", "question mark", "exclamation mark", "new line",
 * "new paragraph" (Hindi: "पूर्ण विराम", "अल्पविराम", "प्रश्नवाचक चिह्न", "नई पंक्ति", "नया पैराग्राफ").
 *
 * Recognition runs fully offline with vosk-browser (WebAssembly, app/vendor/vosk/vosk.js, loaded on first use)
 * and Vosk small models that main/voice.js downloads once per language into userData/voice-models.
 * Page CSP needs: script-src 'wasm-unsafe-eval' blob:; worker-src blob:; connect-src blob:
 *
 * Test hook: BTSVoice._testStream = MediaStream  (used instead of the microphone).
 */
(function () {
  const LANGS = {
    'en-in': { label: 'English (India)', short: 'EN', mb: 36 },
    hi: { label: 'हिंदी', short: 'हि', mb: 42 },
  };
  const SRC = document.currentScript && document.currentScript.src;
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const ic = (n, s = 16) => (window.icon ? window.icon(n, s) : '');
  const store = { get: (k) => { try { return localStorage.getItem(k); } catch { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch { } } };

  const V = window.BTSVoice = {
    lang: LANGS[store.get('bts-voice-lang')] ? store.get('bts-voice-lang') : 'en-in',
    isOn: () => state === 'listening',
    toggle: () => (state === 'listening' || state === 'loading' ? V.stop() : V.start()),
    start, stop, close,
    models: () => (window.bts ? bts.ext('voice-models') : Promise.resolve({ models: [], bytes: 0 })),
    removeModels: async () => { if (!window.bts) return 0; const r = await bts.ext('voice-delete'); return r.freed || 0; },
    _testStream: null,
  };

  /* ---------- where the text goes: last focused editable + its selection ---------- */
  let target = null, saved = null;
  const textInput = (el) => el.tagName === 'TEXTAREA' || (el.tagName === 'INPUT' && /^(text|search|url|email|tel|)$/i.test(el.type || ''));
  function editableRoot(el) {
    if (!el || el.nodeType !== 1 || el.closest('.bts-voice')) return null;
    if (textInput(el)) return el.readOnly || el.disabled ? null : el;
    if (!el.isContentEditable) return null;
    while (el.parentElement && el.parentElement.isContentEditable) el = el.parentElement;
    return el;
  }
  function save() {
    if (!target) return;
    if (textInput(target)) { saved = { s: target.selectionStart ?? target.value.length, e: target.selectionEnd ?? target.value.length }; return; }
    const sel = getSelection();
    if (sel.rangeCount && target.contains(sel.getRangeAt(0).startContainer)) saved = { r: sel.getRangeAt(0).cloneRange() };
  }
  document.addEventListener('focusin', (e) => { const r = editableRoot(e.target); if (r) { target = r; saved = null; save(); } }, true);
  document.addEventListener('focusout', (e) => { if (target && e.target === target) save(); }, true);
  for (const ev of ['selectionchange', 'keyup', 'mouseup', 'input', 'select'])   // text boxes fire selectionchange on the element, not the document
    document.addEventListener(ev, () => { const a = document.activeElement; if (a && target && (a === target || target.contains(a))) save(); }, true);
  function restore() {
    if (!target || !target.isConnected) return false;
    const a = document.activeElement;
    if (a === target || target.contains(a)) { save(); return true; }   // still focused: the live caret is right
    target.focus({ preventScroll: true });
    if (!saved) return true;
    if (saved.r) { const sel = getSelection(); sel.removeAllRanges(); sel.addRange(saved.r); }
    else try { target.setSelectionRange(saved.s, saved.e); } catch { }
    return true;
  }
  function textBefore() {
    if (textInput(target)) return target.value.slice(0, target.selectionStart ?? target.value.length);
    const sel = getSelection(); if (!sel.rangeCount) return '';
    const r = sel.getRangeAt(0), pre = document.createRange();
    let n = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
    const block = (n && n.closest && n.closest('p,li,h1,h2,h3,h4,h5,h6,td,th,blockquote,pre,div')) || target;
    pre.setStart(target.contains(block) ? block : target, 0); pre.setEnd(r.startContainer, r.startOffset);
    const tw = document.createTreeWalker(pre.cloneContents(), NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
    let s = ''; for (let n; (n = tw.nextNode());) s += n.nodeType === 3 ? n.data : n.tagName === 'BR' ? '\n' : '';
    return s.slice(-200);
  }

  /* ---------- spoken commands, spacing and capitals ---------- */
  const CMDS = {
    'en-in': [[/(full stop|period|fullstop)/, '.'], [/comma/, ','], [/question mark/, '?'], [/(exclamation mark|exclamation point)/, '!'], [/(new paragraph|next paragraph)/, '\u00b6'], [/(new line|next line)/, '\u21b5']],
    hi: [[/(पूर्ण विराम|पूर्णविराम|फुल स्टॉप|फुलस्टॉप)/, '।'], [/(अल्पविराम|अल्प विराम|कॉमा)/, ','], [/(प्रश्नवाचक चिह्न|प्रश्नवाचक चिन्ह|प्रश्न चिह्न|प्रश्न चिन्ह|क्वेश्चन मार्क)/, '?'], [/(विस्मयादिबोधक चिह्न|विस्मयादिबोधक चिन्ह)/, '!'], [/(नया पैराग्राफ|नया अनुच्छेद|न्यू पैराग्राफ)/, '\u00b6'], [/(नई पंक्ति|नई लाइन|न्यू लाइन|अगली पंक्ति)/, '\u21b5']],
  };
  function tokens(text, lang) {
    let s = ' ' + text.trim().replace(/\s+/g, ' ') + ' ';
    for (const [re, t] of CMDS[lang]) s = s.replace(new RegExp('(?<=[\\s\\u0001])' + re.source + '(?=[\\s\\u0001])', 'g'), '\u0001' + t + '\u0001');
    return s.split('\u0001').map(x => x.trim()).filter(Boolean);
  }
  V._tokens = tokens;
  function insert(text) {
    if (!text || !text.trim()) return;
    if (!restore()) { window.BTS && BTS.toast && BTS.toast('Click where you want the text to go, then speak.'); return false; }
    const en = V.lang === 'en-in', single = target.tagName === 'INPUT';
    for (const tok of tokens(text, V.lang)) {
      if (tok === '\u21b5' || tok === '\u00b6') {   // new line / new paragraph
        const nl = tok === '\u21b5';
        if (single) document.execCommand('insertText', false, ' ');
        else if (textInput(target)) document.execCommand('insertText', false, nl ? '\n' : '\n\n');
        else document.execCommand(nl ? 'insertLineBreak' : 'insertParagraph');
        continue;
      }
      const raw = textBefore(), before = raw.replace(/[\s\u00a0]+$/, '');
      if (/^[.,?!।]$/.test(tok)) { document.execCommand('insertText', false, tok); continue; }
      let w = tok;
      if (en) { w = w.replace(/\bi\b/g, 'I').replace(/\bi'(m|ll|ve|d)\b/g, "I'$1"); if (!before || /[.?!]$/.test(before)) w = w[0].toUpperCase() + w.slice(1); }
      if (raw && !/[\s\u00a0]$/.test(raw)) w = ' ' + w;
      document.execCommand('insertText', false, w);
    }
    save();
    return true;
  }
  V._insert = insert;

  /* ---------- panel ---------- */
  let panel = null, state = 'idle', partialEl, statusEl, meterEl, micBtn, langSel, bodyEl;
  function css() {
    const st = document.createElement('style');
    st.textContent = `
.bts-voice{position:fixed;right:18px;bottom:40px;z-index:2500;width:320px;background:var(--surface);color:var(--ink);border:1px solid var(--line);border-radius:12px;box-shadow:var(--shadow-pop);font:13px var(--font-ui);user-select:none;animation:fade .12s ease-out}
.bts-voice header{display:flex;align-items:center;gap:8px;padding:9px 10px 0 14px;font-weight:600}
.bts-voice header .t{flex:1}
.bts-voice select{height:26px;border:1px solid var(--line-2);border-radius:6px;background:var(--surface);color:var(--ink);font:inherit;font-size:12px;padding:0 4px}
.bts-voice .x{width:26px;height:26px;border:0;background:transparent;border-radius:6px;color:var(--ink-2);display:grid;place-items:center}
.bts-voice .x:hover{background:var(--hover)}
.bts-voice .main{display:flex;align-items:center;gap:12px;padding:12px 14px 6px}
.bts-voice .mic{flex:none;width:46px;height:46px;border-radius:50%;border:0;background:var(--surface-3);color:var(--ink-2);display:grid;place-items:center;transition:background .15s,box-shadow .15s}
.bts-voice .mic:hover{filter:brightness(.97)}
.bts-voice.on .mic{background:var(--accent);color:var(--accent-ink);box-shadow:0 0 0 calc(3px + var(--lvl,0) * 9px) var(--accent-soft)}
.bts-voice .st{flex:1;min-width:0}
.bts-voice .st b{display:block;font-weight:600;margin-bottom:6px}
.bts-voice .meter{height:4px;border-radius:2px;background:var(--surface-3);overflow:hidden}
.bts-voice .meter i{display:block;height:100%;width:0;background:var(--accent);transition:width .08s}
.bts-voice .bar{height:6px;border-radius:3px;background:var(--surface-3);overflow:hidden;margin-top:2px}
.bts-voice .bar i{display:block;height:100%;width:0;background:var(--accent);transition:width .2s}
.bts-voice .partial{margin:4px 14px 0;min-height:38px;max-height:64px;overflow:hidden;color:var(--ink-3);font-style:italic;line-height:1.45;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical}
.bts-voice:not(.on) .partial:empty{display:none}
.bts-voice .hint{padding:6px 14px 11px;color:var(--ink-3);font-size:11.5px;line-height:1.4;border-top:1px solid var(--line);margin-top:6px}
.bts-voice .err{color:var(--danger)}
.bts-voice .acts{display:flex;gap:8px;margin-top:8px}
.bts-voice .acts .btn{height:28px;padding:0 12px;font-size:12.5px}`;
    document.head.appendChild(st);
  }
  function buildPanel() {
    if (panel) return;
    css();
    panel = document.createElement('div'); panel.className = 'bts-voice'; panel.setAttribute('role', 'dialog'); panel.setAttribute('aria-label', 'Voice typing');
    panel.innerHTML = `<header>${ic('mic', 16)}<span class="t">Voice typing</span><select title="Language">${Object.entries(LANGS).map(([k, l]) => `<option value="${k}">${esc(l.label)}</option>`).join('')}</select><button class="x" data-close title="Close">${ic('x', 16)}</button></header>
<div class="main"><button class="mic" title="Start or stop listening">${ic('mic', 22)}</button><div class="st"><b>Ready</b><div class="meter"><i></i></div></div></div>
<div class="partial"></div><div class="hint"></div>`;
    micBtn = panel.querySelector('.mic'); statusEl = panel.querySelector('.st'); partialEl = panel.querySelector('.partial'); langSel = panel.querySelector('select'); bodyEl = panel.querySelector('.main');
    panel.addEventListener('mousedown', (e) => { if (!e.target.closest('select')) e.preventDefault(); });   // keep the caret in the document
    micBtn.onclick = () => V.toggle();
    panel.querySelector('[data-close]').onclick = () => close();
    langSel.onchange = () => { const on = state === 'listening' || state === 'loading'; stop(); V.lang = langSel.value; store.set('bts-voice-lang', V.lang); hint(); if (on) start(V.lang); else restore(); };
    document.body.appendChild(panel);
    hint();
  }
  function hint() {
    if (!panel) return;
    langSel.value = V.lang;
    panel.querySelector('.hint').innerHTML = V.lang === 'hi'
      ? 'बोलें “पूर्ण विराम”, “अल्पविराम”, “नई पंक्ति”, “नया पैराग्राफ”। Works offline — your voice stays on this computer.'
      : 'Say “full stop”, “comma”, “question mark”, “new line” or “new paragraph”. Works offline — your voice stays on this computer.';
  }
  function setStatus(title, extra = '') {
    statusEl.innerHTML = `<b class="${state === 'error' ? 'err' : ''}">${esc(title)}</b>${extra}`;
    meterEl = statusEl.querySelector('.meter i');
    panel.classList.toggle('on', state === 'listening');
    micBtn.innerHTML = ic(state === 'listening' ? 'mic' : 'mic-off', 22);
    micBtn.title = state === 'listening' ? 'Stop listening' : 'Start listening';
  }
  const meterHTML = '<div class="meter"><i></i></div>';
  function showError(msg, retry = true) {
    state = 'error';
    setStatus('Voice typing stopped', `<div class="err" style="line-height:1.4;white-space:normal">${esc(msg)}</div>${retry ? '<div class="acts"><button class="btn" data-retry>Try again</button></div>' : ''}`);
    const b = statusEl.querySelector('[data-retry]'); if (b) b.onclick = () => start(V.lang);
  }

  /* ---------- model: download once (main process), load into the WebAssembly worker ---------- */
  let voskLoad = null, model = null, modelLang = null, cancelDl = null;
  const mb = (n) => (n / 1048576).toFixed(n < 10485760 ? 1 : 0);
  function loadVosk() {
    return voskLoad || (voskLoad = new Promise((res, rej) => {
      if (window.Vosk) return res();
      const s = document.createElement('script'); s.src = new URL('../vendor/vosk/vosk.js', SRC || location.href).href;
      s.onload = () => res(); s.onerror = () => { voskLoad = null; rej(new Error('Could not load the speech engine.')); };
      document.head.appendChild(s);
    }));
  }
  async function ensureDownloaded(lang) {
    const info = await bts.ext('voice-models');
    const m = info.models.find(x => x.lang === lang);
    if (m && m.downloaded) return true;
    const L = LANGS[lang];
    const ok = await BTS.dialog({
      title: 'Download speech model?',
      body: `<div style="max-width:430px;line-height:1.55;margin:2px 0 6px"><p style="margin:0 0 10px">To type with your voice in <b>${esc(L.label)}</b>, BTS Office needs to download a speech model once (<b>${L.mb} MB</b>).</p><p style="margin:0;color:var(--ink-2)">After that, voice typing works without internet, and what you say never leaves this computer.${m && m.partial ? ` ${mb(m.partial)} MB were already downloaded earlier — the download will continue from there.` : ''}</p></div>`,
      buttons: [{ label: 'Not now', value: false }, { label: `Download (${L.mb} MB)`, primary: true, value: true }],
    });
    if (!ok) return false;
    state = 'loading';
    setStatus(`Downloading ${L.label} model…`, `<div class="bar"><i></i></div><div class="acts" style="justify-content:space-between;align-items:center"><span class="pg" style="color:var(--ink-3);font-size:12px">Starting…</span><button class="btn" data-cancel>Cancel</button></div>`);
    const bar = statusEl.querySelector('.bar i'), pg = statusEl.querySelector('.pg');
    let cancelled = false;
    cancelDl = () => { cancelled = true; bts.ext('voice-cancel', lang); };
    statusEl.querySelector('[data-cancel]').onclick = () => cancelDl();
    const off = bts.onExt('voice-progress', (d) => {
      if (d.lang !== lang) return;
      if (d.phase === 'download') { bar.style.width = (100 * d.received / (d.total || 1)).toFixed(1) + '%'; pg.textContent = `${mb(d.received)} of ${mb(d.total)} MB`; }
      if (d.phase === 'extract') { bar.style.width = '100%'; pg.textContent = 'Setting up…'; }
    });
    const r = await bts.ext('voice-download', lang).finally(() => { off(); cancelDl = null; });
    if (r.ok) return true;
    if (r.cancelled || cancelled) { state = 'idle'; setStatus('Download cancelled', meterHTML); return false; }
    throw new Error(r.error || 'The download failed.');
  }
  function createModel(url) {
    return new Promise((res, rej) => {
      const m = new Vosk.Model(url, -1);
      m.on('load', (e) => (e.result ? res(m) : rej(new Error('The speech model could not be loaded.'))));
      m.on('error', (e) => { m.terminate(); rej(new Error('The speech model could not be loaded. Remove voice models in Options and download again. (' + e.error + ')')); });
    });
  }
  async function ensureModel(lang) {
    if (model && modelLang === lang) return model;
    if (model) { model.terminate(); model = null; modelLang = null; }
    if (!window.bts) throw new Error('Voice typing works in the BTS Office desktop app.');
    if (!(await ensureDownloaded(lang))) return null;
    state = 'loading'; setStatus(`Loading ${LANGS[lang].label} model…`, meterHTML);
    const [bytes] = await Promise.all([bts.ext('voice-model', lang), loadVosk()]);
    if (!bytes) throw new Error('The speech model is missing. Please try again.');
    const url = URL.createObjectURL(new Blob([bytes], { type: 'application/gzip' }));
    try { model = await createModel(url); modelLang = lang; } finally { URL.revokeObjectURL(url); }
    return model;
  }

  /* ---------- microphone -> recognizer ---------- */
  let audio = null, rec = null, runId = 0;
  async function start(lang) {
    lang = LANGS[lang] ? lang : V.lang;
    V.lang = lang; store.set('bts-voice-lang', lang);
    buildPanel(); hint();
    if (state === 'listening' || state === 'loading') return;
    const id = ++runId;
    state = 'loading'; setStatus('Getting ready…', meterHTML); partialEl.textContent = '';
    try {
      const m = await ensureModel(lang);
      if (id !== runId) return;
      if (!m) { if (state === 'loading') { state = 'idle'; setStatus('Ready', meterHTML); } return; }
      setStatus('Starting microphone…', meterHTML);
      const stream = V._testStream || await getMic();
      if (id !== runId) { if (!V._testStream) stream.getTracks().forEach(t => t.stop()); return; }
      const ctx = new AudioContext();
      const src = ctx.createMediaStreamSource(stream), proc = ctx.createScriptProcessor(4096, 1, 1), mute = ctx.createGain();
      mute.gain.value = 0;
      rec = new m.KaldiRecognizer(ctx.sampleRate);
      const r = rec;
      r.on('partialresult', (e) => { if (r === rec) partialEl.textContent = e.result.partial || ''; });
      r.on('result', (e) => { const t = e.result && e.result.text; if (r === rec) partialEl.textContent = ''; if (t) insert(t); });
      r.on('error', (e) => console.warn('[voice]', e.error));
      proc.onaudioprocess = (e) => {
        const d = e.inputBuffer.getChannelData(0);
        let sum = 0; for (let i = 0; i < d.length; i += 4) sum += d[i] * d[i];
        const lvl = Math.min(1, Math.sqrt(sum / (d.length / 4)) * 6);
        panel.style.setProperty('--lvl', lvl.toFixed(2)); if (meterEl) meterEl.style.width = (lvl * 100).toFixed(0) + '%';
        try { r.acceptWaveformFloat(d, ctx.sampleRate); } catch (err) { console.warn('[voice]', err); }
      };
      src.connect(proc); proc.connect(mute); mute.connect(ctx.destination);
      audio = { ctx, src, proc, stream };
      state = 'listening'; setStatus(`Listening — ${LANGS[lang].label}`, meterHTML);
      if (!target) partialEl.textContent = 'Click in your document where the text should go.';
    } catch (err) {
      if (id !== runId) return;
      release();
      showError(err.message || String(err));
    }
  }
  async function getMic() {
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) throw new Error('No microphone found. Plug in a microphone or headset and try again.');
    try {
      return await navigator.mediaDevices.getUserMedia({ audio: { channelCount: 1, echoCancellation: true, noiseSuppression: true, autoGainControl: true }, video: false });
    } catch (e) {
      if (/NotFound|DevicesNotFound|Overconstrained/i.test(e.name)) throw new Error('No microphone found. Plug in a microphone or headset and try again.');
      if (/NotAllowed|Security|PermissionDenied/i.test(e.name)) throw new Error('BTS Office is not allowed to use the microphone. Turn on microphone access for desktop apps in your computer’s privacy settings, then try again.');
      if (/NotReadable|TrackStart|Abort/i.test(e.name)) throw new Error('The microphone is busy or not working. Close other apps that use it and try again.');
      throw new Error('Could not start the microphone (' + (e.message || e.name) + ').');
    }
  }
  function release() {
    if (audio) {
      try { audio.proc.onaudioprocess = null; audio.src.disconnect(); audio.proc.disconnect(); } catch { }
      if (audio.stream !== V._testStream) audio.stream.getTracks().forEach(t => t.stop());
      audio.ctx.close().catch(() => { });
      audio = null;
    }
    if (rec) {   // flush the last words, then free the recognizer
      const r = rec; rec = null;
      try { r.retrieveFinalResult(); } catch { }
      setTimeout(() => { try { r.remove(); } catch { } }, 1500);
    }
    if (panel) { panel.style.setProperty('--lvl', 0); if (meterEl) meterEl.style.width = '0'; }
  }
  function stop() {
    runId++;
    if (cancelDl) cancelDl();
    release();
    if (panel && state !== 'error') { state = 'idle'; setStatus('Paused — click the microphone to continue', meterHTML); partialEl.textContent = ''; }
    else state = state === 'error' ? 'error' : 'idle';
  }
  function close() { stop(); state = 'idle'; if (panel) { panel.remove(); panel = null; } restore(); }
})();
