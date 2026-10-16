/* BTS Docs 1.2 — Review extras: Restrict editing / Mark as final, Accessibility checker, alt text,
 * Word count with reading time and Flesch reading ease. */
(function () {
  const D = window.Docs; const { $, $$ } = D; const esc = BTS.esc;

  /* ================= protection ================= */
  D.isProtected = () => !!(D.design && D.design.protect);
  async function sha256(s) { const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s)); return [...new Uint8Array(b)].map(x => x.toString(16).padStart(2, '0')).join(''); }
  D.sha256 = sha256;
  D.applyProtection = () => {
    const p = D.design.protect; const on = !!p;
    [D.doc, D.hdr, D.ftr].forEach(el => { if (!el) return; el.contentEditable = on ? 'false' : 'true'; });
    document.body.classList.toggle('doc-protected', on);
    const bar = $('#protect-bar'); if (!bar) return;
    if (!on) { bar.classList.add('hidden'); bar.innerHTML = ''; }
    else {
      bar.classList.remove('hidden');
      bar.innerHTML = p.mode === 'final'
        ? `${icon('badge-check', 16)}<span><b>Marked as final.</b> An author has marked this document as final to discourage editing.</span><span style="flex:1"></span><button class="btn" data-cmd="protect-edit-anyway" title="Remove the final mark and edit">Edit anyway</button>`
        : `${icon('lock', 16)}<span><b>Restricted editing.</b> This document is read-only${p.hash ? ' and protected with a password' : ''}. You can read, copy, search and print it.</span><span style="flex:1"></span><button class="btn" data-cmd="protect-stop" title="Stop protection">Stop protection</button>`;
    }
    $('#btn-final')?.classList.toggle('on', on && p.mode === 'final');
    $('#btn-restrict')?.classList.toggle('on', on && p.mode === 'readonly');
  };
  D.markFinal = () => {
    if (D.design.protect && D.design.protect.mode === 'final') { D.editAnyway(); return; }
    if (D.design.protect) { BTS.toast('Stop the editing restriction first.'); return; }
    D.design.protect = { mode: 'final' }; D.design.track = false; D.syncTrackUI && D.syncTrackUI();
    BTS.markDirty(); D.applyProtection(); BTS.toast('Marked as final — save to keep the mark');
  };
  D.editAnyway = () => { D.design.protect = null; BTS.markDirty(); D.applyProtection(); D.restoreFocus(); };
  D.restrictDialog = async () => {
    if (D.design.protect && D.design.protect.mode === 'readonly') return D.stopProtection();
    const v = await BTS.dialog({ title: 'Restrict editing', width: 480, body: `
      <p style="margin:0 0 10px;line-height:1.5">Make this document <b>read-only</b>. People can still read, copy, search and print it.</p>
      <div class="dlg-grid"><div class="field"><label>Password (optional)</label><input id="rp-1" type="password" autocomplete="new-password"></div><div class="field"><label>Confirm password</label><input id="rp-2" type="password" autocomplete="new-password"></div></div>
      <div class="hint" style="line-height:1.5">${icon('info', 14)} This is a <b>convenience lock</b> to prevent accidental changes — it is <b>not encryption</b>. The text stays readable and a determined person with a ZIP tool can remove the lock. Microsoft Word shows the document as read-only too (Word does not ask for this password).</div>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'Start enforcing protection', primary: true, validate: (rt) => { if ($('#rp-1', rt).value !== $('#rp-2', rt).value) { BTS.toast('The passwords don’t match.', { error: true }); return false; } return true; }, value: (rt) => $('#rp-1', rt).value }] });
    if (v == null) return;
    const hash = v ? await sha256('bts-docs:' + v) : null;
    D.design.protect = { mode: 'readonly', hash }; D.design.track = false; D.syncTrackUI && D.syncTrackUI();
    BTS.markDirty(); D.applyProtection(); BTS.toast('Editing is restricted');
  };
  D.stopProtection = async () => {
    const p = D.design.protect; if (!p) return;
    if (p.hash) {
      const v = await BTS.dialog({ title: 'Unprotect document', width: 380, body: '<div class="field"><label>Password</label><input id="up-1" type="password"></div>', buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (rt) => $('#up-1', rt).value }] });
      if (v == null) return;
      if ((await sha256('bts-docs:' + v)) !== p.hash) { BTS.toast('That password is not correct.', { error: true }); return; }
    }
    D.design.protect = null; BTS.markDirty(); D.applyProtection(); BTS.toast('Protection stopped — you can edit again'); D.restoreFocus();
  };
  /** Ribbon commands that never change the document (allowed while protected). */
  D.SAFE_CMDS = new Set(['copy', 'find', 'replace-close', 'find-close', 'find-next', 'find-prev', 'find-toggle-replace', 'select-all', 'word-count', 'read-aloud', 'print-layout', 'focus-mode', 'dark-canvas', 'ruler', 'nav-pane', 'nav-close', 'zoom-menu', 'zoom-page', 'zoom-width', 'goto-page', 'comments-toggle', 'comment-next', 'comment-prev', 'protect-edit-anyway', 'protect-stop', 'restrict-editing', 'mark-final', 'a11y-check', 'a11y-close', 'a11y-recheck', 'marks', 'marks-check', 'spelling', 'styles-pane', 'styles-close', 'tc-next', 'tc-prev', 'markup-menu', 'gal-up', 'gal-down', 'thgal-up', 'thgal-down', 'ssgal-up', 'ssgal-down', 'mm-preview', 'mm-first', 'mm-prev', 'mm-next', 'mm-last', 'mm-highlight', 'mm-edit-list', 'mm-recipients', 'mm-finish', 'footnote-next', 'manage-sources', 'legacy-dialog', 'indic-menu', 'indic-toggle', 'legacy-bar-close']);

  /* ================= alt text ================= */
  D.imgAlt = async (imgArg) => {
    const img = imgArg || D.selImg; if (!img) { BTS.toast('Click a picture first.'); return; }
    const v = await BTS.dialog({ title: 'Alt text', width: 460, body: `<div class="alt-prev"><img src="${esc(img.getAttribute('src'))}" alt=""></div><div class="field"><label>Describe this picture for people who are blind or have low vision (1–2 sentences)</label><textarea id="alt-t" style="height:84px">${esc(img.alt || '')}</textarea></div><label class="check"><input type="checkbox" id="alt-dec" ${img.dataset.decorative ? 'checked' : ''}> Mark as decorative (a border, divider or ornament with no meaning)</label>`, buttons: [{ label: 'Cancel', value: null }, { label: 'Save', primary: true, value: (rt) => ({ t: $('#alt-t', rt).value.trim(), dec: $('#alt-dec', rt).checked }) }] });
    if (!v) return false;
    D.change(() => { img.alt = v.dec ? '' : v.t; if (v.dec) img.dataset.decorative = '1'; else delete img.dataset.decorative; });
    return true;
  };

  /* ================= accessibility checker ================= */
  const lum = (rgb) => { const [r, g, b] = rgb.map(v => { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); }); return 0.2126 * r + 0.7152 * g + 0.0722 * b; };
  const parseRGB = (c) => { const m = String(c).match(/rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?/); if (!m) return null; return { rgb: [+m[1], +m[2], +m[3]], a: m[4] === undefined ? 1 : +m[4] }; };
  const hexRGB = (h) => { const n = parseInt(String(h).replace('#', ''), 16); return [n >> 16, (n >> 8) & 255, n & 255]; };
  D.contrast = (a, b) => { const L1 = lum(a), L2 = lum(b); return (Math.max(L1, L2) + 0.05) / (Math.min(L1, L2) + 0.05); };
  function bgOf(el) { while (el && el !== document.body) { const c = parseRGB(getComputedStyle(el).backgroundColor); if (c && c.a > 0.5) return c.rgb; if (el === D.sheet) break; el = el.parentElement; } return hexRGB(D.design.pageColor || '#ffffff'); }
  function hasHeaderRow(t) { if (t.rows[0] && [...t.rows[0].cells].every(c => c.tagName === 'TH')) return true; return /\bt-(header|banded|dark|grid-accent|light|minimal|hrow)\b/.test(t.className) && !t.classList.contains('t-noheader'); }
  const label = (el) => { const t = (el.textContent || '').trim().replace(/\s+/g, ' '); return t.length > 60 ? t.slice(0, 60) + '…' : t; };
  D.checkA11y = () => {
    const out = [];
    // pictures
    [D.doc, D.hdr, D.ftr].forEach(root => $$('img', root).forEach((img, i) => { if (!img.alt && !img.dataset.decorative) out.push({ sev: 'error', kind: 'alt', title: 'Missing alternative text', detail: (root === D.doc ? 'Picture' : root === D.hdr ? 'Header picture' : 'Footer picture') + ' ' + (i + 1) + (img.getAttribute('src') || '').slice(0, 0), el: img, fix: 'Add alt text' }); }));
    // tables
    $$('table', D.doc).forEach((t, i) => { if (t.closest('.textbox')) return; if (!hasHeaderRow(t)) out.push({ sev: 'error', kind: 'table', title: 'No header row specified', detail: 'Table ' + (i + 1) + ': ' + (t.rows[0] ? [...t.rows[0].cells].map(c => c.textContent.trim()).join(' · ').slice(0, 60) : ''), el: t, fix: 'Use first row as header' }); });
    // headings
    let prev = 0;
    $$('h1,h2,h3,h4,h5,h6', D.doc).forEach(h => {
      if (h.closest('.toc,.biblio')) return;
      const lvl = +h.tagName[1];
      if (!h.textContent.replace(/​/g, '').trim()) { out.push({ sev: 'error', kind: 'empty', title: 'Empty heading', detail: 'Heading ' + lvl + ' has no text — screen readers announce it as a blank heading', el: h, fix: 'Change to Normal' }); return; }
      if ((prev === 0 && lvl > 2) || (prev && lvl > prev + 1)) out.push({ sev: 'warn', kind: 'skip', title: 'Skipped heading level', detail: `Heading ${lvl} follows ${prev ? 'Heading ' + prev : 'the start of the document'}: “${label(h)}”`, el: h, fix: `Change to Heading ${Math.max(1, prev + 1)}`, to: Math.max(1, prev + 1) });
      prev = lvl;
    });
    // contrast (one finding per block)
    const seen = new Set(); const tw = document.createTreeWalker(D.doc, NodeFilter.SHOW_TEXT); let n, k = 0;
    while ((n = tw.nextNode()) && k++ < 6000) {
      if (!n.data.trim()) continue; const el = n.parentElement; if (!el || el.closest('.tc-ins,.tc-del,.toc,.cmt,.mf,del')) continue;
      const blk = D.blockOf(n) || el; if (seen.has(blk)) continue;
      const cs = getComputedStyle(el); const fg = parseRGB(cs.color); if (!fg) continue;
      const ratio = D.contrast(fg.rgb, bgOf(el)); const pt = parseFloat(cs.fontSize) * 0.75; const large = pt >= 18 || (pt >= 14 && +cs.fontWeight >= 700);
      if (ratio < (large ? 3 : 4.5)) { seen.add(blk); out.push({ sev: 'warn', kind: 'contrast', title: 'Hard-to-read text contrast', detail: `Contrast ${ratio.toFixed(2)}:1 (needs ${large ? '3' : '4.5'}:1): “${label(el)}”`, el, fix: 'Use automatic colour' }); }
    }
    // long all-caps paragraphs
    $$('p,li,h1,h2,h3,h4,h5,h6', D.doc).forEach(b => {
      if (D.blockOf(b) !== b || b.closest('.toc,.biblio')) return;
      const t = b.textContent; const letters = t.match(/\p{L}/gu) || []; if (letters.length < 60) return;
      const upper = getComputedStyle(b).textTransform === 'uppercase' ? letters.length : letters.filter(c => c !== c.toLowerCase() && c === c.toUpperCase()).length;
      if (upper / letters.length > 0.9) out.push({ sev: 'warn', kind: 'caps', title: 'Long paragraph in all capitals', detail: `“${label(b)}” — capitals are harder to read`, el: b, fix: 'Change to sentence case' });
    });
    // tips
    if (!$$('h1,h2,h3', D.doc).filter(h => h.textContent.trim()).length && D.countWords(D.docText()) > 400) out.push({ sev: 'tip', kind: 'structure', title: 'Use headings to organise the document', detail: 'Long documents are easier to navigate with Heading styles.', el: D.doc.firstElementChild });
    if (!(BTS.props.title || '').trim()) out.push({ sev: 'tip', kind: 'title', title: 'Document title is missing', detail: 'Add a title on File ▸ Info so screen readers announce it.', el: null, fix: 'Add title' });
    $$('a[href]', D.doc).forEach(a => { const t = a.textContent.trim(); if (/^(click here|here|link|read more|more)$/i.test(t) || /^https?:\/\/\S{40,}$/.test(t)) out.push({ sev: 'tip', kind: 'link', title: 'Unclear link text', detail: `“${label(a)}” doesn’t say where the link goes`, el: a }); });
    return out;
  };
  let findings = [];
  D.toggleA11y = (on) => {
    const p = $('#a11y-pane'); const show = on ?? p.classList.contains('hidden');
    p.classList.toggle('hidden', !show); $('#btn-a11y')?.classList.toggle('on', show);
    if (show) D.renderA11y(); D.scheduleLayout(30);
  };
  D.renderA11y = () => {
    const p = $('#a11y-pane'); if (!p || p.classList.contains('hidden')) return;
    findings = D.checkA11y();
    const groups = [['error', 'Errors', 'circle-x'], ['warn', 'Warnings', 'triangle-alert'], ['tip', 'Tips', 'lightbulb']];
    const body = findings.length ? groups.map(([sev, name, ic]) => { const f = findings.map((x, i) => [x, i]).filter(([x]) => x.sev === sev); if (!f.length) return ''; return `<div class="a11y-group a11y-${sev}"><div class="a11y-gh">${icon(ic, 15)} ${name} <span>${f.length}</span></div>${f.map(([x, i]) => `<div class="a11y-item" data-i="${i}" title="Click to go to it"><div class="a11y-t">${esc(x.title)}</div><div class="a11y-d">${esc(x.detail)}</div>${x.fix ? `<button class="btn a11y-fix" data-fix="${i}">${esc(x.fix)}</button>` : ''}</div>`).join('')}</div>`; }).join('')
      : `<div class="a11y-ok">${icon('circle-check-big', 34)}<b>No accessibility issues found</b><span>People with disabilities should not have difficulty reading this document.</span></div>`;
    $('#a11y-list').innerHTML = `<div class="a11y-sum">${findings.filter(f => f.sev === 'error').length} errors · ${findings.filter(f => f.sev === 'warn').length} warnings · ${findings.filter(f => f.sev === 'tip').length} tips</div>` + body;
  };
  async function fix(i) {
    const f = findings[i]; if (!f) return;
    if (f.kind === 'alt') { await D.imgAlt(f.el); }
    else if (f.kind === 'table') D.change(() => { f.el.classList.remove('t-noheader'); f.el.classList.add('t-hrow'); });
    else if (f.kind === 'empty') D.change(() => { D._setBlockStyle(f.el, 'normal'); });
    else if (f.kind === 'skip') D.change(() => { D._setBlockStyle(f.el, 'h' + f.to); });
    else if (f.kind === 'contrast') D.change(() => { let el = f.el; const blk = D.blockOf(el); while (el && el !== blk && !D.isRoot(el)) { el.style.color = ''; if (el.tagName === 'SPAN' && !el.getAttribute('style')) { const p = el.parentElement; D.unwrap(el); el = p; continue; } el = el.parentElement; } if (blk) blk.style.color = ''; });
    else if (f.kind === 'caps') { D.scrollToEl(f.el); D.selectNode(f.el, true); D.changeCase('sentence'); }
    else if (f.kind === 'title') { const t = await BTS.prompt('Document title', 'Title', BTS.title || ''); if (t != null) { BTS.setProps({ title: t }); BTS.markDirty(); } }
    D.renderA11y();
  }
  function go(i) { const f = findings[i]; if (!f || !f.el || !f.el.isConnected) return; D.scrollToEl(f.el); if (f.el.tagName === 'IMG') { if (D.doc.contains(f.el)) D.selectImage(f.el); } else { const r = document.createRange(); r.selectNodeContents(f.el); const root = D.rootOf(f.el) || D.doc; root.focus({ preventScroll: true }); const s = getSelection(); s.removeAllRanges(); s.addRange(r); } f.el.classList.add('a11y-flash'); setTimeout(() => f.el.classList.remove('a11y-flash'), 1200); }
  function wireA11y() {
    const l = $('#a11y-list'); if (!l) return;
    l.addEventListener('mousedown', e => { if (e.target.closest('.a11y-item')) e.preventDefault(); });
    l.addEventListener('click', e => { const b = e.target.closest('[data-fix]'); if (b) { fix(+b.dataset.fix); return; } const it = e.target.closest('.a11y-item'); if (it) go(+it.dataset.i); });
  }

  /* ================= word count & readability ================= */
  const syllables = (w) => { w = w.toLowerCase().replace(/[^a-z]/g, ''); if (!w) return 0; if (w.length <= 3) return 1; w = w.replace(/(?:[^laeiouy]es|ed|[^laeiouy]e)$/, '').replace(/^y/, ''); const m = w.match(/[aeiouy]{1,2}/g); return Math.max(1, m ? m.length : 1); };
  D.readability = (text) => {
    const words = text.match(/[A-Za-z][A-Za-z'’-]*/g) || []; const sentences = Math.max(1, (text.match(/[^.!?।]+[.!?।]+(\s|$)/g) || []).length || (words.length ? 1 : 0));
    if (words.length < 10) return null;
    const syl = words.reduce((a, w) => a + syllables(w), 0);
    const score = 206.835 - 1.015 * (words.length / sentences) - 84.6 * (syl / words.length);
    const s = Math.max(0, Math.min(100, score));
    const band = s >= 90 ? 'Very easy' : s >= 80 ? 'Easy' : s >= 70 ? 'Fairly easy' : s >= 60 ? 'Plain English' : s >= 50 ? 'Fairly difficult' : s >= 30 ? 'Difficult (college level)' : 'Very difficult (graduate level)';
    const grade = 0.39 * (words.length / sentences) + 11.8 * (syl / words.length) - 15.59;
    return { score: Math.round(s * 10) / 10, band, grade: Math.max(0, Math.round(grade * 10) / 10), sentences, avgWords: Math.round(words.length / sentences * 10) / 10 };
  };
  D.wordCountDialog = () => {
    const s = D.stats(); const sel = D.selectedText(); const text = sel && sel.trim() ? sel : (D.doc.innerText || '');
    const words = sel && sel.trim() ? D.countWords(sel) : s.words;
    const mins = words / 238; const rt = mins < 1 ? 'Less than a minute' : Math.round(mins) + ' minute' + (Math.round(mins) === 1 ? '' : 's');
    const speak = words / 130; const st = speak < 1 ? 'Less than a minute' : Math.round(speak) + ' minute' + (Math.round(speak) === 1 ? '' : 's');
    const rd = D.readability(text);
    const row = (k, v) => `<tr><td>${k}</td><td>${typeof v === 'number' ? v.toLocaleString('en-IN') : v}</td></tr>`;
    BTS.dialog({ title: 'Word count', width: 440, body: `<table class="wc-table">${sel ? row('Words in selection', D.countWords(sel)) : ''}${row('Pages', s.pages)}${row('Words', s.words)}${row('Characters (no spaces)', s.charsNoSpace)}${row('Characters (with spaces)', s.chars)}${row('Paragraphs', s.paras)}${row('Lines', s.lines)}</table>
      <div class="wc-h">Reading</div><table class="wc-table">${row('Reading time', rt + ' <small>(238 words/min)</small>')}${row('Speaking time', st + ' <small>(130 words/min)</small>')}${rd ? row('Flesch reading ease', `<b>${rd.score}</b> — ${rd.band}`) + row('Flesch–Kincaid grade', rd.grade) + row('Sentences', rd.sentences) + row('Words per sentence', rd.avgWords) : row('Readability', 'Needs at least 10 English words')}</table>
      ${rd ? `<div class="wc-meter" title="Flesch reading ease: higher is easier"><span style="left:${rd.score}%"></span></div><div class="wc-scale"><span>Hard</span><span>Easy</span></div>` : ''}`, buttons: [{ label: 'Close', primary: true, value: true }] });
  };

  D.initReview = () => { wireA11y(); };
})();
