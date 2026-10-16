/* BTS Docs — Insert features: pictures (resize handles, wrap), shapes, links, symbols, equations,
 * date & time, text box, horizontal line, drop cap, header/footer & page numbers, TOC, footnotes,
 * captions, comments. */
(function () {
  const D = window.Docs; const { $, $$, PX } = D; const esc = BTS.esc;

  /* ================= pictures ================= */
  D.sniffImage = (bytes) => {
    if (bytes[0] === 0x89 && bytes[1] === 0x50) return 'image/png';
    if (bytes[0] === 0xFF && bytes[1] === 0xD8) return 'image/jpeg';
    if (bytes[0] === 0x47 && bytes[1] === 0x49) return 'image/gif';
    if (bytes[0] === 0x42 && bytes[1] === 0x4D) return 'image/bmp';
    if (bytes[0] === 0x52 && bytes[8] === 0x57) return 'image/webp';
    const head = new TextDecoder().decode(bytes.slice(0, 200)); if (/<svg/i.test(head)) return 'image/svg+xml';
    return 'image/png';
  };
  /** Convert formats Word can't take (webp/svg) to PNG data URL. */
  D.toWordFriendly = (url) => new Promise(res => {
    if (/^data:image\/(png|jpeg|gif|bmp)/.test(url)) return res(url);
    const im = new Image(); im.onload = () => { const c = document.createElement('canvas'); c.width = im.naturalWidth || 400; c.height = im.naturalHeight || 300; c.getContext('2d').drawImage(im, 0, 0, c.width, c.height); res(c.toDataURL('image/png')); }; im.onerror = () => res(url); im.src = url;
  });
  D.insertImageURL = async (url, alt = '') => {
    url = await D.toWordFriendly(url);
    const dims = await new Promise(r => { const im = new Image(); im.onload = () => r([im.naturalWidth, im.naturalHeight]); im.onerror = () => r([300, 200]); im.src = url; });
    const maxW = (D.paperSize().w - D.page.m.l - D.page.m.r) * PX;
    let [w, h] = dims; if (w > maxW) { h = h * maxW / w; w = maxW; }
    D.change(() => {
      const img = document.createElement('img'); img.src = url; img.alt = alt; img.style.width = Math.round(w) + 'px'; img.style.height = Math.round(h) + 'px';
      img.dataset.nw = dims[0]; img.dataset.nh = dims[1];
      const r = D.ensureSel(); r.deleteContents(); r.insertNode(img);
      const after = document.createRange(); after.setStartAfter(img); after.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(after);
      setTimeout(() => D.selectImage(img), 30);
    });
  };
  D.pickPicture = async (replace) => {
    const f = await BTS.pickFile('image/*'); if (!f) return;
    const url = await BTS.bytesToDataURL(f.data, f.type || D.sniffImage(f.data));
    if (replace && D.selImg) { const img = D.selImg; const u2 = await D.toWordFriendly(url); D.change(() => { img.src = u2; img.style.height = 'auto'; }); setTimeout(() => { img.style.height = img.offsetHeight + 'px'; D.updateImgHandles(); }, 50); return; }
    D.insertImageURL(url, f.name.replace(/\.[^.]+$/, ''));
  };

  D.selImg = null;
  D.selectImage = (img) => {
    D.selImg = img;
    $('#tab-picture').classList.toggle('hidden', !img);
    if (img) { BTS.showTab('picture'); syncPicture(); }
    else if (BTS.activeTab === 'picture') BTS.showTab('home');
    D.updateImgHandles();
  };
  function syncPicture() {
    const img = D.selImg; if (!img) return;
    const sab = $('#btn-img-smartart'); if (sab) sab.classList.toggle('hidden', !(D.isSmartArt && D.isSmartArt(img)));
    $('#in-img-w').value = (img.offsetWidth / PX * 2.54).toFixed(2); $('#in-img-h').value = (img.offsetHeight / PX * 2.54).toFixed(2);
    const wrap = img.classList.contains('img-left') ? 'left' : img.classList.contains('img-right') ? 'right' : img.classList.contains('img-center') ? 'center' : 'inline';
    $$('[data-panel="picture"] [data-wrap]').forEach(b => b.classList.toggle('on', b.dataset.wrap === wrap));
    $('#btn-img-border').classList.toggle('on', img.classList.contains('img-border'));
    $('#btn-img-rounded').classList.toggle('on', img.classList.contains('img-rounded'));
    $('#btn-img-shadow').classList.toggle('on', img.classList.contains('img-shadow'));
  }
  D.syncPicture = syncPicture;
  D.updateImgHandles = () => {
    const h = $('#img-handles'); const img = D.selImg;
    if (!img || !img.isConnected) { h.classList.add('hidden'); if (D.selImg && !D.selImg.isConnected) { D.selImg = null; $('#tab-picture').classList.add('hidden'); } return; }
    const sr = D.sheet.getBoundingClientRect(); const ir = img.getBoundingClientRect(); const s = D.zoom / 100;
    Object.assign(h.style, { left: (ir.left - sr.left) / s + 'px', top: (ir.top - sr.top) / s + 'px', width: ir.width / s + 'px', height: ir.height / s + 'px' });
    h.classList.remove('hidden');
  };
  D.imgOp = (op, v) => {
    const img = D.selImg; if (!img) return;
    D.change(() => {
      if (op === 'wrap') { img.classList.remove('img-left', 'img-right', 'img-center'); if (v !== 'inline') img.classList.add('img-' + v); }
      else if (op === 'toggle') img.classList.toggle(v);
      else if (op === 'w' || op === 'h') {
        const px = Math.max(8, v / 2.54 * PX); const ratio = img.offsetHeight / img.offsetWidth; const lock = $('#chk-img-lock').checked;
        if (op === 'w') { img.style.width = Math.round(px) + 'px'; if (lock) img.style.height = Math.round(px * ratio) + 'px'; }
        else { img.style.height = Math.round(px) + 'px'; if (lock) img.style.width = Math.round(px / ratio) + 'px'; }
      }
      else if (op === 'reset') { const nw = +img.dataset.nw || img.naturalWidth, nh = +img.dataset.nh || img.naturalHeight; const maxW = (D.paperSize().w - D.page.m.l - D.page.m.r) * PX; const k = Math.min(1, maxW / nw); img.style.width = Math.round(nw * k) + 'px'; img.style.height = Math.round(nh * k) + 'px'; }
      else if (op === 'delete') { img.remove(); D.selImg = null; }
    });
    if (op === 'delete') D.selectImage(null); else setTimeout(() => { syncPicture(); D.updateImgHandles(); }, 20);
  };
  D.imgAlt = async () => { const img = D.selImg; if (!img) return; const v = await BTS.prompt('Alt text', 'Describe this picture for people who use screen readers', img.alt || ''); if (v != null) D.change(() => { img.alt = v; }); };
  function wireImages() {
    D.doc.addEventListener('mousedown', e => { if (e.target.tagName === 'IMG') { e.preventDefault(); D.selectImage(e.target); const r = document.createRange(); r.selectNode(e.target); const s = getSelection(); s.removeAllRanges(); s.addRange(r); D.doc.focus({ preventScroll: true }); } else if (!e.target.closest('#img-handles')) { if (D.selImg) D.selectImage(null); } });
    [D.hdr, D.ftr].forEach(el => el.addEventListener('mousedown', e => { if (e.target.tagName === 'IMG') { e.preventDefault(); D.selectImage(e.target); } else if (D.selImg) D.selectImage(null); }));
    const h = $('#img-handles'); let drag = null;
    h.addEventListener('mousedown', e => {
      const k = e.target.closest('i[data-h]'); if (!k || !D.selImg) return; e.preventDefault(); e.stopPropagation();
      D.checkpoint(); const img = D.selImg;
      drag = { k: k.dataset.h, x0: e.clientX, y0: e.clientY, w0: img.offsetWidth, h0: img.offsetHeight, s: D.zoom / 100 }; h.classList.add('dragging');
    });
    addEventListener('mousemove', e => {
      if (!drag) return; const img = D.selImg; if (!img) return;
      let dx = (e.clientX - drag.x0) / drag.s, dy = (e.clientY - drag.y0) / drag.s;
      if (drag.k.includes('w')) dx = -dx; if (drag.k.includes('n')) dy = -dy;
      const maxW = (D.paperSize().w * PX);
      let w = drag.w0, hh = drag.h0;
      if (drag.k === 'e') w = drag.w0 + dx; else if (drag.k === 's') hh = drag.h0 + dy;
      else { const k = Math.max((drag.w0 + dx) / drag.w0, (drag.h0 + dy) / drag.h0); w = drag.w0 * k; hh = drag.h0 * k; }
      w = Math.max(12, Math.min(maxW, w)); hh = Math.max(12, hh);
      img.style.width = Math.round(w) + 'px'; img.style.height = Math.round(hh) + 'px';
      $('#img-size-tip').textContent = (w / PX * 2.54).toFixed(2) + ' × ' + (hh / PX * 2.54).toFixed(2) + ' cm';
      D.updateImgHandles();
    });
    addEventListener('mouseup', () => { if (!drag) return; drag = null; h.classList.remove('dragging'); syncPicture(); D.afterChange(); });
    D.canvas.addEventListener('scroll', () => D.updateImgHandles());
  }

  /* ================= shapes (inline SVG images) ================= */
  const SHAPES = {
    rect: '<rect x="4" y="4" width="192" height="112" rx="2"/>', rounded: '<rect x="4" y="4" width="192" height="112" rx="22"/>', ellipse: '<ellipse cx="100" cy="60" rx="96" ry="56"/>',
    triangle: '<path d="M100 4 196 116H4z"/>', diamond: '<path d="M100 4 196 60 100 116 4 60z"/>', star: '<path d="m100 4 26 38 46 6-33 31 9 34-48-20-48 20 9-34-33-31 46-6z"/>',
    arrow: '<path d="M4 40h120V8l72 52-72 52V80H4z"/>', larrow: '<path d="M196 40H76V8L4 60l72 52V80h120z"/>', callout: '<path d="M8 4h184v80H80l-40 32 8-32H8z"/>', line: '<path d="M4 60h192" fill="none"/>',
  };
  D.shapesMenu = (anchor) => {
    const el = document.createElement('div');
    el.innerHTML = '<div class="mprev-head">Basic shapes</div><div class="shape-grid">' + Object.entries(SHAPES).map(([k, p]) => `<button data-s="${k}" title="${k}"><svg viewBox="0 0 200 120" width="34" height="22" fill="none" stroke="currentColor" stroke-width="10">${p}</svg></button>`).join('') + '</div>';
    el.addEventListener('mousedown', e => e.preventDefault());
    el.addEventListener('click', e => { const b = e.target.closest('[data-s]'); if (b) { BTS.resolvePopover(b.dataset.s); } });
    BTS.popover(anchor, el).then(async k => {
      if (!k) return;
      const c = await BTS.colorPicker(anchor, { none: 'Outline only', value: D.theme().accent }); if (c === null) return;
      const fill = c === 'none' ? 'none' : c; const stroke = c === 'none' ? '#1f2533' : shade(c);
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="400" height="240" viewBox="0 0 200 120"><g fill="${fill}" stroke="${stroke}" stroke-width="${k === 'line' ? 4 : 3}" stroke-linejoin="round">${SHAPES[k]}</g></svg>`;
      // rasterize to PNG so it survives DOCX export
      const url = await new Promise(res => { const im = new Image(); im.onload = () => { const cv = document.createElement('canvas'); cv.width = 400; cv.height = 240; cv.getContext('2d').drawImage(im, 0, 0, 400, 240); res(cv.toDataURL('image/png')); }; im.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg); });
      const img = await D.insertImageURL(url, k + ' shape');
      setTimeout(() => { if (D.selImg) { D.selImg.style.width = '200px'; D.selImg.style.height = '120px'; D.updateImgHandles(); } }, 60);
    });
  };
  function shade(hex) { const n = parseInt(hex.slice(1), 16); const f = (x) => Math.max(0, Math.round(x * 0.7)); return '#' + [f(n >> 16), f((n >> 8) & 255), f(n & 255)].map(x => x.toString(16).padStart(2, '0')).join(''); }

  /* ================= links ================= */
  D.linkDialog = async () => {
    D.ensureSel(); const r = D.range(); const sv = D.saveSel();
    const a = r && D.closest(r.startContainer, 'a');
    const text = a ? a.textContent : (r ? r.toString() : '');
    const heads = $$('h1,h2,h3', D.doc).filter(h => h.textContent.trim());
    const v = await BTS.dialog({
      title: a ? 'Edit link' : 'Insert link', width: 460,
      body: `<div class="field"><label>Text to display</label><input id="ln-text" value="${esc(text)}"></div>
        <div class="field"><label>Address (web page, e-mail or heading)</label><input id="ln-url" list="ln-heads" placeholder="https://example.com" value="${esc(a ? a.getAttribute('href') : '')}"><datalist id="ln-heads">${heads.map(h => `<option value="#${esc(ensureId(h))}">${esc(h.textContent.trim())}</option>`).join('')}</datalist></div>
        <div class="hint">Tip: type an e-mail address and it becomes a mailto: link. Choose a heading from the list to link inside this document.</div>`,
      buttons: [...(a ? [{ label: 'Remove link', value: 'remove', danger: true }] : []), { label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (rt) => ({ text: $('#ln-text', rt).value, url: $('#ln-url', rt).value.trim() }) }],
    });
    D.restoreSel(sv);
    if (v === 'remove') { D.change(() => { D.unwrap(a); }); return; }
    if (!v || !v.url) return;
    let url = v.url;
    if (/^[^@\s]+@[^@\s]+\.[a-z]{2,}$/i.test(url)) url = 'mailto:' + url;
    else if (!/^(https?:|mailto:|#|ftp:|tel:)/i.test(url)) url = 'https://' + url;
    D.change(() => {
      if (a) { a.href = url; if (v.text && v.text !== a.textContent) a.textContent = v.text; return; }
      const rr = D.range();
      if (rr && !rr.collapsed && (!v.text || v.text === text)) { D.exec('createLink', url); return; }
      const el = document.createElement('a'); el.href = url; el.textContent = v.text || v.url;
      if (rr) { rr.deleteContents(); rr.insertNode(el); D.setCaret(el.parentNode, [...el.parentNode.childNodes].indexOf(el) + 1); }
    });
  };
  function ensureId(h) { if (!h.id) h.id = '_h' + Math.random().toString(36).slice(2, 8); return h.id; }
  D.ensureId = ensureId;
  D.followLink = (a) => {
    const href = a.getAttribute('href') || '';
    if (href.startsWith('#')) { const t = document.getElementById(href.slice(1)); if (t) D.scrollToEl(t, true); return; }
    if (BTS.desktop) bts.openExternal(href); else window.open(href, '_blank');
  };

  /* ================= symbols ================= */
  const SYMS = {
    'Common': '₹ © ® ™ ° ± × ÷ √ ∞ ≈ ≠ ≤ ≥ € £ $ ¥ ¢ § ¶ † ‡ • … – — ‘ ’ “ ” « » ¡ ¿ ½ ⅓ ¼ ¾ ‰ µ № ✓ ✗ ★ ☆ ♥ ☎ ✉ ☺',
    'Math': '+ − × ÷ = ≠ ≈ ≡ < > ≤ ≥ ± ∓ √ ∛ ∜ ∞ ∝ ∑ ∏ ∫ ∬ ∮ ∂ ∇ Δ ∆ ∈ ∉ ⊂ ⊃ ⊆ ⊇ ∪ ∩ ∅ ∀ ∃ ¬ ∧ ∨ ⊕ ⊗ ⊥ ∥ ∠ ° ′ ″ ∴ ∵ ℕ ℤ ℚ ℝ ℂ π ‰ ⁰ ¹ ² ³ ⁴ ⁿ ₀ ₁ ₂ ₃',
    'Greek': 'α β γ δ ε ζ η θ ι κ λ μ ν ξ ο π ρ σ ς τ υ φ χ ψ ω Α Β Γ Δ Ε Ζ Η Θ Ι Κ Λ Μ Ν Ξ Ο Π Ρ Σ Τ Υ Φ Χ Ψ Ω',
    'Arrows': '← → ↑ ↓ ↔ ↕ ⇐ ⇒ ⇑ ⇓ ⇔ ↖ ↗ ↘ ↙ ↩ ↪ ⟵ ⟶ ⟷ ⟸ ⟹ ⟺ ➔ ➜ ➤ ↺ ↻ ⇄ ⇆',
    'Currency': '₹ $ € £ ¥ ₩ ₽ ₺ ₫ ₦ ₱ ₪ ฿ ₴ ₸ ₡ ¢ ₿',
    'Devanagari': 'अ आ इ ई उ ऊ ए ऐ ओ औ क ख ग घ ङ च छ ज झ ञ ट ठ ड ढ ण त थ द ध न प फ ब भ म य र ल व श ष स ह ळ क्ष ज्ञ ा ि ी ु ू े ै ो ौ ं ः ् । ॥ ० १ २ ३ ४ ५ ६ ७ ८ ९ ॐ',
    'Shapes': '■ □ ▪ ▫ ▲ △ ▼ ▽ ◆ ◇ ○ ● ◎ ◉ ★ ☆ ✦ ✧ ♠ ♣ ♥ ♦ ☐ ☑ ☒ ✔ ✘ ❖ ❤',
  };
  const NAMES = { '₹': 'Indian Rupee sign', '©': 'Copyright', '®': 'Registered', '™': 'Trade mark', '°': 'Degree', '±': 'Plus-minus', '×': 'Multiplication', '÷': 'Division', '√': 'Square root', '∞': 'Infinity', '≈': 'Almost equal', '≠': 'Not equal', '≤': 'Less than or equal', '≥': 'Greater than or equal', 'π': 'Pi', '∑': 'Summation', '∫': 'Integral', 'Δ': 'Delta', '→': 'Right arrow', '←': 'Left arrow', '…': 'Ellipsis', '—': 'Em dash', '–': 'En dash', '§': 'Section', '¶': 'Pilcrow', '•': 'Bullet', 'µ': 'Micro', '½': 'One half', 'ॐ': 'Om', '।': 'Danda' };
  let recentSyms = (() => { try { return JSON.parse(localStorage.getItem('docs-recent-syms') || '[]'); } catch { return []; } })();
  D.symbolDialog = async () => {
    const sv = D.saveSel();
    let chosen = null;
    const body = document.createElement('div');
    const cats = Object.keys(SYMS);
    body.innerHTML = `<div class="sym-tabs">${(recentSyms.length ? ['Recent'] : []).concat(cats).map((c, i) => `<button data-cat="${c}" class="${i === 0 ? 'on' : ''}">${c}</button>`).join('')}</div><div class="sym-grid"></div><div class="sym-info"></div>`;
    const grid = $('.sym-grid', body), info = $('.sym-info', body);
    const show = (cat) => { const list = cat === 'Recent' ? recentSyms : SYMS[cat].split(' '); grid.innerHTML = list.map(s => `<button data-s="${esc(s)}" title="${esc(NAMES[s] || 'U+' + s.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'))}">${esc(s)}</button>`).join(''); $$('.sym-tabs button', body).forEach(b => b.classList.toggle('on', b.dataset.cat === cat)); };
    show(recentSyms.length ? 'Recent' : cats[0]);
    body.addEventListener('click', e => {
      const c = e.target.closest('[data-cat]'); if (c) show(c.dataset.cat);
      const s = e.target.closest('[data-s]'); if (s) { $$('.sym-grid button', body).forEach(b => b.classList.remove('on')); s.classList.add('on'); chosen = s.dataset.s; info.textContent = (NAMES[chosen] || 'Character') + ' · U+' + chosen.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'); }
    });
    body.addEventListener('dblclick', e => { const s = e.target.closest('[data-s]'); if (s) { chosen = s.dataset.s; $('footer .primary', body.closest('.modal')).click(); } });
    const v = await BTS.dialog({ title: 'Symbol', body, width: 520, buttons: [{ label: 'Cancel', value: null }, { label: 'Insert', primary: true, value: () => chosen }] });
    D.restoreSel(sv);
    if (!v) return;
    recentSyms = [v, ...recentSyms.filter(x => x !== v)].slice(0, 24); try { localStorage.setItem('docs-recent-syms', JSON.stringify(recentSyms)); } catch { }
    D.insertText(v);
  };
  D.insertText = (t) => D.change(() => { D.ensureSel(); const live = D.trackBefore ? D.trackBefore() : null; document.execCommand('insertText', false, t); if (live) D.trackAfter(live); });

  /* ================= equations ================= */
  /** Tiny linear-format → HTML: x^2, x_i, a/b (simple), sqrt(x), \alpha, \pi, \sum ... */
  const TEX = { alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', theta: 'θ', lambda: 'λ', mu: 'μ', pi: 'π', sigma: 'σ', phi: 'φ', omega: 'ω', Delta: 'Δ', Sigma: 'Σ', Omega: 'Ω', sum: '∑', prod: '∏', int: '∫', inf: '∞', infty: '∞', pm: '±', times: '×', div: '÷', le: '≤', ge: '≥', ne: '≠', approx: '≈', to: '→', deg: '°', cdot: '·', partial: '∂', nabla: '∇', in: '∈', forall: '∀', exists: '∃' };
  D.eqToHTML = (s) => {
    let t = esc(s);
    t = t.replace(/\\([a-zA-Z]+)/g, (m, w) => TEX[w] || m);
    t = t.replace(/sqrt\(([^()]*)\)/g, '√<span style="text-decoration:overline">$1</span>');
    t = t.replace(/\(([^()]+)\)\/\(([^()]+)\)/g, '<span class="frac"><span>$1</span><span>$2</span></span>');
    t = t.replace(/(\w+)\/(\w+)/g, '<span class="frac"><span>$1</span><span>$2</span></span>');
    t = t.replace(/\^\{([^}]*)\}/g, '<sup>$1</sup>').replace(/\^(-?[\w.]+)/g, '<sup>$1</sup>');
    t = t.replace(/_\{([^}]*)\}/g, '<sub>$1</sub>').replace(/_([\w.]+)/g, '<sub>$1</sub>');
    return t;
  };
  D.equationDialog = async (existing) => {
    const sv = D.saveSel();
    const tpls = [['Quadratic', 'x = (-b ± sqrt(b^2 - 4ac))/(2a)'], ['Pythagoras', 'a^2 + b^2 = c^2'], ['Area of circle', 'A = \\pi r^2'], ['Binomial', '(a + b)^2 = a^2 + 2ab + b^2'], ['Sum', '\\sum_{i=1}^{n} i = (n(n+1))/(2)'], ['Einstein', 'E = mc^2'], ['Water', 'H_2O'], ['Fraction', '(a)/(b)']];
    const body = document.createElement('div');
    body.innerHTML = `<div class="eq-tpls">${tpls.map(([n, e]) => `<button data-e="${esc(e)}" title="${esc(e)}">${n}</button>`).join('')}</div>
      <div class="field"><label>Equation (linear format)</label><input id="eq-in" value="${esc(existing ? existing.dataset.eq : 'a^2 + b^2 = c^2')}"></div>
      <div class="eq-prev" id="eq-prev"></div>
      <div class="hint">Use ^ for superscript, _ for subscript, (a)/(b) for fractions, sqrt(x) for roots and \\alpha, \\pi, \\sum, \\int, \\pm, \\le for symbols.</div>`;
    const inp = $('#eq-in', body), prev = $('#eq-prev', body); const upd = () => prev.innerHTML = D.eqToHTML(inp.value); upd();
    inp.addEventListener('input', upd);
    body.addEventListener('click', e => { const b = e.target.closest('[data-e]'); if (b) { inp.value = b.dataset.e; upd(); inp.focus(); } });
    const v = await BTS.dialog({ title: existing ? 'Edit equation' : 'Insert equation', body, width: 540, buttons: [{ label: 'Cancel', value: null }, { label: existing ? 'Update' : 'Insert', primary: true, value: () => inp.value }] });
    D.restoreSel(sv);
    if (!v) return;
    D.change(() => {
      if (existing) { existing.dataset.eq = v; existing.innerHTML = D.eqToHTML(v); return; }
      const sp = document.createElement('span'); sp.className = 'eq'; sp.contentEditable = 'false'; sp.dataset.eq = v; sp.innerHTML = D.eqToHTML(v);
      const r = D.ensureSel(); r.deleteContents(); r.insertNode(sp); const sp2 = document.createTextNode(' '); sp.after(sp2); D.setCaret(sp2, 1);
    });
  };

  /* ================= date & time ================= */
  D.dateTimeDialog = async () => {
    const sv = D.saveSel(); const d = new Date();
    const fm = [
      d.toLocaleDateString('en-IN'), d.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' }).replace(/\//g, '-'),
      d.toLocaleDateString('en-IN', { day: 'numeric', month: 'long', year: 'numeric' }), d.toLocaleDateString('en-IN', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }),
      d.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' }), d.toISOString().slice(0, 10), d.toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: '2-digit' }),
      d.toLocaleDateString('en-IN', { month: 'long', year: 'numeric' }), d.toLocaleTimeString('en-IN', { hour: 'numeric', minute: '2-digit' }), d.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit', hour12: false }),
      d.toLocaleString('en-IN', { day: 'numeric', month: 'short', year: 'numeric', hour: 'numeric', minute: '2-digit' }), d.toLocaleDateString('hi-IN', { day: 'numeric', month: 'long', year: 'numeric' }),
    ];
    const v = await BTS.dialog({ title: 'Date and time', width: 400, body: `<div class="dt-list">${fm.map((f, i) => `<label><input type="radio" name="dtf" value="${i}" ${i === 2 ? 'checked' : ''}> ${esc(f)}</label>`).join('')}</div>`, buttons: [{ label: 'Cancel', value: null }, { label: 'Insert', primary: true, value: (r) => fm[+$('input[name=dtf]:checked', r).value] }] });
    D.restoreSel(sv); if (v) D.insertText(v);
  };

  /* ================= misc inserts ================= */
  D.insertHR = () => D.change(() => { D.insertBlocks([document.createElement('hr')]); });
  D.insertTextBox = () => D.change(() => {
    const box = document.createElement('div'); box.className = 'textbox'; box.innerHTML = '<p><b>Text box</b></p><p>Type your note here. Text boxes float on the right of the page.</p>';
    D.insertBlocks([box], box.querySelector('p'));
    D.selectNode(box.querySelector('p'), true);
  });
  D.dropCap = () => D.change(() => {
    D.ensureSel(); const b = D.blocks()[0]; if (!b) return;
    const existing = b.querySelector('.dropcap'); if (existing) { D.unwrap(existing); return; }
    const w = document.createTreeWalker(b, NodeFilter.SHOW_TEXT); let n; while ((n = w.nextNode()) && !n.data.trim());
    if (!n) return; const i = n.data.search(/\S/); const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + 1);
    const sp = document.createElement('span'); sp.className = 'dropcap'; r.surroundContents(sp);
  });
  D.caption = async () => {
    const sv = D.saveSel();
    const nFig = $$('.caption[data-cap="Figure"]', D.doc).length + 1, nTab = $$('.caption[data-cap="Table"]', D.doc).length + 1;
    const v = await BTS.dialog({ title: 'Caption', width: 420, body: `<div class="field"><label>Label</label><select id="cp-l"><option>Figure</option><option>Table</option><option>Equation</option></select></div><div class="field"><label>Caption text</label><input id="cp-t" placeholder="e.g. Growth of literacy in Maharashtra"></div>`, buttons: [{ label: 'Cancel', value: null }, { label: 'Insert', primary: true, value: (r) => ({ l: $('#cp-l', r).value, t: $('#cp-t', r).value }) }] });
    D.restoreSel(sv); if (!v) return;
    const n = $$(`.caption[data-cap="${v.l}"]`, D.doc).length + 1;
    D.change(() => { const p = document.createElement('p'); p.className = 'caption'; p.dataset.cap = v.l; p.textContent = `${v.l} ${n}${v.t ? ': ' + v.t : ''}`; D.insertBlocks([p]); });
  };

  /* ================= header / footer / page numbers ================= */
  D.editHF = (which) => { const el = which === 'header' ? D.hdr : D.ftr; if (!el.firstChild) el.innerHTML = '<p><br></p>'; el.focus(); const r = document.createRange(); r.selectNodeContents(el); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); el.scrollIntoView({ block: 'nearest' }); };
  D.hfPreset = (which, kind) => D.change(() => {
    const el = which === 'header' ? D.hdr : D.ftr; const title = esc(BTS.title || 'Document');
    const map = {
      blank: '<p><br></p>', title: `<p style="text-align:right"><span style="color:#7a8499">${title}</span></p>`,
      line: `<p style="border-bottom:1px solid #9aa1ad;padding-bottom:2pt">${title}</p>`,
      student: '<p>Name: ____________________ &nbsp;&nbsp; Roll no.: ________ &nbsp;&nbsp; Class: ________</p>',
      centered: `<p style="text-align:center">${title}</p>`,
      date: `<p style="text-align:right">${esc(new Date().toLocaleDateString('en-IN', { day: 'numeric', month: 'long', year: 'numeric' }))}</p>`,
      remove: '',
    };
    el.innerHTML = map[kind] ?? '';
    D.applySettings();
    if (kind !== 'remove') D.editHF(which);
  });
  D.insertPageNumber = (where, fmt = 'n') => D.change(() => {
    const el = where.startsWith('top') ? D.hdr : D.ftr;
    const align = where.endsWith('left') ? 'left' : where.endsWith('right') ? 'right' : 'center';
    const field = fmt === 'nofm' ? 'Page <span class="fld" data-fld="page" contenteditable="false">1</span> of <span class="fld" data-fld="pages" contenteditable="false">1</span>' : fmt === 'page' ? 'Page <span class="fld" data-fld="page" contenteditable="false">1</span>' : '<span class="fld" data-fld="page" contenteditable="false">1</span>';
    $$('.fld[data-fld="page"]', el).forEach(f => { const p = f.closest('p'); if (p && p.textContent.replace(/Page|of|\d|\s/g, '') === '') p.remove(); });
    el.insertAdjacentHTML('beforeend', `<p style="text-align:${align}">${field}</p>`);
    $$('p', el).forEach(p => { if (!p.textContent.trim() && !p.querySelector('img,.fld') && el.children.length > 1) p.remove(); });
    D.applySettings();
  });
  D.removePageNumbers = () => D.change(() => { [D.hdr, D.ftr].forEach(el => $$('.fld[data-fld="page"]', el).forEach(f => { const p = f.closest('p'); if (p) p.remove(); else f.remove(); })); D.applySettings(); });
  /** Render header/footer with fields filled for page n of N (for print). */
  D.hfHTML = (el, n, N) => { const c = el.cloneNode(true); $$('.fld[data-fld="page"]', c).forEach(f => f.textContent = n); $$('.fld[data-fld="pages"]', c).forEach(f => f.textContent = N); return c.innerHTML; };

  /* ================= table of contents ================= */
  D.tocHTML = (opts = {}) => {
    const levels = opts.levels || 3; const heads = $$('h1,h2,h3', D.doc).filter(h => +h.tagName[1] <= levels && h.textContent.trim() && !h.closest('.toc'));
    D.layout();
    const items = heads.map(h => { const id = ensureId(h); const pg = D.headPage && D.headPage.get(h) || pageOfEl(h); return `<p class="toc-l${h.tagName[1]}" data-target="${id}"><span class="toc-t">${esc(h.textContent.trim())}</span><span class="toc-dots"></span><span class="toc-n">${pg}</span></p>`; }).join('');
    return `<div class="toc-title">${esc(opts.title || 'Contents')}</div>` + (items || '<p class="toc-empty">No headings found. Apply Heading 1–3 styles to your titles, then update the table.</p>');
  };
  function pageOfEl(el) {
    const dr = D.doc.getBoundingClientRect(); const scale = dr.height / (D.doc.offsetHeight || 1) || 1; const y = (el.getBoundingClientRect().top - dr.top) / scale + 1;
    let p = 1; D.pages.forEach((b, i) => { if (y >= b) p = i + 1; }); return p;
  }
  D.pageOfEl = pageOfEl;
  D.insertTOC = (style = 'auto') => D.change(() => {
    let toc = $('.toc:not(.tof)', D.doc);
    if (!toc) { toc = document.createElement('div'); toc.className = 'toc'; toc.contentEditable = 'false'; toc.dataset.style = style; D.insertBlocks([toc]); }
    toc.dataset.style = style; toc.innerHTML = D.tocHTML({ title: style === 'toc-index' ? 'Index of topics' : 'Contents' });
    // recompute numbers after the TOC itself took space
    D.layout(); toc.innerHTML = D.tocHTML({ title: style === 'toc-index' ? 'Index of topics' : 'Contents' });
  });
  D.updateTOC = (silent) => {
    const tocs = $$('.toc:not(.tof)', D.doc); if (!tocs.length) { if (!silent) BTS.toast('There is no table of contents yet — insert one first.'); return; }
    D.change(() => { tocs.forEach(t => { t.innerHTML = D.tocHTML({ title: t.dataset.style === 'toc-index' ? 'Index of topics' : 'Contents' }); }); });
    if (!silent) BTS.toast('Table of contents updated');
  };
  D.removeTOC = () => D.change(() => $$('.toc:not(.tof)', D.doc).forEach(t => t.remove()));

  /* ================= footnotes ================= */
  D.footnotes = []; // [{id, html}]
  D.insertFootnote = async () => {
    const sv = D.saveSel();
    const v = await BTS.prompt('Insert footnote', 'Footnote text', ''); D.restoreSel(sv);
    if (v == null || !v.trim()) return;
    D.change(() => {
      const id = D.uid('fn'); const sup = document.createElement('sup'); sup.className = 'fn-ref'; sup.dataset.fn = id; sup.contentEditable = 'false'; sup.dataset.text = v.trim(); sup.textContent = '0';
      const r = D.ensureSel(); r.collapse(false); r.insertNode(sup); D.setCaret(sup.parentNode, [...sup.parentNode.childNodes].indexOf(sup) + 1);
      D.renderFootnotes();
    });
  };
  D.renderFootnotes = () => {
    const refs = $$('sup.fn-ref', D.doc); const area = $('#fn-area');
    refs.forEach((s, i) => s.textContent = i + 1);
    area.innerHTML = refs.length ? `<div class="footnotes"><div class="fn-title">Notes</div><ol>${refs.map((s, i) => `<li data-fn="${s.dataset.fn}">${esc(s.dataset.text || '')}</li>`).join('')}</ol></div>` : '';
  };
  D.editFootnote = async (sup) => {
    const v = await BTS.dialog({ title: 'Footnote ' + sup.textContent, width: 440, body: `<div class="field"><label>Footnote text</label><textarea id="fn-t">${esc(sup.dataset.text || '')}</textarea></div>`, buttons: [{ label: 'Delete footnote', value: 'del', danger: true }, { label: 'Cancel', value: null }, { label: 'Save', primary: true, value: (r) => $('#fn-t', r).value }] });
    if (v === 'del') D.change(() => { sup.remove(); D.renderFootnotes(); });
    else if (v != null) D.change(() => { sup.dataset.text = v; D.renderFootnotes(); });
  };
  D.nextFootnote = () => { const refs = $$('sup.fn-ref', D.doc); if (!refs.length) { BTS.toast('No footnotes yet'); return; } const r = D.range(); let i = 0; if (r) { i = refs.findIndex(s => r.compareBoundaryPoints(Range.START_TO_START, rangeOf(s)) < 0); if (i < 0) i = 0; } D.scrollToEl(refs[i]); D.setCaret(refs[i].parentNode, [...refs[i].parentNode.childNodes].indexOf(refs[i]) + 1); };
  function rangeOf(n) { const r = document.createRange(); r.selectNode(n); return r; }

  /* ================= comments ================= */
  D.comments = {};
  D.addComment = async () => {
    D.ensureSel(); let r = D.range();
    if (!r || !D.doc.contains(r.commonAncestorContainer)) { BTS.toast('Select some text in the document to comment on.'); return; }
    if (r.collapsed) { D.selectWordIfCollapsed(); r = D.range(); }
    if (!r || r.collapsed) { BTS.toast('Select some text to comment on.'); return; }
    const id = D.uid('c');
    D.change(() => {
      const quote = r.toString().slice(0, 80);
      D.wrapText(r, () => { const s = document.createElement('span'); s.className = 'cmt'; s.dataset.cid = id; return s; });
      D.comments[id] = { text: '', author: D.author(), time: Date.now(), quote };
      getSelection().collapseToEnd();
    });
    D.showComments(true); D.renderComments();
    setTimeout(() => { const ta = $(`.ccard[data-cid="${id}"] textarea`); if (ta) ta.focus(); }, 30);
  };
  D.author = () => { try { return localStorage.getItem('docs-author') || 'Student'; } catch { return 'Student'; } };
  D.deleteComment = (id) => {
    if (!id) { const a = $('.ccard.active'); id = a && a.dataset.cid; }
    if (!id) { const r = D.range(); const c = r && D.closest(r.startContainer, '.cmt'); id = c && c.dataset.cid; }
    if (!id) { BTS.toast('Click a comment first.'); return; }
    D.change(() => { $$(`.cmt[data-cid="${id}"]`, D.doc).forEach(D.unwrap); delete D.comments[id]; });
    D.renderComments();
  };
  D.deleteAllComments = () => { if (!$('.cmt', D.doc)) { BTS.toast('There are no comments.'); return; } D.change(() => { $$('.cmt', D.doc).forEach(D.unwrap); D.comments = {}; }); D.renderComments(); };
  D.showComments = (on) => {
    const rail = $('#crail'); const show = on ?? rail.classList.contains('hidden');
    rail.classList.toggle('hidden', !show); document.body.classList.toggle('hide-comments', !show);
    $('#rt-comments').classList.toggle('on', show); $('#btn-show-comments').classList.toggle('on', show);
    D.renderComments();
  };
  D.renderComments = () => {
    const rail = $('#crail'); const ids = [...new Set($$('.cmt', D.doc).map(s => s.dataset.cid))];
    Object.keys(D.comments).forEach(k => { if (!ids.includes(k)) delete D.comments[k]; });
    ids.forEach(id => { if (!D.comments[id]) D.comments[id] = { text: '', author: D.author(), time: Date.now(), quote: '' }; });
    if (ids.length && rail.classList.contains('hidden') && !D._commentsHiddenByUser) { rail.classList.remove('hidden'); document.body.classList.remove('hide-comments'); $('#rt-comments').classList.add('on'); $('#btn-show-comments').classList.add('on'); }
    if (rail.classList.contains('hidden')) return;
    rail.innerHTML = ids.length ? ids.map(id => { const c = D.comments[id]; return `<div class="ccard" data-cid="${id}"><div class="c-head"><span class="c-av">${esc((c.author || 'S')[0].toUpperCase())}</span><span><div class="c-who">${esc(c.author || 'Student')}</div><div class="c-when">${esc(new Date(c.time).toLocaleString('en-IN', { day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit' }))}</div></span><button class="rb c-del" data-cmd="comment-delete-id" data-cid="${id}" title="Delete comment">${icon('trash-2', 14)}</button></div>${c.quote ? `<div class="c-quote">“${esc(c.quote)}”</div>` : ''}<textarea placeholder="Add a comment…" rows="1">${esc(c.text || '')}</textarea></div>`; }).join('') : '<div class="empty-note" style="padding:10px 4px;font-size:12px">No comments. Select text and choose New comment.</div>';
    D.positionComments();
  };
  D.positionComments = () => {
    const rail = $('#crail'); if (rail.classList.contains('hidden')) return;
    const rr = rail.getBoundingClientRect(); const s = D.zoom / 100; let lastBottom = 0;
    $$('.ccard', rail).forEach(card => {
      const span = $(`.cmt[data-cid="${card.dataset.cid}"]`, D.doc); if (!span) return;
      let top = (span.getBoundingClientRect().top - rr.top) / s - 8; top = Math.max(top, lastBottom + 8);
      card.style.top = top + 'px'; lastBottom = top + card.offsetHeight;
    });
  };
  D.activateComment = (id) => {
    $$('.ccard').forEach(c => c.classList.toggle('active', c.dataset.cid === id));
    $$('.cmt', D.doc).forEach(c => c.classList.toggle('active', c.dataset.cid === id));
  };
  D.stepComment = (dir) => {
    const spans = $$('.cmt', D.doc).filter((s, i, a) => a.findIndex(x => x.dataset.cid === s.dataset.cid) === i); if (!spans.length) { BTS.toast('There are no comments.'); return; }
    const cur = $('.ccard.active'); let i = cur ? spans.findIndex(s => s.dataset.cid === cur.dataset.cid) : -1;
    i = (i + dir + spans.length) % spans.length; D.showComments(true); D.scrollToEl(spans[i]); D.activateComment(spans[i].dataset.cid);
  };
  function wireComments() {
    const rail = $('#crail');
    rail.addEventListener('input', e => { const ta = e.target.closest('textarea'); if (!ta) return; const id = ta.closest('.ccard').dataset.cid; if (D.comments[id]) { D.comments[id].text = ta.value; BTS.markDirty(); } D.positionComments(); });
    rail.addEventListener('focusin', e => { const c = e.target.closest('.ccard'); if (c) D.activateComment(c.dataset.cid); });
    rail.addEventListener('mousedown', e => { const c = e.target.closest('.ccard'); if (c) D.activateComment(c.dataset.cid); });
    rail.addEventListener('change', e => { if (e.target.closest('textarea')) D.checkpoint(); });
  }

  D.initInsert = () => { wireImages(); wireComments(); };
})();
