/* BTS Slides — document model, factories, assets and undo/redo.
 *
 * presentation = { width, height (logical px; 96 px = 1 in; 1280×720 = 16:9, 960×720 = 4:3),
 *   theme:{id, variant}, footer:{…}, slides:[slide] }
 * slide = { id, layout, background|null, hideGraphics, transition:{type,dur,dir,auto}, notes, hidden, comments:[],
 *   layer:[el] (imported master/layout graphics, not editable), elements:[el] }   (array order = z-order)
 * el = { id, type:'text'|'shape'|'image'|'table'|'chart'|'icon', x,y,w,h, rot, opacity, flipH, flipV, g (group id),
 *   html, role, font, size(pt), color, align, valign, lh, pspace, bold, inset, dir, effect,
 *   kind, fill, stroke, strokeWidth, dash, shadow, radius, src (asset id), crop, rows, chart, name, anim:{type,trigger,dur,delay,dir,order},
 *   ph, placeholder, autofit }
 * Images are stored once in SL.assets (id → data URL) so undo snapshots stay small.
 */
(function () {
  const SL = window.SL = window.SL || {};
  let seq = 0;
  SL.uid = (p = 'e') => p + Date.now().toString(36).slice(-5) + (seq++).toString(36) + Math.random().toString(36).slice(2, 5);

  /* ---------- assets ---------- */
  SL.assets = {};
  SL.addAsset = (dataURL) => {
    for (const [k, v] of Object.entries(SL.assets)) if (v === dataURL) return k;
    const id = 'a' + SL.uid(''); SL.assets[id] = dataURL; return id;
  };
  SL.assetURL = (id) => (id && (SL.assets[id] || (id.startsWith('data:') ? id : ''))) || '';

  /* ---------- factories ---------- */
  SL.W = (p) => p.width || 1280;
  SL.H = (p) => p.height || 720;
  SL.newPresentation = (themeId = 'clean') => {
    const p = { width: 1280, height: 720, theme: { id: themeId, variant: 0 }, footer: { date: false, dateText: '', number: false, text: '', skipTitle: true }, slides: [] };
    p.slides.push(SL.newSlide(p, 'title'));
    return p;
  };
  SL.newSlide = (pres, layout = 'titleContent') => {
    const s = { id: SL.uid('s'), layout, background: null, hideGraphics: false, transition: { type: 'none', dur: 700 }, notes: '', hidden: false, comments: [], layer: [], elements: [] };
    SL.layoutPlaceholders(layout, SL.W(pres), SL.H(pres), pres).forEach(spec => s.elements.push(SL.newElement(spec.type, spec)));
    return s;
  };
  const DEF = {
    common: { rot: 0, opacity: 1 },
    text: { html: '', role: 'body', size: 18, align: 'left', valign: 'top', fill: 'none', stroke: 'none', strokeWidth: 0 },
    shape: { kind: 'rect', html: '', role: 'body', size: 18, align: 'center', valign: 'middle', fill: 'accent1', stroke: 'none', strokeWidth: 2, color: '#ffffff' },
    image: { src: null, stroke: 'none', strokeWidth: 0, radius: 0 },
    table: { rows: [], header: true, banded: true, accent: 'accent1', size: 16 },
    chart: { chart: null },
    icon: { name: 'star', color: 'accent1', strokeWidth: 1.75 },
  };
  SL.newElement = (type, props = {}) => Object.assign({ id: SL.uid(), type }, DEF.common, DEF[type] || {}, props);
  SL.newTable = (rows, cols, props = {}) => SL.newElement('table', Object.assign({
    rows: Array.from({ length: rows }, () => Array.from({ length: cols }, () => ({ html: '' }))),
    colW: Array(cols).fill(1 / cols),
  }, props));

  SL.isEmptyText = (el) => !el.html || !SL.plainText(el.html).trim();
  SL.plainText = (html) => { const d = document.createElement('div'); d.innerHTML = String(html || '').replace(/<\/(p|li|div)>/gi, '</$1>\n').replace(/<br\s*\/?>/gi, '\n'); return d.textContent.replace(/\n+$/, ''); };
  SL.textToHTML = (t) => String(t || '').split('\n').map(l => '<p>' + (BTS.esc(l) || '<br>') + '</p>').join('');

  /* ---------- rich text sanitizer (keeps the subset we can render & export) ---------- */
  const KEEP_STYLE = ['color', 'background-color', 'font-size', 'font-family', 'font-weight', 'font-style', 'text-decoration', 'text-decoration-line', 'text-align', 'vertical-align'];
  const INLINE = { B: 'b', STRONG: 'b', I: 'i', EM: 'i', U: 'u', S: 's', STRIKE: 's', DEL: 's', SPAN: 'span', FONT: 'span', SUB: 'sub', SUP: 'sup', A: 'a' };
  SL.sanitize = (html) => {
    const src = document.createElement('div'); src.innerHTML = html || '';
    const out = document.createElement('div');
    let curP = null;
    const ensureP = () => { if (!curP) { curP = document.createElement('p'); out.appendChild(curP); } return curP; };
    const cleanStyle = (node, target) => {
      if (node.style) for (const k of KEEP_STYLE) { const v = node.style.getPropertyValue(k); if (v && k !== 'text-align') target.style.setProperty(k, v); }
      if (node.tagName === 'FONT') { if (node.color) target.style.color = node.color; if (node.face) target.style.fontFamily = node.face; }
    };
    const inline = (node, parent) => {
      for (const ch of [...node.childNodes]) {
        if (ch.nodeType === 3) { parent.appendChild(document.createTextNode(ch.nodeValue)); continue; }
        if (ch.nodeType !== 1) continue;
        const t = ch.tagName;
        if (t === 'BR') { parent.appendChild(document.createElement('br')); continue; }
        if (INLINE[t]) {
          const n = document.createElement(INLINE[t]); cleanStyle(ch, n);
          if (t === 'A' && ch.getAttribute('href')) n.setAttribute('href', ch.getAttribute('href'));
          if (ch.getAttribute('data-field')) n.setAttribute('data-field', ch.getAttribute('data-field'));
          if (n.tagName === 'SPAN' && !n.getAttribute('style') && !n.getAttribute('data-field')) { inline(ch, parent); continue; }
          inline(ch, n); parent.appendChild(n); continue;
        }
        inline(ch, parent); // unknown inline-ish: unwrap
      }
    };
    const blockAttrs = (srcEl, dst) => {
      const al = srcEl.style && srcEl.style.textAlign; if (al) dst.style.textAlign = al;
      const lv = srcEl.getAttribute && srcEl.getAttribute('data-lvl'); if (lv && +lv > 0) dst.setAttribute('data-lvl', lv);
    };
    const walk = (node, lvl) => {
      for (const ch of [...node.childNodes]) {
        if (ch.nodeType === 3) { if (ch.nodeValue.trim() || curP) ensureP().appendChild(document.createTextNode(ch.nodeValue)); continue; }
        if (ch.nodeType !== 1) continue;
        const t = ch.tagName;
        if (t === 'UL' || t === 'OL') {
          curP = null;
          const list = document.createElement(t.toLowerCase());
          for (const li of [...ch.childNodes]) {
            if (li.nodeType === 1 && (li.tagName === 'UL' || li.tagName === 'OL')) { // nested list from execCommand indent
              const tmp = document.createElement('div'); tmp.appendChild(li.cloneNode(true));
              const sub = SL.sanitize(tmp.innerHTML); const d = document.createElement('div'); d.innerHTML = sub;
              d.querySelectorAll('li').forEach(x => { x.setAttribute('data-lvl', String(lvl + 1 + (+x.getAttribute('data-lvl') || 0))); list.appendChild(x); });
              continue;
            }
            if (li.nodeType !== 1 || li.tagName !== 'LI') { if (li.nodeType === 3 && !li.nodeValue.trim()) continue; }
            const n = document.createElement('li'); if (li.nodeType === 1) { blockAttrs(li, n); if (lvl && !n.getAttribute('data-lvl')) n.setAttribute('data-lvl', lvl); }
            if (li.nodeType === 1) {
              // flatten nested lists inside li
              const nested = [...li.childNodes].filter(x => x.nodeType === 1 && (x.tagName === 'UL' || x.tagName === 'OL'));
              nested.forEach(x => x.remove());
              const blocks = [...li.childNodes].filter(x => x.nodeType === 1 && /^(P|DIV)$/.test(x.tagName));
              if (blocks.length) blocks.forEach((b, i) => { if (i) n.appendChild(document.createElement('br')); inline(b, n); }); else inline(li, n);
              if (!n.childNodes.length) n.appendChild(document.createElement('br'));
              list.appendChild(n);
              nested.forEach(x => { const tmp = document.createElement('div'); tmp.appendChild(x); const d = document.createElement('div'); d.innerHTML = SL.sanitize(tmp.innerHTML); d.querySelectorAll('li').forEach(y => { y.setAttribute('data-lvl', String((+n.getAttribute('data-lvl') || 0) + 1 + (+y.getAttribute('data-lvl') || 0))); list.appendChild(y); }); });
            } else { n.appendChild(document.createTextNode(li.nodeValue)); list.appendChild(n); }
          }
          if (list.childNodes.length) out.appendChild(list);
          continue;
        }
        if (/^(P|DIV|H[1-6]|BLOCKQUOTE|PRE|LI|SECTION|ARTICLE|HEADER|FOOTER|TABLE|TBODY|TR|TD|TH)$/.test(t)) {
          if (/^(DIV|BLOCKQUOTE|SECTION|ARTICLE|TABLE|TBODY|TR)$/.test(t) && [...ch.children].some(c => /^(P|DIV|UL|OL|H[1-6]|TABLE|TR|TD|TH|LI)$/.test(c.tagName))) { curP = null; walk(ch, lvl + (t === 'BLOCKQUOTE' ? 1 : 0)); curP = null; continue; }
          const p = document.createElement('p'); blockAttrs(ch, p); if (t === 'BLOCKQUOTE' && !p.getAttribute('data-lvl')) p.setAttribute('data-lvl', String(lvl + 1));
          if (/^H[1-3]$/.test(t)) { const b = document.createElement('b'); inline(ch, b); p.appendChild(b); } else inline(ch, p);
          if (!p.childNodes.length) p.appendChild(document.createElement('br'));
          out.appendChild(p); curP = null; continue;
        }
        if (t === 'BR') { if (curP) { curP.appendChild(document.createElement('br')); } else { const p = document.createElement('p'); p.appendChild(document.createElement('br')); out.appendChild(p); } continue; }
        if (/^(STYLE|SCRIPT|META|TITLE|HEAD|IMG|SVG|IFRAME|OBJECT)$/.test(t)) continue;
        const holder = ensureP(); const tmp = document.createElement('span'); tmp.appendChild(ch.cloneNode(true)); inline(tmp, holder);
      }
    };
    walk(src, 0);
    // merge trailing <br> inside paragraphs that end with br followed by content? keep as-is
    out.querySelectorAll('p,li').forEach(b => { if (b.childNodes.length > 1 && b.lastChild.nodeName === 'BR' && b.textContent.length) b.lastChild.remove(); });
    return out.innerHTML;
  };

  /* ---------- undo / redo (JSON snapshots) ---------- */
  const H = SL.history = { stack: [], idx: -1, max: 150 };
  SL.snapshot = () => JSON.stringify({ p: SL.pres, cur: SL.cur });
  SL.resetHistory = () => { H.stack = [SL.snapshot()]; H.idx = 0; };
  SL.commit = (opts = {}) => {
    const s = SL.snapshot();
    if (H.stack[H.idx] === s) return false;
    H.stack = H.stack.slice(0, H.idx + 1); H.stack.push(s);
    if (H.stack.length > H.max) H.stack.shift();
    H.idx = H.stack.length - 1;
    if (!opts.silent) BTS.markDirty();
    SL.onChange && SL.onChange();
    return true;
  };
  const restore = (s) => { const o = JSON.parse(s); SL.pres = o.p; SL.cur = Math.min(o.cur, SL.pres.slides.length - 1); };
  SL.undo = () => { if (H.idx <= 0) return false; H.idx--; restore(H.stack[H.idx]); BTS.markDirty(); return true; };
  SL.redo = () => { if (H.idx >= H.stack.length - 1) return false; H.idx++; restore(H.stack[H.idx]); BTS.markDirty(); return true; };

  SL.slide = () => SL.pres.slides[SL.cur];
  SL.findEl = (id, slide = SL.slide()) => slide && slide.elements.find(e => e.id === id);
  SL.clone = (o) => JSON.parse(JSON.stringify(o));
  SL.reId = (el) => { el.id = SL.uid(); return el; };

  SL.pres = SL.newPresentation();
  SL.cur = 0;
})();
