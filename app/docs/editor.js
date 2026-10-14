/* BTS Docs — core editor: document state, history (undo/redo), selection helpers,
 * inline & paragraph formatting, styles, pagination/page guides, ruler, zoom, status.
 * Other files extend window.Docs (D). */
(function () {
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const esc = (s) => BTS.esc(s);
  const PX = 96; // CSS px per inch

  const D = window.Docs = {
    $, $$, PX,
    doc: null, hdr: null, ftr: null, sheet: null, canvas: null,
    page: null, design: null, comments: {},
    zoom: 100, pages: [0], pageCount: 1,
  };

  /* ---------------- constants ---------------- */
  D.PAPER = {
    A4: { w: 8.27, h: 11.69, label: 'A4', note: '21 cm × 29.7 cm' },
    Letter: { w: 8.5, h: 11, label: 'Letter', note: '21.59 cm × 27.94 cm' },
    Legal: { w: 8.5, h: 14, label: 'Legal', note: '21.59 cm × 35.56 cm' },
    A5: { w: 5.83, h: 8.27, label: 'A5', note: '14.8 cm × 21 cm' },
  };
  D.MARGINS = {
    normal: { label: 'Normal', t: 1, b: 1, l: 1, r: 1 },
    narrow: { label: 'Narrow', t: 0.5, b: 0.5, l: 0.5, r: 0.5 },
    moderate: { label: 'Moderate', t: 1, b: 1, l: 0.75, r: 0.75 },
    wide: { label: 'Wide', t: 1, b: 1, l: 2, r: 2 },
    college: { label: 'College (binding)', t: 1, b: 1, l: 1.5, r: 1 },
  };
  D.THEMES = [
    { id: 'bts', name: 'BTS', accent: '#3C4FE0', dark: '#1f2a6b', head: 'Calibri Light', body: 'Calibri', colors: ['#3C4FE0', '#1f2a6b', '#0F8A6A', '#D9661A'] },
    { id: 'office', name: 'Classic Blue', accent: '#2F5496', dark: '#1F3763', head: 'Calibri Light', body: 'Calibri', colors: ['#2F5496', '#1F3763', '#C55A11', '#548235'] },
    { id: 'ocean', name: 'Ocean', accent: '#0E7C86', dark: '#0A4F55', head: 'Georgia', body: 'Calibri', colors: ['#0E7C86', '#0A4F55', '#5BB5BD', '#F2A541'] },
    { id: 'saffron', name: 'Saffron', accent: '#D9661A', dark: '#7A3410', head: 'Cambria', body: 'Calibri', colors: ['#D9661A', '#7A3410', '#F2B705', '#2E7D4F'] },
    { id: 'forest', name: 'Forest', accent: '#2E7D4F', dark: '#1B4D30', head: 'Georgia', body: 'Segoe UI', colors: ['#2E7D4F', '#1B4D30', '#8DB580', '#C9A227'] },
    { id: 'royal', name: 'Royal', accent: '#6B3FA0', dark: '#3E2163', head: 'Garamond', body: 'Garamond', colors: ['#6B3FA0', '#3E2163', '#B48EDB', '#E0A526'] },
    { id: 'crimson', name: 'Crimson', accent: '#B3263E', dark: '#6B1626', head: 'Trebuchet MS', body: 'Trebuchet MS', colors: ['#B3263E', '#6B1626', '#E3818F', '#3D5A80'] },
    { id: 'formal', name: 'Formal', accent: '#000000', dark: '#000000', head: 'Times New Roman', body: 'Times New Roman', colors: ['#000000', '#404040', '#7F7F7F', '#BFBFBF'] },
    { id: 'slate', name: 'Slate', accent: '#44546A', dark: '#222A35', head: 'Segoe UI', body: 'Segoe UI', colors: ['#44546A', '#222A35', '#8497B0', '#ED7D31'] },
  ];
  D.FONTS = ['Calibri', 'Calibri Light', 'Arial', 'Times New Roman', 'Cambria', 'Georgia', 'Segoe UI', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Garamond', 'Courier New', 'Comic Sans MS', 'Nirmala UI', 'Mangal'];
  D.SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  D.STYLES = [
    { id: 'normal', name: 'Normal', tag: 'p', cls: '' },
    { id: 'nospace', name: 'No Spacing', tag: 'p', cls: 'nospace' },
    { id: 'h1', name: 'Heading 1', tag: 'h1', cls: '' },
    { id: 'h2', name: 'Heading 2', tag: 'h2', cls: '' },
    { id: 'h3', name: 'Heading 3', tag: 'h3', cls: '' },
    { id: 'title', name: 'Title', tag: 'p', cls: 'title' },
    { id: 'subtitle', name: 'Subtitle', tag: 'p', cls: 'subtitle' },
    { id: 'quote', name: 'Quote', tag: 'blockquote', cls: '' },
    { id: 'code', name: 'Code', tag: 'pre', cls: '' },
    { id: 'caption', name: 'Caption', tag: 'p', cls: 'caption' },
  ];
  const BLOCK_TAGS = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'PRE', 'BLOCKQUOTE', 'DIV', 'TD', 'TH']);
  D.BLOCK_TAGS = BLOCK_TAGS;
  const TEXT_BLOCKS = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'PRE']);
  D.TEXT_BLOCKS = TEXT_BLOCKS;

  D.defaultPage = () => ({ size: 'A4', orient: 'portrait', m: { t: 1, b: 1, l: 1, r: 1 }, hdrDist: 0.5, ftrDist: 0.5 });
  D.defaultDesign = () => ({ theme: 'bts', pageColor: null, watermark: null, border: null, pspacing: null, hdrFirst: false });
  D.page = D.defaultPage();
  D.design = D.defaultDesign();

  /* ---------------- unit helpers ---------------- */
  D.toPx = (v, base = 11 * 4 / 3) => {
    if (v == null || v === '') return null;
    const m = String(v).trim().match(/^(-?[\d.]+)\s*(px|pt|in|cm|mm|em|rem|%)?$/);
    if (!m) return null;
    const n = parseFloat(m[1]);
    switch (m[2]) { case 'pt': return n * 4 / 3; case 'in': return n * 96; case 'cm': return n * 96 / 2.54; case 'mm': return n * 96 / 25.4; case 'em': case 'rem': return n * base; case '%': return null; default: return n; }
  };
  D.paperSize = () => { const p = D.PAPER[D.page.size] || D.PAPER.A4; return D.page.orient === 'landscape' ? { w: p.h, h: p.w } : { w: p.w, h: p.h }; };
  D.theme = () => D.THEMES.find(t => t.id === D.design.theme) || D.THEMES[0];
  D.fontStack = (f) => {
    const map = { Calibri: 'Calibri, Carlito, "Segoe UI", Arial, sans-serif', 'Calibri Light': '"Calibri Light", Calibri, Carlito, "Segoe UI", Arial, sans-serif', Cambria: 'Cambria, Caladea, Georgia, serif', Garamond: 'Garamond, "EB Garamond", Georgia, serif', 'Segoe UI': '"Segoe UI", "Noto Sans", Arial, sans-serif', 'Times New Roman': '"Times New Roman", "Liberation Serif", Tinos, serif' };
    return map[f] || `"${f}", sans-serif`;
  };
  D.uid = (p = 'x') => p + Math.random().toString(36).slice(2, 9);

  /* ---------------- history (snapshot undo, covers formatting + structure + settings) ---------------- */
  const imgStore = new Map(); const imgList = [];
  function tokenize(html) {
    return html.replace(/<hr class="page-break" style="[^"]*">/g, '<hr class="page-break">').replace(/src="(data:[^"]{200,})"/g, (m, url) => { let i = imgStore.get(url); if (i == null) { i = imgList.length; imgList.push(url); imgStore.set(url, i); } return 'src="bts-img:' + i + '"'; });
  }
  const detokenize = (html) => html.replace(/src="bts-img:(\d+)"/g, (m, i) => 'src="' + imgList[+i] + '"');
  D.detokenize = detokenize;
  const hist = { undo: [], redo: [], lastType: 0, lastKind: '', restoring: false };
  function capture() {
    const html = tokenize(D.doc.innerHTML), hdr = tokenize(D.hdr.innerHTML), ftr = tokenize(D.ftr.innerHTML);
    const settings = JSON.stringify({ p: D.page, d: D.design, c: D.comments });
    return { key: html + '\u0001' + hdr + '\u0001' + ftr + '\u0001' + settings, html, hdr, ftr, settings, sel: saveSel() };
  }
  D.checkpoint = () => {
    if (hist.restoring || !D.doc) return;
    const s = capture(); const top = hist.undo[hist.undo.length - 1];
    if (!top || top.key !== s.key) { hist.undo.push(s); if (hist.undo.length > 200) hist.undo.shift(); hist.redo = []; }
  };
  function restore(s) {
    hist.restoring = true;
    try {
      D.doc.innerHTML = detokenize(s.html); D.hdr.innerHTML = detokenize(s.hdr); D.ftr.innerHTML = detokenize(s.ftr);
      const st = JSON.parse(s.settings); D.page = st.p; D.design = st.d; D.comments = st.c || {};
      D.applySettings(); restoreSel(s.sel);
    } finally { hist.restoring = false; }
    D.afterChange({ noHistory: true });
  }
  D.undo = () => {
    const cur = capture();
    while (hist.undo.length && hist.undo[hist.undo.length - 1].key === cur.key) hist.undo.pop();
    const s = hist.undo.pop(); if (!s) { BTS.toast('Nothing to undo'); return; }
    hist.redo.push(cur); restore(s); hist.lastType = 0;
  };
  D.redo = () => {
    const s = hist.redo.pop(); if (!s) return;
    const cur = capture(); hist.undo.push(cur); restore(s); hist.lastType = 0;
  };
  D.clearHistory = () => { hist.undo = []; hist.redo = []; imgStore.clear(); imgList.length = 0; };
  /** Run a modifying operation with a single undo step. */
  D.change = (fn) => {
    D.checkpoint();
    const r = fn();
    D.afterChange();
    hist.lastType = 0;
    return r;
  };
  function onBeforeInput(e) {
    const t = e.inputType || '';
    if (t === 'historyUndo') { e.preventDefault(); D.undo(); return; }
    if (t === 'historyRedo') { e.preventDefault(); D.redo(); return; }
    const kind = t.startsWith('insertText') || t === 'insertCompositionText' ? 'type' : t.startsWith('delete') ? 'del' : t;
    const now = Date.now();
    if (kind !== hist.lastKind || now - hist.lastType > 900 || (kind === 'type' && e.data === ' ' && now - hist.lastType > 300) || !['type', 'del'].includes(kind)) D.checkpoint();
    hist.lastKind = kind; hist.lastType = now;
  }

  /* ---------------- selection save / restore by node path ---------------- */
  function pathOf(node, root) { const p = []; while (node && node !== root) { const par = node.parentNode; if (!par) return null; p.unshift([...par.childNodes].indexOf(node)); node = par; } return node === root ? p : null; }
  function nodeAt(path, root) { let n = root; for (const i of path) { if (!n.childNodes[i]) return null; n = n.childNodes[i]; } return n; }
  function saveSel() {
    const s = getSelection(); if (!s.rangeCount) return null;
    const r = s.getRangeAt(0); const root = D.rootOf(r.startContainer); if (!root) return null;
    const rootId = root.id || null; if (!rootId) return null;
    return { root: rootId, a: pathOf(r.startContainer, root), ao: r.startOffset, b: pathOf(r.endContainer, root), bo: r.endOffset };
  }
  function restoreSel(sv) {
    if (!sv) return;
    const root = document.getElementById(sv.root); if (!root) return;
    try {
      const a = nodeAt(sv.a, root), b = nodeAt(sv.b, root); if (!a || !b) return;
      const r = document.createRange(); r.setStart(a, Math.min(sv.ao, a.nodeType === 3 ? a.length : a.childNodes.length)); r.setEnd(b, Math.min(sv.bo, b.nodeType === 3 ? b.length : b.childNodes.length));
      root.focus({ preventScroll: true }); const s = getSelection(); s.removeAllRanges(); s.addRange(r);
    } catch { }
  }
  D.saveSel = saveSel; D.restoreSel = restoreSel;

  /* ---------------- selection helpers ---------------- */
  D.rootOf = (node) => { while (node) { if (node === D.doc || node === D.hdr || node === D.ftr) return node; node = node.parentNode; } return null; };
  D.range = () => { const s = getSelection(); if (!s.rangeCount) return null; const r = s.getRangeAt(0); return D.rootOf(r.commonAncestorContainer) ? r : null; };
  let lastRange = null;
  D.lastRange = () => lastRange;
  /** Ensure there is a selection inside an editable root (restores last one or puts caret at doc end). */
  D.ensureSel = () => {
    let r = D.range(); if (r) return r;
    if (lastRange && D.rootOf(lastRange.commonAncestorContainer) && lastRange.startContainer.isConnected) {
      const root = D.rootOf(lastRange.commonAncestorContainer); root.focus({ preventScroll: true });
      const s = getSelection(); s.removeAllRanges(); s.addRange(lastRange); return lastRange;
    }
    D.doc.focus({ preventScroll: true });
    if (!D.doc.firstChild) D.doc.innerHTML = '<p><br></p>';
    r = document.createRange(); const first = D.doc.firstElementChild || D.doc; r.selectNodeContents(first); r.collapse(true);
    const s = getSelection(); s.removeAllRanges(); s.addRange(r); return r;
  };
  D.setCaret = (node, offset = 0) => { const r = document.createRange(); r.setStart(node, offset); r.collapse(true); (D.rootOf(node) || D.doc).focus({ preventScroll: true }); const s = getSelection(); s.removeAllRanges(); s.addRange(r); if (D.rootOf(node)) lastRange = r.cloneRange(); };
  D.selectNode = (node, contents) => { const r = document.createRange(); contents ? r.selectNodeContents(node) : r.selectNode(node); const s = getSelection(); s.removeAllRanges(); s.addRange(r); };
  D.closest = (node, sel) => { const el = node && (node.nodeType === 1 ? node : node.parentElement); const c = el && el.closest(sel); return c && D.rootOf(c) ? c : null; };
  D.blockOf = (node) => {
    let el = node && (node.nodeType === 1 ? node : node.parentElement);
    while (el && !D.isRoot(el)) {
      if (TEXT_BLOCKS.has(el.tagName)) return el;
      if ((el.tagName === 'BLOCKQUOTE' || el.tagName === 'DIV' || el.tagName === 'TD' || el.tagName === 'TH') && !hasBlockChild(el)) return el;
      el = el.parentElement;
    }
    return null;
  };
  D.isRoot = (el) => el === D.doc || el === D.hdr || el === D.ftr;
  function hasBlockChild(el) { for (const c of el.children) if (BLOCK_TAGS.has(c.tagName) || c.tagName === 'UL' || c.tagName === 'OL' || c.tagName === 'TABLE' || c.tagName === 'HR') return true; return false; }
  D.hasBlockChild = hasBlockChild;
  /** Wrap stray inline content (text directly inside the root/td) into paragraphs so block ops have a target. */
  function normalizeRootInline(root) {
    let run = [];
    const flush = (before) => { if (!run.length) return; if (run.every(n => n.nodeType === 3 && !n.data.trim())) { run.forEach(n => n.remove()); run = []; return; } const p = document.createElement('p'); root.insertBefore(p, before); run.forEach(n => p.appendChild(n)); run = []; };
    [...root.childNodes].forEach(n => { const inline = n.nodeType === 3 || (n.nodeType === 1 && !BLOCK_TAGS.has(n.tagName) && !['UL', 'OL', 'TABLE', 'HR', 'FIGURE', 'SECTION'].includes(n.tagName)); if (inline) run.push(n); else flush(n); });
    flush(null);
  }
  D.normalizeRootInline = normalizeRootInline;
  /** All text blocks touched by the selection. */
  D.blocks = (r = D.range()) => {
    if (!r) return [];
    const root = D.rootOf(r.commonAncestorContainer); if (!root) return [];
    if (!D.blockOf(r.startContainer) || !D.blockOf(r.endContainer)) {
      const sv = saveSel(); normalizeRootInline(root); restoreSel(sv); r = D.range() || r;
    }
    const sb = D.blockOf(r.startContainer), eb = D.blockOf(r.endContainer);
    if (r.collapsed || sb === eb) return sb ? [sb] : [];
    const out = [];
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT);
    let n = walker.currentNode, on = false;
    while (n) {
      if (n === sb) on = true;
      if (on && D.blockOf(n) === n && !out.includes(n)) out.push(n);
      if (n === eb) break;
      n = walker.nextNode();
    }
    if (!out.includes(eb) && eb) out.push(eb);
    return out.filter(b => !out.some(o => o !== b && o.contains(b) && o.tagName === 'LI' && false));
  };
  D.selectedText = () => { const r = D.range(); return r ? r.toString() : ''; };

  /* ---------------- inline formatting ---------------- */
  D.exec = (cmd, val = null) => { D.ensureSel(); document.execCommand('styleWithCSS', false, true); return document.execCommand(cmd, false, val); };
  D.inline = (cmd, val) => D.change(() => D.exec(cmd, val));
  /** Wrap the selected text nodes, or (collapsed) the word at caret, in styled spans — used for font size etc. */
  function selectWordIfCollapsed() {
    const s = getSelection(); const r = D.range(); if (!r || !r.collapsed) return false;
    const n = r.startContainer; if (n.nodeType !== 3) return false;
    const t = n.data, o = r.startOffset;
    let a = o, b = o; while (a > 0 && /[\p{L}\p{N}_'’-]/u.test(t[a - 1])) a--; while (b < t.length && /[\p{L}\p{N}_'’-]/u.test(t[b])) b++;
    if (a === b) return false;
    const nr = document.createRange(); nr.setStart(n, a); nr.setEnd(n, b); s.removeAllRanges(); s.addRange(nr); return { n, o };
  }
  D.selectWordIfCollapsed = selectWordIfCollapsed;
  D.setFontSize = (pt) => {
    pt = Math.max(1, Math.min(1638, Math.round(pt * 2) / 2));
    D.change(() => {
      D.ensureSel(); const restoreTo = selectWordIfCollapsed();
      const r = D.range();
      if (r && r.collapsed) { insertStyledCaretSpan({ fontSize: pt + 'pt' }); return; }
      D.exec('fontSize', 7);
      const root = D.rootOf(D.range()?.commonAncestorContainer) || D.doc;
      $$('span[style*="xxx-large"], font[size="7"]', root).forEach(el => {
        if (el.tagName === 'FONT') { const sp = document.createElement('span'); sp.style.fontSize = pt + 'pt'; while (el.firstChild) sp.appendChild(el.firstChild); el.replaceWith(sp); }
        else el.style.fontSize = pt + 'pt';
      });
      if (restoreTo) D.setCaret(restoreTo.n.isConnected ? restoreTo.n : root, restoreTo.n.isConnected ? restoreTo.o : 0);
    });
  };
  /** At a collapsed caret, insert an empty styled span holding a zero-width char so the next typed text gets the style. */
  function insertStyledCaretSpan(style) {
    const r = D.range(); if (!r) return;
    const sp = document.createElement('span'); Object.assign(sp.style, style); sp.textContent = '​';
    r.insertNode(sp); D.setCaret(sp.firstChild, 1);
  }
  D.setFontName = (name) => {
    if (!name) return;
    D.change(() => { D.ensureSel(); const restoreTo = selectWordIfCollapsed(); const r = D.range(); if (r && r.collapsed) { insertStyledCaretSpan({ fontFamily: name }); return; } D.exec('fontName', name); if (restoreTo && restoreTo.n.isConnected) D.setCaret(restoreTo.n, restoreTo.o); });
    BTS.markDirty();
  };
  D.currentFontSize = () => {
    const r = D.range(); if (!r) return 11;
    let el = r.startContainer; if (el.nodeType === 3) el = el.parentElement;
    if (r.startContainer.nodeType === 1 && r.startContainer.childNodes[r.startOffset] && r.startContainer.childNodes[r.startOffset].nodeType === 1) el = r.startContainer.childNodes[r.startOffset];
    const px = parseFloat(getComputedStyle(el).fontSize) || 14.667; return Math.round(px * 0.75 * 2) / 2;
  };
  D.growFont = (dir) => {
    const cur = D.currentFontSize();
    let next;
    if (dir > 0) next = D.SIZES.find(s => s > cur + 0.01) || cur + 10; else next = [...D.SIZES].reverse().find(s => s < cur - 0.01) || Math.max(1, cur - 1);
    D.setFontSize(next);
  };
  D.changeCase = (mode) => {
    D.change(() => {
      D.ensureSel(); const restoreTo = selectWordIfCollapsed(); const r = D.range(); if (!r || r.collapsed) return;
      const nodes = textNodesIn(r);
      let sentenceStart = true; const prevText = textBefore(r, 2); if (prevText && !/[.!?]\s*$/.test(prevText) && prevText.trim()) sentenceStart = false;
      let wordStart = !prevText || /\s$/.test(prevText);
      nodes.forEach(({ node, start, end }) => {
        const t = node.data; let seg = t.slice(start, end);
        if (mode === 'upper') seg = seg.toUpperCase();
        else if (mode === 'lower') seg = seg.toLowerCase();
        else if (mode === 'toggle') seg = [...seg].map(c => c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase()).join('');
        else if (mode === 'title') { seg = [...seg].map(c => { const o = wordStart ? c.toUpperCase() : c.toLowerCase(); wordStart = /\s|-|\(/.test(c); return o; }).join(''); }
        else if (mode === 'sentence') { seg = [...seg].map(c => { let o = c.toLowerCase(); if (sentenceStart && /\p{L}/u.test(c)) { o = c.toUpperCase(); sentenceStart = false; } if (/[.!?]/.test(c)) sentenceStart = true; return o; }).join(''); }
        node.data = t.slice(0, start) + seg + t.slice(end);
      });
      const first = nodes[0], last = nodes[nodes.length - 1];
      if (first && last) { const nr = document.createRange(); nr.setStart(first.node, first.start); nr.setEnd(last.node, last.end); const s = getSelection(); s.removeAllRanges(); s.addRange(nr); }
      if (restoreTo && restoreTo.n.isConnected) D.setCaret(restoreTo.n, restoreTo.o);
    });
  };
  function textBefore(r, n) { const c = r.startContainer; if (c.nodeType === 3) return c.data.slice(Math.max(0, r.startOffset - n), r.startOffset); return ''; }
  function textNodesIn(r) {
    const out = []; const root = r.commonAncestorContainer.nodeType === 3 ? r.commonAncestorContainer.parentNode : r.commonAncestorContainer;
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT); let n;
    while ((n = w.nextNode())) {
      if (!r.intersectsNode(n)) continue;
      const start = n === r.startContainer ? r.startOffset : 0, end = n === r.endContainer ? r.endOffset : n.length;
      if (end > start) out.push({ node: n, start, end });
    }
    return out;
  }
  D.textNodesIn = textNodesIn;
  /** Wrap each selected text segment in an element produced by make(); returns created elements. */
  D.wrapText = (r, make) => {
    const segs = textNodesIn(r); const out = [];
    segs.forEach(({ node, start, end }) => {
      let n = node; if (start > 0) n = n.splitText(start); if (end - start < n.length) n.splitText(end - start);
      if (!n.data.length) return;
      const w = make(); n.parentNode.insertBefore(w, n); w.appendChild(n); out.push(w);
    });
    if (out.length) { const nr = document.createRange(); nr.setStartBefore(out[0]); nr.setEndAfter(out[out.length - 1]); const s = getSelection(); s.removeAllRanges(); s.addRange(nr); }
    return out;
  };
  D.clearFormatting = () => {
    D.change(() => {
      D.ensureSel(); const r = D.range(); if (!r) return;
      const blocks = D.blocks();
      if (r.collapsed) selectAllOf(blocks);
      D.exec('removeFormat');
      // also strip spans with font/color styles inside the selection
      const rr = D.range(); if (rr) $$('span, font, mark', D.rootOf(rr.commonAncestorContainer)).forEach(el => { if (rr.intersectsNode(el) && !el.matches('.cmt, .eq, .fld, .tab, .fn-ref, .dropcap')) unwrap(el); });
      blocks.forEach(b => { if (b.tagName !== 'LI' && b.tagName !== 'TD' && b.tagName !== 'TH') setBlockStyle(b, 'normal'); b.removeAttribute('style'); });
    });
  };
  function selectAllOf(blocks) { if (!blocks.length) return; const r = document.createRange(); r.setStart(blocks[0], 0); r.setEnd(blocks[blocks.length - 1], blocks[blocks.length - 1].childNodes.length); const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
  function unwrap(el) {
    const p = el.parentNode; if (!p) return;
    const s = getSelection(); const r = s.rangeCount ? s.getRangeAt(0) : null;
    const pt = r && el.contains(r.startContainer) || r && el.contains(r.endContainer) ? { sc: r.startContainer, so: r.startOffset, ec: r.endContainer, eo: r.endOffset } : null;
    const idx = [...p.childNodes].indexOf(el);
    while (el.firstChild) p.insertBefore(el.firstChild, el); el.remove();
    if (pt) { try { const fix = (c, o) => c === el ? [p, idx + o] : [c, o]; const [a, ao] = fix(pt.sc, pt.so), [b, bo] = fix(pt.ec, pt.eo); const nr = document.createRange(); nr.setStart(a, ao); nr.setEnd(b, bo); s.removeAllRanges(); s.addRange(nr); } catch { } }
  }
  D.unwrap = unwrap;

  /* ---------------- paragraph formatting ---------------- */
  D.align = (a) => D.change(() => { D.ensureSel(); const bl = D.blocks(); bl.forEach(b => { b.style.textAlign = a === 'left' ? '' : a; if (!b.getAttribute('style')) b.removeAttribute('style'); }); });
  D.lineSpacing = (v) => D.change(() => { D.ensureSel(); D.blocks().forEach(b => { b.style.lineHeight = String(v); }); });
  D.paraSpacing = (which, pt) => D.change(() => { D.ensureSel(); D.blocks().forEach(b => { b.style[which === 'before' ? 'marginTop' : 'marginBottom'] = pt + 'pt'; }); });
  D.indentPara = (side, cm) => D.change(() => { D.ensureSel(); D.blocks().forEach(b => { b.style[side === 'left' ? 'marginLeft' : 'marginRight'] = cm ? cm + 'cm' : ''; }); });
  D.indent = (dir) => {
    D.ensureSel();
    const r = D.range(); const li = r && D.closest(r.startContainer, 'li');
    if (li) { D.change(() => { if (dir > 0) indentListItems(); else outdentListItems(); }); return; }
    D.change(() => {
      D.blocks().forEach(b => {
        const cur = D.toPx(b.style.marginLeft) || 0; const step = 0.5 * PX;
        const next = Math.max(0, dir > 0 ? cur + step : cur - step);
        b.style.marginLeft = next ? (next / PX).toFixed(2).replace(/\.?0+$/, '') + 'in' : '';
      });
    });
  };
  function indentListItems() {
    const items = D.blocks().map(b => b.closest('li')).filter((v, i, a) => v && a.indexOf(v) === i);
    items.forEach(li => {
      const prev = li.previousElementSibling; const list = li.parentElement;
      if (!prev || prev.tagName !== 'LI') { return; }
      let sub = [...prev.children].reverse().find(c => c.tagName === 'UL' || c.tagName === 'OL');
      if (!sub) { sub = document.createElement(list.tagName); const lt = nextLevelType(list); if (lt) sub.style.listStyleType = lt; prev.appendChild(sub); }
      sub.appendChild(li);
    });
    keepCaret(items[0]);
  }
  function nextLevelType(list) {
    const t = list.style.listStyleType;
    if (list.tagName === 'UL') return ({ disc: 'circle', circle: 'square', square: 'disc' })[t || 'disc'] || '';
    return ({ decimal: 'lower-alpha', '': 'lower-alpha', 'lower-alpha': 'lower-roman', 'lower-roman': 'decimal', 'upper-alpha': 'lower-alpha', 'upper-roman': 'upper-alpha', 'paren-decimal': 'paren-lower-alpha', 'paren-lower-alpha': 'lower-roman' })[t] || 'lower-alpha';
  }
  function outdentListItems() {
    const items = D.blocks().map(b => b.closest('li')).filter((v, i, a) => v && a.indexOf(v) === i);
    [...items].reverse().forEach(li => {
      const list = li.parentElement; const parentLi = list.parentElement && list.parentElement.closest('li');
      // items after li in same list move into a new sublist of li
      const after = []; let n = li.nextElementSibling; while (n) { after.push(n); n = n.nextElementSibling; }
      if (after.length) { const sub = document.createElement(list.tagName); sub.setAttribute('style', list.getAttribute('style') || ''); after.forEach(a => sub.appendChild(a)); li.appendChild(sub); }
      if (parentLi && D.rootOf(parentLi)) { parentLi.after(li); }
      else {
        // leave the list entirely → becomes paragraph(s)
        const p = document.createElement('p');
        const subs = [...li.children].filter(c => c.tagName === 'UL' || c.tagName === 'OL');
        while (li.firstChild && !subs.includes(li.firstChild)) p.appendChild(li.firstChild);
        [...li.childNodes].forEach(c => { if (!subs.includes(c)) p.appendChild(c); });
        if (!p.childNodes.length) p.innerHTML = '<br>';
        // split the list
        const tail = document.createElement(list.tagName); tail.setAttribute('style', list.getAttribute('style') || '');
        n = li.nextElementSibling; while (n) { const nx = n.nextElementSibling; tail.appendChild(n); n = nx; }
        list.after(p); subs.forEach(s => { const lt = s; tail.prepend(...lt.children); lt.remove(); });
        if (tail.children.length) p.after(tail);
        li.remove(); if (!list.children.length) list.remove();
        keepCaret(p);
      }
      if (!list.children.length) list.remove();
    });
    if (items[0] && items[0].isConnected) keepCaret(items[0]);
  }
  function keepCaret(el) { if (!el || !el.isConnected) return; const s = getSelection(); if (s.rangeCount && el.contains(s.getRangeAt(0).startContainer)) return; const r = document.createRange(); r.selectNodeContents(el); r.collapse(false); s.removeAllRanges(); s.addRange(r); }

  /** Selection as (leaf-block index, char offset in block) — robust to DOM restructuring like li <-> p. */
  function leafBlocks(root) { const out = []; const w = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT); let n; while ((n = w.nextNode())) if (D.blockOf(n) === n) out.push(n); return out; }
  function textOffsets() {
    const r = D.range(); if (!r) return null; const root = D.rootOf(r.commonAncestorContainer); const bl = leafBlocks(root);
    const pt = (c, o) => { const b = D.blockOf(c); const bi = bl.indexOf(b); if (bi < 0) return null; const t = document.createRange(); t.setStart(b, 0); t.setEnd(c, o); return { bi, off: t.toString().length }; };
    return { root, a: pt(r.startContainer, r.startOffset), b: pt(r.endContainer, r.endOffset), collapsed: r.collapsed };
  }
  function restoreOffsets(o) {
    if (!o || !o.root.isConnected || !o.a) return;
    const bl = leafBlocks(o.root);
    const find = (p) => {
      const b = bl[Math.min(p.bi, bl.length - 1)]; if (!b) return null;
      const w = document.createTreeWalker(b, NodeFilter.SHOW_TEXT); let t, pos = 0;
      while ((t = w.nextNode())) { if (D.blockOf(t) !== b) continue; if (pos + t.length >= p.off) return [t, p.off - pos]; pos += t.length; }
      return [b, 0];
    };
    const a = find(o.a), b = o.collapsed || !o.b ? a : find(o.b); if (!a) return;
    const r = document.createRange(); r.setStart(a[0], a[1]); if (b) r.setEnd(b[0], b[1]); else r.collapse(true);
    const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  }
  D.textOffsets = textOffsets; D.restoreOffsets = restoreOffsets;
  D.toggleList = (ordered, styleType) => {
    D.ensureSel();
    const off = textOffsets();
    D.change(() => {
      const r = D.range(); const tag = ordered ? 'OL' : 'UL';
      const li = r && D.closest(r.startContainer, 'li'); const list = li && li.parentElement;
      if (list && list.tagName === tag && styleType === undefined) { removeList(off); return; }
      if (list && list.tagName !== tag) { const nl = document.createElement(tag); [...list.attributes].forEach(a => nl.setAttribute(a.name, a.value)); while (list.firstChild) nl.appendChild(list.firstChild); list.replaceWith(nl); if (styleType !== undefined) nl.style.listStyleType = styleType; return; }
      if (list && styleType !== undefined) { list.style.listStyleType = styleType; return; }
      // convert headings/quotes etc. to p first so lists don't nest weirdly
      D.blocks().forEach(b => { if (/^H[1-6]$|^PRE$|^BLOCKQUOTE$/.test(b.tagName)) setBlockStyle(b, 'normal'); });
      D.exec(ordered ? 'insertOrderedList' : 'insertUnorderedList');
      normalizeRootInline(off ? off.root : D.doc);
      restoreOffsets(off);
      const r2 = D.range(); const nl = r2 && D.closest(r2.startContainer, tag.toLowerCase());
      if (nl && styleType) nl.style.listStyleType = styleType;
      cleanupListMarkup();
    });
    restoreOffsets(off);
  };
  /** Run fn and put the selection back on the same nodes (survives moves of ancestors). */
  D.preserve = (fn) => keepSelection(() => fn());
  function removeList(off) {
    for (let i = 0; i < 10; i++) {
      const items = D.blocks().map(b => b.closest('li')).filter(Boolean); if (!items.length) break;
      outdentListItems(); restoreOffsets(off);
    }
  }
  function cleanupListMarkup() { D.preserve(cleanupListMarkup2); }
  function cleanupListMarkup2() {
    // Chromium sometimes wraps lists in <p> or leaves <span style> junk; unwrap lists inside p
    $$('p > ul, p > ol', D.doc).forEach(l => { const p = l.parentElement; if (p.childNodes.length === 1) p.replaceWith(l); });
    $$('li > p:only-child', D.doc).forEach(p => { if (!p.attributes.length) unwrap(p); });
  }

  /** Run fn(map) where fn may replace block elements; keeps the selection, mapping replaced containers via map. */
  function keepSelection(fn) {
    const r = D.range(); const pt = r ? { sc: r.startContainer, so: r.startOffset, ec: r.endContainer, eo: r.endOffset } : null;
    const map = new Map(); fn(map);
    if (!pt) return;
    const m = (n) => map.get(n) || n;
    try {
      const sc = m(pt.sc), ec = m(pt.ec); if (!sc.isConnected || !ec.isConnected) return;
      const nr = document.createRange(); nr.setStart(sc, Math.min(pt.so, sc.nodeType === 3 ? sc.length : sc.childNodes.length)); nr.setEnd(ec, Math.min(pt.eo, ec.nodeType === 3 ? ec.length : ec.childNodes.length));
      const s = getSelection(); s.removeAllRanges(); s.addRange(nr);
    } catch { }
  }
  D.keepSelection = keepSelection;
  D.setBlockStyle = (id) => {
    D.ensureSel();
    D.change(() => keepSelection(map => {
      const blocks = D.blocks();
      blocks.forEach(b => { const nb = setBlockStyle(b, id); if (nb && nb !== b) map.set(b, nb); });
    }));
  };
  function setBlockStyle(b, id) {
    const st = D.STYLES.find(s => s.id === id) || D.STYLES[0];
    if (b.tagName === 'TD' || b.tagName === 'TH' || b.tagName === 'DIV') {
      // wrap cell/div content into a paragraph so it can take a style
      const p = document.createElement(st.tag); if (st.cls) p.className = st.cls;
      while (b.firstChild) p.appendChild(b.firstChild); if (!p.childNodes.length) p.innerHTML = '<br>'; b.appendChild(p); return p;
    }
    if (b.tagName === 'LI') {
      if (st.tag === 'p') { b.className = ''; return b; }
      // headings inside list: wrap content
      return b;
    }
    let el = b;
    if (b.tagName !== st.tag.toUpperCase()) {
      el = document.createElement(st.tag);
      [...b.attributes].forEach(a => { if (a.name !== 'class') el.setAttribute(a.name, a.value); });
      if (b.tagName === 'BLOCKQUOTE' && hasBlockChild(b)) { while (b.firstChild) el.appendChild(b.firstChild); }
      else while (b.firstChild) el.appendChild(b.firstChild);
      if (st.tag === 'pre') $$('br', el).forEach(br => br.replaceWith('\n'));
      b.replaceWith(el);
    }
    el.className = st.cls;
    if (!el.className) el.removeAttribute('class');
    if (!el.childNodes.length) el.innerHTML = '<br>';
    return el;
  }
  D._setBlockStyle = setBlockStyle;
  D.styleOf = (b) => {
    if (!b) return 'normal';
    const t = b.tagName;
    if (/^H[1-3]$/.test(t)) return 'h' + t[1];
    if (t === 'BLOCKQUOTE' || b.closest('blockquote')) return 'quote';
    if (t === 'PRE') return 'code';
    for (const c of ['title', 'subtitle', 'nospace', 'caption']) if (b.classList.contains(c)) return c;
    return 'normal';
  };
  /* live preview */
  let preview = null;
  D.previewStyle = (id) => {
    D.endPreview();
    const r = D.range() || lastRange; if (!r) return;
    let blocks; try { blocks = D.blocks(r); } catch { return; }
    preview = blocks.filter(b => b.isConnected).map(b => { const pv = 'pv-' + (id === 'normal' ? 'normal' : id); b.classList.add('pv-on', pv); return [b, pv]; });
  };
  D.endPreview = () => { if (!preview) return; preview.forEach(([b, pv]) => { b.classList.remove('pv-on', pv); if (!b.className) b.removeAttribute('class'); }); preview = null; };

  /* paragraph borders & shading */
  D.paraShading = (color) => D.change(() => { D.ensureSel(); D.blocks().forEach(b => { b.style.backgroundColor = color === 'none' ? '' : color; fixPad(b); }); });
  D.paraBorder = (kind) => D.change(() => {
    D.ensureSel(); const blocks = D.blocks(); const line = '1px solid #3c4250';
    blocks.forEach((b, i) => {
      if (kind === 'none') { ['border', 'borderTop', 'borderBottom', 'borderLeft', 'borderRight'].forEach(p => b.style[p] = ''); }
      else if (kind === 'box') { b.style.borderTop = i === 0 ? line : ''; b.style.borderBottom = i === blocks.length - 1 ? line : ''; b.style.borderLeft = line; b.style.borderRight = line; }
      else if (kind === 'top') b.style.borderTop = line;
      else if (kind === 'bottom') b.style.borderBottom = line;
      else if (kind === 'left') b.style.borderLeft = '3px solid #3c4250';
      else if (kind === 'right') b.style.borderRight = line;
      fixPad(b);
    });
  });
  function fixPad(b) {
    const has = b.style.backgroundColor || b.style.borderTop || b.style.borderBottom || b.style.borderLeft || b.style.borderRight;
    b.style.padding = has ? '2pt 4pt' : '';
    if (!b.getAttribute('style')) b.removeAttribute('style');
  }

  /* ---------------- page breaks & structural insert ---------------- */
  /** Top-level child of root containing node. */
  D.topBlock = (node) => { let n = node; while (n && n.parentNode !== D.doc) n = n.parentNode; return n; };
  /** Split the top-level block at the caret, returning the element after which to insert. */
  D.splitAtCaret = () => {
    const r = D.ensureSel(); if (!D.doc.contains(r.startContainer)) { D.doc.focus(); }
    const rr = D.range() || r; rr.deleteContents();
    const top = D.topBlock(rr.startContainer);
    if (!top || top === D.doc) { return { before: D.doc.lastChild, after: null }; }
    if (top.nodeType !== 1 || /^(TABLE|HR|DIV)$/.test(top.tagName)) return { before: top, after: top.nextSibling };
    const tail = document.createRange(); tail.setStart(rr.startContainer, rr.startOffset); tail.setEnd(top, top.childNodes.length);
    const atEnd = !tail.toString().trim() && !tail.cloneContents().querySelector('img,table,hr,.eq');
    const head = document.createRange(); head.setStart(top, 0); head.setEnd(rr.startContainer, rr.startOffset);
    const atStart = !head.toString().trim() && !head.cloneContents().querySelector('img,table,hr,.eq');
    if (atEnd) return { before: top, after: top.nextSibling };
    if (atStart) return { before: top.previousSibling, after: top, emptyBefore: true };
    const frag = tail.extractContents();
    const nb = top.cloneNode(false); nb.removeAttribute('id'); nb.appendChild(frag); top.after(nb);
    if (!top.textContent && !top.querySelector('img,br')) top.innerHTML = '<br>';
    return { before: top, after: nb };
  };
  /** Insert block node(s) at caret (split the current block). Caret goes after. */
  D.insertBlocks = (nodes, caretInto) => {
    const { before, after } = D.splitAtCaret();
    const ref = after || null;
    nodes.forEach(n => D.doc.insertBefore(n, ref));
    let next = ref;
    if (!next || next.nodeType !== 1) { next = document.createElement('p'); next.innerHTML = '<br>'; D.doc.insertBefore(next, ref); }
    const target = caretInto || next;
    const rr = document.createRange(); rr.selectNodeContents(target); rr.collapse(true); D.doc.focus({ preventScroll: true }); const s = getSelection(); s.removeAllRanges(); s.addRange(rr);
    return next;
  };
  D.pageBreak = () => D.change(() => { const hr = document.createElement('hr'); hr.className = 'page-break'; D.insertBlocks([hr]); D.scrollCaret(); });
  D.blankPage = () => D.change(() => {
    const a = document.createElement('hr'); a.className = 'page-break';
    const p = document.createElement('p'); p.innerHTML = '<br>';
    const b = document.createElement('hr'); b.className = 'page-break';
    D.insertBlocks([a, p, b], p); D.scrollCaret();
  });

  /* ---------------- settings ---------------- */
  D.settings = () => ({ page: D.page, design: D.design });
  D.applySettings = () => {
    const th = D.theme();
    [D.doc, D.hdr, D.ftr, $('#fn-area')].forEach(el => {
      if (!el) return;
      el.style.setProperty('--th-accent', th.accent); el.style.setProperty('--th-dark', th.dark);
      el.style.setProperty('--th-head', D.fontStack(D.design.headFont || th.head)); el.style.setProperty('--th-body', D.fontStack(D.design.bodyFont || th.body));
      const ps = D.design.pspacing; if (ps) { el.style.setProperty('--p-after', ps.after + 'pt'); el.style.setProperty('--p-line', ps.line); el.style.setProperty('--p-before', (ps.before || 0) + 'pt'); }
      else { el.style.removeProperty('--p-after'); el.style.removeProperty('--p-line'); el.style.removeProperty('--p-before'); }
    });
    [D.hdr, D.ftr].forEach(el => { el.classList.add('doc-content'); });
    D.layout();
    D.syncRibbon && D.syncRibbon();
  };
  D.setPage = (patch) => D.change(() => { Object.assign(D.page, patch); D.applySettings(); });
  D.setMargins = (m) => D.change(() => { D.page.m = { ...m }; D.applySettings(); });
  D.setDesign = (patch) => D.change(() => { Object.assign(D.design, patch); D.applySettings(); });

  /* ---------------- layout & pagination ---------------- */
  let layoutTimer = 0;
  D.scheduleLayout = (ms = 120) => { clearTimeout(layoutTimer); layoutTimer = setTimeout(() => D.layout(), ms); };
  D.eff = { t: 96, b: 96, l: 96, r: 96 };
  D.layout = () => {
    if (!D.doc) return;
    const ps = D.paperSize(); const m = D.page.m;
    const W = ps.w * PX, H = ps.h * PX;
    const sheet = D.sheet, bw = $('#body-wrap');
    sheet.style.width = W + 'px';
    sheet.style.background = D.design.pageColor || '#fff';
    // header/footer boxes
    const hd = D.page.hdrDist * PX, fd = D.page.ftrDist * PX;
    Object.assign(D.hdr.style, { top: hd + 'px', left: m.l * PX + 'px', right: m.r * PX + 'px' });
    Object.assign(D.ftr.style, { bottom: fd + 'px', left: m.l * PX + 'px', right: m.r * PX + 'px' });
    const hh = D.hdr.offsetHeight, fh = D.ftr.offsetHeight;
    const hasH = !!D.hdr.textContent.trim() || !!D.hdr.querySelector('img'), hasF = !!D.ftr.textContent.trim() || !!D.ftr.querySelector('img');
    const eff = D.eff = {
      t: Math.max(m.t * PX, hasH ? hd + hh + 8 : 0), b: Math.max(m.b * PX, hasF ? fd + fh + 8 : 0), l: m.l * PX, r: m.r * PX,
    };
    bw.style.padding = `${eff.t}px ${eff.r}px ${eff.b}px ${eff.l}px`;
    const contentH = Math.max(60, H - eff.t - eff.b); D.contentH = contentH; D.pageH = H; D.pageW = W;
    bw.style.minHeight = '';
    const pg = paginate(contentH);
    D.pages = pg.breaks; D.pageCount = pg.breaks.length;
    const lastStart = pg.breaks[pg.breaks.length - 1];
    const fnH = $('#fn-area').offsetHeight;
    const natural = D.doc.offsetHeight + fnH;
    const minBody = lastStart + contentH;
    bw.style.minHeight = (Math.max(natural, minBody) + eff.t + eff.b) + 'px';
    drawGuides(pg, W, eff);
    drawRuler(W, m);
    D.positionComments && D.positionComments();
    D.updateImgHandles && D.updateImgHandles();
    D.updateStatus && D.updateStatus();
    D.updateTableTools && D.updateTableTools();
  };
  /** Compute natural (continuous) y-coordinates where each printed page starts. */
  function paginate(H) {
    const doc = D.doc; const dr = doc.getBoundingClientRect(); const scale = dr.height / (doc.offsetHeight || 1) || 1;
    const rel = (rect) => ({ top: (rect.top - dr.top) / scale, bottom: (rect.bottom - dr.top) / scale });
    doc.querySelectorAll('hr.page-break[style]').forEach(h => h.removeAttribute('style'));
    const leaves = []; collectLeaves(doc, leaves);
    const fn = $('#fn-area'); if (fn.firstChild) leaves.push({ el: fn, split: true });
    const breaks = [0]; const headPage = new Map();
    let start = 0;
    for (const lf of leaves) {
      const el = lf.el;
      if (lf.forced) {
        let r = rel(el.getBoundingClientRect());
        while (r.top >= start + H - 0.5) { start += H; breaks.push(start); }
        const fill = Math.max(18, Math.floor(start + H - r.top - 8));
        el.style.height = fill + 'px';
        r = rel(el.getBoundingClientRect()); start = Math.max(r.bottom, start + 1); breaks.push(start); continue;
      }
      let r = lf.rect ? lf.rect() : rel(el.getBoundingClientRect());
      if (lf.rectFn) r = lf.rectFn(rel);
      if (r.bottom - r.top <= 0) continue;
      let guard = 0;
      while (r.bottom > start + H + 0.5 && guard++ < 500) {
        const limit = start + H;
        if (r.top >= limit) { start = limit; breaks.push(start); continue; }
        if (!lf.split || r.bottom - r.top < 30) {
          if (r.top > start + 1) { start = r.top; breaks.push(start); continue; }
          start = limit; breaks.push(start); continue; // taller than a page: let it overflow
        }
        // split between line boxes
        const lines = lineTops(el, rel);
        let bp = null;
        for (let i = 0; i < lines.length; i++) { if (lines[i].bottom > limit + 0.5) { bp = lines[i].top; break; } }
        if (bp == null || bp <= start + 1) bp = (bp != null && bp <= start + 1) ? limit : (r.top > start + 1 ? r.top : limit);
        start = bp; breaks.push(start);
      }
      if (/^H[1-6]$/.test(el.tagName)) headPage.set(el, breaks.length);
    }
    D.headPage = headPage;
    return { breaks, scale };
  }
  D.paginate = paginate;
  function lineTops(el, rel) {
    const rg = document.createRange(); rg.selectNodeContents(el);
    const rects = [...rg.getClientRects()].map(rel).filter(r => r.bottom - r.top > 2);
    rects.sort((a, b) => a.top - b.top);
    const lines = [];
    rects.forEach(r => { const l = lines.find(x => Math.abs(x.top - r.top) < 3 || (r.top >= x.top && r.bottom <= x.bottom)); if (l) { l.top = Math.min(l.top, r.top); l.bottom = Math.max(l.bottom, r.bottom); } else lines.push({ ...r }); });
    lines.sort((a, b) => a.top - b.top);
    return lines;
  }
  function collectLeaves(parent, out) {
    for (const el of parent.children) {
      const t = el.tagName;
      if (t === 'HR' && el.classList.contains('page-break')) { out.push({ el, forced: true }); continue; }
      if (t === 'UL' || t === 'OL') { collectLeaves(el, out); continue; }
      if (t === 'LI') {
        const sub = [...el.children].find(c => c.tagName === 'UL' || c.tagName === 'OL');
        if (sub) { out.push({ el, split: true, rectFn: (rel) => { const a = rel(el.getBoundingClientRect()); const b = rel(sub.getBoundingClientRect()); return { top: a.top, bottom: b.top }; } }); collectLeaves(el, out); }
        else out.push({ el, split: !el.querySelector('img,table') });
        continue;
      }
      if (t === 'TABLE') { for (const row of el.rows) out.push({ el: row, split: false }); continue; }
      if ((t === 'BLOCKQUOTE' || t === 'DIV' || t === 'SECTION') && hasBlockChild(el) && !el.classList.contains('cols') && !el.classList.contains('textbox')) { collectLeaves(el, out); continue; }
      const hasImg = !!el.querySelector('img,table,.textbox');
      out.push({ el, split: !hasImg && !/^H[1-6]$/.test(t) && t !== 'HR' });
    }
  }
  function drawGuides(pg, W, eff) {
    const g = $('#guides'); const wm = $('#wm-layer'); const pb = $('#pborder');
    const H = D.pageH;
    // page seams (dashed) + labels
    let html = '';
    const seamY = pg.breaks.map(b => eff.t + b);
    pg.breaks.forEach((b, i) => {
      if (i === 0) { html += `<div class="pnum-lbl" style="top:8px">Page 1</div>`; return; }
      const y = seamY[i];
      html += `<div class="seam" style="top:${y}px"><span class="seam-l"></span><span class="seam-r"></span></div><div class="pnum-lbl" style="top:${y + 4}px">Page ${i + 1}</div>`;
    });
    g.innerHTML = html;
    // watermark per page segment
    const sheetH = D.sheet.offsetHeight || H;
    const segs = seamY.map((y, i) => ({ top: i === 0 ? 0 : y - eff.t * 0.5, bottom: i === seamY.length - 1 ? sheetH : seamY[i + 1] - eff.t * 0.5 }));
    const w = D.design.watermark;
    wm.innerHTML = w ? segs.map(s => `<div class="wm" style="top:${s.top}px;height:${s.bottom - s.top}px">${watermarkInner(w, W)}</div>`).join('') : '';
    const bd = D.design.border;
    pb.innerHTML = bd ? `<div class="pborder" style="${borderCSS(bd)};top:${0.35 * PX}px;bottom:${0.35 * PX}px;left:${0.35 * PX}px;right:${0.35 * PX}px"></div>` : '';
    D.sheet.classList.toggle('has-wm', !!w);
  }
  function watermarkInner(w, W) {
    const size = Math.min(130, Math.max(28, (W * (w.layout === 'horizontal' ? 0.8 : 1.0)) / Math.max(4, w.text.length) * 1.25));
    return `<span style="color:${w.color || '#c0c4cc'};opacity:${w.opacity ?? 0.35};font-size:${size}px;transform:rotate(${w.layout === 'horizontal' ? 0 : -45}deg);font-family:${D.fontStack(w.font || 'Calibri')}">${esc(w.text)}</span>`;
  }
  D.watermarkInner = watermarkInner;
  function borderCSS(bd) {
    const style = bd.style === 'double' ? 'double' : bd.style === 'dashed' ? 'dashed' : bd.style === 'dotted' ? 'dotted' : 'solid';
    const w = bd.style === 'double' ? Math.max(3, bd.width || 3) : (bd.width || 1.5);
    let s = `border:${w}px ${style} ${bd.color || '#1f2533'}`;
    if (bd.style === 'art') s = `border:3px solid ${bd.color || '#1f2533'};outline:1px solid ${bd.color || '#1f2533'};outline-offset:3px`;
    return s;
  }
  D.borderCSS = borderCSS;

  /* ---------------- ruler ---------------- */
  function drawRuler(W, m) {
    const ri = $('#ruler-inner'); ri.style.width = W + 'px';
    const rail = $('#crail'); $('#ruler').style.paddingRight = rail.classList.contains('hidden') ? '' : ((rail.offsetWidth + 18) * D.zoom / 100) + 'px';
    const metric = true; const unit = metric ? PX / 2.54 : PX / 8;
    let html = `<div class="r-margin" style="left:0;width:${m.l * PX}px"></div><div class="r-margin" style="right:0;width:${m.r * PX}px"></div>`;
    const origin = m.l * PX;
    for (let x = origin % unit, i = -Math.floor(origin / unit); x <= W; x += unit, i++) {
      const isCm = metric; const big = isCm ? true : i % 8 === 0;
      if (i !== 0 && big) html += `<span class="r-num" style="left:${x}px">${Math.abs(i)}</span>`;
      html += `<span class="r-tick" style="left:${x}px;height:${i === 0 ? 10 : 3}px"></span>`;
      if (isCm) html += `<span class="r-tick" style="left:${x + unit / 2}px;height:5px"></span>`;
    }
    html += `<span class="r-handle" data-h="l" style="left:${m.l * PX}px" title="Left margin"></span><span class="r-handle" data-h="r" style="left:${W - m.r * PX}px" title="Right margin"></span>`;
    ri.innerHTML = html;
  }
  function wireRuler() {
    const ri = $('#ruler-inner'); let drag = null; let tip = null;
    ri.addEventListener('mousedown', e => {
      const h = e.target.closest('.r-handle'); if (!h) return; e.preventDefault();
      D.checkpoint();
      drag = { side: h.dataset.h, x0: e.clientX, m0: { ...D.page.m }, scale: ri.getBoundingClientRect().width / ri.offsetWidth };
      tip = document.createElement('div'); tip.id = 'ruler-tip'; document.body.appendChild(tip);
    });
    addEventListener('mousemove', e => {
      if (!drag) return;
      const dx = (e.clientX - drag.x0) / drag.scale / PX; const ps = D.paperSize();
      const m = { ...drag.m0 };
      if (drag.side === 'l') m.l = Math.max(0.2, Math.min(ps.w - m.r - 1, drag.m0.l + dx)); else m.r = Math.max(0.2, Math.min(ps.w - m.l - 1, drag.m0.r - dx));
      m.l = Math.round(m.l * 20) / 20; m.r = Math.round(m.r * 20) / 20;
      D.page.m = m; D.layout();
      tip.textContent = (drag.side === 'l' ? 'Left margin: ' : 'Right margin: ') + (m[drag.side] * 2.54).toFixed(2) + ' cm';
      tip.style.left = e.clientX + 12 + 'px'; tip.style.top = e.clientY + 16 + 'px';
    });
    addEventListener('mouseup', () => { if (!drag) return; drag = null; tip && tip.remove(); D.afterChange(); });
  }

  /* ---------------- zoom ---------------- */
  D.setZoom = (z) => {
    D.zoom = z; const s = z / 100;
    $('#sheet-row').style.zoom = s; $('#ruler-inner').style.zoom = s;
    D.updateImgHandles && D.updateImgHandles();
  };
  D.fitZoom = (mode) => {
    const c = D.canvas; const ps = D.paperSize();
    const availW = c.clientWidth - 70 - ($('#crail').classList.contains('hidden') ? 0 : 270), availH = c.clientHeight - 50;
    let z = mode === 'width' ? availW / (ps.w * PX) : Math.min(availW / (ps.w * PX), availH / (ps.h * PX));
    BTS.setZoom(Math.max(20, Math.min(300, Math.floor(z * 100))));
  };

  /* ---------------- status & ribbon sync ---------------- */
  D.countWords = (text) => { const m = text.replace(/[​]/g, '').match(/[\p{L}\p{N}][\p{L}\p{N}\p{M}'’_-]*/gu); return m ? m.length : 0; };
  D.docText = () => D.doc.innerText || '';
  D.caretPage = () => {
    const r = D.range(); if (!r || !D.doc.contains(r.startContainer)) return D.lastCaretPage || 1;
    let rect = r.getClientRects()[0] || (r.startContainer.nodeType === 1 ? r.startContainer.getBoundingClientRect() : r.startContainer.parentElement.getBoundingClientRect());
    const dr = D.doc.getBoundingClientRect(); const scale = dr.height / (D.doc.offsetHeight || 1) || 1;
    const y = (rect.top - dr.top) / scale + 1;
    let p = 1; D.pages.forEach((b, i) => { if (y >= b) p = i + 1; });
    D.lastCaretPage = p; return p;
  };
  let statusTimer = 0;
  D.updateStatus = () => {
    clearTimeout(statusTimer);
    statusTimer = setTimeout(() => {
      const words = D.countWords(D.docText());
      const sel = D.range(); const selWords = sel && !sel.collapsed ? D.countWords(sel.toString()) : 0;
      BTS.setStatus('page', `<button class="sb-btn" data-cmd="goto-page" title="Go to page">Page ${D.caretPage()} of ${D.pageCount}</button>`);
      BTS.setStatus('words', `<button class="sb-btn" data-cmd="word-count" title="Word count">${selWords ? selWords.toLocaleString('en-IN') + ' of ' : ''}${words.toLocaleString('en-IN')} word${words === 1 ? '' : 's'}</button>`);
      BTS.setStatus('lang', `${icon('languages', 14)}<span title="Proofing language">English (India)</span>`);
      BTS.setStatusRight(`<button class="sb-btn ${D.focusMode ? 'on' : ''}" data-cmd="focus-mode" title="Focus mode">${icon('scan', 14)}</button><button class="sb-btn" data-cmd="read-aloud" title="Read aloud">${icon('volume-2', 14)}</button>`);
    }, 60);
  };

  D.scrollCaret = () => {
    const r = D.range(); if (!r) return;
    const rect = r.getClientRects()[0] || (r.startContainer.nodeType === 1 ? r.startContainer.getBoundingClientRect() : null); if (!rect) return;
    const c = D.canvas.getBoundingClientRect();
    if (rect.bottom > c.bottom - 30) D.canvas.scrollTop += rect.bottom - c.bottom + 80;
    else if (rect.top < c.top + 40) D.canvas.scrollTop -= c.top + 80 - rect.top;
  };
  D.scrollToEl = (el, select) => {
    if (!el) return;
    const c = D.canvas.getBoundingClientRect(); const r = el.getBoundingClientRect();
    D.canvas.scrollTop += r.top - c.top - 90;
    if (select) { const rg = document.createRange(); rg.selectNodeContents(el); rg.collapse(true); (D.rootOf(el) || D.doc).focus({ preventScroll: true }); const s = getSelection(); s.removeAllRanges(); s.addRange(rg); }
  };

  /* ---------------- after change / input ---------------- */
  D.afterChange = (opts = {}) => {
    if (D.loading) return;
    BTS.markDirty();
    ensureNotEmpty();
    D.scheduleLayout(opts.now ? 0 : 140);
    D.onChanged && D.onChanged();
  };
  function ensureNotEmpty() {
    if (!D.doc.firstChild || (D.doc.childNodes.length === 1 && D.doc.firstChild.nodeType === 3 && !D.doc.firstChild.data.trim())) { D.doc.innerHTML = '<p><br></p>'; D.setCaret(D.doc.firstChild, 0); }
    else if (D.doc.childNodes.length === 1 && D.doc.firstChild.nodeName === 'BR') { D.doc.innerHTML = '<p><br></p>'; D.setCaret(D.doc.firstChild, 0); }
  }

  /* ---------------- HTML normalization (for loading/pasting) ---------------- */
  D.normalizeWhitespace = (root) => {
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT); const kill = []; let n;
    while ((n = w.nextNode())) {
      if (n.parentElement && n.parentElement.closest('pre, textarea, .tab')) continue;
      const blockParent = n.parentElement && (BLOCK_TAGS.has(n.parentElement.tagName) || ['UL', 'OL', 'TABLE', 'TBODY', 'THEAD', 'TR', 'BODY', 'SECTION'].includes(n.parentElement.tagName) || n.parentElement === root);
      if (!n.data.trim()) {
        const prev = n.previousSibling, next = n.nextSibling;
        const isBlock = (x) => !x || (x.nodeType === 1 && (BLOCK_TAGS.has(x.tagName) || ['UL', 'OL', 'TABLE', 'HR', 'BR', 'TBODY', 'TR'].includes(x.tagName)));
        if (blockParent && (isBlock(prev) || isBlock(next))) { kill.push(n); continue; }
        if (['UL', 'OL', 'TABLE', 'TBODY', 'THEAD', 'TR'].includes(n.parentElement && n.parentElement.tagName)) { kill.push(n); continue; }
      }
      n.data = n.data.replace(/[\t\n\r ]+/g, ' ');
    }
    kill.forEach(k => k.remove());
  };
  /** Load HTML content into the document body (no history). */
  D.setContent = (html, { hdr = '', ftr = '' } = {}) => {
    D.loading = true;
    const tmp = document.createElement('div'); tmp.innerHTML = html || '';
    D.normalizeWhitespace(tmp);
    normalizeRootInline(tmp);
    D.doc.innerHTML = tmp.innerHTML || '<p><br></p>';
    if (!D.doc.firstElementChild) D.doc.innerHTML = '<p><br></p>';
    D.hdr.innerHTML = hdr; D.ftr.innerHTML = ftr;
    D.loading = false;
    D.renderFootnotes && D.renderFootnotes();
    D.renderComments && D.renderComments();
    D.applySettings();
    D.clearHistory();
  };

  /* ---------------- init ---------------- */
  D.initEditor = () => {
    D.doc = $('#doc'); D.hdr = $('#hdr'); D.ftr = $('#ftr'); D.sheet = $('#sheet'); D.canvas = $('#canvas');
    document.execCommand('defaultParagraphSeparator', false, 'p');
    document.execCommand('styleWithCSS', false, true);
    [D.doc, D.hdr, D.ftr].forEach(root => {
      root.addEventListener('beforeinput', onBeforeInput);
      root.addEventListener('input', () => D.afterChange());
    });
    document.addEventListener('selectionchange', () => {
      const r = D.range(); if (r) lastRange = r.cloneRange();
      D.onSelection && D.onSelection();
    });
    wireRuler();
    D.setContent('<p><br></p>');
    new ResizeObserver(() => D.scheduleLayout(60)).observe(D.doc);
    new ResizeObserver(() => D.scheduleLayout(60)).observe(D.hdr);
    new ResizeObserver(() => D.scheduleLayout(60)).observe(D.ftr);
    if (document.fonts) document.fonts.ready.then(() => D.layout());
  };
})();
