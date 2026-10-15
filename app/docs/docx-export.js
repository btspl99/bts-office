/* BTS Docs — HTML → DOCX converter (docx v9, window.docx).
 * Works from the live DOM + computed styles so what you see is what Word gets:
 * styles/headings, runs (font/size/colour/highlight/b/i/u/s/sub/sup/caps), alignment, spacing,
 * indents, borders & shading, lists (per-list numbering, multilevel), tables (merged cells,
 * widths, borders, shading, valign, header rows), images (inline & floating), links (external +
 * internal bookmarks), page breaks, TOC, footnotes, comments, header/footer with page numbers,
 * watermark, page size/orientation/margins, page colour, page borders and columns. */
(function () {
  const D = window.Docs; const { $, $$, PX } = D;
  const TW = 15; // twips per CSS px
  const tw = (px) => Math.round(px * TW);

  const hex = (c) => {
    if (!c) return null;
    const cs = c.match(/color\(srgb\s+([\d.e-]+)\s+([\d.e-]+)\s+([\d.e-]+)(?:\s*\/\s*([\d.]+))?\)/);
    if (cs) { if (cs[4] !== undefined && +cs[4] < 0.05) return null; return [cs[1], cs[2], cs[3]].map(v => Math.round(Math.max(0, Math.min(1, +v)) * 255).toString(16).padStart(2, '0')).join('').toUpperCase(); }
    const m = c.match(/rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?\)/);
    if (!m) { if (/^#[0-9a-f]{6}$/i.test(c)) return c.slice(1).toUpperCase(); return null; }
    if (m[4] !== undefined && +m[4] < 0.05) return null;
    return [m[1], m[2], m[3]].map(v => Math.round(+v).toString(16).padStart(2, '0')).join('').toUpperCase();
  };
  const firstFamily = (ff) => (ff || '').split(',')[0].trim().replace(/^["']|["']$/g, '');
  const HIGHLIGHTS = { FFFF00: 'yellow', '00FF00': 'green', '00FFFF': 'cyan', FF00FF: 'magenta', '0000FF': 'blue', FF0000: 'red', '000080': 'darkBlue', '008080': 'darkCyan', '008000': 'darkGreen', '800080': 'darkMagenta', '800000': 'darkRed', '808000': 'darkYellow', '808080': 'darkGray', C0C0C0: 'lightGray', '000000': 'black' };

  /** Resolved formatting of each named style, used as the "base" so runs only carry direct formatting. */
  const SID = { Normal: 'normal', NoSpacing: 'nospace', Title: 'title', Subtitle: 'subtitle', Heading1: 'h1', Heading2: 'h2', Heading3: 'h3', Heading4: 'h4', Heading5: 'h5', Heading6: 'h6', Quote: 'quote', IntenseQuote: 'iquote', Caption: 'caption', Code: 'code', ListParagraph: 'listpara' };
  D.DOCX_SID = SID;
  const toBase = (r) => ({ font: r.font, size: r.size || 11, color: String(D.resolveColor(r.color) || '#111111').replace('#', '').toUpperCase(), bold: !!r.bold, italic: !!r.italic, caps: !!r.caps, smallcaps: !!r.smallcaps });
  function bases() {
    const out = { acc: D.accent().slice(1).toUpperCase(), dark: D.darkColor().slice(1).toUpperCase(), head: D.headFont(), body: D.bodyFont() };
    Object.entries(SID).forEach(([k, id]) => { out[k] = toBase(D.resolvedStyle(id)); });
    ['TOC1', 'TOC2', 'TOC3', 'Bibliography', 'TableofFigures'].forEach(k => out[k] = out.Normal);
    (D.design.customStyles || []).forEach(u => { if (u.type !== 'c') out['BTSU' + u.id] = toBase(D.resolvedStyle(u.id)); });
    return out;
  }

  function styleIdOf(el) {
    const t = el.tagName;
    if (/^H[1-6]$/.test(t)) return 'Heading' + t[1];
    if (t === 'PRE') return 'Code';
    const bq = el.closest('blockquote'); if (bq && D.rootOf(el)) return bq.classList.contains('intense') ? 'IntenseQuote' : 'Quote';
    if (el.classList.contains('bib-entry')) return 'Bibliography';
    if (el.classList.contains('listpara')) return 'ListParagraph';
    const us = [...el.classList].find(c => c.startsWith('us-')); if (us && D.findStyle(us.slice(3))) return 'BTSU' + us.slice(3);
    if (el.classList.contains('title')) return 'Title';
    if (el.classList.contains('subtitle')) return 'Subtitle';
    if (el.classList.contains('nospace')) return 'NoSpacing';
    if (el.classList.contains('caption')) return 'Caption';
    return 'Normal';
  }

  async function prepareImages(roots) {
    const map = new Map();
    for (const root of roots) for (const img of $$('img', root)) {
      let src = img.getAttribute('src') || ''; if (!src.startsWith('data:')) { try { src = await toDataURL(img); } catch { continue; } }
      if (!/^data:image\/(png|jpe?g|gif|bmp)/.test(src)) src = await D.toWordFriendly(src);
      const m = src.match(/^data:image\/(png|jpe?g|gif|bmp)/); if (!m) continue;
      map.set(img, { type: m[1] === 'jpeg' ? 'jpg' : m[1], data: BTS.dataURLToBytes(src) });
    }
    return map;
  }
  function toDataURL(img) { return new Promise((res, rej) => { const c = document.createElement('canvas'); c.width = img.naturalWidth; c.height = img.naturalHeight; try { c.getContext('2d').drawImage(img, 0, 0); res(c.toDataURL('image/png')); } catch (e) { rej(e); } }); }

  function watermarkPNG(w, Wpx, Hpx) {
    const c = document.createElement('canvas'); const scale = 2; c.width = Math.round(Wpx * scale); c.height = Math.round(Hpx * scale);
    const g = c.getContext('2d'); g.translate(c.width / 2, c.height / 2); if (w.layout !== 'horizontal') g.rotate(-Math.PI / 4);
    const size = Math.min(130, Math.max(28, (Wpx * (w.layout === 'horizontal' ? 0.8 : 1.0)) / Math.max(4, w.text.length) * 1.25)) * scale;
    g.font = `bold ${size}px ${D.fontStack(w.font || 'Calibri')}`; g.textAlign = 'center'; g.textBaseline = 'middle';
    g.globalAlpha = w.opacity ?? 0.35; g.fillStyle = w.color || '#c0c4cc'; g.fillText(w.text, 0, 0);
    return BTS.dataURLToBytes(c.toDataURL('image/png'));
  }

  D.exportDocx = async () => {
    const X = window.docx; if (!X) throw new Error('DOCX library failed to load');
    D.endPreview(); D.clearCellSel && D.clearCellSel();
    const oldZoom = D.zoom; if (oldZoom !== 100) D.setZoom(100);
    const mmPrev = D.mm && D.mm.preview; if (mmPrev) D.mmPreview(false);
    document.body.classList.add('exporting');
    try { return await build(X); } finally { document.body.classList.remove('exporting'); if (oldZoom !== 100) D.setZoom(oldZoom); if (mmPrev) D.mmPreview(true); }
  };

  async function build(X) {
    D.layout(); D.renderFootnotes();
    const B = bases();
    const ps = D.paperSize(); const m = D.page.m;
    const contentTw = Math.round((ps.w - m.l - m.r) * 1440);
    const images = await prepareImages([D.doc, D.hdr, D.ftr]);
    const numbering = []; let listSeq = 0;
    const footnotes = {}; let fnSeq = 0;
    const cids = [...new Set($$('.cmt', D.doc).map(s => s.dataset.cid))];
    const cidNum = new Map(cids.map((c, i) => [c, i]));
    const lastCmt = new Map(); $$('.cmt', D.doc).forEach(s => lastCmt.set(s.dataset.cid, s));
    const cmtStarted = new Set();

    const AL = { center: X.AlignmentType.CENTER, right: X.AlignmentType.RIGHT, end: X.AlignmentType.RIGHT, justify: X.AlignmentType.JUSTIFIED };
    const safeId = (id) => { let s = String(id).replace(/[^A-Za-z0-9_]/g, '_'); if (!/^[A-Za-z_]/.test(s)) s = '_' + s; return s.slice(0, 40); };

    /* ---------- runs ---------- */
    const CHAR_RSTYLE = { emphasis: 'Emphasis', strong: 'Strong', 'subtle-em': 'SubtleEmphasis', 'intense-em': 'IntenseEmphasis', 'subtle-ref': 'SubtleReference', 'intense-ref': 'IntenseReference', 'book-title': 'BookTitle' };
    const fontAttrs = (fam) => ({ ascii: fam, hAnsi: fam, eastAsia: fam, cs: D.csFontFor(fam) });
    function runProps(el, block, base, inLink) {
      const cs = getComputedStyle(el); const o = {};
      // character style (or citation) → run style; compare against base + style so only direct formatting is written
      let cst = null, isCite = false;
      for (let a = el; a && a !== block && !D.isRoot(a); a = a.parentElement) { if (!cst && a.classList) { const c = [...a.classList].find(x => x.startsWith('cs-')); if (c && D.findStyle(c.slice(3))) cst = c.slice(3); } if (a.classList && a.classList.contains('cite')) isCite = true; }
      if (cst) { const r = D.styleDef(cst); base = Object.assign({}, base); if (r.font) base.font = D.resolveFont(r.font); if (r.size) base.size = r.size; if (r.color) base.color = String(D.resolveColor(r.color)).replace('#', '').toUpperCase(); if (r.bold !== undefined) base.bold = !!r.bold; if (r.italic !== undefined) base.italic = !!r.italic; if (r.caps !== undefined) base.caps = !!r.caps; if (r.smallcaps !== undefined) base.smallcaps = !!r.smallcaps; }
      let deco = new Set(), bg = null, va = null, sizeEl = el;
      for (let a = el; a; a = a.parentElement) {
        const s = a === el ? cs : getComputedStyle(a);
        (s.textDecorationLine || '').split(' ').forEach(d => { if (d && d !== 'none') deco.add(d); });
        if (a === block || D.isRoot(a) || a.tagName === 'TD' || a.tagName === 'TH') break;
        if (!bg && !a.matches('.cmt,.eq,.fld,.fn-ref,a')) bg = hex(s.backgroundColor);
        if (!va && (s.verticalAlign === 'super' || s.verticalAlign === 'sub')) { va = s.verticalAlign; sizeEl = a.parentElement; }
      }
      const fam = firstFamily(cs.fontFamily); if (fam && fam.toLowerCase() !== String(base.font).toLowerCase()) o.font = fontAttrs(fam);
      const sz = Math.round(parseFloat(getComputedStyle(sizeEl).fontSize) * 0.75 * 2) / 2; if (sz && Math.abs(sz - base.size) > 0.01) o.size = Math.round(sz * 2);
      const bold = +cs.fontWeight >= 600; if (bold !== base.bold) o.bold = bold;
      const ital = cs.fontStyle === 'italic' || cs.fontStyle.startsWith('oblique'); if (ital !== base.italic) o.italics = ital;
      if (deco.has('underline')) o.underline = { type: X.UnderlineType.SINGLE };
      if (deco.has('line-through')) o.strike = true;
      const col = hex(cs.color); if (col && col !== base.color && !(inLink && col === '1F5FD6') && !(base.color === '111111' && col === '000000')) o.color = col;
      if (bg) { if (HIGHLIGHTS[bg]) o.highlight = HIGHLIGHTS[bg]; else o.shading = { type: X.ShadingType.CLEAR, fill: bg, color: 'auto' }; }
      if (va === 'super') o.superScript = true; else if (va === 'sub') o.subScript = true;
      if ((cs.textTransform === 'uppercase') !== !!base.caps) o.allCaps = cs.textTransform === 'uppercase';
      const sc = cs.fontVariantCaps === 'small-caps' || cs.fontVariant === 'small-caps'; if (sc !== !!base.smallcaps) o.smallCaps = sc;
      if (cst) o.style = CHAR_RSTYLE[cst] || 'BTSU' + cst;
      if (isCite) o.style = 'BTSCitation';
      if (inLink) o.style = 'Hyperlink';
      return o;
    }
    let curTC = null, tcSeq = 9001;
    const tcDate = (t) => { const d = new Date(t || Date.now()); return (isNaN(d) ? new Date() : d).toISOString().replace(/\.\d+Z$/, 'Z'); };
    function mkRun(o) {
      if (!curTC) return new X.TextRun(o);
      const opts = { ...o, id: tcSeq++, author: curTC.author || 'Unknown', date: tcDate(curTC.time) };
      return curTC.kind === 'ins' ? new X.InsertedTextRun(opts) : new X.DeletedTextRun(opts);
    }
    /** w:fldSimple whose cached result keeps the surrounding run formatting. */
    function field(instr, text, props) { const f = new X.SimpleField(instr); f.root.push(new X.TextRun({ ...props, text })); return f; }
    function textRuns(text, props, pre) {
      text = text.replace(/​/g, '');
      if (!pre) text = text.replace(/ /g, ' ');
      if (!text) return [];
      if (D.hasDevanagari(text)) props = { ...props, language: { value: 'en-IN', bidirectional: 'hi-IN' } };
      const out = [];
      const lines = pre ? text.split('\n') : [text];
      lines.forEach((ln, li) => {
        const parts = ln.split('\t'); const children = [];
        parts.forEach((p, i) => { if (i) children.push(new X.Tab()); if (p) children.push(p); });
        if (li > 0) out.push(mkRun({ ...props, break: 1 }));
        if (children.length) out.push(mkRun({ ...props, children }));
      });
      return out;
    }
    function imageRun(img, base, hf) {
      const im = images.get(img); if (!im) return null;
      let w = parseFloat(img.style.width) || img.offsetWidth || img.naturalWidth || 100;
      let h = parseFloat(img.style.height) || img.offsetHeight || img.naturalHeight || 100;
      if (img.style.height === 'auto' || !h) h = img.offsetHeight || h;
      const maxW = contentTw / TW; if (w > maxW) { h = h * maxW / w; w = maxW; }
      const opts = { type: im.type, data: im.data, transformation: { width: Math.round(w), height: Math.round(h) }, altText: { name: 'Picture', description: img.alt || '', title: img.alt || '' } };
      const side = img.classList.contains('img-left') ? 'left' : img.classList.contains('img-right') ? 'right' : null;
      if (side) opts.floating = { horizontalPosition: { relative: X.HorizontalPositionRelativeFrom.COLUMN, align: side === 'left' ? X.HorizontalPositionAlign.LEFT : X.HorizontalPositionAlign.RIGHT }, verticalPosition: { relative: X.VerticalPositionRelativeFrom.PARAGRAPH, offset: 0 }, wrap: { type: X.TextWrappingType.SQUARE, side: X.TextWrappingSide.BOTH_SIDES }, margins: { left: 114300, right: 114300, top: 38100, bottom: 38100 } };
      if (img.dataset.decorative) opts.altText = { name: 'Decorative picture', description: '', title: '' };
      if (img.classList.contains('img-border')) opts.outline = { type: 'solidFill', solidFillType: 'rgb', value: '3C4250', width: 9525 };
      return new X.ImageRun(opts);
    }
    function inlineChildren(block, base, ctx) {
      const out = []; const pre = block.tagName === 'PRE' || getComputedStyle(block).whiteSpace.startsWith('pre');
      const walk = (node, sink, inLink) => {
        for (const n of node.childNodes) {
          if (n.nodeType === 3) { sink.push(...textRuns(pre ? n.data : n.data.replace(/[\n\r]+/g, ' '), runProps(n.parentElement, block, base, inLink), pre)); continue; }
          if (n.nodeType !== 1) continue;
          const t = n.tagName;
          if (t === 'UL' || t === 'OL' || (ctx.skipBlocks && (D.BLOCK_TAGS.has(t) || t === 'TABLE'))) continue;
          if (t === 'BR') { if (n.nextSibling || n.previousSibling) { const nx = n.nextSibling; if (!(nx == null || (nx.nodeType === 3 && !nx.data.trim() && !nx.nextSibling))) sink.push(new X.TextRun({ break: 1 })); } continue; }
          if (t === 'IMG') { const r = imageRun(n, base, ctx.hf); if (r) sink.push(r); continue; }
          if (n.classList.contains('fn-ref')) { fnSeq++; footnotes[fnSeq] = { children: [new X.Paragraph({ children: [new X.TextRun({ text: n.dataset.text || '', size: 18 })] })] }; sink.push(new X.FootnoteReferenceRun(fnSeq)); continue; }
          if (n.classList.contains('tc-ins') || n.classList.contains('tc-del')) {
            const prev = curTC; curTC = { kind: n.classList.contains('tc-ins') ? 'ins' : 'del', author: n.dataset.author, time: n.dataset.time };
            if (prev && prev.kind === 'del') curTC = prev; // a deletion inside an insertion stays deleted
            walk(n, sink, inLink); curTC = prev; continue;
          }
          if (n.classList.contains('seq')) { if (curTC) { sink.push(...textRuns(n.textContent, runProps(n, block, base, false))); continue; } sink.push(field(`SEQ ${(n.dataset.seq || 'Figure').replace(/[^A-Za-z]/g, '')} \\* ARABIC`, n.textContent || '1', runProps(n, block, base, false))); continue; }
          if (n.classList.contains('mf')) { if (curTC) { sink.push(...textRuns('«' + n.dataset.field + '»', runProps(n, block, base, false))); continue; } sink.push(field(` MERGEFIELD ${/\s/.test(n.dataset.field || '') ? '"' + n.dataset.field.replace(/"/g, '') + '"' : (n.dataset.field || 'Field')} `, '«' + (n.dataset.field || 'Field') + '»', runProps(n, block, base, false))); continue; }
          if (n.classList.contains('fld')) { const p = runProps(n, block, base, false); sink.push(new X.TextRun({ ...p, children: [n.dataset.fld === 'pages' ? X.PageNumber.TOTAL_PAGES : X.PageNumber.CURRENT] })); continue; }
          if (n.classList.contains('tab')) { sink.push(new X.TextRun({ ...runProps(n, block, base, inLink), children: [new X.Tab()] })); continue; }
          if (n.classList.contains('frac')) { const [a, b] = n.children; const p = runProps(n, block, base, inLink); const wrap = (s) => s.length > 1 ? '(' + s + ')' : s; sink.push(new X.TextRun({ ...p, text: wrap(a ? a.textContent : '') + '/' + wrap(b ? b.textContent : '') })); continue; }
          if (t === 'A' && n.getAttribute('href') && !inLink) {
            const kids = []; walk(n, kids, true);
            if (!kids.length) continue;
            const href = n.getAttribute('href');
            if (href.startsWith('#')) sink.push(new X.InternalHyperlink({ anchor: safeId(href.slice(1)), children: kids }));
            else sink.push(new X.ExternalHyperlink({ link: href, children: kids }));
            continue;
          }
          const cid = !ctx.hf && n.classList.contains('cmt') ? n.dataset.cid : null;
          if (cid && cidNum.has(cid) && !cmtStarted.has(cid)) { cmtStarted.add(cid); sink.push(new X.CommentRangeStart(cidNum.get(cid))); }
          walk(n, sink, inLink);
          if (cid && cidNum.has(cid) && lastCmt.get(cid) === n) { const id = cidNum.get(cid); sink.push(new X.CommentRangeEnd(id)); sink.push(new X.TextRun({ children: [new X.CommentReference(id)] })); }
        }
      };
      walk(block, out, false);
      return out;
    }

    /* ---------- paragraphs ---------- */
    const HEADINGS = { Heading1: X.HeadingLevel.HEADING_1, Heading2: X.HeadingLevel.HEADING_2, Heading3: X.HeadingLevel.HEADING_3, Heading4: X.HeadingLevel.HEADING_4, Heading5: X.HeadingLevel.HEADING_5, Heading6: X.HeadingLevel.HEADING_6, Title: X.HeadingLevel.TITLE };
    function borderOf(cs, side) {
      const st = cs['border' + side + 'Style'], w = parseFloat(cs['border' + side + 'Width']) || 0, col = hex(cs['border' + side + 'Color']);
      if (!st || st === 'none' || st === 'hidden' || !w || !col) return null;
      return { style: st === 'double' ? X.BorderStyle.DOUBLE : st === 'dashed' ? X.BorderStyle.DASHED : st === 'dotted' ? X.BorderStyle.DOTTED : X.BorderStyle.SINGLE, size: Math.max(2, Math.round(w * 6)), color: col, space: 2 };
    }
    function paragraph(el, ctx, extra = {}) {
      const sid = extra.styleId || styleIdOf(el); const base = B[sid] || B.Normal;
      const cs = getComputedStyle(el);
      const o = { children: inlineChildren(el, base, ctx) };
      if (HEADINGS[sid]) o.heading = HEADINGS[sid]; else if (sid !== 'Normal') o.style = sid;
      const sdef = SID[sid] ? D.resolvedStyle(SID[sid]) : (/^BTSU/.test(sid) ? D.resolvedStyle(sid.slice(4)) : D.resolvedStyle('normal'));
      const direct = !!(el.style.marginTop || el.style.marginBottom || el.style.lineHeight || el.closest('td,th,li,.textbox,.cols') || ctx.hf || el.closest('.biblio,.toc'));
      const al = AL[cs.textAlign]; const styleAl = sdef.align || 'left'; const curAl = cs.textAlign.replace('start', 'left').replace(/^end$/, 'right');
      if (el.style.textAlign || curAl !== styleAl) o.alignment = al || X.AlignmentType.LEFT;
      // spacing (emulate CSS margin collapsing between siblings); paragraphs without direct spacing use the style's
      const fs = parseFloat(cs.fontSize) || 14.67;
      const before = parseFloat(cs.marginTop) || 0, after = parseFloat(cs.marginBottom) || 0;
      if (direct || ctx.prevAfterDirect) {
        const spacing = { before: tw(Math.max(0, before - (ctx.prevAfter || 0))), after: tw(after) };
        const lh = parseFloat(cs.lineHeight); if (lh && cs.lineHeight !== 'normal') { spacing.line = Math.round(lh / fs * 240); spacing.lineRule = X.LineRuleType.AUTO; }
        o.spacing = spacing;
      }
      ctx.prevAfter = after; ctx.prevAfterDirect = direct && !!el.style.marginBottom;
      // indents relative to the container's content box
      const styleIndentOnly = !el.style.marginLeft && !el.style.textIndent && !el.style.paddingLeft && !el.closest('td,th,.textbox') && !ctx.hf && (sid === 'Quote' || sid === 'IntenseQuote' || sid === 'ListParagraph' || sid === 'Bibliography' || /^BTSU/.test(sid) || sid === 'Code' || ((sdef.indent || sdef.firstLine) && !!ctx.box));
      if (!extra.numbering && ctx.box && !styleIndentOnly) {
        const r = el.getBoundingClientRect();
        let left = r.left - ctx.box.left, right = ctx.box.right - r.right;
        if (ctx.inCols) { left = parseFloat(cs.marginLeft) || 0; right = parseFloat(cs.marginRight) || 0; }
        const ti = parseFloat(cs.textIndent) || 0; const ind = {};
        if (left > 1) ind.left = tw(left + (parseFloat(cs.paddingLeft) || 0) + (parseFloat(cs.borderLeftWidth) || 0)); else if (parseFloat(cs.paddingLeft) > 6) ind.left = tw(parseFloat(cs.paddingLeft));
        if (right > 1 && !el.closest('td,th')) ind.right = tw(right);
        if (ti > 0) ind.firstLine = tw(ti); else if (ti < 0) ind.hanging = tw(-ti);
        if (extra.indentLeft) ind.left = (ind.left || 0) + extra.indentLeft;
        if (Object.keys(ind).length) o.indent = ind;
      }
      if (extra.numbering) o.numbering = extra.numbering;
      // borders & shading (paragraph itself, or enclosing blockquote / pre)
      if (el.tagName !== 'TD' && el.tagName !== 'TH' && !((sid === 'Quote' || sid === 'IntenseQuote' || sid === 'Code') && !el.style.borderLeft && !el.style.backgroundColor)) {
        const bq = el.closest('blockquote'); const src = el.tagName === 'PRE' ? cs : (bq && D.rootOf(bq) ? getComputedStyle(bq) : cs);
        const bd = {}; ['Top', 'Bottom', 'Left', 'Right'].forEach(s => { const b = borderOf(src === cs ? cs : src, s); if (b && (src === cs || s === 'Left')) bd[s.toLowerCase()] = b; });
        if (Object.keys(bd).length) o.border = bd;
        const bg = hex(src.backgroundColor); if (bg && bg !== 'FFFFFF') o.shading = { type: X.ShadingType.CLEAR, fill: bg, color: 'auto' };
      }
      if (ctx.pendingBreak) { o.pageBreakBefore = true; ctx.pendingBreak = false; }
      if (/^Heading/.test(sid)) { o.keepNext = true; o.keepLines = true; }
      if (el.id) o.children = [new X.Bookmark({ id: safeId(el.id), children: o.children })];
      if (extra.tabStops) o.tabStops = extra.tabStops;
      return new X.Paragraph(o);
    }

    /* ---------- lists ---------- */
    const FMT = { decimal: [X.LevelFormat.DECIMAL, '%N.'], 'decimal-leading-zero': [X.LevelFormat.DECIMAL_ZERO, '%N.'], 'lower-alpha': [X.LevelFormat.LOWER_LETTER, '%N.'], 'lower-latin': [X.LevelFormat.LOWER_LETTER, '%N.'], 'upper-alpha': [X.LevelFormat.UPPER_LETTER, '%N.'], 'upper-latin': [X.LevelFormat.UPPER_LETTER, '%N.'], 'lower-roman': [X.LevelFormat.LOWER_ROMAN, '%N.'], 'upper-roman': [X.LevelFormat.UPPER_ROMAN, '%N.'], disc: [X.LevelFormat.BULLET, '•'], circle: [X.LevelFormat.BULLET, '◦'], square: [X.LevelFormat.BULLET, '▪'], none: [X.LevelFormat.NONE, ''] };
    const ORD_CYCLE = ['decimal', 'lower-alpha', 'lower-roman'], BUL_CYCLE = ['disc', 'circle', 'square'];
    function listDef(top) {
      const fmts = []; const starts = [];
      const scan = (list, d) => { if (d > 8) return; if (!fmts[d]) { fmts[d] = getComputedStyle(list).listStyleType; starts[d] = list.tagName === 'OL' ? (+list.getAttribute('start') || 1) : 1; } [...list.children].forEach(li => [...li.children].forEach(c => { if (c.tagName === 'UL' || c.tagName === 'OL') scan(c, d + 1); })); };
      scan(top, 0);
      const ref = 'list' + (++listSeq); const ordered = top.tagName === 'OL';
      const levels = [];
      for (let i = 0; i < 9; i++) {
        let f = fmts[i] || (ordered ? ORD_CYCLE[i % 3] : BUL_CYCLE[i % 3]); if (!FMT[f]) f = /OL/.test(top.tagName) ? 'decimal' : 'disc';
        const [format, text] = FMT[f];
        levels.push({ level: i, format, text: text.replace('N', i + 1), alignment: X.AlignmentType.LEFT, start: starts[i] || 1, style: { paragraph: { indent: { left: 360 + i * 360 + 180, hanging: 270 } }, ...(format === X.LevelFormat.BULLET ? { run: { font: 'Calibri' } } : {}) } });
      }
      numbering.push({ reference: ref, levels });
      return ref;
    }
    function listBlocks(list, level, ref, ctx, out) {
      for (const li of list.children) {
        if (li.tagName !== 'LI') continue;
        const blocks = [...li.children].filter(c => D.BLOCK_TAGS.has(c.tagName) && c.tagName !== 'LI');
        const hasInline = [...li.childNodes].some(n => (n.nodeType === 3 && n.data.trim()) || (n.nodeType === 1 && !D.BLOCK_TAGS.has(n.tagName) && !['UL', 'OL', 'TABLE'].includes(n.tagName)));
        if (hasInline || !blocks.length) out.push(paragraph(li, { ...ctx, skipBlocks: true }, { numbering: { reference: ref, level: Math.min(8, level) }, styleId: 'ListParagraph' }));
        let first = !(hasInline || !blocks.length);
        for (const c of li.children) {
          if (c.tagName === 'UL' || c.tagName === 'OL') listBlocks(c, level + 1, ref, ctx, out);
          else if (c.tagName === 'TABLE') out.push(table(c, ctx), new X.Paragraph({ spacing: { after: 0 } }));
          else if (D.BLOCK_TAGS.has(c.tagName)) { out.push(paragraph(c, ctx, first ? { numbering: { reference: ref, level: Math.min(8, level) } } : { indentLeft: 540 + level * 360 })); first = false; }
        }
        ctx.prevAfter = 0;
      }
    }

    /* ---------- tables ---------- */
    function table(t, ctx) {
      const g = D.tableGrid(t); if (!g.rows || !g.cols) return new X.Paragraph('');
      const tr = t.getBoundingClientRect(); const totalPx = Math.min(tr.width, (ctx.availTw || contentTw) / TW);
      // column widths: prefer colgroup, fall back to measured single-span cells
      const cols = [...(t.querySelector(':scope > colgroup')?.children || [])];
      let wpx = new Array(g.cols).fill(0);
      for (let c = 0; c < g.cols; c++) for (let r = 0; r < g.rows; r++) { const x = g.g[r][c]; if (x && (x.colSpan || 1) === 1) { wpx[c] = x.getBoundingClientRect().width; break; } }
      if (cols.length === g.cols) cols.forEach((col, i) => { const s = col.style.width; if (s.endsWith('%')) wpx[i] = totalPx * parseFloat(s) / 100; else if (D.toPx(s)) wpx[i] = D.toPx(s); });
      const missing = wpx.filter(w => !w).length; if (missing) { const rest = Math.max(20 * missing, totalPx - wpx.reduce((a, b) => a + b, 0)); wpx = wpx.map(w => w || rest / missing); }
      const sum = wpx.reduce((a, b) => a + b, 0); const k = sum > totalPx ? totalPx / sum : 1;
      const colTw = wpx.map(w => Math.max(120, tw(w * k)));
      const NONE = { style: X.BorderStyle.NIL, size: 0, color: 'auto' };
      const rows = [...t.rows].map((row, ri) => {
        const cells = [...row.cells].map(cell => {
          const cs = getComputedStyle(cell); const p = (() => { for (let c = 0; c < g.cols; c++) if (g.g[ri][c] === cell) return c; return 0; })();
          const span = cell.colSpan || 1; const wTw = colTw.slice(p, p + span).reduce((a, b) => a + b, 0);
          const inner = blocksOf(cell, { ...ctx, box: contentBox(cell), availTw: wTw, prevAfter: 0, pendingBreak: false });
          if (!inner.length || !(inner[inner.length - 1] instanceof X.Paragraph)) inner.push(new X.Paragraph({ spacing: { after: 0 } }));
          const o = { children: inner, width: { size: wTw, type: X.WidthType.DXA } };
          if (span > 1) o.columnSpan = span; if ((cell.rowSpan || 1) > 1) o.rowSpan = cell.rowSpan;
          const bg = hex(cs.backgroundColor); if (bg) o.shading = { type: X.ShadingType.CLEAR, fill: bg, color: 'auto' };
          const va = cs.verticalAlign; if (va === 'middle') o.verticalAlign = X.VerticalAlignTable.CENTER; else if (va === 'bottom') o.verticalAlign = X.VerticalAlignTable.BOTTOM;
          o.borders = { top: borderOf(cs, 'Top') || NONE, bottom: borderOf(cs, 'Bottom') || NONE, left: borderOf(cs, 'Left') || NONE, right: borderOf(cs, 'Right') || NONE };
          ['top', 'bottom', 'left', 'right'].forEach(s => { if (o.borders[s].space) delete o.borders[s].space; });
          const pl = parseFloat(cs.paddingLeft) || 0, pt = parseFloat(cs.paddingTop) || 0;
          o.margins = { left: tw(pl), right: tw(parseFloat(cs.paddingRight) || 0), top: tw(pt), bottom: tw(parseFloat(cs.paddingBottom) || 0) };
          return new X.TableCell(o);
        });
        const header = ri === 0 && (/t-(header|banded|dark|grid-accent|light|minimal|hrow)/.test(t.className) || [...row.cells].every(c => c.tagName === 'TH')) && !t.classList.contains('t-noheader');
        return new X.TableRow({ children: cells, tableHeader: header || undefined, cantSplit: true });
      });
      const tbl = { rows, columnWidths: colTw, width: { size: colTw.reduce((a, b) => a + b, 0), type: X.WidthType.DXA }, layout: X.TableLayoutType.FIXED, borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE } };
      const ml = parseFloat(getComputedStyle(t).marginLeft) || 0; if (ml > 1) tbl.indent = { size: tw(ml), type: X.WidthType.DXA };
      if (getComputedStyle(t).marginLeft === getComputedStyle(t).marginRight && ml > 1) { delete tbl.indent; tbl.alignment = X.AlignmentType.CENTER; }
      ctx.prevAfter = parseFloat(getComputedStyle(t).marginBottom) || 0;
      return new X.Table(tbl);
    }
    function contentBox(el) { const r = el.getBoundingClientRect(); const cs = getComputedStyle(el); return { left: r.left + (parseFloat(cs.paddingLeft) || 0) + (parseFloat(cs.borderLeftWidth) || 0), right: r.right - (parseFloat(cs.paddingRight) || 0) - (parseFloat(cs.borderRightWidth) || 0) }; }

    /* ---------- TOC ---------- */
    function tocBlocks(toc, ctx) {
      const out = []; const title = toc.querySelector('.toc-title'); const tof = toc.classList.contains('tof');
      if (title) out.push(new X.Paragraph({ style: 'TOCHeading', children: [new X.TextRun(title.textContent)] }));
      $$('p[data-target]', toc).forEach(p => {
        const lvl = +((p.className.match(/toc-l(\d)/) || [, 1])[1]);
        const text = p.querySelector('.toc-t').textContent, num = p.querySelector('.toc-n').textContent;
        out.push(new X.Paragraph({ style: tof ? 'TableofFigures' : 'TOC' + lvl, tabStops: [{ type: X.TabStopType.RIGHT, position: contentTw - 10, leader: X.LeaderType.DOT }], children: [new X.InternalHyperlink({ anchor: safeId(p.dataset.target), children: [new X.TextRun(text), new X.TextRun({ children: [new X.Tab(), num] })] })] }));
      });
      if (!out.length) out.push(new X.Paragraph(''));
      ctx.prevAfter = 8;
      return out;
    }

    /* ---------- block walker ---------- */
    function blocksOf(container, ctx) {
      const out = [];
      ctx.box = ctx.box || contentBox(container);
      for (const n of container.childNodes) {
        if (n.nodeType === 3) { if (n.data.trim()) out.push(new X.Paragraph({ children: textRuns(n.data, {}) })); continue; }
        if (n.nodeType !== 1) continue;
        const t = n.tagName;
        if (n.id === 'fn-area') continue;
        if (t === 'HR') {
          if (n.classList.contains('page-break')) { if (ctx.pendingBreak) out.push(new X.Paragraph({ pageBreakBefore: true })); ctx.pendingBreak = true; }
          else out.push(new X.Paragraph({ border: { bottom: { style: X.BorderStyle.SINGLE, size: 6, color: '9AA1AD', space: 1 } }, spacing: { before: 120, after: 120 } }));
          continue;
        }
        if (n.classList.contains('toc')) { out.push(...tocBlocks(n, ctx)); continue; }
        if (t === 'UL' || t === 'OL') { const ref = listDef(n); const before = out.length; listBlocks(n, 0, ref, ctx, out); if (ctx.pendingBreak && out.length > before) ctx.pendingBreak = false; continue; }
        if (t === 'TABLE') { if (ctx.pendingBreak) { out.push(new X.Paragraph({ children: [new X.PageBreak()], spacing: { after: 0 } })); ctx.pendingBreak = false; } out.push(table(n, ctx)); continue; }
        if (n.classList.contains('textbox')) {
          const w = n.getBoundingClientRect().width; const cs = getComputedStyle(n); const b = borderOf(cs, 'Top') || { style: X.BorderStyle.SINGLE, size: 6, color: '5D6575' };
          const inner = blocksOf(n, { ...ctx, box: contentBox(n), prevAfter: 0, pendingBreak: false }); if (!inner.length) inner.push(new X.Paragraph(''));
          const fl = n.classList.contains('tb-inline') ? null : (n.classList.contains('tb-left') ? 'left' : 'right');
          const tblOpts = { rows: [new X.TableRow({ children: [new X.TableCell({ children: inner, width: { size: tw(w), type: X.WidthType.DXA }, margins: { left: 120, right: 120, top: 80, bottom: 80 }, borders: { top: b, bottom: b, left: b, right: b } })] })], width: { size: tw(w), type: X.WidthType.DXA }, columnWidths: [tw(w)], layout: X.TableLayoutType.FIXED };
          if (fl) tblOpts.float = { horizontalAnchor: X.TableAnchorType.MARGIN, relativeHorizontalPosition: fl === 'left' ? X.RelativeHorizontalPosition.LEFT : X.RelativeHorizontalPosition.RIGHT, verticalAnchor: X.TableAnchorType.TEXT, leftFromText: 180, rightFromText: 180, topFromText: 60, bottomFromText: 60 };
          out.push(new X.Table(tblOpts)); continue;
        }
        if ((t === 'DIV' || t === 'BLOCKQUOTE' || t === 'SECTION') && D.hasBlockChild(n)) { out.push(...blocksOf(n, { ...ctx, box: ctx.box, prevAfter: ctx.prevAfter })); ctx.pendingBreak = false; continue; }
        if (D.BLOCK_TAGS.has(t) || t === 'IMG' || !D.BLOCK_TAGS.has(t)) {
          if (t === 'IMG') { const wrap = document.createElement('span'); /* image directly in container */ const r = imageRun(n, B.Normal); if (r) out.push(new X.Paragraph({ children: [r], alignment: n.classList.contains('img-center') ? X.AlignmentType.CENTER : undefined })); continue; }
          const p = paragraph(n, ctx);
          if (n.querySelector(':scope > img.img-center')) p.properties && 0; // alignment handled via CSS text-align below
          out.push(p);
        }
      }
      return out;
    }

    /* ---------- sections (columns) ---------- */
    const segs = []; let cur = { cols: 1, nodes: [] };
    for (const n of D.doc.childNodes) {
      if (n.nodeType === 1 && n.classList.contains('cols')) { if (cur.nodes.length) segs.push(cur); const k = n.classList.contains('cols-3') ? 3 : 2; segs.push({ cols: k, nodes: [n], wrapper: true }); cur = { cols: 1, nodes: [] }; }
      else cur.nodes.push(n);
    }
    if (cur.nodes.length || !segs.length) segs.push(cur);

    // centered images in their own paragraph: fix alignment by inspecting class after build (handled in paragraph via CSS: img-center is display:block margin auto → mark)
    $$('img.img-center', D.doc).forEach(img => { const b = D.blockOf(img); if (b && !b.style.textAlign) { b.dataset.tmpAlign = '1'; b.style.textAlign = 'center'; } });

    const ctx0 = { prevAfter: 0, pendingBreak: false };
    const sectionsChildren = segs.map(seg => {
      const holder = seg.wrapper ? seg.nodes[0] : { childNodes: seg.nodes };
      return blocksOf(holder, Object.assign(ctx0, { box: contentBox(D.doc), inCols: !!seg.wrapper }));
    });
    $$('[data-tmp-align]', D.doc).forEach(b => { b.style.textAlign = ''; b.removeAttribute('data-tmp-align'); if (!b.getAttribute('style')) b.removeAttribute('style'); });
    if (ctx0.pendingBreak) sectionsChildren[sectionsChildren.length - 1].push(new X.Paragraph({ children: [new X.PageBreak()] }));
    // footnote area in editor isn't exported as text (real footnotes are)

    /* ---------- header / footer ---------- */
    const hfCtx = () => ({ prevAfter: 0, pendingBreak: false, hf: true });
    const hasHdr = D.hdr.textContent.trim() || D.hdr.querySelector('img'); const hasFtr = D.ftr.textContent.trim() || D.ftr.querySelector('img');
    const W = ps.w * PX, H = ps.h * PX;
    let hdrKids = hasHdr ? blocksOf(D.hdr, { ...hfCtx(), box: contentBox(D.hdr) }) : [];
    if (D.design.watermark) {
      const wmW = W * 0.9, wmH = H * 0.6;
      const wm = new X.ImageRun({ type: 'png', data: watermarkPNG(D.design.watermark, wmW, wmH), transformation: { width: Math.round(wmW), height: Math.round(wmH) }, altText: { name: 'BTS Watermark', description: JSON.stringify(D.design.watermark), title: 'Watermark' }, floating: { horizontalPosition: { relative: X.HorizontalPositionRelativeFrom.PAGE, align: X.HorizontalPositionAlign.CENTER }, verticalPosition: { relative: X.VerticalPositionRelativeFrom.PAGE, align: X.VerticalPositionAlign.CENTER }, behindDocument: true, allowOverlap: true, wrap: { type: X.TextWrappingType.NONE } } });
      hdrKids = [new X.Paragraph({ children: [wm], spacing: { after: 0 } }), ...hdrKids];
    }
    const headers = hdrKids.length ? { default: new X.Header({ children: hdrKids }) } : undefined;
    const footers = hasFtr ? { default: new X.Footer({ children: blocksOf(D.ftr, { ...hfCtx(), box: contentBox(D.ftr) }) }) } : undefined;
    if (D.design.hdrFirst) { if (headers) headers.first = new X.Header({ children: D.design.watermark ? [hdrKids[0]] : [new X.Paragraph('')] }); if (footers) footers.first = new X.Footer({ children: [new X.Paragraph('')] }); }

    /* ---------- page setup ---------- */
    const P = D.PAPER[D.page.size] || D.PAPER.A4;
    const pageProps = {
      size: { width: Math.round(P.w * 1440), height: Math.round(P.h * 1440), orientation: D.page.orient === 'landscape' ? X.PageOrientation.LANDSCAPE : X.PageOrientation.PORTRAIT },
      margin: { top: Math.round(m.t * 1440), bottom: Math.round(m.b * 1440), left: Math.round(m.l * 1440), right: Math.round(m.r * 1440), header: Math.round(D.page.hdrDist * 1440), footer: Math.round(D.page.ftrDist * 1440), gutter: 0 },
    };
    if (D.design.border) {
      const bd = D.design.border; const style = bd.style === 'double' ? X.BorderStyle.DOUBLE : bd.style === 'dashed' ? X.BorderStyle.DASHED : bd.style === 'dotted' ? X.BorderStyle.DOTTED : bd.style === 'art' ? X.BorderStyle.THICK_THIN_SMALL_GAP : X.BorderStyle.SINGLE;
      const b = { style, size: Math.max(4, Math.min(96, Math.round((bd.width || 1.5) * 6))), color: (bd.color || '#1f2533').slice(1), space: 24 };
      pageProps.borders = { pageBorders: { display: X.PageBorderDisplay.ALL_PAGES, offsetFrom: X.PageBorderOffsetFrom.PAGE, zOrder: X.PageBorderZOrder.FRONT }, pageBorderTop: b, pageBorderBottom: b, pageBorderLeft: b, pageBorderRight: b };
    }
    const sections = segs.map((seg, i) => ({
      properties: { page: pageProps, ...(i > 0 ? { type: X.SectionType.CONTINUOUS } : {}), ...(seg.cols > 1 ? { column: { count: seg.cols, space: 576, equalWidth: true } } : {}), ...(i === 0 && D.design.hdrFirst ? { titlePage: true } : {}) },
      headers: i === 0 ? headers : undefined, footers: i === 0 ? footers : undefined,
      children: sectionsChildren[i].length ? sectionsChildren[i] : [new X.Paragraph('')],
    }));

    /* ---------- styles (real named styles in w:styles) ---------- */
    const ps2 = D.design.pspacing || { after: 8, line: 1.15, before: 0 };
    const RS = (id) => D.resolvedStyle(id);
    const normalAfter = RS('normal').after || 0;
    const runOf = (r, isChar) => {
      const o = {};
      if (r.font) { const f = D.resolveFont(r.font); o.font = { ascii: f, hAnsi: f, eastAsia: f, cs: D.csFontFor(f) }; }
      if (r.size) o.size = Math.round(r.size * 2);
      if (r.color) o.color = String(D.resolveColor(r.color)).replace('#', '').toUpperCase();
      if (r.bold !== undefined || !isChar) o.bold = !!r.bold;
      if (r.italic !== undefined || !isChar) o.italics = !!r.italic;
      if (r.underline) o.underline = { type: X.UnderlineType.SINGLE };
      if (r.caps) o.allCaps = true; if (r.smallcaps) o.smallCaps = true;
      if (r.letter) o.characterSpacing = Math.round(r.letter * 20);
      return o;
    };
    const paraOf = (r, id) => {
      const o = { spacing: { before: Math.round(Math.max(0, (r.before || 0) - (id === 'normal' ? 0 : normalAfter)) * 20), after: Math.round((r.after || 0) * 20), line: Math.round((r.line || 1) * 240), lineRule: X.LineRuleType.AUTO } };
      if (r.align && r.align !== 'left') o.alignment = AL[r.align] || X.AlignmentType.LEFT;
      const ind = {}; if (r.indent) ind.left = Math.round(r.indent / 2.54 * 1440); if (r.firstLine > 0) ind.firstLine = Math.round(r.firstLine / 2.54 * 1440); else if (r.firstLine < 0) ind.hanging = Math.round(-r.firstLine / 2.54 * 1440); if (Object.keys(ind).length) o.indent = ind;
      if (/^h[1-6]$|^title$/.test(id)) { o.keepNext = true; o.keepLines = true; }
      return o;
    };
    const full = (id) => { const r = RS(id); return { run: runOf(r), paragraph: paraOf(r, id) }; };
    const nb = RS('normal');
    const accHex = D.accent().slice(1).toUpperCase();
    const pstyle = (docxId, name, id, extra = {}) => { const f = full(id); return { id: docxId, name, basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { ...f.run, ...(extra.run || {}) }, paragraph: { ...f.paragraph, ...(extra.paragraph || {}) } }; };
    const styles = {
      default: {
        document: { run: { font: { ascii: nb.font, hAnsi: nb.font, eastAsia: nb.font, cs: D.csFontFor(nb.font) }, size: Math.round((nb.size || 11) * 2), color: String(nb.color).replace('#', '').toUpperCase() }, paragraph: { spacing: { after: Math.round(ps2.after * 20), before: Math.round((ps2.before || 0) * 20), line: Math.round(ps2.line * 240), lineRule: X.LineRuleType.AUTO } } },
        title: full('title'), heading1: full('h1'), heading2: full('h2'), heading3: full('h3'), heading4: full('h4'), heading5: full('h5'), heading6: full('h6'),
        hyperlink: { run: { color: '1F5FD6', underline: { type: X.UnderlineType.SINGLE } } },
        listParagraph: { run: {}, paragraph: { spacing: { after: 40 }, indent: { left: 720 }, contextualSpacing: true } },
      },
      paragraphStyles: [
        pstyle('Subtitle', 'Subtitle', 'subtitle'),
        pstyle('Quote', 'Quote', 'quote', { paragraph: { indent: { left: 432, right: 576 }, border: { left: { style: X.BorderStyle.SINGLE, size: 18, color: accHex, space: 8 } } } }),
        pstyle('IntenseQuote', 'Intense Quote', 'iquote', { paragraph: { indent: { left: 490, right: 490 }, border: { top: { style: X.BorderStyle.SINGLE, size: 6, color: accHex, space: 8 }, bottom: { style: X.BorderStyle.SINGLE, size: 6, color: accHex, space: 8 } } } }),
        { id: 'NoSpacing', name: 'No Spacing', basedOn: 'Normal', quickFormat: true, paragraph: { spacing: { after: 0, before: 0, line: 240, lineRule: X.LineRuleType.AUTO } } },
        pstyle('Caption', 'caption', 'caption'),
        pstyle('Code', 'Code', 'code', { paragraph: { shading: { type: X.ShadingType.CLEAR, fill: 'F3F4F7', color: 'auto' } } }),
        { id: 'TOCHeading', name: 'TOC Heading', basedOn: 'Heading1', next: 'Normal', quickFormat: true, run: full('h1').run },
        { id: 'TOC1', name: 'toc 1', basedOn: 'Normal', next: 'Normal', paragraph: { spacing: { after: 60 } } },
        { id: 'TOC2', name: 'toc 2', basedOn: 'Normal', next: 'Normal', paragraph: { spacing: { after: 60 }, indent: { left: 220 } } },
        { id: 'TOC3', name: 'toc 3', basedOn: 'Normal', next: 'Normal', paragraph: { spacing: { after: 60 }, indent: { left: 440 } } },
        { id: 'TableofFigures', name: 'table of figures', basedOn: 'Normal', next: 'Normal', paragraph: { spacing: { after: 60 } } },
        { id: 'Bibliography', name: 'Bibliography', basedOn: 'Normal', next: 'Normal', paragraph: { indent: { left: 720, hanging: 720 }, spacing: { after: 120 } } },
        ...(D.design.customStyles || []).filter(u => u.type !== 'c').map(u => pstyle('BTSU' + u.id, u.name, u.id)),
      ],
      characterStyles: [
        ...D.BASE_STYLES.filter(st => st.type === 'c').map(st => ({ id: { emphasis: 'Emphasis', strong: 'Strong', 'subtle-em': 'SubtleEmphasis', 'intense-em': 'IntenseEmphasis', 'subtle-ref': 'SubtleReference', 'intense-ref': 'IntenseReference', 'book-title': 'BookTitle' }[st.id], name: st.name, basedOn: 'DefaultParagraphFont', quickFormat: true, run: runOf(D.styleDef(st.id), true) })),
        ...(D.design.customStyles || []).filter(u => u.type === 'c').map(u => ({ id: 'BTSU' + u.id, name: u.name, basedOn: 'DefaultParagraphFont', quickFormat: true, run: runOf(D.styleDef(u.id), true) })),
        { id: 'BTSCitation', name: 'BTS Citation', basedOn: 'DefaultParagraphFont', run: {} },
      ],
    };

    const comments = cids.length ? { children: cids.map((cid, i) => { const c = D.comments[cid] || {}; return { id: i, author: c.author || 'Student', initials: (c.author || 'S').slice(0, 2).toUpperCase(), date: new Date(c.time || Date.now()), children: (c.text || ' ').split('\n').map(t => new X.Paragraph({ children: [new X.TextRun(t || ' ')] })) }; }) } : undefined;

    const PR = BTS.props || {}; const author = (PR.author || '').trim() || BTS.userName || 'Student';
    const doc = new X.Document({
      creator: author, title: PR.title || BTS.title, subject: PR.subject || undefined, keywords: PR.keywords || undefined, description: PR.comments || undefined, lastModifiedBy: BTS.userName || author,
      styles, numbering: { config: numbering.length ? numbering : [{ reference: 'none', levels: [{ level: 0, format: X.LevelFormat.DECIMAL, text: '%1.', alignment: X.AlignmentType.LEFT }] }] },
      footnotes: fnSeq ? footnotes : undefined, comments,
      background: D.design.pageColor ? { color: D.design.pageColor.slice(1).toUpperCase() } : undefined,
      features: { updateFields: false, trackRevisions: !!D.design.track },
      sections,
    });
    const blob = await X.Packer.toBlob(doc);
    return await postProcess(new Uint8Array(await blob.arrayBuffer()));
  }

  /* ---------- package post-processing: properties, protection, BTS data part ---------- */
  const xesc = (v) => String(v ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  async function postProcess(bytes) {
    const zip = await JSZip.loadAsync(bytes);
    const P = BTS.props || {}; const author = (P.author || '').trim() || BTS.userName || 'Student';
    const ct = await zip.file('[Content_Types].xml').async('string'); let ctx = ct;
    const addOverride = (part, type) => { if (!ctx.includes(`PartName="${part}"`)) ctx = ctx.replace('</Types>', `<Override PartName="${part}" ContentType="${type}"/></Types>`); };
    // core properties
    const coreF = zip.file('docProps/core.xml');
    if (coreF) {
      let core = await coreF.async('string');
      const setTag = (tag, val) => { const re = new RegExp(`<${tag}[^>]*>[\\s\\S]*?</${tag}>|<${tag}\\s*/>`); const el = `<${tag}>${xesc(val)}</${tag}>`; if (!val) { core = core.replace(re, ''); return; } core = re.test(core) ? core.replace(re, el) : core.replace('</cp:coreProperties>', el + '</cp:coreProperties>'); };
      setTag('dc:creator', author); setTag('cp:category', P.category || ''); setTag('cp:contentStatus', D.design.protect && D.design.protect.mode === 'final' ? 'Final' : '');
      zip.file('docProps/core.xml', core);
    }
    // app properties (company)
    const appF = zip.file('docProps/app.xml');
    if (appF) { let app = await appF.async('string'); app = app.replace(/<Company>[\s\S]*?<\/Company>|<Company\/>/, ''); if (P.company) { if (/<Properties[^>]*\/>/.test(app)) app = app.replace(/<Properties([^>]*)\/>/, '<Properties$1></Properties>'); app = app.replace('</Properties>', `<Company>${xesc(P.company)}</Company></Properties>`); } zip.file('docProps/app.xml', app); }
    else if (P.company) {
      zip.file('docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>BTS Docs</Application><Company>${xesc(P.company)}</Company></Properties>`);
      addOverride('/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml');
      let rels = await zip.file('_rels/.rels').async('string'); if (!rels.includes('docProps/app.xml')) rels = rels.replace('</Relationships>', '<Relationship Id="rIdBtsApp" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'); zip.file('_rels/.rels', rels);
    }
    // Mark as final (Word's convention) as a custom property
    if (D.design.protect && D.design.protect.mode === 'final') {
      zip.file('docProps/custom.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="_MarkAsFinal"><vt:bool>true</vt:bool></property></Properties>`);
      addOverride('/docProps/custom.xml', 'application/vnd.openxmlformats-officedocument.custom-properties+xml');
      let rels = await zip.file('_rels/.rels').async('string'); if (!rels.includes('docProps/custom.xml')) rels = rels.replace('</Relationships>', '<Relationship Id="rIdBtsCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>'); zip.file('_rels/.rels', rels);
    }
    // read-only protection in settings.xml (after trackRevisions / displayBackgroundShape, before defaultTabStop)
    const setF = zip.file('word/settings.xml');
    if (setF && D.design.protect && D.design.protect.mode === 'readonly') {
      let st = await setF.async('string'); const el = '<w:documentProtection w:edit="readOnly" w:enforcement="1"/>';
      if (/<w:trackRevisions[^>]*\/>/.test(st)) st = st.replace(/(<w:trackRevisions[^>]*\/>)/, '$1' + el);
      else if (/<w:displayBackgroundShape[^>]*\/>/.test(st)) st = st.replace(/(<w:displayBackgroundShape[^>]*\/>)/, '$1' + el);
      else st = st.replace(/(<w:settings[^>]*>)/, '$1' + el);
      zip.file('word/settings.xml', st);
    }
    // BTS Docs data part (styles, sources, citations, protection) — custom XML, preserved by Word
    const data = { v: 1, app: 'BTS Docs', design: D.design, cites: $$('.cite', D.doc).map(c => ({ src: c.dataset.src || '', page: c.dataset.page || '' })) };
    zip.file('customXml/item1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><btsDocs xmlns="https://office.btspl.tech/ns/btsdocs/2026">${xesc(JSON.stringify(data))}</btsDocs>`);
    zip.file('customXml/itemProps1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="{6F1B6D2A-5C3E-4B8F-9A7D-B75D0C5E2026}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="https://office.btspl.tech/ns/btsdocs/2026"/></ds:schemaRefs></ds:datastoreItem>`);
    zip.file('customXml/_rels/item1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>`);
    addOverride('/customXml/itemProps1.xml', 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml');
    if (!/Extension="xml"/i.test(ctx)) ctx = ctx.replace(/(<Types[^>]*>)/, '$1<Default Extension="xml" ContentType="application/xml"/>');
    const drF = zip.file('word/_rels/document.xml.rels');
    if (drF) { let dr = await drF.async('string'); if (!dr.includes('customXml/item1.xml')) dr = dr.replace('</Relationships>', '<Relationship Id="rIdBtsData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>'); zip.file('word/_rels/document.xml.rels', dr); }
    zip.file('[Content_Types].xml', ctx);
    return await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }
})();
