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

  function bases() {
    const th = D.theme(); const head = D.design.headFont || th.head, body = D.design.bodyFont || th.body;
    const acc = th.accent.slice(1).toUpperCase(), dark = th.dark.slice(1).toUpperCase();
    const N = { font: body, size: 11, color: '111111', bold: false, italic: false };
    return {
      acc, dark, head, body,
      Normal: N, NoSpacing: N, ListParagraph: N, TOC1: N, TOC2: N, TOC3: N,
      Title: { font: head, size: 28, color: dark, bold: false, italic: false },
      Subtitle: { font: body, size: 13, color: '5A6275', bold: false, italic: false },
      Heading1: { font: head, size: 16, color: acc, bold: false, italic: false },
      Heading2: { font: head, size: 13, color: acc, bold: false, italic: false },
      Heading3: { font: head, size: 12, color: dark, bold: false, italic: false },
      Heading4: { font: head, size: 11, color: acc, bold: false, italic: true },
      Heading5: { font: head, size: 11, color: dark, bold: false, italic: false },
      Heading6: { font: head, size: 11, color: dark, bold: false, italic: false },
      Quote: { font: body, size: 11, color: '404654', bold: false, italic: true },
      Code: { font: 'Consolas', size: 10, color: '1D2433', bold: false, italic: false },
      Caption: { font: body, size: 9, color: dark, bold: false, italic: true },
    };
  }

  function styleIdOf(el) {
    const t = el.tagName;
    if (/^H[1-6]$/.test(t)) return 'Heading' + t[1];
    if (t === 'PRE') return 'Code';
    if (el.closest('blockquote') && D.rootOf(el)) return 'Quote';
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
    try { return await build(X); } finally { if (oldZoom !== 100) D.setZoom(oldZoom); }
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
    function runProps(el, block, base, inLink) {
      const cs = getComputedStyle(el); const o = {};
      let deco = new Set(), bg = null, va = null, sizeEl = el;
      for (let a = el; a; a = a.parentElement) {
        const s = a === el ? cs : getComputedStyle(a);
        (s.textDecorationLine || '').split(' ').forEach(d => { if (d && d !== 'none') deco.add(d); });
        if (a === block || D.isRoot(a) || a.tagName === 'TD' || a.tagName === 'TH') break;
        if (!bg && !a.matches('.cmt,.eq,.fld,.fn-ref,a')) bg = hex(s.backgroundColor);
        if (!va && (s.verticalAlign === 'super' || s.verticalAlign === 'sub')) { va = s.verticalAlign; sizeEl = a.parentElement; }
      }
      const fam = firstFamily(cs.fontFamily); if (fam && fam.toLowerCase() !== String(base.font).toLowerCase()) o.font = fam;
      const sz = Math.round(parseFloat(getComputedStyle(sizeEl).fontSize) * 0.75 * 2) / 2; if (sz && Math.abs(sz - base.size) > 0.01) o.size = Math.round(sz * 2);
      const bold = +cs.fontWeight >= 600; if (bold !== base.bold) o.bold = bold;
      const ital = cs.fontStyle === 'italic' || cs.fontStyle.startsWith('oblique'); if (ital !== base.italic) o.italics = ital;
      if (deco.has('underline')) o.underline = { type: X.UnderlineType.SINGLE };
      if (deco.has('line-through')) o.strike = true;
      const col = hex(cs.color); if (col && col !== base.color && !(inLink && col === '1F5FD6') && !(base.color === '111111' && col === '000000')) o.color = col;
      if (bg) { if (HIGHLIGHTS[bg]) o.highlight = HIGHLIGHTS[bg]; else o.shading = { type: X.ShadingType.CLEAR, fill: bg, color: 'auto' }; }
      if (va === 'super') o.superScript = true; else if (va === 'sub') o.subScript = true;
      if (cs.textTransform === 'uppercase') o.allCaps = true;
      if (cs.fontVariantCaps === 'small-caps' || cs.fontVariant === 'small-caps') o.smallCaps = true;
      if (inLink) o.style = 'Hyperlink';
      return o;
    }
    function textRuns(text, props, pre) {
      text = text.replace(/​/g, '');
      if (!pre) text = text.replace(/ /g, ' ');
      if (!text) return [];
      const out = [];
      const lines = pre ? text.split('\n') : [text];
      lines.forEach((ln, li) => {
        const parts = ln.split('\t'); const children = [];
        parts.forEach((p, i) => { if (i) children.push(new X.Tab()); if (p) children.push(p); });
        if (li > 0) out.push(new X.TextRun({ ...props, break: 1 }));
        if (children.length) out.push(new X.TextRun({ ...props, children }));
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
      const al = AL[cs.textAlign]; if (al) o.alignment = al;
      // spacing (emulate CSS margin collapsing between siblings)
      const fs = parseFloat(cs.fontSize) || 14.67;
      const before = parseFloat(cs.marginTop) || 0, after = parseFloat(cs.marginBottom) || 0;
      const spacing = { before: tw(Math.max(0, before - (ctx.prevAfter || 0))), after: tw(after) };
      const lh = parseFloat(cs.lineHeight); if (lh && cs.lineHeight !== 'normal') { spacing.line = Math.round(lh / fs * 240); spacing.lineRule = X.LineRuleType.AUTO; }
      o.spacing = spacing; ctx.prevAfter = after;
      // indents relative to the container's content box
      if (!extra.numbering && ctx.box) {
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
      if (el.tagName !== 'TD' && el.tagName !== 'TH') {
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
        const header = ri === 0 && /t-(header|banded|dark|grid-accent|light|minimal)/.test(t.className) && !t.classList.contains('t-noheader');
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
      const out = []; const title = toc.querySelector('.toc-title');
      if (title) out.push(new X.Paragraph({ style: 'TOCHeading', children: [new X.TextRun(title.textContent)] }));
      $$('p[data-target]', toc).forEach(p => {
        const lvl = +((p.className.match(/toc-l(\d)/) || [, 1])[1]);
        const text = p.querySelector('.toc-t').textContent, num = p.querySelector('.toc-n').textContent;
        out.push(new X.Paragraph({ style: 'TOC' + lvl, tabStops: [{ type: X.TabStopType.RIGHT, position: contentTw - 10, leader: X.LeaderType.DOT }], children: [new X.InternalHyperlink({ anchor: safeId(p.dataset.target), children: [new X.TextRun(text), new X.TextRun({ children: [new X.Tab(), num] })] })] }));
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

    /* ---------- styles ---------- */
    const ps2 = D.design.pspacing || { after: 8, line: 1.15, before: 0 };
    const run = (b, extra = {}) => ({ font: b.font, size: b.size * 2, color: b.color, bold: b.bold, italics: b.italic, ...extra });
    const styles = {
      default: {
        document: { run: { font: B.body, size: 22, color: '111111' }, paragraph: { spacing: { after: Math.round(ps2.after * 20), before: Math.round((ps2.before || 0) * 20), line: Math.round(ps2.line * 240), lineRule: X.LineRuleType.AUTO } } },
        title: { run: run(B.Title), paragraph: { spacing: { after: 80 } } },
        heading1: { run: run(B.Heading1), paragraph: { spacing: { before: 280, after: 80 }, keepNext: true, keepLines: true } },
        heading2: { run: run(B.Heading2), paragraph: { spacing: { before: 200, after: 60 }, keepNext: true, keepLines: true } },
        heading3: { run: run(B.Heading3), paragraph: { spacing: { before: 160, after: 60 }, keepNext: true, keepLines: true } },
        heading4: { run: run(B.Heading4), paragraph: { spacing: { before: 160, after: 40 }, keepNext: true } },
        heading5: { run: run(B.Heading5), paragraph: { spacing: { before: 120, after: 40 }, keepNext: true } },
        heading6: { run: run(B.Heading6), paragraph: { spacing: { before: 120, after: 40 }, keepNext: true } },
        hyperlink: { run: { color: '1F5FD6', underline: { type: X.UnderlineType.SINGLE } } },
        listParagraph: { paragraph: { spacing: { after: 40 } } },
      },
      paragraphStyles: [
        { id: 'Subtitle', name: 'Subtitle', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: run(B.Subtitle, { characterSpacing: 8 }), paragraph: { spacing: { after: 240 } } },
        { id: 'Quote', name: 'Quote', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: run(B.Quote), paragraph: { indent: { left: 432, right: 576 }, border: { left: { style: X.BorderStyle.SINGLE, size: 18, color: B.acc, space: 8 } }, spacing: { before: 200, after: 200 } } },
        { id: 'NoSpacing', name: 'No Spacing', basedOn: 'Normal', quickFormat: true, paragraph: { spacing: { after: 0, before: 0, line: 240, lineRule: X.LineRuleType.AUTO } } },
        { id: 'Caption', name: 'Caption', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: run(B.Caption), paragraph: { spacing: { before: 40, after: 200 } } },
        { id: 'Code', name: 'Code', basedOn: 'Normal', quickFormat: true, run: run(B.Code), paragraph: { spacing: { after: 160, line: 270, lineRule: X.LineRuleType.AUTO }, shading: { type: X.ShadingType.CLEAR, fill: 'F3F4F7', color: 'auto' } } },
        { id: 'TOCHeading', name: 'TOC Heading', basedOn: 'Heading1', next: 'Normal', quickFormat: true, run: run(B.Heading1) },
        { id: 'TOC1', name: 'toc 1', basedOn: 'Normal', next: 'Normal', paragraph: { spacing: { after: 60 } } },
        { id: 'TOC2', name: 'toc 2', basedOn: 'Normal', next: 'Normal', paragraph: { spacing: { after: 60 }, indent: { left: 220 } } },
        { id: 'TOC3', name: 'toc 3', basedOn: 'Normal', next: 'Normal', paragraph: { spacing: { after: 60 }, indent: { left: 440 } } },
      ],
    };

    const comments = cids.length ? { children: cids.map((cid, i) => { const c = D.comments[cid] || {}; return { id: i, author: c.author || 'Student', initials: (c.author || 'S').slice(0, 2).toUpperCase(), date: new Date(c.time || Date.now()), children: (c.text || ' ').split('\n').map(t => new X.Paragraph({ children: [new X.TextRun(t || ' ')] })) }; }) } : undefined;

    const doc = new X.Document({
      creator: D.author(), title: BTS.title, description: 'Created with BTS Docs', lastModifiedBy: D.author(),
      styles, numbering: { config: numbering.length ? numbering : [{ reference: 'none', levels: [{ level: 0, format: X.LevelFormat.DECIMAL, text: '%1.', alignment: X.AlignmentType.LEFT }] }] },
      footnotes: fnSeq ? footnotes : undefined, comments,
      background: D.design.pageColor ? { color: D.design.pageColor.slice(1).toUpperCase() } : undefined,
      features: { updateFields: false },
      sections,
    });
    const blob = await X.Packer.toBlob(doc);
    return new Uint8Array(await blob.arrayBuffer());
  }
})();
