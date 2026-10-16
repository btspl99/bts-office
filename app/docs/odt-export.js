/* BTS Docs 1.3 — HTML → OpenDocument Text (.odt), written by hand with JSZip.
 * Like the .docx exporter it reads the live DOM + computed styles: named paragraph styles (headings with outline
 * levels, Title, Quote, Code…), runs (font/size/colour/highlight/b/i/u/s/sub/sup/caps), alignment, spacing,
 * indents, borders & shading, page breaks, lists (bullets / numbers, nested, start values), tables (column widths,
 * borders, shading, merged cells, header rows), pictures (Pictures/, inline and wrapped), links and bookmarks,
 * footnotes, comments, tracked changes, columns (sections), header / footer with page numbers, page size,
 * orientation, margins, page colour and border, and the document properties (meta.xml). */
(function () {
  const D = window.Docs; const { $$, PX } = D;
  const X = (s) => String(s ?? '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const ptv = (px) => +(px * 0.75).toFixed(2) + 'pt';
  const cm = (inch) => +(inch * 2.54).toFixed(3) + 'cm';
  const cmPx = (px) => +(px / PX * 2.54).toFixed(3) + 'cm';
  const hex = (c) => {
    if (!c) return null;
    const cs = c.match(/color\(srgb\s+([\d.e-]+)\s+([\d.e-]+)\s+([\d.e-]+)(?:\s*\/\s*([\d.]+))?\)/);
    if (cs) { if (cs[4] !== undefined && +cs[4] < 0.05) return null; return '#' + [cs[1], cs[2], cs[3]].map(v => Math.round(Math.max(0, Math.min(1, +v)) * 255).toString(16).padStart(2, '0')).join(''); }
    const m = c.match(/rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?\)/);
    if (!m) return /^#[0-9a-f]{6}$/i.test(c) ? c.toLowerCase() : null;
    if (m[4] !== undefined && +m[4] < 0.05) return null;
    return '#' + [m[1], m[2], m[3]].map(v => Math.round(+v).toString(16).padStart(2, '0')).join('');
  };
  const firstFamily = (ff) => (ff || '').split(',')[0].trim().replace(/^["']|["']$/g, '');
  const NSDECL = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"';

  /* ---------- named styles (Docs style id → ODF style) ---------- */
  const NAMED = { normal: ['Standard', 'Default Paragraph Style'], nospace: ['No_20_Spacing', 'No Spacing'], title: ['Title', 'Title'], subtitle: ['Subtitle', 'Subtitle'], quote: ['Quotations', 'Quotations'], iquote: ['Intense_20_Quote', 'Intense Quote'], code: ['Preformatted_20_Text', 'Preformatted Text'], caption: ['Caption', 'Caption'], listpara: ['List_20_Paragraph', 'List Paragraph'] };
  for (let i = 1; i <= 6; i++) NAMED['h' + i] = ['Heading_20_' + i, 'Heading ' + i];
  function styleIdOf(el) {
    const t = el.tagName;
    if (/^H[1-6]$/.test(t)) return 'h' + t[1];
    if (t === 'PRE') return 'code';
    const bq = el.closest('blockquote'); if (bq && D.rootOf(el)) return bq.classList.contains('intense') ? 'iquote' : 'quote';
    const us = [...el.classList].find(c => c.startsWith('us-')); if (us && D.findStyle(us.slice(3))) return us.slice(3);
    for (const c of ['title', 'subtitle', 'nospace', 'caption', 'listpara']) if (el.classList.contains(c)) return c;
    return 'normal';
  }
  const odfName = (id) => NAMED[id] ? NAMED[id][0] : 'BTSU_' + String(id).replace(/[^A-Za-z0-9_]/g, '_');
  const GENERIC = { sans: 'swiss', serif: 'roman', display: 'decorative', mono: 'modern', indian: 'swiss', script: 'script' };
  const devaFor = (fam) => D.isIndianFont(fam) ? fam : (/serif|display/.test(D.fontCat(fam)) ? 'Noto Serif Devanagari' : 'Noto Sans Devanagari');

  /** A set of automatic styles for one XML file (content.xml or styles.xml). */
  function Bank(prefix) {
    const map = new Map(); const out = []; const n = {};
    return {
      get(family, key, xml, base) {
        const k = family + '|' + key; if (map.has(k)) return map.get(k);
        const p = { paragraph: 'P', text: 'T', table: 'Tbl', 'table-column': 'Col', 'table-cell': 'Cell', graphic: 'fr', section: 'Sect' }[family] || 'S';
        n[p] = (n[p] || 0) + 1; const name = prefix + p + n[p];
        out.push(`<style:style style:name="${name}" style:family="${family}"${base ? ` style:parent-style-name="${base}"` : ''}>${xml}</style:style>`); map.set(k, name); return name;
      },
      list(xml) { n.L = (n.L || 0) + 1; const name = prefix + 'L' + n.L; out.push(xml.replace('%NAME%', name)); return name; },
      xml: () => out.join(''),
    };
  }
  function textPropsXML(o) {
    let a = '';
    if (o.font) { a += ` style:font-name="${X(o.font)}"`; a += ` style:font-name-complex="${X(devaFor(o.font))}"`; }
    if (o.size) a += ` fo:font-size="${o.size}pt" style:font-size-complex="${o.size}pt"`;
    if (o.bold !== undefined) a += ` fo:font-weight="${o.bold ? 'bold' : 'normal'}" style:font-weight-complex="${o.bold ? 'bold' : 'normal'}"`;
    if (o.italic !== undefined) a += ` fo:font-style="${o.italic ? 'italic' : 'normal'}" style:font-style-complex="${o.italic ? 'italic' : 'normal'}"`;
    if (o.underline) a += ' style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"';
    if (o.strike) a += ' style:text-line-through-style="solid" style:text-line-through-type="single"';
    if (o.color) a += ` fo:color="${o.color}"`;
    if (o.bg) a += ` fo:background-color="${o.bg}"`;
    if (o.pos) a += ` style:text-position="${o.pos}"`;
    if (o.caps !== undefined) a += ` fo:text-transform="${o.caps ? 'uppercase' : 'none'}"`;
    if (o.smallcaps !== undefined) a += ` fo:font-variant="${o.smallcaps ? 'small-caps' : 'normal'}"`;
    if (o.letter) a += ` fo:letter-spacing="${o.letter}pt"`;
    return a ? `<style:text-properties${a}/>` : '';
  }

  D.exportOdt = async () => {
    if (!window.JSZip) throw new Error('ZIP library failed to load');
    D.endPreview(); D.clearCellSel && D.clearCellSel();
    const oldZoom = D.zoom; if (oldZoom !== 100) D.setZoom(100);
    const mmPrev = D.mm && D.mm.preview; if (mmPrev) D.mmPreview(false);
    document.body.classList.add('exporting');
    try { return await build(); } finally { document.body.classList.remove('exporting'); if (oldZoom !== 100) D.setZoom(oldZoom); if (mmPrev) D.mmPreview(true); }
  };

  async function build() {
    D.layout(); D.renderFootnotes && D.renderFootnotes();
    const fonts = new Set(['Noto Sans Devanagari', 'Noto Serif Devanagari']);
    const useFont = (f) => { if (f) fonts.add(f); return f; };
    const ps = D.paperSize(); const m = D.page.m;
    const contentWpx = (ps.w - m.l - m.r) * PX;

    /* pictures */
    const pics = new Map(); let picN = 0; const picFiles = [];
    for (const img of $$('img', D.doc).concat($$('img', D.hdr), $$('img', D.ftr))) {
      let src = img.getAttribute('src') || '';
      if (!src.startsWith('data:')) { try { const c = document.createElement('canvas'); c.width = img.naturalWidth; c.height = img.naturalHeight; c.getContext('2d').drawImage(img, 0, 0); src = c.toDataURL('image/png'); } catch { continue; } }
      if (!/^data:image\/(png|jpe?g|gif|svg\+xml)/.test(src)) src = await D.toWordFriendly(src);
      const mt = src.match(/^data:(image\/(png|jpe?g|gif|svg\+xml))/); if (!mt) continue;
      const ext = mt[2] === 'jpeg' ? 'jpg' : mt[2] === 'svg+xml' ? 'svg' : mt[2];
      const path = `Pictures/image${++picN}.${ext}`; picFiles.push({ path, mime: mt[1], data: BTS.dataURLToBytes(src) }); pics.set(img, { path, mime: mt[1] });
    }

    /* comments & tracked changes (content.xml only) */
    const changes = []; let chN = 0;
    const lastCmt = new Map(); $$('.cmt', D.doc).forEach(s => lastCmt.set(s.dataset.cid, s)); const cmtStarted = new Set();
    let fnN = 0, tblN = 0, frN = 0, sectN = 0;

    function makeWriter(bank, hf) {
      /* ----- runs ----- */
      function runProps(el, block, base) {
        const cs = getComputedStyle(el); const o = {};
        let deco = new Set(), bg = null, va = null, sizeEl = el;
        for (let a = el; a; a = a.parentElement) {
          const s = a === el ? cs : getComputedStyle(a);
          (s.textDecorationLine || '').split(' ').forEach(d => { if (d && d !== 'none') deco.add(d); });
          if (a === block || D.isRoot(a) || a.tagName === 'TD' || a.tagName === 'TH') break;
          if (!bg && !a.matches('.cmt,.eq,.fld,.fn-ref,a,ins,del')) bg = hex(s.backgroundColor);
          if (!va && (s.verticalAlign === 'super' || s.verticalAlign === 'sub')) { va = s.verticalAlign; sizeEl = a.parentElement; }
        }
        const fam = firstFamily(cs.fontFamily); if (fam && fam.toLowerCase() !== String(base.font).toLowerCase()) o.font = useFont(fam);
        const sz = Math.round(parseFloat(getComputedStyle(sizeEl).fontSize) * 0.75 * 2) / 2; if (sz && Math.abs(sz - base.size) > 0.01) o.size = sz;
        const bold = +cs.fontWeight >= 600; if (bold !== !!base.bold) o.bold = bold;
        const ital = cs.fontStyle === 'italic' || cs.fontStyle.startsWith('oblique'); if (ital !== !!base.italic) o.italic = ital;
        const inA = el.closest('a'); const inTC = el.closest('ins.tc-ins, del.tc-del');
        if (deco.has('underline') && !inA && !inTC) o.underline = true;
        if (deco.has('line-through') && !inTC) o.strike = true;
        const col = hex(cs.color); const bcol = String(D.resolveColor(base.color) || '#111111').toLowerCase();
        if (col && col !== bcol && !(inA && col === '#1f5fd6') && !inTC && !(bcol === '#111111' && col === '#000000')) o.color = col;
        if (bg) o.bg = bg;
        if (va) o.pos = va === 'super' ? 'super 58%' : 'sub 58%';
        if ((cs.textTransform === 'uppercase') !== !!base.caps) o.caps = cs.textTransform === 'uppercase';
        const sc = cs.fontVariantCaps === 'small-caps'; if (sc !== !!base.smallcaps) o.smallcaps = sc;
        return o;
      }
      function textXML(t, pre) {
        t = t.replace(/\u200B/g, '');
        if (!pre) t = t.replace(/[\n\r]+/g, ' ');
        let out = '';
        const lines = pre ? t.split('\n') : [t];
        lines.forEach((ln, li) => {
          if (li) out += '<text:line-break/>';
          ln.split('\t').forEach((part, pi) => {
            if (pi) out += '<text:tab/>';
            out += X(part).replace(/\u00A0/g, ' ').replace(/^ /, '<text:s/>').replace(/ {2,}/g, (s) => ' <text:s text:c="' + (s.length - 1) + '"/>');
          });
        });
        return out;
      }
      function span(inner, props) {
        if (!inner) return '';
        const k = JSON.stringify(props); if (k === '{}') return inner;
        return `<text:span text:style-name="${bank.get('text', k, textPropsXML(props))}">${inner}</text:span>`;
      }
      function frame(img, inline) {
        const p = pics.get(img); if (!p) return '';
        let w = parseFloat(img.style.width) || img.offsetWidth || img.naturalWidth || 100;
        let h = parseFloat(img.style.height) || img.offsetHeight || img.naturalHeight || 100;
        if (img.style.height === 'auto' || !h) h = img.offsetHeight || h;
        if (w > contentWpx) { h = h * contentWpx / w; w = contentWpx; }
        const side = img.classList.contains('img-left') ? 'left' : img.classList.contains('img-right') ? 'right' : null;
        const bordered = img.classList.contains('img-border');
        const gx = side ? `<style:graphic-properties style:wrap="${side === 'left' ? 'right' : 'left'}" style:number-wrapped-paragraphs="no-limit" style:horizontal-pos="${side}" style:horizontal-rel="paragraph" style:vertical-pos="top" style:vertical-rel="paragraph" fo:margin-left="0.2cm" fo:margin-right="0.2cm" fo:margin-top="0.1cm" fo:margin-bottom="0.1cm"${bordered ? ' fo:border="0.75pt solid #3c4250"' : ' fo:border="none"'}/>`
          : `<style:graphic-properties style:vertical-pos="top" style:vertical-rel="baseline" style:horizontal-pos="center" style:horizontal-rel="paragraph"${bordered ? ' fo:border="0.75pt solid #3c4250"' : ' fo:border="none"'}/>`;
        const fs = bank.get('graphic', gx, gx, 'Graphics');
        const alt = img.dataset.decorative ? '' : (img.alt || '');
        return `<draw:frame draw:style-name="${fs}" draw:name="Picture ${++frN}${hf ? 'h' : ''}" text:anchor-type="${side ? 'char' : 'as-char'}" svg:width="${cmPx(w)}" svg:height="${cmPx(h)}" draw:z-index="${frN}"><draw:image xlink:href="${p.path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" draw:mime-type="${p.mime}"/>${alt ? `<svg:title>${X(alt)}</svg:title><svg:desc>${X(alt)}</svg:desc>` : ''}</draw:frame>`;
      }
      function inline(block, base, opts = {}) {
        const pre = block.tagName === 'PRE' || getComputedStyle(block).whiteSpace.startsWith('pre');
        let out = '';
        const walk = (node, inTC) => {
          let s = '';
          for (const n of node.childNodes) {
            if (n.nodeType === 3) { s += span(textXML(n.data, pre), runProps(n.parentElement, block, base)); continue; }
            if (n.nodeType !== 1) continue;
            const t = n.tagName;
            if (t === 'UL' || t === 'OL' || (opts.skipBlocks && (D.BLOCK_TAGS.has(t) || t === 'TABLE'))) continue;
            if (t === 'BR') { const nx = n.nextSibling; if (nx && !(nx.nodeType === 3 && !nx.data.trim() && !nx.nextSibling)) s += '<text:line-break/>'; continue; }
            if (t === 'IMG') { s += frame(n); continue; }
            if (n.classList.contains('fn-ref')) {
              if (hf) continue; fnN++;
              s += `<text:note text:id="ftn${fnN}" text:note-class="footnote"><text:note-citation>${fnN}</text:note-citation><text:note-body>${String(n.dataset.text || '').split('\n').map(l => `<text:p text:style-name="Footnote">${textXML(l)}</text:p>`).join('')}</text:note-body></text:note>`;
              continue;
            }
            if (n.classList.contains('fld')) { s += span(n.dataset.fld === 'pages' ? `<text:page-count>${X(n.textContent || '1')}</text:page-count>` : `<text:page-number text:select-page="current">${X(n.textContent || '1')}</text:page-number>`, runProps(n, block, base)); continue; }
            if (n.classList.contains('tab')) { s += '<text:tab/>'; continue; }
            if ((n.classList.contains('tc-ins') || n.classList.contains('tc-del')) && !hf && !inTC) {
              const id = 'ct' + (++chN); const info = `<office:change-info><dc:creator>${X(n.dataset.author || 'Unknown')}</dc:creator><dc:date>${X(new Date(n.dataset.time || Date.now()).toISOString().replace(/\.\d+Z$/, ''))}</dc:date></office:change-info>`;
              if (n.classList.contains('tc-ins')) { changes.push(`<text:changed-region xml:id="${id}" text:id="${id}"><text:insertion>${info}</text:insertion></text:changed-region>`); s += `<text:change-start text:change-id="${id}"/>${walk(n, true)}<text:change-end text:change-id="${id}"/>`; }
              else { changes.push(`<text:changed-region xml:id="${id}" text:id="${id}"><text:deletion>${info}<text:p>${span(textXML(n.textContent), runProps(n, block, base))}</text:p></text:deletion></text:changed-region>`); s += `<text:change text:change-id="${id}"/>`; }
              continue;
            }
            if (n.classList.contains('eq') || n.classList.contains('frac') || n.classList.contains('mf') || (n.getAttribute('contenteditable') === 'false' && !n.querySelector('img'))) {
              const txt = n.classList.contains('mf') ? '«' + (n.dataset.field || 'Field') + '»' : n.classList.contains('frac') ? [...n.children].map(c => c.textContent).join('/') : n.textContent;
              s += span(textXML(txt), runProps(n, block, base)); continue;
            }
            if (t === 'A' && n.getAttribute('href')) {
              const inner = walk(n, inTC); if (!inner) continue;
              s += `<text:a xlink:type="simple" xlink:href="${X(n.getAttribute('href'))}" text:style-name="Internet_20_link" text:visited-style-name="Visited_20_Internet_20_Link">${inner}</text:a>`; continue;
            }
            const cid = !hf && n.classList.contains('cmt') ? n.dataset.cid : null;
            if (cid && !cmtStarted.has(cid)) {
              cmtStarted.add(cid); const c = D.comments[cid] || {};
              s += `<office:annotation office:name="${X(cid)}"><dc:creator>${X(c.author || 'Student')}</dc:creator><dc:date>${X(new Date(c.time || Date.now()).toISOString().replace(/\.\d+Z$/, ''))}</dc:date>${String(c.text || ' ').split('\n').map(l => `<text:p>${textXML(l)}</text:p>`).join('')}</office:annotation>`;
            }
            s += walk(n, inTC);
            if (cid && lastCmt.get(cid) === n) s += `<office:annotation-end office:name="${X(cid)}"/>`;
          }
          return s;
        };
        out = walk(block, false);
        return out;
      }

      /* ----- paragraphs ----- */
      function borderXML(cs, side) {
        const st = cs['border' + side + 'Style'], w = parseFloat(cs['border' + side + 'Width']) || 0, col = hex(cs['border' + side + 'Color']);
        if (!st || st === 'none' || st === 'hidden' || !w || !col) return 'none';
        return `${+(w * 0.75).toFixed(2)}pt ${st === 'double' ? 'double' : st === 'dashed' ? 'dashed' : st === 'dotted' ? 'dotted' : 'solid'} ${col}`;
      }
      function paraStyle(el, sid, ctx, extra = {}) {
        const cs = getComputedStyle(el); const def = D.resolvedStyle(sid); let a = '', inner = '';
        const al = cs.textAlign.replace('start', 'left').replace(/^end$/, 'right');
        if (el.style.textAlign || al !== (def.align || 'left') || extra.align) a += ` fo:text-align="${({ left: 'start', right: 'end', center: 'center', justify: 'justify' })[extra.align || al] || 'start'}"`;
        const direct = !!(el.style.marginTop || el.style.marginBottom || el.style.lineHeight || el.closest('td,th,.textbox,.cols') || hf || el.closest('.biblio,.toc'));
        if (direct) {
          a += ` fo:margin-top="${ptv(parseFloat(cs.marginTop) || 0)}" fo:margin-bottom="${ptv(parseFloat(cs.marginBottom) || 0)}"`;
          const lh = parseFloat(cs.lineHeight), fs = parseFloat(cs.fontSize) || 14.67; if (lh && cs.lineHeight !== 'normal') a += ` fo:line-height="${Math.round(lh / fs * 100)}%"`;
        }
        if (!extra.inList && (el.style.marginLeft || el.style.textIndent || el.style.paddingLeft || el.style.marginRight)) {
          a += ` fo:margin-left="${cmPx((parseFloat(cs.marginLeft) || 0) + (parseFloat(cs.paddingLeft) || 0))}" fo:margin-right="${cmPx(parseFloat(cs.marginRight) || 0)}" fo:text-indent="${cmPx(parseFloat(cs.textIndent) || 0)}"`;
        }
        if (el.style.backgroundColor || el.style.borderTop || el.style.borderBottom || el.style.borderLeft || el.style.borderRight || el.style.border) {
          const bg = hex(cs.backgroundColor); if (bg) a += ` fo:background-color="${bg}"`;
          ['Top', 'Bottom', 'Left', 'Right'].forEach(s => { a += ` fo:border-${s.toLowerCase()}="${borderXML(cs, s)}"`; });
          a += ` fo:padding="${ptv(parseFloat(cs.paddingTop) || 0)}"`;
        }
        if (extra.border) a += extra.border;
        if (ctx.pendingBreak) { a += ' fo:break-before="page"'; ctx.pendingBreak = false; }
        if (/^H[1-6]$/.test(el.tagName) && !NAMED[sid]) a += ' fo:keep-with-next="always"';
        const ppx = a ? `<style:paragraph-properties${a}/>` : '';
        return bank.get('paragraph', sid + ppx + inner, ppx, odfName(sid));
      }
      function paragraph(el, ctx, extra = {}) {
        const sid = extra.sid || styleIdOf(el);
        const def = D.resolvedStyle(sid); const base = { font: def.font, size: def.size || 11, color: def.color, bold: !!def.bold, italic: !!def.italic, caps: !!def.caps, smallcaps: !!def.smallcaps };
        useFont(def.font);
        const imgCenter = el.querySelector(':scope > img.img-center') && !el.style.textAlign;
        const name = paraStyle(el, sid, ctx, Object.assign({}, extra, imgCenter ? { align: 'center' } : {}));
        const bm = el.id && !hf ? `<text:bookmark text:name="${X(el.id)}"/>` : '';
        const body = bm + inline(el, base, extra);
        const lvl = /^H([1-6])$/.exec(el.tagName);
        return lvl ? `<text:h text:style-name="${name}" text:outline-level="${lvl[1]}">${body}</text:h>` : `<text:p text:style-name="${name}">${body}</text:p>`;
      }
      const emptyP = (ctx, extraProps = '') => { let a = extraProps; if (ctx.pendingBreak) { a += ' fo:break-before="page"'; ctx.pendingBreak = false; } return `<text:p text:style-name="${a ? bank.get('paragraph', 'empty' + a, `<style:paragraph-properties${a}/>`, 'Standard') : 'Standard'}"/>`; };

      /* ----- lists ----- */
      const NUMFMT = { decimal: ['1', '', '.'], 'decimal-leading-zero': ['1', '', '.'], 'lower-alpha': ['a', '', '.'], 'lower-latin': ['a', '', '.'], 'upper-alpha': ['A', '', '.'], 'upper-latin': ['A', '', '.'], 'lower-roman': ['i', '', '.'], 'upper-roman': ['I', '', '.'], 'paren-decimal': ['1', '(', ')'], 'paren-lower-alpha': ['a', '(', ')'], 'paren-lower-roman': ['i', '(', ')'] };
      const BUL = { disc: '•', circle: '◦', square: '▪', none: '' };
      function listStyle(top) {
        const fmts = [], starts = [], kinds = [];
        const scan = (list, d) => { if (d > 9) return; if (!fmts[d]) { fmts[d] = getComputedStyle(list).listStyleType; starts[d] = list.tagName === 'OL' ? (+list.getAttribute('start') || 1) : 1; kinds[d] = list.tagName; } [...list.children].forEach(li => [...li.children].forEach(c => { if (c.tagName === 'UL' || c.tagName === 'OL') scan(c, d + 1); })); };
        scan(top, 0);
        let lv = '';
        for (let i = 0; i < 10; i++) {
          const kind = kinds[i] || top.tagName; let f = fmts[i] || (kind === 'OL' ? ['decimal', 'lower-alpha', 'lower-roman'][i % 3] : ['disc', 'circle', 'square'][i % 3]);
          if (/^"/.test(f)) f = 'disc'; // custom string markers
          const ml = 1.27 + i * 0.635;
          const props = `<style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="${ml.toFixed(3)}cm" fo:text-indent="-0.635cm" fo:margin-left="${ml.toFixed(3)}cm"/></style:list-level-properties>`;
          if (NUMFMT[f] || (kind === 'OL' && !BUL[f])) { const [nf, pre, suf] = NUMFMT[f] || NUMFMT.decimal; lv += `<text:list-level-style-number text:level="${i + 1}" style:num-prefix="${pre}" style:num-suffix="${suf}" style:num-format="${nf}"${starts[i] > 1 ? ` text:start-value="${starts[i]}"` : ''}>${props}</text:list-level-style-number>`; }
          else lv += `<text:list-level-style-bullet text:level="${i + 1}" text:bullet-char="${X(BUL[f] ?? '•')}">${props}<style:text-properties style:font-name="OpenSymbol"/></text:list-level-style-bullet>`;
        }
        fonts.add('OpenSymbol');
        return bank.list(`<text:list-style style:name="%NAME%" style:display-name="${top.tagName === 'OL' ? 'Numbering' : 'Bullets'}">${lv}</text:list-style>`);
      }
      function list(el, ctx, styleName) {
        let s = `<text:list${styleName ? ` text:style-name="${styleName}"` : ''}>`;
        for (const li of el.children) {
          if (li.tagName !== 'LI') continue;
          s += '<text:list-item>';
          const blocks = [...li.children].filter(c => D.BLOCK_TAGS.has(c.tagName) && c.tagName !== 'LI');
          const hasInline = [...li.childNodes].some(n => (n.nodeType === 3 && n.data.trim()) || (n.nodeType === 1 && !D.BLOCK_TAGS.has(n.tagName) && !['UL', 'OL', 'TABLE'].includes(n.tagName)));
          if (hasInline || !blocks.length) s += paragraph(li, ctx, { skipBlocks: true, sid: 'listpara', inList: true });
          for (const c of li.children) {
            if (c.tagName === 'UL' || c.tagName === 'OL') s += list(c, ctx, null);
            else if (c.tagName === 'TABLE') s += table(c, ctx);
            else if (D.BLOCK_TAGS.has(c.tagName)) s += paragraph(c, ctx, { inList: true });
          }
          s += '</text:list-item>';
        }
        return s + '</text:list>';
      }

      /* ----- tables ----- */
      function table(t, ctx) {
        const g = D.tableGrid(t); if (!g.rows || !g.cols) return '';
        const name = 'Table' + (++tblN) + (hf ? 'h' : '');
        const totalPx = Math.min(t.getBoundingClientRect().width || contentWpx, contentWpx);
        let wpx = new Array(g.cols).fill(0);
        for (let c = 0; c < g.cols; c++) for (let r = 0; r < g.rows; r++) { const x = g.g[r][c]; if (x && (x.colSpan || 1) === 1) { wpx[c] = x.getBoundingClientRect().width; break; } }
        const cols = [...(t.querySelector(':scope > colgroup')?.children || [])];
        if (cols.length === g.cols) cols.forEach((col, i) => { const s = col.style.width; if (s.endsWith('%')) wpx[i] = totalPx * parseFloat(s) / 100; else if (D.toPx(s)) wpx[i] = D.toPx(s); });
        const missing = wpx.filter(w => !w).length; if (missing) { const rest = Math.max(20 * missing, totalPx - wpx.reduce((a, b) => a + b, 0)); wpx = wpx.map(w => w || rest / missing); }
        const sum = wpx.reduce((a, b) => a + b, 0); const k = sum > totalPx ? totalPx / sum : 1; wpx = wpx.map(w => Math.max(12, w * k));
        const tcs = getComputedStyle(t); const ml = parseFloat(tcs.marginLeft) || 0;
        const centered = tcs.marginLeft === tcs.marginRight && ml > 1;
        const tp = `<style:table-properties style:width="${cmPx(wpx.reduce((a, b) => a + b, 0))}" table:align="${centered ? 'center' : 'left'}"${!centered && ml > 1 ? ` fo:margin-left="${cmPx(ml)}"` : ''} fo:margin-top="0cm" fo:margin-bottom="${ptv(parseFloat(tcs.marginBottom) || 0)}"${ctx.pendingBreak ? ' fo:break-before="page"' : ''}/>`;
        ctx.pendingBreak = false;
        const tsn = bank.get('table', tp, tp);
        let s = `<table:table table:name="${name}" table:style-name="${tsn}">`;
        wpx.forEach(w => { const cp = `<style:table-column-properties style:column-width="${cmPx(w)}"/>`; s += `<table:table-column table:style-name="${bank.get('table-column', cp, cp)}"/>`; });
        const origin = new Map();
        for (let r = 0; r < g.rows; r++) for (let c = 0; c < g.cols; c++) { const x = g.g[r][c]; if (x && !origin.has(x)) origin.set(x, [r, c]); }
        const header = (/t-(header|banded|dark|grid-accent|light|minimal|hrow)/.test(t.className) || [...(t.rows[0]?.cells || [])].every(c => c.tagName === 'TH')) && !t.classList.contains('t-noheader');
        for (let r = 0; r < g.rows; r++) {
          const row = t.rows[r]; let rs = '';
          for (let c = 0; c < g.cols; c++) {
            const cell = g.g[r][c];
            if (!cell) { rs += '<table:table-cell office:value-type="string"><text:p text:style-name="Standard"/></table:table-cell>'; continue; }
            const o = origin.get(cell); if (o[0] !== r || o[1] !== c) { rs += '<table:covered-table-cell/>'; continue; }
            const cs = getComputedStyle(cell); let a = '';
            const bg = hex(cs.backgroundColor); a += ` fo:background-color="${bg || 'transparent'}"`;
            ['Top', 'Bottom', 'Left', 'Right'].forEach(sd => { a += ` fo:border-${sd.toLowerCase()}="${borderXML(cs, sd)}"`; });
            a += ` fo:padding-top="${ptv(parseFloat(cs.paddingTop) || 0)}" fo:padding-bottom="${ptv(parseFloat(cs.paddingBottom) || 0)}" fo:padding-left="${ptv(parseFloat(cs.paddingLeft) || 0)}" fo:padding-right="${ptv(parseFloat(cs.paddingRight) || 0)}"`;
            if (cs.verticalAlign === 'middle') a += ' style:vertical-align="middle"'; else if (cs.verticalAlign === 'bottom') a += ' style:vertical-align="bottom"'; else a += ' style:vertical-align="top"';
            const cp = `<style:table-cell-properties${a}/>`;
            const inner = blocks(cell, { pendingBreak: false }) || '<text:p text:style-name="Standard"/>';
            rs += `<table:table-cell table:style-name="${bank.get('table-cell', cp, cp)}" office:value-type="string"${(cell.colSpan || 1) > 1 ? ` table:number-columns-spanned="${cell.colSpan}"` : ''}${(cell.rowSpan || 1) > 1 ? ` table:number-rows-spanned="${cell.rowSpan}"` : ''}>${inner}</table:table-cell>`;
          }
          const tr = `<table:table-row>${rs}</table:table-row>`;
          s += r === 0 && header ? `<table:table-header-rows>${tr}</table:table-header-rows>` : tr;
          void row;
        }
        return s + '</table:table>';
      }

      /* ----- blocks ----- */
      function blocks(container, ctx) {
        let s = '';
        const kids = container.childNodes ? [...container.childNodes] : container;
        for (const n of kids) {
          if (n.nodeType === 3) { if (n.data.trim()) s += `<text:p text:style-name="Standard">${textXML(n.data)}</text:p>`; continue; }
          if (n.nodeType !== 1) continue;
          const t = n.tagName;
          if (n.id === 'fn-area') continue;
          if (t === 'HR') {
            if (n.classList.contains('page-break')) { if (ctx.pendingBreak) s += emptyP(ctx); ctx.pendingBreak = true; }
            else s += emptyP(ctx, ' fo:border-bottom="0.75pt solid #9aa1ad" fo:padding-bottom="0.05cm" fo:margin-top="6pt" fo:margin-bottom="6pt"');
            continue;
          }
          if (n.classList.contains('toc')) {
            const title = n.querySelector('.toc-title'); if (title) s += `<text:p text:style-name="Contents_20_Heading">${textXML(title.textContent)}</text:p>`;
            $$('p[data-target]', n).forEach(p => { const lvl = +((p.className.match(/toc-l(\d)/) || [, 1])[1]); const tt = p.querySelector('.toc-t'), nn = p.querySelector('.toc-n'); s += `<text:p text:style-name="Contents_20_${Math.min(3, lvl)}"><text:a xlink:type="simple" xlink:href="#${X(p.dataset.target)}">${textXML(tt ? tt.textContent : p.textContent)}</text:a><text:tab/>${textXML(nn ? nn.textContent : '')}</text:p>`; });
            continue;
          }
          if (t === 'UL' || t === 'OL') { s += list(n, ctx, listStyle(n)); continue; }
          if (t === 'TABLE') { s += table(n, ctx); continue; }
          if (n.classList.contains('cols') && !hf) {
            const k = n.classList.contains('cols-3') ? 3 : 2; const sp = `<style:section-properties text:dont-balance-text-columns="false"><style:columns fo:column-count="${k}" fo:column-gap="0.635cm"/></style:section-properties>`;
            s += `<text:section text:style-name="${bank.get('section', sp, sp)}" text:name="Columns${++sectN}">${blocks(n, ctx) || '<text:p text:style-name="Standard"/>'}</text:section>`; continue;
          }
          if (n.classList.contains('textbox') || ((t === 'DIV' || t === 'BLOCKQUOTE' || t === 'SECTION') && D.hasBlockChild(n))) { s += blocks(n, ctx); continue; }
          if (t === 'IMG') { s += `<text:p text:style-name="${paraStyleForImg(n, ctx)}">${frame(n)}</text:p>`; continue; }
          s += paragraph(n, ctx);
        }
        return s;
      }
      function paraStyleForImg(img, ctx) { const a = (img.classList.contains('img-center') ? ' fo:text-align="center"' : '') + (ctx.pendingBreak ? ' fo:break-before="page"' : ''); ctx.pendingBreak = false; return a ? bank.get('paragraph', 'img' + a, `<style:paragraph-properties${a}/>`, 'Standard') : 'Standard'; }
      return { blocks };
    }

    /* ----- body ----- */
    const cb = Bank(''); const W = makeWriter(cb, false);
    $$('img.img-center', D.doc).forEach(img => { const b = D.blockOf(img); if (b && !b.style.textAlign) { b.dataset.tmpAlign = '1'; b.style.textAlign = 'center'; } });
    const ctx0 = { pendingBreak: false };
    let bodyXML;
    try { bodyXML = W.blocks(D.doc, ctx0); } finally { $$('[data-tmp-align]', D.doc).forEach(b => { b.style.textAlign = ''; b.removeAttribute('data-tmp-align'); if (!b.getAttribute('style')) b.removeAttribute('style'); }); }
    if (ctx0.pendingBreak) bodyXML += `<text:p text:style-name="${cb.get('paragraph', 'brk', '<style:paragraph-properties fo:break-before="page"/>', 'Standard')}"/>`;
    if (!bodyXML) bodyXML = '<text:p text:style-name="Standard"/>';

    /* ----- header / footer (styles.xml) ----- */
    const sb = Bank('M'); const HW = makeWriter(sb, true);
    const hasHdr = !!(D.hdr.textContent.trim() || D.hdr.querySelector('img')), hasFtr = !!(D.ftr.textContent.trim() || D.ftr.querySelector('img'));
    const hdrXML = hasHdr ? HW.blocks(D.hdr, { pendingBreak: false }) : '';
    const ftrXML = hasFtr ? HW.blocks(D.ftr, { pendingBreak: false }) : '';

    /* ----- named styles ----- */
    const fontsOfStyles = [];
    const namedStyle = (id, fam = 'paragraph') => {
      const [name, disp] = NAMED[id] || [odfName(id), (D.findStyle(id) || {}).name || id];
      const r = D.resolvedStyle(id); useFont(r.font); fontsOfStyles.push(r.font);
      const tp = textPropsXML({ font: r.font, size: r.size || 11, color: String(r.color || '#111111').toLowerCase(), bold: !!r.bold, italic: !!r.italic, caps: r.caps ? true : undefined, smallcaps: r.smallcaps ? true : undefined, letter: r.letter || undefined });
      const al = r.align && r.align !== 'left' ? ` fo:text-align="${({ right: 'end', center: 'center', justify: 'justify' })[r.align] || 'start'}"` : '';
      const ind = id === 'listpara' ? '' : (r.indent ? ` fo:margin-left="${cm(r.indent / 2.54)}"` : '') + (r.firstLine ? ` fo:text-indent="${cm(r.firstLine / 2.54)}"` : '');
      const lvl = /^h([1-6])$/.exec(id);
      const extra = id === 'code' ? ' fo:background-color="#f3f4f7" fo:padding="0.15cm" fo:border="0.5pt solid #e2e5eb"' : id === 'quote' ? ' fo:border-left="2.25pt solid ' + D.accent().toLowerCase() + '" fo:padding-left="0.42cm" fo:margin-right="1cm"' : id === 'iquote' ? ` fo:border-top="0.75pt solid ${D.accent().toLowerCase()}" fo:border-bottom="0.75pt solid ${D.accent().toLowerCase()}" fo:padding="0.28cm"` : '';
      const pp = `<style:paragraph-properties fo:margin-top="${+(r.before || 0)}pt" fo:margin-bottom="${+(r.after || 0)}pt"${r.line ? ` fo:line-height="${Math.round(r.line * 100)}%"` : ''}${al}${ind}${lvl ? ' fo:keep-with-next="always"' : ''}${extra}/>`;
      return `<style:style style:name="${name}" style:display-name="${X(disp)}" style:family="${fam}"${id === 'normal' ? '' : ' style:parent-style-name="Standard"'}${lvl ? ` style:default-outline-level="${lvl[1]}"` : ''} style:next-style-name="Standard"${lvl ? ' style:class="text"' : ''}>${pp}${tp}</style:style>`;
    };
    const ids = ['normal', 'nospace', 'title', 'subtitle', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'quote', 'iquote', 'code', 'caption', 'listpara'].concat((D.design.customStyles || []).filter(u => u.type !== 'c').map(u => u.id));
    const namedXML = ids.map(id => namedStyle(id)).join('');
    const nb = D.resolvedStyle('normal');
    const acc = D.accent().toLowerCase();
    const outline = `<text:outline-style style:name="Outline">${[1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map(i => `<text:outline-level-style text:level="${i}" style:num-format=""><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab"/></style:list-level-properties></text:outline-level-style>`).join('')}</text:outline-style>`;
    const fnConf = '<text:notes-configuration text:note-class="footnote" style:num-format="1" text:start-value="0" text:footnotes-position="page" text:start-numbering-at="document"/>';

    /* ----- page layout ----- */
    const P = D.PAPER[D.page.size] || D.PAPER.A4; const land = D.page.orient === 'landscape';
    const pw = land ? P.h : P.w, ph = land ? P.w : P.h;
    const hd = D.page.hdrDist ?? 0.5, fd = D.page.ftrDist ?? 0.5;
    const pageBg = D.design.pageColor ? ` fo:background-color="${X(D.design.pageColor)}"` : '';
    const bd = D.design.border; const pageBorder = bd ? ` fo:border="${bd.style === 'double' ? '2.25pt double' : (+(bd.width || 1.5) * 0.75).toFixed(2) + 'pt ' + (bd.style === 'dashed' ? 'dashed' : bd.style === 'dotted' ? 'dotted' : 'solid')} ${X(bd.color || '#1f2533')}"` : '';
    // Docs draws the page border 0.35" from the paper edge: the ODF border sits on the margin, so move the margin out to 0.35" and pad the rest
    const eff = { t: hasHdr ? hd : m.t, b: hasFtr ? fd : m.b, l: m.l, r: m.r }; const pm = {}, pad = {};
    Object.keys(eff).forEach(k => { pm[k] = bd ? Math.min(0.35, eff[k]) : eff[k]; pad[k] = eff[k] - pm[k]; });
    const padXML = bd ? ` fo:padding-top="${cm(pad.t)}" fo:padding-bottom="${cm(pad.b)}" fo:padding-left="${cm(pad.l)}" fo:padding-right="${cm(pad.r)}"` : '';
    const layout = `<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="${cm(pw)}" fo:page-height="${cm(ph)}" style:print-orientation="${land ? 'landscape' : 'portrait'}" fo:margin-top="${cm(pm.t)}" fo:margin-bottom="${cm(pm.b)}" fo:margin-left="${cm(pm.l)}" fo:margin-right="${cm(pm.r)}" style:writing-mode="lr-tb"${pageBg}${pageBorder}${padXML}>${D.design.pageColor ? `<style:background-image/>` : ''}<style:footnote-sep style:width="0.018cm" style:distance-before-sep="0.101cm" style:distance-after-sep="0.101cm" style:line-style="solid" style:adjustment="left" style:rel-width="25%" style:color="#000000"/></style:page-layout-properties>`
      + `<style:header-style>${hasHdr ? `<style:header-properties fo:min-height="0cm" fo:margin-bottom="${cm(Math.max(0.1, (m.t - hd) * 0.4))}" style:dynamic-spacing="false"/>` : ''}</style:header-style>`
      + `<style:footer-style>${hasFtr ? `<style:footer-properties fo:min-height="0cm" fo:margin-top="${cm(Math.max(0.1, (m.b - fd) * 0.4))}" style:dynamic-spacing="false"/>` : ''}</style:footer-style></style:page-layout>`;
    const master = `<style:master-page style:name="Standard" style:page-layout-name="pm1">${hasHdr ? `<style:header>${hdrXML}</style:header>${D.design.hdrFirst ? '<style:header-first style:display="false"/>' : ''}` : ''}${hasFtr ? `<style:footer>${ftrXML}</style:footer>${D.design.hdrFirst ? '<style:footer-first style:display="false"/>' : ''}` : ''}</style:master-page>`;

    const fontDecls = () => `<office:font-face-decls>${[...fonts].filter(Boolean).map(f => `<style:font-face style:name="${X(f)}" svg:font-family="${X(/\s/.test(f) ? "'" + f + "'" : f)}" style:font-family-generic="${GENERIC[D.fontCat(f)] || 'swiss'}" style:font-pitch="${D.fontCat(f) === 'mono' ? 'fixed' : 'variable'}"/>`).join('')}</office:font-face-decls>`;

    const stylesXML = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${NSDECL} office:version="1.3">${fontDecls()}<office:styles>`
      + `<style:default-style style:family="paragraph"><style:paragraph-properties style:tab-stop-distance="1.27cm" style:writing-mode="page"/><style:text-properties style:font-name="${X(nb.font)}" fo:font-size="${nb.size || 11}pt" fo:language="en" fo:country="IN" style:font-name-complex="${X(devaFor(nb.font))}" style:font-size-complex="${nb.size || 11}pt" style:language-complex="hi" style:country-complex="IN"/></style:default-style>`
      + `<style:default-style style:family="graphic"><style:graphic-properties svg:stroke-color="#3465a4" draw:fill-color="#729fcf" style:flow-with-text="false"/></style:default-style>`
      + `<style:default-style style:family="table"><style:table-properties table:border-model="collapsing"/></style:default-style>`
      + namedXML
      + `<style:style style:name="Footnote" style:display-name="Footnote" style:family="paragraph" style:parent-style-name="Standard" style:class="extra"><style:paragraph-properties fo:margin-left="0.5cm" fo:text-indent="-0.5cm" fo:margin-top="0pt" fo:margin-bottom="0pt"/><style:text-properties fo:font-size="9pt" style:font-size-complex="9pt"/></style:style>`
      + `<style:style style:name="Contents_20_Heading" style:display-name="Contents Heading" style:family="paragraph" style:parent-style-name="Heading_20_1" style:class="index"/>`
      + [1, 2, 3].map(i => `<style:style style:name="Contents_20_${i}" style:display-name="Contents ${i}" style:family="paragraph" style:parent-style-name="Standard" style:class="index"><style:paragraph-properties fo:margin-left="${((i - 1) * 0.5).toFixed(2)}cm" fo:margin-top="0pt" fo:margin-bottom="3pt"><style:tab-stops><style:tab-stop style:position="${cmPx(contentWpx - (i - 1) * 0.5 / 2.54 * PX - 2)}" style:type="right" style:leader-style="dotted" style:leader-text="."/></style:tab-stops></style:paragraph-properties></style:style>`).join('')
      + `<style:style style:name="Graphics" style:family="graphic"><style:graphic-properties text:anchor-type="paragraph" svg:x="0cm" svg:y="0cm" style:wrap="none" style:vertical-pos="top" style:vertical-rel="paragraph" style:horizontal-pos="center" style:horizontal-rel="paragraph"/></style:style>`
      + `<style:style style:name="Internet_20_link" style:display-name="Internet link" style:family="text"><style:text-properties fo:color="#1f5fd6" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>`
      + `<style:style style:name="Visited_20_Internet_20_Link" style:display-name="Visited Internet Link" style:family="text"><style:text-properties fo:color="#800080" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>`
      + `<style:style style:name="Footnote_20_anchor" style:display-name="Footnote anchor" style:family="text"><style:text-properties style:text-position="super 58%"/></style:style>`
      + outline + fnConf
      + `</office:styles><office:automatic-styles>${layout}${sb.xml()}</office:automatic-styles><office:master-styles>${master}</office:master-styles></office:document-styles>`;
    void acc; void fontsOfStyles;

    const contentXML = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${NSDECL} office:version="1.3">${fontDecls()}<office:automatic-styles>${cb.xml()}</office:automatic-styles><office:body><office:text>`
      + (changes.length ? `<text:tracked-changes text:track-changes="${D.design.track ? 'true' : 'false'}">${changes.join('')}</text:tracked-changes>` : '')
      + bodyXML + '</office:text></office:body></office:document-content>';

    const pr = BTS.props || {}; const now = new Date().toISOString().replace(/\.\d+Z$/, '');
    const words = D.countWords ? D.countWords(D.docText()) : 0;
    const metaXML = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-meta ${NSDECL} office:version="1.3"><office:meta><meta:generator>BTS Docs/1.3 (BTS Office)</meta:generator>`
      + (pr.title ? `<dc:title>${X(pr.title)}</dc:title>` : '') + (pr.subject ? `<dc:subject>${X(pr.subject)}</dc:subject>` : '') + (pr.comments ? `<dc:description>${X(pr.comments)}</dc:description>` : '')
      + String(pr.keywords || '').split(/[,;]/).map(k => k.trim()).filter(Boolean).map(k => `<meta:keyword>${X(k)}</meta:keyword>`).join('')
      + (pr.author ? `<meta:initial-creator>${X(pr.author)}</meta:initial-creator><dc:creator>${X(pr.author)}</dc:creator>` : '')
      + `<meta:creation-date>${now}</meta:creation-date><dc:date>${now}</dc:date><dc:language>en-IN</dc:language>`
      + `<meta:user-defined meta:name="BTS.settings">${X(JSON.stringify({ page: D.page, design: Object.fromEntries(['theme', 'pspacing', 'styleSet', 'styleMods', 'customStyles', 'headFont', 'bodyFont', 'accent', 'dark', 'biblio', 'markup', 'watermark', 'border', 'pageColor', 'hdrFirst'].filter(k => D.design[k] !== undefined).map(k => [k, D.design[k]])) }))}</meta:user-defined>`
      + (pr.company ? `<meta:user-defined meta:name="Company">${X(pr.company)}</meta:user-defined>` : '') + (pr.category ? `<meta:user-defined meta:name="Category">${X(pr.category)}</meta:user-defined>` : '')
      + `<meta:document-statistic meta:page-count="${D.pageCount || 1}" meta:paragraph-count="${$$('p,h1,h2,h3,h4,h5,h6,li', D.doc).length}" meta:word-count="${words}" meta:image-count="${picFiles.length}" meta:table-count="${tblN}"/>`
      + '</office:meta></office:document-meta>';

    const manifest = `<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>${picFiles.map(p => `<manifest:file-entry manifest:full-path="${p.path}" manifest:media-type="${p.mime}"/>`).join('')}</manifest:manifest>`;

    const zip = new JSZip();
    zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' });
    const nf = { createFolders: false };
    zip.file('content.xml', contentXML, nf); zip.file('styles.xml', stylesXML, nf); zip.file('meta.xml', metaXML, nf);
    picFiles.forEach(p => zip.file(p.path, p.data, { compression: 'STORE', createFolders: false }));
    zip.file('META-INF/manifest.xml', manifest, nf);
    D._lastOdt = { contentXML, stylesXML, metaXML, manifest };
    return await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.oasis.opendocument.text' });
  }
})();
