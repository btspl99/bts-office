/* BTS Docs — opening files (.docx via mammoth + XML fidelity pass, .odt, .html, .md, .txt)
 * and simple exporters (.html "web page" with embedded settings, .txt, .md). */
(function () {
  const D = window.Docs; const { $, $$, PX } = D; const esc = BTS.esc;

  /* ================= sanitizing HTML ================= */
  const ALLOWED_TAGS = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'B', 'STRONG', 'I', 'EM', 'U', 'S', 'STRIKE', 'DEL', 'INS', 'SUB', 'SUP', 'SPAN', 'FONT', 'A', 'BR', 'HR', 'UL', 'OL', 'LI', 'TABLE', 'THEAD', 'TBODY', 'TFOOT', 'TR', 'TD', 'TH', 'COLGROUP', 'COL', 'CAPTION', 'IMG', 'BLOCKQUOTE', 'PRE', 'CODE', 'DIV', 'SECTION', 'ARTICLE', 'MARK', 'SMALL', 'BIG', 'CENTER', 'DL', 'DT', 'DD', 'FIGURE', 'FIGCAPTION', 'KBD', 'Q', 'CITE', 'ABBR', 'TT', 'HEADER', 'FOOTER', 'MAIN', 'ASIDE', 'NAV', 'ADDRESS']);
  const DROP_TAGS = new Set(['SCRIPT', 'STYLE', 'IFRAME', 'OBJECT', 'EMBED', 'LINK', 'META', 'NOSCRIPT', 'TEMPLATE', 'FORM', 'INPUT', 'BUTTON', 'SELECT', 'TEXTAREA', 'SVG', 'MATH', 'CANVAS', 'VIDEO', 'AUDIO', 'TITLE', 'HEAD', 'BASE', 'FRAME', 'FRAMESET', 'APPLET']);
  const ALLOWED_ATTR = new Set(['href', 'src', 'alt', 'title', 'style', 'class', 'colspan', 'rowspan', 'width', 'height', 'id', 'align', 'valign', 'color', 'face', 'size', 'start', 'type', 'span', 'contenteditable', 'data-eq', 'data-fn', 'data-text', 'data-cid', 'data-fld', 'data-style', 'data-target', 'data-cap', 'data-nw', 'data-nh', 'data-own-drop', 'data-src', 'data-page', 'data-seq', 'data-field', 'data-author', 'data-time', 'data-tcid', 'data-label', 'data-title', 'data-head', 'data-decorative', 'data-smartart']);
  D.sanitize = (html, { keepClasses = true } = {}) => {
    const doc = new DOMParser().parseFromString('<!doctype html><body>' + html, 'text/html');
    const walk = (el) => {
      [...el.children].forEach(c => {
        if (DROP_TAGS.has(c.tagName)) { c.remove(); return; }
        [...c.attributes].forEach(a => {
          const n = a.name.toLowerCase();
          if (n.startsWith('on') || !ALLOWED_ATTR.has(n) || (n === 'class' && !keepClasses)) c.removeAttribute(a.name);
          else if ((n === 'href' || n === 'src') && /^\s*(javascript|vbscript|data:text\/html)/i.test(a.value)) c.removeAttribute(a.name);
          else if (n === 'style' && /expression\(|url\(\s*['"]?\s*javascript/i.test(a.value)) c.removeAttribute(a.name);
          else if (n === 'contenteditable' && a.value !== 'false') c.removeAttribute(a.name);
        });
        if (c.tagName === 'IMG' && (!c.getAttribute('src') || !/^(data:image\/|https?:|blob:)/i.test(c.getAttribute('src')))) { c.remove(); return; }
        walk(c);
        if (!ALLOWED_TAGS.has(c.tagName)) { const p = c.parentNode; while (c.firstChild) p.insertBefore(c.firstChild, c); c.remove(); }
        else if (c.tagName === 'CENTER') { const d = doc.createElement('div'); d.style.textAlign = 'center'; while (c.firstChild) d.appendChild(c.firstChild); c.replaceWith(d); }
        else if (c.tagName === 'FONT') { const s = doc.createElement('span'); if (c.getAttribute('color')) s.style.color = c.getAttribute('color'); if (c.getAttribute('face')) s.style.fontFamily = c.getAttribute('face'); const sz = { 1: 8, 2: 10, 3: 12, 4: 14, 5: 18, 6: 24, 7: 36 }[c.getAttribute('size')]; if (sz) s.style.fontSize = sz + 'pt'; while (c.firstChild) s.appendChild(c.firstChild); c.replaceWith(s); }
        else if (c.tagName === 'A' && c.getAttribute('href')) c.removeAttribute('target');
      });
    };
    walk(doc.body);
    // Normalize semantic sectioning elements to divs and remove empty wrappers
    $$('section,article,header,footer,main,aside,nav,address,figure', doc.body).forEach(e => { const d = doc.createElement('div'); [...e.attributes].forEach(a => d.setAttribute(a.name, a.value)); while (e.firstChild) d.appendChild(e.firstChild); e.replaceWith(d); });
    $$('figcaption', doc.body).forEach(e => { const p = doc.createElement('p'); p.className = 'caption'; while (e.firstChild) p.appendChild(e.firstChild); e.replaceWith(p); });
    // unwrap plain divs that only hold blocks (keeps structure flat & editable)
    let changed = true; while (changed) { changed = false; $$('div', doc.body).forEach(d => { if (!d.className && !d.getAttribute('style') && D.hasBlockChild(d)) { while (d.firstChild) d.parentNode.insertBefore(d.firstChild, d); d.remove(); changed = true; } }); }
    $$('div', doc.body).forEach(d => { if (!d.className && !D.hasBlockChild(d)) { const p = doc.createElement('p'); [...d.attributes].forEach(a => p.setAttribute(a.name, a.value)); while (d.firstChild) p.appendChild(d.firstChild); d.replaceWith(p); } });
    $$('table', doc.body).forEach(t => { if (!/\btbl\b/.test(t.className)) t.classList.add('tbl', 't-plain'); });
    $$('td,th', doc.body).forEach(c => { if (!c.firstChild) c.innerHTML = '<p><br></p>'; });
    $$('p,h1,h2,h3,h4,h5,h6,li', doc.body).forEach(p => { if (!p.childNodes.length) p.innerHTML = '<br>'; });
    return doc.body.innerHTML;
  };

  /* ================= native (HTML + meta) ================= */
  function docCSS() {
    const out = [];
    for (const sh of document.styleSheets) { let rules; try { rules = sh.cssRules; } catch { continue; } if (!sh.href || !/docs\.css/.test(sh.href)) continue; for (const r of rules) { if (r.selectorText && /\.doc-content/.test(r.selectorText) && !/#doc|body\.|\.pv-|\.paper/.test(r.selectorText)) out.push(r.cssText); } }
    return out.join('\n');
  }
  D.serializeNative = () => {
    const meta = { v: 1, page: D.page, design: D.design, comments: D.comments, hdr: D.hdr.innerHTML, ftr: D.ftr.innerHTML };
    const ps = D.paperSize(); const m = D.page.m;
    const styleVars = D.doc.getAttribute('style') || '';
    return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="generator" content="BTS Docs">
<meta name="bts-docs" content="${esc(JSON.stringify(meta))}">
<title>${esc(BTS.title)}</title>
<style>
body { margin: 0; background: ${D.design.pageColor || '#f3f4f6'}; }
.page { max-width: ${(ps.w * 96).toFixed(0)}px; margin: 24px auto; background: ${D.design.pageColor || '#fff'}; padding: ${m.t}in ${m.r}in ${m.b}in ${m.l}in; box-shadow: 0 2px 14px rgba(0,0,0,.12); box-sizing: border-box; }
@media print { body { background: none; } .page { box-shadow: none; margin: 0; padding: 0; max-width: none; } @page { size: ${ps.w}in ${ps.h}in; margin: ${m.t}in ${m.r}in ${m.b}in ${m.l}in; } hr.page-break { break-after: page; visibility: hidden; } }
${docCSS()}
</style></head>
<body><div class="page doc-content" style="${esc(styleVars)}">
${D.detokenize(cleanHTML(D.doc))}
${$('#fn-area').innerHTML}
</div></body></html>`;
  };
  function cleanHTML(root) { const c = root.cloneNode(true); $$('.cell-sel', c).forEach(e => e.classList.remove('cell-sel')); $$('.pv-on', c).forEach(e => e.className = e.className.replace(/\bpv-\S+/g, '').trim()); $$('.cmt.active', c).forEach(e => e.classList.remove('active')); $$('hr.page-break', c).forEach(h => h.removeAttribute('style')); return c.innerHTML.replace(/​/g, ''); }
  D.cleanHTML = cleanHTML;
  D.loadNative = (html) => {
    const dp = new DOMParser().parseFromString(html, 'text/html');
    const metaEl = dp.querySelector('meta[name="bts-docs"]'); let meta = null;
    try { meta = metaEl && JSON.parse(metaEl.getAttribute('content')); } catch { }
    const body = dp.querySelector('.page.doc-content') || dp.body;
    $$('.footnotes', body).forEach(f => f.remove());
    if (meta) { D.page = Object.assign(D.defaultPage(), meta.page); D.design = Object.assign(D.defaultDesign(), meta.design); D.comments = meta.comments || {}; }
    else { D.page = D.defaultPage(); D.design = D.defaultDesign(); D.comments = {}; }
    D.setContent(D.sanitize(body.innerHTML), { hdr: meta ? D.sanitize(meta.hdr || '') : '', ftr: meta ? D.sanitize(meta.ftr || '') : '' });
  };

  /* ================= plain text & markdown ================= */
  D.textToHTML = (t) => t.replace(/\r\n?/g, '\n').split('\n').map(l => l.length ? '<p>' + esc(l).replace(/\t/g, '<span class="tab">\t</span>') + '</p>' : '<p><br></p>').join('').replace(/\f/g, '');
  D.mdToHTML = (md) => {
    const lines = md.replace(/\r\n?/g, '\n').split('\n'); const out = []; let i = 0;
    const inline = (s) => {
      const codes = []; s = s.replace(/`([^`]+)`/g, (m, c) => { codes.push(c); return '\u0000' + (codes.length - 1) + '\u0000'; });
      s = esc(s)
        .replace(/!\[([^\]]*)\]\(([^)\s]+)(?:\s+"[^"]*")?\)/g, (m, a, u) => /^(data:image|https?:)/.test(u) ? `<img alt="${a}" src="${u}">` : '')
        .replace(/\[([^\]]+)\]\(([^)\s]+)(?:\s+"[^"]*")?\)/g, (m, t, u) => `<a href="${/^(https?:|mailto:|#)/.test(u) ? u : 'https://' + u}">${t}</a>`)
        .replace(/\*\*\*([^*]+)\*\*\*/g, '<b><i>$1</i></b>').replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>').replace(/__([^_]+)__/g, '<b>$1</b>')
        .replace(/(^|[^*])\*([^*\s][^*]*)\*/g, '$1<i>$2</i>').replace(/(^|\W)_([^_\s][^_]*)_(?=\W|$)/g, '$1<i>$2</i>')
        .replace(/~~([^~]+)~~/g, '<s>$1</s>').replace(/==([^=]+)==/g, '<span style="background-color:#ffff00">$1</span>')
        .replace(/(^|\s)(https?:\/\/[^\s<]+)/g, '$1<a href="$2">$2</a>');
      return s.replace(/\u0000(\d+)\u0000/g, (m, k) => `<span style="font-family:Consolas,'Courier New',monospace;background-color:#f1f2f5">${esc(codes[+k])}</span>`);
    };
    const listRe = /^(\s*)([-*+]|\d+[.)])\s+(.*)$/;
    while (i < lines.length) {
      const l = lines[i];
      if (/^```/.test(l)) { const buf = []; i++; while (i < lines.length && !/^```/.test(lines[i])) buf.push(lines[i++]); i++; out.push('<pre>' + esc(buf.join('\n')) + '</pre>'); continue; }
      let m;
      if ((m = l.match(/^(#{1,6})\s+(.*?)\s*#*$/))) { const n = m[1].length; out.push(`<h${Math.min(n, 3) === n ? n : 3}>${inline(m[2])}</h${Math.min(n, 3) === n ? n : 3}>`); i++; continue; }
      if (/^\s*([-*_])\s*\1\s*\1[\s\1]*$/.test(l)) { out.push('<hr>'); i++; continue; }
      if (/^\s*>/.test(l)) { const buf = []; while (i < lines.length && /^\s*>/.test(lines[i])) buf.push(lines[i++].replace(/^\s*>\s?/, '')); out.push('<blockquote>' + D.mdToHTML(buf.join('\n')) + '</blockquote>'); continue; }
      if (/^\s*\|.*\|\s*$/.test(l) && i + 1 < lines.length && /^\s*\|?\s*:?-{2,}/.test(lines[i + 1])) {
        const row = (s) => s.trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim());
        const head = row(l); i += 2; const body = []; while (i < lines.length && /^\s*\|.*\|\s*$/.test(lines[i])) body.push(row(lines[i++]));
        out.push(`<table class="tbl t-header"><colgroup>${head.map(() => `<col style="width:${(100 / head.length).toFixed(2)}%">`).join('')}</colgroup><tbody><tr>${head.map(c => `<td><p>${inline(c)}</p></td>`).join('')}</tr>${body.map(r => `<tr>${head.map((_, k) => `<td><p>${inline(r[k] || '') || '<br>'}</p></td>`).join('')}</tr>`).join('')}</tbody></table>`); continue;
      }
      if (listRe.test(l)) {
        // nested lists by indentation
        const root = { children: [], ordered: /\d/.test(l.match(listRe)[2]) }; const stack = [{ indent: -1, node: root }];
        while (i < lines.length && (listRe.test(lines[i]) || (/^\s{2,}\S/.test(lines[i]) && stack.length > 1))) {
          const mm = lines[i].match(listRe);
          if (!mm) { const last = stack[stack.length - 1].node.children; if (last.length) last[last.length - 1].text += ' ' + lines[i].trim(); i++; continue; }
          const ind = mm[1].replace(/\t/g, '    ').length;
          if (ind === 0 && root.children.length && /\d/.test(mm[2]) !== root.ordered) break;
          while (stack.length > 1 && ind <= stack[stack.length - 1].indent) stack.pop();
          const parent = stack[stack.length - 1];
          let list = parent.node;
          if (parent.indent >= 0) { const lastItem = parent.node.children[parent.node.children.length - 1]; if (!lastItem) { i++; continue; } lastItem.sub = lastItem.sub || { children: [], ordered: /\d/.test(mm[2]) }; list = lastItem.sub; }
          const task = mm[3].match(/^\[([ xX])\]\s+(.*)$/);
          list.children.push({ text: task ? (task[1].trim() ? '☑ ' : '☐ ') + task[2] : mm[3] });
          stack.push({ indent: ind, node: list });
          i++;
        }
        const render = (lst) => `<${lst.ordered ? 'ol' : 'ul'}>${lst.children.map(it => `<li>${inline(it.text)}${it.sub ? render(it.sub) : ''}</li>`).join('')}</${lst.ordered ? 'ol' : 'ul'}>`;
        // flatten stack nodes: root holds top-level items
        out.push(render(root)); continue;
      }
      if (!l.trim()) { i++; continue; }
      const buf = [l]; i++;
      while (i < lines.length && lines[i].trim() && !/^(#{1,6}\s|```|\s*>|\s*([-*+]|\d+[.)])\s)/.test(lines[i]) && !/^\s*\|.*\|\s*$/.test(lines[i])) buf.push(lines[i++]);
      out.push('<p>' + buf.map(inline).join(/  $/.test(buf[0]) ? '<br>' : ' ') + '</p>');
    }
    return out.join('');
  };
  /* HTML → Markdown / text */
  D.toMarkdown = () => {
    const blk = []; const inl = (el) => {
      let s = '';
      el.childNodes.forEach(n => {
        if (n.nodeType === 3) { s += n.data.replace(/​/g, '').replace(/([*_`[\]])/g, '\\$1'); return; }
        if (n.nodeType !== 1) return; const t = n.tagName; const cs = n.style;
        if (n.classList.contains('fn-ref')) { s += `[^${n.textContent}]`; return; }
        if (t === 'BR') { s += '  \n'; return; }
        if (t === 'IMG') { s += `![${n.alt || ''}](${n.src.length > 4000 ? 'image' : n.src})`; return; }
        let inner = inl(n);
        const bold = t === 'B' || t === 'STRONG' || +cs.fontWeight >= 600 || cs.fontWeight === 'bold';
        const ital = t === 'I' || t === 'EM' || cs.fontStyle === 'italic';
        const strike = t === 'S' || t === 'STRIKE' || t === 'DEL' || /line-through/.test(cs.textDecoration);
        if (n.classList.contains('eq')) inner = n.dataset.eq || inner;
        if (strike && inner.trim()) inner = `~~${inner}~~`;
        if (ital && inner.trim()) inner = `*${inner}*`;
        if (bold && inner.trim()) inner = `**${inner}**`;
        if (t === 'A') inner = `[${inner}](${n.getAttribute('href')})`;
        s += inner;
      });
      return s;
    };
    const walk = (el, depth = 0, listType = null) => {
      [...el.children].forEach(c => {
        const t = c.tagName;
        if (/^H[1-6]$/.test(t)) blk.push('#'.repeat(+t[1]) + ' ' + inl(c));
        else if (t === 'P') { if (c.classList.contains('title')) blk.push('# ' + inl(c)); else if (c.classList.contains('subtitle')) blk.push('*' + inl(c) + '*'); else blk.push(inl(c)); }
        else if (t === 'PRE') blk.push('```\n' + c.textContent + '\n```');
        else if (t === 'BLOCKQUOTE') blk.push((D.hasBlockChild(c) ? [...c.children].map(inl).join('\n\n') : inl(c)).split('\n').map(l => '> ' + l).join('\n'));
        else if (t === 'HR') blk.push(c.classList.contains('page-break') ? '<div style="page-break-after:always"></div>' : '---');
        else if (t === 'UL' || t === 'OL') {
          const items = []; let k = 1;
          const rec = (list, d) => [...list.children].forEach(li => { const txt = inl(cloneWithout(li)); items.push('   '.repeat(d) + (list.tagName === 'OL' ? (k++) + '. ' : '- ') + txt); [...li.children].filter(x => x.tagName === 'UL' || x.tagName === 'OL').forEach(sub => rec(sub, d + 1)); });
          rec(c, 0); blk.push(items.join('\n'));
        }
        else if (t === 'TABLE') {
          const rows = [...c.rows].map(r => [...r.cells].map(td => inl(td).replace(/\|/g, '\\|').replace(/\n+/g, ' ').trim()));
          if (rows.length) { const w = Math.max(...rows.map(r => r.length)); rows.forEach(r => { while (r.length < w) r.push(''); }); blk.push(['| ' + rows[0].join(' | ') + ' |', '|' + ' --- |'.repeat(w), ...rows.slice(1).map(r => '| ' + r.join(' | ') + ' |')].join('\n')); }
        }
        else if (c.classList.contains('toc')) blk.push([...c.querySelectorAll('p[data-target]')].map(p => '   '.repeat(+(p.className.match(/toc-l(\d)/) || [, 1])[1] - 1) + '- ' + p.querySelector('.toc-t').textContent).join('\n'));
        else if (t === 'DIV') walk(c, depth);
        else blk.push(inl(c));
      });
    };
    function cloneWithout(li) { const c = li.cloneNode(true); $$('ul,ol', c).forEach(x => x.remove()); return c; }
    walk(D.doc);
    const refs = $$('sup.fn-ref', D.doc); if (refs.length) blk.push(refs.map((s, i) => `[^${i + 1}]: ${s.dataset.text || ''}`).join('\n'));
    return blk.filter(b => b !== undefined).join('\n\n').replace(/\n{3,}/g, '\n\n') + '\n';
  };
  D.toText = () => {
    const out = []; const txt = (el) => { const c = el.cloneNode(true); $$('br', c).forEach(b => b.replaceWith('\n')); $$('ul,ol', c).forEach(x => x.remove()); return c.textContent.replace(/​/g, ''); };
    const walk = (el, d = 0) => [...el.children].forEach(c => {
      const t = c.tagName;
      if (t === 'UL' || t === 'OL') { let k = 1; [...c.children].forEach(li => { out.push('    '.repeat(d) + (t === 'OL' ? (k++) + '. ' : '• ') + txt(li)); [...li.children].filter(x => x.tagName === 'UL' || x.tagName === 'OL').forEach(sub => walk({ children: [sub] }, d + 1)); }); }
      else if (t === 'TABLE') { [...c.rows].forEach(r => out.push([...r.cells].map(td => td.textContent.trim()).join('\t'))); out.push(''); }
      else if (t === 'HR') out.push(c.classList.contains('page-break') ? '\f' : '────────────────────');
      else if (c.classList && c.classList.contains('toc')) [...c.querySelectorAll('p[data-target]')].forEach(p => out.push(p.querySelector('.toc-t').textContent + '\t' + p.querySelector('.toc-n').textContent));
      else if (t === 'DIV' && D.hasBlockChild(c)) walk(c, d);
      else out.push(txt(c));
    });
    walk(D.doc);
    const refs = $$('sup.fn-ref', D.doc); if (refs.length) { out.push('', '────────'); refs.forEach((s, i) => out.push(`[${i + 1}] ${s.dataset.text || ''}`)); }
    return out.join('\r\n').replace(/(\r\n){3,}/g, '\r\n\r\n') + '\r\n';
  };

  /* ================= DOCX import ================= */
  const STYLE_MAP = [
    "p[style-name='Title'] => p.title:fresh", "p[style-name='Subtitle'] => p.subtitle:fresh",
    "p[style-name='Quote'] => blockquote > p:fresh", "p[style-name='Intense Quote'] => blockquote.intense > p:fresh",
    "p[style-name='No Spacing'] => p.nospace:fresh", "p[style-name='Caption'] => p.caption:fresh",
    "p[style-name='Code'] => pre:fresh", "p[style-name='HTML Preformatted'] => pre:fresh", "p[style-name='Source Code'] => pre:fresh",
    "p[style-name='TOC Heading'] => p.toc-head:fresh", "p[style-name='toc 1'] => p.toc-l1:fresh", "p[style-name='toc 2'] => p.toc-l2:fresh", "p[style-name='toc 3'] => p.toc-l3:fresh",
    "p[style-name='TOC 1'] => p.toc-l1:fresh", "p[style-name='TOC 2'] => p.toc-l2:fresh", "p[style-name='TOC 3'] => p.toc-l3:fresh",
    "p[style-name='Heading 4'] => h4:fresh", "p[style-name='Heading 5'] => h5:fresh", "p[style-name='Heading 6'] => h6:fresh",
    "p[style-name='table of figures'] => p.tof-l:fresh", "p[style-name='Table of Figures'] => p.tof-l:fresh",
    "r[style-name='Strong'] =>", "r[style-name='Emphasis'] =>", "u => u", "strike => s",
  ];
  const kids = (el, n) => el ? [...el.children].filter(c => c.nodeName === n) : [];
  const first = (el, n) => el ? [...el.children].find(c => c.nodeName === n) || null : null;
  const att = (el, n) => el ? el.getAttribute(n) : null;
  const HL = { yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000', white: '#ffffff' };

  function parseRPr(rPr) {
    if (!rPr) return null; const o = {};
    const c = att(first(rPr, 'w:color'), 'w:val'); if (c && c !== 'auto' && /^[0-9a-f]{6}$/i.test(c)) o.color = '#' + c;
    const sz = att(first(rPr, 'w:sz'), 'w:val'); if (sz) o.size = +sz / 2;
    const f = first(rPr, 'w:rFonts'); const fn = att(f, 'w:ascii') || att(f, 'w:hAnsi') || att(f, 'w:cs'); if (fn) o.font = fn;
    const h = att(first(rPr, 'w:highlight'), 'w:val'); if (h && HL[h]) o.bg = HL[h];
    const shd = att(first(rPr, 'w:shd'), 'w:fill'); if (!o.bg && shd && shd !== 'auto' && /^[0-9a-f]{6}$/i.test(shd)) o.bg = '#' + shd;
    const caps = first(rPr, 'w:caps'); if (caps && att(caps, 'w:val') !== '0' && att(caps, 'w:val') !== 'false') o.caps = true;
    return Object.keys(o).length ? o : null;
  }
  /** Walk a w:p collecting runs (text with props) and comment boundaries by character offset. */
  function readPara(p, ctx) {
    const runs = []; let text = ''; const dels = [];
    const visit = (el, tc) => {
      for (const c of el.children) {
        const n = c.nodeName;
        if (n === 'w:r') {
          const rPr = first(c, 'w:rPr'); const props = parseRPr(rPr); let t = '';
          for (const k of c.children) { if (k.nodeName === 'w:t') t += k.textContent; else if (k.nodeName === 'w:tab') t += '\t'; else if (k.nodeName === 'w:noBreakHyphen') t += '‑'; }
          const rs = att(first(rPr, 'w:rStyle'), 'w:val');
          if (t) { runs.push({ start: text.length, end: text.length + t.length, props, rs, tc }); text += t; }
        } else if (n === 'w:ins' || n === 'w:moveTo') visit(c, { kind: 'ins', author: att(c, 'w:author') || 'Unknown', date: att(c, 'w:date') || '', id: 'x' + (att(c, 'w:id') || Math.random().toString(36).slice(2, 7)) });
        else if (n === 'w:del' || n === 'w:moveFrom') {
          let t = ''; const rp = []; [...c.getElementsByTagName('w:r')].forEach(r => { for (const k of r.children) { if (k.nodeName === 'w:delText' || k.nodeName === 'w:t') t += k.textContent; else if (k.nodeName === 'w:tab') t += '\t'; } if (!rp.length) rp.push(parseRPr(first(r, 'w:rPr'))); });
          if (t) dels.push({ at: text.length, text: t, author: att(c, 'w:author') || 'Unknown', date: att(c, 'w:date') || '', id: 'x' + (att(c, 'w:id') || ''), props: rp[0] });
        } else if (n === 'w:hyperlink' || n === 'w:smartTag' || n === 'w:fldSimple' || n === 'w:sdt' || n === 'w:sdtContent' || n === 'w:customXml') visit(c, tc);
        else if (n === 'w:commentRangeStart') ctx.cStart[att(c, 'w:id')] = { p: ctx.pi, o: text.length };
        else if (n === 'w:commentRangeEnd') ctx.cEnd[att(c, 'w:id')] = { p: ctx.pi, o: text.length };
      }
    };
    visit(p, null);
    const pPr = first(p, 'w:pPr');
    const numPr = first(pPr, 'w:numPr');
    return { text, runs, dels, pPr, num: numPr ? { id: att(first(numPr, 'w:numId'), 'w:val'), lvl: +(att(first(numPr, 'w:ilvl'), 'w:val') || 0) } : null, style: att(first(pPr, 'w:pStyle'), 'w:val') };
  }
  function paraCSS(pPr, isList) {
    if (!pPr) return {}; const s = {};
    const jc = att(first(pPr, 'w:jc'), 'w:val'); if (jc) s.textAlign = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[jc] || '';
    const sp = first(pPr, 'w:spacing');
    if (sp) { const b = att(sp, 'w:before'), a = att(sp, 'w:after'), l = att(sp, 'w:line'), lr = att(sp, 'w:lineRule'); if (b != null && att(sp, 'w:beforeAutospacing') !== '1') s.marginTop = (+b / 20) + 'pt'; if (a != null && att(sp, 'w:afterAutospacing') !== '1') s.marginBottom = (+a / 20) + 'pt'; if (l && (!lr || lr === 'auto')) s.lineHeight = String(Math.round(+l / 240 * 100) / 100); else if (l && (lr === 'exact' || lr === 'atLeast')) s.lineHeight = (+l / 20) + 'pt'; }
    const ind = first(pPr, 'w:ind');
    if (ind && !isList) { const L = att(ind, 'w:left') ?? att(ind, 'w:start'), R = att(ind, 'w:right') ?? att(ind, 'w:end'), F = att(ind, 'w:firstLine'), H = att(ind, 'w:hanging'); if (L && +L) s.marginLeft = (+L / 1440).toFixed(2) + 'in'; if (R && +R) s.marginRight = (+R / 1440).toFixed(2) + 'in'; if (F && +F) s.textIndent = (+F / 1440).toFixed(2) + 'in'; if (H && +H) { s.textIndent = (-H / 1440).toFixed(2) + 'in'; s.paddingLeft = ''; } }
    const shd = att(first(pPr, 'w:shd'), 'w:fill'); if (shd && shd !== 'auto' && /^[0-9a-f]{6}$/i.test(shd)) s.backgroundColor = '#' + shd;
    const bdr = first(pPr, 'w:pBdr'); if (bdr) { ['top', 'bottom', 'left', 'right'].forEach(side => { const b = first(bdr, 'w:' + side); const css = borderFromXML(b); if (css) s['border' + side[0].toUpperCase() + side.slice(1)] = css; }); }
    if (s.backgroundColor || s.borderTop || s.borderBottom || s.borderLeft) s.padding = '2pt 4pt';
    if (first(pPr, 'w:pageBreakBefore') && att(first(pPr, 'w:pageBreakBefore'), 'w:val') !== '0') s._pbb = true;
    return s;
  }
  function borderFromXML(b) {
    if (!b) return null; const v = att(b, 'w:val'); if (!v || v === 'nil' || v === 'none') return null;
    const w = Math.max(0.5, (+att(b, 'w:sz') || 4) / 8 * 4 / 3); const c = att(b, 'w:color'); const col = c && c !== 'auto' ? '#' + c : '#000';
    const st = /double/.test(v) ? 'double' : /dash/.test(v) ? 'dashed' : /dot/.test(v) ? 'dotted' : 'solid';
    return `${st === 'double' ? Math.max(3, w).toFixed(1) : w.toFixed(1)}px ${st} ${col}`;
  }
  /** Text nodes that make up the visible text of an HTML paragraph element (skipping nested lists & footnote refs). */
  function paraTextNodes(el) {
    const out = []; const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, { acceptNode: n => { let p = n.parentElement; while (p && p !== el) { if (p.tagName === 'UL' || p.tagName === 'OL' || p.classList.contains('fn-ref') || (p.tagName === 'SUP' && p.querySelector('a[href^="#footnote"]'))) return NodeFilter.FILTER_REJECT; p = p.parentElement; } return NodeFilter.FILTER_ACCEPT; } });
    let n; while ((n = w.nextNode())) out.push(n); return out;
  }
  function wrapOffsets(el, start, end, make) {
    const nodes = paraTextNodes(el); let pos = 0; const segs = [];
    nodes.forEach(n => { const a = pos, b = pos + n.length; pos = b; const s = Math.max(start, a), e = Math.min(end, b); if (e > s) segs.push({ n, s: s - a, e: e - a }); });
    segs.forEach(({ n, s, e }) => { let t = n; if (s > 0) t = t.splitText(s); if (e - s < t.length) t.splitText(e - s); const w = make(); t.parentNode.insertBefore(w, t); w.appendChild(t); });
  }
  const normT = (s) => s.replace(/ /g, ' ').replace(/\s+/g, ' ').trim();

  const CHAR_IDS = { Emphasis: 'emphasis', Strong: 'strong', SubtleEmphasis: 'subtle-em', IntenseEmphasis: 'intense-em', SubtleReference: 'subtle-ref', IntenseReference: 'intense-ref', BookTitle: 'book-title' };
  /** BTS Docs data stored in a custom XML part (styles, sources, protection, citation list…). */
  async function readBtsData(zip) {
    for (const name of Object.keys(zip.files).filter(n => /^customXml\/item\d+\.xml$/i.test(n))) {
      try { const x = await zip.file(name).async('string'); if (!/btsDocs/.test(x)) continue; const d = new DOMParser().parseFromString(x, 'application/xml'); const root = d.documentElement; if (root.localName !== 'btsDocs') continue; return JSON.parse(root.textContent); } catch (e) { console.warn('bts data', e); }
    }
    return null;
  }
  async function readProps(zip) {
    const out = {}; const txt = async (n) => { const f = zip.file(n); return f ? new DOMParser().parseFromString(await f.async('string'), 'application/xml') : null; };
    const core = await txt('docProps/core.xml');
    if (core) { const g = (tag) => { const e = core.getElementsByTagName(tag)[0]; return e ? e.textContent.trim() : ''; }; Object.assign(out, { title: g('dc:title'), subject: g('dc:subject'), author: g('dc:creator'), keywords: g('cp:keywords'), comments: g('dc:description'), category: g('cp:category'), status: g('cp:contentStatus') }); }
    const app = await txt('docProps/app.xml'); if (app) { const c = app.getElementsByTagName('Company')[0]; if (c) out.company = c.textContent.trim(); }
    const cust = await txt('docProps/custom.xml'); if (cust) { [...cust.getElementsByTagName('property')].forEach(pr => { if (pr.getAttribute('name') === '_MarkAsFinal' && /true|1/i.test(pr.textContent)) out.markFinal = true; }); }
    return out;
  }
  async function readSettings(zip) {
    const f = zip.file('word/settings.xml'); if (!f) return {};
    const x = new DOMParser().parseFromString(await f.async('string'), 'application/xml'); const out = {};
    const tr = x.getElementsByTagName('w:trackRevisions')[0]; if (tr && !/^(0|false|off)$/.test(att(tr, 'w:val') || '')) out.track = true;
    const dp = x.getElementsByTagName('w:documentProtection')[0]; if (dp && /^(1|true|on)$/.test(att(dp, 'w:enforcement') || '') && att(dp, 'w:edit') === 'readOnly') out.readOnly = true;
    return out;
  }
  /** Word documents (not made by BTS Docs): take theme fonts and built-in style formatting from styles.xml. */
  async function readWordStyles(zip, design) {
    try {
      const th = zip.file('word/theme/theme1.xml');
      if (th) { const t = new DOMParser().parseFromString(await th.async('string'), 'application/xml'); const maj = t.getElementsByTagName('a:majorFont')[0], min = t.getElementsByTagName('a:minorFont')[0]; const lat = (e) => e && e.getElementsByTagName('a:latin')[0] && e.getElementsByTagName('a:latin')[0].getAttribute('typeface'); if (lat(maj)) design.headFont = lat(maj); if (lat(min)) design.bodyFont = lat(min); const clr = (n) => { const e = t.getElementsByTagName('a:' + n)[0]; const c = e && (e.getElementsByTagName('a:srgbClr')[0] || null); return c ? '#' + c.getAttribute('val').toUpperCase() : null; }; if (clr('accent1')) design.accent = clr('accent1'); if (clr('dk2')) design.dark = clr('dk2'); }
      const f = zip.file('word/styles.xml'); if (!f) return;
      const x = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
      const ID = { Normal: 'normal', Title: 'title', Subtitle: 'subtitle', Heading1: 'h1', Heading2: 'h2', Heading3: 'h3', Heading4: 'h4', Heading5: 'h5', Heading6: 'h6', Quote: 'quote', IntenseQuote: 'iquote', Caption: 'caption', NoSpacing: 'nospace' };
      const mods = {};
      const readR = (rPr, o) => {
        if (!rPr) return; const rf = first(rPr, 'w:rFonts');
        if (rf) { const at = att(rf, 'w:asciiTheme'); if (at) o.font = /major/.test(at) ? '@head' : '@body'; else if (att(rf, 'w:ascii')) o.font = att(rf, 'w:ascii'); }
        const sz = att(first(rPr, 'w:sz'), 'w:val'); if (sz) o.size = +sz / 2;
        const c = first(rPr, 'w:color'); if (c) { const tc = att(c, 'w:themeColor'); if (tc && /accent1/.test(tc)) o.color = '@accent'; else if (tc && /text2|dark2|tx2/.test(tc)) o.color = '@dark'; else if (att(c, 'w:val') && /^[0-9a-f]{6}$/i.test(att(c, 'w:val'))) o.color = '#' + att(c, 'w:val').toUpperCase(); }
        const on = (n) => { const e = first(rPr, n); return e ? !/^(0|false)$/.test(att(e, 'w:val') || '') : undefined; };
        if (on('w:b') !== undefined) o.bold = on('w:b'); if (on('w:i') !== undefined) o.italic = on('w:i'); if (on('w:caps') !== undefined) o.caps = on('w:caps');
      };
      const readP = (pPr, o) => { if (!pPr) return; const sp = first(pPr, 'w:spacing'); if (sp) { if (att(sp, 'w:before') != null) o.before = +att(sp, 'w:before') / 20; if (att(sp, 'w:after') != null) o.after = +att(sp, 'w:after') / 20; const l = att(sp, 'w:line'); if (l && (!att(sp, 'w:lineRule') || att(sp, 'w:lineRule') === 'auto')) o.line = Math.round(+l / 240 * 100) / 100; } const jc = att(first(pPr, 'w:jc'), 'w:val'); if (jc) o.align = { center: 'center', right: 'right', both: 'justify' }[jc] || 'left'; };
      const dd = x.getElementsByTagName('w:docDefaults')[0];
      if (dd) { const o = {}; readR(first(first(dd, 'w:rPrDefault'), 'w:rPr'), o); readP(first(first(dd, 'w:pPrDefault'), 'w:pPr'), o); if (o.font === '@head') o.font = '@body'; if (o.after != null || o.line != null) design.pspacing = { after: o.after ?? 0, line: o.line ?? 1, before: o.before || 0 }; delete o.before; delete o.after; delete o.line; mods.normal = o; }
      [...x.getElementsByTagName('w:style')].forEach(st => {
        const id = ID[att(st, 'w:styleId')]; if (!id) return;
        const o = {}; readR(first(st, 'w:rPr'), o); readP(first(st, 'w:pPr'), o);
        if (id === 'normal') { mods.normal = Object.assign(mods.normal || {}, o); if (o.after != null || o.line != null) { design.pspacing = Object.assign({ before: 0, after: 8, line: 1.15 }, design.pspacing || {}, o.after != null ? { after: o.after } : {}, o.line != null ? { line: o.line } : {}); delete mods.normal.after; delete mods.normal.line; delete mods.normal.before; } return; }
        if (Object.keys(o).length) mods[id] = o;
      });
      if (mods.normal && !Object.keys(mods.normal).length) delete mods.normal;
      design.styleMods = mods; design.styleSet = 'classic';
    } catch (e) { console.warn('word styles', e); }
  }

  D.importDocx = async (bytes) => {
    const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    let zip = null; try { zip = await JSZip.loadAsync(ab); } catch (e) { console.warn('zip', e); }
    const bts = zip ? await readBtsData(zip) : null;
    const props = zip ? await readProps(zip) : {};
    const settings = zip ? await readSettings(zip) : {};
    const styleMap = STYLE_MAP.slice();
    ((bts && bts.design && bts.design.customStyles) || []).forEach(u => { if (u.type === 'c') styleMap.push(`r[style-name='${u.name.replace(/'/g, '')}'] =>`); else styleMap.push(`p[style-name='${u.name.replace(/'/g, '')}'] => p.us-${u.id}:fresh`); });
    styleMap.push("r[style-name='BTS Citation'] =>", "r[style-name='Subtle Emphasis'] =>", "r[style-name='Intense Emphasis'] =>", "r[style-name='Subtle Reference'] =>", "r[style-name='Intense Reference'] =>", "r[style-name='Book Title'] =>");
    // mammoth drops simple fields (SEQ, MERGEFIELD …); keep their cached result text as ordinary runs
    let mab = ab;
    if (zip && zip.file('word/document.xml')) {
      const dx = await zip.file('word/document.xml').async('string');
      if (/<w:fldSimple\b/.test(dx)) { zip.file('word/document.xml', dx.replace(/<w:fldSimple\b[^>]*\/>/g, '').replace(/<w:fldSimple\b[^>]*>/g, '').replace(/<\/w:fldSimple>/g, '')); mab = await zip.generateAsync({ type: 'arraybuffer' }); }
    }
    const res = await mammoth.convertToHtml({ arrayBuffer: mab }, {
      styleMap, ignoreEmptyParagraphs: false, includeDefaultStyleMap: true,
      convertImage: mammoth.images.imgElement(img => img.read('base64').then(b => ({ src: 'data:' + (img.contentType || 'image/png') + ';base64,' + b }))),
    });
    const holder = document.createElement('div'); holder.innerHTML = res.value;
    let xml = null; const meta = { page: D.defaultPage(), design: D.defaultDesign(), comments: {}, hdr: '', ftr: '' };
    if (bts && bts.design) { const keep = ['theme', 'pspacing', 'styleSet', 'styleMods', 'customStyles', 'headFont', 'bodyFont', 'accent', 'dark', 'biblio', 'markup', 'protect']; keep.forEach(k => { if (bts.design[k] !== undefined) meta.design[k] = bts.design[k]; }); }
    else if (zip) await readWordStyles(zip, meta.design);
    meta.design.track = !!settings.track;
    if (props.markFinal || /^final$/i.test(props.status || '')) meta.design.protect = { mode: 'final' };
    else if (settings.readOnly) meta.design.protect = { mode: 'readonly', hash: bts && bts.design && bts.design.protect ? bts.design.protect.hash || null : null };
    else meta.design.protect = null;
    const P = Object.assign({}, props); delete P.status; delete P.markFinal; BTS.setProps(P);
    try { if (zip) xml = new DOMParser().parseFromString(await zip.file('word/document.xml').async('string'), 'application/xml'); } catch (e) { console.warn('docx xml pass skipped', e); }
    // footnotes from mammoth → our fn-ref sups
    $$('sup > a[href^="#footnote-"]', holder).forEach(a => {
      const id = a.getAttribute('href').slice(1); const li = holder.querySelector('li#' + CSS.escape(id));
      let text = ''; if (li) { const c = li.cloneNode(true); $$('a[href^="#footnote-ref"]', c).forEach(x => x.remove()); text = c.textContent.trim(); }
      const sup = a.parentElement; const s = document.createElement('sup'); s.className = 'fn-ref'; s.contentEditable = 'false'; s.dataset.fn = D.uid('fn'); s.dataset.text = text; s.textContent = '0'; sup.replaceWith(s);
    });
    $$('li[id^="footnote-"], li[id^="endnote-"]', holder).forEach(li => { const ol = li.parentElement; li.remove(); if (ol && !ol.children.length) ol.remove(); });
    $$('a[id^="footnote-ref"], a[id^="endnote-ref"]', holder).forEach(a => D.unwrap(a));
    if (xml) { try { await fidelityPass(holder, xml, zip, meta); } catch (e) { console.warn('docx fidelity pass failed', e); } }
    // tabs → tab spans, empties
    const tw = document.createTreeWalker(holder, NodeFilter.SHOW_TEXT); const tabs = []; let n; while ((n = tw.nextNode())) if (n.data.includes('\t') && !n.parentElement.closest('pre')) tabs.push(n);
    tabs.forEach(n => { const parts = n.data.split('\t'); const frag = document.createDocumentFragment(); parts.forEach((p, i) => { if (i) { const s = document.createElement('span'); s.className = 'tab'; s.textContent = '\t'; frag.appendChild(s); } if (p) frag.appendChild(document.createTextNode(p)); }); n.replaceWith(frag); });
    // merge consecutive pre (one per paragraph from mammoth)
    $$('pre', holder).forEach(p => { let nx = p.nextElementSibling; while (nx && nx.tagName === 'PRE') { p.append('\n', ...nx.childNodes); const r = nx; nx = nx.nextElementSibling; r.remove(); } });
    $$('blockquote', holder).forEach(b => { if (!holder.contains(b)) return; let nx = b.nextElementSibling; while (nx && nx.tagName === 'BLOCKQUOTE' && nx.className === b.className) { b.append(...nx.childNodes); const r = nx; nx = nx.nextElementSibling; r.remove(); } });
    // table of figures paragraphs → live block
    let hadTof = false;
    $$('p.tof-l', holder).forEach(p0 => {
      if (!holder.contains(p0) || (p0.previousElementSibling && p0.previousElementSibling.classList.contains('tof-l'))) return;
      const group = []; let x = p0; while (x && x.tagName === 'P' && x.classList.contains('tof-l')) { group.push(x); x = x.nextElementSibling; }
      const t = document.createElement('div'); t.className = 'toc tof'; t.contentEditable = 'false';
      const m = group.map(g => g.textContent.trim().match(/^(Figure|Table|Equation)\b/i)).find(Boolean); t.dataset.label = m ? m[1][0].toUpperCase() + m[1].slice(1).toLowerCase() : 'Figure';
      const head = p0.previousElementSibling; if (head && (head.classList.contains('toc-head') || /^(table of figures|list of (figures|tables|equations))$/i.test(head.textContent.trim()))) head.remove();
      p0.before(t); group.forEach(g => g.remove()); hadTof = true;
    });
    // citations, bibliography, merge fields
    const cites = (bts && bts.cites) || [];
    $$('span.cite-imp', holder).forEach((sp, i) => { const c = cites[i]; if (!c) { D.unwrap(sp); return; } sp.className = 'cite'; sp.setAttribute('contenteditable', 'false'); sp.dataset.src = c.src; if (c.page) sp.dataset.page = c.page; });
    const hasSources = !!(meta.design.biblio && meta.design.biblio.sources && meta.design.biblio.sources.length);
    $$('p.bib-entry', holder).forEach(p0 => {
      if (!holder.contains(p0) || !hasSources || (p0.previousElementSibling && p0.previousElementSibling.classList.contains('bib-entry'))) return;
      const b = document.createElement('div'); b.className = 'biblio'; b.setAttribute('contenteditable', 'false');
      const head = p0.previousElementSibling; if (head && /^H[1-3]$/.test(head.tagName) && /^(references|works cited|bibliography|reference list|sources)$/i.test(head.textContent.trim())) { if (/^bibliography$/i.test(head.textContent.trim())) b.dataset.title = 'Bibliography'; b.dataset.head = head.tagName.toLowerCase(); head.remove(); }
      p0.before(b); let x = b.nextElementSibling; while (x && x.tagName === 'P' && x.classList.contains('bib-entry')) { const nx = x.nextElementSibling; x.remove(); x = nx; }
    });
    { const tw2 = document.createTreeWalker(holder, NodeFilter.SHOW_TEXT); const hits = []; let t; while ((t = tw2.nextNode())) if (/«[^«»\n]{1,60}»/.test(t.data)) hits.push(t);
      hits.forEach(t => { const frag = document.createDocumentFragment(); let last = 0; const re = /«([^«»\n]{1,60})»/g; let m; while ((m = re.exec(t.data))) { if (m.index > last) frag.appendChild(document.createTextNode(t.data.slice(last, m.index))); const sp = document.createElement('span'); sp.className = 'mf'; sp.setAttribute('contenteditable', 'false'); sp.dataset.field = m[1]; sp.textContent = m[0]; frag.appendChild(sp); last = m.index + m[0].length; } if (last < t.data.length) frag.appendChild(document.createTextNode(t.data.slice(last))); t.replaceWith(frag); }); }
    // TOC paragraphs → live TOC block
    const tocPs = $$('p.toc-l1, p.toc-l2, p.toc-l3', holder); let hadToc = false;
    if (tocPs.length) {
      const toc = document.createElement('div'); toc.className = 'toc'; toc.contentEditable = 'false';
      const head = tocPs[0].previousElementSibling; if (head && (head.classList.contains('toc-head') || /^(contents|table of contents)$/i.test(head.textContent.trim()))) head.remove();
      tocPs[0].before(toc); tocPs.forEach(p => p.remove()); hadToc = true;
    }
    $$('p.toc-head', holder).forEach(p => p.classList.remove('toc-head'));
    $$('table', holder).forEach(t => { t.classList.add('tbl'); if (!/\bt-/.test(t.className)) t.classList.add('t-plain'); });
    $$('td,th', holder).forEach(c => { if (!c.firstElementChild) { const p = document.createElement('p'); while (c.firstChild) p.appendChild(c.firstChild); if (!p.childNodes.length) p.innerHTML = '<br>'; c.appendChild(p); } });
    $$('p,h1,h2,h3,h4,h5,h6,li', holder).forEach(p => { if (!p.childNodes.length || (!p.textContent && !p.querySelector('img,br,.fn-ref'))) p.innerHTML = '<br>'; });
    // drop mammoth's empty bookmark anchors that aren't link targets
    const targets = new Set($$('a[href^="#"]', holder).map(a => a.getAttribute('href').slice(1)));
    $$('a[id]:not([href])', holder).forEach(a => { if (!a.textContent) { const h = a.closest('h1,h2,h3,h4,h5,h6,p'); if (h && targets.has(a.id) && !h.id) h.id = a.id; a.remove(); } });
    if (meta.cols > 1) { const d = document.createElement('div'); d.className = 'cols cols-' + meta.cols; while (holder.firstChild) d.appendChild(holder.firstChild); holder.appendChild(d); }
    D.page = meta.page; D.design = meta.design; D.comments = meta.comments;
    D.setContent(D.sanitize(holder.innerHTML), { hdr: meta.hdr, ftr: meta.ftr });
    const needFields = hadToc || hadTof || $('.cite,.biblio,.seq', D.doc);
    if (needFields) setTimeout(() => { D.updateFields(true); D.clearHistory(); BTS.markClean(); }, 120);
    return res.messages;
  };

  async function fidelityPass(holder, xml, zip, meta) {
    const body = xml.getElementsByTagName('w:body')[0]; if (!body) return;
    const numbering = await readNumbering(zip);
    // --- paragraphs (document order) ---
    const ctx = { cStart: {}, cEnd: {}, pi: 0 };
    const xps = [...body.getElementsByTagName('w:p')].filter(p => !p.parentElement.closest || true);
    const xparas = xps.map((p, i) => { ctx.pi = i; const r = readPara(p, ctx); r.inTxbx = !!closestXML(p, 'w:txbxContent'); return r; });
    const hparas = $$('p,h1,h2,h3,h4,h5,h6,li,pre', holder).filter(el => !(el.tagName === 'P' && el.parentElement.closest('li') && false));
    const htext = (el) => paraTextNodes(el).map(n => n.data).join('');
    // align sequences by text
    const map = new Map(); let j = 0;
    for (let i = 0; i < xparas.length && j < hparas.length; i++) {
      const xt = normT(xparas[i].text);
      if (normT(htext(hparas[j])) === xt) { map.set(i, j); j++; continue; }
      let found = -1; for (let k = 1; k <= 4 && j + k < hparas.length; k++) if (normT(htext(hparas[j + k])) === xt && xt) { found = j + k; break; }
      if (found >= 0) { map.set(i, found); j = found + 1; continue; }
      // xml paragraph missing in html (e.g. text box) — skip it
    }
    map.forEach((hj, xi) => {
      const xp = xparas[xi]; const el = hparas[hj];
      const css = paraCSS(xp.pPr, el.tagName === 'LI');
      if (css._pbb) { const hr = document.createElement('hr'); hr.className = 'page-break'; const top = el.closest('li,table') || el; (top.tagName === 'LI' ? top.closest('ul,ol') : top).before(hr); delete css._pbb; }
      Object.entries(css).forEach(([k, v]) => { if (v) el.style[k] = v; });
      // runs
      const htxt = htext(el);
      if (el.tagName === 'P' && xp.style) {
        const cls = xp.style === 'ListParagraph' && !xp.num ? 'listpara' : xp.style === 'Bibliography' ? 'bib-entry' : /^TableofFigures$/i.test(xp.style) ? 'tof-l' : /^BTSU/.test(xp.style) ? 'us-' + xp.style.slice(4) : null;
        if (cls) el.classList.add(cls);
      }
      if (htxt === xp.text) {
        xp.runs.forEach(r => { if (!r.props) return; const p = r.props; wrapOffsets(el, r.start, r.end, () => { const s = document.createElement('span'); if (p.color) s.style.color = p.color; if (p.size) s.style.fontSize = p.size + 'pt'; if (p.font) s.style.fontFamily = D.fontStack(p.font); if (p.bg) s.style.backgroundColor = p.bg; if (p.caps) s.style.textTransform = 'uppercase'; return s; }); });
        xp.runs.forEach(r => { if (!r.tc) return; wrapOffsets(el, r.start, r.end, () => D.makeTC('ins', r.tc.id, r.tc.author, r.tc.date || undefined)); });
        xp.runs.forEach(r => {
          if (!r.rs) return; let cls = null;
          if (r.rs === 'BTSCitation') cls = 'cite-imp'; else if (CHAR_IDS[r.rs]) cls = 'cs-' + CHAR_IDS[r.rs]; else if (/^BTSU/.test(r.rs)) cls = 'cs-' + r.rs.slice(4);
          if (cls) wrapOffsets(el, r.start, r.end, () => { const s = document.createElement('span'); s.className = cls; return s; });
        });
        xp._el = el;
      }
      // list style from numbering definitions
      if (el.tagName === 'LI' && xp.num && numbering) {
        const lst = el.parentElement; const fmt = numbering(xp.num.id, xp.num.lvl);
        if (fmt && !lst.dataset.fmtSet) { lst.dataset.fmtSet = '1'; if (fmt !== (lst.tagName === 'OL' ? 'decimal' : 'disc')) lst.style.listStyleType = fmt; }
      }
    });
    $$('[data-fmt-set]', holder).forEach(l => l.removeAttribute('data-fmt-set'));
    // page breaks inside runs (w:br type=page) → hr.page-break before the next paragraph
    xps.forEach((p, i) => {
      const br = [...p.getElementsByTagName('w:br')].find(b => att(b, 'w:type') === 'page'); if (!br) return;
      const hj = map.get(i); if (hj == null) return; const el = hparas[hj];
      const hr = document.createElement('hr'); hr.className = 'page-break';
      const top = el.tagName === 'LI' ? el.closest('ul,ol') : (el.closest('table') || el);
      const txt = normT(xparas[i].text);
      if (!txt) { if (top === el) el.replaceWith(hr); else top.after(hr); } else { (isBrAtStart(p, br) ? top.before(hr) : top.after(hr)); }
      // mammoth renders the break as <br>; remove trailing/leading br
      $$('br', el).forEach(b => { if (!b.nextSibling || !b.previousSibling) b.remove(); });
    });
    // --- comments ---
    try {
      const cf = zip.file('word/comments.xml');
      if (cf) {
        const cx = new DOMParser().parseFromString(await cf.async('string'), 'application/xml');
        [...cx.getElementsByTagName('w:comment')].forEach(c => {
          const id = att(c, 'w:id'); const s = ctx.cStart[id], e = ctx.cEnd[id]; if (!s || !e) return;
          const cid = D.uid('c'); const text = [...c.getElementsByTagName('w:p')].map(p => [...p.getElementsByTagName('w:t')].map(t => t.textContent).join('')).join('\n');
          let quote = '';
          for (let pi = s.p; pi <= e.p; pi++) {
            const xp = xparas[pi]; if (!xp || !xp._el) continue;
            const a = pi === s.p ? s.o : 0, b = pi === e.p ? e.o : xp.text.length; if (b <= a) continue;
            quote += xp.text.slice(a, b) + ' ';
            wrapOffsets(xp._el, a, b, () => { const sp = document.createElement('span'); sp.className = 'cmt'; sp.dataset.cid = cid; return sp; });
          }
          meta.comments[cid] = { text, author: att(c, 'w:author') || 'Author', time: Date.parse(att(c, 'w:date') || '') || Date.now(), quote: quote.trim().slice(0, 80) };
        });
      }
    } catch (e) { console.warn('comments', e); }
    // --- tracked deletions: put the deleted text back as <del> at its offset ---
    xparas.forEach(xp => {
      if (!xp._el || !xp.dels.length) return; const el = xp._el;
      [...xp.dels].reverse().forEach(dl => {
        const d = D.makeTC('del', dl.id, dl.author, dl.date || undefined); d.textContent = dl.text;
        if (dl.props && (dl.props.color || dl.props.font)) { const sp = document.createElement('span'); if (dl.props.color) sp.style.color = dl.props.color; if (dl.props.font) sp.style.fontFamily = D.fontStack(dl.props.font); sp.textContent = dl.text; d.textContent = ''; d.appendChild(sp); }
        const nodes = paraTextNodes(el); let pos = 0, placed = false;
        for (const n of nodes) { if (dl.at <= pos + n.length) { const off = dl.at - pos; if (off === 0) n.before(d); else if (off === n.length) n.after(d); else { const rest = n.splitText(off); rest.before(d); } placed = true; break; } pos += n.length; }
        if (!placed) { const br = [...el.childNodes].find(c => c.nodeName === 'BR' && !c.nextSibling); if (br) br.before(d); else el.appendChild(d); }
      });
    });
    // --- images: sizes & wrapping ---
    const drawings = [...body.getElementsByTagName('w:drawing'), ...body.getElementsByTagName('v:imagedata')].filter(d => d.nodeName === 'v:imagedata' || d.getElementsByTagName('a:blip').length);
    drawings.sort((a, b) => a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1);
    const imgs = $$('img', holder);
    if (drawings.length === imgs.length) drawings.forEach((d, i) => {
      const img = imgs[i];
      if (d.nodeName === 'v:imagedata') { const sh = d.parentElement; const st = att(sh, 'style') || ''; const w = st.match(/width:\s*([\d.]+)pt/), h = st.match(/height:\s*([\d.]+)pt/); if (w) img.style.width = Math.round(+w[1] * 4 / 3) + 'px'; if (h) img.style.height = Math.round(+h[1] * 4 / 3) + 'px'; return; }
      const ext = d.getElementsByTagName('wp:extent')[0]; if (ext) { img.style.width = Math.round(+att(ext, 'cx') / 9525) + 'px'; img.style.height = Math.round(+att(ext, 'cy') / 9525) + 'px'; }
      const anchor = first(d, 'wp:anchor');
      if (anchor) {
        const align = (anchor.getElementsByTagName('wp:align')[0] || {}).textContent || '';
        const wrapNone = first(anchor, 'wp:wrapNone') || first(anchor, 'wp:wrapTopAndBottom');
        if (first(anchor, 'wp:wrapTopAndBottom')) img.classList.add('img-center');
        else if (!wrapNone) img.classList.add(align === 'right' ? 'img-right' : 'img-left');
      }
      const dp = d.getElementsByTagName('wp:docPr')[0]; if (dp && att(dp, 'descr') && !img.alt) img.alt = att(dp, 'descr');
    });
    // --- tables ---
    const xt = [...body.getElementsByTagName('w:tbl')]; const ht = $$('table', holder);
    if (xt.length === ht.length) xt.forEach((tb, i) => {
      const t = ht[i];
      const cols = kids(first(tb, 'w:tblGrid'), 'w:gridCol').map(g => +att(g, 'w:w') || 0); const sum = cols.reduce((a, b) => a + b, 0);
      if (cols.length && sum) { const cg = document.createElement('colgroup'); cols.forEach(w => { const c = document.createElement('col'); c.style.width = (w / sum * 100).toFixed(3) + '%'; cg.appendChild(c); }); t.prepend(cg); }
      const tPr = first(tb, 'w:tblPr'); const tw = first(tPr, 'w:tblW');
      if (tw && att(tw, 'w:type') === 'dxa' && sum) { const contentW = (D.paperSize().w - meta.page.m.l - meta.page.m.r) * 1440; if (+att(tw, 'w:w') < contentW * 0.9) { t.style.width = Math.round(+att(tw, 'w:w') / 15) + 'px'; } }
      const tb_b = first(tPr, 'w:tblBorders');
      const xrows = kids(tb, 'w:tr'); const hrows = [...t.rows]; let anyCellBorders = false;
      xrows.forEach((xr, ri) => {
        const hr = hrows[ri]; if (!hr) return;
        const xcells = kids(xr, 'w:tc').filter(tc => { const vm = first(first(tc, 'w:tcPr'), 'w:vMerge'); return !vm || att(vm, 'w:val') === 'restart'; });
        if (xcells.length !== hr.cells.length) return;
        xcells.forEach((tc, ci) => {
          const cell = hr.cells[ci]; const pr = first(tc, 'w:tcPr'); if (!pr) return;
          const shd = att(first(pr, 'w:shd'), 'w:fill'); if (shd && shd !== 'auto' && /^[0-9a-f]{6}$/i.test(shd)) cell.style.backgroundColor = '#' + shd;
          const va = att(first(pr, 'w:vAlign'), 'w:val'); if (va) cell.style.verticalAlign = va === 'center' ? 'middle' : va;
          const bd = first(pr, 'w:tcBorders');
          if (bd) { anyCellBorders = true; ['top', 'bottom', 'left', 'right'].forEach(side => { const b = first(bd, 'w:' + side) || first(bd, 'w:' + (side === 'left' ? 'start' : side === 'right' ? 'end' : side)); if (!b) return; const css = borderFromXML(b); cell.style['border' + side[0].toUpperCase() + side.slice(1)] = css || '0 none'; }); }
        });
      });
      const noTblBorders = !tb_b || ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'].every(s => { const b = first(tb_b, 'w:' + s); return !b || ['nil', 'none'].includes(att(b, 'w:val')); });
      const styled = att(first(tPr, 'w:tblStyle'), 'w:val');
      if (noTblBorders && !anyCellBorders && !(styled && /Grid/i.test(styled))) t.classList.add('no-borders');
    });
    // --- section: page size, margins, columns, borders, header/footer ---
    const sect = [...body.children].reverse().find(c => c.nodeName === 'w:sectPr');
    if (sect) {
      const sz = first(sect, 'w:pgSz'); const mg = first(sect, 'w:pgMar');
      if (sz) {
        let w = +att(sz, 'w:w') / 1440, h = +att(sz, 'w:h') / 1440; const land = att(sz, 'w:orient') === 'landscape' || w > h;
        if (w > h) [w, h] = [h, w];
        let best = 'A4', bd = 1e9; Object.entries(D.PAPER).forEach(([k, p]) => { const d = Math.abs(p.w - w) + Math.abs(p.h - h); if (d < bd) { bd = d; best = k; } });
        meta.page.size = best; meta.page.orient = land ? 'landscape' : 'portrait';
      }
      if (mg) { const g = (k, d) => { const v = att(mg, 'w:' + k); return v != null ? Math.max(0, +v / 1440) : d; }; meta.page.m = { t: g('top', 1), b: g('bottom', 1), l: g('left', 1), r: g('right', 1) }; meta.page.hdrDist = g('header', 0.5); meta.page.ftrDist = g('footer', 0.5); }
      const cols = first(sect, 'w:cols'); if (cols && +att(cols, 'w:num') > 1) meta.cols = Math.min(3, +att(cols, 'w:num'));
      const pb = first(sect, 'w:pgBorders'); if (pb) { const tp = first(pb, 'w:top'); if (tp && !['nil', 'none'].includes(att(tp, 'w:val'))) { const v = att(tp, 'w:val'); meta.design.border = { style: /double/.test(v) ? 'double' : /dash/.test(v) ? 'dashed' : /dot/.test(v) ? 'dotted' : 'solid', color: att(tp, 'w:color') && att(tp, 'w:color') !== 'auto' ? '#' + att(tp, 'w:color') : '#1f2533', width: Math.max(1, (+att(tp, 'w:sz') || 8) / 8 * 4 / 3) }; } }
      if (first(sect, 'w:titlePg')) meta.design.hdrFirst = true;
      const rels = await readRels(zip);
      for (const kind of ['header', 'footer']) {
        const ref = kids(sect, 'w:' + kind + 'Reference').find(r => att(r, 'w:type') === 'default') || kids(sect, 'w:' + kind + 'Reference')[0];
        if (!ref) continue; const target = rels[att(ref, 'r:id')]; if (!target) continue;
        const f = zip.file('word/' + target.replace(/^\/?word\//, '')); if (!f) continue;
        const hx = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
        meta[kind === 'header' ? 'hdr' : 'ftr'] = await hfToHTML(hx, zip, target, meta);
      }
    }
    const bg = xml.getElementsByTagName('w:background')[0]; if (bg && att(bg, 'w:color') && att(bg, 'w:color') !== 'auto') meta.design.pageColor = '#' + att(bg, 'w:color');
  }
  function closestXML(el, name) { let p = el.parentNode; while (p && p.nodeName !== '#document') { if (p.nodeName === name) return p; p = p.parentNode; } return null; }
  function isBrAtStart(p, br) { const r = closestXML(br, 'w:r'); let prev = r && r.previousElementSibling; while (prev) { if (prev.getElementsByTagName('w:t').length) return false; prev = prev.previousElementSibling; } return true; }
  async function readRels(zip) {
    const out = {}; const f = zip.file('word/_rels/document.xml.rels'); if (!f) return out;
    const x = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
    [...x.getElementsByTagName('Relationship')].forEach(r => out[r.getAttribute('Id')] = r.getAttribute('Target'));
    return out;
  }
  async function readNumbering(zip) {
    const f = zip.file('word/numbering.xml'); if (!f) return null;
    const x = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
    const abs = {}; [...x.getElementsByTagName('w:abstractNum')].forEach(a => abs[att(a, 'w:abstractNumId')] = a);
    const nums = {}; [...x.getElementsByTagName('w:num')].forEach(n => nums[att(n, 'w:numId')] = att(first(n, 'w:abstractNumId'), 'w:val'));
    return (numId, lvl) => {
      const a = abs[nums[numId]]; if (!a) return null;
      const L = kids(a, 'w:lvl').find(l => +att(l, 'w:ilvl') === lvl); if (!L) return null;
      const fmt = att(first(L, 'w:numFmt'), 'w:val'); const txt = att(first(L, 'w:lvlText'), 'w:val') || '';
      if (fmt === 'bullet') return /[◦o○]/.test(txt) ? 'circle' : /[▪■]/.test(txt) ? 'square' : 'disc';
      return { decimal: 'decimal', lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha', lowerRoman: 'lower-roman', upperRoman: 'upper-roman', decimalZero: 'decimal-leading-zero' }[fmt] || null;
    };
  }
  async function hfToHTML(hx, zip, target, meta) {
    const out = [];
    for (const p of hx.getElementsByTagName('w:p')) {
      if (closestXML(p, 'w:txbxContent')) continue;
      // watermark image we wrote ourselves
      const dp = [...p.getElementsByTagName('wp:docPr')].find(d => att(d, 'name') === 'BTS Watermark');
      if (dp) { try { meta.design.watermark = JSON.parse(att(dp, 'descr')); } catch { } }
      let html = ''; let inField = null; let fieldShown = false;
      const runHTML = (r) => {
        let t = ''; for (const k of r.children) { if (k.nodeName === 'w:t') t += esc(k.textContent); else if (k.nodeName === 'w:tab') t += '<span class="tab">\t</span>'; }
        const rp = first(r, 'w:rPr'); if (!t) return '';
        const b = first(rp, 'w:b') && att(first(rp, 'w:b'), 'w:val') !== '0', i = first(rp, 'w:i') && att(first(rp, 'w:i'), 'w:val') !== '0';
        const pr = parseRPr(rp) || {}; let st = ''; if (pr.color) st += `color:${pr.color};`; if (pr.size) st += `font-size:${pr.size}pt;`; if (pr.font) st += `font-family:${pr.font};`;
        if (b) t = `<b>${t}</b>`; if (i) t = `<i>${t}</i>`; return st ? `<span style="${st}">${t}</span>` : t;
      };
      const fld = (instr) => /\bNUMPAGES\b/.test(instr) ? '<span class="fld" data-fld="pages" contenteditable="false">1</span>' : /\bPAGE\b/.test(instr) ? '<span class="fld" data-fld="page" contenteditable="false">1</span>' : null;
      const walk = (el) => {
        for (const c of el.children) {
          if (c.nodeName === 'w:fldSimple') { const f = fld(att(c, 'w:instr') || ''); html += f || [...c.getElementsByTagName('w:r')].map(runHTML).join(''); }
          else if (c.nodeName === 'w:r') {
            if (!first(c, 'w:fldChar') && !first(c, 'w:instrText')) { if (inField != null && fieldShown) continue; html += runHTML(c); continue; }
            let pending = false;
            const flushField = () => { if (pending) return; const f = fld(inField || ''); if (f) { html += f; fieldShown = true; } pending = true; };
            for (const k of c.children) {
              if (k.nodeName === 'w:fldChar') { const ty = att(k, 'w:fldCharType'); if (ty === 'begin') { inField = ''; fieldShown = false; pending = false; } else if (ty === 'separate') flushField(); else if (ty === 'end') { if (inField != null && !fieldShown) flushField(); inField = null; fieldShown = false; } }
              else if (k.nodeName === 'w:instrText') inField = (inField || '') + k.textContent;
            }
          } else if (['w:hyperlink', 'w:sdt', 'w:sdtContent', 'w:smartTag', 'w:ins'].includes(c.nodeName)) walk(c);
        }
      };
      walk(p);
      const css = paraCSS(first(p, 'w:pPr'), false); delete css._pbb;
      const st = Object.entries(css).filter(([, v]) => v).map(([k, v]) => k.replace(/[A-Z]/g, m => '-' + m.toLowerCase()) + ':' + v).join(';');
      if (html || out.length) out.push(`<p${st ? ` style="${st}"` : ''}>${html || '<br>'}</p>`);
    }
    while (out.length && out[out.length - 1].endsWith('><br></p>')) out.pop();
    return out.join('');
  }

  /* ================= ODT import ================= */
  D.importOdt = async (bytes) => {
    const zip = await JSZip.loadAsync(bytes);
    const parse = async (n) => { const f = zip.file(n); return f ? new DOMParser().parseFromString(await f.async('string'), 'application/xml') : null; };
    const x = await parse('content.xml'); if (!x) throw new Error('This isn’t a valid .odt file.');
    const sx = await parse('styles.xml'), mx = await parse('meta.xml');
    const att = (e, n) => (e && e.getAttribute(n)) || '';
    const kids = (e, n) => e ? [...e.children].filter(c => c.nodeName === n) : [];
    const first = (e, n) => e ? e.getElementsByTagName(n)[0] || null : null;
    const styles = {}, lists = {}, sections = {};
    const NAMED_TAG = { Title: 'title', Subtitle: 'subtitle', Quotations: 'quote', Quote: 'quote', Intense_20_Quote: 'quote', Preformatted_20_Text: 'pre', Source_20_Text: 'pre', Caption: 'caption', No_20_Spacing: 'nospace', Footnote: 'footnote' };
    const readStyles = (doc, isAuto) => {
      if (!doc) return;
      [...doc.getElementsByTagName('style:style')].forEach(s => {
        const name = att(s, 'style:name'); const tp = first(s, 'style:text-properties'), pp = first(s, 'style:paragraph-properties'), cp = first(s, 'style:table-cell-properties'), gp = first(s, 'style:graphic-properties'), colp = first(s, 'style:table-column-properties'), sp = first(s, 'style:section-properties');
        const auto = isAuto || (s.parentNode && s.parentNode.nodeName === 'office:automatic-styles');
        let css = '', pcss = '';
        if (tp) {
          const g = (n) => att(tp, n);
          if (g('fo:font-weight') === 'bold' || +g('fo:font-weight') >= 600) css += 'font-weight:bold;'; else if (g('fo:font-weight') === 'normal' && auto) css += 'font-weight:normal;';
          if (g('fo:font-style') === 'italic') css += 'font-style:italic;';
          const deco = []; if (g('style:text-underline-style') && g('style:text-underline-style') !== 'none') deco.push('underline'); if (g('style:text-line-through-style') && g('style:text-line-through-style') !== 'none') deco.push('line-through'); if (deco.length) css += 'text-decoration:' + deco.join(' ') + ';';
          if (g('fo:color')) css += 'color:' + g('fo:color') + ';';
          if (/pt$/.test(g('fo:font-size'))) css += 'font-size:' + g('fo:font-size') + ';';
          if (g('fo:background-color') && g('fo:background-color') !== 'transparent') css += 'background-color:' + g('fo:background-color') + ';';
          const fn = g('style:font-name') || g('fo:font-family'); if (fn) css += 'font-family:' + D.fontStack(fn.replace(/^['"]|['"]$/g, '')).replace(/"/g, "'") + ';';
          const pos = g('style:text-position'); if (/^super|^\d/.test(pos) && !/^0/.test(pos)) css += 'vertical-align:super;font-size:smaller;'; else if (/^sub|^-/.test(pos)) css += 'vertical-align:sub;font-size:smaller;';
          if (g('fo:text-transform') === 'uppercase') css += 'text-transform:uppercase;';
          if (g('fo:font-variant') === 'small-caps') css += 'font-variant:small-caps;';
        }
        if (pp) {
          const g = (n) => att(pp, n);
          if (g('fo:text-align')) pcss += 'text-align:' + ({ start: 'left', end: 'right', left: 'left', right: 'right', center: 'center', justify: 'justify' }[g('fo:text-align')] || 'left') + ';';
          if (auto) {
            if (g('fo:margin-left') && parseFloat(g('fo:margin-left'))) pcss += 'margin-left:' + g('fo:margin-left') + ';';
            if (g('fo:margin-right') && parseFloat(g('fo:margin-right'))) pcss += 'margin-right:' + g('fo:margin-right') + ';';
            if (g('fo:text-indent') && parseFloat(g('fo:text-indent'))) pcss += 'text-indent:' + g('fo:text-indent') + ';';
            if (g('fo:margin-top')) pcss += 'margin-top:' + g('fo:margin-top') + ';';
            if (g('fo:margin-bottom')) pcss += 'margin-bottom:' + g('fo:margin-bottom') + ';';
            if (/%$/.test(g('fo:line-height'))) pcss += 'line-height:' + (parseFloat(g('fo:line-height')) / 100) + ';';
            if (g('fo:background-color') && g('fo:background-color') !== 'transparent') pcss += 'background-color:' + g('fo:background-color') + ';';
            ['top', 'bottom', 'left', 'right'].forEach(sd => { const b = g('fo:border-' + sd) || g('fo:border'); if (b && b !== 'none') pcss += `border-${sd}:${b.replace(/([\d.]+)pt/, (m, v) => Math.max(1, Math.round(v * 4 / 3)) + 'px')};`; });
            if (pcss.includes('border') || pcss.includes('background')) pcss += 'padding:2pt 4pt;';
          }
        }
        let cell = '';
        if (cp) {
          const g = (n) => att(cp, n);
          if (g('fo:background-color') && g('fo:background-color') !== 'transparent') cell += 'background-color:' + g('fo:background-color') + ';';
          ['top', 'bottom', 'left', 'right'].forEach(sd => { const b = g('fo:border-' + sd) || g('fo:border'); if (b) cell += `border-${sd}:${b === 'none' ? 'none' : b.replace(/([\d.]+)pt/, (m, v) => Math.max(1, Math.round(v * 4 / 3)) + 'px')};`; });
          if (g('style:vertical-align') === 'middle') cell += 'vertical-align:middle;'; else if (g('style:vertical-align') === 'bottom') cell += 'vertical-align:bottom;';
        }
        const parent = att(s, 'style:parent-style-name');
        styles[name] = { css, pcss, cell, parent, auto, family: att(s, 'style:family'), breakBefore: pp && att(pp, 'fo:break-before') === 'page', listStyle: att(s, 'style:list-style-name'), hpos: gp ? att(gp, 'style:horizontal-pos') : '', wrap: gp ? att(gp, 'style:wrap') : '', colWidth: colp ? att(colp, 'style:column-width') : '', cols: sp ? +(att(first(sp, 'style:columns'), 'fo:column-count') || 1) : 0, outline: att(s, 'style:default-outline-level') };
      });
      [...doc.getElementsByTagName('text:list-style')].forEach(ls => {
        const lv = [];
        [...ls.children].forEach(l => { const i = (+att(l, 'text:level') || 1) - 1; if (l.nodeName === 'text:list-level-style-number') lv[i] = { num: true, fmt: att(l, 'style:num-format') || '1', start: +att(l, 'text:start-value') || 1, pre: att(l, 'style:num-prefix'), suf: att(l, 'style:num-suffix') }; else if (l.nodeName === 'text:list-level-style-bullet') lv[i] = { num: false, ch: att(l, 'text:bullet-char') }; });
        lists[att(ls, 'style:name')] = lv;
      });
    };
    readStyles(sx, false); readStyles(x, true);
    const named = (n) => { for (let i = 0; n && i < 12; i++) { if (NAMED_TAG[n]) return NAMED_TAG[n]; const h = /^Heading_20_(\d)$/.exec(n); if (h) return 'h' + h[1]; n = styles[n] && styles[n].parent; } return ''; };
    const css = (n) => { let out = ''; const chain = []; for (let i = 0; n && styles[n] && i < 12; i++) { if (!styles[n].auto) break; chain.unshift(styles[n]); n = styles[n].parent; } chain.forEach(s => { out += s.css; }); return out; };
    const pcssOf = (n) => { let out = ''; for (let i = 0; n && styles[n] && i < 12; i++) { if (styles[n].auto) out = styles[n].pcss + out; n = styles[n].parent; } return out; };
    const LT = { '1': 'decimal', a: 'lower-alpha', A: 'upper-alpha', i: 'lower-roman', I: 'upper-roman' };
    const BUL = { '•': 'disc', '●': 'disc', '◦': 'circle', '○': 'circle', '▪': 'square', '■': 'square', '–': '"– "', '-': '"- "' };
    // tracked changes & BTS settings
    const regions = {};
    [...x.getElementsByTagName('text:changed-region')].forEach(r => {
      const id = att(r, 'text:id') || att(r, 'xml:id'); const k = r.firstElementChild; if (!k) return;
      const ci = first(k, 'office:change-info'); regions[id] = { kind: k.nodeName === 'text:insertion' ? 'ins' : k.nodeName === 'text:deletion' ? 'del' : 'fmt', author: ci ? (first(ci, 'dc:creator') || {}).textContent || '' : '', time: ci ? (first(ci, 'dc:date') || {}).textContent || '' : '', text: k.nodeName === 'text:deletion' ? kids(k, 'text:p').map(p => p.textContent).join(' ') : '' };
    });
    const ud = {}; if (mx) [...mx.getElementsByTagName('meta:user-defined')].forEach(u => { ud[att(u, 'meta:name')] = u.textContent; });
    let bts = null; try { bts = ud['BTS.settings'] ? JSON.parse(ud['BTS.settings']) : null; } catch { }
    const pics = {}; const comments = {}; let fnN = 0, tcN = 0;
    const tcTag = (id) => { const r = regions[id] || {}; return `data-author="${esc(r.author || 'Unknown')}" data-time="${esc(r.time ? new Date(r.time + (/Z|[+-]\d\d:?\d\d$/.test(r.time) ? '' : 'Z')).toISOString() : new Date().toISOString())}" data-tcid="t${esc(id)}"`; };
    const conv = async (node, ctx = {}) => {
      let out = '';
      for (const c of node.childNodes) {
        if (c.nodeType === 3) { out += esc(c.data); continue; }
        if (c.nodeType !== 1) continue;
        const n = c.nodeName; const sn = att(c, 'text:style-name');
        if (n === 'text:h' || n === 'text:p') {
          const tag = n === 'text:h' ? 'h' + Math.max(1, Math.min(6, +att(c, 'text:outline-level') || +((styles[sn] || {}).outline) || 1)) : named(sn);
          const st = (styles[sn] || {}); let pre = ''; for (let s2 = sn, i = 0; s2 && styles[s2] && i < 12; i++, s2 = styles[s2].parent) if (styles[s2].breakBefore) { pre = '<hr class="page-break">'; break; }
          if (ctx.inList && !pre && st) pre = '';
          const bm = [...c.children].find(k => k.nodeName === 'text:bookmark' || k.nodeName === 'text:bookmark-start');
          const id = bm && /^[A-Za-z][\w-]*$/.test(att(bm, 'text:name')) ? ` id="${esc(att(bm, 'text:name'))}"` : '';
          const pc = pcssOf(sn), tc = css(sn); const style = pc + tc ? ` style="${esc(pc + tc)}"` : '';
          let inner = await conv(c, ctx);
          if (ctx.raw) { out += inner; continue; }
          if (/^h\d$/.test(tag)) out += `${pre}<${tag}${id}${style}>${inner || '<br>'}</${tag}>`;
          else if (tag === 'pre') out += `${pre}<pre${id}>${inner}</pre>`;
          else if (tag === 'quote') out += `${pre}<blockquote><p${id}${style}>${inner || '<br>'}</p></blockquote>`;
          else out += `${pre}<p${id}${['title', 'subtitle', 'caption', 'nospace'].includes(tag) ? ` class="${tag}"` : ''}${style}>${inner || '<br>'}</p>`;
        }
        else if (n === 'text:span') { const s = css(sn); const inner = await conv(c, ctx); out += s ? `<span style="${esc(s)}">${inner}</span>` : inner; }
        else if (n === 'text:a') out += `<a href="${esc(att(c, 'xlink:href'))}">${await conv(c, ctx)}</a>`;
        else if (n === 'text:s') out += '&nbsp;'.repeat(Math.max(0, (+att(c, 'text:c') || 1) - 1)) + ' ';
        else if (n === 'text:tab') out += '<span class="tab">\t</span>';
        else if (n === 'text:line-break') out += '<br>';
        else if (n === 'text:page-number') out += '<span class="fld" data-fld="page" contenteditable="false">1</span>';
        else if (n === 'text:page-count') out += '<span class="fld" data-fld="pages" contenteditable="false">1</span>';
        else if (n === 'text:note') {
          const body = first(c, 'text:note-body'); const text = body ? kids(body, 'text:p').map(p => p.textContent).join('\n').trim() : '';
          out += `<sup class="fn-ref" data-fn="fo${++fnN}" contenteditable="false" data-text="${esc(text)}">${fnN}</sup>`;
        }
        else if (n === 'office:annotation') {
          const nm = att(c, 'office:name'); const txt = kids(c, 'text:p').map(p => p.textContent).join('\n');
          const id = 'c' + (nm || Math.random().toString(36).slice(2, 8)).replace(/[^\w-]/g, '');
          comments[id] = { text: txt, author: (first(c, 'dc:creator') || {}).textContent || 'Unknown', time: Date.parse((first(c, 'dc:date') || {}).textContent || '') || Date.now(), quote: '', name: nm };
          if (nm) out += `\u0001cs:${id}\u0001`;
        }
        else if (n === 'office:annotation-end') { const nm = att(c, 'office:name'); const id = Object.keys(comments).find(k => comments[k].name === nm); if (id) out += `\u0001ce:${id}\u0001`; }
        else if (n === 'text:change-start') out += `\u0001ts:${att(c, 'text:change-id')}\u0001`;
        else if (n === 'text:change-end') out += `\u0001te:${att(c, 'text:change-id')}\u0001`;
        else if (n === 'text:change') { const r = regions[att(c, 'text:change-id')]; if (r && r.kind === 'del' && r.text) out += `<del class="tc-del" ${tcTag(att(c, 'text:change-id'))}>${esc(r.text)}</del>`; }
        else if (n === 'text:list') {
          const lsName = sn || ctx.listStyle || ''; const lv = (lists[lsName] || [])[ctx.level || 0];
          const ordered = lv ? lv.num : (/Numbering|num/i.test(lsName) && !/bullet/i.test(lsName));
          const items = [];
          for (const li of c.children) if (li.nodeName === 'text:list-item' || li.nodeName === 'text:list-header') items.push(`<li>${(await conv(li, { ...ctx, listStyle: lsName, level: (ctx.level || 0) + 1, inList: true })).replace(/^<p( style="[^"]*")?>([\s\S]*?)<\/p>/, '$2')}</li>`);
          const type = lv ? (lv.num ? (lv.pre === '(' ? 'paren-' : '') + (LT[lv.fmt] || 'decimal') : (BUL[lv.ch] || '')) : '';
          const defType = ordered ? ['decimal', 'lower-alpha', 'lower-roman'][(ctx.level || 0) % 3] : ['disc', 'circle', 'square'][(ctx.level || 0) % 3];
          const st = type && type !== defType ? ` style="list-style-type:${esc(type)}"` : '';
          const start = lv && lv.num && lv.start > 1 ? ` start="${lv.start}"` : '';
          out += `<${ordered ? 'ol' : 'ul'}${start}${st}>${items.join('')}</${ordered ? 'ol' : 'ul'}>`;
        }
        else if (n === 'table:table') {
          const colsW = []; const colEls = []; [...c.children].forEach(k => { if (k.nodeName === 'table:table-column') colEls.push(k); else if (/^table:table-(header-)?columns$/.test(k.nodeName)) colEls.push(...kids(k, 'table:table-column')); });
          for (const col of colEls) { const cw = (styles[att(col, 'table:style-name')] || {}).colWidth || ''; const w = parseFloat(cw) || 0; const unit = cw.replace(/[\d.]/g, ''); const rep = Math.min(64, +att(col, 'table:number-columns-repeated') || 1); for (let i = 0; i < rep; i++) colsW.push(unit === 'in' ? w * 2.54 : unit === 'mm' ? w / 10 : unit === 'pt' ? w / 72 * 2.54 : w); }
          const total = colsW.reduce((a, b) => a + b, 0);
          const rows = [];
          const rowEls = [...c.getElementsByTagName('table:table-row')].filter(r => { let p = r.parentNode; while (p && p.nodeName !== 'table:table') p = p.parentNode; return p === c; });
          for (const r of rowEls) {
            const cells = [];
            for (const cell of r.children) if (cell.nodeName === 'table:table-cell') {
              const cs = att(cell, 'table:number-columns-spanned'), rs = att(cell, 'table:number-rows-spanned'); const st = (styles[att(cell, 'table:style-name')] || {}).cell || '';
              cells.push(`<td${cs > 1 ? ` colspan="${cs}"` : ''}${rs > 1 ? ` rowspan="${rs}"` : ''}${st ? ` style="${esc(st)}"` : ''}>${await conv(cell, ctx) || '<p><br></p>'}</td>`);
            }
            rows.push('<tr>' + cells.join('') + '</tr>');
          }
          const cg = total > 0 ? '<colgroup>' + colsW.map(w => `<col style="width:${(w / total * 100).toFixed(2)}%">`).join('') + '</colgroup>' : '';
          const tst = styles[att(c, 'table:style-name')] || {};
          out += `${tst.breakBefore ? '<hr class="page-break">' : ''}<table class="tbl t-plain">${cg}<tbody>${rows.join('')}</tbody></table>`;
        }
        else if (n === 'draw:frame') {
          const img = first(c, 'draw:image');
          if (img) {
            const href = att(img, 'xlink:href'); const f = href && zip.file(href.replace(/^\.\//, ''));
            if (f) {
              if (!pics[href]) { const b = await f.async('uint8array'); pics[href] = await BTS.bytesToDataURL(b, D.sniffImage(b)); }
              const w = D.toPx(att(c, 'svg:width')), h = D.toPx(att(c, 'svg:height')); const gs = styles[att(c, 'draw:style-name')] || {};
              const cls = att(c, 'text:anchor-type') !== 'as-char' && gs.wrap && gs.wrap !== 'none' ? (gs.hpos === 'right' ? 'img-right' : gs.hpos === 'left' ? 'img-left' : '') : '';
              const alt = ((first(c, 'svg:desc') || first(c, 'svg:title') || {}).textContent || '');
              out += `<img src="${pics[href]}"${cls ? ` class="${cls}"` : ''}${alt ? ` alt="${esc(alt)}"` : ''}${w ? ` style="width:${Math.round(w)}px;height:${Math.round(h)}px"` : ''}>`;
            }
          } else { const tb = first(c, 'draw:text-box'); if (tb) out += await conv(tb, { ...ctx, raw: true }); }
        }
        else if (n === 'text:section') { const st = styles[sn] || {}; const inner = await conv(c, ctx); out += st.cols > 1 ? `<div class="cols cols-${Math.min(3, st.cols)}">${inner}</div>` : inner; }
        else if (n === 'text:soft-page-break' || n === 'text:sequence-decls' || n === 'office:forms' || n === 'text:tracked-changes' || n === 'text:bookmark' || n === 'text:bookmark-start' || n === 'text:bookmark-end' || n === 'text:table-of-content-source' || n === 'svg:desc' || n === 'svg:title') { }
        else out += await conv(c, ctx);
      }
      return out;
    };
    // markers → comment / insertion spans when start and end are in the same paragraph
    const finish = (h) => h
      .replace(/\u0001cs:([\w-]+)\u0001([^\u0001]*?)\u0001ce:\1\u0001/g, (m, id, inner) => /<\/?(p|h\d|li|td|table|ul|ol|blockquote|pre)\b/.test(inner) ? inner : `<span class="cmt" data-cid="${id}">${inner}</span>`)
      .replace(/\u0001ts:([\w-]+)\u0001([^\u0001]*?)\u0001te:\1\u0001/g, (m, id, inner) => /<\/?(p|h\d|li|td|table|ul|ol|blockquote|pre)\b/.test(inner) ? inner : `<ins class="tc-ins" ${tcTag(id)}>${inner}</ins>`)
      .replace(/\u0001[^\u0001]*\u0001/g, '');
    const bodyEl = x.getElementsByTagName('office:text')[0];
    D.page = D.defaultPage(); D.design = D.defaultDesign(); D.comments = {};
    // page layout of the first master page
    let hdr = '', ftr = '';
    if (sx) {
      const mp = sx.getElementsByTagName('style:master-page')[0];
      if (mp) {
        const pl = [...sx.getElementsByTagName('style:page-layout')].find(l => att(l, 'style:name') === att(mp, 'style:page-layout-name'));
        const pp = pl && first(pl, 'style:page-layout-properties');
        if (pp) {
          const inch = (v) => (D.toPx(v) || 0) / PX;
          const w = inch(att(pp, 'fo:page-width')), h = inch(att(pp, 'fo:page-height'));
          const land = att(pp, 'style:print-orientation') === 'landscape' || w > h; const pw = Math.min(w, h), ph = Math.max(w, h);
          const size = Object.keys(D.PAPER).find(k => Math.abs(D.PAPER[k].w - pw) < 0.08 && Math.abs(D.PAPER[k].h - ph) < 0.08) || 'A4';
          const hp = first(first(pl, 'style:header-style'), 'style:header-properties'), fp = first(first(pl, 'style:footer-style'), 'style:footer-properties');
          const padI = (sd) => inch(att(pp, 'fo:padding-' + sd) || att(pp, 'fo:padding')); const mt = inch(att(pp, 'fo:margin-top')) + padI('top'), mb = inch(att(pp, 'fo:margin-bottom')) + padI('bottom');
          D.page = { size, orient: land ? 'landscape' : 'portrait', m: { t: hp ? Math.max(mt + 0.4, mt + inch(att(hp, 'fo:min-height')) + inch(att(hp, 'fo:margin-bottom')) + 0.2) : mt, b: fp ? Math.max(mb + 0.4, mb + inch(att(fp, 'fo:min-height')) + inch(att(fp, 'fo:margin-top')) + 0.2) : mb, l: inch(att(pp, 'fo:margin-left')) + padI('left'), r: inch(att(pp, 'fo:margin-right')) + padI('right') }, hdrDist: hp ? mt : 0.5, ftrDist: fp ? mb : 0.5 };
          ['t', 'b', 'l', 'r'].forEach(k => { D.page.m[k] = Math.round(D.page.m[k] * 100) / 100; });
          if (att(pp, 'fo:background-color') && att(pp, 'fo:background-color') !== 'transparent') D.design.pageColor = att(pp, 'fo:background-color');
        }
        const h = first(mp, 'style:header'), f = first(mp, 'style:footer');
        if (h && att(h, 'style:display') !== 'false') hdr = finish(await conv(h));
        if (f && att(f, 'style:display') !== 'false') ftr = finish(await conv(f));
      }
    }
    const html = finish(await conv(bodyEl));
    if (bts) { if (bts.page) D.page = Object.assign(D.defaultPage(), bts.page); if (bts.design) Object.assign(D.design, bts.design); }
    D.design.track = !!(first(x, 'text:tracked-changes') && att(first(x, 'text:tracked-changes'), 'text:track-changes') === 'true');
    Object.keys(comments).forEach(k => { if (html.includes(`data-cid="${k}"`)) { const c = comments[k]; delete c.name; D.comments[k] = c; } });
    if (mx) {
      const t = (n) => ((mx.getElementsByTagName(n)[0] || {}).textContent || '').trim();
      BTS.setProps({ title: t('dc:title'), subject: t('dc:subject'), comments: t('dc:description'), author: t('meta:initial-creator') || t('dc:creator'), keywords: [...mx.getElementsByTagName('meta:keyword')].map(k => k.textContent.trim()).filter(Boolean).join(', '), company: ud.Company || '', category: ud.Category || '' });
    }
    D.setContent(D.sanitize(html), { hdr: D.sanitize(hdr), ftr: D.sanitize(ftr) });
  };

  /* ================= dispatch ================= */
  D.loadFile = async ({ name, data }) => {
    const ext = BTS.extOf(name);
    const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
    const text = () => new TextDecoder('utf-8').decode(bytes).replace(/^﻿/, '');
    D.endPreview(); D.selectImage && D.selectImage(null);
    if (ext === 'docx' || ext === 'docm' || ext === 'dotx') {
      if (bytes[0] !== 0x50 || bytes[1] !== 0x4B) throw new Error('This isn’t a valid .docx file.');
      await D.importDocx(bytes);
    }
    else if (ext === 'doc') throw new Error('Old Word 97–2003 (.doc) files aren’t supported. Open it in Word or LibreOffice and save as .docx.');
    else if (ext === 'odt') await D.importOdt(bytes);
    else if (ext === 'html' || ext === 'htm') D.loadNative(text());
    else if (ext === 'md' || ext === 'markdown') { D.page = D.defaultPage(); D.design = D.defaultDesign(); D.comments = {}; D.setContent(D.sanitize(D.mdToHTML(text()))); }
    else if (ext === 'rtf') { const t = text().replace(/\\par[d]?/g, '\n').replace(/\{\\\*[^{}]*\}|\\[a-z]+-?\d* ?|[{}]/g, ''); D.page = D.defaultPage(); D.design = D.defaultDesign(); D.comments = {}; D.setContent(D.textToHTML(t.trim())); }
    else { D.page = D.defaultPage(); D.design = D.defaultDesign(); D.comments = {}; D.setContent(D.textToHTML(text())); }
    D.normalizeMergeFields && D.normalizeMergeFields(); if (D.mm && D.mm.preview) D.mmPreview(false);
    D.renderComments(); D.renderNav(); D.canvas.scrollTop = 0; $('#recover-bar').classList.add('hidden');
    const firstBlock = D.doc.querySelector('p,h1,h2,h3,li,td'); if (firstBlock) D.setCaret(firstBlock, 0);
    setTimeout(() => D.layout(), 50);
  };
})();
