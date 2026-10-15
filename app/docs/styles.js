/* BTS Docs 1.2 — fonts, named styles, style sets.
 * Styles are data (D.styleDef(id)) rendered into a generated stylesheet, so "Modify style" updates every
 * paragraph that uses the style. Effective definition = BASE ⊕ style set overrides ⊕ document modifications.
 * Paragraph styles map to tag + class (h1, p.title, p.us-<id> …); character styles to span.cs-<id>. */
(function () {
  const D = window.Docs; const { $, $$ } = D; const esc = BTS.esc;

  /* ================= fonts ================= */
  const FALLBACK = { 'Calibri Light': 'Carlito', Calibri: 'Carlito', Cambria: 'Caladea', Arial: 'Arimo', 'Times New Roman': 'Tinos', 'Courier New': 'Cousine', Georgia: 'Gelasio', Consolas: 'Cousine', 'Segoe UI': 'Noto Sans', Verdana: 'PT Sans', Tahoma: 'Noto Sans', 'Trebuchet MS': 'Noto Sans', 'Century Gothic': 'Poppins', Aptos: 'Inter', Garamond: 'EB Garamond', 'Book Antiqua': 'Gelasio', 'Palatino Linotype': 'Gelasio', 'Comic Sans MS': 'Kalam', 'Nirmala UI': 'Noto Sans Devanagari', Mangal: 'Noto Sans Devanagari', Kokila: 'Tiro Devanagari Hindi' };
  const GENERIC = { sans: 'sans-serif', serif: 'serif', display: 'serif', mono: 'monospace', indian: 'sans-serif', script: 'cursive' };
  D.fontInfo = (name) => (window.BTS_FONTS || []).find(f => f.name.toLowerCase() === String(name || '').toLowerCase()) || null;
  D.fontCat = (name) => { const f = D.fontInfo(name); if (f) return f.cat; if (/mono|consol|courier|code/i.test(name)) return 'mono'; if (/serif|roman|garamond|georgia|cambria|baskerville|playfair|lora|merriweather|tinos|caladea|gelasio|slab|crimson|antiqua|palatino/i.test(name) && !/sans/i.test(name)) return 'serif'; return 'sans'; };
  D.isIndianFont = (name) => D.fontCat(name) === 'indian';
  /** CSS font stack: the font, its metric-compatible bundled stand-in, a generic family, then a bundled Devanagari font. */
  D.fontStack = (name) => {
    if (!name) name = 'Calibri';
    const cat = D.fontCat(name); const q = (n) => '"' + String(n).replace(/"/g, '') + '"';
    const parts = [q(name)]; if (FALLBACK[name]) parts.push(q(FALLBACK[name]));
    if (name === 'Calibri Light') parts.push('"Calibri"');
    parts.push(cat === 'serif' || cat === 'display' ? '"Noto Serif Devanagari"' : '"Noto Sans Devanagari"');
    parts.push(GENERIC[cat] || 'sans-serif');
    return parts.join(', ');
  };
  D.FONTS = (window.BTS_FONTS || []).map(f => f.name).concat(['Calibri Light']).filter((v, i, a) => a.indexOf(v) === i).sort((a, b) => a.localeCompare(b));
  /** Devanagari-capable font to pair with a Latin font in Word's complex-script slot. */
  D.csFontFor = (name) => (D.isIndianFont(name) ? name : 'Nirmala UI');
  D.hasDevanagari = (s) => /[ऀ-ॿ꣠-ꣿ]/.test(s || '');

  /* ================= base style definitions ================= */
  const P = (id, name, tag, cls, def, extra) => ({ id, name, tag, cls, type: 'p', builtin: true, def, extra: extra || '' });
  const C = (id, name, def) => ({ id, name, type: 'c', builtin: true, cls: 'cs-' + id, def });
  D.BASE_STYLES = [
    P('normal', 'Normal', 'p', '', { font: '@body', size: 11, color: '#111111', bold: false, italic: false, before: null, after: null, line: null, align: 'left', indent: 0, firstLine: 0 }),
    P('nospace', 'No Spacing', 'p', 'nospace', { before: 0, after: 0, line: 1 }),
    P('title', 'Title', 'p', 'title', { font: '@head', size: 28, color: '@dark', before: 0, after: 4, line: 1.1, letter: -0.5 }),
    P('subtitle', 'Subtitle', 'p', 'subtitle', { font: '@body', size: 13, color: '#5A6275', before: 0, after: 12, letter: 0.4 }),
    P('h1', 'Heading 1', 'h1', '', { font: '@head', size: 16, color: '@accent', before: 14, after: 4, line: 1.15 }),
    P('h2', 'Heading 2', 'h2', '', { font: '@head', size: 13, color: '@accent', before: 10, after: 3, line: 1.15 }),
    P('h3', 'Heading 3', 'h3', '', { font: '@head', size: 12, color: '@dark', before: 8, after: 3, line: 1.15 }),
    P('h4', 'Heading 4', 'h4', '', { font: '@head', size: 11, color: '@accent', italic: true, before: 8, after: 2, line: 1.15 }),
    P('h5', 'Heading 5', 'h5', '', { font: '@head', size: 11, color: '@dark', before: 6, after: 2, line: 1.15 }),
    P('h6', 'Heading 6', 'h6', '', { font: '@head', size: 11, color: '@dark', italic: true, before: 6, after: 2, line: 1.15 }),
    P('quote', 'Quote', 'blockquote', '', { font: '@body', size: 11, color: '#404654', italic: true, before: 10, after: 10, indent: 0.76 }, 'border-left:3px solid var(--th-accent);padding:2pt 0 2pt 12pt;margin-right:.4in;'),
    P('iquote', 'Intense Quote', 'blockquote', 'intense', { font: '@body', size: 11, color: '@accent', italic: true, before: 14, after: 14, indent: 0.86, align: 'center' }, 'border-left:0;border-top:1px solid var(--th-accent);border-bottom:1px solid var(--th-accent);padding:8pt 0;margin-right:.34in;'),
    P('caption', 'Caption', 'p', 'caption', { font: '@body', size: 9, color: '@dark', italic: true, before: 2, after: 10 }),
    P('code', 'Code', 'pre', '', { font: '@mono', size: 10, color: '#1D2433', before: 0, after: 8, line: 1.35 }, 'background:#f3f4f7;border:1px solid #e2e5eb;border-radius:3px;padding:6pt 8pt;white-space:pre-wrap;'),
    P('listpara', 'List Paragraph', 'p', 'listpara', { indent: 1.27, before: 0, after: 2 }),
    C('emphasis', 'Emphasis', { italic: true }),
    C('strong', 'Strong', { bold: true }),
    C('subtle-em', 'Subtle Emphasis', { italic: true, color: '#404040' }),
    C('intense-em', 'Intense Emphasis', { italic: true, color: '@accent' }),
    C('subtle-ref', 'Subtle Reference', { smallcaps: true, color: '#5A5A5A' }),
    C('intense-ref', 'Intense Reference', { bold: true, smallcaps: true, color: '@accent', letter: 0.25 }),
    C('book-title', 'Book Title', { bold: true, italic: true, letter: 0.25 }),
  ];

  /* ================= style sets (Design ▸ Style sets) ================= */
  D.STYLE_SETS = [
    { id: 'classic', name: 'Classic', head: 'Calibri Light', body: 'Calibri', desc: 'Calibri Light headings with Calibri text — matches Microsoft Word', s: {} },
    { id: 'corporate', name: 'Corporate', head: 'Inter', body: 'Source Serif 4', accent: '#1F3A5F', dark: '#14213D', desc: 'Inter headings, Source Serif 4 body', s: { normal: { size: 11, color: '#1E2430' }, title: { size: 30, bold: true, weight: 700, letter: -0.8 }, subtitle: { color: '#55627A' }, h1: { size: 17, bold: true, weight: 700, before: 18, after: 6 }, h2: { size: 13.5, bold: true, weight: 600 }, h3: { size: 12, bold: true, weight: 600, color: '#3D4B66' } } },
    { id: 'academic', name: 'Academic', head: 'EB Garamond', body: 'EB Garamond', accent: '#222222', dark: '#111111', desc: 'EB Garamond throughout, centred title, 1.5 line spacing', s: { normal: { size: 12, color: '#111111', line: 1.5, after: 6 }, title: { size: 24, align: 'center', color: '#111111', after: 6 }, subtitle: { align: 'center', size: 13, italic: true, color: '#333333' }, h1: { size: 16, bold: true, align: 'center', color: '#111111', before: 18 }, h2: { size: 13, bold: true, color: '#111111' }, h3: { size: 12, bold: true, italic: true, color: '#111111' }, quote: { color: '#222222' }, caption: { color: '#333333' } } },
    { id: 'modern', name: 'Modern', head: 'Montserrat', body: 'Open Sans', accent: '#0F766E', dark: '#134E4A', desc: 'Montserrat headings, Open Sans body', s: { normal: { size: 10.5, color: '#1F2937' }, title: { size: 30, bold: true, weight: 800, letter: -0.5, color: '#0F172A' }, subtitle: { caps: true, letter: 2, size: 10.5, color: '#0F766E' }, h1: { size: 16, bold: true, weight: 700, color: '#0F766E', before: 18 }, h2: { size: 12.5, bold: true, weight: 600, color: '#134E4A' }, h3: { size: 11, bold: true, weight: 600, caps: true, letter: 1, color: '#334155' } } },
    { id: 'elegant', name: 'Elegant', head: 'Playfair Display', body: 'Lato', accent: '#7A4E2D', dark: '#3E2716', desc: 'Playfair Display headings, Lato body', s: { normal: { size: 11, color: '#2A2420', line: 1.3 }, title: { size: 34, color: '#3E2716', align: 'center', letter: 0 }, subtitle: { align: 'center', italic: true, color: '#7A4E2D' }, h1: { size: 19, color: '#3E2716', before: 20, after: 6 }, h2: { size: 14, italic: true, color: '#7A4E2D' }, h3: { size: 12, bold: true, color: '#3E2716' } } },
    { id: 'technical', name: 'Technical', head: 'IBM Plex Sans', body: 'IBM Plex Serif', mono: 'IBM Plex Mono', accent: '#0F62FE', dark: '#161616', desc: 'IBM Plex Sans, Serif and Mono', s: { normal: { size: 10.5, color: '#161616' }, title: { size: 28, bold: true, weight: 700, color: '#161616' }, subtitle: { color: '#525252' }, h1: { size: 16, bold: true, weight: 700, color: '#0F62FE' }, h2: { size: 13, bold: true, weight: 500, color: '#161616' }, h3: { size: 11, bold: true, color: '#393939' }, code: { size: 9.5 } } },
    { id: 'report', name: 'Report', head: 'Source Sans 3', body: 'Source Sans 3', accent: '#2B579A', dark: '#1B365D', desc: 'Source Sans 3, ruled headings', s: { normal: { size: 11, color: '#1B1F27' }, title: { size: 30, bold: true, weight: 700, color: '#1B365D' }, h1: { size: 16, bold: true, weight: 700, color: '#1B365D', before: 18, after: 6 }, h2: { size: 13, bold: true, weight: 600, color: '#2B579A' }, h3: { size: 11.5, bold: true, color: '#1B365D' } }, extra: { h1: 'border-bottom:1.5px solid var(--th-accent);padding-bottom:2pt;' } },
    { id: 'minimal', name: 'Minimal', head: 'DM Sans', body: 'DM Sans', accent: '#111111', dark: '#111111', desc: 'DM Sans, black on white, airy spacing', s: { normal: { size: 10.5, color: '#222222', line: 1.4, after: 9 }, title: { size: 32, bold: true, weight: 700, letter: -1, color: '#111111' }, subtitle: { color: '#6B6B6B' }, h1: { size: 15, bold: true, weight: 700, color: '#111111', before: 22, after: 6 }, h2: { size: 12, bold: true, weight: 500, color: '#111111' }, h3: { size: 10.5, bold: true, color: '#444444', caps: true, letter: 1 } } },
    { id: 'editorial', name: 'Editorial', head: 'Libre Baskerville', body: 'Work Sans', accent: '#B3261E', dark: '#2B1B17', desc: 'Libre Baskerville headings, Work Sans body', s: { normal: { size: 10.5, color: '#231F20', line: 1.35 }, title: { size: 30, color: '#2B1B17' }, subtitle: { italic: true, color: '#B3261E' }, h1: { size: 17, color: '#2B1B17', before: 18 }, h2: { size: 13, italic: true, color: '#B3261E' }, h3: { size: 10.5, bold: true, caps: true, letter: 1.2, color: '#2B1B17' }, quote: { size: 13, color: '#2B1B17' } } },
    { id: 'friendly', name: 'Friendly', head: 'Nunito', body: 'Nunito', accent: '#E8590C', dark: '#7C2D12', desc: 'Nunito, warm colours', s: { normal: { size: 11, color: '#1F2937', line: 1.3 }, title: { size: 30, bold: true, weight: 700, color: '#E8590C' }, subtitle: { color: '#7C2D12' }, h1: { size: 17, bold: true, weight: 700, color: '#E8590C' }, h2: { size: 13.5, bold: true, weight: 700, color: '#7C2D12' }, h3: { size: 12, bold: true, color: '#9A3412' } } },
    { id: 'letter', name: 'Formal letter', head: 'Times New Roman', body: 'Times New Roman', accent: '#000000', dark: '#000000', desc: 'Times New Roman (Tinos) 12 pt, single spacing', s: { normal: { size: 12, color: '#000000', line: 1, after: 10 }, title: { size: 20, bold: true, color: '#000000', align: 'center' }, subtitle: { color: '#000000', align: 'center' }, h1: { size: 14, bold: true, color: '#000000' }, h2: { size: 12, bold: true, color: '#000000' }, h3: { size: 12, bold: true, italic: true, color: '#000000' } } },
    { id: 'hindi', name: 'Hindi / Marathi', head: 'Noto Serif Devanagari', body: 'Mukta', accent: '#B45309', dark: '#78350F', desc: 'Noto Serif Devanagari headings, Mukta body — हिन्दी, मराठी', s: { normal: { size: 12, color: '#1C1917', line: 1.35 }, title: { size: 28, bold: true, weight: 700, color: '#78350F' }, h1: { size: 18, bold: true, weight: 700, color: '#B45309' }, h2: { size: 14, bold: true, weight: 600, color: '#78350F' }, h3: { size: 12.5, bold: true, color: '#92400E' } } },
    { id: 'lively', name: 'Lively', head: 'Poppins', body: 'Lora', accent: '#7C3AED', dark: '#3B0764', desc: 'Poppins headings, Lora body', s: { normal: { size: 11, color: '#1E1B2E', line: 1.3 }, title: { size: 30, bold: true, weight: 600, color: '#3B0764' }, subtitle: { color: '#7C3AED' }, h1: { size: 16, bold: true, weight: 600, color: '#7C3AED' }, h2: { size: 13, bold: true, weight: 500, color: '#3B0764' }, h3: { size: 11.5, bold: true, color: '#5B21B6' } } },
    { id: 'bold', name: 'Bold', head: 'Oswald', body: 'Roboto', accent: '#C2410C', dark: '#1C1917', desc: 'Oswald headings, Roboto body', s: { normal: { size: 10.5, color: '#1C1917' }, title: { size: 34, caps: true, color: '#1C1917', letter: 1 }, subtitle: { caps: true, letter: 2, color: '#C2410C', size: 11 }, h1: { size: 19, caps: true, color: '#C2410C', letter: 0.5 }, h2: { size: 14, color: '#1C1917' }, h3: { size: 12, color: '#57534E' } } },
  ];
  D.styleSet = () => D.STYLE_SETS.find(s => s.id === D.design.styleSet) || D.STYLE_SETS[0];

  /* ================= registry ================= */
  D.STYLES.length = 0;
  D.allStyles = () => D.BASE_STYLES.concat((D.design.customStyles || []).map(u => ({ id: u.id, name: u.name, type: u.type || 'p', tag: u.type === 'c' ? 'span' : 'p', cls: (u.type === 'c' ? 'cs-' : 'us-') + u.id, def: u.def || {}, custom: true })));
  D.findStyle = (id) => D.allStyles().find(s => s.id === id) || null;
  function syncStyleList() { const list = D.allStyles().filter(s => s.type === 'p'); D.STYLES.length = 0; list.forEach(s => D.STYLES.push(s)); }
  /** Effective definition of a style (merged, unresolved tokens like @head stay). */
  D.styleDef = (id) => {
    const st = D.findStyle(id); if (!st) return {};
    const set = D.styleSet(); const mods = (D.design.styleMods || {})[id] || {};
    return Object.assign({}, st.def, st.custom ? {} : ((set.s || {})[id] || {}), mods);
  };
  D.headFont = () => D.design.headFont || D.styleSet().head || D.theme().head;
  D.bodyFont = () => D.design.bodyFont || D.styleSet().body || D.theme().body;
  D.monoFont = () => D.styleSet().mono || 'Consolas';
  D.accent = () => D.design.accent || D.theme().accent;
  D.darkColor = () => D.design.dark || D.theme().dark;
  D.resolveFont = (f) => f === '@head' ? D.headFont() : f === '@body' ? D.bodyFont() : f === '@mono' ? D.monoFont() : f;
  D.resolveColor = (c) => c === '@accent' ? D.accent() : c === '@dark' ? D.darkColor() : c;
  /** Fully resolved paragraph style (inherits Normal), as plain values for export. */
  D.resolvedStyle = (id) => {
    const st = D.findStyle(id); const norm = D.styleDef('normal'); const ps = D.design.pspacing || { after: 8, line: 1.15, before: 0 };
    const d = st && st.type === 'c' ? D.styleDef(id) : Object.assign({}, norm, id === 'normal' ? {} : D.styleDef(id));
    const out = Object.assign({}, d, { font: D.resolveFont(d.font || '@body'), color: D.resolveColor(d.color || '#111111') });
    if (out.before == null) out.before = ps.before || 0; if (out.after == null) out.after = ps.after; if (out.line == null) out.line = ps.line;
    return out;
  };

  /* ================= stylesheet generation ================= */
  const cssColor = (c) => c === '@accent' ? 'var(--th-accent)' : c === '@dark' ? 'var(--th-dark)' : c;
  const cssFont = (f) => f === '@head' ? 'var(--th-head)' : f === '@body' ? 'var(--th-body)' : f === '@mono' ? 'var(--th-mono)' : D.fontStack(f);
  function declare(d, isChar) {
    const o = [];
    if (d.font) o.push('font-family:' + cssFont(d.font));
    if (d.size) o.push('font-size:' + d.size + 'pt');
    if (d.color) o.push('color:' + cssColor(d.color));
    if (d.bold !== undefined || d.weight) o.push('font-weight:' + (d.bold ? (d.weight || 700) : 400));
    if (d.italic !== undefined) o.push('font-style:' + (d.italic ? 'italic' : 'normal'));
    if (d.underline !== undefined) o.push('text-decoration:' + (d.underline ? 'underline' : 'none'));
    if (d.caps !== undefined) o.push('text-transform:' + (d.caps ? 'uppercase' : 'none'));
    if (d.smallcaps !== undefined) o.push('font-variant-caps:' + (d.smallcaps ? 'small-caps' : 'normal'));
    if (d.letter !== undefined) o.push('letter-spacing:' + d.letter + 'pt');
    if (!isChar) {
      if (d.before !== undefined) o.push('margin-top:' + (d.before == null ? 'var(--p-before)' : d.before + 'pt'));
      if (d.after !== undefined) o.push('margin-bottom:' + (d.after == null ? 'var(--p-after)' : d.after + 'pt'));
      if (d.line !== undefined) o.push('line-height:' + (d.line == null ? 'var(--p-line)' : d.line));
      if (d.align) o.push('text-align:' + d.align);
      if (d.indent !== undefined) o.push('margin-left:' + (d.indent || 0) + 'cm');
      if (d.indentRight !== undefined) o.push('margin-right:' + (d.indentRight || 0) + 'cm');
      if (d.firstLine !== undefined) o.push('text-indent:' + (d.firstLine || 0) + 'cm');
    }
    return o.join(';');
  }
  D.styleCSSText = (id) => { const st = D.findStyle(id); return st ? declare(D.styleDef(id), st.type === 'c') : ''; };
  function selectorOf(st) {
    if (st.id === 'normal') return '.doc-content p, .doc-content li > p';
    if (st.type === 'c') return `.doc-content .${st.cls}`;
    return `.doc-content ${st.tag}${st.cls ? '.' + st.cls : (st.tag === 'blockquote' ? ':not(.intense)' : '')}`;
  }
  D.applyStyleSheet = () => {
    let el = document.getElementById('doc-styles'); if (!el) { el = document.createElement('style'); el.id = 'doc-styles'; document.head.appendChild(el); }
    syncStyleList();
    const set = D.styleSet(); const rules = [];
    const nd = D.styleDef('normal');
    rules.push(`.doc-content { font-family:${cssFont(nd.font || '@body')}; font-size:${nd.size || 11}pt; color:${cssColor(nd.color || '#111111')}; ${nd.line != null ? 'line-height:' + nd.line : ''} }`);
    D.allStyles().forEach(st => {
      const d = D.styleDef(st.id); if (st.id === 'normal') { const x = Object.assign({}, d); delete x.font; delete x.size; delete x.color; rules.push(`${selectorOf(st)} { ${declare(x)} }`); return; }
      const extra = (st.extra || '') + ((set.extra || {})[st.id] || '');
      const imp = st.id === 'nospace' ? ' !important' : '';
      const body = declare(d, st.type === 'c').split(';').filter(Boolean).map(x => x + imp).join(';');
      rules.push(`${selectorOf(st)} { ${body}; ${extra} }`);
      if (st.type === 'p') rules.push(`.doc-content .pv-on.pv-${st.id} { ${declare(Object.assign({ font: '@body', size: 11, color: '#111111', bold: false, italic: false, caps: false, letter: 0 }, d))}; ${extra} }`);
    });
    rules.push('.doc-content blockquote p { font: inherit; color: inherit; margin: 0 0 4pt; text-align: inherit; }');
    el.textContent = rules.join('\n');
  };

  /* ================= paragraph & character style helpers ================= */
  D.styleOf = (b) => {
    if (!b) return 'normal';
    const t = b.tagName;
    if (/^H[1-6]$/.test(t)) return 'h' + t[1];
    const bq = t === 'BLOCKQUOTE' ? b : b.closest('blockquote'); if (bq && D.rootOf(bq)) return bq.classList.contains('intense') ? 'iquote' : 'quote';
    if (t === 'PRE') return 'code';
    for (const c of ['title', 'subtitle', 'nospace', 'caption', 'listpara']) if (b.classList.contains(c)) return c;
    const us = [...b.classList].find(c => c.startsWith('us-')); if (us && D.findStyle(us.slice(3))) return us.slice(3);
    return 'normal';
  };
  /** Apply (or toggle off) a character style on the selection. */
  D.applyCharStyle = (id) => {
    const st = D.findStyle(id); if (!st) return;
    D.ensureSel();
    D.change(() => {
      const r0 = D.range(); if (!r0) return; if (r0.collapsed) D.selectWordIfCollapsed();
      const r = D.range(); if (!r || r.collapsed) return;
      const root = D.rootOf(r.commonAncestorContainer);
      const inside = D.closest(r.startContainer, '.' + st.cls);
      // remove any character styles inside the selection first
      $$('[class*="cs-"]', root).forEach(el => { if (r.intersectsNode(el) && [...el.classList].some(c => c.startsWith('cs-'))) { if (el.tagName === 'SPAN' && el.classList.length === 1 && !el.getAttribute('style')) D.unwrap(el); else [...el.classList].filter(c => c.startsWith('cs-')).forEach(c => el.classList.remove(c)); } });
      if (inside) return; // toggled off
      const r2 = D.range(); if (r2 && !r2.collapsed) D.wrapText(r2, () => { const s = document.createElement('span'); s.className = st.cls; return s; });
    });
  };
  D.charStyleAt = (node) => { let el = node && (node.nodeType === 1 ? node : node.parentElement); while (el && !D.isRoot(el)) { const c = el.classList && [...el.classList].find(x => x.startsWith('cs-')); if (c) return c.slice(3); el = el.parentElement; } return null; };

  /* ================= galleries ================= */
  const GAL = ['normal', 'nospace', 'title', 'subtitle', 'h1', 'h2', 'h3', 'h4', 'quote', 'iquote', 'emphasis', 'strong', 'subtle-em', 'intense-em', 'subtle-ref', 'intense-ref', 'book-title', 'caption', 'code', 'listpara'];
  D.renderStylesGallery = () => {
    const g = $('#styles-gallery'); if (!g) return;
    const ids = GAL.concat((D.design.customStyles || []).map(u => u.id));
    g.innerHTML = ids.map(id => D.findStyle(id)).filter(Boolean).map(s => {
      const d = D.styleDef(s.id); const sz = Math.max(10, Math.min(17, (d.size || 11) * (s.id === 'title' ? 0.62 : 1.05)));
      const css = declare(Object.assign({}, s.type === 'c' ? { font: '@body' } : {}, d), true).replace(/font-size:[^;]+;?/, '');
      return `<button class="gi keepfocus" data-cmd="${s.type === 'c' ? 'char-style' : 'style'}" data-style="${s.id}" title="${esc(s.name)}${s.type === 'c' ? ' (character style)' : ''} — right-click to modify"><span class="pv doc-content" style="--th-accent:${D.accent()};--th-dark:${D.darkColor()};--th-head:${esc(D.fontStack(D.headFont()))};--th-body:${esc(D.fontStack(D.bodyFont()))};--th-mono:${esc(D.fontStack(D.monoFont()))}"><span style="${esc(css)};font-size:${sz}px;display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin:0;padding:0;border:0;background:none;line-height:1.2">${s.id === 'code' ? 'code()' : 'AaBbCc'}</span></span><small>${esc(s.name)}</small></button>`;
    }).join('');
  };
  function setThumb(set) {
    const h = D.fontStack(set.head), b = D.fontStack(set.body); const acc = set.accent || D.theme().accent, dk = set.dark || D.theme().dark;
    const t = Object.assign({}, (set.s || {}).title), h1 = Object.assign({}, (set.s || {}).h1);
    return `<span class="ss-thumb"><span style="font-family:${esc(h)};color:${dk};font-size:11px;font-weight:${t.bold ? (t.weight || 700) : 400};${t.caps ? 'text-transform:uppercase;' : ''}${t.align === 'center' ? 'text-align:center;' : ''}">Title</span><span style="font-family:${esc(h)};color:${acc};font-size:7.5px;font-weight:${h1.bold ? (h1.weight || 700) : 400};${h1.caps ? 'text-transform:uppercase;' : ''}">Heading 1</span><span style="font-family:${esc(b)};font-size:5px;line-height:1.3;color:#333">${set.id === 'hindi' ? 'यह अनुच्छेद का नमूना पाठ है। हिन्दी और मराठी में लिखें।' : 'The quick brown fox jumps over the lazy dog. Body text sample for this set.'}</span></span>`;
  }
  D.renderStyleSets = () => {
    const g = $('#stylesets-gallery'); if (!g) return;
    g.innerHTML = D.STYLE_SETS.map(s => `<button class="gi keepfocus" data-cmd="style-set" data-set="${s.id}" title="${esc(s.name)} — ${esc(s.desc)}">${setThumb(s)}<small>${esc(s.name)}</small></button>`).join('');
    $$('.gi', g).forEach(x => x.classList.toggle('on', x.dataset.set === (D.design.styleSet || 'classic')));
  };
  D.applyStyleSet = (id) => {
    const set = D.STYLE_SETS.find(s => s.id === id); if (!set) return;
    D.setDesign({ styleSet: id, styleMods: {}, headFont: null, bodyFont: null, accent: set.accent || null, dark: set.dark || null });
    BTS.toast(`Style set “${set.name}” applied`);
  };

  /* ================= Styles pane ================= */
  D.toggleStylesPane = (on) => {
    const p = $('#styles-pane'); const show = on ?? p.classList.contains('hidden');
    p.classList.toggle('hidden', !show); $$('[data-cmd="styles-pane"]').forEach(b => b.classList.toggle('on', show));
    const chk = $('#chk-styles'); if (chk) chk.checked = show;
    if (show) D.renderStylesPane(); D.scheduleLayout(30);
  };
  D.renderStylesPane = () => {
    const p = $('#styles-pane'); if (!p || p.classList.contains('hidden')) return;
    const r = D.range() || D.lastRange(); const cur = r ? D.styleOf(D.blockOf(r.startContainer)) : 'normal'; const curC = r ? D.charStyleAt(r.startContainer) : null;
    const counts = {}; $$('p,h1,h2,h3,h4,h5,h6,pre,blockquote', D.doc).forEach(b => { if (b.closest('.toc,.biblio')) return; if ((b.tagName === 'BLOCKQUOTE' && !D.hasBlockChild(b)) || (b.tagName !== 'BLOCKQUOTE' && D.blockOf(b) === b)) { const s = D.styleOf(b); counts[s] = (counts[s] || 0) + 1; } });
    const row = (s) => { const d = D.styleDef(s.id); const css = declare(Object.assign({}, d), true).replace(/font-size:[^;]+;?/, ''); return `<div class="sp-row${(s.type === 'p' ? s.id === cur : s.id === curC) ? ' on' : ''}" data-sid="${s.id}" data-type="${s.type}" title="${esc(describe(s.id))}"><span class="sp-name doc-content" style="${esc(css)};font-size:${Math.max(11, Math.min(16, (d.size || 11) * 1.1))}px">${esc(s.name)}</span><span class="sp-kind">${s.type === 'c' ? 'a' : '¶'}</span>${counts[s.id] ? `<span class="sp-count" title="Paragraphs using this style">${counts[s.id]}</span>` : ''}<button class="rb sp-menu" data-cmd="style-row-menu" data-sid="${s.id}" title="Options for ${esc(s.name)}">${icon('ellipsis', 14)}</button></div>`; };
    const all = D.allStyles();
    $('#styles-list').innerHTML = `<div class="sp-head">Paragraph styles</div>${all.filter(s => s.type === 'p').map(row).join('')}<div class="sp-head">Character styles</div>${all.filter(s => s.type === 'c').map(row).join('')}`;
    $('#styles-list').querySelectorAll('.sp-row').forEach(el => { el.style.setProperty('--th-accent', D.accent()); el.style.setProperty('--th-dark', D.darkColor()); el.style.setProperty('--th-head', D.fontStack(D.headFont())); el.style.setProperty('--th-body', D.fontStack(D.bodyFont())); el.style.setProperty('--th-mono', D.fontStack(D.monoFont())); });
    $('#styles-set-name').textContent = D.styleSet().name;
  };
  function describe(id) {
    const d = D.resolvedStyle(id); const st = D.findStyle(id);
    const f = []; if (st.type === 'p' || d.font) f.push(d.font + ', ' + (d.size || 11) + ' pt'); if (d.bold) f.push('Bold'); if (d.italic) f.push('Italic'); if (d.caps) f.push('All caps'); if (d.smallcaps) f.push('Small caps');
    if (st.type === 'p') { f.push('Space before ' + (d.before || 0) + ' pt, after ' + (d.after || 0) + ' pt'); f.push('Line spacing ' + (d.line || 1)); if (d.align && d.align !== 'left') f.push('Aligned ' + d.align); if (d.indent) f.push('Indent ' + d.indent + ' cm'); }
    return st.name + ': ' + f.join(' · ');
  }
  function wireStylesPane() {
    const list = $('#styles-list'); if (!list) return;
    list.addEventListener('mousedown', e => { if (e.target.closest('.sp-row')) e.preventDefault(); });
    list.addEventListener('click', e => {
      if (e.target.closest('.sp-menu')) return;
      const row = e.target.closest('.sp-row'); if (!row) return;
      D.restoreFocus(); if (row.dataset.type === 'c') D.applyCharStyle(row.dataset.sid); else D.setBlockStyle(row.dataset.sid);
      D.renderStylesPane();
    });
    list.addEventListener('contextmenu', e => { const row = e.target.closest('.sp-row'); if (!row) return; e.preventDefault(); D.styleMenu({ x: e.clientX, y: e.clientY }, row.dataset.sid); });
    list.addEventListener('mouseover', e => { const row = e.target.closest('.sp-row[data-type="p"]'); if (row && row.dataset.sid !== hov) { hov = row.dataset.sid; D.previewStyle(hov); } });
    list.addEventListener('mouseleave', () => { hov = null; D.endPreview(); });
    $('#styles-gallery').addEventListener('contextmenu', e => { const b = e.target.closest('.gi'); if (!b) return; e.preventDefault(); D.styleMenu({ x: e.clientX, y: e.clientY }, b.dataset.style); });
  }
  let hov = null;
  D.styleMenu = (anchor, id) => {
    const st = D.findStyle(id); if (!st) return;
    BTS.menu(anchor, [{ head: st.name },
      { label: st.type === 'c' ? 'Apply' : 'Apply to paragraph', icon: 'check', action: () => { D.restoreFocus(); st.type === 'c' ? D.applyCharStyle(id) : D.setBlockStyle(id); } },
      { label: 'Modify…', icon: 'pencil', action: () => D.modifyStyleDialog(id) },
      { label: `Update ${st.name} to match selection`, icon: 'refresh-cw', action: () => D.updateStyleFromSelection(id) },
      { label: 'Select all instances', icon: 'text-select', disabled: st.type === 'c', action: () => selectInstances(id) },
      { sep: true },
      { label: 'Reset to style set default', icon: 'rotate-ccw', disabled: st.custom || !(D.design.styleMods || {})[id], action: () => { D.setDesign({ styleMods: Object.fromEntries(Object.entries(D.design.styleMods || {}).filter(([k]) => k !== id)) }); } },
      { label: 'Delete style', icon: 'trash-2', disabled: !st.custom, action: () => deleteStyle(id) },
    ]);
  };
  function selectInstances(id) {
    const bl = $$('p,h1,h2,h3,h4,h5,h6,pre,blockquote', D.doc).filter(b => (b.tagName === 'BLOCKQUOTE' || D.blockOf(b) === b) && D.styleOf(b) === id);
    if (!bl.length) { BTS.toast('No paragraphs use this style.'); return; }
    const r = document.createRange(); r.setStart(bl[0], 0); r.setEnd(bl[bl.length - 1], bl[bl.length - 1].childNodes.length); D.doc.focus(); const s = getSelection(); s.removeAllRanges(); s.addRange(r);
    BTS.toast(`${bl.length} paragraph${bl.length === 1 ? '' : 's'} selected`);
  }
  function deleteStyle(id) {
    const st = D.findStyle(id);
    D.change(() => {
      $$('.' + st.cls, D.doc).forEach(el => { el.classList.remove(st.cls); if (!el.className) el.removeAttribute('class'); if (st.type === 'c' && el.tagName === 'SPAN' && !el.attributes.length) D.unwrap(el); });
      D.design.customStyles = (D.design.customStyles || []).filter(u => u.id !== id); D.applySettings();
    });
  }

  /* ================= Modify style / New style dialogs ================= */
  const pct = (v) => v == null ? '' : String(v);
  function styleForm(id, d, isChar, isNew) {
    const fontLabel = (f) => f === '@head' ? `Headings font (${D.headFont()})` : f === '@body' ? `Body font (${D.bodyFont()})` : f === '@mono' ? `Code font (${D.monoFont()})` : f || `Inherit (${D.bodyFont()})`;
    const color = D.resolveColor(d.color || '') || '';
    return `<div class="ms-dlg">
      ${isNew ? `<div class="dlg-grid"><div class="field"><label>Name</label><input id="ms-name" placeholder="e.g. Chapter heading" maxlength="40"></div><div class="field"><label>Style type</label><select id="ms-type"><option value="p">Paragraph</option><option value="c">Character</option></select></div></div>` : ''}
      <div class="ms-row"><div class="field" style="flex:1.6"><label>Font</label><button class="btn ms-font" id="ms-font" data-font="${esc(d.font || '')}" title="Choose a font"><span style="font-family:${esc(D.fontStack(D.resolveFont(d.font || '@body')))}">${esc(fontLabel(d.font))}</span>${icon('chevron-down', 14)}</button></div>
        <div class="field" style="width:80px"><label>Size (pt)</label><input id="ms-size" type="number" min="1" max="400" step="0.5" value="${pct(d.size)}" placeholder="${isChar ? 'inherit' : '11'}"></div>
        <div class="field" style="width:110px"><label>Colour</label><div class="ms-color"><input id="ms-color" type="color" value="${/^#[0-9a-f]{6}$/i.test(color) ? color : '#111111'}"><label class="check" title="Use the theme colour / inherit"><input type="checkbox" id="ms-color-auto" ${d.color ? '' : 'checked'}> Auto</label></div></div></div>
      <div class="ms-toggles"><label class="check"><input type="checkbox" id="ms-b" ${d.bold ? 'checked' : ''}> <b>Bold</b></label><label class="check"><input type="checkbox" id="ms-i" ${d.italic ? 'checked' : ''}> <i>Italic</i></label><label class="check"><input type="checkbox" id="ms-u" ${d.underline ? 'checked' : ''}> <u>Underline</u></label><label class="check"><input type="checkbox" id="ms-caps" ${d.caps ? 'checked' : ''}> ALL CAPS</label><label class="check"><input type="checkbox" id="ms-sc" ${d.smallcaps ? 'checked' : ''}> Small caps</label></div>
      <div class="ms-para" ${isChar ? 'hidden' : ''}>
        <div class="ms-row"><div class="field"><label>Alignment</label><select id="ms-align">${['left', 'center', 'right', 'justify'].map(a => `<option value="${a}" ${(d.align || 'left') === a ? 'selected' : ''}>${a[0].toUpperCase() + a.slice(1)}</option>`).join('')}</select></div>
          <div class="field"><label>Before (pt)</label><input id="ms-before" type="number" min="0" max="200" step="1" value="${pct(d.before)}" placeholder="default"></div>
          <div class="field"><label>After (pt)</label><input id="ms-after" type="number" min="0" max="200" step="1" value="${pct(d.after)}" placeholder="default"></div>
          <div class="field"><label>Line spacing</label><select id="ms-line"><option value="">Document default</option>${[1, 1.15, 1.3, 1.5, 2, 2.5, 3].map(x => `<option value="${x}" ${d.line === x ? 'selected' : ''}>${x}</option>`).join('')}${d.line && ![1, 1.15, 1.3, 1.5, 2, 2.5, 3].includes(d.line) ? `<option value="${d.line}" selected>${d.line}</option>` : ''}</select></div></div>
        <div class="ms-row"><div class="field"><label>Indent left (cm)</label><input id="ms-indent" type="number" min="0" max="20" step="0.1" value="${d.indent || 0}"></div>
          <div class="field"><label>First line (cm)</label><input id="ms-first" type="number" min="-10" max="10" step="0.1" value="${d.firstLine || 0}" title="Use a negative number for a hanging indent"></div></div>
      </div>
      <div class="ms-prev-wrap"><div class="ms-prev doc-content" id="ms-prev"><p class="ms-ghost">Previous paragraph previous paragraph previous paragraph.</p><p id="ms-sample">${isChar ? 'Normal text with <span id="ms-sample-c">styled words in it</span> for preview.' : 'Sample text — The quick brown fox jumps over the lazy dog. यह नमूना पाठ है।'}</p><p class="ms-ghost">Following paragraph following paragraph following paragraph.</p></div></div>
      <div class="hint" id="ms-desc"></div>
    </div>`;
  }
  function readForm(rt) {
    const num = (sel) => { const v = $(sel, rt).value.trim(); return v === '' ? undefined : +v; };
    const o = { font: $('#ms-font', rt).dataset.font || undefined, size: num('#ms-size'), color: $('#ms-color-auto', rt).checked ? undefined : $('#ms-color', rt).value.toUpperCase(), bold: $('#ms-b', rt).checked, italic: $('#ms-i', rt).checked, underline: $('#ms-u', rt).checked, caps: $('#ms-caps', rt).checked, smallcaps: $('#ms-sc', rt).checked };
    if (!$('.ms-para', rt).hidden) Object.assign(o, { align: $('#ms-align', rt).value, before: num('#ms-before'), after: num('#ms-after'), line: $('#ms-line', rt).value ? +$('#ms-line', rt).value : undefined, indent: +$('#ms-indent', rt).value || 0, firstLine: +$('#ms-first', rt).value || 0 });
    return o;
  }
  function wireForm(rt, isCharFn) {
    const upd = () => {
      const o = readForm(rt); const isChar = isCharFn();
      const target = isChar ? $('#ms-sample-c', rt) : $('#ms-sample', rt); if (!target) return;
      const pv = $('#ms-prev', rt); pv.style.setProperty('--th-accent', D.accent()); pv.style.setProperty('--th-dark', D.darkColor()); pv.style.setProperty('--th-head', D.fontStack(D.headFont())); pv.style.setProperty('--th-body', D.fontStack(D.bodyFont())); pv.style.setProperty('--th-mono', D.fontStack(D.monoFont()));
      const x = Object.assign({}, o); if (x.line === undefined && !isChar) x.line = null; if (x.before === undefined && !isChar) x.before = null; if (x.after === undefined && !isChar) x.after = null;
      if (!isChar && !x.color) x.color = '#111111';
      target.setAttribute('style', declare(Object.fromEntries(Object.entries(x).filter(([, v]) => v !== undefined)), isChar));
      $('#ms-desc', rt).textContent = [D.resolveFont(o.font || '@body'), (o.size || 'inherited') + ' pt', o.bold && 'bold', o.italic && 'italic', !isChar && `space ${o.before ?? 'default'}/${o.after ?? 'default'} pt`, !isChar && `line ${o.line || 'default'}`].filter(Boolean).join(' · ');
    };
    rt.addEventListener('input', upd); rt.addEventListener('change', upd);
    $('#ms-font', rt).addEventListener('click', async (e) => {
      const b = e.currentTarget; const cur = D.resolveFont(b.dataset.font || '@body');
      const f = await BTS.fontPicker(b, { current: cur, theme: [{ label: 'Headings', name: D.headFont() }, { label: 'Body', name: D.bodyFont() }] });
      if (!f) return;
      b.dataset.font = f === D.headFont() && b.dataset.font === '@head' ? '@head' : f === D.bodyFont() && b.dataset.font === '@body' ? '@body' : f;
      b.querySelector('span').textContent = f; b.querySelector('span').style.fontFamily = D.fontStack(f); upd();
    });
    const auto = $('#ms-color-auto', rt); $('#ms-color', rt).addEventListener('input', () => { auto.checked = false; });
    const ty = $('#ms-type', rt); if (ty) ty.addEventListener('change', () => { $('.ms-para', rt).hidden = ty.value === 'c'; $('#ms-sample', rt).innerHTML = ty.value === 'c' ? 'Normal text with <span id="ms-sample-c">styled words in it</span> for preview.' : 'Sample text — The quick brown fox jumps over the lazy dog.'; upd(); });
    setTimeout(upd);
  }
  function diffDef(id, o) {
    // store only what differs from the style-set default so later style-set changes still flow through
    const st = D.findStyle(id); const base = Object.assign({}, st.def, st.custom ? {} : ((D.styleSet().s || {})[id] || {}));
    const out = {};
    Object.entries(o).forEach(([k, v]) => {
      if (v === undefined) return; // "default" / "auto" → keep the style-set value
      const bv = base[k];
      if (bv === v || (bv === undefined && (v === false || v === 0 || (k === 'align' && v === 'left')))) return;
      if (k === 'color' && bv && D.resolveColor(bv).toUpperCase() === String(v).toUpperCase()) return;
      out[k] = v;
    });
    return out;
  }
  D.modifyStyleDialog = async (id) => {
    const st = D.findStyle(id); if (!st) return; const sv = D.saveSel();
    const d = D.styleDef(id); const isChar = st.type === 'c';
    const v = await BTS.dialog({ title: `Modify style — ${st.name}`, width: 620, body: styleForm(id, d, isChar, false), onOpen: (rt) => wireForm(rt, () => isChar), buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (rt) => readForm(rt) }] });
    D.restoreSel(sv); if (!v) return;
    D.change(() => {
      if (st.custom) { const u = D.design.customStyles.find(x => x.id === id); u.def = Object.fromEntries(Object.entries(v).filter(([, x]) => x !== undefined && x !== false)); }
      else { const mods = Object.assign({}, D.design.styleMods || {}); const df = diffDef(id, v); Object.keys(df).forEach(k => df[k] === undefined && delete df[k]); if (Object.keys(df).length) mods[id] = df; else delete mods[id]; D.design.styleMods = mods; }
      D.applySettings();
    });
    const n = $$(st.type === 'c' ? '.' + st.cls : 'p,h1,h2,h3,h4,h5,h6,pre,blockquote', D.doc).filter(b => st.type === 'c' || ((b.tagName === 'BLOCKQUOTE' || D.blockOf(b) === b) && D.styleOf(b) === id)).length;
    BTS.toast(`${st.name} updated${n ? ` — ${n} ${st.type === 'c' ? 'place' : 'paragraph'}${n === 1 ? '' : 's'} changed` : ''}`);
  };
  /** Capture the formatting at the caret as a definition. */
  function defFromSelection(isChar) {
    const r = D.ensureSel(); let el = r.startContainer; if (el.nodeType === 3) el = el.parentElement;
    const cs = getComputedStyle(el); const b = D.blockOf(r.startContainer); const bs = b ? getComputedStyle(b) : cs;
    const fam = cs.fontFamily.split(',')[0].trim().replace(/^["']|["']$/g, '');
    const hexOf = (c) => { const m = c.match(/rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)/); return m ? '#' + [m[1], m[2], m[3]].map(x => (+x).toString(16).padStart(2, '0')).join('').toUpperCase() : undefined; };
    const o = { font: fam === D.bodyFont() ? '@body' : fam === D.headFont() ? '@head' : fam, size: Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2, color: hexOf(cs.color), bold: +cs.fontWeight >= 600, italic: cs.fontStyle === 'italic', underline: /underline/.test(cs.textDecorationLine), caps: cs.textTransform === 'uppercase', smallcaps: cs.fontVariantCaps === 'small-caps' };
    if (!isChar && b) { const fs = parseFloat(bs.fontSize) || 14.67; Object.assign(o, { align: bs.textAlign.replace('start', 'left').replace('end', 'right'), before: Math.round(parseFloat(bs.marginTop) * 0.75), after: Math.round(parseFloat(bs.marginBottom) * 0.75), line: bs.lineHeight === 'normal' ? 1.15 : Math.round(parseFloat(bs.lineHeight) / fs * 100) / 100, indent: +(parseFloat(bs.marginLeft) / 96 * 2.54).toFixed(2), firstLine: +(parseFloat(bs.textIndent) / 96 * 2.54).toFixed(2) }); }
    return { o, block: b };
  }
  D.updateStyleFromSelection = (id) => {
    const st = D.findStyle(id); const { o } = defFromSelection(st.type === 'c');
    D.change(() => {
      if (st.custom) { const u = D.design.customStyles.find(x => x.id === id); u.def = o; }
      else { const mods = Object.assign({}, D.design.styleMods || {}); mods[id] = diffDef(id, o); Object.keys(mods[id]).forEach(k => mods[id][k] === undefined && delete mods[id][k]); D.design.styleMods = mods; }
      D.applySettings();
    });
    BTS.toast(`${st.name} now matches the selection`);
  };
  D.newStyleDialog = async () => {
    const sv = D.saveSel(); const { o } = defFromSelection(false);
    const v = await BTS.dialog({ title: 'Create a new style from formatting', width: 620, body: styleForm(null, o, false, true), onOpen: (rt) => wireForm(rt, () => $('#ms-type', rt).value === 'c'), buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, validate: (rt) => { const n = $('#ms-name', rt).value.trim(); if (!n) { BTS.toast('Give the style a name.', { error: true }); return false; } if (D.allStyles().some(s => s.name.toLowerCase() === n.toLowerCase())) { BTS.toast('A style with that name already exists.', { error: true }); return false; } return true; }, value: (rt) => ({ name: $('#ms-name', rt).value.trim(), type: $('#ms-type', rt).value, def: readForm(rt) }) }] });
    D.restoreSel(sv); if (!v) return;
    const id = 'u' + Math.random().toString(36).slice(2, 8);
    const def = Object.fromEntries(Object.entries(v.def).filter(([, x]) => x !== undefined && x !== false));
    if (v.type === 'c') ['align', 'before', 'after', 'line', 'indent', 'firstLine'].forEach(k => delete def[k]);
    D.change(() => { D.design.customStyles = (D.design.customStyles || []).concat([{ id, name: v.name, type: v.type, def }]); D.applySettings(); });
    // apply it to the selection, clearing the direct formatting it now captures
    if (v.type === 'c') D.applyCharStyle(id);
    else { D.setBlockStyle(id); D.change(() => D.blocks().forEach(b => { ['textAlign', 'marginTop', 'marginBottom', 'lineHeight', 'marginLeft', 'textIndent'].forEach(k => b.style[k] = ''); if (!b.getAttribute('style')) b.removeAttribute('style'); })); }
    BTS.toast(`Style “${v.name}” created`);
  };

  D.initStyles = () => {
    syncStyleList(); D.applyStyleSheet(); D.renderStylesGallery(); D.renderStyleSets(); wireStylesPane();
  };
})();
