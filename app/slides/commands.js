/* BTS Slides — ribbon commands, text/shape formatting, galleries, menus, dialogs and task panes. */
(function () {
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const esc = (s) => BTS.esc(s);
  const f2 = (n) => Math.round(n * 100) / 100;
  const P = () => SL.pres, S = () => SL.slide();
  const C = window.CMD = {};

  /* ====================================================================== */
  /*  Text formatting                                                        */
  /* ====================================================================== */
  const isTextual = (el) => el && (el.type === 'text' || (el.type === 'shape' && !SL.isLine(el)) || el.type === 'table');
  C.textTargets = () => E.selected().filter(isTextual);
  const editing = () => !!E.editing && !!E.editNode();
  function exec(cmd, val) { document.execCommand('styleWithCSS', false, true); document.execCommand(cmd, false, val); E.syncEditing(); E.renderOverlay(); E.refreshThumb(); BTS.markDirty(); E.updateRibbon(); }

  // Walk element html and remove inline properties / unwrap tags (used when formatting whole objects).
  function stripHTML(html, props = [], tags = []) {
    const d = document.createElement('div'); d.innerHTML = html || '';
    d.querySelectorAll('*').forEach(n => { props.forEach(p => n.style && n.style.removeProperty(p)); if (n.getAttribute('style') === '') n.removeAttribute('style'); });
    const unwrap = (n) => { n.replaceWith(...n.childNodes); };
    d.querySelectorAll(tags.concat(['span:not([style]):not([data-field])']).join(',')).forEach(unwrap);
    return d.innerHTML;
  }
  function eachText(fn) {
    const els = C.textTargets(); if (!els.length) return false;
    els.forEach(el => { fn(el); if (el.type === 'table') el.rows.forEach(r => r.forEach(c => { if (c._fn) c._fn = null; })); });
    E.renderCanvas(); SL.commit(); E.refreshThumb(true); E.updateRibbon(); return true;
  }
  function eachCell(el, fn) { if (el.type === 'table') el.rows.forEach(r => r.forEach(c => { c.html = fn(c.html || ''); })); else el.html = fn(el.html || ''); }

  const TOGGLE = { bold: ['bold', 'font-weight', ['b', 'strong']], italic: ['italic', 'font-style', ['i', 'em']], underline: ['underline', 'text-decoration', ['u']], strike: ['strike', 'text-decoration', ['s', 'strike', 'del']] };
  C.toggleStyle = (k) => {
    if (editing()) return exec(k === 'strike' ? 'strikeThrough' : k);
    const els = C.textTargets(); if (!els.length) return;
    const on = !C.elFlag(els[0], k);
    eachText(el => { el[k] = on; eachCell(el, h => stripHTML(h, [TOGGLE[k][1], ...(k === 'underline' || k === 'strike' ? ['text-decoration-line'] : [])], TOGGLE[k][2])); });
  };
  C.elFlag = (el, k) => k === 'bold' ? !!(el.bold ?? (el.role === 'title' && SL.themeOf(P()).titleBold)) : !!el[k];

  function sizeAtCaret() {
    const s = getSelection(); if (!s.rangeCount) return null;
    let n = s.anchorNode; if (n && n.nodeType === 3) n = n.parentElement; if (!n) return null;
    return Math.round(parseFloat(getComputedStyle(n).fontSize) * .75 * 2) / 2;
  }
  function fontAtCaret() { const s = getSelection(); let n = s.anchorNode; if (n && n.nodeType === 3) n = n.parentElement; return n ? getComputedStyle(n).fontFamily.split(',')[0].replace(/["']/g, '').trim() : ''; }

  C.setFontSize = (pt) => {
    pt = Math.max(1, Math.min(400, +pt)); if (!pt) return;
    if (editing()) {
      const root = E.editNode();
      document.execCommand('styleWithCSS', false, false); document.execCommand('fontSize', false, '7');
      const made = [];
      root.querySelectorAll('font[size="7"]').forEach(f => { const sp = document.createElement('span'); sp.style.fontSize = pt + 'pt'; sp.append(...f.childNodes); f.replaceWith(sp); made.push(sp); });
      root.querySelectorAll('span').forEach(sp => { if (sp.style.fontSize === 'xxx-large' || sp.style.fontSize === '-webkit-xxx-large') { sp.style.fontSize = pt + 'pt'; made.push(sp); } });
      // nested spans inside the new size span keep an older size: drop it
      made.forEach(sp => sp.querySelectorAll('[style*="font-size"]').forEach(n => n.style.removeProperty('font-size')));
      if (made.length) { const r = document.createRange(); r.setStartBefore(made[0]); r.setEndAfter(made[made.length - 1]); const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
      E.syncEditing(); E.renderOverlay(); E.refreshThumb(); BTS.markDirty(); E.updateRibbon(); return;
    }
    eachText(el => { el.size = pt; el.fontScale = 1; eachCell(el, h => stripHTML(h, ['font-size'])); });
  };
  C.stepFont = (dir) => {
    const step = (v) => { const i = SL.SIZES.findIndex(s => s >= v - .01); if (dir > 0) return SL.SIZES.find(s => s > v + .01) || v + 8; return [...SL.SIZES].reverse().find(s => s < v - .01) || Math.max(1, v - 1); void i; };
    if (editing()) { const cur = sizeAtCaret() || 18; return C.setFontSize(step(cur)); }
    eachText(el => {
      const ns = step(el.size || 18); const k = ns / (el.size || 18); el.size = ns;
      eachCell(el, h => h.replace(/font-size:\s*([\d.]+)pt/g, (_, v) => 'font-size: ' + f2(v * k) + 'pt'));
    });
  };
  C.setFont = (name) => {
    if (!name) return;
    if (editing()) return exec('fontName', name);
    eachText(el => { el.font = name; eachCell(el, h => stripHTML(h, ['font-family'])); });
  };
  C.setColor = (c) => {
    if (editing()) return exec('foreColor', c === 'auto' ? SL.textDefaults(SL.findEl(E.editing), P()).color : c);
    eachText(el => { if (c === 'auto') delete el.color; else el.color = c; eachCell(el, h => stripHTML(h, ['color'])); if (el.type === 'table') el.rows.forEach(r => r.forEach(cell => { if (c === 'auto') delete cell.color; })); });
  };
  C.setHighlight = (c) => {
    if (editing()) return exec('hiliteColor', c === 'none' ? 'transparent' : c);
    eachText(el => eachCell(el, h => {
      const stripped = stripHTML(h, ['background-color']); if (c === 'none') return stripped;
      const d = document.createElement('div'); d.innerHTML = stripped; d.querySelectorAll('p,li').forEach(b => { const sp = document.createElement('span'); sp.style.backgroundColor = c; sp.append(...b.childNodes); b.appendChild(sp); }); return d.innerHTML;
    }));
  };
  C.clearFormat = () => {
    if (editing()) { exec('removeFormat'); return; }
    eachText(el => {
      ['font', 'color', 'bold', 'italic', 'underline', 'strike', 'effect', 'caps', 'spacing', 'highlight'].forEach(k => delete el[k]);
      eachCell(el, h => stripHTML(h, ['color', 'background-color', 'font-size', 'font-family', 'font-weight', 'font-style', 'text-decoration', 'text-decoration-line'], ['b', 'strong', 'i', 'em', 'u', 's', 'strike', 'del', 'sub', 'sup']));
    });
  };

  // blocks (p/li) touched by the current selection inside the edit node
  function selBlocks() {
    const root = E.editNode(); const s = getSelection(); if (!root || !s.rangeCount) return [];
    const r = s.getRangeAt(0); const all = $$('p,li', root);
    const hit = all.filter(b => r.intersectsNode(b));
    if (!hit.length) { let n = r.startContainer; if (n.nodeType === 3) n = n.parentElement; const b = n.closest('p,li'); if (b) hit.push(b); }
    return hit;
  }
  C.setAlign = (a) => {
    if (editing()) { const blocks = selBlocks(); blocks.forEach(b => b.style.textAlign = a); if (!blocks.length) exec({ left: 'justifyLeft', center: 'justifyCenter', right: 'justifyRight', justify: 'justifyFull' }[a]); E.syncEditing(); E.refreshThumb(); BTS.markDirty(); E.updateRibbon(); return; }
    eachText(el => { if (el.type === 'table') { const c = E.tableCell; el.rows.forEach(r => r.forEach(cell => cell.align = a)); void c; return; } el.align = a; el.html = stripHTML(el.html, ['text-align']); });
  };
  C.setList = (kind) => { // 'ul' | 'ol' | null(toggle)
    const conv = (root, blocks) => {
      // blocks: p/li elements to convert
      const allList = blocks.length && blocks.every(b => b.tagName === 'LI' && b.parentElement.tagName.toLowerCase() === kind);
      if (allList) { // remove list → paragraphs
        blocks.forEach(li => { const p = document.createElement('p'); p.innerHTML = li.innerHTML; if (li.getAttribute('style')) p.setAttribute('style', li.getAttribute('style')); if (li.dataset.lvl) p.dataset.lvl = li.dataset.lvl; splitOut(li, p); });
      } else {
        blocks.forEach(b => {
          if (b.tagName === 'LI') { const list = b.parentElement; if (list.tagName.toLowerCase() !== kind) { const nl = document.createElement(kind); [...list.attributes].forEach(a => nl.setAttribute(a.name, a.value)); nl.append(...list.childNodes); list.replaceWith(nl); } return; }
          const li = document.createElement('li'); li.innerHTML = b.innerHTML; if (b.getAttribute('style')) li.setAttribute('style', b.getAttribute('style')); if (b.dataset.lvl) li.dataset.lvl = b.dataset.lvl;
          const prev = b.previousElementSibling;
          if (prev && prev.tagName.toLowerCase() === kind) { prev.appendChild(li); b.remove(); }
          else { const l = document.createElement(kind); l.appendChild(li); b.replaceWith(l); }
        });
        // merge adjacent lists of same kind
        $$(kind, root).forEach(l => { const n = l.nextElementSibling; if (n && n.tagName === l.tagName) { l.append(...n.childNodes); n.remove(); } });
      }
      $$('ul,ol', root).forEach(l => { if (!l.children.length) l.remove(); });
    };
    function splitOut(li, p) {
      const list = li.parentElement; const after = document.createElement(list.tagName.toLowerCase());
      let n = li.nextSibling; while (n) { const nx = n.nextSibling; after.appendChild(n); n = nx; }
      list.after(p); if (after.children.length) p.after(after); li.remove();
    }
    if (editing()) {
      const root = E.editNode(); const blocks = selBlocks(); if (!blocks.length) return;
      const s = getSelection(); const r = s.rangeCount ? s.getRangeAt(0) : null; const off = r ? { n: r.startContainer, o: r.startOffset } : null;
      conv(root, blocks);
      try { if (off && root.contains(off.n)) { const rr = document.createRange(); rr.setStart(off.n, off.o); rr.collapse(true); s.removeAllRanges(); s.addRange(rr); } } catch {}
      E.syncEditing(); E.refreshThumb(); BTS.markDirty(); E.updateRibbon(); E.renderOverlay(); return;
    }
    eachText(el => { if (el.type === 'table') return; const d = document.createElement('div'); d.innerHTML = el.html || '<p><br></p>'; conv(d, $$('p,li', d)); el.html = d.innerHTML; el.list = SL.isEmptyText(el) ? kind : el.list; });
  };
  C.indent = (dir) => {
    const bump = (blocks) => blocks.forEach(b => { const v = Math.max(0, Math.min(4, (+b.dataset.lvl || 0) + dir)); if (v) b.dataset.lvl = v; else delete b.dataset.lvl; });
    if (editing()) { bump(selBlocks()); E.syncEditing(); E.refreshThumb(); BTS.markDirty(); E.renderOverlay(); return; }
    eachText(el => { if (el.type === 'table') return; const d = document.createElement('div'); d.innerHTML = el.html; bump($$('p,li', d)); el.html = d.innerHTML; });
  };
  C.setElProp = (k, v, types) => {
    const els = E.selected().filter(e => !types || types.includes(e.type)); if (!els.length) return false;
    els.forEach(e => { if (v === undefined) delete e[k]; else e[k] = v; });
    if (E.editing) E.syncEditing();
    E.renderCanvas(); SL.commit(); E.refreshThumb(true); E.updateRibbon(); E.refreshPane && E.refreshPane('sel'); return true;
  };
  C.changeCase = (mode) => {
    const tf = (t) => mode === 'upper' ? t.toUpperCase() : mode === 'lower' ? t.toLowerCase() : mode === 'title' ? t.toLowerCase().replace(/(^|\s)(\p{L})/gu, (m, a, b) => a + b.toUpperCase()) : mode === 'sentence' ? t.toLowerCase().replace(/(^\s*|[.!?]\s+)(\p{L})/gu, (m, a, b) => a + b.toUpperCase()) : [...t].map(c => c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase()).join('');
    if (editing()) { const s = getSelection(); if (!s.rangeCount || s.isCollapsed) return BTS.toast('Select some text first'); document.execCommand('insertText', false, tf(s.toString())); E.syncEditing(); return; }
    eachText(el => eachCell(el, h => { const d = document.createElement('div'); d.innerHTML = h; const w = document.createTreeWalker(d, NodeFilter.SHOW_TEXT); let n; while ((n = w.nextNode())) n.nodeValue = tf(n.nodeValue); return d.innerHTML; }));
  };

  /* ====================================================================== */
  /*  Shape formatting                                                      */
  /* ====================================================================== */
  C.shapeTargets = () => E.selected().filter(e => e.type === 'shape' || e.type === 'text' || e.type === 'image' || e.type === 'icon' || e.type === 'table' || e.type === 'chart');
  C.setFill = (c) => {
    const els = E.selected(); if (!els.length) return;
    els.forEach(e => { if (e.type === 'icon') { e.color = c === 'none' ? e.color : c; e.fillIcon = c !== 'none' && e.fillIcon; } else if (e.type === 'table') e.accent = c; else if (e.type !== 'image' && e.type !== 'chart') e.fill = c; });
    E.renderCanvas(); SL.commit(); E.refreshThumb(true); E.updateRibbon(); E.refreshPane && E.refreshPane('sel');
  };
  C.QUICK = (pres) => {
    const out = [];
    for (let i = 1; i <= 6; i++) {
      const a = 'accent' + i;
      out.push({ fill: a, stroke: 'none', color: '#ffffff', shadow: false, label: 'Solid' });
      out.push({ fill: { type: 'gradient', color: a, color2: shade(SL.color(a, pres), -.35), angle: 90 }, stroke: 'none', color: '#ffffff', label: 'Gradient' });
      out.push({ fill: 'none', stroke: a, strokeWidth: 2, color: a, label: 'Outline' });
      out.push({ fill: tint(SL.color(a, pres), .82), stroke: a, strokeWidth: 1, color: shade(SL.color(a, pres), -.35), label: 'Light' });
      out.push({ fill: a, stroke: 'none', color: '#ffffff', shadow: true, label: 'Raised' });
    }
    return out;
  };
  function hex2rgb(h) { h = h.replace('#', ''); if (h.length === 3) h = h.split('').map(c => c + c).join(''); return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]; }
  function rgb2hex(r) { return '#' + r.map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join(''); }
  function tint(h, k) { return rgb2hex(hex2rgb(h).map(v => v + (255 - v) * k)); }
  function shade(h, k) { return rgb2hex(hex2rgb(h).map(v => v * (1 + k))); }
  C.tint = tint; C.shade = shade; C.hex2rgb = hex2rgb; C.rgb2hex = rgb2hex;
  C.applyQuick = (st) => {
    const els = E.selected().filter(e => e.type === 'shape' || e.type === 'text'); if (!els.length) return BTS.toast('Select a shape or text box first');
    els.forEach(e => { e.fill = SL.clone(st.fill); e.stroke = st.stroke; e.strokeWidth = st.strokeWidth || 0; e.shadow = !!st.shadow; if (e.type === 'shape' || !SL.isEmptyText(e)) { if (!SL.isLine(e)) e.color = st.color; } if (SL.isLine(e)) { e.stroke = typeof st.fill === 'string' && st.fill !== 'none' ? st.fill : st.stroke; e.strokeWidth = Math.max(2, st.strokeWidth || 2); } });
    E.renderCanvas(); SL.commit(); E.refreshThumb(true);
  };

  /* ====================================================================== */
  /*  Popovers / galleries                                                  */
  /* ====================================================================== */
  function pop(anchor, html, cls, onClick) {
    const el = document.createElement('div'); el.className = cls || ''; el.innerHTML = html;
    el.addEventListener('mousedown', e => { if (!e.target.closest('input,select,textarea')) e.preventDefault(); });
    if (onClick) el.addEventListener('click', e => onClick(e, el));
    const pr = BTS.popover(anchor, el); el.closed = pr; return el;
  }
  C.pop = pop;

  C.layoutPreview = (layout, w) => {
    const s = SL.newSlide(P(), layout); s.background = null;
    s.elements.forEach(e => { if (e.type === 'text') { e.html = `<p>${e.ph === 'title' ? 'Title' : e.ph === 'subtitle' ? 'Subtitle' : 'Text'}</p>`; e.stroke = '#9aa3b2'; e.dash = 'dash'; e.strokeWidth = 1.5; e.color = '#7a8499'; e.list = null; } if (e.type === 'image') { e.type = 'shape'; e.kind = 'rect'; e.fill = '#e6e9ee'; e.stroke = '#9aa3b2'; e.dash = 'dash'; e.strokeWidth = 1.5; e.html = '<p>Picture</p>'; e.color = '#7a8499'; } });
    return SL.thumbHTML(s, P(), w, 0);
  };
  C.layoutPicker = (anchor, onPick, current) => {
    const html = `<div class="menu-head">${esc(SL.themeOf(P()).name)}</div><div class="layout-grid">${SL.LAYOUTS.map(l => `<button class="layout-card ${l.id === current ? 'on' : ''}" data-l="${l.id}">${C.layoutPreview(l.id, 120)}<span>${esc(l.name)}</span></button>`).join('')}</div>` +
      (anchor.dataset && anchor.dataset.cmd === 'new-slide-menu' ? `<div class="menu-sep"></div><div class="menu-item" data-act="dup">${icon('copy-plus', 16)}<span>Duplicate selected slides</span></div><div class="menu-item" data-act="reuse">${icon('folder-open', 16)}<span>Reuse slides from a file…</span></div>` : '');
    pop(anchor, html, 'layout-pop', (e) => {
      const b = e.target.closest('[data-l]'); if (b) { BTS.closePopover(); onPick(b.dataset.l); }
      const a = e.target.closest('[data-act]'); if (a) { BTS.closePopover(); if (a.dataset.act === 'dup') E.duplicateSlides(); else C.reuseSlides(); }
    });
  };
  C.shapePicker = (anchor, onPick) => {
    const groups = [['Lines', ['line', 'arrowLine']], ['Rectangles', ['rect', 'roundRect']], ['Basic shapes', ['ellipse', 'triangle', 'rtTriangle', 'diamond', 'parallelogram', 'trapezoid', 'pentagon', 'hexagon', 'octagon', 'plus', 'donut', 'heart', 'cloud', 'can']],
      ['Block arrows', ['rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'chevron', 'homePlate']], ['Stars', ['star4', 'star5', 'star6']], ['Callouts', ['wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout']]];
    const name = Object.fromEntries(SL.SHAPES);
    const html = groups.map(([g, ks]) => `<div class="menu-head">${g}</div><div class="shape-grid">${ks.map(k => `<button data-k="${k}" title="${esc(name[k])}">${shapeIcon(k)}</button>`).join('')}</div>`).join('');
    pop(anchor, html, 'shape-pop', (e) => { const b = e.target.closest('[data-k]'); if (b) { BTS.closePopover(); onPick(b.dataset.k); } });
  };
  function shapeIcon(k) {
    const el = { id: 'i' + k, kind: k, fill: k === 'line' || k === 'arrowLine' ? 'none' : 'none', stroke: '#47526a', strokeWidth: 1.1, type: 'shape' };
    const box = k === 'line' || k === 'arrowLine' ? [18, 18] : k.includes('Callout') ? [20, 14] : ['rightArrow', 'leftArrow', 'leftRightArrow', 'chevron', 'homePlate', 'parallelogram', 'trapezoid', 'rect', 'roundRect'].includes(k) ? [20, 14] : ['upArrow', 'downArrow', 'can'].includes(k) ? [13, 20] : [18, 18];
    return `<span class="shp-ic" style="width:${box[0]}px;height:${box[1]}px">${SL.shapeSVG(el, P(), box[0], box[1])}</span>`;
  }
  C.shapeIcon = shapeIcon;

  C.buildGalleries = () => {
    // home mini shapes
    const hs = $('#home-shapes'); if (hs) hs.innerHTML = ['textbox', 'rect', 'ellipse', 'triangle', 'line', 'arrowLine', 'roundRect', 'rightArrow', 'star5', 'more'].map(k => k === 'textbox' ? `<button class="rb" data-cmd="insert-textbox" title="Text box">${icon('type', 16)}</button>` : k === 'more' ? `<button class="rb" data-cmd="shapes-menu" title="More shapes">${icon('chevron-down', 14)}</button>` : `<button class="rb" data-cmd="draw-shape" data-kind="${k}" title="${esc(Object.fromEntries(SL.SHAPES)[k])}">${shapeIcon(k)}</button>`).join('');
    C.buildThemeGallery(); C.buildFxGalleries();
  };
  C.themeSample = (theme, variant, w) => {
    const p = { width: SL.W(P()), height: SL.H(P()), theme: { id: theme.id, variant: variant || 0 }, slides: [] };
    const s = SL.newSlide(p, 'title');
    s.elements[0].html = '<p>Aa</p>'; s.elements[0].size = 120; s.elements[1].html = `<p>${esc(theme.name)}</p>`; s.elements[1].size = 40;
    const pal = SL.palette(p);
    return SL.thumbHTML(s, p, w, 0) + `<div class="th-dots">${pal.accents.slice(0, 4).map(c => `<i style="background:${c}"></i>`).join('')}</div>`;
  };
  C.buildThemeGallery = () => {
    const g = $('#theme-gal'); if (!g) return;
    if (!g._wheel) { g._wheel = true; g.addEventListener('wheel', (e) => { if (Math.abs(e.deltaY) > Math.abs(e.deltaX)) { g.scrollLeft += e.deltaY; e.preventDefault(); } }, { passive: false }); }
    g.innerHTML = SL.THEMES.map(t => `<button class="th-card ${P().theme.id === t.id ? 'on' : ''}" data-cmd="apply-theme" data-theme-id="${t.id}" title="${esc(t.name)}">${C.themeSample(t, 0, 92)}</button>`).join('');
    const v = $('#variant-gal'); const th = SL.themeOf(P());
    if (window.SLX && SLX.buildStyleGallery) SLX.buildStyleGallery();
    v.innerHTML = (th.variants || [{}]).map((_, i) => `<button class="th-card var ${(P().theme.variant || 0) === i ? 'on' : ''}" data-cmd="apply-variant" data-v="${i}" title="Variant ${i + 1}">${C.themeSample(th, i, 72)}</button>`).join('');
  };
  C.TRANS = [['none', 'None', 'ban'], ['fade', 'Fade', 'blend'], ['push', 'Push', 'arrow-big-left'], ['wipe', 'Wipe', 'panel-left-open'], ['split', 'Split', 'columns-2'], ['reveal', 'Reveal', 'panel-right-open'], ['cover', 'Cover', 'layers'], ['zoom', 'Zoom', 'zoom-in'], ['flip', 'Flip', 'flip-horizontal-2'], ['cube', 'Cube', 'box'], ['morph', 'Morph', 'sparkles']];
  C.ANIMS = [['none', 'None', 'ban'], ['appear', 'Appear', 'square'], ['fade', 'Fade', 'blend'], ['fly', 'Fly In', 'move-up'], ['float', 'Float In', 'arrow-up-from-line'], ['zoom', 'Zoom', 'zoom-in'], ['wipe', 'Wipe', 'panel-left-open'], ['split', 'Split', 'columns-2'], ['bounce', 'Bounce', 'activity'], ['grow', 'Grow & Turn', 'rotate-cw'], ['swivel', 'Swivel', 'refresh-cw'], ['pulse', 'Pulse', 'heart-pulse']];
  C.buildFxGalleries = () => {
    const missing = (n) => (window.BTS_ICONS || {})[n] ? n : 'sparkles';
    $('#trans-gal').innerHTML = C.TRANS.map(([k, n, ic]) => `<button class="fx-card" data-cmd="set-transition" data-fx="${k}">${icon(missing(ic), 22)}<span>${n}</span></button>`).join('');
    $('#anim-gal').innerHTML = C.ANIMS.map(([k, n, ic]) => `<button class="fx-card anim" data-cmd="set-anim" data-fx="${k}">${icon(missing(ic), 22)}<span>${n}</span></button>`).join('');
  };

  /* ====================================================================== */
  /*  Ribbon state                                                           */
  /* ====================================================================== */
  E.onRibbonState = () => {
    const els = E.selected(); const el = els[0]; const ed = editing();
    const tels = C.textTargets(); const t = tels[0];
    // contextual tab
    const tabF = $('#tab-format');
    const showFmt = !!els.length && (E.view === 'normal' || E.view === 'master');
    tabF.classList.toggle('hidden', !showFmt);
    if (showFmt) {
      const types = new Set(els.map(e => SL.isLine(e) ? 'shape' : e.type));
      tabF.textContent = types.size === 1 ? ({ image: 'Picture Format', table: 'Table Design', chart: 'Chart Design', icon: 'Graphics Format' }[[...types][0]] || 'Shape Format') : 'Shape Format';
      $$('#ribbon [data-panel="format"] .rgroup').forEach(g => { const f = g.dataset.for.split(' '); g.classList.toggle('hidden', !(f.includes('any') || [...types].some(x => f.includes(x)))); });
      if (el) { const w = $('#size-w'), h = $('#size-h'); if (document.activeElement !== w) w.value = f2(el.w / 96); if (document.activeElement !== h) h.value = f2(el.h / 96); $('#size-lock').checked = !!el.lockAspect || el.type === 'image' || el.type === 'icon'; }
      if (el && el.type === 'table') { $('#tbl-header').checked = !!el.header; $('#tbl-banded').checked = !!el.banded; }
      $('#b-crop').classList.toggle('on', !!E.cropping);
    } else if (BTS.activeTab === 'format') BTS.showTab(E.view === 'master' ? 'master' : 'home');
    // font fields (the font name box is updated by styles.js)
    const fs = $('#font-size');
    if (document.activeElement !== fs) fs.value = ed ? (sizeAtCaret() || '') : t ? (t.size || 18) : '';
    const st = (k) => ed ? document.queryCommandState(k === 'strike' ? 'strikeThrough' : k) : t ? C.elFlag(t, k) : false;
    ['bold', 'italic', 'underline', 'strike'].forEach(k => $('#b-' + k).classList.toggle('on', st(k)));
    $('#b-tshadow').classList.toggle('on', !!(t && t.effect === 'shadow'));
    let align = t ? (t.align || 'left') : null;
    if (ed) { const b = selBlocks()[0]; if (b) align = getComputedStyle(b).textAlign.replace('start', 'left').replace('end', 'right'); }
    [['al', 'left'], ['ac', 'center'], ['ar', 'right'], ['aj', 'justify']].forEach(([i, a]) => $('#b-' + i).classList.toggle('on', align === a));
    let list = null; if (ed) { const b = selBlocks()[0]; if (b && b.tagName === 'LI') list = b.parentElement.tagName.toLowerCase(); } else if (t && t.html) { const m = /^<(ul|ol)/.exec(t.html); list = m && m[1]; }
    $('#b-ul').classList.toggle('on', list === 'ul'); $('#b-ol').classList.toggle('on', list === 'ol');
    if (el) { const fc = typeof el.fill === 'string' ? SL.color(el.fill, P()) : el.fill && el.fill.color ? SL.color(el.fill.color, P()) : null; if (fc && fc !== 'none') { $('#sw-fill').style.background = fc; $('#sw-fill2').style.background = fc; } }
    // slide-level
    $('#b-hide').classList.toggle('on', !!S().hidden);
    const tr = S().transition || { type: 'none' };
    $$('#trans-gal .fx-card').forEach(b => b.classList.toggle('on', b.dataset.fx === (tr.type || 'none')));
    const td = $('#trans-dur'); if (document.activeElement !== td) td.value = f2((tr.dur || 700) / 1000);
    $('#trans-auto-on').checked = tr.auto != null && tr.auto !== false; const ta = $('#trans-auto'); if (document.activeElement !== ta) ta.value = tr.auto ?? 0;
    $('#trans-click').checked = tr.click !== false;
    const an = el && el.anim;
    $$('#anim-gal .fx-card').forEach(b => { b.classList.toggle('on', !!el && b.dataset.fx === (an ? an.type : 'none')); b.disabled = !el; });
    ['anim-trigger', 'anim-dur', 'anim-delay'].forEach(id => $('#' + id).disabled = !an);
    if (an) { $('#anim-trigger').value = an.trigger || 'click'; const ad = $('#anim-dur'); if (document.activeElement !== ad) ad.value = f2((an.dur || 500) / 1000); const dl = $('#anim-delay'); if (document.activeElement !== dl) dl.value = f2((an.delay || 0) / 1000); }
    ['b-v-normal', 'b-v-sorter', 'b-v-notes'].forEach(id => $('#' + id).classList.toggle('on', id === 'b-v-' + E.view));
    $('#b-spell').classList.toggle('on', SL.spell !== false);
    $$('#theme-gal .th-card').forEach(b => b.classList.toggle('on', b.dataset.themeId === P().theme.id));
    $$('#variant-gal .th-card').forEach(b => b.classList.toggle('on', +b.dataset.v === (P().theme.variant || 0)));
    $('#b-anim-pane').classList.toggle('on', E.pane === 'anim'); $('#b-comments').classList.toggle('on', E.pane === 'comments');
    const a11 = $('#b-a11y'); if (a11) a11.classList.toggle('on', E.pane === 'a11y');
    (window.SLX && SLX.ribbon || []).forEach(f => { try { f(); } catch (err) { console.error(err); } });
  };

  /* ====================================================================== */
  /*  Dialogs                                                                */
  /* ====================================================================== */
  C.chartDialog = async (existing) => {
    const ch = existing ? SL.clone(existing) : { type: 'bar', title: 'Chart title', categories: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: 'Series 1', values: [12, 19, 15, 24] }, { name: 'Series 2', values: [8, 11, 17, 13] }] };
    const grid = () => `<table class="chart-grid"><tr><th></th>${ch.series.map((s, i) => `<th><input data-sn="${i}" value="${esc(s.name)}"></th>`).join('')}<th><button class="btn sm" data-add="col" title="Add series">+</button></th></tr>` +
      ch.categories.map((c, r) => `<tr><th><input data-cat="${r}" value="${esc(c)}"></th>${ch.series.map((s, i) => `<td><input type="number" data-v="${i},${r}" value="${s.values[r] ?? 0}"></td>`).join('')}<td></td></tr>`).join('') +
      `<tr><th><button class="btn sm" data-add="row" title="Add category">+</button></th></tr></table>`;
    const body = `<div class="field-row"><div class="field"><label>Chart type</label><select id="ch-type">${[['bar', 'Clustered column'], ['line', 'Line'], ['pie', 'Pie'], ['doughnut', 'Doughnut']].map(([k, n]) => `<option value="${k}" ${ch.type === k ? 'selected' : ''}>${n}</option>`).join('')}</select></div><div class="field"><label>Title</label><input id="ch-title" value="${esc(ch.title || '')}"></div></div>
      <div class="check"><input type="checkbox" id="ch-legend" ${ch.legend === false ? '' : 'checked'}><label for="ch-legend">Show legend</label></div>
      <div class="field"><label>Data (pie charts use the first series)</label><div id="ch-grid" class="modal-noenter">${grid()}</div></div>`;
    const read = (root) => {
      $$('[data-sn]', root).forEach(i => ch.series[+i.dataset.sn].name = i.value);
      $$('[data-cat]', root).forEach(i => ch.categories[+i.dataset.cat] = i.value);
      $$('[data-v]', root).forEach(i => { const [s, r] = i.dataset.v.split(',').map(Number); ch.series[s].values[r] = +i.value || 0; });
      ch.type = $('#ch-type', root).value; ch.title = $('#ch-title', root).value; ch.legend = $('#ch-legend', root).checked;
    };
    const r = await BTS.dialog({
      title: existing ? 'Edit chart data' : 'Insert chart', body, width: 560,
      buttons: [{ label: 'Cancel', value: null }, { label: existing ? 'Update' : 'Insert', primary: true, value: (root) => { read(root); return ch; } }],
      onOpen: (root) => root.addEventListener('click', e => { const a = e.target.closest('[data-add]'); if (!a) return; read(root); if (a.dataset.add === 'row') { ch.categories.push('New'); ch.series.forEach(s => s.values.push(0)); } else ch.series.push({ name: 'Series ' + (ch.series.length + 1), values: ch.categories.map(() => 0) }); $('#ch-grid', root).innerHTML = grid(); }),
    });
    return r;
  };
  C.iconDialog = async () => {
    const names = Object.keys(window.BTS_ICONS || {});
    const pick = await new Promise(resolve => {
      BTS.dialog({
        title: 'Insert icons', width: 640,
        body: `<div class="field"><input id="ic-q" placeholder="Search ${names.length} icons (e.g. book, rocket, leaf, user)"></div><div class="icon-grid" id="ic-grid"></div><div class="ic-hint">Click an icon to insert it. Icons from Lucide (ISC licence).</div>`,
        buttons: [{ label: 'Close', value: null }],
        onOpen: (root) => {
          const g = $('#ic-grid', root), q = $('#ic-q', root);
          const draw = () => { const s = q.value.trim().toLowerCase(); const list = (s ? names.filter(n => n.includes(s)) : names.filter((_, i) => i % 7 === 0)).slice(0, 240); g.innerHTML = list.map(n => `<button data-n="${n}" title="${n}">${icon(n, 26)}</button>`).join('') || '<div class="empty-note">No icons match.</div>'; };
          q.addEventListener('input', draw); draw();
          g.addEventListener('click', e => { const b = e.target.closest('[data-n]'); if (b) { resolve(b.dataset.n); root.querySelector('[data-x]').click(); } });
        },
      }).then(() => resolve(null));
    });
    return pick;
  };
  C.SYMBOLS = {
    'Math': '± × ÷ ≠ ≈ ≤ ≥ ∞ √ ∛ ∑ ∏ ∫ ∂ ∆ ∇ π θ λ μ σ φ ω α β γ δ ε Ω ° ′ ″ ∠ ⊥ ∥ ∈ ∉ ⊂ ⊃ ∪ ∩ ∀ ∃ ∴ ∵ ½ ⅓ ¼ ¾ ² ³ ⁿ ₀ ₁ ₂ ₃'.split(' '),
    'Arrows': '← → ↑ ↓ ↔ ↕ ⇐ ⇒ ⇔ ↗ ↘ ↙ ↖ ⟶ ⟵ ➜ ➤ ⮕'.split(' '),
    'Currency & marks': '₹ $ € £ ¥ ¢ © ® ™ § ¶ † ‡ • · … – — « » ‹ › “ ” ‘ ’'.split(' '),
    'Shapes & ticks': '✓ ✔ ✗ ✘ ★ ☆ ● ○ ■ □ ▲ △ ▼ ▽ ◆ ◇ ♥ ♦ ♣ ♠ ☐ ☑ ☒ ☀ ☁ ☂ ☎ ✉ ✈ ⚡ ♪ ♫'.split(' '),
    'Devanagari': 'ॐ । ॥ ० १ २ ३ ४ ५ ६ ७ ८ ९ ₹ अ आ इ ई उ ऊ ए ऐ ओ औ क ख ग घ'.split(' '),
  };
  C.symbolDialog = () => new Promise(resolve => {
    BTS.dialog({
      title: 'Symbol', width: 520, body: Object.entries(C.SYMBOLS).map(([g, l]) => `<div class="menu-head">${g}</div><div class="sym-grid">${l.map(s => `<button data-s="${esc(s)}">${esc(s)}</button>`).join('')}</div>`).join(''),
      buttons: [{ label: 'Close', value: null }],
      onOpen: (root) => root.addEventListener('click', e => { const b = e.target.closest('[data-s]'); if (b) { resolve(b.dataset.s); root.querySelector('[data-x]').click(); } }),
    }).then(() => resolve(null));
  });

  /* ====================================================================== */
  /*  Task panes                                                             */
  /* ====================================================================== */
  E.pane = null;
  E.openPane = (kind) => {
    if (E.pane === kind) { E.closePane(); return; }
    E.pane = kind; $('#side-pane').classList.remove('hidden'); E.refreshPane(); E.updateRibbon();
    requestAnimationFrame(E.layout);
  };
  E.closePane = () => { E.pane = null; $('#side-pane').classList.add('hidden'); E.updateRibbon(); requestAnimationFrame(E.layout); };
  E.refreshPane = (why) => {
    if (!E.pane) return;
    const body = $('#pane-body'), title = $('#pane-title');
    if (why === 'sel' && !['shape', 'anim', 'selection'].includes(E.pane)) return;
    if (body.contains(document.activeElement) && document.activeElement.matches('input[type=text],input:not([type]),textarea,input[type=number]') && why !== 'force') return;
    const fn = PANES[E.pane]; if (!fn) return;
    title.textContent = fn.title; body.innerHTML = fn.html(); fn.wire && fn.wire(body);
  };
  const row = (label, ctl) => `<div class="pf-row"><label>${label}</label><div>${ctl}</div></div>`;
  const swatchBtn = (act, color, label) => `<button class="pf-swatch" data-act="${act}" title="${label || 'Choose color'}"><i style="background:${color && color !== 'none' ? color : 'transparent'}"></i>${color === 'none' ? 'None' : ''}${icon('chevron-down', 12)}</button>`;
  const PANES = {
    bg: {
      title: 'Format Background',
      html: () => {
        const bg = S().background || Object.assign({ inherited: true }, SL.themeBg(P()));
        const ty = bg.type || 'solid';
        return `<div class="pf-sec">Fill</div>
          ${['solid', 'gradient', 'image'].map(t => `<label class="check"><input type="radio" name="bgt" value="${t}" ${ty === t ? 'checked' : ''}> ${({ solid: 'Solid fill', gradient: 'Gradient fill', image: 'Picture fill' })[t]}</label>`).join('')}
          <div class="pf-box">
          ${ty !== 'image' ? row('Color', swatchBtn('bg-c1', SL.color(bg.color, P()))) : ''}
          ${ty === 'gradient' ? row('Second color', swatchBtn('bg-c2', SL.color(bg.color2, P()))) + row('Angle', `<input type="range" min="0" max="360" step="5" value="${bg.angle ?? 135}" data-k="angle"> <span>${bg.angle ?? 135}°</span>`) : ''}
          ${ty === 'image' ? row('Picture', `<button class="btn" data-act="bg-img">${icon('image', 15)} Insert…</button>`) + row('Fit', `<select data-k="fit">${['cover', 'contain', 'stretch'].map(f => `<option ${bg.fit === f ? 'selected' : ''}>${f}</option>`).join('')}</select>`) : ''}
          </div>
          <label class="check"><input type="checkbox" data-act="bg-hide" ${S().hideGraphics ? 'checked' : ''}> Hide background graphics</label>
          <div class="pf-actions"><button class="btn" data-act="bg-reset">Reset background</button><button class="btn primary" data-act="bg-all">Apply to all</button></div>
          ${bg.inherited ? '<p class="pf-note">This slide uses the theme background.</p>' : ''}`;
      },
      wire: (b) => {
        const set = (fn) => { const bg = S().background ? S().background : SL.clone(SL.themeBg(P())); fn(bg); S().background = bg; E.renderCanvas(); E.refreshThumb(true); SL.commit(); E.refreshPane('force'); };
        $$('input[name=bgt]', b).forEach(r => r.addEventListener('change', () => set(bg => { bg.type = r.value; if (r.value === 'gradient' && !bg.color2) { bg.color2 = C.shade(SL.color(bg.color || 'bg', P()), -.25); bg.angle = 135; } if (r.value === 'image' && !bg.src) setTimeout(() => b.querySelector('[data-act="bg-img"]')?.click(), 0); })));
        b.addEventListener('input', e => { const k = e.target.dataset.k; if (k === 'angle') { e.target.nextElementSibling.textContent = e.target.value + '°'; const bg = S().background || SL.clone(SL.themeBg(P())); bg.angle = +e.target.value; S().background = bg; E.renderCanvas(); E.refreshThumb(); } });
        b.addEventListener('change', e => { const k = e.target.dataset.k; if (k === 'angle') SL.commit(); if (k === 'fit') set(bg => bg.fit = e.target.value); });
        b.addEventListener('click', async e => {
          const a = e.target.closest('[data-act]'); if (!a) return; const act = a.dataset.act;
          if (act === 'bg-c1' || act === 'bg-c2') { const c = await BTS.colorPicker(a, {}); if (c) set(bg => bg[act === 'bg-c1' ? 'color' : 'color2'] = c); }
          if (act === 'bg-img') { const f = await BTS.pickFile('image/*'); if (f) { const url = await BTS.bytesToDataURL(f.data, f.type || 'image/png'); set(bg => { bg.type = 'image'; bg.src = SL.addAsset(url); bg.fit = bg.fit || 'cover'; }); } }
          if (act === 'bg-reset') { S().background = null; S().hideGraphics = false; E.renderCanvas(); E.refreshThumb(true); SL.commit(); E.refreshPane('force'); }
          if (act === 'bg-all') { const bg = S().background; P().slides.forEach(s => { s.background = bg ? SL.clone(bg) : null; s.hideGraphics = S().hideGraphics; }); SL.commit(); E.renderAll(); BTS.toast('Background applied to all slides'); }
          if (act === 'bg-hide') { S().hideGraphics = a.checked; E.renderCanvas(); E.refreshThumb(true); SL.commit(); }
        });
      },
    },
    shape: {
      title: 'Format Shape',
      html: () => {
        const el = E.primary(); if (!el) return '<p class="pf-note">Select an object on the slide to format it.</p>';
        const fill = el.fill; const ft = !fill || fill === 'none' ? 'none' : typeof fill === 'object' ? 'gradient' : 'solid';
        const canFill = ['shape', 'text'].includes(el.type) && !SL.isLine(el);
        const isTx = el.type === 'text' || (el.type === 'shape' && !SL.isLine(el));
        const ins = el.inset || [9.6, 4.8, 9.6, 4.8];
        return `${canFill ? `<div class="pf-sec">Fill</div>
            ${['none', 'solid', 'gradient'].map(t => `<label class="check"><input type="radio" name="ft" value="${t}" ${ft === t ? 'checked' : ''}> ${({ none: 'No fill', solid: 'Solid fill', gradient: 'Gradient fill' })[t]}</label>`).join('')}
            <div class="pf-box">${ft === 'solid' ? row('Color', swatchBtn('fill-c', SL.color(fill, P()))) : ''}${ft === 'gradient' ? row('From', swatchBtn('fill-g1', SL.color(fill.color, P()))) + row('To', swatchBtn('fill-g2', SL.color(fill.color2, P()))) + row('Angle', `<input type="number" min="0" max="360" step="15" value="${fill.angle ?? 90}" data-k="gangle">`) : ''}</div>` : ''}
          ${el.type !== 'chart' ? `<div class="pf-sec">Line</div>
          ${row('Color', swatchBtn('line-c', SL.color(el.stroke, P()) || 'none'))}
          ${row('Width', `<input type="number" min="0" max="40" step="0.25" value="${el.strokeWidth ?? 0}" data-k="strokeWidth"> pt`)}
          ${row('Dash', `<select data-k="dash">${SL.DASHES.map(([k, n]) => `<option value="${k}" ${(el.dash || 'solid') === k ? 'selected' : ''}>${n}</option>`).join('')}</select>`)}` : ''}
          <div class="pf-sec">Effects</div>
          ${row('Transparency', `<input type="range" min="0" max="100" value="${Math.round((1 - (el.opacity ?? 1)) * 100)}" data-k="transp"> <span>${Math.round((1 - (el.opacity ?? 1)) * 100)}%</span>`)}
          <label class="check"><input type="checkbox" data-k="shadow" ${el.shadow ? 'checked' : ''}> Shadow</label>
          ${el.type === 'image' ? row('Corner rounding', `<input type="range" min="0" max="50" value="${el.radius || 0}" data-k="radius">`) : ''}
          <div class="pf-sec">Size &amp; position</div>
          <div class="pf-grid">
            <label>X <input type="number" step="0.05" value="${f2(el.x / 96)}" data-k="x"> in</label><label>Y <input type="number" step="0.05" value="${f2(el.y / 96)}" data-k="y"> in</label>
            <label>W <input type="number" step="0.05" min="0.01" value="${f2(el.w / 96)}" data-k="w"> in</label><label>H <input type="number" step="0.05" min="0.01" value="${f2(el.h / 96)}" data-k="h"> in</label>
            <label>Rotate <input type="number" step="1" value="${f2(el.rot || 0)}" data-k="rot"> °</label>
          </div>
          ${isTx ? `<div class="pf-sec">Text box</div>
            ${row('Vertical align', `<select data-k="valign">${['top', 'middle', 'bottom'].map(v => `<option ${el.valign === v ? 'selected' : ''}>${v}</option>`).join('')}</select>`)}
            ${row('Autofit', `<select data-k="autofit"><option value="">Do not autofit</option><option value="shrink" ${el.autofit === 'shrink' ? 'selected' : ''}>Shrink text on overflow</option><option value="resize" ${el.autofit === 'resize' ? 'selected' : ''}>Resize shape to fit text</option></select>`)}
            ${row('Line spacing', `<input type="number" min="0.5" max="4" step="0.1" value="${el.lh || 1}" data-k="lh"> ×`)}
            ${row('Space after ¶', `<input type="number" min="0" max="72" step="1" value="${el.pspace ?? 0}" data-k="pspace"> pt`)}
            <div class="pf-grid"><label>Left <input type="number" step="0.01" value="${f2(ins[0] / 96)}" data-k="ins0"> in</label><label>Right <input type="number" step="0.01" value="${f2(ins[2] / 96)}" data-k="ins2"> in</label><label>Top <input type="number" step="0.01" value="${f2(ins[1] / 96)}" data-k="ins1"> in</label><label>Bottom <input type="number" step="0.01" value="${f2(ins[3] / 96)}" data-k="ins3"> in</label></div>
            <label class="check"><input type="checkbox" data-k="wrap" ${el.wrap === false ? '' : 'checked'}> Wrap text in shape</label>` : ''}
          <div class="pf-sec">Alt text</div><textarea class="pf-ta" data-k="alt" placeholder="Describe this object for screen readers">${esc(el.alt || '')}</textarea>`;
      },
      wire: (b) => {
        const els = () => E.selected();
        const apply = (fn, commit = true) => { els().forEach(fn); E.renderCanvas(); E.refreshThumb(); if (commit) { SL.commit(); E.updateRibbon(); } };
        $$('input[name=ft]', b).forEach(r => r.addEventListener('change', () => { apply(el => { if (r.value === 'none') el.fill = 'none'; else if (r.value === 'solid') el.fill = typeof el.fill === 'object' ? el.fill.color : (el.fill && el.fill !== 'none' ? el.fill : 'accent1'); else el.fill = { type: 'gradient', color: typeof el.fill === 'string' && el.fill !== 'none' ? el.fill : 'accent1', color2: 'accent2', angle: 90 }; }); E.refreshPane('force'); }));
        b.addEventListener('input', e => {
          const k = e.target.dataset.k; const v = e.target.value;
          if (k === 'transp') { e.target.nextElementSibling.textContent = v + '%'; apply(el => el.opacity = 1 - v / 100, false); }
          if (k === 'radius') apply(el => el.radius = +v, false);
        });
        b.addEventListener('change', e => {
          const k = e.target.dataset.k; if (!k) return; const v = e.target.value; const n = +v;
          const map = {
            strokeWidth: el => { el.strokeWidth = n; if (n > 0 && (!el.stroke || el.stroke === 'none')) el.stroke = 'text'; }, dash: el => el.dash = v, transp: () => { }, radius: () => { }, shadow: el => el.shadow = e.target.checked,
            x: el => el.x = n * 96, y: el => el.y = n * 96, w: el => el.w = Math.max(1, n * 96), h: el => el.h = Math.max(1, n * 96), rot: el => el.rot = ((n % 360) + 360) % 360,
            valign: el => el.valign = v, autofit: el => { if (v) el.autofit = v; else delete el.autofit; el.fontScale = 1; }, lh: el => el.lh = n, pspace: el => el.pspace = n,
            ins0: el => { el.inset = (el.inset || [9.6, 4.8, 9.6, 4.8]).slice(); el.inset[0] = n * 96; }, ins1: el => { el.inset = (el.inset || [9.6, 4.8, 9.6, 4.8]).slice(); el.inset[1] = n * 96; },
            ins2: el => { el.inset = (el.inset || [9.6, 4.8, 9.6, 4.8]).slice(); el.inset[2] = n * 96; }, ins3: el => { el.inset = (el.inset || [9.6, 4.8, 9.6, 4.8]).slice(); el.inset[3] = n * 96; },
            wrap: el => el.wrap = e.target.checked, alt: el => el.alt = v, gangle: el => { if (typeof el.fill === 'object') el.fill.angle = n; },
          };
          if (map[k]) apply(map[k]);
          if (k === 'transp' || k === 'radius') SL.commit();
        });
        b.addEventListener('click', async e => {
          const a = e.target.closest('[data-act]'); if (!a) return;
          const c = await BTS.colorPicker(a, a.dataset.act === 'line-c' ? { none: 'No line' } : { none: 'No fill' }); if (!c) return;
          if (a.dataset.act === 'fill-c') apply(el => el.fill = c);
          if (a.dataset.act === 'fill-g1') apply(el => el.fill.color = c);
          if (a.dataset.act === 'fill-g2') apply(el => el.fill.color2 = c);
          if (a.dataset.act === 'line-c') apply(el => { el.stroke = c; if (c !== 'none' && !el.strokeWidth) el.strokeWidth = 1.5; });
          E.refreshPane('force');
        });
      },
    },
    anim: {
      title: 'Animation Pane',
      html: () => {
        const list = C.animOrder(S());
        if (!list.length) return `<p class="pf-note">No animations on this slide yet. Select an object and choose an effect from the Animations tab.</p>`;
        let step = 0;
        return `<button class="btn" data-act="play" style="margin-bottom:10px">${icon('play', 14)} Play from start</button><div class="anim-list">` + list.map((el, i) => {
          if ((el.anim.trigger || 'click') === 'click') step++;
          const nm = C.ANIMS.find(a => a[0] === el.anim.type)?.[1] || el.anim.type;
          return `<div class="anim-row ${E.sel.includes(el.id) ? 'sel' : ''}" data-id="${el.id}"><span class="anim-step">${(el.anim.trigger || 'click') === 'click' ? step : ''}</span><span class="anim-trig" title="${el.anim.trigger || 'click'}">${icon(el.anim.trigger === 'after' ? 'clock' : el.anim.trigger === 'with' ? 'link' : 'mouse-pointer-click', 13)}</span><span class="anim-name"><b>${esc(nm)}</b> ${esc(C.elName(el))}</span><span class="anim-bar" style="width:${Math.min(60, (el.anim.dur || 500) / 50)}px"></span><button class="rb" data-act="up" title="Move earlier">${icon('chevron-up', 14)}</button><button class="rb" data-act="down" title="Move later">${icon('chevron-down', 14)}</button><button class="rb" data-act="del" title="Remove">${icon('x', 14)}</button></div>`;
        }).join('') + '</div>';
      },
      wire: (b) => b.addEventListener('click', e => {
        const a = e.target.closest('[data-act]'); const r = e.target.closest('.anim-row');
        if (a && a.dataset.act === 'play') { BTS.dispatch('preview-anim'); return; }
        if (!r) return; const id = r.dataset.id;
        if (a) { E.select([id]); if (a.dataset.act === 'del') { delete SL.findEl(id).anim; C.renumberAnims(); SL.commit(); E.renderAll(); } else C.moveAnim(a.dataset.act === 'up' ? -1 : 1); return; }
        E.select([id]);
      }),
    },
    selection: {
      title: 'Selection',
      html: () => `<div class="pf-actions" style="margin-top:0"><button class="btn" data-act="showall">Show all</button><button class="btn" data-act="hideall">Hide all</button></div><div class="sel-list">` + [...S().elements].reverse().map(el => `<div class="sel-row ${E.sel.includes(el.id) ? 'sel' : ''}" data-id="${el.id}">${icon({ text: 'type', shape: 'shapes', image: 'image', table: 'table', chart: 'chart-column', icon: 'smile' }[el.type] || 'square', 14)}<span class="sel-name" title="Double-click to rename">${esc(C.elName(el))}</span><button class="rb" data-act="vis" title="${el.hidden ? 'Show' : 'Hide'}">${icon(el.hidden ? 'eye-off' : 'eye', 14)}</button></div>`).join('') + '</div><p class="pf-note">Top of the list is in front. Ctrl+click to select several.</p>',
      wire: (b) => {
        b.addEventListener('click', e => {
          const a = e.target.closest('[data-act]');
          if (a && a.dataset.act === 'showall') { S().elements.forEach(x => delete x.hidden); SL.commit(); E.renderCanvas(); E.refreshPane('force'); return; }
          if (a && a.dataset.act === 'hideall') { S().elements.forEach(x => x.hidden = true); SL.commit(); E.renderCanvas(); E.refreshPane('force'); return; }
          const r = e.target.closest('.sel-row'); if (!r) return; const el = SL.findEl(r.dataset.id);
          if (a && a.dataset.act === 'vis') { el.hidden = !el.hidden; SL.commit(); E.renderCanvas(); E.refreshThumb(true); E.refreshPane('force'); return; }
          if (e.ctrlKey || e.shiftKey) E.select(E.sel.includes(el.id) ? E.sel.filter(x => x !== el.id) : [...E.sel, el.id]); else E.select([el.id]);
        });
        b.addEventListener('dblclick', async e => { const r = e.target.closest('.sel-row'); if (!r) return; const el = SL.findEl(r.dataset.id); const n = await BTS.prompt('Rename object', 'Name', C.elName(el)); if (n) { el.name = n; SL.commit(); E.refreshPane('force'); } });
      },
    },
    comments: {
      title: 'Comments',
      html: () => {
        const cs = S().comments || [];
        return `<button class="btn primary" data-act="new">${icon('message-square-plus', 15)} New comment</button>` + (cs.length ? cs.map((c, i) => `<div class="cmt"><div class="cmt-h"><b>${esc(c.author)}</b><small>${new Date(c.time).toLocaleString()}</small><button class="rb" data-act="del" data-i="${i}" title="Delete">${icon('trash-2', 14)}</button></div><div class="cmt-t">${esc(c.text)}</div></div>`).join('') : '<p class="pf-note">No comments on this slide.</p>');
      },
      wire: (b) => b.addEventListener('click', e => { const a = e.target.closest('[data-act]'); if (!a) return; if (a.dataset.act === 'new') BTS.dispatch('new-comment'); if (a.dataset.act === 'del') { S().comments.splice(+a.dataset.i, 1); SL.commit(); E.renderAll(); } }),
    },
    find: {
      title: 'Find and Replace',
      html: () => `<div class="field"><label>Find what</label><input id="fr-find" value="${esc(C._find || '')}"></div><div class="field"><label>Replace with</label><input id="fr-rep" value="${esc(C._rep || '')}"></div>
        <label class="check"><input type="checkbox" id="fr-case" ${C._case ? 'checked' : ''}> Match case</label>
        <div class="pf-actions"><button class="btn" data-act="next">Find next</button><button class="btn" data-act="rep">Replace</button><button class="btn primary" data-act="all">Replace all</button></div><div class="pf-note" id="fr-msg"></div>`,
      wire: (b) => {
        const rd = () => { C._find = $('#fr-find', b).value; C._rep = $('#fr-rep', b).value; C._case = $('#fr-case', b).checked; };
        b.addEventListener('input', rd); b.addEventListener('change', rd);
        $('#fr-find', b).addEventListener('keydown', e => { if (e.key === 'Enter') { rd(); C.findNext(C._find); } });
        b.addEventListener('click', e => { const a = e.target.closest('[data-act]'); if (!a) return; rd(); const msg = $('#fr-msg', b); if (a.dataset.act === 'next') msg.textContent = C.findNext(C._find) ? '' : 'No more matches.'; if (a.dataset.act === 'rep') { C.replaceOne(C._find, C._rep); } if (a.dataset.act === 'all') msg.textContent = C.replaceAll(C._find, C._rep) + ' replacement(s) made.'; });
        setTimeout(() => { const f = $('#fr-find', b); if (f && f.isConnected) f.focus(); }, 0);
      },
    },
  };
  C.PANES = PANES;
  C.elName = (el) => (el.type !== 'icon' && el.name) || ({ text: el.ph === 'title' ? 'Title' : el.ph ? 'Content placeholder' : 'Text box', shape: Object.fromEntries(SL.SHAPES)[el.kind] || 'Shape', image: 'Picture', table: 'Table', chart: 'Chart', icon: 'Icon: ' + el.name }[el.type]) + (SL.plainText(el.html || '').trim() ? ': ' + SL.plainText(el.html).trim().slice(0, 24) : '');

  /* ---------- animations helpers ---------- */
  C.animOrder = (slide) => slide.elements.filter(e => e.anim && e.anim.type && e.anim.type !== 'none').sort((a, b) => (a.anim.order ?? 0) - (b.anim.order ?? 0));
  C.renumberAnims = (slide = S()) => C.animOrder(slide).forEach((e, i) => e.anim.order = i);
  C.moveAnim = (d) => {
    const list = C.animOrder(S()); const el = E.primary(); if (!el || !el.anim) return;
    const i = list.indexOf(el), j = i + d; if (j < 0 || j >= list.length) return;
    [list[i], list[j]] = [list[j], list[i]]; list.forEach((e, k) => e.anim.order = k);
    SL.commit(); E.renderAll();
  };

  /* ---------- find / replace ---------- */
  const textFields = (el) => el.type === 'table' ? el.rows.flatMap((r, ri) => r.map((c, ci) => ({ get: () => c.html || '', set: (h) => c.html = h, k: ri + ',' + ci }))) : (el.html != null ? [{ get: () => el.html || '', set: (h) => el.html = h, k: '' }] : []);
  C.findNext = (q) => {
    if (!q) return false; if (E.editing) E.endEdit();
    const cs = C._case; const norm = (s) => cs ? s : s.toLowerCase(); const nq = norm(q);
    const n = P().slides.length; const start = C._lastHit || { s: SL.cur, e: -1 };
    for (let k = 0; k <= n; k++) {
      const si = (start.s + k) % n; const els = P().slides[si].elements;
      for (let ei = (k === 0 ? start.e + 1 : 0); ei < els.length; ei++) {
        if (textFields(els[ei]).some(f => norm(SL.plainText(f.get())).includes(nq))) {
          C._lastHit = { s: si, e: ei }; E.goto(si); E.select([els[ei].id]); return true;
        }
      }
    }
    C._lastHit = null; return false;
  };
  function replaceIn(html, q, r, all) {
    const d = document.createElement('div'); d.innerHTML = html; let count = 0;
    const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), C._case ? 'g' : 'gi');
    const w = document.createTreeWalker(d, NodeFilter.SHOW_TEXT); let nd;
    while ((nd = w.nextNode())) { const v = nd.nodeValue; const nv = v.replace(re, (m) => { if (!all && count) return m; count++; return r; }); if (nv !== v) nd.nodeValue = nv; if (!all && count) break; }
    return { html: d.innerHTML, count };
  }
  C.replaceOne = (q, r) => {
    if (!q) return; const el = E.primary();
    if (el) for (const f of textFields(el)) { const res = replaceIn(f.get(), q, r, false); if (res.count) { f.set(res.html); SL.commit(); E.renderAll(); break; } }
    C.findNext(q);
  };
  C.replaceAll = (q, r) => {
    if (!q) return 0; if (E.editing) E.endEdit(); let n = 0;
    P().slides.forEach(s => s.elements.forEach(el => textFields(el).forEach(f => { const res = replaceIn(f.get(), q, r, true); if (res.count) { f.set(res.html); n += res.count; } })));
    if (n) { SL.commit(); E.renderAll(); }
    return n;
  };

  /* ---------- format painter ---------- */
  const PAINT_KEYS = ['fill', 'stroke', 'strokeWidth', 'dash', 'shadow', 'opacity', 'font', 'size', 'color', 'bold', 'italic', 'underline', 'strike', 'align', 'valign', 'effect', 'lh', 'radius', 'pspace'];
  C.startPainter = () => {
    const el = E.primary(); if (!el) return BTS.toast('Select an object whose formatting you want to copy');
    E.painter = Object.fromEntries(PAINT_KEYS.filter(k => el[k] !== undefined).map(k => [k, SL.clone(el[k])]));
    $('#stage').classList.add('painting'); $('[data-cmd="format-painter"]').classList.add('on');
    BTS.toast('Click an object to paste the formatting', { ms: 1800 });
  };
  E.applyPainter = (id) => {
    const t = SL.findEl(id); const f = E.painter; E.painter = null; $('#stage').classList.remove('painting'); $('[data-cmd="format-painter"]').classList.remove('on');
    if (!t || !f) return; Object.assign(t, SL.clone(f));
    if (t.html) t.html = stripHTML(t.html, ['color', 'font-size', 'font-family', 'font-weight', 'font-style'], ['b', 'strong', 'i', 'em']);
    E.sel = [id]; E.renderCanvas(); SL.commit(); E.refreshThumb(true); E.updateRibbon();
  };

  /* ---------- reuse slides from another .pptx (the dialog with thumbnails lives in showtools.js) ---------- */
  C.reuseSlides = () => window.SLX && SLX.reuseSlides ? SLX.reuseSlides() : C.reuseSlidesQuick();
  C.reuseSlidesQuick = async () => {
    const f = await BTS.pickFile('.pptx'); if (!f) return;
    try {
      const other = await window.PPTXImport.parse(f.data);
      const W = SL.W(P()) / SL.W(other), H = SL.H(P()) / SL.H(other);
      const add = other.slides.map(s => { s.id = SL.uid('s'); s.elements.forEach(el => { el.x *= W; el.y *= H; el.w *= W; el.h *= H; }); return s; });
      P().slides.splice(SL.cur + 1, 0, ...add); SL.cur += 1; E.railSel = new Set(add.map((_, i) => SL.cur + i));
      SL.commit(); E.renderAll(); BTS.toast(`Inserted ${add.length} slide(s)`);
    } catch (err) { console.error(err); BTS.toast('Couldn’t read that presentation: ' + err.message, { error: true }); }
  };

  /* ====================================================================== */
  /*  Context menus                                                          */
  /* ====================================================================== */
  E.canvasContext = (ev) => {
    const elNode = ev.target.closest('.sl-slide.editor > .el:not(.el-footer)');
    const at = { x: ev.clientX, y: ev.clientY };
    if (elNode && !E.editing) { const id = elNode.dataset.id; if (!E.sel.includes(id)) E.select(E.groupIds(SL.findEl(id))); }
    const els = E.selected(); const el = els[0];
    if (E.editing) {
      return BTS.menu(at, [{ label: 'Cut', icon: 'scissors', kbd: 'Ctrl+X', action: () => document.execCommand('cut') }, { label: 'Copy', icon: 'copy', kbd: 'Ctrl+C', action: () => document.execCommand('copy') }, { label: 'Paste', icon: 'clipboard-paste', kbd: 'Ctrl+V', action: () => BTS.dispatch('paste') }, { sep: 1 },
        { label: 'Bold', icon: 'bold', action: () => C.toggleStyle('bold') }, { label: 'Italic', icon: 'italic', action: () => C.toggleStyle('italic') }, { label: 'Bullets', icon: 'list', action: () => C.setList('ul') }, { label: 'Numbering', icon: 'list-ordered', action: () => C.setList('ol') }, { sep: 1 },
        { label: 'Font size…', icon: 'a-arrow-up', action: async () => { const v = await BTS.prompt('Font size', 'Size in points', sizeAtCaret() || 18); if (v) C.setFontSize(+v); } }, { label: 'Exit editing', icon: 'check', kbd: 'Esc', action: () => E.endEdit() }]);
    }
    if (el) {
      const items = [
        { label: 'Cut', icon: 'scissors', kbd: 'Ctrl+X', action: () => E.copy(true) }, { label: 'Copy', icon: 'copy', kbd: 'Ctrl+C', action: () => E.copy() }, { label: 'Paste', icon: 'clipboard-paste', kbd: 'Ctrl+V', action: () => BTS.dispatch('paste'), disabled: !E.clip },
        { label: 'Duplicate', icon: 'copy-plus', kbd: 'Ctrl+D', action: () => E.duplicateSelected() }, { label: 'Delete', icon: 'trash-2', kbd: 'Del', action: () => E.deleteSelected() }, { sep: 1 },
      ];
      if (isTextual(el) && els.length === 1) items.push({ label: el.type === 'table' ? 'Edit table' : 'Edit text', icon: 'text-cursor-input', action: () => E.startEdit(el.id) });
      if (el.type === 'image' && el.src) items.push({ label: 'Crop', icon: 'crop', action: () => E.startCrop(el.id) }, { label: 'Change picture…', icon: 'image-plus', action: () => BTS.dispatch('pic-replace') }, { label: 'Save picture as…', icon: 'image-down', action: () => C.savePicture(el) });
      if (el.type === 'chart') items.push({ label: 'Edit data…', icon: 'sheet', action: () => BTS.dispatch('chart-edit') });
      items.push({ label: 'Bring to front', icon: 'bring-to-front', action: () => E.zorder('front') }, { label: 'Bring forward', icon: 'arrow-up-from-line', action: () => E.zorder('forward') }, { label: 'Send backward', icon: 'arrow-down-from-line', action: () => E.zorder('backward') }, { label: 'Send to back', icon: 'send-to-back', action: () => E.zorder('back') }, { sep: 1 });
      if (els.length > 1) items.push({ label: 'Group', icon: 'group', kbd: 'Ctrl+G', action: () => E.group() });
      if (els.some(e => e.g)) items.push({ label: 'Ungroup', icon: 'ungroup', kbd: 'Ctrl+Shift+G', action: () => E.ungroup() });
      items.push({ label: 'Align & distribute…', icon: 'align-start-vertical', action: () => C.alignMenu(at) }, { label: 'Animate…', icon: 'star', action: () => { BTS.showTab('animations'); } }, { label: 'Format shape…', icon: 'paint-bucket', action: () => E.openPane('shape') });
      return BTS.menu(at, items);
    }
    BTS.menu(at, [
      { label: 'Paste', icon: 'clipboard-paste', kbd: 'Ctrl+V', action: () => BTS.dispatch('paste') }, { sep: 1 },
      { label: 'New slide', icon: 'square-plus', kbd: 'Ctrl+M', action: () => E.newSlide() }, { label: 'Layout…', icon: 'layout-template', action: () => C.layoutPicker(at, (l) => { E.applyLayout(l); SL.commit(); E.renderAll(); }, S().layout) },
      { label: 'Reset slide', icon: 'rotate-ccw', action: () => { E.resetSlide(); SL.commit(); E.renderAll(); } }, { label: 'Format background…', icon: 'paint-roller', action: () => E.openPane('bg') }, { sep: 1 },
      { label: 'Select all', icon: 'box-select', kbd: 'Ctrl+A', action: () => E.selectAll() },
      { label: 'Gridlines', checked: E.grid, action: () => BTS.dispatch('toggle-grid') }, { label: 'Guides', checked: E.guides, action: () => BTS.dispatch('toggle-guides') }, { label: 'Ruler', checked: E.ruler, action: () => BTS.dispatch('toggle-ruler') },
    ]);
  };
  E.railContext = (ev) => {
    const at = { x: ev.clientX, y: ev.clientY }; const n = E.railSel.size;
    const onItem = !!ev.target.closest('.rail-item, .sorter-item');
    BTS.menu(at, [
      { label: 'New slide', icon: 'square-plus', kbd: 'Ctrl+M', action: () => E.newSlide() },
      ...(onItem ? [
        { label: n > 1 ? `Duplicate ${n} slides` : 'Duplicate slide', icon: 'copy-plus', action: () => E.duplicateSlides() },
        { label: n > 1 ? `Delete ${n} slides` : 'Delete slide', icon: 'trash-2', kbd: 'Del', action: () => E.deleteSlides() },
        { label: 'Copy', icon: 'copy', kbd: 'Ctrl+C', action: () => C.copySlides() }, { label: 'Cut', icon: 'scissors', kbd: 'Ctrl+X', action: () => { C.copySlides(); E.deleteSlides(); } },
        { label: 'Paste slides', icon: 'clipboard-paste', action: () => C.pasteSlides(), disabled: !E.slideClip }, { sep: 1 },
        { label: 'Add section', icon: 'rows-3', action: () => C.addSection() },
        { label: 'Layout…', icon: 'layout-template', action: () => C.layoutPicker(at, (l) => { [...E.railSel].forEach(i => E.applyLayout(l, P().slides[i])); SL.commit(); E.renderAll(); }, S().layout) },
        { label: 'Reset slide', icon: 'rotate-ccw', action: () => { [...E.railSel].forEach(i => E.resetSlide(P().slides[i])); SL.commit(); E.renderAll(); } },
        { label: 'Format background…', icon: 'paint-roller', action: () => E.openPane('bg') },
        { label: S().hidden ? 'Unhide slide' : 'Hide slide', icon: 'eye-off', action: () => BTS.dispatch('hide-slide') }, { sep: 1 },
        { label: 'Move up', icon: 'chevron-up', kbd: 'Ctrl+↑', action: () => E.moveSlideBy(-1) }, { label: 'Move down', icon: 'chevron-down', kbd: 'Ctrl+↓', action: () => E.moveSlideBy(1) },
      ] : [{ label: 'Paste slides', icon: 'clipboard-paste', action: () => C.pasteSlides(), disabled: !E.slideClip }]),
    ]);
  };
  C.copySlides = () => { E.slideClip = { slides: [...E.railSel].sort((a, b) => a - b).map(i => SL.clone(P().slides[i])), assets: Object.assign({}, SL.assets) }; BTS.toast(`Copied ${E.slideClip.slides.length} slide(s)`); };
  C.pasteSlides = () => {
    if (!E.slideClip || (E.masterGuard && E.masterGuard())) return; Object.assign(SL.assets, E.slideClip.assets);
    const add = E.slideClip.slides.map(s => { const c = SL.clone(s); c.id = SL.uid('s'); c.elements.forEach(SL.reId); E.regroup(c.elements); return c; });
    const at = Math.max(...E.railSel, SL.cur) + 1; P().slides.splice(at, 0, ...add); SL.cur = at; E.railSel = new Set(add.map((_, i) => at + i)); SL.commit(); E.renderAll();
  };
  C.addSection = async () => {
    const name = await BTS.prompt('Add section', 'Section name', 'Untitled Section'); if (!name) return;
    P().sections = P().sections || []; const id = S().id;
    const ex = P().sections.find(s => s.start === id); if (ex) ex.name = name; else P().sections.push({ start: id, name });
    if (SL.cur > 0 && !P().sections.some(s => s.start === P().slides[0].id)) P().sections.push({ start: P().slides[0].id, name: 'Default Section' });
    SL.commit(); E.renderAll();
  };
  C.savePicture = async (el) => {
    const url = SL.assetURL(el.src); const m = /^data:image\/(\w+)/.exec(url); const ext = (m && m[1] === 'jpeg') ? 'jpg' : (m ? m[1].replace('svg+xml', 'svg') : 'png');
    const data = BTS.dataURLToBytes(url);
    if (BTS.desktop) { const p = await bts.saveDialog('Picture.' + ext, [{ name: 'Image', extensions: [ext] }]); if (p) { await bts.writeFile(p, data); BTS.toast('Picture saved'); } }
    else BTS.download('Picture.' + ext, data);
  };
  C.alignMenu = (anchor) => {
    const multi = E.sel.length > 1; const rel = C._alignToSlide || !multi;
    BTS.menu(anchor, [
      { label: 'Align Left', icon: 'align-start-vertical', action: () => E.align('left', rel) }, { label: 'Align Center', icon: 'align-center-vertical', action: () => E.align('center', rel) }, { label: 'Align Right', icon: 'align-end-vertical', action: () => E.align('right', rel) }, { sep: 1 },
      { label: 'Align Top', icon: 'align-start-horizontal', action: () => E.align('top', rel) }, { label: 'Align Middle', icon: 'align-center-horizontal', action: () => E.align('middle', rel) }, { label: 'Align Bottom', icon: 'align-end-horizontal', action: () => E.align('bottom', rel) }, { sep: 1 },
      { label: 'Distribute Horizontally', icon: 'align-horizontal-space-around', action: () => E.align('dist-h', rel) }, { label: 'Distribute Vertically', icon: 'align-vertical-space-around', action: () => E.align('dist-v', rel) }, { sep: 1 },
      { label: 'Align to Slide', checked: rel, action: () => { C._alignToSlide = true; } }, { label: 'Align Selected Objects', checked: !rel, disabled: !multi, action: () => { C._alignToSlide = false; } },
    ]);
  };
  C.arrangeMenu = (anchor) => {
    BTS.menu(anchor, [
      { head: 'Order objects' }, { label: 'Bring to Front', icon: 'bring-to-front', kbd: 'Ctrl+Shift+]', action: () => E.zorder('front') }, { label: 'Send to Back', icon: 'send-to-back', kbd: 'Ctrl+Shift+[', action: () => E.zorder('back') }, { label: 'Bring Forward', icon: 'arrow-up-from-line', kbd: 'Ctrl+]', action: () => E.zorder('forward') }, { label: 'Send Backward', icon: 'arrow-down-from-line', kbd: 'Ctrl+[', action: () => E.zorder('backward') },
      { head: 'Group objects' }, { label: 'Group', icon: 'group', kbd: 'Ctrl+G', action: () => E.group(), disabled: E.sel.length < 2 }, { label: 'Ungroup', icon: 'ungroup', kbd: 'Ctrl+Shift+G', action: () => E.ungroup(), disabled: !E.selected().some(e => e.g) },
      { head: 'Position objects' }, { label: 'Align…', icon: 'align-start-vertical', action: () => setTimeout(() => C.alignMenu(anchor), 0) }, { label: 'Rotate…', icon: 'rotate-cw', action: () => setTimeout(() => C.rotateMenu(anchor), 0) }, { label: 'Selection Pane…', icon: 'layers-2', action: () => E.openPane('selection') },
    ]);
  };
  C.rotateMenu = (anchor) => BTS.menu(anchor, [
    { label: 'Rotate Right 90°', icon: 'rotate-cw', action: () => rot(90) }, { label: 'Rotate Left 90°', icon: 'rotate-ccw', action: () => rot(-90) },
    { label: 'Flip Vertical', icon: 'flip-vertical-2', action: () => flip('flipV') }, { label: 'Flip Horizontal', icon: 'flip-horizontal-2', action: () => flip('flipH') },
    { label: 'Reset rotation', icon: 'rotate-ccw', action: () => { E.selected().forEach(e => e.rot = 0); SL.commit(); E.renderCanvas(); E.refreshThumb(true); } },
  ]);
  function rot(d) { E.selected().forEach(e => e.rot = (((e.rot || 0) + d) % 360 + 360) % 360); SL.commit(); E.renderCanvas(); E.refreshThumb(true); }
  function flip(k) { E.selected().forEach(e => e[k] = !e[k]); SL.commit(); E.renderCanvas(); E.refreshThumb(true); }
})();
