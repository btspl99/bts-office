/* BTS Docs 1.3 — Hindi & Marathi: phonetic typing (common/indic.js) wired into Docs' undo + track changes,
 * and the old-font converter (common/legacy-fonts.js): convert Kruti Dev / DevLys / Shivaji text to Unicode
 * (selection or whole document), Unicode → Kruti Dev 010, and an offer to convert when such a file is opened. */
(function () {
  const D = window.Docs; const { $, $$ } = D;
  const IN = () => window.BTSIndic, LF = () => window.BTSLegacyFonts;
  const UNI = 'Noto Sans Devanagari';

  /* ================= typing ================= */
  /** Replace the del characters before the caret with text — through Docs' history and track changes. */
  function docsReplace(el, del, text, newWord) {
    if (!D.isRoot(el)) return false;
    if (D.isProtected && D.isProtected()) return true; // read-only: swallow the key
    if (newWord) D.checkpoint();
    const tracking = D.tracking() && el === D.doc;
    const b = del ? IN().rangeBefore(el, del) : null;
    const s = getSelection();
    if (tracking) {
      const r = b && b.range;
      if (r) {
        const n = r.startContainer; const ins = D.closestTC(n, 'ins');
        if (n === r.endContainer && n.nodeType === 3 && ins && ins.dataset.author === D.tcAuthor()) {
          n.replaceData(r.startOffset, r.endOffset - r.startOffset, text);
          if (n.length) D.setCaret(n, r.startOffset + text.length); else { const p = n.parentNode, i = [...p.childNodes].indexOf(n); n.remove(); if (ins.isConnected && !ins.textContent) { const pp = ins.parentNode, j = [...pp.childNodes].indexOf(ins); ins.remove(); D.setCaret(pp, j); } else D.setCaret(p, i); }
        } else {
          const res = D.trackDeleteRange(r);
          if (res.last && res.last.isConnected) { const p = res.last.parentNode; D.setCaret(p, [...p.childNodes].indexOf(res.last) + 1); } else { r.collapse(true); s.removeAllRanges(); s.addRange(r); }
          if (text) D.trackInsertText(text);
        }
      } else {
        const rr = D.range(); if (rr && !rr.collapsed) { const res = D.trackDeleteRange(rr); if (res.last && res.last.isConnected) { const p = res.last.parentNode; D.setCaret(p, [...p.childNodes].indexOf(res.last) + 1); } }
        if (text) D.trackInsertText(text);
      }
    } else {
      if (b && b.range) { s.removeAllRanges(); s.addRange(b.range); }
      if (text) document.execCommand('insertText', false, text); else if (del) document.execCommand('delete', false);
    }
    D.afterChange();
    return true;
  }
  const LANGS = { hi: 'Hindi (हिंदी)', mr: 'Marathi (मराठी)' };
  function syncTypingUI() {
    const lang = IN() && IN().lang; const b = $('#btn-indic');
    if (b) { b.classList.toggle('on', !!lang); const l = b.querySelector('span'); if (l) l.firstChild.textContent = lang ? (lang === 'hi' ? 'हिंदी ' : 'मराठी ') : 'Type हिंदी '; b.title = lang ? `Typing in ${LANGS[lang]} — click to turn off (Ctrl+Alt+${lang === 'hi' ? 'H' : 'M'})` : 'Type in Hindi or Marathi using English letters (Ctrl+Alt+H / Ctrl+Alt+M)'; }
    BTS.setStatus && D.updateStatus && D.updateStatus();
  }
  D.indicToggle = () => { if (!IN()) return; IN().toggle(); D.restoreFocus(); };
  D.indicMenu = (anchor) => { if (!IN()) return; BTS.menu(anchor, IN().menuItems().concat([{ sep: true }, { label: 'Convert old fonts to Unicode', icon: 'languages', action: () => D.convertLegacy() }, { label: 'Convert selection to Kruti Dev', icon: 'type', action: () => D.convertToKruti() }, { label: 'Font converter…', icon: 'arrow-left-right', action: () => LF().dialog() }])); };
  /** The status bar pill (placed next to the proofing language). */
  D.mountIndicPill = () => {
    if (!IN()) return;
    BTS.setStatus('indic', ''); const slot = document.querySelector('#statusbar [data-sb="indic"]');
    if (slot && !slot.contains(IN().indicatorEl())) slot.appendChild(IN().indicatorEl());
  };

  /* ================= old fonts ================= */
  const legacyOf = (el) => LF().detect(getComputedStyle(el).fontFamily);
  /** Text nodes (inside the editable roots) whose font is a legacy font; scope = Range or null for whole doc. */
  function legacyNodes(scope) {
    const out = []; const roots = scope ? [D.rootOf(scope.commonAncestorContainer) || D.doc] : [D.doc, D.hdr, D.ftr];
    roots.forEach(root => {
      const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT); let n;
      while ((n = w.nextNode())) {
        if (!n.data.replace(/[\s​]/g, '')) continue;
        if (scope && !scope.intersectsNode(n)) continue;
        if (n.parentElement.closest('[contenteditable="false"], del.tc-del')) continue;
        const f = legacyOf(n.parentElement); if (f) out.push({ node: n, font: f });
      }
    });
    return out;
  }
  /** Element that gives a text node its (legacy) font inline, or null when it comes from a style. */
  function fontHolder(node, root) { for (let e = node.parentElement; e && e !== root; e = e.parentElement) { if (e.style && e.style.fontFamily) return e; if (e.tagName === 'FONT' && e.getAttribute('face')) return e; } return null; }
  /** Convert a list of text nodes; consecutive nodes of one word are joined so ि and reph reorder correctly. */
  function convertNodes(list, forceFont) {
    let words = 0; const stack = D.fontStack(UNI); const touched = new Set();
    for (let i = 0; i < list.length; i++) {
      let { node, font } = list[i]; font = forceFont || font;
      // pull in following text nodes when this one ends inside a syllable (ि typed before its consonant, or a mark follows)
      while (i + 1 < list.length && list[i + 1].node.parentNode && D.blockOf(list[i + 1].node) === D.blockOf(node) && (/[fÇÆÉi]$/.test(node.data) || /^[kZ±ÊhsSaqw`W¡%‚ª+zÓî\-]/.test(list[i + 1].node.data)) && nextTextNode(node) === list[i + 1].node) {
        const nx = list[++i].node; node.data += nx.data; nx.data = '';
      }
      const u = LF().toUnicode(node.data, font); words += D.countWords(u); node.data = u;
      const root = D.rootOf(node) || D.doc; const h = fontHolder(node, root);
      if (h && LF().detect(h.style.fontFamily || h.getAttribute('face'))) { if (!touched.has(h)) { touched.add(h); if (h.tagName === 'FONT') h.removeAttribute('face'); h.style.fontFamily = stack; } }
      else if (legacyOf(node.parentElement)) { const sp = document.createElement('span'); sp.style.fontFamily = stack; node.parentNode.insertBefore(sp, node); sp.appendChild(node); }
    }
    return words;
  }
  function nextTextNode(n) { const root = D.rootOf(n) || D.doc; const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT); w.currentNode = n; let x; while ((x = w.nextNode())) { if (x.data.length) return x; } return null; }
  /** Styles / theme fonts that are legacy fonts → Unicode font (whole-document conversion). */
  function fixStyleFonts() {
    let changed = false; const des = D.design;
    if (LF().detect(D.bodyFont())) { des.bodyFont = UNI; changed = true; }
    if (LF().detect(D.headFont())) { des.headFont = UNI; changed = true; }
    Object.values(des.styleMods || {}).forEach(m => { if (m && m.font && LF().detect(m.font)) { m.font = UNI; changed = true; } });
    (des.customStyles || []).forEach(u => { if (u.def && u.def.font && LF().detect(u.def.font)) { u.def.font = UNI; changed = true; } });
    if (changed) D.applySettings();
    return changed;
  }
  const FONT_NAME = { krutidev: 'Kruti Dev', shivaji: 'Shivaji' };
  D.convertLegacy = async () => {
    if (!LF()) return;
    const r = D.range() || D.lastRange();
    const sel = r && !r.collapsed && D.rootOf(r.commonAncestorContainer) ? r.cloneRange() : null;
    if (sel) {
      let list = legacyNodes(sel); let force = null;
      if (!list.length) {
        const pick = await BTS.dialog({ title: 'Convert old fonts', width: 460, body: '<p style="margin:0 0 6px">The selected text is not in a Kruti Dev, DevLys or Shivaji font.</p><p style="margin:0;color:var(--ink-2)">If it was typed in one of them (it looks like <b>Hkkjr</b> instead of <b>भारत</b>), choose the font to convert from:</p>', buttons: [{ label: 'Cancel', value: null }, { label: 'Shivaji', value: 'shivaji' }, { label: 'Kruti Dev / DevLys', value: 'krutidev', primary: true }] });
        if (!pick) return; force = pick;
        const rr = sel; list = D.textNodesIn(rr).map(s => s).filter(s => s.node.data.trim()).map(s => ({ node: s.node, font: pick, start: s.start, end: s.end }));
        if (!list.length) return;
        D.change(() => {
          // split partly selected nodes so only the selection is converted
          const segs = list.map(({ node, start, end }) => { let n = node; if (start > 0) n = n.splitText(start); if (end - start < n.length) n.splitText(end - start); return { node: n, font: pick }; });
          const words = convertNodes(segs, force); BTS.toast(`Converted ${words.toLocaleString('en-IN')} word${words === 1 ? '' : 's'} from ${FONT_NAME[pick]} to Unicode`);
        });
        return;
      }
      D.change(() => { const words = convertNodes(list); BTS.toast(`Converted ${words.toLocaleString('en-IN')} word${words === 1 ? '' : 's'} to Unicode (${UNI})`); });
      return;
    }
    const list = legacyNodes(null);
    const styleLegacy = LF().detect(D.bodyFont()) || LF().detect(D.headFont());
    if (!list.length && !styleLegacy) { BTS.toast('No Kruti Dev, DevLys or Shivaji text found. To convert text anyway, select it first.'); return; }
    D.change(() => {
      const fonts = [...new Set(list.map(x => x.font))];
      const words = convertNodes(list); fixStyleFonts(); D.hideLegacyBar();
      BTS.toast(`Converted ${words.toLocaleString('en-IN')} word${words === 1 ? '' : 's'} from ${fonts.map(f => FONT_NAME[f]).join(' and ') || 'old fonts'} to Unicode (${UNI})`);
    });
  };
  D.convertToKruti = () => {
    if (!LF()) return;
    const r = D.range() || D.lastRange();
    if (!r || r.collapsed || !D.rootOf(r.commonAncestorContainer)) { BTS.toast('Select the Hindi text you want in Kruti Dev 010 first.'); return; }
    const s = getSelection(); s.removeAllRanges(); s.addRange(r);
    let n = 0;
    D.change(() => {
      const spans = D.wrapText(D.range(), () => { const sp = document.createElement('span'); sp.style.fontFamily = '"Kruti Dev 010"'; return sp; });
      spans.forEach(sp => { const t = sp.textContent; if (!D.hasDevanagari(t)) { sp.replaceWith(...sp.childNodes); return; } sp.textContent = LF().fromUnicode(t, 'krutidev'); n++; });
    });
    if (n) BTS.toast(window.btsFontAvailable && btsFontAvailable('Kruti Dev 010') ? 'Converted to Kruti Dev 010' : 'Converted to Kruti Dev 010 — install the Kruti Dev 010 font to see it as Hindi', { ms: 4200 });
    else BTS.toast('The selection has no Hindi (Devanagari) text to convert.');
  };
  /* ---------- offer after opening a file ---------- */
  D.hideLegacyBar = () => { const b = $('#legacy-bar'); if (b) b.classList.add('hidden'); };
  D.checkLegacyFonts = () => {
    if (!LF()) return;
    const list = legacyNodes(null); const style = LF().detect(D.bodyFont());
    if (!list.length && !style) { D.hideLegacyBar(); return; }
    const fonts = [...new Set(list.map(x => x.font).concat(style ? [style] : []))].map(f => FONT_NAME[f]).join(' and ');
    let bar = $('#legacy-bar'); if (!bar) { bar = document.createElement('div'); bar.id = 'legacy-bar'; $('#canvas').insertBefore(bar, $('#zoomer')); }
    bar.className = '';
    bar.innerHTML = `${icon('languages', 16)}<span>This document uses the old <b>${fonts}</b> font${/ and /.test(fonts) ? 's' : ''}. Convert it to Unicode so it shows correctly everywhere and can be searched and spell-checked.</span><span style="flex:1"></span><button class="btn primary" data-cmd="legacy-convert">Convert to Unicode</button><button class="btn" data-cmd="legacy-bar-close" title="Keep the old font">Not now</button>`;
  };

  D.initIndian = () => {
    if (IN()) {
      IN().setHandler(docsReplace);
      IN().onChange(syncTypingUI);
      syncTypingUI();
    }
  };
})();
