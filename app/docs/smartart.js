/* BTS Docs 1.3 — Insert ▸ SmartArt. The shared module (common/smartart.js) lays the diagram out; Docs draws it
 * into a PNG picture (2× for sharp printing) with the document's body font and theme colours, so it survives
 * .docx / .odt / PDF export like any picture. The layout id and the items are kept on the <img> as
 * data-smartart='{"id":…,"items":[…]}' — double-click it (or right-click ▸ Edit SmartArt) to change it. */
(function () {
  const D = window.Docs; const { PX } = D;
  const SA = () => window.BTSSmartArt;
  const SCALE = 2;

  /** A bundled font from the body font's stack (pictures can't use fonts that aren't really there). */
  function fontName() {
    const fams = D.fontStack(D.bodyFont()).split(',').map(s => s.trim().replace(/^["']|["']$/g, ''));
    const bundled = new Set((window.BTS_FONTS || []).filter(f => f.bundled).map(f => f.name));
    return fams.find(f => bundled.has(f)) || 'Carlito';
  }
  /** Theme colours: accent first, then the theme's other colours, dark colour last. */
  D.smartArtPalette = () => { const t = D.theme(); return [...new Set([D.accent(), ...(t.colors || []), D.darkColor()].filter(Boolean).map(c => String(c).toUpperCase()))]; };

  /** Draw the SVG made by BTSSmartArt.toSVG onto a canvas (uses the page's loaded fonts, unlike an SVG <img>). */
  function drawSVG(svg, W, H) {
    const doc = new DOMParser().parseFromString(svg, 'image/svg+xml'); const root = doc.documentElement;
    const c = document.createElement('canvas'); c.width = Math.round(W * SCALE); c.height = Math.round(H * SCALE);
    const g = c.getContext('2d'); g.scale(SCALE, SCALE); g.lineJoin = 'round';
    const num = (e, a, d = 0) => { const v = parseFloat(e.getAttribute(a)); return isNaN(v) ? d : v; };
    const rot = (e) => { const m = /rotate\(([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\)/.exec(e.getAttribute('transform') || ''); if (m) { g.translate(+m[2], +m[3]); g.rotate(+m[1] * Math.PI / 180); g.translate(-m[2], -m[3]); } };
    for (const e of root.children) {
      g.save(); g.globalAlpha = num(e, 'opacity', 1); rot(e);
      const tag = e.tagName.toLowerCase();
      if (tag === 'rect') { g.fillStyle = e.getAttribute('fill') || '#fff'; g.fillRect(0, 0, num(e, 'width'), num(e, 'height')); }
      else if (tag === 'path') {
        const p = new Path2D(e.getAttribute('d') || ''); const fill = e.getAttribute('fill'), stroke = e.getAttribute('stroke');
        if (fill && fill !== 'none') { g.fillStyle = fill; g.fill(p); }
        if (stroke && stroke !== 'none') { g.strokeStyle = stroke; g.lineWidth = num(e, 'stroke-width', 1); g.stroke(p); }
      } else if (tag === 'line') {
        const x1 = num(e, 'x1'), y1 = num(e, 'y1'), x2 = num(e, 'x2'), y2 = num(e, 'y2'), sw = num(e, 'stroke-width', 1), col = e.getAttribute('stroke') || '#888';
        const arrow = !!e.getAttribute('marker-end'); const a = Math.atan2(y2 - y1, x2 - x1), L = 5 * sw;
        const ex = arrow ? x2 - Math.cos(a) * L * 0.6 : x2, ey = arrow ? y2 - Math.sin(a) * L * 0.6 : y2;
        g.strokeStyle = col; g.lineWidth = sw; g.beginPath(); g.moveTo(x1, y1); g.lineTo(ex, ey); g.stroke();
        if (arrow) { const tx = x2 + Math.cos(a) * L * 0.2, ty = y2 + Math.sin(a) * L * 0.2; g.fillStyle = col; g.beginPath(); g.moveTo(tx, ty); g.lineTo(tx - Math.cos(a - 0.46) * L, ty - Math.sin(a - 0.46) * L); g.lineTo(tx - Math.cos(a + 0.46) * L, ty - Math.sin(a + 0.46) * L); g.closePath(); g.fill(); }
      } else if (tag === 'text') {
        g.font = `${e.getAttribute('font-weight') === '700' ? '700' : '400'} ${num(e, 'font-size', 12)}px ${e.getAttribute('font-family') || 'sans-serif'}`;
        g.fillStyle = e.getAttribute('fill') || '#000'; g.textAlign = ({ start: 'left', end: 'right', middle: 'center' })[e.getAttribute('text-anchor')] || 'left'; g.textBaseline = 'alphabetic';
        for (const t of e.children) g.fillText(t.textContent, num(t, 'x'), num(t, 'y'));
      }
      g.restore();
    }
    return c.toDataURL('image/png');
  }

  /** {id, items} → { url, w, h, alt } */
  async function render(spec, w) {
    const font = fontName();
    try { await Promise.all([document.fonts.load(`400 16px "${font}"`), document.fonts.load(`700 16px "${font}"`)]); } catch { }
    const H = Math.round(w * 0.58);
    const model = SA().layout(spec.id, spec.items, { w, h: H, palette: D.smartArtPalette(), font, textColor: '#1f2533' });
    const url = drawSVG(SA().toSVG(model, { font }), model.w, model.h);
    const name = (SA().LAYOUTS.find(l => l.id === spec.id) || {}).name || 'SmartArt';
    const alt = `${name} diagram: ` + spec.items.filter(it => !it.level).map(it => it.text).join(', ');
    return { url, w: model.w, h: model.h, alt: alt.length > 240 ? alt.slice(0, 237) + '…' : alt };
  }
  const specOf = (img) => { try { const s = JSON.parse(img.dataset.smartart || ''); return s && s.id && Array.isArray(s.items) ? s : null; } catch { return null; } };
  const dlg = (initial) => SA().dialog({ palette: D.smartArtPalette(), font: fontName(), textColor: '#1f2533', initial });

  D.insertSmartArt = async () => {
    if (!SA()) return;
    const keep = D.range() || D.lastRange(); const keepR = keep && keep.cloneRange();
    const spec = await dlg(null); if (!spec || !spec.items.length) { D.restoreFocus(); return; }
    const maxW = (D.paperSize().w - D.page.m.l - D.page.m.r) * PX;
    const r = await render(spec, Math.round(Math.min(maxW, 600)));
    if (keepR) { const s = getSelection(); s.removeAllRanges(); s.addRange(keepR); }
    D.change(() => {
      const img = document.createElement('img'); img.src = r.url; img.alt = r.alt; img.className = 'smartart';
      img.style.width = Math.round(r.w) + 'px'; img.style.height = Math.round(r.h) + 'px';
      img.dataset.nw = Math.round(r.w); img.dataset.nh = Math.round(r.h); img.dataset.smartart = JSON.stringify(spec);
      const rg = D.ensureSel(); rg.deleteContents();
      // a diagram gets its own centred paragraph
      const blk = D.blockOf(rg.startContainer); let host = null;
      if (blk && blk.textContent.trim()) { host = document.createElement('p'); host.style.textAlign = 'center'; host.appendChild(img); blk.after(host); }
      else if (blk && blk.tagName === 'P') { blk.innerHTML = ''; blk.style.textAlign = 'center'; blk.appendChild(img); host = blk; }
      else rg.insertNode(img);
      if (host && !host.nextElementSibling) { const p = document.createElement('p'); p.innerHTML = '<br>'; host.after(p); }
      const after = document.createRange(); after.setStartAfter(img); after.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(after);
      setTimeout(() => D.selectImage && D.selectImage(img), 30);
    });
  };
  D.editSmartArt = async (img) => {
    img = img || D.selImg; const spec = img && specOf(img); if (!spec) return;
    if (D.isProtected && D.isProtected()) { BTS.toast('This document is protected — stop protection to change it.'); return; }
    const ns = await dlg(spec); if (!ns || !ns.items.length) return;
    const curW = parseFloat(img.style.width) || img.offsetWidth || 600;
    const r = await render(ns, Math.round(+img.dataset.nw || 600));
    D.change(() => {
      img.src = r.url; img.alt = r.alt; img.dataset.smartart = JSON.stringify(ns); img.dataset.nw = Math.round(r.w); img.dataset.nh = Math.round(r.h);
      img.style.width = Math.round(curW) + 'px'; img.style.height = Math.round(curW * r.h / r.w) + 'px';
    });
    setTimeout(() => { D.selectImage && D.selectImage(img); D.updateImgHandles && D.updateImgHandles(); }, 30);
  };
  D.isSmartArt = (img) => !!(img && img.tagName === 'IMG' && specOf(img));
  D.initSmartArt = () => {
    [D.doc, D.hdr, D.ftr].forEach(el => el.addEventListener('dblclick', e => { const img = e.target.closest && e.target.closest('img[data-smartart]'); if (img) { e.preventDefault(); e.stopPropagation(); D.editSmartArt(img); } }));
  };
})();
