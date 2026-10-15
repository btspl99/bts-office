/* BTS Slides 1.3 — Save as OpenDocument Presentation (.odp) for LibreOffice Impress.
 * Written with JSZip: slide size, backgrounds (colour / gradient / picture), theme graphics (as a picture at the back),
 * text boxes and shapes (any BTS shape, incl. SmartArt pieces, as custom shapes with their exact outline) with rich text
 * (bold / italic / underline / strike / size / colour / font / alignment / bullets & numbering / line spacing),
 * lines and arrows, pictures (cropped / rounded), icons, charts (as pictures), equations (pictures + alt text), tables,
 * groups, rotation & flips, hidden slides, transitions + timings and speaker notes.
 * Units: model px (96/in) → inches; font sizes in pt.
 */
(function () {
  const X = window.SLX = window.SLX || { cmd: {}, ribbon: [] };
  const O = window.ODPExport = {};
  const P = () => SL.pres;
  const xe = (t) => String(t ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const IN = (px) => (Math.round(px / 96 * 10000) / 10000) + 'in';
  const PT = (pt) => (Math.round(pt * 100) / 100) + 'pt';
  const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"';

  function hexOf(c, pres) {
    let v = SL.color(c, pres); if (!v || v === 'none' || v === 'transparent') return null;
    if (v.startsWith('rgb')) { const m = v.match(/[\d.]+/g).map(Number); return { c: '#' + m.slice(0, 3).map(x => Math.round(x).toString(16).padStart(2, '0')).join(''), a: m[3] ?? 1 }; }
    if (!v.startsWith('#')) { const d = document.createElement('i'); d.style.color = v; document.body.appendChild(d); const cs = getComputedStyle(d).color; d.remove(); const m = cs.match(/[\d.]+/g); if (!m) return null; return { c: '#' + m.slice(0, 3).map(x => (+x).toString(16).padStart(2, '0')).join(''), a: m[3] !== undefined ? +m[3] : 1 }; }
    v = v.slice(1); if (v.length === 3) v = v.split('').map(x => x + x).join('');
    return { c: '#' + v.slice(0, 6).toLowerCase(), a: v.length === 8 ? parseInt(v.slice(6), 16) / 255 : 1 };
  }

  O.build = async (pres = P()) => {
    await SL.ensureFonts(pres);
    const W = SL.W(pres), H = SL.H(pres);
    const zip = new JSZip();
    zip.file('mimetype', 'application/vnd.oasis.opendocument.presentation', { compression: 'STORE' });
    /* ---------- shared resources ---------- */
    const pics = new Map(); const manifest = [];
    const pic = (url) => {
      if (pics.has(url)) return pics.get(url);
      const m = /^data:([^;,]+)[;,]/.exec(url) || []; const mime = m[1] || 'image/png';
      const ext = mime.includes('jpeg') ? 'jpg' : mime.includes('gif') ? 'gif' : mime.includes('svg') ? 'svg' : mime.includes('webp') ? 'webp' : 'png';
      const path = `Pictures/img${pics.size + 1}.${ext}`; zip.file(path, BTS.dataURLToBytes(url)); manifest.push([path, mime]); pics.set(url, path); return path;
    };
    const auto = new Map(); let nAuto = 0; // automatic styles (content.xml)
    const sty = (family, prefix, inner) => { const k = family + '|' + inner; if (auto.has(k)) return auto.get(k).name; const name = prefix + (++nAuto); auto.set(k, { name, xml: `<style:style style:name="${name}" style:family="${family}"${family === 'graphic' ? ' style:parent-style-name="standard"' : ''}>${inner}</style:style>` }); return name; };
    const common = []; const grads = new Map(); const fillImgs = new Map(); const dashes = new Set();
    const gradient = (c1, c2, cssAngle) => { const k = [c1, c2, cssAngle].join(); if (grads.has(k)) return grads.get(k); const nm = 'Gradient_' + (grads.size + 1); grads.set(k, nm); common.push(`<draw:gradient draw:name="${nm}" draw:style="linear" draw:start-color="${c1}" draw:end-color="${c2}" draw:start-intensity="100%" draw:end-intensity="100%" draw:angle="${Math.round(((180 - (cssAngle ?? 180)) % 360 + 360) % 360)}deg" draw:border="0%"/>`); return nm; };
    const fillImage = (path) => { if (fillImgs.has(path)) return fillImgs.get(path); const nm = 'Bitmap_' + (fillImgs.size + 1); fillImgs.set(path, nm); common.push(`<draw:fill-image draw:name="${nm}" xlink:href="${path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`); return nm; };
    const DASH = { dash: [4, 3], lgDash: [8, 3], dot: [1, 2], dashDot: [4, 2], sysDash: [3, 1] };
    const dashStyle = (k) => { const d = DASH[k]; if (!d) return null; const nm = 'Dash_' + k; if (!dashes.has(nm)) { dashes.add(nm); common.push(`<draw:stroke-dash draw:name="${nm}" draw:style="rect" draw:dots1="1" draw:dots1-length="${d[0] * 100}%" draw:distance="${d[1] * 100}%"/>`); } return nm; };

    /* ---------- text ---------- */
    const textXML = (el) => {
      const runs = PPTXExport.runs(el, pres); const paras = []; let cur = [];
      runs.forEach(r => { cur.push(r); if (r.options.breakLine) { paras.push(cur); cur = []; } }); if (cur.length) paras.push(cur);
      let out = ''; let prevNum = false;
      paras.forEach(pa => {
        const o = pa[0].options; const al = { left: 'start', center: 'center', right: 'end', justify: 'justify' }[o.align] || 'start';
        let pp = `fo:text-align="${al}"`; if (o.lineSpacingMultiple) pp += ` fo:line-height="${Math.round(o.lineSpacingMultiple * 100)}%"`; if (o.paraSpaceAfter) pp += ` fo:margin-bottom="${PT(o.paraSpaceAfter)}"`;
        const ps = sty('paragraph', 'P', `<style:paragraph-properties ${pp}/>`);
        let body = '';
        pa.forEach(r => {
          const q = r.options; let tp = '';
          if (q.fontFace) tp += ` style:font-name="${xe(q.fontFace)}" fo:font-family="'${xe(q.fontFace)}'"`;
          if (q.fontSize) tp += ` fo:font-size="${PT(q.fontSize)}"`;
          if (q.color) tp += ` fo:color="#${q.color.toLowerCase()}"`;
          tp += ` fo:font-weight="${q.bold ? 'bold' : 'normal'}" fo:font-style="${q.italic ? 'italic' : 'normal'}"`;
          if (q.underline) tp += ' style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"';
          if (q.strike) tp += ' style:text-line-through-style="solid"';
          if (q.superscript) tp += ' style:text-position="super 58%"'; else if (q.subscript) tp += ' style:text-position="sub 58%"';
          if (q.highlight) tp += ` fo:background-color="#${q.highlight.toLowerCase()}"`;
          if (q.charSpacing) tp += ` fo:letter-spacing="${PT(q.charSpacing)}"`;
          const ts = sty('text', 'T', `<style:text-properties${tp}/>`);
          if (r.softBreakBefore || q.softBreakBefore) body += '<text:line-break/>';
          const t = xe(r.text).replace(/\t/g, '<text:tab/>').replace(/ {2,}/g, m => ' <text:s text:c="' + (m.length - 1) + '"/>');
          body += q.hyperlink && q.hyperlink.url ? `<text:a xlink:type="simple" xlink:href="${xe(q.hyperlink.url)}"><text:span text:style-name="${ts}">${t}</text:span></text:a>` : `<text:span text:style-name="${ts}">${t}</text:span>`;
        });
        const p = `<text:p text:style-name="${ps}">${body}</text:p>`;
        if (o.bullet) {
          const num = typeof o.bullet === 'object' && o.bullet.type === 'number'; const lv = o.indentLevel || 0;
          out += `<text:list text:style-name="${num ? 'LNum' : 'LBul'}"${num && prevNum ? ' text:continue-numbering="true"' : ''}>` + '<text:list-item><text:list>'.repeat(lv) + `<text:list-item>${p}</text:list-item>` + '</text:list></text:list-item>'.repeat(lv) + '</text:list>';
          prevNum = num;
        } else { out += p; prevNum = false; }
      });
      return out;
    };

    /* ---------- geometry helpers ---------- */
    const place = (el) => {
      const w = Math.max(.5, el.w), h = Math.max(.5, el.h);
      if (!el.rot) return `svg:x="${IN(el.x)}" svg:y="${IN(el.y)}" svg:width="${IN(w)}" svg:height="${IN(h)}"`;
      const th = el.rot * Math.PI / 180, cs = Math.cos(th), sn = Math.sin(th), cx = w / 2, cy = h / 2;
      const tx = el.x + cx - (cs * cx - sn * cy), ty = el.y + cy - (sn * cx + cs * cy);
      return `svg:width="${IN(w)}" svg:height="${IN(h)}" draw:transform="rotate (${(-th).toFixed(6)}) translate (${IN(tx)} ${IN(ty)})"`;
    };
    const strokeProps = (el, isLine) => {
      const c = hexOf(el.stroke, pres);
      if (!c || (!isLine && !el.strokeWidth)) { if (!isLine) return 'draw:stroke="none"'; const t = hexOf('text', pres); return `draw:stroke="solid" svg:stroke-color="${t ? t.c : '#000000'}" svg:stroke-width="${PT(1.5)}"`; }
      const ds = el.dash && el.dash !== 'solid' ? dashStyle(el.dash) : null;
      let s = `draw:stroke="${ds ? 'dash' : 'solid'}" svg:stroke-color="${c.c}" svg:stroke-width="${PT(el.strokeWidth || 1.5)}"`; if (ds) s += ` draw:stroke-dash="${ds}"`;
      const op = (el.opacity ?? 1) * c.a; if (op < 1) s += ` svg:stroke-opacity="${Math.round(op * 100)}%"`;
      return s;
    };
    const fillProps = (el) => {
      const f = el.fill; const op = el.opacity ?? 1;
      if (f && typeof f === 'object' && f.type === 'gradient') { const a = hexOf(f.color, pres), b = hexOf(f.color2, pres); if (a && b) return `draw:fill="gradient" draw:fill-gradient-name="${gradient(a.c, b.c, f.angle ?? 90)}"${op < 1 ? ` draw:opacity="${Math.round(op * 100)}%"` : ''}`; }
      const c = hexOf(f, pres); if (!c) return 'draw:fill="none"';
      const o2 = op * c.a; return `draw:fill="solid" draw:fill-color="${c.c}"${o2 < 1 ? ` draw:opacity="${Math.round(o2 * 100)}%"` : ''}`;
    };
    const shadowProps = (el) => el.shadow ? ' draw:shadow="visible" draw:shadow-offset-x="0.03in" draw:shadow-offset-y="0.03in" draw:shadow-color="#000000" draw:shadow-opacity="30%" draw:shadow-blur="0.06in"' : '';
    const titleDesc = (el) => (el.name ? `<svg:title>${xe(el.name)}</svg:title>` : '') + (el.alt ? `<svg:desc>${xe(el.alt)}</svg:desc>` : '');
    const S = (n) => Math.round(n * 100);

    /* ---------- one element → XML ---------- */
    async function elXML(el, layer) {
      if (el.hidden) return '';
      if (el.ph && ((el.type === 'text' && SL.isEmptyText(el)) || (el.type === 'image' && !el.src))) return '';
      const nm = el.name ? ` draw:name="${xe(el.name)}"` : '';
      const lay = layer ? ' draw:layer="layout"' : '';
      if (el.type === 'text' || el.type === 'shape') {
        if (SL.isLine(el)) {
          let x1 = el.flipH ? el.x + el.w : el.x, y1 = el.flipV ? el.y + el.h : el.y, x2 = el.flipH ? el.x : el.x + el.w, y2 = el.flipV ? el.y : el.y + el.h;
          if (el.rot) { const th = el.rot * Math.PI / 180, cx = el.x + el.w / 2, cy = el.y + el.h / 2, r = (x, y) => [cx + (x - cx) * Math.cos(th) - (y - cy) * Math.sin(th), cy + (x - cx) * Math.sin(th) + (y - cy) * Math.cos(th)];[x1, y1] = r(x1, y1);[x2, y2] = r(x2, y2); }
          const gs = sty('graphic', 'gr', `<style:graphic-properties ${strokeProps(el, true)}${el.kind === 'arrowLine' ? ` draw:marker-end="Arrow" draw:marker-end-width="${PT(Math.max(6, (el.strokeWidth || 1.5) * 4))}"` : ''} draw:fill="none"${shadowProps(el)}/>`);
          return `<draw:line draw:style-name="${gs}"${nm}${lay} svg:x1="${IN(x1)}" svg:y1="${IN(y1)}" svg:x2="${IN(x2)}" svg:y2="${IN(y2)}">${titleDesc(el)}</draw:line>`;
        }
        const kind = el.type === 'text' ? 'rect' : (el.kind || 'rect');
        const g = SL.geom(kind, el.w, el.h, el); const ins = el.inset || [9.6, 4.8, 9.6, 4.8];
        const va = { top: 'top', middle: 'middle', bottom: 'bottom' }[el.valign] || 'top';
        let gp = `${fillProps(el)} ${strokeProps(el, false)} draw:textarea-vertical-align="${va}" draw:textarea-horizontal-align="justify" draw:auto-grow-height="false" draw:auto-grow-width="false" fo:min-height="0in" fo:wrap-option="${el.wrap === false ? 'no-wrap' : 'wrap'}" fo:padding-left="${IN(ins[0])}" fo:padding-top="${IN(ins[1])}" fo:padding-right="${IN(ins[2])}" fo:padding-bottom="${IN(ins[3])}"${shadowProps(el)}`;
        if (el.dir === 'vert' || el.dir === 'vert270') gp += ' style:writing-mode="tb-rl"';
        const gs = sty('graphic', 'gr', `<style:graphic-properties ${gp}/>`);
        const path = SL.pathSegs(g.d).map(s => s.c === 'Z' ? 'Z' : s.c + ' ' + s.p.map(S).join(' ')).join(' ') + ' N';
        const ti = g.inset || [0, 0, 0, 0];
        const ta = `${S(ti[0] * el.w)} ${S(ti[1] * el.h)} ${S(el.w - ti[2] * el.w)} ${S(el.h - ti[3] * el.h)}`;
        const geo = `<draw:enhanced-geometry svg:viewBox="0 0 ${S(el.w)} ${S(el.h)}" draw:type="non-primitive" draw:enhanced-path="${path}" draw:text-areas="${ta}"${el.flipH ? ' draw:mirror-horizontal="true"' : ''}${el.flipV ? ' draw:mirror-vertical="true"' : ''}/>`;
        const txt = SL.isEmptyText(el) ? '' : textXML(el);
        return `<draw:custom-shape draw:style-name="${gs}"${nm}${lay} ${place(el)}>${titleDesc(el)}${txt}${geo}</draw:custom-shape>`;
      }
      let url = null, alt = el.alt || '';
      if (el.type === 'image') url = await PPTXExport.processImage(el);
      else if (el.type === 'icon') url = await PPTXExport.svgToPNG(`<svg xmlns="http://www.w3.org/2000/svg" width="${el.w}" height="${el.h}" viewBox="0 0 24 24" preserveAspectRatio="none" fill="${el.fillIcon ? SL.color(el.color || 'accent1', pres) : 'none'}" stroke="${SL.color(el.color || 'accent1', pres)}" stroke-width="${el.strokeWidth || 1.75}" stroke-linecap="round" stroke-linejoin="round">${(window.BTS_ICONS || {})[el.name] || ''}</svg>`, el.w, el.h, 4), alt = alt || 'Icon: ' + el.name;
      else if (el.type === 'chart') url = await PPTXExport.svgToPNG(SL.chartSVG(el.chart, el.w, el.h, pres).replace('<svg ', '<svg xmlns="http://www.w3.org/2000/svg" '), el.w, el.h, 3), alt = alt || 'Chart' + (el.chart && el.chart.title ? ': ' + el.chart.title : '');
      else if (el.type === 'equation' && SL.equationPNG) url = await SL.equationPNG(el, pres, 4), alt = 'Equation: ' + (el.eq || '');
      else if (el.type === 'table') return tableXML(el, nm, lay);
      if (!url) return '';
      const op = el.opacity ?? 1;
      const gs = sty('graphic', 'gr', `<style:graphic-properties draw:stroke="none" draw:fill="none"${op < 1 ? ` draw:image-opacity="${Math.round(op * 100)}%"` : ''}${shadowProps(el)}${el.flipH || el.flipV ? ` style:mirror="${[el.flipH && 'horizontal', el.flipV && 'vertical'].filter(Boolean).join(' ')}"` : ''}/>`);
      return `<draw:frame draw:style-name="${gs}"${nm}${lay} ${place(el)}><draw:image xlink:href="${pic(url)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"><text:p/></draw:image>${el.name ? `<svg:title>${xe(el.name)}</svg:title>` : ''}${alt && !el.decorative ? `<svg:desc>${xe(alt)}</svg:desc>` : ''}</draw:frame>`;
    }
    function tableXML(el, nm, lay) {
      const n = el.rows.length, nc = (el.rows[0] || []).length; if (!n || !nc) return '';
      const cols = (el.colW || Array(nc).fill(1 / nc)).map(c => sty('table-column', 'co', `<style:table-column-properties style:column-width="${IN(c * el.w)}"/>`));
      let x = `<draw:frame draw:style-name="${sty('graphic', 'gr', '<style:graphic-properties draw:stroke="none" draw:fill="none"/>')}"${nm}${lay} ${place(el)}><table:table>`;
      x += cols.map(c => `<table:table-column table:style-name="${c}"/>`).join('');
      el.rows.forEach((row, r) => {
        const rh = el.rowH?.[r] ? el.rowH[r] * el.h : el.h / n;
        x += `<table:table-row table:style-name="${sty('table-row', 'ro', `<style:table-row-properties style:row-height="${IN(rh)}"/>`)}">`;
        row.forEach((c, ci) => {
          if (c.merged) { x += '<table:covered-table-cell/>'; return; }
          const ts = SL.tableCellStyle(el, r, ci, pres);
          const bg = c.fill === 'none' ? null : hexOf(c.fill || ts.bg, pres);
          const bd = ['top', 'right', 'bottom', 'left'].map((s, i) => { const b = ts.b[i]; const bc = b && hexOf(b.c, pres); return `fo:border-${s}="${b && bc ? PT(b.w * .75) + ' solid ' + bc.c : 'none'}"`; }).join(' ');
          const va = { top: 'top', middle: 'middle', bottom: 'bottom' }[c.valign || 'middle'];
          const cs = sty('table-cell', 'ce', `<style:graphic-properties ${bg ? `draw:fill="solid" draw:fill-color="${bg.c}"` : 'draw:fill="none"'} draw:textarea-vertical-align="${va}" fo:padding-top="0.04in" fo:padding-bottom="0.04in" fo:padding-left="0.07in" fo:padding-right="0.07in"/><style:paragraph-properties ${bd}/>`);
          const color = c.color ? SL.color(c.color, pres) : ts.color;
          const cellEl = { html: c.html || '<p></p>', size: el.size || 16, font: el.font || SL.fontOf('body', pres), color: color || '#1f2430', bold: ts.bold || c.bold, align: c.align || 'left', role: 'body', type: 'text' };
          x += `<table:table-cell table:style-name="${cs}"${c.cs > 1 ? ` table:number-columns-spanned="${c.cs}"` : ''}${c.rs > 1 ? ` table:number-rows-spanned="${c.rs}"` : ''}>${textXML(cellEl) || '<text:p/>'}</table:table-cell>`;
        });
        x += '</table:table-row>';
      });
      return x + '</table:table></draw:frame>';
    }

    /* ---------- slides ---------- */
    const decoCache = {}; let pages = '';
    for (let i = 0; i < pres.slides.length; i++) {
      const s = pres.slides[i];
      const bg = s.background || SL.themeBg(pres, s.layout);
      let dp = 'presentation:background-visible="true" presentation:background-objects-visible="true"';
      if (bg.type === 'image' && bg.src) { const d = await PPTXExport.bgImage(bg, W, H); dp += d ? ` draw:fill="bitmap" draw:fill-image-name="${fillImage(pic(d))}" style:repeat="stretch"` : ' draw:fill="solid" draw:fill-color="#ffffff"'; }
      else if (bg.type === 'gradient') { const a = hexOf(bg.color, pres), b = hexOf(bg.color2, pres); dp += ` draw:fill="gradient" draw:fill-gradient-name="${gradient(a ? a.c : '#ffffff', b ? b.c : '#ffffff', bg.angle ?? 135)}"`; }
      else { const c = hexOf(bg.color || 'bg', pres); dp += ` draw:fill="solid" draw:fill-color="${c ? c.c : '#ffffff'}"`; }
      if (s.hidden) dp += ' presentation:visibility="hidden"';
      const tr = s.transition || {};
      if (tr.type && tr.type !== 'none') {
        const d4 = { left: 'fromRight', right: 'fromLeft', up: 'fromBottom', down: 'fromTop' }[tr.dir || 'left'];
        const sm = { fade: ['fade', 'crossfade'], push: ['pushWipe', d4], cover: ['slideWipe', d4], reveal: ['slideWipe', d4], wipe: ['barWipe', 'leftToRight'], split: ['barnDoorWipe', 'vertical'], zoom: ['zoom', 'rotateIn'] }[tr.type] || ['fade', 'crossfade'];
        dp += ` smil:type="${sm[0]}" smil:subtype="${sm[1]}" presentation:transition-speed="${(tr.dur || 700) < 500 ? 'fast' : (tr.dur || 700) < 1000 ? 'medium' : 'slow'}"`;
      }
      if (tr.auto != null && tr.auto !== false) dp += ` presentation:transition-type="automatic" presentation:duration="PT${Math.max(0, +tr.auto)}S"`;
      const dps = sty('drawing-page', 'dp', `<style:drawing-page-properties ${dp}/>`);
      let body = '';
      if (!s.hideGraphics && !s.noDeco) {
        const key = SL.decoKind(s.layout) + (s.layout === 'pictureCaption' ? ':pc' : '');
        if (!(key in decoCache)) { const svg = SL.decoSVG(Object.assign({}, s, { hideGraphics: false }), pres); decoCache[key] = svg ? await PPTXExport.svgToPNG(svg.replace('<svg ', '<svg xmlns="http://www.w3.org/2000/svg" '), W, H, 1) : null; }
        if (decoCache[key]) body += `<draw:frame draw:style-name="${sty('graphic', 'gr', '<style:graphic-properties draw:stroke="none" draw:fill="none"/>')}" draw:name="Theme graphics" svg:x="0in" svg:y="0in" svg:width="${IN(W)}" svg:height="${IN(H)}"><draw:image xlink:href="${pic(decoCache[key])}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"><text:p/></draw:image></draw:frame>`;
      }
      if (!s.hideGraphics && pres.master) for (const el of [...(pres.master.graphics || []), ...((pres.master.layoutGraphics || {})[s.layout] || [])]) body += await elXML(el, false);
      for (const el of (s.hideGraphics ? [] : s.layer || [])) body += await elXML(el, false);
      const done = new Set();
      for (const el of s.elements) {
        if (done.has(el.id)) continue;
        if (el.g) {
          const mem = s.elements.filter(e => e.g === el.g); mem.forEach(m => done.add(m.id));
          let inner = ''; for (const m of mem) inner += await elXML(m, false);
          body += mem.length > 1 && inner ? `<draw:g>${inner}</draw:g>` : inner; continue;
        }
        done.add(el.id); body += await elXML(el, false);
      }
      for (const el of SL.footerEls(s, pres, i)) body += await elXML(el, false);
      const notes = (s.notes || '').trim();
      const nb = notes ? notes.split('\n').map(l => `<text:p>${xe(l)}</text:p>`).join('') : '<text:p/>';
      body += `<presentation:notes><draw:page-thumbnail draw:layer="layout" svg:x="1.1in" svg:y="0.8in" svg:width="6.3in" svg:height="${IN(6.3 * 96 * H / W)}" draw:page-number="${i + 1}" presentation:class="page"/><draw:frame presentation:class="notes" svg:x="0.9in" svg:y="${IN((0.8 + 6.3 * H / W + 0.4) * 96)}" svg:width="6.7in" svg:height="4.6in"><draw:text-box>${nb}</draw:text-box></draw:frame></presentation:notes>`;
      pages += `<draw:page draw:name="page${i + 1}" draw:style-name="${dps}" draw:master-page-name="Default">${body}</draw:page>`;
    }

    /* ---------- parts ---------- */
    const lists = `<text:list-style style:name="LBul">${[0, 1, 2, 3, 4, 5].map(l => `<text:list-level-style-bullet text:level="${l + 1}" text:bullet-char="${['•', '–', '•', '–', '•', '–'][l]}"><style:list-level-properties text:space-before="${(l * .3 + .05).toFixed(2)}in" text:min-label-width="0.25in"/><style:text-properties fo:font-family="sans-serif" style:font-size-rel="100%"/></text:list-level-style-bullet>`).join('')}</text:list-style>` +
      `<text:list-style style:name="LNum">${[0, 1, 2, 3, 4, 5].map(l => `<text:list-level-style-number text:level="${l + 1}" style:num-suffix="." style:num-format="${['1', 'a', 'i', '1', 'a', 'i'][l]}"><style:list-level-properties text:space-before="${(l * .3 + .05).toFixed(2)}in" text:min-label-width="0.3in"/><style:text-properties style:font-size-rel="100%"/></text:list-level-style-number>`).join('')}</text:list-style>`;
    const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3"><office:automatic-styles>${[...auto.values()].map(a => a.xml).join('')}${lists}</office:automatic-styles><office:body><office:presentation>${pages}</office:presentation></office:body></office:document-content>`;
    const fonts = SL.fontOf('body', pres), tfont = SL.fontOf('title', pres);
    const styles = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS} office:version="1.3"><office:styles>` +
      `<draw:marker draw:name="Arrow" svg:viewBox="0 0 20 30" svg:d="M10 0l-10 30h20z"/>${common.join('')}` +
      `<style:default-style style:family="graphic"><style:graphic-properties draw:shadow="hidden"/><style:paragraph-properties fo:text-align="start"/><style:text-properties fo:font-family="'${xe(fonts)}'" style:font-name="${xe(fonts)}" fo:font-size="18pt" fo:language="en" fo:country="IN"/></style:default-style>` +
      `<style:style style:name="standard" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none" fo:wrap-option="wrap"/><style:text-properties fo:font-family="'${xe(fonts)}'" fo:font-size="18pt"/></style:style>` +
      `</office:styles><office:automatic-styles><style:page-layout style:name="PM1"><style:page-layout-properties fo:margin-top="0in" fo:margin-bottom="0in" fo:margin-left="0in" fo:margin-right="0in" fo:page-width="${IN(W)}" fo:page-height="${IN(H)}" style:print-orientation="${W >= H ? 'landscape' : 'portrait'}"/></style:page-layout>` +
      `<style:style style:name="Mdp1" style:family="drawing-page"><style:drawing-page-properties draw:fill="none"/></style:style></office:automatic-styles>` +
      `<office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1" draw:style-name="Mdp1"/></office:master-styles></office:document-styles>`;
    void tfont;
    const pr = Object.assign({}, BTS.props || {});
    const meta = `<?xml version="1.0" encoding="UTF-8"?><office:document-meta ${NS} office:version="1.3"><office:meta><meta:generator>BTS Slides (BTS Office)</meta:generator><dc:title>${xe(pr.title || BTS.title || 'Presentation')}</dc:title>${pr.subject ? `<dc:subject>${xe(pr.subject)}</dc:subject>` : ''}${pr.comments ? `<dc:description>${xe(pr.comments)}</dc:description>` : ''}${pr.keywords ? `<meta:keyword>${xe(pr.keywords)}</meta:keyword>` : ''}<meta:initial-creator>${xe(pr.author || BTS.userName || 'BTS Slides')}</meta:initial-creator><dc:creator>${xe(BTS.userName || pr.author || 'BTS Slides')}</dc:creator><dc:date>${new Date().toISOString().replace(/\.\d+Z$/, '')}</dc:date></office:meta></office:document-meta>`;
    zip.file('content.xml', content); zip.file('styles.xml', styles); zip.file('meta.xml', meta);
    zip.file('META-INF/manifest.xml', `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.presentation"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>${manifest.map(([p, m]) => `<manifest:file-entry manifest:full-path="${p}" manifest:media-type="${m}"/>`).join('')}</manifest:manifest>`);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.oasis.opendocument.presentation' });
  };

  X.cmd['save-odp'] = () => BTS.saveAs('odp');
})();
