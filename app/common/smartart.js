/* BTS Office — SmartArt-style diagrams (shared by BTS Slides and BTS Docs). Self-contained: no app dependencies
 * (the dialog uses BTS.dialog when it is present).
 *
 *   BTSSmartArt.LAYOUTS = [{ id, name, category, desc, thumbSVG, sample:[items] }]
 *   BTSSmartArt.layout(id, items, { w, h, palette:[hex…], font, textColor, lineColor, maxFont, minFont })
 *       items = [{ text, level }]  (level 0 = main item, 1 = sub-item, 2 …)
 *     → { w, h, id, shapes:[shape] }   all sizes (x, y, w, h, strokeW, fontSize, pad) are in the caller's units.
 *   shape = { kind, x, y, w, h, rot, fill, stroke, strokeW, opacity, text, textColor, fontSize, bold, align, valign,
 *             pad:[l,t,r,b] (text insets incl. the shape's own geometry), flipV, inset, path, role, item }
 *     kind: 'rect' | 'roundRect' | 'ellipse' | 'chevron' | 'pentagon' (arrow pentagon / "home plate") | 'arrowRight' |
 *           'triangle' | 'trapezoid' (top edge narrower by `inset` on each side; flipV:true → bottom edge narrower) |
 *           'circleArrow' (curved block arrow; `path` = SVG path in the shape's own box, "x,y" pairs, M/L/C/Z only) |
 *           'line' | 'connector' (from (x, y) to (x + w, y + h); w/h may be negative; connector has an arrow head at the end)
 *     text: lines separated by "\n"; lines starting with "• " are sub-item bullets. fill/stroke: hex or 'none'.
 *     item: index of the source item when the shape shows exactly one item's text (for editing in place).
 *   BTSSmartArt.toSVG(model, { font, background }) → standalone SVG string (width/height = model.w/h).
 *   BTSSmartArt.dialog({ palette, initial:{ id, items }, font, title }) → Promise<{ id, items } | null>
 *   BTSSmartArt.parseOutline(text) / toOutline(items)  — text outline (Tab = one level) ⇄ items.
 *   BTSSmartArt.shapePath(shape) → SVG path data (absolute coordinates) for a non-line shape.
 */
(function (root) {
  const SA = {};
  const f2 = (n) => Math.round(n * 100) / 100;
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const DEF_PAL = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];

  /* ---------- colours ---------- */
  function rgb(hex) { let h = String(hex || '#888').replace('#', ''); if (h.length === 3) h = h.split('').map(c => c + c).join(''); const n = parseInt(h.slice(0, 6), 16) || 0; return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; }
  const toHex = (a) => '#' + a.map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('');
  /** t > 0 → towards white, t < 0 → towards black */
  SA.mix = (hex, t) => { const c = rgb(hex); return toHex(c.map(v => t >= 0 ? v + (255 - v) * t : v * (1 + t))); };
  const lum = (hex) => { const c = rgb(hex).map(v => { v /= 255; return v <= .03928 ? v / 12.92 : Math.pow((v + .055) / 1.055, 2.4); }); return .2126 * c[0] + .7152 * c[1] + .0722 * c[2]; };
  const onColor = (fill, dark) => lum(fill) > .45 ? dark : '#ffffff';

  /* ---------- text measuring & fitting ---------- */
  let ctx = null;
  const fontCSS = (f) => `'${String(f || 'Inter').replace(/['"]/g, '')}', sans-serif`;
  function measure(text, size, font, bold) {
    if (!ctx && typeof document !== 'undefined') { try { ctx = document.createElement('canvas').getContext('2d'); } catch { } }
    if (!ctx) return text.length * size * .56 * (bold ? 1.06 : 1);
    ctx.font = (bold ? '700 ' : '400 ') + size + 'px ' + fontCSS(font);
    return ctx.measureText(text).width;
  }
  function wrap(text, maxW, size, font, bold) {
    const out = [];
    String(text || '').split('\n').forEach(par => {
      const words = par.split(/\s+/).filter(Boolean); if (!words.length) { out.push(''); return; }
      let line = '';
      words.forEach(w => {
        const t = line ? line + ' ' + w : w;
        if (!line || measure(t, size, font, bold) <= maxW) { line = t; }
        else { out.push(line); line = w; }
        while (measure(line, size, font, bold) > maxW && line.length > 1) { // a single long word: break it
          let k = line.length - 1; while (k > 1 && measure(line.slice(0, k), size, font, bold) > maxW) k--;
          out.push(line.slice(0, k)); line = line.slice(k);
        }
      });
      out.push(line);
    });
    return out;
  }
  SA.wrap = wrap;
  const LH = 1.2;
  function fits(text, w, h, size, font, bold) {
    if (w <= 1 || h <= 1) return false;
    if (String(text).split(/\s+/).some(wd => measure(wd, size, font, bold) > w)) return false; // shrink rather than break a word
    const lines = wrap(text, w, size, font, bold);
    return lines.length * size * LH <= h && lines.every(l => measure(l, size, font, bold) <= w + .5);
  }
  function fitSize(text, w, h, font, bold, max, min) {
    if (!String(text || '').trim()) return max;
    if (fits(text, w, h, max, font, bold)) return max;
    let lo = min, hi = max;
    for (let i = 0; i < 14; i++) { const m = (lo + hi) / 2; if (fits(text, w, h, m, font, bold)) lo = m; else hi = m; }
    return lo;
  }

  /* ---------- geometry ---------- */
  // fractions [l, t, r, b] of the box that the shape's outline takes away from the text area (same as BTS Slides)
  function geoInset(s) {
    const w = s.w, h = s.h, m = Math.min(w, h);
    switch (s.kind) {
      case 'roundRect': return [.05, .05, .05, .05];
      case 'ellipse': return [.146, .146, .146, .146];
      case 'triangle': return [.25, .5, .25, 0];
      case 'trapezoid': { const o = Math.min(s.inset ?? m * .25, w / 2) / w; return [o * .8, 0, o * .8, 0]; }
      case 'chevron': { const o = Math.min(w / 2, h / 2) / w; return [o, 0, o, 0]; }
      case 'pentagon': { const o = Math.min(w / 2, h / 2) / w; return [0, 0, o * .7, 0]; }
      case 'arrowRight': { const hx = w - Math.min(w, h) * .5; return [0, .25, (w - hx) / w * .5, .25]; }
      default: return [0, 0, 0, 0];
    }
  }
  const poly = (pts) => 'M' + pts.map(p => f2(p[0]) + ',' + f2(p[1])).join(' L') + ' Z';
  /** SVG path (absolute coordinates) of a closed shape. */
  SA.shapePath = (s) => {
    const { x, y, w, h } = s; const m = Math.min(w, h); const P = (px, py) => [x + px, y + py];
    switch (s.kind) {
      case 'ellipse': { const rx = w / 2, ry = h / 2, cx = x + rx, cy = y + ry, k = .5523; return `M${f2(cx - rx)},${f2(cy)} C${f2(cx - rx)},${f2(cy - ry * k)} ${f2(cx - rx * k)},${f2(cy - ry)} ${f2(cx)},${f2(cy - ry)} C${f2(cx + rx * k)},${f2(cy - ry)} ${f2(cx + rx)},${f2(cy - ry * k)} ${f2(cx + rx)},${f2(cy)} C${f2(cx + rx)},${f2(cy + ry * k)} ${f2(cx + rx * k)},${f2(cy + ry)} ${f2(cx)},${f2(cy + ry)} C${f2(cx - rx * k)},${f2(cy + ry)} ${f2(cx - rx)},${f2(cy + ry * k)} ${f2(cx - rx)},${f2(cy)} Z`; }
      case 'roundRect': { const r = Math.min(m * .1667, w / 2, h / 2), k = r * .4477; return `M${f2(x + r)},${f2(y)} L${f2(x + w - r)},${f2(y)} C${f2(x + w - k)},${f2(y)} ${f2(x + w)},${f2(y + k)} ${f2(x + w)},${f2(y + r)} L${f2(x + w)},${f2(y + h - r)} C${f2(x + w)},${f2(y + h - k)} ${f2(x + w - k)},${f2(y + h)} ${f2(x + w - r)},${f2(y + h)} L${f2(x + r)},${f2(y + h)} C${f2(x + k)},${f2(y + h)} ${f2(x)},${f2(y + h - k)} ${f2(x)},${f2(y + h - r)} L${f2(x)},${f2(y + r)} C${f2(x)},${f2(y + k)} ${f2(x + k)},${f2(y)} ${f2(x + r)},${f2(y)} Z`; }
      case 'triangle': return poly([P(w / 2, 0), P(w, h), P(0, h)]);
      case 'trapezoid': { const o = Math.min(s.inset ?? m * .25, w / 2); return s.flipV ? poly([P(0, 0), P(w, 0), P(w - o, h), P(o, h)]) : poly([P(o, 0), P(w - o, 0), P(w, h), P(0, h)]); }
      case 'chevron': { const o = Math.min(w / 2, h / 2); return poly([P(0, 0), P(w - o, 0), P(w, h / 2), P(w - o, h), P(0, h), P(o, h / 2)]); }
      case 'pentagon': { const o = Math.min(w / 2, h / 2); return poly([P(0, 0), P(w - o, 0), P(w, h / 2), P(w - o, h), P(0, h)]); }
      case 'arrowRight': { const hx = w - Math.min(w, h) * .5; return poly([P(0, h * .25), P(hx, h * .25), P(hx, 0), P(w, h / 2), P(hx, h), P(hx, h * .75), P(0, h * .75)]); }
      case 'circleArrow': return String(s.path || '').replace(/(-?\d*\.?\d+(?:e-?\d+)?),(-?\d*\.?\d+(?:e-?\d+)?)/g, (_, a, b) => f2(+a + x) + ',' + f2(+b + y));
      default: return poly([P(0, 0), P(w, 0), P(w, h), P(0, h)]);
    }
  };
  // arc (centre cx,cy radius r, angles a0→a1 in radians, clockwise on screen) as cubic segments "C …" (start point excluded)
  function arcC(cx, cy, r, a0, a1) {
    const n = Math.max(1, Math.ceil(Math.abs(a1 - a0) / (Math.PI / 2))); const d = (a1 - a0) / n; const k = 4 / 3 * Math.tan(d / 4); let s = '';
    for (let i = 0; i < n; i++) {
      const b0 = a0 + d * i, b1 = b0 + d;
      const p1 = [cx + r * (Math.cos(b0) - k * Math.sin(b0)), cy + r * (Math.sin(b0) + k * Math.cos(b0))];
      const p2 = [cx + r * (Math.cos(b1) + k * Math.sin(b1)), cy + r * (Math.sin(b1) - k * Math.cos(b1))];
      const p3 = [cx + r * Math.cos(b1), cy + r * Math.sin(b1)];
      s += ` C${f2(p1[0])},${f2(p1[1])} ${f2(p2[0])},${f2(p2[1])} ${f2(p3[0])},${f2(p3[1])}`;
    }
    return s;
  }
  // curved block arrow along a circle from angle a0 to a1 (clockwise), thickness t → shape with a box-relative path
  function circleArrow(cx, cy, R, t, a0, a1) {
    const head = Math.min((t * 1.5) / R, (a1 - a0) * .6); const ah = a1 - head;
    const ro = R + t / 2, ri = R - t / 2, hw = t * 1.05;
    const pt = (r, a) => [cx + r * Math.cos(a), cy + r * Math.sin(a)];
    const pts = [pt(ro, a0), pt(ro, ah), pt(R + hw, ah), pt(R, a1), pt(R - hw, ah), pt(ri, ah), pt(ri, a0)];
    // bounding box (arc extremes included)
    const all = pts.slice(); for (let a = Math.ceil(a0 / (Math.PI / 2)) * Math.PI / 2; a < ah; a += Math.PI / 2) all.push(pt(ro, a));
    const minX = Math.min(...all.map(p => p[0])), minY = Math.min(...all.map(p => p[1])), maxX = Math.max(...all.map(p => p[0])), maxY = Math.max(...all.map(p => p[1]));
    const L = (p) => f2(p[0] - minX) + ',' + f2(p[1] - minY);
    const rel = (str) => str.replace(/(-?\d*\.?\d+),(-?\d*\.?\d+)/g, (_, a, b) => f2(+a - minX) + ',' + f2(+b - minY));
    const path = `M${L(pts[0])}${rel(arcC(cx, cy, ro, a0, ah))} L${L(pts[2])} L${L(pts[3])} L${L(pts[4])} L${L(pts[5])}${rel(arcC(cx, cy, ri, ah, a0))} Z`;
    return { kind: 'circleArrow', x: minX, y: minY, w: Math.max(1, maxX - minX), h: Math.max(1, maxY - minY), path };
  }

  /* ---------- items → structure ---------- */
  function norm(items) { return (items || []).map(it => typeof it === 'string' ? { text: it, level: 0 } : { text: String(it.text ?? ''), level: Math.max(0, Math.min(6, +it.level || 0)) }).filter(it => it.text.trim() || it.level === 0); }
  // top-level items with their sub-items as bullet lines
  function tops(items) {
    const out = []; items.forEach((it, i) => { if (it.level === 0 || !out.length) out.push({ text: it.text, idx: i, kids: [] }); else out[out.length - 1].kids.push({ text: it.text, idx: i, level: it.level }); });
    return out;
  }
  const withKids = (t) => t.text + (t.kids.length ? '\n' + t.kids.map(k => '• ' + k.text).join('\n') : '');
  const kidsText = (t) => t.kids.map(k => '• ' + k.text).join('\n');
  function tree(items) {
    const rootN = { kids: [], level: -1 }; const stack = [rootN];
    items.forEach((it, i) => { const n = { text: it.text, idx: i, level: it.level, kids: [] }; while (stack.length > 1 && stack[stack.length - 1].level >= it.level) stack.pop(); stack[stack.length - 1].kids.push(n); stack.push(n); });
    return rootN.kids;
  }

  /* ---------- layouts ---------- */
  const S = (o) => Object.assign({ rot: 0, stroke: 'none', strokeW: 0, opacity: 1, text: '', bold: false, align: 'center', valign: 'middle' }, o);
  const L = {};
  const sample = {
    steps: ['Plan', 'Research', 'Create', 'Review'].map(t => ({ text: t, level: 0 })),
    cycle: ['Plant', 'Grow', 'Harvest', 'Sell', 'Save'].map(t => ({ text: t, level: 0 })),
    tree: [{ text: 'Principal', level: 0 }, { text: 'Vice principal', level: 1 }, { text: 'Science', level: 2 }, { text: 'Maths', level: 2 }, { text: 'Admin office', level: 1 }, { text: 'Accounts', level: 2 }],
    radial: [{ text: 'Healthy life', level: 0 }, { text: 'Sleep', level: 1 }, { text: 'Exercise', level: 1 }, { text: 'Food', level: 1 }, { text: 'Friends', level: 1 }, { text: 'Study', level: 1 }],
    pyramid: ['Self-actualization', 'Esteem', 'Belonging', 'Safety', 'Basic needs'].map(t => ({ text: t, level: 0 })),
    funnel: ['Visitors', 'Interested', 'Applied', 'Admitted'].map(t => ({ text: t, level: 0 })),
    venn: ['Skills', 'Passion', 'Need'].map(t => ({ text: t, level: 0 })),
    timeline: ['1947 Independence', '1950 Republic', '1969 ISRO founded', '1991 Reforms', '2023 Chandrayaan-3'].map(t => ({ text: t, level: 0 })),
    list: [{ text: 'Reading', level: 0 }, { text: 'Two chapters a day', level: 1 }, { text: 'Writing', level: 0 }, { text: 'One essay a week', level: 1 }, { text: 'Revision', level: 0 }, { text: 'Every Sunday', level: 1 }],
    matrix: ['Urgent & important', 'Important, not urgent', 'Urgent, not important', 'Neither'].map(t => ({ text: t, level: 0 })),
    blocks: ['Physics', 'Chemistry', 'Biology', 'Maths', 'English', 'Hindi'].map(t => ({ text: t, level: 0 })),
  };

  L.process = (T, W, H, c) => {
    const N = T.length, ar = .38, bw = W / (N + ar * (N - 1)), g = bw * ar;
    const bh = Math.min(H, bw * (T.some(t => t.kids.length) ? 1.1 : .66)), y = (H - bh) / 2; const out = [];
    T.forEach((t, i) => {
      const x = i * (bw + g), col = c.col(i);
      out.push(S({ kind: 'roundRect', x, y, w: bw, h: bh, fill: col, text: withKids(t), role: 'main', item: t.kids.length ? undefined : t.idx }));
      if (i < N - 1) { const aw = g * .62, ah = Math.min(bw * .3, bh * .45); out.push(S({ kind: 'arrowRight', x: x + bw + g * .19, y: H / 2 - ah / 2, w: aw, h: ah, fill: SA.mix(c.col(i), .45) })); }
    });
    return out;
  };
  L.chevron = (T, W, H, c) => {
    const N = T.length, kids = T.some(t => t.kids.length);
    let cw = W / (1 + .832 * (N - 1)); let ch = Math.min(kids ? H * .32 : H, cw * .4);
    const step = cw - ch * .5 + ch * .08; cw = (W - step * (N - 1)) > 0 ? Math.min(cw, W - step * (N - 1)) : cw;
    const total = step * (N - 1) + cw, ox = (W - total) / 2, y = kids ? 0 : (H - ch) / 2; const out = [];
    T.forEach((t, i) => {
      out.push(S({ kind: i === 0 ? 'pentagon' : 'chevron', x: ox + i * step, y, w: cw, h: ch, fill: c.col(i), text: t.text, role: 'main', item: t.idx }));
      if (t.kids.length) out.push(S({ kind: 'rect', x: ox + i * step + ch * .25, y: ch + H * .03, w: cw - ch * .5, h: H - ch - H * .03, fill: 'none', text: kidsText(t), textColor: c.text, role: 'child', align: 'left', valign: 'top' }));
    });
    return out;
  };
  L.stepUp = (T, W, H, c) => {
    const N = T.length, g = Math.min(W, H) * .02, cw = (W - g * (N - 1)) / N, sh = H / N; const out = [];
    T.forEach((t, i) => {
      const hh = sh * (i + 1), col = SA.mix(c.col(0), N > 1 ? .55 * (1 - i / (N - 1)) : 0);
      out.push(S({ kind: 'rect', x: i * (cw + g), y: H - hh, w: cw, h: hh, fill: col, text: withKids(t), role: 'main', valign: 'top', fitH: sh, align: t.kids.length ? 'left' : 'center', item: t.kids.length ? undefined : t.idx }));
    });
    return out;
  };
  L.timeline = (T, W, H, c) => {
    const N = T.length, bh = Math.max(4, H * .07), out = [];
    out.push(S({ kind: 'arrowRight', x: 0, y: H / 2 - bh / 2, w: W, h: bh, fill: SA.mix(c.col(0), .55) }));
    const d = Math.min(H * .07, W / N * .3), cell = (W - bh) / N, alt = N > 2;
    T.forEach((t, i) => {
      const cx = cell * (i + .5), up = !alt || i % 2 === 0;
      out.push(S({ kind: 'ellipse', x: cx - d / 2, y: H / 2 - d / 2, w: d, h: d, fill: c.col(i), stroke: '#ffffff', strokeW: Math.max(1, d * .08) }));
      const lw = alt ? Math.min(cell * 1.8, W * .4) : cell * .94, lh = H / 2 - d * 1.4;
      const lx = Math.max(0, Math.min(W - lw, cx - lw / 2));
      out.push(S({ kind: 'line', x: cx, y: up ? H / 2 - d / 2 : H / 2 + d / 2, w: 0, h: up ? -d * .7 : d * .7, stroke: c.col(i), strokeW: Math.max(1, d * .06) }));
      out.push(S({ kind: 'rect', x: lx, y: up ? 0 : H / 2 + d * 1.4, w: lw, h: lh, fill: 'none', text: withKids(t), textColor: c.text, role: 'label', valign: up ? 'bottom' : 'top', item: t.kids.length ? undefined : t.idx }));
    });
    return out;
  };
  L.cycle = (T, W, H, c) => {
    const N = T.length, M = Math.min(W, H), cx = W / 2, cy = H / 2, out = [];
    if (N === 1) return [S({ kind: 'ellipse', x: cx - M * .2, y: cy - M * .2, w: M * .4, h: M * .4, fill: c.col(0), text: withKids(T[0]), role: 'main', item: T[0].kids.length ? undefined : T[0].idx })];
    const s = Math.sin(Math.PI / N); let d = .58 * M * s / (1 + .58 * s); d = Math.min(d, M * .32);
    const R = M / 2 - d / 2, t = Math.max(2, d * .2);
    T.forEach((it, i) => {
      const a = -Math.PI / 2 + i * 2 * Math.PI / N;
      out.push(S({ kind: 'ellipse', x: cx + R * Math.cos(a) - d / 2, y: cy + R * Math.sin(a) - d / 2, w: d, h: d, fill: c.col(i), text: withKids(it), role: 'main', item: it.kids.length ? undefined : it.idx }));
    });
    const gap = Math.asin(Math.min(.99, (d / 2 + t * .9) / R)) + .02;
    for (let i = 0; i < N; i++) {
      const a0 = -Math.PI / 2 + i * 2 * Math.PI / N + gap, a1 = -Math.PI / 2 + (i + 1) * 2 * Math.PI / N - gap;
      if (a1 - a0 < .08) continue;
      out.push(S(Object.assign(circleArrow(cx, cy, R, t, a0, a1), { fill: SA.mix(c.col(i), .4) })));
    }
    return out;
  };
  L.radial = (T, W, H, c) => {
    let center = T[0], sats = T[0].kids.length ? T[0].kids.map(k => ({ text: k.text, idx: k.idx, kids: [] })) : T.slice(1);
    if (!T[0].kids.length && T.length > 1) sats = T.slice(1).map(t => ({ text: withKids(t), idx: t.kids.length ? undefined : t.idx, kids: [] }));
    const n = sats.length, M = Math.min(W, H), cx = W / 2, cy = H / 2, out = [];
    if (!n) return [S({ kind: 'ellipse', x: cx - M * .2, y: cy - M * .2, w: M * .4, h: M * .4, fill: c.col(0), text: center.text, role: 'center', item: center.idx })];
    const s = Math.sin(Math.PI / Math.max(n, 2)); let ds = .8 * M * s / (1 + .8 * s); ds = Math.min(ds, M * .27);
    const Rx = Math.min(W / 2 - ds / 2, (H / 2 - ds / 2) * 1.5), Ry = H / 2 - ds / 2;
    const d0 = Math.min(M * .36, (Math.min(Rx, Ry) - ds / 2) * 1.5);
    const lines = [], nodes = [];
    sats.forEach((t, i) => {
      const a = -Math.PI / 2 + i * 2 * Math.PI / n; const sx = cx + Rx * Math.cos(a), sy = cy + Ry * Math.sin(a);
      const dx = sx - cx, dy = sy - cy, L0 = Math.hypot(dx, dy) || 1, ux = dx / L0, uy = dy / L0;
      const p0 = [cx + ux * d0 / 2, cy + uy * d0 / 2], p1 = [sx - ux * ds / 2, sy - uy * ds / 2];
      if (Math.hypot(p1[0] - p0[0], p1[1] - p0[1]) > 2) lines.push(S({ kind: 'line', x: p0[0], y: p0[1], w: p1[0] - p0[0], h: p1[1] - p0[1], stroke: SA.mix(c.col(0), .3), strokeW: Math.max(1, M * .006) }));
      nodes.push(S({ kind: 'ellipse', x: sx - ds / 2, y: sy - ds / 2, w: ds, h: ds, fill: c.col(i + 1), text: t.text, role: 'sat', item: t.idx }));
    });
    out.push(...lines, S({ kind: 'ellipse', x: cx - d0 / 2, y: cy - d0 / 2, w: d0, h: d0, fill: c.col(0), text: center.text, role: 'center', bold: true, item: center.idx }), ...nodes);
    return out;
  };
  L.hierarchy = (T0, W, H, c, items) => {
    const roots = tree(items); if (!roots.length) return [];
    let leaves = 0, depth = 0;
    const walk = (n, lv) => { depth = Math.max(depth, lv + 1); n.lv = lv; if (!n.kids.length) { n.slot = leaves++ + .5; } else { n.kids.forEach(k => walk(k, lv + 1)); n.slot = (n.kids[0].slot + n.kids[n.kids.length - 1].slot) / 2; } };
    roots.forEach(r => walk(r, 0));
    const slotW = W / leaves, rowH = H / depth;
    const bw = Math.min(slotW * .86, W * .3), bh = Math.min(rowH * .62, Math.max(bw * .5, rowH * .4));
    const out = [], lines = []; const lc = c.line; const lw = Math.max(1, Math.min(W, H) * .004);
    const pos = (n) => ({ x: n.slot * slotW - bw / 2, y: n.lv * rowH + (rowH - bh) / 2 });
    const visit = (n) => {
      const p = pos(n);
      out.push(S({ kind: 'roundRect', x: p.x, y: p.y, w: bw, h: bh, fill: c.col(n.lv), text: n.text, role: 'lv' + Math.min(n.lv, 2), bold: n.lv === 0, item: n.idx }));
      if (n.kids.length) {
        const pb = p.y + bh, ct = pos(n.kids[0]).y, my = (pb + ct) / 2, px = n.slot * slotW;
        lines.push(S({ kind: 'line', x: px, y: pb, w: 0, h: my - pb, stroke: lc, strokeW: lw }));
        if (n.kids.length > 1) { const a = n.kids[0].slot * slotW, b = n.kids[n.kids.length - 1].slot * slotW; lines.push(S({ kind: 'line', x: a, y: my, w: b - a, h: 0, stroke: lc, strokeW: lw })); }
        n.kids.forEach(k => { const kx = k.slot * slotW; lines.push(S({ kind: 'line', x: kx, y: my, w: 0, h: ct - my, stroke: lc, strokeW: lw })); visit(k); });
      }
    };
    roots.forEach(visit);
    return lines.concat(out);
  };
  L.pyramid = (T, W, H, c) => {
    const N = T.length, gap = N > 1 ? H * .015 : 0, baseW = Math.min(W, H * 1.35), cx = W / 2, out = [];
    const ap = N > 1 ? 1.7 : 1, unit = (H - gap * (N - 1)) / (N - 1 + ap); // the top piece is taller so its text stays readable
    const half = (y) => baseW / 2 * y / H;
    T.forEach((t, i) => {
      const rh = i === 0 ? unit * ap : unit, y0 = i === 0 ? 0 : unit * ap + gap + (i - 1) * (unit + gap), y1 = y0 + rh, hw1 = half(y1), hw0 = half(y0);
      const s = i === 0 ? S({ kind: 'triangle', x: cx - hw1, y: y0, w: hw1 * 2, h: rh }) : S({ kind: 'trapezoid', x: cx - hw1, y: y0, w: hw1 * 2, h: rh, inset: hw1 - hw0 });
      out.push(Object.assign(s, { fill: c.col(i), text: withKids(t), role: i === 0 && N > 1 ? 'apex' : 'main', item: t.kids.length ? undefined : t.idx }));
    });
    return out;
  };
  L.funnel = (T, W, H, c) => {
    const N = T.length, gap = N > 1 ? H * .015 : 0, topW = Math.min(W, H * 1.5), botW = topW * .28, cx = W / 2, rh = (H - gap * (N - 1)) / N, out = [];
    const wAt = (y) => topW - (topW - botW) * y / H;
    T.forEach((t, i) => {
      const y0 = i * (rh + gap), w0 = wAt(y0), w1 = wAt(y0 + rh);
      out.push(S({ kind: 'trapezoid', flipV: true, x: cx - w0 / 2, y: y0, w: w0, h: rh, inset: (w0 - w1) / 2, fill: c.col(i), text: withKids(t), role: 'main', item: t.kids.length ? undefined : t.idx }));
    });
    return out;
  };
  L.venn = (T, W, H, c) => {
    const n = Math.min(4, T.length), out = [], cx = W / 2, cy = H / 2; let d, C;
    if (n === 1) { d = Math.min(W, H) * .8; C = [[cx, cy]]; }
    else if (n === 2) { d = Math.min(H * .95, W / 1.65); C = [[cx - d * .32, cy], [cx + d * .32, cy]]; }
    else if (n === 3) { d = Math.min(H / 1.58, W / 1.7); C = [[cx, cy - d * .27], [cx - d * .31, cy + d * .26], [cx + d * .31, cy + d * .26]]; }
    else { d = Math.min(H, W) / 1.6; C = [[cx - d * .3, cy - d * .3], [cx + d * .3, cy - d * .3], [cx - d * .3, cy + d * .3], [cx + d * .3, cy + d * .3]]; }
    const gx = C.reduce((a, p) => a + p[0], 0) / n, gy = C.reduce((a, p) => a + p[1], 0) / n;
    C.forEach((p, i) => out.push(S({ kind: 'ellipse', x: p[0] - d / 2, y: p[1] - d / 2, w: d, h: d, fill: c.col(i), opacity: .55 })));
    C.forEach((p, i) => {
      let ux = p[0] - gx, uy = p[1] - gy; const l = Math.hypot(ux, uy); if (l > .1) { ux /= l; uy /= l; } else { ux = 0; uy = 0; }
      const lx = p[0] + ux * d * .2, ly = p[1] + uy * d * .2, lw = d * .52, lh = d * .34;
      const t = T[i]; out.push(S({ kind: 'rect', x: lx - lw / 2, y: ly - lh / 2, w: lw, h: lh, fill: 'none', text: withKids(t), textColor: onColor(SA.mix(c.col(i), .45), c.text), bold: true, role: 'label', item: t.kids.length ? undefined : t.idx }));
    });
    return out;
  };
  L.vlist = (T, W, H, c) => {
    const N = T.length, gap = N > 1 ? H * .035 : 0, rh = (H - gap * (N - 1)) / N, kids = T.some(t => t.kids.length), out = [];
    T.forEach((t, i) => {
      const y = i * (rh + gap);
      if (!kids) { out.push(S({ kind: 'roundRect', x: 0, y, w: W, h: rh, fill: c.col(i), text: t.text, role: 'main', item: t.idx })); return; }
      const hw = W * .3;
      out.push(S({ kind: 'rect', x: hw - W * .02, y: y + rh * .08, w: W - hw + W * .02, h: rh * .84, fill: SA.mix(c.col(i), .85), stroke: SA.mix(c.col(i), .35), strokeW: Math.max(.75, H * .002), text: kidsText(t), textColor: c.text, align: 'left', role: 'child', padX: W * .04 }));
      out.push(S({ kind: 'roundRect', x: 0, y, w: hw, h: rh, fill: c.col(i), text: t.text, role: 'main', bold: true, item: t.idx }));
    });
    return out;
  };
  L.matrix = (T, W, H, c) => {
    const center = T.length >= 5 ? T[0] : null, Q = (center ? T.slice(1, 5) : T.slice(0, 4)); const g = Math.min(W, H) * .03, out = [];
    const qw = (W - g) / 2, qh = (H - g) / 2;
    Q.forEach((t, i) => out.push(S({ kind: 'roundRect', x: (i % 2) * (qw + g), y: Math.floor(i / 2) * (qh + g), w: qw, h: qh, fill: c.col(i), text: withKids(t), role: 'main', item: t.kids.length ? undefined : t.idx })));
    if (center) out.push(S({ kind: 'roundRect', x: W * .34, y: H * .36, w: W * .32, h: H * .28, fill: '#ffffff', stroke: c.col(0), strokeW: Math.max(1, H * .006), text: center.text, textColor: c.text, bold: true, role: 'center', item: center.idx }));
    return out;
  };
  L.blocks = (T, W, H, c) => {
    const N = T.length; let best = 1, bs = 1e9;
    for (let cols = 1; cols <= N; cols++) { const rows = Math.ceil(N / cols); const ar = (W / cols) / (H / rows); const sc = Math.abs(Math.log(ar / 1.6)) + (rows * cols - N) * .15; if (sc < bs) { bs = sc; best = cols; } }
    const cols = best, rows = Math.ceil(N / cols), g = Math.min(W, H) * .03, bw = (W - g * (cols - 1)) / cols, bh = (H - g * (rows - 1)) / rows, out = [];
    T.forEach((t, i) => {
      const r = Math.floor(i / cols), inRow = r === rows - 1 ? N - r * cols : cols, off = (cols - inRow) * (bw + g) / 2;
      out.push(S({ kind: 'roundRect', x: off + (i % cols) * (bw + g), y: r * (bh + g), w: bw, h: bh, fill: c.col(i), text: withKids(t), role: 'main', item: t.kids.length ? undefined : t.idx }));
    });
    return out;
  };

  SA.LAYOUTS = [
    { id: 'process', name: 'Basic process', category: 'Process', desc: 'Steps in order, left to right, with arrows.', sample: sample.steps },
    { id: 'chevron', name: 'Chevron process', category: 'Process', desc: 'Steps as arrow shapes. Sub-items appear under each step.', sample: sample.steps },
    { id: 'stepUp', name: 'Step-up process', category: 'Process', desc: 'Steps that climb like stairs — good for progress or levels.', sample: sample.steps },
    { id: 'timeline', name: 'Timeline', category: 'Process', desc: 'Events along a line. Start each line with a year or date.', sample: sample.timeline },
    { id: 'cycle', name: 'Cycle', category: 'Cycle', desc: 'Stages that repeat in a loop.', sample: sample.cycle },
    { id: 'radial', name: 'Radial', category: 'Cycle', desc: 'A central idea (first line) with related ideas around it (sub-items).', sample: sample.radial },
    { id: 'hierarchy', name: 'Hierarchy / org chart', category: 'Hierarchy', desc: 'Levels of a team or topic. Use Tab to put an item under the one above.', sample: sample.tree },
    { id: 'vlist', name: 'Vertical box list', category: 'List', desc: 'Headings with their details (sub-items) beside them.', sample: sample.list },
    { id: 'blocks', name: 'Block list', category: 'List', desc: 'Equal boxes for a group of items.', sample: sample.blocks },
    { id: 'venn', name: 'Venn', category: 'Relationship', desc: 'Two to four overlapping ideas.', sample: sample.venn },
    { id: 'funnel', name: 'Funnel', category: 'Relationship', desc: 'Stages that get smaller, like a filter.', sample: sample.funnel },
    { id: 'matrix', name: 'Matrix 2×2', category: 'Matrix', desc: 'Four quarters. Add a fifth line first to put a title in the middle.', sample: sample.matrix },
    { id: 'pyramid', name: 'Pyramid', category: 'Pyramid', desc: 'Levels from the top (smallest) to the base.', sample: sample.pyramid },
  ];
  SA.CATEGORIES = [...new Set(SA.LAYOUTS.map(l => l.category))];

  /** Lay a diagram out in a w×h box. */
  SA.layout = (id, items, o = {}) => {
    const W = Math.max(10, +o.w || 600), H = Math.max(10, +o.h || 400);
    const pal = (o.palette && o.palette.length ? o.palette : DEF_PAL).map(String);
    const text = o.textColor || '#1f2430';
    const c = { col: (i) => pal[((i % pal.length) + pal.length) % pal.length], text, line: o.lineColor || '#8b95a7' };
    const its = norm(items); const T = tops(its);
    const fn = L[id] || L.process;
    let shapes = T.length ? fn(T, W, H, c, its) : [];
    shapes = shapes.filter(s => s.kind === 'line' || s.kind === 'connector' || (s.w > .5 && s.h > .5));
    // text colour, insets and fitted font size (one size per role, like PowerPoint)
    const M = Math.min(W, H), maxF = o.maxFont || Math.max(6, Math.min(M * .085, 40 * M / 470)), minF = o.minFont || Math.max(3, maxF * .22);
    const sizes = {};
    shapes.forEach(s => {
      if (s.kind === 'line' || s.kind === 'connector' || !s.text) return;
      if (!s.textColor) s.textColor = s.fill === 'none' ? text : onColor(s.fill, text);
      const gi = geoInset(s); const p = Math.max(1, Math.min(s.w, s.h) * .06);
      s.pad = [gi[0] * s.w + p + (s.padX || 0), gi[1] * s.h + p, gi[2] * s.w + p, gi[3] * s.h + p];
      const tw = s.w - s.pad[0] - s.pad[2], th = (s.fitH || s.h) - s.pad[1] - s.pad[3];
      const role = s.role || 'main', mx = role === 'child' || role === 'label' ? maxF * .8 : role === 'center' ? maxF * 1.1 : maxF;
      const f = fitSize(s.text, tw * .97, th * .97, o.font, s.bold, mx, minF);
      sizes[role] = Math.min(sizes[role] ?? 1e9, f);
    });
    shapes.forEach(s => { if (s.text && s.pad) s.fontSize = f2(sizes[s.role || 'main']); delete s.padX; delete s.fitH; });
    return { id, w: W, h: H, shapes };
  };

  /* ---------- SVG ---------- */
  SA.toSVG = (model, o = {}) => {
    const font = o.font || 'Inter'; let defs = '', body = '';
    if (o.background) body += `<rect width="${f2(model.w)}" height="${f2(model.h)}" fill="${esc(o.background)}"/>`;
    model.shapes.forEach((s, i) => {
      const op = s.opacity != null && s.opacity < 1 ? ` opacity="${s.opacity}"` : '';
      if (s.kind === 'line' || s.kind === 'connector') {
        let mk = ''; if (s.kind === 'connector') { defs += `<marker id="sa-ah${i}" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="5" markerHeight="5" orient="auto"><path d="M0 0 L10 5 L0 10 z" fill="${esc(s.stroke)}"/></marker>`; mk = ` marker-end="url(#sa-ah${i})"`; }
        body += `<line x1="${f2(s.x)}" y1="${f2(s.y)}" x2="${f2(s.x + s.w)}" y2="${f2(s.y + s.h)}" stroke="${esc(s.stroke || '#888')}" stroke-width="${f2(s.strokeW || 1)}"${mk}${op}/>`;
        return;
      }
      const tr = s.rot ? ` transform="rotate(${s.rot} ${f2(s.x + s.w / 2)} ${f2(s.y + s.h / 2)})"` : '';
      const st = s.stroke && s.stroke !== 'none' && s.strokeW ? ` stroke="${esc(s.stroke)}" stroke-width="${f2(s.strokeW)}"` : '';
      body += `<path d="${SA.shapePath(s)}" fill="${esc(s.fill || 'none')}"${st}${op}${tr}/>`;
      if (s.text && s.fontSize) {
        const p = s.pad || [0, 0, 0, 0], tw = s.w - p[0] - p[2], fs = s.fontSize;
        const lines = wrap(s.text, tw, fs, font, s.bold); const lh = fs * LH, th = lines.length * lh;
        const top = s.valign === 'top' ? s.y + p[1] : s.valign === 'bottom' ? s.y + s.h - p[3] - th : s.y + p[1] + (s.h - p[1] - p[3] - th) / 2;
        const anchor = s.align === 'left' ? 'start' : s.align === 'right' ? 'end' : 'middle';
        const tx = s.align === 'left' ? s.x + p[0] : s.align === 'right' ? s.x + s.w - p[2] : s.x + p[0] + tw / 2;
        body += `<text font-family="${esc(fontCSS(font))}" font-size="${f2(fs)}" fill="${esc(s.textColor || '#000')}" text-anchor="${anchor}"${s.bold ? ' font-weight="700"' : ''}${tr}>${lines.map((l, k) => `<tspan x="${f2(tx)}" y="${f2(top + k * lh + fs * .95)}">${esc(l)}</tspan>`).join('')}</text>`;
      }
    });
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${f2(model.w)}" height="${f2(model.h)}" viewBox="0 0 ${f2(model.w)} ${f2(model.h)}">${defs ? '<defs>' + defs + '</defs>' : ''}${body}</svg>`;
  };

  // gallery thumbnails: the layout drawn without text
  SA.LAYOUTS.forEach(l => {
    const m = SA.layout(l.id, l.sample.map(it => ({ text: 'x', level: it.level })), { w: 96, h: 60, palette: ['#4472C4', '#5B9BD5', '#2F5597', '#8FAADC', '#1F4E79', '#9DC3E6'] });
    m.shapes.forEach(s => { s.text = ''; });
    l.thumbSVG = SA.toSVG(m);
  });

  /* ---------- outline text ⇄ items ---------- */
  SA.parseOutline = (txt) => String(txt || '').replace(/\r/g, '').split('\n').map(line => {
    const m = /^([\t ]*)(.*)$/.exec(line); const ind = m[1].replace(/ {4}/g, '\t').replace(/ +/g, ''); return { text: m[2].trim(), level: ind.length };
  }).filter(it => it.text);
  SA.toOutline = (items) => norm(items).map(it => '\t'.repeat(it.level) + it.text).join('\n');

  /* ---------- dialog ---------- */
  const CSS = `
.sa-dlg{display:grid;grid-template-columns:208px minmax(200px,1fr) 270px;gap:14px;height:440px;min-width:0}
.sa-gal{overflow:auto;border:1px solid var(--line,#d8dce6);border-radius:8px;padding:6px 8px;background:var(--surface,#fff)}
.sa-cat{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:var(--ink-2,#5b6475);margin:8px 2px 5px}
.sa-cat:first-child{margin-top:2px}
.sa-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px}
.sa-it{display:flex;flex-direction:column;align-items:center;gap:3px;padding:6px 4px 5px;border:1px solid transparent;border-radius:7px;background:none;color:inherit;font:inherit;font-size:11px;line-height:1.2;text-align:center;cursor:default}
.sa-it svg{width:76px;height:48px;display:block}
.sa-it:hover{background:var(--accent-soft,#e8ecff)}
.sa-it.on{border-color:var(--accent,#3c4fe0);background:var(--accent-soft,#e8ecff)}
.sa-mid{display:flex;flex-direction:column;gap:6px;min-width:0}
.sa-mid label{font-weight:600;font-size:12.5px}
.sa-outline{flex:1;resize:none;font:13.5px/1.6 var(--font-ui,inherit);tab-size:3;-moz-tab-size:3;padding:8px 10px;border:1px solid var(--line,#d8dce6);border-radius:8px;background:var(--surface,#fff);color:var(--ink,#1f2430);white-space:pre;overflow:auto}
.sa-outline:focus{outline:2px solid var(--accent,#3c4fe0);outline-offset:-1px}
.sa-tools{display:flex;gap:6px;align-items:center}
.sa-tools button{white-space:nowrap;height:26px;padding:0 9px;border:1px solid var(--line,#d8dce6);border-radius:6px;background:var(--surface,#fff);color:inherit;font:inherit;font-size:12px}
.sa-tools button:hover{border-color:var(--accent,#3c4fe0)}
.sa-hint{font-size:11.5px;color:var(--ink-2,#5b6475);line-height:1.4}
.sa-side{display:flex;flex-direction:column;gap:8px;min-width:0}
.sa-prev{border:1px solid var(--line,#d8dce6);border-radius:8px;background:#fff;height:230px;display:flex;align-items:center;justify-content:center;padding:10px}
.sa-prev svg{max-width:100%;max-height:100%}
.sa-name{font-weight:700}
.sa-desc{font-size:12px;color:var(--ink-2,#5b6475);line-height:1.45}`;
  function injectCSS() { if (typeof document === 'undefined' || document.getElementById('sa-css')) return; const st = document.createElement('style'); st.id = 'sa-css'; st.textContent = CSS; document.head.appendChild(st); }

  SA.dialog = (opts = {}) => {
    const B = root.BTS; if (!B || !B.dialog) return Promise.resolve(null);
    injectCSS();
    const init = opts.initial || {};
    let cur = SA.LAYOUTS.some(l => l.id === init.id) ? init.id : 'process';
    const lay = () => SA.LAYOUTS.find(l => l.id === cur);
    let touched = !!(init.items && init.items.length);
    const body = document.createElement('div'); body.className = 'sa-dlg modal-noenter';
    body.innerHTML = `<div class="sa-gal">${SA.CATEGORIES.map(cat => `<div class="sa-cat">${esc(cat)}</div><div class="sa-grid">${SA.LAYOUTS.filter(l => l.category === cat).map(l => `<button class="sa-it" data-id="${l.id}" title="${esc(l.desc)}">${l.thumbSVG}<span>${esc(l.name)}</span></button>`).join('')}</div>`).join('')}</div>
      <div class="sa-mid"><label for="sa-text">Text — one line per item</label><textarea id="sa-text" class="sa-outline" spellcheck="true"></textarea>
        <div class="sa-tools"><button data-a="out" title="Move the line up a level (Shift+Tab)">Promote</button><button data-a="in" title="Make the line a sub-item (Tab)">Demote</button><span style="flex:1"></span><button data-a="sample" title="Fill in example text for this layout">Example</button></div>
        <div class="sa-hint">Press <b>Tab</b> to make a line a sub-item of the line above, <b>Shift+Tab</b> to move it back.</div></div>
      <div class="sa-side"><div class="sa-name"></div><div class="sa-prev"></div><div class="sa-desc"></div></div>`;
    const ta = body.querySelector('#sa-text');
    ta.value = SA.toOutline(touched ? init.items : lay().sample);
    const draw = () => {
      body.querySelectorAll('.sa-it').forEach(b => b.classList.toggle('on', b.dataset.id === cur));
      body.querySelector('.sa-name').textContent = lay().name; body.querySelector('.sa-desc').textContent = lay().desc;
      const m = SA.layout(cur, SA.parseOutline(ta.value), { w: 400, h: 280, palette: opts.palette, font: opts.font, textColor: opts.textColor });
      body.querySelector('.sa-prev').innerHTML = SA.toSVG(m, { font: opts.font });
    };
    const shift = (d) => {
      const v = ta.value, s = ta.selectionStart, e = ta.selectionEnd;
      const ls = v.lastIndexOf('\n', s - 1) + 1; let le = v.indexOf('\n', Math.max(s, e - (e > s && v[e - 1] === '\n' ? 1 : 0))); if (le < 0) le = v.length;
      const block = v.slice(ls, le).split('\n').map(l => d > 0 ? '\t' + l : l.replace(/^(\t| {1,4})/, '')).join('\n');
      ta.value = v.slice(0, ls) + block + v.slice(le);
      ta.selectionStart = ls; ta.selectionEnd = ls + block.length; if (s === e) { const pos = Math.max(ls, s + (block.length - (le - ls))); ta.selectionStart = ta.selectionEnd = pos; }
      touched = true; draw();
    };
    ta.addEventListener('keydown', e => { if (e.key === 'Tab') { e.preventDefault(); e.stopPropagation(); shift(e.shiftKey ? -1 : 1); } });
    ta.addEventListener('input', () => { touched = true; draw(); });
    body.addEventListener('click', e => {
      const b = e.target.closest('.sa-it'); if (b) { cur = b.dataset.id; if (!touched) ta.value = SA.toOutline(lay().sample); draw(); return; }
      const a = e.target.closest('[data-a]'); if (!a) return;
      if (a.dataset.a === 'in' || a.dataset.a === 'out') { ta.focus(); shift(a.dataset.a === 'in' ? 1 : -1); }
      if (a.dataset.a === 'sample') { ta.value = SA.toOutline(lay().sample); touched = false; draw(); ta.focus(); }
    });
    body.addEventListener('dblclick', e => { if (e.target.closest('.sa-it')) body.closest('.modal')?.querySelector('footer .primary')?.click(); });
    return B.dialog({
      title: opts.title || (init.id ? 'Change SmartArt' : 'Insert SmartArt'), width: 900, body,
      buttons: [{ label: 'Cancel', value: null }, { label: opts.okLabel || (init.id ? 'Update' : 'Insert'), primary: true, value: () => ({ id: cur, items: SA.parseOutline(ta.value) }), validate: () => { if (!SA.parseOutline(ta.value).length) { ta.focus(); return false; } return true; } }],
      onOpen: (r) => { const m = r.querySelector('.modal'); if (m) m.style.maxWidth = 'min(920px, 95vw)'; draw(); setTimeout(() => { ta.focus(); ta.selectionStart = ta.selectionEnd = ta.value.length; }, 0); },
    });
  };

  root.BTSSmartArt = SA;
})(typeof window !== 'undefined' ? window : globalThis);
