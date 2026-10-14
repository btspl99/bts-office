/* BTS Slides — themes, colour tokens, fonts and slide layouts.
 * A theme = background + fonts + palette + decoration (SVG drawn behind content).
 * Element colours may be hex ("#rrggbb", "#rrggbbaa") or theme tokens:
 *   bg, bg2, text, title, accent1 … accent6  — resolved by SL.color(token, pres).
 */
(function () {
  const SL = window.SL = window.SL || {};

  SL.FONTS = ['Inter', 'Poppins', 'Montserrat', 'Lora', 'Playfair Display', 'Figtree', 'Arial', 'Calibri', 'Cambria', 'Georgia',
    'Times New Roman', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Segoe UI', 'Courier New', 'Comic Sans MS', 'Nirmala UI', 'Mangal'];
  SL.SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];

  // deco(kind, W, H, c) → SVG inner markup. kind: 'title' | 'section' | 'content'
  const T = (o) => o;
  SL.THEMES = [
    T({ id: 'clean', name: 'Clean', titleFont: 'Inter', bodyFont: 'Inter', titleBold: true,
      colors: { bg: '#ffffff', bg2: '#f4f5f7', text: '#2b3140', title: '#121722', accents: ['#D9661A', '#2F6FDE', '#1C9C74', '#E0A526', '#8A4FD3', '#D0384E'] },
      variants: [{}, { accents: ['#2F6FDE', '#D9661A', '#1C9C74', '#E0A526', '#8A4FD3', '#D0384E'] }, { accents: ['#1C9C74', '#2F6FDE', '#D9661A', '#E0A526', '#8A4FD3', '#D0384E'] }, { accents: ['#8A4FD3', '#D0384E', '#2F6FDE', '#E0A526', '#1C9C74', '#D9661A'] }],
      deco: (k, W, H, c) => k === 'title' ? `<rect x="0" y="${H - 14}" width="${W}" height="14" fill="${c.accents[0]}"/><rect x="${W / 2 - 60}" y="${H * .538}" width="120" height="6" rx="3" fill="${c.accents[0]}"/>`
        : k === 'section' ? `<rect x="0" y="0" width="22" height="${H}" fill="${c.accents[0]}"/>` : `<rect x="0" y="0" width="${W}" height="8" fill="${c.accents[0]}"/>` }),
    T({ id: 'midnight', name: 'Midnight', titleFont: 'Poppins', bodyFont: 'Inter', titleBold: true,
      bg: { type: 'gradient', color: '#0B1222', color2: '#18233F', angle: 135 },
      colors: { bg: '#0F172A', bg2: '#1E293B', text: '#CBD5E1', title: '#FFFFFF', accents: ['#38BDF8', '#A78BFA', '#F472B6', '#34D399', '#FBBF24', '#F87171'] },
      variants: [{}, { accents: ['#A78BFA', '#38BDF8', '#F472B6', '#34D399', '#FBBF24', '#F87171'] }, { accents: ['#34D399', '#38BDF8', '#A78BFA', '#F472B6', '#FBBF24', '#F87171'] }, { accents: ['#FBBF24', '#F87171', '#38BDF8', '#A78BFA', '#34D399', '#F472B6'] }],
      deco: (k, W, H, c) => `<defs><radialGradient id="mg1"><stop offset="0" stop-color="${c.accents[0]}" stop-opacity=".55"/><stop offset="1" stop-color="${c.accents[0]}" stop-opacity="0"/></radialGradient><radialGradient id="mg2"><stop offset="0" stop-color="${c.accents[1]}" stop-opacity=".45"/><stop offset="1" stop-color="${c.accents[1]}" stop-opacity="0"/></radialGradient></defs>` +
        (k === 'content' ? `<circle cx="${W}" cy="0" r="${H * .45}" fill="url(#mg1)"/>` : `<circle cx="${W * .9}" cy="${H * .1}" r="${H * .7}" fill="url(#mg1)"/><circle cx="${W * .05}" cy="${H}" r="${H * .6}" fill="url(#mg2)"/>`) }),
    T({ id: 'sunrise', titleArea: { x: .0625, w: .54 }, name: 'Sunrise', titleFont: 'Montserrat', bodyFont: 'Inter', titleBold: true,
      bg: { type: 'gradient', color: '#FFF4EA', color2: '#FFE3E8', angle: 160 },
      colors: { bg: '#FFF4EA', bg2: '#FFFFFF', text: '#4A3540', title: '#2A1622', accents: ['#F06A4B', '#E8467C', '#F2A93B', '#7A5AF0', '#2BA39A', '#3B82F6'] },
      variants: [{}, { accents: ['#E8467C', '#F06A4B', '#F2A93B', '#7A5AF0', '#2BA39A', '#3B82F6'] }, { accents: ['#7A5AF0', '#E8467C', '#F06A4B', '#F2A93B', '#2BA39A', '#3B82F6'] }, { accents: ['#2BA39A', '#F2A93B', '#F06A4B', '#E8467C', '#7A5AF0', '#3B82F6'] }],
      deco: (k, W, H, c) => k === 'content' ? `<circle cx="${W - 40}" cy="${H - 30}" r="90" fill="${c.accents[0]}" opacity=".12"/><circle cx="${W - 110}" cy="${H + 10}" r="60" fill="${c.accents[1]}" opacity=".12"/>`
        : `<circle cx="${W * .86}" cy="${H * .74}" r="${H * .42}" fill="${c.accents[0]}" opacity=".9"/><circle cx="${W * .76}" cy="${H * .96}" r="${H * .3}" fill="${c.accents[1]}" opacity=".8"/><circle cx="${W * .95}" cy="${H * .3}" r="${H * .1}" fill="${c.accents[2]}" opacity=".9"/>` }),
    T({ id: 'forest', name: 'Forest', titleFont: 'Lora', bodyFont: 'Inter', titleBold: true,
      colors: { bg: '#0F3B2E', bg2: '#16503F', text: '#D7E8DF', title: '#F4F1E6', accents: ['#B7E36B', '#F2C14E', '#6CC4A1', '#F08A5D', '#9AD1F0', '#E6E6E6'] },
      variants: [{}, { accents: ['#F2C14E', '#B7E36B', '#6CC4A1', '#F08A5D', '#9AD1F0', '#E6E6E6'] }, { bg: '#1F2A44', bg2: '#2A3858', accents: ['#9AD1F0', '#F2C14E', '#B7E36B', '#F08A5D', '#6CC4A1', '#E6E6E6'] }, { bg: '#3B1F2B', bg2: '#4E2A3A', accents: ['#F08A5D', '#F2C14E', '#B7E36B', '#9AD1F0', '#6CC4A1', '#E6E6E6'] }],
      deco: (k, W, H, c) => k === 'content' ? `<path d="M${W - 160} ${H} Q${W - 60} ${H - 120} ${W} ${H - 150} V${H} Z" fill="${c.accents[0]}" opacity=".18"/>`
        : `<path d="M0 ${H} Q${W * .3} ${H * .72} ${W * .55} ${H * .86} T${W} ${H * .7} V${H} Z" fill="${c.accents[0]}" opacity=".25"/><path d="M0 ${H} Q${W * .4} ${H * .82} ${W * .7} ${H * .93} T${W} ${H * .85} V${H} Z" fill="${c.accents[0]}" opacity=".35"/>` }),
    T({ id: 'paper', name: 'Paper', titleFont: 'Playfair Display', bodyFont: 'Lora', titleBold: true,
      colors: { bg: '#FBF7EF', bg2: '#F2EBDD', text: '#3D3529', title: '#1E1A14', accents: ['#9E2A2B', '#335C67', '#E09F3E', '#540B0E', '#6B8F71', '#7F5539'] },
      variants: [{}, { accents: ['#335C67', '#9E2A2B', '#E09F3E', '#540B0E', '#6B8F71', '#7F5539'] }, { accents: ['#6B8F71', '#9E2A2B', '#335C67', '#E09F3E', '#540B0E', '#7F5539'] }, { accents: ['#E09F3E', '#335C67', '#9E2A2B', '#540B0E', '#6B8F71', '#7F5539'] }],
      deco: (k, W, H, c) => k === 'title' ? `<rect x="40" y="40" width="${W - 80}" height="${H - 80}" fill="none" stroke="${c.accents[0]}" stroke-width="2"/><rect x="50" y="50" width="${W - 100}" height="${H - 100}" fill="none" stroke="${c.accents[0]}" stroke-width=".8"/>`
        : k === 'section' ? `<line x1="80" y1="${H * .5}" x2="${W - 80}" y2="${H * .5}" stroke="${c.accents[0]}" stroke-width="1.5"/>` : `<line x1="60" y1="${H - 40}" x2="${W - 60}" y2="${H - 40}" stroke="${c.accents[0]}" stroke-width="1"/>` }),
    T({ id: 'ocean', name: 'Ocean', titleFont: 'Poppins', bodyFont: 'Poppins', titleBold: true,
      bg: { type: 'gradient', color: '#1E3A8A', color2: '#0EA5E9', angle: 135 },
      colors: { bg: '#1E40AF', bg2: '#1D4ED8', text: '#E0F2FE', title: '#FFFFFF', accents: ['#FDE68A', '#5EEAD4', '#FCA5A5', '#C4B5FD', '#86EFAC', '#FFFFFF'] },
      variants: [{}, { bg: '#5B21B6', accents: ['#5EEAD4', '#FDE68A', '#FCA5A5', '#C4B5FD', '#86EFAC', '#FFFFFF'] }, { bg: '#065F46', accents: ['#FDE68A', '#FCA5A5', '#5EEAD4', '#C4B5FD', '#86EFAC', '#FFFFFF'] }, { bg: '#9F1239', accents: ['#FDE68A', '#5EEAD4', '#C4B5FD', '#FCA5A5', '#86EFAC', '#FFFFFF'] }],
      bgFor: (v) => v.bg ? { type: 'gradient', color: v.bg, color2: v.bg2 || '#0EA5E9', angle: 135 } : null,
      deco: (k, W, H, c) => `<path d="M0 ${H * .82} C${W * .25} ${H * .74} ${W * .5} ${H * .92} ${W} ${H * .8} V${H} H0 Z" fill="#fff" opacity=".08"/><path d="M0 ${H * .9} C${W * .3} ${H * .82} ${W * .6} ${H * .98} ${W} ${H * .88} V${H} H0 Z" fill="#fff" opacity=".1"/>` }),
    T({ id: 'geo', titleArea: { x: .0625, w: .6 }, name: 'Geometric', titleFont: 'Montserrat', bodyFont: 'Montserrat', titleBold: true,
      colors: { bg: '#FFFFFF', bg2: '#F3F4F8', text: '#30364A', title: '#161A2B', accents: ['#5A4FCF', '#FF6B6B', '#FFC145', '#21B6A8', '#3A86FF', '#FF8FAB'] },
      variants: [{}, { accents: ['#21B6A8', '#FFC145', '#FF6B6B', '#5A4FCF', '#3A86FF', '#FF8FAB'] }, { accents: ['#FF6B6B', '#5A4FCF', '#FFC145', '#21B6A8', '#3A86FF', '#FF8FAB'] }, { accents: ['#3A86FF', '#FF8FAB', '#FFC145', '#21B6A8', '#5A4FCF', '#FF6B6B'] }],
      deco: (k, W, H, c) => k === 'content' ? `<polygon points="${W},0 ${W},90 ${W - 90},0" fill="${c.accents[0]}"/><polygon points="${W - 60},0 ${W - 120},0 ${W - 90},40" fill="${c.accents[2]}"/>`
        : `<polygon points="${W},0 ${W},${H * .5} ${W * .8},0" fill="${c.accents[0]}"/><polygon points="${W},${H * .42} ${W},${H} ${W * .84},${H}" fill="${c.accents[1]}"/><polygon points="${W * .7},0 ${W * .82},0 ${W * .76},${H * .16}" fill="${c.accents[2]}"/><circle cx="${W * .93}" cy="${H * .56}" r="30" fill="${c.accents[3]}"/><rect x="0" y="${H - 14}" width="${W * .35}" height="14" fill="${c.accents[0]}"/>` }),
    T({ id: 'block', titleArea: { x: .0625, w: .55 }, name: 'Bold Block', titleFont: 'Poppins', bodyFont: 'Inter', titleBold: true,
      colors: { bg: '#FFFFFF', bg2: '#F5F3EE', text: '#2E2E2E', title: '#111111', accents: ['#FF5A36', '#111111', '#2563EB', '#16A34A', '#EAB308', '#9333EA'] },
      variants: [{}, { accents: ['#2563EB', '#111111', '#FF5A36', '#16A34A', '#EAB308', '#9333EA'] }, { accents: ['#16A34A', '#111111', '#2563EB', '#FF5A36', '#EAB308', '#9333EA'] }, { accents: ['#9333EA', '#111111', '#2563EB', '#16A34A', '#EAB308', '#FF5A36'] }],
      deco: (k, W, H, c) => k === 'title' ? `<rect x="${W * .64}" y="0" width="${W * .36}" height="${H}" fill="${c.accents[0]}"/><rect x="${W * .64 - 26}" y="${H * .38}" width="52" height="52" fill="${c.accents[1]}"/>`
        : k === 'section' ? `<rect x="0" y="0" width="${W}" height="${H}" fill="${c.accents[0]}" opacity=".08"/><rect x="0" y="0" width="${W * .03}" height="${H}" fill="${c.accents[0]}"/>` : `<rect x="0" y="0" width="12" height="${H}" fill="${c.accents[0]}"/>` }),
    T({ id: 'saffron', titleArea: { x: .086, w: .7 }, name: 'Saffron', titleFont: 'Poppins', bodyFont: 'Inter', titleBold: true,
      colors: { bg: '#FFFBF3', bg2: '#FFF1DB', text: '#3A3226', title: '#1F1A12', accents: ['#F28C28', '#128807', '#0B3D91', '#C2410C', '#B45309', '#7C3AED'] },
      variants: [{}, { accents: ['#128807', '#F28C28', '#0B3D91', '#C2410C', '#B45309', '#7C3AED'] }, { accents: ['#0B3D91', '#F28C28', '#128807', '#C2410C', '#B45309', '#7C3AED'] }, { accents: ['#C2410C', '#B45309', '#128807', '#0B3D91', '#F28C28', '#7C3AED'] }],
      deco: (k, W, H, c) => k === 'content' ? `<rect x="0" y="0" width="${W}" height="6" fill="${c.accents[0]}"/><rect x="0" y="${H - 6}" width="${W}" height="6" fill="${c.accents[1]}"/>`
        : `<rect x="0" y="0" width="${W}" height="18" fill="${c.accents[0]}"/><rect x="0" y="${H - 18}" width="${W}" height="18" fill="${c.accents[1]}"/><g transform="translate(${W - 140} ${H / 2})" fill="none" stroke="${c.accents[2]}" stroke-width="2" opacity=".55"><circle r="70"/><circle r="8" fill="${c.accents[2]}"/>${Array.from({ length: 24 }, (_, i) => `<line x1="0" y1="0" x2="${(70 * Math.cos(i * Math.PI / 12)).toFixed(1)}" y2="${(70 * Math.sin(i * Math.PI / 12)).toFixed(1)}"/>`).join('')}</g>` }),
    T({ id: 'slate', titleArea: { x: .0625, w: .46 }, name: 'Slate', titleFont: 'Inter', bodyFont: 'Inter', titleBold: true,
      colors: { bg: '#F1F5F9', bg2: '#FFFFFF', text: '#334155', title: '#0F172A', accents: ['#0EA5E9', '#6366F1', '#F43F5E', '#10B981', '#F59E0B', '#64748B'] },
      variants: [{}, { accents: ['#6366F1', '#0EA5E9', '#F43F5E', '#10B981', '#F59E0B', '#64748B'] }, { accents: ['#F43F5E', '#6366F1', '#0EA5E9', '#10B981', '#F59E0B', '#64748B'] }, { accents: ['#10B981', '#0EA5E9', '#6366F1', '#F43F5E', '#F59E0B', '#64748B'] }],
      deco: (k, W, H, c) => k === 'content' ? `<rect x="40" y="44" width="6" height="64" rx="3" fill="${c.accents[0]}"/>`
        : `<rect x="${W * .55}" y="${H * .1}" width="${W * .4}" height="${H * .8}" rx="28" fill="${c.bg2}"/><rect x="${W * .6}" y="${H * .2}" width="${W * .3}" height="12" rx="6" fill="${c.accents[0]}"/><rect x="${W * .6}" y="${H * .2 + 30}" width="${W * .2}" height="12" rx="6" fill="${c.accents[1]}" opacity=".7"/><rect x="${W * .6}" y="${H * .2 + 60}" width="${W * .25}" height="12" rx="6" fill="${c.accents[2]}" opacity=".5"/>` }),
    T({ id: 'neon', name: 'Neon', titleFont: 'Montserrat', bodyFont: 'Inter', titleBold: true,
      colors: { bg: '#0A0A0F', bg2: '#17171F', text: '#D4D4E0', title: '#FFFFFF', accents: ['#FF2E88', '#00E5FF', '#B6FF00', '#FFB800', '#8B5CF6', '#FF5C00'] },
      variants: [{}, { accents: ['#00E5FF', '#FF2E88', '#B6FF00', '#FFB800', '#8B5CF6', '#FF5C00'] }, { accents: ['#B6FF00', '#00E5FF', '#FF2E88', '#FFB800', '#8B5CF6', '#FF5C00'] }, { accents: ['#FFB800', '#FF2E88', '#00E5FF', '#B6FF00', '#8B5CF6', '#FF5C00'] }],
      deco: (k, W, H, c) => `<defs><linearGradient id="ng" x1="0" x2="1"><stop offset="0" stop-color="${c.accents[0]}"/><stop offset="1" stop-color="${c.accents[1]}"/></linearGradient></defs>` +
        (k === 'content' ? `<rect x="0" y="${H - 5}" width="${W}" height="5" fill="url(#ng)"/>` : `<rect x="0" y="${H - 10}" width="${W}" height="10" fill="url(#ng)"/>${Array.from({ length: 6 }, (_, i) => `<line x1="${W * .55 + i * 60}" y1="${H}" x2="${W * .8 + i * 60}" y2="0" stroke="url(#ng)" stroke-width="2" opacity="${.15 + i * .1}"/>`).join('')}`) }),
  ];
  SL.themeById = (id) => SL.THEMES.find(t => t.id === id) || SL.THEMES[0];
  /* Theme object for a presentation; imported files carry their own palette/fonts (theme.custom). */
  SL.themeOf = (pres) => {
    const t = pres.theme || {};
    if (t.id === 'custom' && t.custom) return Object.assign({ id: 'custom', name: t.custom.name || 'Imported theme', titleBold: false, variants: [{}], deco: null }, t.custom);
    return SL.themeById(t.id);
  };

  /* Effective palette for the presentation theme (+variant). */
  SL.palette = (pres) => {
    const th = SL.themeOf(pres);
    const v = th.variants?.[pres.theme?.variant || 0] || {};
    const c = Object.assign({}, th.colors, v);
    c.accents = v.accents || th.colors.accents;
    return c;
  };
  SL.themeBg = (pres) => {
    const th = SL.themeOf(pres); const v = th.variants?.[pres.theme?.variant || 0] || {}; const c = SL.palette(pres);
    if (th.bgFor) { const b = th.bgFor(v); if (b) return b; }
    if (v.bg) return { type: 'solid', color: v.bg };
    return th.bg || { type: 'solid', color: c.bg };
  };
  SL.color = (v, pres) => {
    if (!v || v === 'none') return v;
    if (typeof v === 'object') return v.color ? SL.color(v.color, pres) : null;
    if (v[0] === '#' || v.startsWith('rgb')) return v;
    const c = SL.palette(pres);
    const m = /^accent([1-6])$/.exec(v); if (m) return c.accents[+m[1] - 1];
    return c[v] || v;
  };
  SL.fontOf = (role, pres) => { const th = SL.themeOf(pres); return role === 'title' ? th.titleFont : th.bodyFont; };

  /* ---------- Layouts (positions for a W×H slide) ---------- */
  SL.LAYOUTS = [
    { id: 'title', name: 'Title Slide' },
    { id: 'titleContent', name: 'Title and Content' },
    { id: 'section', name: 'Section Header' },
    { id: 'twoContent', name: 'Two Content' },
    { id: 'comparison', name: 'Comparison' },
    { id: 'titleOnly', name: 'Title Only' },
    { id: 'blank', name: 'Blank' },
    { id: 'pictureCaption', name: 'Picture with Caption' },
  ];
  // Returns element specs (without ids). Sizes in pt, geometry in logical px.
  SL.layoutPlaceholders = (layout, W, H, pres) => {
    const ta = layout === 'title' && pres ? SL.themeOf(pres).titleArea : null;
    const sx = W / 1280, sy = H / 720, R = (x, y, w, h) => ({ x: Math.round(x * sx), y: Math.round(y * sy), w: Math.round(w * sx), h: Math.round(h * sy) });
    const title = (g, extra) => Object.assign({ type: 'text', ph: 'title', role: 'title', size: 40, valign: 'middle', placeholder: 'Click to add title' }, g, extra);
    const body = (g, extra) => Object.assign({ type: 'text', ph: 'body', role: 'body', size: 24, valign: 'top', pspace: 8, list: 'ul', placeholder: 'Click to add text' }, g, extra);
    switch (layout) {
      case 'title': {
        if (ta) { const x = Math.round(ta.x * W), w = Math.round(ta.w * W); // left-aligned in the space the theme graphics leave free
          return [title(Object.assign(R(0, 150, 0, 230), { x, w }), { size: 54, align: 'left', valign: 'bottom' }), { type: 'text', ph: 'subtitle', role: 'body', size: 24, align: 'left', valign: 'top', placeholder: 'Click to add subtitle', ...Object.assign(R(0, 400, 0, 130), { x, w }) }]; }
        return [title(R(110, 190, 1060, 190), { size: 54, align: 'center', valign: 'bottom' }), { type: 'text', ph: 'subtitle', role: 'body', size: 24, align: 'center', valign: 'top', placeholder: 'Click to add subtitle', ...R(110, 400, 1060, 110) }];
      }
      case 'titleContent': return [title(R(80, 40, 1120, 110)), body(R(80, 170, 1120, 490))];
      case 'section': return [title(R(110, 210, 1060, 180), { size: 48, valign: 'bottom' }), { type: 'text', ph: 'body', role: 'body', size: 22, valign: 'top', placeholder: 'Click to add text', ...R(110, 410, 1060, 90) }];
      case 'twoContent': return [title(R(80, 40, 1120, 110)), body(R(80, 170, 545, 490), { size: 22 }), body(R(655, 170, 545, 490), { size: 22 })];
      case 'comparison': return [title(R(80, 40, 1120, 100)),
        { type: 'text', ph: 'body', role: 'body', size: 24, bold: true, valign: 'bottom', placeholder: 'Click to add text', ...R(80, 160, 545, 60) }, body(R(80, 230, 545, 430), { size: 20 }),
        { type: 'text', ph: 'body', role: 'body', size: 24, bold: true, valign: 'bottom', placeholder: 'Click to add text', ...R(655, 160, 545, 60) }, body(R(655, 230, 545, 430), { size: 20 })];
      case 'titleOnly': return [title(R(80, 40, 1120, 110))];
      case 'pictureCaption': return [title(R(80, 90, 440, 160), { size: 32, valign: 'bottom' }), { type: 'image', ph: 'pic', ...R(580, 70, 620, 580) }, { type: 'text', ph: 'body', role: 'body', size: 18, valign: 'top', placeholder: 'Click to add text', ...R(80, 270, 440, 380) }];
      default: return [];
    }
  };
  SL.decoKind = (layout) => layout === 'title' ? 'title' : layout === 'section' ? 'section' : 'content';
})();
