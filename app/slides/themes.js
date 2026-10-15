/* BTS Slides — themes, colour tokens, fonts and slide layouts.
 * A theme = background + fonts + palette + decoration (SVG drawn behind content) + 4 colour variants.
 * Element colours may be hex ("#rrggbb", "#rrggbbaa") or theme tokens:
 *   bg, bg2, text, title, accent1 … accent6  — resolved by SL.color(token, pres).
 *   A token may carry a modifier: "accent1/l80" (80% lighter) or "accent1/d35" (35% darker).
 * Presentation-level overrides (saved in the file): pres.theme.fonts {title, body}, pres.theme.colors {bg,bg2,text,title,accents[6]},
 * pres.master {bg, layoutBg{}, title{font,size,color,bold,italic}, body{…}, layouts{layout:{key:{x,y,w,h,size}}}, graphics[], layoutGraphics{}}.
 *
 * Decorations must never sit under a placeholder: themes that draw on the left/right of the title slide declare
 * `areas` so the placeholders move into the free space (see SL.layoutPlaceholders). Elements with class="wash"
 * are subtle full-slide textures and may sit behind text (they keep contrast).
 */
(function () {
  const SL = window.SL = window.SL || {};

  SL.FONTS = (window.BTS_FONTS || []).map(f => f.name);
  if (!SL.FONTS.length) SL.FONTS = ['Inter', 'Poppins', 'Montserrat', 'Lora', 'Playfair Display', 'Arial', 'Calibri', 'Cambria', 'Georgia', 'Times New Roman'];
  SL.SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];

  const rot = (a, n) => a.slice(n).concat(a.slice(0, n));
  // rangoli / mandala motif: concentric petals centred at (cx, cy)
  const rangoli = (cx, cy, r, c1, c2, c3) => {
    let s = `<g transform="translate(${cx} ${cy})">`;
    for (let i = 0; i < 16; i++) s += `<ellipse cx="0" cy="${-r * .62}" rx="${r * .11}" ry="${r * .3}" fill="${i % 2 ? c2 : c1}" opacity=".85" transform="rotate(${i * 22.5})"/>`;
    for (let i = 0; i < 8; i++) s += `<ellipse cx="0" cy="${-r * .3}" rx="${r * .08}" ry="${r * .17}" fill="${c3}" transform="rotate(${i * 45 + 22.5})"/>`;
    return s + `<circle r="${r * .14}" fill="${c2}"/><circle r="${r * .96}" fill="none" stroke="${c1}" stroke-width="2" stroke-dasharray="3 6"/></g>`;
  };
  const chakra = (cx, cy, r, c) => `<g transform="translate(${cx} ${cy})" fill="none" stroke="${c}" stroke-width="2"><circle r="${r}"/><circle r="${r * .12}" fill="${c}"/>${Array.from({ length: 24 }, (_, i) => `<line x1="0" y1="0" x2="${(r * Math.cos(i * Math.PI / 12)).toFixed(1)}" y2="${(r * Math.sin(i * Math.PI / 12)).toFixed(1)}"/>`).join('')}</g>`;
  const grid = (W, H, step, color, op) => { let s = `<g class="wash" stroke="${color}" stroke-width="1" opacity="${op}">`; for (let x = step; x < W; x += step) s += `<line x1="${x}" y1="0" x2="${x}" y2="${H}"/>`; for (let y = step; y < H; y += step) s += `<line x1="0" y1="${y}" x2="${W}" y2="${y}"/>`; return s + '</g>'; };

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
        (k === 'content' ? `<circle class="wash" cx="${W}" cy="0" r="${H * .45}" fill="url(#mg1)"/>` : `<circle class="wash" cx="${W * .9}" cy="${H * .1}" r="${H * .7}" fill="url(#mg1)"/><circle class="wash" cx="${W * .05}" cy="${H}" r="${H * .6}" fill="url(#mg2)"/>`) }),
    T({ id: 'sunrise', titleArea: { x: .0625, w: .54 }, areas: { section: { x: .0625, w: .54 } }, name: 'Sunrise', titleFont: 'Montserrat', bodyFont: 'Inter', titleBold: true,
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
        : k === 'section' ? `<line x1="${W * .086}" y1="${H * .555}" x2="${W * .914}" y2="${H * .555}" stroke="${c.accents[0]}" stroke-width="1.5"/>` : `<line x1="60" y1="${H - 14}" x2="${W - 60}" y2="${H - 14}" stroke="${c.accents[0]}" stroke-width="1"/>` }),
    T({ id: 'ocean', name: 'Ocean', titleFont: 'Poppins', bodyFont: 'Poppins', titleBold: true,
      bg: { type: 'gradient', color: '#1E3A8A', color2: '#0EA5E9', angle: 135 },
      colors: { bg: '#1E40AF', bg2: '#1D4ED8', text: '#E0F2FE', title: '#FFFFFF', accents: ['#FDE68A', '#5EEAD4', '#FCA5A5', '#C4B5FD', '#86EFAC', '#FFFFFF'] },
      variants: [{}, { bg: '#5B21B6', accents: ['#5EEAD4', '#FDE68A', '#FCA5A5', '#C4B5FD', '#86EFAC', '#FFFFFF'] }, { bg: '#065F46', accents: ['#FDE68A', '#FCA5A5', '#5EEAD4', '#C4B5FD', '#86EFAC', '#FFFFFF'] }, { bg: '#9F1239', accents: ['#FDE68A', '#5EEAD4', '#C4B5FD', '#FCA5A5', '#86EFAC', '#FFFFFF'] }],
      bgFor: (v) => v.bg ? { type: 'gradient', color: v.bg, color2: v.bg2 || '#0EA5E9', angle: 135 } : null,
      deco: (k, W, H, c) => `<path d="M0 ${H * .82} C${W * .25} ${H * .74} ${W * .5} ${H * .92} ${W} ${H * .8} V${H} H0 Z" fill="#fff" opacity=".08"/><path d="M0 ${H * .9} C${W * .3} ${H * .82} ${W * .6} ${H * .98} ${W} ${H * .88} V${H} H0 Z" fill="#fff" opacity=".1"/>` }),
    T({ id: 'geo', titleArea: { x: .0625, w: .6 }, areas: { section: { x: .0625, w: .6 } }, name: 'Geometric', titleFont: 'Montserrat', bodyFont: 'Montserrat', titleBold: true,
      colors: { bg: '#FFFFFF', bg2: '#F3F4F8', text: '#30364A', title: '#161A2B', accents: ['#5A4FCF', '#FF6B6B', '#FFC145', '#21B6A8', '#3A86FF', '#FF8FAB'] },
      variants: [{}, { accents: ['#21B6A8', '#FFC145', '#FF6B6B', '#5A4FCF', '#3A86FF', '#FF8FAB'] }, { accents: ['#FF6B6B', '#5A4FCF', '#FFC145', '#21B6A8', '#3A86FF', '#FF8FAB'] }, { accents: ['#3A86FF', '#FF8FAB', '#FFC145', '#21B6A8', '#5A4FCF', '#FF6B6B'] }],
      deco: (k, W, H, c) => k === 'content' ? `<polygon points="${W},0 ${W},90 ${W - 90},0" fill="${c.accents[0]}"/><polygon points="${W - 60},0 ${W - 120},0 ${W - 90},40" fill="${c.accents[2]}"/>`
        : `<polygon points="${W},0 ${W},${H * .5} ${W * .8},0" fill="${c.accents[0]}"/><polygon points="${W},${H * .42} ${W},${H} ${W * .84},${H}" fill="${c.accents[1]}"/><polygon points="${W * .7},0 ${W * .82},0 ${W * .76},${H * .16}" fill="${c.accents[2]}"/><circle cx="${W * .93}" cy="${H * .56}" r="30" fill="${c.accents[3]}"/><rect x="0" y="${H - 14}" width="${W * .35}" height="14" fill="${c.accents[0]}"/>` }),
    T({ id: 'block', titleArea: { x: .0625, w: .55 }, areas: { section: { x: .0625, w: .55 } }, name: 'Bold Block', titleFont: 'Poppins', bodyFont: 'Inter', titleBold: true,
      colors: { bg: '#FFFFFF', bg2: '#F5F3EE', text: '#2E2E2E', title: '#111111', accents: ['#FF5A36', '#111111', '#2563EB', '#16A34A', '#EAB308', '#9333EA'] },
      variants: [{}, { accents: ['#2563EB', '#111111', '#FF5A36', '#16A34A', '#EAB308', '#9333EA'] }, { accents: ['#16A34A', '#111111', '#2563EB', '#FF5A36', '#EAB308', '#9333EA'] }, { accents: ['#9333EA', '#111111', '#2563EB', '#16A34A', '#EAB308', '#FF5A36'] }],
      deco: (k, W, H, c) => k === 'title' ? `<rect x="${W * .64}" y="0" width="${W * .36}" height="${H}" fill="${c.accents[0]}"/><rect x="${W * .64 - 26}" y="${H * .38}" width="52" height="52" fill="${c.accents[1]}"/>`
        : k === 'section' ? `<rect x="0" y="0" width="${W}" height="${H}" fill="${c.accents[0]}" opacity=".08"/><rect x="0" y="0" width="${W * .03}" height="${H}" fill="${c.accents[0]}"/>` : `<rect x="0" y="0" width="12" height="${H}" fill="${c.accents[0]}"/>` }),
    T({ id: 'saffron', titleArea: { x: .086, w: .7 }, areas: { section: { x: .086, w: .7 } }, name: 'Saffron', titleFont: 'Poppins', bodyFont: 'Inter', titleBold: true,
      colors: { bg: '#FFFBF3', bg2: '#FFF1DB', text: '#3A3226', title: '#1F1A12', accents: ['#F28C28', '#128807', '#0B3D91', '#C2410C', '#B45309', '#7C3AED'] },
      variants: [{}, { accents: ['#128807', '#F28C28', '#0B3D91', '#C2410C', '#B45309', '#7C3AED'] }, { accents: ['#0B3D91', '#F28C28', '#128807', '#C2410C', '#B45309', '#7C3AED'] }, { accents: ['#C2410C', '#B45309', '#128807', '#0B3D91', '#F28C28', '#7C3AED'] }],
      deco: (k, W, H, c) => k === 'content' ? `<rect x="0" y="0" width="${W}" height="6" fill="${c.accents[0]}"/><rect x="0" y="${H - 6}" width="${W}" height="6" fill="${c.accents[1]}"/>`
        : `<rect x="0" y="0" width="${W}" height="18" fill="${c.accents[0]}"/><rect x="0" y="${H - 18}" width="${W}" height="18" fill="${c.accents[1]}"/><g transform="translate(${W - 140} ${H / 2})" fill="none" stroke="${c.accents[2]}" stroke-width="2" opacity=".55"><circle r="70"/><circle r="8" fill="${c.accents[2]}"/>${Array.from({ length: 24 }, (_, i) => `<line x1="0" y1="0" x2="${(70 * Math.cos(i * Math.PI / 12)).toFixed(1)}" y2="${(70 * Math.sin(i * Math.PI / 12)).toFixed(1)}"/>`).join('')}</g>` }),
    T({ id: 'slate', titleArea: { x: .0625, w: .46 }, areas: { section: { x: .0625, w: .46 } }, name: 'Slate', titleFont: 'Inter', bodyFont: 'Inter', titleBold: true,
      colors: { bg: '#F1F5F9', bg2: '#FFFFFF', text: '#334155', title: '#0F172A', accents: ['#0EA5E9', '#6366F1', '#F43F5E', '#10B981', '#F59E0B', '#64748B'] },
      variants: [{}, { accents: ['#6366F1', '#0EA5E9', '#F43F5E', '#10B981', '#F59E0B', '#64748B'] }, { accents: ['#F43F5E', '#6366F1', '#0EA5E9', '#10B981', '#F59E0B', '#64748B'] }, { accents: ['#10B981', '#0EA5E9', '#6366F1', '#F43F5E', '#F59E0B', '#64748B'] }],
      deco: (k, W, H, c) => k === 'content' ? `<rect x="40" y="44" width="6" height="64" rx="3" fill="${c.accents[0]}"/>`
        : `<rect x="${W * .55}" y="${H * .1}" width="${W * .4}" height="${H * .8}" rx="28" fill="${c.bg2}"/><rect x="${W * .6}" y="${H * .2}" width="${W * .3}" height="12" rx="6" fill="${c.accents[0]}"/><rect x="${W * .6}" y="${H * .2 + 30}" width="${W * .2}" height="12" rx="6" fill="${c.accents[1]}" opacity=".7"/><rect x="${W * .6}" y="${H * .2 + 60}" width="${W * .25}" height="12" rx="6" fill="${c.accents[2]}" opacity=".5"/>` }),
    T({ id: 'neon', name: 'Neon', areas: { title: { x: .0625, w: .5 }, section: { x: .0625, w: .5 } }, titleFont: 'Montserrat', bodyFont: 'Inter', titleBold: true,
      colors: { bg: '#0A0A0F', bg2: '#17171F', text: '#D4D4E0', title: '#FFFFFF', accents: ['#FF2E88', '#00E5FF', '#B6FF00', '#FFB800', '#8B5CF6', '#FF5C00'] },
      variants: [{}, { accents: ['#00E5FF', '#FF2E88', '#B6FF00', '#FFB800', '#8B5CF6', '#FF5C00'] }, { accents: ['#B6FF00', '#00E5FF', '#FF2E88', '#FFB800', '#8B5CF6', '#FF5C00'] }, { accents: ['#FFB800', '#FF2E88', '#00E5FF', '#B6FF00', '#8B5CF6', '#FF5C00'] }],
      deco: (k, W, H, c) => `<defs><linearGradient id="ng" x1="0" x2="1"><stop offset="0" stop-color="${c.accents[0]}"/><stop offset="1" stop-color="${c.accents[1]}"/></linearGradient></defs>` +
        (k === 'content' ? `<rect x="0" y="${H - 5}" width="${W}" height="5" fill="url(#ng)"/>` : `<rect x="0" y="${H - 10}" width="${W}" height="10" fill="url(#ng)"/>${Array.from({ length: 6 }, (_, i) => `<line x1="${W * .55 + i * 60}" y1="${H}" x2="${W * .8 + i * 60}" y2="0" stroke="url(#ng)" stroke-width="2" opacity="${.15 + i * .1}"/>`).join('')}`) }),
    /* ------------------------------------------------------------------ new in 1.2 ------------------------------------------------------------------ */
    T({ id: 'corporate', name: 'Corporate', titleFont: 'Inter', bodyFont: 'Inter', titleBold: true, areas: { title: { x: .07, w: .6 } },
      colors: { bg: '#FFFFFF', bg2: '#F2F5FA', text: '#2A3342', title: '#0B1F3A', accents: ['#1F5AA6', '#00A3E0', '#F2A900', '#5B6770', '#2E8B57', '#C8102E'] },
      variants: [{}, { accents: ['#00897B', '#1F5AA6', '#F2A900', '#5B6770', '#7CB342', '#C8102E'] }, { accents: ['#6A3FA0', '#00A3E0', '#F2A900', '#5B6770', '#2E8B57', '#C8102E'] }, { bg: '#0B1F3A', bg2: '#13294B', text: '#D5DEEA', title: '#FFFFFF', accents: ['#4DA3FF', '#00C2E0', '#F2B632', '#9AA7B4', '#4CC38A', '#FF5A6E'] }],
      deco: (k, W, H, c) => k === 'title' ? `<rect x="${W * .74}" y="0" width="${W * .26}" height="${H}" fill="${c.accents[0]}"/><rect x="${W * .74 - 12}" y="0" width="12" height="${H}" fill="${c.accents[1]}"/><rect x="${W * .07}" y="${H * .15}" width="64" height="6" fill="${c.accents[0]}"/><rect x="${W * .81}" y="${H * .72}" width="${W * .12}" height="${W * .12}" fill="none" stroke="#fff" stroke-opacity=".5" stroke-width="2"/>`
        : k === 'section' ? `<rect x="0" y="0" width="${W * .038}" height="${H}" fill="${c.accents[0]}"/><rect x="${W * .038}" y="0" width="6" height="${H}" fill="${c.accents[1]}"/>`
          : `<rect x="0" y="0" width="${W}" height="10" fill="${c.accents[0]}"/><rect x="0" y="0" width="${W * .18}" height="10" fill="${c.accents[1]}"/><rect x="${W * .0625}" y="${H - 12}" width="${W * .875}" height="2" fill="${c.accents[0]}" opacity=".35"/>` }),
    T({ id: 'executive', name: 'Executive', titleFont: 'Source Serif 4', bodyFont: 'Source Sans 3', titleBold: true,
      colors: { bg: '#FBFAF7', bg2: '#F1EDE4', text: '#33302B', title: '#1C1A17', accents: ['#A67C2E', '#1E3A5F', '#7A1F2B', '#4F6D5A', '#8C8C8C', '#C9A14A'] },
      variants: [{}, { bg: '#1E2023', bg2: '#2A2D31', text: '#D8D4CC', title: '#F5F1E8', accents: ['#C9A14A', '#7FA7D9', '#D9707E', '#8DB59A', '#B0B0B0', '#E3C27A'] }, { accents: ['#1E3A5F', '#A67C2E', '#7A1F2B', '#4F6D5A', '#8C8C8C', '#C9A14A'] }, { accents: ['#7A1F2B', '#A67C2E', '#1E3A5F', '#4F6D5A', '#8C8C8C', '#C9A14A'] }],
      deco: (k, W, H, c, L) => k === 'title' ? `<line x1="${W * .086}" y1="${H * .07}" x2="${W * .914}" y2="${H * .07}" stroke="${c.accents[0]}" stroke-width="1.5"/><line x1="${W * .086}" y1="${H * .93}" x2="${W * .914}" y2="${H * .93}" stroke="${c.accents[0]}" stroke-width="1.5"/><rect x="${W / 2 - 50}" y="${H * .538}" width="100" height="3" fill="${c.accents[0]}"/><circle cx="${W / 2}" cy="${H * .07}" r="5" fill="${c.accents[0]}"/>`
        : k === 'section' ? `<rect x="${W * .086}" y="${H * .552}" width="${W * .18}" height="3" fill="${c.accents[0]}"/><line x1="${W * .04}" y1="${H * .12}" x2="${W * .04}" y2="${H * .88}" stroke="${c.accents[0]}" stroke-width="1.5"/>`
          : (L === 'pictureCaption' ? '' : `<line x1="${W * .0625}" y1="${H * .222}" x2="${W * .9375}" y2="${H * .222}" stroke="${c.accents[0]}" stroke-width="1.2"/>`) + `<line x1="${W * .0625}" y1="${H - 14}" x2="${W * .9375}" y2="${H - 14}" stroke="${c.accents[0]}" stroke-width=".8" opacity=".6"/>` }),
    T({ id: 'academic', name: 'Academic', titleFont: 'EB Garamond', bodyFont: 'EB Garamond', titleBold: true,
      colors: { bg: '#FFFDF8', bg2: '#F3EEE2', text: '#2B2B2B', title: '#0E2A47', accents: ['#0E2A47', '#8B1E3F', '#B8860B', '#3C6E71', '#6B4E9B', '#5A5A5A'] },
      variants: [{}, { title: '#5E1129', accents: ['#8B1E3F', '#0E2A47', '#B8860B', '#3C6E71', '#6B4E9B', '#5A5A5A'] }, { title: '#21494B', accents: ['#3C6E71', '#B8860B', '#0E2A47', '#8B1E3F', '#6B4E9B', '#5A5A5A'] }, { bg: '#0E2A47', bg2: '#163A5F', text: '#E6E1D3', title: '#FFFFFF', accents: ['#D4AF37', '#E07A9A', '#F2C14E', '#7FC4C7', '#B9A3E3', '#C9C9C9'] }],
      deco: (k, W, H, c, L) => k === 'title' || k === 'section' ? `<rect x="0" y="0" width="${W}" height="${H * .12}" fill="${c.accents[0]}"/><rect x="0" y="${H * .12}" width="${W}" height="4" fill="${c.accents[2]}"/><line x1="${W * .2}" y1="${H * .9}" x2="${W * .8}" y2="${H * .9}" stroke="${c.accents[0]}" stroke-width="1"/><line x1="${W * .25}" y1="${H * .915}" x2="${W * .75}" y2="${H * .915}" stroke="${c.accents[0]}" stroke-width="1"/>`
        : `<line x1="${W * .03}" y1="${H * .056}" x2="${W * .03}" y2="${H * .917}" stroke="${c.accents[0]}" stroke-width="2"/><line x1="${W * .037}" y1="${H * .056}" x2="${W * .037}" y2="${H * .917}" stroke="${c.accents[0]}" stroke-width=".8"/>${L === 'pictureCaption' ? '' : `<line x1="${W * .0625}" y1="${H * .222}" x2="${W * .38}" y2="${H * .222}" stroke="${c.accents[2]}" stroke-width="2"/>`}` }),
    T({ id: 'minimal', name: 'Minimal', titleFont: 'DM Sans', bodyFont: 'DM Sans', titleBold: true, areas: { title: { x: .086, w: .74 } },
      colors: { bg: '#FFFFFF', bg2: '#F5F5F7', text: '#3A3A3C', title: '#111111', accents: ['#0071E3', '#FF375F', '#30D158', '#FF9F0A', '#BF5AF2', '#40C8E0'] },
      variants: [{}, { bg: '#000000', bg2: '#1C1C1E', text: '#D1D1D6', title: '#FFFFFF', accents: ['#0A84FF', '#FF375F', '#30D158', '#FF9F0A', '#BF5AF2', '#64D2FF'] }, { bg: '#F5F5F7', bg2: '#FFFFFF' }, { accents: ['#FF9F0A', '#0071E3', '#30D158', '#FF375F', '#BF5AF2', '#40C8E0'] }],
      deco: (k, W, H, c) => k === 'title' ? `<circle cx="${W * .086 + 8}" cy="${H * .17}" r="8" fill="${c.accents[0]}"/>`
        : k === 'section' ? `<rect x="${W * .086}" y="${H * .552}" width="${W * .06}" height="4" rx="2" fill="${c.accents[0]}"/>` : `<circle cx="${W * .0625 + 5}" cy="22" r="5" fill="${c.accents[0]}"/>` }),
    T({ id: 'bold', name: 'Bold', titleFont: 'Montserrat', bodyFont: 'Montserrat', titleBold: true, areas: { title: { x: .44, w: .5 }, section: { x: .44, w: .5 } },
      colors: { bg: '#FFFFFF', bg2: '#F4F4F4', text: '#222222', title: '#111111', accents: ['#FF4D00', '#111111', '#FFC400', '#00B3A6', '#3D5AFE', '#E91E63'] },
      variants: [{}, { accents: ['#3D5AFE', '#111111', '#FFC400', '#00B3A6', '#FF4D00', '#E91E63'] }, { accents: ['#00B3A6', '#111111', '#FFC400', '#FF4D00', '#3D5AFE', '#E91E63'] }, { bg: '#111111', bg2: '#1E1E1E', text: '#DDDDDD', title: '#FFFFFF', accents: ['#FFC400', '#FFFFFF', '#FF4D00', '#00B3A6', '#5C7CFF', '#FF4F8B'] }],
      deco: (k, W, H, c) => k === 'title' ? `<rect x="0" y="0" width="${W * .38}" height="${H}" fill="${c.accents[0]}"/><rect x="${W * .06}" y="${H * .78}" width="${W * .1}" height="${W * .1}" fill="${c.accents[2]}"/><rect x="${W * .38}" y="${H * .08}" width="${W * .04}" height="${W * .04}" fill="${c.accents[1]}"/>`
        : k === 'section' ? `<rect x="0" y="0" width="${W * .38}" height="${H * .9}" fill="${c.accents[0]}"/><rect x="${W * .06}" y="${H * .62}" width="${W * .1}" height="${W * .1}" fill="${c.accents[2]}"/><rect x="${W * .38}" y="${H * .08}" width="${W * .04}" height="${W * .04}" fill="${c.accents[1]}"/>`
        : `<rect x="0" y="0" width="${W * .03}" height="${H}" fill="${c.accents[0]}"/><rect x="0" y="${H * .82}" width="${W * .03}" height="${H * .18}" fill="${c.accents[2]}"/>` }),
    T({ id: 'editorial', name: 'Editorial', titleFont: 'Playfair Display', bodyFont: 'Lato', titleBold: true, areas: { title: { x: .086, w: .6 } },
      colors: { bg: '#FAF8F5', bg2: '#EFEAE2', text: '#2F2F2F', title: '#141414', accents: ['#C0392B', '#1F3A5F', '#D4A017', '#2E7D6B', '#6D4C7D', '#7F8C8D'] },
      variants: [{}, { accents: ['#1F3A5F', '#C0392B', '#D4A017', '#2E7D6B', '#6D4C7D', '#7F8C8D'] }, { accents: ['#2E7D6B', '#D4A017', '#C0392B', '#1F3A5F', '#6D4C7D', '#7F8C8D'] }, { bg: '#161616', bg2: '#242424', text: '#D6D2CB', title: '#FFFFFF', accents: ['#E5584A', '#7BA4D9', '#E8C04A', '#5FB8A2', '#B48FC4', '#A0AAAB'] }],
      deco: (k, W, H, c) => k === 'title' ? `<rect x="${W * .086}" y="${H * .09}" width="${W * .828}" height="3" fill="${c.title}"/><rect x="${W * .086}" y="${H * .09 + 8}" width="${W * .828}" height="1" fill="${c.title}"/><line x1="${W * .72}" y1="${H * .2}" x2="${W * .72}" y2="${H * .86}" stroke="${c.text}" stroke-opacity=".35"/>${Array.from({ length: 9 }, (_, i) => `<rect x="${W * .75}" y="${H * .24 + i * 30}" width="${W * (i % 3 === 2 ? .1 : .16)}" height="5" rx="2" fill="${c.text}" opacity=".18"/>`).join('')}<rect x="${W * .75}" y="${H * .24 - 26}" width="${W * .05}" height="10" fill="${c.accents[0]}"/>`
        : k === 'section' ? `<rect x="0" y="0" width="${W * .03}" height="${H}" fill="${c.accents[0]}"/><rect x="${W * .086}" y="${H * .2}" width="${W * .08}" height="6" fill="${c.accents[0]}"/>`
          : `<rect x="${W * .0625}" y="18" width="${W * .875}" height="2.5" fill="${c.title}"/><rect x="${W * .0625}" y="25" width="${W * .875}" height="1" fill="${c.title}"/><rect x="${W * .0625}" y="${H - 14}" width="${W * .875}" height="1" fill="${c.text}" opacity=".4"/>` }),
    T({ id: 'tech', name: 'Tech', titleFont: 'IBM Plex Sans', bodyFont: 'IBM Plex Sans', titleBold: true,
      colors: { bg: '#0D1117', bg2: '#161B22', text: '#C9D1D9', title: '#F0F6FC', accents: ['#58A6FF', '#3FB950', '#D29922', '#F778BA', '#A371F7', '#FF7B72'] },
      variants: [{}, { bg: '#F6F8FA', bg2: '#FFFFFF', text: '#24292F', title: '#0D1117', accents: ['#0969DA', '#1A7F37', '#9A6700', '#BF3989', '#8250DF', '#CF222E'] }, { accents: ['#3FB950', '#58A6FF', '#D29922', '#F778BA', '#A371F7', '#FF7B72'] }, { accents: ['#A371F7', '#58A6FF', '#3FB950', '#F778BA', '#D29922', '#FF7B72'] }],
      deco: (k, W, H, c) => grid(W, H, 40, c.text, .05) + (k === 'content' ? `<rect x="0" y="0" width="${W}" height="4" fill="${c.accents[0]}"/><g fill="${c.accents[0]}"><circle cx="${W - 26}" cy="22" r="4"/><circle cx="${W - 42}" cy="22" r="4" opacity=".6"/><circle cx="${W - 58}" cy="22" r="4" opacity=".3"/></g>`
        : `<path d="M${W * .04 + 60} ${H * .1} H${W * .04} V${H * .1 + 60}" fill="none" stroke="${c.accents[0]}" stroke-width="4"/><path d="M${W * .96 - 60} ${H * .9} H${W * .96} V${H * .9 - 60}" fill="none" stroke="${c.accents[0]}" stroke-width="4"/><text x="${W * .96}" y="${H * .1 + 20}" text-anchor="end" font-family="IBM Plex Mono, monospace" font-size="22" fill="${c.accents[1]}">&lt;/&gt;</text>${k === 'title' ? `<rect x="${W / 2 - 40}" y="${H * .538}" width="80" height="4" fill="${c.accents[1]}"/>` : `<rect x="${W * .086}" y="${H * .552}" width="80" height="4" fill="${c.accents[1]}"/>`}`) }),
    T({ id: 'startup', name: 'Startup', titleFont: 'Sora', bodyFont: 'Outfit', titleBold: true, areas: { title: { x: .07, w: .55 }, section: { x: .086, w: .55 } },
      colors: { bg: '#FFFFFF', bg2: '#F4F2FF', text: '#2D2A3E', title: '#16132B', accents: ['#6C47FF', '#FF6B9A', '#14C8A8', '#FFB020', '#2D9CFF', '#FF5A36'] },
      variants: [{}, { bg: '#14112A', bg2: '#1F1A3D', text: '#D9D5F2', title: '#FFFFFF', accents: ['#8B6BFF', '#FF7FA8', '#2FE0C0', '#FFC24D', '#55B0FF', '#FF7A5C'] }, { accents: ['#14C8A8', '#2D9CFF', '#6C47FF', '#FFB020', '#FF6B9A', '#FF5A36'] }, { accents: ['#FF5A36', '#FFB020', '#6C47FF', '#14C8A8', '#2D9CFF', '#FF6B9A'] }],
      deco: (k, W, H, c) => `<defs><linearGradient id="stg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="${c.accents[0]}"/><stop offset="1" stop-color="${c.accents[1]}"/></linearGradient></defs>` +
        (k === 'content' ? `<circle cx="${W}" cy="0" r="66" fill="url(#stg)"/><circle cx="${W - 72}" cy="14" r="9" fill="${c.accents[2]}"/>`
          : `<circle cx="${W * .84}" cy="${H * .42}" r="${H * .33}" fill="url(#stg)"/><circle cx="${W * .71}" cy="${H * .82}" r="${H * .06}" fill="${c.accents[2]}"/><circle cx="${W * .95}" cy="${H * .86}" r="${H * .035}" fill="${c.accents[3]}"/><circle cx="${W * .7}" cy="${H * .12}" r="${H * .025}" fill="${c.accents[4]}"/>`) }),
    T({ id: 'classic', name: 'Classic', titleFont: 'Cambria', bodyFont: 'Calibri', titleBold: false,
      colors: { bg: '#FFFFFF', bg2: '#EEF2F7', text: '#262626', title: '#1F3864', accents: ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'] },
      variants: [{}, { title: '#17365D', accents: ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646'] }, { title: '#375623', accents: ['#70AD47', '#4472C4', '#FFC000', '#ED7D31', '#5B9BD5', '#A5A5A5'] }, { bg: '#1F3864', bg2: '#2E4A7D', text: '#E3E9F3', title: '#FFFFFF', accents: ['#8FB4F0', '#F4A261', '#C9C9C9', '#FFD34D', '#9CC9EE', '#9ED07A'] }],
      deco: (k, W, H, c, L) => k === 'content' ? `<rect x="0" y="${H - 8}" width="${W}" height="8" fill="${c.accents[0]}"/>${L === 'pictureCaption' ? '' : `<line x1="${W * .0625}" y1="${H * .222}" x2="${W * .9375}" y2="${H * .222}" stroke="${c.accents[0]}" stroke-width="1"/>`}`
        : k === 'section' ? `<rect x="0" y="0" width="${W}" height="${H * .14}" fill="${c.accents[0]}"/><rect x="0" y="${H * .14 + 2}" width="${W}" height="4" fill="${c.accents[1]}"/><rect x="0" y="${H - 8}" width="${W}" height="8" fill="${c.accents[0]}"/>`
          : `<rect x="0" y="${H * .86}" width="${W}" height="${H * .14}" fill="${c.accents[0]}"/><rect x="0" y="${H * .86 - 6}" width="${W}" height="4" fill="${c.accents[1]}"/>` }),
    T({ id: 'chalkboard', name: 'Chalkboard', titleFont: 'Kalam', bodyFont: 'Kalam', titleBold: true,
      bg: { type: 'gradient', color: '#2B4536', color2: '#37543F', angle: 160 },
      colors: { bg: '#2F4A3A', bg2: '#3A5A48', text: '#F2F2EA', title: '#FFFFFF', accents: ['#F7E36B', '#8FD3FF', '#FF9BB3', '#B6F09C', '#FFB86B', '#E0E0E0'] },
      variants: [{}, { bg: '#222826', bg2: '#2E3532' }, { bg: '#1F3B57', bg2: '#2A4C6E' }, { bg: '#3B2F2A', bg2: '#4A3B34' }],
      bgFor: (v) => v.bg ? { type: 'gradient', color: v.bg, color2: v.bg2 || v.bg, angle: 160 } : null,
      deco: (k, W, H, c) => `<g class="wash" fill="#ffffff"><ellipse cx="${W * .3}" cy="${H * .35}" rx="${W * .22}" ry="${H * .12}" opacity=".035"/><ellipse cx="${W * .7}" cy="${H * .66}" rx="${W * .25}" ry="${H * .1}" opacity=".03"/></g><rect x="9" y="9" width="${W - 18}" height="${H - 18}" fill="none" stroke="#8B5A2B" stroke-width="18"/><rect x="18" y="18" width="${W - 36}" height="${H - 36}" fill="none" stroke="#5E3B1B" stroke-width="2"/>` +
        (k === 'title' ? `<rect x="${W * .3}" y="${H - 16}" width="60" height="7" rx="3" fill="#F2F2EA"/><rect x="${W * .3 + 72}" y="${H - 16}" width="40" height="7" rx="3" fill="${c.accents[0]}"/>` : '') }),
    T({ id: 'bharat', name: 'Bharat', titleFont: 'Mukta', bodyFont: 'Hind', titleBold: true,
      colors: { bg: '#FFFDF7', bg2: '#FFF3E0', text: '#2E2A24', title: '#1A1712', accents: ['#FF9933', '#138808', '#000080', '#C2185B', '#F9A825', '#6A1B9A'] },
      variants: [{}, { accents: ['#138808', '#FF9933', '#000080', '#C2185B', '#F9A825', '#6A1B9A'] }, { accents: ['#000080', '#FF9933', '#138808', '#C2185B', '#F9A825', '#6A1B9A'] }, { bg: '#4A0E1E', bg2: '#5C1628', text: '#F6E7D7', title: '#FFE9B0', accents: ['#FFB347', '#6FD36F', '#9FB4FF', '#FF7FAF', '#FFD54F', '#D1A3FF'] }],
      deco: (k, W, H, c) => k === 'content' ? `<rect x="0" y="0" width="${W}" height="6" fill="#FF9933"/><rect x="0" y="6" width="${W}" height="3" fill="${c.bg2}"/><rect x="0" y="9" width="${W}" height="6" fill="#138808"/>${rangoli(W, 0, 34, c.accents[0], c.accents[1], c.accents[3])}`
        : `<rect x="0" y="0" width="${W}" height="${H * .022}" fill="#FF9933"/><rect x="0" y="${H * .978}" width="${W}" height="${H * .022}" fill="#138808"/>${rangoli(W, 0, H * .2, c.accents[0], c.accents[1], c.accents[3])}${k === 'title' ? rangoli(0, H, H * .2, c.accents[1], c.accents[0], c.accents[4]) + chakra(W / 2, H * .17, 20, '#000080') : rangoli(0, 0, H * .14, c.accents[1], c.accents[0], c.accents[4])}` }),
    T({ id: 'darkpro', name: 'Dark Pro', titleFont: 'Manrope', bodyFont: 'Manrope', titleBold: true, areas: { title: { x: .086, w: .66 }, section: { x: .086, w: .66 } },
      bg: { type: 'gradient', color: '#0B0D12', color2: '#161B26', angle: 150 },
      colors: { bg: '#0F1218', bg2: '#1A1F2B', text: '#C7CEDB', title: '#FFFFFF', accents: ['#4F8CFF', '#22D3EE', '#A78BFA', '#F59E0B', '#10B981', '#F43F5E'] },
      variants: [{}, { accents: ['#22D3EE', '#4F8CFF', '#A78BFA', '#F59E0B', '#10B981', '#F43F5E'] }, { accents: ['#A78BFA', '#F472B6', '#4F8CFF', '#F59E0B', '#10B981', '#22D3EE'] }, { bg: '#0A1630', bg2: '#10244A', accents: ['#5EA0FF', '#22D3EE', '#A78BFA', '#FBBF24', '#34D399', '#FB7185'] }],
      bgFor: (v) => v.bg ? { type: 'gradient', color: v.bg, color2: v.bg2 || v.bg, angle: 150 } : null,
      deco: (k, W, H, c) => `<defs><radialGradient id="dpg"><stop offset="0" stop-color="${c.accents[0]}" stop-opacity=".35"/><stop offset="1" stop-color="${c.accents[0]}" stop-opacity="0"/></radialGradient></defs>` +
        (k === 'content' ? `<rect x="${W * .0625}" y="20" width="48" height="4" rx="2" fill="${c.accents[0]}"/><circle cx="${W}" cy="${H}" r="${H * .3}" fill="url(#dpg)" class="wash"/>`
          : `<circle cx="${W * 1.02}" cy="${H * 1.05}" r="${H * .62}" fill="url(#dpg)" class="wash"/><circle cx="${W * 1.02}" cy="${H * 1.05}" r="${H * .42}" fill="none" stroke="${c.accents[0]}" stroke-width="2" opacity=".7"/><circle cx="${W * 1.02}" cy="${H * 1.05}" r="${H * .5}" fill="none" stroke="${c.accents[1]}" stroke-width="1" opacity=".4"/><rect x="${W * .086}" y="${H * .16}" width="56" height="5" rx="2.5" fill="${c.accents[0]}"/>`) }),
    T({ id: 'gradient', name: 'Gradient', titleFont: 'Outfit', bodyFont: 'Outfit', titleBold: true,
      bg: { type: 'gradient', color: '#5B2BE0', color2: '#E0457B', angle: 135 },
      colors: { bg: '#7A3AD0', bg2: '#8E44C8', text: '#F3EFFF', title: '#FFFFFF', accents: ['#FFE066', '#7CF7D4', '#FFFFFF', '#FFB3C7', '#9AD0FF', '#FFC56B'] },
      variants: [{}, { bg: '#1E3C72', bg2: '#2A9D8F' }, { bg: '#C2410C', bg2: '#E11D48' }, { bg: '#134E5E', bg2: '#4C8C5C' }],
      bgFor: (v) => v.bg ? { type: 'gradient', color: v.bg, color2: v.bg2 || v.bg, angle: 135 } : null,
      deco: (k, W, H, c) => `<g class="wash" fill="#ffffff"><circle cx="${W * .9}" cy="${H * .15}" r="${H * .4}" opacity=".08"/><circle cx="${W * .08}" cy="${H * .95}" r="${H * .35}" opacity=".07"/></g>` + (k === 'title' ? `<rect x="${W / 2 - 40}" y="${H * .538}" width="80" height="4" rx="2" fill="#ffffff" opacity=".85"/>` : k === 'section' ? `<rect x="${W * .086}" y="${H * .552}" width="80" height="4" rx="2" fill="#ffffff" opacity=".85"/>` : '') }),
    T({ id: 'nordic', name: 'Nordic', titleFont: 'Work Sans', bodyFont: 'Work Sans', titleBold: true, areas: { title: { x: .086, w: .6 } },
      colors: { bg: '#F4F6F5', bg2: '#FFFFFF', text: '#3B4A54', title: '#1E2B33', accents: ['#5E8C9A', '#D08C60', '#8FA98F', '#C9B79C', '#7A8FA6', '#B5656B'] },
      variants: [{}, { accents: ['#D08C60', '#5E8C9A', '#8FA98F', '#C9B79C', '#7A8FA6', '#B5656B'] }, { accents: ['#8FA98F', '#5E8C9A', '#D08C60', '#C9B79C', '#7A8FA6', '#B5656B'] }, { bg: '#1E2B33', bg2: '#2A3A44', text: '#D3DCE0', title: '#FFFFFF', accents: ['#8DBCC9', '#E8A77C', '#AFC7AF', '#E0D1B8', '#A3B6CC', '#D98A90'] }],
      deco: (k, W, H, c) => k === 'title' ? `<rect x="${W * .74}" y="${H * .18}" width="${W * .2}" height="${H * .5}" rx="${W * .1}" fill="${c.accents[0]}" opacity=".85"/><circle cx="${W * .76}" cy="${H * .74}" r="${H * .1}" fill="${c.accents[1]}" opacity=".8"/><rect x="${W * .86}" y="${H * .7}" width="${W * .1}" height="${H * .2}" rx="${W * .05}" fill="${c.accents[2]}" opacity=".75"/>`
        : k === 'section' ? `<circle cx="${W}" cy="0" r="${H * .26}" fill="${c.accents[0]}" opacity=".6"/><circle cx="${W * .88}" cy="${H * .2}" r="${H * .05}" fill="${c.accents[1]}"/>`
          : `<ellipse cx="${W - 30}" cy="30" rx="40" ry="26" fill="${c.accents[0]}" opacity=".5"/><circle cx="${W - 64}" cy="14" r="8" fill="${c.accents[1]}" opacity=".7"/>` }),
    T({ id: 'royal', name: 'Royal', titleFont: 'Cormorant Garamond', bodyFont: 'Raleway', titleBold: true,
      bg: { type: 'gradient', color: '#1F1233', color2: '#2D1B4E', angle: 150 },
      colors: { bg: '#24163C', bg2: '#2E1D4A', text: '#E9E1F5', title: '#F7E7B4', accents: ['#D4AF37', '#B57EDC', '#E8C4A0', '#7FB3D5', '#E57373', '#A5D6A7'] },
      variants: [{}, { bg: '#0F2A22', bg2: '#173D31' }, { bg: '#121A33', bg2: '#1B2749' }, { bg: '#3A0D1A', bg2: '#521428' }],
      bgFor: (v) => v.bg ? { type: 'gradient', color: v.bg, color2: v.bg2, angle: 150 } : null,
      deco: (k, W, H, c) => `<rect x="12" y="12" width="${W - 24}" height="${H - 24}" fill="none" stroke="${c.accents[0]}" stroke-width="1.5"/><rect x="17" y="17" width="${W - 34}" height="${H - 34}" fill="none" stroke="${c.accents[0]}" stroke-width=".6" opacity=".7"/>` +
        [[12, 12], [W - 12, 12], [12, H - 12], [W - 12, H - 12]].map(([x, y]) => `<rect x="${x - 7}" y="${y - 7}" width="14" height="14" transform="rotate(45 ${x} ${y})" fill="${c.accents[0]}"/>`).join('') +
        (k === 'title' ? `<path d="M${W / 2 - 90} ${H * .542} H${W / 2 - 12} M${W / 2 + 12} ${H * .542} H${W / 2 + 90}" stroke="${c.accents[0]}" stroke-width="1.2"/><rect x="${W / 2 - 5}" y="${H * .542 - 5}" width="10" height="10" transform="rotate(45 ${W / 2} ${H * .542})" fill="${c.accents[0]}"/>` : '') }),
    T({ id: 'metro', name: 'Metro', titleFont: 'Oswald', bodyFont: 'Roboto', titleBold: true, areas: { title: { x: .07, w: .58 }, section: { x: .1, w: .8 } },
      colors: { bg: '#FFFFFF', bg2: '#F0F0F0', text: '#333333', title: '#1A1A1A', accents: ['#0078D7', '#E81123', '#FFB900', '#107C10', '#5C2D91', '#00B7C3'] },
      variants: [{}, { accents: ['#E81123', '#0078D7', '#FFB900', '#107C10', '#5C2D91', '#00B7C3'] }, { accents: ['#107C10', '#00B7C3', '#0078D7', '#FFB900', '#5C2D91', '#E81123'] }, { bg: '#1A1A1A', bg2: '#2B2B2B', text: '#E0E0E0', title: '#FFFFFF' }],
      deco: (k, W, H, c) => {
        if (k === 'content') return c.accents.slice(0, 4).map((a, i) => `<rect x="${W - 20 - (4 - i) * 18}" y="12" width="14" height="14" fill="${a}"/>`).join('');
        if (k === 'section') return c.accents.slice(0, 5).map((a, i) => `<rect x="0" y="${i * H / 5}" width="${W * .042}" height="${H / 5 - 4}" fill="${a}"/>`).join('');
        const s = W * .11, g = 8, x0 = W * .7, y0 = H * .14; let o = '';
        [[0, 0, 0], [1, 0, 1], [0, 1, 2], [1, 1, 3], [0, 2, 4], [1, 2, 5]].forEach(([cx, cy, i]) => { o += `<rect x="${x0 + cx * (s + g)}" y="${y0 + cy * (s + g)}" width="${s}" height="${s}" fill="${c.accents[i]}"/>`; });
        return o;
      } }),
    T({ id: 'scholar', name: 'Scholar', titleFont: 'Merriweather', bodyFont: 'Open Sans', titleBold: true, areas: { title: { x: .1, w: .82 }, section: { x: .1, w: .82 } },
      colors: { bg: '#FFFFFF', bg2: '#F7F3F3', text: '#333333', title: '#5B0E1C', accents: ['#7A1324', '#C99A2E', '#2F4858', '#33658A', '#86BBD8', '#758E4F'] },
      variants: [{}, { title: '#1E3A4C', accents: ['#2F4858', '#C99A2E', '#7A1324', '#33658A', '#86BBD8', '#758E4F'] }, { title: '#2D4A1E', accents: ['#4D6B2F', '#C99A2E', '#7A1324', '#33658A', '#86BBD8', '#2F4858'] }, { bg: '#2A0A10', bg2: '#3B1219', text: '#EADCDC', title: '#F3D9A4', accents: ['#E07A8A', '#E3BE63', '#8FB3C7', '#7DAFD6', '#B9DCEB', '#A9C38A'] }],
      deco: (k, W, H, c) => k === 'content' ? `<rect x="0" y="0" width="${W * .02}" height="${H}" fill="${c.accents[0]}"/><rect x="${W * .02}" y="0" width="3" height="${H}" fill="${c.accents[1]}"/>`
        : `<rect x="0" y="0" width="${W * .04}" height="${H}" fill="${c.accents[0]}"/><rect x="${W * .04}" y="0" width="4" height="${H}" fill="${c.accents[1]}"/><rect x="${W * .1}" y="${H * .17}" width="${W * .06}" height="4" fill="${c.accents[1]}"/>` }),
    T({ id: 'blueprint', name: 'Blueprint', titleFont: 'Roboto Slab', bodyFont: 'Roboto', titleBold: true,
      colors: { bg: '#0F3D6E', bg2: '#154C85', text: '#DCEBFA', title: '#FFFFFF', accents: ['#FFFFFF', '#7FD1FF', '#FFD166', '#9BF6A5', '#FF9F80', '#C3B1FF'] },
      variants: [{}, { bg: '#0B2340', bg2: '#12345C' }, { bg: '#F4F8FC', bg2: '#FFFFFF', text: '#1C3D5A', title: '#0F3D6E', accents: ['#0F5FA8', '#2BA3D9', '#D98E04', '#2E9E5B', '#D9534F', '#6A5ACD'] }, { bg: '#0E4D4A', bg2: '#13625E' }],
      deco: (k, W, H, c) => grid(W, H, 20, c.text, .07) + grid(W, H, 100, c.text, .1) + (k === 'content' ? `<rect x="${W - 210}" y="8" width="190" height="24" fill="none" stroke="${c.text}" stroke-width="1" opacity=".55"/><line x1="${W - 115}" y1="8" x2="${W - 115}" y2="32" stroke="${c.text}" stroke-width="1" opacity=".55"/>`
        : `<g fill="none" stroke="${c.text}" stroke-width="1.2" opacity=".7"><rect x="${W * .72}" y="${H * .76}" width="${W * .24}" height="${H * .12}"/><line x1="${W * .72}" y1="${H * .82}" x2="${W * .96}" y2="${H * .82}"/><line x1="${W * .84}" y1="${H * .82}" x2="${W * .84}" y2="${H * .88}"/></g><circle cx="${W * .05}" cy="${H * .1}" r="22" fill="none" stroke="${c.accents[1]}" stroke-width="2"/><line x1="${W * .05 - 32}" y1="${H * .1}" x2="${W * .05 + 32}" y2="${H * .1}" stroke="${c.accents[1]}" stroke-width="1.2"/><line x1="${W * .05}" y1="${H * .1 - 32}" x2="${W * .05}" y2="${H * .1 + 32}" stroke="${c.accents[1]}" stroke-width="1.2"/>`) }),
    T({ id: 'pastel', name: 'Pastel', titleFont: 'Baloo 2', bodyFont: 'Baloo 2', titleBold: true,
      colors: { bg: '#FFF9F2', bg2: '#FFFFFF', text: '#4A4058', title: '#2E2340', accents: ['#FF8FAB', '#5BB8F5', '#FFC94D', '#5ED3B0', '#A98BFF', '#FF9F68'] },
      variants: [{}, { bg: '#F6F1FF', accents: ['#A98BFF', '#FF8FAB', '#5BB8F5', '#FFC94D', '#5ED3B0', '#FF9F68'] }, { bg: '#F0FAF6', accents: ['#5ED3B0', '#5BB8F5', '#FFC94D', '#FF8FAB', '#A98BFF', '#FF9F68'] }, { bg: '#FFF4F7', accents: rot(['#FF8FAB', '#5BB8F5', '#FFC94D', '#5ED3B0', '#A98BFF', '#FF9F68'], 2) }],
      deco: (k, W, H, c) => k === 'content' ? `<circle cx="${W - 22}" cy="22" r="16" fill="${c.accents[0]}" opacity=".8"/><circle cx="${W - 52}" cy="14" r="8" fill="${c.accents[2]}"/><circle cx="22" cy="22" r="10" fill="${c.accents[1]}" opacity=".7"/>`
        : k === 'section' ? `<circle cx="${W * .05}" cy="${H * .1}" r="${H * .14}" fill="${c.accents[0]}" opacity=".75"/><circle cx="${W * .14}" cy="${H * .06}" r="${H * .05}" fill="${c.accents[2]}"/><circle cx="${W * .95}" cy="${H * .12}" r="${H * .15}" fill="${c.accents[1]}" opacity=".7"/><circle cx="${W * .86}" cy="${H * .06}" r="${H * .04}" fill="${c.accents[3]}"/>`
        : `<circle cx="${W * .05}" cy="${H * .1}" r="${H * .14}" fill="${c.accents[0]}" opacity=".75"/><circle cx="${W * .14}" cy="${H * .06}" r="${H * .05}" fill="${c.accents[2]}"/><circle cx="${W * .95}" cy="${H * .9}" r="${H * .15}" fill="${c.accents[1]}" opacity=".7"/><circle cx="${W * .86}" cy="${H * .95}" r="${H * .05}" fill="${c.accents[3]}"/><path d="M${W * .9} ${H * .12} l8 -16 l8 16 l-8 16 z" fill="${c.accents[4]}"/><path d="M${W * .07} ${H * .86} l7 -14 l7 14 l-7 14 z" fill="${c.accents[5]}"/>` }),
    T({ id: 'ink', name: 'Ink', titleFont: 'Libre Baskerville', bodyFont: 'Source Sans 3', titleBold: true, areas: { title: { x: .12, w: .76 }, section: { x: .55, w: .38 } },
      colors: { bg: '#FFFFFF', bg2: '#F2F2F2', text: '#1A1A1A', title: '#000000', accents: ['#000000', '#E63946', '#555555', '#999999', '#1D3557', '#457B9D'] },
      variants: [{}, { bg: '#000000', bg2: '#1A1A1A', text: '#E6E6E6', title: '#FFFFFF', accents: ['#FFFFFF', '#FF5A67', '#BBBBBB', '#888888', '#8FB3E0', '#7FB7D9'] }, { accents: ['#1D3557', '#E63946', '#555555', '#999999', '#000000', '#457B9D'] }, { bg: '#F7F3EA', bg2: '#EFE8D8' }],
      deco: (k, W, H, c) => k === 'title' ? `<rect x="${W * .07}" y="${H * .22}" width="10" height="${H * .5}" fill="${c.accents[0]}"/><rect x="${W * .07}" y="${H * .22}" width="10" height="${H * .08}" fill="${c.accents[1]}"/>`
        : k === 'section' ? `<rect x="0" y="0" width="${W * .48}" height="${H * .88}" fill="${c.accents[0]}"/><rect x="${W * .48 - 60}" y="${H * .5 - 30}" width="60" height="60" fill="${c.accents[1]}"/>`
          : `<rect x="0" y="0" width="${W}" height="6" fill="${c.accents[0]}"/><rect x="${W * .0625}" y="6" width="40" height="6" fill="${c.accents[1]}"/>` }),
  ];
  // gallery order: Clean first, then the 1.2 professional themes, then the classic set
  (() => { const NEW = ['corporate', 'executive', 'academic', 'minimal', 'bold', 'editorial', 'tech', 'startup', 'classic', 'chalkboard', 'bharat', 'darkpro', 'gradient', 'nordic', 'royal', 'metro', 'scholar', 'blueprint', 'pastel', 'ink'];
    const by = (id) => SL.THEMES.find(t => t.id === id); SL.THEMES = [by('clean'), ...NEW.map(by), ...SL.THEMES.filter(t => t.id !== 'clean' && !NEW.includes(t.id))]; })();
  SL.themeById = (id) => SL.THEMES.find(t => t.id === id) || SL.THEMES[0];
  /* Theme object for a presentation; imported files carry their own palette/fonts (theme.custom). */
  SL.themeOf = (pres) => {
    const t = (pres && pres.theme) || {};
    if (t.id === 'custom' && t.custom) return Object.assign({ id: 'custom', name: t.custom.name || 'Imported theme', titleBold: false, variants: [{}], deco: null }, t.custom);
    return SL.themeById(t.id);
  };

  /* Effective palette for the presentation theme (+variant, + customised colours). */
  SL.palette = (pres) => {
    const th = SL.themeOf(pres);
    const v = th.variants?.[pres?.theme?.variant || 0] || {};
    const c = Object.assign({}, th.colors, v);
    c.accents = (v.accents || th.colors.accents).slice();
    const o = pres?.theme?.colors;
    if (o) { ['bg', 'bg2', 'text', 'title'].forEach(k => { if (o[k]) c[k] = o[k]; }); if (o.accents) o.accents.forEach((a, i) => { if (a) c.accents[i] = a; }); }
    return c;
  };
  /* Background used by slides that don't set their own: master (per layout) → customised colours → theme. */
  SL.themeBg = (pres, layout) => {
    const m = pres && pres.master;
    if (m) { if (layout && m.layoutBg && m.layoutBg[layout]) return m.layoutBg[layout]; if (m.bg) return m.bg; }
    const th = SL.themeOf(pres); const v = th.variants?.[pres?.theme?.variant || 0] || {}; const c = SL.palette(pres);
    const o = pres?.theme?.colors;
    if (o && o.bg) { const base = th.bgFor ? th.bgFor(v) : null; const b = base || th.bg; if (b && b.type === 'gradient' && !o.bgSolid) return { type: 'gradient', color: o.bg, color2: o.bg2 || o.bg, angle: b.angle ?? 135 }; return { type: 'solid', color: o.bg }; }
    if (th.bgFor) { const b = th.bgFor(v); if (b) return b; }
    if (v.bg) return { type: 'solid', color: v.bg };
    return th.bg || { type: 'solid', color: c.bg };
  };
  function hex2rgb(h) { h = String(h).replace('#', ''); if (h.length === 3) h = h.split('').map(c => c + c).join(''); return [parseInt(h.slice(0, 2), 16) || 0, parseInt(h.slice(2, 4), 16) || 0, parseInt(h.slice(4, 6), 16) || 0]; }
  function rgb2hex(r) { return '#' + r.map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join(''); }
  SL.mixColor = (hex, k) => k >= 0 ? rgb2hex(hex2rgb(hex).map(v => v + (255 - v) * k)) : rgb2hex(hex2rgb(hex).map(v => v * (1 + k)));
  SL.color = (v, pres) => {
    if (!v || v === 'none') return v;
    if (typeof v === 'object') return v.color ? SL.color(v.color, pres) : null;
    if (v[0] === '#' || v.startsWith('rgb')) return v;
    const mod = /^([\w]+)\/([ld])(\d+)$/.exec(v);
    if (mod) { const base = SL.color(mod[1], pres); if (!base || base[0] !== '#') return base; return SL.mixColor(base.slice(0, 7), (mod[2] === 'l' ? 1 : -1) * (+mod[3] / 100)); }
    const c = SL.palette(pres);
    const m = /^accent([1-6])$/.exec(v); if (m) return c.accents[+m[1] - 1];
    return c[v] || v;
  };
  SL.fontOf = (role, pres) => {
    const th = SL.themeOf(pres); const o = pres?.theme?.fonts || {}; const m = pres?.master || {};
    if (role === 'title') return (m.title && m.title.font) || o.title || th.titleFont;
    return (m.body && m.body.font) || o.body || th.bodyFont;
  };
  /* Theme fonts ignoring master overrides (for Design ▸ Fonts). */
  SL.themeFonts = (pres) => { const th = SL.themeOf(pres); const o = pres?.theme?.fonts || {}; return { title: o.title || th.titleFont, body: o.body || th.bodyFont }; };

  /* ---------- Theme fonts (Design ▸ Fonts) and colour schemes (Design ▸ Colors) ---------- */
  SL.FONT_PAIRS = [
    ['Inter', 'Inter', 'Inter'], ['Corporate', 'Inter', 'Source Sans 3'], ['Executive', 'Source Serif 4', 'Source Sans 3'], ['Academic', 'EB Garamond', 'EB Garamond'],
    ['Minimal', 'DM Sans', 'DM Sans'], ['Bold', 'Montserrat', 'Open Sans'], ['Editorial', 'Playfair Display', 'Lato'], ['Tech', 'IBM Plex Sans', 'IBM Plex Serif'],
    ['Startup', 'Sora', 'Outfit'], ['Classic Office', 'Cambria', 'Calibri'], ['Modern Office', 'Calibri', 'Calibri'], ['Friendly', 'Poppins', 'Inter'],
    ['Research', 'Merriweather', 'Open Sans'], ['Headline', 'Oswald', 'Roboto'], ['Elegant', 'Cormorant Garamond', 'Raleway'], ['Engineering', 'Roboto Slab', 'Roboto'],
    ['Literary', 'Libre Baskerville', 'Source Sans 3'], ['Humanist', 'Lora', 'Work Sans'], ['Geometric', 'Manrope', 'Manrope'], ['Display', 'Bebas Neue', 'Montserrat'],
    ['Handwriting', 'Kalam', 'Kalam'], ['Hindi + English (Mukta)', 'Mukta', 'Hind'], ['Hindi + English (Noto)', 'Noto Serif Devanagari', 'Noto Sans Devanagari'], ['Playful (Baloo)', 'Baloo 2', 'Baloo 2'],
  ].map(([name, title, body]) => ({ name, title, body }));
  SL.COLOR_SCHEMES = [
    ['Office', ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']],
    ['Office 2007', ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646']],
    ['Blue', ['#0F6FC6', '#009DD9', '#0BD0D9', '#10CF9B', '#7CCA62', '#A5C249']],
    ['Blue Green', ['#3494BA', '#58B6C0', '#75BDA7', '#7A8C8E', '#84ACB6', '#2683C6']],
    ['Green', ['#549E39', '#8AB833', '#C0CF3A', '#029676', '#4AB5C4', '#0989B1']],
    ['Yellow Orange', ['#FFCA08', '#F8931D', '#CE8D3E', '#EC7016', '#E64823', '#9C6A6A']],
    ['Red Orange', ['#E84C22', '#FFBD47', '#B64926', '#FF8427', '#CC9900', '#B22600']],
    ['Red Violet', ['#E32D91', '#C830CC', '#4EA6DC', '#4775E7', '#8971E1', '#D54773']],
    ['Violet', ['#AD84C6', '#8784C7', '#5D739A', '#6997AF', '#84ACB6', '#6F8183']],
    ['Marigold', ['#F2A900', '#6B4E9B', '#C8102E', '#00A3E0', '#2E8B57', '#5B6770']],
    ['Saffron India', ['#FF9933', '#138808', '#000080', '#C2185B', '#F9A825', '#6A1B9A']],
    ['Ocean', ['#0077B6', '#00B4D8', '#48CAE4', '#023E8A', '#90E0EF', '#03045E']],
    ['Forest', ['#2D6A4F', '#40916C', '#74C69D', '#B7E4C7', '#1B4332', '#95D5B2']],
    ['Sunset', ['#F94144', '#F3722C', '#F8961E', '#F9C74F', '#90BE6D', '#43AA8B']],
    ['Berry', ['#7B2CBF', '#C77DFF', '#E0AAFF', '#9D4EDD', '#5A189A', '#3C096C']],
    ['Earth', ['#8C5E3C', '#C08552', '#DAB49D', '#5E6E4A', '#A3B18A', '#6B4226']],
    ['Pastel', ['#FF8FAB', '#5BB8F5', '#FFC94D', '#5ED3B0', '#A98BFF', '#FF9F68']],
    ['Neon', ['#FF2E88', '#00E5FF', '#B6FF00', '#FFB800', '#8B5CF6', '#FF5C00']],
    ['Grayscale', ['#404040', '#5F5F5F', '#7F7F7F', '#A5A5A5', '#BFBFBF', '#262626']],
    ['Corporate Navy', ['#1F5AA6', '#00A3E0', '#F2A900', '#5B6770', '#2E8B57', '#C8102E']],
  ].map(([name, accents]) => ({ name, accents }));

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
  SL.layoutName = (id) => (SL.LAYOUTS.find(l => l.id === id) || { name: id }).name;
  /* Key of a placeholder inside its layout: ph + index among same-ph placeholders (title0, body0, body1, subtitle0, pic0). */
  SL.phKeys = (specs) => { const n = {}; return specs.map(s => { const k = s.ph + (n[s.ph] = (n[s.ph] ?? -1) + 1); return k; }); };
  // Returns element specs (without ids). Sizes in pt, geometry in logical px.
  SL.layoutPlaceholders = (layout, W, H, pres, opts = {}) => {
    const th = pres ? SL.themeOf(pres) : null;
    const areas = (th && th.areas) || {};
    const ta = layout === 'title' && th ? (areas.title || th.titleArea) : layout === 'section' && th ? areas.section : null;
    const sx = W / 1280, sy = H / 720, R = (x, y, w, h) => ({ x: Math.round(x * sx), y: Math.round(y * sy), w: Math.round(w * sx), h: Math.round(h * sy) });
    const title = (g, extra) => Object.assign({ type: 'text', ph: 'title', role: 'title', size: 40, valign: 'middle', placeholder: 'Click to add title' }, g, extra);
    const body = (g, extra) => Object.assign({ type: 'text', ph: 'body', role: 'body', size: 24, valign: 'top', pspace: 8, list: 'ul', placeholder: 'Click to add text' }, g, extra);
    let specs;
    switch (layout) {
      case 'title': {
        if (ta) { const x = Math.round(ta.x * W), w = Math.round(ta.w * W); // left-aligned in the space the theme graphics leave free
          specs = [title(Object.assign(R(0, 150, 0, 230), { x, w }), { size: 54, align: 'left', valign: 'bottom' }), { type: 'text', ph: 'subtitle', role: 'body', size: 24, align: 'left', valign: 'top', placeholder: 'Click to add subtitle', ...Object.assign(R(0, 400, 0, 130), { x, w }) }]; break; }
        specs = [title(R(110, 190, 1060, 190), { size: 54, align: 'center', valign: 'bottom' }), { type: 'text', ph: 'subtitle', role: 'body', size: 24, align: 'center', valign: 'top', placeholder: 'Click to add subtitle', ...R(110, 400, 1060, 110) }]; break;
      }
      case 'titleContent': specs = [title(R(80, 40, 1120, 110)), body(R(80, 170, 1120, 490))]; break;
      case 'section': {
        if (ta) { const x = Math.round(ta.x * W), w = Math.round(ta.w * W); specs = [title(Object.assign(R(0, 210, 0, 180), { x, w }), { size: 48, valign: 'bottom' }), { type: 'text', ph: 'body', role: 'body', size: 22, valign: 'top', placeholder: 'Click to add text', ...Object.assign(R(0, 410, 0, 90), { x, w }) }]; break; }
        specs = [title(R(110, 210, 1060, 180), { size: 48, valign: 'bottom' }), { type: 'text', ph: 'body', role: 'body', size: 22, valign: 'top', placeholder: 'Click to add text', ...R(110, 410, 1060, 90) }]; break;
      }
      case 'twoContent': specs = [title(R(80, 40, 1120, 110)), body(R(80, 170, 545, 490), { size: 22 }), body(R(655, 170, 545, 490), { size: 22 })]; break;
      case 'comparison': specs = [title(R(80, 40, 1120, 100)),
        { type: 'text', ph: 'body', role: 'body', size: 24, bold: true, valign: 'bottom', placeholder: 'Click to add text', ...R(80, 160, 545, 60) }, body(R(80, 230, 545, 430), { size: 20 }),
        { type: 'text', ph: 'body', role: 'body', size: 24, bold: true, valign: 'bottom', placeholder: 'Click to add text', ...R(655, 160, 545, 60) }, body(R(655, 230, 545, 430), { size: 20 })]; break;
      case 'titleOnly': specs = [title(R(80, 40, 1120, 110))]; break;
      case 'pictureCaption': specs = [title(R(80, 90, 440, 160), { size: 32, valign: 'bottom' }), { type: 'image', ph: 'pic', ...R(580, 70, 620, 580) }, { type: 'text', ph: 'body', role: 'body', size: 18, valign: 'top', placeholder: 'Click to add text', ...R(80, 270, 440, 380) }]; break;
      default: specs = [];
    }
    // content layouts squeezed into a theme content area (fractions of the slide), if the theme reserves a side band
    if (areas.content && !['title', 'section', 'blank'].includes(layout)) {
      const E0 = { x: 80 * sx, y: 40 * sy, w: 1120 * sx, h: 620 * sy }, A = areas.content, E1 = { x: (A.x ?? .0625) * W, y: (A.y ?? .0556) * H, w: (A.w ?? .875) * W, h: (A.h ?? .861) * H };
      specs.forEach(s => { s.x = Math.round(E1.x + (s.x - E0.x) * E1.w / E0.w); s.w = Math.round(s.w * E1.w / E0.w); s.y = Math.round(E1.y + (s.y - E0.y) * E1.h / E0.h); s.h = Math.round(s.h * E1.h / E0.h); });
    }
    // slide master overrides (View ▸ Slide Master)
    const m = !opts.noMaster && pres && pres.master;
    if (m) {
      const keys = SL.phKeys(specs); const lo = (m.layouts || {})[layout] || {};
      specs.forEach((s, i) => {
        const role = s.ph === 'title' ? 'title' : 'body';
        const ms = m[role]; if (ms && ms.size && s.type === 'text' && !['title', 'section'].includes(layout) && !(s.ph === 'body' && s.bold)) s.size = Math.round(ms.size * (s.size / (s.ph === 'title' ? 40 : 24)) * 2) / 2;
        const o = lo[keys[i]]; if (!o) return;
        if (o.x != null) { s.x = Math.round(o.x * W); s.y = Math.round(o.y * H); s.w = Math.round(o.w * W); s.h = Math.round(o.h * H); }
        if (o.size) s.size = o.size; if (o.align) s.align = o.align;
      });
    }
    return specs;
  };
  SL.decoKind = (layout) => layout === 'title' ? 'title' : layout === 'section' ? 'section' : 'content';

  /* ---------- fonts: make sure the bundled fonts a presentation uses are loaded before drawing thumbnails, shows and exports ---------- */
  SL.usedFonts = (pres) => {
    const set = new Set([SL.fontOf('title', pres), SL.fontOf('body', pres)]);
    const scan = (html) => String(html || '').replace(/font-family:\s*([^;"]+)/g, (m, f) => { set.add(f.split(',')[0].replace(/['"&]|quot;/g, '').trim()); return m; });
    (pres.slides || []).forEach(s => [...(s.elements || []), ...(s.layer || [])].forEach(e => { if (e.font) set.add(e.font); scan(e.html); if (e.rows) e.rows.forEach(r => r.forEach(c => scan(c.html))); }));
    if (pres.master) ['title', 'body'].forEach(k => pres.master[k] && pres.master[k].font && set.add(pres.master[k].font));
    return [...set].filter(Boolean);
  };
  const loaded = new Set();
  SL.ensureFonts = async (list) => {
    if (!document.fonts) return;
    if (!Array.isArray(list)) list = SL.usedFonts(list || SL.pres);
    const jobs = [];
    for (const f of list) {
      if (loaded.has(f)) continue; loaded.add(f);
      const reg = (window.BTS_FONTS || []).find(x => x.name === f);
      const fam = `"${f.replace(/"/g, '')}"`;
      for (const w of [400, 700]) { jobs.push(document.fonts.load(`${w} 24px ${fam}`).catch(() => {})); jobs.push(document.fonts.load(`italic ${w} 24px ${fam}`).catch(() => {})); }
      jobs.push(document.fonts.load(`400 24px ${fam}`, 'अआक').catch(() => {}));
      if (reg && reg.cat === 'indian') jobs.push(document.fonts.load(`700 24px ${fam}`, 'अआक').catch(() => {}));
    }
    ['Noto Sans Devanagari', 'Noto Serif Devanagari'].forEach(f => { if (!loaded.has(f)) { loaded.add(f); jobs.push(document.fonts.load(`400 24px "${f}"`, 'अआक').catch(() => {}), document.fonts.load(`700 24px "${f}"`, 'अआक').catch(() => {})); } });
    await Promise.all(jobs);
    try { await document.fonts.ready; } catch { }
  };
})();
