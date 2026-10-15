#!/usr/bin/env node
/* Builds the bundled font library from @fontsource packages (SIL Open Font License 1.1 or Apache 2.0).
 *   node tools/build-fonts.js
 * Output:
 *   app/fonts/lib/<id>/*.woff2      the font files (latin, latin-ext and devanagari subsets only)
 *   app/common/fonts.css            @font-face rules, plus local-first aliases for Office fonts
 *   app/common/fonts.js             window.BTS_FONTS registry used by every app's font picker
 * Edit FAMILIES below to add or remove fonts, then run this script again.
 */
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const nm = path.join(root, 'node_modules', '@fontsource');
const SUBSETS = ['latin', 'latin-ext', 'devanagari'];

// cat: sans | serif | display | mono | indian | script.  alias: Office font this one stands in for (same metrics).
const FAMILIES = [
  // Office-compatible (metric-identical replacements, used automatically when the Office font isn't installed)
  { id: 'carlito', name: 'Carlito', cat: 'sans', alias: 'Calibri', w: [400, 700], i: true },
  { id: 'caladea', name: 'Caladea', cat: 'serif', alias: 'Cambria', w: [400, 700], i: true },
  { id: 'arimo', name: 'Arimo', cat: 'sans', alias: 'Arial', w: [400, 700], i: true },
  { id: 'tinos', name: 'Tinos', cat: 'serif', alias: 'Times New Roman', w: [400, 700], i: true },
  { id: 'cousine', name: 'Cousine', cat: 'mono', alias: 'Courier New', w: [400, 700], i: true },
  { id: 'gelasio', name: 'Gelasio', cat: 'serif', alias: 'Georgia', w: [400, 700], i: true },
  // Sans serif
  { id: 'inter', name: 'Inter', cat: 'sans', w: [400, 500, 600, 700], i: true },
  { id: 'roboto', name: 'Roboto', cat: 'sans', w: [400, 500, 700], i: true },
  { id: 'open-sans', name: 'Open Sans', cat: 'sans', w: [400, 600, 700], i: true },
  { id: 'lato', name: 'Lato', cat: 'sans', w: [400, 700], i: true },
  { id: 'source-sans-3', name: 'Source Sans 3', cat: 'sans', w: [400, 600, 700], i: true },
  { id: 'noto-sans', name: 'Noto Sans', cat: 'sans', w: [400, 700], i: true },
  { id: 'ibm-plex-sans', name: 'IBM Plex Sans', cat: 'sans', w: [400, 500, 700], i: true },
  { id: 'pt-sans', name: 'PT Sans', cat: 'sans', w: [400, 700], i: true },
  { id: 'work-sans', name: 'Work Sans', cat: 'sans', w: [400, 600, 700], i: true },
  { id: 'nunito', name: 'Nunito', cat: 'sans', w: [400, 700], i: true },
  { id: 'dm-sans', name: 'DM Sans', cat: 'sans', w: [400, 500, 700], i: true },
  { id: 'manrope', name: 'Manrope', cat: 'sans', w: [400, 600, 700, 800], i: false },
  { id: 'montserrat', name: 'Montserrat', cat: 'sans', w: [400, 600, 700, 800], i: true },
  { id: 'poppins', name: 'Poppins', cat: 'sans', w: [400, 500, 600, 700], i: true },
  { id: 'raleway', name: 'Raleway', cat: 'sans', w: [400, 600, 700], i: true },
  { id: 'outfit', name: 'Outfit', cat: 'sans', w: [400, 600, 700], i: false },
  { id: 'sora', name: 'Sora', cat: 'sans', w: [400, 600, 700], i: false },
  // Serif
  { id: 'source-serif-4', name: 'Source Serif 4', cat: 'serif', w: [400, 600, 700], i: true },
  { id: 'merriweather', name: 'Merriweather', cat: 'serif', w: [400, 700], i: true },
  { id: 'lora', name: 'Lora', cat: 'serif', w: [400, 600, 700], i: true },
  { id: 'noto-serif', name: 'Noto Serif', cat: 'serif', w: [400, 700], i: true },
  { id: 'ibm-plex-serif', name: 'IBM Plex Serif', cat: 'serif', w: [400, 600, 700], i: true },
  { id: 'pt-serif', name: 'PT Serif', cat: 'serif', w: [400, 700], i: true },
  { id: 'eb-garamond', name: 'EB Garamond', cat: 'serif', w: [400, 600, 700], i: true },
  { id: 'libre-baskerville', name: 'Libre Baskerville', cat: 'serif', w: [400, 700], i: true },
  { id: 'crimson-pro', name: 'Crimson Pro', cat: 'serif', w: [400, 600, 700], i: true },
  { id: 'roboto-slab', name: 'Roboto Slab', cat: 'serif', w: [400, 600, 700], i: false },
  // Display / headings
  { id: 'playfair-display', name: 'Playfair Display', cat: 'display', w: [400, 700, 800], i: true },
  { id: 'cormorant-garamond', name: 'Cormorant Garamond', cat: 'display', w: [400, 600, 700], i: true },
  { id: 'oswald', name: 'Oswald', cat: 'display', w: [400, 600, 700], i: false },
  { id: 'bebas-neue', name: 'Bebas Neue', cat: 'display', w: [400], i: false },
  // Monospace
  { id: 'roboto-mono', name: 'Roboto Mono', cat: 'mono', w: [400, 700], i: true },
  { id: 'jetbrains-mono', name: 'JetBrains Mono', cat: 'mono', w: [400, 700], i: true },
  { id: 'fira-code', name: 'Fira Code', cat: 'mono', w: [400, 700], i: false },
  { id: 'ibm-plex-mono', name: 'IBM Plex Mono', cat: 'mono', w: [400, 700], i: true },
  // Indian scripts (Hindi, Marathi, Sanskrit …)
  { id: 'noto-sans-devanagari', name: 'Noto Sans Devanagari', cat: 'indian', w: [400, 600, 700], i: false },
  { id: 'noto-serif-devanagari', name: 'Noto Serif Devanagari', cat: 'indian', w: [400, 700], i: false },
  { id: 'mukta', name: 'Mukta', cat: 'indian', w: [400, 600, 700], i: false },
  { id: 'hind', name: 'Hind', cat: 'indian', w: [400, 600, 700], i: false },
  { id: 'tiro-devanagari-hindi', name: 'Tiro Devanagari Hindi', cat: 'indian', w: [400], i: true },
  { id: 'baloo-2', name: 'Baloo 2', cat: 'indian', w: [400, 600, 700], i: false },
  { id: 'rozha-one', name: 'Rozha One', cat: 'indian', w: [400], i: false },
  { id: 'yatra-one', name: 'Yatra One', cat: 'indian', w: [400], i: false },
  // Handwriting
  { id: 'kalam', name: 'Kalam', cat: 'script', w: [400, 700], i: false },
];

// Fonts that are usually installed with Windows/Office — offered in pickers, used when present.
const SYSTEM = [
  ['Calibri', 'sans'], ['Arial', 'sans'], ['Segoe UI', 'sans'], ['Verdana', 'sans'], ['Tahoma', 'sans'], ['Trebuchet MS', 'sans'], ['Century Gothic', 'sans'], ['Aptos', 'sans'],
  ['Cambria', 'serif'], ['Times New Roman', 'serif'], ['Georgia', 'serif'], ['Garamond', 'serif'], ['Book Antiqua', 'serif'], ['Palatino Linotype', 'serif'],
  ['Courier New', 'mono'], ['Consolas', 'mono'], ['Comic Sans MS', 'script'], ['Nirmala UI', 'indian'], ['Mangal', 'indian'], ['Kokila', 'indian'],
];

const outLib = path.join(root, 'app', 'fonts', 'lib');
fs.rmSync(outLib, { recursive: true, force: true });
let css = '/* BTS Office font library — generated by tools/build-fonts.js. Fonts: SIL OFL 1.1 or Apache 2.0 (see app/fonts/lib/LICENSES.md). */\n';
const registry = [];
const licenses = new Set();
let bytes = 0, files = 0;

for (const f of FAMILIES) {
  const pkg = path.join(nm, f.id);
  if (!fs.existsSync(pkg)) { console.warn('skip (not installed):', f.id); continue; }
  const dest = path.join(outLib, f.id); fs.mkdirSync(dest, { recursive: true });
  const styles = []; for (const w of f.w) { styles.push([w, 'normal']); if (f.i) styles.push([w, 'italic']); }
  const faces = [];
  for (const [w, st] of styles) {
    const cssFile = path.join(pkg, `${w}${st === 'italic' ? '-italic' : ''}.css`);
    if (!fs.existsSync(cssFile)) continue;
    const src = fs.readFileSync(cssFile, 'utf8');
    for (const block of src.split('@font-face').slice(1)) {
      const file = (block.match(/url\(\.\/files\/([^)]+?\.woff2)\)/) || [])[1]; if (!file) continue;
      const subset = SUBSETS.find(s => file.includes(`-${s}-${w}-`)); if (!subset) continue;
      fs.copyFileSync(path.join(pkg, 'files', file), path.join(dest, file));
      bytes += fs.statSync(path.join(dest, file)).size; files++;
      const range = (block.match(/unicode-range:\s*([^;]+);/) || [])[1];
      faces.push({ file, w, st, range });
    }
  }
  const lic = JSON.parse(fs.readFileSync(path.join(pkg, 'package.json'), 'utf8')).license || 'OFL-1.1';
  licenses.add(`- **${f.name}** — ${lic}${f.alias ? ` (metric-compatible with ${f.alias})` : ''}`);
  for (const fc of faces) {
    css += `@font-face{font-family:'${f.name}';font-style:${fc.st};font-weight:${fc.w};font-display:swap;src:url(../fonts/lib/${f.id}/${fc.file}) format('woff2');${fc.range ? `unicode-range:${fc.range};` : ''}}\n`;
    // Local-first alias: "Calibri" uses the real Calibri when installed, otherwise Carlito (same metrics, so layout matches).
    if (f.alias) css += `@font-face{font-family:'${f.alias}';font-style:${fc.st};font-weight:${fc.w};src:local('${f.alias}${fc.w >= 700 ? ' Bold' : ''}${fc.st === 'italic' ? ' Italic' : ''}'),local('${f.alias}'),url(../fonts/lib/${f.id}/${fc.file}) format('woff2');${fc.range ? `unicode-range:${fc.range};` : ''}}\n`;
  }
  if (faces.length) registry.push({ name: f.name, cat: f.cat, bundled: true, weights: f.w, italic: !!f.i, alias: f.alias || null });
}
for (const [name, cat] of SYSTEM) if (!registry.some(r => r.name === name)) registry.push({ name, cat, bundled: false, system: true, alias: null });

const CATS = { sans: 'Sans serif', serif: 'Serif', display: 'Display and headings', mono: 'Monospace', indian: 'Indian scripts', script: 'Handwriting' };
const FALLBACK = { sans: 'sans-serif', serif: 'serif', display: 'serif', mono: 'monospace', indian: 'sans-serif', script: 'cursive' };
const js = `/* Generated by tools/build-fonts.js — the font registry shared by Docs, Sheets and Slides. */
window.BTS_FONTS = ${JSON.stringify(registry)};
window.BTS_FONT_CATEGORIES = ${JSON.stringify(CATS)};
window.BTS_FONT_FALLBACK = ${JSON.stringify(FALLBACK)};
/* stack('Inter') → "'Inter', sans-serif" */
window.btsFontStack = function (name) { var f = (window.BTS_FONTS || []).find(function (x) { return x.name === name; }); return "'" + name + "', " + (f ? window.BTS_FONT_FALLBACK[f.cat] : 'sans-serif'); };
/* Is a system (non-bundled) font installed on this computer? */
window.btsFontAvailable = function (name) { try { return document.fonts.check('12px "' + name + '"'); } catch (e) { return true; } };
/* Recently used fonts (per app), most recent first. */
window.btsRecentFonts = function (add) { var k = 'bts-recent-fonts'; var l = []; try { l = JSON.parse(localStorage.getItem(k) || '[]'); } catch (e) {} if (add) { l = [add].concat(l.filter(function (x) { return x !== add; })).slice(0, 6); try { localStorage.setItem(k, JSON.stringify(l)); } catch (e) {} } return l; };
`;
fs.writeFileSync(path.join(root, 'app', 'common', 'fonts.css'), css);
fs.writeFileSync(path.join(root, 'app', 'common', 'fonts.js'), js);
fs.writeFileSync(path.join(outLib, 'LICENSES.md'), '# Bundled fonts\n\nAll fonts are free to use, embed and redistribute. Sources: https://fontsource.org (Google Fonts).\n\n' + [...licenses].join('\n') + '\n');
console.log(`✔ ${registry.filter(r => r.bundled).length} bundled families, ${files} files, ${(bytes / 1048576).toFixed(1)} MB; ${registry.length} fonts in registry`);
