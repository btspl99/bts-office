/* BTS PDF — shared state and small helpers used by every module */
import * as pdfjsLib from '../vendor/pdf.min.mjs';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('../vendor/pdf.worker.min.mjs', import.meta.url).href;

export { pdfjsLib };
export const CSS_UNITS = 96 / 72;            // 100% zoom = real size on a 96-dpi screen
export const $ = (s, r = document) => r.querySelector(s);
export const $$ = (s, r = document) => [...r.querySelectorAll(s)];
export const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
export const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

/* One mutable state object. Modules read and write it directly. */
export const S = {
  doc: null,            // PDFDocumentProxy
  loadingTask: null,
  bytes: null,          // Uint8Array of the current document (original, or merged after "Combine")
  modified: false,      // structure changed (combine) — save must rewrite even without annotations
  encrypted: false,
  pages: [],            // [{ i, proxy, vp1, el, canvas, text, links, svg, ... }]
  scale: CSS_UNITS,     // CSS px per PDF point
  rot: 0,               // user view rotation (0/90/180/270), added to each page's own /Rotate
  mode: 'continuous',   // continuous | single | spread
  fit: 'auto',          // auto | width | page | null
  cur: 0,               // current page index (0-based)
  tool: 'text',         // text | hand | highlight | underline | strike | note | textbox | pen | rect | ellipse | arrow | eraser
  colors: { highlight: '#FFD400', underline: '#1A56DB', strike: '#D93025', note: '#FFD400', textbox: '#172033', pen: '#D93025', rect: '#D93025', ellipse: '#1A56DB', arrow: '#D93025' },
  width: 2,
  anns: [],             // annotations: { id, p (page index), type, color, ... } in PDF user-space coordinates
  sel: null,            // selected annotation id
  undo: [], redo: [],
  textCache: new Map(), // page index -> textContent
  meta: null,           // cached document properties
  fileName: '', fileSize: 0,
  darkRead: false,
};

export const MARKUP = new Set(['highlight', 'underline', 'strike']);
export const DRAW_TOOLS = new Set(['pen', 'rect', 'ellipse', 'arrow', 'note', 'textbox', 'eraser']);

let nextId = 1;
export const newId = () => 'a' + (nextId++) + '_' + Date.now().toString(36);

/** Base viewport (scale 1 = PDF points) honouring the page's /Rotate plus the view rotation. */
export function baseViewport(proxy, rot = S.rot) {
  return proxy.getViewport({ scale: 1, rotation: (proxy.rotate + rot) % 360 });
}

/** Client (screen) point -> PDF user-space point for page p. */
export function clientToPdf(p, cx, cy) {
  const r = p.el.getBoundingClientRect();
  const vx = (cx - r.left) / r.width * p.vp1.width;
  const vy = (cy - r.top) / r.height * p.vp1.height;
  return p.vp1.convertToPdfPoint(vx, vy);
}
/** PDF user-space point -> position in % of the page box (for HTML overlays). */
export function pdfToPct(p, x, y) {
  const [vx, vy] = p.vp1.convertToViewportPoint(x, y);
  return [vx / p.vp1.width * 100, vy / p.vp1.height * 100];
}

/** Lazily load pdf-lib (UMD → window.PDFLib). */
let pdfLibPromise = null;
export function loadPdfLib() {
  if (window.PDFLib) return Promise.resolve(window.PDFLib);
  if (!pdfLibPromise) pdfLibPromise = new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = new URL('../vendor/pdf-lib.min.js', import.meta.url).href;
    s.onload = () => res(window.PDFLib);
    s.onerror = () => { pdfLibPromise = null; rej(new Error('Could not load pdf-lib')); };
    document.head.appendChild(s);
  });
  return pdfLibPromise;
}

export function hexToRgb01(hex) {
  const h = hex.replace('#', '');
  const n = parseInt(h.length === 3 ? h.split('').map(c => c + c).join('') : h, 16);
  return [(n >> 16 & 255) / 255, (n >> 8 & 255) / 255, (n & 255) / 255];
}

export const isTyping = (el) => !!el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.tagName === 'SELECT' || el.isContentEditable);

/* Note callout layout (shared by on-screen print preview and pdf-lib burn-in). */
export const NOTE = { icon: 18, fs: 9, lh: 11.5, pad: 5, maxW: 170 };
export function wrapText(text, maxW, measure) {
  const out = [];
  for (const para of String(text || '').split('\n')) {
    const words = para.split(/(\s+)/); let line = '';
    for (const w of words) {
      const t = line + w;
      if (measure(t.trimEnd()) > maxW && line.trim()) { out.push(line.trimEnd()); line = w.trimStart(); }
      else line = t;
      while (measure(line) > maxW && line.length > 1) {        // very long word: hard-break
        let k = line.length - 1; while (k > 1 && measure(line.slice(0, k)) > maxW) k--;
        out.push(line.slice(0, k)); line = line.slice(k);
      }
    }
    out.push(line.trimEnd());
  }
  return out;
}
let mctx = null;
export function measureHelv(size) {
  mctx ||= document.createElement('canvas').getContext('2d');
  mctx.font = `${size}px Helvetica, Arial, "Liberation Sans", sans-serif`;
  return (t) => mctx.measureText(t).width;
}
