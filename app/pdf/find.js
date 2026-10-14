/* BTS PDF — find in document: searches every page's text, highlights matches in the text layer */
import { S, $, $$, esc } from './state.js';
import { hooks, getText, goTo } from './viewer.js';

const F = { q: '', caseSens: false, matches: [], idx: -1, token: 0, pageStr: new Map(), done: false };
export const findState = F;

/** Page text as one string + map from string offset to [itemIndex, charOffset]. */
async function pageIndex(i) {
  if (F.pageStr.has(i)) return F.pageStr.get(i);
  const tc = await getText(i);
  let str = ''; const map = [];
  let k = 0;   // index among items that have .str (== TextLayer textDivs index)
  for (const it of tc.items) {
    if (it.str === undefined) continue;
    for (let c = 0; c < it.str.length; c++) { map.push([k, c]); }
    str += it.str;
    if (it.hasEOL) { str += ' '; map.push([-1, 0]); }
    else if (it.str && !/\s$/.test(it.str)) { /* items on a line are usually adjacent; no gap */ }
    k++;
  }
  const entry = { str, low: str.toLowerCase(), map };
  F.pageStr.set(i, entry);
  return entry;
}

const norm = (s) => s.replace(/\s+/g, ' ');

export async function runFind(q, { caseSens = F.caseSens } = {}) {
  q = norm(q || '').trim();
  const token = ++F.token;
  clearMarks();
  F.q = q; F.caseSens = caseSens; F.matches = []; F.idx = -1; F.done = false;
  updateCount();
  if (!q || !S.doc) return;
  const needle = caseSens ? q : q.toLowerCase();
  const start = S.cur; const n = S.pages.length;
  // search from the current page so the first hit is near the reader
  const order = [...Array(n).keys()].map(k => (start + k) % n);
  const found = new Map();
  let jumped = false;
  for (let k = 0; k < n; k++) {
    const i = order[k];
    const P = await pageIndex(i);
    if (token !== F.token) return;
    // whitespace-collapsed view of the page text, with a map back to raw offsets
    const { s, back } = collapse(caseSens ? P.str : P.low);
    const list = [];
    let pos = s.indexOf(needle);
    while (pos !== -1) { list.push({ p: i, start: back[pos], end: back[pos + needle.length - 1] + 1 }); pos = s.indexOf(needle, pos + Math.max(1, needle.length)); }
    found.set(i, list);
    if (!jumped && list.length) {
      jumped = true;
      F.matches = list.slice(); F.idx = 0; show(); updateCount(true);
    }
    if (k % 8 === 7) { await new Promise(r => setTimeout(r)); if (token !== F.token) return; }
  }
  // final ordered list by page, keep the current match
  const cur = F.matches[F.idx];
  F.matches = [...found.keys()].sort((a, b) => a - b).flatMap(i => found.get(i));
  F.idx = cur ? F.matches.findIndex(m => m.p === cur.p && m.start === cur.start) : (F.matches.length ? 0 : -1);
  F.done = true;
  updateCount();
  paintAll();
  if (!F.matches.length) $('#findbar')?.classList.add('nores');
}
function collapse(str) {
  let s = ''; const back = []; let prevSpace = false;
  for (let i = 0; i < str.length; i++) {
    const ch = str[i]; const sp = /\s/.test(ch);
    if (sp && prevSpace) continue;
    s += sp ? ' ' : ch; back.push(i); prevSpace = sp;
  }
  return { s, back };
}

function updateCount(partial = false) {
  const el = $('#find-count'); if (!el) return;
  $('#findbar')?.classList.remove('nores');
  if (!F.q) { el.textContent = ''; return; }
  if (!F.matches.length) { el.textContent = F.done ? 'No results' : 'Searching…'; return; }
  el.textContent = partial && !F.done ? `${F.idx + 1} of ${F.matches.length}+` : `${F.idx + 1} of ${F.matches.length}`;
}

export function step(dir) {
  if (!F.matches.length) { if (F.q && F.done) BTS.toast('No matches for “' + F.q + '”'); return; }
  F.idx = (F.idx + dir + F.matches.length) % F.matches.length;
  updateCount(); show();
}

/** Scroll to the current match and mark it as selected. */
function show() {
  const m = F.matches[F.idx]; if (!m) return;
  paintAll();
  const p = S.pages[m.p];
  const el = p && $('.textLayer .highlight.selected', p.el);
  if (el) { scrollToEl(el); return; }
  // page not rendered yet — jump to it, the textReady hook will finish the job
  F.pendingScroll = true;
  goTo(m.p);
}
function scrollToEl(el) {
  const v = $('#viewer'); const r = el.getBoundingClientRect(), vr = v.getBoundingClientRect();
  if (r.top < vr.top + 40 || r.bottom > vr.bottom - 40) v.scrollTop += r.top - vr.top - vr.height / 3;
  if (r.left < vr.left || r.right > vr.right) v.scrollLeft += r.left - vr.left - vr.width / 3;
}

/* ---------------- painting into text layers ---------------- */
function clearMarks(p) {
  const pages = p ? [p] : S.pages;
  for (const q of pages) {
    if (!q.text?.layer || !q.marked) continue;
    const divs = q.text.layer.textDivs, strs = q.text.layer.textContentItemsStr;
    for (const k of q.marked) if (divs[k]) divs[k].textContent = strs[k];
    q.marked = null;
  }
}
function paint(p) {
  clearMarks(p);
  if (!F.q || !p.text?.layer) return;
  const P = F.pageStr.get(p.i); if (!P) return;
  const list = F.matches.filter(m => m.p === p.i); if (!list.length) return;
  const cur = F.matches[F.idx];
  const divs = p.text.layer.textDivs, strs = p.text.layer.textContentItemsStr;
  const per = new Map();   // div index -> [[s,e,selected]]
  for (const m of list) {
    const sel = m === cur;
    let i = m.start;
    while (i < m.end) {
      const [k, c] = P.map[i] || [-1, 0];
      if (k < 0) { i++; continue; }
      let e = i; while (e < m.end && P.map[e] && P.map[e][0] === k) e++;
      const endC = P.map[e - 1][1] + 1;
      if (!per.has(k)) per.set(k, []);
      per.get(k).push([c, endC, sel]);
      i = e;
    }
  }
  p.marked = [...per.keys()];
  for (const [k, ranges] of per) {
    const div = divs[k]; if (!div) continue;
    const s = strs[k]; ranges.sort((a, b) => a[0] - b[0]);
    let html = '', pos = 0;
    for (const [a, b, sel] of ranges) {
      if (a < pos) continue;
      html += esc(s.slice(pos, a)) + `<span class="highlight${sel ? ' selected' : ''}">${esc(s.slice(a, b))}</span>`;
      pos = b;
    }
    html += esc(s.slice(pos));
    div.innerHTML = html;
  }
}
function paintAll() { S.pages.forEach(p => { if (p.text?.layer) paint(p); }); }

hooks.textReady.push((p) => {
  p.marked = null;
  if (!F.q) return;
  paint(p);
  const m = F.matches[F.idx];
  if (F.pendingScroll && m && m.p === p.i) {
    F.pendingScroll = false;
    const el = $('.textLayer .highlight.selected', p.el); if (el) scrollToEl(el);
  }
});

/* ---------------- find bar UI ---------------- */
export function openFind(text) {
  if (!S.doc) return;
  const bar = $('#findbar'); bar.classList.remove('hidden');
  const inp = $('#find-input');
  if (text != null) inp.value = text;
  else { const sel = getSelection().toString().trim(); if (sel && sel.length < 80 && !sel.includes('\n')) inp.value = sel; }
  inp.focus(); inp.select();
  if (inp.value && inp.value !== F.q) runFind(inp.value);
}
export function closeFind() {
  $('#findbar').classList.add('hidden');
  F.token++; F.q = ''; F.matches = []; clearMarks(); updateCount();
  $('#viewer').focus({ preventScroll: true });
}
export function resetFind() { F.token++; F.q = ''; F.matches = []; F.idx = -1; F.pageStr.clear(); $('#findbar')?.classList.add('hidden'); updateCount(); }
export function toggleCase() {
  F.caseSens = !F.caseSens; $('#find-case').classList.toggle('on', F.caseSens);
  if (F.q || $('#find-input').value) runFind($('#find-input').value);
}

export function initFind() {
  const inp = $('#find-input');
  let t = 0;
  inp.addEventListener('input', () => { clearTimeout(t); t = setTimeout(() => runFind(inp.value), 220); });
  inp.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      e.preventDefault(); clearTimeout(t);
      if (norm(inp.value).trim() !== F.q) runFind(inp.value); else step(e.shiftKey ? -1 : 1);
    } else if (e.key === 'Escape') { e.preventDefault(); closeFind(); }
  });
}
