/* BTS Slides 1.2 — Review: threaded comments (anchored to a slide or an object, replies, resolve), accessibility
 * checker (alt text, slide titles, contrast, table headers, reading order) and the Alt Text dialog.
 * Comment model: slide.comments = [{ id, author, text, time, x, y, elId?, resolved?, replies:[{id, author, text, time}] }]
 */
(function () {
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const esc = (s) => BTS.esc(s);
  const X = window.SLX = window.SLX || { cmd: {}, ribbon: [] };
  const C = () => window.CMD;
  const P = () => SL.pres, S = () => SL.slide();

  const me = () => BTS.userName || 'You';
  const initials = (n) => String(n || '?').trim().split(/\s+/).map(w => w[0]).slice(0, 2).join('').toUpperCase() || '?';
  const hue = (n) => { let h = 0; for (const ch of String(n)) h = (h * 31 + ch.charCodeAt(0)) % 360; return h; };
  const avatar = (n) => `<span class="cm-av" style="background:hsl(${hue(n)} 55% 45%)" title="${esc(n)}">${esc(initials(n))}</span>`;
  const when = (t) => { const d = Date.now() - t; if (d < 60e3) return 'just now'; if (d < 3600e3) return Math.floor(d / 60e3) + ' min ago'; if (d < 86400e3) return Math.floor(d / 3600e3) + ' h ago'; return new Date(t).toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' }); };
  const normalize = (c) => { if (!c.id) c.id = SL.uid('c'); if (!c.replies) c.replies = []; return c; };
  X.normalizeComments = (pres) => (pres.slides || []).forEach(s => (s.comments || []).forEach(normalize));

  /* ---------------- comments ---------------- */
  const st = X.cmtState = { scope: 'slide', showResolved: true, composing: false, active: null, editing: null };
  X.cmd['new-comment'] = () => {
    if (E.view === 'master') return BTS.toast('Close Slide Master view to add comments');
    st.composing = true; st.anchor = E.primary() ? E.primary().id : null; st.scope = 'slide';
    if (E.pane !== 'comments') E.openPane('comments'); else E.refreshPane('force');
    setTimeout(() => { const t = $('#cm-new'); t && t.focus(); }, 0);
  };
  X.cmd['comments-pane'] = () => E.openPane('comments');
  const allThreads = () => P().slides.flatMap((s, si) => (s.comments || []).map(c => ({ c: normalize(c), si })));
  X.cmd['next-comment'] = X.cmd['prev-comment'] = (el, ev, dir) => {
    const list = allThreads(); if (!list.length) return BTS.toast('There are no comments in this presentation');
    const d = dir || ((el && el.dataset && el.dataset.cmd === 'prev-comment') ? -1 : 1);
    let i = list.findIndex(t => t.c.id === st.active); if (i < 0) i = d > 0 ? list.findIndex(t => t.si >= SL.cur) - 1 : list.findIndex(t => t.si >= SL.cur);
    if (i < -1) i = -1;
    const j = ((i + d) % list.length + list.length) % list.length; const t = list[j];
    st.active = t.c.id; if (t.si !== SL.cur) E.goto(t.si);
    if (t.c.elId && SL.findEl(t.c.elId)) E.select([t.c.elId]);
    if (E.pane !== 'comments') E.openPane('comments'); else E.refreshPane('force');
    setTimeout(() => { const card = $(`.cm-card[data-id="${t.c.id}"]`); card && card.scrollIntoView({ block: 'nearest' }); }, 0);
  };
  X.cmd['delete-comments'] = (el) => BTS.menu(el, [
    { label: 'Delete selected comment', icon: 'message-square-x', disabled: !st.active || !findThread(st.active), action: () => delThread(st.active) },
    { label: 'Delete all comments on this slide', icon: 'trash-2', disabled: !(S().comments || []).length, action: () => { S().comments = []; SL.commit(); E.renderAll(); } },
    { label: 'Delete all comments in this presentation', icon: 'trash', disabled: !allThreads().length, action: async () => { if (await BTS.confirm('Delete all comments', 'Delete every comment on every slide?', 'Delete all')) { P().slides.forEach(s => s.comments = []); SL.commit(); E.renderAll(); } } },
  ]);
  function findThread(id) { for (const s of P().slides) { const c = (s.comments || []).find(x => x.id === id); if (c) return { c, s }; } return null; }
  function delThread(id) { const t = findThread(id); if (!t) return; t.s.comments = t.s.comments.filter(x => x !== t.c); if (st.active === id) st.active = null; SL.commit(); E.renderAll(); }
  function anchorLabel(c, slide) { if (!c.elId) return 'Slide'; const el = (slide.elements || []).find(e => e.id === c.elId); return el ? C().elName(el) : 'Object (deleted)'; }
  X.addComment = (text, anchorId) => {
    const s = S(); const el = anchorId ? SL.findEl(anchorId) : null;
    const n = (s.comments || []).length;
    const at = el ? { x: el.x + el.w - 8, y: Math.max(0, el.y - 8) } : { x: 20 + n * 30, y: 20 };
    const c = normalize({ author: me(), text, time: Date.now(), x: at.x, y: at.y });
    if (el) c.elId = el.id;
    (s.comments = s.comments || []).push(c); st.active = c.id;
    SL.commit(); E.renderAll(); return c;
  };
  const thread = (c, si, slide) => {
    const own = c.author === me();
    const edit = st.editing === c.id;
    return `<div class="cm-card ${c.resolved ? 'resolved' : ''} ${st.active === c.id ? 'active' : ''}" data-id="${c.id}" data-si="${si}">
      <div class="cm-head">${avatar(c.author)}<div class="cm-who"><b>${esc(c.author)}</b><small>${when(c.time)}${st.scope === 'all' ? ` · Slide ${si + 1}` : ''}</small></div>
        <div class="cm-acts"><button class="rb" data-act="resolve" title="${c.resolved ? 'Reopen this thread' : 'Resolve this thread'}">${icon(c.resolved ? 'rotate-ccw' : 'circle-check', 15)}</button>${own ? `<button class="rb" data-act="edit" title="Edit comment">${icon('pencil', 14)}</button>` : ''}<button class="rb" data-act="del" title="Delete thread">${icon('trash-2', 14)}</button></div></div>
      <div class="cm-anchor">${icon(c.elId ? 'square-dashed-mouse-pointer' : 'presentation', 12)} ${esc(anchorLabel(c, slide))}${c.resolved ? ' · <b>Resolved</b>' : ''}</div>
      ${edit ? `<textarea class="cm-edit" data-act-k="edit-text">${esc(c.text)}</textarea><div class="cm-row"><button class="btn sm" data-act="edit-cancel">Cancel</button><button class="btn sm primary" data-act="edit-save">Save</button></div>` : `<div class="cm-text">${esc(c.text)}</div>`}
      ${(c.replies || []).map(r => `<div class="cm-reply">${avatar(r.author)}<div><div class="cm-who"><b>${esc(r.author)}</b><small>${when(r.time)}</small>${r.author === me() ? `<button class="rb cm-rdel" data-act="rdel" data-rid="${r.id}" title="Delete reply">${icon('x', 12)}</button>` : ''}</div><div class="cm-text">${esc(r.text)}</div></div></div>`).join('')}
      ${c.resolved ? '' : `<div class="cm-row cm-replybox"><input class="cm-rin" placeholder="Reply…" title="Write a reply and press Enter"><button class="btn sm" data-act="reply" title="Post reply">${icon('send', 13)}</button></div>`}
    </div>`;
  };
  C().PANES.comments = {
    title: 'Comments',
    html: () => {
      const list = st.scope === 'all' ? allThreads() : (S().comments || []).map(c => ({ c: normalize(c), si: SL.cur }));
      const shown = list.filter(t => st.showResolved || !t.c.resolved);
      const anchor = st.anchor && SL.findEl(st.anchor);
      return `<div class="cm-bar"><button class="btn primary" data-act="new" title="New comment on this slide or the selected object">${icon('message-square-plus', 15)} New</button>
          <select data-k="scope" title="Which comments to show"><option value="slide" ${st.scope === 'slide' ? 'selected' : ''}>This slide</option><option value="all" ${st.scope === 'all' ? 'selected' : ''}>All slides</option></select>
          <label class="check" title="Show resolved threads"><input type="checkbox" data-k="resolved" ${st.showResolved ? 'checked' : ''}> Resolved</label></div>
        ${st.composing ? `<div class="cm-compose">${avatar(me())}<div style="flex:1"><div class="cm-anchor">${icon(anchor ? 'square-dashed-mouse-pointer' : 'presentation', 12)} ${anchor ? 'On ' + esc(C().elName(anchor)) : 'On this slide'}</div><textarea id="cm-new" placeholder="Type your comment. Ctrl+Enter to post." title="Comment text"></textarea><div class="cm-row"><button class="btn sm" data-act="cancel">Cancel</button><button class="btn sm primary" data-act="post">Post</button></div></div></div>` : ''}
        ${shown.length ? shown.map(t => thread(t.c, t.si, P().slides[t.si])).join('') : `<p class="pf-note">${list.length ? 'All comments here are resolved.' : st.scope === 'all' ? 'No comments in this presentation.' : 'No comments on this slide. Select an object and click New to comment on it.'}</p>`}
        ${BTS.userName ? '' : '<p class="pf-note">Tip: set your name in File ▸ Options so people know who wrote each comment.</p>'}`;
    },
    wire: (b) => {
      b.addEventListener('change', e => { const k = e.target.dataset.k; if (k === 'scope') { st.scope = e.target.value; E.refreshPane('force'); } if (k === 'resolved') { st.showResolved = e.target.checked; E.refreshPane('force'); } });
      const post = () => { const t = $('#cm-new', b); const v = t && t.value.trim(); if (!v) return; st.composing = false; X.addComment(v, st.anchor); st.anchor = null; E.refreshPane('force'); };
      b.addEventListener('keydown', e => {
        if (e.target.id === 'cm-new' && e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); post(); }
        if (e.target.classList.contains('cm-rin') && e.key === 'Enter') { e.preventDefault(); e.target.closest('.cm-card').querySelector('[data-act="reply"]').click(); }
        if (e.key === 'Escape' && st.composing) { st.composing = false; E.refreshPane('force'); }
      });
      b.addEventListener('click', e => {
        const a = e.target.closest('[data-act]'); const card = e.target.closest('.cm-card');
        if (a) {
          const act = a.dataset.act;
          if (act === 'new') { X.cmd['new-comment'](); return; }
          if (act === 'cancel') { st.composing = false; E.refreshPane('force'); return; }
          if (act === 'post') { post(); return; }
          const t = card && findThread(card.dataset.id); if (!t) return;
          if (act === 'resolve') { t.c.resolved = !t.c.resolved; SL.commit(); E.renderAll(); return; }
          if (act === 'del') { delThread(t.c.id); return; }
          if (act === 'edit') { st.editing = t.c.id; E.refreshPane('force'); return; }
          if (act === 'edit-cancel') { st.editing = null; E.refreshPane('force'); return; }
          if (act === 'edit-save') { const v = card.querySelector('.cm-edit').value.trim(); if (v) { t.c.text = v; t.c.edited = Date.now(); } st.editing = null; SL.commit(); E.renderAll(); return; }
          if (act === 'reply') { const inp = card.querySelector('.cm-rin'); const v = inp.value.trim(); if (!v) return; t.c.replies.push({ id: SL.uid('r'), author: me(), text: v, time: Date.now() }); st.active = t.c.id; SL.commit(); E.renderAll(); setTimeout(() => { const ni = $(`.cm-card[data-id="${t.c.id}"] .cm-rin`); ni && ni.focus(); }, 0); return; }
          if (act === 'rdel') { t.c.replies = t.c.replies.filter(r => r.id !== a.dataset.rid); SL.commit(); E.renderAll(); return; }
          return;
        }
        if (card && !e.target.closest('input,textarea')) {
          st.active = card.dataset.id; const si = +card.dataset.si; const t = findThread(card.dataset.id);
          if (si !== SL.cur) E.goto(si); else E.refreshPane('force');
          if (t && t.c.elId && SL.findEl(t.c.elId)) E.select([t.c.elId]); else E.renderOverlay();
        }
      });
    },
  };
  // comment pins on the canvas: anchored to their object, faded when resolved
  E.commentPins = (z) => (S().comments || []).map((c, i) => {
    normalize(c); const el = c.elId && SL.findEl(c.elId);
    const x = el ? el.x + el.w - 10 : (c.x ?? 10 + i * 30), y = el ? Math.max(0, el.y - 12) : (c.y ?? 10);
    return `<div class="cmt-pin ${c.resolved ? 'resolved' : ''} ${st.active === c.id ? 'active' : ''}" data-cmt-id="${c.id}" style="left:${Math.round(x * z)}px;top:${Math.round(y * z)}px" title="${esc(c.author + ': ' + c.text)}${c.replies.length ? ` (${c.replies.length} repl${c.replies.length > 1 ? 'ies' : 'y'})` : ''}">${icon('message-square', 14)}${c.replies.length ? `<i>${c.replies.length + 1}</i>` : ''}</div>`;
  }).join('');

  /* ---------------- alt text ---------------- */
  const altable = (el) => el && ['image', 'icon', 'chart', 'shape', 'table', 'text'].includes(el.type);
  X.suggestAlt = (el) => {
    if (el.type === 'icon') return (el.name || 'icon').replace(/-/g, ' ') + ' icon';
    if (el.type === 'chart' && el.chart) { const ch = el.chart; const t = { bar: 'Column', line: 'Line', pie: 'Pie', doughnut: 'Doughnut' }[ch.type] || 'Chart'; return `${t} chart${ch.title ? ': ' + ch.title : ''}. ` + (ch.series || []).map(s => `${s.name}: ` + (ch.categories || []).map((c, i) => `${c} ${s.values[i]}`).join(', ')).join('; '); }
    if (el.type === 'image') return (el.name || '').replace(/\.[a-z0-9]+$/i, '').replace(/[_-]+/g, ' ').trim();
    if (el.type === 'table') return `Table with ${el.rows.length} rows and ${(el.rows[0] || []).length} columns`;
    return SL.plainText(el.html || '').slice(0, 120);
  };
  X.altTextDialog = async (el) => {
    el = el || E.primary();
    if (!altable(el)) return BTS.toast('Select a picture, icon, chart or shape first');
    const r = await BTS.dialog({
      title: 'Alt Text', width: 480,
      body: `<p class="pf-note" style="margin-top:0">How would you describe this ${el.type === 'image' ? 'picture' : el.type} and its context to someone who is blind or has low vision?</p>
        <div class="field"><textarea id="alt-t" rows="4" style="width:100%;min-height:90px" placeholder="For example: Bar chart showing plant height rising from 3 cm to 14 cm over two weeks">${esc(el.alt || '')}</textarea></div>
        <div class="cm-row" style="justify-content:space-between"><label class="check"><input type="checkbox" id="alt-dec" ${el.decorative ? 'checked' : ''}> Mark as decorative</label><button class="btn sm" id="alt-sug" title="Fill in a description from the object (icon name, chart data or file name)">${icon('wand-sparkles', 13)} Suggest</button></div>
        <p class="pf-note">Decorative objects (borders, background shapes) add visual interest but aren’t informative; screen readers skip them.</p>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (root) => ({ alt: $('#alt-t', root).value.trim(), dec: $('#alt-dec', root).checked }) }],
      onOpen: (root) => {
        const t = $('#alt-t', root), d = $('#alt-dec', root); const sync = () => { t.disabled = d.checked; }; sync(); d.addEventListener('change', sync);
        $('#alt-sug', root).addEventListener('click', () => { t.value = X.suggestAlt(el); d.checked = false; sync(); t.focus(); });
      },
    });
    if (!r) return;
    if (r.dec) { el.decorative = true; delete el.alt; } else { delete el.decorative; if (r.alt) el.alt = r.alt; else delete el.alt; }
    SL.commit(); E.refreshPane('force'); E.updateRibbon();
  };
  X.cmd['alt-text'] = () => X.altTextDialog();

  /* ---------------- accessibility checker ---------------- */
  function lum(hex) { const c = String(hex).replace('#', ''); const v = [0, 2, 4].map(i => parseInt(c.slice(i, i + 2), 16) / 255).map(x => x <= .03928 ? x / 12.92 : Math.pow((x + .055) / 1.055, 2.4)); return .2126 * v[0] + .7152 * v[1] + .0722 * v[2]; }
  X.contrast = (a, b) => { const l1 = lum(a), l2 = lum(b); return (Math.max(l1, l2) + .05) / (Math.min(l1, l2) + .05); };
  const hex6 = (c) => { if (!c || c === 'none' || c === 'transparent') return null; if (c.startsWith('rgb')) { const m = c.match(/[\d.]+/g).map(Number); return '#' + m.slice(0, 3).map(x => Math.round(x).toString(16).padStart(2, '0')).join(''); } return /^#[0-9a-f]{6}/i.test(c) ? c.slice(0, 7) : null; };
  function slideBgColors(slide, pres) {
    const bg = slide.background || SL.themeBg(pres, slide.layout);
    if (bg.type === 'image') return null;
    if (bg.type === 'gradient') return [hex6(SL.color(bg.color, pres)), hex6(SL.color(bg.color2, pres))].filter(Boolean);
    return [hex6(SL.color(bg.color || 'bg', pres))].filter(Boolean);
  }
  X.checkAccessibility = (pres = P()) => {
    const out = []; const titles = {};
    pres.slides.forEach((s, si) => {
      const title = s.elements.find(e => e.ph === 'title' && e.type === 'text' && !SL.isEmptyText(e));
      if (!title) out.push({ sev: 'error', kind: 'title', si, msg: 'Missing slide title', why: 'Slide titles let people navigate and skim with a screen reader.' });
      else { const t = SL.plainText(title.html).trim().toLowerCase(); (titles[t] = titles[t] || []).push(si); }
      const bgs = slideBgColors(s, pres);
      s.elements.forEach((el, ei) => {
        if (el.hidden) return;
        if ((el.type === 'image' && el.src) || el.type === 'icon' || el.type === 'chart') if (!el.alt && !el.decorative) out.push({ sev: 'error', kind: 'alt', si, elId: el.id, msg: 'Missing alt text: ' + C().elName(el), why: 'Screen readers read alt text to describe pictures, icons and charts.' });
        if (el.type === 'table' && !el.header) out.push({ sev: 'warn', kind: 'table', si, elId: el.id, msg: 'Table has no header row', why: 'A header row tells screen readers what each column contains.' });
        if ((el.type === 'text' || (el.type === 'shape' && !SL.isLine(el))) && !SL.isEmptyText(el)) {
          const d = SL.textDefaults(el, pres); const fg = hex6(d.color); if (!fg) return;
          const fill = typeof el.fill === 'string' && el.fill !== 'none' ? hex6(SL.color(el.fill, pres)) : (el.fill && el.fill.type === 'gradient' ? hex6(SL.color(el.fill.color, pres)) : null);
          const backs = fill && (el.opacity ?? 1) > .6 ? [fill] : bgs; if (!backs || !backs.length) return;
          const ratio = Math.min(...backs.map(b => X.contrast(fg, b)));
          const large = (el.size || 18) >= 18 || (d.bold && (el.size || 18) >= 14);
          if (ratio < (large ? 3 : 4.5)) out.push({ sev: 'warn', kind: 'contrast', si, elId: el.id, ratio, msg: `Hard-to-read text contrast (${ratio.toFixed(1)}:1): ${C().elName(el)}`, why: `Text needs a contrast of at least ${large ? '3' : '4.5'}:1 against its background.` });
        }
        void ei;
      });
      const order = s.elements.filter(e => !e.hidden && !e.decorative);
      if (title && order.length > 1 && order[0] !== title) out.push({ sev: 'tip', kind: 'order', si, elId: title.id, msg: 'Title is not first in the reading order', why: 'Screen readers read objects in reading order; the title should come first.' });
    });
    Object.values(titles).filter(l => l.length > 1).forEach(l => l.forEach(si => out.push({ sev: 'tip', kind: 'duptitle', si, msg: 'Duplicate slide title', why: 'Give each slide a unique title so slides are easy to tell apart.' })));
    return out;
  };
  const SEV = { error: ['Errors', 'circle-x'], warn: ['Warnings', 'triangle-alert'], tip: ['Tips', 'lightbulb'] };
  C().PANES.a11y = {
    title: 'Accessibility',
    html: () => {
      const iss = X.checkAccessibility();
      const grp = (sev) => { const l = iss.filter(i => i.sev === sev); if (!l.length) return ''; return `<div class="a11y-sec ${sev}">${icon(SEV[sev][1], 15)} ${SEV[sev][0]} <span>${l.length}</span></div>` + l.map((i) => `<div class="a11y-row" data-k="${iss.indexOf(i)}" title="${esc(i.why)}"><span class="a11y-s">${i.si + 1}</span><span class="a11y-m">${esc(i.msg)}</span>${i.kind === 'alt' ? '<button class="btn sm" data-fix="alt" title="Add alt text">Add alt text</button>' : i.kind === 'title' ? '<button class="btn sm" data-fix="title" title="Add a title to this slide">Add title</button>' : i.kind === 'contrast' ? '<button class="btn sm" data-fix="contrast" title="Switch the text to black or white, whichever is easier to read">Fix</button>' : i.kind === 'order' ? '<button class="btn sm" data-fix="order" title="Move the title to the start of the reading order">Fix</button>' : i.kind === 'table' ? '<button class="btn sm" data-fix="table" title="Turn on the header row">Fix</button>' : ''}</div>`).join(''); };
      const els = S().elements;
      return (iss.length ? `<p class="pf-note" style="margin-top:0">${iss.filter(i => i.sev !== 'tip').length ? 'Fix these issues so everyone can read your presentation.' : 'No errors. A few tips:'}</p>` + grp('error') + grp('warn') + grp('tip') : `<div class="a11y-ok">${icon('circle-check', 28)}<b>No accessibility issues found</b><span>People with disabilities should have no trouble reading this presentation.</span></div>`) +
        `<div class="pf-sec" style="margin-top:18px">Reading order · slide ${SL.cur + 1}</div><p class="pf-note">Screen readers read objects in this order (first = read first). Moving an object also changes which objects are in front.</p>
        <div class="ro-list">${els.map((el, i) => `<div class="ro-row ${E.sel.includes(el.id) ? 'sel' : ''}" data-id="${el.id}"><span class="ro-n">${i + 1}</span><span class="ro-name">${esc(C().elName(el))}${el.decorative ? ' <small>(decorative)</small>' : ''}</span><button class="rb" data-ro="up" title="Read earlier">${icon('arrow-up', 13)}</button><button class="rb" data-ro="down" title="Read later">${icon('arrow-down', 13)}</button></div>`).join('') || '<p class="pf-note">This slide is empty.</p>'}</div>
        <div class="pf-actions"><button class="btn" data-act="recheck">${icon('refresh-cw', 14)} Check again</button></div>`;
    },
    wire: (b) => b.addEventListener('click', e => {
      const iss = X.checkAccessibility();
      if (e.target.closest('[data-act="recheck"]')) { E.refreshPane('force'); BTS.toast('Accessibility check updated'); return; }
      const ro = e.target.closest('[data-ro]'); const rr = e.target.closest('.ro-row');
      if (ro && rr) { const els = S().elements; const i = els.findIndex(x => x.id === rr.dataset.id), j = i + (ro.dataset.ro === 'up' ? -1 : 1); if (j < 0 || j >= els.length) return; [els[i], els[j]] = [els[j], els[i]]; E.sel = [els[j].id]; SL.commit(); E.renderCanvas(); E.refreshThumb(true); E.refreshPane('force'); return; }
      if (rr) { E.select([rr.dataset.id]); E.refreshPane('force'); return; }
      const row = e.target.closest('.a11y-row'); if (!row) return; const it = iss[+row.dataset.k]; if (!it) return;
      if (it.si !== SL.cur) E.goto(it.si);
      const el = it.elId && SL.findEl(it.elId); if (el) E.select([el.id]);
      const fix = e.target.closest('[data-fix]'); if (!fix) return;
      if (fix.dataset.fix === 'alt') X.altTextDialog(el);
      if (fix.dataset.fix === 'contrast' && el) { const bgs = slideBgColors(S(), P()) || ['#ffffff']; const fill = typeof el.fill === 'string' && el.fill !== 'none' ? hex6(SL.color(el.fill, P())) : null; const b0 = fill || bgs[0]; el.color = X.contrast('#000000', b0) >= X.contrast('#ffffff', b0) ? '#111111' : '#ffffff'; el.html = (el.html || '').replace(/(<[^>]+style="[^"]*?)color:[^;"]*;?/gi, '$1'); SL.commit(); E.renderCanvas(); E.refreshThumb(true); E.refreshPane('force'); }
      if (fix.dataset.fix === 'order' && el) { const els = S().elements; els.splice(els.indexOf(el), 1); els.unshift(el); SL.commit(); E.renderCanvas(); E.refreshThumb(true); E.refreshPane('force'); }
      if (fix.dataset.fix === 'table' && el) { el.header = true; SL.commit(); E.renderCanvas(); E.refreshThumb(true); E.refreshPane('force'); }
      if (fix.dataset.fix === 'title') {
        let t = S().elements.find(x => x.ph === 'title');
        if (!t) { const spec = SL.layoutPlaceholders('titleOnly', SL.W(P()), SL.H(P()), P())[0]; t = SL.newElement('text', spec); S().elements.unshift(t); SL.commit(); E.renderAll(); }
        E.startEdit(t.id);
      }
    }),
  };
  X.cmd['a11y-pane'] = () => E.openPane('a11y');
  // keep the accessibility pane fresh as people edit
  X.ribbon.push(() => { if (E.pane === 'a11y' && !$('#pane-body').contains(document.activeElement)) { clearTimeout(X._a11yT); X._a11yT = setTimeout(() => E.pane === 'a11y' && E.refreshPane(), 400); } });
})();
