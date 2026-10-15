/* BTS Sheets — threaded comments (separate from notes).
 * s.threads: Map(cellKey -> { id, items:[{ id, author, text, time }], resolved }) — shown in the Comments pane, marked purple on the grid. */
(function () {
  const Threads = window.Threads = {};
  const { key, unkey } = SU;
  let wb, draftFor = null;
  Threads.init = (w) => { wb = w; };
  const S = () => wb.active;
  const esc = SU.esc;
  const me = () => BTS.userName || 'Me';
  const initials = (n) => (n || '?').split(/\s+/).filter(Boolean).slice(0, 2).map(x => x[0].toUpperCase()).join('') || '?';
  const hue = (n) => { let h = 0; for (const ch of String(n)) h = (h * 31 + ch.charCodeAt(0)) % 360; return h; };
  const when = (iso) => { const d = new Date(iso); if (isNaN(d)) return ''; const diff = (Date.now() - d) / 1000; if (diff < 60) return 'just now'; if (diff < 3600) return Math.floor(diff / 60) + ' min ago'; if (diff < 86400) return Math.floor(diff / 3600) + ' h ago'; return d.toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: d.getFullYear() === new Date().getFullYear() ? undefined : 'numeric' }) + ' ' + d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' }); };
  const uid = () => 'tc' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

  function setThreads(s, m, label) { wb.run(label, () => wb.setProp(s, 'threads', m)); Grid.render(); Threads.render(); }
  Threads.add = (s, r, c, text, author = me()) => {
    const m = new Map(s.threads); const k = key(r, c); const cur = m.get(k);
    const item = { id: uid(), author, text, time: new Date().toISOString() };
    m.set(k, cur ? { ...cur, items: [...cur.items, item], resolved: false } : { id: uid(), items: [item], resolved: false });
    setThreads(s, m, cur ? 'Reply' : 'New comment');
  };
  Threads.resolve = (s, k, v) => { const m = new Map(s.threads); const t = m.get(k); if (!t) return; m.set(k, { ...t, resolved: v }); setThreads(s, m, v ? 'Resolve thread' : 'Reopen thread'); };
  Threads.removeThread = (s, k) => { const m = new Map(s.threads); m.delete(k); setThreads(s, m, 'Delete comment'); };
  Threads.removeItem = (s, k, id) => { const m = new Map(s.threads); const t = m.get(k); if (!t) return; const items = t.items.filter(x => x.id !== id); if (items.length) m.set(k, { ...t, items }); else m.delete(k); setThreads(s, m, 'Delete reply'); };
  Threads.edit = (s, k, id, text) => { const m = new Map(s.threads); const t = m.get(k); if (!t) return; m.set(k, { ...t, items: t.items.map(x => x.id === id ? { ...x, text, edited: new Date().toISOString() } : x) }); setThreads(s, m, 'Edit comment'); };

  /* ---------- pane ---------- */
  function paneEl() {
    let el = document.getElementById('cm-pane'); if (el) return el;
    el = document.createElement('aside'); el.className = 'pane side-pane hidden'; el.id = 'cm-pane';
    el.innerHTML = `<header><span>Comments</span><button class="rb" data-cmd="comment-new" data-icon="message-square-plus" title="New comment (Ctrl+Alt+M)" style="margin-left:auto">${icon('message-square-plus', 16)}</button><button class="rb" data-cmd="comments-pane-close" title="Close the Comments pane" style="margin-left:2px">${icon('x', 16)}</button></header>
      <div class="pane-body cm-body"><label class="check cm-filter"><input type="checkbox" id="cm-show-resolved" checked> Show resolved threads</label><div class="cm-list"></div></div>`;
    document.getElementById('workspace').appendChild(el); wire(el); return el;
  }
  Threads.isOpen = () => { const el = document.getElementById('cm-pane'); return !!el && !el.classList.contains('hidden'); };
  Threads.open = (newDraft) => {
    const el = paneEl(); const was = el.classList.contains('hidden');
    if (window.Pivot) { Pivot.paneClosed = true; Pivot.refreshPane(); }
    el.classList.remove('hidden'); if (was) App.layoutPanes();
    if (newDraft) { const s = S(); draftFor = { s, k: key(s.active.r, s.active.c) }; }
    Threads.render();
    if (newDraft) setTimeout(() => { const t = el.querySelector('.cm-card.draft textarea, .cm-card.on textarea'); t && t.focus(); }, 30);
    App.refreshRibbon();
  };
  Threads.close = (silent) => { const el = document.getElementById('cm-pane'); if (!el || el.classList.contains('hidden')) return; el.classList.add('hidden'); draftFor = null; if (!silent) App.layoutPanes(); App.refreshRibbon(); };
  Threads.toggle = () => Threads.isOpen() ? Threads.close() : Threads.open();
  Threads.render = () => {
    const el = document.getElementById('cm-pane'); if (!el || el.classList.contains('hidden')) return;
    const s = S(); const act = key(s.active.r, s.active.c); const showRes = el.querySelector('#cm-show-resolved').checked;
    const list = [...s.threads.entries()].map(([k, t]) => ({ k, t, ...unkey(k) })).sort((a, b) => a.r - b.r || a.c - b.c).filter(x => showRes || !x.t.resolved || x.k === act);
    let h = '';
    if (draftFor && draftFor.s === s && !s.threads.has(draftFor.k)) { const p = unkey(draftFor.k); h += `<div class="cm-card draft on" data-k="${draftFor.k}"><div class="cm-addr">${SU.addr(p.r, p.c)}</div><div class="cm-item"><span class="cm-av" style="background:hsl(${hue(me())} 55% 42%)">${esc(initials(me()))}</span><b>${esc(me())}</b></div><textarea placeholder="Start a conversation… (Ctrl+Enter to post)"></textarea><div class="cm-actions"><button class="btn" data-a="cancel">Cancel</button><button class="btn primary" data-a="post">Post</button></div></div>`; }
    for (const { k, t, r, c } of list) {
      h += `<div class="cm-card ${k === act ? 'on' : ''} ${t.resolved ? 'resolved' : ''}" data-k="${k}"><div class="cm-addr">${SU.addr(r, c)}${t.resolved ? '<span class="cm-tag">Resolved</span>' : ''}<span class="cm-tools"><button class="rb" data-a="${t.resolved ? 'reopen' : 'resolve'}" title="${t.resolved ? 'Reopen thread' : 'Resolve thread'}">${icon(t.resolved ? 'rotate-ccw' : 'check-check', 15)}</button><button class="rb" data-a="delthread" title="Delete thread">${icon('trash-2', 15)}</button></span></div>`;
      h += t.items.map((it, i) => `<div class="cm-item ${i ? 'reply' : ''}" data-id="${it.id}"><span class="cm-av" style="background:hsl(${hue(it.author)} 55% 42%)">${esc(initials(it.author))}</span><div class="cm-main"><div class="cm-meta"><b>${esc(it.author)}</b><small title="${esc(new Date(it.time).toLocaleString())}">${esc(when(it.time))}${it.edited ? ' · edited' : ''}</small>${it.author === me() ? `<button class="rb cm-mini" data-a="edit" title="Edit">${icon('pencil', 13)}</button>` : ''}${i ? `<button class="rb cm-mini" data-a="delitem" title="Delete reply">${icon('x', 13)}</button>` : ''}</div><div class="cm-text">${esc(it.text)}</div></div></div>`).join('');
      if (!t.resolved) h += `<textarea class="cm-reply" placeholder="Reply…"></textarea><div class="cm-actions hidden"><button class="btn" data-a="cancel">Cancel</button><button class="btn primary" data-a="reply">Reply</button></div>`;
      h += '</div>';
    }
    el.querySelector('.cm-list').innerHTML = h || `<div class="empty-note cm-empty">${icon('messages-square', 28)}<p>No comments on this sheet yet.</p><p>Select a cell and choose <b>New Comment</b> to start a conversation${BTS.userName ? '' : '. Set your name in File ▸ Options so people know who wrote it'}.</p></div>`;
    const on = el.querySelector('.cm-card.on'); if (on) on.scrollIntoView({ block: 'nearest' });
  };
  function wire(el) {
    el.querySelector('#cm-show-resolved').addEventListener('change', () => Threads.render());
    el.addEventListener('focusin', (e) => { if (e.target.classList.contains('cm-reply')) { const a = e.target.nextElementSibling; a && a.classList.remove('hidden'); } });
    el.addEventListener('keydown', (e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey) && e.target.tagName === 'TEXTAREA') { e.preventDefault(); const card = e.target.closest('.cm-card'); const b = card.querySelector('[data-a="post"],[data-a="reply"]'); b && b.click(); } if (e.key === 'Escape' && e.target.tagName === 'TEXTAREA') { e.stopPropagation(); e.target.blur(); } });
    el.addEventListener('click', async (e) => {
      const card = e.target.closest('.cm-card'); if (!card) return; const s = S(); const k = +card.dataset.k; const p = unkey(k);
      const a = e.target.closest('[data-a]');
      if (!a) { if (!e.target.closest('textarea')) { if (s.active.r !== p.r || s.active.c !== p.c) { Grid.select(p.r, p.c); } } return; }
      const act = a.dataset.a; const ta = card.querySelector('textarea');
      if (act === 'cancel') { if (card.classList.contains('draft')) { draftFor = null; Threads.render(); Grid.focus(); } else { ta.value = ''; a.parentElement.classList.add('hidden'); } return; }
      if (act === 'post' || act === 'reply') { const t = ta.value.trim(); if (!t) { ta.focus(); return; } if (act === 'post') draftFor = null; Threads.add(s, p.r, p.c, t); return; }
      if (act === 'resolve' || act === 'reopen') return Threads.resolve(s, k, act === 'resolve');
      if (act === 'delthread') return Threads.removeThread(s, k);
      const item = e.target.closest('.cm-item'); const id = item && item.dataset.id;
      if (act === 'delitem') return Threads.removeItem(s, k, id);
      if (act === 'edit') { const t = s.threads.get(k).items.find(x => x.id === id); const v = await BTS.dialog({ title: 'Edit comment', body: `<textarea id="ec" style="width:360px;height:100px">${esc(t.text)}</textarea>`, buttons: [{ label: 'Cancel', value: null }, { label: 'Save', primary: true, value: (r) => r.querySelector('#ec').value }] }); if (v && v.trim()) Threads.edit(s, k, id, v.trim()); }
    });
  }
  Threads.newComment = () => { if (A.protectedCheck(false, true, 'objects')) return; Threads.open(true); };
  Threads.go = (dir) => {
    const s = S(); const keys = [...s.threads.keys()].map(unkey).sort((a, b) => a.r - b.r || a.c - b.c); if (!keys.length) { BTS.toast('This sheet has no comments.', { error: true }); return; }
    const a = s.active; let i = dir > 0 ? keys.findIndex(k => k.r > a.r || (k.r === a.r && k.c > a.c)) : keys.map((k, j) => [k, j]).filter(([k]) => k.r < a.r || (k.r === a.r && k.c < a.c)).map(x => x[1]).pop();
    if (i === undefined || i < 0) i = dir > 0 ? 0 : keys.length - 1;
    Grid.select(keys[i].r, keys[i].c); Threads.open();
  };
  Threads.deleteAtActive = () => { const s = S(); const k = key(s.active.r, s.active.c); if (!s.threads.has(k)) { BTS.toast('The selected cell has no comment.', { error: true }); return; } Threads.removeThread(s, k); };
  /** hover preview text */
  Threads.preview = (s, r, c) => { const t = s.threads.get(key(r, c)); if (!t) return null; const f = t.items[0]; return { title: f.author + (t.resolved ? ' · resolved' : ''), text: f.text + (t.items.length > 1 ? `\n— ${t.items.length - 1} repl${t.items.length > 2 ? 'ies' : 'y'}` : '') }; };
  /** legacy note text written to xlsx for apps that don't read our metadata */
  Threads.asNote = (t) => t.items.map((it, i) => (i ? 'Reply ' : '') + it.author + ':\n' + it.text).join('\n\n');
})();
