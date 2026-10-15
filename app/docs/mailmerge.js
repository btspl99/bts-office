/* BTS Docs 1.2 — Mailings: mail merge from a CSV or Excel list.
 * Merge fields are <span class="mf" data-field="Name">«Name»</span> (exported to .docx as real MERGEFIELDs).
 * {{Name}} placeholders typed in the text are converted to fields automatically. */
(function () {
  const D = window.Docs; const { $, $$ } = D; const esc = BTS.esc;
  const M = D.mm = { cols: [], rows: [], include: [], source: '', preview: false, rec: 0, highlight: false };

  /* ---------- data ---------- */
  D.mmSetData = (cols, rows, source) => {
    M.cols = cols.map((c, i) => String(c == null || c === '' ? 'Column ' + (i + 1) : c).trim());
    M.rows = rows.filter(r => r.some(v => String(v ?? '').trim() !== '')).map(r => M.cols.map((_, i) => r[i] == null ? '' : String(r[i])));
    M.include = M.rows.map(() => true); M.source = source || 'list'; M.rec = 0;
    syncUI();
  };
  function fromSheetRows(aoa, name) {
    if (!aoa.length) throw new Error('That file has no rows.');
    D.mmSetData(aoa[0], aoa.slice(1), name);
  }
  D.mmLoadFile = async (f) => {
    if (!window.XLSX) throw new Error('Spreadsheet reader failed to load.');
    const ext = (f.name.split('.').pop() || '').toLowerCase();
    let wb;
    if (ext === 'csv' || ext === 'txt' || ext === 'tsv') { const text = new TextDecoder('utf-8').decode(f.data).replace(/^﻿/, ''); wb = XLSX.read(text, { type: 'string', raw: true }); }
    else wb = XLSX.read(f.data, { type: 'array', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
    fromSheetRows(aoa, f.name);
  };
  D.mmSelectRecipients = async () => {
    const f = await BTS.pickFile('.csv,.xlsx,.xls,.ods,.tsv,.txt'); if (!f) return;
    try { await D.mmLoadFile(f); convertPlaceholders(true); BTS.toast(`${M.rows.length} recipient${M.rows.length === 1 ? '' : 's'} loaded from ${f.name}`); }
    catch (e) { BTS.toast(e.message || String(e), { error: true }); }
  };
  D.mmTypeList = async () => {
    const sample = M.cols.length ? [M.cols.join(','), ...M.rows.map(r => r.map(v => /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v).join(','))].join('\n') : 'Name,Course,Grade\nAarav Patil,Python Basics,A\nPriya Sharma,Python Basics,A+';
    const v = await BTS.dialog({ title: 'Type a new list', width: 560, body: `<div class="field"><label>One person per line. The first line holds the column names; separate values with commas.</label><textarea id="mm-csv" style="height:220px;font-family:Consolas,Cousine,monospace" spellcheck="false">${esc(sample)}</textarea></div>`, buttons: [{ label: 'Cancel', value: null }, { label: 'Use this list', primary: true, value: (rt) => $('#mm-csv', rt).value }] });
    if (!v) return;
    const wb = XLSX.read(v, { type: 'string', raw: true }); fromSheetRows(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: '' }), 'typed list');
    convertPlaceholders(true); BTS.toast(`${M.rows.length} recipients`);
  };
  const included = () => M.rows.map((r, i) => [r, i]).filter(([, i]) => M.include[i]);
  D.mmRecord = (i) => { const r = M.rows[i]; if (!r) return null; return Object.fromEntries(M.cols.map((c, k) => [c, r[k]])); };
  const valueFor = (rec, field) => { if (!rec) return ''; if (field in rec) return rec[field]; const k = Object.keys(rec).find(x => x.toLowerCase().replace(/\s+/g, '') === String(field).toLowerCase().replace(/\s+/g, '')); return k ? rec[k] : ''; };

  /* ---------- recipient list dialog ---------- */
  D.mmEditList = async () => {
    if (!M.rows.length) { BTS.toast('Choose recipients first (Mailings ▸ Select recipients).'); return; }
    const inc = M.include.slice();
    const body = document.createElement('div'); body.className = 'mm-list-dlg';
    body.innerHTML = `<div class="mm-list-tools"><div class="src-search">${icon('search', 14)}<input id="mm-q" placeholder="Filter recipients" spellcheck="false"></div><button class="btn" data-a="all">Select all</button><button class="btn" data-a="none">Clear all</button><span class="mm-count"></span></div>
      <div class="mm-table-wrap"><table class="mm-table"><thead><tr><th></th>${M.cols.map(c => `<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>${M.rows.map((r, i) => `<tr data-i="${i}"><td><input type="checkbox" ${inc[i] ? 'checked' : ''}></td>${r.map(v => `<td>${esc(v)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>
      <div class="hint">Data source: ${esc(M.source)} · untick the people you don’t want in this merge.</div>`;
    const count = () => { $('.mm-count', body).textContent = `${inc.filter(Boolean).length} of ${M.rows.length} selected`; };
    body.addEventListener('change', e => { const tr = e.target.closest('tr[data-i]'); if (tr) { inc[+tr.dataset.i] = e.target.checked; count(); } });
    body.addEventListener('click', e => { const a = e.target.closest('[data-a]'); if (!a) return; $$('tbody tr', body).forEach(tr => { if (tr.style.display === 'none') return; const i = +tr.dataset.i; inc[i] = a.dataset.a === 'all'; tr.querySelector('input').checked = inc[i]; }); count(); });
    body.addEventListener('input', e => { if (e.target.id !== 'mm-q') return; const t = e.target.value.toLowerCase(); $$('tbody tr', body).forEach(tr => tr.style.display = tr.textContent.toLowerCase().includes(t) ? '' : 'none'); });
    count();
    const v = await BTS.dialog({ title: 'Mail merge recipients', body, width: Math.min(980, 260 + M.cols.length * 140), buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: true }] });
    if (!v) return;
    M.include = inc; if (!M.include[M.rec]) M.rec = Math.max(0, M.include.indexOf(true)); syncUI();
  };

  /* ---------- fields ---------- */
  function fieldEl(name) { const s = document.createElement('span'); s.className = 'mf'; s.contentEditable = 'false'; s.dataset.field = name; s.textContent = '«' + name + '»'; return s; }
  D.mmInsertField = (name) => {
    if (!name) return;
    D.change(() => {
      const r = D.ensureSel(); const live = D.trackBefore ? D.trackBefore() : null; const rr = D.range() || r; if (!live) rr.deleteContents();
      const el = fieldEl(name); rr.insertNode(el); const sp = document.createTextNode(''); el.after(sp); D.setCaret(el.parentNode, [...el.parentNode.childNodes].indexOf(el) + 1);
      if (live) D.trackAfter(live);
    });
    if (M.preview) renderPreview();
  };
  D.mmFieldMenu = (anchor) => {
    const fromDoc = [...new Set($$('.mf', D.doc).map(m => m.dataset.field))];
    const cols = M.cols.length ? M.cols : fromDoc;
    BTS.menu(anchor, [...(cols.length ? [{ head: M.cols.length ? 'Fields in ' + M.source : 'Fields in this document' }, ...cols.map(c => ({ label: c, icon: 'braces', action: () => D.mmInsertField(c) }))] : [{ label: 'Choose recipients first to see their columns', disabled: true }]), { sep: true }, { label: 'Other field name…', icon: 'pencil', action: async () => { const sv = D.saveSel(); const n = await BTS.prompt('Insert merge field', 'Field (column) name', ''); D.restoreSel(sv); if (n && n.trim()) D.mmInsertField(n.trim()); } }, { label: 'Convert {{placeholders}} to fields', icon: 'wand-sparkles', action: () => convertPlaceholders(false) }]);
  };
  /** {{Name}} text → merge fields. */
  function convertPlaceholders(quiet) {
    const hits = []; const tw = document.createTreeWalker(D.doc, NodeFilter.SHOW_TEXT); let n;
    while ((n = tw.nextNode())) if (/\{\{\s*[^{}]+?\s*\}\}/.test(n.data) && !n.parentElement.closest('.mf')) hits.push(n);
    if (!hits.length) { if (!quiet) BTS.toast('No {{placeholders}} found. Type for example {{Name}} where a value should go.'); return 0; }
    let count = 0;
    D.change(() => hits.forEach(t => {
      const frag = document.createDocumentFragment(); let last = 0; const re = /\{\{\s*([^{}]+?)\s*\}\}/g; let m;
      while ((m = re.exec(t.data))) { if (m.index > last) frag.appendChild(document.createTextNode(t.data.slice(last, m.index))); frag.appendChild(fieldEl(m[1])); last = m.index + m[0].length; count++; }
      if (last < t.data.length) frag.appendChild(document.createTextNode(t.data.slice(last)));
      t.replaceWith(frag);
    }));
    if (!quiet || count) BTS.toast(`${count} placeholder${count === 1 ? '' : 's'} converted to merge fields`);
    if (M.preview) renderPreview();
    return count;
  }
  D.mmConvertPlaceholders = convertPlaceholders;
  D.normalizeMergeFields = (root = D.doc) => $$('.mf', root).forEach(m => { const t = '«' + m.dataset.field + '»'; if (m.textContent !== t) m.textContent = t; m.classList.remove('mf-filled', 'mf-missing'); });

  /* ---------- preview ---------- */
  function renderPreview() {
    const rec = D.mmRecord(M.rec);
    $$('.mf', D.doc).forEach(m => { if (M.preview && rec) { const v = valueFor(rec, m.dataset.field); m.textContent = v; m.classList.toggle('mf-filled', true); m.classList.toggle('mf-missing', !v); } else { m.textContent = '«' + m.dataset.field + '»'; m.classList.remove('mf-filled', 'mf-missing'); } });
    D.scheduleLayout(30);
  }
  D.mmPreview = (on) => {
    const now = on ?? !M.preview;
    if (now && !M.rows.length) { BTS.toast('Choose recipients first (Mailings ▸ Select recipients).'); return; }
    if (now) convertPlaceholders(true);
    M.preview = now; if (now && !M.include[M.rec]) M.rec = Math.max(0, M.include.indexOf(true));
    renderPreview(); syncUI();
  };
  D.mmGo = (where) => {
    if (!M.rows.length) { BTS.toast('Choose recipients first.'); return; }
    const idx = included().map(([, i]) => i); if (!idx.length) return;
    let p = idx.indexOf(M.rec); if (p < 0) p = 0;
    if (where === 'first') p = 0; else if (where === 'last') p = idx.length - 1; else if (where === 'next') p = Math.min(idx.length - 1, p + 1); else if (where === 'prev') p = Math.max(0, p - 1); else if (typeof where === 'number') p = Math.max(0, Math.min(idx.length - 1, where - 1));
    M.rec = idx[p]; if (!M.preview) D.mmPreview(true); else { renderPreview(); syncUI(); }
  };
  D.mmHighlight = (on) => { M.highlight = on ?? !M.highlight; document.body.classList.toggle('mm-hl', M.highlight); syncUI(); };
  function syncUI() {
    const idx = included().map(([, i]) => i); const p = idx.indexOf(M.rec);
    const inp = $('#mm-rec'); if (inp && document.activeElement !== inp) inp.value = M.rows.length ? String(p + 1) : '';
    const tot = $('#mm-total'); if (tot) tot.textContent = M.rows.length ? 'of ' + idx.length : '';
    $('#btn-mm-preview')?.classList.toggle('on', M.preview); $('#btn-mm-hl')?.classList.toggle('on', M.highlight);
    document.body.classList.toggle('mm-preview', M.preview);
    BTS.setStatus('mm', M.rows.length ? `<span title="Mail merge data source" style="display:inline-flex;align-items:center;gap:5px;white-space:nowrap">${icon('mail', 14)} ${esc(M.source)} · ${idx.length} recipient${idx.length === 1 ? '' : 's'}${M.preview ? ' · previewing ' + (p + 1) : ''}</span>` : '');
  }
  D.mmSync = syncUI;

  /* ---------- finish ---------- */
  /** HTML of the merged document: one copy per included record, separated by page breaks. */
  D.mmMergedHTML = () => {
    const recs = included(); if (!recs.length) throw new Error('No recipients are selected.');
    const src = D.doc.cloneNode(true); D.normalizeMergeFields(src);
    $$('.cmt', src).forEach(D.unwrap); $$('.cell-sel', src).forEach(e => e.classList.remove('cell-sel')); $$('.pv-on', src).forEach(e => e.className = e.className.replace(/\bpv-\S+/g, '').trim());
    const tpl = src.innerHTML;
    return recs.map(([, i], k) => {
      const rec = D.mmRecord(i); const box = document.createElement('div'); box.innerHTML = tpl;
      $$('.mf', box).forEach(m => m.replaceWith(document.createTextNode(valueFor(rec, m.dataset.field))));
      const tw = document.createTreeWalker(box, NodeFilter.SHOW_TEXT); const ts = []; let n; while ((n = tw.nextNode())) if (n.data.includes('{{')) ts.push(n);
      ts.forEach(t => { t.data = t.data.replace(/\{\{\s*([^{}]+?)\s*\}\}/g, (m0, f) => valueFor(rec, f)); });
      if (k > 0) $$('[id]', box).forEach(e => e.removeAttribute('id'));
      $$('hr.page-break', box).forEach(h => h.removeAttribute('style'));
      return box.innerHTML;
    }).join('<hr class="page-break">');
  };
  /** Build a .docx from arbitrary body HTML using the live layout engine, without touching undo or the dirty flag. */
  D.docxFromHTML = async (html) => {
    const keep = { html: D.doc.innerHTML, dirty: BTS.dirty, sel: D.saveSel() };
    D.loading = true;
    try { D.doc.innerHTML = html; D.renderFootnotes(); D.layout(); return await D.exportDocx(); }
    finally { D.doc.innerHTML = keep.html; D.loading = false; D.renderFootnotes(); D.layout(); D.restoreSel(keep.sel); if (!keep.dirty) BTS.markClean(); }
  };
  D.mmFinish = async (mode) => {
    if (!M.rows.length) { BTS.toast('Choose recipients first (Mailings ▸ Select recipients).'); return; }
    convertPlaceholders(true);
    const wasPreview = M.preview; if (wasPreview) { M.preview = false; renderPreview(); }
    let html; try { html = D.mmMergedHTML(); } catch (e) { BTS.toast(e.message, { error: true }); return; }
    const n = included().length;
    if (mode === 'here') {
      const ok = await BTS.dialog({ title: 'Merge into this window', body: `<p>This replaces the template in this window with ${n} merged cop${n === 1 ? 'y' : 'ies'}. Save the template first if you want to keep it.</p>`, buttons: [{ label: 'Cancel', value: false }, { label: 'Merge', primary: true, value: true }] });
      if (!ok) { if (wasPreview) D.mmPreview(true); return; }
      D.change(() => { D.doc.innerHTML = html; }); D.renderFootnotes(); BTS.setTitle((BTS.title || 'Document') + ' (merged)'); BTS.filePath = null; M.preview = false; syncUI();
      BTS.toast(`Merged ${n} document${n === 1 ? '' : 's'}`); return;
    }
    const bytes = await D.docxFromHTML(html);
    if (wasPreview) D.mmPreview(true);
    if (!BTS.desktop) { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([bytes])); a.download = (BTS.title || 'Document') + ' - merged.docx'; a.click(); return; }
    const path = await bts.saveDialog((BTS.title || 'Document') + ' - merged.docx', [{ name: 'Word document', extensions: ['docx'] }]); if (!path) return;
    await bts.writeFile(path, bytes); bts.openApp('docs', path);
    BTS.toast(`Merged ${n} document${n === 1 ? '' : 's'} — opening in a new window`);
  };
  D.mmFinishMenu = (anchor) => BTS.menu(anchor, [{ label: 'Edit individual documents…', icon: 'files', action: () => D.mmFinish('window') }, { label: 'Merge into this window (replace template)', icon: 'file-stack', action: () => D.mmFinish('here') }]);
  D.mmRecipientsMenu = (anchor) => BTS.menu(anchor, [{ label: 'Use an existing list (CSV or Excel)…', icon: 'file-spreadsheet', action: D.mmSelectRecipients }, { label: 'Type a new list…', icon: 'keyboard', action: D.mmTypeList }, { sep: true }, { label: 'Remove the recipient list', icon: 'trash-2', disabled: !M.rows.length, action: () => { if (M.preview) D.mmPreview(false); D.mmSetData([], [], ''); M.cols = []; syncUI(); } }]);

  D.initMailMerge = () => {
    const inp = $('#mm-rec'); if (inp) inp.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); D.mmGo(+inp.value || 1); D.restoreFocus(); } });
    syncUI();
  };
})();
