/* BTS Sheets — protection: protect sheet (optional password, allowed actions), protect workbook structure, allow-edit ranges.
 * Passwords are never stored: we keep the Excel-compatible SHA-512 hash (salt + spin count), written to xlsx as-is. */
(function () {
  const Prot = window.Prot = {};
  let wb;
  Prot.init = (w) => { wb = w; };
  const S = () => wb.active;
  const esc = SU.esc;
  Prot.SPIN = 10000; // Excel reads the spin count from the file; 10k keeps protecting quick
  Prot.OPTIONS = [['selectLockedCells', 'Select locked cells', true], ['selectUnlockedCells', 'Select unlocked cells', true], ['formatCells', 'Format cells'], ['formatColumns', 'Format columns'], ['formatRows', 'Format rows'], ['insertColumns', 'Insert columns'], ['insertRows', 'Insert rows'], ['insertHyperlinks', 'Insert hyperlinks'], ['deleteColumns', 'Delete columns'], ['deleteRows', 'Delete rows'], ['sort', 'Sort'], ['autoFilter', 'Use AutoFilter'], ['pivotTables', 'Use PivotTable reports'], ['objects', 'Edit objects']];

  /* ---------- Excel's password hash: H0 = SHA512(salt + UTF16LE(pw)); Hn = SHA512(Hn-1 + LE32(n)) ---------- */
  const b64 = (u8) => { let s = ''; for (const x of u8) s += String.fromCharCode(x); return btoa(s); };
  const unb64 = (s) => Uint8Array.from(atob(s), c => c.charCodeAt(0));
  Prot.hash = async (pw, saltB64, spin = Prot.SPIN) => {
    const salt = saltB64 ? unb64(saltB64) : crypto.getRandomValues(new Uint8Array(16));
    const pwb = new Uint8Array(pw.length * 2); for (let i = 0; i < pw.length; i++) { const c = pw.charCodeAt(i); pwb[i * 2] = c & 255; pwb[i * 2 + 1] = c >> 8; }
    const first = new Uint8Array(salt.length + pwb.length); first.set(salt); first.set(pwb, salt.length);
    let h = new Uint8Array(await crypto.subtle.digest('SHA-512', first));
    const buf = new Uint8Array(68);
    for (let i = 0; i < spin; i++) { buf.set(h); buf[64] = i & 255; buf[65] = (i >>> 8) & 255; buf[66] = (i >>> 16) & 255; buf[67] = (i >>> 24) & 255; h = new Uint8Array(await crypto.subtle.digest('SHA-512', buf)); }
    return { algorithmName: 'SHA-512', hashValue: b64(h), saltValue: b64(salt), spinCount: spin };
  };
  Prot.verify = async (pw, h) => { if (h && h.password) return legacy(pw) === String(h.password).toUpperCase(); if (!h || !h.hashValue) return true; const x = await Prot.hash(pw, h.saltValue, h.spinCount || Prot.SPIN); return x.hashValue === h.hashValue; };
  /** old Excel 16-bit "password" attribute */
  function legacy(pw) { let h = 0; for (let i = pw.length - 1; i >= 0; i--) { h = ((h >> 14) & 1) | ((h << 1) & 0x7fff); h ^= pw.charCodeAt(i); } h = ((h >> 14) & 1) | ((h << 1) & 0x7fff); h ^= pw.length; h ^= 0xCE4B; return h.toString(16).toUpperCase().padStart(4, '0'); }
  Prot.legacy = legacy;
  async function busy(fn) { const t = setTimeout(() => BTS.toast('Securing…', { ms: 1500 }), 250); try { return await fn(); } finally { clearTimeout(t); } }

  /* ---------- protect sheet ---------- */
  Prot.toggleSheet = async () => {
    const s = S();
    if (s.protected) {
      const h = s.protection && s.protection.hash;
      if (h && h.hashValue || h && h.password) { const pw = await askPassword('Unprotect sheet'); if (pw === null) return; if (!(await busy(() => Prot.verify(pw, h)))) { BTS.toast('The password you supplied is not correct.', { error: true }); return; } }
      wb.run('Unprotect sheet', () => wb.setProp(s, 'protected', false));
      BTS.toast('Sheet unprotected'); after(); return;
    }
    const opts = (s.protection && s.protection.opts) || {};
    const res = await BTS.dialog({ title: 'Protect sheet', width: 400, body: `<label class="check"><input type="checkbox" checked disabled> Protect worksheet and contents of locked cells</label>
      <div class="field"><label>Password to unprotect sheet (optional)</label><input id="ps-pw" type="password" autocomplete="new-password"></div>
      <div class="field"><label>Reenter password</label><input id="ps-pw2" type="password" autocomplete="new-password"></div>
      <div class="dlg-note" style="margin:6px 0 4px">Allow all users of this worksheet to:</div>
      <div class="ps-opts">${Prot.OPTIONS.map(([k, l, def]) => `<label class="check"><input type="checkbox" data-k="${k}" ${(opts[k] ?? !!def) ? 'checked' : ''}> ${l}</label>`).join('')}</div>
      <p class="dlg-note">Tip: unlock input cells first (Format Cells ▸ Locked) so people can still fill them in.</p>`,
      buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ pw: r.querySelector('#ps-pw').value, pw2: r.querySelector('#ps-pw2').value, opts: Object.fromEntries([...r.querySelectorAll('[data-k]')].map(x => [x.dataset.k, x.checked])) }) }] });
    if (!res) return;
    if (res.pw !== res.pw2) { BTS.toast('The passwords don’t match.', { error: true }); return; }
    const hash = res.pw ? await busy(() => Prot.hash(res.pw)) : null;
    wb.run('Protect sheet', () => { wb.setProp(s, 'protection', { opts: res.opts, hash }); wb.setProp(s, 'protected', true); });
    BTS.toast(res.pw ? 'Sheet protected with a password' : 'Sheet protected. Only unlocked cells can be edited.'); after();
  };
  function askPassword(title) { return BTS.dialog({ title, width: 340, body: `<div class="field"><label>Password</label><input id="pp" type="password" autocomplete="current-password"></div>`, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => r.querySelector('#pp').value }] }); }
  function after() { App.renderTabs(); App.refreshRibbon(); App.refreshStatus(); Grid.render(); }
  Prot.allowed = (s, kind) => !!(s.protection && s.protection.opts && s.protection.opts[kind]);

  /* ---------- protect workbook structure ---------- */
  Prot.toggleWorkbook = async () => {
    if (wb.protectStructure) {
      const h = wb.protectStructure.hash;
      if (h && (h.hashValue || h.password)) { const pw = await askPassword('Unprotect workbook'); if (pw === null) return; if (!(await busy(() => Prot.verify(pw, h)))) { BTS.toast('The password you supplied is not correct.', { error: true }); return; } }
      wb.run('Unprotect workbook', () => wb.setProp(wb, 'protectStructure', null)); BTS.toast('Workbook structure unprotected'); after(); return;
    }
    const res = await BTS.dialog({ title: 'Protect structure and windows', width: 380, body: `<label class="check"><input type="checkbox" checked disabled> Structure (sheets can’t be added, deleted, renamed, moved or recoloured)</label><div class="field"><label>Password (optional)</label><input id="pw-1" type="password" autocomplete="new-password"></div><div class="field"><label>Reenter password</label><input id="pw-2" type="password" autocomplete="new-password"></div>`, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: (r) => ({ pw: r.querySelector('#pw-1').value, pw2: r.querySelector('#pw-2').value }) }] });
    if (!res) return; if (res.pw !== res.pw2) { BTS.toast('The passwords don’t match.', { error: true }); return; }
    const hash = res.pw ? await busy(() => Prot.hash(res.pw)) : null;
    wb.run('Protect workbook', () => wb.setProp(wb, 'protectStructure', { hash })); BTS.toast('Workbook structure protected'); after();
  };
  Prot.structureCheck = () => { if (!wb.protectStructure) return false; BTS.toast('The workbook structure is protected. Unprotect it on the Review tab to change sheets.', { error: true, ms: 3500 }); return true; };

  /* ---------- allow edit ranges ---------- */
  Prot.allowRangesDialog = async () => {
    const s = S(); let list = s.allowRanges.map(x => ({ ...x }));
    const body = document.createElement('div');
    const render = () => { body.innerHTML = `<p class="dlg-note" style="margin:0 0 6px">Ranges unlocked by a password when the sheet is protected — here they are simply editable by everyone.</p><div class="nm-wrap" style="min-width:420px"><table class="nm-table"><thead><tr><th>Title</th><th>Refers to cells</th><th></th></tr></thead><tbody>${list.map((a, i) => `<tr><td>${esc(a.title)}</td><td>${SU.absRangeStr(a.range)}</td><td style="width:30px"><button class="rb" data-del="${i}" title="Delete range">${icon('trash-2', 14)}</button></td></tr>`).join('') || '<tr><td colspan="3" class="empty-hint">No ranges yet.</td></tr>'}</tbody></table></div><div class="field-row" style="margin-top:8px;gap:6px"><input id="ar-title" placeholder="Title" value="Range${list.length + 1}" style="flex:1;height:30px;border:1px solid var(--line-2);border-radius:7px;padding:0 8px"><input id="ar-ref" value="${SU.absRangeStr(A.effData(s.sel))}" style="flex:1;height:30px;border:1px solid var(--line-2);border-radius:7px;padding:0 8px"><button class="btn" data-a="add">${icon('plus', 15)} New</button></div>`; };
    body.addEventListener('click', (e) => { const d = e.target.closest('[data-del]'); if (d) { list.splice(+d.dataset.del, 1); render(); return; } if (e.target.closest('[data-a="add"]')) { const g = SU.parseRange(body.querySelector('#ar-ref').value); if (!g) { BTS.toast('That reference isn’t valid.', { error: true }); return; } list.push({ title: body.querySelector('#ar-title').value.trim() || 'Range' + (list.length + 1), range: g }); render(); } });
    render();
    const ok = await BTS.dialog({ title: 'Allow users to edit ranges', width: 520, body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, value: true }] });
    if (ok) wb.run('Allow edit ranges', () => wb.setProp(s, 'allowRanges', list));
  };
  /** may this cell be edited on a protected sheet? */
  Prot.cellEditable = (s, r, c) => s.style(r, c).locked === false || s.allowRanges.some(a => SU.inRange(a.range, r, c));
  Prot.rangeEditable = (s, g) => { const e = A.eff(g, s); const n = (e.r1 - e.r0 + 1) * (e.c1 - e.c0 + 1); if (n > 200000) return false; for (let r = e.r0; r <= e.r1; r++) for (let c = e.c0; c <= e.c1; c++) if (!Prot.cellEditable(s, r, c)) return false; return true; };
})();
