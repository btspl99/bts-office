/* BTS Slides 1.3 — autosave & crash recovery (like BTS Docs, but in IndexedDB because slides hold big pictures).
 * Each window writes its own slot (key = window id) in the IndexedDB store "bts-slides-recovery/slots", debounced,
 * and beats a heartbeat in localStorage ("bts-slides-hb:<id>") every 5 s. On startup we only offer slots whose window
 * stopped beating (crashed or closed without saving) — never another open window's work. A successful save
 * (BTS.markClean) or a clean close clears the slot.
 */
(function () {
  const $ = (s, r = document) => r.querySelector(s);
  const esc = (s) => BTS.esc(s);
  const SID = Math.random().toString(36).slice(2, 10);
  const HB = 'bts-slides-hb:';
  const ls = { get: (k) => { try { return localStorage.getItem(k); } catch { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch { } }, del: (k) => { try { localStorage.removeItem(k); } catch { } } };
  let dbp = null;
  const db = () => dbp || (dbp = new Promise((res, rej) => { const r = indexedDB.open('bts-slides-recovery', 1); r.onupgradeneeded = () => r.result.createObjectStore('slots'); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }));
  const tx = async (mode, fn) => { const d = await db(); return new Promise((res, rej) => { const t = d.transaction('slots', mode); const req = fn(t.objectStore('slots')); t.oncomplete = () => res(req && req.result); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error); }); };
  const R = window.SLRecovery = { sid: SID };
  R.put = (k, v) => tx('readwrite', s => s.put(v, k));
  R.del = (k) => tx('readwrite', s => s.delete(k)).catch(() => { });
  R.all = async () => { const keys = await tx('readonly', s => s.getAllKeys()); const vals = await tx('readonly', s => s.getAll()); return keys.map((k, i) => [k, vals[i]]); };

  R.snapshot = () => {
    const used = {}; const json = JSON.stringify(SL.pres);
    json.replace(/"(a[\w]+)"/g, (m, id) => { if (SL.assets[id]) used[id] = SL.assets[id]; });
    return { t: Date.now(), title: BTS.title, path: BTS.filePath || null, pres: json, assets: used, slides: SL.pres.slides.length };
  };
  let timer = 0, writing = false;
  // File > Options > Save: "Save recovery info every N seconds" (BTS.prefs.autosaveSeconds); 0 = soon after each change
  R.schedule = () => {
    const every = +((BTS.prefs && BTS.prefs.autosaveSeconds) || 0) * 1000;
    if (every && timer) return; // already planned; the snapshot is taken when it runs
    clearTimeout(timer);
    timer = setTimeout(async () => {
      timer = 0;
      if (!BTS.dirty || writing) return; writing = true;
      try { await R.put(SID, R.snapshot()); } catch (e) { console.warn('autosave', e); } finally { writing = false; }
    }, every || 2000);
  };
  R.clear = () => { clearTimeout(timer); timer = 0; return R.del(SID); };
  const beat = () => ls.set(HB + SID, String(Date.now()));
  beat(); setInterval(beat, 5000);

  R.check = async () => {
    let list = []; try { list = await R.all(); } catch { return; }
    const now = Date.now(); let best = null;
    for (const [k, v] of list) {
      if (k === SID) continue;
      if (!v || !v.pres || now - v.t > 14 * 864e5) { R.del(k); ls.del(HB + k); continue; }
      if (now - (+ls.get(HB + k) || 0) < 15000) continue; // that window is still open
      if (!best || v.t > best.v.t) best = { k, v };
    }
    if (!best) return null;
    const bar = $('#recover-bar'); if (!bar) return best;
    const when = new Date(best.v.t).toLocaleString('en-IN', { day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit' });
    bar.innerHTML = `${icon('history', 16)}<span>Recover unsaved presentation “${esc(best.v.title || 'Presentation')}” (${best.v.slides || '?'} slide${best.v.slides === 1 ? '' : 's'}) from ${esc(when)}?</span><span style="flex:1"></span><button class="btn primary" id="rec-yes">Recover</button><button class="btn" id="rec-no">Dismiss</button>`;
    bar.classList.remove('hidden'); requestAnimationFrame(E.layout);
    const close = () => { bar.classList.add('hidden'); bar.innerHTML = ''; requestAnimationFrame(E.layout); };
    $('#rec-yes', bar).onclick = () => {
      try {
        Object.assign(SL.assets, best.v.assets || {}); SLApp.setPresentation(JSON.parse(best.v.pres));
        BTS.title = best.v.title || 'Presentation'; BTS.filePath = null; BTS.markDirty();
        if (BTS.setTitle) BTS.setTitle(BTS.title);
        R.del(best.k); ls.del(HB + best.k); BTS.toast('Presentation recovered — save it to keep it');
      } catch (e) { BTS.toast('Couldn’t recover: ' + e.message, { error: true }); }
      close();
    };
    $('#rec-no', bar).onclick = () => { R.del(best.k); ls.del(HB + best.k); close(); };
    return best;
  };

  function hook() {
    const md = BTS.markDirty; BTS.markDirty = () => { md(); R.schedule(); };
    const mc = BTS.markClean; BTS.markClean = () => { mc(); R.clear(); };
    addEventListener('beforeunload', () => { ls.del(HB + SID); if (!BTS.dirty) R.clear(); });
    setTimeout(() => { if (!BTS.loadedExt && !BTS.filePath) R.check(); }, 1500);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => setTimeout(hook, 0)); else setTimeout(hook, 0);
})();
