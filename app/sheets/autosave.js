/* BTS Sheets — autosave & crash recovery.
 * Each window keeps its own slot: the workbook snapshot lives in IndexedDB (workbooks with charts/pictures can be
 * far bigger than localStorage allows), and a tiny index entry with a heartbeat lives in localStorage
 * (bts-sheets-autosave:<sid>). On startup we only offer slots whose window stopped beating (crashed or closed
 * without saving), never the live work of another open window. A save, a deliberate new file or a clean close
 * removes the slot. The snapshot is serialised from the workbook model itself (cells with formulas, styles,
 * merges, charts, pictures, comments, protection, sparklines, PivotTables, names, theme, sheet tabs). */
(function () {
  const AS = window.SAuto = {};
  const LK = 'bts-sheets-autosave:'; const SID = Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
  const STALE = 15000, MAX_AGE = 14 * 864e5, DEBOUNCE = 1500, MAX_WAIT = 10000;
  AS.sid = SID;
  const ls = { get: (k) => { try { return localStorage.getItem(k); } catch { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); return true; } catch { return false; } }, del: (k) => { try { localStorage.removeItem(k); } catch { } }, keys: () => { try { return Object.keys(localStorage); } catch { return []; } } };
  const readIdx = (k) => { try { return JSON.parse(ls.get(k)); } catch { return null; } };

  /* ---------- IndexedDB (one store, key = window id) ---------- */
  let dbp = null;
  const db = () => dbp || (dbp = new Promise((res, rej) => {
    let rq; try { rq = indexedDB.open('bts-sheets-recovery', 1); } catch (e) { rej(e); return; }
    rq.onupgradeneeded = () => rq.result.createObjectStore('slots');
    rq.onsuccess = () => res(rq.result); rq.onerror = () => rej(rq.error); rq.onblocked = () => rej(new Error('blocked'));
  }).catch(e => { dbp = null; throw e; }));
  const tx = async (mode, fn) => { const d = await db(); return new Promise((res, rej) => { const t = d.transaction('slots', mode); const st = t.objectStore('slots'); const rq = fn(st); t.oncomplete = () => res(rq && rq.result); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error || new Error('aborted')); }); };
  AS.idb = { put: (k, v) => tx('readwrite', s => s.put(v, k)), get: (k) => tx('readonly', s => s.get(k)), del: (k) => tx('readwrite', s => s.delete(k)), keys: () => tx('readonly', s => s.getAllKeys()) };

  /* ---------- snapshot of the model ---------- */
  const SHEET_PROPS = ['styles', 'colW', 'rowH', 'hiddenR', 'hiddenC', 'merges', 'comments', 'links', 'validations', 'cf', 'tables', 'filter', 'charts', 'images', 'page', 'freeze', 'rowCount', 'colCount', 'filterHidden', 'sparklines', 'threads', 'outlineR', 'outlineC', 'pivots', 'allowRanges', 'tabColor', 'showGrid', 'showHeadings', 'protected', 'protection', 'sel', 'active', 'scroll'];
  const replacer = function (k, v) {
    if (k && k[0] === '_') return undefined;
    if (typeof v === 'function') return undefined;
    const raw = this[k];
    if (raw instanceof Map) return { $m: [...raw] };
    if (raw instanceof Set) return { $s: [...raw] };
    if (raw instanceof Date) return { $d: raw.toISOString() };
    return v;
  };
  const reviver = (k, v) => { if (v && typeof v === 'object' && !Array.isArray(v)) { if (v.$m) return new Map(v.$m); if (v.$s) return new Set(v.$s); if (v.$d) return new Date(v.$d); } return v; };
  function cellsOf(wb, s) {
    let rows; try { rows = wb.hf.getSheetSerialized(s.id); } catch { return []; }
    return rows.map(row => row && row.map(v => {
      if (v === undefined || v === null || v === '') return null;
      if (typeof v === 'string') { if (v[0] === '=') return SU.fromHF(v); if (v[0] === "'") return v.slice(1); return v; }
      if (typeof v === 'object' && 'value' in v) return v.value;
      return v;
    }));
  }
  AS.serialize = (wb) => {
    const names = []; try { wb.hf.getAllNamedExpressionsSerialized().forEach(n => names.push({ name: n.name, expression: n.expression, scope: n.scope === undefined ? null : wb.sheets.findIndex(s => s.id === n.scope), options: n.options })); } catch { }
    const snap = {
      v: 1, t: Date.now(), title: BTS.title, props: BTS.props, filePath: BTS.filePath || null,
      book: { theme: wb.theme, themeOver: wb.themeOver, customStyles: wb.customStyles, styleMods: wb.styleMods, scenarios: wb.scenarios, protectStructure: wb.protectStructure, activeIndex: wb.activeIndex, wbProps: wb.props, names },
      sheets: wb.sheets.map(s => { const meta = {}; SHEET_PROPS.forEach(p => { if (s[p] !== undefined) meta[p] = s[p]; }); return { name: s.name, data: cellsOf(wb, s), meta }; }),
    };
    return JSON.stringify(snap, replacer);
  };
  /** rebuild the workbook from a snapshot string */
  AS.restore = (wb, json) => {
    const snap = JSON.parse(json, reviver);
    wb.resetBook();
    const glob = snap.book.names.filter(n => n.scope === null || n.scope === undefined || n.scope < 0).map(n => ({ name: n.name, expression: SU.fromHF(n.expression) }));
    IO.buildWorkbook(wb, snap.sheets.map(s => ({ name: s.name, data: s.data, meta: s.meta })), glob);
    for (const n of snap.book.names) if (n.scope !== null && n.scope >= 0 && wb.sheets[n.scope]) { try { wb.hf.addNamedExpression(n.name, n.expression, wb.sheets[n.scope].id, n.options); } catch (e) { console.warn('name', n.name, e); } }
    const b = snap.book; Object.assign(wb, { theme: b.theme || 'bts', themeOver: b.themeOver || null, customStyles: b.customStyles || [], styleMods: b.styleMods || {}, scenarios: b.scenarios || [], protectStructure: b.protectStructure || null });
    if (b.wbProps) wb.props = b.wbProps;
    wb.activeIndex = Math.min(Math.max(0, b.activeIndex || 0), wb.sheets.length - 1);
    wb.undoStack = []; wb.redoStack = [];
    if (snap.props) BTS.setProps(snap.props);
    return snap;
  };

  /* ---------- writing ---------- */
  let wbRef = null, timer = 0, firstPending = 0, gen = 0, writing = null;
  AS.init = (wb) => { wbRef = wb; };
  AS.schedule = () => {
    if (!wbRef) return; const now = Date.now(); if (!firstPending) firstPending = now;
    clearTimeout(timer);
    timer = setTimeout(AS.flush, now - firstPending > MAX_WAIT ? 0 : DEBOUNCE);
  };
  AS.flush = async () => {
    clearTimeout(timer); firstPending = 0;
    if (!wbRef || !BTS.dirty) return;
    if (writing) { await writing.catch(() => { }); }
    const g = gen; let json;
    try { json = AS.serialize(wbRef); } catch (e) { console.warn('autosave: serialise failed', e); return; }
    const t = Date.now();
    writing = AS.idb.put(SID, { t, title: BTS.title, json }).then(() => {
      if (g !== gen) { AS.idb.del(SID).catch(() => { }); return; } // cleared meanwhile (saved / new file)
      ls.set(LK + SID, JSON.stringify({ t, hb: Date.now(), title: BTS.title, size: json.length }));
    }).catch(e => console.warn('autosave failed', e));
    await writing; writing = null;
  };
  /* heartbeat: tells other windows this slot is still alive */
  setInterval(() => { const r = readIdx(LK + SID); if (!r) return; r.hb = Date.now(); ls.set(LK + SID, JSON.stringify(r)); }, 5000);
  AS.clear = () => { gen++; clearTimeout(timer); firstPending = 0; ls.del(LK + SID); AS.idb.del(SID).catch(() => { }); };
  AS.hasSlot = () => !!readIdx(LK + SID);

  /* ---------- recovery ---------- */
  /** slots left behind by windows that are no longer running, newest first */
  AS.orphans = () => {
    const now = Date.now(); const out = [];
    for (const k of ls.keys()) {
      if (!k.startsWith(LK) || k === LK + SID) continue;
      const r = readIdx(k); if (!r || !r.t) { ls.del(k); continue; }
      if (now - r.t > MAX_AGE) { ls.del(k); AS.idb.del(k.slice(LK.length)).catch(() => { }); continue; }
      if (now - (r.hb || r.t) < STALE) continue; // that window is still open
      out.push({ sid: k.slice(LK.length), ...r });
    }
    return out.sort((a, b) => b.t - a.t);
  };
  /* IndexedDB records without an index entry (e.g. index removed on a clean close while the delete was still pending) */
  AS.sweep = async () => {
    try { const keys = await AS.idb.keys(); for (const k of keys) { if (k === SID || readIdx(LK + k)) continue; const rec = await AS.idb.get(k); if (!rec || Date.now() - rec.t > 60000) await AS.idb.del(k); } } catch { }
  };
  AS.discard = (sid) => { ls.del(LK + sid); return AS.idb.del(sid).catch(() => { }); };
  AS.load = async (sid) => { const rec = await AS.idb.get(sid); return rec && rec.json ? rec : null; };
})();
