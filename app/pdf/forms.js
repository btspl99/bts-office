/* BTS PDF — fillable forms (AcroForm): editable boxes over the page, answers saved with pdf-lib.
 * Values live in S.form.values (field name -> string | boolean | null) until the PDF is saved. */
import { S, $, $$, esc, APP } from './state.js';

const FILLABLE = new Set(['text', 'checkbox', 'radiobutton', 'combobox', 'listbox']);
let fieldObjs = null;      // pdf.js getFieldObjects() result
let token = 0;

/** Called by app.js right after a document is loaded (before pages are built). */
export async function detect() {
  token++;
  S.form.values = new Map();
  fieldObjs = null;
  try { fieldObjs = await S.doc.getFieldObjects(); } catch {}
  S.form.has = !!fieldObjs && Object.values(fieldObjs).some(list => list.some(f => FILLABLE.has(f.type) && !f.hidden));
  document.body.classList.toggle('has-form', S.form.has);
  document.body.classList.toggle('form-hl', S.form.has && S.form.highlight);
  syncButtons();
  return S.form.has;
}
export function syncButtons() {
  $$('#ribbon [data-cmd="form-lock"]').forEach(b => b.classList.toggle('on', S.form.lock));
  $$('#ribbon [data-cmd="form-highlight"]').forEach(b => b.classList.toggle('on', S.form.has && S.form.highlight));
}

/* ---------------- on-page layer ---------------- */
const pct = (v, total) => (v / total * 100) + '%';
function boxOf(p, rect) {
  const a = p.vp1.convertToViewportPoint(rect[0], rect[1]), b = p.vp1.convertToViewportPoint(rect[2], rect[3]);
  const l = Math.min(a[0], b[0]), t = Math.min(a[1], b[1]), w = Math.abs(a[0] - b[0]), h = Math.abs(a[1] - b[1]);
  return { l, t, w, h, css: `left:${pct(l, p.vp1.width)};top:${pct(t, p.vp1.height)};width:${pct(w, p.vp1.width)};height:${pct(h, p.vp1.height)}` };
}
const initialValue = (a) => {
  if (a.fieldType === 'Btn') {
    if (a.checkBox) return !!a.fieldValue && a.fieldValue !== 'Off';
    if (a.radioButton) return a.fieldValue && a.fieldValue !== 'Off' ? a.fieldValue : null;
  }
  if (a.fieldType === 'Ch') return Array.isArray(a.fieldValue) ? (a.fieldValue[0] ?? '') : (a.fieldValue ?? '');
  return a.fieldValue ?? '';
};
const valueOf = (a) => S.form.values.has(a.fieldName) ? S.form.values.get(a.fieldName) : initialValue(a);

export async function buildLayer(p) {
  if (!S.form.has || p.formRot === S.rot) return;
  const tk = token;
  let annots;
  try { annots = p.annotsCache ||= await p.proxy.getAnnotations({ intent: 'display' }); } catch { return; }
  if (tk !== token || !S.pages.includes(p)) return;
  p.formRot = S.rot;
  $('.formLayer', p.el)?.remove();
  const widgets = annots.filter(a => a.subtype === 'Widget' && a.rect && !a.hidden && a.fieldName !== undefined);
  if (!widgets.length) return;
  const layer = document.createElement('div'); layer.className = 'formLayer';
  for (const a of widgets) {
    const B = boxOf(p, a.rect);
    const name = a.fieldName; const tip = a.alternativeText || name;
    const ro = !!a.readOnly;
    let el = null;
    if (a.fieldType === 'Tx') {
      const multi = !!a.multiLine || B.h > 40;
      el = document.createElement(multi ? 'textarea' : 'input');
      if (!multi) el.type = 'text';
      el.value = valueOf(a);
      if (a.maxLen) el.maxLength = a.maxLen;
      const da = a.defaultAppearanceData?.fontSize || 0;   // 0 = "auto" in the PDF
      const fs = multi ? Math.min(da || 10, 12) : Math.max(6, Math.min(da || 12, B.h * 0.7));
      el.style.cssText = `${B.css};font-size:calc(var(--scale-factor) * ${fs}px);text-align:${['left', 'center', 'right'][a.textAlignment || 0]}`;
      if (a.comb && a.maxLen) { el.classList.add('comb'); el.style.letterSpacing = `calc(var(--scale-factor) * ${(B.w / a.maxLen - fs * 0.6).toFixed(2)}px)`; el.style.paddingLeft = `calc(var(--scale-factor) * ${(B.w / a.maxLen / 2 - fs * 0.3).toFixed(2)}px)`; }
      el.addEventListener('input', () => setValue(name, el.value, el));
    } else if (a.fieldType === 'Btn' && (a.checkBox || a.radioButton)) {
      el = document.createElement('input');
      el.type = a.checkBox ? 'checkbox' : 'radio';
      if (a.radioButton) { el.name = 'pdfr-' + token + '-' + name; el.value = a.buttonValue; el.checked = valueOf(a) === a.buttonValue; }
      else { el.checked = !!valueOf(a); el.dataset.on = a.exportValue || 'Yes'; }
      el.style.cssText = B.css;
      el.addEventListener('change', () => setValue(name, a.checkBox ? el.checked : (el.checked ? a.buttonValue : null), el));
      // a radio button can be un-ticked by clicking it again (like paper forms)
      if (a.radioButton) el.addEventListener('click', () => { if (el.dataset.was === '1') { el.checked = false; setValue(name, null, el); } el.dataset.was = el.checked ? '1' : ''; });
      if (a.radioButton) el.addEventListener('mousedown', () => { el.dataset.was = el.checked ? '1' : ''; });
    } else if (a.fieldType === 'Ch') {
      el = document.createElement('select');
      if (!a.combo) el.size = Math.max(2, Math.min(6, (a.options || []).length));
      const v = valueOf(a);
      const opts = (a.options || []).map(o => ({ v: o.exportValue ?? o.displayValue, d: o.displayValue ?? o.exportValue }));
      if (a.combo && !opts.some(o => o.v === '')) opts.unshift({ v: '', d: '' });
      el.innerHTML = opts.map(o => `<option value="${esc(o.v)}">${esc(o.d)}</option>`).join('');
      el.value = v;
      const fs = a.defaultAppearanceData?.fontSize || Math.max(6, Math.min(11, B.h * 0.6));
      el.style.cssText = `${B.css};font-size:calc(var(--scale-factor) * ${fs}px)`;
      el.addEventListener('change', () => setValue(name, el.value, el));
    } else if (a.fieldType === 'Sig') {
      el = document.createElement('button');
      el.className = 'ff-sig'; el.type = 'button';
      el.innerHTML = `${icon('signature', 14)}<span>Sign here</span>`;
      el.style.cssText = B.css;
      el.addEventListener('click', () => window.dispatchEvent(new CustomEvent('bts-pdf-sign-field', { detail: { page: p.i, rect: a.rect } })));
    }
    if (!el) continue;
    el.classList.add('ff'); el.dataset.name = name; el.title = tip;
    if (ro) { el.disabled = true; el.classList.add('ro'); }
    el.addEventListener('keydown', e => e.stopPropagation());
    layer.appendChild(el);
  }
  p.el.appendChild(layer);
}

function setValue(name, v, src) {
  S.form.values.set(name, v);
  // keep other boxes of the same field (e.g. a name repeated on page 2) in step
  $$(`.formLayer .ff`).forEach(el => {
    if (el === src || el.dataset.name !== name) return;
    if (el.type === 'checkbox') el.checked = !!v;
    else if (el.type === 'radio') el.checked = el.value === v;
    else el.value = v ?? '';
  });
  BTS.markDirty();
}

/** All fields, reset to empty. */
export function clearAll() {
  if (!S.form.has) return BTS.toast('This PDF has no form to clear');
  for (const [name, list] of Object.entries(fieldObjs || {})) {
    const f = list.find(x => FILLABLE.has(x.type)); if (!f || f.editable === false) continue;
    S.form.values.set(name, f.type === 'checkbox' ? false : f.type === 'radiobutton' ? null : '');
  }
  S.pages.forEach(p => { p.formRot = -1; buildLayer(p); });
  BTS.markDirty();
  BTS.toast('All answers cleared');
}

export function rebuildAll() { S.pages.forEach(p => { p.formRot = -1; if (p.canvas) buildLayer(p); }); }

/** "Fill form" button: jump to the first empty box, or explain politely that there is nothing to fill. */
export async function startFilling() {
  if (!S.doc) return;
  if (!S.form.has) {
    const r = await BTS.dialog({
      title: 'No form boxes in this PDF',
      body: `<p style="margin:2px 0 6px;max-width:420px;line-height:1.55;color:var(--ink-2)">This PDF doesn’t have boxes you can type into. You can still fill it in: click <b>Add text</b> and click where you want to type, add ticks ✓, and sign it.</p>`,
      buttons: [{ label: 'OK', value: null }, { label: 'Add text', primary: true, value: 'text' }],
    });
    if (r === 'text') APP.setTool?.('textbox');
    return;
  }
  S.form.highlight = true; document.body.classList.add('form-hl'); syncButtons();
  APP.setTool?.('text');
  const first = $$('.formLayer .ff:not(.ro):not(.ff-sig)').find(el => el.type === 'checkbox' || el.type === 'radio' ? false : !el.value) || $('.formLayer .ff:not(.ro)');
  if (first) { first.scrollIntoView({ block: 'center' }); first.focus({ preventScroll: true }); }
  else {   // layers are built lazily: go to the first page that has fields
    for (const [, list] of Object.entries(fieldObjs || {})) { const f = list.find(x => FILLABLE.has(x.type)); if (f && Number.isInteger(f.page)) { APP.goTo?.(f.page); break; } }
  }
  BTS.toast('Click a box to type · Tab moves to the next box', { ms: 3000 });
}
export function toggleHighlight() {
  S.form.highlight = !S.form.highlight;
  document.body.classList.toggle('form-hl', S.form.has && S.form.highlight); syncButtons();
}
export function toggleLock() {
  S.form.lock = !S.form.lock; syncButtons();
  if (S.doc) BTS.markDirty();
  BTS.toast(S.form.lock ? 'Answers will be locked when you save (they can’t be changed later)' : 'Answers will stay editable after saving', { ms: 3200 });
}
export function fieldCount() { return fieldObjs ? Object.values(fieldObjs).filter(l => l.some(f => FILLABLE.has(f.type))).length : 0; }

/* ---------------- saving with pdf-lib ---------------- */
export function needsSave(base) {
  if (base) return S.form.values.size > 0;
  return S.form.values.size > 0 || (S.form.lock && S.form.has);
}

/** Write the answers into pdf-lib document `doc`. flatten = lock answers (fields become plain page content). */
export async function applyToDoc(doc, L, { flatten = false } = {}) {
  let form;
  try { form = doc.getForm(); } catch { return; }
  const helv = await doc.embedFont(L.StandardFonts.Helvetica);
  const charset = new Set(helv.getCharacterSet());
  const latin = (t) => [...t].every(ch => ch === '\n' || ch === '\r' || charset.has(ch.codePointAt(0)));
  const special = [];   // text fields with characters Helvetica can't draw (e.g. Hindi)
  const optionMap = new Map();
  Object.entries(fieldObjs || {}).forEach(([n, list]) => list.forEach(f => { if (f.items) optionMap.set(n, f.items); }));
  for (const [name, v] of S.form.values) {
    let f; try { f = form.getField(name); } catch { continue; }
    try {
      if (f instanceof L.PDFTextField) {
        let t = String(v ?? '');
        if (!f.isMultiline()) t = t.replace(/\r?\n/g, ' ');
        const ml = f.getMaxLength(); if (ml && t.length > ml) t = t.slice(0, ml);
        f.setText(t);
        // "auto" font size makes pdf-lib fill a tall box with huge letters: use normal text size instead
        if (f.isMultiline() && /(^|\s)0(\.0+)?\s+Tf/.test(f.acroField.getDefaultAppearance() || '')) { try { f.setFontSize(10); } catch {} }
        if (!latin(t)) special.push([f, t]);
      } else if (f instanceof L.PDFCheckBox) { v ? f.check() : f.uncheck(); }
      else if (f instanceof L.PDFRadioGroup) {
        if (!v) f.clear();
        else if (f.getOptions().includes(v)) f.select(v);
        else selectRadioByState(f, v, L);
      } else if (f instanceof L.PDFDropdown || f instanceof L.PDFOptionList) {
        if (!v) { f.clear(); continue; }
        const opts = f.getOptions();
        const items = optionMap.get(name) || [];
        const disp = items.find(it => it.exportValue === v)?.displayValue;
        if (opts.includes(v)) f.select(v);
        else if (disp && opts.includes(disp)) f.select(disp);
        else if (f instanceof L.PDFDropdown && f.isEditable()) f.select(v, true);
      }
    } catch (e) { console.warn('Could not fill field', name, e); }
  }
  for (const [f] of special) form.markFieldAsClean(f.ref);
  try { form.updateFieldAppearances(helv); } catch (e) { console.warn(e); }
  if (special.length) {
    if (flatten) {
      for (const [f, t] of special) {
        for (const w of f.acroField.getWidgets()) {
          const ref = doc.context.getObjectRef(w.dict);
          const page = ref ? doc.findPageForAnnotationRef(ref) : null;
          if (!page) continue;
          const r = w.getRectangle();
          const png = await textPicture(t, r.width, r.height, f.isMultiline());
          const img = await doc.embedPng(png);
          page.drawImage(img, { x: r.x, y: r.y, width: r.width, height: r.height });
        }
        form.removeField(f);
      }
    } else form.acroForm.dict.set(L.PDFName.of('NeedAppearances'), L.PDFBool.True);
  }
  if (flatten) form.flatten({ updateFieldAppearances: false });
}

/** Radio state names from pdf.js may differ from pdf-lib's option list: set the on-state directly. */
function selectRadioByState(f, state, L) {
  const on = L.PDFName.of(state);   // PDFName instances are interned, so === compares names
  f.acroField.dict.set(L.PDFName.of('V'), on);
  for (const w of f.acroField.getWidgets()) w.setAppearanceState(w.getOnValue() === on ? on : L.PDFName.of('Off'));
}

/** Render text (any script) into a PNG for a box of w×h points. */
async function textPicture(text, w, h, multi) {
  const k = 4; const c = document.createElement('canvas');
  c.width = Math.max(4, Math.round(w * k)); c.height = Math.max(4, Math.round(h * k));
  const x = c.getContext('2d');
  const fam = '"BTS Deva", "Noto Sans Devanagari", sans-serif';
  let fs = multi ? 10 : Math.max(6, Math.min(12, h * 0.62));
  try { await document.fonts.load(`${fs * k}px ${fam}`, text); } catch {}
  x.fillStyle = '#000'; x.textBaseline = 'middle';
  x.font = `${fs * k}px ${fam}`;
  if (!multi) {
    while (fs > 5 && x.measureText(text).width > (w - 4) * k) { fs -= 0.5; x.font = `${fs * k}px ${fam}`; }
    x.fillText(text, 2 * k, c.height / 2);
  } else {
    const lines = []; for (const para of text.split('\n')) { let line = ''; for (const word of para.split(' ')) { const t = line ? line + ' ' + word : word; if (x.measureText(t).width > (w - 4) * k && line) { lines.push(line); line = word; } else line = t; } lines.push(line); }
    lines.forEach((l, i) => x.fillText(l, 2 * k, (2 + fs * 0.7 + i * fs * 1.25) * k));
  }
  const blob = await new Promise(r => c.toBlob(r, 'image/png'));
  return new Uint8Array(await blob.arrayBuffer());
}
