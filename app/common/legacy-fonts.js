/* BTS Office — old Hindi / Marathi font converter (works offline).
 *
 * Before Unicode, Hindi and Marathi were typed in "legacy" fonts that draw Devanagari shapes on top of English
 * letters: Kruti Dev 010 (and the Kruti Dev / DevLys family, Remington typewriter layout) and Shivaji 01/02
 * (Maharashtra). Text in those fonts is gibberish anywhere else ("Hkkjr" instead of भारत). This module turns it
 * into real Unicode Devanagari, and Unicode back into Kruti Dev 010 for offices that still ask for it.
 * Tables follow the long-standing community converters (Kruti Dev 010 table of the "Krutidev ⇄ Unicode
 * converter" V24 and the Shivaji ⇄ Unicode converter by Narayan Prasad), with the vowel-sign / reph reordering
 * done by pattern rules instead of chained replaces.
 *
 * API: window.BTSLegacyFonts = {
 *   toUnicode(text, font)    font: 'krutidev' (default) | 'shivaji'
 *   fromUnicode(text, font)  font: 'krutidev' (only Kruti Dev 010 output is supported)
 *   detect(fontFamily)       'krutidev' | 'shivaji' | null  — for a CSS font-family value or a font name
 *   dialog()                 paste-in / convert / copy converter window (needs BTS.dialog)
 *   unicodeFont              suggested Unicode replacement font ('Noto Sans Devanagari')
 * }
 */
(function () {
  'use strict';
  const PI = '\uE001', PIM = '\uE002', RE = '\uE003'; // markers: ि prefix, िं prefix, reph
  const CONS = 'क-हक़-य़'; // क…ह and nukta letters
  const SIGNS = 'ा-ौॅॉॢॣँंः'; // vowel signs + ँ ं ः

  /* ---------- Kruti Dev 010 → Unicode ---------- */
  const KD = {
    // digits & quotes
    'ƒ': '१', '„': '२', '…': '३', '†': '४', '‡': '५', 'ˆ': '६', '‰': '७', 'Š': '८', '‹': '९', 'Œ': '०', 'å': '०',
    '‘': '"', '’': '"', '“': "'", '”': "'",
    // independent vowels
    'v‚': 'ऑ', 'vks': 'ओ', 'vkS': 'औ', 'vkW': 'ऑ', 'vk': 'आ', 'v': 'अ', 'b±': 'ईं', 'Ã': 'ई', 'bZ': 'ई', 'b': 'इ', 'm': 'उ', 'Å': 'ऊ', ',s': 'ऐ', ',': 'ए', '_': 'ऋ',
    // nukta letters
    'd+': 'क़', '[+': 'ख़्', 'x+': 'ग़', 'T+': 'ज़्', 't+': 'ज़', 'M+': 'ड़', '<+': 'ढ़', '¶+': 'फ़्', 'Q+': 'फ़', ';+': 'य़', 'j+': 'ऱ', 'u+': 'ऩ',
    // consonants (capital / special glyph = half form; half + k (ा stem) = full form)
    'd': 'क', 'D': 'क्', '£': 'ख्र', '[': 'ख्', 'x': 'ग', 'X': 'ग्', 'Ä': 'घ', '?': 'घ्', '³': 'ङ',
    'p': 'च', 'P': 'च्', 'N': 'छ', 't': 'ज', 'T': 'ज्', '>': 'झ', '÷': 'झ्', 'Ö': 'झ्', '¥': 'ञ',
    'V': 'ट', 'B': 'ठ', 'M': 'ड', '<': 'ढ', '.': 'ण्', 'r': 'त', 'R': 'त्', 'F': 'थ्', 'n': 'द', '/': 'ध्', 'Ë': 'ध्', 'è': 'ध्', 'u': 'न', 'U': 'न्',
    'i': 'प', 'I': 'प्', 'Q': 'फ', '¶': 'फ्', 'c': 'ब', 'C': 'ब्', 'Ò': 'भ', 'H': 'भ्', 'e': 'म', 'E': 'म्',
    ';': 'य', '¸': 'य्', 'j': 'र', 'y': 'ल', 'Y': 'ल्', 'G': 'ळ', 'Üo': 'श्व', 'Ük': 'श', 'Üz': 'श्र्', 'o': 'व', 'O': 'व्', "'": 'श्', '"': 'ष्', 'l': 'स', 'L': 'स्', 'g': 'ह',
    // conjuncts
    'Ñ': 'कृ', '—': 'कृ', 'ô': 'क्क', 'ä': 'क्त', '{': 'क्ष्', 'K': 'ज्ञ', 'ê': 'ट्ट', 'Í': 'ट्ट', 'ë': 'ट्ठ', 'Î': 'ट्ठ', 'ð': 'ठ्ठ', 'Ï': 'ड्ड', 'ì': 'ड्ड', 'ï': 'ड्ढ', 'Ô': 'ड्ढ',
    'Ù': 'त्त्', '=': 'त्र', '«': 'त्र्', '–': 'दृ', 'Ì': 'द्द', 'í': 'द्द', ')': 'द्ध', '˜': 'द्भ', 'ö': 'द्भ', '|': 'द्य', '}': 'द्व', 'é': 'न्न', '™': 'न्न्', 'ó': 'स्त्र',
    'â': 'हृ', 'à': 'ह्न', 'ã': 'ह्म', 'á': 'ह्य', 'º': 'ह्', 'J': 'श्र', 'Ø': 'क्र', 'Ý': 'फ्र', 'æ': 'द्र', 'ç': 'प्र', 'Á': 'प्र', '#': 'रु', ':': 'रू', 'Ó': '्य', 'î': '्य', 'z': '्र', 'ª': '्र',
    // vowel signs and marks
    'f': PI, 'Ç': PIM, 'È': 'ीं', 'Ê': RE + 'ी', '›': RE + 'ैं', 'õ': RE + 'ैं', '±': RE + 'ं', 'Æ': 'र्' + PI, 'É': 'र्' + PIM, 'Z': RE,
    '‚': 'ॉ', '¨': 'ो', '®': 'ो', 'kW': 'ॉ', 'ks': 'ो', '©': 'ौ', 'kS': 'ौ', 'h': 'ी', 'q': 'ु', 'w': 'ू', '`': 'ृ', 's': 'े', '¢': 'े', 'S': 'ै',
    'a': 'ं', '¡': 'ँ', '%': 'ः', 'W': 'ॅ', '•': 'ऽ', '·': 'ऽ', '∙': 'ऽ', '~': '्', '+': '़', 'k': 'ा', 'A': '।', 'ñ': '॰',
    // punctuation
    '\\': '?', '^': '‘', '*': '’', 'Þ': '“', 'ß': '”', '(': ';', '¼': '(', '½': ')', '¿': '{', 'À': '}', '¾': '=', '-': '.', '&': '-', ']': ',', '@': '/', '$': '+', 'þ': '%', 'ü': '#',
  };
  /* ---------- Shivaji 01/02 → Unicode ---------- */
  const SH = {
    '|': 'ऽ', 'K': 'ख', '#': 'ख्', 'k': 'क', '@': 'क्', 'G': 'घ्', 'g': 'ग्', '=': 'ङ', 'C': 'छ', 'c': 'च्', 'J': 'झ्', 'j': 'ज्', 'H': 'ञ्',
    'z': 'ठ', 'T': 'ट', 'Z': 'ढ', 'D': 'ड', 'N': 'ण्', 'q': 'थ्', 't': 'त', '%': 'त्', 'Q': 'ध्', 'd': 'द', 'n': 'न्',
    'f': 'फ', 'p': 'प', 'P': 'प्', 'B': 'भ्', 'b': 'ब्', 'm': 'म्', 'y': 'य्', 'r': 'र', 'l': 'ल्', 'v': 'व्', 'S': 'श्', 'Y': 'ष्', 's': 'स्', 'h': 'ह',
    'p`': 'प्र्', '`': '्र', '/': '्र', 'x': 'क्ष्', '~': 'त्र', '&': 'ज्ञ', 'V': 'द्य', 'Ù': 'न्न', 'E': 'श्र्', '(': 'ह्म', 'w': 'द्ध', 'É': 'रु', '$': 'रू', 'Ë': 'क्र', 'W': 'द्व', '<': 'त्त्', '>': 'क्त', '_': 'द्द',
    'AaO': 'औ', 'AÝ': 'औ', 'Aao': 'ओ', 'Aa': 'आ', 'A': 'अ', 'eo': 'ऐ', 'e': 'ए', 'š': 'ई', '[': 'इ', ']': 'उ', '?': 'ऋ', '¸': ',', '.': '।',
    '\\': '्', 'Ý': 'ौ', 'aO': 'ौ', 'O': 'ै', 'I': 'ी', 'U': 'ू', 'u': 'ु', 'R': 'ृ', 'ao': 'ो', 'o': 'े', 'a': 'ा', '^M': 'ँ', 'M': 'ं', 'Á': 'ः',
    '³': '(', '´': ')', 'i': PI, '-': RE,
  };
  const keysOf = (m) => Object.keys(m).sort((a, b) => b.length - a.length);
  const KD_KEYS = keysOf(KD), SH_KEYS = keysOf(SH);
  function tokenize(s, map, keys) {
    const max = keys[0].length; let out = '';
    for (let i = 0; i < s.length;) {
      let hit = null;
      for (let L = Math.min(max, s.length - i); L > 0; L--) { const k = s.substr(i, L); if (map[k] !== undefined) { hit = k; break; } }
      if (hit) { out += map[hit]; i += hit.length; } else { out += s[i]; i++; }
    }
    return out;
  }
  const reCluster = new RegExp(`([${PI}${PIM}])((?:[${CONS}]़?्)*[${CONS}]़?)`, 'g');
  const reReph1 = new RegExp(`([${CONS}]़?)([${SIGNS}]*)${RE}`, 'g');
  const reReph2 = new RegExp(`((?:[${CONS}]़?्)+)${RE}`, 'g');
  function fixUp(u, font) {
    u = u.replace(/्ा/g, '') // half form + ा stem = full consonant
      .replace(/्([ोौॉ])/g, (m, v) => ({ 'ो': 'े', 'ौ': 'ै', 'ॉ': 'ॅ' })[v]) // half + "ks" = full + े (the k is the ा stem)
      .replace(/ाे/g, 'ो').replace(/ाै/g, 'ौ').replace(/ाॅ/g, 'ॉ').replace(/े्र/g, '्रे').replace(/अौ/g, 'औ').replace(/अो/g, 'ओ').replace(/आॅ/g, 'ऑ').replace(/आे/g, 'ओ').replace(/आै/g, 'औ')
      .replace(/़़+/g, '़');
    u = u.replace(reCluster, '$2$1').replace(new RegExp(PI, 'g'), 'ि').replace(new RegExp(PIM, 'g'), 'िं'); // ि typed before the consonant
    u = u.replace(reReph1, RE + '$1$2').replace(reReph2, RE + '$1'); // reph typed after the syllable
    u = u.replace(new RegExp(RE + `(?=[${CONS}])`, 'g'), 'र्').replace(new RegExp(RE, 'g'), font === 'shivaji' ? '-' : 'र्');
    u = u.replace(/([ंँ॰])([ािीुूृेैोौॉ])/g, '$2$1').replace(/ि([़])/g, '$1ि').replace(/ं़/g, '़ं');
    u = u.replace(/(^|[\s(])ः/g, '$1:');
    return u;
  }
  function toUnicode(text, font = 'krutidev') {
    font = normFont(font) || 'krutidev';
    let s = String(text ?? '');
    if (!s) return s;
    if (font === 'shivaji') {
      s = s.replace(/iM/g, 'Mi').replace(/(^|\s)aa(?=\s|$)/g, '$1॥').replace(/(^|\s)a(?=\s|$)/g, '$1।');
      return fixUp(tokenize(s, SH, SH_KEYS), font).replace(/््/g, '्');
    }
    s = s.replace(/aa/g, 'a').replace(/ZZ/g, 'Z').replace(/=kk/g, '=k').replace(/f=k/g, 'f=').replace(/Q\+Z/g, 'QZ+').replace(/sas/g, 'sa').replace(/as/g, 'sa');
    return fixUp(tokenize(s, KD, KD_KEYS), font);
  }

  /* ---------- Unicode → Kruti Dev 010 ---------- */
  const UK = {
    "'": '*', '"': 'Þ', '—': '&', '–': '&', '‘': '^', '’': '*', '“': 'Þ', '”': 'ß', '(': '¼', ')': '½', '{': '¿', '}': 'À', '[': '¼', ']': '½', '_': '&', '-': '&', '+': '$', '=': '¾',
    '।': 'A', '॥': 'AA', '॰': 'ñ', '.': '-', '?': '\\', ',': ']', '%': 'þ', ':': '%', ';': '(', '/': '@', '#': 'ü',
    '०': 'å', '१': 'ƒ', '२': '„', '३': '…', '४': '†', '५': '‡', '६': 'ˆ', '७': '‰', '८': 'Š', '९': '‹',
    'क़्': 'D+', 'ख़्': '[+', 'ग़्': 'X+', 'ज़्': 'T+', 'ऩ्': 'U+', 'फ़्': '¶+', 'य़्': '¸+',
    'क़': 'd+', 'ख़': '[k+', 'ग़': 'x+', 'ज़': 't+', 'ड़': 'M+', 'ढ़': '<+', 'ऩ': 'u+', 'फ़': 'Q+', 'य़': ';+', 'ऱ': 'j+',
    'कृ': '—', 'क्क': 'ô', 'क्त': 'ä', 'क्ष्': '{', 'क्ष': '{k', 'ज्ञ्': 'K~', 'ज्ञ': 'K', 'ट्ट': 'ê', 'ट्ठ': 'ë', 'ड्ड': 'ì', 'ड्ढ': 'ï', 'त्त्': 'Ù', 'त्त': 'Ùk',
    'त्र्': '«', 'त्र': '=', 'दृ': '–', 'द्द': 'í', 'द्ध': ')', 'द्य': '|', 'द्व': '}', 'छ्य': 'Nî', 'ट्य': 'Vî', 'ठ्य': 'Bî', 'ड्य': 'Mî', 'ढ्य': '<î',
    'छ्र': 'Nª', 'ट्र': 'Vª', 'ठ्र': 'Bª', 'ड्र': 'Mª', 'ढ्र': '<ª', 'क्र': 'Ø', 'द्र': 'æ', 'प्र': 'iz', 'फ्र': 'Ý', 'श्र्': 'Üz', 'श्र': 'J', 'शृ': 'Ük`', 'श्व': 'Üo',
    'हृ': 'â', 'ह्न': 'à', 'ह्म': 'ã', 'ह्य': 'á', 'ह्र': 'ºz', 'ह्': 'º', '्र': 'z', 'रु': '#', 'रू': ':',
    'ओ': 'vks', 'औ': 'vkS', 'ऑ': 'v‚', 'आ': 'vk', 'अ': 'v', 'ईं': 'b±', 'ई': 'Ã', 'इ': 'b', 'उ': 'm', 'ऊ': 'Å', 'ऐ': ',s', 'ए': ',', 'ऋ': '_', 'ॲ': 'vW',
    'क्': 'D', 'क': 'd', 'ख्': '[', 'ख': '[k', 'ग्': 'X', 'ग': 'x', 'घ्': '?', 'घ': '?k', 'ङ': '³', 'च्': 'P', 'च': 'p', 'छ': 'N', 'ज्': 'T', 'ज': 't', 'झ्': '÷', 'झ': '>', 'ञ': '¥',
    'ट': 'V', 'ठ': 'B', 'ड': 'M', 'ढ': '<', 'ण्': '.', 'ण': '.k', 'त्': 'R', 'त': 'r', 'थ्': 'F', 'थ': 'Fk', 'द': 'n', 'ध्': '/', 'ध': '/k', 'न्': 'U', 'न': 'u',
    'प्': 'I', 'प': 'i', 'फ्': '¶', 'फ': 'Q', 'ब्': 'C', 'ब': 'c', 'भ्': 'H', 'भ': 'Hk', 'म्': 'E', 'म': 'e', 'य्': '¸', 'य': ';', 'र': 'j', 'ल्': 'Y', 'ल': 'y', 'ळ': 'G',
    'व्': 'O', 'व': 'o', 'श्': "'", 'श': "'k", 'ष्': '"', 'ष': '"k', 'स्': 'L', 'स': 'l', 'ह': 'g',
    'ॉ': '‚', 'ॅ': 'W', 'ो': 'ks', 'ौ': 'kS', 'ा': 'k', 'ीं': 'È', 'ी': 'h', 'ु': 'q', 'ू': 'w', 'ृ': '`', 'े': 's', 'ै': 'S', 'ं': 'a', 'ँ': '¡', 'ः': '%', 'ऽ': '·', '्': '~', '़': '+',
    [PI]: 'f', [PIM]: 'Ç', [RE]: 'Z', '\uE004': 'Æ', '\uE005': 'É', '\uE006': 'Ê', '\uE007': 'õ', '\uE008': '±',
  };
  for (const k of Object.keys(UK)) { const n = k.normalize('NFD'); if (n !== k) { if (UK[n] === undefined) UK[n] = UK[k]; delete UK[k]; } }
  const UK_KEYS = keysOf(UK);
  const reIRev = new RegExp(`((?:[${CONS}]़?्)*[${CONS}]़?)([${PI}${PIM}])`, 'g');
  const reRephRev = new RegExp(`${RE}((?:[${CONS}]़?्)*[${CONS}]़?)([${SIGNS}]*)`, 'g');
  function fromUnicode(text, font = 'krutidev') {
    font = normFont(font) || 'krutidev';
    if (font !== 'krutidev') return null;
    let s = String(text ?? '');
    s = s.replace(/[\u0929\u0931\u0934\u0958-\u095F]/g, c => c.normalize('NFD'));
    s = s.replace(/र्\u200D/g, 'र~').replace(/[\u200C\u200D]/g, '');
    s = s.replace(new RegExp(`(?<![्])र्(?=[${CONS}])`, 'g'), RE); // reph = र् starting a syllable before a consonant
    s = s.replace(/िं/g, PIM).replace(/ि/g, PI).replace(reIRev, '$2$1'); // ि goes before its consonant cluster
    s = s.replace(new RegExp(RE + PI, 'g'), '\uE004').replace(new RegExp(RE + PIM, 'g'), '\uE005');
    s = s.replace(reRephRev, '$1$2' + RE); // reph goes after the syllable
    s = s.replace(new RegExp('ी' + RE, 'g'), '\uE006').replace(new RegExp('ैं' + RE, 'g'), '\uE007').replace(new RegExp('ं' + RE, 'g'), '\uE008');
    s = s.replace(/्(?=[\s,;.।\-:?!)'"]|$)/g, '~'); // a halant at the end of a word stays visible
    s = s.replace(/~/g, '\uE009');
    return tokenize(s, UK, UK_KEYS).replace(/\uE009/g, '~');
  }

  /* ---------- detection ---------- */
  function normFont(f) { if (!f) return null; f = String(f).toLowerCase().replace(/[\s_-]/g, ''); if (/shivaji/.test(f)) return 'shivaji'; if (/kruti|kriti|devlys/.test(f)) return 'krutidev'; return null; }
  /** Legacy font of a CSS font-family value (first family) or a font name. */
  function detect(fontFamily) {
    if (!fontFamily) return null;
    const first = String(fontFamily).split(',')[0].replace(/["']/g, '').trim();
    return normFont(first);
  }

  /* ---------- dialog ---------- */
  function dialog() {
    if (!(window.BTS && BTS.dialog)) return null;
    const esc = BTS.esc;
    const body = document.createElement('div'); body.className = 'lf-dlg modal-noenter';
    body.innerHTML = `<style>.lf-dlg{display:flex;flex-direction:column;gap:10px;min-width:0}.lf-dlg .lf-row{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.lf-dlg .lf-cols{display:grid;grid-template-columns:1fr 1fr;gap:12px}.lf-dlg textarea{width:100%;height:220px;resize:vertical;box-sizing:border-box;border:1px solid var(--line);border-radius:6px;padding:8px 10px;font-size:15px;line-height:1.5;background:var(--surface);color:var(--ink)}
.lf-dlg #lf-out{font-family:'Noto Sans Devanagari','Mukta',sans-serif}.lf-dlg label{font-size:12px;color:var(--ink-2);display:block;margin-bottom:4px}.lf-dlg .hint{font-size:12px;color:var(--ink-3)}</style>
      <div class="lf-row"><label style="margin:0">Convert</label><select id="lf-mode" class="rsel"><option value="kd">Kruti Dev / DevLys → Unicode</option><option value="sh">Shivaji 01 / 02 → Unicode</option><option value="uk">Unicode → Kruti Dev 010</option></select>
        <span class="hint">Paste text on the left — the result appears on the right.</span></div>
      <div class="lf-cols"><div><label id="lf-in-l">Old font text (looks like English letters, e.g. Hkkjr)</label><textarea id="lf-in" spellcheck="false" data-indic="off"></textarea></div>
        <div><label id="lf-out-l">Unicode text (works everywhere)</label><textarea id="lf-out" spellcheck="false" readonly data-indic="off"></textarea></div></div>
      <div class="lf-row"><button class="btn" id="lf-swap">${window.icon ? icon('arrow-left-right', 15) : ''} Use result as input</button><span style="flex:1"></span><button class="btn primary" id="lf-copy">${window.icon ? icon('copy', 15) : ''} Copy result</button></div>`;
    const run = () => {
      const m = body.querySelector('#lf-mode').value; const v = body.querySelector('#lf-in').value;
      body.querySelector('#lf-out').value = m === 'uk' ? fromUnicode(v, 'krutidev') : toUnicode(v, m === 'sh' ? 'shivaji' : 'krutidev');
      body.querySelector('#lf-in-l').textContent = m === 'uk' ? 'Unicode text (e.g. भारत)' : 'Old font text (looks like English letters, e.g. Hkkjr)';
      body.querySelector('#lf-out-l').textContent = m === 'uk' ? 'Kruti Dev 010 text — paste it into a document set in Kruti Dev 010' : 'Unicode text (works everywhere)';
      body.querySelector('#lf-out').style.fontFamily = m === 'uk' ? "'Kruti Dev 010', monospace" : '';
    };
    return BTS.dialog({
      title: 'Convert old Hindi / Marathi fonts', width: 820, body, buttons: [{ label: 'Close', value: null }],
      onOpen: (root) => {
        const inp = root.querySelector('#lf-in'); inp.addEventListener('input', run); root.querySelector('#lf-mode').addEventListener('change', run);
        root.querySelector('#lf-swap').addEventListener('click', () => { inp.value = root.querySelector('#lf-out').value; const sel = root.querySelector('#lf-mode'); sel.value = sel.value === 'uk' ? 'kd' : 'uk'; run(); });
        root.querySelector('#lf-copy').addEventListener('click', async () => { const t = root.querySelector('#lf-out').value; try { await navigator.clipboard.writeText(t); } catch { const o = root.querySelector('#lf-out'); o.select(); document.execCommand('copy'); } BTS.toast && BTS.toast('Copied'); });
        setTimeout(() => inp.focus(), 30);
      },
    });
  }

  window.BTSLegacyFonts = { toUnicode, fromUnicode, detect, dialog, unicodeFont: 'Noto Sans Devanagari', fontKey: normFont };
})();
