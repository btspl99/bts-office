/* BTS Office — Hindi & Marathi phonetic typing (works offline, no IME needed).
 *
 * Include this file on a page; nothing happens until it is switched on. While on, typing Latin letters in any
 * contenteditable, <textarea> or text <input> produces Devanagari, word by word (the word you are typing is
 * re-transliterated on every key and committed on space, punctuation or Enter):
 *   namaste → नमस्ते · bhaarat / bharat → भारत · kRShNa → कृष्ण · shrI → श्री · dnyaan / jnaan → ज्ञान
 * Scheme: ITRANS-like and forgiving (aa/A, ii/ee/I, uu/oo/U, T Th D Dh N for retroflex, sh/Sh, x = क्ष,
 * M = ं, H = ः, .N = ँ, .h = ्, O = ऑ, | = ।, || = ॥), plus a small dictionary of everyday words typed the
 * casual way. Marathi: L = ळ, z = झ (maza → माझा), ~a/E = ॲ, r^ = eyelash र (र्\u200D), Devanagari digits optional.
 * Backspace inside the word removes the last Latin letter; Esc turns the current word back into English
 * letters. Ctrl+Alt+H = Hindi, Ctrl+Alt+M = Marathi, the same shortcut again = off.
 *
 * API: window.BTSIndic = {
 *   enable(lang|false)     'hi' | 'mr' | false (also accepts 'hindi' / 'marathi')
 *   lang                    current language or null when off
 *   isOn()
 *   translit(str, lang)     convert a whole string (non-letters are kept)
 *   indicatorEl()           a small pill ("अ हिंदी") with a menu; place it anywhere (status bar etc.)
 *   onChange(cb)            cb(lang|null) whenever typing is switched on/off or the language changes
 *   setOptions({ digits: 'latin'|'deva' }), options
 *   setHandler(fn)          app hook: fn(el, deleteCount, text, isNewWord) → true if the app did the replace
 *                           itself (BTS Docs uses this for undo + track changes)
 *   skip                    CSS selector of fields to leave alone (default: [data-indic="off"], #ribbon input …)
 * }
 */
(function () {
  'use strict';
  /* ---------------- transliteration tables ---------------- */
  // vowels: [independent, matra]
  const V = {
    a: ['अ', ''], aa: ['आ', 'ा'], A: ['आ', 'ा'], i: ['इ', 'ि'], ii: ['ई', 'ी'], ee: ['ई', 'ी'], I: ['ई', 'ी'],
    u: ['उ', 'ु'], uu: ['ऊ', 'ू'], oo: ['ऊ', 'ू'], U: ['ऊ', 'ू'], R: ['ऋ', 'ृ'], RRi: ['ऋ', 'ृ'], 'R^i': ['ऋ', 'ृ'],
    RRI: ['ॠ', 'ॄ'], 'R^I': ['ॠ', 'ॄ'], e: ['ए', 'े'], ai: ['ऐ', 'ै'], ei: ['ऐ', 'ै'], o: ['ओ', 'ो'], au: ['औ', 'ौ'], ou: ['औ', 'ौ'],
    O: ['ऑ', 'ॉ'], '~o': ['ऑ', 'ॉ'], '~a': ['ॲ', 'ॅ'], 'A.c': ['ॲ', 'ॅ'],
  };
  const C = {
    k: 'क', kh: 'ख', g: 'ग', gh: 'घ', '~N': 'ङ', 'N^': 'ङ', c: 'च', ch: 'च', chh: 'छ', Ch: 'छ', j: 'ज', jh: 'झ', '~n': 'ञ', JN: 'ञ',
    T: 'ट', Th: 'ठ', D: 'ड', Dh: 'ढ', N: 'ण', t: 'त', th: 'थ', d: 'द', dh: 'ध', n: 'न', p: 'प', ph: 'फ', f: 'फ़', b: 'ब', bh: 'भ', m: 'म',
    y: 'य', r: 'र', l: 'ल', L: 'ळ', v: 'व', w: 'व', sh: 'श', Sh: 'ष', shh: 'ष', s: 'स', h: 'ह',
    x: 'क्ष', ksh: 'क्ष', kSh: 'क्ष', GY: 'ज्ञ', dny: 'ज्ञ', 'j~n': 'ज्ञ', jny: 'ज्ञ', gny: 'ज्ञ',
    q: 'क़', z: 'ज़', '.D': 'ड़', '.Dh': 'ढ़', Rh: 'ड़', '.kh': 'ख़', '.g': 'ग़',
  };
  // signs that follow a syllable
  const S = { M: 'ं', '.n': 'ं', '.m': 'ं', '.N': 'ँ', '~': 'ँ', H: 'ः', '.h': '्', '.a': 'ऽ', '|': '।', '||': '॥', OM: 'ॐ', AUM: 'ॐ', 'r^': 'R^' };
  const MR = { z: 'झ', E: ['ॲ', 'ॅ'] };
  const DIG = '०१२३४५६७८९';
  const KEYS = [...new Set([...Object.keys(V), ...Object.keys(C), ...Object.keys(S), 'E'])].sort((a, b) => b.length - a.length);
  const MAXK = KEYS[0].length;
  const KEYSET = new Set(KEYS);

  /* everyday words typed the casual way (lowercase only; strict spellings always work too) */
  const LEX_COMMON = 'bharat=भारत,bharath=भारत,india=इंडिया,hindi=हिंदी,marathi=मराठी,maharashtra=महाराष्ट्र,maharashtr=महाराष्ट्र,mumbai=मुंबई,pune=पुणे,delhi=दिल्ली,dilli=दिल्ली,thane=ठाणे,bhiwandi=भिवंडी,nagpur=नागपूर,nashik=नाशिक,' +
    'namaste=नमस्ते,namaskar=नमस्कार,dhanyavad=धन्यवाद,dhanyawad=धन्यवाद,dhanyavaad=धन्यवाद,swagat=स्वागत,shubh=शुभ,shri=श्री,shree=श्री,shrimati=श्रीमती,ji=जी,' +
    'gyan=ज्ञान,gyaan=ज्ञान,vigyan=विज्ञान,vigyaan=विज्ञान,shiksha=शिक्षा,vidyalaya=विद्यालय,vidyalay=विद्यालय,vidyarthi=विद्यार्थी,shikshak=शिक्षक,pariksha=परीक्षा,prashn=प्रश्न,uttar=उत्तर,' +
    'pradhanmantri=प्रधानमंत्री,mantri=मंत्री,sarkar=सरकार,desh=देश,rashtra=राष्ट्र,rajya=राज्य,bhasha=भाषा,pustak=पुस्तक,samay=समय,jeevan=जीवन,jivan=जीवन,' +
    'krishna=कृष्ण,krishn=कृष्ण,ram=राम,shiv=शिव,ganesh=गणेश,ganpati=गणपती,ganga=गंगा,diwali=दिवाली,holi=होली,hindu=हिंदू,hindustan=हिंदुस्तान,' +
    'school=स्कूल,college=कॉलेज,doctor=डॉक्टर,computer=कंप्यूटर,mobile=मोबाइल,internet=इंटरनेट,' +
    'sundar=सुंदर,prem=प्रेम,mitra=मित्र,mata=माता,pita=पिता,parivar=परिवार,shahar=शहर,janmdin=जन्मदिन,abhinandan=अभिनंदन,jai=जय,jay=जय';
  const LEX = {
    hi: LEX_COMMON + ',main=मैं,hum=हम,tum=तुम,aap=आप,hai=है,hain=हैं,tha=था,thi=थी,the=थे,ka=का,ke=के,ko=को,se=से,mein=में,men=में,par=पर,aur=और,ya=या,' +
      'nahi=नहीं,nahin=नहीं,kya=क्या,kyon=क्यों,kyun=क्यों,kaise=कैसे,kaun=कौन,kab=कब,kahan=कहाँ,yah=यह,vah=वह,woh=वह,yeh=ये,wo=वो,' +
      'ek=एक,do=दो,teen=तीन,char=चार,paanch=पाँच,panch=पाँच,das=दस,sau=सौ,hazar=हज़ार,kitab=किताब,ghar=घर,pani=पानी,khana=खाना,raat=रात,din=दिन,saal=साल,varsh=वर्ष,' +
      'log=लोग,accha=अच्छा,achha=अच्छा,acha=अच्छा,bahut=बहुत,sab=सब,kuch=कुछ,kuchh=कुछ,abhi=अभी,phir=फिर,lekin=लेकिन,kyonki=क्योंकि,isliye=इसलिए,agar=अगर,toh=तो,bhi=भी,jo=जो,' +
      'apna=अपना,apni=अपनी,mera=मेरा,meri=मेरी,mere=मेरे,tera=तेरा,hamara=हमारा,tumhara=तुम्हारा,uska=उसका,unka=उनका,kaam=काम,naam=नाम,bhai=भाई,behan=बहन,beta=बेटा,beti=बेटी,dost=दोस्त,' +
      'pyar=प्यार,duniya=दुनिया,gaon=गाँव,bada=बड़ा,badi=बड़ी,chhota=छोटा,chota=छोटा,ladka=लड़का,ladki=लड़की,padhai=पढ़ाई,karna=करना,raha=रहा,rahi=रही,rahe=रहे,gaya=गया,gayi=गई,hua=हुआ,hui=हुई,' +
      'sakta=सकता,chahiye=चाहिए,badhai=बधाई,adhyapak=अध्यापक',
    mr: LEX_COMMON + ',maza=माझा,majha=माझा,mazi=माझी,majhi=माझी,maze=माझे,majhe=माझे,tumcha=तुमचा,tumchi=तुमची,aahe=आहे,ahe=आहे,aahet=आहेत,nahi=नाही,kay=काय,kasa=कसा,kashi=कशी,kase=कसे,' +
      'mi=मी,tu=तू,ti=ती,aamhi=आम्ही,amhi=आम्ही,tumhi=तुम्ही,ani=आणि,aani=आणि,pan=पण,shala=शाळा,shaala=शाळा,ghari=घरी,aaj=आज,udya=उद्या,mulga=मुलगा,mulgi=मुलगी,aai=आई,baba=बाबा,' +
      'kolhapur=कोल्हापूर,bappa=बाप्पा,morya=मोरया,shivaji=शिवाजी,maharaj=महाराज,kiti=किती,kuthe=कुठे,ithe=इथे,tithe=तिथे,karto=करतो,karte=करते,zala=झाला,jhala=झाला,zali=झाली,jhali=झाली,' +
      'hota=होता,hoti=होती,pani=पाणी,lok=लोक,khup=खूप,chan=छान,ganpati=गणपती,mahiti=माहिती,vachan=वाचन',
  };
  const lexOf = {};
  for (const k in LEX) { lexOf[k] = {}; LEX[k].split(',').forEach(p => { const [a, b] = p.split('='); if (a && b) lexOf[k][a] = b; }); }

  function normLang(l) { if (!l) return null; l = String(l).toLowerCase(); return l.startsWith('m') ? 'mr' : 'hi'; }
  const opts = { digits: 'latin' };

  /** Transliterate one word (letters and scheme symbols only). */
  function word(w, lang) {
    if (!w) return '';
    if (lang !== 'mr' && lang !== 'hi') lang = 'hi';
    const lx = lexOf[lang][w]; if (lx && w === w.toLowerCase()) return lx;
    let out = '', cons = false, i = 0;
    const atStart = (j) => j === 0;
    while (i < w.length) {
      const ch = w[i];
      if (/[0-9]/.test(ch)) { out += opts.digits === 'deva' ? DIG[+ch] : ch; cons = false; i++; continue; }
      let key = null, val = null, kind = null;
      for (let L = Math.min(MAXK, w.length - i); L > 0 && !key; L--) {
        const cand = w.substr(i, L);
        if (lang === 'mr' && MR[cand] !== undefined) { key = cand; val = MR[cand]; kind = Array.isArray(val) ? 'v' : 'c'; }
        else if (cand === 'E') { key = cand; val = ['ए', 'े']; kind = 'v'; }
        else if (V[cand]) { key = cand; val = V[cand]; kind = 'v'; }
        else if (C[cand]) { key = cand; val = C[cand]; kind = 'c'; }
        else if (S[cand]) { key = cand; val = S[cand]; kind = 's'; }
      }
      if (w.substr(i, 2) === 'jn' && atStart(i) && (!key || key.length < 2)) { key = 'jn'; val = 'ज्ञ'; kind = 'c'; }
      if (!key) {
        // forgiving: unknown capital → its small letter (B → b, P → p …)
        const lo = ch.toLowerCase();
        if (lo !== ch) { const rest = lo + w.slice(i + 1); const sub = word(rest, lang === 'mr' ? 'mr' : 'hi'); return out + joinCons(cons, sub); }
        out += ch; cons = false; i++; continue;
      }
      i += key.length;
      if (kind === 'c') {
        // forgiving: श before a retroflex is always ष (mahaaraashTra → महाराष्ट्र, kRshNa → कृष्ण)
        if (cons && /^[टठडढण]/.test(val) && out.endsWith('श')) out = out.slice(0, -1) + 'ष';
        out += (cons ? '्' : '') + val; cons = true; continue;
      }
      if (kind === 'v') { out += cons ? val[1] : val[0]; cons = false; continue; }
      // signs
      if (val === 'R^') { out += (cons ? '्' : '') + 'र्\u200D'; cons = false; continue; }
      if (val === '्') { if (cons) out += '्'; cons = false; continue; }
      out += val; cons = false;
    }
    return out;
  }
  // Join the translit of the rest of a word (computed separately) to a prefix that may end in a consonant.
  function joinCons(cons, sub) { if (!cons || !sub) return sub; const f = sub[0]; return /[क-हक़-य़ळ]/.test(f) ? '्' + sub : vowelToMatra(sub); }
  function vowelToMatra(s) { const M = { 'अ': '', 'आ': 'ा', 'इ': 'ि', 'ई': 'ी', 'उ': 'ु', 'ऊ': 'ू', 'ऋ': 'ृ', 'ए': 'े', 'ऐ': 'ै', 'ओ': 'ो', 'औ': 'ौ', 'ऑ': 'ॉ', 'ॲ': 'ॅ' }; return M[s[0]] !== undefined ? M[s[0]] + s.slice(1) : s; }

  const WORD_RE = /[A-Za-z0-9.~^|]+/g;
  const isLetter = (c) => /[A-Za-z]/.test(c);
  /** Transliterate a whole string. Only runs of letters (with the scheme's . ~ ^ | marks) change. */
  function translit(str, lang) {
    lang = normLang(lang) || 'hi';
    return String(str).replace(WORD_RE, (w) => {
      if (!/[A-Za-z|]/.test(w)) return opts.digits === 'deva' ? w.replace(/[0-9]/g, d => DIG[+d]) : w;
      // a trailing full stop is punctuation, not part of the word
      const m = w.match(/^(.*?)(\.*)$/); return word(m[1], lang) + m[2];
    });
  }

  /* ---------------- state & UI ---------------- */
  const KEY = 'bts-indic-lang', OPTKEY = 'bts-indic-opts';
  let lang = null; let lastLang = 'hi';
  try { lastLang = normLang(localStorage.getItem(KEY)) || 'hi'; Object.assign(opts, JSON.parse(localStorage.getItem(OPTKEY) || '{}')); } catch { }
  const listeners = [];
  let pill = null; let handler = null;
  const NAMES = { hi: 'हिंदी', mr: 'मराठी' };
  function enable(l) {
    const nl = l === false || l == null || l === 'off' ? null : normLang(l);
    reset();
    lang = nl; if (nl) { lastLang = nl; try { localStorage.setItem(KEY, nl); } catch { } }
    API.lang = lang;
    renderPill();
    listeners.forEach(cb => { try { cb(lang); } catch (e) { console.error(e); } });
    return lang;
  }
  function renderPill() {
    if (!pill) return;
    pill.classList.toggle('on', !!lang);
    pill.innerHTML = lang ? `<b>अ</b><span>${NAMES[lang]}</span>` : `<b>A</b><span>English</span>`;
    pill.title = lang ? `Typing in ${lang === 'hi' ? 'Hindi' : 'Marathi'} — type in English letters, get Devanagari. Click for options (Ctrl+Alt+${lang === 'hi' ? 'H' : 'M'} to turn off)` : 'Type in Hindi or Marathi using English letters (Ctrl+Alt+H / Ctrl+Alt+M)';
  }
  function menuItems() {
    return [
      { head: 'Typing language' },
      { label: 'English (off)', checked: !lang, action: () => enable(false) },
      { label: 'हिंदी — Hindi', kbd: 'Ctrl+Alt+H', checked: lang === 'hi', action: () => enable('hi') },
      { label: 'मराठी — Marathi', kbd: 'Ctrl+Alt+M', checked: lang === 'mr', action: () => enable('mr') },
      { sep: true },
      { label: 'Devanagari digits (१२३)', checked: opts.digits === 'deva', action: () => setOptions({ digits: opts.digits === 'deva' ? 'latin' : 'deva' }) },
      { label: 'How to type…', icon: 'circle-help', action: help },
    ];
  }
  function indicatorEl() {
    if (pill) return pill;
    pill = document.createElement('button'); pill.type = 'button'; pill.className = 'indic-pill'; pill.setAttribute('data-indic', 'off');
    if (!document.getElementById('indic-css')) {
      const st = document.createElement('style'); st.id = 'indic-css';
      st.textContent = `.indic-pill{display:inline-flex;align-items:center;gap:5px;height:22px;padding:0 9px 0 7px;border:1px solid var(--line,#d5d9e2);border-radius:11px;background:var(--surface,#fff);color:var(--ink-2,#47526a);font:inherit;font-size:12px;cursor:pointer;white-space:nowrap}
.indic-pill b{font-family:'Noto Sans Devanagari','Mukta',sans-serif;font-weight:700;font-size:13px;line-height:1}.indic-pill:hover{border-color:var(--accent,#3C4FE0)}
.indic-pill.on{background:var(--accent-soft,#e8ebff);border-color:var(--accent,#3C4FE0);color:var(--accent,#3C4FE0)}html[data-theme="dark"] .indic-pill.on{color:color-mix(in srgb,var(--accent,#3C4FE0) 40%,#fff)}.indic-pill span{font-family:'Noto Sans Devanagari','Mukta',inherit}
.indic-menu{position:fixed;z-index:9999;background:var(--surface,#fff);color:var(--ink,#1f2533);border:1px solid var(--line,#d5d9e2);border-radius:8px;box-shadow:0 8px 28px rgba(0,0,0,.18);padding:4px;min-width:220px;font-size:13px}
.indic-menu button{display:flex;width:100%;gap:8px;align-items:center;border:0;background:none;color:inherit;font:inherit;padding:6px 10px;border-radius:5px;text-align:left;cursor:pointer}.indic-menu button:hover{background:var(--accent-soft,#e8ebff)}
.indic-menu .h{font-size:11px;color:var(--ink-3,#7a8399);padding:6px 10px 2px}.indic-menu hr{border:0;border-top:1px solid var(--line,#d5d9e2);margin:4px 0}.indic-menu kbd{margin-left:auto;font:inherit;font-size:11px;color:var(--ink-3,#7a8399)}`;
      document.head.appendChild(st);
    }
    pill.addEventListener('mousedown', e => e.preventDefault()); // keep the caret where it is
    pill.addEventListener('click', () => {
      if (window.BTS && BTS.menu) { BTS.menu(pill, menuItems()); return; }
      ownMenu(pill, menuItems());
    });
    renderPill();
    return pill;
  }
  function ownMenu(anchor, items) {
    document.querySelectorAll('.indic-menu').forEach(m => m.remove());
    const m = document.createElement('div'); m.className = 'indic-menu';
    items.forEach(it => {
      if (it.sep) { m.appendChild(document.createElement('hr')); return; }
      if (it.head) { const h = document.createElement('div'); h.className = 'h'; h.textContent = it.head; m.appendChild(h); return; }
      const b = document.createElement('button'); b.textContent = (it.checked ? '✓ ' : '   ') + it.label;
      if (it.kbd) { const k = document.createElement('kbd'); k.textContent = it.kbd; b.appendChild(k); }
      b.addEventListener('mousedown', e => e.preventDefault());
      b.addEventListener('click', () => { m.remove(); it.action && it.action(); }); m.appendChild(b);
    });
    document.body.appendChild(m);
    const r = anchor.getBoundingClientRect(); const h = m.offsetHeight;
    m.style.left = Math.min(r.left, innerWidth - m.offsetWidth - 8) + 'px'; m.style.top = (r.top - h - 6 > 0 ? r.top - h - 6 : r.bottom + 6) + 'px';
    setTimeout(() => document.addEventListener('mousedown', function off(e) { if (!m.contains(e.target)) { m.remove(); document.removeEventListener('mousedown', off, true); } }, true), 0);
  }
  function help() {
    const rows = [['namaste', 'नमस्ते'], ['bhaarat or bharat', 'भारत'], ['kRShNa', 'कृष्ण'], ['shrI', 'श्री'], ['dnyaan / jnaan / GYaan', 'ज्ञान'], ['mahaaraashTra', 'महाराष्ट्र'], ['hiMdI / hindii', 'हिंदी / हिन्दी'], ['shaaLaa (Marathi)', 'शाळा'], ['DOkTar', 'डॉक्टर'], ['raamaH . |', 'रामः . ।']];
    const keys = [['a aa/A i ii/ee/I u uu/oo/U', 'अ आ इ ई उ ऊ'], ['e ai o au O R', 'ए ऐ ओ औ ऑ ऋ'], ['T Th D Dh N', 'ट ठ ड ढ ण'], ['t th d dh n', 'त थ द ध न'], ['sh Sh s h L', 'श ष स ह ळ'], ['x / ksh, GY / dny', 'क्ष, ज्ञ'], ['M  .N  H  .h', 'ं  ँ  ः  ्'], ['q K G z f .D', 'क़ ख़ ग़ ज़ फ़ ड़'], ['Marathi: z  ~a  r^', 'झ  ॲ  र्\u200D'], ['| ||', '। ॥']];
    const tbl = (rs) => '<table style="border-collapse:collapse;width:100%">' + rs.map(([a, b]) => `<tr><td style="padding:3px 10px 3px 0;font-family:Consolas,monospace;color:var(--ink-2)">${a}</td><td style="padding:3px 0;font-family:'Noto Sans Devanagari',sans-serif;font-size:15px">${b}</td></tr>`).join('') + '</table>';
    const body = `<div style="display:grid;grid-template-columns:1fr 1fr;gap:18px;font-size:13px"><div><b>Examples</b>${tbl(rows)}</div><div><b>Keys</b>${tbl(keys)}</div></div>
      <p style="font-size:12.5px;color:var(--ink-2);margin:12px 0 0">Type a word in English letters — it changes to Devanagari as you type and is fixed when you press Space. <b>Backspace</b> removes the last letter you typed, <b>Esc</b> keeps the word in English. Turn it on or off with <b>Ctrl+Alt+H</b> (Hindi) or <b>Ctrl+Alt+M</b> (Marathi).</p>`;
    if (window.BTS && BTS.dialog) BTS.dialog({ title: 'Typing in Hindi and Marathi', width: 680, body, buttons: [{ label: 'Close', value: 1, primary: true }] });
  }
  function setOptions(o) { Object.assign(opts, o || {}); try { localStorage.setItem(OPTKEY, JSON.stringify(opts)); } catch { } listeners.forEach(cb => { try { cb(lang); } catch { } }); }

  /* ---------------- live typing engine ---------------- */
  let st = { el: null, buf: '', out: '', literal: false }; let busy = false;
  function reset() { st = { el: null, buf: '', out: '', literal: false }; }
  const TEXT_TYPES = new Set(['text', 'search', '']);
  function editableOf(t) {
    if (!t || t.nodeType !== 1) t = t && t.parentElement; if (!t) return null;
    if (t.closest && t.closest(API.skip)) return null;
    if (t.tagName === 'TEXTAREA') return t.readOnly || t.disabled ? null : t;
    if (t.tagName === 'INPUT') return TEXT_TYPES.has((t.getAttribute('type') || '').toLowerCase()) && !t.readOnly && !t.disabled && t.inputMode !== 'numeric' && t.inputMode !== 'decimal' ? t : null;
    if (t.isContentEditable) { let h = t; while (h.parentElement && h.parentElement.isContentEditable) h = h.parentElement; return h; }
    return null;
  }
  const isField = (el) => el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA');
  /** Text right before the caret (up to n chars) and the Range covering it (contenteditable). */
  function before(el, n) {
    if (isField(el)) { const a = el.selectionStart, b = el.selectionEnd; if (a == null || a !== b) return null; return { text: el.value.slice(Math.max(0, a - n), a), pos: a }; }
    const s = getSelection(); if (!s.rangeCount) return null; const r = s.getRangeAt(0); if (!r.collapsed || !el.contains(r.startContainer)) return null;
    let node = r.startContainer, off = r.startOffset;
    if (node.nodeType !== 3) { // caret between nodes: step to the text before it
      let p = node.childNodes[off - 1] || null;
      if (!p) { node = prevText(el, node); off = node ? node.length : 0; }
      else { while (p && p.nodeType === 1 && p.lastChild) p = p.lastChild; if (p && p.nodeType === 3) { node = p; off = p.length; } else { node = prevText(el, p); off = node ? node.length : 0; } }
    }
    const range = document.createRange(); range.setEnd(r.startContainer, r.startOffset);
    let text = ''; let cur = node, co = off;
    if (!cur) { range.collapse(false); return { text: '', range }; }
    range.setStart(cur, co);
    while (cur && text.length < n) {
      const take = Math.min(co, n - text.length);
      text = cur.data.slice(co - take, co) + text; range.setStart(cur, co - take);
      if (text.length >= n) break;
      const pv = prevText(el, cur); if (!pv) break;
      // stop at block boundaries
      if (blockOf(pv, el) !== blockOf(cur, el)) break;
      cur = pv; co = pv.length;
    }
    return { text, range };
  }
  function blockOf(n, host) { let e = n.parentElement; while (e && e !== host) { const d = getComputedStyle(e).display; if (d !== 'inline' && d !== 'contents') return e; e = e.parentElement; } return host; }
  function prevText(host, node) {
    const w = document.createTreeWalker(host, NodeFilter.SHOW_TEXT); w.currentNode = node;
    let p = w.previousNode(); return p && p !== node ? p : null;
  }
  /** Replace the del chars before the caret with text. */
  function replace(el, del, text, newWord) {
    busy = true;
    try {
      if (handler && handler(el, del, text, newWord)) return;
      if (isField(el)) {
        const a = el.selectionStart; el.setSelectionRange(a - del, a);
        if (!document.execCommand('insertText', false, text) || el.value.slice(a - del, a - del + text.length) !== text) {
          el.setRangeText(text, a - del, a, 'end'); el.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: text }));
        }
        return;
      }
      if (del) { const b = before(el, del); if (b && b.range) { const s = getSelection(); s.removeAllRanges(); s.addRange(b.range); } }
      if (text) document.execCommand('insertText', false, text); else if (del) document.execCommand('delete', false);
    } finally { busy = false; }
  }
  const continuing = (el) => { if (!st.buf || st.el !== el) return false; const b = before(el, st.out.length); return !!b && b.text === st.out; };
  const wordChar = (c, inWord) => isLetter(c) || c === '|' || (inWord && /[.~^]/.test(c));
  const render = (buf, literal) => literal ? buf : word(buf, lang);

  function onBeforeInput(e) {
    if (!lang || busy) return;
    if (e.isComposing || e.inputType === 'insertCompositionText') { reset(); return; }
    const el = editableOf(e.target); if (!el) return;
    if (e.inputType !== 'insertText' || !e.data) { reset(); return; }
    if (isField(el) && /^=/.test(el.value) && el.selectionStart > 0) return; // spreadsheet formulas stay in English
    const ch = e.data;
    const cont = continuing(el);
    if (ch.length !== 1 || !wordChar(ch, cont)) {
      // digits: Devanagari digits if asked for
      if (ch.length === 1 && /[0-9]/.test(ch) && opts.digits === 'deva') { e.preventDefault(); e.stopPropagation(); replace(el, 0, DIG[+ch], true); }
      // a trailing '.' that the scheme did not use is punctuation: keep it as typed
      reset(); return;
    }
    if (!cont) reset();
    const buf = st.buf + ch;
    const out = render(buf, st.literal);
    e.preventDefault(); e.stopPropagation();
    replace(el, st.out.length, out, !st.buf);
    st = { el, buf, out, literal: st.literal };
  }
  function onKeyDown(e) {
    if (e.ctrlKey && e.altKey && !e.shiftKey && !e.metaKey && (e.code === 'KeyH' || e.code === 'KeyM' || /^[hm]$/i.test(e.key))) {
      const want = (e.code === 'KeyH' || /^h$/i.test(e.key)) ? 'hi' : 'mr';
      e.preventDefault(); e.stopPropagation();
      enable(lang === want ? false : want);
      if (window.BTS && BTS.toast) BTS.toast(lang ? `Typing in ${lang === 'hi' ? 'Hindi (हिंदी)' : 'Marathi (मराठी)'} — type in English letters` : 'Typing in English');
      return;
    }
    if (!lang || !st.buf) return;
    if (e.isComposing) { reset(); return; }
    const el = editableOf(e.target);
    if (e.key === 'Backspace' && !e.ctrlKey && !e.altKey && !e.metaKey && el && continuing(el)) {
      e.preventDefault(); e.stopPropagation();
      const buf = st.buf.slice(0, -1); const out = buf ? render(buf, st.literal) : '';
      replace(el, st.out.length, out, false);
      if (buf) st = { el, buf, out, literal: st.literal }; else reset();
      return;
    }
    if (e.key === 'Escape' && el && continuing(el) && !st.literal) {
      e.preventDefault(); e.stopPropagation();
      replace(el, st.out.length, st.buf, false); st = { el, buf: st.buf, out: st.buf, literal: true };
      return;
    }
    if (e.key.length > 1 && !/^(Shift|Control|Alt|Meta|CapsLock|Backspace|Escape)$/.test(e.key)) reset();
    else if ((e.ctrlKey || e.metaKey) && !/^(Control|Meta|Shift|Alt)$/.test(e.key)) reset();
  }
  document.addEventListener('beforeinput', onBeforeInput, true);
  document.addEventListener('keydown', onKeyDown, true);
  document.addEventListener('mousedown', (e) => { if (!pill || !pill.contains(e.target)) reset(); }, true);
  document.addEventListener('compositionstart', reset, true);
  document.addEventListener('focusin', () => reset(), true);

  const API = window.BTSIndic = {
    lang: null,
    enable, isOn: () => !!lang, translit, word: (w, l) => word(w, normLang(l) || 'hi'),
    indicatorEl, onChange: (cb) => { listeners.push(cb); return () => { const i = listeners.indexOf(cb); if (i >= 0) listeners.splice(i, 1); }; },
    setOptions, get options() { return Object.assign({}, opts); },
    setHandler: (fn) => { handler = fn; },
    rangeBefore: (el, n) => before(el, n),
    /* App hook: the last `latin` characters before the caret in `el` were typed while focus was elsewhere
       (e.g. the first key that opens a spreadsheet cell editor) — convert them and continue the word from there. */
    adopt: (el, latin) => { if (!lang || !el || !latin) return; const b = before(el, latin.length); if (!b || b.text !== latin) return; const out = render(latin, false); replace(el, latin.length, out, true); st = { el, buf: latin, out, literal: false }; },
    menuItems, help, lastLang: () => lastLang,
    toggle: (l) => enable(lang ? false : (normLang(l) || lastLang)),
    skip: '[data-indic="off"], #ribbon input, #ribbon select, input[type="number"], #gkeys',
  };
})();
