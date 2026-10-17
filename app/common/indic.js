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
  const opts = { digits: 'latin', layout: 'phonetic', suggest: true };

  /* Keyboard layouts.
   *  phonetic  — type how it sounds in English letters (like Microsoft Indic Language Input Tool / Google Input Tools)
   *  inscript  — the Government of India standard InScript layout, as in Windows' "Hindi Traditional" / "Marathi" keyboards
   *  remington — Remington (GAIL) typewriter layout, the one Kruti Dev typists and government typing exams use */
  const LAYOUTS = [
    { id: 'phonetic', label: 'Phonetic (type how it sounds)', short: 'Phonetic', desc: 'namaste → नमस्ते. Suggestions as you type.' },
    { id: 'inscript', label: 'InScript (Microsoft standard)', short: 'InScript', desc: 'The Windows Hindi / Marathi keyboard: k = क, e = ा, d = ्.' },
    { id: 'remington', label: 'Remington (GAIL) — Kruti Dev typing', short: 'Remington', desc: 'Typewriter layout used with Kruti Dev: Hkkjr → भारत.' },
  ];
  // InScript, keyed by the character a US keyboard produces (normal / Shift).
  const INSCRIPT = {
    '`': 'ॊ', '~': 'ऒ', '1': '१', '!': 'ऍ', '2': '२', '@': 'ॅ', '3': '३', '#': '्र', '4': '४', '$': 'र्', '5': '५', '%': 'ज्ञ',
    '6': '६', '^': 'त्र', '7': '७', '&': 'क्ष', '8': '८', '*': 'श्र', '9': '९', '(': '(', '0': '०', ')': ')', '-': '-', '_': 'ः', '=': 'ृ', '+': 'ऋ',
    q: 'ौ', Q: 'औ', w: 'ै', W: 'ऐ', e: 'ा', E: 'आ', r: 'ी', R: 'ई', t: 'ू', T: 'ऊ', y: 'ब', Y: 'भ', u: 'ह', U: 'ङ', i: 'ग', I: 'घ',
    o: 'द', O: 'ध', p: 'ज', P: 'झ', '[': 'ड', '{': 'ढ', ']': '़', '}': 'ञ', '\\': 'ॉ', '|': 'ऑ',
    a: 'ो', A: 'ओ', s: 'े', S: 'ए', d: '्', D: 'अ', f: 'ि', F: 'इ', g: 'ु', G: 'उ', h: 'प', H: 'फ', j: 'र', J: 'ऱ', k: 'क', K: 'ख',
    l: 'त', L: 'थ', ';': 'च', ':': 'छ', "'": 'ट', '"': 'ठ',
    z: 'ॆ', Z: 'ऎ', x: 'ं', X: 'ँ', c: 'म', C: 'ण', v: 'न', V: 'ऩ', b: 'व', B: 'ऴ', n: 'ल', N: 'ळ', m: 'स', M: 'श',
    ',': ',', '<': 'ष', '.': '.', '>': '।', '/': 'य', '?': 'य़',
  };
  const INS_LATIN_DIGITS = { '१': '1', '२': '2', '३': '3', '४': '4', '५': '5', '६': '6', '७': '7', '८': '8', '९': '9', '०': '0' };
  // US keyboard rows, for the on-screen keyboard
  const KB_ROWS = [
    [['`', '~'], ['1', '!'], ['2', '@'], ['3', '#'], ['4', '$'], ['5', '%'], ['6', '^'], ['7', '&'], ['8', '*'], ['9', '('], ['0', ')'], ['-', '_'], ['=', '+']],
    [['q', 'Q'], ['w', 'W'], ['e', 'E'], ['r', 'R'], ['t', 'T'], ['y', 'Y'], ['u', 'U'], ['i', 'I'], ['o', 'O'], ['p', 'P'], ['[', '{'], [']', '}'], ['\\', '|']],
    [['a', 'A'], ['s', 'S'], ['d', 'D'], ['f', 'F'], ['g', 'G'], ['h', 'H'], ['j', 'J'], ['k', 'K'], ['l', 'L'], [';', ':'], ["'", '"']],
    [['z', 'Z'], ['x', 'X'], ['c', 'C'], ['v', 'V'], ['b', 'B'], ['n', 'N'], ['m', 'M'], [',', '<'], ['.', '>'], ['/', '?']],
  ];
  // Character palette (phonetic layout's on-screen keyboard)
  const PALETTE = [
    ['अ', 'आ', 'इ', 'ई', 'उ', 'ऊ', 'ऋ', 'ए', 'ऐ', 'ओ', 'औ', 'ऑ', 'ॲ', 'अं', 'अः'],
    ['ा', 'ि', 'ी', 'ु', 'ू', 'ृ', 'े', 'ै', 'ो', 'ौ', 'ॉ', 'ॅ', 'ं', 'ँ', 'ः', '्', '़'],
    ['क', 'ख', 'ग', 'घ', 'ङ', 'च', 'छ', 'ज', 'झ', 'ञ', 'ट', 'ठ', 'ड', 'ढ', 'ण'],
    ['त', 'थ', 'द', 'ध', 'न', 'प', 'फ', 'ब', 'भ', 'म', 'य', 'र', 'ल', 'व', 'ळ'],
    ['श', 'ष', 'स', 'ह', 'क्ष', 'त्र', 'ज्ञ', 'श्र', 'र्‍', '।', '॥', 'ॐ', '₹', '०', '१', '२', '३', '४', '५', '६', '७', '८', '९'],
  ];

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
    pill.title = lang ? `Typing in ${lang === 'hi' ? 'Hindi' : 'Marathi'} (${(LAYOUTS.find(l => l.id === opts.layout) || LAYOUTS[0]).short} keyboard). Click for options (Ctrl+Alt+${lang === 'hi' ? 'H' : 'M'} to turn off)` : 'Type in Hindi or Marathi using English letters (Ctrl+Alt+H / Ctrl+Alt+M)';
  }
  function menuItems() {
    return [
      { head: 'Typing language' },
      { label: 'English (off)', checked: !lang, action: () => enable(false) },
      { label: 'हिंदी — Hindi', kbd: 'Ctrl+Alt+H', checked: lang === 'hi', action: () => enable('hi') },
      { label: 'मराठी — Marathi', kbd: 'Ctrl+Alt+M', checked: lang === 'mr', action: () => enable('mr') },
      { head: 'Keyboard' },
      ...LAYOUTS.map(l => ({ label: l.label, checked: opts.layout === l.id, action: () => setLayout(l.id) })),
      { label: 'On-screen keyboard', icon: 'keyboard', action: showKeyboard },
      { sep: true },
      opts.layout === 'inscript' ? { label: 'Devanagari digits (१२३)', checked: opts.inscriptDigits !== 'latin', action: () => setOptions({ inscriptDigits: opts.inscriptDigits === 'latin' ? 'deva' : 'latin' }) }
        : { label: 'Devanagari digits (१२३)', checked: opts.digits === 'deva', action: () => setOptions({ digits: opts.digits === 'deva' ? 'latin' : 'deva' }) },
      ...(opts.layout === 'phonetic' ? [{ label: 'Show suggestions while typing', checked: opts.suggest !== false, action: () => setOptions({ suggest: opts.suggest === false }) }] : []),
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
    if (opts.layout !== 'phonetic') {
      const lay = LAYOUTS.find(l => l.id === opts.layout);
      const txt = opts.layout === 'inscript'
        ? 'InScript is the standard Hindi and Marathi keyboard built into Windows. Each key always gives the same letter: <b>k</b> = क, <b>e</b> = ा, <b>f</b> = ि (typed after the letter), and <b>d</b> (्) joins two letters, so <b>k d k</b> = क्क. Shift gives the second letter on each key.'
        : 'Remington (GAIL) is the typewriter layout used with Kruti Dev and in government typing tests. Type exactly as you would in Kruti Dev — <b>Hkkjr</b> gives भारत, <b>f</b> before a letter gives ि — and BTS Office writes proper Unicode text that works everywhere.';
      const body = `<p style="margin-top:0;line-height:1.55;font-size:13px">${txt}</p><p style="font-size:12.5px;color:var(--ink-2)">Open the on-screen keyboard to see where every letter is. Turn typing on or off with <b>Ctrl+Alt+H</b> (Hindi) or <b>Ctrl+Alt+M</b> (Marathi).</p>`;
      if (window.BTS && BTS.dialog) BTS.dialog({ title: lay.label, body, buttons: [{ label: 'Show on-screen keyboard', value: 'kb' }, { label: 'Close', value: 1, primary: true }] }).then(v => { if (v === 'kb') showKeyboard(); });
      return;
    }
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
  function reset() { st = { el: null, buf: '', out: '', literal: false }; hideSuggest(); }
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
  const wordChar = (c, inWord) => opts.layout === 'remington' ? (c.length === 1 && c > ' ' && c !== '\u007f') : (isLetter(c) || c === '|' || (inWord && /[.~^]/.test(c)));
  // A word typed on Remington can start with a vowel sign (ि is typed before its letter). Until the letter comes,
  // show the sign on a dotted circle, as Windows does — a lone sign would merge with the space before it.
  const remington = (buf) => { const o = window.BTSLegacyFonts ? BTSLegacyFonts.toUnicode(buf, 'krutidev') : buf; return /^[\u0900-\u0903\u093A-\u094F\u0955-\u0957\u0962\u0963]/.test(o) ? '\u25CC' + o : o; };
  const render = (buf, literal) => literal ? buf : opts.layout === 'remington' ? remington(buf) : word(buf, lang);

  function onBeforeInput(e) {
    if (!lang || busy) return;
    if (e.isComposing || e.inputType === 'insertCompositionText') { reset(); return; }
    const el = editableOf(e.target); if (!el) return;
    if (e.inputType !== 'insertText' || !e.data) { reset(); return; }
    if (isField(el) && /^=/.test(el.value) && el.selectionStart > 0) return; // spreadsheet formulas stay in English
    const ch = e.data;
    if (opts.layout === 'inscript') {
      if (ch.length !== 1 || !(ch in INSCRIPT)) { reset(); return; }
      let out = INSCRIPT[ch];
      if (opts.inscriptDigits === 'latin' && INS_LATIN_DIGITS[out]) out = INS_LATIN_DIGITS[out];
      e.preventDefault(); e.stopPropagation(); replace(el, 0, out, true); return;
    }
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
    suggestSoon(el);
  }
  function onKeyDown(e) {
    if (e.ctrlKey && e.altKey && !e.shiftKey && !e.metaKey && (e.code === 'KeyH' || e.code === 'KeyM' || /^[hm]$/i.test(e.key))) {
      const want = (e.code === 'KeyH' || /^h$/i.test(e.key)) ? 'hi' : 'mr';
      e.preventDefault(); e.stopPropagation();
      enable(lang === want ? false : want);
      if (window.BTS && BTS.toast) BTS.toast(lang ? `Typing in ${lang === 'hi' ? 'Hindi (हिंदी)' : 'Marathi (मराठी)'} — ${opts.layout === 'phonetic' ? 'type in English letters' : (LAYOUTS.find(l => l.id === opts.layout) || {}).short + ' keyboard'}` : 'Typing in English');
      return;
    }
    if (!lang || !st.buf) return;
    if (e.isComposing) { reset(); return; }
    const el = editableOf(e.target);
    if (sug && sug.items.length > 1 && el && continuing(el) && !e.ctrlKey && !e.altKey && !e.metaKey) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); e.stopPropagation(); sug.sel = (sug.sel + (e.key === 'ArrowDown' ? 1 : sug.items.length - 1)) % sug.items.length; drawSuggest(); return; }
      if ((e.key === 'Enter' || e.key === 'Tab') && sug.sel > 0) { e.preventDefault(); e.stopPropagation(); pickSuggest(sug.sel); return; }
    }
    if (e.key === 'Backspace' && !e.ctrlKey && !e.altKey && !e.metaKey && el && continuing(el)) {
      e.preventDefault(); e.stopPropagation();
      const buf = st.buf.slice(0, -1); const out = buf ? render(buf, st.literal) : '';
      replace(el, st.out.length, out, false);
      if (buf) { st = { el, buf, out, literal: st.literal }; suggestSoon(el); } else reset();
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
  document.addEventListener('focusin', (e) => { reset(); const ed = editableOf(e.target); if (ed) lastField = ed; }, true);

  /* ---------------- suggestions (phonetic) ---------------- */
  let sug = null, sugEl = null, sugTimer = 0, lastField = null;
  let dictOK = null; // can the Hindi spelling dictionary rank words? (null = not checked yet)
  const isWord = (w) => { try { return window.bts && bts.isWord ? bts.isWord(w) : null; } catch { return null; } };
  function dictReady() {
    if (dictOK !== null) return dictOK;
    const real = isWord('भारत'), junk = isWord('क्ज़्ज़्फ़');
    dictOK = real === true && junk === false; return dictOK;
  }
  function variants(buf) {
    const out = new Set();
    const subs = [[/(^|[^a])a(?!a|i|u)/g, (m, p) => p + 'aa'], [/(^|[^ae])i(?!i)/g, (m, p) => p + 'ee'], [/(^|[^ao])u(?!u)/g, (m, p) => p + 'oo'],
      [/n(?=[bcdfghjklmpqrstvwxyz])/g, () => 'M'], [/n$/g, () => 'M'], [/t(?!h)/g, () => 'T'], [/d(?!h)/g, () => 'D'],
      [/s(?!h)/g, () => 'sh'], [/sh/g, () => 'Sh'], [/ri/g, () => 'R'], [/aa/g, () => 'a'], [/ee/g, () => 'i'], [/v/g, () => 'w'], [/n/g, () => 'N'], [/e$/g, () => 'ey']];
    for (const [re, fn] of subs) {
      if (!fn) continue; re.lastIndex = 0; let m;
      const r = new RegExp(re.source, 'g');
      while ((m = r.exec(buf))) { const v = buf.slice(0, m.index) + fn(m[0], m[1]) + buf.slice(m.index + m[0].length); if (v !== buf) out.add(v); if (out.size > 60) break; if (m[0] === '') r.lastIndex++; }
    }
    return [...out];
  }
  function suggestSoon(el) { clearTimeout(sugTimer); if (opts.layout !== 'phonetic' || opts.suggest === false || st.literal) { hideSuggest(); return; } sugTimer = setTimeout(() => buildSuggest(el), 60); }
  function buildSuggest(el) {
    if (!st.buf || st.el !== el || !lang) { hideSuggest(); return; }
    const base = st.out; const items = [base];
    const lexHit = Object.values(lexOf[lang]);
    if (st.buf.length >= 2) {
      const cands = [...new Set(variants(st.buf).map(v => word(v, lang)))].filter(w => w && w !== base);
      const useDict = dictReady();
      const good = cands.filter(w => lexHit.includes(w) || (useDict && isWord(w) === true));
      good.slice(0, 4).forEach(w => items.push(w));
    }
    items.push(st.buf); // the English letters, as typed
    sug = { el, items: [...new Set(items)], sel: 0 }; drawSuggest();
  }
  function caretRect(el) {
    if (isField(el)) { const r = el.getBoundingClientRect(); return { left: r.left + 6, bottom: Math.min(r.bottom, r.top + 28) }; }
    const s = getSelection(); if (s.rangeCount) { const rr = s.getRangeAt(0).getBoundingClientRect(); if (rr.width || rr.height) return { left: rr.left, bottom: rr.bottom }; }
    const r = el.getBoundingClientRect(); return { left: r.left, bottom: r.top + 24 };
  }
  function drawSuggest() {
    if (!sug) return;
    injectCSS();
    if (!sugEl) { sugEl = document.createElement('div'); sugEl.className = 'indic-sug'; sugEl.setAttribute('data-indic', 'off'); sugEl.addEventListener('mousedown', e => { e.preventDefault(); const b = e.target.closest('[data-i]'); if (b) pickSuggest(+b.dataset.i); }); document.body.appendChild(sugEl); }
    sugEl.innerHTML = sug.items.map((w, i) => `<button data-i="${i}" class="${i === sug.sel ? 'sel' : ''}${i === sug.items.length - 1 ? ' en' : ''}"><small>${i + 1}</small>${w.replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]))}</button>`).join('') + '<i>↑↓ choose · Enter pick · Space keep first</i>';
    const r = caretRect(sug.el); sugEl.style.display = 'flex';
    const w = sugEl.offsetWidth, h = sugEl.offsetHeight;
    sugEl.style.left = Math.max(6, Math.min(r.left, innerWidth - w - 8)) + 'px';
    sugEl.style.top = (r.bottom + 6 + h > innerHeight ? r.bottom - h - 30 : r.bottom + 6) + 'px';
  }
  function hideSuggest() { sug = null; clearTimeout(sugTimer); if (sugEl) sugEl.style.display = 'none'; }
  function pickSuggest(i) {
    if (!sug || !st.buf) return; const w = sug.items[i]; const el = sug.el; if (w == null || !continuing(el)) { hideSuggest(); return; }
    replace(el, st.out.length, w, false); reset();
  }

  /* ---------------- on-screen keyboard ---------------- */
  let kbEl = null, kbShift = false;
  function keyOut(ch) {
    if (opts.layout === 'inscript') { const o = INSCRIPT[ch]; return o && opts.inscriptDigits === 'latin' && INS_LATIN_DIGITS[o] ? INS_LATIN_DIGITS[o] : o; }
    if (opts.layout === 'remington') return window.BTSLegacyFonts ? BTSLegacyFonts.toUnicode(ch, 'krutidev') : ch;
    return null;
  }
  function insertText(t) {
    const el = (document.activeElement && editableOf(document.activeElement)) || lastField;
    if (!el) { if (window.BTS && BTS.toast) BTS.toast('Click where you want the text to go first'); return; }
    if (document.activeElement !== el) el.focus();
    if (opts.layout === 'remington' && lang) { // feed the key through the typing engine so letters combine
      const ev = new InputEvent('beforeinput', { inputType: 'insertText', data: t, bubbles: true, cancelable: true });
      if (el.dispatchEvent(ev)) document.execCommand('insertText', false, t); return;
    }
    reset(); replace(el, 0, t, true);
  }
  function drawKeyboard() {
    if (!kbEl) return;
    const lay = LAYOUTS.find(l => l.id === opts.layout) || LAYOUTS[0];
    let body = '';
    if (opts.layout === 'phonetic') body = PALETTE.map(r => `<div class="r">${r.map(c => `<button class="k ch" data-t="${c}">${c}</button>`).join('')}</div>`).join('');
    else body = KB_ROWS.map((r, ri) => `<div class="r r${ri}">${r.map(([n, sh]) => { const key = kbShift ? sh : n; const o = keyOut(key) || ''; return `<button class="k" data-key="${key.replace(/"/g, '&quot;')}"><em>${(kbShift ? sh : n).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]))}</em><b>${o.replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]))}</b></button>`; }).join('')}</div>`).join('')
      + `<div class="r"><button class="k wide ${kbShift ? 'on' : ''}" data-act="shift">⇧ Shift</button><button class="k space" data-t=" ">Space</button><button class="k wide" data-act="bksp">⌫</button></div>`;
    kbEl.innerHTML = `<div class="kb-h"><b>${lang ? NAMES[lang] : 'हिंदी / मराठी'} · ${lay.short}</b><select data-act="layout">${LAYOUTS.map(l => `<option value="${l.id}" ${l.id === opts.layout ? 'selected' : ''}>${l.short}</option>`).join('')}</select><button data-act="close" title="Close">✕</button></div><div class="kb-b">${body}</div>`;
  }
  function showKeyboard() {
    injectCSS();
    if (!kbEl) {
      kbEl = document.createElement('div'); kbEl.className = 'indic-kb'; kbEl.setAttribute('data-indic', 'off');
      kbEl.addEventListener('mousedown', e => { if (!e.target.closest('select')) e.preventDefault(); });
      kbEl.addEventListener('click', e => {
        const b = e.target.closest('button'); if (!b) return;
        const act = b.dataset.act;
        if (act === 'close') return hideKeyboard();
        if (act === 'shift') { kbShift = !kbShift; return drawKeyboard(); }
        if (act === 'bksp') { const el = (document.activeElement && editableOf(document.activeElement)) || lastField; if (el) { el.focus(); reset(); document.execCommand('delete'); } return; }
        if (b.dataset.t != null) { insertText(b.dataset.t); return; }
        if (b.dataset.key != null) { if (opts.layout === 'remington') insertText(b.dataset.key); else insertText(keyOut(b.dataset.key) || ''); if (kbShift) { kbShift = false; drawKeyboard(); } }
      });
      kbEl.addEventListener('change', e => { if (e.target.dataset.act === 'layout') setLayout(e.target.value); });
      // drag by the title row
      kbEl.addEventListener('mousedown', e => {
        if (!e.target.closest('.kb-h') || e.target.closest('select,button')) return;
        const r = kbEl.getBoundingClientRect(), dx = e.clientX - r.left, dy = e.clientY - r.top;
        const mv = (ev) => { kbEl.style.left = Math.max(0, Math.min(innerWidth - r.width, ev.clientX - dx)) + 'px'; kbEl.style.top = Math.max(0, Math.min(innerHeight - r.height, ev.clientY - dy)) + 'px'; kbEl.style.right = 'auto'; kbEl.style.bottom = 'auto'; };
        const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); };
        document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
      });
      document.body.appendChild(kbEl);
    }
    kbEl.style.display = 'block'; drawKeyboard();
    if (!lang) enable(lastLang);
  }
  function hideKeyboard() { if (kbEl) kbEl.style.display = 'none'; }
  function setLayout(id) {
    if (!LAYOUTS.some(l => l.id === id)) return;
    setOptions({ layout: id }); reset();
    if (kbEl && kbEl.style.display !== 'none') drawKeyboard();
    if (window.BTS && BTS.toast && lang) BTS.toast(`Keyboard: ${LAYOUTS.find(l => l.id === id).label}`);
  }
  function injectCSS() {
    if (document.getElementById('indic-css2')) return;
    const st2 = document.createElement('style'); st2.id = 'indic-css2';
    st2.textContent = `.indic-sug{position:fixed;z-index:9998;display:none;align-items:center;gap:2px;flex-wrap:wrap;max-width:560px;background:var(--surface,#fff);color:var(--ink,#1f2533);border:1px solid var(--line,#d5d9e2);border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,.18);padding:3px;font-family:'Noto Sans Devanagari','Mukta',sans-serif}
.indic-sug button{border:0;background:none;color:inherit;font:inherit;font-size:16px;padding:3px 9px;border-radius:5px;cursor:pointer;display:flex;align-items:baseline;gap:4px}.indic-sug button small{font:11px system-ui,sans-serif;color:var(--ink-3,#7a8399)}
.indic-sug button.sel{background:var(--accent-soft,#e8ebff);color:var(--accent,#3C4FE0)}html[data-theme="dark"] .indic-sug button.sel{color:color-mix(in srgb,var(--accent,#3C4FE0) 40%,#fff)}.indic-sug button.en{font:13px system-ui,sans-serif;color:var(--ink-2,#47526a)}.indic-sug button:hover{background:var(--surface-2,#f3f5f9)}
.indic-sug i{font:11px system-ui,sans-serif;font-style:normal;color:var(--ink-3,#7a8399);padding:0 8px;width:100%}
.indic-kb{position:fixed;z-index:9997;right:18px;bottom:44px;display:none;background:var(--surface,#fff);color:var(--ink,#1f2533);border:1px solid var(--line,#d5d9e2);border-radius:12px;box-shadow:0 12px 36px rgba(0,0,0,.22);padding:8px 10px 10px;user-select:none}
.indic-kb .kb-h{cursor:move;display:flex;align-items:center;gap:8px;margin-bottom:6px;font-size:13px}.indic-kb .kb-h b{font-family:'Noto Sans Devanagari','Mukta',system-ui,sans-serif;flex:1}.indic-kb .kb-h select{font:inherit;font-size:12px;border:1px solid var(--line,#d5d9e2);border-radius:6px;background:var(--surface,#fff);color:inherit;padding:2px 4px}
.indic-kb .kb-h button{border:0;background:none;color:var(--ink-2,#47526a);font-size:14px;cursor:pointer;padding:2px 6px;border-radius:5px}.indic-kb .kb-h button:hover{background:var(--surface-2,#f3f5f9)}
.indic-kb .r{display:flex;gap:4px;margin-top:4px}.indic-kb .r1{padding-left:14px}.indic-kb .r2{padding-left:24px}.indic-kb .r3{padding-left:40px}
.indic-kb .k{position:relative;min-width:38px;height:40px;border:1px solid var(--line,#d5d9e2);border-radius:7px;background:var(--surface-2,#f6f7fa);color:inherit;cursor:pointer;font-family:'Noto Sans Devanagari','Mukta',sans-serif;font-size:17px;display:grid;place-items:center;padding:0 4px}
.indic-kb .k:hover{border-color:var(--accent,#3C4FE0);background:var(--accent-soft,#e8ebff)}.indic-kb .k:active{transform:translateY(1px)}
.indic-kb .k em{position:absolute;left:4px;top:2px;font:10px system-ui,sans-serif;font-style:normal;color:var(--ink-3,#7a8399)}.indic-kb .k b{font-weight:500}
.indic-kb .k.ch{min-width:32px;height:34px;font-size:16px}.indic-kb .k.wide{min-width:84px;font:12px system-ui,sans-serif}.indic-kb .k.wide.on{background:var(--accent-soft,#e8ebff);border-color:var(--accent,#3C4FE0)}.indic-kb .k.space{flex:1;font:12px system-ui,sans-serif}`;
    document.head.appendChild(st2);
  }

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
    layouts: LAYOUTS.map(l => ({ id: l.id, label: l.label, desc: l.desc })), get layout() { return opts.layout; }, setLayout, showKeyboard, hideKeyboard, inscript: INSCRIPT,
    toggle: (l) => enable(lang ? false : (normLang(l) || lastLang)),
    skip: '[data-indic="off"], #ribbon input, #ribbon select, input[type="number"], #gkeys',
  };
})();
