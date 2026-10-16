/* BTS Docs 1.2 — References: sources, citations & bibliography (APA 7, MLA 9, IEEE, Harvard, Chicago author-date),
 * captions with live numbers, table of figures, and Update fields. Sources live in D.design.biblio so they are
 * saved with the document (and in the .docx custom XML part). */
(function () {
  const D = window.Docs; const { $, $$ } = D; const esc = BTS.esc;

  D.CITE_STYLES = { apa: 'APA 7th', mla: 'MLA 9th', ieee: 'IEEE', harvard: 'Harvard', chicago: 'Chicago (author-date)' };
  D.SOURCE_TYPES = { book: 'Book', journal: 'Journal article', website: 'Web page', report: 'Report', thesis: 'Thesis / dissertation' };
  const BIB_TITLE = { apa: 'References', mla: 'Works Cited', ieee: 'References', harvard: 'Reference list', chicago: 'References' };
  const bib = () => { if (!D.design.biblio) D.design.biblio = { style: 'apa', sources: [] }; if (!D.design.biblio.sources) D.design.biblio.sources = []; return D.design.biblio; };
  D.biblio = bib;
  D.sources = () => bib().sources;
  D.sourceById = (id) => bib().sources.find(s => s.id === id);

  /* ---------- names ---------- */
  const ORG = /\b(university|organi[sz]ation|ministry|institute|association|department|inc|ltd|limited|society|council|board|government|committee|bank|foundation|agency|press|group|corporation|company|news|bureau|commission|trust|centre|center|school|college|office|authority|india|nations|who|unesco|unicef|isro|nasa)\b/i;
  D.parseAuthors = (s) => String(s || '').split(/\s*[;\n]\s*/).map(x => x.trim()).filter(Boolean).map(x => {
    if (/^\{.*\}$/.test(x)) return { org: x.slice(1, -1) };
    if (x.includes(',')) { const [last, ...rest] = x.split(','); return { last: last.trim(), first: rest.join(',').trim() }; }
    const words = x.split(/\s+/);
    if (words.length === 1) return { last: x, first: '' };
    if (words.length <= 4 && !ORG.test(x) && words.every(w => /^[\p{Lu}ऀ-ॿ]/u.test(w) || /^(de|van|von|da|di|del|la|le|bin|al)$/i.test(w))) return { last: words[words.length - 1], first: words.slice(0, -1).join(' ') };
    return { org: x };
  });
  const initials = (first, spaced = true) => String(first || '').split(/\s+/).filter(Boolean).map(w => w.split('-').map(p => p ? p[0].toUpperCase() + '.' : '').join('-')).join(spaced ? ' ' : '');
  const lastOf = (a) => a.org || a.last;
  const listJoin = (arr, sep, lastSep) => arr.length <= 1 ? (arr[0] || '') : arr.length === 2 ? arr[0] + lastSep + arr[1] : arr.slice(0, -1).join(sep) + lastSep + arr[arr.length - 1];
  const doiOf = (s) => String(s.doi || '').replace(/^https?:\/\/(dx\.)?doi\.org\//i, '').trim();
  const pageLoc = (p, style) => { if (!p) return ''; const multi = /[-–,]/.test(p); if (style === 'apa' || style === 'harvard') return (multi ? 'pp. ' : 'p. ') + p; if (style === 'ieee') return (multi ? 'pp. ' : 'p. ') + p; return p; };
  const tidy = (h) => h.replace(/\s+([.,;:])/g, '$1').replace(/([.?!])\.(?!\.)/g, '$1').replace(/,\s*,/g, ',').replace(/,\s*\./g, '.').replace(/\(\s*\)/g, '').replace(/\[\s*\]/g, '').replace(/\s{2,}/g, ' ').replace(/:\s*\./g, '.').trim();

  /* ---------- in-text ---------- */
  function authorText(s, style) {
    const a = D.parseAuthors(s.authors); if (!a.length) return style === 'mla' ? '“' + esc(shortTitle(s.title)) + '”' : '<i>' + esc(shortTitle(s.title)) + '</i>';
    const L = a.map(x => esc(lastOf(x)));
    if (style === 'apa') return a.length === 1 ? L[0] : a.length === 2 ? L[0] + ' &amp; ' + L[1] : L[0] + ' et al.';
    if (style === 'mla' || style === 'harvard') return a.length === 1 ? L[0] : a.length === 2 ? L[0] + ' and ' + L[1] : (style === 'harvard' && a.length === 3 ? L[0] + ', ' + L[1] + ' and ' + L[2] : L[0] + ' et al.');
    if (style === 'chicago') return a.length === 1 ? L[0] : a.length === 2 ? L[0] + ' and ' + L[1] : a.length === 3 ? L[0] + ', ' + L[1] + ', and ' + L[2] : L[0] + ' et al.';
    return L[0];
  }
  const shortTitle = (t) => { const w = String(t || 'Untitled').split(/\s+/); return w.length > 4 ? w.slice(0, 4).join(' ') + '…' : w.join(' '); };
  const year = (s) => esc(String(s.year || '').trim() || 'n.d.');
  /** In-text citation HTML for the given source ids. */
  D.formatCitation = (ids, page, style = bib().style, numbers) => {
    const srcs = ids.map(D.sourceById).filter(Boolean);
    if (!srcs.length) return '(Source missing)';
    if (style === 'ieee') { const n = srcs.map(s => (numbers && numbers.get(s.id)) || '?'); if (srcs.length === 1) return '[' + n[0] + (page ? ', ' + esc(pageLoc(page, 'ieee')) : '') + ']'; return n.map(x => '[' + x + ']').join(', '); }
    const one = (s, withPage) => {
      const au = authorText(s, style); const pg = withPage && page ? page : '';
      if (style === 'apa' || style === 'harvard') return au + ', ' + year(s) + (pg ? ', ' + esc(pageLoc(pg, style)) : '');
      if (style === 'mla') return au + (pg ? ' ' + esc(pg) : '');
      if (style === 'chicago') return au + ' ' + year(s) + (pg ? ', ' + esc(pg) : '');
      return au;
    };
    return '(' + srcs.map((s, i) => one(s, i === srcs.length - 1)).join('; ') + ')';
  };

  /* ---------- reference list entries ---------- */
  function names(s, style) {
    const a = D.parseAuthors(s.authors); if (!a.length) return '';
    if (style === 'apa') { const n = a.map(x => x.org ? esc(x.org) : esc(x.last) + (x.first ? ', ' + esc(initials(x.first)) : '')); return n.length === 1 ? n[0] : n.length === 2 ? n[0] + ', &amp; ' + n[1] : n.slice(0, -1).join(', ') + ', &amp; ' + n[n.length - 1]; }
    if (style === 'mla') { const f = (x, inv) => x.org ? esc(x.org) : inv ? esc(x.last) + (x.first ? ', ' + esc(x.first) : '') : esc((x.first ? x.first + ' ' : '') + x.last); return a.length === 1 ? f(a[0], true) : a.length === 2 ? f(a[0], true) + ', and ' + f(a[1]) : f(a[0], true) + ', et al'; }
    if (style === 'ieee') { const n = a.map(x => x.org ? esc(x.org) : esc((x.first ? initials(x.first) + ' ' : '') + x.last)); if (n.length > 6) return n[0] + ' <i>et al.</i>'; return listJoin(n, ', ', n.length > 2 ? ', and ' : ' and '); }
    if (style === 'harvard') { const n = a.map(x => x.org ? esc(x.org) : esc(x.last) + (x.first ? ', ' + esc(initials(x.first, false)) : '')); if (n.length > 3) return n[0] + ' <i>et al.</i>'; return listJoin(n, ', ', ' and '); }
    if (style === 'chicago') { const n = a.map((x, i) => x.org ? esc(x.org) : i === 0 ? esc(x.last) + (x.first ? ', ' + esc(x.first) : '') : esc((x.first ? x.first + ' ' : '') + x.last)); if (n.length > 10) return n.slice(0, 7).join(', ') + ', et al'; return listJoin(n, ', ', n.length > 2 ? ', and ' : ', and '); }
    return '';
  }
  D.formatEntry = (s, style = bib().style, num) => {
    const e = (k) => esc(String(s[k] || '').trim()); const T = e('title') || 'Untitled'; const N = names(s, style); const Y = year(s);
    const doi = doiOf(s); const url = e('url'); const link = doi ? 'https://doi.org/' + esc(doi) : url;
    const t = s.type || 'book'; let h = '';
    if (style === 'apa') {
      const head = N ? `${N} (${Y}).` : '';
      if (t === 'journal') h = `${head} ${T}. <i>${e('journal')}</i>${s.volume ? `, <i>${e('volume')}</i>` : ''}${s.issue ? `(${e('issue')})` : ''}${s.pages ? `, ${e('pages')}` : ''}. ${link}`;
      else if (t === 'website') h = N ? `${head} <i>${T}</i>. ${e('journal')}. ${url}` : `<i>${T}</i>. (${Y}). ${e('journal')}. ${url}`;
      else if (t === 'report') h = `${head} <i>${T}</i>${s.issue ? ` (Report No. ${e('issue')})` : ''}. ${e('publisher')}. ${link}`;
      else if (t === 'thesis') h = `${head} <i>${T}</i> [${e('kind') || 'Doctoral dissertation'}, ${e('publisher')}]. ${url}`;
      else h = `${head} <i>${T}</i>${s.edition ? ` (${e('edition')} ed.)` : ''}. ${e('publisher')}. ${link}`;
      if (!N && t !== 'website') h = `<i>${T}</i>. (${Y}). ` + h.replace(/^\s*<i>[^<]*<\/i>/, '').replace(/^\s*\.?\s*/, '');
    } else if (style === 'mla') {
      const head = N ? N + '.' : '';
      if (t === 'journal') h = `${head} “${T}.” <i>${e('journal')}</i>${s.volume ? `, vol. ${e('volume')}` : ''}${s.issue ? `, no. ${e('issue')}` : ''}, ${Y}${s.pages ? `, pp. ${e('pages')}` : ''}.${doi ? ` https://doi.org/${esc(doi)}.` : ''}`;
      else if (t === 'website') h = `${head} “${T}.” <i>${e('journal')}</i>, ${Y}, ${url.replace(/^https?:\/\//, '')}.${s.accessed ? ` Accessed ${e('accessed')}.` : ''}`;
      else if (t === 'thesis') h = `${head} <i>${T}</i>. ${Y}. ${e('publisher')}, ${e('kind') || 'PhD dissertation'}.`;
      else h = `${head} <i>${T}</i>. ${s.edition ? e('edition') + ' ed., ' : ''}${e('publisher')}, ${Y}.`;
    } else if (style === 'ieee') {
      if (t === 'journal') h = `${N}, “${T},” <i>${e('journal')}</i>${s.volume ? `, vol. ${e('volume')}` : ''}${s.issue ? `, no. ${e('issue')}` : ''}${s.pages ? `, pp. ${e('pages')}` : ''}, ${Y}${doi ? `, doi: ${esc(doi)}` : ''}.`;
      else if (t === 'website') h = `${N ? N + ', ' : ''}“${T},” ${e('journal')}. [Online]. Available: ${url}${s.accessed ? ` (accessed ${e('accessed')})` : ''}.`;
      else if (t === 'report') h = `${N}, “${T},” ${e('publisher')}${s.city ? ', ' + e('city') : ''}${s.issue ? `, Rep. ${e('issue')}` : ''}, ${Y}.`;
      else if (t === 'thesis') h = `${N}, “${T},” ${/master/i.test(s.kind || '') ? 'M.S. thesis' : 'Ph.D. dissertation'}, ${e('publisher')}${s.city ? ', ' + e('city') : ''}, ${Y}.`;
      else h = `${N}, <i>${T}</i>${s.edition ? `, ${e('edition')} ed` : ''}. ${s.city ? e('city') + ': ' : ''}${e('publisher')}, ${Y}.`;
    } else if (style === 'harvard') {
      const head = N ? `${N} (${Y})` : `<i>${T}</i> (${Y})`;
      if (t === 'journal') h = `${head} ‘${T}’, <i>${e('journal')}</i>, ${e('volume')}${s.issue ? `(${e('issue')})` : ''}${s.pages ? `, pp. ${e('pages')}` : ''}.${doi ? ` Available at: https://doi.org/${esc(doi)}.` : ''}`;
      else if (t === 'website') h = `${head} ${N ? `<i>${T}</i>. ` : ''}Available at: ${url}${s.accessed ? ` (Accessed: ${e('accessed')})` : ''}.`;
      else if (t === 'thesis') h = `${head} ${N ? `<i>${T}</i>. ` : ''}${e('kind') || 'PhD thesis'}. ${e('publisher')}.`;
      else h = `${head} ${N ? `<i>${T}</i>. ` : ''}${s.edition ? e('edition') + ' edn. ' : ''}${s.city ? e('city') + ': ' : ''}${e('publisher')}.`;
    } else { // chicago author-date
      const head = N ? `${N}. ${Y}.` : `<i>${T}</i>. ${Y}.`;
      if (t === 'journal') h = `${head} “${T}.” <i>${e('journal')}</i> ${e('volume')}${s.issue ? ` (${e('issue')})` : ''}${s.pages ? `: ${e('pages')}` : ''}.${doi ? ` https://doi.org/${esc(doi)}.` : (url ? ' ' + url + '.' : '')}`;
      else if (t === 'website') h = `${head} “${T}.” ${e('journal')}. ${s.accessed ? `Accessed ${e('accessed')}. ` : ''}${url}.`;
      else if (t === 'thesis') h = `${head} “${T}.” ${e('kind') || 'PhD diss.'}, ${e('publisher')}.`;
      else h = `${head} ${N ? `<i>${T}</i>. ` : ''}${s.city ? e('city') + ': ' : ''}${e('publisher')}.`;
    }
    return tidy(h);
  };
  const sortKey = (s) => { const a = D.parseAuthors(s.authors)[0]; return ((a ? lastOf(a) : s.title) || '').toLowerCase() + ' ' + (s.year || ''); };

  /* ---------- numbering / refresh ---------- */
  function ieeeNumbers() {
    const m = new Map(); let n = 0;
    $$('.cite', D.doc).forEach(c => (c.dataset.src || '').split(/[;,]/).filter(Boolean).forEach(id => { if (!m.has(id) && D.sourceById(id)) m.set(id, ++n); }));
    bib().sources.forEach(s => { if (!m.has(s.id)) m.set(s.id, ++n); });
    return m;
  }
  D.updateCitations = () => {
    const style = bib().style; const nums = style === 'ieee' ? ieeeNumbers() : null;
    $$('.cite', D.doc).forEach(c => { const ids = (c.dataset.src || '').split(/[;,]/).filter(Boolean); const h = D.formatCitation(ids, c.dataset.page || '', style, nums); if (c.innerHTML !== h) c.innerHTML = h; });
    $$('.biblio', D.doc).forEach(b => renderBiblio(b, style, nums));
  };
  function renderBiblio(b, style, nums) {
    const list = [...bib().sources];
    if (style === 'ieee') list.sort((a, c) => nums.get(a.id) - nums.get(c.id)); else list.sort((a, c) => sortKey(a).localeCompare(sortKey(c)));
    const title = b.dataset.title || BIB_TITLE[style];
    const tag = /^h[1-3]$/.test(b.dataset.head || '') ? b.dataset.head : 'h1';
    b.innerHTML = `<${tag} class="bib-title">${esc(BIB_TITLE[style] && !b.dataset.title ? BIB_TITLE[style] : title)}</${tag}>` + (list.length ? list.map(s => `<p class="bib-entry${style === 'ieee' ? ' ieee' : ''}" data-src="${esc(s.id)}">${style === 'ieee' ? `<span class="bib-n">[${nums.get(s.id)}]</span><span class="tab">\t</span>` : ''}${D.formatEntry(s, style)}</p>`).join('') : '<p class="bib-entry bib-empty">There are no sources in this document. Use References ▸ Manage sources to add some.</p>');
    b.dataset.style = style;
  }

  /* ---------- look up a source online (DOI → Crossref, ISBN → Open Library, then Google Books) ---------- */
  const strip = (h) => { const d = document.createElement('div'); d.innerHTML = String(h || '').replace(/<[^>]+>/g, ''); return d.textContent.replace(/\s+/g, ' ').trim(); };
  const lastFirst = (name) => { name = String(name || '').trim(); if (!name || name.includes(',')) return name; const w = name.split(/\s+/); return w.length < 2 ? name : w[w.length - 1] + ', ' + w.slice(0, -1).join(' '); };
  const ordinal = (n) => { n = String(n || ''); if (!/^\d+$/.test(n)) return n; const t = +n % 100, u = +n % 10; return n + (t >= 11 && t <= 13 ? 'th' : u === 1 ? 'st' : u === 2 ? 'nd' : u === 3 ? 'rd' : 'th'); };
  /** What the user typed: { doi } | { isbn } | { bad: message }. */
  D.parseLookup = (q) => {
    q = String(q || '').trim();
    const doi = q.match(/10\.\d{4,9}\/[^\s"<>]+/i);
    if (doi) { let d = doi[0]; try { d = decodeURIComponent(d); } catch { } return { doi: d.replace(/[.,;)\]]+$/, '') }; }
    const n = q.replace(/^isbn(-1[03])?[:\s]*/i, '').replace(/[\s-]/g, '').toUpperCase();
    if (/^\d{9}[\dX]$/.test(n)) { let sum = 0; for (let i = 0; i < 10; i++) sum += (n[i] === 'X' ? 10 : +n[i]) * (10 - i); return sum % 11 === 0 ? { isbn: n } : { bad: 'That ISBN has a typing mistake — its last digit doesn’t match. Check the number.' }; }
    if (/^97[89]\d{10}$/.test(n)) { let sum = 0; for (let i = 0; i < 13; i++) sum += +n[i] * (i % 2 ? 3 : 1); return sum % 10 === 0 ? { isbn: n } : { bad: 'That ISBN has a typing mistake — its last digit doesn’t match. Check the number.' }; }
    return { bad: 'That doesn’t look like a DOI or an ISBN. A DOI starts with 10. (for example 10.1038/nature12373); an ISBN has 10 or 13 digits.' };
  };
  async function getJSON(url) {
    if (D.lookupTransport) { const r = await D.lookupTransport(url); if (r && r.ok) return r.data; const e = new Error((r && r.error) || 'Lookup failed'); e.code = /No internet/i.test(e.message) ? 'offline' : /Not found/i.test(e.message) ? 'notfound' : 'other'; throw e; } // tests
    if (!(window.bts && bts.ext)) { const e = new Error('Looking up sources works in the BTS Office desktop app. You can still type the details yourself.'); e.code = 'nodesk'; throw e; }
    let r; try { r = await bts.ext('fetch-json', url); } catch { const e = new Error('Looking up sources isn’t available right now. You can still type the details yourself.'); e.code = 'nodesk'; throw e; }
    if (r && r.ok) return r.data;
    const e = new Error((r && r.error) || 'Lookup failed'); e.code = /No internet/i.test(e.message) ? 'offline' : /Not found/i.test(e.message) ? 'notfound' : 'other'; throw e;
  }
  const OFFLINE = 'You’re offline. Connect to the internet to look up sources, or type the details yourself.';
  /** Source fields for a DOI / ISBN / link with a DOI. Throws an Error with a friendly message. */
  D.lookupSource = async (q) => {
    const p = D.parseLookup(q); if (p.bad) throw new Error(p.bad);
    try {
      if (p.doi) {
        const m = (await getJSON('https://api.crossref.org/works/' + encodeURIComponent(p.doi))).message || {};
        const dp = (m.issued || m['published-print'] || m['published-online'] || m.created || {})['date-parts'];
        const year = dp && dp[0] && dp[0][0] ? String(dp[0][0]) : '';
        const authors = (m.author || m.editor || []).map(a => a.family ? a.family + (a.given ? ', ' + a.given : '') : (a.name || '')).filter(Boolean).join('; ');
        const t = m.type || '';
        const type = /journal-article|proceedings-article|posted-content|peer-review|article/.test(t) ? 'journal' : /report/.test(t) ? 'report' : /dissertation/.test(t) ? 'thesis' : 'book';
        const title = strip((m.title || [])[0]) + (m.subtitle && m.subtitle[0] ? ': ' + strip(m.subtitle[0]) : '');
        const f = { type, authors, title, year, doi: m.DOI || p.doi };
        if (type === 'journal') Object.assign(f, { journal: strip((m['container-title'] || [])[0] || (m['short-container-title'] || [])[0] || ''), volume: m.volume || '', issue: m.issue || '', pages: String(m.page || '').replace(/-/g, '–'), publisher: m.publisher || '' });
        else if (type === 'thesis') Object.assign(f, { publisher: (m.institution && m.institution[0] && m.institution[0].name) || m.publisher || '', kind: 'Doctoral dissertation' });
        else Object.assign(f, { publisher: m.publisher || '', city: m['publisher-location'] || '', edition: m['edition-number'] ? ordinal(m['edition-number']) : '' });
        return f;
      }
      // ISBN: Open Library first, Google Books as a fallback
      try {
        const b = await getJSON('https://openlibrary.org/isbn/' + p.isbn + '.json');
        const names = [];
        for (const a of (b.authors || []).slice(0, 8)) { try { const au = await getJSON('https://openlibrary.org' + a.key + '.json'); if (au && (au.personal_name || au.name)) names.push(lastFirst(au.personal_name || au.name)); } catch (e) { if (e.code === 'offline') throw e; } }
        const year = ((b.publish_date || '').match(/\d{4}/) || [''])[0];
        return { type: 'book', authors: names.join('; ') || lastFirst(String(b.by_statement || '').replace(/^by\s+/i, '').replace(/\.$/, '')), title: (b.title || '') + (b.subtitle ? ': ' + b.subtitle : ''), year, publisher: (b.publishers || [])[0] || '', city: (b.publish_places || [])[0] || '', edition: String(b.edition_name || '').replace(/[[\]]/g, '').replace(/\.\s*\.$/, '.').trim(), url: b.key ? 'https://openlibrary.org' + b.key : '' };
      } catch (e) {
        if (e.code === 'offline' || e.code === 'nodesk') throw e;
        const g = await getJSON('https://www.googleapis.com/books/v1/volumes?q=isbn:' + p.isbn);
        const v = g && g.items && g.items[0] && g.items[0].volumeInfo; if (!v) { const nf = new Error('Not found'); nf.code = 'notfound'; throw nf; }
        return { type: 'book', authors: (v.authors || []).map(lastFirst).join('; '), title: (v.title || '') + (v.subtitle ? ': ' + v.subtitle : ''), year: ((v.publishedDate || '').match(/\d{4}/) || [''])[0], publisher: v.publisher || '', url: v.infoLink || '' };
      }
    } catch (e) {
      if (e.code === 'offline') throw new Error(OFFLINE);
      if (e.code === 'nodesk') throw e;
      if (e.code === 'notfound') throw new Error(p.doi ? 'No article or book was found for that DOI. Check it and try again.' : 'No book was found for that ISBN. Check the number, or type the details yourself.');
      throw new Error('The lookup didn’t work (' + e.message + '). Try again, or type the details yourself.');
    }
  };

  /* ---------- manage sources dialog ---------- */
  const FIELDS = {
    book: [['authors', 'Author(s)'], ['title', 'Title'], ['year', 'Year'], ['edition', 'Edition (e.g. 2nd)'], ['publisher', 'Publisher'], ['city', 'City'], ['doi', 'DOI'], ['url', 'URL']],
    journal: [['authors', 'Author(s)'], ['title', 'Article title'], ['journal', 'Journal'], ['year', 'Year'], ['volume', 'Volume'], ['issue', 'Issue'], ['pages', 'Pages (e.g. 12–19)'], ['doi', 'DOI'], ['url', 'URL']],
    website: [['authors', 'Author(s) or organisation'], ['title', 'Page title'], ['journal', 'Website name'], ['year', 'Year published'], ['url', 'URL'], ['accessed', 'Date accessed (e.g. 14 October 2026)']],
    report: [['authors', 'Author(s) or organisation'], ['title', 'Report title'], ['year', 'Year'], ['issue', 'Report number'], ['publisher', 'Publisher / institution'], ['city', 'City'], ['doi', 'DOI'], ['url', 'URL']],
    thesis: [['authors', 'Author'], ['title', 'Title'], ['year', 'Year'], ['kind', 'Type (e.g. Master’s thesis, Doctoral dissertation)'], ['publisher', 'University'], ['city', 'City'], ['url', 'URL']],
  };
  D.manageSources = async (opts = {}) => {
    const sv = D.saveSel();
    const work = JSON.parse(JSON.stringify(bib().sources)); let cur = opts.newSource ? null : (work[0] ? work[0].id : null);
    const body = document.createElement('div'); body.className = 'src-dlg modal-noenter';
    const render = () => {
      const s = work.find(x => x.id === cur);
      body.innerHTML = `<div class="src-side"><div class="src-search">${icon('search', 14)}<input id="src-q" placeholder="Search sources" spellcheck="false"></div><div class="src-list">${work.length ? work.map(x => `<button class="src-item${x.id === cur ? ' on' : ''}" data-id="${esc(x.id)}"><b>${esc(D.parseAuthors(x.authors).map(lastOf).slice(0, 2).join(', ') || x.title || 'Untitled')}${x.year ? ' (' + esc(x.year) + ')' : ''}</b><small>${esc(D.SOURCE_TYPES[x.type] || '')} · ${esc(x.title || '')}</small></button>`).join('') : '<div class="empty-note" style="padding:10px">No sources yet.</div>'}</div><div class="src-side-btns"><button class="btn primary" data-a="new">${icon('plus', 15)} New source</button>${s ? `<button class="btn danger" data-a="del" title="Delete this source">${icon('trash-2', 15)}</button>` : ''}</div></div>
        <div class="src-main"><div class="src-lookup"><div class="src-search">${icon('search', 14)}<input id="src-lookup" placeholder="Look up by DOI, ISBN or a link with a DOI" spellcheck="false" data-indic="off"></div><button class="btn" data-a="lookup" title="Fill in the details from Crossref (articles) or Open Library / Google Books (books)">${icon('globe', 15)} Look up</button><span class="src-net" title="This is the only part of BTS Docs that uses the internet">Needs internet</span></div><div class="src-lookup-msg" id="src-lookup-msg"></div>${s ? `<div class="field"><label>Type of source</label><select id="src-type">${Object.entries(D.SOURCE_TYPES).map(([k, v]) => `<option value="${k}" ${s.type === k ? 'selected' : ''}>${v}</option>`).join('')}</select></div>
          <div class="src-fields">${FIELDS[s.type || 'book'].map(([k, l]) => `<div class="field${k === 'title' || k === 'authors' || k === 'url' ? ' wide' : ''}"><label>${l}</label><input data-f="${k}" value="${esc(s[k] || '')}" spellcheck="false"></div>`).join('')}</div>
          <div class="hint">Authors: separate people with semicolons and write each as <b>Last, First</b> — e.g. <i>Sharma, Priya; Rao, K. N.</i> For an organisation just type its name.</div>
          <div class="src-prev"><div class="src-prev-h">Preview — ${esc(D.CITE_STYLES[bib().style])}</div><div class="src-prev-e"><span>${D.formatCitation([s.id], '', bib().style, new Map([[s.id, 1]])).replace(/^\(Source missing\)$/, '')}</span><p>${D.formatEntry(s, bib().style)}</p></div></div>` : '<div class="empty-note" style="padding:30px 10px">Add the books, articles and web pages you use. They are saved inside this document.</div>'}</div>`;
      const q = $('#src-q', body); q.addEventListener('input', () => { const t = q.value.toLowerCase(); $$('.src-item', body).forEach(b => b.style.display = b.textContent.toLowerCase().includes(t) ? '' : 'none'); });
    };
    const newSrc = () => { const s = { id: 's' + Math.random().toString(36).slice(2, 8), type: 'book', authors: '', title: '', year: String(new Date().getFullYear()) }; work.unshift(s); cur = s.id; render(); setTimeout(() => { const i = $('[data-f="authors"]', body); i && i.focus(); }); };
    body.addEventListener('click', e => {
      const it = e.target.closest('.src-item'); if (it) { cur = it.dataset.id; render(); return; }
      const a = e.target.closest('[data-a]'); if (!a) return;
      if (a.dataset.a === 'new') newSrc();
      if (a.dataset.a === 'lookup') doLookup();
      if (a.dataset.a === 'del') { const used = $$('.cite', D.doc).some(c => (c.dataset.src || '').split(/[;,]/).includes(cur)); const i = work.findIndex(x => x.id === cur); if (i >= 0) work.splice(i, 1); cur = work[0] ? work[0].id : null; render(); if (used) BTS.toast('That source was cited — those citations will show “Source missing”.', { error: true }); }
    });
    body.addEventListener('input', e => { const f = e.target.dataset.f; if (!f) return; const s = work.find(x => x.id === cur); if (!s) return; s[f] = e.target.value; const pe = $('.src-prev-e', body); if (pe) pe.innerHTML = `<span>${D.formatCitation([s.id], '', bib().style, new Map([[s.id, 1]]))}</span><p>${D.formatEntry(s, bib().style)}</p>`; const lb = $(`.src-item[data-id="${CSS.escape(s.id)}"] b`, body); if (lb) lb.textContent = (D.parseAuthors(s.authors).map(lastOf).slice(0, 2).join(', ') || s.title || 'Untitled') + (s.year ? ' (' + s.year + ')' : ''); });
    body.addEventListener('change', e => { if (e.target.id === 'src-type') { const s = work.find(x => x.id === cur); s.type = e.target.value; render(); } });
    body.addEventListener('keydown', e => { if (e.key === 'Enter' && e.target.id === 'src-lookup') { e.preventDefault(); e.stopPropagation(); doLookup(); } });
    let looking = false;
    const doLookup = async () => {
      const inp = $('#src-lookup', body); const q = inp ? inp.value.trim() : ''; const msg = (t, err) => { const m = $('#src-lookup-msg', body); if (m) { m.textContent = t; m.classList.toggle('err', !!err); } };
      if (!q) { msg('Paste a DOI (like 10.1038/nature12373), an ISBN (like 978-0-14-303943-3) or a link that contains a DOI.', true); inp && inp.focus(); return; }
      if (looking) return; looking = true; const btn = $('[data-a="lookup"]', body); if (btn) btn.disabled = true; msg('Looking up…');
      try {
        const f = await D.lookupSource(q);
        let s = work.find(x => x.id === cur);
        if (!s || ((s.title || '').trim() && (s.title || '').trim() !== (f.title || '').trim())) { s = { id: 's' + Math.random().toString(36).slice(2, 8) }; work.unshift(s); cur = s.id; }
        Object.keys(s).forEach(k => { if (k !== 'id') delete s[k]; }); Object.assign(s, f);
        render(); msg(`Found: ${f.title || 'Untitled'}${f.year ? ' (' + f.year + ')' : ''}. Check the details, then Save.`);
      } catch (err) { msg(err.message || String(err), true); }
      finally { looking = false; const b2 = $('[data-a="lookup"]', body); if (b2) b2.disabled = false; const i2 = $('#src-lookup', body); if (i2 && !i2.value) i2.value = q; }
    };
    // temporarily swap so previews resolve against the working list
    const real = bib().sources; bib().sources = work;
    if (opts.newSource) newSrc(); else render();
    const v = await BTS.dialog({ title: 'Manage sources', body, width: 860, buttons: [{ label: 'Cancel', value: null }, { label: opts.pick ? 'Save and continue' : 'Save', primary: true, value: true }] });
    bib().sources = real;
    D.restoreSel(sv);
    if (!v) return null;
    const clean = work.filter(s => (s.title || '').trim() || (s.authors || '').trim());
    D.change(() => { bib().sources = clean; D.updateCitations(); });
    return cur;
  };

  /* ---------- insert / edit citation ---------- */
  D.insertCitation = async (existing) => {
    if (!bib().sources.length) { BTS.toast('Add a source first — it is saved in this document.'); const id = await D.manageSources({ newSource: true, pick: true }); if (!id || !bib().sources.length) return; }
    const sv = D.saveSel();
    const chosen = new Set(existing ? (existing.dataset.src || '').split(/[;,]/).filter(Boolean) : []);
    const body = document.createElement('div'); body.className = 'cite-dlg';
    const render = () => {
      body.innerHTML = `<div class="src-search">${icon('search', 14)}<input id="ct-q" placeholder="Search by author or title" spellcheck="false"></div>
      <div class="cite-list">${bib().sources.map(s => `<label class="cite-item"><input type="checkbox" value="${esc(s.id)}" ${chosen.has(s.id) ? 'checked' : ''}><span><b>${esc(D.parseAuthors(s.authors).map(lastOf).slice(0, 3).join(', ') || 'No author')} ${s.year ? '(' + esc(s.year) + ')' : ''}</b><small>${esc(s.title || 'Untitled')}</small></span></label>`).join('')}</div>
      <div class="dlg-grid"><div class="field"><label>Page or pages (optional)</label><input id="ct-page" value="${esc(existing ? existing.dataset.page || '' : '')}" placeholder="e.g. 45 or 45–47"></div><div class="field"><label>Citation style</label><select id="ct-style">${Object.entries(D.CITE_STYLES).map(([k, v]) => `<option value="${k}" ${bib().style === k ? 'selected' : ''}>${v}</option>`).join('')}</select></div></div>
      <div class="cite-prev">Preview: <span id="ct-prev"></span></div><button class="linkish" data-a="add">+ Add a new source…</button>`;
      const upd = () => { const ids = [...chosen]; const st = $('#ct-style', body).value; const nums = new Map(ids.map((id, i) => [id, i + 1])); $('#ct-prev', body).innerHTML = ids.length ? D.formatCitation(ids, $('#ct-page', body).value.trim(), st, nums) : '<i>Tick one or more sources</i>'; };
      body.oninput = (e) => { if (e.target.id === 'ct-q') { const t = e.target.value.toLowerCase(); $$('.cite-item', body).forEach(b => b.style.display = b.textContent.toLowerCase().includes(t) ? '' : 'none'); } upd(); };
      body.onchange = (e) => { if (e.target.type === 'checkbox') { e.target.checked ? chosen.add(e.target.value) : chosen.delete(e.target.value); } upd(); };
      upd();
    };
    body.addEventListener('click', async e => { if (e.target.closest('[data-a="add"]')) { const id = await D.manageSources({ newSource: true, pick: true }); if (id && D.sourceById(id)) chosen.add(id); render(); } });
    render();
    const v = await BTS.dialog({ title: existing ? 'Edit citation' : 'Insert citation', body, width: 560, buttons: [...(existing ? [{ label: 'Remove citation', value: 'remove', danger: true }] : []), { label: 'Cancel', value: null }, { label: existing ? 'Update' : 'Insert', primary: true, value: (rt) => ({ ids: [...chosen], page: $('#ct-page', rt).value.trim(), style: $('#ct-style', rt).value }) }] });
    D.restoreSel(sv);
    if (v === 'remove') { D.change(() => { existing.remove(); D.updateCitations(); }); return; }
    if (!v || !v.ids.length) return;
    D.change(() => {
      bib().style = v.style;
      let el = existing;
      if (!el) {
        el = document.createElement('span'); el.className = 'cite'; el.contentEditable = 'false';
        const r = D.ensureSel(); if (!D.doc.contains(r.startContainer)) return;
        const live = D.trackBefore ? D.trackBefore() : null;
        const rr = D.range() || r; if (!live) rr.deleteContents();
        // a space before the citation if needed
        const prev = rr.startContainer.nodeType === 3 ? rr.startContainer.data.slice(0, rr.startOffset) : '';
        if (prev && !/\s$/.test(prev)) rr.insertNode(document.createTextNode(' '));
        const nextTxt = rr.endContainer.nodeType === 3 ? rr.endContainer.data.slice(rr.endOffset, rr.endOffset + 1) : '';
        rr.collapse(false); rr.insertNode(el);
        if (nextTxt && /[\p{L}\p{N}(]/u.test(nextTxt)) el.after(document.createTextNode(' '));
        D.setCaret(el.parentNode, [...el.parentNode.childNodes].indexOf(el) + 1);
        if (live) { const t = document.createRange(); t.setStart(live.startContainer, live.startOffset); t.setEndAfter(el); const ins = D.makeTC('ins'); try { t.surroundContents(ins); } catch { } D.setCaret(ins.parentNode, [...ins.parentNode.childNodes].indexOf(ins) + 1); }
      }
      el.dataset.src = v.ids.join(';'); if (v.page) el.dataset.page = v.page; else delete el.dataset.page;
      D.updateCitations();
    });
  };
  D.setCiteStyle = (st) => { if (!D.CITE_STYLES[st]) return; D.change(() => { bib().style = st; D.updateCitations(); }); BTS.toast('Citation style: ' + D.CITE_STYLES[st] + (($$('.cite', D.doc).length || $('.biblio', D.doc)) ? ' — citations and bibliography reformatted' : '')); syncCiteStyle(); };
  function syncCiteStyle() { const s = $('#sel-cite-style'); if (s && document.activeElement !== s) s.value = bib().style; }
  D.syncCiteStyle = syncCiteStyle;
  D.insertBibliography = (kind) => {
    const existing = $('.biblio', D.doc);
    if (existing && kind !== 'new') { D.change(() => D.updateCitations()); D.scrollToEl(existing); BTS.toast('Bibliography updated'); return; }
    if (!bib().sources.length) { BTS.toast('Add sources first (References ▸ Manage sources).'); }
    D.change(() => {
      const b = document.createElement('div'); b.className = 'biblio'; b.contentEditable = 'false';
      if (kind === 'bibliography') b.dataset.title = 'Bibliography';
      D.insertBlocks([b]); D.updateCitations();
    });
  };
  D.bibliographyMenu = (anchor) => BTS.menu(anchor, [{ head: 'Built-in' }, { label: `Insert ${BIB_TITLE[bib().style]} (${D.CITE_STYLES[bib().style]})`, icon: 'book-open', action: () => D.insertBibliography($('.biblio', D.doc) ? 'new' : undefined) }, { label: 'Insert “Bibliography”', icon: 'book-marked', action: () => D.insertBibliography('bibliography') }, { sep: true }, { label: 'Update citations and bibliography', icon: 'refresh-cw', action: () => { D.change(() => D.updateCitations()); BTS.toast('Citations and bibliography updated'); } }, { label: 'Remove bibliography', icon: 'trash-2', disabled: !$('.biblio', D.doc), action: () => D.change(() => $$('.biblio', D.doc).forEach(b => b.remove())) }]);

  /* ---------- captions ---------- */
  D.CAPTION_LABELS = ['Figure', 'Table', 'Equation'];
  D.caption = async () => {
    const sv = D.saveSel(); const r0 = D.range() || D.lastRange();
    const tbl = r0 && D.closest(r0.startContainer, 'table'); const img = !tbl && D.selImg && D.selImg.isConnected ? D.selImg : null;
    const defLabel = img ? 'Figure' : tbl ? 'Table' : 'Figure';
    const v = await BTS.dialog({ title: 'Insert caption', width: 460, body: `<div class="dlg-grid"><div class="field"><label>Label</label><select id="cp-l">${D.CAPTION_LABELS.map(l => `<option ${l === defLabel ? 'selected' : ''}>${l}</option>`).join('')}</select></div><div class="field"><label>Position</label><select id="cp-pos"><option value="below" ${defLabel !== 'Table' ? 'selected' : ''}>Below selected item</option><option value="above" ${defLabel === 'Table' ? 'selected' : ''}>Above selected item</option></select></div></div><div class="field"><label>Caption</label><input id="cp-t" placeholder="e.g. Growth of literacy in Maharashtra"></div><div class="hint">Numbers update automatically (References ▸ Update fields). Use Insert table of figures to list your captions.</div>`, onOpen: (rt) => { $('#cp-l', rt).addEventListener('change', e => { $('#cp-pos', rt).value = e.target.value === 'Table' ? 'above' : 'below'; }); setTimeout(() => $('#cp-t', rt).focus()); }, buttons: [{ label: 'Cancel', value: null }, { label: 'Insert', primary: true, value: (rt) => ({ l: $('#cp-l', rt).value, t: $('#cp-t', rt).value.trim(), pos: $('#cp-pos', rt).value }) }] });
    D.restoreSel(sv); if (!v) return;
    D.change(() => {
      const p = document.createElement('p'); p.className = 'caption'; p.dataset.cap = v.l;
      p.innerHTML = `${esc(v.l)} <span class="seq" data-seq="${esc(v.l)}" contenteditable="false">1</span>${v.t ? ': ' + esc(v.t) : ''}`;
      let anchor = null;
      if (img && img.isConnected) anchor = D.topBlock(img);
      else if (tbl) anchor = D.topBlock(tbl);
      if (anchor && anchor.parentNode === D.doc) { if (v.pos === 'above') anchor.before(p); else anchor.after(p); }
      else {
        const r = D.ensureSel(); const blk = D.blockOf(r.startContainer); const top = blk ? D.topBlock(blk) : null;
        if (top && top.parentNode === D.doc && top.textContent.trim()) { if (v.pos === 'above') top.before(p); else top.after(p); }
        else D.insertBlocks([p]);
      }
      D.updateSeq();
      const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); D.doc.focus({ preventScroll: true }); const s = getSelection(); s.removeAllRanges(); s.addRange(r);
    });
  };
  /** Renumber caption sequence fields (and upgrade plain-text captions from older files). */
  D.updateSeq = () => {
    $$('p.caption', D.doc).forEach(p => {
      if (p.querySelector('.seq')) return;
      const m = p.textContent.match(/^\s*(Figure|Table|Equation|Chart|Diagram|Image|Fig\.)\s+(\d+)/i); if (!m) return;
      const tw = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); let n; while ((n = tw.nextNode())) { const i = n.data.indexOf(m[2]); if (i >= 0) { const num = n.splitText(i); num.splitText(m[2].length); const sp = document.createElement('span'); sp.className = 'seq'; sp.dataset.seq = m[1]; sp.contentEditable = 'false'; num.before(sp); sp.appendChild(num); p.dataset.cap = m[1]; break; } }
    });
    const counts = {};
    $$('.seq', D.doc).forEach(s => { if (s.closest('del.tc-del')) return; const l = s.dataset.seq || 'Figure'; counts[l] = (counts[l] || 0) + 1; if (s.textContent !== String(counts[l])) s.textContent = counts[l]; });
  };

  /* ---------- table of figures ---------- */
  const TOF_TITLE = { Figure: 'Table of Figures', Table: 'List of Tables', Equation: 'List of Equations' };
  D.tofHTML = (label) => {
    D.layout();
    const caps = $$('p.caption', D.doc).filter(p => (p.dataset.cap || (p.querySelector('.seq') || {}).dataset?.seq) === label && !p.closest('.toc'));
    const items = caps.map(p => { const id = D.ensureId(p); return `<p class="toc-l1" data-target="${id}"><span class="toc-t">${esc(p.textContent.trim())}</span><span class="toc-dots"></span><span class="toc-n">${D.pageOfEl(p)}</span></p>`; }).join('');
    return `<div class="toc-title">${esc(TOF_TITLE[label] || 'Table of ' + label + 's')}</div>` + (items || `<p class="toc-empty">No ${label.toLowerCase()} captions found. Use References ▸ Insert caption first, then update this table.</p>`);
  };
  D.insertTOF = async () => {
    const sv = D.saveSel();
    const v = await BTS.dialog({ title: 'Insert table of figures', width: 400, body: `<div class="field"><label>Caption label</label><select id="tf-l">${D.CAPTION_LABELS.map(l => `<option>${l}</option>`).join('')}</select></div><div class="hint">Lists every caption with that label and its page number. Click an entry to jump to it.</div>`, buttons: [{ label: 'Cancel', value: null }, { label: 'Insert', primary: true, value: (rt) => $('#tf-l', rt).value }] });
    D.restoreSel(sv); if (!v) return;
    D.change(() => { D.updateSeq(); const t = document.createElement('div'); t.className = 'toc tof'; t.contentEditable = 'false'; t.dataset.label = v; D.insertBlocks([t]); t.innerHTML = D.tofHTML(v); D.layout(); t.innerHTML = D.tofHTML(v); });
  };
  D.updateTOF = () => { $$('.toc.tof', D.doc).forEach(t => { t.innerHTML = D.tofHTML(t.dataset.label || 'Figure'); }); };

  /* ---------- update all fields ---------- */
  D.updateFields = (silent) => {
    D.change(() => {
      D.updateSeq(); D.updateCitations(); D.layout();
      $$('.toc:not(.tof)', D.doc).forEach(t => { t.innerHTML = D.tocHTML({ title: t.dataset.style === 'toc-index' ? 'Index of topics' : 'Contents' }); });
      D.updateTOF();
    });
    if (!silent) BTS.toast('Fields updated — captions, contents, figures, citations and bibliography');
  };

  D.citeMenu = (anchor, el) => BTS.menu(anchor, [{ head: 'Citation' }, { label: 'Edit citation…', icon: 'quote', action: () => D.insertCitation(el) }, { label: 'Edit sources…', icon: 'library', action: () => D.manageSources() }, { label: 'Remove citation', icon: 'trash-2', action: () => D.change(() => { el.remove(); D.updateCitations(); }) }]);

  D.initRefs = () => {
    D.doc.addEventListener('click', e => { const c = e.target.closest('.cite'); if (c && D.doc.contains(c)) { e.preventDefault(); D.citeMenu({ x: e.clientX, y: e.clientY + 6 }, c); } const b = e.target.closest('.biblio'); if (b && !c && e.detail === 2) D.insertBibliography(); });
    syncCiteStyle();
  };
})();
