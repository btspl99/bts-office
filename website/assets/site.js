/* BTS Office website behaviour: icons, live download info from the update feed, release notes. */
(function () {
  const S = window.BTS_SITE || {};
  const $$ = (s) => [...document.querySelectorAll(s)];
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  // icons
  $$('[data-i]').forEach(el => {
    const body = (window.SITE_ICONS || {})[el.dataset.i]; if (!body) return;
    el.outerHTML = '<svg class="ic" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + body + '</svg>';
  });

  // links from config
  $$('[data-github]').forEach(a => { a.href = S.github; a.target = '_blank'; a.rel = 'noopener'; });
  $$('[data-github-issues]').forEach(a => { a.href = S.github + '/issues'; a.target = '_blank'; a.rel = 'noopener'; });
  $$('[data-mail]').forEach(a => { a.href = 'mailto:' + S.supportEmail; });
  $$('[data-org]').forEach(e => { e.textContent = S.organisation || e.textContent; });

  // sticky header border
  const top = document.querySelector('.top');
  addEventListener('scroll', () => top.classList.toggle('scrolled', scrollY > 8), { passive: true });

  // one orchestrated moment: the sheets fan out on load
  const stack = document.querySelector('.stack');
  if (stack && !matchMedia('(prefers-reduced-motion: reduce)').matches) {
    stack.classList.add('closed');
    requestAnimationFrame(() => setTimeout(() => stack.classList.remove('closed'), 120));
  }

  // minimal reader for electron-builder's latest.yml
  function parseLatest(txt) {
    const get = (k) => { const m = txt.match(new RegExp('^' + k + ':\\s*[\'"]?([^\'"\\n]+)', 'm')); return m ? m[1].trim() : null; };
    const size = txt.match(/^\s+size:\s*(\d+)/m);
    return { version: get('version'), file: get('path'), date: (get('releaseDate') || '').slice(0, 10), size: size ? +size[1] : 0 };
  }
  const fmtSize = (b) => b ? Math.round(b / 1048576) + ' MB' : '';
  const fmtDate = (d) => { try { return new Date(d + 'T00:00:00').toLocaleDateString('en-IN', { day: 'numeric', month: 'long', year: 'numeric' }); } catch { return d; } };

  function applyDownload(info) {
    if (!info || !info.file) return;
    const href = 'updates/' + encodeURIComponent(info.file);
    $$('[data-download-link]').forEach(a => { a.href = href; a.setAttribute('download', ''); });
    const parts = ['Version ' + info.version, fmtSize(info.size), 'Windows 10 and 11, 64-bit'].filter(Boolean);
    $$('[data-dl-meta]').forEach(p => { p.textContent = parts.join(', '); });
  }
  applyDownload(S.fallback);
  fetch('updates/latest.yml', { cache: 'no-store' }).then(r => r.ok ? r.text() : Promise.reject()).then(t => applyDownload(parseLatest(t))).catch(() => {});

  // release notes
  fetch('releases.json', { cache: 'no-cache' }).then(r => r.json()).then(list => {
    const el = document.getElementById('releases');
    el.innerHTML = list.slice(0, 6).map((r, i) => `<article class="rel ${i === 0 ? 'latest' : ''}"><div><h3>${esc(r.version)}</h3><time datetime="${esc(r.date)}">${esc(fmtDate(r.date))}</time></div><ul>${(r.notes || []).map(n => `<li>${esc(n)}</li>`).join('')}</ul></article>`).join('');
  }).catch(() => { document.getElementById('releases').innerHTML = '<p class="muted">Release notes are available on GitHub.</p>'; });
})();
