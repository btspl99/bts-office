// BTS Office — small, allow-listed internet lookups used by optional online features
// (citation lookup by DOI/ISBN). Pages cannot reach the internet themselves (CSP); they ask here.
// Uses Electron's net module, so the computer's proxy settings (common in college labs) apply.
const ALLOWED = ['api.crossref.org', 'openlibrary.org', 'www.googleapis.com'];

async function getJSON(url) {
  const { net } = require('electron');
  let u; try { u = new URL(url); } catch { throw new Error('Bad address'); }
  if (u.protocol !== 'https:' || !ALLOWED.includes(u.hostname)) throw new Error('Address not allowed');
  const ctl = new AbortController(); const t = setTimeout(() => ctl.abort(), 15000);
  try {
    const res = await net.fetch(u.href, { signal: ctl.signal, redirect: 'follow', headers: { 'User-Agent': 'BTS-Office (https://office.btspl.tech; mailto:btspl99@gmail.com)', Accept: 'application/json' } });
    if (!ALLOWED.includes(new URL(res.url || u.href).hostname)) throw new Error('Address not allowed');
    if (res.status === 404) throw new Error('Not found');
    if (!res.ok) throw new Error('Server answered ' + res.status);
    const body = await res.text();
    if (body.length > 4e6) throw new Error('Too large');
    try { return JSON.parse(body); } catch { throw new Error('Unexpected answer'); }
  } catch (err) {
    if (err.name === 'AbortError') throw new Error('Timed out');
    throw err;
  } finally { clearTimeout(t); }
}

function init({ ipcMain }) {
  ipcMain.handle('bts:ext:fetch-json', async (e, url) => {
    try { return { ok: true, data: await getJSON(url) }; }
    catch (err) { return { ok: false, error: /ERR_INTERNET_DISCONNECTED|ERR_NAME_NOT_RESOLVED|ERR_CONNECTION|ERR_PROXY|ERR_ADDRESS_UNREACHABLE|ENOTFOUND|EAI_AGAIN|ECONN|ENETUNREACH|Timed out|net::/i.test(err.message) ? 'No internet connection' : err.message }; }
  });
}
module.exports = { init, getJSON };
