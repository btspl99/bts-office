// BTS Office — small, allow-listed internet lookups used by optional online features
// (citation lookup by DOI/ISBN). Pages cannot reach the internet themselves (CSP); they ask here.
const https = require('https');
const ALLOWED = ['api.crossref.org', 'openlibrary.org', 'www.googleapis.com'];

function getJSON(url, redirects = 3) {
  return new Promise((resolve, reject) => {
    let u; try { u = new URL(url); } catch { return reject(new Error('Bad address')); }
    if (u.protocol !== 'https:' || !ALLOWED.includes(u.hostname)) return reject(new Error('Address not allowed'));
    const req = https.get(u, { headers: { 'User-Agent': 'BTS-Office (https://office.btspl.tech; mailto:btspl99@gmail.com)', Accept: 'application/json' }, timeout: 15000 }, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location && redirects > 0) { res.resume(); return resolve(getJSON(new URL(res.headers.location, u).href, redirects - 1)); }
      if (res.statusCode === 404) { res.resume(); return reject(new Error('Not found')); }
      if (res.statusCode !== 200) { res.resume(); return reject(new Error('Server answered ' + res.statusCode)); }
      let body = ''; res.setEncoding('utf8');
      res.on('data', c => { body += c; if (body.length > 4e6) req.destroy(new Error('Too large')); });
      res.on('end', () => { try { resolve(JSON.parse(body)); } catch { reject(new Error('Unexpected answer')); } });
    });
    req.on('timeout', () => req.destroy(new Error('Timed out')));
    req.on('error', reject);
  });
}

function init({ ipcMain }) {
  ipcMain.handle('bts:ext:fetch-json', async (e, url) => {
    try { return { ok: true, data: await getJSON(url) }; }
    catch (err) { return { ok: false, error: /ENOTFOUND|EAI_AGAIN|ECONN|ENETUNREACH|Timed out/i.test(err.message) ? 'No internet connection' : err.message }; }
  });
}
module.exports = { init, getJSON };
