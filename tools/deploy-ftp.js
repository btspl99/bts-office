#!/usr/bin/env node
/* Uploads dist/site/ to Hostinger (or any FTP host) in a safe order:
 *   1. website files          2. installer + .blockmap          3. updates/latest.yml (last)
 * Installed apps only see a new version once latest.yml changes, so they never point at a
 * half-uploaded installer. Nothing on the server is deleted, so older .blockmap files stay
 * available for small differential updates.
 *
 * Settings (environment variables or a git-ignored .env.deploy file in the project root):
 *   FTP_HOST      e.g. ftp.btspl.tech  or the IP shown in Hostinger ▸ Files ▸ FTP Accounts
 *   FTP_USER      FTP username
 *   FTP_PASSWORD  FTP password
 *   FTP_DIR       folder of the office.btspl.tech site, e.g. /domains/office.btspl.tech/public_html
 *                 (use / if the FTP account is already limited to that folder)
 *   FTP_SECURE    "true" (default, FTPS) or "false" if your host only allows plain FTP
 *   FTP_PORT      default 21
 */
const fs = require('fs');
const path = require('path');
const ftp = require('basic-ftp');

const root = path.join(__dirname, '..');
const envFile = path.join(root, '.env.deploy');
if (fs.existsSync(envFile)) {
  for (const line of fs.readFileSync(envFile, 'utf8').split(/\r?\n/)) {
    const m = line.match(/^\s*([A-Z_]+)\s*=\s*(.*)\s*$/); if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
  }
}
const { FTP_HOST, FTP_USER, FTP_PASSWORD } = process.env;
const FTP_DIR = (process.env.FTP_DIR || '/').replace(/\/+$/, '') || '/';
const secure = (process.env.FTP_SECURE || 'true') !== 'false';
const site = path.join(root, 'dist', 'site');
if (!FTP_HOST || !FTP_USER || !FTP_PASSWORD) { console.error('✖ Set FTP_HOST, FTP_USER and FTP_PASSWORD (see the top of tools/deploy-ftp.js).'); process.exit(1); }
if (!fs.existsSync(path.join(site, 'updates', 'latest.yml'))) { console.error('✖ dist/site is missing. Run "npm run release" first.'); process.exit(1); }

function walk(dir, base = '') {
  const out = [];
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const rel = base ? base + '/' + e.name : e.name;
    if (e.isDirectory()) out.push(...walk(path.join(dir, e.name), rel)); else out.push(rel);
  }
  return out;
}

(async () => {
  const client = new ftp.Client(60_000);
  client.ftp.verbose = !!process.env.FTP_VERBOSE;
  try {
    await client.access({ host: FTP_HOST, port: +(process.env.FTP_PORT || 21), user: FTP_USER, password: FTP_PASSWORD, secure, secureOptions: { rejectUnauthorized: process.env.FTP_INSECURE_TLS !== '1' } });
    const files = walk(site);
    const isBinary = (f) => /^updates\/.+\.(exe|blockmap|AppImage)$/.test(f);
    const order = [
      ...files.filter(f => !f.startsWith('updates/') || (!isBinary(f) && f !== 'updates/latest.yml')),
      ...files.filter(f => isBinary(f)),
      'updates/latest.yml',
    ];
    for (const rel of order) {
      const remote = (FTP_DIR === '/' ? '' : FTP_DIR) + '/' + rel;
      await client.ensureDir(path.posix.dirname(remote));
      await client.cd('/');
      const local = path.join(site, rel);
      const size = fs.statSync(local).size;
      if (isBinary(rel)) {
        // skip installers that are already fully uploaded
        try { if ((await client.size(remote)) === size) { console.log('= ' + rel + ' (already on server)'); continue; } } catch {}
        process.stdout.write('↑ ' + rel + ' (' + Math.round(size / 1048576) + ' MB)… ');
        await client.uploadFrom(local, remote + '.part');
        await client.rename(remote + '.part', remote);
        const remoteSize = await client.size(remote);
        if (remoteSize !== size) throw new Error(`${rel} uploaded ${remoteSize} of ${size} bytes`);
        console.log('done');
      } else {
        await client.uploadFrom(local, remote);
        console.log('↑ ' + rel);
      }
    }
    console.log('\n✔ Published. Installed copies will pick up the update within a few hours (or at next launch).');
  } catch (err) {
    console.error('\n✖ Upload failed: ' + err.message);
    if (/certificate|TLS|SSL/i.test(err.message)) console.error('  Tip: if your host has no FTPS, set FTP_SECURE=false.');
    process.exitCode = 1;
  } finally { client.close(); }
})();
