#!/usr/bin/env node
/* Release helper.
 *
 *   node tools/assemble-site.js notes   → writes build/release-notes.md from website/releases.json
 *                                          (run before building, so the notes go into latest.yml)
 *   node tools/assemble-site.js         → builds dist/site/ = the website + updates/ feed,
 *                                          ready to upload to office.btspl.tech
 *
 * Before tagging a release: bump "version" in package.json and add an entry at the top of
 * website/releases.json with the same version.
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const root = path.join(__dirname, '..');
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
const version = pkg.version;
const releases = JSON.parse(fs.readFileSync(path.join(root, 'website', 'releases.json'), 'utf8'));
const entry = releases.find(r => r.version === version);
const die = (m) => { console.error('✖ ' + m); process.exit(1); };

if (!entry) die(`website/releases.json has no entry for version ${version}. Add one at the top, then try again.`);
if (releases[0].version !== version) die(`The first entry in website/releases.json is ${releases[0].version}, but package.json says ${version}.`);

if (process.argv[2] === 'notes') {
  fs.writeFileSync(path.join(root, 'build', 'release-notes.md'), entry.notes.map(n => '- ' + n).join('\n') + '\n');
  console.log(`✔ build/release-notes.md written for ${version}`);
  process.exit(0);
}

const dist = path.join(root, 'dist');
const exe = `BTS-Office-Setup-${version}.exe`;
const need = [exe, exe + '.blockmap', 'latest.yml'];
for (const f of need) if (!fs.existsSync(path.join(dist, f))) die(`dist/${f} is missing. Run "npm run dist:win" first.`);

// sanity: latest.yml must describe this version and match the installer's checksum
const yml = fs.readFileSync(path.join(dist, 'latest.yml'), 'utf8');
const ymlVersion = (yml.match(/^version:\s*(.+)$/m) || [])[1];
const ymlSha = (yml.match(/^sha512:\s*(.+)$/m) || [])[1];
if (ymlVersion !== version) die(`dist/latest.yml is for ${ymlVersion}, expected ${version}.`);
const sha = crypto.createHash('sha512').update(fs.readFileSync(path.join(dist, exe))).digest('base64');
if (sha !== ymlSha) die('The installer checksum does not match latest.yml. Rebuild with "npm run dist:win".');
const size = fs.statSync(path.join(dist, exe)).size;

// copy website → dist/site
const out = path.join(dist, 'site');
fs.rmSync(out, { recursive: true, force: true });
fs.cpSync(path.join(root, 'website'), out, { recursive: true });
for (const f of need) fs.copyFileSync(path.join(dist, f), path.join(out, 'updates', f));

// keep the offline fallback in config.js current
const cfgPath = path.join(out, 'config.js');
const date = (yml.match(/^releaseDate:\s*'?([0-9-]{10})/m) || [])[1] || new Date().toISOString().slice(0, 10);
let cfg = fs.readFileSync(cfgPath, 'utf8');
cfg = cfg.replace(/fallback:\s*\{[^}]*\}/, `fallback: { version: '${version}', file: '${exe}', size: ${size}, date: '${date}' }`);
fs.writeFileSync(cfgPath, cfg);
fs.writeFileSync(path.join(root, 'website', 'config.js'), cfg); // so the repo copy stays in step

console.log(`✔ dist/site is ready for version ${version}`);
console.log(`  updates/${exe}  (${Math.round(size / 1048576)} MB)`);
console.log('  updates/latest.yml   ← upload this LAST, after the installer and .blockmap');
console.log('  Upload: npm run deploy  (FTP), or the Hostinger File Manager. See docs/HOSTING.md.');
