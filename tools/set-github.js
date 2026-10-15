#!/usr/bin/env node
// One-time: node tools/set-github.js <account>/<repo>   e.g. node tools/set-github.js btspl/bts-office
// Replaces the YOUR-GITHUB-ACCOUNT/bts-office placeholder everywhere it appears.
const fs = require('fs'), path = require('path');
const target = process.argv[2];
if (!/^[\w.-]+\/[\w.-]+$/.test(target || '')) { console.error('Usage: node tools/set-github.js <account>/<repo>'); process.exit(1); }
const root = path.join(__dirname, '..');
const files = ['package.json', 'README.md', 'CONTRIBUTING.md', 'website/config.js'];
for (const f of files) {
  const p = path.join(root, f); let s = fs.readFileSync(p, 'utf8');
  const n = s.split('YOUR-GITHUB-ACCOUNT/bts-office').length - 1;
  if (n) { fs.writeFileSync(p, s.split('YOUR-GITHUB-ACCOUNT/bts-office').join(target)); console.log(`✔ ${f} (${n})`); }
}
