#!/usr/bin/env bash
# Build a release and prepare dist/site for office.btspl.tech.
#   npm run release            build + assemble
#   npm run release -- --deploy   …and upload via FTP (needs .env.deploy or FTP_* variables)
set -euo pipefail
cd "$(dirname "$0")/.."
node tools/assemble-site.js notes
npm run dist:win
node tools/assemble-site.js
if [[ "${1:-}" == "--deploy" ]]; then node tools/deploy-ftp.js; fi
