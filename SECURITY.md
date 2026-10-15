# Security policy

## Reporting a vulnerability
Please **don't** open a public issue for security problems. Email **btspl99@gmail.com** with the subject "BTS Office security", and include steps to reproduce. We aim to reply within 7 days and to ship a fix through the automatic updater as soon as possible.

## Supported versions
Only the latest release receives fixes. Installed copies update themselves, so most users are always on it.

## How updates are protected
* Updates are downloaded only over HTTPS from `https://office.btspl.tech/updates/`.
* Each installer's SHA-512 checksum is published in `latest.yml` and verified before installing; a mismatched file is discarded.
* Downgrades are refused.
* Once releases are code-signed, setting `publisherName` makes the app also reject installers that weren't signed with the project's certificate.

## App hardening
Every document window runs with `contextIsolation`, the Chromium `sandbox`, and no Node.js integration. Pages have a strict Content-Security-Policy and load no remote code. Links open in the system browser, and HTML documents are stripped of scripts and event handlers when opened.
