# Hosting the website and update server on Hostinger

The website at **https://office.btspl.tech** does two jobs:

1. It is the download page students visit.
2. Its `updates/` folder is the update server. Every installed copy of BTS Office reads
   `https://office.btspl.tech/updates/latest.yml` to find out whether a new version exists.

It is all static files, so any Hostinger plan works. No database or PHP is needed.

```
public_html/  (of office.btspl.tech)
├── index.html, config.js, releases.json, favicon.ico, .htaccess
├── assets/                 css, js, fonts, screenshots
└── updates/
    ├── .htaccess           no-cache for latest.yml, download headers for installers
    ├── latest.yml          ← the update feed (upload last)
    ├── BTS-Office-Setup-1.1.0.exe
    └── BTS-Office-Setup-1.1.0.exe.blockmap
```

## One-time setup

### 1. Create the subdomain
hPanel → **Websites** → your btspl.tech site → **Domains → Subdomains** → create `office`.
Hostinger gives it its own folder, usually `domains/office.btspl.tech/public_html` (hPanel shows the exact path).

**DNS:** if btspl.tech's DNS is managed somewhere other than Hostinger (for example AWS Route 53, because other btspl.tech services run on AWS), add an **A record** there:
`office` → the server IP shown in hPanel (**Websites → Dashboard → Plan details**). If the DNS is at Hostinger, this happens automatically.

### 2. Turn on SSL (required)
hPanel → **Security → SSL** → install the free SSL certificate for `office.btspl.tech`.
The app refuses updates over plain HTTP, and the `.htaccess` redirects every visit to HTTPS.
Wait until `https://office.btspl.tech` opens without a warning before publishing.

### 3. Create an FTP account for publishing (optional, for automatic uploads)
hPanel → **Files → FTP Accounts** → create an account whose directory is the subdomain's `public_html`.
Note the host, username and password. Then either:

* **From your PC:** create a file named `.env.deploy` in the project folder (it is git-ignored):
  ```
  FTP_HOST=ftp.btspl.tech
  FTP_USER=u123456789.office
  FTP_PASSWORD=your-password
  FTP_DIR=/
  FTP_SECURE=true
  ```
  Use `FTP_DIR=/` when the FTP account is limited to the subdomain folder; otherwise give the full path, for example `/domains/office.btspl.tech/public_html`.
  If the upload fails with a TLS/certificate error, set `FTP_SECURE=false`.
* **From GitHub Actions:** add the same five values as repository secrets
  (**Settings → Secrets and variables → Actions**): `FTP_HOST`, `FTP_USER`, `FTP_PASSWORD`, `FTP_DIR`, `FTP_SECURE`.

## Publishing the website (and each update)

Build first (see [RELEASING.md](RELEASING.md)); the result is in `dist/site/`. Then use one of these:

**A. Automatic (recommended):** `npm run deploy` from your PC, or push a version tag and let GitHub Actions do it.
The script uploads in a safe order: website files first, then the installer and `.blockmap`, and **`latest.yml` last**. Installed copies never see a version whose installer isn't fully uploaded yet. It never deletes anything on the server.

**B. By hand with the File Manager:**
1. Open hPanel → **Files → File Manager** → the subdomain's `public_html`.
2. Upload everything in `dist/site/` **except** `updates/latest.yml`. Tip: zip `dist/site`, upload the zip, and use **Extract**, then delete the zip. Remove `updates/latest.yml` from the zip before uploading it.
3. Once the `.exe` and `.exe.blockmap` have finished uploading, upload `updates/latest.yml`.
4. Make sure hidden files (`.htaccess`) were uploaded. The File Manager shows them under **Settings → Show hidden files**.

## Check it works

* `https://office.btspl.tech/` shows the page, and the button says the new version.
* `https://office.btspl.tech/updates/latest.yml` shows the version, file name and checksum.
* Run `curl -I https://office.btspl.tech/updates/latest.yml`; it should include `Cache-Control: no-cache`.
* Run `curl -I -H "Range: bytes=0-99" https://office.btspl.tech/updates/BTS-Office-Setup-1.1.0.exe`; it should answer `206 Partial Content`. That allows small differential updates.
* Open BTS Office → **File → Options & about → Check for updates**.

## Housekeeping

* Keep old `.blockmap` files; they let PCs download only the changed parts of a new version. They are small.
* Old `.exe` files can be deleted after a few weeks to save space. Keep at least the current and previous versions.
* If Hostinger's CDN or LiteSpeed cache is on, exclude `/updates/latest.yml` from caching, or purge the cache after each release. Otherwise PCs may not see the new version straight away.

## School or lab mirror

A school can copy the `updates/` folder to its own web server and set `"updateUrl": "https://their-server/path/"` in
`%APPDATA%\BTS Office\settings.json` on each PC. `"autoUpdate": false` turns automatic checks off entirely.
