# Publishing an update

Every installed copy checks `https://office.btspl.tech/updates/latest.yml` a few seconds after starting, and again every 4 hours.
When the version there is higher than its own, it downloads the new installer in the background, checks its SHA-512, and installs it silently the next time the user closes BTS Office (or right away if they click **Restart now**).

## Steps

1. **Bump the version** in `package.json`, for example `1.1.0` → `1.2.0`. Use [semantic versioning](https://semver.org): patch for fixes, minor for features.
2. **Write the release notes**: add an entry at the top of `website/releases.json`:
   ```json
   { "version": "1.2.0", "date": "2026-11-01", "notes": ["What changed, in plain words students understand."] }
   ```
   These notes appear on the website, in the app's update card ("What's new") and on the GitHub Release.
3. Commit, then publish one of two ways:

### Option 1: GitHub Actions (no Windows PC needed)
```bash
git commit -am "Release 1.2.0"
git tag v1.2.0
git push && git push --tags
```
The **Release** workflow builds the installer on Windows, creates the GitHub Release and, if the FTP secrets are set, uploads everything to office.btspl.tech. Without FTP secrets, download the `site-v1.2.0` artifact from the workflow run and upload it by hand (see [HOSTING.md](HOSTING.md)).

### Option 2: from your own computer
```bash
npm run release              # builds dist/BTS-Office-Setup-1.2.0.exe and dist/site/
npm run deploy               # uploads dist/site/ to Hostinger (needs .env.deploy)
# or in one go:
npm run release -- --deploy
```
On Linux, building the Windows installer needs Wine (`sudo apt install wine64 wine32:i386`). On Windows it works natively, but run `npm run notes`, `npm run dist:win` and `npm run assemble` instead of `npm run release`, since that script uses bash.

## Rules that keep updates safe

* **Never reuse a version number.** PCs compare versions, and the installer name includes the version.
* **Upload `latest.yml` last.** The deploy script does this for you.
* **Don't edit files after they are published.** If something is wrong, release a new version.
* **Roll back** by releasing a higher version that contains the old code. Downgrades are disabled.
* Test the installer on a real Windows PC before publishing. One option is to publish to a test folder first and point a test PC at it with `"updateUrl"` in `%APPDATA%\BTS Office\settings.json`.

## Code signing (recommended once you have users)

Unsigned installers show "Windows protected your PC". With a code-signing certificate (an EV certificate also builds SmartScreen reputation faster):
1. Remove `"signAndEditExecutable": false` from `package.json` → `build.win`.
2. Set `CSC_LINK` (path to the .pfx, or base64 of it) and `CSC_KEY_PASSWORD`. In GitHub, add them as the `WIN_CSC_LINK` and `WIN_CSC_KEY_PASSWORD` secrets.
3. Optionally set `build.win.publisherName` to the certificate's subject name. electron-updater will then refuse any update not signed by you.
