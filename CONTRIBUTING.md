# Contributing to BTS Office

Thanks for helping make a free office suite for students. Bug reports, sample files that open wrongly, translations, design fixes and code are all welcome.

## Reporting a problem
Open an issue using the **Bug report** form. Include the app, the version (File > Options & about), the steps, and a sample file if a document opens or saves wrongly. Remove anything private from the file first.

## Working on the code
```bash
git clone https://github.com/btspl99/bts-office.git
cd bts-office
npm install
npm start                # opens the start screen
npm start -- --app=docs  # or open one app directly: docs, sheets, slides, pdf
```
There is no build step for the app itself: edit the HTML/CSS/JS under `app/` and restart. Read [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) first. It explains the shared shell (`app/common/shell.js`), how ribbon commands are wired, and the headless screenshot test tool.

## Guidelines
* **Keep it offline.** The app must work with no internet. Put libraries in `app/vendor/`; never load anything from a CDN.
* **Keep pages sandboxed.** Pages have no Node.js access; anything that touches files goes through `preload.js` → `main.js`.
* **Use the design tokens** in `app/common/base.css` (`--accent`, `--ink`, `--surface`, …) so light and dark themes both work.
* **Write for students.** Use plain, short labels in sentence case, and say what a button does ("Save a copy", not "Submit").
* **Test file formats both ways.** If you touch import/export, save a file and open it in another office app, and open a file made elsewhere.
* One change per pull request, with a screenshot for anything visual.

## License
By contributing you agree that your work is released under the project's GPL-3.0-or-later license.
