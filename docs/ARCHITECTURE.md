# BTS Office — architecture for contributors

Electron desktop app. No bundler, no framework: plain HTML/CSS/JS pages, libraries are pre-built browser bundles in `app/vendor/`.

```
main.js            main process: windows, file dialogs, read/write, print-to-PDF, recent files, single instance, file associations
preload.js         exposes window.bts (secure IPC bridge). Pages never touch Node.
app/index.html     Start hub
app/docs.html      BTS Docs  (word processor)      → app/docs/*.js, *.css
app/sheets.html    BTS Sheets (spreadsheet)        → app/sheets/*.js, *.css
app/slides.html    BTS Slides (presentations)      → app/slides/*.js, *.css
app/pdf.html       BTS PDF   (viewer)              → app/pdf/*.js, *.css
app/common/base.css   design tokens + titlebar/ribbon/backstage/menus/dialogs/status bar styles
app/common/shell.js   window.BTS — shared shell (read the header comment for the full API)
app/common/icons.js   window.icon(name, size) → inline SVG (all Lucide icons, https://lucide.dev/icons)
app/vendor/           mammoth, docx (window.docx), xlsx (SheetJS, window.XLSX), exceljs (window.ExcelJS),
                      hyperformula (window.HyperFormula), chart.umd.js (window.Chart), pptxgen.bundle.js (window.PptxGenJS),
                      jszip.min.js (window.JSZip), pdf.min.mjs + pdf.worker.min.mjs (ES module, pdf.js 4.10)
app/fonts/            Figtree (UI font)
```

## Page skeleton every editor uses

```html
<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'self'; img-src 'self' data: blob:; style-src 'self' 'unsafe-inline'; script-src 'self' 'unsafe-inline' blob:; font-src 'self' data:; worker-src 'self' blob:; connect-src 'self' data: blob:">
<link rel="stylesheet" href="common/base.css"><link rel="stylesheet" href="docs/docs.css">
<script src="common/icons.js"></script><script src="common/shell.js"></script> ...vendor... <script src="docs/docs.js" defer></script>
</head><body data-app="docs">
<nav id="ribbon-tabs"><button class="file-tab">File</button><button data-tab="home">Home</button> ... </nav>
<div id="ribbon">
  <section class="ribbon-panel" data-panel="home">
    <div class="rgroup"><div class="rgroup-body">
        <button class="rb lg" data-cmd="paste" data-icon="clipboard-paste" title="Paste (Ctrl+V)"><span>Paste</span></button>
        <div class="rcol"><div class="rrow"><button class="rb" data-cmd="bold" data-icon="bold" title="Bold (Ctrl+B)"></button>...</div></div>
      </div><div class="rgroup-label">Clipboard</div></div>
  </section>
</div>
<main id="workspace"> ...editor... </main>
<footer id="statusbar"></footer>    <!-- shell appends zoom controls -->
</body></html>
```

* The shell builds `#titlebar` (prepended) and `#backstage` itself. Don't write them.
* `data-icon="name"` on any element inserts the icon at load (DOMContentLoaded). For elements created later use `icon(name, size)`.
* Every element with `data-cmd` inside ribbon / statusbar / panes is routed to `cfg.onCommand(cmd, el, event)`. `<select data-cmd>` / `<input data-cmd>` are routed on `change`.
* Ribbon buttons (`#ribbon .rb`) don't take focus on mousedown, so editor selections survive clicks.
* Built-in commands handled by the shell: file, save, saveas, open, new, print, export-pdf, toggle-ribbon, theme-toggle, zoom-in, zoom-out, zoom-100.
* Button classes: `.rb` (small icon), `.rb.md` (icon+label), `.rb.lg` (big, label in `<span>`), toggled state: `.on` or `aria-pressed="true"`. `.rsel`/`.rinput` for selects/inputs. `.rgallery > .gi` for style galleries. `.pane` for side task panes.
* Colors: use CSS vars (`--accent`, `--ink`, `--ink-2/3`, `--surface`, `--surface-2/3`, `--canvas`, `--line`, `--hover`) so dark mode works. Document content itself (paper, cells, slides) stays light-colored as real documents.
* Mark edits with `BTS.markDirty()`. The shell handles save prompts on close, Ctrl+S/O/N/P, title rename, recent files, drag-and-drop of files.
* `cfg.load(file)` gets `{name, path?, data: Uint8Array}`; `cfg.save(ext)` returns `Uint8Array | ArrayBuffer→Uint8Array | string`.
* `cfg.exportPDF()` should switch the page into print layout (add a class to body) and return printToPDF options (`{ pageSize: 'A4' | 'Letter' | 'Legal' | {width, height} (inches), landscape: bool }`); `cfg.afterExportPDF()` restores. Use `@page` CSS for size/margins (preferCSSPageSize is on by default).

## Testing (headless)

`tools/shot.sh <app> <out.png> [eval.js]` launches the real app under Xvfb, optionally runs `eval.js` in the page
(body of an async function; `return` a value to print it, use `await new Promise(r=>setTimeout(r,300))` for waits),
then saves a screenshot. Renderer console errors are printed. Example:
`tools/shot.sh docs /tmp/claude-0/shots/docs.png /tmp/claude-0/t/docs-test.js`
Then look at the PNG with the Read tool. In tests, call internal functions / BTS.cfg.save(ext) directly
(native dialogs can't be clicked). You can verify generated files by writing them via `bts.writeFile('/tmp/...', data)`
then inspecting them with Python (python-docx, openpyxl, python-pptx may be pip-installable with --break-system-packages).

## Automatic updates

`updater.js` (main process) uses electron-updater with a generic feed at `https://office.btspl.tech/updates/`.
It checks 8 s after launch and every 4 h, downloads in the background, verifies SHA-512, and installs silently on quit
(or immediately via "Restart now", which first refuses if any window has unsaved changes). State is broadcast to every
window over `bts:update-state`. `app/common/updates.js` (loaded by every page) shows the "Update ready" card and the
Updates panel on File ▸ Options & about. Overrides: env `BTS_UPDATE_URL`, or `updateUrl` / `autoUpdate` in
`%APPDATA%\BTS Office\settings.json`. Dev runs (`npm start`) don't check unless `BTS_UPDATE_URL` is set (they use
`dev-app-update.yml`). See docs/RELEASING.md for publishing and docs/HOSTING.md for the server.

## Fonts, command search, properties, versions (1.2)

* **Font library:** `tools/build-fonts.js` copies 50 OFL/Apache families from `@fontsource/*` into `app/fonts/lib/` and generates
  `app/common/fonts.css` (@font-face + local-first aliases: Calibri→Carlito, Cambria→Caladea, Arial→Arimo, Times New Roman→Tinos,
  Courier New→Cousine, Georgia→Gelasio) and `app/common/fonts.js` (`window.BTS_FONTS` registry: `{name, cat, bundled, weights, italic, alias}`,
  categories in `BTS_FONT_CATEGORIES`, helpers `btsFontStack(name)`, `btsFontAvailable(name)`, `btsRecentFonts(add?)`).
  Docs/Sheets/Slides include both files. Use `BTS.fontPicker(anchor, { current, theme:[{label,name}] })` → Promise(name|null) for every font dropdown.
* **Command search:** `BTS.commandSearch()` (Alt+Q, Ctrl+Shift+P, lightbulb in title bar) indexes every `#ribbon [data-cmd]` button by its title/text,
  group and tab. Give every ribbon button a clear `title` so it is findable.
* **Properties:** `BTS.props = {title, subject, author, company, keywords, comments, category}` is edited on File ▸ Info. Apps must write it into
  saved files (docx/xlsx/pptx core properties) and fill it with `BTS.setProps({...})` when loading. `BTS.resetProps()` runs before `cfg.load` and on new.
  `BTS.userName` (File ▸ Options) is the default author and the name to use for comments and tracked changes.
* **Version history:** every save of .docx/.xlsx/.pptx is copied by `main.js` to `userData/versions/` (last 30 per file). File ▸ Info lists them;
  clicking one opens it as a copy in a new window. Nothing for apps to do.
