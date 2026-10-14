# BTS Office

A free, open-source office suite for students. Write documents, build spreadsheets, make presentations and read PDFs, using the file formats schools and colleges already use.

| App | Opens | Saves |
|---|---|---|
| **BTS Docs** (word processor) | .docx, .odt, .html, .md, .txt, .rtf | .docx, .html, .md, .txt, PDF |
| **BTS Sheets** (spreadsheet) | .xlsx, .xls, .ods, .csv, .tsv | .xlsx, .csv, .ods, .html, PDF |
| **BTS Slides** (presentations) | .pptx | .pptx, PDF |
| **BTS PDF** (reader & annotator) | .pdf | annotated .pdf |

## Install (Windows)

Run `BTS-Office-Setup-1.0.0.exe`. You can choose the install folder. The installer adds a desktop shortcut and Start menu entries for BTS Office, Docs, Sheets, Slides and PDF, and registers the file types above so double-clicking a file opens it.

The installer isn't code-signed yet, so Windows SmartScreen may show "Windows protected your PC". Click **More info → Run anyway**. To remove that warning for everyone, sign the installer with a code-signing certificate (see *Signing* below).

## Highlights

- **Docs:** ribbon with Home, Insert, Design, Layout, References, Review and View. Real pages (A4, Letter, Legal), styles, lists, tables with merge, pictures with text wrap, headers and footers with page numbers, table of contents, footnotes, comments, watermark, find and replace, word count, read aloud, spell check, crash recovery and templates.
- **Sheets:** formula engine with 400+ Excel-compatible functions, fill handle, freeze panes, merge, number formats (₹ by default), conditional formatting, tables, sort and filter, data validation, charts, named ranges, multiple sheets, full undo/redo and templates.
- **Slides:** themes, layouts, shapes, pictures, tables, charts, icons, transitions, animations, a full-screen slide show with pen and laser pointer, slide sorter, notes and templates.
- **PDF:** fast viewing, thumbnails, bookmarks, search, zoom and rotate. Highlight, draw and add notes, then save an annotated copy.
- Light and dark themes, drag-and-drop to open, recent files, and keyboard shortcuts that match the usual office apps.

## Build from source

Requirements: Node.js 20+.

```bash
npm install
npm start                 # run the app
npm run dist:win          # Windows installer -> dist/BTS-Office-Setup-<version>.exe
npm run dist:linux        # AppImage + .deb
```

Building the Windows installer on Linux needs Wine (`wine64` and `wine32`). On Windows it builds natively.

After upgrading an icon or library package, refresh the copies in `app/vendor/` (and run `npm run icons` for the icon set). The app loads only vendored files, so it works fully offline.

## Project layout

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md). In short: `main.js` (windows and file access), `preload.js` (secure bridge), `app/common/` (shared shell, ribbon, dialogs, design tokens) and one folder per app under `app/`.

## Signing (optional, recommended for wide distribution)

Set `CSC_LINK` (path or base64 of a .pfx) and `CSC_KEY_PASSWORD`, then remove `"signAndEditExecutable": false` from `package.json` and run `npm run dist:win`.

## License

GPL-3.0-or-later. See `LICENSE`. Third-party components keep their own licenses: Electron (MIT), HyperFormula (GPL-3.0), ExcelJS (MIT), SheetJS Community (Apache-2.0), docx (MIT), Mammoth (BSD-2), PptxGenJS (MIT), JSZip (MIT/GPL-3.0), PDF.js (Apache-2.0), pdf-lib (MIT), Chart.js (MIT), Lucide icons (ISC), Figtree and bundled slide fonts (SIL OFL 1.1).

BTS Office is an independent project. It is not affiliated with Microsoft or The Document Foundation; product names belong to their owners.
