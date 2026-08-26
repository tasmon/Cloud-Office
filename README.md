# Cloud Office 2.0.1

Cloud Word, Cloud Sheet and Cloud PowerPoint. Three office apps that run
entirely in your browser. No server, no account, no network requests at
runtime. Everything is saved to this browser's local storage (IndexedDB)
and stays on your device.

## Files

Every file sits in one flat folder, no subfolders. This is intentional so
the whole thing can be unzipped and opened directly, or dropped onto any
static file host with zero configuration.

```
index.html          hub / launcher, cross-app recent files list
word.html            Cloud Word
sheet.html           Cloud Sheet
slides.html          Cloud PowerPoint
help.html            help center
about.html           about / version / credits

theme.css             design tokens, light/dark theme, [hidden] fix
common.css            shared chrome: menu bar, sidebar, modals
common.js              shared behavior: theme toggle, menu bar, modals
db.js                  tiny IndexedDB helper used by all three apps
undo.js                 generic undo/redo history manager

word.css / word.js
sheet.css / sheet.js
slides.css / slides.js

manifest.json         PWA manifest for the whole suite
sw.js                  service worker, caches the app shell for offline use

icon-*.svg / icon-*.png    app icons

mammoth.browser.min.js     reads .docx (Cloud Word)
html-docx.js                writes .docx (Cloud Word)
xlsx.full.min.js            reads/writes .xlsx and .csv (Cloud Sheet)
pptxgen.bundle.js           writes .pptx (Cloud PowerPoint)
```

The four libraries above are the only third-party code included. They are
bundled as plain local files: nothing is fetched from a CDN or any other
URL at runtime. Open any `.html` file directly (`file://`) or serve the
folder with any static file server and it works offline.

## Run it locally

```bash
cd cloud-office
python3 -m http.server 8080
# open http://localhost:8080
```

Opening `index.html` directly by double-clicking also works in most
browsers; a local server is only slightly more reliable for IndexedDB and
the service worker.

## Deploy

Any static host works: GitHub Pages, Vercel, Netlify, an S3 bucket, or an
internal file share. Push the contents of this folder (not the folder
itself) to the root of the site. No build step.

## What's new in 2.0.1

- Every file is local. No CDN dependencies, works fully offline.
- Flat file layout: one folder, no subfolders.
- Redesigned hub page: an app launcher with a cross-app recent-files list,
  not a marketing page.
- Menu bars (File, Edit, Insert, Format, ...) in all three apps, in
  addition to the icon toolbar.
- Cloud Sheet and Cloud PowerPoint now support multiple saved files with
  the same sidebar pattern Cloud Word already had (New / switch / delete).
- Settings (light/dark/system theme), Help center, and About pages.
- Fixed: Cloud Word's Find & Replace close button, and Cloud Sheet's chart
  dialog, could get stuck open. The cause was a CSS rule that set
  `display` unconditionally on the same element the `hidden` attribute
  targeted, so the attribute had no visible effect. `theme.css` now
  carries one rule, `[hidden] { display: none !important; }`, that makes
  `hidden` authoritative everywhere in the suite going forward.
- Cloud Word: expanded font list, working line spacing (previously it
  could silently do nothing if a toolbar dropdown had stolen the
  browser's text selection before the change fired).

## Known limits

- Cloud Sheet's cell merge hides borders and redirects edits to the
  top-left cell rather than using a true HTML colspan/rowspan.
- Copy/paste in Cloud Sheet does not shift relative formula references.
- Row/column insert-delete does not renumber existing merges.
- Cloud Word's Find & Replace matches within a single formatting run; it
  won't find a phrase split across two different bold/italic spans.

## Third-party licenses

Mammoth.js, html-docx-js, SheetJS (xlsx), and PptxGenJS each ship under
their own open-source licenses, included in their minified files' headers
where present. Check each project's repository for full license text if
you need it.

---

Built by Tasmon Islam (tasmon@outlook.com)
