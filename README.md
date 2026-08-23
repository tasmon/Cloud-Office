# Cloud Office

Three private, offline-capable office apps that run entirely in the browser:

- **Cloud Word** — rich-text document editor. Opens/saves `.docx`, `.html`, `.txt`.
- **Cloud Sheet** — spreadsheet with formulas (`SUM`, `AVERAGE`, `IF`, ranges, etc). Opens/saves `.xlsx`, `.csv`.
- **Cloud PowerPoint** — slide editor with shapes, images and a present mode. Exports `.pptx`, prints to PDF.

Everything is stored **only in the visitor's browser** (IndexedDB). There is no
backend, no account, and no data ever leaves the device — which is also what
makes this trivial to host for free.

## Project structure

```
cloud-office/
  index.html         ← hub / landing page (About section with your name + email)
  shared/             ← shared CSS tokens + tiny IndexedDB helper, used by all 3 apps
  word/               ← Cloud Word (index.html, word.css, word.js, manifest.json, sw.js)
  sheet/              ← Cloud Sheet
  slides/              ← Cloud PowerPoint
  icons/              ← generated PWA icons + favicons
```

Each app is a self-contained folder with its own `manifest.json` and service
worker (`sw.js`), so each one installs as its **own** PWA (separate icon,
separate offline cache) even though they share the hub page and the `shared/`
CSS + storage helper.

## Run it locally

No build step — it's static HTML/CSS/JS. Any static file server works:

```bash
cd cloud-office
python3 -m http.server 8080
# open http://localhost:8080
```

(Opening `index.html` directly via `file://` mostly works too, but IndexedDB
and service workers are more reliable over `http://localhost`.)

## Deploy to GitHub Pages

1. Push this folder to a GitHub repo (the contents of `cloud-office/`, not the
   folder itself, should be at the repo root — or set Pages to serve from a
   subfolder).
2. Repo → **Settings → Pages** → Source: **Deploy from a branch** → pick
   `main` and `/ (root)`.
3. Your site will be live at `https://<username>.github.io/<repo>/`.

## Deploy to Vercel

1. Push the folder to a GitHub repo (or run `vercel` from inside
   `cloud-office/` with the Vercel CLI).
2. In the Vercel dashboard, **Add New Project** → import the repo.
3. Framework preset: **Other** (it's static — no build command, no output
   directory override needed since everything is already static HTML).
4. Deploy. Vercel serves it instantly over HTTPS (required for service
   workers/PWA install to work).

Both hosts serve everything over HTTPS by default, which is required for the
"Install app" prompt and for the service workers to register.

## Notes on the third-party libraries used

Each app pulls in one small, well-known library from a public CDN purely to
read/write real Office file formats — none of them ever send your document
anywhere; they run fully client-side:

- **Cloud Word** — [mammoth.js](https://github.com/mwilliamson/mammoth.js) (.docx → HTML) and [html-docx-js](https://github.com/evidenceprime/html-docx-js) (HTML → .docx)
- **Cloud Sheet** — [SheetJS / xlsx](https://github.com/SheetJS/sheetjs) (.xlsx/.csv read+write)
- **Cloud PowerPoint** — [PptxGenJS](https://github.com/gitbrent/PptxGenJS) (.pptx generation)

If you'd rather have zero external requests (fully offline from first load),
download these libraries into each app folder and change the `<script src>`
tags to point at the local copies — everything else already works offline via
the service workers.

## Customizing

- Colors/fonts: `shared/theme.css` (CSS custom properties at the top).
- Rename the apps: update the `<title>`, `manifest.json` `name`/`short_name`,
  and the on-page headings.
- About section (name + email) lives at the bottom of `index.html`.
