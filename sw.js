const CACHE = 'cloud-office-v2';
const SHELL = [
  './index.html', './word.html', './sheet.html', './slides.html', './help.html', './about.html',
  './theme.css', './common.css', './word.css', './sheet.css', './slides.css',
  './common.js', './db.js', './undo.js', './word.js', './sheet.js', './slides.js',
  './manifest.json',
  './mammoth.browser.min.js', './html-docx.js', './xlsx.full.min.js', './pptxgen.bundle.js',
  './icon-hub-192.png', './icon-hub-512.png',
  './icon-word-192.png', './icon-word-512.png',
  './icon-sheet-192.png', './icon-sheet-512.png',
  './icon-slides-192.png', './icon-slides-512.png',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))));
  self.clients.claim();
});

self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request).then((cached) => {
      const network = fetch(e.request)
        .then((res) => {
          if (res && res.status === 200) {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(e.request, copy));
          }
          return res;
        })
        .catch(() => cached);
      return cached || network;
    })
  );
});
