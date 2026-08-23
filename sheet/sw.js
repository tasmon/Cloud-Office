const CACHE = 'cloud-sheet-v1';
const SHELL = [
  './index.html',
  './sheet.css',
  './sheet.js',
  './manifest.json',
  '../shared/theme.css',
  '../shared/db.js',
  '../icons/sheet-192.png',
  '../icons/sheet-512.png',
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
        .then((res) => { if (res && res.status === 200) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(e.request, copy)); } return res; })
        .catch(() => cached);
      return cached || network;
    })
  );
});
