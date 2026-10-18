const CACHE_NAME = 'hyrox-pwa-v2';
const DATA_CACHE = 'hyrox-data-v1';
const ASSETS = [
  'index.html',
  'manifest.json',
  'icon-192.png',
  'icon-512.png',
  'https://cdn.tailwindcss.com'
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) =>
      // Add files one by one so a single missing file (e.g. an icon) doesn't break the install
      Promise.all(ASSETS.map((a) => cache.add(a).catch((err) => console.warn('SW: could not cache', a, err))))
    )
  );
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(
        keys.filter((k) => k !== CACHE_NAME && k !== DATA_CACHE).map((k) => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // 1) CSV data from GitHub: network first, fall back to the last copy when offline.
  //    The page adds ?t=<timestamp> to every request, so cache by origin + path only.
  if (url.hostname === 'raw.githubusercontent.com') {
    const key = url.origin + url.pathname;
    e.respondWith(
      fetch(req)
        .then((res) => {
          if (res.ok) {
            const copy = res.clone();
            caches.open(DATA_CACHE).then((c) => c.put(key, copy));
          }
          return res;
        })
        .catch(() => caches.match(key).then((hit) => hit || Response.error()))
    );
    return;
  }

  // 2) Our own files (index.html etc.): network first so updates show up immediately,
  //    cached copy when offline.
  if (url.origin === self.location.origin) {
    e.respondWith(
      fetch(req)
        .then((res) => {
          if (res.ok) {
            const copy = res.clone();
            caches.open(CACHE_NAME).then((c) => c.put(req, copy));
          }
          return res;
        })
        .catch(() => caches.match(req).then((hit) => hit || caches.match('index.html')))
    );
    return;
  }

  // 3) Everything else (Tailwind CDN): cache first.
  e.respondWith(caches.match(req).then((hit) => hit || fetch(req)));
});
