// Minimal service worker for PWA "Add to Home Screen" support
const CACHE_NAME = 'match-subs-v1';

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll([
        'index.html',
        'manifest.json'
      ]);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  // Don't cache POST or GET API requests
  if (event.request.method === 'POST') return;
  if (event.request.url.includes('script.google.com')) return;

  // Network-first for the app shell, so a new deploy shows up on the next
  // load instead of serving a stale index.html until the cache is bumped.
  // Falls back to cache when offline.
  const isShell = event.request.mode === 'navigate' ||
                  event.request.url.endsWith('index.html') ||
                  event.request.url.endsWith('/');

  if (isShell) {
    event.respondWith(
      fetch(event.request)
        .then((resp) => {
          const copy = resp.clone();
          caches.open(CACHE_NAME).then((c) => c.put(event.request, copy));
          return resp;
        })
        .catch(() => caches.match(event.request).then((c) => c || caches.match('index.html')))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => cached || fetch(event.request))
  );
});
