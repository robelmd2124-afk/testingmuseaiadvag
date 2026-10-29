// AdVerse Agency Service Worker — installability + safe offline fallback.
//
// Why this exists: Chrome only offers "Install app" when a service worker
// with a fetch handler controls the page, and the manifest icons resolve.
// The previous worker unregistered itself, so install never fired.
//
// Update rule: bump SW_VERSION whenever the precache list, icons or the
// manifest change. /sw.js is served no-cache by server.ts, so browsers
// pick the new worker up on the next visit and old caches are purged.
//
// Freshness rule (learned from deploy desyncs): JS/CSS/HTML are NEVER
// cached here — they are always network-first. Only brand images, fonts
// and the manifest use cache-first, and /api/* (sheet RPC) is never
// cached at all, so a save can never be served stale.
const SW_VERSION = 'adverse-pwa-v1';
const STATIC_CACHE = `adverse-static-${SW_VERSION}`;
const PAGE_CACHE = `adverse-pages-${SW_VERSION}`;
const STATIC_RE = /\.(png|svg|ico|webmanifest|json|woff2?|ttf)$/i;

self.addEventListener('install', (event) => {
  // Take over as soon as possible; old caches are purged on activate.
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(
        keys.filter((k) => k !== STATIC_CACHE && k !== PAGE_CACHE).map((k) => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // Sheet RPC must never be cached or replayed.
  if (url.pathname.startsWith('/api/')) return;

  // Navigation: network-first so a new deploy is never hidden behind
  // an old shell; fall back to the last cached page when offline.
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(PAGE_CACHE).then((cache) => cache.put(req, copy));
          return res;
        })
        .catch(() => caches.match(req).then((hit) => hit || caches.match('/')))
    );
    return;
  }

  // Same-origin brand assets: cache-first with background refresh.
  if (url.origin === self.location.origin && STATIC_RE.test(url.pathname)) {
    event.respondWith(
      caches.open(STATIC_CACHE).then((cache) =>
        cache.match(req).then((hit) => {
          const network = fetch(req)
            .then((res) => {
              if (res.ok) cache.put(req, res.clone());
              return res;
            })
            .catch(() => hit);
          return hit || network;
        })
      )
    );
    return;
  }

  // Everything else (app JS/CSS/HTML): network only, never cached here.
});
