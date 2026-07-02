// Minimal service worker — enables PWA installability, no aggressive caching
// (the app needs live server data, so we pass all requests through)

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));

self.addEventListener('fetch', event => {
  // Pass through to network — no offline caching for this app
  event.respondWith(fetch(event.request));
});
