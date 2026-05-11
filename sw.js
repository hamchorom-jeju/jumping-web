// Service Worker for Nohyung Jumping PWA
// This is a minimal service worker to enable "Add to Home Screen" prompts.

const CACHE_NAME = 'nohyung-jumping-v1';
const ASSETS_TO_CACHE = [
  '/',
  '/index.html',
  '/final-logo.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS_TO_CACHE);
    })
  );
});

self.addEventListener('fetch', (event) => {
  event.respondWith(
    caches.match(event.request).then((response) => {
      return response || fetch(event.request);
    })
  );
});
