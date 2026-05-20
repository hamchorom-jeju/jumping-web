// Service Worker for Nohyung Jumping PWA (v45.120)
// Strategy: Network First for HTML, Cache First for assets

const CACHE_NAME = 'nohyung-jumping-v45.120';
const ASSETS_TO_CACHE = [
  '/',
  '/index.html',
  '/index.css',
  '/dashboard.js',
  '/therapy-icon.png',
  '/final-logo.png',
  '/club-logo.png',
  '/badge_v2_seed.png',
  '/badge_v2_sprout.png',
  '/badge_v2_tree.png',
  '/badge_v2_flower.png',
  '/badge_v2_fairy.png',
  '/badge_v2_legend.png',
  '/badge_v2_guardian.png'
];

// 🛠️ Install: Cache essential assets
self.addEventListener('install', (event) => {
  self.skipWaiting(); // Force the waiting service worker to become active
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS_TO_CACHE);
    })
  );
});

// 🧹 Activate: Clean up old caches (Crucial for fixing the "Purple Portal" issue)
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            console.log('[SW] Deleting old cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// 🌐 Fetch: Network-First for better update visibility
self.addEventListener('fetch', (event) => {
  // Only handle GET requests
  if (event.request.method !== 'GET') return;

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        // If network request succeeds, update the cache
        const responseClone = response.clone();
        caches.open(CACHE_NAME).then((cache) => {
          cache.put(event.request, responseClone);
        });
        return response;
      })
      .catch(() => {
        // If network fails, try the cache
        return caches.match(event.request);
      })
  );
});
