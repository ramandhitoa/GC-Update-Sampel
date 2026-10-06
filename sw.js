const CACHE_PREFIX = 'gc-update-sampel-pwa-';
const CACHE_VERSION = `${CACHE_PREFIX}v1`;
const SHELL_ASSETS = [
  './',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png'
];
const CDN_ASSETS = [
  'https://cdn.tailwindcss.com/',
  'https://unpkg.com/lucide@latest',
  'https://cdn.jsdelivr.net/npm/chart.js'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_VERSION)
      .then(async cache => {
        await cache.addAll(SHELL_ASSETS);
        await Promise.allSettled(CDN_ASSETS.map(async assetUrl => {
          const request = new Request(assetUrl, { mode: 'no-cors' });
          const response = await fetch(request);
          await cache.put(request, response);
        }));
      })
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys
        .filter(key => key.startsWith(CACHE_PREFIX) && key !== CACHE_VERSION)
        .map(key => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const request = event.request;
  if (request.method !== 'GET') return;

  const url = new URL(request.url);
  if (url.origin === self.location.origin) {
    if (request.mode === 'navigate') {
      event.respondWith(networkFirst(request));
      return;
    }
    if (url.pathname.endsWith('/manifest.webmanifest') || url.pathname.includes('/icons/')) {
      event.respondWith(cacheFirst(request));
    }
    return;
  }

  if (CDN_ASSETS.includes(url.href)) {
    event.respondWith(cacheFirst(request));
    return;
  }

  if (url.hostname === 'script.google.com' && url.pathname.includes('/macros/s/')) {
    event.respondWith(networkFirst(request));
  }
});

async function cacheFirst(request) {
  const cached = await caches.match(request);
  if (cached) return cached;

  const response = await fetch(request);
  if (response.ok) {
    try {
      const cache = await caches.open(CACHE_VERSION);
      await cache.put(request, response.clone());
    } catch (error) {
      console.warn('PWA static asset could not be cached:', error);
    }
  }
  return response;
}

async function networkFirst(request) {
  try {
    const response = await fetch(request);
    if (response.ok) {
      try {
        const cache = await caches.open(CACHE_VERSION);
        await cache.put(request, response.clone());
      } catch (error) {
        console.warn('PWA GET response could not be cached:', error);
      }
    }
    return response;
  } catch (error) {
    const cached = await caches.match(request);
    if (cached) return cached;
    throw error;
  }
}
