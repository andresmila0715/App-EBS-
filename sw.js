const CACHE_NAME = 'ebs-app-v14';

// Archivos que se guardan para funcionar sin internet
const ASSETS = [
  '/',
  '/index.html',
  '/manifest.json',
  '/icon-192.png',
  '/icon-512.png'
];

// Al instalar: guarda todos los archivos en caché
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(ASSETS);
    })
  );
  self.skipWaiting();
});

// Al activar: elimina cachés viejos
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      )
    )
  );
  self.clients.claim();
});

// Al hacer fetch:
// - Navegación e index.html: network-first (mejor para actualizaciones)
// - Estáticos: cache-first
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;

  // Las llamadas a Google Sheets siempre van por red
  if (event.request.url.includes('script.google.com')) {
    event.respondWith(fetch(event.request).catch(() => new Response('', { status: 503 })));
    return;
  }

  const reqUrl = new URL(event.request.url);
  const isSameOrigin = reqUrl.origin === self.location.origin;
  const isNavRequest = event.request.mode === 'navigate';
  const isIndexRequest = reqUrl.pathname === '/' || reqUrl.pathname.endsWith('/index.html');
  const isConfigRequest = reqUrl.pathname.endsWith('/config.js');

  // HTML principal y config.js: network-first para reflejar cambios de usuarios/municipios
  if (isSameOrigin && (isNavRequest || isIndexRequest || isConfigRequest)) {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          const clone = response.clone();
          const key = isConfigRequest ? '/config.js' : '/index.html';
          caches.open(CACHE_NAME).then(cache => cache.put(key, clone));
          return response;
        })
        .catch(async () => {
          const key = isConfigRequest ? '/config.js' : '/index.html';
          const cached = await caches.match(key);
          return cached || new Response('Sin conexion y sin cache disponible', { status: 503 });
        })
    );
    return;
  }

  // Demas recursos: cache-first con respaldo de red.
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        // Guarda respuestas del mismo origen para offline.
        const clone = response.clone();
        if (isSameOrigin) {
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => caches.match('/index.html'));
    })
  );
});
