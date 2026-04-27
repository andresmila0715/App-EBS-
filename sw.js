const CACHE_NAME = 'ebs-app-v40';

// Archivos precacheados al instalar — config.js incluido para login offline
const ASSETS = [
  '/',
  '/index.html',
  '/config.js',
  '/manifest.json',
  '/icon-192.png',
  '/icon-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;

  // Google Sheets: siempre por red, nunca cachear
  if (event.request.url.includes('script.google.com')) {
    event.respondWith(
      fetch(event.request).catch(() => new Response('', { status: 503 }))
    );
    return;
  }

  const reqUrl = new URL(event.request.url);
  const isSameOrigin = reqUrl.origin === self.location.origin;
  const isConfig = reqUrl.pathname.endsWith('/config.js');
  const isIndex  = event.request.mode === 'navigate'
                || reqUrl.pathname === '/'
                || reqUrl.pathname.endsWith('/index.html');

  // config.js e index.html: stale-while-revalidate
  // → sirve caché INMEDIATAMENTE (funciona offline en iOS)
  // → actualiza en segundo plano para la próxima apertura
  if (isSameOrigin && (isConfig || isIndex)) {
    event.respondWith(
      caches.open(CACHE_NAME).then(async cache => {
        const cached = await cache.match(isConfig ? '/config.js' : '/index.html');

        // Actualizar en segundo plano sin bloquear
        const fetchPromise = fetch(event.request).then(response => {
          if (response.ok) {
            cache.put(isConfig ? '/config.js' : '/index.html', response.clone());
          }
          return response;
        }).catch(() => null);

        // Si hay caché, devolver inmediatamente (no esperar red)
        if (cached) return cached;

        // Si no hay caché (primera vez), esperar la red
        return fetchPromise.then(r => r || new Response('Sin conexión y sin caché', { status: 503 }));
      })
    );
    return;
  }

  // Resto de recursos: cache-first con respaldo de red
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        if (isSameOrigin && response.ok) {
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, response.clone()));
        }
        return response;
      }).catch(() => caches.match('/index.html'));
    })
  );
});
