/* PAVIMAX service worker — habilita PWA + notificaciones desde SW. */
const CACHE = 'pavimax-v8';
const ASSETS = [
  './',
  './index.html',
  './cargar.html',
  './manifest.json',
  './logo.png',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS).catch(() => {})));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.filter(k => k !== CACHE).map(k => caches.delete(k))
    )).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const url = e.request.url;
  // Llamadas al Apps Script y a Drive siempre van directo a la red (datos vivos)
  if (url.indexOf('script.google.com') !== -1
   || url.indexOf('googleusercontent.com') !== -1
   || url.indexOf('drive.google.com') !== -1) {
    return;
  }
  if (e.request.method !== 'GET') return;
  // Network-first para el HTML (así actualizaciones llegan rápido), cache-first para el resto
  if (e.request.mode === 'navigate' || url.endsWith('/index.html') || url.endsWith('/')) {
    e.respondWith(
      fetch(e.request)
        .then(res => {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, copy)).catch(() => {});
          return res;
        })
        .catch(() => caches.match(e.request))
    );
  } else {
    e.respondWith(
      caches.match(e.request).then(cached => cached || fetch(e.request))
    );
  }
});

// Clicks en notificación abren la app
self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(list => {
      for (const c of list) {
        if ('focus' in c) return c.focus();
      }
      if (clients.openWindow) return clients.openWindow('./');
    })
  );
});
