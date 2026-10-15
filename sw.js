/* Service Worker: App-Huelle offline halten.
   Version bei jeder Aenderung hochzaehlen, sonst behaelt das iPhone die alte. */

const VERSION = 'abrechnung-v4';
const HUELLE = [
  './', './index.html', './style.css', './app.js',
  './manifest.webmanifest', './icon-192.png', './icon-512.png',
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(VERSION).then((c) => c.addAll(HUELLE)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((k) => Promise.all(k.filter((x) => x !== VERSION).map((x) => caches.delete(x))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET') return;

  // Die App spricht ausschliesslich mit der eigenen Herkunft. Alles andere
  // wird nicht angefasst, es gibt keine externen Aufrufe.
  // Eigene Dateien: erst Cache, im Hintergrund auffrischen.
  if (u.origin === self.location.origin) {
    e.respondWith(
      caches.match(e.request).then((treffer) => {
        const netz = fetch(e.request).then((a) => {
          if (a && a.status === 200) {
            const kopie = a.clone();
            caches.open(VERSION).then((c) => c.put(e.request, kopie));
          }
          return a;
        }).catch(() => treffer);
        return treffer || netz;
      })
    );
  }
});
