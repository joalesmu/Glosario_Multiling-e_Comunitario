/* Visor del catálogo · service worker: guarda el visor en el aparato para que abra sin internet.
   Copyright (c) 2026 Alejandro Estrada · SPDX-License-Identifier: AGPL-3.0-only
   Solo guarda los archivos del visor. Los datos del catálogo, las imágenes y los audios los guarda el propio visor
   (copia del catálogo y «📥 Descargar para usar sin conexión»); las consultas al catálogo pasan directo, sin tocarse.
   Al publicar una versión nueva, cambia VERSION: el visor avisa «Hay una versión nueva». */
const VERSION = 'visor-1.2';
const ARCHIVOS = ['./', './index.html', './manifest.webmanifest', './iconos/icono-192.png', './iconos/icono-512.png', './iconos/icono-maskable-512.png', './iconos/apple-touch-icon.png', './iconos/favicon-32.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(ARCHIVOS.map(u => new Request(u, { cache: 'reload' })))));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('visor-') && k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
// Solo los archivos del propio visor: primero lo guardado (abre al instante aunque la señal sea mala)
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then(r => r || fetch(e.request).catch(() => caches.match('./index.html'))));
});
self.addEventListener('message', e => { if (e.data === 'saltar') self.skipWaiting(); });
