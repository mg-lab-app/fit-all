// FIT-ALL: funciona sin conexión. Cambia VERSION al publicar una versión nueva.
const VERSION = "fitall-v5";
const BASE = ["./", "./index.html", "./manifest.webmanifest", "./icons/icon-192.png", "./icons/icon-512.png", "./icons/apple-touch-icon.png", "./icons/favicon.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(BASE)));
  self.skipWaiting();
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", e => {
  const r = e.request;
  if (r.method !== "GET") return;
  // La app: primero internet (para tener siempre la última versión); sin conexión, la guardada.
  if (r.mode === "navigate") {
    e.respondWith(fetch(r)
      .then(res => { const c = res.clone(); caches.open(VERSION).then(ca => ca.put("./index.html", c)); return res; })
      .catch(() => caches.match("./index.html")));
    return;
  }
  // Iconos, fuentes y demás: la copia guardada al momento y se actualiza por detrás.
  e.respondWith(caches.match(r).then(hit => {
    const red = fetch(r).then(res => {
      if (res && (res.ok || res.type === "opaque")) { const c = res.clone(); caches.open(VERSION).then(ca => ca.put(r, c)); }
      return res;
    }).catch(() => hit);
    return hit || red;
  }));
});
