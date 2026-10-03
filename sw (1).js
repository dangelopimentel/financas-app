const CACHE = "fin-v1";
const ARQUIVOS = ["./", "index.html", "manifest.json", "icon-192.png", "icon-512.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARQUIVOS)));
  self.skipWaiting();
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (u.origin !== location.origin) return; // os dados financeiros nunca passam pelo cache
  e.respondWith(fetch(e.request).catch(() => caches.match(e.request)));
});
