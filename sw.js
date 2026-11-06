/* Keeps the app openable without internet. Only this site's own files are cached, never reports. */
const CACHE = "hours-v3";
const FILES = ["./", "index.html", "app.js", "xlsx.js", "manifest.webmanifest", "icon-180.png", "icon-512.png"];

self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", event => {
  event.waitUntil(caches.keys()
    .then(names => Promise.all(names.filter(n => n !== CACHE).map(n => caches.delete(n))))
    .then(() => self.clients.claim()));
});

/* Shows the saved copy at once and refreshes it in the background, so updates appear on the next open. */
self.addEventListener("fetch", event => {
  const url = new URL(event.request.url);
  if (event.request.method !== "GET" || url.origin !== self.location.origin) return;
  event.respondWith(caches.open(CACHE).then(async cache => {
    const saved = await cache.match(event.request, { ignoreSearch: true });
    const fresh = fetch(event.request).then(response => {
      if (response.ok) cache.put(event.request, response.clone());
      return response;
    });
    if (saved) { fresh.catch(() => {}); return saved; }
    return fresh;
  }));
});
