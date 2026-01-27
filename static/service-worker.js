self.addEventListener("install", event => {
  event.waitUntil(
    caches.open("carga-cache").then(cache => {
      return cache.addAll([
        "/",
        "/static/manifest.json"
      ]);
    })
  );
});

self.addEventListener("fetch", event => {
  event.respondWith(
    caches.match(event.request).then(resp => {
      return resp || fetch(event.request);
    })
  );
});
