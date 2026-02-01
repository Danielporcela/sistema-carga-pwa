const CACHE_NAME = "app-frota-v1";

// Arquivos essenciais (mínimo para NÃO travar)
const FILES_TO_CACHE = [
  "/",
  "/index.html",
  "/manifest.json"
];

// ===== INSTALAÇÃO =====
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(FILES_TO_CACHE);
    })
  );
  self.skipWaiting();
});

// ===== ATIVAÇÃO =====
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cache) => {
          if (cache !== CACHE_NAME) {
            return caches.delete(cache);
          }
        })
      );
    })
  );
  self.clients.claim();
});

// ===== FETCH (SEM LOOP) =====
self.addEventListener("fetch", (event) => {
  // NÃO intercepta requisições POST (Flask precisa disso)
  if (event.request.method !== "GET") {
    return;
  }

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        // Se resposta válida, retorna direto
        return response;
      })
      .catch(() => {
        // Se offline, tenta cache
        return caches.match(event.request);
      })
  );
});



