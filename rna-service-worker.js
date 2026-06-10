const CACHE_NAME = "rna-form-offline-v25";
const CORE_SHELL = [
  "./",
  "./index.html",
  "./rna-form.html",
  "./rna-manifest.webmanifest",
  "./rna-icon.svg"
];
const OPTIONAL_SHELL = [
  "./psgc-offline-data.json"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => Promise.all([
        cache.addAll(CORE_SHELL),
        Promise.allSettled(OPTIONAL_SHELL.map((url) => cache.add(url)))
      ]))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;

  const requestUrl = new URL(event.request.url);
  if (requestUrl.origin !== self.location.origin) return;

  const isHtmlShell = requestUrl.pathname.endsWith("/") || requestUrl.pathname.endsWith("/index.html") || requestUrl.pathname.endsWith("/rna-form.html");
  if (isHtmlShell) {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
          return response;
        })
        .catch(() => caches.match(event.request).then((cached) => cached || caches.match("./index.html")))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request)
        .then((response) => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
          if (requestUrl.pathname.endsWith("/psgc-offline-data.json")) {
            caches.open("rna-psgc-offline-data-v1").then((cache) => cache.put(event.request, response.clone()));
          }
          return response;
        })
        .catch(() => caches.match("./index.html"));
    })
  );
});
