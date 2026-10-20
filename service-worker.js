// Семейный бюджет — офлайн-кэш. Страница: сначала сеть (всегда свежая версия), офлайн — из кэша.
// Остальные файлы (иконки, манифест): сначала кэш. Поднимай версию при обновлении файлов.
const CACHE = "budget-v7";
const ASSETS = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./icon-192.png",
  "./icon-512.png",
  "./icon-180.png",
  "./icon-maskable.png"
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return; // чужие запросы (Apps Script) не трогаем

  const isPage = req.mode === "navigate" || url.pathname.endsWith("/") || url.pathname.endsWith("/index.html");
  if (isPage) {
    e.respondWith(
      fetch(req).then((resp) => {
        const copy = resp.clone();
        caches.open(CACHE).then((c) => c.put("./index.html", copy)).catch(() => {});
        return resp;
      }).catch(() => caches.match("./index.html"))
    );
    return;
  }

  e.respondWith(
    caches.match(req).then((hit) => hit || fetch(req).then((resp) => {
      const copy = resp.clone();
      caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
      return resp;
    }).catch(() => caches.match("./index.html")))
  );
});
