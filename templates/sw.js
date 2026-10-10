// Caches only static assets and the offline page. Pages with financial data are never cached.
const CACHE = "kucu-v1";
const ASSETS = ["/offline/", "{{ css }}", "{{ logo }}", "{{ icon }}"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const r = e.request;
  if (r.method !== "GET") return;
  const url = new URL(r.url);
  if (url.origin !== location.origin) return;
  if (url.pathname.startsWith("/static/")) {
    e.respondWith(fetch(r).then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put(r, copy)); return res; }).catch(() => caches.match(r)));
  } else if (r.mode === "navigate") {
    e.respondWith(fetch(r).catch(() => caches.match("/offline/")));
  }
});
