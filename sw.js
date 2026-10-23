// Web version: lets the POS page open with no internet. Online it always loads the newest page and keeps a copy,
// offline it serves that copy. Data goes to Google with POST requests, which this never touches.
const CACHE = "pos-page";
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;
  e.respondWith(
    fetch(e.request, {cache: "no-cache"})
      .then(r => {
        if (r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(url.pathname, copy)); }
        return r;
      })
      .catch(() => caches.match(url.pathname).then(r => r || Response.error()))
  );
});
