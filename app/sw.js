// Interview Recorder service worker: keeps the app usable without a connection.
const CACHE = 'dsint-v0.1.0';
const INDEX = new URL('./index.html', self.registration.scope).href;
const ASSETS = [INDEX, new URL('./', self.registration.scope).href,
  new URL('./manifest.webmanifest', self.registration.scope).href,
  new URL('./icon-192.png', self.registration.scope).href,
  new URL('./icon-512.png', self.registration.scope).href,
  new URL('./apple-touch-icon.png', self.registration.scope).href];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;          // never touch the Google server calls
  // Open pages always map to index.html (this also keeps setup links with a passcode out of the cache)
  const key = req.mode === 'navigate' ? INDEX : req;
  e.respondWith(caches.match(key).then(hit => {
    const net = fetch(key).then(res => {
      if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(key, copy)); }
      return res;
    }).catch(() => hit);
    return hit || net;
  }));
});
