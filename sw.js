const CACHE = 'dtm-v1';
const PRECACHE = ['.', 'contacts.js', 'icon.svg', 'manifest.json'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(PRECACHE)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  // Only cache same-origin GET requests for the shell assets; let API/auth calls pass through
  if (e.request.method !== 'GET' || url.origin !== self.location.origin) return;
  if (url.pathname.includes('login') || url.hostname.includes('microsoft') || url.hostname.includes('graph')) return;
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});
