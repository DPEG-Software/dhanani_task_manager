const CACHE = 'dtm-v56';
const PRECACHE = [
  '.',
  'bootstrap.js',
  'styles.css',
  'app-state.js',
  'app-helpers.js',
  'auth.js',
  'data-sync.js',
  'settings-admin.js',
  'ai-summary.js',
  'navigation.js',
  'outlook-shell.js',
  'outlook-mail.js',
  'outlook-calendar.js',
  'notifications.js',
  'notification-center.js',
  'dashboard.js',
  'action-log.js',
  'task-details.js',
  'people-departments.js',
  'discussion-notes.js',
  'contacts.js',
  'tasks-hub.js',
  'notes-archive.js',
  'summary-sheet.js',
  'app.js',
  'msal.min.js',
  'icon.svg',
  'manifest.json'
];

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
  if (e.request.method !== 'GET' || url.origin !== self.location.origin) return;
  e.respondWith(fetch(e.request).catch(() => caches.match(e.request)));
});
