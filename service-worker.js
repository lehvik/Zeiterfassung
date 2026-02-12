const CACHE_NAME = 'arbeitszeit-tracker-v4';
const urlsToCache = [
  './timetracker.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

// Install Service Worker
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        console.log('✅ Cache geöffnet');
        return cache.addAll(urlsToCache);
      })
  );
  self.skipWaiting();
});

// Activate Service Worker
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            console.log('🗑️ Alter Cache gelöscht:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
  self.clients.claim();
});

// Fetch Strategy: Network First, fallback to Cache
self.addEventListener('fetch', (event) => {
  // Skip cross-origin requests
  if (!event.request.url.startsWith(self.location.origin)) {
    // Allow Google Sheets API calls to pass through
    if (event.request.url.includes('googleapis.com')) {
      return;
    }
  }

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        // Clone the response
        const responseToCache = response.clone();
        
        caches.open(CACHE_NAME)
          .then((cache) => {
            cache.put(event.request, responseToCache);
          });
        
        return response;
      })
      .catch(() => {
        // If network fails, try cache
        return caches.match(event.request)
          .then((response) => {
            if (response) {
              return response;
            }
            
            // If not in cache either, return offline page
            if (event.request.destination === 'document') {
              return caches.match('./timetracker.html');
            }
          });
      })
  );
});

// Background Sync (for future offline tracking)
self.addEventListener('sync', (event) => {
  if (event.tag === 'sync-timetracking') {
    event.waitUntil(syncTimeData());
  }
});

async function syncTimeData() {
  // Placeholder for future offline sync functionality
  console.log('🔄 Synchronisiere Zeitdaten...');
}
