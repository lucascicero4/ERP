const CACHE_NAME = 'erp-lucas-v5.1';
const ASSETS = [
    '/',
    '/index.html',
    '/manifest.json',
    '/icon-192.png',
    '/icon-512.png'
];

// Install
self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => cache.addAll(ASSETS))
            .then(() => self.skipWaiting())
    );
});

// Activate
self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys().then(keys => {
            return Promise.all(
                keys.filter(key => key !== CACHE_NAME)
                    .map(key => caches.delete(key))
            );
        }).then(() => self.clients.claim())
    );
});

// Fetch - Network first, then cache
self.addEventListener('fetch', event => {
    // Skip non-GET requests
    if (event.request.method !== 'GET') return;

    // Skip external API calls (cotizaciones)
    if (event.request.url.includes('api.') || 
        event.request.url.includes('dolarapi') ||
        event.request.url.includes('coingecko')) {
        return;
    }

    event.respondWith(
        fetch(event.request)
            .then(response => {
                // Clone the response
                const responseClone = response.clone();
                
                // Cache the response
                caches.open(CACHE_NAME).then(cache => {
                    cache.put(event.request, responseClone);
                });

                return response;
            })
            .catch(() => {
                // If network fails, try cache
                return caches.match(event.request)
                    .then(response => response || caches.match('/index.html'));
            })
    );
});

// Background sync for offline data
self.addEventListener('sync', event => {
    if (event.tag === 'sync-expenses') {
        event.waitUntil(syncExpenses());
    }
});

async function syncExpenses() {
    // Get pending expenses from IndexedDB and sync
    // This would be implemented with the actual sync logic
    console.log('Syncing expenses...');
}
