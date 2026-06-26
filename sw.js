const CACHE_NAME = 'messstellen-v33';
const TILE_CACHE = 'messstellen-tiles-v2';

// Core app shell — always cache these
const APP_SHELL = [
    './',
    './index.html',
    './app.js',
    './formatting.js',
    './styles.css',
    './manifest.json',
    './app-icon.png',
    './idb_logo.jpg'
];

// CDN resources — stale-while-revalidate
const CDN_ASSETS = [
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css',
    'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css',
    'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js',
    'https://unpkg.com/leaflet-draw@1.0.4/dist/leaflet.draw.css',
    'https://unpkg.com/leaflet-draw@1.0.4/dist/leaflet.draw.js',
    'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js',
    'https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
    'https://cdn.jsdelivr.net/npm/chart.js',
    'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap'
];

// Install: cache app shell immediately
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => {
                // Cache app shell (fail gracefully for individual items)
                return Promise.allSettled(
                    APP_SHELL.map(url => cache.add(url).catch(e => console.warn('Cache miss:', url, e)))
                ).then(() => {
                    // Try to cache CDN assets (non-blocking)
                    return Promise.allSettled(
                        CDN_ASSETS.map(url => cache.add(url).catch(e => console.warn('CDN cache miss:', url, e)))
                    );
                });
            })
            .then(() => self.skipWaiting())
    );
});

// Activate: clean old caches, claim clients
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then(keys =>
            Promise.all(
                keys.filter(k => k !== CACHE_NAME && k !== TILE_CACHE)
                    .map(k => caches.delete(k))
            )
        ).then(() => self.clients.claim())
    );
});

// Fetch strategy based on request type
self.addEventListener('fetch', (event) => {
    const url = new URL(event.request.url);

    // Skip non-GET requests
    if (event.request.method !== 'GET') return;

    // 1. Network-first for geocoding API
    if (url.hostname === 'nominatim.openstreetmap.org') {
        event.respondWith(
            fetch(event.request)
                .then(response => {
                    // Cache successful geocoding results
                    if (response.ok) {
                        const clone = response.clone();
                        caches.open(CACHE_NAME).then(c => c.put(event.request, clone));
                    }
                    return response;
                })
                .catch(() => caches.match(event.request))
        );
        return;
    }

    // 2. Stale-while-revalidate for CDN resources
    if (url.hostname.includes('cdnjs.cloudflare.com') ||
        url.hostname.includes('unpkg.com') ||
        url.hostname.includes('cdn.sheetjs.com') ||
        url.hostname.includes('cdn.jsdelivr.net') ||
        url.hostname.includes('fonts.googleapis.com') ||
        url.hostname.includes('fonts.gstatic.com')) {
        event.respondWith(
            caches.match(event.request).then(cached => {
                const fetchPromise = fetch(event.request).then(response => {
                    if (response.ok) {
                        const clone = response.clone();
                        caches.open(CACHE_NAME).then(c => c.put(event.request, clone));
                    }
                    return response;
                }).catch(() => cached);

                return cached || fetchPromise;
            })
        );
        return;
    }

    // 3. Cache-first for map tiles (with dedicated tile cache + size limit)
    if (url.hostname.includes('google.com') && url.pathname.includes('/vt/')) {
        event.respondWith(
            caches.match(event.request).then(cached => {
                if (cached) return cached;
                return fetch(event.request).then(response => {
                    if (response.ok) {
                        const clone = response.clone();
                        caches.open(TILE_CACHE).then(cache => {
                            cache.put(event.request, clone);
                            // Limit tile cache to ~500 entries
                            cache.keys().then(keys => {
                                if (keys.length > 500) {
                                    // Remove oldest 100 tiles
                                    keys.slice(0, 100).forEach(key => cache.delete(key));
                                }
                            });
                        });
                    }
                    return response;
                }).catch(() => {
                    // Return a transparent 1x1 png for missing tiles
                    return new Response('', { status: 404 });
                });
            })
        );
        return;
    }

    // 4. Network-first for app shell, falling back to cache
    if (url.origin === self.location.origin) {
        event.respondWith(
            fetch(event.request)
                .then(response => {
                    if (response.ok) {
                        const clone = response.clone();
                        caches.open(CACHE_NAME).then(c => c.put(event.request, clone));
                    }
                    return response;
                })
                .catch(() => caches.match(event.request).then(cached => {
                    if (cached) return cached;
                    if (event.request.mode === 'navigate') {
                        return caches.match('./index.html');
                    }
                }))
        );
        return;
    }
});

self.addEventListener('message', (event) => {
    if (event.data && event.data.action === 'skipWaiting') {
        self.skipWaiting();
    }
});
