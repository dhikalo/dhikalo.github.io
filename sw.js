const CACHE_NAME = 'messstellen-v6';
const ASSETS = [
    './',
    './index.html',
    './app.js',
    './formatting.js',
    './styles.css',
    './manifest.json',
    './app-icon.png',
    './idb_logo.jpg',
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css',
    'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css',
    'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js',
    'https://unpkg.com/leaflet-draw@1.0.4/dist/leaflet.draw.css',
    'https://unpkg.com/leaflet-draw@1.0.4/dist/leaflet.draw.js',
    'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js',
    'https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
    'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap'
];

// Install: cache all assets
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => cache.addAll(ASSETS))
            .then(() => self.skipWaiting())
    );
});

// Activate: clean old caches
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then(keys =>
            Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
        ).then(() => self.clients.claim())
    );
});

// Fetch: cache-first for assets, network-first for API calls
self.addEventListener('fetch', (event) => {
    const url = new URL(event.request.url);

    // Network-first for geocoding API
    if (url.hostname === 'nominatim.openstreetmap.org') {
        event.respondWith(
            fetch(event.request).catch(() => caches.match(event.request))
        );
        return;
    }

    // Cache-first for everything else
    event.respondWith(
        caches.match(event.request).then(cached => {
            if (cached) return cached;
            return fetch(event.request).then(response => {
                // Cache tile requests for offline map
                if (url.hostname.includes('google.com') && response.ok) {
                    const clone = response.clone();
                    caches.open(CACHE_NAME + '-tiles').then(cache => cache.put(event.request, clone));
                }
                return response;
            }).catch(() => {
                // Return offline fallback for navigation
                if (event.request.mode === 'navigate') {
                    return caches.match('./index.html');
                }
            });
        })
    );
});
