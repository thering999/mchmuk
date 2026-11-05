// Bump this string on every deploy to invalidate old caches across all clients.
const CACHE_NAME = 'mchmuk-v3-20261105';
const DATASET_URL = 'https://raw.githubusercontent.com/thering999/mchmuk/main/tmp_exchange_data.xlsx';
// Must match IDB constants in app.js
const IDB_NAME = 'mchmuk-offline';
const IDB_STORE = 'datasets';
const IDB_KEY = 'mchmuk-xlsx-buffer';
const STATIC_ASSETS = [
    './',
    './index.html',
    './app.js',
    './style.css',
    './auth.js',
    './hospitals_db.js',
    './excel-worker.js',
    './manifest.json',
    'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
    'https://cdn.jsdelivr.net/npm/apexcharts',
    'https://unpkg.com/lucide@latest',
];

self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME).then(cache => cache.addAll(STATIC_ASSETS).catch(() => {}))
    );
    self.skipWaiting();
});

self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys().then(keys => Promise.all(
            keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
        ))
    );
    self.clients.claim();
});

function idbPut(value) {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(IDB_NAME, 1);
        req.onupgradeneeded = e => {
            const db = e.target.result;
            if (!db.objectStoreNames.contains(IDB_STORE)) db.createObjectStore(IDB_STORE);
        };
        req.onerror = e => reject(e.target.error);
        req.onsuccess = e => {
            const db = e.target.result;
            const tx = db.transaction(IDB_STORE, 'readwrite');
            tx.objectStore(IDB_STORE).put(value, IDB_KEY);
            tx.oncomplete = () => { db.close(); resolve(); };
            tx.onerror = ev => { db.close(); reject(ev.target.error); };
        };
    });
}

async function refreshOfflineDataset() {
    const res = await fetch(`${DATASET_URL}?cb=${Date.now()}`);
    if (!res.ok) return;
    const buffer = await res.arrayBuffer();
    await idbPut({ buffer, savedAt: Date.now() });
}

self.addEventListener('periodicsync', event => {
    if (event.tag === 'mchmuk-data-refresh') event.waitUntil(refreshOfflineDataset().catch(() => {}));
});

self.addEventListener('fetch', event => {
    const url = new URL(event.request.url);
    // Network-first for GitHub raw data
    if (url.hostname === 'raw.githubusercontent.com' || url.hostname === 'api.github.com') {
        event.respondWith(
            fetch(event.request).catch(() => caches.match(event.request))
        );
        return;
    }
    // Cache-first for static assets
    event.respondWith(
        caches.match(event.request).then(cached => {
            if (cached) return cached;
            return fetch(event.request).then(response => {
                if (response && response.status === 200 && response.type === 'basic') {
                    const clone = response.clone();
                    caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
                }
                return response;
            });
        })
    );
});
