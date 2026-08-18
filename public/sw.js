// Service Worker for POS PWA
const CACHE_NAME = 'pos-v1.2.0';
const urlsToCache = [
  '/pos.html',
  '/css/style.css?v=2',
  '/vendor/bootstrap/css/bootstrap.min.css',
  '/vendor/bootstrap-icons/bootstrap-icons.css',
  '/js/performance-optimizer.js',
  '/js/optimized-loader.js?v=20260209',
  '/js/pos.js?v=15',
  '/js/back-to-top.js?v=2',
  '/vendor/jquery/jquery.min.js',
  '/vendor/bootstrap/js/bootstrap.bundle.min.js',
  '/vendor/jspdf/jspdf.umd.min.js',
  '/vendor/html2canvas/html2canvas.min.js',
  '/vendor/html2pdf/html2pdf.bundle.min.js',
  '/js/vendor/zxing-browser.min.js',
  '/js/vendor/zxing-library.min.js',
  '/Library/browser/jsDelivr/Versi%200.1.5/zxing-browser.min.js',
  '/Library/browser/UNPKG/Versi%200.1.5/zxing-browser.min.js',
  '/Library/browser/jsDelivr/Versi%200.1.4/zxing-browser.min.js',
  '/Library/browser/UNPKG/Versi%200.1.4/zxing-browser.min.js',
  '/Library/browser/jsDelivr/Versi%200.1.1/zxing-browser.min.js',
  '/Library/browser/UNPKG/Versi%200.1.1/zxing-browser.min.js',
  '/Library/browser/unminified/Versi%200.1.5/zxing-browser.js',
  '/Library/browser/unminified/Versi%200.1.4/zxing-browser.js',
  '/Library/library/jsDelivr/Versi%200.21.3/index.min.js',
  '/Library/library/UNPKG/Versi%200.21.3/index.min.js',
  '/Library/library/UNPKG/Versi%200.20.0/index.min.js',
  '/Library/library/UNPKG/Versi%200.19.3/index.min.js',
  '/Library/library/UNPKG/versi%200.12.3/index.min.js',
  '/Library/library/unminified/Versi%200.12.3/index.min.js'
];

const APP_SHELL_PATHS = new Set([
  '/',
  '/kasir',
  '/admin',
  '/login',
  '/pos.html',
  '/admin.html',
  '/index.html',
  '/receipt-print.html',
  '/debt-receipt-print.html',
  '/mutasi-barang.html'
]);

function isAppShellRequest(reqUrl) {
  try {
    const path = String(reqUrl.pathname || '').replace(/\/+$/, '') || '/';
    if (APP_SHELL_PATHS.has(path) || APP_SHELL_PATHS.has(reqUrl.pathname)) return true;
    // Jangan ganggu route HTML autentikasi / halaman utama.
    if (path === '/kasir' || path.startsWith('/kasir/')) return true;
    if (path === '/admin' || path.startsWith('/admin/')) return true;
    if (path === '/login' || path.startsWith('/login/')) return true;
  } catch (e) {}
  return false;
}

function offlineJson(message) {
  return new Response(JSON.stringify({ success: false, offline: true, message: message || 'Network unavailable' }), {
    status: 503,
    headers: { 'Content-Type': 'application/json' }
  });
}

self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(async cache => {
        console.log('Opened cache');
        await Promise.allSettled(
          urlsToCache.map(async (url) => {
            try {
              await cache.add(url);
            } catch (error) {
              console.warn('Cache add failed:', url);
            }
          })
        );
      })
  );
});

self.addEventListener('fetch', event => {
  try {
    const req = event.request;
    if (!req || req.method !== 'GET') {
      return;
    }

    let reqUrl;
    try {
      reqUrl = new URL(req.url);
    } catch (e) {
      return;
    }

    if (reqUrl.origin !== self.location.origin) {
      return;
    }

    const accept = (req.headers && req.headers.get('accept')) || '';
    const isSse = accept.includes('text/event-stream') || reqUrl.pathname.includes('/events');
    const isNavigate = req.mode === 'navigate' || req.destination === 'document';
    const isApi = reqUrl.pathname.startsWith('/api/');

    // Biarkan browser menangani navigasi & halaman app shell (hindari error FetchEvent /kasir).
    if (isNavigate || isAppShellRequest(reqUrl)) {
      return;
    }

    // SSE: network only, jangan Response.error() (bising di console).
    if (isSse) {
      event.respondWith(
        fetch(req).catch(() => new Response(null, { status: 204, statusText: 'No Content' }))
      );
      return;
    }

    // API: network first, fallback JSON 503 (bukan Response.error).
    if (isApi) {
      event.respondWith(
        fetch(req).catch(() => offlineJson('Network unavailable'))
      );
      return;
    }

    // Static assets: cache first, lalu network, lalu cache ulang.
    event.respondWith(
      caches.match(req).then(async (cached) => {
        if (cached) return cached;
        try {
          const networkResponse = await fetch(req);
          return networkResponse;
        } catch (error) {
          console.warn('Network request failed for:', req.url);
          const cachedFallback = await caches.match(req);
          if (cachedFallback) return cachedFallback;
          // Jangan pakai Response.error() — memicu "FetchEvent ... error response object".
          return new Response('Offline', {
            status: 503,
            statusText: 'Service Unavailable',
            headers: { 'Content-Type': 'text/plain; charset=utf-8' }
          });
        }
      })
    );
  } catch (error) {
    console.error('Service worker fetch error:', error);
    // Jangan respondWith(Response.error()); biarkan request jalan normal.
  }
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (cacheName !== CACHE_NAME) {
            console.log('Deleting old cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});
