// NTS Field Acceptance — Service Worker
// Strategia: Network First pentru HTML/JS/CSS (mereu versiunea noua)
//            Cache First pentru fonturi externe (stabile)
// Versiunea se schimba automat la fiecare deploy nou prin timestamp

const SW_VERSION   = '__SW_VERSION__'; // inlocuit la build
const CACHE_NAME   = 'nts-field-' + SW_VERSION;
const FONTS_CACHE  = 'nts-fonts-v1';

// La install — precacheaza nimic, lasa Network First sa faca treaba
self.addEventListener('install', event => {
  console.log('[SW] Install version:', SW_VERSION);
  // skipWaiting = activare imediata fara sa astepte tab-urile vechi
  self.skipWaiting();
});

// La activate — sterge TOATE cache-urile vechi
self.addEventListener('activate', event => {
  console.log('[SW] Activate — clearing old caches');
  event.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys
          .filter(k => k !== CACHE_NAME && k !== FONTS_CACHE)
          .map(k => {
            console.log('[SW] Deleting old cache:', k);
            return caches.delete(k);
          })
      );
    }).then(() => {
      // Preia controlul imediat peste toate tab-urile deschise
      return self.clients.claim();
    })
  );
});

self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // ── Google Fonts — Cache First (nu se schimba) ──
  if (url.hostname === 'fonts.googleapis.com' ||
      url.hostname === 'fonts.gstatic.com') {
    event.respondWith(
      caches.open(FONTS_CACHE).then(cache =>
        cache.match(event.request).then(cached => {
          if (cached) return cached;
          return fetch(event.request).then(resp => {
            cache.put(event.request, resp.clone());
            return resp;
          });
        })
      )
    );
    return;
  }

  // ── GAS API calls — Network Only (niciodata din cache) ──
  if (url.hostname.includes('script.google.com') ||
      url.hostname.includes('googleapis.com')) {
    event.respondWith(
      fetch(event.request).catch(() =>
        new Response(JSON.stringify({ error: 'offline' }), {
          headers: { 'Content-Type': 'application/json' }
        })
      )
    );
    return;
  }

  // ── HTML / JS / CSS / Imagini — Network First ──
  // Incearca intotdeauna reteaua, fallback pe cache doar daca offline
  event.respondWith(
    fetch(event.request, { cache: 'no-store' })
      .then(response => {
        // Salveaza in cache versiunea noua
        if (response && response.status === 200 && event.request.method === 'GET') {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      })
      .catch(() => {
        // Offline — intoarce din cache daca exista
        return caches.match(event.request).then(cached => {
          if (cached) return cached;
          // Fallback pentru navigare offline
          if (event.request.destination === 'document') {
            return caches.match('./index.html');
          }
          return new Response('Offline', { status: 503 });
        });
      })
  );
});

// Mesaj de la pagina pentru a forta refresh
self.addEventListener('message', event => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
