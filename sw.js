const CACHE = 'haebing-v4';
const STATIC = [
  '/inca-dashboard/icon-192.png',
  '/inca-dashboard/icon-512.png',
  'https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css',
  'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js',
];
// viewer.html, manifest.json 은 캐시 제외 → 항상 최신 버전 로드

// 설치: 정적 자산 캐시
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(STATIC)).then(() => self.skipWaiting())
  );
});

// 활성화: 구버전 캐시 삭제
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// fetch: GAS API는 네트워크 우선, HTML은 항상 네트워크, 나머지는 캐시 우선
self.addEventListener('fetch', e => {
  const url = e.request.url;

  // HTML 파일 → 항상 네트워크 직접 (최신 버전 보장)
  if (url.endsWith('.html') || url.includes('/inca-dashboard/?') || url.endsWith('/inca-dashboard/') || url.endsWith('/inca-dashboard')) {
    e.respondWith(fetch(e.request, { cache: 'no-store' }));
    return;
  }

  // Google Apps Script (데이터 API) → 네트워크 우선, 실패 시 캐시
  if (url.includes('script.google.com')) {
    e.respondWith(
      fetch(e.request)
        .then(res => {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
          return res;
        })
        .catch(() => caches.match(e.request))
    );
    return;
  }

  // 정적 자산 → 캐시 우선, 없으면 네트워크
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(res => {
        if (res && res.status === 200) {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        return res;
      });
    })
  );
});
