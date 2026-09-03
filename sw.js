// 판매/재고/수급 현황 PWA 서비스워커 — index.html·data.json은 네트워크 우선(실패 시 캐시), 아이콘 등 정적 파일은 캐시 우선
const VERSION = 2;   // 아이콘 교체(2026-09-03)
const CACHE = 'stock-supply-v' + VERSION;
const STATIC = [
  './', './index.html', './manifest.json',
  './icons/icon-192.png', './icons/icon-512.png',
  './icons/icon-maskable-192.png', './icons/icon-maskable-512.png',
  './icons/favicon-32.png', './icons/apple-touch-icon-180.png'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(STATIC)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (url.origin !== location.origin) return;          // 폰트 CDN 등 외부 요청은 관여하지 않음
  const isHTML = e.request.mode === 'navigate' || url.pathname.endsWith('/index.html');
  const isData = url.pathname.endsWith('/data/data.json');
  if (isHTML || isData) {
    const key = isData ? './data/data.json' : './index.html';
    e.respondWith(
      fetch(e.request).then(res => {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(key, copy));
        return res;
      }).catch(() => caches.match(key))
    );
    return;
  }
  e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request)));
});
