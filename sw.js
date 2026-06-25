/* 선교 오프라인 회계 — 앱 껍데기 캐시 (오프라인 로딩용)
   파일을 고치면 아래 CACHE 버전 숫자를 올리세요: v1 → v2 ... */
const CACHE = "igw-mission-v19";
const SHELL = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png"
];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

/* 캐시 우선 → 없으면 네트워크 (오프라인에서도 앱이 뜨도록) */
self.addEventListener("fetch", e => {
  if (e.request.method !== "GET") return;
  e.respondWith(
    caches.match(e.request).then(hit => hit || fetch(e.request).catch(() => caches.match("./index.html")))
  );
});
