const CACHE = 'opic-v2';
const FILES = ['./index.html', './manifest.json'];

self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.filter(k => k !== CACHE).map(k => caches.delete(k))
    )).then(() => self.clients.claim())
  );
});

// 네트워크 우선: 온라인이면 항상 최신 파일을 받아오고, 오프라인일 때만 캐시로 대체한다.
// (예전 버전은 캐시가 있으면 무조건 캐시만 써서, 서버를 업데이트해도 설치된 앱에는 영영 반영되지 않았음)
self.addEventListener('fetch', e => {
  var url = new URL(e.request.url);
  // 구글 시트(Apps Script) 같은 다른 도메인 요청은 건드리지 않고 그대로 통과시킨다.
  // 서비스워커가 리다이렉트를 끼고 있는 크로스오리진 요청을 가로채면 브라우저가
  // 응답을 스크립트로 실행하길 거부하는 문제가 있었음 (별점 저장 JSONP 요청이 계속 실패한 원인).
  if (url.origin !== self.location.origin) return;
  e.respondWith(
    fetch(e.request).then(res => {
      var copy = res.clone();
      caches.open(CACHE).then(c => c.put(e.request, copy));
      return res;
    }).catch(() => caches.match(e.request))
  );
});
