// Offline support: cache app files, serve from cache, update in background.
const CACHE = 'bihar-heritage-v2';
const FILES = ['./', './index.html', './data.js', './data2.js', './credits.js', './manifest.json', './icon-192.png', './icon-512.png',
  './img/aryabhata.jpg', './img/ashoka.jpg', './img/barabar.jpg', './img/bismillah.jpg', './img/bodhgaya.jpg', './img/champaran.jpg', './img/chanakya.jpg', './img/chandragupta.jpg', './img/chhath.jpg', './img/dinkar.jpg', './img/gobind.jpg', './img/golghar.jpg', './img/jp.jpg', './img/kesariya.jpg', './img/kunwar.jpg', './img/litchi.jpg', './img/litti.jpg', './img/madhubani.jpg', './img/mahavira.jpg', './img/makhana.jpg', './img/manjusha.jpg', './img/nalanda.jpg', './img/rajendra.jpg', './img/rajgir.jpg', './img/sama.jpg', './img/shershah.jpg', './img/silk.jpg', './img/sonepur.jpg', './img/tilkut.jpg', './img/vaishali.jpg', './img/valmiki.jpg', './img/vidyapati.jpg', './img/vikramshila.jpg'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  e.respondWith(caches.match(e.request).then(cached => {
    const fresh = fetch(e.request).then(res => {
      if (res.ok) caches.open(CACHE).then(c => c.put(e.request, res.clone()));
      return res;
    }).catch(() => cached);
    return cached || fresh;
  }));
});
