// Service Worker — ให้เปิดแอปได้แม้ไม่มีเน็ต
// หลักการ: ออนไลน์ก่อนเสมอ — โหลดจากเน็ตทุกครั้ง (ได้ของใหม่) และเก็บสำเนาไว้
//          ใช้สำเนาในเครื่องเฉพาะตอนโหลดจากเน็ตไม่ได้ (ไม่มีเน็ต / เซิร์ฟเวอร์ไม่ตอบ)
// - ไฟล์ของแอป (หน้าเว็บ / css / js / ข้อมูล) และ ฟอนต์ / ตัวอ่านบาร์โค้ดจาก CDN ใช้หลักการเดียวกัน
// - Google Apps Script (บันทึก/ซิงค์): ไม่ยุ่ง ให้แอปจัดการคิวรอส่งเอง
var VERSION = '20261127-0432'; // ตั้งโดย tools/bump-version.py
var CACHE = 'h9pl-' + VERSION;
var CDN_CACHE = 'h9pl-cdn';
var PRECACHE = [
  './',
  'index.html',
  'version.json',
  'css/style.css?v=' + VERSION,
  'js/config.js?v=' + VERSION,
  'js/data.js?v=' + VERSION,
  'js/app.js?v=' + VERSION,
  'js/vendor/jsqr.js?v=2',
  'js/partialqr.js?v=1',
  'js/vendor/qrcode.js?v=1'
];
var CDN_HOSTS = ['fonts.googleapis.com', 'fonts.gstatic.com', 'cdn.jsdelivr.net'];

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) {
    // ทีละไฟล์: ไฟล์ไหนโหลดไม่ได้ไม่ทำให้ทั้งหมดล้ม
    return Promise.all(PRECACHE.map(function (u) { return c.add(new Request(u, { cache: 'reload' })).catch(function () {}); }));
  }).then(function () { return self.skipWaiting(); }));
});

self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k.indexOf('h9pl-') === 0 && k !== CACHE && k !== CDN_CACHE; })
      .map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

// กันเน็ตค้างไม่ตอบเลย (เช่น ต่อ Wi-Fi ได้แต่ไม่มีอินเทอร์เน็ต) — รอนานพอให้เน็ตช้ายังโหลดจากเน็ตได้
function timeout(ms) { return new Promise(function (_, bad) { setTimeout(function () { bad(new Error('timeout')); }, ms); }); }
var NET_WAIT = 20000;

// เน็ตก่อน → โหลดไม่ได้จริงๆ ค่อยใช้ที่เก็บไว้ (เทียบแบบไม่สน ?v= เป็นทางสุดท้าย)
function networkFirst(req, cacheName) {
  return Promise.race([fetch(req), timeout(NET_WAIT)]).then(function (res) {
    if (res && (res.ok || res.type === 'opaque')) { var copy = res.clone(); caches.open(cacheName || CACHE).then(function (c) { c.put(req, copy); }); }
    return res;
  }).catch(function () {
    return caches.match(req).then(function (hit) {
      if (hit) return hit;
      return caches.match(req, { ignoreSearch: true }).then(function (h2) {
        if (h2) return h2;
        if (req.mode === 'navigate') return caches.match('index.html', { ignoreSearch: true });
        return Response.error();
      });
    });
  });
}

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);
  if (url.origin === self.location.origin) { e.respondWith(networkFirst(req)); return; }
  if (CDN_HOSTS.indexOf(url.hostname) >= 0) { e.respondWith(networkFirst(req, CDN_CACHE)); return; }
});
