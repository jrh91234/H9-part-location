// Service Worker — ให้เปิดแอปได้แม้ไม่มีเน็ต
// - ไฟล์ของแอป (หน้าเว็บ / css / js / ข้อมูล): ลองโหลดจากเน็ตก่อน (ได้ของใหม่เสมอ) ถ้าไม่มีเน็ตหรือช้าเกิน 4 วินาที → ใช้ที่เก็บไว้
// - ฟอนต์ / ตัวอ่านบาร์โค้ดจาก CDN: เก็บไว้หลังโหลดครั้งแรก แล้วใช้ของในเครื่อง
// - Google Apps Script (บันทึก/ซิงค์): ไม่ยุ่ง ให้แอปจัดการคิวรอส่งเอง
var VERSION = '20261127-0136'; // ตั้งโดย tools/bump-version.py
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

function timeout(ms) { return new Promise(function (_, bad) { setTimeout(function () { bad(new Error('timeout')); }, ms); }); }

// เน็ตก่อน → ไม่ได้ค่อยใช้ที่เก็บไว้ (เทียบแบบไม่สน ?v= เป็นทางสุดท้าย)
function networkFirst(req) {
  return Promise.race([fetch(req), timeout(4000)]).then(function (res) {
    if (res && res.ok) { var copy = res.clone(); caches.open(CACHE).then(function (c) { c.put(req, copy); }); }
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

// ของในเครื่องก่อน → ไม่มีค่อยโหลดแล้วเก็บไว้
function cacheFirst(req) {
  return caches.open(CDN_CACHE).then(function (c) {
    return c.match(req).then(function (hit) {
      return hit || fetch(req).then(function (res) {
        if (res && (res.ok || res.type === 'opaque')) c.put(req, res.clone());
        return res;
      });
    });
  });
}

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);
  if (url.origin === self.location.origin) { e.respondWith(networkFirst(req)); return; }
  if (CDN_HOSTS.indexOf(url.hostname) >= 0) { e.respondWith(cacheFirst(req)); return; }
});
