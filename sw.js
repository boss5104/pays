// Service worker: ทำให้แอปเปิดได้แม้ออฟไลน์ และติดตั้งลงหน้าจอได้
// เปลี่ยนชื่อแคชทุกครั้งที่ออกเวอร์ชันใหม่ (ให้ตรงกับ APP_VERSION ใน index.html และ version.json)
const CACHE = 'smart-expense-2.4.0';
const SHELL = [
  './',
  './index.html',
  './manifest.webmanifest',
  './google-apps-script.gs',
  './icons/icon.svg',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/maskable-512.png',
  './icons/apple-touch-icon.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // ไฟล์ตรวจเวอร์ชัน: ดึงจากเน็ตเสมอ ไม่ใช้แคช
  if (url.origin === self.location.origin && url.pathname.endsWith('/version.json')) return;

  // หน้าเว็บ: ลองดึงจากเน็ตก่อน (ได้เวอร์ชันล่าสุด) ถ้าออฟไลน์ใช้จากแคช
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then((res) => { const copy = res.clone(); caches.open(CACHE).then((c) => c.put('./index.html', copy)); return res; })
        .catch(() => caches.match('./index.html'))
    );
    return;
  }

  // ไฟล์ในแอปและฟอนต์: ใช้แคชก่อน แล้วอัปเดตเบื้องหลัง
  const isFont = url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com';
  if (url.origin === self.location.origin || isFont) {
    event.respondWith(
      caches.match(req).then((cached) => {
        const network = fetch(req)
          .then((res) => {
            if (res && (res.ok || res.type === 'opaque')) {
              const copy = res.clone();
              caches.open(CACHE).then((c) => c.put(req, copy));
            }
            return res;
          })
          .catch(() => cached);
        return cached || network;
      })
    );
  }
});
