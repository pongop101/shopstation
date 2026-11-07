// ShopStation service worker — ทำให้แอปเปิดได้แม้ไม่มีอินเทอร์เน็ต
const VER = 'shopstation-v3';
const SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png',
  'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css',
  'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js'];
const MEDIA = 'shopstation-media';
const MEDIA_MAX = 600;

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VER).then(c => Promise.all(SHELL.map(u => c.add(u).catch(() => null)))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VER && k !== MEDIA).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
async function trim() {
  const c = await caches.open(MEDIA); const ks = await c.keys();
  for (let i = 0; i < ks.length - MEDIA_MAX; i++) await c.delete(ks[i]);
}
self.addEventListener('fetch', e => {
  const req = e.request; if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // API / Google Maps — ไม่แคช
  if (/script\.google(usercontent)?\.com|maps\.googleapis\.com|maps\.gstatic\.com/.test(url.host)) return;
  // หน้าแอป — ลองเน็ตก่อน (ได้เวอร์ชันใหม่) ถ้าไม่มีเน็ตใช้แคช
  if (req.mode === 'navigate' || (url.origin === location.origin && /\/(index\.html)?$/.test(url.pathname))) {
    e.respondWith(fetch(req).then(r => { const cp = r.clone(); caches.open(VER).then(c => c.put('./index.html', cp)); return r; })
      .catch(() => caches.match('./index.html')));
    return;
  }
  // แผ่นแผนที่ และรูปจาก Google Drive — แคชไว้ใช้ออฟไลน์
  if (/tile\.openstreetmap\.org|arcgisonline\.com|drive\.google\.com|googleusercontent\.com/.test(url.host)) {
    e.respondWith(caches.open(MEDIA).then(async c => {
      const hit = await c.match(req);
      const net = fetch(req).then(r => { if (r && (r.ok || r.type === 'opaque')) { c.put(req, r.clone()); trim(); } return r; }).catch(() => hit);
      return hit || net;
    }));
    return;
  }
  // ไฟล์อื่น ๆ (Leaflet, ฟอนต์, ไอคอน) — แคชก่อน
  if (url.origin === location.origin || /cdnjs\.cloudflare\.com|fonts\.(googleapis|gstatic)\.com/.test(url.host)) {
    e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(r => {
      if (r && (r.ok || r.type === 'opaque')) { const cp = r.clone(); caches.open(VER).then(c => c.put(req, cp)); }
      return r;
    })));
  }
});
