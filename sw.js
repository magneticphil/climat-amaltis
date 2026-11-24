/* Climat — Amaltis : l'écran s'ouvre même sans réseau (réseau d'abord, copie locale après 4 s).
   Les données ne passent jamais par ici : elles viennent du script Google. */
const CACHE = 'climat-v1';
const FICHIERS = ['./', './index.html', './manifest.webmanifest', './icone-180.png', './icone-192.png', './icone-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FICHIERS)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('climat-') && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== self.location.origin) return;
  const copie = () => caches.match(e.request, { ignoreSearch: true }).then(r => r || caches.match('./index.html'));
  const reseau = fetch(e.request).then(r => { if (r.ok) { const c = r.clone(); caches.open(CACHE).then(ca => ca.put(e.request, c)); } return r; });
  const delai = new Promise(ok => setTimeout(ok, 4000)).then(copie);
  e.respondWith(Promise.race([reseau.catch(copie), delai.then(r => r || reseau)]));
});
