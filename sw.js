/* HUFC Session Planner — keeps the planner on the phone so it opens with no signal */
const VERSION = '20261027102334-ebd5eeab';
const CACHE = 'hufc-' + VERSION;
const FILES = ["./", "index.html", "lib-pdf.js", "lib-pdffonts.js", "lib-pptx.js", "logos.json", "manifest.webmanifest", "apple-touch-icon.png", "icon-180.png", "icon-192.png", "icon-512.png", "icon-512-maskable.png"];
self.addEventListener('install', function(e){
  e.waitUntil(caches.open(CACHE).then(function(c){ return c.addAll(FILES.map(function(f){ return new Request(f, {cache:'reload'}); })); }).then(function(){ return self.skipWaiting(); }));
});
self.addEventListener('activate', function(e){
  e.waitUntil(caches.keys().then(function(keys){
    const old = keys.filter(function(k){ return k.indexOf('hufc-') === 0 && k !== CACHE; });
    return Promise.all(old.map(function(k){ return caches.delete(k); })).then(function(){ return self.clients.claim(); }).then(function(){
      /* a newer planner replaced an older one: tell any open page */
      if(old.length) return self.clients.matchAll().then(function(cs){ cs.forEach(function(c){ c.postMessage('hufc-updated'); }); });
    });
  }));
});
/* the planner always opens from the copy on the phone (instant, works offline) */
self.addEventListener('fetch', function(e){
  const req = e.request;
  if(req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  const nav = req.mode === 'navigate';
  e.respondWith(caches.open(CACHE).then(function(c){
    return c.match(nav ? 'index.html' : req, {ignoreSearch:true}).then(function(hit){
      return hit || fetch(req).catch(function(){ return nav ? c.match('index.html') : Response.error(); });
    });
  }));
});
