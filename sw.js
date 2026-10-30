// Кэш приложения: после первого открытия работает без интернета. Данные КП хранятся только в памяти браузера.
var CACHE='kp-hub-v6';
var CORE=['./','index.html','manifest.webmanifest','icon-192.png','icon-512.png'];
self.addEventListener('install',function(e){e.waitUntil(caches.open(CACHE).then(function(c){return c.addAll(CORE)}).then(function(){return self.skipWaiting()}))});
self.addEventListener('activate',function(e){e.waitUntil(caches.keys().then(function(ks){return Promise.all(ks.filter(function(k){return k!==CACHE}).map(function(k){return caches.delete(k)}))}).then(function(){return self.clients.claim()}))});
self.addEventListener('fetch',function(e){
  if(e.request.method!=='GET')return;
  var url=new URL(e.request.url);
  // страница: сначала сеть (чтобы приходили обновления), без сети — из кэша
  if(e.request.mode==='navigate'){
    e.respondWith(fetch(e.request).then(function(r){var cp=r.clone();caches.open(CACHE).then(function(c){c.put('index.html',cp)});return r}).catch(function(){return caches.match('index.html')}));
    return;
  }
  // остальное (иконки, шрифты): из кэша, иначе сеть с сохранением
  e.respondWith(caches.match(e.request).then(function(hit){
    return hit||fetch(e.request).then(function(r){if(url.origin===location.origin||url.host.indexOf('fonts.')>=0){var cp=r.clone();caches.open(CACHE).then(function(c){c.put(e.request,cp)})}return r});
  }));
});
