/* Night Walker サービスワーカー
 *
 * アプリ本体を端末に保存し、電波が悪い場所でも起動できるようにする。
 * 地図タイル・天気・見守り送信は通信が必要なため、圏外では使えない。
 */
var APP_VERSION = "2.3.0";
var CACHE = "night-walker-v" + APP_VERSION;

var ASSETS = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./icon-192.png",
  "./icon-512.png",
  "./apple-touch-icon.png"
];

self.addEventListener("install", function(e){
  e.waitUntil(
    caches.open(CACHE).then(function(c){
      return c.addAll(ASSETS);
    }).then(function(){ return self.skipWaiting(); })
  );
});

self.addEventListener("activate", function(e){
  e.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.map(function(k){
        if(k !== CACHE) return caches.delete(k);
      }));
    }).then(function(){ return self.clients.claim(); })
  );
});

self.addEventListener("fetch", function(e){
  var req = e.request;
  if(req.method !== "GET") return;

  var url = new URL(req.url);

  // 外部への通信（地図タイル・天気・スプレッドシート）はキャッシュせず、そのまま通す
  if(url.origin !== self.location.origin){
    return;
  }

  // アプリ本体は「まずネットワーク、だめならキャッシュ」で、更新を取り込みやすくする
  e.respondWith(
    fetch(req).then(function(res){
      var copy = res.clone();
      caches.open(CACHE).then(function(c){ c.put(req, copy); }).catch(function(){});
      return res;
    }).catch(function(){
      return caches.match(req).then(function(hit){
        return hit || caches.match("./index.html");
      });
    })
  );
});
