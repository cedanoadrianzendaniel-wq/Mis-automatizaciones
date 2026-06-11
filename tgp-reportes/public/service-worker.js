// ═══════════════════════════════════════════════════════════════════════════
// SERVICE WORKER — TGP Reportes Offline
// Cachea archivos estaticos y respuestas de listas para que los formularios
// abran sin conexion. Los envios POST NO se cachean; los maneja offline-queue.js
// ═══════════════════════════════════════════════════════════════════════════

const CACHE_VERSION = "tgp-reportes-v7";
const STATIC_CACHE  = CACHE_VERSION + "-static";
const API_CACHE     = CACHE_VERSION + "-api";

// Archivos que se precargan al instalar el SW
const PRECACHE_URLS = [
  "/",
  "/hse",
  "/formulario.html",
  "/formulario-hse.html",
  "/offline-queue.js",
  "/manifest.json"
];

// Endpoints de listas (cache-first con actualizacion en background)
const API_LIST_ENDPOINTS = [
  "/api/proyectos-capex",
  "/api/supervisores",
  "/api/todos-frentes"
];

// ─── INSTALL ────────────────────────────────────────────────────────────────
self.addEventListener("install", function(event) {
  event.waitUntil(
    caches.open(STATIC_CACHE).then(function(cache) {
      return cache.addAll(PRECACHE_URLS).catch(function(err) {
        console.warn("[SW] Precache parcial:", err);
      });
    }).then(function() { return self.skipWaiting(); })
  );
});

// ─── ACTIVATE ───────────────────────────────────────────────────────────────
self.addEventListener("activate", function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(keys.map(function(k) {
        if (k !== STATIC_CACHE && k !== API_CACHE) return caches.delete(k);
      }));
    }).then(function() { return self.clients.claim(); })
  );
});

// ─── FETCH ──────────────────────────────────────────────────────────────────
self.addEventListener("fetch", function(event) {
  const req = event.request;
  const url = new URL(req.url);

  // Solo manejamos GET; POST de reportes pasa directo (lo maneja la cola JS)
  if (req.method !== "GET") return;

  // Solo same-origin
  if (url.origin !== self.location.origin) return;

  // ─── Dashboard y portal: network-only (paginas de coordinador, online-only) ──
  // No las cacheamos para que actualizaciones se vean siempre y no por SW stale.
  if (url.pathname === "/dashboard" ||
      url.pathname === "/portal" ||
      url.pathname === "/dashboard.html" ||
      url.pathname === "/portal.html") {
    return; // dejar que el browser haga fetch normal
  }

  // ─── Listas API: stale-while-revalidate ─────────
  const isListEndpoint = API_LIST_ENDPOINTS.some(function(p) {
    return url.pathname === p || url.pathname.startsWith(p + "?");
  });
  if (isListEndpoint) {
    event.respondWith(staleWhileRevalidate(req));
    return;
  }

  // ─── Otros API: network-only (no cachear) ──────
  if (url.pathname.startsWith("/api/")) {
    return; // dejar que el browser haga fetch normal
  }

  // ─── Estaticos: cache-first con fallback de red ─────
  event.respondWith(cacheFirst(req));
});

// ─── ESTRATEGIAS ────────────────────────────────────────────────────────────
async function cacheFirst(req) {
  const cached = await caches.match(req);
  if (cached) return cached;
  try {
    const fresh = await fetch(req);
    if (fresh.ok) {
      const cache = await caches.open(STATIC_CACHE);
      cache.put(req, fresh.clone());
    }
    return fresh;
  } catch (err) {
    // Si pidio HTML y no hay cache, servir formulario.html como fallback
    if (req.headers.get("accept") && req.headers.get("accept").indexOf("text/html") !== -1) {
      const fallback = await caches.match("/formulario.html");
      if (fallback) return fallback;
    }
    throw err;
  }
}

async function staleWhileRevalidate(req) {
  const cache = await caches.open(API_CACHE);
  const cached = await cache.match(req);
  const networkPromise = fetch(req).then(function(fresh) {
    if (fresh.ok) cache.put(req, fresh.clone());
    return fresh;
  }).catch(function() { return null; });

  if (cached) {
    networkPromise; // actualiza en background
    return cached;
  }
  const fresh = await networkPromise;
  if (fresh) return fresh;

  // Sin cache y sin red: intentar fallback inteligente
  const url = new URL(req.url);

  // Para /api/supervisores?subcategoria=X, intentar /api/supervisores (sin filtro)
  // y filtrar del lado del SW
  if (url.pathname === "/api/supervisores" && url.searchParams.get("subcategoria")) {
    const fullCache = await cache.match(self.location.origin + "/api/supervisores");
    if (fullCache) {
      try {
        const all = await fullCache.json();
        const subcat = url.searchParams.get("subcategoria");
        const filtered = all.filter(function(s) { return s.subcategoria === subcat; });
        return new Response(JSON.stringify(filtered), {
          headers: { "Content-Type": "application/json" }
        });
      } catch (e) {}
    }
    // Fallback final: array vacio
    return new Response("[]", { headers: { "Content-Type": "application/json" } });
  }

  // Fallbacks por endpoint
  if (url.pathname === "/api/supervisores") {
    return new Response("[]", { headers: { "Content-Type": "application/json" } });
  }
  if (url.pathname === "/api/proyectos-capex") {
    return new Response(JSON.stringify({ pep: [], cuentaOrden: [] }), {
      headers: { "Content-Type": "application/json" }
    });
  }
  if (url.pathname === "/api/todos-frentes") {
    return new Response("{}", { headers: { "Content-Type": "application/json" } });
  }
  return new Response("{}", { headers: { "Content-Type": "application/json" } });
}
