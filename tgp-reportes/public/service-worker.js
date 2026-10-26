// ═══════════════════════════════════════════════════════════════════════════
// SERVICE WORKER — KILL SWITCH
// Reemplazo del SW original. Al activarse:
//   1. Borra TODAS las cachés que hubiera.
//   2. Fuerza reload de todos los clientes abiertos.
//   3. Se auto-desregistra.
// Después de esto, el navegador deja de usar SW en este dominio.
// Firma: killswitch-2026-10-26
// ═══════════════════════════════════════════════════════════════════════════

const KILLSWITCH_VERSION = "killswitch-2026-10-26";

self.addEventListener("install", (event) => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    try {
      const cacheNames = await caches.keys();
      await Promise.all(cacheNames.map((name) => caches.delete(name)));
    } catch (e) {}

    try {
      await self.clients.claim();
    } catch (e) {}

    try {
      const windowClients = await self.clients.matchAll({
        type: "window",
        includeUncontrolled: true
      });
      for (const client of windowClients) {
        try { client.navigate(client.url); } catch (e) {}
      }
    } catch (e) {}

    try {
      await self.registration.unregister();
    } catch (e) {}
  })());
});

self.addEventListener("fetch", (event) => {
  // No-op: no interceptamos nada. El navegador va al network directo.
  return;
});
