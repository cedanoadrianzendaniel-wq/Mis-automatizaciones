// ═══════════════════════════════════════════════════════════════════════════
// OFFLINE QUEUE — TGP Reportes
// Guarda envios fallidos en IndexedDB y los reintenta cuando vuelve la senal.
// Tambien expone el indicador verde/rojo/amarillo en pantalla.
// ═══════════════════════════════════════════════════════════════════════════
(function(window) {
  "use strict";

  var DB_NAME = "tgp-offline-queue";
  var DB_VERSION = 1;
  var STORE = "pendientes";

  // ─── INDEXEDDB ────────────────────────────────────────────────────────────
  function openDB() {
    return new Promise(function(resolve, reject) {
      var req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = function(e) {
        var db = e.target.result;
        if (!db.objectStoreNames.contains(STORE)) {
          db.createObjectStore(STORE, { keyPath: "id", autoIncrement: true });
        }
      };
      req.onsuccess = function(e) { resolve(e.target.result); };
      req.onerror = function(e) { reject(e.target.error); };
    });
  }

  function addToQueue(item) {
    return openDB().then(function(db) {
      return new Promise(function(resolve, reject) {
        var tx = db.transaction(STORE, "readwrite");
        var req = tx.objectStore(STORE).add(item);
        req.onsuccess = function() { resolve(req.result); };
        req.onerror = function() { reject(req.error); };
      });
    });
  }

  function getAll() {
    return openDB().then(function(db) {
      return new Promise(function(resolve, reject) {
        var tx = db.transaction(STORE, "readonly");
        var req = tx.objectStore(STORE).getAll();
        req.onsuccess = function() { resolve(req.result || []); };
        req.onerror = function() { reject(req.error); };
      });
    });
  }

  function removeFromQueue(id) {
    return openDB().then(function(db) {
      return new Promise(function(resolve, reject) {
        var tx = db.transaction(STORE, "readwrite");
        var req = tx.objectStore(STORE).delete(id);
        req.onsuccess = function() { resolve(); };
        req.onerror = function() { reject(req.error); };
      });
    });
  }

  function updateItem(item) {
    return openDB().then(function(db) {
      return new Promise(function(resolve, reject) {
        var tx = db.transaction(STORE, "readwrite");
        var req = tx.objectStore(STORE).put(item);
        req.onsuccess = function() { resolve(); };
        req.onerror = function() { reject(req.error); };
      });
    });
  }

  // ─── INDICADOR DE ESTADO ──────────────────────────────────────────────────
  function ensureIndicator() {
    var el = document.getElementById("conn-indicator");
    if (el) return el;
    el = document.createElement("div");
    el.id = "conn-indicator";
    el.style.cssText = [
      "position:fixed","top:0","left:0","right:0","z-index:9999",
      "padding:8px 14px","font-family:Arial,sans-serif","font-size:12px",
      "font-weight:bold","text-align:center","color:white",
      "transition:background .3s","box-shadow:0 1px 4px rgba(0,0,0,.2)"
    ].join(";");
    document.body.appendChild(el);
    // Empuja el contenido hacia abajo para no taparlo
    document.body.style.paddingTop = "36px";
    return el;
  }

  function setIndicator(state, count) {
    var el = ensureIndicator();
    if (state === "online") {
      el.style.background = "#059669"; // verde
      el.textContent = count > 0
        ? "🟢 En linea — " + count + " reporte(s) en cola"
        : "🟢 En linea";
    } else if (state === "offline") {
      el.style.background = "#DC2626"; // rojo
      el.textContent = count > 0
        ? "🔴 Sin conexion — " + count + " reporte(s) pendiente(s)"
        : "🔴 Sin conexion";
    } else if (state === "syncing") {
      el.style.background = "#D97706"; // amarillo/naranja
      el.textContent = "🟡 Sincronizando " + count + " reporte(s)...";
    }
  }

  // ─── REFRESCAR ESTADO VISUAL ──────────────────────────────────────────────
  function refreshIndicator() {
    return getAll().then(function(items) {
      var pending = items.length;
      setIndicator(navigator.onLine ? "online" : "offline", pending);
    }).catch(function() {
      setIndicator(navigator.onLine ? "online" : "offline", 0);
    });
  }

  // ─── ENVIO CON FALLBACK A COLA ────────────────────────────────────────────
  // Intenta hacer fetch normal. Si falla por red, guarda en IndexedDB.
  function submitOrQueue(endpoint, payload, formType) {
    if (!navigator.onLine) {
      return queueAndAck(endpoint, payload, formType);
    }
    return fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    })
    .then(function(r) {
      if (!r.ok) throw new Error("HTTP " + r.status);
      return r.json();
    })
    .then(function(json) {
      if (!json.ok) throw new Error(json.error || "Error servidor");
      refreshIndicator();
      return { ok: true, queued: false, response: json };
    })
    .catch(function(err) {
      // Si el error es de red (no de validacion del server), encolar
      if (err.message.indexOf("HTTP 4") === 0) {
        // 4xx = error de cliente, no encolar
        throw err;
      }
      return queueAndAck(endpoint, payload, formType);
    });
  }

  function queueAndAck(endpoint, payload, formType) {
    return addToQueue({
      endpoint: endpoint,
      payload: payload,
      formType: formType || "campo",
      timestamp: Date.now(),
      retries: 0
    }).then(function() {
      refreshIndicator();
      return { ok: true, queued: true };
    });
  }

  // ─── SYNC: procesar la cola cuando hay senal ──────────────────────────────
  var syncing = false;
  function syncQueue() {
    if (syncing || !navigator.onLine) return Promise.resolve();
    syncing = true;
    return getAll().then(function(items) {
      if (items.length === 0) {
        syncing = false;
        refreshIndicator();
        return;
      }
      setIndicator("syncing", items.length);
      // Procesar de a uno con espacio entre requests (evita saturar Google APIs)
      return processOne(items, 0);
    }).then(function() {
      syncing = false;
      refreshIndicator();
    }).catch(function(err) {
      console.warn("[Queue] Error sync:", err);
      syncing = false;
      refreshIndicator();
    });
  }

  function processOne(items, idx) {
    if (idx >= items.length) return Promise.resolve();
    var item = items[idx];
    return fetch(item.endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(item.payload)
    })
    .then(function(r) { return r.json().then(function(j) { return { ok: r.ok, json: j }; }); })
    .then(function(res) {
      if (res.ok && res.json && res.json.ok) {
        return removeFromQueue(item.id);
      } else {
        // Server respondio pero rechazo: marcar como fallido pero dejar para reintento
        item.retries = (item.retries || 0) + 1;
        item.lastError = (res.json && res.json.error) || "rechazado";
        if (item.retries >= 5) {
          // tras 5 intentos fallidos del server, lo dejamos pero no bloqueamos
          console.warn("[Queue] Item con muchos reintentos:", item);
        }
        return updateItem(item);
      }
    })
    .catch(function(err) {
      // Error de red: dejar en cola, abortar sync
      console.warn("[Queue] Red caida durante sync:", err);
      throw err;
    })
    .then(function() {
      // Espacio entre envios para no saturar Google APIs
      return new Promise(function(r) { setTimeout(r, 1500); });
    })
    .then(function() {
      return getAll().then(function(remaining) {
        setIndicator("syncing", remaining.length);
        if (remaining.length === 0) return;
        return processOne(remaining, 0);
      });
    });
  }

  // ─── EVENTOS ──────────────────────────────────────────────────────────────
  window.addEventListener("online",  function() { refreshIndicator(); syncQueue(); });
  window.addEventListener("offline", function() { refreshIndicator(); });

  // Reintento periodico cada 60 seg (por si el evento online no dispara)
  setInterval(function() {
    if (navigator.onLine) syncQueue();
    else refreshIndicator();
  }, 60000);

  // Inicializar al cargar
  document.addEventListener("DOMContentLoaded", function() {
    refreshIndicator();
    if (navigator.onLine) syncQueue();
  });

  // ─── REGISTRO DE SERVICE WORKER ───────────────────────────────────────────
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", function() {
      navigator.serviceWorker.register("/service-worker.js")
        .then(function(reg) { console.log("[SW] Registrado:", reg.scope); })
        .catch(function(err) { console.warn("[SW] No se pudo registrar:", err); });
    });
  }

  // ─── API PUBLICA ──────────────────────────────────────────────────────────
  window.TGPOffline = {
    submit: submitOrQueue,
    sync:   syncQueue,
    pending: getAll,
    refresh: refreshIndicator
  };

})(window);
