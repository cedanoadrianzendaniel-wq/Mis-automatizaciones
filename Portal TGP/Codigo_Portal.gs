// ═══════════════════════════════════════════════════════════════════════════
// PORTAL TGP — Bureau Veritas | Codigo_Portal.gs v1
// Proyecto independiente — NO modifica formularios existentes
// ═══════════════════════════════════════════════════════════════════════════

var CLAVE_DEFAULT   = "tgp2026";
var SHEET_AVANCE    = "TGP - Avance Frentes";
var SHEET_CAMPO     = "Reportes de Campo 2026 - Registros";
var CARPETA_DOCS    = "Portal TGP";

// ─── FRENTES POR SECTOR ──────────────────────────────────────────────────────
var FRENTES_PORTAL = {
  "Costa": {
    subcategoria: "Geotecnia",
    frentes: [
      { nombre: "SOPORTE A INGENIERIA",                    cuenta: "63800018", orden: "TGGA/COS-26-02-01" },
      { nombre: "VIAL",                                    cuenta: "63800018", orden: "TG3CDV1"           },
      { nombre: "URGENCIA VIAL KP 472+700 AL KP 482+000",  cuenta: "63440004", orden: "TGCI-2644"         },
      { nombre: "KP 714+155 AL KP 730+698",                cuenta: "63440002", orden: "TGEO-2629"         },
      { nombre: "MG KP 519+526 AL KP 540+839",             cuenta: "63440002", orden: "TGEO-2631"         }
    ]
  },
  "Sierra": {
    subcategoria: "Geotecnia",
    frentes: [
      { nombre: "MG KP 170+000 AL KP 179+850", cuenta: "63440002", orden: "TGEO-2621"          },
      { nombre: "MG KP 179+850 AL KP 194+000", cuenta: "63440002", orden: "TGEO-2622"          },
      { nombre: "MG KP 194+000 AL KP 209+360", cuenta: "63440002", orden: "TGEO-2623"          },
      { nombre: "APOYO A INGENIERIA",           cuenta: "63290002", orden: "TGGA/SIE-26-02-02" }
    ]
  },
  "Selva": {
    subcategoria: "Geotecnia",
    frentes: [
      { nombre: "M.G. KP 43+830 - KP 53+000 - ETAPA 1", cuenta: "63800018", orden: "TGEO-2610"          },
      { nombre: "M.G. KP 0+000 - KP 12+000 - ETAPA 1",  cuenta: "63800018", orden: "TGEO-2606"          },
      { nombre: "M.G. KP 25+000 - KP 35+000",            cuenta: "63800018", orden: "TGEO-2608"          },
      { nombre: "TAI KP 112+300",                         cuenta: "63800018", orden: "TGEO-2601"          },
      { nombre: "Reparacion de F.O. KP61+150",           cuenta: "63800018", orden: "TG1CDV1"            },
      { nombre: "Perforaciones del KP126",               cuenta: "63800018", orden: "TGGA/SEL-25-08-07"  },
      { nombre: "Perforaciones del KP55+118",            cuenta: "63800018", orden: "TGGA/SEL-26-02-01"  },
      { nombre: "Apoyo / Ingenieria",                    cuenta: "63800018", orden: "TGGA/SEL-26-03-01"  }
    ]
  }
};

// ─── ROUTING ──────────────────────────────────────────────────────────────────
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("Portal")
    .setTitle("Portal TGP — Bureau Veritas")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── AUTENTICACIÓN ────────────────────────────────────────────────────────────
function verificarClavePortal(clave) {
  var claveGuardada = PropertiesService.getScriptProperties()
    .getProperty("CLAVE_PORTAL") || CLAVE_DEFAULT;
  return String(clave).trim() === claveGuardada;
}

// ─── HOME: RESUMEN DE SECTORES ────────────────────────────────────────────────
function obtenerResumenSectores() {
  try {
    return ["Costa", "Sierra", "Selva"].map(function (sector) {
      var frentes         = FRENTES_PORTAL[sector].frentes;
      var avances         = frentes.map(function (f) { return obtenerUltimoAvance(sector, f.nombre); });
      var avancePromedio  = Math.round(avances.reduce(function (a, b) { return a + b; }, 0) / avances.length);
      return {
        sector:          sector,
        numFreentes:     frentes.length,
        avancePromedio:  avancePromedio,
        ultimaActividad: obtenerUltimaActSector(sector)
      };
    });
  } catch (e) {
    Logger.log("obtenerResumenSectores: " + e);
    return [
      { sector: "Costa",  numFreentes: 5, avancePromedio: 0, ultimaActividad: "-" },
      { sector: "Sierra", numFreentes: 4, avancePromedio: 0, ultimaActividad: "-" },
      { sector: "Selva",  numFreentes: 8, avancePromedio: 0, ultimaActividad: "-" }
    ];
  }
}

// ─── SECTOR: FRENTES CON AVANCE ───────────────────────────────────────────────
function obtenerFrentesConAvance(sector) {
  var config = FRENTES_PORTAL[sector];
  if (!config) return [];
  return config.frentes.map(function (f) {
    return {
      nombre:          f.nombre,
      cuenta:          f.cuenta,
      orden:           f.orden,
      avance:          obtenerUltimoAvance(sector, f.nombre),
      ultimaActividad: obtenerUltimaActFrente(sector, f.nombre)
    };
  });
}

// ─── FRENTE: DETALLE GENERAL ──────────────────────────────────────────────────
function obtenerDetalleFrente(sector, frente) {
  var config     = FRENTES_PORTAL[sector] || { subcategoria: "", frentes: [] };
  var frenteInfo = null;
  config.frentes.forEach(function (f) { if (f.nombre === frente) frenteInfo = f; });
  return {
    sector:       sector,
    subcategoria: config.subcategoria,
    nombre:       frente,
    cuenta:       frenteInfo ? frenteInfo.cuenta : "",
    orden:        frenteInfo ? frenteInfo.orden  : "",
    avanceActual: obtenerUltimoAvance(sector, frente)
  };
}

// ─── CURVA DE AVANCE ──────────────────────────────────────────────────────────
function obtenerCurvaAvance(sector, frente) {
  try {
    var ss = buscarSheet(SHEET_AVANCE);
    if (!ss) return [];
    var hoja = ss.getSheetByName("Avance");
    if (!hoja || hoja.getLastRow() <= 1) return [];

    var datos  = hoja.getRange(2, 1, hoja.getLastRow() - 1, 7).getValues();
    var puntos = [];
    datos.forEach(function (r) {
      var fechaFila = r[1] instanceof Date
        ? Utilities.formatDate(r[1], "America/Lima", "yyyy-MM-dd")
        : String(r[1]).substring(0, 10);
      if (String(r[2]).toLowerCase() === sector.toLowerCase() && String(r[3]) === frente) {
        puntos.push({ fecha: fechaFila, avance: parseFloat(r[4]) || 0, obs: String(r[5]) });
      }
    });
    puntos.sort(function (a, b) { return a.fecha.localeCompare(b.fecha); });
    return puntos;
  } catch (e) {
    Logger.log("obtenerCurvaAvance: " + e);
    return [];
  }
}

// ─── ACTUALIZAR AVANCE (manual por coordinador) ────────────────────────────────
function actualizarAvanceFrente(sector, frente, porcentaje, observacion, usuario) {
  try {
    var ss   = abrirOCrearSheetAvance();
    var hoja = ss.getSheetByName("Avance");
    var ahora = Utilities.formatDate(new Date(), "America/Lima", "yyyy-MM-dd HH:mm");
    var hoy   = Utilities.formatDate(new Date(), "America/Lima", "yyyy-MM-dd");
    hoja.appendRow([
      ahora, hoy, sector, frente,
      parseFloat(porcentaje) || 0,
      observacion || "",
      usuario     || "Portal"
    ]);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

// ─── DOCUMENTOS ───────────────────────────────────────────────────────────────
function obtenerDocumentosFrente(sector, frente) {
  try {
    var carpeta = obtenerCarpetaDocs(sector, frente, false);
    if (!carpeta) return [];
    var docs  = [];
    var files = carpeta.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      docs.push({
        nombre: f.getName(),
        url:    f.getUrl(),
        fecha:  Utilities.formatDate(f.getDateCreated(), "America/Lima", "dd/MM/yyyy"),
        tamano: (Math.round(f.getSize() / 1024)) + " KB"
      });
    }
    docs.reverse();
    return docs;
  } catch (e) {
    Logger.log("obtenerDocumentosFrente: " + e);
    return [];
  }
}

function subirDocumentoFrente(datos) {
  try {
    var carpeta   = obtenerCarpetaDocs(datos.sector, datos.frente, true);
    var contenido = Utilities.base64Decode(datos.archivoBase64);
    var blob      = Utilities.newBlob(contenido, datos.mimeType, datos.nombreArchivo);
    var archivo   = carpeta.createFile(blob);
    archivo.setDescription(datos.sector + " | " + datos.frente);
    return { ok: true, url: archivo.getUrl(), nombre: archivo.getName() };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

// ─── ACTIVIDADES RECIENTES ────────────────────────────────────────────────────
function obtenerActividadesFrente(sector, frente) {
  try {
    var ss = buscarSheet(SHEET_CAMPO);
    if (!ss) return [];
    var hoja = ss.getSheetByName("RAW_DATA");
    if (!hoja || hoja.getLastRow() <= 1) return [];

    var filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 23).getValues();
    var result = [];
    filas.forEach(function (r) {
      if (String(r[4]).toLowerCase() === sector.toLowerCase() && String(r[6]) === frente) {
        var fechaFila = r[1] instanceof Date
          ? Utilities.formatDate(r[1], "America/Lima", "dd/MM/yyyy")
          : String(r[1]).substring(0, 10);
        result.push({
          fecha:       fechaFila,
          responsable: r[2],
          puesto:      r[3],
          descripcion: r[16] || "",
          avance:      r[17] ? parseFloat(r[17]).toFixed(1) + "%" : "-"
        });
      }
    });
    result.reverse();
    return result.slice(0, 15);
  } catch (e) {
    Logger.log("obtenerActividadesFrente: " + e);
    return [];
  }
}

// ─── HELPERS INTERNOS ─────────────────────────────────────────────────────────
function obtenerUltimoAvance(sector, frente) {
  try {
    var ss = buscarSheet(SHEET_AVANCE);
    if (!ss) return 0;
    var hoja = ss.getSheetByName("Avance");
    if (!hoja || hoja.getLastRow() <= 1) return 0;
    var datos  = hoja.getRange(2, 1, hoja.getLastRow() - 1, 7).getValues();
    var ultimo = 0;
    datos.forEach(function (r) {
      if (String(r[2]).toLowerCase() === sector.toLowerCase() && String(r[3]) === frente) {
        ultimo = parseFloat(r[4]) || 0;
      }
    });
    return ultimo;
  } catch (e) { return 0; }
}

function obtenerUltimaActSector(sector) {
  try {
    var ss = buscarSheet(SHEET_CAMPO);
    if (!ss) return "-";
    var hoja = ss.getSheetByName("RAW_DATA");
    if (!hoja || hoja.getLastRow() <= 1) return "-";
    for (var i = hoja.getLastRow(); i >= 2; i--) {
      var r = hoja.getRange(i, 1, 1, 5).getValues()[0];
      if (String(r[4]).toLowerCase() === sector.toLowerCase()) {
        return r[1] instanceof Date
          ? Utilities.formatDate(r[1], "America/Lima", "dd/MM/yyyy")
          : String(r[1]).substring(0, 10);
      }
    }
    return "-";
  } catch (e) { return "-"; }
}

function obtenerUltimaActFrente(sector, frente) {
  try {
    var ss = buscarSheet(SHEET_CAMPO);
    if (!ss) return "Sin actividad";
    var hoja = ss.getSheetByName("RAW_DATA");
    if (!hoja || hoja.getLastRow() <= 1) return "Sin actividad";
    for (var i = hoja.getLastRow(); i >= 2; i--) {
      var r = hoja.getRange(i, 1, 1, 7).getValues()[0];
      if (String(r[4]).toLowerCase() === sector.toLowerCase() && String(r[6]) === frente) {
        return r[1] instanceof Date
          ? Utilities.formatDate(r[1], "America/Lima", "dd/MM/yyyy")
          : String(r[1]).substring(0, 10);
      }
    }
    return "Sin actividad";
  } catch (e) { return "-"; }
}

function buscarSheet(nombre) {
  var f = DriveApp.getFilesByName(nombre);
  return f.hasNext() ? SpreadsheetApp.open(f.next()) : null;
}

function abrirOCrearSheetAvance() {
  var f = DriveApp.getFilesByName(SHEET_AVANCE);
  if (f.hasNext()) return SpreadsheetApp.open(f.next());
  var ss   = SpreadsheetApp.create(SHEET_AVANCE);
  var hoja = ss.getActiveSheet();
  hoja.setName("Avance");
  hoja.appendRow(["Timestamp", "Fecha", "Sector", "Frente", "% Avance", "Observacion", "Usuario"]);
  hoja.getRange(1, 1, 1, 7)
    .setFontWeight("bold").setBackground("#1F3864").setFontColor("white");
  hoja.setFrozenRows(1);
  return ss;
}

function obtenerCarpetaDocs(sector, frente, crear) {
  var raiz = buscarOCrear(CARPETA_DOCS, DriveApp.getRootFolder(), crear);
  if (!raiz) return null;
  var sec  = buscarOCrear(sector,       raiz,                    crear);
  if (!sec)  return null;
  return     buscarOCrear(frente,       sec,                     crear);
}

function buscarOCrear(nombre, padre, crear) {
  var i = padre.getFoldersByName(nombre);
  if (i.hasNext()) return i.next();
  return crear ? padre.createFolder(nombre) : null;
}

// ─── AUTORIZAR PERMISOS (ejecutar 1 vez) ──────────────────────────────────────
function autorizarPermisos() {
  Logger.log("Drive: "  + DriveApp.getRootFolder().getName());
  Logger.log("Sheets: " + SpreadsheetApp.create("test_delete_me").getId());
  Logger.log("OK");
}
