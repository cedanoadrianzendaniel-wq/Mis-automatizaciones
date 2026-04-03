// ═══════════════════════════════════════════════════════════════════════════
// PORTAL TGP — Bureau Veritas | Codigo_Portal.gs v2
// Contrato BV-TGP: 01/02/2026 — 31/01/2029
// ═══════════════════════════════════════════════════════════════════════════

var CLAVE_DEFAULT  = "tgp2026";
var SHEET_AVANCE   = "TGP - Avance Frentes";
var SHEET_TAREAS   = "TGP - Tareas";
var SHEET_CAMPO    = "Reportes de Campo 2026 - Registros";
var SHEET_HSE      = "Reportes HSE 2026 - Registros";
var CARPETA_DOCS   = "Portal TGP";
var FECHA_INICIO   = "2026-02-01";
var FECHA_FIN      = "2029-01-31";

// ─── FRENTES POR SECTOR ──────────────────────────────────────────────────────
var FRENTES_PORTAL = {
  "Costa": {
    subcategoria: "Geotecnia",
    frentes: [
      { nombre: "SOPORTE A INGENIERIA",                   cuenta: "63800018", orden: "TGGA/COS-26-02-01" },
      { nombre: "VIAL",                                   cuenta: "63800018", orden: "TG3CDV1"           },
      { nombre: "URGENCIA VIAL KP 472+700 AL KP 482+000", cuenta: "63440004", orden: "TGCI-2644"         },
      { nombre: "KP 714+155 AL KP 730+698",               cuenta: "63440002", orden: "TGEO-2629"         },
      { nombre: "MG KP 519+526 AL KP 540+839",            cuenta: "63440002", orden: "TGEO-2631"         }
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
      { nombre: "M.G. KP 43+830 - KP 53+000 - ETAPA 1", cuenta: "63800018", orden: "TGEO-2610"         },
      { nombre: "M.G. KP 0+000 - KP 12+000 - ETAPA 1",  cuenta: "63800018", orden: "TGEO-2606"         },
      { nombre: "M.G. KP 25+000 - KP 35+000",            cuenta: "63800018", orden: "TGEO-2608"         },
      { nombre: "TAI KP 112+300",                         cuenta: "63800018", orden: "TGEO-2601"         },
      { nombre: "Reparacion de F.O. KP61+150",           cuenta: "63800018", orden: "TG1CDV1"           },
      { nombre: "Perforaciones del KP126",               cuenta: "63800018", orden: "TGGA/SEL-25-08-07" },
      { nombre: "Perforaciones del KP55+118",            cuenta: "63800018", orden: "TGGA/SEL-26-02-01" },
      { nombre: "Apoyo / Ingenieria",                    cuenta: "63800018", orden: "TGGA/SEL-26-03-01" }
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
  var guardada = PropertiesService.getScriptProperties().getProperty("CLAVE_PORTAL") || CLAVE_DEFAULT;
  return String(clave).trim() === guardada;
}

// ═══════════════════════════════════════════════════════════════════════════
// TABLERO
// ═══════════════════════════════════════════════════════════════════════════
function obtenerKPIsTablero() {
  try {
    var hoy       = new Date();
    var inicio    = new Date(FECHA_INICIO);
    var fin       = new Date(FECHA_FIN);
    var totalDias = Math.round((fin - inicio) / 86400000);
    var diasTrans = Math.round((hoy - inicio) / 86400000);
    var diasRest  = Math.max(0, totalDias - diasTrans);
    var pctTiempo = Math.min(100, Math.round(diasTrans / totalDias * 100));

    // Avance promedio de frentes
    var totalFrente = 0, sumAvance = 0;
    ["Costa","Sierra","Selva"].forEach(function(s) {
      FRENTES_PORTAL[s].frentes.forEach(function(f) {
        sumAvance += obtenerUltimoAvance(s, f.nombre);
        totalFrente++;
      });
    });
    var avancePromedio = totalFrente > 0 ? Math.round(sumAvance / totalFrente) : 0;

    // Tareas
    var tareas     = leerTareas();
    var totalT     = tareas.length;
    var completadT = tareas.filter(function(t) { return t.estado === "Completado"; }).length;
    var pctTareas  = totalT > 0 ? Math.round(completadT / totalT * 100) : 0;

    // Alertas activas
    var alertas    = calcularAlertas(tareas);
    var numAlertas = alertas.length;

    // Cumplimiento HSE (últimos 30 días)
    var hse = calcularCumplimientoHSE();

    return {
      fechaInicio:     FECHA_INICIO,
      fechaFin:        FECHA_FIN,
      pctTiempo:       pctTiempo,
      diasTranscurridos: diasTrans,
      diasRestantes:   diasRest,
      totalFrentes:    totalFrente,
      avancePromedio:  avancePromedio,
      totalTareas:     totalT,
      tareasCompletadas: completadT,
      pctTareas:       pctTareas,
      alertasActivas:  numAlertas,
      cumplimientoHSE: hse.pct
    };
  } catch(e) {
    Logger.log("obtenerKPIsTablero: " + e);
    return { pctTiempo:0, diasTranscurridos:0, diasRestantes:0, totalFrentes:17,
             avancePromedio:0, pctTareas:0, alertasActivas:0, cumplimientoHSE:0 };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// SEGURIDAD (desde Reportes HSE)
// ═══════════════════════════════════════════════════════════════════════════
function obtenerResumenSeguridad() {
  try {
    var hse = calcularCumplimientoHSE();
    return hse;
  } catch(e) {
    Logger.log("obtenerResumenSeguridad: " + e);
    return { pct: 0, porTipo: {}, recientes: [] };
  }
}

function calcularCumplimientoHSE() {
  try {
    var ss = buscarSheet(SHEET_HSE);
    if (!ss) return { pct: 0, porTipo: {}, recientes: [] };
    var hoja = ss.getSheetByName("RAW_DATA");
    if (!hoja || hoja.getLastRow() <= 1) return { pct: 0, porTipo: {}, recientes: [] };

    var hoy    = new Date();
    var hace30 = new Date(hoy.getTime() - 30 * 86400000);
    var filas  = hoja.getRange(2, 1, hoja.getLastRow() - 1, 11).getValues();

    var porTipo  = {};
    var recientes = [];
    var total = 0, conArchivo = 0;

    filas.forEach(function(r) {
      var fecha = r[1] instanceof Date ? r[1] : new Date(r[1]);
      if (isNaN(fecha.getTime())) return;

      var tipo = String(r[7]);
      porTipo[tipo] = (porTipo[tipo] || 0) + 1;
      total++;
      if (r[9]) conArchivo++;

      if (fecha >= hace30) {
        recientes.push({
          fecha:       Utilities.formatDate(fecha, "America/Lima", "dd/MM/yyyy"),
          responsable: r[2],
          tipo:        tipo,
          sector:      r[4]
        });
      }
    });

    recientes.reverse();
    var pct = total > 0 ? Math.round(conArchivo / total * 100) : 0;
    return { pct: pct, total: total, conArchivo: conArchivo, porTipo: porTipo, recientes: recientes.slice(0, 10) };
  } catch(e) {
    return { pct: 0, porTipo: {}, recientes: [] };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// PROGRAMACIÓN — TAREAS
// ═══════════════════════════════════════════════════════════════════════════
function obtenerTodasTareas() {
  return leerTareas();
}

function obtenerTareasPorSemana(filtroSector, filtroFrente) {
  var tareas = leerTareas();

  if (filtroSector) tareas = tareas.filter(function(t) { return t.sector === filtroSector; });
  if (filtroFrente) tareas = tareas.filter(function(t) { return t.frente === filtroFrente; });

  // Ordenar por fecha inicio desc
  tareas.sort(function(a, b) { return b.fechaInicio.localeCompare(a.fechaInicio); });

  // Agrupar por semana
  var semanas = {};
  tareas.forEach(function(t) {
    var key = obtenerSemanaKey(t.fechaInicio);
    if (!semanas[key]) semanas[key] = { label: key, tareas: [] };
    semanas[key].tareas.push(t);
  });

  return Object.keys(semanas).sort(function(a,b){ return b.localeCompare(a); })
    .map(function(k) { return semanas[k]; });
}

function agregarTarea(datos) {
  try {
    var ss   = abrirOCrearSheetTareas();
    var hoja = ss.getSheetByName("Tareas");
    var id   = "T-" + new Date().getTime();
    var ts   = Utilities.formatDate(new Date(), "America/Lima", "yyyy-MM-dd HH:mm");

    hoja.appendRow([
      ts, id,
      datos.sector        || "",
      datos.frente        || "",
      datos.nombre        || "",
      datos.descripcion   || "",
      datos.fechaInicio   || "",
      datos.fechaFin      || "",
      parseFloat(datos.planCantidad) || 0,
      datos.planUnidad    || "",
      0,
      "Pendiente",
      datos.responsable   || ""
    ]);
    return { ok: true, id: id };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

function actualizarTarea(id, estado, ejecutadoCantidad) {
  try {
    var ss   = buscarSheet(SHEET_TAREAS);
    if (!ss) return { ok: false, error: "Sheet no encontrada" };
    var hoja = ss.getSheetByName("Tareas");
    if (!hoja || hoja.getLastRow() <= 1) return { ok: false, error: "Sin tareas" };

    var datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 13).getValues();
    for (var i = 0; i < datos.length; i++) {
      if (String(datos[i][1]) === String(id)) {
        var fila = i + 2;
        hoja.getRange(fila, 11).setValue(parseFloat(ejecutadoCantidad) || 0);
        hoja.getRange(fila, 12).setValue(estado);
        return { ok: true };
      }
    }
    return { ok: false, error: "Tarea no encontrada" };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

function eliminarTarea(id) {
  try {
    var ss   = buscarSheet(SHEET_TAREAS);
    if (!ss) return { ok: false };
    var hoja = ss.getSheetByName("Tareas");
    if (!hoja || hoja.getLastRow() <= 1) return { ok: false };

    var datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < datos.length; i++) {
      if (String(datos[i][1]) === String(id)) {
        hoja.deleteRow(i + 2);
        return { ok: true };
      }
    }
    return { ok: false };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// ALERTAS
// ═══════════════════════════════════════════════════════════════════════════
function obtenerAlertas() {
  try {
    var tareas = leerTareas();
    return calcularAlertas(tareas);
  } catch(e) {
    return [];
  }
}

function calcularAlertas(tareas) {
  var hoy     = new Date();
  hoy.setHours(0,0,0,0);
  var alertas = [];

  tareas.forEach(function(t) {
    if (t.estado === "Completado") return;
    if (!t.fechaFin) return;

    var fin = new Date(t.fechaFin);
    fin.setHours(0,0,0,0);
    if (fin < hoy) {
      var diasVencido = Math.round((hoy - fin) / 86400000);
      alertas.push({
        id:          t.id,
        tipo:        diasVencido > 7 ? "critical" : "warning",
        titulo:      'Tarea "' + t.nombre + '" — Plazo vencido',
        detalle:     t.frente + " · " + t.sector,
        vence:       t.fechaFin,
        diasVencido: diasVencido,
        estado:      t.estado
      });
    }
  });

  alertas.sort(function(a,b) { return b.diasVencido - a.diasVencido; });
  return alertas;
}

// ═══════════════════════════════════════════════════════════════════════════
// INFORMES — FRENTES + DOCUMENTOS + AVANCE
// ═══════════════════════════════════════════════════════════════════════════
function obtenerResumenSectores() {
  try {
    return ["Costa","Sierra","Selva"].map(function(sector) {
      var frentes        = FRENTES_PORTAL[sector].frentes;
      var avances        = frentes.map(function(f) { return obtenerUltimoAvance(sector, f.nombre); });
      var avancePromedio = Math.round(avances.reduce(function(a,b){ return a+b; }, 0) / avances.length);
      return {
        sector:          sector,
        numFrentes:      frentes.length,
        avancePromedio:  avancePromedio,
        ultimaActividad: obtenerUltimaActSector(sector)
      };
    });
  } catch(e) {
    return [
      { sector:"Costa",  numFrentes:5, avancePromedio:0, ultimaActividad:"-" },
      { sector:"Sierra", numFrentes:4, avancePromedio:0, ultimaActividad:"-" },
      { sector:"Selva",  numFrentes:8, avancePromedio:0, ultimaActividad:"-" }
    ];
  }
}

function obtenerFrentesConAvance(sector) {
  var config = FRENTES_PORTAL[sector];
  if (!config) return [];
  return config.frentes.map(function(f) {
    return {
      nombre:          f.nombre,
      cuenta:          f.cuenta,
      orden:           f.orden,
      avance:          obtenerUltimoAvance(sector, f.nombre),
      ultimaActividad: obtenerUltimaActFrente(sector, f.nombre)
    };
  });
}

function obtenerDetalleFrente(sector, frente) {
  var config     = FRENTES_PORTAL[sector] || { subcategoria:"", frentes:[] };
  var frenteInfo = null;
  config.frentes.forEach(function(f) { if (f.nombre === frente) frenteInfo = f; });
  return {
    sector: sector, subcategoria: config.subcategoria, nombre: frente,
    cuenta: frenteInfo ? frenteInfo.cuenta : "",
    orden:  frenteInfo ? frenteInfo.orden  : "",
    avanceActual: obtenerUltimoAvance(sector, frente)
  };
}

function obtenerCurvaAvance(sector, frente) {
  try {
    var ss = buscarSheet(SHEET_AVANCE);
    if (!ss) return [];
    var hoja = ss.getSheetByName("Avance");
    if (!hoja || hoja.getLastRow() <= 1) return [];
    var datos  = hoja.getRange(2, 1, hoja.getLastRow()-1, 7).getValues();
    var puntos = [];
    datos.forEach(function(r) {
      var f = r[1] instanceof Date
        ? Utilities.formatDate(r[1], "America/Lima", "yyyy-MM-dd")
        : String(r[1]).substring(0,10);
      if (String(r[2]).toLowerCase() === sector.toLowerCase() && String(r[3]) === frente)
        puntos.push({ fecha: f, avance: parseFloat(r[4])||0, obs: String(r[5]) });
    });
    puntos.sort(function(a,b){ return a.fecha.localeCompare(b.fecha); });
    return puntos;
  } catch(e) { return []; }
}

function actualizarAvanceFrente(sector, frente, porcentaje, observacion, usuario) {
  try {
    var ss   = abrirOCrearSheetAvance();
    var hoja = ss.getSheetByName("Avance");
    var ahora = Utilities.formatDate(new Date(), "America/Lima", "yyyy-MM-dd HH:mm");
    var hoy   = Utilities.formatDate(new Date(), "America/Lima", "yyyy-MM-dd");
    hoja.appendRow([ahora, hoy, sector, frente, parseFloat(porcentaje)||0, observacion||"", usuario||"Portal"]);
    return { ok: true };
  } catch(e) { return { ok: false, error: e.toString() }; }
}

function obtenerDocumentosFrente(sector, frente) {
  try {
    var carpeta = obtenerCarpetaDocs(sector, frente, false);
    if (!carpeta) return [];
    var docs = [], files = carpeta.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      docs.push({
        nombre: f.getName(),
        url:    f.getUrl(),
        fecha:  Utilities.formatDate(f.getDateCreated(), "America/Lima", "dd/MM/yyyy"),
        tamano: Math.round(f.getSize()/1024) + " KB"
      });
    }
    docs.reverse();
    return docs;
  } catch(e) { return []; }
}

function subirDocumentoFrente(datos) {
  try {
    var carpeta   = obtenerCarpetaDocs(datos.sector, datos.frente, true);
    var contenido = Utilities.base64Decode(datos.archivoBase64);
    var blob      = Utilities.newBlob(contenido, datos.mimeType, datos.nombreArchivo);
    var archivo   = carpeta.createFile(blob);
    archivo.setDescription(datos.sector + " | " + datos.frente);
    return { ok: true, url: archivo.getUrl(), nombre: archivo.getName() };
  } catch(e) { return { ok: false, error: e.toString() }; }
}

function obtenerActividadesFrente(sector, frente) {
  try {
    var ss = buscarSheet(SHEET_CAMPO);
    if (!ss) return [];
    var hoja = ss.getSheetByName("RAW_DATA");
    if (!hoja || hoja.getLastRow() <= 1) return [];
    var filas = hoja.getRange(2, 1, hoja.getLastRow()-1, 23).getValues();
    var result = [];
    filas.forEach(function(r) {
      if (String(r[4]).toLowerCase() === sector.toLowerCase() && String(r[6]) === frente) {
        var f = r[1] instanceof Date
          ? Utilities.formatDate(r[1], "America/Lima", "dd/MM/yyyy")
          : String(r[1]).substring(0,10);
        result.push({ fecha: f, responsable: r[2], puesto: r[3], descripcion: r[16]||"", avance: r[17] ? parseFloat(r[17]).toFixed(1)+"%" : "-" });
      }
    });
    result.reverse();
    return result.slice(0, 15);
  } catch(e) { return []; }
}

// ═══════════════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════════════
function leerTareas() {
  try {
    var ss = buscarSheet(SHEET_TAREAS);
    if (!ss) return [];
    var hoja = ss.getSheetByName("Tareas");
    if (!hoja || hoja.getLastRow() <= 1) return [];
    var datos = hoja.getRange(2, 1, hoja.getLastRow()-1, 13).getValues();
    return datos.map(function(r) {
      return {
        timestamp:      String(r[0]),
        id:             String(r[1]),
        sector:         String(r[2]),
        frente:         String(r[3]),
        nombre:         String(r[4]),
        descripcion:    String(r[5]),
        fechaInicio:    r[6] instanceof Date ? Utilities.formatDate(r[6],"America/Lima","yyyy-MM-dd") : String(r[6]).substring(0,10),
        fechaFin:       r[7] instanceof Date ? Utilities.formatDate(r[7],"America/Lima","yyyy-MM-dd") : String(r[7]).substring(0,10),
        planCantidad:   parseFloat(r[8])  || 0,
        planUnidad:     String(r[9]),
        ejecutado:      parseFloat(r[10]) || 0,
        estado:         String(r[11]),
        responsable:    String(r[12])
      };
    });
  } catch(e) { return []; }
}

function obtenerSemanaKey(fechaStr) {
  try {
    var d   = new Date(fechaStr + "T12:00:00");
    var dia = d.getDay();
    var lun = new Date(d);
    lun.setDate(d.getDate() - (dia === 0 ? 6 : dia - 1));
    var dom = new Date(lun);
    dom.setDate(lun.getDate() + 6);
    var meses = ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"];
    return lun.getDate() + " " + meses[lun.getMonth()] + " – " + dom.getDate() + " " + meses[dom.getMonth()] + " " + dom.getFullYear();
  } catch(e) { return fechaStr; }
}

function obtenerUltimoAvance(sector, frente) {
  try {
    var ss = buscarSheet(SHEET_AVANCE);
    if (!ss) return 0;
    var hoja = ss.getSheetByName("Avance");
    if (!hoja || hoja.getLastRow() <= 1) return 0;
    var datos = hoja.getRange(2, 1, hoja.getLastRow()-1, 7).getValues();
    var ultimo = 0;
    datos.forEach(function(r) {
      if (String(r[2]).toLowerCase() === sector.toLowerCase() && String(r[3]) === frente)
        ultimo = parseFloat(r[4]) || 0;
    });
    return ultimo;
  } catch(e) { return 0; }
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
        return r[1] instanceof Date ? Utilities.formatDate(r[1],"America/Lima","dd/MM/yyyy") : String(r[1]).substring(0,10);
      }
    }
    return "-";
  } catch(e) { return "-"; }
}

function obtenerUltimaActFrente(sector, frente) {
  try {
    var ss = buscarSheet(SHEET_CAMPO);
    if (!ss) return "Sin actividad";
    var hoja = ss.getSheetByName("RAW_DATA");
    if (!hoja || hoja.getLastRow() <= 1) return "Sin actividad";
    for (var i = hoja.getLastRow(); i >= 2; i--) {
      var r = hoja.getRange(i, 1, 1, 7).getValues()[0];
      if (String(r[4]).toLowerCase() === sector.toLowerCase() && String(r[6]) === frente)
        return r[1] instanceof Date ? Utilities.formatDate(r[1],"America/Lima","dd/MM/yyyy") : String(r[1]).substring(0,10);
    }
    return "Sin actividad";
  } catch(e) { return "-"; }
}

function buscarSheet(nombre) {
  var f = DriveApp.getFilesByName(nombre);
  return f.hasNext() ? SpreadsheetApp.open(f.next()) : null;
}

function abrirOCrearSheetAvance() {
  var f = DriveApp.getFilesByName(SHEET_AVANCE);
  if (f.hasNext()) return SpreadsheetApp.open(f.next());
  var ss = SpreadsheetApp.create(SHEET_AVANCE);
  var h  = ss.getActiveSheet(); h.setName("Avance");
  h.appendRow(["Timestamp","Fecha","Sector","Frente","% Avance","Observacion","Usuario"]);
  h.getRange(1,1,1,7).setFontWeight("bold").setBackground("#1F3864").setFontColor("white");
  h.setFrozenRows(1);
  return ss;
}

function abrirOCrearSheetTareas() {
  var f = DriveApp.getFilesByName(SHEET_TAREAS);
  if (f.hasNext()) return SpreadsheetApp.open(f.next());
  var ss = SpreadsheetApp.create(SHEET_TAREAS);
  var h  = ss.getActiveSheet(); h.setName("Tareas");
  h.appendRow(["Timestamp","ID","Sector","Frente","Nombre","Descripcion","FechaInicio","FechaFin","PlanCantidad","PlanUnidad","Ejecutado","Estado","Responsable"]);
  h.getRange(1,1,1,13).setFontWeight("bold").setBackground("#1F3864").setFontColor("white");
  h.setFrozenRows(1);
  return ss;
}

function obtenerCarpetaDocs(sector, frente, crear) {
  var raiz = buscarOCrear(CARPETA_DOCS,  DriveApp.getRootFolder(), crear);
  if (!raiz) return null;
  var sec  = buscarOCrear(sector,         raiz,                    crear);
  if (!sec)  return null;
  return     buscarOCrear(frente,         sec,                     crear);
}

function buscarOCrear(nombre, padre, crear) {
  var i = padre.getFoldersByName(nombre);
  if (i.hasNext()) return i.next();
  return crear ? padre.createFolder(nombre) : null;
}

function autorizarPermisos() {
  Logger.log("Drive: "  + DriveApp.getRootFolder().getName());
  Logger.log("Gmail: "  + GmailApp.getDrafts().length);
  Logger.log("Sheets OK");
}
