// ═══════════════════════════════════════════════════════════════════════════
// GENERADOR DE PARTE DIARIO DE SERVICIOS (PDS) — v1
// Se agrega al proyecto de Reportes de Campo (mismo Apps Script)
// ═══════════════════════════════════════════════════════════════════════════

// ─── IDs DE IMÁGENES EN DRIVE ────────────────────────────────────────────────
var ID_LOGO_TGP    = "1xzkhhOpG1tPwUvMTqiEz-y5pgjXufVTB";
var ID_FIRMA_YURI  = "1zhEKAF6qQrSJBn-U5MiIT6Bk-Zr6H19V";

// ─── TABLA CUENTA / ORDEN POR FRENTE ─────────────────────────────────────────
var CUENTAS_ORDENES = {
  // COSTA
  "SOPORTE A INGENIERIA":                   { cuenta: "63800018", orden: "TGGA/COS-26-02-01"  },
  "VIAL":                                   { cuenta: "63800018", orden: "TG3CDV1"             },
  "URGENCIA VIAL KP 472+700 AL KP 482+000": { cuenta: "63440004", orden: "TGCI-2644"          },
  "KP 714+155 AL KP 730+698":               { cuenta: "63440002", orden: "TGEO-2629"          },
  "MG KP 519+526 AL KP 540+839":            { cuenta: "63440002", orden: "TGEO-2631"          },
  // SIERRA
  "MG KP 170+000 AL KP 179+850":            { cuenta: "63440002", orden: "TGEO-2621"          },
  "MG KP 179+850 AL KP 194+000":            { cuenta: "63440002", orden: "TGEO-2622"          },
  "MG KP 194+000 AL KP 209+360":            { cuenta: "63440002", orden: "TGEO-2623"          },
  "APOYO A INGENIERIA":                      { cuenta: "63290002", orden: "TGGA/SIE-26-02-02" },
  // SELVA
  "M.G. KP 43+830 - KP 53+000 - ETAPA 1":  { cuenta: "63800018", orden: "TGEO-2610"          },
  "M.G. KP 0+000 - KP 12+000 - ETAPA 1":   { cuenta: "63800018", orden: "TGEO-2606"          },
  "M.G. KP 25+000 - KP 35+000":             { cuenta: "63800018", orden: "TGEO-2608"          },
  "TAI KP 112+300":                          { cuenta: "63800018", orden: "TGEO-2601"          },
  "Reparacion de F.O. KP61+150":            { cuenta: "63800018", orden: "TG1CDV1"            },
  "Perforaciones del KP126":                { cuenta: "63800018", orden: "TGGA/SEL-25-08-07"  },
  "Perforaciones del KP55+118":             { cuenta: "63800018", orden: "TGGA/SEL-26-02-01"  },
  "Apoyo / Ingenieria":                     { cuenta: "63800018", orden: "TGGA/SEL-26-03-01"  }
};

// ─── COLORES PDS ──────────────────────────────────────────────────────────────
var PDS_AZUL_OSCURO  = "#1A5276";
var PDS_AZUL_MEDIO  = "#1F6B8E";
var PDS_AZUL_COL    = "#2E86AB";
var PDS_AZUL_CLARO  = "#D6EAF8";
var PDS_GRIS_FILA   = "#EBF5FB";
var PDS_BLANCO      = "#FFFFFF";

// ─── INTEGRACIÓN: llamar desde procesarReporte() ──────────────────────────────
// En tu Codigo.gs, dentro de procesarReporte(), AGREGA esta línea
// justo después de enviarEmail(...):
//
//   generarPDS(datos.sector, datos.fecha || hoy);
//
// ─────────────────────────────────────────────────────────────────────────────

// ─── FUNCIÓN PRINCIPAL ────────────────────────────────────────────────────────
// Genera o actualiza el tab del día dentro del archivo mensual del sector
function generarPDS(sector, fecha) {
  try {
    var reportes = obtenerReportesPorSectorFecha(sector, fecha);
    if (reportes.length === 0) return null;

    // Carpeta: Parte Diario / Sector
    var carpeta = obtenerCarpetaPDS(sector);

    // Archivo mensual: PDS_Costa_2026-04
    var mes    = fecha.substring(0, 7); // "2026-04"
    var nombre = "PDS_" + sector + "_" + mes;

    // Buscar archivo existente o crear nuevo
    var pdsSS;
    var existentes = carpeta.getFilesByName(nombre);
    if (existentes.hasNext()) {
      pdsSS = SpreadsheetApp.open(existentes.next());
    } else {
      pdsSS    = SpreadsheetApp.create(nombre);
      var file = DriveApp.getFileById(pdsSS.getId());
      carpeta.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    }

    // Tab por fecha: "2026-04-01"
    var tabNombre = fecha; // formato simple yyyy-MM-dd sin caracteres especiales
    var pdsHoja   = pdsSS.getSheetByName(tabNombre);
    if (pdsHoja) {
      // Eliminar y recrear para evitar conflictos con imágenes previas
      pdsSS.deleteSheet(pdsHoja);
    }
    pdsHoja = pdsSS.insertSheet(tabNombre, 0); // posicion 0 = primera pestaña

    // Eliminar la hoja por defecto "Hoja 1" / "Sheet1" si existe
    ["Hoja 1", "Sheet1", "Hoja1"].forEach(function(n) {
      var h = pdsSS.getSheetByName(n);
      if (h && pdsSS.getSheets().length > 1) pdsSS.deleteSheet(h);
    });

    construirPDS(pdsHoja, sector, fecha, reportes);
    SpreadsheetApp.flush();
    return pdsSS.getUrl();

  } catch (err) {
    Logger.log("Error generarPDS: " + err.toString());
    return null;
  }
}

// ─── REGENERAR HISTÓRICOS ─────────────────────────────────────────────────────
// Ejecutar una sola vez para generar PDS de todos los días anteriores
function regenerarTodosPDS() {
  var ss   = abrirOCrearSheet();
  var hoja = ss.getSheetByName("RAW_DATA");
  if (!hoja || hoja.getLastRow() <= 1) { Logger.log("Sin datos"); return; }

  var filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 10).getValues();
  var combinaciones = {};

  filas.forEach(function(r) {
    var fecha = r[1] instanceof Date
      ? Utilities.formatDate(r[1], "America/Lima", "yyyy-MM-dd")
      : String(r[1]).substring(0, 10);
    var sector = String(r[4]).trim();
    if (fecha && sector) combinaciones[sector + "|" + fecha] = { sector: sector, fecha: fecha };
  });

  var keys = Object.keys(combinaciones);
  Logger.log("Combinaciones encontradas: " + keys.length);
  keys.forEach(function(k) {
    var c = combinaciones[k];
    var url = generarPDS(c.sector, c.fecha);
    Logger.log(c.sector + " | " + c.fecha + " → " + (url || "sin datos"));
  });
  Logger.log("Regeneracion completada");
}

// ─── CARPETA PARTE DIARIO ─────────────────────────────────────────────────────
function obtenerCarpetaPDS(sector) {
  var raiz = carpetaEnPadre("Parte Diario", DriveApp.getRootFolder());
  return carpetaEnPadre(sector, raiz);
}

// ─── LEER RAW_DATA FILTRADO ───────────────────────────────────────────────────
function obtenerReportesPorSectorFecha(sector, fecha) {
  var ss   = abrirOCrearSheet();
  var hoja = ss.getSheetByName("RAW_DATA");
  if (!hoja || hoja.getLastRow() <= 1) return [];

  var filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 23).getValues();

  return filas.filter(function (r) {
    // r[1] = Fecha, r[4] = Sector
    var fechaFila = r[1] instanceof Date
      ? Utilities.formatDate(r[1], "America/Lima", "yyyy-MM-dd")
      : String(r[1]).substring(0, 10);
    return fechaFila === fecha && String(r[4]).toLowerCase() === sector.toLowerCase();
  });
}

// ─── CONSTRUCCIÓN DEL PDS ─────────────────────────────────────────────────────
function construirPDS(hoja, sector, fecha, reportes) {

  // Formatear fecha legible: "sáb 28-Feb-26"
  var fechaFormateada = formatearFechaPDS(fecha);

  var f = 1; // fila actual

  // ── TÍTULO con logo TGP ──────────────────────────────────────────────────
  estiloBloque(hoja, f, 1, 1, 10, "", PDS_AZUL_OSCURO, 14, true, "center");
  hoja.getRange(f, 3, 1, 6).merge()
    .setValue("PARTE DIARIO DE SERVICIOS")
    .setBackground(PDS_AZUL_OSCURO)
    .setFontColor(PDS_BLANCO)
    .setFontSize(14).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  hoja.setRowHeight(f, 50);
  // Logo TGP pendiente de configurar
  f++;

  // ── DATOS GENERALES ───────────────────────────────────────────────────────
  estiloBloque(hoja, f, 1, 1, 10, "DATOS GENERALES", PDS_AZUL_MEDIO, 10, true, "center");
  hoja.setRowHeight(f, 20);
  f++;

  // Etiquetas
  hoja.getRange(f, 1, 1, 4).merge().setValue("Empresa colaboradora:").setFontWeight("bold").setBackground(PDS_AZUL_CLARO);
  hoja.getRange(f, 5, 1, 3).merge().setValue("Zona:").setFontWeight("bold").setBackground(PDS_AZUL_CLARO);
  hoja.getRange(f, 8, 1, 3).merge().setValue("Fecha:").setFontWeight("bold").setBackground(PDS_AZUL_CLARO);
  hoja.setRowHeight(f, 18);
  f++;

  // Valores
  hoja.getRange(f, 1, 1, 4).merge().setValue("BUREAU VERITAS DEL PERÚ")
    .setFontWeight("bold").setHorizontalAlignment("center");
  hoja.getRange(f, 5, 1, 3).merge().setValue(sector.toUpperCase())
    .setFontWeight("bold").setHorizontalAlignment("center");
  hoja.getRange(f, 8, 1, 3).merge().setValue(fechaFormateada)
    .setHorizontalAlignment("center");
  hoja.setRowHeight(f, 20);
  f++;

  // ── PERSONAL ──────────────────────────────────────────────────────────────
  estiloBloque(hoja, f, 1, 1, 10, "PERSONAL", PDS_AZUL_MEDIO, 10, true, "center");
  hoja.setRowHeight(f, 20);
  f++;

  // Headers personal
  var hPersonal = [["N","NOMBRE","","","ORIGEN","DESTINO","SERVICIO","ACTIVIDAD","CUENTA","ORDEN"]];
  hoja.getRange(f, 1, 1, 10).setValues(hPersonal)
    .setBackground(PDS_AZUL_COL).setFontColor(PDS_BLANCO)
    .setFontWeight("bold").setHorizontalAlignment("center");
  hoja.getRange(f, 2, 1, 3).merge();
  hoja.setRowHeight(f, 20);
  f++;

  // Filas de personal
  reportes.forEach(function (r, idx) {
    var co  = CUENTAS_ORDENES[r[6]] || { cuenta: "", orden: "" };
    var bg  = idx % 2 === 1 ? PDS_GRIS_FILA : PDS_BLANCO;

    hoja.getRange(f, 1).setValue(idx + 1).setHorizontalAlignment("center");
    hoja.getRange(f, 2, 1, 3).merge().setValue(r[2]); // Nombre
    hoja.getRange(f, 5).setValue(r[11]);               // Origen
    hoja.getRange(f, 6).setValue(r[12]);               // Destino
    hoja.getRange(f, 7).setValue(r[3]);                // Servicio = Puesto
    hoja.getRange(f, 8).setValue(r[6]);                // Actividad = Frente
    hoja.getRange(f, 9).setValue(co.cuenta).setHorizontalAlignment("center");
    hoja.getRange(f, 10).setValue(co.orden).setHorizontalAlignment("center");
    hoja.getRange(f, 1, 1, 10).setBackground(bg);
    hoja.setRowHeight(f, 20);
    f++;
  });

  // ── EQUIPOS ───────────────────────────────────────────────────────────────
  estiloBloque(hoja, f, 1, 1, 10, "EQUIPOS", PDS_AZUL_MEDIO, 10, true, "center");
  hoja.setRowHeight(f, 20);
  f++;

  // Headers equipos
  var hEquipos = [["N","EQUIPO","KM INICIO","KM FIN","ORIGEN","DESTINO","SERVICIO","ACTIVIDAD","CUENTA","ORDEN"]];
  hoja.getRange(f, 1, 1, 10).setValues(hEquipos)
    .setBackground(PDS_AZUL_COL).setFontColor(PDS_BLANCO)
    .setFontWeight("bold").setHorizontalAlignment("center");
  hoja.setRowHeight(f, 20);
  f++;

  // Filas de equipos (solo quienes tienen camioneta)
  var numEquipos = 0;
  reportes.forEach(function (r) {
    if (String(r[7]).toLowerCase() !== "si" || !r[8]) return; // sin camioneta

    var co      = CUENTAS_ORDENES[r[6]] || { cuenta: "", orden: "" };
    var kmIni   = parseFloat(r[9])  || 0;
    var kmFin   = parseFloat(r[10]) || 0;
    var recorr  = Math.max(0, kmFin - kmIni).toFixed(1);
    var bg      = numEquipos % 2 === 1 ? PDS_GRIS_FILA : PDS_BLANCO;

    numEquipos++;
    hoja.getRange(f, 1).setValue(numEquipos).setHorizontalAlignment("center");
    hoja.getRange(f, 2).setValue("CAMIONETA " + String(r[8]).toUpperCase());
    hoja.getRange(f, 3).setValue(kmIni).setHorizontalAlignment("center");
    hoja.getRange(f, 4).setValue(kmFin).setHorizontalAlignment("center");
    hoja.getRange(f, 5).setValue(r[11]);  // Origen
    hoja.getRange(f, 6).setValue(r[12]);  // Destino
    hoja.getRange(f, 7).setValue("ALQUILER DE CAMIONETA SIN CONDUCTOR");
    hoja.getRange(f, 8).setValue(r[6]);   // Actividad = Frente
    hoja.getRange(f, 9).setValue(co.cuenta).setHorizontalAlignment("center");
    hoja.getRange(f, 10).setValue(co.orden).setHorizontalAlignment("center");
    hoja.getRange(f, 1, 1, 10).setBackground(bg);
    hoja.setRowHeight(f, 20);
    f++;
  });

  if (numEquipos === 0) {
    hoja.getRange(f, 1, 1, 10).merge()
      .setValue("Sin equipos registrados")
      .setFontStyle("italic").setFontColor("#999999")
      .setHorizontalAlignment("center");
    hoja.setRowHeight(f, 20);
    f++;
  }

  // ── COMENTARIOS ───────────────────────────────────────────────────────────
  estiloBloque(hoja, f, 1, 1, 10, "COMENTARIOS", PDS_AZUL_MEDIO, 10, true, "center");
  hoja.setRowHeight(f, 20);
  f++;

  var texto = generarComentarios(reportes, numEquipos);
  var numLineas = Math.max(6, texto.split("\n").length + 1);
  hoja.getRange(f, 1, numLineas, 10).merge()
    .setValue(texto)
    .setVerticalAlignment("top")
    .setWrap(true)
    .setFontSize(10);
  for (var i = 0; i < numLineas; i++) hoja.setRowHeight(f + i, 18);
  f += numLineas;

  // ── FIRMAS ────────────────────────────────────────────────────────────────
  f++; // espacio
  estiloBloque(hoja, f, 1, 1, 5, "REPRESENTANTE BUREAU VERITAS", PDS_AZUL_MEDIO, 10, true, "center");
  estiloBloque(hoja, f, 6, 1, 5, "REPRESENTANTE DE TGP",          PDS_AZUL_MEDIO, 10, true, "center");
  hoja.setRowHeight(f, 20);
  f++;

  hoja.getRange(f, 1, 1, 5).merge().setValue("Firma:").setFontWeight("bold");
  hoja.getRange(f, 6, 1, 5).merge().setValue("Firma:").setFontWeight("bold");
  f++;

  hoja.setRowHeight(f, 60); // espacio para firma manual
  f++;

  // Línea punteada y datos del firmante
  hoja.getRange(f, 1, 1, 5).merge()
    .setValue("........................................\nIng. Yuri Arangoitia Rendon\nJefe de Supervisión (BV)\nReg. CIP N° 206381")
    .setWrap(true).setFontSize(9).setHorizontalAlignment("center");
  hoja.getRange(f, 6, 1, 5).merge().setValue("Nombre:").setFontWeight("bold");
  hoja.setRowHeight(f, 60);

  // ── ANCHOS DE COLUMNA ─────────────────────────────────────────────────────
  hoja.setColumnWidth(1,  40);   // N
  hoja.setColumnWidth(2,  180);  // NOMBRE / EQUIPO
  hoja.setColumnWidth(3,  75);   // KM INI
  hoja.setColumnWidth(4,  75);   // KM FIN
  hoja.setColumnWidth(5,  90);   // ORIGEN
  hoja.setColumnWidth(6,  90);   // DESTINO
  hoja.setColumnWidth(7,  230);  // SERVICIO
  hoja.setColumnWidth(8,  190);  // ACTIVIDAD
  hoja.setColumnWidth(9,  85);   // CUENTA
  hoja.setColumnWidth(10, 130);  // ORDEN

  // ── BORDES GENERALES ──────────────────────────────────────────────────────
  hoja.getRange(1, 1, f, 10)
    .setBorder(true, true, true, true, true, true,
               "#AAAAAA", SpreadsheetApp.BorderStyle.SOLID);

  SpreadsheetApp.flush();
}

// ─── TEXTO DE COMENTARIOS ─────────────────────────────────────────────────────
function generarComentarios(reportes, numEquipos) {
  var numSup  = reportes.length;
  var numAlim = reportes.filter(function(r) { return String(r[13]).toLowerCase() === "si"; }).length;
  var numHosp = reportes.filter(function(r) { return String(r[14]).toLowerCase() === "si"; }).length;

  var lineas = [];
  lineas.push(
    numSup + " Supervisor" + (numSup > 1 ? "es" : "") +
    (numEquipos > 0 ? ", " + numEquipos + " Camioneta" + (numEquipos > 1 ? "s" : "") : "")
  );
  lineas.push("");

  reportes.forEach(function(r) {
    if (r[8]) lineas.push("- Camioneta " + r[8] + " a servicio de " + r[2]);
  });

  lineas.push("");
  lineas.push("Actividades:");
  lineas.push("");

  reportes.forEach(function(r) {
    lineas.push(r[2]);
    var desc = r[16] || ("Actividades en " + r[6]);
    lineas.push("1.- " + desc);
    if (r[18]) lineas.push("   Observaciones: " + r[18]);
    lineas.push("");
  });

  lineas.push("GASTOS LOGÍSTICOS:");
  lineas.push("- ALIMENTACION: " + String(numAlim).padStart(2, "0"));
  lineas.push("- HOSPEDAJE: "    + String(numHosp).padStart(2, "0"));

  return lineas.join("\n");
}

// ─── HELPERS DE ESTILO ────────────────────────────────────────────────────────
function estiloBloque(hoja, fila, col, numFilas, numCols, texto, bg, fontSize, bold, align) {
  var r = hoja.getRange(fila, col, numFilas, numCols);
  if (numCols > 1 || numFilas > 1) r.merge();
  r.setValue(texto)
   .setBackground(bg)
   .setFontColor(PDS_BLANCO)
   .setFontSize(fontSize || 10)
   .setFontWeight(bold ? "bold" : "normal")
   .setHorizontalAlignment(align || "left")
   .setVerticalAlignment("middle");
}

function formatearFechaPDS(fechaStr) {
  try {
    var partes = fechaStr.split("-");
    var d = new Date(
      parseInt(partes[0]),
      parseInt(partes[1]) - 1,
      parseInt(partes[2])
    );
    var dias   = ["dom","lun","mar","mié","jue","vie","sáb"];
    var meses  = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
    return dias[d.getDay()] + " " + d.getDate() + "-" + meses[d.getMonth()] + "-" + String(d.getFullYear()).slice(2);
  } catch(e) {
    return fechaStr;
  }
}
