// ═══════════════════════════════════════════════════════════════════════════
// SISTEMA DE REPORTES HSE — Codigo_HSE.gs v1
// ═══════════════════════════════════════════════════════════════════════════

var EMAIL_HSE        = "tgpcontroldocumentos@gmail.com";
var CARPETA_RAIZ_HSE = "Reportes HSE 2026";

var TIPOS_HSE = [
  "Evaluacion ATS",
  "Evaluacion PTS",
  "Tarjetas STOP",
  "Inspecciones de Seguridad",
  "Check List de Conduccion"
];

var FRENTES_HSE = {
  "Costa_Geotecnia": [
    "SOPORTE A INGENIERIA",
    "VIAL",
    "URGENCIA VIAL KP 472+700 AL KP 482+000",
    "KP 714+155 AL KP 730+698",
    "MG KP 519+526 AL KP 540+839"
  ],
  "Sierra_Geotecnia": [
    "MG KP 170+000 AL KP 179+850",
    "MG KP 179+850 AL KP 194+000",
    "MG KP 194+000 AL KP 209+360",
    "APOYO A INGENIERIA"
  ],
  "Selva_Geotecnia": [
    "M.G. KP 43+830 - KP 53+000 - ETAPA 1",
    "M.G. KP 0+000 - KP 12+000 - ETAPA 1",
    "M.G. KP 25+000 - KP 35+000",
    "TAI KP 112+300",
    "Reparacion de F.O. KP61+150",
    "Perforaciones del KP126",
    "Perforaciones del KP55+118",
    "Apoyo / Ingenieria"
  ]
};

var SUPERVISORES_HSE = [
  { nombre: "CRISTHIAN BAQUERIZO" },
  { nombre: "WALTER JESUS" },
  { nombre: "LIZ GUERRERO" },
  { nombre: "CARLOS DE LA CRUZ" },
  { nombre: "ROY HERRADA" },
  { nombre: "ABEL SANCHEZ QUIHUI" },
  { nombre: "SAMUEL JARA MAYTA" },
  { nombre: "NIKOLAI ARANGOITIA" },
  { nombre: "ROGELIO CHAMPI CHOQUEPATA" },
  { nombre: "JHON FUENTES" },
  { nombre: "RUBEN NUÑEZ" },
  { nombre: "JORDAN GALLO" },
  { nombre: "ABRAHAM JIMENEZ" },
  { nombre: "NEISSER MAMANI" },
  { nombre: "PAUL PACSI ALAVE" },
  { nombre: "DANIEL ATAYUPANQUI TARCO" },
  { nombre: "CARLOS PUENTE" }
];

// ─── ROUTING ─────────────────────────────────────────────────────────────────
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("FormularioHSE")
    .setTitle("Reporte HSE")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── PROCESAR REPORTE ─────────────────────────────────────────────────────────
function procesarReporteHSE(datos) {
  try {
    var hoy   = Utilities.formatDate(new Date(), "America/Lima", "yyyy-MM-dd");
    var ahora = Utilities.formatDate(new Date(), "America/Lima", "yyyy-MM-dd HH:mm");
    var urlArchivo = "";

    if (datos.archivoBase64 && datos.archivoBase64.length > 100) {
      var carpeta   = obtenerCarpetaHSE(datos.sector, datos.subcategoria, datos.frente, datos.tipoReporte, hoy);
      var contenido = Utilities.base64Decode(datos.archivoBase64);
      var blob      = Utilities.newBlob(contenido, datos.mimeType, datos.nombreArchivo);
      var archivo   = carpeta.createFile(blob);
      archivo.setDescription(datos.responsable + " | " + datos.frente + " | " + datos.tipoReporte);
      urlArchivo = archivo.getUrl();
    }

    registrarEnSheetHSE(datos, urlArchivo, ahora, hoy);
    enviarEmailHSE(datos, urlArchivo, ahora);
    return { ok: true };

  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ─── CARPETAS DRIVE ───────────────────────────────────────────────────────────
function obtenerCarpetaHSE(sector, subcategoria, frente, tipo, fecha) {
  var r  = carpetaHSE(CARPETA_RAIZ_HSE, DriveApp.getRootFolder());
  var s  = carpetaHSE(sector,           r);
  var sc = carpetaHSE(subcategoria,     s);
  var fr = carpetaHSE(frente,           sc);
  var t  = carpetaHSE(tipo,             fr);
  return    carpetaHSE(fecha,           t);
}

function carpetaHSE(nombre, padre) {
  var i = padre.getFoldersByName(nombre);
  return i.hasNext() ? i.next() : padre.createFolder(nombre);
}

// ─── GOOGLE SHEET ─────────────────────────────────────────────────────────────
function registrarEnSheetHSE(datos, urlArchivo, ahora, fecha) {
  var ss   = abrirOCrearSheetHSE();
  var hoja = ss.getSheetByName("RAW_DATA");

  if (hoja.getLastRow() === 0) {
    hoja.appendRow([
      "Timestamp", "Fecha", "Responsable", "Puesto",
      "Sector", "Subcategoria", "Frente de Trabajo",
      "Tipo Reporte", "Nombre Archivo", "Link Archivo", "Semana"
    ]);
    hoja.getRange(1, 1, 1, 11)
      .setFontWeight("bold")
      .setBackground("#1A5276")
      .setFontColor("white");
    hoja.setFrozenRows(1);
  }

  var sem = Math.ceil(((new Date() - new Date(new Date().getFullYear(), 0, 1)) / 86400000
    + new Date(new Date().getFullYear(), 0, 1).getDay() + 1) / 7);

  hoja.appendRow([
    ahora,
    fecha,
    datos.responsable,
    datos.puesto        || "",
    datos.sector,
    datos.subcategoria,
    datos.frente        || "",
    datos.tipoReporte,
    datos.nombreArchivo || "",
    urlArchivo          || "",
    sem
  ]);
}

function abrirOCrearSheetHSE() {
  var nombre = CARPETA_RAIZ_HSE + " - Registros";
  var f = DriveApp.getFilesByName(nombre);
  if (f.hasNext()) return SpreadsheetApp.open(f.next());
  var ss = SpreadsheetApp.create(nombre);
  ss.getActiveSheet().setName("RAW_DATA");
  return ss;
}

// ─── EMAIL ────────────────────────────────────────────────────────────────────
function enviarEmailHSE(datos, urlArchivo, ahora) {
  var adjunto = urlArchivo
    ? '<br><a href="' + urlArchivo + '" style="color:#1A5276;font-weight:bold">Ver archivo en Drive</a>'
    : "";

  var asunto = "[HSE] " + datos.tipoReporte + " | " + datos.sector
    + " | " + datos.frente + " | " + datos.responsable;

  var cuerpo = '<div style="font-family:Arial,sans-serif;max-width:580px">'
    + '<div style="background:#1A5276;color:#fff;padding:16px;border-radius:8px 8px 0 0">'
    + '<b style="font-size:15px">Nuevo Reporte HSE</b><br>'
    + '<span style="font-size:11px;opacity:.8">' + ahora + '</span></div>'
    + '<div style="padding:16px;border:1px solid #ddd;border-top:0;border-radius:0 0 8px 8px;background:#fafafa">'
    + filaHSE("Tipo de Reporte",   datos.tipoReporte)
    + filaHSE("Responsable",       datos.responsable)
    + filaHSE("Puesto",            datos.puesto)
    + filaHSE("Sector",            datos.sector)
    + filaHSE("Subcategoria",      datos.subcategoria)
    + filaHSE("Frente de Trabajo", datos.frente)
    + filaHSE("Archivo",           datos.nombreArchivo)
    + adjunto
    + '</div></div>';

  GmailApp.sendEmail(EMAIL_HSE, asunto, "", { htmlBody: cuerpo });
}

function filaHSE(k, v) {
  return '<div style="padding:6px 0;border-bottom:1px solid #eee;font-size:13px">'
    + '<b style="color:#555">' + k + ':</b> ' + (v || "-") + '</div>';
}

// ─── DATOS PARA FRONTEND ──────────────────────────────────────────────────────
function obtenerSupervisoresHSE()          { return SUPERVISORES_HSE; }
function obtenerFrentesHSE(sector, subcat) { return FRENTES_HSE[sector + "_" + subcat] || []; }
function obtenerTiposHSE()                 { return TIPOS_HSE; }

// ─── AUTORIZAR PERMISOS ───────────────────────────────────────────────────────
function autorizarPermisosHSE() {
  Logger.log("Drive OK: "  + DriveApp.getRootFolder().getName());
  var drafts = GmailApp.getDrafts();
  Logger.log("Gmail OK: "  + drafts.length + " borradores");
}
