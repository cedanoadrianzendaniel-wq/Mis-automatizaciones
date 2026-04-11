// ═══════════════════════════════════════════════════════════════════════════
// PDS GENERATOR — Parte Diario de Servicios (Excel)
// Genera/actualiza el PDS por Sector y Mes desde datos de RAW_DATA
// ═══════════════════════════════════════════════════════════════════════════

const ExcelJS = require("exceljs");
const { Readable } = require("stream");
const path = require("path");
const fs = require("fs");

// ─── COLORES ─────────────────────────────────────────────────────────────────
const BLUE_HEADER = "00709C";
const WHITE = "FFFFFF";
const BORDER_THIN = { style: "thin", color: { argb: "FF000000" } };
const ALL_BORDERS = { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN };

// ─── MAPEO: Frente → Cuenta / Orden (Geotecnia) ─────────────────────────────
const FRENTES_MAPPING = {
  // SELVA
  "M.G. KP 43+830 - KP 53+000 - ETAPA 1": { cuenta: "63800018", orden: "TGEO-2610" },
  "M.G. KP 0+000 - KP 12+000 - ETAPA 1":  { cuenta: "63800018", orden: "TGEO-2606" },
  "M.G. KP 25+000 - KP 35+000":            { cuenta: "63800018", orden: "TGEO-2601" },
  "TAI KP 112+300":                          { cuenta: "63800018", orden: "TGEO-2629" },
  "Reparacion de F.O. KP61+150":             { cuenta: "63800018", orden: "TGTCDV1" },
  "Perforaciones del KP126":                 { cuenta: "63800018", orden: "TGASEL-25-08-07" },
  "Perforaciones del KP55+118":              { cuenta: "63800018", orden: "TGASEL-26-02-01" },
  "Apoyo / Ingenieria":                      { cuenta: "63800018", orden: "TGASEL-26-03-01" },
  // COSTA
  "SOPORTE A INGENIERIA":                    { cuenta: "63800018", orden: "TOGA/COS-26-02-01" },
  "VIAL":                                    { cuenta: "63800018", orden: "TGEO-2871" },
  "URGENCIA VIAL KP 472+700 AL KP 482+000": { cuenta: "63440002", orden: "TGEO-2644" },
  "KP 714+155 AL KP 730+698":               { cuenta: "8344002",  orden: "TGEO-2629" },
  "MG KP 519+526 AL KP 540+839":            { cuenta: "63440002", orden: "TGEO-2857" },
  // SIERRA
  "MG KP 170+000 AL KP 179+850":            { cuenta: "63440002", orden: "TGEO-2821" },
  "MG KP 179+850 AL KP 194+000":            { cuenta: "63440002", orden: "TGEO-2822" },
  "MG KP 194+000 AL KP 209+360":            { cuenta: "63440002", orden: "TGEO-2823" },
  "APOYO A INGENIERIA":                      { cuenta: "63290002", orden: "TOGA/SIE-26-02-02" }
};

// ─── MAPEO: Proyecto CAPEX → Cuenta / Orden / ElementoPEP ───────────────────
const CAPEX_MAPPING = {
  // Proyectos con Elemento PEP (sin cuenta/orden tradicional)
  "Protecciones mecanicas ductos NG/NG":                  { elementoPEP: "TGPY/OPE-1501-2-2", cuenta: "", orden: "" },
  "Mejoras Skids Gas combustible en PS's":                { elementoPEP: "TGPY/OPE-1901-2-4", cuenta: "", orden: "" },
  "Construccion nuevas instalaciones Lurin":              { elementoPEP: "TGPY/OPE-1902-2-2", cuenta: "", orden: "" },
  "Mejora instalaciones Aerodromo Kiteni":                { elementoPEP: "TGPY/OPE-2101-1-2", cuenta: "", orden: "" },
  "Actualizacion Sistema de Automatizacion":              { elementoPEP: "TGPY/OPE-2201-2-4", cuenta: "", orden: "" },
  "Cambio Tableros / Luminarias Areas Clasificadas":      { elementoPEP: "TGPY/OPE-2301-2-3", cuenta: "", orden: "" },
  "Plan Mitigacion Ruido PC Kamani (venteo)":             { elementoPEP: "TGPY/OPE-2302-2-4", cuenta: "", orden: "" },
  "Adecuacion Valvula Sobrepresion NG32 PS1":             { elementoPEP: "TGPY/OPE-2304-2-4", cuenta: "", orden: "" },
  "Upgrade motores Waukesha PS's":                        { elementoPEP: "TGPY/OPE-2602-2-3", cuenta: "", orden: "" },
  "Cerco perimetrico KP12 XV-10000 / XV-50001":           { elementoPEP: "TGPY/OPE-2603-1-3", cuenta: "", orden: "" },
  "Plan multianual reemplazo valvulas NG-NGL":            { elementoPEP: "TGPY/OPE-2305-2-4", cuenta: "", orden: "" },
  "Instalacion Sistema Monitoreo de fuego en PS's":       { elementoPEP: "TGPY/OPE-2403-1-4", cuenta: "", orden: "" },
  "Adecuacion Sistema contra incendios BOK":              { elementoPEP: "TGPY/OPE-2406-2-3", cuenta: "", orden: "" },
  "Instalacion motogenerador GN Camp PS3":                { elementoPEP: "TGPY/OPE-2408-2-3", cuenta: "", orden: "" },
  "Reemplazo de Pisos Campamentos Geotecnia":             { elementoPEP: "TGPY/OPE-2409-2-3", cuenta: "", orden: "" },
  "Medicion calidad gas puntos de entrega":               { elementoPEP: "TGPY/OPE-2414-2-3", cuenta: "", orden: "" },
  "Actualizacion Computador Flujo CG Lurin":              { elementoPEP: "TGPY/OPE-2503-2-3", cuenta: "", orden: "" },
  "Mejoras Sala de Servidores Torre Panama":              { elementoPEP: "TGPY/OPE-2509-2-3", cuenta: "", orden: "" },
  "Cerco perimetrico KP75 XV-10002 / XV-50003":           { elementoPEP: "TGPY/OPE-2604-1-3", cuenta: "", orden: "" },
  "Nuevo cerco valvulas XV-50014 / XV-50018":             { elementoPEP: "TGPY/OPE-2608-1-2", cuenta: "", orden: "" },
  "Supervision Instalacion KP43 (Selva)":                 { elementoPEP: "TGPY/OPE-2609-2-2", cuenta: "", orden: "" },
  // Proyectos con Cuenta/Orden
  "Reemplazo de PTARD en PS3":                            { elementoPEP: "", cuenta: "A111078",  orden: "TGP6-2502" },
  "Servicio de Supervision HSE - Selva":                  { elementoPEP: "", cuenta: "6325000",  orden: "TG3CDV1" },
  "Mantenimiento Mayor Puente Comercial KP151+850":       { elementoPEP: "", cuenta: "6323004",  orden: "TGCI-2503" }
};

function getCuentaOrden(frente, subcategoria) {
  if (subcategoria === "CAPEX") {
    const capex = CAPEX_MAPPING[frente];
    if (capex) {
      // Para PEP, usar elementoPEP como "orden" para que aparezca en el PDS
      if (capex.elementoPEP && !capex.cuenta) {
        return { cuenta: capex.elementoPEP, orden: "PEP" };
      }
      return { cuenta: capex.cuenta, orden: capex.orden };
    }
    return { cuenta: "", orden: "" };
  }
  return FRENTES_MAPPING[frente] || { cuenta: "", orden: "" };
}

// ─── CACHÉ DE CARPETAS (evita búsquedas repetidas a Drive) ──────────────────
const folderCache = {};

async function getOrCreateFolder(drive, name, parentId) {
  const cacheKey = `${parentId}/${name}`;
  if (folderCache[cacheKey]) return folderCache[cacheKey];

  const q = `name='${name}' and mimeType='application/vnd.google-apps.folder' and '${parentId}' in parents and trashed=false`;
  const res = await drive.files.list({ q, fields: "files(id)", pageSize: 1, supportsAllDrives: true, includeItemsFromAllDrives: true });

  let folderId;
  if (res.data.files.length > 0) {
    folderId = res.data.files[0].id;
  } else {
    const created = await drive.files.create({
      requestBody: { name, mimeType: "application/vnd.google-apps.folder", parents: [parentId] },
      fields: "id", supportsAllDrives: true
    });
    folderId = created.data.id;
  }
  folderCache[cacheKey] = folderId;
  return folderId;
}

// ─── REINTENTOS EXPONENCIALES ────────────────────────────────────────────────
async function withRetry(fn, maxRetries = 3) {
  for (let i = 0; i < maxRetries; i++) {
    try {
      return await fn();
    } catch (err) {
      if (i === maxRetries - 1) throw err;
      const delay = Math.pow(2, i) * 1000 + Math.random() * 500;
      console.warn(`[PDS] Reintento ${i + 1}/${maxRetries} en ${Math.round(delay)}ms: ${err.message}`);
      await new Promise(r => setTimeout(r, delay));
    }
  }
}

// ─── OBTENER REPORTES DIARIOS DESDE RAW_DATA ─────────────────────────────────
async function getReportesDiarios(sheets, spreadsheetId, sector, yearMonth, filtroSubcategoria) {
  const result = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "RAW_DATA!A2:W"
  });
  const filas = result.data.values || [];

  // Filtrar: sector coincide, fecha empieza con yearMonth, tipo = Reporte Diario
  return filas.filter(r => {
    const fecha = r[1] || "";
    const sectorRow = r[4] || "";
    const tipo = r[15] || "";
    const subcat = r[5] || "";
    const matchBase = sectorRow === sector
      && fecha.startsWith(yearMonth)
      && tipo.toLowerCase().includes("diario");
    if (!matchBase) return false;
    // Filtrar por subcategoria si se especifica
    if (filtroSubcategoria) return subcat === filtroSubcategoria;
    return true;
  }).map(r => ({
    timestamp:    r[0]  || "",
    fecha:        r[1]  || "",
    responsable:  r[2]  || "",
    puesto:       r[3]  || "",
    sector:       r[4]  || "",
    subcategoria: r[5]  || "",
    frente:       r[6]  || "",
    camioneta:    r[7]  || "No",
    placa:        r[8]  || "",
    kmInicial:    r[9]  || "",
    kmFinal:      r[10] || "",
    origen:       r[11] || "",
    destino:      r[12] || "",
    alimentacion: r[13] || "No",
    hospedaje:    r[14] || "No",
    tipoReporte:  r[15] || "",
    descripcion:  r[16] || "",
    avance:       r[17] || "",
    observaciones:r[18] || ""
  }));
}

// ─── AGRUPAR POR DÍA ────────────────────────────────────────────────────────
function agruparPorDia(reportes) {
  const grupos = {};
  for (const r of reportes) {
    const dia = r.fecha.substring(8, 10); // "03" de "2026-04-03"
    if (!grupos[dia]) grupos[dia] = [];
    grupos[dia].push(r);
  }
  return grupos;
}

// ─── CONSTRUIR HOJA DE UN DÍA ────────────────────────────────────────────────
function buildDaySheet(wb, dayStr, sector, yearMonth, reportes, logoImageId, firmaImageId) {
  const ws = wb.addWorksheet(dayStr, {
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
  });

  // ─ Anchos de columna
  ws.getColumn("A").width = 3;
  ws.getColumn("B").width = 5;
  ws.getColumn("C").width = 34;
  ws.getColumn("D").width = 12;
  ws.getColumn("E").width = 12;
  ws.getColumn("F").width = 13;
  ws.getColumn("G").width = 16;
  ws.getColumn("H").width = 16;
  ws.getColumn("I").width = 23;
  ws.getColumn("J").width = 35;
  ws.getColumn("K").width = 14;
  ws.getColumn("L").width = 21;
  ws.getColumn("M").width = 5;
  ws.getColumn("N").width = 12;
  ws.getColumn("O").width = 16;
  ws.getColumn("P").width = 14;

  const fecha = `${yearMonth}-${dayStr}`;
  let row = 1;

  // ─ Row 1: espacio para logo
  ws.getRow(row).height = 15;
  row++;

  // ─ Row 2: TÍTULO
  ws.getRow(row).height = 70;
  ws.mergeCells(`B${row}:L${row}`);
  const titleCell = ws.getCell(`B${row}`);
  titleCell.value = "PARTE DIARIO DE SERVICIOS";
  titleCell.font = { name: "Calibri", size: 24, bold: true };
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  titleCell.border = ALL_BORDERS;

  // Logo TGP (esquina superior izquierda del título)
  if (logoImageId !== undefined) {
    ws.addImage(logoImageId, {
      tl: { col: 1.2, row: 1.1 },
      ext: { width: 90, height: 70 }
    });
  }
  row++;

  // ─ Row 3: DATOS GENERALES
  ws.getRow(row).height = 20;
  ws.mergeCells(`B${row}:L${row}`);
  const datosCell = ws.getCell(`B${row}`);
  datosCell.value = "DATOS GENERALES";
  datosCell.font = { name: "Calibri", size: 12, bold: true, color: { argb: WHITE } };
  datosCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLUE_HEADER } };
  datosCell.alignment = { horizontal: "center", vertical: "middle" };
  datosCell.border = ALL_BORDERS;
  row++;

  // ─ Row 4: Empresa / Zona / Fecha labels
  ws.getRow(row).height = 20;
  ws.mergeCells(`B${row}:I${row}`);
  ws.getCell(`B${row}`).value = "Empresa colaboradora:";
  ws.getCell(`B${row}`).font = { name: "Calibri", size: 12 };
  ws.getCell(`B${row}`).alignment = { vertical: "middle" };
  ws.getCell(`B${row}`).border = ALL_BORDERS;
  ws.getCell(`J${row}`).value = "Zona:";
  ws.getCell(`J${row}`).font = { name: "Calibri", size: 12 };
  ws.getCell(`J${row}`).alignment = { vertical: "middle" };
  ws.getCell(`J${row}`).border = ALL_BORDERS;
  ws.mergeCells(`K${row}:L${row}`);
  ws.getCell(`K${row}`).value = "Fecha:";
  ws.getCell(`K${row}`).font = { name: "Calibri", size: 12 };
  ws.getCell(`K${row}`).alignment = { vertical: "middle" };
  ws.getCell(`K${row}`).border = ALL_BORDERS;
  row++;

  // ─ Row 5: Values
  ws.getRow(row).height = 20;
  ws.mergeCells(`B${row}:I${row}`);
  ws.getCell(`B${row}`).value = "BUREAU VERITAS DEL PERÚ";
  ws.getCell(`B${row}`).font = { name: "Calibri", size: 14, bold: true };
  ws.getCell(`B${row}`).alignment = { horizontal: "center", vertical: "middle" };
  ws.getCell(`B${row}`).border = ALL_BORDERS;
  ws.getCell(`J${row}`).value = sector.toUpperCase();
  ws.getCell(`J${row}`).font = { name: "Calibri", size: 12, bold: true };
  ws.getCell(`J${row}`).alignment = { horizontal: "center", vertical: "middle" };
  ws.getCell(`J${row}`).border = ALL_BORDERS;
  ws.mergeCells(`K${row}:L${row}`);
  ws.getCell(`K${row}`).value = fecha;
  ws.getCell(`K${row}`).font = { name: "Calibri", size: 12 };
  ws.getCell(`K${row}`).alignment = { horizontal: "center", vertical: "middle" };
  ws.getCell(`K${row}`).border = ALL_BORDERS;
  row++;

  // ─ PERSONAL Header
  ws.getRow(row).height = 20;
  ws.mergeCells(`B${row}:L${row}`);
  const persHeader = ws.getCell(`B${row}`);
  persHeader.value = "PERSONAL";
  persHeader.font = { name: "Calibri", size: 12, bold: true, color: { argb: WHITE } };
  persHeader.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLUE_HEADER } };
  persHeader.alignment = { horizontal: "center", vertical: "middle" };
  persHeader.border = ALL_BORDERS;
  row++;

  // ─ Personal Column Headers
  const persHeaderRow = row;
  ws.getRow(row).height = 20;
  const personalHeaders = [
    { col: "B", val: "N", width: 1 },
    { col: "C", val: "NOMBRE", merge: `C${row}:E${row}` },
    { col: "F", val: "ORIGEN" },
    { col: "G", val: "DESTINO" },
    { col: "H", val: "SERVICIO", merge: `H${row}:I${row}` },
    { col: "J", val: "ACTIVIDAD" },
    { col: "K", val: "CUENTA" },
    { col: "L", val: "ORDEN" }
  ];
  for (const h of personalHeaders) {
    if (h.merge) ws.mergeCells(h.merge);
    const cell = ws.getCell(`${h.col}${row}`);
    cell.value = h.val;
    cell.font = { name: "Calibri", size: 12, bold: true, color: { argb: WHITE } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLUE_HEADER } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = ALL_BORDERS;
  }
  row++;

  // ─ Personal Data Rows
  const personalStartRow = row;
  let alimentacionTotal = 0;
  let hospedajeTotal = 0;

  for (let i = 0; i < reportes.length; i++) {
    const r = reportes[i];
    const { cuenta, orden } = getCuentaOrden(r.frente, r.subcategoria);
    ws.getRow(row).height = 20;
    ws.mergeCells(`C${row}:E${row}`);
    ws.mergeCells(`H${row}:I${row}`);

    ws.getCell(`B${row}`).value = i + 1;
    ws.getCell(`B${row}`).alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`B${row}`).border = ALL_BORDERS;

    ws.getCell(`C${row}`).value = r.responsable;
    ws.getCell(`C${row}`).font = { name: "Calibri", size: 11 };
    ws.getCell(`C${row}`).border = ALL_BORDERS;

    ws.getCell(`F${row}`).value = r.origen;
    ws.getCell(`F${row}`).font = { name: "Calibri", size: 11 };
    ws.getCell(`F${row}`).alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`F${row}`).border = ALL_BORDERS;

    ws.getCell(`G${row}`).value = r.destino;
    ws.getCell(`G${row}`).font = { name: "Calibri", size: 11 };
    ws.getCell(`G${row}`).alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`G${row}`).border = ALL_BORDERS;

    ws.getCell(`H${row}`).value = r.puesto || (r.subcategoria === "CAPEX" ? "SUPERVISOR CAPEX" : "SUPERVISOR DE GEOTECNIA SENIOR");
    ws.getCell(`H${row}`).font = { name: "Calibri", size: 10 };
    ws.getCell(`H${row}`).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    ws.getCell(`H${row}`).border = ALL_BORDERS;

    ws.getCell(`J${row}`).value = r.frente;
    ws.getCell(`J${row}`).font = { name: "Calibri", size: 10 };
    ws.getCell(`J${row}`).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    ws.getCell(`J${row}`).border = ALL_BORDERS;

    ws.getCell(`K${row}`).value = cuenta;
    ws.getCell(`K${row}`).font = { name: "Calibri", size: 11 };
    ws.getCell(`K${row}`).alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`K${row}`).border = ALL_BORDERS;

    ws.getCell(`L${row}`).value = orden;
    ws.getCell(`L${row}`).font = { name: "Calibri", size: 11 };
    ws.getCell(`L${row}`).alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`L${row}`).border = ALL_BORDERS;

    if (r.alimentacion === "Si") alimentacionTotal++;
    if (r.hospedaje === "Si") hospedajeTotal++;
    row++;
  }

  // Si no hay reportes, dejar una fila vacía
  if (reportes.length === 0) {
    ws.getRow(row).height = 20;
    ws.mergeCells(`C${row}:E${row}`);
    ws.mergeCells(`H${row}:I${row}`);
    ws.getCell(`B${row}`).value = 1;
    ws.getCell(`B${row}`).border = ALL_BORDERS;
    for (const col of ["C", "F", "G", "H", "J", "K", "L"]) {
      ws.getCell(`${col}${row}`).border = ALL_BORDERS;
    }
    row++;
  }

  // ─ EQUIPOS Header
  ws.getRow(row).height = 20;
  ws.mergeCells(`B${row}:L${row}`);
  const eqHeader = ws.getCell(`B${row}`);
  eqHeader.value = "EQUIPOS";
  eqHeader.font = { name: "Calibri", size: 12, bold: true, color: { argb: WHITE } };
  eqHeader.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLUE_HEADER } };
  eqHeader.alignment = { horizontal: "center", vertical: "middle" };
  eqHeader.border = ALL_BORDERS;
  row++;

  // ─ Equipment Column Headers
  ws.getRow(row).height = 20;
  const eqHeaders = [
    { col: "B", val: "N" },
    { col: "C", val: "EQUIPO" },
    { col: "D", val: "KM INICIO" },
    { col: "E", val: "KM FIN" },
    { col: "F", val: "ORIGEN" },
    { col: "G", val: "DESTINO" },
    { col: "H", val: "SERVICIO", merge: `H${row}:I${row}` },
    { col: "J", val: "ACTIVIDAD" },
    { col: "K", val: "CUENTA" },
    { col: "L", val: "ORDEN" },
    { col: "N", val: "EQUIPO" },
    { col: "O", val: "RECORRIDO" },
    { col: "P", val: "A CERTIFICAR" }
  ];
  for (const h of eqHeaders) {
    if (h.merge) ws.mergeCells(h.merge);
    const cell = ws.getCell(`${h.col}${row}`);
    cell.value = h.val;
    cell.font = { name: "Calibri", size: 12, bold: true, color: { argb: WHITE } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLUE_HEADER } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = ALL_BORDERS;
  }
  row++;

  // ─ Equipment Data Rows
  const vehiculos = reportes.filter(r => r.camioneta === "Si" && r.placa);
  let eqNum = 1;
  for (const r of vehiculos) {
    const { cuenta, orden } = getCuentaOrden(r.frente, r.subcategoria);
    const kmI = parseFloat(r.kmInicial) || 0;
    const kmF = parseFloat(r.kmFinal) || 0;
    const recorrido = kmF - kmI;

    ws.getRow(row).height = 20;
    ws.mergeCells(`H${row}:I${row}`);

    ws.getCell(`B${row}`).value = eqNum++;
    ws.getCell(`B${row}`).alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`B${row}`).border = ALL_BORDERS;

    ws.getCell(`C${row}`).value = `CAMIONETA ${r.placa}`;
    ws.getCell(`C${row}`).font = { name: "Calibri", size: 11 };
    ws.getCell(`C${row}`).border = ALL_BORDERS;

    ws.getCell(`D${row}`).value = kmI || "";
    ws.getCell(`D${row}`).font = { name: "Calibri", size: 11 };
    ws.getCell(`D${row}`).alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`D${row}`).border = ALL_BORDERS;

    ws.getCell(`E${row}`).value = kmF || "";
    ws.getCell(`E${row}`).font = { name: "Calibri", size: 11 };
    ws.getCell(`E${row}`).alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`E${row}`).border = ALL_BORDERS;

    ws.getCell(`F${row}`).value = r.origen;
    ws.getCell(`F${row}`).font = { name: "Calibri", size: 11 };
    ws.getCell(`F${row}`).alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`F${row}`).border = ALL_BORDERS;

    ws.getCell(`G${row}`).value = r.destino;
    ws.getCell(`G${row}`).font = { name: "Calibri", size: 11 };
    ws.getCell(`G${row}`).alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`G${row}`).border = ALL_BORDERS;

    ws.getCell(`H${row}`).value = "ALQUILER DE CAMIONETA SIN CONDUCTOR";
    ws.getCell(`H${row}`).font = { name: "Calibri", size: 9 };
    ws.getCell(`H${row}`).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    ws.getCell(`H${row}`).border = ALL_BORDERS;

    ws.getCell(`J${row}`).value = r.frente;
    ws.getCell(`J${row}`).font = { name: "Calibri", size: 10 };
    ws.getCell(`J${row}`).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    ws.getCell(`J${row}`).border = ALL_BORDERS;

    ws.getCell(`K${row}`).value = cuenta;
    ws.getCell(`K${row}`).font = { name: "Calibri", size: 11 };
    ws.getCell(`K${row}`).alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`K${row}`).border = ALL_BORDERS;

    ws.getCell(`L${row}`).value = orden;
    ws.getCell(`L${row}`).font = { name: "Calibri", size: 11 };
    ws.getCell(`L${row}`).alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`L${row}`).border = ALL_BORDERS;

    // Columnas extra de resumen equipo
    ws.getCell(`N${row}`).value = r.placa;
    ws.getCell(`N${row}`).font = { name: "Calibri", size: 11 };
    ws.getCell(`N${row}`).alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`N${row}`).border = ALL_BORDERS;

    ws.getCell(`O${row}`).value = recorrido > 0 ? recorrido : 0;
    ws.getCell(`O${row}`).font = { name: "Calibri", size: 11 };
    ws.getCell(`O${row}`).alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`O${row}`).border = ALL_BORDERS;

    ws.getCell(`P${row}`).value = recorrido > 0 ? recorrido : 0;
    ws.getCell(`P${row}`).font = { name: "Calibri", size: 11 };
    ws.getCell(`P${row}`).alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`P${row}`).border = ALL_BORDERS;

    row++;
  }

  // Fila vacía si no hay vehículos
  if (vehiculos.length === 0) {
    ws.getRow(row).height = 20;
    ws.mergeCells(`H${row}:I${row}`);
    ws.getCell(`B${row}`).value = 1;
    ws.getCell(`B${row}`).border = ALL_BORDERS;
    for (const col of ["C", "D", "E", "F", "G", "H", "J", "K", "L"]) {
      ws.getCell(`${col}${row}`).border = ALL_BORDERS;
    }
    row++;
  }

  // ─ COMENTARIOS Header
  ws.getRow(row).height = 20;
  ws.mergeCells(`B${row}:L${row}`);
  const comHeader = ws.getCell(`B${row}`);
  comHeader.value = "COMENTARIOS";
  comHeader.font = { name: "Calibri", size: 12, bold: true, color: { argb: WHITE } };
  comHeader.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLUE_HEADER } };
  comHeader.alignment = { horizontal: "center", vertical: "middle" };
  comHeader.border = ALL_BORDERS;
  row++;

  // ─ Comentarios Text
  // Construir texto de comentarios con resumen de personal + actividades
  let comentarios = "";
  const supervisorCount = reportes.length;
  const camionetaCount = vehiculos.length;

  if (supervisorCount > 0 || camionetaCount > 0) {
    const parts = [];
    if (supervisorCount > 0) parts.push(`${String(supervisorCount).padStart(2, "0")} Supervisor${supervisorCount > 1 ? "es" : ""}`);
    if (camionetaCount > 0) parts.push(`${String(camionetaCount).padStart(2, "0")} camioneta${camionetaCount > 1 ? "s" : ""}`);
    comentarios += parts.join(", ") + "\n\n";

    // Detalle de camionetas
    for (const r of vehiculos) {
      comentarios += `- Camioneta ${r.placa}, a Servicio de ${r.responsable}\n`;
    }
    if (vehiculos.length > 0) comentarios += "\n";

    comentarios += "Actividades:\n\n";
    for (const r of reportes) {
      comentarios += `${r.responsable}\n`;
      if (r.descripcion) comentarios += `* ${r.descripcion}\n`;
      if (r.observaciones) comentarios += `* Obs: ${r.observaciones}\n`;
      comentarios += "\n";
    }

    // Gastos logísticos
    comentarios += "GASTOS LOGISTICOS:\n";
    comentarios += `- ALIMENTACION: ${String(alimentacionTotal).padStart(2, "0")}\n`;
    comentarios += `- HOSPEDAJE: ${String(hospedajeTotal).padStart(2, "0")}`;
  }

  ws.mergeCells(`B${row}:L${row + 1}`);
  const comCell = ws.getCell(`B${row}`);
  comCell.value = comentarios;
  comCell.font = { name: "Calibri", size: 10 };
  comCell.alignment = { vertical: "top", wrapText: true };
  comCell.border = ALL_BORDERS;
  ws.getRow(row).height = 44;
  ws.getRow(row + 1).height = 200;
  row += 2;

  // ─ REPRESENTANTES
  ws.getRow(row).height = 20;
  ws.mergeCells(`B${row}:H${row}`);
  ws.getCell(`B${row}`).value = "REPRESENTANTE BUREAU VERITAS";
  ws.getCell(`B${row}`).font = { name: "Calibri", size: 12, bold: true, color: { argb: WHITE } };
  ws.getCell(`B${row}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLUE_HEADER } };
  ws.getCell(`B${row}`).alignment = { horizontal: "center", vertical: "middle" };
  ws.getCell(`B${row}`).border = ALL_BORDERS;
  ws.mergeCells(`I${row}:L${row}`);
  ws.getCell(`I${row}`).value = "REPRESENTANTE DE TGP";
  ws.getCell(`I${row}`).font = { name: "Calibri", size: 12, bold: true, color: { argb: WHITE } };
  ws.getCell(`I${row}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLUE_HEADER } };
  ws.getCell(`I${row}`).alignment = { horizontal: "center", vertical: "middle" };
  ws.getCell(`I${row}`).border = ALL_BORDERS;
  row++;

  // ─ Firma
  ws.getRow(row).height = 50;
  ws.mergeCells(`B${row}:H${row}`);
  ws.getCell(`B${row}`).value = "Firma:";
  ws.getCell(`B${row}`).font = { name: "Calibri", size: 11 };
  ws.getCell(`B${row}`).alignment = { vertical: "top" };
  ws.getCell(`B${row}`).border = ALL_BORDERS;
  ws.mergeCells(`I${row}:L${row}`);
  ws.getCell(`I${row}`).value = "Firma:";
  ws.getCell(`I${row}`).font = { name: "Calibri", size: 11 };
  ws.getCell(`I${row}`).alignment = { vertical: "top" };
  ws.getCell(`I${row}`).border = ALL_BORDERS;

  // Firma image
  if (firmaImageId !== undefined) {
    ws.addImage(firmaImageId, {
      tl: { col: 2.5, row: row - 0.8 },
      ext: { width: 120, height: 50 }
    });
  }
  row++;

  // ─ Blank row
  ws.getRow(row).height = 20;
  ws.mergeCells(`B${row}:H${row}`);
  ws.getCell(`B${row}`).border = ALL_BORDERS;
  ws.mergeCells(`I${row}:L${row}`);
  ws.getCell(`I${row}`).border = ALL_BORDERS;
  row++;

  // ─ Nombre
  ws.mergeCells(`B${row}:H${row + 1}`);
  ws.getCell(`B${row}`).value = "Nombre: Yuri Arangoitia. R";
  ws.getCell(`B${row}`).font = { name: "Calibri", size: 11, bold: true };
  ws.getCell(`B${row}`).alignment = { horizontal: "center", vertical: "middle" };
  ws.getCell(`B${row}`).border = ALL_BORDERS;
  ws.mergeCells(`I${row}:L${row + 1}`);
  ws.getCell(`I${row}`).value = "Nombre:";
  ws.getCell(`I${row}`).font = { name: "Calibri", size: 11 };
  ws.getCell(`I${row}`).alignment = { horizontal: "center", vertical: "middle" };
  ws.getCell(`I${row}`).border = ALL_BORDERS;

  return ws;
}

// ─── CREAR WORKBOOK Y SUBIR A DRIVE ─────────────────────────────────────────
async function buildAndUploadPDS(drive, sector, yearMonth, reportes, folderParentId, fileNamePrefix, description, pdsRootFolderId) {
  const porDia = agruparPorDia(reportes);
  console.log(`[PDS] ${fileNamePrefix} — Días con datos: ${Object.keys(porDia).join(", ")}`);

  const wb = new ExcelJS.Workbook();
  wb.creator = "TGP Reportes v10";
  wb.created = new Date();

  const logoPath = path.join(__dirname, "templates", "logo_tgp.png");
  const firmaPath = path.join(__dirname, "templates", "firma_sello.jpeg");

  let logoImageId, firmaImageId;
  try {
    if (fs.existsSync(logoPath)) {
      logoImageId = wb.addImage({ filename: logoPath, extension: "png" });
    }
  } catch (e) {
    console.warn("[PDS] Logo no encontrado, omitiendo.");
  }
  try {
    if (fs.existsSync(firmaPath)) {
      firmaImageId = wb.addImage({ filename: firmaPath, extension: "jpeg" });
    }
  } catch (e) {
    console.warn("[PDS] Firma no encontrada, omitiendo.");
  }

  const dias = Object.keys(porDia).sort();
  for (const dia of dias) {
    buildDaySheet(wb, dia, sector, yearMonth, porDia[dia], logoImageId, firmaImageId);
  }

  const buffer = await wb.xlsx.writeBuffer();
  console.log(`[PDS] Excel generado: ${Math.round(buffer.length / 1024)}KB, ${dias.length} pestañas`);

  const fileName = `${fileNamePrefix}.xlsx`;
  const existingQ = `name='${fileName}' and '${folderParentId}' in parents and trashed=false`;
  const existing = await withRetry(() =>
    drive.files.list({ q: existingQ, fields: "files(id)", pageSize: 1, supportsAllDrives: true })
  );

  let fileId;
  if (existing.data.files.length > 0) {
    fileId = existing.data.files[0].id;
    await withRetry(() =>
      drive.files.update({
        fileId,
        media: {
          mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          body: Readable.from(buffer)
        },
        supportsAllDrives: true
      })
    );
    console.log(`[PDS] Archivo actualizado: ${fileName} (${fileId})`);
  } else {
    const created = await withRetry(() =>
      drive.files.create({
        requestBody: {
          name: fileName,
          parents: [folderParentId],
          description
        },
        media: {
          mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          body: Readable.from(buffer)
        },
        fields: "id,webViewLink",
        supportsAllDrives: true
      })
    );
    fileId = created.data.id;
    console.log(`[PDS] Archivo creado: ${fileName} (${fileId})`);
  }

  return fileId;
}

// ─── GENERAR PDS COMPLETO ────────────────────────────────────────────────────
async function generatePDS(sheets, drive, spreadsheetId, sector, yearMonth, pdsRootFolderId) {
  console.log(`[PDS] Generando PDS para ${sector} / ${yearMonth}...`);

  // Carpeta raíz PDS
  const raizId = await withRetry(() =>
    getOrCreateFolder(drive, "PARTE DIARIO DE SERVICIO", pdsRootFolderId || "root")
  );
  const sectorId = await withRetry(() =>
    getOrCreateFolder(drive, sector.toUpperCase(), raizId)
  );

  // ═══ PDS GEOTECNIA (comportamiento original) ═══
  const reportesGeo = await withRetry(() =>
    getReportesDiarios(sheets, spreadsheetId, sector, yearMonth, "Geotecnia")
  );

  if (reportesGeo.length > 0) {
    await buildAndUploadPDS(
      drive, sector, yearMonth, reportesGeo, sectorId,
      `PDS_${sector.toUpperCase()}_${yearMonth}`,
      `Parte Diario de Servicios - Geotecnia - ${sector} - ${yearMonth}`,
      pdsRootFolderId
    );
  } else {
    console.log("[PDS] Sin reportes Geotecnia para este sector/mes.");
  }

  // ═══ PDS CAPEX (un PDS por proyecto independiente) ═══
  const reportesCapex = await withRetry(() =>
    getReportesDiarios(sheets, spreadsheetId, sector, yearMonth, "CAPEX")
  );

  if (reportesCapex.length > 0) {
    // Crear carpeta CAPEX dentro del sector
    const capexFolderId = await withRetry(() =>
      getOrCreateFolder(drive, "CAPEX", sectorId)
    );

    // Agrupar por proyecto (frente)
    const porProyecto = {};
    for (const r of reportesCapex) {
      if (!porProyecto[r.frente]) porProyecto[r.frente] = [];
      porProyecto[r.frente].push(r);
    }

    // Generar un PDS por cada proyecto CAPEX
    for (const [proyecto, reportes] of Object.entries(porProyecto)) {
      // Nombre de archivo limpio (sin caracteres especiales)
      const nombreLimpio = proyecto
        .replace(/[\/\\:*?"<>|]/g, "-")
        .replace(/\s+/g, "_")
        .substring(0, 60);
      const filePrefix = `PDS_CAPEX_${nombreLimpio}_${yearMonth}`;

      console.log(`[PDS-CAPEX] Proyecto: ${proyecto} — ${reportes.length} reportes`);

      await buildAndUploadPDS(
        drive, sector, yearMonth, reportes, capexFolderId,
        filePrefix,
        `PDS CAPEX - ${proyecto} - ${sector} - ${yearMonth}`,
        pdsRootFolderId
      );
    }

    console.log(`[PDS-CAPEX] ${Object.keys(porProyecto).length} proyectos procesados para ${sector}`);
  } else {
    console.log("[PDS] Sin reportes CAPEX para este sector/mes.");
  }

  return true;
}

module.exports = {
  generatePDS,
  FRENTES_MAPPING,
  CAPEX_MAPPING,
  getCuentaOrden
};
