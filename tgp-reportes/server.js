// ═══════════════════════════════════════════════════════════════════════════
// SISTEMA DE REPORTES DE CAMPO — server.js v10.1 (+ HSE + PDS + Portal Gestion)
// ═══════════════════════════════════════════════════════════════════════════

require("dotenv").config();
const express    = require("express");
const cors       = require("cors");
const path       = require("path");
const { google } = require("googleapis");
const nodemailer = require("nodemailer");
const { generatePDS } = require("./pds-generator");

const app  = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: "60mb" }));
app.use(express.urlencoded({ extended: true, limit: "60mb" }));
app.use(express.static(path.join(__dirname, "public")));

// ─── CONFIG — CAMPO ──────────────────────────────────────────────────────────
const EMAIL_COORDINADOR = process.env.EMAIL_COORDINADOR ||
  "yuri.arangoitia@bureauveritas.com, daniel.cedano@bureauveritas.com, gustavo.fernandez@bureauveritas.com, fiorella.diaz@bureauveritas.com";
const CARPETA_RAIZ      = process.env.CARPETA_RAIZ || "Reportes de Campo 2026";
const CARPETA_RAIZ_ID   = process.env.CARPETA_RAIZ_ID || "";
const CLAVE_DASHBOARD   = process.env.CLAVE_DASHBOARD || "campo2026";
const SPREADSHEET_ID    = process.env.SPREADSHEET_ID || "";

// ─── CONFIG — HSE ────────────────────────────────────────────────────────────
const CARPETA_RAIZ_HSE    = process.env.CARPETA_RAIZ_HSE || "Reportes HSE 2026";
const CARPETA_RAIZ_HSE_ID = process.env.CARPETA_RAIZ_HSE_ID || "";
const SPREADSHEET_ID_HSE  = process.env.SPREADSHEET_ID_HSE || "";

// ─── DATOS ESTÁTICOS ─────────────────────────────────────────────────────────
const FRENTES = {
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

// ─── CAPEX — Proyectos con Elemento PEP ─────────────────────────────────────
const PROYECTOS_CAPEX_PEP = [
  { proyecto: "Protecciones mecanicas ductos NG/NG",                  elementoPEP: "TGPY/OPE-1501-2-2" },
  { proyecto: "Mejoras Skids Gas combustible en PS's",                elementoPEP: "TGPY/OPE-1901-2-4" },
  { proyecto: "Construccion nuevas instalaciones Lurin",              elementoPEP: "TGPY/OPE-1902-2-2" },
  { proyecto: "Mejora instalaciones Aerodromo Kiteni",                elementoPEP: "TGPY/OPE-2101-1-2" },
  { proyecto: "Actualizacion Sistema de Automatizacion",              elementoPEP: "TGPY/OPE-2201-2-4" },
  { proyecto: "Cambio Tableros / Luminarias Areas Clasificadas",      elementoPEP: "TGPY/OPE-2301-2-3" },
  { proyecto: "Plan Mitigacion Ruido PC Kamani (venteo)",             elementoPEP: "TGPY/OPE-2302-2-4" },
  { proyecto: "Adecuacion Valvula Sobrepresion NG32 PS1",             elementoPEP: "TGPY/OPE-2304-2-4" },
  { proyecto: "Upgrade motores Waukesha PS's",                        elementoPEP: "TGPY/OPE-2602-2-3" },
  { proyecto: "Cerco perimetrico KP12 XV-10000 / XV-50001",           elementoPEP: "TGPY/OPE-2603-1-3" },
  { proyecto: "Plan multianual reemplazo valvulas NG-NGL",            elementoPEP: "TGPY/OPE-2305-2-4" },
  { proyecto: "Instalacion Sistema Monitoreo de fuego en PS's",       elementoPEP: "TGPY/OPE-2403-1-4" },
  { proyecto: "Adecuacion Sistema contra incendios BOK",              elementoPEP: "TGPY/OPE-2406-2-3" },
  { proyecto: "Instalacion motogenerador GN Camp PS3",                elementoPEP: "TGPY/OPE-2408-2-3" },
  { proyecto: "Reemplazo de Pisos Campamentos Geotecnia",             elementoPEP: "TGPY/OPE-2409-2-3" },
  { proyecto: "Medicion calidad gas puntos de entrega",               elementoPEP: "TGPY/OPE-2414-2-3" },
  { proyecto: "Actualizacion Computador Flujo CG Lurin",              elementoPEP: "TGPY/OPE-2503-2-3" },
  { proyecto: "Mejoras Sala de Servidores Torre Panama",              elementoPEP: "TGPY/OPE-2509-2-3" },
  { proyecto: "Cerco perimetrico KP75 XV-10002 / XV-50003",           elementoPEP: "TGPY/OPE-2604-1-3" },
  { proyecto: "Nuevo cerco valvulas XV-50014 / XV-50018",             elementoPEP: "TGPY/OPE-2608-1-2" },
  { proyecto: "Supervision Instalacion KP43 (Selva)",                 elementoPEP: "TGPY/OPE-2609-2-2" }
];

// ─── CAPEX — Proyectos con Cuenta/Orden ─────────────────────────────────────
const PROYECTOS_CAPEX_CO = [
  { proyecto: "Reemplazo de PTARD en PS3",                    cuenta: "A111078",  orden: "TGP6-2502" },
  { proyecto: "Servicio de Supervision HSE - Selva",          cuenta: "6325000",  orden: "TG3CDV1" },
  { proyecto: "Mantenimiento Mayor Puente Comercial KP151+850", cuenta: "6323004", orden: "TGCI-2503" }
];

const SUPERVISORES = [
  // ── Geotecnia ──
  { nombre: "CRISTHIAN BAQUERIZO",       sector: "", subcategoria: "Geotecnia" },
  { nombre: "WALTER JESUS",              sector: "", subcategoria: "Geotecnia" },
  { nombre: "LIZ GUERRERO",              sector: "", subcategoria: "Geotecnia" },
  { nombre: "CARLOS DE LA CRUZ",         sector: "", subcategoria: "Geotecnia" },
  { nombre: "ROY HERRADA",               sector: "", subcategoria: "Geotecnia" },
  { nombre: "ABEL SANCHEZ QUIHUI",       sector: "", subcategoria: "Geotecnia" },
  { nombre: "SAMUEL JARA MAYTA",         sector: "", subcategoria: "Geotecnia" },
  { nombre: "NIKOLAI ARANGOITIA",        sector: "", subcategoria: "Geotecnia" },
  { nombre: "ROGELIO CHAMPI CHOQUEPATA", sector: "", subcategoria: "Geotecnia" },
  { nombre: "JHON FUENTES",              sector: "", subcategoria: "Geotecnia" },
  { nombre: "RUBEN NUÑEZ",               sector: "", subcategoria: "Geotecnia" },
  { nombre: "JORDAN GALLO",              sector: "", subcategoria: "Geotecnia" },
  { nombre: "ABRAHAM JIMENEZ",           sector: "", subcategoria: "Geotecnia" },
  { nombre: "NEISSER MAMANI",            sector: "", subcategoria: "Geotecnia" },
  { nombre: "PAUL PACSI ALAVE",          sector: "", subcategoria: "Geotecnia" },
  { nombre: "DANIEL ATAYUPANQUI TARCO",  sector: "", subcategoria: "Geotecnia" },
  { nombre: "CARLOS PUENTE",             sector: "", subcategoria: "Geotecnia" },
  { nombre: "CHRISTIAN RAMIREZ",         sector: "", subcategoria: "Geotecnia" },
  { nombre: "ORLANDO SALHUANA",          sector: "", subcategoria: "Geotecnia" },
  // ── CAPEX ──
  { nombre: "FERNANDO DAVILA",           sector: "", subcategoria: "CAPEX" },
  { nombre: "GREGORY VELASQUEZ",         sector: "", subcategoria: "CAPEX" },
  { nombre: "JAIME GALVAN",              sector: "", subcategoria: "CAPEX" },
  { nombre: "SANTIAGO ROJAS",            sector: "", subcategoria: "CAPEX" },
  { nombre: "RICARDO QUEZADA",           sector: "", subcategoria: "CAPEX" },
  { nombre: "GUSTAVO CANDIOTTI",         sector: "", subcategoria: "CAPEX" },
  { nombre: "RONALDO MONTERO",           sector: "", subcategoria: "CAPEX" },
  { nombre: "EDWIN HERBOZO",             sector: "", subcategoria: "CAPEX" },
  { nombre: "LUIS JAUREGUI",             sector: "", subcategoria: "CAPEX" },
  { nombre: "ABRAHAM ARCE",              sector: "", subcategoria: "CAPEX" },
  { nombre: "PETTER BLAS",               sector: "", subcategoria: "CAPEX" }
];

// ─── GOOGLE AUTH (OAuth2) ────────────────────────────────────────────────────
function getGoogleAuth() {
  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    "https://developers.google.com/oauthplayground"
  );
  oauth2Client.setCredentials({
    refresh_token: process.env.GOOGLE_REFRESH_TOKEN
  });
  return oauth2Client;
}

// ─── DRIVE: obtener o crear carpeta ─────────────────────────────────────────
async function carpetaEnPadre(drive, nombre, padreId) {
  const q = `name='${nombre}' and mimeType='application/vnd.google-apps.folder' and '${padreId}' in parents and trashed=false`;
  const res = await drive.files.list({
    q,
    fields: "files(id)",
    pageSize: 1,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true
  });
  if (res.data.files.length > 0) return res.data.files[0].id;
  const created = await drive.files.create({
    requestBody: { name: nombre, mimeType: "application/vnd.google-apps.folder", parents: [padreId] },
    fields: "id",
    supportsAllDrives: true
  });
  return created.data.id;
}

async function obtenerRaizId(drive, raizNombre, raizIdEnv) {
  if (raizIdEnv) return raizIdEnv;
  const q = `name='${raizNombre}' and mimeType='application/vnd.google-apps.folder' and 'root' in parents and trashed=false`;
  const res = await drive.files.list({ q, fields: "files(id)", pageSize: 1 });
  if (res.data.files.length > 0) return res.data.files[0].id;
  const r = await drive.files.create({
    requestBody: { name: raizNombre, mimeType: "application/vnd.google-apps.folder" },
    fields: "id"
  });
  return r.data.id;
}

async function obtenerCarpetaDrive(drive, sector, subcategoria, frente, tipo, fecha) {
  const raizId   = await obtenerRaizId(drive, CARPETA_RAIZ, CARPETA_RAIZ_ID);
  const sectorId = await carpetaEnPadre(drive, sector,       raizId);
  const subcatId = await carpetaEnPadre(drive, subcategoria, sectorId);
  const frenteId = await carpetaEnPadre(drive, frente,       subcatId);
  const tipoId   = await carpetaEnPadre(drive, tipo,         frenteId);
  return           await carpetaEnPadre(drive, fecha,        tipoId);
}

async function obtenerCarpetaDriveHSE(drive, sector, subcategoria, frente, tipo, fecha) {
  const raizId   = await obtenerRaizId(drive, CARPETA_RAIZ_HSE, CARPETA_RAIZ_HSE_ID);
  const sectorId = await carpetaEnPadre(drive, sector,       raizId);
  const subcatId = await carpetaEnPadre(drive, subcategoria, sectorId);
  const frenteId = await carpetaEnPadre(drive, frente,       subcatId);
  const tipoId   = await carpetaEnPadre(drive, tipo,         frenteId);
  return           await carpetaEnPadre(drive, fecha,        tipoId);
}

// ─── SHEETS: registrar fila — CAMPO ─────────────────────────────────────────
async function registrarEnSheet(sheets, datos, urlArchivo, ahora, fecha) {
  const SHEET_NAME = "RAW_DATA";
  if (!SPREADSHEET_ID) {
    console.warn("SPREADSHEET_ID no configurado — omitiendo registro en Sheets");
    return;
  }
  let sheetData;
  try {
    sheetData = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A1:W1`
    });
  } catch (e) {
    sheetData = { data: { values: [] } };
  }
  const hasHeader = sheetData.data.values && sheetData.data.values.length > 0;
  if (!hasHeader) {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A1`,
      valueInputOption: "RAW",
      requestBody: {
        values: [[
          "Timestamp","Fecha","Responsable","Puesto","Sector","Subcategoria","Frente de Trabajo",
          "Camioneta","Placa","KM Inicial","KM Final","Origen","Destino",
          "Alimentacion","Hospedaje","Tipo Reporte","Descripcion",
          "% Avance","Observaciones","Requiere Oficina",
          "Nombre Archivo","Link Archivo","Semana"
        ]]
      }
    });
  }
  const sem = semanaDelAno(new Date());
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A1`,
    valueInputOption: "RAW",
    requestBody: {
      values: [[
        ahora, fecha,
        datos.responsable, datos.puesto || "",
        datos.sector, datos.subcategoria, datos.frente || "",
        datos.camioneta || "No", datos.placa || "",
        datos.kmInicial || "", datos.kmFinal || "",
        datos.origen || "", datos.destino || "",
        datos.alimentacion || "No", datos.hospedaje || "No",
        datos.tipoReporte, datos.descripcion,
        datos.avance || "", datos.problemas || "",
        "No", datos.nombreArchivo || "", urlArchivo || "", sem
      ]]
    }
  });
}

// ─── SHEETS: registrar fila — HSE ────────────────────────────────────────────
async function registrarEnSheetHSE(sheets, datos, urlArchivo, ahora, fecha) {
  const SHEET_NAME = "RAW_DATA";
  if (!SPREADSHEET_ID_HSE) {
    console.warn("SPREADSHEET_ID_HSE no configurado — omitiendo registro en Sheets HSE");
    return;
  }
  let sheetData;
  try {
    sheetData = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID_HSE,
      range: `${SHEET_NAME}!A1:K1`
    });
  } catch (e) {
    sheetData = { data: { values: [] } };
  }
  const hasHeader = sheetData.data.values && sheetData.data.values.length > 0;
  if (!hasHeader) {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID_HSE,
      range: `${SHEET_NAME}!A1`,
      valueInputOption: "RAW",
      requestBody: {
        values: [[
          "Timestamp", "Fecha", "Responsable", "Puesto",
          "Sector", "Subcategoria", "Frente",
          "Tipo Reporte", "Nombre Archivo", "Link Archivo", "Semana"
        ]]
      }
    });
  }
  const sem = semanaDelAno(new Date());
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID_HSE,
    range: `${SHEET_NAME}!A1`,
    valueInputOption: "RAW",
    requestBody: {
      values: [[
        ahora,
        fecha,
        datos.responsable,
        datos.puesto       || "",
        datos.sector,
        datos.subcategoria,
        datos.frente       || "",
        datos.tipoReporte,
        datos.nombreArchivo || "",
        urlArchivo          || "",
        sem
      ]]
    }
  });
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────
function semanaDelAno(d) {
  const ini = new Date(d.getFullYear(), 0, 1);
  return Math.ceil(((d - ini) / 86400000 + ini.getDay() + 1) / 7);
}
function fechaLima() {
  return new Date().toLocaleString("sv-SE", { timeZone: "America/Lima" });
}
function fechaCorta() {
  return fechaLima().substring(0, 10);
}

// ─── EMAIL — CAMPO ───────────────────────────────────────────────────────────
async function enviarEmail(datos, urlArchivo, ahora) {
  if (!process.env.SMTP_HOST) { console.warn("SMTP no configurado — omitiendo email"); return; }
  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST, port: parseInt(process.env.SMTP_PORT || "587"),
    secure: process.env.SMTP_SECURE === "true",
    auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS }
  });
  const adjunto = urlArchivo ? `<br><a href="${urlArchivo}" style="color:#2E75B6;font-weight:bold">Ver archivo en Drive</a>` : "";
  const asunto  = `[Campo] ${datos.tipoReporte} | ${datos.sector} | ${datos.frente} | ${datos.responsable}`;
  const fila    = (k, v) => `<div style="padding:6px 0;border-bottom:1px solid #eee;font-size:13px"><b style="color:#555">${k}:</b> ${v || "-"}</div>`;
  const cuerpo  = `<div style="font-family:Arial,sans-serif;max-width:580px">
    <div style="background:#1F3864;color:#fff;padding:16px;border-radius:8px 8px 0 0">
      <b style="font-size:15px">Nuevo Reporte de Campo</b><br>
      <span style="font-size:11px;opacity:.8">${ahora}</span>
    </div>
    <div style="padding:16px;border:1px solid #ddd;border-top:0;border-radius:0 0 8px 8px;background:#fafafa">
      ${fila("Responsable", datos.responsable)}
      ${fila("Puesto", datos.puesto)}
      ${fila("Sector", datos.sector)}
      ${fila("Subcategoria", datos.subcategoria)}
      ${fila("Frente de Trabajo", datos.frente)}
      ${fila("Camioneta", datos.camioneta)}
      ${datos.camioneta === "Si" ? fila("Placa", datos.placa) : ""}
      ${datos.kmInicial ? fila("KM Inicial / Final", datos.kmInicial + " / " + datos.kmFinal) : ""}
      ${datos.origen ? fila("Origen / Destino", datos.origen + " > " + datos.destino) : ""}
      ${fila("Alimentacion", datos.alimentacion)}
      ${fila("Hospedaje", datos.hospedaje)}
      ${fila("Tipo", datos.tipoReporte)}
      ${fila("Descripcion", datos.descripcion)}
      ${datos.avance ? fila("% Avance", parseFloat(datos.avance).toFixed(2) + "%") : ""}
      ${datos.problemas ? fila("Observaciones", datos.problemas) : ""}
      ${adjunto}
    </div>
  </div>`;
  await transporter.sendMail({
    from: `"Reportes TGP" <${process.env.SMTP_USER}>`,
    to: EMAIL_COORDINADOR, subject: asunto, html: cuerpo
  });
}

// ─── EMAIL — HSE ─────────────────────────────────────────────────────────────
async function enviarEmailHSE(datos, urlArchivo, ahora) {
  if (!process.env.SMTP_HOST) { console.warn("SMTP no configurado — omitiendo email HSE"); return; }
  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST, port: parseInt(process.env.SMTP_PORT || "587"),
    secure: process.env.SMTP_SECURE === "true",
    auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS }
  });
  const adjunto = urlArchivo ? `<br><a href="${urlArchivo}" style="color:#2E75B6;font-weight:bold">Ver archivo en Drive</a>` : "";
  const asunto  = `[HSE] ${datos.tipoReporte} | ${datos.sector} | ${datos.frente} | ${datos.responsable}`;
  const fila    = (k, v) => `<div style="padding:6px 0;border-bottom:1px solid #eee;font-size:13px"><b style="color:#555">${k}:</b> ${v || "-"}</div>`;
  const cuerpo  = `<div style="font-family:Arial,sans-serif;max-width:580px">
    <div style="background:#1F3864;color:#fff;padding:16px;border-radius:8px 8px 0 0">
      <b style="font-size:15px">Nuevo Reporte HSE</b><br>
      <span style="font-size:11px;opacity:.8">${ahora}</span>
    </div>
    <div style="padding:16px;border:1px solid #ddd;border-top:0;border-radius:0 0 8px 8px;background:#fafafa">
      ${fila("Responsable", datos.responsable)}
      ${fila("Puesto", datos.puesto)}
      ${fila("Sector", datos.sector)}
      ${fila("Subcategoria", datos.subcategoria)}
      ${fila("Frente de Trabajo", datos.frente)}
      ${fila("Tipo de Reporte", datos.tipoReporte)}
      ${datos.observaciones ? fila("Observaciones", datos.observaciones) : ""}
      ${adjunto}
    </div>
  </div>`;
  await transporter.sendMail({
    from: `"Reportes TGP" <${process.env.SMTP_USER}>`,
    to: EMAIL_COORDINADOR, subject: asunto, html: cuerpo
  });
}

// ─── RUTAS — CAMPO ───────────────────────────────────────────────────────────
app.get("/",          (req, res) => res.sendFile(path.join(__dirname, "public", "formulario.html")));
app.get("/dashboard", (req, res) => res.sendFile(path.join(__dirname, "public", "dashboard.html")));
app.get("/api/supervisores", (req, res) => {
  const { subcategoria, excluir } = req.query;
  if (subcategoria) {
    res.json(SUPERVISORES.filter(s => s.subcategoria === subcategoria));
  } else if (excluir) {
    res.json(SUPERVISORES.filter(s => s.subcategoria !== excluir));
  } else {
    res.json(SUPERVISORES);
  }
});
app.get("/api/todos-frentes",  (req, res) => res.json(FRENTES));
app.get("/api/frentes", (req, res) => {
  const { sector, subcat } = req.query;
  res.json(FRENTES[`${sector}_${subcat}`] || []);
});
app.get("/api/proyectos-capex", (req, res) => {
  res.json({ pep: PROYECTOS_CAPEX_PEP, cuentaOrden: PROYECTOS_CAPEX_CO });
});
app.post("/api/verificar-clave", (req, res) => {
  const { clave } = req.body;
  res.json({ ok: String(clave).trim() === CLAVE_DASHBOARD });
});

// ─── RUTA — PORTAL ───────────────────────────────────────────────────────────
app.get("/portal", (req, res) => res.sendFile(path.join(__dirname, "public", "portal.html")));

// ─── RUTAS — HSE ─────────────────────────────────────────────────────────────
app.get("/hse", (req, res) => res.sendFile(path.join(__dirname, "public", "formulario-hse.html")));

app.post("/api/reporte-hse", async (req, res) => {
  const datos = req.body;
  try {
    const ahora = fechaLima().replace("T", " ");
    const hoy   = datos.fecha || fechaCorta();
    let urlArchivo = "";

    const auth   = getGoogleAuth();
    const drive  = google.drive({ version: "v3", auth });
    const sheets = google.sheets({ version: "v4", auth });

    if (datos.archivoBase64 && datos.archivoBase64.length > 100) {
      const carpetaId = await obtenerCarpetaDriveHSE(
        drive, datos.sector, datos.subcategoria,
        datos.frente, datos.tipoReporte, hoy
      );
      const buffer    = Buffer.from(datos.archivoBase64, "base64");
      const uploadRes = await drive.files.create({
        requestBody: {
          name: datos.nombreArchivo,
          parents: [carpetaId],
          description: `${datos.responsable} | ${datos.frente} | HSE`
        },
        media: {
          mimeType: datos.mimeType || "application/octet-stream",
          body: require("stream").Readable.from(buffer)
        },
        fields: "id,webViewLink",
        supportsAllDrives: true
      });
      urlArchivo = uploadRes.data.webViewLink || "";
    }

    await registrarEnSheetHSE(sheets, datos, urlArchivo, ahora, hoy);

    // Email no bloquea
    try {
      await enviarEmailHSE(datos, urlArchivo, ahora);
    } catch (emailErr) {
      console.warn("[Email HSE] Fallo (no critico):", emailErr.message);
    }

    res.json({ ok: true });
  } catch (err) {
    console.error("[ERROR] reporte-hse:", err.message);
    res.json({ ok: false, error: err.message });
  }
});

// ─── DIAGNOSTICO ─────────────────────────────────────────────────────────────
app.get("/api/diagnostico", async (req, res) => {
  const resultado = {
    env: {
      CLIENT_ID:        process.env.GOOGLE_CLIENT_ID     ? "OK" : "FALTA",
      CLIENT_SECRET:    process.env.GOOGLE_CLIENT_SECRET ? "OK" : "FALTA",
      REFRESH_TOKEN:    process.env.GOOGLE_REFRESH_TOKEN
        ? "OK (" + process.env.GOOGLE_REFRESH_TOKEN.substring(0, 10) + "...)" : "FALTA",
      SPREADSHEET_ID:     process.env.SPREADSHEET_ID     || "FALTA",
      SPREADSHEET_ID_HSE: process.env.SPREADSHEET_ID_HSE || "FALTA"
    },
    drive: null,
    sheets: null,
    sheetsHSE: null
  };
  try {
    const auth  = getGoogleAuth();
    const drive = google.drive({ version: "v3", auth });
    const r     = await drive.files.list({ pageSize: 1, fields: "files(id,name)" });
    resultado.drive = { ok: true, archivos: r.data.files.length };
  } catch (e) {
    resultado.drive = { ok: false, error: e.message, code: e.code };
  }
  try {
    const auth   = getGoogleAuth();
    const sheets = google.sheets({ version: "v4", auth });
    if (SPREADSHEET_ID) {
      await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
      resultado.sheets = { ok: true };
    } else {
      resultado.sheets = { ok: false, error: "SPREADSHEET_ID no configurado" };
    }
  } catch (e) {
    resultado.sheets = { ok: false, error: e.message, code: e.code };
  }
  try {
    const auth   = getGoogleAuth();
    const sheets = google.sheets({ version: "v4", auth });
    if (SPREADSHEET_ID_HSE) {
      await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID_HSE });
      resultado.sheetsHSE = { ok: true };
    } else {
      resultado.sheetsHSE = { ok: false, error: "SPREADSHEET_ID_HSE no configurado" };
    }
  } catch (e) {
    resultado.sheetsHSE = { ok: false, error: e.message, code: e.code };
  }
  res.json(resultado);
});

// ─── POST /api/reporte — CAMPO ────────────────────────────────────────────────
app.post("/api/reporte", async (req, res) => {
  const datos = req.body;
  try {
    const ahora = fechaLima().replace("T", " ");
    const hoy   = datos.fecha || fechaCorta();
    let urlArchivo = "";

    const auth   = getGoogleAuth();
    const drive  = google.drive({ version: "v3", auth });
    const sheets = google.sheets({ version: "v4", auth });

    if (datos.archivoBase64 && datos.archivoBase64.length > 100) {
      const carpetaId = await obtenerCarpetaDrive(
        drive, datos.sector, datos.subcategoria,
        datos.frente, datos.tipoReporte, hoy
      );
      const buffer    = Buffer.from(datos.archivoBase64, "base64");
      const uploadRes = await drive.files.create({
        requestBody: {
          name: datos.nombreArchivo, parents: [carpetaId],
          description: `${datos.responsable} | ${datos.frente} | ${datos.descripcion}`
        },
        media: {
          mimeType: datos.mimeType || "application/octet-stream",
          body: require("stream").Readable.from(buffer)
        },
        fields: "id,webViewLink",
        supportsAllDrives: true
      });
      urlArchivo = uploadRes.data.webViewLink || "";
    }

    await registrarEnSheet(sheets, datos, urlArchivo, ahora, hoy);

    // Email no bloquea
    try {
      await enviarEmail(datos, urlArchivo, ahora);
    } catch (emailErr) {
      console.warn("[Email] Fallo (no critico):", emailErr.message);
    }

    // PDS no bloquea: si falla, el reporte ya fue guardado
    if (datos.tipoReporte && datos.tipoReporte.toLowerCase().includes("diario")) {
      const yearMonth = hoy.substring(0, 7); // "2026-04"
      generatePDS(sheets, drive, SPREADSHEET_ID, datos.sector, yearMonth)
        .then(() => console.log("[PDS] Generado exitosamente"))
        .catch(pdsErr => console.warn("[PDS] Fallo (no critico):", pdsErr.message));
    }

    res.json({ ok: true });
  } catch (err) {
    console.error("[ERROR] procesarReporte:", err.message);
    res.json({ ok: false, error: err.message });
  }
});

// ─── GET /api/datos — CAMPO ───────────────────────────────────────────────────
app.get("/api/datos", async (req, res) => {
  if (!SPREADSHEET_ID) return res.json({ reportes: [] });
  try {
    const auth   = getGoogleAuth();
    const sheets = google.sheets({ version: "v4", auth });
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID, range: "RAW_DATA!A2:W"
    });
    const filas = result.data.values || [];
    const reportes = filas.map(r => ({
      fecha: r[1] || "", responsable: r[2] || "", sector: r[4] || "",
      frente: r[6] || "", tipoReporte: r[15] || "", link: r[21] || "", semana: r[22] || ""
    })).reverse();
    res.json({ reportes });
  } catch (err) {
    console.error("Error obtenerDatos:", err.message);
    res.json({ error: err.message });
  }
});

// ─── GET /api/datos-hse — HSE ────────────────────────────────────────────────
app.get("/api/datos-hse", async (req, res) => {
  if (!SPREADSHEET_ID_HSE) return res.json({ reportes: [] });
  try {
    const auth   = getGoogleAuth();
    const sheets = google.sheets({ version: "v4", auth });
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID_HSE, range: "RAW_DATA!A2:K"
    });
    const filas = result.data.values || [];
    const reportes = filas.map(r => ({
      fecha: r[1] || "", responsable: r[2] || "", sector: r[4] || "",
      frente: r[6] || "", tipoReporte: r[7] || "", link: r[9] || "", semana: r[10] || ""
    })).reverse();
    res.json({ reportes });
  } catch (err) {
    console.error("Error obtenerDatosHSE:", err.message);
    res.json({ error: err.message });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// PORTAL TGP-BV — ENDPOINTS
// ═══════════════════════════════════════════════════════════════════════════

const fs = require("fs");
const PORTAL_DATA_FILE = path.join(__dirname, "portal-data.json");

// Utilidades para portal-data.json
function readPortalData() {
  try {
    if (fs.existsSync(PORTAL_DATA_FILE)) {
      return JSON.parse(fs.readFileSync(PORTAL_DATA_FILE, "utf8"));
    }
  } catch (e) { console.warn("[Portal] Error leyendo portal-data.json:", e.message); }
  return { certificaciones: {}, ganttFiles: {}, fechasInicioFrente: {} };
}

function writePortalData(data) {
  fs.writeFileSync(PORTAL_DATA_FILE, JSON.stringify(data, null, 2), "utf8");
}

// Cache para RAW_DATA (evita llamadas excesivas a Google)
let rawDataCache = { campo: null, hse: null, timestamp: 0 };
const CACHE_TTL = 5 * 60 * 1000; // 5 minutos

async function getCachedData() {
  const now = Date.now();
  if (rawDataCache.campo && (now - rawDataCache.timestamp) < CACHE_TTL) {
    return rawDataCache;
  }
  const auth   = getGoogleAuth();
  const sheets = google.sheets({ version: "v4", auth });

  let campoDatos = [];
  if (SPREADSHEET_ID) {
    try {
      const r = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID, range: "RAW_DATA!A2:W"
      });
      campoDatos = r.data.values || [];
    } catch (e) { console.warn("[Cache] Error campo:", e.message); }
  }

  let hseDatos = [];
  if (SPREADSHEET_ID_HSE) {
    try {
      const r = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID_HSE, range: "RAW_DATA!A2:K"
      });
      hseDatos = r.data.values || [];
    } catch (e) { console.warn("[Cache] Error HSE:", e.message); }
  }

  rawDataCache = { campo: campoDatos, hse: hseDatos, timestamp: now };
  return rawDataCache;
}

// ─── GET /api/portal/resumen ────────────────────────────────────────────────
app.get("/api/portal/resumen", async (req, res) => {
  try {
    const cached = await getCachedData();
    const filas = cached.campo || [];
    const filasHSE = cached.hse || [];
    const portalData = readPortalData();
    const hoy = fechaCorta();
    const mes = hoy.substring(0, 7);
    const hace7 = new Date(Date.now() - 7 * 86400000).toISOString().substring(0, 10);
    const hace48h = new Date(Date.now() - 48 * 3600000).toISOString().substring(0, 10);

    // Reportes formateados
    const reportesCampo = filas.map(r => ({
      fecha: r[1]||"", responsable: r[2]||"", sector: r[4]||"",
      frente: r[6]||"", tipoReporte: r[15]||"", avance: parseFloat(r[17])||0,
      link: r[21]||"", semana: r[22]||""
    })).reverse();

    const reportesHSE = filasHSE.map(r => ({
      fecha: r[1]||"", responsable: r[2]||"", sector: r[4]||"",
      frente: r[6]||"", tipoReporte: r[7]||"", link: r[9]||"", semana: r[10]||""
    })).reverse();

    // Calcular datos por sector
    const sectores = {};
    const alertas = [];

    ["Costa", "Sierra", "Selva"].forEach(sector => {
      const frentesDelSector = FRENTES[`${sector}_Geotecnia`] || [];
      const reportesSector = filas.filter(r => r[4] === sector);
      const reportesMes = reportesSector.filter(r => (r[1]||"").substring(0,7) === mes).length;

      // Ultimo reporte del sector
      const fechas = reportesSector.map(r => r[1]||"").filter(f => f).sort().reverse();
      const ultimoReporte = fechas[0] ? fechas[0].substring(0,10) : "-";

      // Supervisores activos (ultimos 7 dias)
      const supsSet = new Set();
      reportesSector.filter(r => (r[1]||"") >= hace7).forEach(r => { if(r[2]) supsSet.add(r[2]); });

      // Datos por frente
      const frentesData = {};
      let sumaAvance = 0;
      let frentesConAvance = 0;
      let frentesActivos = 0;

      frentesDelSector.forEach(frente => {
        const repFrente = filas.filter(r => r[4] === sector && r[6] === frente);
        const repFrenteReciente = repFrente.filter(r => (r[1]||"") >= hace7);

        // Avance: tomar el ultimo % reportado
        const conAvance = repFrente.filter(r => r[17] && parseFloat(r[17]) > 0)
          .sort((a,b) => (b[1]||"").localeCompare(a[1]||""));
        const avance = conAvance.length > 0 ? parseFloat(conAvance[0][17]) : 0;

        if (avance > 0) { sumaAvance += avance; frentesConAvance++; }
        if (repFrenteReciente.length > 0) frentesActivos++;

        // Supervisores en este frente
        const supsFrSet = new Set();
        repFrenteReciente.forEach(r => { if(r[2]) supsFrSet.add(r[2]); });

        // Dias activo (desde primer reporte)
        const fechasFrente = repFrente.map(r => r[1]||"").filter(f=>f).sort();
        const primerReporte = fechasFrente[0] || "";
        const diasActivo = primerReporte
          ? Math.round((new Date() - new Date(primerReporte)) / 86400000)
          : 0;

        const ultimoRepFrente = fechasFrente.length > 0
          ? fechasFrente[fechasFrente.length-1].substring(0,10) : "-";

        // Alertas
        if (repFrente.length > 0 && ultimoRepFrente < hace48h) {
          alertas.push({
            severity: "high",
            message: sector + " / " + frente.substring(0,35) + " — Sin reportes en 48+ horas"
          });
        }

        frentesData[frente] = {
          avance: avance,
          supervisores: Array.from(supsFrSet),
          diasActivo: diasActivo,
          ultimoReporte: ultimoRepFrente,
          totalReportes: repFrente.length
        };
      });

      const avanceSector = frentesConAvance > 0 ? sumaAvance / frentesConAvance : 0;

      sectores[sector] = {
        avance: avanceSector,
        frentesActivos: frentesActivos,
        supervisores: Array.from(supsSet),
        reportesMes: reportesMes,
        ultimoReporte: ultimoReporte,
        frentes: frentesData
      };
    });

    // Avance global
    const avances = Object.values(sectores).map(s => s.avance).filter(a => a > 0);
    const avanceGlobal = avances.length > 0 ? avances.reduce((a,b)=>a+b,0) / avances.length : 0;

    // Alertas adicionales: supervisores inactivos
    SUPERVISORES.forEach(sup => {
      const repSup = filas.filter(r => r[2] === sup.nombre);
      if (repSup.length > 0) {
        const ultimoSup = repSup.map(r => r[1]||"").sort().reverse()[0] || "";
        const hace72h = new Date(Date.now() - 72 * 3600000).toISOString().substring(0, 10);
        if (ultimoSup < hace72h) {
          alertas.push({
            severity: "low",
            message: sup.nombre + " — Sin actividad en 72+ horas"
          });
        }
      }
    });

    // Curvas S por sector (agrupado por semana)
    const scurves = {};
    ["Costa", "Sierra", "Selva"].forEach(sector => {
      const repSector = filas.filter(r => r[4] === sector && r[17] && parseFloat(r[17]) > 0);
      const porSemana = {};
      repSector.forEach(r => {
        const sem = r[22] || "0";
        const av = parseFloat(r[17]) || 0;
        if (!porSemana[sem] || av > porSemana[sem]) porSemana[sem] = av;
      });
      const semanas = Object.keys(porSemana).sort((a,b) => parseInt(a)-parseInt(b));
      scurves[sector] = {
        labels: semanas.map(s => "Sem " + s),
        real: semanas.map(s => porSemana[s]),
        planificado: semanas.map(() => 0) // sin datos planificados aun
      };
    });

    res.json({
      reportesCampo, reportesHSE, sectores, avanceGlobal, alertas, scurves
    });
  } catch (err) {
    console.error("[Portal] Error resumen:", err.message);
    res.json({ error: err.message, sectores: {}, alertas: [], scurves: {}, reportesCampo: [], reportesHSE: [] });
  }
});

// ─── GET /api/portal/certificaciones ────────────────────────────────────────
app.get("/api/portal/certificaciones", (req, res) => {
  const data = readPortalData();
  res.json(data.certificaciones || {});
});

// ─── POST /api/portal/certificacion ─────────────────────────────────────────
app.post("/api/portal/certificacion", (req, res) => {
  try {
    const data = readPortalData();
    data.certificaciones = req.body;
    writePortalData(data);
    res.json({ ok: true });
  } catch (err) {
    console.error("[Portal] Error guardando certificacion:", err.message);
    res.json({ ok: false, error: err.message });
  }
});

// ─── GET /api/portal/gantt-list ─────────────────────────────────────────────
app.get("/api/portal/gantt-list", (req, res) => {
  const data = readPortalData();
  res.json(data.ganttFiles || {});
});

// ─── POST /api/portal/gantt-upload ──────────────────────────────────────────
app.post("/api/portal/gantt-upload", (req, res) => {
  try {
    const { sector, filename, mimeType, base64 } = req.body;
    const data = readPortalData();
    if (!data.ganttFiles) data.ganttFiles = {};
    data.ganttFiles[sector] = {
      filename, mimeType, base64, uploadDate: fechaCorta()
    };
    writePortalData(data);
    res.json({ ok: true });
  } catch (err) {
    console.error("[Portal] Error subiendo gantt:", err.message);
    res.json({ ok: false, error: err.message });
  }
});

// Servir templates (logos)
app.use("/templates", express.static(path.join(__dirname, "templates")));

// ─── HEALTH ──────────────────────────────────────────────────────────────────
app.get("/health", (req, res) => {
  res.json({ status: "ok", version: "10.1.0", time: new Date().toISOString() });
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`TGP Reportes v10.1 corriendo en puerto ${PORT}`);
});
