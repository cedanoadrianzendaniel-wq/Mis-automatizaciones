// ═══════════════════════════════════════════════════════════════════════════
// SISTEMA DE REPORTES DE CAMPO — server.js v4 (Node.js + Express)
// Migrado desde Google Apps Script
// ═══════════════════════════════════════════════════════════════════════════

require("dotenv").config();
const express    = require("express");
const cors       = require("cors");
const path       = require("path");
const { google } = require("googleapis");
const nodemailer = require("nodemailer");

const app  = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: "60mb" }));
app.use(express.urlencoded({ extended: true, limit: "60mb" }));
app.use(express.static(path.join(__dirname, "public")));

// ─── CONFIG ─────────────────────────────────────────────────────────────────
const EMAIL_COORDINADOR = process.env.EMAIL_COORDINADOR ||
  "yuri.arangoitia@bureauveritas.com, daniel.cedano@bureauveritas.com, gustavo.fernandez@bureauveritas.com, fiorella.diaz@bureauveritas.com";
const CARPETA_RAIZ      = process.env.CARPETA_RAIZ || "Reportes de Campo 2026";
  const CARPETA_RAIZ_ID   = process.env.CARPETA_RAIZ_ID || "";
const CLAVE_DASHBOARD   = process.env.CLAVE_DASHBOARD || "campo2026";
const SPREADSHEET_ID    = process.env.SPREADSHEET_ID || "";

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

const SUPERVISORES = [
  { nombre: "CRISTHIAN BAQUERIZO",       sector: "", subcategoria: "" },
  { nombre: "WALTER JESUS",              sector: "", subcategoria: "" },
  { nombre: "LIZ GUERRERO",              sector: "", subcategoria: "" },
  { nombre: "CARLOS DE LA CRUZ",         sector: "", subcategoria: "" },
  { nombre: "ROY HERRADA",               sector: "", subcategoria: "" },
  { nombre: "ABEL SANCHEZ QUIHUI",       sector: "", subcategoria: "" },
  { nombre: "SAMUEL JARA MAYTA",         sector: "", subcategoria: "" },
  { nombre: "NIKOLAI ARANGOITIA",        sector: "", subcategoria: "" },
  { nombre: "ROGELIO CHAMPI CHOQUEPATA", sector: "", subcategoria: "" },
  { nombre: "JHON FUENTES",              sector: "", subcategoria: "" },
  { nombre: "RUBEN NUÑEZ",               sector: "", subcategoria: "" },
  { nombre: "JORDAN GALLO",              sector: "", subcategoria: "" },
  { nombre: "ABRAHAM JIMENEZ",           sector: "", subcategoria: "" },
  { nombre: "NEISSER MAMANI",            sector: "", subcategoria: "" },
  { nombre: "PAUL PACSI ALAVE",          sector: "", subcategoria: "" },
  { nombre: "DANIEL ATAYUPANQUI TARCO",  sector: "", subcategoria: "" },
  { nombre: "CARLOS PUENTE",             sector: "", subcategoria: "" }
];

// ─── GOOGLE AUTH (Service Account) ──────────────────────────────────────────
function getGoogleAuth() {
  const credentials = process.env.GOOGLE_CREDENTIALS_JSON
    ? JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON)
    : null;
  if (!credentials) {
    throw new Error("GOOGLE_CREDENTIALS_JSON no configurado en .env");
  }
  return new google.auth.GoogleAuth({
    credentials,
    scopes: [
      "https://www.googleapis.com/auth/drive",
      "https://www.googleapis.com/auth/spreadsheets"
    ]
  });
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


async function obtenerCarpetaDrive(drive, sector, subcategoria, frente, tipo, fecha) {
      let raizId;
      if (CARPETA_RAIZ_ID) {
              raizId = CARPETA_RAIZ_ID;
      } else {
              const qRaiz = `name='${CARPETA_RAIZ}' and mimeType='application/vnd.google-apps.folder' and 'root' in parents and trashed=false`;
              const raizRes = await drive.files.list({ q: qRaiz, fields: "files(id)", pageSize: 1 });
              if (raizRes.data.files.length > 0) {
                        raizId = raizRes.data.files[0].id;
              } else {
                        const r = await drive.files.create({
                                    requestBody: { name: CARPETA_RAIZ, mimeType: "application/vnd.google-apps.folder" },
                                    fields: "id"
                        });
                        raizId = r.data.id;
              }
      }
  const sectorId = await carpetaEnPadre(drive, sector,       raizId);
  const subcatId = await carpetaEnPadre(drive, subcategoria, sectorId);
  const frenteId = await carpetaEnPadre(drive, frente,       subcatId);
  const tipoId   = await carpetaEnPadre(drive, tipo,         frenteId);
  const fechaId  = await carpetaEnPadre(drive, fecha,        tipoId);
  return fechaId;
}

// ─── SHEETS: registrar fila ──────────────────────────────────────────────────
async function registrarEnSheet(sheets, datos, urlArchivo, ahora, fecha) {
  const SHEET_NAME = "RAW_DATA";
  const spreadsheetId = SPREADSHEET_ID;
  if (!spreadsheetId) {
    console.warn("SPREADSHEET_ID no configurado — omitiendo registro en Sheets");
    return;
  }
  let sheetData;
  try {
    sheetData = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${SHEET_NAME}!A1:W1`
    });
  } catch (e) {
    sheetData = { data: { values: [] } };
  }
  const hasHeader = sheetData.data.values && sheetData.data.values.length > 0;
  if (!hasHeader) {
    await sheets.spreadsheets.values.append({
      spreadsheetId,
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
    spreadsheetId,
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

// ─── EMAIL ───────────────────────────────────────────────────────────────────
async function enviarEmail(datos, urlArchivo, ahora) {
  if (!process.env.SMTP_HOST) { console.warn("SMTP no configurado — omitiendo email"); return; }
  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST, port: parseInt(process.env.SMTP_PORT || "587"),
    secure: process.env.SMTP_SECURE === "true",
    auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS }
  });
  const adjunto = urlArchivo ? `<br><a href="${urlArchivo}" style="color:#2E75B6;font-weight:bold">Ver archivo en Drive</a>` : "";
  const asunto = `[Campo] ${datos.tipoReporte} | ${datos.sector} | ${datos.frente} | ${datos.responsable}`;
  const fila = (k, v) => `<div style="padding:6px 0;border-bottom:1px solid #eee;font-size:13px"><b style="color:#555">${k}:</b> ${v || "-"}</div>`;
  const cuerpo = `<div style="font-family:Arial,sans-serif;max-width:580px">
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

// ─── RUTAS ───────────────────────────────────────────────────────────────────
app.get("/", (req, res) => { res.sendFile(path.join(__dirname, "public", "formulario.html")); });
app.get("/dashboard", (req, res) => { res.sendFile(path.join(__dirname, "public", "dashboard.html")); });
app.get("/api/supervisores", (req, res) => { res.json(SUPERVISORES); });
app.get("/api/frentes", (req, res) => {
  const { sector, subcat } = req.query;
  res.json(FRENTES[`${sector}_${subcat}`] || []);
});
app.get("/api/todos-frentes", (req, res) => { res.json(FRENTES); });
app.post("/api/verificar-clave", (req, res) => {
  const { clave } = req.body;
  res.json({ ok: String(clave).trim() === CLAVE_DASHBOARD });
});

app.post("/api/reporte", async (req, res) => {
  const datos = req.body;
  try {
    const ahora = fechaLima().replace("T", " ");
    const hoy = fechaCorta();
    let urlArchivo = "";
    const auth = getGoogleAuth();
    const client = await auth.getClient();
    const drive = google.drive({ version: "v3", auth: client });
    const sheets = google.sheets({ version: "v4", auth: client });
    if (datos.archivoBase64 && datos.archivoBase64.length > 100) {
      const carpetaId = await obtenerCarpetaDrive(drive, datos.sector, datos.subcategoria, datos.frente, datos.tipoReporte, hoy);
      const buffer = Buffer.from(datos.archivoBase64, "base64");
      const uploadRes = await drive.files.create({
        requestBody: {
          name: datos.nombreArchivo, parents: [carpetaId],
          description: `${datos.responsable} | ${datos.frente} | ${datos.descripcion}`
        },
        media: { mimeType: datos.mimeType || "application/octet-stream", body: require("stream").Readable.from(buffer) },
        fields: "id,webViewLink",
        supportsAllDrives: true
      });
      urlArchivo = uploadRes.data.webViewLink || "";
    }
    await registrarEnSheet(sheets, datos, urlArchivo, ahora, hoy);
    await enviarEmail(datos, urlArchivo, ahora);
    res.json({ ok: true });
  } catch (err) {
    console.error("Error procesarReporte:", err.message);
    res.json({ ok: false, error: err.message });
  }
});

app.get("/api/datos", async (req, res) => {
  const spreadsheetId = SPREADSHEET_ID;
  if (!spreadsheetId) return res.json({ reportes: [] });
  try {
    const auth = getGoogleAuth();
    const client = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });
    const result = await sheets.spreadsheets.values.get({ spreadsheetId, range: "RAW_DATA!A2:W" });
    const filas = result.data.values || [];
    const reportes = filas.map(r => ({
      fecha: r[1] || "", responsable: r[2] || "", sector: r[4] || "",
      frente: r[6] || "", link: r[21] || "", semana: r[22] || ""
    })).reverse();
    res.json({ reportes });
  } catch (err) {
    console.error("Error obtenerDatos:", err.message);
    res.json({ error: err.message });
  }
});

app.get("/health", (req, res) => {
  res.json({ status: "ok", version: "4.0.0", time: new Date().toISOString() });
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`TGP Reportes de Campo v4 corriendo en puerto ${PORT}`);
});
