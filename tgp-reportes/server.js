// ═══════════════════════════════════════════════════════════════════════════
// SISTEMA DE REPORTES DE CAMPO — server.js v10 (+ HSE + PDS + Portal)
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
app.get("/api/supervisores",   (req, res) => res.json(SUPERVISORES));
app.get("/api/todos-frentes",  (req, res) => res.json(FRENTES));
app.get("/api/frentes", (req, res) => {
  const { sector, subcat } = req.query;
  res.json(FRENTES[`${sector}_${subcat}`] || []);
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
    const hoy   = fechaCorta();
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
    const hoy   = fechaCorta();
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

// ─── HEALTH ──────────────────────────────────────────────────────────────────
app.get("/health", (req, res) => {
  res.json({ status: "ok", version: "10.0.0", time: new Date().toISOString() });
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`TGP Reportes v10 corriendo en puerto ${PORT}`);
});
