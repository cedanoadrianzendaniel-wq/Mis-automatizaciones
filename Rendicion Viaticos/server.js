// ═══════════════════════════════════════════════════════════════════════════
// PLATAFORMA DE RENDICIÓN DE VIÁTICOS — server.js v1.0
// Escaneo OCR de comprobantes + Generación Excel/PDF + Envío por correo
// ═══════════════════════════════════════════════════════════════════════════

require("dotenv").config();
const express    = require("express");
const cors       = require("cors");
const path       = require("path");
const multer     = require("multer");
const Tesseract  = require("tesseract.js");
const ExcelJS    = require("exceljs");
const PDFDocument = require("pdfkit");
const nodemailer = require("nodemailer");
const fs         = require("fs");

const app  = express();
const PORT = process.env.PORT || 3001;

// ─── Middleware ──────────────────────────────────────────────────────────────
app.use(cors());
app.use(express.json({ limit: "60mb" }));
app.use(express.urlencoded({ extended: true, limit: "60mb" }));
app.use(express.static(path.join(__dirname, "public")));

// ─── Multer para subida de imágenes ─────────────────────────────────────────
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const dir = path.join(__dirname, "uploads");
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    cb(null, dir);
  },
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname);
    cb(null, `comprobante_${Date.now()}${ext}`);
  }
});
const upload = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowed = /jpeg|jpg|png|gif|bmp|webp|pdf/;
    const ext = allowed.test(path.extname(file.originalname).toLowerCase());
    const mime = allowed.test(file.mimetype);
    if (ext && mime) return cb(null, true);
    cb(new Error("Solo se permiten imágenes (JPG, PNG, BMP, WEBP) o PDF"));
  }
});

// ─── CONFIG ─────────────────────────────────────────────────────────────────
const EMAIL_USER          = process.env.EMAIL_USER || "";
const EMAIL_PASS          = process.env.EMAIL_PASS || "";
const EMAIL_CONTABILIDAD  = process.env.EMAIL_CONTABILIDAD || "";
const EMPRESA_NOMBRE      = process.env.EMPRESA_NOMBRE || "Bureau Veritas";

// ─── RUTA PRINCIPAL ─────────────────────────────────────────────────────────
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

// ═══════════════════════════════════════════════════════════════════════════
// ENDPOINT: OCR — Escanear comprobante
// ═══════════════════════════════════════════════════════════════════════════
app.post("/api/escanear", upload.single("comprobante"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ ok: false, error: "No se recibió archivo" });
    }

    const filePath = req.file.path;

    const { data: { text } } = await Tesseract.recognize(filePath, "spa", {
      logger: m => {} // silenciar logs
    });

    // Extraer datos del comprobante
    const datos = extraerDatosComprobante(text);

    // Limpiar archivo temporal
    fs.unlink(filePath, () => {});

    res.json({
      ok: true,
      textoCompleto: text,
      datos
    });
  } catch (err) {
    console.error("Error OCR:", err);
    res.status(500).json({ ok: false, error: "Error al procesar el comprobante" });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// ENDPOINT: Generar Excel
// ═══════════════════════════════════════════════════════════════════════════
app.post("/api/generar-excel", async (req, res) => {
  try {
    const { empleado, periodo, viaticoAsignado, centroCostos, nroContrato, comprobantes, declaraciones, movilidad } = req.body;
    const montoAsignado = parseFloat(viaticoAsignado) || 0;

    const workbook = new ExcelJS.Workbook();
    workbook.creator = EMPRESA_NOMBRE;
    workbook.created = new Date();

    // ── Hoja 1: Comprobantes ──────────────────────────────────────────────
    const wsComp = workbook.addWorksheet("Comprobantes de Pago");
    agregarEncabezado(wsComp, empleado, periodo, "RENDICIÓN DE COMPROBANTES DE PAGO", { centroCostos, nroContrato });

    wsComp.getRow(5).values = ["N°", "Fecha", "Tipo Comprobante", "N° Comprobante", "Concepto / Detalle", "Subtotal (S/)", "IGV (S/)", "Total (S/)"];
    const headerRow1 = wsComp.getRow(5);
    estiloEncabezadoTabla(headerRow1, 8);

    wsComp.getColumn(1).width = 6;
    wsComp.getColumn(2).width = 14;
    wsComp.getColumn(3).width = 22;
    wsComp.getColumn(4).width = 20;
    wsComp.getColumn(5).width = 35;
    wsComp.getColumn(6).width = 14;
    wsComp.getColumn(7).width = 14;
    wsComp.getColumn(8).width = 14;

    let totalComp = 0;
    let totalSubtotal = 0;
    let totalIGV = 0;
    if (comprobantes && comprobantes.length > 0) {
      comprobantes.forEach((c, i) => {
        const sub = parseFloat(c.subtotal) || 0;
        const igv = parseFloat(c.igv) || 0;
        const tot = parseFloat(c.monto) || 0;
        const row = wsComp.addRow([i + 1, c.fecha, c.tipo, c.numero, c.concepto, sub, igv, tot]);
        [6, 7, 8].forEach(col => { row.getCell(col).numFmt = '#,##0.00'; });
        row.eachCell(cell => {
          cell.border = borderThin();
          cell.alignment = { vertical: "middle" };
        });
        totalSubtotal += sub;
        totalIGV += igv;
        totalComp += tot;
      });
    }

    const totalRow1 = wsComp.addRow(["", "", "", "", "TOTALES", totalSubtotal, totalIGV, totalComp]);
    [5, 6, 7, 8].forEach(col => {
      totalRow1.getCell(col).font = { bold: true, size: 11 };
      if (col >= 6) totalRow1.getCell(col).numFmt = '#,##0.00';
    });
    totalRow1.eachCell(cell => { cell.border = borderThin(); });

    // ── Hoja 2: Declaraciones Juradas ─────────────────────────────────────
    const wsDJ = workbook.addWorksheet("Declaraciones Juradas");
    agregarEncabezado(wsDJ, empleado, periodo, "DECLARACIONES JURADAS — GASTOS SIN COMPROBANTE", { centroCostos, nroContrato });

    wsDJ.getRow(5).values = ["N°", "Fecha", "Concepto / Detalle", "Motivo (sin comprobante)", "Monto (S/)"];
    const headerRow2 = wsDJ.getRow(5);
    estiloEncabezadoTabla(headerRow2, 5);

    wsDJ.getColumn(1).width = 6;
    wsDJ.getColumn(2).width = 14;
    wsDJ.getColumn(3).width = 35;
    wsDJ.getColumn(4).width = 35;
    wsDJ.getColumn(5).width = 16;

    let totalDJ = 0;
    if (declaraciones && declaraciones.length > 0) {
      declaraciones.forEach((d, i) => {
        const row = wsDJ.addRow([i + 1, d.fecha, d.concepto, d.motivo, parseFloat(d.monto) || 0]);
        row.getCell(5).numFmt = '#,##0.00';
        row.eachCell(cell => {
          cell.border = borderThin();
          cell.alignment = { vertical: "middle" };
        });
        totalDJ += parseFloat(d.monto) || 0;
      });
    }

    const totalRow2 = wsDJ.addRow(["", "", "", "TOTAL", totalDJ]);
    totalRow2.getCell(4).font = { bold: true, size: 11 };
    totalRow2.getCell(5).font = { bold: true, size: 11 };
    totalRow2.getCell(5).numFmt = '#,##0.00';
    totalRow2.eachCell(cell => { cell.border = borderThin(); });

    // ── Hoja 3: Movilización ──────────────────────────────────────────────
    const wsMov = workbook.addWorksheet("Movilización");
    agregarEncabezado(wsMov, empleado, periodo, "DECLARACIÓN DE MOVILIZACIÓN", { centroCostos, nroContrato });

    wsMov.getRow(5).values = ["N°", "Fecha", "Origen", "Destino", "Medio de Transporte", "Motivo", "Monto (S/)"];
    const headerRow3 = wsMov.getRow(5);
    estiloEncabezadoTabla(headerRow3, 7);

    wsMov.getColumn(1).width = 6;
    wsMov.getColumn(2).width = 14;
    wsMov.getColumn(3).width = 22;
    wsMov.getColumn(4).width = 22;
    wsMov.getColumn(5).width = 22;
    wsMov.getColumn(6).width = 30;
    wsMov.getColumn(7).width = 16;

    let totalMov = 0;
    if (movilidad && movilidad.length > 0) {
      movilidad.forEach((m, i) => {
        const row = wsMov.addRow([i + 1, m.fecha, m.origen, m.destino, m.transporte, m.motivo, parseFloat(m.monto) || 0]);
        row.getCell(7).numFmt = '#,##0.00';
        row.eachCell(cell => {
          cell.border = borderThin();
          cell.alignment = { vertical: "middle" };
        });
        totalMov += parseFloat(m.monto) || 0;
      });
    }

    const totalRow3 = wsMov.addRow(["", "", "", "", "", "TOTAL", totalMov]);
    totalRow3.getCell(6).font = { bold: true, size: 11 };
    totalRow3.getCell(7).font = { bold: true, size: 11 };
    totalRow3.getCell(7).numFmt = '#,##0.00';
    totalRow3.eachCell(cell => { cell.border = borderThin(); });

    // ── Hoja 4: Resumen ───────────────────────────────────────────────────
    const wsRes = workbook.addWorksheet("Resumen");
    agregarEncabezado(wsRes, empleado, periodo, "RESUMEN DE RENDICIÓN DE VIÁTICOS", { centroCostos, nroContrato });

    wsRes.getRow(5).values = ["Categoría", "Total (S/)"];
    const headerRow4 = wsRes.getRow(5);
    estiloEncabezadoTabla(headerRow4, 2);

    wsRes.getColumn(1).width = 40;
    wsRes.getColumn(2).width = 20;

    const totalGeneral = totalComp + totalDJ + totalMov;
    const saldo = montoAsignado - totalGeneral;

    const resData = [
      ["Viático Asignado", montoAsignado],
      ["Comprobantes de Pago", totalComp],
      ["Declaraciones Juradas", totalDJ],
      ["Movilización", totalMov],
      ["TOTAL GASTADO", totalGeneral],
      [saldo >= 0 ? "SALDO A FAVOR (devolver)" : "MONTO EXCEDIDO (por reembolsar)", Math.abs(saldo)]
    ];
    resData.forEach(([cat, monto], i) => {
      const row = wsRes.addRow([cat, monto]);
      row.getCell(2).numFmt = '#,##0.00';
      row.eachCell(cell => {
        cell.border = borderThin();
        cell.alignment = { vertical: "middle" };
      });
      if (i === 0) {
        row.getCell(1).font = { bold: true, size: 11 };
        row.getCell(2).font = { bold: true, size: 11, color: { argb: "FF1A56DB" } };
        row.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8F0FE" } };
        row.getCell(2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8F0FE" } };
      }
      if (i === 4) {
        row.getCell(1).font = { bold: true, size: 12 };
        row.getCell(2).font = { bold: true, size: 12, color: { argb: "FF0066CC" } };
        row.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8F0FE" } };
        row.getCell(2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8F0FE" } };
      }
      if (i === 5) {
        const color = saldo >= 0 ? "FF059669" : "FFDC2626";
        row.getCell(1).font = { bold: true, size: 12 };
        row.getCell(2).font = { bold: true, size: 14, color: { argb: color } };
        const bgColor = saldo >= 0 ? "FFD1FAE5" : "FFFEE2E2";
        row.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: bgColor } };
        row.getCell(2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: bgColor } };
      }
    });

    // ── Hoja 5: Asiento Contable PCGE ─────────────────────────────────────
    // Formato genérico Plan Contable General Empresarial (Perú) compatible
    // con la mayoría de ERPs contables (SAP B1, Defontana, Concar, Siigo,
    // Starsoft, Contasis, etc.) con adaptación mínima.
    const wsAsiento = workbook.addWorksheet("Asiento Contable PCGE");
    agregarEncabezado(wsAsiento, empleado, periodo, "ASIENTO CONTABLE — RENDICIÓN DE VIÁTICOS (PCGE)", { centroCostos, nroContrato });

    wsAsiento.getRow(5).values = [
      "Fecha", "Cuenta", "Glosa", "Debe (S/)", "Haber (S/)",
      "C. Costo", "Tipo Doc", "Serie-N°", "RUC Proveedor", "Razón Social"
    ];
    estiloEncabezadoTabla(wsAsiento.getRow(5), 10);

    wsAsiento.getColumn(1).width = 12;
    wsAsiento.getColumn(2).width = 10;
    wsAsiento.getColumn(3).width = 38;
    wsAsiento.getColumn(4).width = 13;
    wsAsiento.getColumn(5).width = 13;
    wsAsiento.getColumn(6).width = 12;
    wsAsiento.getColumn(7).width = 10;
    wsAsiento.getColumn(8).width = 18;
    wsAsiento.getColumn(9).width = 14;
    wsAsiento.getColumn(10).width = 30;

    // Mapeo SUNAT de tipos de documento
    const tipoDocSunat = {
      "Factura Electrónica": "01",
      "Boleta de Venta": "03",
      "Recibo por Honorarios": "02",
      "Nota de Crédito": "07",
      "Ticket": "12",
      "Otro": "00"
    };

    let totalDebe = 0;
    let totalHaber = 0;
    const cc = centroCostos || "";

    // Asientos por comprobantes (cargo a gasto + IGV, abono a entregas a rendir)
    if (comprobantes && comprobantes.length > 0) {
      comprobantes.forEach(c => {
        const sub = parseFloat(c.subtotal) || 0;
        const igv = parseFloat(c.igv) || 0;
        const tot = parseFloat(c.monto) || 0;
        const cuenta = c.cuenta || "631";
        const tipoDoc = tipoDocSunat[c.tipo] || "00";
        const glosa = `${c.concepto || c.tipo} - ${c.numero}`;

        // Cargo a gasto (subtotal o total si no hay IGV)
        const montoGasto = sub > 0 ? sub : tot;
        const filaGasto = wsAsiento.addRow([
          c.fecha, cuenta, glosa, montoGasto, 0,
          cc, tipoDoc, c.numero, c.ruc || "", c.razonSocial || ""
        ]);
        filaGasto.getCell(4).numFmt = '#,##0.00';
        filaGasto.getCell(5).numFmt = '#,##0.00';
        filaGasto.eachCell(cell => { cell.border = borderThin(); cell.alignment = { vertical: "middle" }; });
        totalDebe += montoGasto;

        // Cargo a IGV crédito fiscal (solo si hay IGV)
        if (igv > 0) {
          const filaIgv = wsAsiento.addRow([
            c.fecha, "40111", `IGV - ${c.numero}`, igv, 0,
            cc, tipoDoc, c.numero, c.ruc || "", c.razonSocial || ""
          ]);
          filaIgv.getCell(4).numFmt = '#,##0.00';
          filaIgv.getCell(5).numFmt = '#,##0.00';
          filaIgv.eachCell(cell => { cell.border = borderThin(); cell.alignment = { vertical: "middle" }; });
          totalDebe += igv;
        }
      });
    }

    // Asientos por declaraciones juradas (sin IGV, sin proveedor)
    if (declaraciones && declaraciones.length > 0) {
      declaraciones.forEach(d => {
        const monto = parseFloat(d.monto) || 0;
        const fila = wsAsiento.addRow([
          d.fecha, "6315", `DJ - ${d.concepto || ""}`, monto, 0,
          cc, "00", "", "", ""
        ]);
        fila.getCell(4).numFmt = '#,##0.00';
        fila.getCell(5).numFmt = '#,##0.00';
        fila.eachCell(cell => { cell.border = borderThin(); cell.alignment = { vertical: "middle" }; });
        totalDebe += monto;
      });
    }

    // Asientos por movilización
    if (movilidad && movilidad.length > 0) {
      movilidad.forEach(m => {
        const monto = parseFloat(m.monto) || 0;
        const fila = wsAsiento.addRow([
          m.fecha, "6311", `Movilidad ${m.origen || ""}-${m.destino || ""}`, monto, 0,
          cc, "00", "", "", ""
        ]);
        fila.getCell(4).numFmt = '#,##0.00';
        fila.getCell(5).numFmt = '#,##0.00';
        fila.eachCell(cell => { cell.border = borderThin(); cell.alignment = { vertical: "middle" }; });
        totalDebe += monto;
      });
    }

    // Contrapartida: abono a "Entregas a rendir cuenta" (cuenta 1411)
    if (totalDebe > 0) {
      const fechaCierre = new Date().toLocaleDateString("es-PE");
      const filaAbono = wsAsiento.addRow([
        fechaCierre, "1411", `Rendición viáticos - ${empleado || ""} - ${periodo || ""}`,
        0, totalDebe, cc, "", "", "", ""
      ]);
      filaAbono.getCell(4).numFmt = '#,##0.00';
      filaAbono.getCell(5).numFmt = '#,##0.00';
      filaAbono.eachCell(cell => { cell.border = borderThin(); cell.alignment = { vertical: "middle" }; });
      totalHaber = totalDebe;
    }

    // Fila de totales
    const filaTotales = wsAsiento.addRow(["", "", "TOTALES", totalDebe, totalHaber, "", "", "", "", ""]);
    filaTotales.getCell(3).font = { bold: true, size: 11 };
    filaTotales.getCell(4).font = { bold: true, size: 11 };
    filaTotales.getCell(5).font = { bold: true, size: 11 };
    filaTotales.getCell(4).numFmt = '#,##0.00';
    filaTotales.getCell(5).numFmt = '#,##0.00';
    filaTotales.eachCell(cell => { cell.border = borderThin(); });

    // Nota informativa al pie
    wsAsiento.addRow([]);
    const notaRow = wsAsiento.addRow(["Nota: Asiento generado en formato PCGE (Plan Contable General Empresarial - Perú). Adaptable a SAP B1, Defontana, Concar, Siigo, Starsoft, Contasis y otros ERPs."]);
    notaRow.getCell(1).font = { italic: true, size: 9, color: { argb: "FF64748B" } };
    wsAsiento.mergeCells(`A${notaRow.number}:J${notaRow.number}`);

    // ── Hoja 6: Registro de Compras (formato PLE SUNAT) ───────────────────
    const wsPle = workbook.addWorksheet("Registro Compras PLE");
    agregarEncabezado(wsPle, empleado, periodo, "REGISTRO DE COMPRAS — FORMATO PLE 8.1 (SUNAT)", { centroCostos, nroContrato });

    wsPle.getRow(5).values = [
      "Período", "CUO", "Fecha Emisión", "Fecha Vcto.",
      "Tipo CP", "Serie", "Número", "Tipo Doc. Prov.", "N° Doc. Prov.",
      "Razón Social", "B. Imponible Gravada", "IGV", "Importe Total",
      "Moneda", "Tipo Cambio"
    ];
    estiloEncabezadoTabla(wsPle.getRow(5), 15);

    [10, 8, 12, 12, 8, 8, 14, 10, 14, 30, 14, 12, 14, 8, 10].forEach((w, i) => {
      wsPle.getColumn(i + 1).width = w;
    });

    const periodoPle = (periodo || "").replace(/\D/g, "").padEnd(6, "0").slice(0, 6) || "000000";
    let cuo = 1;

    if (comprobantes && comprobantes.length > 0) {
      comprobantes.forEach(c => {
        const sub = parseFloat(c.subtotal) || 0;
        const igv = parseFloat(c.igv) || 0;
        const tot = parseFloat(c.monto) || 0;
        const tipoCp = tipoDocSunat[c.tipo] || "00";
        const tipoDocProv = (c.ruc || "").length === 11 ? "6" : ((c.ruc || "").length === 8 ? "1" : "0");

        // Separar serie y número (E001-00012345)
        let serie = "", numero = c.numero || "";
        if (numero.includes("-")) {
          const partes = numero.split("-");
          serie = partes[0];
          numero = partes.slice(1).join("-");
        }

        const fila = wsPle.addRow([
          periodoPle,
          String(cuo++).padStart(6, "0"),
          c.fecha, c.fecha,
          tipoCp, serie, numero,
          tipoDocProv, c.ruc || "",
          c.razonSocial || "",
          sub, igv, tot,
          "PEN", 1.000
        ]);
        [11, 12, 13].forEach(col => fila.getCell(col).numFmt = '#,##0.00');
        fila.getCell(15).numFmt = '0.000';
        fila.eachCell(cell => { cell.border = borderThin(); cell.alignment = { vertical: "middle" }; });
      });
    }

    wsPle.addRow([]);
    const notaPle = wsPle.addRow(["Nota: Formato compatible con PLE 8.1 SUNAT. Tipo Doc. Prov.: 6=RUC, 1=DNI, 0=Sin doc. Tipo CP: 01=Factura, 03=Boleta, 02=RxH, 07=NC, 12=Ticket."]);
    notaPle.getCell(1).font = { italic: true, size: 9, color: { argb: "FF64748B" } };
    wsPle.mergeCells(`A${notaPle.number}:O${notaPle.number}`);

    // Generar buffer y enviar
    const buffer = await workbook.xlsx.writeBuffer();
    const filename = `Rendicion_Viaticos_${(empleado || "").replace(/\s+/g, "_")}_${periodo || "sin_periodo"}.xlsx`;

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(Buffer.from(buffer));

  } catch (err) {
    console.error("Error generando Excel:", err);
    res.status(500).json({ ok: false, error: "Error al generar el archivo Excel" });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// ENDPOINT: Generar PDF
// ═══════════════════════════════════════════════════════════════════════════
app.post("/api/generar-pdf", async (req, res) => {
  try {
    const { empleado, periodo, centroCostos, nroContrato, firmaEmpleado, firmaAprobador, comprobantes, declaraciones, movilidad } = req.body;

    const doc = new PDFDocument({ size: "A4", margin: 40, bufferPages: true });
    const chunks = [];
    doc.on("data", chunk => chunks.push(chunk));
    doc.on("end", () => {
      const buffer = Buffer.concat(chunks);
      const filename = `Rendicion_Viaticos_${(empleado || "").replace(/\s+/g, "_")}_${periodo || "sin_periodo"}.pdf`;
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
      res.send(buffer);
    });

    const azul = "#1a56db";
    const grisClaro = "#f3f4f6";

    // ── Encabezado ──────────────────────────────────────────────────────
    doc.rect(0, 0, 595.28, 70).fill(azul);
    doc.fontSize(20).fillColor("#ffffff").text("RENDICIÓN DE VIÁTICOS", 40, 20, { align: "center" });
    doc.fontSize(10).text(EMPRESA_NOMBRE, 40, 45, { align: "center" });

    doc.fillColor("#333333");
    doc.moveDown(2);
    doc.fontSize(11).text(`Empleado: ${empleado || "—"}`, 40);
    doc.text(`Período: ${periodo || "—"}`);
    if (centroCostos) doc.text(`Centro de Costos: ${centroCostos}`);
    if (nroContrato) doc.text(`N° Contrato: ${nroContrato}`);
    doc.text(`Fecha de emisión: ${new Date().toLocaleDateString("es-PE")}`);
    doc.moveDown(1);

    // ── Sección Comprobantes ────────────────────────────────────────────
    let totalComp = 0;
    doc.fontSize(14).fillColor(azul).text("1. COMPROBANTES DE PAGO", 40);
    doc.moveDown(0.5);

    if (comprobantes && comprobantes.length > 0) {
      // Cabecera tabla
      const y0 = doc.y;
      doc.rect(40, y0, 515, 20).fill(azul);
      doc.fontSize(8).fillColor("#ffffff");
      doc.text("N°", 45, y0 + 5, { width: 25 });
      doc.text("Fecha", 70, y0 + 5, { width: 65 });
      doc.text("Tipo", 135, y0 + 5, { width: 80 });
      doc.text("N° Comprobante", 215, y0 + 5, { width: 90 });
      doc.text("Concepto", 305, y0 + 5, { width: 170 });
      doc.text("Monto", 475, y0 + 5, { width: 75, align: "right" });

      doc.fillColor("#333333");
      let yRow = y0 + 22;
      comprobantes.forEach((c, i) => {
        if (yRow > 750) { doc.addPage(); yRow = 50; }
        const bg = i % 2 === 0 ? grisClaro : "#ffffff";
        doc.rect(40, yRow, 515, 18).fill(bg);
        doc.fillColor("#333333").fontSize(8);
        doc.text(String(i + 1), 45, yRow + 4, { width: 25 });
        doc.text(c.fecha || "", 70, yRow + 4, { width: 65 });
        doc.text(c.tipo || "", 135, yRow + 4, { width: 80 });
        doc.text(c.numero || "", 215, yRow + 4, { width: 90 });
        doc.text(c.concepto || "", 305, yRow + 4, { width: 170 });
        const monto = parseFloat(c.monto) || 0;
        doc.text(`S/ ${monto.toFixed(2)}`, 475, yRow + 4, { width: 75, align: "right" });
        totalComp += monto;
        yRow += 18;
      });
      doc.rect(40, yRow, 515, 20).fill(azul);
      doc.fontSize(9).fillColor("#ffffff");
      doc.text("TOTAL COMPROBANTES", 305, yRow + 5, { width: 170 });
      doc.text(`S/ ${totalComp.toFixed(2)}`, 475, yRow + 5, { width: 75, align: "right" });
      doc.y = yRow + 30;
    } else {
      doc.fontSize(9).fillColor("#666").text("No se registraron comprobantes.", 40);
    }

    doc.fillColor("#333333").moveDown(1);

    // ── Sección Declaraciones Juradas ───────────────────────────────────
    let totalDJ = 0;
    if (doc.y > 680) doc.addPage();
    doc.fontSize(14).fillColor(azul).text("2. DECLARACIONES JURADAS", 40);
    doc.moveDown(0.5);

    if (declaraciones && declaraciones.length > 0) {
      const y0 = doc.y;
      doc.rect(40, y0, 515, 20).fill(azul);
      doc.fontSize(8).fillColor("#ffffff");
      doc.text("N°", 45, y0 + 5, { width: 25 });
      doc.text("Fecha", 70, y0 + 5, { width: 65 });
      doc.text("Concepto", 135, y0 + 5, { width: 180 });
      doc.text("Motivo", 315, y0 + 5, { width: 160 });
      doc.text("Monto", 475, y0 + 5, { width: 75, align: "right" });

      doc.fillColor("#333333");
      let yRow = y0 + 22;
      declaraciones.forEach((d, i) => {
        if (yRow > 750) { doc.addPage(); yRow = 50; }
        const bg = i % 2 === 0 ? grisClaro : "#ffffff";
        doc.rect(40, yRow, 515, 18).fill(bg);
        doc.fillColor("#333333").fontSize(8);
        doc.text(String(i + 1), 45, yRow + 4, { width: 25 });
        doc.text(d.fecha || "", 70, yRow + 4, { width: 65 });
        doc.text(d.concepto || "", 135, yRow + 4, { width: 180 });
        doc.text(d.motivo || "", 315, yRow + 4, { width: 160 });
        const monto = parseFloat(d.monto) || 0;
        doc.text(`S/ ${monto.toFixed(2)}`, 475, yRow + 4, { width: 75, align: "right" });
        totalDJ += monto;
        yRow += 18;
      });
      doc.rect(40, yRow, 515, 20).fill(azul);
      doc.fontSize(9).fillColor("#ffffff");
      doc.text("TOTAL DECLARACIONES JURADAS", 315, yRow + 5, { width: 160 });
      doc.text(`S/ ${totalDJ.toFixed(2)}`, 475, yRow + 5, { width: 75, align: "right" });
      doc.y = yRow + 30;
    } else {
      doc.fontSize(9).fillColor("#666").text("No se registraron declaraciones juradas.", 40);
    }

    doc.fillColor("#333333").moveDown(1);

    // ── Sección Movilización ────────────────────────────────────────────
    let totalMov = 0;
    if (doc.y > 680) doc.addPage();
    doc.fontSize(14).fillColor(azul).text("3. MOVILIZACIÓN", 40);
    doc.moveDown(0.5);

    if (movilidad && movilidad.length > 0) {
      const y0 = doc.y;
      doc.rect(40, y0, 515, 20).fill(azul);
      doc.fontSize(8).fillColor("#ffffff");
      doc.text("N°", 45, y0 + 5, { width: 20 });
      doc.text("Fecha", 65, y0 + 5, { width: 55 });
      doc.text("Origen", 120, y0 + 5, { width: 80 });
      doc.text("Destino", 200, y0 + 5, { width: 80 });
      doc.text("Transporte", 280, y0 + 5, { width: 70 });
      doc.text("Motivo", 350, y0 + 5, { width: 125 });
      doc.text("Monto", 475, y0 + 5, { width: 75, align: "right" });

      doc.fillColor("#333333");
      let yRow = y0 + 22;
      movilidad.forEach((m, i) => {
        if (yRow > 750) { doc.addPage(); yRow = 50; }
        const bg = i % 2 === 0 ? grisClaro : "#ffffff";
        doc.rect(40, yRow, 515, 18).fill(bg);
        doc.fillColor("#333333").fontSize(8);
        doc.text(String(i + 1), 45, yRow + 4, { width: 20 });
        doc.text(m.fecha || "", 65, yRow + 4, { width: 55 });
        doc.text(m.origen || "", 120, yRow + 4, { width: 80 });
        doc.text(m.destino || "", 200, yRow + 4, { width: 80 });
        doc.text(m.transporte || "", 280, yRow + 4, { width: 70 });
        doc.text(m.motivo || "", 350, yRow + 4, { width: 125 });
        const monto = parseFloat(m.monto) || 0;
        doc.text(`S/ ${monto.toFixed(2)}`, 475, yRow + 4, { width: 75, align: "right" });
        totalMov += monto;
        yRow += 18;
      });
      doc.rect(40, yRow, 515, 20).fill(azul);
      doc.fontSize(9).fillColor("#ffffff");
      doc.text("TOTAL MOVILIZACIÓN", 350, yRow + 5, { width: 125 });
      doc.text(`S/ ${totalMov.toFixed(2)}`, 475, yRow + 5, { width: 75, align: "right" });
      doc.y = yRow + 30;
    } else {
      doc.fontSize(9).fillColor("#666").text("No se registraron gastos de movilización.", 40);
    }

    // ── Resumen Final ───────────────────────────────────────────────────
    doc.fillColor("#333333").moveDown(2);
    if (doc.y > 680) doc.addPage();

    doc.fontSize(14).fillColor(azul).text("RESUMEN GENERAL", 40);
    doc.moveDown(0.5);

    const totalGeneral = totalComp + totalDJ + totalMov;
    const yRes = doc.y;
    doc.rect(40, yRes, 515, 22).fill(grisClaro);
    doc.fontSize(10).fillColor("#333");
    doc.text("Comprobantes de Pago:", 50, yRes + 6, { width: 300 });
    doc.text(`S/ ${totalComp.toFixed(2)}`, 400, yRes + 6, { width: 150, align: "right" });

    doc.rect(40, yRes + 22, 515, 22).fill("#ffffff");
    doc.text("Declaraciones Juradas:", 50, yRes + 28, { width: 300 });
    doc.text(`S/ ${totalDJ.toFixed(2)}`, 400, yRes + 28, { width: 150, align: "right" });

    doc.rect(40, yRes + 44, 515, 22).fill(grisClaro);
    doc.text("Movilización:", 50, yRes + 50, { width: 300 });
    doc.text(`S/ ${totalMov.toFixed(2)}`, 400, yRes + 50, { width: 150, align: "right" });

    doc.rect(40, yRes + 66, 515, 26).fill(azul);
    doc.fontSize(12).fillColor("#ffffff");
    doc.text("TOTAL GENERAL:", 50, yRes + 72, { width: 300 });
    doc.text(`S/ ${totalGeneral.toFixed(2)}`, 400, yRes + 72, { width: 150, align: "right" });

    // ── Firmas ───────────────────────────────────────────────────────────
    doc.fillColor("#333333");
    let yFirma = yRes + 120;
    if (yFirma > 620) { doc.addPage(); yFirma = 60; }

    const firmaW = 220;
    const firmaX1 = 60;
    const firmaX2 = 315;

    // Firma del empleado
    if (firmaEmpleado && firmaEmpleado.imagen) {
      try {
        const imgData = firmaEmpleado.imagen.replace(/^data:image\/\w+;base64,/, "");
        const imgBuffer = Buffer.from(imgData, "base64");
        doc.image(imgBuffer, firmaX1, yFirma, { width: firmaW, height: 80, fit: [firmaW, 80] });
      } catch (e) { /* ignore */ }
    }
    doc.moveTo(firmaX1, yFirma + 85).lineTo(firmaX1 + firmaW, yFirma + 85).stroke("#999");
    doc.fontSize(9).fillColor("#333");
    doc.text("Firma del Empleado (Rindente)", firmaX1, yFirma + 90, { width: firmaW, align: "center" });
    doc.fontSize(9).fillColor("#555");
    doc.text(firmaEmpleado?.nombre || empleado || "", firmaX1, yFirma + 103, { width: firmaW, align: "center" });
    doc.text(firmaEmpleado?.cargo || "", firmaX1, yFirma + 115, { width: firmaW, align: "center" });

    // Firma del aprobador
    if (firmaAprobador && firmaAprobador.imagen) {
      try {
        const imgData = firmaAprobador.imagen.replace(/^data:image\/\w+;base64,/, "");
        const imgBuffer = Buffer.from(imgData, "base64");
        doc.image(imgBuffer, firmaX2, yFirma, { width: firmaW, height: 80, fit: [firmaW, 80] });
      } catch (e) { /* ignore */ }
    }
    doc.moveTo(firmaX2, yFirma + 85).lineTo(firmaX2 + firmaW, yFirma + 85).stroke("#999");
    doc.fontSize(9).fillColor("#333");
    doc.text("Firma del Aprobador", firmaX2, yFirma + 90, { width: firmaW, align: "center" });
    doc.fontSize(9).fillColor("#555");
    doc.text(firmaAprobador?.nombre || "", firmaX2, yFirma + 103, { width: firmaW, align: "center" });
    doc.text(firmaAprobador?.cargo || "", firmaX2, yFirma + 115, { width: firmaW, align: "center" });

    // ── Anexo: Evidencias de Comprobantes ────────────────────────────────
    const compConEvidencia = (comprobantes || []).filter(c => c.evidencia);
    if (compConEvidencia.length > 0) {
      doc.addPage();
      doc.rect(0, 0, 595.28, 50).fill(azul);
      doc.fontSize(16).fillColor("#ffffff").text("ANEXO: EVIDENCIAS DE COMPROBANTES", 40, 15, { align: "center" });

      let yEvid = 70;
      compConEvidencia.forEach((c, i) => {
        if (yEvid > 550) { doc.addPage(); yEvid = 50; }

        // Título del comprobante
        doc.fontSize(10).fillColor(azul).text(`${i + 1}. ${c.tipo} — ${c.numero} — ${c.fecha} — S/ ${parseFloat(c.monto).toFixed(2)}`, 40, yEvid);
        yEvid += 18;

        // Imagen
        try {
          const imgData = c.evidencia.replace(/^data:image\/\w+;base64,/, "");
          const imgBuffer = Buffer.from(imgData, "base64");
          const imgWidth = 400;
          const imgHeight = 280;
          doc.image(imgBuffer, 95, yEvid, { fit: [imgWidth, imgHeight], align: "center" });
          yEvid += imgHeight + 20;
        } catch (e) {
          doc.fontSize(8).fillColor("#999").text("(Imagen no disponible)", 40, yEvid);
          yEvid += 20;
        }

        // Separador
        if (i < compConEvidencia.length - 1) {
          doc.moveTo(40, yEvid).lineTo(555, yEvid).stroke("#e5e7eb");
          yEvid += 15;
        }
      });
    }

    doc.end();

  } catch (err) {
    console.error("Error generando PDF:", err);
    res.status(500).json({ ok: false, error: "Error al generar el PDF" });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// ENDPOINT: Enviar por correo
// ═══════════════════════════════════════════════════════════════════════════
app.post("/api/enviar-correo", async (req, res) => {
  try {
    const { destinatario, empleado, periodo, archivoBase64, nombreArchivo, formato } = req.body;

    if (!EMAIL_USER || !EMAIL_PASS) {
      return res.status(400).json({ ok: false, error: "Correo no configurado en el servidor. Configure EMAIL_USER y EMAIL_PASS en .env" });
    }

    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: { user: EMAIL_USER, pass: EMAIL_PASS }
    });

    const mimeType = formato === "pdf"
      ? "application/pdf"
      : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    await transporter.sendMail({
      from: `"Rendición de Viáticos" <${EMAIL_USER}>`,
      to: destinatario || EMAIL_CONTABILIDAD,
      subject: `Rendición de Viáticos — ${empleado || "Empleado"} — ${periodo || ""}`,
      html: `
        <div style="font-family:Arial,sans-serif;max-width:600px;">
          <div style="background:#1a56db;color:#fff;padding:20px;text-align:center;border-radius:8px 8px 0 0;">
            <h2 style="margin:0;">Rendición de Viáticos</h2>
            <p style="margin:5px 0 0;">${EMPRESA_NOMBRE}</p>
          </div>
          <div style="padding:20px;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 8px 8px;">
            <p>Estimado/a,</p>
            <p>Adjunto la rendición de viáticos correspondiente a:</p>
            <ul>
              <li><strong>Empleado:</strong> ${empleado || "—"}</li>
              <li><strong>Período:</strong> ${periodo || "—"}</li>
            </ul>
            <p>Por favor revise el archivo adjunto.</p>
            <p>Saludos cordiales.</p>
          </div>
        </div>
      `,
      attachments: [{
        filename: nombreArchivo,
        content: archivoBase64,
        encoding: "base64",
        contentType: mimeType
      }]
    });

    res.json({ ok: true, mensaje: "Correo enviado exitosamente" });

  } catch (err) {
    console.error("Error enviando correo:", err);
    res.status(500).json({ ok: false, error: `Error al enviar correo: ${err.message}` });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// FUNCIONES AUXILIARES
// ═══════════════════════════════════════════════════════════════════════════

function extraerDatosComprobante(texto) {
  const datos = {
    fecha: "",
    tipo: "",
    numero: "",
    concepto: "",
    monto: ""
  };

  // Buscar fecha (formatos: dd/mm/yyyy, dd-mm-yyyy, dd.mm.yyyy)
  const fechaMatch = texto.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})/);
  if (fechaMatch) {
    const [, d, m, y] = fechaMatch;
    const year = y.length === 2 ? `20${y}` : y;
    datos.fecha = `${d.padStart(2, "0")}/${m.padStart(2, "0")}/${year}`;
  }

  // Detectar tipo de comprobante
  const textoUpper = texto.toUpperCase();
  if (/FACTURA\s*(ELECTR[OÓ]NICA)?/i.test(texto)) {
    datos.tipo = "Factura Electrónica";
  } else if (/BOLETA/i.test(texto)) {
    datos.tipo = "Boleta de Venta";
  } else if (/RECIBO.*HONORARIO/i.test(texto)) {
    datos.tipo = "Recibo por Honorarios";
  } else if (/NOTA\s*DE\s*CR[EÉ]DITO/i.test(texto)) {
    datos.tipo = "Nota de Crédito";
  } else if (/TICKET/i.test(texto)) {
    datos.tipo = "Ticket";
  }

  // Buscar número de comprobante (E001-xxx, F001-xxx, B001-xxx, etc.)
  const numMatch = texto.match(/([EFBefb]\d{3})\s*[-–—]\s*(\d+)/);
  if (numMatch) {
    datos.numero = `${numMatch[1].toUpperCase()}-${numMatch[2]}`;
  } else {
    const numMatch2 = texto.match(/(\d{3,4})\s*[-–—]\s*(\d{4,10})/);
    if (numMatch2) {
      datos.numero = `${numMatch2[1]}-${numMatch2[2]}`;
    }
  }

  // Buscar monto total
  const montoPatterns = [
    /TOTAL\s*:?\s*S\/?\s*\.?\s*([\d,]+\.?\d{0,2})/i,
    /IMPORTE\s*TOTAL\s*:?\s*S\/?\s*\.?\s*([\d,]+\.?\d{0,2})/i,
    /TOTAL\s*A\s*PAGAR\s*:?\s*S\/?\s*\.?\s*([\d,]+\.?\d{0,2})/i,
    /TOTAL\s*:?\s*([\d,]+\.\d{2})/i,
    /S\/?\s*\.?\s*([\d,]+\.\d{2})/i
  ];

  for (const pattern of montoPatterns) {
    const match = texto.match(pattern);
    if (match) {
      datos.monto = match[1].replace(/,/g, "");
      break;
    }
  }

  // Buscar concepto/detalle — extraer líneas descriptivas
  const lineas = texto.split("\n").filter(l => l.trim().length > 5);
  const conceptoExclude = /fecha|ruc|factura|boleta|total|igv|subtotal|direc|telef|email|electr|serie/i;
  const conceptoLineas = lineas.filter(l => !conceptoExclude.test(l)).slice(0, 2);
  if (conceptoLineas.length > 0) {
    datos.concepto = conceptoLineas.join(" | ").substring(0, 120).trim();
  }

  return datos;
}

function agregarEncabezado(ws, empleado, periodo, titulo, extra) {
  ws.mergeCells("A1:H1");
  const titleCell = ws.getCell("A1");
  titleCell.value = titulo;
  titleCell.font = { bold: true, size: 14, color: { argb: "FF1A56DB" } };
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(1).height = 30;

  ws.getCell("A2").value = `Empleado: ${empleado || "—"}`;
  ws.getCell("A2").font = { size: 10 };
  ws.getCell("D2").value = `Centro de Costos: ${extra?.centroCostos || "—"}`;
  ws.getCell("D2").font = { size: 10 };
  ws.getCell("A3").value = `Período: ${periodo || "—"}`;
  ws.getCell("A3").font = { size: 10 };
  ws.getCell("D3").value = `N° Contrato: ${extra?.nroContrato || "—"}`;
  ws.getCell("D3").font = { size: 10 };
  ws.getCell("A4").value = `Fecha de emisión: ${new Date().toLocaleDateString("es-PE")}`;
  ws.getCell("A4").font = { size: 10, color: { argb: "FF666666" } };
}

function estiloEncabezadoTabla(row, cols) {
  row.font = { bold: true, size: 10, color: { argb: "FFFFFFFF" } };
  row.height = 22;
  for (let i = 1; i <= cols; i++) {
    const cell = row.getCell(i);
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1A56DB" } };
    cell.border = borderThin();
    cell.alignment = { horizontal: "center", vertical: "middle" };
  }
}

function borderThin() {
  return {
    top: { style: "thin", color: { argb: "FFD1D5DB" } },
    left: { style: "thin", color: { argb: "FFD1D5DB" } },
    bottom: { style: "thin", color: { argb: "FFD1D5DB" } },
    right: { style: "thin", color: { argb: "FFD1D5DB" } }
  };
}

// ─── INICIO ─────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n╔═══════════════════════════════════════════════════════════╗`);
  console.log(`║  RENDICIÓN DE VIÁTICOS — Servidor activo                 ║`);
  console.log(`║  Puerto: ${PORT}                                            ║`);
  console.log(`║  URL: http://localhost:${PORT}                              ║`);
  console.log(`╚═══════════════════════════════════════════════════════════╝\n`);
});
