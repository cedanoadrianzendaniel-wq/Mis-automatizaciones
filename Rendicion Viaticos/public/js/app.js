// ═══════════════════════════════════════════════════════════════════════════
// RENDIGASTOS — app.js v2.0
// ═══════════════════════════════════════════════════════════════════════════

const state = { comprobantes: [], declaraciones: [], movilidad: [] };
const $ = (s) => document.querySelector(s);
const $$ = (s) => document.querySelectorAll(s);

// ─── INIT ───────────────────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
  const modal = document.getElementById("modalCorreo");
  if (modal) modal.style.display = "none";

  initTabs();
  initScanArea();
  initFormularios();
  initExportacion();
  initModal();
  cargarDatosGuardados();
});

// ═══════════════════════════════════════════════════════════════════════════
// TABS
// ═══════════════════════════════════════════════════════════════════════════
function initTabs() {
  $$(".tab").forEach(tab => {
    tab.addEventListener("click", () => {
      $$(".tab").forEach(t => t.classList.remove("active"));
      $$(".tab-panel").forEach(p => p.classList.remove("active"));
      tab.classList.add("active");
      $(`#tab-${tab.dataset.tab}`).classList.add("active");
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SCAN / CAPTURE
// ═══════════════════════════════════════════════════════════════════════════
function initScanArea() {
  const scanArea = $("#scanArea");
  const fileCamara = $("#fileCamara");
  const fileGaleria = $("#fileGaleria");
  const placeholder = $("#scanPlaceholder");
  const preview = $("#scanPreview");
  const previewImg = $("#previewImg");
  const btnRemove = $("#btnRemoveImg");
  const btnEscanear = $("#btnEscanear");
  const scanProgress = $("#scanProgress");

  let archivoSeleccionado = null;

  $("#btnTomarFoto").addEventListener("click", () => fileCamara.click());
  $("#btnSubirArchivo").addEventListener("click", () => fileGaleria.click());

  scanArea.addEventListener("click", (e) => {
    if (e.target.closest("#btnRemoveImg") || e.target.closest(".capture-img")) return;
    fileGaleria.click();
  });

  scanArea.addEventListener("dragover", (e) => { e.preventDefault(); scanArea.classList.add("dragover"); });
  scanArea.addEventListener("dragleave", () => scanArea.classList.remove("dragover"));
  scanArea.addEventListener("drop", (e) => {
    e.preventDefault(); scanArea.classList.remove("dragover");
    if (e.dataTransfer.files.length > 0) cargarArchivo(e.dataTransfer.files[0]);
  });

  fileCamara.addEventListener("change", () => { if (fileCamara.files.length > 0) cargarArchivo(fileCamara.files[0]); });
  fileGaleria.addEventListener("change", () => { if (fileGaleria.files.length > 0) cargarArchivo(fileGaleria.files[0]); });

  function cargarArchivo(file) {
    if (file.size > 10 * 1024 * 1024) { toast("Archivo excede 10MB", "error"); return; }
    archivoSeleccionado = file;
    const reader = new FileReader();
    reader.onload = (e) => {
      previewImg.src = e.target.result;
      placeholder.hidden = true;
      preview.hidden = false;
      btnEscanear.disabled = false;
    };
    reader.readAsDataURL(file);
  }

  btnRemove.addEventListener("click", (e) => {
    e.stopPropagation();
    archivoSeleccionado = null;
    fileCamara.value = ""; fileGaleria.value = "";
    placeholder.hidden = false; preview.hidden = true;
    btnEscanear.disabled = true;
  });

  btnEscanear.addEventListener("click", async () => {
    if (!archivoSeleccionado) return;
    btnEscanear.disabled = true;
    scanProgress.hidden = false;

    const formData = new FormData();
    formData.append("comprobante", archivoSeleccionado);

    try {
      const resp = await fetch("/api/escanear", { method: "POST", body: formData });
      const data = await resp.json();

      if (data.ok) {
        if (data.datos.fecha) {
          const p = data.datos.fecha.split("/");
          if (p.length === 3) $("#compFecha").value = `${p[2]}-${p[1]}-${p[0]}`;
        }
        if (data.datos.tipo) {
          const sel = $("#compTipo");
          for (const opt of sel.options) { if (opt.value === data.datos.tipo) { sel.value = data.datos.tipo; break; } }
        }
        if (data.datos.numero) $("#compNumero").value = data.datos.numero;
        if (data.datos.concepto) $("#compConcepto").value = data.datos.concepto;
        if (data.datos.monto) $("#compMonto").value = data.datos.monto;

        toast("Comprobante escaneado. Verifique los datos.", "success");
      } else {
        toast(data.error || "Error al escanear", "error");
      }
    } catch (err) {
      toast("Error de conexi\u00f3n al escanear", "error");
    } finally {
      btnEscanear.disabled = false;
      scanProgress.hidden = true;
    }
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// FORMULARIOS
// ═══════════════════════════════════════════════════════════════════════════
function initFormularios() {
  // Comprobantes
  $("#btnAgregarComp").addEventListener("click", () => {
    const fecha = $("#compFecha").value, tipo = $("#compTipo").value;
    const numero = $("#compNumero").value.trim(), concepto = $("#compConcepto").value.trim();
    const monto = $("#compMonto").value;

    if (!fecha || !tipo || !numero || !monto) { toast("Complete: Fecha, Tipo, N\u00b0 y Monto", "warning"); return; }

    state.comprobantes.push({ fecha: fmtFecha(fecha), tipo, numero, concepto, monto: parseFloat(monto).toFixed(2) });
    limpiarFormComp(); renderComprobantes(); actualizarTotales(); guardarDatos();
    toast("Comprobante agregado", "success");
  });

  $("#btnLimpiarComp").addEventListener("click", limpiarFormComp);

  // Declaraciones
  $("#btnAgregarDJ").addEventListener("click", () => {
    const fecha = $("#djFecha").value, concepto = $("#djConcepto").value.trim();
    const motivo = $("#djMotivo").value.trim(), monto = $("#djMonto").value;

    if (!fecha || !concepto || !monto) { toast("Complete: Fecha, Concepto y Monto", "warning"); return; }

    state.declaraciones.push({ fecha: fmtFecha(fecha), concepto, motivo, monto: parseFloat(monto).toFixed(2) });
    $("#djFecha").value = ""; $("#djConcepto").value = ""; $("#djMotivo").value = ""; $("#djMonto").value = "";
    renderDeclaraciones(); actualizarTotales(); guardarDatos();
    toast("Declaraci\u00f3n jurada agregada", "success");
  });

  // Movilidad
  $("#btnAgregarMov").addEventListener("click", () => {
    const fecha = $("#movFecha").value, transporte = $("#movTransporte").value;
    const origen = $("#movOrigen").value.trim(), destino = $("#movDestino").value.trim();
    const monto = $("#movMonto").value, motivo = $("#movMotivo").value.trim();

    if (!fecha || !transporte || !monto) { toast("Complete: Fecha, Transporte y Monto", "warning"); return; }

    state.movilidad.push({ fecha: fmtFecha(fecha), transporte, origen, destino, monto: parseFloat(monto).toFixed(2), motivo });
    ["movFecha","movTransporte","movOrigen","movDestino","movMonto","movMotivo"].forEach(id => $(`#${id}`).value = "");
    renderMovilidad(); actualizarTotales(); guardarDatos();
    toast("Movilizaci\u00f3n agregada", "success");
  });
}

function limpiarFormComp() {
  ["compFecha","compTipo","compNumero","compConcepto","compMonto"].forEach(id => $(`#${id}`).value = "");
}

// ═══════════════════════════════════════════════════════════════════════════
// RENDER
// ═══════════════════════════════════════════════════════════════════════════
const trashSVG = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="15" height="15"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>';

function renderComprobantes() {
  const tb = $("#tbodyComp");
  if (!state.comprobantes.length) { tb.innerHTML = '<tr class="row-empty"><td colspan="7">Sin comprobantes registrados</td></tr>'; return; }
  tb.innerHTML = state.comprobantes.map((c, i) => `<tr>
    <td>${i+1}</td><td>${c.fecha}</td><td>${c.tipo}</td><td><strong>${c.numero}</strong></td><td>${c.concepto}</td>
    <td class="text-right"><strong>S/ ${parseFloat(c.monto).toFixed(2)}</strong></td>
    <td><button class="btn-icon" onclick="eliminarComprobante(${i})" title="Eliminar">${trashSVG}</button></td></tr>`).join("");
}

function renderDeclaraciones() {
  const tb = $("#tbodyDJ");
  if (!state.declaraciones.length) { tb.innerHTML = '<tr class="row-empty"><td colspan="6">Sin declaraciones registradas</td></tr>'; return; }
  tb.innerHTML = state.declaraciones.map((d, i) => `<tr>
    <td>${i+1}</td><td>${d.fecha}</td><td>${d.concepto}</td><td>${d.motivo}</td>
    <td class="text-right"><strong>S/ ${parseFloat(d.monto).toFixed(2)}</strong></td>
    <td><button class="btn-icon" onclick="eliminarDeclaracion(${i})" title="Eliminar">${trashSVG}</button></td></tr>`).join("");
}

function renderMovilidad() {
  const tb = $("#tbodyMov");
  if (!state.movilidad.length) { tb.innerHTML = '<tr class="row-empty"><td colspan="8">Sin movilizaciones registradas</td></tr>'; return; }
  tb.innerHTML = state.movilidad.map((m, i) => `<tr>
    <td>${i+1}</td><td>${m.fecha}</td><td>${m.origen}</td><td>${m.destino}</td><td>${m.transporte}</td><td>${m.motivo}</td>
    <td class="text-right"><strong>S/ ${parseFloat(m.monto).toFixed(2)}</strong></td>
    <td><button class="btn-icon" onclick="eliminarMovilidad(${i})" title="Eliminar">${trashSVG}</button></td></tr>`).join("");
}

function eliminarComprobante(i) { state.comprobantes.splice(i,1); renderComprobantes(); actualizarTotales(); guardarDatos(); }
function eliminarDeclaracion(i) { state.declaraciones.splice(i,1); renderDeclaraciones(); actualizarTotales(); guardarDatos(); }
function eliminarMovilidad(i) { state.movilidad.splice(i,1); renderMovilidad(); actualizarTotales(); guardarDatos(); }

// ═══════════════════════════════════════════════════════════════════════════
// TOTALES
// ═══════════════════════════════════════════════════════════════════════════
function actualizarTotales() {
  const totalComp = state.comprobantes.reduce((s,c) => s + parseFloat(c.monto), 0);
  const totalDJ = state.declaraciones.reduce((s,d) => s + parseFloat(d.monto), 0);
  const totalMov = state.movilidad.reduce((s,m) => s + parseFloat(m.monto), 0);
  const total = totalComp + totalDJ + totalMov;

  $("#subtotalComp").textContent = `Subtotal: S/ ${totalComp.toFixed(2)}`;
  $("#subtotalDJ").textContent = `Subtotal: S/ ${totalDJ.toFixed(2)}`;
  $("#subtotalMov").textContent = `Subtotal: S/ ${totalMov.toFixed(2)}`;

  $("#countComp").textContent = state.comprobantes.length;
  $("#countDJ").textContent = state.declaraciones.length;
  $("#countMov").textContent = state.movilidad.length;

  $("#resComp").textContent = `S/ ${totalComp.toFixed(2)}`;
  $("#resDJ").textContent = `S/ ${totalDJ.toFixed(2)}`;
  $("#resMov").textContent = `S/ ${totalMov.toFixed(2)}`;
  $("#resTotal").textContent = `S/ ${total.toFixed(2)}`;

  // Topbar pill
  const pill = $("#badgeTotal");
  pill.querySelector(".pill-value").textContent = `S/ ${total.toFixed(2)}`;

  actualizarSaldo(total);
}

function actualizarSaldo(totalGastado) {
  const asignado = parseFloat($("#viaticoAsignado").value) || 0;
  const saldo = asignado - totalGastado;

  $("#saldoAsignado").textContent = `S/ ${asignado.toFixed(2)}`;
  $("#saldoGastado").textContent = `S/ ${totalGastado.toFixed(2)}`;
  $("#saldoRestante").textContent = `S/ ${Math.abs(saldo).toFixed(2)}`;

  const el = $("#balanceResult");
  const dot = $("#balanceDot");
  const det = $("#saldoDetalle");

  el.classList.remove("positivo","negativo","neutro");
  dot.classList.remove("dot--green","dot--red","dot--amber");

  if (asignado === 0) {
    el.classList.add("neutro"); dot.classList.add("dot--amber");
    det.textContent = "Ingrese vi\u00e1tico asignado";
  } else if (saldo > 0) {
    el.classList.add("positivo"); dot.classList.add("dot--green");
    det.textContent = "Saldo a favor (devolver)";
  } else if (saldo < 0) {
    el.classList.add("negativo"); dot.classList.add("dot--red");
    det.textContent = "Excedido (por reembolsar)";
  } else {
    el.classList.add("positivo"); dot.classList.add("dot--green");
    det.textContent = "Rendici\u00f3n exacta";
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// EXPORT
// ═══════════════════════════════════════════════════════════════════════════
function initExportacion() {
  $("#btnExportExcel").addEventListener("click", () => exportar("excel"));
  $("#btnExportPDF").addEventListener("click", () => exportar("pdf"));
}

async function exportar(formato) {
  const empleado = $("#empleado").value.trim();
  const periodo = $("#periodo").value.trim();
  if (!empleado) { toast("Ingrese nombre del empleado", "warning"); return; }
  if (!state.comprobantes.length && !state.declaraciones.length && !state.movilidad.length) { toast("No hay datos", "warning"); return; }

  const viaticoAsignado = $("#viaticoAsignado").value || "0";
  const body = { empleado, periodo, viaticoAsignado, comprobantes: state.comprobantes, declaraciones: state.declaraciones, movilidad: state.movilidad };
  const endpoint = formato === "excel" ? "/api/generar-excel" : "/api/generar-pdf";
  const ext = formato === "excel" ? "xlsx" : "pdf";

  try {
    toast("Generando archivo\u2026", "info");
    const resp = await fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    if (!resp.ok) throw new Error();
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `Rendicion_${empleado.replace(/\s+/g,"_")}_${periodo||"sin_periodo"}.${ext}`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    toast(`${ext.toUpperCase()} descargado`, "success");
  } catch { toast(`Error al generar ${ext.toUpperCase()}`, "error"); }
}

// ═══════════════════════════════════════════════════════════════════════════
// MODAL
// ═══════════════════════════════════════════════════════════════════════════
function initModal() {
  const modal = $("#modalCorreo");

  $("#btnEnviarCorreo").addEventListener("click", () => {
    if (!$("#empleado").value.trim()) { toast("Ingrese nombre del empleado", "warning"); return; }
    if (!state.comprobantes.length && !state.declaraciones.length && !state.movilidad.length) { toast("No hay datos", "warning"); return; }
    modal.style.display = "flex";
  });

  function cerrarModal() { modal.style.display = "none"; }
  $("#btnCerrarModal").addEventListener("click", cerrarModal);
  $("#btnCancelarCorreo").addEventListener("click", cerrarModal);
  modal.addEventListener("click", (e) => { if (e.target === modal) cerrarModal(); });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape" && modal.style.display !== "none") cerrarModal(); });

  $("#btnConfirmarEnvio").addEventListener("click", enviarPorCorreo);
}

async function enviarPorCorreo() {
  const destinatario = $("#correoDestino").value.trim();
  const formato = document.querySelector('input[name="formatoCorreo"]:checked').value;
  const empleado = $("#empleado").value.trim();
  const periodo = $("#periodo").value.trim();
  if (!destinatario) { toast("Ingrese correo del destinatario", "warning"); return; }

  const btn = $("#btnConfirmarEnvio");
  btn.disabled = true; btn.textContent = "Enviando\u2026";

  try {
    const viaticoAsignado = $("#viaticoAsignado").value || "0";
    const body = { empleado, periodo, viaticoAsignado, comprobantes: state.comprobantes, declaraciones: state.declaraciones, movilidad: state.movilidad };
    const endpoint = formato === "excel" ? "/api/generar-excel" : "/api/generar-pdf";
    const resp = await fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    if (!resp.ok) throw new Error();

    const blob = await resp.blob();
    const reader = new FileReader();
    reader.onload = async () => {
      const base64 = reader.result.split(",")[1];
      const ext = formato === "excel" ? "xlsx" : "pdf";
      try {
        const r = await fetch("/api/enviar-correo", {
          method: "POST", headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ destinatario, empleado, periodo, archivoBase64: base64, nombreArchivo: `Rendicion_${empleado.replace(/\s+/g,"_")}.${ext}`, formato })
        });
        const data = await r.json();
        $("#modalCorreo").style.display = "none";
        toast(data.ok ? "Correo enviado" : (data.error || "Error al enviar"), data.ok ? "success" : "error");
      } catch { toast("Error de conexi\u00f3n", "error"); }
      finally { resetBtn(); }
    };
    reader.readAsDataURL(blob);
  } catch { toast("Error al preparar archivo", "error"); resetBtn(); }

  function resetBtn() {
    btn.disabled = false;
    btn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg> Enviar';
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// PERSISTENCE
// ═══════════════════════════════════════════════════════════════════════════
function guardarDatos() {
  localStorage.setItem("rendicion_viaticos", JSON.stringify({
    empleado: $("#empleado").value, cargo: $("#cargo").value,
    area: $("#area").value, periodo: $("#periodo").value,
    viaticoAsignado: $("#viaticoAsignado").value,
    comprobantes: state.comprobantes, declaraciones: state.declaraciones, movilidad: state.movilidad
  }));
}

function cargarDatosGuardados() {
  const saved = localStorage.getItem("rendicion_viaticos");
  if (saved) {
    try {
      const d = JSON.parse(saved);
      if (d.empleado) $("#empleado").value = d.empleado;
      if (d.cargo) $("#cargo").value = d.cargo;
      if (d.area) $("#area").value = d.area;
      if (d.periodo) $("#periodo").value = d.periodo;
      if (d.viaticoAsignado) $("#viaticoAsignado").value = d.viaticoAsignado;
      if (d.comprobantes) state.comprobantes = d.comprobantes;
      if (d.declaraciones) state.declaraciones = d.declaraciones;
      if (d.movilidad) state.movilidad = d.movilidad;
      renderComprobantes(); renderDeclaraciones(); renderMovilidad(); actualizarTotales();
    } catch {}
  }

  ["empleado","cargo","area","periodo"].forEach(id => $(`#${id}`).addEventListener("input", guardarDatos));
  $("#viaticoAsignado").addEventListener("input", () => { guardarDatos(); actualizarTotales(); });
}

// ═══════════════════════════════════════════════════════════════════════════
// UTILS
// ═══════════════════════════════════════════════════════════════════════════
function fmtFecha(d) { if (!d) return ""; const [y,m,dd] = d.split("-"); return `${dd}/${m}/${y}`; }

function toast(msg, type = "info") {
  const c = $("#toastContainer");
  const el = document.createElement("div");
  el.className = `toast toast-${type}`;
  el.textContent = msg;
  c.appendChild(el);
  setTimeout(() => {
    el.style.opacity = "0"; el.style.transform = "translateX(100%)";
    el.style.transition = "all .3s ease";
    setTimeout(() => el.remove(), 300);
  }, 3500);
}
