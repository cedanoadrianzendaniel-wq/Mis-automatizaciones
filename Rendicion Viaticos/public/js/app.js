// ═══════════════════════════════════════════════════════════════════════════
// RENDICIÓN DE VIÁTICOS — app.js v1.0
// Lógica del cliente: escaneo OCR, gestión de datos, exportación
// ═══════════════════════════════════════════════════════════════════════════

// ─── ESTADO ─────────────────────────────────────────────────────────────────
const state = {
  comprobantes: [],
  declaraciones: [],
  movilidad: []
};

// ─── ELEMENTOS DOM ──────────────────────────────────────────────────────────
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// ─── INICIALIZACIÓN ─────────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
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
      $$(".tab-content").forEach(tc => tc.classList.remove("active"));
      tab.classList.add("active");
      $(`#tab-${tab.dataset.tab}`).classList.add("active");
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// ESCANEO OCR
// ═══════════════════════════════════════════════════════════════════════════
function initScanArea() {
  const scanArea = $("#scanArea");
  const fileInput = $("#fileInput");
  const placeholder = $("#scanPlaceholder");
  const preview = $("#scanPreview");
  const previewImg = $("#previewImg");
  const btnRemove = $("#btnRemoveImg");
  const btnEscanear = $("#btnEscanear");
  const scanProgress = $("#scanProgress");

  let archivoSeleccionado = null;

  // Click para seleccionar archivo
  scanArea.addEventListener("click", (e) => {
    if (e.target.closest("#btnRemoveImg")) return;
    fileInput.click();
  });

  // Drag & drop
  scanArea.addEventListener("dragover", (e) => {
    e.preventDefault();
    scanArea.classList.add("dragover");
  });

  scanArea.addEventListener("dragleave", () => {
    scanArea.classList.remove("dragover");
  });

  scanArea.addEventListener("drop", (e) => {
    e.preventDefault();
    scanArea.classList.remove("dragover");
    if (e.dataTransfer.files.length > 0) {
      cargarArchivo(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener("change", () => {
    if (fileInput.files.length > 0) {
      cargarArchivo(fileInput.files[0]);
    }
  });

  function cargarArchivo(file) {
    if (file.size > 10 * 1024 * 1024) {
      toast("El archivo excede el límite de 10MB", "error");
      return;
    }
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
    fileInput.value = "";
    placeholder.hidden = false;
    preview.hidden = true;
    btnEscanear.disabled = true;
  });

  // Escanear con OCR
  btnEscanear.addEventListener("click", async () => {
    if (!archivoSeleccionado) return;

    btnEscanear.disabled = true;
    scanProgress.hidden = false;

    const formData = new FormData();
    formData.append("comprobante", archivoSeleccionado);

    try {
      const resp = await fetch("/api/escanear", {
        method: "POST",
        body: formData
      });

      const data = await resp.json();

      if (data.ok) {
        // Rellenar formulario con datos extraídos
        if (data.datos.fecha) {
          // Convertir dd/mm/yyyy a yyyy-mm-dd para input date
          const partes = data.datos.fecha.split("/");
          if (partes.length === 3) {
            $("#compFecha").value = `${partes[2]}-${partes[1]}-${partes[0]}`;
          }
        }
        if (data.datos.tipo) {
          const select = $("#compTipo");
          for (const opt of select.options) {
            if (opt.value === data.datos.tipo) {
              select.value = data.datos.tipo;
              break;
            }
          }
        }
        if (data.datos.numero) {
          $("#compNumero").value = data.datos.numero;
        }
        if (data.datos.concepto) {
          $("#compConcepto").value = data.datos.concepto;
        }
        if (data.datos.monto) {
          $("#compMonto").value = data.datos.monto;
        }

        toast("Comprobante escaneado exitosamente. Verifique los datos extraídos.", "success");
      } else {
        toast(data.error || "Error al escanear", "error");
      }
    } catch (err) {
      toast("Error de conexión al escanear", "error");
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
  // ── Comprobantes ─────────────────────────────────────────────────────
  $("#btnAgregarComp").addEventListener("click", () => {
    const fecha = $("#compFecha").value;
    const tipo = $("#compTipo").value;
    const numero = $("#compNumero").value.trim();
    const concepto = $("#compConcepto").value.trim();
    const monto = $("#compMonto").value;

    if (!fecha || !tipo || !numero || !monto) {
      toast("Complete los campos obligatorios: Fecha, Tipo, N° y Monto", "warning");
      return;
    }

    state.comprobantes.push({
      fecha: formatearFecha(fecha),
      tipo,
      numero,
      concepto,
      monto: parseFloat(monto).toFixed(2)
    });

    limpiarFormComp();
    renderComprobantes();
    actualizarTotales();
    guardarDatos();
    toast("Comprobante agregado", "success");
  });

  $("#btnLimpiarComp").addEventListener("click", limpiarFormComp);

  // ── Declaraciones Juradas ───────────────────────────────────────────
  $("#btnAgregarDJ").addEventListener("click", () => {
    const fecha = $("#djFecha").value;
    const concepto = $("#djConcepto").value.trim();
    const motivo = $("#djMotivo").value.trim();
    const monto = $("#djMonto").value;

    if (!fecha || !concepto || !monto) {
      toast("Complete los campos obligatorios: Fecha, Concepto y Monto", "warning");
      return;
    }

    state.declaraciones.push({
      fecha: formatearFecha(fecha),
      concepto,
      motivo,
      monto: parseFloat(monto).toFixed(2)
    });

    $("#djFecha").value = "";
    $("#djConcepto").value = "";
    $("#djMotivo").value = "";
    $("#djMonto").value = "";

    renderDeclaraciones();
    actualizarTotales();
    guardarDatos();
    toast("Declaración jurada agregada", "success");
  });

  // ── Movilización ────────────────────────────────────────────────────
  $("#btnAgregarMov").addEventListener("click", () => {
    const fecha = $("#movFecha").value;
    const transporte = $("#movTransporte").value;
    const origen = $("#movOrigen").value.trim();
    const destino = $("#movDestino").value.trim();
    const monto = $("#movMonto").value;
    const motivo = $("#movMotivo").value.trim();

    if (!fecha || !transporte || !monto) {
      toast("Complete los campos obligatorios: Fecha, Transporte y Monto", "warning");
      return;
    }

    state.movilidad.push({
      fecha: formatearFecha(fecha),
      transporte,
      origen,
      destino,
      monto: parseFloat(monto).toFixed(2),
      motivo
    });

    $("#movFecha").value = "";
    $("#movTransporte").value = "";
    $("#movOrigen").value = "";
    $("#movDestino").value = "";
    $("#movMonto").value = "";
    $("#movMotivo").value = "";

    renderMovilidad();
    actualizarTotales();
    guardarDatos();
    toast("Movilización agregada", "success");
  });
}

function limpiarFormComp() {
  $("#compFecha").value = "";
  $("#compTipo").value = "";
  $("#compNumero").value = "";
  $("#compConcepto").value = "";
  $("#compMonto").value = "";
}

// ═══════════════════════════════════════════════════════════════════════════
// RENDER TABLAS
// ═══════════════════════════════════════════════════════════════════════════
function renderComprobantes() {
  const tbody = $("#tbodyComp");
  if (state.comprobantes.length === 0) {
    tbody.innerHTML = '<tr class="empty-row"><td colspan="7">No hay comprobantes registrados</td></tr>';
    return;
  }
  tbody.innerHTML = state.comprobantes.map((c, i) => `
    <tr>
      <td>${i + 1}</td>
      <td>${c.fecha}</td>
      <td>${c.tipo}</td>
      <td><strong>${c.numero}</strong></td>
      <td>${c.concepto}</td>
      <td><strong>S/ ${parseFloat(c.monto).toFixed(2)}</strong></td>
      <td>
        <button class="btn-icon" onclick="eliminarComprobante(${i})" title="Eliminar">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16">
            <polyline points="3 6 5 6 21 6"/>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
          </svg>
        </button>
      </td>
    </tr>
  `).join("");
}

function renderDeclaraciones() {
  const tbody = $("#tbodyDJ");
  if (state.declaraciones.length === 0) {
    tbody.innerHTML = '<tr class="empty-row"><td colspan="6">No hay declaraciones registradas</td></tr>';
    return;
  }
  tbody.innerHTML = state.declaraciones.map((d, i) => `
    <tr>
      <td>${i + 1}</td>
      <td>${d.fecha}</td>
      <td>${d.concepto}</td>
      <td>${d.motivo}</td>
      <td><strong>S/ ${parseFloat(d.monto).toFixed(2)}</strong></td>
      <td>
        <button class="btn-icon" onclick="eliminarDeclaracion(${i})" title="Eliminar">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16">
            <polyline points="3 6 5 6 21 6"/>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
          </svg>
        </button>
      </td>
    </tr>
  `).join("");
}

function renderMovilidad() {
  const tbody = $("#tbodyMov");
  if (state.movilidad.length === 0) {
    tbody.innerHTML = '<tr class="empty-row"><td colspan="8">No hay movilizaciones registradas</td></tr>';
    return;
  }
  tbody.innerHTML = state.movilidad.map((m, i) => `
    <tr>
      <td>${i + 1}</td>
      <td>${m.fecha}</td>
      <td>${m.origen}</td>
      <td>${m.destino}</td>
      <td>${m.transporte}</td>
      <td>${m.motivo}</td>
      <td><strong>S/ ${parseFloat(m.monto).toFixed(2)}</strong></td>
      <td>
        <button class="btn-icon" onclick="eliminarMovilidad(${i})" title="Eliminar">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16">
            <polyline points="3 6 5 6 21 6"/>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
          </svg>
        </button>
      </td>
    </tr>
  `).join("");
}

// ─── Eliminar registros ─────────────────────────────────────────────────────
function eliminarComprobante(i) {
  state.comprobantes.splice(i, 1);
  renderComprobantes();
  actualizarTotales();
  guardarDatos();
}

function eliminarDeclaracion(i) {
  state.declaraciones.splice(i, 1);
  renderDeclaraciones();
  actualizarTotales();
  guardarDatos();
}

function eliminarMovilidad(i) {
  state.movilidad.splice(i, 1);
  renderMovilidad();
  actualizarTotales();
  guardarDatos();
}

// ═══════════════════════════════════════════════════════════════════════════
// TOTALES
// ═══════════════════════════════════════════════════════════════════════════
function actualizarTotales() {
  const totalComp = state.comprobantes.reduce((s, c) => s + parseFloat(c.monto), 0);
  const totalDJ = state.declaraciones.reduce((s, d) => s + parseFloat(d.monto), 0);
  const totalMov = state.movilidad.reduce((s, m) => s + parseFloat(m.monto), 0);
  const totalGeneral = totalComp + totalDJ + totalMov;

  // Subtotales
  $("#subtotalComp").textContent = `Subtotal: S/ ${totalComp.toFixed(2)}`;
  $("#subtotalDJ").textContent = `Subtotal: S/ ${totalDJ.toFixed(2)}`;
  $("#subtotalMov").textContent = `Subtotal: S/ ${totalMov.toFixed(2)}`;

  // Contadores en tabs
  $("#countComp").textContent = state.comprobantes.length;
  $("#countDJ").textContent = state.declaraciones.length;
  $("#countMov").textContent = state.movilidad.length;

  // Resumen
  $("#resComp").textContent = `S/ ${totalComp.toFixed(2)}`;
  $("#resDJ").textContent = `S/ ${totalDJ.toFixed(2)}`;
  $("#resMov").textContent = `S/ ${totalMov.toFixed(2)}`;
  $("#resTotal").textContent = `S/ ${totalGeneral.toFixed(2)}`;

  // Badge header
  $("#badgeTotal").textContent = `Total: S/ ${totalGeneral.toFixed(2)}`;
}

// ═══════════════════════════════════════════════════════════════════════════
// EXPORTACIÓN
// ═══════════════════════════════════════════════════════════════════════════
function initExportacion() {
  $("#btnExportExcel").addEventListener("click", () => exportar("excel"));
  $("#btnExportPDF").addEventListener("click", () => exportar("pdf"));
}

async function exportar(formato) {
  const empleado = $("#empleado").value.trim();
  const periodo = $("#periodo").value.trim();

  if (!empleado) {
    toast("Ingrese el nombre del empleado antes de exportar", "warning");
    return;
  }

  if (state.comprobantes.length === 0 && state.declaraciones.length === 0 && state.movilidad.length === 0) {
    toast("No hay datos para exportar", "warning");
    return;
  }

  const body = {
    empleado,
    periodo,
    comprobantes: state.comprobantes,
    declaraciones: state.declaraciones,
    movilidad: state.movilidad
  };

  const endpoint = formato === "excel" ? "/api/generar-excel" : "/api/generar-pdf";
  const ext = formato === "excel" ? "xlsx" : "pdf";

  try {
    toast("Generando archivo...", "info");

    const resp = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });

    if (!resp.ok) throw new Error("Error al generar archivo");

    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Rendicion_Viaticos_${empleado.replace(/\s+/g, "_")}_${periodo || "sin_periodo"}.${ext}`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    toast(`Archivo ${ext.toUpperCase()} descargado exitosamente`, "success");
  } catch (err) {
    toast(`Error al generar ${ext.toUpperCase()}`, "error");
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// MODAL DE CORREO
// ═══════════════════════════════════════════════════════════════════════════
function initModal() {
  const modal = $("#modalCorreo");

  $("#btnEnviarCorreo").addEventListener("click", () => {
    const empleado = $("#empleado").value.trim();
    if (!empleado) {
      toast("Ingrese el nombre del empleado", "warning");
      return;
    }
    if (state.comprobantes.length === 0 && state.declaraciones.length === 0 && state.movilidad.length === 0) {
      toast("No hay datos para enviar", "warning");
      return;
    }
    modal.hidden = false;
  });

  $("#btnCerrarModal").addEventListener("click", () => { modal.hidden = true; });
  $("#btnCancelarCorreo").addEventListener("click", () => { modal.hidden = true; });

  modal.addEventListener("click", (e) => {
    if (e.target === modal) modal.hidden = true;
  });

  $("#btnConfirmarEnvio").addEventListener("click", enviarPorCorreo);
}

async function enviarPorCorreo() {
  const destinatario = $("#correoDestino").value.trim();
  const formato = document.querySelector('input[name="formatoCorreo"]:checked').value;
  const empleado = $("#empleado").value.trim();
  const periodo = $("#periodo").value.trim();

  if (!destinatario) {
    toast("Ingrese el correo del destinatario", "warning");
    return;
  }

  const btn = $("#btnConfirmarEnvio");
  btn.disabled = true;
  btn.textContent = "Enviando...";

  try {
    // Generar el archivo primero
    const body = {
      empleado,
      periodo,
      comprobantes: state.comprobantes,
      declaraciones: state.declaraciones,
      movilidad: state.movilidad
    };

    const endpoint = formato === "excel" ? "/api/generar-excel" : "/api/generar-pdf";
    const resp = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });

    if (!resp.ok) throw new Error("Error al generar archivo");

    const blob = await resp.blob();
    const reader = new FileReader();

    reader.onload = async () => {
      const base64 = reader.result.split(",")[1];
      const ext = formato === "excel" ? "xlsx" : "pdf";
      const nombreArchivo = `Rendicion_Viaticos_${empleado.replace(/\s+/g, "_")}_${periodo || ""}.${ext}`;

      try {
        const envioResp = await fetch("/api/enviar-correo", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            destinatario,
            empleado,
            periodo,
            archivoBase64: base64,
            nombreArchivo,
            formato
          })
        });

        const data = await envioResp.json();

        if (data.ok) {
          toast("Correo enviado exitosamente", "success");
          $("#modalCorreo").hidden = true;
        } else {
          toast(data.error || "Error al enviar correo", "error");
        }
      } catch (err) {
        toast("Error de conexión al enviar correo", "error");
      } finally {
        btn.disabled = false;
        btn.innerHTML = `
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16">
            <line x1="22" y1="2" x2="11" y2="13"/>
            <polygon points="22 2 15 22 11 13 2 9 22 2"/>
          </svg>
          Enviar
        `;
      }
    };

    reader.readAsDataURL(blob);
  } catch (err) {
    toast("Error al preparar el archivo", "error");
    btn.disabled = false;
    btn.innerHTML = `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16">
        <line x1="22" y1="2" x2="11" y2="13"/>
        <polygon points="22 2 15 22 11 13 2 9 22 2"/>
      </svg>
      Enviar
    `;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// PERSISTENCIA LOCAL
// ═══════════════════════════════════════════════════════════════════════════
function guardarDatos() {
  const datos = {
    empleado: $("#empleado").value,
    cargo: $("#cargo").value,
    area: $("#area").value,
    periodo: $("#periodo").value,
    comprobantes: state.comprobantes,
    declaraciones: state.declaraciones,
    movilidad: state.movilidad
  };
  localStorage.setItem("rendicion_viaticos", JSON.stringify(datos));
}

function cargarDatosGuardados() {
  const saved = localStorage.getItem("rendicion_viaticos");
  if (!saved) return;

  try {
    const datos = JSON.parse(saved);
    if (datos.empleado) $("#empleado").value = datos.empleado;
    if (datos.cargo) $("#cargo").value = datos.cargo;
    if (datos.area) $("#area").value = datos.area;
    if (datos.periodo) $("#periodo").value = datos.periodo;
    if (datos.comprobantes) state.comprobantes = datos.comprobantes;
    if (datos.declaraciones) state.declaraciones = datos.declaraciones;
    if (datos.movilidad) state.movilidad = datos.movilidad;

    renderComprobantes();
    renderDeclaraciones();
    renderMovilidad();
    actualizarTotales();
  } catch (e) {
    // datos corruptos, ignorar
  }

  // Guardar al cambiar datos del empleado
  ["empleado", "cargo", "area", "periodo"].forEach(id => {
    $(`#${id}`).addEventListener("input", guardarDatos);
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// UTILIDADES
// ═══════════════════════════════════════════════════════════════════════════
function formatearFecha(dateStr) {
  if (!dateStr) return "";
  const [y, m, d] = dateStr.split("-");
  return `${d}/${m}/${y}`;
}

function toast(mensaje, tipo = "info") {
  const container = $("#toastContainer");
  const div = document.createElement("div");
  div.className = `toast toast-${tipo}`;
  div.textContent = mensaje;
  container.appendChild(div);
  setTimeout(() => {
    div.style.opacity = "0";
    div.style.transform = "translateX(100%)";
    div.style.transition = "all 0.3s ease";
    setTimeout(() => div.remove(), 300);
  }, 4000);
}
