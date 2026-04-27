/**
 * Genera un reporte consolidado de gestiones (Activas + Históricas) para pago de incentivos.
 * Escanea tanto el archivo histórico como las planillas vivas de los agentes.
 */
function GenerarReporteIncentivos() {

  // --- 1. CONFIGURACIÓN LOCAL (IDs) ---
  const IDs = {
    supervisora:       "1gMQJkimpLeiHafcWTumJLh4lQuMlXjk9PN119KIWnUI",
    historicoAgentes:  "1zngpyCyipGSmM93Yl98vUy_W_GIkDT-KS0L-qLexP84",
    carpetaAgentes:    "1kSrb-zocfdEVFxHZGsjcefZMTido0X4N"
  };

  const ui = SpreadsheetApp.getUi();

  // --- 2. SOLICITAR RANGO DE FECHAS ---
  const input = ui.prompt(
    'Generar Reporte de Incentivos',
    'Ingresa el rango de fechas (Formato: AAAA-MM-DD,AAAA-MM-DD)\nEjemplo: 2026-01-01,2026-01-31',
    ui.ButtonSet.OK_CANCEL
  );

  if (input.getSelectedButton() !== ui.Button.OK) return;

  const textosFecha = input.getResponseText().split(',');
  if (textosFecha.length !== 2) {
    ui.alert('Formato incorrecto. Debe ser: Inicio,Fin');
    return;
  }

  const fInicio = new Date(textosFecha[0].trim() + "T00:00:00");
  const fFin    = new Date(textosFecha[1].trim() + "T23:59:59");

  if (isNaN(fInicio.getTime()) || isNaN(fFin.getTime())) {
    ui.alert('Fechas inválidas.');
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Iniciando escaneo masivo (Esto puede tardar unos minutos)...',
    'Reporte Incentivos', -1
  );

  const ssSup = SpreadsheetApp.openById(IDs.supervisora);
  let reporteFinal = [];

  // Encabezados
  reporteFinal.push([
    "Fecha Gestión", "Agente", "ID Caso", "Origen",
    "Tipo Planilla", "Estado Contacto", "Estado Adherencia"
  ]);

  // ==========================================
  // FASE 1: ESCANEAR HISTÓRICO (Lo ya cerrado)
  // Estructura hoja "Gestiones":
  //   A(0)=ID_Caso, L(11)=Estado Fono/Mail, N(13)=Estado Adherencia,
  //   X(23)=Fecha Gestión, Z(25)=Tipo Gestión, AA(26)=Agente
  // ==========================================
  try {
    const ssHist  = SpreadsheetApp.openById(IDs.historicoAgentes);
    const hojaHist = ssHist.getSheetByName("Gestiones");

    if (hojaHist) {
      const datosHist = hojaHist.getDataRange().getValues();

      for (let i = 1; i < datosHist.length; i++) {
        const fila   = datosHist[i];
        const fechaG = fila[23]; // Col X — Fecha de la Gestión

        // Validar que sea fecha real y esté en rango
        if (!(fechaG instanceof Date) || isNaN(fechaG.getTime())) continue;
        if (fechaG < fInicio || fechaG > fFin) continue;

        const estadoContacto   = (fila[11] || "").toString().trim(); // Col L
        const estadoAdherencia = (fila[13] || "").toString().trim(); // Col N

        reporteFinal.push([
          fechaG,
          fila[26],              // Col AA — Agente
          fila[0],               // Col A  — ID Caso
          "HISTÓRICO",
          fila[25] || "General", // Col Z  — Tipo Gestión
          estadoContacto,
          estadoAdherencia
        ]);
      }
    }
  } catch (e) {
    console.warn("Error leyendo histórico: " + e.message);
  }

  // ==========================================
  // FASE 2: ESCANEAR PLANILLAS ACTIVAS
  // Estructura hojas "Agendar" y "Notificar":
  //   A(0)=ID Caso, X(23)=Fecha Gestión
  //   Agendar:   L(11)=Estado Contacto, N(13)=Estado Adherencia
  //   Notificar: L(11)=Estado Contacto, N(13)=Estado Adherencia
  // ==========================================
  try {
    const carpeta = DriveApp.getFolderById(IDs.carpetaAgentes);
    const archivos = carpeta.getFiles();

    while (archivos.hasNext()) {
      const archivo = archivos.next();
      if (archivo.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

      const nombreAgente = archivo.getName().replace("[BACKUP] ", "").trim();

      try {
        const ssAgente      = SpreadsheetApp.open(archivo);
        const hojasAProcesar = ["Agendar", "Notificar"];

        hojasAProcesar.forEach(nombreHoja => {
        const hoja = ssAgente.getSheetByName(nombreHoja);
        if (!hoja || hoja.getLastRow() <= 1) return;

        const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 24).getValues();

        // Índices según pestaña
        const isAgendar      = (nombreHoja === "Agendar");
        const idxContacto    = isAgendar ? 11 : 13; // Col L vs Col N
        const idxAdherencia  = isAgendar ? 13 : 15; // Col N vs Col P
        const idxFecha       = 23;                  // Col X — igual en ambas

        datos.forEach(fila => {
          const fechaG = fila[idxFecha];

          if (!(fechaG instanceof Date) || isNaN(fechaG.getTime())) return;
          if (fechaG < fInicio || fechaG > fFin) return;

          reporteFinal.push([
            fechaG,
            nombreAgente,
            fila[0],
            "ACTIVO (EN CURSO)",
            nombreHoja,
            (fila[idxContacto]   || "").toString().trim(),
            (fila[idxAdherencia] || "").toString().trim()
          ]);
        });
      });
      } catch (e) {
        console.warn(`Error leyendo planilla de ${nombreAgente}: ${e.message}`);
      }
    }
  } catch (e) {
    console.warn("Error accediendo a carpeta de agentes: " + e.message);
  }

  // ==========================================
  // FASE 3: ESCRIBIR REPORTE EN SUPERVISORA
  // ==========================================
  let hojaReporte = ssSup.getSheetByName("Reporte_Consolidado_Gestiones");
  if (!hojaReporte) {
    hojaReporte = ssSup.insertSheet("Reporte_Consolidado_Gestiones");
  } else {
    hojaReporte.clear();
  }

  if (reporteFinal.length > 1) {
    hojaReporte.getRange(1, 1, reporteFinal.length, 7).setValues(reporteFinal);
    hojaReporte.getRange(2, 1, reporteFinal.length - 1, 1).setNumberFormat("dd/mm/yyyy hh:mm");
    hojaReporte.autoResizeColumns(1, 7);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Reporte generado con ${reporteFinal.length - 1} gestiones.`,
    'Finalizado', -1
  );

  ui.alert(
    `Proceso completado.\n\n` +
    `Se encontraron ${reporteFinal.length - 1} gestiones en el periodo seleccionado.\n` +
    `Revisa la hoja 'Reporte_Consolidado_Gestiones' en la Planilla Supervisora.`
  );
}


// ==========================================
// TRIGGER DIARIO — Ejecutar a medianoche
// ==========================================
function configurarTriggerDiario() {
  // Eliminar triggers existentes para evitar duplicados
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "GenerarReporteIncentivosAuto") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("GenerarReporteIncentivosAuto")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();

  SpreadsheetApp.getUi().alert("Trigger diario configurado. El reporte se actualizará cada día a medianoche.");
}

/**
 * Versión automática para el trigger — usa siempre los últimos 30 días
 * como ventana de datos frescos.
 */
function GenerarReporteIncentivosAuto() {
  const hoy     = new Date();
  const fFin    = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate(), 23, 59, 59);
  const fInicio = new Date(hoy.getFullYear(), hoy.getMonth(), 1, 0, 0, 0); // Primer día del mes actual

  _ejecutarReporte(fInicio, fFin);
}

/**
 * Versión manual con prompt de fechas.
 */
function GenerarReporteIncentivos() {
  const ui    = SpreadsheetApp.getUi();
  const input = ui.prompt(
    'Generar Reporte de Incentivos',
    'Ingresa el rango de fechas (Formato: AAAA-MM-DD,AAAA-MM-DD)\nEjemplo: 2026-01-01,2026-01-31',
    ui.ButtonSet.OK_CANCEL
  );

  if (input.getSelectedButton() !== ui.Button.OK) return;

  const textosFecha = input.getResponseText().split(',');
  if (textosFecha.length !== 2) {
    ui.alert('Formato incorrecto. Debe ser: Inicio,Fin');
    return;
  }

  const fInicio = new Date(textosFecha[0].trim() + "T00:00:00");
  const fFin    = new Date(textosFecha[1].trim() + "T23:59:59");

  if (isNaN(fInicio.getTime()) || isNaN(fFin.getTime())) {
    ui.alert('Fechas inválidas.');
    return;
  }

  _ejecutarReporte(fInicio, fFin);
}

/**
 * Lógica central compartida por ambas versiones.
 */
function _ejecutarReporte(fInicio, fFin) {

  const IDs = {
    supervisora:      "1gMQJkimpLeiHafcWTumJLh4lQuMlXjk9PN119KIWnUI",
    historicoAgentes: "1zngpyCyipGSmM93Yl98vUy_W_GIkDT-KS0L-qLexP84",
    carpetaAgentes:   "1kSrb-zocfdEVFxHZGsjcefZMTido0X4N"
  };

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Iniciando escaneo masivo (Esto puede tardar unos minutos)...',
    'Reporte Incentivos', -1
  );

  const ssSup = SpreadsheetApp.openById(IDs.supervisora);
  let reporteFinal = [];

  reporteFinal.push([
    "Fecha Gestión", "Agente", "ID Caso", "Origen",
    "Tipo Planilla", "Estado Contacto", "Estado Adherencia"
  ]);

  // --- FASE 1: HISTÓRICO ---
  try {
    const ssHist   = SpreadsheetApp.openById(IDs.historicoAgentes);
    const hojaHist = ssHist.getSheetByName("Gestiones");

    if (hojaHist) {
      const datosHist = hojaHist.getDataRange().getValues();

      for (let i = 1; i < datosHist.length; i++) {
        const fila   = datosHist[i];
        const fechaG = fila[23];

        if (!(fechaG instanceof Date) || isNaN(fechaG.getTime())) continue;
        if (fechaG < fInicio || fechaG > fFin) continue;

        reporteFinal.push([
          fechaG,
          fila[26],
          fila[0],
          "HISTÓRICO",
          fila[25] || "General",
          (fila[11] || "").toString().trim(),
          (fila[13] || "").toString().trim()
        ]);
      }
    }
  } catch (e) {
    console.warn("Error leyendo histórico: " + e.message);
  }

  // --- FASE 2: PLANILLAS ACTIVAS ---
  try {
    const carpeta  = DriveApp.getFolderById(IDs.carpetaAgentes);
    const archivos = carpeta.getFiles();

    while (archivos.hasNext()) {
      const archivo = archivos.next();
      if (archivo.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

      const nombreAgente = archivo.getName().replace("[BACKUP] ", "").trim();

      try {
        const ssAgente       = SpreadsheetApp.open(archivo);
        const hojasAProcesar = ["Agendar", "Notificar"];

        hojasAProcesar.forEach(nombreHoja => {
          const hoja = ssAgente.getSheetByName(nombreHoja);
          if (!hoja || hoja.getLastRow() <= 1) return;

          const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 24).getValues();

          // ---> SOLUCIÓN: ÍNDICES DINÁMICOS SEGÚN LA PESTAÑA <---
          const isAgendar     = (nombreHoja === "Agendar");
          const idxContacto   = isAgendar ? 11 : 13; // Col L (11) para Agendar, Col N (13) para Notificar
          const idxAdherencia = isAgendar ? 13 : 15; // Col N (13) para Agendar, Col P (15) para Notificar
          const idxFecha      = 23;                  // Col X (23) igual en ambas

          datos.forEach(fila => {
            const fechaG = fila[idxFecha];

            if (!(fechaG instanceof Date) || isNaN(fechaG.getTime())) return;
            if (fechaG < fInicio || fechaG > fFin) return;

            reporteFinal.push([
              fechaG,
              nombreAgente,
              fila[0],
              "ACTIVO (EN CURSO)",
              nombreHoja,
              (fila[idxContacto] || "").toString().trim(),   // Ahora lee la columna correcta
              (fila[idxAdherencia] || "").toString().trim()  // Ahora lee la columna correcta
            ]);
          });
        });
      } catch (e) {
        console.warn(`Error leyendo planilla de ${nombreAgente}: ${e.message}`);
      }
    }
  } catch (e) {
    console.warn("Error accediendo a carpeta de agentes: " + e.message);
  }

  // --- FASE 3: ESCRIBIR REPORTE ---
  let hojaReporte = ssSup.getSheetByName("Reporte_Consolidado_Gestiones");
  if (!hojaReporte) {
    hojaReporte = ssSup.insertSheet("Reporte_Consolidado_Gestiones");
  } else {
    hojaReporte.clear();
  }

  if (reporteFinal.length > 1) {
    hojaReporte.getRange(1, 1, reporteFinal.length, 7).setValues(reporteFinal);
    hojaReporte.getRange(2, 1, reporteFinal.length - 1, 1).setNumberFormat("dd/mm/yyyy hh:mm");
    hojaReporte.autoResizeColumns(1, 7);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Reporte generado con ${reporteFinal.length - 1} gestiones.`,
    'Finalizado', -1
  );

  try {
    SpreadsheetApp.getUi().alert(
      `Proceso completado.\n\n` +
      `Se encontraron ${reporteFinal.length - 1} gestiones en el periodo seleccionado.\n` +
      `Revisa la hoja 'Reporte_Consolidado_Gestiones' en la Planilla Supervisora.`
    );
  } catch(e) {
    // Si se ejecuta por trigger no hay UI disponible, ignorar
  }
}
