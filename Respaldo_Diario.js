function SISTEMA_RESPALDO_POR_FECHAS() {
  // --- CONFIGURACIÓN ---
  const ID_CARPETA_AGENTES = "1kSrb-zocfdEVFxHZGsjcefZMTido0X4N"; // Carpeta donde trabajan los agentes
  const ID_CARPETA_BACKUP_RAIZ = "13EM95hPXJSaB3gfd8iRhbdifE3G2RyXf"; // Carpeta "RESPALDOS DIARIOS AGENTES"
  
  const fechaHoy = new Date();
  // Formato de nombre de carpeta: "YYYY-MM-DD" (Ej: 2026-01-09)
  const nombreCarpetaDia = Utilities.formatDate(fechaHoy, Session.getScriptTimeZone(), "yyyy-MM-dd");

  const carpetaRaiz = DriveApp.getFolderById(ID_CARPETA_BACKUP_RAIZ);
  
  // 1. CREAR (O SELECCIONAR) LA CARPETA DEL DÍA
  let carpetaDia;
  const carpetasExistentes = carpetaRaiz.getFoldersByName(nombreCarpetaDia);
  
  if (carpetasExistentes.hasNext()) {
    carpetaDia = carpetasExistentes.next();
    console.log(`Usando carpeta existente: ${nombreCarpetaDia}`);
  } else {
    carpetaDia = carpetaRaiz.createFolder(nombreCarpetaDia);
    console.log(`Carpeta creada: ${nombreCarpetaDia}`);
  }

  // 2. COPIAR ARCHIVOS (SIN BORRAR NADA ANTIGUO)
  const carpetaAgentes = DriveApp.getFolderById(ID_CARPETA_AGENTES);
  const archivos = carpetaAgentes.getFiles();
  let count = 0;
  
  console.log("Iniciando respaldo...");
  
  while (archivos.hasNext()) {
    const archivoOriginal = archivos.next();
    
    // Solo copiamos Hojas de Cálculo
    if (archivoOriginal.getMimeType() === MimeType.GOOGLE_SHEETS) {
      // Nombre del backup: "[BACKUP] Nombre Agente" (La fecha ya está en la carpeta)
      const nombreBackup = `[BACKUP] ${archivoOriginal.getName()}`;
      archivoOriginal.makeCopy(nombreBackup, carpetaDia);
      count++;
    }
  }
  
  console.log(`✅ Respaldo finalizado con éxito.\n📂 Carpeta: ${nombreCarpetaDia}\n📄 Archivos copiados: ${count}`);
}
