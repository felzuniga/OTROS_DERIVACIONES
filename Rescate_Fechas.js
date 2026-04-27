function RESCATE_GLOBAL_MULTI_FECHA() {
  const ui = SpreadsheetApp.getUi();
  const start = new Date();

  // --- CONFIGURACIÓN ---
  const ID_HISTORICO_AGENTES = "1zngpyCyipGSmM93Yl98vUy_W_GIkDT-KS0L-qLexP84"; 
  const ID_CARPETA_BACKUP_RAIZ = "13EM95hPXJSaB3gfd8iRhbdifE3G2RyXf"; 
  
  // **NUEVA CONFIGURACIÓN: RANGO DE AGENTES A PROCESAR**
  const AGENTE_INICIO = 9;  // Cambiar este número para cada ejecución
  const AGENTE_FIN = 11;     // Procesar 5 agentes por vez
  
  // Índices Histórico (Base 0)
  const COL_ID = 0;
  const COL_CITA = 14;      // O
  const COL_HORA = 15;      // P
  const COL_GESTION = 23;   // X
  const COL_AGENTE = 26;    // AA

  SpreadsheetApp.getActiveSpreadsheet().toast('Inventariando todas las carpetas de respaldo...', 'Rescate Global', -1);

  // 1. OBTENER TODAS LAS CARPETAS DE RESPALDO (ORDENADAS DE LA MÁS NUEVA A LA MÁS VIEJA)
  const carpetaRaiz = DriveApp.getFolderById(ID_CARPETA_BACKUP_RAIZ);
  const iteradorCarpetas = carpetaRaiz.getFolders();
  let listaCarpetas = [];
  
  while (iteradorCarpetas.hasNext()) {
    listaCarpetas.push(iteradorCarpetas.next());
  }
  
  // Ordenar descendente por nombre (2026-01-10 va antes que 2026-01-09)
  listaCarpetas.sort((a, b) => b.getName().localeCompare(a.getName()));

  if (listaCarpetas.length === 0) {
    ui.alert("No se encontraron carpetas dentro del directorio de respaldos.");
    return;
  }
  
  console.log(`📂 Se encontraron ${listaCarpetas.length} carpetas de respaldo. Se revisarán en este orden:`);
  listaCarpetas.forEach(c => console.log(`   - ${c.getName()}`));

  // 2. LEER HISTÓRICO Y DETECTAR HUECOS
  SpreadsheetApp.getActiveSpreadsheet().toast('Analizando el Histórico...', 'Rescate Global', -1);
  
  const ssHist = SpreadsheetApp.openById(ID_HISTORICO_AGENTES);
  const hojaHist = ssHist.getSheetByName("Gestiones");
  const ultimaFila = hojaHist.getLastRow();
  const dataHist = hojaHist.getRange(2, 1, ultimaFila - 1, 27).getValues(); 
  
  // Mapa: NombreAgente -> [ {rowIndex, id, faltaGestion, faltaCita, reparado} ]
  let mapaTrabajo = new Map();
  let totalHuecos = 0;

  for (let i = 0; i < dataHist.length; i++) {
    const fGestion = dataHist[i][COL_GESTION];
    const fCita = dataHist[i][COL_CITA];
    const idRaw = dataHist[i][COL_ID];
    const agente = String(dataHist[i][COL_AGENTE]).trim();

    if ((fGestion === "" || fCita === "") && idRaw && agente) {
      if (!mapaTrabajo.has(agente)) {
        mapaTrabajo.set(agente, []);
      }
      mapaTrabajo.get(agente).push({
        rowIndex: i + 2,
        id: limpiarId(idRaw),
        faltaGestion: (fGestion === ""),
        faltaCita: (fCita === ""),
        reparado: false
      });
      totalHuecos++;
    }
  }

  if (totalHuecos === 0) {
    ui.alert("El Histórico está perfecto. No hay huecos que reparar.");
    return;
  }

  // **FILTRAR AGENTES SEGÚN EL RANGO CONFIGURADO**
  const todosLosAgentes = Array.from(mapaTrabajo.keys());
  const agentesAProcesar = todosLosAgentes.slice(AGENTE_INICIO - 1, AGENTE_FIN);
  
  console.log(`\n🎯 CONFIGURACIÓN DE EJECUCIÓN:`);
  console.log(`   Total de agentes con huecos: ${todosLosAgentes.length}`);
  console.log(`   Rango a procesar: ${AGENTE_INICIO} - ${AGENTE_FIN}`);
  console.log(`   Agentes en esta ejecución: ${agentesAProcesar.length}`);
  console.log(`   Agentes: ${agentesAProcesar.join(', ')}`);
  
  if (agentesAProcesar.length === 0) {
    ui.alert(
      'Rango vacío',
      `No hay agentes en el rango ${AGENTE_INICIO}-${AGENTE_FIN}.\n\n` +
      `Total de agentes con huecos: ${todosLosAgentes.length}\n\n` +
      `Ajusta AGENTE_INICIO y AGENTE_FIN en el código.`
    );
    return;
  }

  let reparadosTotal = 0;
  let huecosEnRango = 0;
  
  // Contar huecos solo en los agentes que vamos a procesar
  agentesAProcesar.forEach(agente => {
    huecosEnRango += mapaTrabajo.get(agente).length;
  });
  
  console.log(`   Huecos a reparar en este rango: ${huecosEnRango}`);

  // 3. PROCESAR AGENTE POR AGENTE (SOLO EL RANGO SELECCIONADO)
  let contadorAgentes = 0;
  
  for (const nombreAgente of agentesAProcesar) {
    const casosAReparar = mapaTrabajo.get(nombreAgente);
    contadorAgentes++;
    
    // Filtramos solo los casos que aún necesitan reparación
    let pendientes = casosAReparar.filter(c => !c.reparado);
    if (pendientes.length === 0) continue;

    console.log(`\n${'='.repeat(70)}`);
    console.log(`🔎 AGENTE [${contadorAgentes}/${agentesAProcesar.length}]: ${nombreAgente}`);
    console.log(`   Casos a reparar: ${pendientes.length}`);
    console.log(`${'='.repeat(70)}`);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Procesando ${nombreAgente} (${contadorAgentes}/${agentesAProcesar.length})...`, 
      'Rescate Global', 
      -1
    );

    // RECORREMOS LAS CARPETAS (DE HOY HACIA ATRÁS)
    for (const carpeta of listaCarpetas) {
      
      // Si ya arreglamos todo de este agente, pasamos al siguiente
      if (pendientes.length === 0) break;

      // Buscar archivo del agente en ESTA carpeta
      const archivos = carpeta.getFiles();
      let archivoBackup = null;
      while (archivos.hasNext()) {
        const f = archivos.next();
        if (f.getName().toUpperCase().includes(nombreAgente.toUpperCase())) {
          archivoBackup = f;
          break;
        }
      }

      if (!archivoBackup) {
        continue; 
      }

      // LEER BACKUP
      console.log(`   📖 Leyendo backup en ${carpeta.getName()}: "${archivoBackup.getName()}"`);
      const ssBackup = SpreadsheetApp.open(archivoBackup);
      const memoriaBackup = extraerDataEnListas(ssBackup); 

      // INTENTAR REPARAR CON ESTA DATA
      let cambiosEnEstaPasada = 0;
      
      for (let idx = 0; idx < casosAReparar.length; idx++) {
        const target = casosAReparar[idx];
        
        if (target.reparado) continue; // Ya está completo
        
        if (memoriaBackup.has(target.id)) {
          const listaDatos = memoriaBackup.get(target.id);
          
          if (listaDatos.length > 0) {
            const fuente = listaDatos[0]; // Tomamos el primero disponible
            let huboCambio = false;

            // Reparar GESTIÓN si falta
            if (target.faltaGestion && fuente.gestion !== "") {
              console.log(`      ➤ Fila ${target.rowIndex}: Escribiendo F.Gestión = ${fuente.gestion}`);
              hojaHist.getRange(target.rowIndex, COL_GESTION + 1).setValue(fuente.gestion);
              target.faltaGestion = false;
              huboCambio = true;
            }
            
            // Reparar CITA si falta
            if (target.faltaCita && fuente.cita !== "") {
              console.log(`      ➤ Fila ${target.rowIndex}: Escribiendo F.Cita = ${fuente.cita}`);
              hojaHist.getRange(target.rowIndex, COL_CITA + 1).setValue(fuente.cita);
              
              if (fuente.hora !== "") {
                console.log(`      ➤ Fila ${target.rowIndex}: Escribiendo Hora = ${fuente.hora}`);
                hojaHist.getRange(target.rowIndex, COL_HORA + 1).setValue(fuente.hora);
              }
              target.faltaCita = false;
              huboCambio = true;
            }

            // Si ya no le falta nada, marcarlo como reparado
            if (!target.faltaGestion && !target.faltaCita) {
              target.reparado = true;
              console.log(`      ✅ Caso ID ${target.id} COMPLETAMENTE REPARADO`);
            }
            
            if (huboCambio) {
              cambiosEnEstaPasada++;
              reparadosTotal++;
            }
          }
        }
      }
      
      // Forzar que se escriban los cambios
      SpreadsheetApp.flush();
      
      if (cambiosEnEstaPasada > 0) {
         console.log(`      ✅ Recuperados ${cambiosEnEstaPasada} datos en carpeta ${carpeta.getName()}`);
      }
      
      // Actualizamos la lista de pendientes para la siguiente carpeta
      pendientes = casosAReparar.filter(c => !c.reparado);
      
      if (pendientes.length > 0) {
        console.log(`      ⏳ Aún faltan ${pendientes.length} casos por reparar en este agente`);
      }
    }
    
    const noReparados = casosAReparar.filter(c => !c.reparado).length;
    if (noReparados > 0) {
      console.log(`   ⚠️  ${noReparados} casos quedaron sin reparar (no se encontraron en los backups)`);
    }
  }

  const tiempo = ((new Date() - start) / 1000).toFixed(1);
  
  // Calcular siguiente rango sugerido
  const siguienteInicio = AGENTE_FIN + 1;
  const siguienteFin = Math.min(AGENTE_FIN + 5, todosLosAgentes.length);
  
  console.log(`\n${'='.repeat(70)}`);
  console.log(`✅ RESCATE PARCIAL FINALIZADO (Rango ${AGENTE_INICIO}-${AGENTE_FIN})`);
  console.log(`   Carpetas revisadas: ${listaCarpetas.length}`);
  console.log(`   Agentes procesados en esta ejecución: ${agentesAProcesar.length}`);
  console.log(`   Huecos detectados en este rango: ${huecosEnRango}`);
  console.log(`   Datos recuperados: ${reparadosTotal}`);
  console.log(`   Tiempo total: ${tiempo}s`);
  
  if (siguienteInicio <= todosLosAgentes.length) {
    console.log(`\n📋 SIGUIENTE EJECUCIÓN:`);
    console.log(`   Cambiar AGENTE_INICIO = ${siguienteInicio}`);
    console.log(`   Cambiar AGENTE_FIN = ${siguienteFin}`);
  } else {
    console.log(`\n🎉 ¡TODOS LOS AGENTES HAN SIDO PROCESADOS!`);
  }
  console.log(`${'='.repeat(70)}\n`);
  
  let mensaje = `Rango procesado: Agentes ${AGENTE_INICIO}-${AGENTE_FIN}\n` +
    `Agentes procesados: ${agentesAProcesar.length}\n` +
    `Huecos en este rango: ${huecosEnRango}\n` +
    `Datos recuperados: ${reparadosTotal}\n` +
    `Tiempo: ${tiempo}s\n\n`;
  
  if (siguienteInicio <= todosLosAgentes.length) {
    mensaje += `📋 SIGUIENTE EJECUCIÓN:\n` +
      `AGENTE_INICIO = ${siguienteInicio}\n` +
      `AGENTE_FIN = ${siguienteFin}`;
  } else {
    mensaje += `🎉 ¡TODOS LOS AGENTES COMPLETADOS!`;
  }
  
  ui.alert('Rescate Parcial Finalizado', mensaje, ui.ButtonSet.OK);
}

// --- FUNCIONES AUXILIARES ---

function extraerDataEnListas(ss) {
  const mapa = new Map();
  const hojas = ["Agendar", "Notificar"];

  hojas.forEach(nombreHoja => {
    const hoja = ss.getSheetByName(nombreHoja);
    if (hoja && hoja.getLastRow() > 1) {
      const data = hoja.getRange(2, 1, hoja.getLastRow() - 1, 24).getValues();
      const isAgendar = (nombreHoja === "Agendar");

      data.forEach(fila => {
        const id = limpiarId(fila[0]);
        if (!id) return;

        const fGestion = fila[23]; 
        let fCita, hCita;

        if (isAgendar) {
           fCita = fila[14]; 
           hCita = fila[15]; 
        } else {
           fCita = fila[10]; 
           hCita = fila[11]; 
        }

        const datoFila = { gestion: fGestion, cita: fCita, hora: hCita };

        if (!mapa.has(id)) mapa.set(id, []);
        
        if (fGestion !== "" || fCita !== "") {
           mapa.get(id).push(datoFila);
        }
      });
    }
  });
  return mapa;
}

function limpiarId(id) {
  if (!id) return null;
  return String(id).replace(/['"]/g, '').trim().toUpperCase();
}
