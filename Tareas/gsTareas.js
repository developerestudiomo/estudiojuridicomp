const SS_TAREAS_ID = "1Ibh97nTTL7hMPXWLKWhMoJKIi9co3gZzZMeGXdjQ0nE";
const CARPETAS_RAIZ_TIPOS = { 
"PARTICULAR": "1dJNip5JuUtmYljHudu-Ztk8ao8nyfU38", 
"ANNYA": "1AY7X5-HjdeojTnuEhla15PU4MURb1Dh-"
};

function guardarNuevaTarea(obj) {
  try {
    const ss = SpreadsheetApp.openById(SS_TAREAS_ID);
    const sheet = ss.getSheetByName("baseTareas");
    const idTarea = "TASK-" + Math.floor(Math.random() * 1000) + "-" + new Date().getTime().toString().slice(-4);
    
    sheet.appendRow([
      idTarea, new Date(), obj.fechaFin, obj.usuario, 
      obj.categoria, obj.referencia, obj.prioridad, obj.detalle, "PENDIENTE"
    ]);
    
    return { success: true, id: idTarea };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getFuerosTarea(tipo) {
  const carpeta = DriveApp.getFolderById(CARPETAS_RAIZ_TIPOS[tipo]);
  const carpetas = carpeta.getFolders();
  let fueros = [];
  while (carpetas.hasNext()) fueros.push(carpetas.next().getName());
  return fueros.sort();
}

function getDeptosTarea(tipo, fueroNombre) {
  const raizId = CARPETAS_RAIZ_TIPOS[tipo];
  const fueroFolder = DriveApp.getFolderById(raizId).getFoldersByName(fueroNombre).next();
  const subFolders = fueroFolder.getFolders();
  let deptos = [];
  while (subFolders.hasNext()) deptos.push(subFolders.next().getName());
  return deptos.sort();
}

function getExpedientesTarea(tipo, fuero, depto) {
  try {
    const raizId = CARPETAS_RAIZ_TIPOS[tipo];
    const fueroF = DriveApp.getFolderById(raizId).getFoldersByName(fuero).next();
    const deptoF = fueroF.getFoldersByName(depto).next();
    const archivos = deptoF.getFiles();
    
    if (archivos.hasNext()) {
      const ss = SpreadsheetApp.open(archivos.next());
      const sheet = ss.getSheetByName("INDICE");
      
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        // Quitamos cabecera (slice(1))
        // r[0] es la Columna A (Número), r[1] es la Columna B (Carátula/Nombre)
        return data.slice(1)
          .filter(r => r[0] || r[1]) // Evitar filas vacías
          .map(r => {
            const numero = r[0] || "S/N";
            const nombre = r[1] || "Sin nombre";
            return `${numero} - ${nombre}`; // Resultado: "12345 - PEREZ JUAN"
          });
      }
    }
    return ["Sin expedientes en carpeta"];
  } catch (e) {
    console.error("Error en getExpedientesTarea: " + e.message);
    return ["Error al cargar datos"];
  }
}

function finalizarTareaServidor(idTarea) {
  try {
    const ss = SpreadsheetApp.openById(SS_TAREAS_ID);
    const sheetBase = ss.getSheetByName("baseTareas");
    const sheetHistorial = ss.getSheetByName("historialTareas");
    const data = sheetBase.getDataRange().getValues();
    const index = data.findIndex(r => r[0] == idTarea);
    if (index === -1) throw new Error("No encontrada");
    let fila = data[index];
    fila.push(new Date()); 
    sheetHistorial.appendRow(fila);
    sheetBase.deleteRow(index + 1);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function getListaTareas() {
  try {
    const ss = SpreadsheetApp.openById(SS_TAREAS_ID);
    const sheet = ss.getSheetByName("baseTareas");
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    data.shift();
    
    const safeFormat = (val) => {
      if (!val || !(val instanceof Date)) return val;
      return Utilities.formatDate(val, "GMT-3", "dd/MM/yyyy");
    };

    return data.map(r => ({
      id: r[0],
      fechaInicio: safeFormat(r[1]),
      fechaFin: safeFormat(r[2]),
      usuario: r[3],
      categoria: r[4],
      referencia: r[5],
      prioridad: r[6],
      detalle: r[7],
      estado: r[8]
    })).reverse();
  } catch (e) {
    console.error("Error en servidor: " + e.message);
    return [];
  }
}

function guardarNuevaTareaServidor(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Tareas"); // Cambia "Tareas" por el nombre real de tu solapa
    
    // Generar un ID único simple (puedes adaptarlo a tu lógica de IDs)
    const idUnico = Utilities.getUuid();
    
    // Estructura de columnas típica: ID, Referencia, Responsable, Prioridad, Fecha, Detalle, Estado
    hoja.appendRow([
      idUnico,
      datos.referencia,
      datos.usuario,
      datos.prioridad,
      datos.fechaFin,
      datos.detalle,
      "PENDIENTE" // Estado inicial
    ]);
    
    return true;
  } catch (error) {
    throw new Error("Error al guardar en la hoja: " + error.message);
  }
}

// Función para actualizar una tarea existente
function actualizarTareaServidor(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Tareas"); // Cambia "Tareas" por el nombre real de tu solapa
    const rangos = hoja.getDataRange().getValues();
    
    // Buscar la fila por ID (asumiendo que el ID está en la primera columna, índice 0)
    let filaEncontrada = -1;
    for (let i = 1; i < rangos.length; i++) {
      if (rangos[i][0].toString() === datos.id.toString()) {
        filaEncontrada = i + 1; // Las filas en Sheets empiezan en 1
        break;
      }
    }
    
    if (filaEncontrada !== -1) {
      // Actualizar las celdas de esa fila según tu orden de columnas
      hoja.getRange(filaEncontrada, 2).setValue(datos.referencia);
      hoja.getRange(filaEncontrada, 3).setValue(datos.usuario);
      hoja.getRange(filaEncontrada, 4).setValue(datos.prioridad);
      hoja.getRange(filaEncontrada, 5).setValue(datos.fechaFin);
      hoja.getRange(filaEncontrada, 6).setValue(datos.detalle);
      return true;
    } else {
      throw new Error("No se encontró la tarea con el ID especificado.");
    }
  } catch (error) {
    throw new Error("Error al actualizar en la hoja: " + error.message);
  }
}

