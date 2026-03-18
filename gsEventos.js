//----gsEventos---//
const RAIZ_SHEET = {
  "EVENTOS": "17x9WOGoUxJ52yC_bxwS3ovftNK41GTBYQiOUfBn4YdI",
};

const CARPETAS_RAIZ = {
  "PARTICULAR": "1dJNip5JuUtmYljHudu-Ztk8ao8nyfU38",
  "ANNYA": "1dJNip5JuUtmYljHudu-Ztk8ao8nyfU38" 
};

//Obtiene los Fueros dependiendo de la rama seleccionada//
function getFuerosDriveRaiz(tipoRama) {
  try {
    const idRaiz = CARPETAS_RAIZ[tipoRama];
    if (!idRaiz) throw new Error("Rama no definida o ID de carpeta faltante.");

    const carpetaRaiz = DriveApp.getFolderById(idRaiz);
    const carpetas = carpetaRaiz.getFolders();
    let fueros = [];

    while (carpetas.hasNext()) {
      let f = carpetas.next();
      fueros.push({ id: f.getId(), name: f.getName() });
    }
    return fueros.sort((a, b) => a.name.localeCompare(b.name));
  } catch (e) {
    console.error("Error en getFuerosDriveRaiz: " + e.toString());
    return [];
  }
}

//Obtiene Deptos Judiciales//
function getDeptosJudiciales(idFuero) {
  try {
    const carpetaFuero = DriveApp.getFolderById(idFuero);
    const subcarpetas = carpetaFuero.getFolders();
    let deptos = [];

    while (subcarpetas.hasNext()) {
      let d = subcarpetas.next();
      deptos.push({ id: d.getId(), name: d.getName() });
    }
    return deptos.sort((a, b) => a.name.localeCompare(b.name));
  } catch (e) {
    return [];
  }
}

//Obtiene Juzgados//
function getExptesFromDrive(idFuero, idDepto) {
  try {
    const carpetaDepto = DriveApp.getFolderById(idDepto);
    const archivos = carpetaDepto.getFilesByType(MimeType.GOOGLE_SHEETS);
    let juzgados = [];

    while (archivos.hasNext()) {
      let a = archivos.next();
      juzgados.push({ id: a.getId(), name: a.getName() });
    }
    return juzgados.sort((a, b) => a.name.localeCompare(b.name));
  } catch (e) {
    return [];
  }
}

//Obtiene la lista de expedientes desde el INDICE del juzgado seleccionado//
function getListaExptesdeId(idJuzgado) {
  try {
    if (!idJuzgado) return ["Error: ID no recibido"];

    // Se asume que la función getData existe en tu proyecto para leer Sheets
    const resultado = getData(idJuzgado, "INDICE");

    if (!resultado.success) throw new Error(resultado.error);

    // Mapeamos para el selector: "N° RECEPTORIA - CARATULA"
    return resultado.data.map(r => `${r[0]} - ${r[1]}`);

  } catch (e) {
    console.error("Error en getListaExptesdeId: " + e.message);
    return ["No se pudieron cargar expedientes"];
  }
}

//Guarda el evento en Sheet y sincroniza con Google Calendar//
function guardarEvento(obj) {
  try {
    const ss = SpreadsheetApp.openById(RAIZ_SHEET.EVENTOS);
    let sheet = ss.getSheetByName("Gestor eventos");
    
    if (!sheet) {
      sheet = ss.insertSheet("Gestor eventos");
      // Encabezados claros
      sheet.appendRow(["ID EVENTO", "FECHA HORA", "MOTIVO", "TIPO RAMA", "FUERO", "JUZGADO", "EXPEDIENTE", "DETALLES", "CALENDAR ID"]);
      sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#F3E5F5");
    }

    const calendar = CalendarApp.getDefaultCalendar();
    const inicio = new Date(obj.fecha);
    const fin = new Date(inicio.getTime() + (60 * 60 * 1000)); 

    // Sincronización con Calendar
    const calEvent = calendar.createEvent(`[${obj.tipoRama}] ${obj.motivo}: ${obj.expediente}`, inicio, fin, {
      description: `Fuero: ${obj.fuero}\nJuzgado: ${obj.juzgadoInfo}\nDetalles: ${obj.detalle}`
    });

    const id = "EV-" + new Date().getTime();
    
    sheet.appendRow([
      id, 
      obj.fecha, 
      obj.motivo, 
      obj.tipoRama,
      obj.fuero, 
      obj.juzgadoInfo,
      obj.expediente, 
      obj.detalle, 
      calEvent.getId()
    ]);

    return { success: true, id: id };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

//Obtención de eventos para la vista de listado/
function obtenerEventos() {
  try {
    const ss = SpreadsheetApp.openById(RAIZ_SHEET.EVENTOS);
    const sheet = ss.getSheetByName("Gestor eventos");
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

    return data.map(r => ({
      id: r[0],
      fecha: r[1] instanceof Date ? r[1].toISOString() : r[1],
      motivo: r[2],
      tipoRama: r[3],
      fuero: r[4],
      juzgado: r[5],
      expediente: r[6],
      detalle: r[7],
      calId: r[8]
    })).sort((a, b) => new Date(a.fecha) - new Date(b.fecha));

  } catch (e) {
    console.error("Error en obtenerEventos: " + e.message);
    return [];
  }
}

//Elimina evento de Sheet y Calendar//
function eliminarEvento(id, calId) {
  try {
    if (calId) {
      const cal = CalendarApp.getDefaultCalendar();
      const ev = cal.getEventById(calId);
      if (ev) ev.deleteEvent();
    }

    const ss = SpreadsheetApp.openById(RAIZ_SHEET.EVENTOS);
    const sheet = ss.getSheetByName("Gestor eventos");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/* Mueve un evento de la hoja activa a la hoja de Historial*/
function finalizarEvento(id, calId) {
  try {
    const ss = SpreadsheetApp.openById(RAIZ_SHEET.EVENTOS);
    const sheetOrigen = ss.getSheetByName("Gestor eventos");
    let sheetDestino = ss.getSheetByName("Historial");

    //Crear la hoja Historial si no existe
    if (!sheetDestino) {
      sheetDestino = ss.insertSheet("Historial");
      sheetDestino.appendRow(["ID EVENTO", "FECHA HORA", "MOTIVO", "TIPO RAMA", "FUERO", "JUZGADO", "EXPEDIENTE", "DETALLES", "CALENDAR ID", "FECHA FINALIZADO"]);
      sheetDestino.getRange("A1:J1").setFontWeight("bold").setBackground("#D1FFD6");
    }

    const data = sheetOrigen.getDataRange().getValues();
    let filaEncontrada = -1;
    let datosEvento = [];

    //Buscar la fila por ID
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        filaEncontrada = i + 1;
        datosEvento = data[i];
        break;
      }
    }

    if (filaEncontrada !== -1) {
      //Eliminar de Google Calendar
      if (calId) {
        try {
          const cal = CalendarApp.getDefaultCalendar();
          const ev = cal.getEventById(calId);
          if (ev) ev.deleteEvent();
        } catch (e) { console.warn("No se pudo borrar de Calendar, tal vez ya no existe."); }
      }

      //Agregar a Historial con fecha de finalización
      const ahora = new Date();
      datosEvento.push(ahora); // Añadimos la fecha de cierre
      sheetDestino.appendRow(datosEvento);

      //Eliminar de la hoja activa
      sheetOrigen.deleteRow(filaEncontrada);
      
      return { success: true };
    } else {
      return { success: false, error: "No se encontró el evento." };
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
function obtenerHistorial() {
  try {
    const ss = SpreadsheetApp.openById(RAIZ_SHEET.EVENTOS);
    const sheet = ss.getSheetByName("Historial");
    
    // Si la hoja no existe aún, devolvemos lista vacía
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    // Si solo está el encabezado, devolvemos lista vacía
    if (lastRow <= 1) return [];

    // Obtenemos los datos desde la fila 2 hasta la última, columnas A a J (10 columnas)
    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

    return data.map(r => ({
      id: String(r[0]),
      fecha: r[1] instanceof Date ? r[1].toISOString() : r[1],
      motivo: String(r[2]),
      tipoRama: String(r[3]),
      fuero: String(r[4]),
      juzgado: String(r[5]),
      expediente: String(r[6]),
      detalle: String(r[7]),
      calId: String(r[8]),
      fechaFinalizado: r[9] instanceof Date ? r[9].toISOString() : r[9]
    })).sort((a, b) => new Date(b.fechaFinalizado) - new Date(a.fechaFinalizado)); // Ordenar por fecha de cierre

  } catch (e) {
    console.error("Error en obtenerHistorial: " + e.message);
    return [];
  }
}
