
// --- CONFIGURACIÓN GLOBAL ---
const ssId = "1dAT2YaqSck9dAz97-TATKZGQ9S2XnyZ8D35pqBTvUYM"; // Base Turnos
const ssNovedadesId = "1L8hA3eMtjvxHOn53on0zFPpEyKdrG8JMNjaim2PyifQ"; // Novedades
const folderIdNovedades = "1kbBwe9S7j3eCkbFAVpWoo36WgfW7iM8C"; // Tu carpeta de Drive - 1kbBwe9S7j3eCkbFAVpWoo36WgfW7iM8C


function obtenerid() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(ss);
  var id= ss.getId()
  Logger.log(id);
}
function getUrl() {
  return ScriptApp.getService().getUrl();
}


//-----------------------INICIO-----------------------//


//---------------------code.gs--------------------------//


// --- FUNCIÓN PRINCIPAL doGet ---
function doGet(e) {
  // 1. Caso para obtener Novedades (JSON)
  if (e && e.parameter && e.parameter.action === "getNovedades") {
    try {
      const datos = obtenerNovedadesSheet();
      return ContentService.createTextOutput(JSON.stringify(datos))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ "error": err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 2. Caso para cargar la Web App (HTML)
  try {
    return HtmlService.createTemplateFromFile('htmlEstructura/index')
      .evaluate()
      .setTitle('Gestión Estudio Jurídico')
      .setFaviconUrl('https://img.icons8.com/ios-filled/50/000000/law.png')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput("Error al cargar la página: " + err.toString());
  }
}



function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* Carga el contenido HTML de una página.*/
function loadPage(pageName) {
  try {
    return HtmlService.createTemplateFromFile(pageName).evaluate().getContent();
  } catch (e) {
    console.error("ERROR CRÍTICO: El archivo '" + pageName + "' no existe o tiene errores. " + e.message);
    console.error(`SERVER: Error fatal al cargar plantilla ${pageName}.html: ${e.message}`);
    throw new Error(`No se pudo cargar la página ${pageName}.`);
  }
}//----------------FIN FUNCIONES DE CARGA E INICIO-----------------------//


//-------------funciones para leer las carpetas del drive---------------//
  const RAICES_DRIVE = {
  "ANNYA": "1VoNdNaLxXXGpmvhVmfp4jun-J8O5JdWl",
  "PARTICULAR": "17S2nQCPQH0-SQ4AYUmDOkXxuANKejpQ2"
};

/*Obtiene el OBJETO carpeta "FUEROS"*/
function getCarpetaFuerosObjeto(tipoCaso) {
  const rootId = RAICES_DRIVE[tipoCaso] || RAICES_DRIVE["PARTICULAR"];
  const carpetaRaiz = DriveApp.getFolderById(rootId);
  const carpetas = carpetaRaiz.getFoldersByName("FUEROS");
  
  if (carpetas.hasNext()) {
    return carpetas.next();
  } else {
    throw new Error("No se encontró la carpeta 'FUEROS' en la gestión: " + tipoCaso);
  }
}

/*Obtiene los Fueros*/
function getFuerosDesdeDrive(tipoCaso) {
  try {
    const carpetaFuerosPadre = getCarpetaFuerosObjeto(tipoCaso);
    const lista = [];
    const subs = carpetaFuerosPadre.getFolders();
    while (subs.hasNext()) {
      const carpeta = subs.next();
      var nombre = String(carpeta.getName());
      if (nombre) lista.push(nombre);
      //lista.push(subs.next().getName().toString());
    }
    return lista.sort();
  } catch (e) {
    Logger.log("Error en getFuerosDesdeDrive: " + e.message);
    throw new Error("Error al leer carpetas de Drive: " + e.message);
  }
}

/*Obtiene los Departamentos*/
function getDepartamentosJudiciales(nombreFuero, tipoCaso) {
  try {
    const carpetaFuerosPadre = getCarpetaFuerosObjeto(tipoCaso);
    const carpetasFuero = carpetaFuerosPadre.getFoldersByName(nombreFuero);
    
    if (!carpetasFuero.hasNext()) return [];
    
    const carpetaFueroElegido = carpetasFuero.next();
    const lista = [];
    const subs = carpetaFueroElegido.getFolders();
    while (subs.hasNext()) {
      lista.push(subs.next().getName());
    }
    return lista.sort();
  } catch (e) {
    throw new Error("Error en Departamentos: " + e.message);
  }
}

/*Obtiene los Juzgados dentro del Departamento*/
function getExpedientesFromDrive(nombreFuero, nombreDepartamento, tipoCaso) {
  try {
    const carpetaFuerosPadre = getCarpetaFuerosObjeto(tipoCaso);
    
    // Navegación segura hacia el departamento
    const f = carpetaFuerosPadre.getFoldersByName(nombreFuero);
    if (!f.hasNext()) return [];
    const d = f.next().getFoldersByName(nombreDepartamento);
    if (!d.hasNext()) return [];
    
    const carpetaDepto = d.next();
    const spreadsheets = carpetaDepto.getFilesByType(MimeType.GOOGLE_SHEETS);
    const lista = [];

    while (spreadsheets.hasNext()) {
      const file = spreadsheets.next();
      const nombreJuzgado = file.getName();
      const subCarpetas = carpetaDepto.getFoldersByName(nombreJuzgado);
      let folderId = subCarpetas.hasNext() ? subCarpetas.next().getId() : "";

      lista.push({
        id: file.getId(), 
        name: nombreJuzgado, 
        folderId: folderId 
      });
    }
    return lista.sort((a, b) => a.name.localeCompare(b.name, 'es', { numeric: true }));
  } catch (error) {
    Logger.log("Error en getExpedientesFromDrive: " + error.message);
    throw new Error("No se pudieron cargar los juzgados.");
  }
}
    //-----FIN DE LAS FUNCIONES DE CARGA DE LOS LISTADOS----/


//------FUNCION QUE Crea la nueva hoja con datos ingresados---------//
function createNewExpedienteSheet(data) {
  try {
    if (!data.juzgadoId || !data.juzgadoFolderId) {
      throw new Error("Faltan identificadores del Juzgado o su Carpeta de destino.");
    }
    const spreadsheet = SpreadsheetApp.openById(data.juzgadoId);
    const sheetName = data.sheetName;
    const parentFolderId = data.juzgadoFolderId; 

    if (spreadsheet.getSheetByName(sheetName)) {
      return { error: `El expediente ${sheetName} ya existe en este juzgado.` };
    }
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const newExpedienteFolder = parentFolder.createFolder(sheetName);
    const newExpedienteFolderId = newExpedienteFolder.getId();

    const sheet = spreadsheet.insertSheet(sheetName);
    const COLOR_ROSA = "#f48fb1"; 

    //Encabezados principales
    sheet.getRange("A1:D1").setValues([["JUZGADO DE", data.fueroId, data.juzgadoName, "TIPO EXPTE"]]);
    //Carga del Tipo de Caso
    sheet.getRange("D2").setValue(data.tipoCaso || "PARTICULAR");
    //Datos de identificación
    sheet.getRange("A2:C2").setValues([["PREFIJO", "NÚMERO DE RECEPTORIA", "AÑO"]]);
    sheet.getRange("A3:C3").setValues([[data.prefijo, data.numeroReceptoria, data.anio]]);
    //Carátula con combinación de celdas
    sheet.getRange("A4").setValue("CARÁTULA");
    sheet.getRange("B4:D4").merge().setValue(data.caratula);
    //Novedades y Fecha Inicio
    sheet.getRange("A5:B5").setValues([["FECHA NOVEDAD", "PASOS PROCESALES DEL EXPEDIENTE"]]);
    sheet.getRange("C5").setValue("FECHA INICIO");
    sheet.getRange("D5").setValue(data.fechaInicial).setNumberFormat('dd/MM/yyyy');

    const titlesRange = sheet.getRangeList(['A1', 'B1', 'D1', 'C5', 'A2:C2', 'A4', 'A5:B5']);
    titlesRange.setBackground(COLOR_ROSA).setFontColor('white').setFontWeight('bold');
    sheet.getRange("D2").setFontWeight("bold").setHorizontalAlignment("center")
      .setBorder(true, true, true, true, null, null, COLOR_ROSA, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // Ajuste de anchos de columna
    sheet.setColumnWidth(1, 150); 
    sheet.setColumnWidth(2, 600); 
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 200); 

    sheet.getRange("A6").setValue(data.fechaInicial).setNumberFormat('dd/MM/yyyy');
    sheet.getRange("B6").setValue(data.novedadesIniciales).setWrap(true);

    //ACTUALIZACIÓN DEL ÍNDICE
    if (typeof actualizarListaExpedientesPorId === "function") {
      actualizarListaExpedientesPorId(spreadsheet);
    }

    // Retornamos el ID de la carpeta
    return {
      success: true,
      folderId: newExpedienteFolderId 
    };

  } catch (e) {
    Logger.log("Error en createNewExpedienteSheet: " + e.toString());
    return { error: "Error en el servidor: " + e.message };
  }
}
//--------------------------------------------------//


//----------------------------getdata----------------------------------//
//----Función que obtiene el listado de los expedientes de la hoja llamada INDICE---//
function getData(id, sheetName = "INDICE") {
  let response = { headers: [], data: [], success: false, error: null };
  try {
    if (!id) throw new Error("ID de Juzgado no válido.");
    const ss = SpreadsheetApp.openById(id);
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      if (typeof actualizarListaExpedientesPorId === 'function') {
        actualizarListaExpedientesPorId(ss);
        sheet = ss.getSheetByName(sheetName);
      }
      if (!sheet) throw new Error("No se pudo encontrar ni generar la hoja INDICE.");
    }
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      response.headers = ["N° EXPTE", "CARÁTULA", "FECHA INICIO", "TIPO EXPTE"];
      response.data = [];
      response.success = true;
      return response;
    }
    const values = sheet.getRange(1, 1, lastRow, 4).getValues();
    // Procesamos encabezados
    response.headers = values[0].map(h => h.toString().toUpperCase().trim());
    response.data = values.slice(1)
      .filter(row => row[0] && row[0].toString().trim() !== "") // Filtra filas sin N° Expte
      .map(row => {
        return row.map(cell => {
          if (cell instanceof Date) {
            return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
          }
          return cell;
        });
      });
    response.success = true;

  } catch (e) {
    console.error("Error en getData:", e.message);
    response.error = "Error en el servidor: " + e.message;
    response.success = false;
  }
  return response;
}

//-------FUNCION QUE ACTUALIZA LISTADO DE LA SHEET INDICE---------------//
 function actualizarListaExpedientesPorId(ss) {
    try {
      if (!ss || typeof ss.getSheetByName !== 'function') {
        ss = SpreadsheetApp.getActiveSpreadsheet();
      }
      
      const NOMBRE_HOJA_LISTADO = 'INDICE';
      let hojaDestino = ss.getSheetByName(NOMBRE_HOJA_LISTADO);

      if (!hojaDestino) {
        hojaDestino = ss.insertSheet(NOMBRE_HOJA_LISTADO, 0);
      } else {
        hojaDestino.clear().clearFormats();
      }
      const encabezados = [['N° EXPTE', 'CARÁTULA', 'FECHA INICIO', 'TIPO EXPTE']];
      const numCols = encabezados[0].length;
      hojaDestino.getRange(1, 1, 1, numCols)
        .setValues(encabezados)
        .setBackground('#ED6AFF')
        .setFontColor('white')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      
      const hojas = ss.getSheets();
      const nuevosDatos = [];

      for (let i = 0; i < hojas.length; i++) {
        const hoja = hojas[i];
        const nombreHoja = hoja.getName();
        if (nombreHoja === NOMBRE_HOJA_LISTADO || nombreHoja.includes('Sheet') || nombreHoja.includes('Hoja') || nombreHoja === 'dataExpedientes') continue;

        try {
          const prefijo = hoja.getRange("A3").getValue();
          const nroReceptoria = hoja.getRange("B3").getValue();
          const anio = hoja.getRange("C3").getValue();
          const caratula = hoja.getRange("B4").getValue();
          const fecha = hoja.getRange('D5').getDisplayValue();
          const tipo = hoja.getRange('D2').getDisplayValue();

          if (!prefijo && !caratula) continue;

          nuevosDatos.push([
            `${prefijo}-${nroReceptoria}-${anio}`,
            caratula.toString().trim() || '(Sin Carátula)',
            fecha || '(Sin Fecha)',
            tipo.toString().trim() || 'NO DEFINIDO'
          ]);
        } catch (err) { console.warn("Error en hoja " + nombreHoja); }
      }

      if (nuevosDatos.length > 0) {
        nuevosDatos.sort((a, b) => a[1].toLowerCase().localeCompare(b[1].toLowerCase()));
        
        const rangoDatos = hojaDestino.getRange(2, 1, nuevosDatos.length, numCols);
        const rangoCompleto = hojaDestino.getRange(1, 1, nuevosDatos.length + 1, numCols);
        rangoDatos.setValues(nuevosDatos);
        rangoCompleto.setBorder(true, true, true, true, true, true).setVerticalAlignment('middle');
        hojaDestino.getRange(2, 2, nuevosDatos.length, 1).setWrap(true);
        SpreadsheetApp.flush(); 
        hojaDestino.autoResizeColumns(1, 4);
        const anchoCol4 = hojaDestino.getColumnWidth(4);
        hojaDestino.setColumnWidth(4, anchoCol4 + 20); 
        hojaDestino.setColumnWidth(2, 500); 
      }
    } catch (e) {
      throw new Error("Error actualizando INDICE: " + e.message);
    }
  }//---FIN FUNCION DE ACTUALIZACION SHEET INDICE--//

//------------------------------------------------------------------//
function handleFormSubmission(formData) {
  try {
    const spreadsheetId = formData.spreadsheetId;
    const data = {
      "N° Expte": formData["expte"],
      "Carátula": formData["caratula"],
      "Fecha inicio": formData["fechaInicio"],
      "Pasos procesales": formData["pasos"],
      "Fecha Novedad": formData["fechaNovedad"],
      "modo": formData["accionExpediente"] // El radio button
    };
    guardarCambiosExpediente(spreadsheetId, data); 
    return HtmlService.createTemplateFromFile('Expedientes/dataExpedientes')
        .evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  } catch (error) {
    Logger.log("ERROR CRÍTICO en handleFormSubmission: " + error.message);
    return HtmlService.createHtmlOutput('<script>showAlert("Error al guardar: ' + error.message + '"); window.history.back();</script>')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }
}
//-----------------------------------------------//

//--Función que Elimina la hoja, la carpeta de Drive y actualiza el índice.--//
function eliminarExpedienteCompleto(spreadsheetId, sheetName, folderId) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    //Eliminar la pestaña del expediente
    const hojaAEliminar = ss.getSheetByName(sheetName);
    if (hojaAEliminar) {
      ss.deleteSheet(hojaAEliminar);
    }

    // Eliminar la carpeta en Google Drive
    if (folderId && folderId !== "null" && folderId !== "") {
      try {
        DriveApp.getFolderById(folderId).setTrashed(true);
      } catch (e) {
        console.warn("No se pudo eliminar la carpeta: " + e.message);
      }
    }
    actualizarListaExpedientesPorId(ss);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

//---NOVEDADES---//
function obtenerNovedadesSheet() {
  const ss = SpreadsheetApp.openById(ssId);
  const hoja = ss.getSheetByName("Novedades");
  const datos = hoja.getDataRange().getValues();
  
  // Quitamos cabecera y filtramos solo los que dicen "SI" en la columna E (index 4)
  return datos.slice(1)
    .filter(fila => fila[4] === "SI")
    .map(fila => ({
      titulo: fila[0],
      descripcion: fila[1],
      etiqueta: fila[2],
      imagen: fila[3]
    }));
}