//----gsClientes-----//
const SHEET_CLIENTES_ID = "1aHFKMsh9oVbvwUPHv2OtijSD4xxPLiD2-3_Fe1HSMZA";
const FOLDER_PARTICULAR = "1dJNip5JuUtmYljHudu-Ztk8ao8nyfU38";
const FOLDER_ANNYA = "1AY7X5-HjdeojTnuEhla15PU4MURb1Dh-";

// --- SECCIÓN: GESTIÓN DE CLIENTES (Spreadsheet) ---

function guardarCliente(datos) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_CLIENTES_ID);
    let sheet = ss.getSheetByName("baseClientes");
    
    if (!sheet) {
      sheet = ss.insertSheet("baseClientes");
      sheet.appendRow(["FECHA REGISTRO", "NOMBRE", "DNI", "TEL", "EMAIL", "DOMICILIO", "EXP VINCULADO", "ID"]);
    }

    const idUnico = "CLI-" + new Date().getTime();
    sheet.appendRow([
      datos.fecha,
      datos.nombre,
      datos.dni,
      datos.tel,
      datos.email,
      datos.domicilio,
      datos.expediente || "SIN VINCULAR",
      idUnico
    ]);

    return { success: true, id: idUnico };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function obtenerClientes() {
  console.log("Iniciando obtenerClientes...");
  try {
    const ss = SpreadsheetApp.openById(SHEET_CLIENTES_ID);
    const sheet = ss.getSheetByName("baseClientes");
    
    if (!sheet) {
      console.error("No se encontró la pestaña 'baseClientes'.");
      return [{ nombre: "ERROR DE CARGA", expediente: "No se encontró la pestaña baseClientes" }];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; 

    // Obtenemos los datos de las 8 columnas
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    
    // IMPORTANTE: Convertimos cada celda a String para evitar errores de serialización (null)
    const resultado = data.map(r => {
      return {
        fecha: r[0] ? r[0].toString() : "", 
        nombre: String(r[1] || ""),
        dni: String(r[2] || ""),
        tel: String(r[3] || ""),
        email: String(r[4] || ""),
        domicilio: String(r[5] || ""),
        expediente: String(r[6] || ""),
        id: String(r[7] || "")
      };
    });

    console.log("Carga exitosa. Clientes encontrados: " + resultado.length);
    return resultado;

  } catch (e) {
    console.error("Error en servidor: " + e.toString());
    // Enviamos el error dentro del array para que el SuccessHandler lo capture
    return [{ nombre: "ERROR DE CARGA", expediente: e.toString() }];
  }
}

  function obtenerConfiguracionId() {
    return SHEET_CLIENTES_ID;
  }


/*FUNCIÓN AUXILIAR PARA EVITAR REPETIR CÓDIGO*/
const PREFIJO_NOMBRES = {
    'MO': 'MORON', 'LM': 'LA MATANZA', 'SI': 'SAN ISIDRO', 'SM': 'SAN MARTIN',
    'MZ': 'MERCEDES LUJÁN', 'CA': 'CABA', 'AV': 'AVELLANEDA', 
    'MGRZ': 'MORENO GRAL RODRIGUEZ', 'LZ': 'LOMAS DE ZAMORA' 
};

/** 1. Obtiene Fueros (Subcarpetas de la Rama) */
function getFuerosDesdeDrive(tipo) {
  const folderId = (tipo === "ANNYA") ? FOLDER_ANNYA : FOLDER_PARTICULAR;
  const folder = DriveApp.getFolderById(folderId);
  const sub = folder.getFolders();
  let res = [];
  while (sub.hasNext()) {
    res.push(sub.next().getName());
  }
  return res.sort();
}

/** 2. Obtiene Deptos (Subcarpetas del Fuero) */
function getDepartamentosJudiciales(fueroNombre, tipo) {
  const rootId = (tipo === "ANNYA") ? FOLDER_ANNYA : FOLDER_PARTICULAR;
  const rootFolder = DriveApp.getFolderById(rootId);
  const fueroFolder = rootFolder.getFoldersByName(fueroNombre).next();
  
  const sub = fueroFolder.getFolders();
  let res = [];
  while (sub.hasNext()) {
    res.push(sub.next().getName());
  }
  return res.sort();
}

/** 3. Obtiene los JUZGADOS (Que son archivos de Google Sheets) */
function getExpedientesFromDrive(fuero, depto, tipo) {
  const rootId = (tipo === "ANNYA") ? FOLDER_ANNYA : FOLDER_PARTICULAR;
  const fueroFolder = DriveApp.getFolderById(rootId).getFoldersByName(fuero).next();
  const deptoFolder = fueroFolder.getFoldersByName(depto).next();
  
  const files = deptoFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  let res = [];
  while (files.hasNext()) {
    let f = files.next();
    res.push({ id: f.getId(), name: f.getName() });
  }
  return res.sort((a, b) => a.name.localeCompare(b.name));
}

/** 4. EXTRAE LA LISTA DE EXPEDIENTES DESDE EL INTERIOR DEL SHEET */
function getListaExptesdeId(spreadsheetId) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheets()[0]; 
    const data = sheet.getDataRange().getValues();
    
    // Asumimos que la columna 'EXPEDIENTE' es la segunda (índice 1)
    // Saltamos la cabecera (fila 0)
    let expedientes = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) expedientes.push(data[i][1].toString());
    }
    return [...new Set(expedientes)].sort(); 
  } catch (e) {
    return ["Error al leer archivo del juzgado"];
  }
}
