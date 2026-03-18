//-----------------------------CODIGO GS-------------------------------//
//FUNCION QUE CARGA datos de expediente en PAGINA verExpediente.html//
function getListData(spreadsheetId, sheetName, nombreFuero, nombreDepartamento, tipoCaso) {
  let log = [];
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const ssName = ss.getName();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: "Hoja no encontrada." };

    let folderId = "";
    const nombreExpediente = sheetName.trim();
    
    //Obtenego carpeta Padre del archivo
    const ssFile = DriveApp.getFileById(spreadsheetId);
    const parentFolders = ssFile.getParents();
    
    if (parentFolders.hasNext()) {
      const carpetaLM = parentFolders.next();
      log.push("📂 En carpeta Departamento: " + carpetaLM.getName());
      
      //Busco la carpeta hija
      const juzgadoFolders = carpetaLM.getFoldersByName(ssName);
      if (juzgadoFolders.hasNext()) {
        const carpetaJuzgado = juzgadoFolders.next();
        log.push("📂 Entrando a carpeta Juzgado: " + carpetaJuzgado.getName());
        
        //Busco carpeta del expediente
        const expedienteFolders = carpetaJuzgado.getFoldersByName(nombreExpediente);
        if (expedienteFolders.hasNext()) {
          const fldFinal = expedienteFolders.next();
          folderId = fldFinal.getId();
          log.push("✅ Carpeta de expediente localizada: " + fldFinal.getName());
        } else {
          log.push("❌ No existe la carpeta del expediente '" + nombreExpediente + "' dentro de '" + ssName + "'");
        }
      } else {
        log.push("❌ No se encontró la carpeta del Juzgado '" + ssName + "' dentro de '" + carpetaLM.getName() + "'");
      }
    }

    const data = sheet.getRange("A1:C5").getValues();
    const header = {
      fueroId: data[0][1] || "S/D",
      juzgadoName: data[0][2] || "S/D",
      expedienteNumero: sheetName,
      caratula: data[3][1] || "Sin Carátula",
      folderId: folderId 
    };

    const novedades = sheet.getDataRange().getValues().slice(5)
      .filter(row => row[0] !== "" && row[0] !== null)
      .map(row => {
        let f = row[0] instanceof Date ? Utilities.formatDate(row[0], "GMT-3", "dd/MM/yyyy") : row[0];
        return [f, row[1] || ""];
      });

    return { success: true, header: header, novedades: novedades, headers: ["FECHA", "DETALLE"], serverLog: log };
  } catch (e) {
    return { success: false, error: e.message, serverLog: log };
  }
}
//------------------------------//
function addNovedadExpediente(spreadsheetId, sheetName, fecha, texto) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    const lastRow = Math.max(sheet.getLastRow(), 5);
    
    const fechaObj = new Date(fecha + "T05:00:00"); // Ajuste zona horaria manual simple
    sheet.getRange(lastRow + 1, 1).setValue(fechaObj).setNumberFormat("dd/MM/yyyy");
    sheet.getRange(lastRow + 1, 2).setValue(texto).setWrap(true);
    
    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}
function uploadBase64File(base64Data, contentType, fileName, folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, fileId: file.getId() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

