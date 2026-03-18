function obtenerNovedadesSheet() {
  const ss = SpreadsheetApp.openById(ssNovedadesId);
  const hoja = ss.getSheetByName("Novedades");
  if (!hoja) throw new Error("No se encontró la pestaña 'Novedades'");
  
  const valores = hoja.getDataRange().getValues();
  if (valores.length <= 1) return []; // Si solo están los encabezados

  // Procesamos las filas saltando la primera (encabezados)
  return valores.slice(1)
    .filter(function(fila) { 
      return fila[4] === "SI"; // Columna E: Activo
    })
    .map(function(fila) {
      let imagenLink = fila[3] ? fila[3].toString().trim() : "";
      
      // Si el campo tiene texto pero no es un link, buscamos en Drive
      if (imagenLink && !imagenLink.toLowerCase().startsWith("http")) {
        imagenLink = obtenerLinkImagenDirecto(imagenLink);
      }

      return {
        titulo: fila[0],      // Columna A
        descripcion: fila[1], // Columna B
        etiqueta: fila[2],    // Columna C
        imagen: imagenLink    // Columna D
      };
    });
}

// --- BUSCADOR DE IMÁGENES EN DRIVE ---
function obtenerLinkImagenDirecto(nombreArchivo) {
  try {
    const carpeta = DriveApp.getFolderById(folderIdNovedades);
    const archivos = carpeta.getFilesByName(nombreArchivo.toString().trim());
    
    if (archivos.hasNext()) {
      const archivo = archivos.next();
      archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      // ESTA ES LA URL SECRETA QUE MEJOR FUNCIONA PARA IMÁGENES DE FONDO
      return "https://lh3.googleusercontent.com/d/" + archivo.getId();
    }
  } catch (e) {
    console.error("Error: " + e.toString());
  }
  return "";
}