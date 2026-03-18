/*---Obtiene la lista de archivos de una carpeta específica.---*/
function listarArchivosExpediente(folderId) {
  try {
    if (!folderId || folderId === "null" || folderId === "") return [];
    const carpeta = DriveApp.getFolderById(folderId);
    const archivos = carpeta.getFiles();
    const resultado = [];

    while (archivos.hasNext()) {
      const archivo = archivos.next();
      // Filtramos para ver solo imágenes y PDFs si se desea, o todo
      resultado.push({
        nombre: archivo.getName(),
        embedUrl: archivo.getUrl().replace('/view', '/preview'), 
        id: archivo.getId()
      });
    }
    return resultado;
  } catch (e) {
    console.error("Error al listar archivos: " + e.message);
    return [];
  }
}