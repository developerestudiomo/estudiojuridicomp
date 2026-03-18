//-------------gsConvenios--------//
const CARPETA_RAIZ_PARTICULAR_1 = "1dJNip5JuUtmYljHudu-Ztk8ao8nyfU38";
const SHEET_CLIENTES_ID_1 = "1aHFKMsh9oVbvwUPHv2OtijSD4xxPLiD2-3_Fe1HSMZA";
const SHEET_CONVENIOS_ID ="1CNeTz3y_DWnQ8RCdEfey4eVXO6V7rWBYJlEyw5TRb6c";


function getListaClientesConvenio() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_CLIENTES_ID_1);
    const sheet = ss.getSheetByName("baseClientes");
    const data = sheet.getDataRange().getValues();
    const clientes = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) { // Verifica columna B
        clientes.push({
          nombre: String(data[i][1]),    
          dni: String(data[i][2]),       
          domicilio: String(data[i][5]) 
        });
      }
    }
    return clientes.sort((a, b) => a.nombre.localeCompare(b.nombre));
  } catch (e) { throw new Error("Error Clientes: " + e.message); }
}

function getFuerosConvenio() {
  try {
    const root = DriveApp.getFolderById(CARPETA_RAIZ_PARTICULAR_1);
    const carpetas = root.getFolders();
    let lista = [];

    while (carpetas.hasNext()) {
      let f = carpetas.next();
      lista.push({ 
        id: String(f.getId()), 
        name: String(f.getName()) 
      });
    }

    if (lista.length === 0) return [{ name: "No se encontraron carpetas", id: "" }];

    return lista.sort((a, b) => a.name.localeCompare(b.name));
  } catch (e) {
    return [{ name: "Error: " + e.message, id: "" }];
  }
}

function getDeptosConvenio(nombreFuero) {
  try {
    const root = DriveApp.getFolderById(CARPETA_RAIZ_PARTICULAR_1);
    const carpetasFuero = root.getFoldersByName(nombreFuero);
    
    if (!carpetasFuero.hasNext()) {
      console.error("No se encontró la carpeta: " + nombreFuero);
      return []; 
    }
    
    const carpetaPadre = carpetasFuero.next();
    const subCarpetas = carpetaPadre.getFolders();
    let lista = [];
    
    while (subCarpetas.hasNext()) {
      lista.push(subCarpetas.next().getName());
    }
    return lista.sort();
  } catch (e) {
    throw new Error("Error en Deptos: " + e.message);
  }
}

function getJuzgadosConvenio(nombreFuero, nombreDepto) {
  try {
    const root = DriveApp.getFolderById(CARPETA_RAIZ_PARTICULAR_1);
    const fItr = root.getFoldersByName(nombreFuero);
    if (!fItr.hasNext()) return [];
    
    const dItr = fItr.next().getFoldersByName(nombreDepto);
    if (!dItr.hasNext()) return [];
    
    const carpetaFinal = dItr.next();
    const files = carpetaFinal.getFilesByType(MimeType.GOOGLE_SHEETS);
    let lista = [];
    
    while (files.hasNext()) {
      lista.push(files.next().getName());
    }
    return lista.sort();
  } catch (e) {
    throw new Error("Error en Juzgados: " + e.message);
  }
}

function guardarConvenioFinal(datos) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_CONVENIOS_ID);
    const sheet = ss.getSheetByName("baseConvenios");
    const idGenerado = "CONV-" + new Date().getTime();
    
    sheet.appendRow([
      idGenerado, 
      new Date(),
      datos.cliente,
      datos.materia,
      datos.tipo,
      datos.cantJus,
      datos.porcentajeJus,
      datos.cuotaJus
    ]);
    return idGenerado;
  } catch (e) { throw new Error("Error al guardar: " + e.message); }
}

function generarPDFConvenio(htmlCuerpo, nombreCliente, tipoConvenio) {
  try {
    const ID_CARPETA_DESTINO = "1J4cQlxvswHy33eqOW-isbyPhelGcAU1x";
    const carpetaDestino = DriveApp.getFolderById(ID_CARPETA_DESTINO);

    const fecha = Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd");
    const nombreArchivo = `${tipoConvenio} - ${nombreCliente} - ${fecha}.pdf`;

    const htmlFinal = `<div style="font-family: 'Times New Roman', serif; padding: 30px; text-align: justify;">${htmlCuerpo}</div>`;
    
    const blob = Utilities.newBlob(htmlFinal, "text/html", "temp.html");
    const pdfBlob = blob.getAs("application/pdf").setName(nombreArchivo);
    const archivoPDF = carpetaDestino.createFile(pdfBlob);
    
    return { success: true, url: archivoPDF.getUrl() };
  } catch (e) { throw new Error("Error PDF: " + e.message); }
}

/** Obtener el logo del estudio en formato Base64 **/
function getLogoEstudio() {
  try {
    const fileId = "1tZ-aMLpZqkCSOKxDzeobMyhVIX_qdke8";
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const contentType = blob.getContentType();
    const base64 = Utilities.base64Encode(blob.getBytes());
    
    return `data:${contentType};base64,${base64}`;
  } catch (e) {
    console.error("Error al obtener el logo: " + e.message);
    return "";
  }
}

// Obtener Fueros según la Raíz (Particular o NNYA)
function getFuerosDesdeRaiz(raizId) {
  const root = DriveApp.getFolderById(raizId);
  const carpetas = root.getFolders();
  let lista = [];
  while (carpetas.hasNext()) {
    let f = carpetas.next();
    lista.push({ name: f.getName(), id: f.getId() });
  }
  return lista.sort((a, b) => a.name.localeCompare(b.name));
}

// Deptos Dinámicos
function getDeptosDinamicos(raizId, fuero) {
  const fueroFolder = DriveApp.getFolderById(raizId).getFoldersByName(fuero).next();
  const sub = fueroFolder.getFolders();
  let lista = [];
  while (sub.hasNext()) lista.push(sub.next().getName());
  return lista.sort();
}

// Juzgados Dinámicos (Busca los archivos Excel)
function getJuzgadosDinamicos(raizId, fuero, depto) {
  const fueroFolder = DriveApp.getFolderById(raizId).getFoldersByName(fuero).next();
  const deptoFolder = fueroFolder.getFoldersByName(depto).next();
  const files = deptoFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  let lista = [];
  while (files.hasNext()) lista.push(files.next().getName());
  return lista.sort();
}


function getListaExpedientesDinamicos(raizId, fuero, depto, nombreExcel) {
  try {

    const fueroFolder = DriveApp.getFolderById(raizId).getFoldersByName(fuero).next();
    const deptoFolder = fueroFolder.getFoldersByName(depto).next();
    const files = deptoFolder.getFilesByName(nombreExcel);
    
    if (!files.hasNext()) return ["Archivo no encontrado"];
    const idArchivo = files.next().getId();


    const resultado = getData(idArchivo, "INDICE");

    if (!resultado.success) throw new Error(resultado.error);
    return resultado.data.map(r => `${r[0]} - ${r[1]}`).sort();

  } catch (e) {
    console.error("Error en getListaExpedientesDinamicos: " + e.message);
    return ["No se pudieron cargar expedientes"];
  }
}
//----------------GESTOR DE PAGOS--------------//

function abrirModalCargarPagos() {
  Swal.fire('Gestor de Pagos', 'Cargando base de pagos...', 'info');
}
function listarConveniosVigentes() {
    loadContent('getListaConveniosParaPago');
}

function sumarDiasHabiles(fecha, dias) {
  let f = new Date(fecha);
  let count = 0;
  while (count < dias) {
    f.setDate(f.getDate() + 1);
    if (f.getDay() !== 0 && f.getDay() !== 6) { // 0=Dom, 6=Sab
      count++;
    }
  }
  return f;
}

function registrarEnBasePagos(datos) {
  const ssId = "1CNeTz3y_DWnQ8RCdEfey4eVXO6V7rWBYJlEyw5TRb6c";
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName("basePagos");
  const lastRow = sheet.getLastRow() + 1;
  
  const hoy = new Date();
  // Vencimientos base (mensuales)
  const v1 = new Date(); v1.setMonth(hoy.getMonth() + 1);
  const v2 = new Date(); v2.setMonth(hoy.getMonth() + 2);
  const v3 = new Date(); v3.setMonth(hoy.getMonth() + 3);

  // Preparamos la fila. 
  // Nota: Las columnas D, E, F, G y N llevan TUS FÓRMULAS de la hoja.
  // Usamos setFormulasR1C1 para que se adapten a la fila automáticamente.
  
  const filaData = [
    datos.id,      
    datos.tipo,    
    datos.cliente, 
    "", "", "", "",
    v1, "",        
    v2, "",        
    v3, "",        
    "",            
    "AL DÍA" 
  ];

  sheet.appendRow(filaData);
  sheet.getRange(lastRow, 4).setFormula(`=IFERROR(VLOOKUP(A${lastRow};'baseConvenios'!A:H;2;FALSE);"")`); // Ejemplo para Fecha
  sheet.getRange(lastRow, 5).setFormula(`=IFERROR(VLOOKUP(A${lastRow};'baseConvenios'!A:H;6;FALSE);"")`); // Ejemplo para Total
  sheet.getRange(lastRow, 6).setFormula(`=IFERROR(VLOOKUP(A${lastRow};'baseConvenios'!A:H;7;FALSE);"")`); // Ejemplo para 40%
  sheet.getRange(lastRow, 7).setFormula(`=IFERROR(VLOOKUP(A${lastRow};'baseConvenios'!A:H;8;FALSE);"")`); // Ejemplo para Cuota
  sheet.getRange(lastRow, 14).setFormula(`=SI(E${lastRow}="";"";(E${lastRow}-F${lastRow})-I${lastRow}-K${lastRow}-M${lastRow})`);
}

function getListaConveniosParaPago() {
  const ssId = "1CNeTz3y_DWnQ8RCdEfey4eVXO6V7rWBYJlEyw5TRb6c";
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName("basePagos");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  data.shift(); 
  return data
    .filter(r => r[0] !== "") 
    .map(r => ({ id: String(r[0]), cliente: String(r[2]) }))
    .sort((a, b) => a.cliente.localeCompare(b.cliente));
}

function getDetalleConvenioPorId(id) {
  const ssId = "1CNeTz3y_DWnQ8RCdEfey4eVXO6V7rWBYJlEyw5TRb6c";
  const sheetPagos = SpreadsheetApp.openById(ssId).getSheetByName("basePagos");
  const dataPagos = sheetPagos.getDataRange().getValues();
  
  const fila = dataPagos.find(r => r[0] == id);
  if (!fila) return null;

  const safeFormat = (val) => {
    if (!val || !(val instanceof Date) || isNaN(val.getTime())) return "";
    return Utilities.formatDate(val, "GMT-3", "dd/MM/yyyy");
  };

  return {
    id: fila[0],
    cliente: fila[2],
    valorCuota: fila[6],
    venc1: safeFormat(fila[7]),
    pago1: safeFormat(fila[8]),
    venc2: safeFormat(fila[9]),
    pago2: safeFormat(fila[10]),
    venc3: safeFormat(fila[11]),
    pago3: safeFormat(fila[12])
  };
}

function registrarPagoEnSheet(id, numCuota) {
  const ssId = "1CNeTz3y_DWnQ8RCdEfey4eVXO6V7rWBYJlEyw5TRb6c";
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName("basePagos");
  const data = sheet.getDataRange().getValues();
  
  let indexFila = data.findIndex(r => r[0] == id);
  if (indexFila === -1) return false;
  
  let numFila = indexFila + 1;
  let colPago;
  
  if (numCuota == 1) colPago = 9;
  else if (numCuota == 2) colPago = 11;
  else if (numCuota == 3) colPago = 13;
  
  sheet.getRange(numFila, colPago).setValue(new Date());
  
  sheet.getRange(numFila, 15).setValue("AL DÍA");
  
  return true;
}




