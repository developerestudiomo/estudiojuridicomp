//const ssId = "1dAT2YaqSck9dAz97-TATKZGQ9S2XnyZ8D35pqBTvUYM";
const sheetEventosId = "17x9WOGoUxJ52yC_bxwS3ovftNK41GTBYQiOUfBn4YdI";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(ssId);
    const hoja = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
    
    hoja.appendRow([
      new Date(),
      data.nombre,
      data.telefono,
      data.motivo,
      data.origen || "Web"
    ]);
    
    return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput("Error: " + err.toString());
  }
}

function obtenerBandejaEntrada() {
  try {
    const ss = SpreadsheetApp.openById(ssId);
    const hoja = ss.getSheetByName("baseTurnos");
    
    if (!hoja) return [];

    const lastRow = hoja.getLastRow();
    if (lastRow <= 1) return [];

    // Leemos los valores visibles (textos)
    const data = hoja.getRange(2, 1, lastRow - 1, 5).getDisplayValues();

    const resultado = data.map((fila, index) => {
      const motivoVal = fila[3] ? fila[3].toString().toLowerCase() : "";
      const colorHex = motivoVal.includes("familia") ? "#ED6AFF" : "#AD46FF";
      
      return {
        idFila: index + 2,
        fechaTexto: fila[0], // Guardamos el texto tal cual: "24/02/2026 16:22:50"
        nombre: fila[1],
        whatsapp: fila[2],
        motivo: fila[3],
        estado: fila[4] ? fila[4].toString().trim() : "",
        color: colorHex
      };
    })
    .filter(item => item.estado.toLowerCase() === "pendiente");

    // Ordenar por texto de fecha (al ser DD/MM/YYYY requiere un pequeño ajuste para ordenar bien)
    return resultado.sort((a, b) => {
      const [fechaA, horaA] = a.fechaTexto.split(' ');
      const [d1, m1, y1] = fechaA.split('/');
      const [fechaB, horaB] = b.fechaTexto.split(' ');
      const [d2, m2, y2] = fechaB.split('/');
      
      // Comparamos Año-Mes-Día para un orden correcto
      return new Date(y1, m1-1, d1) - new Date(y2, m2-1, d2);
    });

  } catch (e) {
    console.error("Error: " + e.message);
    return [];
  }
}

function marcarTurnoComoAgendado(idFila) {
  try {
    const ss = SpreadsheetApp.openById(ssId);
    const hoja = ss.getSheetByName("baseTurnos");
    // La columna 5 es la del Estado (E)
    hoja.getRange(idFila, 5).setValue("Agendado");
    return true;
  } catch (e) {
    return false;
  }
}

function procesarAgendadoCompleto(obj) {
  try {
    // 1. Marcar como Agendado en la baseTurnos
    const ssOrigen = SpreadsheetApp.openById(ssId);
    ssOrigen.getSheetByName("baseTurnos").getRange(obj.idFila, 5).setValue("Agendado");

    // 2. Insertar en la Hoja "Gestor eventos"
    const ssDestino = SpreadsheetApp.openById(sheetEventosId);
    const hojaEventos = ssDestino.getSheetByName("Gestor eventos");
    
    // Combinar fecha y hora para el objeto Date
    const startDateTime = new Date(obj.fechaCita + 'T' + obj.horaCita);
    const endDateTime = new Date(startDateTime.getTime() + 60 * 60 * 1000); // +1 hora por defecto

    hojaEventos.appendRow([
      new Date(), // Fecha registro
      obj.fechaCita,
      obj.horaCita,
      obj.nombre,
      obj.whatsapp,
      obj.motivo,
      obj.notas,
      "Programado" // Estado inicial en el gestor
    ]);

    // 3. Crear evento en Google Calendar (Calendario principal)
    const cal = CalendarApp.getDefaultCalendar();
    cal.createEvent(
      "Cita: " + obj.nombre + " - " + obj.motivo,
      startDateTime,
      endDateTime,
      {
        description: "WhatsApp: " + obj.whatsapp + "\nNotas: " + obj.notas,
        location: "Estudio Jurídico"
      }
    );

    return true;
  } catch (e) {
    console.error("Error en procesarAgendadoCompleto: " + e.toString());
    throw new Error("No se pudo completar el registro: " + e.message);
  }
}
