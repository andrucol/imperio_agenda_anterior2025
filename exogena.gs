/**
 * Función para consolidar la información ("Exógena") de las hojas diarias (1 al 31).
 * Código organizado y optimizado para mejorar el rendimiento en Google Apps Script.
 */
function exogena() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const nombreLibro = libro.getName();
  
  // 1. Obtener o crear la hoja de resumen 'Exogena'
  let hojaResumen = libro.getSheetByName("Exogena");
  if (!hojaResumen) {
    hojaResumen = libro.insertSheet("Exogena");
  } else {
    hojaResumen.clear(); // Limpiar historial previo para evitar mezclar datos
  }

  // Array maestro donde consolidaremos toda la data junta para escribirla de golpe. 
  // (Las llamadas a servicio como setValues son costosas, mejor hacer una sola al final).
  const datosConsolidados = [];

  // 2. Iterar por las posibles hojas de días del mes (1 al 31)
  for (let dia = 1; dia <= 31; dia++) {
    const hojaDia = libro.getSheetByName(dia.toString());
    
    // Si la hoja no existe (ej: no hay hoja 31 en febrero), pasar a la siguiente
    if (!hojaDia) continue;

    const produccion = hojaDia.getRange("B11").getValue();
    
    // Solo procesar si hay producción reportada mayor a 0
    if (produccion > 0) {
      const ultimaFila = hojaDia.getLastRow();
      
      // Asegurarse de que existan datos a partir de la fila 11
      if (ultimaFila < 11) continue;

      // Optimización: getRange(fila, columna, numFilas, numColumnas)
      // Columna E = 5, hasta Columna I = 5 columnas en total (E, F, G, H, I)
      const rangoCampos = hojaDia.getRange(11, 5, ultimaFila - 10, 5);
      const datosDia = rangoCampos.getValues();

      // Recorrer los datos extraídos para filtrar los que vienen vacíos y formatearlos
      for (const fila of datosDia) {
        // Analizamos si la celda de la columna E (índice 0) tiene texto válido
        if (fila[0] !== undefined && fila[0] !== null && fila[0].toString().trim() !== "") {
          // Agregar la estructura: [Día, NombreLibro, Col_E, Col_F, Col_G, Col_H, Col_I]
          datosConsolidados.push([dia, nombreLibro, ...fila]);
        }
      }
    }
  }

  // 3. Imprimir todos los datos recolectados de una sola vez
  if (datosConsolidados.length > 0) {
    const numFilas = datosConsolidados.length;
    const numCols = datosConsolidados[0].length;
    
    hojaResumen
      .getRange(1, 1, numFilas, numCols)
      .setValues(datosConsolidados);
  }
}
