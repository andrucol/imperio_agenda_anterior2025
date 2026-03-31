/**
 * ============================================
 * exogena.gs - REPORTE CONSOLIDADO MENSUAL
 * ============================================
 *
 * Recorre las hojas del libro activo nombradas del "1" al "31"
 * (cada una representa un día del mes con datos de producción),
 * extrae los registros válidos y los consolida en una sola hoja
 * llamada "Exogena" para generar un reporte unificado.
 *
 * OPTIMIZACIÓN: Toda la data se acumula en un array en memoria
 * y se escribe con un solo setValues() al final, evitando
 * múltiples llamadas costosas al servicio de Sheets.
 */
function exogena() {
  // Referencia al libro de cálculo activo (el que tiene esta macro)
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  // Nombre del libro, se incluye como columna en cada fila del reporte.
  const nombreLibro = libro.getName();
  
  // ─── 1. Obtener o crear la hoja de resumen "Exogena" ───
  let hojaResumen = libro.getSheetByName("Exogena");
  if (!hojaResumen) {
    // Si no existe, la crea vacía.
    hojaResumen = libro.insertSheet("Exogena");
  } else {
    // Si ya existe, la limpia para evitar mezclar datos de meses anteriores.
    hojaResumen.clear();
  }

  // Array maestro donde se acumula toda la información antes de escribir.
  // Usar un solo setValues() al final es ~50x más rápido que appendRow() en un loop.
  const datosConsolidados = [];

  // ─── 2. Iterar por las posibles hojas de días del mes (1 al 31) ───
  for (let dia = 1; dia <= 31; dia++) {
    // Buscar la hoja cuyo nombre sea el número del día (ej: "1", "15", "31")
    const hojaDia = libro.getSheetByName(dia.toString());
    
    // Si la hoja no existe (ej: febrero no tiene hoja "30"), salta al siguiente día.
    if (!hojaDia) continue;

    // Leer el valor de producción del día (celda B11 de cada hoja diaria).
    const produccion = hojaDia.getRange("B11").getValue();
    
    // Solo procesar si hay producción mayor a cero (evita días vacíos o sin actividad).
    if (produccion > 0) {
      const ultimaFila = hojaDia.getLastRow();
      
      // Si no hay datos desde la fila 11 hacia abajo, salta.
      if (ultimaFila < 11) continue;

      // Leer las columnas E a I (5 columnas) desde la fila 11 hasta la última con datos.
      // Columna E = 5, se leen 5 columnas (E, F, G, H, I) y (ultimaFila - 10) filas.
      const rangoCampos = hojaDia.getRange(11, 5, ultimaFila - 10, 5);
      const datosDia = rangoCampos.getValues();

      // Filtrar filas vacías y formatear cada registro válido.
      for (const fila of datosDia) {
        // Solo incluir si la columna E (índice 0) tiene contenido válido.
        if (fila[0] !== undefined && fila[0] !== null && fila[0].toString().trim() !== "") {
          // Estructura de cada fila: [Día, NombreLibro, ColE, ColF, ColG, ColH, ColI]
          datosConsolidados.push([dia, nombreLibro, ...fila]);
        }
      }
    }
  }

  // ─── 3. Escribir todos los datos recolectados DE UNA SOLA VEZ ───
  if (datosConsolidados.length > 0) {
    const numFilas = datosConsolidados.length;
    const numCols = datosConsolidados[0].length;
    
    // setValues() escribe todo el bloque de datos en una sola operación al servidor.
    hojaResumen
      .getRange(1, 1, numFilas, numCols)
      .setValues(datosConsolidados);
  }
}
