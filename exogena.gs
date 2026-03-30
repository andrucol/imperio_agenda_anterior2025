function exogena() {
  const libro = SpreadsheetApp.getActive();
  const hojaResumen =
    libro.getSheetByName("Exogena") || libro.insertSheet("Exogena");
  hojaResumen.clear();
  let nombre = libro.getName();
  for (let i = 1; i < 32; i++) {
    let hoja = libro.getSheetByName(i.toString());
    let produccion = hoja.getRange("B11").getValue();
    if (produccion > 0) {
      let datos = hoja.getRange("E11:I").getValues();
      let datosFiltrados = datos
        .filter((fila) => {
          return fila[0].toString().trim() !== "";
        })
        .map((fila) => [i, nombre, ...fila]);
      hojaResumen
        .getRange(
          hojaResumen.getLastRow() + 1,
          1,
          datosFiltrados.length,
          datosFiltrados[0].length,
        )
        .setValues(datosFiltrados);
    }
  }
}
