/**
 * ============================================
 * menuDespegable.gs - MENÚ PERSONALIZADO EN GOOGLE SHEETS
 * ============================================
 *
 * Crea un menú llamado "Imperio" en la barra superior de Google Sheets.
 * Este menú aparece automáticamente cada vez que un usuario abre la hoja de cálculo
 * gracias al trigger reservado onOpen() de Google Apps Script.
 *
 * Desde aquí se abren las ventanas HTML para:
 *   - Crear clientes nuevos (ventana modal/popup)
 *   - Gestionar/eliminar clientes (barra lateral derecha)
 *   - Agendar citas desde el panel interno (barra lateral derecha)
 */

/**
 * onOpen() - Trigger automático de Google Apps Script.
 * Se ejecuta cada vez que cualquier usuario abre la hoja de cálculo.
 * Genera el menú "Imperio" con 3 opciones que invocan funciones de este archivo.
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Imperio")
    .addItem("Crear cliente", "crearHTMLCliente")       // Abre popup de registro
    .addItem("Gestionar cliente", "eliminarHTMLCliente") // Abre sidebar de gestión/eliminación
    .addItem("Agendar", "agendarHTMLCliente")            // Abre sidebar de agendamiento
    .addToUi();
}

/**
 * eliminarHTMLCliente() - Abre la barra lateral (sidebar) con la vista de gestión de clientes.
 * Carga el archivo "eliminar.html" y lo muestra en un panel lateral a la derecha del Sheets.
 */
function eliminarHTMLCliente() {
  const html =
    HtmlService.createHtmlOutputFromFile("eliminar").setTitle("Neto gestión");
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * crearHTMLCliente() - Abre una ventana modal (popup centrado) con el formulario de registro.
 * Carga el archivo "crear.html" y lo muestra como un diálogo flotante de 650x650px.
 */
function crearHTMLCliente() {
  const html = HtmlService.createHtmlOutputFromFile("crear")
    .setHeight(650)
    .setWidth(650);
  SpreadsheetApp.getUi().showModalDialog(html, "Neto registro");
}

/**
 * agendarHTMLCliente() - Abre la barra lateral (sidebar) con la vista de agendamiento.
 * Carga el archivo "agendar.html" y lo muestra en un panel lateral a la derecha del Sheets.
 */
function agendarHTMLCliente() {
  const html =
    HtmlService.createHtmlOutputFromFile("agendar").setTitle("Neto agenda");
  SpreadsheetApp.getUi().showSidebar(html);
}
