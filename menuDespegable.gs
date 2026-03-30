function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Imperio")
    .addItem("Crear cliente", "crearHTMLCliente")
    .addItem("Gestionar cliente", "eliminarHTMLCliente")
    .addItem("Agendar", "agendarHTMLCliente")
    .addToUi();
}

function eliminarHTMLCliente() {
  const html =
    HtmlService.createHtmlOutputFromFile("eliminar").setTitle("Neto gestión");
  SpreadsheetApp.getUi().showSidebar(html);
}

function crearHTMLCliente() {
  const html = HtmlService.createHtmlOutputFromFile("crear")
    .setHeight(650)
    .setWidth(650);
  SpreadsheetApp.getUi().showModalDialog(html, "Neto registro");
}

function agendarHTMLCliente() {
  const html =
    HtmlService.createHtmlOutputFromFile("agendar").setTitle("Neto agenda");
  SpreadsheetApp.getUi().showSidebar(html);
}
