/**
 * AgendamientoWeb.gs
 * Lógica para el formulario HTML de agendamiento.
 */

const ID_BASE_DATOS_WEB = "1_f4uTfUENXCGFdd0eUnu5jZQJTiJOC2dGh_HWuMz9Bs";
const NOMBRE_HOJA_BD_WEB = "Reporte";

const HORAS_LABORALES_WEB = [
  "8 AM",
  "8:30 AM",
  "9 AM",
  "9:30 AM",
  "10 AM",
  "10:30 AM",
  "11 AM",
  "11:30 AM",
  "12 PM",
  "12:30 PM",
  "1 PM",
  "1:30 PM",
  "2 PM",
  "2:30 PM",
  "3 PM",
  "3:30 PM",
  "4 PM",
  "4:30 PM",
  "5 PM",
  "5:30 PM",
  "6 PM",
  "6:30 PM",
  "7 PM",
];

/* ==========================================
   1. LISTAS PARA DROPDOWNS (WEB)
   ========================================== */

function getListaProfesionales() {
  const hoja = SpreadsheetApp.getActive().getSheetByName("Info");
  if (!hoja) return [];
  const ultFila = hoja.getLastRow();
  // Validación: Si hay menos de 3 filas (headers + vacío), retorna array vacío para evitar error en getRange
  if (ultFila < 3) return [];

  // Columna E (5), desde fila 3
  return hoja
    .getRange(3, 5, ultFila - 2, 1)
    .getValues()
    .flat()
    .filter((item) => item !== "");
}

function getListaClientes() {
  const hoja = SpreadsheetApp.getActive().getSheetByName("Clientes");
  if (!hoja) return [];
  const ultFila = hoja.getLastRow();
  if (ultFila < 2) return [];

  // Columna B (2), desde fila 2
  return hoja
    .getRange(2, 2, ultFila - 1, 1)
    .getValues()
    .flat()
    .filter(String);
}

function getListaProcedimientos() {
  const hoja = SpreadsheetApp.getActive().getSheetByName("Info");
  if (!hoja) return [];
  const ultFila = hoja.getLastRow();
  if (ultFila < 1) return []; // Validación extra

  // Columna L (12), desde fila 1
  return hoja
    .getRange(1, 12, ultFila, 1)
    .getValues()
    .flat()
    .filter((item) => item !== "");
}

/* ==========================================
   2. VERIFICAR DISPONIBILIDAD (WEB)
   ========================================== */

function verAgenda(profesional, dia, mes, anio) {
  let bdSheet;
  try {
    bdSheet =
      SpreadsheetApp.openById(ID_BASE_DATOS_WEB).getSheetByName(
        NOMBRE_HOJA_BD_WEB,
      );
  } catch (e) {
    // Fallback por si falla el ID, intenta usar la hoja activa
    bdSheet = SpreadsheetApp.getActive().getSheetByName(NOMBRE_HOJA_BD_WEB);
  }

  // Si no existe la hoja, devolvemos todo libre
  if (!bdSheet)
    return { horasDisponibles: HORAS_LABORALES_WEB, citasOcupadas: [] };

  const ultimaFila = bdSheet.getLastRow();
  let citasOcupadas = [];

  // Solo leemos si hay datos (fila > 1 asumiendo fila 1 es encabezado)
  if (ultimaFila > 1) {
    // Indices (0-based): [1]Dia, [2]Mes, [3]Año, [4]Profesional, [6]Hora, [8]Proceso, [9]Obs
    const datos = bdSheet.getRange(2, 1, ultimaFila - 1, 15).getValues();

    // Filtramos
    const agendaFiltrada = datos.filter((fila) => {
      // Usamos == para permitir comparación laxa (string vs number)
      return (
        fila[4] == profesional &&
        fila[1] == dia &&
        fila[2] == mes &&
        fila[3] == anio
      );
    });

    // Mapeamos resultado para enviar al HTML
    citasOcupadas = agendaFiltrada.map((fila) => {
      return {
        hora: fila[6],
        cliente: fila[7], // <--- NUEVO: Índice 7 es el Celular/Cliente
        proceso: fila[8],
        preferencia: fila[10], // <--- NUEVO: Índice 10 es la Preferencia
        observaciones: fila[9],
      };
    });
  }

  // Calculamos libres
  const listaHorasOcupadas = citasOcupadas.map((c) => c.hora);
  const horasDisponibles = HORAS_LABORALES_WEB.filter(
    (hora) => !listaHorasOcupadas.includes(hora),
  );

  return {
    horasDisponibles: horasDisponibles,
    citasOcupadas: citasOcupadas,
  };
}

/* ==========================================
   3. GUARDAR CITA (WEB) - MODIFICADO
   ========================================== */

function guardarCita(form) {
  // 1. Validar Cliente
  if (!validarClienteWeb(form.celular)) {
    throw new Error("El celular " + form.celular + " no está registrado.");
  }

  // 2. Abrir Base de Datos
  let bdBook;
  try {
    bdBook = SpreadsheetApp.openById(ID_BASE_DATOS_WEB);
  } catch (e) {
    bdBook = SpreadsheetApp.getActive();
  }

  const bdSheet = bdBook.getSheetByName(NOMBRE_HOJA_BD_WEB);
  const ultimaFila = bdSheet.getLastRow() + 1;

  // 3. Generar Datos
  const fechaInicio = new Date(2025, 2, 1);
  const fechaActual = new Date();
  const id = Math.floor((fechaActual.getTime() - fechaInicio.getTime()) / 1000);
  const correo = Session.getActiveUser().getEmail();
  const estado = "Agendada";

  // 4. Guardar en Hoja
  const filaNueva = [
    id, // 0
    form.dia, // 1
    form.mes, // 2
    form.anio, // 3
    form.profesional, // 4
    form.area, // 5
    form.hora, // 6
    form.celular, // 7
    form.procedimiento, // 8
    form.observacion, // 9
    form.preferencia, // 10
    estado, // 11
    correo, // 12
    fechaActual, // 13
  ];

  bdSheet.getRange(ultimaFila, 1, 1, filaNueva.length).setValues([filaNueva]);

  // 5. RETORNAR OBJETO PARA WHATSAPP (No abrimos el link aquí, devolvemos los datos)
  return {
    resultado: "OK",
    celular: form.celular,
    mensaje: `Hola, te confirmamos tu cita de *${form.area}* para el día *${form.dia} de ${form.mes}* a las *${form.hora}*. ¿Confirmas asistencia?`,
  };
}

/* --- Helpers Internos --- */

function validarClienteWeb(celular) {
  const hoja = SpreadsheetApp.getActive().getSheetByName("Clientes");
  if (!hoja) return false;
  // Validación de seguridad para hoja vacía
  if (hoja.getLastRow() < 2) return false;

  const datos = hoja
    .getRange("B2:B" + hoja.getLastRow())
    .getValues()
    .flat();
  // Comparación estricta numérica para evitar falsos negativos por formato texto/número
  return datos.some((c) => Number(c) === Number(celular));
}

function doGet(e) {
  // 1. Detectamos qué página quiere ver el usuario
  // Si no hay parámetro, por defecto mostramos 'agendar'
  var ruta = e.parameter.v || "agendar";

  var template;
  var titulo;

  // 2. Elegimos el archivo HTML correcto
  switch (ruta) {
    case "crear":
      template = HtmlService.createTemplateFromFile("crear");
      titulo = "Crear Nuevo Registro";
      break;

    case "eliminar":
      template = HtmlService.createTemplateFromFile("eliminar");
      titulo = "Eliminar Registro";
      break;

    case "agendar":
    default: // 'default' atrapa cualquier error o ruta desconocida
      template = HtmlService.createTemplateFromFile("agendar");
      titulo = "Agendar Cita";
      break;
  }

  // 3. IMPORTANTE: Obtenemos la URL pública del script
  // y se la pasamos a la plantilla para usarla en los botones
  template.pubUrl = ScriptApp.getService().getUrl();

  // 4. Renderizamos la página
  return template
    .evaluate()
    .setTitle(titulo)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
