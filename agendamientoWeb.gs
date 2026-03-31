/**
 * ============================================
 * agendamientoWeb.gs - LÓGICA DE LA WEB APP PÚBLICA
 * ============================================
 *
 * Este archivo maneja la aplicación web que se despliega como
 * "Web App" de Google Apps Script. Es la interfaz que usan los
 * usuarios EXTERNOS (clientes o recepcionistas desde celular)
 * para agendar, crear y eliminar clientes vía navegador.
 *
 * FUNCIONES PRINCIPALES:
 *   - doGet(): Punto de entrada de la web (decide qué página mostrar)
 *   - getListaXxx(): Proveen datos a los dropdowns del formulario HTML
 *   - verAgenda(): Consulta disponibilidad de horarios (usada por agendar.html)
 *   - guardarCita(): Guarda una cita nueva (con LockService)
 *   - validarClienteWeb(): Verifica que el celular esté registrado
 *
 * NOTA: Este archivo tiene su propia copia de HORAS_LABORALES_WEB y los
 * IDs de BD porque la web app se ejecuta en un contexto separado del
 * libro de Sheets. En una futura mejora se podrían unificar con baseDatos.gs.
 */

// ─── Constantes de la Web App ───
// ID del libro de Google Sheets con los datos de citas
const ID_BASE_DATOS_WEB = "1_f4uTfUENXCGFdd0eUnu5jZQJTiJOC2dGh_HWuMz9Bs";
// Nombre de la pestaña donde se guardan las citas
const NOMBRE_HOJA_BD_WEB = "Reporte";

// Horario laboral completo (mismo que en baseDatos.gs, duplicado aquí
// porque la web app puede ejecutarse sin acceso directo a baseDatos.gs)
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


// ╔══════════════════════════════════════════════╗
// ║   1. LISTAS PARA DROPDOWNS (WEB)            ║
// ╚══════════════════════════════════════════════╝

/**
 * getListaProfesionales()
 * Lee la lista de profesionales desde la hoja "Info" del libro activo
 * (columna E, desde fila 3) y la retorna como array para llenar
 * el dropdown de profesionales en agendar.html.
 * @returns {Array<string>} Nombres de profesionales (sin vacíos).
 */
function getListaProfesionales() {
  const hoja = SpreadsheetApp.getActive().getSheetByName("Info");
  if (!hoja) return [];
  const ultFila = hoja.getLastRow();
  // Si hay menos de 3 filas (headers + mínimo 1 dato), retorna vacío
  if (ultFila < 3) return [];

  // Columna E = columna 5, desde fila 3 hasta la última con datos
  return hoja
    .getRange(3, 5, ultFila - 2, 1)
    .getValues()
    .flat()                             // Aplana [[val1],[val2]] a [val1, val2]
    .filter((item) => item !== "");     // Elimina celdas vacías
}

/**
 * getListaClientes()
 * Lee la lista de celulares de clientes desde la hoja "Clientes"
 * (columna B, desde fila 2) para el buscador de clientes en agendar.html.
 * @returns {Array<string>} Lista de celulares registrados.
 */
function getListaClientes() {
  const hoja = SpreadsheetApp.getActive().getSheetByName("Clientes");
  if (!hoja) return [];
  const ultFila = hoja.getLastRow();
  if (ultFila < 2) return [];

  // Columna B = columna 2, desde fila 2
  return hoja
    .getRange(2, 2, ultFila - 1, 1)
    .getValues()
    .flat()
    .filter(String);                    // filter(String) equivale a eliminar vacíos/nulls
}

/**
 * getListaProcedimientos()
 * Lee la lista de procedimientos/servicios disponibles
 * desde la hoja "Info" (columna L, desde fila 1).
 * @returns {Array<string>} Lista de procedimientos (ej: "Manicure Gel", "Pestañas pelo a pelo").
 */
function getListaProcedimientos() {
  const hoja = SpreadsheetApp.getActive().getSheetByName("Info");
  if (!hoja) return [];
  const ultFila = hoja.getLastRow();
  if (ultFila < 1) return [];

  // Columna L = columna 12, desde fila 1
  return hoja
    .getRange(1, 12, ultFila, 1)
    .getValues()
    .flat()
    .filter((item) => item !== "");
}


// ╔══════════════════════════════════════════════╗
// ║   2. VERIFICAR DISPONIBILIDAD (WEB)         ║
// ╚══════════════════════════════════════════════╝

/**
 * verAgenda(profesional, dia, mes, anio)
 * Consulta la base de datos de citas y retorna un objeto con:
 *   - horasDisponibles: array de horas libres para esa fecha/profesional
 *   - citasOcupadas: array de objetos con los datos de cada cita existente
 *
 * Esta función es llamada desde agendar.html cuando el usuario presiona
 * "Ver Disponibilidad". Es diferente a la verAgenda(numAgenda) de codigo.gs
 * porque esta recibe los parámetros directamente y retorna datos al HTML,
 * mientras que la otra lee/escribe directamente en las celdas de Sheets.
 *
 * @param {string} profesional - Nombre del profesional a consultar.
 * @param {string|number} dia - Día del mes.
 * @param {string} mes - Nombre del mes (ej: "marzo").
 * @param {string|number} anio - Año (ej: 2025).
 * @returns {Object} { horasDisponibles: string[], citasOcupadas: Object[] }
 */
function verAgenda(profesional, dia, mes, anio) {
  let bdSheet;
  try {
    // Intentar abrir la BD externa por su ID
    bdSheet =
      SpreadsheetApp.openById(ID_BASE_DATOS_WEB).getSheetByName(
        NOMBRE_HOJA_BD_WEB,
      );
  } catch (e) {
    // Fallback: si falla el ID (ej: permisos), usar la hoja activa
    bdSheet = SpreadsheetApp.getActive().getSheetByName(NOMBRE_HOJA_BD_WEB);
  }

  // Si no existe la hoja, devolver todo como libre
  if (!bdSheet)
    return { horasDisponibles: HORAS_LABORALES_WEB, citasOcupadas: [] };

  const ultimaFila = bdSheet.getLastRow();
  let citasOcupadas = [];

  // Solo procesar si hay datos (fila > 1 = hay algo además del encabezado)
  if (ultimaFila > 1) {
    // Leer todas las filas de datos (15 columnas: A-O)
    // Índices 0-based del array:
    //   [0]=ID, [1]=Día, [2]=Mes, [3]=Año, [4]=Profesional, [5]=Área,
    //   [6]=Hora, [7]=Celular, [8]=Procedimiento, [9]=Observación,
    //   [10]=Preferencia, [11]=Estado, [12]=Correo, [13]=Fecha, [14]=extra
    const datos = bdSheet.getRange(2, 1, ultimaFila - 1, 15).getValues();

    // Filtrar por los criterios del usuario (== para comparación laxa texto/número)
    const agendaFiltrada = datos.filter((fila) => {
      return (
        fila[4] == profesional &&
        fila[1] == dia &&
        fila[2] == mes &&
        fila[3] == anio
      );
    });

    // Mapear a objetos legibles para el frontend HTML
    citasOcupadas = agendaFiltrada.map((fila) => {
      return {
        hora: fila[6],             // Hora de la cita
        cliente: fila[7],          // Celular del cliente
        proceso: fila[8],          // Procedimiento solicitado
        preferencia: fila[10],     // SI/NO preferencia
        observaciones: fila[9],    // Notas adicionales
      };
    });
  }

  // Calcular horas libres = todas las horas - las que ya están ocupadas
  const listaHorasOcupadas = citasOcupadas.map((c) => c.hora);
  const horasDisponibles = HORAS_LABORALES_WEB.filter(
    (hora) => !listaHorasOcupadas.includes(hora),
  );

  return {
    horasDisponibles: horasDisponibles,
    citasOcupadas: citasOcupadas,
  };
}


// ╔══════════════════════════════════════════════╗
// ║   3. GUARDAR CITA (WEB)                     ║
// ╚══════════════════════════════════════════════╝

/**
 * guardarCita(form)
 * Guarda una nueva cita agendada desde la web app en la base de datos.
 * Usa LockService para evitar que dos usuarios web agenden al mismo tiempo
 * y se sobreescriban datos en la misma fila.
 *
 * FLUJO:
 *   1. Validar que el celular esté registrado en la base de clientes
 *   2. Abrir la BD de agenda
 *   3. Adquirir candado (Lock) para escritura segura
 *   4. Generar ID único y timestamp
 *   5. Escribir la fila de la cita
 *   6. Retornar datos al HTML para generar el enlace de WhatsApp
 *
 * @param {Object} form - Datos del formulario HTML:
 *   { dia, mes, anio, profesional, area, hora, celular, procedimiento, observacion, preferencia }
 * @returns {Object} { resultado: "OK", celular, mensaje } para WhatsApp.
 * @throws {Error} Si el celular no está registrado.
 */
function guardarCita(form) {
  // Paso 1: Validar que el cliente exista antes de continuar
  if (!validarClienteWeb(form.celular)) {
    throw new Error("El celular " + form.celular + " no está registrado.");
  }

  // Paso 2: Abrir la base de datos de agenda
  let bdBook;
  try {
    bdBook = SpreadsheetApp.openById(ID_BASE_DATOS_WEB);
  } catch (e) {
    // Fallback a la hoja activa si falla el ID
    bdBook = SpreadsheetApp.getActive();
  }

  const bdSheet = bdBook.getSheetByName(NOMBRE_HOJA_BD_WEB);

  // Paso 3: Adquirir candado para escritura segura multi-usuario
  const lock = LockService.getScriptLock();
  lock.tryLock(10000); // Esperar hasta 10 segundos
  
  try {
    // La lectura de la última fila se hace DENTRO del lock para evitar
    // que dos usuarios lean la misma "última fila" simultáneamente
    const ultimaFila = bdSheet.getLastRow() + 1;

    // Paso 4: Generar ID único y datos de auditoría
    const fechaInicio = new Date(2025, 2, 1);   // Referencia: 1 de marzo 2025
    const fechaActual = new Date();
    const id = Math.floor((fechaActual.getTime() - fechaInicio.getTime()) / 1000);
    const correo = Session.getActiveUser().getEmail();
    const estado = "Agendada";

    // Paso 5: Armar la fila completa (14 columnas: A-N)
    const filaNueva = [
      id,                  // A: ID único
      form.dia,            // B: Día
      form.mes,            // C: Mes
      form.anio,           // D: Año
      form.profesional,    // E: Nombre del profesional
      form.area,           // F: Área/especialidad
      form.hora,           // G: Hora de la cita
      form.celular,        // H: Celular del cliente
      form.procedimiento,  // I: Procedimiento solicitado
      form.observacion,    // J: Observaciones
      form.preferencia,    // K: Preferencia (SI/NO)
      estado,              // L: Estado de la cita
      correo,              // M: Email del creador
      fechaActual,         // N: Fecha y hora de creación
    ];

    // Escribir la fila en la BD
    bdSheet.getRange(ultimaFila, 1, 1, filaNueva.length).setValues([filaNueva]);

    // Paso 6: Retornar datos al frontend para el enlace de WhatsApp
    return {
      resultado: "OK",
      celular: form.celular,
      mensaje: `Hola, te confirmamos tu cita de *${form.area}* para el día *${form.dia} de ${form.mes}* a las *${form.hora}*. ¿Confirmas asistencia?`,
    };
  } finally {
    // SIEMPRE liberar el candado, incluso si hubo un error
    lock.releaseLock();
  }
}


// ╔══════════════════════════════════════════════╗
// ║   HELPERS INTERNOS                          ║
// ╚══════════════════════════════════════════════╝

/**
 * validarClienteWeb(celular)
 * Verifica si un número de celular existe en la hoja "Clientes" (columna B).
 * Se usa antes de guardar una cita para asegurar que el cliente está registrado.
 * Usa comparación numérica estricta para evitar falsos negativos
 * cuando el dato en Sheets es texto y el parámetro es número (o viceversa).
 *
 * @param {string|number} celular - Número de celular a validar.
 * @returns {boolean} true si el celular existe en la base, false si no.
 */
function validarClienteWeb(celular) {
  const hoja = SpreadsheetApp.getActive().getSheetByName("Clientes");
  if (!hoja) return false;
  // Si la hoja tiene solo encabezado (fila 1), no hay clientes
  if (hoja.getLastRow() < 2) return false;

  // Leer todos los celulares de la columna B
  const datos = hoja
    .getRange("B2:B" + hoja.getLastRow())
    .getValues()
    .flat();

  // Comparar numéricamente: Number("3102875348") === Number(3102875348)
  return datos.some((c) => Number(c) === Number(celular));
}


// ╔══════════════════════════════════════════════╗
// ║   PUNTO DE ENTRADA DE LA WEB APP            ║
// ╚══════════════════════════════════════════════╝

/**
 * doGet(e)
 * Función especial de Google Apps Script que se ejecuta automáticamente
 * cuando alguien accede a la URL pública de la web app.
 *
 * ROUTING:
 *   La URL puede incluir un parámetro "v" para elegir la página:
 *     ?v=agendar   → Formulario de agendamiento (default)
 *     ?v=crear     → Formulario de registro de cliente
 *     ?v=eliminar  → Panel de gestión/eliminación de cliente
 *
 * IMPORTANTE: Pasa la URL pública del script (template.pubUrl) a la plantilla
 * para que los botones de navegación entre páginas funcionen correctamente.
 *
 * @param {Object} e - Evento HTTP con parámetros de la URL.
 * @returns {HtmlOutput} Página HTML renderizada para el navegador.
 */
function doGet(e) {
  // Leer el parámetro "v" de la URL; por defecto mostrar "agendar"
  var ruta = e.parameter.v || "agendar";

  var template;
  var titulo;

  // Seleccionar el archivo HTML y el título según la ruta
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
    default: // Cualquier ruta desconocida cae aquí también
      template = HtmlService.createTemplateFromFile("agendar");
      titulo = "Agendar Cita";
      break;
  }

  // Inyectar la URL pública del script en la plantilla HTML
  // para que los links de navegación (Agendar/Crear/Eliminar) funcionen
  template.pubUrl = ScriptApp.getService().getUrl();

  // Renderizar el template con evaluate() y configurar metadatos
  return template
    .evaluate()
    .setTitle(titulo)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")    // Responsive
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);    // Permitir embeber en iframes
}
