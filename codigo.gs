/**
 * ============================================
 * codigo.gs - LÓGICA PRINCIPAL DEL SISTEMA DE AGENDAMIENTO
 * ============================================
 *
 * Contiene las funciones de:
 *   1. GESTIÓN DE CLIENTES: Consultar, crear, eliminar y buscar clientes.
 *   2. AGENDAS (Agenda_1, Agenda_2): Ver disponibilidad, agendar citas, bloquear días.
 *   3. VISTAS DE DÍA (Dia_1, Dia_2): Actualizar, cambiar estados, enviar WhatsApp.
 *   4. VISTA MENSUAL (Mes): Reporte mensual por profesional.
 *   5. WRAPPERS: Funciones "puente" para los botones de Google Sheets.
 *
 * DEPENDENCIAS:
 *   - baseDatos.gs (IDs de hojas, horarios, colores)
 *   - agendamientoWeb.gs (lógica de la web app pública)
 *
 * CONCURRENCIA:
 *   Las funciones de escritura usan LockService.getScriptLock() para evitar
 *   que dos usuarios escriban al mismo tiempo y se sobreescriban datos.
 */


// ╔══════════════════════════════════════════════╗
// ║   SECCIÓN 1: GESTIÓN DE CLIENTES            ║
// ╚══════════════════════════════════════════════╝

/**
 * obtenerNumerosClientes()
 * Devuelve un array con todos los números de celular de la base de datos de clientes.
 * Se usa en el frontend (eliminar.html) para llenar el selector con buscador (Select2).
 * @returns {Array<string>} Lista plana de celulares, sin vacíos.
 */
function obtenerNumerosClientes() {
  let libro = SpreadsheetApp.openById(idBaseDeDatosClientes);
  let base = libro.getSheetByName(hojaBaseDeDatosClientes);
  // Lee toda la columna A desde la fila 2 (la fila 1 es el encabezado)
  let datos = base.getRange("A2:A").getValues();
  // flat() convierte [[val1],[val2]] en [val1, val2], filter(String) elimina vacíos
  return datos.flat().filter(String);
}

/**
 * obtenerDatosClientes(numero)
 * Busca un cliente por su número de celular usando BÚSQUEDA BINARIA.
 * La columna A debe estar ORDENADA numéricamente para que esto funcione.
 * Retorna un string formateado con los datos del cliente o un mensaje de error.
 *
 * @param {number} numero - Número de celular a buscar.
 * @returns {string} Datos del cliente en formato legible, o mensaje de no encontrado.
 */
function obtenerDatosClientes(numero) {
  try {
    const libro = SpreadsheetApp.openById(idBaseDeDatosClientes);
    const base = libro.getSheetByName(hojaBaseDeDatosClientes);

    // Búsqueda binaria: inicio en fila 2 (después del header), fin en la última fila.
    let inicio = 2, final = base.getLastRow();
    while (inicio <= final) {
      let mitad = Math.floor((inicio + final) / 2);
      let posicionDato = Number(base.getRange(`A${mitad}`).getValue());

      if (numero === posicionDato) {
        // ¡Encontrado! Leer las 6 columnas de esa fila (A-F)
        let fila = base.getRange(mitad, 1, 1, 6).getValues()[0];

        // Formatear la fecha de cumpleaños (columna F, índice 5)
        let fecha = new Date(fila[5]);
        let dia = fecha.getDate().toString().padStart(2, "0");
        let mes = (fecha.getMonth() + 1).toString().padStart(2, "0");
        let cumpleañosFormateado = `${dia}/${mes}`;

        // Retornar string multilinea con la info del cliente
        return `Nombre: ${fila[1]}\nTipo Doc.: ${fila[2]}\nNúmero Doc.: ${fila[3]}\nCorreo: ${fila[4]}\nCumpleaños: ${cumpleañosFormateado}`;
      } else if (numero < posicionDato) {
        // El número buscado es menor, mover el final hacia la izquierda.
        final = mitad - 1;
      } else {
        // El número buscado es mayor, mover el inicio hacia la derecha.
        inicio = mitad + 1;
      }
    }
    return `Cliente con número ${numero} no encontrado.`;
  } catch (error) {
    Logger.log("Error en obtenerDatosClientes: " + error.message);
    return "Error al buscar el cliente.";
  }
}

/**
 * eliminarDatosClientes(numero)
 * Busca un cliente por celular (búsqueda binaria), lo mueve a la hoja "Eliminados"
 * como respaldo, y luego borra la fila de la hoja principal.
 *
 * @param {number} numero - Número de celular del cliente a eliminar.
 * @returns {string} Mensaje de éxito o error.
 */
function eliminarDatosClientes(numero) {
  try {
    const libro = SpreadsheetApp.openById(idBaseDeDatosClientes);
    const base = libro.getSheetByName(hojaBaseDeDatosClientes);
    const hojaEliminados = libro.getSheetByName("Eliminados");

    let inicio = 2, final = base.getLastRow();
    while (inicio <= final) {
      let mitad = Math.floor((inicio + final) / 2);
      let posicionDato = Number(base.getRange(`A${mitad}`).getValue());

      if (numero === posicionDato) {
        // Leer 9 columnas (A-I) de la fila encontrada
        let fila = base.getRange(mitad, 1, 1, 9).getValues()[0];
        // Guardar en "Eliminados" con la fecha de eliminación al inicio
        hojaEliminados.appendRow([new Date(), ...fila.slice(0, 9)]);
        // Borrar la fila original de la base
        base.deleteRow(mitad);
        return `Cliente con número ${numero} eliminado correctamente.`;
      } else if (numero < posicionDato) {
        final = mitad - 1;
      } else {
        inicio = mitad + 1;
      }
    }
    return `Cliente con número ${numero} no encontrado.`;
  } catch (error) {
    return "Error al eliminar el cliente.";
  }
}

/**
 * apiWhatsApp(celular, area, dia, hora)
 * Genera un enlace de WhatsApp con un mensaje predefinido de confirmación de cita
 * y lo muestra al usuario en una ventana emergente dentro de Google Sheets.
 * Al hacer clic se abre WhatsApp Web/App con el texto prellenado.
 *
 * @param {string|number} celular - Número de celular del cliente (sin código de país).
 * @param {string} area - Tipo de servicio (ej: MANICURISTA, LASHISTA).
 * @param {string|number} dia - Día de la cita.
 * @param {string} hora - Hora de la cita (ej: "10 AM").
 */
function apiWhatsApp(celular, area, dia, hora) {
  // Construir la URL de la API de WhatsApp con el mensaje codificado
  let mensaje = `https://api.whatsapp.com/send?phone=+57${celular}&text=¡Hola%20con%20mucho%20amor!%20😊%20Queremos%20recordarte%20que%20tu%20cita%20de%20*${area}*%20ha%20sido%20generada%20para%20el%20día%20*${dia}*%20a%20las%20*${hora}*.%20¡Te%20esperamos!%20💕`;
  // Crear un mini HTML con el enlace para que el usuario haga clic
  let html = HtmlService.createHtmlOutput(`<html><head><style>a { font-weight: 800; font-size: 18px; color: Lime; } a:hover { color: gray; }</style></head><body><a href="${mensaje}" target="_blank">¡Enviar confirmacion!</a></body></html>`).setWidth(300).setHeight(300);
  // Mostrar como diálogo modal en Google Sheets
  SpreadsheetApp.getUi().showModalDialog(html, "Enlace WhatsApp");
}

/**
 * busquedaBinaria(celular)
 * Busca un número de celular en la base de datos usando búsqueda binaria.
 * Retorna [true, fila] si existe, [false, null] si no.
 * Se usa desde crear.html para verificar si el cliente ya está registrado
 * ANTES de intentar guardarlo.
 *
 * @param {number} celular - Número de celular a verificar.
 * @returns {Array} [boolean existe, number|null filaEncontrada]
 */
function busquedaBinaria(celular) {
  try {
    let base = SpreadsheetApp.openById(idBaseDeDatosClientes).getSheetByName(hojaBaseDeDatosClientes);
    let inicio = 1, final = base.getLastRow();
    while (inicio <= final) {
      let mitad = Math.floor((inicio + final) / 2);
      let posicionDato = Number(base.getRange(`A${mitad}`).getValue());
      if (celular === posicionDato) return [true, mitad];
      else if (celular < posicionDato) final = mitad - 1;
      else inicio = mitad + 1;
    }
    return [false, null];
  } catch (error) {
    return [false, null];
  }
}

/**
 * guardarCliente(data)
 * Inserta un nuevo cliente en la base de datos externa.
 * Usa LockService para evitar que dos usuarios creen clientes al mismo tiempo
 * y se pierda uno de los registros.
 * Después de insertar, reordena toda la tabla por la columna A (celular)
 * para mantener la búsqueda binaria funcional.
 *
 * @param {Object} data - Objeto con los campos del formulario:
 *   { customerNumber, name, docType, docNumber, email, birthDate }
 */
function guardarCliente(data) {
  // Adquirir candado de script para evitar escrituras simultáneas
  let lock = LockService.getScriptLock();
  lock.tryLock(10000); // Esperar hasta 10 segundos si otro usuario está escribiendo

  try {
    const hoja = SpreadsheetApp.openById(idBaseDeDatosClientes).getSheetByName(hojaBaseDeDatosClientes);

    let fechaActual = new Date();
    // Generar un ID único basado en segundos transcurridos desde el 1 de enero 2025
    let id = Math.floor((fechaActual.getTime() - new Date(2025, 0, 1).getTime()) / 1000);

    // Insertar nueva fila con: [Celular, Nombre, TipoDoc, NumDoc, Email, Cumpleaños, ID, CorreoCreador, FechaCreación]
    hoja.appendRow([
      data.customerNumber,
      data.name.toUpperCase(),
      data.docType.toUpperCase(),
      data.docNumber,
      data.email.toUpperCase(),
      data.birthDate,
      id,
      Session.getActiveUser().getEmail().toUpperCase(),
      fechaActual
    ]);

    // Reordenar por columna A (celular) ascendente para que la búsqueda binaria siga funcionando
    hoja.getRange("A2:I").sort({ column: 1, ascending: true });
  } finally {
    // SIEMPRE liberar el candado, incluso si hubo error
    lock.releaseLock();
  }
}


// ╔══════════════════════════════════════════════╗
// ║   SECCIÓN 2: FUNCIONES DE AGENDA            ║
// ║   (Agenda_1, Agenda_2)                       ║
// ╚══════════════════════════════════════════════╝

/**
 * limpiarVer(numAgenda)
 * Restaura las celdas de la hoja "Agenda_X" a su estado inicial (blanco/vacío).
 * Se ejecuta antes de cargar datos nuevos para evitar que queden residuos.
 *
 * Celdas afectadas:
 *   G3  = Celular del cliente
 *   G6  = Procedimiento
 *   G9  = Hora seleccionada
 *   F15:G = Columnas de hora y celular (datos de citas existentes)
 *   I15:K = Columnas de procedimiento, observación y preferencia
 *   L3  = Checkbox de preferencia (se desmarca)
 *   K8  = Campo de observación (placeholder)
 *
 * @param {number} numAgenda - Número de agenda (1 o 2).
 */
function limpiarVer(numAgenda) {
  let hoja = SpreadsheetApp.getActive().getSheetByName("Agenda_" + numAgenda);
  if (!hoja) return;
  hoja.getRange("G3").setValue("").setBackground("white");    // Limpiar celular
  hoja.getRange("G6").setValue("").setBackground("white");    // Limpiar procedimiento
  hoja.getRange("G9").setValue("").setBackground("white");    // Limpiar hora
  hoja.getRange("F15:G").setValue("").setBackground("white"); // Limpiar tabla de citas (hora + celular)
  hoja.getRange("I15:K").setValue("").setBackground("white"); // Limpiar tabla de citas (proc + obs + pref)
  hoja.getRange("L3").setValue(false).setBackground("white"); // Desmarcar checkbox de preferencia
  hoja.getRange("K8").setValue("Escribe una observacion...").setBackground("white"); // Placeholder de observación
}

/**
 * verAgenda(numAgenda)
 * Lee los filtros (profesional, día, mes, año) de la hoja "Agenda_X",
 * consulta la base de datos de agenda externa y muestra:
 *   - Las citas ya agendadas para esa fecha y profesional (en filas lavanda)
 *   - Las horas disponibles en un dropdown con validación de datos
 *
 * FLUJO:
 *   1. Limpiar datos previos con limpiarVer()
 *   2. Leer filtros del Sheets (C3=profesional, C8=día, C11=mes, C14=año)
 *   3. Filtrar la BD de agenda por esos criterios
 *   4. Ordenar las citas por hora usando HORAS_LABORALES como referencia
 *   5. Calcular horas libres = HORAS_LABORALES - horas ocupadas
 *   6. Crear dropdown en G9 con las horas libres
 *   7. Mostrar las citas existentes en las filas 15+
 *   8. Activar la bandera en la hoja "Info" (columna J) para habilitar el botón "Agendar"
 *
 * @param {number} numAgenda - Número de agenda (1 o 2).
 */
function verAgenda(numAgenda) {
  let ui = SpreadsheetApp.getUi();
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Agenda_" + numAgenda);
  let hojaInfo = libro.getSheetByName("Info");
  if (!hoja) return ui.alert("Hoja Agenda_" + numAgenda + " no encontrada.");

  // Paso 1: Limpiar residuos de consultas anteriores
  limpiarVer(numAgenda);

  // Paso 2: Leer los filtros que el usuario seleccionó en la hoja
  let profesional = hoja.getRange("C3").getValue();
  let dia = hoja.getRange("C8").getValue();
  let mes = hoja.getRange("C11").getValue();
  let anio = hoja.getRange("C14").getValue();

  if (profesional && dia && mes && anio) {
    // Paso 3: Traer TODA la base de datos de agenda y filtrar en memoria
    let baseDeDatosAgenda = SpreadsheetApp.openById(idBaseDeDatosAgenda).getSheetByName(hojaBaseDeDatosAgenda).getDataRange().getValues();
    let baseDeDatosAgendaFiltro = baseDeDatosAgenda.filter(fila => fila[4] === profesional && fila[1] === dia && fila[2] === mes && fila[3] === anio);

    if (baseDeDatosAgendaFiltro.length > 0) {
      // Paso 4: Ordenar las citas filtradas por hora (según el orden de HORAS_LABORALES)
      let ordenHoras = {};
      HORAS_LABORALES.forEach((hora, index) => ordenHoras[hora] = index);
      baseDeDatosAgendaFiltro.sort((a, b) => ordenHoras[a[6]] - ordenHoras[b[6]]);

      // Paso 5: Calcular las horas que NO están ocupadas
      let horasNoDisponibles = baseDeDatosAgendaFiltro.map(fila => fila[6]);
      let horasDisponibles = HORAS_LABORALES.filter(hora => !horasNoDisponibles.includes(hora));

      // Paso 6: Crear dropdown en G9 con SOLO las horas libres (fondo cyan)
      hoja.getRange("G9").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(horasDisponibles).setAllowInvalid(false).build()).setBackground("cyan");

      // Paso 7: Extraer datos para mostrar en las celdas. Las columnas de BD son:
      //   [0]=ID, [1]=Día, [2]=Mes, [3]=Año, [4]=Profesional, [5]=Área
      //   [6]=Hora, [7]=Celular, [8]=Procedimiento, [9]=Observación, [10]=Preferencia
      let datosAgenda1 = baseDeDatosAgendaFiltro.map(fila => [fila[6], fila[7]]);           // Hora + Celular
      let datosAgenda2 = baseDeDatosAgendaFiltro.map(fila => [fila[8], fila[9], fila[10]]); // Proc + Obs + Pref

      if (datosAgenda1.length > 0 && datosAgenda2.length > 0) {
        // Escribir las citas en las filas 15+ con fondo lavanda
        hoja.getRange(15, 6, datosAgenda1.length, 2).setValues(datosAgenda1).setBackground("Lavender");
        hoja.getRange(15, 9, datosAgenda2.length, 3).setValues(datosAgenda2).setBackground("Lavender");
      } else {
        hoja.getRange("J15").setValue("Libre, hay cita(s) cancelada(s).").setBackground("Khaki");
      }
    } else {
      // No hay citas: mostrar TODAS las horas como disponibles (fondo rosa)
       hoja.getRange("G9").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(HORAS_LABORALES).setAllowInvalid(false).build()).setBackground("HotPink");
    }

    // Paso 8: Marcar bandera = 1 en hoja Info para que el botón "Agendar" sepa que ya se consultó
    // Agenda 1 -> celda J2, Agenda 2 -> celda J3
    hojaInfo.getRange(1 + numAgenda, 10).setValue(1); 
  } else {
    ui.alert(`Error: profesional: ${profesional} - dia: ${dia} - mes: ${mes} - año: ${anio}`);
  }
}

/**
 * agendar(numAgenda)
 * Guarda una nueva cita en la base de datos de agenda externa.
 * Usa LockService para evitar que dos recepcionistas agenden en la misma hora al mismo tiempo.
 *
 * VALIDACIONES:
 *   - La bandera en Info debe ser = 1 (se ejecutó verAgenda() primero)
 *   - El celular del cliente debe existir en la base de datos (busquedaBinaria)
 *   - Todos los campos deben estar llenos
 *
 * FLUJO:
 *   1. Adquirir candado (Lock)
 *   2. Validar bandera + existencia del cliente
 *   3. Leer todos los campos del formulario de la hoja Agenda_X
 *   4. Generar ID único y timestamp
 *   5. Escribir la fila nueva en la BD de agenda
 *   6. Mostrar el enlace de WhatsApp para confirmar la cita
 *   7. Refrescar la vista de la agenda y resetear la bandera
 *
 * @param {number} numAgenda - Número de agenda (1 o 2).
 */
function agendar(numAgenda) {
  let lock = LockService.getScriptLock();
  let success = lock.tryLock(10000); // Intentar adquirir candado por hasta 10 segundos
  let ui = SpreadsheetApp.getUi();

  if (!success) {
    ui.alert("El sistema está ocupado guardando otra cita. Por favor intenta en unos segundos.");
    return;
  }
  
  try {
    let libro = SpreadsheetApp.getActive();
    let hoja = libro.getSheetByName("Agenda_" + numAgenda);
    let hojaInfo = libro.getSheetByName("Info");
    let bd = SpreadsheetApp.openById(idBaseDeDatosAgenda).getSheetByName(hojaBaseDeDatosAgenda);

    // Leer la bandera de "agenda consultada" y el celular ingresado
    let bandera = hojaInfo.getRange(1 + numAgenda, 10).getValue();
    let celular = hoja.getRange("G3").getValue();
    // Verificar que el celular exista en la base de datos de clientes
    let bandera2 = celular ? busquedaBinaria(celular)[0] : false;

    if (bandera === 1 && bandera2) {
      // Leer todos los campos del formulario en la hoja de agenda
      let profesional = hoja.getRange("C3").getValue();
      let area = hoja.getRange("C6").getValue();
      let dia = hoja.getRange("C8").getValue();
      let mes = hoja.getRange("C11").getValue();
      let anio = hoja.getRange("C14").getValue();
      let procedimiento = hoja.getRange("G6").getValue();
      let hora = hoja.getRange("G9").getValue();
      let preferencia = hoja.getRange("L3").getValue();     // Checkbox true/false
      let observacion = hoja.getRange("K8").getValue();

      if (profesional && dia && mes && anio && celular && procedimiento && hora && observacion) {
        let ultimaFila = bd.getLastRow() + 1;
        let fechaActual = new Date();
        // ID único: segundos transcurridos desde el 1 de marzo 2025
        let id = Math.floor((fechaActual.getTime() - new Date(2025, 2, 1).getTime()) / 1000);
        let correo = Session.getActiveUser().getEmail();
        // Convertir checkbox booleano a texto "SI"/"NO"
        preferencia = preferencia ? "SI" : "NO";

        // Escribir la nueva cita en la base de datos (14 columnas: A-N)
        bd.getRange(ultimaFila, 1, 1, 14).setValues([[id, dia, mes, anio, profesional, area, hora, celular, procedimiento, observacion, preferencia, "Agendada", correo, fechaActual]]);

        // Mostrar enlace de WhatsApp para confirmar la cita al cliente
        apiWhatsApp(celular, area, dia, hora);

        // Refrescar la vista y resetear la bandera
        limpiarVer(numAgenda);
        verAgenda(numAgenda);
        hojaInfo.getRange(1 + numAgenda, 10).setValue(0);
      } else {
        ui.alert("Faltan datos por llenar.");
      }
    } else {
      // Si falló la validación, recargar la agenda y avisar
      verAgenda(numAgenda);
      ui.alert("El celular no esta registrado o error al visualizar agenda, cita NO asignada.");
    }
  } finally {
    // SIEMPRE liberar el candado
    lock.releaseLock();
  }
}

/**
 * bloquear(numAgenda)
 * Bloquea TODAS las horas de un día para un profesional específico.
 * Se usa para cerrar un día completo (ej: festivos, incapacidades).
 * Requiere ingresar la contraseña "qwe123+" como medida de seguridad
 * para que solo administradores puedan bloquear agendas.
 *
 * FUNCIONAMIENTO:
 *   Crea una fila de tipo "Bloqueada" para CADA hora de HORAS_LABORALES,
 *   llenando la agenda completa e impidiendo que se agende nada más ese día.
 *
 * @param {number} numAgenda - Número de agenda (1 o 2).
 */
function bloquear(numAgenda) {
  let lock = LockService.getScriptLock();
  let success = lock.tryLock(10000);
  let ui = SpreadsheetApp.getUi();
  if (!success) { return ui.alert("El sistema está ocupado. Intenta de nuevo."); }

  try {
    let libro = SpreadsheetApp.getActive();
    let hoja = libro.getSheetByName("Agenda_" + numAgenda);
    let hojaInfo = libro.getSheetByName("Info");

    // Solicitar código de seguridad al usuario
    let codigo = ui.prompt("Ingrese el codigo.");
    
    if (codigo.getSelectedButton() === ui.Button.OK) {
      // Verificar contraseña
      if (codigo.getResponseText() === "qwe123+") {
        let profesional = hoja.getRange("C3").getValue();
        let area = hoja.getRange("C6").getValue();
        let dia = hoja.getRange("C8").getValue();
        let mes = hoja.getRange("C11").getValue();
        let anio = hoja.getRange("C14").getValue();

        if (profesional && dia && mes && anio) {
          let bd = SpreadsheetApp.openById(idBaseDeDatosAgenda).getSheetByName(hojaBaseDeDatosAgenda);
          let baseDeDatosDeAgendaBloqueada = [];
          
          let fechaActual = new Date();
          let id = Math.floor((fechaActual.getTime() - new Date(2025, 2, 1).getTime()) / 1000);
          let correo = Session.getActiveUser().getEmail();
          let observacion = `Agenda bloqueada ${fechaActual} por ${correo}`;

          // Crear una fila "Bloqueada" para cada hora del día
          HORAS_LABORALES.forEach(hora => {
            baseDeDatosDeAgendaBloqueada.push([id, dia, mes, anio, profesional, area, hora, "NA", "NA", observacion, "NO", "Bloqueada", correo, fechaActual]);
          });

          // Escribir todas las filas de bloqueo de una sola vez
          bd.getRange(bd.getLastRow() + 1, 1, baseDeDatosDeAgendaBloqueada.length, 14).setValues(baseDeDatosDeAgendaBloqueada);
          ui.alert("¡Agenda bloqueada!");
          verAgenda(numAgenda);
          hojaInfo.getRange(1 + numAgenda, 10).setValue(0);
        } else {
          ui.alert("Faltan datos por llenar.");
        }
      } else {
        ui.alert("¡Codigo incorrecto!, intentar nuevamente.");
      }
    }
  } finally {
    lock.releaseLock();
  }
}


// ╔══════════════════════════════════════════════╗
// ║   SECCIÓN 3: FUNCIONES DE VISTA DIARIA      ║
// ║   (Dia_1, Dia_2)                             ║
// ╚══════════════════════════════════════════════╝

/**
 * btnActualizarDia(numDia)
 * Lee la fecha de la hoja "Dia_X" (A2=año, C2=mes, E2=día),
 * consulta la BD de agenda externa y muestra TODAS las citas de ese día
 * (de todos los profesionales), ordenadas por profesional y por hora.
 * Aplica colores pastel para distinguir visualmente a cada profesional.
 *
 * @param {number} numDia - Número de la vista de día (1 o 2).
 */
function btnActualizarDia(numDia) {
  let ui = SpreadsheetApp.getUi();
  let hoja = SpreadsheetApp.getActive().getSheetByName("Dia_" + numDia);
  if (!hoja) return;

  // Leer la fecha de los filtros de la hoja
  let dia = hoja.getRange("E2").getValue();
  let mes = hoja.getRange("C2").getValue();
  let anio = hoja.getRange("A2").getValue();

  if (dia && mes && anio) {
    // Limpiar datos anteriores desde la fila 6 hacia abajo
    hoja.getRange("A6:K").setValue("").setBackground("white");

    // Traer toda la BD y filtrar por día/mes/año
    let baseDeDatosAgenda = SpreadsheetApp.openById(idBaseDeDatosAgenda).getSheetByName(hojaBaseDeDatosAgenda).getDataRange().getValues();
    let datosDia = baseDeDatosAgenda.filter(f => f[1] === dia && f[2] === mes && f[3] === anio);

    if (datosDia.length > 0) {
      // Extraer columnas relevantes: [ID, Profesional, Área, Hora, Celular, Procedimiento, Observación, Preferencia, Estado]
      datosDia = datosDia.map(f => [f[0], f[4], f[5], f[6], f[7], f[8], f[9], f[10], f[11]]);

      // Ordenar primero por nombre del profesional, luego por hora dentro del mismo profesional
      datosDia.sort((a, b) => {
        const nA = a[1].toString().toLowerCase(), nB = b[1].toString().toLowerCase();
        if (nA !== nB) return nA < nB ? -1 : 1;
        const iA = HORAS_LABORALES.indexOf(a[3]), iB = HORAS_LABORALES.indexOf(b[3]);
        return iA - iB;
      });

      // Escribir las citas desde la fila 6
      hoja.getRange(6, 1, datosDia.length, 9).setValues(datosDia);
      // Aplicar colores pastel alternados por profesional (columna 2 = profesional, 8 cols de ancho)
      calendarioColor(hoja, 2, 8);
    } else {
      ui.alert(`El ${dia} no tiene citas.`);
    }
  } else {
    ui.alert(`Datos incompletos`);
  }
}

/**
 * calendarioColor(hoja, startCol, numCols)
 * Aplica colores de fondo alternados (de COLORES_PASTEL_AGENDA) a las filas
 * de una hoja, cambiando de color cada vez que cambia el nombre del profesional.
 * Esto permite distinguir visualmente los bloques de cada profesional.
 *
 * @param {Sheet} hoja - Referencia a la hoja de Google Sheets.
 * @param {number} startCol - Columna de inicio para leer el nombre (y colorear).
 * @param {number} numCols - Cantidad de columnas a colorear.
 */
function calendarioColor(hoja, startCol, numCols) {
  // Leer los nombres de profesional desde la fila 6 hasta la última
  let datos = hoja.getRange(6, startCol, hoja.getLastRow() - 5, numCols).getValues();
  let nombreAnterior = null, m = 0;

  for (let i = 0; i < datos.length; i++) {
    let nombreActual = datos[i][0]; // Primera columna del rango = nombre del profesional

    // Si cambió el profesional, avanzar al siguiente color
    if (nombreActual !== nombreAnterior) {
      m = (m + 1) % COLORES_PASTEL_AGENDA.length; // Ciclar entre los colores
      nombreAnterior = nombreActual;
    }

    // Aplicar el color de fondo a toda la fila
    hoja.getRange(i + 6, startCol, 1, numCols).setBackground(COLORES_PASTEL_AGENDA[m]);
  }
}

/**
 * btnCambiarEstadoDia(numDia)
 * Sincroniza los cambios de estado hechos por el usuario en la vista "Dia_X"
 * de vuelta a la base de datos de agenda externa.
 * Usa LockService para evitar conflictos de escritura.
 *
 * REGLAS DE NEGOCIO:
 *   - Si el estado nuevo es "Cancelada": limpia el celular y agrega nota de cancelación.
 *   - No permite modificar estados que ya son "Cancelada" o "Bloqueada".
 *   - Si el estado cambió a cualquier otro: actualiza la observación y el estado.
 *
 * @param {number} numDia - Número de la vista de día (1 o 2).
 */
function btnCambiarEstadoDia(numDia) {
  let lock = LockService.getScriptLock();
  if(!lock.tryLock(10000)) return SpreadsheetApp.getUi().alert("Espera un momento.");
  
  try {
    let hoja = SpreadsheetApp.getActive().getSheetByName("Dia_" + numDia);
    let bd = SpreadsheetApp.openById(idBaseDeDatosAgenda).getSheetByName(hojaBaseDeDatosAgenda);

    // Leer TODA la BD en memoria para buscar por ID
    let dataBD = bd.getDataRange().getValues();
    // Leer las filas editadas por el usuario en la vista de día (columnas A-I desde fila 6)
    let datos = hoja.getRange(6, 1, hoja.getLastRow() - 5, 9).getValues();

    for (let f of datos) {
      // Desestructurar: id=col1, obs=col7, est=col9 (los demás se ignoran con comas)
      let [id, , , , , , obs, , est] = f;
      if (!id) continue; // Saltar filas vacías
      
      // Buscar la fila correspondiente en la BD por su ID
      let idx = dataBD.findIndex(row => row[0] === id);
      if (idx >= 0) {
        let currentEst = dataBD[idx][11]; // Columna L = índice 11 = Estado actual en BD

        // Solo actualizar si el estado cambió Y no es una cita cancelada/bloqueada previamente
        if (currentEst !== est && currentEst !== "Cancelada" && currentEst !== "Bloqueada") {
          if (est === "Cancelada") {
            // Si se cancela: limpiar el celular (columna H = col 8) y agregar nota
            bd.getRange(idx + 1, 8).setValue("");
            bd.getRange(idx + 1, 11).setValue(`Cita cancelada: ${new Date()}`);
          } else {
            // Para otros cambios de estado: actualizar la observación
            bd.getRange(idx + 1, 11).setValue(obs);
          }
          // Actualizar el estado en la BD (columna L = col 12)
          bd.getRange(idx + 1, 12).setValue(est);
        }
      }
    }

    // Refrescar la vista después de los cambios
    btnActualizarDia(numDia);
  } finally {
    lock.releaseLock();
  }
}

/**
 * btnWhatsAppDia(numDia)
 * Genera enlaces de WhatsApp personalizados para CADA cita del día,
 * con mensajes diferentes según el estado de la cita
 * (Agendada, Confirmacion 1, Confirmacion 2, Asistencia, Cancelada, u otro).
 * Los enlaces se insertan como fórmulas HYPERLINK en la columna K.
 *
 * @param {number} numDia - Número de la vista de día (1 o 2).
 */
function btnWhatsAppDia(numDia) {
  let hoja = SpreadsheetApp.getActive().getSheetByName("Dia_" + numDia);
  let dia = hoja.getRange("E2").getValue();

  // Limpiar la columna K (enlaces previos)
  hoja.getRange("K6:K").setValue("").setBackground("white");
  // Refrescar la vista antes de generar enlaces
  btnActualizarDia(numDia);

  // Leer los datos de las citas (desde columna B hasta columna I, 8 columnas)
  let datos = hoja.getRange(6, 2, hoja.getLastRow() - 5, 8).getValues();

  if (datos.length > 0) {
    for (let i = 0; i < datos.length; i++) {
      let area = datos[i][1];     // Col C = Área (MANICURISTA, etc.)
      let hora = datos[i][2];     // Col D = Hora
      let celular = datos[i][3];  // Col E = Celular
      let estado = datos[i][7];   // Col I = Estado
      let msg = "";

      // Función auxiliar para codificar caracteres especiales en URL
      let uri = (text) => encodeURI(text);

      // Seleccionar el mensaje según el estado actual de la cita
      if (estado === "Agendada") {
        // Recordatorio del día anterior con info de servicios adicionales
        msg = `¡Hola Bell@! ✨ Espero que estés teniendo un día increíble. Te escribimos para confirmar tu cita de ${area} para el día mañana *${dia}* en la *SEDE DE ${sedeMensaje}* a las *${hora}*. 💅🏼👁😊🕓\n\n¡Recuerda que también tenemos servicios de *MANOS*, *PIES* 🦶🏼, *CEJAS* y *PESTAÑAS* para ti! 👁 Pregunta por nuestros catálogos.\n\nTen en cuenta:\n1. Pago con tarjeta 💳 tiene incremento de $3.000.\n2. Facturación electrónica debe solicitarse en caja.💕`;
      } else if (estado === "Confirmacion 1") {
        // Primera confirmación el mismo día
        msg = `Hola bell@, como estas?, confirmo tu cita con tu *${area}* del dia de hoy *${dia}* a las *${hora}* en la sede de *${sedeMensaje}*? 🥰💅\nRecuerda que te esperamos con mucho amor.`;
      } else if (estado === "Confirmacion 2") {
        // Segunda confirmación (más corta, recordatorio final)
        msg = `Recuerda que estas pronto para tu cita. Te estamos esperando 💞🥰`;
      } else if (estado === "Asistencia") {
        // Encuesta de satisfacción post-servicio
        msg = `¡Hola, mi corazón! 💞🌞\nQueremos saber cómo te fue en tu servicio con tu *${area}*. ¿Estás feliz con el resultado?\nNos encantaría recibir tu calificación:\n1. 😡 Pésimo servicio\n2. 🥺 Mal servicio\n3. 🫤 Servicio regular\n4. 😊 Buen servicio\n5. 🥰 ¡Excelente servicio!\n\n¡Siempre será un placer atenderte! ❤🫶🏼`;
      } else if (estado === "Cancelada") {
        // Aviso de cancelación con oferta de reagendar
        msg = `¡Hola! Lamentamos informarte que tu cita con tu *${area}* para el día *${dia}* ha sido cancelada.\nDeseas reagendar?`;
      } else {
        // Mensaje genérico para estados personalizados
        msg = `¡Hola! Queremos informarte que tu cita con tu *${area}* para el día *${dia}* tiene novedades.\nPor favor, contáctanos.`;
      }

      // Insertar fórmula HYPERLINK en columna K con el enlace de WhatsApp
      hoja.getRange(i + 6, 11).setFormula(`=HYPERLINK("https://api.whatsapp.com/send?phone=+57${celular}&text=${uri(msg)}"; "Enviar mensaje")`);
    }
  } else {
    SpreadsheetApp.getUi().alert("No hay agenda disponible.");
  }
}


// ╔══════════════════════════════════════════════╗
// ║   SECCIÓN 4: VISTA MENSUAL                  ║
// ╚══════════════════════════════════════════════╝

/**
 * agendaMensual()
 * Muestra TODAS las citas de un mes y profesional específico
 * en la hoja "Mes", ordenadas por día y hora.
 * Los filtros se leen de: B2=año, D2=mes, H2=profesional.
 * Aplica colores pastel para distinguir cada día.
 */
function agendaMensual() {
  let ui = SpreadsheetApp.getUi();
  let hoja = SpreadsheetApp.getActive().getSheetByName("Mes");
  let anio = hoja.getRange("B2").getValue();
  let mes = hoja.getRange("D2").getValue();
  let profesional = hoja.getRange("H2").getValue();

  if (mes && anio && profesional) {
    // Limpiar datos anteriores
    hoja.getRange("A6:I").setValue("").setBackground("white");

    // Filtrar BD por mes, año y profesional
    let baseDeDatosAgenda = SpreadsheetApp.openById(idBaseDeDatosAgenda).getSheetByName(hojaBaseDeDatosAgenda).getDataRange().getValues();
    let datosMes = baseDeDatosAgenda.filter(f => f[2] === mes && f[3] === anio && f[4] === profesional);

    if (datosMes.length > 0) {
      // Extraer columnas: [ID, Día, Área, Hora, Celular, Procedimiento, Observación, Preferencia, Estado]
      datosMes = datosMes.map(f => [f[0], f[1], f[5], f[6], f[7], f[8], f[9], f[10], f[11]]);

      // Ordenar por día y luego por hora dentro del mismo día
      datosMes.sort((a, b) => {
        const nA = a[1].toString().toLowerCase(), nB = b[1].toString().toLowerCase();
        if (nA !== nB) return nA < nB ? -1 : 1;
        return HORAS_LABORALES.indexOf(a[3]) - HORAS_LABORALES.indexOf(b[3]);
      });

      // Escribir y colorear
      hoja.getRange(6, 1, datosMes.length, 9).setValues(datosMes);
      calendarioColor(hoja, 1, 9);
    } else {
      ui.alert(`El ${mes} con profesional ${profesional} no tiene citas.`);
    }
  } else {
    ui.alert("Datos incompletos.");
  }
}


// ╔══════════════════════════════════════════════╗
// ║   SECCIÓN 5: WRAPPERS PARA BOTONES          ║
// ╚══════════════════════════════════════════════╝
//
// Google Sheets requiere que los botones (formas/imágenes) invoquen
// funciones SIN parámetros. Estas funciones "envoltorio" simplemente
// llaman a la función genérica con el número de agenda o día correcto.
//
// Si se agrega una Agenda_3, solo se necesita añadir:
//   function limpiarVer3() { limpiarVer(3); }
//   function verAgenda3() { verAgenda(3); }
//   ... etc.
//

// --- Wrappers de Agenda 1 ---
function limpiarVer1() { limpiarVer(1); }
function verAgenda1() { verAgenda(1); }
function agendar1() { agendar(1); }
function bloquear1() { bloquear(1); }

// --- Wrappers de Agenda 2 ---
function limpiarVer2() { limpiarVer(2); }
function verAgenda2() { verAgenda(2); }
function agendar2() { agendar(2); }
function bloquear2() { bloquear(2); }

// --- Wrappers de Vista Dia 1 ---
function btnActualizar1() { btnActualizarDia(1); }
function calendarioColor1() { btnActualizarDia(1); } // Alias para compatibilidad
function btnCambiarEstado1() { btnCambiarEstadoDia(1); }
function btnWhatsApp1() { btnWhatsAppDia(1); }

// --- Wrappers de Vista Dia 2 ---
function btnActualizar2() { btnActualizarDia(2); }
function calendarioColor2() { btnActualizarDia(2); } // Alias para compatibilidad
function btnCambiarEstado2() { btnCambiarEstadoDia(2); }
function btnWhatsApp2() { btnWhatsAppDia(2); }
