// Funciones globales para clientes
function obtenerNumerosClientes() {
  let libro = SpreadsheetApp.openById(idBaseDeDatosClientes);
  let base = libro.getSheetByName(hojaBaseDeDatosClientes);
  let datos = base.getRange("A2:A").getValues();
  return datos.flat().filter(String);
}

function obtenerDatosClientes(numero) {
  try {
    const libro = SpreadsheetApp.openById(idBaseDeDatosClientes);
    const base = libro.getSheetByName(hojaBaseDeDatosClientes);
    let inicio = 2, final = base.getLastRow();
    while (inicio <= final) {
      let mitad = Math.floor((inicio + final) / 2);
      let posicionDato = Number(base.getRange(`A${mitad}`).getValue());
      if (numero === posicionDato) {
        let fila = base.getRange(mitad, 1, 1, 6).getValues()[0];
        let fecha = new Date(fila[5]);
        let dia = fecha.getDate().toString().padStart(2, "0");
        let mes = (fecha.getMonth() + 1).toString().padStart(2, "0");
        let cumpleañosFormateado = `${dia}/${mes}`;
        return `Nombre: ${fila[1]}\nTipo Doc.: ${fila[2]}\nNúmero Doc.: ${fila[3]}\nCorreo: ${fila[4]}\nCumpleaños: ${cumpleañosFormateado}`;
      } else if (numero < posicionDato) final = mitad - 1;
      else inicio = mitad + 1;
    }
    return `Cliente con número ${numero} no encontrado.`;
  } catch (error) {
    Logger.log("Error en obtenerDatosClientes: " + error.message);
    return "Error al buscar el cliente.";
  }
}

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
        let fila = base.getRange(mitad, 1, 1, 9).getValues()[0];
        hojaEliminados.appendRow([new Date(), ...fila.slice(0, 9)]);
        base.deleteRow(mitad);
        return `Cliente con número ${numero} eliminado correctamente.`;
      } else if (numero < posicionDato) final = mitad - 1;
      else inicio = mitad + 1;
    }
    return `Cliente con número ${numero} no encontrado.`;
  } catch (error) {
    return "Error al eliminar el cliente.";
  }
}

function apiWhatsApp(celular, area, dia, hora) {
  let mensaje = `https://api.whatsapp.com/send?phone=+57${celular}&text=¡Hola%20con%20mucho%20amor!%20😊%20Queremos%20recordarte%20que%20tu%20cita%20de%20*${area}*%20ha%20sido%20generada%20para%20el%20día%20*${dia}*%20a%20las%20*${hora}*.%20¡Te%20esperamos!%20💕`;
  let html = HtmlService.createHtmlOutput(`<html><head><style>a { font-weight: 800; font-size: 18px; color: Lime; } a:hover { color: gray; }</style></head><body><a href="${mensaje}" target="_blank">¡Enviar confirmacion!</a></body></html>`).setWidth(300).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, "Enlace WhatsApp");
}

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

function guardarCliente(data) {
  let lock = LockService.getScriptLock();
  lock.tryLock(10000); // 10s wait Multiple User LockService
  try {
    const hoja = SpreadsheetApp.openById(idBaseDeDatosClientes).getSheetByName(hojaBaseDeDatosClientes);
    let fechaActual = new Date();
    let id = Math.floor((fechaActual.getTime() - new Date(2025, 0, 1).getTime()) / 1000);
    hoja.appendRow([data.customerNumber, data.name.toUpperCase(), data.docType.toUpperCase(), data.docNumber, data.email.toUpperCase(), data.birthDate, id, Session.getActiveUser().getEmail().toUpperCase(), fechaActual]);
    hoja.getRange("A2:I").sort({ column: 1, ascending: true });
  } finally {
    lock.releaseLock();
  }
}

// ============================================
// FUNCIONES GENERALIZADAS DE AGENDA
// ============================================

function limpiarVer(numAgenda) {
  let hoja = SpreadsheetApp.getActive().getSheetByName("Agenda_" + numAgenda);
  if (!hoja) return;
  hoja.getRange("G3").setValue("").setBackground("white");
  hoja.getRange("G6").setValue("").setBackground("white");
  hoja.getRange("G9").setValue("").setBackground("white");
  hoja.getRange("F15:G").setValue("").setBackground("white");
  hoja.getRange("I15:K").setValue("").setBackground("white");
  hoja.getRange("L3").setValue(false).setBackground("white");
  hoja.getRange("K8").setValue("Escribe una observacion...").setBackground("white");
}

function verAgenda(numAgenda) {
  let ui = SpreadsheetApp.getUi();
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Agenda_" + numAgenda);
  let hojaInfo = libro.getSheetByName("Info");
  if (!hoja) return ui.alert("Hoja Agenda_" + numAgenda + " no encontrada.");

  limpiarVer(numAgenda);
  let profesional = hoja.getRange("C3").getValue();
  let dia = hoja.getRange("C8").getValue();
  let mes = hoja.getRange("C11").getValue();
  let anio = hoja.getRange("C14").getValue();

  if (profesional && dia && mes && anio) {
    let baseDeDatosAgenda = SpreadsheetApp.openById(idBaseDeDatosAgenda).getSheetByName(hojaBaseDeDatosAgenda).getDataRange().getValues();
    let baseDeDatosAgendaFiltro = baseDeDatosAgenda.filter(fila => fila[4] === profesional && fila[1] === dia && fila[2] === mes && fila[3] === anio);

    if (baseDeDatosAgendaFiltro.length > 0) {
      let ordenHoras = {};
      HORAS_LABORALES.forEach((hora, index) => ordenHoras[hora] = index);
      baseDeDatosAgendaFiltro.sort((a, b) => ordenHoras[a[6]] - ordenHoras[b[6]]);
      let horasNoDisponibles = baseDeDatosAgendaFiltro.map(fila => fila[6]);
      let horasDisponibles = HORAS_LABORALES.filter(hora => !horasNoDisponibles.includes(hora));

      hoja.getRange("G9").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(horasDisponibles).setAllowInvalid(false).build()).setBackground("cyan");

      let datosAgenda1 = baseDeDatosAgendaFiltro.map(fila => [fila[6], fila[7]]);
      let datosAgenda2 = baseDeDatosAgendaFiltro.map(fila => [fila[8], fila[9], fila[10]]);

      if (datosAgenda1.length > 0 && datosAgenda2.length > 0) {
        hoja.getRange(15, 6, datosAgenda1.length, 2).setValues(datosAgenda1).setBackground("Lavender");
        hoja.getRange(15, 9, datosAgenda2.length, 3).setValues(datosAgenda2).setBackground("Lavender");
      } else {
        hoja.getRange("J15").setValue("Libre, hay cita(s) cancelada(s).").setBackground("Khaki");
      }
    } else {
       hoja.getRange("G9").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(HORAS_LABORALES).setAllowInvalid(false).build()).setBackground("HotPink");
    }
    // Set flag in info sheet: Agenda 1 -> J2, Agenda 2 -> J3, etc.
    hojaInfo.getRange(1 + numAgenda, 10).setValue(1); 
  } else {
    ui.alert(`Error: profesional: ${profesional} - dia: ${dia} - mes: ${mes} - año: ${anio}`);
  }
}

function agendar(numAgenda) {
  let lock = LockService.getScriptLock();
  let success = lock.tryLock(10000); // 10s wait LockService
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

    let bandera = hojaInfo.getRange(1 + numAgenda, 10).getValue();
    let celular = hoja.getRange("G3").getValue();
    let bandera2 = celular ? busquedaBinaria(celular)[0] : false;

    if (bandera === 1 && bandera2) {
      let profesional = hoja.getRange("C3").getValue();
      let area = hoja.getRange("C6").getValue();
      let dia = hoja.getRange("C8").getValue();
      let mes = hoja.getRange("C11").getValue();
      let anio = hoja.getRange("C14").getValue();
      let procedimiento = hoja.getRange("G6").getValue();
      let hora = hoja.getRange("G9").getValue();
      let preferencia = hoja.getRange("L3").getValue();
      let observacion = hoja.getRange("K8").getValue();

      if (profesional && dia && mes && anio && celular && procedimiento && hora && observacion) {
        let ultimaFila = bd.getLastRow() + 1;
        let fechaActual = new Date();
        let id = Math.floor((fechaActual.getTime() - new Date(2025, 2, 1).getTime()) / 1000);
        let correo = Session.getActiveUser().getEmail();
        preferencia = preferencia ? "SI" : "NO";

        bd.getRange(ultimaFila, 1, 1, 14).setValues([[id, dia, mes, anio, profesional, area, hora, celular, procedimiento, observacion, preferencia, "Agendada", correo, fechaActual]]);
        apiWhatsApp(celular, area, dia, hora);
        limpiarVer(numAgenda);
        verAgenda(numAgenda);
        hojaInfo.getRange(1 + numAgenda, 10).setValue(0);
      } else {
        ui.alert("Faltan datos por llenar.");
      }
    } else {
      verAgenda(numAgenda);
      ui.alert("El celular no esta registrado o error al visualizar agenda, cita NO asignada.");
    }
  } finally {
    lock.releaseLock();
  }
}

function bloquear(numAgenda) {
  let lock = LockService.getScriptLock();
  let success = lock.tryLock(10000);
  let ui = SpreadsheetApp.getUi();
  if (!success) { return ui.alert("El sistema está ocupado. Intenta de nuevo."); }

  try {
    let libro = SpreadsheetApp.getActive();
    let hoja = libro.getSheetByName("Agenda_" + numAgenda);
    let hojaInfo = libro.getSheetByName("Info");
    let codigo = ui.prompt("Ingrese el codigo.");
    
    if (codigo.getSelectedButton() === ui.Button.OK) {
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

          HORAS_LABORALES.forEach(hora => {
            baseDeDatosDeAgendaBloqueada.push([id, dia, mes, anio, profesional, area, hora, "NA", "NA", observacion, "NO", "Bloqueada", correo, fechaActual]);
          });

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

// ============================================
// FUNCIONES GENERALIZADAS DE DÍA
// ============================================

function btnActualizarDia(numDia) {
  let ui = SpreadsheetApp.getUi();
  let hoja = SpreadsheetApp.getActive().getSheetByName("Dia_" + numDia);
  if (!hoja) return;

  let dia = hoja.getRange("E2").getValue();
  let mes = hoja.getRange("C2").getValue();
  let anio = hoja.getRange("A2").getValue();

  if (dia && mes && anio) {
    hoja.getRange("A6:K").setValue("").setBackground("white");
    let baseDeDatosAgenda = SpreadsheetApp.openById(idBaseDeDatosAgenda).getSheetByName(hojaBaseDeDatosAgenda).getDataRange().getValues();
    let datosDia = baseDeDatosAgenda.filter(f => f[1] === dia && f[2] === mes && f[3] === anio);

    if (datosDia.length > 0) {
      datosDia = datosDia.map(f => [f[0], f[4], f[5], f[6], f[7], f[8], f[9], f[10], f[11]]);
      datosDia.sort((a, b) => {
        const nA = a[1].toString().toLowerCase(), nB = b[1].toString().toLowerCase();
        if (nA !== nB) return nA < nB ? -1 : 1;
        const iA = HORAS_LABORALES.indexOf(a[3]), iB = HORAS_LABORALES.indexOf(b[3]);
        return iA - iB;
      });
      hoja.getRange(6, 1, datosDia.length, 9).setValues(datosDia);
      calendarioColor(hoja, 2, 8);
    } else {
      ui.alert(`El ${dia} no tiene citas.`);
    }
  } else {
    ui.alert(`Datos incompletos`);
  }
}

function calendarioColor(hoja, startCol, numCols) {
  let datos = hoja.getRange(6, startCol, hoja.getLastRow() - 5, numCols).getValues();
  let nombreAnterior = null, m = 0;
  for (let i = 0; i < datos.length; i++) {
    let nombreActual = datos[i][0];
    if (nombreActual !== nombreAnterior) {
      m = (m + 1) % COLORES_PASTEL_AGENDA.length;
      nombreAnterior = nombreActual;
    }
    hoja.getRange(i + 6, startCol, 1, numCols).setBackground(COLORES_PASTEL_AGENDA[m]);
  }
}

function btnCambiarEstadoDia(numDia) {
  let lock = LockService.getScriptLock();
  if(!lock.tryLock(10000)) return SpreadsheetApp.getUi().alert("Espera un momento.");
  
  try {
    let hoja = SpreadsheetApp.getActive().getSheetByName("Dia_" + numDia);
    let bd = SpreadsheetApp.openById(idBaseDeDatosAgenda).getSheetByName(hojaBaseDeDatosAgenda);
    let dataBD = bd.getDataRange().getValues();
    let datos = hoja.getRange(6, 1, hoja.getLastRow() - 5, 9).getValues();

    for (let f of datos) {
      let [id, , , , , , obs, , est] = f;
      if (!id) continue;
      
      let idx = dataBD.findIndex(row => row[0] === id);
      if (idx >= 0) {
        let currentEst = dataBD[idx][11]; // index 11 is Estado (Column L -> index 11)
        if (currentEst !== est && currentEst !== "Cancelada" && currentEst !== "Bloqueada") {
          if (est === "Cancelada") {
            bd.getRange(idx + 1, 8).setValue(""); // Clear celular
            bd.getRange(idx + 1, 11).setValue(`Cita cancelada: ${new Date()}`);
          } else {
            bd.getRange(idx + 1, 11).setValue(obs);
          }
          bd.getRange(idx + 1, 12).setValue(est); // Estado
        }
      }
    }
    btnActualizarDia(numDia);
  } finally {
    lock.releaseLock();
  }
}

function btnWhatsAppDia(numDia) {
  let hoja = SpreadsheetApp.getActive().getSheetByName("Dia_" + numDia);
  let dia = hoja.getRange("E2").getValue();
  hoja.getRange("K6:K").setValue("").setBackground("white");
  btnActualizarDia(numDia);

  let datos = hoja.getRange(6, 2, hoja.getLastRow() - 5, 8).getValues();
  if (datos.length > 0) {
    for (let i = 0; i < datos.length; i++) {
      let area = datos[i][1], hora = datos[i][2], celular = datos[i][3], estado = datos[i][7];
      let msg = "";
      let uri = (text) => encodeURI(text);
      if (estado === "Agendada") msg = `¡Hola Bell@! ✨ Espero que estés teniendo un día increíble. Te escribimos para confirmar tu cita de ${area} para el día mañana *${dia}* en la *SEDE DE ${sedeMensaje}* a las *${hora}*. 💅🏼👁😊🕓\n\n¡Recuerda que también tenemos servicios de *MANOS*, *PIES* 🦶🏼, *CEJAS* y *PESTAÑAS* para ti! 👁 Pregunta por nuestros catálogos.\n\nTen en cuenta:\n1. Pago con tarjeta 💳 tiene incremento de $3.000.\n2. Facturación electrónica debe solicitarse en caja.💕`;
      else if (estado === "Confirmacion 1") msg = `Hola bell@, como estas?, confirmo tu cita con tu *${area}* del dia de hoy *${dia}* a las *${hora}* en la sede de *${sedeMensaje}*? 🥰💅\nRecuerda que te esperamos con mucho amor.`;
      else if (estado === "Confirmacion 2") msg = `Recuerda que estas pronto para tu cita. Te estamos esperando 💞🥰`;
      else if (estado === "Asistencia") msg = `¡Hola, mi corazón! 💞🌞\nQueremos saber cómo te fue en tu servicio con tu *${area}*. ¿Estás feliz con el resultado?\nNos encantaría recibir tu calificación:\n1. 😡 Pésimo servicio\n2. 🥺 Mal servicio\n3. 🫤 Servicio regular\n4. 😊 Buen servicio\n5. 🥰 ¡Excelente servicio!\n\n¡Siempre será un placer atenderte! ❤🫶🏼`;
      else if (estado === "Cancelada") msg = `¡Hola! Lamentamos informarte que tu cita con tu *${area}* para el día *${dia}* ha sido cancelada.\nDeseas reagendar?`;
      else msg = `¡Hola! Queremos informarte que tu cita con tu *${area}* para el día *${dia}* tiene novedades.\nPor favor, contáctanos.`;

      hoja.getRange(i + 6, 11).setFormula(`=HYPERLINK("https://api.whatsapp.com/send?phone=+57${celular}&text=${uri(msg)}"; "Enviar mensaje")`);
    }
  } else {
    SpreadsheetApp.getUi().alert("No hay agenda disponible.");
  }
}

// ============================================
// FUNCIONES DE AGENDA MENSUAL
// ============================================

function agendaMensual() {
  let ui = SpreadsheetApp.getUi();
  let hoja = SpreadsheetApp.getActive().getSheetByName("Mes");
  let anio = hoja.getRange("B2").getValue();
  let mes = hoja.getRange("D2").getValue();
  let profesional = hoja.getRange("H2").getValue();

  if (mes && anio && profesional) {
    hoja.getRange("A6:I").setValue("").setBackground("white");
    let baseDeDatosAgenda = SpreadsheetApp.openById(idBaseDeDatosAgenda).getSheetByName(hojaBaseDeDatosAgenda).getDataRange().getValues();
    let datosMes = baseDeDatosAgenda.filter(f => f[2] === mes && f[3] === anio && f[4] === profesional);

    if (datosMes.length > 0) {
      datosMes = datosMes.map(f => [f[0], f[1], f[5], f[6], f[7], f[8], f[9], f[10], f[11]]);
      datosMes.sort((a, b) => {
        const nA = a[1].toString().toLowerCase(), nB = b[1].toString().toLowerCase();
        if (nA !== nB) return nA < nB ? -1 : 1;
        return HORAS_LABORALES.indexOf(a[3]) - HORAS_LABORALES.indexOf(b[3]);
      });
      hoja.getRange(6, 1, datosMes.length, 9).setValues(datosMes);
      calendarioColor(hoja, 1, 9);
    } else {
      ui.alert(`El ${mes} con profesional ${profesional} no tiene citas.`);
    }
  } else {
    ui.alert("Datos incompletos.");
  }
}

// ============================================
// WRAPPERS PARA BOTONES DE GUIs ESTÁTICOS EXCEL
// ============================================

function limpiarVer1() { limpiarVer(1); }
function verAgenda1() { verAgenda(1); }
function agendar1() { agendar(1); }
function bloquear1() { bloquear(1); }

function limpiarVer2() { limpiarVer(2); }
function verAgenda2() { verAgenda(2); }
function agendar2() { agendar(2); }
function bloquear2() { bloquear(2); }

function btnActualizar1() { btnActualizarDia(1); }
function calendarioColor1() { btnActualizarDia(1); } 
function btnCambiarEstado1() { btnCambiarEstadoDia(1); }
function btnWhatsApp1() { btnWhatsAppDia(1); }

function btnActualizar2() { btnActualizarDia(2); }
function calendarioColor2() { btnActualizarDia(2); } 
function btnCambiarEstado2() { btnCambiarEstadoDia(2); }
function btnWhatsApp2() { btnWhatsAppDia(2); }
