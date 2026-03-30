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

    let inicio = 2;
    let final = base.getLastRow();

    while (inicio <= final) {
      let mitad = Math.floor((inicio + final) / 2);
      let posicionDato = Number(base.getRange(`A${mitad}`).getValue());

      if (numero === posicionDato) {
        let fila = base.getRange(mitad, 1, 1, 6).getValues()[0];

        // Formatear cumpleaños como DD/MM
        let fecha = new Date(fila[5]);
        let dia = fecha.getDate().toString().padStart(2, "0");
        let mes = (fecha.getMonth() + 1).toString().padStart(2, "0");
        let cumpleañosFormateado = `${dia}/${mes}`;

        return (
          `Nombre: ${fila[1]}\n` +
          `Tipo Doc.: ${fila[2]}\n` +
          `Número Doc.: ${fila[3]}\n` +
          `Correo: ${fila[4]}\n` +
          `Cumpleaños: ${cumpleañosFormateado}`
        );
      } else if (numero < posicionDato) {
        final = mitad - 1;
      } else {
        inicio = mitad + 1;
      }
    }

    return `Cliente con número ${numero} no encontrado.`;
  } catch (error) {
    Logger.log("Error en obtenerDatosClientes: " + error.message);
    return "Error al buscar el cliente. Revisa el registro de Apps Script.";
  }
}

function eliminarDatosClientes(numero) {
  try {
    const libro = SpreadsheetApp.openById(idBaseDeDatosClientes);
    const base = libro.getSheetByName(hojaBaseDeDatosClientes);
    const hojaEliminados = libro.getSheetByName("Eliminados");

    let inicio = 2;
    let final = base.getLastRow();

    while (inicio <= final) {
      let mitad = Math.floor((inicio + final) / 2);
      let posicionDato = Number(base.getRange(`A${mitad}`).getValue());

      if (numero === posicionDato) {
        let fila = base.getRange(mitad, 1, 1, 9).getValues()[0];
        hojaEliminados.appendRow([
          new Date(),
          fila[0],
          fila[1],
          fila[2],
          fila[3],
          fila[4],
          fila[5],
          fila[6],
          fila[7],
          fila[8],
        ]);
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
    Logger.log("Error en eliminarDatosClientes: " + error.message);
    return "Error al eliminar el cliente. Revisa el registro de Apps Script.";
  }
}

/** */ function apiWhatsApp(celular, area, dia, hora /** */) {
  let mensaje = `https://api.whatsapp.com/send?phone=+57${celular}&text=¡Hola%20con%20mucho%20amor!%20😊%20Queremos%20recordarte%20que%20tu%20cita%20de%20*${area}*%20ha%20sido%20generada%20para%20el%20día%20*${dia}*%20a%20las%20*${hora}*.%20¡Te%20esperamos!%20💕`;

  let html = HtmlService.createHtmlOutput(
    `
    
    <html>
      <head>
          <style>
            a {
              font-weight: 800; /* Normal */
              font-size: 18px;
              color: Lime;
            }
            a:hover {
              color: gray;
            }
          </style>
      </head>
      <body>
        <a href="${mensaje}" target="_blank">¡Enviar confirmacion!</a>
      </body>
    </html>
  `,
  )
    .setWidth(300)
    .setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(html, "Enlace WhatsApp");
}

/** */ function busquedaBinaria(celular /** */) {
  let libro = SpreadsheetApp.openById(idBaseDeDatosClientes);
  let base = libro.getSheetByName(hojaBaseDeDatosClientes);

  try {
    let inicio = 1;
    let final = base.getLastRow();
    while (inicio <= final) {
      let mitad = Math.floor((inicio + final) / 2);
      let posicionDato = Number(base.getRange(`A${mitad}`).getValue());
      if (celular === posicionDato) {
        return [true, mitad];
      } else if (celular < posicionDato) {
        final = mitad - 1;
      } else {
        inicio = mitad + 1;
      }
    }
    return [false, null];
  } catch (error) {
    Logger.log(error);
  }
}

/** */ function guardarCliente(data /** */) {
  const hoja = SpreadsheetApp.openById(idBaseDeDatosClientes).getSheetByName(
    hojaBaseDeDatosClientes,
  );

  let fechaInicio = new Date(2025, 0, 1);
  let fechaActual = new Date();
  let correo = Session.getActiveUser().getEmail();
  let id = Math.floor((fechaActual.getTime() - fechaInicio.getTime()) / 1000);

  hoja.appendRow([
    data.customerNumber,
    data.name.toUpperCase(),
    data.docType.toUpperCase(),
    data.docNumber,
    data.email.toUpperCase(),
    data.birthDate,
    id,
    correo.toUpperCase(),
    fechaActual,
  ]);

  hoja.getRange("A2:I").sort({ column: 1, ascending: true });
}

/** */ function verAgenda1 /** */() {
  let ui = SpreadsheetApp.getUi();
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Agenda_1");
  let hojaInfo = libro.getSheetByName("Info");

  limpiarVer1();

  let profesional = hoja.getRange("C3").getValue();
  let dia = hoja.getRange("C8").getValue();
  let mes = hoja.getRange("C11").getValue();
  let anio = hoja.getRange("C14").getValue();

  if (profesional != "" && dia != "" && mes != "" && anio != "") {
    let baseDeDatosAgenda = SpreadsheetApp.openById(idBaseDeDatosAgenda)
      .getSheetByName(hojaBaseDeDatosAgenda)
      .getDataRange()
      .getValues();

    let baseDeDatosAgendaFiltro = baseDeDatosAgenda.filter((fila) => {
      let profesionalAgenda = fila[4];
      let diaAgenda = fila[1];
      let mesAgenda = fila[2];
      let anioAgenda = fila[3];
      return (
        profesional === profesionalAgenda &&
        dia === diaAgenda &&
        mes === mesAgenda &&
        anio === anioAgenda
      );
    });

    let horasLaborales = [
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

    if (baseDeDatosAgendaFiltro.length > 0) {
      let ordenHoras = {};

      horasLaborales.forEach((hora, index) => {
        ordenHoras[hora] = index;
      });

      baseDeDatosAgendaFiltro.sort((a, b) => {
        return ordenHoras[a[6]] - ordenHoras[b[6]];
      });

      let horasNoDisponibles = baseDeDatosAgendaFiltro.map((fila) => fila[6]);

      let horasDisponibles = horasLaborales.filter((hora) => {
        return !horasNoDisponibles.includes(hora);
      });

      let listaHorasDisponible = SpreadsheetApp.newDataValidation()
        .requireValueInList(horasDisponibles)
        .setAllowInvalid(false)
        .build();

      hoja
        .getRange("G9")
        .setDataValidation(listaHorasDisponible)
        .setBackground("cyan");

      baseDeDatosAgendaFiltro1 = baseDeDatosAgendaFiltro.map((fila) => {
        return [fila[6], fila[7]];
      });

      baseDeDatosAgendaFiltro2 = baseDeDatosAgendaFiltro.map((fila) => {
        return [fila[8], fila[9], fila[10]];
      });

      if (
        baseDeDatosAgendaFiltro1.length > 0 &&
        baseDeDatosAgendaFiltro2.length > 0
      ) {
        hoja
          .getRange(15, 6, baseDeDatosAgendaFiltro1.length, 2)
          .setValues(baseDeDatosAgendaFiltro1)
          .setBackground("Lavender");
        hoja
          .getRange(15, 9, baseDeDatosAgendaFiltro2.length, 3)
          .setValues(baseDeDatosAgendaFiltro2)
          .setBackground("Lavender");
      } else {
        hoja
          .getRange("J15")
          .setValue("Libre, hay cita(s) cancelada(s).")
          .setBackground("Khaki");
      }
    } else {
      let listaHorasDisponible = SpreadsheetApp.newDataValidation()
        .requireValueInList(horasLaborales)
        .setAllowInvalid(false)
        .build();
      hoja
        .getRange("G9")
        .setDataValidation(listaHorasDisponible)
        .setBackground("HotPink");
    }

    hojaInfo.getRange("J2").setValue(1);
  } else {
    ui.alert(
      `Error: profesional: ${profesional} - dia: ${dia} - mes: ${mes} - año: ${anio}`,
    );
  }
}

/** */ function limpiarVer1 /** */() {
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Agenda_1");
  hoja.getRange("G3").setValue("").setBackground("white");
  hoja.getRange("G6").setValue("").setBackground("white");
  hoja.getRange("G9").setValue("").setBackground("white");
  hoja.getRange("F15:G").setValue("").setBackground("white");
  hoja.getRange("I15:K").setValue("").setBackground("white");
  hoja.getRange("L3").setValue(false).setBackground("white");
  hoja
    .getRange("K8")
    .setValue("Escribe una observacion...")
    .setBackground("white");
}

/** */ function agendar1 /** */() {
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Agenda_1");
  let hojaInfo = libro.getSheetByName("Info");
  let ui = SpreadsheetApp.getUi();
  let bd = SpreadsheetApp.openById(idBaseDeDatosAgenda).getSheetByName(
    hojaBaseDeDatosAgenda,
  );

  let bandera = hojaInfo.getRange("J2").getValue();
  let celular = hoja.getRange("G3").getValue();
  let bandera2 = false;

  if (celular != "") {
    let resultado = busquedaBinaria(celular);
    if (resultado[0]) {
      bandera2 = true;
    }
  }

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

    if (
      profesional != "" &&
      dia != "" &&
      mes != "" &&
      anio != "" &&
      celular != "" &&
      procedimiento != "" &&
      hora != "" &&
      observacion != ""
    ) {
      let ultimaFila = bd.getLastRow() + 1;
      let fechaInicio = new Date(2025, 2, 1);
      let fechaActual = new Date();
      let id = Math.floor(
        (fechaActual.getTime() - fechaInicio.getTime()) / 1000,
      );
      let correo = Session.getActiveUser().getEmail();
      if (preferencia) {
        preferencia = "SI";
      } else {
        preferencia = "NO";
      }
      let estado = "Agendada";

      let escritura = [
        id,
        dia,
        mes,
        anio,
        profesional,
        area,
        hora,
        celular,
        procedimiento,
        observacion,
        preferencia,
        estado,
        correo,
        fechaActual,
      ];

      bd.getRange(ultimaFila, 1, 1, escritura.length).setValues([escritura]);

      apiWhatsApp(celular, area, dia, hora);

      limpiarVer1();
      verAgenda1();
      hojaInfo.getRange("J2").setValue(0);
    } else {
      ui.alert("Faltan datos por llenar.");
    }
  } else {
    verAgenda1();
    ui.alert(
      "El celular no esta registrado o error al visualizar agenda, cita NO asignada.",
    );
  }
}

/** */ function bloquear1 /** */() {
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Agenda_1");
  let bd = SpreadsheetApp.openById(idBaseDeDatosAgenda).getSheetByName(
    hojaBaseDeDatosAgenda,
  );
  let ui = SpreadsheetApp.getUi();
  let codigo = ui.prompt("Ingrese el codigo.");
  let respuesta = codigo.getResponseText();
  let codigoCorrecto = "qwe123+";
  let hojaInfo = libro.getSheetByName("Info");

  if (codigo.getSelectedButton() === ui.Button.OK) {
    if (respuesta === codigoCorrecto) {
      let profesional = hoja.getRange("C3").getValue();
      let area = hoja.getRange("C6").getValue();
      let dia = hoja.getRange("C8").getValue();
      let mes = hoja.getRange("C11").getValue();
      let anio = hoja.getRange("C14").getValue();

      if (profesional != "" && dia != "" && mes != "" && anio != "") {
        let horasLaborales = [
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

        let ultimaFila = bd.getLastRow() + 1;
        let fechaInicio = new Date(2025, 2, 1);
        let fechaActual = new Date();
        let id = Math.floor(
          (fechaActual.getTime() - fechaInicio.getTime()) / 1000,
        );
        let correo = Session.getActiveUser().getEmail();
        let preferencia = "NO";
        let estado = "Bloqueada";
        let observacion = `Agenda bloqueada ${fechaActual} por ${correo}`;
        let celular = "NA";
        let procedimiento = "NA";

        let baseDeDatosDeAgendaBloqueada = [];

        horasLaborales.forEach((hora) => {
          let escritura = [
            id,
            dia,
            mes,
            anio,
            profesional,
            area,
            hora,
            celular,
            procedimiento,
            observacion,
            preferencia,
            estado,
            correo,
            fechaActual,
          ];
          baseDeDatosDeAgendaBloqueada.push(escritura);
        });

        bd.getRange(
          ultimaFila,
          1,
          baseDeDatosDeAgendaBloqueada.length,
          baseDeDatosDeAgendaBloqueada[0].length,
        ).setValues(baseDeDatosDeAgendaBloqueada);

        ui.alert("¡Agenda bloqueada!");
        verAgenda1();
        hojaInfo.getRange("J2").setValue(0);
      } else {
        ui.alert("Faltan datos por llenar.");
      }
    } else {
      ui.alert("¡Codigo incorrecto!, intentar nuevamente.");
    }
  }
}

/** */ function verAgenda2 /** */() {
  let ui = SpreadsheetApp.getUi();
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Agenda_2");
  let hojaInfo = libro.getSheetByName("Info");

  limpiarVer2();

  let profesional = hoja.getRange("C3").getValue();
  let dia = hoja.getRange("C8").getValue();
  let mes = hoja.getRange("C11").getValue();
  let anio = hoja.getRange("C14").getValue();

  if (profesional != "" && dia != "" && mes != "" && anio != "") {
    let baseDeDatosAgenda = SpreadsheetApp.openById(idBaseDeDatosAgenda)
      .getSheetByName(hojaBaseDeDatosAgenda)
      .getDataRange()
      .getValues();

    let baseDeDatosAgendaFiltro = baseDeDatosAgenda.filter((fila) => {
      let profesionalAgenda = fila[4];
      let diaAgenda = fila[1];
      let mesAgenda = fila[2];
      let anioAgenda = fila[3];
      return (
        profesional === profesionalAgenda &&
        dia === diaAgenda &&
        mes === mesAgenda &&
        anio === anioAgenda
      );
    });

    let horasLaborales = [
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

    if (baseDeDatosAgendaFiltro.length > 0) {
      let ordenHoras = {};

      horasLaborales.forEach((hora, index) => {
        ordenHoras[hora] = index;
      });

      baseDeDatosAgendaFiltro.sort((a, b) => {
        return ordenHoras[a[6]] - ordenHoras[b[6]];
      });

      let horasNoDisponibles = baseDeDatosAgendaFiltro.map((fila) => fila[6]);

      let horasDisponibles = horasLaborales.filter((hora) => {
        return !horasNoDisponibles.includes(hora);
      });

      let listaHorasDisponible = SpreadsheetApp.newDataValidation()
        .requireValueInList(horasDisponibles)
        .setAllowInvalid(false)
        .build();

      hoja
        .getRange("G9")
        .setDataValidation(listaHorasDisponible)
        .setBackground("cyan");

      baseDeDatosAgendaFiltro1 = baseDeDatosAgendaFiltro.map((fila) => {
        return [fila[6], fila[7]];
      });

      baseDeDatosAgendaFiltro2 = baseDeDatosAgendaFiltro.map((fila) => {
        return [fila[8], fila[9], fila[10]];
      });

      if (
        baseDeDatosAgendaFiltro1.length > 0 &&
        baseDeDatosAgendaFiltro2.length > 0
      ) {
        hoja
          .getRange(15, 6, baseDeDatosAgendaFiltro1.length, 2)
          .setValues(baseDeDatosAgendaFiltro1)
          .setBackground("Lavender");
        hoja
          .getRange(15, 9, baseDeDatosAgendaFiltro2.length, 3)
          .setValues(baseDeDatosAgendaFiltro2)
          .setBackground("Lavender");
      } else {
        hoja
          .getRange("J15")
          .setValue("Libre, hay cita(s) cancelada(s).")
          .setBackground("Khaki");
      }
    } else {
      let listaHorasDisponible = SpreadsheetApp.newDataValidation()
        .requireValueInList(horasLaborales)
        .setAllowInvalid(false)
        .build();
      hoja
        .getRange("G9")
        .setDataValidation(listaHorasDisponible)
        .setBackground("HotPink");
    }

    hojaInfo.getRange("J3").setValue(1);
  } else {
    ui.alert(
      `Error: profesional: ${profesional} - dia: ${dia} - mes: ${mes} - año: ${anio}`,
    );
  }
}

/** */ function limpiarVer2 /** */() {
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Agenda_2");
  hoja.getRange("G3").setValue("").setBackground("white");
  hoja.getRange("G6").setValue("").setBackground("white");
  hoja.getRange("G9").setValue("").setBackground("white");
  hoja.getRange("F15:G").setValue("").setBackground("white");
  hoja.getRange("I15:K").setValue("").setBackground("white");
  hoja.getRange("L3").setValue(false).setBackground("white");
  hoja
    .getRange("K8")
    .setValue("Escribe una observacion...")
    .setBackground("white");
}

/** */ function agendar2 /** */() {
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Agenda_2");
  let hojaInfo = libro.getSheetByName("Info");
  let ui = SpreadsheetApp.getUi();
  let bd = SpreadsheetApp.openById(idBaseDeDatosAgenda).getSheetByName(
    hojaBaseDeDatosAgenda,
  );

  let bandera = hojaInfo.getRange("J3").getValue();
  let celular = hoja.getRange("G3").getValue();
  let bandera2 = false;

  if (celular != "") {
    let resultado = busquedaBinaria(celular);
    if (resultado[0]) {
      bandera2 = true;
    }
  }

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

    if (
      profesional != "" &&
      dia != "" &&
      mes != "" &&
      anio != "" &&
      celular != "" &&
      procedimiento != "" &&
      hora != "" &&
      observacion != ""
    ) {
      let ultimaFila = bd.getLastRow() + 1;
      let fechaInicio = new Date(2025, 2, 1);
      let fechaActual = new Date();
      let id = Math.floor(
        (fechaActual.getTime() - fechaInicio.getTime()) / 1000,
      );
      let correo = Session.getActiveUser().getEmail();
      if (preferencia) {
        preferencia = "SI";
      } else {
        preferencia = "NO";
      }
      let estado = "Agendada";

      let escritura = [
        id,
        dia,
        mes,
        anio,
        profesional,
        area,
        hora,
        celular,
        procedimiento,
        observacion,
        preferencia,
        estado,
        correo,
        fechaActual,
      ];

      bd.getRange(ultimaFila, 1, 1, escritura.length).setValues([escritura]);

      apiWhatsApp(celular, area, dia, hora);

      limpiarVer2();
      verAgenda2();
      hojaInfo.getRange("J3").setValue(0);
    } else {
      ui.alert("Faltan datos por llenar.");
    }
  } else {
    verAgenda2();
    ui.alert(
      "El celular no esta registrado o error al visualizar agenda, cita NO asignada.",
    );
  }
}

/** */ function bloquear2 /** */() {
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Agenda_2");
  let bd = SpreadsheetApp.openById(idBaseDeDatosAgenda).getSheetByName(
    hojaBaseDeDatosAgenda,
  );
  let ui = SpreadsheetApp.getUi();
  let codigo = ui.prompt("Ingrese el codigo.");
  let respuesta = codigo.getResponseText();
  let codigoCorrecto = "qwe123+";
  let hojaInfo = libro.getSheetByName("Info");

  if (codigo.getSelectedButton() === ui.Button.OK) {
    if (respuesta === codigoCorrecto) {
      let profesional = hoja.getRange("C3").getValue();
      let area = hoja.getRange("C6").getValue();
      let dia = hoja.getRange("C8").getValue();
      let mes = hoja.getRange("C11").getValue();
      let anio = hoja.getRange("C14").getValue();

      if (profesional != "" && dia != "" && mes != "" && anio != "") {
        let horasLaborales = [
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

        let ultimaFila = bd.getLastRow() + 1;
        let fechaInicio = new Date(2025, 2, 1);
        let fechaActual = new Date();
        let id = Math.floor(
          (fechaActual.getTime() - fechaInicio.getTime()) / 1000,
        );
        let correo = Session.getActiveUser().getEmail();
        let preferencia = "NO";
        let estado = "Bloqueada";
        let observacion = `Agenda bloqueada ${fechaActual} por ${correo}`;
        let celular = "NA";
        let procedimiento = "NA";

        let baseDeDatosDeAgendaBloqueada = [];

        horasLaborales.forEach((hora) => {
          let escritura = [
            id,
            dia,
            mes,
            anio,
            profesional,
            area,
            hora,
            celular,
            procedimiento,
            observacion,
            preferencia,
            estado,
            correo,
            fechaActual,
          ];
          baseDeDatosDeAgendaBloqueada.push(escritura);
        });

        bd.getRange(
          ultimaFila,
          1,
          baseDeDatosDeAgendaBloqueada.length,
          baseDeDatosDeAgendaBloqueada[0].length,
        ).setValues(baseDeDatosDeAgendaBloqueada);

        ui.alert("¡Agenda bloqueada!");
        verAgenda2();
        hojaInfo.getRange("J3").setValue(0);
      } else {
        ui.alert("Faltan datos por llenar.");
      }
    } else {
      ui.alert("¡Codigo incorrecto!, intentar nuevamente.");
    }
  }
}

function btnActualizar1() {
  let ui = SpreadsheetApp.getUi();
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Dia_1");

  let dia = hoja.getRange("E2").getValue();
  let mes = hoja.getRange("C2").getValue();
  let anio = hoja.getRange("A2").getValue();

  let ordenCronologico = [
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

  if (dia != "" && mes != "" && anio != "") {
    hoja.getRange("A6:K").setValue("").setBackground("white");

    let baseDeDatosAgenda = SpreadsheetApp.openById(idBaseDeDatosAgenda)
      .getSheetByName(hojaBaseDeDatosAgenda)
      .getDataRange()
      .getValues();

    let datosDia = baseDeDatosAgenda.filter((fila) => {
      let diaSeleccionado = fila[1];
      let mesSeleccionado = fila[2];
      let anioSeleccionado = fila[3];
      return (
        diaSeleccionado === dia &&
        mesSeleccionado === mes &&
        anioSeleccionado === anio
      );
    });

    if (datosDia.length > 0) {
      datosDia = datosDia
        .map((fila) => {
          return [
            fila[0],
            fila[4],
            fila[5],
            fila[6],
            fila[7],
            fila[8],
            fila[9],
            fila[10],
            fila[11],
          ];
        })
        .sort((a, b) => {
          // Ordenamiento primario: por nombre (columna 0)
          const nombreA = a[1].toString().toLowerCase();
          const nombreB = b[1].toString().toLowerCase();

          if (nombreA < nombreB) {
            return -1;
          }
          if (nombreA > nombreB) {
            return 1;
          }

          // Si los nombres son iguales, entonces aplicamos el ordenamiento secundario: por hora (columna 2)
          const horaA = a[3];
          const horaB = b[3];

          const indiceA = ordenCronologico.indexOf(horaA);
          const indiceB = ordenCronologico.indexOf(horaB);

          if (indiceA < indiceB) {
            return -1;
          }
          if (indiceA > indiceB) {
            return 1;
          }

          // Si tanto el nombre como la hora son iguales, mantenemos su orden relativo
          return 0;
        });

      hoja
        .getRange(6, 1, datosDia.length, datosDia[0].length)
        .setValues(datosDia);
      calendarioColor1();
    } else {
      ui.alert(`El ${dia} no tiene ninguna cita asignada.`);
    }
  } else {
    ui.alert(`Datos incompletos: dia ${dia} - mes ${mes} - año ${anio}`);
  }
}

function calendarioColor1() {
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Dia_1");
  let ultimaFilaCalendario = hoja.getLastRow();
  let datos = hoja.getRange(6, 2, ultimaFilaCalendario - 5, 8).getValues();
  let numeroFilasDatos = datos.length;
  let nombreAnterior = null;

  const coloresPastel = [
    "#FFFFFF", // Blanco
    "#FADADD", // Rosa pálido
    "#E0F7FA", // Azul claro cielo
    "#FFFFE0", // Amarillo claro
    "#D7FFDB", // Verde menta
    "#F3E5F5", // Lavanda pálido
    "#FFF9C4", // Crema
    "#FFE0B2", // Durazno pálido
  ];

  let m = 0;

  for (let i = 0; i < numeroFilasDatos; i++) {
    let nombreActual = datos[i][0];
    let fila = i + 6;
    let rango = hoja.getRange("B" + fila + ":H" + fila);

    if (nombreActual !== nombreAnterior) {
      m += 1;
      if (m === coloresPastel.length) {
        m = 0;
      }
      nombreAnterior = nombreActual;
    }

    rango.setBackground(coloresPastel[m]);
  }
}

function btnCambiarEstado1() {
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Dia_1");
  let bd_agenda_sede = SpreadsheetApp.openById(
    idBaseDeDatosAgenda,
  ).getSheetByName(hojaBaseDeDatosAgenda);

  let baseDeDatosAgenda = bd_agenda_sede.getDataRange().getValues();

  let ultimaFilaCalendario = hoja.getLastRow();
  let datos = hoja.getRange(6, 1, ultimaFilaCalendario - 5, 9).getValues();

  for (let i = 0; i < datos.length; i++) {
    let id = datos[i][0];
    let observacion = datos[i][6];
    let estado = datos[i][8];

    let indice = baseDeDatosAgenda.findIndex((fila) => {
      return id === fila[0];
    });

    if (indice >= 0) {
      let cita = bd_agenda_sede.getRange(indice + 1, 12).getValue();

      if (cita !== estado && cita !== "Cancelada" && cita !== "Bloqueada") {
        if (estado === "Cancelada") {
          bd_agenda_sede.getRange(indice + 1, 7).setValue("");
          bd_agenda_sede
            .getRange(indice + 1, 10)
            .setValue(
              `Cita cancelada: ${new Date()} ${Session.getActiveUser().getEmail()}`,
            );
          bd_agenda_sede.getRange(indice + 1, 12).setValue(estado);
        } else {
          bd_agenda_sede.getRange(indice + 1, 10).setValue(observacion);
          bd_agenda_sede.getRange(indice + 1, 12).setValue(estado);
        }
      }
    } else {
      ui.alert(`Error -> la cita ${id} no se encuentra asignda.`);
    }
  }

  btnActualizar1();
}

function btnWhatsApp1() {
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Dia_1");

  let dia = hoja.getRange("E2").getValue();

  hoja.getRange("K6:K").setValue("").setBackground("white");

  btnActualizar1();

  let ultimaFilaCalendario = hoja.getLastRow();
  let datos = hoja.getRange(6, 2, ultimaFilaCalendario - 5, 8).getValues();

  if (datos.length > 0) {
    for (let fila = 0; fila < datos.length; fila++) {
      let area = datos[fila][1];
      let hora = datos[fila][2];
      let celular = datos[fila][3];
      let estado = datos[fila][7];
      let mensajeBase = `https://api.whatsapp.com/send?phone=+57${celular}&text=`;
      let mensajeTexto = "";

      switch (estado) {
        case "Agendada":
          mensajeTexto = `¡Hola%20Bell@!%20✨%20Espero%20que%20estés%20teniendo%20un%20día%20increíble.%20Te%20escribimos%20para%20confirmar%20tu%20cita%20de%20${area}%20para%20el%20día%20mañana%20*${dia}*%20en%20la%20*SEDE%20DE%20${sedeMensaje}*%20a%20las%20*${hora}*.%20💅🏼👁😊🕓%0A%0A¡Recuerda%20que%20también%20tenemos%20servicios%20de%20*MANOS*,%20*PIES*%20🦶🏼,%20*CEJAS*%20y%20*PESTAÑAS*%20para%20ti!%20👁%20Pregunta%20por%20nuestros%20catálogos%20para%20que%20conozcas%20todos%20nuestros%20servicios.%0A%0ATen%20en%20cuenta:%0A1.%20Si%20vas%20a%20realizar%20el%20pago%20con%20tarjeta%20💳%20de%20débito%20o%20crédito,%20el%20valor%20total%20tendrá%20un%20incremento%20de%20$3.000.%0A2.%20Si%20deseas%20facturación%20electrónica%20de%20tu%20servicio,%20debes%20solicitarla%20al%20momento%20de%20pagar%20en%20caja.💕`;
          break;
        case "Confirmacion 1":
          mensajeTexto = `Hola%20bell@,%20como%20estas?,%20confirmo%20tu%20cita%20con%20tu%20*${area}*%20del%20dia%20de%20hoy%20*${dia}*%20a%20las%20*${hora}*%20en%20la%20sede%20de%20*${sedeMensaje}*?%20🥰💅%0ARecuerda%20que%20te%20esperamos%20con%20mucho%20amor%20y%20empatia.%20Por%20favor,%20contáctanos%20para%20confirmarla.%20¡Gracias!`;
          break;
        case "Confirmacion 2":
          mensajeTexto = `Recuerda%20que%20estas%20pronto%20para%20tu%20cita.%20Te%20estamos%20esperando%20💞🥰`;
          break;
        case "Asistencia":
          mensajeTexto = `¡Hola,%20mi%20corazón!%20Te%20deseamos%20un%20día%20lleno%20de%20luz%20y%20alegría.%20💞🌞%0A%0AEsperamos%20que%20te%20encuentres%20súper%20bien.%20🫶😊%0A%0AQueremos%20saber%20cómo%20te%20fue%20en%20tu%20servicio%20con%20tu%20*${area}*.%20¿Estás%20feliz%20con%20el%20resultado?%20¿Hay%20algo%20que%20podríamos%20mejorar%20para%20tu%20próxima%20visita?%20¡Tu%20opinión%20es%20muy%20valiosa%20para%20nosotros!%20❤💅%0A%0ANos%20encantaría%20recibir%20tu%20calificación%20para%20seguir%20ofreciéndote%20la%20mejor%20experiencia:%0A%0A*¿Cómo%20calificarías%20tu%20servicio?*%0A1.%20😡%20Pésimo%20servicio%0A2.%20🥺%20Mal%20servicio%0A3.%20🫤%20Servicio%20regular%0A4.%20😊%20Buen%20servicio%0A5.%20🥰%20¡Excelente%20servicio!%0A%0A¡Recuerda%20que%20te%20esperamos%20con%20los%20brazos%20abiertos%20para%20tu%20próxima%20cita!%20No%20dudes%20en%20agendar%20cuando%20quieras.%20¡Siempre%20será%20un%20placer%20atenderte!%20❤🫶🏼`;
          break;
        case "Cancelada":
          mensajeTexto = `¡Hola!%20Lamentamos%20informarte%20que%20tu%20cita%20con%20tu%20*${area}*%20para%20el%20día%20*${dia}*%20ha%20sido%20cancelada.%0ADeseas%20reagendar?%0ASi%20tienes%20alguna%20pregunta,%20no%20dudes%20en%20contactarnos.`;
          break;
        default:
          mensajeTexto = `¡Hola!%20Queremos%20informarte%20que%20tu%20cita%20con%20tu%20*${area}*%20para%20el%20día%20*${dia}*%20tiene%20novedades.%0APor%20favor,%20verifica%20los%20detalles%20o%20contáctanos%20para%20más%20información.%0AEstamos%20a%20tu%20disposición.`;
          break;
      }

      let mensaje = mensajeBase + mensajeTexto;

      let formulaHipervinculo = `=HYPERLINK("${mensaje}"; "Enviar mensaje")`;

      hoja.getRange("K" + (fila + 6)).setFormula(formulaHipervinculo);
    }
  } else {
    ui.alert("No hay agenda disponible.");
  }
}

function btnActualizar2() {
  let ui = SpreadsheetApp.getUi();
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Dia_2");

  let dia = hoja.getRange("E2").getValue();
  let mes = hoja.getRange("C2").getValue();
  let anio = hoja.getRange("A2").getValue();

  let ordenCronologico = [
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

  if (dia != "" && mes != "" && anio != "") {
    hoja.getRange("A6:K").setValue("").setBackground("white");

    let baseDeDatosAgenda = SpreadsheetApp.openById(idBaseDeDatosAgenda)
      .getSheetByName(hojaBaseDeDatosAgenda)
      .getDataRange()
      .getValues();

    let datosDia = baseDeDatosAgenda.filter((fila) => {
      let diaSeleccionado = fila[1];
      let mesSeleccionado = fila[2];
      let anioSeleccionado = fila[3];
      return (
        diaSeleccionado === dia &&
        mesSeleccionado === mes &&
        anioSeleccionado === anio
      );
    });

    if (datosDia.length > 0) {
      datosDia = datosDia
        .map((fila) => {
          return [
            fila[0],
            fila[4],
            fila[5],
            fila[6],
            fila[7],
            fila[8],
            fila[9],
            fila[10],
            fila[11],
          ];
        })
        .sort((a, b) => {
          // Ordenamiento primario: por nombre (columna 0)
          const nombreA = a[1].toString().toLowerCase();
          const nombreB = b[1].toString().toLowerCase();

          if (nombreA < nombreB) {
            return -1;
          }
          if (nombreA > nombreB) {
            return 1;
          }

          // Si los nombres son iguales, entonces aplicamos el ordenamiento secundario: por hora (columna 2)
          const horaA = a[3];
          const horaB = b[3];

          const indiceA = ordenCronologico.indexOf(horaA);
          const indiceB = ordenCronologico.indexOf(horaB);

          if (indiceA < indiceB) {
            return -1;
          }
          if (indiceA > indiceB) {
            return 1;
          }

          // Si tanto el nombre como la hora son iguales, mantenemos su orden relativo
          return 0;
        });

      hoja
        .getRange(6, 1, datosDia.length, datosDia[0].length)
        .setValues(datosDia);
      calendarioColor2();
    } else {
      ui.alert(`El ${dia} no tiene ninguna cita asignada.`);
    }
  } else {
    ui.alert(`Datos incompletos: dia ${dia} - mes ${mes} - año ${anio}`);
  }
}

function calendarioColor2() {
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Dia_2");
  let ultimaFilaCalendario = hoja.getLastRow();
  let datos = hoja.getRange(6, 2, ultimaFilaCalendario - 5, 8).getValues();
  let numeroFilasDatos = datos.length;
  let nombreAnterior = null;

  const coloresPastel = [
    "#FFFFFF", // Blanco
    "#FADADD", // Rosa pálido
    "#E0F7FA", // Azul claro cielo
    "#FFFFE0", // Amarillo claro
    "#D7FFDB", // Verde menta
    "#F3E5F5", // Lavanda pálido
    "#FFF9C4", // Crema
    "#FFE0B2", // Durazno pálido
  ];

  let m = 0;

  for (let i = 0; i < numeroFilasDatos; i++) {
    let nombreActual = datos[i][0];
    let fila = i + 6;
    let rango = hoja.getRange("B" + fila + ":H" + fila);

    if (nombreActual !== nombreAnterior) {
      m += 1;
      if (m === coloresPastel.length) {
        m = 0;
      }
      nombreAnterior = nombreActual;
    }

    rango.setBackground(coloresPastel[m]);
  }
}

function btnCambiarEstado2() {
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Dia_2");
  let bd_agenda_sede = SpreadsheetApp.openById(
    idBaseDeDatosAgenda,
  ).getSheetByName(hojaBaseDeDatosAgenda);

  let baseDeDatosAgenda = bd_agenda_sede.getDataRange().getValues();

  let ultimaFilaCalendario = hoja.getLastRow();
  let datos = hoja.getRange(6, 1, ultimaFilaCalendario - 5, 9).getValues();

  for (let i = 0; i < datos.length; i++) {
    let id = datos[i][0];
    let observacion = datos[i][6];
    let estado = datos[i][8];

    let indice = baseDeDatosAgenda.findIndex((fila) => {
      return id === fila[0];
    });

    if (indice >= 0) {
      let cita = bd_agenda_sede.getRange(indice + 1, 12).getValue();

      if (cita !== estado && cita !== "Cancelada" && cita !== "Bloqueada") {
        if (estado === "Cancelada") {
          bd_agenda_sede.getRange(indice + 1, 7).setValue("");
          bd_agenda_sede
            .getRange(indice + 1, 10)
            .setValue(
              `Cita cancelada: ${new Date()} ${Session.getActiveUser().getEmail()}`,
            );
          bd_agenda_sede.getRange(indice + 1, 12).setValue(estado);
        } else {
          bd_agenda_sede.getRange(indice + 1, 10).setValue(observacion);
          bd_agenda_sede.getRange(indice + 1, 12).setValue(estado);
        }
      }
    } else {
      ui.alert(`Error -> la cita ${id} no se encuentra asignda.`);
    }
  }

  btnActualizar2();
}

function btnWhatsApp2() {
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Dia_2");

  let dia = hoja.getRange("E2").getValue();

  hoja.getRange("K6:K").setValue("").setBackground("white");

  btnActualizar2();

  let ultimaFilaCalendario = hoja.getLastRow();
  let datos = hoja.getRange(6, 2, ultimaFilaCalendario - 5, 8).getValues();

  if (datos.length > 0) {
    for (let fila = 0; fila < datos.length; fila++) {
      let area = datos[fila][1];
      let hora = datos[fila][2];
      let celular = datos[fila][3];
      let estado = datos[fila][7];
      let mensajeBase = `https://api.whatsapp.com/send?phone=+57${celular}&text=`;
      let mensajeTexto = "";

      switch (estado) {
        case "Agendada":
          mensajeTexto = `¡Hola%20Bell@!%20✨%20Espero%20que%20estés%20teniendo%20un%20día%20increíble.%20Te%20escribimos%20para%20confirmar%20tu%20cita%20de%20${area}%20para%20el%20día%20mañana%20*${dia}*%20en%20la%20*SEDE%20DE%20${sedeMensaje}*%20a%20las%20*${hora}*.%20💅🏼👁😊🕓%0A%0A¡Recuerda%20que%20también%20tenemos%20servicios%20de%20*MANOS*,%20*PIES*%20🦶🏼,%20*CEJAS*%20y%20*PESTAÑAS*%20para%20ti!%20👁%20Pregunta%20por%20nuestros%20catálogos%20para%20que%20conozcas%20todos%20nuestros%20servicios.%0A%0ATen%20en%20cuenta:%0A1.%20Si%20vas%20a%20realizar%20el%20pago%20con%20tarjeta%20💳%20de%20débito%20o%20crédito,%20el%20valor%20total%20tendrá%20un%20incremento%20de%20$3.000.%0A2.%20Si%20deseas%20facturación%20electrónica%20de%20tu%20servicio,%20debes%20solicitarla%20al%20momento%20de%20pagar%20en%20caja.💕`;
          break;
        case "Confirmacion 1":
          mensajeTexto = `Hola%20bell@,%20como%20estas?,%20confirmo%20tu%20cita%20con%20tu%20*${area}*%20del%20dia%20de%20hoy%20*${dia}*%20a%20las%20*${hora}*%20en%20la%20sede%20de%20*${sedeMensaje}*?%20🥰💅%0ARecuerda%20que%20te%20esperamos%20con%20mucho%20amor%20y%20empatia.%20Por%20favor,%20contáctanos%20para%20confirmarla.%20¡Gracias!`;
          break;
        case "Confirmacion 2":
          mensajeTexto = `Recuerda%20que%20estas%20pronto%20para%20tu%20cita.%20Te%20estamos%20esperando%20💞🥰`;
          break;
        case "Asistencia":
          mensajeTexto = `¡Hola,%20mi%20corazón!%20Te%20deseamos%20un%20día%20lleno%20de%20luz%20y%20alegría.%20💞🌞%0A%0AEsperamos%20que%20te%20encuentres%20súper%20bien.%20🫶😊%0A%0AQueremos%20saber%20cómo%20te%20fue%20en%20tu%20servicio%20con%20tu%20*${area}*.%20¿Estás%20feliz%20con%20el%20resultado?%20¿Hay%20algo%20que%20podríamos%20mejorar%20para%20tu%20próxima%20visita?%20¡Tu%20opinión%20es%20muy%20valiosa%20para%20nosotros!%20❤💅%0A%0ANos%20encantaría%20recibir%20tu%20calificación%20para%20seguir%20ofreciéndote%20la%20mejor%20experiencia:%0A%0A*¿Cómo%20calificarías%20tu%20servicio?*%0A1.%20😡%20Pésimo%20servicio%0A2.%20🥺%20Mal%20servicio%0A3.%20🫤%20Servicio%20regular%0A4.%20😊%20Buen%20servicio%0A5.%20🥰%20¡Excelente%20servicio!%0A%0A¡Recuerda%20que%20te%20esperamos%20con%20los%20brazos%20abiertos%20para%20tu%20próxima%20cita!%20No%20dudes%20en%20agendar%20cuando%20quieras.%20¡Siempre%20será%20un%20placer%20atenderte!%20❤🫶🏼`;
          break;
        case "Cancelada":
          mensajeTexto = `¡Hola!%20Lamentamos%20informarte%20que%20tu%20cita%20con%20tu%20*${area}*%20para%20el%20día%20*${dia}*%20ha%20sido%20cancelada.%0ADeseas%20reagendar?%0ASi%20tienes%20alguna%20pregunta,%20no%20dudes%20en%20contactarnos.`;
          break;
        default:
          mensajeTexto = `¡Hola!%20Queremos%20informarte%20que%20tu%20cita%20con%20tu%20*${area}*%20para%20el%20día%20*${dia}*%20tiene%20novedades.%0APor%20favor,%20verifica%20los%20detalles%20o%20contáctanos%20para%20más%20información.%0AEstamos%20a%20tu%20disposición.`;
          break;
      }

      let mensaje = mensajeBase + mensajeTexto;

      let formulaHipervinculo = `=HYPERLINK("${mensaje}"; "Enviar mensaje")`;

      hoja.getRange("K" + (fila + 6)).setFormula(formulaHipervinculo);
    }
  } else {
    ui.alert("No hay agenda disponible.");
  }
}

function agendaMensual() {
  let ui = SpreadsheetApp.getUi();
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Mes");

  let anio = hoja.getRange("B2").getValue();
  let mes = hoja.getRange("D2").getValue();
  let profesional = hoja.getRange("H2").getValue();

  let ordenCronologico = [
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

  if (mes != "" && anio != "" && profesional != "") {
    hoja.getRange("A6:I").setValue("").setBackground("white");

    let baseDeDatosAgenda = SpreadsheetApp.openById(idBaseDeDatosAgenda)
      .getSheetByName(hojaBaseDeDatosAgenda)
      .getDataRange()
      .getValues();

    let datosMes = baseDeDatosAgenda.filter((fila) => {
      let mesSeleccionado = fila[2];
      let anioSeleccionado = fila[3];
      let profesionalSeleccionado = fila[4];
      return (
        mesSeleccionado === mes &&
        anioSeleccionado === anio &&
        profesionalSeleccionado === profesional
      );
    });

    if (datosMes.length > 0) {
      datosMes = datosMes
        .map((fila) => {
          return [
            fila[0],
            fila[1],
            fila[5],
            fila[6],
            fila[7],
            fila[8],
            fila[9],
            fila[10],
            fila[11],
          ];
        })
        .sort((a, b) => {
          // Ordenamiento primario: por nombre (columna 0)
          const nombreA = a[1].toString().toLowerCase();
          const nombreB = b[1].toString().toLowerCase();

          if (nombreA < nombreB) {
            return -1;
          }
          if (nombreA > nombreB) {
            return 1;
          }

          // Si los nombres son iguales, entonces aplicamos el ordenamiento secundario: por hora (columna 2)
          const horaA = a[3];
          const horaB = b[3];

          const indiceA = ordenCronologico.indexOf(horaA);
          const indiceB = ordenCronologico.indexOf(horaB);

          if (indiceA < indiceB) {
            return -1;
          }
          if (indiceA > indiceB) {
            return 1;
          }

          // Si tanto el nombre como la hora son iguales, mantenemos su orden relativo
          return 0;
        });

      hoja
        .getRange(6, 1, datosMes.length, datosMes[0].length)
        .setValues(datosMes);
      calendarioColorMes();
    } else {
      ui.alert(
        `El ${mes} con el profesional ${profesional} no tiene ninguna citas asignada.`,
      );
    }
  } else {
    ui.alert(
      `Datos incompletos: profesional ${profesional} - mes ${mes} - año ${anio}`,
    );
  }
}

function calendarioColorMes() {
  let libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName("Mes");
  let ultimaFilaCalendario = hoja.getLastRow();
  let datos = hoja.getRange(6, 2, ultimaFilaCalendario - 5, 8).getValues();
  let numeroFilasDatos = datos.length;
  let diaAnterior = null;

  const coloresPastel = [
    "#FFFFFF", // Blanco
    "#FFE0B2", // Durazno pálido
  ];

  let m = 0;

  for (let i = 0; i < numeroFilasDatos; i++) {
    let diaActual = datos[i][0];
    let fila = i + 6;
    let rango = hoja.getRange("A" + fila + ":I" + fila);

    if (diaActual !== diaAnterior) {
      m += 1;
      if (m === coloresPastel.length) {
        m = 0;
      }
      diaAnterior = diaActual;
    }

    rango.setBackground(coloresPastel[m]);
  }
}
