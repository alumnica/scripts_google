var activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();

//Crea un evento en el calendario de Alúmnica con el nombre del empleado y el periódo
// de vacaciones que solicito.
function crearEvento_(fechaInicial, fechaFinal, nombreEvento) {
  // var fechaFinal = new Date(namedValues["Último día de tus vacaciones"])
  // Se agrega un día a la fecha final por que CalendarApp hace el evento con un día menos
  var fechaFinalAjuste = new Date(fechaFinal.setDate(fechaFinal.getDate() + 1));
  //Busca el calendario de Alúmnica y creo el evento.
  CalendarApp.getCalendarsByName("Alúmnica")[0].createAllDayEvent(
    nombreEvento,
    fechaInicial,
    fechaFinalAjuste
  );
}

function getLinkedForm() {
  return FormApp.openByUrl(SpreadsheetApp.getActiveSpreadsheet().getFormUrl());
}

function borrarUltimaFilaSheet(sheetName) {
  var sheet = activeSpreadSheet.getSheetByName(sheetName);
  sheet.deleteRow(sheet.getLastRow());
}

function borrarUltimaRespuestaForm(form) {
  var formResponses = form.getResponses();
  var lastformResponseID = formResponses[formResponses.length - 1].getId();
  form.deleteResponse(lastformResponseID);
}

//Le da formato dd/mm/aaaa a un Date()
function DarFormatoAFecha(fecha) {
  var date = fecha;
  var dd = String(date.getDate());
  var mm = String(date.getMonth() + 1);
  var aaaa = date.getFullYear();
  return dd + "/" + mm + "/" + aaaa;
}

function GetInfoEmpleado(email) {
  var sheetDataEmpleados = activeSpreadSheet.getSheetByName("DATA");
  var dataEmpleados = sheetDataEmpleados.getDataRange().getValues();
  var infoEmpleado;
  dataEmpleados.forEach(function(row) {
    if (row[0] === email) {
      infoEmpleado = row;
    }
  });
  return infoEmpleado;
}
function getDataDiasSolicitados(email) {
  var sheet = activeSpreadSheet.getSheetByName("DíasSolicitados");
  var dataSolicitudes = sheet.getDataRange().getValues();
  var solicitudesEmpleados = dataSolicitudes.filter(function(row) {
    return row[1] === email;
  });
  return [
    solicitudesEmpleados[solicitudesEmpleados.length - 1][6],
    solicitudesEmpleados[solicitudesEmpleados.length - 1][8],
    solicitudesEmpleados[solicitudesEmpleados.length - 1][9],
    solicitudesEmpleados[solicitudesEmpleados.length - 1][7]
  ];
}
function EmailGoogleFormData(e) {
  if (!e) {
    throw new Error("No hay Evento");
  }
  try {
    if (MailApp.getRemainingDailyQuota() > 0) {
      //fecha de la solicitud
      var fechaSolicitud = new Date();
      var fechaSolicitudConFormato = DarFormatoAFecha(fechaSolicitud);

      //Información del evento del Form
      var email = e.namedValues["Email Address"][0];
      var fechaInicioVacaciones = new Date(
        e.namedValues["Primer día de tus vacaciones"][0]
      );
      var fechaFinVacaciones = new Date(
        e.namedValues["Último día de tus vacaciones"][0]
      );
      var timeStamp = e.namedValues["Timestamp"][0];

      // Busca información del empleado que utilizo el Form
      var infoEmpleadoArray = GetInfoEmpleado(email);
      var nombreEmpleado = infoEmpleadoArray[1];
      var diasRestantes = infoEmpleadoArray[7];
      var emailJefe = infoEmpleadoArray[8];
      var fechaFinPeriodoLaboral = new Date(infoEmpleadoArray[6]);
      //Fechas
      var fechaInicioVacacionesConFormato = DarFormatoAFecha(
        fechaInicioVacaciones
      );
      var fechaFinVacacionesConFormato = DarFormatoAFecha(fechaFinVacaciones);
      var fechaFinPeriodoLaboralAjuste = new Date(
        fechaFinPeriodoLaboral.setDate(fechaFinPeriodoLaboral.getDate() + 1)
      );
      var fechaFinPeriodoLaboralAjusteConFormato = DarFormatoAFecha(
        fechaFinPeriodoLaboralAjuste
      );
      //INFO extra
      var dataDiasSolicitados = getDataDiasSolicitados(email);
      var diasSolicitados = dataDiasSolicitados[0];
      var fueraAnoLaboral = dataDiasSolicitados[1];
      var entreAnosLaborales = dataDiasSolicitados[2];
      var diasDentro = dataDiasSolicitados[3];
      //EMAIL
      var subject =
        "Solicitud de vacaciones " +
        nombreEmpleado +
        " " +
        fechaSolicitudConFormato;
      var message = "<h2>¡Hola " + nombreEmpleado + "!</h2><br />";
      //Email to:
      var emails = email + "," + emailJefe;
      +",diego@fundacionmanuelmoreno.org" + ",karina@fundacionmanuelmoreno.org";

      if (
        fechaInicioVacaciones > fechaFinVacaciones ||
        fechaSolicitud > fechaInicioVacaciones ||
        fechaSolicitud > fechaFinVacaciones
      ) {
        message =
          message +
          "<p>Llenaste mal tu solicitud de vacaciones. " +
          "Por favor, <a href='https://docs.google.com/forms/d/e/1FAIpQLSdevz-z" +
          "Prvs4f7jRIWBRlNlklomH5D8gHmt4fSXmTXCI26svg/viewform?usp=sf_link'>vuelvela a llenar.</a>";

        var activeForm = getLinkedForm();
        borrarUltimaRespuestaForm(activeForm);
        // Borra la ultima respuesta en el sheets
        borrarUltimaFilaSheet("Respuestas");
        return MailApp.sendEmail(email, "ERROR - Solicitud de vacaciones", "", {
          htmlBody: message
        });
      } else if (diasRestantes < 0 && !fueraAnoLaboral) {
        message =
          message +
          "<p>En tu solicitud de vacaciones pediste " +
          diasRestantes * -1 +
          " días más de los que te corresponden este año laboral." +
          " Cualquier duda o modificación que necesites hacer a tu solicitud" +
          " de vacaciones por favor escribe a:" +
          " <a href='mailto:contacto@alumnica.org?subject=Dudas/modificaciones" +
          " a solicitud de vacaciones " +
          timeStamp +
          "'>contacto@alumnica.org</a>.</p>";

        var activeForm = getLinkedForm();
        borrarUltimaRespuestaForm(activeForm);
        // Borra la ultima respuesta en el sheets
        borrarUltimaFilaSheet("Respuestas");
        return MailApp.sendEmail(emails, subject, "", { htmlBody: message });
      } else if (entreAnosLaborales && diasRestantes >= 0) {
        message =
          message +
          "<p>Pediste <strong>" +
          diasSolicitados +
          " días</strong> de vacaciones para salir del <strong>" +
          fechaInicioVacacionesConFormato +
          "</strong> al <strong>" +
          fechaFinVacacionesConFormato +
          "</strong>. <strong>" +
          diasDentro +
          " de estos días </strong>cuentan para este año laboral pero los otros <strong>" +
          (diasSolicitados - diasDentro) +
          " días </strong>contaran para tu próximo año laboral" +
          ". Recuerda que este año laboral te quedan <strong>" +
          diasRestantes +
          "</strong> días de vacaciones que tienes que tomar antes de <strong>" +
          fechaFinPeriodoLaboralAjusteConFormato +
          "</strong>. Cualquier duda o modificación que necesites " +
          "hacer a tu solicitud de vacaciones por favor escribe a:" +
          " <a href='mailto:contacto@alumnica.org?subject=Dudas/modificaciones " +
          "a solicitud de vacaciones " +
          timeStamp +
          "'>contacto@alumnica.org</a>.</p>";

        crearEvento_(
          fechaInicioVacaciones,
          fechaFinVacaciones,
          "Vacaciones " + nombreEmpleado
        );

        return MailApp.sendEmail(emails, subject, "", { htmlBody: message });
      } else if (diasRestantes >= 0 && !fueraAnoLaboral) {
        message =
          message +
          "<p>Pediste <strong>" +
          diasSolicitados +
          " días</strong> de vacaciones para salir del <strong>" +
          fechaInicioVacacionesConFormato +
          "</strong> al <strong>" +
          fechaFinVacacionesConFormato +
          "</strong>. Recuerda que te quedan <strong>" +
          diasRestantes +
          "</strong> días de vacaciones que tienes que tomar antes de <strong>" +
          fechaFinPeriodoLaboralAjusteConFormato +
          "</strong>. Cualquier duda o modificación que necesites " +
          "hacer a tu solicitud de vacaciones por favor escribe a:" +
          " <a href='mailto:contacto@alumnica.org?subject=Dudas/modificaciones " +
          "a solicitud de vacaciones " +
          timeStamp +
          "'>contacto@alumnica.org</a>.</p>";

        crearEvento_(
          fechaInicioVacaciones,
          fechaFinVacaciones,
          "Vacaciones " + nombreEmpleado
        );

        return MailApp.sendEmail(emails, subject, "", { htmlBody: message });
      } else if (fueraAnoLaboral) {
        message =
          message +
          "<p>Pediste <strong>" +
          diasSolicitados +
          " días</strong> de vacaciones para salir del <strong>" +
          fechaInicioVacacionesConFormato +
          "</strong> al <strong>" +
          fechaFinVacacionesConFormato +
          "</strong>" +
          ". Cualquier duda o modificación que necesites " +
          "hacer a tu solicitud de vacaciones por favor escribe a:" +
          " <a href='mailto:contacto@alumnica.org?subject=Dudas/modificaciones " +
          "a solicitud de vacaciones " +
          timeStamp +
          "'>contacto@alumnica.org</a>.</p>";

        crearEvento_(
          fechaInicioVacaciones,
          fechaFinVacaciones,
          "Vacaciones " + nombreEmpleado
        );

        return MailApp.sendEmail(emails, subject, "", { htmlBody: message });
      }
    }
  } catch (error) {
    Logger.log(error.toString());
  }
}
