var activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();

//Crea un evento en el calendario de Alúmnica con el nombre del empleado y el periódo
// de vacaciones que solicito.
function crearEvento_(nombreEvento, fecha) {
  // var fechaFinal = new Date(namedValues["Último día de tus vacaciones"])

  //Busca el calendario de Alúmnica y creo el evento.

  CalendarApp.getCalendarsByName("Alúmnica")[0].createAllDayEvent(
    nombreEvento,
    fecha
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

function DaysBetweenDates(fechaI, fechaF) {
  var diffMilliseconds = fechaF.getTime() - fechaI.getTime();
  // To calculate the no. of days between two dates
  var diffDays = diffMilliseconds / (1000 * 3600 * 24);
  return Math.ceil(diffDays);
}

function GetDataEmpleado(email) {
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

function EmailGoogleFormData(e) {
  if (!e) {
    throw new Error("No hay Evento");
  }
  try {
    if (MailApp.getRemainingDailyQuota() > 0) {
      //fecha de la solicitud

      //Información del evento del Form
      var email = e.namedValues["Email Address"][0];
      var fechaHomeOfficeRaw = new Date(
        e.namedValues['¿Qué día requieres de "Home Office"?'][0]
      );
      var fechaHomeOffice = DarFormatoAFecha(fechaHomeOfficeRaw);
      var fechaSolicitudRaw = new Date();
      var fechaSolicitud = DarFormatoAFecha(fechaSolicitudRaw);

      // Busca información del empleado que utilizo el Form
      var dataEmpleadoArray = GetDataEmpleado(email);
      var nombreEmpleado = dataEmpleadoArray[1];
      var diasRestantes = dataEmpleadoArray[12];
      var emailJefe = dataEmpleadoArray[6];
      var nombreJefe = dataEmpleadoArray[7];
      var fechaFinPeriodoLaboral = DarFormatoAFecha(
        new Date(dataEmpleadoArray[11])
      );

      //Diff en días entre solicitud y fecha home Office
      var diasEntreFechas = DaysBetweenDates(
        fechaSolicitudRaw,
        fechaHomeOfficeRaw
      );

      //Contenido Correo
      var message = "<h2>¡Hola " + nombreEmpleado + "!,</h2><br />";
      var subject =
        "Home Office " +
        fechaHomeOffice +
        "- solicitado por " +
        nombreEmpleado +
        " el " +
        fechaSolicitud;
      var emails =
        email + "," + emailJefe + ",karina@fundacionmanuelmoreno.org";

      if (diasRestantes < 0) {
        subject = "RECHAZADO " + subject;
        message +=
          "<p>Ya tomaste 6 días de Home Office este año laboral." +
          " Si quieres saber más " +
          "<a href='mailto:karina@fundacionmanuelmoreno.org?cc=" +
          emailJefe +
          "&subject=Dudas " +
          subject +
          "'>" +
          'escríbenos aquí</a>. Recuerda que los días de "Home Office" tienen' +
          " que ser aceptados o contarán como falta." +
          " Si tienes alguna duda, consulta las" +
          ' políticas de <a href="https://docs.google.com/document/d/1fhdp-yJOpEDlPVTjnCDIWetxQMODB-4ZLj3wTB_JqGg/preview">"Home Office"</a>.</p>';
        var activeForm = getLinkedForm();
        borrarUltimaRespuestaForm(activeForm);
        // Borra la ultima respuesta en el sheets
        borrarUltimaFilaSheet("Respuestas");
      } else if (diasEntreFechas < 4) {
        subject = "RECHAZADO " + subject;
        message +=
          '<p>Solo puedes pedir "Home Office" con <strong>4 días de anticipación</strong>' +
          " Si quieres saber más " +
          "<a href='mailto:karina@fundacionmanuelmoreno.org?cc=" +
          emailJefe +
          "&subject=Dudas " +
          subject +
          "'>" +
          'escríbenos aquí</a>. Recuerda que los días de "Home Office" tienen' +
          " que ser aceptados o contarán como falta." +
          " Si tienes alguna duda, consulta las" +
          ' políticas de <a href="https://docs.google.com/document/d/1fhdp-yJOpEDlPVTjnCDIWetxQMODB-4ZLj3wTB_JqGg/preview">"Home Office"</a>.</p>';
        var activeForm = getLinkedForm();
        borrarUltimaRespuestaForm(activeForm);
        // Borra la ultima respuesta en el sheets
        borrarUltimaFilaSheet("Respuestas");
      } else if (diasRestantes >= 0) {
        //Fechas
        var mensajeDiasRestantes = "";
        if (diasRestantes === 0) {
          mensajeDiasRestantes +=
            ' Este es tu <strong>último día</strong> de "Home Office" este año laboral.';
        } else if (diasRestantes > 0) {
          var dias =
            diasRestantes === 1
              ? "queda <strong>1 día</strong>"
              : "quedan <strong>" + diasRestantes + " días</strong>";
          mensajeDiasRestantes +=
            "Recuerda que solo te  " +
            dias +
            ' de "Home Office" este año laboral.';
        }

        message =
          message +
          '<p>Quieres hacer "Home Office" el <strong>' +
          fechaHomeOffice +
          "</strong>. " +
          mensajeDiasRestantes +
          " Si quieres modificar o cancelar tu solicitud  " +
          "<a href='mailto:karina@fundacionmanuelmoreno.org?cc=" +
          emailJefe +
          "&subject=MODIFICACIÓN " +
          subject +
          "'>" +
          'escríbenos aquí</a>. Recuerda que los días de "Home Office" tienen' +
          " que ser aceptados o contarán como falta. " +
          "<strong>" +
          nombreJefe +
          "</strong> te dirá tu lista de tareas un día hábil antés de la fecha." +
          " Es tu responsabilidad recordarle a " +
          nombreJefe +
          " que te de la lista de tareas. Si tienes alguna duda, consulta las" +
          ' políticas de <a href="https://docs.google.com/document/d/1fhdp-yJOpEDlPVTjnCDIWetxQMODB-4ZLj3wTB_JqGg/preview">"Home Office"</a>.</p>';

        crearEvento_("Home Office " + nombreEmpleado, fechaHomeOfficeRaw);
      }
      return MailApp.sendEmail(emails, subject, "", { htmlBody: message });
    }
  } catch (error) {
    Logger.log(error.toString());
  }
}
