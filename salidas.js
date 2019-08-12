var activeSS = SpreadsheetApp.getActiveSpreadsheet();
var emailAdmin = "karina@fundacionmanuelmoreno.org";

//Le da formato dd/mm/aaaa a un Date()
function darFormatoFecha(fecha) {
  var date = fecha;
  var dd = String(date.getDate());
  var mm = String(date.getMonth() + 1);
  var aaaa = date.getFullYear();
  return dd + "/" + mm + "/" + aaaa;
}

function getDataEmpleado(email) {
  var sheetDataEmpleados = activeSS.getSheetByName("DATA");
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
  try {
    if (MailApp.getRemainingDailyQuota() > 0) {
      //INFO SOLICITUD
      ////fecha
      var fechaSolicitudRaw = new Date();
      var fechaSolicitud = darFormatoFecha(fechaSolicitudRaw);

      ////data del evento de google forms
      var emailEmpleado = e.namedValues["Email Address"][0];
      var fechaActividadRaw = new Date(
        e.namedValues["Fecha de la actividad:"][0]
      );
      var fechaActividad = darFormatoFecha(fechaActividadRaw);
      var lugar = e.namedValues["Lugar:"][0];
      var motivo = e.namedValues["Motivo:"][0];

      ////data empleado
      var dataEmpleado = getDataEmpleado(emailEmpleado);
      var nombreEmpleado = dataEmpleado[1];
      var emailJefe = dataEmpleado[6];

      //contenido email
      var emails = [emailEmpleado, emailJefe, emailAdmin].join(",");
      var subject = nombreEmpleado + " saldrá el " + fechaActividad;

      var body = "<h3>¡Hola " + nombreEmpleado + "!</h3>";
      body += "<p>Registraste una actividad fuera de las oficinas de alúmnica ";
      body += "para el día <b>" + fechaActividad + "</b> en ";
      body += "<b>" + lugar + "</b> para ";
      body += "<b>" + motivo + "</b>. Cualquier cambio o cancelación ";
      body += "escribir a <a href='mailto:" + emailAdmin + "?subject=";
      body += "CAMBIO " + subject + "'>" + emailAdmin + "</a>.</p>";

      return MailApp.sendEmail({
        to: emails,
        subject: subject,
        htmlBody: body
      });
    }
  } catch (error) {
    Logger.log(error.toString());
  }
}
