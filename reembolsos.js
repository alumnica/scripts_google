var activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();

function borrarUltimaFilaSheet(sheetName) {
  var sheet = activeSpreadSheet.getSheetByName(sheetName);
  sheet.deleteRow(sheet.getLastRow());
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
function getLinks(stringComprobantes) {
  var arrayComprobantes = stringComprobantes.split(", ");
  var links = "";
  for (var i = 0; i < arrayComprobantes.length; i++) {
    links +=
      i === 0
        ? "<a href='" + arrayComprobantes[i] + "'>" + (i + 1) + "</a>"
        : ", <a href='" + arrayComprobantes[i] + "'>" + (i + 1) + "</a>";
  }
  return links;
}

function DiasSolicitados(email) {
  var sheet = activeSpreadSheet.getSheetByName("DíasSolicitados");
  var dataSolicitudes = sheet.getDataRange().getValues();
  var solicitudesEmpleados = dataSolicitudes.filter(function(row) {
    return row[1] === email;
  });
  return solicitudesEmpleados[solicitudesEmpleados.length - 1][6];
}
function EmailGoogleFormData(e) {
  if (!e) {
    throw new Error("No hay Evento");
  }
  try {
    if (MailApp.getRemainingDailyQuota() > 0) {
      //fecha de la solicitud
      var fechaSolicitud = DarFormatoAFecha(new Date());

      //Información del evento del Form
      var email = e.namedValues["Email Address"][0];
      var importe = e.namedValues["Importe"][0];
      var fechaGasto = DarFormatoAFecha(new Date(e.namedValues["Fecha"][0]));
      var tipoGasto = e.namedValues["¿Cuál fue el gasto?"][0];
      var areaGasto = e.namedValues["¿Para qué fue el gasto?"][0];
      var comoSePago = e.namedValues["¿Cómo lo pagaste?"][0];
      var tipoComprobante = e.namedValues["¿Qué comprobante obtuviste?"][0];
      var comprobantes =
        e.namedValues["Favor de adjuntar los comprobantes."][0];

      // Busca información del empleado que utilizo el Form
      var infoEmpleadoArray = GetInfoEmpleado(email);
      var nombreEmpleado = infoEmpleadoArray[1];
      var emailJefe = infoEmpleadoArray[6];
      var emails = email + ",karina@fundacionmanuelmoreno.org," + emailJefe;
      var linksComprobantes = getLinks(comprobantes);

      //CONTENT EMAIL
      var subject =
        "Solicitud de reembolso de " + nombreEmpleado + " - " + fechaSolicitud;
      var message =
        "<h2>¡Hola " +
        nombreEmpleado +
        "!,</h2>" +
        "<h3>Solicitaste un reembolso con esta información:</h3>" +
        "<b>Fecha del Gasto:</b> " +
        fechaGasto +
        "<br />" +
        "<b>Importe:</b> $" +
        importe +
        " pesos<br />" +
        "<b>Tipo de Gasto:</b> " +
        tipoGasto +
        "<br />" +
        "<b>Area:</b> " +
        areaGasto +
        "<br />" +
        "<b>Forma de Pago:</b> " +
        comoSePago +
        "<br />" +
        "<b>Tipo de Comprobante:</b> " +
        tipoComprobante +
        "<br />" +
        "<b>Comprobantes:</b> " +
        linksComprobantes;

      return MailApp.sendEmail(emails, subject, "", { htmlBody: message });
    }
  } catch (error) {
    Logger.log(error.toString());
  }
}
