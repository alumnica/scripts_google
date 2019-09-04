var ss = SpreadsheetApp.getActiveSpreadsheet();
var ui = SpreadsheetApp.getUi();

var holidays = ss
  .getSheetByName("HOLIDAYS")
  .getDataRange()
  .getValues()
  .filter(function(e) {
    return e[0].length !== 0;
  });

var users = ss
  .getSheetByName("USERS")
  .getDataRange()
  .getValues()
  .filter(function(e) {
    return e[0].length !== 0;
  });

var relojData = ss.getSheetByName("Resumen");

var lastRow = relojData.getLastRow();

var timeRange = relojData.getRange("B2").getValues()[0][0];
var dates = getDates(timeRange);

function onOpen() {
  // Or DocumentApp or FormApp.
  ui.createMenu("Alúmnica")
    .addItem("Calcular Horas", "getNetworkHours")
    .addItem("Generar Reportes", "createReportDocuments")
    .addToUi();
}

function formatDateESMX(date) {
  var dias = [
    "domingo",
    "lunes",
    "martes",
    "miércoles",
    "jueves",
    "viernes",
    "sábado"
  ];
  var meses = [
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre"
  ];
  return (
    dias[date.getDay()] +
    " " +
    date.getDate() +
    " de " +
    meses[date.getMonth()] +
    " de " +
    date.getFullYear()
  );
}

function getDates(rawString) {
  if (rawString === "") {
    return;
  }
  var datesStringsArray = rawString.match(/\d{2}[\/]\d{2}\s/gm);
  var yearString = rawString.match(/\d{4}/gm)[0];

  var iDateString = datesStringsArray[0].split("/");
  var fDateString = datesStringsArray[1].split("/");

  var iDate = new Date(yearString, iDateString[0] - 1, iDateString[1], 12);
  var fDate = new Date(yearString, fDateString[0] - 1, fDateString[1], 12);

  return [iDate, fDate];
}

function countCertainDays(days, d0, d1) {
  d0.setHours(12, 0, 0, 0);
  d1.setHours(12, 0, 0, 0);
  var ndays = 1 + Math.round((d1 - d0) / (24 * 3600 * 1000));
  var sum = function(a, b) {
    return a + Math.floor((ndays + ((d0.getDay() + 6 - b) % 7)) / 7);
  };
  return days.reduce(sum, 0);
}

/**
 * Calcula el número de horas de trabajo entre dos fechas utilizando la información de Fecha de la hoja "Resumen de Asistencia" del reloj checador.
 *
 * @param {B2}  rawString La celda que contiene la información del rango de fechas del reloj checador, ej. 2019/08/07 ~ 08/16        ( alumnica )
 * @returns                        El total de horas trabajadas dentro de dos fechas, según el horario de alúmnica
 * @customfunction
 */
function getNetworkHours() {
  //ej. rawString from clock 2019/08/07 ~ 08/16        ( alumnica )

  var startDate = typeof dates[0] == "object" ? dates[0] : new Date(dates[0]);
  var endDate = typeof dates[1] == "object" ? dates[1] : new Date(dates[1]);
  if (endDate > startDate) {
    var days = Math.ceil(
      (endDate.setHours(23, 59, 59, 999) - startDate.setHours(0, 0, 0, 1)) /
        (86400 * 1000)
    );
    var weeks = Math.floor(
      Math.ceil(
        (endDate.setHours(23, 59, 59, 999) - startDate.setHours(0, 0, 0, 1)) /
          (86400 * 1000)
      ) / 7
    );
    days = days - weeks * 2;
    var hours = days * 7;
    hours = startDate.getDay() - endDate.getDay() > 1 ? hours - 14 : hours;
    hours =
      startDate.getDay() == 0 && endDate.getDay() != 6 ? hours - 7 : hours;
    hours =
      endDate.getDay() == 6 && startDate.getDay() != 0 ? hours - 7 : hours;

    var fridays = countCertainDays([5], startDate, endDate);
    //el viernes se trabajan 3 horas menos
    var days = hours / 7;

    hours -= fridays * 3;

    holidays.forEach(function(day) {
      if (day[0] >= startDate && day[0] <= endDate) {
        // Si el feriado cae en viernes se quitan 4 horas del total
        if (day[0].getDay() === 5) {
          hours -= 4;
          days--;
        } else if (day[0].getDay() % 6 != 0) {
          hours -= 7;
          days--;
        }
      }
    });
    generateReport(hours, days, startDate, endDate);
    relojData.getRange("J2").setValues([[hours]]);
  }
  return null;
}
function generateReport(hours, days, startDate, endDate) {
  var dataRelojChecadorUsuarios = relojData
    .getRange("A5:E" + lastRow)
    .getValues();
  var tolerancia = (days * 10 * -1) / 60;
  var result = [];
  dataRelojChecadorUsuarios.forEach(function(row) {
    var difHours = row[4] - hours;
    //AMartinez Abraham Martinez hace un día de HomeOffice todos los lunes.
    var mondays = countCertainDays([1], startDate, endDate);
    difHours += row[1] === "AMartinez" ? 7 * mondays : 0;
    var rep = difHours < tolerancia ? " ✖" : " ✔";
    result.push([row[1], difHours.toFixed(2) + rep]);
  });
  setResult(result);
}

function setResult(result) {
  relojData
    .getRange("F5:G" + lastRow)
    .setValues(result)
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      false,
      "#434343",
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    )
    .setFontFamily("Nunito")
    .setBackground("white")
    .setFontColor("#434343");

  relojData
    .getRange("F3:G4")
    .setValues([["Alúmnica", ""], ["Nombre", "Reporte"]])
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "#434343",
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    )
    .setFontFamily("Comfortaa")
    .setFontColor("white")
    .setBackground("#434343")
    .setWrap(false);

  relojData.autoResizeColumns(5, 8);
}
function createReportDocuments() {
  var response = ui.prompt(
    "¿Cómo quieres nombrar el reporte que se va a generar?",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() == ui.Button.OK) {
    var reportName = response.getResponseText();
  } else {
    return;
  }

  var templateFileID = DriveApp.getFileById(
    "1wXwsL-rtUTocyyG7-0YePrlvIv0mcx81feyebAM7IXk"
  )
    .makeCopy(
      reportName,
      DriveApp.getFolderById("1V1nwC7hAbo3S1aENXSXlvu6HZmz2M5lN")
    )
    .getId();

  if (relojData.getRange("J2").getValues()[0][0] === "") {
    getNetworkHours();
  }
  var doc = DocumentApp.openById(templateFileID);
  var body = doc.getBody();
  body.appendParagraph("");
  body.appendParagraph("");
  body.appendParagraph("");
  body.appendParagraph("");
  body.appendParagraph("");
  body.appendParagraph("");
  body.appendParagraph("");
  body.appendParagraph("");
  body.appendParagraph("");
  body.appendParagraph("");
  body.appendParagraph("");
  body.appendParagraph("");
  body
    .appendParagraph("Reporte de Asistencia")
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body
    .appendParagraph(
      "del " + formatDateESMX(dates[0]) + " al " + formatDateESMX(dates[1])
    )
    .setHeading(DocumentApp.ParagraphHeading.SUBTITLE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendPageBreak();

  var usersInfoAndHours = relojData.getRange("A5:G" + lastRow).getValues();
  var userHours = ss
    .getSheetByName("Registros")
    .getDataRange()
    .getValues();

  var user = 1;

  usersInfoAndHours.forEach(function(row) {
    var userData = users.filter(function(user) {
      return user[5] === row[1];
    })[0];

    var paragraph =
      "Dentro de estas fechas, " +
      userData[1] +
      " " +
      userData[2] +
      " " +
      userData[3] +
      " trabajo " +
      row[4] +
      " horas en las oficinas de Alúmnica como detalla la siguiente tabla:";

    var cells = [["FECHA", "ENTRADA", "SALIDA", "LUGAR"]];

    var currentUserHoursH = [userHours[0 + 3 * user], userHours[2 + 3 * user]];

    var currentUserHoursV = currentUserHoursH[0].map(function(col, i) {
      return currentUserHoursH.map(function(row) {
        return row[i];
      });
    });

    currentUserHoursV.forEach(function(row) {
      if (row[0] === "") {
        return;
      }

      var inOutHours = row[1].match(/\d{2}:\d{2}/gm);
      if (inOutHours) {
        row[1] = inOutHours[0];
        if (inOutHours[1]) {
          row.push(inOutHours[1]);
          row.push("Oficinas Alúmnica");
        } else {
          row.push("");
          row.push("Oficinas Alúmnica");
        }
      } else {
        row.push("");
        row.push("DÍA NO LABORABLE");
      }

      cells.push(row);
    });
    body
      .appendParagraph("Reporte de Asistencia")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    body
      .appendParagraph(
        "del " + formatDateESMX(dates[0]) + " al " + formatDateESMX(dates[1])
      )
      .setHeading(DocumentApp.ParagraphHeading.SUBTITLE)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph("");
    body.appendParagraph("");
    body.appendParagraph(paragraph);
    body.appendParagraph("");
    var table = body.appendTable(cells);

    var styles = {};
    styles[DocumentApp.Attribute.PADDING_TOP] = 0.3;
    styles[DocumentApp.Attribute.PADDING_BOTTOM] = 0.3;
    styles[DocumentApp.Attribute.PADDING_LEFT] = 0.5;
    styles[DocumentApp.Attribute.PADDING_RIGHT] = 0;
    styles[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
      DocumentApp.HorizontalAlignment.CENTER;

    var numberOfRows = table.getNumRows();

    for (var i = 0; i < numberOfRows; i++) {
      // for each column...
      var cols = table.getRow(i).getNumChildren();

      for (var j = 0; j < cols; j++) {
        var cell = table.getRow(i).getCell(j);
        if (i === 0) {
          styles[DocumentApp.Attribute.BOLD] = true;
        }

        // Set the TableCell attributes to each
        cell.setAttributes(styles);

        styles[DocumentApp.Attribute.BOLD] = false;
      }
    }

    body.appendParagraph("");
    body.appendParagraph("");
    body.appendParagraph("");
    body.appendParagraph("");
    body.appendParagraph("");
    body.appendHorizontalRule();

    body
      .appendParagraph(userData[1] + " " + userData[2])
      .setHeading(DocumentApp.ParagraphHeading.SUBTITLE)
      .setIndentStart(0)
      .setIndentFirstLine(0)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    body.appendPageBreak();
    user++;
  });
}

function test() {
  Logger.log(usersInfoAndHours);
}
