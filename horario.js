var ss = SpreadsheetApp.getActiveSpreadsheet();
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu("Alúmnica")
    .addItem("Generar Reporte", "getNetworkHours")
    .addToUi();
}

var holidays = ss
  .getSheetByName("DATA")
  .getDataRange()
  .getValues()
  .filter(function(e) {
    return e[0].length !== 0;
  });
var as = ss.getActiveSheet();

var lastRow = as.getLastRow();
var dataRelojChecadorUsuarios = as.getRange("A5:E" + lastRow).getValues();

var dataRelojChecadorFecha = as.getRange("B2").getValues()[0][0];

function getDates(rawString) {
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
  var dates = getDates(dataRelojChecadorFecha);
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
    generateReport(hours, days);
    as.getRange("J2").setValues([[hours]]);
  }
  return null;
}
function generateReport(hours, days) {
  var tolerancia = (days * 10 * -1) / 60;
  var result = [];
  dataRelojChecadorUsuarios.forEach(function(row) {
    var difHours = row[4] - hours;
    //AMartinez Abraham Martinez hace un día de HomeOffice todos las semanas.
    difHours += row[1] === "AMartinez" ? 7 : 0;
    var rep = difHours < tolerancia ? " ✖" : " ✔";
    result.push([row[1], difHours.toFixed(2) + rep]);
  });
  setResult(result);
}

function setResult(result) {
  as.getRange("F5:G" + lastRow)
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

  as.getRange("F3:G4")
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

  as.autoResizeColumns(5, 8);
}

function test() {
  Logger.log(dataRelojChecadorUsuarios);
}
