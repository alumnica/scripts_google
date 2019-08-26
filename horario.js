var holidays = SpreadsheetApp.getActiveSpreadsheet()
  .getSheetByName("DATA")
  .getDataRange()
  .getValues()
  .filter(function(e) {
    return e[0].length !== 0;
  });

function getDates(rawString) {
  var datesStringsArray = rawString.match(/\d{2}[\/]\d{2}\s/gm);
  var yearString = rawString.match(/\d{4}/gm)[0];

  var iDateString = datesStringsArray[0].split("/");
  var fDateString = datesStringsArray[1].split("/");

  var iDate = new Date(yearString, iDateString[0] - 1, iDateString[1], 12);
  var fDate = new Date(yearString, fDateString[0] - 1, fDateString[1], 12);

  return [iDate, fDate];
}

function countCertainDays(d0, d1) {
  var ndays = 1 + Math.round((d1 - d0) / (24 * 3600 * 1000));
  return Math.floor((ndays + ((d0.getDay() + 1) % 7)) / 7);
}

function getNetworkHours(rawString) {
  var dates = getDates(rawString);
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
    var fridays = countCertainDays(startDate, endDate);
    //el viernes se trabajan 3 horas menos
    hours -= fridays * 3;

    holidays.forEach(function(day) {
      if (day[0] >= startDate && day[0] <= endDate) {
        if (day[0].getDay() % 6 != 0) {
          hours -= 7;
          // Si el feriado cae en viernes se quitan 4 horas del total
        } else if (day[0].getDay() === 5) {
          hours -= 4;
        }
      }
    });

    return hours;
  }
  return null;
}
