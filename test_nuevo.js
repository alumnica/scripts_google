var ss = SpreadsheetApp.getActiveSpreadsheet();
var resultsRaw = ss
  .getSheetByName("Respuestas")
  .getDataRange()
  .getValues();

var testKolb = ss
  .getSheetByName("kolb")
  .getDataRange()
  .getValues();

var testAlumnica = ss
  .getSheetByName("alumnica")
  .getDataRange()
  .getValues();

function getAnswersReference(twoDRespuestasRaw) {
  Logger.log(Object.getOwnPropertyNames(String.prototype));
  Logger.log(Object.getOwnPropertyNames(Array.prototype));
  Logger.log(transpose2DArray(twoDRespuestasRaw));
  return twoDRespuestasRaw[0];
}

function extractAnswersText(array_raw) {
  var answersText = array_raw.map(function(text) {
    var matches = text.match(/\[(.*?)\]/);
    return matches ? matches[1] : text;
  });
  return answersText;
}

function transpose2DArray(twoDArray) {
  return twoDArray[0].map(function(col, i) {
    return twoDArray.map(function(row) {
      return row[i];
    });
  });
}

function makeTestObject(testRawTrans) {
  var test = {};
  for (var i = 0; i < testRawTrans.length; i++) {
    var row = testRawTrans[i]
    test[row[0]] =

  }
}

function allResults() {
  var headerResults = getAnswersReference(resultsRaw);
  var headerAnswers = extractAnswersText(headerResults);
  return [headerAnwsers];
}

function addKeysToSum(keys, sum) {
  var testValuesKeys = [];
  for (var i = 0; i < keys.length; i++) {
    testValuesKeys.push([keys[i], sum[i]]);
  }
  return testValuesKeys;
}

function sortBigToSmall(keysWithValues, result) {
  keysWithValues.sort(function(a, b) {
    return b[1] - a[1];
  });
  for (var i = 0; i < keysWithValues.length; i++) {
    var style_result = keysWithValues[i];
    result.push(style_result[0]);
    result.push(style_result[1]);
  }
  return keysWithValues;
}

function sacarEstilosPrincipales(sortedResults, result) {
  if (sortedResults[0][1] !== sortedResults[1][1]) {
    result.push(sortedResults[0][0]);
    result.push("");
    result.push("");
    result.push("");
  } else if (sortedResults[0][1] === sortedResults[3][1]) {
    result.push(sortedResults[0][0]);
    result.push(sortedResults[1][0]);
    result.push(sortedResults[2][0]);
    result.push(sortedResults[3][0]);
  } else if (
    sortedResults[0][1] === sortedResults[1][1] &&
    sortedResults[1][1] !== sortedResults[2][1]
  ) {
    result.push(sortedResults[0][0]);
    result.push(sortedResults[1][0]);
    result.push("");
    result.push("");
  } else if (
    sortedResults[0][1] === sortedResults[2][1] &&
    sortedResults[2][1] !== sortedResults[3][1]
  ) {
    result.push(sortedResults[0][0]);
    result.push(sortedResults[1][0]);
    result.push(sortedResults[2][0]);
    result.push("");
  }
}

/**
 * Suma los valores del resultado de un solo test de un solo alumno.
 *
 * @param {Resultados!B2:R2}  range El rango de celdas que contienen el resultado del test de un alumno. Selecciona todas las de UN MISMO TEST.
 * @param {"kolb"}  test_name El nombre del test que se utilizó.
 * @returns                        Los estilos con sus valores y los estilos principales.
 * @customfunction
 */
function calcularResultados(range, test_name) {
  var sumRaw = sumaValores(range[0]);
  var result = [test_name + ":"];

  //Gets test DATA to evaluate results
  var test = getTest(test_name);
  if (test === "error") {
    return 'Solo puedes usar "kolb" o "alumnica" como nombre del test';
  }
  var testKeys = test[0].slice(2);

  //Join values with keys and sort them from big to small
  var testValuesKeys = addKeysToSum(testKeys, sumRaw);
  //Sort big to small

  var sortedResults = sortBigToSmall(testValuesKeys, result);

  //para que no se vea tan junto el resultadosCompatibles
  result.push("PRINCIPAL:");

  var estilos = sacarEstilosPrincipales(sortedResults, result);

  result.push("||||");
  return [result];
}

function getTest(test_name) {
  var test;
  switch (test_name) {
    case "kolb":
      test = ss
        .getSheetByName("kolb")
        .getDataRange()
        .getValues();
      break;
    case "alumnica":
      test = ss
        .getSheetByName("alumnica")
        .getDataRange()
        .getValues();
      break;
    default:
      test = "error";
      break;
  }
  return test;
}
function sumaValores(values_array) {
  var a = 0;
  var b = 0;
  var c = 0;
  var d = 0;
  var j = 0;
  for (var i = 0; i < values_array.length; i++) {
    switch (j) {
      case 0:
        a += values_array[i];
        break;
      case 1:
        b += values_array[i];
        break;
      case 2:
        c += values_array[i];
        break;
      case 3:
        d += values_array[i];
        break;
      default:
        break;
    }
    j++;
    if (j === 4) {
      j = 0;
    }
  }
  return [a, b, c, d];
}

function test() {
  var test = getAnswersReference(resultsRaw);
  Logger.log(extractAnswersText(test));
}
