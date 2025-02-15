var ss = SpreadsheetApp.getActiveSpreadsheet();

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
  var principales = [];
  var num;
  if (sortedResults[0][1] !== sortedResults[1][1]) {
    principales.push(sortedResults[0][0]);
    num = 1;
  } else if (sortedResults[0][1] === sortedResults[3][1]) {
    principales.push(sortedResults[0][0]);
    principales.push(sortedResults[1][0]);
    principales.push(sortedResults[2][0]);
    principales.push(sortedResults[3][0]);
    num = 4;
  } else if (
    sortedResults[0][1] === sortedResults[1][1] &&
    sortedResults[1][1] !== sortedResults[2][1]
  ) {
    principales.push(sortedResults[0][0]);
    principales.push(sortedResults[1][0]);
    num = 2;
  } else if (
    sortedResults[0][1] === sortedResults[2][1] &&
    sortedResults[2][1] !== sortedResults[3][1]
  ) {
    principales.push(sortedResults[0][0]);
    principales.push(sortedResults[1][0]);
    principales.push(sortedResults[2][0]);
    num = 3;
  }
  result.push("PRINCIPAL:");
  result.push(num);
  result.push(principales.join(", "));
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
