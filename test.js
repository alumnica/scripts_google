var ss = SpreadsheetApp.getActiveSpreadsheet();

/**
 * Suma los valores del resultado de un solo test de un solo alumno.
 *
 * @param {Resultados!B2:R2}  range El rango de celdas que contienen el resultado del test de un alumno.
 * @param {"kolb"}  test_name El nombre del test que se utilizó.
 * @returns                        The time the reference was last changed.
 * @customfunction
 */
function sumaResultadosTest(range) {
  return sumaValores(range[0]);
}

function sacarEstilo(range, test_name) {
  var test = getTest(test_name);
  if (test === "error") {
    return 'Solo puedes usar "kolb" o "alumnica" como nombre del test';
  }
  var testKeys = test[0].slice(2);
  var testValuesKeys = [];
  for (var i = 0; i < testKeys.length; i++) {
    testValuesKeys.push([testKeys[i], range[0][i]]);
  }
  //Sort big to small
  testValuesKeys.sort(function(a, b) {
    return b[1] - a[1];
  });

  if (testValuesKeys[0][1] !== testValuesKeys[1][1]) {
  } else {
  }
  var result = [];
  for (var i = 0; i < testValuesKeys.length; i++) {
    var style_result = testValuesKeys[i];
    result.push(style_result[0]);
    result.push(style_result[1]);
  }

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
  return [[a, b, c, d, "---->"]];
}

function test() {
  sumaResultadosTest(1, "kolb");
}
