//******************************************************************************
//********************************GENERALES*************************************

//El rango de valores seleccionado viene como un array de dos dimensiones
function oneDimArray(datosCrudos) {
  // convierto los datosCrudos de un array2d a uno 1d
  // solo sí se selecciono una solafila
  return !!datosCrudos[1] ? "Error: Selecciona solo una fila" : datosCrudos[0];
}

// Determina si todos los valores son iguales en un array1d
function todosIguales(valores) {
  return valores.every(function(valor) {
    return valor === valores[0];
  });
}
//******************************************************************************
//*********************************VALORES**************************************
//Funciones que utilizan el valor númerico de los tests de aprendizaje

// Calcula del resultado númerico del test las dos capacidades principales
function dosSecundariasIguales(valores) {
  //Ordena los valores de mayor a menor
  var valoresOrdenados = ordenarMayoraMenor(valores);
  //Si el segundo y el tercer valor son iguales no se puede calcular estilo
  return valoresOrdenados[1] === valoresOrdenados[2];
}

// Ordenar un array1d con valores de mayor a menor
function ordenarMayoraMenor(valores) {
  //Duplica Array par ano modificar valores originales
  var valoresOrdenados = valores.slice(0);

  valoresOrdenados.sort(function(a, b) {
    return b - a;
  });

  return valoresOrdenados;
}

// Incorpora el nombre de la capacidad con el valor númerico del test
function asignarValorCapacidades(valores) {
  var capacidadesV = [
    ["a", valores[0]],
    ["b", valores[1]],
    ["c", valores[2]],
    ["d", valores[3]]
  ];

  return capacidadesV;
}

//Ordena las capacidadesV de mayor a menor de acuerdo al valor
function ordenarCapacidadesVMayorMenor(capacidadesV) {
  // Genera una copia del array para no modificar permanentemente el array original
  var capacidadesVOrdenadas = capacidadesV.slice(0);

  capacidadesVOrdenadas.sort(function(a, b) {
    return b[1] - a[1];
  });

  return capacidadesVOrdenadas;
}

//Genera un array 1d con las capacidades quitandole el valor asigando
function quitarValores(capacidadesVOrdenadas) {
  var capacidadesOrdenadas = [
    capacidadesVOrdenadas[0][0],
    capacidadesVOrdenadas[1][0],
    capacidadesVOrdenadas[2][0],
    capacidadesVOrdenadas[3][0]
  ];
  return capacidadesOrdenadas;
}

//Regresa las dos capacidades principales
function separarCapacidadePrincipales(capacidadesOrdenadas) {
  return capacidadesOrdenadas.slice(0, 2);
}

//Determina el estilo de aprendizaje, requiere array capacidades ordenadas sin valor
function determinarEstiloAprendizaje(capacidadesOrdenadas) {
  var capacidadesPrincipales = separarCapacidadePrincipales(
    capacidadesOrdenadas
  );
  if (capacidadesPrincipales.indexOf("a") !== -1) {
    return capacidadesPrincipales.indexOf("d") !== -1
      ? "Acomodador"
      : "Divergente";
  } else {
    return capacidadesPrincipales.indexOf("c") !== -1
      ? "Asimilador"
      : "Convergente";
  }
}

// Determina el estilo de aprendizaje a partir del valor del test
function estiloDeAprendizaje(datosCrudos) {
  var valores = oneDimArray(datosCrudos);
  if (todosIguales(valores)) {
    return "Todos Iguales";
  } else if (dosSecundariasIguales(valores)) {
    return "Dos capacidades secundarias iguales";
  } else {
    var capacidadesV = asignarValorCapacidades(valores);
    var capacidadesVOrdenadas = ordenarCapacidadesVMayorMenor(capacidadesV);
    var soloCapacidadesOrdenadas = quitarValores(capacidadesVOrdenadas);

    return determinarEstiloAprendizaje(soloCapacidadesOrdenadas);
  }
}

//******************************************************************************
//**************************ESTILOS DE APRENDIZAJE******************************
//Funciones que utilizan el estilos de aprendizaje (Resultados finales del test)
//Requiere formato [Test1,Corto1,Test2,Corto2]

function soloResultadosTestsCortos(estilos) {
  var estilosTestCorto = [estilos[1], estilos[3]];
  return estilosTestCorto;
}

function soloResultadosTestsCompletos(estilos) {
  var estilosTestCompleto = [estilos[0], estilos[2]];
  return estilosTestCompleto;
}

//Acepta un set de estilo ya sea del test completo o del test completo
function resultadosSetCompatibles(estilos, tipo) {
  //Completo forma un set de los test completos, cualquier otra cosa lo forma con los test cortos
  var set =
    tipo === "Completo" ? [estilos[0], estilos[2]] : [estilos[1], estilos[3]];
  //Estilos juntos son complementarios, Acomodador y Convergente también (es un ciclo)
  var reglas = ["Acomodador", "Divergente", "Asimilador", "Convergente"];
  //no importa el símbolo
  var diferencia = Math.abs(reglas.indexOf(set[0]) - reglas.indexOf(set[1]));
  return diferencia;
}

function resultadosCompatibles(datosCrudos) {
  var estilos = oneDimArray(datosCrudos);
  var diferenciaCompleto = resultadosSetCompatibles(estilos, "Completo");

  //Test Completo 1 igual a Test Completo 2
  if (diferenciaCompleto === 0) {
    return "Global";
    //Resultados Complementarios
  } else if (diferenciaCompleto === 1 || diferenciaCompleto === 3) {
    return "Compatibles";
    //Resultados Contradictorios
  } else {
    return "Incompatibles";
  }
}
//MIENTRAS SE AUTOMATIZA MÁS EL PROCESO Y SE CAMBIA EL TEST
//preferencia se saca de la ultima pregunta
function estiloFinal(estilos, preferencia) {
  var estilosOneD = oneDimArray(estilos);
  var resultadosTestsCompletos = soloResultadosTestsCompletos(estilosOneD);
  if (preferencia === "Global") {
    return resultadosTestsCompletos[0];
  } else if (preferencia === "Mecanismos") {
    return resultadosTestsCompletos[0];
  } else if (preferencia === "Carácter") {
    return resultadosTestsCompletos[1];
  } else {
    return "Utiliza la última pregunta del formulario para determinar la preferencia del alumno por Mecanismos o Carácter, y escríbelo en la columna Preferencia.";
  }
}
//Experimentos
//Requiere formato [Test1,Corto1,Test2,Corto2,Global,Manual]
function rasgosComplementarios(estilosRaw) {
  var estilos = oneDimArray(estilosRaw);
  var estiloManual = estilos[5];
  for (var i = 0; i < estilos.length - 1; i++) {
    if (estilos[i] === estiloManual) {
      estilos.splice(i, 1);
    }
  }
  return estilos;
}
