/**
 * Este script calcula las horas laborales entre fechas y horas específicas en una hoja de cálculo de Google Sheets,
 * excluyendo días festivos y fines de semana, y descontando el tiempo de almuerzo.
 * 
 * Requisitos:
 * 
 * 1. Hojas de Cálculo:
 *    - Deben existir dos hojas de cálculo con los nombres exactos: "HConfig" y "festivos".
 * 
 * 2. Estructura de las Hojas de Cálculo:
 * 
 *    a. HConfig:
 *       Esta hoja contiene la configuración de las filas y columnas que se utilizarán para los cálculos,
 *       así como la hoja donde se encuentran los datos. Los datos deben estar organizados de la siguiente manera:
 * 
 *       | Fecha1 | Hora1 | Fecha2 | Hora2 | Salida | Nombre Hoja       |
 *       |--------|-------|--------|-------|--------|-------------------|
 *       | E      | F     | G      | H     | J      | Seg. Productividad |
 *       | G      | H     | Q      | R     | U      | Seg. Productividad |
 * 
 *    b. festivos:
 *       Esta hoja contiene los días festivos de los próximos dos años. Debe actualizarse periódicamente.
 *       Los datos deben estar organizados de la siguiente manera:
 * 
 *       | Año | Día          | Fecha       | Festividad                | Fecha     |
 *       |-----|--------------|-------------|---------------------------|-----------|
 *       | 2023| 1 de enero   | Martes      | Año Nuevo                 | 01/01/2023|
 *       | 2023| 6 de enero   | Lunes       | Día de los Reyes Magos    | 06/01/2023|
 *       | 2023| 20 de marzo  | Lunes       | Día de San José           | 20/03/2023|
 *       | 2023| 1 de abril   | Domingo     | Domingo de Ramos          | 01/04/2023|
 *       | 2023| 6 de abril   | Jueves      | Jueves Santo              | 06/04/2023|
 *       | 2023| 7 de abril   | Viernes     | Viernes Santo             | 07/04/2023|
 *       | 2023| 9 de abril   | Domingo     | Domingo de Resurrección   | 09/04/2023|
 *       | 2023| 1 de mayo    | Domingo     | Día del Trabajo           | 01/05/2023|
 * 
 * Funcionamiento:
 * 
 * La función `onButtonPress` es la función principal que se ejecuta al presionar un botón en la hoja de cálculo.
 * Esta función lee las configuraciones de la hoja "HConfig" y luego, para cada configuración, lee las fechas y horas de las columnas especificadas,
 * calcula las horas laborales excluyendo los festivos y fines de semana, y finalmente escribe los resultados en la columna especificada.
 * 
 * Las funciones auxiliares se utilizan para convertir y validar fechas y horas, y para realizar los cálculos de las horas laborales.
 * 
 * Nota:
 * Asegúrate de actualizar periódicamente la hoja "festivos" para incluir los días festivos de los próximos años.
 */



// Nombre de la hoja que contiene los días festivos - Variable Global
var Hojafestivos = "festivos";

/**
 * Función principal que se ejecuta al presionar el botón.
 * Esta función calcula las horas laborales entre fechas y horas especificadas en una hoja de cálculo de Google Sheets.
 *
 * @param {string} sheetName - El nombre de la hoja de cálculo.
 * @param {string} fechaInicioColLetra - La letra de la columna de la fecha de inicio.
 * @param {string} horaInicioColLetra - La letra de la columna de la hora de inicio.
 * @param {string} fechaFinColLetra - La letra de la columna de la fecha de fin.
 * @param {string} horaFinColLetra - La letra de la columna de la hora de fin.
 * @param {string} resultadoColLetra - La letra de la columna donde se escribirá el resultado.
 */
function onButtonPress(sheetName, fechaInicioColLetra, horaInicioColLetra, fechaFinColLetra, horaFinColLetra, resultadoColLetra) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);

  // Convertir las letras de las columnas a índices de columna
  var fechaInicioCol = getColumnIndex(fechaInicioColLetra) + 1;
  var horaInicioCol = getColumnIndex(horaInicioColLetra) + 1;
  var fechaFinCol = getColumnIndex(fechaFinColLetra) + 1;
  var horaFinCol = getColumnIndex(horaFinColLetra) + 1;
  var resultadoCol = getColumnIndex(resultadoColLetra) + 1;

  // Obtener la última fila con datos, limitado a un máximo de 3000 filas
  var lastRowWithData = Math.min(3000, sheet.getRange('A:A').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow());

  // Obtener el rango de datos desde la segunda fila hasta la última fila con datos
  var dataRange = sheet.getRange(2, 1, lastRowWithData - 1, sheet.getLastColumn());
  var displayValues = dataRange.getDisplayValues();

  Logger.log(displayValues.length);

  // Obtener los valores de la columna de resultados
  var resultados = sheet.getRange(2, resultadoCol, lastRowWithData - 1, 1).getValues();

  // Recorrer cada fila de datos
  for (var i = 0; i &lt; displayValues.length; i++) {
    Logger.log(i);
    if (!resultados[i][0]) {
      // Convertir las fechas y horas a los formatos adecuados
      var fechaInicio = convertirFormatoFecha(displayValues[i][fechaInicioCol - 1]);
      var horaInicio = formatearHora(displayValues[i][horaInicioCol - 1]);
      var fechaFin = convertirFormatoFecha(displayValues[i][fechaFinCol - 1]);
      var horaFin = formatearHora(displayValues[i][horaFinCol - 1]);

      Logger.log("Fecha inicio: " + fechaInicio);
      Logger.log("Hora inicio: " + horaInicio);
      Logger.log("Fecha Fin: " + fechaFin);
      Logger.log("Hora Fin: " + horaFin);

      // Verificar si las fechas y horas son válidas
      if (fechaInicio &amp;&amp; horaInicio &amp;&amp; fechaFin &amp;&amp; horaFin) {
        var resultado = horasLaborales(fechaInicio, horaInicio, fechaFin, horaFin);
        resultados[i][0] = resultado;
      } else if (fechaInicio == 'Fecha no válida' || fechaFin == 'Fecha no válida') {
        resultados[i][0] = 'errorFecha';
      } else {
        resultados[i][0] = 'sd';
      }

      Logger.log("resultado " + (i + 1) + " : " + resultados[i][0]);
    }
  }

  // Escribir los resultados en la hoja de cálculo
  sheet.getRange(2, resultadoCol, resultados.length, 1).setValues(resultados);
}


/**
 * Convierte el nombre de una columna (letra) en el índice de columna.
 * Esta función convierte un nombre de columna en formato alfabético (por ejemplo, "A", "B", "Z", "AA") en su correspondiente índice numérico
 * utilizado en Google Apps Script, donde "A" es 0, "B" es 1, "Z" es 25, "AA" es 26, y así sucesivamente.
 * 
 * @param {string} columnName - El nombre de la columna en formato alfabético.
 * @returns {number} El índice de la columna, donde "A" es 0.
 * 
 * Ejemplo:
 *   getColumnIndex("A") devuelve 0.
 *   getColumnIndex("B") devuelve 1.
 *   getColumnIndex("Z") devuelve 25.
 *   getColumnIndex("AA") devuelve 26.
 * 
 * Detalles del funcionamiento:
 *   - La función itera sobre cada carácter del nombre de la columna.
 *   - Convierte cada carácter en su correspondiente valor numérico.
 *   - Utiliza un sistema de base 26 para calcular el índice, similar a cómo se calcularía el valor en un sistema numérico de base 26.
 * 
 * Nota:
 *   - El índice devuelto es 0-based, es decir, "A" corresponde a 0, "B" a 1, etc.
 */
function getColumnIndex(columnName) {
  var sum = 0; // Inicializar la suma que contendrá el índice de la columna

  // Iterar sobre cada carácter en el nombre de la columna
  for (var i = 0; i &lt; columnName.length; i++) {
    // Multiplicar la suma actual por 26 para desplazar el valor en el sistema de base 26
    sum *= 26;

    // Convertir el carácter actual en su valor numérico (A=1, B=2, ..., Z=26)
    sum += (columnName[i].charCodeAt(0) - 'A'.charCodeAt(0) + 1);
  }

  // Restar 1 porque los índices en Apps Script empiezan en 0 (A=0, B=1, ...)
  return sum - 1;
}

/**
 * Convierte una fecha en formato texto a formato yyyy-MM-dd.
 * Esta función toma una fecha en formato texto (dd/MM/yyyy) y la convierte a un formato estándar ISO (yyyy-MM-dd) utilizado en muchas aplicaciones y sistemas.
 * Si la fecha es anterior al año 2023 o si el formato es inválido, la función devuelve 'Fecha no válida'.
 * 
 * @param {string} fechaTexto - La fecha en formato texto (dd/MM/yyyy).
 * @returns {string} La fecha en formato yyyy-MM-dd o 'Fecha no válida' si la fecha es inválida.
 * 
 * Ejemplo:
 *   convertirFormatoFecha("01/01/2023") devuelve "2023-01-01".
 *   convertirFormatoFecha("31/12/2022") devuelve "Fecha no válida".
 * 
 * Detalles del funcionamiento:
 *   - La función primero intenta hacer coincidir la fecha con un formato específico utilizando una expresión regular.
 *   - Si la fecha coincide con el formato dd/MM/yyyy, se valida y se convierte a yyyy-MM-dd.
 *   - Si la fecha no coincide con el formato esperado, se intenta parsear directamente utilizando el constructor Date de JavaScript.
 *   - En ambos casos, se verifica que el año sea 2023 o posterior.
 * 
 * Nota:
 *   - La función maneja fechas inválidas devolviendo 'Fecha no válida'.
 *   - Asegúrate de pasar las fechas en el formato correcto (dd/MM/yyyy) para obtener resultados precisos.
 */
function convertirFormatoFecha(fechaTexto) {
  // Definir una expresión regular para validar el formato dd/MM/yyyy
  var regexFecha = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
  
  // Intentar hacer coincidir la fecha con la expresión regular
  var coincidencia = fechaTexto.match(regexFecha);

  // Si la fecha coincide con el formato esperado
  if (coincidencia) {
    var dia = coincidencia[1];
    var mes = coincidencia[2];
    var año = parseInt(coincidencia[3], 10);

    // Validar el año
    if (año &lt; 2023) {
      return 'Fecha no válida';
    }

    // Asegurar que el día y el mes tengan dos dígitos
    dia = dia.length === 1 ? '0' + dia : dia;
    mes = mes.length === 1 ? '0' + mes : mes;

    // Devolver la fecha en formato yyyy-MM-dd
    return año + '-' + mes + '-' + dia;
  } else {
    // Si la fecha no coincide con el formato, intentar parsear directamente
    var fecha = new Date(fechaTexto);
    if (!isNaN(fecha.getTime())) {
      // Verificar que el año sea 2023 o posterior
      var año = fecha.getFullYear();
      if (año &lt; 2023) {
        return 'Fecha no válida';
      }

      // Formatear la fecha a yyyy-MM-dd
      var mes = fecha.getMonth() + 1; // getMonth() devuelve el mes del 0 al 11
      mes = mes &lt; 10 ? '0' + mes : mes;
      var dia = fecha.getDate();
      dia = dia &lt; 10 ? '0' + dia : dia;
      
      return año + '-' + mes + '-' + dia;
    } else {
      // Devolver 'Fecha no válida' si el formato es incorrecto
      return 'Fecha no válida';
    }
  }
}



/**
 * Formatea una hora proporcionada como texto para asegurar que tanto las horas como los minutos
 * tengan dos dígitos. Si el texto proporcionado no está en un formato reconocible de hora
 * (es decir, no contiene al menos una separación de ':'), la función devuelve la entrada original.
 * 
 * @param {string} horaTexto - La hora en formato de texto, típicamente en formato 'H:M' o 'HH:MM'.
 * @returns {string} La hora formateada en formato 'HH:MM'. Si el formato original no es adecuado,
 *                   devuelve el texto original.
 * 
 * Ejemplos:
 *   - formatearHora("9:5") devuelve "09:05".
 *   - formatearHora("23:9") devuelve "23:09".
 *   - formatearHora("12:45") devuelve "12:45".
 *   - formatearHora("textoIncorrecto") devuelve "textoIncorrecto".
 *
 * Detalles del funcionamiento:
 *   - La función intenta dividir el texto de entrada en partes utilizando ':' como delimitador.
 *   - Si el texto se divide correctamente en al menos dos partes (horas y minutos), cada parte es
 *     evaluada y formateada para asegurar que tiene dos dígitos. Esto se logra añadiendo un '0'
 *     delante si es necesario.
 *   - Si el texto no se puede dividir en al menos dos partes, indica que no está en un formato
 *     de hora reconocible y devuelve el texto original.
 */
function formatearHora(horaTexto) {
  var partes = horaTexto.split(':');
  if (partes.length &gt;= 2) {
    var horas = partes[0];
    var minutos = partes[1];

    // Añadir un cero si es necesario para asegurar que ambos, horas y minutos, tengan dos dígitos
    horas = horas.length === 1 ? '0' + horas : horas;
    minutos = minutos.length === 1 ? '0' + minutos : minutos;

    return horas + ':' + minutos;
  } else {
    // Devolver la hora original si no está en el formato esperado
    return horaTexto;
  }
}


/**
 * Calcula las horas laborales entre dos fechas y horas dadas.
 * Esta función toma una fecha y hora de inicio, y una fecha y hora de fin, y calcula la cantidad de horas laborales entre estos dos momentos,
 * excluyendo los días festivos y fines de semana, y descontando el tiempo de almuerzo.
 * 
 * @param {string} fechaInicio - La fecha de inicio en formato yyyy-MM-dd.
 * @param {string} horaInicio - La hora de inicio en formato hh:mm.
 * @param {string} fechaFin - La fecha de fin en formato yyyy-MM-dd.
 * @param {string} horaFin - La hora de fin en formato hh:mm.
 * @returns {string} Las horas laborales en formato h,m o "MesO+" si el intervalo excede un mes.
 * 
 * Ejemplo:
 *   horasLaborales("2023-01-01", "08:00", "2023-01-01", "17:00") devuelve "8,0".
 *   horasLaborales("2023-01-01", "08:00", "2023-01-02", "17:00") devuelve "8,0" excluyendo almuerzo y fines de semana.
 * 
 * Detalles del funcionamiento:
 *   - La función primero verifica si la fecha de inicio y la fecha de fin son el mismo día.
 *   - Si las fechas son el mismo día, llama a `calcularMinutosLaboralesMismoDia` para calcular los minutos laborales en ese día.
 *   - Si las fechas son diferentes, llama a `calcularMinutosLaboralesVariosDias` para calcular los minutos laborales a través de varios días.
 *   - Si el intervalo entre las fechas excede un mes, devuelve "MesO+".
 *   - Convierte los minutos laborales totales a horas y minutos, y los devuelve en formato h,m.
 */
function horasLaborales(fechaInicio, horaInicio, fechaFin, horaFin) {
    // Verificar si la fecha de inicio y la fecha de fin son el mismo día
    var mismoDia = (fechaInicio === fechaFin);
    
    // Crear objetos Date para la fecha y hora de inicio y fin
    var inicio = new Date(fechaInicio + "T" + horaInicio);
    var fin = new Date(fechaFin + "T" + horaFin);

    var totalMinutosLaborales;
    
    // Calcular los minutos laborales dependiendo de si es el mismo día o varios días
    if (mismoDia) {
        totalMinutosLaborales = calcularMinutosLaboralesMismoDia(inicio, fin);
    } else {
        totalMinutosLaborales = calcularMinutosLaboralesVariosDias(inicio, fin);
        if(totalMinutosLaborales == "MesO+"){
          return "MesO+"; // Devolver "MesO+" si el intervalo excede un mes
        }
    }
    
    // Convertir los minutos laborales totales a horas y minutos
    var horasLaborales = totalMinutosLaborales / 60;
    var horas = Math.floor(horasLaborales);
    var minutos = Math.round((horasLaborales - horas) * 60);

    return horas + "," + minutos;
}

/**
 * Descuenta el tiempo de almuerzo del tiempo laboral.
 * Esta función está diseñada para calcular y devolver la cantidad de minutos que el intervalo
 * de tiempo especificado entre 'horarioInicio' y 'horarioFin' se superpone con un periodo fijo
 * de almuerzo entre las 13:00 y las 14:00. Si hay una superposición, calcula los minutos
 * de superposición y los devuelve. 
 
 *
 * @param {Date} horarioInicio - La hora de inicio del trabajo, como objeto Date.
 * @param {Date} horarioFin - La hora de fin del trabajo, como objeto Date.
 * @returns {number} Los minutos de almuerzo descontados si el horario laboral se superpone
 * con el intervalo de almuerzo fijo.
 *
 */

function descontarAlmuerzo(horarioInicio, horarioFin) {
    var inicioAlmuerzo = new Date(horarioInicio.getFullYear(), horarioInicio.getMonth(), horarioInicio.getDate(), 13, 0, 0);
    var finAlmuerzo = new Date(horarioInicio.getFullYear(), horarioInicio.get;
  var horaFinCol = getColumnIndex(horaFinColLetra) + 1;
  var resultadoCol = getColumnIndex(resultadoColLetra) + 1;

  // Encontrar la última fila con datos
  var lastRowWithData = Math.min(3000, sheet.getRange('A:A').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow());

  // Obtener los valores de las celdas relevantes
  var dataRange = sheet.getRange(2, 1, lastRowWithData - 1, sheet.getLastColumn());
  var displayValues = dataRange.getDisplayValues();

  Logger.log(displayValues.length);

  var resultados = sheet.getRange(2, resultadoCol, lastRowWithData - 1, 1).getValues();
  for (var i = 0; i &lt; displayValues.length; i++) {
    Logger.log(i);
    if (!resultados[i][0]) {
      var fechaInicio = convertirFormatoFecha(displayValues[i][fechaInicioCol - 1]);
      var horaInicio = formatearHora(displayValues[i][horaInicioCol - 1]);
      var fechaFin = convertirFormatoFecha(displayValues[i][fechaFinCol - 1]);
      var horaFin = formatearHora(displayValues[i][horaFinCol - 1]);

      Logger.log("Fecha inicio: " + fechaInicio);
      Logger.log("Hora inicio: " + horaInicio);
      Logger.log("Fecha Fin: " + fechaFin);
      Logger.log("Hora Fin: " + horaFin);

      if (fechaInicio &amp;&amp; horaInicio &amp;&amp; fechaFin &amp;&amp; horaFin) {
        var resultado = horasLaborales(fechaInicio, horaInicio, fechaFin, horaFin);
        resultados[i][0] = resultado;
      } else if (fechaInicio == 'Fecha no válida' || fechaFin == 'Fecha no válida') {
        resultados[i][0] = 'errorFecha';
      } else {
        resultados[i][0] = 'sd';
      }

      Logger.log("resultado " + (i + 1) + " : " + resultados[i][0]);
    }
  }

  // Establecer los valores de los resultados en la hoja
  sheet.getRange(2, resultadoCol, resultados.length, 1).setValues(resultados);
}

/**
 * Convierte el nombre de una columna (letra) en su índice correspondiente (número).
 * @param {string} columnName - El nombre de la columna en formato de letra.
 * @returns {number} - El índice de la columna.
 */
function getColumnIndex(columnName) {
  var sum = 0;
  for (var i = 0; i &lt; columnName.length; i++) {
    sum *= 26;
    sum += (columnName[i].charCodeAt(0) - 'A'.charCodeAt(0) + 1);
  }
  return sum - 1; // Restar 1 porque los índices en Apps Script empiezan en 0
}


/**
 * Calcula los minutos laborales entre dos fechas y horas en diferentes días.
 * @param {Date} inicio - La fecha y hora de inicio.
 * @param {Date} fin - La fecha y hora de fin.
 * @returns {number|string} - Los minutos laborales totales o "MesO+" si excede un mes.
 */
function calcularMinutosLaboralesVariosDias(inicio, fin) {
  var diferenciaDias = (fin - inicio) / (1000 * 60 * 60 * 24);
  var diferenciaMeses = diferenciaDias / 30;

  if (diferenciaMeses &gt; 1) {
    return "MesO+";
  }

  var totalMinutosLaborales = 0;
  var fechaInicio = new Date(inicio.getFullYear(), inicio.getMonth(), inicio.getDate());
  var fechaFin = new Date(fin.getFullYear(), fin.getMonth(), fin.getDate());

  for (var dia = new Date(fechaInicio); dia &lt;= fechaFin; dia.setDate(dia.getDate() + 1)) {
    if (!esFestivo(dia)) {
      var diaSemana = dia.getDay();
      if (diaSemana !== 0 &amp;&amp; diaSemana !== 6) {
        var horarioInicio = new Date(dia);
        horarioInicio.setHours(7, 0, 0, 0);
        var horarioFin = new Date(dia);
        horarioFin.setHours(17, 0, 0, 0);

        if (dia.toDateString() === fechaInicio.toDateString()) {
          horarioInicio = new Date(Math.max(inicio, horarioInicio));
        }
        if (dia.toDateString() === fechaFin.toDateString()) {
          horarioFin = new Date(Math.min(fin, horarioFin));
        }

        var minutosDia = (horarioFin - horarioInicio) / (1000 * 60);
        minutosDia = Math.max(0, minutosDia);
        minutosDia -= descontarAlmuerzo(horarioInicio, horarioFin);
        totalMinutosLaborales += minutosDia;
      }
    }
  }

  return totalMinutosLaborales;
}

/**
 * Calcula los minutos laborales entre dos horas en el mismo día.
 * @param {Date} inicio - La fecha y hora de inicio.
 * @param {Date} fin - La fecha y hora de fin.
 * @returns {number} - Los minutos laborales totales en el mismo día.
 */
function calcularMinutosLaboralesMismoDia(inicio, fin) {
  var horarioInicioLaboral = new Date(Math.max(inicio, new Date(inicio.getFullYear(), inicio.getMonth(), inicio.getDate(), 7, 0, 0)));
  var horarioFinLaboral = new Date(Math.min(fin, new Date(fin.getFullYear(), fin.getMonth(), fin.getDate(), 17, 0, 0)));
  var minutosDia = (horarioFinLaboral - horarioInicioLaboral) / (1000 * 60);
  minutosDia -= descontarAlmuerzo(horarioInicioLaboral, horarioFinLaboral);
  return Math.max(0, minutosDia);
}

/**
 * Verifica si una fecha es festiva.
 * @param {Date} dia - La fecha a verificar.
 * @returns {boolean} - Verdadero si la fecha es festiva, falso en caso contrario.
 */
function esFestivo(dia) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Hojafestivos);
  var festivosData = sheet.getDataRange().getValues();
  var fechaDia = Utilities.formatDate(dia, "GMT-0500", "dd/MM/yyyy");

  for (var i = 1; i &lt; festivosData.length; i++) {
    var fechaFestivo = festivosData[i][4];
    if (fechaFestivo) {
      var fechaFestivoFormateada = Utilities.formatDate(fechaFestivo, "GMT-0500", "dd/MM/yyyy");
      if (fechaDia === fechaFestivoFormateada) {
        return true;
      }
    }
  }
  return false;
}
