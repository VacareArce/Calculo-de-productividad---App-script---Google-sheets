// Nombre de la hoja que contiene los días festivos
var Hojafestivos = "festivos";

/**
 * Función principal que se ejecuta al presionar el botón.
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

  // Obtener la última fila con datos
  var lastRowWithData = Math.min(3000, sheet.getRange('A:A').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow());

  // Obtener el rango de datos
  var dataRange = sheet.getRange(2, 1, lastRowWithData - 1, sheet.getLastColumn());
  var displayValues = dataRange.getDisplayValues();

  Logger.log(displayValues.length);

  // Obtener los valores de la columna de resultados
  var resultados = sheet.getRange(2, resultadoCol, lastRowWithData - 1, 1).getValues();

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

      if (fechaInicio &amp;&amp; horaInicio &amp;&amp; fechaFin &amp;&amp; horaFin) {
        var resultado = horasLaborales(fechaInicio, horaInicio, fechaFin, horaFin);
        resultados[i][0] = resultado;
      } else if (fechaInicio == 'Fecha no válida' || fechaFin == 'Fecha no válida') {
        resultados[i][0] = 'errorFecha';
      } else {
        resultados[i][0] = 'sd';
      }

      Logger.log("resultado " + (i+1) + " : " + resultados[i][0]);
    }
  }

  // Escribir los resultados en la hoja de cálculo
  sheet.getRange(2, resultadoCol, resultados.length, 1).setValues(resultados);
}

/**
 * Convierte el nombre de una columna (letra) en el índice de columna.
 * @param {string} columnName - El nombre de la columna.
 * @returns {number} El índice de la columna.
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
 * Convierte una fecha en formato texto a formato yyyy-MM-dd.
 * @param {string} fechaTexto - La fecha en formato texto (dd/MM/yyyy).
 * @returns {string} La fecha en formato yyyy-MM-dd o 'Fecha no válida' si la fecha es inválida.
 */
function convertirFormatoFecha(fechaTexto) {
  var regexFecha = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/; // Formato dd/MM/yyyy
  var coincidencia = fechaTexto.match(regexFecha);

  if (coincidencia) {
    var dia = coincidencia[1];
    var mes = coincidencia[2];
    var año = parseInt(coincidencia[3], 10);

    // Validación del año
    if (año &lt; 2023) {
      return 'Fecha no válida';
    }

    // Asegurarse de que el día y el mes tengan dos dígitos
    dia = dia.length === 1 ? '0' + dia : dia;
    mes = mes.length === 1 ? '0' + mes : mes;

    return año + '-' + mes + '-' + dia; // Convertir a formato yyyy-MM-dd
  } else {
    // En caso de no coincidir con el formato esperado, se intenta parsear directamente
    var fecha = new Date(fechaTexto);
    if (!isNaN(fecha.getTime())) {
      // Verificar si el año es 2023 o posterior
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
      // Devuelve un valor predeterminado o maneja el error según sea necesario
      return 'Fecha no válida';
    }
  }
}

/**
 * Formatea una hora en formato texto (hh:mm) asegurándose de que tenga dos dígitos para horas y minutos.
 * @param {string} horaTexto - La hora en formato texto.
 * @returns {string} La hora formateada en formato hh:mm.
 */
function formatearHora(horaTexto) {
  var partes = horaTexto.split(':');
  if (partes.length &gt;= 2) {
    var horas = partes[0];
    var minutos = partes[1];

    // Añadir un cero si es necesario
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
 * @param {string} fechaInicio - La fecha de inicio en formato yyyy-MM-dd.
 * @param {string} horaInicio - La hora de inicio en formato hh:mm.
 * @param {string} fechaFin - La fecha de fin en formato yyyy-MM-dd.
 * @param {string} horaFin - La hora de fin en formato hh:mm.
 * @returns {string} Las horas laborales en formato h,m.
 */
function horasLaborales(fechaInicio, horaInicio, fechaFin, horaFin) {
    // Comparar las fechas directamente
    var mismoDia = (fechaInicio === fechaFin);
    var inicio = new Date(fechaInicio + "T" + horaInicio);
    var fin = new Date(fechaFin + "T" + horaFin);

    var totalMinutosLaborales;
    if (mismoDia) {
        totalMinutosLaborales = calcularMinutosLaboralesMismoDia(inicio, fin);
    } else {
        totalMinutosLaborales = calcularMinutosLaboralesVariosDias(inicio, fin);
        if(totalMinutosLaborales == "MesO+"){
          return "MesO+"
        }
    }
   
    var horasLaborales = totalMinutosLaborales / 60;
    var horas = Math.floor(horasLaborales);
    var minutos = Math.round((horasLaborales - horas) * 60);

    return horas + "," + minutos;
}

/**
 * Descuenta el tiempo de almuerzo del tiempo laboral.
 * @param {Date} horarioInicio - La hora de inicio del trabajo.
 * @param {Date} horarioFin - La hora de fin del trabajo.
 * @returns {number} Los minutos de almuerzo descontados.
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
 * Convierte una fecha en formato de texto a un formato estándar.
 * @param {string} fechaTexto - La fecha en formato de texto.
 * @returns {string} - La fecha en formato yyyy-MM-dd o 'Fecha no válida' si es incorrecta.
 */
function convertirFormatoFecha(fechaTexto) {
  var regexFecha = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/; // Formato dd/MM/yyyy
  var coincidencia = fechaTexto.match(regexFecha);

  if (coincidencia) {
    var dia = coincidencia[1];
    var mes = coincidencia[2];
    var año = parseInt(coincidencia[3], 10);

    // Validación del año
    if (año &lt; 2023) {
      return 'Fecha no válida';
    }

    // Asegurarse de que el día y el mes tengan dos dígitos
    dia = dia.length === 1 ? '0' + dia : dia;
    mes = mes.length === 1 ? '0' + mes : mes;

    return año + '-' + mes + '-' + dia; // Convertir a formato yyyy-MM-dd
  } else {
    // En caso de no coincidir con el formato esperado, se intenta parsear directamente
    var fecha = new Date(fechaTexto);
    if (!isNaN(fecha.getTime())) {
      // Verificar si el año es 2023 o posterior
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
      // Devuelve un valor predeterminado o maneja el error según sea necesario
      return 'Fecha no válida';
    }
  }
}

/**
 * Formatea una hora en texto asegurando que tenga dos dígitos para horas y minutos.
 * @param {string} horaTexto - La hora en formato de texto.
 * @returns {string} - La hora formateada en HH:mm.
 */
function formatearHora(horaTexto) {
  var partes = horaTexto.split(':');
  if (partes.length &gt;= 2) {
    var horas = partes[0];
    var minutos = partes[1];

    // Añadir un cero si es necesario
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
 * @param {string} fechaInicio - La fecha de inicio en formato yyyy-MM-dd.
 * @param {string} horaInicio - La hora de inicio en formato HH:mm.
 * @param {string} fechaFin - La fecha de fin en formato yyyy-MM-dd.
 * @param {string} horaFin - La hora de fin en formato HH:mm.
 * @returns {string} - El número de horas y minutos laborales en formato h,m.
 */
function horasLaborales(fechaInicio, horaInicio, fechaFin, horaFin) {
  var mismoDia = (fechaInicio === fechaFin);
  var inicio = new Date(fechaInicio + "T" + horaInicio);
  var fin = new Date(fechaFin + "T" + horaFin);

  var totalMinutosLaborales;
  if (mismoDia) {
    totalMinutosLaborales = calcularMinutosLaboralesMismoDia(inicio, fin);
  } else {
    totalMinutosLaborales = calcularMinutosLaboralesVariosDias(inicio, fin);
    if (totalMinutosLaborales == "MesO+") {
      return "MesO+";
    }
  }

  var horasLaborales = totalMinutosLaborales / 60;
  var horas = Math.floor(horasLaborales);
  var minutos = Math.round((horasLaborales - horas) * 60);

  return horas + "," + minutos;
}

/**
 * Descuenta el tiempo de almuerzo de la duración total de las horas laborales.
 * @param {Date} horarioInicio - El horario de inicio del trabajo.
 * @param {Date} horarioFin - El horario de fin del trabajo.
 * @returns {number} - Los minutos descontados por el almuerzo.
 */
function descontarAlmuerzo(horarioInicio, horarioFin) {
  var inicioAlmuerzo = new Date(horarioInicio.getFullYear(), horarioInicio.getMonth(), horarioInicio.getDate(), 13, 0, 0);
  var finAlmuerzo = new Date(horarioInicio.getFullYear(), horarioInicio.getMonth(), horarioInicio.getDate(), 14, 0, 0);
  if (horarioFin &gt; inicioAlmuerzo &amp;&amp; horarioInicio &lt; finAlmuerzo) {
    var almuerzoInicio = new Date(Math.max(horarioInicio, inicioAlmuerzo));
    var almuerzoFin = new Date(Math.min(horarioFin, finAlmuerzo));
    return (almuerzoFin - almuerzoInicio) / (1000 * 60);
  }
  return 0;
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