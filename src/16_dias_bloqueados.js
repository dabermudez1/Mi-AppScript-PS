/***********************
 * BLOQUE 16
 * DÍAS BLOQUEADOS
 ***********************/

function gestionarDiasBloqueados() {
  const html = HtmlService
    .createHtmlOutputFromFile('DiasBloqueadosForm')
    .setWidth(760)
    .setHeight(620);

  SpreadsheetApp.getUi().showModalDialog(html, 'Días bloqueados');
}

function obtenerDiasBloqueadosFormulario() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DIAS_BLOQUEADOS');
  if (!sheet) {
    return [];
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return [];
  }

  // Normalizamos encabezados a minúsculas para evitar errores de teclado
  const headers = data[0].map(h => String(h).toLowerCase());
  const colFecha = headers.indexOf('fecha');
  const colBloqueado = headers.indexOf('bloqueado');
  const colMotivo = headers.indexOf('motivo');

  const out = [];

  for (let i = 1; i < data.length; i++) {
    const valFecha = data[i][colFecha];
    const fecha = valFecha instanceof Date ? valFecha : new Date(valFecha);
    
    if (!fecha || isNaN(fecha.getTime())) continue;

    out.push({
      fecha: formatearFecha_(fecha),
      bloqueado: esValorVerdadero_(data[i][colBloqueado]),
      motivo: data[i][colMotivo] || ''
    });
  }

  out.sort((a, b) => {
    const fa = parseFechaES_(a.fecha);
    const fb = parseFechaES_(b.fecha);
    return compararFechas_(fa, fb);
  });

  return out;
}

function guardarDiaBloqueadoFormulario(formData) {
  const fechaDesde = parseFechaES_(formData.fechaDesde || formData.fecha || '');
  const fechaHasta = parseFechaES_(formData.fechaHasta || formData.fechaDesde || formData.fecha || '');
  const motivo = String(formData.motivo || '').trim();

  if (!fechaDesde) {
    throw new Error('La fecha desde no es válida. Usa formato dd/mm/yyyy.');
  }

  if (!fechaHasta) {
    throw new Error('La fecha hasta no es válida. Usa formato dd/mm/yyyy.');
  }

  const desde = normalizarFecha_(fechaDesde);
  const hasta = normalizarFecha_(fechaHasta);

  if (hasta.getTime() < desde.getTime()) {
    throw new Error('La fecha hasta no puede ser anterior a la fecha desde.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('DIAS_BLOQUEADOS');

  if (!sheet) {
    sheet = ss.insertSheet('DIAS_BLOQUEADOS');
    sheet.getRange(1, 1, 1, 3).setValues([['Fecha', 'Bloqueado', 'Motivo']]);
  }

  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);

  const existentes = {};
  for (let i = 1; i < data.length; i++) {
    const fechaExistente = data[i][idx.Fecha];
    if (!fechaExistente) continue;

    const claveExistente = obtenerClaveFecha_(fechaExistente);
    existentes[claveExistente] = i + 1;
  }

  let actualizados = 0;
  let insertados = 0;
  let omitidosFinSemana = 0;

  let actual = new Date(desde);

  while (actual.getTime() <= hasta.getTime()) {
    const fechaActual = normalizarFecha_(actual);

    if (esFinDeSemana_(fechaActual)) {
      omitidosFinSemana++;
      actual = sumarDiasNaturales_(actual, 1);
      continue;
    }

    const clave = obtenerClaveFecha_(fechaActual);

    if (existentes[clave]) {
      const rowNum = existentes[clave];
      sheet.getRange(rowNum, idx.Bloqueado + 1).setValue(true);
      sheet.getRange(rowNum, idx.Motivo + 1).setValue(motivo);
      actualizados++;
    } else {
      sheet.appendRow([fechaActual, true, motivo]);
      insertados++;
    }

    actual = sumarDiasNaturales_(actual, 1);
  }

  let mensaje =
    'Bloqueo guardado correctamente.\n\n' +
    'Insertados: ' + insertados + '\n' +
    'Actualizados: ' + actualizados;

  if (omitidosFinSemana > 0) {
    mensaje += '\nFines de semana omitidos: ' + omitidosFinSemana;
  }

  // Sincronizar con Google Calendar después de modificar los días bloqueados
  sincronizarDiasBloqueadosAGoogleCalendar();

  return { mensaje: mensaje };
}

function eliminarLoteDiasBloqueadosFormulario(fechasTexto) {
  if (!Array.isArray(fechasTexto) || fechasTexto.length === 0) {
    throw new Error('No se han seleccionado fechas para eliminar.');
  }

  const clavesObjetivo = new Set(fechasTexto.map(f => {
    const fObj = parseFechaES_(f);
    if (!fObj) throw new Error('Fecha no válida: ' + f);
    return obtenerClaveFecha_(fObj);
  }));

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DIAS_BLOQUEADOS');
  if (!sheet) throw new Error('No existe la hoja DIAS_BLOQUEADOS.');

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { mensaje: 'No hay datos para eliminar.' };

  const idx = indexByHeader_(data[0]);
  let eliminados = 0;

  // Recorremos de abajo hacia arriba para mantener la validez de los índices al borrar filas
  for (let i = data.length - 1; i >= 1; i--) {
    const fechaFila = data[i][idx.Fecha];
    if (!fechaFila) continue;

    const claveFila = obtenerClaveFecha_(fechaFila);
    if (clavesObjetivo.has(claveFila)) {
      sheet.deleteRow(i + 1);
      eliminados++;
    }
  }

  // Sincronizar con Google Calendar después de modificar los días bloqueados
  sincronizarDiasBloqueadosAGoogleCalendar();

  return { mensaje: `Se han eliminado ${eliminados} días bloqueados correctamente.` };
}

function eliminarDiaBloqueadoFormulario(fechaTexto) {
  return eliminarLoteDiasBloqueadosFormulario([fechaTexto]);
}