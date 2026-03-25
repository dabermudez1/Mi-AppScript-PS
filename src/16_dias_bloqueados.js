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

  const idx = indexByHeader_(data[0]);
  const out = [];

  // Obtener el primer y último día del próximo mes
  const fechaHoy = new Date();
  const primerDiaProximoMes = new Date(fechaHoy.getFullYear(), fechaHoy.getMonth() + 1, 1);
  const ultimoDiaProximoMes = new Date(fechaHoy.getFullYear(), fechaHoy.getMonth() + 2, 0); // El 0 nos da el último día del mes anterior al mes siguiente

  for (let i = 1; i < data.length; i++) {
    const fecha = new Date(data[i][idx.Fecha]); // Asegúrate de que la fecha esté en el formato adecuado
    const bloqueado = data[i][idx.Bloqueado];
    const motivo = data[i][idx.Motivo];

    if (!fecha || fecha < primerDiaProximoMes || fecha > ultimoDiaProximoMes) continue; // Filtra los días fuera del rango del próximo mes

    out.push({
      fecha: formatearFecha_(fecha),  // Asegúrate de que esta función esté bien implementada
      bloqueado: esValorVerdadero_(bloqueado),
      motivo: motivo || ''
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

  return { mensaje: mensaje };
}

function eliminarDiaBloqueadoFormulario(fechaTexto) {
  const fecha = parseFechaES_(fechaTexto);
  if (!fecha) {
    throw new Error('La fecha indicada no es válida.');
  }

  const claveObjetivo = obtenerClaveFecha_(fecha);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DIAS_BLOQUEADOS');
  if (!sheet) {
    throw new Error('No existe la hoja DIAS_BLOQUEADOS.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('No hay días bloqueados para eliminar.');
  }

  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    const fechaFila = data[i][idx.Fecha];
    if (!fechaFila) continue;

    const claveFila = obtenerClaveFecha_(fechaFila);

    if (claveFila === claveObjetivo) {
      sheet.deleteRow(i + 1);
      return { mensaje: 'Día bloqueado eliminado correctamente.' };
    }
  }

  throw new Error('No se encontró el día bloqueado indicado.');
}