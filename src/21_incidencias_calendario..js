/***********************
 * BLOQUE 21
 * INCIDENCIAS CALENDARIO
 ***********************/

function obtenerResumenIncidenciasCalendario(sesData, sesIdx) {
  const cantidad = obtenerCantidadSesionesAfectadasPorDiasBloqueados_(sesData, sesIdx);
  return { cantidad: cantidad };
}

function obtenerCantidadSesionesAfectadasPorDiasBloqueados_(sesData, sesIdx) {
  let data = sesData;
  let idx = sesIdx;

  if (!data) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
    if (!sheet) throw new Error('No existe la hoja ' + SHEET_SESIONES + '.');

    data = sheet.getDataRange().getValues();
  }

  if (!data || data.length < 2) return 0;

  if (!idx) idx = indexByHeader_(data[0]);

  const hoy = normalizarFecha_(new Date());
  const mapaDiasBloqueados = obtenerMapaDiasBloqueados_();

  let cantidad = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const fechaSesion = row[idx.FechaSesion];
    const estadoSesion = row[idx.EstadoSesion] || '';

    if (!fechaSesion || !(fechaSesion instanceof Date)) continue;

    const fechaNormalizada = normalizarFecha_(fechaSesion);

    if (fechaNormalizada.getTime() < hoy.getTime()) continue;

    if (
      estadoSesion !== ESTADOS_SESION.PENDIENTE &&
      estadoSesion !== ESTADOS_SESION.REPROGRAMADA
    ) {
      continue;
    }

    const detalleBloqueo = obtenerDetalleBloqueoFechaConMapa_(fechaNormalizada, mapaDiasBloqueados);
    if (!detalleBloqueo.bloqueada) continue;

    cantidad++;
  }

  return cantidad;
}

function verIncidenciasCalendario() {
  const html = HtmlService
    .createHtmlOutputFromFile('IncidenciasCalendarioForm')
    .setWidth(1100)
    .setHeight(720);

  SpreadsheetApp.getUi().showModalDialog(html, 'Incidencias de calendario');
}

function obtenerIncidenciasCalendarioFormulario() {
  return obtenerSesionesAfectadasPorDiasBloqueados_();
}

function obtenerSesionesAfectadasPorDiasBloqueados_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_SESIONES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const idx = indexByHeader_(data[0]);
  const out = [];
  const hoy = normalizarFecha_(new Date());
  const mapaDiasBloqueados = obtenerMapaDiasBloqueados_();

  for (let i = 1; i < data.length; i++) {
    const sesionId = data[i][idx.SesionID] || '';
    const pacienteId = data[i][idx.PacienteID] || '';
    const cicloId = data[i][idx.CicloID] || '';
    const modalidad = data[i][idx.Modalidad] || '';
    const nombrePaciente = data[i][idx.NombrePaciente] || '';
    const numeroSesion = Number(data[i][idx.NumeroSesion] || 0);
    const fechaSesion = data[i][idx.FechaSesion];
    const estadoSesion = data[i][idx.EstadoSesion] || '';

    if (!fechaSesion || !(fechaSesion instanceof Date)) continue;

    const fechaNormalizada = normalizarFecha_(fechaSesion);

    if (fechaNormalizada.getTime() < hoy.getTime()) continue;

    if (
      estadoSesion !== ESTADOS_SESION.PENDIENTE &&
      estadoSesion !== ESTADOS_SESION.REPROGRAMADA
    ) {
      continue;
    }

    const detalleBloqueo = obtenerDetalleBloqueoFechaConMapa_(fechaNormalizada, mapaDiasBloqueados);
    if (!detalleBloqueo.bloqueada) continue;

    out.push({
      sesionId: sesionId,
      pacienteId: pacienteId,
      cicloId: cicloId,
      modalidad: modalidad,
      nombrePaciente: nombrePaciente,
      numeroSesion: numeroSesion,
      fechaSesion: formatearFecha_(fechaNormalizada),
      estadoSesion: estadoSesion,
      tipoBloqueo: detalleBloqueo.tipo || '',
      motivoBloqueo: detalleBloqueo.motivo || ''
    });
  }

  out.sort(function(a, b) {
    const fa = parseFechaES_(a.fechaSesion);
    const fb = parseFechaES_(b.fechaSesion);

    if (fa && fb && fa.getTime() !== fb.getTime()) {
      return fa.getTime() - fb.getTime();
    }

    if ((a.modalidad || '') !== (b.modalidad || '')) {
      return String(a.modalidad || '').localeCompare(String(b.modalidad || ''));
    }

    return String(a.nombrePaciente || '').localeCompare(String(b.nombrePaciente || ''));
  });

  return out;
}