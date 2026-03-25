/***********************
 * BLOQUE 13
 * PANTALLA VISUAL SESIONES
 ***********************/

function abrirPantallaSesiones() {
  const template = HtmlService.createTemplateFromFile('PantallaSesiones');
  template.pacientePreseleccionadoId = '';

  const html = template
    .evaluate()
    .setWidth(1180)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, 'Sesiones');
}

function obtenerDatosPantallaSesiones() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_SESIONES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return {
      sesiones: [],
      modalidades: obtenerValoresCatalogo_('MODALIDADES'),
      estadosSesion: obtenerValoresCatalogo_('ESTADOS_SESION'),
      estadosSync: []
    };
  }

  const idx = indexByHeader_(data[0]);
  const sesiones = [];
  const estadosSyncSet = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const syncStatus = row[idx.CalendarSyncStatus] || '';

    if (syncStatus) {
      estadosSyncSet[String(syncStatus)] = true;
    }

    sesiones.push({
      sesionId: row[idx.SesionID] || '',
      pacienteId: row[idx.PacienteID] || '',
      cicloId: row[idx.CicloID] || '',
      asignacionId: row[idx.AsignacionID] || '',
      modalidad: row[idx.Modalidad] || '',
      nombrePaciente: row[idx.NombrePaciente] || '',
      numeroSesion: Number(row[idx.NumeroSesion] || 0),
      fechaSesion: formatearFecha_(row[idx.FechaSesion]),
      estadoSesion: row[idx.EstadoSesion] || '',
      fechaOriginal: formatearFecha_(row[idx.FechaOriginal]),
      modificadaManual: row[idx.ModificadaManual] === true,
      notas: row[idx.Notas] || '',
      calendarEventId: row[idx.CalendarEventId] || '',
      calendarSyncStatus: syncStatus,
      calendarLastSync: formatearFecha_(row[idx.CalendarLastSync]),
      calendarEventTitle: row[idx.CalendarEventTitle] || ''
    });
  }

  return {
    sesiones: sesiones,
    modalidades: obtenerValoresCatalogo_('MODALIDADES'),
    estadosSesion: obtenerValoresCatalogo_('ESTADOS_SESION'),
    estadosSync: Object.keys(estadosSyncSet).sort((a, b) => a.localeCompare(b))
  };
}

function abrirPantallaSesionesDesdePaciente(pacienteId) {
  if (!pacienteId) {
    throw new Error('No se indicó paciente.');
  }

  const template = HtmlService.createTemplateFromFile('PantallaSesiones');
  template.pacientePreseleccionadoId = String(pacienteId);

  const html = template
    .evaluate()
    .setWidth(1180)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, 'Sesiones');
}