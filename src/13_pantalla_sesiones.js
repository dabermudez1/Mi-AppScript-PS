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
  const repo = new SessionRepository();
  const data = repo.findAll();
  const estadosSyncSet = {};

  const sesiones = data.map(s => {
    if (s.CalendarSyncStatus) estadosSyncSet[String(s.CalendarSyncStatus)] = true;
    return {
      sesionId: s.SesionID,
      pacienteId: s.PacienteID,
      cicloId: s.CicloID,
      asignacionId: s.AsignacionID,
      modalidad: s.Modalidad,
      nombrePaciente: s.NombrePaciente,
      numeroSesion: Number(s.NumeroSesion || 0),
      fechaSesion: formatearFecha_(s.FechaSesion),
      estadoSesion: s.EstadoSesion,
      fechaOriginal: formatearFecha_(s.FechaOriginal),
      modificadaManual: s.ModificadaManual === true,
      calendarSyncStatus: s.CalendarSyncStatus,
      calendarLastSync: formatearFecha_(s.CalendarLastSync),
      calendarEventId: s.CalendarEventId || '',
      calendarEventTitle: s.CalendarEventTitle || '',
      notas: s.Notas || ''
    };
  });

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