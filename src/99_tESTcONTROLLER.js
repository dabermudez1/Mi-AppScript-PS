/**
 * Funciones de entrada para el front-end.
 */
function getProposedSlots(config) {
  try {
    const service = new PlanningService();
    const ciclo = { fechaInicio: new Date(config.fechaInicio) };
    const sesiones = service.planificarCicloSeguimiento(ciclo, config.numSesiones, config.frecuencia);

    return sesiones.map(s => ({
      numero: s.NumeroSesion,
      fecha: Utilities.formatDate(s.FechaSesion, Session.getScriptTimeZone(), "dd/MM/yyyy"),
      hora: s.HoraInicio
    }));
  } catch (e) {
    throw new Error(e.message);
  }
}

function openTestPlanningDialog() {
  const html = HtmlService.createHtmlOutputFromFile('TestPlanningUI')
      .setWidth(450)
      .setHeight(550)
      .setTitle('Planificador por Slots');
  SpreadsheetApp.getUi().showModalDialog(html, 'Simulación de Agenda');
}
