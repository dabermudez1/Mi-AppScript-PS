/***********************
 * BLOQUE 12
 * PANTALLA VISUAL CICLOS
 ***********************/

function abrirPantallaCiclos() {
  const html = HtmlService
    .createHtmlOutputFromFile('PantallaCiclos')
    .setWidth(1120)
    .setHeight(720);

  SpreadsheetApp.getUi().showModalDialog(html, 'Ciclos');
}

function obtenerDatosPantallaCiclos() {
  const ciclos = new CicloRepository().findAll().map(c => ({
    ...c,
    fechaInicioCiclo: formatearFecha_(c.FechaInicioCiclo),
    fechaFinCiclo: formatearFecha_(c.FechaFinCiclo),
    fechaBaseUsada: formatearFecha_(c.FechaBaseUsada)
  }));

  return {
    ciclos: ciclos,
    modalidades: obtenerValoresCatalogo_('MODALIDADES'),
    estadosCiclo: obtenerValoresCatalogo_('ESTADOS_CICLO')
  };
}