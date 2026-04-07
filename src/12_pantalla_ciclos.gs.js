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
  const repo = new CicloRepository();
  const data = repo.findAll();

  const ciclosMapeados = data.map(c => ({
    cicloId: c.CicloID,
    modalidad: c.Modalidad,
    numeroCiclo: c.NumeroCiclo,
    estadoCiclo: c.EstadoCiclo,
    fechaInicioCiclo: formatearFecha_(c.FechaInicioCiclo),
    fechaFinCiclo: formatearFecha_(c.FechaFinCiclo),
    diaSemana: c.DiaSemana,
    frecuenciaDias: c.FrecuenciaDias,
    capacidadMaxima: c.CapacidadMaxima,
    plazasOcupadas: c.PlazasOcupadas,
    plazasLibres: c.PlazasLibres,
    notas: c.Notas || ''
  }));

  return {
    ciclos: ciclosMapeados,
    modalidades: obtenerValoresCatalogo_('MODALIDADES'),
    estadosCiclo: obtenerValoresCatalogo_('ESTADOS_CICLO')
  };
}