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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CICLOS);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_CICLOS + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return {
      ciclos: [],
      modalidades: obtenerValoresCatalogo_('MODALIDADES'),
      estadosCiclo: obtenerValoresCatalogo_('ESTADOS_CICLO')
    };
  }

  const idx = indexByHeader_(data[0]);
  const ciclos = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    ciclos.push({
      cicloId: row[idx.CicloID] || '',
      modalidad: row[idx.Modalidad] || '',
      numeroCiclo: Number(row[idx.NumeroCiclo] || 0),
      estadoCiclo: row[idx.EstadoCiclo] || '',
      fechaInicioCiclo: formatearFecha_(row[idx.FechaInicioCiclo]),
      fechaFinCiclo: formatearFecha_(row[idx.FechaFinCiclo]),
      fechaBaseUsada: formatearFecha_(row[idx.FechaBaseUsada]),
      diaSemana: row[idx.DiaSemana] || '',
      frecuenciaDias: Number(row[idx.FrecuenciaDias] || 0),
      sesionesPorCiclo: Number(row[idx.SesionesPorCiclo] || 0),
      capacidadMaxima: Number(row[idx.CapacidadMaxima] || 0),
      plazasOcupadas: Number(row[idx.PlazasOcupadas] || 0),
      plazasLibres: Number(row[idx.PlazasLibres] || 0),
      generadoManual: row[idx.GeneradoManual] === true,
      notas: row[idx.Notas] || ''
    });
  }

  return {
    ciclos: ciclos,
    modalidades: obtenerValoresCatalogo_('MODALIDADES'),
    estadosCiclo: obtenerValoresCatalogo_('ESTADOS_CICLO')
  };
}