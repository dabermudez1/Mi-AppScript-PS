/***********************
 * BLOQUE 22
 * ESTADÍSTICAS FICHAS PACIENTES
 ***********************/

function estadisticasFichasPacientes() {
  const html = HtmlService
    .createHtmlOutputFromFile('EstadisticasFichasForm')
    .setWidth(1650)
    .setHeight(820);

  SpreadsheetApp.getUi().showModalDialog(html, 'Estadísticas fichas pacientes');
}

/**
 * DEPRECATED: La lógica se ha movido a 19_datos_clinicos_pacientes.js para centralizarla.
 * Esta función ahora es un simple wrapper para mantener la compatibilidad con la UI.
 */
function obtenerDatosEstadisticasFichasFormulario() {
  // Redirigimos la llamada a la función correcta y centralizada.
  return _obtenerDatosEstadisticasFichas();
}