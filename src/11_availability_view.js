/***********************
 * BLOQUE 11
 * PANTALLA DE DISPONIBILIDAD
 ***********************/

/**
 * Punto de entrada para abrir la nueva pantalla de disponibilidad desde el menú.
 */
function abrirPantallaDisponibilidad() {
  const html = HtmlService.createHtmlOutputFromFile('AvailabilityView')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Vista de Disponibilidad Semanal');
}

/**
 * Obtiene los datos procesados para la vista de disponibilidad semanal.
 * Esta es la función que será llamada por el HTML.
 * @param {string} startDateISO - Fecha de inicio de la semana en formato YYYY-MM-DD.
 * @returns {Object} Un objeto que contiene los datos de la semana y la fecha actual.
 */
function getAvailabilityViewData(startDateISO) {
  try {
    const availabilityService = new AvailabilityService();
    const startDate = startDateISO ? parseFechaISO_(startDateISO) : new Date();
    
    const data = availabilityService.getWeeklyState(startDate);

    return {
      weekData: data
    };
  } catch (e) {
    Logger.log(`Error en getAvailabilityViewData: ${e.message} ${e.stack}`);
    throw new Error(`No se pudieron cargar los datos de disponibilidad: ${e.message}`);
  }
}

/**
 * Parsea una fecha en formato ISO (YYYY-MM-DD) a un objeto Date.
 * @param {string} texto - La fecha en formato ISO.
 * @returns {Date|null}
 */
function parseFechaISO_(texto) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec((texto || '').trim());
  if (!m) return null;

  const year = Number(m[1]);
  const month = Number(m[2]) - 1;
  const day = Number(m[3]);

  const fecha = new Date(year, month, day);

  if (fecha.getFullYear() !== year || fecha.getMonth() !== month || fecha.getDate() !== day) return null;

  return normalizarFecha_(fecha); // normalizarFecha_ está en 01_base.js y es global
}