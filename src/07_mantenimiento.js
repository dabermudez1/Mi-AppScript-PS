/***********************
 * BLOQUE 7
 * MANTENIMIENTO
 *
 * Este archivo ahora contiene solo los entry points del menú.
 ***********************/

function recalcularOcupacionCiclos() {
  const ui = SpreadsheetApp.getUi();

  try {
    const actualizados = new MaintenanceService().recalculateCycleOccupancy();

    ui.alert(
      'Ocupación recalculada correctamente.\n\n' +
      'Ciclos actualizados: ' + actualizados
    );
  } catch (error) {
    ui.alert('Error al recalcular ocupación: ' + error.message);
    throw error;
  }
}