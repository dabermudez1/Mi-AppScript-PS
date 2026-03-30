/***********************
 * BLOQUE 5
 * MOTOR DE ESTADOS
 ***********************/

function actualizarEstadosAutomaticos() {
  const ui = SpreadsheetApp.getUi();

  try {
    const stateService = new StateService();
    const stats = stateService.runAutomaticTransitions();

    try {
      refrescarDashboard();
    } catch (e) {
      // No rompemos el proceso por un fallo visual del dashboard
    }

    ui.alert(
      `Actualización completada.\n\nCiclos procesados: ${stats.ciclos}\nPacientes actualizados: ${stats.pacientes}\nSesiones auto-completadas: ${stats.sesiones}`
    );

  } catch (error) {
    ui.alert('Error al actualizar estados: ' + error.message);
    throw error;
  }
}

function crearTriggerEstadosAutomaticos() {
  const ui = SpreadsheetApp.getUi();

  try {
    const triggers = ScriptApp.getProjectTriggers();
    const yaExiste = triggers.some(t => t.getHandlerFunction() === 'actualizarEstadosAutomaticosTrigger');

    if (yaExiste) {
      ui.alert('El trigger diario de estados automáticos ya existe.');
      return;
    }

    ScriptApp.newTrigger('actualizarEstadosAutomaticosTrigger')
      .timeBased()
      .everyDays(1)
      .atHour(2)
      .create();

    ui.alert('Trigger diario de estados automáticos creado correctamente.');
  } catch (error) {
    ui.alert('Error al crear trigger: ' + error.message);
    throw error;
  }
}

function eliminarTriggerEstadosAutomaticos() {
  const ui = SpreadsheetApp.getUi();

  try {
    const triggers = ScriptApp.getProjectTriggers();
    let eliminados = 0;

    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'actualizarEstadosAutomaticosTrigger') {
        ScriptApp.deleteTrigger(trigger);
        eliminados++;
      }
    });

    ui.alert('Triggers eliminados: ' + eliminados);
  } catch (error) {
    ui.alert('Error al eliminar trigger: ' + error.message);
    throw error;
  }
}

function actualizarEstadosAutomaticosTrigger() {
  const stateService = new StateService();
  stateService.runAutomaticTransitions();

  try {
    refrescarDashboard();
  } catch (e) {
    // Evitamos romper el trigger por el dashboard
  }
}
// Eliminada lógica duplicada: StateService ahora gestiona estas transiciones.