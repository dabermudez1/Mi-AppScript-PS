/***********************
 * BLOQUE 20
 * UTILIDADES DEPURACIÓN
 ***********************/

function limpiarDatosOperativosDepuracion_(confirmacion) {
  if (confirmacion !== 'CONFIRMAR_RESET_OPERATIVO') {
    throw new Error(
      'Confirmación inválida para limpieza operativa.\n\n' +
      'Usa exactamente: CONFIRMAR_RESET_OPERATIVO'
    );
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojasObjetivo = [
    SHEET_CICLOS,
    SHEET_ASIGNACIONES_CICLO,
    SHEET_SESIONES,
    SHEET_DATOS_CLINICOS_PACIENTES || 'DATOS_CLINICOS_PACIENTES'
  ];

  hojasObjetivo.forEach(function(nombreHoja) {
    vaciarHojaManteniendoCabecera_(ss, nombreHoja);
  });

  return {
    mensaje:
      'Limpieza operativa completada.\n\n' +
      'Se han vaciado:\n' +
      '- ' + SHEET_CICLOS + '\n' +
      '- ' + SHEET_ASIGNACIONES_CICLO + '\n' +
      '- ' + SHEET_SESIONES + '\n' +
      '- DATOS_CLINICOS_PACIENTES'
  };
}

function resetProyectoDepuracion_(confirmacion) {
  if (confirmacion !== 'CONFIRMAR_RESET_TOTAL') {
    throw new Error(
      'Confirmación inválida para reset total.\n\n' +
      'Usa exactamente: CONFIRMAR_RESET_TOTAL'
    );
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojasObjetivo = [
    SHEET_PACIENTES,
    SHEET_CICLOS,
    SHEET_ASIGNACIONES_CICLO,
    SHEET_SESIONES,
    SHEET_DATOS_CLINICOS_PACIENTES || 'DATOS_CLINICOS_PACIENTES'
  ];

  hojasObjetivo.forEach(function(nombreHoja) {
    vaciarHojaManteniendoCabecera_(ss, nombreHoja);
  });

  return {
    mensaje:
      'Reset total de depuración completado.\n\n' +
      'Se han vaciado:\n' +
      '- ' + SHEET_PACIENTES + '\n' +
      '- ' + SHEET_CICLOS + '\n' +
      '- ' + SHEET_ASIGNACIONES_CICLO + '\n' +
      '- ' + SHEET_SESIONES + '\n' +
      '- DATOS_CLINICOS_PACIENTES\n\n' +
      'No se han tocado:\n' +
      '- DIAS_BLOQUEADOS\n' +
      '- ' + SHEET_CATALOGOS + '\n' +
      '- ' + SHEET_CONFIG_MODALIDADES
  };
}

function vaciarHojaManteniendoCabecera_(ss, nombreHoja) {
  const sheet = ss.getSheetByName(nombreHoja);
  if (!sheet) {
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow <= 1 || lastColumn <= 0) {
    return;
  }

  sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
}

function limpiarIntegracionCalendarDepuracion_(confirmacion) {
  if (confirmacion !== 'CONFIRMAR_RESET_CALENDAR') {
    throw new Error(
      'Confirmación inválida para reset de calendar.\n\n' +
      'Usa exactamente: CONFIRMAR_RESET_CALENDAR'
    );
  }

  try {
    limpiarCamposSyncCalendarSesiones();
  } catch (e) {
    Logger.log('Error limpiando sync: ' + e.message);
  }

  try {
    resetCalendarioConsultaVinculado();
  } catch (e) {
    Logger.log('Error reseteando vínculo calendario: ' + e.message);
  }

  return {
    mensaje:
      'Integración de Google Calendar reseteada.\n\n' +
      '- Sync de sesiones limpiado\n' +
      '- Calendario desvinculado\n\n' +
      'NO se han borrado eventos del calendario.'
  };
}

function resetEntornoCompletoDepuracion_(confirmacion) {
  if (confirmacion !== 'CONFIRMAR_RESET_TOTAL_ABSOLUTO') {
    throw new Error(
      'Confirmación inválida.\n\n' +
      'Usa: CONFIRMAR_RESET_TOTAL_ABSOLUTO'
    );
  }

  // 1. Datos
  resetProyectoDepuracion_('CONFIRMAR_RESET_TOTAL');

  // 2. Calendar
  limpiarIntegracionCalendarDepuracion_('CONFIRMAR_RESET_CALENDAR');

  return {
    mensaje:
      'RESET TOTAL COMPLETO realizado.\n\n' +
      '- Datos borrados\n' +
      '- Calendar desvinculado\n' +
      '- Sync limpio\n\n' +
      'Sistema listo para pruebas desde cero.'
  };
}

function limpiarIntegracionCalendarDepuracion_(confirmacion) {
  if (confirmacion !== 'CONFIRMAR_RESET_CALENDAR') {
    throw new Error(
      'Confirmación inválida para reset de calendar.\n\n' +
      'Usa exactamente: CONFIRMAR_RESET_CALENDAR'
    );
  }

  try {
    limpiarCamposSyncCalendarSesiones();
  } catch (e) {
    Logger.log('Error limpiando campos sync calendar: ' + e.message);
  }

  try {
    resetCalendarioConsultaVinculado();
  } catch (e) {
    Logger.log('Error reseteando calendario vinculado: ' + e.message);
  }

  return {
    mensaje:
      'Integración de Google Calendar reseteada.\n\n' +
      '- Sync de sesiones limpiado\n' +
      '- Calendario desvinculado\n\n' +
      'No se han borrado eventos del calendario.'
  };
}

function resetEntornoCompletoDepuracion_(confirmacion) {
  if (confirmacion !== 'CONFIRMAR_RESET_TOTAL_ABSOLUTO') {
    throw new Error(
      'Confirmación inválida.\n\n' +
      'Usa exactamente: CONFIRMAR_RESET_TOTAL_ABSOLUTO'
    );
  }

  resetProyectoDepuracion_('CONFIRMAR_RESET_TOTAL');
  limpiarIntegracionCalendarDepuracion_('CONFIRMAR_RESET_CALENDAR');

  return {
    mensaje:
      'RESET TOTAL COMPLETO realizado.\n\n' +
      '- Datos vaciados\n' +
      '- Integración Google Calendar reseteada\n\n' +
      'Sistema listo para pruebas desde cero.'
  };
}

function ejecutarLimpiarDatosOperativosDepuracion_() {
  const res = limpiarDatosOperativosDepuracion_('CONFIRMAR_RESET_OPERATIVO');
  SpreadsheetApp.getUi().alert(res.mensaje);
}

function ejecutarResetProyectoDepuracion_() {
  const res = resetProyectoDepuracion_('CONFIRMAR_RESET_TOTAL');
  SpreadsheetApp.getUi().alert(res.mensaje);
}

function ejecutarLimpiarIntegracionCalendarDepuracion_() {
  const res = limpiarIntegracionCalendarDepuracion_('CONFIRMAR_RESET_CALENDAR');
  SpreadsheetApp.getUi().alert(res.mensaje);
}

function ejecutarResetEntornoCompletoDepuracion_() {
  const res = resetEntornoCompletoDepuracion_('CONFIRMAR_RESET_TOTAL_ABSOLUTO');
  SpreadsheetApp.getUi().alert(res.mensaje);
}