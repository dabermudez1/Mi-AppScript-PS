/***********************
 * BLOQUE 5
 * MOTOR DE ESTADOS
 ***********************/

function actualizarEstadosAutomaticos() {
  const ui = SpreadsheetApp.getUi();

  try {
    recalcularOcupacionCiclosInterno_();
    const res = StatusService.runDailyUpdate();

    try {
      refrescarDashboard();
    } catch (e) {
      // No rompemos el proceso por un fallo visual del dashboard
    }

    ui.alert(
      'Actualización completada.\n\n' +
      'Ciclos recalculados (antes): ' + resumen.ocupacionAntes + '\n' +
      'Ciclos → EN_CURSO: ' + resumen.ciclosEnCurso + '\n' +
      'Ciclos → CERRADO: ' + resumen.ciclosCerrados + '\n' +
      'Asignaciones → ACTIVO: ' + resumen.asignacionesActivadas + '\n' +
      'Asignaciones → FINALIZADO: ' + resumen.asignacionesFinalizadas + '\n' +
      'Pacientes → ACTIVO: ' + resumen.pacientesActivados + '\n' +
      'Sesiones auto-completadas: ' + resumen.sesionesAutoCompletadas + '\n' +
      'Pacientes → ALTA: ' + resumen.pacientesAlta + '\n' +
      'Ciclos recalculados (después): ' + resumen.ocupacionDespues
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
  recalcularOcupacionCiclosInterno_();
  actualizarCiclosAEnCurso_();
  actualizarCiclosACerrado_();
  actualizarAsignacionesAActivas_();
  actualizarAsignacionesAFinalizadas_();
  actualizarPacientesAPorInicioDeCiclo_();
  actualizarSesionesVencidas_();
  actualizarPacientesAAlta_();
  recalcularOcupacionCiclosInterno_();

  try {
    refrescarDashboard();
  } catch (e) {
    // Evitamos romper el trigger por el dashboard
  }
}

/***************
 * CICLOS
 ***************/
function actualizarCiclosAEnCurso_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CICLOS);
  if (!sheet) throw new Error('No existe la hoja ' + SHEET_CICLOS + '.');

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;

  const headers = data[0];
  const idx = indexByHeader_(headers);
  const hoy = normalizarFecha_(new Date());

  let cambios = 0;

  for (let i = 1; i < data.length; i++) {
    const estado = data[i][idx.EstadoCiclo];
    const fechaInicio = data[i][idx.FechaInicioCiclo];

    if (
      estado === ESTADOS_CICLO.PLANIFICADO &&
      fechaInicio instanceof Date &&
      normalizarFecha_(fechaInicio).getTime() <= hoy.getTime()
    ) {
      sheet.getRange(i + 1, idx.EstadoCiclo + 1).setValue(ESTADOS_CICLO.EN_CURSO);
      cambios++;
    }
  }

  return cambios;
}

function actualizarCiclosACerrado_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CICLOS);
  if (!sheet) throw new Error('No existe la hoja ' + SHEET_CICLOS + '.');

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;

  const headers = data[0];
  const idx = indexByHeader_(headers);
  const hoy = normalizarFecha_(new Date());

  let cambios = 0;

  for (let i = 1; i < data.length; i++) {
    const estado = data[i][idx.EstadoCiclo];
    const fechaFin = data[i][idx.FechaFinCiclo];

    if (
      (estado === ESTADOS_CICLO.PLANIFICADO || estado === ESTADOS_CICLO.EN_CURSO) &&
      fechaFin instanceof Date &&
      normalizarFecha_(fechaFin).getTime() < hoy.getTime()
    ) {
      sheet.getRange(i + 1, idx.EstadoCiclo + 1).setValue(ESTADOS_CICLO.CERRADO);
      cambios++;
    }
  }

  return cambios;
}

/***************
 * ASIGNACIONES
 ***************/
function actualizarAsignacionesAActivas_() {
  const sheetAsign = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ASIGNACIONES_CICLO);
  const sheetCiclos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CICLOS);

  if (!sheetAsign || !sheetCiclos) {
    throw new Error('Faltan hojas de asignaciones o ciclos.');
  }

  const asignData = sheetAsign.getDataRange().getValues();
  const ciclosData = sheetCiclos.getDataRange().getValues();

  if (asignData.length < 2 || ciclosData.length < 2) return 0;

  const aIdx = indexByHeader_(asignData[0]);
  const cIdx = indexByHeader_(ciclosData[0]);

  const estadoCicloPorId = {};
  for (let i = 1; i < ciclosData.length; i++) {
    estadoCicloPorId[String(ciclosData[i][cIdx.CicloID])] = ciclosData[i][cIdx.EstadoCiclo];
  }

  let cambios = 0;

  for (let i = 1; i < asignData.length; i++) {
    const estadoAsign = asignData[i][aIdx.EstadoAsignacion];
    const cicloId = String(asignData[i][aIdx.CicloID] || '');
    const estadoCiclo = estadoCicloPorId[cicloId] || '';

    if (
      estadoAsign === ESTADOS_ASIGNACION.RESERVADO &&
      estadoCiclo === ESTADOS_CICLO.EN_CURSO
    ) {
      sheetAsign.getRange(i + 1, aIdx.EstadoAsignacion + 1).setValue(ESTADOS_ASIGNACION.ACTIVO);
      cambios++;
    }
  }

  return cambios;
}

function actualizarAsignacionesAFinalizadas_() {
  const sheetAsign = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ASIGNACIONES_CICLO);
  const sheetCiclos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CICLOS);

  if (!sheetAsign || !sheetCiclos) {
    throw new Error('Faltan hojas de asignaciones o ciclos.');
  }

  const asignData = sheetAsign.getDataRange().getValues();
  const ciclosData = sheetCiclos.getDataRange().getValues();

  if (asignData.length < 2 || ciclosData.length < 2) return 0;

  const aIdx = indexByHeader_(asignData[0]);
  const cIdx = indexByHeader_(ciclosData[0]);

  const estadoCicloPorId = {};
  for (let i = 1; i < ciclosData.length; i++) {
    estadoCicloPorId[String(ciclosData[i][cIdx.CicloID])] = ciclosData[i][cIdx.EstadoCiclo];
  }

  let cambios = 0;

  for (let i = 1; i < asignData.length; i++) {
    const estadoAsign = asignData[i][aIdx.EstadoAsignacion];
    const cicloId = String(asignData[i][aIdx.CicloID] || '');
    const estadoCiclo = estadoCicloPorId[cicloId] || '';

    if (
      (estadoAsign === ESTADOS_ASIGNACION.RESERVADO || estadoAsign === ESTADOS_ASIGNACION.ACTIVO) &&
      estadoCiclo === ESTADOS_CICLO.CERRADO
    ) {
      sheetAsign.getRange(i + 1, aIdx.EstadoAsignacion + 1).setValue(ESTADOS_ASIGNACION.FINALIZADO);
      cambios++;
    }
  }

  return cambios;
}

/***************
 * PACIENTES
 ***************/
function actualizarPacientesAPorInicioDeCiclo_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPac = ss.getSheetByName(SHEET_PACIENTES);
  const sheetCiclos = ss.getSheetByName(SHEET_CICLOS);

  if (!sheetPac || !sheetCiclos) {
    throw new Error('Faltan hojas PACIENTES o CICLOS.');
  }

  const pacData = sheetPac.getDataRange().getValues();
  const ciclosData = sheetCiclos.getDataRange().getValues();

  if (pacData.length < 2 || ciclosData.length < 2) return 0;

  const pIdx = indexByHeader_(pacData[0]);
  const cIdx = indexByHeader_(ciclosData[0]);

  const ciclosPorId = {};
  for (let i = 1; i < ciclosData.length; i++) {
    ciclosPorId[String(ciclosData[i][cIdx.CicloID])] = {
      EstadoCiclo: ciclosData[i][cIdx.EstadoCiclo],
      FechaInicioCiclo: ciclosData[i][cIdx.FechaInicioCiclo]
    };
  }

  let cambios = 0;

  for (let i = 1; i < pacData.length; i++) {
    const estadoPaciente = pacData[i][pIdx.EstadoPaciente];
    const cicloObjetivoId = String(pacData[i][pIdx.CicloObjetivoID] || '');

    if (estadoPaciente !== ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO) continue;
    if (!cicloObjetivoId) continue;

    const ciclo = ciclosPorId[cicloObjetivoId];
    if (!ciclo) continue;

    if (ciclo.EstadoCiclo === ESTADOS_CICLO.EN_CURSO) {
      sheetPac.getRange(i + 1, pIdx.EstadoPaciente + 1).setValue(ESTADOS_PACIENTE.ACTIVO);
      sheetPac.getRange(i + 1, pIdx.CicloActivoID + 1).setValue(cicloObjetivoId);
      cambios++;
    }
  }

  return cambios;
}

function actualizarPacientesAAlta_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;

  const idx = indexByHeader_(data[0]);
  const hoy = normalizarFecha_(new Date());

  let cambios = 0;

  for (let i = 1; i < data.length; i++) {
    const estado = data[i][idx.EstadoPaciente];
    const planificadas = Number(data[i][idx.SesionesPlanificadas] || 0);
    const completadas = Number(data[i][idx.SesionesCompletadas] || 0);

    if (
      (estado === ESTADOS_PACIENTE.ACTIVO || estado === ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO) &&
      planificadas > 0 &&
      completadas >= planificadas
    ) {
      sheet.getRange(i + 1, idx.EstadoPaciente + 1).setValue(ESTADOS_PACIENTE.ALTA);
      sheet.getRange(i + 1, idx.FechaCierre + 1).setValue(hoy);
      sheet.getRange(i + 1, idx.ProximaSesion + 1).setValue('');
      cambios++;
    }
  }

  return cambios;
}

/***************
 * SESIONES
 ***************/
function actualizarSesionesVencidas_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  if (!sheet) throw new Error('No existe la hoja ' + SHEET_SESIONES + '.');

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;

  const idx = indexByHeader_(data[0]);
  const hoy = normalizarFecha_(new Date());

  let cambios = 0;

  for (let i = 1; i < data.length; i++) {
    const estado = data[i][idx.EstadoSesion];
    const fechaSesion = data[i][idx.FechaSesion];

    if (
      estado === ESTADOS_SESION.PENDIENTE &&
      fechaSesion instanceof Date &&
      normalizarFecha_(fechaSesion).getTime() < hoy.getTime()
    ) {
      sheet.getRange(i + 1, idx.EstadoSesion + 1).setValue(ESTADOS_SESION.COMPLETADA_AUTO);
      cambios++;
    }
  }

  if (cambios > 0) {
    recalcularMetricasBasicas_();
  }

  return cambios;
}

/***************
 * MÉTRICAS BÁSICAS
 ***************/
function recalcularMetricasBasicas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPac = ss.getSheetByName(SHEET_PACIENTES);
  const sheetSes = ss.getSheetByName(SHEET_SESIONES);

  if (!sheetPac || !sheetSes) {
    throw new Error('Faltan hojas PACIENTES o SESIONES.');
  }

  const pacData = sheetPac.getDataRange().getValues();
  const sesData = sheetSes.getDataRange().getValues();

  if (pacData.length < 2) return;

  const pIdx = indexByHeader_(pacData[0]);
  const sIdx = indexByHeader_(sesData[0]);

  const statsPorPaciente = {};

  for (let i = 1; i < sesData.length; i++) {
    const row = sesData[i];
    const pacienteId = String(row[sIdx.PacienteID] || '');
    const estadoSesion = row[sIdx.EstadoSesion];
    const fechaSesion = row[sIdx.FechaSesion];

    if (!pacienteId) continue;

    if (!statsPorPaciente[pacienteId]) {
      statsPorPaciente[pacienteId] = {
        completadas: 0,
        pendientes: 0,
        proximaSesion: ''
      };
    }

    if (
      estadoSesion === ESTADOS_SESION.COMPLETADA_AUTO ||
      estadoSesion === ESTADOS_SESION.COMPLETADA_MANUAL
    ) {
      statsPorPaciente[pacienteId].completadas++;
    }

    if (
      estadoSesion === ESTADOS_SESION.PENDIENTE ||
      estadoSesion === ESTADOS_SESION.REPROGRAMADA
    ) {
      statsPorPaciente[pacienteId].pendientes++;

      if (fechaSesion instanceof Date) {
        const fechaNorm = normalizarFecha_(fechaSesion);
        const actual = statsPorPaciente[pacienteId].proximaSesion;

        if (!actual || fechaNorm.getTime() < actual.getTime()) {
          statsPorPaciente[pacienteId].proximaSesion = fechaNorm;
        }
      }
    }
  }

  const numPacFilas = Math.max(pacData.length - 1, 0);
  const completadasOut = new Array(numPacFilas);
  const pendientesOut = new Array(numPacFilas);
  const proximaOut = new Array(numPacFilas);

  for (let i = 1; i < pacData.length; i++) {
    const outIdx = i - 1;
    const pacienteId = String(pacData[i][pIdx.PacienteID] || '');
    const stats = statsPorPaciente[pacienteId] || {
      completadas: 0,
      pendientes: 0,
      proximaSesion: ''
    };

    completadasOut[outIdx] = [stats.completadas];
    pendientesOut[outIdx] = [stats.pendientes];
    proximaOut[outIdx] = [stats.proximaSesion || ''];
  }

  if (numPacFilas > 0) {
    sheetPac.getRange(2, pIdx.SesionesCompletadas + 1, numPacFilas, 1).setValues(completadasOut);
    sheetPac.getRange(2, pIdx.SesionesPendientes + 1, numPacFilas, 1).setValues(pendientesOut);
    sheetPac.getRange(2, pIdx.ProximaSesion + 1, numPacFilas, 1).setValues(proximaOut);
  }
}