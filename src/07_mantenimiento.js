/***********************
 * BLOQUE 7
 * MANTENIMIENTO
 ***********************/

function recalcularOcupacionCiclos() {
  const ui = SpreadsheetApp.getUi();

  try {
    const actualizados = recalcularOcupacionCiclosInterno_();

    ui.alert(
      'Ocupación recalculada correctamente.\n\n' +
      'Ciclos actualizados: ' + actualizados
    );
  } catch (error) {
    ui.alert('Error al recalcular ocupación: ' + error.message);
    throw error;
  }
}

function recalcularOcupacionCiclosInterno_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCiclos = ss.getSheetByName(SHEET_CICLOS);
  const sheetAsign = ss.getSheetByName(SHEET_ASIGNACIONES_CICLO);

  if (!sheetCiclos || !sheetAsign) {
    throw new Error('Faltan hojas CICLOS o ASIGNACIONES_CICLO.');
  }

  const ciclosData = sheetCiclos.getDataRange().getValues();
  const asignData = sheetAsign.getDataRange().getValues();

  if (ciclosData.length < 2) {
    return 0;
  }

  const cIdx = indexByHeader_(ciclosData[0]);
  const aIdx = asignData.length > 0 ? indexByHeader_(asignData[0]) : {};

  ['CicloID', 'CapacidadMaxima', 'PlazasOcupadas', 'PlazasLibres']
    .forEach(col => {
      if (cIdx[col] === undefined) {
        throw new Error('Falta columna "' + col + '" en CICLOS.');
      }
    });

  ['CicloID', 'EstadoAsignacion']
    .forEach(col => {
      if (aIdx[col] === undefined) {
        throw new Error('Falta columna "' + col + '" en ASIGNACIONES_CICLO.');
      }
    });

  const ocupacionPorCiclo = {};

  for (let i = 1; i < asignData.length; i++) {
    const row = asignData[i];

    const cicloId = String(row[aIdx.CicloID] || '');
    const estado = row[aIdx.EstadoAsignacion];

    if (!cicloId) continue;

    if (
      estado === ESTADOS_ASIGNACION.ACTIVO ||
      estado === ESTADOS_ASIGNACION.RESERVADO
    ) {
      if (!ocupacionPorCiclo[cicloId]) {
        ocupacionPorCiclo[cicloId] = 0;
      }
      ocupacionPorCiclo[cicloId]++;
    }
  }

  let actualizados = 0;

  for (let i = 1; i < ciclosData.length; i++) {
    const cicloId = String(ciclosData[i][cIdx.CicloID] || '');
    const capacidad = Number(ciclosData[i][cIdx.CapacidadMaxima] || 0);

    const ocupadas = ocupacionPorCiclo[cicloId] || 0;
    const libres = Math.max(0, capacidad - ocupadas);

    sheetCiclos.getRange(i + 1, cIdx.PlazasOcupadas + 1).setValue(ocupadas);
    sheetCiclos.getRange(i + 1, cIdx.PlazasLibres + 1).setValue(libres);

    actualizados++;
  }

  return actualizados;
}