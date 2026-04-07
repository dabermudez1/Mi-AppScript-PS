/***********************
 * BLOQUE 24
 * ELIMINAR CICLO POR ERROR
 ***********************/

function eliminarCicloPorError() {
  const html = HtmlService
    .createHtmlOutputFromFile('EliminarCicloForm')
    .setWidth(720)
    .setHeight(620);

  SpreadsheetApp.getUi().showModalDialog(html, 'Eliminar grupo/ciclo por error');
}

function obtenerCiclosEliminablesFormulario() {
  const repo = new CicloRepository();
  const data = repo.findAll() || [];
  
  return data.map(c => ({
    cicloId: String(c.CicloID || ''),
    label: `${c.Modalidad || 'SIN_MOD'} | Ciclo ${c.NumeroCiclo || '?'} | Inicio: ${formatearFecha_(c.FechaInicioCiclo)} | (${c.EstadoCiclo || ''})`
  })).sort((a, b) => String(b.cicloId).localeCompare(String(a.cicloId)));
}

function obtenerImpactoEliminacionCicloFormulario(cicloId) {
  const cicloRepo = new CicloRepository();
  const sessionRepo = new SessionRepository();
  const asignacionRepo = new AsignacionRepository();
  const patientRepo = new PatientRepository();

  const ciclo = cicloRepo.findOneBy('CicloID', cicloId);
  if (!ciclo) throw new Error('Ciclo no encontrado.');

  const sesiones = sessionRepo.findAll().filter(s => String(s.CicloID) === String(cicloId));
  const asignaciones = asignacionRepo.findAll().filter(a => String(a.CicloID) === String(cicloId));
  const pacientesAfectados = patientRepo.findAll().filter(p => 
    String(p.CicloObjetivoID) === String(cicloId) || String(p.CicloActivoID) === String(cicloId)
  );

  return {
    cicloId: ciclo.CicloID,
    modalidad: ciclo.Modalidad,
    numeroCiclo: ciclo.NumeroCiclo,
    estado: ciclo.EstadoCiclo,
    sesionesContar: sesiones.length,
    asignacionesContar: asignaciones.length,
    pacientesContar: pacientesAfectados.length,
    nombresPacientes: pacientesAfectados.map(p => p.Nombre).join(', ')
  };
}

function confirmarEliminacionCicloFormulario(formData) {
  const cicloId = String(formData.cicloId || '').trim();
  if (!cicloId) throw new Error('Debes seleccionar un ciclo.');

  const cicloRepo = new CicloRepository();
  const ciclo = cicloRepo.findOneBy('CicloID', cicloId);
  if (!ciclo) throw new Error('Ciclo no encontrado.');

  // 1. Limpiar sesiones vinculadas (Pattern B: Batch write)
  borrarSesionesPorCiclo_(cicloId);

  // 2. Limpiar asignaciones vinculadas
  borrarAsignacionesPorCiclo_(cicloId);

  // 3. Devolver pacientes a ESPERA y limpiar sus campos de ciclo
  restaurarPacientesDeCicloEliminado_(cicloId);

  // 4. Eliminar el ciclo físicamente
  const sheet = cicloRepo.getSheet();
  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);
  const filtrados = data.filter((row, i) => i === 0 || String(row[idx.CicloID]) !== cicloId);
  
  sheet.clearContents();
  sheet.getRange(1, 1, filtrados.length, data[0].length).setValues(filtrados);
  __EXECUTION_CACHE__[SHEET_CICLOS] = null;
  
  eliminarCacheDashboard_();

  return {
    mensaje: `Ciclo ${ciclo.NumeroCiclo} eliminado correctamente.\nSesiones y asignaciones eliminadas.\nPacientes devueltos a ESPERA.`
  };
}

/**
 * Borrado masivo de sesiones por CicloID
 */
function borrarSesionesPorCiclo_(cicloId) {
  const repo = new SessionRepository();
  const sheet = repo.getSheet();
  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);
  
  const filtrados = data.filter((row, i) => i === 0 || String(row[idx.CicloID]) !== String(cicloId));
  
  if (filtrados.length !== data.length) {
    sheet.clearContents();
    sheet.getRange(1, 1, filtrados.length, data[0].length).setValues(filtrados);
    __EXECUTION_CACHE__[SHEET_SESIONES] = null;
  }
}

/**
 * Borrado masivo de asignaciones por CicloID
 */
function borrarAsignacionesPorCiclo_(cicloId) {
  const repo = new AsignacionRepository();
  const sheet = repo.getSheet();
  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);
  
  const filtrados = data.filter((row, i) => i === 0 || String(row[idx.CicloID]) !== String(cicloId));
  
  if (filtrados.length !== data.length) {
    sheet.clearContents();
    sheet.getRange(1, 1, filtrados.length, data[0].length).setValues(filtrados);
    __EXECUTION_CACHE__[SHEET_ASIGNACIONES_CICLO] = null;
  }
}

/**
 * Actualiza los pacientes que estaban en el ciclo para que vuelvan a estar disponibles
 */
function restaurarPacientesDeCicloEliminado_(cicloId) {
  const repo = new PatientRepository();
  const pacientes = repo.findAll();
  const modificados = [];

  pacientes.forEach(p => {
    if (String(p.CicloObjetivoID) === String(cicloId) || String(p.CicloActivoID) === String(cicloId)) {
      p.EstadoPaciente = ESTADOS_PACIENTE.ESPERA;
      p.CicloObjetivoID = '';
      p.CicloActivoID = '';
      p.MotivoEspera = 'CICLO_ELIMINADO';
      p.FechaPrimeraSesionReal = '';
      p.SesionesPlanificadas = 0;
      p.SesionesPendientes = 0;
      p.ProximaSesion = '';
      modificados.push(p);
    }
  });

  if (modificados.length > 0) {
    repo.saveAll(modificados);
  }
}