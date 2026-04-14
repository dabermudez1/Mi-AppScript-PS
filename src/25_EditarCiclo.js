/***********************
 * BLOQUE 25
 * EDICIÓN DE CICLOS / GRUPOS
 ***********************/

/**
 * Abre el diálogo de edición de ciclo.
 */
function editarCiclo() {
  const html = HtmlService.createHtmlOutputFromFile('EditarCicloForm')
    .setWidth(500)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Editar grupo / Cerrar cupo');
}

/**
 * Obtiene la lista de ciclos editables para el selector.
 */
function obtenerCiclosEdicionFormulario() {
  const repo = new CicloRepository();
  const data = repo.findAll() || [];
  
  // Mostramos solo ciclos que no estén terminados o cancelados
  return data
    .filter(c => c.EstadoCiclo === ESTADOS_CICLO.PLANIFICADO || c.EstadoCiclo === ESTADOS_CICLO.EN_CURSO)
    .map(c => ({
      cicloId: String(c.CicloID || ''),
      label: `${c.Modalidad} | Ciclo ${c.NumeroCiclo} | Inicio: ${formatearFecha_(c.FechaInicioCiclo)} (${c.EstadoCiclo})`
    }))
    .sort((a, b) => b.label.localeCompare(a.label));
}

/**
 * Obtiene los detalles actuales de un ciclo para cargar el formulario.
 */
function obtenerDetalleCicloEdicion(cicloId) {
  const repo = new CicloRepository();
  const ciclo = repo.findOneBy('CicloID', cicloId);
  if (!ciclo) throw new Error('Ciclo no encontrado.');

  return {
    cicloId: ciclo.CicloID,
    modalidad: ciclo.Modalidad,
    numeroCiclo: ciclo.NumeroCiclo,
    capacidadMaxima: Number(ciclo.CapacidadMaxima || 0),
    bloqueoInscripciones: ciclo.BloqueoInscripciones === true || ciclo.BloqueoInscripciones === 'TRUE',
    notas: ciclo.Notas || '',
    plazasOcupadas: Number(ciclo.PlazasOcupadas || 0)
  };
}

/**
 * Guarda los cambios realizados en el ciclo.
 */
function guardarEdicionCiclo(formData) {
  const repo = new CicloRepository();
  const ciclo = repo.findOneBy('CicloID', formData.cicloId);
  if (!ciclo) throw new Error('Ciclo no encontrado.');

  const nuevaCapacidad = Number(formData.capacidadMaxima);
  const ocupadas = Number(ciclo.PlazasOcupadas || 0);

  if (nuevaCapacidad < ocupadas) {
    throw new Error(`La capacidad (${nuevaCapacidad}) no puede ser menor a las plazas ya ocupadas (${ocupadas}).`);
  }

  ciclo.CapacidadMaxima = nuevaCapacidad;
  ciclo.PlazasLibres = nuevaCapacidad - ocupadas;
  ciclo.BloqueoInscripciones = formData.bloqueoInscripciones === true;
  ciclo.Notas = formData.notas;

  repo.save(ciclo);
  
  // Forzar recálculo y limpieza de caché
  new MaintenanceService().recalculateCycleOccupancy();
  if (typeof __EXECUTION_CACHE__ !== 'undefined') __EXECUTION_CACHE__[SHEET_CICLOS] = null;
  eliminarCacheDashboard_();

  return {
    mensaje: `Ciclo ${ciclo.NumeroCiclo} actualizado correctamente.`
  };
}