/***********************
 * BLOQUE 14
 * REPROGRAMACIÓN SESIONES
 ***********************/

function abrirReprogramarSesion() {
  const html = HtmlService
    .createHtmlOutputFromFile('ReprogramarSesionForm')
    .setWidth(500)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Reprogramar sesión');
}

/***************
 * DATOS FORMULARIO
 ***************/
function obtenerDatosReprogramacion() {
  const patientRepo = new PatientRepository();
  const cicloRepo = new CicloRepository();

  const pacientesIndividuales = patientRepo.findAll()
    .filter(p => p.ModalidadSolicitada === MODALIDADES.INDIVIDUAL)
    .map(p => ({ id: p.PacienteID, nombre: p.Nombre }));

  const ciclosGrupo = cicloRepo.findAll()
    .filter(c => c.Modalidad !== MODALIDADES.INDIVIDUAL)
    .map(c => ({
      id: c.CicloID,
      label: `${c.Modalidad} | Ciclo ${c.NumeroCiclo}`
    }));

  return {
    modalidades: Object.values(MODALIDADES),
    pacientes: pacientesIndividuales,
    ciclos: ciclosGrupo
  };
}

/***************
 * GUARDAR
 ***************/
function guardarReprogramacion(data) {
  const modalidad = data.modalidad;

  if (modalidad === MODALIDADES.INDIVIDUAL) {
    return reprogramarSesionIndividual_(data);
  }

  return reprogramarSesionGrupo_(data);
}

/***************
 * INDIVIDUAL
 ***************/
function reprogramarSesionIndividual_(data) {
  const sessionService = new SessionService();

  const pacienteId = data.pacienteId;
  const numeroSesion = Number(data.numeroSesion);
  const nuevaFecha = parseFechaES_(data.fecha);
  const nuevaHora = data.hora; // Capturamos la hora del formulario

  if (!nuevaFecha) {
    throw new Error('La nueva fecha no es válida. Usa formato dd/mm/yyyy.');
  }

  if (!esFechaOperativaValida_(nuevaFecha)) {
    throw new Error(construirMensajeFechaNoOperativa_(nuevaFecha));
  }

  sessionService.rescheduleSession(pacienteId, numeroSesion, nuevaFecha, nuevaHora);

  new StateService().refreshPatientMetrics(new PatientRepository().findById(pacienteId));

  return {
    mensaje: 'Sesión individual reprogramada correctamente.'
  };
}

/***************
 * GRUPO
 ***************/
function reprogramarSesionGrupo_(data) {
  const sessionRepo = new SessionRepository();
  const patientRepo = new PatientRepository();
  const stateService = new StateService();

  const cicloId = data.cicloId;
  const numeroSesion = Number(data.numeroSesion);
  const nuevaFecha = parseFechaES_(data.fecha);
  const nuevaHora = data.hora;

  if (!nuevaFecha) {
    throw new Error('La nueva fecha no es válida. Usa formato dd/mm/yyyy.');
  }

  if (!esFechaOperativaValida_(nuevaFecha)) {
  throw new Error(construirMensajeFechaNoOperativa_(nuevaFecha));
  }

  const diaSemanaCiclo = obtenerDiaSemanaCiclo_(cicloId);

  if (!fechaCoincideConDiaSemana_(nuevaFecha, diaSemanaCiclo)) {
    throw new Error(
      'La nueva fecha no respeta el día fijo del grupo.\n\n' +
      'Día esperado: ' + diaSemanaCiclo + '\n' +
      'Fecha propuesta: ' + formatearFecha_(nuevaFecha)
    );
  }

  let cambios = 0;
  const sesiones = sessionRepo.findByCicloId(cicloId).filter(s => 
    Number(s.NumeroSesion) === numeroSesion && 
    s.EstadoSesion === ESTADOS_SESION.PENDIENTE
  );

  sesiones.forEach(sesion => {
    if (!sesion.FechaOriginal) {
      sesion.FechaOriginal = sesion.FechaSesion;
    }
    sesion.FechaSesion = nuevaFecha;
    if (nuevaHora) sesion.HoraInicio = nuevaHora;
    sesion.ModificadaManual = true;
    sessionRepo.save(sesion);
    
    // Actualizar métricas del paciente afectado
    const paciente = patientRepo.findById(sesion.PacienteID);
    if (paciente) stateService.refreshPatientMetrics(paciente);
    
    cambios++;
  });

  return {
    mensaje: 'Sesión grupal reprogramada.\nSesiones afectadas: ' + cambios
  };
}

function obtenerSesionesPendientesIndividualFormulario(pacienteId) {
  const sessionRepo = new SessionRepository();
  const sesiones = sessionRepo.findPendientesByPaciente(pacienteId);

  return sesiones.map(s => ({
    numeroSesion: s.NumeroSesion,
    fechaActual: formatearFecha_(s.FechaSesion),
    label: `Sesión ${s.NumeroSesion} | ${formatearFecha_(s.FechaSesion)}`
  })).sort((a, b) => a.numeroSesion - b.numeroSesion);
}

function obtenerSesionesPendientesGrupoFormulario(cicloId) {
  const sessionRepo = new SessionRepository();
  const sesiones = sessionRepo.findByCicloId(cicloId).filter(s => s.EstadoSesion === ESTADOS_SESION.PENDIENTE);

  // Agrupar por número de sesión para el dropdown (ya que la sesión es común al grupo)
  const unicas = [...new Map(sesiones.map(s => [s.NumeroSesion, s])).values()];

  return unicas.map(s => ({
    numeroSesion: s.NumeroSesion,
    fechaActual: formatearFecha_(s.FechaSesion),
    label: `Sesión ${s.NumeroSesion} | ${formatearFecha_(s.FechaSesion)}`
  })).sort((a, b) => a.numeroSesion - b.numeroSesion);
}

function obtenerDiaSemanaCiclo_(cicloId) {
  const ciclo = new CicloRepository().findOneBy('CicloID', cicloId);
  if (!ciclo) throw new Error('No se encontró el ciclo indicado.');
  return ciclo.DiaSemana || '';
}

function abrirReprogramarSesionDesdeSesion(sesionId) {
  if (!sesionId) {
    throw new Error('No se indicó la sesión a reprogramar.');
  }

  const template = HtmlService.createTemplateFromFile('ReprogramarSesionForm');
  template.sesionPreseleccionadaId = String(sesionId);
  template.origenPantalla = 'SESIONES';

  const html = template
    .evaluate()
    .setWidth(500)
    .setHeight(540);

  SpreadsheetApp.getUi().showModalDialog(html, 'Reprogramar sesión');
}

function abrirReprogramarSesionDesdeIncidencia(sesionId) {
  if (!sesionId) {
    throw new Error('No se indicó la sesión a reprogramar.');
  }

  const template = HtmlService.createTemplateFromFile('ReprogramarSesionForm');
  template.sesionPreseleccionadaId = String(sesionId);
  template.origenPantalla = 'INCIDENCIAS';

  const html = template
    .evaluate()
    .setWidth(500)
    .setHeight(540);

  SpreadsheetApp.getUi().showModalDialog(html, 'Reprogramar sesión');
}

function obtenerSesionParaReprogramacionFormulario(sesionId) {
  const sessionRepo = new SessionRepository();
  const sesion = sessionRepo.findOneBy('SesionID', sesionId);

  if (!sesion) {
    throw new Error('No se encontró la sesión indicada.');
  }

  if (sesion.EstadoSesion !== ESTADOS_SESION.PENDIENTE) {
    throw new Error('Solo se pueden abrir para reprogramación sesiones en estado PENDIENTE.');
  }

  return {
    sesionId: sesion.SesionID,
    modalidad: sesion.Modalidad,
    pacienteId: sesion.PacienteID,
    cicloId: sesion.CicloID,
    numeroSesion: Number(sesion.NumeroSesion || 0),
    fechaActual: formatearFecha_(sesion.FechaSesion),
    nombrePaciente: sesion.NombrePaciente || ''
  };
}