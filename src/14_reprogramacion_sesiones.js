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
  const patientRepo = new PatientRepository();
  const paciente = patientRepo.findById(data.pacienteId);
  const sessionRepo = new SessionRepository();
  const sesionOriginal = sessionRepo.findPendientesByPaciente(data.pacienteId).find(s => Number(s.NumeroSesion) === Number(data.numeroSesion));
  if (!sesionOriginal) throw new Error("No se encontró la sesión original para reprogramar.");
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

  const sesionActualizada = sessionService.rescheduleSession(pacienteId, numeroSesion, nuevaFecha, nuevaHora);

  new StateService().refreshPatientMetrics(new PatientRepository().findById(pacienteId));

  // Refrescar caché del dashboard para actualizar incidencias y métricas en el panel
  try {
    refrescarDashboard();
  } catch (e) {
    eliminarCacheDashboard_();
  }

  return {
    mensaje: 'Sesión individual reprogramada correctamente.',
    pacienteId: pacienteId, // Devolvemos el ID para el refresco
    detalle: {
      // --- CORRECCIÓN: Usar datos de la sesión ANTES de la modificación ---
      nombre: sesionOriginal.NombrePaciente,
      numeroSesion: sesionOriginal.NumeroSesion,
      fechaAnterior: formatearFecha_(sesionOriginal.FechaSesion),
      horaAnterior: formatearHora_(sesionOriginal.HoraInicio),
      // --- Usar datos de la sesión DESPUÉS de la modificación ---
      fechaNueva: formatearFecha_(sesionActualizada.FechaSesion),
      horaNueva: formatearHora_(sesionActualizada.HoraInicio)
    }
  };
}

/***************
 * GRUPO
 ***************/
function reprogramarSesionGrupo_(data) {
  const sessionRepo = new SessionRepository();
  const patientRepo = new PatientRepository();
  const cicloRepo = new CicloRepository();
  const availabilityService = new AvailabilityService();
  const stateService = new StateService();

  const cicloId = data.cicloId;
  const numeroSesion = Number(data.numeroSesion);
  const nuevaFecha = parseFechaES_(data.fecha);
  const nuevaHora = data.hora;
  const enCascada = data.cascada === true;

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

  const ciclo = cicloRepo.findOneBy('CicloID', cicloId);
  if (!ciclo) throw new Error('Ciclo no encontrado.');

  const config = obtenerConfigModalidad_(ciclo.Modalidad);
  // La frecuencia de grupo viene en semanas (1 o 2)
  const frecuenciaSemanas = Math.max(1, Number(config.FrecuenciaDias || 1));
  const todasLasSesionesDelCiclo = sessionRepo.findByCicloId(cicloId);
  
  let sesionesAfectadasPrincipal = 0;
  let numeroSesionesMovidasCascada = 0;
  let manualesSobreescritas = 0;
  let currentBaseDate = nuevaFecha;

  const maxNumSesion = Math.max(...todasLasSesionesDelCiclo.map(s => Number(s.NumeroSesion)));
  const ultimoNumeroAProcesar = enCascada ? maxNumSesion : numeroSesion;

  for (let n = numeroSesion; n <= ultimoNumeroAProcesar; n++) {
    const sesionesDelNumero = todasLasSesionesDelCiclo.filter(s => 
      Number(s.NumeroSesion) === n && 
      (s.EstadoSesion === ESTADOS_SESION.PENDIENTE || s.EstadoSesion === ESTADOS_SESION.REPROGRAMADA)
    );

    if (sesionesDelNumero.length === 0) continue;

    let targetDate;
    let targetHora;

    if (n === numeroSesion) {
      targetDate = nuevaFecha;
      targetHora = nuevaHora;
    } else {
      // Buscamos el siguiente hueco disponible respetando la frecuencia base
      const proximaMinima = sumarSemanasManteniendoDia_(currentBaseDate, frecuenciaSemanas);
      const slot = availabilityService.findNextAvailableSlot(
        normalizarFechaHora_(proximaMinima, "00:00"), 
        ciclo.Modalidad, 
        'SEGUIMIENTO/GRUPO', // Especificamos el tipo de sesión
        90, 
        cicloId // Ignoramos colisiones con nuestras propias sesiones futuras
      );
      
      if (!slot) {
        throw new Error(`Interrupción en cascada: No se encontró un hueco libre para la sesión número ${n} a partir del ${formatearFecha_(proximaMinima)}.`);
      }
      targetDate = slot.startDateTime;
      targetHora = formatearHora_(slot.startDateTime);
      numeroSesionesMovidasCascada++;
    }

    sesionesDelNumero.forEach(sesion => {
      if (sesion.ModificadaManual && n !== numeroSesion) manualesSobreescritas++;
      if (!sesion.FechaOriginal) sesion.FechaOriginal = sesion.FechaSesion;
      
      sesion.FechaSesion = normalizarFecha_(targetDate);
      sesion.HoraInicio = targetHora;
      sesion.ModificadaManual = true;
      sessionRepo.save(sesion);
      
      const paciente = patientRepo.findById(sesion.PacienteID);
      if (paciente) stateService.refreshPatientMetrics(paciente);
      
      if (n === numeroSesion) sesionesAfectadasPrincipal++;
    });

    currentBaseDate = targetDate;
  }

  // Recalcular y actualizar la Fecha Fin del Ciclo
  const sesionesActualizadas = sessionRepo.findByCicloId(cicloId).sort((a, b) => Number(b.NumeroSesion) - Number(a.NumeroSesion));
  if (sesionesActualizadas.length > 0) {
    ciclo.FechaFinCiclo = normalizarFecha_(sesionesActualizadas[0].FechaSesion);
    cicloRepo.save(ciclo);
  }

  // Refrescar caché del dashboard para actualizar incidencias y métricas en el panel
  try {
    refrescarDashboard();
  } catch (e) {
    eliminarCacheDashboard_();
  }

  let msgResult = `Reprogramación completada.\nPacientes en sesión actual: ${sesionesAfectadasPrincipal}`;
  if (enCascada) {
    msgResult += `\nSesiones sucesivas desplazadas: ${numeroSesionesMovidasCascada}`;
    if (manualesSobreescritas > 0) {
      msgResult += `\n⚠️ Se han sobrescrito ${manualesSobreescritas} cambios manuales previos en sesiones sucesivas.`;
    }
  }

  return {
    mensaje: msgResult,
    cicloId: cicloId, // Devolvemos el ID para el refresco
    detalle: {
      nombre: `Grupo ${ciclo.Modalidad} - Ciclo ${ciclo.NumeroCiclo}`,
      numeroSesion: numeroSesion,
      fechaNueva: formatearFecha_(nuevaFecha),
      horaNueva: nuevaHora
    }
  };
}

/**
 * Obtiene los slots libres para una fecha y modalidad específica.
 * Busca en los próximos días y devuelve una estructura agrupada por fecha.
 */
function obtenerMultiplesSlotsDisponibles(fechaISO, modalidad, cicloId) {
  if (!fechaISO || !modalidad) return [];
  
  const availabilityService = new AvailabilityService();
  const agendaService = new AgendaService();
  const fecha = parseFechaISO_(fechaISO);
  if (!fecha) return [];

  const diasConSlots = [];
  let currentDay = normalizarFechaHora_(fecha, "00:00");
  let diasBuscados = 0;
  const maxDiasBusqueda = 60; // Límite de seguridad
  const diasAEncontrar = 7; // Buscamos los 7 próximos días con huecos

  const allSessions = availabilityService.sessionRepo.findAll();

  while (diasConSlots.length < diasAEncontrar && diasBuscados < maxDiasBusqueda) {
    const dateKey = obtenerClaveFecha_(currentDay);
    const agendaForDay = agendaService.getAgendaForDay(currentDay);
    
    const sessionsForDay = allSessions.filter(s => 
      s.FechaSesion instanceof Date && 
      obtenerClaveFecha_(s.FechaSesion) === dateKey &&
      s.EstadoSesion !== ESTADOS_SESION.CANCELADA &&
      (!cicloId || String(s.CicloID) !== String(cicloId))
    );
    
    const occupiedSlots = availabilityService._getOccupiedSlotsFromSessions(sessionsForDay);
    const requiredDuration = (modalidad === MODALIDADES.INDIVIDUAL) ? 30 : 90;

    const libres = agendaForDay.filter(slot => {
        return slot.type !== 'DESCANSO' &&
               availabilityService._isSlotCompatible(slot, modalidad, requiredDuration) &&
               !availabilityService._isSlotOccupied(slot, occupiedSlots);
      }).map(s => ({
        hora: formatearHora_(s.startDateTime),
        label: `${formatearHora_(s.startDateTime)} (${s.type})`
      }));

    if (libres.length > 0) {
      diasConSlots.push({
        fechaISO: formatearFechaISOInput_(currentDay),
        fechaLabel: formatearFecha_(currentDay) + " (" + convertirDiaSemanaATexto_(currentDay) + ")",
        slots: libres
      });
    }
    
    currentDay = sumarDiasNaturales_(currentDay, 1);
    diasBuscados++;
  }

  return diasConSlots;
}

function obtenerSesionesPendientesIndividualFormulario(pacienteId) {
  const sessionRepo = new SessionRepository();
  const sesiones = sessionRepo.findPendientesByPaciente(pacienteId);

  return sesiones.map(s => ({
    numeroSesion: s.NumeroSesion,
    fechaActual: formatearFecha_(s.FechaSesion),
    horaActual: formatearHora_(s.HoraInicio),
    label: `Sesión ${s.NumeroSesion} | ${formatearFecha_(s.FechaSesion)} (${formatearHora_(s.HoraInicio)})`
  })).sort((a, b) => a.numeroSesion - b.numeroSesion);
}

function obtenerSesionesPendientesGrupoFormulario(cicloId) {
  const sessionRepo = new SessionRepository();
  const sesiones = sessionRepo.findByCicloId(cicloId).filter(s => 
    s.EstadoSesion === ESTADOS_SESION.PENDIENTE || 
    s.EstadoSesion === ESTADOS_SESION.REPROGRAMADA
  );

  // Agrupar por número de sesión para el dropdown (ya que la sesión es común al grupo)
  const unicas = [...new Map(sesiones.map(s => [s.NumeroSesion, s])).values()];

  return unicas.map(s => ({
    numeroSesion: s.NumeroSesion,
    fechaActual: formatearFecha_(s.FechaSesion),
    horaActual: formatearHora_(s.HoraInicio),
    label: `Sesión ${s.NumeroSesion} | ${formatearFecha_(s.FechaSesion)} (${formatearHora_(s.HoraInicio)})`
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

  if (sesion.EstadoSesion !== ESTADOS_SESION.PENDIENTE && sesion.EstadoSesion !== ESTADOS_SESION.REPROGRAMADA) {
    throw new Error('Solo se pueden abrir para reprogramación sesiones en estado PENDIENTE o REPROGRAMADA.');
  }

  return {
    sesionId: sesion.SesionID,
    modalidad: sesion.Modalidad,
    pacienteId: sesion.PacienteID,
    cicloId: sesion.CicloID,
    numeroSesion: Number(sesion.NumeroSesion || 0),
    fechaActual: formatearFecha_(sesion.FechaSesion),
    horaActual: formatearHora_(sesion.HoraInicio),
    nombrePaciente: sesion.NombrePaciente || ''
  };
}