/***********************
 * BLOQUE 3
 * GESTIÓN DE PACIENTES
 ***********************/

function nuevoPaciente() {
  const html = HtmlService
    .createHtmlOutputFromFile('NuevoPacienteForm')
    .setWidth(420)
    .setHeight(850);

  SpreadsheetApp.getUi().showModalDialog(html, 'Nuevo paciente');
}

/**
 * Abre el formulario para reservar una primera consulta (2.1).
 */
function reservarPrimeraConsulta() {
  const html = HtmlService
    .createHtmlOutputFromFile('ReservarPrimeraForm')
    .setWidth(420)
    .setHeight(550); // Altura ajustada para este formulario
  SpreadsheetApp.getUi().showModalDialog(html, 'Reservar Primera Consulta (2.1)');
}

/**
 * Abre el formulario para editar o eliminar una reserva de primera consulta (2.1).
 */
function editarEliminarReserva21() {
  const html = HtmlService
    .createHtmlOutputFromFile('EditarReserva21Form')
    .setWidth(420)
    .setHeight(750); // Altura ajustada
  SpreadsheetApp.getUi().showModalDialog(html, 'Editar/Eliminar Reserva 2.1');
}

/***************
 * ENTRADAS UI
 ***************/
function pedirTextoObligatorio_(ui, titulo, mensaje) {
  const resp = ui.prompt(titulo, mensaje, ui.ButtonSet.OK_CANCEL);

  if (resp.getSelectedButton() !== ui.Button.OK) return null;

  const valor = (resp.getResponseText() || '').trim();
  if (!valor) {
    ui.alert('El valor es obligatorio.');
    return null;
  }

  return valor;
}

function pedirModalidadPaciente_(ui) {
  const mensaje =
    'Selecciona modalidad:\n\n' +
    '1 = INDIVIDUAL\n' +
    '2 = GRUPO_1\n' +
    '3 = GRUPO_2\n' +
    '4 = GRUPO_3';

  const resp = ui.prompt('Modalidad', mensaje, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return null;

  const valor = (resp.getResponseText() || '').trim();

  if (valor === '1') return MODALIDADES.INDIVIDUAL;
  if (valor === '2') return MODALIDADES.GRUPO_1;
  if (valor === '3') return MODALIDADES.GRUPO_2;
  if (valor === '4') return MODALIDADES.GRUPO_3;

  ui.alert('Modalidad no válida.');
  return null;
}

/***************
 * ORQUESTACIÓN
 ***************/
function crearPacienteSegunModalidad_({
    nombre,
    nhc,
    sexoGenero,
    motivoConsultaDiagnostico,
    motivoConsultaOtros,
    modalidad,
    fechaPrimeraConsulta,
    slotSeguimiento // Nuevo: Objeto { fecha, hora } seleccionado manualmente
  }) {
  const config = obtenerConfigModalidad_(modalidad);

  // 1. Caso INDIVIDUAL: Verificamos cupo y calculamos primera sesión
  if (config.TipoModalidad === TIPOS_MODALIDAD.INDIVIDUAL) {
    // Validación de seguridad: Para individuales es obligatorio haber elegido slot
    if (!slotSeguimiento || !slotSeguimiento.fecha || !slotSeguimiento.hora) {
      throw new Error('Es obligatorio seleccionar la primera sesión de seguimiento para pacientes individuales.');
    }

    if (!hayCapacidadIndividual_()) {
      const pacienteId = crearPacienteEnSheet_({
        nombre, nhc, sexoGenero, motivoConsultaDiagnostico, motivoConsultaOtros,
        modalidadSolicitada: modalidad, fechaPrimeraConsulta,
        estadoPaciente: ESTADOS_PACIENTE.ESPERA, motivoEspera: 'SIN_CUPO_INDIVIDUAL',
        sesionesPlanificadas: Number(config.SesionesPorCiclo || 0),
        sesionesPendientes: Number(config.SesionesPorCiclo || 0)
      });
      return {
        pacienteId,
        mensaje: `Paciente creado en ESPERA.\nMotivo: SIN_CUPO_INDIVIDUAL`
      };
    }

    const fechaHoraInicio = normalizarFechaHora_(slotSeguimiento.fecha, slotSeguimiento.hora);

    const pacienteId = crearPacienteEnSheet_({
      nombre, nhc, sexoGenero, motivoConsultaDiagnostico, motivoConsultaOtros,
      modalidadSolicitada: modalidad, fechaPrimeraConsulta,
      estadoPaciente: ESTADOS_PACIENTE.ACTIVO,
      fechaPrimeraSesionReal: fechaHoraInicio,
      proximaSesion: fechaHoraInicio,
      sesionesPlanificadas: Number(config.SesionesPorCiclo || 0),
      sesionesPendientes: Number(config.SesionesPorCiclo || 0)
    });

    const patientRepo = new PatientRepository();
    const pCompleto = patientRepo.findById(pacienteId); // Forzar recarga
    if (pCompleto) generarSesionesPacienteIndividual_(pacienteId, fechaHoraInicio);
    
    return {
      pacienteId,
      mensaje: `Paciente creado correctamente.\nEstado: ACTIVO\nPrimera sesión: ${formatearFecha_(fechaHoraInicio)} a las ${formatearHora_(fechaHoraInicio)}`
    };
  }

  // 2. Caso GRUPO: Delegamos en la lógica ya existente de grupos
  return procesarAltaGrupo_({
    nombre, nhc, sexoGenero, motivoConsultaDiagnostico, motivoConsultaOtros,
    modalidad, fechaPrimeraConsulta, config
  });
}

function calcularPrimeraSesionIndividual_(fechaPrimeraConsulta, modalidad) {
  const config = obtenerConfigModalidad_(modalidad);
  const intervaloDias = Number(config.FrecuenciaDias || 0);

  if (intervaloDias <= 0) {
    throw new Error(
      'La frecuencia de la modalidad individual no es válida.\n\n' +
      'Modalidad: ' + modalidad
    );
  }

  const availabilityService = new AvailabilityService();
  // Buscar el primer slot '2.2' a partir de la fecha de primera consulta + intervalo
  let startSearchDate = sumarDiasNaturales_(fechaPrimeraConsulta, intervaloDias);

  // AJUSTE: Garantizar que no se proponen citas en el pasado
  const ahora = normalizarFechaHora_(new Date());
  if (startSearchDate.getTime() < ahora.getTime()) {
    startSearchDate = ahora;
  }

  const slot = availabilityService.findNextAvailableSlot(startSearchDate, modalidad, 30);

  return slot ? { fecha: slot.startDateTime, hora: formatearHora_(slot.startDateTime) } : null;
}

function hayCapacidadIndividual_() {
  const config = obtenerConfigModalidad_(MODALIDADES.INDIVIDUAL);
  const capacidadMaxima = Number(config.CapacidadMaxima || 0);
  if (capacidadMaxima <= 0) return false;

  // Usar repositorio para aprovechar la caché de ejecución
  const repo = new PatientRepository();
  const pacientes = repo.findAll();
  if (pacientes.length === 0) return true;

  const activos = pacientes.filter(p => 
    p.ModalidadSolicitada === MODALIDADES.INDIVIDUAL && 
    p.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO
  ).length;

  return activos < capacidadMaxima;
}

/***************
 * GRUPOS
 ***************/
function procesarAltaGrupo_({
    nombre,
    nhc,
    sexoGenero,
    motivoConsultaDiagnostico,
    motivoConsultaOtros,
    modalidad,
    fechaPrimeraConsulta,
    config
  }) {
  validarConfigGrupo_(modalidad, config);

  let ciclo = buscarPrimerCicloFuturoDisponible_(modalidad, fechaPrimeraConsulta);
  
  // Si el ciclo encontrado tiene el cupo bloqueado manualmente, lo ignoramos para nuevas altas
  if (ciclo && (ciclo.BloqueoInscripciones === true || ciclo.BloqueoInscripciones === 'TRUE')) {
    ciclo = null;
  }

  if (!ciclo) {
    const motivo = existeCicloFuturoPeroLleno_(modalidad, fechaPrimeraConsulta)
      ? 'SIN_PLAZA_CICLO'
      : 'SIN_CICLO_DISPONIBLE';

    const pacienteId = crearPacienteEnSheet_({
      nombre,
      nhc,
      sexoGenero,
      motivoConsultaDiagnostico,
      motivoConsultaOtros,
      modalidadSolicitada: modalidad,
      fechaPrimeraConsulta,
      estadoPaciente: ESTADOS_PACIENTE.ESPERA,
      motivoEspera: motivo,
      cicloObjetivoId: '',
      cicloActivoId: '',
      fechaPrimeraSesionReal: '',
      sesionesPlanificadas: Number(config.SesionesPorCiclo || 7),
      sesionesCompletadas: 0,
      sesionesPendientes: Number(config.SesionesPorCiclo || 7),
      proximaSesion: '',
      fechaCierre: '',
      observaciones: '',
      recalcularSecuencia: false
    });

    return {
      pacienteId,
      mensaje:
        'Paciente creado en ESPERA.\n\n' +
        'Nombre: ' + nombre + '\n' +
        'Modalidad: ' + modalidad + '\n' +
        'Motivo: ' + motivo
    };
  }

  // Revalidación final de capacidad real del ciclo en el momento de reservar.
  const reservaOk = actualizarPlazasCiclo_(ciclo.CicloID, +1, true);

  if (!reservaOk) {
    const pacienteId = crearPacienteEnSheet_({
      nombre,
      nhc,
      sexoGenero,
      motivoConsultaDiagnostico,
      motivoConsultaOtros,
      modalidadSolicitada: modalidad,
      fechaPrimeraConsulta,
      estadoPaciente: ESTADOS_PACIENTE.ESPERA,
      motivoEspera: 'SIN_PLAZA_CICLO',
      cicloObjetivoId: '',
      cicloActivoId: '',
      fechaPrimeraSesionReal: '',
      sesionesPlanificadas: Number(config.SesionesPorCiclo || 7),
      sesionesCompletadas: 0,
      sesionesPendientes: Number(config.SesionesPorCiclo || 7),
      proximaSesion: '',
      fechaCierre: '',
      observaciones: '',
      recalcularSecuencia: false
    });

    return {
      pacienteId,
      mensaje:
        'Paciente creado en ESPERA.\n\n' +
        'Nombre: ' + nombre + '\n' +
        'Modalidad: ' + modalidad + '\n' +
        'Motivo: SIN_PLAZA_CICLO'
    };
  }

  const pacienteId = crearPacienteEnSheet_({
    nombre,
    nhc,
    sexoGenero,
    motivoConsultaDiagnostico,
    motivoConsultaOtros,
    modalidadSolicitada: modalidad,
    fechaPrimeraConsulta,
    estadoPaciente: ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO,
    motivoEspera: '',
    cicloObjetivoId: ciclo.CicloID,
    cicloActivoId: '',
    fechaPrimeraSesionReal: ciclo.FechaInicioCiclo,
    sesionesPlanificadas: Number(ciclo.SesionesPorCiclo || config.SesionesPorCiclo || 7),
    sesionesCompletadas: 0,
    sesionesPendientes: Number(ciclo.SesionesPorCiclo || config.SesionesPorCiclo || 7),
    proximaSesion: ciclo.FechaInicioCiclo,
    fechaCierre: '',
    observaciones: '',
    recalcularSecuencia: false
  });

  crearAsignacionCiclo_({
    pacienteId,
    cicloId: ciclo.CicloID,
    modalidad,
    estadoAsignacion: ESTADOS_ASIGNACION.RESERVADO,
    observaciones: 'Asignación automática al crear paciente'
  });

  // Forzar guardado de asignación antes de generar sesiones
  SpreadsheetApp.flush();
  if (typeof __EXECUTION_CACHE__ !== 'undefined') __EXECUTION_CACHE__[SHEET_ASIGNACIONES_CICLO] = null;

  // Generar sesiones automáticas (sin try-catch para ver errores reales en desarrollo)
  generarSesionesPacienteGrupo_(pacienteId, ciclo.CicloID);
  
  return {
    pacienteId,
    mensaje:
      'Paciente creado y asignado a ciclo.\n\n' +
      'Nombre: ' + nombre + '\n' +
      'Modalidad: ' + modalidad + '\n' +
      'Estado: ACTIVO_PENDIENTE_INICIO\n' +
      'CicloID: ' + ciclo.CicloID + '\n' +
      'Inicio ciclo: ' + formatearFecha_(ciclo.FechaInicioCiclo)
  };
}

/**
 * Genera las sesiones para un paciente individual.
 * @param {string} pacienteId - ID del paciente.
 * @param {Date} fechaHoraPrimeraSesion - La fecha y hora pactada para la primera cita 2.2.
 */
function generarSesionesPacienteIndividual_(pacienteId, fechaHoraPrimeraSesion) {
  const patientRepo = new PatientRepository();
  const sessionRepo = new SessionRepository();
  const availabilityService = new AvailabilityService();
  const paciente = patientRepo.findById(pacienteId);

  if (!paciente) throw new Error('Paciente no encontrado: ' + pacienteId);
  if (paciente.ModalidadSolicitada !== MODALIDADES.INDIVIDUAL) {
    throw new Error('Esta función es solo para pacientes individuales.');
  }

  const config = obtenerConfigModalidad_(paciente.ModalidadSolicitada);
  const sesionesPlanificadas = Number(config.SesionesPorCiclo || 0);
  const frecuenciaDias = Number(config.FrecuenciaDias || 0);
  const duracionSlot = 30; // Minutos estándar para 2.2

  if (!fechaHoraPrimeraSesion) throw new Error('Se requiere la fecha de la primera sesión para generar el ciclo.');

  if (sesionesPlanificadas <= 0) {
    console.warn('No hay sesiones planificadas configuradas para esta modalidad.');
    return;
  }

  // Limpiar cualquier sesión previa por error para evitar duplicados si se reintenta
  borrarSesionesPaciente_(pacienteId);

  let currentSearchDateTime = fechaHoraPrimeraSesion;
  
  const generatedSessions = [];

  for (let i = 0; i < sesionesPlanificadas; i++) {
    let nextSlot;
    
    if (i === 0) {
      // La primera sesión es la que ya eligió el usuario
      nextSlot = { startDateTime: currentSearchDateTime, durationMinutes: duracionSlot };
    } else {
      // Las siguientes se buscan automáticamente
      nextSlot = availabilityService.findNextAvailableSlot(
        currentSearchDateTime,
        paciente.ModalidadSolicitada,
        duracionSlot
      );
    }

    if (!nextSlot) {
      console.error(`CRÍTICO: No hay slots para la sesión ${i+1}. Búsqueda falló en ${currentSearchDateTime}`);
      throw new Error(`Error de Planificación: No hay huecos libres en la agenda para la sesión ${i + 1} de ${paciente.Nombre}. ` +
                      `Revisa la 'Plantilla de Agenda' y que existan slots de tipo '2.2' o 'SEGUIMIENTO'. ` +
                      `Búsqueda iniciada en: ${formatearFecha_(currentSearchDateTime)}`);
    }

    const sesionId = generarId_('SES');
    const nuevaSesion = {
      SesionID: sesionId,
      PacienteID: paciente.PacienteID,
      CicloID: '', // No aplica para individuales
      AsignacionID: '', // No aplica para individuales
      Modalidad: paciente.ModalidadSolicitada,
      NombrePaciente: paciente.Nombre,
      NumeroSesion: i + 1,
      FechaSesion: normalizarFecha_(nextSlot.startDateTime),
      EstadoSesion: ESTADOS_SESION.PENDIENTE,
      FechaOriginal: normalizarFecha_(nextSlot.startDateTime),
      ModificadaManual: false,
      Notas: '',
      CalendarEventId: '',
      CalendarSyncStatus: '',
      CalendarLastSync: '',
      CalendarEventTitle: '',
      CalendarHash: '',
      HoraInicio: formatearHora_(nextSlot.startDateTime),
      Duracion: nextSlot.durationMinutes || duracionSlot
    };
    generatedSessions.push(nuevaSesion);

    // IMPORTANTE: Para la siguiente sesión del ciclo, saltamos los días de frecuencia
    // y reseteamos la hora a "00:00" para que busque desde el principio del día del siguiente hueco.
    let proximaFechaBusqueda = sumarDiasNaturales_(nextSlot.startDateTime, frecuenciaDias);
    currentSearchDateTime = normalizarFechaHora_(proximaFechaBusqueda, "00:00");
  }

  if (generatedSessions.length > 0) {
    console.log(`Insertando ${generatedSessions.length} sesiones en la hoja.`);
    sessionRepo.insertAll(generatedSessions);
    SpreadsheetApp.flush(); 
  } else {
    console.warn("No se generaron sesiones. Verifica la configuración de la modalidad.");
  }

  // Actualizar la próxima sesión del paciente
  if (generatedSessions.length > 0) {
    paciente.ProximaSesion = normalizarFechaHora_(generatedSessions[0].FechaSesion, generatedSessions[0].HoraInicio);
    paciente.SesionesPlanificadas = sesionesPlanificadas;
    paciente.SesionesPendientes = sesionesPlanificadas;
    patientRepo.save(paciente);
  }
}

/**
 * Genera las sesiones para un paciente de grupo.
 * @param {string} pacienteId - ID del paciente.
 * @param {string} cicloId - ID del ciclo al que está asignado el paciente.
 */
function generarSesionesPacienteGrupo_(pacienteId, cicloId) {
  const patientRepo = new PatientRepository();
  const cicloRepo = new CicloRepository();
  const sessionService = new SessionService();

  const paciente = patientRepo.findById(pacienteId);
  const ciclo = cicloRepo.findOneBy('CicloID', cicloId);
  
  if (!paciente || !ciclo) {
    Logger.log(`CRÍTICO: No se pudo generar sesiones de grupo. Paciente: ${pacienteId}, Ciclo: ${cicloId}`);
    return;
  }

  const config = obtenerConfigModalidad_(paciente.ModalidadSolicitada);
  const horaBase = config.HoraBase || '09:00';
  
  console.log(`Generando sesiones grupales para paciente ${pacienteId} en ciclo ${cicloId}`);
  
  // ¡IMPORTANTE! Usamos los datos del CICLO ya guardado para no tener que buscar slots de nuevo
  const slots = generarSlotsCiclo_({ 
    fechaInicio: ciclo.FechaInicioCiclo,
    horaInicio: ciclo.HoraInicio || horaBase,
    modalidad: paciente.ModalidadSolicitada
  });

  if (slots.length > 0) {
    const fechas = slots.map(s => s.startDateTime);
    console.log(`Slots encontrados: ${fechas.length}. Procediendo a crear sesiones.`);
    sessionService.createInitialSessions(paciente, fechas, cicloId);
    SpreadsheetApp.flush(); // Forzar sincronización con la hoja
  }
}

function buscarPrimerCicloFuturoDisponible_(modalidad, fechaPrimeraConsulta) {
  const repo = new CicloRepository();
  return repo.findNextAvailable(modalidad, fechaPrimeraConsulta);
}

function existeCicloFuturoPeroLleno_(modalidad, fechaPrimeraConsulta) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CICLOS);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_CICLOS + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return false;

  const headers = data[0];
  const idx = indexByHeader_(headers);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowModalidad = row[idx.Modalidad];
    const estadoCiclo = row[idx.EstadoCiclo];
    const fechaInicio = row[idx.FechaInicioCiclo];

    if (rowModalidad !== modalidad) continue;
    if (estadoCiclo !== ESTADOS_CICLO.PLANIFICADO) continue;
    if (!(fechaInicio instanceof Date)) continue;
    if (!(normalizarFecha_(fechaInicio) > normalizarFecha_(fechaPrimeraConsulta))) continue;

    return true;
  }

  return false;
}

/***************
 * ESCRITURA EN SHEETS
 ***************/
function crearPacienteEnSheet_(params) {
  const repo = new PatientRepository();
  const pacienteId = generarId_('PAC');
  const fechaAlta = normalizarFecha_(new Date());

  const nuevoPaciente = {
    PacienteID: pacienteId,
    Nombre: params.nombre,
    NHC: params.nhc || '',
    SexoGenero: params.sexoGenero || '',
    MotivoConsultaDiagnostico: params.motivoConsultaDiagnostico || '',
    MotivoConsultaOtros: params.motivoConsultaOtros || '',
    ModalidadSolicitada: params.modalidadSolicitada,
    FechaAlta: fechaAlta,
    FechaPrimeraConsulta: normalizarFecha_(params.fechaPrimeraConsulta),
    EstadoPaciente: params.estadoPaciente,
    MotivoEspera: params.motivoEspera || '',
    CicloObjetivoID: params.cicloObjetivoId || '',
    CicloActivoID: params.cicloActivoId || '',
    FechaPrimeraSesionReal: params.fechaPrimeraSesionReal,
    SesionesPlanificadas: Number(params.sesionesPlanificadas || 0),
    SesionesCompletadas: 0,
    SesionesPendientes: Number(params.sesionesPendientes || 0),
    ProximaSesion: params.proximaSesion,
    FechaCierre: params.fechaCierre || '',
    FechaAltaEfectiva: '',
    MotivoAltaCodigo: '',
    MotivoAltaTexto: '',
    ComentarioAlta: '',
    Observaciones: params.observaciones || '',
    RecalcularSecuencia: params.recalcularSecuencia === true
  };

  repo.save(nuevoPaciente);
  SpreadsheetApp.flush(); // Forzar escritura inmediata
  if (typeof __EXECUTION_CACHE__ !== 'undefined') __EXECUTION_CACHE__[SHEET_PACIENTES] = null;
  eliminarCacheDashboard_(); // Limpiar caché de UI para que el nuevo paciente aparezca
  return pacienteId;
}

function crearAsignacionCiclo_({ pacienteId, cicloId, modalidad, estadoAsignacion, observaciones }) {
  const repo = new BaseRepository(SHEET_ASIGNACIONES_CICLO, HEADERS[SHEET_ASIGNACIONES_CICLO]);
  const asignacionId = generarId_('ASI');

  repo.save({
    AsignacionID: asignacionId,
    PacienteID: pacienteId,
    CicloID: cicloId,
    Modalidad: modalidad,
    FechaAsignacion: normalizarFecha_(new Date()),
    EstadoAsignacion: estadoAsignacion,
    Observaciones: observaciones || ''
  });
  return asignacionId;
}

function actualizarPlazasCiclo_(cicloId, delta, devolverFalseSiNoHayPlaza) {
  const repo = new CicloRepository();
  const ciclo = repo.findOneBy('CicloID', cicloId);
  
  if (!ciclo) throw new Error('No existe el ciclo con ID ' + cicloId + '.');

  const capacidadMaxima = Number(ciclo.CapacidadMaxima || 0);
  const ocupadasActuales = Number(ciclo.PlazasOcupadas || 0);
  const nuevasOcupadas = ocupadasActuales + delta;

  if (nuevasOcupadas < 0) throw new Error('Las plazas ocupadas no pueden quedar en negativo.');
  
  if (nuevasOcupadas > capacidadMaxima) {
    if (devolverFalseSiNoHayPlaza === true) return false;
    throw new Error('El ciclo supera la capacidad máxima.');
  }

  ciclo.PlazasOcupadas = nuevasOcupadas;
  ciclo.PlazasLibres = capacidadMaxima - nuevasOcupadas;

  repo.save(ciclo);
  return true;
}

function obtenerOpcionesModalidadFormulario() {
  return [
    { value: MODALIDADES.INDIVIDUAL, label: 'INDIVIDUAL' },
    { value: MODALIDADES.GRUPO_1, label: 'GRUPO_1' },
    { value: MODALIDADES.GRUPO_2, label: 'GRUPO_2' },
    { value: MODALIDADES.GRUPO_3, label: 'GRUPO_3' }
  ];
}

/**
 * Busca slots disponibles para la reserva de primera consulta (2.1).
 * Es similar a la búsqueda de seguimiento, pero para slots de tipo '2.1' o 'PRIMERA'.
 * @param {string} fechaInicioISO - Fecha en formato YYYY-MM-DD para iniciar la búsqueda.
 * @param {number} diasABuscar - Número de días hacia adelante para buscar.
 * @returns {Array<Object>} Un array de días con slots disponibles.
 */
function obtenerSlotsDisponiblesParaReserva21(fechaInicioISO, diasABuscar = 7) {
  const fechaInicio = parseFechaISO_(fechaInicioISO);
  if (!fechaInicio) {
    throw new Error("La fecha de inicio proporcionada no es válida.");
  }

  const availabilityService = new AvailabilityService();
  const agendaService = new AgendaService();
  const diasConSlots = [];
  let currentDay = normalizarFechaHora_(fechaInicio, "00:00");
  let diasBuscados = 0;
  const maxDiasBusqueda = 60; // Límite de seguridad

  while (diasConSlots.length < diasABuscar && diasBuscados < maxDiasBusqueda) {
    const slotsDelDia = agendaService.getAgendaForDay(currentDay);
    const sessionsRepo = new SessionRepository();
    const sesionesExistentes = sessionsRepo.findAll().filter(s => 
      normalizarFecha_(s.FechaSesion).getTime() === normalizarFecha_(currentDay).getTime()
    );
    
    const occupied = availabilityService._getOccupiedSlotsFromSessions(sesionesExistentes);
    
    const libres = slotsDelDia.filter(slot => {
      // Solo slots compatibles con primera consulta (2.1)
      if (slot.type !== '2.1' && slot.type !== 'PRIMERA') return false;
      // RESTRICCIÓN: Las sesiones 2.1 solo pueden empezar en horas en punto (00 minutos)
      if (slot.startDateTime.getMinutes() !== 0) return false;
      return !availabilityService._isSlotOccupied(slot, occupied);
    }).map(s => ({
      hora: formatearHora_(s.startDateTime),
      label: formatearHora_(s.startDateTime)
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

/**
 * Busca slots disponibles para la primera sesión de seguimiento (2.2).
 * Devuelve los primeros 7 días que tengan al menos un slot libre.
 */
function obtenerSlotsDisponiblesParaSeguimiento(fechaConsultaISO, modalidad) {
  const fechaConsulta = parseFechaISO_(fechaConsultaISO);
  const config = obtenerConfigModalidad_(modalidad);
  const frecuencia = Number(config.FrecuenciaDias || 15);
  
  const availabilityService = new AvailabilityService();
  const agendaService = new AgendaService();
  
  // Empezamos a buscar desde FechaConsulta + Frecuencia
  let startSearch = sumarDiasNaturales_(fechaConsulta, frecuencia);
  const hoy = new Date();
  if (startSearch < hoy) startSearch = hoy;

  const diasConSlots = [];
  let currentDay = normalizarFechaHora_(startSearch, "00:00");
  let diasBuscados = 0;
  const maxDiasBúsqueda = 60; // Límite de seguridad para evitar loops infinitos

  while (diasConSlots.length < 7 && diasBuscados < maxDiasBúsqueda) {
    const slotsDelDia = agendaService.getAgendaForDay(currentDay);
    const sessionsRepo = new SessionRepository();
    const sesionesExistentes = sessionsRepo.findAll().filter(s => 
      normalizarFecha_(s.FechaSesion).getTime() === normalizarFecha_(currentDay).getTime()
    );
    
    const occupied = availabilityService._getOccupiedSlotsFromSessions(sesionesExistentes);
    
    const libres = slotsDelDia.filter(slot => {
      // Solo slots compatibles con seguimiento individual (2.2)
      if (slot.type !== '2.2' && slot.type !== 'SEGUIMIENTO') return false;
      return !availabilityService._isSlotOccupied(slot, occupied);
    }).map(s => ({
      hora: formatearHora_(s.startDateTime),
      label: formatearHora_(s.startDateTime)
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

/**
 * Obtiene una lista de reservas 2.1 provisionales para el formulario de edición/eliminación.
 * @returns {Array<Object>} Lista de objetos de reserva.
 */
function obtenerReservas21Formulario() {
  const sessionRepo = new SessionRepository();
  const reservas = sessionRepo.findAll().filter(s =>
    s.EstadoSesion === ESTADOS_SESION.RESERVADA_PROVISIONAL &&
    s.PacienteID === PACIENTE_ID_RESERVA_21_GENERICA
  );

  return reservas.map(r => ({
    sesionId: r.SesionID,
    label: `${formatearFecha_(r.FechaSesion)} ${formatearHora_(r.HoraInicio)} - ${r.NombrePaciente || 'Sin Nombre'}`
  })).sort((a, b) => a.label.localeCompare(b.label));
}

/**
 * Obtiene los detalles de una reserva 2.1 provisional específica.
 * @param {string} sesionId - ID de la sesión de reserva.
 * @returns {Object} Detalles de la reserva.
 */
function obtenerDetalleReserva21Formulario(sesionId) {
  const sessionRepo = new SessionRepository();
  const reserva = sessionRepo.findOneBy('SesionID', sesionId);

  if (!reserva || reserva.EstadoSesion !== ESTADOS_SESION.RESERVADA_PROVISIONAL) {
    throw new Error('Reserva no encontrada o no es una reserva provisional 2.1.');
  }

  const detalleReserva = {
    sesionId: String(reserva.SesionID || ''),
    fechaISO: formatearFechaISOInput_(reserva.FechaSesion),
    hora: formatearHora_(reserva.HoraInicio),
    nombreProvisional: reserva.NombrePaciente === NOMBRE_RESERVA_21_GENERICA ? '' : String(reserva.NombrePaciente || ''),
    nhcProvisional: reserva.NHC === NHC_RESERVA_21_GENERICA ? '' : String(reserva.NHC || '')
  };
  
  Logger.log(`[SERVER] obtenerDetalleReserva21Formulario returning POJO: ${JSON.stringify(detalleReserva)}`);
  return JSON.parse(JSON.stringify(detalleReserva));
}

/**
 * Elimina una reserva 2.1 provisional. La eliminación del evento de Calendar
 * se delega en el proceso de sincronización general.
 * @param {string} sesionId - ID de la sesión a eliminar.
 * @returns {Object} Mensaje de confirmación.
 */
function eliminarReserva21(sesionId) {
  const sessionRepo = new SessionRepository();
  const reserva = sessionRepo.findOneBy('SesionID', sesionId);
  
  if (!reserva || reserva.EstadoSesion !== ESTADOS_SESION.RESERVADA_PROVISIONAL) {
    throw new Error('Reserva no encontrada o no es una reserva provisional 2.1.');
  }

  // NO se elimina el evento de Google Calendar aquí.
  // El proceso de sincronización general detectará que esta sesión ya no existe
  // y se encargará de eliminar el evento huérfano del calendario.
  // Esto mejora la respuesta de la UI y centraliza la lógica de Calendar.

  sessionRepo.delete(reserva);
  SpreadsheetApp.flush();
  if (typeof __EXECUTION_CACHE__ !== 'undefined') __EXECUTION_CACHE__[SHEET_SESIONES] = null;
  eliminarCacheDashboard_();

  return { mensaje: 'Reserva eliminada correctamente.' };
}

/**
 * Guarda los cambios en una reserva 2.1 provisional.
 * @param {Object} formData - Datos del formulario.
 * @param {string} formData.sesionId - ID de la sesión a editar.
 * @param {string} formData.fechaISO - Nueva fecha en formato ISO.
 * @param {string} formData.hora - Nueva hora.
 * @param {string} [formData.nombreProvisional] - Nuevo nombre provisional.
 * @param {string} [formData.nhcProvisional] - Nuevo NHC provisional.
 * @returns {Object} Mensaje de confirmación.
 */
function guardarEdicionReserva21(formData) {
  const sessionRepo = new SessionRepository();
  const reserva = sessionRepo.findOneBy('SesionID', formData.sesionId);

  if (!reserva || reserva.EstadoSesion !== ESTADOS_SESION.RESERVADA_PROVISIONAL) {
    throw new Error('Reserva no encontrada o no es una reserva provisional 2.1.');
  }

  const nuevaFechaHoraInicio = normalizarFechaHora_(parseFechaISO_(formData.fechaISO), formData.hora);
  const duracionSlot = 60; // Duración estándar para 2.1

  // Si la fecha/hora ha cambiado, verificar disponibilidad
  if (compararFechasHoras_(reserva.FechaSesion, nuevaFechaHoraInicio) !== 0 || reserva.HoraInicio !== formData.hora) {
    const availabilityService = new AvailabilityService();
    const agendaService = new AgendaService();

    const agendaForDay = agendaService.getAgendaForDay(nuevaFechaHoraInicio);
    const targetSlot = agendaForDay.find(slot =>
      compararFechasHoras_(slot.startDateTime, nuevaFechaHoraInicio) === 0 &&
      (slot.type === '2.1' || slot.type === 'PRIMERA') &&
      slot.durationMinutes >= duracionSlot
    );

    if (!targetSlot) {
      throw new Error(`El nuevo slot ${formatearFecha_(nuevaFechaHoraInicio)} a las ${formatearHora_(nuevaFechaHoraInicio)} no es un slot de primera consulta (2.1) válido o no existe en la plantilla.`);
    }

    const sesionesExistentes = sessionRepo.findAll().filter(s =>
      s.SesionID !== reserva.SesionID && // Ignorar la propia sesión que estamos moviendo
      normalizarFecha_(s.FechaSesion).getTime() === normalizarFecha_(nuevaFechaHoraInicio).getTime()
    );
    const occupied = availabilityService._getOccupiedSlotsFromSessions(sesionesExistentes);

    if (availabilityService._isSlotOccupied(targetSlot, occupied)) {
      throw new Error(`El nuevo slot ${formatearFecha_(nuevaFechaHoraInicio)} a las ${formatearHora_(nuevaFechaHoraInicio)} ya está ocupado.`);
    }
  }

  // Actualizar campos de la reserva
  reserva.FechaSesion = normalizarFecha_(nuevaFechaHoraInicio);
  reserva.HoraInicio = formatearHora_(nuevaFechaHoraInicio);
  reserva.NombrePaciente = formData.nombreProvisional || NOMBRE_RESERVA_21_GENERICA;
  reserva.NHC = formData.nhcProvisional || NHC_RESERVA_21_GENERICA;
  reserva.Notas = `Reserva provisional de 2.1. Nombre: ${reserva.NombrePaciente || 'N/A'}, NHC: ${reserva.NHC || 'N/A'}`;
  reserva.ModificadaManual = true; // Se ha editado manualmente

  sessionRepo.save(reserva);
  SpreadsheetApp.flush();
  if (typeof __EXECUTION_CACHE__ !== 'undefined') __EXECUTION_CACHE__[SHEET_SESIONES] = null;
  eliminarCacheDashboard_();
  
  // No es necesario forzar la sincronización de Calendar aquí.
  // El evento se actualizará en la próxima sincronización programada o manual.

  return { mensaje: 'Reserva actualizada correctamente.' };
}

/**
 * Crea una nueva sesión de reserva provisional 2.1.
 * @param {Object} formData - Datos del formulario.
 * @param {string} formData.fechaISO - Fecha en formato ISO.
 * @param {string} formData.hora - Hora.
 * @param {string} [formData.nombreProvisional] - Nombre provisional.
 * @param {string} [formData.nhcProvisional] - NHC provisional.
 * @returns {Object} Mensaje de confirmación.
 */
function reservarPrimeraSesion(formData) {
  const sessionRepo = new SessionRepository();
  const availabilityService = new AvailabilityService();
  const agendaService = new AgendaService();

  const fechaHoraInicio = normalizarFechaHora_(parseFechaISO_(formData.fechaISO), formData.hora);
  const duracionSlot = 60; // Duración estándar para 2.1

  // Verificar disponibilidad
  const agendaForDay = agendaService.getAgendaForDay(fechaHoraInicio);
  const targetSlot = agendaForDay.find(slot =>
    compararFechasHoras_(slot.startDateTime, fechaHoraInicio) === 0 &&
    (slot.type === '2.1' || slot.type === 'PRIMERA') &&
    slot.durationMinutes >= duracionSlot
  );

  if (!targetSlot) {
    throw new Error(`El slot ${formatearFecha_(fechaHoraInicio)} a las ${formatearHora_(fechaHoraInicio)} no es un slot de primera consulta (2.1) válido o no existe en la plantilla.`);
  }

  const sesionesExistentes = sessionRepo.findAll().filter(s =>
    normalizarFecha_(s.FechaSesion).getTime() === normalizarFecha_(fechaHoraInicio).getTime()
  );
  const occupied = availabilityService._getOccupiedSlotsFromSessions(sesionesExistentes);

  if (availabilityService._isSlotOccupied(targetSlot, occupied)) {
    throw new Error(`El slot ${formatearFecha_(fechaHoraInicio)} a las ${formatearHora_(fechaHoraInicio)} ya está ocupado.`);
  }

  // Crear la sesión de reserva
  const sesionId = generarId_('SES');
  const nuevaReserva = {
    SesionID: sesionId,
    PacienteID: PACIENTE_ID_RESERVA_21_GENERICA,
    CicloID: '',
    AsignacionID: '',
    Modalidad: 'INDIVIDUAL', // Las 2.1 son siempre individuales
    NombrePaciente: formData.nombreProvisional || NOMBRE_RESERVA_21_GENERICA,
    NHC: formData.nhcProvisional || NHC_RESERVA_21_GENERICA,
    NumeroSesion: 0, // No es una sesión de ciclo
    FechaSesion: normalizarFecha_(fechaHoraInicio),
    HoraInicio: formatearHora_(fechaHoraInicio),
    Duracion: duracionSlot,
    EstadoSesion: ESTADOS_SESION.RESERVADA_PROVISIONAL,
    FechaOriginal: normalizarFecha_(fechaHoraInicio),
    Notas: `Reserva provisional de 2.1. Nombre: ${formData.nombreProvisional || 'N/A'}, NHC: ${formData.nhcProvisional || 'N/A'}`
  };

  sessionRepo.save(nuevaReserva);
  eliminarCacheDashboard_();

  return { mensaje: 'Reserva creada correctamente.' };
}

function guardarNuevoPacienteDesdeFormulario(formData) {
  const nombre = String(formData.nombre || '').trim();
  const nhc = String(formData.nhc || '').trim();
  const sexoGenero = String(formData.sexoGenero || '').trim();
  const motivoConsultaDiagnostico = String(formData.motivoConsultaDiagnostico || '').trim();
  const motivoConsultaOtros = String(formData.motivoConsultaOtros || '').trim();
  const modalidad = String(formData.modalidad || '').trim();
  const fechaISO = String(formData.fechaPrimeraConsulta || '').trim();

  if (!nombre) {
    throw new Error('El nombre es obligatorio.');
  }

  if (!nhc) {
    throw new Error('El NHC es obligatorio.');
  }

  if (motivoConsultaDiagnostico === 'Otros' && !motivoConsultaOtros) {
    throw new Error('Debes indicar el motivo cuando eliges "Otros".');
  }

  const modalidadesValidas = Object.values(MODALIDADES);
  if (!modalidadesValidas.includes(modalidad)) {
    throw new Error('La modalidad no es válida.');
  }

  if (!fechaISO) {
    throw new Error('La fecha de primera consulta es obligatoria.');
  }

  const fechaPrimeraConsulta = parseFechaISO_(fechaISO);
  if (!(fechaPrimeraConsulta instanceof Date)) {
    throw new Error('La fecha de primera consulta no es válida.');
  }

  return crearPacienteSegunModalidad_({
    nombre,
    nhc,
    sexoGenero,
    motivoConsultaDiagnostico,
    motivoConsultaOtros,
    modalidad,
    fechaPrimeraConsulta,
    slotSeguimiento: formData.slotSeguimiento // Objeto { fecha, hora }
  });
}

function parseFechaISO_(texto) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec((texto || '').trim());
  if (!m) return null;

  const year = Number(m[1]);
  const month = Number(m[2]) - 1;
  const day = Number(m[3]);

  const fecha = new Date(year, month, day);

  if (
    fecha.getFullYear() !== year ||
    fecha.getMonth() !== month ||
    fecha.getDate() !== day
  ) {
    return null;
  }

  return normalizarFecha_(fecha);
}

/***************
 * ESPERA -> CICLO (MANUAL)
 ***************/
function asignarPacienteEnEsperaACiclo() {
  const html = HtmlService
    .createHtmlOutputFromFile('AsignarEsperaCicloForm')
    .setWidth(620)
    .setHeight(620);

  SpreadsheetApp.getUi().showModalDialog(html, 'Asignar paciente en espera a ciclo');
}

function obtenerPacientesEnEspera_() {
  const repo = new PatientRepository();
  const out = repo.findAll()
    .filter(p => p.EstadoPaciente === ESTADOS_PACIENTE.ESPERA)
    .map(p => ({
      PacienteID: p.PacienteID,
      Nombre: p.Nombre,
      ModalidadSolicitada: p.ModalidadSolicitada,
      FechaPrimeraConsulta: p.FechaPrimeraConsulta, // Sigue siendo Date interno, pero filtrado por callers
      MotivoEspera: p.MotivoEspera || '',
      fila: p._row 
    }));

  out.sort((a, b) => compararFechas_(a.FechaPrimeraConsulta, b.FechaPrimeraConsulta));
  return out;
}

function obtenerCiclosDisponiblesParaPacienteEnEspera_(paciente) {
  const cicloRepo = new CicloRepository();
  const todosLosCiclos = cicloRepo.findAll(); // Usa el repositorio y su caché

  const ciclosFiltrados = todosLosCiclos.filter(c => {
    const fechaInicio = c.FechaInicioCiclo;
    const capacidad = Number(c.CapacidadMaxima || 0);
    const ocupadas = Number(c.PlazasOcupadas || 0);
    const libresReales = capacidad - ocupadas;

    return (
      c.Modalidad === paciente.ModalidadSolicitada &&
      c.EstadoCiclo === ESTADOS_CICLO.PLANIFICADO &&
      (fechaInicio instanceof Date) &&
      // CAMBIO: Permitir ciclos que empiecen el mismo día de la consulta
      (normalizarFecha_(fechaInicio).getTime() >= normalizarFecha_(paciente.FechaPrimeraConsulta).getTime()) &&
      libresReales > 0 &&
      !(c.BloqueoInscripciones === true || c.BloqueoInscripciones === 'TRUE')
    );
  });

  ciclosFiltrados.sort((a, b) => compararFechas_(a.FechaInicioCiclo, b.FechaInicioCiclo));

  // Mapear a un formato compatible con la UI, asegurando que las fechas sean strings
  return ciclosFiltrados.map(c => ({
    ...c,
    FechaInicioCiclo: formatearFecha_(c.FechaInicioCiclo) // Convertir Date a string
  }));
}

function actualizarPacienteEsperaAAssignado_(pacienteId, ciclo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('La hoja PACIENTES no tiene datos.');
  }

  const idx = indexByHeader_(data[0]);

  const columnasNecesarias = [
    'PacienteID',
    'EstadoPaciente',
    'MotivoEspera',
    'CicloObjetivoID',
    'CicloActivoID',
    'FechaPrimeraSesionReal',
    'SesionesPlanificadas',
    'SesionesPendientes',
    'ProximaSesion'
  ];

  columnasNecesarias.forEach(col => {
    if (idx[col] === undefined) {
      throw new Error('Falta la columna "' + col + '" en ' + SHEET_PACIENTES + '.');
    }
  });

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.PacienteID]) === String(pacienteId)) {
      // Cargamos la fila actual, modificamos el array en memoria y guardamos de una vez
      const filaActual = data[i];
      filaActual[idx.EstadoPaciente] = ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO;
      filaActual[idx.MotivoEspera] = '';
      filaActual[idx.CicloObjetivoID] = ciclo.CicloID;
      filaActual[idx.CicloActivoID] = '';
      filaActual[idx.FechaPrimeraSesionReal] = ciclo.FechaInicioCiclo;
      filaActual[idx.SesionesPlanificadas] = ciclo.SesionesPorCiclo;
      filaActual[idx.SesionesPendientes] = ciclo.SesionesPorCiclo;
      filaActual[idx.ProximaSesion] = ciclo.FechaInicioCiclo;

      sheet.getRange(i + 1, 1, 1, filaActual.length).setValues([filaActual]);
      return;
    }
  }

  throw new Error('No existe el paciente con ID ' + pacienteId + '.');
}

function asegurarSesionesPacienteGrupoSiFaltan_(pacienteId, cicloId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_SESIONES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length >= 2) {
    const idx = indexByHeader_(data[0]);

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idx.PacienteID] || '') === String(pacienteId)) {
        return; // ya tiene sesiones, no generamos
      }
    }
  }

  generarSesionesPacienteGrupo_(pacienteId, cicloId);
}

function obtenerPacientesEnEsperaFormulario() {
  const pacientes = obtenerPacientesEnEspera_();

  return pacientes.map(p => ({
    pacienteId: p.PacienteID,
    nombre: p.Nombre || '',
    modalidad: p.ModalidadSolicitada || '',
    fechaPrimeraConsulta: formatearFecha_(p.FechaPrimeraConsulta),
    motivoEspera: p.MotivoEspera || ''
  }));
}

function obtenerCiclosDisponiblesPacienteFormulario(pacienteId) {
  const pacientes = obtenerPacientesEnEspera_();
  const paciente = pacientes.find(p => String(p.PacienteID) === String(pacienteId));

  if (!paciente) {
    throw new Error('Paciente en espera no encontrado.');
  }

  const ciclos = obtenerCiclosDisponiblesParaPacienteEnEspera_(paciente);

  return ciclos.map(c => ({
    cicloId: c.CicloID,
    modalidad: c.Modalidad,
    numeroCiclo: c.NumeroCiclo,
    estadoCiclo: c.EstadoCiclo,
    fechaInicio: formatearFecha_(c.FechaInicioCiclo),
    plazasLibres: c.PlazasLibres
  }));
}

function confirmarAsignacionPacienteEnEsperaFormulario(formData) {
  const pacienteId = String(formData.pacienteId || '').trim();
  const cicloId = String(formData.cicloId || '').trim();

  if (!pacienteId) {
    throw new Error('Debes seleccionar un paciente.');
  }

  if (!cicloId) {
    throw new Error('Debes seleccionar un ciclo.');
  }

  const pacientes = obtenerPacientesEnEspera_();
  const paciente = pacientes.find(p => String(p.PacienteID) === String(pacienteId));

  if (!paciente) {
    throw new Error('Paciente en espera no encontrado.');
  }

  const ciclosDisponibles = obtenerCiclosDisponiblesParaPacienteEnEspera_(paciente);
  const ciclo = ciclosDisponibles.find(c => String(c.CicloID) === String(cicloId));

  if (!ciclo) {
    throw new Error('El ciclo ya no está disponible para este paciente.');
  }

  const reservaOk = actualizarPlazasCiclo_(ciclo.CicloID, +1, true);

  if (!reservaOk) {
    throw new Error('Ese ciclo ya no tiene plazas disponibles.');
  }

  actualizarPacienteEsperaAAssignado_(paciente.PacienteID, ciclo);
  
  // Limpiar caché para que el cambio de estado se refleje en los siguientes pasos
  __EXECUTION_CACHE__[SHEET_PACIENTES] = null;

  crearAsignacionCiclo_({
    pacienteId: paciente.PacienteID,
    cicloId: ciclo.CicloID,
    modalidad: ciclo.Modalidad,
    estadoAsignacion: ESTADOS_ASIGNACION.RESERVADO,
    observaciones: 'Asignación manual desde espera'
  });

  asegurarSesionesPacienteGrupoSiFaltan_(paciente.PacienteID, ciclo.CicloID);
  new MaintenanceService().recalculateCycleOccupancy();
  new StateService().runAutomaticTransitions();

  try {
    refrescarDashboard();
  } catch (e) {
    // no rompemos por dashboard
  }

  return {
    mensaje:
      'Paciente asignado correctamente.\n\n' +
      'Paciente: ' + paciente.Nombre + '\n' +
      'Modalidad: ' + paciente.ModalidadSolicitada + '\n' +
      'Ciclo: ' + ciclo.NumeroCiclo + '\n' +
      'Inicio: ' + formatearFecha_(ciclo.FechaInicioCiclo)
  };
}

/***************
 * EDICIÓN DE PACIENTE
 ***************/
function editarPaciente() {
  const html = HtmlService
    .createHtmlOutputFromFile('EditarPacienteForm')
    .setWidth(700)
    .setHeight(720);

  SpreadsheetApp.getUi().showModalDialog(html, 'Editar paciente');
}

function obtenerPacientesFormularioEdicion() {
  const repo = new PatientRepository();
  return repo.findAll().map(p => ({
    pacienteId: p.PacienteID,
    label: `${p.Nombre || 'SIN_NOMBRE'} | ${p.ModalidadSolicitada || ''} | ${p.EstadoPaciente || ''}`
  })).sort((a, b) => a.label.localeCompare(b.label));
}

function obtenerDetallePacienteParaEdicion(pacienteId) {
  const patientRepo = new PatientRepository();
  const paciente = patientRepo.findById(pacienteId); // Usa el repositorio

  if (!paciente) throw new Error('Paciente no encontrado.');

  const estado = paciente.EstadoPaciente;
  const puedeEditarModalidad = estado === ESTADOS_PACIENTE.ESPERA;
  const puedeEditarFechaConsulta = estado === ESTADOS_PACIENTE.ESPERA;

  return {
    pacienteId: paciente.PacienteID,
    nombre: paciente.Nombre || '',
    modalidad: paciente.ModalidadSolicitada || '',
    fechaPrimeraConsulta: formatearFechaISOInput_(paciente.FechaPrimeraConsulta),
    estadoPaciente: estado || '',
    motivoEspera: paciente.MotivoEspera || '',
    cicloObjetivoId: paciente.CicloObjetivoID || '',
    cicloActivoId: paciente.CicloActivoID || '',
    fechaPrimeraSesionReal: formatearFecha_(paciente.FechaPrimeraSesionReal),
    sesionesPlanificadas: Number(paciente.SesionesPlanificadas || 0),
    sesionesCompletadas: Number(paciente.SesionesCompletadas || 0),
    sesionesPendientes: Number(paciente.SesionesPendientes || 0),
    proximaSesion: formatearFecha_(paciente.ProximaSesion),
    fechaCierre: formatearFecha_(paciente.FechaCierre),
    observaciones: paciente.Observaciones || '',
    restricciones: {
      puedeEditarModalidad: puedeEditarModalidad,
      puedeEditarFechaConsulta: puedeEditarFechaConsulta,
      puedeEditarNombre: true,
      puedeEditarObservaciones: true
    }
  };
}

function guardarEdicionPaciente(formData) {
  const patientRepo = new PatientRepository();
  const paciente = patientRepo.findById(formData.pacienteId);
  if (!paciente) throw new Error('Paciente no encontrado.');

  const nombre = String(formData.nombre || '').trim();
  const modalidadNueva = String(formData.modalidad || '').trim();
  const fechaConsultaISO = String(formData.fechaPrimeraConsulta || '').trim();
  const observaciones = String(formData.observaciones || '').trim();

  if (!nombre) throw new Error('El nombre es obligatorio.');

  const estado = paciente.EstadoPaciente;
  const modalidadActual = paciente.ModalidadSolicitada;

  paciente.Nombre = nombre;
  paciente.Observaciones = observaciones;

  let mensajeAdicional = '';

  if (estado === ESTADOS_PACIENTE.ESPERA) {
    if (!Object.values(MODALIDADES).includes(modalidadNueva)) {
      throw new Error('La modalidad no es válida.');
    }
    const fechaConsulta = parseFechaISO_(fechaConsultaISO);
    if (!(fechaConsulta instanceof Date)) {
      throw new Error('La fecha de primera consulta no es válida.');
    }

    paciente.ModalidadSolicitada = modalidadNueva;
    paciente.FechaPrimeraConsulta = fechaConsulta;
    mensajeAdicional = `\nModalidad anterior: ${modalidadActual}\nModalidad nueva: ${modalidadNueva}`;
  } else {
    if (modalidadNueva && modalidadNueva !== modalidadActual) {
      throw new Error('Solo se puede cambiar la modalidad si el paciente está en ESPERA.');
    }
    if (fechaConsultaISO) {
      const fechaOriginal = paciente.FechaPrimeraConsulta;
      const fechaNueva = parseFechaISO_(fechaConsultaISO);
      if (!(fechaNueva instanceof Date)) {
        throw new Error('La fecha de primera consulta no es válida.');
      }
      const originalTime = fechaOriginal instanceof Date ? normalizarFecha_(fechaOriginal).getTime() : null;
      const nuevaTime = normalizarFecha_(fechaNueva).getTime();

      if (originalTime !== null && originalTime !== nuevaTime) {
        throw new Error('Solo se puede cambiar la fecha de primera consulta si el paciente está en ESPERA.');
      }
    }
    mensajeAdicional = '\nSolo se han permitido los campos editables para ese estado.';
  }

  patientRepo.save(paciente); // Guardado por bloques
  eliminarCacheDashboard_(); // Invalida la caché del dashboard

  return {
    mensaje:
      `Paciente actualizado correctamente.\n\n` +
      `Nombre: ${nombre}\n` +
      `Estado: ${estado}` +
      mensajeAdicional
  };
}

function formatearFechaISOInput_(fecha) {
  if (!fecha) return '';
  const d = (fecha instanceof Date) ? fecha : new Date(fecha);
  if (isNaN(d.getTime())) return '';
  
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function obtenerCatalogosFormularioEdicionPaciente() {
  return {
    modalidades: obtenerValoresCatalogo_('MODALIDADES'),
    estadosPaciente: obtenerValoresCatalogo_('ESTADOS_PACIENTE')
  };
}

/***************
 * REASIGNACIÓN AVANZADA
 ***************/
function reasignacionAvanzadaPaciente() {
  const html = HtmlService
    .createHtmlOutputFromFile('ReasignarPacienteForm')
    .setWidth(720)
    .setHeight(620);

  SpreadsheetApp.getUi().showModalDialog(html, 'Reasignación avanzada de paciente');
}

function obtenerPacientesReasignablesFormulario() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const idx = indexByHeader_(data[0]);
  const out = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const estado = row[idx.EstadoPaciente];

    if (
      estado !== ESTADOS_PACIENTE.ESPERA &&
      estado !== ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO
    ) {
      continue;
    }

    out.push({
      pacienteId: row[idx.PacienteID],
      label:
        (row[idx.Nombre] || 'SIN_NOMBRE') +
        ' | ' + (row[idx.ModalidadSolicitada] || '') +
        ' | ' + (estado || '')
    });
  }

  out.sort((a, b) => String(a.label).localeCompare(String(b.label)));
  return out;
}

function obtenerDetallePacienteReasignacionFormulario(pacienteId) {
  const paciente = obtenerPacienteCompletoPorId_(pacienteId);
  if (!paciente) {
    throw new Error('Paciente no encontrado.');
  }

  if (
    paciente.EstadoPaciente !== ESTADOS_PACIENTE.ESPERA &&
    paciente.EstadoPaciente !== ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO
  ) {
    throw new Error('Solo se pueden reasignar pacientes en ESPERA o ACTIVO_PENDIENTE_INICIO.');
  }

  return {
    pacienteId: paciente.PacienteID,
    nombre: paciente.Nombre || '',
    modalidad: paciente.ModalidadSolicitada || '',
    estadoPaciente: paciente.EstadoPaciente || '',
    fechaPrimeraConsulta: formatearFecha_(paciente.FechaPrimeraConsulta),
    motivoEspera: paciente.MotivoEspera || '',
    cicloObjetivoId: paciente.CicloObjetivoID || '',
    cicloActivoId: paciente.CicloActivoID || '',
    proximaSesion: formatearFecha_(paciente.ProximaSesion),
    observaciones: paciente.Observaciones || ''
  };
}

function obtenerCiclosReasignacionFormulario(pacienteId) {
  const paciente = obtenerPacienteCompletoPorId_(pacienteId);
  if (!paciente) {
    throw new Error('Paciente no encontrado.');
  }

  if (
    paciente.EstadoPaciente !== ESTADOS_PACIENTE.ESPERA &&
    paciente.EstadoPaciente !== ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO
  ) {
    throw new Error('Solo se pueden reasignar pacientes en ESPERA o ACTIVO_PENDIENTE_INICIO.');
  }

  const ciclos = obtenerCiclosDisponiblesParaPacienteEnEspera_({
    PacienteID: paciente.PacienteID,
    Nombre: paciente.Nombre,
    ModalidadSolicitada: paciente.ModalidadSolicitada,
    FechaPrimeraConsulta: paciente.FechaPrimeraConsulta,
    MotivoEspera: paciente.MotivoEspera || ''
  });

  const cicloActualRef = String(paciente.CicloObjetivoID || paciente.CicloActivoID || '');

  return ciclos
    .filter(c => String(c.CicloID) !== cicloActualRef)
    .map(c => ({
      cicloId: c.CicloID,
      modalidad: c.Modalidad,
      numeroCiclo: c.NumeroCiclo,
      estadoCiclo: c.EstadoCiclo,
      fechaInicio: formatearFecha_(c.FechaInicioCiclo),
      plazasLibres: c.PlazasLibres
    }));
}

function confirmarReasignacionPacienteFormulario(formData) {
  const pacienteId = String(formData.pacienteId || '').trim();
  const nuevoCicloId = String(formData.cicloId || '').trim();

  if (!pacienteId) {
    throw new Error('Debes seleccionar un paciente.');
  }

  if (!nuevoCicloId) {
    throw new Error('Debes seleccionar un ciclo destino.');
  }

  const paciente = obtenerPacienteCompletoPorId_(pacienteId);
  if (!paciente) {
    throw new Error('Paciente no encontrado.');
  }

  if (
    paciente.EstadoPaciente !== ESTADOS_PACIENTE.ESPERA &&
    paciente.EstadoPaciente !== ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO
  ) {
    throw new Error('Solo se pueden reasignar pacientes en ESPERA o ACTIVO_PENDIENTE_INICIO.');
  }

  const ciclosDisponibles = obtenerCiclosDisponiblesParaPacienteEnEspera_({
    PacienteID: paciente.PacienteID,
    Nombre: paciente.Nombre,
    ModalidadSolicitada: paciente.ModalidadSolicitada,
    FechaPrimeraConsulta: paciente.FechaPrimeraConsulta,
    MotivoEspera: paciente.MotivoEspera || ''
  });

  const nuevoCiclo = ciclosDisponibles.find(c => String(c.CicloID) === String(nuevoCicloId));
  if (!nuevoCiclo) {
    throw new Error('El ciclo destino ya no está disponible.');
  }

  const cicloAnteriorId = String(paciente.CicloObjetivoID || paciente.CicloActivoID || '');
  const tieneAsignacionAnterior = !!cicloAnteriorId;

  const reservaOk = actualizarPlazasCiclo_(nuevoCiclo.CicloID, +1, true);
  if (!reservaOk) {
    throw new Error('El ciclo destino ya no tiene plazas disponibles.');
  }

  try {
    if (tieneAsignacionAnterior) {
      cancelarAsignacionVigentePaciente_(paciente.PacienteID, cicloAnteriorId);
      actualizarPlazasCiclo_(cicloAnteriorId, -1, false);
      cancelarSesionesPendientesPaciente_(paciente.PacienteID);
    }

    actualizarPacienteTrasReasignacion_(paciente.PacienteID, nuevoCiclo);

    crearAsignacionCiclo_({
      pacienteId: paciente.PacienteID,
      cicloId: nuevoCiclo.CicloID,
      modalidad: nuevoCiclo.Modalidad,
      estadoAsignacion: ESTADOS_ASIGNACION.RESERVADO,
      observaciones: tieneAsignacionAnterior
        ? 'Reasignación avanzada a nuevo ciclo'
        : 'Asignación desde espera'
    });

    generarSesionesPacienteGrupo_(paciente.PacienteID, nuevoCiclo.CicloID);

    new MaintenanceService().recalculateCycleOccupancy();
    new StateService().runAutomaticTransitions();
    
    try {
      refrescarDashboard();
    } catch (e) {
      // no rompemos por dashboard
    }

    return {
      mensaje:
        'Reasignación completada correctamente.\n\n' +
        'Paciente: ' + paciente.Nombre + '\n' +
        'Modalidad: ' + paciente.ModalidadSolicitada + '\n' +
        'Nuevo ciclo: ' + nuevoCiclo.NumeroCiclo + '\n' +
        'Inicio: ' + formatearFecha_(nuevoCiclo.FechaInicioCiclo)
    };

  } catch (error) {
    // ROLLBACK: Revertimos los cambios realizados para evitar el "limbo administrativo"
    try {
      actualizarPlazasCiclo_(nuevoCiclo.CicloID, -1, false);
      if (tieneAsignacionAnterior) {
        // Devolvemos la plaza al ciclo viejo y restauramos estados
        actualizarPlazasCiclo_(cicloAnteriorId, +1, false);
        restaurarAsignacionVigente_(paciente.PacienteID, cicloAnteriorId);
        restaurarSesionesCanceladas_(paciente.PacienteID);
      }
    } catch (e) {
      console.error("Error crítico en el rollback: " + e.message);
    }
    throw error;
  }
}

function obtenerPacienteCompletoPorId_(pacienteId) {
  const repo = new PatientRepository();
  return repo.findById(pacienteId);
}

function cancelarAsignacionVigentePaciente_(pacienteId, cicloId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ASIGNACIONES_CICLO);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_ASIGNACIONES_CICLO + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    const rowPacienteId = String(data[i][idx.PacienteID] || '');
    const rowCicloId = String(data[i][idx.CicloID] || '');
    const estado = data[i][idx.EstadoAsignacion];

    if (
      rowPacienteId === String(pacienteId) &&
      rowCicloId === String(cicloId) &&
      (estado === ESTADOS_ASIGNACION.RESERVADO || estado === ESTADOS_ASIGNACION.ACTIVO)
    ) {
      sheet.getRange(i + 1, idx.EstadoAsignacion + 1).setValue(ESTADOS_ASIGNACION.CANCELADO);
    }
  }
}

function cancelarSesionesPendientesPaciente_(pacienteId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_SESIONES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    const rowPacienteId = String(data[i][idx.PacienteID] || '');
    const estadoSesion = data[i][idx.EstadoSesion];

    if (rowPacienteId !== String(pacienteId)) continue;

    if (
      estadoSesion === ESTADOS_SESION.PENDIENTE ||
      estadoSesion === ESTADOS_SESION.REPROGRAMADA
    ) {
      sheet.getRange(i + 1, idx.EstadoSesion + 1).setValue(ESTADOS_SESION.CANCELADA);
    }
  }
}

function actualizarPacienteTrasReasignacion_(pacienteId, ciclo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('La hoja PACIENTES no tiene datos.');
  }

  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.PacienteID]) !== String(pacienteId)) continue;

    // PATRÓN ATÓMICO: Modificamos toda la fila en memoria y escribimos una sola vez
    const fila = data[i];
    fila[idx.EstadoPaciente] = ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO;
    fila[idx.MotivoEspera] = '';
    fila[idx.CicloObjetivoID] = ciclo.CicloID;
    fila[idx.CicloActivoID] = '';
    fila[idx.FechaPrimeraSesionReal] = ciclo.FechaInicioCiclo;
    fila[idx.SesionesPlanificadas] = ciclo.SesionesPorCiclo;
    fila[idx.SesionesCompletadas] = 0;
    fila[idx.SesionesPendientes] = ciclo.SesionesPorCiclo;
    fila[idx.ProximaSesion] = ciclo.FechaInicioCiclo;
    fila[idx.FechaCierre] = '';
    
    sheet.getRange(i + 1, 1, 1, fila.length).setValues([fila]);
    return;
  }

  throw new Error('Paciente no encontrado para actualizar.');
}

/***************
 * ELIMINAR PACIENTE POR ERROR
 ***************/
function eliminarPacientePorError() {
  const html = HtmlService
    .createHtmlOutputFromFile('EliminarPacienteForm')
    .setWidth(720)
    .setHeight(620);

  SpreadsheetApp.getUi().showModalDialog(html, 'Eliminar paciente por error');
}

function obtenerPacientesEliminablesFormulario() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const idx = indexByHeader_(data[0]);
  const out = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const estado = row[idx.EstadoPaciente];
    const pacienteId = row[idx.PacienteID];

    if (!pacienteId) continue;

    out.push({
      pacienteId: pacienteId,
      label: (row[idx.Nombre] || 'SIN_NOMBRE') + ' | ' + (row[idx.ModalidadSolicitada] || '') + ' | ' + (estado || '')
    });
  }

  out.sort((a, b) => String(a.label).localeCompare(String(b.label)));
  return out;
}

function obtenerImpactoEliminacionPacienteFormulario(pacienteId) {
  const paciente = obtenerPacienteCompletoPorId_(pacienteId);
  if (!paciente) {
    throw new Error('Paciente no encontrado.');
  }

  const asignaciones = contarAsignacionesPaciente_(paciente.PacienteID);
  const sesiones = contarSesionesPaciente_(paciente.PacienteID);
  const cicloId = String(paciente.CicloObjetivoID || paciente.CicloActivoID || '');

  // Ahora liberamos plaza si el paciente está activo o pendiente, para limpiar duplicados correctamente
  const liberaraPlaza = (paciente.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO || 
                         paciente.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO) && !!cicloId;

  return {
    pacienteId: paciente.PacienteID,
    nombre: paciente.Nombre || '',
    modalidad: paciente.ModalidadSolicitada || '',
    estado: paciente.EstadoPaciente || '',
    fechaPrimeraConsulta: formatearFecha_(paciente.FechaPrimeraConsulta),
    motivoEspera: paciente.MotivoEspera || '',
    cicloId: cicloId,
    asignaciones: asignaciones,
    sesiones: sesiones,
    liberaraPlaza: liberaraPlaza
  };
}

function confirmarEliminacionPacienteFormulario(formData) {
  const pacienteId = String(formData.pacienteId || '').trim();

  if (!pacienteId) {
    throw new Error('Debes seleccionar un paciente.');
  }

  const paciente = obtenerPacienteCompletoPorId_(pacienteId);
  if (!paciente) {
    throw new Error('Paciente no encontrado.');
  }

  const cicloId = String(paciente.CicloObjetivoID || paciente.CicloActivoID || '');
  
  // Lógica de liberación de plaza ampliada a ACTIVO para permitir limpieza de duplicados
  const liberarPlaza = (paciente.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO || 
                        paciente.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO) && !!cicloId;

  if (liberarPlaza) {
    actualizarPlazasCiclo_(cicloId, -1, false);
  }

  borrarAsignacionesPaciente_(paciente.PacienteID);
  borrarSesionesPaciente_(paciente.PacienteID);
  borrarPacientePorId_(paciente.PacienteID);

  new MaintenanceService().recalculateCycleOccupancy();
  new StateService().runAutomaticTransitions();

  try {
    refrescarDashboard();
  } catch (e) {
    eliminarCacheDashboard_();
    // no rompemos por dashboard
  }

  return {
    mensaje:
      'Paciente eliminado correctamente.\n\n' +
      'Paciente: ' + (paciente.Nombre || '') + '\n' +
      'Modalidad: ' + (paciente.ModalidadSolicitada || '') + '\n' +
      'Estado anterior: ' + (paciente.EstadoPaciente || '')
  };
}

function contarAsignacionesPaciente_(pacienteId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ASIGNACIONES_CICLO);
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;

  const idx = indexByHeader_(data[0]);
  let total = 0;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.PacienteID] || '') === String(pacienteId)) {
      total++;
    }
  }

  return total;
}

function contarSesionesPaciente_(pacienteId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;

  const idx = indexByHeader_(data[0]);
  let total = 0;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.PacienteID] || '') === String(pacienteId)) {
      total++;
    }
  }

  return total;
}

function borrarAsignacionesPaciente_(pacienteId) {
  const repo = new AsignacionRepository();
  const sheet = repo.getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const idx = indexByHeader_(data[0]);
  const pidStr = String(pacienteId);
  
  // PATRÓN B: Escritura por bloques (Filtramos en memoria y volcamos)
  const filtrados = data.filter((row, i) => i === 0 || String(row[idx.PacienteID]) !== pidStr);

  if (filtrados.length !== data.length) {
    sheet.clearContents();
    sheet.getRange(1, 1, filtrados.length, data[0].length).setValues(filtrados);
    __EXECUTION_CACHE__[SHEET_ASIGNACIONES_CICLO] = null;
  }
}

function borrarSesionesPaciente_(pacienteId) {
  const repo = new SessionRepository();
  const sheet = repo.getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const idx = indexByHeader_(data[0]);
  const pidStr = String(pacienteId);

  const filtrados = data.filter((row, i) => i === 0 || String(row[idx.PacienteID]) !== pidStr);
  const soloDatos = filtrados.slice(1);

  if (filtrados.length !== data.length) {
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    if (soloDatos.length > 0) {
      sheet.getRange(2, 1, soloDatos.length, data[0].length).setValues(soloDatos);
    }
    __EXECUTION_CACHE__[SHEET_SESIONES] = null;
  }
}

function borrarPacientePorId_(pacienteId) {
  const repo = new PatientRepository();
  const sheet = repo.getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const idx = indexByHeader_(data[0]);
  const pidStr = String(pacienteId);

  const filtrados = data.filter((row, i) => i === 0 || String(row[idx.PacienteID]) !== pidStr);

  if (filtrados.length !== data.length) {
    sheet.clearContents();
    sheet.getRange(1, 1, filtrados.length, data[0].length).setValues(filtrados);
    __EXECUTION_CACHE__[SHEET_PACIENTES] = null;
    eliminarCacheDashboard_();
  } else {
    throw new Error('Paciente no encontrado para borrado.');
  }
}

/***********************
 * GESTIÓN UNIFICADA ESPERA / REASIGNACIÓN
 ***********************/

function gestionarEsperaYCicloPaciente() {
  const html = HtmlService
    .createHtmlOutputFromFile('GestionEsperaCicloForm')
    .setWidth(500)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'Gestión espera / ciclo');
}


/***************
 * PACIENTES
 ***************/
function obtenerPacientesGestionEsperaCicloFormulario() {
  const enEspera = obtenerPacientesEnEsperaFormulario();
  const reasignables = obtenerPacientesReasignablesFormulario();

  const map = {};

  [...enEspera, ...reasignables].forEach(p => {
    map[p.pacienteId] = p;
  });

  return Object.values(map);
}


/***************
 * DETALLE PACIENTE
 ***************/
function obtenerDetallePacienteGestionEsperaCicloFormulario(pacienteId) {
  const detalle = obtenerDetallePacienteReasignacionFormulario(pacienteId);

  let tipoAccion = 'ASIGNAR';
  let labelAccion = 'Asignación a nuevo ciclo (desde Espera)';

  if (detalle.cicloActivoId || detalle.cicloObjetivoId) {
    tipoAccion = 'REASIGNAR';
    labelAccion = 'Cambio de grupo / Reasignación';
  }

  return {
    ...detalle,
    tipoAccion,
    labelAccion
  };
}


/***************
 * CICLOS DISPONIBLES
 ***************/
function obtenerCiclosGestionEsperaCicloFormulario(pacienteId, tipoAccion) {
  const accion = tipoAccion || obtenerDetallePacienteGestionEsperaCicloFormulario(pacienteId).tipoAccion;

  if (accion === 'REASIGNAR') {
    return obtenerCiclosReasignacionFormulario(pacienteId);
  }

  return obtenerCiclosDisponiblesPacienteFormulario(pacienteId);
}


/***************
 * CONFIRMAR
 ***************/
function confirmarGestionEsperaCicloFormulario(formData) {
  const { tipoAccion } = formData;

  if (tipoAccion === 'REASIGNAR') {
    return confirmarReasignacionPacienteFormulario(formData);
  }

  return confirmarAsignacionPacienteEnEsperaFormulario(formData);
}

function editarPacienteDesdeId_(pacienteId) {
  if (!pacienteId) {
    throw new Error('No se indicó paciente para editar.');
  }

  const template = HtmlService.createTemplateFromFile('EditarPacienteForm');
  template.pacientePreseleccionadoId = String(pacienteId);

  const html = template
    .evaluate()
    .setWidth(760)
    .setHeight(720);

  SpreadsheetApp.getUi().showModalDialog(html, 'Editar paciente');
}

/**
 * Restauradores de emergencia para Rollback (H-02)
 */
function restaurarAsignacionVigente_(pacienteId, cicloId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ASIGNACIONES_CICLO);
  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.PacienteID]) === String(pacienteId) && 
        String(data[i][idx.CicloID]) === String(cicloId) &&
        data[i][idx.EstadoAsignacion] === ESTADOS_ASIGNACION.CANCELADO) {
      
      // Lo devolvemos a RESERVADO (asumiendo que venía de reasignación desde activo_pendiente)
      sheet.getRange(i + 1, idx.EstadoAsignacion + 1).setValue(ESTADOS_ASIGNACION.RESERVADO);
      return;
    }
  }
}

function restaurarSesionesCanceladas_(pacienteId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);
  
  let restauradas = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.PacienteID]) === String(pacienteId) && 
        data[i][idx.EstadoSesion] === ESTADOS_SESION.CANCELADA) {
      
      sheet.getRange(i + 1, idx.EstadoSesion + 1).setValue(ESTADOS_SESION.PENDIENTE);
      restauradas++;
    }
  }
  if (restauradas > 0) {
    if (typeof __EXECUTION_CACHE__ !== 'undefined') __EXECUTION_CACHE__[SHEET_SESIONES] = null;
  }
}

/**
 * Reconcilia Google Calendar eliminando eventos que ya no tienen una fila correspondiente
 * en la hoja SESIONES. Útil para limpiar sesiones borradas manualmente o por error.
 * 
 * @param {number} diasAtras - Días hacia atrás para revisar (default 30).
 * @param {number} diasVista - Días hacia adelante para revisar (default 180).
 * @returns {number} Cantidad de eventos huérfanos eliminados.
 */
function limpiarEventosHuerfanosCalendar(diasAtras = 30, diasVista = 180) {
  const sessionRepo = new SessionRepository();
  const sesionesEnHoja = sessionRepo.findAll();
  
  // 1. Indexar todos los CalendarEventId que SÍ existen en Sheets para búsqueda O(1)
  const idsValidos = new Set();
  sesionesEnHoja.forEach(s => {
    if (s.CalendarEventId) idsValidos.add(String(s.CalendarEventId));
  });

  const calendars = CalendarApp.getCalendarsByName(GOOGLE_CALENDAR_NAME);
  if (calendars.length === 0) return 0;
  const calendar = calendars[0];

  // 2. Definir rango de búsqueda para no saturar la API
  const hoy = new Date();
  const inicio = sumarDiasNaturales_(hoy, -diasAtras);
  const fin = sumarDiasNaturales_(hoy, diasVista);
  
  const eventos = calendar.getEvents(inicio, fin);
  let borrados = 0;

  // 3. Comparar y eliminar
  eventos.forEach(event => {
    const eventId = event.getId();
    
    // Si el evento está en nuestro calendario pero su ID no consta en nuestra base de datos...
    if (!idsValidos.has(eventId)) {
       // SEGURIDAD CRÍTICA: Solo borramos si el evento tiene la "firma" de nuestra aplicación.
      // Evita que el sistema borre reuniones personales añadidas a mano en Calendar.
      const desc = event.getDescription() || '';
      const titulo = event.getTitle() || '';
      const esDelSistema = desc.includes('PACIENTE:') || titulo.includes('✅') || titulo.includes('🗓️') || titulo.includes('🔁') || titulo.includes('PRIMERA 2.1');
      
      if (esDelSistema && !esEventoDiaBloqueado_(event)) {
        event.deleteEvent();
        borrados++;
      }
    }
  });

  console.log(`Limpieza de calendario completada. Eventos huérfanos eliminados: ${borrados}`);
  return borrados;
}