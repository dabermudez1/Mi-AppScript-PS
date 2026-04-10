/***********************
 * BLOQUE 2
 * GESTIÓN DE CICLOS
 ***********************/
function crearCicloGrupo() {
  const html = HtmlService
    .createHtmlOutputFromFile('CrearCicloGrupoForm')
    .setWidth(420)
    .setHeight(320);

  SpreadsheetApp.getUi().showModalDialog(html, 'Crear ciclo de grupo');
}

function obtenerConfigModalidad_(modalidad) {
  const config = new ConfigRepository().findByModalidad(modalidad);
  if (!config) throw new Error('No existe configuración para la modalidad ' + modalidad + '.');
  return config;
}

function validarConfigGrupo_(modalidad, config) {
  if (!config.Activa) {
    throw new Error('La modalidad está inactiva: ' + modalidad);
  }

  if (config.TipoModalidad !== TIPOS_MODALIDAD.GRUPO) {
    throw new Error('La modalidad no es de tipo grupo: ' + modalidad);
  }

  if (!config.DiaSemana) {
    throw new Error('Falta DiaSemana en CONFIG_MODALIDADES para ' + modalidad + '.');
  }

  if (!config.FrecuenciaDias || config.FrecuenciaDias <= 0) {
    throw new Error('FrecuenciaDias no válida para ' + modalidad + '.');
  }

  if (!config.SesionesPorCiclo || config.SesionesPorCiclo <= 0) {
    throw new Error('SesionesPorCiclo no válida para ' + modalidad + '.');
  }

  if (!config.CapacidadMaxima || config.CapacidadMaxima <= 0) {
    throw new Error('CapacidadMaxima no válida para ' + modalidad + '.');
  }

  if (!(config.FechaBase instanceof Date)) {
    throw new Error('FechaBase obligatoria y válida para ' + modalidad + '.');
  }
}

function validarFechaInicioCiclo_(fechaInicio, config) {
  const diaEsperado = convertirDiaSemanaATexto_(fechaInicio);

  if (diaEsperado !== config.DiaSemana) {
    throw new Error(
      'La fecha introducida no cae en el día configurado para el grupo.\n\n' +
      'Esperado: ' + config.DiaSemana + '\n' +
      'Recibido: ' + diaEsperado
    );
  }
}

/**
 * Genera los slots horarios para un ciclo de grupo.
 * @param {Object} params - Parámetros para la generación.
 * @param {Date} params.fechaInicio - Fecha de inicio del ciclo.
 * @param {string} params.horaInicio - Hora de inicio del primer slot del ciclo.
 * @param {string} params.modalidad - Modalidad del grupo.
 * @returns {Array<AgendaSlot>} Lista de slots generados.
 */
function generarSlotsCiclo_({ fechaInicio, horaInicio, modalidad }) {
  const config = obtenerConfigModalidad_(modalidad);
  const sesiones = Number(config.SesionesPorCiclo || 7);
  const frecuenciaSemanas = Number(config.FrecuenciaDias || 1);
  
  const availabilityService = new AvailabilityService();
  const slots = [];
  let currentSearchDateTime = normalizarFechaHora_(fechaInicio, horaInicio);

  for (let i = 0; i < sesiones; i++) {
    // Los grupos suelen ocupar 60 min (2 slots de 30) según la lógica de AgendaService
    const slot = availabilityService.findNextAvailableSlot(currentSearchDateTime, modalidad, 60);
    
    if (!slot) {
      Logger.log(`No se encontró slot para sesión ${i + 1} del ciclo ${modalidad}`);
      break; 
    }
    
    slots.push(slot);
    
    // Avanzar a la misma hora pero N semanas después
    const siguienteFecha = sumarSemanasManteniendoDia_(slot.startDateTime, frecuenciaSemanas);
    currentSearchDateTime = normalizarFechaHora_(siguienteFecha, horaInicio);
  }
  
  return slots;
}

function crearCicloEnSheet_({ modalidad, fechaInicio, fechas, config }) {
  const cicloRepo = new CicloRepository();
  const todos = cicloRepo.findAll();
  
  const numeroCiclo = obtenerSiguienteNumeroCiclo_(modalidad, todos);
  const cicloId = generarId_('CIC');
  const fechaFin = fechas[fechas.length - 1];

  const nuevoCiclo = {
    CicloID: cicloId,
    Modalidad: modalidad,
    NumeroCiclo: numeroCiclo,
    EstadoCiclo: ESTADOS_CICLO.PLANIFICADO,
    FechaInicioCiclo: normalizarFecha_(fechaInicio),
    FechaFinCiclo: normalizarFecha_(fechaFin),
    FechaBaseUsada: normalizarFecha_(config.FechaBase),
    DiaSemana: config.DiaSemana,
    FrecuenciaDias: config.FrecuenciaDias,
    SesionesPorCiclo: config.SesionesPorCiclo,
    CapacidadMaxima: config.CapacidadMaxima,
    PlazasOcupadas: 0,
    PlazasLibres: config.CapacidadMaxima,
    GeneradoManual: true,
    Notas: ''
  };

  cicloRepo.save(nuevoCiclo);

  return {
    cicloId,
    numeroCiclo
  };
}

function obtenerSiguienteNumeroCiclo_(modalidad, todosLosCiclos) {
  let maximo = 0;

  todosLosCiclos.forEach(c => {
    if (c.Modalidad === modalidad) {
      const numero = Number(c.NumeroCiclo || 0);
      if (numero > maximo) {
        maximo = numero;
      }
    }
  });

  return maximo + 1;
}

function convertirDiaSemanaATexto_(fecha) {
  const dias = [
    DIAS_SEMANA.DOMINGO,
    DIAS_SEMANA.LUNES,
    DIAS_SEMANA.MARTES,
    DIAS_SEMANA.MIERCOLES,
    DIAS_SEMANA.JUEVES,
    DIAS_SEMANA.VIERNES,
    DIAS_SEMANA.SABADO
  ];

  return dias[fecha.getDay()];
}

function formatearFecha_(fecha) {
  if (!(fecha instanceof Date)) return '';
  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function obtenerOpcionesGrupoFormulario() {
  return [
    { value: MODALIDADES.GRUPO_1, label: 'GRUPO_1' },
    { value: MODALIDADES.GRUPO_2, label: 'GRUPO_2' },
    { value: MODALIDADES.GRUPO_3, label: 'GRUPO_3' }
  ];
}

function guardarCicloGrupoDesdeFormulario(formData) {
  const modalidad = String(formData.modalidad || '').trim();
  const fechaISO = String(formData.fechaInicio || '').trim();

  const modalidadesValidas = [
    MODALIDADES.GRUPO_1,
    MODALIDADES.GRUPO_2,
    MODALIDADES.GRUPO_3
  ];

  if (!modalidadesValidas.includes(modalidad)) {
    throw new Error('La modalidad de grupo no es válida.');
  }

  if (!fechaISO) {
    throw new Error('La fecha de inicio es obligatoria.');
  }

  const fechaInicio = parseFechaISO_(fechaISO);
  if (!(fechaInicio instanceof Date)) {
    throw new Error('La fecha de inicio no es válida.');
  }

  const config = obtenerConfigModalidad_(modalidad);
  validarConfigGrupo_(modalidad, config);
  validarFechaInicioCiclo_(fechaInicio, config);

  const slots = generarSlotsCiclo_({
  fechaInicio,
  horaInicio: config.HoraBase,
  modalidad: modalidad
  });

  // Temporalmente usamos la fecha de inicio para el resto de la lógica hasta completar el stub
  const fechas = slots.length > 0 ? slots.map(s => s.startDateTime) : [fechaInicio];
  const avisos = [];

  const ciclo = crearCicloEnSheet_({
    modalidad,
    fechaInicio,
    fechas,
    config
  });
  const fechaFin = fechas[fechas.length - 1];

  let mensaje =
  'Ciclo creado correctamente.\n\n' +
  'Modalidad: ' + modalidad + '\n' +
  'Inicio: ' + formatearFecha_(fechaInicio) + '\n' +
  'Fin: ' + formatearFecha_(fechaFin);

  if (avisos.length > 0) {
    mensaje += '\n\nAvisos:\n- ' + avisos.join('\n- ');
  }

  return {
    cicloId: ciclo.cicloId,
    mensaje: mensaje
  };
}