/***********************
 * BLOQUE 4
 * GENERACIÓN DE SESIONES
 ***********************/

/***************
 * ENTRY POINTS
 ***************/
function generarSesionesPacienteIndividual_(pacienteId) {
  const patientRepo = new PatientRepository();
  const configRepo = new ConfigRepository();
  const sessionService = new SessionService();
  const paciente = patientRepo.findById(pacienteId);

  if (!paciente) {
    throw new Error('Paciente no encontrado: ' + pacienteId);
  }

  if (paciente.EstadoPaciente !== ESTADOS_PACIENTE.ACTIVO) {
    return { avisos: [] };
  }

  const config = configRepo.findByModalidad(paciente.ModalidadSolicitada);
  const intervaloDias = Number(config.FrecuenciaDias || 0);

  if (intervaloDias <= 0) {
    throw new Error(
      'La frecuencia de la modalidad individual no es válida.\n\n' +
      'Modalidad: ' + paciente.ModalidadSolicitada
    );
  }

  const resultado = generarFechasIndividualConAvisos_({
    fechaInicio: paciente.FechaPrimeraSesionReal,
    intervaloDias: intervaloDias,
    sesiones: paciente.SesionesPlanificadas
  });

  sessionService.createInitialSessions(paciente, resultado.fechas);

  return {
    avisos: resultado.avisos || []
  };
}

function generarSesionesPacienteGrupo_(pacienteId, cicloId) {
  const patientRepo = new PatientRepository();
  const cicloRepo = new CicloRepository();
  const sessionService = new SessionService();
  const paciente = patientRepo.findById(pacienteId);
  const ciclo = cicloRepo.findOneBy('CicloID', cicloId);

  if (!paciente || !ciclo) {
    throw new Error('Paciente o ciclo no encontrado.');
  }

  const resultado = generarFechasCiclo_({
    fechaInicio: ciclo.FechaInicioCiclo,
    diaSemana: ciclo.DiaSemana,
    frecuenciaDias: ciclo.FrecuenciaDias,
    sesiones: ciclo.SesionesPorCiclo
  });

  sessionService.createInitialSessions(paciente, resultado.fechas, cicloId);

  return {
    avisos: resultado.avisos || []
  };
}

/***************
 * MANTENIMIENTO
 ***************/
function generarSesionesFaltantes() {
  const props = PropertiesService.getUserProperties();
  const patientRepo = new PatientRepository();
  const sessionRepo = new SessionRepository();

  props.setProperty('TASK_GENERATE_SESSIONS_RUNNING', 'true');
  props.setProperty('TASK_GENERATE_SESSIONS_PROGRESS', '0');

  try {
    const pacientes = patientRepo.findAll();
    const todasLasSesiones = sessionRepo.findAll();

    if (pacientes.length === 0) {
      props.setProperty('TASK_GENERATE_SESSIONS_RUNNING', 'false');
      props.setProperty('TASK_GENERATE_SESSIONS_RESULT', 'No hay pacientes registrados.');
      return;
    }

    const sesionesPorPaciente = sessionRepo.mapByPatient(todasLasSesiones);

    let generadasIndividual = 0;
    let generadasGrupo = 0;
    let omitidasConSesiones = 0;
    let omitidasSinCondiciones = 0;
    let avisos = [];

    const total = pacientes.length;

    pacientes.forEach((paciente, index) => {
      if (index % 5 === 0) props.setProperty('TASK_GENERATE_SESSIONS_PROGRESS', Math.round((index / total) * 100).toString());

      const totalExistentes = (sesionesPorPaciente[String(paciente.PacienteID)] || []).length;

      if (totalExistentes > 0) {
        omitidasConSesiones++;
        return;
      }

      if (paciente.ModalidadSolicitada === MODALIDADES.INDIVIDUAL) {
        if (
          paciente.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO &&
          paciente.FechaPrimeraSesionReal instanceof Date &&
          Number(paciente.SesionesPlanificadas || 0) > 0
        ) {
          const resultadoIndividual = generarSesionesPacienteIndividual_(paciente.PacienteID);

          if (resultadoIndividual && resultadoIndividual.avisos && resultadoIndividual.avisos.length) {
            avisos = avisos.concat(
              resultadoIndividual.avisos.map(a => paciente.Nombre + ' (' + paciente.PacienteID + '): ' + a)
            );
          }

          generadasIndividual++;
        } else {
          omitidasSinCondiciones++;
        }
        return;
      }

      const esGrupo =
        paciente.ModalidadSolicitada === MODALIDADES.GRUPO_1 ||
        paciente.ModalidadSolicitada === MODALIDADES.GRUPO_2 ||
        paciente.ModalidadSolicitada === MODALIDADES.GRUPO_3;

      if (esGrupo) {
        const cicloId = paciente.CicloObjetivoID || paciente.CicloActivoID || '';

        if (
          (paciente.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO ||
          paciente.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO) &&
          cicloId
        ) {
          const resultadoGrupo = generarSesionesPacienteGrupo_(paciente.PacienteID, cicloId);

          if (resultadoGrupo && resultadoGrupo.avisos && resultadoGrupo.avisos.length) {
            avisos = avisos.concat(
              resultadoGrupo.avisos.map(a => paciente.Nombre + ' (' + paciente.PacienteID + '): ' + a)
            );
          }

          generadasGrupo++;
        } else {
          omitidasSinCondiciones++;
        }
      }
    });

    let mensaje =
      'Generación de sesiones faltantes completada.\n\n' +
      'Individuales generadas: ' + generadasIndividual + '\n' +
      'Grupales generadas: ' + generadasGrupo + '\n' +
      'Omitidas (ya tenían sesiones): ' + omitidasConSesiones + '\n' +
      'Omitidas (sin condiciones válidas): ' + omitidasSinCondiciones;

    props.setProperty('TASK_GENERATE_SESSIONS_PROGRESS', '100');
    props.setProperty('TASK_GENERATE_SESSIONS_RESULT', mensaje);
    props.setProperty('TASK_GENERATE_SESSIONS_RUNNING', 'false');

  } catch (error) {
    props.setProperty('TASK_GENERATE_SESSIONS_RUNNING', 'false');
    props.setProperty('TASK_GENERATE_SESSIONS_RESULT', 'Error: ' + error.message);
  }
}

function generarFechasIndividualConAvisos_({
  fechaInicio,
  intervaloDias,
  sesiones
}) {
  const fechas = [];
  const avisos = [];

  for (let i = 0; i < sesiones; i++) {
    const base = sumarDiasNaturales_(fechaInicio, i * intervaloDias);
    const ajuste = ajustarASiguienteFechaOperativaConAviso_(base);

    fechas.push(ajuste.fecha);

    if (ajuste.ajustada) {
      avisos.push('Sesión ' + (i + 1) + ': ' + ajuste.aviso);
    }
  }

  return {
    fechas,
    avisos
  };
}

function obtenerConfigModalidadPorNombre_(modalidad) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG_MODALIDADES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_CONFIG_MODALIDADES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('No hay datos en la hoja ' + SHEET_CONFIG_MODALIDADES + '.');
  }

  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.Modalidad]) === String(modalidad)) {
      return {
        Modalidad: data[i][idx.Modalidad],
        TipoModalidad: data[i][idx.TipoModalidad],
        FrecuenciaDias: Number(data[i][idx.FrecuenciaDias] || 0),
        SesionesPorCiclo: Number(data[i][idx.SesionesPorCiclo] || 0)
      };
    }
  }

  throw new Error('No se encontró configuración para la modalidad: ' + modalidad);
}

function abrirPantallaSesionesDesdePantalla() {
  abrirPantallaSesiones();
}