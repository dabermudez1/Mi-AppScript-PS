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
 * HELPERS DATOS
 ***************/
function obtenerPacientePorId_(pacienteId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = indexByHeader_(headers);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.PacienteID]) === String(pacienteId)) {
      return mapPaciente_(data[i], idx);
    }
  }

  return null;
}

function obtenerCicloPorId_(cicloId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CICLOS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = indexByHeader_(headers);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.CicloID]) === String(cicloId)) {
      return mapCiclo_(data[i], idx);
    }
  }

  return null;
}

function mapPaciente_(row, idx) {
  return {
    PacienteID: row[idx.PacienteID],
    Nombre: row[idx.Nombre],
    ModalidadSolicitada: row[idx.ModalidadSolicitada],
    EstadoPaciente: row[idx.EstadoPaciente],
    FechaPrimeraSesionReal: row[idx.FechaPrimeraSesionReal],
    SesionesPlanificadas: row[idx.SesionesPlanificadas]
  };
}

function mapCiclo_(row, idx) {
  return {
    CicloID: row[idx.CicloID],
    FechaInicioCiclo: row[idx.FechaInicioCiclo],
    DiaSemana: row[idx.DiaSemana],
    FrecuenciaDias: Number(row[idx.FrecuenciaDias] || 0),
    SesionesPorCiclo: Number(row[idx.SesionesPorCiclo] || 0)
  };
}

/***************
 * MANTENIMIENTO
 ***************/
function generarSesionesFaltantes() {
  const ui = SpreadsheetApp.getUi();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPac = ss.getSheetByName(SHEET_PACIENTES);
    const sheetSes = ss.getSheetByName(SHEET_SESIONES);

    if (!sheetPac || !sheetSes) {
      throw new Error('Faltan hojas PACIENTES o SESIONES.');
    }

    const pacData = sheetPac.getDataRange().getValues();
    const sesData = sheetSes.getDataRange().getValues();

    if (pacData.length < 2) {
      ui.alert('No hay pacientes.');
      return;
    }

    const pIdx = indexByHeader_(pacData[0]);
    const sIdx = sesData.length > 0 ? indexByHeader_(sesData[0]) : {};

    const sesionesPorPaciente = contarSesionesPorPaciente_(sesData, sIdx);

    let generadasIndividual = 0;
    let generadasGrupo = 0;
    let omitidasConSesiones = 0;
    let omitidasSinCondiciones = 0;
    let avisos = [];

    for (let i = 1; i < pacData.length; i++) {
      const paciente = mapPacienteCompletoDesdeFila_(pacData[i], pIdx);
      const totalExistentes = sesionesPorPaciente[String(paciente.PacienteID)] || 0;

      if (totalExistentes > 0) {
        omitidasConSesiones++;
        continue;
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
        continue;
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
    }

    let mensaje =
      'Generación de sesiones faltantes completada.\n\n' +
      'Individuales generadas: ' + generadasIndividual + '\n' +
      'Grupales generadas: ' + generadasGrupo + '\n' +
      'Omitidas (ya tenían sesiones): ' + omitidasConSesiones + '\n' +
      'Omitidas (sin condiciones válidas): ' + omitidasSinCondiciones;

    if (avisos.length > 0) {
      mensaje += '\n\nAvisos:\n- ' + avisos.join('\n- ');
    }

    ui.alert(mensaje);

  } catch (error) {
    ui.alert('Error al generar sesiones faltantes: ' + error.message);
    throw error;
  }
}

function contarSesionesPorPaciente_(sesData, sIdx) {
  const map = {};

  if (!sesData || sesData.length < 2 || sIdx.PacienteID === undefined) {
    return map;
  }

  for (let i = 1; i < sesData.length; i++) {
    const pacienteId = String(sesData[i][sIdx.PacienteID] || '');
    if (!pacienteId) continue;

    if (!map[pacienteId]) {
      map[pacienteId] = 0;
    }
    map[pacienteId]++;
  }

  return map;
}

function mapPacienteCompletoDesdeFila_(row, idx) {
  return {
    PacienteID: row[idx.PacienteID],
    Nombre: row[idx.Nombre],
    ModalidadSolicitada: row[idx.ModalidadSolicitada],
    EstadoPaciente: row[idx.EstadoPaciente],
    CicloObjetivoID: row[idx.CicloObjetivoID],
    CicloActivoID: row[idx.CicloActivoID],
    FechaPrimeraSesionReal: row[idx.FechaPrimeraSesionReal],
    SesionesPlanificadas: Number(row[idx.SesionesPlanificadas] || 0),
    SesionesCompletadas: Number(row[idx.SesionesCompletadas] || 0),
    SesionesPendientes: Number(row[idx.SesionesPendientes] || 0),
    ProximaSesion: row[idx.ProximaSesion]
  };
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