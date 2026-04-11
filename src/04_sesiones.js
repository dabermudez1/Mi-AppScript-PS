/***********************
 * BLOQUE 4
 * GENERACIÓN DE SESIONES
 ***********************/

/***************
 * ENTRY POINTS
 ***************/
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
          // Llamamos a la lógica moderna de 03_pacientes.js
          generarSesionesPacienteGrupo_(paciente.PacienteID, cicloId);
          
          const resultadoGrupo = { avisos: [] }; // Mock para compatibilidad con el resto del bucle

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