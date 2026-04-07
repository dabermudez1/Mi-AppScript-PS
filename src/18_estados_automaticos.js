/***********************
 * BLOQUE 18
 * ESTADOS AUTOMÁTICOS
 ***********************/

function recalcularEstadosAutomaticamente() {
  const patientRepo = new PatientRepository();
  const sessionRepo = new SessionRepository();
  
  const pacientes = patientRepo.findAll();
  const todasLasSesiones = sessionRepo.findAll();
  
  if (pacientes.length === 0) {
    return { mensaje: 'No hay pacientes para recalcular.' };
  }

  // Patrón C: Mapa de sesiones por paciente para búsqueda instantánea
  const sesionesPorPaciente = sessionRepo.mapByPatient(todasLasSesiones);
  const pacientesAActualizar = [];
  const errores = [];
  let actualizados = 0;

  pacientes.forEach(paciente => {
    const estadoActual = paciente.EstadoPaciente;
    const nombre = paciente.Nombre || '';
    const pacienteId = paciente.PacienteID;

    if (estadoActual === ESTADOS_PACIENTE.ALTA) {
      return;
    }

    const cicloObjetivoId = paciente.CicloObjetivoID || '';
    const cicloActivoId = paciente.CicloActivoID || '';
    const tieneCiclo = !!(cicloObjetivoId || cicloActivoId);

    const sesiones = sesionesPorPaciente[String(pacienteId)] || [];
    const analisis = analizarSesionesPaciente_(sesiones);

    const estadoCalculado = calcularEstadoPacienteAutomatico_({
      modalidad: paciente.ModalidadSolicitada,
      tieneCiclo,
      analisis
    });

    const erroresPaciente = detectarErroresEstadoPaciente_({ ...paciente, tieneCiclo, analisis });

    if (erroresPaciente.length) {
      errores.push.apply(errores, erroresPaciente);
    }

    if (estadoCalculado && estadoCalculado !== estadoActual) {
      paciente.EstadoPaciente = estadoCalculado;
      pacientesAActualizar.push(paciente);
      actualizados++;
    }
  });

  // Patrón B: Un solo volcado a Sheets
  if (pacientesAActualizar.length > 0) {
    patientRepo.saveAll(pacientesAActualizar);
  }

  let mensaje =
    'Recalculo de estados completado.\n\n' +
    'Pacientes actualizados: ' + actualizados + '\n' +
    'Errores detectados: ' + errores.length;

  if (errores.length > 0) {
    mensaje += '\n\nErrores:\n- ' + errores.join('\n- ');
  }

  return { mensaje: mensaje };
}

function agruparSesionesPorPaciente_(sesData, sIdx) {
  const map = {};

  if (!sesData || sesData.length < 2 || sIdx.PacienteID === undefined) {
    return map;
  }

  for (let i = 1; i < sesData.length; i++) {
    const pacienteId = String(sesData[i][sIdx.PacienteID] || '');
    if (!pacienteId) continue;

    if (!map[pacienteId]) {
      map[pacienteId] = [];
    }

    map[pacienteId].push({
      estadoSesion: sesData[i][sIdx.EstadoSesion] || '',
      fechaSesion: sesData[i][sIdx.FechaSesion] || ''
    });
  }

  return map;
}

function analizarSesionesPaciente_(sesiones) {
  const hoy = normalizarFecha_(new Date());

  let completadas = 0;
  let pendientesFuturas = 0;
  let pendientesVencidas = 0;
  let total = 0;

  (sesiones || []).forEach(function(s) {
    total++;

    const estado = s.estadoSesion || '';
    const fecha = s.fechaSesion ? normalizarFecha_(s.fechaSesion) : null;

    if (
      estado === ESTADOS_SESION.COMPLETADA_AUTO ||
      estado === ESTADOS_SESION.COMPLETADA_MANUAL
    ) {
      completadas++;
      return;
    }

    if (
      estado === ESTADOS_SESION.PENDIENTE ||
      estado === ESTADOS_SESION.REPROGRAMADA
    ) {
      if (fecha && fecha.getTime() >= hoy.getTime()) {
        pendientesFuturas++;
      } else {
        pendientesVencidas++;
      }
    }
  });

  return {
    total,
    completadas,
    pendientesFuturas,
    pendientesVencidas
  };
}

function calcularEstadoPacienteAutomatico_({ modalidad, tieneCiclo, analisis }) {
  const esIndividual = modalidad === MODALIDADES.INDIVIDUAL;

  if (esIndividual) {
    if (analisis.completadas > 0) {
      return ESTADOS_PACIENTE.ACTIVO;
    }

    if (analisis.pendientesFuturas > 0) {
      return ESTADOS_PACIENTE.ACTIVO;
    }

    if (analisis.pendientesVencidas > 0) {
      return ESTADOS_PACIENTE.ACTIVO;
    }

    return ESTADOS_PACIENTE.ESPERA;
  }

  if (!tieneCiclo) {
    return ESTADOS_PACIENTE.ESPERA;
  }

  if (analisis.completadas > 0) {
    return ESTADOS_PACIENTE.ACTIVO;
  }

  if (analisis.pendientesFuturas > 0 || analisis.total === 0) {
    return ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO;
  }

  if (analisis.pendientesVencidas > 0) {
    return ESTADOS_PACIENTE.ACTIVO;
  }

  return ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO;
}

function detectarErroresEstadoPaciente_(ctx) {
  const errores = [];

  if (ctx.tieneCiclo && ctx.analisis.total === 0) {
    errores.push(
      ctx.nombre + ' (' + ctx.pacienteId + '): tiene ciclo asignado pero no tiene sesiones generadas.'
    );
  }

  if (!ctx.tieneCiclo && ctx.estadoActual === ESTADOS_PACIENTE.ACTIVO) {
    errores.push(
      ctx.nombre + ' (' + ctx.pacienteId + '): figura ACTIVO pero no tiene ciclo asignado.'
    );
  }

  if (!ctx.tieneCiclo && ctx.estadoActual === ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO) {
    errores.push(
      ctx.nombre + ' (' + ctx.pacienteId + '): figura ACTIVO_PENDIENTE_INICIO pero no tiene ciclo asignado.'
    );
  }

  if (ctx.tieneCiclo && ctx.analisis.pendientesFuturas === 0 && ctx.analisis.completadas === 0) {
    errores.push(
      ctx.nombre + ' (' + ctx.pacienteId + '): tiene ciclo asignado pero no tiene sesiones futuras ni completadas.'
    );
  }

  return errores;
}

function recalcularEstadosAutomaticamenteConModal() {
  const resultado = recalcularEstadosAutomaticamente();

  const template = HtmlService.createTemplateFromFile('ResultadoProcesoForm');
  template.titulo = 'Resultado de actualización de estados';
  template.mensaje = (resultado && resultado.mensaje)
    ? resultado.mensaje
    : 'Proceso completado.';

  const html = template
    .evaluate()
    .setWidth(760)
    .setHeight(560);

  SpreadsheetApp.getUi().showModalDialog(html, 'Resultado');
}