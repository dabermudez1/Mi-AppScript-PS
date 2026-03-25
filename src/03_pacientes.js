/***********************
 * BLOQUE 3
 * GESTIÓN DE PACIENTES
 ***********************/

function nuevoPaciente() {
  const html = HtmlService
    .createHtmlOutputFromFile('NuevoPacienteForm')
    .setWidth(420)
    .setHeight(360);

  SpreadsheetApp.getUi().showModalDialog(html, 'Nuevo paciente');
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
    fechaPrimeraConsulta
  }) {
  const config = obtenerConfigModalidad_(modalidad);

  if (!config.Activa) {
    throw new Error('La modalidad está inactiva: ' + modalidad);
  }

  if (modalidad === MODALIDADES.INDIVIDUAL) {
    return procesarAltaIndividual_({
      nombre,
      nhc,
      sexoGenero,
      motivoConsultaDiagnostico,
      motivoConsultaOtros,
      modalidad,
      fechaPrimeraConsulta,
      config
    });
  }

  return procesarAltaGrupo_({
    nombre,
    nhc,
    sexoGenero,
    motivoConsultaDiagnostico,
    motivoConsultaOtros,
    modalidad,
    fechaPrimeraConsulta,
    config
  });
}

/***************
 * INDIVIDUAL
 ***************/
function procesarAltaIndividual_({
    nombre,
    nhc,
    sexoGenero,
    motivoConsultaDiagnostico,
    motivoConsultaOtros,
    modalidad,
    fechaPrimeraConsulta,
    config
  }) {
  const tieneCapacidad = hayCapacidadIndividual_();

  if (!tieneCapacidad) {
    const pacienteId = crearPacienteEnSheet_({
      nombre,
      nhc,
      sexoGenero,
      motivoConsultaDiagnostico,
      motivoConsultaOtros,
      modalidadSolicitada: modalidad,
      fechaPrimeraConsulta,
      estadoPaciente: ESTADOS_PACIENTE.ESPERA,
      motivoEspera: 'SIN_CAPACIDAD_INDIVIDUAL',
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
        'Motivo: SIN_CAPACIDAD_INDIVIDUAL'
    };
  }

  const fechaPrimeraSesionReal = calcularPrimeraSesionIndividual_(fechaPrimeraConsulta, modalidad);
  const totalSesiones = Number(config.SesionesPorCiclo || 7);

  const pacienteId = crearPacienteEnSheet_({
    nombre,
    nhc,
    sexoGenero,
    motivoConsultaDiagnostico,
    motivoConsultaOtros,
    modalidadSolicitada: modalidad,
    fechaPrimeraConsulta,
    estadoPaciente: ESTADOS_PACIENTE.ACTIVO,
    motivoEspera: '',
    cicloObjetivoId: '',
    cicloActivoId: '',
    fechaPrimeraSesionReal,
    sesionesPlanificadas: totalSesiones,
    sesionesCompletadas: 0,
    sesionesPendientes: totalSesiones,
    proximaSesion: fechaPrimeraSesionReal,
    fechaCierre: '',
    observaciones: '',
    recalcularSecuencia: false
  });

  const resultadoSesiones = generarSesionesPacienteIndividual_(pacienteId);
  const avisos = (resultadoSesiones && resultadoSesiones.avisos) ? resultadoSesiones.avisos : [];

  let mensaje =
    'Paciente creado correctamente.\n\n' +
    'Nombre: ' + nombre + '\n' +
    'Modalidad: ' + modalidad + '\n' +
    'Estado: ACTIVO\n' +
    'Primera sesión real: ' + formatearFecha_(fechaPrimeraSesionReal);

  if (avisos.length > 0) {
    mensaje += '\n\nAvisos:\n- ' + avisos.join('\n- ');
  }

  return {
    pacienteId,
    mensaje: mensaje
  };
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

  const fechaBase = sumarDiasNaturales_(fechaPrimeraConsulta, intervaloDias);
  return ajustarASiguienteFechaOperativa_(fechaBase);
}

function hayCapacidadIndividual_() {
  const config = obtenerConfigModalidad_(MODALIDADES.INDIVIDUAL);
  const capacidadMaxima = Number(config.CapacidadMaxima || 0);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return capacidadMaxima > 0;
  }

  const headers = data[0];
  const idx = indexByHeader_(headers);

  const columnasNecesarias = ['ModalidadSolicitada', 'EstadoPaciente'];
  columnasNecesarias.forEach(col => {
    if (idx[col] === undefined) {
      throw new Error('Falta la columna "' + col + '" en ' + SHEET_PACIENTES + '.');
    }
  });

  let activos = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const modalidad = row[idx.ModalidadSolicitada];
    const estado = row[idx.EstadoPaciente];

    if (
      modalidad === MODALIDADES.INDIVIDUAL &&
      estado === ESTADOS_PACIENTE.ACTIVO
    ) {
      activos++;
    }
  }

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

  const ciclo = buscarPrimerCicloFuturoDisponible_(modalidad, fechaPrimeraConsulta);

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
    observaciones: ''
  });
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

function buscarPrimerCicloFuturoDisponible_(modalidad, fechaPrimeraConsulta) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CICLOS);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_CICLOS + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  const headers = data[0];
  const idx = indexByHeader_(headers);

  const columnasNecesarias = [
    'CicloID',
    'Modalidad',
    'EstadoCiclo',
    'FechaInicioCiclo',
    'SesionesPorCiclo',
    'CapacidadMaxima',
    'PlazasOcupadas',
    'PlazasLibres'
  ];

  columnasNecesarias.forEach(col => {
    if (idx[col] === undefined) {
      throw new Error('Falta la columna "' + col + '" en ' + SHEET_CICLOS + '.');
    }
  });

  const candidatos = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowModalidad = row[idx.Modalidad];
    const estadoCiclo = row[idx.EstadoCiclo];
    const fechaInicio = row[idx.FechaInicioCiclo];
    const capacidadMaxima = Number(row[idx.CapacidadMaxima] || 0);
    const plazasOcupadas = Number(row[idx.PlazasOcupadas] || 0);
    const plazasLibresReales = capacidadMaxima - plazasOcupadas;

    if (rowModalidad !== modalidad) continue;
    if (estadoCiclo !== ESTADOS_CICLO.PLANIFICADO) continue;
    if (!(fechaInicio instanceof Date)) continue;
    if (!(normalizarFecha_(fechaInicio) > normalizarFecha_(fechaPrimeraConsulta))) continue;
    if (plazasLibresReales <= 0) continue;

    candidatos.push({
      fila: i + 1,
      CicloID: row[idx.CicloID],
      Modalidad: row[idx.Modalidad],
      EstadoCiclo: row[idx.EstadoCiclo],
      FechaInicioCiclo: normalizarFecha_(row[idx.FechaInicioCiclo]),
      SesionesPorCiclo: Number(row[idx.SesionesPorCiclo] || 0),
      CapacidadMaxima: capacidadMaxima,
      PlazasOcupadas: plazasOcupadas,
      PlazasLibres: plazasLibresReales
    });
  }

  if (candidatos.length === 0) return null;

  candidatos.sort((a, b) => a.FechaInicioCiclo - b.FechaInicioCiclo);
  return candidatos[0];
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');
  }

  const pacienteId = generarId_('PAC');
  const fechaAlta = normalizarFecha_(new Date());

  const row = [
    pacienteId,
    params.nombre,
    params.nhc || '',
    params.sexoGenero || '',
    params.motivoConsultaDiagnostico || '',
    params.motivoConsultaOtros || '',
    params.modalidadSolicitada,
    fechaAlta,
    normalizarFecha_(params.fechaPrimeraConsulta),
    params.estadoPaciente,
    params.motivoEspera || '',
    params.cicloObjetivoId || '',
    params.cicloActivoId || '',
    params.fechaPrimeraSesionReal instanceof Date ? normalizarFecha_(params.fechaPrimeraSesionReal) : '',
    Number(params.sesionesPlanificadas || 0),
    Number(params.sesionesCompletadas || 0),
    Number(params.sesionesPendientes || 0),
    params.proximaSesion instanceof Date ? normalizarFecha_(params.proximaSesion) : '',
    params.fechaCierre instanceof Date ? normalizarFecha_(params.fechaCierre) : '',
    params.observaciones || '',
    params.recalcularSecuencia === true
  ];

  sheet.appendRow(row);
  return pacienteId;
}

function crearAsignacionCiclo_({ pacienteId, cicloId, modalidad, estadoAsignacion, observaciones }) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ASIGNACIONES_CICLO);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_ASIGNACIONES_CICLO + '.');
  }

  const asignacionId = generarId_('ASI');

  const row = [
    asignacionId,
    pacienteId,
    cicloId,
    modalidad,
    normalizarFecha_(new Date()),
    estadoAsignacion,
    observaciones || ''
  ];

  sheet.appendRow(row);
  return asignacionId;
}

function actualizarPlazasCiclo_(cicloId, delta, devolverFalseSiNoHayPlaza) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CICLOS);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_CICLOS + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('La hoja CICLOS no tiene datos.');
  }

  const headers = data[0];
  const idx = indexByHeader_(headers);

  const columnasNecesarias = ['CicloID', 'CapacidadMaxima', 'PlazasOcupadas', 'PlazasLibres'];
  columnasNecesarias.forEach(col => {
    if (idx[col] === undefined) {
      throw new Error('Falta la columna "' + col + '" en ' + SHEET_CICLOS + '.');
    }
  });

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.CicloID]) === String(cicloId)) {
      const capacidadMaxima = Number(data[i][idx.CapacidadMaxima] || 0);
      const ocupadasActuales = Number(data[i][idx.PlazasOcupadas] || 0);
      const nuevasOcupadas = ocupadasActuales + delta;

      if (nuevasOcupadas < 0) {
        throw new Error('Las plazas ocupadas no pueden quedar en negativo.');
      }

      if (nuevasOcupadas > capacidadMaxima) {
        if (devolverFalseSiNoHayPlaza === true) {
          return false;
        }
        throw new Error('El ciclo supera la capacidad máxima.');
      }

      const nuevasLibres = capacidadMaxima - nuevasOcupadas;

      sheet.getRange(i + 1, idx.PlazasOcupadas + 1).setValue(nuevasOcupadas);
      sheet.getRange(i + 1, idx.PlazasLibres + 1).setValue(nuevasLibres);
      return true;
    }
  }

  throw new Error('No existe el ciclo con ID ' + cicloId + '.');
}

function obtenerOpcionesModalidadFormulario() {
  return [
    { value: MODALIDADES.INDIVIDUAL, label: 'INDIVIDUAL' },
    { value: MODALIDADES.GRUPO_1, label: 'GRUPO_1' },
    { value: MODALIDADES.GRUPO_2, label: 'GRUPO_2' },
    { value: MODALIDADES.GRUPO_3, label: 'GRUPO_3' }
  ];
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
    fechaPrimeraConsulta
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
    .setHeight(520);

  SpreadsheetApp.getUi().showModalDialog(html, 'Asignar paciente en espera a ciclo');
}

function obtenerPacientesEnEspera_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const idx = indexByHeader_(data[0]);

  const columnasNecesarias = [
    'PacienteID',
    'Nombre',
    'ModalidadSolicitada',
    'FechaPrimeraConsulta',
    'EstadoPaciente',
    'MotivoEspera'
  ];

  columnasNecesarias.forEach(col => {
    if (idx[col] === undefined) {
      throw new Error('Falta la columna "' + col + '" en ' + SHEET_PACIENTES + '.');
    }
  });

  const out = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (row[idx.EstadoPaciente] === ESTADOS_PACIENTE.ESPERA) {
      out.push({
        fila: i + 1,
        PacienteID: row[idx.PacienteID],
        Nombre: row[idx.Nombre],
        ModalidadSolicitada: row[idx.ModalidadSolicitada],
        FechaPrimeraConsulta: row[idx.FechaPrimeraConsulta],
        MotivoEspera: row[idx.MotivoEspera] || ''
      });
    }
  }

  out.sort((a, b) => compararFechas_(a.FechaPrimeraConsulta, b.FechaPrimeraConsulta));
  return out;
}

function obtenerCiclosDisponiblesParaPacienteEnEspera_(paciente) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CICLOS);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_CICLOS + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const idx = indexByHeader_(data[0]);

  const columnasNecesarias = [
    'CicloID',
    'Modalidad',
    'NumeroCiclo',
    'EstadoCiclo',
    'FechaInicioCiclo',
    'SesionesPorCiclo',
    'CapacidadMaxima',
    'PlazasOcupadas',
    'PlazasLibres'
  ];

  columnasNecesarias.forEach(col => {
    if (idx[col] === undefined) {
      throw new Error('Falta la columna "' + col + '" en ' + SHEET_CICLOS + '.');
    }
  });

  const ciclos = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const modalidad = row[idx.Modalidad];
    const estado = row[idx.EstadoCiclo];
    const fechaInicio = row[idx.FechaInicioCiclo];
    const capacidad = Number(row[idx.CapacidadMaxima] || 0);
    const ocupadas = Number(row[idx.PlazasOcupadas] || 0);
    const libresReales = capacidad - ocupadas;

    if (modalidad !== paciente.ModalidadSolicitada) continue;
    if (estado !== ESTADOS_CICLO.PLANIFICADO) continue;
    if (!(fechaInicio instanceof Date)) continue;
    if (!(normalizarFecha_(fechaInicio).getTime() > normalizarFecha_(paciente.FechaPrimeraConsulta).getTime())) continue;
    if (libresReales <= 0) continue;

    ciclos.push({
      fila: i + 1,
      CicloID: row[idx.CicloID],
      Modalidad: modalidad,
      NumeroCiclo: Number(row[idx.NumeroCiclo] || 0),
      EstadoCiclo: estado,
      FechaInicioCiclo: normalizarFecha_(fechaInicio),
      SesionesPorCiclo: Number(row[idx.SesionesPorCiclo] || 0),
      CapacidadMaxima: capacidad,
      PlazasOcupadas: ocupadas,
      PlazasLibres: libresReales
    });
  }

  ciclos.sort((a, b) => compararFechas_(a.FechaInicioCiclo, b.FechaInicioCiclo));
  return ciclos;
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
      sheet.getRange(i + 1, idx.EstadoPaciente + 1).setValue(ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO);
      sheet.getRange(i + 1, idx.MotivoEspera + 1).setValue('');
      sheet.getRange(i + 1, idx.CicloObjetivoID + 1).setValue(ciclo.CicloID);
      sheet.getRange(i + 1, idx.CicloActivoID + 1).setValue('');
      sheet.getRange(i + 1, idx.FechaPrimeraSesionReal + 1).setValue(ciclo.FechaInicioCiclo);
      sheet.getRange(i + 1, idx.SesionesPlanificadas + 1).setValue(ciclo.SesionesPorCiclo);
      sheet.getRange(i + 1, idx.SesionesPendientes + 1).setValue(ciclo.SesionesPorCiclo);
      sheet.getRange(i + 1, idx.ProximaSesion + 1).setValue(ciclo.FechaInicioCiclo);
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

  crearAsignacionCiclo_({
    pacienteId: paciente.PacienteID,
    cicloId: ciclo.CicloID,
    modalidad: ciclo.Modalidad,
    estadoAsignacion: ESTADOS_ASIGNACION.RESERVADO,
    observaciones: 'Asignación manual desde espera'
  });

  asegurarSesionesPacienteGrupoSiFaltan_(paciente.PacienteID, ciclo.CicloID);
  recalcularOcupacionCiclosInterno_();

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
    .setHeight(620);

  SpreadsheetApp.getUi().showModalDialog(html, 'Editar paciente');
}

function obtenerPacientesFormularioEdicion() {
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
    const pacienteId = row[idx.PacienteID];
    const nombre = row[idx.Nombre];
    const modalidad = row[idx.ModalidadSolicitada];
    const estado = row[idx.EstadoPaciente];

    if (!pacienteId) continue;

    out.push({
      pacienteId: pacienteId,
      label: (nombre || 'SIN_NOMBRE') + ' | ' + (modalidad || '') + ' | ' + (estado || '')
    });
  }

  out.sort((a, b) => String(a.label).localeCompare(String(b.label)));
  return out;
}

function obtenerDetallePacienteParaEdicion(pacienteId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('No hay pacientes.');
  }

  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (String(row[idx.PacienteID]) === String(pacienteId)) {
      const estado = row[idx.EstadoPaciente];
      const puedeEditarModalidad = estado === ESTADOS_PACIENTE.ESPERA;
      const puedeEditarFechaConsulta = estado === ESTADOS_PACIENTE.ESPERA;

      return {
        pacienteId: row[idx.PacienteID],
        nombre: row[idx.Nombre] || '',
        modalidad: row[idx.ModalidadSolicitada] || '',
        fechaPrimeraConsulta: formatearFechaISOInput_(row[idx.FechaPrimeraConsulta]),
        estadoPaciente: estado || '',
        motivoEspera: row[idx.MotivoEspera] || '',
        cicloObjetivoId: row[idx.CicloObjetivoID] || '',
        cicloActivoId: row[idx.CicloActivoID] || '',
        fechaPrimeraSesionReal: formatearFecha_(row[idx.FechaPrimeraSesionReal]),
        sesionesPlanificadas: Number(row[idx.SesionesPlanificadas] || 0),
        sesionesCompletadas: Number(row[idx.SesionesCompletadas] || 0),
        sesionesPendientes: Number(row[idx.SesionesPendientes] || 0),
        proximaSesion: formatearFecha_(row[idx.ProximaSesion]),
        fechaCierre: formatearFecha_(row[idx.FechaCierre]),
        observaciones: row[idx.Observaciones] || '',
        restricciones: {
          puedeEditarModalidad: puedeEditarModalidad,
          puedeEditarFechaConsulta: puedeEditarFechaConsulta,
          puedeEditarNombre: true,
          puedeEditarObservaciones: true
        }
      };
    }
  }

  throw new Error('Paciente no encontrado.');
}

function guardarEdicionPaciente(formData) {
  const pacienteId = String(formData.pacienteId || '').trim();
  const nombre = String(formData.nombre || '').trim();
  const modalidadNueva = String(formData.modalidad || '').trim();
  const fechaConsultaISO = String(formData.fechaPrimeraConsulta || '').trim();
  const observaciones = String(formData.observaciones || '').trim();

  if (!pacienteId) {
    throw new Error('Falta el paciente.');
  }

  if (!nombre) {
    throw new Error('El nombre es obligatorio.');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('No hay pacientes.');
  }

  const idx = indexByHeader_(data[0]);

  const columnasNecesarias = [
    'PacienteID',
    'Nombre',
    'ModalidadSolicitada',
    'FechaPrimeraConsulta',
    'EstadoPaciente',
    'Observaciones'
  ];

  columnasNecesarias.forEach(col => {
    if (idx[col] === undefined) {
      throw new Error('Falta la columna "' + col + '" en ' + SHEET_PACIENTES + '.');
    }
  });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (String(row[idx.PacienteID]) !== String(pacienteId)) continue;

    const estado = row[idx.EstadoPaciente];
    const modalidadActual = row[idx.ModalidadSolicitada];

    sheet.getRange(i + 1, idx.Nombre + 1).setValue(nombre);
    sheet.getRange(i + 1, idx.Observaciones + 1).setValue(observaciones);

    if (estado === ESTADOS_PACIENTE.ESPERA) {
      if (!Object.values(MODALIDADES).includes(modalidadNueva)) {
        throw new Error('La modalidad no es válida.');
      }

      const fechaConsulta = parseFechaISO_(fechaConsultaISO);
      if (!(fechaConsulta instanceof Date)) {
        throw new Error('La fecha de primera consulta no es válida.');
      }

      sheet.getRange(i + 1, idx.ModalidadSolicitada + 1).setValue(modalidadNueva);
      sheet.getRange(i + 1, idx.FechaPrimeraConsulta + 1).setValue(fechaConsulta);

      return {
        mensaje:
          'Paciente actualizado correctamente.\n\n' +
          'Nombre: ' + nombre + '\n' +
          'Modalidad anterior: ' + modalidadActual + '\n' +
          'Modalidad nueva: ' + modalidadNueva + '\n' +
          'Estado: ' + estado
      };
    }

    if (modalidadNueva && modalidadNueva !== modalidadActual) {
      throw new Error('Solo se puede cambiar la modalidad si el paciente está en ESPERA.');
    }

    if (fechaConsultaISO) {
      const fechaOriginal = row[idx.FechaPrimeraConsulta];
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

    return {
      mensaje:
        'Paciente actualizado correctamente.\n\n' +
        'Nombre: ' + nombre + '\n' +
        'Estado: ' + estado + '\n' +
        'Solo se han permitido los campos editables para ese estado.'
    };
  }

  throw new Error('Paciente no encontrado.');
}

function formatearFechaISOInput_(fecha) {
  if (!(fecha instanceof Date)) return '';
  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd');
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

    recalcularOcupacionCiclosInterno_();
    recalcularMetricasBasicas_();

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
    // rollback mínimo de plaza nueva si algo falla después de reservar
    try {
      actualizarPlazasCiclo_(nuevoCiclo.CicloID, -1, false);
    } catch (e) {
      // evitamos encadenar errores
    }
    throw error;
  }
}

function obtenerPacienteCompletoPorId_(pacienteId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (String(row[idx.PacienteID]) === String(pacienteId)) {
      return {
        fila: i + 1,
        PacienteID: row[idx.PacienteID],
        Nombre: row[idx.Nombre],
        NHC: idx.NHC !== undefined ? row[idx.NHC] : '',
        SexoGenero: idx.SexoGenero !== undefined ? row[idx.SexoGenero] : '',
        MotivoConsultaDiagnostico: idx.MotivoConsultaDiagnostico !== undefined ? row[idx.MotivoConsultaDiagnostico] : '',
        MotivoConsultaOtros: idx.MotivoConsultaOtros !== undefined ? row[idx.MotivoConsultaOtros] : '',
        ModalidadSolicitada: row[idx.ModalidadSolicitada],
        FechaAlta: row[idx.FechaAlta],
        FechaPrimeraConsulta: row[idx.FechaPrimeraConsulta],
        EstadoPaciente: row[idx.EstadoPaciente],
        MotivoEspera: row[idx.MotivoEspera],
        CicloObjetivoID: row[idx.CicloObjetivoID],
        CicloActivoID: row[idx.CicloActivoID],
        FechaPrimeraSesionReal: row[idx.FechaPrimeraSesionReal],
        SesionesPlanificadas: Number(row[idx.SesionesPlanificadas] || 0),
        SesionesCompletadas: Number(row[idx.SesionesCompletadas] || 0),
        SesionesPendientes: Number(row[idx.SesionesPendientes] || 0),
        ProximaSesion: row[idx.ProximaSesion],
        FechaCierre: row[idx.FechaCierre],
        Observaciones: row[idx.Observaciones] || ''
      };
    }
  }

  return null;
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

    sheet.getRange(i + 1, idx.EstadoPaciente + 1).setValue(ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO);
    sheet.getRange(i + 1, idx.MotivoEspera + 1).setValue('');
    sheet.getRange(i + 1, idx.CicloObjetivoID + 1).setValue(ciclo.CicloID);
    sheet.getRange(i + 1, idx.CicloActivoID + 1).setValue('');
    sheet.getRange(i + 1, idx.FechaPrimeraSesionReal + 1).setValue(ciclo.FechaInicioCiclo);
    sheet.getRange(i + 1, idx.SesionesPlanificadas + 1).setValue(ciclo.SesionesPorCiclo);
    sheet.getRange(i + 1, idx.SesionesCompletadas + 1).setValue(0);
    sheet.getRange(i + 1, idx.SesionesPendientes + 1).setValue(ciclo.SesionesPorCiclo);
    sheet.getRange(i + 1, idx.ProximaSesion + 1).setValue(ciclo.FechaInicioCiclo);
    sheet.getRange(i + 1, idx.FechaCierre + 1).setValue('');
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

    if (
      estado === ESTADOS_PACIENTE.ESPERA ||
      estado === ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO
    ) {
      out.push({
        pacienteId: row[idx.PacienteID],
        label:
          (row[idx.Nombre] || 'SIN_NOMBRE') +
          ' | ' + (row[idx.ModalidadSolicitada] || '') +
          ' | ' + (estado || '')
      });
    }
  }

  out.sort((a, b) => String(a.label).localeCompare(String(b.label)));
  return out;
}

function obtenerImpactoEliminacionPacienteFormulario(pacienteId) {
  const paciente = obtenerPacienteCompletoPorId_(pacienteId);
  if (!paciente) {
    throw new Error('Paciente no encontrado.');
  }

  if (
    paciente.EstadoPaciente !== ESTADOS_PACIENTE.ESPERA &&
    paciente.EstadoPaciente !== ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO
  ) {
    throw new Error('Solo se puede eliminar físicamente un paciente en ESPERA o ACTIVO_PENDIENTE_INICIO.');
  }

  const asignaciones = contarAsignacionesPaciente_(paciente.PacienteID);
  const sesiones = contarSesionesPaciente_(paciente.PacienteID);
  const cicloId = String(paciente.CicloObjetivoID || paciente.CicloActivoID || '');

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
    liberaraPlaza: paciente.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO && !!cicloId
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

  if (
    paciente.EstadoPaciente !== ESTADOS_PACIENTE.ESPERA &&
    paciente.EstadoPaciente !== ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO
  ) {
    throw new Error('Solo se puede eliminar físicamente un paciente en ESPERA o ACTIVO_PENDIENTE_INICIO.');
  }

  const cicloId = String(paciente.CicloObjetivoID || paciente.CicloActivoID || '');
  const liberarPlaza = paciente.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO && !!cicloId;

  if (liberarPlaza) {
    actualizarPlazasCiclo_(cicloId, -1, false);
  }

  borrarAsignacionesPaciente_(paciente.PacienteID);
  borrarSesionesPaciente_(paciente.PacienteID);
  borrarPacientePorId_(paciente.PacienteID);

  recalcularOcupacionCiclosInterno_();
  recalcularMetricasBasicas_();

  try {
    refrescarDashboard();
  } catch (e) {
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ASIGNACIONES_CICLO);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const idx = indexByHeader_(data[0]);

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idx.PacienteID] || '') === String(pacienteId)) {
      sheet.deleteRow(i + 1);
    }
  }
}

function borrarSesionesPaciente_(pacienteId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const idx = indexByHeader_(data[0]);

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idx.PacienteID] || '') === String(pacienteId)) {
      sheet.deleteRow(i + 1);
    }
  }
}

function borrarPacientePorId_(pacienteId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('No hay pacientes.');
  }

  const idx = indexByHeader_(data[0]);

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idx.PacienteID] || '') === String(pacienteId)) {
      sheet.deleteRow(i + 1);
      return;
    }
  }

  throw new Error('Paciente no encontrado para borrado.');
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

  if (detalle.cicloActivoId || detalle.cicloObjetivoId) {
    tipoAccion = 'REASIGNAR';
  }

  return {
    ...detalle,
    tipoAccion
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