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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG_MODALIDADES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_CONFIG_MODALIDADES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('La hoja CONFIG_MODALIDADES no tiene datos.');
  }

  const headers = data[0];
  const idx = indexByHeader_(headers);

  const columnasNecesarias = [
    'Modalidad',
    'TipoModalidad',
    'Activa',
    'DiaSemana',
    'FrecuenciaDias',
    'FechaBase',
    'CapacidadMaxima',
    'SesionesPorCiclo'
  ];

  columnasNecesarias.forEach(col => {
    if (idx[col] === undefined) {
      throw new Error('Falta la columna "' + col + '" en ' + SHEET_CONFIG_MODALIDADES + '.');
    }
  });

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.Modalidad]) === String(modalidad)) {
      return {
        Modalidad: data[i][idx.Modalidad],
        TipoModalidad: data[i][idx.TipoModalidad],
        Activa: data[i][idx.Activa] === true,
        DiaSemana: data[i][idx.DiaSemana],
        FrecuenciaDias: Number(data[i][idx.FrecuenciaDias] || 0),
        FechaBase: data[i][idx.FechaBase],
        CapacidadMaxima: Number(data[i][idx.CapacidadMaxima] || 0),
        SesionesPorCiclo: Number(data[i][idx.SesionesPorCiclo] || 0)
      };
    }
  }

  throw new Error('No existe configuración para la modalidad ' + modalidad + '.');
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

function generarFechasCiclo_({ fechaInicio, diaSemana, frecuenciaDias, sesiones }) {
  return generarFechasGrupoPorSemanasConAvisos_({
    fechaInicio: normalizarFecha_(fechaInicio),
    diaSemana: diaSemana,
    intervaloSemanas: Number(frecuenciaDias || 0),
    sesiones: Number(sesiones || 0)
  });
}

function crearCicloEnSheet_({ modalidad, fechaInicio, fechas, config }) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CICLOS);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_CICLOS + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 1) {
    throw new Error('La hoja CICLOS no tiene encabezados.');
  }

  const headers = data[0];
  const idx = indexByHeader_(headers);

  const numeroCiclo = obtenerSiguienteNumeroCiclo_(modalidad, data, idx);
  const cicloId = generarId_('CIC');
  const fechaFin = fechas[fechas.length - 1];

  const row = [
    cicloId,
    modalidad,
    numeroCiclo,
    ESTADOS_CICLO.PLANIFICADO,
    normalizarFecha_(fechaInicio),
    normalizarFecha_(fechaFin),
    normalizarFecha_(config.FechaBase),
    config.DiaSemana,
    config.FrecuenciaDias,
    config.SesionesPorCiclo,
    config.CapacidadMaxima,
    0,
    config.CapacidadMaxima,
    true,
    ''
  ];

  sheet.appendRow(row);

  return {
    cicloId,
    numeroCiclo
  };
}

function obtenerSiguienteNumeroCiclo_(modalidad, data, idx) {
  let maximo = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idx.Modalidad] === modalidad) {
      const numero = Number(row[idx.NumeroCiclo] || 0);
      if (numero > maximo) {
        maximo = numero;
      }
    }
  }

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

  const resultadoFechas = generarFechasCiclo_({
  fechaInicio,
  diaSemana: config.DiaSemana,
  frecuenciaDias: config.FrecuenciaDias,
  sesiones: config.SesionesPorCiclo
  });

const fechas = resultadoFechas.fechas;
const avisos = resultadoFechas.avisos || [];

  const ciclo = crearCicloEnSheet_({
    modalidad,
    fechaInicio,
    fechas,
    config
  });

  let mensaje =
  'Ciclo creado correctamente.\n\n' +
  'Modalidad: ' + modalidad + '\n' +
  'Inicio: ' + formatearFecha_(fechaInicio) + '\n' +
  'Fin: ' + formatearFecha_(fechaFin);

  if (avisos.length > 0) {
    mensaje += '\n\nAvisos:\n- ' + avisos.join('\n- ');
  }

  return {
    cicloId,
    mensaje: mensaje
  };
}