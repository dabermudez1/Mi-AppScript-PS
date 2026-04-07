/***********************
 * BLOQUE 15
 * CALENDARIO OPERATIVO
 ***********************/

function obtenerNumeroDiaSemana_(diaTexto) {
  const mapa = {
    'DOMINGO': 0,
    'LUNES': 1,
    'MARTES': 2,
    'MIERCOLES': 3,
    'JUEVES': 4,
    'VIERNES': 5,
    'SABADO': 6
  };

  const clave = String(diaTexto || '')
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');

  return mapa[clave];
}

function fechaCoincideConDiaSemana_(fecha, diaSemanaTexto) {
  if (!(fecha instanceof Date)) return false;

  const esperado = obtenerNumeroDiaSemana_(diaSemanaTexto);
  if (esperado === undefined) return false;

  return normalizarFecha_(fecha).getDay() === esperado;
}

function validarFechaBaseGrupo_(fechaBase, diaSemanaTexto) {
  if (!(fechaBase instanceof Date)) {
    throw new Error('La FechaBase no es válida.');
  }

  if (!fechaCoincideConDiaSemana_(fechaBase, diaSemanaTexto)) {
    throw new Error(
      'La FechaBase no coincide con el día configurado para la modalidad.\n\n' +
      'FechaBase: ' + formatearFecha_(fechaBase) + '\n' +
      'DiaSemana esperado: ' + diaSemanaTexto
    );
  }

  if (esFinDeSemana_(fechaBase)) {
    throw new Error('La FechaBase no puede caer en sábado o domingo.');
  }
}

function sumarSemanasManteniendoDia_(fecha, semanas) {
  const f = normalizarFecha_(fecha);
  f.setDate(f.getDate() + (Number(semanas || 0) * 7));
  return normalizarFecha_(f);
}

function buscarSiguienteFechaValidaIndividual_(fecha) {
  return moverASiguienteLaborable_(fecha);
}

function buscarSiguienteFechaValidaGrupo_(fecha, diaSemanaTexto) {
  let f = normalizarFecha_(fecha);

  if (!fechaCoincideConDiaSemana_(f, diaSemanaTexto)) {
    throw new Error(
      'La fecha calculada no respeta el día fijo del grupo.\n\n' +
      'Fecha: ' + formatearFecha_(f) + '\n' +
      'DiaSemana esperado: ' + diaSemanaTexto
    );
  }

  if (esFinDeSemana_(f)) {
    throw new Error(
      'La fecha del grupo cae en fin de semana, lo cual no es válido.\n\n' +
      'Fecha: ' + formatearFecha_(f)
    );
  }

  return f;
}

function generarFechasGrupoPorSemanas_({
  fechaInicio,
  diaSemana,
  intervaloSemanas,
  sesiones
}) {
  validarFechaBaseGrupo_(fechaInicio, diaSemana);

  const fechas = [];
  let actual = normalizarFecha_(fechaInicio);

  for (let i = 0; i < sesiones; i++) {
    fechas.push(new Date(actual));

    // Avanzar a la siguiente fecha según el intervalo
    actual = sumarSemanasManteniendoDia_(actual, intervaloSemanas);
  }

  return fechas;
}

function generarFechasIndividualPorDiasLaborables_({
  fechaInicio,
  intervaloDias,
  sesiones
}) {
  const fechas = [];

  for (let i = 0; i < sesiones; i++) {
    const base = sumarDiasNaturales_(fechaInicio, i * intervaloDias);
    const ajustada = buscarSiguienteFechaValidaIndividual_(base);
    fechas.push(ajustada);
  }

  return fechas;
}

function esFechaBloqueada_(fecha) {
  const f = normalizarFecha_(fecha);

  if (esFinDeSemana_(f)) return true;

  // Usar la versión con CacheService para evitar lecturas repetidas a la hoja
  const bloqueadas = obtenerMapaDiasBloqueados_();
  const clave = obtenerClaveFecha_(f);

  return bloqueadas[clave] === true;
}

function esFechaOperativaValida_(fecha) {
  if (!(fecha instanceof Date)) return false;
  return !esFechaBloqueada_(fecha);
}

function ajustarASiguienteFechaOperativa_(fecha) {
  let f = normalizarFecha_(fecha);

  while (!esFechaOperativaValida_(f)) {
    f.setDate(f.getDate() + 1);
  }

  return f;
}

function obtenerDiasBloqueados_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DIAS_BLOQUEADOS');
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  const idx = indexByHeader_(data[0]);
  const bloqueadas = {};

  for (let i = 1; i < data.length; i++) {
    const fechaRaw = data[i][idx.Fecha];
    const bloqueadoRaw = data[i][idx.Bloqueado];

    if (!fechaRaw) continue;
    if (!esValorVerdadero_(bloqueadoRaw)) continue;

    const clave = obtenerClaveFecha_(fechaRaw);
    bloqueadas[clave] = true;
  }

  return bloqueadas;
}

function obtenerClaveFecha_(fecha) {
  const f = normalizarFecha_(fecha);
  return Utilities.formatDate(f, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function esValorVerdadero_(valor) {
  if (valor === true) return true;
  if (valor === 1) return true;

  const texto = String(valor || '').trim().toUpperCase();
  return texto === 'TRUE' || texto === 'SI' || texto === 'SÍ' || texto === 'YES' || texto === '1';
}

function obtenerDetalleBloqueoFecha_(fecha) {
  return obtenerDetalleBloqueoFechaConMapa_(fecha, obtenerMapaDiasBloqueados_());
}

function obtenerDetalleBloqueoFechaConMapa_(fecha, mapaDiasBloqueados) {
  const f = normalizarFecha_(fecha);

  if (esFinDeSemana_(f)) {
    return {
      bloqueada: true,
      tipo: 'FIN_DE_SEMANA',
      motivo: 'Sábado o domingo'
    };
  }

  const clave = obtenerClaveFecha_(f);
  const detalle = (mapaDiasBloqueados && mapaDiasBloqueados[clave]) || null;

  if (detalle) {
    return {
      bloqueada: true,
      tipo: detalle.tipo || 'DIA_BLOQUEADO',
      motivo: detalle.motivo || ''
    };
  }

  return {
    bloqueada: false,
    tipo: '',
    motivo: ''
  };
}

function construirMensajeFechaNoOperativa_(fecha) {
  const detalle = obtenerDetalleBloqueoFecha_(fecha);

  if (!detalle.bloqueada) {
    return 'No se puede programar una sesión en una fecha no operativa.';
  }

  if (detalle.tipo === 'FIN_DE_SEMANA') {
    return 'No se puede programar una sesión en sábado o domingo.';
  }

  if (detalle.tipo === 'DIA_BLOQUEADO') {
    if (detalle.motivo) {
      return 'No se puede programar una sesión en una fecha bloqueada.\n\nMotivo: ' + detalle.motivo;
    }

    return 'No se puede programar una sesión en una fecha bloqueada.';
  }

  return 'No se puede programar una sesión en una fecha no operativa.';
}

function ajustarASiguienteFechaOperativaConAviso_(fecha) {
  const original = normalizarFecha_(fecha);
  const detalleOriginal = obtenerDetalleBloqueoFecha_(original);

  if (!detalleOriginal.bloqueada) {
    return {
      fecha: original,
      ajustada: false,
      aviso: ''
    };
  }

  const ajustada = ajustarASiguienteFechaOperativa_(original);
  const detalle = obtenerDetalleBloqueoFecha_(original);

  let motivo = 'fecha no operativa';
  if (detalle.tipo === 'FIN_DE_SEMANA') {
    motivo = 'fin de semana';
  } else if (detalle.tipo === 'DIA_BLOQUEADO' && detalle.motivo) {
    motivo = 'día bloqueado (' + detalle.motivo + ')';
  } else if (detalle.tipo === 'DIA_BLOQUEADO') {
    motivo = 'día bloqueado';
  }

  return {
    fecha: ajustada,
    ajustada: true,
    aviso:
      'Se ajustó automáticamente la fecha de ' +
      formatearFecha_(original) +
      ' a ' +
      formatearFecha_(ajustada) +
      ' por ' + motivo + '.'
  };
}

function ajustarAGrupoSiguienteFechaValidaConAviso_(fecha, diaSemanaTexto, intervaloSemanas) {
  const original = normalizarFecha_(fecha);

  if (!fechaCoincideConDiaSemana_(original, diaSemanaTexto)) {
    throw new Error(
      'La fecha del grupo no respeta el día fijo configurado.\n\n' +
      'Fecha: ' + formatearFecha_(original) + '\n' +
      'Día esperado: ' + diaSemanaTexto
    );
  }

  const detalleOriginal = obtenerDetalleBloqueoFecha_(original);

  if (!detalleOriginal.bloqueada) {
    return {
      fecha: original,
      ajustada: false,
      aviso: ''
    };
  }

  let ajustada = normalizarFecha_(original);
  let intentos = 0;

  while (esFechaBloqueada_(ajustada)) {
    ajustada = sumarSemanasManteniendoDia_(ajustada, intervaloSemanas);
    intentos++;

    if (intentos > 100) {
      throw new Error(
        'No se encontró una fecha válida para el grupo después de múltiples intentos.\n\n' +
        'Fecha inicial: ' + formatearFecha_(original) + '\n' +
        'Día del grupo: ' + diaSemanaTexto
      );
    }
  }

  let motivo = 'fecha no operativa';
  if (detalleOriginal.tipo === 'FIN_DE_SEMANA') {
    motivo = 'fin de semana';
  } else if (detalleOriginal.tipo === 'DIA_BLOQUEADO' && detalleOriginal.motivo) {
    motivo = 'día bloqueado (' + detalleOriginal.motivo + ')';
  } else if (detalleOriginal.tipo === 'DIA_BLOQUEADO') {
    motivo = 'día bloqueado';
  }

  return {
    fecha: ajustada,
    ajustada: true,
    aviso:
      'Se ajustó automáticamente la fecha del grupo de ' +
      formatearFecha_(original) +
      ' a ' +
      formatearFecha_(ajustada) +
      ' por ' + motivo + '.'
  };
}

function generarFechasGrupoPorSemanasConAvisos_({
  fechaInicio,
  diaSemana,
  intervaloSemanas,
  sesiones
}) {
  validarFechaBaseGrupo_(fechaInicio, diaSemana);

  const fechas = [];
  const avisos = [];
  let actual = normalizarFecha_(fechaInicio);

  for (let i = 0; i < sesiones; i++) {
    const ajuste = ajustarAGrupoSiguienteFechaValidaConAviso_(
      actual,
      diaSemana,
      intervaloSemanas
    );

    fechas.push(ajuste.fecha);

    if (ajuste.ajustada) {
      avisos.push('Sesión ' + (i + 1) + ': ' + ajuste.aviso);
    }

    actual = sumarSemanasManteniendoDia_(ajuste.fecha, intervaloSemanas);
  }

  return {
    fechas,
    avisos
  };
}

function obtenerMapaDiasBloqueados_() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'mapaDiasBloqueados_cache';
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* ignore */ }
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DIAS_BLOQUEADOS');
  const mapa = {};

  if (!sheet) return mapa;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return mapa;

  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    const fechaRaw = data[i][idx.Fecha];
    const bloqueadoRaw = data[i][idx.Bloqueado];
    const motivoRaw = data[i][idx.Motivo];

    if (!fechaRaw) continue;
    if (!esValorVerdadero_(bloqueadoRaw)) continue;

    const clave = obtenerClaveFecha_(fechaRaw);

    mapa[clave] = {
      bloqueada: true,
      tipo: 'DIA_BLOQUEADO',
      motivo: motivoRaw || ''
    };
  }

  // 60s: Home y formularios pueden consultarlo muchas veces seguidas.
  try { cache.put(cacheKey, JSON.stringify(mapa), 60); } catch (e) { /* ignore */ }
  return mapa;
}