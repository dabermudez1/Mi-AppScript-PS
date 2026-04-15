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

function esFechaBloqueada_(fecha) {
  const f = normalizarFecha_(fecha);

  if (esFinDeSemana_(f)) return true;

  // Usar la versión con CacheService para evitar lecturas repetidas a la hoja
  const bloqueadas = obtenerMapaDiasBloqueados_();
  const clave = obtenerClaveFecha_(f);

  // El mapa contiene objetos para los días bloqueados, por lo que verificamos si existe la clave
  return !!bloqueadas[clave];
}

/**
 * Obtiene el detalle de por qué una fecha está bloqueada usando un mapa precargado.
 * @param {Date} fecha - La fecha a consultar.
 * @param {Object} mapaBloqueos - El mapa devuelto por obtenerMapaDiasBloqueados_.
 * @returns {Object|null} Objeto con {tipo, motivo} o null si no está bloqueada.
 */
function obtenerDetalleBloqueoFechaConMapa_(fecha, mapaBloqueos) {
  const f = normalizarFecha_(fecha);
  
  if (esFinDeSemana_(f)) {
    return {
      bloqueada: true,
      tipo: 'FIN_DE_SEMANA',
      motivo: 'Sábado o Domingo'
    };
  }

  const clave = obtenerClaveFecha_(f);
  const bloqueoEspecifico = mapaBloqueos[clave];
  
  return bloqueoEspecifico || {
    bloqueada: false,
    tipo: 'OPERATIVO',
    motivo: 'Día operativo'
  };
}
function esFechaOperativaValida_(fecha) {
  return !esFechaBloqueada_(fecha);
}

function construirMensajeFechaNoOperativa_(fecha) {
  const f = normalizarFecha_(fecha);
  const bloqueos = obtenerMapaDiasBloqueados_();
  const detalle = obtenerDetalleBloqueoFechaConMapa_(f, bloqueos);
  
  if (detalle) {
    return `La fecha ${formatearFecha_(f)} no es operativa.\nMotivo: ${detalle.motivo || detalle.tipo}`;
  }
  return `La fecha ${formatearFecha_(f)} no es válida para la programación.`;
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
  // Asegurarse de que 'Bloqueado' y 'Motivo' existan en el índice
  if (idx.Bloqueado === undefined || idx.Motivo === undefined) {
    Logger.log('Advertencia: Las columnas "Bloqueado" o "Motivo" no se encontraron en la hoja DIAS_BLOQUEADOS.');
    return mapa;
  }
  for (let i = 1; i < data.length; i++) {
    const fechaRaw = data[i][idx.Fecha];
    const bloqueadoRaw = data[i][idx.Bloqueado];
    const motivoRaw = data[i][idx.Motivo];

    if (!fechaRaw) continue;
    // Usar esValorVerdadero_ para interpretar correctamente el checkbox o texto
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