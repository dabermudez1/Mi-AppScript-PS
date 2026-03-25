/***********************
 * BLOQUE 23
 * SINCRONIZACIÓN DE DÍAS BLOQUEADOS CON GOOGLE CALENDAR
 ***********************/

function obtenerDiasBloqueados() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DIAS_BLOQUEADOS');
  if (!sheet) {
    throw new Error('La hoja DIAS_BLOQUEADOS no existe.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('No hay datos en la hoja DIAS_BLOQUEADOS.');
  }

  const diasBloqueados = [];
  const headers = data[0];
  const idx = {
    Fecha: headers.indexOf('Fecha'),
    Bloqueado: headers.indexOf('Bloqueado'),
    Motivo: headers.indexOf('Motivo')
  };

  if (idx.Fecha === -1 || idx.Bloqueado === -1 || idx.Motivo === -1) {
    throw new Error('Faltan las columnas Fecha, Bloqueado o Motivo en DIAS_BLOQUEADOS.');
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idx.Bloqueado] === true) {
      diasBloqueados.push({
        fecha: new Date(row[idx.Fecha]),
        motivo: row[idx.Motivo] || 'Sin motivo'
      });
    }
  }

  return diasBloqueados;
}

function sincronizarDiasBloqueadosAGoogleCalendar() {
  const calendar = obtenerOCrearCalendarioConsulta_(); // Obtener o crear el calendario de Google
  const diasBloqueados = obtenerDiasBloqueados(); // Obtener los días bloqueados desde Sheets
  if (!diasBloqueados || diasBloqueados.length === 0) return;

  // Normalizamos para comparar por día (ignora la hora) y buscamos el evento por (día + motivo)
  function normalizarDiaKey_(fecha) {
    // normalizarFecha_ vive en 01_base.js
    return normalizarFecha_(new Date(fecha)).getTime();
  }

  function obtenerMotivoDesdeTitulo_(titulo) {
    const t = String(titulo || '');
    // Normaliza para soportar posibles diferencias de acentos
    const tn = t
      .toLowerCase()
      .replace(/á/g, 'a')
      .replace(/é/g, 'e')
      .replace(/í/g, 'i')
      .replace(/ó/g, 'o')
      .replace(/ú/g, 'u');

    const prefijo1 = 'dia bloqueado:';
    const prefijo2 = 'd\u00eda bloqueado:'; // fallback por si ya venía con carácter

    let idx = tn.indexOf(prefijo1);
    if (idx !== 0) {
      idx = tn.indexOf(prefijo2);
    }
    if (idx !== 0) return null;

    const motivo = t.slice(t.indexOf(':') + 1).trim();
    return motivo || null;
  }

  const desired = new Map(); // key => { fecha: Date, motivo: string }
  let minTime = null;
  let maxTime = null;

  diasBloqueados.forEach(dia => {
    const motivo = String(dia.motivo || 'Sin motivo');
    const fechaDia = normalizarFecha_(new Date(dia.fecha));
    const key = normalizarDiaKey_(fechaDia) + '|' + motivo;
    desired.set(key, { fecha: fechaDia, motivo });

    const tk = fechaDia.getTime();
    if (minTime === null || tk < minTime) minTime = tk;
    if (maxTime === null || tk > maxTime) maxTime = tk;
  });

  // Prefetch de eventos del rango para evitar getEventsForDay() por cada día
  const rangoInicio = new Date(minTime);
  const rangoFin = new Date(maxTime);
  const eventosEnRango = calendar.getEvents(rangoInicio, rangoFin);

  const existing = new Map(); // key => CalendarEvent

  eventosEnRango.forEach(evento => {
    const titulo = evento.getTitle();
    const motivo = obtenerMotivoDesdeTitulo_(titulo);
    if (!motivo) return;

    // Para all-day, getAllDayStartDate suele existir; si no, usamos startTime.
    let start = null;
    if (typeof evento.getAllDayStartDate === 'function') {
      try {
        start = evento.getAllDayStartDate();
      } catch (e) {
        start = null;
      }
    }
    if (!(start instanceof Date)) {
      start = evento.getStartTime();
    }

    const key = normalizarDiaKey_(start) + '|' + motivo;
    if (!existing.has(key)) {
      existing.set(key, evento);
    } else {
      // Si hay eventos duplicados para el mismo (día + motivo), nos quedamos con el primero
      // y borramos los demás para evitar acumulación.
      try { evento.deleteEvent(); } catch (e) { /* ignore */ }
    }
  });

  // Borra eventos existentes que ya no están en la base de datos
  existing.forEach((evento, key) => {
    if (!desired.has(key)) {
      evento.deleteEvent();
    }
  });

  // Crea/actualiza los eventos deseados
  desired.forEach((d, key) => {
    const eventoExistente = existing.get(key);

    const titulo = `Día Bloqueado: ${d.motivo}`;
    const descripcion = `Este día está bloqueado por el motivo: ${d.motivo}`;

    if (eventoExistente) {
      eventoExistente.setTitle(titulo);
      eventoExistente.setDescription(descripcion);
      if (typeof eventoExistente.setColor === 'function') {
        try { eventoExistente.setColor(8); } catch (e) { /* ignore */ }
      }
    } else {
      const nuevoEvento = calendar.createAllDayEvent(titulo, d.fecha, { description: descripcion });
      if (typeof nuevoEvento.setColor === 'function') {
        try { nuevoEvento.setColor(8); } catch (e) { /* ignore */ }
      }
    }
  });

  Logger.log('Sincronización de días bloqueados completada.');
}