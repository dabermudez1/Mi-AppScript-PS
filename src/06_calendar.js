/***********************
 * BLOQUE 6
 * GOOGLE CALENDAR
 ***********************/

function crearCalendarioConsulta() {
  const ui = SpreadsheetApp.getUi();

  try {
    const calendar = obtenerOCrearCalendarioConsulta_();
    ui.alert('Calendario listo: ' + calendar.getName());
  } catch (error) {
    ui.alert('Error con Google Calendar: ' + error.message);
    throw error;
  }
}

function sincronizarSesionesAGoogleCalendar(calendarParam) {
  const ui = SpreadsheetApp.getUi();
  const calendar = calendarParam || obtenerOCrearCalendarioConsulta_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);

  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_SESIONES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0];
  const idx = indexByHeader_(headers);

  // Mapa PacienteID -> NHC (para que el operador vea NHC en el evento)
  const sheetPacientes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  let nhcByPacienteId = new Map();
  if (sheetPacientes) {
    const pacData = sheetPacientes.getDataRange().getValues();
    if (pacData && pacData.length > 1) {
      const pacIdx = indexByHeader_(pacData[0]);
      const idCol = pacIdx.PacienteID;
      const nhcCol = pacIdx.NHC;
      if (idCol !== undefined && nhcCol !== undefined) {
        nhcByPacienteId = new Map(
          pacData.slice(1).map(row => [String(row[idCol]), String(row[nhcCol] || '')])
        );
      }
    }
  }

  function esEventoDiaBloqueado_(evento) {
    const titulo = String(evento && typeof evento.getTitle === 'function' ? evento.getTitle() : '');
    const t = titulo.toLowerCase();
    return t.includes('día bloqueado') || t.includes('dia bloqueado');
  }

  // 1) BORRAR eventos que ya no existen en "base de datos"
  // Set con todos los SesionID que existen en la hoja (equivale a "obtenerSesionDesdeBDPorId(...)" en boolean)
  const sesionIdsBD = new Set(
    data.slice(1).map(row => String(row[idx.SesionID]))
  );

  // Rango de fechas basado en la base de datos (acelera bastante si el calendario tiene muchos eventos antiguos)
  let minTime = null;
  let maxTime = null;

  for (let i = 1; i < data.length; i++) {
    const fechaSesion = data[i][idx.FechaSesion];
    const t = normalizarFecha_(new Date(fechaSesion)).getTime();
    if (Number.isNaN(t)) continue;
    if (minTime === null || t < minTime) minTime = t;
    if (maxTime === null || t > maxTime) maxTime = t;
  }

  const eventosGoogle =
    minTime === null
      ? []
      : calendar.getEvents(new Date(minTime), new Date(maxTime));

  let eventosBorrados = 0;
  eventosGoogle.forEach(evento => {
    // No tocamos los eventos de días bloqueados con la sincronización de sesiones.
    if (esEventoDiaBloqueado_(evento)) return;

    const existeSesion = sesionIdsBD.has(String(evento.getId()));
    if (!existeSesion) {
      evento.deleteEvent();
      eventosBorrados++;
      try { Logger.log('Evento eliminado: ' + evento.getTitle()); } catch (e) { /* opcional */ }
    }
  });

  // 2) PREFETCH de eventos para evitar getEvents() por fila
  // Mapa: dia(ms normalizado) => primer evento encontrado en ese día
  const eventosPorFechaKey = new Map();

  if (minTime !== null) {
    const eventosEnRango = calendar.getEvents(new Date(minTime), new Date(maxTime));
    eventosEnRango.forEach(evento => {
      // Ignoramos eventos de "día bloqueado" al buscar el evento existente del día.
      if (esEventoDiaBloqueado_(evento)) return;

      let allDayStart = null;
      if (typeof evento.getAllDayStartDate === 'function') {
        try {
          allDayStart = evento.getAllDayStartDate();
        } catch (e) {
          allDayStart = null;
        }
      }
      const start = allDayStart instanceof Date ? allDayStart : evento.getStartTime();
      const key = normalizarFecha_(start).getTime();
      if (!eventosPorFechaKey.has(key)) {
        eventosPorFechaKey.set(key, evento);
      }
    });
  }

  // 3) Actualizar/crear por fila como evento "1 día entero" (all-day)
  data.slice(1).forEach(row => {
    const pacienteId = row[idx.PacienteID];
    const modalidad = row[idx.Modalidad];
    const fechaSesion = row[idx.FechaSesion];

    const fechaSesionDate = new Date(fechaSesion);
    const fechaDia = normalizarFecha_(fechaSesionDate);
    const diaKey = fechaDia.getTime();

    const color = obtenerColorPorModalidad_(modalidad);
    const sesion = {
      SesionID: row[idx.SesionID],
      PacienteID: pacienteId,
      NombrePaciente: row[idx.NombrePaciente],
      NumeroSesion: row[idx.NumeroSesion],
      Modalidad: modalidad,
      EstadoSesion: row[idx.EstadoSesion],
      FechaSesion: fechaSesionDate,
      FechaOriginal: row[idx.FechaOriginal],
      CicloID: row[idx.CicloID],
      Notas: row[idx.Notas],
      NHC: nhcByPacienteId.get(String(pacienteId)) || ''
    };

    const titulo = construirTituloEventoSesion_(sesion);
    const descripcion = construirDescripcionEventoSesion_(sesion);

    let eventoExistente;
    if (Number.isNaN(diaKey)) {
      // Mantener comportamiento "tal cual" para fechas inválidas (incluida posible respuesta/throw)
      eventoExistente = calendar.getEvents(fechaSesionDate, fechaSesionDate)[0];
    } else {
      eventoExistente = eventosPorFechaKey.get(diaKey);
    }

    if (eventoExistente) {
      const esAllDay =
        (() => {
          if (typeof eventoExistente.getAllDayStartDate !== 'function') return false;
          try {
            return !!eventoExistente.getAllDayStartDate();
          } catch (e) {
            return false;
          }
        })();

      eventoExistente.setTitle(titulo);
      eventoExistente.setDescription(descripcion);

      if (typeof eventoExistente.setColor === 'function') {
        try { eventoExistente.setColor(color); } catch (e) { /* ignorar si no soporta color */ }
      }

      // Cumplir requisito: todo debe quedar como evento all-day.
      if (!esAllDay) {
        eventoExistente.deleteEvent();
        const nuevoEvento = calendar.createAllDayEvent(titulo, fechaDia, { description: descripcion });
        if (typeof nuevoEvento.setColor === 'function') {
          try { nuevoEvento.setColor(color); } catch (e) { /* ignorar */ }
        }
        if (!Number.isNaN(diaKey)) {
          eventosPorFechaKey.set(diaKey, nuevoEvento);
        }
      }
    } else {
      const nuevoEvento = calendar.createAllDayEvent(titulo, fechaDia, { description: descripcion });
      if (typeof nuevoEvento.setColor === 'function') {
        try { nuevoEvento.setColor(color); } catch (e) { /* ignorar */ }
      }

      // Si hay más filas con la misma fecha, mantenemos el comportamiento previo: mapear al primer evento encontrado.
      if (!Number.isNaN(diaKey)) {
        eventosPorFechaKey.set(diaKey, nuevoEvento);
      }
    }
  });

  ui.alert(
    'Sincronización de Google Calendar completada.' +
    (typeof eventosBorrados === 'number' && eventosBorrados > 0 ? ' (eventos eliminados: ' + eventosBorrados + ')' : '')
  );
}


/***************
 * CALENDAR CORE
 ***************/
function obtenerOCrearCalendarioConsulta_() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const savedCalendarId = props.getProperty('CONSULTA_CALENDAR_ID');

  let calendar;
  
  if (savedCalendarId) {
    calendar = CalendarApp.getCalendarById(savedCalendarId);
    if (calendar) {
      return calendar;
    }
  }

  // Si no existe el calendario vinculado, preguntamos el nombre a crear.
  let nombreCalendario = GOOGLE_CALENDAR_NAME;
  try {
    const resp = ui.prompt(
      'Crear calendario de Google Calendar',
      'Indica el nombre del calendario a crear (se usará este si no existe el vinculado):',
      ui.ButtonSet.OK_CANCEL
    );
    if (resp && resp.getSelectedButton && resp.getSelectedButton() === ui.Button.OK) {
      const texto = (resp.getResponseText && resp.getResponseText()) || '';
      const t = String(texto).trim();
      if (t) nombreCalendario = t;
    }
  } catch (e) {
    // Si no se puede mostrar prompt (ej. contexto no interactivo), usamos el valor por defecto.
  }

  // Intentar re-vincular por nombre antes de crear (si ya existiera).
  const porNombre = CalendarApp.getCalendarsByName(nombreCalendario);
  if (porNombre && porNombre.length > 0) {
    calendar = porNombre[0];
  } else {
    ui.alert('El calendario "' + nombreCalendario + '" no existe. Se creará uno nuevo.');
    calendar = CalendarApp.createCalendar(nombreCalendario);
  }

  props.setProperty('CONSULTA_CALENDAR_ID', calendar.getId());
  
  return calendar;
}

/**
 * Devuelve la URL web del calendario vinculado, si existe.
 * No crea nada ni pregunta: se usa para mostrar botones en el panel.
 */
function obtenerCalendarConsultaUrl_() {
  const props = PropertiesService.getUserProperties();
  const savedCalendarId = props.getProperty('CONSULTA_CALENDAR_ID');

  // Intento 1: por ID guardado
  if (savedCalendarId) {
    try {
      const cal = CalendarApp.getCalendarById(savedCalendarId);
      if (cal) {
        return 'https://calendar.google.com/calendar/u/0?cid=' + encodeURIComponent(cal.getId());
      }
    } catch (e) {
      // sigue con fallback
    }
  }

  // Intento 2: por nombre (si el usuario lo creó manualmente)
  try {
    const porNombre = CalendarApp.getCalendarsByName(GOOGLE_CALENDAR_NAME);
    if (porNombre && porNombre.length > 0) {
      return 'https://calendar.google.com/calendar/u/0?cid=' + encodeURIComponent(porNombre[0].getId());
    }
  } catch (e) {
    // sin calendario
  }

  return null;
}


function obtenerEventoSeguro_(calendar, eventId) {
  try {
    return calendar.getEventById(eventId);
  } catch (error) {
    return null;
  }
}

/***************
 * CONSTRUCCIÓN EVENTO
 ***************/
function construirTituloEventoSesion_(sesion) {
  const paciente = sesion.NombrePaciente || 'Paciente';
  const sesionLabel = 'Sesión S' + (sesion.NumeroSesion || '');
  const modalidad = sesion.Modalidad || '';

  const base = [paciente, sesionLabel, modalidad].filter(Boolean).join(' - ');

  if (
    sesion.EstadoSesion === ESTADOS_SESION.COMPLETADA_AUTO ||
    sesion.EstadoSesion === ESTADOS_SESION.COMPLETADA_MANUAL
  ) {
    return '✅ ' + base;
  }

  if (sesion.EstadoSesion === ESTADOS_SESION.CANCELADA) {
    return '❌ ' + base;
  }

  if (sesion.EstadoSesion === ESTADOS_SESION.REPROGRAMADA) {
    return '🔁 ' + base;
  }

  return base;
}

function construirDescripcionEventoSesion_(sesion) {
  const paciente = sesion.NombrePaciente || '';
  const nhc = sesion.NHC || '';
  const modalidad = sesion.Modalidad || '';

  const lineas = [
    'PACIENTE: ' + paciente,
    'NHC: ' + (nhc || '-'),
    'SESION: S' + (sesion.NumeroSesion || ''),
    'MODALIDAD: ' + modalidad,
    'ESTADO: ' + (sesion.EstadoSesion || ''),
    'FECHA (DIA ENTERO): ' + formatearFecha_(sesion.FechaSesion),
    'FECHA ORIGINAL: ' + formatearFecha_(sesion.FechaOriginal),
    'CICLO: ' + (sesion.CicloID || '-'),
    '',
    'NOTAS:',
    String(sesion.Notas || '-')
  ];

  return lineas.join('\n');
}

function generarHashSesionCalendar_(sesion) {
  const str = [
    sesion.SesionID || '',
    sesion.NombrePaciente || '',
    sesion.Modalidad || '',
    sesion.NumeroSesion || '',
    sesion.FechaSesion instanceof Date ? normalizarFecha_(sesion.FechaSesion).getTime() : '',
    sesion.EstadoSesion || '',
    sesion.Notas || '',
    sesion.CicloID || ''
  ].join('|');

  return Utilities.base64Encode(str);
}

function obtenerColorPorModalidad_(modalidad) {
  switch (modalidad) {
    case MODALIDADES.INDIVIDUAL:
      return CalendarApp.EventColor.GREEN;
    case MODALIDADES.GRUPO_1:
      return CalendarApp.EventColor.BLUE;
    case MODALIDADES.GRUPO_2:
      return CalendarApp.EventColor.PURPLE;
    case MODALIDADES.GRUPO_3:
      return CalendarApp.EventColor.ORANGE;
    default:
      return CalendarApp.EventColor.GRAY;
  }
}

/***************
 * ESCRITURA RESULTADO SYNC
 ***************/
function guardarResultadoSyncSesion_(sheet, idx, fila, resultado) {
  sheet.getRange(fila, idx.CalendarEventId + 1).setValue(resultado.eventId || '');
  sheet.getRange(fila, idx.CalendarSyncStatus + 1).setValue(resultado.syncStatus || '');
  sheet.getRange(fila, idx.CalendarLastSync + 1).setValue(new Date());
  sheet.getRange(fila, idx.CalendarEventTitle + 1).setValue(resultado.title || '');
  sheet.getRange(fila, idx.CalendarHash + 1).setValue(resultado.hash || '');
}

/***************
 * FUNCION OCULTA PARA VERIFICAR
 ***************/

function diagnosticarGoogleCalendar() {
  const ui = SpreadsheetApp.getUi();

  try {
    const calendar = obtenerOCrearCalendarioConsulta_();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);

    if (!sheet) {
      throw new Error('No existe la hoja ' + SHEET_SESIONES + '.');
    }

    const data = sheet.getDataRange().getValues();
    const idx = indexByHeader_(data[0]);

    let sesionesConEventId = 0;
    let sesionesCreadas = 0;
    let sesionesActualizadas = 0;
    let sesionesError = 0;

    for (let i = 1; i < data.length; i++) {
      const eventId = idx.CalendarEventId !== undefined ? data[i][idx.CalendarEventId] : '';
      const syncStatus = idx.CalendarSyncStatus !== undefined ? data[i][idx.CalendarSyncStatus] : '';

      if (eventId) sesionesConEventId++;
      if (syncStatus === 'CREADO' || syncStatus === 'RECREADO') sesionesCreadas++;
      if (syncStatus === 'ACTUALIZADO') sesionesActualizadas++;
      if (syncStatus === 'ERROR') sesionesError++;
    }

    const hoy = normalizarFecha_(new Date());
    const haceUnAno = new Date(hoy.getFullYear() - 1, hoy.getMonth(), hoy.getDate());
    const dentroDeDosAnos = new Date(hoy.getFullYear() + 2, hoy.getMonth(), hoy.getDate());
    const eventos = calendar.getEvents(haceUnAno, dentroDeDosAnos);

    ui.alert(
      'Diagnóstico Google Calendar\n\n' +
      'Nombre: ' + calendar.getName() + '\n' +
      'ID: ' + calendar.getId() + '\n' +
      'Eventos en calendario (-1 año / +2 años): ' + eventos.length + '\n\n' +
      'SESIONES con CalendarEventId: ' + sesionesConEventId + '\n' +
      'SESIONES con status CREADO/RECREADO: ' + sesionesCreadas + '\n' +
      'SESIONES con status ACTUALIZADO: ' + sesionesActualizadas + '\n' +
      'SESIONES con status ERROR: ' + sesionesError
    );
  } catch (error) {
    ui.alert('Error en diagnóstico de calendar: ' + error.message);
    throw error;
  }
}

function verCalendarioConsultaActual() {
  const ui = SpreadsheetApp.getUi();

  try {
    const props = PropertiesService.getScriptProperties();
    const savedCalendarId = props.getProperty('CONSULTA_CALENDAR_ID') || '(no guardado)';
    const calendar = obtenerOCrearCalendarioConsulta_();

    ui.alert(
      'Calendario actual\n\n' +
      'Nombre: ' + calendar.getName() + '\n' +
      'ID real: ' + calendar.getId() + '\n' +
      'ID guardado: ' + savedCalendarId
    );
  } catch (error) {
    ui.alert('Error al consultar calendario actual: ' + error.message);
    throw error;
  }
}

function resetCalendarioConsultaVinculado() {
  const ui = SpreadsheetApp.getUi();

  try {
    PropertiesService.getScriptProperties().deleteProperty('CONSULTA_CALENDAR_ID');
    const calendar = obtenerOCrearCalendarioConsulta_();

    ui.alert(
      'Vinculación de calendario reiniciada.\n\n' +
      'Ahora se está usando:\n' +
      calendar.getName() + '\n' +
      calendar.getId()
    );
  } catch (error) {
    ui.alert('Error al resetear la vinculación del calendario: ' + error.message);
    throw error;
  }
}

function limpiarCamposSyncCalendarSesiones() {
  const ui = SpreadsheetApp.getUi();

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
    if (!sheet) {
      throw new Error('No existe la hoja ' + SHEET_SESIONES + '.');
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      ui.alert('No hay sesiones.');
      return;
    }

    const idx = indexByHeader_(data[0]);
    const columnas = [
      'CalendarEventId',
      'CalendarSyncStatus',
      'CalendarLastSync',
      'CalendarEventTitle',
      'CalendarHash'
    ];

    columnas.forEach(col => {
      if (idx[col] === undefined) {
        throw new Error('Falta la columna "' + col + '" en ' + SHEET_SESIONES + '.');
      }
    });

    for (let i = 1; i < data.length; i++) {
      sheet.getRange(i + 1, idx.CalendarEventId + 1).setValue('');
      sheet.getRange(i + 1, idx.CalendarSyncStatus + 1).setValue('');
      sheet.getRange(i + 1, idx.CalendarLastSync + 1).setValue('');
      sheet.getRange(i + 1, idx.CalendarEventTitle + 1).setValue('');
      sheet.getRange(i + 1, idx.CalendarHash + 1).setValue('');
    }

    ui.alert('Campos de sincronización limpiados en SESIONES.');
  } catch (error) {
    ui.alert('Error al limpiar sync de sesiones: ' + error.message);
    throw error;
  }
}

function crearTriggerAutomatico() {
  ScriptApp.newTrigger('sincronizarDiasBloqueadosAGoogleCalendar')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()  // Se activa cada vez que se edita el documento
    .create();
}

function sincronizarCalendario() {
  const calendar = CalendarApp.getCalendarById('your-calendar-id@group.calendar.google.com');
  
  // Este es un ejemplo de la lista de sesiones
  const sesiones = obtenerSesionesParaSincronizar();  // Suponiendo que esta función ya existe y retorna las sesiones que quieres sincronizar.
  
  sesiones.forEach(function(sesion) {
    let colorId;
    let tituloEvento = 'Consulta Psicología';  // Por defecto, puedes cambiarlo si es necesario.

    // Asignar color según la modalidad
    switch (sesion.modalidad) {
      case 'INDIVIDUAL':
        colorId = 8;  // Violeta
        tituloEvento = 'Consulta Psicología Individual';
        break;
      case 'GRUPO_1':
        colorId = 6;  // Verde claro
        tituloEvento = 'Consulta Psicología Grupo 1';
        break;
      case 'GRUPO_2':
        colorId = 9;  // Azul claro
        tituloEvento = 'Consulta Psicología Grupo 2';
        break;
      case 'GRUPO_3':
        colorId = 1;  // Azul
        tituloEvento = 'Consulta Psicología Grupo 3';
        break;
      default:
        colorId = 1;  // Azul por defecto
        break;
    }

    // Crear evento de todo el día
    calendar.createAllDayEvent(tituloEvento, new Date(sesion.fecha), { 
      description: sesion.descripcion || '', 
      colorId: colorId  // Aplicar el color al evento
    });
  });
}

// Legacy/ejemplo: mantenido para referencia, pero RENOMBRADO para que no interfiera
// con la sincronización real que vive en `23_dias_bloqueados_calendar_sync.js`.
function sincronizarDiasBloqueadosAGoogleCalendarLegacy_(calendar) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DIAS_BLOQUEADOS);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_DIAS_BLOQUEADOS + '.');
  }

  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);
  
  const eventosEnCalendar = calendar.getEvents(new Date(), new Date('2025-12-31')); // Obtener eventos del calendario

  for (let i = 1; i < data.length; i++) {
    const fechaBloqueada = data[i][idx.Fecha];
    const motivo = data[i][idx.Motivo];
    
    if (!fechaBloqueada || !motivo) continue;

    const eventoExistente = eventosEnCalendar.find(evento => evento.getTitle().includes(motivo));
    
    if (!eventoExistente) {
      // Crear evento si no existe
      calendar.createAllDayEvent(
        'Día bloqueado: ' + motivo,
        new Date(fechaBloqueada),
        { description: 'Motivo: ' + motivo }
      );
    } else {
      // Eliminar evento si está presente pero ya no es relevante
      eventoExistente.deleteEvent();
    }
  }
}

function eliminarEventosAntiguos(calendar, eventosEnCalendar) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);

  // Obtener todos los eventos existentes en Google Calendar
  const eventIdsEnBaseDeDatos = data.slice(1).map(row => row[idx.CalendarEventId]);

  // Eliminar eventos en el calendario que no están en la base de datos
  eventosEnCalendar.forEach(event => {
    if (!eventIdsEnBaseDeDatos.includes(event.getId())) {
      event.deleteEvent();
    }
  });
}

function crearTriggerAutomatico() {
  ScriptApp.newTrigger('sincronizarCalendario')
    .timeBased()
    .everyHours(1)  // Sincronización cada hora
    .create();
}

function eliminarEventosNoRepresentados(calendar) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);

  const eventosEnBaseDeDatos = data.slice(1).map(row => row[idx.CalendarEventId]);

  // Rango de fechas para los eventos
  const hoy = new Date();
  const dentroDeDosAnos = new Date(hoy.getFullYear() + 2, hoy.getMonth(), hoy.getDate());

  // Obtener eventos en el calendario en el rango de fechas (últimos 2 años por ejemplo)
  const eventosEnCalendar = calendar.getEvents(hoy, dentroDeDosAnos);

  eventosEnCalendar.forEach(event => {
    // Si el evento no está en la base de datos, lo eliminamos
    if (!eventosEnBaseDeDatos.includes(event.getId())) {
      event.deleteEvent();
    }
  });
}
