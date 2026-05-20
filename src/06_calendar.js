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
  const props = PropertiesService.getUserProperties();
  // Si el usuario le da al botón, forzamos inicio limpio
  props.setProperty('TASK_SYNC_CALENDAR_RUNNING', 'true');
  props.setProperty('TASK_SYNC_CALENDAR_PROGRESS', '0');
  props.deleteProperty('TASK_SYNC_CALENDAR_INDEX');
  props.deleteProperty('TASK_SYNC_CALENDAR_MODIFIED');
  
  limpiarTriggersSyncCalendar_();
  
  // Iniciamos el proceso por lotes
  continuarSincronizacionCalendarLote();
}

function limpiarTriggersSyncCalendar_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'continuarSincronizacionCalendarLote') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function continuarSincronizacionCalendarLote() {
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 4 * 60 * 1000; // Límite de seguridad: 4 minutos (Google corta a los 6m)
  
  const props = PropertiesService.getUserProperties();
  const calendar = obtenerOCrearCalendarioConsulta_();
  const sessionRepo = new SessionRepository();
  const patientRepo = new PatientRepository();

  const startIndex = parseInt(props.getProperty('TASK_SYNC_CALENDAR_INDEX') || '0', 10);
  let totalActualizadosGlobales = parseInt(props.getProperty('TASK_SYNC_CALENDAR_MODIFIED') || '0', 10);

  // 1) Cargar NHCs (es muy rápido en memoria)
  const pacientes = patientRepo.findAll();
  const nhcByPacienteId = new Map(pacientes.map(p => [String(p.PacienteID), String(p.NHC || '')]));

  // 2) Obtener TODAS las sesiones
  const todasLasSesiones = sessionRepo.findAll();
  const sesionesDeseadas = todasLasSesiones.map(s => ({
    ...s,
    NHC: nhcByPacienteId.get(String(s.PacienteID)) || ''
  }));
  
  if (sesionesDeseadas.length === 0) {
    finalizarSyncCalendar_(props, 0, 0);
    return;
  }

  let actualizadosEnLote = 0;
  const sesionesModificadas = [];
  let timeoutAlcanzado = false;
  let currentIndex = startIndex;

  for (; currentIndex < sesionesDeseadas.length; currentIndex++) {
    const sesion = sesionesDeseadas[currentIndex];
    const currentHash = generarHashSesionCalendar_(sesion);

    // OPTIMIZACIÓN CLAVE: Si el hash coincide, ya tiene ID y no está en ERROR, saltamos.
    // Esto ahorra el 99% de las llamadas a la API de Calendar.
    if (currentHash === sesion.CalendarHash && 
        sesion.CalendarEventId && 
        sesion.CalendarSyncStatus !== 'ERROR') {
      continue;
    }

    try {
      const titulo = construirTituloEventoSesion_(sesion);
      const desc = construirDescripcionEventoSesion_(sesion);
      const color = obtenerColorPorModalidad_(sesion.Modalidad);
      const fecha = normalizarFecha_(sesion.FechaSesion);
      const horaInicioStr = sesion.HoraInicio;
      const duracion = Number(sesion.Duracion || 30);

      let start = null;
      let end = null;
      if (horaInicioStr) {
        start = normalizarFechaHora_(sesion.FechaSesion, horaInicioStr);
        end = sumarMinutos_(start, duracion);
      }

      let ev = sesion.CalendarEventId ? obtenerEventoSeguro_(calendar, sesion.CalendarEventId) : null;

      if (ev) {
        if (start && end) ev.setTime(start, end); else ev.setAllDayDate(fecha);
        ev.setTitle(titulo);
        ev.setDescription(desc);
        try { ev.setColor(color); } catch(e){}
        sesion.CalendarSyncStatus = 'ACTUALIZADO';
      } else {
        const nuevoEv = start && end
          ? calendar.createEvent(titulo, start, end, { description: desc })
          : calendar.createAllDayEvent(titulo, fecha, { description: desc });
        try { nuevoEv.setColor(color); } catch(e){}
        sesion.CalendarEventId = nuevoEv.getId();
        sesion.CalendarSyncStatus = 'CREADO';
      }

      sesion.CalendarHash = currentHash;
      sesion.CalendarLastSync = new Date();
      
      sesionesModificadas.push(sesion);
      actualizadosEnLote++;
      totalActualizadosGlobales++;

    } catch (error) {
      console.error(`Error sync sesion ${sesion.SesionID}: ${error.message}`);
      sesion.CalendarSyncStatus = 'ERROR';
      sesion.CalendarLastSync = new Date();
      sesionesModificadas.push(sesion);
    }

    if (currentIndex % 10 === 0) {
      const prg = Math.max(1, Math.round((currentIndex / sesionesDeseadas.length) * 100));
      props.setProperty('TASK_SYNC_CALENDAR_PROGRESS', prg.toString());
    }

    // Evaluar timeout
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
      timeoutAlcanzado = true;
      currentIndex++; // Para que el trigger empiece en la siguiente sesión
      break;
    }
  }

  // Volcar a Sheets solo las que hemos modificado para ahorrar tiempo
  if (sesionesModificadas.length > 0) {
    sessionRepo.saveAll(sesionesModificadas);
  }

  if (timeoutAlcanzado) {
    // Pausar y programar siguiente lote
    props.setProperty('TASK_SYNC_CALENDAR_INDEX', currentIndex.toString());
    props.setProperty('TASK_SYNC_CALENDAR_MODIFIED', totalActualizadosGlobales.toString());
    
    ScriptApp.newTrigger('continuarSincronizacionCalendarLote')
      .timeBased()
      .after(1000) // Reanudar lo antes posible (1 segundo)
      .create();
  } else {
    // Fin de la iteración. Procedemos a limpiar huérfanos.
    props.setProperty('TASK_SYNC_CALENDAR_PROGRESS', '95');
    
    let eliminados = 0;
    try {
      if (typeof limpiarEventosHuerfanosCalendar === 'function') {
        // Buscamos 60 días atrás y 180 adelante como zona de seguridad
        eliminados = limpiarEventosHuerfanosCalendar(60, 180);
      }
    } catch(e) {
       console.warn("No se pudo limpiar huérfanos: " + e.message);
    }

    finalizarSyncCalendar_(props, totalActualizadosGlobales, eliminados);
  }
}

function finalizarSyncCalendar_(props, actualizados, eliminados) {
  limpiarTriggersSyncCalendar_();
  const resultMsg = `Sincronización completada: ${actualizados} actualizados, ${eliminados} eliminados.`;
  props.setProperty('TASK_SYNC_CALENDAR_PROGRESS', '100');
  props.setProperty('TASK_SYNC_CALENDAR_RESULT', resultMsg);
  props.setProperty('TASK_SYNC_CALENDAR_RUNNING', 'false');
  props.deleteProperty('TASK_SYNC_CALENDAR_INDEX');
  props.deleteProperty('TASK_SYNC_CALENDAR_MODIFIED');
}

/***************
 * CALENDAR CORE
 ***************/
function obtenerOCrearCalendarioConsulta_() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getUserProperties();
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

  if (sesion.EstadoSesion === ESTADOS_SESION.RESERVADA_PROVISIONAL) {
    return '📝 RESERVA 2.1 - ' + (sesion.NombrePaciente || 'Sin Nombre') + ' - ' + (sesion.HoraInicio || '');
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
    'HORA: ' + (sesion.HoraInicio || '-'), // Nuevo
    'DURACION: ' + (sesion.Duracion || '30') + ' min',
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
    sesion.NHC || '', // Ahora sí incluimos NHC para detectar cambios de paciente
    sesion.Modalidad || '',
    sesion.NumeroSesion || '',
    sesion.HoraInicio || '', // Nuevo
    sesion.Duracion || '',
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
    const props = PropertiesService.getUserProperties();
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
    PropertiesService.getUserProperties().deleteProperty('CONSULTA_CALENDAR_ID');
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

// Eliminadas funciones duplicadas y legacy

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

/**
 * Helper para identificar si un evento de Google Calendar 
 * es un "Día Bloqueado" para no sobreescribirlo ni borrarlo
 * durante la sync de sesiones.
 */
function esEventoDiaBloqueado_(evento) {
  const titulo = String(evento && typeof evento.getTitle === 'function' ? evento.getTitle() : '');
  return titulo.toLowerCase().includes('bloqueado');
}
