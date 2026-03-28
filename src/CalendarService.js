/**
 * CalendarService
 * Gestiona la integración bidireccional con Google Calendar.
 */
class CalendarService {
  /**
   * Obtiene o crea el calendario de la consulta.
   */
  static getCalendar() {
    const props = PropertiesService.getUserProperties();
    const calendarId = props.getProperty('CONSULTA_CALENDAR_ID');
    let calendar;

    if (calendarId) {
      calendar = CalendarApp.getCalendarById(calendarId);
      if (calendar) return calendar;
    }

    // Fallback por nombre si el ID no es válido o no existe
    const calendars = CalendarApp.getCalendarsByName(GOOGLE_CALENDAR_NAME);
    if (calendars.length > 0) {
      calendar = calendars[0];
    } else {
      calendar = CalendarApp.createCalendar(GOOGLE_CALENDAR_NAME);
    }

    props.setProperty('CONSULTA_CALENDAR_ID', calendar.getId());
    return calendar;
  }

  /**
   * Sincroniza todo lo pendiente (Sesiones y Días Bloqueados).
   */
  static syncAll() {
    const calendar = this.getCalendar();
    const resSesiones = this.syncSessions(calendar);
    const resBloqueos = this.syncBlockedDays(calendar);
    return { sesiones: resSesiones, bloqueos: resBloqueos };
  }

  /**
   * Sincroniza las sesiones de la base de datos con el calendario.
   */
  static syncSessions(calendar) {
    const sesiones = sessionRepo.findAll();
    const pacientes = patientRepo.findAll();
    const pacMap = new Map(pacientes.map(p => [p.PacienteID, p]));
    
    let syncCount = 0;

    sesiones.forEach(sesion => {
      const paciente = pacMap.get(sesion.PacienteID);
      if (!paciente) return;

      // Solo sincronizamos si ha cambiado o no tiene ID de evento
      const currentHash = this._generateHash(sesion);
      if (sesion.CalendarHash === currentHash && sesion.CalendarEventId) return;

      try {
        const event = this._upsertSessionEvent(calendar, sesion, paciente);
        if (event) {
          sesion.CalendarEventId = event.getId();
          sesion.CalendarSyncStatus = 'SINCRONIZADO';
          sesion.CalendarLastSync = new Date();
          sesion.CalendarHash = currentHash;
          sesion.CalendarEventTitle = event.getTitle();
          sessionRepo.save(sesion);
          syncCount++;
        }
      } catch (e) {
        sesion.CalendarSyncStatus = 'ERROR';
        sessionRepo.save(sesion);
        Logger.log(`Error sync sesión ${sesion.SesionID}: ${e.message}`);
      }
    });

    return syncCount;
  }

  /**
   * Sincroniza los días bloqueados de la hoja al calendario.
   */
  static syncBlockedDays(calendar) {
    // Aquí movemos la lógica del bloque 23 de forma más limpia
    const bloqueos = obtenerDiasBloqueados(); // Función existente en 23_...
    // ... lógica de limpieza de eventos antiguos de bloqueos y creación de nuevos ...
    // (Para brevedad, asumimos la integración de la lógica ya funcional del Bloque 23)
    sincronizarDiasBloqueadosAGoogleCalendar(calendar);
    return bloqueos.length;
  }

  /** @private */
  static _upsertSessionEvent(calendar, sesion, paciente) {
    const titulo = this._buildTitle(sesion, paciente);
    const descripcion = this._buildDescription(sesion, paciente);
    const fecha = normalizarFecha_(sesion.FechaSesion);
    const color = this._getColor(sesion.Modalidad);

    let event = null;
    if (sesion.CalendarEventId) {
      try { event = calendar.getEventById(sesion.CalendarEventId); } catch(e) {}
    }

    if (sesion.EstadoSesion === ESTADOS_SESION.CANCELADA) {
      if (event) event.deleteEvent();
      return null;
    }

    if (event) {
      event.setTitle(titulo);
      event.setDescription(descripcion);
      event.setAllDayDate(fecha);
    } else {
      event = calendar.createAllDayEvent(titulo, fecha, { description: descripcion });
    }
    
    try { event.setColor(color); } catch(e) {}
    return event;
  }

  /** @private */
  static _buildTitle(sesion, paciente) {
    const prefijos = {
      [ESTADOS_SESION.COMPLETADA_AUTO]: '✅ ',
      [ESTADOS_SESION.COMPLETADA_MANUAL]: '✅ ',
      [ESTADOS_SESION.CANCELADA]: '❌ ',
      [ESTADOS_SESION.REPROGRAMADA]: '🔁 '
    };
    const prefijo = prefijos[sesion.EstadoSesion] || '';
    return `${prefijo}${paciente.Nombre} - S${sesion.NumeroSesion} - ${sesion.Modalidad}`;
  }

  /** @private */
  static _buildDescription(sesion, paciente) {
    return `PACIENTE: ${paciente.Nombre}\nNHC: ${paciente.NHC || '-'}\nESTADO: ${sesion.EstadoSesion}\nNOTAS: ${sesion.Notas || '-'}`;
  }

  /** @private */
  static _getColor(modalidad) {
    const colores = {
      [MODALIDADES.INDIVIDUAL]: CalendarApp.EventColor.GREEN,
      [MODALIDADES.GRUPO_1]: CalendarApp.EventColor.BLUE,
      [MODALIDADES.GRUPO_2]: CalendarApp.EventColor.PURPLE,
      [MODALIDADES.GRUPO_3]: CalendarApp.EventColor.ORANGE
    };
    return colores[modalidad] || CalendarApp.EventColor.GRAY;
  }

  /** @private */
  static _generateHash(s) {
    return Utilities.base64Encode(`${s.SesionID}|${s.FechaSesion}|${s.EstadoSesion}|${s.Notas}`);
  }
}