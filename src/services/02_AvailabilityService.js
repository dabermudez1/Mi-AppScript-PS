/**
 * Servicio para determinar la disponibilidad de slots horarios.
 * Combina la agenda (plantilla + excepciones) con las sesiones ya programadas.
 */
class AvailabilityService {
  constructor() {
    this.agendaService = new AgendaService();
    this.sessionRepo = new SessionRepository();
    this._allSessions = null; // Caché interna de la instancia
  }

  /**
   * Encuentra el siguiente slot disponible compatible con la modalidad y duración requerida.
   * @param {Date} startSearchDateTime - Fecha y hora a partir de la cual empezar a buscar.
   * @param {string} modality - Modalidad del paciente (ej. INDIVIDUAL, GRUPO_1).   
   * @param {number} [requiredDurationMinutes] - Duración requerida. Si no se pasa, se deduce de la modalidad.
   * @param {string} [sessionType='SEGUIMIENTO'] - El tipo de sesión a buscar ('PRIMERA', 'SEGUIMIENTO').
   * @param {string} [ignoreSessionId] - ID de la sesión a ignorar en la búsqueda de ocupación.
   * @returns {AgendaSlot|null} El primer slot disponible encontrado, o null si no hay.
   */
  findNextAvailableSlot(startSearchDateTime, modality, sessionType = 'SEGUIMIENTO', requiredDurationMinutes = null, ignoreSessionId = null) {
    // Aseguramos que empezamos a buscar con la hora correcta
    let currentDateTime = new Date(startSearchDateTime.getTime());

    // --- NUEVA LÓGICA ---
    // Si no se especifica una duración, la deducimos de la modalidad para evitar errores.
    if (!requiredDurationMinutes) {
      requiredDurationMinutes = this.agendaService._getSlotDuration(sessionType);
    }
    // --- FIN NUEVA LÓGICA ---

    // Para ciclos, 60 días de búsqueda es más que suficiente y evita Timeouts
    let searchLimitDate = sumarDiasNaturales_(currentDateTime, 60); 

    // Cargar configuración para validar restricciones de día de la semana (Grupo)
    const config = (new ConfigRepository()).findByModalidad(modality);
    const targetDay = (config && config.TipoModalidad === 'GRUPO') ? String(config.DiaSemana).trim().toUpperCase() : null;

    // Si estamos en una ejecución larga, forzamos limpieza de caché de sesiones antes de empezar un ciclo
    if (typeof __EXECUTION_CACHE__ !== 'undefined') __EXECUTION_CACHE__[SHEET_SESIONES] = null;

    // OPTIMIZACIÓN: Solo cargamos todas las sesiones una vez por cada AvailabilityService
    if (!this._allSessions) {
      this._allSessions = this.sessionRepo.findAll();
    }
    const allSessions = this._allSessions;

    // VALIDACIÓN PREVIA: Si no hay slots de esta modalidad en la plantilla, no busques (evita loops inútiles)
    const hasTemplate = this.agendaService.getWeeklyTemplate().some(s => {
      // Mapeamos el slot de la plantilla al formato que espera _isSlotCompatible
      return this._isSlotCompatible({
        type: this._normalizeTypeForCompatibility(s.TipoSlot),
        durationMinutes: this.agendaService._getSlotDuration(s.TipoSlot)
      }, modality, requiredDurationMinutes);
    });
    if (!hasTemplate) throw new Error(`No hay slots de tipo ${modality} configurados en la 'Plantilla de Agenda'.`);

    const sessionsMap = {};
    allSessions.forEach(s => {
      // Intentamos parsear la fecha si no es objeto Date (seguridad extra)
      const fecha = s.FechaSesion instanceof Date ? s.FechaSesion : parseFechaES_(s.FechaSesion);
      
      if (fecha && s.EstadoSesion !== ESTADOS_SESION.CANCELADA) {
        const key = obtenerClaveFecha_(fecha);
        if (!sessionsMap[key]) sessionsMap[key] = [];
        sessionsMap[key].push(s);
      }
    });

    const weeklyTemplate = this.agendaService.getWeeklyTemplate();

    while (compararFechasHoras_(currentDateTime, searchLimitDate) <= 0) {
      const dayOfWeek = convertirDiaSemanaATexto_(currentDateTime);
      
      // Si es un grupo con día fijo configurado, saltamos los días que no coincidan
      if (targetDay && dayOfWeek !== targetDay) {
        currentDateTime = sumarDiasNaturales_(currentDateTime, 1);
        continue;
      }

      // Solo procesamos si el día de la semana existe en la plantilla (vía rápida)
      if (weeklyTemplate.some(t => t.DiaSemana === dayOfWeek)) {
        const agendaForDay = this.agendaService.getAgendaForDay(currentDateTime);
      
      // Filtramos sesiones del día, ignorando las del ciclo actual si se solicita
      const sessionsForDay = (sessionsMap[obtenerClaveFecha_(currentDateTime)] || []).filter(s => 
        !ignoreSessionId || String(s.SesionID) !== String(ignoreSessionId)
      );

      // Obtener slots ocupados por sesiones existentes
      const occupiedSlots = this._getOccupiedSlotsFromSessions(sessionsForDay);

      for (const agendaSlot of agendaForDay) {
        // Si el slot de la agenda ya pasó la hora de inicio de búsqueda, o es el mismo slot
        if (agendaSlot.startDateTime.getTime() >= currentDateTime.getTime()) {
          // Verificar si el slot es compatible con la modalidad y duración
          if (this._isSlotCompatible(agendaSlot, sessionType, requiredDurationMinutes)) {
            // Verificar si el slot está ocupado por una sesión existente
            if (!this._isSlotOccupied(agendaSlot, occupiedSlots)) {
              // Verificar si el día está completamente bloqueado (ej. por DIAS_BLOQUEADOS)
              // FIX: Añadimos comprobación de existencia del helper para evitar ReferenceError
              if (typeof esFechaBloqueada_ !== 'function' || !esFechaBloqueada_(agendaSlot.startDateTime)) {
                return agendaSlot; // ¡Slot encontrado!
              }
            }
          }
        }
      }
      }

      // Si no se encontró slot en el día actual, avanzar al siguiente día a la primera hora de la plantilla
      currentDateTime = sumarDiasNaturales_(currentDateTime, 1);
      // CORRECCIÓN: Reiniciar siempre a medianoche para escanear el día completo.
      currentDateTime = normalizarFechaHora_(currentDateTime, '00:00');
    }

    return null; // No se encontró ningún slot disponible
  }

  /**
   * Determina si un slot de la agenda es compatible con una modalidad y duración requerida.
   * @private
   * @param {AgendaSlot} agendaSlot - El slot de la agenda.
   * @param {string} sessionType - El tipo de sesión a buscar ('PRIMERA', 'SEGUIMIENTO', 'GRUPO').
   * @param {number} requiredDurationMinutes - Duración requerida.
   * @returns {boolean} True si es compatible, false en caso contrario.
   */
  _isSlotCompatible(agendaSlot, sessionType, requiredDurationMinutes) {
    // --- CORRECCIÓN CRÍTICA ---
    // El parámetro 'sessionType' a veces recibe una Modalidad ('INDIVIDUAL', 'GRUPO_1') en lugar 
    // de un Tipo de Sesión ('SEGUIMIENTO', 'GRUPO'). Normalizamos esto primero.
    let searchType = String(sessionType || '').trim().toUpperCase();
    if (searchType.startsWith('GRUPO')) {
      searchType = 'GRUPO';
    } else if (searchType === 'INDIVIDUAL') {
      searchType = 'SEGUIMIENTO';
    }

    const normalizedTemplateType = this._normalizeTypeForCompatibility(agendaSlot.type);
    const normalizedSearchType = this._normalizeTypeForCompatibility(searchType);

    // Reglas de compatibilidad de tipo de slot
    if (normalizedTemplateType === 'DESCANSO' || normalizedTemplateType === '') return false;

    switch (normalizedSearchType) {
      case 'SEGUIMIENTO':
        // Una sesión de seguimiento individual puede ir en un slot de SEGUIMIENTO o en uno de PRIMERA.
        if (normalizedTemplateType !== 'SEGUIMIENTO' && normalizedTemplateType !== 'PRIMERA') {
          return false;
        }
        break;
      
      case 'PRIMERA':
        // Una primera consulta solo puede ir en un slot de PRIMERA.
        if (normalizedTemplateType !== 'PRIMERA') {
          return false;
        }
        break;

      case 'GRUPO':
      case 'SEGUIMIENTO/GRUPO':
        // Una sesión de grupo solo puede ir en un slot de GRUPO.
        if (normalizedTemplateType !== 'SEGUIMIENTO/GRUPO') {
          return false;
        }
        break;

      default:
        // Tipo de búsqueda no reconocido.
        return false;
    }

    // Reglas de compatibilidad de duración
    return agendaSlot.durationMinutes >= requiredDurationMinutes;
  }

  /**
   * Obtiene una lista de rangos de tiempo ocupados por sesiones existentes.
   * @private
   * @param {Array<Object>} sessions - Lista de objetos de sesión.
   * @returns {Array<{start: Date, end: Date}>} Rangos de tiempo ocupados.
   */
  _getOccupiedSlotsFromSessions(sessions) {
    const occupied = [];
    sessions.forEach(s => {
      if (s.EstadoSesion !== ESTADOS_SESION.CANCELADA) {
        // Aseguramos que HoraInicio se trate correctamente tanto si es String como Date
        let horaStr = s.HoraInicio;
        if (horaStr instanceof Date) {
          horaStr = formatearHora_(horaStr);
        }
        const start = normalizarFechaHora_(s.FechaSesion, horaStr);
                
        // Preferimos la duración guardada en la sesión si existe
        let duration = Number(s.Duracion);
        
        if (!duration) {
          let sessionSlotType = s.Modalidad === MODALIDADES.INDIVIDUAL ? '2.2' : '2.2/GRUPO';
          duration = this.agendaService._getSlotDuration(sessionSlotType);
        }

        const end = sumarMinutos_(start, duration);
        occupied.push({ start, end });
      }
    });
    return occupied;
  }

  /**
   * Verifica si un slot de la agenda se solapa con algún slot ocupado.
   * @private
   * @param {AgendaSlot} agendaSlot - El slot de la agenda a verificar.
   * @param {Array<{start: Date, end: Date}>} occupiedSlots - Lista de slots ya ocupados.
   * @returns {boolean} True si el slot está ocupado, false en caso contrario.
   */
  _isSlotOccupied(agendaSlot, occupiedSlots) {
    const slotEnd = sumarMinutos_(agendaSlot.startDateTime, agendaSlot.durationMinutes);
    return occupiedSlots.some(occ =>
      (agendaSlot.startDateTime < occ.end && slotEnd > occ.start)
    );
  }

  /**
   * Genera un resumen de huecos libres para los próximos 7 días.
   * @returns {Array<Object>} Resumen por día.
   */
  getFreeSlotsSummary() {
    const today = new Date();
    const summary = [];

    // 1. Carga masiva de datos (Una sola lectura a disco)
    const allSessions = this.sessionRepo.findAll(); 
    const blockedDays = obtenerMapaDiasBloqueados_();
    const weeklyTemplate = this.agendaService.getWeeklyTemplate();

    // 2. Indexación rápida de sesiones
    const sessionsMap = {};
    allSessions.forEach(s => {
      // Intentamos parsear la fecha si no es objeto Date
      const fecha = s.FechaSesion instanceof Date ? s.FechaSesion : parseFechaES_(s.FechaSesion);
      if (fecha && s.EstadoSesion !== ESTADOS_SESION.CANCELADA) {
        const key = obtenerClaveFecha_(fecha);
        if (!sessionsMap[key]) sessionsMap[key] = [];
        sessionsMap[key].push(s); 
      }
    });

    for (let i = 0; i < 7; i++) {
      const date = sumarDiasNaturales_(today, i);
      const isToday = i === 0;
      const dateKey = obtenerClaveFecha_(date);
      const diaSemanaStr = convertirDiaSemanaATexto_(date);
      
      let dayInfo = {
        fecha: formatearFecha_(date),
        diaSemana: diaSemanaStr,
        blocked: false,
        reason: '',
        slots: []
      };

      if (esFinDeSemana_(date)) {
        dayInfo.blocked = true;
        dayInfo.reason = 'Fin de semana';
      } else if (blockedDays[dateKey]) {
        dayInfo.blocked = true;
        dayInfo.reason = blockedDays[dateKey].motivo || 'Festivo / Bloqueado';
      } else {
        const agendaForDay = this.agendaService.getAgendaForDay(date);
        const sessionsForDay = sessionsMap[dateKey] || [];
        const occupiedSlots = this._getOccupiedSlotsFromSessions(sessionsForDay);

        const freeSlots = agendaForDay.filter(slot => {
          if (slot.type === 'DESCANSO') return false;
          if (isToday && slot.startDateTime.getTime() <= today.getTime()) return false;
          return !this._isSlotOccupied(slot, occupiedSlots);
        });

        dayInfo.slots = freeSlots.map(s => ({ 
          hora: formatearHora_(s.startDateTime), 
          tipo: this._normalizeTypeForUI(s.type) 
        }));
        
        if (dayInfo.slots.length === 0) {
          dayInfo.reason = 'Sin huecos disponibles';
        }
      }
      summary.push(dayInfo);
    }
    return summary;
  }

  /**
   * Mapea nombres descriptivos a códigos técnicos para el CSS del Dashboard
   * @private
   */
  _normalizeTypeForUI(type) {
    const map = {
      'SEGUIMIENTO': '2.2',
      'PRIMERA': '2.1',
      'SEGUIMIENTO/GRUPO': '2.2/GRUPO',
      '2.1/RESERVA': '2.1 (Reservado)' // Nuevo mapeo para reservas 2.1
    };
    return map[type] || type;
  }

  /**
   * Normaliza los diferentes alias de tipos de slot a un estándar para la lógica interna.
   * @private
   */
  _normalizeTypeForCompatibility(type) {
    const t = String(type || '').trim().toUpperCase();
    switch (t) {
      case '2.1': case 'PRIMERA': return 'PRIMERA';
      case '2.2': case 'SEGUIMIENTO': return 'SEGUIMIENTO';
      case '2.2/GRUPO': case 'SEGUIMIENTO/GRUPO': case 'GRUPO': return 'SEGUIMIENTO/GRUPO';
      case 'DESCANSO': return 'DESCANSO';
      default: return t;
    }
  }

  /**
   * Genera una representación detallada del estado de la agenda para una semana completa.
   * @param {Date} startDate - Cualquier día de la semana que se quiere visualizar.
   * @returns {Array<Object>} Un array de 7 objetos, uno por cada día de la semana.
   */
  getWeeklyState(startDate) {
    const weekData = [];
    const startOfWeek = this._getStartOfWeek(startDate);

    // Carga masiva de datos para toda la semana
    const allSessions = this.sessionRepo.findAll();
    const blockedDaysMap = obtenerMapaDiasBloqueados_();

    for (let i = 0; i < 7; i++) {
      const currentDate = sumarDiasNaturales_(startOfWeek, i);
      const dateKey = obtenerClaveFecha_(currentDate);
      const dayOfWeekLabel = convertirDiaSemanaATexto_(currentDate);

      const dayInfo = {
        dateLabel: formatearFecha_(currentDate),
        dayOfWeekLabel: dayOfWeekLabel,
        isBlocked: false,
        blockReason: '',
        slots: []
      };

      // Comprobar si es fin de semana o un día bloqueado
      if (esFinDeSemana_(currentDate)) {
        dayInfo.isBlocked = true;
        dayInfo.blockReason = 'Fin de semana';
      } else if (blockedDaysMap[dateKey]) {
        dayInfo.isBlocked = true;
        dayInfo.blockReason = blockedDaysMap[dateKey].motivo || 'Día bloqueado';
      }

      if (dayInfo.isBlocked) {
        weekData.push(dayInfo);
        continue;
      }

      // Obtener plantilla y sesiones para el día
      const agendaForDay = this.agendaService.getAgendaForDay(currentDate);
      const sessionsForDay = allSessions.filter(s =>
        s.FechaSesion && normalizarFecha_(s.FechaSesion).getTime() === currentDate.getTime() &&
        s.EstadoSesion !== ESTADOS_SESION.CANCELADA
      );

      const occupiedSlots = this._getOccupiedSlotsFromSessions(sessionsForDay);

      dayInfo.slots = agendaForDay.map(agendaSlot => {
        const slotState = {
          time: formatearHora_(agendaSlot.startDateTime),
          templateType: agendaSlot.type,
          status: '',
          occupiedBy: '',
          sessionNumber: null,
          durationMinutes: agendaSlot.durationMinutes, // Añadido para el frontend
          sessionStatus: null // Nuevo campo para el estado de la sesión
        };

        if (agendaSlot.type === 'DESCANSO') {
          slotState.status = 'TEMPLATE_REST';
        } else {
          const isOccupied = this._isSlotOccupied(agendaSlot, occupiedSlots);
          if (isOccupied) {
            slotState.status = 'OCCUPIED';
            // Encontrar qué sesión ocupa este slot
            const occupyingSession = sessionsForDay.find(s => {
                const start = normalizarFechaHora_(s.FechaSesion, s.HoraInicio);
                const end = sumarMinutos_(start, Number(s.Duracion || 30));
                return agendaSlot.startDateTime < end && sumarMinutos_(agendaSlot.startDateTime, agendaSlot.durationMinutes) > start;
            });
            if (occupyingSession) {
              slotState.occupiedBy = occupyingSession.NombrePaciente || 'N/A';
              slotState.sessionNumber = occupyingSession.NumeroSesion || null;
              // --- CORRECCIÓN CLAVE ---
              // La duración de la sesión ocupada debe basarse en su MODALIDAD, no en la plantilla.
              // Esto evita que una sesión individual de 30min se "estire" si cae en un slot de grupo de 90min.
              slotState.durationMinutes = this._getSlotDuration(occupyingSession.Modalidad);
              slotState.sessionModality = occupyingSession.Modalidad; // NUEVO: Enviamos la modalidad real
              slotState.sessionStatus = occupyingSession.EstadoSesion || null;
            }
          } else {
            slotState.status = 'FREE';
          }
        }
        return slotState;
      });

      weekData.push(dayInfo);
    }

    return weekData;
  }

  /**
   * Calcula el inicio de la semana (Lunes) para una fecha dada.
   * @param {Date} date - La fecha.
   * @returns {Date} El lunes de esa semana.
   * @private
   */
  _getStartOfWeek(date) {
    const d = new Date(date);
    const day = d.getDay(); // Domingo = 0, Lunes = 1, ...
    const diff = d.getDate() - day + (day === 0 ? -6 : 1); // Ajuste para que Lunes sea el primer día
    return new Date(d.setDate(diff));
  }
}
