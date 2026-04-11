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
   * @param {number} requiredDurationMinutes - Duración requerida del slot en minutos (ej. 30, 90).
   * @returns {AgendaSlot|null} El primer slot disponible encontrado, o null si no hay.
   */
  findNextAvailableSlot(startSearchDateTime, modality, requiredDurationMinutes) {
    // Aseguramos que empezamos a buscar con la hora correcta
    let currentDateTime = new Date(startSearchDateTime.getTime());
    // Reducimos el límite a 90 días para búsquedas iniciales, suficiente para encontrar un hueco
    let searchLimitDate = sumarDiasNaturales_(currentDateTime, 90); 

    // OPTIMIZACIÓN: Solo cargamos todas las sesiones una vez por cada AvailabilityService
    if (!this._allSessions) {
      this._allSessions = this.sessionRepo.findAll();
    }
    const allSessions = this._allSessions;

    // VALIDACIÓN PREVIA: Si no hay slots de esta modalidad en la plantilla, no busques (evita loops inútiles)
    const hasTemplate = this.agendaService.getWeeklyTemplate().some(s => {
      // Mapeamos el slot de la plantilla al formato que espera _isSlotCompatible
      return this._isSlotCompatible({
        type: s.TipoSlot,
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

    while (compararFechasHoras_(currentDateTime, searchLimitDate) <= 0) {
      const agendaForDay = this.agendaService.getAgendaForDay(currentDateTime);
      const sessionsForDay = sessionsMap[obtenerClaveFecha_(currentDateTime)] || [];

      // Obtener slots ocupados por sesiones existentes
      const occupiedSlots = this._getOccupiedSlotsFromSessions(sessionsForDay);

      for (const agendaSlot of agendaForDay) {
        // Si el slot de la agenda ya pasó la hora de inicio de búsqueda, o es el mismo slot
        if (agendaSlot.startDateTime.getTime() >= currentDateTime.getTime()) {
          // Verificar si el slot es compatible con la modalidad y duración
          if (this._isSlotCompatible(agendaSlot, modality, requiredDurationMinutes)) {
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

      // Si no se encontró slot en el día actual, avanzar al siguiente día a la primera hora de la plantilla
      currentDateTime = sumarDiasNaturales_(currentDateTime, 1);
      // Establecer la hora de inicio del día a la primera hora de la plantilla si existe
      const nextDayTemplate = this.agendaService.getAgendaForDay(currentDateTime);
      if (nextDayTemplate.length > 0) {
        currentDateTime = normalizarFechaHora_(currentDateTime, formatearHora_(nextDayTemplate[0].startDateTime));
      } else {
        // Si el siguiente día no tiene plantilla, simplemente ir a medianoche
        currentDateTime = normalizarFechaHora_(currentDateTime, '00:00');
      }
    }

    return null; // No se encontró ningún slot disponible
  }

  /**
   * Determina si un slot de la agenda es compatible con una modalidad y duración requerida.
   * @param {AgendaSlot} agendaSlot - El slot de la agenda.
   * @param {string} modality - La modalidad del paciente.
   * @param {number} requiredDurationMinutes - Duración requerida.
   * @returns {boolean} True si es compatible, false en caso contrario.
   */
  _isSlotCompatible(agendaSlot, modality, requiredDurationMinutes) {
    const slotType = String(agendaSlot.type || '').trim().toUpperCase();
    const mod = String(modality || '').trim().toUpperCase();

    // Reglas de compatibilidad de tipo de slot
    if (slotType === 'DESCANSO' || slotType === '') return false;

    if (mod === 'INDIVIDUAL') {
      // Las sesiones 2.2 generadas solo deben ir en slots tipo 2.2 o SEGUIMIENTO
      if (slotType !== '2.2' && slotType !== 'SEGUIMIENTO' && slotType !== '2.1' && slotType !== 'PRIMERA') return false;
    } else if (mod.startsWith('GRUPO')) {
      // Un ciclo (ej. GRUPO_1) puede usar slots específicos o el genérico 'GRUPO'
      const tiposValidos = [mod, 'GRUPO', '2.2/GRUPO', 'SEGUIMIENTO/GRUPO'];
      if (!tiposValidos.includes(slotType)) return false;
    } else {
      // Otras modalidades no soportadas por la generación automática
      return false;
    }

    // Reglas de compatibilidad de duración
    return agendaSlot.durationMinutes >= requiredDurationMinutes;
  }

  /**
   * Obtiene una lista de rangos de tiempo ocupados por sesiones existentes.
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
      if (s.FechaSesion instanceof Date && s.EstadoSesion !== ESTADOS_SESION.CANCELADA) {
        const key = obtenerClaveFecha_(s.FechaSesion);
        if (!sessionsMap[key]) sessionsMap[key] = [];
        sessionsMap[key].push(s); 
      }
    });

    for (let i = 0; i < 7; i++) {
      const date = sumarDiasNaturales_(today, i);
      const isToday = i === 0;
      const dateKey = obtenerClaveFecha_(date);
      
      // Saltamos si el día está bloqueado (Festivos/Fines de semana)
      if (esFinDeSemana_(date) || blockedDays[dateKey]) continue;

      const agendaForDay = this.agendaService.getAgendaForDay(date);
      const sessionsForDay = sessionsMap[dateKey] || [];
      const occupiedSlots = this._getOccupiedSlotsFromSessions(sessionsForDay);

      const freeSlots = agendaForDay.filter(slot => {
        if (slot.type === 'DESCANSO') return false;
        // Si es hoy, solo mostramos slots futuros
        if (isToday && slot.startDateTime.getTime() <= today.getTime()) return false;
        return !this._isSlotOccupied(slot, occupiedSlots);
      });

      if (freeSlots.length > 0) {
        const diaSemanaStr = convertirDiaSemanaATexto_(date);
        summary.push({
          fecha: formatearFecha_(date),
          diaSemana: diaSemanaStr,
          // Normalizamos el tipo para que el CSS del Dashboard funcione
          slots: freeSlots.map(s => ({ 
            hora: formatearHora_(s.startDateTime), 
            tipo: this._normalizeTypeForUI(s.type) 
          }))
        });
      }
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
      'SEGUIMIENTO/GRUPO': '2.2/GRUPO'
    };
    return map[type] || type;
  }
}
