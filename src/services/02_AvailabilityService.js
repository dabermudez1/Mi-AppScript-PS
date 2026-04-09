/**
 * Servicio para determinar la disponibilidad de slots horarios.
 * Combina la agenda (plantilla + excepciones) con las sesiones ya programadas.
 */
class AvailabilityService {
  constructor() {
    this.agendaService = new AgendaService();
    this.sessionRepo = new SessionRepository();
    // Podríamos necesitar un repositorio para días bloqueados si queremos consultar el motivo
    // this.diasBloqueadosRepo = new DiasBloqueadosRepository();
  }

  /**
   * Encuentra el siguiente slot disponible compatible con la modalidad y duración requerida.
   * @param {Date} startSearchDateTime - Fecha y hora a partir de la cual empezar a buscar.
   * @param {string} modality - Modalidad del paciente (ej. INDIVIDUAL, GRUPO_1).
   * @param {number} requiredDurationMinutes - Duración requerida del slot en minutos (ej. 30, 90).
   * @returns {AgendaSlot|null} El primer slot disponible encontrado, o null si no hay.
   */
  findNextAvailableSlot(startSearchDateTime, modality, requiredDurationMinutes) {
    let currentDateTime = normalizarFechaHora_(startSearchDateTime);
    let searchLimitDate = sumarDiasNaturales_(currentDateTime, 365); // Buscar hasta 1 año en el futuro

    while (compararFechasHoras_(currentDateTime, searchLimitDate) <= 0) {
      const agendaForDay = this.agendaService.getAgendaForDay(currentDateTime);
      const sessionsForDay = this.sessionRepo.findAll().filter(s =>
        normalizarFecha_(s.FechaSesion).getTime() === normalizarFecha_(currentDateTime).getTime()
      );

      // Obtener slots ocupados por sesiones existentes
      const occupiedSlots = this._getOccupiedSlotsFromSessions(sessionsForDay);

      for (const agendaSlot of agendaForDay) {
        // Si el slot de la agenda ya pasó la hora de inicio de búsqueda, o es el mismo slot
        if (compararFechasHoras_(agendaSlot.startDateTime, currentDateTime) >= 0) {
          // Verificar si el slot es compatible con la modalidad y duración
          if (this._isSlotCompatible(agendaSlot, modality, requiredDurationMinutes)) {
            // Verificar si el slot está ocupado por una sesión existente
            if (!this._isSlotOccupied(agendaSlot, occupiedSlots)) {
              // Verificar si el día está completamente bloqueado (ej. por DIAS_BLOQUEADOS)
              if (!esFechaBloqueada_(agendaSlot.startDateTime)) { // Reutilizamos la función existente para días bloqueados
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
    // Reglas de compatibilidad de tipo de slot
    if (agendaSlot.type === 'DESCANSO') return false;

    if (modality === MODALIDADES.INDIVIDUAL) {
      // Individual busca slots '2.2'
      if (agendaSlot.type !== '2.2') return false;
    } else if (modality.startsWith('GRUPO')) {
      // Grupos buscan slots '2.2/GRUPO'
      if (agendaSlot.type !== '2.2/GRUPO') return false;
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
        const start = normalizarFechaHora_(s.FechaSesion, s.HoraInicio);
        // Asumimos que la duración de una sesión es 30 min si no se especifica o es 2.2
        // Esto podría ser más sofisticado leyendo la configuración de modalidad o el tipo de slot
        const duration = 30; // TODO: Obtener duración real de la sesión o tipo de slot
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
    
    for (let i = 0; i < 7; i++) {
      const date = sumarDiasNaturales_(today, i);
      if (esFechaBloqueada_(date)) continue;

      const agendaForDay = this.agendaService.getAgendaForDay(date);
      const sessionsForDay = this.sessionRepo.findAll().filter(s =>
        s.FechaSesion instanceof Date && 
        normalizarFecha_(s.FechaSesion).getTime() === normalizarFecha_(date).getTime() &&
        s.EstadoSesion !== ESTADOS_SESION.CANCELADA
      );

      const occupiedSlots = this._getOccupiedSlotsFromSessions(sessionsForDay);
      const freeSlots = agendaForDay.filter(slot => 
        slot.type !== 'DESCANSO' && !this._isSlotOccupied(slot, occupiedSlots)
      );

      if (freeSlots.length > 0) {
        summary.push({
          fecha: formatearFecha_(date),
          diaSemana: convertirDiaSemanaATexto_(date),
          slots: freeSlots.map(s => ({ hora: formatearHora_(s.startDateTime), tipo: s.type }))
        });
      }
    }
    return summary;
  }
}
