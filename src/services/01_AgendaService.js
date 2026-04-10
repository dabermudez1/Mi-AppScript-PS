/**
 * Servicio para gestionar la plantilla semanal y las excepciones de la agenda.
 * Consolida la información de AGENDA_PLANTILLA y AGENDA_EXCEPCIONES.
 */
class AgendaService {
  constructor() {
    this.templateRepo = new AgendaTemplateRepository();
    this.exceptionRepo = new AgendaExceptionRepository();
    // Cache de datos para evitar lecturas repetitivas en bucles
    this._cachedTemplate = null;
    this._cachedExceptions = null;
  }

  /**
   * Obtiene la plantilla semanal de slots.
   * @returns {Array<Object>} Lista de objetos de plantilla.
   */
  getWeeklyTemplate() {
    if (!this._cachedTemplate) {
      this._cachedTemplate = this.templateRepo.findAll();
    }
    return this._cachedTemplate;
  }

  /**
   * Obtiene las excepciones de la agenda en un rango de fechas.
   * @param {Date} startDate - Fecha de inicio del rango.
   * @param {Date} endDate - Fecha de fin del rango.
   * @returns {Array<Object>} Lista de objetos de excepción.
   */
  getExceptions(startDate, endDate) {
    if (!this._cachedExceptions) {
      this._cachedExceptions = this.exceptionRepo.findAll();
    }
    const normalizedStartDate = normalizarFecha_(startDate);
    const normalizedEndDate = normalizarFecha_(endDate);

    return this._cachedExceptions.filter(ex => {
      const exDate = normalizarFecha_(ex.Fecha);
      return exDate >= normalizedStartDate && exDate <= normalizedEndDate;
    });
  }

  /**
   * Combina la plantilla semanal y las excepciones para un día específico,
   * devolviendo una lista de slots disponibles para ese día.
   * @param {Date} date - La fecha para la que se quiere obtener la agenda.
   * @returns {Array<AgendaSlot>} Lista de objetos AgendaSlot para el día.
   */
  getAgendaForDay(date) {
    const normalizedDate = normalizarFecha_(date);
    const dayOfWeek = convertirDiaSemanaATexto_(normalizedDate);

    const templateSlots = this.getWeeklyTemplate()
      .filter(slot => {
        const slotDay = String(slot.DiaSemana || '').trim().toUpperCase();
        return slotDay === dayOfWeek;
      })
      .map(slot => ({
        startDateTime: normalizarFechaHora_(normalizedDate, slot.HoraInicio),
        type: String(slot.TipoSlot || '').trim().toUpperCase(),
        durationMinutes: this._getSlotDuration(slot.TipoSlot)
      }));

    const exceptions = this.getExceptions(normalizedDate, normalizedDate);

    // Aplicar excepciones
    let finalSlots = [...templateSlots];

    exceptions.forEach(ex => {
      const exStartDateTime = ex.HoraInicio ? normalizarFechaHora_(normalizedDate, ex.HoraInicio) : null;

      if (exStartDateTime) {
        // Excepción de slot específico
        const index = finalSlots.findIndex(slot =>
          compararFechasHoras_(slot.startDateTime, exStartDateTime) === 0
        );
        if (index !== -1) {
          if (ex.TipoSlot === 'LIBRE') {
            // Si la excepción es "LIBRE", eliminamos el slot de la plantilla
            finalSlots.splice(index, 1);
          } else {
            // Si es otro tipo, actualizamos el tipo de slot
            finalSlots[index].type = ex.TipoSlot;
            finalSlots[index].durationMinutes = this._getSlotDuration(ex.TipoSlot);
          }
        } else if (ex.TipoSlot !== 'LIBRE') {
          // Si no existe en plantilla y no es LIBRE, lo añadimos (ej. un slot extra)
          finalSlots.push({
            startDateTime: exStartDateTime,
            type: ex.TipoSlot,
            durationMinutes: this._getSlotDuration(ex.TipoSlot)
          });
        }
      } else {
        // Excepción de día completo
        if (ex.TipoSlot === 'DESCANSO') {
          finalSlots = []; // Bloquea todo el día
        } else if (ex.TipoSlot === 'LIBRE') {
          // Un día completo "LIBRE" podría anular un día de descanso de la plantilla,
          // pero la plantilla ya se ha aplicado. Esto es más complejo y quizás no necesario para MVP.
          // Por ahora, si es día completo LIBRE, no hace nada si no hay DESCANSO en plantilla.
        }
      }
    });

    // Ordenar por hora de inicio
    finalSlots.sort((a, b) => compararFechasHoras_(a.startDateTime, b.startDateTime));

    return finalSlots;
  }

  _getSlotDuration(slotType) {
    const type = String(slotType || '').trim().toUpperCase();
    // Reglas de duración de slots
    switch (type) {
      case '2.2': case 'SEGUIMIENTO': return 30;
      case '2.1': case 'PRIMERA': return 60; // 2 slots de 30 min
      case '2.2/GRUPO': case 'SEGUIMIENTO/GRUPO': return 60; // 2 slots de 30 min (según nueva instrucción)
      case 'DESCANSO': return 30; // Aunque bloquea, un slot de descanso es de 30 min
      default: return 30; // Por defecto, 30 minutos
    }
  }
}

/**
 * Representa un slot horario disponible en la agenda.
 * @typedef {Object} AgendaSlot
 * @property {Date} startDateTime - Fecha y hora de inicio del slot.
 * @property {string} type - Tipo de slot (ej. "2.2", "2.2/GRUPO", "DESCANSO").
 * @property {number} durationMinutes - Duración del slot en minutos (ej. 30).
 */