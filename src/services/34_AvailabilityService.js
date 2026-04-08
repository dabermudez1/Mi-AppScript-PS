/**
 * Servicio encargado de la lógica de disponibilidad y asignación de slots.
 */
class AvailabilityService {
  constructor(agendaRepo, sessionRepo) {
    this.agendaRepo = agendaRepo;
    this.sessionRepo = sessionRepo;
  }

  /**
   * Encuentra el siguiente slot disponible respetando la frecuencia.
   * @param {Date} targetDate Fecha sugerida (ej. ultimaSesion + frecuencia)
   * @param {string} modalidad Tipo de modalidad (ej. "2.2")
   */
  findNextAvailableSlot(targetDate, modalidad) {
    let searchDate = new Date(targetDate);
    const weeklyTemplate = this.agendaRepo.getWeeklyTemplate();
    const exceptions = this.agendaRepo.getExceptions();
    
    // Mapeo de reglas de negocio
    const requiredType = this._getRequiredSlotType(modalidad);

    // Buscamos hasta en 30 días posteriores si no hay hueco
    for (let i = 0; i < 30; i++) {
      const dateStr = Utilities.formatDate(searchDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
      const dayName = this._getDayName(searchDate);

      // 1. Obtener slots del día (Excepciones > Plantilla)
      let dailySlots = exceptions.filter(e => e.fecha === dateStr);
      if (dailySlots.length === 0) {
        dailySlots = weeklyTemplate[dayName] || [];
      }

      // 2. Filtrar y buscar primer slot libre del tipo correcto
      const freeSlot = dailySlots.find(slot => {
        if (slot.tipo !== requiredType) return false;
        return !this.sessionRepo.isSlotOccupied(dateStr, slot.hora);
      });

      if (freeSlot) {
        return {
          fecha: new Date(searchDate),
          hora: freeSlot.hora
        };
      }

      // 3. Si no hay hueco, pasar al siguiente día
      searchDate.setDate(searchDate.getDate() + 1);
    }
    
    throw new Error(`No se encontró disponibilidad para ${modalidad} tras 30 días de búsqueda.`);
  }

  _getRequiredSlotType(modalidad) {
    if (modalidad === "2.2") return "SEGUIMIENTO";
    if (modalidad === "2.1") return "PRIMERA";
    if (modalidad.includes("GRUPO")) return "SEGUIMIENTO/GRUPO";
    return "SEGUIMIENTO";
  }

  _getDayName(date) {
    const dias = ["DOMINGO", "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO"];
    return dias[date.getDay()];
  }
}
