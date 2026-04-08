/**
 * Servicio para calcular disponibilidad basado en slots horarios.
 */
class AvailabilityService {
  constructor(agendaRepo, sessionRepo) {
    this.agendaRepo = agendaRepo;
    this.sessionRepo = sessionRepo;
  }

  findNextAvailableSlot(targetDate, modalidad) {
    let searchDate = new Date(targetDate);
    const weeklyTemplate = this.agendaRepo.getWeeklyTemplate();
    const exceptions = this.agendaRepo.getExceptions();
    const requiredType = this._mapModalidadToSlot(modalidad);

    for (let i = 0; i < 30; i++) {
      const dateStr = Utilities.formatDate(searchDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
      const dayName = this._getDayName(searchDate);

      let dailySlots = exceptions.filter(e => e.fecha === dateStr);
      if (dailySlots.length === 0) {
        dailySlots = weeklyTemplate[dayName] || [];
      }

      const freeSlot = dailySlots.find(slot => {
        if (slot.tipo !== requiredType) return false;
        return !this.sessionRepo.isSlotOccupied(dateStr, slot.hora);
      });

      if (freeSlot) {
        return { fecha: new Date(searchDate), hora: freeSlot.hora };
      }
      searchDate.setDate(searchDate.getDate() + 1);
    }
    throw new Error(`No hay huecos libres para ${modalidad} en los próximos 30 días.`);
  }

  _mapModalidadToSlot(mod) {
    const map = { "2.2": "SEGUIMIENTO", "2.1": "PRIMERA", "GRUPO": "SEGUIMIENTO/GRUPO" };
    return map[mod] || "SEGUIMIENTO";
  }

  _getDayName(date) {
    const dias = ["DOMINGO", "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO"];
    return dias[date.getDay()];
  }
}
