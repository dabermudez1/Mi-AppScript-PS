/**
 * Servicio de planificación de alto nivel.
 */
class PlanningService {
  constructor() {
    this.agendaRepo = new AgendaRepository();
    this.sessionRepo = new SessionRepository();
    this.availabilityService = new AvailabilityService(this.agendaRepo, this.sessionRepo);
  }

  /**
   * Genera automáticamente las sesiones 2.2 de un ciclo.
   * @param {Object} ciclo Datos del ciclo
   * @param {number} numSesiones Cantidad de sesiones a generar
   * @param {number} frecuencia Dias entre sesiones
   */
  planificarCicloSeguimiento(ciclo, numSesiones, frecuencia) {
    let ultimaFecha = new Date(ciclo.fechaInicio); // O la fecha de la última sesión existente
    const sesionesGeneradas = [];

    for (let i = 0; i < numSesiones; i++) {
      // Calcular fecha objetivo basada en frecuencia
      let fechaObjetivo = new Date(ultimaFecha);
      if (i > 0) {
        fechaObjetivo.setDate(fechaObjetivo.getDate() + frecuencia);
      }

      // Buscar el slot disponible (esta función ya gestiona excepciones y saturación)
      const slot = this.availabilityService.findNextAvailableSlot(fechaObjetivo, "2.2");

      const nuevaSesion = {
        PacienteID: ciclo.PacienteID,
        CicloID: ciclo.CicloID,
        Modalidad: "2.2",
        NumeroSesion: i + 1,
        FechaSesion: slot.fecha,
        HoraInicio: slot.hora,
        EstadoSesion: "PROGRAMADA"
      };

      sesionesGeneradas.push(nuevaSesion);
      ultimaFecha = slot.fecha; // La siguiente sesión cuenta a partir de esta
    }

    return sesionesGeneradas;
  }
}
