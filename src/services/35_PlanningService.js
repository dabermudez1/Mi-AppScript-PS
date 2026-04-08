/**
 * Orquesta la generación de sesiones del ciclo.
 */
class PlanningService {
  constructor() {
    this.availabilityService = new AvailabilityService(new AgendaRepository(), new SessionRepository());
  }

  planificarCicloSeguimiento(ciclo, numSesiones, frecuencia) {
    let ultimaFecha = new Date(ciclo.fechaInicio);
    const sesiones = [];

    for (let i = 0; i < numSesiones; i++) {
      let fechaObjetivo = new Date(ultimaFecha);
      if (i > 0) fechaObjetivo.setDate(fechaObjetivo.getDate() + frecuencia);

      const slot = this.availabilityService.findNextAvailableSlot(fechaObjetivo, "2.2");
      sesiones.push({
        NumeroSesion: i + 1,
        FechaSesion: slot.fecha,
        HoraInicio: slot.hora,
        Modalidad: "2.2"
      });
      ultimaFecha = slot.fecha;
    }
    return sesiones;
  }
}
