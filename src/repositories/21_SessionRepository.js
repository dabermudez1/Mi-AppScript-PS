/**
 * 21_SessionRepository.js
 * Repositorio para la gestión de Sesiones.
 */
class SessionRepository extends BaseRepository {
  constructor() {
    super(SHEET_SESIONES, HEADERS[SHEET_SESIONES]);
  }

  findByPacienteId(pacienteId) {
    return this.findAll().filter(s => String(s.PacienteID) === String(pacienteId));
  }

  findByCicloId(cicloId) {
    return this.findAll().filter(s => String(s.CicloID) === String(cicloId));
  }

  findPendientesByPaciente(pacienteId) {
    return this.findByPacienteId(pacienteId).filter(s => 
      s.EstadoSesion === ESTADOS_SESION.PENDIENTE || 
      s.EstadoSesion === ESTADOS_SESION.REPROGRAMADA
    );
  }

  /**
   * Agrupa una lista de sesiones por PacienteID.
   */
  mapByPatient(sesiones) {
    const data = sesiones || this.findAll();
    const map = {};
    data.forEach(s => {
      if (!map[s.PacienteID]) map[s.PacienteID] = [];
      map[s.PacienteID].push(s);
    });
    return map;
  }
}