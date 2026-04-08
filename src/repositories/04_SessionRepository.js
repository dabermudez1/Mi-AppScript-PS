/**
 * Repositorio para la hoja SESIONES.
 */
class SessionRepository extends BaseRepository {
  constructor() {
    super(SHEET_SESIONES, HEADERS[SHEET_SESIONES]);
  }

  findByPacienteId(pacienteId) {
    const all = this.findAll();
    return all.filter(s => String(s.PacienteID) === String(pacienteId));
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