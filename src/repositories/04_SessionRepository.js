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
}