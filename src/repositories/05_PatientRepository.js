/**
 * Repositorio para la hoja PACIENTES.
 */
class PatientRepository extends BaseRepository {
  constructor() {
    super(SHEET_PACIENTES, HEADERS[SHEET_PACIENTES]);
  }

  findById(pacienteId) {
    const all = this.findAll();
    return all.find(p => String(p.PacienteID) === String(pacienteId)) || null;
  }

  findActivos() {
    return this.findAll().filter(p => p.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO);
  }

  findEnEspera() {
    return this.findAll().filter(p => p.EstadoPaciente === ESTADOS_PACIENTE.ESPERA);
  }
}