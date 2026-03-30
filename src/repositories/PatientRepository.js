/**
 * Repositorio específico para Pacientes.
 */
class PatientRepository extends BaseRepository {
  constructor() {
    super(SHEET_PACIENTES, HEADERS[SHEET_PACIENTES]);
  }

  findById(pacienteId) {
    return this.findOneBy('PacienteID', pacienteId);
  }
  
  // Aquí se pueden añadir métodos específicos como findActivos(), findEnEspera(), etc.
}