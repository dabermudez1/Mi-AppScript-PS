/**
 * PatientRepository
 * Maneja la persistencia de los datos de pacientes.
 */
class PatientRepository extends BaseRepository {
  constructor() {
    // SHEET_PACIENTES está definido en 01_base.js
    super(SHEET_PACIENTES, 'PacienteID');
  }

  /**
   * Cuenta cuántos pacientes activos hay para una modalidad específica.
   * Reemplaza parte de la lógica ineficiente de hayCapacidadIndividual_.
   */
  countActiveByModality(modalidad) {
    return this.findAll().filter(p => 
      p.ModalidadSolicitada === modalidad && 
      p.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO
    ).length;
  }
}

// Instancia única para ser usada en los servicios
const patientRepo = new PatientRepository();