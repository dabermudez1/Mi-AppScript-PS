/**
 * @file 20_PatientRepository.js
 * Repositorio específico para Pacientes.
 */
class PatientRepository extends BaseRepository {
  constructor() {
    super(SHEET_PACIENTES, HEADERS[SHEET_PACIENTES]);
  }

  /**
   * Busca un paciente por su ID único.
   * @param {string} pacienteId 
   * @returns {Object|null}
   */
  findById(pacienteId) {
    return this.findOneBy('PacienteID', pacienteId);
  }
  
  /**
   * Obtiene todos los pacientes con estado ACTIVO.
   * @returns {Object[]}
   */
  findActivos() {
    return this.findAll().filter(p => p.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO);
  }

  /**
   * Obtiene todos los pacientes que están actualmente en lista de espera.
   * @returns {Object[]}
   */
  findEnEspera() {
    return this.findAll().filter(p => p.EstadoPaciente === ESTADOS_PACIENTE.ESPERA);
  }

  /**
   * Busca pacientes por una modalidad específica (INDIVIDUAL, GRUPO_1, etc).
   */
  findPorModalidad(modalidad) {
    return this.findAll().filter(p => p.ModalidadSolicitada === modalidad);
  }
}