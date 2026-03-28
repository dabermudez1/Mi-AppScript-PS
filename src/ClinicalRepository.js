/**
 * ClinicalRepository
 * Maneja la persistencia de la ficha clínica detallada de los pacientes.
 */
class ClinicalRepository extends BaseRepository {
  constructor() {
    super('DATOS_CLINICOS_PACIENTES', 'PacienteID');
  }
}
const clinicalRepo = new ClinicalRepository();