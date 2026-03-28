/**
 * AssignmentRepository
 * Maneja el registro histórico de asignaciones de pacientes a ciclos.
 */
class AssignmentRepository extends BaseRepository {
  constructor() {
    super(SHEET_ASIGNACIONES_CICLO, 'AsignacionID');
  }
}
const assignmentRepo = new AssignmentRepository();