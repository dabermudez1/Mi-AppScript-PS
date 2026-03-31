/**
 * Repositorio para gestionar las vinculaciones entre Pacientes y Ciclos.
 */
class AsignacionRepository extends BaseRepository {
  constructor() {
    super(SHEET_ASIGNACIONES_CICLO, HEADERS[SHEET_ASIGNACIONES_CICLO]);
  }
}