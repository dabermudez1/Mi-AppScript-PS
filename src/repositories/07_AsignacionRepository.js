/**
 * Repositorio para la hoja ASIGNACIONES_CICLO.
 */
class AsignacionRepository extends BaseRepository {
  constructor() {
    super(SHEET_ASIGNACIONES_CICLO, HEADERS[SHEET_ASIGNACIONES_CICLO]);
  }
}