/**
 * SessionRepository
 * Maneja la persistencia de las sesiones de terapia.
 */
class SessionRepository extends BaseRepository {
  constructor() {
    super(SHEET_SESIONES, 'SesionID');
  }
}
const sessionRepo = new SessionRepository();