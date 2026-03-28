/**
 * SessionRepository
 * Maneja la persistencia de las sesiones de terapia.
 */
class SessionRepository extends BaseRepository {
  constructor() {
    super(SHEET_SESIONES, 'SesionID');
  }
}
var sessionRepo;
if (!sessionRepo) {
  sessionRepo = new SessionRepository();
}