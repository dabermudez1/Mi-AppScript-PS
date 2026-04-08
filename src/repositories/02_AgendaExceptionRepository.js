/**
 * Repositorio para la hoja AGENDA_EXCEPCIONES.
 */
class AgendaExceptionRepository extends BaseRepository {
  constructor() {
    super(SHEET_AGENDA_EXCEPCIONES, HEADERS[SHEET_AGENDA_EXCEPCIONES]);
  }
}