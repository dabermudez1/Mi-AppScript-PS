/**
 * Repositorio para la hoja AGENDA_PLANTILLA.
 */
class AgendaTemplateRepository extends BaseRepository {
  constructor() {
    super(SHEET_AGENDA_PLANTILLA, HEADERS[SHEET_AGENDA_PLANTILLA]);
  }
}