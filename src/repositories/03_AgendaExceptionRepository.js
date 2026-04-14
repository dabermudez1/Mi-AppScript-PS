/**
 * Repositorio para la hoja AGENDA_EXCEPCIONES.
 * Gestiona bloqueos puntuales, festivos y cambios de horario específicos.
 */
class AgendaExceptionRepository extends BaseRepository {
  constructor() {
    super(SHEET_AGENDA_EXCEPCIONES, HEADERS[SHEET_AGENDA_EXCEPCIONES]);
  }

  /**
   * Normaliza los tipos de slot de las excepciones para coherencia lógica.
   */
  findAll() {
    const data = super.findAll();
    return data.map(ex => ({
      ...ex,
      TipoSlot: String(ex.TipoSlot || '').trim().toUpperCase()
    }));
  }
}