/**
 * Repositorio para la hoja AGENDA_PLANTILLA.
 * Gestiona la disponibilidad semanal base.
 */
class AgendaTemplateRepository extends BaseRepository {
  constructor() {
    super(SHEET_AGENDA_PLANTILLA, HEADERS[SHEET_AGENDA_PLANTILLA]);
  }

  /**
   * Sobrescribimos findAll para asegurar la normalización de strings
   * necesaria para el motor de disponibilidad.
   */
  findAll() {
    const data = super.findAll();
    return data.map(slot => ({
      ...slot,
      DiaSemana: String(slot.DiaSemana || '').trim().toUpperCase(),
      TipoSlot: String(slot.TipoSlot || '').trim().toUpperCase()
    }));
  }
}