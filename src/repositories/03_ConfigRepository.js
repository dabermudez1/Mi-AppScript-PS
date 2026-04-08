/**
 * Repositorio para la hoja CONFIG_MODALIDADES.
 */
class ConfigRepository extends BaseRepository {
  constructor() {
    super(SHEET_CONFIG_MODALIDADES, HEADERS[SHEET_CONFIG_MODALIDADES]);
  }

  findByModalidad(modalidad) {
    const all = this.findAll();
    return all.find(cfg => cfg.Modalidad === modality) || null;
  }
}