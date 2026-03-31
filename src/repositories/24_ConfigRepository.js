/**
 * Repositorio para la configuración de modalidades.
 */
class ConfigRepository extends BaseRepository {
  constructor() {
    super(SHEET_CONFIG_MODALIDADES, HEADERS[SHEET_CONFIG_MODALIDADES]);
  }

  findByModalidad(modalidad) {
    return this.findOneBy('Modalidad', modalidad);
  }
}