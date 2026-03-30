/**
 * Repositorio para la gestión de Ciclos de terapia grupal.
 */
class CicloRepository extends BaseRepository {
  constructor() {
    super(SHEET_CICLOS, HEADERS[SHEET_CICLOS]);
  }

  /**
   * Busca el primer ciclo planificado en el futuro con plazas libres.
   */
  findNextAvailable(modalidad, fechaReferencia) {
    const ref = normalizarFecha_(fechaReferencia);
    
    return this.findAll()
      .filter(c => 
        c.Modalidad === modalidad &&
        c.EstadoCiclo === ESTADOS_CICLO.PLANIFICADO &&
        c.FechaInicioCiclo instanceof Date &&
        normalizarFecha_(c.FechaInicioCiclo) > ref &&
        Number(c.PlazasLibres) > 0
      )
      .sort((a, b) => a.FechaInicioCiclo - b.FechaInicioCiclo)[0] || null;
  }

  /**
   * Verifica si existe algún ciclo (aunque esté lleno) en el futuro.
   * Útil para diferenciar el motivo de espera.
   */
  existsPlannedInFuture(modalidad, fechaReferencia) {
    const ref = normalizarFecha_(fechaReferencia);
    return this.findAll().some(c => 
      c.Modalidad === modalidad &&
      c.EstadoCiclo === ESTADOS_CICLO.PLANIFICADO &&
      c.FechaInicioCiclo instanceof Date &&
      normalizarFecha_(c.FechaInicioCiclo) > ref
    );
  }
}