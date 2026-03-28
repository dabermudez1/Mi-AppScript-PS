/**
 * CycleRepository
 * Maneja la persistencia y consultas de los ciclos de terapia grupal.
 */
class CycleRepository extends BaseRepository {
  constructor() {
    super(SHEET_CICLOS, 'CicloID');
  }

  /**
   * Busca el primer ciclo planificado disponible para una modalidad después de una fecha.
   */
  findAvailableCycle(modalidad, fechaConsulta) {
    const fechaMin = normalizarFecha_(fechaConsulta);
    
    return this.findAll()
      .filter(c => 
        c.Modalidad === modalidad && 
        c.EstadoCiclo === ESTADOS_CICLO.PLANIFICADO &&
        c.FechaInicioCiclo instanceof Date &&
        c.FechaInicioCiclo.getTime() > fechaMin.getTime() &&
        Number(c.PlazasLibres || 0) > 0
      )
      .sort((a, b) => a.FechaInicioCiclo - b.FechaInicioCiclo)[0] || null;
  }

  /**
   * Verifica si existe algún ciclo futuro (aunque esté lleno) para determinar el motivo de espera.
   */
  existsFutureCycle(modalidad, fechaConsulta) {
    const fechaMin = normalizarFecha_(fechaConsulta);
    return this.findAll().some(c => 
      c.Modalidad === modalidad && 
      c.EstadoCiclo === ESTADOS_CICLO.PLANIFICADO &&
      c.FechaInicioCiclo instanceof Date &&
      c.FechaInicioCiclo.getTime() > fechaMin.getTime()
    );
  }

  /**
   * Actualiza las plazas de un ciclo de forma segura.
   * @returns {boolean} True si se pudo actualizar, false si no había capacidad.
   */
  updatePlazas(cicloId, delta) {
    const ciclo = this.findById(cicloId);
    if (!ciclo) throw new Error(`Ciclo ${cicloId} no encontrado.`);

    const nuevasOcupadas = Number(ciclo.PlazasOcupadas || 0) + delta;
    if (nuevasOcupadas < 0) throw new Error("Las plazas ocupadas no pueden ser negativas.");
    if (nuevasOcupadas > ciclo.CapacidadMaxima && delta > 0) return false;

    ciclo.PlazasOcupadas = nuevasOcupadas;
    ciclo.PlazasLibres = Math.max(0, ciclo.CapacidadMaxima - nuevasOcupadas);
    
    this.save(ciclo);
    return true;
  }
}
const cycleRepo = new CycleRepository();