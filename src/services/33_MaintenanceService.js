/**
 * Servicio para tareas de mantenimiento y recálculo de métricas globales.
 */
class MaintenanceService {
  constructor() {
    this.cicloRepo = new CicloRepository();
    this.asignacionRepo = new AsignacionRepository();
  }

  /**
   * Recalcula la ocupación de todos los ciclos basándose en las asignaciones activas.
   * @returns {number} Número de ciclos actualizados.
   */
  recalculateCycleOccupancy() {
    const ciclos = this.cicloRepo.findAll();
    const asignaciones = this.asignacionRepo.findAll();

    if (ciclos.length === 0) {
      return 0;
    }

    const ocupacionPorCiclo = {};

    asignaciones.forEach(asignacion => {
      const cicloId = String(asignacion.CicloID || '');
      const estado = asignacion.EstadoAsignacion;

      if (!cicloId) return;

      if (
        estado === ESTADOS_ASIGNACION.ACTIVO ||
        estado === ESTADOS_ASIGNACION.RESERVADO
      ) {
        if (!ocupacionPorCiclo[cicloId]) {
          ocupacionPorCiclo[cicloId] = 0;
        }
        ocupacionPorCiclo[cicloId]++;
      }
    });

    let actualizados = 0;

    ciclos.forEach(ciclo => {
      const capacidad = Number(ciclo.CapacidadMaxima || 0);
      const ocupadas = ocupacionPorCiclo[ciclo.CicloID] || 0;
      const libres = Math.max(0, capacidad - ocupadas);

      if (ciclo.PlazasOcupadas !== ocupadas || ciclo.PlazasLibres !== libres) {
        ciclo.PlazasOcupadas = ocupadas;
        ciclo.PlazasLibres = libres;
        this.cicloRepo.save(ciclo);
        actualizados++;
      }
    });

    return actualizados;
  }
}