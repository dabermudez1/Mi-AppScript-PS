/**
 * Repositorio para la hoja CICLOS.
 */
class CicloRepository extends BaseRepository {
  constructor() {
    super(SHEET_CICLOS, HEADERS[SHEET_CICLOS]);
  }

  findNextAvailable(modalidad, fechaPrimeraConsulta) {
    const todosLosCiclos = this.findAll();
    const hoy = normalizarFecha_(new Date());

    const ciclosFiltrados = todosLosCiclos.filter(c => {
      const fechaInicio = c.FechaInicioCiclo;
      return c.Modalidad === modalidad &&
             c.EstadoCiclo === ESTADOS_CICLO.PLANIFICADO &&
             (fechaInicio instanceof Date) &&
             (normalizarFecha_(fechaInicio).getTime() >= normalizarFecha_(fechaPrimeraConsulta).getTime()) &&
             (Number(c.PlazasLibres || 0) > 0);
    });

    return ciclosFiltrados.sort((a, b) => compararFechas_(a.FechaInicioCiclo, b.FechaInicioCiclo))[0] || null;
  }

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