/**
 * Servicio para la gestión de transiciones de estado automáticas.
 */
class StateService {
  constructor() {
    this.patientRepo = new PatientRepository();
    this.cicloRepo = new CicloRepository();
    this.sessionRepo = new SessionRepository();
    this.asignacionRepo = new AsignacionRepository();
  }

  /**
   * Ejecuta el proceso completo de actualización de estados.
   */
  runAutomaticTransitions() {
    const hoy = normalizarFecha_(new Date());
    let stats = { ciclos: 0, pacientes: 0, sesiones: 0 };

    // 1. Ciclos: PLANIFICADO -> EN_CURSO -> CERRADO
    const ciclos = this.cicloRepo.findAll();
    ciclos.forEach(c => {
      const fechaInicio = normalizarFecha_(new Date(c.FechaInicioCiclo));
      const fechaFin = normalizarFecha_(new Date(c.FechaFinCiclo));

      if (c.EstadoCiclo === ESTADOS_CICLO.PLANIFICADO && fechaInicio <= hoy) {
        c.EstadoCiclo = ESTADOS_CICLO.EN_CURSO;
        this.cicloRepo.save(c);
        stats.ciclos++;
      } else if ((c.EstadoCiclo === ESTADOS_CICLO.EN_CURSO || c.EstadoCiclo === ESTADOS_CICLO.PLANIFICADO) && fechaFin < hoy) {
        c.EstadoCiclo = ESTADOS_CICLO.CERRADO;
        this.cicloRepo.save(c);
        stats.ciclos++;
      }
    });

    // 2. Sesiones: PENDIENTE -> COMPLETADA_AUTO (si la fecha pasó)
    const sesiones = this.sessionRepo.findAll();
    sesiones.forEach(s => {
      const fechaSesion = normalizarFecha_(new Date(s.FechaSesion));
      if (s.EstadoSesion === ESTADOS_SESION.PENDIENTE && fechaSesion < hoy) {
        s.EstadoSesion = ESTADOS_SESION.COMPLETADA_AUTO;
        this.sessionRepo.save(s);
        stats.sesiones++;
      }
    });

    // 3. Pacientes: ACTIVO_PENDIENTE_INICIO -> ACTIVO (si su ciclo empezó)
    // Y ACTIVO -> ALTA (si terminaron sesiones)
    const pacientes = this.patientRepo.findAll();
    pacientes.forEach(p => {
      if (p.EstadoPaciente === ESTADOS_PACIENTE.ALTA) return;

      // Lógica de inicio de ciclo
      if (p.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO && p.CicloObjetivoID) {
        const ciclo = ciclos.find(c => c.CicloID === p.CicloObjetivoID);
        if (ciclo && ciclo.EstadoCiclo === ESTADOS_CICLO.EN_CURSO) {
          p.EstadoPaciente = ESTADOS_PACIENTE.ACTIVO;
          p.CicloActivoID = p.CicloObjetivoID;
          this.patientRepo.save(p);
          stats.pacientes++;
        }
      }

      // Lógica de fin de sesiones (Auto-Alta)
      this.refreshPatientMetrics(p);
      if (p.SesionesPlanificadas > 0 && p.SesionesCompletadas >= p.SesionesPlanificadas) {
        p.EstadoPaciente = ESTADOS_PACIENTE.ALTA;
        p.FechaCierre = hoy;
        this.patientRepo.save(p);
        stats.pacientes++;
      }
    });

    return stats;
  }

  /**
   * Recalcula sesiones completadas/pendientes para un objeto paciente.
   */
  refreshPatientMetrics(patient) {
    const sesiones = this.sessionRepo.findByPacienteId(patient.PacienteID);
    patient.SesionesCompletadas = sesiones.filter(s => 
      s.EstadoSesion === ESTADOS_SESION.COMPLETADA_AUTO || s.EstadoSesion === ESTADOS_SESION.COMPLETADA_MANUAL
    ).length;
    patient.SesionesPendientes = sesiones.filter(s => 
      s.EstadoSesion === ESTADOS_SESION.PENDIENTE || s.EstadoSesion === ESTADOS_SESION.REPROGRAMADA
    ).length;
    
    const proximas = sesiones.filter(s => (s.EstadoSesion === ESTADOS_SESION.PENDIENTE || s.EstadoSesion === ESTADOS_SESION.REPROGRAMADA) && s.FechaSesion instanceof Date);
    if (proximas.length > 0) {
      patient.ProximaSesion = proximas.sort((a,b) => a.FechaSesion - b.FechaSesion)[0].FechaSesion;
    }
  }
}