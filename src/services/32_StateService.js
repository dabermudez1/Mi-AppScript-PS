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
	const props = PropertiesService.getUserProperties();
    props.setProperty('TASK_UPDATE_STATES_PROGRESS', '2');
    const hoy = normalizarFecha_(new Date());
    let stats = { ciclos: 0, pacientes: 0, sesiones: 0 };
    props.setProperty('TASK_UPDATE_STATES_PROGRESS', '5');

    // 1. Ciclos: PLANIFICADO -> EN_CURSO -> CERRADO
    const ciclos = this.cicloRepo.findAll();
    props.setProperty('TASK_UPDATE_STATES_PROGRESS', '15');

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
    const totalSesiones = sesiones.length || 1;
    props.setProperty('TASK_UPDATE_STATES_PROGRESS', '30');

    sesiones.forEach((s, i) => {
      if (i % 10 === 0) {
        const p = Math.round(30 + (i / totalSesiones * 20));
        props.setProperty('TASK_UPDATE_STATES_PROGRESS', p.toString());
      }
      const fechaSesion = normalizarFecha_(new Date(s.FechaSesion));
      const sesionDateTime = normalizarFechaHora_(s.FechaSesion, s.HoraInicio);
      if (s.EstadoSesion === ESTADOS_SESION.PENDIENTE && compararFechasHoras_(sesionDateTime, hoy) < 0) {
        s.EstadoSesion = ESTADOS_SESION.COMPLETADA_AUTO;
        this.sessionRepo.save(s);
        stats.sesiones++;
      }
    });

    // 3. Pacientes: ACTIVO_PENDIENTE_INICIO -> ACTIVO (si su ciclo empezó)
    // Y ACTIVO -> ALTA (si terminaron sesiones)
    const pacientes = this.patientRepo.findAll();
    const totalPacientes = pacientes.length;
    props.setProperty('TASK_UPDATE_STATES_PROGRESS', '50');

    // OPTIMIZACIÓN: Precargar sesiones y agruparlas por PacienteID en un solo paso
    const sesionesPorPaciente = this._mapSessionsByPatient();

    pacientes.forEach((p, i) => {
      if (i % 5 === 0) props.setProperty('TASK_UPDATE_STATES_PROGRESS', Math.round(50 + (i / totalPacientes * 45)).toString());
      if (p.EstadoPaciente === ESTADOS_PACIENTE.ALTA) return;

      const susSesiones = sesionesPorPaciente[p.PacienteID] || [];

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
      this.refreshPatientMetrics(p, susSesiones);
      if (p.SesionesPlanificadas > 0 && p.SesionesCompletadas >= p.SesionesPlanificadas) {
        p.EstadoPaciente = ESTADOS_PACIENTE.ALTA;
        p.FechaCierre = hoy;
        this.patientRepo.save(p);
        stats.pacientes++;
      }
    });
	props.setProperty('TASK_UPDATE_STATES_PROGRESS', '100');
    return stats;
  }

  /**
   * Recalcula sesiones completadas/pendientes para un objeto paciente.
   * @param {Object} patient - El objeto paciente.
   * @param {Array} [providedSessions] - Opcional: Lista de sesiones ya filtrada.
   */
  refreshPatientMetrics(patient, providedSessions) {
    const sesiones = providedSessions || this.sessionRepo.findByPacienteId(patient.PacienteID);
    patient.SesionesCompletadas = sesiones.filter(s => 
      s.EstadoSesion === ESTADOS_SESION.COMPLETADA_AUTO || s.EstadoSesion === ESTADOS_SESION.COMPLETADA_MANUAL
    ).length;
    patient.SesionesPendientes = sesiones.filter(s => 
      s.EstadoSesion === ESTADOS_SESION.PENDIENTE || s.EstadoSesion === ESTADOS_SESION.REPROGRAMADA
    ).length;

    const proximas = sesiones.filter(s =>
      (s.EstadoSesion === ESTADOS_SESION.PENDIENTE || s.EstadoSesion === ESTADOS_SESION.REPROGRAMADA) &&
      s.FechaSesion instanceof Date && s.HoraInicio
    ).map(s => ({ ...s, fullDateTime: normalizarFechaHora_(s.FechaSesion, s.HoraInicio) }));
    if (proximas.length > 0) {
      patient.ProximaSesion = proximas.sort((a,b) => compararFechasHoras_(a.fullDateTime, b.fullDateTime))[0].fullDateTime;
    }
  }

  /**
   * Helper para agrupar todas las sesiones por PacienteID
   * @private
   */
  _mapSessionsByPatient() {
    const all = this.sessionRepo.findAll();
    const map = {};
    all.forEach(s => {
      if (!map[s.PacienteID]) map[s.PacienteID] = [];
      map[s.PacienteID].push(s);
    });
    return map;
  }
}