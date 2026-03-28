/**
 * StatusService
 * Orquestador central para la actualización de estados y métricas.
 * Implementa la lógica de transición de estados de forma atómica y eficiente.
 */
class StatusService {
  /**
   * Ejecuta el proceso completo de actualización diaria.
   */
  static runDailyUpdate() {
    const hoy = normalizarFecha_(new Date());
    
    const resultados = {
      sesiones: this.updateExpiredSessions(hoy),
      ciclos: this.updateCycleStatuses(hoy),
      pacientes: this.recalculateAllPatientStates(hoy)
    };

    // Al final, sincronizamos las métricas de todos los pacientes activos
    this.syncAllActivePatientMetrics();
    
    // NUEVO: Sincronización con Google Calendar
    CalendarService.syncAll();
    
    return resultados;
  }

  /**
   * Pasa a COMPLETADA_AUTO las sesiones PENDIENTES de fechas pasadas.
   */
  static updateExpiredSessions(hoy) {
    const sesiones = sessionRepo.findAll();
    let cambios = 0;

    sesiones.forEach(s => {
      if (s.EstadoSesion === ESTADOS_SESION.PENDIENTE && 
          s.FechaSesion instanceof Date && 
          s.FechaSesion.getTime() < hoy.getTime()) {
        s.EstadoSesion = ESTADOS_SESION.COMPLETADA_AUTO;
        sessionRepo.save(s);
        cambios++;
      }
    });
    return cambios;
  }

  /**
   * Gestiona el ciclo de vida de los ciclos (PLANIFICADO -> EN_CURSO -> CERRADO).
   */
  static updateCycleStatuses(hoy) {
    const ciclos = cycleRepo.findAll();
    let aEnCurso = 0;
    let aCerrado = 0;

    ciclos.forEach(c => {
      const inicio = c.FechaInicioCiclo instanceof Date ? c.FechaInicioCiclo.getTime() : 0;
      const fin = c.FechaFinCiclo instanceof Date ? c.FechaFinCiclo.getTime() : 0;
      const tHoy = hoy.getTime();

      if (c.EstadoCiclo === ESTADOS_CICLO.PLANIFICADO && inicio <= tHoy) {
        c.EstadoCiclo = ESTADOS_CICLO.EN_CURSO;
        cycleRepo.save(c);
        aEnCurso++;
      } else if (c.EstadoCiclo === ESTADOS_CICLO.EN_CURSO && fin < tHoy) {
        c.EstadoCiclo = ESTADOS_CICLO.CERRADO;
        cycleRepo.save(c);
        aCerrado++;
      }
    });
    return { aEnCurso, aCerrado };
  }

  /**
   * Recalcula estados de pacientes basándose en sus sesiones y ciclos.
   */
  static recalculateAllPatientStates(hoy) {
    const pacientes = patientRepo.findAll();
    const sesiones = sessionRepo.findAll();
    let actualizados = 0;

    pacientes.forEach(p => {
      if (p.EstadoPaciente === ESTADOS_PACIENTE.ALTA) return;

      const misSesiones = sesiones.filter(s => s.PacienteID === p.PacienteID);
      const stats = this._analyzeSessions(misSesiones, hoy);
      
      let nuevoEstado = p.EstadoPaciente;

      // Lógica de transición
      if (p.ModalidadSolicitada === MODALIDADES.INDIVIDUAL) {
        nuevoEstado = (stats.total > 0) ? ESTADOS_PACIENTE.ACTIVO : ESTADOS_PACIENTE.ESPERA;
      } else {
        const tieneCiclo = !!(p.CicloObjetivoID || p.CicloActivoID);
        if (!tieneCiclo) {
          nuevoEstado = ESTADOS_PACIENTE.ESPERA;
        } else if (stats.completadas > 0 || stats.vencidas > 0) {
          nuevoEstado = ESTADOS_PACIENTE.ACTIVO;
        } else {
          nuevoEstado = ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO;
        }
      }

      // Auto-alta si terminaron todas las sesiones
      if (p.SesionesPlanificadas > 0 && stats.completadas >= p.SesionesPlanificadas) {
        nuevoEstado = ESTADOS_PACIENTE.ALTA;
        p.FechaCierre = hoy;
      }

      if (nuevoEstado !== p.EstadoPaciente) {
        p.EstadoPaciente = nuevoEstado;
        patientRepo.save(p);
        actualizados++;
      }
    });
    return actualizados;
  }

  /** @private */
  static _analyzeSessions(sesiones, hoy) {
    return {
      total: sesiones.length,
      completadas: sesiones.filter(s => s.EstadoSesion.startsWith('COMPLETADA')).length,
      vencidas: sesiones.filter(s => s.EstadoSesion === ESTADOS_SESION.PENDIENTE && s.FechaSesion.getTime() < hoy.getTime()).length,
      proxima: sesiones.filter(s => s.EstadoSesion === ESTADOS_SESION.PENDIENTE && s.FechaSesion.getTime() >= hoy.getTime())
                       .sort((a,b) => a.FechaSesion - b.FechaSesion)[0]?.FechaSesion || null
    };
  }

  /**
   * Sincroniza contadores y fechas de próxima sesión en la hoja PACIENTES.
   */
  static syncAllActivePatientMetrics() {
    const pacientes = patientRepo.findAll();
    const sesiones = sessionRepo.findAll();
    const hoy = normalizarFecha_(new Date());

    pacientes.forEach(p => {
      const misSesiones = sesiones.filter(s => s.PacienteID === p.PacienteID);
      const stats = this._analyzeSessions(misSesiones, hoy);
      
      p.SesionesCompletadas = stats.completadas;
      p.SesionesPendientes = p.SesionesPlanificadas - stats.completadas;
      p.ProximaSesion = stats.proxima;
      
      patientRepo.save(p);
    });
  }
}