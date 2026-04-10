/**
 * Servicio para la lógica de negocio de Sesiones.
 */
class SessionService {
  constructor() {
    this.sessionRepo = new SessionRepository();
  }

  /**
   * Reprograma una sesión específica.
   */
  rescheduleSession(pacienteId, numeroSesion, nuevaFecha, nuevaHora) {
    const sesiones = this.sessionRepo.findPendientesByPaciente(pacienteId);
    const sesion = sesiones.find(s => Number(s.NumeroSesion) === Number(numeroSesion));

    if (!sesion) throw new Error("No se encontró la sesión pendiente.");

    // Lógica de negocio: guardar fecha original si es la primera vez que se cambia
    if (!sesion.FechaOriginal) {
      sesion.FechaOriginal = sesion.FechaSesion;
    }

    sesion.FechaSesion = normalizarFecha_(nuevaFecha);
    sesion.HoraInicio = nuevaHora || formatearHora_(sesion.HoraInicio);
    sesion.ModificadaManual = true;
    sesion.EstadoSesion = ESTADOS_SESION.REPROGRAMADA;

    this.sessionRepo.save(sesion);
    return sesion;
  }

  /**
   * Cancela todas las sesiones pendientes de un paciente.
   */
  cancelPendingSessions(pacienteId) {
    const pendientes = this.sessionRepo.findPendientesByPaciente(pacienteId);
    pendientes.forEach(s => {
      s.EstadoSesion = ESTADOS_SESION.CANCELADA;
      this.sessionRepo.save(s);
    });
    return pendientes.length;
  }

  /**
   * Crea sesiones masivamente para un paciente (Alta inicial).
   */
  createInitialSessions(paciente, fechas, cicloId = '') {
    const sesiones = fechas.map((fecha, index) => {
      const nuevaSesion = {
        SesionID: generarId_('SES'),
        PacienteID: paciente.PacienteID,
        CicloID: cicloId,
        AsignacionID: '',
        Modalidad: paciente.ModalidadSolicitada,
        NombrePaciente: paciente.Nombre,
        NumeroSesion: index + 1,
        FechaSesion: normalizarFecha_(fecha),
        EstadoSesion: ESTADOS_SESION.PENDIENTE,
        FechaOriginal: normalizarFecha_(fecha),
        ModificadaManual: false,
        Notas: '',
        CalendarEventId: '',
        CalendarSyncStatus: 'PENDIENTE',
        CalendarLastSync: '',
        CalendarEventTitle: '',
        CalendarHash: '',
        HoraInicio: formatearHora_(fecha)
      };
      return nuevaSesion;
    });

    if (sesiones.length > 0) {
      this.sessionRepo.saveAll(sesiones);
    }
    return sesiones.length;
  }
}