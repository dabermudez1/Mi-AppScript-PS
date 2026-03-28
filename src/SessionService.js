/**
 * SessionService
 * Lógica de negocio para la planificación y gestión de sesiones.
 */
class SessionService {
  /**
   * Genera el cronograma de sesiones para un paciente individual.
   */
  static generateForIndividual(patientId) {
    const patient = patientRepo.findById(patientId);
    if (!patient || patient.EstadoPaciente !== ESTADOS_PACIENTE.ACTIVO) return { avisos: [] };

    const config = obtenerConfigModalidad_(patient.ModalidadSolicitada);
    const intervaloDias = Number(config.FrecuenciaDias || 0);

    const resultado = generarFechasIndividualConAvisos_({
      fechaInicio: patient.FechaPrimeraSesionReal,
      intervaloDias: intervaloDias,
      sesiones: patient.SesionesPlanificadas
    });

    this._bulkCreate(patient, resultado.fechas);
    return { avisos: resultado.avisos || [] };
  }

  /**
   * Genera el cronograma de sesiones para un paciente de grupo basándose en el ciclo.
   */
  static generateForGroup(patientId, cicloId) {
    const patient = patientRepo.findById(patientId);
    const ciclo = cycleRepo.findById(cicloId);
    if (!patient || !ciclo) throw new Error('Paciente o ciclo no encontrado.');

    const resultado = generarFechasCiclo_({
      fechaInicio: ciclo.FechaInicioCiclo,
      diaSemana: ciclo.DiaSemana,
      frecuenciaDias: ciclo.FrecuenciaDias,
      sesiones: ciclo.SesionesPorCiclo
    });

    this._bulkCreate(patient, resultado.fechas, cicloId);
    return { avisos: resultado.avisos || [] };
  }

  /**
   * Crea las entidades de sesión y las persiste de forma masiva.
   * @private
   */
  static _bulkCreate(patient, fechas, cicloId = '') {
    const sessions = fechas.map((fecha, index) => ({
      SesionID: generarId_('SES'),
      PacienteID: patient.PacienteID,
      CicloID: cicloId,
      AsignacionId: '',
      Modalidad: patient.ModalidadSolicitada,
      NombrePaciente: patient.Nombre,
      NumeroSesion: index + 1,
      FechaSesion: fecha,
      EstadoSesion: ESTADOS_SESION.PENDIENTE,
      FechaOriginal: fecha,
      ModificadaManual: false,
      Notas: '',
      CalendarEventId: '',
      CalendarSyncStatus: 'PENDIENTE',
      CalendarLastSync: '',
      CalendarEventTitle: '',
      CalendarHash: ''
    }));

    sessionRepo.insertMany(sessions);
  }
}