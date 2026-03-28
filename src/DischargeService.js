/**
 * DischargeService
 * Gestiona el proceso de salida de un paciente del sistema.
 */
class DischargeService {
  /**
   * Ejecuta el proceso de alta completo.
   */
  static execute(formData) {
    const patientId = formData.pacienteId;
    const patient = patientRepo.findById(patientId);
    
    if (!patient || patient.EstadoPaciente === ESTADOS_PACIENTE.ALTA) {
      throw new Error('El paciente no existe o ya está en estado ALTA.');
    }

    const cicloId = patient.CicloActivoID || patient.CicloObjetivoID;

    // 1. Eliminar sesiones que aún no han ocurrido
    const sessions = sessionRepo.findAll().filter(s => 
      s.PacienteID === patientId && 
      (s.EstadoSesion === ESTADOS_SESION.PENDIENTE || s.EstadoSesion === ESTADOS_SESION.REPROGRAMADA)
    );
    sessions.forEach(s => sessionRepo.deleteById(s.SesionID));

    // 2. Finalizar asignación de ciclo
    const assignments = assignmentRepo.findAll().filter(a => 
      a.PacienteID === patientId && 
      (a.EstadoAsignacion === ESTADOS_ASIGNACION.RESERVADO || a.EstadoAsignacion === ESTADOS_ASIGNACION.ACTIVO)
    );
    assignments.forEach(a => {
      a.EstadoAsignacion = ESTADOS_ASIGNACION.FINALIZADO;
      a.Observaciones = 'Alta paciente: ' + (formData.motivoTexto || '');
      assignmentRepo.save(a);
    });

    // 3. Actualizar paciente a estado ALTA
    patient.EstadoPaciente = ESTADOS_PACIENTE.ALTA;
    patient.FechaCierre = parseFechaES_(formData.fechaAlta);
    patient.MotivoAltaCodigo = Number(formData.motivoCodigo);
    patient.MotivoAltaTexto = formData.motivoTexto;
    patient.ComentarioAlta = formData.comentario;
    patient.ProximaSesion = '';
    patient.SesionesPendientes = 0;
    patient.CicloObjetivoID = '';
    patient.CicloActivoID = '';

    patientRepo.save(patient);
    return { success: true };
  }
}