/**
 * PatientService
 * Orquestador de la lógica de negocio para pacientes.
 */
class PatientService {
  /**
   * Registra un nuevo paciente evaluando capacidad y calculando fechas.
   */
  static register(formData) {
    const modalidad = formData.modalidad;
    // obtenerConfigModalidad_ viene de 02_ciclos.js (se refactorizará luego)
    const config = obtenerConfigModalidad_(modalidad);
    
    if (modalidad === MODALIDADES.INDIVIDUAL) {
      return this._handleIndividualRegistration(formData, config);
    } else {
      // La lógica de grupos requiere el CycleRepository, que haremos a continuación.
      throw new Error("Lógica de grupos pendiente de migración a Service.");
    }
  }

  static _handleIndividualRegistration(formData, config) {
    const capacidadMaxima = Number(config.CapacidadMaxima || 0);
    const activos = patientRepo.countActiveByModality(MODALIDADES.INDIVIDUAL);
    const tieneCapacidad = activos < capacidadMaxima;

    const sesionesPlanificadas = Number(config.SesionesPorCiclo || 7);
    
    const nuevoPaciente = new PatientEntity({
      PacienteID: generarId_('PAC'),
      Nombre: formData.nombre,
      NHC: formData.nhc,
      SexoGenero: formData.sexoGenero,
      MotivoConsultaDiagnostico: formData.motivoConsultaDiagnostico,
      MotivoConsultaOtros: formData.motivoConsultaOtros,
      ModalidadSolicitada: formData.modalidad,
      FechaAlta: normalizarFecha_(new Date()),
      FechaPrimeraConsulta: parseFechaISO_(formData.fechaPrimeraConsulta),
      Observaciones: formData.observaciones || '',
      SesionesPlanificadas: sesionesPlanificadas,
      SesionesCompletadas: 0,
      SesionesPendientes: sesionesPlanificadas,
      RecalcularSecuencia: false
    });

    // Validación inmediata antes de cualquier lógica adicional
    nuevoPaciente.validate();

    if (!tieneCapacidad) {
      nuevoPaciente.EstadoPaciente = ESTADOS_PACIENTE.ESPERA;
      nuevoPaciente.MotivoEspera = 'SIN_CAPACIDAD_INDIVIDUAL';
      
      patientRepo.save(nuevoPaciente);
      
      return {
        success: true,
        pacienteId: nuevoPaciente.PacienteID,
        mensaje: `Paciente creado en ESPERA por falta de capacidad individual.`
      };
    }

    // Lógica para paciente ACTIVO
    const fechaInicio = calcularPrimeraSesionIndividual_(nuevoPaciente.FechaPrimeraConsulta, formData.modalidad);
    
    nuevoPaciente.EstadoPaciente = ESTADOS_PACIENTE.ACTIVO;
    nuevoPaciente.FechaPrimeraSesionReal = fechaInicio;
    nuevoPaciente.ProximaSesion = fechaInicio;

    patientRepo.save(nuevoPaciente);

    // Invocamos la generación de sesiones (Bloque 4)
    const resultadoSesiones = generarSesionesPacienteIndividual_(nuevoPaciente.PacienteID);

    return {
      success: true,
      pacienteId: nuevoPaciente.PacienteID,
      mensaje: `Paciente creado correctamente en estado ACTIVO.\nPrimera sesión: ${formatearFecha_(fechaInicio)}`,
      avisos: resultadoSesiones.avisos || []
    };
  }
  
  static delete(id) {
    // Aquí podríamos añadir lógica de "antes de borrar" (ej: cancelar eventos en Calendar)
    return patientRepo.deleteById(id);
  }
}