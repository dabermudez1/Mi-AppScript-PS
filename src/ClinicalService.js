/**
 * ClinicalService
 * Gestiona la lógica de la ficha clínica y la integridad de los datos de salud.
 */
class ClinicalService {
  /**
   * Obtiene la ficha completa asegurando que existan los datos y estén sincronizados.
   */
  static getFullFicha(pacienteId) {
    this._ensureFichaExists(pacienteId);
    this.syncFromPatient(pacienteId);
    return clinicalRepo.findById(pacienteId);
  }

  /**
   * Sincroniza campos automáticos desde la tabla de Pacientes hacia la Ficha Clínica.
   */
  static syncFromPatient(pacienteId) {
    const patient = patientRepo.findById(pacienteId);
    const ficha = clinicalRepo.findById(pacienteId);
    if (!patient || !ficha) return;

    ficha.Nombre = patient.Nombre;
    ficha.NHC = patient.NHC;
    ficha.FechaAltaPrograma = patient.FechaAlta;
    ficha.FechaPrimeraConsulta = patient.FechaPrimeraConsulta;
    ficha.EstadoPacienteActual = patient.EstadoPaciente;
    ficha.FechaAltaEfectiva = patient.FechaCierre || '';
    ficha.TipoIntervencionPrincipal = patient.ModalidadSolicitada === MODALIDADES.INDIVIDUAL ? 'Individual' : 'Grupal';
    
    if (patient.EstadoPaciente === ESTADOS_PACIENTE.ALTA) {
      ficha.FinTratamientoCodigo = patient.MotivoAltaCodigo || '';
      ficha.FinTratamientoTexto = patient.MotivoAltaTexto || '';
    } else {
      ficha.FinTratamientoCodigo = 7;
      ficha.FinTratamientoTexto = 'Activo en el programa';
    }

    clinicalRepo.save(ficha);
  }

  /**
   * Guarda los datos clínicos y actualiza los campos básicos en el repositorio de pacientes.
   */
  static saveFicha(formData) {
    const patientId = formData.pacienteId;
    const ficha = clinicalRepo.findById(patientId);
    if (!ficha) throw new Error("Ficha clínica no encontrada.");

    // Actualizamos campos de la ficha
    Object.assign(ficha, formData);
    clinicalRepo.save(ficha);

    // Sincronización inversa: Actualizamos datos básicos en el objeto Paciente
    const patient = patientRepo.findById(patientId);
    patient.NHC = formData.nhc;
    patient.SexoGenero = formData.sexoGenero;
    patient.MotivoConsultaDiagnostico = formData.motivoConsultaDiagnostico;
    patient.MotivoConsultaOtros = formData.motivoConsultaOtros;
    
    patientRepo.save(patient);
    return { success: true };
  }

  static _ensureFichaExists(pacienteId) {
    const existing = clinicalRepo.findById(pacienteId);
    if (!existing) {
      clinicalRepo.save({ 
        PacienteID: pacienteId,
        Nombre: '',
        NHC: '',
        EstadoPacienteActual: ESTADOS_PACIENTE.ESPERA
      });
    }
  }
}