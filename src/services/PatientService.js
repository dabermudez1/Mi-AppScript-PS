/**
 * Servicio para gestionar la lógica de negocio de Pacientes.
 */
class PatientService {
  constructor() {
    this.patientRepo = new PatientRepository();
    this.configRepo = new ConfigRepository();
  }

  /**
   * Lógica principal de alta de un nuevo paciente.
   */
  createPatient(data) {
    const config = this.configRepo.findByModalidad(data.modalidad);
    if (!config || !config.Activa) {
      throw new Error(`La modalidad no existe o está inactiva: ${data.modalidad}`);
    }

    if (data.modalidad === MODALIDADES.INDIVIDUAL) {
      return this._processIndividualHigh(data, config);
    } else {
      // Para grupos, de momento mantenemos la delegación al sistema anterior
      // mientras terminamos de migrar la lógica de ciclos.
      return null; 
    }
  }

  _processIndividualHigh(data, config) {
    const capacityInfo = this._hasIndividualCapacity(config);
    
    const patientData = {
      PacienteID: generarId_('PAC'),
      Nombre: data.nombre,
      NHC: data.nhc,
      SexoGenero: data.sexoGenero,
      MotivoConsultaDiagnostico: data.motivoConsultaDiagnostico,
      MotivoConsultaOtros: data.motivoConsultaOtros,
      ModalidadSolicitada: data.modalidad,
      FechaAlta: normalizarFecha_(new Date()),
      FechaPrimeraConsulta: normalizarFecha_(data.fechaPrimeraConsulta),
      SesionesPlanificadas: Number(config.SesionesPorCiclo || 7),
      SesionesCompletadas: 0,
      SesionesPendientes: Number(config.SesionesPorCiclo || 7),
      RecalcularSecuencia: false
    };

    if (!capacityInfo.hasCapacity) {
      patientData.EstadoPaciente = ESTADOS_PACIENTE.ESPERA;
      patientData.MotivoEspera = 'SIN_CAPACIDAD_INDIVIDUAL';
      this.patientRepo.save(patientData);
      
      return {
        pacienteId: patientData.PacienteID,
        mensaje: `Paciente ${data.nombre} en ESPERA por falta de capacidad.`
      };
    }

    const firstSessionDate = this._calculateFirstIndividualSession(data.fechaPrimeraConsulta, config);
    patientData.EstadoPaciente = ESTADOS_PACIENTE.ACTIVO;
    patientData.FechaPrimeraSesionReal = firstSessionDate;
    patientData.ProximaSesion = firstSessionDate;

    this.patientRepo.save(patientData);

    return {
      pacienteId: patientData.PacienteID,
      status: 'ACTIVE',
      firstSessionDate
    };
  }

  _hasIndividualCapacity(config) {
    const capacity = Number(config.CapacidadMaxima || 0);
    const activePatients = this.patientRepo.findAll().filter(p => 
      p.ModalidadSolicitada === MODALIDADES.INDIVIDUAL && 
      p.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO
    ).length;

    return {
      hasCapacity: activePatients < capacity
    };
  }

  _calculateFirstIndividualSession(date, config) {
    const interval = Number(config.FrecuenciaDias || 15);
    const base = sumarDiasNaturales_(date, interval);
    return ajustarASiguienteFechaOperativa_(base);
  }

  /**
   * Actualiza los datos básicos del paciente usando el repositorio.
   */
  updateBasicData(pacienteId, data) {
    const patient = this.patientRepo.findById(pacienteId);
    if (!patient) throw new Error("Paciente no encontrado.");

    // Solo actualizamos campos permitidos para no romper estados
    patient.Nombre = data.nombre || patient.Nombre;
    patient.NHC = data.nhc || patient.NHC;
    patient.Observaciones = data.observaciones || patient.Observaciones;

    this.patientRepo.save(patient);
    return patient;
  }
}