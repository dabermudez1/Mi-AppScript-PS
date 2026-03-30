/**
 * Servicio para gestionar la lógica de negocio de Pacientes.
 */
class PatientService {
  constructor() {
    this.patientRepo = new PatientRepository();
    this.configRepo = new ConfigRepository();
    this.cicloRepo = new CicloRepository();
    this.asignacionRepo = new AsignacionRepository();
    this.sessionService = new SessionService();
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
      return this._processGroupHigh(data, config);
    }
  }

  /**
   * Procesa el alta médica (Bloque 17).
   */
  dischargePatient(pacienteId, data) {
    const patient = this.patientRepo.findById(pacienteId);
    if (!patient) throw new Error("Paciente no encontrado.");

    // 1. Cancelar sesiones futuras
    this.sessionService.cancelPendingSessions(pacienteId);

    // 2. Finalizar asignaciones de ciclo si existen
    const asignaciones = this.asignacionRepo.findAll().filter(a => 
      String(a.PacienteID) === String(pacienteId) && 
      (a.EstadoAsignacion === 'ACTIVO' || a.EstadoAsignacion === 'RESERVADO')
    );
    
    asignaciones.forEach(a => {
      a.EstadoAsignacion = 'FINALIZADO';
      a.Observaciones = 'Alta paciente';
      this.asignacionRepo.save(a);
    });

    // 3. Actualizar estado del paciente
    patient.EstadoPaciente = ESTADOS_PACIENTE.ALTA;
    patient.FechaCierre = data.fechaAlta;
    patient.FechaAltaEfectiva = data.fechaAlta;
    patient.MotivoAltaCodigo = data.motivoCodigo;
    patient.MotivoAltaTexto = data.motivoTexto;
    patient.ComentarioAlta = data.comentario;
    patient.ProximaSesion = '';
    patient.SesionesPendientes = 0;
    patient.CicloObjetivoID = '';
    patient.CicloActivoID = '';

    this.patientRepo.save(patient);
    return patient;
  }

  _processGroupHigh(data, config) {
    const ciclo = this.cicloRepo.findNextAvailable(data.modalidad, data.fechaPrimeraConsulta);
    
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
      SesionesCompletadas: 0,
      RecalcularSecuencia: false
    };

    if (!ciclo) {
      const motivo = this.cicloRepo.existsPlannedInFuture(data.modalidad, data.fechaPrimeraConsulta)
        ? 'SIN_PLAZA_CICLO'
        : 'SIN_CICLO_DISPONIBLE';

      patientData.EstadoPaciente = ESTADOS_PACIENTE.ESPERA;
      patientData.MotivoEspera = motivo;
      patientData.SesionesPlanificadas = Number(config.SesionesPorCiclo || 7);
      patientData.SesionesPendientes = patientData.SesionesPlanificadas;
      
      this.patientRepo.save(patientData);
      return { pacienteId: patientData.PacienteID, status: 'WAITING', motivo };
    }

    // Intentar reservar plaza
    ciclo.PlazasOcupadas = Number(ciclo.PlazasOcupadas) + 1;
    ciclo.PlazasLibres = Number(ciclo.CapacidadMaxima) - ciclo.PlazasOcupadas;
    this.cicloRepo.save(ciclo);

    // Datos de paciente activo en ciclo
    patientData.EstadoPaciente = ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO;
    patientData.CicloObjetivoID = ciclo.CicloID;
    patientData.FechaPrimeraSesionReal = ciclo.FechaInicioCiclo;
    patientData.SesionesPlanificadas = Number(ciclo.SesionesPorCiclo || config.SesionesPorCiclo || 7);
    patientData.SesionesPendientes = patientData.SesionesPlanificadas;
    patientData.ProximaSesion = ciclo.FechaInicioCiclo;

    this.patientRepo.save(patientData);

    // Crear la asignación técnica
    this.asignacionRepo.save({
      AsignacionID: generarId_('ASI'),
      PacienteID: patientData.PacienteID,
      CicloID: ciclo.CicloID,
      Modalidad: data.modalidad,
      FechaAsignacion: normalizarFecha_(new Date()),
      EstadoAsignacion: ESTADOS_ASIGNACION.RESERVADO
    });

    return {
      pacienteId: patientData.PacienteID,
      status: 'ACTIVE_PENDING',
      cicloId: ciclo.CicloID,
      numeroCiclo: ciclo.NumeroCiclo,
      fechaInicio: ciclo.FechaInicioCiclo
    };
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