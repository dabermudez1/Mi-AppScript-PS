/**
 * PatientEntity
 * Representa un Paciente en el sistema con validaciones de negocio.
 */
class PatientEntity {
  constructor(data = {}) {
    this.PacienteID = data.PacienteID || '';
    this.Nombre = data.Nombre || '';
    this.NHC = data.NHC || '';
    this.SexoGenero = data.SexoGenero || '';
    this.MotivoConsultaDiagnostico = data.MotivoConsultaDiagnostico || '';
    this.MotivoConsultaOtros = data.MotivoConsultaOtros || '';
    this.ModalidadSolicitada = data.ModalidadSolicitada || '';
    this.FechaAlta = data.FechaAlta || null;
    this.FechaPrimeraConsulta = data.FechaPrimeraConsulta || null;
    this.EstadoPaciente = data.EstadoPaciente || '';
    this.MotivoEspera = data.MotivoEspera || '';
    this.CicloObjetivoID = data.CicloObjetivoID || '';
    this.CicloActivoID = data.CicloActivoID || '';
    this.FechaPrimeraSesionReal = data.FechaPrimeraSesionReal || null;
    this.SesionesPlanificadas = Number(data.SesionesPlanificadas || 0);
    this.SesionesCompletadas = Number(data.SesionesCompletadas || 0);
    this.SesionesPendientes = Number(data.SesionesPendientes || 0);
    this.ProximaSesion = data.ProximaSesion || null;
    this.FechaCierre = data.FechaCierre || null;
    this.Observaciones = data.Observaciones || '';
    this.RecalcularSecuencia = !!data.RecalcularSecuencia;
  }

  /**
   * Valida que el paciente tenga los datos mínimos coherentes.
   * @throws {Error} Si los datos no son válidos.
   */
  validate() {
    if (!this.Nombre.trim()) throw new Error('El nombre del paciente es obligatorio.');
    if (!this.NHC.trim()) throw new Error('El NHC (Número de Historia) es obligatorio.');
    
    const modalidadesValidas = Object.values(MODALIDADES);
    if (!modalidadesValidas.includes(this.ModalidadSolicitada)) {
      throw new Error(`Modalidad no válida: ${this.ModalidadSolicitada}`);
    }

    if (!(this.FechaPrimeraConsulta instanceof Date)) {
      throw new Error('La fecha de primera consulta debe ser un objeto Date válido.');
    }
  }

  isIndividual() {
    return this.ModalidadSolicitada === MODALIDADES.INDIVIDUAL;
  }
}