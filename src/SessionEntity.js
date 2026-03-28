/**
 * SessionEntity
 * Representa una sesión de terapia individual o grupal.
 */
class SessionEntity {
  constructor(data = {}) {
    this.SesionID = data.SesionID || '';
    this.PacienteID = data.PacienteID || '';
    this.CicloID = data.CicloID || '';
    this.AsignacionId = data.AsignacionId || '';
    this.Modalidad = data.Modalidad || '';
    this.NombrePaciente = data.NombrePaciente || '';
    this.NumeroSesion = Number(data.NumeroSesion || 0);
    this.FechaSesion = data.FechaSesion || null;
    this.EstadoSesion = data.EstadoSesion || '';
    this.FechaOriginal = data.FechaOriginal || null;
    this.ModificadaManual = !!data.ModificadaManual;
    this.Notas = data.Notas || '';
    this.CalendarEventId = data.CalendarEventId || '';
    this.CalendarSyncStatus = data.CalendarSyncStatus || '';
    this.CalendarLastSync = data.CalendarLastSync || null;
    this.CalendarEventTitle = data.CalendarEventTitle || '';
    this.CalendarHash = data.CalendarHash || '';
  }

  validate() {
    if (!this.PacienteID) throw new Error('La sesión debe estar vinculada a un PacienteID.');
    if (!(this.FechaSesion instanceof Date)) throw new Error('La fecha de la sesión es inválida.');
    
    const estadosValidos = Object.values(ESTADOS_SESION);
    if (!estadosValidos.includes(this.EstadoSesion)) {
      throw new Error(`Estado de sesión no reconocido: ${this.EstadoSesion}`);
    }
  }
}