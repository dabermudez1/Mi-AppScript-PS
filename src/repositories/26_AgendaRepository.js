/**
 * Repository para la configuración de la agenda (Plantilla y Excepciones).
 */
class AgendaRepository {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.SHEET_PLANTILLA = "AGENDA_PLANTILLA";
    this.SHEET_EXCEPCIONES = "AGENDA_EXCEPCIONES";
  }

  /**
   * Obtiene la plantilla semanal base.
   * @returns {Object} { "LUNES": [{hora: "09:30", tipo: "SEGUIMIENTO"}, ...] }
   */
  getWeeklyTemplate() {
    const sheet = this.ss.getSheetByName(this.SHEET_PLANTILLA);
    if (!sheet) throw new Error(`No se encontró la hoja ${this.SHEET_PLANTILLA}`);
    
    const data = sheet.getDataRange().getValues();
    const [, ...rows] = data; // Ignorar cabecera

    return rows.reduce((acc, row) => {
      const [dia, hora, tipo] = row;
      if (!dia) return acc;
      const diaUpper = dia.toString().toUpperCase().trim();
      if (!acc[diaUpper]) acc[diaUpper] = [];
      
      acc[diaUpper].push({ 
        hora: this._formatTime(hora), 
        tipo: tipo 
      });
      return acc;
    }, {});
  }

  /**
   * Obtiene las excepciones configuradas.
   */
  getExceptions() {
    const sheet = this.ss.getSheetByName(this.SHEET_EXCEPCIONES);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const [, ...rows] = data;

    return rows.map(row => ({
      fecha: this._formatDate(row[0]),
      hora: this._formatTime(row[1]),
      tipo: row[2]
    }));
  }

  _formatTime(timeValue) {
    if (timeValue instanceof Date) {
      return Utilities.formatDate(timeValue, Session.getScriptTimeZone(), "HH:mm");
    }
    // Si viene como string "HH:mm:ss" o similar de la celda
    return timeValue.toString().substring(0, 5);
  }

  _formatDate(dateValue) {
    return Utilities.formatDate(new Date(dateValue), Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
}
