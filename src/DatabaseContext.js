/**
 * DatabaseContext
 * Centraliza el acceso a Google Sheets y cachea referencias para optimizar rendimiento.
 * Al usar clasp, este archivo se gestionará localmente en la carpeta infrastructure.
 */
class DatabaseContext {
  constructor() {
    this._ss = null;
    this._sheets = new Map();
    this._headerMaps = new Map();
  }

  get ss() {
    if (!this._ss) {
      this._ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    return this._ss;
  }

  getSheet(name) {
    if (!this._sheets.has(name)) {
      const sheet = this.ss.getSheetByName(name);
      if (!sheet) throw new Error(`La hoja "${name}" no existe en el sistema.`);
      this._sheets.set(name, sheet);
    }
    return this._sheets.get(name);
  }

  /**
   * Obtiene un mapa de { NombreColumna: Indice0 } para una hoja.
   * Esto sustituye la necesidad de llamar a indexByHeader_ en cada función.
   */
  getHeaderMap(name) {
    if (!this._headerMaps.has(name)) {
      const sheet = this.getSheet(name);
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) return {};
      
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const map = {};
      headers.forEach((h, i) => {
        if (h) map[h] = i;
      });
      this._headerMaps.set(name, map);
    }
    return this._headerMaps.get(name);
  }

  clearCache() {
    this._sheets.clear();
    this._headerMaps.clear();
  }
}

const dbContext = new DatabaseContext();