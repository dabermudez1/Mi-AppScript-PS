/**
 * 00_BaseRepository.js
 * Clase Base para manejar la persistencia en Google Sheets.
 * Implementa el mapeo automático entre Filas y Objetos.
 */

// Objeto global para cachear lecturas durante una misma ejecución del script
if (typeof __EXECUTION_CACHE__ === 'undefined') {
  var __EXECUTION_CACHE__ = {};
}

class BaseRepository {
  constructor(sheetName, headers) {
    this.sheetName = sheetName;
    this.headers = headers;
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  getSheet() {
    const sheet = this.ss.getSheetByName(this.sheetName);
    if (!sheet) throw new Error(`Hoja ${this.sheetName} no encontrada.`);
    return sheet;
  }

  /**
   * Obtiene todos los registros como una lista de objetos.
   */
  findAll() {
    // Si ya leímos esta hoja en esta ejecución, devolver caché
    if (__EXECUTION_CACHE__[this.sheetName]) {
      return __EXECUTION_CACHE__[this.sheetName];
    }

    const sheet = this.getSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const headerRow = data[0];
    const idx = this._indexHeaders(headerRow);

    const results = data.slice(1).map((row, i) => {
      const obj = { _row: i + 2 }; // Guardamos la referencia a la fila física
      this.headers.forEach(h => {
        obj[h] = row[idx[h]];
      });
      return obj;
    });

    __EXECUTION_CACHE__[this.sheetName] = results;
    return results;
  }

  /**
   * Busca un registro por una propiedad y valor específicos.
   */
  findOneBy(property, value) {
    const all = this.findAll();
    return all.find(item => String(item[property]) === String(value)) || null;
  }

  /**
   * Crea o actualiza un registro a partir de un objeto.
   */
  save(obj, idPropertyName) {
    __EXECUTION_CACHE__[this.sheetName] = null; // Invalida caché SIEMPRE al guardar
    
    const sheet = this.getSheet();
    const headerRow = this.headers; // Usar headers definidos en el constructor para consistencia
    const idx = this._indexHeaders(headerRow);

    const rowValues = headerRow.map(h => {
      let val = obj[h];
      if (val === undefined || val === null) return "";
      // Normalización de fechas para Sheets
      if (val instanceof Date) return val; 
      return val;
    });

    if (obj._row) {
      sheet.getRange(obj._row, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
      obj._row = sheet.getLastRow();
    }
    return obj;
  }

  /**
   * Inserta múltiples objetos nuevos en la hoja.
   * Cada objeto se añade como una nueva fila.
   * @param {Array<Object>} objects - Array de objetos a insertar.
   */
  insertAll(objects) {
    if (!objects || objects.length === 0) return;

    const sheet = this.getSheet();
    const headerRow = this.headers;
    
    const rowsToAppend = objects.map(obj => {
      return headerRow.map(h => {
        let val = obj[h];
        if (val === undefined || val === null) return "";
        if (val instanceof Date) return val;
        return val;
      });
    });

    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rowsToAppend.length, headerRow.length).setValues(rowsToAppend);
    __EXECUTION_CACHE__[this.sheetName] = null; // Invalida caché
  }

  /**
  * Guarda múltiples objetos en la hoja de una sola vez.
   * Extremadamente eficiente para procesos masivos.
   */
  saveAll(objects) {
    if (!objects || objects.length === 0) return;
    
    const sheet = this.getSheet();
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];
    const idx = this._indexHeaders(headerRow);

    objects.forEach(obj => {
      if (!obj._row) return; // Solo procesa actualizaciones de filas existentes
      const rowIndex = obj._row - 1;
      const rowValues = headerRow.map(h => (obj[h] !== undefined ? obj[h] : data[rowIndex][idx[h]]));
      data[rowIndex] = rowValues;
    });

    // Escribimos toda la tabla de una sola vez
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    
    // Invalidar caché
    __EXECUTION_CACHE__[this.sheetName] = null;
  }

  /**
   * Elimina un registro por su número de fila.
   */
  delete(obj) {
    if (!obj._row) throw new Error("No se puede eliminar un objeto sin referencia a fila.");
    this.getSheet().deleteRow(obj._row);
  }

  _indexHeaders(headerRow) {
    const map = {};
    headerRow.forEach((h, i) => {
      map[h] = i;
    });
    return map;
  }
}