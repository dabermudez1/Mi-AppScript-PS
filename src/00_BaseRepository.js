/**
 * BaseRepository
 * Abstracción CRUD para cualquier hoja de cálculo tratada como tabla.
 * Implementa la conversión automática entre Filas (Arrays) y Entidades (Objetos).
 */
class BaseRepository {
  constructor(sheetName, idColumnName) {
    this.sheetName = sheetName;
    this.idColumnName = idColumnName;
  }

  /** @protected */
  get sheet() {
    return dbContext.getSheet(this.sheetName);
  }

  /** @protected */
  get headerMap() {
    return dbContext.getHeaderMap(this.sheetName);
  }

  /**
   * Valida la integridad y tipos de datos de la entidad antes de guardar.
   * @private
   */
  _validateEntity(entity) {
    if (!entity[this.idColumnName]) {
      throw new Error(`Operación cancelada: El campo ID "${this.idColumnName}" es obligatorio.`);
    }

    const map = this.headerMap;
    for (const key in entity) {
      if (map[key] === undefined) continue;

      const value = entity[key];
      
      // Validación y conversión automática de fechas
      if (key.toLowerCase().includes('fecha') && value && typeof value === 'string') {
        const dateValue = new Date(value);
        if (isNaN(dateValue.getTime())) {
          throw new Error(`El campo "${key}" no tiene un formato de fecha válido: ${value}`);
        }
        entity[key] = dateValue;
      }

      // Aquí se podrían añadir validaciones adicionales para números o enums
    }
  }

  /**
   * Mapea una fila (array de valores) a un objeto JS basado en los encabezados de la hoja.
   */
  _rowToEntity(row) {
    const map = this.headerMap;
    const entity = {};
    for (const key in map) {
      entity[key] = row[map[key]];
    }
    return entity;
  }

  /**
   * Convierte un objeto JS en una fila (array) respetando el orden de las columnas de la hoja.
   */
  _entityToRow(entity) {
    const map = this.headerMap;
    const indices = Object.values(map);
    const maxIndex = indices.length > 0 ? Math.max(...indices) : 0;
    const row = new Array(maxIndex + 1).fill('');
    
    for (const key in entity) {
      if (map[key] !== undefined) {
        row[map[key]] = entity[key];
      }
    }
    return row;
  }

  findAll() {
    const data = this.sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    return data.slice(1).map(row => this._rowToEntity(row));
  }

  findById(id) {
    const data = this.sheet.getDataRange().getValues();
    const map = this.headerMap;
    const idIdx = map[this.idColumnName];
    if (idIdx === undefined) throw new Error(`Columna ID "${this.idColumnName}" no encontrada en ${this.sheetName}`);

    const row = data.find(r => String(r[idIdx]) === String(id));
    return row ? this._rowToEntity(row) : null;
  }

  /**
   * Elimina físicamente un registro de la hoja basándose en su ID.
   * @param {string|number} id 
   * @returns {boolean} True si se eliminó, false si no se encontró.
   */
  deleteById(id) {
    const data = this.sheet.getDataRange().getValues();
    const map = this.headerMap;
    const idIdx = map[this.idColumnName];
    
    const rowIndex = data.findIndex((r, i) => i > 0 && String(r[idIdx]) === String(id));

    if (rowIndex !== -1) {
      // rowIndex es 0-based para el array, sumamos 1 para que sea 1-based para la hoja
      this.sheet.deleteRow(rowIndex + 1);
      return true;
    }
    return false;
  }

  /**
   * Crea o actualiza una entidad basándose en su ID.
   */
  save(entity) {
    this._validateEntity(entity);

    const id = entity[this.idColumnName];
    const data = this.sheet.getDataRange().getValues();
    const map = this.headerMap;
    const idIdx = map[this.idColumnName];
    
    const rowIndex = data.findIndex((r, i) => i > 0 && String(r[idIdx]) === String(id));
    const rowData = this._entityToRow(entity);

    if (rowIndex === -1) {
      this.sheet.appendRow(rowData);
    } else {
      this.sheet.getRange(rowIndex + 1, 1, 1, rowData.length).setValues([rowData]);
    }
    return entity;
  }

  /**
   * Inserta múltiples entidades de una sola vez para optimizar rendimiento.
   * @param {Array<Object>} entities 
   */
  insertMany(entities) {
    if (!entities.length) return;
    entities.forEach(e => this._validateEntity(e));
    const rows = entities.map(e => this._entityToRow(e));
    const lastRow = this.sheet.getLastRow();
    this.sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}