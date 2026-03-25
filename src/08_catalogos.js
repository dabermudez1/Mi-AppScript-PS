/***********************
 * BLOQUE 8
 * GESTIÓN HTML DE CATALOGOS
 ***********************/

function gestionarCatalogos() {
  const html = HtmlService
    .createHtmlOutputFromFile('CatalogosForm')
    .setWidth(760)
    .setHeight(620);

  SpreadsheetApp.getUi().showModalDialog(html, 'Gestión de catálogos');
}

function obtenerCatalogosAgrupadosFormulario() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CATALOGOS);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_CATALOGOS + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return [];
  }

  const idx = indexByHeader_(data[0]);

  if (idx.Catalogo === undefined || idx.Valor === undefined) {
    throw new Error('La hoja CATALOGOS debe tener columnas Catalogo y Valor.');
  }

  const agrupado = {};

  for (let i = 1; i < data.length; i++) {
    const catalogo = String(data[i][idx.Catalogo] || '').trim();
    const valor = String(data[i][idx.Valor] || '').trim();

    if (!catalogo || !valor) continue;

    if (!agrupado[catalogo]) {
      agrupado[catalogo] = [];
    }

    agrupado[catalogo].push(valor);
  }

  return Object.keys(agrupado)
    .sort((a, b) => a.localeCompare(b))
    .map(nombre => ({
      catalogo: nombre,
      valores: agrupado[nombre].slice().sort((a, b) => a.localeCompare(b))
    }));
}

function obtenerNombresCatalogosFormulario() {
  return obtenerCatalogosAgrupadosFormulario().map(x => x.catalogo);
}

function obtenerValoresDeCatalogoFormulario(nombreCatalogo) {
  if (!nombreCatalogo) return [];
  return obtenerValoresCatalogo_(String(nombreCatalogo).trim());
}

function agregarValorCatalogoFormulario(formData) {
  const catalogo = String(formData.catalogo || '').trim();
  const valor = String(formData.valor || '').trim();

  if (!catalogo) {
    throw new Error('El nombre del catálogo es obligatorio.');
  }

  if (!valor) {
    throw new Error('El valor es obligatorio.');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CATALOGOS);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_CATALOGOS + '.');
  }

  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    const c = String(data[i][idx.Catalogo] || '').trim();
    const v = String(data[i][idx.Valor] || '').trim();

    if (c === catalogo && v === valor) {
      throw new Error('Ese valor ya existe en el catálogo.');
    }
  }

  sheet.appendRow([catalogo, valor]);

  return {
    mensaje:
      'Valor añadido correctamente.\n\n' +
      'Catálogo: ' + catalogo + '\n' +
      'Valor: ' + valor
  };
}

function eliminarValorCatalogoFormulario(formData) {
  const catalogo = String(formData.catalogo || '').trim();
  const valor = String(formData.valor || '').trim();

  if (!catalogo) {
    throw new Error('Falta el catálogo.');
  }

  if (!valor) {
    throw new Error('Falta el valor.');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CATALOGOS);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_CATALOGOS + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('No hay datos en CATALOGOS.');
  }

  const idx = indexByHeader_(data[0]);

  for (let i = data.length - 1; i >= 1; i--) {
    const c = String(data[i][idx.Catalogo] || '').trim();
    const v = String(data[i][idx.Valor] || '').trim();

    if (c === catalogo && v === valor) {
      sheet.deleteRow(i + 1);

      return {
        mensaje:
          'Valor eliminado correctamente.\n\n' +
          'Catálogo: ' + catalogo + '\n' +
          'Valor: ' + valor
      };
    }
  }

  throw new Error('No se encontró el valor en el catálogo.');
}