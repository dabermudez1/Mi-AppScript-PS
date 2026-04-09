/***********************
 * BLOQUE 10
 * CONFIG_MODALIDADES HTML
 ***********************/

function gestionarConfigModalidades() {
  const html = HtmlService
    .createHtmlOutputFromFile('ConfigModalidadesForm')
    .setWidth(860)
    .setHeight(640);

  SpreadsheetApp.getUi().showModalDialog(html, 'Configuración de modalidades');
}

function obtenerConfigModalidadesFormulario() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG_MODALIDADES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_CONFIG_MODALIDADES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return {
      modalidades: [],
      catalogos: obtenerCatalogosConfigModalidades_()
    };
  }

  const idx = indexByHeader_(data[0]);

  const modalidades = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const modalidad = String(row[idx.Modalidad] || '').trim();
    if (!modalidad) continue;

    modalidades.push({
      modalidad: modalidad,
      tipoModalidad: row[idx.TipoModalidad] || '',
      activa: row[idx.Activa] === true,
      diaSemana: row[idx.DiaSemana] || '',
      frecuenciaDias: Number(row[idx.FrecuenciaDias] || 0),
      fechaBase: formatearFechaISOInput_(row[idx.FechaBase]),
      horaBase: row[idx.HoraBase] || '', // Nuevo campo
      capacidadMaxima: Number(row[idx.CapacidadMaxima] || 0),
      sesionesPorCiclo: Number(row[idx.SesionesPorCiclo] || 0),
      notas: row[idx.Notas] || ''
    });
  }

  modalidades.sort((a, b) => String(a.modalidad).localeCompare(String(b.modalidad)));

  return {
    modalidades: modalidades,
    catalogos: obtenerCatalogosConfigModalidades_()
  };
}

function obtenerCatalogosConfigModalidades_() {
  return {
    diasSemana: obtenerValoresCatalogo_('DIAS_SEMANA'),
    tiposModalidad: obtenerValoresCatalogo_('TIPOS_MODALIDAD'),
    modalidades: obtenerValoresCatalogo_('MODALIDADES')
  };
}

function guardarConfigModalidadFormulario(formData) {
  const modalidad = String(formData.modalidad || '').trim();
  const activa = formData.activa === true;
  const diaSemana = String(formData.diaSemana || '').trim();
  const frecuenciaDias = Number(formData.frecuenciaDias || 0);
  const fechaBaseISO = String(formData.fechaBase || '').trim();
  const horaBase = String(formData.horaBase || '').trim(); // Nuevo
  const capacidadMaxima = Number(formData.capacidadMaxima || 0);
  const sesionesPorCiclo = Number(formData.sesionesPorCiclo || 0);
  const notas = String(formData.notas || '').trim();

  if (!modalidad) {
    throw new Error('La modalidad es obligatoria.');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG_MODALIDADES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_CONFIG_MODALIDADES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('No hay datos en CONFIG_MODALIDADES.');
  }

  const idx = indexByHeader_(data[0]);

  const diasValidos = obtenerValoresCatalogo_('DIAS_SEMANA');
  const modalidadesValidas = obtenerValoresCatalogo_('MODALIDADES');

  if (!modalidadesValidas.includes(modalidad)) {
    throw new Error('La modalidad no es válida.');
  }

  if (!Number.isFinite(frecuenciaDias) || frecuenciaDias <= 0) {
    throw new Error('FrecuenciaDias debe ser un número mayor que 0.');
  }

  if (!Number.isFinite(capacidadMaxima) || capacidadMaxima < 0) {
    throw new Error('CapacidadMaxima debe ser un número mayor o igual que 0.');
  }

  if (!Number.isFinite(sesionesPorCiclo) || sesionesPorCiclo <= 0) {
    throw new Error('SesionesPorCiclo debe ser un número mayor que 0.');
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.Modalidad] || '').trim() !== modalidad) continue;

    const tipoModalidad = String(data[i][idx.TipoModalidad] || '').trim();

    if (tipoModalidad === TIPOS_MODALIDAD.INDIVIDUAL) {
      if (diaSemana) {
        throw new Error('INDIVIDUAL no debe tener DiaSemana.');
      }
      if (fechaBaseISO) {
        throw new Error('INDIVIDUAL no debe tener FechaBase.');
      }
      if (horaBase) {
        throw new Error('INDIVIDUAL no debe tener HoraBase.');
      }
    }

    if (tipoModalidad === TIPOS_MODALIDAD.GRUPO) {
      if (!diaSemana) {
        throw new Error('Las modalidades de grupo requieren DiaSemana.');
      }
      if (!diasValidos.includes(diaSemana)) {
        throw new Error('DiaSemana no es válido.');
      }
      if (!fechaBaseISO) {
        throw new Error('Las modalidades de grupo requieren FechaBase.');
      }

      if (!horaBase) {
        throw new Error('Las modalidades de grupo requieren HoraBase.');
      }
      if (!/^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/.test(horaBase)) {
        throw new Error('HoraBase no es válida. Se espera formato HH:mm.');
      }

      const fechaBase = parseFechaISO_(fechaBaseISO);
      if (!(fechaBase instanceof Date)) {
        throw new Error('FechaBase no es válida.');
      }

      const diaReal = convertirDiaSemanaATexto_(fechaBase);
      // La validación de día de la semana se mantiene para la fecha base
      if (diaReal !== diaSemana) {
        throw new Error(
          'La FechaBase no coincide con el DiaSemana configurado.\n\n' +
          'DiaSemana: ' + diaSemana + '\n' +
          'FechaBase: ' + formatearFecha_(fechaBase) + ' (' + diaReal + ')'
        );
      }

      // Guardamos la fecha y la hora base
      sheet.getRange(i + 1, idx.FechaBase + 1).setValue(fechaBase);
      sheet.getRange(i + 1, idx.HoraBase + 1).setValue(horaBase);
    } else {
      sheet.getRange(i + 1, idx.FechaBase + 1).setValue('');
      sheet.getRange(i + 1, idx.HoraBase + 1).setValue('');
    }

    sheet.getRange(i + 1, idx.Activa + 1).setValue(activa);
    sheet.getRange(i + 1, idx.DiaSemana + 1).setValue(tipoModalidad === TIPOS_MODALIDAD.GRUPO ? diaSemana : '');
    sheet.getRange(i + 1, idx.FrecuenciaDias + 1).setValue(frecuenciaDias);
    sheet.getRange(i + 1, idx.CapacidadMaxima + 1).setValue(capacidadMaxima);
    sheet.getRange(i + 1, idx.SesionesPorCiclo + 1).setValue(sesionesPorCiclo);
    sheet.getRange(i + 1, idx.Notas + 1).setValue(notas);

    return {
      mensaje:
        'Configuración guardada correctamente.\n\n' +
        'Modalidad: ' + modalidad + '\n' +
        'Activa: ' + (activa ? 'Sí' : 'No') + '\n' +
        'Frecuencia: ' + frecuenciaDias + '\n' +
        'Capacidad: ' + capacidadMaxima + '\n' +
        'Sesiones por ciclo: ' + sesionesPorCiclo
    };
  }

  throw new Error('No se encontró la modalidad a actualizar.');
}

/***********************
 * GESTIÓN DE AGENDA UI
 ***********************/

/**
 * Abre el diálogo unificado para gestionar la agenda.
 */
function gestionarAgenda() {
  const html = HtmlService
    .createHtmlOutputFromFile('AgendaForm')
    .setWidth(1000)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'Gestión de Agenda');
}

/**
 * Función unificada para carga inicial.
 * Optimizada para reducir latencia y round-trips al servidor.
 */
function obtenerDatosInicialesAgenda() {
  return {
    catalogos: obtenerCatalogosAgenda(),
    // Cargamos directamente los datos procesados para la UI
    plantilla: obtenerAgendaPlantillaParaUI(),
    excepciones: obtenerAgendaExcepcionesParaUI(),
    fechaHoy: new Date().toISOString().split('T')[0]
  };
}

/**
 * Obtiene todos los slots de la plantilla para la UI.
 */
function obtenerAgendaPlantillaParaUI() {
  const repo = new AgendaTemplateRepository();
  return repo.findAll().map(slot => ({
    diaSemana: slot.DiaSemana,
    horaInicio: formatearHora_(slot.HoraInicio),
    tipoSlot: slot.TipoSlot,
    row: slot._row
  }));
}

/**
 * Obtiene los catálogos necesarios para la agenda.
 */
function obtenerCatalogosAgenda() {
  return {
    diasSemana: obtenerValoresCatalogo_('DIAS_SEMANA'),
    // Tipos de slot para la lógica de planificación y la UI
    tiposSlot: ['2.1', '2.2', '2.2/GRUPO', 'PRIMERA', 'SEGUIMIENTO', 'SEGUIMIENTO/GRUPO', 'DESCANSO']
  };
}

/**
 * Guarda o actualiza un slot en la plantilla semanal.
 */
function guardarSlotPlantilla(formData) {
  const repo = new AgendaTemplateRepository();
  const slot = {
    DiaSemana: formData.diaSemana,
    HoraInicio: formData.horaInicio,
    TipoSlot: formData.tipoSlot,
    _row: formData.row ? Number(formData.row) : null
  };
  repo.save(slot);
  return { mensaje: 'Slot de plantilla guardado correctamente.' };
}

/**
 * Actualiza la plantilla completa desde la vista de rejilla.
 * @param {Array<Object>} listaSlots - Lista de objetos {diaSemana, horaInicio, tipoSlot}.
 */
function actualizarPlantillaCompleta(listaSlots) {
  const repo = new AgendaTemplateRepository();
  const sheet = repo.getSheet();
  
  // 1. Limpiar datos actuales (manteniendo cabecera)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  
  if (!listaSlots || listaSlots.length === 0) return { mensaje: 'Plantilla vaciada.' };

  // 2. Preparar filas para guardado masivo
  const headers = HEADERS[SHEET_AGENDA_PLANTILLA];
  const idx = indexByHeader_(headers);
  const dataValues = listaSlots.map(slot => {
    const row = new Array(headers.length).fill('');
    row[idx.DiaSemana] = slot.diaSemana;
    row[idx.HoraInicio] = slot.horaInicio;
    row[idx.TipoSlot] = slot.tipoSlot;
    return row;
  });

  // 3. Escribir en bloque
  sheet.getRange(2, 1, dataValues.length, headers.length).setValues(dataValues);
  aplicarFormatoAgenda_();
  return { mensaje: 'Plantilla actualizada correctamente.' };
}

/**
 * Elimina un slot de la plantilla.
 */
function eliminarSlotPlantilla(row) {
  const repo = new AgendaTemplateRepository();
  repo.delete({ _row: Number(row) });
  return { mensaje: 'Slot eliminado.' };
}

/**
 * Obtiene las excepciones para la UI.
 */
function obtenerAgendaExcepcionesParaUI() {
  const repo = new AgendaExceptionRepository();
  const data = repo.findAll();
  
  return data.map(ex => {
    const dt = ex.Fecha instanceof Date ? ex.Fecha : new Date(ex.Fecha);
    return {
      fecha: formatearFecha_(dt),
      timestamp: dt.getTime(), // Sort key pre-calculada
      horaInicio: formatearHora_(ex.HoraInicio),
      tipoSlot: ex.TipoSlot,
      row: ex._row
    };
  }).sort((a, b) => a.timestamp - b.timestamp);
}

/**
 * Guarda una excepción (semana concreta o día específico).
 */
function guardarExcepcionAgenda(formData) {
  const repo = new AgendaExceptionRepository();
  const fecha = parseFechaISO_(formData.fecha);
  
  if (!fecha) throw new Error('Fecha no válida.');

  const excepcion = {
    Fecha: fecha,
    HoraInicio: formData.horaInicio || '', // Si es vacío, afecta a todo el día
    TipoSlot: formData.tipoSlot,
    _row: formData.row ? Number(formData.row) : null
  };
  
  repo.save(excepcion);
  return { mensaje: 'Excepción guardada correctamente.' };
}

/**
 * Elimina una excepción.
 */
function eliminarExcepcionAgenda(row) {
  const repo = new AgendaExceptionRepository();
  repo.delete({ _row: Number(row) });
  return { mensaje: 'Excepción eliminada.' };
}