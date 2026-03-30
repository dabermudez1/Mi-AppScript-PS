/***********************
 * BLOQUE 14
 * REPROGRAMACIÓN SESIONES
 ***********************/

function abrirReprogramarSesion() {
  const html = HtmlService
    .createHtmlOutputFromFile('ReprogramarSesionForm')
    .setWidth(500)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Reprogramar sesión');
}

/***************
 * DATOS FORMULARIO
 ***************/
function obtenerDatosReprogramacion() {
  const patientRepo = new PatientRepository();
  const cicloRepo = new CicloRepository();

  const pacientesIndividuales = patientRepo.findAll()
    .filter(p => p.ModalidadSolicitada === MODALIDADES.INDIVIDUAL)
    .map(p => ({ id: p.PacienteID, nombre: p.Nombre }));

  const ciclosGrupo = cicloRepo.findAll()
    .filter(c => c.Modalidad !== MODALIDADES.INDIVIDUAL)
    .map(c => ({
      id: c.CicloID,
      label: `${c.Modalidad} | Ciclo ${c.NumeroCiclo}`
    }));

  return {
    modalidades: Object.values(MODALIDADES),
    pacientes: pacientesIndividuales,
    ciclos: ciclosGrupo
  };
}

/***************
 * GUARDAR
 ***************/
function guardarReprogramacion(data) {
  const modalidad = data.modalidad;

  if (modalidad === MODALIDADES.INDIVIDUAL) {
    return reprogramarSesionIndividual_(data);
  }

  return reprogramarSesionGrupo_(data);
}

/***************
 * INDIVIDUAL
 ***************/
function reprogramarSesionIndividual_(data) {
  const sessionService = new SessionService();

  const pacienteId = data.pacienteId;
  const numeroSesion = Number(data.numeroSesion);
  const nuevaFecha = parseFechaES_(data.fecha);

  if (!nuevaFecha) {
    throw new Error('La nueva fecha no es válida. Usa formato dd/mm/yyyy.');
  }

  if (!esFechaOperativaValida_(nuevaFecha)) {
    throw new Error(construirMensajeFechaNoOperativa_(nuevaFecha));
  }

  sessionService.rescheduleSession(pacienteId, numeroSesion, nuevaFecha);

  recalcularMetricasBasicas_();

  return {
    mensaje: 'Sesión individual reprogramada.\nCambios: ' + cambios
  };
}

/***************
 * GRUPO
 ***************/
function reprogramarSesionGrupo_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SESIONES);

  const sesData = sheet.getDataRange().getValues();
  const idx = indexByHeader_(sesData[0]);

  const cicloId = data.cicloId;
  const numeroSesion = Number(data.numeroSesion);
  const nuevaFecha = parseFechaES_(data.fecha);

  if (!nuevaFecha) {
    throw new Error('La nueva fecha no es válida. Usa formato dd/mm/yyyy.');
  }

  if (!esFechaOperativaValida_(nuevaFecha)) {
  throw new Error(construirMensajeFechaNoOperativa_(nuevaFecha));
  }

  const diaSemanaCiclo = obtenerDiaSemanaCiclo_(cicloId);

  if (!fechaCoincideConDiaSemana_(nuevaFecha, diaSemanaCiclo)) {
    throw new Error(
      'La nueva fecha no respeta el día fijo del grupo.\n\n' +
      'Día esperado: ' + diaSemanaCiclo + '\n' +
      'Fecha propuesta: ' + formatearFecha_(nuevaFecha)
    );
  }

  let cambios = 0;

  for (let i = 1; i < sesData.length; i++) {
    if (
      sesData[i][idx.CicloID] === cicloId &&
      Number(sesData[i][idx.NumeroSesion]) === numeroSesion &&
      sesData[i][idx.EstadoSesion] === ESTADOS_SESION.PENDIENTE
    ) {
      const fechaActual = sesData[i][idx.FechaSesion];

      if (!sesData[i][idx.FechaOriginal]) {
        sheet.getRange(i + 1, idx.FechaOriginal + 1).setValue(fechaActual);
      }

      sheet.getRange(i + 1, idx.FechaSesion + 1).setValue(nuevaFecha);
      sheet.getRange(i + 1, idx.ModificadaManual + 1).setValue(true);

      cambios++;
    }
  }

  recalcularMetricasBasicas_();

  return {
    mensaje: 'Sesión grupal reprogramada.\nSesiones afectadas: ' + cambios
  };
}

function obtenerSesionesPendientesIndividualFormulario(pacienteId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_SESIONES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const idx = indexByHeader_(data[0]);
  const sesionesMap = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (
      String(row[idx.PacienteID] || '') === String(pacienteId) &&
      row[idx.EstadoSesion] === ESTADOS_SESION.PENDIENTE
    ) {
      const numeroSesion = Number(row[idx.NumeroSesion] || 0);
      const fechaSesion = row[idx.FechaSesion];

      if (!sesionesMap[numeroSesion]) {
        sesionesMap[numeroSesion] = {
          numeroSesion: numeroSesion,
          fechaActual: formatearFecha_(fechaSesion),
          label: 'Sesión ' + numeroSesion + ' | ' + formatearFecha_(fechaSesion)
        };
      }
    }
  }

  return Object.values(sesionesMap).sort((a, b) => a.numeroSesion - b.numeroSesion);
}

function obtenerSesionesPendientesGrupoFormulario(cicloId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_SESIONES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const idx = indexByHeader_(data[0]);
  const sesionesMap = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (
      String(row[idx.CicloID] || '') === String(cicloId) &&
      row[idx.EstadoSesion] === ESTADOS_SESION.PENDIENTE
    ) {
      const numeroSesion = Number(row[idx.NumeroSesion] || 0);
      const fechaSesion = row[idx.FechaSesion];

      if (!sesionesMap[numeroSesion]) {
        sesionesMap[numeroSesion] = {
          numeroSesion: numeroSesion,
          fechaActual: formatearFecha_(fechaSesion),
          label: 'Sesión ' + numeroSesion + ' | ' + formatearFecha_(fechaSesion)
        };
      }
    }
  }

  return Object.values(sesionesMap).sort((a, b) => a.numeroSesion - b.numeroSesion);
}

function obtenerDiaSemanaCiclo_(cicloId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CICLOS);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_CICLOS + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('No hay datos en la hoja ' + SHEET_CICLOS + '.');
  }

  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.CicloID]) === String(cicloId)) {
      return data[i][idx.DiaSemana] || '';
    }
  }

  throw new Error('No se encontró el ciclo indicado.');
}

function abrirReprogramarSesionDesdeSesion(sesionId) {
  if (!sesionId) {
    throw new Error('No se indicó la sesión a reprogramar.');
  }

  const template = HtmlService.createTemplateFromFile('ReprogramarSesionForm');
  template.sesionPreseleccionadaId = String(sesionId);
  template.origenPantalla = 'SESIONES';

  const html = template
    .evaluate()
    .setWidth(500)
    .setHeight(540);

  SpreadsheetApp.getUi().showModalDialog(html, 'Reprogramar sesión');
}

function abrirReprogramarSesionDesdeIncidencia(sesionId) {
  if (!sesionId) {
    throw new Error('No se indicó la sesión a reprogramar.');
  }

  const template = HtmlService.createTemplateFromFile('ReprogramarSesionForm');
  template.sesionPreseleccionadaId = String(sesionId);
  template.origenPantalla = 'INCIDENCIAS';

  const html = template
    .evaluate()
    .setWidth(500)
    .setHeight(540);

  SpreadsheetApp.getUi().showModalDialog(html, 'Reprogramar sesión');
}

function obtenerSesionParaReprogramacionFormulario(sesionId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_SESIONES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('No hay sesiones registradas.');
  }

  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (String(row[idx.SesionID] || '') !== String(sesionId)) continue;

    const estadoSesion = row[idx.EstadoSesion] || '';
    if (estadoSesion !== ESTADOS_SESION.PENDIENTE) {
      throw new Error('Solo se pueden abrir para reprogramación sesiones en estado PENDIENTE.');
    }

    return {
      sesionId: row[idx.SesionID] || '',
      modalidad: row[idx.Modalidad] || '',
      pacienteId: row[idx.PacienteID] || '',
      cicloId: row[idx.CicloID] || '',
      numeroSesion: Number(row[idx.NumeroSesion] || 0),
      fechaActual: formatearFecha_(row[idx.FechaSesion]),
      nombrePaciente: row[idx.NombrePaciente] || ''
    };
  }

  throw new Error('No se encontró la sesión indicada.');
}