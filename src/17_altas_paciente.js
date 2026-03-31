/***********************
 * BLOQUE 17
 * ALTAS DE PACIENTE
 ***********************/

function procesarAltaPacienteFormulario(formData) {
  if (!formData || typeof formData !== 'object') {
    throw new Error(
      'No se recibieron datos del formulario de alta.\n\n' +
      'Debes abrir el formulario de alta y enviarlo desde la interfaz.'
    );
  }

  const pacienteId = formData.pacienteId;
  const fechaAlta = parseFechaES_(formData.fechaAlta);
  const motivoCodigo = Number(formData.motivoCodigo);
  const motivoTexto = formData.motivoTexto || '';
  const comentario = String(formData.comentario || '').trim();

  if (!pacienteId) {
    throw new Error('Paciente no válido.');
  }

  if (!fechaAlta) {
    throw new Error('Fecha de alta no válida.');
  }

  if (!motivoCodigo) {
    throw new Error('Debes seleccionar un motivo de alta.');
  }

  if (motivoCodigo === 8 && !comentario) {
    throw new Error('Debes indicar un comentario para "Otros".');
  }

  const service = new PatientService();
  service.dischargePatient(pacienteId, {
    fechaAlta,
    motivoCodigo,
    motivoTexto,
    comentario
  });

  // Recalcular ocupación de ciclos de forma global tras el alta
  new MaintenanceService().recalculateCycleOccupancy();

  return {
    mensaje: 'Alta generada correctamente.'
  };
}

function altaPaciente() {
  const html = HtmlService
    .createHtmlOutputFromFile('AltaPacienteForm')
    .setWidth(760)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, 'Alta de paciente');
}

function altaPacienteDesdePaciente(pacienteId) {
  if (!pacienteId) {
    throw new Error('No se indicó paciente para generar alta.');
  }

  const template = HtmlService.createTemplateFromFile('AltaPacienteForm');
  template.pacientePreseleccionadoId = String(pacienteId);

  const html = template
    .evaluate()
    .setWidth(760)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, 'Alta de paciente');
}

function obtenerPacientesAltaFormulario() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const idx = indexByHeader_(data[0]);
  const out = [];

  for (let i = 1; i < data.length; i++) {
    const estado = data[i][idx.EstadoPaciente];

    if (
      estado !== ESTADOS_PACIENTE.ACTIVO &&
      estado !== ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO &&
      estado !== ESTADOS_PACIENTE.ESPERA
    ) {
      continue;
    }

    out.push({
      pacienteId: data[i][idx.PacienteID],
      label:
        (data[i][idx.Nombre] || 'SIN_NOMBRE') +
        ' | ' + (data[i][idx.ModalidadSolicitada] || '') +
        ' | ' + (estado || '')
    });
  }

  out.sort((a, b) => String(a.label).localeCompare(String(b.label)));
  return out;
}

function obtenerDetallePacienteAltaFormulario(pacienteId) {
  const paciente = obtenerPacienteCompletoPorId_(pacienteId);

  if (!paciente) {
    throw new Error('Paciente no encontrado.');
  }

  return {
    pacienteId: paciente.PacienteID,
    nombre: paciente.Nombre || '',
    modalidad: paciente.ModalidadSolicitada || '',
    estadoPaciente: paciente.EstadoPaciente || '',
    fechaPrimeraConsulta: formatearFecha_(paciente.FechaPrimeraConsulta),
    cicloObjetivoId: paciente.CicloObjetivoID || '',
    cicloActivoId: paciente.CicloActivoID || '',
    proximaSesion: formatearFecha_(paciente.ProximaSesion),
    sesionesPendientes: Number(paciente.SesionesPendientes || 0)
  };
}

function obtenerMotivosAltaFormulario() {
  return [
    { codigo: 1, texto: 'Alta terapéutica' },
    { codigo: 2, texto: 'Abandono' },
    { codigo: 3, texto: 'No acude a 1ª consulta' },
    { codigo: 4, texto: 'Alta por fuerza mayor' },
    { codigo: 5, texto: 'Derivación a Salud Mental' },
    { codigo: 6, texto: 'No asumido / excluído' },
    { codigo: 8, texto: 'Otros' }
  ];
}