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

  const paciente = obtenerPacienteCompletoPorId_(pacienteId);

  if (!paciente) {
    throw new Error('Paciente no encontrado.');
  }

  if (paciente.EstadoPaciente === ESTADOS_PACIENTE.ALTA) {
    throw new Error('El paciente ya está en estado ALTA.');
  }

  // 1. Eliminar sesiones futuras
  eliminarSesionesPendientesPaciente_(pacienteId);

  // 2. Finalizar asignación de ciclo
  finalizarAsignacionPaciente_(pacienteId);

  // 3. Limpiar ciclo en paciente
  const cicloId = paciente.CicloActivoID || paciente.CicloObjetivoID || '';

  // 4. Actualizar paciente
  actualizarPacienteAlta_(pacienteId, {
    fechaAlta,
    motivoCodigo,
    motivoTexto,
    comentario
  });

  // 5. Recalcular ciclo si aplica
  if (cicloId) {
    recalcularCiclo_(cicloId);
  }

  return {
    mensaje: 'Alta generada correctamente.'
  };
}

function eliminarSesionesPendientesPaciente_(pacienteId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];

    if (String(row[idx.PacienteID]) !== String(pacienteId)) continue;

    const estado = row[idx.EstadoSesion];

    if (
      estado === ESTADOS_SESION.PENDIENTE ||
      estado === ESTADOS_SESION.REPROGRAMADA
    ) {
      sheet.deleteRow(i + 1);
    }
  }
}

function finalizarAsignacionPaciente_(pacienteId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ASIGNACIONES_CICLO);
  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.PacienteID]) !== String(pacienteId)) continue;

    const estado = data[i][idx.EstadoAsignacion];

    if (estado === 'ACTIVO' || estado === 'RESERVADO') {
      sheet.getRange(i + 1, idx.EstadoAsignacion + 1).setValue('FINALIZADO');
      sheet.getRange(i + 1, idx.Observaciones + 1).setValue('Alta paciente');
    }
  }
}

function actualizarPacienteAlta_(pacienteId, datos) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.PacienteID]) !== String(pacienteId)) continue;

    sheet.getRange(i + 1, idx.EstadoPaciente + 1).setValue(ESTADOS_PACIENTE.ALTA);
    sheet.getRange(i + 1, idx.FechaCierre + 1).setValue(datos.fechaAlta);

    sheet.getRange(i + 1, idx.FechaAltaEfectiva + 1).setValue(datos.fechaAlta);
    sheet.getRange(i + 1, idx.MotivoAltaCodigo + 1).setValue(datos.motivoCodigo);
    sheet.getRange(i + 1, idx.MotivoAltaTexto + 1).setValue(datos.motivoTexto);
    sheet.getRange(i + 1, idx.ComentarioAlta + 1).setValue(datos.comentario);

    sheet.getRange(i + 1, idx.ProximaSesion + 1).setValue('');
    sheet.getRange(i + 1, idx.SesionesPendientes + 1).setValue(0);

    sheet.getRange(i + 1, idx.CicloObjetivoID + 1).setValue('');
    sheet.getRange(i + 1, idx.CicloActivoID + 1).setValue('');

    return;
  }

  throw new Error('No se pudo actualizar el paciente.');
}

function recalcularCiclo_(cicloId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CICLOS);
  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.CicloID]) !== String(cicloId)) continue;

    const capacidad = Number(data[i][idx.CapacidadMaxima] || 0);
    let ocupadas = 0;

    const asignaciones = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_ASIGNACIONES_CICLO)
      .getDataRange()
      .getValues();

    const aIdx = indexByHeader_(asignaciones[0]);

    for (let j = 1; j < asignaciones.length; j++) {
      if (String(asignaciones[j][aIdx.CicloID]) !== String(cicloId)) continue;

      if (asignaciones[j][aIdx.EstadoAsignacion] === 'ACTIVO') {
        ocupadas++;
      }
    }

    sheet.getRange(i + 1, idx.PlazasOcupadas + 1).setValue(ocupadas);
    sheet.getRange(i + 1, idx.PlazasLibres + 1).setValue(capacidad - ocupadas);

    return;
  }
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