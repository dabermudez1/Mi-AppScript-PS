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

  // 1. Invalida caché de ejecución para asegurar que los datos frescos se propaguen en esta sesión
  if (typeof __EXECUTION_CACHE__ !== 'undefined') {
    __EXECUTION_CACHE__[SHEET_PACIENTES] = null;
    __EXECUTION_CACHE__[SHEET_SESIONES] = null;
    __EXECUTION_CACHE__[SHEET_DATOS_CLINICOS_PACIENTES] = null;
  }
  SpreadsheetApp.flush();

  // 2. Propagar el cambio a la hoja de DATOS_CLINICOS_PACIENTES (Sincronización Clínica inmediata)
  sincronizarFichasClinicasPacientes();

  // 3. Recalcular ocupación de ciclos de forma global tras el alta
  new MaintenanceService().recalculateCycleOccupancy();

  // 4. Invalidar caché del dashboard para reflejar el alta en las métricas visuales
  eliminarCacheDashboard_();

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
  const repo = new PatientRepository();
  const pacientes = repo.findAll();

  const out = pacientes
    .filter(p => 
      p.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO ||
      p.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO ||
      p.EstadoPaciente === ESTADOS_PACIENTE.ESPERA
    )
    .map(p => ({
      pacienteId: p.PacienteID,
      label: `${p.Nombre || 'SIN_NOMBRE'} | ${p.ModalidadSolicitada || ''} | ${p.EstadoPaciente || ''}`
    }));

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