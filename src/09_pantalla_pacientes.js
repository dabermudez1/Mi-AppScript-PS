/***********************
 * BLOQUE 9
 * PANTALLA VISUAL PACIENTES
 ***********************/

function abrirPantallaPacientes() {
  const html = HtmlService
    .createHtmlOutputFromFile('PantallaPacientes')
    .setWidth(1100)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, 'Pacientes');
}

function obtenerDatosPantallaPacientes() {
  const repo = new PatientRepository();
  const pacientes = repo.findAll().map(p => ({
    pacienteId: p.PacienteID,
    nombre: p.Nombre,
    modalidad: p.ModalidadSolicitada,
    estadoPaciente: p.EstadoPaciente, // Clave vital para los filtros
    fechaAlta: formatearFecha_(p.FechaAlta),
    fechaPrimeraConsulta: formatearFecha_(p.FechaPrimeraConsulta),
    fechaPrimeraSesionReal: formatearFecha_(p.FechaPrimeraSesionReal),
    proximaSesion: formatearFecha_(p.ProximaSesion),
    fechaCierre: formatearFecha_(p.FechaCierre),
    cicloObjetivoId: p.CicloObjetivoID || '',
    cicloActivoId: p.CicloActivoID || '',
    sesionesCompletadas: Number(p.SesionesCompletadas || 0),
    sesionesPendientes: Number(p.SesionesPendientes || 0),
    observaciones: p.Observaciones || ''
  }));

  return {
    pacientes: pacientes,
    estadosPaciente: obtenerValoresCatalogo_('ESTADOS_PACIENTE'),
    modalidades: obtenerValoresCatalogo_('MODALIDADES')
  };
}

function abrirAltaPacienteDesdePantalla(pacienteId) {
  altaPacienteDesdePaciente(pacienteId);
}

function abrirEditarPacienteDesdePantalla(pacienteId) {
  if (!pacienteId) {
    throw new Error('No se indicó paciente para editar.');
  }

  editarPacienteDesdeId_(pacienteId);
}

function abrirSesionesPacienteDesdePantalla(pacienteId) {
  if (!pacienteId) {
    throw new Error('No se indicó paciente para ver sesiones.');
  }

  abrirPantallaSesionesDesdePaciente(pacienteId);
}

function abrirPantallaPacientesDesdePantalla() {
  abrirPantallaPacientes();
}

function abrirFichaClinicaPacienteDesdePantalla(pacienteId) {
  if (!pacienteId) {
    throw new Error('No se indicó paciente para la ficha clínica.');
  }

  fichaClinicaPacienteDesdePaciente(pacienteId);
}