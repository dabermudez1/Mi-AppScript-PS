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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return {
      pacientes: [],
      estadosPaciente: obtenerValoresCatalogo_('ESTADOS_PACIENTE'),
      modalidades: obtenerValoresCatalogo_('MODALIDADES')
    };
  }

  const idx = indexByHeader_(data[0]);

  const pacientes = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    pacientes.push({
      pacienteId: row[idx.PacienteID] || '',
      nombre: row[idx.Nombre] || '',
      modalidad: row[idx.ModalidadSolicitada] || '',
      fechaAlta: formatearFecha_(row[idx.FechaAlta]),
      fechaPrimeraConsulta: formatearFecha_(row[idx.FechaPrimeraConsulta]),
      estadoPaciente: row[idx.EstadoPaciente] || '',
      motivoEspera: row[idx.MotivoEspera] || '',
      cicloObjetivoId: row[idx.CicloObjetivoID] || '',
      cicloActivoId: row[idx.CicloActivoID] || '',
      fechaPrimeraSesionReal: formatearFecha_(row[idx.FechaPrimeraSesionReal]),
      sesionesPlanificadas: Number(row[idx.SesionesPlanificadas] || 0),
      sesionesCompletadas: Number(row[idx.SesionesCompletadas] || 0),
      sesionesPendientes: Number(row[idx.SesionesPendientes] || 0),
      proximaSesion: formatearFecha_(row[idx.ProximaSesion]),
      fechaCierre: formatearFecha_(row[idx.FechaCierre]),
      observaciones: row[idx.Observaciones] || ''
    });
  }

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