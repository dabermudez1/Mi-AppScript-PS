/***********************
 * BLOQUE 19
 * DATOS CLÍNICOS PACIENTES
 ***********************/

function fichaClinicaPaciente() {
  const template = HtmlService.createTemplateFromFile('DatosClinicosPacienteForm');
  template.pacientePreseleccionadoId = '';

  const html = template
    .evaluate()
    .setWidth(820)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, 'Ficha clínica paciente');
}

function fichaClinicaPacienteDesdePaciente(pacienteId) {
  if (!pacienteId) {
    throw new Error('No se indicó paciente para la ficha clínica.');
  }

  const template = HtmlService.createTemplateFromFile('DatosClinicosPacienteForm');
  template.pacientePreseleccionadoId = String(pacienteId);

  const html = template
    .evaluate()
    .setWidth(820)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, 'Ficha clínica paciente');
}

function obtenerPacientesFichaClinicaFormulario() {
  const patientRepo = new PatientRepository();
  const pacientes = patientRepo.findAll(); // Usa el repositorio

  return pacientes.map(p => ({
    pacienteId: p.PacienteID,
    label:
      (p.Nombre || 'SIN_NOMBRE') +
      ' | ' + (p.NHC || 'SIN_NHC') +
      ' | ' + (p.ModalidadSolicitada || '') +
      ' | ' + (p.EstadoPaciente || '')
  })).sort((a, b) => String(a.label).localeCompare(String(b.label)));
}

function obtenerCatalogosFichaClinicaFormulario() {
  return {
    sexoGenero: ['Varón', 'Mujer', 'Otro'],
    nivelEstudios: ['Bajo', 'Medio', 'Superior', 'No consta'],
    motivoConsultaDiagnostico: [
      'Trastorno adaptativo',
      'Duelo',
      'Episodio depresivo',
      'Trastorno de ansiedad',
      'Trastorno de síntomas somáticos',
      'Otros'
    ],
    comorbilidad: ['Sí', 'No'],
    antecedentesSM: ['Sí', 'No', 'Desconocido'],
    psicofarmacos: [
      'No toma',
      'Ansiolíticos',
      'Antidepresivos',
      'Ansiolíticos + antidepresivos',
      'Otros',
      'No consta'
    ],
    situacionLaboralPrevia: [
      'Trabaja',
      'IT',
      'Incapacidad',
      'Desempleo',
      'Trabajo no remunerado',
      'Estudiante',
      'Jubilado',
      'No consta'
    ],
    cambioSituacionLaboralAlta: [
      'Sin cambios',
      'Inicia IT',
      'Finaliza IT y se reincorpora',
      'No valorable'
    ],
    cambioFarmacologicoAlta: [
      'Sin cambios',
      'Deprescripción',
      'Descenso',
      'Inicio o aumento',
      'No valorable'
    ],
    escalaSatisfaccion: ['Nada', 'Un poco', 'Bastante', 'Mucho']
  };
}

function obtenerFichaClinicaPacienteFormulario(pacienteId) {
  if (!pacienteId) {
    throw new Error('No se indicó paciente.');
  }

  asegurarFilaFichaClinicaPaciente_(pacienteId);
  sincronizarCamposAutomaticosFichaClinica_(pacienteId);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATOS_CLINICOS_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_DATOS_CLINICOS_PACIENTES + '.');
  }

  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.PacienteID] || '') !== String(pacienteId)) continue;

    return {
      pacienteId: data[i][idx.PacienteID] || '',
      nombre: data[i][idx.Nombre] || '',
      nhc: data[i][idx.NHC] || '',
      fechaAltaPrograma: formatearFecha_(data[i][idx.FechaAltaPrograma]),
      fechaPrimeraConsulta: formatearFecha_(data[i][idx.FechaPrimeraConsulta]),
      fechaAltaEfectiva: formatearFecha_(data[i][idx.FechaAltaEfectiva]),
      estadoPacienteActual: data[i][idx.EstadoPacienteActual] || '',
      tipoIntervencionPrincipal: data[i][idx.TipoIntervencionPrincipal] || '',
      finTratamientoCodigo: data[i][idx.FinTratamientoCodigo] || '',
      finTratamientoTexto: data[i][idx.FinTratamientoTexto] || '',
      numeroSesionesTotal: Number(data[i][idx.NumeroSesionesTotal] || 0),
      tiempoEsperaHastaPrimeraConsultaDias: data[i][idx.TiempoEsperaHastaPrimeraConsultaDias] || '',
      sexoGenero: data[i][idx.SexoGenero] || '',
      edad: data[i][idx.Edad] || '',
      nivelEstudios: data[i][idx.NivelEstudios] || '',
      motivoConsultaDiagnostico: data[i][idx.MotivoConsultaDiagnostico] || '',
      motivoConsultaOtros: data[i][idx.MotivoConsultaOtros] || '',
      comorbilidad: data[i][idx.Comorbilidad] || '',
      antecedentesSM: data[i][idx.AntecedentesSM] || '',
      psicofarmacos: data[i][idx.Psicofarmacos] || '',
      situacionLaboralPrevia: data[i][idx.SituacionLaboralPrevia] || '',
      cambioSituacionLaboralAlta: data[i][idx.CambioSituacionLaboralAlta] || '',
      cambioFarmacologicoAlta: data[i][idx.CambioFarmacologicoAlta] || '',
      gad7Pre: data[i][idx.GAD7_PRE] || '',
      phq9Pre: data[i][idx.PHQ9_PRE] || '',
      whoqolPre: data[i][idx.WHOQOLBREF_PRE] || '',
      gad7Post: data[i][idx.GAD7_POST] || '',
      phq9Post: data[i][idx.PHQ9_POST] || '',
      whoqolPost: data[i][idx.WHOQOLBREF_POST] || '',
      escalaSatisfaccion: data[i][idx.EscalaSatisfaccion] || '',
      otrosComentarios: data[i][idx.OtrosComentarios] || ''
    };
  }

  throw new Error('No se encontró la ficha clínica del paciente.');
}

function guardarFichaClinicaPacienteFormulario(formData) {
  if (!formData || !formData.pacienteId) {
    throw new Error('No se indicó paciente.');
  }

  asegurarFilaFichaClinicaPaciente_(formData.pacienteId);
  sincronizarCamposAutomaticosFichaClinica_(formData.pacienteId);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATOS_CLINICOS_PACIENTES);
  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.PacienteID] || '') !== String(formData.pacienteId)) continue;

    sheet.getRange(i + 1, idx.NHC + 1).setValue(String(formData.nhc || '').trim());

    actualizarCamposClinicosBasicosEnPacientes_({
      pacienteId: formData.pacienteId,
      nhc: String(formData.nhc || '').trim(),
      sexoGenero: formData.sexoGenero || '',
      motivoConsultaDiagnostico: formData.motivoConsultaDiagnostico || '',
      motivoConsultaOtros: formData.motivoConsultaOtros || ''
    });

    sheet.getRange(i + 1, idx.SexoGenero + 1).setValue(formData.sexoGenero || '');
    sheet.getRange(i + 1, idx.Edad + 1).setValue(formData.edad || '');
    sheet.getRange(i + 1, idx.NivelEstudios + 1).setValue(formData.nivelEstudios || '');
    sheet.getRange(i + 1, idx.MotivoConsultaDiagnostico + 1).setValue(formData.motivoConsultaDiagnostico || '');
    sheet.getRange(i + 1, idx.MotivoConsultaOtros + 1).setValue(formData.motivoConsultaOtros || '');
    sheet.getRange(i + 1, idx.Comorbilidad + 1).setValue(formData.comorbilidad || '');
    sheet.getRange(i + 1, idx.AntecedentesSM + 1).setValue(formData.antecedentesSM || '');
    sheet.getRange(i + 1, idx.Psicofarmacos + 1).setValue(formData.psicofarmacos || '');
    sheet.getRange(i + 1, idx.SituacionLaboralPrevia + 1).setValue(formData.situacionLaboralPrevia || '');
    sheet.getRange(i + 1, idx.CambioSituacionLaboralAlta + 1).setValue(formData.cambioSituacionLaboralAlta || '');
    sheet.getRange(i + 1, idx.CambioFarmacologicoAlta + 1).setValue(formData.cambioFarmacologicoAlta || '');

    sheet.getRange(i + 1, idx.GAD7_PRE + 1).setValue(formData.gad7Pre || '');
    sheet.getRange(i + 1, idx.PHQ9_PRE + 1).setValue(formData.phq9Pre || '');
    sheet.getRange(i + 1, idx.WHOQOLBREF_PRE + 1).setValue(formData.whoqolPre || '');
    sheet.getRange(i + 1, idx.GAD7_POST + 1).setValue(formData.gad7Post || '');
    sheet.getRange(i + 1, idx.PHQ9_POST + 1).setValue(formData.phq9Post || '');
    sheet.getRange(i + 1, idx.WHOQOLBREF_POST + 1).setValue(formData.whoqolPost || '');
    sheet.getRange(i + 1, idx.EscalaSatisfaccion + 1).setValue(formData.escalaSatisfaccion || '');
    sheet.getRange(i + 1, idx.OtrosComentarios + 1).setValue(formData.otrosComentarios || '');

    sincronizarCamposAutomaticosFichaClinica_(formData.pacienteId);

    return {
      mensaje: 'Ficha clínica guardada correctamente.'
    };
  }

  throw new Error('No se pudo guardar la ficha clínica.');
}

function asegurarFilaFichaClinicaPaciente_(pacienteId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DATOS_CLINICOS_PACIENTES);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DATOS_CLINICOS_PACIENTES);
    const headers = HEADERS[SHEET_DATOS_CLINICOS_PACIENTES];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.PacienteID] || '') === String(pacienteId)) {
      return;
    }
  }

  sheet.appendRow(new Array(HEADERS[SHEET_DATOS_CLINICOS_PACIENTES].length).fill('').map((_, i) => i === 0 ? pacienteId : ''));
}

function sincronizarCamposAutomaticosFichaClinica_(pacienteId) {
  const paciente = obtenerPacienteCompletoPorId_(pacienteId);
  if (!paciente) {
    throw new Error('Paciente no encontrado.');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATOS_CLINICOS_PACIENTES);
  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);

  const sesiones = obtenerResumenSesionesClinico_(pacienteId);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.PacienteID] || '') !== String(pacienteId)) continue;

    const nhcActual = paciente.NHC || data[i][idx.NHC] || '';
    const sexoGeneroActual = paciente.SexoGenero || '';
    const motivoConsultaDiagnosticoActual = paciente.MotivoConsultaDiagnostico || '';
    const motivoConsultaOtrosActual = paciente.MotivoConsultaOtros || '';
    const fechaAltaEfectiva = paciente.FechaAltaEfectiva || paciente.FechaCierre || '';
    const tipoIntervencion = paciente.ModalidadSolicitada === MODALIDADES.INDIVIDUAL ? 'Individual' : 'Grupal';

    let finCodigo = '';
    let finTexto = '';

    if (paciente.EstadoPaciente === ESTADOS_PACIENTE.ALTA) {
      finCodigo = paciente.MotivoAltaCodigo || '';
      finTexto = paciente.MotivoAltaTexto || '';
    } else {
      finCodigo = 7;
      finTexto = 'Activo en el programa';
    }

    let tiempoEspera = '';
    if (paciente.FechaPrimeraConsulta instanceof Date && paciente.FechaPrimeraSesionReal instanceof Date) {
      tiempoEspera = diferenciaDiasFechas_(paciente.FechaPrimeraConsulta, paciente.FechaPrimeraSesionReal);
    }

    sheet.getRange(i + 1, idx.Nombre + 1).setValue(paciente.Nombre || '');
    sheet.getRange(i + 1, idx.NHC + 1).setValue(nhcActual);
    sheet.getRange(i + 1, idx.SexoGenero + 1).setValue(sexoGeneroActual);
    sheet.getRange(i + 1, idx.MotivoConsultaDiagnostico + 1).setValue(motivoConsultaDiagnosticoActual);
    sheet.getRange(i + 1, idx.MotivoConsultaOtros + 1).setValue(motivoConsultaOtrosActual);
    sheet.getRange(i + 1, idx.FechaAltaPrograma + 1).setValue(paciente.FechaAlta || '');
    sheet.getRange(i + 1, idx.FechaPrimeraConsulta + 1).setValue(paciente.FechaPrimeraConsulta || '');
    sheet.getRange(i + 1, idx.FechaAltaEfectiva + 1).setValue(fechaAltaEfectiva);
    sheet.getRange(i + 1, idx.EstadoPacienteActual + 1).setValue(paciente.EstadoPaciente || '');
    sheet.getRange(i + 1, idx.TipoIntervencionPrincipal + 1).setValue(tipoIntervencion);
    sheet.getRange(i + 1, idx.FinTratamientoCodigo + 1).setValue(finCodigo);
    sheet.getRange(i + 1, idx.FinTratamientoTexto + 1).setValue(finTexto);
    sheet.getRange(i + 1, idx.NumeroSesionesTotal + 1).setValue(sesiones.numeroSesionesTotal);
    sheet.getRange(i + 1, idx.TiempoEsperaHastaPrimeraConsultaDias + 1).setValue(tiempoEspera);

    return;
  }

  throw new Error('No se encontró fila clínica para sincronizar.');
}

function obtenerResumenSesionesClinico_(pacienteId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  if (!sheet) {
    return { completadas: 0, numeroSesionesTotal: 0 };
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return { completadas: 0, numeroSesionesTotal: 0 };
  }

  const idx = indexByHeader_(data[0]);
  let completadas = 0;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.PacienteID] || '') !== String(pacienteId)) continue;

    const estado = data[i][idx.EstadoSesion] || '';
    if (
      estado === ESTADOS_SESION.COMPLETADA_AUTO ||
      estado === ESTADOS_SESION.COMPLETADA_MANUAL
    ) {
      completadas++;
    }
  }

  const paciente = obtenerPacienteCompletoPorId_(pacienteId);
  const numeroSesionesTotal = (paciente && paciente.FechaPrimeraConsulta instanceof Date ? 1 : 0) + completadas;

  return {
    completadas: completadas,
    numeroSesionesTotal: numeroSesionesTotal
  };
}

function diferenciaDiasFechas_(fechaA, fechaB) {
  const a = normalizarFecha_(fechaA);
  const b = normalizarFecha_(fechaB);
  const msDia = 24 * 60 * 60 * 1000;
  return Math.round((b.getTime() - a.getTime()) / msDia);
}

function actualizarCamposClinicosBasicosEnPacientes_(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');
  }

  const data = sheet.getDataRange().getValues();
  const idx = indexByHeader_(data[0]);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.PacienteID] || '') !== String(params.pacienteId)) continue;

    if (idx.NHC !== undefined) {
      sheet.getRange(i + 1, idx.NHC + 1).setValue(params.nhc || '');
    }

    if (idx.SexoGenero !== undefined) {
      sheet.getRange(i + 1, idx.SexoGenero + 1).setValue(params.sexoGenero || '');
    }

    if (idx.MotivoConsultaDiagnostico !== undefined) {
      sheet.getRange(i + 1, idx.MotivoConsultaDiagnostico + 1).setValue(params.motivoConsultaDiagnostico || '');
    }

    if (idx.MotivoConsultaOtros !== undefined) {
      sheet.getRange(i + 1, idx.MotivoConsultaOtros + 1).setValue(params.motivoConsultaOtros || '');
    }

    return;
  }
}

function sincronizarFichasClinicasPacientes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_PACIENTES + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return { mensaje: 'No hay pacientes para sincronizar.' };
  }

  const idx = indexByHeader_(data[0]);
  let procesados = 0;

  for (let i = 1; i < data.length; i++) {
    const pacienteId = data[i][idx.PacienteID] || '';
    if (!pacienteId) continue;

    asegurarFilaFichaClinicaPaciente_(pacienteId);
    sincronizarCamposAutomaticosFichaClinica_(pacienteId);
    procesados++;
  }

  return {
    mensaje:
      'Sincronización de fichas clínicas completada.\n\n' +
      'Pacientes procesados: ' + procesados
  };
}

function ejecutarSincronizarFichasClinicasPacientes() {
  const res = sincronizarFichasClinicasPacientes();
  SpreadsheetApp.getUi().alert(res.mensaje);
}
