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

/**
 * Obtiene todos los datos para el formulario de estadísticas, 
 * forzando una sincronización previa de los estados de alta.
 */
function obtenerDatosEstadisticasFichasFormulario() {
  // 1. Invalida caché de ejecución para leer datos reales
  if (typeof __EXECUTION_CACHE__ !== 'undefined') {
    Object.keys(__EXECUTION_CACHE__).forEach(key => __EXECUTION_CACHE__[key] = null);
  }
  SpreadsheetApp.flush();

  // 2. Sincroniza fichas para capturar las "Altas" recientes de la hoja PACIENTES
  sincronizarFichasClinicasPacientes();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATOS_CLINICOS_PACIENTES);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { filas: [], resumen: {} };

  const idx = indexByHeader_(data[0]);
  const filas = data.slice(1).map(row => {
    const pre = [row[idx.GAD7_PRE], row[idx.PHQ9_PRE], row[idx.WHOQOLBREF_PRE]];
    const post = [row[idx.GAD7_POST], row[idx.PHQ9_POST], row[idx.WHOQOLBREF_POST]];
    
    // Cálculos de mejora (Delta)
    const deltaGad = (row[idx.GAD7_POST] !== '' && row[idx.GAD7_PRE] !== '') ? row[idx.GAD7_POST] - row[idx.GAD7_PRE] : null;
    const deltaPhq = (row[idx.PHQ9_POST] !== '' && row[idx.PHQ9_PRE] !== '') ? row[idx.PHQ9_POST] - row[idx.PHQ9_PRE] : null;
    const deltaWho = (row[idx.WHOQOLBREF_POST] !== '' && row[idx.WHOQOLBREF_PRE] !== '') ? row[idx.WHOQOLBREF_POST] - row[idx.WHOQOLBREF_PRE] : null;

    return {
      pacienteId: row[idx.PacienteID],
      nombre: row[idx.Nombre],
      nhc: row[idx.NHC],
      estadoPacienteActual: row[idx.EstadoPacienteActual],
      tipoIntervencionPrincipal: row[idx.TipoIntervencionPrincipal],
      finTratamientoTexto: row[idx.FinTratamientoTexto],
      sexoGenero: row[idx.SexoGenero],
      edad: row[idx.Edad],
      motivoConsultaDiagnostico: row[idx.MotivoConsultaDiagnostico],
      numeroSesionesTotal: row[idx.NumeroSesionesTotal],
      tiempoEsperaHastaPrimeraConsultaDias: row[idx.TiempoEsperaHastaPrimeraConsultaDias],
      gad7Pre: row[idx.GAD7_PRE],
      gad7Post: row[idx.GAD7_POST],
      deltaGad7: deltaGad,
      phq9Pre: row[idx.PHQ9_PRE],
      phq9Post: row[idx.PHQ9_POST],
      deltaPhq9: deltaPhq,
      whoqolPre: row[idx.WHOQOLBREF_PRE],
      whoqolPost: row[idx.WHOQOLBREF_POST],
      deltaWhoqol: deltaWho
    };
  });

  const total = filas.length;
  const conPre = filas.filter(f => f.gad7Pre !== '' || f.phq9Pre !== '' || f.whoqolPre !== '').length;
  const conPost = filas.filter(f => f.gad7Post !== '' || f.phq9Post !== '' || f.whoqolPost !== '').length;
  
  // Helper para media
  const media = (arr) => {
    const vals = arr.filter(v => v !== null && v !== '').map(Number);
    return vals.length ? (vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(1) : 0;
  };

  return {
    filas: filas,
    resumen: {
      totalFichas: total,
      conPre: conPre,
      conPost: conPost,
      comparables: filas.filter(f => (f.gad7Pre !== '' && f.gad7Post !== '')).length,
      activos: filas.filter(f => f.estadoPacienteActual === 'ACTIVO').length,
      alta: filas.filter(f => f.estadoPacienteActual === 'ALTA').length,
      mediaGadPre: media(filas.map(f => f.gad7Pre)),
      mediaGadPost: media(filas.map(f => f.gad7Post)),
      deltaGad: media(filas.map(f => f.deltaGad7)),
      mediaPhqPre: media(filas.map(f => f.phq9Pre)),
      mediaPhqPost: media(filas.map(f => f.phq9Post)),
      deltaPhq: media(filas.map(f => f.deltaPhq9)),
      mediaWhoqolPre: media(filas.map(f => f.whoqolPre)),
      mediaWhoqolPost: media(filas.map(f => f.whoqolPost)),
      deltaWhoqol: media(filas.map(f => f.deltaWhoqol))
    }
  };
}

function sincronizarCamposAutomaticosFichaClinica_(pacienteId) {
  // Redirigimos al motor de sincronización masiva optimizado
  sincronizarFichasClinicasPacientes();
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName(SHEET_PACIENTES);
  const cSheet = ss.getSheetByName(SHEET_DATOS_CLINICOS_PACIENTES);
  const sSheet = ss.getSheetByName(SHEET_SESIONES);

  if (!pSheet || !cSheet || !sSheet) throw new Error('Faltan hojas críticas para sincronización.');

  const pData = pSheet.getDataRange().getValues();
  const cData = cSheet.getDataRange().getValues();
  const sData = sSheet.getDataRange().getValues();

  if (pData.length < 2) return { mensaje: 'No hay pacientes.' };

  const pIdx = indexByHeader_(pData[0]);
  const cIdx = indexByHeader_(cData[0]);
  const sIdx = indexByHeader_(sData[0]);

  // 1. Mapa de sesiones completadas por paciente (Batch Load)
  const sMap = {};
  for (let i = 1; i < sData.length; i++) {
    const pid = String(sData[i][sIdx.PacienteID]);
    const estado = sData[i][sIdx.EstadoSesion];
    if (estado === ESTADOS_SESION.COMPLETADA_AUTO || estado === ESTADOS_SESION.COMPLETADA_MANUAL) {
      sMap[pid] = (sMap[pid] || 0) + 1;
    }
  }

  // 2. Mapa de filas clínicas existentes (PacienteID -> índice en cData)
  const cMap = {};
  for (let i = 1; i < cData.length; i++) {
    cMap[String(cData[i][cIdx.PacienteID])] = i;
  }

  let procesados = 0;
  let nuevos = 0;

  // 3. Procesar todos los pacientes (Batch Process)
  for (let i = 1; i < pData.length; i++) {
    const pid = String(pData[i][pIdx.PacienteID]);
    if (!pid) continue;

    const p = pData[i];
    let cRowIdx = cMap[pid];
    let row;

    if (cRowIdx !== undefined) {
      row = cData[cRowIdx];
    } else {
      row = new Array(cData[0].length).fill('');
      row[cIdx.PacienteID] = pid;
      cData.push(row);
      cRowIdx = cData.length - 1;
      nuevos++;
    }

    // Actualizar campos calculados/mapeados
    row[cIdx.Nombre] = p[pIdx.Nombre] || '';
    row[cIdx.NHC] = p[pIdx.NHC] || row[cIdx.NHC] || '';
    row[cIdx.SexoGenero] = p[pIdx.SexoGenero] || '';
    row[cIdx.MotivoConsultaDiagnostico] = p[pIdx.MotivoConsultaDiagnostico] || '';
    row[cIdx.MotivoConsultaOtros] = p[pIdx.MotivoConsultaOtros] || '';
    row[cIdx.FechaAltaPrograma] = p[pIdx.FechaAlta] || '';
    row[cIdx.FechaPrimeraConsulta] = p[pIdx.FechaPrimeraConsulta] || '';
    row[cIdx.FechaAltaEfectiva] = p[pIdx.FechaAltaEfectiva] || p[pIdx.FechaCierre] || '';
    row[cIdx.EstadoPacienteActual] = p[pIdx.EstadoPaciente] || '';
    row[cIdx.TipoIntervencionPrincipal] = p[pIdx.ModalidadSolicitada] === MODALIDADES.INDIVIDUAL ? 'Individual' : 'Grupal';
    row[cIdx.FinTratamientoCodigo] = p[pIdx.EstadoPaciente] === ESTADOS_PACIENTE.ALTA ? p[pIdx.MotivoAltaCodigo] : 7;
    row[cIdx.FinTratamientoTexto] = p[pIdx.EstadoPaciente] === ESTADOS_PACIENTE.ALTA ? p[pIdx.MotivoAltaTexto] : 'Activo en el programa';
    row[cIdx.NumeroSesionesTotal] = (p[pIdx.FechaPrimeraConsulta] instanceof Date ? 1 : 0) + (sMap[pid] || 0);
    
    if (p[pIdx.FechaPrimeraConsulta] instanceof Date && p[pIdx.FechaPrimeraSesionReal] instanceof Date) {
      row[cIdx.TiempoEsperaHastaPrimeraConsultaDias] = diferenciaDiasFechas_(p[pIdx.FechaPrimeraConsulta], p[pIdx.FechaPrimeraSesionReal]);
    }

    procesados++;
  }

  // 4. Volcado masivo a la hoja clínica
  cSheet.getRange(1, 1, cData.length, cData[0].length).setValues(cData);
  return { mensaje: `Sincronización completada. Procesados: ${procesados} (${nuevos} nuevos).` };
}

function ejecutarSincronizarFichasClinicasPacientes() {
  const res = sincronizarFichasClinicasPacientes();
  SpreadsheetApp.getUi().alert(res.mensaje);
}
