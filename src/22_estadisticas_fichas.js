/***********************
 * BLOQUE 22
 * ESTADÍSTICAS FICHAS PACIENTES
 ***********************/

function estadisticasFichasPacientes() {
  const html = HtmlService
    .createHtmlOutputFromFile('EstadisticasFichasForm')
    .setWidth(1650)
    .setHeight(820);

  SpreadsheetApp.getUi().showModalDialog(html, 'Estadísticas fichas pacientes');
}

function obtenerDatosEstadisticasFichasFormulario() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATOS_CLINICOS_PACIENTES');
  if (!sheet) {
    throw new Error('No existe la hoja DATOS_CLINICOS_PACIENTES.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return {
      filas: [],
      resumen: {
        totalFichas: 0,
        conPre: 0,
        conPost: 0,
        comparables: 0,
        activos: 0,
        alta: 0,
        mediaGadPre: '',
        mediaGadPost: '',
        deltaGad: '',
        mediaPhqPre: '',
        mediaPhqPost: '',
        deltaPhq: '',
        mediaWhoqolPre: '',
        mediaWhoqolPost: '',
        deltaWhoqol: ''
      }
    };
  }

  const idx = indexByHeader_(data[0]);
  const filas = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const gadPre = convertirNumeroFicha_(row[idx.GAD7_PRE]);
    const gadPost = convertirNumeroFicha_(row[idx.GAD7_POST]);
    const phqPre = convertirNumeroFicha_(row[idx.PHQ9_PRE]);
    const phqPost = convertirNumeroFicha_(row[idx.PHQ9_POST]);
    const whoqolPre = convertirNumeroFicha_(row[idx.WHOQOLBREF_PRE]);
    const whoqolPost = convertirNumeroFicha_(row[idx.WHOQOLBREF_POST]);

    filas.push({
      pacienteId: row[idx.PacienteID] || '',
      nombre: row[idx.Nombre] || '',
      nhc: row[idx.NHC] || '',
      estadoPacienteActual: row[idx.EstadoPacienteActual] || '',
      tipoIntervencionPrincipal: row[idx.TipoIntervencionPrincipal] || '',
      finTratamientoTexto: row[idx.FinTratamientoTexto] || '',
      sexoGenero: row[idx.SexoGenero] || '',
      edad: row[idx.Edad] || '',
      nivelEstudios: row[idx.NivelEstudios] || '',
      motivoConsultaDiagnostico: row[idx.MotivoConsultaDiagnostico] || '',
      numeroSesionesTotal: convertirNumeroFicha_(row[idx.NumeroSesionesTotal]),
      tiempoEsperaHastaPrimeraConsultaDias: convertirNumeroFicha_(row[idx.TiempoEsperaHastaPrimeraConsultaDias]),
      psicofarmacos: row[idx.Psicofarmacos] || '',
      situacionLaboralPrevia: row[idx.SituacionLaboralPrevia] || '',
      gad7Pre: gadPre,
      gad7Post: gadPost,
      deltaGad7: calcularDeltaFicha_(gadPre, gadPost),
      phq9Pre: phqPre,
      phq9Post: phqPost,
      deltaPhq9: calcularDeltaFicha_(phqPre, phqPost),
      whoqolPre: whoqolPre,
      whoqolPost: whoqolPost,
      deltaWhoqol: calcularDeltaFicha_(whoqolPre, whoqolPost)
    });
  }

  const resumen = construirResumenEstadisticasFichas_(filas);

  return {
    filas: filas,
    resumen: resumen
  };
}

function construirResumenEstadisticasFichas_(filas) {
  const totalFichas = filas.length;
  const conPre = filas.filter(f =>
    f.gad7Pre !== null || f.phq9Pre !== null || f.whoqolPre !== null
  ).length;

  const conPost = filas.filter(f =>
    f.gad7Post !== null || f.phq9Post !== null || f.whoqolPost !== null
  ).length;

  const comparables = filas.filter(f =>
    (f.gad7Pre !== null && f.gad7Post !== null) ||
    (f.phq9Pre !== null && f.phq9Post !== null) ||
    (f.whoqolPre !== null && f.whoqolPost !== null)
  ).length;

  const activos = filas.filter(f => f.estadoPacienteActual === 'ACTIVO').length;
  const alta = filas.filter(f => f.estadoPacienteActual === 'ALTA').length;

  return {
    totalFichas: totalFichas,
    conPre: conPre,
    conPost: conPost,
    comparables: comparables,
    activos: activos,
    alta: alta,
    mediaGadPre: calcularMediaFicha_(filas.map(f => f.gad7Pre)),
    mediaGadPost: calcularMediaFicha_(filas.map(f => f.gad7Post)),
    deltaGad: calcularMediaFicha_(filas.map(f => f.deltaGad7)),
    mediaPhqPre: calcularMediaFicha_(filas.map(f => f.phq9Pre)),
    mediaPhqPost: calcularMediaFicha_(filas.map(f => f.phq9Post)),
    deltaPhq: calcularMediaFicha_(filas.map(f => f.deltaPhq9)),
    mediaWhoqolPre: calcularMediaFicha_(filas.map(f => f.whoqolPre)),
    mediaWhoqolPost: calcularMediaFicha_(filas.map(f => f.whoqolPost)),
    deltaWhoqol: calcularMediaFicha_(filas.map(f => f.deltaWhoqol))
  };
}

function convertirNumeroFicha_(valor) {
  if (valor === '' || valor === null || valor === undefined) return null;
  const n = Number(valor);
  return isNaN(n) ? null : n;
}

function calcularDeltaFicha_(pre, post) {
  if (pre === null || post === null) return null;
  return post - pre;
}

function calcularMediaFicha_(valores) {
  const nums = (valores || []).filter(v => v !== null && !isNaN(v));
  if (!nums.length) return '';
  const suma = nums.reduce((a, b) => a + b, 0);
  return Math.round((suma / nums.length) * 100) / 100;
}