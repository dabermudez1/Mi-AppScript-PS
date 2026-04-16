/***********************
 * BLOQUE 00_DASHBOARD_WRITER
 * DASHBOARD REAL
 ***********************/

function construirDashboardReal_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DASHBOARD);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DASHBOARD);
  }

  sheet.clear();
  sheet.clearFormats();
  sheet.setHiddenGridlines(true);

  // Aseguramos que las métricas estén actualizadas antes de leer para el dashboard
  new StateService().runAutomaticTransitions();

  const data = obtenerDatosDashboard_();

  aplicarBaseDashboard_(sheet);

  escribirTituloDashboard_(sheet);
  escribirTarjetasResumen_(sheet, data);
  escribirBloqueIndividual_(sheet, data);
  escribirBloqueGrupos_(sheet, data);
  escribirBloqueEspera_(sheet, data);
  escribirBloquePacientesActivos_(sheet, data);
  escribirBloqueCiclosProximos_(sheet, data);

  ajustarLayoutDashboard_(sheet);
}

// La función `recalcularMetricasBasicas_` ha sido eliminada ya que su lógica
// está ahora integrada en `StateService.runAutomaticTransitions()`
// y se llama antes de obtener los datos del dashboard.

function obtenerDatosDashboard_() {
  // Usar repositorios para aprovechar la caché de ejecución (__EXECUTION_CACHE__)
  const patientRepo = new PatientRepository();
  const cicloRepo = new BaseRepository(SHEET_CICLOS, HEADERS[SHEET_CICLOS]);
  const configRepo = new BaseRepository(SHEET_CONFIG_MODALIDADES, HEADERS[SHEET_CONFIG_MODALIDADES]);

  // findAll() ya devuelve objetos mapeados y usa caché
  const pacientes = patientRepo.findAll();
  const ciclos = cicloRepo.findAll();
  const configArray = configRepo.findAll();
  
  // Convertir config a mapa para compatibilidad con lógica existente
  const config = {};
  configArray.forEach(c => { config[c.Modalidad] = c; });

  const hoy = normalizarFecha_(new Date());

  const pacientesActivos = pacientes.filter(p =>
    p.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO ||
    p.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO
  );

  const pacientesEspera = pacientes.filter(p => p.EstadoPaciente === ESTADOS_PACIENTE.ESPERA);
  const pacientesAlta = pacientes.filter(p => p.EstadoPaciente === ESTADOS_PACIENTE.ALTA);

  const individualCfg = config[MODALIDADES.INDIVIDUAL] || {};
  const individualActivos = pacientes.filter(p =>
    p.ModalidadSolicitada === MODALIDADES.INDIVIDUAL &&
    p.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO
  );

  const individualEspera = pacientes.filter(p =>
    p.ModalidadSolicitada === MODALIDADES.INDIVIDUAL &&
    p.EstadoPaciente === ESTADOS_PACIENTE.ESPERA
  );

  const individualCapacidad = Number(individualCfg.CapacidadMaxima || 0);
  const individualOcupadas = individualActivos.length;
  const individualLibres = Math.max(0, individualCapacidad - individualOcupadas);

  const grupos = [MODALIDADES.GRUPO_1, MODALIDADES.GRUPO_2, MODALIDADES.GRUPO_3].map(modalidad => {
    const ciclosModalidad = ciclos
      .filter(c => c.Modalidad === modalidad)
      .sort((a, b) => compararFechas_(a.FechaInicioCiclo, b.FechaInicioCiclo));

    const cicloActual = ciclosModalidad.find(c => c.EstadoCiclo === ESTADOS_CICLO.EN_CURSO) || null;

    const cicloPlanificado = ciclosModalidad.find(c =>
      c.EstadoCiclo === ESTADOS_CICLO.PLANIFICADO &&
      c.FechaInicioCiclo instanceof Date &&
      c.FechaInicioCiclo.getTime() >= hoy.getTime()
    ) || null;

    const espera = pacientes.filter(p =>
      p.ModalidadSolicitada === modalidad &&
      p.EstadoPaciente === ESTADOS_PACIENTE.ESPERA
    ).length;

    return {
      Modalidad: modalidad,
      CicloActual: cicloActual,
      ProximoCiclo: cicloPlanificado,
      Espera: espera
    };
  });

  const esperaPorModalidad = [
    MODALIDADES.INDIVIDUAL,
    MODALIDADES.GRUPO_1,
    MODALIDADES.GRUPO_2,
    MODALIDADES.GRUPO_3
  ].map(modalidad => ({
    Modalidad: modalidad,
    Total: pacientes.filter(p =>
      p.ModalidadSolicitada === modalidad &&
      p.EstadoPaciente === ESTADOS_PACIENTE.ESPERA
    ).length
  }));

  const topPacientesActivos = pacientesActivos
    .slice()
    .sort((a, b) => {
      const fechaA = a.ProximaSesion instanceof Date ? a.ProximaSesion.getTime() : Number.MAX_SAFE_INTEGER;
      const fechaB = b.ProximaSesion instanceof Date ? b.ProximaSesion.getTime() : Number.MAX_SAFE_INTEGER;
      if (fechaA !== fechaB) return fechaA - fechaB;
      return String(a.Nombre || '').localeCompare(String(b.Nombre || ''));
    })
    .slice(0, 12);

  const ciclosProximos = ciclos
    .filter(c => 
      c.EstadoCiclo === ESTADOS_CICLO.PLANIFICADO || c.EstadoCiclo === ESTADOS_CICLO.EN_CURSO
    )
    .sort((a, b) => compararFechas_(a.FechaInicioCiclo, b.FechaInicioCiclo))
    .slice(0, 12);

  return {
    hoy,
    pacientesActivos,
    pacientesEspera,
    pacientesAlta,
    individual: {
      capacidad: individualCapacidad,
      ocupadas: individualOcupadas,
      libres: individualLibres,
      espera: individualEspera.length
    },
    grupos,
    esperaPorModalidad,
    topPacientesActivos,
    ciclosProximos
  };
}

/***************
 * MAPEOS
 ***************/
function mapearPacientesDashboard_(data) {
  if (!data || data.length < 2) return [];

  const idx = indexByHeader_(data[0]);

  return data.slice(1).map(row => ({
    PacienteID: row[idx.PacienteID],
    Nombre: row[idx.Nombre],
    ModalidadSolicitada: row[idx.ModalidadSolicitada],
    EstadoPaciente: row[idx.EstadoPaciente],
    CicloObjetivoID: row[idx.CicloObjetivoID],
    CicloActivoID: row[idx.CicloActivoID],
    FechaPrimeraSesionReal: row[idx.FechaPrimeraSesionReal],
    SesionesPlanificadas: Number(row[idx.SesionesPlanificadas] || 0),
    SesionesCompletadas: Number(row[idx.SesionesCompletadas] || 0),
    SesionesPendientes: Number(row[idx.SesionesPendientes] || 0),
    ProximaSesion: row[idx.ProximaSesion]
  }));
}

function mapearCiclosDashboard_(data) {
  if (!data || data.length < 2) return [];

  const idx = indexByHeader_(data[0]);

  return data.slice(1).map(row => ({
    CicloID: row[idx.CicloID],
    Modalidad: row[idx.Modalidad],
    NumeroCiclo: Number(row[idx.NumeroCiclo] || 0),
    EstadoCiclo: row[idx.EstadoCiclo],
    FechaInicioCiclo: row[idx.FechaInicioCiclo],
    FechaFinCiclo: row[idx.FechaFinCiclo],
    CapacidadMaxima: Number(row[idx.CapacidadMaxima] || 0),
    PlazasOcupadas: Number(row[idx.PlazasOcupadas] || 0),
    PlazasLibres: Number(row[idx.PlazasLibres] || 0)
  }));
}

function mapearConfigDashboard_(data) {
  if (!data || data.length < 2) return {};

  const idx = indexByHeader_(data[0]);
  const out = {};

  data.slice(1).forEach(row => {
    const modalidad = row[idx.Modalidad];
    if (!modalidad) return;

    out[modalidad] = {
      TipoModalidad: row[idx.TipoModalidad],
      Activa: row[idx.Activa] === true,
      CapacidadMaxima: Number(row[idx.CapacidadMaxima] || 0),
      SesionesPorCiclo: Number(row[idx.SesionesPorCiclo] || 0)
    };
  });

  return out;
}

/***************
 * ESCRITURA DASHBOARD
 ***************/
function escribirTituloDashboard_(sheet) {
  sheet.getRange('A1:L1').merge();
  sheet.getRange('A1').setValue('DASHBOARD CONSULTA');
  sheet.getRange('A1').setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#b6d7a8')
    .setBorder(true, true, true, true, false, false);
}

function escribirTarjetasResumen_(sheet, data) {
  escribirTarjetaDashboard_(sheet, 'A3:B5', 'Pacientes activos', data.pacientesActivos.length, '#d9ead3');
  escribirTarjetaDashboard_(sheet, 'D3:E5', 'Pacientes en espera', data.pacientesEspera.length, '#fce5cd');
  escribirTarjetaDashboard_(sheet, 'G3:H5', 'Pacientes alta', data.pacientesAlta.length, '#d0e0e3');

  const gruposEnCurso = data.grupos.filter(g => g.CicloActual).length;
  escribirTarjetaDashboard_(sheet, 'J3:K5', 'Grupos en curso', gruposEnCurso, '#ead1dc');
}

function escribirBloqueIndividual_(sheet, data) {
  sheet.getRange('A7:D7').merge();
  sheet.getRange('A7').setValue('INDIVIDUAL');
  aplicarTituloBloqueDashboard_(sheet.getRange('A7:D7'));

  const headers = [['Capacidad', 'Ocupadas', 'Libres', 'Espera']];
  sheet.getRange(8, 1, 1, 4).setValues(headers);
  aplicarCabeceraDashboard_(sheet.getRange(8, 1, 1, 4));

  const row = [[
    data.individual.capacidad,
    data.individual.ocupadas,
    data.individual.libres,
    data.individual.espera
  ]];

  sheet.getRange(9, 1, 1, 4).setValues(row);
  aplicarTablaDashboard_(sheet.getRange(9, 1, 1, 4));
}

function escribirBloqueGrupos_(sheet, data) {
  sheet.getRange('F7:L7').merge();
  sheet.getRange('F7').setValue('SITUACIÓN DE GRUPOS');
  aplicarTituloBloqueDashboard_(sheet.getRange('F7:L7'));

  const headers = [[
    'Grupo',
    'Ciclo actual',
    'Estado',
    'Ocupadas',
    'Libres',
    'Inicio',
    'Espera'
  ]];
  sheet.getRange(8, 6, 1, 7).setValues(headers);
  aplicarCabeceraDashboard_(sheet.getRange(8, 6, 1, 7));

  const rows = data.grupos.map(g => {
    const ciclo = g.CicloActual || g.ProximoCiclo || null;

    return [
      g.Modalidad,
      ciclo ? ciclo.NumeroCiclo : '',
      ciclo ? ciclo.EstadoCiclo : 'SIN_CICLO',
      ciclo ? ciclo.PlazasOcupadas : '',
      ciclo ? ciclo.PlazasLibres : '',
      ciclo ? formatearFecha_(ciclo.FechaInicioCiclo) : '',
      g.Espera
    ];
  });

  if (rows.length > 0) {
    sheet.getRange(9, 6, rows.length, 7).setValues(rows);
    aplicarTablaDashboard_(sheet.getRange(9, 6, rows.length, 7));
  }
}

function escribirBloqueEspera_(sheet, data) {
  sheet.getRange('A12:D12').merge();
  sheet.getRange('A12').setValue('ESPERA POR MODALIDAD');
  aplicarTituloBloqueDashboard_(sheet.getRange('A12:D12'));

  sheet.getRange(13, 1, 1, 2).setValues([['Modalidad', 'Total espera']]);
  aplicarCabeceraDashboard_(sheet.getRange(13, 1, 1, 2));

  const rows = data.esperaPorModalidad.map(r => [r.Modalidad, r.Total]);
  sheet.getRange(14, 1, rows.length, 2).setValues(rows);
  aplicarTablaDashboard_(sheet.getRange(14, 1, rows.length, 2));
}

function escribirBloquePacientesActivos_(sheet, data) {
  sheet.getRange('F12:L12').merge();
  sheet.getRange('F12').setValue('PRÓXIMOS PACIENTES ACTIVOS');
  aplicarTituloBloqueDashboard_(sheet.getRange('F12:L12'));

  sheet.getRange(13, 6, 1, 7).setValues([[
    'Nombre',
    'Modalidad',
    'Estado',
    'Hechas',
    'Pendientes',
    'Próxima sesión',
    'Ciclo'
  ]]);
  aplicarCabeceraDashboard_(sheet.getRange(13, 6, 1, 7));

  const rows = data.topPacientesActivos.map(p => [
    p.Nombre || '',
    p.ModalidadSolicitada || '',
    p.EstadoPaciente || '',
    p.SesionesCompletadas || 0,
    p.SesionesPendientes || 0,
    p.ProximaSesion instanceof Date ? 
      formatearFecha_(p.ProximaSesion) + ' ' + formatearHora_(p.ProximaSesion) : 
      formatearFecha_(p.ProximaSesion),
    p.CicloActivoID || p.CicloObjetivoID || ''
  ]);

  if (rows.length > 0) {
    sheet.getRange(14, 6, rows.length, 7).setValues(rows);
    aplicarTablaDashboard_(sheet.getRange(14, 6, rows.length, 7));
  } else {
    sheet.getRange('F14:L14').merge();
    sheet.getRange('F14').setValue('No hay pacientes activos.');
    aplicarCajaVaciaDashboard_(sheet.getRange('F14:L14'));
  }
}

function escribirBloqueCiclosProximos_(sheet, data) {
  sheet.getRange('A20:L20').merge();
  sheet.getRange('A20').setValue('PRÓXIMOS GRUPOS (VIGENTES)');
  aplicarTituloBloqueDashboard_(sheet.getRange('A20:L20'));

  sheet.getRange(21, 1, 1, 7).setValues([[
    'Modalidad',
    'Ciclo',
    'Estado',
    'Inicio',
    'Fin',
    'Ocupadas',
    'Libres'
  ]]);
  aplicarCabeceraDashboard_(sheet.getRange(21, 1, 1, 7));

  const rows = data.ciclosProximos.map(c => [
    c.Modalidad,
    c.NumeroCiclo,
    c.EstadoCiclo,
    formatearFecha_(c.FechaInicioCiclo),
    formatearFecha_(c.FechaFinCiclo),
    c.PlazasOcupadas,
    c.PlazasLibres
  ]);

  if (rows.length > 0) {
    sheet.getRange(22, 1, rows.length, 7).setValues(rows);
    aplicarTablaDashboard_(sheet.getRange(22, 1, rows.length, 7));
  } else {
    sheet.getRange('A22:G22').merge();
    sheet.getRange('A22').setValue('No hay ciclos planificados.');
    aplicarCajaVaciaDashboard_(sheet.getRange('A22:G22'));
  }
}

/***************
 * ESTILOS
 ***************/
function aplicarBaseDashboard_(sheet) {
  sheet.getRange('A:L').setFontFamily('Arial');
  sheet.getRange('A:L').setBackground('#f7f7f7');
  sheet.setFrozenRows(1);
}

function escribirTarjetaDashboard_(sheet, a1Notation, titulo, valor, color) {
  const range = sheet.getRange(a1Notation);
  range.merge();
  range.setBackground(color);
  range.setBorder(true, true, true, true, false, false, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');

  const cell = range.getCell(1, 1);
  cell.setValue(titulo + '\n' + valor);
  cell.setFontSize(15);
  cell.setFontWeight('bold');
  cell.setWrap(true);
}

function aplicarTituloBloqueDashboard_(range) {
  range.setBackground('#d9d9d9');
  range.setFontWeight('bold');
  range.setFontSize(11);
  range.setBorder(true, true, true, true, false, false);
}

function aplicarCabeceraDashboard_(range) {
  range.setBackground('#ececec');
  range.setFontWeight('bold');
  range.setHorizontalAlignment('center');
  range.setBorder(true, true, true, true, true, true);
}

function aplicarTablaDashboard_(range) {
  range.setBackground('#ffffff');
  range.setBorder(true, true, true, true, true, true, '#d0d0d0', SpreadsheetApp.BorderStyle.SOLID);
  range.setVerticalAlignment('middle');
}

function aplicarCajaVaciaDashboard_(range) {
  range.setBackground('#ffffff');
  range.setFontStyle('italic');
  range.setHorizontalAlignment('center');
  range.setBorder(true, true, true, true, false, false);
}

function ajustarLayoutDashboard_(sheet) {
  const widths = {
    1: 150, 2: 110, 3: 110, 4: 120,
    5: 24,
    6: 170, 7: 110, 8: 110, 9: 90, 10: 90, 11: 115, 12: 170
  };

  Object.keys(widths).forEach(col => {
    sheet.setColumnWidth(Number(col), widths[col]);
  });

  sheet.getRange('A:L').setVerticalAlignment('middle');
}

/***************
 * HELPERS
 ***************/
function compararFechas_(a, b) {
  const tA = a instanceof Date ? a.getTime() : Number.MAX_SAFE_INTEGER;
  const tB = b instanceof Date ? b.getTime() : Number.MAX_SAFE_INTEGER;
  return tA - tB;
}