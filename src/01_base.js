/***********************
 * BLOQUE 1
 * BASE ESTRUCTURAL V2
 ***********************/

/***************
 * HOJAS
 ***************/
const SHEET_CATALOGOS = 'CATALOGOS';
const SHEET_CONFIG_MODALIDADES = 'CONFIG_MODALIDADES';
const SHEET_PACIENTES = 'PACIENTES';
const SHEET_CICLOS = 'CICLOS';
const SHEET_ASIGNACIONES_CICLO = 'ASIGNACIONES_CICLO';
const SHEET_SESIONES = 'SESIONES';
const SHEET_DASHBOARD = 'DASHBOARD';

/***************
 * CALENDARIO
 ***************/
const GOOGLE_CALENDAR_NAME = 'Consulta Psicologia';

/***************
 * MODALIDADES
 ***************/
const MODALIDADES = {
  INDIVIDUAL: 'INDIVIDUAL',
  GRUPO_1: 'GRUPO_1',
  GRUPO_2: 'GRUPO_2',
  GRUPO_3: 'GRUPO_3'
};

const TIPOS_MODALIDAD = {
  INDIVIDUAL: 'INDIVIDUAL',
  GRUPO: 'GRUPO'
};

/***************
 * ESTADOS PACIENTE
 ***************/
const ESTADOS_PACIENTE = {
  ACTIVO: 'ACTIVO',
  ACTIVO_PENDIENTE_INICIO: 'ACTIVO_PENDIENTE_INICIO',
  ESPERA: 'ESPERA',
  ALTA: 'ALTA'
};

/***************
 * ESTADOS CICLO
 ***************/
const ESTADOS_CICLO = {
  PLANIFICADO: 'PLANIFICADO',
  EN_CURSO: 'EN_CURSO',
  CERRADO: 'CERRADO',
  CANCELADO: 'CANCELADO'
};

/***************
 * ESTADOS ASIGNACIÓN
 ***************/
const ESTADOS_ASIGNACION = {
  RESERVADO: 'RESERVADO',
  ACTIVO: 'ACTIVO',
  FINALIZADO: 'FINALIZADO',
  CANCELADO: 'CANCELADO'
};

/***************
 * ESTADOS SESIÓN
 ***************/
const ESTADOS_SESION = {
  PENDIENTE: 'PENDIENTE',
  COMPLETADA_AUTO: 'COMPLETADA_AUTO',
  COMPLETADA_MANUAL: 'COMPLETADA_MANUAL',
  REPROGRAMADA: 'REPROGRAMADA',
  CANCELADA: 'CANCELADA'
};

/***************
 * DÍAS SEMANA
 ***************/
const DIAS_SEMANA = {
  LUNES: 'LUNES',
  MARTES: 'MARTES',
  MIERCOLES: 'MIERCOLES',
  JUEVES: 'JUEVES',
  VIERNES: 'VIERNES',
  SABADO: 'SABADO',
  DOMINGO: 'DOMINGO'
};

/***************
 * ENCABEZADOS OFICIALES
 ***************/
const HEADERS = {
  [SHEET_CONFIG_MODALIDADES]: [
    'Modalidad',
    'TipoModalidad',
    'Activa',
    'DiaSemana',
    'FrecuenciaDias',
    'FechaBase',
    'CapacidadMaxima',
    'SesionesPorCiclo',
    'Notas'
  ],

  [SHEET_PACIENTES]: [
    'PacienteID',
    'Nombre',
    'ModalidadSolicitada',
    'FechaAlta',
    'FechaPrimeraConsulta',
    'EstadoPaciente',
    'MotivoEspera',
    'CicloObjetivoID',
    'CicloActivoID',
    'FechaPrimeraSesionReal',
    'SesionesPlanificadas',
    'SesionesCompletadas',
    'SesionesPendientes',
    'ProximaSesion',
    'FechaCierre',
    'Observaciones',
    'RecalcularSecuencia'
  ],

  [SHEET_CICLOS]: [
    'CicloID',
    'Modalidad',
    'NumeroCiclo',
    'EstadoCiclo',
    'FechaInicioCiclo',
    'FechaFinCiclo',
    'FechaBaseUsada',
    'DiaSemana',
    'FrecuenciaDias',
    'SesionesPorCiclo',
    'CapacidadMaxima',
    'PlazasOcupadas',
    'PlazasLibres',
    'GeneradoManual',
    'Notas'
  ],

  [SHEET_ASIGNACIONES_CICLO]: [
    'AsignacionID',
    'PacienteID',
    'CicloID',
    'Modalidad',
    'FechaAsignacion',
    'EstadoAsignacion',
    'Observaciones'
  ],

  [SHEET_SESIONES]: [
    'SesionID',
    'PacienteID',
    'CicloID',
    'AsignacionID',
    'Modalidad',
    'NombrePaciente',
    'NumeroSesion',
    'FechaSesion',
    'EstadoSesion',
    'FechaOriginal',
    'ModificadaManual',
    'Notas',
    'CalendarEventId',
    'CalendarSyncStatus',
    'CalendarLastSync',
    'CalendarEventTitle',
    'CalendarHash'
  ]
};

/***************
 * MENÚ
 ***************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Gestión Consulta');

  menu.addItem('Home / Panel de control', 'abrirHomeDashboard');

  menu.addSubMenu(
    ui.createMenu('Sistema')
      .addItem('Inicializar / verificar sistema', 'inicializarSistema')
      .addItem('Recargar catálogos', 'cargarCatalogosBase_')
      .addItem('Recargar configuración base modalidades', 'cargarConfiguracionModalidadesBase_')
      .addItem('Gestionar catálogos', 'gestionarCatalogos')
      .addItem('Gestionar configuración modalidades', 'gestionarConfigModalidades')
      .addItem('Gestionar días bloqueados', 'gestionarDiasBloqueados')
  );

  menu.addSubMenu(
  ui.createMenu('Operativa diaria')
    .addItem('Nuevo paciente', 'nuevoPaciente')
    .addItem('Editar paciente', 'editarPaciente')
    .addItem('Reprogramar sesión', 'abrirReprogramarSesion')
    .addItem('Alta de paciente', 'altaPaciente')
    .addItem('Eliminar paciente por error', 'eliminarPacientePorError')
    .addItem('Pantalla pacientes', 'abrirPantallaPacientes')
    .addItem('Pantalla ciclos', 'abrirPantallaCiclos')
    .addItem('Pantalla sesiones', 'abrirPantallaSesiones')
    .addItem('Crear ciclo de grupo', 'crearCicloGrupo')
    .addItem('Actualizar estados automáticos', 'actualizarEstadosAutomaticos')
    .addItem('Refrescar dashboard', 'refrescarDashboard')
  );

  menu.addSubMenu(
    ui.createMenu('Espera y ciclos')
      .addItem('Gestionar espera / cambio de grupo', 'gestionarEsperaYCicloPaciente')
      .addItem('Recalcular ocupación ciclos', 'recalcularOcupacionCiclos')
      .addItem('Generar sesiones faltantes', 'generarSesionesFaltantes')
  );

  menu.addSubMenu(
    ui.createMenu('Google Calendar')
      .addItem('Crear / obtener Google Calendar', 'crearCalendarioConsulta')
      .addItem('Sincronizar sesiones a Google Calendar', 'sincronizarSesionesAGoogleCalendar')
      .addItem('Sincronizar días bloqueados a Google Calendar', 'sincronizarDiasBloqueadosAGoogleCalendar')  // Nueva opción
      .addItem('Diagnóstico Google Calendar', 'diagnosticarGoogleCalendar')
      .addItem('Ver calendario actual vinculado', 'verCalendarioConsultaActual')
      .addItem('Resetear calendario vinculado', 'resetCalendarioConsultaVinculado')
      .addItem('Limpiar sync de Google Calendar', 'limpiarCamposSyncCalendarSesiones')
  );

  menu.addSubMenu(
    ui.createMenu('Automatización')
      .addItem('Crear trigger estados automáticos', 'crearTriggerEstadosAutomaticos')
      .addItem('Eliminar trigger estados automáticos', 'eliminarTriggerEstadosAutomaticos')
  );

  menu.addSubMenu(
  ui.createMenu('Desarrollador')
    .addItem('Recalcular estados', 'recalcularEstadosAutomaticamenteConModal')
    .addItem('Sincronizar fichas clínicas', 'ejecutarSincronizarFichasClinicasPacientes')
    .addSeparator()
    .addItem('Limpiar datos operativos (depuración)', 'ejecutarLimpiarDatosOperativosDepuracion_')
    .addItem('Reset total datos (depuración)', 'ejecutarResetProyectoDepuracion_')
    .addSeparator()
    .addItem('Limpiar integración Google Calendar', 'ejecutarLimpiarIntegracionCalendarDepuracion_')
    .addItem('Reset total completo', 'ejecutarResetEntornoCompletoDepuracion_')
  );  

  menu.addToUi();
}

/***************
 * INICIALIZACIÓN PRINCIPAL
 ***************/
function inicializarSistema() {
  const ui = SpreadsheetApp.getUi();

  try {
    crearHojasSiNoExisten_();
    asegurarEstructuraCompleta_();
    cargarCatalogosBase_();
    cargarConfiguracionModalidadesBase_();
    aplicarFormatoBasico_();

    ui.alert(
      'Sistema inicializado/verificado correctamente.\n\n' +
      'Se han revisado hojas, encabezados, catálogos y configuración base.'
    );
  } catch (error) {
    ui.alert('Error al inicializar el sistema: ' + error.message);
    throw error;
  }
}

/***************
 * CREACIÓN DE HOJAS SI FALTAN
 ***************/
function crearHojasSiNoExisten_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojasNecesarias = [
    SHEET_CATALOGOS,
    SHEET_CONFIG_MODALIDADES,
    SHEET_PACIENTES,
    SHEET_CICLOS,
    SHEET_ASIGNACIONES_CICLO,
    SHEET_SESIONES,
    SHEET_DASHBOARD
  ];

  const existentes = ss.getSheets().map(s => s.getName());

  hojasNecesarias.forEach(nombreHoja => {
    if (!existentes.includes(nombreHoja)) {
      ss.insertSheet(nombreHoja);
    }
  });
}

/***************
 * VERIFICAR / CREAR ENCABEZADOS
 ***************/
function asegurarEstructuraCompleta_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(HEADERS).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`No existe la hoja ${sheetName}.`);
    }

    asegurarEncabezadosExactos_(sheet, HEADERS[sheetName]);
  });

  asegurarHojaCatalogosEstructura_();
  asegurarHojaDashboardBase_();
}

/**
 * Si la hoja está vacía, escribe encabezados.
 * Si ya tiene encabezados, exige coincidencia exacta en orden y nombre.
 */
function asegurarEncabezadosExactos_(sheet, headersEsperados) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow === 0 || lastColumn === 0) {
    sheet.getRange(1, 1, 1, headersEsperados.length).setValues([headersEsperados]);
    return;
  }

  const headersActuales = sheet.getRange(1, 1, 1, Math.max(lastColumn, headersEsperados.length)).getValues()[0];

  const headersActualesRecortados = headersActuales.slice(0, headersEsperados.length);

  const coincideLongitud = lastColumn >= headersEsperados.length;
  const coincideContenido = headersEsperados.every((h, i) => headersActualesRecortados[i] === h);

  if (!coincideLongitud || !coincideContenido) {
    throw new Error(
      `Los encabezados de la hoja ${sheet.getName()} no coinciden con la estructura esperada.\n\n` +
      `Esperados:\n${headersEsperados.join(' | ')}\n\n` +
      `Actuales:\n${headersActualesRecortados.join(' | ')}`
    );
  }
}

function asegurarHojaCatalogosEstructura_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CATALOGOS);

  if (!sheet) {
    throw new Error(`No existe la hoja ${SHEET_CATALOGOS}.`);
  }

  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, 2).setValues([['Catalogo', 'Valor']]);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, 2).getValues()[0];
  const esperados = ['Catalogo', 'Valor'];

  if (headers[0] !== esperados[0] || headers[1] !== esperados[1]) {
    throw new Error(
      `Los encabezados de ${SHEET_CATALOGOS} deben ser exactamente: ${esperados.join(' | ')}`
    );
  }
}

function asegurarHojaDashboardBase_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DASHBOARD);

  if (!sheet) {
    throw new Error(`No existe la hoja ${SHEET_DASHBOARD}.`);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange('A1').setValue('DASHBOARD');
  }
}

/***************
 * CARGA DE CATÁLOGOS
 ***************/
function cargarCatalogosBase_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CATALOGOS);

  if (!sheet) {
    throw new Error(`No existe la hoja ${SHEET_CATALOGOS}.`);
  }

  const filas = [
    ['MODALIDADES', MODALIDADES.INDIVIDUAL],
    ['MODALIDADES', MODALIDADES.GRUPO_1],
    ['MODALIDADES', MODALIDADES.GRUPO_2],
    ['MODALIDADES', MODALIDADES.GRUPO_3],

    ['TIPOS_MODALIDAD', TIPOS_MODALIDAD.INDIVIDUAL],
    ['TIPOS_MODALIDAD', TIPOS_MODALIDAD.GRUPO],

    ['ESTADOS_PACIENTE', ESTADOS_PACIENTE.ACTIVO],
    ['ESTADOS_PACIENTE', ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO],
    ['ESTADOS_PACIENTE', ESTADOS_PACIENTE.ESPERA],
    ['ESTADOS_PACIENTE', ESTADOS_PACIENTE.ALTA],

    ['ESTADOS_CICLO', ESTADOS_CICLO.PLANIFICADO],
    ['ESTADOS_CICLO', ESTADOS_CICLO.EN_CURSO],
    ['ESTADOS_CICLO', ESTADOS_CICLO.CERRADO],
    ['ESTADOS_CICLO', ESTADOS_CICLO.CANCELADO],

    ['ESTADOS_ASIGNACION', ESTADOS_ASIGNACION.RESERVADO],
    ['ESTADOS_ASIGNACION', ESTADOS_ASIGNACION.ACTIVO],
    ['ESTADOS_ASIGNACION', ESTADOS_ASIGNACION.FINALIZADO],
    ['ESTADOS_ASIGNACION', ESTADOS_ASIGNACION.CANCELADO],

    ['ESTADOS_SESION', ESTADOS_SESION.PENDIENTE],
    ['ESTADOS_SESION', ESTADOS_SESION.COMPLETADA_AUTO],
    ['ESTADOS_SESION', ESTADOS_SESION.COMPLETADA_MANUAL],
    ['ESTADOS_SESION', ESTADOS_SESION.REPROGRAMADA],
    ['ESTADOS_SESION', ESTADOS_SESION.CANCELADA],

    ['DIAS_SEMANA', DIAS_SEMANA.LUNES],
    ['DIAS_SEMANA', DIAS_SEMANA.MARTES],
    ['DIAS_SEMANA', DIAS_SEMANA.MIERCOLES],
    ['DIAS_SEMANA', DIAS_SEMANA.JUEVES],
    ['DIAS_SEMANA', DIAS_SEMANA.VIERNES],
    ['DIAS_SEMANA', DIAS_SEMANA.SABADO],
    ['DIAS_SEMANA', DIAS_SEMANA.DOMINGO]
  ];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 2).setValues([['Catalogo', 'Valor']]);
  sheet.getRange(2, 1, filas.length, 2).setValues(filas);

  aplicarFormatoCatalogos_();
}

/***************
 * CONFIGURACIÓN BASE DE MODALIDADES
 ***************/
function cargarConfiguracionModalidadesBase_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CONFIG_MODALIDADES);

  if (!sheet) {
    throw new Error(`No existe la hoja ${SHEET_CONFIG_MODALIDADES}.`);
  }

  asegurarEncabezadosExactos_(sheet, HEADERS[SHEET_CONFIG_MODALIDADES]);

  const datosActuales = sheet.getDataRange().getValues();

  // Si solo tiene encabezados o está vacía de datos, cargamos base.
  if (datosActuales.length <= 1) {
    const filas = [
      [
        MODALIDADES.INDIVIDUAL,
        TIPOS_MODALIDAD.INDIVIDUAL,
        true,
        '',
        15,
        '',
        5,
        7,
        'La primera sesión real se calcula desde la primera consulta (+15 días naturales y ajuste a laborable).'
      ],
      [
        MODALIDADES.GRUPO_1,
        TIPOS_MODALIDAD.GRUPO,
        true,
        DIAS_SEMANA.MARTES,
        15,
        '',
        5,
        7,
        'Grupo alterno 1. Requiere fecha base. Solo entrada al inicio de ciclo.'
      ],
      [
        MODALIDADES.GRUPO_2,
        TIPOS_MODALIDAD.GRUPO,
        true,
        DIAS_SEMANA.MARTES,
        15,
        '',
        5,
        7,
        'Grupo alterno 2. Requiere fecha base. Solo entrada al inicio de ciclo.'
      ],
      [
        MODALIDADES.GRUPO_3,
        TIPOS_MODALIDAD.GRUPO,
        true,
        DIAS_SEMANA.JUEVES,
        7,
        '',
        5,
        7,
        'Grupo semanal. Requiere fecha base. Solo entrada al inicio de ciclo.'
      ]
    ];

    if (filas.length > 0) {
      sheet.getRange(2, 1, filas.length, filas[0].length).setValues(filas);
    }
  }

  aplicarFormatoConfigModalidades_();
}

/***************
 * FORMATO BÁSICO
 ***************/
function aplicarFormatoBasico_() {
  aplicarFormatoConfigModalidades_();
  aplicarFormatoPacientes_();
  aplicarFormatoCiclos_();
  aplicarFormatoAsignaciones_();
  aplicarFormatoSesiones_();
  aplicarFormatoCatalogos_();
}

function aplicarFormatoConfigModalidades_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG_MODALIDADES);
  if (!sheet) return;

  aplicarFormatoCabecera_(sheet, HEADERS[SHEET_CONFIG_MODALIDADES].length);
  sheet.setFrozenRows(1);

  const idx = indexByHeader_(HEADERS[SHEET_CONFIG_MODALIDADES]);
  if (idx.FechaBase !== undefined && sheet.getLastRow() > 1) {
    sheet.getRange(2, idx.FechaBase + 1, Math.max(sheet.getLastRow() - 1, 1), 1).setNumberFormat('dd/MM/yyyy');
  }
}

function aplicarFormatoPacientes_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PACIENTES);
  if (!sheet) return;

  aplicarFormatoCabecera_(sheet, HEADERS[SHEET_PACIENTES].length);
  sheet.setFrozenRows(1);

  const idx = indexByHeader_(HEADERS[SHEET_PACIENTES]);
  const lastRows = Math.max(sheet.getLastRow() - 1, 1);

  ['FechaAlta', 'FechaPrimeraConsulta', 'FechaPrimeraSesionReal', 'ProximaSesion', 'FechaCierre']
    .forEach(col => {
      if (idx[col] !== undefined) {
        sheet.getRange(2, idx[col] + 1, lastRows, 1).setNumberFormat('dd/MM/yyyy');
      }
    });
}

function aplicarFormatoCiclos_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CICLOS);
  if (!sheet) return;

  aplicarFormatoCabecera_(sheet, HEADERS[SHEET_CICLOS].length);
  sheet.setFrozenRows(1);

  const idx = indexByHeader_(HEADERS[SHEET_CICLOS]);
  const lastRows = Math.max(sheet.getLastRow() - 1, 1);

  ['FechaInicioCiclo', 'FechaFinCiclo', 'FechaBaseUsada']
    .forEach(col => {
      if (idx[col] !== undefined) {
        sheet.getRange(2, idx[col] + 1, lastRows, 1).setNumberFormat('dd/MM/yyyy');
      }
    });
}

function aplicarFormatoAsignaciones_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ASIGNACIONES_CICLO);
  if (!sheet) return;

  aplicarFormatoCabecera_(sheet, HEADERS[SHEET_ASIGNACIONES_CICLO].length);
  sheet.setFrozenRows(1);

  const idx = indexByHeader_(HEADERS[SHEET_ASIGNACIONES_CICLO]);
  const lastRows = Math.max(sheet.getLastRow() - 1, 1);

  if (idx.FechaAsignacion !== undefined) {
    sheet.getRange(2, idx.FechaAsignacion + 1, lastRows, 1).setNumberFormat('dd/MM/yyyy');
  }
}

function aplicarFormatoSesiones_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESIONES);
  if (!sheet) return;

  aplicarFormatoCabecera_(sheet, HEADERS[SHEET_SESIONES].length);
  sheet.setFrozenRows(1);

  const idx = indexByHeader_(HEADERS[SHEET_SESIONES]);
  const lastRows = Math.max(sheet.getLastRow() - 1, 1);

  ['FechaSesion', 'FechaOriginal', 'CalendarLastSync']
    .forEach(col => {
      if (idx[col] !== undefined) {
        sheet.getRange(2, idx[col] + 1, lastRows, 1).setNumberFormat('dd/MM/yyyy');
      }
    });
}

function aplicarFormatoCatalogos_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CATALOGOS);
  if (!sheet) return;

  aplicarFormatoCabecera_(sheet, 2);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);
}

function aplicarFormatoCabecera_(sheet, totalColumnas) {
  const range = sheet.getRange(1, 1, 1, totalColumnas);
  range.setFontWeight('bold');
  range.setBackground('#d9ead3');
  range.setFontColor('#1f1f1f');
  range.setHorizontalAlignment('center');
  range.setBorder(true, true, true, true, true, true);
}

/***************
 * HELPERS TÉCNICOS
 ***************/
function indexByHeader_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    map[h] = i;
  });
  return map;
}

function generarId_(prefijo) {
  const ahora = new Date();
  const stamp = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const rand = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
  return `${prefijo}_${stamp}_${rand}`;
}

function normalizarFecha_(fecha) {
  return new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
}

function sumarDiasNaturales_(fecha, dias) {
  const f = new Date(fecha);
  f.setDate(f.getDate() + dias);
  return normalizarFecha_(f);
}

function esFinDeSemana_(fecha) {
  const day = fecha.getDay(); // 0 domingo, 6 sábado
  return day === 0 || day === 6;
}

function moverASiguienteLaborable_(fecha) {
  let f = normalizarFecha_(fecha);
  while (esFinDeSemana_(f)) {
    f = sumarDiasNaturales_(f, 1);
  }
  return f;
}

function parseFechaES_(texto) {
  if (!texto) return null;
  if (texto instanceof Date) return normalizarFecha_(texto);

  // Soporte para formato ISO (yyyy-mm-dd) que viene de los inputs de HTML
  if (texto.includes('-')) {
    const partes = texto.split('-');
    return normalizarFecha_(new Date(partes[0], partes[1] - 1, partes[2]));
  }

  // Soporte para formato dd/mm/yyyy
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(texto).trim());
  if (!m) return null;

  const day = Number(m[1]);
  const month = Number(m[2]) - 1;
  const year = Number(m[3]);

  const fecha = new Date(year, month, day);

  if (
    fecha.getFullYear() !== year ||
    fecha.getMonth() !== month ||
    fecha.getDate() !== day
  ) {
    return null;
  }

  return normalizarFecha_(fecha);
}

/***************
 * STUB TEMPORAL
 * Se rehace más adelante
 ***************/
function refrescarDashboard() {
  construirDashboardReal_();
}

/***************
 * HELPERS DE ENTRADA
 ***************/
function pedirFecha_(ui, titulo, mensaje) {
  const resp = ui.prompt(titulo, mensaje, ui.ButtonSet.OK_CANCEL);

  if (resp.getSelectedButton() !== ui.Button.OK) return null;

  const texto = (resp.getResponseText() || '').trim();
  if (!texto) {
    ui.alert('La fecha es obligatoria.');
    return null;
  }

  const fecha = parseFechaES_(texto);
  if (!fecha) {
    ui.alert('Fecha no válida. Usa el formato dd/mm/yyyy, por ejemplo 20/03/2026.');
    return null;
  }

  return fecha;
}

function obtenerValoresCatalogo_(nombreCatalogo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CATALOGOS);
  if (!sheet) {
    throw new Error('No existe la hoja ' + SHEET_CATALOGOS + '.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const idx = indexByHeader_(headers);

  if (idx.Catalogo === undefined || idx.Valor === undefined) {
    throw new Error('La hoja CATALOGOS debe tener las columnas Catalogo y Valor.');
  }

  const valores = [];

  for (let i = 1; i < data.length; i++) {
    const catalogo = String(data[i][idx.Catalogo] || '').trim();
    const valor = String(data[i][idx.Valor] || '').trim();

    if (catalogo === nombreCatalogo && valor) {
      valores.push(valor);
    }
  }

  return valores;
}