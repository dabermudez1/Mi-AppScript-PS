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
const SHEET_AGENDA_PLANTILLA = 'AGENDA_PLANTILLA';
// El usuario ya ha añadido estas hojas y sus encabezados.
// Necesitamos asegurarnos de que el sistema las reconozca.
const SHEET_AGENDA_EXCEPCIONES = 'AGENDA_EXCEPCIONES';
const SHEET_DATOS_CLINICOS_PACIENTES = 'DATOS_CLINICOS_PACIENTES';

const PACIENTE_ID_RESERVA_21_GENERICA = 'SYS_RES_21';
const NOMBRE_RESERVA_21_GENERICA = 'RESERVA 2.1';
const NHC_RESERVA_21_GENERICA = 'SYS_RES_21';

// Objeto global para caché de ejecución. Evita lecturas repetidas a Sheets en el mismo script.
// Usamos 'var' para evitar el SyntaxError: Identifier has already been declared, 
// que es común en Apps Script cuando hay colisiones de archivos o push duplicados.
var __EXECUTION_CACHE__ = {
  [SHEET_PACIENTES]: null,
  [SHEET_SESIONES]: null,
  [SHEET_CICLOS]: null,
  [SHEET_DATOS_CLINICOS_PACIENTES]: null
};

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

const TIPOS_SESION_AGENDA = {
  S22: '2.2',
  S21: '2.1',
  GRUPO: '2.2/GRUPO',
  DESCANSO: 'DESCANSO',
  S21_RESERVA: '2.1/RESERVA' // Nuevo tipo para reservas provisionales de 2.1
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
  CANCELADA: 'CANCELADA',
  RESERVADA_PROVISIONAL: 'RESERVADA_PROVISIONAL' // Nuevo estado para reservas 2.1
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
    'HoraBase', // Nuevo campo para modalidades de grupo
    'CapacidadMaxima',
    'SesionesPorCiclo',
    'Notas'
  ],

  [SHEET_PACIENTES]: [
    'PacienteID',
    'Nombre',
    'NHC',
    'SexoGenero',
    'MotivoConsultaDiagnostico',
    'MotivoConsultaOtros',
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
    'FechaAltaEfectiva',
    'MotivoAltaCodigo',
    'MotivoAltaTexto',
    'ComentarioAlta',
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
    'Notas',
    'BloqueoInscripciones'
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
    'CalendarHash',
    'HoraInicio',
    'Duracion'
  ],

  [SHEET_AGENDA_PLANTILLA]: [
    'DiaSemana',
    'HoraInicio',
    'TipoSlot'
  ],

  [SHEET_AGENDA_EXCEPCIONES]: [
    'Fecha',
    'HoraInicio',
    'TipoSlot'
  ],

  [SHEET_DATOS_CLINICOS_PACIENTES]: [
    'PacienteID',
    'Nombre',
    'NHC',
    'FechaAltaPrograma',
    'FechaPrimeraConsulta',
    'FechaAltaEfectiva',
    'EstadoPacienteActual',
    'TipoIntervencionPrincipal',
    'FinTratamientoCodigo',
    'FinTratamientoTexto',
    'NumeroSesionesTotal',
    'TiempoEsperaHastaPrimeraConsultaDias',
    'SexoGenero',
    'Edad',
    'NivelEstudios',
    'MotivoConsultaDiagnostico',
    'MotivoConsultaOtros',
    'Comorbilidad',
    'AntecedentesSM',
    'Psicofarmacos',
    'SituacionLaboralPrevia',
    'CambioSituacionLaboralAlta',
    'CambioFarmacologicoAlta',
    'GAD7_PRE',
    'PHQ9_PRE',
    'WHOQOLBREF_FISICO_PRE',
    'WHOQOLBREF_PSICO_PRE',
    'WHOQOLBREF_SOCIAL_PRE',
    'WHOQOLBREF_AMBIENTE_PRE',
    'GAD7_POST',
    'PHQ9_POST',
    'WHOQOLBREF_FISICO_POST',
    'WHOQOLBREF_PSICO_POST',
    'WHOQOLBREF_SOCIAL_POST',
    'WHOQOLBREF_AMBIENTE_POST',
    'EscalaSatisfaccion',
    'OtrosComentarios'
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
      .addItem('Ver Disponibilidad Semanal', 'abrirPantallaDisponibilidad') // <-- NUEVA PANTALLA
      .addSeparator()
      .addItem('Gestionar Agenda (Slots y Excepciones)', 'gestionarAgenda')
  );

  menu.addSubMenu(
  ui.createMenu('Operativa diaria')
    .addItem('Nuevo paciente', 'nuevoPaciente')
    .addItem('Editar/Eliminar Reserva 2.1', 'editarEliminarReserva21') // Nuevo elemento de menú
    .addItem('Reservar 1ª Consulta (2.1)', 'reservarPrimeraConsulta') // Nuevo elemento de menú
    .addItem('Editar paciente', 'editarPaciente')
    .addItem('Reprogramar sesión', 'abrirReprogramarSesion')
    .addItem('Alta de paciente', 'altaPaciente')
    .addItem('Eliminar paciente por error', 'eliminarPacientePorError')
    .addItem('Pantalla pacientes', 'abrirPantallaPacientes')
    .addItem('Pantalla ciclos', 'abrirPantallaCiclos')
    .addItem('Pantalla sesiones', 'abrirPantallaSesiones')
    .addItem('Crear ciclo de grupo', 'crearCicloGrupo')
    .addItem('Actualizar estados automáticos', 'actualizarEstadosAutomaticos')
    .addItem('Generar sesiones faltantes (Mantenimiento)', 'generarSesionesFaltantes')
    .addItem('Refrescar panel y datos', 'refrescarDashboard')
  );

  menu.addSubMenu(
    ui.createMenu('Espera y ciclos')
      .addItem('Gestionar espera / cambio de grupo', 'gestionarEsperaYCicloPaciente')
      .addItem('Recalcular ocupación ciclos', 'recalcularOcupacionCiclos')
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
    .addItem('Auditar y resaltar columnas sobrantes', 'resaltarEstructuraNoOficial')
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
    const deseaResaltar = ui.alert(
      'Verificación del Sistema',
      '¿Deseas realizar también una auditoría visual resaltando en rojo las columnas no oficiales?',
      ui.ButtonSet.YES_NO
    ) === ui.Button.YES;

    // 1. Asegurar Hojas con encabezados fijos o sin estructura de datos
    asegurarEstructuraHoja_(SHEET_CATALOGOS, ['Catalogo', 'Valor'], deseaResaltar);
    asegurarEstructuraHoja_('DIAS_BLOQUEADOS', ['Fecha', 'Bloqueado', 'Motivo'], deseaResaltar);
    asegurarEstructuraHoja_(SHEET_DASHBOARD, [], false);

    // 2. Asegurar Hojas basadas en el esquema global HEADERS
    const esquemasPrincipales = [
      SHEET_CONFIG_MODALIDADES,
      SHEET_PACIENTES,
      SHEET_CICLOS,
      SHEET_ASIGNACIONES_CICLO,
      SHEET_SESIONES,
      SHEET_AGENDA_PLANTILLA,
      SHEET_AGENDA_EXCEPCIONES,
      SHEET_DATOS_CLINICOS_PACIENTES
    ];

    esquemasPrincipales.forEach(nombreHoja => {
      asegurarEstructuraHoja_(nombreHoja, HEADERS[nombreHoja], deseaResaltar);
    });

    // 3. Carga de datos base y formatos
    cargarCatalogosBase_();
    cargarConfiguracionModalidadesBase_();
    aplicarFormatoBasico_();

    ui.alert(
      'Verificación defensiva completada con éxito.\n\n' +
      'El sistema ha revisado las 9 hojas principales. Se han añadido columnas faltantes ' +
      'al final de las hojas existentes sin alterar los datos o el orden de los psicólogos.'
    );
  } catch (error) {
    ui.alert('Fallo en la inicialización defensiva: ' + error.message);
    throw error;
  }
}

/**
 * Crea una hoja si no existe y asegura que contenga los encabezados requeridos.
 * Si la hoja existe, añade al final de la fila 1 aquellos encabezados que falten.
 * Enfoque defensivo: Resalta en rojo los encabezados que no pertenecen al esquema oficial.
 */
function asegurarEstructuraHoja_(nombreHoja, encabezadosRequeridos, resaltar = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(nombreHoja);

  // 1. Crear la hoja si no existe
  if (!sheet) {
    sheet = ss.insertSheet(nombreHoja);
    if (encabezadosRequeridos && encabezadosRequeridos.length > 0) {
      sheet.getRange(1, 1, 1, encabezadosRequeridos.length).setValues([encabezadosRequeridos]);
    }
    return;
  }

  // 2. Si es el Dashboard o no hay requerimientos, no validamos estructura
  if (nombreHoja === SHEET_DASHBOARD || !encabezadosRequeridos || encabezadosRequeridos.length === 0) return;

  // 3. Verificar estructura existente
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const fila1 = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headersActuales = fila1.map(h => String(h).trim());

  // 4. Identificar encabezados faltantes
  const faltantes = encabezadosRequeridos.filter(req => !headersActuales.includes(req));

  if (faltantes.length > 0) {
    // Añadir columnas faltantes al final de la fila 1
    sheet.getRange(1, lastCol + 1, 1, faltantes.length).setValues([faltantes]);
    
    // Sincronizar cambios antes de proceder con la auditoría visual
    SpreadsheetApp.flush();
  }

  // 5. Auditoría Visual: Resaltar oficiales (Verde) y sobrantes (Rojo)
  const finalCol = resaltar ? sheet.getLastColumn() : 0;
  if (finalCol > 0) {
    const rangeFila1 = sheet.getRange(1, 1, 1, finalCol);
    const headersFinales = rangeFila1.getValues()[0];
    const backgrounds = headersFinales.map(h => {
      const nombre = String(h).trim();
      if (nombre === "") return null; // Sin formato para celdas vacías
      if (encabezadosRequeridos.includes(nombre)) return '#d9ead3'; // Verde suave (Oficial)
      return '#f4cccc'; // Rojo suave (No oficial/Sobrante)
    });
    
    rangeFila1.setBackgrounds([backgrounds]);
    rangeFila1.setFontWeight('bold');
    rangeFila1.setHorizontalAlignment('center');
  }
}

/**
 * Función administrativa para auditar visualmente todas las hojas.
 * Aplica el resaltado rojo a las columnas que no coinciden con HEADERS.
 */
function resaltarEstructuraNoOficial() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('Auditoría de Estructura', 
    'El sistema resaltará en ROJO los encabezados que no pertenezcan al diseño oficial.\n\n' +
    '¿Deseas continuar?', ui.ButtonSet.YES_NO);
  
  if (resp !== ui.Button.YES) return;

  try {
    // 1. Auditar hojas con encabezados fijos
    asegurarEstructuraHoja_(SHEET_CATALOGOS, ['Catalogo', 'Valor'], true);
    asegurarEstructuraHoja_('DIAS_BLOQUEADOS', ['Fecha', 'Bloqueado', 'Motivo'], true);

    // 2. Auditar resto de hojas según objeto HEADERS
    Object.keys(HEADERS).forEach(nombreHoja => {
      asegurarEstructuraHoja_(nombreHoja, HEADERS[nombreHoja], true);
    });
    ui.alert('Auditoría completada. Revisa los encabezados en rojo en las pestañas.');
  } catch (e) {
    ui.alert('Error durante la auditoría: ' + e.message);
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
        '09:00', // HoraBase añadida
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
        '09:00', // HoraBase añadida
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
        '17:00', // HoraBase añadida
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
  aplicarFormatoAgenda_(); // Nuevo
  aplicarFormatoExcepciones_(); // Nuevo
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

  if (idx.HoraInicio !== undefined) {
    sheet.getRange(2, idx.HoraInicio + 1, lastRows, 1).setNumberFormat('HH:mm');
  }
}

function aplicarFormatoCatalogos_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CATALOGOS);
  if (!sheet) return;

  aplicarFormatoCabecera_(sheet, 2);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);
}

function aplicarFormatoAgenda_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_AGENDA_PLANTILLA);
  if (!sheet) return;

  aplicarFormatoCabecera_(sheet, HEADERS[SHEET_AGENDA_PLANTILLA].length);
  sheet.setFrozenRows(1);

  const idx = indexByHeader_(HEADERS[SHEET_AGENDA_PLANTILLA]);
  const lastRows = Math.max(sheet.getLastRow() - 1, 1);

  if (idx.HoraInicio !== undefined) {
    sheet.getRange(2, idx.HoraInicio + 1, lastRows, 1).setNumberFormat('HH:mm');
  }
}

function aplicarFormatoExcepciones_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_AGENDA_EXCEPCIONES);
  if (!sheet) return;

  aplicarFormatoCabecera_(sheet, HEADERS[SHEET_AGENDA_EXCEPCIONES].length);
  sheet.setFrozenRows(1);

  const idx = indexByHeader_(HEADERS[SHEET_AGENDA_EXCEPCIONES]);
  const lastRows = Math.max(sheet.getLastRow() - 1, 1);

  sheet.getRange(2, idx.Fecha + 1, lastRows, 1).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(2, idx.HoraInicio + 1, lastRows, 1).setNumberFormat('HH:mm');
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

/**
 * Normaliza un objeto Date para incluir la hora, o combina una fecha y una cadena de hora.
 * Si se pasa solo `date`, se normaliza a la fecha y hora exactas.
 * Si se pasa `date` y `timeString`, combina la fecha de `date` con la hora de `timeString`.
 * @param {Date} date - El objeto Date base.
 * @param {string} [timeString] - Cadena de hora en formato "HH:mm".
 * @returns {Date} Un nuevo objeto Date con la fecha y hora normalizadas.
 */
function normalizarFechaHora_(date, timeString) {
  let d;
  // --- CORRECCIÓN CRÍTICA DE TIMEZONE ---
  // Si 'date' es un string (ej: "2026-10-20" de un input), new Date(date) lo interpreta como UTC.
  // Esto causa el bug de "un día antes" en zonas horarias positivas.
  // La solución es parsear el string manualmente para construir la fecha en la zona horaria del script.
  if (typeof date === 'string' && date.includes('-')) {
    const parts = date.split('T')[0].split('-');
    d = new Date(parts[0], parts[1] - 1, parts[2]);
  } else {
    d = (date instanceof Date) ? new Date(date.getTime()) : new Date(date);
  }
  
  if (isNaN(d.getTime())) {
    throw new Error('El primer argumento debe ser una fecha válida.');
  }

  if (timeString && String(timeString).trim() !== '') {
    // Usamos formatearHora_ para asegurar que tratamos con un string "HH:mm" 
    // incluso si timeString es un objeto Date de Sheets
    const sTime = formatearHora_(timeString);
    const partes = sTime.split(':');
    
    if (partes.length >= 2) {
      d.setHours(Number(partes[0]), Number(partes[1]), 0, 0);
    } 
    // Si el split falla, simplemente mantenemos la hora que ya trajera el objeto
  } else {
    // Si no se proporciona timeString, se normaliza la fecha y hora existentes
    d.setHours(d.getHours(), d.getMinutes(), 0, 0);
  }
  return d;
}

function sumarMinutos_(date, minutes) {
  return new Date(date.getTime() + minutes * 60 * 1000);
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

function compararFechasHoras_(date1, date2) {
  return date1.getTime() - date2.getTime();
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

/**
 * Parsea una cadena de fecha y hora en formato "dd/mm/yyyy HH:mm" o "dd/mm/yyyy" a un objeto Date.
 * @param {string} texto - La cadena de fecha y hora.
 * @returns {Date|null} Objeto Date si el parseo es exitoso, null en caso contrario.
 */
function parseFechaHoraES_(texto) {
  if (!texto) return null;
  if (texto instanceof Date) return texto; // Ya es un Date, devolver tal cual

  const partes = String(texto).trim().split(' ');
  const fechaParte = partes[0];
  const horaParte = partes[1] || '00:00'; // Si no hay hora, asumir medianoche

  const fecha = parseFechaES_(fechaParte);
  if (!fecha) return null;

  const [horas, minutos] = horaParte.split(':').map(Number);
  if (isNaN(horas) || isNaN(minutos)) return null;

  fecha.setHours(horas, minutos, 0, 0);
  return fecha;
}

/**
 * Formateador de alto rendimiento para UI. 
 * Evita la sobrecarga de Utilities.formatDate en bucles masivos.
 */
function formatearHora_(dateOrTimeString) {
  if (!dateOrTimeString && dateOrTimeString !== 0) return '';
  
  if (dateOrTimeString instanceof Date && !isNaN(dateOrTimeString.getTime())) {
    const h = dateOrTimeString.getHours().toString().padStart(2, '0');
    const m = dateOrTimeString.getMinutes().toString().padStart(2, '0');
    return `${h}:${m}`;
  }
  
  // Si es un string o número, intentamos extraer HH:mm
  const str = String(dateOrTimeString).trim();
  return str.match(/^\d{1,2}:\d{2}/) ? str.substring(0, 5) : (str.substring(0, 5) || '');
}

/***************
 * STUB TEMPORAL
 * Se rehace más adelante
 ***************/
/**
 * Punto de entrada unificado para refrescar todos los datos del sistema.
 * Limpia cachés de memoria y persistencia, y actualiza la hoja Dashboard.
 */
function refrescarDashboard() {
  // 1. Limpiar caché de memoria de la ejecución actual (TODAS)
  if (typeof __EXECUTION_CACHE__ !== 'undefined') {
    Object.keys(__EXECUTION_CACHE__).forEach(key => __EXECUTION_CACHE__[key] = null);
  }
  SpreadsheetApp.flush();

  // 2. Forzar recálculo de ocupación de ciclos (Mantenimiento profundo)
  try {
    if (typeof MaintenanceService === 'function') {
      new MaintenanceService().recalculateCycleOccupancy();
    }
  } catch (e) {
    console.warn("No se pudo completar el recálculo de ciclos: " + e.message);
  }

  // 3. Limpiar caché persistente (Dashboard HTML y lock de métricas)
  eliminarCacheDashboard_();

  // 4. Reconstruir la hoja Dashboard (Esto internamente ejecuta StateService.runAutomaticTransitions)
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

function pedirFechaHora_(ui, titulo, mensaje) {
  // Implementación similar a pedirFecha_, pero usando parseFechaHoraES_
  // y validando formato "dd/mm/yyyy HH:mm"
}

function obtenerValoresCatalogo_(nombreCatalogo) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `catalogo_${nombreCatalogo}`;
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* Fallback a lectura normal */ }
  }

  // Si no está en caché, leer de la hoja
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

  // Guardar en caché por 5 minutos (300 segundos)
  try {
    cache.put(cacheKey, JSON.stringify(valores), 300);
  } catch (e) { /* Ignorar errores de caché */ }

  return valores;
}

function esValorVerdadero_(val) {
  if (typeof val === 'boolean') return val;
  if (typeof val === 'string') {
    const s = val.trim().toUpperCase();
    return s === 'TRUE' || s === 'SÍ' || s === 'SI' || s === '1';
  }
  return !!val;
}

/**
 * Punto de entrada para la actualización de estados, usable desde el menú o triggers de tiempo.
 */
function actualizarEstadosAutomaticos() {
  try {
    const service = new StateService();
    const stats = service.runAutomaticTransitions();
    
    // Intentar mostrar notificación si hay una sesión de usuario activa
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Actualización automática finalizada. Sesiones procesadas: ${stats.sesiones}`, 
        "📦 Motor de Estados"
      );
    } catch(e) {
      // Ignorar si se ejecuta en segundo plano por un trigger
    }
    
    console.log(`Ejecución del motor de estados completada con éxito: ${JSON.stringify(stats)}`);
  } catch (error) {
    console.error("Error en la ejecución del motor de estados: " + error.message);
  }
}

/**
 * Configura un disparador (trigger) para que el sistema se actualice solo cada hora.
 */
function crearTriggerEstadosAutomaticos() {
  const ui = SpreadsheetApp.getUi();
  const handlers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  
  if (handlers.includes('actualizarEstadosAutomaticos')) {
    ui.alert('El trigger automático ya está configurado y en funcionamiento.');
    return;
  }

  ScriptApp.newTrigger('actualizarEstadosAutomaticos')
    .timeBased()
    .everyHours(1)
    .create();

  ui.alert('✅ Éxito: El sistema ahora se actualizará solo cada hora, incluso con la hoja cerrada.');
}

/**
 * Elimina los triggers automáticos existentes para el motor de estados.
 */
function eliminarTriggerEstadosAutomaticos() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'actualizarEstadosAutomaticos') {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  });
  SpreadsheetApp.getUi().alert(`Se han desactivado ${count} disparador(es) automáticos.`);
}