/***********************
 * BLOQUE 11
 * HOME / DASHBOARD HTML
 ***********************/

function abrirHomeDashboard() {
  const html = HtmlService
    .createHtmlOutputFromFile('HomeDashboard')
    .setWidth(1280)
    .setHeight(820);

  SpreadsheetApp.getUi().showModalDialog(html, 'Panel de control');
}

function obtenerDatosHomeDashboard() {
  // `recalcularMetricasBasicas_()` puede ser costoso si hay muchas sesiones.
  // Cacheamos el recálculo breve para que el panel sea fluido.
  const cache = CacheService.getScriptCache();
  const cacheKey = 'metricasBasicas_home_lastRun';
  const ttlSeconds = 60;

  const yaReciente = cache.get(cacheKey);
  if (!yaReciente) {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(1500)) {
      try {
        const yaReciente2 = cache.get(cacheKey);
        if (!yaReciente2) {
          const stateService = new StateService();
          stateService.runAutomaticTransitions(); // Esto actualiza las métricas de pacientes
          cache.put(cacheKey, '1', ttlSeconds);
        }
      } finally {
        lock.releaseLock();
      }
    }
  }

  const patientRepo = new PatientRepository();
  const cicloRepo = new CicloRepository();
  const sessionRepo = new SessionRepository();
  const configRepo = new ConfigRepository();

  const pacientes = patientRepo.findAll();
  const ciclos = cicloRepo.findAll();
  const sesiones = sessionRepo.findAll();
  const modalidadesCfg = configRepo.findAll();

  const hoy = normalizarFecha_(new Date());

  const resumen = {
    totalPacientes: pacientes.length,
    activos: pacientes.filter(p => p.EstadoPaciente === 'ACTIVO').length,
    espera: pacientes.filter(p => p.EstadoPaciente === 'ESPERA').length,
    alta: pacientes.filter(p => p.EstadoPaciente === 'ALTA').length,
    pendienteInicio: pacientes.filter(p => p.EstadoPaciente === 'ACTIVO_PENDIENTE_INICIO').length,
    gruposEnCurso: ciclos.filter(c => c.EstadoCiclo === 'EN_CURSO').length,
    sesionesPendientes: sesiones.filter(s => s.EstadoSesion === 'PENDIENTE').length,
    sesionesErrorSync: sesiones.filter(s => s.CalendarSyncStatus === 'ERROR').length
  };

  const ocupacionPorModalidad = Object.values(MODALIDADES).map(modalidad => {
    if (modalidad === MODALIDADES.INDIVIDUAL) {
      const cfg = modalidadesCfg.find(c => c.Modalidad === modalidad) || {};
      const capacidad = Number(cfg.CapacidadMaxima || 0);
      const ocupadas = pacientes.filter(p =>
        p.ModalidadSolicitada === modalidad &&
        p.EstadoPaciente === 'ACTIVO'
      ).length;

      return {
        modalidad,
        capacidad,
        ocupadas,
        libres: Math.max(0, capacidad - ocupadas),
        porcentaje: capacidad > 0 ? Math.round((ocupadas / capacidad) * 100) : 0
      };
    }

    const ciclosVigentes = ciclos.filter(c =>
      c.Modalidad === modalidad &&
      (c.EstadoCiclo === 'PLANIFICADO' || c.EstadoCiclo === 'EN_CURSO')
    );

    if (ciclosVigentes.length === 0) {
      return {
        modalidad,
        capacidad: 0,
        ocupadas: 0,
        libres: 0,
        porcentaje: 0
      };
    }

    const cicloReferencia = ciclosVigentes
      .slice()
      .sort((a, b) => {
        const tA = a.FechaInicioCiclo instanceof Date ? a.FechaInicioCiclo.getTime() : 0;
        const tB = b.FechaInicioCiclo instanceof Date ? b.FechaInicioCiclo.getTime() : 0;
        return tA - tB;
      })[0];

    const capacidad = Number(cicloReferencia.CapacidadMaxima || 0);
    const ocupadas = Number(cicloReferencia.PlazasOcupadas || 0);

    return {
      modalidad,
      capacidad,
      ocupadas,
      libres: Math.max(0, capacidad - ocupadas),
      porcentaje: capacidad > 0 ? Math.round((ocupadas / capacidad) * 100) : 0
    };
  });

  const estadoCiclos = {
    planificados: ciclos.filter(c => c.EstadoCiclo === 'PLANIFICADO').length,
    enCurso: ciclos.filter(c => c.EstadoCiclo === 'EN_CURSO').length,
    cerrados: ciclos.filter(c => c.EstadoCiclo === 'CERRADO').length,
    cancelados: ciclos.filter(c => c.EstadoCiclo === 'CANCELADO').length
  };

  const esperaPorModalidad = Object.values(MODALIDADES).map(mod => ({
    modalidad: mod,
    total: pacientes.filter(p => p.ModalidadSolicitada === mod && p.EstadoPaciente === 'ESPERA').length
  }));

  const proximosCiclos = ciclos
    .filter(c => c.EstadoCiclo === 'PLANIFICADO')
    .sort((a, b) => {
      const tA = a.FechaInicioCiclo instanceof Date ? a.FechaInicioCiclo.getTime() : Infinity;
      const tB = b.FechaInicioCiclo instanceof Date ? b.FechaInicioCiclo.getTime() : Infinity;
      return tA - tB;
    })
    .slice(0, 8)
    .map(c => ({ 
      CicloID: c.CicloID, 
      Modalidad: c.Modalidad, 
      NumeroCiclo: c.NumeroCiclo, 
      FechaInicioCiclo: formatearFecha_(c.FechaInicioCiclo),
      PlazasLibres: c.PlazasLibres 
    }));

  const proximosPacientes = pacientes
    .filter(p => p.EstadoPaciente === 'ACTIVO' || p.EstadoPaciente === 'ACTIVO_PENDIENTE_INICIO')
    .sort((a, b) => {
      const tA = a.ProximaSesion instanceof Date ? a.ProximaSesion.getTime() : Infinity;
      const tB = b.ProximaSesion instanceof Date ? b.ProximaSesion.getTime() : Infinity;
      return tA - tB;
    })
    .slice(0, 10)
    .map(p => ({ 
      PacienteID: p.PacienteID, 
      Nombre: p.Nombre, 
      EstadoPaciente: p.EstadoPaciente, 
      ProximaSesion: formatearFecha_(p.ProximaSesion) 
    }));

  const alertas = construirAlertasHome_(pacientes, ciclos, sesiones, hoy);

  return {
    fechaHoy: formatearFecha_(hoy),
    fechaActualizacion: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    resumen,
    ocupacionPorModalidad,
    estadoCiclos,
    esperaPorModalidad,
    proximosCiclos,
    proximosPacientes,
    resumenIncidenciasCalendario: obtenerResumenIncidenciasCalendario(sesiones),
    calendarUrl: (typeof obtenerCalendarConsultaUrl_ === 'function') ? obtenerCalendarConsultaUrl_() : null,
    alertas
  };
}

function construirAlertasHome_(pacientes, ciclos, sesiones, hoy) {
  const alertas = [];

  const esperaSinCiclo = pacientes.filter(p =>
    p.EstadoPaciente === 'ESPERA' &&
    (!p.CicloObjetivoID && !p.CicloActivoID)
  ).length;

  if (esperaSinCiclo > 0) {
    alertas.push({
      tipo: 'warning',
      titulo: 'En espera sin ciclo',
      detalle: String(esperaSinCiclo)
    });
  }

  const erroresSync = sesiones.filter(s => s.CalendarSyncStatus === 'ERROR').length;
  if (erroresSync > 0) {
    alertas.push({
      tipo: 'danger',
      titulo: 'Errores Sync Calendar',
      detalle: String(erroresSync)
    });
  }

  const ciclosLlenos = ciclos.filter(c =>
    (c.EstadoCiclo === 'PLANIFICADO' || c.EstadoCiclo === 'EN_CURSO') &&
    Number(c.PlazasLibres || 0) <= 0
  ).length;

  if (ciclosLlenos > 0) {
    alertas.push({
      tipo: 'warning',
      titulo: 'Ciclos llenos',
      detalle: String(ciclosLlenos)
    });
  }

  [MODALIDADES.GRUPO_1, MODALIDADES.GRUPO_2, MODALIDADES.GRUPO_3].forEach(mod => {
    const tienePlanificado = ciclos.some(c =>
      c.Modalidad === mod &&
      c.EstadoCiclo === 'PLANIFICADO'
    );

    const esperaMod = pacientes.filter(p =>
      p.ModalidadSolicitada === mod && p.EstadoPaciente === 'ESPERA'
    ).length;

    if (esperaMod > 0 && !tienePlanificado) {
      alertas.push({
        tipo: 'warning',
        titulo: `Espera sin planificar: ${mod}`,
        detalle: mod + ' (' + esperaMod + ' en espera)'
      });
    }
  });

  const inicianPronto = pacientes.filter(p => {
    if (p.EstadoPaciente !== 'ACTIVO_PENDIENTE_INICIO') return false;
    const fecha = p.ProximaSesion;
    if (!fecha) return false;
    const diff = Math.round((fecha.getTime() - hoy.getTime()) / (24 * 60 * 60 * 1000));
    return diff >= 0 && diff <= 7;
  }).length;

  if (inicianPronto > 0) {
    alertas.push({
      tipo: 'info',
      titulo: 'Inicios próximos (7d)',
      detalle: String(inicianPronto)
    });
  }

  return alertas;
}
// Eliminadas funciones de mapeo redundantes y duplicados

function actualizarDashboard() {
  const sheetDiasBloqueados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DIAS_BLOQUEADOS');
  const sheetSesiones = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SESIONES');

  const diasData = sheetDiasBloqueados.getDataRange().getValues();
  const sesionesData = sheetSesiones.getDataRange().getValues();

  const diasBloqueados = diasData.filter(row => row[1] === true);  // Filtrar solo los días bloqueados
  const sesiones = sesionesData.filter(row => row[7] === 'ACTIVO');  // Filtrar solo las sesiones activas

  // Actualizar la sección del Dashboard con los días bloqueados
  const dashboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD');
  const diasRange = dashboardSheet.getRange('A1');  // Asumimos que queremos mostrarlo en la celda A1
  const diasTexto = diasBloqueados.map(dia => `${dia[0]}: ${dia[2] || 'Sin motivo'}`).join('\n');
  diasRange.setValue(`Días Bloqueados:\n${diasTexto}`);

  // Actualizar la sección del Dashboard con las sesiones
  const sesionesRange = dashboardSheet.getRange('B1');  // Asumimos que las sesiones irán en la celda B1
  const sesionesTexto = sesiones.map(sesion => `Sesion ${sesion[4]}: ${sesion[3]}`).join('\n');
  sesionesRange.setValue(`Sesiones Activas:\n${sesionesTexto}`);
}


/***************
 * LANZADORES DESDE HOME
 ***************/
function homeNuevoPaciente() { nuevoPaciente(); }
function homeEditarPaciente() { editarPaciente(); }
function homeEliminarPacientePorError() { eliminarPacientePorError(); }
function homeCrearCicloGrupo() { crearCicloGrupo(); }
function homeGestionEsperaYCicloPaciente() { gestionarEsperaYCicloPaciente(); }
function homeAbrirPantallaPacientes() { abrirPantallaPacientes(); }
function homeAbrirPantallaCiclos() { abrirPantallaCiclos(); }
function homeAbrirPantallaSesiones() { abrirPantallaSesiones(); }
function homeGestionarConfigModalidades() { gestionarConfigModalidades(); }
function homeGestionarCatalogos() { gestionarCatalogos(); }
function homeGestionarDiasBloqueados() { gestionarDiasBloqueados(); }
function homeSincronizarGoogleCalendar() {
  const calendar = obtenerOCrearCalendarioConsulta_(); // Se crea uno por usuario
  // Sincronizamos sesiones
  sincronizarSesionesAGoogleCalendar(calendar);
  // Pasamos el calendar explícitamente para evitar ambigüedades de firma
  sincronizarDiasBloqueadosAGoogleCalendar(calendar);
}
function homeActualizarEstadosAutomaticos() { actualizarEstadosAutomaticos(); }
function homeRefrescarDashboardHoja() { refrescarDashboard(); }
function abrirHomeDashboardDesdePantalla() {  abrirHomeDashboard(); }
function homeAbrirReprogramarSesion() { abrirReprogramarSesion(); }
function homeAltaPaciente() { altaPaciente(); }
function refrescarPanel() {  abrirHomeDashboard(); }
function homeRecalcularEstadosAutomaticamente() { recalcularEstadosAutomaticamenteConModal(); }
function homeFichaClinicaPaciente() { fichaClinicaPaciente(); }
function homeVerIncidenciasCalendario() { verIncidenciasCalendario(); }
function homeObtenerResumenIncidenciasCalendario() { return obtenerResumenIncidenciasCalendario(); }

function volverAlPanelDesdeDiasBloqueados() {
  abrirHomeDashboard(); // Re-abre el panel de control
}