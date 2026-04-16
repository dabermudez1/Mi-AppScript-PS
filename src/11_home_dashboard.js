/***********************
 * BLOQUE 11
 * HOME / DASHBOARD HTML
 ***********************/

function abrirHomeDashboard() {
  const html = HtmlService
    .createHtmlOutputFromFile('HomeDashboard')
    .setWidth(1280)
    .setHeight(970);

  SpreadsheetApp.getUi().showModalDialog(html, 'Panel de control');
}

function obtenerDatosHomeDashboard() {
  const cache = CacheService.getScriptCache();
  const cacheKeyData = 'dashboard_full_data_cache';
  
  // Intentar obtener el objeto completo de la caché para carga instantánea
  const cachedData = cache.get(cacheKeyData);
  if (cachedData) {
    try {
      return JSON.parse(cachedData);
    } catch (e) { /* fallback a carga normal */ }
  }

  // Si no hay caché, realizamos el proceso pesado
  const lockKey = 'metricasBasicas_home_lastRun';
  const ttlSeconds = 60;
  const yaReciente = cache.get(lockKey);

  if (!yaReciente) {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(1500)) {
      try {
        const yaReciente2 = cache.get(lockKey);
        if (!yaReciente2) {
          const stateService = new StateService();
          stateService.runAutomaticTransitions(); // Esto actualiza las métricas de pacientes
          cache.put(lockKey, '1', ttlSeconds);
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
  const ahora = new Date();
  const ciclos = cicloRepo.findAll();
  const sesiones = sessionRepo.findAll();
  const modalidadesCfg = configRepo.findAll();
  const hoy = normalizarFecha_(new Date());
  const hoyMs = hoy.getTime();

  // Helper para normalizar comparaciones de texto (evita fallos por espacios o mayúsculas)
  const normalize = (str) => String(str || '').trim().toUpperCase();
  
  // Calcular altas del mes actual
  const primerDiaMesActual = new Date(ahora.getFullYear(), ahora.getMonth(), 1);
  const altasMesActual = pacientes.filter(p => {
    if (normalize(p.EstadoPaciente) !== ESTADOS_PACIENTE.ALTA) return false;
    const fCierre = (p.FechaCierre instanceof Date) ? p.FechaCierre : parseFechaES_(p.FechaCierre);
    return fCierre && !isNaN(fCierre.getTime()) && fCierre >= primerDiaMesActual;
  }).length;

  // Calcular espera media
  const pacientesEspera = pacientes.filter(p => normalize(p.EstadoPaciente) === ESTADOS_PACIENTE.ESPERA);
  let esperaMedia = 0;
  if (pacientesEspera.length > 0) {
    const sumaDias = pacientesEspera.reduce((acc, p) => {
      const fAlta = (p.FechaAlta instanceof Date) ? p.FechaAlta : parseFechaES_(p.FechaAlta);
      if (!fAlta || isNaN(fAlta.getTime())) return acc;
      const diff = Math.max(0, Math.floor((hoyMs - fAlta.getTime()) / (1000 * 60 * 60 * 24)));
      return acc + diff;
    }, 0);
    esperaMedia = Math.round(sumaDias / pacientesEspera.length);
  }

  const resumen = {
    totalPacientes: pacientes.length,
    activos: pacientes.filter(p => normalize(p.EstadoPaciente) === ESTADOS_PACIENTE.ACTIVO).length,
    activosInd: pacientes.filter(p => 
      normalize(p.EstadoPaciente) === ESTADOS_PACIENTE.ACTIVO && 
      normalize(p.ModalidadSolicitada) === MODALIDADES.INDIVIDUAL).length,
    activosGrp: pacientes.filter(p => 
      normalize(p.EstadoPaciente) === ESTADOS_PACIENTE.ACTIVO && 
      normalize(p.ModalidadSolicitada) !== MODALIDADES.INDIVIDUAL).length,
    espera: pacientesEspera.length,
    esperaMedia: esperaMedia,
    alta: pacientes.filter(p => normalize(p.EstadoPaciente) === ESTADOS_PACIENTE.ALTA).length,
    pendienteInicio: pacientes.filter(p => normalize(p.EstadoPaciente) === ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO).length,
    gruposEnCurso: ciclos.filter(c => normalize(c.EstadoCiclo) === ESTADOS_CICLO.EN_CURSO).length,
    altasMesActual: altasMesActual
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
      c.Modalidad === modalidad && (c.EstadoCiclo === 'PLANIFICADO' || c.EstadoCiclo === 'EN_CURSO')
    );

    if (ciclosVigentes.length === 0) {
      return { modalidad, capacidad: 0, ocupadas: 0, libres: 0, porcentaje: 0 };
    }

    // FIX: Cálculo agregado para modalidades de grupo
    const capacidad = ciclosVigentes.reduce((sum, c) => sum + Number(c.CapacidadMaxima || 0), 0);
    const ocupadas = ciclosVigentes.reduce((sum, c) => sum + Number(c.PlazasOcupadas || 0), 0);

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
    .filter(c => c.EstadoCiclo === ESTADOS_CICLO.PLANIFICADO || c.EstadoCiclo === ESTADOS_CICLO.EN_CURSO)
    .sort((a, b) => compararFechas_(a.FechaInicioCiclo, b.FechaInicioCiclo))
    .slice(0, 8)
    .map(c => ({ 
      CicloID: c.CicloID, 
      Modalidad: c.Modalidad, 
      NumeroCiclo: c.NumeroCiclo, 
      FechaInicioCiclo: formatearFecha_(c.FechaInicioCiclo),
      PlazasLibres: c.PlazasLibres,
      EstadoCiclo: c.EstadoCiclo
    }));

  const proximosPacientes = pacientes
    .filter(p => (p.EstadoPaciente === 'ACTIVO' || p.EstadoPaciente === 'ACTIVO_PENDIENTE_INICIO') && 
                 p.ProximaSesion instanceof Date && p.ProximaSesion >= ahora)
    .sort((a, b) => a.ProximaSesion.getTime() - b.ProximaSesion.getTime())
    .slice(0, 10)
    .map(p => ({ 
      PacienteID: p.PacienteID, 
      Nombre: p.Nombre, 
      EstadoPaciente: p.EstadoPaciente, 
      ProximaSesion: p.ProximaSesion instanceof Date ? 
        formatearFecha_(p.ProximaSesion) + ' ' + formatearHora_(p.ProximaSesion) : 
        formatearFecha_(p.ProximaSesion)
    }));

  const alertas = construirAlertasHome_(pacientes, ciclos, sesiones, hoy);

  const finalResult = {
    fechaHoy: formatearFecha_(hoy),
    fechaActualizacion: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    resumen,
    ocupacionPorModalidad,
    estadoCiclos,
    esperaPorModalidad,
    proximosCiclos,
    proximosPacientes,
    disponibilidadSemanal: new AvailabilityService().getFreeSlotsSummary(),
    resumenIncidenciasCalendario: obtenerResumenIncidenciasCalendario(sesiones),
    taskStatus: getBackgroundTaskStatus_(),
    calendarUrl: (typeof obtenerCalendarConsultaUrl_ === 'function') ? obtenerCalendarConsultaUrl_() : null,
    alertas
  };

  // Guardar en caché el resultado final (durante 10 minutos o hasta actualización manual)
  try { cache.put(cacheKeyData, JSON.stringify(finalResult), 600); } catch (e) {}

  return finalResult;
}

function eliminarCacheDashboard_() {
  const cache = CacheService.getScriptCache();
  cache.remove('dashboard_full_data_cache');
  cache.remove('metricasBasicas_home_lastRun'); // Permite que el panel vuelva a ejecutar transiciones inmediatamente
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


/**
 * Obtiene el estado de las tareas de segundo plano.
 */
function getBackgroundTaskStatus_() {
  const props = PropertiesService.getUserProperties();
  return {
    syncCalendar: {
      running: props.getProperty('TASK_SYNC_CALENDAR_RUNNING') === 'true',
      progress: parseInt(props.getProperty('TASK_SYNC_CALENDAR_PROGRESS') || '0'),
      lastResult: props.getProperty('TASK_SYNC_CALENDAR_RESULT') || ''
    },
    updateStates: {
      running: props.getProperty('TASK_UPDATE_STATES_RUNNING') === 'true',
      progress: parseInt(props.getProperty('TASK_UPDATE_STATES_PROGRESS') || '0'),
      lastResult: props.getProperty('TASK_UPDATE_STATES_RESULT') || ''
    }
  };
}

/***************
 * LANZADORES DESDE HOME
 ***************/
function homeNuevoPaciente() { nuevoPaciente(); }
function homeEditarPaciente() { editarPaciente(); }
function homeEliminarPacientePorError() { eliminarPacientePorError(); }
function homeEliminarCicloPorError() { eliminarCicloPorError(); }
function homeCrearCicloGrupo() { crearCicloGrupo(); }
function homeEditarCiclo() { editarCiclo(); } // Nueva función
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
function abrirHomeDashboardDesdePantalla() {  abrirHomeDashboard(); }
function homeAbrirReprogramarSesion() { abrirReprogramarSesion(); }
function homeAltaPaciente() { altaPaciente(); }
/**
 * Unificado: Limpia caché y reabre el dashboard para asegurar datos frescos.
 */
function refrescarPanel() {
  refrescarDashboard();
  abrirHomeDashboard();
}
function homeEstadisticasFichasPacientes() { estadisticasFichasPacientes(); }
function homeRecalcularEstadosAutomaticamente() { recalcularEstadosAutomaticamenteConModal(); }
function homeFichaClinicaPaciente() { fichaClinicaPaciente(); }
function homeVerIncidenciasCalendario() { verIncidenciasCalendario(); }
function homeObtenerResumenIncidenciasCalendario() { return obtenerResumenIncidenciasCalendario(); }

function volverAlPanelDesdeDiasBloqueados() {
  abrirHomeDashboard(); // Re-abre el panel de control
}