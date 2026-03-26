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
          recalcularMetricasBasicas_();
          cache.put(cacheKey, '1', ttlSeconds);
        }
      } finally {
        lock.releaseLock();
      }
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPac = ss.getSheetByName(SHEET_PACIENTES);
  const sheetCic = ss.getSheetByName(SHEET_CICLOS);
  const sheetSes = ss.getSheetByName(SHEET_SESIONES);
  const sheetCfg = ss.getSheetByName(SHEET_CONFIG_MODALIDADES);

  if (!sheetPac || !sheetCic || !sheetSes || !sheetCfg) {
    throw new Error('Faltan hojas necesarias para construir el panel.');
  }

  const pacData = sheetPac.getDataRange().getValues();
  const cicData = sheetCic.getDataRange().getValues();
  const sesData = sheetSes.getDataRange().getValues();
  const cfgData = sheetCfg.getDataRange().getValues();

  const pacientes = mapHomePacientes_(pacData);
  const ciclos = mapHomeCiclos_(cicData);
  const sesiones = mapHomeSesiones_(sesData);
  const config = mapHomeConfig_(cfgData);
  const sesIdx = (sesData && sesData.length > 0) ? indexByHeader_(sesData[0]) : null;

  const hoy = normalizarFecha_(new Date());

  const resumen = {
    totalPacientes: pacientes.length,
    activos: pacientes.filter(p => p.estadoPaciente === ESTADOS_PACIENTE.ACTIVO).length,
    espera: pacientes.filter(p => p.estadoPaciente === ESTADOS_PACIENTE.ESPERA).length,
    alta: pacientes.filter(p => p.estadoPaciente === ESTADOS_PACIENTE.ALTA).length,
    pendienteInicio: pacientes.filter(p => p.estadoPaciente === ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO).length,
    gruposEnCurso: ciclos.filter(c => c.estadoCiclo === ESTADOS_CICLO.EN_CURSO).length,
    sesionesPendientes: sesiones.filter(s => s.estadoSesion === ESTADOS_SESION.PENDIENTE).length,
    sesionesErrorSync: sesiones.filter(s => s.calendarSyncStatus === 'ERROR').length
  };

  const ocupacionPorModalidad = Object.values(MODALIDADES).map(modalidad => {
    if (modalidad === MODALIDADES.INDIVIDUAL) {
      const capacidad = Number((config[modalidad] || {}).capacidadMaxima || 0);
      const ocupadas = pacientes.filter(p =>
        p.modalidad === modalidad &&
        p.estadoPaciente === ESTADOS_PACIENTE.ACTIVO
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
      c.modalidad === modalidad &&
      (c.estadoCiclo === ESTADOS_CICLO.PLANIFICADO || c.estadoCiclo === ESTADOS_CICLO.EN_CURSO)
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
      .sort((a, b) => compararFechas_(parseFechaHome_(a.fechaInicioCiclo), parseFechaHome_(b.fechaInicioCiclo)))[0];

    const capacidad = Number(cicloReferencia.capacidadMaxima || 0);
    const ocupadas = Number(cicloReferencia.plazasOcupadas || 0);

    return {
      modalidad,
      capacidad,
      ocupadas,
      libres: Math.max(0, capacidad - ocupadas),
      porcentaje: capacidad > 0 ? Math.round((ocupadas / capacidad) * 100) : 0
    };
  });

  const estadoCiclos = {
    planificados: ciclos.filter(c => c.estadoCiclo === ESTADOS_CICLO.PLANIFICADO).length,
    enCurso: ciclos.filter(c => c.estadoCiclo === ESTADOS_CICLO.EN_CURSO).length,
    cerrados: ciclos.filter(c => c.estadoCiclo === ESTADOS_CICLO.CERRADO).length,
    cancelados: ciclos.filter(c => c.estadoCiclo === ESTADOS_CICLO.CANCELADO).length
  };

  const esperaPorModalidad = Object.values(MODALIDADES).map(mod => ({
    modalidad: mod,
    total: pacientes.filter(p =>
      p.modalidad === mod &&
      p.estadoPaciente === ESTADOS_PACIENTE.ESPERA
    ).length
  }));

  const proximosCiclos = ciclos
    .filter(c => c.estadoCiclo === ESTADOS_CICLO.PLANIFICADO)
    .sort((a, b) => compararFechas_(parseFechaHome_(a.fechaInicioCiclo), parseFechaHome_(b.fechaInicioCiclo)))
    .slice(0, 8);

  const proximosPacientes = pacientes
    .filter(p =>
      p.estadoPaciente === ESTADOS_PACIENTE.ACTIVO ||
      p.estadoPaciente === ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO
    )
    .sort((a, b) => compararFechas_(parseFechaHome_(a.proximaSesion), parseFechaHome_(b.proximaSesion)))
    .slice(0, 10);

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
    resumenIncidenciasCalendario: obtenerResumenIncidenciasCalendario(sesData, sesIdx),
    calendarUrl: (typeof obtenerCalendarConsultaUrl_ === 'function') ? obtenerCalendarConsultaUrl_() : null,
    alertas
  };
}

function construirAlertasHome_(pacientes, ciclos, sesiones, hoy) {
  const alertas = [];

  const esperaSinCiclo = pacientes.filter(p =>
    p.estadoPaciente === ESTADOS_PACIENTE.ESPERA &&
    (!p.cicloObjetivoId && !p.cicloActivoId)
  ).length;

  if (esperaSinCiclo > 0) {
    alertas.push({
      tipo: 'warning',
      titulo: 'Pacientes en espera sin ciclo asignado',
      detalle: String(esperaSinCiclo)
    });
  }

  const erroresSync = sesiones.filter(s => s.calendarSyncStatus === 'ERROR').length;
  if (erroresSync > 0) {
    alertas.push({
      tipo: 'danger',
      titulo: 'Sesiones con error de sincronización',
      detalle: String(erroresSync)
    });
  }

  const ciclosLlenos = ciclos.filter(c =>
    (c.estadoCiclo === ESTADOS_CICLO.PLANIFICADO || c.estadoCiclo === ESTADOS_CICLO.EN_CURSO) &&
    Number(c.plazasLibres || 0) <= 0
  ).length;

  if (ciclosLlenos > 0) {
    alertas.push({
      tipo: 'warning',
      titulo: 'Ciclos sin plazas libres',
      detalle: String(ciclosLlenos)
    });
  }

  const grupos = [MODALIDADES.GRUPO_1, MODALIDADES.GRUPO_2, MODALIDADES.GRUPO_3];

  grupos.forEach(mod => {
    const tienePlanificado = ciclos.some(c =>
      c.modalidad === mod &&
      c.estadoCiclo === ESTADOS_CICLO.PLANIFICADO
    );

    const esperaMod = pacientes.filter(p =>
      p.modalidad === mod &&
      p.estadoPaciente === ESTADOS_PACIENTE.ESPERA
    ).length;

    if (esperaMod > 0 && !tienePlanificado) {
      alertas.push({
        tipo: 'warning',
        titulo: 'Grupo con espera sin próximo ciclo',
        detalle: mod + ' (' + esperaMod + ' en espera)'
      });
    }
  });

  const inicianPronto = pacientes.filter(p => {
    if (p.estadoPaciente !== ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO) return false;
    const fecha = parseFechaHome_(p.proximaSesion);
    if (!fecha) return false;
    const diff = Math.round((fecha.getTime() - hoy.getTime()) / (24 * 60 * 60 * 1000));
    return diff >= 0 && diff <= 7;
  }).length;

  if (inicianPronto > 0) {
    alertas.push({
      tipo: 'info',
      titulo: 'Pacientes que empiezan ciclo en 7 días',
      detalle: String(inicianPronto)
    });
  }

  if (alertas.length === 0) {
    alertas.push({
      tipo: 'ok',
      titulo: 'Sin alertas relevantes',
      detalle: 'Todo en orden'
    });
  }

  return alertas;
}

function mapHomePacientes_(data) {
  if (!data || data.length < 2) return [];

  const idx = indexByHeader_(data[0]);

  return data.slice(1).map(row => ({
    pacienteId: row[idx.PacienteID] || '',
    nombre: row[idx.Nombre] || '',
    modalidad: row[idx.ModalidadSolicitada] || '',
    estadoPaciente: row[idx.EstadoPaciente] || '',
    proximaSesion: formatearFecha_(row[idx.ProximaSesion]),
    cicloObjetivoId: row[idx.CicloObjetivoID] || '',
    cicloActivoId: row[idx.CicloActivoID] || '',
    sesionesCompletadas: Number(row[idx.SesionesCompletadas] || 0),
    sesionesPendientes: Number(row[idx.SesionesPendientes] || 0)
  }));
}

function mapHomeCiclos_(data) {
  if (!data || data.length < 2) return [];

  const idx = indexByHeader_(data[0]);

  return data.slice(1).map(row => ({
    cicloId: row[idx.CicloID] || '',
    modalidad: row[idx.Modalidad] || '',
    numeroCiclo: Number(row[idx.NumeroCiclo] || 0),
    estadoCiclo: row[idx.EstadoCiclo] || '',
    fechaInicioCiclo: formatearFecha_(row[idx.FechaInicioCiclo]),
    fechaFinCiclo: formatearFecha_(row[idx.FechaFinCiclo]),
    capacidadMaxima: Number(row[idx.CapacidadMaxima] || 0),
    plazasOcupadas: Number(row[idx.PlazasOcupadas] || 0),
    plazasLibres: Number(row[idx.PlazasLibres] || 0)
  }));
}

function mapHomeSesiones_(data) {
  if (!data || data.length < 2) return [];

  const idx = indexByHeader_(data[0]);

  return data.slice(1).map(row => ({
    sesionId: row[idx.SesionID] || '',
    estadoSesion: row[idx.EstadoSesion] || '',
    calendarSyncStatus: row[idx.CalendarSyncStatus] || ''
  }));
}

function mapHomeConfig_(data) {
  if (!data || data.length < 2) return {};

  const idx = indexByHeader_(data[0]);
  const out = {};

  data.slice(1).forEach(row => {
    const modalidad = row[idx.Modalidad];
    if (!modalidad) return;

    out[modalidad] = {
      capacidadMaxima: Number(row[idx.CapacidadMaxima] || 0)
    };
  });

  return out;
}

function parseFechaHome_(texto) {
  if (!texto) return null;
  const partes = String(texto).split('/');
  if (partes.length !== 3) return null;

  const day = Number(partes[0]);
  const month = Number(partes[1]) - 1;
  const year = Number(partes[2]);

  const fecha = new Date(year, month, day);
  return isNaN(fecha.getTime()) ? null : fecha;
}

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
function homeEstadisticasFichasPacientes() { estadisticasFichasPacientes(); }

function volverAlPanelDesdeDiasBloqueados() {
  abrirHomeDashboard(); // Re-abre el panel de control
}
function homeEstadisticasFichasPacientes() { estadisticasFichasPacientes(); }