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
  // 1. Obtención de datos mediante Repositorios
  const pacientes = patientRepo.findAll();
  const ciclos = cycleRepo.findAll();
  const sesiones = sessionRepo.findAll();
  
  const hoy = normalizarFecha_(new Date());

  // 2. Cálculo de métricas generales (Single-pass sobre Entidades)
  const resumen = {
    totalPacientes: pacientes.length,
    activos: pacientes.filter(p => p.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO).length,
    espera: pacientes.filter(p => p.EstadoPaciente === ESTADOS_PACIENTE.ESPERA).length,
    alta: pacientes.filter(p => p.EstadoPaciente === ESTADOS_PACIENTE.ALTA).length,
    pendienteInicio: pacientes.filter(p => p.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO_PENDIENTE_INICIO).length,
    gruposEnCurso: ciclos.filter(c => c.EstadoCiclo === ESTADOS_CICLO.EN_CURSO).length,
    sesionesPendientes: sesiones.filter(s => s.EstadoSesion === ESTADOS_SESION.PENDIENTE).length,
    sesionesErrorSync: sesiones.filter(s => s.CalendarSyncStatus === 'ERROR').length
  };

  // 3. Ocupación por modalidad
  const ocupacionPorModalidad = Object.values(MODALIDADES).map(modalidad => {
    if (modalidad === MODALIDADES.INDIVIDUAL) {
      const config = obtenerConfigModalidad_(modalidad);
      const capacidad = Number(config.CapacidadMaxima || 0);
      const ocupadas = pacientes.filter(p =>
        p.ModalidadSolicitada === modalidad &&
        p.EstadoPaciente === ESTADOS_PACIENTE.ACTIVO
      ).length;

      return {
        modalidad,
        capacidad,
        ocupadas,
        libres: Math.max(0, capacidad - ocupadas)
      };
    }

    // Para grupos, buscamos el ciclo activo o el próximo planificado
    const cicloRef = ciclos
      .filter(c => c.Modalidad === modalidad && c.EstadoCiclo !== ESTADOS_CICLO.CERRADO)
      .sort((a, b) => a.FechaInicioCiclo - b.FechaInicioCiclo)[0];

    return {
      modalidad,
      capacidad: cicloRef ? cicloRef.CapacidadMaxima : 0,
      ocupadas: cicloRef ? cicloRef.PlazasOcupadas : 0,
      libres: cicloRef ? cicloRef.PlazasLibres : 0
    };
  });

  // 4. Listados próximos
  const proximosCiclos = ciclos
    .filter(c => c.EstadoCiclo === ESTADOS_CICLO.PLANIFICADO)
    .sort((a, b) => a.FechaInicioCiclo - b.FechaInicioCiclo)
    .slice(0, 8);

  const proximosPacientes = pacientes
    .filter(p => p.EstadoPaciente !== ESTADOS_PACIENTE.ALTA && p.ProximaSesion)
    .sort((a, b) => a.ProximaSesion - b.ProximaSesion)
    .slice(0, 10);

  const alertas = construirAlertasHome_(pacientes, ciclos, sesiones, hoy);

  return {
    fechaHoy: formatearFecha_(hoy),
    fechaActualizacion: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    resumen,
    ocupacionPorModalidad,
    proximosCiclos,
    proximosPacientes,
    calendarUrl: CalendarService.getCalendar().getId(),
    alertas
  };
}

function construirAlertasHome_(pacientes, ciclos, sesiones, hoy) {
  const alertas = [];

  const esperaSinCiclo = pacientes.filter(p =>
    p.EstadoPaciente === ESTADOS_PACIENTE.ESPERA &&
    (!p.CicloObjetivoID && !p.CicloActivoID)
  ).length;

  if (esperaSinCiclo > 0) {
    alertas.push({
      tipo: 'warning',
      titulo: 'Pacientes en espera sin ciclo asignado',
      detalle: `${esperaSinCiclo} pacientes`
    });
  }

  const erroresSync = sesiones.filter(s => s.CalendarSyncStatus === 'ERROR').length;
  if (erroresSync > 0) {
    alertas.push({
      tipo: 'danger',
      titulo: 'Sesiones con error de sincronización',
      detalle: `${erroresSync} sesiones`
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