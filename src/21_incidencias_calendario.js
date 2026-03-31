/***********************
 * BLOQUE 21
 * INCIDENCIAS CALENDARIO
 ***********************/

function obtenerResumenIncidenciasCalendario(sesionesObjects) {
  const cantidad = obtenerCantidadSesionesAfectadasPorDiasBloqueados_(sesionesObjects);
  return { cantidad: cantidad };
}

function obtenerCantidadSesionesAfectadasPorDiasBloqueados_(sesData) {
  const data = sesData || new SessionRepository().findAll();

  if (!data || data.length === 0) return 0;

  

  const hoy = normalizarFecha_(new Date());
  const mapaDiasBloqueados = obtenerMapaDiasBloqueados_();

  let cantidad = 0;

  data.forEach(s => {
    const fechaSesion = s.FechaSesion;
    const estadoSesion = s.EstadoSesion || '';

    if (!fechaSesion || !(fechaSesion instanceof Date)) return;

    const fechaNormalizada = normalizarFecha_(fechaSesion);

    if (fechaNormalizada.getTime() < hoy.getTime()) return;

    if (estadoSesion !== ESTADOS_SESION.PENDIENTE && estadoSesion !== ESTADOS_SESION.REPROGRAMADA) return;
    const detalleBloqueo = obtenerDetalleBloqueoFechaConMapa_(fechaNormalizada, mapaDiasBloqueados);
    if (!detalleBloqueo.bloqueada) return;

    cantidad++;
  });
  return cantidad;
}

function verIncidenciasCalendario() {
  const html = HtmlService
    .createHtmlOutputFromFile('IncidenciasCalendarioForm')
    .setWidth(1100)
    .setHeight(720);

  SpreadsheetApp.getUi().showModalDialog(html, 'Incidencias de calendario');
}

function obtenerIncidenciasCalendarioFormulario() {
  return obtenerSesionesAfectadasPorDiasBloqueados_();
}

function obtenerSesionesAfectadasPorDiasBloqueados_() {
  const sessionRepo = new SessionRepository();
  const data = sessionRepo.findAll();
  if (data.length === 0) return [];

  const out = [];
  const hoy = normalizarFecha_(new Date());
  const mapaDiasBloqueados = obtenerMapaDiasBloqueados_();

  data.forEach(s => {
    const fechaSesion = s.FechaSesion;
    const estadoSesion = s.EstadoSesion || '';

    if (!fechaSesion || !(fechaSesion instanceof Date)) return;

    const fechaNormalizada = normalizarFecha_(fechaSesion);
    if (fechaNormalizada.getTime() < hoy.getTime()) return;
    if (estadoSesion !== ESTADOS_SESION.PENDIENTE && estadoSesion !== ESTADOS_SESION.REPROGRAMADA) return;

    const detalleBloqueo = obtenerDetalleBloqueoFechaConMapa_(fechaNormalizada, mapaDiasBloqueados);
    if (!detalleBloqueo.bloqueada) return;

    out.push({
      sesionId: s.SesionID,
      pacienteId: s.PacienteID,
      cicloId: s.CicloID,
      modalidad: s.Modalidad,
      nombrePaciente: s.NombrePaciente,
      numeroSesion: s.NumeroSesion,
      fechaSesion: formatearFecha_(fechaNormalizada),
      estadoSesion: estadoSesion,
      tipoBloqueo: detalleBloqueo.tipo || '',
      motivoBloqueo: detalleBloqueo.motivo || ''
    });
  });

  out.sort(function(a, b) {
    const fa = parseFechaES_(a.fechaSesion);
    const fb = parseFechaES_(b.fechaSesion);

    if (fa && fb && fa.getTime() !== fb.getTime()) {
      return fa.getTime() - fb.getTime();
    }

    if ((a.modalidad || '') !== (b.modalidad || '')) {
      return String(a.modalidad || '').localeCompare(String(b.modalidad || ''));
    }

    return String(a.nombrePaciente || '').localeCompare(String(b.nombrePaciente || ''));
  });

  return out;
}