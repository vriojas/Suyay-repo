/**
 * Suyay Voluntarios — backend (Google Apps Script + Google Sheets)
 * Se instala desde la cuenta comunicacion.suyay@gmail.com.
 * Todos los correos salen desde esa cuenta.
 */

const TZ = 'America/Lima';
const SHEET_ID = '1xFxvAM80l0NG-iB_FGxtTGQPAZoUkaPvZbTu54ippN8';
function libro_() { return SpreadsheetApp.openById(SHEET_ID); }

const HOJAS = {
  Voluntarios: ['id', 'nombre', 'dni', 'celular', 'correo', 'universidad', 'nacimiento', 'ingreso_mes', 'ingreso_anio', 'fecha_registro', 'token'],
  Actividades: ['id', 'titulo', 'programa', 'fecha', 'hora_inicio', 'hora_fin', 'lugar', 'lat', 'lng', 'radio_m', 'cupos', 'cierre_inscripcion', 'resumen', 'whatsapp', 'horas', 'estado'],
  Inscripciones: ['fecha_hora', 'actividad_id', 'voluntario_id', 'estado'],
  Asistencias: ['fecha_hora', 'actividad_id', 'voluntario_id', 'lat', 'lng', 'distancia_m', 'horas'],
  Fotos: ['fecha_hora', 'actividad_id', 'voluntario_id', 'url'],
  Feedback: ['fecha_hora', 'actividad_id', 'voluntario_id', 'disfrute_1a5', 'organizacion_1a5', 'volverias_1a5', 'comentario'],
  Certificados: ['fecha_hora', 'voluntario_id', 'tipo', 'horas', 'jornadas', 'codigo', 'url_pdf', 'enviado'],
  Constancias: ['dni (escribe aquí para emitir un certificado actualizado)', 'estado'],
  Config: ['clave', 'valor', 'nota']
};

const CONFIG_INICIAL = [
  ['PLANTILLA_ID', '', 'ID de la plantilla del certificado en Google Slides'],
  ['RUC', '20613847066', 'RUC de la Asociación Suyay Global Impact'],
  ['REGISTRO_MIMP', '', 'N.° de inscripción en el Registro de Voluntariado del MIMP (dejar vacío si aún no lo tienen)'],
  ['HITO_HORAS', 20, 'Cada cuántas horas se entrega certificado e insignia'],
  ['MARGEN_ANTES_MIN', 30, 'Minutos antes del inicio en que ya se puede marcar asistencia'],
  ['MARGEN_DESPUES_MIN', 30, 'Minutos después del fin en que aún se puede marcar asistencia'],
  ['RADIO_DEFECTO_M', 200, 'Distancia al lugar (metros) si la actividad no indica otra'],
  ['RADIO_MAX_M', 300, 'Tope: ninguna actividad acepta asistencias a más de esta distancia'],
  ['FIRMANTE2_NOMBRE', 'Paolo Bazán', 'Segunda firma del certificado'],
  ['FIRMANTE2_CARGO', 'Coordinador de Voluntariado', ''],
  ['FIRMANTE2_ARCHIVO', 'firma-paolo.png', 'Nombre del archivo de la firma en la carpeta Recursos'],
  ['APP_URL', '', 'Link de la app'],
  ['REMITENTE', 'Suyay Global Impact', 'Nombre que ven los voluntarios en los correos'],
  ['CARPETA_RAIZ_ID', '', 'Se llena solo con setup()'],
  ['CARPETA_RECURSOS_ID', '', 'Se llena solo'],
  ['CARPETA_INSIGNIAS_ID', '', 'Se llena solo'],
  ['CARPETA_CERTIFICADOS_ID', '', 'Se llena solo'],
  ['CARPETA_FOTOS_ID', '', 'Se llena solo'],
  ['CARPETA_TEMP_ID', '', 'Se llena solo']
];

/* =========================================================
 * INSTALACIÓN — correr una sola vez desde el editor
 * ========================================================= */
function setup() {
  const ss = libro_();
  Object.keys(HOJAS).forEach(nombre => {
    let sh = ss.getSheetByName(nombre);
    if (!sh) sh = ss.insertSheet(nombre);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HOJAS[nombre]);
      sh.getRange(1, 1, 1, HOJAS[nombre].length).setFontWeight('bold').setBackground('#112EA7').setFontColor('#F4F4EE');
      sh.setFrozenRows(1);
    }
  });
  const cfg = ss.getSheetByName('Config');
  if (cfg.getLastRow() === 1) CONFIG_INICIAL.forEach(r => cfg.appendRow(r));
  const hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && ss.getSheets().length > 1) ss.deleteSheet(hoja1);

  // Formatos de la hoja Actividades para que Lucía la llene fácil
  const act = ss.getSheetByName('Actividades');
  act.getRange('D2:D').setNumberFormat('dd/mm/yyyy');
  act.getRange('L2:L').setNumberFormat('dd/mm/yyyy');
  act.getRange('E2:F').setNumberFormat('@'); // horas como texto "09:00"
  act.getRange('P2:P').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['activa', 'cancelada'], true).build());
  ss.getSheetByName('Voluntarios').getRange('C2:D').setNumberFormat('@');

  // Carpetas en Drive
  let raizId = getConfig_('CARPETA_RAIZ_ID');
  let raiz = raizId ? DriveApp.getFolderById(raizId) : DriveApp.createFolder('Suyay App Voluntarios');
  setConfig_('CARPETA_RAIZ_ID', raiz.getId());
  [['CARPETA_RECURSOS_ID', 'Recursos (logo y firmas)'], ['CARPETA_INSIGNIAS_ID', 'Insignias'],
   ['CARPETA_CERTIFICADOS_ID', 'Certificados'], ['CARPETA_FOTOS_ID', 'Fotos de actividades'], ['CARPETA_TEMP_ID', 'temp']]
    .forEach(([k, n]) => { if (!getConfig_(k)) setConfig_(k, raiz.createFolder(n).getId()); });

  // Automatizaciones diarias
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('enviarCumpleanos').timeBased().atHour(8).everyDays(1).inTimezone(TZ).create();
  ScriptApp.newTrigger('enviarPedidoFeedback').timeBased().atHour(20).everyDays(1).inTimezone(TZ).create();
  ScriptApp.newTrigger('alEditar').forSpreadsheet(SHEET_ID).onEdit().create();
  return 'Listo: hojas, carpetas y automatizaciones creadas.';
}

/** Se ejecuta solo cuando alguien edita la hoja. */
function alEditar(e) {
  const sh = e.range.getSheet();
  if (sh.getName() === 'Actividades') { asignarIds_(); return; }
  if (sh.getName() !== 'Constancias' || e.range.getColumn() !== 1 || e.range.getRow() < 2) return;
  const dni = String(e.range.getValue()).trim();
  if (!dni) return;
  const est = sh.getRange(e.range.getRow(), 2);
  try { est.setValue('Generando…'); est.setValue(emitirConstancia(dni)); }
  catch (err) { est.setValue('Error: ' + err.message); }
}

/* =========================================================
 * API — la app llama aquí con POST {action, ...}
 * ========================================================= */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.ruta = (e && e.parameter && e.parameter.r) || '';
  t.appUrl = ScriptApp.getService().getUrl();
  return t.evaluate().setTitle('Suyay Voluntarios')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** La app llama aquí con google.script.run */
function apiCall(reqJson) {
  const req = JSON.parse(reqJson);
  try {
    const fn = API[req.action];
    if (!fn) return JSON.stringify({ ok: false, error: 'Acción desconocida' });
    return JSON.stringify(Object.assign({ ok: true }, fn(req)));
  } catch (err) {
    return JSON.stringify({ ok: false, error: err.message || String(err) });
  }
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Solicitud inválida' }); }
  try {
    const fn = API[req.action];
    if (!fn) return json_({ ok: false, error: 'Acción desconocida' });
    return json_(Object.assign({ ok: true }, fn(req)));
  } catch (err) {
    return json_({ ok: false, error: err.message || String(err) });
  }
}

const API = {
  registrar(r) {
    const d = r.datos || {};
    const dni = String(d.dni || '').trim();
    if (!/^\d{8}$/.test(dni)) throw new Error('El DNI debe tener 8 dígitos.');
    ['nombre', 'celular', 'correo', 'nacimiento', 'ingreso_mes', 'ingreso_anio'].forEach(k => {
      if (!String(d[k] || '').trim()) throw new Error('Falta completar: ' + k.replace('_', ' '));
    });
    if (!/^\S+@\S+\.\S+$/.test(d.correo)) throw new Error('Revisa tu correo.');
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      if (filas_('Voluntarios').some(v => String(v.dni) === dni)) throw new Error('Ese DNI ya está registrado. Ingresa con tu DNI y fecha de nacimiento.');
      const id = 'V' + Utilities.formatString('%04d', filas_('Voluntarios').length + 1);
      const token = Utilities.getUuid();
      hoja_('Voluntarios').appendRow([id, d.nombre.trim(), "'" + dni, "'" + String(d.celular).trim(), d.correo.trim().toLowerCase(),
        (d.universidad || '').trim(), d.nacimiento, Number(d.ingreso_mes), Number(d.ingreso_anio), new Date(), token]);
      return { token };
    } finally { lock.releaseLock(); }
  },

  login(r) {
    const dni = String(r.dni || '').trim();
    const nac = String(r.nacimiento || '').trim();
    const v = filas_('Voluntarios').find(x => String(x.dni) === dni && fechaISO_(x.nacimiento) === nac);
    if (!v) throw new Error('No encontramos ese DNI con esa fecha de nacimiento.');
    return { token: v.token };
  },

  perfil(r) {
    const v = voluntario_(r.token);
    const acts = mapa_(filas_('Actividades'), 'id');
    const asis = filas_('Asistencias').filter(a => a.voluntario_id === v.id);
    const horas = asis.reduce((s, a) => s + Number(a.horas || 0), 0);
    const fb = new Set(filas_('Feedback').filter(f => f.voluntario_id === v.id).map(f => f.actividad_id));
    const hace14 = Date.now() - 14 * 864e5;
    return {
      voluntario: {
        id: v.id, nombre: v.nombre, correo: v.correo, celular: String(v.celular), universidad: v.universidad,
        nacimiento: fechaISO_(v.nacimiento), ingreso_mes: Number(v.ingreso_mes), ingreso_anio: Number(v.ingreso_anio)
      },
      horas,
      hito: Number(getConfig_('HITO_HORAS')) || 20,
      asistencias: asis.map(a => ({ actividad_id: a.actividad_id, titulo: (acts[a.actividad_id] || {}).titulo || a.actividad_id, fecha: fechaISO_(a.fecha_hora), horas: Number(a.horas) })),
      inscripciones: filas_('Inscripciones').filter(i => i.voluntario_id === v.id && i.estado === 'inscrito').map(i => i.actividad_id),
      certificados: filas_('Certificados').filter(c => c.voluntario_id === v.id).map(c => ({ tipo: c.tipo, horas: Number(c.horas), jornadas: Number(c.jornadas), url: c.url_pdf, fecha: fechaISO_(c.fecha_hora), codigo: c.codigo })),
      feedbackPendiente: asis.filter(a => !fb.has(a.actividad_id) && new Date(a.fecha_hora).getTime() > hace14)
        .map(a => ({ actividad_id: a.actividad_id, titulo: (acts[a.actividad_id] || {}).titulo || '' }))
    };
  },

  actividades() {
    asignarIds_();
    const insc = filas_('Inscripciones').filter(i => i.estado === 'inscrito');
    const hoy = fechaISO_(new Date());
    const lista = filas_('Actividades')
      .filter(a => a.id && a.titulo && a.estado !== 'cancelada' && fechaISO_(a.fecha) >= restarDias_(hoy, 31))
      .map(a => {
        const inscritos = insc.filter(i => i.actividad_id === a.id).length;
        return {
          id: a.id, titulo: a.titulo, programa: a.programa, fecha: fechaISO_(a.fecha),
          hora_inicio: hora_(a.hora_inicio), hora_fin: hora_(a.hora_fin), lugar: a.lugar,
          cupos: Number(a.cupos) || 0, inscritos, cierre: a.cierre_inscripcion ? fechaISO_(a.cierre_inscripcion) : '',
          resumen: a.resumen, whatsapp: a.whatsapp, horas: Number(a.horas) || 0
        };
      })
      .sort((x, y) => (x.fecha + x.hora_inicio).localeCompare(y.fecha + y.hora_inicio));
    return { actividades: lista, hoy };
  },

  inscribir(r) {
    const v = voluntario_(r.token);
    const a = actividad_(r.actividad_id);
    const hoy = fechaISO_(new Date());
    if (a.cierre_inscripcion && hoy > fechaISO_(a.cierre_inscripcion)) throw new Error('Las inscripciones para esta actividad ya cerraron.');
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const insc = filas_('Inscripciones').filter(i => i.actividad_id === a.id && i.estado === 'inscrito');
      if (insc.some(i => i.voluntario_id === v.id)) return { whatsapp: a.whatsapp };
      if (Number(a.cupos) > 0 && insc.length >= Number(a.cupos)) throw new Error('¡Uy! Ya no quedan cupos para esta actividad.');
      hoja_('Inscripciones').appendRow([new Date(), a.id, v.id, 'inscrito']);
      return { whatsapp: a.whatsapp };
    } finally { lock.releaseLock(); }
  },

  desinscribir(r) {
    const v = voluntario_(r.token);
    const sh = hoja_('Inscripciones');
    const data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === r.actividad_id && data[i][2] === v.id && data[i][3] === 'inscrito') sh.getRange(i + 1, 4).setValue('cancelado');
    }
    return {};
  },

  asistencia(r) {
    const v = voluntario_(r.token);
    const a = actividad_(r.actividad_id);
    const ahora = new Date();
    if (fechaISO_(a.fecha) !== fechaISO_(ahora)) throw new Error('La asistencia solo se puede marcar el día de la actividad.');
    const ini = fechaHora_(a.fecha, a.hora_inicio), fin = fechaHora_(a.fecha, a.hora_fin);
    const antes = Number(getConfig_('MARGEN_ANTES_MIN')) || 30, despues = Number(getConfig_('MARGEN_DESPUES_MIN')) || 30;
    if (ahora < new Date(ini.getTime() - antes * 6e4)) throw new Error('Todavía es temprano. Podrás marcar desde ' + antes + ' minutos antes del inicio.');
    if (ahora > new Date(fin.getTime() + despues * 6e4)) throw new Error('El horario para marcar asistencia ya terminó.');
    const coords = coordenadas_(a);
    if (!coords) throw new Error('Esta actividad aún no tiene ubicación configurada. Avísale a tu coordinadora.');
    if (typeof r.lat !== 'number' || typeof r.lng !== 'number') throw new Error('No pudimos leer tu ubicación.');
    const dist = Math.round(distancia_(coords.lat, coords.lng, r.lat, r.lng));
    const radio = Math.min(Number(getConfig_('RADIO_MAX_M')) || 300, Number(a.radio_m) || Number(getConfig_('RADIO_DEFECTO_M')) || 200);
    if (dist > radio) throw new Error('Parece que no estás en el lugar de la actividad (estás a ' + dist + ' m). Acércate e intenta otra vez.');
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    let nuevos = [];
    try {
      if (filas_('Asistencias').some(x => x.actividad_id === a.id && x.voluntario_id === v.id)) return { yaMarcada: true };
      hoja_('Asistencias').appendRow([ahora, a.id, v.id, r.lat, r.lng, dist, Number(a.horas) || 0]);
    } finally { lock.releaseLock(); }
    try { nuevos = revisarHitos_(v); } catch (err) { console.error(err); }
    return { horasSumadas: Number(a.horas) || 0, nuevosCertificados: nuevos };
  },

  foto(r) {
    const v = voluntario_(r.token);
    const a = actividad_(r.actividad_id);
    const m = /^data:(image\/\w+);base64,(.+)$/.exec(r.imagen || '');
    if (!m) throw new Error('Imagen inválida.');
    const raiz = DriveApp.getFolderById(getConfig_('CARPETA_FOTOS_ID'));
    const nombreCarpeta = fechaISO_(a.fecha) + ' · ' + a.titulo;
    const it = raiz.getFoldersByName(nombreCarpeta);
    const carpeta = it.hasNext() ? it.next() : raiz.createFolder(nombreCarpeta);
    const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], v.nombre + ' ' + Date.now() + '.jpg');
    const f = carpeta.createFile(blob);
    hoja_('Fotos').appendRow([new Date(), a.id, v.id, f.getUrl()]);
    return {};
  },

  feedback(r) {
    const v = voluntario_(r.token);
    const ok = n => Number.isInteger(n) && n >= 1 && n <= 5;
    if (![r.p1, r.p2, r.p3].every(ok)) throw new Error('Responde las tres preguntas.');
    if (!filas_('Asistencias').some(x => x.actividad_id === r.actividad_id && x.voluntario_id === v.id)) throw new Error('Solo puedes opinar sobre actividades a las que asististe.');
    if (filas_('Feedback').some(f => f.actividad_id === r.actividad_id && f.voluntario_id === v.id)) return {};
    hoja_('Feedback').appendRow([new Date(), r.actividad_id, v.id, r.p1, r.p2, r.p3, String(r.comentario || '').slice(0, 1000)]);
    return {};
  }
};

/* =========================================================
 * CERTIFICADOS
 * ========================================================= */
function revisarHitos_(v) {
  const hito = Number(getConfig_('HITO_HORAS')) || 20;
  const horas = filas_('Asistencias').filter(a => a.voluntario_id === v.id).reduce((s, a) => s + Number(a.horas || 0), 0);
  const hechos = new Set(filas_('Certificados').filter(c => c.voluntario_id === v.id && c.tipo === 'hito').map(c => Number(c.horas)));
  const nuevos = [];
  for (let h = hito; h <= horas; h += hito) {
    if (!hechos.has(h)) { generarCertificado_(v, h, 'hito'); nuevos.push(h); }
  }
  return nuevos;
}

/** Revisa a todos (útil si se cargaron asistencias a mano en la hoja). */
function revisarTodosLosHitos() {
  let total = 0;
  filas_('Voluntarios').forEach(v => { total += revisarHitos_(v).length; });
  return total + ' certificado(s) nuevo(s) enviados.';
}

/**
 * Genera el certificado (hoja 1) + detalle de jornadas (hoja 2 en adelante).
 * tipo: 'hito' (automático cada 20 h), 'constancia' (a pedido) o 'prueba'.
 */
function generarCertificado_(v, horas, tipo) {
  const plantillaId = getConfig_('PLANTILLA_ID');
  if (!plantillaId) throw new Error('Falta PLANTILLA_ID en Config.');

  // Detalle de jornadas del voluntario
  let jornadas;
  if (tipo === 'prueba') {
    jornadas = [['15/03/2025', 'Inducción al voluntariado', 'Capacitación', 3], ['22/03/2025', 'Biblioteca móvil Tiwinza', 'Lectura', 4],
      ['05/04/2025', 'Taller de teatro', 'Belaúnde Terry', 4], ['10/05/2025', 'Mediación lectora', 'Capacitación', 3]];
  } else {
    const acts = mapa_(filas_('Actividades'), 'id');
    jornadas = filas_('Asistencias').filter(a => a.voluntario_id === v.id)
      .map(a => { const act = acts[a.actividad_id] || {}; return [Utilities.formatDate(new Date(a.fecha_hora), TZ, 'dd/MM/yyyy'), act.titulo || a.actividad_id, act.programa || '', Number(a.horas) || 0]; });
  }
  const capacitaciones = jornadas.filter(j => /capacit|inducci|taller para volunt/i.test(j[2] + ' ' + j[1])).map(j => j[1] + ' (' + j[0] + ')');

  const meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const hoy = new Date();
  const fecha = Utilities.formatDate(hoy, TZ, 'd') + ' de ' + meses[Number(Utilities.formatDate(hoy, TZ, 'M')) - 1] + ' de ' + Utilities.formatDate(hoy, TZ, 'yyyy');
  const codigo = 'SUY-' + Utilities.formatString('%04d', horas) + '-' + Utilities.getUuid().slice(0, 4).toUpperCase();
  const registro = getConfig_('REGISTRO_MIMP');
  const reemplazos = {
    '{{NOMBRE}}': v.nombre, '{{DNI}}': String(v.dni), '{{HORAS}}': String(horas), '{{JORNADAS}}': String(jornadas.length),
    '{{DESDE}}': meses[Number(v.ingreso_mes) - 1] + ' de ' + v.ingreso_anio,
    '{{FECHA}}': fecha, '{{FECHA_CORTE}}': fecha, '{{CODIGO}}': codigo,
    '{{RUC}}': getConfig_('RUC') || '—',
    '{{INSCRIPCION}}': registro ? ', inscrita en el Registro de Voluntariado del MIMP con N.° ' + registro : '',
    '{{CAPACITACIONES}}': capacitaciones.length ? capacitaciones.join('; ') : 'ninguna registrada a la fecha',
    '{{FIRMANTE2_NOMBRE}}': getConfig_('FIRMANTE2_NOMBRE'), '{{FIRMANTE2_CARGO}}': getConfig_('FIRMANTE2_CARGO')
  };
  const imagenes = {
    '{{LOGO}}': archivoEn_('CARPETA_RECURSOS_ID', 'logo.png'),
    '{{FIRMA1}}': archivoEn_('CARPETA_RECURSOS_ID', 'firma-valeria.png'),
    '{{FIRMA2}}': archivoEn_('CARPETA_RECURSOS_ID', getConfig_('FIRMANTE2_ARCHIVO')),
    '{{INSIGNIA}}': insigniaPara_(horas)
  };

  const temp = DriveApp.getFolderById(getConfig_('CARPETA_TEMP_ID'));
  const copia = DriveApp.getFileById(plantillaId).makeCopy('cert-' + v.id + '-' + horas, temp);
  const pres = SlidesApp.openById(copia.getId());

  // Fondos diseñados en Canva (si están en Recursos, reemplazan los de la plantilla)
  ponerFondo_(pres.getSlides()[0], archivoEn_('CARPETA_RECURSOS_ID', 'fondo-certificado.png'), pres);
  ponerFondo_(pres.getSlides()[1], archivoEn_('CARPETA_RECURSOS_ID', 'fondo-anexo.png'), pres);

  // Hoja de detalle: 14 jornadas por página
  const POR_PAG = 14;
  const anexo = pres.getSlides()[1];
  const paginas = Math.max(1, Math.ceil(jornadas.length / POR_PAG));
  const anexos = [anexo];
  for (let i = 1; i < paginas; i++) anexos.push(anexos[anexos.length - 1].duplicate());
  anexos.forEach((sl, i) => llenarTabla_(sl, jornadas.slice(i * POR_PAG, (i + 1) * POR_PAG)));

  pres.getSlides().forEach(sl => reemplazarImagenes_(sl, imagenes));
  Object.keys(reemplazos).forEach(k => pres.replaceAllText(k, reemplazos[k]));
  pres.saveAndClose();
  if (tablasPendientes_.length) {
    Slides.Presentations.batchUpdate({ requests: tablasPendientes_.flatMap(tb => [0.16, 0.48, 0.26, 0.10].map((f, j) => ({ updateTableColumnProperties: { objectId: tb.id, columnIndices: [j], tableColumnProperties: { columnWidth: { magnitude: tb.W * f, unit: 'PT' } }, fields: 'columnWidth' } }))) }, copia.getId());
    tablasPendientes_ = [];
  }

  const pdf = copia.getAs('application/pdf').setName('Certificado Suyay ' + horas + 'h - ' + v.nombre + '.pdf');
  copia.setTrashed(true);
  if (tipo === 'prueba') return pdf;

  const archivo = DriveApp.getFolderById(getConfig_('CARPETA_CERTIFICADOS_ID')).createFile(pdf);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  let enviado = 'sí';
  try { enviarCorreoCertificado_(v, horas, pdf, tipo); } catch (err) { enviado = 'error: ' + err.message; }
  hoja_('Certificados').appendRow([new Date(), v.id, tipo, horas, jornadas.length, codigo, archivo.getUrl(), enviado]);
  return archivo;
}

function ponerFondo_(slide, archivo, pres) {
  if (!archivo) return;
  const ancho = pres.getPageWidth();
  slide.getImages().filter(im => im.getWidth() > ancho * 0.9).forEach(im => im.replace(archivo.getBlob()));
}

function reemplazarImagenes_(slide, imagenes) {
  slide.getShapes().forEach(sh => {
    const t = sh.getText().asString().trim();
    if (!(t in imagenes)) return;
    const f = imagenes[t];
    if (f) {
      const img = slide.insertImage(f.getBlob());
      const L = sh.getLeft(), T = sh.getTop(), W = sh.getWidth(), H = sh.getHeight();
      const esc = Math.min(W / img.getWidth(), H / img.getHeight());
      const w = img.getWidth() * esc, h = img.getHeight() * esc;
      img.setWidth(w).setHeight(h).setLeft(t === '{{LOGO}}' ? L : L + (W - w) / 2).setTop(t === '{{LOGO}}' ? T : T + (H - h) / 2);
    }
    sh.remove();
  });
}

var tablasPendientes_ = [];
function llenarTabla_(slide, filas) {
  const marcador = slide.getShapes().find(sh => sh.getText().asString().trim() === '{{TABLA}}');
  if (!marcador) return;
  const L = marcador.getLeft(), T = marcador.getTop(), W = marcador.getWidth();
  marcador.remove();
  if (!filas.length) {
    const tb = slide.insertTextBox('Aún no hay jornadas registradas.', L, T, W, 30);
    tb.getText().getTextStyle().setFontFamily('Quicksand').setFontSize(11);
    return;
  }
  const datos = [['Fecha', 'Actividad', 'Programa', 'Horas']].concat(filas.map(f => f.map(String)));
  const tabla = slide.insertTable(datos.length, 4, L, T, W, 22 * datos.length);
  const anchos = [0.16, 0.48, 0.26, 0.10];
  tablasPendientes_.push({ id: tabla.getObjectId(), W: W });
  datos.forEach((fila, i) => fila.forEach((txt, j) => {
    const c = tabla.getCell(i, j);
    c.getText().setText(txt);
    const st = c.getText().getTextStyle().setFontFamily('Quicksand').setFontSize(i === 0 ? 10.5 : 10);
    if (i === 0) { c.getFill().setSolidFill('#112EA7'); st.setForegroundColor('#F4F4EE').setBold(true); }
    else { c.getFill().setSolidFill(i % 2 ? '#FFFFFF' : '#EEF0F8'); st.setForegroundColor('#2A2A3A'); }
  }));
}

function enviarCorreoCertificado_(v, horas, pdf, tipo) {
  const nombre = String(v.nombre).split(' ')[0];
  const app = getConfig_('APP_URL');
  if (tipo === 'constancia') {
    const html = plantillaCorreo_('Tu certificado de voluntariado 📄',
      '<p>Hola ' + nombre + ',</p><p>Te enviamos tu certificado de voluntariado actualizado, con <b>' + horas + ' horas</b> acreditadas y el detalle de tus jornadas.</p>' +
      '<p>Gracias por ser parte de la familia Suyay 💛</p>');
    MailApp.sendEmail({ to: v.correo, subject: 'Tu certificado de voluntariado · Suyay', htmlBody: html, attachments: [pdf], name: getConfig_('REMITENTE') });
    return;
  }
  const ins = INSIGNIAS.filter(i => i.horas <= horas).pop();
  const html = plantillaCorreo_(
    '¡' + horas + ' horas encendiendo la llama del cambio! 🔥',
    '<p>Hola ' + nombre + ',</p>' +
    '<p>Acabas de cumplir <b>' + horas + ' horas de voluntariado</b> en Suyay y queremos agradecértelo como se merece. Cada hora que nos regalas se convierte en un niño o niña que lee, juega y se siente acompañado.</p>' +
    (ins ? '<p>Desbloqueaste la insignia <b>' + ins.nombre + '</b>. Ya la puedes ver en tu app.</p>' : '') +
    '<p>Te adjuntamos tu certificado con el detalle de tus jornadas. También queda guardado en <i>Mis certificados</i>' + (app ? ' de <a href="' + app + '">la app</a>' : '') + '.</p>' +
    '<p>Gracias por ser parte de la familia Suyay 💛</p>'
  );
  MailApp.sendEmail({ to: v.correo, subject: '🎉 Tu certificado de ' + horas + ' horas en Suyay', htmlBody: html, attachments: [pdf], name: getConfig_('REMITENTE') });
}

/** Menú > Probar certificado: te manda uno de muestra a ti, sin guardarlo. */
function probarCertificado() {
  const correo = Session.getActiveUser().getEmail();
  const v = { id: 'PRUEBA', nombre: 'Camila Fernández Quispe', dni: '71234567', ingreso_mes: 3, ingreso_anio: 2025, correo };
  const pdf = generarCertificado_(v, 40, 'prueba');
  MailApp.sendEmail({ to: correo, subject: 'Prueba de certificado Suyay', body: 'Así se ve el certificado de muestra.', attachments: [pdf], name: getConfig_('REMITENTE') });
  return 'Enviado a ' + correo;
}

/** Menú > Emitir constancia: certificado con las horas actuales (para MIMP, universidades, etc.). */
function emitirConstancia(dni) {
  const v = filas_('Voluntarios').find(x => String(x.dni) === String(dni).trim());
  if (!v) throw new Error('No hay ningún voluntario con ese DNI.');
  const horas = filas_('Asistencias').filter(a => a.voluntario_id === v.id).reduce((s, a) => s + Number(a.horas || 0), 0);
  generarCertificado_(v, horas, 'constancia');
  return 'Enviado a ' + v.correo + ' (' + horas + ' h) el ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
}

const INSIGNIAS = [
  { horas: 20, archivo: '020-chispa.png', nombre: 'Chispa' }, { horas: 40, archivo: '040-llamita.png', nombre: 'Llamita' },
  { horas: 60, archivo: '060-fogata.png', nombre: 'Fogata' }, { horas: 80, archivo: '080-antorcha.png', nombre: 'Antorcha' },
  { horas: 100, archivo: '100-faro.png', nombre: 'Faro' }, { horas: 120, archivo: '120-lucero.png', nombre: 'Lucero' },
  { horas: 140, archivo: '140-estrella.png', nombre: 'Estrella' }, { horas: 160, archivo: '160-constelacion.png', nombre: 'Constelación' },
  { horas: 180, archivo: '180-sol.png', nombre: 'Sol' }, { horas: 200, archivo: '200-corazon.png', nombre: 'Corazón Suyay' }
];
function insigniaPara_(horas) {
  const ins = INSIGNIAS.filter(i => i.horas <= horas).pop();
  return ins ? archivoEn_('CARPETA_INSIGNIAS_ID', ins.archivo) : null;
}

/* =========================================================
 * AUTOMATIZACIONES DIARIAS
 * ========================================================= */
function enviarCumpleanos() {
  const hoy = Utilities.formatDate(new Date(), TZ, 'MM-dd');
  const app = getConfig_('APP_URL');
  filas_('Voluntarios').forEach(v => {
    if (fechaISO_(v.nacimiento).slice(5) !== hoy || !v.correo) return;
    const nombre = String(v.nombre).split(' ')[0];
    const html = plantillaCorreo_('¡Feliz cumpleaños, ' + nombre + '! 🎂',
      '<p>Hoy toda la familia Suyay celebra contigo.</p>' +
      '<p>Gracias por compartir tu tiempo, tu energía y tu cariño con los niños y niñas. Que este nuevo año te traiga tantas alegrías como las que tú ayudas a encender.</p>' +
      '<p>¡Un abrazo enorme!</p>' + (app ? '<p><a href="' + app + '">Abre tu app</a>, hoy tiene una sorpresa para ti.</p>' : ''));
    MailApp.sendEmail({ to: v.correo, subject: '🎉 ¡Feliz cumpleaños de parte de Suyay!', htmlBody: html, name: getConfig_('REMITENTE') });
  });
}

function enviarPedidoFeedback() {
  const hoy = fechaISO_(new Date());
  const app = getConfig_('APP_URL');
  const acts = filas_('Actividades').filter(a => a.id && fechaISO_(a.fecha) === hoy);
  if (!acts.length) return;
  const vols = mapa_(filas_('Voluntarios'), 'id');
  const asis = filas_('Asistencias');
  acts.forEach(a => {
    asis.filter(x => x.actividad_id === a.id).forEach(x => {
      const v = vols[x.voluntario_id]; if (!v || !v.correo) return;
      const html = plantillaCorreo_('¿Cómo te fue hoy? 💬',
        '<p>Hola ' + String(v.nombre).split(' ')[0] + ', gracias por estar hoy en <b>' + a.titulo + '</b>.</p>' +
        '<p>Cuéntanos en menos de un minuto cómo la pasaste y qué mejorarías. Tu opinión nos ayuda a hacerlo mejor cada vez.</p>' +
        (app ? '<p style="text-align:center"><a href="' + app + '?r=feedback/' + a.id + '" style="background:#E34B14;color:#fff;padding:12px 22px;border-radius:30px;text-decoration:none;font-weight:bold">Dar mi opinión</a></p>' : ''));
      MailApp.sendEmail({ to: v.correo, subject: 'Tu opinión sobre ' + a.titulo, htmlBody: html, name: getConfig_('REMITENTE') });
    });
  });
}

function plantillaCorreo_(titulo, cuerpo) {
  return '<div style="background:#F4F4EE;padding:24px;font-family:Arial,sans-serif">' +
    '<div style="max-width:520px;margin:auto;background:#fff;border-radius:18px;overflow:hidden">' +
    '<div style="background:#112EA7;padding:22px 26px;color:#F4F4EE;font-size:22px;font-weight:bold">' + titulo + '</div>' +
    '<div style="padding:22px 26px;color:#2A2A3A;font-size:15px;line-height:1.55">' + cuerpo + '</div>' +
    '<div style="padding:14px 26px;background:#F2D13C;color:#112EA7;font-size:12px">Suyay Global Impact · enciende la llama del cambio</div>' +
    '</div></div>';
}

/* =========================================================
 * UBICACIONES — Lucía escribe la dirección y el sistema la convierte en coordenadas
 * ========================================================= */
function completarCoordenadas() {
  let n = 0;
  asignarIds_();
  filas_('Actividades').forEach(a => { if (a.lugar && (a.lat === '' || a.lng === '')) { if (coordenadas_(a)) n++; } });
  return n + ' ubicación(es) completada(s).';
}

/** Pone un código a cada actividad nueva (A0001, A0002...) para que Lucía no tenga que hacerlo. */
function asignarIds_() {
  const sh = hoja_('Actividades');
  const filas = filas_('Actividades');
  let max = filas.reduce((m, a) => Math.max(m, Number(String(a.id).replace(/\D/g, '')) || 0), 0);
  filas.forEach(a => { if (!a.id && a.titulo) sh.getRange(a._fila, 1).setValue('A' + Utilities.formatString('%04d', ++max)); });
}

function coordenadas_(a) {
  if (a.lat !== '' && a.lng !== '') return { lat: Number(a.lat), lng: Number(a.lng) };
  if (!a.lugar) return null;
  const res = Maps.newGeocoder().setRegion('pe').setLanguage('es').geocode(a.lugar);
  if (res.status !== 'OK' || !res.results.length) return null;
  const loc = res.results[0].geometry.location;
  const sh = hoja_('Actividades');
  const col = HOJAS.Actividades.indexOf('lat') + 1;
  sh.getRange(a._fila, col, 1, 2).setValues([[loc.lat, loc.lng]]);
  return { lat: loc.lat, lng: loc.lng };
}

/* =========================================================
 * UTILIDADES
 * ========================================================= */
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function hoja_(n) { return libro_().getSheetByName(n); }
function filas_(n) {
  const v = hoja_(n).getDataRange().getValues();
  const h = v.shift();
  return v.map((r, i) => { const o = { _fila: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; });
}
function mapa_(arr, k) { const m = {}; arr.forEach(x => m[x[k]] = x); return m; }
function getConfig_(k) {
  const c = CacheService.getScriptCache(); const hit = c.get('cfg_' + k); if (hit !== null) return hit;
  const f = filas_('Config').find(r => r.clave === k); const val = f ? String(f.valor) : '';
  c.put('cfg_' + k, val, 300); return val;
}
function setConfig_(k, val) {
  const f = filas_('Config').find(r => r.clave === k);
  if (f) hoja_('Config').getRange(f._fila, 2).setValue(val); else hoja_('Config').appendRow([k, val, '']);
  CacheService.getScriptCache().remove('cfg_' + k);
}
function voluntario_(token) {
  const v = token && filas_('Voluntarios').find(x => x.token === token);
  if (!v) throw new Error('SESION');
  return v;
}
function actividad_(id) {
  const a = filas_('Actividades').find(x => x.id === id);
  if (!a || a.estado === 'cancelada') throw new Error('Actividad no encontrada.');
  return a;
}
function archivoEn_(cfgCarpeta, nombre) {
  if (!nombre) return null;
  const it = DriveApp.getFolderById(getConfig_(cfgCarpeta)).getFilesByName(nombre);
  return it.hasNext() ? it.next() : null;
}
function fechaISO_(d) {
  if (!d) return '';
  if (d instanceof Date) return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  const s = String(d).trim();
  let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s); if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s); if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return s;
}
function restarDias_(iso, n) { const d = new Date(iso + 'T12:00:00'); d.setDate(d.getDate() - n); return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function hora_(h) {
  if (h instanceof Date) return Utilities.formatDate(h, TZ, 'HH:mm');
  const m = /^(\d{1,2}):(\d{2})/.exec(String(h || '').trim());
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}
function fechaHora_(fecha, hora) {
  // Lima es UTC-5 todo el año (sin horario de verano)
  return new Date(fechaISO_(fecha) + 'T' + (hora_(hora) || '00:00') + ':00-05:00');
}
function distancia_(la1, lo1, la2, lo2) {
  const R = 6371000, rad = x => x * Math.PI / 180;
  const dLa = rad(la2 - la1), dLo = rad(lo2 - lo1);
  const a = Math.sin(dLa / 2) ** 2 + Math.cos(rad(la1)) * Math.cos(rad(la2)) * Math.sin(dLo / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}

/** Guarda en Recursos una imagen desde un link (lo usa Claude para traer los fondos exportados de Canva). */
function guardarFondoDesdeUrl(url, nombre) {
  const blob = UrlFetchApp.fetch(url).getBlob().setName(nombre);
  const carpeta = DriveApp.getFolderById(getConfig_('CARPETA_RECURSOS_ID'));
  const it = carpeta.getFilesByName(nombre);
  while (it.hasNext()) it.next().setTrashed(true);
  return carpeta.createFile(blob).getUrl();
}

/* =========================================================
 * PLANTILLA DEL CERTIFICADO — se crea sola (A4 apaisado) con los fondos de Canva
 * ========================================================= */
function crearPlantilla(urlPptxA4) {
  const fondo1 = archivoEn_('CARPETA_RECURSOS_ID', 'fondo-certificado.png');
  const fondo2 = archivoEn_('CARPETA_RECURSOS_ID', 'fondo-anexo.png');
  if (!fondo1 || !fondo2) throw new Error('Faltan fondo-certificado.png y fondo-anexo.png en la carpeta Recursos.');
  // Google Slides no permite elegir A4 al crear, así que partimos de un archivo A4 exportado desde Canva
  const viejo = getConfig_('PLANTILLA_ID');
  if (viejo) { try { DriveApp.getFileById(viejo).setTrashed(true); } catch (e) {} }
  const nueva = Drive.Files.create({ name: 'Plantilla Certificado Suyay', mimeType: 'application/vnd.google-apps.presentation', parents: [getConfig_('CARPETA_RAIZ_ID')] }, UrlFetchApp.fetch(urlPptxA4).getBlob());
  const pres = SlidesApp.openById(nueva.id);
  const [s1, s2] = pres.getSlides();
  pres.getSlides().forEach(sl => sl.getPageElements().forEach(p => p.remove()));
  const W = pres.getPageWidth(), H = pres.getPageHeight();
  [[s1, fondo1], [s2, fondo2]].forEach(([sl, f]) => sl.insertImage(f.getBlob()).setLeft(0).setTop(0).setWidth(W).setHeight(H));
  const px = v => v * 0.75, CX = 581.5;
  const AZUL = '#112EA7', TXT = '#2A2A3A';
  function caja(sl, x, y, w, h, texto, fam, peso, tam, color, alin, extra) {
    const tb = sl.insertTextBox(texto, px(x), px(y), px(w), px(h));
    const tr = tb.getText();
    const st = tr.getTextStyle().setFontFamilyAndWeight(fam, peso).setFontSize(tam).setForegroundColor(color);
    if (extra && extra.italic) st.setItalic(true);
    tr.getParagraphStyle().setParagraphAlignment(alin || SlidesApp.ParagraphAlignment.CENTER).setLineSpacing((extra && extra.ls) || 100).setSpaceAbove(0).setSpaceBelow(0);
    if (extra && extra.medio) tb.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    return tb;
  }
  const L = SlidesApp.ParagraphAlignment.START;
  caja(s1, CX - 390, 274, 780, 52, 'La Asociación Suyay Global Impact, con RUC N.° {{RUC}}{{INSCRIPCION}}, en el marco de la Ley N.° 28238, Ley General del Voluntariado, certifica que:', 'Merriweather', 400, 11, TXT, null, { italic: true, ls: 115 });
  caja(s1, CX - 420, 334, 840, 64, '{{NOMBRE}}', 'Fredoka', 700, 32, AZUL, null, { medio: true });
  const cuerpo = caja(s1, CX - 400, 416, 800, 108, 'identificado(a) con DNI N.° {{DNI}}, participa como voluntario(a) de Suyay Global Impact desde {{DESDE}} en acciones de fomento de la lectura y desarrollo integral de niñas, niños y adolescentes de comunidades vulnerables, habiendo acreditado al {{FECHA_CORTE}} un total de {{JORNADAS}} jornadas y {{HORAS}} horas de voluntariado, conforme al detalle adjunto.', 'Quicksand', 400, 12, TXT, null, { ls: 125 });
  const t = cuerpo.getText(), marca = '{{JORNADAS}} jornadas y {{HORAS}} horas de voluntariado', i = t.asString().indexOf(marca);
  t.getRange(i, i + marca.length).getTextStyle().setFontFamilyAndWeight('Quicksand', 700).setForegroundColor(AZUL);
  caja(s1, CX - 400, 524, 800, 24, 'Se expide el presente certificado para los fines que el(la) interesado(a) estime conveniente.', 'Quicksand', 400, 10, '#55556A');
  [[CX - 210, 'Valeria Riojas', 'Presidenta y representante legal'], [CX + 210, '{{FIRMANTE2_NOMBRE}}', '{{FIRMANTE2_CARGO}}']].forEach(([cx, n, c]) => {
    caja(s1, cx - 150, 656, 300, 24, n, 'Fredoka', 700, 12.5, AZUL);
    caja(s1, cx - 150, 678, 300, 22, c, 'Quicksand', 400, 9.5, TXT);
  });
  caja(s1, CX - 400, 724, 800, 20, 'Lima, {{FECHA}}  ·  Código de verificación: {{CODIGO}}', 'Quicksand', 400, 8, '#7A7A88');
  caja(s2, 104, 194, 880, 24, '{{NOMBRE}}  ·  DNI N.° {{DNI}}  ·  Código {{CODIGO}}', 'Quicksand', 400, 11, TXT, L);
  caja(s2, 104, 232, 880, 400, '{{TABLA}}', 'Quicksand', 400, 9, '#B0B0BC', L);
  caja(s2, 104, 668, 880, 44, 'Capacitaciones recibidas: {{CAPACITACIONES}}', 'Quicksand', 400, 10, TXT, L, { ls: 110 });
  caja(s2, 104, 724, 880, 20, 'Registro generado por el sistema de voluntariado de Suyay Global Impact a partir de asistencias validadas por ubicación y horario.', 'Quicksand', 400, 8, '#7A7A88', L, { italic: true });
  const url = pres.getUrl();
  pres.saveAndClose();
  setConfig_('PLANTILLA_ID', nueva.id);
  return url;
}

/** Trae los fondos exportados de Canva (links temporales) a Recursos y arma la plantilla. */
function instalarFondos(url1, url2) {
  guardarFondoDesdeUrl(url1, 'fondo-certificado.png');
  guardarFondoDesdeUrl(url2, 'fondo-anexo.png');
  return 'fondos guardados';
}
