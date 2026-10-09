/***********************************************************************
 * QUALITY CCI v2 — Control de Calidad Interno de Parasitología
 * Versión multi-marca
 * -----------------------------------------------------------------------
 * Este script debe estar VINCULADO a una Google Sheet (Extensiones >
 * Apps Script desde la hoja). Antes de publicarlo, ejecutar UNA VEZ
 * desde el editor la función "configurarInicial" (ver README.md).
 *
 * HOJAS QUE SE CREAN AUTOMÁTICAMENTE:
 *   Marcas         -> laboratorios participantes (código, zona horaria…)
 *   Usuarios       -> cuentas (contraseñas cifradas, nunca en texto)
 *   BancoImagenes  -> imágenes de cada marca (las de la marca de origen
 *                     forman el banco general que se comparte)
 *   Asignaciones   -> par de imágenes que recibió cada marca cada día
 *   Registros      -> respuestas y su calificación
 *
 * CICLO DE VIDA DE LAS IMÁGENES:
 *   - No se consumen: rotan sin repetirse para la misma marca en 90 días
 *     (si el banco alcanza), eligiendo la que no ha visto o vio hace más.
 *   - A los 6 meses de subida salen de la rotación (se archivan).
 *   - A los 9 meses se borran definitivamente.
 *
 * Todas las celdas se guardan como TEXTO para que Sheets no convierta
 * fechas u horas en valores de tipo Fecha.
 ***********************************************************************/

var CONFIG = {
  CARPETA_IMAGENES: 'QualityCCI_Imagenes',
  CARPETA_ARCHIVO: 'QualityCCI_Archivo',
  MESES_ROTACION: 6,
  MESES_ELIMINACION: 9,
  DIAS_SIN_REPETIR: 90,
  MAX_INTENTOS: 3,
  HORAS_SESION: 6,            // máximo que permite CacheService
  ZONA_DEFAULT: 'America/Mexico_City',
  ITERACIONES_HASH: 200,
  MAX_FALLOS_LOGIN: 5,
  MINUTOS_BLOQUEO: 15,
  BANCO_POR_PAGINA: 6,
  HISTORIAL_POR_PAGINA: 10
};

// [código, número, nombre, ¿es la marca de origen?]
var MARCAS_INICIALES = [
  ['LST', '01', 'Lister', 'SI'],
  ['AZT', '02', 'Azteca', 'NO'],
  ['SWL', '03', 'Swiss Lab', 'NO'],
  ['PLB', '04', 'Polab', 'NO'],
  ['JNR', '05', 'Jenner', 'NO'],
  ['MRA', '06', 'Moreira', 'NO'],
  ['LCS', '07', 'Liacsa', 'NO'],
  ['BMD', '08', 'Biomedica', 'NO'],
  ['EXK', '09', 'Exakta', 'NO'],
  ['PMD', '10', 'Promedic', 'NO'],
  ['FML', '11', 'FamilyLabs', 'NO']
];

var HOJAS = {
  Marcas: ['codigo', 'numero', 'nombre', 'zonaHoraria', 'activa', 'esOrigen'],
  Usuarios: ['usuario', 'nombre', 'marca', 'rol', 'admin', 'hash', 'salt', 'temporal', 'activo', 'creado'],
  BancoImagenes: ['id', 'marca', 'compartida', 'tipo', 'parasito', 'fuente', 'fileId', 'thumbId', 'fechaSubida', 'estado'],
  Asignaciones: ['fecha', 'marca', 'orden', 'tipo', 'idImagen', 'fileId', 'parasito', 'fuente'],
  Registros: ['id', 'fecha', 'marca', 'usuario', 'hora', 'respuesta1', 'respuesta2', 'estado', 'intento']
};

var NEGATIVO = 'Negativo a parásitos';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Quality CCI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============================ hojas ============================ */

function hoja_(nombre) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(nombre);
  if (!sh) {
    var headers = HOJAS[nombre];
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function texto_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.ZONA_DEFAULT, 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

// Devuelve las filas como objetos {columna: valor}, con _fila = número de fila real.
function leer_(nombre) {
  var data = hoja_(nombre).getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i].join('') === '') continue;
    var obj = { _fila: i + 1 };
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = texto_(data[i][j]);
    out.push(obj);
  }
  return out;
}

function agregar_(nombre, obj) {
  agregarVarios_(nombre, [obj]);
}

function agregarVarios_(nombre, objs) {
  if (!objs.length) return;
  var sh = hoja_(nombre);
  var headers = HOJAS[nombre];
  var filas = objs.map(function (obj) {
    return headers.map(function (h) { return texto_(obj[h]); });
  });
  var inicio = sh.getLastRow() + 1;
  var faltan = inicio + filas.length - 1 - sh.getMaxRows();
  if (faltan > 0) sh.insertRowsAfter(sh.getMaxRows(), faltan + 100);
  var rango = sh.getRange(inicio, 1, filas.length, headers.length);
  rango.setNumberFormat('@');
  rango.setValues(filas);
}

function actualizar_(nombre, fila, cambios) {
  var sh = hoja_(nombre);
  var headers = HOJAS[nombre];
  Object.keys(cambios).forEach(function (col) {
    var idx = headers.indexOf(col);
    if (idx === -1) return;
    sh.getRange(fila, idx + 1).setNumberFormat('@').setValue(texto_(cambios[col]));
  });
}

function conBloqueo_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ============================ fechas ============================ */

function zona_(perfil) { return (perfil && perfil.zona) || CONFIG.ZONA_DEFAULT; }
function hoy_(perfil) { return Utilities.formatDate(new Date(), zona_(perfil), 'yyyy-MM-dd'); }
function ahora_(perfil) { return Utilities.formatDate(new Date(), zona_(perfil), 'HH:mm:ss'); }

function fechaUTC_(fecha) { return new Date(fecha + 'T12:00:00Z'); }
function sumarDias_(fecha, dias) {
  var d = fechaUTC_(fecha);
  d.setUTCDate(d.getUTCDate() + dias);
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}
function sumarMeses_(fecha, meses) {
  var d = fechaUTC_(fecha);
  d.setUTCMonth(d.getUTCMonth() + meses);
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}
function diasEntre_(desde, hasta) {
  return Math.round((fechaUTC_(hasta) - fechaUTC_(desde)) / 86400000);
}

/* ======================= contraseñas y sesión ======================= */

function hash_(password, salt) {
  var v = salt + '|' + password;
  for (var i = 0; i < CONFIG.ITERACIONES_HASH; i++) {
    v = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, v, Utilities.Charset.UTF_8));
  }
  return v;
}

function generarPassword_() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  var out = '';
  for (var i = 0; i < 8; i++) out += chars.charAt(Math.floor(Math.random() * chars.length));
  return out;
}

function nuevoId_() { return Utilities.getUuid().replace(/-/g, '').substring(0, 10); }

function cache_() { return CacheService.getScriptCache(); }

function perfilDe_(u, marca) {
  return {
    usuario: u.usuario,
    nombre: u.nombre,
    marca: u.marca,
    numero: marca.numero,
    rol: u.rol,
    admin: u.admin === 'SI',
    esOrigen: marca.esOrigen === 'SI',
    zona: marca.zonaHoraria || CONFIG.ZONA_DEFAULT,
    temporal: u.temporal === 'SI'
  };
}

function guardarSesion_(token, perfil) {
  cache_().put('s_' + token, JSON.stringify(perfil), CONFIG.HORAS_SESION * 3600);
}

// Valida el token. requisito: undefined (cualquiera), 'analista', 'supervisor' o 'admin'.
function sesion_(token, requisito, permitirTemporal) {
  var raw = token ? cache_().get('s_' + token) : null;
  if (!raw) throw new Error('SESION_EXPIRADA');
  var p = JSON.parse(raw);
  if (p.temporal && !permitirTemporal) throw new Error('Debes cambiar tu contraseña temporal');
  if (requisito === 'analista' && p.rol !== 'analista') throw new Error('No autorizado');
  if (requisito === 'supervisor' && p.rol !== 'supervisor') throw new Error('No autorizado');
  if (requisito === 'admin' && !(p.rol === 'supervisor' && p.admin)) throw new Error('No autorizado');
  return p;
}

function marcas_() { return leer_('Marcas'); }
function marca_(codigo) { return marcas_().filter(function (m) { return m.codigo === codigo; })[0]; }

function login(usuario, password) {
  usuario = String(usuario || '').trim().toUpperCase();
  password = String(password || '');
  if (!usuario || !password) return { success: false, error: 'Escribe tu usuario y contraseña' };

  var claveFallos = 'f_' + usuario;
  var fallos = Number(cache_().get(claveFallos) || 0);
  if (fallos >= CONFIG.MAX_FALLOS_LOGIN) {
    return { success: false, error: 'Demasiados intentos. Espera ' + CONFIG.MINUTOS_BLOQUEO + ' minutos.' };
  }

  var u = leer_('Usuarios').filter(function (x) { return x.usuario === usuario; })[0];
  if (!u || u.activo !== 'SI' || hash_(password, u.salt) !== u.hash) {
    cache_().put(claveFallos, String(fallos + 1), CONFIG.MINUTOS_BLOQUEO * 60);
    return { success: false, error: 'Usuario o contraseña incorrectos' };
  }
  var marca = marca_(u.marca);
  if (!marca || marca.activa !== 'SI') return { success: false, error: 'Tu laboratorio no está activo' };

  cache_().remove(claveFallos);
  var token = Utilities.getUuid() + nuevoId_();
  var perfil = perfilDe_(u, marca);
  guardarSesion_(token, perfil);
  return { success: true, token: token, perfil: perfil };
}

function getPerfil(token) {
  return { success: true, perfil: sesion_(token, undefined, true) };
}

function cerrarSesion(token) {
  if (token) cache_().remove('s_' + token);
  return { success: true };
}

function cambiarPassword(token, actual, nueva) {
  var p = sesion_(token, undefined, true);
  nueva = String(nueva || '');
  if (nueva.length < 6) return { success: false, error: 'La nueva contraseña debe tener al menos 6 caracteres' };
  if (nueva === actual) return { success: false, error: 'La nueva contraseña debe ser distinta a la actual' };
  var u = leer_('Usuarios').filter(function (x) { return x.usuario === p.usuario; })[0];
  if (!u || hash_(String(actual || ''), u.salt) !== u.hash) return { success: false, error: 'La contraseña actual no es correcta' };
  var salt = nuevoId_();
  actualizar_('Usuarios', u._fila, { hash: hash_(nueva, salt), salt: salt, temporal: 'NO' });
  p.temporal = false;
  guardarSesion_(token, p);
  return { success: true, perfil: p };
}

/* ============================ Drive ============================ */

function carpeta_(nombre) {
  var it = DriveApp.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : DriveApp.createFolder(nombre);
}

function hacerPublico_(file) {
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
}

function urlImagen_(fileId, ancho) {
  return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w' + (ancho || 1400);
}

/* ===================== selección de imágenes ===================== */

// Imágenes que puede recibir una marca: las suyas + (si no es la de origen)
// las compartidas del banco general.
function poolMarca_(banco, perfil, tipo) {
  return banco.filter(function (b) {
    return b.estado === 'activa' && b.tipo === tipo &&
      (b.marca === perfil.marca || (!perfil.esOrigen && b.compartida === 'SI'));
  });
}

function ultimaVez_(asignaciones, marca) {
  var mapa = {};
  asignaciones.forEach(function (a) {
    if (a.marca !== marca) return;
    if (!mapa[a.idImagen] || a.fecha > mapa[a.idImagen]) mapa[a.idImagen] = a.fecha;
  });
  return mapa;
}

// Elige la imagen que la marca nunca ha visto; si ya vio todas, la que vio hace más tiempo.
function elegirImagen_(pool, vistas) {
  if (!pool.length) return null;
  var nuevas = pool.filter(function (b) { return !vistas[b.id]; });
  var candidatas = nuevas;
  if (!candidatas.length) {
    var minima = pool.reduce(function (min, b) { return vistas[b.id] < min ? vistas[b.id] : min; }, '9999-12-31');
    candidatas = pool.filter(function (b) { return vistas[b.id] === minima; });
  }
  return candidatas[Math.floor(Math.random() * candidatas.length)];
}

function asignacionDe_(filas) {
  var a = filas.filter(function (f) { return f.orden === 'A'; })[0];
  var b = filas.filter(function (f) { return f.orden === 'B'; })[0];
  return (a && b) ? { A: a, B: b } : null;
}

// Devuelve el par de imágenes de hoy para la marca; si no existe y crear=true, lo genera.
function asignacionHoy_(perfil, crear) {
  var fecha = hoy_(perfil);
  var clave = 'a_' + perfil.marca + '_' + fecha;
  var enCache = cache_().get(clave);
  if (enCache) return JSON.parse(enCache);

  var deHoy = function () {
    return leer_('Asignaciones').filter(function (a) { return a.marca === perfil.marca && a.fecha === fecha; });
  };
  var asignacion = asignacionDe_(deHoy());
  if (!asignacion && crear) {
    asignacion = conBloqueo_(function () {
      var filas = deHoy();
      if (!asignacionDe_(filas)) filas = filas.concat(crearAsignacion_(perfil, fecha, filas));
      return asignacionDe_(filas);
    });
  }
  if (asignacion) cache_().put(clave, JSON.stringify(asignacion), 21600);
  return asignacion;
}

function crearAsignacion_(perfil, fecha, existentes) {
  var banco = leer_('BancoImagenes');
  var vistas = ultimaVez_(leer_('Asignaciones'), perfil.marca);
  var ordenes = Math.random() < 0.5 ? ['A', 'B'] : ['B', 'A'];
  var ocupadas = existentes.map(function (e) { return e.orden; });
  var nuevas = [];
  ['Positiva', 'Negativa'].forEach(function (tipo, i) {
    if (existentes.some(function (e) { return e.tipo === tipo; })) return;
    var elegida = elegirImagen_(poolMarca_(banco, perfil, tipo), vistas);
    if (!elegida) return;
    var orden = ocupadas.indexOf(ordenes[i]) === -1 ? ordenes[i] : ordenes[1 - i];
    ocupadas.push(orden);
    nuevas.push({
      fecha: fecha, marca: perfil.marca, orden: orden, tipo: tipo, idImagen: elegida.id,
      fileId: elegida.fileId, parasito: elegida.parasito, fuente: elegida.fuente
    });
  });
  agregarVarios_('Asignaciones', nuevas);
  return nuevas;
}

/* ========================= control diario ========================= */

function registrosHoy_(perfil) {
  var fecha = hoy_(perfil);
  return leer_('Registros').filter(function (r) { return r.marca === perfil.marca && r.fecha === fecha; });
}

function estadoDeRegistros_(registros) {
  var aprobado = registros.filter(function (r) { return r.estado === 'Aprobado'; })[0];
  var ultimo = registros[registros.length - 1];
  return {
    intentosUsados: registros.length,
    intentosRestantes: Math.max(0, CONFIG.MAX_INTENTOS - registros.length),
    aprobado: !!aprobado,
    agotado: !aprobado && registros.length >= CONFIG.MAX_INTENTOS,
    usuario: ultimo ? ultimo.usuario : '',
    hora: ultimo ? ultimo.hora : ''
  };
}

// El analista nunca recibe el parásito correcto ni el tipo de imagen.
function getControlHoy(token) {
  var p = sesion_(token);
  var asignacion = asignacionHoy_(p, true);
  var estado = estadoDeRegistros_(registrosHoy_(p));
  if (!asignacion) return { success: true, hay: false, estado: estado, fecha: hoy_(p) };

  var esSup = p.rol === 'supervisor';
  var positiva = asignacion.A.tipo === 'Positiva' ? asignacion.A : asignacion.B;
  var imagenes = ['A', 'B'].map(function (orden) {
    var a = asignacion[orden];
    var img = { url: urlImagen_(a.fileId, 1400) };
    if (esSup) { img.parasito = a.parasito; img.tipo = a.tipo; }
    return img;
  });
  return { success: true, hay: true, estado: estado, fecha: hoy_(p), fuente: positiva.fuente, imagenes: imagenes };
}

function normalizarRespuesta_(s) {
  return String(s || '').split('+').map(function (x) { return x.trim().toLowerCase(); })
    .filter(function (x) { return x; }).sort().join('+');
}

function enviarRespuesta(token, respuesta1, respuesta2) {
  var p = sesion_(token, 'analista');
  respuesta1 = String(respuesta1 || '').trim();
  respuesta2 = String(respuesta2 || '').trim();
  if (!respuesta1 || !respuesta2) return { success: false, error: 'Responde ambas imágenes' };

  return conBloqueo_(function () {
    var estado = estadoDeRegistros_(registrosHoy_(p));
    if (estado.aprobado) return { success: false, error: 'El control de hoy ya fue aprobado' };
    if (estado.agotado) return { success: false, error: 'Ya se usaron los ' + CONFIG.MAX_INTENTOS + ' intentos de hoy' };

    var asignacion = asignacionHoy_(p, false);
    if (!asignacion) return { success: false, error: 'No hay imágenes asignadas hoy' };

    var correcto = normalizarRespuesta_(asignacion.A.parasito) === normalizarRespuesta_(respuesta1) &&
      normalizarRespuesta_(asignacion.B.parasito) === normalizarRespuesta_(respuesta2);
    var intento = estado.intentosUsados + 1;
    var registro = {
      id: nuevoId_(), fecha: hoy_(p), marca: p.marca, usuario: p.usuario, hora: ahora_(p),
      respuesta1: respuesta1, respuesta2: respuesta2, estado: correcto ? 'Aprobado' : 'Rechazado', intento: intento
    };
    agregar_('Registros', registro);
    return {
      success: true, estado: registro.estado, intento: intento,
      intentosRestantes: CONFIG.MAX_INTENTOS - intento, hora: registro.hora, usuario: p.usuario
    };
  });
}

/* ========================= historial ========================= */

function getHistorial(token, pagina) {
  var p = sesion_(token);
  pagina = Number(pagina) || 1;
  var hoy = hoy_(p);
  var esSup = p.rol === 'supervisor';

  var porFecha = {};
  leer_('Asignaciones').forEach(function (a) {
    if (a.marca !== p.marca) return;
    porFecha[a.fecha] = porFecha[a.fecha] || {};
    porFecha[a.fecha][a.orden] = a;
  });
  var registros = leer_('Registros').filter(function (r) { return r.marca === p.marca; });
  var conRespuesta = {};
  registros.forEach(function (r) { conRespuesta[r.fecha] = true; });

  var filas = registros.map(function (r) {
    var f = { fecha: r.fecha, usuario: r.usuario, hora: r.hora, respuesta1: r.respuesta1, respuesta2: r.respuesta2, intento: r.intento, estado: r.estado };
    if (esSup) {
      var dia = porFecha[r.fecha] || {};
      f.correcta1 = dia.A ? dia.A.parasito : '';
      f.correcta2 = dia.B ? dia.B.parasito : '';
    }
    return f;
  });
  Object.keys(porFecha).forEach(function (fecha) {
    if (fecha < hoy && !conRespuesta[fecha]) {
      filas.push({ fecha: fecha, usuario: '—', hora: '', respuesta1: '', respuesta2: '', intento: '', estado: 'Sin respuesta' });
    }
  });

  filas.sort(function (a, b) {
    var ka = a.fecha + 'T' + (a.hora || '00:00:00'), kb = b.fecha + 'T' + (b.hora || '00:00:00');
    return kb > ka ? 1 : (kb < ka ? -1 : 0);
  });
  var totalPaginas = Math.max(1, Math.ceil(filas.length / CONFIG.HISTORIAL_POR_PAGINA));
  pagina = Math.min(pagina, totalPaginas);
  return {
    success: true, pagina: pagina, totalPaginas: totalPaginas, total: filas.length,
    registros: filas.slice((pagina - 1) * CONFIG.HISTORIAL_POR_PAGINA, pagina * CONFIG.HISTORIAL_POR_PAGINA)
  };
}

// Indicadores de los últimos 30 días (desde que la marca empezó a usar el sistema).
function getIndicadores(token) {
  var p = sesion_(token, 'supervisor');
  var hoy = hoy_(p);
  var desde = sumarDias_(hoy, -29);
  var asignadas = leer_('Asignaciones').filter(function (a) { return a.marca === p.marca; });
  var inicio = asignadas.reduce(function (min, a) { return a.fecha < min ? a.fecha : min; }, hoy);
  if (inicio > desde) desde = inicio;

  var porDia = {};
  leer_('Registros').forEach(function (r) {
    if (r.marca !== p.marca || r.fecha < desde || r.fecha > hoy) return;
    (porDia[r.fecha] = porDia[r.fecha] || []).push(r);
  });

  var res = { dias: 0, primerIntento: 0, aprobadoReintento: 0, noAprobado: 0, sinRespuesta: 0, desde: desde };
  for (var f = desde; f <= hoy; f = sumarDias_(f, 1)) {
    var regs = porDia[f] || [];
    if (f === hoy && !regs.length) continue; // hoy todavía puede responderse
    res.dias++;
    var aprobado = regs.filter(function (r) { return r.estado === 'Aprobado'; })[0];
    if (!regs.length) res.sinRespuesta++;
    else if (aprobado && aprobado.intento === '1') res.primerIntento++;
    else if (aprobado) res.aprobadoReintento++;
    else res.noAprobado++;
  }
  res.porcentajePrimerIntento = res.dias ? Math.round(res.primerIntento * 100 / res.dias) : 0;
  return { success: true, indicadores: res };
}

/* ======================== banco de imágenes ======================== */

function subirImagenBanco(token, base64, mimeType, parasito, fuente, tipo, thumbBase64) {
  var p = sesion_(token, 'supervisor');
  tipo = tipo === 'Negativa' ? 'Negativa' : 'Positiva';
  parasito = String(parasito || '').trim();
  if (!base64) return { success: false, error: 'Falta la imagen' };
  if (tipo === 'Positiva' && !parasito) return { success: false, error: 'Falta el nombre del parásito' };

  var id = nuevoId_();
  var folder = carpeta_(CONFIG.CARPETA_IMAGENES);
  var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(base64.split(',').pop()), mimeType || 'image/jpeg', id + '.jpg'));
  hacerPublico_(file);
  var thumbId = '';
  if (thumbBase64) {
    var thumb = folder.createFile(Utilities.newBlob(Utilities.base64Decode(thumbBase64.split(',').pop()), 'image/jpeg', id + '_thumb.jpg'));
    hacerPublico_(thumb);
    thumbId = thumb.getId();
  }
  agregar_('BancoImagenes', {
    id: id, marca: p.marca, compartida: p.esOrigen ? 'SI' : 'NO', tipo: tipo,
    parasito: tipo === 'Negativa' ? NEGATIVO : parasito,
    fuente: tipo === 'Negativa' ? '' : String(fuente || '').trim(),
    fileId: file.getId(), thumbId: thumbId, fechaSubida: hoy_(p), estado: 'activa'
  });
  return { success: true, id: id };
}

function getBanco(token, tipo, pagina) {
  var p = sesion_(token, 'supervisor');
  tipo = tipo === 'Negativa' ? 'Negativa' : 'Positiva';
  pagina = Number(pagina) || 1;
  var hoy = hoy_(p);
  var banco = leer_('BancoImagenes');

  var propias = banco.filter(function (b) { return b.marca === p.marca && b.estado === 'activa' && b.tipo === tipo; });
  propias.sort(function (a, b) { return b.fechaSubida > a.fechaSubida ? 1 : -1; });

  var pool = poolMarca_(banco, p, tipo);
  var vistas = ultimaVez_(leer_('Asignaciones'), p.marca);
  var limite = sumarDias_(hoy, -CONFIG.DIAS_SIN_REPETIR);
  var disponibles = pool.filter(function (b) { return !vistas[b.id] || vistas[b.id] <= limite; }).length;

  var totalPaginas = Math.max(1, Math.ceil(propias.length / CONFIG.BANCO_POR_PAGINA));
  pagina = Math.min(pagina, totalPaginas);
  var items = propias.slice((pagina - 1) * CONFIG.BANCO_POR_PAGINA, pagina * CONFIG.BANCO_POR_PAGINA).map(function (b) {
    return {
      id: b.id, parasito: b.parasito, fuente: b.fuente, fechaSubida: b.fechaSubida,
      diasEnRotacion: Math.max(0, diasEntre_(hoy, sumarMeses_(b.fechaSubida, CONFIG.MESES_ROTACION))),
      url: urlImagen_(b.thumbId || b.fileId, 200)
    };
  });
  return {
    success: true, tipo: tipo, pagina: pagina, totalPaginas: totalPaginas, banco: items,
    resumen: {
      propias: propias.length,
      general: p.esOrigen ? 0 : pool.length - propias.length,
      totalRotacion: pool.length,
      disponiblesSinRepetir: disponibles,
      recomendado: CONFIG.DIAS_SIN_REPETIR
    }
  };
}

function eliminarImagenBanco(token, id) {
  var p = sesion_(token, 'supervisor');
  var fila = leer_('BancoImagenes').filter(function (b) { return b.id === id && b.marca === p.marca && b.estado !== 'eliminada'; })[0];
  if (!fila) return { success: false, error: 'Imagen no encontrada' };
  borrarArchivos_(fila);
  actualizar_('BancoImagenes', fila._fila, { estado: 'eliminada' });
  return { success: true };
}

function borrarArchivos_(fila) {
  [fila.fileId, fila.thumbId].forEach(function (fid) {
    if (fid) { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {} }
  });
}

/* ========================= usuarios ========================= */

function usuarioPublico_(u) {
  return { usuario: u.usuario, nombre: u.nombre, marca: u.marca, rol: u.rol, admin: u.admin === 'SI', activo: u.activo === 'SI', temporal: u.temporal === 'SI', creado: u.creado };
}

function getUsuarios(token) {
  var p = sesion_(token, 'supervisor');
  var lista = leer_('Usuarios').filter(function (u) { return p.admin || u.marca === p.marca; }).map(usuarioPublico_);
  lista.sort(function (a, b) { return a.usuario > b.usuario ? 1 : -1; });
  return { success: true, usuarios: lista };
}

function puedeGestionar_(p, u) {
  if (u.usuario === p.usuario) return false;
  if (p.admin) return true;
  return u.marca === p.marca && u.rol === 'analista';
}

function crearUsuario(token, nombre, rol, marca) {
  var p = sesion_(token, 'supervisor');
  nombre = String(nombre || '').trim();
  if (!nombre) return { success: false, error: 'Escribe el nombre completo' };
  rol = rol === 'supervisor' ? 'supervisor' : 'analista';
  marca = p.admin ? String(marca || p.marca) : p.marca;
  if (!p.admin && rol !== 'analista') return { success: false, error: 'No autorizado' };
  if (!marca_(marca)) return { success: false, error: 'Marca no válida' };

  return conBloqueo_(function () {
    var usuarios = leer_('Usuarios');
    var codigo;
    if (rol === 'supervisor') {
      codigo = marca + '-SUP';
      if (usuarios.some(function (u) { return u.usuario === codigo; })) return { success: false, error: 'Esa marca ya tiene supervisor (' + codigo + ')' };
    } else {
      var n = 1;
      do { codigo = marca + '-' + (n < 10 ? '0' : '') + n; n++; }
      while (usuarios.some(function (u) { return u.usuario === codigo; }));
    }
    var temporal = generarPassword_();
    var salt = nuevoId_();
    agregar_('Usuarios', {
      usuario: codigo, nombre: nombre, marca: marca, rol: rol, admin: 'NO', hash: hash_(temporal, salt),
      salt: salt, temporal: 'SI', activo: 'SI', creado: hoy_(p)
    });
    return { success: true, usuario: codigo, password: temporal };
  });
}

function restablecerPassword(token, usuario) {
  var p = sesion_(token, 'supervisor');
  var u = leer_('Usuarios').filter(function (x) { return x.usuario === usuario; })[0];
  if (!u || !puedeGestionar_(p, u)) return { success: false, error: 'No autorizado' };
  var temporal = generarPassword_();
  var salt = nuevoId_();
  actualizar_('Usuarios', u._fila, { hash: hash_(temporal, salt), salt: salt, temporal: 'SI' });
  return { success: true, usuario: u.usuario, password: temporal };
}

function cambiarEstadoUsuario(token, usuario, activo) {
  var p = sesion_(token, 'supervisor');
  var u = leer_('Usuarios').filter(function (x) { return x.usuario === usuario; })[0];
  if (!u || !puedeGestionar_(p, u)) return { success: false, error: 'No autorizado' };
  actualizar_('Usuarios', u._fila, { activo: activo ? 'SI' : 'NO' });
  return { success: true };
}

/* ====================== marcas (solo admin) ====================== */

function getMarcas(token) {
  sesion_(token, 'admin');
  var usuarios = leer_('Usuarios');
  var lista = marcas_().map(function (m) {
    return {
      codigo: m.codigo, numero: m.numero, nombre: m.nombre, zonaHoraria: m.zonaHoraria,
      activa: m.activa === 'SI', esOrigen: m.esOrigen === 'SI',
      tieneSupervisor: usuarios.some(function (u) { return u.marca === m.codigo && u.rol === 'supervisor'; }),
      usuarios: usuarios.filter(function (u) { return u.marca === m.codigo && u.activo === 'SI'; }).length
    };
  });
  return { success: true, marcas: lista };
}

function cambiarEstadoMarca(token, codigo, activa) {
  sesion_(token, 'admin');
  var m = marca_(codigo);
  if (!m) return { success: false, error: 'Marca no encontrada' };
  if (m.esOrigen === 'SI' && !activa) return { success: false, error: 'La marca de origen no se puede desactivar' };
  actualizar_('Marcas', m._fila, { activa: activa ? 'SI' : 'NO' });
  return { success: true };
}

/* ====================== mantenimiento (disparador) ====================== */

// Se ejecuta sola cada día (ver configurarInicial). A los 6 meses archiva,
// a los 9 meses borra definitivamente.
function mantenimientoDiario() {
  var hoy = Utilities.formatDate(new Date(), CONFIG.ZONA_DEFAULT, 'yyyy-MM-dd');
  var limiteArchivo = sumarMeses_(hoy, -CONFIG.MESES_ROTACION);
  var limiteBorrado = sumarMeses_(hoy, -CONFIG.MESES_ELIMINACION);
  var archivo = null;
  var archivadas = 0, borradas = 0;
  leer_('BancoImagenes').forEach(function (b) {
    if (b.estado === 'eliminada' || !b.fechaSubida) return;
    if (b.fechaSubida <= limiteBorrado) {
      borrarArchivos_(b);
      actualizar_('BancoImagenes', b._fila, { estado: 'eliminada' });
      borradas++;
    } else if (b.estado === 'activa' && b.fechaSubida <= limiteArchivo) {
      archivo = archivo || carpeta_(CONFIG.CARPETA_ARCHIVO);
      [b.fileId, b.thumbId].forEach(function (fid) {
        if (fid) { try { DriveApp.getFileById(fid).moveTo(archivo); } catch (e) {} }
      });
      actualizar_('BancoImagenes', b._fila, { estado: 'archivada' });
      archivadas++;
    }
  });
  return 'Archivadas: ' + archivadas + ' · Borradas: ' + borradas;
}

/* ====================== configuración inicial ====================== */

// Ejecutar UNA VEZ desde el editor. Crea hojas, marcas, la cuenta LST-SUP
// (administradora) y el disparador diario. Es seguro volver a ejecutarla.
function configurarInicial() {
  Object.keys(HOJAS).forEach(hoja_);

  var existentes = marcas_().map(function (m) { return m.codigo; });
  agregarVarios_('Marcas', MARCAS_INICIALES.filter(function (m) { return existentes.indexOf(m[0]) === -1; }).map(function (m) {
    return { codigo: m[0], numero: m[1], nombre: m[2], zonaHoraria: CONFIG.ZONA_DEFAULT, activa: m[3], esOrigen: m[3] };
  }));

  var mensaje = 'Configuración lista.';
  var origen = MARCAS_INICIALES.filter(function (m) { return m[3] === 'SI'; })[0][0];
  var admin = origen + '-SUP';
  if (!leer_('Usuarios').some(function (u) { return u.usuario === admin; })) {
    var temporal = generarPassword_();
    var salt = nuevoId_();
    agregar_('Usuarios', {
      usuario: admin, nombre: 'Supervisor ' + origen, marca: origen, rol: 'supervisor', admin: 'SI',
      hash: hash_(temporal, salt), salt: salt, temporal: 'SI', activo: 'SI',
      creado: Utilities.formatDate(new Date(), CONFIG.ZONA_DEFAULT, 'yyyy-MM-dd')
    });
    mensaje += ' Usuario: ' + admin + ' · Contraseña temporal: ' + temporal;
  }

  var tieneDisparador = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'mantenimientoDiario'; });
  if (!tieneDisparador) ScriptApp.newTrigger('mantenimientoDiario').timeBased().everyDays(1).atHour(3).create();

  Logger.log(mensaje);
  return mensaje;
}

/* ================ migración desde la versión anterior ================ */

// Si los datos anteriores están en OTRA hoja de cálculo, pegar aquí su ID
// (lo que va entre /d/ y /edit en la URL). Si están en esta misma, dejar vacío.
var ID_HOJA_ANTERIOR = '';

// Ejecutar desde el editor, después de configurarInicial. Se puede repetir
// sin duplicar datos.
// Pasa a la marca de origen: las imágenes del banco anterior (incluidas las
// ya usadas, que vuelven a rotar), el historial de imágenes del día y las
// respuestas. Las imágenes migradas cuentan sus 6 meses desde hoy.
function migrarDatosAnteriores() {
  var ss = ID_HOJA_ANTERIOR ? SpreadsheetApp.openById(ID_HOJA_ANTERIOR) : SpreadsheetApp.getActiveSpreadsheet();
  var nombresViejos = ['Banco', 'ImagenDia', 'Respuestas'];
  var encontradas = nombresViejos.filter(function (n) { return ss.getSheetByName(n); });
  if (!encontradas.length) {
    var aviso = 'No se encontraron las hojas anteriores (' + nombresViejos.join(', ') + ') en la hoja de cálculo "' +
      ss.getName() + '". Hojas que sí tiene: ' + ss.getSheets().map(function (h) { return h.getName(); }).join(', ') +
      '. Si tus datos están en otra hoja de cálculo, pega su ID en ID_HOJA_ANTERIOR.';
    Logger.log(aviso);
    return aviso;
  }
  var origen = MARCAS_INICIALES.filter(function (m) { return m[3] === 'SI'; })[0][0];
  var hoy = Utilities.formatDate(new Date(), CONFIG.ZONA_DEFAULT, 'yyyy-MM-dd');

  var leerViejo = function (nombre) {
    var sh = ss.getSheetByName(nombre);
    if (!sh) return [];
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    return data.slice(1).filter(function (r) { return r.join('') !== ''; }).map(function (r) {
      var o = {};
      headers.forEach(function (h, j) {
        var v = r[j];
        if (v instanceof Date) v = Utilities.formatDate(v, ss.getSpreadsheetTimeZone(), h === 'hora' ? 'HH:mm:ss' : 'yyyy-MM-dd');
        o[h] = v === null || v === undefined ? '' : String(v);
      });
      if (!o.fileId && o.url) {
        var m = o.url.match(/[?&](?:img|id)=([^&]+)/);
        if (m) o.fileId = decodeURIComponent(m[1]);
      }
      return o;
    });
  };

  var bancoNuevo = leer_('BancoImagenes');
  var porFileId = {};
  bancoNuevo.forEach(function (b) { porFileId[b.fileId] = b.id; });

  var nuevasImagenes = [];
  var agregarImagen = function (viejo, tipo) {
    if (!viejo.fileId || porFileId[viejo.fileId]) return;
    var id = nuevoId_();
    porFileId[viejo.fileId] = id;
    nuevasImagenes.push({
      id: id, marca: origen, compartida: 'SI', tipo: tipo,
      parasito: tipo === 'Negativa' ? NEGATIVO : viejo.parasito, fuente: tipo === 'Negativa' ? '' : viejo.fuente,
      fileId: viejo.fileId, thumbId: viejo.thumbId || '', fechaSubida: hoy, estado: 'activa'
    });
  };
  leerViejo('Banco').forEach(function (b) { agregarImagen(b, b.tipo === 'Negativa' ? 'Negativa' : 'Positiva'); });

  var dias = leerViejo('ImagenDia');
  dias.forEach(function (d) { agregarImagen(d, d.tipo === 'Negativa' ? 'Negativa' : 'Positiva'); });
  agregarVarios_('BancoImagenes', nuevasImagenes);

  // Evita duplicados si la migración se ejecuta más de una vez.
  var yaAsignado = {};
  leer_('Asignaciones').forEach(function (a) { yaAsignado[a.marca + a.fecha + a.tipo] = true; });
  var yaRegistrado = {};
  leer_('Registros').forEach(function (r) { yaRegistrado[r.id] = true; });

  // En la versión anterior la imagen 1 era la positiva y la imagen 2 la negativa.
  agregarVarios_('Asignaciones', dias.filter(function (d) {
    return d.fileId && !yaAsignado[origen + d.fecha + (d.tipo === 'Negativa' ? 'Negativa' : 'Positiva')];
  }).map(function (d) {
    var tipo = d.tipo === 'Negativa' ? 'Negativa' : 'Positiva';
    return {
      fecha: d.fecha, marca: origen, orden: tipo === 'Positiva' ? 'A' : 'B', tipo: tipo,
      idImagen: porFileId[d.fileId], fileId: d.fileId, parasito: d.parasito, fuente: d.fuente
    };
  }));

  var respuestas = leerViejo('Respuestas').filter(function (r) { return !r.id || !yaRegistrado[r.id]; });
  agregarVarios_('Registros', respuestas.map(function (r) {
    return {
      id: r.id || nuevoId_(), fecha: r.fecha, marca: origen, usuario: r.analista, hora: r.hora,
      respuesta1: r.respuesta, respuesta2: r.respuestaImg2, estado: r.estado, intento: r.intento || '1'
    };
  }));

  var mensaje = 'Migración lista (hojas encontradas: ' + encontradas.join(', ') + '). Imágenes: ' + nuevasImagenes.length + ' · Días: ' + dias.length + ' · Respuestas: ' + respuestas.length;
  Logger.log(mensaje);
  return mensaje;
}
