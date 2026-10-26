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
  LARGO_PASSWORD: 4,
  MAX_MB_VIDEO: 20,
  MAX_FALLOS_LOGIN: 5,
  MINUTOS_BLOQUEO: 15
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
  Usuarios: ['usuario', 'nombre', 'marca', 'rol', 'admin', 'hash', 'salt', 'temporal', 'activo', 'creado', 'iniciales', 'numero', 'puesto'],
  BancoImagenes: ['id', 'marca', 'compartida', 'tipo', 'parasito', 'fuente', 'fileId', 'thumbId', 'fechaSubida', 'estado', 'medio'],
  Asignaciones: ['fecha', 'marca', 'orden', 'tipo', 'idImagen', 'fileId', 'parasito', 'fuente', 'medio', 'thumbId'],
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

var HOJAS_LISTAS = {};  // hojas ya verificadas en esta ejecución
var LECTURAS = {};      // lecturas en memoria durante esta ejecución

// ID de la hoja de cálculo de la aplicación. Normalmente no hace falta
// llenarlo: se guarda solo la primera vez que se abre la aplicación.
var ID_HOJA = '';

// Hoja de cálculo de la aplicación. Al ejecutar funciones desde el editor
// puede no haber hoja activa; entonces se abre por su ID guardado.
function libro_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  if (ss) {
    if (props.getProperty('ID_HOJA') !== ss.getId()) props.setProperty('ID_HOJA', ss.getId());
    return ss;
  }
  var id = ID_HOJA || props.getProperty('ID_HOJA');
  if (!id) throw new Error('No se encontró la hoja de cálculo. Abre la aplicación una vez o escribe su ID en ID_HOJA (Code.gs) y vuelve a ejecutar.');
  return SpreadsheetApp.openById(id);
}

function hoja_(nombre) {
  var ss = libro_();
  var sh = ss.getSheetByName(nombre);
  var headers = HOJAS[nombre];
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else if (!HOJAS_LISTAS[nombre]) {
    // Agrega al final las columnas nuevas que una versión anterior no tenía.
    var actuales = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
    if (actuales.length < headers.length) {
      sh.getRange(1, actuales.length + 1, 1, headers.length - actuales.length).setValues([headers.slice(actuales.length)]);
    }
  }
  HOJAS_LISTAS[nombre] = true;
  return sh;
}

function texto_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.ZONA_DEFAULT, 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

// Devuelve las filas como objetos {columna: valor}, con _fila = número de fila real.
function leer_(nombre) {
  if (LECTURAS[nombre]) return LECTURAS[nombre];
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
  LECTURAS[nombre] = out;
  return out;
}

function agregar_(nombre, obj) {
  agregarVarios_(nombre, [obj]);
}

function agregarVarios_(nombre, objs) {
  if (!objs.length) return;
  delete LECTURAS[nombre];
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
  delete LECTURAS[nombre];
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
  LECTURAS = {}; // dentro del bloqueo siempre se lee lo más reciente
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

// Contraseña temporal de 4 dígitos; se cambia en el primer ingreso.
function generarPassword_() {
  var out = '';
  for (var i = 0; i < CONFIG.LARGO_PASSWORD; i++) out += Math.floor(Math.random() * 10);
  return out;
}

function nuevoId_() { return Utilities.getUuid().replace(/-/g, '').substring(0, 10); }

function cache_() { return CacheService.getScriptCache(); }

function etiquetaMarca_(m) { return m.codigo + '-' + m.numero + ' ' + m.nombre; }

function perfilDe_(u, marca) {
  return {
    usuario: u.usuario,
    nombre: u.nombre,
    iniciales: u.iniciales || '',
    puesto: u.puesto || (u.rol === 'supervisor' ? 'Supervisor' : 'Analista'),
    laboratorio: etiquetaMarca_(marca),
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
  return { success: true, token: token, perfil: perfil, datos: perfil.temporal ? null : datos_(perfil) };
}

// Al recargar la página: valida la sesión y entrega todos los datos en una sola llamada.
function getSesion(token) {
  var p = sesion_(token, undefined, true);
  return { success: true, perfil: p, datos: p.temporal ? null : datos_(p) };
}

function getDatos(token) {
  return { success: true, datos: datos_(sesion_(token)) };
}

function cerrarSesion(token) {
  if (token) cache_().remove('s_' + token);
  return { success: true };
}

function cambiarPassword(token, actual, nueva) {
  var p = sesion_(token, undefined, true);
  nueva = String(nueva || '');
  if (nueva.length !== CONFIG.LARGO_PASSWORD) return { success: false, error: 'La contraseña debe tener ' + CONFIG.LARGO_PASSWORD + ' caracteres' };
  if (nueva === actual) return { success: false, error: 'La nueva contraseña debe ser distinta a la actual' };
  var u = leer_('Usuarios').filter(function (x) { return x.usuario === p.usuario; })[0];
  if (!u || hash_(String(actual || ''), u.salt) !== u.hash) return { success: false, error: 'La contraseña actual no es correcta' };
  var salt = nuevoId_();
  actualizar_('Usuarios', u._fila, { hash: hash_(nueva, salt), salt: salt, temporal: 'NO' });
  p.temporal = false;
  guardarSesion_(token, p);
  return { success: true, perfil: p, datos: datos_(p) };
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

// Datos para mostrar una imagen o un video del banco (o de una asignación).
function medio_(x, ancho) {
  var esVideo = x.medio === 'video';
  var m = { medio: esVideo ? 'video' : 'imagen', url: urlImagen_(esVideo && x.thumbId ? x.thumbId : x.fileId, ancho) };
  if (esVideo) m.fileId = x.fileId;
  return m;
}

// El video se entrega como datos (no como enlace de Drive) para que el
// navegador lo reproduzca de inmediato, en bucle y sin opción de descarga.
function getVideo(token, fileId) {
  var p = sesion_(token);
  var fila = leer_('BancoImagenes').filter(function (b) {
    return b.fileId === fileId && b.medio === 'video' && (b.marca === p.marca || b.compartida === 'SI');
  })[0];
  if (!fila) return { success: false, error: 'Video no disponible' };
  var blob = DriveApp.getFileById(fileId).getBlob();
  return { success: true, mime: blob.getContentType() || 'video/mp4', data: Utilities.base64Encode(blob.getBytes()) };
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
      fileId: elegida.fileId, parasito: elegida.parasito, fuente: elegida.fuente,
      medio: elegida.medio || 'imagen', thumbId: elegida.thumbId || ''
    });
  });
  agregarVarios_('Asignaciones', nuevas);
  return nuevas;
}

/* ========================= control diario ========================= */

function registrosHoy_(perfil) {
  var fecha = hoy_(perfil);
  return registrosMarca_(perfil.marca).filter(function (r) { return r.fecha === fecha; });
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

// El analista nunca recibe el parásito correcto, el tipo de imagen ni la fuente.
function controlHoy_(p) {
  var asignacion = asignacionHoy_(p, true);
  var regs = registrosHoy_(p);
  var out = { hay: false, estado: estadoDeRegistros_(regs), fecha: hoy_(p) };
  // Respuestas ya enviadas hoy, para consultarlas sin poder editarlas.
  out.respuestas = regs.map(function (r) {
    return { intento: r.intento, hora: r.hora, usuario: r.usuario, respuesta1: r.respuesta1, respuesta2: r.respuesta2, estado: r.estado };
  });
  if (!asignacion) return out;
  var esSup = p.rol === 'supervisor';
  out.hay = true;
  out.imagenes = ['A', 'B'].map(function (orden) {
    var a = asignacion[orden];
    var img = medio_(a, 2400);
    if (esSup) { img.parasito = a.parasito; img.tipo = a.tipo; img.fuente = a.fuente; }
    return img;
  });
  return out;
}

function getControlHoy(token) {
  var r = controlHoy_(sesion_(token));
  r.success = true;
  return r;
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

/* ==================== datos para la aplicación ==================== */

// Todo lo que la aplicación necesita, en una sola llamada. El navegador lo
// guarda en memoria y así el cambio de pestaña es inmediato.
function datos_(p) {
  var d = { hoy: controlHoy_(p), indicadores: indicadores_(p), historial: historial_(p), atlas: atlas_(p), especiesBanco: especiesBanco_(p), estadisticas: estadisticas_(p) };
  if (p.rol === 'supervisor') {
    d.banco = { Positiva: banco_(p, 'Positiva'), Negativa: banco_(p, 'Negativa') };
    d.usuarios = usuarios_(p);
    if (p.admin) d.marcas = marcasLista_();
  }
  return d;
}

function asignacionesPorFecha_(marca) {
  var porFecha = {};
  leer_('Asignaciones').forEach(function (a) {
    if (a.marca !== marca) return;
    porFecha[a.fecha] = porFecha[a.fecha] || {};
    porFecha[a.fecha][a.orden] = a;
  });
  return porFecha;
}

// Un registro sin fecha o sin resultado (por ejemplo, un renglón incompleto
// copiado de la versión anterior) no se toma en cuenta.
function registroValido_(r) {
  return /^\d{4}-\d{2}-\d{2}$/.test(r.fecha) && (r.estado === 'Aprobado' || r.estado === 'Rechazado');
}

function registrosMarca_(marca) {
  return leer_('Registros').filter(function (r) { return r.marca === marca && registroValido_(r); });
}

function registrosPorFecha_(marca) {
  var porFecha = {};
  registrosMarca_(marca).forEach(function (r) {
    (porFecha[r.fecha] = porFecha[r.fecha] || []).push(r);
  });
  return porFecha;
}

/* ========================= historial ========================= */

function historial_(p) {
  var hoy = hoy_(p);
  var esSup = p.rol === 'supervisor';
  var porFecha = asignacionesPorFecha_(p.marca);
  var registros = registrosMarca_(p.marca);
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
  return filas;
}

function getHistorial(token) {
  return { success: true, registros: historial_(sesion_(token)) };
}

// Indicadores de los últimos 30 días (desde que la marca empezó a usar el sistema).
function indicadores_(p) {
  var hoy = hoy_(p);
  var desde = sumarDias_(hoy, -29);
  var asignadas = leer_('Asignaciones').filter(function (a) { return a.marca === p.marca; });
  var inicio = asignadas.reduce(function (min, a) { return a.fecha < min ? a.fecha : min; }, hoy);
  if (inicio > desde) desde = inicio;
  var porDia = registrosPorFecha_(p.marca);

  var res = { dias: 0, primerIntento: 0, aprobadoReintento: 0, noAprobado: 0, sinRespuesta: 0, desde: desde, ultimo: null };
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
  var pct = function (n) { return res.dias ? Math.round(n * 100 / res.dias) : 0; };
  res.concordancia = pct(res.primerIntento + res.aprobadoReintento);
  res.porcentajePrimerIntento = pct(res.primerIntento);
  res.cumplimiento = pct(res.dias - res.sinRespuesta);
  res.controlesAprobados = res.primerIntento + res.aprobadoReintento;
  res.controlesRealizados = res.dias - res.sinRespuesta;

  var fechas = Object.keys(porDia).sort();
  for (var i = fechas.length - 1; i >= 0; i--) {
    var dia = porDia[fechas[i]];
    var aprob = dia.filter(function (r) { return r.estado === 'Aprobado'; })[0];
    var ult = aprob || dia[dia.length - 1];
    if (!aprob && dia.length < CONFIG.MAX_INTENTOS && fechas[i] === hoy) continue; // aún en curso
    res.ultimo = { fecha: fechas[i], estado: aprob ? 'Aprobado' : 'No aprobado', usuario: ult.usuario, hora: ult.hora, intento: ult.intento };
    break;
  }
  return res;
}

function getIndicadores(token) {
  return { success: true, indicadores: indicadores_(sesion_(token)) };
}

/* ========================= estadísticas ========================= */

function inicioDeMes_(fecha, meses) { return sumarMeses_(fecha.substring(0, 8) + '01', meses || 0); }
function finDeMes_(mes) { return sumarDias_(sumarMeses_(mes + '-01', 1), -1); }

// Días con control entre dos fechas, con sus imágenes y sus respuestas en
// orden de intento. Hoy solo cuenta cuando el control ya terminó.
function diasControl_(p, desde, hasta) {
  var hoy = hoy_(p);
  var asignaciones = asignacionesPorFecha_(p.marca);
  var registros = registrosPorFecha_(p.marca);
  var dias = [];
  Object.keys(asignaciones).sort().forEach(function (f) {
    var a = asignaciones[f];
    if (f < desde || f > hasta || f > hoy || !a.A || !a.B) return;
    var regs = (registros[f] || []).slice().sort(function (x, y) { return Number(x.intento) - Number(y.intento); });
    var e = estadoDeRegistros_(regs);
    if (f === hoy && !e.aprobado && !e.agotado) return;
    dias.push({ fecha: f, A: a.A, B: a.B, regs: regs, aprobado: e.aprobado });
  });
  return dias;
}

function pct_(n, d) { return d ? Math.round(n * 100 / d) : null; }

// Indicadores de calidad de un conjunto de días.
// concordancia: muestras identificadas bien en el primer intento / muestras evaluadas.
// aprobacionDirecta: controles aprobados en el primer intento / controles realizados.
// cumplimiento: controles realizados / días con control.
function metricas_(dias) {
  var m = { controles: dias.length, realizados: 0, aprobados: 0, directos: 0, muestras: 0, correctas: 0, intentos: 0,
    errores: { falsoNegativo: 0, falsoPositivo: 0, especie: 0 } };
  dias.forEach(function (d) {
    if (!d.regs.length) return;
    m.realizados++;
    if (d.aprobado) m.aprobados++;
    if (d.regs[0].estado === 'Aprobado') m.directos++;
    m.intentos += d.regs.length;
    d.regs.forEach(function (r, i) {
      [[d.A, r.respuesta1], [d.B, r.respuesta2]].forEach(function (par) {
        var bien = normalizarRespuesta_(par[1]) === normalizarRespuesta_(par[0].parasito);
        if (i === 0) { m.muestras++; if (bien) m.correctas++; }
        if (bien) return;
        if (par[0].tipo === 'Negativa') m.errores.falsoPositivo++;
        else if (normalizarRespuesta_(par[1]) === normalizarRespuesta_(NEGATIVO)) m.errores.falsoNegativo++;
        else m.errores.especie++;
      });
    });
  });
  m.noAprobados = m.realizados - m.aprobados;
  m.concordancia = pct_(m.correctas, m.muestras);
  m.aprobacionDirecta = pct_(m.directos, m.realizados);
  m.cumplimiento = pct_(m.realizados, m.controles);
  m.intentosPromedio = m.realizados ? Math.round(m.intentos * 10 / m.realizados) / 10 : null;
  return m;
}

// Parásitos con menos aciertos en el primer intento dentro del periodo.
function especiesDificiles_(dias) {
  var mapa = {};
  dias.forEach(function (d) {
    if (!d.regs.length) return;
    [[d.A, d.regs[0].respuesta1], [d.B, d.regs[0].respuesta2]].forEach(function (par) {
      if (par[0].tipo !== 'Positiva') return;
      var e = mapa[par[0].parasito] = mapa[par[0].parasito] || { nombre: par[0].parasito, veces: 0, correctas: 0 };
      e.veces++;
      if (normalizarRespuesta_(par[1]) === normalizarRespuesta_(par[0].parasito)) e.correctas++;
    });
  });
  return Object.keys(mapa).map(function (k) { return mapa[k]; })
    .sort(function (a, b) { return (a.correctas / a.veces - b.correctas / b.veces) || (b.veces - a.veces); })
    .slice(0, 5);
}

// Concordancia de un solo día (primer intento de sus dos muestras).
function concordanciaDia_(d) {
  if (!d.regs.length) return null;
  var bien = 0;
  [[d.A, d.regs[0].respuesta1], [d.B, d.regs[0].respuesta2]].forEach(function (par) {
    if (normalizarRespuesta_(par[1]) === normalizarRespuesta_(par[0].parasito)) bien++;
  });
  return bien * 50;
}

// Resultado final de un día: el intento aprobado o, si no hubo, el último.
function resultadoFinal_(d) {
  if (!d.regs.length) return null;
  return d.regs.filter(function (r) { return r.estado === 'Aprobado'; })[0] || d.regs[d.regs.length - 1];
}

// Estadísticas por mes de calendario (solo los meses con controles), con el
// detalle de cada día para la vista diaria.
function estadisticas_(p) {
  var hoy = hoy_(p);
  var todos = diasControl_(p, inicioDeMes_(hoy, -(CONFIG.MESES_ELIMINACION - 1)), hoy);
  var porMes = {};
  todos.forEach(function (d) { (porMes[d.fecha.substring(0, 7)] = porMes[d.fecha.substring(0, 7)] || []).push(d); });
  var claves = Object.keys(porMes).sort();
  var out = { mesActual: hoy.substring(0, 7), meses: [], serieMensual: [] };
  claves.forEach(function (mes) {
    var dias = porMes[mes];
    var anterior = porMes[inicioDeMes_(mes + '-01', -1).substring(0, 7)] || [];
    var m = {
      mes: mes, actual: metricas_(dias), anterior: metricas_(anterior), especies: especiesDificiles_(dias),
      dias: dias.map(function (d) {
        var f = resultadoFinal_(d);
        return { dia: Number(d.fecha.substring(8, 10)), concordancia: concordanciaDia_(d), resultado: f ? (f.estado === 'Aprobado' ? 'Aprobado' : 'Rechazado') : 'Sin respuesta',
          intento: f ? Number(f.intento) : null, analista: f ? f.usuario : '' };
      })
    };
    if (p.rol === 'supervisor') {
      var porUsuario = {};
      dias.forEach(function (d) { if (d.regs.length) (porUsuario[d.regs[0].usuario] = porUsuario[d.regs[0].usuario] || []).push(d); });
      m.porAnalista = Object.keys(porUsuario).sort().map(function (u) {
        var x = metricas_(porUsuario[u]);
        return { usuario: u, controles: x.realizados, concordancia: x.concordancia, aprobacionDirecta: x.aprobacionDirecta, noAprobados: x.noAprobados };
      });
    }
    out.meses.push(m);
    if (m.actual.concordancia != null) out.serieMensual.push({ mes: mes, valor: m.actual.concordancia });
  });
  out.meses.reverse();
  return out;
}

var MESES_NOMBRE = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

function etiquetaMes_(mes) { return MESES_NOMBRE[Number(mes.substring(5, 7)) - 1] + ' ' + mes.substring(0, 4); }

// Un renglón por día del mes con el resultado final (el intento aprobado o el
// último), más la concordancia acumulada para la gráfica.
function filasExportacion_(p, mes) {
  var hoy = hoy_(p);
  var dias = diasControl_(p, mes + '-01', finDeMes_(mes) < hoy ? finDeMes_(mes) : hoy);
  var muestras = 0, correctas = 0;
  var filas = dias.map(function (d) {
    var dia = Number(d.fecha.substring(8, 10));
    var final = resultadoFinal_(d);
    if (!final) return { dia: dia, hora: '', a: '', b: '', intento: '', analista: '', resultado: 'Sin respuesta', acumulada: muestras ? Math.round(correctas * 100 / muestras) : null };
    var c = concordanciaDia_(d);
    muestras += 2; correctas += c / 50;
    return {
      dia: dia, hora: String(final.hora).substring(0, 5), a: final.respuesta1, b: final.respuesta2, intento: Number(final.intento),
      analista: final.usuario, resultado: final.estado === 'Aprobado' ? 'Aprobado' : 'Rechazado',
      acumulada: Math.round(correctas * 100 / muestras)
    };
  });
  return { etiqueta: etiquetaMes_(mes), filas: filas };
}

// Crea una hoja con formato y gráfica, la descarga como Excel y la borra.
function exportarExcel(token, mes) {
  var p = sesion_(token);
  if (!/^\d{4}-\d{2}$/.test(String(mes))) mes = hoy_(p).substring(0, 7);
  var datos = filasExportacion_(p, mes);
  var libro = SpreadsheetApp.create('Control de calidad ' + p.marca + ' ' + datos.etiqueta);
  try {
    var hoja = libro.getSheets()[0];
    hoja.setName('Control diario');
    var n = datos.filas.length;
    hoja.getRange('A1:H2').setValues([
      ['Mes:', datos.etiqueta, '', '', 'Área:', 'Parasitología', '', ''],
      ['Laboratorio:', p.laboratorio, '', '', 'Exportado por:', p.usuario, '', '']
    ]);
    hoja.getRange('B1:D1').merge(); hoja.getRange('B2:D2').merge();
    hoja.getRange('F1:H1').merge(); hoja.getRange('F2:H2').merge();
    hoja.getRange('A1:A2').setFontWeight('bold'); hoja.getRange('E1:E2').setFontWeight('bold');
    hoja.getRange('B1:D2').setBorder(null, null, true, null, null, null);
    hoja.getRange('F1:H2').setBorder(null, null, true, null, null, null);

    var cab = ['No.', 'Día', 'Hora', 'Muestra A', 'Muestra B', 'Intento', 'Analista', 'Resultado'];
    hoja.getRange(4, 1, 1, 8).setValues([cab]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#312e81');
    if (n) {
      hoja.getRange(5, 1, n, 8).setValues(datos.filas.map(function (f, i) { return [i + 1, f.dia, f.hora, f.a, f.b, f.intento, f.analista, f.resultado]; }));
      hoja.getRange(5, 2, n, 1).setBackground('#e5e7eb');
      var res = hoja.getRange(5, 8, n, 1);
      hoja.setConditionalFormatRules([
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Aprobado').setBackground('#d1fae5').setFontColor('#065f46').setRanges([res]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Rechazado').setBackground('#fee2e2').setFontColor('#991b1b').setRanges([res]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Sin respuesta').setBackground('#f3f4f6').setFontColor('#6b7280').setRanges([res]).build()
      ]);
    }
    var tabla = hoja.getRange(4, 1, Math.max(n, 1) + 1, 8);
    tabla.setBorder(true, true, true, true, true, true, '#9ca3af', SpreadsheetApp.BorderStyle.SOLID);
    hoja.getRange(1, 1, Math.max(n + 4, 4), 8).setFontFamily('Arial').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    [44, 84, 54, 190, 190, 56, 76, 92].forEach(function (w, i) { hoja.setColumnWidth(i + 1, w); });

    // Datos de la gráfica en una hoja aparte; la gráfica queda a la derecha de la tabla.
    var g = libro.insertSheet('Datos de la gráfica');
    var puntos = datos.filas.filter(function (f) { return f.acumulada != null; });
    g.getRange(1, 1, 1, 3).setValues([['Día', 'Concordancia acumulada (%)', 'Meta (%)']]);
    if (puntos.length) g.getRange(2, 1, puntos.length, 3).setValues(puntos.map(function (f) { return [String(f.dia), f.acumulada, 90]; }));
    g.getRange(1, 1, Math.max(puntos.length, 1) + 1, 3).setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center');
    if (puntos.length) {
      var grafica = hoja.newChart().asLineChart()
        .addRange(g.getRange(1, 1, puntos.length + 1, 3))
        .setTitle('Concordancia acumulada del mes')
        .setOption('vAxis', { viewWindow: { min: 60, max: 150 }, title: 'Concordancia (%)' })
        .setOption('hAxis', { title: 'Día' })
        .setOption('series', {
          0: { color: '#4f46e5', pointSize: 7, lineWidth: 2 },
          1: { color: '#dc2626', pointSize: 0, lineWidth: 2, lineDashStyle: [8, 5] }
        })
        .setOption('legend', { position: 'bottom' })
        .setOption('width', 560).setOption('height', 320)
        .setPosition(4, 10, 0, 0)
        .build();
      hoja.insertChart(grafica);
    }
    SpreadsheetApp.flush();
    var url = 'https://docs.google.com/spreadsheets/d/' + libro.getId() + '/export?format=xlsx';
    var archivo = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
    return {
      success: true, nombre: 'Control de calidad ' + p.marca + ' ' + datos.etiqueta.replace(/\//g, '-') + '.xlsx',
      data: Utilities.base64Encode(archivo.getBytes())
    };
  } finally {
    DriveApp.getFileById(libro.getId()).setTrashed(true);
  }
}

/* ========================= atlas ========================= */

function grupoParasito_(nombre) {
  return /huevo|larva|progl/i.test(nombre) ? 'Helminto' : 'Protozoario';
}

// Parásitos que han aparecido en el control de la marca y cuántas veces se
// identificaron bien en el primer intento. El día de hoy solo cuenta cuando
// el control ya terminó, para no revelar la respuesta.
function atlas_(p) {
  var hoy = hoy_(p);
  var asignaciones = asignacionesPorFecha_(p.marca);
  var registros = registrosPorFecha_(p.marca);
  var mapa = {};
  Object.keys(asignaciones).forEach(function (fecha) {
    var regs = (registros[fecha] || []).slice().sort(function (x, y) { return Number(x.intento) - Number(y.intento); });
    if (!regs.length) return;
    var e = estadoDeRegistros_(regs);
    if (fecha === hoy && !e.aprobado && !e.agotado) return;
    var primero = regs[0];
    var final = regs.filter(function (r) { return r.estado === 'Aprobado'; })[0] || regs[regs.length - 1];
    ['A', 'B'].forEach(function (orden) {
      var a = asignaciones[fecha][orden];
      if (!a || a.tipo !== 'Positiva') return;
      var campo = orden === 'A' ? 'respuesta1' : 'respuesta2';
      var esperada = normalizarRespuesta_(a.parasito);
      var alPrimero = normalizarRespuesta_(primero[campo]) === esperada;
      var alFinal = normalizarRespuesta_(final[campo]) === esperada;
      var item = mapa[a.parasito] = mapa[a.parasito] || { nombre: a.parasito, grupo: grupoParasito_(a.parasito), veces: 0, correctas: 0, ultima: '', url: '', hist: [] };
      item.veces++;
      if (alFinal) item.correctas++;
      // 2: bien a la primera, 1: bien en un reintento, 0: no se identificó.
      item.hist.push([fecha, alPrimero ? 2 : alFinal ? 1 : 0]);
      if (fecha > item.ultima) {
        item.ultima = fecha;
        var m = medio_(a, 1600);
        item.url = m.url; item.medio = m.medio; item.fileId = m.fileId || '';
      }
    });
  });
  return Object.keys(mapa).map(function (k) {
    var it = mapa[k];
    it.hist = it.hist.sort(function (x, y) { return x[0] > y[0] ? 1 : -1; }).map(function (x) { return x[1]; }).slice(-12);
    return it;
  }).sort(function (a, b) { return b.ultima > a.ultima ? 1 : -1; });
}

function especiesBanco_(p) {
  var vistos = {};
  poolMarca_(leer_('BancoImagenes'), p, 'Positiva').forEach(function (b) { if (b.parasito) vistos[b.parasito] = true; });
  return Object.keys(vistos).sort();
}

/* ======================== banco de imágenes ======================== */

function subirImagenBanco(token, base64, mimeType, parasito, fuente, tipo, thumbBase64) {
  var p = sesion_(token, 'supervisor');
  tipo = tipo === 'Negativa' ? 'Negativa' : 'Positiva';
  parasito = String(parasito || '').trim();
  if (!base64) return { success: false, error: 'Falta la imagen' };
  var esVideo = /^video\//.test(String(mimeType || ''));
  if (esVideo && base64.length > CONFIG.MAX_MB_VIDEO * 1.4 * 1024 * 1024) return { success: false, error: 'El video pesa más de ' + CONFIG.MAX_MB_VIDEO + ' MB' };
  if (tipo === 'Positiva' && !parasito) return { success: false, error: 'Falta el nombre del parásito' };

  var id = nuevoId_();
  var folder = carpeta_(CONFIG.CARPETA_IMAGENES);
  var extension = esVideo ? '.' + (String(mimeType).split('/')[1] || 'mp4').replace('quicktime', 'mov') : '.jpg';
  var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(base64.split(',').pop()), mimeType || 'image/jpeg', id + extension));
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
    fileId: file.getId(), thumbId: thumbId, fechaSubida: hoy_(p), estado: 'activa', medio: esVideo ? 'video' : 'imagen'
  });
  return { success: true, id: id };
}

function banco_(p, tipo) {
  var hoy = hoy_(p);
  var banco = leer_('BancoImagenes');
  var propias = banco.filter(function (b) { return b.marca === p.marca && b.estado === 'activa' && b.tipo === tipo; });
  propias.sort(function (a, b) { return b.fechaSubida > a.fechaSubida ? 1 : -1; });

  var pool = poolMarca_(banco, p, tipo);
  var vistas = ultimaVez_(leer_('Asignaciones'), p.marca);
  var limite = sumarDias_(hoy, -CONFIG.DIAS_SIN_REPETIR);
  return {
    items: propias.map(function (b) {
      return {
        id: b.id, parasito: b.parasito, fuente: b.fuente, fechaSubida: b.fechaSubida,
        diasEnRotacion: Math.max(0, diasEntre_(hoy, sumarMeses_(b.fechaSubida, CONFIG.MESES_ROTACION))),
        url: urlImagen_(b.thumbId || b.fileId, 200),
        urlGrande: urlImagen_(b.fileId, 2400), fileId: b.fileId,
        medio: b.medio === 'video' ? 'video' : 'imagen'
      };
    }),
    resumen: {
      propias: propias.length,
      general: p.esOrigen ? 0 : pool.length - propias.length,
      totalRotacion: pool.length,
      disponiblesSinRepetir: pool.filter(function (b) { return !vistas[b.id] || vistas[b.id] <= limite; }).length,
      recomendado: CONFIG.DIAS_SIN_REPETIR
    }
  };
}

function eliminarImagenBanco(token, id) {
  var p = sesion_(token, 'supervisor');
  var fila = leer_('BancoImagenes').filter(function (b) { return b.id === id && b.marca === p.marca && b.estado !== 'eliminada'; })[0];
  if (!fila) return { success: false, error: 'Imagen no encontrada' };
  return conBloqueo_(function () {
    borrarArchivos_(fila);
    actualizar_('BancoImagenes', fila._fila, { estado: 'eliminada' });
    var reemplazadas = reemplazarEnHoy_(fila);
    return { success: true, reemplazadas: reemplazadas };
  });
}

// Si la imagen eliminada estaba asignada hoy a algún laboratorio, se sustituye
// por otra del mismo tipo, aunque el control ya se esté respondiendo. Los
// controles que ya terminaron (aprobados o con intentos agotados) no se tocan,
// para conservar el registro de lo que se evaluó.
function reemplazarEnHoy_(eliminada) {
  var marcas = {};
  marcas_().forEach(function (m) { marcas[m.codigo] = m; });
  var asignaciones = leer_('Asignaciones');
  var cambios = 0;
  asignaciones.filter(function (a) { return a.idImagen === eliminada.id; }).forEach(function (a) {
    var m = marcas[a.marca];
    if (!m) return;
    var perfil = { marca: m.codigo, esOrigen: m.esOrigen === 'SI', zona: m.zonaHoraria || CONFIG.ZONA_DEFAULT };
    if (a.fecha !== hoy_(perfil)) return;
    var estado = estadoDeRegistros_(registrosHoy_(perfil));
    if (estado.aprobado || estado.agotado) return;

    var usadasHoy = asignaciones.filter(function (x) { return x.marca === a.marca && x.fecha === a.fecha; }).map(function (x) { return x.idImagen; });
    var pool = poolMarca_(leer_('BancoImagenes'), perfil, a.tipo).filter(function (b) { return usadasHoy.indexOf(b.id) === -1; });
    var nueva = elegirImagen_(pool, ultimaVez_(asignaciones, a.marca));
    if (!nueva) return;
    actualizar_('Asignaciones', a._fila, { idImagen: nueva.id, fileId: nueva.fileId, parasito: nueva.parasito, fuente: nueva.fuente, medio: nueva.medio || 'imagen', thumbId: nueva.thumbId || '' });
    cache_().remove('a_' + a.marca + '_' + a.fecha);
    cambios++;
  });
  return cambios;
}

function borrarArchivos_(fila) {
  [fila.fileId, fila.thumbId].forEach(function (fid) {
    if (fid) { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {} }
  });
}

/* ========================= usuarios ========================= */

function usuarios_(p) {
  var marcas = {};
  marcas_().forEach(function (m) { marcas[m.codigo] = etiquetaMarca_(m); });
  // Las cuentas eliminadas no se listan; su fila se conserva para el historial.
  var lista = leer_('Usuarios').filter(function (u) { return u.activo === 'SI' && (p.admin || u.marca === p.marca); }).map(function (u) {
    return {
      usuario: u.usuario, nombre: u.nombre, iniciales: u.iniciales, numero: u.numero, puesto: u.puesto,
      marca: u.marca, laboratorio: marcas[u.marca] || u.marca, rol: u.rol, admin: u.admin === 'SI',
      activo: u.activo === 'SI', temporal: u.temporal === 'SI', creado: u.creado
    };
  });
  lista.sort(function (a, b) { return (a.marca + a.usuario) > (b.marca + b.usuario) ? 1 : -1; });
  return lista;
}

function puedeGestionar_(p, u) {
  if (u.usuario === p.usuario) return false;
  if (p.admin) return true;
  return u.marca === p.marca && u.rol === 'analista';
}

// Valida los datos de una cuenta y arma su usuario: iniciales + número de
// analista (ej. MPWN1) o, para supervisores, iniciales + "-SUP" (ej. MPWN-SUP).
function validarDatosUsuario_(datos, rol) {
  var d = {
    nombre: String(datos.nombre || '').trim(),
    iniciales: String(datos.iniciales || '').toUpperCase().replace(/\s+/g, ''),
    numero: String(datos.numero || '').trim(),
    puesto: String(datos.puesto || '').trim()
  };
  if (!d.nombre) return { error: 'Escribe el nombre completo' };
  if (!/^[A-ZÑ]{2,6}$/.test(d.iniciales)) return { error: 'Las iniciales deben ser de 2 a 6 letras' };
  if (rol === 'supervisor') {
    d.numero = '';
    d.usuario = d.iniciales + '-SUP';
  } else {
    if (!/^\d{1,3}$/.test(d.numero)) return { error: 'El número de analista debe ser de 1 a 3 dígitos' };
    d.numero = String(Number(d.numero));
    d.usuario = d.iniciales + d.numero;
  }
  if (!d.puesto) d.puesto = rol === 'supervisor' ? 'Supervisor' : 'Analista';
  return d;
}

// datos: { nombre, iniciales, numero, puesto, rol, marca }
function crearUsuario(token, datos) {
  var p = sesion_(token, 'supervisor');
  datos = datos || {};
  var rol = datos.rol === 'supervisor' ? 'supervisor' : 'analista';
  var marca = p.admin ? String(datos.marca || p.marca) : p.marca;
  if (!p.admin && rol !== 'analista') return { success: false, error: 'No autorizado' };
  if (!marca_(marca)) return { success: false, error: 'Laboratorio no válido' };
  var d = validarDatosUsuario_(datos, rol);
  if (d.error) return { success: false, error: d.error };

  return conBloqueo_(function () {
    if (leer_('Usuarios').some(function (u) { return u.usuario === d.usuario; })) {
      return { success: false, error: 'El usuario ' + d.usuario + ' ya existe.' };
    }
    var temporal = generarPassword_();
    var salt = nuevoId_();
    agregar_('Usuarios', {
      usuario: d.usuario, nombre: d.nombre, marca: marca, rol: rol, admin: 'NO', hash: hash_(temporal, salt),
      salt: salt, temporal: 'SI', activo: 'SI', creado: hoy_(p),
      iniciales: d.iniciales, numero: d.numero, puesto: d.puesto
    });
    return { success: true, usuario: d.usuario, password: temporal };
  });
}

// Un supervisor edita sus propios datos o los de sus analistas; solo la
// administración cambia el laboratorio. Si cambian las iniciales o el número,
// cambia el usuario para iniciar sesión (los registros pasados conservan el anterior).
function editarUsuario(token, usuario, datos) {
  var p = sesion_(token, 'supervisor');
  datos = datos || {};
  return conBloqueo_(function () {
    var usuarios = leer_('Usuarios');
    var u = usuarios.filter(function (x) { return x.usuario === usuario && x.activo === 'SI'; })[0];
    if (!u || (u.usuario !== p.usuario && !puedeGestionar_(p, u))) return { success: false, error: 'No autorizado' };
    var d = validarDatosUsuario_(datos, u.rol);
    if (d.error) return { success: false, error: d.error };
    var marca = p.admin && datos.marca ? String(datos.marca) : u.marca;
    if (!marca_(marca)) return { success: false, error: 'Laboratorio no válido' };
    if (d.usuario !== u.usuario && usuarios.some(function (x) { return x.usuario === d.usuario; })) {
      return { success: false, error: 'El usuario ' + d.usuario + ' ya existe.' };
    }
    actualizar_('Usuarios', u._fila, { usuario: d.usuario, nombre: d.nombre, iniciales: d.iniciales, numero: d.numero, puesto: d.puesto, marca: marca });

    var res = { success: true, usuario: d.usuario };
    if (u.usuario === p.usuario) {
      var nuevo = leer_('Usuarios').filter(function (x) { return x.usuario === d.usuario; })[0];
      var perfil = perfilDe_(nuevo, marca_(nuevo.marca));
      guardarSesion_(token, perfil);
      res.perfil = perfil;
    }
    return res;
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

function marcasLista_() {
  var usuarios = leer_('Usuarios');
  return marcas_().map(function (m) {
    return {
      codigo: m.codigo, numero: m.numero, nombre: m.nombre, etiqueta: etiquetaMarca_(m),
      activa: m.activa === 'SI', esOrigen: m.esOrigen === 'SI',
      supervisores: usuarios.filter(function (u) { return u.marca === m.codigo && u.rol === 'supervisor'; }).map(function (u) { return u.usuario; }),
      usuarios: usuarios.filter(function (u) { return u.marca === m.codigo && u.activo === 'SI'; }).length
    };
  });
}

function getMarcas(token) {
  sesion_(token, 'admin');
  return { success: true, marcas: marcasLista_() };
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

/* ============ unificación de cuentas de la versión anterior ============ */

// Antes de esta versión algunas personas respondían con su usuario sin número.
// Ejecutar desde el editor para que sus respuestas pasen a la cuenta actual.
var CUENTAS_ANTERIORES = { 'EMAR': 'EMAR1', 'MPWN': 'MPWN1' };

function unificarUsuariosAnteriores() {
  var cambios = 0;
  conBloqueo_(function () {
    leer_('Registros').forEach(function (r) {
      var nuevo = CUENTAS_ANTERIORES[String(r.usuario).trim().toUpperCase()];
      if (nuevo && r.usuario !== nuevo) { actualizar_('Registros', r._fila, { usuario: nuevo }); cambios++; }
    });
  });
  var mensaje = 'Respuestas actualizadas: ' + cambios;
  Logger.log(mensaje);
  return mensaje;
}

/* ================== cambio de nombre en el catálogo ================== */

// Si se cambia el nombre de un parásito en el catálogo (Index), las imágenes
// que ya estaban en el banco conservan el nombre anterior y nadie podría
// acertarlas. Ejecutar desde el editor después de cambiar el catálogo.
// Actualiza el banco, las imágenes asignadas y las respuestas guardadas, para
// que el atlas y el historial muestren el nombre nuevo. Se puede repetir.
function actualizarNombreGiardia() {
  return renombrarParasitos_({
    'Giardia spp quiste': 'Giardia intestinalis quiste',
    'Giardia spp trofozoíto': 'Giardia intestinalis trofozoíto'
  });
}

function renombrarParasitos_(cambios) {
  var cambiar = function (texto) {
    var partes = String(texto || '').split(' + ');
    if (!partes.some(function (x) { return cambios[x]; })) return null;
    return partes.map(function (x) { return cambios[x] || x; }).sort().join(' + ');
  };
  return conBloqueo_(function () {
    var cuenta = { BancoImagenes: 0, Asignaciones: 0, Registros: 0 };
    [['BancoImagenes', ['parasito']], ['Asignaciones', ['parasito']], ['Registros', ['respuesta1', 'respuesta2']]].forEach(function (h) {
      leer_(h[0]).forEach(function (fila) {
        var nuevos = {};
        h[1].forEach(function (col) { var v = cambiar(fila[col]); if (v !== null) nuevos[col] = v; });
        if (!Object.keys(nuevos).length) return;
        actualizar_(h[0], fila._fila, nuevos);
        if (h[0] === 'Asignaciones') cache_().remove('a_' + fila.marca + '_' + fila.fecha);
        cuenta[h[0]]++;
      });
    });
    var mensaje = 'Imágenes del banco: ' + cuenta.BancoImagenes + ' · Imágenes asignadas: ' + cuenta.Asignaciones + ' · Respuestas: ' + cuenta.Registros;
    Logger.log(mensaje);
    return mensaje;
  });
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
  var ss = ID_HOJA_ANTERIOR ? SpreadsheetApp.openById(ID_HOJA_ANTERIOR) : libro_();
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

  var respuestas = leerViejo('Respuestas').filter(function (r) {
    return r.fecha && r.estado && (!r.id || !yaRegistrado[r.id]);
  });
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
