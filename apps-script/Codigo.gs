/* Buzón de reportes del dashboard de Cargas (Google Apps Script).
 *
 * La página https://jmunozjm93-blip.github.io/cargas/ manda aquí cada reporte de una supervisora
 * (OC · tienda · estado · comentario · fotos). El script guarda las fotos en una carpeta de Drive
 * y agrega una fila en la hoja "Reportes" del Sheet; la página lee esa hoja para marcar qué cargas
 * ya están reportadas.
 *
 * IMPORTANTE: no borrar el Sheet "Cargas - Reportes" ni la carpeta "Cargas - Fotos" de Drive.
 * Si se borran, la página deja de poder reportar (el /exec responde 404).
 *
 * Instalación (una sola vez):
 *   1. script.google.com → Nuevo proyecto → pegar este archivo entero → Guardar.
 *   2. Elegir la función "configurar" → Ejecutar → autorizar.
 *   3. Implementar → Nueva implementación → "Aplicación web":
 *      Ejecutar como: "Yo" · Quién tiene acceso: "Cualquier usuario" → Implementar.
 *   4. Copiar la "URL de la aplicación web" (termina en /exec) y pegarla en index.html, en API_URL.
 */

var NOMBRE_SHEET = 'Cargas - Reportes';
var NOMBRE_CARPETA = 'Cargas - Fotos';
var NOMBRE_CARPETA_RAIZ = 'Cargas - Grupo Depor';   // carpeta que agrupa todo lo de este sistema
var HOJA = 'Reportes';
var COLUMNAS = ['Fecha', 'OC', 'Cliente', 'Departamento', 'Cod tienda', 'Tienda', 'Supervisora', 'Estado', 'Comentario', 'Fotos', 'Unidades', 'Origen'];

// ---------- configuración inicial ----------
function configurar() {
  var props = PropertiesService.getScriptProperties();
  var ss = null;
  var id = props.getProperty('SHEET_ID');
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) { ss = SpreadsheetApp.create(NOMBRE_SHEET); props.setProperty('SHEET_ID', ss.getId()); }
  organizar();
  var hoja = ss.getSheetByName(HOJA);
  if (!hoja) {
    var primera = ss.getSheets()[0];
    hoja = (primera.getName() === 'Hoja 1' || primera.getName() === 'Sheet1') ? primera.setName(HOJA) : ss.insertSheet(HOJA);
  }
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(COLUMNAS);
    hoja.getRange(1, 1, 1, COLUMNAS.length).setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  var carpeta = null;
  var cid = props.getProperty('CARPETA_ID');
  if (cid) { try { carpeta = DriveApp.getFolderById(cid); } catch (e) { carpeta = null; } }
  if (!carpeta) {
    var it = DriveApp.getFoldersByName(NOMBRE_CARPETA);
    carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(NOMBRE_CARPETA);
    props.setProperty('CARPETA_ID', carpeta.getId());
  }
  carpeta.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  Logger.log('Sheet: ' + ss.getUrl());
  Logger.log('Carpeta de fotos: ' + carpeta.getUrl());
  Logger.log('Listo. Ahora: Implementar → Nueva implementación → Aplicación web.');
}

// ---------- carpeta madre: agrupa el Sheet, las fotos y el propio script ----------
function raiz_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('RAIZ_ID'), f = null;
  if (id) { try { f = DriveApp.getFolderById(id); if (f.isTrashed()) f = null; } catch (e) { f = null; } }
  if (!f) {
    var it = DriveApp.getFoldersByName(NOMBRE_CARPETA_RAIZ);
    f = it.hasNext() ? it.next() : DriveApp.createFolder(NOMBRE_CARPETA_RAIZ);
    props.setProperty('RAIZ_ID', f.getId());
  }
  return f;
}

// Mueve a la carpeta madre el Sheet de reportes, la carpeta de fotos y este mismo script.
// Se puede volver a ejecutar cuando sea: si ya están dentro, no hace nada.
function organizar() {
  var props = PropertiesService.getScriptProperties();
  var raiz = raiz_();
  var movidos = [];

  var sid = props.getProperty('SHEET_ID');
  if (sid) { try { DriveApp.getFileById(sid).moveTo(raiz); movidos.push('Sheet de reportes'); } catch (e) { Logger.log('Sheet: ' + e); } }

  var cid = props.getProperty('CARPETA_ID');
  if (cid) { try { DriveApp.getFolderById(cid).moveTo(raiz); movidos.push('Carpeta de fotos'); } catch (e) { Logger.log('Fotos: ' + e); } }

  try { DriveApp.getFileById(ScriptApp.getScriptId()).moveTo(raiz); movidos.push('Script'); } catch (e) { Logger.log('Script: ' + e); }

  // carpetas sueltas de fotos que hayan quedado de versiones anteriores
  var it = DriveApp.getFoldersByName(NOMBRE_CARPETA);
  while (it.hasNext()) {
    var f = it.next();
    if (f.getId() === cid) continue;
    try { f.moveTo(raiz); movidos.push('Carpeta de fotos antigua (' + f.getId() + ')'); } catch (e) { Logger.log('Antigua: ' + e); }
  }

  Logger.log('Carpeta: ' + raiz.getUrl());
  Logger.log('Movidos: ' + (movidos.length ? movidos.join(' · ') : 'nada, ya estaba todo dentro'));
  return raiz.getUrl();
}

function hoja_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Falta ejecutar configurar()');
  var h = SpreadsheetApp.openById(id).getSheetByName(HOJA);
  if (!h) throw new Error('El Sheet no tiene la hoja ' + HOJA);
  return h;
}
function carpeta_() {
  return DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('CARPETA_ID'));
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- GET: la página pide los reportes ----------
function doGet(e) {
  try {
    var accion = (e && e.parameter && e.parameter.accion) || 'reportes';
    if (accion === 'ping') return json_({ ok: true, hora: new Date().toISOString() });
    var hoja = hoja_();
    var n = hoja.getLastRow();
    var sheetUrl = hoja.getParent().getUrl();
    var filas = n > 1 ? hoja.getRange(2, 1, n - 1, COLUMNAS.length).getValues() : [];
    var reportes = filas.map(function (r) {
      return {
        fecha: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
        oc: String(r[1]), cliente: String(r[2]), depto: String(r[3]), cod: String(r[4]), tienda: String(r[5]),
        sup: String(r[6]), estado: String(r[7]), comentario: String(r[8]),
        fotos: String(r[9] || '').split(',').map(function (s) { return s.trim(); }).filter(String),
        uds: r[10] === '' ? null : Number(r[10]),
      };
    });
    return json_({ ok: true, reportes: reportes, sheetUrl: sheetUrl, generado: new Date().toISOString() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

// ---------- POST: la página manda un reporte ----------
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var d = JSON.parse(e.postData.contents);
    if (!d.oc || !d.cod || !d.sup || !d.estado) throw new Error('Faltan datos (oc, tienda, supervisora o estado)');
    var carpeta = carpeta_();
    var ids = [];
    (d.fotos || []).slice(0, 10).forEach(function (f, i) {
      if (!f || !f.b64) return;
      var bytes = Utilities.base64Decode(f.b64);
      var ext = (f.tipo === 'image/png') ? 'png' : 'jpg';
      var nombre = d.oc + '_' + d.cod + '_' + Utilities.formatDate(new Date(), 'America/Santiago', 'yyyyMMdd-HHmm') + '_' + (i + 1) + '.' + ext;
      var blob = Utilities.newBlob(bytes, f.tipo || 'image/jpeg', nombre);
      var file = carpeta.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      ids.push(file.getId());
    });
    lock.waitLock(20000);
    hoja_().appendRow([new Date(), String(d.oc), d.cliente || '', d.depto || '', String(d.cod), d.tienda || '', d.sup, d.estado,
      d.comentario || '', ids.join(','), d.uds == null ? '' : d.uds, 'pagina']);
    return json_({ ok: true, fotos: ids, fecha: new Date().toISOString() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
