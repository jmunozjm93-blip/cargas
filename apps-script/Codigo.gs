/* Buzón de reportes del dashboard de Cargas (Google Apps Script).
 *
 * La página https://jmunozjm93-blip.github.io/cargas/ manda aquí cada reporte de una supervisora
 * (OC · tienda · estado · comentario · fotos). El script guarda las fotos en una carpeta de Drive
 * y agrega una fila en la hoja "Reportes" del Sheet; la página lee esa hoja para marcar qué cargas
 * ya están reportadas.
 *
 * Instalación (una sola vez, con la cuenta de Google donde quieres que quede el Sheet):
 *   1. script.google.com → Nuevo proyecto → pegar este archivo entero (reemplaza lo que haya) → Guardar.
 *   2. Arriba, elegir la función "configurar" → Ejecutar → autorizar con tu cuenta.
 *      Crea el Sheet "Cargas - Reportes" y la carpeta "Cargas - Fotos" en tu Drive (mira el Registro
 *      de ejecución: muestra los links).
 *   3. Implementar → Nueva implementación → tipo "Aplicación web":
 *      Ejecutar como: "Yo" · Quién tiene acceso: "Cualquier usuario" → Implementar.
 *   4. Copiar la "URL de la aplicación web" (termina en /exec) y pegarla en index.html, en API_URL.
 *
 * Si cambias este código, hay que hacer "Implementar → Administrar implementaciones → editar →
 * Versión: nueva" para que la URL /exec use el código nuevo.
 */

var NOMBRE_SHEET = 'Cargas - Reportes';
var NOMBRE_CARPETA = 'Cargas - Fotos';
var HOJA = 'Reportes';
var COLUMNAS = ['Fecha', 'OC', 'Cliente', 'Departamento', 'Cod tienda', 'Tienda', 'Supervisora', 'Estado', 'Comentario', 'Fotos', 'Unidades', 'Origen'];

// ---------- configuración inicial ----------
function configurar() {
  var props = PropertiesService.getScriptProperties();
  var ss;
  if (props.getProperty('SHEET_ID')) {
    ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  } else {
    ss = SpreadsheetApp.create(NOMBRE_SHEET);
    props.setProperty('SHEET_ID', ss.getId());
  }
  var hoja = ss.getSheetByName(HOJA);
  if (!hoja) {
    hoja = ss.getSheets()[0].getName() === 'Hoja 1' || ss.getSheets()[0].getName() === 'Sheet1' ? ss.getSheets()[0].setName(HOJA) : ss.insertSheet(HOJA);
  }
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(COLUMNAS);
    hoja.getRange(1, 1, 1, COLUMNAS.length).setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  var carpeta;
  if (props.getProperty('CARPETA_ID')) {
    carpeta = DriveApp.getFolderById(props.getProperty('CARPETA_ID'));
  } else {
    var it = DriveApp.getFoldersByName(NOMBRE_CARPETA);
    carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(NOMBRE_CARPETA);
    props.setProperty('CARPETA_ID', carpeta.getId());
  }
  // las fotos se ven en la página por link público; la carpeta entera queda "cualquiera con el enlace puede ver"
  carpeta.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  Logger.log('Sheet: ' + ss.getUrl());
  Logger.log('Carpeta de fotos: ' + carpeta.getUrl());
  Logger.log('Listo. Ahora: Implementar → Nueva implementación → Aplicación web.');
}

function hoja_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Falta ejecutar configurar()');
  return SpreadsheetApp.openById(id).getSheetByName(HOJA);
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
// body (JSON, como texto plano para evitar el preflight CORS):
// { oc, cliente, depto, cod, tienda, sup, estado: 'completa'|'incompleta'|'nollego', comentario, uds,
//   fotos: [{ nombre, tipo, b64 }] }
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
