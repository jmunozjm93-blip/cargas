/* Web Worker: lee el Excel y lo convierte sin congelar la pantalla */
importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'convertir.js');

// maestro de supervisoras publicado (data/supervisores.json): se pide una vez y se reutiliza
var pMaestro = null;
function maestro() {
  if (!pMaestro) pMaestro = fetch('../data/supervisores.json?t=' + Date.now())
    .then(function (r) { return r.ok ? r.json() : null; }).catch(function () { return null; });
  return pMaestro;
}

self.onmessage = function (e) {
  const { id, nombre, buffer } = e.data;
  const t0 = Date.now();
  try {
    const tipo = Convertir.identificar(nombre);
    self.postMessage({ id, etapa: 'leyendo' });
    const wb = XLSX.read(buffer, Convertir.opcionesLectura(tipo));
    const tLectura = Date.now() - t0;
    self.postMessage({ id, etapa: 'convirtiendo', hojas: wb.SheetNames });
    if (tipo === 'imagenes') {
      const json = Convertir.parseImagenes(wb);
      self.postMessage({ id, etapa: 'listo', tipo, texto: JSON.stringify(json), resumen: { filasExcel: json.filasExcel, modelos: json.modelos }, ms: { lectura: tLectura, total: Date.now() - t0 } });
      return;
    }
    if (tipo === 'supervisores') {
      const json = Convertir.parseSupervisores(wb);
      pMaestro = Promise.resolve(json); // lo recién subido manda para los Excel que se conviertan después
      const porSup = {};
      Object.keys(json.sup).forEach(function (k) { const s = json.sup[k][1]; porSup[s] = (porSup[s] || 0) + 1; });
      self.postMessage({ id, etapa: 'listo', tipo, texto: JSON.stringify(json), resumen: { tiendas: json.tiendas, filasExcel: json.filasExcel, sinSupervisor: json.sinSupervisor, porSup }, ms: { lectura: tLectura, total: Date.now() - t0 } });
      return;
    }
    maestro().then(function (m) {
      const ocs = Convertir.convertir(wb, nombre, m);
      const salida = ocs.map(function (o) { return { oc: o.oc, texto: JSON.stringify(o), resumen: Convertir.resumenDe(o) }; });
      self.postMessage({ id, etapa: 'listo', tipo, ocs: salida, maestro: m ? m.tiendas : 0, ms: { lectura: tLectura, total: Date.now() - t0 } });
    }).catch(function (err) { self.postMessage({ id, etapa: 'error', error: err.message || String(err) }); });
  } catch (err) {
    self.postMessage({ id, etapa: 'error', error: err.message || String(err) });
  }
};
