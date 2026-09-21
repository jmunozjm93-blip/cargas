/* Web Worker: lee el Excel y lo convierte sin congelar la pantalla */
importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'convertir.js');

self.onmessage = function (e) {
  const { id, nombre, buffer } = e.data;
  const t0 = Date.now();
  try {
    self.postMessage({ id, etapa: 'leyendo' });
    const wb = XLSX.read(buffer, Convertir.opcionesLectura());
    const tLectura = Date.now() - t0;
    self.postMessage({ id, etapa: 'convirtiendo', hojas: wb.SheetNames });
    const ocs = Convertir.convertir(wb, nombre);
    const salida = ocs.map(o => ({ oc: o.oc, texto: JSON.stringify(o), resumen: Convertir.resumenDe(o) }));
    self.postMessage({ id, etapa: 'listo', ocs: salida, ms: { lectura: tLectura, total: Date.now() - t0 } });
  } catch (err) {
    self.postMessage({ id, etapa: 'error', error: err.message || String(err) });
  }
};
