/* Web Worker de armar.html: lee los Excel sin congelar la pantalla.
 * Reconoce por el contenido: Ordenes_Picking (columnas NumAtCard + ItemCode), el maestro de supervisoras
 * (hoja Supervisores / columna Cadena) o, si no es ninguno, el listado de OC (primera hoja tal cual). */
importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'convertir.js', 'armar.js');

function filasDe(ws) {
  if (ws['!data']) return ws['!data'].map(function (row) { return row ? row.map(function (c) { return c ? c.v : null; }) : []; });
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
}
const t = v => v == null ? '' : String(v).trim();

self.onmessage = function (e) {
  const { id, nombre, buffer } = e.data;
  const t0 = Date.now();
  try {
    self.postMessage({ id, etapa: 'leyendo' });
    const wb = XLSX.read(buffer, { type: 'array', dense: true, cellFormula: false, cellHTML: false, cellText: false, cellStyles: false });
    self.postMessage({ id, etapa: 'revisando' });

    if (Convertir.identificar(nombre) === 'supervisores' || wb.SheetNames.some(n => n.trim().toLowerCase() === 'supervisores')) {
      const json = Convertir.parseSupervisores(wb);
      self.postMessage({ id, etapa: 'listo', tipo: 'supervisores', maestro: json, ms: Date.now() - t0 });
      return;
    }

    if (Convertir.identificar(nombre) === 'imagenes' || wb.SheetNames.some(n => n.trim().toLowerCase() === 'imagenes')) {
      const json = Convertir.parseImagenes(wb);
      self.postMessage({ id, etapa: 'listo', tipo: 'imagenes', imagenes: json, ms: Date.now() - t0 });
      return;
    }

    // picking: la primera hoja que tenga NumAtCard e ItemCode en sus títulos
    for (const hoja of wb.SheetNames) {
      const filas = filasDe(wb.Sheets[hoja]);
      const iHdr = filas.slice(0, 20).findIndex(r => r && r.some(c => t(c) === 'NumAtCard') && r.some(c => t(c) === 'ItemCode'));
      if (iHdr < 0) continue;
      const hdr = filas[iHdr].map(t);
      const idx = [], col = {}, faltan = [];
      Armar.COLS_PICKING.forEach(nc => {
        const obligatoria = nc.endsWith('*'), n = nc.replace('*', '');
        const i = hdr.indexOf(n);
        if (i < 0) { if (obligatoria) faltan.push(n); return; }
        col[n] = idx.length; idx.push(i);
      });
      if (faltan.length) throw new Error(`Hoja "${hoja}": faltan las columnas ${faltan.join(', ')}`);
      const filasPk = [];
      for (let k = iHdr + 1; k < filas.length; k++) {
        const r = filas[k];
        if (!r || r[hdr.indexOf('NumAtCard')] == null) continue;
        filasPk.push(idx.map(i => r[i] == null ? '' : r[i]));
      }
      self.postMessage({ id, etapa: 'listo', tipo: 'picking', filas: filasPk, col, totalFilas: filas.length - iHdr - 1, ms: Date.now() - t0 });
      return;
    }

    // si no, es el listado de OC: se devuelve la primera hoja como texto con tabuladores
    const filas = filasDe(wb.Sheets[wb.SheetNames[0]]).filter(r => r && r.some(c => t(c)));
    if (!filas.length) throw new Error('El Excel está vacío');
    const texto = filas.slice(0, 2000).map(r => r.map(t).join('\t')).join('\n');
    self.postMessage({ id, etapa: 'listo', tipo: 'lista', texto, ms: Date.now() - t0 });
  } catch (err) {
    self.postMessage({ id, etapa: 'error', error: err.message || String(err) });
  }
};
