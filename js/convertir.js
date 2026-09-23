/* Conversión del Excel de cargas (órdenes de compra por tienda) al JSON que consume el dashboard.
 * Corre en el navegador sobre libros abiertos con SheetJS (XLSX.read).
 * Regla de oro: los números son los del Excel, tal cual. No se estima nada.
 *
 * Formato del Excel (fijo): trae una hoja "Datos" con una fila por OC · modelo · talla · tienda:
 *   NumAtCard · Cliente · [Departamento] · modelo · Descripcion modelo · ShipToCode · Tienda · Supervisor ·
 *   ItemCode · Dscription · Quantity · DocNum · Direccion · StatusOrden
 * Opcionalmente una hoja "Resumen" con "OC (NumAtCard) · Comentario · Modelos · Tiendas · UND OC · …" y,
 * en la fila 2, "Fuente: <archivo de picking>". Las hojas con nombre de OC (dinámicas) no se usan.
 */
(function (global) {
  'use strict';

  // Excel_Macro.xlsx (hoja Imagenes): catálogo de fotos de los modelos, "MODELO.jpg" → ID de Google Drive
  const RE_IMAGENES = /^Excel_Macro\.xlsx$/i;
  // Supervisores_Consolidado.xlsx (hoja Supervisores): maestro Cadena · Cod · Tienda · Supervisor
  const RE_SUPERVISORES = /^Supervisores[_ ]?Consolidado\.xlsx$/i;
  function identificar(nombre) {
    const n = nombre.trim();
    if (RE_IMAGENES.test(n)) return 'imagenes';
    if (RE_SUPERVISORES.test(n)) return 'supervisores';
    return 'cargas';
  }

  // cliente del Excel de cargas / cadena del maestro → la misma clave
  function claveCliente(s) {
    const n = String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z]/g, '');
    if (n.indexOf('FALABELLA') >= 0) return 'FALABELLA';
    if (n.indexOf('RIPLEY') >= 0) return 'RIPLEY';
    if (n.indexOf('PARIS') >= 0) return 'PARIS';
    if (n.indexOf('POLAR') >= 0) return 'LAPOLAR';
    if (n.indexOf('HITES') >= 0) return 'HITES';
    return n;
  }

  const COLUMNAS = ['NumAtCard', 'Cliente', 'modelo', 'Descripcion modelo', 'ShipToCode', 'Tienda', 'Supervisor', 'Quantity'];
  // "Departamento" es opcional: si no viene, se toma del nombre del archivo (Cargas_21.09_Paris-deporte.xlsx → Deporte)
  const DEPTOS = [[/calzado[\s_-]*(dama|mujer)/i, 'Calzado dama'], [/deporte[\s_-]*mujer/i, 'Deporte mujer'], [/deporte[\s_-]*hombre/i, 'Deporte hombre'], [/deporte/i, 'Deporte'], [/kids|ni[ñn]o/i, 'Kids'], [/juvenil/i, 'Juvenil'], [/accesorio/i, 'Accesorios'], [/ropa|vestuario/i, 'Ropa']];
  function deptoDeNombre(nombre) { const n = String(nombre || '').replace(/\.xlsx$/i, ''); for (const [re, d] of DEPTOS) if (re.test(n)) return d; return ''; }

  // primer segmento de la descripción ("CONV|CALZ |DAY ONE…") → marca que se muestra
  const MARCAS = { CONV: 'Converse', CONVERSE: 'Converse', FILA: 'Fila', UMB: 'Umbro', UMBR: 'Umbro', UMBRO: 'Umbro' };

  // valores de la columna Supervisor que significan "nadie"
  const SIN_SUPERVISOR = ['', 'Z', '-', 'N/A', 'NA', 'SIN ASIGNAR', 'SIN SUPERVISOR', 'RUTA NUEVO', 'RUTA NUEVA'];

  const vacio = v => v == null || v === '';
  const txt = v => vacio(v) ? '' : String(v).trim();
  const nombrePersona = v => txt(v).replace(/\s+/g, ' ').toUpperCase();
  const cod = v => (typeof v === 'number') ? String(v) : txt(v);
  const num = v => { if (typeof v === 'number') return v; const n = parseFloat(String(v).replace(',', '.')); return isNaN(n) ? 0 : n; };

  function filasDe(ws) {
    if (ws['!data']) return ws['!data'].map(row => row ? row.map(c => (c ? c.v : null)) : []);
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  }

  function marcaDe(desc, modelo) {
    const seg = txt(desc).split('|')[0].trim().toUpperCase();
    if (MARCAS[seg]) return MARCAS[seg];
    if (/^[A-Z]?\d{5,6}[A-Z]-/.test(modelo)) return 'Converse'; // A20633C-102, 1U646C-…
    return seg ? seg.charAt(0) + seg.slice(1).toLowerCase() : '';
  }

  function opcionesLectura(tipo) {
    const o = { type: 'array', dense: true, cellFormula: false, cellHTML: false, cellText: false, cellStyles: false };
    if (tipo === 'imagenes') o.sheets = ['Imagenes'];
    if (tipo === 'supervisores') o.sheets = ['Supervisores'];
    return o;
  }

  // ---------- maestro de supervisoras (Supervisores_Consolidado.xlsx, hoja Supervisores) ----------
  // "CADENA|cod" → [nombre de tienda, supervisora]. Manda sobre la columna Supervisor del Excel de cargas.
  function parseSupervisores(wb) {
    const nombre = wb.SheetNames.find(n => n.trim().toLowerCase() === 'supervisores') || wb.SheetNames[0];
    const filas = filasDe(wb.Sheets[nombre]);
    const iHdr = filas.findIndex(r => r && txt(r[0]).toLowerCase() === 'cadena');
    if (iHdr < 0) throw new Error(`La hoja "${nombre}" no tiene la fila de títulos (Cadena · Cod · Tienda · Supervisor)`);
    const sup = {};
    let leidas = 0, sinSup = 0;
    for (const r of filas.slice(iHdr + 1)) {
      if (!r) continue;
      const cad = claveCliente(r[0]), c = cod(r[1]), tienda = txt(r[2]), quien = txt(r[3]);
      if (!cad || !c) continue;
      leidas++;
      if (SIN_SUPERVISOR.includes(nombrePersona(quien))) { sinSup++; continue; }
      sup[cad + '|' + c] = [tienda || c, quien];
    }
    if (!leidas) throw new Error(`La hoja "${nombre}" no tiene filas con Cadena y Cod`);
    return { generado: new Date().toISOString(), archivo: 'Supervisores_Consolidado.xlsx', tiendas: Object.keys(sup).length, filasExcel: leidas, sinSupervisor: sinSup, sup };
  }

  // ---------- imágenes (Excel_Macro.xlsx, hoja Imagenes) ----------
  // Si un modelo aparece varias veces, manda la última fila.
  function parseImagenes(wb) {
    const nombre = wb.SheetNames.find(n => n.trim().toLowerCase() === 'imagenes') || wb.SheetNames[0];
    const filas = filasDe(wb.Sheets[nombre]);
    const img = {};
    let leidas = 0;
    for (const r of filas) {
      const archivo = txt(r[0]), link = txt(r[1]);
      if (!archivo || /^nombre de/i.test(archivo)) continue;
      const m = link.match(/[?&]id=([A-Za-z0-9_-]+)/);
      if (!m) continue;
      leidas++;
      img[archivo.replace(/\.(jpe?g|png|webp)$/i, '').toUpperCase()] = m[1];
    }
    if (!leidas) throw new Error(`La hoja "${nombre}" no tiene links de Google Drive (columnas Nombre del Archivo · Link)`);
    return { generado: new Date().toISOString(), filasExcel: leidas, modelos: Object.keys(img).length, img };
  }

  // ---------- hoja Resumen (opcional): comentario y unidades declaradas por OC, archivo de origen ----------
  function leerResumen(wb) {
    const nombre = wb.SheetNames.find(n => n.toLowerCase() === 'resumen');
    const out = { fuente: '', ocs: {} };
    if (!nombre) return out;
    const filas = filasDe(wb.Sheets[nombre]);
    for (const r of filas.slice(0, 5)) { const m = /^fuente:\s*(.+)$/i.exec(txt(r[0])); if (m) out.fuente = m[1].trim(); }
    const iHdr = filas.findIndex(r => /^oc\b/i.test(txt(r[0])));
    if (iHdr < 0) return out;
    const hdr = filas[iHdr].map(h => txt(h).toLowerCase());
    const iCom = hdr.findIndex(h => h.startsWith('comentario'));
    const iUnd = hdr.findIndex(h => h === 'und oc' || h.startsWith('unidades oc'));
    for (const r of filas.slice(iHdr + 1)) {
      const oc = cod(r[0]);
      if (!/^\d+$/.test(oc)) continue;
      out.ocs[oc] = { comentario: iCom >= 0 ? txt(r[iCom]) : '', udsOC: iUnd >= 0 ? num(r[iUnd]) : null };
    }
    return out;
  }

  // ---------- hoja Datos ----------
  function convertir(wb, nombreArchivo, maestro) {
    const mSup = (maestro && maestro.sup) || {};
    const nombre = wb.SheetNames.find(n => n.toLowerCase() === 'datos');
    if (!nombre) throw new Error(`El Excel no tiene la hoja "Datos" (hojas: ${wb.SheetNames.join(' · ')})`);
    const filas = filasDe(wb.Sheets[nombre]);
    const iHdr = filas.findIndex(r => r && r.some(c => txt(c) === 'NumAtCard'));
    if (iHdr < 0) throw new Error('Hoja "Datos": no encontré la fila de títulos (debe tener la columna NumAtCard)');
    const hdr = filas[iHdr].map(txt);
    const col = {};
    for (const c of COLUMNAS) { const i = hdr.indexOf(c); if (i < 0) throw new Error(`Hoja "Datos": falta la columna "${c}"`); col[c] = i; }
    const iStatus = hdr.indexOf('StatusOrden');
    const iDepto = hdr.indexOf('Departamento');
    const deptoArchivo = deptoDeNombre(nombreArchivo);

    const resumen = leerResumen(wb);
    const ocs = {};
    let filasLeidas = 0;
    for (const r of filas.slice(iHdr + 1)) {
      if (!r) continue;
      const oc = cod(r[col.NumAtCard]);
      if (!oc) continue;
      filasLeidas++;
      const o = ocs[oc] || (ocs[oc] = {
        oc, cliente: txt(r[col.Cliente]), depto: iDepto >= 0 ? txt(r[iDepto]) : deptoArchivo, marcas: {}, tiendas: {}, modelos: {}, estados: {},
      });
      const modelo = txt(r[col.modelo]), desc = txt(r[col['Descripcion modelo']]);
      const tCod = cod(r[col.ShipToCode]), tNom = txt(r[col.Tienda]);
      const sup = nombrePersona(r[col.Supervisor]);
      const q = num(r[col.Quantity]);
      const marca = marcaDe(desc, modelo);
      o.marcas[marca] = (o.marcas[marca] || 0) + q;
      if (iStatus >= 0) { const s = txt(r[iStatus]); if (s) o.estados[s] = (o.estados[s] || 0) + q; }
      const m = o.modelos[modelo] || (o.modelos[modelo] = { modelo, desc, marca, uds: 0 });
      m.uds += q;
      // el maestro (Supervisores_Consolidado) manda; si la tienda no está ahí, se usa lo que traiga el Excel
      const mt = mSup[claveCliente(o.cliente) + '|' + tCod];
      const t = o.tiendas[tCod] || (o.tiendas[tCod] = {
        cod: tCod, nombre: (mt && mt[0]) || tNom || tCod,
        sup: mt ? nombrePersona(mt[1]) : (SIN_SUPERVISOR.includes(sup) ? '' : sup),
        deMaestro: !!mt, uds: 0, items: {},
      });
      t.uds += q;
      t.items[modelo] = (t.items[modelo] || 0) + q;
    }
    if (!filasLeidas) throw new Error('Hoja "Datos": no tiene filas con OC');

    const generado = new Date().toISOString();
    const salida = [];
    for (const oc of Object.keys(ocs).sort()) {
      const o = ocs[oc];
      const marcas = Object.entries(o.marcas).sort((a, b) => b[1] - a[1]).map(x => x[0]).filter(Boolean);
      const res = resumen.ocs[oc] || {};
      const tiendas = Object.values(o.tiendas)
        .sort((a, b) => (a.sup || '~').localeCompare(b.sup || '~') || a.nombre.localeCompare(b.nombre))
        .map(t => ({ cod: t.cod, nombre: t.nombre, sup: t.sup, uds: t.uds,
          items: Object.entries(t.items).map(([modelo, uds]) => ({ modelo, uds })).sort((a, b) => a.modelo.localeCompare(b.modelo)) }));
      salida.push({
        oc, cliente: o.cliente, depto: o.depto, marca: marcas.join(' · '),
        comentario: res.comentario || '', udsOC: res.udsOC == null ? null : res.udsOC,
        fuente: resumen.fuente, archivo: nombreArchivo, generado,
        uds: tiendas.reduce((s, t) => s + t.uds, 0),
        estados: o.estados,
        modelos: Object.values(o.modelos).sort((a, b) => a.modelo.localeCompare(b.modelo)),
        tiendas,
      });
    }
    return salida;
  }

  // lo que va al manifiesto (sin el detalle de modelos por tienda)
  function resumenDe(o) {
    return {
      oc: o.oc, cliente: o.cliente, depto: o.depto, marca: o.marca, comentario: o.comentario, udsOC: o.udsOC,
      fuente: o.fuente, archivo: o.archivo, uds: o.uds, modelos: o.modelos.length,
      sinSupervisor: o.tiendas.filter(t => !t.sup).map(t => t.nombre),
      tiendas: o.tiendas.map(t => ({ cod: t.cod, nombre: t.nombre, sup: t.sup, uds: t.uds, modelos: t.items.length })),
    };
  }

  global.Convertir = { identificar, opcionesLectura, convertir, parseImagenes, parseSupervisores, claveCliente, resumenDe, COLUMNAS };
})(typeof self !== 'undefined' ? self : this);
