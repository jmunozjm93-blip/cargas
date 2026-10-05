/* Armado del Excel de cargas a partir de:
 *   - el listado de OC (Marca · Departamento · N°DEOC · Familia · UND OC), pegado desde Excel
 *   - Ordenes_Picking <dd.mm>.xlsx (una fila por DocNum · ItemCode · tienda)
 *   - el maestro de supervisoras (data/supervisores.json o Supervisores_Consolidado.xlsx)
 *   - las fotos de los modelos (data/imagenes.json → miniaturas de Google Drive)
 * Sale el mismo formato que Cargas_<dd.mm>_<Cliente>.xlsx: hojas Resumen, Datos y una por OC
 * (filas = modelo, columnas = tienda con su supervisora arriba, valores = suma de Quantity).
 * Regla de oro: los números son los del picking, tal cual. Sin filtro por StatusOrden.
 * Necesita convertir.js (claveTienda) y, para escribir el Excel, ExcelJS.
 */
(function (global) {
  'use strict';

  const vacio = v => v == null || v === '';
  const txt = v => vacio(v) ? '' : String(v).trim();
  const cod = v => (typeof v === 'number') ? String(v) : txt(v);
  const num = v => { if (typeof v === 'number') return v; const n = parseFloat(String(v).replace(/[.\s]/g, '').replace(',', '.')); return isNaN(n) ? 0 : n; };
  const nombrePersona = v => txt(v).replace(/\s+/g, ' ').toUpperCase();
  const titulo = s => txt(s).toLowerCase().replace(/(^|\s)\S/g, c => c.toUpperCase());
  const SIN_SUPERVISOR = 'SIN ASIGNAR';

  // ---------- listado de OC (pegado desde Excel: columnas separadas por tabulador) ----------
  const CAMPOS = [
    ['oc', /n\W*°?\W*de\W*oc|^n.?\s*oc$|^oc\b|numatcard|orden/i],
    ['marca', /marca/i],
    ['depto', /depart/i],
    ['familia', /famil/i],
    ['udsOC', /und|unid|cant/i],
  ];
  // sin fila de títulos se asume el orden de la planilla: Marca · Departamento · N°DEOC · Familia · UND OC
  const ORDEN_FIJO = { marca: 0, depto: 1, oc: 2, familia: 3, udsOC: 4 };

  function filasDeTexto(texto) {
    return String(texto || '').replace(/\r/g, '').split('\n')
      .map(l => (l.indexOf('\t') >= 0 ? l.split('\t') : l.split(/;|\s{2,}/)).map(c => c.trim()))
      .filter(r => r.some(c => c));
  }

  function parseLista(filas) {
    const avisos = [];
    let col = null, desde = 0;
    const iHdr = filas.findIndex(r => r.some(c => /n\W*°?\W*de\W*oc|numatcard|^oc\b/i.test(txt(c))));
    if (iHdr >= 0) {
      col = {};
      const hdr = filas[iHdr].map(txt);
      for (const [k, re] of CAMPOS) { const i = hdr.findIndex(h => re.test(h)); if (i >= 0) col[k] = i; }
      desde = iHdr + 1;
    }
    const items = [], vistos = {};
    for (let n = desde; n < filas.length; n++) {
      const r = filas[n].map(txt);
      let c = col;
      if (!c) {
        if (r.length >= 5) c = ORDEN_FIJO;
        else { const i = r.findIndex(x => /^\d{5,}$/.test(x.replace(/\s/g, ''))); c = { oc: i < 0 ? 0 : i }; }
      }
      const oc = txt(r[c.oc]).replace(/\s/g, '').replace(/\.0+$/, '');
      if (!oc) continue;
      if (!/^\d+$/.test(oc)) { avisos.push(`Fila ${n + 1}: "${oc}" no es un número de OC, se ignora`); continue; }
      const it = {
        oc, fila: n + 1,
        marca: c.marca != null ? txt(r[c.marca]).toUpperCase() : '',
        depto: c.depto != null ? titulo(r[c.depto]) : '',
        familia: c.familia != null ? titulo(r[c.familia]) : '',
        udsOC: c.udsOC != null && txt(r[c.udsOC]) !== '' ? num(r[c.udsOC]) : null,
      };
      if (vistos[oc]) { avisos.push(`La OC ${oc} está repetida (filas ${vistos[oc].fila} y ${it.fila}): se usa la primera`); continue; }
      vistos[oc] = it;
      items.push(it);
    }
    return { items, avisos };
  }

  // ---------- Ordenes_Picking ----------
  // columnas que se leen del picking; las marcadas con * son obligatorias
  const COLS_PICKING = ['NumAtCard*', 'CardName*', 'ShipToCode*', 'ItemCode*', 'Dscription*', 'Quantity*', 'DocNum*',
    'ObjTypeDesc', 'CANCELED', 'LineNum', 'Address2', 'StatusOrden', 'WhsCode', 'DocDate'];

  function clienteDe(cardName, whs) {
    const n = String(cardName || '').toUpperCase(), w = String(whs || '').toUpperCase();
    if (/CENCOSUD|PARIS/.test(n)) return 'Paris';
    if (/ECCSA|RIPLEY/.test(n)) return 'Ripley';
    if (/FALABELLA/.test(n)) return 'Falabella';
    if (/POLAR/.test(n)) return 'La Polar';
    if (/HITES/.test(n)) return 'Hites';
    if (w === 'C_PARIS') return 'Paris';
    if (w === 'C_RIPLEY') return 'Ripley';
    if (w === 'C_FALAB') return 'Falabella';
    if (w === 'C_HITES') return 'Hites';
    if (/POLAR/.test(w)) return 'La Polar';
    return titulo(cardName);
  }

  // ItemCode sin la talla (A24087C-234-7 → A24087C-234); Dscription sin el último segmento (la talla)
  const modeloDe = item => { const s = txt(item); const i = s.lastIndexOf('-'); return i > 0 ? s.slice(0, i) : s; };
  const descModeloDe = d => { const s = txt(d); const i = s.lastIndexOf('|'); return (i > 0 ? s.slice(0, i) : s).trim(); };
  // "MARINA ARAUCO  SANTIAGO CHILE" → "MARINA ARAUCO" (el picking separa la ciudad con doble espacio)
  const lugarDe = a => txt(a).split(/\s{2,}/)[0].trim();

  // Filas del picking (arrays) → índice por OC. En el worker ya vienen solo las columnas de COLS_PICKING.
  function indexarPicking(filas, col) {
    const porOC = {};
    let canceladas = 0;
    for (const r of filas) {
      const oc = cod(r[col.NumAtCard]).replace(/\.0+$/, '');
      if (!oc) continue;
      if (col.CANCELED != null && txt(r[col.CANCELED]).toUpperCase() === 'Y') { canceladas++; continue; }
      (porOC[oc] || (porOC[oc] = [])).push(r);
    }
    // totales por OC para sugerir cuando una OC del listado no aparece (ej. un dígito de menos)
    const totales = {};
    for (const oc in porOC) {
      const rs = porOC[oc];
      const tipo = col.ObjTypeDesc != null ? txt(rs[0][col.ObjTypeDesc]) : '';
      totales[oc] = { uds: rs.reduce((s, r) => s + num(r[col.Quantity]), 0), cliente: clienteDe(rs[0][col.CardName], col.WhsCode != null ? rs[0][col.WhsCode] : ''), tipo };
    }
    return { porOC, totales, col, canceladas };
  }

  function distancia(a, b) {
    const m = a.length, n = b.length, d = Array.from({ length: m + 1 }, (_, i) => [i]);
    for (let j = 1; j <= n; j++) d[0][j] = j;
    for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++)
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    return d[m][n];
  }

  // OC del picking parecidas a la que no se encontró: primero las que calzan en unidades con UND OC
  function sugerencias(oc, udsOC, pk) {
    const out = [];
    for (const [otra, t] of Object.entries(pk.totales)) {
      if (t.tipo && t.tipo !== 'Ordenes') continue;
      const d = distancia(oc, otra);
      const mismasUds = udsOC != null && t.uds === udsOC;
      if (d <= 2 || (mismasUds && (otra.indexOf(oc) >= 0 || oc.indexOf(otra) >= 0 || d <= 3))) out.push({ oc: otra, uds: t.uds, cliente: t.cliente, d, mismasUds });
    }
    return out.sort((a, b) => (b.mismasUds - a.mismasUds) || (a.d - b.d) || a.oc.localeCompare(b.oc)).slice(0, 3);
  }

  // ---------- armado ----------
  function armar(lista, pk, maestro) {
    const mSup = (maestro && maestro.sup) || {};
    const c = pk.col;
    const ocs = [];
    for (const it of lista) {
      const filas = pk.porOC[it.oc];
      if (!filas) { ocs.push(Object.assign({}, it, { encontrada: false, sugerencias: sugerencias(it.oc, it.udsOC, pk) })); continue; }

      // duplicados: la misma OC cargada más de una vez en SAP con distinto DocNum → por OC + tienda + ItemCode manda el DocNum más reciente
      const maxDoc = {};
      for (const r of filas) {
        const k = cod(r[c.ShipToCode]) + '|' + txt(r[c.ItemCode]);
        const d = num(r[c.DocNum]);
        if (!(k in maxDoc) || d > maxDoc[k]) maxDoc[k] = d;
      }
      let descartadas = 0;
      const vivas = filas.filter(r => {
        const ok = num(r[c.DocNum]) === maxDoc[cod(r[c.ShipToCode]) + '|' + txt(r[c.ItemCode])];
        if (!ok) descartadas++;
        return ok;
      });

      const cliente = clienteDe(filas[0][c.CardName], c.WhsCode != null ? filas[0][c.WhsCode] : '');
      const tiendas = {}, modelos = {}, estados = {}, tipos = {}, datos = [];
      for (const r of vivas) {
        const tCod = cod(r[c.ShipToCode]);
        const q = num(r[c.Quantity]);
        let t = tiendas[tCod];
        if (!t) {
          const mt = mSup[Convertir.claveTienda(cliente, tCod)];
          const lugar = c.Address2 != null ? lugarDe(r[c.Address2]) : '';
          t = tiendas[tCod] = { cod: tCod, nombre: mt ? mt[0] : (lugar ? `${tCod}-${lugar}` : tCod), sup: mt ? nombrePersona(mt[1]) : SIN_SUPERVISOR, deMaestro: !!mt, uds: 0 };
        }
        t.uds += q;
        const modelo = modeloDe(r[c.ItemCode]);
        const m = modelos[modelo] || (modelos[modelo] = { modelo, desc: descModeloDe(r[c.Dscription]), porTienda: {}, uds: 0 });
        m.porTienda[tCod] = (m.porTienda[tCod] || 0) + q;
        m.uds += q;
        const st = c.StatusOrden != null ? txt(r[c.StatusOrden]) : '';
        if (st) estados[st] = (estados[st] || 0) + q;
        const tp = c.ObjTypeDesc != null ? txt(r[c.ObjTypeDesc]) : '';
        if (tp) tipos[tp] = (tipos[tp] || 0) + 1;
        datos.push([it.oc, cliente, it.depto, modelo, m.desc, tCod, t.nombre, t.sup, txt(r[c.ItemCode]), txt(r[c.Dscription]), q,
          cod(r[c.DocNum]), c.Address2 != null ? txt(r[c.Address2]) : '', st, it.marca, it.familia]);
      }
      // columnas ordenadas por supervisora y luego por tienda; las sin asignar al final
      const listaTiendas = Object.values(tiendas).sort((a, b) =>
        ((a.sup === SIN_SUPERVISOR) - (b.sup === SIN_SUPERVISOR)) || a.sup.localeCompare(b.sup) || a.nombre.localeCompare(b.nombre, 'es', { numeric: true }));
      const uds = vivas.reduce((s, r) => s + num(r[c.Quantity]), 0);
      ocs.push(Object.assign({}, it, {
        encontrada: true, cliente, uds, descartadas, estados, tipos, datos,
        tiendas: listaTiendas,
        sinSupervisor: listaTiendas.filter(t => t.sup === SIN_SUPERVISOR),
        modelos: Object.values(modelos).sort((a, b) => a.modelo.localeCompare(b.modelo)),
      }));
    }
    return ocs;
  }

  // ---------- nombre del archivo ----------
  function fechaPicking(nombre) {
    const m = /(\d{1,2})[.\-_](\d{1,2})(?!\d)/.exec(String(nombre || ''));
    if (m) return m[1].padStart(2, '0') + '.' + m[2].padStart(2, '0');
    const d = new Date();
    return String(d.getDate()).padStart(2, '0') + '.' + String(d.getMonth() + 1).padStart(2, '0');
  }
  function nombreArchivo(ocs, fuente) {
    const ok = ocs.filter(o => o.encontrada);
    const unico = xs => [...new Set(xs.filter(Boolean))];
    const clientes = unico(ok.map(o => o.cliente)).map(s => s.replace(/\s+/g, '_'));
    const deptos = unico(ok.map(o => o.depto));
    let n = `Cargas_${fechaPicking(fuente)}_${clientes.join('_') || 'OC'}`;
    if (deptos.length === 1) n += '-' + deptos[0].toLowerCase().replace(/\s+/g, '_');
    return n + '.xlsx';
  }

  // ---------- Excel (ExcelJS) ----------
  const COLOR = { hdr: 'FF44546A', hdrTxt: 'FFFFFFFF', celeste: 'FFBDD7EE', amarillo: 'FFFFF2CC', gris: 'FFF2F2F2', total: 'FF1F4E79', rojo: 'FFC00000' };
  const relleno = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const BORDE = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  const FOTO_PX = 78;

  function hojaResumen(wb, ocs, fuente) {
    const ws = wb.addWorksheet('Resumen');
    const ok = ocs.filter(o => o.encontrada);
    const clientes = [...new Set(ok.map(o => o.cliente))].join(' · ');
    ws.getCell('A1').value = `Cargas ${fechaPicking(fuente)}${clientes ? ' - ' + clientes : ''}`;
    ws.getCell('A1').font = { bold: true, size: 14 };
    ws.getCell('A2').value = `Fuente: ${fuente}`;
    ws.getCell('A2').font = { italic: true, size: 9 };
    const hdr = ['OC (NumAtCard)', 'Cliente', 'Marca', 'Departamento', 'Familia', 'Modelos', 'Tiendas', 'UND OC', 'Unidades picking', 'Diferencia'];
    ws.getRow(4).values = hdr;
    ws.getRow(4).eachCell(cl => { cl.font = { bold: true, color: { argb: COLOR.hdrTxt } }; cl.fill = relleno(COLOR.hdr); cl.border = BORDE; cl.alignment = { vertical: 'middle' }; });
    let r = 5, totOC = 0, totPk = 0;
    for (const o of ocs) {
      const dif = o.encontrada && o.udsOC != null ? o.uds - o.udsOC : null;
      ws.getRow(r).values = [Number(o.oc), o.cliente || '', o.marca, o.depto, o.familia,
        o.encontrada ? o.modelos.length : null, o.encontrada ? o.tiendas.length : null, o.udsOC, o.encontrada ? o.uds : 'No está en el picking', dif];
      ws.getRow(r).eachCell({ includeEmpty: true }, (cl, i) => {
        if (i > hdr.length) return;
        cl.border = BORDE; cl.fill = relleno(COLOR.amarillo);
        if (i >= 6) cl.numFmt = '#,##0';
      });
      ws.getCell(r, 1).numFmt = '0';
      if (!o.encontrada) ws.getCell(r, 9).font = { bold: true, color: { argb: COLOR.rojo } };
      if (dif) ws.getCell(r, 10).font = { bold: true, color: { argb: COLOR.rojo } };
      totOC += o.udsOC || 0; totPk += o.encontrada ? o.uds : 0;
      r++;
    }
    ws.getRow(r).values = ['TOTAL', '', '', '', '', '', '', totOC, totPk, totPk - totOC];
    ws.getRow(r).eachCell({ includeEmpty: true }, (cl, i) => { if (i > hdr.length) return; cl.font = { bold: true }; cl.fill = relleno(COLOR.gris); cl.border = BORDE; if (i >= 8) cl.numFmt = '#,##0'; });
    r += 2;
    const sinSup = [...new Set(ok.flatMap(o => o.sinSupervisor.map(t => `${o.cliente} ${t.cod}`)))];
    const notas = [
      'Estructura: Filas = modelo (ItemCode sin talla) / Columnas = ShipToCode / Valores = Suma de Quantity / Filtro = NumAtCard.',
      'Duplicados: si la misma OC viene más de una vez en SAP bajo distinto DocNum, por OC + tienda + ItemCode se conserva el DocNum más reciente.',
      'Fotos desde el catálogo Excel_Macro.xlsx (miniaturas de Google Drive), cruzadas por modelo.',
      'Descripcion = campo Dscription del origen, sin el último segmento (la talla).',
      'Sin filtro por StatusOrden. Columnas ordenadas por supervisor y luego por tienda.',
      'Tienda y supervisor cruzados por ShipToCode contra Supervisores_Consolidado.xlsx.',
    ];
    if (sinSup.length) notas.push('Tiendas sin supervisor en el maestro: ' + sinSup.join(', '));
    const faltan = ocs.filter(o => !o.encontrada).map(o => o.oc);
    if (faltan.length) notas.push('OC del listado que no están en el picking: ' + faltan.join(', '));
    ws.getCell(r, 1).value = 'NOTAS'; ws.getCell(r, 1).font = { bold: true, size: 9 };
    for (const n of notas) { r++; ws.getCell(r, 1).value = n; ws.getCell(r, 1).font = { size: 9 }; }
    [17, 12, 14, 16, 16, 10, 10, 11, 17, 12].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    ws.views = [{ state: 'frozen', ySplit: 4 }];
  }

  const HDR_DATOS = ['NumAtCard', 'Cliente', 'Departamento', 'modelo', 'Descripcion modelo', 'ShipToCode', 'Tienda', 'Supervisor', 'ItemCode', 'Dscription', 'Quantity', 'DocNum', 'Direccion', 'StatusOrden', 'Marca', 'Familia'];
  function hojaDatos(wb, ocs) {
    const ws = wb.addWorksheet('Datos');
    ws.addRow(HDR_DATOS);
    ws.getRow(1).eachCell(cl => { cl.font = { bold: true, color: { argb: COLOR.hdrTxt } }; cl.fill = relleno(COLOR.hdr); });
    for (const o of ocs) if (o.encontrada) for (const d of o.datos) {
      const fila = d.slice();
      fila[0] = Number(fila[0]);
      if (/^\d+$/.test(fila[5])) fila[5] = Number(fila[5]);
      if (/^\d+$/.test(fila[11])) fila[11] = Number(fila[11]);
      ws.addRow(fila);
    }
    [11, 10, 13, 16, 48, 11, 30, 22, 20, 52, 9, 11, 34, 13, 11, 13].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: HDR_DATOS.length } };
    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  function hojaOC(wb, o, fotos, fuente) {
    const ws = wb.addWorksheet(o.oc.slice(0, 31), { views: [{ state: 'frozen', xSplit: 3, ySplit: 2 }] });
    const nT = o.tiendas.length, cTot = 4 + nT;
    // fila 1: supervisora encima de cada tienda · fila 2: títulos y tiendas (texto vertical, como la tabla dinámica)
    const f1 = ws.getRow(1), f2 = ws.getRow(2);
    f2.getCell(1).value = 'Etiquetas de fila'; f2.getCell(2).value = 'Descripcion'; f2.getCell(3).value = 'Foto';
    o.tiendas.forEach((t, i) => { f1.getCell(4 + i).value = t.sup; f2.getCell(4 + i).value = t.nombre; });
    f2.getCell(cTot).value = 'Total general';
    for (let cI = 1; cI <= cTot; cI++) for (const f of [f1, f2]) {
      const cl = f.getCell(cI);
      cl.fill = relleno(COLOR.celeste); cl.border = BORDE; cl.font = { name: 'Arial', size: cI >= 4 ? 9 : 10, bold: true };
      cl.alignment = cI >= 4 && cI < cTot ? { textRotation: 90, horizontal: 'center', vertical: 'bottom' } : cI === cTot ? { wrapText: true, horizontal: 'center', vertical: 'middle' } : { vertical: 'middle' };
      if (cI < cTot && o.tiendas[cI - 4] && o.tiendas[cI - 4].sup === SIN_SUPERVISOR && f === f1) cl.font = { name: 'Arial', size: 9, bold: true, color: { argb: COLOR.rojo } };
    }
    f1.height = 120; f2.height = 170;

    let r = 3;
    for (const m of o.modelos) {
      const fila = ws.getRow(r);
      fila.getCell(1).value = m.modelo;
      fila.getCell(2).value = m.desc;
      o.tiendas.forEach((t, i) => { const q = m.porTienda[t.cod]; if (q) fila.getCell(4 + i).value = q; });
      fila.getCell(cTot).value = m.uds;
      for (let cI = 1; cI <= cTot; cI++) {
        const cl = fila.getCell(cI);
        cl.border = BORDE; cl.font = { name: 'Arial', size: cI === 2 ? 9 : 10 };
        cl.alignment = cI === 2 ? { wrapText: true, vertical: 'middle' } : cI >= 4 ? { horizontal: 'center', vertical: 'middle' } : { vertical: 'middle' };
      }
      fila.getCell(cTot).font = { name: 'Arial', size: 10, bold: true, color: { argb: COLOR.total } };
      const f = fotos && fotos.fotos[m.modelo.toUpperCase()];
      const link = fotos && fotos.links[m.modelo.toUpperCase()];
      fila.height = f ? 64 : 30;
      if (link) { fila.getCell(3).value = { text: 'ver foto', hyperlink: link }; fila.getCell(3).font = { name: 'Arial', size: 9, underline: true, color: { argb: 'FF0563C1' } }; }
      if (f) {
        const id = wb.addImage({ buffer: f.buffer, extension: f.ext });
        const esc = Math.min(FOTO_PX / f.w, FOTO_PX / f.h);
        ws.addImage(id, { tl: { col: 2.08, row: r - 1 + 0.06 }, ext: { width: Math.round(f.w * esc), height: Math.round(f.h * esc) }, editAs: 'oneCell' });
      }
      r++;
    }
    const ft = ws.getRow(r);
    ft.getCell(1).value = 'Total general';
    o.tiendas.forEach((t, i) => { ft.getCell(4 + i).value = t.uds; });
    ft.getCell(cTot).value = o.uds;
    for (let cI = 1; cI <= cTot; cI++) {
      const cl = ft.getCell(cI);
      cl.fill = relleno(COLOR.gris); cl.border = BORDE;
      cl.font = { name: 'Arial', size: 10, bold: true, color: cI >= 4 ? { argb: COLOR.total } : undefined };
      cl.alignment = { horizontal: cI >= 4 ? 'center' : 'left', vertical: 'middle' };
    }

    r += 2;
    const dif = o.udsOC != null ? o.uds - o.udsOC : null;
    const est = Object.entries(o.estados).map(([s, q]) => `${s} ${q}`).join(' · ');
    const notas = [
      `OC (NumAtCard): ${o.oc} | Cliente: ${o.cliente} | Marca: ${o.marca || '—'} | Departamento: ${o.depto || '—'} | Familia: ${o.familia || '—'}`,
      o.udsOC != null ? `Unidades OC declaradas: ${o.udsOC} | Unidades en este picking: ${o.uds}${dif ? `  <-- DIFERENCIA ${dif > 0 ? '+' : ''}${dif}` : ' (cuadra)'}` : `Unidades en este picking: ${o.uds}`,
      est ? `Estado en el picking (unidades): ${est}` : '',
      `Valores = Suma de Quantity, sin filtro de estado. Detalle línea a línea en la hoja Datos. Fuente: ${fuente}.`,
      o.descartadas ? `Duplicados: se descartaron ${o.descartadas} línea(s) de un DocNum anterior (por OC + ItemCode + tienda se conservó el DocNum más reciente).` : '',
      o.sinSupervisor.length ? `Tiendas sin supervisor en el maestro: ${o.sinSupervisor.map(t => t.cod).join(', ')}` : '',
    ].filter(Boolean);
    for (const n of notas) {
      const cl = ws.getCell(r++, 1); cl.value = n;
      cl.font = { name: 'Arial', size: 9, bold: /DIFERENCIA/.test(n), color: /DIFERENCIA/.test(n) ? { argb: COLOR.rojo } : undefined };
    }

    ws.getColumn(1).width = 19; ws.getColumn(2).width = 46; ws.getColumn(3).width = 13;
    for (let i = 0; i < nT; i++) ws.getColumn(4 + i).width = 7.5;
    ws.getColumn(cTot).width = 10;
  }

  // fotos: miniaturas de Google Drive (lh3 permite leerlas desde la página). Google corta con 429 si se piden
  // muchas seguidas: de a 3 en paralelo, reintentando con espera. Lo ya bajado queda en memoria para el próximo Excel.
  const cacheFotos = {};
  const esperar = ms => new Promise(r => setTimeout(r, ms));
  async function bajarFoto(id) {
    for (let intento = 0; ; intento++) {
      const r = await fetch(`https://lh3.googleusercontent.com/d/${id}=w220`);
      if (r.ok) return r.blob();
      if ((r.status === 429 || r.status >= 500) && intento < 5) { await esperar(1500 * 2 ** intento + Math.random() * 500); continue; }
      throw new Error(r.status);
    }
  }
  async function traerFotos(modelos, imagenes, avance) {
    const img = (imagenes && imagenes.img) || {};
    const pedir = [...new Set(modelos.map(m => m.toUpperCase()))].filter(m => img[m]);
    const out = {}, fallidas = [];
    let hechas = 0;
    const uno = async m => {
      const id = img[m];
      try {
        if (!cacheFotos[id]) {
          const blob = await bajarFoto(id);
          const ext = /png/.test(blob.type) ? 'png' : /gif/.test(blob.type) ? 'gif' : 'jpeg';
          let w = 1, h = 1;
          try { const bmp = await createImageBitmap(blob); w = bmp.width; h = bmp.height; bmp.close && bmp.close(); } catch (e) { /* sin tamaño: cuadrada */ }
          cacheFotos[id] = { buffer: await blob.arrayBuffer(), ext, w, h };
        }
        out[m] = cacheFotos[id];
      } catch (e) { fallidas.push(m); }
      hechas++;
      if (avance) avance(hechas, pedir.length);
    };
    const cola = pedir.slice();
    await Promise.all(Array.from({ length: 3 }, async () => { while (cola.length) await uno(cola.shift()); }));
    // las que no bajaron quedan con un link a la foto en Drive
    const links = {};
    for (const m of fallidas) links[m] = `https://drive.google.com/file/d/${img[m]}/view`;
    return { fotos: out, links, conFoto: Object.keys(out).length, sinCatalogo: modelos.length - pedir.length, fallidas: fallidas.length };
  }

  // fotos = lo que devuelve traerFotos ({ fotos, links }) o null para un Excel sin fotos
  async function excel(ocs, fuente, fotos) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Cargas · Grupo Depor';
    wb.created = new Date();
    hojaResumen(wb, ocs, fuente);
    hojaDatos(wb, ocs);
    for (const o of ocs) if (o.encontrada) hojaOC(wb, o, fotos, fuente);
    return wb.xlsx.writeBuffer();
  }

  global.Armar = { filasDeTexto, parseLista, COLS_PICKING, indexarPicking, armar, sugerencias, clienteDe, nombreArchivo, fechaPicking, traerFotos, excel, SIN_SUPERVISOR };
})(typeof self !== 'undefined' ? self : this);
