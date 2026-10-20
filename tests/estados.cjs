// Run: node tests/estados.cjs. No network requests or production writes.
const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const html = fs.readFileSync(__dirname + '/../index.html', 'utf8');
const script = html.match(/<script>\s*([\s\S]*?)<\/script>/)[1].replace(/\ncargar\(\);\s*$/, '');
const nodes = {};
function node(id) {
  return nodes[id] ||= {
    value: '', innerHTML: '', textContent: '', style: {}, classList: { add() {}, remove() {}, toggle() {} },
    addEventListener() {}, querySelectorAll() { return []; }, appendChild() {}, setAttribute() {}
  };
}
const context = {
  document: { getElementById: node, querySelectorAll: () => [], addEventListener() {}, createElement: () => node('created') },
  localStorage: { getItem: () => null, setItem() {}, removeItem() {} },
  navigator: {}, location: { protocol: 'https:' }, setTimeout: () => 1, clearTimeout() {},
  setInterval: () => 1, clearInterval() {}, AbortController, scrollTo() {}
};
context.window = context;
vm.createContext(context);
vm.runInContext(script, context);
const c = context;
const reports = { ok: true, reportes: [
  { oc: '100001', cod: '1', sup: 'Ana', fecha: '2026-10-01T12:00:00Z', estado: 'nollego', comentario: '<anterior>' },
  { oc: '100001', cod: '1', sup: 'Ana', fecha: '2026-10-02T12:00:00Z', estado: 'completa' },
  { oc: '100001', cod: '2', sup: 'Ana', fecha: '2026-10-02T12:00:00Z', estado: 'incompleta' }
] };
c.MAN = { actualizado: '2026-10-20' };
c.OCS = [{ oc: '100001', cliente: 'Paris', tiendas: ['1','2','3'].map(cod => ({cod, nombre: 'Tienda '+cod, sup: 'Ana', uds: 10, modelos: 1})) }];
c.RESP = c.leerReportes(reports);
c.SUP = 'ANA';
c.cruzar();
assert.deepEqual(Array.from(c.ITEMS, i => i.estado), ['ok', 'obs', 'pend']);
assert.equal(c.ITEMS.length, 3, 'Multiple reports must not create extra loads');
assert.equal(c.ITEMS[0].reps.length, 2, 'Keep historical reports');
const history = c.historialReportes(c.ITEMS[0]);
assert(history.includes('&lt;anterior&gt;'));
assert(history.indexOf('Llegó completa') < history.indexOf('No llegó'));
c.VISTA = 'res'; c.render(); c.actualizarMeta();
assert.equal(node('totalOC').textContent, '1');
assert.equal(node('totalCargas').textContent, '3');
assert.equal(node('totalEnvios').textContent, '3');
assert(node('totalReportadas').textContent.includes('2 cargas distintas'));
assert.equal(node('rkPend').textContent, '1');

(async () => {
  c.API_ERR = true;
  c.cruzar();
  assert(c.ITEMS.every(i => i.estado === 'unknown'), 'Failure must never fabricate pending or complete states');
  node('rEst').value = 'ok';
  c.render();
  assert.equal(c.itemsRes().length, 3, 'A stale state filter must not conceal unknown loads');
  for (const id of ['rkAv','rkPend','rkOk','rkObs','totalEnvios']) assert.equal(node(id).textContent, '—');
  assert(!node('rOcs').innerHTML.includes('Todo ok'));
  assert(node('rOcs').innerHTML.includes('Sin confirmar'));
  assert(node('estadoConsulta').innerHTML.includes('Reintentar'));
  c.VISTA = 'arch'; c.render();
  assert(node('sLista').innerHTML.includes('No se pueden confirmar'));
  c.VISTA = 'sup'; c.render();
  assert(node('sLista').innerHTML.includes('Sin confirmar'));
  assert(!node('sLista').innerHTML.includes('Reportar llegada'));
  c.fetch = async () => { throw Error('offline'); };
  await c.reintentarReportes();
  assert.equal(c.API_ERR, true);
  assert.equal(c.REP_LOADING, false);
  assert(c.ITEMS.every(i => i.estado === 'unknown'));
  c.fetch = async () => ({ ok: true, json: async () => reports });
  await c.reintentarReportes();
  assert.equal(c.API_ERR, false);
  assert.deepEqual(Array.from(c.ITEMS, i => i.estado), ['ok', 'obs', 'pend']);
  assert.equal(node('rEst').value, 'ok', 'Retry preserves filters');
  assert(node('sLista').innerHTML.includes('Reportar llegada'));
  assert(!node('estadoConsulta').innerHTML.includes('Reintentar'));
  assert.equal(c.RESP.length, 3);
  console.log('PASS: counts, latest report, history, escaping, failure views, failed retry and recovery.');
})().catch(e => { console.error(e); process.exitCode = 1; });
