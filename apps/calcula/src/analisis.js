// Calcula Datos — parte 3 (como LibreOffice Calc): escenarios (crear, mostrar y resumen), operaciones
// múltiples (tabla de datos de una o dos variables), listas de ordenamiento personalizadas (para
// autorrellenar y ordenar), detective (marcar datos no válidos y rastrear error) e insertar la fecha
// o la hora actual. GPL-3.0
import { HyperFormula } from 'hyperformula';
import { h, toast, t, tf, modal } from '../../shared/office.js';
import { tabbed, field, sel, chk } from '../../shared/dialogs.js';
import { colName, addr, parseAddr, parseRange, rangeStr } from './workbook.js';
import { validate } from './features.js';

let C = null; // { wb, grid, shell, act, refresh, sel, clamp, switchSheet, renderTabs, refreshCharts }
const marks = { invalid: [], errors: [] }; // círculos rojos y flechas de error (por hoja)
export function init(ctx) {
  C = ctx;
  const g = C.grid; const prev = g.overlay;
  g.overlay = (ctx2, ...a) => { prev?.(ctx2, ...a); drawMarks(ctx2); };
}
const X3 = (s = C.wb.sheet) => { if (!s.x3) s.x3 = { sparklines: [], rowGroups: [], colGroups: [], slicers: [] }; return s.x3; };
const v = (body, n) => { const el = body.querySelector(`[name="${n}"]`); return el.type === 'checkbox' ? el.checked : el.value; };

// Copia aparte del libro para calcular «qué pasaría si…» sin tocar el libro ni su historial
function sandbox() {
  const hf = C.wb.hf;
  const sheets = hf.getAllSheetsSerialized();
  const names = (hf.getAllNamedExpressionsSerialized?.() || []);
  const box = HyperFormula.buildFromSheets(sheets, { ...hf.getConfig(), undoLimit: 0 }, names);
  const id = (s) => box.getSheetId(hf.getSheetName(s.id));
  return { box, id, done: () => box.destroy() };
}
const rangeCells = (g) => { const o = []; for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) o.push([r, c]); return o; };

// ================================================================ escenarios
export async function createScenarioDialog() {
  const g = C.clamp(C.sel()); const n = (X3().scenarios || []).length + 1;
  const body = await tabbed('Crear escenario', [{ label: 'Escenario', body: h('div', { class: 'fx-col' },
    field('Nombre del escenario', h('input', { name: 'name', value: tf('Escenario {0}', n) })),
    field('Celdas cambiantes', h('input', { name: 'range', value: rangeStr(g) })),
    field('Comentario', h('textarea', { name: 'comment', rows: 3, style: { width: '100%', fontFamily: 'inherit' } }, tf('Creado el {0}', new Date().toLocaleDateString('es-MX'))))) }], { width: '460px' });
  if (!body) return null;
  return createScenario(v(body, 'name').trim(), v(body, 'range'), v(body, 'comment'));
}
export function createScenario(name, range, comment = '') {
  const g = parseRange(range); if (!g || !name) { toast('Intervalo no válido', { error: true }); return null; }
  const wb = C.wb; const values = [];
  for (let r = g.r1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) row.push(wb.raw(r, c) ?? null); values.push(row); }
  const sc = { name, range: rangeStr(g), values, comment };
  C.act(t('Crear escenario'), () => { const x = X3(); x.scenarios = [...(x.scenarios || []).filter((q) => q.name !== name), sc]; }, { hf: false });
  return sc;
}
export function showScenario(name) {
  const sc = (X3().scenarios || []).find((q) => q.name === name); if (!sc) return false;
  const g = parseRange(sc.range); const wb = C.wb;
  C.act(tf('Escenario: {0}', name), () => sc.values.forEach((row, i) => row.forEach((val, j) => wb.setRaw(g.r1 + i, g.c1 + j, val))));
  C.refreshCharts?.();
  return true;
}
export async function scenariosDialog() {
  const list = h('select', { size: 7, style: { width: '100%' }, name: 'list', 'aria-label': t('Escenarios') });
  const info = h('p', { class: 'fx-note' });
  const fill = () => { const sc = X3().scenarios || []; list.replaceChildren(...sc.map((q) => h('option', { value: q.name }, `${q.name} — ${q.range}`))); if (sc.length) list.value = sc[0].name; upd(); };
  const upd = () => { const q = (X3().scenarios || []).find((x) => x.name === list.value); info.textContent = q?.comment || ''; };
  list.addEventListener('change', upd);
  const btn = (label, fn) => h('button', { type: 'button', class: 'btn', onclick: fn }, t(label));
  const body = h('div', { class: 'cshow' }, h('div', { class: 'fx-col' }, list, info), h('div', { class: 'fx-col' },
    btn('Mostrar', () => { if (list.value) showScenario(list.value); }),
    btn('Agregar…', async () => { await createScenarioDialog(); fill(); }),
    btn('Eliminar', () => { const n = list.value; if (!n) return; C.act(t('Eliminar escenario'), () => { X3().scenarios = (X3().scenarios || []).filter((q) => q.name !== n); }, { hf: false }); fill(); })));
  fill();
  const r = await modal({ title: t('Administrador de escenarios'), body, width: '560px', buttons: [{ label: t('Resumen…'), value: 'summary' }, { label: t('Cerrar'), value: true, primary: true }] });
  if (r === 'summary') return scenarioSummaryDialog();
  return r;
}
export async function scenarioSummaryDialog() {
  const g = C.clamp(C.sel());
  const body = await tabbed('Resumen del escenario', [{ label: 'Resumen', body: h('div', { class: 'fx-col' },
    field('Celdas de resultado', h('input', { name: 'results', value: rangeStr(g) })), h('p', { class: 'fx-note' }, t('Se crea una hoja nueva con los valores de cada escenario y los resultados que producen.'))) }], { width: '460px' });
  if (!body) return null;
  const rg = parseRange(v(body, 'results')); if (!rg) { toast('Intervalo no válido', { error: true }); return null; }
  return scenarioSummary(rg);
}
// hoja nueva: una columna por escenario con sus celdas cambiantes y los resultados calculados
export function scenarioSummary(results) {
  const wb = C.wb; const src = wb.sheet; const scs = X3(src).scenarios || [];
  if (!scs.length) { toast('Esta hoja no tiene escenarios.', { error: true }); return null; }
  const { box, id, done } = sandbox(); const sid = id(src);
  const cols = scs.map((sc) => {
    const g = parseRange(sc.range);
    sc.values.forEach((row, i) => row.forEach((val, j) => box.setCellContents({ sheet: sid, row: g.r1 + i, col: g.c1 + j }, [[val]])));
    const inputs = rangeCells(g).map(([r, c]) => box.getCellValue({ sheet: sid, row: r, col: c }));
    const outs = rangeCells(results).map(([r, c]) => box.getCellValue({ sheet: sid, row: r, col: c }));
    return { name: sc.name, inputs, outs, g };
  });
  done();
  const plainV = (x) => (x && typeof x === 'object' ? String(x.value ?? '#ERR') : x);
  const s = wb.addSheet(tf('Resumen de escenarios {0}', src.name).slice(0, 31));
  const rowsIn = rangeCells(cols[0].g); const rowsOut = rangeCells(results);
  wb.action(t('Resumen de escenarios'), () => {
    const put = (r, c, val, st) => { wb.setRaw(r, c, val ?? null, s); if (st) wb.setStyle(r, c, st, s); };
    put(0, 0, tf('Resumen de escenarios — {0}', src.name), { b: true, size: 13 });
    put(2, 0, t('Celdas cambiantes'), { b: true, bg: '#D9E1F2' });
    cols.forEach((col, k) => put(2, 1 + k, col.name, { b: true, bg: '#D9E1F2', align: 'center' }));
    rowsIn.forEach(([r, c], i) => { put(3 + i, 0, addr(r, c)); cols.forEach((col, k) => put(3 + i, 1 + k, plainV(col.inputs[i]))); });
    const r0 = 4 + rowsIn.length;
    put(r0, 0, t('Celdas de resultado'), { b: true, bg: '#D9E1F2' }); cols.forEach((_, k) => put(r0, 1 + k, null, { bg: '#D9E1F2' }));
    rowsOut.forEach(([r, c], i) => { put(r0 + 1 + i, 0, addr(r, c)); cols.forEach((col, k) => put(r0 + 1 + i, 1 + k, plainV(col.outs[i]))); });
    s.colW.set(0, 150);
  }, { sheets: [s] });
  C.renderTabs(); C.switchSheet(wb.sheets.indexOf(s)); C.shell.markDirty();
  return cols;
}

// ================================================================ operaciones múltiples (tabla de datos)
export async function multipleOpsDialog() {
  const g = C.clamp(C.sel());
  const body = await tabbed('Operaciones múltiples', [{ label: 'Valores predeterminados', body: h('div', { class: 'fx-col' },
    field('Intervalo de la tabla', h('input', { name: 'range', value: rangeStr(g) })),
    field('Fórmulas', h('input', { name: 'formulas', placeholder: t('Por ejemplo: B5 o B5:D5') })),
    field('Celda de entrada de fila', h('input', { name: 'rowCell', placeholder: t('(opcional)') })),
    field('Celda de entrada de columna', h('input', { name: 'colCell', placeholder: t('(opcional)') })),
    h('p', { class: 'fx-note' }, t('Con una columna de valores, la primera columna del intervalo tiene los valores a probar y los resultados se escriben a su derecha (uno por fórmula). Con dos variables, la primera fila y la primera columna tienen los valores y se usa una sola fórmula.'))) }], { width: '520px' });
  if (!body) return null;
  return multipleOps({ range: v(body, 'range'), formulas: v(body, 'formulas'), rowCell: v(body, 'rowCell'), colCell: v(body, 'colCell') });
}
export function multipleOps({ range, formulas, rowCell, colCell }) {
  const wb = C.wb; const g = parseRange(range); const f = parseRange(formulas || '');
  const rc = String(rowCell || '').trim() ? parseAddr(rowCell) : null; const cc = String(colCell || '').trim() ? parseAddr(colCell) : null;
  if (!g || !f || (!rc && !cc)) { toast('Faltan el intervalo, las fórmulas o una celda de entrada.', { error: true }); return null; }
  const { box, id, done } = sandbox(); const sid = id(wb.sheet);
  const at = (p) => ({ sheet: sid, row: p.r, col: p.c });
  const fcells = rangeCells(f).map(([r, c]) => ({ r, c }));
  const eval1 = (assign, fcell) => { for (const [p, val] of assign) box.setCellContents(at(p), [[val]]); return box.getCellValue(at(fcell)); };
  const restore = (cells) => { for (const p of cells) box.setCellContents(at(p), [[wb.raw(p.r, p.c) ?? null]]); };
  const out = []; // [r, c, valor]
  const val = (x) => (x && typeof x === 'object' ? (x.value ?? '#¡VALOR!') : x);
  if (rc && cc) {
    const fc = fcells[0];
    for (let r = g.r1 + 1; r <= g.r2; r++) for (let c = g.c1 + 1; c <= g.c2; c++) {
      out.push([r, c, val(eval1([[cc, wb.value(r, g.c1)], [rc, wb.value(g.r1, c)]], fc))]);
    }
    restore([rc, cc]);
  } else if (cc) {
    for (let r = g.r1; r <= g.r2; r++) fcells.forEach((fc, k) => out.push([r, g.c1 + 1 + k, val(eval1([[cc, wb.value(r, g.c1)]], fc))]));
    restore([cc]);
  } else {
    for (let c = g.c1; c <= g.c2; c++) fcells.forEach((fc, k) => out.push([g.r1 + 1 + k, c, val(eval1([[rc, wb.value(g.r1, c)]], fc))]));
    restore([rc]);
  }
  done();
  C.act(t('Operaciones múltiples'), () => out.forEach(([r, c, x]) => wb.setRaw(r, c, x ?? null)));
  return out.length;
}

// ================================================================ listas de ordenamiento
const LS = 'oficina.calc.lists';
export const BUILTIN_LISTS = [
  'lun,mar,mié,jue,vie,sáb,dom', 'lunes,martes,miércoles,jueves,viernes,sábado,domingo',
  'ene,feb,mar,abr,may,jun,jul,ago,sep,oct,nov,dic', 'enero,febrero,marzo,abril,mayo,junio,julio,agosto,septiembre,octubre,noviembre,diciembre',
  'Mon,Tue,Wed,Thu,Fri,Sat,Sun', 'Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday',
  'Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec', 'January,February,March,April,May,June,July,August,September,October,November,December',
].map((l) => l.split(','));
export function userLists() { try { const j = JSON.parse(localStorage.getItem(LS) || '[]'); return Array.isArray(j) ? j.filter((l) => Array.isArray(l) && l.length > 1) : []; } catch { return []; } }
export function setUserLists(lists) { try { localStorage.setItem(LS, JSON.stringify(lists)); } catch { /* sin almacenamiento */ } }
export const allLists = () => [...userLists(), ...BUILTIN_LISTS];
export async function sortListsDialog() {
  const ta = h('textarea', { name: 'lists', rows: 8, style: { width: '100%', fontFamily: 'inherit' } }, userLists().map((l) => l.join(', ')).join('\n'));
  const fromSel = h('button', { type: 'button', class: 'btn', onclick: () => {
    const g = C.clamp(C.sel()); const items = []; for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { const x = C.wb.display(r, c).text; if (x !== '' && x != null) items.push(String(x)); }
    if (items.length > 1) ta.value = `${ta.value.trim() ? `${ta.value.trim()}\n` : ''}${items.join(', ')}`;
  } }, t('Copiar de la selección'));
  const body = await tabbed('Listas de ordenamiento', [{ label: 'Listas', body: h('div', { class: 'fx-col' },
    h('p', { class: 'fx-note' }, t('Una lista por línea, con los elementos separados por comas. Sirven para autorrellenar (escribe el primero y arrastra) y para ordenar con un orden personalizado. Los días y los meses ya vienen incluidos.')), ta, fromSel) }], { width: '560px' });
  if (!body) return null;
  const lists = ta.value.split('\n').map((l) => l.split(',').map((x) => x.trim()).filter(Boolean)).filter((l) => l.length > 1);
  setUserLists(lists); toast(tf('{0} listas guardadas', lists.length));
  return lists;
}
// posición de un texto en alguna lista (para ordenar con orden personalizado)
export function listRank(list, x) { const k = list.findIndex((w) => w.toLowerCase() === String(x ?? '').toLowerCase()); return k < 0 ? Infinity : k; }

// ================================================================ detective
export function markInvalid() {
  const wb = C.wb; const s = wb.sheet; const out = [];
  for (const vd of s.validations || []) {
    const g = parseRange(vd.range); if (!g) continue;
    const u = wb.usedRange(); const r2 = Math.min(g.r2, u.r2); const c2 = Math.min(g.c2, u.c2);
    for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) {
      const txt = wb.editText(r, c); if (txt === '' || txt == null) continue;
      if (!validate(r, c, String(txt)).ok) out.push({ r, c, sheet: s.id });
    }
  }
  marks.invalid = [...marks.invalid.filter((m) => m.sheet !== s.id), ...out];
  C.grid.draw();
  toast(out.length ? tf('{0} celdas con datos no válidos', out.length) : t('No hay datos no válidos.'));
  return out.length;
}
// desde una celda con error, flechas rojas hasta las celdas con error que la causan
export function traceError() {
  const wb = C.wb; const s = wb.sheet; const { r, c } = C.grid.sel;
  if (!wb.isError(wb.value(r, c))) { toast('La celda no tiene un error.', { error: true }); return 0; }
  const seen = new Set(); const found = [];
  const walk = (rr, cc, depth) => {
    if (depth > 30 || seen.has(`${rr},${cc}`)) return; seen.add(`${rr},${cc}`);
    let pre = []; try { pre = wb.hf.getCellPrecedents({ sheet: s.id, row: rr, col: cc }); } catch { pre = []; }
    for (const a of pre) {
      const cells = a.start ? (a.start.sheet === s.id ? rangeCells({ r1: a.start.row, c1: a.start.col, r2: a.end.row, c2: a.end.col }) : []) : (a.sheet === s.id ? [[a.row, a.col]] : []);
      for (const [pr, pc] of cells) if (wb.isError(wb.value(pr, pc))) { found.push({ from: { r: pr, c: pc }, to: { r: rr, c: cc }, sheet: s.id }); walk(pr, pc, depth + 1); }
    }
  };
  walk(r, c, 0);
  marks.errors = [...marks.errors.filter((m) => m.sheet !== s.id), ...found];
  if (!found.length) marks.invalid.push({ r, c, sheet: s.id, err: true });
  C.grid.draw();
  return found.length;
}
export function clearMarks() { marks.invalid = []; marks.errors = []; C.grid.draw(); }
export const markCount = () => ({ invalid: marks.invalid.length, errors: marks.errors.length });
function drawMarks(ctx) {
  if (!marks.invalid.length && !marks.errors.length) return;
  const s = C.wb.sheet; ctx.save(); ctx.strokeStyle = '#D32F2F'; ctx.fillStyle = '#D32F2F'; ctx.lineWidth = 1.6;
  for (const m of marks.invalid) {
    if (m.sheet !== s.id) continue; const b = C.grid.cellRect(m.r, m.c);
    ctx.beginPath(); ctx.ellipse(b.x + b.w / 2, b.y + b.h / 2, b.w / 2 + 3, b.h / 2 + 3, 0, 0, Math.PI * 2); ctx.stroke();
  }
  for (const a of marks.errors) {
    if (a.sheet !== s.id) continue;
    const f = C.grid.cellRect(a.from.r, a.from.c); const tt = C.grid.cellRect(a.to.r, a.to.c);
    const x1 = f.x + f.w / 2, y1 = f.y + f.h / 2, x2 = tt.x + tt.w / 2, y2 = tt.y + tt.h / 2;
    ctx.beginPath(); ctx.arc(x1, y1, 3, 0, Math.PI * 2); ctx.fill();
    ctx.beginPath(); ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); ctx.stroke();
    const ang = Math.atan2(y2 - y1, x2 - x1);
    ctx.beginPath(); ctx.moveTo(x2, y2); ctx.lineTo(x2 - 9 * Math.cos(ang - 0.4), y2 - 9 * Math.sin(ang - 0.4)); ctx.lineTo(x2 - 9 * Math.cos(ang + 0.4), y2 - 9 * Math.sin(ang + 0.4)); ctx.closePath(); ctx.fill();
  }
  ctx.restore();
}

// ================================================================ fecha y hora actuales (Ctrl+; y Ctrl+Mayús+;)
const EPOCH = Date.UTC(1899, 11, 30);
export function insertNow(kind = 'date') {
  const wb = C.wb; const { r, c } = C.grid.sel; const d = new Date();
  const local = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds());
  const serial = kind === 'date' ? Math.floor((local - EPOCH) / 86400000) : ((local - EPOCH) / 86400000) % 1;
  C.act(kind === 'date' ? t('Insertar fecha') : t('Insertar hora'), () => { wb.setRaw(r, c, Math.round(serial * 1e8) / 1e8); wb.setStyle(r, c, { fmt: kind === 'date' ? 'dd/mm/yyyy' : 'hh:mm:ss' }); });
  return serial;
}
void colName; void chk; void sel;
