// Calcula Datos — Obtener y transformar datos (tipo Power Query): una consulta lee un origen
// (CSV, libro de Excel, JSON o un rango del libro), le aplica pasos guardados y carga el resultado
// en una hoja. "Actualizar" vuelve a leer el origen y repite los pasos. GPL-3.0
import ExcelJS from 'exceljs';
import { h, toast, modal, form, dropdown, esc, api, t, tf } from '../../shared/office.js';
import { parseCsv } from './csv.js';
import { parseRange, rangeStr } from './workbook.js';

let C = null; // { wb, grid, shell, act, refresh, switchSheet, refreshCharts }
export function init(ctx) { C = ctx; if (!C.wb.queries) C.wb.queries = []; }
const uid = () => `q${Math.random().toString(36).slice(2, 8)}`;

// ================================================================ lectura de orígenes
const cellVal = (v) => {
  if (v == null) return null;
  if (typeof v === 'object') {
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    if ('result' in v) return cellVal(v.result);
    if (v.richText) return v.richText.map((r) => r.text).join('');
    if ('text' in v) return v.text;
    if (v.error) return v.error;
    return null;
  }
  return v;
};
function guess(v) {
  if (typeof v !== 'string') return v;
  const s = v.trim(); if (s === '') return null;
  if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
  if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) return Number(s.replace(/,/g, ''));
  if (/^(true|verdadero)$/i.test(s)) return true; if (/^(false|falso)$/i.test(s)) return false;
  return v;
}
function toTable(rows, headers = true) {
  const width = Math.max(0, ...rows.map((r) => r.length));
  const pad = (r) => Array.from({ length: width }, (_, i) => (r[i] === undefined ? null : r[i]));
  let cols; let body;
  if (headers && rows.length) { cols = pad(rows[0]).map((x, i) => String(x ?? '').trim() || `${t('Columna')}${i + 1}`); body = rows.slice(1).map(pad); }
  else { cols = Array.from({ length: width }, (_, i) => `${t('Columna')}${i + 1}`); body = rows.map(pad); }
  // nombres únicos
  const seen = {}; cols = cols.map((c) => { if (seen[c] == null) { seen[c] = 1; return c; } seen[c]++; return `${c}_${seen[c]}`; });
  return { cols, rows: body };
}
function flatten(o, pre = '', out = {}) {
  for (const [k, v] of Object.entries(o || {})) { const key = pre ? `${pre}.${k}` : k; if (v && typeof v === 'object' && !Array.isArray(v)) flatten(v, key, out); else out[key] = Array.isArray(v) ? JSON.stringify(v) : v; }
  return out;
}
export async function readSource(src, bytes) {
  if (src.kind === 'range') {
    const s = C.wb.sheets.find((x) => x.name === src.sheet) || C.wb.sheet; const g = parseRange(src.range);
    if (!g) throw new Error(t('Rango no válido'));
    const rows = []; for (let r = g.r1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) { const v = C.wb.value(r, c, s); row.push(v && typeof v === 'object' ? null : v); } rows.push(row); }
    return toTable(rows, src.headers !== false);
  }
  if (!bytes) { const f = await api.readData(src.path); if (!f?.bytes) throw new Error(tf('No se encontró el archivo de origen: {0}', src.path)); bytes = f.bytes; }
  if (src.kind === 'csv') {
    const text = new TextDecoder(src.encoding || 'utf-8').decode(bytes);
    return toTable(parseCsv(text).map((r) => r.map(guess)), src.headers !== false);
  }
  if (src.kind === 'json') {
    let d = JSON.parse(new TextDecoder().decode(bytes));
    if (!Array.isArray(d)) { const arr = Object.values(d || {}).find(Array.isArray); d = arr || [d]; }
    const objs = d.map((x) => (x && typeof x === 'object' && !Array.isArray(x) ? flatten(x) : { valor: x }));
    const cols = [...new Set(objs.flatMap((o) => Object.keys(o)))];
    return { cols, rows: objs.map((o) => cols.map((c) => (o[c] === undefined ? null : o[c]))) };
  }
  if (src.kind === 'xlsx') {
    const x = new ExcelJS.Workbook(); await x.xlsx.load(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
    const ws = (src.sheet && x.getWorksheet(src.sheet)) || x.worksheets.find((w) => w.name !== '_calcula_datos');
    if (!ws) throw new Error(t('El libro no tiene hojas'));
    const rows = []; ws.eachRow({ includeEmpty: true }, (row, rn) => { const r = []; row.eachCell({ includeEmpty: true }, (cell, cn) => { r[cn - 1] = cellVal(cell.value); }); rows[rn - 1] = r; });
    return toTable(Array.from(rows, (r) => r || []), src.headers !== false);
  }
  throw new Error(t('Origen desconocido'));
}

// ================================================================ pasos
const isNum = (v) => typeof v === 'number' && Number.isFinite(v);
const cmp = (a, b) => (a == null ? (b == null ? 0 : 1) : b == null ? -1 : isNum(a) && isNum(b) ? a - b : String(a).localeCompare(String(b), 'es', { numeric: true, sensitivity: 'base' }));
// Expresión de columna personalizada: [Col] * 2, [A] & " " & [B], números, + - * / ( ), texto entre comillas
export function compileExpr(expr, cols) {
  const toks = []; const s = String(expr); let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (/\s/.test(ch)) { i++; continue; }
    if (ch === '[') { const j = s.indexOf(']', i); if (j < 0) throw new Error(t('Falta ]')); const name = s.slice(i + 1, j); const k = cols.indexOf(name); if (k < 0) throw new Error(tf('No existe la columna {0}', name)); toks.push({ t: 'col', k }); i = j + 1; continue; }
    if (ch === '"') { const j = s.indexOf('"', i + 1); if (j < 0) throw new Error(t('Falta "')); toks.push({ t: 'str', v: s.slice(i + 1, j) }); i = j + 1; continue; }
    const m = /^\d+(\.\d+)?/.exec(s.slice(i)); if (m) { toks.push({ t: 'num', v: Number(m[0]) }); i += m[0].length; continue; }
    if ('+-*/()&'.includes(ch)) { toks.push({ t: 'op', v: ch }); i++; continue; }
    throw new Error(tf('Carácter no válido: {0}', ch));
  }
  let p = 0;
  const peek = () => toks[p]; const eat = (v) => { if (peek()?.v === v) { p++; return true; } return false; };
  const prim = () => { const tk = toks[p++]; if (!tk) throw new Error(t('Expresión incompleta')); if (tk.t === 'num' || tk.t === 'str') return () => tk.v; if (tk.t === 'col') return (row) => row[tk.k]; if (tk.v === '(') { const e = cat(); if (!eat(')')) throw new Error(t('Falta )')); return e; } if (tk.v === '-') { const e = prim(); return (row) => -Number(e(row)); } throw new Error(t('Expresión no válida')); };
  const mul = () => { let a = prim(); for (;;) { if (eat('*')) { const l = a, r = prim(); a = (row) => Number(l(row)) * Number(r(row)); } else if (eat('/')) { const l = a, r = prim(); a = (row) => { const d = Number(r(row)); return d === 0 ? null : Number(l(row)) / d; }; } else return a; } };
  const add = () => { let a = mul(); for (;;) { if (eat('+')) { const l = a, r = mul(); a = (row) => Number(l(row)) + Number(r(row)); } else if (eat('-')) { const l = a, r = mul(); a = (row) => Number(l(row)) - Number(r(row)); } else return a; } };
  const cat = () => { let a = add(); while (eat('&')) { const l = a, r = add(); a = (row) => `${l(row) ?? ''}${r(row) ?? ''}`; } return a; };
  const fn = cat(); if (p < toks.length) throw new Error(t('Expresión no válida'));
  return (row) => { const v = fn(row); return typeof v === 'number' && !Number.isFinite(v) ? null : v; };
}
const OPS = { '=': (a, b) => cmp(a, b) === 0, '<>': (a, b) => cmp(a, b) !== 0, '>': (a, b) => cmp(a, b) > 0, '<': (a, b) => cmp(a, b) < 0, '>=': (a, b) => cmp(a, b) >= 0, '<=': (a, b) => cmp(a, b) <= 0, contains: (a, b) => String(a ?? '').toLowerCase().includes(String(b).toLowerCase()), notContains: (a, b) => !String(a ?? '').toLowerCase().includes(String(b).toLowerCase()), empty: (a) => a == null || a === '', notEmpty: (a) => !(a == null || a === '') };
export function applyStep(tb, st) {
  const ci = (name) => { const k = tb.cols.indexOf(name); if (k < 0) throw new Error(tf('No existe la columna {0}', name)); return k; };
  const map = (col, f) => { const k = ci(col); return { cols: tb.cols, rows: tb.rows.map((r) => { const n = r.slice(); n[k] = f(r[k]); return n; }) }; };
  switch (st.op) {
    case 'removeCols': { const ks = st.cols.map(ci); return { cols: tb.cols.filter((_, i) => !ks.includes(i)), rows: tb.rows.map((r) => r.filter((_, i) => !ks.includes(i))) }; }
    case 'keepCols': { const ks = st.cols.map(ci); return { cols: ks.map((k) => tb.cols[k]), rows: tb.rows.map((r) => ks.map((k) => r[k])) }; }
    case 'rename': { const k = ci(st.from); const cols = tb.cols.slice(); cols[k] = st.to; return { cols, rows: tb.rows }; }
    case 'filter': { const k = ci(st.col); const f = OPS[st.cmp]; const v = guess(String(st.value ?? '')); return { cols: tb.cols, rows: tb.rows.filter((r) => f(r[k], v ?? st.value)) }; }
    case 'filterValues': { const k = ci(st.col); const set = new Set(st.values.map(String)); return { cols: tb.cols, rows: tb.rows.filter((r) => set.has(String(r[k] ?? ''))) }; }
    case 'sort': { const k = ci(st.col); return { cols: tb.cols, rows: tb.rows.slice().sort((a, b) => (st.desc ? -1 : 1) * cmp(a[k], b[k])) }; }
    case 'dedupe': { const ks = (st.cols?.length ? st.cols : tb.cols).map(ci); const seen = new Set(); return { cols: tb.cols, rows: tb.rows.filter((r) => { const key = JSON.stringify(ks.map((k) => r[k])); if (seen.has(key)) return false; seen.add(key); return true; }) }; }
    case 'type': return map(st.col, (v) => {
      if (v == null || v === '') return null;
      if (st.type === 'number') { const n = typeof v === 'number' ? v : Number(String(v).replace(/[$,\s%]/g, '')); return Number.isFinite(n) ? n : null; }
      if (st.type === 'int') { const n = Math.round(typeof v === 'number' ? v : Number(String(v).replace(/[$,\s]/g, ''))); return Number.isFinite(n) ? n : null; }
      if (st.type === 'date') { const m = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/.exec(String(v).trim()); if (m) { const y = m[3].length === 2 ? `20${m[3]}` : m[3]; return `${y}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`; } const d = new Date(v); return Number.isNaN(d.getTime()) ? v : d.toISOString().slice(0, 10); }
      return String(v);
    });
    case 'replace': return map(st.col, (v) => (v == null ? v : guess(String(v).split(st.find).join(st.repl ?? ''))));
    case 'trim': return map(st.col, (v) => (typeof v === 'string' ? v.trim().replace(/\s+/g, ' ') : v));
    case 'upper': return map(st.col, (v) => (typeof v === 'string' ? v.toLocaleUpperCase('es') : v));
    case 'lower': return map(st.col, (v) => (typeof v === 'string' ? v.toLocaleLowerCase('es') : v));
    case 'proper': return map(st.col, (v) => (typeof v === 'string' ? v.toLocaleLowerCase('es').replace(/(^|\s)\p{L}/gu, (m) => m.toLocaleUpperCase('es')) : v));
    case 'fillDown': { const k = ci(st.col); let last = null; return { cols: tb.cols, rows: tb.rows.map((r) => { const n = r.slice(); if (n[k] == null || n[k] === '') n[k] = last; else last = n[k]; return n; }) }; }
    case 'split': {
      const k = ci(st.col); const parts = tb.rows.map((r) => (r[k] == null ? [] : String(r[k]).split(st.delim)));
      const n = Math.max(1, ...parts.map((p) => p.length)); const names = Array.from({ length: n }, (_, i) => `${tb.cols[k]}.${i + 1}`);
      return { cols: [...tb.cols.slice(0, k), ...names, ...tb.cols.slice(k + 1)], rows: tb.rows.map((r, j) => [...r.slice(0, k), ...names.map((_, i) => guess(parts[j][i] ?? null)), ...r.slice(k + 1)]) };
    }
    case 'addCol': { const f = compileExpr(st.expr, tb.cols); return { cols: [...tb.cols, st.name], rows: tb.rows.map((r) => [...r, f(r)]) }; }
    case 'removeEmpty': return { cols: tb.cols, rows: tb.rows.filter((r) => r.some((v) => v != null && v !== '')) };
    case 'top': return { cols: tb.cols, rows: tb.rows.slice(0, Math.max(0, st.n | 0)) };
    case 'skip': return { cols: tb.cols, rows: tb.rows.slice(Math.max(0, st.n | 0)) };
    case 'group': {
      const ks = st.by.map(ci); const groups = new Map();
      for (const r of tb.rows) { const key = JSON.stringify(ks.map((k) => r[k])); if (!groups.has(key)) groups.set(key, { keys: ks.map((k) => r[k]), rows: [] }); groups.get(key).rows.push(r); }
      const aggs = st.aggs.map((a) => ({ ...a, k: a.fn === 'count' ? -1 : ci(a.col) }));
      const calc = (a, rows) => { if (a.fn === 'count') return rows.length; const vs = rows.map((r) => r[a.k]).filter(isNum); if (a.fn === 'sum') return vs.reduce((x, y) => x + y, 0); if (!vs.length) return null; if (a.fn === 'avg') return vs.reduce((x, y) => x + y, 0) / vs.length; if (a.fn === 'min') return Math.min(...vs); if (a.fn === 'max') return Math.max(...vs); if (a.fn === 'distinct') return new Set(rows.map((r) => r[a.k])).size; return null; };
      return { cols: [...ks.map((k) => tb.cols[k]), ...aggs.map((a) => a.as)], rows: [...groups.values()].map((g) => [...g.keys, ...aggs.map((a) => calc(a, g.rows))]) };
    }
    case 'append': {
      const other = st.__data; if (!other) return tb;
      const cols = [...tb.cols, ...other.cols.filter((c) => !tb.cols.includes(c))];
      const pick = (src, r) => cols.map((c) => { const k = src.cols.indexOf(c); return k < 0 ? null : r[k]; });
      return { cols, rows: [...tb.rows.map((r) => pick(tb, r)), ...other.rows.map((r) => pick(other, r))] };
    }
    case 'merge': {
      const other = st.__data; if (!other) return tb;
      const k1 = ci(st.left); const k2 = other.cols.indexOf(st.right); if (k2 < 0) throw new Error(tf('No existe la columna {0}', st.right));
      const idx = new Map(); other.rows.forEach((r) => { const key = String(r[k2] ?? ''); if (!idx.has(key)) idx.set(key, r); });
      const add = other.cols.map((c, i) => [c, i]).filter(([, i]) => i !== k2); const names = add.map(([c]) => (tb.cols.includes(c) ? `${st.query}.${c}` : c));
      const rows = []; for (const r of tb.rows) { const m = idx.get(String(r[k1] ?? '')); if (!m && st.kind === 'inner') continue; rows.push([...r, ...add.map(([, i]) => (m ? m[i] : null))]); }
      return { cols: [...tb.cols, ...names], rows };
    }
    default: return tb;
  }
}
export const STEP_LABEL = (st) => ({
  removeCols: () => tf('Columnas quitadas: {0}', st.cols.join(', ')), keepCols: () => tf('Columnas conservadas: {0}', st.cols.join(', ')), rename: () => tf('Cambiar nombre: {0} → {1}', st.from, st.to),
  filter: () => tf('Filtrar {0} {1} {2}', st.col, { contains: t('contiene'), notContains: t('no contiene'), empty: t('vacío'), notEmpty: t('no vacío') }[st.cmp] || st.cmp, st.value ?? ''), filterValues: () => tf('Filtrar {0}: {1} valores', st.col, st.values.length),
  sort: () => tf('Ordenar por {0} {1}', st.col, st.desc ? '↓' : '↑'), dedupe: () => t('Quitar duplicados'), type: () => tf('Tipo de {0}: {1}', st.col, { number: t('número'), int: t('entero'), date: t('fecha'), text: t('texto') }[st.type]),
  replace: () => tf('Reemplazar en {0}', st.col), trim: () => tf('Recortar {0}', st.col), upper: () => tf('Mayúsculas en {0}', st.col), lower: () => tf('Minúsculas en {0}', st.col), proper: () => tf('Nombre propio en {0}', st.col),
  fillDown: () => tf('Rellenar hacia abajo {0}', st.col), split: () => tf('Dividir {0} por "{1}"', st.col, st.delim), addCol: () => tf('Columna {0} = {1}', st.name, st.expr), removeEmpty: () => t('Quitar filas vacías'),
  top: () => tf('Conservar las primeras {0} filas', st.n), skip: () => tf('Quitar las primeras {0} filas', st.n), group: () => tf('Agrupar por {0}', st.by.join(', ')), append: () => tf('Anexar {0}', st.query), merge: () => tf('Combinar con {0}', st.query),
}[st.op] || (() => st.op))();

// Ejecuta una consulta completa
export async function runQuery(q, { bytes = null, upTo = Infinity, stack = [] } = {}) {
  let tb = await readSource(q.source, bytes);
  for (let i = 0; i < Math.min(q.steps.length, upTo); i++) {
    const st = q.steps[i];
    if (st.op === 'append' || st.op === 'merge') {
      const other = C.wb.queries.find((x) => x.name === st.query);
      if (!other || stack.includes(other.id) || other.id === q.id) throw new Error(tf('No se encontró la consulta {0}', st.query));
      st.__data = await runQuery(other, { stack: [...stack, q.id] });
    }
    tb = applyStep(tb, st);
    delete st.__data;
  }
  return tb;
}

// ================================================================ cargar en una hoja
function writeTable(q, tb) {
  const wb = C.wb;
  let s = wb.sheets.find((x) => x.id === q.sheetId) || wb.sheets.find((x) => x.name === q.sheetName);
  if (!s) s = wb.addSheet(q.name.replace(/[\\/?*[\]:]/g, ' ').slice(0, 31));
  q.sheetId = s.id; q.sheetName = s.name;
  // fechas ISO → número de serie con formato de fecha; textos que parecen fórmula → literal
  const dateCols = new Set();
  const val = (v, c) => {
    if (typeof v === 'string') {
      const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
      if (m) { dateCols.add(c); return Math.round((Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 864e5); }
      if (/^[=+\-@']/.test(v)) return `'${v}`;
    }
    return typeof v === 'boolean' ? v : v;
  };
  const rows = [tb.cols.map((c) => (/^[=+\-@']/.test(c) ? `'${c}` : c)), ...tb.rows.map((r) => r.map(val))];
  C.act(tf('Cargar {0}', q.name), () => {
    const old = wb.hf.getSheetDimensions(s.id);
    wb.hf.setSheetContent(s.id, rows);
    // limpia el formato de las filas que ya no existen
    for (const k of [...s.styles.keys()]) { const [r] = k.split(',').map(Number); if (r >= rows.length && r < old.height) s.styles.delete(k); }
    for (let c = 0; c < tb.cols.length; c++) {
      wb.setStyle(0, c, { b: true, fill: '#DFF5EE', color: '#0A4A3A' }, s);
      if (dateCols.has(c)) for (let r = 1; r < rows.length; r++) { if (typeof rows[r][c] === 'number') wb.setStyle(r, c, { fmt: 'dd/mm/yyyy' }, s); }
      if (!s.colW.has(c)) s.colW.set(c, Math.min(260, Math.max(70, 8 * Math.max(String(tb.cols[c]).length, ...tb.rows.slice(0, 50).map((r) => String(r[c] ?? '').length)) + 18)));
    }
    s.freeze = { rows: 1, cols: 0 };
    s.nRows = Math.max(1000, rows.length + 200); s.nCols = Math.max(52, tb.cols.length + 20);
    q.lastRows = tb.rows.length; q.loadedAt = new Date().toISOString();
  }, { batch: true, sheets: [s] });
  C.switchSheet(wb.sheets.indexOf(s));
  C.refreshCharts?.();
}
export async function refreshQuery(q, quiet = false) {
  try { const tb = await runQuery(q); writeTable(q, tb); if (!quiet) toast(tf('Consulta {0} actualizada: {1} filas', q.name, tb.rows.length)); return true; } catch (e) { toast(tf('No se pudo actualizar {0}: {1}', q.name, e.message), { error: true, ms: 7000 }); return false; }
}
export async function refreshAll() {
  let n = 0; for (const q of C.wb.queries) if (await refreshQuery(q, true)) n++;
  toast(tf('{0} consultas actualizadas', n));
}

// ================================================================ obtener datos
const KIND = { csv: ['csv', 'txt', 'tsv'], xlsx: ['xlsx', 'xlsm'], json: ['json'] };
export async function getData(kind, path = null) {
  let source; let bytes = null; let name = '';
  if (kind === 'range') {
    const g = C.grid.range; const reg = g.r1 === g.r2 && g.c1 === g.c2 ? C.wb.region(g.r1, g.c1) : g;
    source = { kind: 'range', sheet: C.wb.sheet.name, range: rangeStr(reg), headers: true }; name = `${C.wb.sheet.name}_${rangeStr(reg).replace(':', '_')}`;
  } else {
    const f = path ? await api.readData(path) : await api.pickData({ title: t('Obtener datos'), filters: [{ name: { csv: 'CSV / TXT', xlsx: 'Excel', json: 'JSON' }[kind], extensions: KIND[kind] }] });
    if (!f?.bytes) { if (path) toast(tf('No se encontró el archivo de origen: {0}', path), { error: true }); return; }
    bytes = f.bytes;
    source = { kind, path: f.path, headers: true }; name = f.name.replace(/\.[^.]+$/, '');
    if (kind === 'xlsx') {
      const x = new ExcelJS.Workbook(); await x.xlsx.load(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
      const sheets = x.worksheets.map((w) => w.name).filter((n) => n !== '_calcula_datos');
      if (sheets.length > 1) { const v = await form('Obtener datos', [{ name: 'sheet', label: 'Hoja del libro', type: 'select', value: sheets[0], options: sheets.map((s) => [s, s]) }], 'Aceptar'); if (!v) return; source.sheet = v.sheet; } else source.sheet = sheets[0];
      name = `${name}_${source.sheet}`;
    }
  }
  let base = name.replace(/[^\p{L}\p{N}_]+/gu, '_').slice(0, 40) || 'Consulta'; let k = 1; let nm = base;
  while (C.wb.queries.some((q) => q.name === nm)) nm = `${base}_${++k}`;
  const q = { id: uid(), name: nm, source, steps: [] };
  return editQuery(q, bytes, true);
}

// ================================================================ editor de consultas
export async function editQuery(q0, bytes = null, isNew = false) {
  const q = JSON.parse(JSON.stringify(q0));
  let sel = q.steps.length; // vista previa después de este número de pasos
  let cur = null; let err = '';
  const stepsEl = h('ol', { class: 'qe-steps no-i18n' });
  const grid = h('div', { class: 'qe-grid no-i18n' });
  const info = h('div', { class: 'qe-info muted' });
  const nameIn = h('input', { class: 'qe-name', value: q.name, spellcheck: 'false' });
  const hdrChk = h('input', { type: 'checkbox', checked: q.source.headers !== false ? true : null });
  const redraw = async () => {
    try { cur = await runQuery(q, { bytes, upTo: sel }); err = ''; } catch (e) { cur = null; err = e.message; }
    stepsEl.replaceChildren(h('li', { class: `qe-step${sel === 0 ? ' on' : ''}`, onclick: () => { sel = 0; redraw(); } }, h('span', {}, t('Origen')), h('small', {}, q.source.kind === 'range' ? q.source.range : (q.source.path || '').split(/[\\/]/).pop())),
      ...q.steps.map((st, i) => h('li', { class: `qe-step${sel === i + 1 ? ' on' : ''}`, onclick: () => { sel = i + 1; redraw(); } }, h('span', {}, STEP_LABEL(st)), h('button', { class: 'qe-del', title: t('Quitar paso'), onclick: (e) => { e.stopPropagation(); q.steps.splice(i, 1); sel = Math.min(sel, q.steps.length); redraw(); } }, '✕'))));
    if (!cur) { grid.replaceChildren(h('div', { class: 'qe-err' }, err)); info.textContent = ''; return; }
    const show = cur.rows.slice(0, 200);
    const tbl = h('table', {}, h('thead', {}, h('tr', {}, ...cur.cols.map((c) => h('th', { title: t('Clic para transformar esta columna'), onclick: (e) => colMenu(e.currentTarget, c) }, c, h('span', { class: 'qe-caret' }, '▾'))))),
      h('tbody', {}, ...show.map((r) => h('tr', {}, ...r.map((v) => h('td', { class: typeof v === 'number' ? 'num' : v == null ? 'nul' : '' }, v == null ? 'null' : String(v)))))));
    grid.replaceChildren(tbl);
    info.textContent = tf('{0} columnas · {1} filas', cur.cols.length, cur.rows.length) + (cur.rows.length > 200 ? ` · ${t('vista previa de 200')}` : '');
  };
  const add = (st) => { q.steps.splice(sel, 0, st); sel += 1; redraw(); };
  const colMenu = (anchor, col) => dropdown(anchor, [
    { label: 'Quitar columna', run: () => add({ op: 'removeCols', cols: [col] }) },
    { label: 'Quitar otras columnas', run: () => add({ op: 'keepCols', cols: [col] }) },
    { label: 'Cambiar nombre…', run: async () => { const v = await form('Cambiar nombre', [{ name: 'to', label: 'Nombre nuevo', value: col }], 'Aceptar'); if (v?.to && v.to !== col) add({ op: 'rename', from: col, to: v.to }); } },
    '-',
    { label: 'Ordenar ascendente', run: () => add({ op: 'sort', col, desc: false }) },
    { label: 'Ordenar descendente', run: () => add({ op: 'sort', col, desc: true }) },
    { label: 'Filtrar…', run: async () => { const v = await form('Filtrar filas', [{ name: 'cmp', label: tf('Conservar las filas donde {0}', col), type: 'select', value: '=', options: [['=', 'es igual a'], ['<>', 'no es igual a'], ['>', 'es mayor que'], ['>=', 'es mayor o igual que'], ['<', 'es menor que'], ['<=', 'es menor o igual que'], ['contains', 'contiene'], ['notContains', 'no contiene'], ['notEmpty', 'no está vacío'], ['empty', 'está vacío']] }, { name: 'value', label: 'Valor', value: '' }], 'Aceptar'); if (v) add({ op: 'filter', col, cmp: v.cmp, value: v.value }); } },
    { label: 'Elegir valores…', run: async () => { const k = cur.cols.indexOf(col); const vals = [...new Set(cur.rows.map((r) => String(r[k] ?? '')))].sort((a, b) => a.localeCompare(b, 'es', { numeric: true })).slice(0, 300); const box = h('div', { class: 'qe-vals no-i18n' }, ...vals.map((v) => h('label', {}, h('input', { type: 'checkbox', checked: true, value: v }), ` ${v || t('(Vacías)')}`))); const ok = await modal({ title: tf('Valores de {0}', col), body: box, buttons: [{ label: 'Cancelar', value: null }, { label: 'Aceptar', primary: true, value: () => [...box.querySelectorAll('input:checked')].map((i) => i.value) }] }); if (ok && ok.length < vals.length) add({ op: 'filterValues', col, values: ok }); } },
    '-',
    { label: 'Tipo: número', run: () => add({ op: 'type', col, type: 'number' }) },
    { label: 'Tipo: número entero', run: () => add({ op: 'type', col, type: 'int' }) },
    { label: 'Tipo: fecha', run: () => add({ op: 'type', col, type: 'date' }) },
    { label: 'Tipo: texto', run: () => add({ op: 'type', col, type: 'text' }) },
    '-',
    { label: 'Reemplazar valores…', run: async () => { const v = await form('Reemplazar valores', [{ name: 'find', label: 'Buscar', value: '' }, { name: 'repl', label: 'Reemplazar con', value: '' }], 'Aceptar'); if (v && v.find) add({ op: 'replace', col, find: v.find, repl: v.repl }); } },
    { label: 'Dividir columna…', run: async () => { const v = await form('Dividir columna', [{ name: 'delim', label: 'Delimitador', value: ',' }], 'Aceptar'); if (v && v.delim) add({ op: 'split', col, delim: v.delim }); } },
    { label: 'Recortar espacios', run: () => add({ op: 'trim', col }) },
    { label: 'MAYÚSCULAS', run: () => add({ op: 'upper', col }) },
    { label: 'minúsculas', run: () => add({ op: 'lower', col }) },
    { label: 'Nombre Propio', run: () => add({ op: 'proper', col }) },
    { label: 'Rellenar hacia abajo', run: () => add({ op: 'fillDown', col }) },
    '-',
    { label: 'Agrupar por esta columna…', run: () => groupDialog([col]) },
  ]);
  const groupDialog = async (by) => {
    const nums = cur.cols.filter((c) => !by.includes(c));
    const v = await form('Agrupar por', [
      { name: 'by', label: 'Agrupar por (columnas separadas por coma)', value: by.join(', ') },
      { name: 'fn', label: 'Operación', type: 'select', value: 'sum', options: [['sum', 'Suma'], ['count', 'Contar filas'], ['avg', 'Promedio'], ['min', 'Mínimo'], ['max', 'Máximo'], ['distinct', 'Contar valores distintos']] },
      { name: 'col', label: 'Columna', type: 'select', value: nums[0] || '', options: nums.map((c) => [c, c]) },
      { name: 'as', label: 'Nombre de la columna nueva', value: '' },
    ], 'Aceptar');
    if (!v) return;
    const byCols = v.by.split(',').map((x) => x.trim()).filter(Boolean);
    const fnName = { sum: t('Suma'), count: t('Filas'), avg: t('Promedio'), min: t('Mínimo'), max: t('Máximo'), distinct: t('Distintos') }[v.fn];
    add({ op: 'group', by: byCols, aggs: [{ fn: v.fn, col: v.col, as: v.as || (v.fn === 'count' ? fnName : `${fnName} ${v.col}`) }] });
  };
  const toolbar = h('div', { class: 'qe-tools' },
    h('button', { class: 'btn', onclick: () => add({ op: 'dedupe' }) }, t('Quitar duplicados')),
    h('button', { class: 'btn', onclick: () => add({ op: 'removeEmpty' }) }, t('Quitar filas vacías')),
    h('button', { class: 'btn', onclick: async () => { const v = await form('Conservar primeras filas', [{ name: 'n', label: 'Número de filas', type: 'number', value: 100 }], 'Aceptar'); if (v) add({ op: 'top', n: v.n }); } }, t('Conservar primeras filas')),
    h('button', { class: 'btn', onclick: async () => { const v = await form('Quitar primeras filas', [{ name: 'n', label: 'Número de filas', type: 'number', value: 1 }], 'Aceptar'); if (v) add({ op: 'skip', n: v.n }); } }, t('Quitar primeras filas')),
    h('button', { class: 'btn', onclick: async () => { const v = await form('Columna personalizada', [{ name: 'name', label: 'Nombre de la columna', value: t('Personalizada') }, { name: 'expr', label: 'Fórmula, p. ej. [Precio] * [Cantidad]  o  [Nombre] & " " & [Apellido]', value: '' }, { type: 'note', label: tf('Columnas: {0}', cur?.cols.map((c) => `[${c}]`).join(' ') || '') }], 'Aceptar'); if (!v) return; try { compileExpr(v.expr, cur.cols); add({ op: 'addCol', name: v.name, expr: v.expr }); } catch (e) { toast(e.message, { error: true }); } } }, t('Columna personalizada')),
    h('button', { class: 'btn', onclick: () => groupDialog([cur?.cols[0]].filter(Boolean)) }, t('Agrupar por')),
    h('button', { class: 'btn', onclick: async () => { const others = C.wb.queries.filter((x) => x.id !== q.id).map((x) => [x.name, x.name]); if (!others.length) { toast('Primero crea otra consulta.'); return; } const v = await form('Anexar consultas', [{ name: 'query', label: 'Anexar las filas de', type: 'select', value: others[0][0], options: others }], 'Aceptar'); if (v) add({ op: 'append', query: v.query }); } }, t('Anexar')),
    h('button', { class: 'btn', onclick: async () => {
      const others = C.wb.queries.filter((x) => x.id !== q.id); if (!others.length) { toast('Primero crea otra consulta.'); return; }
      const v1 = await form('Combinar consultas', [{ name: 'query', label: 'Combinar con', type: 'select', value: others[0].name, options: others.map((x) => [x.name, x.name]) }, { name: 'left', label: 'Columna de esta consulta', type: 'select', value: cur.cols[0], options: cur.cols.map((c) => [c, c]) }, { name: 'kind', label: 'Tipo', type: 'select', value: 'left', options: [['left', 'Conservar todas las filas de esta consulta'], ['inner', 'Solo las filas que coinciden']] }], 'Siguiente');
      if (!v1) return; const o = await runQuery(others.find((x) => x.name === v1.query));
      const v2 = await form('Combinar consultas', [{ name: 'right', label: tf('Columna de {0}', v1.query), type: 'select', value: o.cols[0], options: o.cols.map((c) => [c, c]) }], 'Aceptar');
      if (v2) add({ op: 'merge', query: v1.query, left: v1.left, right: v2.right, kind: v1.kind });
    } }, t('Combinar')),
    h('label', { class: 'qe-hdr' }, hdrChk, ` ${t('La primera fila son encabezados')}`));
  hdrChk.addEventListener('change', () => { q.source.headers = hdrChk.checked; redraw(); });
  const body = h('div', { class: 'qe' }, h('div', { class: 'qe-top' }, h('label', {}, t('Nombre de la consulta'), nameIn), info), toolbar, h('div', { class: 'qe-main' }, h('div', { class: 'qe-side' }, h('b', {}, t('Pasos aplicados')), stepsEl), grid));
  await redraw();
  const r = await modal({ title: isNew ? t('Obtener y transformar datos') : tf('Editar consulta {0}', q.name), body, width: 'min(1200px, 94vw)', buttons: [{ label: 'Cancelar', value: null }, { label: 'Cerrar y cargar', primary: true, value: true }] });
  if (!r) return null;
  q.name = nameIn.value.trim() || q.name;
  if (C.wb.queries.some((x) => x.name === q.name && x.id !== q.id)) { toast('Ya existe una consulta con ese nombre.', { error: true }); return null; }
  let tb; try { tb = await runQuery(q, { bytes }); } catch (e) { toast(e.message, { error: true }); return null; }
  const i = C.wb.queries.findIndex((x) => x.id === q.id);
  if (i >= 0) C.wb.queries[i] = q; else C.wb.queries.push(q);
  writeTable(q, tb);
  toast(tf('{0} filas cargadas en la hoja {1}', tb.rows.length, q.sheetName));
  return q;
}
export async function manageQueries() {
  const list = h('div', { class: 'q-list' });
  let toEdit = null;
  const draw = () => list.replaceChildren(...(C.wb.queries.length ? C.wb.queries.map((q) => h('div', { class: 'q-item' },
    h('div', { class: 'no-i18n' }, h('b', {}, q.name), h('div', { class: 'muted small' }, `${q.source.kind === 'range' ? `${q.source.sheet}!${q.source.range}` : (q.source.path || '')} · ${tf('{0} pasos', q.steps.length)} · ${q.lastRows != null ? tf('{0} filas', q.lastRows) : ''}`)),
    h('div', { class: 'row-btns' },
      h('button', { class: 'btn ghost', onclick: async () => { await refreshQuery(q); draw(); } }, t('Actualizar')),
      h('button', { class: 'btn ghost', onclick: () => { toEdit = q; list.closest('.modal')?.querySelector('.btn.primary')?.click(); } }, t('Editar')),
      h('button', { class: 'btn ghost', onclick: () => { C.wb.queries = C.wb.queries.filter((x) => x !== q); C.shell.markDirty(); draw(); } }, t('Eliminar'))))) : [h('p', { class: 'muted' }, t('Todavía no hay consultas. Usa “Obtener datos” para crear una.'))]));
  draw();
  await modal({ title: t('Consultas y conexiones'), body: list, width: '640px', buttons: [{ label: 'Cerrar', value: true, primary: true }] });
  if (toEdit) await editQuery(toEdit);
}
export { esc };
