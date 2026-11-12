// Calcula Datos — parte 2 (como LibreOffice Calc): ordenar con varios criterios y opciones, rellenar
// series y rellenar hacia abajo/derecha/arriba/izquierda, números aleatorios, estadísticas (descriptiva,
// correlación, covarianza, media móvil, suavizado exponencial, muestreo y regresión), consolidar,
// gestionar duplicados, estilos de celda, mayúsculas y minúsculas, y mover, copiar, ocultar y mostrar
// hojas. GPL-3.0
import { h, toast, t, tf } from '../../shared/office.js';
import { tabbed, field, numIn, sel, chk } from '../../shared/dialogs.js';
import { colName, addr, parseAddr, parseRange } from './workbook.js';
import { allLists, listRank } from './analisis.js';

let C = null; // { wb, grid, shell, act, refresh, sel, clamp, switchSheet, renderTabs, refreshCharts }
export function init(ctx) { C = ctx; }
const abs = (r, c) => `$${colName(c)}$${r + 1}`;
const absR = (r1, c1, r2, c2) => `${abs(r1, c1)}:${abs(r2, c2)}`;
const isText = (v) => typeof v === 'string' && v !== '';
const regionOf = (g) => (g.r1 === g.r2 && g.c1 === g.c2 ? C.wb.region(g.r1, g.c1) : C.clamp(g));
function guessHeader(g) {
  const wb = C.wb; let allText = true; let below = false;
  for (let c = g.c1; c <= g.c2; c++) { const v = wb.value(g.r1, c); if (!isText(v)) allText = false; if (g.r2 > g.r1 && !isText(wb.value(g.r1 + 1, c))) below = true; }
  return allText && below;
}
const v = (body, n) => body.querySelector(`[name="${n}"]`);
const valOf = (body, n) => { const el = v(body, n); return el.type === 'checkbox' ? el.checked : el.value; };

// ================================================================ ordenar
export async function sortDialog() {
  const wb = C.wb; const g = regionOf(C.sel());
  if (g.r1 === g.r2 && g.c1 === g.c2) { toast('Selecciona los datos que quieres ordenar.', { error: true }); return null; }
  const hdr = guessHeader(g);
  const colOpts = () => { const o = [['', t('(ninguno)')]]; for (let c = g.c1; c <= g.c2; c++) o.push([String(c), hdr ? `${colName(c)} — ${String(wb.display(g.r1, c).text || '').slice(0, 24)}` : `${t('Columna')} ${colName(c)}`]); return o; };
  const key = (i, def) => [field(tf('Criterio {0}', i), sel(`k${i}`, def, colOpts())), field('Orden', sel(`o${i}`, 'asc', [['asc', 'Ascendente'], ['desc', 'Descendente']]))];
  const body = await tabbed('Ordenar', [
    { label: 'Criterios de ordenación', body: h('div', { class: 'fx-grid' }, ...key(1, String(C.grid.sel.c >= g.c1 && C.grid.sel.c <= g.c2 ? C.grid.sel.c : g.c1)), ...key(2, ''), ...key(3, '')) },
    { label: 'Opciones', body: h('div', { class: 'fx-grid' },
      chk('header', 'El intervalo contiene etiquetas de columna', hdr), chk('cs', 'Distinguir mayúsculas y minúsculas', false),
      chk('natural', 'Orden natural (2 antes de 10, «a2» antes de «a10»)', true), chk('formats', 'Incluir los formatos', true),
      field('Orden personalizado (criterio 1)', sel('list', '', [['', t('(ninguno)')], ...allLists().map((l, i) => [String(i), l.slice(0, 4).join(', ') + (l.length > 4 ? '…' : '')])])),
      field('Dirección', sel('dir', 'rows', [['rows', 'De arriba abajo (ordenar filas)'], ['cols', 'De izquierda a derecha (ordenar columnas)']])),
      field('Copiar el resultado en', h('input', { name: 'copyTo', placeholder: t('Por ejemplo: H1 (vacío: en su lugar)') }))) },
  ], { width: '560px' });
  if (!body) return null;
  const keys = [1, 2, 3].map((i) => ({ c: valOf(body, `k${i}`), asc: valOf(body, `o${i}`) === 'asc' })).filter((k) => k.c !== '').map((k) => ({ ...k, c: Number(k.c) }));
  if (!keys.length) return null;
  let dest = null; if (valOf(body, 'copyTo').trim()) { dest = parseAddr(valOf(body, 'copyTo').trim().split(':')[0]); if (!dest) { toast('Dirección no válida', { error: true }); return null; } }
  const opts = { header: valOf(body, 'header'), cs: valOf(body, 'cs'), natural: valOf(body, 'natural'), formats: valOf(body, 'formats'), dir: valOf(body, 'dir'), dest, list: valOf(body, 'list') !== '' ? allLists()[Number(valOf(body, 'list'))] : null };
  if (opts.dir === 'cols') keys.forEach((k) => { k.c = g.r1 + (k.c - g.c1); }); // en columnas, el criterio es una fila
  sortMulti(g, keys, opts);
  return { g, keys, opts };
}
export function sortMulti(g, keys, { header = false, cs = false, natural = true, formats = true, dir = 'rows', dest = null, list = null } = {}) {
  const wb = C.wb; const s = wb.sheet;
  const byCols = dir === 'cols';
  const lines = []; // cada fila (o columna) con sus celdas
  const from = byCols ? (header ? g.c1 + 1 : g.c1) : (header ? g.r1 + 1 : g.r1);
  const to = byCols ? g.c2 : g.r2;
  for (let i = from; i <= to; i++) {
    const cells = [];
    if (byCols) for (let r = g.r1; r <= g.r2; r++) cells.push({ raw: wb.raw(r, i), st: s.style(r, i) });
    else for (let c = g.c1; c <= g.c2; c++) cells.push({ raw: wb.raw(i, c), st: s.style(i, c) });
    lines.push({ cells, vals: keys.map((k) => (byCols ? wb.value(k.c, i) : wb.value(i, k.c))) });
  }
  const coll = new Intl.Collator(document.documentElement.lang === 'en' ? 'en' : 'es', { numeric: natural, sensitivity: cs ? 'case' : 'base', caseFirst: cs ? 'upper' : false });
  const rank = (x) => (x == null || x === '' ? 3 : typeof x === 'number' ? 0 : typeof x === 'string' ? 1 : 2);
  lines.sort((a, b) => {
    for (let k = 0; k < keys.length; k++) {
      const x = a.vals[k]; const y = b.vals[k];
      if (k === 0 && list) { const d0 = listRank(list, x) - listRank(list, y); if (d0 && Number.isFinite(d0)) return keys[k].asc ? d0 : -d0; if (listRank(list, x) !== listRank(list, y)) return listRank(list, x) === Infinity ? 1 : -1; }
      const rx = rank(x); const ry = rank(y);
      if (rx !== ry) return rx - ry; // vacíos al final
      let d = 0; if (rx === 0) d = x - y; else if (rx === 1) d = coll.compare(x, y); else if (rx === 2) d = String(x).localeCompare(String(y));
      if (d) return keys[k].asc ? d : -d;
    }
    return 0;
  });
  const R0 = dest ? dest.r : g.r1; const C0 = dest ? dest.c : g.c1;
  C.act(t('Ordenar'), () => {
    const put = (r, c, cell) => { wb.setRaw(r, c, cell.raw ?? null); if (formats) { if (cell.st) s.styles.set(`${r},${c}`, cell.st); else s.styles.delete(`${r},${c}`); } };
    if (dest && header) { if (byCols) for (let r = g.r1; r <= g.r2; r++) put(R0 + r - g.r1, C0, { raw: wb.raw(r, g.c1), st: s.style(r, g.c1) }); else for (let c = g.c1; c <= g.c2; c++) put(R0, C0 + c - g.c1, { raw: wb.raw(g.r1, c), st: s.style(g.r1, c) }); }
    lines.forEach((ln, i) => ln.cells.forEach((cell, j) => {
      const off = (header ? 1 : 0) + i;
      if (byCols) put(R0 + j, C0 + off, cell); else put(R0 + off, C0 + j, cell);
    }));
  });
  C.refreshCharts?.();
}

// ================================================================ rellenar
// Rellenar hacia abajo, la derecha, arriba o la izquierda: copia la primera celda (las fórmulas se ajustan)
export function fillDir(dir) {
  const wb = C.wb; const s = wb.sheet; const g = C.clamp(C.sel());
  const vert = dir === 'down' || dir === 'up';
  if ((vert && g.r1 === g.r2) || (!vert && g.c1 === g.c2)) { toast('Selecciona varias celdas para rellenar.', { error: true }); return; }
  C.act(t('Rellenar'), () => {
    if (vert) for (let c = g.c1; c <= g.c2; c++) {
      const sr = dir === 'down' ? g.r1 : g.r2;
      wb.hf.copy({ start: { sheet: s.id, row: sr, col: c }, end: { sheet: s.id, row: sr, col: c } });
      for (let r = g.r1; r <= g.r2; r++) if (r !== sr) { wb.hf.paste({ sheet: s.id, row: r, col: c }); const st = s.style(sr, c); if (st) s.styles.set(`${r},${c}`, st); else s.styles.delete(`${r},${c}`); }
    } else for (let r = g.r1; r <= g.r2; r++) {
      const sc = dir === 'right' ? g.c1 : g.c2;
      wb.hf.copy({ start: { sheet: s.id, row: r, col: sc }, end: { sheet: s.id, row: r, col: sc } });
      for (let c = g.c1; c <= g.c2; c++) if (c !== sc) { wb.hf.paste({ sheet: s.id, row: r, col: c }); const st = s.style(r, sc); if (st) s.styles.set(`${r},${c}`, st); else s.styles.delete(`${r},${c}`); }
    }
  });
}
// días de Excel/LibreOffice (1899-12-30) ↔ fecha
const EPOCH = Date.UTC(1899, 11, 30);
const toDate = (n) => new Date(EPOCH + Math.round(n * 86400000));
const toSerial = (d) => (d.getTime() - EPOCH) / 86400000;
export async function fillSeriesDialog() {
  const wb = C.wb; const g = C.clamp(C.sel());
  const first = wb.value(g.r1, g.c1);
  const fmt = wb.sheet.style(g.r1, g.c1)?.fmt || '';
  const isDate = typeof first === 'number' && /[dy]/i.test(fmt.replace(/"[^"]*"/g, '')) && !/[0#]/.test(fmt);
  const body = await tabbed('Rellenar serie', [{ label: 'Serie', body: h('div', { class: 'fx-grid' },
    field('Dirección', sel('dir', g.r2 > g.r1 || g.c1 === g.c2 ? 'down' : 'right', [['down', 'Abajo'], ['right', 'Derecha'], ['up', 'Arriba'], ['left', 'Izquierda']])),
    field('Tipo de serie', sel('type', isDate ? 'date' : 'linear', [['linear', 'Lineal'], ['growth', 'Crecimiento'], ['date', 'Fecha'], ['auto', 'Autorrelleno']])),
    field('Unidad de tiempo', sel('unit', 'day', [['day', 'Día'], ['weekday', 'Día laborable'], ['month', 'Mes'], ['year', 'Año']])), h('span'),
    field('Valor inicial', h('input', { name: 'start', value: typeof first === 'number' ? (isDate ? toDate(first).toISOString().slice(0, 10) : String(first)) : '' })),
    field('Valor final', h('input', { name: 'end', placeholder: t('(opcional)') })),
    field('Incremento', h('input', { name: 'step', value: '1' }))) }], { width: '520px' });
  if (!body) return null;
  return fillSeries(g, { dir: valOf(body, 'dir'), type: valOf(body, 'type'), unit: valOf(body, 'unit'), start: valOf(body, 'start'), end: valOf(body, 'end'), step: valOf(body, 'step') });
}
export function fillSeries(g, { dir = 'down', type = 'linear', unit = 'day', start, end, step = '1' }) {
  const wb = C.wb; const s = wb.sheet;
  const num = (x) => Number(String(x).replace(',', '.'));
  const date = (x) => { if (/^\d{4}-\d{2}-\d{2}$/.test(String(x))) return toSerial(new Date(`${x}T00:00:00Z`)); const n = num(x); return Number.isNaN(n) ? null : n; };
  const vert = dir === 'down' || dir === 'up';
  let n = vert ? g.r2 - g.r1 + 1 : g.c2 - g.c1 + 1;
  const inc = num(step) || (type === 'growth' ? 2 : 1);
  let x0 = type === 'date' ? date(start) : num(start);
  if (type === 'auto') return autoFill(g, dir);
  if (x0 == null || Number.isNaN(x0)) { toast('El valor inicial no es válido.', { error: true }); return null; }
  const endV = String(end ?? '').trim() ? (type === 'date' ? date(end) : num(end)) : null;
  let step_i = 0;
  const nextOf = (x) => {
    if (type === 'growth') return x * inc;
    if (type !== 'date') return x + inc;
    const d = toDate(x);
    if (unit === 'month' || unit === 'year') {
      // como LibreOffice: se conserva el día de la fecha inicial y se ajusta al último día del mes
      step_i += 1; const d0 = toDate(x0); const m = d0.getUTCMonth() + step_i * inc * (unit === 'year' ? 12 : 1);
      const y = d0.getUTCFullYear() + Math.floor(m / 12); const mm = ((m % 12) + 12) % 12; const dim = new Date(Date.UTC(y, mm + 1, 0)).getUTCDate();
      return toSerial(new Date(Date.UTC(y, mm, Math.min(d0.getUTCDate(), dim))));
    }
    if (unit === 'weekday') { let k = 0; const stp = Math.sign(inc) || 1; while (k < Math.abs(inc)) { d.setUTCDate(d.getUTCDate() + stp); if (d.getUTCDay() % 6) k++; } return toSerial(d); }
    return x + inc;
  };
  const vals = []; let x = x0;
  const single = vert ? g.r1 === g.r2 : g.c1 === g.c2;
  const limit = endV != null && single ? 100000 : n;
  for (let i = 0; i < limit; i++) {
    if (endV != null && ((inc > 0 && x > endV + 1e-9) || (inc < 0 && x < endV - 1e-9))) break;
    vals.push(Math.round(x * 1e10) / 1e10); x = nextOf(x);
  }
  n = vals.length;
  const st0 = s.style(g.r1, g.c1);
  C.act(t('Rellenar serie'), () => {
    vals.forEach((val, i) => {
      const r = vert ? (dir === 'down' ? g.r1 + i : g.r2 - i) : g.r1;
      const c = vert ? g.c1 : (dir === 'right' ? g.c1 + i : g.c2 - i);
      if (r < 0 || c < 0) return;
      for (let k = 0; k <= (vert ? g.c2 - g.c1 : g.r2 - g.r1); k++) {
        const rr = vert ? r : g.r1 + k; const cc = vert ? g.c1 + k : c;
        wb.setRaw(rr, cc, val);
        if (type === 'date') wb.setStyle(rr, cc, { fmt: st0?.fmt || 'dd/mm/yyyy' });
      }
    });
  });
  return vals;
}
// autorrelleno con lo que hay en la primera celda (números, días, meses y texto con número)
function autoFill(g, dir) {
  const wb = C.wb; const vert = dir === 'down' || dir === 'up';
  const lists = allLists();
  const n = vert ? g.r2 - g.r1 + 1 : g.c2 - g.c1 + 1;
  const sr = dir === 'up' ? g.r2 : g.r1; const sc = dir === 'left' ? g.c2 : g.c1;
  const v0 = wb.value(sr, sc); const out = [];
  for (let i = 0; i < n; i++) {
    if (typeof v0 === 'number') out.push(v0 + i);
    else if (typeof v0 === 'string') {
      const L = lists.find((l) => l.some((w) => w.toLowerCase() === v0.toLowerCase()));
      if (L) { const k = L.findIndex((w) => w.toLowerCase() === v0.toLowerCase()); const w = L[(k + i) % L.length]; out.push(v0[0] === v0[0].toUpperCase() ? w[0].toUpperCase() + w.slice(1) : w); continue; }
      const m = /^(.*?)(\d+)(\D*)$/.exec(v0); out.push(m ? `${m[1]}${String(Number(m[2]) + i).padStart(m[2].length, '0')}${m[3]}` : v0);
    } else out.push(v0 ?? null);
  }
  C.act(t('Rellenar serie'), () => out.forEach((val, i) => {
    const r = vert ? (dir === 'down' ? g.r1 + i : g.r2 - i) : sr; const c = vert ? sc : (dir === 'right' ? g.c1 + i : g.c2 - i);
    wb.setRaw(r, c, val);
  }));
  return out;
}

// ================================================================ números aleatorios
export async function randomDialog() {
  const g = C.clamp(C.sel());
  const body = await tabbed('Rellenar con números aleatorios', [{ label: 'Opciones', body: h('div', { class: 'fx-grid' },
    field('Intervalo de celdas', h('input', { name: 'range', value: `${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)}` })),
    field('Distribución', sel('dist', 'uniform', [['uniform', 'Uniforme'], ['int', 'Uniforme entera'], ['normal', 'Normal'], ['bernoulli', 'Bernoulli'], ['binomial', 'Binomial'], ['poisson', 'Poisson']])),
    field('Mínimo / media / probabilidad', h('input', { name: 'a', value: '0' })), field('Máximo / desviación / ensayos', h('input', { name: 'b', value: '100' })),
    field('Decimales', numIn('dec', 2, { step: 1, min: 0, max: 10 })), field('Semilla (opcional)', h('input', { name: 'seed', placeholder: t('Vacío: al azar') }))) }], { width: '540px' });
  if (!body) return null;
  const rg = parseRange(valOf(body, 'range')); if (!rg) { toast('Intervalo no válido', { error: true }); return null; }
  return fillRandom(rg, { dist: valOf(body, 'dist'), a: Number(valOf(body, 'a')), b: Number(valOf(body, 'b')), dec: Number(valOf(body, 'dec')), seed: valOf(body, 'seed') });
}
function rng(seed) {
  if (seed === '' || seed == null) return Math.random;
  let s = 0; for (const ch of String(seed)) s = (Math.imul(s, 31) + ch.charCodeAt(0)) | 0;
  return () => { s = (s + 0x6D2B79F5) | 0; let x = Math.imul(s ^ (s >>> 15), 1 | s); x = (x + Math.imul(x ^ (x >>> 7), 61 | x)) ^ x; return ((x ^ (x >>> 14)) >>> 0) / 4294967296; };
}
export function fillRandom(g, { dist = 'uniform', a = 0, b = 1, dec = 2, seed = '' }) {
  const R = rng(seed); const round = (x) => Math.round(x * 10 ** dec) / 10 ** dec;
  const gen = () => {
    switch (dist) {
      case 'int': return Math.floor(a + R() * (Math.floor(b) - Math.ceil(a) + 1));
      case 'normal': { const u = 1 - R(); const w = R(); return round(a + b * Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * w)); }
      case 'bernoulli': return R() < a ? 1 : 0;
      case 'binomial': { let k = 0; for (let i = 0; i < Math.round(b); i++) if (R() < a) k++; return k; }
      case 'poisson': { const L = Math.exp(-a); let k = 0; let p = 1; do { k++; p *= R(); } while (p > L); return k - 1; }
      default: return round(a + R() * (b - a));
    }
  };
  const vals = [];
  C.act(t('Números aleatorios'), () => { for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { const x = gen(); vals.push(x); C.wb.setRaw(r, c, x); } });
  return vals;
}

// ================================================================ estadísticas (resultados con fórmulas, como LibreOffice)
const KINDS = [['descriptive', 'Estadística descriptiva'], ['correlation', 'Correlación'], ['covariance', 'Covarianza'], ['moving', 'Media móvil'], ['smoothing', 'Suavizado exponencial'], ['sampling', 'Muestreo'], ['regression', 'Regresión lineal']];
export async function statsDialog(kind = 'descriptive') {
  const g = regionOf(C.sel());
  const body = await tabbed('Estadísticas', [{ label: 'Datos', body: h('div', { class: 'fx-grid' },
    field('Análisis', sel('kind', kind, KINDS), 'wide'),
    field('Intervalo de entrada', h('input', { name: 'input', value: `${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)}` })),
    field('Resultados en', h('input', { name: 'output', value: addr(g.r1, g.c2 + 2) })),
    field('Agrupado por', sel('by', 'cols', [['cols', 'Columnas'], ['rows', 'Filas']])), chk('labels', 'La primera fila (o columna) tiene etiquetas', guessHeader(g)),
    field('Intervalo (media móvil)', numIn('n', 3, { step: 1, min: 2, max: 1000 })), field('Factor de suavizado', numIn('alpha', 0.3, { step: 0.05, min: 0.01, max: 1 })),
    field('Muestreo', sel('method', 'random', [['random', 'Aleatorio'], ['periodic', 'Periódico']])), field('Tamaño de la muestra / período', numIn('size', 5, { step: 1, min: 1, max: 100000 })),
    chk('replace', 'Con reemplazo (muestreo aleatorio)', false)) }], { width: '600px' });
  if (!body) return null;
  const inp = parseRange(valOf(body, 'input')); const outA = parseAddr(valOf(body, 'output'));
  if (!inp || !outA) { toast('Intervalo no válido', { error: true }); return null; }
  return runStats(valOf(body, 'kind'), inp, outA, { by: valOf(body, 'by'), labels: valOf(body, 'labels'), n: Number(valOf(body, 'n')), alpha: Number(valOf(body, 'alpha')), method: valOf(body, 'method'), size: Number(valOf(body, 'size')), replace: valOf(body, 'replace') });
}
export function runStats(kind, g, o, { by = 'cols', labels = false, n = 3, alpha = 0.3, method = 'random', size = 5, replace = false } = {}) {
  const wb = C.wb; const cells = []; // [r, c, raw, style]
  const put = (r, c, raw, st) => cells.push([r, c, raw, st]);
  const B = { b: true };
  // series de datos
  const series = [];
  if (by === 'cols') for (let c = g.c1; c <= g.c2; c++) series.push({ name: labels ? String(wb.value(g.r1, c) ?? colName(c)) : `${t('Columna')} ${colName(c)}`, ref: absR(labels ? g.r1 + 1 : g.r1, c, g.r2, c), cells: Array.from({ length: g.r2 - g.r1 + 1 - (labels ? 1 : 0) }, (_, i) => [g.r1 + (labels ? 1 : 0) + i, c]) });
  else for (let r = g.r1; r <= g.r2; r++) series.push({ name: labels ? String(wb.value(r, g.c1) ?? r + 1) : `${t('Fila')} ${r + 1}`, ref: absR(r, labels ? g.c1 + 1 : g.c1, r, g.c2), cells: Array.from({ length: g.c2 - g.c1 + 1 - (labels ? 1 : 0) }, (_, i) => [r, g.c1 + (labels ? 1 : 0) + i]) });
  const ref1 = (rc) => abs(rc[0], rc[1]);
  if (kind === 'descriptive') {
    const rows = [['Media', 'AVERAGE(X)'], ['Error estándar', 'SQRT(VAR(X)/COUNT(X))'], ['Moda', 'MODE(X)'], ['Mediana', 'MEDIAN(X)'], ['Primer cuartil', 'QUARTILE(X,1)'], ['Tercer cuartil', 'QUARTILE(X,3)'],
      ['Varianza', 'VAR(X)'], ['Desviación estándar', 'STDEV(X)'], ['Curtosis', 'KURT(X)'], ['Asimetría', 'SKEW(X)'], ['Rango', 'MAX(X)-MIN(X)'], ['Mínimo', 'MIN(X)'], ['Máximo', 'MAX(X)'], ['Suma', 'SUM(X)'], ['Cuenta', 'COUNT(X)']];
    put(o.r, o.c, t('Estadística descriptiva'), B);
    series.forEach((s, j) => put(o.r, o.c + 1 + j, s.name, B));
    rows.forEach(([lab, f], i) => { put(o.r + 1 + i, o.c, t(lab)); series.forEach((s, j) => put(o.r + 1 + i, o.c + 1 + j, `=${f.replace(/X/g, s.ref)}`, { fmt: '0.0000' })); });
  } else if (kind === 'correlation' || kind === 'covariance') {
    const fn = kind === 'correlation' ? 'CORREL' : 'COVARIANCE.P';
    put(o.r, o.c, t(kind === 'correlation' ? 'Correlaciones' : 'Covarianzas'), B);
    series.forEach((s, j) => { put(o.r, o.c + 1 + j, s.name, B); put(o.r + 1 + j, o.c, s.name, B); });
    series.forEach((a, i) => series.forEach((b2, j) => { if (j <= i) put(o.r + 1 + i, o.c + 1 + j, `=${fn}(${a.ref},${b2.ref})`, { fmt: '0.0000' }); }));
  } else if (kind === 'moving' || kind === 'smoothing') {
    series.forEach((s, j) => {
      const isRow = by === 'rows';
      const at = (i) => (isRow ? [o.r + j, o.c + i] : [o.r + i, o.c + j]);
      put(...at(0), kind === 'moving' ? tf('Media móvil ({0})', s.name) : tf('Suavizado ({0})', s.name), B);
      s.cells.forEach((rc, i) => {
        if (kind === 'moving') { if (i >= n - 1) put(...at(i + 1), `=AVERAGE(${absR(...s.cells[i - n + 1], ...rc)})`, { fmt: '0.0000' }); else put(...at(i + 1), '=NA()'); }
        else if (i === 0) put(...at(1), `=${ref1(rc)}`, { fmt: '0.0000' });
        else { const prev = at(i); put(...at(i + 1), `=${alpha}*${ref1(rc)}+(1-${alpha})*${colName(prev[1])}${prev[0] + 1}`, { fmt: '0.0000' }); }
      });
    });
  } else if (kind === 'sampling') {
    const R = Math.random;
    series.forEach((s, j) => {
      put(o.r, o.c + j, tf('Muestra ({0})', s.name), B);
      let pick = [];
      if (method === 'periodic') pick = s.cells.filter((_, i) => (i + 1) % size === 0);
      else if (replace) pick = Array.from({ length: size }, () => s.cells[Math.floor(R() * s.cells.length)]);
      else { const pool = [...s.cells]; for (let i = 0; i < Math.min(size, pool.length); i++) { const k = i + Math.floor(R() * (pool.length - i)); [pool[i], pool[k]] = [pool[k], pool[i]]; } pick = pool.slice(0, Math.min(size, pool.length)); }
      pick.forEach((rc, i) => put(o.r + 1 + i, o.c + j, wb.value(rc[0], rc[1]) ?? null));
    });
  } else if (kind === 'regression') {
    if (series.length < 2) { toast('La regresión necesita dos series: la variable independiente (X) y la dependiente (Y).', { error: true }); return null; }
    const X = series[0].ref; const Y = series[1].ref;
    put(o.r, o.c, t('Regresión lineal'), B); put(o.r, o.c + 1, `${series[1].name} ~ ${series[0].name}`, B);
    [['Pendiente', `=SLOPE(${Y},${X})`], ['Intersección', `=INTERCEPT(${Y},${X})`], ['R²', `=RSQ(${Y},${X})`], ['Correlación (r)', `=CORREL(${Y},${X})`], ['Error típico de la estimación', `=STEYX(${Y},${X})`], ['Observaciones', `=COUNT(${X})`]]
      .forEach(([lab, f], i) => { put(o.r + 1 + i, o.c, t(lab)); put(o.r + 1 + i, o.c + 1, f, { fmt: i === 5 ? undefined : '0.0000' }); });
  } else return null;
  C.act(t('Estadísticas'), () => { for (const [r, c, raw, st] of cells) { wb.setRaw(r, c, raw ?? null); if (st) wb.setStyle(r, c, st); } });
  return cells.length;
}

// ================================================================ consolidar
const CFN = [['SUM', 'Suma'], ['AVERAGE', 'Promedio'], ['COUNT', 'Cuenta'], ['MAX', 'Máximo'], ['MIN', 'Mínimo'], ['PRODUCT', 'Producto']];
export async function consolidateDialog() {
  const g = C.clamp(C.sel()); const cur = C.wb.sheet.name;
  const body = await tabbed('Consolidar', [{ label: 'Consolidar', body: h('div', { class: 'fx-col' },
    field('Función', sel('fn', 'SUM', CFN)),
    field('Intervalos de origen (uno por línea; por ejemplo: Hoja2!A1:C10)', h('textarea', { name: 'sources', rows: 4, style: { width: '100%', fontFamily: 'monospace' } }, `${cur}!${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)}`)),
    field('Copiar los resultados en', h('input', { name: 'dest', value: addr(g.r1, g.c2 + 2) })),
    chk('rowLabels', 'Usar la primera columna como etiquetas', true), chk('colLabels', 'Usar la primera fila como etiquetas', true)) }], { width: '560px' });
  if (!body) return null;
  const dest = parseAddr(valOf(body, 'dest')); if (!dest) { toast('Dirección no válida', { error: true }); return null; }
  return consolidate(valOf(body, 'sources').split('\n').map((x) => x.trim()).filter(Boolean), dest, { fn: valOf(body, 'fn'), rowLabels: valOf(body, 'rowLabels'), colLabels: valOf(body, 'colLabels') });
}
export function consolidate(sources, dest, { fn = 'SUM', rowLabels = true, colLabels = true } = {}) {
  const wb = C.wb;
  const acc = new Map(); const rowKeys = []; const colKeys = [];
  for (const src of sources) {
    const m = /^(?:'?([^'!]+)'?!)?(.+)$/.exec(src); const sh = m[1] ? wb.sheets.find((s) => s.name.toLowerCase() === m[1].toLowerCase()) : wb.sheet; const g = parseRange(m[2]);
    if (!sh || !g) { toast(tf('Intervalo de origen no válido: {0}', src), { error: true }); return null; }
    for (let r = g.r1 + (colLabels ? 1 : 0); r <= g.r2; r++) for (let c = g.c1 + (rowLabels ? 1 : 0); c <= g.c2; c++) {
      const rk = rowLabels ? String(wb.value(r, g.c1, sh) ?? '') : String(r - g.r1); const ck = colLabels ? String(wb.value(g.r1, c, sh) ?? '') : String(c - g.c1);
      if (!rowKeys.includes(rk)) rowKeys.push(rk); if (!colKeys.includes(ck)) colKeys.push(ck);
      const x = wb.value(r, c, sh); const k = `${rk}\u0001${ck}`; const a = acc.get(k) || []; if (typeof x === 'number') a.push(x); acc.set(k, a);
    }
  }
  const agg = (a) => { if (fn === 'COUNT') return a.length; if (!a.length) return null; if (fn === 'SUM') return a.reduce((p, q) => p + q, 0); if (fn === 'AVERAGE') return a.reduce((p, q) => p + q, 0) / a.length; if (fn === 'MAX') return Math.max(...a); if (fn === 'MIN') return Math.min(...a); return a.reduce((p, q) => p * q, 1); };
  const r0 = dest.r + (colLabels ? 1 : 0); const c0 = dest.c + (rowLabels ? 1 : 0);
  C.act(t('Consolidar'), () => {
    if (colLabels) colKeys.forEach((k, j) => { wb.setRaw(dest.r, c0 + j, k); wb.setStyle(dest.r, c0 + j, { b: true }); });
    if (rowLabels) rowKeys.forEach((k, i) => { wb.setRaw(r0 + i, dest.c, k); wb.setStyle(r0 + i, dest.c, { b: true }); });
    rowKeys.forEach((rk, i) => colKeys.forEach((ck, j) => wb.setRaw(r0 + i, c0 + j, agg(acc.get(`${rk}\u0001${ck}`) || []))));
  });
  return { rows: rowKeys.length, cols: colKeys.length };
}

// ================================================================ duplicados
export async function duplicatesDialog() {
  const wb = C.wb; const g = regionOf(C.sel()); const hdr = guessHeader(g);
  const list = h('div', { class: 'fx-col', style: { maxHeight: '180px', overflow: 'auto' } });
  for (let c = g.c1; c <= g.c2; c++) list.append(chk(`col${c}`, `${colName(c)}${hdr ? ` — ${String(wb.display(g.r1, c).text || '').slice(0, 24)}` : ''}`, true));
  const body = await tabbed('Gestionar duplicados', [{ label: 'Duplicados', body: h('div', { class: 'fx-col' },
    h('p', { class: 'fx-note' }, tf('Intervalo: {0}', `${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)}`)), h('strong', {}, t('Comparar las columnas')), list,
    chk('header', 'El intervalo contiene etiquetas de columna', hdr), chk('cs', 'Distinguir mayúsculas y minúsculas', false),
    field('Acción', sel('action', 'remove', [['remove', 'Quitar los duplicados'], ['mark', 'Resaltar los duplicados']]))) }], { width: '480px' });
  if (!body) return null;
  const cols = []; for (let c = g.c1; c <= g.c2; c++) if (valOf(body, `col${c}`)) cols.push(c);
  return handleDuplicates(g, { cols, header: valOf(body, 'header'), cs: valOf(body, 'cs'), action: valOf(body, 'action') });
}
export function handleDuplicates(g, { cols, header = true, cs = false, action = 'remove' }) {
  const wb = C.wb; const s = wb.sheet;
  const seen = new Set(); const keep = []; const dups = [];
  for (let r = g.r1 + (header ? 1 : 0); r <= g.r2; r++) {
    const k = cols.map((c) => { const x = String(wb.display(r, c).text ?? ''); return cs ? x : x.toLowerCase(); }).join('\u0001');
    if (seen.has(k)) dups.push(r); else { seen.add(k); keep.push(r); }
  }
  if (!dups.length) { toast('No se encontraron duplicados.'); return 0; }
  if (action === 'mark') { C.act(t('Resaltar duplicados'), () => dups.forEach((r) => { for (let c = g.c1; c <= g.c2; c++) wb.setStyle(r, c, { bg: '#FFC7CE', color: '#9C0006' }); }), { hf: false }); toast(tf('{0} filas duplicadas resaltadas', dups.length)); return dups.length; }
  const rows = keep.map((r) => { const o = []; for (let c = g.c1; c <= g.c2; c++) o.push({ raw: wb.raw(r, c), st: s.style(r, c) }); return o; });
  const start = g.r1 + (header ? 1 : 0);
  C.act(t('Quitar duplicados'), () => {
    for (let i = 0; i <= g.r2 - start; i++) for (let j = 0; j <= g.c2 - g.c1; j++) {
      const cell = rows[i]?.[j]; const r = start + i; const c = g.c1 + j;
      wb.setRaw(r, c, cell ? cell.raw ?? null : null);
      if (cell?.st) s.styles.set(`${r},${c}`, cell.st); else s.styles.delete(`${r},${c}`);
    }
  });
  toast(tf('Se quitaron {0} filas duplicadas; quedan {1}.', dups.length, keep.length));
  return dups.length;
}

// ================================================================ estilos de celda (los de LibreOffice)
export const CELL_STYLES = [
  ['Predeterminado', null], ['Título', { b: true, size: 18 }], ['Título 1', { b: true, size: 16, color: '#1F3864', border: { b: '#1F3864 medium' } }], ['Título 2', { b: true, size: 13, color: '#1F3864', border: { b: '#8EA9DB' } }],
  ['Encabezado', { b: true, align: 'center', bg: '#D9E1F2', border: { b: '#1F3864' } }], ['Texto', { valign: 'top', wrap: true }], ['Nota', { bg: '#FFFFCC', border: { t: '#B2B2B2', b: '#B2B2B2', l: '#B2B2B2', r: '#B2B2B2' } }],
  ['Nota al pie', { i: true, size: 9, color: '#595959' }], ['Resultado', { b: true, border: { t: '#1F3864', b: '#1F3864 double' } }],
  ['Bueno', { bg: '#C6EFCE', color: '#006100' }], ['Neutral', { bg: '#FFEB9C', color: '#9C5700' }], ['Malo', { bg: '#FFC7CE', color: '#9C0006' }],
  ['Advertencia', { color: '#C00000' }], ['Error', { b: true, bg: '#C00000', color: '#FFFFFF' }], ['Acento', { b: true }], ['Acento 1', { b: true, bg: '#000000', color: '#FFFFFF' }], ['Acento 2', { b: true, bg: '#808080', color: '#FFFFFF' }], ['Acento 3', { b: true, bg: '#DDDDDD' }],
  ['Hipervínculo', { u: true, color: '#0563C1' }],
];
const STYLE_KEYS = ['b', 'i', 'u', 'size', 'color', 'bg', 'border', 'align', 'valign', 'wrap'];
export function applyCellStyle(name) {
  const st = CELL_STYLES.find(([n]) => n === name); if (!st) return;
  const patch = Object.fromEntries(STYLE_KEYS.map((k) => [k, undefined])); Object.assign(patch, st[1] || {});
  C.act(t('Estilo de celda'), () => C.wb.styleRange(C.clamp(C.sel()), patch), { hf: false });
}
export function cellStylesMenu() {
  return CELL_STYLES.map(([name, st]) => ({
    label: name, run: () => applyCellStyle(name),
    style: `${st?.bg ? `background:${st.bg};` : ''}${st?.color ? `color:${st.color};` : ''}${st?.b ? 'font-weight:700;' : ''}${st?.i ? 'font-style:italic;' : ''}${st?.u ? 'text-decoration:underline;' : ''}`,
  }));
}

// ================================================================ mayúsculas y minúsculas
export function textCase(mode) {
  const wb = C.wb; const g = C.clamp(C.sel());
  const conv = (s) => {
    if (mode === 'upper') return s.toLocaleUpperCase();
    if (mode === 'lower') return s.toLocaleLowerCase();
    if (mode === 'title') return s.toLocaleLowerCase().replace(/(^|[\s\-(¿¡"'«])(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase());
    if (mode === 'sentence') return s.toLocaleLowerCase().replace(/(^\s*|[.!?]\s+)(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase());
    if (mode === 'toggle') return [...s].map((ch) => (ch === ch.toLocaleUpperCase() ? ch.toLocaleLowerCase() : ch.toLocaleUpperCase())).join('');
    return s;
  };
  C.act(t('Cambiar mayúsculas y minúsculas'), () => {
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { const raw = wb.raw(r, c); if (typeof raw === 'string' && raw && !raw.startsWith('=')) wb.setRaw(r, c, conv(raw)); }
  });
}

// ================================================================ hojas: mover, copiar, ocultar y mostrar
export async function moveCopyDialog(i = C.wb.active) {
  const wb = C.wb;
  const body = await tabbed('Mover o copiar hoja', [{ label: 'Hoja', body: h('div', { class: 'fx-col' },
    field('Acción', sel('action', 'move', [['move', 'Mover'], ['copy', 'Copiar']])),
    field('Insertar antes de', sel('pos', String(i), [...wb.sheets.map((s, k) => [String(k), s.name]), [String(wb.sheets.length), t('- mover al final -')]])),
    field('Nombre nuevo', h('input', { name: 'name', value: wb.sheets[i].name }))) }], { width: '420px' });
  if (!body) return null;
  return moveCopySheet(i, Number(valOf(body, 'pos')), valOf(body, 'action') === 'copy', valOf(body, 'name').trim());
}
export function moveCopySheet(i, pos, copy = false, name = '') {
  const wb = C.wb;
  if (copy) {
    C.dup(i); i = wb.sheets.length - 1; // la copia queda al final
    const ns = wb.sheets[i];
    if (name && !wb.sheets.some((x) => x !== ns && x.name === name)) { try { wb.hf.renameSheet(ns.id, name); ns.name = name; } catch { /* nombre no válido */ } }
  } else if (name && name !== wb.sheets[i].name && !wb.sheets.some((x) => x.name === name)) { wb.hf.renameSheet(wb.sheets[i].id, name); wb.sheets[i].name = name; }
  const s = wb.sheets[i]; if (!s) return null;
  const target = Math.max(0, Math.min(wb.sheets.length, pos));
  wb.sheets.splice(i, 1); wb.sheets.splice(target > i ? target - 1 : target, 0, s);
  wb.active = wb.sheets.indexOf(s);
  C.renderTabs(); C.switchSheet(wb.active); C.shell.markDirty();
  return wb.sheets.map((x) => x.name);
}
export function hideSheet(i = C.wb.active) {
  const wb = C.wb;
  if (wb.sheets.filter((s) => !s.x3?.hidden).length <= 1) { toast('Debe quedar al menos una hoja visible.', { error: true }); return; }
  const s = wb.sheets[i]; s.x3 = { ...(s.x3 || {}), hidden: true };
  const next = wb.sheets.findIndex((x) => !x.x3?.hidden);
  C.renderTabs(); C.switchSheet(next); C.shell.markDirty();
}
export async function showSheetDialog() {
  const wb = C.wb; const hid = wb.sheets.map((s, i) => [s, i]).filter(([s]) => s.x3?.hidden);
  if (!hid.length) { toast('No hay hojas ocultas.'); return; }
  const body = await tabbed('Mostrar hoja', [{ label: 'Hojas ocultas', body: h('div', { class: 'fx-col' }, ...hid.map(([s, i]) => chk(`s${i}`, s.name, hid.length === 1))) }], { width: '360px' });
  if (!body) return;
  hid.forEach(([s, i]) => { if (valOf(body, `s${i}`)) s.x3 = { ...s.x3, hidden: undefined }; });
  C.renderTabs(); C.shell.markDirty();
}
