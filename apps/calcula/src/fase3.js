// Calcula Datos — fase 3: buscar objetivo, Solver, minigráficos, agrupar y esquema,
// segmentación de datos, función IMAGEN y gráficos que siguen matrices dinámicas. GPL-3.0
import { h, toast, modal, form, dropdown, esc, t, tf } from '../../shared/office.js';
import { parseRange, parseAddr, rangeStr, addr, colName } from './workbook.js';

let C = null; // { wb, grid, shell, act, refresh, applyFilter, mountCharts, refreshCharts }
const uid = () => Math.random().toString(36).slice(2, 9);
const X3 = (s = C.wb.sheet) => { if (!s.x3) s.x3 = { sparklines: [], rowGroups: [], colGroups: [], slicers: [] }; return s.x3; };
const num = (v) => (typeof v === 'number' ? v : NaN);

export function init(ctx) {
  C = ctx;
  const g = C.grid;
  g.overlay = drawOverlay;
  g.skipText = (r, c) => !!imageAt(r, c) || !!sparkAt(r, c);
  g.headerOverlay = drawOutline;
  g.overlayClick = outlineClick;
  C.wb.onChange(() => { clearTimeout(init.t); init.t = setTimeout(() => renderSlicers(), 60); });
}

// ================================================================ rangos de matrices dinámicas (A1#)
// "E2#" = toda la matriz que se derrama desde E2
export function resolveRange(text, s = C.wb.sheet) {
  const m = /^\s*(?:'?[^!]*'?!)?\$?([A-Z]{1,3})\$?(\d+)#\s*$/i.exec(String(text || ''));
  if (!m) return parseRange(text);
  const a = parseAddr(`${m[1]}${m[2]}`); if (!a) return null;
  const hf = C.wb.hf; const part = (r, c) => { try { return hf.isCellPartOfArray({ sheet: s.id, row: r, col: c }); } catch { return false; } };
  if (!part(a.r, a.c)) return { r1: a.r, c1: a.c, r2: a.r, c2: a.c };
  let r2 = a.r, c2 = a.c;
  while (part(r2 + 1, a.c) && C.wb.value(r2 + 1, a.c, s) !== null) r2++;
  while (part(a.r, c2 + 1) && C.wb.value(a.r, c2 + 1, s) !== null) c2++;
  return { r1: a.r, c1: a.c, r2, c2 };
}

// ================================================================ buscar objetivo
const ur = () => C.wb.hf._crudOperations?.undoRedo;
function probe(set, get) {
  // Evalúa sin dejar rastro en la pila de deshacer
  const st = ur(); const n0 = st ? st.undoStack.length : 0;
  return { run: (x) => { C.wb.hf.setCellContents({ sheet: C.wb.sheet.id, row: set.r, col: set.c }, [[x]]); return num(C.wb.value(get.r, get.c)); }, end: () => { if (st) st.undoStack.length = n0; } };
}
export function goalSeekCore(target, want, change) {
  const raw0 = C.wb.raw(change.r, change.c);
  const x0 = num(C.wb.value(change.r, change.c)) || 0;
  const P = probe(change, target);
  const f = (x) => P.run(x) - want;
  let best = null;
  try {
    let a = x0, fa = f(a);
    if (Math.abs(fa) < 1e-9) best = a;
    // secante con respaldo de bisección cuando encuentra un cambio de signo
    let b = a === 0 ? 1 : a * 1.1 + 0.1, fb = f(b);
    for (let i = 0; i < 100 && best == null; i++) {
      if (!Number.isFinite(fa) || !Number.isFinite(fb)) break;
      if (Math.abs(fb) < 1e-9) { best = b; break; }
      let c = fb === fa ? b + (b - a || 1) : b - (fb * (b - a)) / (fb - fa);
      if (!Number.isFinite(c)) break;
      a = b; fa = fb; b = c; fb = f(b);
    }
    if (best == null) {
      // búsqueda de un intervalo con cambio de signo y bisección
      let lo = x0, flo = f(lo), step = Math.max(1, Math.abs(x0));
      let hi = null;
      for (let k = 0; k < 60 && hi == null; k++) { for (const cand of [x0 + step, x0 - step]) { const fc = f(cand); if (Number.isFinite(fc) && Math.sign(fc) !== Math.sign(flo)) { hi = cand; break; } } step *= 2; }
      if (hi != null) { let fhi = f(hi); for (let i = 0; i < 200; i++) { const mid = (lo + hi) / 2, fm = f(mid); if (Math.abs(fm) < 1e-9 || Math.abs(hi - lo) < 1e-12) { best = mid; break; } if (Math.sign(fm) === Math.sign(flo)) { lo = mid; flo = fm; } else { hi = mid; fhi = fm; } } if (best == null) best = (lo + hi) / 2; void fhi; }
    }
  } finally {
    C.wb.hf.setCellContents({ sheet: C.wb.sheet.id, row: change.r, col: change.c }, [[raw0 ?? null]]);
    P.end();
  }
  return best;
}
export async function goalSeek() {
  const sel = C.grid.sel;
  const v = await form('Buscar objetivo', [
    { name: 'target', label: 'Definir la celda (con fórmula)', value: addr(sel.r, sel.c) },
    { name: 'want', label: 'Con el valor', type: 'number', value: '' },
    { name: 'change', label: 'Cambiando la celda', value: '' },
  ], 'Buscar');
  if (!v) return;
  const target = parseAddr(v.target.trim().toUpperCase()), change = parseAddr(v.change.trim().toUpperCase());
  if (!target || !change) { toast('Escribe celdas válidas (p. ej. B5).', { error: true }); return; }
  if (typeof C.wb.raw(target.r, target.c) !== 'string' || !String(C.wb.raw(target.r, target.c)).startsWith('=')) { toast('La celda objetivo debe tener una fórmula.', { error: true }); return; }
  if (String(C.wb.raw(change.r, change.c) ?? '').startsWith('=')) { toast('La celda que cambia debe tener un valor, no una fórmula.', { error: true }); return; }
  const x = goalSeekCore(target, Number(v.want), change);
  if (x == null) { toast('No se encontró una solución.', { error: true }); return; }
  const rounded = Math.abs(x - Math.round(x)) < 1e-7 ? Math.round(x) : Math.round(x * 1e10) / 1e10;
  C.act(t('Buscar objetivo'), () => C.wb.setInput(change.r, change.c, String(rounded)));
  C.refreshCharts();
  toast(tf('Solución encontrada: {0} = {1}', v.change.toUpperCase(), rounded.toLocaleString()), { ms: 6000 });
}

// ================================================================ Solver (búsqueda por patrones con penalización)
function parseConstraints(text, varCells) {
  const out = [];
  for (const line of String(text || '').split('\n').map((l) => l.trim()).filter(Boolean)) {
    const mi = /^([A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?)\s+(entero|int|integer|binario|bin)$/i.exec(line);
    if (mi) { const g = parseRange(mi[1].toUpperCase()); for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) out.push({ kind: /bin/i.test(mi[2]) ? 'bin' : 'int', r, c }); continue; }
    const m = /^([A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?)\s*(<=|>=|=|<|>)\s*(.+)$/i.exec(line);
    if (!m) throw new Error(tf('Restricción no válida: {0}', line));
    const g = parseRange(m[1].toUpperCase()); const rhsCell = parseAddr(m[3].trim().toUpperCase());
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) out.push({ kind: 'cmp', r, c, op: m[2].replace('<', '<=').replace('<==', '<=').replace('>', '>=').replace('>==', '>='), rhs: rhsCell ? { cell: rhsCell } : Number(m[3].replace(',', '.')) });
  }
  void varCells;
  return out;
}
// Simplex de dos fases: maximiza c·y con filas A·y (<=, >=, =) b y y >= 0
export function simplex(c, A, b, types) {
  const EPS = 1e-9; const m = A.length, n = c.length;
  A = A.map((r) => r.slice()); b = b.slice(); types = types.slice();
  for (let i = 0; i < m; i++) if (b[i] < 0) { A[i] = A[i].map((v) => -v); b[i] = -b[i]; types[i] = types[i] === '<=' ? '>=' : types[i] === '>=' ? '<=' : '='; }
  const nS = types.filter((x) => x !== '=').length, nA = types.filter((x) => x !== '<=').length;
  const cols = n + nS + nA; const T = []; const basis = []; let si = 0, ai = 0;
  for (let i = 0; i < m; i++) {
    const row = new Array(cols + 1).fill(0); A[i].forEach((v, j) => { row[j] = v; });
    if (types[i] === '<=') { row[n + si] = 1; basis[i] = n + si; si++; }
    else if (types[i] === '>=') { row[n + si] = -1; si++; row[n + nS + ai] = 1; basis[i] = n + nS + ai; ai++; }
    else { row[n + nS + ai] = 1; basis[i] = n + nS + ai; ai++; }
    row[cols] = b[i]; T.push(row);
  }
  const banned = new Set();
  const pivot = (r, k) => { const pv = T[r][k]; for (let j = 0; j <= cols; j++) T[r][j] /= pv; for (let i = 0; i < m; i++) if (i !== r && Math.abs(T[i][k]) > EPS) { const f = T[i][k]; for (let j = 0; j <= cols; j++) T[i][j] -= f * T[r][j]; } basis[r] = k; };
  const optimize = (cost) => {
    for (let it = 0; it < 5000; it++) {
      let enter = -1;
      for (let j = 0; j < cols; j++) { if (banned.has(j)) continue; let rc = cost[j]; for (let i = 0; i < m; i++) rc -= cost[basis[i]] * T[i][j]; if (rc > EPS) { enter = j; break; } }
      if (enter < 0) return 'ok';
      let row = -1, best = Infinity;
      for (let i = 0; i < m; i++) if (T[i][enter] > EPS) { const q = T[i][cols] / T[i][enter]; if (q < best - EPS || (Math.abs(q - best) <= EPS && basis[i] < basis[row])) { best = q; row = i; } }
      if (row < 0) return 'unbounded';
      pivot(row, enter);
    }
    return 'ok';
  };
  if (nA) {
    const c1 = new Array(cols).fill(0); for (let j = n + nS; j < cols; j++) c1[j] = -1;
    optimize(c1);
    const art = basis.reduce((a, k, i) => a + (k >= n + nS ? T[i][cols] : 0), 0);
    if (art > 1e-7) return { status: 'infeasible' };
    for (let i = 0; i < m; i++) if (basis[i] >= n + nS) { const k = [...Array(n + nS).keys()].find((j) => Math.abs(T[i][j]) > EPS); if (k != null) pivot(i, k); }
    for (let j = n + nS; j < cols; j++) banned.add(j);
  }
  const c2 = new Array(cols).fill(0); c.forEach((v, j) => { c2[j] = v; });
  if (optimize(c2) === 'unbounded') return { status: 'unbounded' };
  const y = new Array(n).fill(0); basis.forEach((k, i) => { if (k < n) y[k] = T[i][cols]; });
  return { status: 'ok', y, z: c.reduce((a, v, j) => a + v * y[j], 0) };
}

export function solverCore({ target, goal, value, vars, constraints }) {
  const st = ur(); const n0 = st ? st.undoStack.length : 0;
  const sid = C.wb.sheet.id;
  const raw0 = vars.map((p) => C.wb.raw(p.r, p.c));
  const isInt = vars.map((p) => constraints.some((k) => (k.kind === 'int' || k.kind === 'bin') && k.r === p.r && k.c === p.c));
  const isBin = vars.map((p) => constraints.some((k) => k.kind === 'bin' && k.r === p.r && k.c === p.c));
  const setAll = (xs) => C.wb.hf.batch(() => vars.forEach((p, i) => C.wb.hf.setCellContents({ sheet: sid, row: p.r, col: p.c }, [[xs[i]]])));
  const val = (p) => num(C.wb.value(p.r, p.c));
  const cmps = constraints.filter((k) => k.kind === 'cmp');
  const evalAll = (xs) => { setAll(xs); return { obj: val(target), lhs: cmps.map((k) => val(k)), rhs: cmps.map((k) => (typeof k.rhs === 'number' ? k.rhs : val(k.rhs.cell))) }; };
  const score = (xs0) => {
    const xs = xs0.map((x, i) => (isBin[i] ? Math.min(1, Math.max(0, Math.round(x))) : isInt[i] ? Math.round(x) : x));
    const e = evalAll(xs); if (!Number.isFinite(e.obj)) return { f: Infinity, xs };
    let pen = 0;
    cmps.forEach((k, q) => { const d = k.op === '<=' ? e.lhs[q] - e.rhs[q] : k.op === '>=' ? e.rhs[q] - e.lhs[q] : Math.abs(e.lhs[q] - e.rhs[q]); if (d > 1e-9) pen += d; });
    const base = goal === 'max' ? -e.obj : goal === 'min' ? e.obj : Math.abs(e.obj - value);
    return { f: base + pen * 1e6, xs, pen, obj: e.obj };
  };
  // ¿Es un modelo lineal? Se mide la pendiente de cada fórmula y se comprueba en otro punto.
  const linear = () => {
    const n = vars.length; const x0 = vars.map(() => 0);
    const e0 = evalAll(x0); const fs = [e0.obj, ...e0.lhs, ...e0.rhs]; if (!fs.every(Number.isFinite)) return null;
    const G = fs.map(() => new Array(n).fill(0));
    for (let i = 0; i < n; i++) { const x = x0.slice(); x[i] = 1; const e = evalAll(x); [e.obj, ...e.lhs, ...e.rhs].forEach((v, k) => { G[k][i] = v - fs[k]; }); }
    for (const probe of [vars.map((_, i) => 3.7 + i * 1.3), vars.map((_, i) => -2.1 * (i + 1))]) {
      const e = evalAll(probe); const got = [e.obj, ...e.lhs, ...e.rhs];
      if (!got.every((v, k) => Math.abs(v - (fs[k] + G[k].reduce((a, g, i) => a + g * probe[i], 0))) <= 1e-7 * (1 + Math.abs(v)))) return null;
    }
    return { fs, G };
  };
  const solveLinear = (L) => {
    const n = vars.length, q = cmps.length;
    // x = p - m (variables libres); filas: (lhs - rhs) op 0
    const rows = [], rb = [], rt = [];
    cmps.forEach((k, j) => { const g = L.G[1 + j].map((v, i) => v - L.G[1 + q + j][i]); const c0 = L.fs[1 + j] - L.fs[1 + q + j]; rows.push([...g, ...g.map((v) => -v)]); rb.push(-c0); rt.push(k.op); });
    if (goal === 'value') { rows.push([...L.G[0], ...L.G[0].map((v) => -v)]); rb.push(value - L.fs[0]); rt.push('='); }
    vars.forEach((_, i) => { if (isBin[i]) { const r = new Array(2 * n).fill(0); r[i] = 1; r[n + i] = -1; rows.push(r); rb.push(1); rt.push('<='); const r2 = r.slice(); rows.push(r2); rb.push(0); rt.push('>='); } });
    const sign = goal === 'min' ? -1 : 1; const cvec = goal === 'value' ? new Array(2 * n).fill(0) : [...L.G[0].map((v) => sign * v), ...L.G[0].map((v) => -sign * v)];
    const toX = (y) => vars.map((_, i) => y[i] - y[n + i]);
    // ramificación y acotamiento para enteros
    let best = null, nodes = 0;
    const bb = (extraRows) => {
      if (++nodes > 400) return;
      const r = simplex(cvec, [...rows, ...extraRows.map((e) => e.row)], [...rb, ...extraRows.map((e) => e.b)], [...rt, ...extraRows.map((e) => e.t)]);
      if (r.status !== 'ok') return;
      if (best && r.z <= best.z + 1e-9) return;
      const x = toX(r.y); const k = x.findIndex((v, i) => isInt[i] && Math.abs(v - Math.round(v)) > 1e-6);
      if (k < 0) { best = { z: r.z, x }; return; }
      const unit = new Array(2 * n).fill(0); unit[k] = 1; unit[n + k] = -1;
      bb([...extraRows, { row: unit, b: Math.floor(x[k]), t: '<=' }]); bb([...extraRows, { row: unit, b: Math.ceil(x[k]), t: '>=' }]);
    };
    bb([]);
    return best ? best.x : null;
  };
  let bestRes = null;
  try {
    const L = linear();
    const xl = L ? solveLinear(L) : null;
    if (xl) bestRes = score(xl);
    if (!bestRes || bestRes.pen > 1e-6) {
      // modelo no lineal: búsqueda por patrones con penalización
      const start = raw0.map((v) => (typeof v === 'number' ? v : 0));
      const t0 = performance.now();
      for (const seed of [start, start.map(() => 0), start.map(() => 1)]) {
        if (bestRes && performance.now() - t0 > 6000) break;
        let x = seed.slice(); let cur = score(x); let step = Math.max(1, ...x.map(Math.abs)) / 2;
        for (let it = 0; it < 3000 && step > 1e-6 && performance.now() - t0 < 12000; it++) {
          let improved = false;
          const dirs = [];
          for (let i = 0; i < x.length; i++) for (const d of [1, -1]) { const v = new Array(x.length).fill(0); v[i] = d; dirs.push(v); }
          for (let i = 0; i < x.length; i++) for (let j = i + 1; j < x.length; j++) for (const [a, b] of [[1, 1], [1, -1], [-1, 1], [-1, -1]]) { const v = new Array(x.length).fill(0); v[i] = a; v[j] = b; dirs.push(v); }
          for (const d of dirs) {
            const y = x.map((xi, i) => xi + d[i] * (isInt[i] ? Math.max(1, Math.round(step)) : step));
            const r = score(y); if (r.f < cur.f - 1e-12) { x = y; cur = r; improved = true; break; }
          }
          if (!improved) { if (isInt.every(Boolean) && step <= 1) break; step /= 2; }
        }
        if (!bestRes || cur.f < bestRes.f) bestRes = cur;
      }
    }
  } finally {
    C.wb.hf.batch(() => vars.forEach((p, i) => C.wb.hf.setCellContents({ sheet: sid, row: p.r, col: p.c }, [[raw0[i] ?? null]])));
    if (st) st.undoStack.length = n0;
  }
  return bestRes;
}
const snap = (x) => (Math.abs(x - Math.round(x)) < 1e-4 ? Math.round(x) : Math.round(x * 1e6) / 1e6);
export async function solver() {
  const s = C.wb.sheet; const last = X3(s).solver || {};
  const v = await form('Solver', [
    { name: 'target', label: 'Celda objetivo (con fórmula)', value: last.target || addr(C.grid.sel.r, C.grid.sel.c) },
    { name: 'goal', label: 'Objetivo', type: 'select', value: last.goal || 'max', options: [['max', 'Máximo'], ['min', 'Mínimo'], ['value', 'Valor de:']] },
    { name: 'value', label: 'Valor (si elegiste “Valor de”)', type: 'number', value: last.value ?? 0 },
    { name: 'vars', label: 'Cambiando las celdas', value: last.vars || '' },
    { name: 'cons', label: 'Restricciones, una por línea (p. ej. B2 <= 100, B2:B4 >= 0, B3 entero, B4 binario)', type: 'textarea', rows: 4, value: last.cons || '' },
    { type: 'note', label: 'Método: Simplex (con enteros por ramificación) si el modelo es lineal; si no, búsqueda directa con penalización.' },
  ], 'Resolver');
  if (!v) return;
  const target = parseAddr(v.target.trim().toUpperCase()); const g = parseRange(v.vars.trim().toUpperCase());
  if (!target || !g) { toast('Escribe la celda objetivo y las celdas que cambian.', { error: true }); return; }
  let constraints; try { constraints = parseConstraints(v.cons); } catch (e) { toast(e.message, { error: true }); return; }
  const vars = []; for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) vars.push({ r, c });
  if (vars.length > 30) { toast('Solver admite hasta 30 celdas variables.', { error: true }); return; }
  X3(s).solver = { target: v.target, goal: v.goal, value: v.value, vars: v.vars, cons: v.cons };
  const res = solverCore({ target, goal: v.goal, value: Number(v.value), vars, constraints });
  if (!res || !Number.isFinite(res.f)) { toast('Solver no encontró una solución.', { error: true }); return; }
  const feasible = !(res.pen > 1e-6);
  const keep = await modal({ title: feasible ? t('Solver encontró una solución') : t('Solver no pudo cumplir todas las restricciones'), body: h('div', {}, h('p', {}, tf('Valor de la celda objetivo: {0}', snap(res.obj))), h('p', { class: 'muted' }, vars.map((p, i) => `${addr(p.r, p.c)} = ${snap(res.xs[i])}`).join(' · '))), buttons: [{ label: 'Restaurar valores originales', value: false }, { label: 'Conservar la solución', primary: true, value: true }] });
  if (!keep) return;
  C.act(t('Solver'), () => vars.forEach((p, i) => C.wb.setInput(p.r, p.c, String(snap(res.xs[i])))), { batch: true });
  C.refreshCharts();
}

// ================================================================ minigráficos
const sparkAt = (r, c) => X3().sparklines.find((sp) => sp.r === r && sp.c === c);
export async function sparklineDialog(type = 'line') {
  const g = C.grid.range;
  const v = await form('Insertar minigráficos', [
    { name: 'type', label: 'Tipo', type: 'select', value: type, options: [['line', 'Línea'], ['column', 'Columna'], ['winloss', 'Ganancia o pérdida']] },
    { name: 'data', label: 'Rango de datos (p. ej. B2:E5)', value: rangeStr(g) },
    { name: 'loc', label: 'Ubicación (una celda por fila o columna de datos)', value: '' },
    { name: 'color', label: 'Color', type: 'select', value: '#0E9C74', options: [['#0E9C74', 'Verde'], ['#5B48E8', 'Índigo'], ['#1F8FD6', 'Azul'], ['#DC3A63', 'Rosa'], ['#E07B2A', 'Naranja']] },
  ], 'Insertar');
  if (!v) return;
  const data = parseRange(v.data.trim().toUpperCase()), loc = parseRange(v.loc.trim().toUpperCase());
  if (!data || !loc) { toast('Escribe el rango de datos y la ubicación.', { error: true }); return; }
  const cells = []; for (let r = loc.r1; r <= loc.r2; r++) for (let c = loc.c1; c <= loc.c2; c++) cells.push({ r, c });
  const vertical = loc.c1 === loc.c2 && cells.length > 1;
  const nSeries = vertical ? data.r2 - data.r1 + 1 : data.c2 - data.c1 + 1;
  if (cells.length !== 1 && cells.length !== nSeries) { toast(tf('La ubicación debe tener {0} celdas.', nSeries), { error: true }); return; }
  const list = cells.map((cell, i) => {
    const rg = cells.length === 1 ? data : vertical ? { r1: data.r1 + i, r2: data.r1 + i, c1: data.c1, c2: data.c2 } : { r1: data.r1, r2: data.r2, c1: data.c1 + i, c2: data.c1 + i };
    return { id: uid(), r: cell.r, c: cell.c, range: rangeStr(rg), type: v.type, color: v.color };
  });
  const s = C.wb.sheet;
  C.act(t('Minigráficos'), () => { const x = X3(s); x.sparklines = [...x.sparklines.filter((sp) => !list.some((n) => n.r === sp.r && n.c === sp.c)), ...list]; }, { hf: false });
}
export function clearSparklines() {
  const g = C.grid.range; const s = C.wb.sheet;
  C.act(t('Borrar minigráficos'), () => { const x = X3(s); x.sparklines = x.sparklines.filter((sp) => !(sp.r >= g.r1 && sp.r <= g.r2 && sp.c >= g.c1 && sp.c <= g.c2)); }, { hf: false });
}
function sparkValues(sp) {
  const g = parseRange(sp.range); if (!g) return [];
  const out = []; for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { const v = C.wb.value(r, c); out.push(typeof v === 'number' ? v : null); }
  return out;
}
function drawSpark(ctx, sp, rc) {
  const vals = sparkValues(sp); const nums = vals.filter((v) => v != null); if (!nums.length) return;
  const pad = 3, x0 = rc.x + pad, y0 = rc.y + pad, w = rc.w - pad * 2, hh = rc.h - pad * 2;
  const min = Math.min(...nums), max = Math.max(...nums); const span = max - min || 1;
  ctx.save(); ctx.beginPath(); ctx.rect(rc.x + 1, rc.y + 1, rc.w - 2, rc.h - 2); ctx.clip();
  const n = vals.length; const step = n > 1 ? w / (n - 1) : 0;
  if (sp.type === 'line') {
    ctx.strokeStyle = sp.color; ctx.lineWidth = 1.5; ctx.beginPath(); let started = false;
    vals.forEach((v, i) => { if (v == null) return; const px = x0 + i * step, py = y0 + hh - ((v - min) / span) * hh; if (!started) { ctx.moveTo(px, py); started = true; } else ctx.lineTo(px, py); });
    ctx.stroke();
    const iMax = vals.indexOf(max), iMin = vals.indexOf(min);
    for (const [i, col] of [[iMax, sp.color], [iMin, '#D93025']]) { ctx.fillStyle = col; ctx.beginPath(); ctx.arc(x0 + i * step, y0 + hh - ((vals[i] - min) / span) * hh, 2.2, 0, Math.PI * 2); ctx.fill(); }
  } else {
    const bw = Math.max(1, w / n - 2); const zeroY = sp.type === 'winloss' ? y0 + hh / 2 : y0 + hh - ((Math.max(0, min) - min) / span) * hh;
    vals.forEach((v, i) => {
      if (v == null) return; const bx = rc.x + pad + i * (w / n) + 1;
      if (sp.type === 'winloss') { ctx.fillStyle = v >= 0 ? sp.color : '#D93025'; ctx.fillRect(bx, v >= 0 ? y0 + 1 : zeroY + 1, bw, hh / 2 - 2); return; }
      const top = y0 + hh - ((v - min) / span) * hh; ctx.fillStyle = v < 0 ? '#D93025' : sp.color; ctx.fillRect(bx, Math.min(top, zeroY), bw, Math.max(1, Math.abs(zeroY - top)));
    });
  }
  ctx.restore();
}

// ================================================================ IMAGEN en celdas
const imgCache = new Map();
function imageAt(r, c) {
  const raw = C.wb.raw(r, c);
  if (typeof raw !== 'string' || !/^=\s*IMAGE\s*\(/i.test(raw)) return null;
  const v = C.wb.value(r, c); return typeof v === 'string' ? v : null;
}
function getImage(src) {
  if (imgCache.has(src)) return imgCache.get(src);
  const im = new Image(); im.referrerPolicy = 'no-referrer'; im.onload = () => C.grid.draw(); im.onerror = () => { im.failed = true; C.grid.draw(); }; im.src = src;
  imgCache.set(src, im); return im;
}

// ================================================================ dibujo sobre la cuadrícula
function drawOverlay(ctx, r1, r2, c1, c2) {
  const x = X3();
  for (const sp of x.sparklines) { if (sp.r < r1 || sp.r > r2 || sp.c < c1 || sp.c > c2) continue; const rc = C.grid.cellRect(sp.r, sp.c); if (rc.h > 2) drawSpark(ctx, sp, rc); }
  const dims = C.wb.hf.getSheetDimensions(C.wb.sheet.id); if (!dims.height) return; // imágenes = valores de celda; no hace falta recorrer los formatos
  const used = { r1: 0, c1: 0, r2: dims.height - 1, c2: dims.width - 1 };
  for (let r = Math.max(r1, used.r1); r <= Math.min(r2, used.r2); r++) for (let c = Math.max(c1, used.c1); c <= Math.min(c2, used.c2); c++) {
    const src = imageAt(r, c); if (!src) continue;
    const rc = C.grid.cellRect(r, c); const im = getImage(src);
    if (im.failed) { ctx.fillStyle = '#9AA0B4'; ctx.font = '11px Inter, sans-serif'; ctx.fillText('🖼 ✕', rc.x + 4, rc.y + rc.h / 2); continue; }
    if (!im.complete || !im.naturalWidth) continue;
    const k = Math.min((rc.w - 4) / im.naturalWidth, (rc.h - 4) / im.naturalHeight); const w = im.naturalWidth * k, hh = im.naturalHeight * k;
    try { ctx.drawImage(im, rc.x + (rc.w - w) / 2, rc.y + (rc.h - hh) / 2, w, hh); } catch { /* */ }
  }
}

// ================================================================ agrupar y esquema
export function group(ungroup = false) {
  const g = C.grid.range, s = C.wb.sheet;
  const byCols = g.r1 === 0 && g.r2 >= s.nRows - 1 && !(g.c1 === 0 && g.c2 >= s.nCols - 1);
  const asRows = !byCols;
  C.act(ungroup ? t('Desagrupar') : t('Agrupar'), () => {
    const x = X3(s); const key = asRows ? 'rowGroups' : 'colGroups';
    const a = asRows ? g.r1 : g.c1, b = asRows ? g.r2 : g.c2;
    if (ungroup) {
      for (const gr of x[key].filter((q) => q.a <= b && q.b >= a)) expandGroup(s, gr, asRows);
      x[key] = x[key].filter((q) => !(q.a <= b && q.b >= a));
    } else x[key] = [...x[key].filter((q) => !(q.a === a && q.b === b)), { id: uid(), a, b, collapsed: false }].sort((p, q) => p.a - q.a);
  }, { hf: false });
}
function collapseGroup(s, gr, rows) {
  gr.collapsed = true;
  if (rows) for (let r = gr.a; r <= gr.b; r++) s.hiddenRows.add(r);
  else { gr.widths = {}; for (let c = gr.a; c <= gr.b; c++) { gr.widths[c] = s.colW.get(c); s.colW.set(c, 0); } }
}
function expandGroup(s, gr, rows) {
  gr.collapsed = false;
  if (rows) for (let r = gr.a; r <= gr.b; r++) s.hiddenRows.delete(r);
  else { for (let c = gr.a; c <= gr.b; c++) { const w = gr.widths?.[c]; if (w == null) s.colW.delete(c); else s.colW.set(c, w); } gr.widths = null; }
}
export function toggleGroup(kind, id, collapse) {
  const s = C.wb.sheet; const x = X3(s); const gr = x[kind === 'row' ? 'rowGroups' : 'colGroups'].find((q) => q.id === id); if (!gr) return;
  C.act(collapse ? t('Ocultar detalle') : t('Mostrar detalle'), () => { const copy = { ...gr }; if (collapse) collapseGroup(s, copy, kind === 'row'); else expandGroup(s, copy, kind === 'row'); x[kind === 'row' ? 'rowGroups' : 'colGroups'] = x[kind === 'row' ? 'rowGroups' : 'colGroups'].map((q) => (q.id === id ? copy : q)); }, { hf: false });
  C.mountCharts();
}
export function showDetail(show) {
  const g = C.grid.range; const x = X3();
  for (const gr of x.rowGroups) if (gr.a <= g.r2 + 1 && gr.b >= g.r1 - 1 && gr.collapsed === show) toggleGroup('row', gr.id, !show);
  for (const gr of x.colGroups) if (gr.a <= g.c2 + 1 && gr.b >= g.c1 - 1 && gr.collapsed === show) toggleGroup('col', gr.id, !show);
}
// Filas ocultas por grupos plegados (el filtro las respeta)
export function outlineHiddenRows(s = C.wb.sheet) { const out = []; for (const gr of X3(s).rowGroups) if (gr.collapsed) for (let r = gr.a; r <= gr.b; r++) out.push(r); return out; }
const BOX = 9;
function outlineBoxes() {
  const x = X3(); const g = C.grid; const out = [];
  for (const gr of x.rowGroups) { const r = gr.b + 1; const y = g.rowTop(r), hh = C.wb.sheet.rh(r) || 18; out.push({ kind: 'row', gr, x: 2, y: y + (hh - BOX) / 2, line: { y1: g.rowTop(gr.a), y2: y } }); }
  for (const gr of x.colGroups) { const c = gr.b + 1; const xx = g.colLeft(c), w = C.wb.sheet.cw(c) || 40; out.push({ kind: 'col', gr, x: xx + 2, y: 2, line: { x1: g.colLeft(gr.a), x2: xx }, w }); }
  return out;
}
function drawOutline(ctx) {
  for (const b of outlineBoxes()) {
    ctx.strokeStyle = '#7B8094'; ctx.lineWidth = 1;
    if (!b.gr.collapsed) { ctx.beginPath(); if (b.kind === 'row') { ctx.moveTo(6.5, b.line.y1 + 3); ctx.lineTo(6.5, b.y); } else { ctx.moveTo(b.line.x1 + 3, 6.5); ctx.lineTo(b.x, 6.5); } ctx.stroke(); }
    ctx.fillStyle = '#fff'; ctx.fillRect(b.x, b.y, BOX, BOX); ctx.strokeRect(b.x + 0.5, b.y + 0.5, BOX - 1, BOX - 1);
    ctx.fillStyle = '#3A3F52'; ctx.fillRect(b.x + 2, b.y + 4, BOX - 4, 1); if (b.gr.collapsed) ctx.fillRect(b.x + 4, b.y + 2, 1, BOX - 4);
  }
}
function outlineClick(px, py) {
  for (const b of outlineBoxes()) if (px >= b.x - 1 && px <= b.x + BOX + 1 && py >= b.y - 1 && py <= b.y + BOX + 1) { toggleGroup(b.kind, b.gr.id, !b.gr.collapsed); return true; }
  return false;
}

// ================================================================ segmentación de datos
export async function insertSlicer() {
  const s = C.wb.sheet; const g0 = C.grid.range;
  const reg = g0.r1 === g0.r2 && g0.c1 === g0.c2 ? C.wb.region(g0.r1, g0.c1) : g0;
  if (reg.r2 <= reg.r1) { toast('Selecciona una tabla con encabezados.', { error: true }); return; }
  const heads = []; for (let c = reg.c1; c <= reg.c2; c++) heads.push([String(c), C.wb.display(reg.r1, c).text || colName(c)]);
  const v = await form('Insertar segmentación de datos', [{ name: 'col', label: 'Columna para filtrar', type: 'select', value: heads[0][0], options: heads }], 'Insertar');
  if (!v) return;
  const col = Number(v.col); const p = C.grid.prefix();
  C.act(t('Segmentación de datos'), () => {
    if (!s.filter || s.filter.range.r1 !== reg.r1 || s.filter.range.c1 !== reg.c1) s.filter = { range: { ...reg }, hidden: new Map() };
    X3(s).slicers.push({ id: uid(), col, title: C.wb.display(reg.r1, col).text || colName(col), x: (p.cx[reg.c2 + 1] || 0) + 24 + X3(s).slicers.length * 20, y: (p.ry[reg.r1] || 0) + X3(s).slicers.length * 20 });
  }, { hf: false });
  renderSlicers();
}
export function renderSlicers() {
  if (!C) return;
  const layer = C.grid.chartLayer; layer.querySelectorAll('.slicer').forEach((n) => n.remove());
  const s = C.wb.sheet; const f = s.filter;
  for (const sl of X3(s).slicers) {
    if (!f) continue;
    const values = new Set(); for (let r = f.range.r1 + 1; r <= f.range.r2; r++) values.add(C.wb.display(r, sl.col).text);
    const hidden = f.hidden.get(sl.col) || new Set();
    const sorted = [...values].sort((a, b) => (a === '' ? 1 : b === '' ? -1 : a.localeCompare(b, 'es', { numeric: true })));
    const box = h('div', { class: 'slicer chartbox', 'data-slicer': sl.id, style: { width: '190px' } },
      h('div', { class: 'sl-head' }, h('b', {}, sl.title), h('button', { class: 'sl-clear', title: t('Borrar filtro'), onclick: (e) => { e.stopPropagation(); setSlicer(sl, null); } }, '⨯'), h('button', { class: 'sl-del', title: t('Quitar segmentación'), onclick: (e) => { e.stopPropagation(); removeSlicer(sl.id); } }, '🗑')),
      h('div', { class: 'sl-items' }, ...sorted.map((val) => h('button', { class: `sl-item${hidden.has(val) ? '' : ' on'}`, onclick: (e) => { e.stopPropagation(); setSlicer(sl, val, e.ctrlKey || e.metaKey); } }, val || t('(Vacías)')))));
    box.dataset.x = sl.x; box.dataset.y = sl.y;
    box.addEventListener('mousedown', (e) => {
      if (e.target.closest('button')) return;
      const x0 = e.clientX, y0 = e.clientY, sx = sl.x, sy = sl.y;
      const move = (ev) => { sl.x = Math.max(0, sx + (ev.clientX - x0) / C.grid.zoom); sl.y = Math.max(0, sy + (ev.clientY - y0) / C.grid.zoom); C.grid.positionCharts(); positionSlicers(); };
      const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); C.shell.markDirty(); };
      addEventListener('mousemove', move); addEventListener('mouseup', up);
    });
    layer.append(box);
  }
  positionSlicers();
}
export function positionSlicers() {
  const g = C.grid; const p = g.prefix(); const { fr, fc, FW, FH } = g.frozen();
  for (const box of g.chartLayer.querySelectorAll('.slicer')) {
    const sl = X3().slicers.find((q) => q.id === box.dataset.slicer); if (!sl) continue;
    const x = 46 + FW + sl.x - p.cx[fc] - g.scrollX, y = 24 + FH + sl.y - p.ry[fr] - g.scrollY;
    box.style.left = `${x}px`; box.style.top = `${y}px`;
  }
}
function setSlicer(sl, val, add) {
  const s = C.wb.sheet; const f = s.filter; if (!f) return;
  const values = new Set(); for (let r = f.range.r1 + 1; r <= f.range.r2; r++) values.add(C.wb.display(r, sl.col).text);
  C.act(t('Segmentación de datos'), () => {
    if (val == null) { f.hidden.delete(sl.col); } else {
      const cur = new Set(f.hidden.get(sl.col) || []);
      const shown = new Set([...values].filter((v) => !cur.has(v)));
      let next;
      if (add) { if (shown.has(val) && shown.size > 1) shown.delete(val); else shown.add(val); next = shown; } else next = cur.size === 0 || !(shown.size === 1 && shown.has(val)) ? new Set([val]) : new Set(values);
      const hid = new Set([...values].filter((v) => !next.has(v)));
      if (hid.size) f.hidden.set(sl.col, hid); else f.hidden.delete(sl.col);
    }
    C.applyFilter();
  }, { hf: false });
  C.refreshCharts(); renderSlicers();
}
function removeSlicer(id) { const s = C.wb.sheet; C.act(t('Quitar segmentación'), () => { X3(s).slicers = X3(s).slicers.filter((q) => q.id !== id); }, { hf: false }); renderSlicers(); }
