// Calcula Datos — funciones nuevas para HyperFormula (fase 3):
// LAMBDA y sus compañeras (MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY), LET,
// TAKE/DROP/TOCOL/TOROW/CHOOSECOLS/CHOOSEROWS/WRAPROWS, TEXTBEFORE/TEXTAFTER e IMAGE. GPL-3.0
import { HyperFormula, FunctionPlugin, FunctionArgumentType, ArraySize, SimpleRangeValue, CellError, ErrorType } from 'hyperformula';

const ERR = (msg) => new CellError(ErrorType.VALUE, msg);
const num = (v) => (typeof v === 'number' ? v : (v && typeof v === 'object' && typeof v.val === 'number' ? v.val : (typeof v === 'boolean' ? Number(v) : NaN)));
const isRange = (v) => v instanceof SimpleRangeValue;
const rows2d = (v) => (isRange(v) ? v.data : [[v]]);
const plain = (v) => (v && typeof v === 'object' && !(v instanceof CellError) && typeof v.val === 'number' ? v.val : v);
const out = (m) => (m.length === 1 && m[0].length === 1 ? m[0][0] : SimpleRangeValue.onlyValues(m));

// Un valor calculado -> nodo del árbol de la fórmula (para sustituir parámetros de LAMBDA)
function valueAst(v) {
  if (isRange(v)) return { type: 'ARRAY', args: v.data.map((row) => row.map(scalarAst)) };
  return scalarAst(v);
}
function scalarAst(v) {
  v = plain(v);
  if (typeof v === 'number') return v < 0 ? { type: 'MINUS_UNARY_OP', value: { type: 'NUMBER', value: -v } } : { type: 'NUMBER', value: v };
  if (typeof v === 'string') return { type: 'STRING', value: v };
  if (typeof v === 'boolean') return { type: 'FUNCTION_CALL', procedureName: v ? 'TRUE' : 'FALSE', args: [] };
  if (v instanceof CellError) return { type: 'ERROR', error: v };
  return { type: 'EMPTY' };
}
function substitute(ast, binds) {
  if (!ast || typeof ast !== 'object') return ast;
  if (ast.type === 'NAMED_EXPRESSION') { const k = ast.expressionName.toUpperCase(); if (binds.has(k)) return binds.get(k); return ast; }
  const o = { ...ast };
  if (Array.isArray(ast.args)) o.args = ast.args.map((a) => (Array.isArray(a) ? a.map((b) => substitute(b, binds)) : substitute(a, binds)));
  for (const k of ['left', 'right', 'value', 'expression']) if (ast[k] && typeof ast[k] === 'object' && ast[k].type) o[k] = substitute(ast[k], binds);
  return o;
}
const isLambda = (ast) => ast?.type === 'FUNCTION_CALL' && ast.procedureName === 'LAMBDA';
const unwrap = (ast) => (ast?.type === 'PARENTHESES' ? unwrap(ast.expression) : ast);

class OficinaPlugin extends FunctionPlugin {
  // Aplica una LAMBDA (nodo del árbol) a una lista de valores
  apply(lam, vals, state) {
    lam = unwrap(lam);
    if (!isLambda(lam)) return ERR('Se esperaba una función LAMBDA');
    const params = lam.args.slice(0, -1), body = lam.args[lam.args.length - 1];
    if (params.length !== vals.length || params.some((p) => p.type !== 'NAMED_EXPRESSION')) return ERR('LAMBDA: número de parámetros incorrecto');
    const binds = new Map(params.map((p, i) => [p.expressionName.toUpperCase(), valueAst(vals[i])]));
    return this.evaluateAst(substitute(body, binds), state);
  }
  val(ast, state) { return this.evaluateAst(ast, state); }
  lit(ast) { ast = unwrap(ast); if (ast?.type === 'NUMBER') return ast.value; if (ast?.type === 'MINUS_UNARY_OP' && ast.value?.type === 'NUMBER') return -ast.value.value; return null; }
  size(ast, state) { const s = this.arraySizeForAst(ast, state); return s.isScalar?.() && !s.isRef ? s : new ArraySize(s.width, s.height); }

  lambda() { return ERR('LAMBDA se usa dentro de MAP, REDUCE, SCAN, BYROW, BYCOL o MAKEARRAY'); }
  let_(ast, state) {
    const a = ast.args; if (a.length < 3 || a.length % 2 === 0) return ERR('LET necesita pares nombre, valor y un cálculo');
    const binds = new Map();
    for (let i = 0; i < a.length - 1; i += 2) {
      if (a[i].type !== 'NAMED_EXPRESSION') return ERR('LET: nombre no válido');
      const v = this.evaluateAst(substitute(a[i + 1], binds), state);
      binds.set(a[i].expressionName.toUpperCase(), valueAst(v));
    }
    return this.evaluateAst(substitute(a[a.length - 1], binds), state);
  }
  letSize(ast, state) { const a = ast.args; return a.length ? this.size(a[a.length - 1], state) : ArraySize.scalar(); }
  map(ast, state) {
    const [arr, lam] = [ast.args[0], ast.args[ast.args.length - 1]];
    const m = rows2d(this.val(arr, state));
    return out(m.map((row) => row.map((v) => plain(this.apply(lam, [v], state)))));
  }
  mapSize(ast, state) { return this.size(ast.args[0], state); }
  reduce(ast, state) {
    if (ast.args.length < 3) return ERR('REDUCE(inicial, matriz, LAMBDA(acum, valor, …))');
    let acc = this.val(ast.args[0], state);
    for (const row of rows2d(this.val(ast.args[1], state))) for (const v of row) acc = plain(this.apply(ast.args[2], [acc, v], state));
    return acc;
  }
  scan(ast, state) {
    let acc = this.val(ast.args[0], state);
    const m = rows2d(this.val(ast.args[1], state));
    return out(m.map((row) => row.map((v) => { acc = plain(this.apply(ast.args[2], [acc, v], state)); return acc; })));
  }
  scanSize(ast, state) { return this.size(ast.args[1], state); }
  byrow(ast, state) {
    const m = rows2d(this.val(ast.args[0], state));
    return out(m.map((row) => [plain(this.apply(ast.args[1], [SimpleRangeValue.onlyValues([row])], state))]));
  }
  byrowSize(ast, state) { const s = this.size(ast.args[0], state); return new ArraySize(1, s.height); }
  bycol(ast, state) {
    const m = rows2d(this.val(ast.args[0], state));
    return out([m[0].map((_, c) => plain(this.apply(ast.args[1], [SimpleRangeValue.onlyValues(m.map((r) => [r[c]]))], state)))]);
  }
  bycolSize(ast, state) { const s = this.size(ast.args[0], state); return new ArraySize(s.width, 1); }
  makearray(ast, state) {
    const R = num(this.val(ast.args[0], state)), Cc = num(this.val(ast.args[1], state));
    if (!(R >= 1 && Cc >= 1)) return ERR('MAKEARRAY: tamaño no válido');
    return out(Array.from({ length: R }, (_, i) => Array.from({ length: Cc }, (__, j) => plain(this.apply(ast.args[2], [i + 1, j + 1], state)))));
  }
  makearraySize(ast) { const r = this.lit(ast.args[0]), c = this.lit(ast.args[1]); return r >= 1 && c >= 1 ? new ArraySize(c, r) : ArraySize.error(); }

  // ---- forma de matrices
  take(ast, state) {
    const m = rows2d(this.val(ast.args[0], state)); const r = ast.args[1] ? num(this.val(ast.args[1], state)) : m.length; const c = ast.args[2] ? num(this.val(ast.args[2], state)) : m[0].length;
    const rows = r >= 0 ? m.slice(0, r) : m.slice(r); return out(rows.map((row) => (c >= 0 ? row.slice(0, c) : row.slice(c))));
  }
  drop(ast, state) {
    const m = rows2d(this.val(ast.args[0], state)); const r = ast.args[1] ? num(this.val(ast.args[1], state)) : 0; const c = ast.args[2] ? num(this.val(ast.args[2], state)) : 0;
    const rows = r >= 0 ? m.slice(r) : m.slice(0, r); const res = rows.map((row) => (c >= 0 ? row.slice(c) : row.slice(0, c)));
    return res.length && res[0].length ? out(res) : new CellError(ErrorType.VALUE, 'Matriz vacía');
  }
  takeSize(ast, state) { const s = this.size(ast.args[0], state); const r = ast.args[1] ? this.lit(ast.args[1]) : s.height; const c = ast.args[2] ? this.lit(ast.args[2]) : s.width; if (r == null || c == null) return ArraySize.error(); return new ArraySize(Math.min(Math.abs(c), s.width), Math.min(Math.abs(r), s.height)); }
  dropSize(ast, state) { const s = this.size(ast.args[0], state); const r = ast.args[1] ? this.lit(ast.args[1]) : 0; const c = ast.args[2] ? this.lit(ast.args[2]) : 0; if (r == null || c == null) return ArraySize.error(); return new ArraySize(Math.max(1, s.width - Math.abs(c)), Math.max(1, s.height - Math.abs(r))); }
  tocol(ast, state) { return out(rows2d(this.val(ast.args[0], state)).flat().map((v) => [v])); }
  tocolSize(ast, state) { const s = this.size(ast.args[0], state); return new ArraySize(1, s.width * s.height); }
  torow(ast, state) { return out([rows2d(this.val(ast.args[0], state)).flat()]); }
  torowSize(ast, state) { const s = this.size(ast.args[0], state); return new ArraySize(s.width * s.height, 1); }
  choosecols(ast, state) {
    const m = rows2d(this.val(ast.args[0], state)); const idx = ast.args.slice(1).map((a) => num(this.val(a, state)));
    if (idx.some((i) => !i || Math.abs(i) > m[0].length)) return ERR('Columna fuera del rango');
    return out(m.map((row) => idx.map((i) => row[i > 0 ? i - 1 : row.length + i])));
  }
  choosecolsSize(ast, state) { const s = this.size(ast.args[0], state); return new ArraySize(ast.args.length - 1, s.height); }
  chooserows(ast, state) {
    const m = rows2d(this.val(ast.args[0], state)); const idx = ast.args.slice(1).map((a) => num(this.val(a, state)));
    if (idx.some((i) => !i || Math.abs(i) > m.length)) return ERR('Fila fuera del rango');
    return out(idx.map((i) => m[i > 0 ? i - 1 : m.length + i]));
  }
  chooserowsSize(ast, state) { const s = this.size(ast.args[0], state); return new ArraySize(s.width, ast.args.length - 1); }
  wraprows(ast, state) {
    const flat = rows2d(this.val(ast.args[0], state)).flat(); const n = num(this.val(ast.args[1], state));
    if (!(n >= 1)) return ERR('WRAPROWS: tamaño no válido');
    const pad = ast.args[2] ? this.val(ast.args[2], state) : new CellError(ErrorType.NA);
    const res = []; for (let i = 0; i < flat.length; i += n) { const row = flat.slice(i, i + n); while (row.length < n) row.push(pad); res.push(row); }
    return out(res);
  }
  wraprowsSize(ast, state) { const s = this.size(ast.args[0], state); const n = this.lit(ast.args[1]); return n >= 1 ? new ArraySize(n, Math.ceil((s.width * s.height) / n)) : ArraySize.error(); }

  // ---- texto
  textbefore(ast, state) { return this.textPart(ast, state, true); }
  textafter(ast, state) { return this.textPart(ast, state, false); }
  textPart(ast, state, before) {
    const text = String(plain(this.val(ast.args[0], state)) ?? ''); const del = String(plain(this.val(ast.args[1], state)) ?? '');
    let n = ast.args[2] ? num(this.val(ast.args[2], state)) : 1; if (!n) n = 1;
    if (!del) return before ? '' : text;
    let pos = -1;
    if (n > 0) { let from = 0; for (let k = 0; k < n; k++) { pos = text.indexOf(del, from); if (pos < 0) break; from = pos + del.length; } } else { let from = text.length; for (let k = 0; k < -n; k++) { pos = text.lastIndexOf(del, from); if (pos < 0) break; from = pos - 1; } }
    if (pos < 0) return new CellError(ErrorType.NA, 'No se encontró el delimitador');
    return before ? text.slice(0, pos) : text.slice(pos + del.length);
  }

  // ---- IMAGEN: devuelve la dirección; la cuadrícula dibuja la imagen en la celda
  image(ast, state) {
    const src = String(plain(this.val(ast.args[0], state)) ?? '').trim();
    if (!/^(https?:|data:image\/)/i.test(src)) return ERR('IMAGEN necesita una dirección https:// o data:image');
    return src;
  }
}
// ---- estadística que faltaba en HyperFormula (como en LibreOffice Calc y Excel)
const nums = (v) => rows2d(v).flat().map(plain).filter((x) => typeof x === 'number');
const pairs = (a, b) => { const A = rows2d(a).flat().map(plain); const B = rows2d(b).flat().map(plain); const o = []; for (let i = 0; i < Math.min(A.length, B.length); i++) if (typeof A[i] === 'number' && typeof B[i] === 'number') o.push([A[i], B[i]]); return o; };
const NUMERR = () => new CellError(ErrorType.NUM);
const NAERR = () => new CellError(ErrorType.NA);
const DIV0 = () => new CellError(ErrorType.DIV_BY_ZERO);
Object.assign(OficinaPlugin.prototype, {
  kurt(ast, state) {
    const x = ast.args.flatMap((a) => nums(this.val(a, state))); const n = x.length; if (n < 4) return DIV0();
    const m = x.reduce((a, b) => a + b, 0) / n; const sd = Math.sqrt(x.reduce((a, b) => a + (b - m) ** 2, 0) / (n - 1)); if (!sd) return DIV0();
    const s4 = x.reduce((a, b) => a + ((b - m) / sd) ** 4, 0);
    return (n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3)) * s4 - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
  },
  mode(ast, state) {
    const x = ast.args.flatMap((a) => nums(this.val(a, state))); const cnt = new Map(); let best = null; let bc = 1;
    for (const v of x) { const c = (cnt.get(v) || 0) + 1; cnt.set(v, c); if (c > bc) { bc = c; best = v; } }
    return best == null ? NAERR() : best;
  },
  intercept(ast, state) {
    const p = pairs(this.val(ast.args[0], state), this.val(ast.args[1], state)); if (p.length < 2) return DIV0();
    const my = p.reduce((a, q) => a + q[0], 0) / p.length; const mx = p.reduce((a, q) => a + q[1], 0) / p.length;
    const sxx = p.reduce((a, q) => a + (q[1] - mx) ** 2, 0); if (!sxx) return DIV0();
    const b = p.reduce((a, q) => a + (q[1] - mx) * (q[0] - my), 0) / sxx; return my - b * mx;
  },
  forecast(ast, state) {
    const x0 = num(plain(this.val(ast.args[0], state))); if (Number.isNaN(x0)) return ERR('PRONOSTICO: valor no numérico');
    const p = pairs(this.val(ast.args[1], state), this.val(ast.args[2], state)); if (p.length < 2) return DIV0();
    const my = p.reduce((a, q) => a + q[0], 0) / p.length; const mx = p.reduce((a, q) => a + q[1], 0) / p.length;
    const sxx = p.reduce((a, q) => a + (q[1] - mx) ** 2, 0); if (!sxx) return DIV0();
    const b = p.reduce((a, q) => a + (q[1] - mx) * (q[0] - my), 0) / sxx; return my + b * (x0 - mx);
  },
});
void NUMERR;
const LAZY = { doesNotNeedArgumentsToBeComputed: true };
const any = (n) => Array.from({ length: n }, () => ({ argumentType: FunctionArgumentType.ANY }));
OficinaPlugin.implementedFunctions = {
  LAMBDA: { method: 'lambda', ...LAZY, parameters: any(1), repeatLastArgs: 1 },
  LET: { method: 'let_', ...LAZY, parameters: any(3), repeatLastArgs: 1, sizeOfResultArrayMethod: 'letSize' },
  MAP: { method: 'map', ...LAZY, parameters: any(2), sizeOfResultArrayMethod: 'mapSize' },
  REDUCE: { method: 'reduce', ...LAZY, parameters: any(3) },
  SCAN: { method: 'scan', ...LAZY, parameters: any(3), sizeOfResultArrayMethod: 'scanSize' },
  BYROW: { method: 'byrow', ...LAZY, parameters: any(2), sizeOfResultArrayMethod: 'byrowSize' },
  BYCOL: { method: 'bycol', ...LAZY, parameters: any(2), sizeOfResultArrayMethod: 'bycolSize' },
  MAKEARRAY: { method: 'makearray', ...LAZY, parameters: any(3), sizeOfResultArrayMethod: 'makearraySize' },
  TAKE: { method: 'take', ...LAZY, parameters: any(3), sizeOfResultArrayMethod: 'takeSize' },
  DROP: { method: 'drop', ...LAZY, parameters: any(3), sizeOfResultArrayMethod: 'dropSize' },
  TOCOL: { method: 'tocol', ...LAZY, parameters: any(1), sizeOfResultArrayMethod: 'tocolSize' },
  TOROW: { method: 'torow', ...LAZY, parameters: any(1), sizeOfResultArrayMethod: 'torowSize' },
  CHOOSECOLS: { method: 'choosecols', ...LAZY, parameters: any(2), repeatLastArgs: 1, sizeOfResultArrayMethod: 'choosecolsSize' },
  CHOOSEROWS: { method: 'chooserows', ...LAZY, parameters: any(2), repeatLastArgs: 1, sizeOfResultArrayMethod: 'chooserowsSize' },
  WRAPROWS: { method: 'wraprows', ...LAZY, parameters: any(3), sizeOfResultArrayMethod: 'wraprowsSize' },
  TEXTBEFORE: { method: 'textbefore', ...LAZY, parameters: any(3) },
  TEXTAFTER: { method: 'textafter', ...LAZY, parameters: any(3) },
  IMAGE: { method: 'image', ...LAZY, parameters: any(1), repeatLastArgs: 1 },
  KURT: { method: 'kurt', ...LAZY, parameters: any(1), repeatLastArgs: 1 },
  MODE: { method: 'mode', ...LAZY, parameters: any(1), repeatLastArgs: 1 },
  'MODE.SNGL': { method: 'mode', ...LAZY, parameters: any(1), repeatLastArgs: 1 },
  INTERCEPT: { method: 'intercept', ...LAZY, parameters: any(2) },
  FORECAST: { method: 'forecast', ...LAZY, parameters: any(3) },
  'FORECAST.LINEAR': { method: 'forecast', ...LAZY, parameters: any(3) },
};
const names = Object.keys(OficinaPlugin.implementedFunctions);
HyperFormula.registerFunctionPlugin(OficinaPlugin, { enGB: Object.fromEntries(names.map((n) => [n, n])) });

// Nombres en español (como en Excel en español)
export const ES_NAMES = {
  VSTACK: 'APILARV', HSTACK: 'APILARH', TAKE: 'TOMAR', DROP: 'EXCLUIR', TOCOL: 'ENCOL', TOROW: 'ENFILA', CHOOSECOLS: 'ELEGIRCOLS', CHOOSEROWS: 'ELEGIRFILAS',
  WRAPROWS: 'AJUSTARFILAS', TEXTBEFORE: 'TEXTOANTES', TEXTAFTER: 'TEXTODESPUES', IMAGE: 'IMAGEN',
  KURT: 'CURTOSIS', MODE: 'MODA', 'MODE.SNGL': 'MODA.UNO', INTERCEPT: 'INTERSECCION.EJE', FORECAST: 'PRONOSTICO', 'FORECAST.LINEAR': 'PRONOSTICO.LINEAL',
  LAMBDA: 'LAMBDA', LET: 'LET', MAP: 'MAP', REDUCE: 'REDUCE', SCAN: 'SCAN', BYROW: 'BYROW', BYCOL: 'BYCOL', MAKEARRAY: 'MAKEARRAY',
};
