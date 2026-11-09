// Modelo del libro de Calcula Datos: valores y fórmulas en HyperFormula (GPLv3), formato aparte. GPL-3.0
import { HyperFormula } from 'hyperformula';
import { toEnglish, toSpanish, errorText } from './i18n.js';
import { parseInput, formatValue } from './format.js';

export const DEF_W = 92, DEF_H = 24;
export const colName = (c) => { let s = ''; c += 1; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
export const colIndex = (s) => s.toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
export const addr = (r, c) => `${colName(c)}${r + 1}`;
export const parseAddr = (a) => { const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a.trim()); return m ? { r: Number(m[2]) - 1, c: colIndex(m[1]) } : null; };
export const parseRange = (s) => {
  const [a, b] = s.split(':');
  const p = parseAddr(a); const q = b ? parseAddr(b) : p;
  if (!p || !q) return null;
  return { r1: Math.min(p.r, q.r), c1: Math.min(p.c, q.c), r2: Math.max(p.r, q.r), c2: Math.max(p.c, q.c) };
};
export const rangeStr = (g) => (g.r1 === g.r2 && g.c1 === g.c2 ? addr(g.r1, g.c1) : `${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)}`);

const key = (r, c) => `${r},${c}`;
const sameMeta = (a, b) => a.styles.size === b.styles.size && [...a.styles].every(([k, v]) => b.styles.get(k) === v)
  && JSON.stringify([[...a.colW], [...a.rowH], [...a.hiddenRows], a.merges, a.freeze, a.charts, a.nRows === b.nRows, a.name]) === JSON.stringify([[...b.colW], [...b.rowH], [...b.hiddenRows], b.merges, b.freeze, b.charts, true, b.name])
  && !!a.filter === !!b.filter;

export class Sheet {
  constructor(id, name) {
    this.id = id; this.name = name;
    this.styles = new Map(); // "r,c" -> objeto de estilo (inmutable)
    this.colW = new Map(); this.rowH = new Map();
    this.hiddenRows = new Set();
    this.merges = []; // {r1,c1,r2,c2}
    this.freeze = { rows: 0, cols: 0 };
    this.filter = null; // { range, col criteria: Map col -> Set de valores ocultos }
    this.charts = []; // { id, type, range, title, x, y, w, h }
    this.nRows = 1000; this.nCols = 52;
  }
  cw(c) { return this.colW.get(c) ?? DEF_W; }
  rh(r) { return this.hiddenRows.has(r) ? 0 : (this.rowH.get(r) ?? DEF_H); }
  style(r, c) { return this.styles.get(key(r, c)) || null; }
  meta() {
    return {
      styles: new Map(this.styles), colW: new Map(this.colW), rowH: new Map(this.rowH), hiddenRows: new Set(this.hiddenRows),
      merges: this.merges.map((m) => ({ ...m })), freeze: { ...this.freeze }, filter: this.filter ? { range: { ...this.filter.range }, hidden: new Map([...this.filter.hidden].map(([k, v]) => [k, new Set(v)])) } : null,
      charts: this.charts.map((c) => ({ ...c })), nRows: this.nRows, nCols: this.nCols, name: this.name,
    };
  }
  setMeta(m) { Object.assign(this, m); }
  mergeAt(r, c) { return this.merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) || null; }
}

export class Workbook {
  constructor() {
    this.hf = HyperFormula.buildEmpty({
      licenseKey: 'gpl-v3', language: 'enGB', functionArgSeparator: ',', decimalSeparator: '.', thousandSeparator: '',
      dateFormats: ['DD/MM/YYYY', 'DD/MM/YY'], timeFormats: ['hh:mm', 'hh:mm:ss.sss'], useArrayArithmetic: true,
      precisionRounding: 10, smartRounding: true, undoLimit: 200, maxRows: 1048576, maxColumns: 16384,
      nullDate: { year: 1899, month: 12, day: 30 }, leapYear1900: true,
    });
    this.sheets = [];
    this.active = 0;
    this.undoStack = []; this.redoStack = [];
    this.listeners = new Set();
  }
  get sheet() { return this.sheets[this.active]; }
  get sid() { return this.sheet.id; }
  onChange(fn) { this.listeners.add(fn); }
  emit(kind) { this.listeners.forEach((fn) => fn(kind)); }

  reset() {
    for (const s of this.sheets) this.hf.removeSheet(s.id);
    this.sheets = []; this.active = 0; this.undoStack = []; this.redoStack = [];
    this.hf.clearUndoStack?.(); this.hf.clearRedoStack?.();
  }
  addSheet(name) {
    let n = name || `Hoja${this.sheets.length + 1}`;
    let i = 1; while (this.sheets.some((s) => s.name.toLowerCase() === n.toLowerCase())) n = `${name || 'Hoja'}${++i}`;
    const added = this.hf.addSheet(n);
    const s = new Sheet(this.hf.getSheetId(added), added);
    this.sheets.push(s);
    return s;
  }

  // ---------------- deshacer: cada acción guarda el formato antes/después y si cambió HyperFormula
  action(label, fn, { hf = true, sheets, batch = false } = {}) {
    const targets = sheets || [this.sheet];
    const before = targets.map((s) => [s, s.meta()]);
    // Se cuentan las entradas que agrega HyperFormula a su pila de deshacer para revertirlas juntas.
    const ur = this.hf._crudOperations?.undoRedo;
    const n0 = ur ? ur.undoStack.length : 0;
    let err = null;
    try { if (hf && batch) this.hf.batch(() => fn()); else fn(); } catch (e) { err = e; }
    let hfSteps = 0;
    if (hf && ur) {
      hfSteps = Math.max(0, ur.undoStack.length - n0);
      const last = ur.undoStack[ur.undoStack.length - 1];
      if (hfSteps === 1 && last?.operations && last.operations.length === 0) { ur.undoStack.pop(); hfSteps = 0; }
    }
    const after = targets.map((s) => [s, s.meta()]);
    if (hfSteps || before.some(([s, m], i) => !sameMeta(m, after[i][1]))) {
      this.undoStack.push({ label, hfSteps, before, after, active: this.active });
      if (this.undoStack.length > 200) this.undoStack.shift();
      this.redoStack = [];
    }
    this.emit('edit');
    if (err) throw err;
  }
  undo() {
    const a = this.undoStack.pop(); if (!a) return false;
    for (let i = 0; i < a.hfSteps; i++) if (this.hf.isThereSomethingToUndo()) this.hf.undo();
    a.before.forEach(([s, m]) => s.setMeta(m));
    this.redoStack.push(a); this.emit('edit'); return true;
  }
  redo() {
    const a = this.redoStack.pop(); if (!a) return false;
    for (let i = 0; i < a.hfSteps; i++) if (this.hf.isThereSomethingToRedo()) this.hf.redo();
    a.after.forEach(([s, m]) => s.setMeta(m));
    this.undoStack.push(a); this.emit('edit'); return true;
  }

  // ---------------- valores
  raw(r, c, s = this.sheet) { return this.hf.getCellSerialized({ sheet: s.id, row: r, col: c }); }
  value(r, c, s = this.sheet) {
    const v = this.hf.getCellValue({ sheet: s.id, row: r, col: c });
    return v;
  }
  isError(v) { return v && typeof v === 'object' && 'type' in v && 'value' in v; }
  display(r, c, s = this.sheet) {
    const v = this.value(r, c, s);
    const st = s.style(r, c);
    if (this.isError(v)) return { text: errorText(v), align: 'center', error: true };
    return formatValue(v, st?.fmt);
  }
  // Lo que se muestra en la barra de fórmulas
  editText(r, c, s = this.sheet) {
    const raw = this.raw(r, c, s);
    if (raw == null) return '';
    if (typeof raw === 'string' && raw.startsWith('=')) return toSpanish(raw);
    const st = s.style(r, c);
    if (typeof raw === 'number' && st?.fmt && st.fmt !== 'General' && st.fmt !== '@') {
      if (/%/.test(st.fmt)) return `${Math.round(raw * 100 * 1e8) / 1e8}%`;
      const d = formatValue(raw, st.fmt);
      if (/[dmy]/i.test(st.fmt.replace(/"[^"]*"/g, '')) && !/[0#]/.test(st.fmt)) return d.text;
    }
    if (typeof raw === 'boolean') return raw ? 'VERDADERO' : 'FALSO';
    return String(raw);
  }
  setInput(r, c, text) {
    const s = this.sheet;
    const p = parseInput(text);
    let v = p.value;
    if (typeof v === 'string' && v.startsWith('=')) v = toEnglish(v);
    this.hf.setCellContents({ sheet: s.id, row: r, col: c }, [[p.text && typeof v === 'string' ? v : v]]);
    if (p.fmt) { const st = s.style(r, c); if (!st?.fmt || st.fmt === 'General') this.setStyle(r, c, { fmt: p.fmt }); }
    this.grow(r, c);
  }
  setRaw(r, c, v, s = this.sheet) { this.hf.setCellContents({ sheet: s.id, row: r, col: c }, [[v]]); this.grow(r, c, s); }
  grow(r, c, s = this.sheet) { if (r >= s.nRows - 5) s.nRows = r + 200; if (c >= s.nCols - 2) s.nCols = c + 20; }

  // ---------------- formato
  setStyle(r, c, patch, s = this.sheet) {
    const k = key(r, c);
    const cur = s.styles.get(k) || {};
    const next = { ...cur, ...patch };
    for (const p of Object.keys(next)) if (next[p] === undefined || next[p] === null || next[p] === false) delete next[p];
    if (Object.keys(next).length) s.styles.set(k, Object.freeze(next)); else s.styles.delete(k);
  }
  styleRange(g, patch) { for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) this.setStyle(r, c, patch); }

  // ---------------- uso de datos
  usedRange(s = this.sheet) {
    const dims = this.hf.getSheetDimensions(s.id);
    let r2 = dims.height - 1, c2 = dims.width - 1;
    for (const k of s.styles.keys()) { const [r, c] = k.split(',').map(Number); if (r > r2) r2 = r; if (c > c2) c2 = c; }
    return { r1: 0, c1: 0, r2: Math.max(0, r2), c2: Math.max(0, c2), empty: dims.height === 0 && s.styles.size === 0 };
  }
  // Región contigua con datos alrededor de una celda (como Ctrl+A / ordenar)
  region(r, c) {
    const filled = (rr, cc) => rr >= 0 && cc >= 0 && this.raw(rr, cc) != null && this.raw(rr, cc) !== '';
    let g = { r1: r, c1: c, r2: r, c2: c };
    let changed = true;
    while (changed) {
      changed = false;
      const tryExp = (dr1, dc1, dr2, dc2) => {
        const n = { r1: g.r1 + dr1, c1: g.c1 + dc1, r2: g.r2 + dr2, c2: g.c2 + dc2 };
        if (n.r1 < 0 || n.c1 < 0) return;
        let any = false;
        if (dr1) for (let cc = n.c1; cc <= n.c2; cc++) any = any || filled(n.r1, cc);
        if (dr2) for (let cc = n.c1; cc <= n.c2; cc++) any = any || filled(n.r2, cc);
        if (dc1) for (let rr = n.r1; rr <= n.r2; rr++) any = any || filled(rr, n.c1);
        if (dc2) for (let rr = n.r1; rr <= n.r2; rr++) any = any || filled(rr, n.c2);
        if (any) { g = n; changed = true; }
      };
      tryExp(-1, 0, 0, 0); tryExp(0, 0, 1, 0); tryExp(0, -1, 0, 0); tryExp(0, 0, 0, 1);
    }
    return g;
  }

  // ---------------- filas y columnas
  shiftMeta(s, axis, at, count) {
    // Mueve formato, tamaños y combinaciones cuando se insertan (count>0) o eliminan (count<0) filas/columnas.
    const moveKey = (k) => { const [r, c] = k.split(',').map(Number); const v = axis === 'row' ? r : c; if (count < 0 && v >= at && v < at - count) return null; if (v >= at) return axis === 'row' ? key(r + count, c) : key(r, c + count); return k; };
    const ns = new Map();
    for (const [k, v] of s.styles) { const nk = moveKey(k); if (nk) ns.set(nk, v); }
    s.styles = ns;
    const sizes = axis === 'row' ? s.rowH : s.colW;
    const nsz = new Map();
    for (const [i, v] of sizes) { if (count < 0 && i >= at && i < at - count) continue; nsz.set(i >= at ? i + count : i, v); }
    if (axis === 'row') s.rowH = nsz; else s.colW = nsz;
    if (axis === 'row') { const nh = new Set(); for (const i of s.hiddenRows) { if (count < 0 && i >= at && i < at - count) continue; nh.add(i >= at ? i + count : i); } s.hiddenRows = nh; }
    s.merges = s.merges.map((m) => {
      const [a, b] = axis === 'row' ? ['r1', 'r2'] : ['c1', 'c2'];
      const n = { ...m };
      if (n[a] >= at) n[a] = Math.max(at, n[a] + count);
      if (n[b] >= at) n[b] = Math.max(at - 1, n[b] + count);
      return n[b] >= n[a] ? n : null;
    }).filter(Boolean);
  }
  insertRows(at, n = 1) { const s = this.sheet; this.hf.addRows(s.id, [at, n]); this.shiftMeta(s, 'row', at, n); s.nRows += n; }
  deleteRows(at, n = 1) { const s = this.sheet; this.hf.removeRows(s.id, [at, n]); this.shiftMeta(s, 'row', at, -n); }
  insertCols(at, n = 1) { const s = this.sheet; this.hf.addColumns(s.id, [at, n]); this.shiftMeta(s, 'col', at, n); s.nCols += n; }
  deleteCols(at, n = 1) { const s = this.sheet; this.hf.removeColumns(s.id, [at, n]); this.shiftMeta(s, 'col', at, -n); }

  // ---------------- ordenar
  sortRange(g, col, asc = true, hasHeader = false) {
    const s = this.sheet;
    const start = hasHeader ? g.r1 + 1 : g.r1;
    const rows = [];
    for (let r = start; r <= g.r2; r++) {
      const cells = []; const styles = [];
      for (let c = g.c1; c <= g.c2; c++) { cells.push(this.raw(r, c)); styles.push(s.style(r, c)); }
      rows.push({ cells, styles, v: this.value(r, col) });
    }
    const rank = (v) => (v == null || v === '' ? 3 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
    rows.sort((a, b) => {
      const ra = rank(a.v), rb = rank(b.v);
      if (ra !== rb) return ra - rb; // vacíos siempre al final
      let d = 0;
      if (ra === 0) d = a.v - b.v; else if (ra === 1) d = a.v.localeCompare(b.v, 'es', { sensitivity: 'base', numeric: true });
      return asc ? d : -d;
    });
    // Las fórmulas se escriben como valores ya calculados si hacen referencia relativa a otra fila (como hacen otras hojas de cálculo, se recalculan).
    rows.forEach((row, i) => {
      const r = start + i;
      row.cells.forEach((v, j) => {
        this.hf.setCellContents({ sheet: s.id, row: r, col: g.c1 + j }, [[v]]);
        const st = row.styles[j];
        if (st) s.styles.set(key(r, g.c1 + j), st); else s.styles.delete(key(r, g.c1 + j));
      });
    });
  }

  // ---------------- serializar para recuperación
  toJSON() {
    return {
      v: 1, active: this.active,
      sheets: this.sheets.map((s) => ({
        name: s.name, data: this.hf.getSheetSerialized(s.id),
        styles: [...s.styles], colW: [...s.colW], rowH: [...s.rowH], hiddenRows: [...s.hiddenRows], merges: s.merges, freeze: s.freeze,
        charts: s.charts, filter: s.filter ? { range: s.filter.range, hidden: [...s.filter.hidden].map(([k, v]) => [k, [...v]]) } : null,
      })),
    };
  }
  fromJSON(j) {
    this.reset();
    for (const d of j.sheets) {
      const s = this.addSheet(d.name);
      this.hf.setSheetContent(s.id, d.data || []);
      s.styles = new Map(d.styles); s.colW = new Map(d.colW); s.rowH = new Map(d.rowH); s.hiddenRows = new Set(d.hiddenRows);
      s.merges = d.merges || []; s.freeze = d.freeze || { rows: 0, cols: 0 }; s.charts = d.charts || [];
      s.filter = d.filter ? { range: d.filter.range, hidden: new Map(d.filter.hidden.map(([k, v]) => [k, new Set(v)])) } : null;
      const u = this.usedRange(s); s.nRows = Math.max(1000, u.r2 + 200); s.nCols = Math.max(52, u.c2 + 20);
    }
    this.active = Math.min(j.active || 0, this.sheets.length - 1);
    this.hf.clearUndoStack?.(); this.undoStack = []; this.redoStack = [];
  }
}
