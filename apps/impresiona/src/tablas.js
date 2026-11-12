// Impresiona — tablas como en LibreOffice Impress (Diseño de tabla): insertar y eliminar filas y
// columnas, combinar y dividir celdas, estilos de tabla con fila de encabezado, filas con bandas,
// primera columna y última fila, y relleno, color, negrita y alineación de cada celda. GPL-3.0
import { h, t, toast } from '../../shared/office.js';
import { tabbed, field, numIn } from '../../shared/dialogs.js';

let C = null; // { slide, commit, selObjs, stopEditing, theme }
export function init(ctx) { C = ctx; }
let curCell = { id: null, r: 0, c: 0 };
export const setCurrent = (id, r, c) => { curCell = { id, r, c }; };
export const current = () => curCell;

// estilos (los colores del tema se usan con «Tema»)
export const PRESETS = {
  tema: null,
  sinestilo: { head: 'transparent', headText: null, band: 'transparent', border: '#9AA0B4' },
  gris: { head: '#595959', headText: '#FFFFFF', band: '#EDEDED', border: '#BFBFBF' },
  azul: { head: '#2F5597', headText: '#FFFFFF', band: '#DAE3F3', border: '#8FAADC' },
  verde: { head: '#548235', headText: '#FFFFFF', band: '#E2F0D9', border: '#A9D18E' },
  naranja: { head: '#C55A11', headText: '#FFFFFF', band: '#FBE5D6', border: '#F4B183' },
  morado: { head: '#7030A0', headText: '#FFFFFF', band: '#E9DCF3', border: '#C09BDB' },
};
export const PRESET_NAMES = [['tema', 'Tema'], ['sinestilo', 'Sin estilo'], ['gris', 'Gris'], ['azul', 'Azul'], ['verde', 'Verde'], ['naranja', 'Naranja'], ['morado', 'Morado']];
export const opts = (o) => ({ preset: 'tema', header: true, banded: false, firstCol: false, lastRow: false, ...(o.tbl || {}) });
const tint = (hex, k) => { const n = (i) => Math.round(parseInt(hex.slice(i, i + 2), 16) * (1 - k) + 255 * k); return /^#[0-9a-f]{6}$/i.test(hex || '') ? `#${[1, 3, 5].map((i) => n(i).toString(16).padStart(2, '0')).join('')}` : '#EEEEEE'; };
// aspecto final de una celda (estilo de la tabla + formato propio de la celda)
export function cellLook(o, r, c, th) {
  const T = opts(o); const R = o.rows.length; const cell = o.rows[r][c] || {};
  const P = PRESETS[T.preset] || { head: th?.accent || '#DC3A63', headText: '#FFFFFF', band: tint(th?.accent || '#DC3A63', 0.82), border: '#9AA0B4' };
  let bg = ''; let color = ''; let bold = false; let borderTop = '';
  const isHead = T.header && r === 0;
  if (isHead) { bg = P.head; color = P.headText || ''; bold = true; }
  else if (T.banded && ((r - (T.header ? 1 : 0)) % 2 === 1)) bg = P.band;
  if (T.firstCol && c === 0 && !isHead) bold = true;
  if (T.lastRow && r === R - 1 && R > 1) { bold = true; borderTop = `2px solid ${P.head && P.head !== 'transparent' ? P.head : '#555'}`; }
  if (bg === 'transparent') bg = '';
  return { bg: cell.bg || bg, color: cell.color || color, bold: cell.bold != null ? !!cell.bold : bold, align: cell.align || '', border: P.border, borderTop };
}

// ---------------------------------------------------------------- modelo de celdas combinadas
function anchors(rows) {
  const out = [];
  rows.forEach((row, r) => row.forEach((cell, c) => { if (!cell.hidden) out.push({ r, c, rs: cell.rs || 1, cs: cell.cs || 1, cell }); }));
  return out;
}
function rebuild(list, R, Cn) {
  const rows = Array.from({ length: R }, () => Array.from({ length: Cn }, () => ({ text: '', hidden: true, rs: 1, cs: 1 })));
  for (const a of list) {
    const rs = Math.max(1, Math.min(a.rs, R - a.r)); const cs = Math.max(1, Math.min(a.cs, Cn - a.c));
    rows[a.r][a.c] = { ...a.cell, hidden: false, rs, cs };
  }
  // las celdas que nadie cubre quedan visibles
  const cover = rows.map((row) => row.map(() => false));
  rows.forEach((row, r) => row.forEach((cell, c) => { if (!cell.hidden) for (let i = r; i < r + cell.rs; i++) for (let j = c; j < c + cell.cs; j++) if (i !== r || j !== c) cover[i][j] = true; }));
  rows.forEach((row, r) => row.forEach((cell, c) => { if (cell.hidden && !cover[r][c]) row[c] = { text: '', hidden: false, rs: 1, cs: 1 }; }));
  return rows;
}
const owner = (list, r, c) => list.find((a) => r >= a.r && r < a.r + a.rs && c >= a.c && c < a.c + a.cs);
const fresh = () => ({ text: '', rs: 1, cs: 1 });

function tableObj() { const o = C.selObjs().find((x) => x.type === 'table'); if (!o) toast('Selecciona una tabla.', { error: true }); return o; }
function edit(label, fn) {
  C.stopEditing();
  const sel = C.selObjs().find((x) => x.type === 'table'); if (!sel) { toast('Selecciona una tabla.', { error: true }); return null; }
  const id = sel.id; const pos = curCell.id === id ? curCell : { r: 0, c: 0 };
  let res = null;
  C.commit(label, () => { const o = C.slide().objects.find((x) => x.id === id); res = fn(o, Math.min(pos.r, o.rows.length - 1), Math.min(pos.c, o.rows[0].length - 1)); });
  return res;
}

export function insertRow(below = true) {
  return edit(t('Insertar fila'), (o, r) => {
    const R = o.rows.length; const Cn = o.rows[0].length; const list = anchors(o.rows);
    const a0 = owner(list, r, 0); const at = below ? (a0 ? a0.r + a0.rs : r + 1) : (a0 ? a0.r : r);
    for (const a of list) { if (a.r >= at) a.r += 1; else if (a.r + a.rs > at) a.rs += 1; }
    for (let c = 0; c < Cn; c++) if (!owner(list, at, c)) list.push({ r: at, c, rs: 1, cs: 1, cell: fresh() });
    o.rows = rebuild(list, R + 1, Cn); o.h += Math.round(o.h / R); curCell.r = at; return at;
  });
}
export function insertCol(right = true) {
  return edit(t('Insertar columna'), (o, r, c) => {
    const R = o.rows.length; const Cn = o.rows[0].length; const list = anchors(o.rows);
    const a0 = owner(list, r, c); const at = right ? (a0 ? a0.c + a0.cs : c + 1) : (a0 ? a0.c : c);
    for (const a of list) { if (a.c >= at) a.c += 1; else if (a.c + a.cs > at) a.cs += 1; }
    for (let rr = 0; rr < R; rr++) if (!owner(list, rr, at)) list.push({ r: rr, c: at, rs: 1, cs: 1, cell: fresh() });
    if (o.colW) o.colW.splice(at, 0, o.colW[Math.max(0, at - 1)] || 1);
    o.rows = rebuild(list, R, Cn + 1); o.w += Math.round(o.w / Cn); curCell.c = at; return at;
  });
}
export function deleteRow() {
  return edit(t('Eliminar fila'), (o, r) => {
    const R = o.rows.length; const Cn = o.rows[0].length; if (R <= 1) return false;
    let list = anchors(o.rows);
    list = list.filter((a) => !(a.r === r && a.rs === 1));
    for (const a of list) { if (a.r > r) a.r -= 1; else if (a.r + a.rs > r) a.rs -= 1; }
    o.rows = rebuild(list, R - 1, Cn); o.h = Math.max(30, o.h - Math.round(o.h / R)); curCell.r = Math.min(r, R - 2); return true;
  });
}
export function deleteCol() {
  return edit(t('Eliminar columna'), (o, r, c) => {
    const R = o.rows.length; const Cn = o.rows[0].length; if (Cn <= 1) return false;
    let list = anchors(o.rows);
    list = list.filter((a) => !(a.c === c && a.cs === 1));
    for (const a of list) { if (a.c > c) a.c -= 1; else if (a.c + a.cs > c) a.cs -= 1; }
    if (o.colW) o.colW.splice(c, 1);
    o.rows = rebuild(list, R, Cn - 1); o.w = Math.max(40, o.w - Math.round(o.w / Cn)); curCell.c = Math.min(c, Cn - 2); return true;
  });
}
export async function mergeDialog() {
  const o = tableObj(); if (!o) return null;
  const body = await tabbed('Combinar celdas', [{ label: 'Combinar', body: h('div', { class: 'fx-grid' },
    h('p', { class: 'fx-note', style: { gridColumn: '1 / -1' } }, t('Se combinan desde la celda actual hacia la derecha y hacia abajo.')),
    field('Columnas', numIn('cols', 2, { step: 1, min: 1, max: 50 })), field('Filas', numIn('rows', 1, { step: 1, min: 1, max: 50 }))) }], { width: '400px' });
  if (!body) return null;
  return mergeCells(Number(body.querySelector('[name=rows]').value) || 1, Number(body.querySelector('[name=cols]').value) || 1);
}
export function mergeCells(nr, nc) {
  return edit(t('Combinar celdas'), (o, r, c) => {
    const R = o.rows.length; const Cn = o.rows[0].length; let list = anchors(o.rows);
    const a0 = owner(list, r, c); r = a0.r; c = a0.c;
    const r2 = Math.min(R - 1, r + nr - 1); const c2 = Math.min(Cn - 1, c + nc - 1);
    const inside = list.filter((a) => a.r >= r && a.c >= c && a.r <= r2 && a.c <= c2);
    const text = inside.map((a) => a.cell.text).filter(Boolean).join(' ');
    list = list.filter((a) => !inside.includes(a) || a === a0);
    a0.rs = r2 - r + 1; a0.cs = c2 - c + 1; a0.cell = { ...a0.cell, text };
    o.rows = rebuild(list, R, Cn); return [a0.rs, a0.cs];
  });
}
export function splitCell() {
  return edit(t('Dividir celda'), (o, r, c) => {
    const R = o.rows.length; const Cn = o.rows[0].length; const list = anchors(o.rows);
    const a = owner(list, r, c); if (!a || (a.rs === 1 && a.cs === 1)) return false;
    for (let i = a.r; i < a.r + a.rs; i++) for (let j = a.c; j < a.c + a.cs; j++) if (i !== a.r || j !== a.c) list.push({ r: i, c: j, rs: 1, cs: 1, cell: fresh() });
    a.rs = 1; a.cs = 1;
    o.rows = rebuild(list, R, Cn); return true;
  });
}
export function setOption(key, value) {
  return edit(t('Estilo de tabla'), (o) => { o.tbl = { ...opts(o), [key]: value === undefined ? !opts(o)[key] : value }; return o.tbl[key]; });
}
export function cellFormat(patch) {
  return edit(t('Formato de celda'), (o, r, c) => { const list = anchors(o.rows); const a = owner(list, r, c); Object.assign(o.rows[a.r][a.c], patch); for (const k of Object.keys(patch)) if (patch[k] == null) delete o.rows[a.r][a.c][k]; return true; });
}
export const presetMenu = () => PRESET_NAMES.map(([k, l]) => ({ label: l, run: () => setOption('preset', k) }));
export const optionsMenu = () => [['header', 'Fila de encabezado'], ['banded', 'Filas con bandas'], ['firstCol', 'Primera columna'], ['lastRow', 'Última fila']].map(([k, l]) => {
  const o = C.selObjs().find((x) => x.type === 'table'); const on = o ? opts(o)[k] : false;
  return { label: `${on ? '✓ ' : '  '}${t(l)}`, run: () => setOption(k) };
});
