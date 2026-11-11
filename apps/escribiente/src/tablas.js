// Escribiente — tablas como en LibreOffice Writer: selección de celdas, combinar y dividir, filas y
// columnas que respetan las celdas combinadas, ancho de columnas arrastrando el borde, cuadro
// Propiedades de tabla, ordenar, fórmulas (=SUM(A1:A3), =SUM(ABOVE)…), estilos de tabla y
// convertir texto ↔ tabla. Todo es HTML con estilos en línea, así que se guarda en .docx y .odt. GPL-3.0
import { h, modal, toast, form, t } from '../../shared/office.js';
import { tabbed, field, numIn, sel, chk } from './formato.js';

let C = null; // { editor, shell, currentBlock, restoreSelection, saveSelection, schedulePaginate, selectedBlocks }
export function init(ctx) { C = ctx; bind(); }
const PX_CM = 2.54 / 96;
const r2 = (n) => Math.round(n * 100) / 100;
const dirty = () => { C.shell.markDirty(); C.schedulePaginate(); };
const toHex = (c) => { if (!c) return ''; if (c.startsWith('#')) return c.slice(0, 7); const m = /rgba?\(([^)]+)\)/.exec(c); if (!m) return ''; const [r, g, b, a] = m[1].split(',').map((x) => parseFloat(x)); if (a === 0) return ''; return `#${[r, g, b].map((n) => Math.round(n).toString(16).padStart(2, '0')).join('')}`; };

// ------------------------------------------------------------------ cuadrícula (cuenta colspan y rowspan)
export const rowsOf = (table) => [...table.querySelectorAll(':scope > tbody > tr, :scope > thead > tr, :scope > tfoot > tr, :scope > tr')].filter((r) => !r.classList.contains('tpush') && !r.classList.contains('trep'));
export function grid(table) {
  const rows = rowsOf(table); const m = []; const cells = [];
  rows.forEach((tr, ri) => {
    m[ri] ||= []; let ci = 0;
    for (const td of tr.children) {
      if (!/^T[DH]$/.test(td.tagName)) continue;
      while (m[ri][ci]) ci++;
      const cs = td.colSpan || 1; const rs = Math.min(td.rowSpan || 1, rows.length - ri);
      const info = { td, r: ri, c: ci, cs, rs };
      cells.push(info);
      for (let r = ri; r < ri + rs; r++) { m[r] ||= []; for (let c = ci; c < ci + cs; c++) m[r][c] = info; }
      ci += cs;
    }
  });
  const cols = Math.max(0, ...m.map((r) => r.length));
  return { rows, m, cells, cols, at: (td) => cells.find((x) => x.td === td) };
}
const mkCell = (tag = 'td', like) => { const c = document.createElement(tag); c.innerHTML = '<br>'; if (like) { const st = like.getAttribute('style'); if (st) c.setAttribute('style', st); } return c; };
// inserta la celda en la fila tr en la posición de columna c de la cuadrícula
function placeCell(G, tr, ri, c, cell) {
  const after = [...tr.children].find((td) => { const x = G.at(td); return x && x.r === ri && x.c > c; });
  tr.insertBefore(cell, after || null);
}
const curCell = () => { const b = C.currentBlock(); return b?.closest?.('td,th') || null; };
const tableOfSel = () => (selCells()[0] || curCell())?.closest('table') || null;

// ------------------------------------------------------------------ selección de celdas (arrastrar, Mayús+clic)
let anchor = null;
export const selCells = () => [...C.editor.querySelectorAll('td.csel, th.csel')];
export function clearSel() { C.editor.querySelectorAll('.csel').forEach((x) => x.classList.remove('csel')); C.editor.classList.remove('cells-selected'); }
function selectRect(a, b) {
  const table = a.closest('table'); if (!table || b.closest('table') !== table) return;
  const G = grid(table); const A = G.at(a); const B = G.at(b); if (!A || !B) return;
  let r0 = Math.min(A.r, B.r), r1 = Math.max(A.r + A.rs, B.r + B.rs) - 1, c0 = Math.min(A.c, B.c), c1 = Math.max(A.c + A.cs, B.c + B.cs) - 1;
  // ampliar hasta cubrir completas las celdas combinadas
  for (let changed = true; changed;) {
    changed = false;
    for (const x of G.cells) {
      const inter = x.r <= r1 && x.r + x.rs - 1 >= r0 && x.c <= c1 && x.c + x.cs - 1 >= c0;
      if (!inter) continue;
      if (x.r < r0) { r0 = x.r; changed = true; } if (x.r + x.rs - 1 > r1) { r1 = x.r + x.rs - 1; changed = true; }
      if (x.c < c0) { c0 = x.c; changed = true; } if (x.c + x.cs - 1 > c1) { c1 = x.c + x.cs - 1; changed = true; }
    }
  }
  clearSel();
  const list = G.cells.filter((x) => x.r >= r0 && x.r + x.rs - 1 <= r1 && x.c >= c0 && x.c + x.cs - 1 <= c1);
  if (list.length > 1) { list.forEach((x) => x.td.classList.add('csel')); C.editor.classList.add('cells-selected'); }
}
export function selectWhole(kind) {
  const cell = selCells()[0] || curCell(); if (!cell) return;
  const table = cell.closest('table'); const G = grid(table); const X = G.at(cell);
  clearSel();
  const list = kind === 'table' ? G.cells : kind === 'row' ? G.cells.filter((x) => x.r <= X.r && x.r + x.rs > X.r) : kind === 'col' ? G.cells.filter((x) => x.c <= X.c && x.c + x.cs > X.c) : [X];
  list.forEach((x) => x.td.classList.add('csel')); if (list.length > 1) C.editor.classList.add('cells-selected');
}
function bind() {
  const ed = C.editor;
  ed.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (resizeStart(e)) return;
    const td = e.target.closest?.('td,th');
    if (e.shiftKey && td && anchor && anchor.isConnected && anchor.closest('table') === td.closest('table') && td !== anchor) { e.preventDefault(); selectRect(anchor, td); return; }
    clearSel(); anchor = td || null;
    if (!td) return;
    const move = (ev) => { const o = document.elementFromPoint(ev.clientX, ev.clientY)?.closest?.('td,th'); if (o && o !== td && o.closest('table') === td.closest('table')) { selectRect(td, o); getSelection().collapseToStart?.(); } else if (o === td) clearSel(); };
    const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); };
    addEventListener('mousemove', move); addEventListener('mouseup', up);
  });
  ed.addEventListener('mousemove', hoverBorder);
  ed.addEventListener('keydown', (e) => {
    if (!selCells().length) return;
    if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); selCells().forEach((c) => { c.innerHTML = '<br>'; delete c.dataset.formula; }); dirty(); return; }
    if (!e.ctrlKey && !e.metaKey && !e.altKey && e.key.length === 1) clearSel();
    if (e.key === 'Escape') clearSel();
  });
  ed.addEventListener('input', (e) => { const td = e.target?.closest?.('td,th') || curCell(); if (td || ed.querySelector('td[data-formula],th[data-formula]')) scheduleRecalc(); });
}

// ------------------------------------------------------------------ ancho de columnas arrastrando el borde
const EDGE = 4;
function edgeAt(e) {
  const td = e.target.closest?.('td,th'); if (!td || !C.editor.contains(td)) return null;
  const r = td.getBoundingClientRect();
  if (Math.abs(e.clientX - r.right) <= EDGE) return { td, side: 'right' };
  if (Math.abs(e.clientX - r.left) <= EDGE && td.previousElementSibling) return { td: td.previousElementSibling, side: 'right' };
  return null;
}
function hoverBorder(e) { if (e.buttons) return; C.editor.classList.toggle('col-resize', !!edgeAt(e)); }
function ensureCols(table) {
  const G = grid(table);
  let cg = table.querySelector(':scope > colgroup');
  const tw = table.getBoundingClientRect().width || 600;
  const widths = Array(G.cols).fill(0);
  const cur = cg ? [...cg.children].map((c) => c.getBoundingClientRect().width || 0) : [];
  for (let c = 0; c < G.cols; c++) {
    if (cur[c]) { widths[c] = cur[c]; continue; }
    const one = G.cells.find((x) => x.c === c && x.cs === 1); widths[c] = one ? one.td.getBoundingClientRect().width : tw / G.cols;
  }
  if (!cg) { cg = document.createElement('colgroup'); table.prepend(cg); }
  cg.innerHTML = ''; const total = widths.reduce((a, b) => a + b, 0) || tw;
  widths.forEach((w) => { const col = document.createElement('col'); col.style.width = `${r2((w / total) * 100)}%`; cg.append(col); });
  return { cg, widths, total };
}
function resizeStart(e) {
  const hit = edgeAt(e); if (!hit) return false;
  e.preventDefault();
  const table = hit.td.closest('table'); const G = grid(table); const X = G.at(hit.td); const ci = X.c + X.cs - 1;
  const { cg, widths } = ensureCols(table);
  const startX = e.clientX; const w0 = [...widths]; const tw = table.getBoundingClientRect().width;
  const last = ci >= G.cols - 1;
  const move = (ev) => {
    const d = ev.clientX - startX; const w = [...w0];
    if (last) { w[ci] = Math.max(16, w0[ci] + d); const nt = Math.min(tw + d, C.editor.clientWidth * 1.2); table.style.width = `${Math.round(nt)}px`; }
    else { const tot = w0[ci] + w0[ci + 1]; w[ci] = Math.max(16, Math.min(tot - 16, w0[ci] + d)); w[ci + 1] = tot - w[ci]; }
    const sum = w.reduce((a, b) => a + b, 0);
    [...cg.children].forEach((col, k) => { col.style.width = `${r2((w[k] / sum) * 100)}%`; });
  };
  const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); dirty(); };
  addEventListener('mousemove', move); addEventListener('mouseup', up);
  return true;
}

// ------------------------------------------------------------------ filas y columnas (con celdas combinadas)
export function insertRows(where, n = 1) {
  const cell = curCell() || selCells()[0]; if (!cell) return;
  const table = cell.closest('table');
  for (let k = 0; k < n; k++) {
    const G = grid(table); const X = G.at(cell);
    const ri = where === 'above' ? X.r : X.r + X.rs - 1; // fila de referencia
    const ref = G.rows[ri];
    const nr = document.createElement('tr');
    const seen = new Set();
    for (let c = 0; c < G.cols; c++) {
      const o = G.m[ri][c]; if (!o) { nr.append(mkCell()); continue; }
      const crosses = where === 'above' ? o.r < ri : o.r + o.rs - 1 > ri;
      if (crosses) { if (!seen.has(o)) { o.td.rowSpan = o.rs + 1; seen.add(o); } c += o.cs - 1; continue; }
      const nc = mkCell(o.td.tagName === 'TH' && where === 'above' ? 'th' : 'td', o.td); if (o.cs > 1) nc.colSpan = o.cs;
      nr.append(nc); c += o.cs - 1;
    }
    ref.parentElement.insertBefore(nr, where === 'above' ? ref : ref.nextSibling);
  }
  restyle(table); dirty();
}
export function insertCols(where, n = 1) {
  const cell = curCell() || selCells()[0]; if (!cell) return;
  const table = cell.closest('table');
  for (let k = 0; k < n; k++) {
    const G = grid(table); const X = G.at(cell);
    const ci = where === 'left' ? X.c : X.c + X.cs - 1;
    const seen = new Set();
    G.rows.forEach((tr, ri) => {
      const o = G.m[ri][ci];
      if (!o) { tr.append(mkCell()); return; }
      const crosses = where === 'left' ? o.c < ci : o.c + o.cs - 1 > ci;
      if (crosses) { if (!seen.has(o)) { o.td.colSpan = o.cs + 1; seen.add(o); } return; }
      if (o.r !== ri) return; // la fila la cubre una celda combinada hacia abajo: se crea una sola celda con el mismo alto
      if (seen.has(`n${o.r}`)) return;
      const nc = mkCell(o.td.tagName, o.td); if (o.rs > 1) nc.rowSpan = o.rs;
      if (where === 'left') tr.insertBefore(nc, o.td); else o.td.after(nc);
    });
    const cg = table.querySelector(':scope > colgroup');
    if (cg) { const cols = [...cg.children]; const nc = document.createElement('col'); const ref = cols[ci]; const w = parseFloat(ref?.style.width) || 10; if (ref) { ref.style.width = `${r2(w / 2)}%`; nc.style.width = `${r2(w / 2)}%`; } if (where === 'left') cg.insertBefore(nc, ref || null); else (ref ? ref.after(nc) : cg.append(nc)); }
  }
  restyle(table); dirty();
}
export function deleteRows() {
  const cells = selCells().length ? selCells() : [curCell()].filter(Boolean); if (!cells.length) return;
  const table = cells[0].closest('table'); let G = grid(table);
  const rowsIdx = new Set(); for (const td of cells) { const X = G.at(td); for (let r = X.r; r < X.r + X.rs; r++) rowsIdx.add(r); }
  if (rowsIdx.size >= G.rows.length) { removeTable(table); return; }
  for (const ri of [...rowsIdx].sort((a, b) => b - a)) {
    G = grid(table); const tr = G.rows[ri];
    const seen = new Set();
    for (let c = 0; c < G.cols; c++) {
      const o = G.m[ri][c]; if (!o || seen.has(o)) continue; seen.add(o);
      if (o.rs > 1) {
        if (o.r === ri) { // la celda baja a la fila siguiente
          const next = G.rows[ri + 1]; o.td.rowSpan = o.rs - 1; placeCell(G, next, ri + 1, o.c, o.td);
        } else o.td.rowSpan = o.rs - 1;
      }
    }
    tr.remove();
  }
  clearSel(); dirty();
}
export function deleteCols() {
  const cells = selCells().length ? selCells() : [curCell()].filter(Boolean); if (!cells.length) return;
  const table = cells[0].closest('table'); let G = grid(table);
  const colsIdx = new Set(); for (const td of cells) { const X = G.at(td); for (let c = X.c; c < X.c + X.cs; c++) colsIdx.add(c); }
  if (colsIdx.size >= G.cols) { removeTable(table); return; }
  for (const ci of [...colsIdx].sort((a, b) => b - a)) {
    G = grid(table); const seen = new Set();
    for (let r = 0; r < G.rows.length; r++) { const o = G.m[r][ci]; if (!o || seen.has(o)) continue; seen.add(o); if (o.cs > 1) o.td.colSpan = o.cs - 1; else o.td.remove(); }
    const cg = table.querySelector(':scope > colgroup'); if (cg?.children[ci]) { const w = parseFloat(cg.children[ci].style.width) || 0; cg.children[ci].remove(); const rest = [...cg.children]; const sum = rest.reduce((a, c) => a + (parseFloat(c.style.width) || 0), 0) || 1; rest.forEach((c) => { c.style.width = `${r2(((parseFloat(c.style.width) || 0) / sum) * (sum + w))}%`; }); }
  }
  clearSel(); dirty();
}
function removeTable(table) { const p = document.createElement('p'); p.innerHTML = '<br>'; table.replaceWith(p); const r = document.createRange(); r.setStart(p, 0); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); clearSel(); dirty(); }
export function deleteTable() { const tb = tableOfSel(); if (tb) removeTable(tb); }

// ------------------------------------------------------------------ combinar y dividir
const isEmptyCell = (td) => !td.textContent.trim() && !td.querySelector('img,table,.eq');
export function mergeCells() {
  const cells = selCells(); if (cells.length < 2) { toast('Selecciona varias celdas arrastrando el ratón (o con Mayús+clic).'); return; }
  const table = cells[0].closest('table'); const G = grid(table);
  const xs = cells.map((td) => G.at(td));
  const r0 = Math.min(...xs.map((x) => x.r)), r1 = Math.max(...xs.map((x) => x.r + x.rs - 1)), c0 = Math.min(...xs.map((x) => x.c)), c1 = Math.max(...xs.map((x) => x.c + x.cs - 1));
  const area = xs.reduce((a, x) => a + x.cs * x.rs, 0);
  if (area !== (r1 - r0 + 1) * (c1 - c0 + 1)) { toast('Las celdas seleccionadas deben formar un rectángulo.', { error: true }); return; }
  xs.sort((a, b) => a.r - b.r || a.c - b.c);
  const first = xs[0].td;
  // el contenido de todas las celdas se junta, un párrafo por celda
  const parts = xs.filter((x) => !isEmptyCell(x.td)).map((x) => x.td.innerHTML);
  first.innerHTML = parts.length > 1 ? parts.map((p) => (/^\s*<(p|h\d|ul|ol|table|div)\b/i.test(p) ? p : `<p>${p}</p>`)).join('') : (parts[0] || '<br>');
  for (const x of xs.slice(1)) x.td.remove();
  first.colSpan = c1 - c0 + 1; first.rowSpan = r1 - r0 + 1;
  if (first.colSpan === 1) first.removeAttribute('colspan'); if (first.rowSpan === 1) first.removeAttribute('rowspan');
  // filas que quedaron vacías (todas sus celdas combinadas arriba)
  for (const tr of rowsOf(table)) if (!tr.children.length) { const G2 = grid(table); const ri = G2.rows.indexOf(tr); G2.cells.forEach((x) => { if (x.r < ri && x.r + x.rs > ri) x.td.rowSpan = Math.max(1, x.rs - 1); }); tr.remove(); }
  clearSel(); placeCaret(first); dirty();
}
export async function splitDialog() {
  const cell = selCells()[0] || curCell(); if (!cell) return;
  const v = await form('Dividir celdas', [{ name: 'n', label: 'Dividir en', type: 'number', value: 2, min: 2, max: 20 }, { name: 'dir', label: 'Dirección', type: 'select', value: 'v', options: [['v', t('Verticalmente (columnas)')], ['h', t('Horizontalmente (filas)')]] }], 'Aceptar');
  if (v) splitCell(cell, Math.max(2, v.n | 0), v.dir);
}
export function splitCell(cell, n, dir) {
  const table = cell.closest('table');
  let G = grid(table); const X = G.at(cell);
  // celda combinada en menos partes de las pedidas: primero se separa y luego se divide la última parte
  const span = dir === 'v' ? X.cs : X.rs;
  if (span > 1 && span < n) {
    splitCell(cell, span, dir);
    const G2 = grid(table); const last = dir === 'v' ? G2.m[X.r][X.c + X.cs - 1].td : G2.m[X.r + X.rs - 1][X.c].td;
    return splitCell(last, n - span + 1, dir);
  }
  if (dir === 'v') {
    if (X.cs >= n) { // repartir la combinación existente
      const base = Math.floor(X.cs / n); let extra = X.cs % n; cell.colSpan = base + (extra-- > 0 ? 1 : 0);
      let prev = cell; for (let k = 1; k < n; k++) { const nc = mkCell(cell.tagName, cell); nc.colSpan = base + (extra-- > 0 ? 1 : 0); if (X.rs > 1) nc.rowSpan = X.rs; prev.after(nc); prev = nc; }
    } else {
      const add = n - X.cs; const ci = X.c + X.cs - 1;
      const seen = new Set();
      G.rows.forEach((tr, ri) => { const o = G.m[ri][ci]; if (!o || seen.has(o) || o === X) return; seen.add(o); o.td.colSpan = o.cs + add; });
      cell.colSpan = 1; let prev = cell; for (let k = 1; k < n; k++) { const nc = mkCell(cell.tagName, cell); if (X.rs > 1) nc.rowSpan = X.rs; prev.after(nc); prev = nc; }
      const cg = table.querySelector(':scope > colgroup'); if (cg?.children[ci]) { const w = parseFloat(cg.children[ci].style.width) || 10; cg.children[ci].style.width = `${r2(w / n)}%`; for (let k = 1; k < n; k++) { const c2 = document.createElement('col'); c2.style.width = `${r2(w / n)}%`; cg.children[ci].after(c2); } }
    }
  } else if (X.rs >= n) {
    const base = Math.floor(X.rs / n); let extra = X.rs % n; cell.rowSpan = base + (extra-- > 0 ? 1 : 0);
    let r = X.r + cell.rowSpan;
    for (let k = 1; k < n; k++) { G = grid(table); const nc = mkCell(cell.tagName, cell); nc.rowSpan = base + (extra-- > 0 ? 1 : 0); if (X.cs > 1) nc.colSpan = X.cs; placeCell(G, G.rows[r], r, X.c, nc); r += nc.rowSpan; }
  } else {
    const add = n - X.rs; const last = X.r + X.rs - 1;
    const seen = new Set();
    for (let c = 0; c < G.cols; c++) { const o = G.m[last][c]; if (!o || seen.has(o) || o === X) continue; seen.add(o); if (o.r + o.rs - 1 === last) o.td.rowSpan = o.rs + add; else o.td.rowSpan = o.rs + add; }
    let ref = G.rows[last];
    for (let k = 0; k < add; k++) { const tr = document.createElement('tr'); const nc = mkCell(cell.tagName, cell); if (X.cs > 1) nc.colSpan = X.cs; tr.append(nc); ref.after(tr); ref = tr; }
    // las celdas nuevas quedan en la columna correcta porque las demás columnas están cubiertas por rowspan
  }
  normalize(table); clearSel(); dirty();
}
// quita colspan/rowspan iguales a 1
function normalize(table) { table.querySelectorAll('td,th').forEach((td) => { if (td.colSpan === 1) td.removeAttribute('colspan'); if (td.rowSpan === 1) td.removeAttribute('rowspan'); }); }
export function splitTable() {
  const cell = curCell(); if (!cell) return;
  const table = cell.closest('table'); const G = grid(table); const X = G.at(cell);
  if (X.r === 0) { toast('Coloca el cursor en una fila que no sea la primera.'); return; }
  if (G.cells.some((x) => x.r < X.r && x.r + x.rs > X.r)) { toast('No se puede dividir por una celda combinada verticalmente.', { error: true }); return; }
  const nt = table.cloneNode(false); const cg = table.querySelector(':scope > colgroup'); if (cg) nt.append(cg.cloneNode(true));
  const tb = document.createElement('tbody'); nt.append(tb);
  G.rows.slice(X.r).forEach((tr) => tb.append(tr));
  const p = document.createElement('p'); p.innerHTML = '<br>';
  table.after(p, nt); placeCaret(p); dirty();
}
function placeCaret(el) { const r = document.createRange(); r.selectNodeContents(el); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); C.saveSelection(); }

// ------------------------------------------------------------------ tamaño
export function distributeCols() {
  const table = tableOfSel(); if (!table) return;
  const G = grid(table); let cg = table.querySelector(':scope > colgroup');
  if (!cg) { cg = document.createElement('colgroup'); table.prepend(cg); }
  cg.innerHTML = ''; for (let c = 0; c < G.cols; c++) { const col = document.createElement('col'); col.style.width = `${r2(100 / G.cols)}%`; cg.append(col); }
  dirty();
}
export function distributeRows() {
  const table = tableOfSel(); if (!table) return;
  const rows = rowsOf(table); const hmax = Math.max(...rows.map((r) => r.getBoundingClientRect().height));
  rows.forEach((r) => { r.style.height = `${r2(hmax * 0.75)}pt`; }); dirty();
}
export function optimalWidth() {
  const table = tableOfSel(); if (!table) return;
  table.querySelector(':scope > colgroup')?.remove();
  const prev = table.style.width; table.style.tableLayout = 'auto'; table.style.width = 'auto';
  const G = grid(table); const w = Array(G.cols).fill(0);
  for (const x of G.cells) if (x.cs === 1) w[x.c] = Math.max(w[x.c], x.td.getBoundingClientRect().width);
  const tw = table.getBoundingClientRect().width; const maxW = C.editor.clientWidth;
  table.style.tableLayout = ''; table.style.width = tw < maxW ? `${Math.round(tw)}px` : (prev || '');
  const sum = w.reduce((a, b) => a + b, 0) || 1;
  const cg = document.createElement('colgroup'); w.forEach((x) => { const col = document.createElement('col'); col.style.width = `${r2((x / sum) * 100)}%`; cg.append(col); }); table.prepend(cg);
  dirty();
}
export function cellVAlign(v) { const cells = selCells().length ? selCells() : [curCell()].filter(Boolean); cells.forEach((c) => { c.style.verticalAlign = v; }); dirty(); }

// ------------------------------------------------------------------ bordes y fondo
export const BORDER_PRESETS = [['all', 'Todos los bordes'], ['outer', 'Solo contorno'], ['inner', 'Solo bordes interiores'], ['outerH', 'Contorno y líneas horizontales'], ['horiz', 'Solo líneas horizontales'], ['none', 'Sin bordes']];
export function applyBorders(preset, { style = 'solid', width = 0.75, color = '#000000', cells, table } = {}) {
  cells ||= selCells().length ? selCells() : null;
  table ||= (cells?.[0] || curCell())?.closest('table'); if (!table) return;
  const G = grid(table);
  const list = cells ? cells.map((td) => G.at(td)) : G.cells;
  const r0 = Math.min(...list.map((x) => x.r)), r1 = Math.max(...list.map((x) => x.r + x.rs - 1)), c0 = Math.min(...list.map((x) => x.c)), c1 = Math.max(...list.map((x) => x.c + x.cs - 1));
  const line = `${width}pt ${style} ${color}`; const none = '0px none transparent';
  for (const x of list) {
    const edge = { top: x.r === r0, bottom: x.r + x.rs - 1 === r1, left: x.c === c0, right: x.c + x.cs - 1 === c1 };
    for (const s of ['top', 'bottom', 'left', 'right']) {
      const outer = edge[s]; const horiz = s === 'top' || s === 'bottom';
      const on = preset === 'all' || (preset === 'outer' && outer) || (preset === 'inner' && !outer) || (preset === 'outerH' && (outer || horiz)) || (preset === 'horiz' && horiz);
      x.td.style.setProperty(`border-${s}`, on ? line : none);
    }
  }
  dirty();
}
export function shadeCells(color) { const cells = selCells().length ? selCells() : [curCell()].filter(Boolean); cells.forEach((c) => { c.style.backgroundColor = color === 'transparent' || !color ? '' : color; }); dirty(); }

// ------------------------------------------------------------------ estilos de tabla (autoformato)
export const TABLE_STYLES = {
  default: { label: 'Predeterminado', border: '#9A9EAD', head: '#EEECFB' },
  grid: { label: 'Rejilla', border: '#000000' },
  none: { label: 'Sin bordes', border: null },
  blue: { label: 'Encabezado azul', border: '#8EAADB', head: '#2F5496', headInk: '#FFFFFF', band: '#DEEAF6' },
  green: { label: 'Verde', border: '#A8D08D', head: '#538135', headInk: '#FFFFFF', band: '#E2EFD9' },
  orange: { label: 'Naranja', border: '#F4B083', head: '#C55A11', headInk: '#FFFFFF', band: '#FBE4D5' },
  gray: { label: 'Gris', border: '#BFBFBF', head: '#595959', headInk: '#FFFFFF', band: '#F2F2F2' },
  elegant: { label: 'Elegante', rules: true, headBold: true },
  bands: { label: 'Filas alternas', border: null, band: '#EEF1F8', headBold: true, headLine: '#4F5D85' },
};
export function applyTableStyle(key, table = tableOfSel()) {
  const S = TABLE_STYLES[key]; if (!S || !table) return;
  table.dataset.tstyle = key; restyle(table); dirty();
}
// vuelve a aplicar el estilo de tabla (después de añadir filas o columnas)
function restyle(table) {
  const S = TABLE_STYLES[table.dataset.tstyle]; if (!S) return;
  const G = grid(table); const n = G.rows.length; const hasHead = G.rows[0] && [...G.rows[0].children].every((c) => c.tagName === 'TH') || !!S.head || S.headBold;
  for (const x of G.cells) {
    const st = x.td.style; const head = hasHead && x.r === 0;
    const b = S.border ? `0.75pt solid ${S.border}` : '0px none transparent';
    ['top', 'bottom', 'left', 'right'].forEach((s) => st.setProperty(`border-${s}`, b));
    if (S.rules) { ['left', 'right'].forEach((s) => st.setProperty(`border-${s}`, '0px none transparent')); st.borderTop = x.r === 0 ? '1.5pt solid #000000' : x.r === 1 ? '0.75pt solid #000000' : '0px none transparent'; st.borderBottom = x.r + x.rs === n ? '1.5pt solid #000000' : '0px none transparent'; }
    if (S.headLine && head) st.borderBottom = `1.5pt solid ${S.headLine}`;
    st.backgroundColor = head && S.head ? S.head : !head && S.band && (x.r - (hasHead ? 1 : 0)) % 2 === 1 ? S.band : (head || key0(S)) ? 'transparent' : '';
    st.color = head && S.headInk ? S.headInk : '';
    st.fontWeight = head && (S.head || S.headBold) ? 'bold' : '';
  }
}
const key0 = (S) => !S.head; // sin color de encabezado: el encabezado queda transparente
export const tableStyleMenu = () => Object.entries(TABLE_STYLES).map(([k, S]) => ({ label: S.label, run: () => applyTableStyle(k) }));

// ------------------------------------------------------------------ cuadro Propiedades de tabla
export async function propertiesDialog(start = 0) {
  const cell = selCells()[0] || curCell(); if (!cell) { toast('Coloca el cursor en una tabla.'); return; }
  const table = cell.closest('table'); const G = grid(table); const X = G.at(cell);
  const ts = table.style; const tw = table.getBoundingClientRect().width; const cw = C.editor.clientWidth;
  const align = ts.marginLeft === 'auto' && ts.marginRight === 'auto' ? 'center' : ts.marginLeft === 'auto' ? 'right' : 'left';
  const widthPct = /%$/.test(ts.width) ? parseFloat(ts.width) : r2((tw / cw) * 100);
  const { widths } = ensureCols(table); const sumW = widths.reduce((a, b) => a + b, 0) || 1;
  const cs = getComputedStyle(cell); const row = G.rows[X.r];
  const head = [...G.rows[0].children].every((c) => c.tagName === 'TH');
  const colInputs = widths.map((w, i) => field(`${t('Columna')} ${i + 1}`, numIn(`col${i}`, r2(((w / sumW) * tw) * PX_CM), { step: 0.1, min: 0.3, max: 50, unit: 'cm' })));
  const body = await tabbed('Propiedades de tabla', [
    { label: 'Tabla', body: h('div', { class: 'fx-grid' },
      h('h4', {}, t('Tamaño')), field('Ancho (% de la página)', numIn('tw', Math.min(100, Math.round(widthPct)), { step: 1, min: 10, max: 120, unit: '%' })), field('Alineación', sel('talign', align, [['left', 'Izquierda'], ['center', 'Centrado'], ['right', 'Derecha']])),
      field('Sangría izquierda', numIn('tind', r2((parseFloat(ts.marginLeft) || 0) * PX_CM), { step: 0.1, min: 0, max: 20, unit: 'cm' })),
      h('h4', {}, t('Encabezado')), chk('head', 'Repetir la primera fila como encabezado', head), chk('nosplit', 'No dividir la tabla entre páginas', table.dataset.nosplit === '1')) },
    { label: 'Columnas', body: h('div', { class: 'fx-grid' }, ...colInputs, h('p', { class: 'fx-note', style: { gridColumn: '1 / -1' } }, t('También puedes arrastrar los bordes de las columnas en el documento.'))) },
    { label: 'Fila', body: h('div', { class: 'fx-grid' }, field('Alto mínimo de la fila', numIn('rh', r2((parseFloat(row.style.height) || 0) / 0.75 * PX_CM * (/(pt)$/.test(row.style.height) ? 1 : 0.75)), { step: 0.1, min: 0, max: 30, unit: 'cm' })), chk('rkeep', 'No dividir la fila entre páginas', row.dataset.keep === '1')) },
    { label: 'Celda', body: h('div', { class: 'fx-grid' },
      field('Alineación vertical', sel('va', cs.verticalAlign === 'middle' ? 'middle' : cs.verticalAlign === 'bottom' ? 'bottom' : 'top', [['top', 'Arriba'], ['middle', 'Centro'], ['bottom', 'Abajo']])),
      field('Relleno (márgenes de celda)', numIn('pad', r2(parseFloat(cs.paddingLeft) * 0.75), { step: 0.5, min: 0, max: 30, unit: 'pt' })),
      field('Color de fondo', h('input', { type: 'color', name: 'bg', value: toHex(cs.backgroundColor) || '#ffffff' })), chk('nobg', 'Sin fondo', !toHex(cs.backgroundColor))) },
    { label: 'Bordes', body: h('div', { class: 'fx-grid' },
      field('Disposición', sel('bp', '', [['', '(sin cambios)'], ...BORDER_PRESETS])), field('Estilo', sel('bs', 'solid', [['solid', 'Sólido'], ['double', 'Doble'], ['dotted', 'Punteado'], ['dashed', 'Discontinuo']])),
      field('Ancho', numIn('bw', 0.75, { step: 0.25, min: 0.25, max: 6, unit: 'pt' })), field('Color', h('input', { type: 'color', name: 'bc', value: '#000000' })),
      field('Aplicar a', sel('bto', selCells().length ? 'sel' : 'table', [['table', 'Toda la tabla'], ['sel', 'Celdas seleccionadas']]))) },
  ], { start, width: '600px' });
  if (!body) return;
  const v = (n) => body.querySelector(`[name="${n}"]`); const num = (n) => Number(v(n).value) || 0;
  // tabla
  const pct = num('tw'); table.style.width = `${pct}%`;
  const a = v('talign').value; table.style.marginLeft = a === 'left' ? (num('tind') ? `${num('tind')}cm` : '') : 'auto'; table.style.marginRight = a === 'center' ? 'auto' : '';
  if (v('nosplit').checked) table.dataset.nosplit = '1'; else delete table.dataset.nosplit;
  if (v('head').checked !== head) setHeaderRow(table, v('head').checked);
  // columnas
  const cmW = widths.map((_, i) => Math.max(0.3, num(`col${i}`)));
  const tot = cmW.reduce((p, q) => p + q, 0); const cg = table.querySelector(':scope > colgroup');
  if (cg) [...cg.children].forEach((col, i) => { col.style.width = `${r2((cmW[i] / tot) * 100)}%`; });
  // fila (las filas seleccionadas o la actual)
  const rowsSel = new Set((selCells().length ? selCells() : [cell]).map((td) => td.parentElement));
  for (const tr of rowsSel) { tr.style.height = num('rh') ? `${r2(num('rh') / PX_CM * 0.75)}pt` : ''; if (v('rkeep').checked) tr.dataset.keep = '1'; else delete tr.dataset.keep; }
  // celdas
  const cells = selCells().length ? selCells() : [cell];
  for (const td of cells) { td.style.verticalAlign = v('va').value; const p = `${num('pad')}pt`; td.style.padding = `${r2(num('pad') * 0.6)}pt ${p}`; td.style.backgroundColor = v('nobg').checked ? '' : v('bg').value; }
  if (v('bp').value) applyBorders(v('bp').value, { style: v('bs').value, width: num('bw') || 0.75, color: v('bc').value, cells: v('bto').value === 'sel' && selCells().length ? selCells() : null, table });
  dirty();
}
export function setHeaderRow(table, on) {
  const r0 = rowsOf(table)[0]; if (!r0) return;
  [...r0.children].forEach((c) => { if ((c.tagName === 'TH') === on) return; const n = document.createElement(on ? 'th' : 'td'); for (const at of c.attributes) n.setAttribute(at.name, at.value); n.innerHTML = c.innerHTML; c.replaceWith(n); });
  dirty();
}
export function toggleHeaderRow() { const tb = tableOfSel(); if (!tb) return; const r0 = rowsOf(tb)[0]; setHeaderRow(tb, ![...r0.children].every((c) => c.tagName === 'TH')); }

// ------------------------------------------------------------------ ordenar
const parseNum = (s) => {
  let x = String(s).replace(/[\s$€%]/g, ''); if (!x) return NaN;
  if (/,\d{1,2}$/.test(x) && !/\.\d{1,2}$/.test(x)) x = x.replace(/\./g, '').replace(',', '.'); else x = x.replace(/,/g, '');
  return /^[-+]?\d*\.?\d+(e[-+]?\d+)?$/i.test(x) ? Number(x) : NaN;
};
export async function sortDialog() {
  const cell = curCell() || selCells()[0]; if (!cell) { toast('Coloca el cursor en una tabla.'); return; }
  const table = cell.closest('table'); const G = grid(table);
  if (G.cells.some((x) => x.rs > 1)) { toast('No se puede ordenar una tabla con celdas combinadas verticalmente.', { error: true }); return; }
  const head = [...G.rows[0].children].every((c) => c.tagName === 'TH');
  const colOpts = Array.from({ length: G.cols }, (_, i) => [String(i), `${t('Columna')} ${i + 1}${head ? ` (${(G.m[0][i]?.td.textContent || '').trim().slice(0, 20)})` : ''}`]);
  const v = await form('Ordenar', [
    { name: 'k1', label: 'Ordenar por', type: 'select', value: String(G.at(cell).c), options: colOpts },
    { name: 'type', label: 'Tipo de clave', type: 'select', value: 'auto', options: [['auto', t('Automático')], ['num', t('Numérico')], ['text', t('Alfanumérico')]] },
    { name: 'order', label: 'Orden', type: 'select', value: 'asc', options: [['asc', t('Ascendente')], ['desc', t('Descendente')]] },
    { name: 'k2', label: 'Después por', type: 'select', value: '', options: [['', t('(ninguna)')], ...colOpts] },
    { name: 'head', label: 'La primera fila es encabezado', type: 'checkbox', value: head || G.rows.length > 1 && head },
  ], 'Ordenar');
  if (!v) return;
  sortTable(table, { keys: [Number(v.k1), v.k2 === '' ? null : Number(v.k2)].filter((k) => k != null), type: v.type, desc: v.order === 'desc', head: v.head });
}
export function sortTable(table, { keys, type = 'auto', desc = false, head = false }) {
  const G = grid(table);
  const body = G.rows.slice(head ? 1 : 0);
  const val = (tr, k) => { const ri = G.rows.indexOf(tr); return (G.m[ri][k]?.td.textContent || '').trim(); };
  const coll = new Intl.Collator(document.documentElement.lang || 'es', { numeric: true, sensitivity: 'base' });
  const cmp = (a, b) => {
    for (const k of keys) {
      const x = val(a, k), y = val(b, k); const nx = parseNum(x), ny = parseNum(y);
      let r;
      if (type === 'num' || (type === 'auto' && !Number.isNaN(nx) && !Number.isNaN(ny))) r = (Number.isNaN(nx) ? -Infinity : nx) - (Number.isNaN(ny) ? -Infinity : ny);
      else r = coll.compare(x, y);
      if (r) return desc ? -r : r;
    }
    return 0;
  };
  const sorted = [...body].sort(cmp);
  const parent = body[0]?.parentElement; if (!parent) return;
  for (const tr of sorted) parent.append(tr);
  restyle(table); recalcAll(); dirty();
}

// ------------------------------------------------------------------ fórmulas en celdas
const colName = (c) => { let s = ''; c += 1; while (c) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
const colIdx = (s) => [...s.toUpperCase()].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
export const cellName = (td) => { const G = grid(td.closest('table')); const X = G.at(td); return `${colName(X.c)}${X.r + 1}`; };
function cellValue(G, r, c) { const o = G.m[r]?.[c]; if (!o) return NaN; return parseNum(o.td.textContent.trim()); }
// evalúa una fórmula estilo Word/Writer: =SUM(A1:B3), =A1*2+B2, =AVERAGE(ABOVE), =MAX(LEFT)
export function evalFormula(table, td, formula) {
  const G = grid(table); const X = G.at(td);
  let src = String(formula).trim().replace(/^=/, '').replace(/<([A-Z]+\d+(?::[A-Z]+\d+)?)>/gi, '$1');
  const range = (a, b) => {
    const p = (s) => { const m = /^([A-Z]+)(\d+)$/i.exec(s); return m && [Number(m[2]) - 1, colIdx(m[1])]; };
    const A = p(a); const B = b ? p(b) : A; if (!A || !B) throw new Error('ref');
    const out = []; for (let r = Math.min(A[0], B[0]); r <= Math.max(A[0], B[0]); r++) for (let c = Math.min(A[1], B[1]); c <= Math.max(A[1], B[1]); c++) { const o = G.m[r]?.[c]; if (o && o.r === r && o.c === c && o !== X) out.push(cellValue(G, r, c)); }
    return out;
  };
  const dirRange = (d) => {
    const out = [];
    if (d === 'ABOVE') for (let r = X.r - 1; r >= 0; r--) out.push(cellValue(G, r, X.c));
    if (d === 'BELOW') for (let r = X.r + X.rs; r < G.rows.length; r++) out.push(cellValue(G, r, X.c));
    if (d === 'LEFT') for (let c = X.c - 1; c >= 0; c--) out.push(cellValue(G, X.r, c));
    if (d === 'RIGHT') for (let c = X.c + X.cs; c < G.cols; c++) out.push(cellValue(G, X.r, c));
    return out;
  };
  const FN = {
    SUM: (a) => a.reduce((p, q) => p + q, 0), AVERAGE: (a) => (a.length ? FN.SUM(a) / a.length : 0), MIN: (a) => Math.min(...a), MAX: (a) => Math.max(...a),
    COUNT: (a) => a.length, PRODUCT: (a) => a.reduce((p, q) => p * q, 1), ABS: (a) => Math.abs(a[0]), ROUND: (a) => { const k = 10 ** (a[1] || 0); return Math.round(a[0] * k) / k; }, INT: (a) => Math.trunc(a[0]),
  };
  const ALIAS = { SUMA: 'SUM', PROMEDIO: 'AVERAGE', MEAN: 'AVERAGE', MAX: 'MAX', MIN: 'MIN', CONTAR: 'COUNT', PRODUCTO: 'PRODUCT', REDONDEAR: 'ROUND', ENTERO: 'INT' };
  // tokenizador y analizador recursivo (sin eval)
  const toks = []; const re = /\s*(?:(\d+(?:[.,]\d+)?%?)|([A-Za-zÁÉÍÓÚáéíóú]+\d*(?::[A-Za-z]+\d+)?)|(\*\*|[-+*/^(),;]))/y;
  let m; while (re.lastIndex < src.length && (m = re.exec(src))) toks.push(m[1] ? { n: m[1] } : m[2] ? { id: m[2].toUpperCase() } : { op: m[3] });
  if (re.lastIndex < src.trim().length && src.slice(re.lastIndex).trim()) throw new Error('sintaxis');
  let i = 0; const peek = () => toks[i]; const take = () => toks[i++];
  const isOp = (o) => peek()?.op === o;
  const args = () => { const out = []; if (isOp(')')) { take(); return out; } for (;;) { const save = i; const tk = peek(); if (tk?.id && /^(ABOVE|BELOW|LEFT|RIGHT)$/.test(tk.id)) { take(); out.push(...dirRange(tk.id)); } else if (tk?.id && tk.id.includes(':')) { take(); const [a, b] = tk.id.split(':'); out.push(...range(a, b)); } else { i = save; out.push(expr()); } if (isOp(',') || isOp(';')) { take(); continue; } if (isOp(')')) { take(); return out; } throw new Error('sintaxis'); } };
  const prim = () => {
    const tk = take(); if (!tk) throw new Error('sintaxis');
    if (tk.n) { const pc = tk.n.endsWith('%'); const n = Number(tk.n.replace('%', '').replace(',', '.')); return pc ? n / 100 : n; }
    if (tk.op === '(') { const v = expr(); if (!isOp(')')) throw new Error('sintaxis'); take(); return v; }
    if (tk.op === '-') return -prim(); if (tk.op === '+') return prim();
    if (tk.id) {
      const name = ALIAS[tk.id] || tk.id;
      if (isOp('(')) { take(); const f = FN[name]; if (!f) throw new Error(`función ${tk.id}`); return f(args().filter((x) => !Number.isNaN(x))); }
      if (/^[A-Z]+\d+$/.test(tk.id)) { const v = range(tk.id)[0]; return Number.isNaN(v) || v == null ? 0 : v; }
    }
    throw new Error('sintaxis');
  };
  const pow = () => { let v = prim(); while (isOp('^') || isOp('**')) { take(); v **= prim(); } return v; };
  const term = () => { let v = pow(); while (isOp('*') || isOp('/')) { const o = take().op; const w = pow(); v = o === '*' ? v * w : v / w; } return v; };
  const expr = () => { let v = term(); while (isOp('+') || isOp('-')) { const o = take().op; const w = term(); v = o === '+' ? v + w : v - w; } return v; };
  const v = expr(); if (i < toks.length) throw new Error('sintaxis');
  return v;
}
const fmtNum = (n) => (Number.isFinite(n) ? n.toLocaleString(document.documentElement.lang === 'en' ? 'en-US' : 'es-MX', { maximumFractionDigits: 2 }) : '#ERROR');
export function setFormula(td, formula) {
  const table = td.closest('table');
  td.dataset.formula = formula.startsWith('=') ? formula : `=${formula}`;
  recalcCell(table, td);
}
function recalcCell(table, td) {
  let txt; try { txt = fmtNum(evalFormula(table, td, td.dataset.formula)); } catch { txt = '#ERROR'; }
  if (td.textContent.trim() !== txt) { const keep = td.querySelector('p'); if (keep) { keep.textContent = txt; [...td.children].slice(1).forEach((x) => x.remove()); } else td.textContent = txt; }
}
export function recalcAll() {
  // dos pasadas: una fórmula puede usar el resultado de otra
  for (let k = 0; k < 2; k++) for (const td of C.editor.querySelectorAll('td[data-formula], th[data-formula]')) recalcCell(td.closest('table'), td);
}
let rt = 0;
function scheduleRecalc() {
  clearTimeout(rt);
  rt = setTimeout(() => {
    const active = curCell();
    for (let k = 0; k < 2; k++) for (const td of C.editor.querySelectorAll('td[data-formula], th[data-formula]')) if (td !== active) recalcCell(td.closest('table'), td);
  }, 250);
}
export async function formulaDialog() {
  const td = curCell() || selCells()[0]; if (!td) { toast('Coloca el cursor en la celda donde va el resultado.'); return; }
  const v = await form(`${t('Fórmula')} — ${cellName(td)}`, [
    { name: 'f', label: 'Fórmula', value: td.dataset.formula || '=SUM(ABOVE)', placeholder: '=SUM(A2:A5)' },
    { name: 'help', type: 'note', label: t('Funciones: SUM, AVERAGE, MIN, MAX, COUNT, PRODUCT, ROUND, ABS. Referencias: A1, A1:B3, ABOVE (arriba), LEFT (izquierda), BELOW, RIGHT. Operadores: + − * / ^ %.') },
  ], 'Aceptar');
  if (!v) return;
  if (!v.f.trim() || v.f.trim() === '=') { delete td.dataset.formula; dirty(); return; }
  try { evalFormula(td.closest('table'), td, v.f); } catch (e) { toast(`${t('Fórmula no válida')}: ${e.message}`, { error: true }); return; }
  setFormula(td, v.f.trim()); dirty();
}

// ------------------------------------------------------------------ convertir texto ↔ tabla
export async function textToTableDialog() {
  C.restoreSelection();
  const blocks = C.selectedBlocks().filter((b) => /^(P|H[1-6]|DIV|LI)$/.test(b.tagName) && !b.closest('table'));
  if (!blocks.length) { toast('Selecciona los párrafos que quieres convertir.'); return; }
  const v = await form('Convertir texto en tabla', [
    { name: 'sep', label: 'Separar el texto en', type: 'select', value: 'tab', options: [['tab', t('Tabulaciones')], [';', t('Punto y coma')], [',', t('Comas')], ['para', t('Párrafos (una columna)')], ['other', t('Otro carácter')]] },
    { name: 'other', label: 'Otro carácter', value: '|' },
    { name: 'head', label: 'La primera fila es encabezado', type: 'checkbox', value: true },
  ], 'Convertir');
  if (!v) return;
  textToTable(blocks, v.sep === 'other' ? v.other || '|' : v.sep, v.head);
}
const textWithTabs = (b) => { const c = b.cloneNode(true); c.querySelectorAll('.otab').forEach((x) => x.replaceWith('\t')); c.querySelectorAll('br').forEach((x) => x.replaceWith('\n')); return c.textContent; };
export function textToTable(blocks, sep = 'tab', head = true) {
  const rows = blocks.map((b) => (sep === 'para' ? [b.textContent] : textWithTabs(b).split(sep === 'tab' ? /\t+/ : sep).map((s) => s.trim())));
  const cols = Math.max(...rows.map((r) => r.length));
  const esc = (s) => String(s).replace(/[&<>]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[ch]));
  const table = document.createElement('table'); const tb = document.createElement('tbody'); table.append(tb);
  rows.forEach((r, ri) => { const tr = document.createElement('tr'); for (let c = 0; c < cols; c++) { const cell = document.createElement(head && ri === 0 ? 'th' : 'td'); cell.innerHTML = r[c] ? esc(r[c]) : '<br>'; tr.append(cell); } tb.append(tr); });
  blocks[0].before(table); blocks.forEach((b) => b.remove());
  if (!table.nextElementSibling) { const p = document.createElement('p'); p.innerHTML = '<br>'; table.after(p); }
  placeCaret(table.querySelector('td,th')); dirty();
  return table;
}
export async function tableToTextDialog() {
  const table = tableOfSel(); if (!table) { toast('Coloca el cursor en una tabla.'); return; }
  const v = await form('Convertir tabla en texto', [{ name: 'sep', label: 'Separar las celdas con', type: 'select', value: 'tab', options: [['tab', t('Tabulaciones')], [';', t('Punto y coma')], [',', t('Comas')], ['para', t('Párrafos')]] }], 'Convertir');
  if (v) tableToText(table, v.sep);
}
export function tableToText(table, sep = 'tab') {
  const G = grid(table); const frag = document.createDocumentFragment();
  const esc = (s) => String(s).replace(/[&<>]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[ch]));
  G.rows.forEach((tr, ri) => {
    const vals = G.cells.filter((x) => x.r === ri).sort((a, b) => a.c - b.c).map((x) => x.td.innerText.replace(/\n+/g, ' ').trim());
    if (sep === 'para') vals.forEach((s) => { const p = document.createElement('p'); p.innerHTML = s ? esc(s) : '<br>'; frag.append(p); });
    else { const p = document.createElement('p'); p.innerHTML = sep === 'tab' ? vals.map(esc).join('<span class="otab" contenteditable="false">\t</span>') : esc(vals.join(`${sep} `)); frag.append(p); }
  });
  const first = frag.firstChild; table.replaceWith(frag); if (first) placeCaret(first);
  C.layoutTabs?.(); dirty();
}

// ------------------------------------------------------------------ insertar varias filas o columnas
export async function insertDialog(kind) {
  const v = await form(kind === 'row' ? 'Insertar filas' : 'Insertar columnas', [{ name: 'n', label: 'Cantidad', type: 'number', value: 1, min: 1, max: 100 }, { name: 'where', label: 'Posición', type: 'select', value: kind === 'row' ? 'below' : 'right', options: kind === 'row' ? [['above', t('Antes')], ['below', t('Después')]] : [['left', t('Antes')], ['right', t('Después')]] }], 'Insertar');
  if (!v) return;
  if (kind === 'row') insertRows(v.where, Math.max(1, v.n | 0)); else insertCols(v.where, Math.max(1, v.n | 0));
}

// datos de las fórmulas para guardarlos dentro del archivo (Word y LibreOffice ven el resultado)
export function formulasMeta(root) { const tables = [...root.querySelectorAll('table')]; const out = []; tables.forEach((tb, ti) => { const G = grid(tb); G.cells.forEach((x, k) => { if (x.td.dataset.formula) out.push({ ti, k, f: x.td.dataset.formula }); }); }); return out; }
export function applyFormulasMeta(root, list) { const tables = [...root.querySelectorAll('table')]; for (const it of list || []) { const tb = tables[it.ti]; if (!tb) continue; const G = grid(tb); const x = G.cells[it.k]; if (x) x.td.dataset.formula = it.f; } }
