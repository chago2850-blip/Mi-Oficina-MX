// Calcula Datos — hoja de cálculo de Mi Oficina MX. GPL-3.0-or-later
import { OfficeShell, h, $, $$, icon, toast, modal, form, dropdown, closeDropdown, esc, api, confirmBox, onLang, t, announce } from '../../shared/office.js';
import * as echarts from 'echarts/dist/echarts.esm.min.mjs';
import { Workbook, colName, addr, rangeStr, parseRange, parseAddr } from './workbook.js';
import { Grid } from './grid.js';
import { readXlsx, writeXlsx, parseCsv, toCsv } from './xlsx.js';
import { toSpanish, toEnglish, FUNCTION_HELP } from './i18n.js';
import { FORMATS, changeDecimals, parseInput, formatValue } from './format.js';
import * as F from './features.js';
import * as F3 from './fase3.js';
import * as Q from './query.js';
import { readOds, writeOds } from './ods.js';
import { injectVba } from './xlsx.js';
import * as CE from './celdas.js';
import * as DT from './datos.js';
import * as AN from './analisis.js';

const wb = new Workbook();
let shell, grid, nameBox, fxInput, tabsBar, statusInfo, zoom;
const FONTS = ['Calibri', 'Arial', 'Cambria', 'Times New Roman', 'Verdana', 'Tahoma', 'Georgia', 'Courier New', 'Inter'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48];

// ------------------------------------------------------------------ helpers de acciones
const sel = () => grid.range;
const act = (label, fn, opts) => { wb.action(label, fn, opts); grid.touchMeta(); grid.invalidate(); grid.updateSpacer(); grid.draw(); refresh(); shell.markDirty(); };
const styleSel = (patch) => act('Formato', () => wb.styleRange(clampToUsed(sel()), patch), { hf: false });
// Al seleccionar columnas/filas enteras se limita al área usada (+ margen) para no crear millones de estilos
function clampToUsed(g) {
  const s = wb.sheet;
  const u = wb.usedRange();
  const out = { ...g };
  if (g.r2 - g.r1 > 5000) out.r2 = Math.min(g.r2, Math.max(u.r2 + 50, g.r1 + 100));
  if (g.c2 - g.c1 > 200) out.c2 = Math.min(g.c2, Math.max(u.c2 + 10, g.c1 + 26));
  void s;
  return out;
}
const curStyle = () => wb.sheet.style(grid.sel.r, grid.sel.c) || {};
const toggleStyle = (k) => { const on = !curStyle()[k]; styleSel({ [k]: on || undefined }); };

// ------------------------------------------------------------------ espacio de trabajo
function buildWorkspace(ws) {
  ws.classList.add('calc');
  nameBox = h('input', { class: 'namebox', 'aria-label': 'Cuadro de nombres', title: 'Escribe una celda o rango (p. ej. B5 o A1:C10) y pulsa Enter' });
  fxInput = h('textarea', { class: 'fx-input', rows: 1, spellcheck: 'false', 'aria-label': 'Barra de fórmulas' });
  const fxBtn = h('button', { class: 'fx-btn', title: 'Insertar función (Mayús+F3)', html: '<i>fx</i>', onclick: () => insertFunction() });
  const bar = h('div', { class: 'formulabar' }, nameBox, h('div', { class: 'fx-sep' }), fxBtn, fxInput);
  const host = h('div', { class: 'gridhost' });
  tabsBar = h('div', { class: 'sheettabs' });
  ws.append(bar, host, tabsBar, h('div', { id: 'printRoot' }));
  grid = new Grid(host, wb, {
    accent: '#0E9C74',
    onSelect: () => refresh(),
    onEdit: (v) => { if (document.activeElement !== fxInput) fxInput.value = v; },
    onChange: () => { shell.markDirty(); refreshCharts(); refresh(); },
    onContext: (e, kind) => contextMenu(e, kind),
  });
  grid.onClear = () => { if (!F.canEdit(grid.sel.r, grid.sel.c)) { toast('La hoja está protegida.', { error: true }); return; } clearContents(); };
  grid.cfStyle = (r, c) => F.cfStyle(r, c);
  grid.validationAt = (r, c) => F.validationAt(r, c);
  grid.validate = (r, c, v) => F.validate(r, c, v);
  grid.canEdit = (r, c) => F.canEdit(r, c);
  grid.onListButton = (r, c, pos) => { const v = F.validationAt(r, c); const anchor = h('div', { style: { position: 'fixed', left: `${pos.left - 160}px`, top: `${pos.top + 8}px`, width: '1px', height: '1px' } }); document.body.append(anchor); dropdown(anchor, v.list.map((x) => ({ label: x, run: () => { act('Elegir de la lista', () => wb.setInput(r, c, x)); refreshCharts(); } }))); anchor.remove(); };
  grid.onFilterButton = (c, pos) => filterMenu(c, pos);
  grid.onZoomWheel = (d) => zoom.set(zoom.get() + d);
  nameBox.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter') return;
    const g = parseRange(nameBox.value);
    if (g) { grid.selectRange(g); grid.canvas.focus(); } else toast('Escribe una celda (B5) o un rango (A1:C10).', { error: true });
  });
  nameBox.addEventListener('focus', () => nameBox.select());
  fxInput.addEventListener('focus', () => { if (!grid.editing) { grid.startEdit(null); fxInput.focus(); fxInput.value = grid.editor.value; } });
  fxInput.addEventListener('input', () => { if (grid.editing) { grid.editor.value = fxInput.value; grid.positionEditor(); } });
  fxInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); grid.editor.value = fxInput.value; grid.commitEdit([1, 0]); }
    if (e.key === 'Escape') { e.preventDefault(); grid.cancelEdit(); }
  });
}

function refresh() {
  if (!grid) return;
  const g = sel(), { r, c } = grid.sel;
  nameBox.value = document.activeElement === nameBox ? nameBox.value : (g.r1 === g.r2 && g.c1 === g.c2 ? addr(r, c) : (g.r1 === 0 && g.r2 >= wb.sheet.nRows - 1 ? `${colName(g.c1)}:${colName(g.c2)}` : g.c1 === 0 && g.c2 >= wb.sheet.nCols - 1 ? `${g.r1 + 1}:${g.r2 + 1}` : rangeStr(g)));
  if (!grid.editing) fxInput.value = wb.sheet.protected && wb.sheet.style(r, c)?.hidef ? '' : wb.editText(r, c);
  // lector de pantalla: celda activa y su contenido
  if (document.activeElement === grid.canvas) { const key = `${wb.active}:${r}:${c}`; const txt = wb.display(r, c); if (key !== refresh.last || txt !== refresh.lastTxt) { refresh.last = key; refresh.lastTxt = txt; announce(`${addr(r, c)}${txt !== '' ? `, ${txt}` : ''}${wb.editText(r, c).startsWith('=') ? `, ${t('fórmula')} ${wb.editText(r, c)}` : ''}`); } }
  // estado de la cinta
  const st = curStyle(), rb = shell.ribbon;
  rb.setOn('bold', st.b); rb.setOn('italic', st.i); rb.setOn('underline', st.u); rb.setOn('strike', st.s);
  rb.setOn('alignLeft', st.align === 'left'); rb.setOn('alignCenter', st.align === 'center'); rb.setOn('alignRight', st.align === 'right');
  rb.setOn('vTop', st.valign === 'top'); rb.setOn('vMiddle', st.valign === 'middle'); rb.setOn('vBottom', !st.valign || st.valign === 'bottom');
  rb.setOn('wrap', st.wrap); rb.setOn('filter', !!wb.sheet.filter);
  rb.setOn('merge', !!wb.sheet.mergeAt(r, c));
  rb.setValue('fontName', st.font || 'Calibri'); rb.setValue('fontSize', String(st.size || 11));
  rb.setValue('numFmt', FORMATS.some((f) => f[0] === (st.fmt || 'General')) ? (st.fmt || 'General') : 'custom');
  // estadísticas de la selección
  const cg = clampToUsed(g);
  let sum = 0, cnt = 0, nums = 0;
  if ((cg.r2 - cg.r1 + 1) * (cg.c2 - cg.c1 + 1) <= 200000 && (cg.r1 !== cg.r2 || cg.c1 !== cg.c2)) {
    for (let rr = cg.r1; rr <= cg.r2; rr++) { if (wb.sheet.hiddenRows.has(rr)) continue; for (let cc = cg.c1; cc <= cg.c2; cc++) { const v = wb.value(rr, cc); if (v === null || v === '' ) continue; cnt++; if (typeof v === 'number') { sum += v; nums++; } } }
  }
  const f = (n) => formatValue(n, st.fmt && st.fmt !== '@' ? st.fmt : '#,##0.########').text;
  statusInfo.textContent = cnt > 1 ? `${nums ? `Promedio: ${f(sum / nums)}   Recuento: ${cnt}   Suma: ${f(sum)}` : `Recuento: ${cnt}`}` : '';
  renderTabs();
}

// ------------------------------------------------------------------ hojas
function renderTabs() {
  const key = wb.sheets.map((s) => `${s.name}${s.x3?.tabColor || ''}${s.x3?.hidden ? '~' : ''}`).join('|') + wb.active;
  if (tabsBar.dataset.key === key) return;
  tabsBar.dataset.key = key;
  tabsBar.replaceChildren(h('button', { class: 'tabadd', title: 'Hoja nueva (Mayús+F11)', html: icon('plus', 16), onclick: () => newSheet() }));
  wb.sheets.forEach((s, i) => {
    if (s.x3?.hidden) return;
    const t = h('button', { class: `stab${i === wb.active ? ' on' : ''}`, style: s.x3?.tabColor ? `--tabc:${s.x3.tabColor}` : null, onclick: () => switchSheet(i), ondblclick: () => renameSheet(i), oncontextmenu: (e) => { e.preventDefault(); switchSheet(i); dropdown(t, [{ label: 'Insertar hoja', icon: 'sheetAdd', run: newSheet }, { label: 'Cambiar nombre', icon: 'textbox', run: () => renameSheet(i) }, { label: 'Duplicar', icon: 'copy', run: () => duplicateSheet(i) }, { label: 'Mover o copiar…', icon: 'copy', run: () => DT.moveCopyDialog(i) }, { label: 'Color de pestaña…', icon: 'fill', run: () => setTimeout(() => CE.tabColorPicker(t, i), 0) }, '-', { label: 'Ocultar', icon: 'minus', run: () => DT.hideSheet(i) }, { label: 'Mostrar…', icon: 'plus', run: () => DT.showSheetDialog() }, '-', { label: 'Eliminar', icon: 'trash', run: () => deleteSheet(i) }]); } }, s.name);
    tabsBar.append(t);
  });
}
function switchSheet(i) {
  if (grid.editing) grid.commitEdit(null);
  wb.active = i; grid.scrollX = grid.scrollY = 0; grid.invalidate(); grid.touchMeta(); grid.updateSpacer(); grid.syncScroll();
  grid.select(0, 0); mountCharts(); refresh();
}
function bookLocked() { if (wb.protectedStructure) { toast('La estructura del libro está protegida (Revisar → Proteger libro).', { error: true }); return true; } return false; }
function newSheet() { if (bookLocked()) return; wb.addSheet(); switchSheet(wb.sheets.length - 1); shell.markDirty(); }
async function renameSheet(i) {
  if (bookLocked()) return;
  const v = await form('Cambiar nombre de la hoja', [{ name: 'n', label: 'Nombre', value: wb.sheets[i].name }]);
  if (!v || !v.n.trim()) return;
  const name = v.n.trim().replace(/[\\/?*[\]:]/g, '').slice(0, 31);
  if (wb.sheets.some((s, j) => j !== i && s.name.toLowerCase() === name.toLowerCase())) { toast('Ya existe una hoja con ese nombre.', { error: true }); return; }
  wb.hf.renameSheet(wb.sheets[i].id, name); wb.sheets[i].name = name;
  tabsBar.dataset.key = ''; refresh(); shell.markDirty();
}
async function deleteSheet(i) {
  if (bookLocked()) return;
  if (wb.sheets.length === 1) { toast('El libro necesita al menos una hoja.', { error: true }); return; }
  if (!(await confirmBox('Eliminar hoja', `¿Eliminar "${wb.sheets[i].name}"? No se puede deshacer.`, 'Eliminar'))) return;
  wb.hf.removeSheet(wb.sheets[i].id); wb.sheets.splice(i, 1);
  wb.undoStack = []; wb.redoStack = []; wb.hf.clearUndoStack?.();
  switchSheet(Math.max(0, i - 1)); shell.markDirty();
}
function duplicateSheet(i) {
  const src = wb.sheets[i];
  const s = wb.addSheet(`${src.name} (2)`);
  wb.hf.setSheetContent(s.id, wb.hf.getSheetSerialized(src.id));
  s.setMeta(src.meta()); s.name = wb.hf.getSheetName(s.id); s.charts = src.charts.map((c) => ({ ...c, id: uid() }));
  switchSheet(wb.sheets.length - 1); shell.markDirty();
}

// ------------------------------------------------------------------ acciones de celdas
function clearContents() {
  const g = clampToUsed(sel());
  act('Borrar', () => { for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if (wb.raw(r, c) != null) wb.setRaw(r, c, null); });
  refreshCharts();
}
function clearAll() { const g = clampToUsed(sel()); act('Borrar todo', () => { for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { if (wb.raw(r, c) != null) wb.setRaw(r, c, null); wb.sheet.styles.delete(`${r},${c}`); } }); }
function clearFormats() { const g = clampToUsed(sel()); act('Borrar formato', () => { for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) wb.sheet.styles.delete(`${r},${c}`); }, { hf: false }); }

function insertRows(above = true) { const g = sel(); const n = g.r2 - g.r1 + 1; act('Insertar filas', () => wb.insertRows(above ? g.r1 : g.r2 + 1, Math.min(n, 1000))); }
function insertCols(left = true) { const g = sel(); const n = g.c2 - g.c1 + 1; act('Insertar columnas', () => wb.insertCols(left ? g.c1 : g.c2 + 1, Math.min(n, 100))); }
function deleteRows() { const g = sel(); act('Eliminar filas', () => wb.deleteRows(g.r1, Math.min(g.r2 - g.r1 + 1, wb.sheet.nRows - 1))); grid.select(g.r1, g.c1); }
function deleteCols() { const g = sel(); act('Eliminar columnas', () => wb.deleteCols(g.c1, Math.min(g.c2 - g.c1 + 1, wb.sheet.nCols - 1))); grid.select(g.r1, g.c1); }

function toggleMerge() {
  const g = sel(), s = wb.sheet;
  const existing = s.mergeAt(g.r1, g.c1);
  if (existing && existing.r1 === g.r1 && existing.c1 === g.c1 && existing.r2 === g.r2 && existing.c2 === g.c2) {
    act('Separar celdas', () => { s.merges = s.merges.filter((m) => m !== existing); }, { hf: false }); return;
  }
  if (g.r1 === g.r2 && g.c1 === g.c2) return;
  let others = 0;
  for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if ((r !== g.r1 || c !== g.c1) && wb.raw(r, c) != null) others++;
  const doIt = () => act('Combinar y centrar', () => {
    s.merges = s.merges.filter((m) => m.r2 < g.r1 || m.r1 > g.r2 || m.c2 < g.c1 || m.c1 > g.c2);
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if ((r !== g.r1 || c !== g.c1) && wb.raw(r, c) != null) wb.setRaw(r, c, null);
    s.merges.push({ ...g });
    wb.setStyle(g.r1, g.c1, { align: 'center', valign: 'middle' });
  });
  if (others) confirmBox('Combinar celdas', 'Al combinar solo se conserva el valor de la celda superior izquierda.', 'Combinar').then((ok) => ok && doIt()); else doIt();
}

function setBorders(kind) {
  const g = clampToUsed(sel()), col = '#000000';
  act('Bordes', () => {
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      const cur = { ...(wb.sheet.style(r, c)?.border || {}) };
      let b;
      if (kind === 'none') b = undefined;
      else if (kind === 'all') b = { t: col, b: col, l: col, r: col };
      else if (kind === 'outer') { b = cur; if (r === g.r1) b.t = col; if (r === g.r2) b.b = col; if (c === g.c1) b.l = col; if (c === g.c2) b.r = col; }
      else if (kind === 'bottom') { b = cur; if (r === g.r2) b.b = col; }
      else if (kind === 'top') { b = cur; if (r === g.r1) b.t = col; }
      else if (kind === 'thickBottom') { b = cur; if (r === g.r2) b.b = '#000000'; }
      wb.setStyle(r, c, { border: b && Object.keys(b).length ? b : undefined });
    }
  }, { hf: false });
}

function sortSel(asc) {
  const g = sel();
  const region = g.r1 === g.r2 && g.c1 === g.c2 ? wb.region(g.r1, g.c1) : clampToUsed(g);
  if (region.r1 === region.r2) return;
  const v0 = wb.value(region.r1, grid.sel.c), v1 = wb.value(region.r1 + 1, grid.sel.c);
  const hasHeader = typeof v0 === 'string' && (typeof v1 !== 'string' || [...Array(region.c2 - region.c1 + 1).keys()].every((i) => typeof wb.value(region.r1, region.c1 + i) === 'string'));
  act(asc ? 'Ordenar A-Z' : 'Ordenar Z-A', () => wb.sortRange(region, grid.sel.c, asc, hasHeader));
  refreshCharts();
}

function toggleFilter() {
  const s = wb.sheet;
  if (s.filter) { act('Quitar filtro', () => { s.filter = null; s.hiddenRows = new Set(); }, { hf: false }); return; }
  const g = sel();
  const region = g.r1 === g.r2 ? wb.region(g.r1, g.c1) : clampToUsed(g);
  if (region.r1 === region.r2) { toast('Selecciona una tabla con encabezados para filtrar.', { error: true }); return; }
  act('Filtro', () => { s.filter = { range: region, hidden: new Map() }; }, { hf: false });
}
function applyFilter() {
  const s = wb.sheet, f = s.filter;
  s.hiddenRows = new Set();
  for (let r = f.range.r1 + 1; r <= f.range.r2; r++) {
    for (const [c, hidden] of f.hidden) { if (hidden.has(wb.display(r, c).text)) { s.hiddenRows.add(r); break; } }
  }
  for (const r of F3.outlineHiddenRows(s)) s.hiddenRows.add(r);
}
function filterMenu(c, pos) {
  const s = wb.sheet, f = s.filter;
  const values = new Map();
  for (let r = f.range.r1 + 1; r <= f.range.r2; r++) { const t = wb.display(r, c).text; values.set(t, (values.get(t) || 0) + 1); }
  const hidden = new Set(f.hidden.get(c) || []);
  const list = h('div', { style: { maxHeight: '240px', overflow: 'auto', padding: '4px 8px' } });
  const sorted = [...values.keys()].sort((a, b) => (a === '' ? 1 : b === '' ? -1 : a.localeCompare(b, 'es', { numeric: true })));
  const all = h('input', { type: 'checkbox', checked: hidden.size === 0 ? true : null });
  all.addEventListener('change', () => { list.querySelectorAll('input').forEach((i) => { i.checked = all.checked; }); });
  sorted.forEach((v) => list.append(h('label', { class: 'check', style: { margin: '2px 0' } }, h('input', { type: 'checkbox', value: v, checked: hidden.has(v) ? null : true }), v === '' ? '(Vacías)' : v)));
  const box = h('div', { style: { width: '240px' } },
    h('button', { class: 'mi', html: `${icon('sortAsc', 18)}<span>Ordenar de A a Z</span>`, onclick: () => { closeDropdown(); grid.sel.c = c; act('Ordenar', () => wb.sortRange(f.range, c, true, true)); applyFilterRefresh(); } }),
    h('button', { class: 'mi', html: `${icon('sortDesc', 18)}<span>Ordenar de Z a A</span>`, onclick: () => { closeDropdown(); act('Ordenar', () => wb.sortRange(f.range, c, false, true)); applyFilterRefresh(); } }),
    h('hr'), h('label', { class: 'check', style: { padding: '0 8px', fontWeight: 600 } }, all, 'Seleccionar todo'), list,
    h('div', { style: { display: 'flex', gap: '6px', justifyContent: 'flex-end', padding: '6px' } },
      h('button', { class: 'btn', onclick: () => closeDropdown() }, 'Cancelar'),
      h('button', { class: 'btn primary', onclick: () => {
        const hid = new Set([...list.querySelectorAll('input')].filter((i) => !i.checked).map((i) => i.value));
        closeDropdown();
        act('Filtrar', () => { if (hid.size) f.hidden.set(c, hid); else f.hidden.delete(c); applyFilter(); }, { hf: false });
      } }, 'Aceptar')));
  const anchor = h('div', { style: { position: 'fixed', left: `${pos.left}px`, top: `${pos.top}px`, width: '1px', height: '1px' } });
  document.body.append(anchor); dropdown(anchor, box); anchor.remove();
}
function applyFilterRefresh() { if (wb.sheet.filter) { applyFilter(); grid.touchMeta(); grid.invalidate(); grid.draw(); } }

function freeze(kind) {
  const s = wb.sheet;
  const f = kind === 'none' ? { rows: 0, cols: 0 } : kind === 'row' ? { rows: 1, cols: 0 } : kind === 'col' ? { rows: 0, cols: 1 } : { rows: grid.sel.r, cols: grid.sel.c };
  act('Inmovilizar', () => { s.freeze = f; }, { hf: false });
  grid.scrollX = grid.scrollY = 0; grid.syncScroll(); grid.draw();
}

function autoSum(fn = 'SUM') {
  const g = sel(), es = toSpanish(`=${fn}(`).slice(1, -1);
  if (g.r1 === g.r2 && g.c1 === g.c2) {
    // Busca números arriba (o a la izquierda) de la celda activa
    let { r, c } = grid.sel; let r0 = r - 1;
    while (r0 >= 0 && typeof wb.value(r0, c) === 'number') r0--;
    if (r0 < r - 1) { grid.startEdit(`=${es}(${addr(r0 + 1, c)}:${addr(r - 1, c)})`); return; }
    let c0 = c - 1; while (c0 >= 0 && typeof wb.value(r, c0) === 'number') c0--;
    if (c0 < c - 1) { grid.startEdit(`=${es}(${addr(r, c0 + 1)}:${addr(r, c - 1)})`); return; }
    grid.startEdit(`=${es}()`); grid.editor.setSelectionRange(grid.editor.value.length - 1, grid.editor.value.length - 1); return;
  }
  // con rango: total debajo de cada columna
  act('Autosuma', () => { for (let c = g.c1; c <= g.c2; c++) wb.setInput(g.r2 + 1, c, `=${es}(${addr(g.r1, c)}:${addr(g.r2, c)})`); });
  refreshCharts();
}

// ------------------------------------------------------------------ diálogos
async function insertFunction() {
  const cats = ['Todas', ...new Set(FUNCTION_HELP.map((f) => f.cat))];
  const q = h('input', { placeholder: 'Busca una función, p. ej. “suma” o “buscar”' });
  const cat = h('select', {}, ...cats.map((c) => h('option', { value: c }, c)));
  const list = h('div', { class: 'fnlist' });
  const info = h('div', { class: 'fninfo' });
  let chosen = null;
  const render = () => {
    const nq = q.value.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
    list.replaceChildren();
    FUNCTION_HELP.filter((f) => (cat.value === 'Todas' || f.cat === cat.value) && (!nq || `${f.es} ${f.en} ${f.desc}`.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').includes(nq))).forEach((f, i) => {
      const it = h('button', { class: 'fnitem', onclick: () => pick(f, it), ondblclick: () => { pick(f, it); ok(); } }, h('b', {}, f.es), h('small', {}, f.cat));
      list.append(it);
      if (i === 0) pick(f, it);
    });
  };
  const pick = (f, it) => { chosen = f; list.querySelectorAll('.on').forEach((x) => x.classList.remove('on')); it.classList.add('on'); info.innerHTML = `<b>${esc(f.syntax)}</b><p>${esc(f.desc)}</p>`; };
  let closeFn = null;
  const ok = () => { if (chosen) { closeFn?.(true); } };
  q.addEventListener('input', render); cat.addEventListener('change', render);
  const body = h('div', { class: 'fndialog' }, h('div', { class: 'row' }, q, cat), list, info);
  render();
  const res = await modal({ title: 'Insertar función', body, width: '560px', buttons: [{ label: 'Cancelar', value: false }, { label: 'Insertar', value: true, primary: true }], onOpen: (_c, close) => { closeFn = close; } });
  if (res && chosen) { grid.select(grid.sel.r, grid.sel.c); grid.startEdit(`=${chosen.es}()`); grid.editor.setSelectionRange(grid.editor.value.length - 1, grid.editor.value.length - 1); }
}

async function formatCells(tab = 0) { return CE.formatCellsDialog(tab); }
async function formatCellsOld() {
  const st = curStyle();
  const fmtSel = h('select', { name: 'fmt' }, ...FORMATS.map(([k, n]) => h('option', { value: k, selected: k === (st.fmt || 'General') ? true : null }, n)), h('option', { value: 'custom', selected: !FORMATS.some((f) => f[0] === (st.fmt || 'General')) ? true : null }, 'Personalizado…'));
  const custom = h('input', { name: 'custom', value: st.fmt || 'General' });
  const preview = h('div', { class: 'fmtprev' });
  const sample = wb.value(grid.sel.r, grid.sel.c);
  const upd = () => { const f = fmtSel.value === 'custom' ? custom.value : fmtSel.value; if (fmtSel.value !== 'custom') custom.value = f; preview.textContent = `Ejemplo: ${formatValue(typeof sample === 'number' ? sample : 1234.5678, f).text}`; };
  fmtSel.addEventListener('change', upd); custom.addEventListener('input', () => { fmtSel.value = 'custom'; upd(); }); upd();
  const al = h('select', { name: 'align' }, ...[['', 'General'], ['left', 'Izquierda'], ['center', 'Centro'], ['right', 'Derecha']].map(([v, l]) => h('option', { value: v, selected: (st.align || '') === v ? true : null }, l)));
  const va = h('select', { name: 'valign' }, ...[['bottom', 'Inferior'], ['middle', 'Centro'], ['top', 'Superior']].map(([v, l]) => h('option', { value: v, selected: (st.valign || 'bottom') === v ? true : null }, l)));
  const wrap = h('input', { type: 'checkbox', name: 'wrap', checked: st.wrap ? true : null });
  const body = h('div', { style: { width: '420px' } },
    h('label', { class: 'field' }, h('span', {}, 'Categoría'), fmtSel), h('label', { class: 'field' }, h('span', {}, 'Código de formato'), custom), preview,
    h('div', { class: 'fields2' }, h('label', { class: 'field' }, h('span', {}, 'Alineación horizontal'), al), h('label', { class: 'field' }, h('span', {}, 'Alineación vertical'), va)),
    h('label', { class: 'check' }, wrap, 'Ajustar texto'));
  const ok = await modal({ title: 'Formato de celdas', body, buttons: [{ label: 'Cancelar', value: false }, { label: 'Aceptar', value: true, primary: true }] });
  if (!ok) return;
  const f = fmtSel.value === 'custom' ? custom.value : fmtSel.value;
  styleSel({ fmt: f === 'General' ? undefined : f, align: al.value || undefined, valign: va.value === 'bottom' ? undefined : va.value, wrap: wrap.checked || undefined });
}

function contextMenu(e, kind) {
  const anchor = h('div', { style: { position: 'fixed', left: `${e.clientX}px`, top: `${e.clientY}px`, width: '1px', height: '1px' } });
  document.body.append(anchor);
  const items = [
    { label: 'Cortar', icon: 'cut', kbd: 'Ctrl+X', run: () => { grid.canvas.focus(); document.execCommand('cut'); } },
    { label: 'Copiar', icon: 'copy', kbd: 'Ctrl+C', run: () => { grid.canvas.focus(); document.execCommand('copy'); } },
    { label: 'Pegar', icon: 'paste', kbd: 'Ctrl+V', run: () => pasteFromClipboard() },
    '-',
    ...(kind !== 'col' ? [{ label: 'Insertar fila arriba', icon: 'rowAdd', run: () => insertRows(true) }, { label: 'Insertar fila abajo', icon: 'rowAdd', run: () => insertRows(false) }] : []),
    ...(kind !== 'row' ? [{ label: 'Insertar columna a la izquierda', icon: 'colAdd', run: () => insertCols(true) }, { label: 'Insertar columna a la derecha', icon: 'colAdd', run: () => insertCols(false) }] : []),
    ...(kind !== 'col' ? [{ label: 'Eliminar filas', icon: 'rowDel', run: deleteRows }] : []),
    ...(kind !== 'row' ? [{ label: 'Eliminar columnas', icon: 'colDel', run: deleteCols }] : []),
    '-',
    { label: 'Borrar contenido', icon: 'clear', kbd: 'Supr', run: clearContents },
    { label: 'Ordenar de A a Z', icon: 'sortAsc', run: () => sortSel(true) },
    { label: 'Formato de celdas…', icon: 'brush', kbd: 'Ctrl+1', run: formatCells },
  ];
  dropdown(anchor, items); anchor.remove();
}

async function pasteFromClipboard() {
  try { const t = await navigator.clipboard.readText(); grid.paste(t); } catch { toast('Usa Ctrl+V para pegar.'); }
}

// ------------------------------------------------------------------ gráficos
const uid = () => Math.random().toString(36).slice(2, 10);
const CHART_TYPES = [['column', 'Columnas'], ['bar', 'Barras'], ['line', 'Líneas'], ['area', 'Área'], ['pie', 'Circular'], ['scatter', 'Dispersión']];
const PALETTE = ['#0E9C74', '#5B48E8', '#DC3A63', '#F08A24', '#1F8FD6', '#8E6CD9', '#2FB4A8', '#E0B000'];
const charts = new Map(); // id -> instancia echarts

function chartData(ch, s = wb.sheet) {
  const g = F3.resolveRange(ch.range, s); if (!g) return null;
  const val = (r, c) => wb.value(r, c, s);
  const headerRow = [...Array(g.c2 - g.c1 + 1).keys()].some((i) => typeof val(g.r1, g.c1 + i) === 'string') && g.r2 > g.r1;
  const catCol = [...Array(g.r2 - g.r1 + 1).keys()].slice(headerRow ? 1 : 0).some((i) => typeof val(g.r1 + i, g.c1) === 'string') || (g.c2 > g.c1 && headerRow && typeof val(g.r1, g.c1) !== 'number' && val(g.r1, g.c1) == null);
  const r0 = headerRow ? g.r1 + 1 : g.r1, c0 = catCol ? g.c1 + 1 : g.c1;
  const cats = []; for (let r = r0; r <= g.r2; r++) { if (s.hiddenRows.has(r)) continue; cats.push(catCol ? wb.display(r, g.c1, s).text : String(r - r0 + 1)); }
  const series = [];
  for (let c = c0; c <= g.c2; c++) {
    const data = []; for (let r = r0; r <= g.r2; r++) { if (s.hiddenRows.has(r)) continue; const v = val(r, c); data.push(typeof v === 'number' ? v : null); }
    series.push({ name: headerRow ? wb.display(g.r1, c, s).text : `Serie ${c - c0 + 1}`, data });
  }
  return { cats, series };
}
function chartOption(ch, s) {
  const d = chartData(ch, s) || { cats: [], series: [] };
  const base = { color: PALETTE, animation: !document.body.classList.contains('printing'), title: ch.title ? { text: ch.title, left: 'center', top: 8, textStyle: { fontSize: 15, fontWeight: 600, color: '#1D2030', fontFamily: 'Inter, sans-serif' } } : undefined, tooltip: { trigger: ch.type === 'pie' ? 'item' : 'axis' }, legend: d.series.length > 1 || ch.type === 'pie' ? { bottom: 4, type: 'scroll' } : undefined, grid: { left: 50, right: 20, top: ch.title ? 48 : 20, bottom: d.series.length > 1 ? 44 : 28, containLabel: true } };
  if (ch.type === 'pie') return { ...base, series: [{ type: 'pie', radius: ['0%', '62%'], center: ['50%', '52%'], data: d.cats.map((c, i) => ({ name: c, value: d.series[0]?.data[i] })), label: { formatter: '{b}: {d}%' } }] };
  if (ch.type === 'scatter') return { ...base, xAxis: { type: 'value' }, yAxis: { type: 'value' }, series: d.series.slice(1).map((sr) => ({ type: 'scatter', name: sr.name, data: sr.data.map((v, i) => [d.series[0].data[i], v]) })) };
  const horizontal = ch.type === 'bar';
  const cat = { type: 'category', data: d.cats, axisLabel: { color: '#4B5066' } }, valAx = { type: 'value', axisLabel: { color: '#4B5066' }, splitLine: { lineStyle: { color: '#ECEDF2' } } };
  return { ...base, xAxis: horizontal ? valAx : cat, yAxis: horizontal ? cat : valAx, series: d.series.map((sr) => ({ type: ch.type === 'column' || ch.type === 'bar' ? 'bar' : 'line', name: sr.name, data: sr.data, smooth: false, areaStyle: ch.type === 'area' ? { opacity: 0.25 } : undefined, barMaxWidth: 48 })) };
}
function mountCharts() {
  charts.forEach((c) => c.dispose()); charts.clear();
  grid.chartLayer.replaceChildren();
  for (const ch of wb.sheet.charts) mountChart(ch);
  grid.positionCharts();
  F3.renderSlicers();
}
function mountChart(ch) {
  const el = h('div', { class: 'chartbox', 'data-id': ch.id, style: { width: `${ch.w}px`, height: `${ch.h}px` }, tabindex: 0 });
  const body = h('div', { class: 'chartbody' });
  const handle = h('div', { class: 'chartresize' });
  el.append(body, handle);
  grid.chartLayer.append(el);
  const inst = echarts.init(body, null, { renderer: 'canvas' });
  inst.setOption(chartOption(ch));
  charts.set(ch.id, inst);
  el.addEventListener('mousedown', (e) => {
    $$('.chartbox.sel').forEach((x) => x.classList.remove('sel')); el.classList.add('sel'); el.focus();
    const resizing = e.target === handle;
    const x0 = e.clientX, y0 = e.clientY, start = { ...ch };
    const before = wb.sheet.meta();
    const move = (ev) => {
      const dx = (ev.clientX - x0) / grid.zoom, dy = (ev.clientY - y0) / grid.zoom;
      if (resizing) { ch.w = Math.max(220, start.w + dx); ch.h = Math.max(160, start.h + dy); el.style.width = `${ch.w}px`; el.style.height = `${ch.h}px`; inst.resize(); }
      else { ch.x = Math.max(0, start.x + dx); ch.y = Math.max(0, start.y + dy); }
      grid.positionCharts();
    };
    const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); if (ch.x !== start.x || ch.y !== start.y || ch.w !== start.w || ch.h !== start.h) { wb.undoStack.push({ label: 'Mover gráfico', hf: false, before: [[wb.sheet, before]], after: [[wb.sheet, wb.sheet.meta()]] }); shell.markDirty(); } };
    addEventListener('mousemove', move); addEventListener('mouseup', up);
    e.stopPropagation();
  });
  el.addEventListener('dblclick', () => editChart(ch));
  el.addEventListener('keydown', (e) => { if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); act('Eliminar gráfico', () => { wb.sheet.charts = wb.sheet.charts.filter((x) => x.id !== ch.id); }, { hf: false }); mountCharts(); } });
}
function refreshCharts() {
  clearTimeout(refreshCharts.t);
  refreshCharts.t = setTimeout(() => {
    const ids = new Set(wb.sheet.charts.map((c) => c.id));
    if ([...charts.keys()].some((k) => !ids.has(k)) || wb.sheet.charts.some((c) => !charts.has(c.id))) { mountCharts(); return; }
    for (const ch of wb.sheet.charts) charts.get(ch.id)?.setOption(chartOption(ch), true);
  }, 120);
}
async function insertChart(type) {
  let g = sel();
  if (g.r1 === g.r2 && g.c1 === g.c2) g = wb.region(g.r1, g.c1);
  if (g.r1 === g.r2 && g.c1 === g.c2) { toast('Selecciona primero los datos que quieres graficar.', { error: true }); return; }
  g = clampToUsed(g);
  const p = grid.prefix();
  const ch = { id: uid(), type: type || 'column', range: rangeStr(g), title: '', x: p.cx[g.c2 + 1] + 24, y: p.ry[g.r1], w: 480, h: 300 };
  if (!type) { const v = await chartDialog(ch, 'Insertar gráfico'); if (!v) return; Object.assign(ch, v); }
  act('Insertar gráfico', () => { wb.sheet.charts.push(ch); }, { hf: false });
  mountCharts();
}
async function chartDialog(ch, title) {
  return form(title, [
    { name: 'type', label: 'Tipo de gráfico', type: 'select', value: ch.type, options: CHART_TYPES },
    { name: 'title', label: 'Título', value: ch.title, placeholder: 'Opcional' },
    { name: 'range', label: 'Datos (rango)', value: ch.range },
  ], 'Aceptar');
}
async function editChart(ch) {
  const v = await chartDialog(ch, 'Editar gráfico');
  if (!v) return;
  if (!F3.resolveRange(v.range)) { toast('Rango no válido', { error: true }); return; }
  act('Editar gráfico', () => Object.assign(ch, v), { hf: false });
  refreshCharts();
}

// ------------------------------------------------------------------ impresión: tabla HTML con el área usada
function buildPrint() {
  const root = $('#printRoot'); root.replaceChildren();
  const s = wb.sheet, u = wb.usedRange();
  const p = grid.prefix();
  const ps = s.pageSetup || {};
  const area = s.printArea ? parseRange(s.printArea) : null;
  let minR = area ? area.r1 : 0, minC = area ? area.c1 : 0;
  let maxR = area ? area.r2 : u.r2, maxC = area ? area.c2 : u.c2;
  if (!area) for (const ch of s.charts) { maxC = Math.max(maxC, grid.bsearch(p.cx, ch.x + ch.w)); maxR = Math.max(maxR, grid.bsearch(p.ry, ch.y + ch.h)); }
  const titleR = /^(\d+):(\d+)$/.exec(ps.titleRows || '');
  const table = h('table', { class: 'ptable' });
  if (ps.gridlines === false) table.classList.add('nogrid');
  const cg = h('colgroup'); for (let c = minC; c <= maxC; c++) cg.append(h('col', { style: { width: `${s.cw(c)}px` } }));
  table.append(cg);
  const covered = new Set();
  const thead = h('thead'); const tbody = h('tbody'); table.append(thead, tbody);
  for (let r = minR; r <= maxR; r++) {
    if (s.hiddenRows.has(r)) continue;
    const tr = h('tr', { style: { height: `${s.rh(r)}px` } });
    for (let c = minC; c <= maxC; c++) {
      if (covered.has(`${r},${c}`)) continue;
      const m = s.mergeAt(r, c);
      const st = grid.eff(r, c) || {};
      const d = wb.display(r, c);
      const td = h('td', {}, d.text);
      if (st.bar) td.style.backgroundImage = `linear-gradient(90deg, ${st.bar.color} ${Math.round(st.bar.pct * 100)}%, transparent ${Math.round(st.bar.pct * 100)}%)`;
      if (m) { td.rowSpan = m.r2 - m.r1 + 1; td.colSpan = m.c2 - m.c1 + 1; for (let rr = m.r1; rr <= m.r2; rr++) for (let cc = m.c1; cc <= m.c2; cc++) covered.add(`${rr},${cc}`); }
      Object.assign(td.style, {
        fontWeight: st.b ? '700' : '', fontStyle: st.i ? 'italic' : '', textDecoration: [st.u && 'underline', st.s && 'line-through'].filter(Boolean).join(' '),
        color: d.error ? '#C62828' : (d.color || st.color || ''), background: st.bg || '', textAlign: st.align || d.align || 'left', verticalAlign: { top: 'top', middle: 'middle' }[st.valign] || 'bottom',
        fontSize: `${st.size || 11}pt`, fontFamily: st.font ? `"${st.font}", Carlito, sans-serif` : '', whiteSpace: st.wrap ? 'pre-wrap' : 'nowrap',
        borderTop: st.border?.t ? `1px solid ${st.border.t}` : '', borderBottom: st.border?.b ? `1px solid ${st.border.b}` : '', borderLeft: st.border?.l ? `1px solid ${st.border.l}` : '', borderRight: st.border?.r ? `1px solid ${st.border.r}` : '',
      });
      tr.append(td);
    }
    (titleR && r >= Number(titleR[1]) - 1 && r <= Number(titleR[2]) - 1 ? thead : tbody).append(tr);
  }
  const wrap = h('div', { class: 'pwrap' }, table);
  for (const ch of s.charts) { const inst = charts.get(ch.id); if (!inst) continue; wrap.append(h('img', { src: inst.getDataURL({ pixelRatio: 2, backgroundColor: '#fff' }), style: { position: 'absolute', left: `${ch.x}px`, top: `${ch.y}px`, width: `${ch.w}px`, height: `${ch.h}px` } })); }
  root.append(h('div', { class: 'ptitle' }, `${shell.doc.name.replace(/\.[^.]+$/, '')} — ${s.name}`), wrap);
  const totalW = p.cx[maxC + 1] - p.cx[minC];
  const land = ps.orientation === 'landscape' || (ps.orientation !== 'portrait' && totalW > 760);
  root.dataset.landscape = land ? '1' : '';
  const pageW = (land ? 11 : 8.5) * 96 - 2 * 45, pageH = (land ? 8.5 : 11) * 96 - 2 * 57 - 30;
  if (ps.fit === 'width' || ps.fit === 'page') {
    const tableH = p.ry[maxR + 1] - p.ry[minR];
    const k = Math.min(1, pageW / totalW, ps.fit === 'page' ? pageH / tableH : 1);
    wrap.style.zoom = String(k);
  }
  let style = $('#pagestyle'); if (!style) { style = h('style', { id: 'pagestyle' }); document.head.append(style); }
  style.textContent = `@page { size: letter ${land ? 'landscape' : 'portrait'}; margin: 1.5cm 1.2cm; }`;
}

// ------------------------------------------------------------------ adaptador
function emptyBook(tpl) {
  // plantillas en el idioma de la interfaz
  const L = (rows) => rows.map((r) => r.map((v) => (typeof v === 'string' && v && !v.startsWith('=') ? t(v) : v)));
  wb.reset();
  const s = wb.addSheet();
  if (tpl === 'presupuesto') {
    const rows = [['Presupuesto mensual', '', '', ''], ['', '', '', ''], ['Concepto', 'Presupuesto', 'Real', 'Diferencia'], ['Renta', 8500, 8500, '=B4-C4'], ['Luz y agua', 900, 1040, '=B5-C5'], ['Internet y teléfono', 650, 650, '=B6-C6'], ['Despensa', 4500, 4820, '=B7-C7'], ['Transporte', 1600, 1350, '=B8-C8'], ['Ahorro', 2000, 2000, '=B9-C9'], ['Total', '=SUM(B4:B9)', '=SUM(C4:C9)', '=B10-C10']];
    wb.hf.setSheetContent(s.id, L(rows));
    s.merges.push({ r1: 0, c1: 0, r2: 0, c2: 3 });
    wb.setStyle(0, 0, { b: true, size: 16, color: '#0A6B50' }); s.rowH.set(0, 32);
    for (let c = 0; c < 4; c++) { wb.setStyle(2, c, { b: true, bg: '#DFF5EE', border: { b: '#0E9C74' } }); wb.setStyle(9, c, { b: true, border: { t: '#000000' } }); }
    for (let r = 3; r <= 9; r++) for (let c = 1; c < 4; c++) wb.setStyle(r, c, { ...(wb.sheet.style(r, c) || {}), fmt: '"$"#,##0.00;[Red]-"$"#,##0.00' });
    s.colW.set(0, 170); s.colW.set(1, 110); s.colW.set(2, 110); s.colW.set(3, 110);
    s.charts.push({ id: uid(), type: 'column', range: 'A3:C9', title: t('Presupuesto vs. real'), x: 530, y: 48, w: 460, h: 290 });
  } else if (tpl === 'ventas') {
    const meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio'];
    const rows = [['Mes', 'Ventas', 'Costos', 'Utilidad', 'Margen']];
    meses.forEach((m, i) => { const r = i + 2; rows.push([m, 42000 + i * 3500 + (i % 2) * 2100, 26000 + i * 1500, `=B${r}-C${r}`, `=D${r}/B${r}`]); });
    rows.push(['Total', '=SUM(B2:B7)', '=SUM(C2:C7)', '=SUM(D2:D7)', '=D8/B8']);
    wb.hf.setSheetContent(s.id, L(rows));
    for (let c = 0; c < 5; c++) { wb.setStyle(0, c, { b: true, color: '#FFFFFF', bg: '#0E9C74', align: c ? 'center' : undefined }); wb.setStyle(7, c, { b: true, border: { t: '#000000' } }); }
    for (let r = 1; r <= 7; r++) { for (let c = 1; c < 4; c++) wb.setStyle(r, c, { ...(wb.sheet.style(r, c) || {}), fmt: '"$"#,##0' }); wb.setStyle(r, 4, { ...(wb.sheet.style(r, 4) || {}), fmt: '0.0%' }); }
    s.colW.set(0, 110); s.freeze = { rows: 1, cols: 0 };
    s.charts.push({ id: uid(), type: 'line', range: 'A1:D7', title: t('Ventas, costos y utilidad'), x: 520, y: 30, w: 480, h: 300 });
  } else if (tpl === 'asistencia') {
    const rows = [['Lista de asistencia', '', '', '', '', '', ''], ['Nombre', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Asistencias']];
    ['Ana López', 'Bruno Díaz', 'Carla Ruiz', 'Diego Mora', 'Elena Paz'].forEach((n, i) => { const r = i + 3; rows.push([n, '✓', '✓', '', '✓', '✓', `=COUNTIF(B${r}:F${r},"✓")`]); });
    wb.hf.setSheetContent(s.id, L(rows));
    s.merges.push({ r1: 0, c1: 0, r2: 0, c2: 6 });
    wb.setStyle(0, 0, { b: true, size: 15, align: 'center' }); s.rowH.set(0, 30);
    for (let c = 0; c < 7; c++) wb.setStyle(1, c, { b: true, bg: '#DFF5EE', align: 'center', border: { t: '#9AA0B4', b: '#9AA0B4', l: '#9AA0B4', r: '#9AA0B4' } });
    for (let r = 2; r < 7; r++) for (let c = 0; c < 7; c++) wb.setStyle(r, c, { align: c ? 'center' : undefined, border: { t: '#9AA0B4', b: '#9AA0B4', l: '#9AA0B4', r: '#9AA0B4' } });
    s.colW.set(0, 160); for (let c = 1; c < 6; c++) s.colW.set(c, 56); s.colW.set(6, 100);
  }
  wb.hf.clearUndoStack?.(); wb.undoStack = []; wb.redoStack = [];
}

const blankThumb = '<svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.4"><path d="M12 5v14 M5 12h14"/></svg>';
const gridThumb = (head, accent) => `<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:1px;background:#E2E4EA;height:100%">${Array.from({ length: 30 }, (_, i) => `<div style="background:${i < 5 && head ? accent : '#fff'}"></div>`).join('')}</div>`;

const adapter = {
  i18nSkip: '.stab, .grid-cell-editor, .cnote-tip',
  exts: ['xlsx', 'xlsm', 'ods', 'csv', 'tsv'],
  writable: ['xlsx', 'xlsm', 'ods', 'csv'],
  saveExt: 'xlsx',
  docIcon: 'sheet',
  homeHint: 'Haz cuentas, presupuestos y reportes con fórmulas en español.',
  templates: [
    { id: 'blank', name: 'Libro en blanco', blank: true, thumb: blankThumb },
    { id: 'presupuesto', name: 'Presupuesto mensual', thumb: gridThumb(true, '#DFF5EE') },
    { id: 'ventas', name: 'Ventas por mes', thumb: gridThumb(true, '#0E9C74') },
    { id: 'asistencia', name: 'Lista de asistencia', thumb: gridThumb(true, '#B9E9D9') },
  ],
  exportFormats: [
    { ext: 'xlsx', ic: 'sheet', label: 'Libro de Excel (.xlsx)', desc: 'Compatible con otras hojas de cálculo.' },
    { ext: 'ods', ic: 'sheet', label: 'Hoja de cálculo ODF (.ods)', desc: 'Formato abierto, compatible con LibreOffice Calc.' },
    { ext: 'csv', ic: 'words', label: 'CSV (valores separados por comas)', desc: 'Solo la hoja actual, sin formato. Útil para importar en otros sistemas.' },
  ],
  async newDoc(tpl) { emptyBook(tpl); afterLoad(); },
  async load(f) {
    if (f.ext === 'csv' || f.ext === 'tsv') {
      const text = new TextDecoder('utf-8').decode(f.bytes);
      wb.reset(); const s = wb.addSheet(f.name.replace(/\.[^.]+$/, '').slice(0, 31) || 'Hoja1');
      const rows = parseCsv(text);
      // se interpreta cada valor y se carga todo de una vez (mucho más rápido que celda por celda)
      let width = 0;
      const data = rows.map((row, r) => { if (row.length > width) width = row.length; return row.map((v, c) => { if (v === '') return null; const p = parseInput(v); let x = p.value; if (typeof x === 'string' && x.startsWith('=')) x = toEnglish(x); if (p.fmt) s.styles.set(`${r},${c}`, Object.freeze({ fmt: p.fmt })); return x; }); });
      wb.hf.setSheetContent(s.id, data);
      s.nRows = Math.max(1000, rows.length + 200); s.nCols = Math.max(52, width + 20);
      wb.hf.clearUndoStack?.(); wb.undoStack = [];
      afterLoad(); return {};
    }
    if (f.ext === 'ods') { const r = await readOds(wb, f.bytes); afterLoad(); return { ...r, props: await odsProps(f.bytes) }; }
    const r = await readXlsx(wb, f.bytes);
    afterLoad();
    return r;
  },
  async serialize(ext) {
    if (ext === 'csv') return toCsv(wb);
    if (ext === 'ods') return writeOds(wb, shell.docProps());
    const bytes = await writeXlsx(wb, shell.docProps());
    if (ext === 'xlsm' && wb.vba) return injectVba(bytes, wb.vba);
    return bytes;
  },
  snapshot() { return JSON.stringify(wb.toJSON()); },
  async restore(str) { wb.fromJSON(JSON.parse(str)); afterLoad(); },
  isPristine() { return !shell.doc.dirty && wb.sheets.length === 1 && wb.usedRange().empty; },
  pdfOptions() { return { landscape: !!$('#printRoot').dataset.landscape, pageSize: 'Letter' }; },
  beforePrint() { buildPrint(); },
  afterPrint() { $('#printRoot').replaceChildren(); },
  focus() { grid?.canvas.focus(); },
  stats() { const u = wb.usedRange(); return [['Hojas', wb.sheets.length], ['Área usada (hoja actual)', u.empty ? '—' : rangeStr(u)], ['Fórmulas (hoja actual)', countFormulas()]]; },
  onCommand,
};
async function odsProps(bytes) { try { const JSZip = (await import('jszip')).default; const z = await JSZip.loadAsync(bytes); const t = await z.file('meta.xml')?.async('string'); if (!t) return {}; const d = new DOMParser().parseFromString(t, 'application/xml'); const g = (n) => d.getElementsByTagName(n)[0]?.textContent || ''; return { title: g('dc:title'), author: g('dc:creator') || g('meta:initial-creator') }; } catch { return {}; } }
function countFormulas() { let n = 0; const u = wb.usedRange(); for (let r = 0; r <= u.r2; r++) for (let c = 0; c <= u.c2; c++) if (wb.hf.doesCellHaveFormula({ sheet: wb.sid, row: r, col: c })) n++; return n; }
function afterLoad() {
  tabsBar.dataset.key = '';
  if (wb.sheet?.x3?.hidden) { const v = wb.sheets.findIndex((x) => !x.x3?.hidden); if (v >= 0) wb.active = v; }
  grid.scrollX = grid.scrollY = 0; grid.invalidate(); grid.touchMeta(); grid.updateSpacer(); grid.syncScroll();
  grid.select(0, 0);
  mountCharts(); refresh();
  setTimeout(() => grid.canvas.focus(), 50);
}

// ------------------------------------------------------------------ comandos
async function onCommand(cmd, arg) {
  if (grid.editing && !['undo', 'redo'].includes(cmd) && cmd !== 'findText') grid.commitEdit(null);
  switch (cmd) {
    case 'undo': if (grid.editing) { document.execCommand('undo'); return; } if (wb.undo()) { grid.touchMeta(); grid.invalidate(); grid.updateSpacer(); grid.draw(); refresh(); refreshCharts(); mountCharts(); shell.markDirty(); } return;
    case 'redo': if (wb.redo()) { grid.touchMeta(); grid.invalidate(); grid.updateSpacer(); grid.draw(); refresh(); mountCharts(); shell.markDirty(); } return;
    case 'selectAll': grid.selectRange({ r1: 0, c1: 0, r2: wb.sheet.nRows - 1, c2: wb.sheet.nCols - 1 }); return;
    case 'cut': grid.canvas.focus(); document.execCommand('cut'); return;
    case 'copy': grid.canvas.focus(); document.execCommand('copy'); return;
    case 'paste': return pasteFromClipboard();
    case 'bold': return toggleStyle('b');
    case 'italic': return toggleStyle('i');
    case 'underline': return toggleStyle('u');
    case 'strike': return toggleStyle('s');
    case 'fontName': return styleSel({ font: arg === 'Calibri' ? undefined : arg });
    case 'fontSize': return styleSel({ size: Number(arg) === 11 ? undefined : Number(arg) });
    case 'grow': return styleSel({ size: SIZES.find((s) => s > (curStyle().size || 11)) || 72 });
    case 'shrink': return styleSel({ size: [...SIZES].reverse().find((s) => s < (curStyle().size || 11)) || 8 });
    case 'color': return styleSel({ color: arg === 'transparent' ? undefined : arg });
    case 'fill': return styleSel({ bg: arg === 'transparent' ? undefined : arg });
    case 'alignLeft': return styleSel({ align: curStyle().align === 'left' ? undefined : 'left' });
    case 'alignCenter': return styleSel({ align: curStyle().align === 'center' ? undefined : 'center' });
    case 'alignRight': return styleSel({ align: curStyle().align === 'right' ? undefined : 'right' });
    case 'vTop': return styleSel({ valign: 'top' });
    case 'vMiddle': return styleSel({ valign: 'middle' });
    case 'vBottom': return styleSel({ valign: undefined });
    case 'wrap': return toggleStyle('wrap');
    case 'merge': return toggleMerge();
    case 'borders': return;
    case 'numFmt': if (arg === 'custom') return formatCells(); return styleSel({ fmt: arg === 'General' ? undefined : arg });
    case 'currency': return styleSel({ fmt: '"$"#,##0.00' });
    case 'percent': return styleSel({ fmt: '0%' });
    case 'thousands': return styleSel({ fmt: '#,##0.00' });
    case 'decInc': case 'decDec': {
      const sample = wb.value(grid.sel.r, grid.sel.c);
      return styleSel({ fmt: changeDecimals(curStyle().fmt, cmd === 'decInc' ? 1 : -1, typeof sample === 'number' ? sample : null) });
    }
    case 'formatCells': return formatCells();
    case 'pasteSpecial': return CE.pasteSpecial();
    case 'pasteValues': return CE.doPasteSpecial({ what: 'values' });
    case 'pasteTranspose': return CE.doPasteSpecial({ what: 'all', transpose: true });
    case 'textToColumns': return CE.textToColumns();
    case 'insertCells': return CE.insertCellsDialog();
    case 'deleteCells': return CE.deleteCellsDialog();
    case 'showFormulas': return CE.toggleFormulas();
    case 'hyperlink': return CE.linkDialog();
    case 'tracePre': return CE.trace('pre');
    case 'traceDep': return CE.trace('dep');
    case 'clearArrows': return CE.clearArrows();
    case 'subtotals': return CE.subtotalsDialog();
    case 'stdFilter': return CE.standardFilter();
    case 'borderStyle': return formatCells(3);
    case 'alignDialog': return formatCells(2);
    case 'clearFormat': return clearFormats();
    case 'clearAll': return clearAll();
    case 'clearContent': return clearContents();
    case 'insertRowAbove': return insertRows(true);
    case 'insertColLeft': return insertCols(true);
    case 'deleteRows': return deleteRows();
    case 'deleteCols': return deleteCols();
    case 'autofit': return grid.autofit(grid.sel.c);
    case 'newSheet': return newSheet();
    case 'sortAsc': return sortSel(true);
    case 'sortDialog': return DT.sortDialog();
    case 'scenarioNew': return AN.createScenarioDialog();
    case 'scenarios': return AN.scenariosDialog();
    case 'multipleOps': return AN.multipleOpsDialog();
    case 'sortLists': return AN.sortListsDialog();
    case 'markInvalid': return AN.markInvalid();
    case 'traceError': return AN.traceError();
    case 'clearMarks': return AN.clearMarks();
    case 'insertDate': return AN.insertNow('date');
    case 'insertTime': return AN.insertNow('time');
    case 'fillDown': return DT.fillDir('down');
    case 'fillRight': return DT.fillDir('right');
    case 'fillUp': return DT.fillDir('up');
    case 'fillLeft': return DT.fillDir('left');
    case 'fillSeries': return DT.fillSeriesDialog();
    case 'fillRandom': return DT.randomDialog();
    case 'stats': return DT.statsDialog(arg || 'descriptive');
    case 'consolidate': return DT.consolidateDialog();
    case 'duplicates': return DT.duplicatesDialog();
    case 'cellStyle': return DT.applyCellStyle(arg);
    case 'caseUpper': return DT.textCase('upper');
    case 'caseLower': return DT.textCase('lower');
    case 'caseTitle': return DT.textCase('title');
    case 'caseSentence': return DT.textCase('sentence');
    case 'caseToggle': return DT.textCase('toggle');
    case 'moveCopySheet': return DT.moveCopyDialog();
    case 'hideSheet': return DT.hideSheet();
    case 'showSheets': return DT.showSheetDialog();
    case 'sortDesc': return sortSel(false);
    case 'filter': return toggleFilter();
    case 'freezePanes': return;
    case 'autoSum': return autoSum('SUM');
    case 'insertFunction': return insertFunction();
    case 'insertChart': return insertChart(typeof arg === 'string' ? arg : null);
    case 'find': case 'replace': return findReplace(cmd === 'replace');
    case 'findText': return findNext(arg);
    case 'zoomIn': return zoom.set(zoom.get() + 10);
    case 'zoomOut': return zoom.set(zoom.get() - 10);
    case 'zoomReset': return zoom.set(100);
    case 'gridlines': return;
    case 'exportCsv': return shell.saveAsFormat('csv', 'CSV');
    case 'exportOds': return shell.saveAsFormat('ods', 'Hoja de cálculo ODF');
    case 'cf': return;
    case 'validation': return F.validationDialog();
    case 'clearValidation': return F.clearValidation();
    case 'defineName': return F.defineName();
    case 'nameManager': return F.manageNames();
    case 'note': return F.noteDialog();
    case 'deleteNote': return F.setNote(grid.sel.r, grid.sel.c, '');
    case 'protect': return F.protectDialog();
    case 'protectBook': return F.protectBook();
    case 'unlockCells': return F.toggleUnlock();
    case 'pivot': return F.pivotDialog();
    case 'goalSeek': return F3.goalSeek();
    case 'solver': return F3.solver();
    case 'getDataCsv': return Q.getData('csv');
    case 'getDataXlsx': return Q.getData('xlsx');
    case 'getDataJson': return Q.getData('json');
    case 'getDataRange': return Q.getData('range');
    case 'queries': return Q.manageQueries();
    case 'refreshAll': return Q.refreshAll();
    case 'sparkline': return F3.sparklineDialog(arg || 'line');
    case 'clearSparklines': return F3.clearSparklines();
    case 'group': return F3.group(false);
    case 'ungroup': return F3.group(true);
    case 'hideDetail': return F3.showDetail(false);
    case 'showDetail': return F3.showDetail(true);
    case 'slicer': return F3.insertSlicer();
    case 'sparklineCol': return F3.sparklineDialog('column');
    case 'sparklineWin': return F3.sparklineDialog('winloss');
    case 'refreshPivot': return F.refreshPivot();
    case 'pageSetup': return F.pageSetupDialog();
    case 'printArea': return F.setPrintArea(false);
    case 'clearPrintArea': return F.setPrintArea(true);
    case 'orientPortrait': act('Orientación', () => { wb.sheet.pageSetup = { ...(wb.sheet.pageSetup || {}), orientation: 'portrait' }; }, { hf: false }); return;
    case 'orientLandscape': act('Orientación', () => { wb.sheet.pageSetup = { ...(wb.sheet.pageSetup || {}), orientation: 'landscape' }; }, { hf: false }); return;
    case 'fitWidth': act('Ajustar', () => { wb.sheet.pageSetup = { ...(wb.sheet.pageSetup || {}), fit: wb.sheet.pageSetup?.fit === 'width' ? 'none' : 'width' }; }, { hf: false }); return;
    default: return undefined;
  }
}

// ------------------------------------------------------------------ buscar y reemplazar
const nrm = (s) => String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
function findNext(q, from = grid.sel) {
  if (!q) return false;
  const u = wb.usedRange(), nq = nrm(q);
  const total = (u.r2 + 1) * (u.c2 + 1);
  let idx = from.r * (u.c2 + 1) + from.c;
  for (let i = 1; i <= total; i++) {
    const k = (idx + i) % total; const r = Math.floor(k / (u.c2 + 1)), c = k % (u.c2 + 1);
    if (nrm(wb.display(r, c).text).includes(nq) || nrm(wb.editText(r, c)).includes(nq)) { grid.select(r, c); return true; }
  }
  toast(`No se encontró "${q}"`); return false;
}
async function findReplace(replace) {
  const v = await form(replace ? 'Reemplazar' : 'Buscar', [{ name: 'q', label: 'Buscar', value: shell.searchBox.value }, ...(replace ? [{ name: 'rep', label: 'Reemplazar con', value: '' }] : [])], replace ? 'Reemplazar todo' : 'Buscar siguiente');
  if (!v || !v.q) return;
  if (!replace) { findNext(v.q); return; }
  const u = wb.usedRange(); let n = 0;
  const re = new RegExp(v.q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
  act('Reemplazar', () => {
    for (let r = 0; r <= u.r2; r++) for (let c = 0; c <= u.c2; c++) {
      const raw = wb.raw(r, c);
      if (typeof raw === 'string' && !raw.startsWith('=') && re.test(raw)) { re.lastIndex = 0; wb.setRaw(r, c, raw.replace(re, v.rep)); n++; }
      re.lastIndex = 0;
    }
  });
  toast(`${n} celdas reemplazadas`);
}

// ------------------------------------------------------------------ cinta
const borderMenu = () => [
  { label: 'Todos los bordes', icon: 'borders', run: () => setBorders('all') },
  { label: 'Bordes externos', icon: 'rect', run: () => setBorders('outer') },
  { label: 'Borde inferior', icon: 'underline', run: () => setBorders('bottom') },
  { label: 'Borde superior', icon: 'hr', run: () => setBorders('top') },
  '-', { label: 'Sin bordes', icon: 'close', run: () => setBorders('none') },
];
const freezeMenu = () => [
  { label: wb.sheet.freeze.rows || wb.sheet.freeze.cols ? 'Movilizar paneles' : 'Inmovilizar paneles', icon: 'freeze', run: () => freeze(wb.sheet.freeze.rows || wb.sheet.freeze.cols ? 'none' : 'sel') },
  { label: 'Inmovilizar fila superior', icon: 'rowAdd', run: () => freeze('row') },
  { label: 'Inmovilizar primera columna', icon: 'colAdd', run: () => freeze('col') },
];
const sumMenu = () => [['SUM', 'Suma'], ['AVERAGE', 'Promedio'], ['COUNT', 'Contar números'], ['MAX', 'Máximo'], ['MIN', 'Mínimo']].map(([f, l]) => ({ label: l, run: () => autoSum(f) })).concat(['-', { label: 'Más funciones…', run: insertFunction }]);
const chartMenu = () => CHART_TYPES.map(([k, l]) => ({ label: l, icon: { pie: 'chartPie', line: 'chartLine', area: 'chartLine' }[k] || 'chart', run: () => insertChart(k) }));
const fillMenu = () => [{ label: 'Abajo', icon: 'arrow', kbd: 'Ctrl+D', run: () => DT.fillDir('down') }, { label: 'Derecha', icon: 'arrow', run: () => DT.fillDir('right') }, { label: 'Arriba', icon: 'arrow', run: () => DT.fillDir('up') }, { label: 'Izquierda', icon: 'arrow', run: () => DT.fillDir('left') }, '-', { label: 'Serie…', icon: 'sortAsc', run: () => DT.fillSeriesDialog() }, { label: 'Números aleatorios…', icon: 'func', run: () => DT.randomDialog() }];
const caseMenu = () => [{ label: 'MAYÚSCULAS', run: () => DT.textCase('upper') }, { label: 'minúsculas', run: () => DT.textCase('lower') }, { label: 'Tipo oración', run: () => DT.textCase('sentence') }, { label: 'Tipo Título', run: () => DT.textCase('title') }, { label: 'aLTERNAR mAYÚSCULAS', run: () => DT.textCase('toggle') }];
const statsMenu = () => [['descriptive', 'Estadística descriptiva…'], ['correlation', 'Correlación…'], ['covariance', 'Covarianza…'], ['moving', 'Media móvil…'], ['smoothing', 'Suavizado exponencial…'], ['sampling', 'Muestreo…'], ['regression', 'Regresión lineal…']].map(([k, l]) => ({ label: l, run: () => DT.statsDialog(k) }));
const clearMenu = () => [{ label: 'Borrar todo', icon: 'trash', run: clearAll }, { label: 'Borrar formatos', icon: 'clear', run: clearFormats }, { label: 'Borrar contenido', icon: 'close', run: clearContents }];
const insMenu = () => [{ label: 'Insertar celdas…', icon: 'rowAdd', kbd: 'Ctrl++', run: () => CE.insertCellsDialog() }, { label: 'Insertar filas', icon: 'rowAdd', run: () => insertRows(true) }, { label: 'Insertar columnas', icon: 'colAdd', run: () => insertCols(true) }, { label: 'Insertar hoja', icon: 'sheetAdd', run: newSheet }];
const delMenu = () => [{ label: 'Eliminar celdas…', icon: 'rowDel', kbd: 'Ctrl+-', run: () => CE.deleteCellsDialog() }, { label: 'Eliminar filas', icon: 'rowDel', run: deleteRows }, { label: 'Eliminar columnas', icon: 'colDel', run: deleteCols }, { label: 'Eliminar hoja', icon: 'trash', run: () => deleteSheet(wb.active) }];
const fmtMenu = () => [{ label: 'Autoajustar ancho de columna', icon: 'columns', run: () => grid.autofit(grid.sel.c) }, { label: 'Ancho de columna…', icon: 'columns', run: async () => { const v = await form('Ancho de columna', [{ name: 'w', label: 'Ancho (px)', type: 'number', value: wb.sheet.cw(grid.sel.c), min: 8, max: 800 }]); if (v) { const g = sel(); act('Ancho', () => { for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 200); c++) wb.sheet.colW.set(c, v.w); }, { hf: false }); } } }, { label: 'Alto de fila…', icon: 'lineSpacing', run: async () => { const v = await form('Alto de fila', [{ name: 'h', label: 'Alto (px)', type: 'number', value: wb.sheet.rh(grid.sel.r), min: 10, max: 400 }]); if (v) { const g = clampToUsed(sel()); act('Alto', () => { for (let r = g.r1; r <= g.r2; r++) wb.sheet.rowH.set(r, v.h); }, { hf: false }); } } }, '-', { label: 'Cambiar nombre de hoja…', icon: 'textbox', run: () => renameSheet(wb.active) }, { label: 'Mover o copiar hoja…', icon: 'copy', run: () => DT.moveCopyDialog() }, { label: 'Ocultar hoja', icon: 'minus', run: () => DT.hideSheet() }, { label: 'Mostrar hoja…', icon: 'plus', run: () => DT.showSheetDialog() }, '-', ...caseMenu(), '-', { label: 'Formato de celdas…', icon: 'brush', kbd: 'Ctrl+1', run: formatCells }];

const TABS = [
  { id: 'home', label: 'Inicio', groups: [
    { label: 'Portapapeles', items: [{ cmd: 'paste', label: 'Pegar', icon: 'paste', size: 'lg', split: () => [{ label: 'Pegar', run: () => onCommand('paste') }, { label: 'Pegar solo valores', run: () => onCommand('pasteValues') }, { label: 'Transponer', run: () => onCommand('pasteTranspose') }, '-', { label: 'Pegado especial…', kbd: 'Ctrl+Mayús+V', run: () => onCommand('pasteSpecial') }] }, { stack: [{ cmd: 'cut', label: 'Cortar', icon: 'cut' }, { cmd: 'copy', label: 'Copiar', icon: 'copy' }, { cmd: 'clearFormat', label: 'Borrar formato', icon: 'clear' }] }] },
    { label: 'Fuente', items: [{ stack: [
      { row: [{ cmd: 'fontName', label: 'Fuente', select: { options: FONTS, width: '130px' } }, { cmd: 'fontSize', label: 'Tamaño', select: { options: SIZES.map(String), width: '56px' } }, { cmd: 'grow', label: 'Aumentar', icon: 'grow' }, { cmd: 'shrink', label: 'Reducir', icon: 'shrink' }] },
      { row: [{ cmd: 'bold', label: 'Negrita', icon: 'bold', kbd: 'Ctrl+B' }, { cmd: 'italic', label: 'Cursiva', icon: 'italic', kbd: 'Ctrl+I' }, { cmd: 'underline', label: 'Subrayado', icon: 'underline', kbd: 'Ctrl+U' }, { cmd: 'strike', label: 'Tachado', icon: 'strike' }, { cmd: 'borders', label: 'Bordes', icon: 'borders', menu: borderMenu }, { cmd: 'fill', label: 'Color de relleno', icon: 'fill', color: true, colorValue: '#FFF59D' }, { cmd: 'color', label: 'Color de fuente', icon: 'textColor', color: true, colorValue: '#C62828', colorOpts: { allowNone: false } }] },
    ] }] },
    { label: 'Alineación', items: [{ stack: [
      { row: [{ cmd: 'vTop', label: 'Arriba', icon: 'alignObj' }, { cmd: 'vMiddle', label: 'Centro vertical', icon: 'alignCenter' }, { cmd: 'vBottom', label: 'Abajo', icon: 'footer' }, { cmd: 'wrap', label: 'Ajustar texto', icon: 'wrap', showLabel: true }] },
      { row: [{ cmd: 'alignLeft', label: 'Izquierda', icon: 'alignLeft' }, { cmd: 'alignCenter', label: 'Centrar', icon: 'alignCenter' }, { cmd: 'alignRight', label: 'Derecha', icon: 'alignRight' }, { cmd: 'merge', label: 'Combinar y centrar', icon: 'merge', showLabel: true }] },
    ] }, { cmd: 'alignDialog', label: 'Alineación…', icon: 'dialogLauncher' }] },
    { label: 'Número', items: [{ stack: [
      { row: [{ cmd: 'numFmt', label: 'Formato de número', select: { options: [...FORMATS.map(([k, n]) => [k, n]), ['custom', 'Más formatos…']], width: '170px' } }] },
      { row: [{ cmd: 'currency', label: 'Moneda', icon: 'currency' }, { cmd: 'percent', label: 'Porcentaje', icon: 'percent' }, { cmd: 'thousands', label: 'Separador de miles', icon: 'hr' }, { cmd: 'decInc', label: 'Aumentar decimales', icon: 'decInc' }, { cmd: 'decDec', label: 'Disminuir decimales', icon: 'decDec' }] },
    ] }] },
    { label: 'Estilos', items: [{ cmd: 'cf', label: 'Formato condicional', icon: 'fill', size: 'lg', menu: () => F.cfMenu() }, { cmd: 'tableStyle', label: 'Formato de tabla', icon: 'table', size: 'lg', menu: () => F.tableStyleMenu() }, { cmd: 'cellStyles', label: 'Estilos de celda', icon: 'brush', size: 'lg', menu: () => DT.cellStylesMenu() }] },
    { label: 'Celdas', items: [{ cmd: 'insert', label: 'Insertar', icon: 'rowAdd', size: 'lg', menu: insMenu }, { cmd: 'delete', label: 'Eliminar', icon: 'rowDel', size: 'lg', menu: delMenu }, { cmd: 'format', label: 'Formato', icon: 'brush', size: 'lg', menu: fmtMenu }] },
    { label: 'Edición', items: [{ stack: [{ cmd: 'autoSum', label: 'Autosuma', icon: 'sum', showLabel: true, menu: sumMenu }, { cmd: 'fillMenu', label: 'Rellenar', icon: 'arrow', showLabel: true, menu: fillMenu }, { cmd: 'clear', label: 'Borrar', icon: 'clear', showLabel: true, menu: clearMenu }, { cmd: 'sort', label: 'Ordenar y filtrar', icon: 'filter', showLabel: true, menu: () => [{ label: 'Ordenar de A a Z', icon: 'sortAsc', run: () => sortSel(true) }, { label: 'Ordenar de Z a A', icon: 'sortDesc', run: () => sortSel(false) }, { label: 'Ordenar…', icon: 'sortAsc', run: () => DT.sortDialog() }, '-', { label: 'Filtro', icon: 'filter', run: toggleFilter }] }] }] },
  ] },
  { id: 'insert', label: 'Insertar', groups: [
    { label: 'Tablas', items: [{ cmd: 'pivot', label: 'Tabla dinámica', icon: 'layout', size: 'lg' }, { cmd: 'tableStyle', label: 'Tabla', icon: 'table', size: 'lg', menu: () => F.tableStyleMenu() }] },
    { label: 'Gráficos', items: [{ cmd: 'insertChart', label: 'Gráfico recomendado', icon: 'chart', size: 'lg' }, { cmd: 'charts', label: 'Todos los gráficos', icon: 'chartPie', size: 'lg', menu: chartMenu }] },
    { label: 'Minigráficos', items: [{ stack: [{ cmd: 'sparkline', label: 'Línea', icon: 'chartLine', showLabel: true }, { cmd: 'sparklineCol', label: 'Columna', icon: 'chart', showLabel: true }, { cmd: 'sparklineWin', label: 'Ganancia o pérdida', icon: 'decDec', showLabel: true }] }] },
    { label: 'Filtros', items: [{ cmd: 'slicer', label: 'Segmentación de datos', icon: 'filter', size: 'lg' }] },
    { label: 'Funciones', items: [{ cmd: 'insertFunction', label: 'Función', icon: 'func', size: 'lg' }, { cmd: 'dateTime', label: 'Fecha y hora', icon: 'recover', size: 'lg', menu: () => [{ label: 'Fecha de hoy', kbd: 'Ctrl+;', run: () => onCommand('insertDate') }, { label: 'Hora actual', kbd: 'Ctrl+Mayús+;', run: () => onCommand('insertTime') }] }, { cmd: 'autoSum', label: 'Autosuma', icon: 'sum', size: 'lg', menu: sumMenu }] },
    { label: 'Vínculos', items: [{ cmd: 'hyperlink', label: 'Hipervínculo', icon: 'link', size: 'lg', kbd: 'Ctrl+K' }] },
    { label: 'Hojas', items: [{ cmd: 'newSheet', label: 'Hoja nueva', icon: 'sheetAdd', size: 'lg' }] },
  ] },
  { id: 'page', label: 'Disposición de página', groups: [
    { label: 'Configurar página', items: [{ cmd: 'orient', label: 'Orientación', icon: 'orientation', size: 'lg', menu: () => [{ label: 'Vertical', run: () => onCommand('orientPortrait') }, { label: 'Horizontal', run: () => onCommand('orientLandscape') }] }, { cmd: 'area', label: 'Área de impresión', icon: 'pageSetup', size: 'lg', menu: () => [{ label: 'Establecer con la selección', run: () => onCommand('printArea') }, { label: 'Borrar área de impresión', run: () => onCommand('clearPrintArea') }] }, { cmd: 'fitWidth', label: 'Ajustar al ancho', icon: 'columns', size: 'lg' }, { cmd: 'pageSetup', label: 'Más opciones', icon: 'pageSetup', size: 'lg' }] },
  ] },
  { id: 'formulas', label: 'Fórmulas', groups: [
    { label: 'Biblioteca de funciones', items: [{ cmd: 'insertFunction', label: 'Insertar función', icon: 'func', size: 'lg' }, { cmd: 'autoSum', label: 'Autosuma', icon: 'sum', size: 'lg', menu: sumMenu }] },
    { label: 'Nombres definidos', items: [{ cmd: 'defineName', label: 'Definir nombre', icon: 'textbox', size: 'lg' }, { cmd: 'nameManager', label: 'Administrador de nombres', icon: 'layout', size: 'lg' }] },
    { label: 'Auditoría de fórmulas', items: [{ stack: [{ cmd: 'tracePre', label: 'Rastrear precedentes', icon: 'tracePre', showLabel: true }, { cmd: 'traceDep', label: 'Rastrear dependientes', icon: 'traceDep', showLabel: true }, { cmd: 'clearArrows', label: 'Quitar flechas', icon: 'close', showLabel: true }] }, { stack: [{ cmd: 'traceError', label: 'Rastrear error', icon: 'tracePre', showLabel: true }, { cmd: 'markInvalid', label: 'Marcar datos no válidos', icon: 'check', showLabel: true }, { cmd: 'clearMarks', label: 'Quitar marcas', icon: 'close', showLabel: true }] }, { cmd: 'showFormulas', label: 'Mostrar fórmulas', icon: 'func', size: 'lg', kbd: 'Ctrl+`' }] },
  ] },
  { id: 'data', label: 'Datos', groups: [
    { label: 'Obtener y transformar', items: [{ cmd: 'getData', label: 'Obtener datos', icon: 'open', size: 'lg', menu: () => [{ label: 'Desde texto/CSV…', run: () => onCommand('getDataCsv') }, { label: 'Desde un libro de Excel…', run: () => onCommand('getDataXlsx') }, { label: 'Desde JSON…', run: () => onCommand('getDataJson') }, '-', { label: 'Desde tabla o rango', run: () => onCommand('getDataRange') }] }, { cmd: 'refreshAll', label: 'Actualizar todo', icon: 'recover', size: 'lg' }, { cmd: 'queries', label: 'Consultas y conexiones', icon: 'layout', size: 'lg' }] },
    { label: 'Ordenar y filtrar', items: [{ cmd: 'sortAsc', label: 'A → Z', icon: 'sortAsc', size: 'lg' }, { cmd: 'sortDesc', label: 'Z → A', icon: 'sortDesc', size: 'lg' }, { cmd: 'sortDialog', label: 'Ordenar…', icon: 'sortAsc', size: 'lg' }, { cmd: 'sortLists', label: 'Listas de ordenamiento…', icon: 'bullets', size: 'lg' }, { cmd: 'filter', label: 'Filtro', icon: 'filter', size: 'lg', kbd: 'Ctrl+Mayús+L' }, { cmd: 'stdFilter', label: 'Filtro estándar…', icon: 'filter', size: 'lg' }] },
    { label: 'Herramientas de datos', items: [{ cmd: 'textToColumns', label: 'Texto en columnas…', icon: 'columns', size: 'lg' }, { cmd: 'subtotals', label: 'Subtotales…', icon: 'sum', size: 'lg' }, { cmd: 'duplicates', label: 'Quitar duplicados…', icon: 'trash', size: 'lg' }, { cmd: 'consolidate', label: 'Consolidar…', icon: 'layout', size: 'lg' }, { cmd: 'validation', label: 'Validación de datos', icon: 'check', size: 'lg' }, { cmd: 'clearValidation', label: 'Quitar validación', icon: 'close', size: 'lg' }, { cmd: 'refreshPivot', label: 'Actualizar dinámica', icon: 'recover', size: 'lg' }] },
    { label: 'Análisis', items: [{ cmd: 'statsMenu', label: 'Estadísticas', icon: 'chart', size: 'lg', menu: statsMenu }, { cmd: 'fillRandom', label: 'Números aleatorios…', icon: 'func', size: 'lg' }] },
    { label: 'Previsión', items: [{ cmd: 'whatIf', label: 'Análisis de hipótesis', icon: 'func', size: 'lg', menu: () => [{ label: 'Administrador de escenarios…', run: () => onCommand('scenarios') }, { label: 'Crear escenario…', run: () => onCommand('scenarioNew') }, { label: 'Operaciones múltiples (tabla de datos)…', run: () => onCommand('multipleOps') }, '-', { label: 'Buscar objetivo…', run: () => onCommand('goalSeek') }, { label: 'Solver…', run: () => onCommand('solver') }] }] },
    { label: 'Esquema', items: [{ stack: [{ cmd: 'group', label: 'Agrupar', icon: 'group', showLabel: true }, { cmd: 'ungroup', label: 'Desagrupar', icon: 'close', showLabel: true }] }, { stack: [{ cmd: 'hideDetail', label: 'Ocultar detalle', icon: 'minus', showLabel: true }, { cmd: 'showDetail', label: 'Mostrar detalle', icon: 'plus', showLabel: true }] }] },
    { label: 'Exportar', items: [{ cmd: 'exportCsv', label: 'Guardar CSV', icon: 'words', size: 'lg' }, { cmd: 'exportOds', label: 'Guardar ODS', icon: 'sheet', size: 'lg' }] },
  ] },
  { id: 'review', label: 'Revisar', groups: [
    { label: 'Notas', items: [{ cmd: 'note', label: 'Nueva nota', icon: 'comment', size: 'lg' }, { cmd: 'deleteNote', label: 'Eliminar nota', icon: 'trash', size: 'lg' }] },
    { label: 'Proteger', items: [{ cmd: 'protect', label: 'Proteger hoja', icon: 'freeze', size: 'lg' }, { cmd: 'protectBook', label: 'Proteger libro', icon: 'sheet', size: 'lg' }, { cmd: 'unlockCells', label: 'Desbloquear celdas', icon: 'check', size: 'lg' }] },
  ] },
  { id: 'view', label: 'Vista', groups: [
    { label: 'Ventana', items: [{ cmd: 'freezePanes', label: 'Inmovilizar', icon: 'freeze', size: 'lg', menu: freezeMenu }] },
    { label: 'Zoom', items: [{ cmd: 'zoomReset', label: '100 %', icon: 'zoomIn', size: 'lg' }, { cmd: 'zoomIn', label: 'Acercar', icon: 'zoomIn', size: 'lg' }, { cmd: 'zoomOut', label: 'Alejar', icon: 'zoomOut', size: 'lg' }] },
  ] },
];

// ------------------------------------------------------------------ arranque
shell = new OfficeShell({ appId: 'calcula', appName: 'Calcula Datos', logo: '../shared/img/calcula.svg', tabs: TABS, adapter });
buildWorkspace(shell.workspace);
F.initFeatures({ wb, grid, shell, act, refresh, switchSheet, sel: () => grid.range, clamp: (g) => clampToUsed(g) });
F3.init({ wb, grid, shell, act, refresh, applyFilter, mountCharts, refreshCharts });
Q.init({ wb, grid, shell, act, refresh, switchSheet, refreshCharts });
AN.init({ wb, grid, shell, act, refresh, switchSheet, refreshCharts, sel: () => grid.range, clamp: (g) => clampToUsed(g), renderTabs: () => { tabsBar.dataset.key = ''; renderTabs(); } });
DT.init({ wb, grid, shell, act, refresh, switchSheet, dup: (i) => duplicateSheet(i), refreshCharts, sel: () => grid.range, clamp: (g) => clampToUsed(g), renderTabs: () => { tabsBar.dataset.key = ''; renderTabs(); } });
CE.init({ wb, grid, shell, act, refresh, switchSheet, sel: () => grid.range, clamp: (g) => clampToUsed(g), curStyle: () => curStyle(), FONTS, renderTabs: () => { tabsBar.dataset.key = ''; renderTabs(); } });
{ const pc = grid.positionCharts.bind(grid); grid.positionCharts = () => { pc(); F3.positionSlicers(); }; }
statusInfo = h('span', { class: 'selstats' });
zoom = shell.zoomControl({ value: 100, min: 50, max: 250, onChange: (v) => { grid.zoom = v / 100; grid.updateSpacer(); grid.syncScroll(); grid.draw(); } });
shell.status.append(h('span', {}, 'Listo'), h('div', { class: 'grow' }), statusInfo, zoom);
addEventListener('keydown', (e) => {
  const mod = e.ctrlKey || e.metaKey;
  if (mod && e.key === '1') { e.preventDefault(); formatCells(); }
  if (mod && (e.key === ';' || e.key === ':' || e.code === 'Semicolon') && document.activeElement === grid.canvas) { e.preventDefault(); AN.insertNow(e.shiftKey ? 'time' : 'date'); }
  if (mod && e.shiftKey && (e.key === 'V' || e.key === 'v') && document.activeElement === grid.canvas) { e.preventDefault(); CE.pasteSpecial(); }
  if (mod && !e.shiftKey && (e.key === 'k' || e.key === 'K') && document.activeElement === grid.canvas) { e.preventDefault(); CE.linkDialog(); }
  if (mod && (e.key === '`' || e.code === 'Backquote')) { e.preventDefault(); CE.toggleFormulas(); }
  if (mod && (e.key === '+' || (e.shiftKey && e.key === '=')) && document.activeElement === grid.canvas) { e.preventDefault(); CE.insertCellsDialog(); }
  if (mod && e.key === '-' && document.activeElement === grid.canvas) { e.preventDefault(); CE.deleteCellsDialog(); }
  if (mod && e.shiftKey && (e.key === 'L' || e.key === 'l')) { e.preventDefault(); toggleFilter(); }
  if (e.shiftKey && e.key === 'F3') { e.preventDefault(); insertFunction(); }
  if (e.shiftKey && e.key === 'F11') { e.preventDefault(); newSheet(); }
  if (e.altKey && e.key === '=') { e.preventDefault(); autoSum(); }
  if (mod && !e.shiftKey && ['b', 'i', 'u'].includes(e.key.toLowerCase()) && document.activeElement === grid.canvas) { e.preventDefault(); toggleStyle({ b: 'b', i: 'i', u: 'u' }[e.key.toLowerCase()]); }
});
emptyBook('blank');
afterLoad();
shell.start();
// Al cambiar el idioma, las fórmulas y errores se vuelven a mostrar en el idioma nuevo
onLang(() => { grid.invalidate(); grid.draw(); refresh(); });

window.__calcula = { wb, grid, onCommand, emptyBook, afterLoad, buildPrint, refresh, F, F3, Q, CE, DT, AN };
