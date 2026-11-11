// Calcula Datos — funciones de LibreOffice Calc que faltaban: cuadro Formato de celdas completo
// (números, fuente, alineación con sangría y giro, bordes con estilo, fondo, protección), pegado
// especial (valores, formatos, operaciones, transponer), texto en columnas, insertar y eliminar
// celdas desplazando, mostrar fórmulas, color de pestaña, hipervínculos, rastrear precedentes y
// dependientes, subtotales, filtro estándar y autocompletar. GPL-3.0
import { h, toast, form, t, tf, dropdown, closeDropdown } from '../../shared/office.js';
import { tabbed, field, numIn, sel, chk } from '../../shared/dialogs.js';
import { colName, addr, rangeStr, parseRange } from './workbook.js';
import { formatValue, FORMATS } from './format.js';

let C = null; // { wb, grid, shell, act, refresh, sel, clamp, curStyle, FONTS, switchSheet, renderTabs }
const X3 = (s = C.wb.sheet) => { if (!s.x3) s.x3 = { sparklines: [], rowGroups: [], colGroups: [], slicers: [] }; return s.x3; };
export function init(ctx) {
  C = ctx;
  const g = C.grid;
  g.linkAt = (r, c) => X3().links?.[`${r},${c}`] || null;
  const prev = g.overlay; g.overlay = (ctx2, ...a) => { prev?.(ctx2, ...a); drawArrows(ctx2); };
  // Ctrl+clic abre el hipervínculo
  g.canvas.addEventListener('mousedown', (e) => { if (!(e.ctrlKey || e.metaKey)) return; const { r, c } = g.hit(e.offsetX, e.offsetY); if (r >= 0 && c >= 0 && followLink(r, c)) { e.preventDefault(); e.stopImmediatePropagation(); } }, true);
  g.canvas.addEventListener('mousemove', (e) => { if (!(e.ctrlKey || e.metaKey)) return; const { r, c } = g.hit(e.offsetX, e.offsetY); if (g.linkAt(r, c)) g.canvas.style.cursor = 'pointer'; });
  // autocompletar al escribir texto (lo sugerido queda seleccionado; Supr lo quita)
  const ed = g.editor;
  ed.addEventListener('input', (e) => {
    if (!g.editing || !/^insert(Text|CompositionText)$/.test(e.inputType || '') || ed.selectionStart !== ed.value.length || ed.value.includes('\n')) return;
    const typed = ed.value; const sug = autoComplete(g.editing.r, g.editing.c, typed);
    if (sug) { ed.value = typed + sug.slice(typed.length); ed.setSelectionRange(typed.length, ed.value.length); g.onEdit?.(ed.value); }
  });
}

// ================================================================ Formato de celdas
const CATS = [
  ['general', 'General'], ['number', 'Número'], ['percent', 'Porcentaje'], ['currency', 'Moneda'], ['date', 'Fecha'], ['time', 'Hora'],
  ['sci', 'Científico'], ['fraction', 'Fracción'], ['bool', 'Valor lógico'], ['text', 'Texto'], ['custom', 'Definido por el usuario'],
];
const DATE_FMTS = ['dd/mm/yyyy', 'dd/mm/yy', 'd "de" mmmm "de" yyyy', 'dddd, d "de" mmmm "de" yyyy', 'mmm-yy', 'yyyy-mm-dd', 'dd/mm/yyyy hh:mm'];
const TIME_FMTS = ['hh:mm', 'hh:mm:ss', 'h:mm AM/PM', '[h]:mm:ss'];
function catOf(f) {
  if (!f || f === 'General') return 'general';
  if (f === '@') return 'text';
  if (/E\+/i.test(f)) return 'sci';
  if (/\?\/\?/.test(f)) return 'fraction';
  if (/%/.test(f)) return 'percent';
  if (/"\$"|\$/.test(f)) return 'currency';
  if (DATE_FMTS.includes(f) || (/[dy]/i.test(f.replace(/"[^"]*"/g, '')) && !/[0#]/.test(f))) return 'date';
  if (TIME_FMTS.includes(f) || (/h|s/i.test(f.replace(/"[^"]*"/g, '')) && !/[0#]/.test(f))) return 'time';
  if (/^"VERDADERO";"VERDADERO";"FALSO"$|TRUE/i.test(f)) return 'bool';
  if (/^[#,0.;[\]Red\-]+$/i.test(f)) return 'number';
  return 'custom';
}
// arma el código a partir de las opciones (como el panel de Opciones de LibreOffice)
function buildCode(cat, { dec = 2, lead = 1, neg = false, thou = false }) {
  const z = `${thou ? '#,##' : ''}${'0'.repeat(Math.max(1, lead))}${dec > 0 ? `.${'0'.repeat(dec)}` : ''}`.replace(/^#,##0{2,}/, (m) => `#,##${m.slice(4)}`);
  if (cat === 'number') return neg ? `${z};[Red]-${z}` : z;
  if (cat === 'percent') return `${z}%`;
  if (cat === 'currency') return neg ? `"$"${z};[Red]-"$"${z}` : `"$"${z}`;
  if (cat === 'sci') return `0${dec > 0 ? `.${'0'.repeat(dec)}` : ''}E+00`;
  return null;
}
function optsOf(f) {
  const s = String(f || '').split(';')[0].replace(/"[^"]*"/g, '');
  const m = /\.(0+)/.exec(s); const dec = m ? m[1].length : 0;
  const lead = (/(0+)(?:\.|%|$|E)/.exec(s.replace(/#,##/, '')) || [, '0'])[1].length || 1;
  return { dec, lead, neg: /\[Red\]|\[Rojo\]/i.test(f || ''), thou: /,/.test(s) };
}
const SIDES = [['t', 'Superior'], ['b', 'Inferior'], ['l', 'Izquierdo'], ['r', 'Derecho']];
export const LINE_STYLES = [['thin', 'Fino'], ['medium', 'Medio'], ['thick', 'Grueso'], ['dashed', 'Discontinuo'], ['dotted', 'Punteado'], ['double', 'Doble'], ['hair', 'Muy fino']];
const BORDER_PRESETS = [['', '(sin cambios)'], ['none', 'Sin bordes'], ['outer', 'Solo contorno'], ['all', 'Contorno y todas las líneas interiores'], ['outerH', 'Contorno y líneas horizontales'], ['inner', 'Solo líneas interiores'], ['bottom', 'Solo inferior'], ['custom', 'Personalizado (lados marcados)']];

export async function formatCellsDialog(start = 0) {
  const g = C.clamp(C.sel()); const st = C.curStyle(); const wb = C.wb;
  const sample = wb.value(C.grid.sel.r, C.grid.sel.c);
  const sampleN = typeof sample === 'number' ? sample : -1234.5678;
  // --- Números
  let cat = catOf(st.fmt);
  const o = optsOf(st.fmt);
  const catSel = h('select', { name: 'cat', size: 11, class: 'cf-cats' }, ...CATS.map(([k, l]) => h('option', { value: k, selected: k === cat ? true : null }, t(l))));
  const fmtList = h('select', { name: 'fmtlist', size: 7, class: 'cf-list' });
  const code = h('input', { name: 'code', value: st.fmt || 'General', class: 'no-i18n' });
  const prev = h('div', { class: 'fx-preview cf-prev no-i18n' });
  const opt = { dec: numIn('dec', o.dec, { step: 1, min: 0, max: 15 }), lead: numIn('lead', o.lead, { step: 1, min: 1, max: 15 }), neg: chk('neg', 'Negativos en rojo', o.neg), thou: chk('thou', 'Separador de miles', o.thou) };
  const optBox = h('div', { class: 'fx-grid' }, h('h4', {}, t('Opciones')), field('Decimales', opt.dec), field('Ceros a la izquierda', opt.lead), opt.neg, opt.thou);
  const listFor = (k) => (k === 'date' ? DATE_FMTS : k === 'time' ? TIME_FMTS : k === 'fraction' ? ['# ?/?', '# ??/??', '# ?/4', '# ?/100'] : k === 'bool' ? ['"VERDADERO";"VERDADERO";"FALSO"'] : k === 'custom' ? FORMATS.map((f) => f[0]) : k === 'general' ? ['General'] : k === 'text' ? ['@'] : []);
  const v = (box, n) => box.querySelector(`[name="${n}"]`);
  const fillList = () => { fmtList.replaceChildren(...listFor(cat).map((f) => h('option', { value: f, selected: f === code.value ? true : null }, formatValue(cat === 'text' ? 'Texto' : (cat === 'date' || cat === 'time') && typeof sample !== 'number' ? 46000.5625 : sampleN, f).text || f))); fmtList.hidden = !fmtList.children.length; optBox.hidden = !['number', 'percent', 'currency', 'sci'].includes(cat); };
  const upd = () => { const r = formatValue(cat === 'text' && typeof sample !== 'number' ? String(sample ?? 'Texto') : (cat === 'date' || cat === 'time') && typeof sample !== 'number' ? 46000.5625 : sampleN, code.value || 'General'); prev.textContent = r.text; prev.style.color = r.color || ''; };
  const fromOpts = () => { const c = buildCode(cat, { dec: Number(v(optBox, 'dec').value), lead: Number(v(optBox, 'lead').value), neg: v(optBox, 'neg').checked, thou: v(optBox, 'thou').checked }); if (c) code.value = c; upd(); };
  catSel.addEventListener('change', () => { cat = catSel.value; const l = listFor(cat); if (l.length && cat !== 'custom') code.value = l[0]; else fromOpts(); fillList(); upd(); if (['number', 'percent', 'currency', 'sci'].includes(cat)) fromOpts(); });
  fmtList.addEventListener('change', () => { code.value = fmtList.value; upd(); });
  optBox.addEventListener('input', fromOpts); optBox.addEventListener('change', fromOpts);
  code.addEventListener('input', upd);
  fillList(); upd();
  const numTab = h('div', { class: 'cf-num' }, h('div', { class: 'cf-numcols' }, field('Categoría', catSel), h('div', { class: 'fx-col' }, field('Formato', fmtList), optBox)), field('Código de formato', code, 'wide'), prev);
  // --- Fuente
  const fontTab = h('div', { class: 'fx-grid' },
    field('Tipo de letra', h('input', { name: 'font', value: st.font || 'Calibri', list: 'cf-fonts' })), h('datalist', { id: 'cf-fonts' }, ...(C.FONTS || []).map((f) => h('option', { value: f }))),
    field('Estilo', sel('fstyle', st.b && st.i ? 'bi' : st.b ? 'b' : st.i ? 'i' : 'n', [['n', 'Normal'], ['b', 'Negrita'], ['i', 'Cursiva'], ['bi', 'Negrita cursiva']])),
    field('Tamaño', numIn('size', st.size || 11, { step: 0.5, min: 4, max: 400, unit: 'pt' })), field('Color de fuente', h('input', { type: 'color', name: 'color', value: st.color || '#000000' })),
    h('h4', {}, t('Efectos')), chk('u', 'Subrayado', st.u), chk('s', 'Tachado', st.s));
  // --- Alineación
  const rotVal = st.rot === 'v' ? 0 : Number(st.rot || 0);
  const alignTab = h('div', { class: 'fx-grid' },
    h('h4', {}, t('Alineación del texto')),
    field('Horizontal', sel('align', st.align || '', [['', 'Predeterminado'], ['left', 'Izquierda'], ['center', 'Centro'], ['right', 'Derecha'], ['justify', 'Justificado'], ['fill', 'Rellenado']])),
    field('Sangría', numIn('indent', st.indent || 0, { step: 1, min: 0, max: 15 })),
    field('Vertical', sel('valign', st.valign || 'bottom', [['top', 'Superior'], ['middle', 'Centro'], ['bottom', 'Inferior']])), h('span'),
    h('h4', {}, t('Orientación del texto')),
    field('Grados', numIn('rot', rotVal, { step: 5, min: -90, max: 90, unit: '°' })), chk('stack', 'Apilado verticalmente', st.rot === 'v'),
    h('h4', {}, t('Propiedades')), chk('wrap', 'Ajustar texto automáticamente', st.wrap), chk('shrink', 'Reducir para ajustar al tamaño de la celda', st.shrink));
  // --- Bordes
  const b0 = st.border || {}; const first = Object.values(b0)[0] || '#000000 thin'; const [c0, s0 = 'thin'] = String(first).split(' ');
  const borderTab = h('div', { class: 'fx-grid' },
    field('Disposición', sel('bp', '', BORDER_PRESETS), 'wide'),
    h('h4', {}, t('Lados (para «Personalizado»)')), ...SIDES.map(([k, l]) => chk(`side_${k}`, l, !!b0[k])),
    h('h4', {}, t('Línea')), field('Estilo', sel('bstyle', s0, LINE_STYLES)), field('Color', h('input', { type: 'color', name: 'bcolor', value: c0 })));
  // --- Fondo
  const bgTab = h('div', { class: 'fx-grid' }, field('Color de fondo', h('input', { type: 'color', name: 'bg', value: st.bg || '#ffffff' })), chk('nobg', 'Sin relleno', !st.bg));
  // --- Protección
  const protTab = h('div', { class: 'fx-col' }, chk('locked', 'Protegida (bloqueada)', !st.unlocked), chk('hidef', 'Ocultar fórmula', st.hidef),
    h('p', { class: 'fx-note' }, t('La protección solo se aplica cuando la hoja está protegida (Revisar → Proteger hoja).')));
  const body = await tabbed('Formato de celdas', [
    { label: 'Números', body: numTab }, { label: 'Fuente', body: fontTab }, { label: 'Alineación', body: alignTab },
    { label: 'Bordes', body: borderTab }, { label: 'Fondo', body: bgTab }, { label: 'Protección', body: protTab },
  ], { start, width: '620px' });
  if (!body) return;
  const q = (n) => body.querySelector(`[name="${n}"]`);
  const fmt = q('code').value.trim();
  const fs = q('fstyle').value; const font = q('font').value.trim();
  const patch = {
    fmt: !fmt || fmt === 'General' ? undefined : fmt,
    font: !font || font === 'Calibri' ? undefined : font, b: fs === 'b' || fs === 'bi' || undefined, i: fs === 'i' || fs === 'bi' || undefined,
    size: Number(q('size').value) === 11 ? undefined : Number(q('size').value) || undefined, color: q('color').value.toLowerCase() === '#000000' ? undefined : q('color').value,
    u: q('u').checked || undefined, s: q('s').checked || undefined,
    align: q('align').value || undefined, indent: Number(q('indent').value) || undefined, valign: q('valign').value === 'bottom' ? undefined : q('valign').value,
    rot: q('stack').checked ? 'v' : Number(q('rot').value) || undefined, wrap: q('wrap').checked || undefined, shrink: q('shrink').checked || undefined,
    bg: q('nobg').checked ? undefined : q('bg').value, unlocked: q('locked').checked ? undefined : true, hidef: q('hidef').checked || undefined,
  };
  const bp = q('bp').value; const line = `${q('bcolor').value} ${q('bstyle').value}`.replace(/ thin$/, '');
  C.act('Formato de celdas', () => {
    C.wb.styleRange(g, patch);
    if (bp) applyBorderPreset(g, bp, line, Object.fromEntries(SIDES.map(([k]) => [k, q(`side_${k}`).checked])));
  }, { hf: false });
}
// bordes por disposición sobre un rango (como el cuadro de LibreOffice)
export function applyBorderPreset(g, preset, line = '#000000', sides = {}) {
  const wb = C.wb;
  for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
    const cur = { ...(wb.sheet.style(r, c)?.border || {}) };
    const edge = { t: r === g.r1, b: r === g.r2, l: c === g.c1, r: c === g.c2 };
    const set = (k, on) => { if (on) cur[k] = line; else delete cur[k]; };
    for (const k of ['t', 'b', 'l', 'r']) {
      const horiz = k === 't' || k === 'b';
      if (preset === 'none') set(k, false);
      else if (preset === 'all') set(k, true);
      else if (preset === 'outer') set(k, edge[k]);
      else if (preset === 'outerH') set(k, edge[k] || horiz);
      else if (preset === 'inner') set(k, !edge[k]);
      else if (preset === 'bottom') { if (k === 'b' && edge.b) set(k, true); }
      else if (preset === 'custom') { if (sides[k] && edge[k]) set(k, true); else if (!sides[k] && edge[k]) set(k, false); }
    }
    wb.setStyle(r, c, { border: Object.keys(cur).length ? cur : undefined });
  }
}

// ================================================================ pegado especial
export async function pasteSpecial() {
  const clip = C.grid.clip;
  if (!clip) { toast('Primero copia celdas de esta hoja de cálculo (Ctrl+C).'); return; }
  const v = await form('Pegado especial', [
    { name: 'what', label: 'Pegar', type: 'select', value: 'all', options: [['all', t('Todo')], ['values', t('Solo valores')], ['formulas', t('Solo fórmulas')], ['formats', t('Solo formatos')], ['valfmt', t('Valores y formatos')]] },
    { name: 'op', label: 'Operación', type: 'select', value: '', options: [['', t('Ninguna')], ['add', t('Sumar')], ['sub', t('Restar')], ['mul', t('Multiplicar')], ['div', t('Dividir')]] },
    { name: 'skip', label: 'Omitir celdas vacías', type: 'checkbox', value: false },
    { name: 'transpose', label: 'Transponer', type: 'checkbox', value: false },
  ], 'Pegar');
  if (v) doPasteSpecial(v);
}
export function doPasteSpecial({ what = 'all', op = '', skip = false, transpose = false }) {
  const clip = C.grid.clip; if (!clip) return;
  const wb = C.wb; const src = clip.range; const S = clip.sheet; const D = wb.sheet;
  const H = src.r2 - src.r1 + 1, W = src.c2 - src.c1 + 1;
  const dst = { r: C.sel().r1, c: C.sel().c1 };
  const cells = [];
  for (let r = 0; r < H; r++) for (let c = 0; c < W; c++) {
    const sr = src.r1 + r, sc = src.c1 + c;
    cells.push({ r, c, raw: wb.raw(sr, sc, S), val: wb.value(sr, sc, S), st: S.style(sr, sc), sr, sc });
  }
  C.act('Pegado especial', () => {
    for (const x of cells) {
      const tr = dst.r + (transpose ? x.c : x.r); const tc = dst.c + (transpose ? x.r : x.c);
      if (skip && (x.raw == null || x.raw === '')) continue;
      const formula = typeof x.raw === 'string' && x.raw.startsWith('=');
      if (what !== 'formats') {
        let nv;
        if (what === 'formulas' || what === 'all') {
          if (formula && !op && !transpose) { wb.hf.copy({ start: { sheet: S.id, row: x.sr, col: x.sc }, end: { sheet: S.id, row: x.sr, col: x.sc } }); wb.hf.paste({ sheet: D.id, row: tr, col: tc }); nv = undefined; }
          else nv = formula ? x.raw : x.raw;
          if (what === 'formulas' && !formula) nv = undefined;
        } else nv = wb.isError(x.val) ? x.raw : x.val;
        if (nv !== undefined) {
          if (op && typeof nv === 'number') { const cur = wb.value(tr, tc); const a = typeof cur === 'number' ? cur : 0; nv = op === 'add' ? a + nv : op === 'sub' ? a - nv : op === 'mul' ? a * nv : nv === 0 ? a : a / nv; }
          wb.setRaw(tr, tc, nv ?? null);
        }
      }
      if (what === 'all' || what === 'formats' || what === 'valfmt') { if (x.st) D.styles.set(`${tr},${tc}`, x.st); else D.styles.delete(`${tr},${tc}`); }
    }
    wb.grow(dst.r + (transpose ? W : H), dst.c + (transpose ? H : W));
  }, { sheets: S === D ? [D] : [D, S] });
  C.grid.selectRange({ r1: dst.r, c1: dst.c, r2: dst.r + (transpose ? W : H) - 1, c2: dst.c + (transpose ? H : W) - 1 });
}

// ================================================================ texto en columnas
export async function textToColumns() {
  const g = C.clamp(C.sel()); const wb = C.wb;
  const c = g.c1; const lines = []; for (let r = g.r1; r <= g.r2; r++) lines.push(String(wb.raw(r, c) ?? ''));
  if (!lines.some(Boolean)) { toast('Selecciona una columna con texto.'); return; }
  const guess = lines.some((l) => l.includes('\t')) ? 'tab' : lines.some((l) => l.includes(';')) ? ';' : lines.some((l) => l.includes(',')) ? ',' : ' ';
  const v = await form('Texto en columnas', [
    { name: 'sep', label: 'Separado por', type: 'select', value: guess, options: [['tab', t('Tabulación')], [';', t('Punto y coma')], [',', t('Coma')], [' ', t('Espacio')], ['other', t('Otro')]] },
    { name: 'other', label: 'Otro separador', value: '|' },
    { name: 'merge', label: 'Unir delimitadores', type: 'checkbox', value: guess === ' ' },
    { name: 'quote', label: 'Las comillas delimitan texto', type: 'checkbox', value: true },
  ], 'Aceptar');
  if (!v) return;
  splitColumn(g, v.sep === 'other' ? v.other || '|' : v.sep === 'tab' ? '\t' : v.sep, { merge: v.merge, quote: v.quote });
}
export function splitColumn(g, sep, { merge = false, quote = true } = {}) {
  const wb = C.wb; const c = g.c1;
  const split = (s) => {
    const out = []; let cur = ''; let q = false;
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (quote && ch === '"') { if (q && s[i + 1] === '"') { cur += '"'; i++; } else q = !q; continue; }
      if (!q && s.startsWith(sep, i)) { out.push(cur); cur = ''; i += sep.length - 1; if (merge) while (s.startsWith(sep, i + 1)) i += sep.length; continue; }
      cur += ch;
    }
    out.push(cur); return out;
  };
  const raws = []; for (let r = g.r1; r <= g.r2; r++) raws.push([r, wb.raw(r, c)]);
  C.act('Texto en columnas', () => {
    for (const [r, raw] of raws) {
      if (raw == null || (typeof raw === 'string' && raw.startsWith('='))) continue;
      split(String(raw)).forEach((part, j) => wb.setInput(r, c + j, part.trim()));
    }
  }, { batch: true });
}

// ================================================================ insertar y eliminar celdas
export async function insertCellsDialog() {
  const v = await form('Insertar celdas', [{ name: 'how', label: 'Selección', type: 'select', value: 'down', options: [['down', t('Desplazar celdas hacia abajo')], ['right', t('Desplazar celdas a la derecha')], ['rows', t('Insertar filas completas')], ['cols', t('Insertar columnas completas')]] }], 'Aceptar');
  if (v) shiftCells(v.how, true);
}
export async function deleteCellsDialog() {
  const v = await form('Eliminar celdas', [{ name: 'how', label: 'Selección', type: 'select', value: 'up', options: [['up', t('Desplazar celdas hacia arriba')], ['left', t('Desplazar celdas a la izquierda')], ['rows', t('Eliminar filas completas')], ['cols', t('Eliminar columnas completas')]] }], 'Aceptar');
  if (v) shiftCells(v.how, false);
}
export function shiftCells(how, insert) {
  const wb = C.wb; const s = wb.sheet; const g = C.sel();
  if (how === 'rows') return insert ? C.act('Insertar filas', () => wb.insertRows(g.r1, g.r2 - g.r1 + 1)) : C.act('Eliminar filas', () => wb.deleteRows(g.r1, g.r2 - g.r1 + 1));
  if (how === 'cols') return insert ? C.act('Insertar columnas', () => wb.insertCols(g.c1, g.c2 - g.c1 + 1)) : C.act('Eliminar columnas', () => wb.deleteCols(g.c1, g.c2 - g.c1 + 1));
  const dims = wb.hf.getSheetDimensions(s.id); const lastR = Math.max(dims.height, g.r2 + 1) - 1; const lastC = Math.max(dims.width, g.c2 + 1) - 1;
  const n = how === 'down' || how === 'up' ? g.r2 - g.r1 + 1 : g.c2 - g.c1 + 1;
  const mv = (r1, c1, r2, c2, dr, dc) => { if (r2 < r1 || c2 < c1) return; wb.hf.moveCells({ start: { sheet: s.id, row: r1, col: c1 }, end: { sheet: s.id, row: r2, col: c2 } }, { sheet: s.id, row: r1 + dr, col: c1 + dc }); };
  // formatos: se mueven con las celdas
  const moveStyles = (r1, c1, r2, c2, dr, dc) => {
    const keep = []; for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) { const k = `${r},${c}`; keep.push([r + dr, c + dc, s.styles.get(k)]); s.styles.delete(k); }
    for (const [r, c, st] of keep) if (st) s.styles.set(`${r},${c}`, st);
  };
  C.act(insert ? 'Insertar celdas' : 'Eliminar celdas', () => {
    if (insert && how === 'down') { wb.grow(lastR + n + 1, g.c2 + 1); mv(g.r1, g.c1, lastR, g.c2, n, 0); moveStyles(g.r1, g.c1, lastR, g.c2, n, 0); }
    if (insert && how === 'right') { wb.grow(g.r2 + 1, lastC + n + 1); mv(g.r1, g.c1, g.r2, lastC, 0, n); moveStyles(g.r1, g.c1, g.r2, lastC, 0, n); }
    if (!insert && how === 'up') {
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { wb.setRaw(r, c, null); s.styles.delete(`${r},${c}`); }
      mv(g.r2 + 1, g.c1, lastR, g.c2, -n, 0); moveStyles(g.r2 + 1, g.c1, lastR, g.c2, -n, 0);
    }
    if (!insert && how === 'left') {
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { wb.setRaw(r, c, null); s.styles.delete(`${r},${c}`); }
      mv(g.r1, g.c2 + 1, g.r2, lastC, 0, -n); moveStyles(g.r1, g.c2 + 1, g.r2, lastC, 0, -n);
    }
  });
}

// ================================================================ mostrar fórmulas
export function toggleFormulas() {
  C.wb.showFormulas = !C.wb.showFormulas; C.shell.ribbon.setOn('showFormulas', C.wb.showFormulas);
  C.grid.invalidate(); C.grid.draw();
}

// ================================================================ color de pestaña
export const TAB_COLORS = ['#C62828', '#EF6C00', '#F9A825', '#2E7D32', '#00838F', '#1565C0', '#6A1B9A', '#AD1457', '#5D4037', '#546E7A'];
export function setTabColor(i, color) {
  const s = C.wb.sheets[i]; C.act('Color de pestaña', () => { X3(s).tabColor = color || undefined; }, { hf: false, sheets: [s] });
  C.renderTabs(true);
}
export function tabColorPicker(anchor, i) {
  const box = h('div', { class: 'tc-pick' }, ...TAB_COLORS.map((col) => h('button', { class: 'tc-sw', title: col, style: { background: col }, 'aria-label': col, onclick: () => { closeDropdown(); setTabColor(i, col); } })),
    h('button', { class: 'mi tc-none', onclick: () => { closeDropdown(); setTabColor(i, null); } }, t('Sin color')));
  dropdown(anchor, box);
}

// ================================================================ hipervínculos
export async function linkDialog() {
  const { r, c } = C.grid.sel; const x = X3(); const cur = x.links?.[`${r},${c}`] || '';
  const text = C.wb.display(r, c).text;
  const v = await form('Hipervínculo', [
    { name: 'url', label: 'Dirección (web, correo o celda como #Hoja1!A1)', value: cur || 'https://' },
    { name: 'text', label: 'Texto que se muestra', value: text },
  ], cur ? 'Actualizar' : 'Insertar', cur ? { extra: [{ label: 'Quitar vínculo', value: 'remove' }] } : undefined);
  if (!v) return;
  setLink(r, c, v === 'remove' ? null : v.url.trim(), v.text);
}
export function setLink(r, c, url, text) {
  const s = C.wb.sheet;
  C.act('Hipervínculo', () => {
    const x = X3(s); x.links = { ...(x.links || {}) };
    if (url && url !== 'https://') { x.links[`${r},${c}`] = url; if (text != null && text !== C.wb.display(r, c).text) C.wb.setInput(r, c, text || url); else if (C.wb.raw(r, c) == null) C.wb.setInput(r, c, url); }
    else delete x.links[`${r},${c}`];
  });
}
export function followLink(r, c) {
  const url = X3().links?.[`${r},${c}`]; if (!url) return false;
  const m = /^#'?([^'!]+)'?!([A-Z]+\d+)$/i.exec(url);
  if (m) { const i = C.wb.sheets.findIndex((s) => s.name.toLowerCase() === m[1].toLowerCase()); if (i >= 0) { C.switchSheet(i); const g = parseRange(m[2]); C.grid.select(g.r1, g.c1); } return true; }
  if (/^(https?:|mailto:)/i.test(url)) { window.oficina?.openExternal?.(url) ?? window.open(url, '_blank'); return true; }
  return false;
}

// ================================================================ rastrear precedentes y dependientes
let arrows = []; // { from:{r,c}, to:{r,c}, sheet, kind }
export function trace(kind) {
  const wb = C.wb; const s = wb.sheet; const { r, c } = C.grid.sel;
  let list = [];
  try { list = kind === 'pre' ? wb.hf.getCellPrecedents({ sheet: s.id, row: r, col: c }) : wb.hf.getCellDependents({ sheet: s.id, row: r, col: c }); } catch { list = []; }
  const cells = [];
  for (const a of list) {
    if (a.start) { if (a.start.sheet !== s.id) continue; cells.push({ r: a.start.row, c: a.start.col, r2: a.end.row, c2: a.end.col }); }
    else if (a.sheet === s.id) cells.push({ r: a.row, c: a.col });
  }
  if (!cells.length) { toast(kind === 'pre' ? 'Esta celda no tiene precedentes.' : 'Ninguna celda depende de esta.'); return; }
  for (const x of cells) arrows.push(kind === 'pre' ? { from: x, to: { r, c }, sheet: s.id } : { from: { r, c }, to: x, sheet: s.id });
  C.grid.draw();
}
export function clearArrows() { arrows = []; C.grid.draw(); }
export const arrowCount = () => arrows.length;
function drawArrows(ctx) {
  if (!arrows.length) return;
  const s = C.wb.sheet; ctx.save(); ctx.strokeStyle = '#1565C0'; ctx.fillStyle = '#1565C0'; ctx.lineWidth = 1.5;
  for (const a of arrows) {
    if (a.sheet !== s.id) continue;
    const f = C.grid.cellRect(a.from.r, a.from.c); const tt = C.grid.cellRect(a.to.r, a.to.c);
    if (a.from.r2 != null) { const f2 = C.grid.cellRect(a.from.r2, a.from.c2); ctx.strokeRect(f.x + 1, f.y + 1, f2.x + f2.w - f.x - 2, f2.y + f2.h - f.y - 2); }
    const x1 = f.x + f.w / 2, y1 = f.y + f.h / 2, x2 = tt.x + tt.w / 2, y2 = tt.y + tt.h / 2;
    ctx.beginPath(); ctx.arc(x1, y1, 3, 0, Math.PI * 2); ctx.fill();
    ctx.beginPath(); ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); ctx.stroke();
    const ang = Math.atan2(y2 - y1, x2 - x1);
    ctx.beginPath(); ctx.moveTo(x2, y2); ctx.lineTo(x2 - 9 * Math.cos(ang - 0.4), y2 - 9 * Math.sin(ang - 0.4)); ctx.lineTo(x2 - 9 * Math.cos(ang + 0.4), y2 - 9 * Math.sin(ang + 0.4)); ctx.closePath(); ctx.fill();
  }
  ctx.restore();
}

// ================================================================ subtotales
const SUBFN = [['9', 'Suma'], ['1', 'Promedio'], ['2', 'Cuenta (números)'], ['3', 'Cuenta'], ['4', 'Máximo'], ['5', 'Mínimo'], ['6', 'Producto']];
export async function subtotalsDialog() {
  const wb = C.wb; let g = C.clamp(C.sel());
  if (g.r1 === g.r2) g = currentRegion(g.r1, g.c1);
  const heads = []; for (let c = g.c1; c <= g.c2; c++) heads.push([String(c), `${colName(c)} — ${String(wb.display(g.r1, c).text || '').slice(0, 24)}`]);
  const numCols = heads.filter(([c]) => typeof wb.value(g.r1 + 1, Number(c)) === 'number').map(([c]) => c);
  const v = await form('Subtotales', [
    { name: 'by', label: 'Agrupar por', type: 'select', value: String(g.c1), options: heads },
    { name: 'fn', label: 'Usar función', type: 'select', value: '9', options: SUBFN.map(([k, l]) => [k, t(l)]) },
    { name: 'cols', label: 'Calcular subtotales para (columnas, p. ej. C,D)', value: numCols.map((c) => colName(Number(c))).join(',') || colName(g.c2) },
    { name: 'sort', label: 'Ordenar primero por el grupo', type: 'checkbox', value: true },
  ], 'Aceptar');
  if (!v) return;
  const cols = v.cols.split(/[,; ]+/).filter(Boolean).map((x) => x.toUpperCase().charCodeAt(0) - 65 + (x.length > 1 ? 26 * (x.toUpperCase().charCodeAt(0) - 64) - (x.toUpperCase().charCodeAt(0) - 65) + x.toUpperCase().charCodeAt(1) - 65 - (x.toUpperCase().charCodeAt(0) - 65) : 0));
  subtotals(g, Number(v.by), v.fn, cols, v.sort);
}
function currentRegion(r, c) {
  const wb = C.wb; const has = (rr, cc) => rr >= 0 && cc >= 0 && wb.raw(rr, cc) != null && wb.raw(rr, cc) !== '';
  let g = { r1: r, c1: c, r2: r, c2: c }; let grow = true;
  while (grow) { grow = false;
    if (g.r1 > 0 && [...Array(g.c2 - g.c1 + 1).keys()].some((k) => has(g.r1 - 1, g.c1 + k))) { g.r1--; grow = true; }
    if ([...Array(g.c2 - g.c1 + 1).keys()].some((k) => has(g.r2 + 1, g.c1 + k))) { g.r2++; grow = true; }
    if (g.c1 > 0 && [...Array(g.r2 - g.r1 + 1).keys()].some((k) => has(g.r1 + k, g.c1 - 1))) { g.c1--; grow = true; }
    if ([...Array(g.r2 - g.r1 + 1).keys()].some((k) => has(g.r1 + k, g.c2 + 1))) { g.c2++; grow = true; }
  }
  return g;
}
export function subtotals(g, by, fn, cols, sort = true) {
  const wb = C.wb; const s = wb.sheet;
  const W = g.c2 - g.c1 + 1;
  const head = []; for (let c = g.c1; c <= g.c2; c++) head.push(wb.raw(g.r1, c));
  const rows = []; for (let r = g.r1 + 1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) row.push({ raw: wb.raw(r, c), st: s.style(r, c) }); rows.push(row); }
  const key = (row) => String(row[by - g.c1].raw ?? '');
  if (sort) rows.sort((a, b) => key(a).localeCompare(key(b), undefined, { numeric: true }));
  const fnName = SUBFN.find((f) => f[0] === fn)?.[1] || 'Suma';
  const out = []; const groups = []; let start = 0;
  const pushTotal = (label, from, to, grand) => {
    const row = Array(W).fill(null).map(() => ({ raw: null, st: { b: true } }));
    row[by - g.c1] = { raw: grand ? `${t('Total general')}` : `${t('Total')} ${label}`, st: { b: true } };
    // total general: los rangos de datos de cada grupo (sin los subtotales), válido en cualquier hoja de cálculo
    for (const c of cols) if (c >= g.c1 && c <= g.c2) row[c - g.c1] = { raw: `=SUBTOTAL(${fn},${grand ? groups.map((gg) => `${colName(c)}${gg.a + 1}:${colName(c)}${gg.b + 1}`).join(',') : `${colName(c)}${from + 1}:${colName(c)}${to + 1}`})`, st: { b: true } };
    out.push({ row, total: true });
  };
  let r0 = g.r1 + 1; const totalRows = [];
  rows.forEach((row, i) => {
    out.push({ row });
    const next = rows[i + 1];
    if (!next || key(next) !== key(row)) {
      const last = g.r1 + out.length; // fila (base 0) del último dato del grupo
      pushTotal(key(row), r0, last); totalRows.push(g.r1 + out.length);
      groups.push({ a: r0, b: last }); r0 = g.r1 + out.length + 1;
    }
  });
  pushTotal('', g.r1 + 1, g.r1 + out.length, true);

  void start; void fnName;
  const extra = out.length - rows.length;
  C.act('Subtotales', () => {
    if (extra > 0) wb.insertRows(g.r2 + 1, extra);
    out.forEach((o, i) => o.row.forEach((cell, j) => { wb.setRaw(g.r1 + 1 + i, g.c1 + j, cell.raw ?? null); if (cell.st) s.styles.set(`${g.r1 + 1 + i},${g.c1 + j}`, Object.freeze({ ...cell.st })); else s.styles.delete(`${g.r1 + 1 + i},${g.c1 + j}`); }));
    // el total general de SUBTOTAL ignora los subtotales anidados (como en LibreOffice)
    const x = X3(s); x.rowGroups = [...(x.rowGroups || []), ...groups.map((gg) => ({ a: gg.a, b: gg.b }))];
  });
  void head;
}

// ================================================================ filtro estándar
const OPS = [['=', '='], ['<>', '<>'], ['>', '>'], ['<', '<'], ['>=', '>='], ['<=', '<='], ['contains', 'Contiene'], ['notcontains', 'No contiene'], ['begins', 'Comienza por'], ['ends', 'Termina en']];
export async function standardFilter() {
  const wb = C.wb; let g = C.clamp(C.sel()); if (g.r1 === g.r2) g = currentRegion(g.r1, g.c1);
  const cols = []; for (let c = g.c1; c <= g.c2; c++) cols.push([String(c), `${colName(c)} — ${String(wb.display(g.r1, c).text || '').slice(0, 24)}`]);
  const v = await form('Filtro estándar', [
    { name: 'c1', label: 'Campo', type: 'select', value: String(g.c1), options: cols }, { name: 'o1', label: 'Condición', type: 'select', value: '=', options: OPS.map(([k, l]) => [k, t(l)]) }, { name: 'v1', label: 'Valor', value: '' },
    { name: 'join', label: 'Y / O', type: 'select', value: '', options: [['', t('(sin segunda condición)')], ['and', t('Y')], ['or', t('O')]] },
    { name: 'c2', label: 'Campo 2', type: 'select', value: String(g.c1), options: cols }, { name: 'o2', label: 'Condición 2', type: 'select', value: '=', options: OPS.map(([k, l]) => [k, t(l)]) }, { name: 'v2', label: 'Valor 2', value: '' },
    { name: 'case', label: 'Distinguir mayúsculas y minúsculas', type: 'checkbox', value: false },
  ], 'Filtrar', { extra: [{ label: 'Quitar filtro', value: 'clear' }] });
  if (!v) return;
  if (v === 'clear') return clearStdFilter(g);
  const conds = [{ c: Number(v.c1), op: v.o1, v: v.v1 }]; if (v.join) conds.push({ c: Number(v.c2), op: v.o2, v: v.v2 });
  applyStdFilter(g, conds, v.join || 'and', v.case);
}
export function applyStdFilter(g, conds, join = 'and', cs = false) {
  const wb = C.wb; const s = wb.sheet;
  const test = (r, { c, op, v }) => {
    const val = wb.value(r, c); const n = Number(String(v).replace(',', '.')); const isNum = v !== '' && !Number.isNaN(n) && typeof val === 'number';
    const A = cs ? String(wb.display(r, c).text) : String(wb.display(r, c).text).toLowerCase(); const B = cs ? String(v) : String(v).toLowerCase();
    switch (op) {
      case '=': return isNum ? val === n : A === B; case '<>': return isNum ? val !== n : A !== B;
      case '>': return isNum ? val > n : A > B; case '<': return isNum ? val < n : A < B; case '>=': return isNum ? val >= n : A >= B; case '<=': return isNum ? val <= n : A <= B;
      case 'contains': return A.includes(B); case 'notcontains': return !A.includes(B); case 'begins': return A.startsWith(B); case 'ends': return A.endsWith(B); default: return true;
    }
  };
  C.act('Filtro estándar', () => {
    for (let r = g.r1 + 1; r <= g.r2; r++) { const ok = join === 'or' ? conds.some((x) => test(r, x)) : conds.every((x) => test(r, x)); if (ok) s.hiddenRows.delete(r); else s.hiddenRows.add(r); }
    X3(s).stdFilter = { range: rangeStr(g) };
  }, { hf: false });
}
export function clearStdFilter(g) {
  const s = C.wb.sheet; const x = X3(s); const rg = x.stdFilter ? parseRange(x.stdFilter.range) : g;
  C.act('Quitar filtro', () => { for (let r = rg.r1 + 1; r <= rg.r2; r++) s.hiddenRows.delete(r); delete x.stdFilter; }, { hf: false });
}

// ================================================================ autocompletar (como LibreOffice: texto que ya está en la columna)
export function autoComplete(r, c, typed) {
  if (!typed || typed.startsWith('=') || /^[\d.,\-+]/.test(typed)) return null;
  const wb = C.wb; const low = typed.toLowerCase(); const seen = new Set();
  const dims = wb.hf.getSheetDimensions(wb.sheet.id);
  for (let k = 1; k < Math.min(dims.height + 1, 5000); k++) {
    for (const rr of [r - k, r + k]) {
      if (rr < 0 || rr >= dims.height) continue;
      const v = wb.value(rr, c); if (typeof v !== 'string' || v.length <= typed.length) continue;
      if (v.toLowerCase().startsWith(low) && !seen.has(v)) return v;
      seen.add(v);
    }
  }
  return null;
}

void addr; void tf;
