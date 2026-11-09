// Calcula Datos — fase 2: formato condicional, validación, nombres, notas, tabla dinámica,
// formato de tabla, protección y configuración de impresión. GPL-3.0
import { h, icon, toast, modal, form, dropdown, closeDropdown, esc, confirmBox } from '../../shared/office.js';
import { parseRange, rangeStr, addr, colName } from './workbook.js';
import { parseInput, formatValue } from './format.js';

let C = null; // { wb, grid, shell, act, refresh, switchSheet, sel }
let version = 0;
export function initFeatures(ctx) { C = ctx; C.wb.onChange(() => { version++; }); }
const uid = () => Math.random().toString(36).slice(2, 9);
const inRange = (g, r, c) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2;

// ================================================================ formato condicional
export const CF_LABELS = {
  gt: 'Mayor que', lt: 'Menor que', between: 'Entre', eq: 'Igual a', text: 'Texto que contiene', dup: 'Valores duplicados',
  top: '10 superiores', bottom: '10 inferiores', above: 'Por encima del promedio', below: 'Por debajo del promedio',
  bar: 'Barras de datos', scale2: 'Escala de 2 colores', scale3: 'Escala de 3 colores', icons: 'Conjunto de íconos',
};
export function addRule(rule) {
  const s = C.wb.sheet;
  const r = { id: uid(), ...rule };
  C.act('Formato condicional', () => { s.cf = [...(s.cf || []), r]; }, { hf: false });
  return r;
}
const statsCache = new Map();
function ruleStats(rule, s) {
  const key = `${s.id}|${rule.id}|${version}`;
  if (statsCache.has(key)) return statsCache.get(key);
  const g = parseRange(rule.range); const nums = []; const texts = new Map();
  if (g) for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 20000); r++) for (let c = g.c1; c <= g.c2; c++) {
    const v = C.wb.value(r, c, s);
    if (typeof v === 'number') nums.push(v);
    if (v != null && v !== '') { const k = String(v).toLowerCase(); texts.set(k, (texts.get(k) || 0) + 1); }
  }
  const sorted = [...nums].sort((a, b) => a - b);
  const st = { min: sorted[0] ?? 0, max: sorted[sorted.length - 1] ?? 0, avg: nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0, sorted, texts };
  if (statsCache.size > 400) statsCache.clear();
  statsCache.set(key, st);
  return st;
}
const mix = (a, b, t) => { const p = (x) => [1, 3, 5].map((i) => parseInt(x.slice(i, i + 2), 16)); const A = p(a), B = p(b); return `#${A.map((v, i) => Math.round(v + (B[i] - v) * t).toString(16).padStart(2, '0')).join('')}`; };
export function cfStyle(r, c, s = C?.wb.sheet) {
  if (!s?.cf?.length) return null;
  let out = null;
  for (const rule of s.cf) {
    const g = parseRange(rule.range); if (!g || !inRange(g, r, c)) continue;
    const v = C.wb.value(r, c, s);
    const num = typeof v === 'number';
    let hit = false;
    const st = ruleStats(rule, s);
    switch (rule.type) {
      case 'gt': hit = num && v > Number(rule.value); break;
      case 'lt': hit = num && v < Number(rule.value); break;
      case 'between': hit = num && v >= Math.min(Number(rule.value), Number(rule.value2)) && v <= Math.max(Number(rule.value), Number(rule.value2)); break;
      case 'eq': hit = num ? v === Number(rule.value) : String(v ?? '').toLowerCase() === String(rule.value ?? '').toLowerCase(); break;
      case 'text': hit = v != null && String(v).toLowerCase().includes(String(rule.text || '').toLowerCase()); break;
      case 'dup': hit = v != null && v !== '' && st.texts.get(String(v).toLowerCase()) > 1; break;
      case 'top': { const n = Number(rule.n || 10); hit = num && st.sorted.length && v >= st.sorted[Math.max(0, st.sorted.length - n)]; break; }
      case 'bottom': { const n = Number(rule.n || 10); hit = num && st.sorted.length && v <= st.sorted[Math.min(st.sorted.length - 1, n - 1)]; break; }
      case 'above': hit = num && v > st.avg; break;
      case 'below': hit = num && v < st.avg; break;
      case 'bar': if (num && st.max !== st.min) { out = { ...(out || {}), bar: { pct: (v - Math.min(0, st.min)) / (st.max - Math.min(0, st.min)), color: rule.color || '#5B9BD5' } }; } else if (num) out = { ...(out || {}), bar: { pct: 1, color: rule.color || '#5B9BD5' } }; break;
      case 'scale2': case 'scale3': if (num) {
        const cols = rule.colors || (rule.type === 'scale3' ? ['#F8696B', '#FFEB84', '#63BE7B'] : ['#FFFFFF', '#63BE7B']);
        const t = st.max === st.min ? 0.5 : (v - st.min) / (st.max - st.min);
        const bg = rule.type === 'scale3' ? (t < 0.5 ? mix(cols[0], cols[1], t * 2) : mix(cols[1], cols[2], (t - 0.5) * 2)) : mix(cols[0], cols[1], t);
        out = { ...(out || {}), bg };
      } break;
      case 'icons': if (num) { const t = st.max === st.min ? 1 : (v - st.min) / (st.max - st.min); out = { ...(out || {}), icon: t >= 0.67 ? 2 : t >= 0.33 ? 1 : 0, iconSet: rule.set || 'arrows' }; } break;
      default: break;
    }
    if (hit) out = { ...(out || {}), ...(rule.style || { bg: '#FFC7CE', color: '#9C0006' }) };
  }
  return out;
}
export async function cfDialog(type) {
  const g = C.sel(); const range = rangeStr(C.clamp(g));
  const presets = [['#FFC7CE|#9C0006', 'Relleno rojo claro con texto rojo oscuro'], ['#FFEB9C|#9C5700', 'Relleno amarillo con texto amarillo oscuro'], ['#C6EFCE|#006100', 'Relleno verde con texto verde oscuro'], ['#FFC7CE|', 'Relleno rojo claro'], ['|#C00000', 'Texto rojo'], ['#BDD7EE|#1F4E78', 'Relleno azul con texto azul oscuro']];
  const needs = { gt: ['value'], lt: ['value'], eq: ['value'], between: ['value', 'value2'], text: ['text'], top: ['n'], bottom: ['n'] }[type] || [];
  const fields = [{ name: 'range', label: 'Aplicar a', value: range }];
  if (needs.includes('value')) fields.push({ name: 'value', label: type === 'between' ? 'De' : 'Valor', value: '' });
  if (needs.includes('value2')) fields.push({ name: 'value2', label: 'Hasta', value: '' });
  if (needs.includes('text')) fields.push({ name: 'text', label: 'Texto que contiene', value: '' });
  if (needs.includes('n')) fields.push({ name: 'n', label: 'Cantidad de elementos', type: 'number', value: 10, min: 1 });
  if (!['bar', 'scale2', 'scale3', 'icons'].includes(type)) fields.push({ name: 'fmt', label: 'Con', type: 'select', value: presets[0][0], options: presets });
  if (type === 'bar') fields.push({ name: 'color', label: 'Color de la barra', type: 'select', value: '#5B9BD5', options: [['#5B9BD5', 'Azul'], ['#63BE7B', 'Verde'], ['#F8696B', 'Rojo'], ['#FFB628', 'Naranja'], ['#8E7FFF', 'Morado']] });
  if (type === 'icons') fields.push({ name: 'set', label: 'Íconos', type: 'select', value: 'arrows', options: [['arrows', 'Flechas (▲ ► ▼)'], ['lights', 'Semáforo (● ● ●)']] });
  const v = await form(CF_LABELS[type], fields, 'Aplicar');
  if (!v) return;
  if (!parseRange(v.range)) { toast('Rango no válido', { error: true }); return; }
  const rule = { range: v.range.toUpperCase(), type };
  if (v.value !== undefined) rule.value = v.value; if (v.value2 !== undefined) rule.value2 = v.value2; if (v.text !== undefined) rule.text = v.text; if (v.n !== undefined) rule.n = v.n;
  if (v.fmt) { const [bg, color] = v.fmt.split('|'); rule.style = { bg: bg || undefined, color: color || undefined }; }
  if (v.color) rule.color = v.color; if (v.set) rule.set = v.set;
  addRule(rule);
}
export function cfMenu() {
  const it = (t, ic) => ({ label: CF_LABELS[t], icon: ic, run: () => cfDialog(t) });
  return [
    { header: 'Resaltar reglas de celdas' }, it('gt', 'chevronDown'), it('lt', 'chevronDown'), it('between', 'alignCenter'), it('eq', 'hr'), it('text', 'find'), it('dup', 'copy'),
    { header: 'Reglas superiores e inferiores' }, it('top', 'sortDesc'), it('bottom', 'sortAsc'), it('above', 'chartLine'), it('below', 'chartLine'),
    '-', it('bar', 'chart'), it('scale2', 'fill'), it('scale3', 'fill'), it('icons', 'star'),
    '-', { label: 'Borrar reglas de la selección', icon: 'clear', run: () => clearRules(true) }, { label: 'Borrar reglas de toda la hoja', icon: 'trash', run: () => clearRules(false) }, { label: 'Administrar reglas…', icon: 'layout', run: manageRules },
  ];
}
function clearRules(onlySel) {
  const s = C.wb.sheet; const g = C.sel();
  C.act('Borrar reglas', () => { s.cf = onlySel ? (s.cf || []).filter((r) => { const rg = parseRange(r.range); return !(rg && rg.r1 <= g.r2 && rg.r2 >= g.r1 && rg.c1 <= g.c2 && rg.c2 >= g.c1); }) : []; }, { hf: false });
}
async function manageRules() {
  const s = C.wb.sheet;
  const list = h('div', { style: { minWidth: '460px', display: 'grid', gap: '6px' } });
  const render = () => {
    list.replaceChildren();
    if (!(s.cf || []).length) list.append(h('p', { style: { color: 'var(--ink-3)' } }, 'Esta hoja no tiene reglas de formato condicional.'));
    (s.cf || []).forEach((r) => {
      const sw = h('span', { style: { display: 'inline-block', width: '40px', height: '18px', borderRadius: '3px', background: r.style?.bg || r.color || (r.colors ? `linear-gradient(90deg,${r.colors.join(',')})` : 'linear-gradient(90deg,#F8696B,#FFEB84,#63BE7B)'), border: '1px solid var(--line)' } });
      list.append(h('div', { style: { display: 'flex', alignItems: 'center', gap: '10px', border: '1px solid var(--line)', borderRadius: '8px', padding: '6px 10px' } },
        sw, h('b', { style: { flex: 1 } }, `${CF_LABELS[r.type]}${r.value != null ? ` ${r.value}` : ''}${r.value2 != null ? ` y ${r.value2}` : ''}${r.text ? ` “${r.text}”` : ''}`), h('code', {}, r.range),
        h('button', { class: 'btn ghost', onclick: () => { C.act('Eliminar regla', () => { s.cf = s.cf.filter((x) => x.id !== r.id); }, { hf: false }); render(); } }, 'Eliminar')));
    });
  };
  render();
  await modal({ title: 'Administrar reglas de formato condicional', body: list });
}

// ================================================================ validación de datos
export function addValidation(v) { const s = C.wb.sheet; C.act('Validación', () => { s.validations = [...(s.validations || []).filter((x) => x.range !== v.range), { id: uid(), ...v }]; }, { hf: false }); }
export function validationAt(r, c, s = C.wb.sheet) { return (s.validations || []).find((v) => { const g = parseRange(v.range); return g && inRange(g, r, c); }) || null; }
export function validate(r, c, text) {
  const v = validationAt(r, c); if (!v || text === '' || text == null) return { ok: true };
  const p = parseInput(text); const val = p.value;
  const fail = (msg) => ({ ok: false, msg: v.message || msg });
  if (typeof val === 'string' && val.startsWith('=')) return { ok: true };
  if (v.type === 'list') return v.list.map((x) => String(x).toLowerCase()).includes(String(text).trim().toLowerCase()) ? { ok: true } : fail(`Elige un valor de la lista: ${v.list.join(', ')}`);
  if (v.type === 'whole' || v.type === 'decimal' || v.type === 'date') {
    if (typeof val !== 'number') return fail('Escribe un número.');
    if (v.type === 'whole' && !Number.isInteger(val)) return fail('Escribe un número entero.');
    const min = v.min !== '' && v.min != null ? (v.type === 'date' ? parseInput(String(v.min)).value : Number(v.min)) : -Infinity;
    const max = v.max !== '' && v.max != null ? (v.type === 'date' ? parseInput(String(v.max)).value : Number(v.max)) : Infinity;
    return val >= min && val <= max ? { ok: true } : fail(`El valor debe estar entre ${v.min ?? '—'} y ${v.max ?? '—'}.`);
  }
  if (v.type === 'length') { const n = String(text).length; return n >= (Number(v.min) || 0) && n <= (Number(v.max) || Infinity) ? { ok: true } : fail(`El texto debe tener entre ${v.min || 0} y ${v.max} caracteres.`); }
  return { ok: true };
}
export async function validationDialog() {
  const g = C.clamp(C.sel()); const cur = validationAt(g.r1, g.c1);
  const v = await form('Validación de datos', [
    { name: 'range', label: 'Aplicar a', value: cur?.range || rangeStr(g) },
    { name: 'type', label: 'Permitir', type: 'select', value: cur?.type || 'list', options: [['list', 'Lista de valores'], ['whole', 'Número entero'], ['decimal', 'Decimal'], ['date', 'Fecha'], ['length', 'Longitud del texto']] },
    { name: 'list', label: 'Valores de la lista (separados por comas)', value: cur?.list?.join(', ') || '' },
    { name: 'min', label: 'Mínimo', value: cur?.min ?? '' },
    { name: 'max', label: 'Máximo', value: cur?.max ?? '' },
    { name: 'message', label: 'Mensaje de error (opcional)', value: cur?.message || '' },
  ], 'Aplicar');
  if (!v) return;
  if (!parseRange(v.range)) { toast('Rango no válido', { error: true }); return; }
  const rule = { range: v.range.toUpperCase(), type: v.type, message: v.message.trim() || undefined };
  if (v.type === 'list') { rule.list = v.list.split(/[,;]/).map((x) => x.trim()).filter(Boolean); if (!rule.list.length) { toast('Escribe al menos un valor para la lista.', { error: true }); return; } }
  else { rule.min = v.min; rule.max = v.max; }
  addValidation(rule);
  toast('Validación aplicada');
}
export function clearValidation() { const s = C.wb.sheet; const g = C.sel(); C.act('Quitar validación', () => { s.validations = (s.validations || []).filter((v) => { const rg = parseRange(v.range); return !(rg && rg.r1 <= g.r2 && rg.r2 >= g.r1 && rg.c1 <= g.c2 && rg.c2 >= g.c1); }); }, { hf: false }); }

// ================================================================ nombres definidos
export function addName(name, expr) {
  const wb = C.wb;
  if (!/^[A-Za-zÁÉÍÓÚÑáéíóúñ_][\wÁÉÍÓÚÑáéíóúñ.]*$/.test(name) || /^[A-Z]{1,3}\d+$/i.test(name)) throw new Error('Nombre no válido: usa letras, números y guion bajo, sin espacios.');
  const e = expr.startsWith('=') ? expr : `=${expr}`;
  if (wb.hf.isItPossibleToAddNamedExpression(name, e)) wb.hf.addNamedExpression(name, e);
  else if (wb.hf.isItPossibleToChangeNamedExpression?.(name, e)) wb.hf.changeNamedExpression(name, e);
  else throw new Error('No se pudo definir ese nombre.');
  wb.names = [...(wb.names || []).filter((n) => n.name.toLowerCase() !== name.toLowerCase()), { name, expr: e }];
  version++;
}
export function removeName(name) { const wb = C.wb; try { wb.hf.removeNamedExpression(name); } catch { /* */ } wb.names = (wb.names || []).filter((n) => n.name !== name); version++; }
export async function defineName() {
  const g = C.clamp(C.sel());
  const s = C.wb.sheet;
  const abs = `${/^[A-Za-z_]\w*$/.test(s.name) ? s.name : `'${s.name}'`}!$${colName(g.c1)}$${g.r1 + 1}${g.r1 !== g.r2 || g.c1 !== g.c2 ? `:$${colName(g.c2)}$${g.r2 + 1}` : ''}`;
  const head = C.wb.value(g.r1 > 0 ? g.r1 - 1 : g.r1, g.c1);
  const v = await form('Definir nombre', [{ name: 'name', label: 'Nombre', value: typeof head === 'string' ? head.replace(/[^\wÁÉÍÓÚÑáéíóúñ]+/g, '_').replace(/^(\d)/, '_$1') : '' }, { name: 'expr', label: 'Se refiere a', value: abs }], 'Aceptar');
  if (!v) return;
  try { addName(v.name.trim(), v.expr.trim()); toast(`Nombre “${v.name}” definido. Úsalo en fórmulas, p. ej. =SUMA(${v.name}).`); C.refresh(); } catch (err) { toast(err.message, { error: true }); }
}
export async function manageNames() {
  const box = h('div', { style: { minWidth: '480px' } });
  const render = () => {
    box.replaceChildren();
    const names = C.wb.names || [];
    if (!names.length) box.append(h('p', { style: { color: 'var(--ink-3)' } }, 'No hay nombres definidos. Selecciona un rango y usa “Definir nombre”.'));
    names.forEach((n) => box.append(h('div', { style: { display: 'flex', gap: '10px', alignItems: 'center', padding: '6px 0', borderBottom: '1px solid var(--line)' } },
      h('b', { style: { width: '140px' } }, n.name), h('code', { style: { flex: 1 } }, n.expr), h('span', { style: { color: 'var(--ink-3)', width: '90px' } }, String(formatValue(C.wb.hf.getNamedExpressionValue(n.name)?.constructor === Array ? '' : C.wb.hf.getNamedExpressionValue(n.name)).text || '…')),
      h('button', { class: 'btn ghost', onclick: () => { removeName(n.name); render(); C.refresh(); } }, 'Eliminar'))));
    box.append(h('p', {}, h('button', { class: 'btn', onclick: async () => { await defineName(); render(); } }, 'Nuevo…')));
  };
  render();
  await modal({ title: 'Administrador de nombres', body: box });
}

// ================================================================ notas en celdas
export function setNote(r, c, text) {
  const s = C.wb.sheet;
  C.act('Nota', () => { const m = new Map(s.notes || []); if (text) m.set(`${r},${c}`, text); else m.delete(`${r},${c}`); s.notes = m; }, { hf: false });
}
export async function noteDialog() {
  const { r, c } = C.grid.sel; const s = C.wb.sheet;
  const v = await form(`Nota en ${addr(r, c)}`, [{ name: 't', label: 'Texto de la nota (vacío para eliminarla)', type: 'textarea', rows: 4, value: s.notes?.get(`${r},${c}`) || '' }], 'Guardar');
  if (v) setNote(r, c, v.t.trim());
}

// ================================================================ dar formato como tabla
export const TABLE_STYLES = { verde: ['#0E9C74', '#E3F5EF'], azul: ['#1F6FB2', '#E4EEF8'], morado: ['#5B48E8', '#ECE9FD'], naranja: ['#D9692F', '#FCEDE3'], gris: ['#4B5066', '#EEEFF3'], rosa: ['#DC3A63', '#FCE6EC'] };
export function formatAsTable(g, styleName = 'verde') {
  const [dark, light] = TABLE_STYLES[styleName] || TABLE_STYLES.verde;
  const s = C.wb.sheet; const wb = C.wb;
  C.act('Formato de tabla', () => {
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      const base = { ...(s.style(r, c) || {}) };
      const border = { b: r === g.r2 ? dark : '#D9DCE5' };
      if (r === g.r1) Object.assign(base, { b: true, color: '#FFFFFF', bg: dark, border });
      else Object.assign(base, { bg: (r - g.r1) % 2 === 0 ? light : undefined, border });
      wb.setStyle(r, c, base);
    }
    s.filter = { range: { ...g }, hidden: new Map() };
  }, { hf: false });
}
export function tableStyleMenu() {
  return Object.entries(TABLE_STYLES).map(([k, [d, l]]) => ({ label: `Estilo ${k}`, style: { borderLeft: `14px solid ${d}`, background: `linear-gradient(90deg, ${l}, transparent)` }, run: () => { let g = C.sel(); if (g.r1 === g.r2 && g.c1 === g.c2) g = C.wb.region(g.r1, g.c1); formatAsTable(C.clamp(g), k); } }));
}

// ================================================================ proteger
export function protectSheet(on, password = '') { const s = C.wb.sheet; C.act(on ? 'Proteger hoja' : 'Desproteger hoja', () => { s.protected = on ? { password } : null; }, { hf: false }); C.shell.ribbon.setOn('protect', on); }
export function canEdit(r, c, s = C.wb.sheet) { if (!s.protected) return true; return !!s.style(r, c)?.unlocked; }
export async function protectDialog() {
  const s = C.wb.sheet;
  if (s.protected) {
    if (s.protected.password) { const v = await form('Desproteger hoja', [{ name: 'p', label: 'Contraseña', type: 'password' }], 'Desproteger'); if (!v) return; if (v.p !== s.protected.password) { toast('Contraseña incorrecta', { error: true }); return; } }
    protectSheet(false); toast('La hoja ya se puede editar'); return;
  }
  const v = await form('Proteger hoja', [{ name: 'p', label: 'Contraseña (opcional)', type: 'password' }, { type: 'note', label: 'Las celdas marcadas como “desbloqueadas” (Formato → Desbloquear celdas) se podrán seguir editando.' }], 'Proteger');
  if (!v) return;
  protectSheet(true, v.p); toast('Hoja protegida');
}
export function toggleUnlock() { const g = C.clamp(C.sel()); const on = !C.wb.sheet.style(g.r1, g.c1)?.unlocked; C.act('Desbloquear', () => C.wb.styleRange(g, { unlocked: on || undefined }), { hf: false }); toast(on ? 'Celdas desbloqueadas (editables con la hoja protegida)' : 'Celdas bloqueadas'); }
export function protectBook() { C.wb.protectedStructure = !C.wb.protectedStructure; C.shell.ribbon.setOn('protectBook', C.wb.protectedStructure); toast(C.wb.protectedStructure ? 'Estructura del libro protegida: no se pueden agregar, quitar ni renombrar hojas' : 'Estructura del libro desprotegida'); C.shell.markDirty(); }

// ================================================================ tabla dinámica
const AGG = { sum: 'Suma', count: 'Cuenta', avg: 'Promedio', max: 'Máximo', min: 'Mínimo' };
export function buildPivot(def, targetSheet = null) {
  const wb = C.wb;
  const [sheetName, ref] = def.source.includes('!') ? def.source.split('!') : [wb.sheet.name, def.source];
  const src = wb.sheets.find((x) => x.name === sheetName.replace(/^'|'$/g, '')) || wb.sheet;
  const g = parseRange(ref); if (!g) throw new Error('Rango de origen no válido');
  const headers = []; for (let c = g.c1; c <= g.c2; c++) headers.push(String(wb.value(g.r1, c, src) ?? `Columna ${c - g.c1 + 1}`));
  const ci = (n) => headers.indexOf(n);
  if (ci(def.rows) < 0 || ci(def.value) < 0) throw new Error('Elige los campos de filas y valores');
  const rowsKeys = [], colKeys = [];
  const cells = new Map(); const rowTot = new Map(); const colTot = new Map(); let grand = [];
  const push = (m, k, v) => { if (!m.has(k)) m.set(k, []); m.get(k).push(v); };
  for (let r = g.r1 + 1; r <= g.r2; r++) {
    const rk = String(wb.display(g.c1 + ci(def.rows) >= 0 ? r : r, g.c1 + ci(def.rows), src).text || '(vacío)');
    const ck = def.cols && ci(def.cols) >= 0 ? String(wb.display(r, g.c1 + ci(def.cols), src).text || '(vacío)') : '';
    const v = wb.value(r, g.c1 + ci(def.value), src);
    if (rk === '(vacío)' && v == null) continue;
    if (!rowsKeys.includes(rk)) rowsKeys.push(rk);
    if (!colKeys.includes(ck)) colKeys.push(ck);
    const val = typeof v === 'number' ? v : (def.agg === 'count' && v != null && v !== '' ? 1 : null);
    if (val == null) continue;
    push(cells, `${rk}\u0000${ck}`, val); push(rowTot, rk, val); push(colTot, ck, val); grand.push(val);
  }
  const sortKeys = (a, b) => a.localeCompare(b, 'es', { numeric: true });
  rowsKeys.sort(sortKeys); colKeys.sort(sortKeys);
  const agg = (arr) => { if (!arr || !arr.length) return null; if (def.agg === 'count') return arr.length; if (def.agg === 'avg') return arr.reduce((a, b) => a + b, 0) / arr.length; if (def.agg === 'max') return Math.max(...arr); if (def.agg === 'min') return Math.min(...arr); return arr.reduce((a, b) => a + b, 0); };
  const data = [[`${AGG[def.agg] || 'Suma'} de ${def.value}`]];
  const hasCols = def.cols && colKeys.length && !(colKeys.length === 1 && colKeys[0] === '');
  data.push([def.rows, ...(hasCols ? colKeys : []), 'Total general']);
  for (const rk of rowsKeys) data.push([rk, ...(hasCols ? colKeys.map((ck) => agg(cells.get(`${rk}\u0000${ck}`))) : []), agg(rowTot.get(rk))]);
  data.push(['Total general', ...(hasCols ? colKeys.map((ck) => agg(colTot.get(ck))) : []), agg(grand)]);
  let s = targetSheet;
  if (!s) { let n = 1; while (wb.sheets.some((x) => x.name === `Dinámica${n}`)) n++; s = wb.addSheet(`Dinámica${n}`); }
  wb.hf.setSheetContent(s.id, data);
  s.styles = new Map(); s.pivot = def;
  const W = data[1].length, last = data.length - 1;
  wb.setStyle(0, 0, { b: true, size: 13, color: '#0A6B50' }, s);
  for (let c = 0; c < W; c++) { wb.setStyle(1, c, { b: true, bg: '#DFF5EE', border: { b: '#0E9C74' } }, s); wb.setStyle(last, c, { b: true, border: { t: '#0E9C74' }, bg: '#F0F8F5', ...(c ? { fmt: '#,##0.##' } : {}) }, s); }
  for (let r = 2; r < last; r++) { for (let c = 1; c < W; c++) wb.setStyle(r, c, { fmt: '#,##0.##', ...(c === W - 1 ? { b: true } : {}) }, s); }
  s.colW.set(0, 150);
  wb.undoStack = []; wb.redoStack = []; wb.hf.clearUndoStack?.();
  return s;
}
export async function pivotDialog() {
  let g = C.sel(); if (g.r1 === g.r2 && g.c1 === g.c2) g = C.wb.region(g.r1, g.c1);
  g = C.clamp(g);
  if (g.r2 <= g.r1) { toast('Selecciona una tabla con encabezados (al menos 2 filas).', { error: true }); return; }
  const headers = []; for (let c = g.c1; c <= g.c2; c++) headers.push(String(C.wb.value(g.r1, c) ?? `Columna ${c - g.c1 + 1}`));
  const numeric = headers.filter((_, i) => typeof C.wb.value(g.r1 + 1, g.c1 + i) === 'number');
  const v = await form('Crear tabla dinámica', [
    { name: 'source', label: 'Datos de origen', value: `${C.wb.sheet.name}!${rangeStr(g)}` },
    { name: 'rows', label: 'Filas', type: 'select', value: headers[0], options: headers },
    { name: 'cols', label: 'Columnas (opcional)', type: 'select', value: '', options: [['', '(ninguna)'], ...headers.map((x) => [x, x])] },
    { name: 'value', label: 'Valores', type: 'select', value: numeric[numeric.length - 1] || headers[headers.length - 1], options: headers },
    { name: 'agg', label: 'Resumir con', type: 'select', value: 'sum', options: Object.entries(AGG) },
    { type: 'note', label: 'La tabla se crea en una hoja nueva. Usa “Actualizar” si cambian los datos de origen.' },
  ], 'Crear');
  if (!v) return;
  try { const s = buildPivot(v); C.switchSheet(C.wb.sheets.indexOf(s)); C.shell.markDirty(); toast('Tabla dinámica creada'); } catch (err) { toast(err.message, { error: true }); }
}
export function refreshPivot() {
  const s = C.wb.sheet; if (!s.pivot) { toast('Esta hoja no es una tabla dinámica.'); return; }
  buildPivot(s.pivot, s); C.switchSheet(C.wb.active); C.shell.markDirty(); toast('Tabla dinámica actualizada');
}

// ================================================================ configuración de página
export async function pageSetupDialog() {
  const s = C.wb.sheet; const ps = s.pageSetup || {};
  const v = await form('Configurar página', [
    { name: 'orientation', label: 'Orientación', type: 'select', value: ps.orientation || 'auto', options: [['auto', 'Automática'], ['portrait', 'Vertical'], ['landscape', 'Horizontal']] },
    { name: 'printArea', label: 'Área de impresión (vacío = área usada)', value: s.printArea || '' },
    { name: 'fit', label: 'Ajustar', type: 'select', value: ps.fit || 'none', options: [['none', 'Tamaño real'], ['width', 'Todas las columnas en una página'], ['page', 'Toda la hoja en una página']] },
    { name: 'titleRows', label: 'Repetir filas de título (p. ej. 1:1)', value: ps.titleRows || '' },
    { name: 'gridlines', label: 'Imprimir líneas de cuadrícula', type: 'checkbox', value: ps.gridlines !== false },
  ], 'Aplicar');
  if (!v) return;
  if (v.printArea && !parseRange(v.printArea)) { toast('Área no válida', { error: true }); return; }
  C.act('Configurar página', () => { s.printArea = v.printArea ? v.printArea.toUpperCase() : null; s.pageSetup = { orientation: v.orientation, fit: v.fit, titleRows: v.titleRows || '', gridlines: v.gridlines }; }, { hf: false });
}
export function setPrintArea(clear = false) { const s = C.wb.sheet; C.act('Área de impresión', () => { s.printArea = clear ? null : rangeStr(C.clamp(C.sel())); }, { hf: false }); toast(clear ? 'Área de impresión borrada' : `Área de impresión: ${s.printArea}`); }
export { AGG, esc, closeDropdown, dropdown, icon, confirmBox };
