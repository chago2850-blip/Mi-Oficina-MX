// Escribiente — Insertar como en LibreOffice Writer: campos (fecha, hora, página, total de páginas,
// título, autor, archivo, palabras), marcadores, referencias cruzadas, títulos de objetos (Figura 1,
// Tabla 1…), índice de ilustraciones y de tablas, índice alfabético con entradas marcadas y salto de
// columna. Los campos se actualizan solos al paginar y con F9. GPL-3.0
import { h, toast, form, t, tf, esc } from '../../shared/office.js';

let C = null; // { editor, state, shell, insertHTMLAtCursor, restoreSelection, saveSelection, currentBlock, schedulePaginate }
export function init(ctx) { C = ctx; }
// inserta HTML en línea donde está el cursor (insertHTML de Chromium saca los elementos no editables del párrafo)
export function insertInline(html) {
  C.restoreSelection(); C.editor.focus({ preventScroll: true });
  const s = getSelection(); if (!s.rangeCount) { C.insertHTMLAtCursor(html); return; }
  let r = s.getRangeAt(0);
  if (!C.editor.contains(r.startContainer)) { C.insertHTMLAtCursor(html); return; }
  if (r.startContainer === C.editor) { const blk = C.editor.children[Math.max(0, r.startOffset - 1)] || C.editor.firstElementChild; r = document.createRange(); r.selectNodeContents(blk); r.collapse(false); }
  r.deleteContents();
  const tpl = document.createElement('template'); tpl.innerHTML = html; const nodes = [...tpl.content.childNodes];
  r.insertNode(tpl.content);
  const last = nodes[nodes.length - 1]; const after = document.createRange(); after.setStartAfter(last); after.collapse(true); s.removeAllRanges(); s.addRange(after);
  C.saveSelection(); C.shell.markDirty(); C.schedulePaginate();
}
const uid = () => Math.random().toString(36).slice(2, 8);

// ================================================================ campos
export const FIELDS = [
  ['date', 'Fecha'], ['time', 'Hora'], ['page', 'Número de página'], ['pages', 'Total de páginas'],
  ['title', 'Título del documento'], ['author', 'Autor'], ['file', 'Nombre de archivo'], ['words', 'Recuento de palabras'],
];
export const DATE_FMTS = ['dd/MM/yyyy', 'd/M/yy', "d 'de' MMMM 'de' yyyy", "EEEE, d 'de' MMMM 'de' yyyy", 'yyyy-MM-dd', 'MMMM yyyy'];
export const TIME_FMTS = ['HH:mm', 'HH:mm:ss', 'h:mm a'];
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const DIAS = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
// formato estilo Word/Java: d dd M MM MMM MMMM yy yyyy EEEE H HH h mm ss a, 'literal'
export function fmtDate(d, f) {
  const en = document.documentElement.lang === 'en';
  const M = en ? MONTHS : MESES, D = en ? DAYS : DIAS;
  const p2 = (n) => String(n).padStart(2, '0');
  return f.replace(/'([^']*)'|yyyy|yy|MMMM|MMM|MM|M|dd|d|EEEE|EEE|HH|H|hh|h|mm|ss|a/g, (m, lit) => {
    if (lit !== undefined) return lit;
    switch (m) {
      case 'yyyy': return String(d.getFullYear()); case 'yy': return p2(d.getFullYear() % 100);
      case 'MMMM': return M[d.getMonth()]; case 'MMM': return M[d.getMonth()].slice(0, 3); case 'MM': return p2(d.getMonth() + 1); case 'M': return String(d.getMonth() + 1);
      case 'dd': return p2(d.getDate()); case 'd': return String(d.getDate()); case 'EEEE': return D[d.getDay()]; case 'EEE': return D[d.getDay()].slice(0, 3);
      case 'HH': return p2(d.getHours()); case 'H': return String(d.getHours()); case 'hh': return p2(d.getHours() % 12 || 12); case 'h': return String(d.getHours() % 12 || 12);
      case 'mm': return p2(d.getMinutes()); case 'ss': return p2(d.getSeconds()); case 'a': return d.getHours() < 12 ? 'a. m.' : 'p. m.';
      default: return m;
    }
  });
}
const fieldHtml = (kind, attrs = {}) => `<span class="fld" data-fld="${kind}"${Object.entries(attrs).map(([k, v]) => (v == null ? '' : ` data-${k}="${esc(String(v))}"`)).join('')} contenteditable="false">${esc(fieldText({ dataset: { fld: kind, ...attrs } }) || '…')}</span>`;
export function insertField(kind, attrs = {}) {
  if (kind === 'date' && !attrs.fmt) attrs.fmt = DATE_FMTS[0];
  if (kind === 'time' && !attrs.fmt) attrs.fmt = TIME_FMTS[0];
  insertInline(fieldHtml(kind, attrs));
  setTimeout(updateFields, 50);
}
export async function fieldDialog() {
  const v = await form('Campos', [
    { name: 'kind', label: 'Tipo', type: 'select', value: 'date', options: FIELDS.map(([k, l]) => [k, t(l)]) },
    { name: 'dfmt', label: 'Formato de fecha', type: 'select', value: DATE_FMTS[0], options: DATE_FMTS.map((f) => [f, fmtDate(new Date(), f)]) },
    { name: 'tfmt', label: 'Formato de hora', type: 'select', value: TIME_FMTS[0], options: TIME_FMTS.map((f) => [f, fmtDate(new Date(), f)]) },
    { name: 'fixed', label: 'Fijo (no se actualiza)', type: 'checkbox', value: false },
  ], 'Insertar');
  if (!v) return;
  insertField(v.kind, { fmt: v.kind === 'date' ? v.dfmt : v.kind === 'time' ? v.tfmt : undefined, fixed: v.fixed && (v.kind === 'date' || v.kind === 'time') ? '1' : undefined });
}
// página (base 1) en la que está un nodo, según la última paginación
export function pageOfNode(n) {
  const ed = C.editor; let b = n; while (b && b.parentElement !== ed) b = b.parentElement;
  const i = b ? [...ed.children].indexOf(b) : -1;
  return (C.state.pageOf?.[i] ?? 0) + 1;
}
function fieldText(el) {
  const d = el.dataset; const props = C?.shell?.docProps?.() || {};
  switch (d.fld) {
    case 'date': case 'time': return d.fixed && el.textContent && el.textContent !== '…' ? el.textContent : fmtDate(d.value ? new Date(d.value) : new Date(), d.fmt || (d.fld === 'date' ? DATE_FMTS[0] : TIME_FMTS[0]));
    case 'page': return el.isConnected ? String(pageOfNode(el)) : '1';
    case 'pages': return String(C?.state?.pages || 1);
    case 'title': return props.title || C?.shell?.doc?.name?.replace(/\.[^.]+$/, '') || t('Sin título');
    case 'author': return props.author || '';
    case 'file': return C?.shell?.doc?.name || t('Sin título');
    case 'words': return String(((C?.editor?.innerText || '').match(/[\p{L}\p{N}]+(?:['’-][\p{L}\p{N}]+)*/gu) || []).length);
    case 'seq': return el.textContent || '1';
    case 'ref': return refText(el) ?? el.textContent;
    default: return el.textContent;
  }
}

// ================================================================ títulos (Figura 1, Tabla 2…) y numeración
export const CATEGORIES = ['Figura', 'Tabla', 'Ilustración', 'Dibujo', 'Texto'];
export function renumberCaptions() {
  const count = new Map();
  for (const s of C.editor.querySelectorAll('.fld[data-fld="seq"]')) { const n = (count.get(s.dataset.seq) || 0) + 1; count.set(s.dataset.seq, n); if (s.textContent !== String(n)) s.textContent = String(n); }
}
export async function captionDialog() {
  C.restoreSelection();
  const b = C.currentBlock();
  const near = b && (b.tagName === 'TABLE' ? b : b.closest?.('table') || (b.querySelector?.('img') ? b : null));
  const isTable = near && (near.tagName === 'TABLE' || near.closest?.('table'));
  const v = await form('Insertar título', [
    { name: 'cat', label: 'Categoría', type: 'select', value: isTable ? 'Tabla' : 'Figura', options: CATEGORIES.map((c) => [c, t(c)]) },
    { name: 'text', label: 'Título', value: '' },
    { name: 'sep', label: 'Separador', value: ': ' },
    { name: 'pos', label: 'Posición', type: 'select', value: isTable ? 'above' : 'below', options: [['above', t('Encima')], ['below', t('Debajo')]] },
  ], 'Insertar');
  if (!v) return;
  insertCaption(v.cat, v.text, { sep: v.sep, pos: v.pos, near: near && (near.tagName === 'TABLE' ? near : near.closest?.('table') || near) });
}
export function insertCaption(cat, text = '', { sep = ': ', pos = 'below', near = null } = {}) {
  const id = `_Ref${uid()}`;
  const p = document.createElement('p'); p.className = 'caption'; p.dataset.seq = cat;
  p.innerHTML = `<a class="bm" id="${id}"></a>${esc(t(cat))} <span class="fld" data-fld="seq" data-seq="${esc(cat)}" contenteditable="false">1</span>${esc(sep)}${esc(text) || '<br>'}`;
  let anchor = near; while (anchor && anchor.parentElement !== C.editor) anchor = anchor.parentElement;
  if (anchor) { if (pos === 'above') anchor.before(p); else anchor.after(p); } else { const blk = C.currentBlock(); let top = blk; while (top && top.parentElement !== C.editor) top = top.parentElement; if (top) top.after(p); else C.editor.append(p); }
  renumberCaptions();
  const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); C.saveSelection();
  C.shell.markDirty(); C.schedulePaginate();
  return p;
}

// ================================================================ marcadores
export const bookmarks = () => [...C.editor.querySelectorAll('a.bm[id]')].filter((a) => !a.id.startsWith('_'));
export async function bookmarkDialog() {
  const list = bookmarks();
  const v = await form('Marcador', [
    { name: 'name', label: 'Nombre', value: `Marcador${list.length + 1}` },
    ...(list.length ? [{ name: 'existing', label: 'Marcadores del documento', type: 'select', value: '', options: [['', t('(ninguno)')], ...list.map((a) => [a.id, a.id])] }, { name: 'act', label: 'Con el seleccionado', type: 'select', value: 'insert', options: [['insert', t('Insertar el nuevo')], ['goto', t('Ir a él')], ['delete', t('Eliminarlo')]] }] : []),
  ], 'Aceptar');
  if (!v) return;
  if (v.act === 'goto' && v.existing) return gotoBookmark(v.existing);
  if (v.act === 'delete' && v.existing) { C.editor.querySelector(`a.bm[id="${CSS.escape(v.existing)}"]`)?.remove(); C.shell.markDirty(); toast(tf('Marcador {0} eliminado', v.existing)); return; }
  const name = v.name.trim().replace(/\s+/g, '_');
  if (!/^[\p{L}_][\p{L}\p{N}_.-]*$/u.test(name)) { toast('El nombre del marcador no es válido.', { error: true }); return; }
  if (C.editor.querySelector(`a.bm[id="${CSS.escape(name)}"]`)) { toast(tf('Ya existe un marcador llamado {0}.', name), { error: true }); return; }
  insertBookmark(name);
}
export function insertBookmark(name) { insertInline(`<a class="bm" id="${esc(name)}"></a>`); toast(tf('Marcador {0} insertado', name)); }
export function gotoBookmark(name) {
  const a = C.editor.querySelector(`a.bm[id="${CSS.escape(name)}"]`); if (!a) return false;
  const r = document.createRange(); r.setStartAfter(a); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); C.saveSelection();
  (a.parentElement || a).scrollIntoView({ block: 'center' }); return true;
}

// ================================================================ referencias cruzadas
function targets() {
  const out = [];
  for (const hd of C.editor.querySelectorAll(':scope > h1, :scope > h2, :scope > h3, :scope > h4')) out.push({ kind: 'h', label: hd.textContent.trim().slice(0, 60), el: hd });
  for (const a of bookmarks()) out.push({ kind: 'bm', label: a.id, id: a.id, el: a });
  for (const c of C.editor.querySelectorAll('p.caption')) { const bm = c.querySelector('a.bm[id]'); out.push({ kind: 'cap', label: c.textContent.trim().slice(0, 60), id: bm?.id, el: c, cat: c.dataset.seq }); }
  return out;
}
export async function crossRefDialog() {
  const list = targets();
  if (!list.length) { toast('No hay títulos, marcadores ni títulos de objetos a los que hacer referencia.'); return; }
  const KIND = { h: t('Título de sección'), bm: t('Marcador'), cap: t('Título de objeto') };
  const v = await form('Referencia cruzada', [
    { name: 'target', label: 'Referencia a', type: 'select', value: '0', options: list.map((x, i) => [String(i), `${KIND[x.kind]}: ${x.label}`]) },
    { name: 'show', label: 'Insertar referencia a', type: 'select', value: 'text', options: [['text', t('Texto de referencia')], ['page', t('Página')], ['label', t('Categoría y número (títulos)')], ['number', t('Número (títulos)')], ['above', t('Arriba/abajo')]] },
  ], 'Insertar');
  if (!v) return;
  const x = list[Number(v.target)];
  let id = x.id;
  if (!id) { id = `_Ref${uid()}`; const a = document.createElement('a'); a.className = 'bm'; a.id = id; (x.kind === 'h' ? x.el : x.el).prepend(a); }
  insertInline(`<span class="fld" data-fld="ref" data-ref="${esc(id)}" data-show="${v.show}" contenteditable="false">…</span>`);
  setTimeout(updateFields, 50);
}
function refText(el) {
  const id = el.dataset.ref; const a = id && C.editor.querySelector(`a.bm[id="${CSS.escape(id)}"]`);
  if (!a) return t('¡Error: referencia no encontrada!');
  const para = a.closest('p,h1,h2,h3,h4,h5,h6,li,td,th') || a.parentElement;
  const cap = a.closest('p.caption');
  switch (el.dataset.show) {
    case 'page': return String(pageOfNode(a));
    case 'number': return cap ? cap.querySelector('.fld[data-fld="seq"]')?.textContent || '' : String([...C.editor.querySelectorAll(':scope > h1, :scope > h2, :scope > h3')].indexOf(para) + 1);
    case 'label': return cap ? `${t(cap.dataset.seq || cap.querySelector('.fld[data-fld="seq"]')?.dataset.seq || '')} ${cap.querySelector('.fld[data-fld="seq"]')?.textContent || ''}` : (para?.textContent || '').trim();
    case 'above': { const pos = el.compareDocumentPosition(a); return pos & Node.DOCUMENT_POSITION_PRECEDING ? t('arriba') : t('abajo'); }
    default: { const c = para.cloneNode(true); c.querySelectorAll('.fld[data-fld="ref"], sup, .ix').forEach((n) => n.remove()); return c.textContent.trim(); }
  }
}

// ================================================================ actualizar campos (al paginar y con F9)
let busy = false;
export function updateFields() {
  if (busy || !C) return; busy = true;
  try {
    renumberCaptions();
    for (const f of C.editor.querySelectorAll('.fld')) { if (f.dataset.fld === 'seq') continue; const v = fieldText(f); if (v != null && f.textContent !== v) f.textContent = v; }
  } finally { busy = false; }
}
export function updateAll() {
  updateFields();
  for (const tb of C.editor.querySelectorAll('.toc[data-kind]')) tb.outerHTML = buildIndex(tb.dataset.kind, tb.dataset.seq);
  C.schedulePaginate();
  setTimeout(() => { updateFields(); for (const tb of C.editor.querySelectorAll('.toc[data-kind]')) tb.outerHTML = buildIndex(tb.dataset.kind, tb.dataset.seq); }, 120);
  toast('Campos e índices actualizados');
}

// ================================================================ índices: de ilustraciones / tablas y alfabético
export async function figureIndexDialog() {
  const v = await form('Índice de ilustraciones', [{ name: 'seq', label: 'Categoría', type: 'select', value: 'Figura', options: CATEGORIES.map((c) => [c, t(c)]) }], 'Insertar');
  if (v) insertIndex('fig', v.seq);
}
export function insertIndex(kind, seq) {
  C.insertHTMLAtCursor(`${buildIndex(kind, seq)}<p><br></p>`);
  setTimeout(() => { for (const tb of C.editor.querySelectorAll(`.toc[data-kind="${kind}"]`)) tb.outerHTML = buildIndex(kind, tb.dataset.seq); }, 120);
}
const line = (txt, pg) => `<div class="toc-l l1"><span>${esc(txt)}</span><span></span><span>${esc(String(pg))}</span></div>`;
function buildIndex(kind, seq) {
  if (kind === 'fig') {
    const caps = [...C.editor.querySelectorAll(`p.caption[data-seq="${CSS.escape(seq || 'Figura')}"]`)];
    const lines = caps.map((c) => { const cl = c.cloneNode(true); cl.querySelectorAll('.ix').forEach((n) => n.remove()); return line(cl.textContent.trim(), pageOfNode(c)); }).join('') || `<div class="toc-l"><span>${esc(t('No hay títulos de esta categoría.'))}</span></div>`;
    const title = seq === 'Tabla' ? t('Índice de tablas') : t('Índice de ilustraciones');
    return `<div class="toc" data-kind="fig" data-seq="${esc(seq || 'Figura')}" contenteditable="false"><div class="toc-h">${esc(title)}</div>${lines}</div>`;
  }
  // alfabético: entradas marcadas, con sus páginas, agrupadas por letra
  const map = new Map();
  for (const e of C.editor.querySelectorAll('.ix[data-ix]')) { const k = e.dataset.ix.trim(); if (!k) continue; const pg = pageOfNode(e); if (!map.has(k)) map.set(k, new Set()); map.get(k).add(pg); }
  const coll = new Intl.Collator(document.documentElement.lang || 'es', { sensitivity: 'base' });
  const keys = [...map.keys()].sort(coll.compare);
  let lines = ''; let letter = '';
  for (const k of keys) { const L = k[0].toLocaleUpperCase(); if (L !== letter) { letter = L; lines += `<div class="toc-l l1 aidx-l"><span><b>${esc(L)}</b></span></div>`; } lines += line(k, [...map.get(k)].sort((a, b) => a - b).join(', ')).replace('l1', 'l2'); }
  if (!keys.length) lines = `<div class="toc-l"><span>${esc(t('Marca entradas con Referencias → Marcar entrada.'))}</span></div>`;
  return `<div class="toc" data-kind="alpha" contenteditable="false"><div class="toc-h">${esc(t('Índice alfabético'))}</div>${lines}</div>`;
}
export async function markEntryDialog() {
  C.restoreSelection();
  const sel = getSelection(); const txt = sel.toString().trim();
  const v = await form('Marcar entrada de índice', [{ name: 'term', label: 'Entrada', value: txt }], 'Marcar');
  if (!v || !v.term.trim()) return;
  C.restoreSelection(); const s2 = getSelection(); if (s2.rangeCount) s2.collapseToEnd();
  insertInline(`<span class="ix" data-ix="${esc(v.term.trim())}" contenteditable="false"></span>`);
  toast(tf('Entrada «{0}» marcada', v.term.trim()));
}

// ================================================================ salto de columna
export function insertColumnBreak() { insertInline('<span class="cbk" contenteditable="false"></span>'); }

void h;
