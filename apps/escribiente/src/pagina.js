// Escribiente — Página como en LibreOffice Writer: encabezado y pie con primera página distinta y
// páginas pares/impares distintas, en tres partes (izquierda, centro y derecha, separadas con
// tabulación), con campos ({pagina}, {paginas}, {fecha}, {titulo}, {autor}, {archivo}) y formato;
// borde de página y numeración de líneas. GPL-3.0
import { h, t } from '../../shared/office.js';
import { tabbed, field, numIn, sel, chk } from '../../shared/dialogs.js';
import { fmtDate } from './insertar.js';

let C = null; // { state, editor, pagesEl, shell, schedulePaginate }
export function init(ctx) { C = ctx; }
const PX_CM = 2.54 / 96;
const r2 = (n) => Math.round(n * 100) / 100;

// ================================================================ encabezado y pie
// texto que toca a la página i (0 = primera)
export function hfFor(kind, i) {
  const s = C.state;
  if (s.hfFirst && i === 0) return s[`${kind}First`] ?? '';
  if (s.hfOddEven && (i + 1) % 2 === 0) return s[`${kind}Even`] ?? s[kind] ?? '';
  return s[kind] || '';
}
export function hfTokens(tpl, i, n) {
  const props = C.shell?.docProps?.() || {};
  return String(tpl || '').replace(/\{pagina\}/gi, String(i + 1 + ((C.state.firstPageNum || 1) - 1))).replace(/\{paginas\}/gi, String(n))
    .replace(/\{fecha\}/gi, fmtDate(new Date(), 'dd/MM/yyyy')).replace(/\{titulo\}/gi, props.title || C.shell?.doc?.name?.replace(/\.[^.]+$/, '') || '')
    .replace(/\{autor\}/gi, props.author || '').replace(/\{archivo\}/gi, C.shell?.doc?.name || '');
}
// una línea "izquierda<TAB>centro<TAB>derecha" se muestra en tres partes, como las tabulaciones de Writer
export function hfNode(kind, i, n, style) {
  const txt = hfFor(kind, i); if (!txt) return null;
  const st = C.state.hfStyle || {};
  const box = h('div', { class: `hf ${kind === 'header' ? 'top' : 'bottom'}${st.line ? ' hf-line' : ''}`, style: { ...style, fontSize: st.size ? `${st.size}pt` : '', fontWeight: st.bold ? '700' : '', fontStyle: st.italic ? 'italic' : '', color: st.color || '' } });
  for (const line of String(txt).split('\n')) {
    const parts = hfTokens(line, i, n).split('\t');
    if (parts.length === 1) box.append(h('div', { class: 'hf-l', style: { textAlign: st.align || 'left' } }, parts[0]));
    else box.append(h('div', { class: 'hf-l hf-3' }, h('span', {}, parts[0]), h('span', {}, parts[1] || ''), h('span', {}, parts.slice(2).join(' '))));
  }
  return box;
}
export async function headerFooterDialog(start = 0) {
  const s = C.state; const st = s.hfStyle || {};
  const ta = (name, value, ph) => h('textarea', { name, rows: 2, placeholder: t(ph), style: { width: '100%', resize: 'vertical', fontFamily: 'inherit' } }, value || '');
  const hfTab = (kind) => h('div', { class: 'fx-col' },
    field(kind === 'header' ? 'Encabezado' : 'Pie de página', ta(kind, s[kind], 'Izquierda[Tab]centro[Tab]derecha')),
    field('Primera página', ta(`${kind}First`, s[`${kind}First`], 'Vacío: sin encabezado en la primera página')),
    field('Páginas pares', ta(`${kind}Even`, s[`${kind}Even`] ?? s[kind], '')));
  const body = await tabbed('Encabezado y pie de página', [
    { label: 'Encabezado', body: hfTab('header') },
    { label: 'Pie de página', body: hfTab('footer') },
    { label: 'Opciones', body: h('div', { class: 'fx-grid' },
      chk('hfFirst', 'Primera página distinta', s.hfFirst), chk('hfOddEven', 'Páginas pares e impares distintas', s.hfOddEven),
      field('Números de página', sel('pageNumbers', s.pageNumbers || 'none', [['none', 'Sin números'], ['footer-center', 'Abajo, al centro'], ['footer-right', 'Abajo, a la derecha'], ['header-right', 'Arriba, a la derecha']])),
      field('Primer número de página', numIn('firstPageNum', s.firstPageNum || 1, { step: 1, min: 0, max: 9999 })),
      h('h4', {}, t('Formato')), field('Tamaño', numIn('size', st.size || 10, { step: 0.5, min: 6, max: 36, unit: 'pt' })), field('Alineación', sel('align', st.align || 'left', [['left', 'Izquierda'], ['center', 'Centro'], ['right', 'Derecha']])),
      chk('bold', 'Negrita', st.bold), chk('italic', 'Cursiva', st.italic), field('Color', h('input', { type: 'color', name: 'color', value: st.color || '#555555' })), chk('line', 'Línea separadora', st.line),
      h('p', { class: 'fx-note', style: { gridColumn: '1 / -1' } }, t('Campos: {pagina}, {paginas}, {fecha}, {titulo}, {autor}, {archivo}. Separa con tabulación para izquierda, centro y derecha (en el cuadro, escribe \\t o pega una tabulación).'))) },
  ], { start, width: '600px' });
  if (!body) return;
  const v = (n) => body.querySelector(`[name="${n}"]`);
  const txt = (n) => v(n).value.replace(/\\t/g, '\t');
  Object.assign(s, {
    header: txt('header'), footer: txt('footer'), headerFirst: txt('headerFirst'), footerFirst: txt('footerFirst'), headerEven: txt('headerEven'), footerEven: txt('footerEven'),
    hfFirst: v('hfFirst').checked, hfOddEven: v('hfOddEven').checked, pageNumbers: v('pageNumbers').value, firstPageNum: Number(v('firstPageNum').value) || 1,
    hfStyle: { size: Number(v('size').value) !== 10 ? Number(v('size').value) : undefined, align: v('align').value !== 'left' ? v('align').value : undefined, bold: v('bold').checked || undefined, italic: v('italic').checked || undefined, color: v('color').value !== '#555555' ? v('color').value : undefined, line: v('line').checked || undefined },
  });
  C.shell.markDirty(); C.schedulePaginate();
}

// ================================================================ borde de página
export async function pageBorderDialog() {
  const b = C.state.pageBorder || {};
  const body = await tabbed('Borde de página', [{ label: 'Borde', body: h('div', { class: 'fx-grid' },
    field('Estilo', sel('style', b.style || 'none', [['none', 'Sin borde'], ['solid', 'Sólido'], ['double', 'Doble'], ['dashed', 'Discontinuo'], ['dotted', 'Punteado']])),
    field('Ancho', numIn('width', b.width || 1, { step: 0.25, min: 0.25, max: 12, unit: 'pt' })), field('Color', h('input', { type: 'color', name: 'color', value: b.color || '#000000' })),
    field('Distancia al texto', numIn('space', r2((b.space ?? 18) * PX_CM), { step: 0.1, min: 0, max: 3, unit: 'cm' })), chk('shadow', 'Sombra', b.shadow)) }], { width: '480px' });
  if (!body) return;
  const v = (n) => body.querySelector(`[name="${n}"]`);
  C.state.pageBorder = v('style').value === 'none' ? null : { style: v('style').value, width: Number(v('width').value) || 1, color: v('color').value, space: Math.round(Number(v('space').value) / PX_CM), shadow: v('shadow').checked || undefined };
  C.shell.markDirty(); C.schedulePaginate();
}
export function borderNode() {
  const b = C.state.pageBorder; if (!b) return null; const p = C.state.page; const sp = b.space ?? 18;
  return h('div', { class: 'pborder', style: { left: `${Math.max(4, p.ml - sp)}px`, right: `${Math.max(4, p.mr - sp)}px`, top: `${Math.max(4, p.mt - sp)}px`, bottom: `${Math.max(4, p.mb - sp)}px`, border: `${b.width * (b.style === 'double' ? 3 : 1) / 0.75}px ${b.style} ${b.color}`, boxShadow: b.shadow ? '4px 4px 0 rgba(0,0,0,.35)' : '' } });
}

// ================================================================ numeración de líneas
export async function lineNumDialog() {
  const l = C.state.lineNum || {};
  const body = await tabbed('Numeración de líneas', [{ label: 'Numeración', body: h('div', { class: 'fx-grid' },
    chk('on', 'Mostrar números de línea', l.on), h('span'),
    field('Cada', numIn('every', l.every || 1, { step: 1, min: 1, max: 100, unit: t('líneas') })), field('Distancia al texto', numIn('dist', r2((l.dist ?? 18) * PX_CM), { step: 0.1, min: 0.1, max: 3, unit: 'cm' })),
    field('Reiniciar', sel('restart', l.restart || 'doc', [['doc', 'Continuo en todo el documento'], ['page', 'En cada página']])), chk('blank', 'Contar líneas vacías', l.blank !== false)) }], { width: '480px' });
  if (!body) return;
  const v = (n) => body.querySelector(`[name="${n}"]`);
  C.state.lineNum = v('on').checked ? { on: true, every: Math.max(1, Number(v('every').value) || 1), dist: Math.round(Number(v('dist').value) / PX_CM), restart: v('restart').value, blank: v('blank').checked } : null;
  C.shell.markDirty(); C.schedulePaginate();
}
// dibuja los números en el margen izquierdo de cada hoja (después de paginar)
export function drawLineNumbers() {
  const pe = C.pagesEl; pe.querySelectorAll('.lnums').forEach((x) => x.remove());
  const L = C.state.lineNum; if (!L?.on) return;
  const z = (C.state.zoom || 100) / 100; const base = pe.getBoundingClientRect();
  const S = C.state.page.h + (C.state.page.gap || 0);
  const layer = h('div', { class: 'lnums' });
  let count = 0; let lastPage = -1;
  const blocks = [...C.editor.querySelectorAll(':scope > p, :scope > h1, :scope > h2, :scope > h3, :scope > h4, :scope > h5, :scope > h6, :scope > blockquote, :scope > ul > li, :scope > ol > li, :scope > div.cols > p, :scope > div.cols > h1, :scope > div.cols > h2')];
  for (const b of blocks) {
    const r = document.createRange(); r.selectNodeContents(b);
    let rects = [...r.getClientRects()].filter((x) => x.height > 2);
    if (!rects.length) { if (L.blank === false) continue; rects = [b.getBoundingClientRect()]; }
    const tops = []; for (const rc of rects) { const y = Math.round(rc.top); if (!tops.some((t2) => Math.abs(t2.y - y) < 3)) tops.push({ y, h: rc.height, x: rc.left }); }
    for (const ln of tops.sort((a, b2) => a.y - b2.y)) {
      const top = (ln.y - base.top) / z; const page = Math.floor(top / S);
      if (L.restart === 'page' && page !== lastPage) count = 0; lastPage = page;
      count += 1;
      if (count % L.every) continue;
      layer.append(h('div', { class: 'lnum', style: { top: `${top}px`, height: `${ln.h / z}px`, lineHeight: `${ln.h / z}px`, left: `${Math.max(2, C.state.page.ml - (L.dist ?? 18) - 40)}px` } }, String(count)));
    }
  }
  pe.append(layer);
}
