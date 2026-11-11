// Escribiente — estilos (como LibreOffice Writer): estilos de párrafo y de carácter con herencia,
// panel de estilos, nuevo/actualizar/modificar/eliminar, estilo siguiente; listas multinivel y
// numeración de capítulos. Los estilos viven en state.styles y se dibujan con una hoja CSS. GPL-3.0
import { h, modal, toast, form, dropdown, t, tf } from '../../shared/office.js';

let C = null; // { editor, state, shell, selectedBlocks, currentBlock, restoreSelection, saveSelection, schedulePaginate, setBlockStyle, applySpanStyles, FONTS, openSide, renderSide }
export function init(ctx) { C = ctx; sheet = document.createElement('style'); sheet.id = 'docStyles'; document.head.append(sheet); render(); }
let sheet = null;
const CM = 96 / 2.54;
const r2 = (n) => Math.round(n * 100) / 100;

// ------------------------------------------------------------------ modelo
export const BUILTIN = [
  { id: 'Normal', name: 'Normal', base: 'p' }, { id: 'Title', name: 'Título', base: 'title' }, { id: 'Subtitle', name: 'Subtítulo', base: 'subtitle' },
  ...[1, 2, 3, 4, 5, 6].map((n) => ({ id: `Heading${n}`, name: `Título ${n}`, base: `h${n}` })),
  { id: 'Quote', name: 'Cita', base: 'blockquote' },
];
export const BUILTIN_CHAR = [
  { id: 'Emphasis', name: 'Énfasis', css: { 'font-style': 'italic' } },
  { id: 'Strong', name: 'Énfasis fuerte', css: { 'font-weight': 'bold' } },
  { id: 'SourceText', name: 'Texto de código', css: { 'font-family': "'Liberation Mono', 'Courier New', monospace" } },
  { id: 'Quotation', name: 'Cita textual', css: { 'font-style': 'italic', color: '#555555' } },
];
const builtinOf = (id) => BUILTIN.find((b) => b.id === id);
const baseOfTag = (el) => { if (!el) return 'Normal'; if (el.dataset?.style && styles()[el.dataset.style]) return el.dataset.style; const tg = el.tagName; if (/^H[1-6]$/.test(tg)) return `Heading${tg[1]}`; if (tg === 'BLOCKQUOTE') return 'Quote'; if (el.classList?.contains('title')) return 'Title'; if (el.classList?.contains('subtitle')) return 'Subtitle'; return 'Normal'; };
export const styles = () => { C.state.styles ||= {}; return C.state.styles; };
// todos los estilos de párrafo (predeterminados + del documento)
export function allPara() {
  const st = styles();
  const out = BUILTIN.map((b) => ({ ...b, builtin: true, kind: 'p', css: st[b.id]?.css || {}, next: st[b.id]?.next || (b.base.startsWith('h') || b.base === 'title' || b.base === 'subtitle' ? 'Normal' : b.id) }));
  for (const s of Object.values(st)) if (s.kind === 'p' && !builtinOf(s.id)) out.push(s);
  return out;
}
export function allChar() {
  const st = styles();
  const out = BUILTIN_CHAR.map((b) => ({ ...b, builtin: true, kind: 'c', css: { ...b.css, ...(st[b.id]?.css || {}) } }));
  for (const s of Object.values(st)) if (s.kind === 'c' && !BUILTIN_CHAR.some((b) => b.id === s.id)) out.push(s);
  return out;
}
const find = (id) => allPara().find((s) => s.id === id) || allChar().find((s) => s.id === id);
// estilo predeterminado del que desciende (para saber la etiqueta HTML)
function rootBase(id) { let s = find(id); let g = 0; while (s && !s.builtin && g++ < 20) s = find(s.parent || 'Normal'); return s?.base || 'p'; }
// CSS efectivo: cadena de herencia
export function effectiveCss(id) {
  const chain = []; let s = find(id); let g = 0;
  while (s && g++ < 20) { chain.unshift(s); if (s.builtin) break; s = find(s.parent || 'Normal'); }
  return Object.assign({}, ...chain.map((x) => x.css || {}));
}
const cssText = (o) => Object.entries(o).filter(([, v]) => v != null && v !== '').map(([k, v]) => `${k}:${v}`).join(';');
const esc = (s) => String(s).replace(/["\\]/g, '\\$&');

// ------------------------------------------------------------------ hoja CSS de estilos, listas y capítulos
const SEL = { Normal: 'p:not([data-style]):not(.title):not(.subtitle)', Title: 'p.title:not([data-style])', Subtitle: 'p.subtitle:not([data-style])', Quote: 'blockquote:not([data-style])' };
export function render() {
  if (!sheet || !C) return;
  const rules = [];
  for (const s of allPara()) {
    const css = s.builtin ? s.css : effectiveCss(s.id);
    if (!Object.keys(css).length) continue;
    const sel = s.builtin ? (SEL[s.id] || `${s.base}:not([data-style])`) : `[data-style="${esc(s.id)}"]`;
    rules.push(`.doc.doc ${sel}{${cssText(css)}}`);
    if (s.id === 'Normal') rules.push(`.doc.doc li{${cssText(Object.fromEntries(Object.entries(css).filter(([k]) => /^(font|color|letter|text-transform|line-height)/.test(k))))}}`);
  }
  for (const s of allChar()) if (Object.keys(s.css || {}).length) rules.push(`.doc.doc span[data-cstyle="${esc(s.id)}"]{${cssText(s.css)}}`);
  rules.push(LIST_CSS_GET());
  rules.push(chapterCss(C.state.chapnum));
  sheet.textContent = rules.join('\n');
  C.onChange?.();
}

// ------------------------------------------------------------------ aplicar
const PARA_PROPS = ['margin-top', 'margin-bottom', 'margin-left', 'margin-right', 'text-indent', 'line-height', 'text-align', 'font-family', 'font-size', 'color', 'font-weight', 'font-style', 'background-color', 'border-top', 'border-bottom', 'border-left', 'border-right', 'padding-top', 'padding-bottom', 'padding-left', 'padding-right', 'letter-spacing', 'text-transform', 'font-variant'];
export function applyPara(id) {
  C.restoreSelection();
  const s = find(id); if (!s) return;
  const base = s.builtin ? s.base : rootBase(id);
  C.setBlockStyle(base, { silent: true });
  for (const b of C.selectedBlocks()) {
    if (!/^(P|H[1-6]|BLOCKQUOTE|LI)$/.test(b.tagName)) continue;
    // aplicar un estilo de párrafo quita el formato directo del párrafo (como en Writer); se conservan las tabulaciones
    const push = b.dataset.push; const mt0 = b.dataset.mt0;
    b.removeAttribute('style');
    if (push) { b.dataset.push = push; b.dataset.mt0 = mt0 || ''; }
    if (s.builtin) delete b.dataset.style; else b.dataset.style = id;
  }
  C.saveSelection(); C.shell.markDirty(); C.schedulePaginate(); C.reflectState?.(); C.renderSide?.();
}
export function applyChar(id) {
  C.restoreSelection();
  C.applySpanStyles({}, { attrs: { 'data-cstyle': id === '__none' ? null : id } });
  C.renderSide?.();
}
export function currentStyleId() { const b = C.currentBlock(); return baseOfTag(b?.closest?.('p,h1,h2,h3,h4,h5,h6,blockquote,li')); }

// estilo siguiente al pulsar Intro (p. ej. después de un título, Normal)
export function onNewParagraph() {
  const b = C.currentBlock(); const prev = b?.previousElementSibling;
  if (!b || !prev || b.textContent.trim()) return;
  const pid = baseOfTag(prev); const ps = find(pid); if (!ps?.next || ps.next === pid) return;
  const ns = find(ps.next); if (!ns) return;
  const base = ns.builtin ? ns.base : rootBase(ns.id);
  const tag = base === 'title' || base === 'subtitle' ? 'p' : base;
  let el = b;
  if (el.tagName.toLowerCase() !== tag) { const n = document.createElement(tag); n.innerHTML = el.innerHTML || '<br>'; el.replaceWith(n); el = n; const r = document.createRange(); r.setStart(n, 0); r.collapse(true); const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r); }
  el.classList.remove('title', 'subtitle'); if (base === 'title' || base === 'subtitle') el.classList.add(base);
  if (ns.builtin) delete el.dataset.style; else el.dataset.style = ns.id;
  el.removeAttribute('style');
}

// ------------------------------------------------------------------ crear, actualizar, modificar, eliminar
// formato directo del párrafo actual (y de su texto) como CSS de estilo
function cssFromSelection() {
  const b = C.currentBlock()?.closest('p,h1,h2,h3,h4,h5,h6,blockquote,li'); if (!b) return {};
  const out = {};
  for (const k of PARA_PROPS) { const v = b.style.getPropertyValue(k); if (v) out[k] = v; }
  if (b.dataset.push) { if (b.dataset.mt0) out['margin-top'] = b.dataset.mt0; else delete out['margin-top']; }
  // letra del texto donde está el cursor
  const sel = getSelection(); const n = sel.anchorNode?.nodeType === 3 ? sel.anchorNode.parentElement : sel.anchorNode;
  if (n && b.contains(n) && n !== b) {
    const cs = getComputedStyle(n), bs = getComputedStyle(b);
    for (const [k, prop] of [['font-family', 'fontFamily'], ['font-size', 'fontSize'], ['color', 'color'], ['font-weight', 'fontWeight'], ['font-style', 'fontStyle']]) if (cs[prop] !== bs[prop]) out[k] = k === 'font-size' ? `${r2(parseFloat(cs[prop]) * 0.75)}pt` : cs[prop];
  }
  return out;
}
function stripInline(id, css) {
  for (const el of C.editor.querySelectorAll(id === 'Normal' ? 'p:not([data-style]):not(.title):not(.subtitle)' : builtinOf(id) ? `${builtinOf(id).base.replace(/^(title|subtitle)$/, 'p.$1')}:not([data-style])` : `[data-style="${esc(id)}"]`)) {
    for (const k of Object.keys(css)) el.style.removeProperty(k);
    if (!el.getAttribute('style')) el.removeAttribute('style');
  }
}
export async function newFromSelection() {
  C.restoreSelection();
  const base = currentStyleId(); const css = cssFromSelection();
  const v = await form('Nuevo estilo a partir de la selección', [{ name: 'name', label: 'Nombre del estilo', value: '' }], 'Aceptar');
  if (!v?.name?.trim()) return;
  const name = v.name.trim();
  if (allPara().some((s) => s.name.toLowerCase() === name.toLowerCase())) { toast(tf('Ya existe un estilo llamado {0}.', name), { error: true }); return; }
  const id = `S${Math.random().toString(36).slice(2, 8)}`;
  styles()[id] = { id, name, kind: 'p', parent: base, next: id, css };
  render(); applyPara(id);
  toast(tf('Estilo {0} creado', name));
}
export function updateFromSelection(id = currentStyleId()) {
  C.restoreSelection();
  const css = cssFromSelection();
  const s = styles()[id] || (styles()[id] = { id, name: find(id)?.name || id, kind: 'p', css: {} });
  s.css = { ...(s.css || {}), ...css };
  stripInline(id, css); render(); C.shell.markDirty(); C.schedulePaginate(); C.renderSide?.();
  toast(tf('Estilo {0} actualizado', find(id)?.name || id));
}
export function deleteStyle(id) {
  const s = styles()[id]; if (!s || builtinOf(id) || BUILTIN_CHAR.some((b) => b.id === id)) return;
  const parent = s.parent || 'Normal';
  // los párrafos con ese estilo pasan al estilo del que heredaba
  C.editor.querySelectorAll(`[data-style="${esc(id)}"]`).forEach((el) => { if (builtinOf(parent)) delete el.dataset.style; else el.dataset.style = parent; });
  C.editor.querySelectorAll(`[data-cstyle="${esc(id)}"]`).forEach((el) => el.removeAttribute('data-cstyle'));
  for (const o of Object.values(styles())) if (o.parent === id) o.parent = parent;
  delete styles()[id];
  render(); C.shell.markDirty(); C.schedulePaginate(); C.renderSide?.();
}

// Cuadro Modificar estilo (Organizador, Fuente, Sangría y espaciado, Bordes)
export async function modifyStyle(id, kind = 'p') {
  const s = find(id); if (!s) return;
  // se calcula el aspecto actual con un elemento de prueba dentro del documento
  const probeTag = kind === 'c' ? 'span' : (s.builtin ? s.base : rootBase(id)).replace(/^(title|subtitle)$/, 'p');
  const probe = document.createElement(probeTag);
  if (kind === 'c') probe.dataset.cstyle = id; else { if (!s.builtin) probe.dataset.style = id; if (/^(title|subtitle)$/.test(s.base || '')) probe.className = s.base; }
  probe.textContent = 'x'; probe.style.position = 'absolute'; probe.style.visibility = 'hidden';
  C.editor.append(probe);
  const cs = getComputedStyle(probe);
  const cur = { font: cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(), size: r2(parseFloat(cs.fontSize) * 0.75), color: toHex(cs.color), bold: Number(cs.fontWeight) >= 600, italic: cs.fontStyle === 'italic', smallcaps: cs.fontVariantCaps === 'small-caps', caps: cs.textTransform === 'uppercase', ls: r2((parseFloat(cs.letterSpacing) || 0) * 0.75), mt: r2(parseFloat(cs.marginTop) * 0.75), mb: r2(parseFloat(cs.marginBottom) * 0.75), ml: r2(parseFloat(cs.marginLeft) / CM), mr: r2(parseFloat(cs.marginRight) / CM), ti: r2(parseFloat(cs.textIndent) / CM), lh: r2(parseFloat(cs.lineHeight) / parseFloat(cs.fontSize)) || 1.15, align: cs.textAlign === 'start' ? 'left' : cs.textAlign, border: cs.borderBottomStyle !== 'none' ? (cs.borderTopStyle !== 'none' ? 'box' : 'bottom') : 'none', bcolor: toHex(cs.borderBottomColor) || '#000000', bg: toHex(cs.backgroundColor) };
  probe.remove();
  const fonts = [...new Set([cur.font, ...C.FONTS.map((f) => f[0])])];
  const parents = (kind === 'c' ? allChar() : allPara()).filter((x) => x.id !== id);
  const fld = (label, el) => h('label', { class: 'fx-field' }, h('span', {}, t(label)), el);
  const inp = (name, value, type = 'number', step = '0.1') => h('input', { name, type, value, step });
  const sel = (name, value, opts) => h('select', { name }, ...opts.map(([v, l]) => h('option', { value: v, selected: String(v) === String(value) ? true : null }, t(l))));
  const chk = (name, label, on) => h('label', { class: 'fx-check' }, h('input', { type: 'checkbox', name, checked: on ? true : null }), ` ${t(label)}`);
  const org = h('div', { class: 'fx-grid' },
    fld('Nombre', h('input', { name: 'name', value: t(s.name), disabled: s.builtin ? true : null })),
    kind === 'p' ? fld('Estilo siguiente', sel('next', s.next || id, allPara().map((x) => [x.id, x.name]))) : h('span'),
    !s.builtin ? fld('Heredado de', sel('parent', s.parent || (kind === 'c' ? '' : 'Normal'), [...(kind === 'c' ? [['', '(ninguno)']] : []), ...parents.map((x) => [x.id, x.name])])) : h('p', { class: 'muted', style: { gridColumn: '1/-1' } }, t('Estilo predeterminado: se puede modificar, no cambiar de nombre ni eliminar.')));
  const font = h('div', { class: 'fx-grid' },
    fld('Fuente', sel('font', cur.font, fonts.map((f) => [f, f]))), fld('Tamaño (pt)', inp('size', cur.size, 'number', '0.5')),
    fld('Color', inp('color', cur.color || '#000000', 'color')), fld('Espaciado entre caracteres (pt)', inp('ls', cur.ls)),
    h('div', { class: 'fx-two', style: { gridColumn: '1/-1' } }, chk('bold', 'Negrita', cur.bold), chk('italic', 'Cursiva', cur.italic), chk('smallcaps', 'Versalitas', cur.smallcaps), chk('caps', 'Mayúsculas', cur.caps)));
  const para = h('div', { class: 'fx-grid' },
    fld('Antes del texto (cm)', inp('ml', cur.ml)), fld('Después del texto (cm)', inp('mr', cur.mr)), fld('Primera línea (cm)', inp('ti', cur.ti)),
    fld('Interlineado (múltiplo)', inp('lh', cur.lh, 'number', '0.05')), fld('Encima del párrafo (pt)', inp('mt', cur.mt, 'number', '1')), fld('Debajo del párrafo (pt)', inp('mb', cur.mb, 'number', '1')),
    fld('Alineación', sel('align', cur.align, [['left', 'Izquierda'], ['center', 'Centrado'], ['right', 'Derecha'], ['justify', 'Justificado']])));
  const bord = h('div', { class: 'fx-grid' },
    fld('Borde', sel('border', cur.border, [['none', 'Sin borde'], ['bottom', 'Solo abajo'], ['box', 'Cuadro']])), fld('Color del borde', inp('bcolor', cur.bcolor, 'color')),
    chk('bgOn', 'Color de fondo', !!cur.bg), fld('Fondo', inp('bg', cur.bg || '#F2F2F2', 'color')));
  const tabs = kind === 'c' ? [['Organizador', org], ['Fuente', font]] : [['Organizador', org], ['Fuente', font], ['Sangría y espaciado', para], ['Bordes', bord]];
  const heads = h('div', { class: 'fx-tabs' }); const panes = h('div', { class: 'fx-panes' });
  tabs.forEach(([l, el], i) => { heads.append(h('button', { type: 'button', class: `fx-tab${i ? '' : ' on'}`, onclick: () => { heads.querySelectorAll('.fx-tab').forEach((x, k) => x.classList.toggle('on', k === i)); panes.querySelectorAll('.fx-pane').forEach((x, k) => { x.hidden = k !== i; }); } }, t(l))); panes.append(h('div', { class: 'fx-pane', hidden: i ? true : null }, el)); });
  const body = h('div', { class: 'fx' }, heads, panes);
  const ok = await modal({ title: tf('Modificar estilo: {0}', t(s.name)), body, width: '560px', buttons: [{ label: 'Cancelar', value: null }, { label: 'Aceptar', primary: true, value: true }] });
  if (!ok) return;
  const v = (n) => { const el = body.querySelector(`[name="${n}"]`); if (!el) return undefined; return el.type === 'checkbox' ? el.checked : el.type === 'number' ? Number(el.value) : el.value; };
  const css = {
    'font-family': `'${v('font')}'`, 'font-size': `${v('size')}pt`, color: v('color'), 'font-weight': v('bold') ? 'bold' : 'normal', 'font-style': v('italic') ? 'italic' : 'normal',
    'font-variant': v('smallcaps') ? 'small-caps' : 'normal', 'text-transform': v('caps') ? 'uppercase' : 'none', 'letter-spacing': v('ls') ? `${v('ls')}pt` : 'normal',
  };
  if (kind === 'p') Object.assign(css, {
    'margin-left': `${v('ml')}cm`, 'margin-right': `${v('mr')}cm`, 'text-indent': `${v('ti')}cm`, 'line-height': String(v('lh') || 1.15), 'margin-top': `${v('mt')}pt`, 'margin-bottom': `${v('mb')}pt`, 'text-align': v('align'),
    'border-top': v('border') === 'box' ? `1pt solid ${v('bcolor')}` : 'none', 'border-left': v('border') === 'box' ? `1pt solid ${v('bcolor')}` : 'none', 'border-right': v('border') === 'box' ? `1pt solid ${v('bcolor')}` : 'none',
    'border-bottom': v('border') !== 'none' ? `1pt solid ${v('bcolor')}` : 'none', 'padding-bottom': v('border') !== 'none' ? '2pt' : '0', 'background-color': v('bgOn') ? v('bg') : 'transparent',
  });
  const rec = styles()[id] || (styles()[id] = { id, name: s.name, kind, css: {} });
  rec.css = css;
  if (!s.builtin) { rec.name = v('name')?.trim() || s.name; rec.parent = v('parent') || undefined; }
  if (kind === 'p') rec.next = v('next');
  stripInline(id, css); render(); C.shell.markDirty(); C.schedulePaginate(); C.renderSide?.();
}

// ------------------------------------------------------------------ panel de estilos (F11)
let paneKind = 'p'; let paneFilter = '';
export function renderPane() {
  const curId = currentStyleId();
  const list = (paneKind === 'p' ? allPara() : allChar()).filter((s) => !paneFilter || t(s.name).toLowerCase().includes(paneFilter.toLowerCase()));
  const used = new Set([...C.editor.querySelectorAll('[data-style]')].map((e) => e.dataset.style));
  const items = list.map((s) => {
    const css = s.builtin && paneKind === 'p' ? s.css : effectiveCss(s.id);
    const prevCss = Object.fromEntries(Object.entries(css).filter(([k]) => /^(font-family|font-weight|font-style|color|text-transform|font-variant)$/.test(k)));
    const tag = paneKind === 'p' ? (s.builtin ? s.base : rootBase(s.id)) : 'span';
    const sz = tag === 'h1' ? 16 : tag === 'h2' ? 14.5 : tag === 'title' ? 18 : tag === 'h3' ? 13.5 : 12.5;
    const b = h('div', { class: `st-item${s.id === curId ? ' on' : ''}${used.has(s.id) ? ' used' : ''}`, role: 'button', tabindex: '0', title: t('Doble clic para aplicar'), 'data-id': s.id },
      h('span', { class: `st-prev st-${tag}`, style: `${cssText(prevCss)};font-size:${sz}px` }, t(s.name)),
      h('button', { class: 'st-more', title: t('Opciones'), onclick: (e) => { e.stopPropagation(); dropdown(e.currentTarget, [
        { label: 'Aplicar', run: () => (paneKind === 'p' ? applyPara(s.id) : applyChar(s.id)) },
        { label: 'Modificar…', run: () => modifyStyle(s.id, paneKind) },
        ...(paneKind === 'p' ? [{ label: 'Actualizar con la selección', run: () => updateFromSelection(s.id) }] : []),
        ...(!s.builtin ? ['-', { label: 'Eliminar', run: () => deleteStyle(s.id) }] : []),
      ]); } }, '⋯'));
    b.addEventListener('mousedown', (e) => { if (!e.target.closest('.st-more')) e.preventDefault(); });
    b.addEventListener('click', (e) => { if (e.target.closest('.st-more')) return; if (paneKind === 'p') applyPara(s.id); else applyChar(s.id); });
    b.addEventListener('keydown', (e) => { if (e.key === 'Enter') b.click(); });
    return b;
  });
  const filter = h('input', { class: 'st-filter', placeholder: t('Buscar estilo…'), value: paneFilter });
  filter.addEventListener('input', () => { paneFilter = filter.value; const box = filter.closest('.st-pane'); box?.replaceWith(renderPane()); });
  return h('div', { class: 'st-pane' },
    h('div', { class: 'st-kinds' }, h('button', { class: paneKind === 'p' ? 'on' : '', onclick: () => { paneKind = 'p'; C.renderSide(); } }, t('Párrafo')), h('button', { class: paneKind === 'c' ? 'on' : '', onclick: () => { paneKind = 'c'; C.renderSide(); } }, t('Carácter'))),
    filter,
    h('div', { class: 'st-list' }, ...items, paneKind === 'c' ? h('div', { class: 'st-item', role: 'button', onclick: () => applyChar('__none') }, h('span', { class: 'st-prev' }, t('(Quitar estilo de carácter)'))) : ''),
    h('div', { class: 'st-actions' },
      h('button', { class: 'btn', onclick: () => newFromSelection() }, t('Nuevo a partir de la selección')),
      h('button', { class: 'btn', onclick: () => updateFromSelection() }, t('Actualizar estilo')),
      h('button', { class: 'btn', onclick: () => modifyStyle(paneKind === 'p' ? currentStyleId() : (getSelection().anchorNode?.parentElement?.closest('[data-cstyle]')?.dataset.cstyle || 'Emphasis'), paneKind) }, t('Modificar…'))));
}

// ------------------------------------------------------------------ listas
// Biblioteca: tipo de cada nivel ('1', 'a', 'A', 'i', 'I', viñeta) y si los números se encadenan (1.1.1)
export const LIST_PRESETS = {
  'n-1': { label: '1. 2. 3.', levels: ['decimal'] },
  'n-1p': { label: '1) 2) 3)', levels: ['decimal'], suffix: ')' },
  'n-a': { label: 'a. b. c.', levels: ['lower-alpha'] },
  'n-A': { label: 'A. B. C.', levels: ['upper-alpha'] },
  'n-i': { label: 'i. ii. iii.', levels: ['lower-roman'] },
  'n-I': { label: 'I. II. III.', levels: ['upper-roman'] },
  'm-1ai': { label: '1. / a. / i.', levels: ['decimal', 'lower-alpha', 'lower-roman'], multi: true },
  'm-legal': { label: '1. / 1.1. / 1.1.1.', levels: ['decimal', 'decimal', 'decimal'], multi: true, legal: true },
  'm-IA1': { label: 'I. / A. / 1. / a.', levels: ['upper-roman', 'upper-alpha', 'decimal', 'lower-alpha'], multi: true },
  'b-disc': { label: '•', bullet: '•' }, 'b-circle': { label: '◦', bullet: '◦' }, 'b-square': { label: '▪', bullet: '▪' },
  'b-check': { label: '✓', bullet: '✓' }, 'b-arrow': { label: '➢', bullet: '➢' }, 'b-dash': { label: '–', bullet: '–' },
};
const topList = (el) => { let l = el?.closest('ol,ul'); while (l && l.parentElement?.closest('ol,ul') && C.editor.contains(l.parentElement)) l = l.parentElement.closest('ol,ul'); return l; };
export function applyListPreset(key) {
  C.restoreSelection();
  const P = LIST_PRESETS[key]; if (!P) return;
  const want = P.bullet ? 'UL' : 'OL';
  let li = C.currentBlock()?.closest('li');
  if (!li || li.closest('ol,ul').tagName !== want) { C.listCmd(want === 'UL' ? 'insertUnorderedList' : 'insertOrderedList'); li = C.currentBlock()?.closest('li'); }
  const top = topList(li); if (!top) return;
  top.dataset.ml = key;
  top.style.listStyleType = P.bullet ? `'${P.bullet} '` : P.levels[0];
  top.querySelectorAll('ol,ul').forEach((sub) => { sub.style.listStyleType = ''; delete sub.dataset.ml; });
  render(); C.saveSelection(); C.shell.markDirty(); C.schedulePaginate();
}
// CSS de las listas de la biblioteca: formato por nivel, sufijo y números encadenados
function buildListCss() {
  const out = [];
  for (const [k, P] of Object.entries(LIST_PRESETS)) {
    if (P.bullet) { out.push(`.doc [data-ml="${k}"], .doc [data-ml="${k}"] ul { list-style-type: '${P.bullet} '; }`); continue; }
    for (let i = 0; i < 6; i++) {
      const sel = `.doc ol[data-ml="${k}"]${' ol'.repeat(i)}`;
      const fmt = P.levels[Math.min(i, P.levels.length - 1)];
      if (P.legal) out.push(`${sel} > li::marker { content: counters(list-item, ".") ". "; }`);
      else if (P.suffix) out.push(`${sel} > li::marker { content: counter(list-item, ${fmt}) "${P.suffix} "; }`);
      else out.push(`${sel} { list-style-type: ${fmt}; }`);
    }
  }
  return out.join('\n');
}
let LIST_CSS_CACHE = null;
const LIST_CSS_GET = () => (LIST_CSS_CACHE ||= buildListCss());
export function restartNumbering(at = 1) {
  C.restoreSelection();
  const li = C.currentBlock()?.closest('li'); const ol = li?.closest('ol'); if (!ol) { toast('Coloca el cursor en una lista numerada.'); return; }
  // la lista se parte en el elemento actual y vuelve a empezar
  const items = [...ol.children]; const i = items.indexOf(li);
  if (i > 0) { const nol = ol.cloneNode(false); items.slice(i).forEach((x) => nol.append(x)); ol.after(nol); nol.setAttribute('start', String(at)); }
  else if (at === 1) ol.removeAttribute('start'); else ol.setAttribute('start', String(at));
  C.shell.markDirty(); C.schedulePaginate();
}
export function continueNumbering() {
  C.restoreSelection();
  const ol = C.currentBlock()?.closest('ol'); if (!ol) return;
  let prev = ol.previousElementSibling; while (prev && prev.tagName !== 'OL') prev = prev.previousElementSibling;
  if (!prev) { ol.removeAttribute('start'); return; }
  const n = (Number(prev.getAttribute('start')) || 1) + prev.querySelectorAll(':scope > li').length;
  ol.setAttribute('start', String(n)); if (prev.dataset.ml) ol.dataset.ml = prev.dataset.ml;
  C.shell.markDirty(); C.schedulePaginate();
}
export async function setNumberingValue() {
  const v = await form('Establecer valor de numeración', [{ name: 'n', label: 'Empezar en', type: 'number', value: 1, min: 0 }], 'Aceptar');
  if (v) restartNumbering(Math.max(0, v.n | 0));
}
export function numberingMenu() {
  const nums = Object.entries(LIST_PRESETS).filter(([k]) => k.startsWith('n-') || k.startsWith('m-'));
  return [
    ...nums.map(([k, P]) => ({ label: P.label, run: () => applyListPreset(k) })),
    '-',
    { label: 'Reiniciar numeración', run: () => restartNumbering(1) },
    { label: 'Continuar numeración', run: () => continueNumbering() },
    { label: 'Establecer valor de numeración…', run: () => setNumberingValue() },
  ];
}
export const bulletMenu = () => Object.entries(LIST_PRESETS).filter(([k]) => k.startsWith('b-')).map(([k, P]) => ({ label: `${P.label}  ${t('Viñeta')}`, run: () => applyListPreset(k) }));

// ------------------------------------------------------------------ numeración de capítulos (títulos 1, 1.1, 1.1.1)
export const CHAPTER_PRESETS = {
  none: { label: 'Ninguna' },
  '1': { label: '1 / 1.1 / 1.1.1', fmts: ['decimal', 'decimal', 'decimal'], sep: '.', after: ' ' },
  '1.': { label: '1. / 1.1. / 1.1.1.', fmts: ['decimal', 'decimal', 'decimal'], sep: '.', after: '. ' },
  IA1: { label: 'I. / A. / 1.', fmts: ['upper-roman', 'upper-alpha', 'decimal'], own: true, after: '. ' },
  cap: { label: 'Capítulo 1 / 1.1 / 1.1.1', fmts: ['decimal', 'decimal', 'decimal'], sep: '.', after: ' ', prefix1: 'Capítulo ' },
};
function chapterCss(key) {
  const P = CHAPTER_PRESETS[key]; if (!P?.fmts) return '';
  const r = ['.doc.doc { counter-reset: ch1 ch2 ch3; }', '.doc.doc h1 { counter-increment: ch1; counter-reset: ch2 ch3; }', '.doc.doc h2 { counter-increment: ch2; counter-reset: ch3; }', '.doc.doc h3 { counter-increment: ch3; }'];
  for (let lv = 1; lv <= 3; lv++) {
    const parts = P.own ? [`counter(ch${lv}, ${P.fmts[lv - 1]})`] : Array.from({ length: lv }, (_, i) => `counter(ch${i + 1}, ${P.fmts[i]})`);
    const content = `${lv === 1 && P.prefix1 ? `"${t(P.prefix1)}" ` : ''}${parts.join(` "${P.sep || '.'}" `)} "${P.after}"`;
    r.push(`.doc.doc h${lv}::before { content: ${content}; }`);
  }
  return r.join('\n');
}
export function setChapterNumbering(key) {
  C.state.chapnum = key === 'none' ? null : key;
  render(); C.shell.markDirty(); C.schedulePaginate();
}
export const chapterMenu = () => Object.entries(CHAPTER_PRESETS).map(([k, P]) => ({ label: P.label, icon: (C.state.chapnum || 'none') === k ? 'check' : 'blank', run: () => setChapterNumbering(k) }));
// texto del número de un título (para el índice y la exportación)
export function chapterNumberText(levels, key = C.state.chapnum) {
  const P = CHAPTER_PRESETS[key]; if (!P?.fmts) return '';
  const fmt = (n, f) => (f === 'upper-roman' ? roman(n).toUpperCase() : f === 'lower-roman' ? roman(n) : f === 'upper-alpha' ? String.fromCharCode(64 + n) : f === 'lower-alpha' ? String.fromCharCode(96 + n) : String(n));
  const lv = levels.length;
  const body = P.own ? fmt(levels[lv - 1], P.fmts[lv - 1]) : levels.map((n, i) => fmt(n, P.fmts[i])).join(P.sep || '.');
  return `${lv === 1 && P.prefix1 ? t(P.prefix1) : ''}${body}${P.after}`;
}
const roman = (n) => { let s = ''; for (const [v, r] of [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]) while (n >= v) { s += r; n -= v; } return s; };

function toHex(c) {
  if (!c) return ''; if (c.startsWith('#')) return c.slice(0, 7);
  const m = /rgba?\(([^)]+)\)/.exec(c); if (!m) return '';
  const [r, g, b, a] = m[1].split(',').map((x) => parseFloat(x)); if (a === 0) return '';
  return `#${[r, g, b].map((n) => Math.round(n).toString(16).padStart(2, '0')).join('')}`;
}
