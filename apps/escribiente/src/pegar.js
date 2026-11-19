// Escribiente — opciones de pegado como en Word y OnlyOffice: después de pegar aparece el botón
// «(Ctrl)» junto al texto pegado para cambiar cómo quedó: mantener el formato original, combinar
// formato (conserva negritas, cursivas, listas, tablas y vínculos, pero toma la letra, el tamaño y
// los colores del documento) o mantener solo el texto. Se abre también con la tecla Ctrl sola, y
// cada opción tiene su letra (K, M, T). Ctrl+Mayús+V pega solo el texto. GPL-3.0
import { h, t } from '../../shared/office.js';
import { sanitize } from './docx-in.js';
import { icon } from '../../shared/icons.js';

let C = null; // { editor, scroller, shell, schedulePaginate }
export function init(ctx) { C = ctx; }

let last = null;   // { html, text, mode }
let ctl = null;    // el botón flotante
let busy = false;  // mientras se cambia la forma de pegado
let ctrlAlone = false;

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

// formato original (lo que ya hacía Escribiente)
// formato original (lo que ya hacía Escribiente). Se quitan los interlineados fijos de las páginas
// web (en px, pt, em o %), que con otra letra o tamaño encimaban las líneas.
export function keepHtml(html) {
  const clean = sanitize(html.replace(/<!--[\s\S]*?-->/g, '')).replace(/\sclass="(?!title|subtitle)[^"]*"/g, '');
  const doc = new DOMParser().parseFromString(`<body>${clean}</body>`, 'text/html');
  doc.body.querySelectorAll('[style]').forEach((el) => {
    const lh = el.style.lineHeight;
    if (lh && !(/^\d*\.?\d+$/.test(lh) && Number(lh) >= 1)) el.style.removeProperty('line-height');
    if (!el.getAttribute('style')?.trim()) el.removeAttribute('style');
  });
  return doc.body.innerHTML;
}

// combinar formato: conserva la estructura y el énfasis, quita letra, tamaño, colores y espaciados
export function mergeHtml(html) {
  const doc = new DOMParser().parseFromString(`<body>${keepHtml(html)}</body>`, 'text/html');
  const B = doc.body;
  // el énfasis escrito como estilo se convierte en etiquetas
  B.querySelectorAll('[style]').forEach((el) => {
    const st = el.style;
    const wrap = (tag) => { const w = doc.createElement(tag); while (el.firstChild) w.append(el.firstChild); el.append(w); };
    if (!/^(B|STRONG|H[1-6]|TH)$/.test(el.tagName) && (st.fontWeight === 'bold' || Number(st.fontWeight) >= 600)) wrap('b');
    if (!/^(I|EM)$/.test(el.tagName) && st.fontStyle === 'italic') wrap('i');
    if (/underline/.test(st.textDecoration || st.textDecorationLine || '') && el.tagName !== 'A' && el.tagName !== 'U') wrap('u');
    if (/line-through/.test(st.textDecoration || st.textDecorationLine || '') && el.tagName !== 'S') wrap('s');
    const align = /^(P|H[1-6]|LI|DIV|TD|TH|BLOCKQUOTE)$/.test(el.tagName) ? st.textAlign : '';
    el.removeAttribute('style');
    if (align && align !== 'start' && align !== 'left') el.style.textAlign = align;
  });
  B.querySelectorAll('font').forEach((f) => f.replaceWith(...f.childNodes));
  B.querySelectorAll('[color],[face],[size],[bgcolor],[class],[width],[height]:not(img)').forEach((el) => { for (const a of ['color', 'face', 'size', 'bgcolor', 'class']) el.removeAttribute(a); if (el.tagName !== 'IMG') { el.removeAttribute('width'); el.removeAttribute('height'); } });
  // los <span> sin atributos sobran
  B.querySelectorAll('span').forEach((s) => { if (!s.attributes.length) s.replaceWith(...s.childNodes); });
  return B.innerHTML;
}

// solo texto: cada línea es un párrafo
export function textHtml(text) {
  const lines = String(text || '').replace(/\r\n?/g, '\n').replace(/\n+$/, '').split('\n');
  if (lines.length <= 1) return null; // una línea: se inserta como texto
  return lines.map((l) => `<p>${esc(l).replace(/\t/g, '&emsp;') || '<br>'}</p>`).join('');
}
const htmlToText = (html) => { const d = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html'); d.querySelectorAll('p,div,li,h1,h2,h3,h4,h5,h6,tr,br').forEach((e) => e.after('\n')); d.querySelectorAll('td,th').forEach((e) => e.after('\t')); return d.body.textContent.replace(/\t\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim(); };

function insert(mode) {
  const { html, text } = last;
  if (mode === 'text') {
    const plain = text || htmlToText(html);
    const multi = textHtml(plain);
    if (multi) document.execCommand('insertHTML', false, multi); else document.execCommand('insertText', false, plain);
  } else document.execCommand('insertHTML', false, mode === 'merge' ? mergeHtml(html) : keepHtml(html));
  last.mode = mode;
}

// pega el contenido del portapapeles; html puede venir vacío (solo texto)
export function paste({ html = '', text = '' }, mode = 'keep') {
  hide();
  last = { html, text, mode };
  busy = true;
  try { insert(html ? mode : 'text'); } finally { busy = false; }
  C.shell.markDirty(); C.schedulePaginate();
  if (html) requestAnimationFrame(show);
}

// cambia la forma del último pegado: se deshace y se vuelve a pegar
export function repaste(mode) {
  if (!last || mode === last.mode) { closeMenu(); return; }
  C.editor.focus();
  busy = true;
  try { document.execCommand('undo'); insert(mode); } finally { busy = false; }
  C.shell.markDirty(); C.schedulePaginate();
  requestAnimationFrame(show);
}

const OPTS = [['keep', 'Mantener el formato original', 'K'], ['merge', 'Combinar formato', 'M'], ['text', 'Mantener solo texto', 'T']];
function show() {
  if (!last) return;
  const s = getSelection(); if (!s.rangeCount) return;
  const r = s.getRangeAt(0).cloneRange(); r.collapse(false);
  let rect = r.getBoundingClientRect();
  if (!rect.height) { const el = s.anchorNode?.nodeType === 1 ? s.anchorNode : s.anchorNode?.parentElement; rect = el?.getBoundingClientRect() || rect; }
  const sr = C.scroller.getBoundingClientRect();
  if (!ctl) {
    ctl = h('div', { class: 'paste-ctl', role: 'group', 'aria-label': t('Opciones de pegado') },
      h('button', { type: 'button', class: 'pc-btn', title: t('Opciones de pegado (Ctrl)'), 'aria-haspopup': 'menu', onmousedown: (e) => e.preventDefault(), onclick: () => toggleMenu() },
        h('span', { class: 'pc-ico', 'aria-hidden': 'true', html: icon('paste', 14) }), h('span', {}, `(${t('Ctrl')})`), h('span', { class: 'pc-car', 'aria-hidden': 'true' }, '▾')),
      h('div', { class: 'pc-menu', role: 'menu', hidden: true }));
    C.scroller.append(ctl);
  }
  const menu = ctl.querySelector('.pc-menu');
  menu.replaceChildren(...OPTS.map(([m, label, k]) => h('button', { type: 'button', role: 'menuitemradio', 'aria-checked': String(last.mode === m), 'data-mode': m, class: `pc-opt${last.mode === m ? ' on' : ''}`, onmousedown: (e) => e.preventDefault(), onclick: () => repaste(m) },
    h('span', { class: 'pc-chk', 'aria-hidden': 'true' }, last.mode === m ? '✓' : ''), `${t(label)} (${k})`)));
  menu.hidden = true;
  ctl.style.left = `${Math.round(rect.right - sr.left + C.scroller.scrollLeft + 4)}px`;
  ctl.style.top = `${Math.round(rect.bottom - sr.top + C.scroller.scrollTop + 4)}px`;
  ctl.hidden = false;
}
export function hide() { if (ctl) { ctl.hidden = true; closeMenu(); } }
export function dismiss() { hide(); last = null; }
function closeMenu() { const m = ctl?.querySelector('.pc-menu'); if (m) m.hidden = true; }
function toggleMenu() { const m = ctl?.querySelector('.pc-menu'); if (!m) return; m.hidden = !m.hidden; if (!m.hidden) m.querySelector('.pc-opt.on, .pc-opt')?.focus({ preventScroll: true }); }
export const isOpen = () => !!(ctl && !ctl.hidden);
export const menuOpen = () => !!(ctl && !ctl.hidden && !ctl.querySelector('.pc-menu').hidden);
export const lastMode = () => last?.mode || null;

// teclado: Ctrl solo abre el menú; K, M, T eligen; Esc lo cierra. Cualquier otra edición lo quita.
export function onKeyDown(e) {
  if (!isOpen()) return false;
  if (e.key === 'Control' || e.key === 'Meta') { ctrlAlone = true; return false; }
  ctrlAlone = false;
  if (menuOpen()) {
    const k = e.key.toLowerCase();
    const opt = OPTS.find((o) => o[2].toLowerCase() === k);
    if (opt && !e.ctrlKey && !e.metaKey && !e.altKey) { e.preventDefault(); repaste(opt[0]); C.editor.focus(); return true; }
    if (e.key === 'Escape') { e.preventDefault(); closeMenu(); C.editor.focus(); return true; }
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault(); const items = [...ctl.querySelectorAll('.pc-opt')]; const i = items.indexOf(document.activeElement);
      items[(i + (e.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length].focus(); return true;
    }
    if (e.key === 'Enter' && document.activeElement?.classList.contains('pc-opt')) { e.preventDefault(); document.activeElement.click(); C.editor.focus(); return true; }
    return false;
  }
  if (e.key === 'Escape') { e.preventDefault(); dismiss(); return true; }
  if (!['Shift', 'Alt', 'CapsLock'].includes(e.key) && !(e.ctrlKey || e.metaKey)) dismiss();
  return false;
}
export function onKeyUp(e) {
  if ((e.key === 'Control' || e.key === 'Meta') && ctrlAlone && isOpen()) { ctrlAlone = false; toggleMenu(); }
}
export function onEdited() { if (!busy && isOpen()) dismiss(); }

// el menú también recibe las teclas cuando tiene el foco
document.addEventListener('keydown', (e) => { if (ctl && ctl.contains(document.activeElement)) onKeyDown(e); }, true);
document.addEventListener('mousedown', (e) => { if (isOpen() && !ctl.contains(e.target)) dismiss(); }, true);
