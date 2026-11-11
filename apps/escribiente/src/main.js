// Escribiente — procesador de textos de Mi Oficina MX. GPL-3.0-or-later
import { OfficeShell, h, $, $$, icon, toast, modal, form, dropdown, esc, api, closeDropdown, localizeHtml, tf, t } from '../../shared/office.js';
import { TEMPLATES, thumb } from './templates.js';
import { htmlToDocx } from './docx-out.js';
import { docxToHtml, textToHtml, markdownToHtml, sanitize, htmlToText } from './docx-in.js';
import { readCoreProps } from '../../shared/coreprops.js';
import * as F from './features.js';
import * as F3 from './fase3.js';
import * as FX from './formato.js';
import * as ES from './estilos.js';
import { odtToHtml, htmlToOdt } from './odt.js';

const CM = 37.7953; // px por cm
const PAGE_SIZES = {
  carta: { name: 'Carta (21.6 × 27.9 cm)', w: 816, h: 1056 },
  oficio: { name: 'Oficio México (21.6 × 34 cm)', w: 816, h: 1285 },
  legal: { name: 'Legal (21.6 × 35.6 cm)', w: 816, h: 1344 },
  a4: { name: 'A4 (21 × 29.7 cm)', w: 794, h: 1123 },
  a5: { name: 'A5 (14.8 × 21 cm)', w: 559, h: 794 },
};
const FONTS = [
  ['Calibri', 'Calibri, Carlito, "Liberation Sans", sans-serif'],
  ['Cambria', 'Cambria, Caladea, "Liberation Serif", serif'],
  ['Arial', 'Arial, "Liberation Sans", Helvetica, sans-serif'],
  ['Times New Roman', '"Times New Roman", "Liberation Serif", Times, serif'],
  ['Georgia', 'Georgia, "DejaVu Serif", serif'],
  ['Verdana', 'Verdana, "DejaVu Sans", sans-serif'],
  ['Tahoma', 'Tahoma, "DejaVu Sans", sans-serif'],
  ['Trebuchet MS', '"Trebuchet MS", "DejaVu Sans", sans-serif'],
  ['Garamond', 'Garamond, "EB Garamond", serif'],
  ['Courier New', '"Courier New", "Liberation Mono", monospace'],
  ['Inter', 'Inter, sans-serif'],
  ['Noto Sans', '"Noto Sans", sans-serif'],
  ['DejaVu Sans', '"DejaVu Sans", sans-serif'],
];
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 60, 72];
const BLOCK_STYLES = [
  ['p', 'Normal'], ['title', 'Título'], ['subtitle', 'Subtítulo'], ['h1', 'Título 1'], ['h2', 'Título 2'], ['h3', 'Título 3'], ['h4', 'Título 4'], ['blockquote', 'Cita'],
];
const BLOCK_TAGS = /^(P|H[1-6]|BLOCKQUOTE|LI|DIV|TD|TH)$/;

// ------------------------------------------------------------------ estado del documento
const state = {
  page: { size: 'carta', w: 816, h: 1056, mt: 2.5 * CM, mb: 2.5 * CM, ml: 3 * CM, mr: 3 * CM, landscape: false, gap: 24 },
  header: '', footer: '', pageNumbers: 'footer-center',
  comments: [], footnotes: {}, endnotes: {}, track: false, watermark: null, pageColor: null, merge: null, pageFns: [], enPos: null,
  zoom: 100, pages: 1, pageOf: [],
};

// ------------------------------------------------------------------ DOM del área de trabajo
let shell, editor, pagesEl, scroller, ruler, statusPages, statusWords, imgBox, selImg = null;

function buildWorkspace(ws) {
  scroller = h('div', { class: 'scroller', role: 'region', 'aria-label': 'Páginas del documento', tabindex: '0' });
  ruler = h('div', { class: 'ruler', 'aria-hidden': 'true' });
  pagesEl = h('div', { class: 'pages' });
  editor = h('div', { id: 'editor', class: 'doc', contenteditable: 'true', spellcheck: 'true', role: 'textbox', 'aria-multiline': 'true', 'aria-label': 'Documento' });
  imgBox = h('div', { class: 'imgbox', hidden: true }, ...['nw', 'ne', 'sw', 'se'].map((c) => h('div', { class: `hd ${c}`, 'data-h': c })));
  pagesEl.append(editor, imgBox);
  scroller.append(ruler, pagesEl);
  ws.append(scroller, h('div', { id: 'printRoot' }));

  editor.addEventListener('input', (e) => { if (e.inputType === 'insertParagraph') ES.onNewParagraph(); shell.markDirty(); schedulePaginate(); updateCounts(); });
  editor.addEventListener('keydown', onEditorKey);
  editor.addEventListener('paste', onPaste);
  editor.addEventListener('click', (e) => {
    if (e.target.tagName === 'IMG') selectImage(e.target); else selectImage(null);
    const a = e.target.closest('a');
    if (a && (e.ctrlKey || e.metaKey)) { e.preventDefault(); window.open(a.href); }
  });
  pagesEl.addEventListener('mousedown', (e) => {
    // Clic en el margen o debajo del texto: el cursor va al final más cercano
    if (e.target === pagesEl || e.target.classList.contains('sheet')) {
      e.preventDefault();
      editor.focus();
      placeCaretAtEnd();
    }
  });
  document.addEventListener('selectionchange', () => { if (editor.contains(document.getSelection().anchorNode)) { saveSelection(); reflectState(); F.refreshStylesPane(); } });
  scroller.addEventListener('scroll', () => { updateStatusPage(); positionImgBox(); });
  new ResizeObserver(() => schedulePaginate()).observe(editor);
  imgBox.addEventListener('mousedown', startImgResize);
}

// ------------------------------------------------------------------ selección
let savedRange = null;
function saveSelection() {
  const s = document.getSelection();
  if (s.rangeCount && editor.contains(s.anchorNode)) savedRange = s.getRangeAt(0).cloneRange();
}
function restoreSelection() {
  if (document.activeElement !== editor) editor.focus({ preventScroll: true });
  if (savedRange) { const s = document.getSelection(); s.removeAllRanges(); s.addRange(savedRange); }
}
function placeCaretAtEnd() {
  const r = document.createRange();
  const last = editor.lastElementChild || editor;
  r.selectNodeContents(last); r.collapse(false);
  const s = document.getSelection(); s.removeAllRanges(); s.addRange(r);
}
function currentBlock() {
  const s = document.getSelection();
  let n = s.anchorNode;
  if (!n || !editor.contains(n)) return null;
  if (n.nodeType === 3) n = n.parentElement;
  while (n && n !== editor && !BLOCK_TAGS.test(n.tagName)) n = n.parentElement;
  return n === editor ? null : n;
}
function selectedBlocks() {
  const s = document.getSelection();
  if (!s.rangeCount) return [];
  const r = s.getRangeAt(0);
  const all = [...editor.querySelectorAll('p,h1,h2,h3,h4,h5,h6,blockquote,li,td,th')].filter((b) => !b.querySelector('p,h1,h2,h3,h4,h5,h6,li'));
  const out = all.filter((b) => r.intersectsNode(b));
  if (!out.length) { const b = currentBlock(); if (b) out.push(b); }
  return out;
}
const exec = (cmd, val = null) => { restoreSelection(); if (/List$/.test(cmd)) listCmd(cmd); else { document.execCommand(cmd, false, val); if (/dent$/.test(cmd)) fixLists(); } saveSelection(); shell.markDirty(); schedulePaginate(); reflectState(); };
// listas: Chromium mueve el cursor al principio; se deja donde estaba
function listCmd(cmd) {
  const s = getSelection(); const r0 = s.rangeCount ? s.getRangeAt(0) : null;
  // posición como número de caracteres desde el principio del párrafo (Chromium rehace los nodos de texto)
  const blk = r0 && (r0.startContainer.nodeType === 3 ? r0.startContainer.parentElement : r0.startContainer).closest('p,h1,h2,h3,h4,h5,h6,li,blockquote,div');
  let off = -1;
  if (blk && r0.collapsed && editor.contains(blk)) { const pr = document.createRange(); pr.selectNodeContents(blk); pr.setEnd(r0.startContainer, r0.startOffset); off = pr.toString().length; }
  document.execCommand(cmd); fixLists();
  if (off < 0) return;
  const s2 = getSelection(); const n = s2.anchorNode; const li = n && (n.nodeType === 3 ? n.parentElement : n).closest('li,p,h1,h2,h3,h4,h5,h6,blockquote,div');
  if (!li || !editor.contains(li)) return;
  const tw = document.createTreeWalker(li, NodeFilter.SHOW_TEXT, { acceptNode: (x) => (x.parentElement.closest('ol,ul') !== li.closest('ol,ul') && li.tagName === 'LI' ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) });
  let left = off; let tn;
  while ((tn = tw.nextNode())) { if (left <= tn.length) { const r = document.createRange(); r.setStart(tn, left); r.collapse(true); s2.removeAllRanges(); s2.addRange(r); return; } left -= tn.length; }
}
// Chromium deja la lista dentro del párrafo (<p><ol>…</ol></p>): se saca la lista y se conserva el texto de alrededor
function fixLists() {
  const sel = getSelection(); const r0 = sel.rangeCount ? sel.getRangeAt(0) : null;
  const keep = r0 && { sc: r0.startContainer, so: r0.startOffset, ec: r0.endContainer, eo: r0.endOffset }; // nodos, no un rango vivo (se movería al mover la lista)
  for (const l of [...editor.querySelectorAll('p > ol, p > ul, h1 > ol, h1 > ul, h2 > ol, h2 > ul, h3 > ol, h3 > ul, h4 > ol, h4 > ul, h5 > ol, h5 > ul, h6 > ol, h6 > ul, blockquote > ol, blockquote > ul')]) {
    const p = l.parentElement; if (!p || p === editor) continue;
    const after = p.cloneNode(false); let n = l.nextSibling; while (n) { const nx = n.nextSibling; after.append(n); n = nx; }
    p.after(l); l.after(after);
    const empty = (x) => !x.textContent.trim() && !x.querySelector('img,span.eq,sup');
    if (empty(after)) after.remove(); if (empty(p)) p.remove();
  }
  // elementos de lista dentro de otro elemento (al quitar sangría): pasan detrás
  for (const li of [...editor.querySelectorAll('li > li')]) { const par = li.parentElement; let n = li; const move = []; while (n) { move.push(n); n = n.nextSibling; } par.after(...move); }
  if (keep && editor.contains(keep.sc) && editor.contains(keep.ec)) { try { const r = document.createRange(); r.setStart(keep.sc, keep.so); r.setEnd(keep.ec, keep.eo); sel.removeAllRanges(); sel.addRange(r); } catch { /* */ } }
}

// Aplica un estilo de texto (tamaño, fuente) a la selección mediante <span style>
// Aplica varias propiedades CSS a la selección (cadena vacía = quitar la propiedad)
function applySpanStyles(styles, { hidden, attrs } = {}) {
  restoreSelection();
  const s = document.getSelection();
  if (!s.rangeCount) return;
  const setOn = (span) => {
    for (const [k, v] of Object.entries(styles)) { if (v === '' || v == null) span.style[k] = ''; else span.style[k] = v; }
    if (hidden != null) span.classList.toggle('hidden-text', !!hidden);
    for (const [k, v] of Object.entries(attrs || {})) { if (v == null) span.removeAttribute(k); else span.setAttribute(k, v); span.querySelectorAll(`[${k}]`).forEach((d) => d.removeAttribute(k)); }
    if (!span.getAttribute('style')) span.removeAttribute('style');
    // los descendientes con el mismo estilo directo dejan de pisarlo
    span.querySelectorAll('span').forEach((d) => { for (const k of Object.keys(styles)) d.style[k] = ''; if (hidden != null) d.classList.remove('hidden-text'); if (!d.getAttribute('style')) d.removeAttribute('style'); });
  };
  if (s.isCollapsed) {
    const span = document.createElement('span'); setOn(span); span.textContent = '\u200B';
    const r = s.getRangeAt(0); r.insertNode(span);
    r.setStart(span.firstChild, 1); r.collapse(true); s.removeAllRanges(); s.addRange(r);
  } else {
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('fontSize', false, '7');
    const marked = [...editor.querySelectorAll('[style*="xxx-large"], font[size="7"]')];
    for (const el of marked) {
      const span = document.createElement('span');
      if (el.tagName === 'SPAN' || el.tagName === 'FONT') {
        span.setAttribute('style', (el.getAttribute('style') || '').replace(/font-size:\s*xxx-large;?/, ''));
        span.innerHTML = el.innerHTML; el.replaceWith(span);
      } else {
        // Chrome puso el tamaño en un <b>/<i>/<u>…: se quita de ahí y el estilo va en un span interior
        el.style.fontSize = ''; if (!el.getAttribute('style')) el.removeAttribute('style');
        span.append(...el.childNodes); el.append(span);
      }
      setOn(span);
    }
  }
  saveSelection(); shell.markDirty(); schedulePaginate(); reflectState();
}
function applySpanStyle(prop, value) {
  restoreSelection();
  const s = document.getSelection();
  if (!s.rangeCount) return;
  document.execCommand('styleWithCSS', false, true);
  if (s.isCollapsed) {
    // Sin selección: se inserta un span vacío para que lo siguiente que se escriba tenga ese estilo.
    const span = document.createElement('span');
    span.style[prop] = value;
    span.textContent = '\u200B';
    const r = s.getRangeAt(0); r.insertNode(span);
    r.setStart(span.firstChild, 1); r.collapse(true); s.removeAllRanges(); s.addRange(r);
  } else {
    document.execCommand('fontSize', false, '7');
    editor.querySelectorAll('span[style*="xxx-large"], font[size="7"]').forEach((el) => {
      const span = document.createElement('span');
      span.setAttribute('style', (el.getAttribute('style') || '').replace(/font-size:\s*xxx-large;?/, ''));
      span.style[prop] = value;
      if (prop !== 'fontSize' && el.dataset.keepSize) span.style.fontSize = el.dataset.keepSize;
      span.innerHTML = el.innerHTML;
      el.replaceWith(span);
    });
  }
  saveSelection(); shell.markDirty(); schedulePaginate(); reflectState();
}
function applyFontFamily(name) {
  const stack = (FONTS.find((f) => f[0] === name) || [name, `"${name}"`])[1];
  restoreSelection();
  document.execCommand('styleWithCSS', false, true);
  document.execCommand('fontName', false, stack);
  saveSelection(); shell.markDirty(); schedulePaginate(); reflectState();
}
function setBlockStyle(kind) {
  restoreSelection();
  if (['p', 'h1', 'h2', 'h3', 'h4', 'blockquote'].includes(kind)) document.execCommand('formatBlock', false, kind);
  else { document.execCommand('formatBlock', false, 'p'); }
  selectedBlocks().forEach((b) => {
    if (b.tagName === 'P') { b.classList.remove('title', 'subtitle'); if (kind === 'title' || kind === 'subtitle') b.classList.add(kind); }
  });
  saveSelection(); shell.markDirty(); schedulePaginate(); reflectState();
}
function adjustSize(delta) {
  const cur = parseFloat(currentFontSizePt()) || 11;
  const idx = SIZES.findIndex((s) => s >= cur);
  const next = delta > 0 ? SIZES[Math.min(SIZES.length - 1, (SIZES[idx] === cur ? idx + 1 : idx))] : SIZES[Math.max(0, (idx === -1 ? SIZES.length : idx) - 1)];
  applySpanStyle('fontSize', `${next}pt`);
}
function currentFontSizePt() {
  const s = document.getSelection();
  let n = s.anchorNode; if (!n) return '11';
  if (n.nodeType === 3) n = n.parentElement;
  if (!editor.contains(n)) return '11';
  const px = parseFloat(getComputedStyle(n).fontSize);
  return String(Math.round(px * 0.75 * 2) / 2);
}
function setLineHeight(v) {
  selectedBlocks().forEach((b) => { b.style.lineHeight = v; });
  shell.markDirty(); schedulePaginate();
}
function setIndent(delta) {
  const b = selectedBlocks();
  if (b.some((x) => x.tagName === 'LI')) { exec(delta > 0 ? 'indent' : 'outdent'); return; }
  b.forEach((x) => { const cur = x.style.marginLeft ? parseFloat(getComputedStyle(x).marginLeft) || 0 : 0; const n = Math.max(0, cur + delta * 48); x.style.marginLeft = n ? `${n}px` : ''; });
  shell.markDirty(); schedulePaginate();
}

// ------------------------------------------------------------------ teclado y pegado
function onEditorKey(e) {
  const mod = e.ctrlKey || e.metaKey;
  if (mod && !e.shiftKey && !e.altKey) {
    const k = e.key.toLowerCase();
    if (k === 'b') { e.preventDefault(); exec('bold'); }
    else if (k === 'i') { e.preventDefault(); exec('italic'); }
    else if (k === 'u') { e.preventDefault(); exec('underline'); }
    else if (k === 'e') { e.preventDefault(); exec('justifyCenter'); }
    else if (k === 'l') { e.preventDefault(); exec('justifyLeft'); }
    else if (k === 'r') { e.preventDefault(); exec('justifyRight'); }
    else if (k === 'j') { e.preventDefault(); exec('justifyFull'); }
    else if (k === 'k') { e.preventDefault(); insertLink(); }
    else if (k === 'enter') { e.preventDefault(); insertPageBreak(); }
    else if (k === ']') { e.preventDefault(); adjustSize(1); }
    else if (k === '[') { e.preventDefault(); adjustSize(-1); }
    else if (k === 'd' && !e.shiftKey) { e.preventDefault(); saveSelection(); FX.fontDialog(); }
    else if (k === 'f10') { e.preventDefault(); onCommand('marks'); }
  }
  if (mod && e.key === 'F11') { e.preventDefault(); onCommand('stylesPane'); }
  if (e.shiftKey && !mod && e.key === 'F3') { e.preventDefault(); saveSelection(); const sel = document.getSelection().toString(); FX.changeCase(sel === sel.toLocaleUpperCase('es') ? 'lower' : sel === sel.toLocaleLowerCase('es') ? 'title' : 'upper'); }
  if (mod && e.altKey && /^[1-3]$/.test(e.key)) { e.preventDefault(); setBlockStyle(`h${e.key}`); }
  if (mod && e.altKey && e.key === '0') { e.preventDefault(); setBlockStyle('p'); }
  if (e.key === 'Tab') {
    const b = currentBlock();
    const cell = b && b.closest('td,th');
    if (cell) { e.preventDefault(); moveCell(cell, e.shiftKey ? -1 : 1); return; }
    if (b && b.tagName === 'LI') { e.preventDefault(); exec(e.shiftKey ? 'outdent' : 'indent'); return; }
    e.preventDefault(); if (e.shiftKey) return; F.insertTab(); shell.markDirty(); schedulePaginate();
  }
  if ((e.key === 'Delete' || e.key === 'Backspace') && selImg) { e.preventDefault(); selImg.remove(); selectImage(null); shell.markDirty(); schedulePaginate(); }
  // Evita que el editor quede vacío sin párrafo
  if (e.key === 'Backspace' || e.key === 'Delete') setTimeout(() => { if (!editor.firstElementChild) editor.innerHTML = '<p><br></p>'; }, 0);
}
function moveCell(cell, dir) {
  const cells = [...cell.closest('table').querySelectorAll('td,th')];
  let i = cells.indexOf(cell) + dir;
  if (i >= cells.length && dir > 0) { tableOp('rowBelow'); return moveCell(cell, dir); }
  i = Math.max(0, Math.min(cells.length - 1, i));
  const r = document.createRange(); r.selectNodeContents(cells[i]);
  const s = document.getSelection(); s.removeAllRanges(); s.addRange(r);
}
function onPaste(e) {
  const html = e.clipboardData.getData('text/html');
  if (html) {
    e.preventDefault();
    const clean = sanitize(html.replace(/<!--[\s\S]*?-->/g, '')).replace(/\sclass="(?!title|subtitle)[^"]*"/g, '');
    document.execCommand('insertHTML', false, clean);
  } else if ([...e.clipboardData.items].some((i) => i.type.startsWith('image/'))) {
    e.preventDefault();
    const f = [...e.clipboardData.items].find((i) => i.type.startsWith('image/')).getAsFile();
    const rd = new FileReader();
    rd.onload = () => insertImageDataUrl(rd.result);
    rd.readAsDataURL(f);
  }
}

// ------------------------------------------------------------------ paginación
let paginateRaf = 0;
function schedulePaginate() { cancelAnimationFrame(paginateRaf); paginateRaf = requestAnimationFrame(paginate); }

function pageDims() {
  const p = state.page;
  return { W: p.w, Hp: p.h, H: p.h - p.mt - p.mb, S: p.h + p.gap, cw: p.w - p.ml - p.mr };
}

function paginate() {
  const p = state.page;
  const { W, Hp, H, S, cw } = pageDims();
  editor.style.left = `${p.ml}px`;
  editor.style.top = `${p.mt}px`;
  editor.style.width = `${cw}px`;
  pagesEl.style.width = `${W}px`;
  F.layoutTabs();
  const blocks = [...editor.children];
  // se quita el empuje anterior y se restaura el margen superior propio del párrafo
  for (const b of blocks) if (b.dataset.push) { b.style.marginTop = b.dataset.mt0 || ''; delete b.dataset.push; delete b.dataset.mt0; }
  const tops = blocks.map((b) => b.offsetTop);
  const hs = blocks.map((b) => b.offsetHeight);
  // se leen todos los márgenes antes de escribir estilos (leerlos dentro del bucle forzaba recalcular la página en cada salto: muy lento en documentos largos)
  const css = blocks.map((b) => getComputedStyle(b));
  const mbs = css.map((c) => parseFloat(c.marginBottom) || 0);
  const mts = css.map((c) => parseFloat(c.marginTop) || 0);
  const shiftAt = [];
  let shift = 0, forceNext = false;
  const pageOf = [];
  // notas al pie: reducen el espacio de la página donde aparece su llamada
  const fnOrder = F.numberFootnotes();
  const fnH = fnOrder.length ? F.footnoteHeights(cw) : {};
  const pageFns = []; const pageFnH = [];
  const fnsOf = (b) => [...new Set([...(b.matches('sup.fn') ? [b] : []), ...b.querySelectorAll('sup.fn')].map((x) => x.dataset.fid))].filter((f) => state.footnotes[f] != null);
  const extraFor = (page, fids) => { const cur = pageFns[page] || []; const add = fids.filter((f) => !cur.includes(f)); if (!add.length) return 0; return (cur.length ? 0 : 14) + add.reduce((a, f) => a + (fnH[f] || 16), 0); };
  const pushBlock = (j, add) => {
    const bj = blocks[j];
    const prevMB = j > 0 ? mbs[j - 1] : 0;
    if (!bj.dataset.push) bj.dataset.mt0 = bj.style.marginTop || '';
    bj.style.marginTop = `${Math.max(prevMB, mts[j]) + add}px`;
    bj.dataset.push = String(add);
  };
  blocks.forEach((b, i) => {
    shiftAt[i] = shift;
    let top = tops[i] + shift;
    let page = Math.floor(top / S);
    const fids = fnOrder.length ? fnsOf(b) : [];
    const zoneEnd = page * S + H - (pageFnH[page] || 0) - extraFor(page, fids);
    let add = 0;
    if (b.classList.contains('pb')) { pageOf.push(page); forceNext = true; return; }
    let overflow = false;
    if (forceNext || (b.dataset.pbb && i > 0 && pageOf[i - 1] === page && !blocks[i - 1].classList.contains('pb'))) { add = (page + 1) * S - top; forceNext = false; }
    else if (top >= zoneEnd - 0.5) { add = (page + 1) * S - top; overflow = true; }
    else if (top + hs[i] > zoneEnd && hs[i] <= H && top > page * S + 1) { add = (page + 1) * S - top; overflow = true; }
    // conservar con el siguiente: los párrafos marcados se van a la página nueva junto con este
    if (overflow && add > 0) {
      let j = i;
      while (j > 0 && blocks[j - 1].dataset.keepnext && pageOf[j - 1] === page && !blocks[j - 1].dataset.push && !blocks[j - 1].classList.contains('pb')) j--;
      const chainTop = tops[j] + shiftAt[j];
      if (j < i && chainTop > Math.floor(chainTop / S) * S + 1 && (tops[i] + hs[i]) - tops[j] <= H) {
        const pg = Math.floor(chainTop / S);
        const addJ = (pg + 1) * S - chainTop;
        pushBlock(j, addJ);
        shift = shiftAt[j] + addJ;
        for (let k = j; k < i; k++) pageOf[k] = pg + 1;
        top = tops[i] + shift; page = Math.floor(top / S); add = 0;
      }
    }
    if (add > 0) {
      pushBlock(i, add);
      shift += add; top += add; page = Math.floor(top / S);
    }
    if (fids.length) { const ex = extraFor(page, fids); pageFns[page] = [...new Set([...(pageFns[page] || []), ...fids])]; pageFnH[page] = (pageFnH[page] || 0) + ex; }
    pageOf.push(page);
  });
  state.pageFns = pageFns;
  const lastBottom = blocks.length ? tops[blocks.length - 1] + shift + hs[blocks.length - 1] : 0;
  let n = Math.max(1, Math.floor(Math.max(0, lastBottom - 1) / S) + 1, (pageOf[pageOf.length - 1] ?? 0) + 1);
  // notas al final: después del último párrafo; si no caben, en la página siguiente
  state.enPos = null;
  const enOrder = F.endnoteOrder();
  if (enOrder.length) {
    const eh = F.endnotesHeight(cw);
    let start = lastBottom + 24; let pg = Math.floor(start / S);
    if (start + eh > pg * S + H - (pageFnH[pg] || 0) && start > pg * S + 1) { pg += 1; start = pg * S; }
    state.enPos = { page: pg, y: start - pg * S };
    n = Math.max(n, pg + 1);
  }
  state.pages = n; state.pageOf = pageOf;
  // hojas
  const sheets = $$('.sheet', pagesEl);
  for (let i = sheets.length; i < n; i++) pagesEl.insertBefore(h('div', { class: 'sheet' }), editor);
  $$('.sheet', pagesEl).forEach((sh, i) => {
    if (i >= n) { sh.remove(); return; }
    sh.style.top = `${i * S}px`; sh.style.width = `${W}px`; sh.style.height = `${Hp}px`;
    renderHeaderFooter(sh, i, n);
  });
  pagesEl.style.height = `${n * S - p.gap}px`;
  editor.style.minHeight = `${H}px`;
  applyZoom();
  drawRuler();
  updateStatusPage();
  positionImgBox();
  clearTimeout(paginate.navT); paginate.navT = setTimeout(() => { F.renderNav(); }, 400);
}

function hfText(tpl, i, n) { return (tpl || '').replace(/\{pagina\}/gi, String(i + 1)).replace(/\{paginas\}/gi, String(n)); }
function renderHeaderFooter(sh, i, n) {
  const p = state.page;
  sh.replaceChildren();
  const pad = { left: `${p.ml}px`, right: `${p.mr}px` };
  if (state.header) sh.append(h('div', { class: 'hf top', style: { ...pad, top: `${Math.max(12, p.mt / 2 - 10)}px` } }, hfText(state.header, i, n)));
  if (state.footer) sh.append(h('div', { class: 'hf bottom', style: { ...pad, bottom: `${Math.max(12, p.mb / 2 - 10)}px` } }, hfText(state.footer, i, n)));
  const pn = state.pageNumbers;
  if (state.pageColor) sh.style.background = state.pageColor; else sh.style.background = '';
  if (state.watermark?.text) sh.append(h('div', { class: 'wm', style: { color: state.watermark.color || '#C9CBD6' } }, state.watermark.text));
  const fns = state.pageFns?.[i];
  if (fns?.length) {
    const order = [...editor.querySelectorAll('sup.fn')].map((x) => x.dataset.fid).filter((v, k, a) => a.indexOf(v) === k);
    sh.append(h('div', { class: 'fnotes doc', style: { left: `${p.ml}px`, width: `${p.w - p.ml - p.mr}px`, bottom: `${p.mb}px` } }, ...fns.map((f) => h('div', { class: 'fnote' }, h('sup', {}, String(order.indexOf(f) + 1)), ` ${state.footnotes[f]}`))));
  }
  if (state.enPos && state.enPos.page === i) sh.append(F.endnotesEl({ left: `${p.ml}px`, width: `${p.w - p.ml - p.mr}px`, top: `${p.mt + state.enPos.y}px` }));
  if (pn && pn !== 'none') {
    const label = tf('Página {0} de {1}', i + 1, n);
    const style = pn.startsWith('header') ? { top: `${Math.max(12, p.mt / 2 - 10)}px` } : { bottom: `${Math.max(12, p.mb / 2 - 10)}px` };
    if (pn.endsWith('center')) Object.assign(style, { left: '0', right: '0', textAlign: 'center' });
    else Object.assign(style, { right: `${p.mr}px` });
    sh.append(h('div', { class: 'num', style }, label));
  }
}

function applyZoom() {
  const z = state.zoom / 100;
  pagesEl.style.transform = `scale(${z})`;
  pagesEl.style.marginBottom = `${(parseFloat(pagesEl.style.height) || 0) * (z - 1) + 60}px`;
  // centrado cuando la página es más angosta que la ventana
  const W = state.page.w * z;
  pagesEl.style.marginLeft = pagesEl.style.marginRight = W + 48 < scroller.clientWidth ? 'auto' : '24px';
  pagesEl.style.transformOrigin = W + 48 < scroller.clientWidth ? 'top center' : 'top left';
}

function drawRuler() {
  const z = state.zoom / 100, p = state.page;
  const W = p.w * z;
  const rect = pagesEl.getBoundingClientRect(), sr = scroller.getBoundingClientRect();
  const left = rect.left - sr.left + scroller.scrollLeft;
  ruler.replaceChildren();
  const track = h('div', { class: 'track', style: { left: `${left}px`, width: `${W}px` } },
    h('div', { class: 'm', style: { left: 0, width: `${p.ml * z}px` } }), h('div', { class: 'm', style: { right: 0, width: `${p.mr * z}px` } }));
  ruler.append(track);
  const cmW = CM * z;
  for (let c = 0; c * cmW <= W - p.ml * z; c += 0.5) {
    const x = left + p.ml * z + c * cmW;
    if (c % 1 === 0 && c > 0) ruler.append(h('div', { class: 'lab', style: { left: `${x}px` } }, String(c)));
    else ruler.append(h('div', { class: 'tick', style: { left: `${x}px`, height: c % 1 ? '4px' : '7px' } }));
  }
  F.drawIndentMarkers(left + p.ml * z, z, p.w - p.ml - p.mr);
  F.drawTabMarkers(left + p.ml * z, z, p.w - p.ml - p.mr);
}

// ------------------------------------------------------------------ barra de estado
function updateCounts() {
  clearTimeout(updateCounts.t);
  updateCounts.t = setTimeout(() => {
    const text = editor.innerText || '';
    const words = (text.match(/[\p{L}\p{N}]+(?:['’-][\p{L}\p{N}]+)*/gu) || []).length;
    statusWords.textContent = `${words.toLocaleString('es-MX')} palabras`;
    statusWords.dataset.chars = String(text.replace(/\s/g, '').length);
  }, 200);
}
function updateStatusPage() {
  if (!statusPages) return;
  const S = pageDims().S * (state.zoom / 100);
  const cur = Math.min(state.pages, Math.floor((scroller.scrollTop + scroller.clientHeight / 3) / S) + 1);
  statusPages.textContent = `Página ${cur} de ${state.pages}`;
}

// ------------------------------------------------------------------ estado de la cinta
function reflectState() {
  if (!shell) return;
  const r = shell.ribbon;
  const q = (c) => { try { return document.queryCommandState(c); } catch { return false; } };
  ['bold', 'italic', 'underline', 'strikeThrough', 'superscript', 'subscript', 'insertUnorderedList', 'insertOrderedList'].forEach((c) => r.setOn(c, q(c)));
  const b = currentBlock();
  const align = b ? getComputedStyle(b).textAlign : 'left';
  r.setOn('justifyLeft', align === 'left' || align === 'start');
  r.setOn('justifyCenter', align === 'center');
  r.setOn('justifyRight', align === 'right' || align === 'end');
  r.setOn('justifyFull', align === 'justify');
  r.setValue('fontSize', currentFontSizePt());
  const s = document.getSelection();
  let n = s.anchorNode; if (n && n.nodeType === 3) n = n.parentElement;
  if (n && editor.contains(n)) {
    const fam = getComputedStyle(n).fontFamily.split(',')[0].replace(/["']/g, '').trim();
    const known = FONTS.find((f) => f[1].split(',').map((x) => x.trim().replace(/["']/g, '')).includes(fam));
    r.setValue('fontName', known ? known[0] : fam);
  }
  let kind = 'p';
  if (b) {
    if (/^H[1-6]$/.test(b.tagName)) kind = b.tagName.toLowerCase();
    else if (b.tagName === 'BLOCKQUOTE') kind = 'blockquote';
    else if (b.classList.contains('title')) kind = 'title';
    else if (b.classList.contains('subtitle')) kind = 'subtitle';
  }
  const sid = b?.dataset?.style && state.styles?.[b.dataset.style] ? b.dataset.style : null;
  $$('.stylegal .gi').forEach((g) => g.classList.toggle('on', sid ? g.dataset.sid === sid : (g.dataset.style === kind && !g.dataset.sid)));
  const inTable = !!(b && b.closest('table'));
  r.setContextual('table', inTable);
  r.setContextual('picture', !!selImg);
}

// ------------------------------------------------------------------ insertar
function insertHTMLAtCursor(html) { restoreSelection(); document.execCommand('insertHTML', false, html); saveSelection(); shell.markDirty(); schedulePaginate(); }

function insertTable(rows, cols) {
  const head = `<tr>${'<th><br></th>'.repeat(cols)}</tr>`;
  const body = Array.from({ length: rows - 1 }, () => `<tr>${'<td><br></td>'.repeat(cols)}</tr>`).join('');
  insertHTMLAtCursor(`<table><tbody>${head}${body}</tbody></table><p><br></p>`);
}
function tablePicker(anchor) {
  const box = h('div', {});
  const lbl = h('div', { class: 'hdr' }, 'Insertar tabla');
  const grid = h('div', { class: 'gridpick' });
  const cells = [];
  for (let r = 0; r < 8; r++) for (let c = 0; c < 10; c++) {
    const s = h('span', { 'data-r': r, 'data-c': c });
    cells.push(s); grid.append(s);
  }
  grid.addEventListener('mouseover', (e) => {
    const t = e.target.closest('span'); if (!t) return;
    const R = +t.dataset.r, C = +t.dataset.c;
    cells.forEach((s) => s.classList.toggle('on', +s.dataset.r <= R && +s.dataset.c <= C));
    lbl.textContent = `Tabla de ${C + 1} × ${R + 1}`;
  });
  grid.addEventListener('click', (e) => { const t = e.target.closest('span'); if (!t) return; closeDropdown(); insertTable(+t.dataset.r + 1, +t.dataset.c + 1); });
  box.append(lbl, grid, h('button', { class: 'mi', onclick: async () => { closeDropdown(); const v = await form('Insertar tabla', [{ name: 'cols', label: 'Columnas', type: 'number', value: 3, min: 1, max: 30 }, { name: 'rows', label: 'Filas', type: 'number', value: 4, min: 1, max: 300 }], 'Insertar'); if (v) insertTable(Math.max(1, v.rows), Math.max(1, v.cols)); }, html: `${icon('table', 18)}<span>Más opciones…</span>` }));
  dropdown(anchor, box);
}
function tableOp(op, arg) {
  const b = currentBlock();
  const cell = b && b.closest('td,th');
  if (!cell) return;
  const tr = cell.parentElement, table = cell.closest('table');
  const ci = [...tr.children].indexOf(cell);
  const mkCell = (tag = 'td') => { const c = document.createElement(tag); c.innerHTML = '<br>'; return c; };
  if (op === 'rowAbove' || op === 'rowBelow') {
    const nr = document.createElement('tr');
    [...tr.children].forEach(() => nr.append(mkCell()));
    tr.parentElement.insertBefore(nr, op === 'rowAbove' ? tr : tr.nextSibling);
  } else if (op === 'colLeft' || op === 'colRight') {
    [...table.rows].forEach((row) => { const ref = row.children[ci]; const c = mkCell(ref?.tagName === 'TH' ? 'th' : 'td'); row.insertBefore(c, op === 'colLeft' ? ref : ref?.nextSibling || null); });
  } else if (op === 'delRow') { if (table.rows.length <= 1) table.remove(); else tr.remove(); }
  else if (op === 'delCol') { if (tr.children.length <= 1) table.remove(); else [...table.rows].forEach((row) => row.children[ci]?.remove()); }
  else if (op === 'delTable') table.remove();
  else if (op === 'shade') cell.style.backgroundColor = arg === 'transparent' ? '' : arg;
  else if (op === 'header') { const first = table.rows[0]; [...first.children].forEach((c) => { const n = document.createElement(c.tagName === 'TH' ? 'td' : 'th'); n.innerHTML = c.innerHTML; c.replaceWith(n); }); }
  shell.markDirty(); schedulePaginate();
}

function insertImageDataUrl(url) {
  const img = new Image();
  img.onload = () => {
    const max = pageDims().cw;
    const w = Math.min(img.naturalWidth, max);
    insertHTMLAtCursor(`<img src="${url}" width="${Math.round(w)}" alt="">`);
    setTimeout(schedulePaginate, 50);
  };
  img.src = url;
}
async function insertImage() {
  if (!api) return;
  const f = await api.openImage();
  if (!f) return;
  const b64 = btoa(Array.from(f.bytes, (c) => String.fromCharCode(c)).join(''));
  insertImageDataUrl(`data:${f.mime};base64,${b64}`);
}

async function insertLink() {
  saveSelection();
  const s = document.getSelection();
  const a = s.anchorNode && (s.anchorNode.nodeType === 3 ? s.anchorNode.parentElement : s.anchorNode).closest('a');
  const v = await form('Insertar vínculo', [
    { name: 'text', label: 'Texto a mostrar', value: a ? a.textContent : s.toString() },
    { name: 'url', label: 'Dirección (URL o correo)', value: a ? a.getAttribute('href') : 'https://', placeholder: 'https://…' },
  ], 'Aceptar');
  if (!v || !v.url) return;
  let url = v.url.trim();
  if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(url)) url = `mailto:${url}`;
  if (!/^(https?:|mailto:)/i.test(url)) url = `https://${url}`;
  if (a) { a.href = url; a.textContent = v.text || url; shell.markDirty(); return; }
  insertHTMLAtCursor(`<a href="${esc(url)}">${esc(v.text || url)}</a>`);
}

function insertPageBreak() {
  insertHTMLAtCursor('<div class="pb" contenteditable="false"></div><p><br></p>');
}

async function insertSymbol(anchor) {
  const syms = '© ® ™ ° ± × ÷ ≠ ≤ ≥ ≈ ∞ √ π Σ µ € $ £ ¥ ¢ § ¶ † • … – — « » ¿ ¡ ñ Ñ á é í ó ú ü Á É Í Ó Ú Ü ← → ↑ ↓ ✓ ✗ ★ ☆ ♥ ☎ ✉ ½ ¼ ¾ ² ³ ‰'.split(' ');
  const box = h('div', { class: 'swatches', style: { gridTemplateColumns: 'repeat(10, 26px)' } });
  syms.forEach((s) => box.append(h('button', { style: { width: '26px', height: '26px', background: '#fff', fontSize: '15px' }, onclick: () => { closeDropdown(); insertHTMLAtCursor(esc(s)); } }, s)));
  dropdown(anchor, box);
}

function buildTOC() {
  const heads = [...editor.querySelectorAll(':scope > h1, :scope > h2, :scope > h3')];
  const lines = heads.map((hd) => {
    const idx = [...editor.children].indexOf(hd);
    return `<div class="toc-l l${hd.tagName[1]}"><span>${esc(hd.textContent)}</span><span></span><span>${(state.pageOf[idx] ?? 0) + 1}</span></div>`;
  }).join('') || '<div class="toc-l"><span>Usa los estilos Título 1, 2 y 3 para generar el índice.</span></div>';
  return `<div class="toc" contenteditable="false"><div class="toc-h">Contenido</div>${lines}</div>`;
}
function insertTOC() {
  const existing = editor.querySelector('.toc');
  if (existing) { existing.outerHTML = buildTOC(); schedulePaginate(); setTimeout(() => { const t = editor.querySelector('.toc'); if (t) t.outerHTML = buildTOC(); }, 60); toast('Índice actualizado'); shell.markDirty(); return; }
  insertHTMLAtCursor(`${buildTOC()}<p><br></p>`);
  setTimeout(() => { const t = editor.querySelector('.toc'); if (t) t.outerHTML = buildTOC(); }, 80);
}

// ------------------------------------------------------------------ imágenes: seleccionar y cambiar tamaño
function selectImage(img) {
  $$('img.sel', editor).forEach((i) => i.classList.remove('sel'));
  selImg = img;
  if (img) img.classList.add('sel');
  imgBox.hidden = !img;
  positionImgBox();
  reflectState();
}
function positionImgBox() {
  if (!selImg || !selImg.isConnected) { imgBox.hidden = true; return; }
  imgBox.style.left = `${selImg.offsetLeft + editor.offsetLeft}px`;
  imgBox.style.top = `${selImg.offsetTop + editor.offsetTop}px`;
  imgBox.style.width = `${selImg.offsetWidth}px`;
  imgBox.style.height = `${selImg.offsetHeight}px`;
}
function startImgResize(e) {
  const hd = e.target.dataset.h; if (!hd || !selImg) return;
  e.preventDefault();
  const z = state.zoom / 100, x0 = e.clientX, w0 = selImg.offsetWidth, max = pageDims().cw;
  const move = (ev) => {
    const dx = (ev.clientX - x0) / z * (hd.includes('w') ? -1 : 1);
    selImg.width = Math.round(Math.max(24, Math.min(max, w0 + dx)));
    selImg.removeAttribute('height');
    positionImgBox();
  };
  const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); shell.markDirty(); schedulePaginate(); };
  addEventListener('mousemove', move); addEventListener('mouseup', up);
}
function imageOp(op) {
  if (!selImg) return;
  const b = selImg.closest('p,div,li,td') || selImg.parentElement;
  if (op === 'left' || op === 'center' || op === 'right') b.style.textAlign = op;
  if (op === 'size') {
    form('Tamaño de la imagen', [{ name: 'w', label: 'Ancho (px)', type: 'number', value: selImg.width, min: 16, max: Math.round(pageDims().cw) }]).then((v) => { if (v) { selImg.width = v.w; selImg.removeAttribute('height'); shell.markDirty(); schedulePaginate(); } });
    return;
  }
  if (op === 'reset') { selImg.width = Math.min(selImg.naturalWidth, pageDims().cw); }
  if (op === 'delete') { selImg.remove(); selectImage(null); }
  shell.markDirty(); schedulePaginate();
}

// ------------------------------------------------------------------ buscar y reemplazar
const norm = (s) => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
let findState = { ranges: [], idx: -1 };
function findAll(query, { matchCase = false, accents = false } = {}) {
  const ranges = [];
  if (!query) return ranges;
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
  const nodes = []; let full = '';
  while (walker.nextNode()) { nodes.push({ n: walker.currentNode, start: full.length }); full += walker.currentNode.nodeValue; }
  const prep = (s) => { let x = s; if (!accents) x = x.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); if (!matchCase) x = x.toLowerCase(); return x; };
  // la normalización NFD puede cambiar longitudes: se trabaja carácter por carácter
  const mapIdx = []; let hay = '';
  for (let i = 0; i < full.length; i++) { const p = prep(full[i]); for (let k = 0; k < p.length; k++) mapIdx.push(i); hay += p; }
  const needle = prep(query);
  let from = 0;
  while (needle && (from = hay.indexOf(needle, from)) !== -1) {
    const s = mapIdx[from], e = mapIdx[from + needle.length - 1] + 1;
    const locate = (pos) => { let lo = 0; for (let j = 0; j < nodes.length; j++) if (nodes[j].start <= pos) lo = j; return lo; };
    const a = locate(s), b = locate(e - 1);
    const r = document.createRange();
    r.setStart(nodes[a].n, s - nodes[a].start); r.setEnd(nodes[b].n, e - nodes[b].start);
    ranges.push(r);
    from += needle.length;
  }
  return ranges;
}
function paintFind() {
  if (!window.CSS?.highlights) return;
  CSS.highlights.set('find', new Highlight(...findState.ranges));
  const cur = findState.ranges[findState.idx];
  if (cur) { CSS.highlights.set('find-cur', new Highlight(cur)); scrollRangeIntoView(cur); } else CSS.highlights.delete('find-cur');
}
function scrollRangeIntoView(r) {
  const rect = r.getBoundingClientRect(), sr = scroller.getBoundingClientRect();
  if (rect.top < sr.top + 40 || rect.bottom > sr.bottom - 40) scroller.scrollTop += rect.top - sr.top - sr.height / 3;
}
let findbar = null;
function openFind(replace = false, initial = '') {
  if (!findbar) {
    const q = h('input', { placeholder: 'Buscar', 'aria-label': 'Buscar' });
    const rep = h('input', { placeholder: 'Reemplazar con', 'aria-label': 'Reemplazar con' });
    const count = h('span', { class: 'count' });
    const mc = h('input', { type: 'checkbox' }), acc = h('input', { type: 'checkbox' });
    const run = () => {
      findState.ranges = findAll(q.value, { matchCase: mc.checked, accents: acc.checked });
      findState.idx = findState.ranges.length ? 0 : -1;
      count.textContent = q.value ? (findState.ranges.length ? `1 de ${findState.ranges.length}` : 'Sin resultados') : '';
      paintFind();
    };
    const step = (d) => {
      if (!findState.ranges.length) return;
      findState.idx = (findState.idx + d + findState.ranges.length) % findState.ranges.length;
      count.textContent = `${findState.idx + 1} de ${findState.ranges.length}`;
      paintFind();
    };
    const replaceOne = () => {
      const r = findState.ranges[findState.idx]; if (!r) return;
      const s = document.getSelection(); s.removeAllRanges(); s.addRange(r);
      document.execCommand('insertText', false, rep.value);
      shell.markDirty(); run();
    };
    const replaceAll = () => {
      const n = findState.ranges.length; if (!n) return;
      for (let i = n - 1; i >= 0; i--) { const s = document.getSelection(); s.removeAllRanges(); s.addRange(findState.ranges[i]); document.execCommand('insertText', false, rep.value); }
      shell.markDirty(); run(); toast(`${n} reemplazos`);
    };
    q.addEventListener('input', run);
    q.addEventListener('keydown', (e) => { if (e.key === 'Enter') step(e.shiftKey ? -1 : 1); if (e.key === 'Escape') closeFind(); });
    rep.addEventListener('keydown', (e) => { if (e.key === 'Enter') replaceOne(); if (e.key === 'Escape') closeFind(); });
    mc.addEventListener('change', run); acc.addEventListener('change', run);
    const repRow = h('div', { class: 'row' }, rep, h('button', { class: 'btn', onclick: replaceOne }, 'Reemplazar'), h('button', { class: 'btn', onclick: replaceAll }, 'Todos'));
    findbar = h('div', { class: 'findbar' },
      h('div', { class: 'row' }, q, count, h('button', { class: 'rb sm', title: 'Anterior', html: icon('chevronDown', 16).replace('M6 9l6 6 6-6', 'M6 15l6-6 6 6'), onclick: () => step(-1) }), h('button', { class: 'rb sm', title: 'Siguiente', html: icon('chevronDown', 16), onclick: () => step(1) }), h('button', { class: 'rb sm', title: 'Cerrar', html: icon('close', 16), onclick: () => closeFind() })),
      repRow,
      h('div', { class: 'row', style: { gap: '14px', fontSize: '12px' } }, h('label', { class: 'check' }, mc, 'Mayúsculas'), h('label', { class: 'check' }, acc, 'Distinguir acentos')));
    findbar.q = q; findbar.rep = repRow; findbar.run = run;
    shell.workspace.append(findbar);
  }
  findbar.hidden = false;
  findbar.rep.hidden = !replace;
  const sel = document.getSelection().toString();
  if (initial || (sel && sel.length < 80)) findbar.q.value = initial || sel;
  findbar.q.focus(); findbar.q.select();
  findbar.run();
}
function closeFind() { if (findbar) findbar.hidden = true; CSS.highlights?.clear(); editor.focus(); }

// ------------------------------------------------------------------ página, encabezado y pie
async function pageSetup() {
  const p = state.page;
  const v = await form('Configurar página', [
    { name: 'size', label: 'Tamaño del papel', type: 'select', value: p.size, options: Object.entries(PAGE_SIZES).map(([k, s]) => [k, s.name]) },
    { name: 'orient', label: 'Orientación', type: 'select', value: p.landscape ? 'h' : 'v', options: [['v', 'Vertical'], ['h', 'Horizontal']] },
    { name: 'mt', label: 'Margen superior (cm)', type: 'number', step: 0.1, value: (p.mt / CM).toFixed(1) },
    { name: 'mb', label: 'Margen inferior (cm)', type: 'number', step: 0.1, value: (p.mb / CM).toFixed(1) },
    { name: 'ml', label: 'Margen izquierdo (cm)', type: 'number', step: 0.1, value: (p.ml / CM).toFixed(1) },
    { name: 'mr', label: 'Margen derecho (cm)', type: 'number', step: 0.1, value: (p.mr / CM).toFixed(1) },
  ], 'Aplicar');
  if (!v) return;
  setPage({ size: v.size, landscape: v.orient === 'h', mt: v.mt * CM, mb: v.mb * CM, ml: v.ml * CM, mr: v.mr * CM });
}
function setPage(o) {
  const p = state.page;
  Object.assign(p, o);
  const s = PAGE_SIZES[p.size] || { w: p.w, h: p.h };
  if (o.size || o.landscape !== undefined) { p.w = p.landscape ? Math.max(s.w, s.h) : Math.min(s.w, s.h); p.h = p.landscape ? Math.min(s.w, s.h) : Math.max(s.w, s.h); }
  shell.markDirty(); schedulePaginate();
}
const MARGIN_PRESETS = [
  ['Normal (2.5 / 3 cm)', { mt: 2.5, mb: 2.5, ml: 3, mr: 3 }], ['Estrecho (1.27 cm)', { mt: 1.27, mb: 1.27, ml: 1.27, mr: 1.27 }],
  ['Moderado (2.54 / 1.91 cm)', { mt: 2.54, mb: 2.54, ml: 1.91, mr: 1.91 }], ['Ancho (2.54 / 5.08 cm)', { mt: 2.54, mb: 2.54, ml: 5.08, mr: 5.08 }],
];
async function headerFooter() {
  const v = await form('Encabezado y pie de página', [
    { name: 'header', label: 'Encabezado (se repite en cada página)', value: state.header, placeholder: 'Ej.: Nombre de la empresa' },
    { name: 'footer', label: 'Pie de página', value: state.footer, placeholder: 'Ej.: Documento confidencial' },
    { name: 'pageNumbers', label: 'Números de página', type: 'select', value: state.pageNumbers, options: [['none', 'Sin números'], ['footer-center', 'Abajo, al centro'], ['footer-right', 'Abajo, a la derecha'], ['header-right', 'Arriba, a la derecha']] },
    { type: 'note', label: 'Puedes escribir {pagina} y {paginas} dentro del texto.' },
  ], 'Aplicar');
  if (!v) return;
  Object.assign(state, v);
  shell.markDirty(); schedulePaginate();
}

async function watermarkDialog() {
  const v = await form('Marca de agua', [
    { name: 'text', label: 'Texto (déjalo vacío para quitarla)', value: state.watermark?.text || '', placeholder: 'Ej.: CONFIDENCIAL, BORRADOR, MUESTRA' },
    { name: 'color', label: 'Color', type: 'select', value: state.watermark?.color || '#C9CBD6', options: [['#C9CBD6', 'Gris claro'], ['#F4B6C2', 'Rosa'], ['#B9B2E3', 'Lila'], ['#A8DCC8', 'Verde']] },
    { type: 'note', label: 'La marca de agua aparece en pantalla, al imprimir y en el PDF.' },
  ], 'Aplicar');
  if (!v) return;
  state.watermark = v.text.trim() ? { text: v.text.trim(), color: v.color } : null;
  shell.markDirty(); schedulePaginate();
}

// ------------------------------------------------------------------ impresión / PDF (páginas idénticas a la pantalla)
let pageStyleEl = null;
function buildPrint() {
  paginate();
  const p = state.page, root = $('#printRoot');
  root.replaceChildren();
  const n = state.pages;
  const pages = Array.from({ length: n }, (_, i) => {
    const pg = h('div', { class: 'ppage', style: { width: `${p.w}px`, height: `${p.h}px` } });
    const sheet = h('div', { class: 'sheet', style: { position: 'absolute', inset: '0', boxShadow: 'none' } });
    renderHeaderFooter(sheet, i, n);
    const content = h('div', { class: 'pcontent doc', style: { left: `${p.ml}px`, top: `${p.mt}px`, width: `${p.w - p.ml - p.mr}px` } });
    pg.append(sheet, content);
    root.append(pg);
    return content;
  });
  [...editor.children].forEach((b, i) => {
    const c = b.cloneNode(true);
    c.style.marginTop = '';
    delete c.dataset.push;
    c.querySelectorAll('img.sel').forEach((x) => x.classList.remove('sel'));
    const target = pages[Math.min(n - 1, state.pageOf[i] ?? 0)];
    target.append(c);
  });
  pageStyleEl?.remove();
  pageStyleEl = h('style', {}, `@page { size: ${p.w / 96}in ${p.h / 96}in; margin: 0; }`);
  document.head.append(pageStyleEl);
}

// ------------------------------------------------------------------ adaptador del archivo
function snapshotSettings() {
  return { page: { ...state.page }, header: state.header, footer: state.footer, pageNumbers: state.pageNumbers, comments: state.comments, footnotes: state.footnotes, endnotes: state.endnotes, sources: state.sources, track: state.track, watermark: state.watermark, pageColor: state.pageColor, defTab: state.defTab || null, styles: state.styles || {}, chapnum: state.chapnum || null };
}
function setContent(html, settings) {
  editor.innerHTML = html || '<p><br></p>';
  if (settings) {
    if (settings.page) Object.assign(state.page, settings.page);
    ['header', 'footer', 'pageNumbers', 'comments', 'footnotes', 'endnotes', 'sources', 'track', 'watermark', 'pageColor', 'defTab', 'styles', 'chapnum'].forEach((k) => { if (settings[k] !== undefined) state[k] = settings[k]; });
    editor.classList.toggle('tracking', !!state.track);
  }
  ES.render();
  selectImage(null);
  editor.querySelectorAll('img').forEach((img) => img.addEventListener('load', schedulePaginate, { once: true }));
  schedulePaginate(); updateCounts();
  scroller.scrollTop = 0;
}

const adapter = {
  i18nSkip: '#pages, .doc, .fnmeasure, .cm-text, .cmcard .reply-text, .nav-item',
  exts: ['docx', 'odt', 'doc', 'txt', 'html', 'htm', 'md'],
  writable: ['docx', 'odt', 'html', 'htm', 'txt', 'md'],
  saveExt: 'docx',
  docIcon: 'doc',
  homeHint: 'Escribe cartas, reportes y documentos que se abren en cualquier procesador de textos.',
  templates: TEMPLATES.map((t) => ({ ...t, thumb: thumb(t) })),
  exportFormats: [
    { ext: 'docx', ic: 'doc', label: 'Documento de Word (.docx)', desc: 'Para editar en cualquier procesador de textos compatible.' },
    { ext: 'odt', ic: 'doc', label: 'Documento ODF (.odt)', desc: 'Formato abierto, compatible con LibreOffice y otros procesadores.' },
    { ext: 'html', ic: 'link', label: 'Página web (.html)', desc: 'Para publicar en internet o abrir en el navegador.' },
    { ext: 'txt', ic: 'words', label: 'Texto sin formato (.txt)', desc: 'Solo el texto, sin estilos ni imágenes.' },
    { ext: 'md', ic: 'words', label: 'Markdown (.md)', desc: 'Texto con títulos, listas, negritas y tablas, para GitHub, blogs y notas.' },
  ],
  async newDoc(tplId) {
    const t = TEMPLATES.find((x) => x.id === tplId) || TEMPLATES[0];
    state.page = { size: 'carta', w: 816, h: 1056, mt: 2.5 * CM, mb: 2.5 * CM, ml: 3 * CM, mr: 3 * CM, landscape: false, gap: 24 };
    state.header = ''; state.footer = ''; state.defTab = null; state.styles = {}; state.chapnum = null; state.pageNumbers = tplId === 'reporte' ? 'footer-center' : 'none';
    Object.assign(state, { comments: [], footnotes: {}, endnotes: {}, sources: [], track: false, watermark: null, pageColor: null, merge: null });
    editor.classList.remove('tracking');
    setContent(localizeHtml(t.html()));
    editor.focus();
    const first = editor.querySelector('p,h1');
    if (first && t.blank) { const r = document.createRange(); r.setStart(first, 0); r.collapse(true); document.getSelection().removeAllRanges(); document.getSelection().addRange(r); }
  },
  async load(f) {
    state.styles = {}; state.chapnum = null; // estilos del documento anterior
    const text = () => new TextDecoder('utf-8').decode(f.bytes);
    if (f.ext === 'docx') {
      const r = await docxToHtml(f.bytes);
      state.header = r.header || ''; state.footer = r.footer || ''; state.pageNumbers = 'none'; state.defTab = r.defTab || null; state.styles = r.styles || {}; state.chapnum = r.chapnum || null;
      Object.assign(state, { comments: r.comments || [], footnotes: r.footnotes || {}, endnotes: r.endnotes || {}, sources: r.sources || [], track: false, watermark: null, pageColor: null, merge: null });
      setContent(r.html, r.page ? { page: { ...state.page, ...r.page, size: matchSize(r.page) } } : null);
      return { warnings: r.warnings, props: await readCoreProps(f.bytes) };
    }
    if (f.ext === 'odt') {
      const r = await odtToHtml(f.bytes);
      Object.assign(state, { styles: r.styles || {}, chapnum: r.chapnum || null, defTab: r.defTab || null, header: r.header || '', footer: r.footer || '', pageNumbers: 'none', comments: r.comments || [], footnotes: r.footnotes || {}, endnotes: r.endnotes || {}, sources: r.sources || [], track: false, watermark: null, pageColor: null, merge: null });
      setContent(r.html, r.page ? { page: { ...state.page, ...r.page, size: matchSize(r.page) } } : null);
      return { props: r.props };
    }
    if (f.ext === 'doc') {
      // Documento binario antiguo de Word: se recupera el texto (modo compatibilidad)
      const d = f.doc || {};
      Object.assign(state, { header: (d.headers || '').trim().split('\n')[0] || '', footer: (d.footers || '').trim().split('\n')[0] || '', pageNumbers: 'none', comments: [], footnotes: {}, endnotes: {}, sources: [], track: false, watermark: null, pageColor: null, merge: null });
      const paras = String(d.body || '').replace(/\r\n?/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
      let html = d.html ? sanitize(d.html) : textToHtml(paras);
      const notes = [d.footnotes, d.endnotes].map((x) => String(x || '').trim()).filter(Boolean);
      if (notes.length) html += `<h2>${esc(t('Notas'))}</h2>${textToHtml(notes.join('\n'))}`;
      setContent(html);
      return { warnings: [d.html ? t('Documento de Word 97-2003 (.doc): se recuperaron el texto, los títulos, las tablas, las listas y el formato básico. Se guardará como .docx (modo compatibilidad)') : t('Documento de Word 97-2003 (.doc): se recuperó el texto; el formato se pierde (modo compatibilidad)')] };
    }
    if (f.ext === 'txt') { setContent(textToHtml(text())); return {}; }
    if (f.ext === 'md') { setContent(markdownToHtml(text())); return {}; }
    if (f.ext === 'html' || f.ext === 'htm') {
      const raw = text();
      const meta = /<script type="application\/json" id="escribiente">([\s\S]*?)<\/script>/.exec(raw);
      const body = /<body[^>]*>([\s\S]*)<\/body>/i.exec(raw);
      setContent(sanitize(body ? body[1] : raw), meta ? JSON.parse(meta[1]) : null);
      return {};
    }
    throw new Error('Formato no compatible');
  },
  async serialize(ext) {
    if (ext === 'txt') return htmlToText(editor);
    if (ext === 'md') return F3.htmlToMarkdown(editor);
    if (ext === 'html' || ext === 'htm') return exportHtml();
    if (ext === 'odt') return htmlToOdt(editor, { ...snapshotSettings(), ...shell.docProps() });
    return htmlToDocx(editor, { ...snapshotSettings(), ...shell.docProps() });
  },
  snapshot() { return JSON.stringify({ html: editor.innerHTML, settings: snapshotSettings() }); },
  async restore(s) { const j = JSON.parse(s); setContent(j.html, j.settings); },
  isPristine() { return !shell.doc.dirty && (editor.innerText || '').trim() === '' && !editor.querySelector('img,table'); },
  pdfOptions() { return { landscape: false, pageSize: { width: state.page.w / 96, height: state.page.h / 96 } }; },
  beforePrint() { buildPrint(); },
  afterPrint() { $('#printRoot').replaceChildren(); },
  focus() { editor.focus({ preventScroll: true }); },
  stats() {
    const text = editor.innerText || '';
    return [['Páginas', state.pages], ['Palabras', (text.match(/[\p{L}\p{N}]+/gu) || []).length], ['Caracteres (sin espacios)', text.replace(/\s/g, '').length], ['Párrafos', editor.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li,blockquote').length]];
  },
  onDropFile(file) {
    if (file.type.startsWith('image/')) { const rd = new FileReader(); rd.onload = () => insertImageDataUrl(rd.result); rd.readAsDataURL(file); return true; }
    return false;
  },
  onCommand,
};

function matchSize(pg) {
  const w = Math.min(pg.w, pg.h), hh = Math.max(pg.w, pg.h);
  return Object.entries(PAGE_SIZES).find(([, s]) => Math.abs(s.w - w) < 6 && Math.abs(s.h - hh) < 6)?.[0] || 'custom';
}

function exportHtml() {
  const clone = editor.cloneNode(true);
  clone.querySelectorAll('[data-push]').forEach((b) => { b.style.marginTop = ''; b.removeAttribute('data-push'); });
  clone.querySelectorAll('.sel').forEach((b) => b.classList.remove('sel'));
  const css = [...document.styleSheets].flatMap((ss) => { try { return [...ss.cssRules].map((r) => r.cssText).filter((t) => t.startsWith('.doc')); } catch { return []; } }).join('\n');
  const title = esc(shell.doc.name.replace(/\.[^.]+$/, ''));
  return `<!doctype html>\n<html lang="es"><head><meta charset="utf-8"><title>${title}</title><meta name="generator" content="Escribiente · Mi Oficina MX">\n<style>body{background:#eee;margin:0;padding:24px}.doc{position:static!important;background:#fff;max-width:${state.page.w - state.page.ml - state.page.mr}px;margin:0 auto;padding:${state.page.mt}px ${state.page.mr}px ${state.page.mb}px ${state.page.ml}px;box-shadow:0 2px 10px rgba(0,0,0,.1)}\n${css}</style>\n<script type="application/json" id="escribiente">${JSON.stringify(snapshotSettings()).replace(/</g, '\\u003c')}</script></head>\n<body><div class="doc">${clone.innerHTML}</div></body></html>`;
}

// ------------------------------------------------------------------ comandos
async function onCommand(cmd, arg) {
  switch (cmd) {
    case 'undo': return exec('undo');
    case 'redo': return exec('redo');
    case 'selectAll': editor.focus(); document.getSelection().selectAllChildren(editor); return;
    case 'bold': case 'italic': case 'underline': case 'strikeThrough': case 'superscript': case 'subscript':
    case 'justifyLeft': case 'justifyCenter': case 'justifyRight': case 'justifyFull':
    case 'insertUnorderedList': case 'insertOrderedList':
      return exec(cmd);
    case 'removeFormat': exec('removeFormat'); exec('unlink'); return;
    case 'fontDialog': return FX.fontDialog();
    case 'paraDialog': return FX.paragraphDialog();
    case 'paraBorders': return FX.paragraphDialog('borders');
    case 'tabs': return F.tabsDialog();
    case 'stylesPane': return F.toggleSide('styles');
    case 'newStyle': return ES.newFromSelection();
    case 'chapters': return;
    case 'changeCase': return;
    case 'marks': { const on = FX.toggleMarks(); shell.ribbon.setOn('marks', on); return; }
    case 'fontName': return applyFontFamily(arg);
    case 'fontSize': return applySpanStyle('fontSize', `${parseFloat(arg) || 11}pt`);
    case 'grow': return adjustSize(1);
    case 'shrink': return adjustSize(-1);
    case 'foreColor': restoreSelection(); document.execCommand('styleWithCSS', false, true); return exec('foreColor', arg === 'transparent' ? '#000000' : arg);
    case 'hiliteColor': restoreSelection(); document.execCommand('styleWithCSS', false, true); return exec('hiliteColor', arg);
    case 'indent': return setIndent(1);
    case 'outdent': return setIndent(-1);
    case 'lineSpacing': return; // menú
    case 'style': return setBlockStyle(arg);
    case 'insertTable': return tablePicker(arg && arg.nodeType ? arg : $('[data-cmd="insertTable"]') || shell.ribbonEl);
    case 'insertImage': return insertImage();
    case 'insertLink': return insertLink();
    case 'pageBreak': return insertPageBreak();
    case 'hr': return insertHTMLAtCursor('<hr><p><br></p>');
    case 'symbol': return insertSymbol(arg);
    case 'toc': return insertTOC();
    case 'headerFooter': return headerFooter();
    case 'pageSetup': return pageSetup();
    case 'orientation': return setPage({ landscape: !state.page.landscape });
    case 'find': return openFind(false);
    case 'findText': return openFind(false, arg);
    case 'replace': return openFind(true);
    case 'zoomIn': return zoom.set(state.zoom + 10);
    case 'zoomOut': return zoom.set(state.zoom - 10);
    case 'zoomReset': return zoom.set(100);
    case 'zoomWidth': return zoom.set(Math.floor(((scroller.clientWidth - 60) / state.page.w) * 100));
    case 'zoomPage': return zoom.set(Math.floor(((scroller.clientHeight - 60) / state.page.h) * 100));
    case 'wordCount': {
      const st = adapter.stats();
      return modal({ title: 'Contar palabras', body: h('table', {}, ...st.map(([k, v]) => h('tr', {}, h('td', { style: { padding: '4px 30px 4px 0', color: 'var(--ink-2)' } }, k), h('td', { style: { textAlign: 'right' } }, Number(v).toLocaleString('es-MX'))))) });
    }
    case 'spell': {
      const on = editor.spellcheck = !editor.spellcheck;
      shell.ribbon.setOn('spell', on);
      editor.innerHTML = editor.innerHTML; // fuerza a repintar los subrayados
      schedulePaginate();
      return toast(on ? 'Revisión ortográfica activada' : 'Revisión ortográfica desactivada');
    }
    case 'rowAbove': case 'rowBelow': case 'colLeft': case 'colRight': case 'delRow': case 'delCol': case 'delTable': case 'header':
      return tableOp(cmd);
    case 'cellShade': return tableOp('shade', arg);
    case 'imgLeft': return imageOp('left');
    case 'imgCenter': return imageOp('center');
    case 'imgRight': return imageOp('right');
    case 'imgSize': return imageOp('size');
    case 'imgReset': return imageOp('reset');
    case 'imgDelete': return imageOp('delete');
    case 'exportDocx': return shell.saveAsFormat('docx', 'Documento de Word');
    case 'painter': return F.startPainter(false);
    case 'navPane': return F.toggleNav();
    case 'comment': return F.newComment();
    case 'showComments': return F.toggleSide('comments');
    case 'showChanges': return F.toggleSide('changes');
    case 'track': return F.setTrack(!state.track);
    case 'revPrev': return F.reviewAction('prev');
    case 'revNext': return F.reviewAction('next');
    case 'revAccept': return F.reviewAction('accept');
    case 'revReject': return F.reviewAction('reject');
    case 'revAcceptAll': return F.reviewAction('acceptAll');
    case 'revRejectAll': return F.reviewAction('rejectAll');
    case 'footnote': return F.insertFootnote();
    case 'endnote': return F.insertEndnote();
    case 'equation': return F3.insertEquation();
    case 'smartart': return F3.insertSmartArt();
    case 'sources': return F3.manageSources();
    case 'bibliography': return F3.updateBibliography();
    case 'compare': return F3.compareDocuments();
    case 'labels': return F3.labels();
    case 'envelope': return F3.envelope();
    case 'focusMode': return F3.focusMode();
    case 'readingMode': return F3.readingMode();
    case 'exportMd': return shell.saveAsFormat('md', 'Markdown');
    case 'textbox': return F.insertTextBox();
    case 'cols1': return F.setColumns(1);
    case 'cols2': return F.setColumns(2);
    case 'cols3': return F.setColumns(3);
    case 'recipients': return F.chooseRecipients();
    case 'mergePreview': return F.previewMerge(0);
    case 'mergePrev': return F.previewMerge(-1);
    case 'mergeNext': return F.previewMerge(1);
    case 'mergeFinish': return F.finishMerge();
    case 'watermark': return watermarkDialog();
    case 'pageColor': state.pageColor = arg === 'transparent' || arg === '#FFFFFF' ? null : arg; shell.markDirty(); schedulePaginate(); return;
    case 'exportOdt': return shell.saveAsFormat('odt', 'Documento ODF');
    default: return undefined;
  }
}

// ------------------------------------------------------------------ cinta
const GAL_ID = { p: 'Normal', title: 'Title', subtitle: 'Subtitle', h1: 'Heading1', h2: 'Heading2', h3: 'Heading3', h4: 'Heading4', blockquote: 'Quote' };
let galleryEl = null;
function fillGallery() {
  const g = galleryEl; if (!g) return;
  const prev = {
    p: 'font-size:13px', title: 'font-size:19px;font-weight:300', subtitle: 'font-size:13px;color:#6B6F80', h1: 'font-size:15px;font-weight:600;color:#2F2496',
    h2: 'font-size:14px;font-weight:600;color:#3D30B5', h3: 'font-size:13px;font-weight:600;color:#3D30B5', h4: 'font-size:12.5px;font-weight:600;font-style:italic;color:#3D30B5', blockquote: 'font-size:12.5px;font-style:italic;color:#555',
  };
  const items = BLOCK_STYLES.map(([k, name]) => ({ k, name, sid: null, css: prev[k] }));
  // estilos propios del documento después de los predeterminados
  for (const st of Object.values(state.styles || {})) if (st.kind === 'p' && !Object.values(GAL_ID).includes(st.id)) { const c = ES.effectiveCss(st.id); items.push({ k: 'p', name: st.name, sid: st.id, css: ['font-family', 'font-weight', 'font-style', 'color'].filter((x) => c[x]).map((x) => `${x}:${c[x]}`).join(';') + ';font-size:13px' }); }
  g.replaceChildren(...items.map((it) => {
    const b = h('button', { class: 'gi', 'data-style': it.k, ...(it.sid ? { 'data-sid': it.sid } : {}), title: it.name, html: `<b style="${it.css}">AaBbCc</b><small>${esc(it.name)}</small>` });
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => ES.applyPara(it.sid || GAL_ID[it.k]));
    b.addEventListener('contextmenu', (e) => { e.preventDefault(); dropdown(b, [{ label: 'Aplicar', run: () => ES.applyPara(it.sid || GAL_ID[it.k]) }, { label: 'Modificar…', run: () => ES.modifyStyle(it.sid || GAL_ID[it.k]) }, { label: 'Actualizar con la selección', run: () => ES.updateFromSelection(it.sid || GAL_ID[it.k]) }]); });
    return b;
  }));
}
function styleGallery() {
  galleryEl = h('div', { class: 'gallery stylegal' });
  fillGallery();
  return galleryEl;
}
const lineMenu = () => ['1.0', '1.15', '1.5', '2.0', '2.5', '3.0'].map((v) => ({ label: v, run: () => setLineHeight(v) }));
const marginMenu = () => MARGIN_PRESETS.map(([l, m]) => ({ label: l, run: () => setPage({ mt: m.mt * CM, mb: m.mb * CM, ml: m.ml * CM, mr: m.mr * CM }) })).concat(['-', { label: 'Márgenes personalizados…', run: pageSetup }]);
const sizeMenu = () => Object.entries(PAGE_SIZES).map(([k, s]) => ({ label: s.name, run: () => setPage({ size: k }) }));

const TABS = [
  {
    id: 'home', label: 'Inicio', groups: [
      { label: 'Portapapeles', items: [
        { cmd: 'paste', label: 'Pegar', icon: 'paste', size: 'lg', title: 'Pegar (Ctrl+V)' },
        { stack: [{ cmd: 'cut', label: 'Cortar', icon: 'cut', title: 'Cortar (Ctrl+X)' }, { cmd: 'copy', label: 'Copiar', icon: 'copy', title: 'Copiar (Ctrl+C)' }, { cmd: 'painter', label: 'Copiar formato', icon: 'brush', title: 'Copiar formato (doble clic: varias veces)' }] },
      ] },
      { label: 'Fuente', items: [
        { stack: [
          { row: [{ cmd: 'fontName', label: 'Fuente', select: { options: FONTS.map((f) => f[0]), width: '150px' } }, { cmd: 'fontSize', label: 'Tamaño', select: { options: SIZES.map(String), width: '62px' } }, { cmd: 'grow', label: 'Aumentar tamaño', icon: 'grow', kbd: 'Ctrl+]' }, { cmd: 'shrink', label: 'Reducir tamaño', icon: 'shrink', kbd: 'Ctrl+[' }] },
          { row: [
            { cmd: 'bold', label: 'Negrita', icon: 'bold', kbd: 'Ctrl+B' }, { cmd: 'italic', label: 'Cursiva', icon: 'italic', kbd: 'Ctrl+I' }, { cmd: 'underline', label: 'Subrayado', icon: 'underline', kbd: 'Ctrl+U' },
            { cmd: 'strikeThrough', label: 'Tachado', icon: 'strike' }, { cmd: 'subscript', label: 'Subíndice', icon: 'sub' }, { cmd: 'superscript', label: 'Superíndice', icon: 'sup' },
            { cmd: 'hiliteColor', label: 'Color de resaltado', icon: 'highlight', color: true, colorValue: '#FFF59D', colorOpts: { noneLabel: 'Sin resaltado' } },
            { cmd: 'foreColor', label: 'Color de fuente', icon: 'textColor', color: true, colorValue: '#C62828', colorOpts: { allowNone: false } },
            { cmd: 'changeCase', label: 'Cambiar mayúsculas y minúsculas', icon: 'changeCase', menu: () => FX.caseMenu() },
            { cmd: 'removeFormat', label: 'Borrar formato', icon: 'clear' },
            { cmd: 'fontDialog', label: 'Fuente… (Ctrl+D)', icon: 'dialogLauncher' },
          ] },
        ] },
      ] },
      { label: 'Párrafo', items: [
        { stack: [
          { row: [{ cmd: 'insertUnorderedList', label: 'Viñetas', icon: 'bullets', split: () => ES.bulletMenu() }, { cmd: 'insertOrderedList', label: 'Numeración', icon: 'numbers', split: () => ES.numberingMenu() }, { cmd: 'outdent', label: 'Disminuir sangría', icon: 'outdent' }, { cmd: 'indent', label: 'Aumentar sangría', icon: 'indent' }] },
          { row: [{ cmd: 'justifyLeft', label: 'Alinear a la izquierda', icon: 'alignLeft', kbd: 'Ctrl+L' }, { cmd: 'justifyCenter', label: 'Centrar', icon: 'alignCenter', kbd: 'Ctrl+E' }, { cmd: 'justifyRight', label: 'Alinear a la derecha', icon: 'alignRight', kbd: 'Ctrl+R' }, { cmd: 'justifyFull', label: 'Justificar', icon: 'justify', kbd: 'Ctrl+J' }, { cmd: 'lineSpacing', label: 'Interlineado', icon: 'lineSpacing', menu: lineMenu }, { cmd: 'paraBorders', label: 'Bordes y sombreado', icon: 'paraBorder' }, { cmd: 'marks', label: 'Marcas de formato (Ctrl+F10)', icon: 'pilcrow' }, { cmd: 'tabs', label: 'Tabulaciones…', icon: 'tabs' }, { cmd: 'paraDialog', label: 'Párrafo…', icon: 'dialogLauncher' }] },
        ] },
      ] },
      { label: 'Estilos', items: [{ render: styleGallery }, { stack: [{ cmd: 'stylesPane', label: 'Panel de estilos (Ctrl+F11)', icon: 'styles' }, { cmd: 'newStyle', label: 'Nuevo estilo a partir de la selección…', icon: 'plus' }, { cmd: 'chapters', label: 'Numeración de capítulos', icon: 'chapters', menu: () => ES.chapterMenu() }] }] },
      { label: 'Edición', items: [{ stack: [{ cmd: 'find', label: 'Buscar', icon: 'find', showLabel: true, kbd: 'Ctrl+F' }, { cmd: 'replace', label: 'Reemplazar', icon: 'replace', showLabel: true, kbd: 'Ctrl+H' }, { cmd: 'selectAll', label: 'Seleccionar', icon: 'check', showLabel: true }] }] },
    ],
  },
  {
    id: 'insert', label: 'Insertar', groups: [
      { label: 'Páginas', items: [{ cmd: 'pageBreak', label: 'Salto de página', icon: 'pageBreak', size: 'lg', kbd: 'Ctrl+Enter' }] },
      { label: 'Tablas', items: [{ cmd: 'insertTable', label: 'Tabla', icon: 'table', size: 'lg' }] },
      { label: 'Ilustraciones', items: [{ cmd: 'insertImage', label: 'Imagen', icon: 'image', size: 'lg' }, { cmd: 'shapes', label: 'Formas', icon: 'star', size: 'lg', menu: () => F.shapeMenu() }, { cmd: 'smartart', label: 'SmartArt', icon: 'layout', size: 'lg' }] },
      { label: 'Texto', items: [{ cmd: 'textbox', label: 'Cuadro de texto', icon: 'textbox', size: 'lg' }] },
      { label: 'Vínculos', items: [{ cmd: 'insertLink', label: 'Vínculo', icon: 'link', size: 'lg', kbd: 'Ctrl+K' }] },
      { label: 'Encabezado y pie', items: [{ cmd: 'headerFooter', label: 'Encabezado y pie', icon: 'header', size: 'lg' }, { cmd: 'headerFooter', label: 'Número de página', icon: 'pageNum', size: 'lg' }] },
      { label: 'Símbolos', items: [{ cmd: 'equation', label: 'Ecuación', icon: 'func', size: 'lg' }, { cmd: 'symbol', label: 'Símbolo', icon: 'symbol', size: 'lg' }, { cmd: 'hr', label: 'Línea horizontal', icon: 'hr', size: 'lg' }] },
    ],
  },
  {
    id: 'layout', label: 'Disposición', groups: [
      { label: 'Configurar página', items: [
        { cmd: 'margins', label: 'Márgenes', icon: 'margins', size: 'lg', menu: marginMenu },
        { cmd: 'orientation', label: 'Orientación', icon: 'orientation', size: 'lg' },
        { cmd: 'size', label: 'Tamaño', icon: 'pageSetup', size: 'lg', menu: sizeMenu },
        { cmd: 'columns', label: 'Columnas', icon: 'columns', size: 'lg', menu: () => [{ label: 'Una columna', run: () => F.setColumns(1) }, { label: 'Dos columnas', run: () => F.setColumns(2) }, { label: 'Tres columnas', run: () => F.setColumns(3) }, '-', { label: 'Se aplica a los párrafos seleccionados', run: () => {} }] },
        { cmd: 'pageSetup', label: 'Más opciones', icon: 'pageSetup', size: 'lg' },
      ] },
      { label: 'Párrafo', items: [{ stack: [{ cmd: 'indent', label: 'Aumentar sangría', icon: 'indent', showLabel: true }, { cmd: 'outdent', label: 'Disminuir sangría', icon: 'outdent', showLabel: true }, { cmd: 'lineSpacing', label: 'Interlineado', icon: 'lineSpacing', showLabel: true, menu: lineMenu }] }] },
    ],
  },
  {
    id: 'design', label: 'Diseño', groups: [
      { label: 'Fondo de página', items: [{ cmd: 'watermark', label: 'Marca de agua', icon: 'background', size: 'lg' }, { cmd: 'pageColor', label: 'Color de página', icon: 'fill', size: 'lg', color: true, colorValue: '#FFF8E1', colorOpts: { noneLabel: 'Sin color' } }] },
    ],
  },
  {
    id: 'refs', label: 'Referencias', groups: [
      { label: 'Tabla de contenido', items: [{ cmd: 'toc', label: 'Índice / actualizar', icon: 'styles', size: 'lg' }] },
      { label: 'Notas al pie', items: [{ cmd: 'footnote', label: 'Insertar nota al pie', icon: 'sup', size: 'lg' }, { cmd: 'endnote', label: 'Insertar nota al final', icon: 'sup' }] },
      { label: 'Citas y bibliografía', items: [{ cmd: 'cite', label: 'Insertar cita', icon: 'comment', size: 'lg', menu: () => F3.citeMenu() }, { stack: [{ cmd: 'sources', label: 'Administrar fuentes', icon: 'folder', showLabel: true }, { cmd: 'bibliography', label: 'Bibliografía', icon: 'styles', showLabel: true }] }] },
    ],
  },
  {
    id: 'mail', label: 'Correspondencia', groups: [
      { label: 'Crear', items: [{ cmd: 'envelope', label: 'Sobres', icon: 'doc', size: 'lg' }, { cmd: 'labels', label: 'Etiquetas', icon: 'table', size: 'lg' }] },
      { label: 'Iniciar combinación', items: [{ cmd: 'recipients', label: 'Elegir destinatarios', icon: 'open', size: 'lg' }] },
      { label: 'Escribir e insertar campos', items: [{ cmd: 'fields', label: 'Insertar campo', icon: 'textbox', size: 'lg', menu: () => F.fieldMenu() }] },
      { label: 'Vista previa', items: [{ cmd: 'mergePreview', label: 'Vista previa', icon: 'find', size: 'lg' }, { stack: [{ cmd: 'mergePrev', label: 'Anterior', icon: 'undo', showLabel: true }, { cmd: 'mergeNext', label: 'Siguiente', icon: 'redo', showLabel: true }] }] },
      { label: 'Finalizar', items: [{ cmd: 'mergeFinish', label: 'Finalizar y combinar', icon: 'check', size: 'lg' }] },
    ],
  },
  {
    id: 'review', label: 'Revisar', groups: [
      { label: 'Revisión', items: [{ cmd: 'spell', label: 'Ortografía', icon: 'spell', size: 'lg' }, { cmd: 'wordCount', label: 'Contar palabras', icon: 'words', size: 'lg' }] },
      { label: 'Comentarios', items: [{ cmd: 'comment', label: 'Nuevo comentario', icon: 'comment', size: 'lg' }, { cmd: 'showComments', label: 'Mostrar comentarios', icon: 'notes', size: 'lg' }] },
      { label: 'Seguimiento', items: [{ cmd: 'track', label: 'Control de cambios', icon: 'replace', size: 'lg' }, { cmd: 'showChanges', label: 'Panel de revisiones', icon: 'words', size: 'lg' }] },
      { label: 'Comparar', items: [{ cmd: 'compare', label: 'Comparar', icon: 'columns', size: 'lg' }] },
      { label: 'Cambios', items: [{ stack: [{ cmd: 'revAccept', label: 'Aceptar', icon: 'check', showLabel: true, menu: null }, { cmd: 'revReject', label: 'Rechazar', icon: 'close', showLabel: true }, { cmd: 'revAcceptAll', label: 'Aceptar todos', icon: 'check', showLabel: true }] }, { stack: [{ cmd: 'revPrev', label: 'Anterior', icon: 'undo', showLabel: true }, { cmd: 'revNext', label: 'Siguiente', icon: 'redo', showLabel: true }, { cmd: 'revRejectAll', label: 'Rechazar todos', icon: 'trash', showLabel: true }] }] },
    ],
  },
  {
    id: 'view', label: 'Vista', groups: [
      { label: 'Vistas', items: [{ cmd: 'readingMode', label: 'Modo lectura', icon: 'doc', size: 'lg' }, { cmd: 'focusMode', label: 'Modo enfoque', icon: 'zoomIn', size: 'lg' }] },
      { label: 'Mostrar', items: [{ cmd: 'navPane', label: 'Panel de navegación', icon: 'styles', size: 'lg' }] },
      { label: 'Zoom', items: [{ cmd: 'zoomReset', label: '100 %', icon: 'zoomIn', size: 'lg' }, { cmd: 'zoomWidth', label: 'Ancho de página', icon: 'columns', size: 'lg' }, { cmd: 'zoomPage', label: 'Página completa', icon: 'doc', size: 'lg' }] },
    ],
  },
  {
    id: 'table', label: 'Tabla', contextual: true, groups: [
      { label: 'Filas y columnas', items: [
        { cmd: 'rowAbove', label: 'Fila arriba', icon: 'rowAdd', size: 'lg' }, { cmd: 'rowBelow', label: 'Fila abajo', icon: 'rowAdd', size: 'lg' },
        { cmd: 'colLeft', label: 'Columna a la izquierda', icon: 'colAdd', size: 'lg' }, { cmd: 'colRight', label: 'Columna a la derecha', icon: 'colAdd', size: 'lg' },
      ] },
      { label: 'Eliminar', items: [{ stack: [{ cmd: 'delRow', label: 'Eliminar fila', icon: 'rowDel', showLabel: true }, { cmd: 'delCol', label: 'Eliminar columna', icon: 'colDel', showLabel: true }, { cmd: 'delTable', label: 'Eliminar tabla', icon: 'trash', showLabel: true }] }] },
      { label: 'Estilo', items: [{ cmd: 'cellShade', label: 'Sombreado', icon: 'fill', color: true, colorValue: '#E3F2FD', size: 'sm' }, { cmd: 'header', label: 'Fila de encabezado', icon: 'table', showLabel: true }] },
    ],
  },
  {
    id: 'picture', label: 'Imagen', contextual: true, groups: [
      { label: 'Posición', items: [{ cmd: 'imgLeft', label: 'Izquierda', icon: 'alignLeft', size: 'lg' }, { cmd: 'imgCenter', label: 'Centro', icon: 'alignCenter', size: 'lg' }, { cmd: 'imgRight', label: 'Derecha', icon: 'alignRight', size: 'lg' }] },
      { label: 'Tamaño', items: [{ cmd: 'imgSize', label: 'Tamaño…', icon: 'zoomIn', size: 'lg' }, { cmd: 'imgReset', label: 'Restablecer', icon: 'undo', size: 'lg' }, { cmd: 'imgDelete', label: 'Eliminar', icon: 'trash', size: 'lg' }] },
    ],
  },
];

// ------------------------------------------------------------------ arranque
let zoom;
shell = new OfficeShell({ appId: 'escribiente', appName: 'Escribiente', logo: '../shared/img/escribiente.svg', tabs: TABS, adapter });
// Cortar/copiar/pegar desde la cinta usan el portapapeles del sistema
const origCommand = shell.command.bind(shell);
shell.command = async (cmd, arg) => {
  if (cmd === 'cut' || cmd === 'copy') { restoreSelection(); document.execCommand(cmd); return; }
  if (cmd === 'paste') {
    restoreSelection();
    try {
      const items = await navigator.clipboard.read();
      for (const it of items) {
        if (it.types.includes('text/html')) { document.execCommand('insertHTML', false, sanitize(await (await it.getType('text/html')).text())); shell.markDirty(); return; }
        if (it.types.includes('text/plain')) { document.execCommand('insertText', false, await (await it.getType('text/plain')).text()); shell.markDirty(); return; }
      }
    } catch { toast('Usa Ctrl+V para pegar.'); }
    return;
  }
  return origCommand(cmd, arg);
};
shell.ribbon.onBeforeSelect = saveSelection;
shell.ribbon.onRender = () => setTimeout(reflectState, 0);
buildWorkspace(shell.workspace);
F.initFeatures({
  renderStylesPane: () => ES.renderPane(),
  editor, state, shell, scroller, ruler, pagesEl, schedulePaginate, paginate, insertHTMLAtCursor, restoreSelection, saveSelection, selectedBlocks, currentBlock, applySpanStyle,
  setContent: (html) => setContent(html), toDocx: (el) => htmlToDocx(el, { ...snapshotSettings(), ...shell.docProps() }),
});
ES.init({ editor, state, shell, selectedBlocks, currentBlock, restoreSelection, saveSelection, schedulePaginate, setBlockStyle, applySpanStyles, FONTS, reflectState, fixLists, listCmd, renderSide: () => F.renderSide(), onChange: () => fillGallery() });
FX.init({ editor, selectedBlocks, restoreSelection, saveSelection, markDirty: () => shell.markDirty(), schedulePaginate, reflectState, FONTS, applySpanStyles, exec });
F3.init({ editor, state, shell, insertHTMLAtCursor, restoreSelection, saveSelection, schedulePaginate, paginate, setContent: (html, s) => setContent(html, s), zoomWidth: () => onCommand('zoomWidth'), setZoom: (v) => zoom.set(v) });
adapter.onPrefs = (p) => { editor.spellcheck = p.spellcheck !== false; };
statusPages = h('span', {}, 'Página 1 de 1');
statusWords = h('button', { class: 'sb-btn', title: 'Contar palabras', onclick: () => onCommand('wordCount') }, '0 palabras');
zoom = shell.zoomControl({ value: 100, min: 30, max: 300, onChange: (v) => { state.zoom = v; applyZoom(); drawRuler(); positionImgBox(); } });
// Control de cambios visible en la barra de estado (como Word): un clic lo activa o desactiva
const statusTrack = h('button', { class: 'sb-btn sb-track', title: 'Clic para activar o desactivar el control de cambios', onclick: () => onCommand('track') });
const syncTrack = () => { const on = editor.classList.contains('tracking'); statusTrack.textContent = t(on ? 'Control de cambios: activado' : 'Control de cambios: desactivado'); statusTrack.classList.toggle('on', on); shell.ribbon.setOn('track', on); };
new MutationObserver(syncTrack).observe(editor, { attributes: true, attributeFilter: ['class'] });
shell.status.append(statusPages, statusWords, statusTrack, h('span', {}, 'Español (México)'), h('div', { class: 'grow' }), zoom);
syncTrack();
addEventListener('resize', () => { applyZoom(); drawRuler(); });
editor.innerHTML = '<p><br></p>';
schedulePaginate();
shell.start();

// Para pruebas automáticas
window.__escribiente = { editor, state, paginate, setContent, onCommand, buildPrint, F, F3, F_X: FX, ES, saveSelection, insertCite: (s) => F3.insertCitation(s) };
