// Escribiente — Revisión de documento (recorrido de ortografía como «Ortografía y gramática» de
// LibreOffice/Word: omitir, omitir todo, agregar al diccionario, cambiar y cambiar todo) y
// traducción con servicios en línea cuando hay Internet. GPL-3.0
import { h, toast, modal, api, t, tf } from '../../shared/office.js';
import { tabbed, field, sel } from '../../shared/dialogs.js';

let C = null; // { editor, shell, saveSelection, restoreSelection, schedulePaginate, selectedBlocks }
export function init(ctx) { C = ctx; }

// ------------------------------------------------------------------ ortografía
const WORD = /[\p{L}][\p{L}\p{M}]*(?:['’][\p{L}]+)*/gu;
const SKIP = '.otab,.cbk,.fld,sup.fn,code,pre,del,.mf,.pb,.toc,.eq,[contenteditable="false"],a[href^="http"],script,style';
const ignoreAll = new Set();
const changeAll = new Map();
let box = null; let cur = null; // { node, start, end, word }

export const isMisspelled = (w) => !!api?.spell?.misspelled(w);
function suggestions(w) { try { return (api?.spell?.suggest(w) || []).slice(0, 8); } catch { return []; } }

function textNodes(from) {
  const tw = document.createTreeWalker(C.editor, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.parentElement.closest(SKIP) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) });
  if (from) tw.currentNode = from;
  return tw;
}
// siguiente palabra mal escrita a partir de (node, offset); aplica «Cambiar todo» en el camino
function findNext(node, offset) {
  const tw = textNodes(node && C.editor.contains(node) ? node : null);
  let n = node && C.editor.contains(node) && node.nodeType === 3 ? node : tw.nextNode();
  let off = n === node ? offset : 0;
  while (n) {
    if (!n.parentElement.closest(SKIP)) {
      WORD.lastIndex = 0; let m;
      const txt = n.nodeValue;
      while ((m = WORD.exec(txt))) {
        if (m.index < off) continue;
        const w = m[0];
        if (w.length < 2 || ignoreAll.has(w.toLowerCase())) continue;
        // palabras unidas a números o con mayúsculas en medio (siglas, códigos) no se revisan
        const before = txt[m.index - 1] || ''; const after = txt[m.index + w.length] || '';
        if (/\d/.test(before + after) || (w.length > 1 && w === w.toUpperCase() && w.length <= 6)) continue;
        if (changeAll.has(w)) {
          const r = document.createRange(); r.setStart(n, m.index); r.setEnd(n, m.index + w.length);
          replaceRange(r, changeAll.get(w));
          return findNext(n, m.index + changeAll.get(w).length);
        }
        if (isMisspelled(w)) return { node: n, start: m.index, end: m.index + w.length, word: w };
      }
    }
    n = tw.nextNode(); off = 0;
  }
  return null;
}
function rangeOf(c) { const r = document.createRange(); r.setStart(c.node, c.start); r.setEnd(c.node, c.end); return r; }
function replaceRange(r, text) {
  C.editor.focus();
  const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  document.execCommand('insertText', false, text);
  C.shell.markDirty(); C.schedulePaginate();
}
function highlight(r) {
  try { if (r) CSS.highlights.set('spell-cur', new Highlight(r)); else CSS.highlights.delete('spell-cur'); } catch { /* sin Highlight API */ }
}

function closeBox() { highlight(null); box?.remove(); box = null; cur = null; document.removeEventListener('keydown', onBoxKey, true); }
function onBoxKey(e) { if (e.key === 'Escape' && box) { e.preventDefault(); e.stopPropagation(); closeBox(); C.editor.focus(); } }

function show(c) {
  cur = c;
  if (!c) {
    highlight(null);
    box.querySelector('.sp-ctx').replaceChildren(h('b', {}, t('La revisión ortográfica ha terminado.')));
    box.querySelector('.sp-word').value = ''; box.querySelector('.sp-list').replaceChildren();
    box.querySelectorAll('.sp-act').forEach((b) => { b.disabled = true; });
    box.querySelector('.sp-close').focus();
    box.dataset.done = '1';
    return;
  }
  const r = rangeOf(c);
  highlight(r);
  const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  (c.node.parentElement).scrollIntoView({ block: 'center', behavior: 'instant' });
  const txt = c.node.nodeValue;
  const a = Math.max(0, c.start - 60); const b = Math.min(txt.length, c.end + 60);
  box.querySelector('.sp-ctx').replaceChildren(a > 0 ? '…' : '', txt.slice(a, c.start), h('mark', {}, c.word), txt.slice(c.end, b), b < txt.length ? '…' : '');
  const list = box.querySelector('.sp-list'); list.replaceChildren();
  const sug = suggestions(c.word);
  const inp = box.querySelector('.sp-word'); inp.value = sug[0] || c.word;
  if (!sug.length) list.append(h('div', { class: 'sp-none' }, t('(Sin sugerencias)')));
  sug.forEach((w, i) => {
    const it = h('button', { type: 'button', class: `sp-sug${i === 0 ? ' on' : ''}`, onclick: () => { list.querySelectorAll('.sp-sug').forEach((x) => x.classList.remove('on')); it.classList.add('on'); inp.value = w; }, ondblclick: () => { inp.value = w; act('change'); } }, w);
    list.append(it);
  });
  box.querySelectorAll('.sp-act').forEach((x) => { x.disabled = false; });
  box.dataset.done = '';
  box.querySelector('.sp-count').textContent = tf('Palabras revisadas: {0}', String(++box._n));
}
function next(fromNode, fromOff) { show(findNext(fromNode, fromOff)); }

async function act(kind) {
  if (!cur) return;
  const c = cur; const inp = box.querySelector('.sp-word'); const val = inp.value.trim();
  if (kind === 'ignore') return next(c.node, c.end);
  if (kind === 'ignoreAll') { ignoreAll.add(c.word.toLowerCase()); return next(c.node, c.end); }
  if (kind === 'add') { ignoreAll.add(c.word.toLowerCase()); try { await api?.spell?.addWord(c.word); } catch { /* nada */ } return next(c.node, c.end); }
  if (!val || val === c.word) return next(c.node, c.end);
  // conserva la mayúscula inicial
  const fixed = c.word[0] === c.word[0].toUpperCase() && c.word[0] !== c.word[0].toLowerCase() ? val[0].toUpperCase() + val.slice(1) : val;
  if (kind === 'changeAll') changeAll.set(c.word, fixed);
  replaceRange(rangeOf(c), fixed);
  const s = getSelection();
  const n = s.anchorNode; const o = s.anchorOffset;
  box.querySelector('.sp-word').focus();
  next(n && n.nodeType === 3 ? n : c.node, n && n.nodeType === 3 ? o : c.start + fixed.length);
}

// Revisar → Revisión de documento (F7)
export function documentReview() {
  if (!api?.spell) { toast('La revisión ortográfica no está disponible.', { error: true }); return; }
  closeBox();
  C.saveSelection?.();
  const B = (label, kind, primary) => h('button', { type: 'button', class: `btn sp-act${primary ? ' primary' : ''}`, 'data-act': kind, onclick: () => act(kind) }, t(label));
  box = h('div', { class: 'modal spellbox', role: 'dialog', 'aria-label': t('Revisión de documento') },
    h('header', { class: 'sp-head' }, t('Revisión de documento'), h('button', { type: 'button', class: 'sp-x', 'aria-label': t('Cerrar'), onclick: () => { closeBox(); C.editor.focus(); } }, '×')),
    h('div', { class: 'body' },
      h('div', { class: 'sp-lab' }, t('No está en el diccionario:')),
      h('div', { class: 'sp-ctx', 'aria-live': 'polite' }),
      h('div', { class: 'sp-cols' },
        h('div', { class: 'sp-left' },
          h('label', { class: 'sp-lab' }, t('Cambiar por:'), h('input', { class: 'sp-word', type: 'text', spellcheck: 'false', onkeydown: (e) => { if (e.key === 'Enter') { e.preventDefault(); act('change'); } } })),
          h('div', { class: 'sp-lab' }, t('Sugerencias:')),
          h('div', { class: 'sp-list', role: 'listbox' })),
        h('div', { class: 'sp-btns' }, B('Omitir una vez', 'ignore'), B('Omitir todo', 'ignoreAll'), B('Agregar al diccionario', 'add'), B('Cambiar', 'change', true), B('Cambiar todo', 'changeAll'))),
      h('div', { class: 'sp-foot' }, h('span', { class: 'sp-count' }), h('button', { type: 'button', class: 'btn sp-close', onclick: () => { closeBox(); C.editor.focus(); } }, t('Cerrar')))));
  box._n = 0;
  document.body.append(box);
  document.addEventListener('keydown', onBoxKey, true);
  // empieza en el cursor, como Word; si no hay cursor, desde el principio
  const s = getSelection();
  let start = null; let off = 0;
  if (s.rangeCount && C.editor.contains(s.anchorNode) && s.anchorNode.nodeType === 3) {
    start = s.anchorNode; off = s.anchorOffset;
    // retrocede al principio de la palabra
    while (off > 0 && /[\p{L}\p{M}'’]/u.test(start.nodeValue[off - 1])) off--;
  }
  const first = findNext(start, off) || (start ? findNext(null, 0) : null);
  show(first);
  if (first) box.querySelector('.sp-word').focus();
}
// para las pruebas
export const _state = () => ({ open: !!box, word: cur?.word || null, done: box?.dataset.done === '1' });

// ------------------------------------------------------------------ traducción
export const LANGS = [['es', 'Español'], ['en', 'Inglés'], ['pt', 'Portugués'], ['fr', 'Francés'], ['it', 'Italiano'], ['de', 'Alemán'], ['ca', 'Catalán'], ['gl', 'Gallego'], ['eu', 'Euskera'], ['nl', 'Neerlandés'], ['pl', 'Polaco'], ['ru', 'Ruso'], ['uk', 'Ucraniano'], ['zh', 'Chino'], ['ja', 'Japonés'], ['ko', 'Coreano'], ['ar', 'Árabe'], ['hi', 'Hindi'], ['tr', 'Turco'], ['el', 'Griego'], ['sv', 'Sueco'], ['da', 'Danés'], ['fi', 'Finés'], ['nb', 'Noruego'], ['cs', 'Checo'], ['ro', 'Rumano'], ['hu', 'Húngaro'], ['he', 'Hebreo'], ['id', 'Indonesio'], ['vi', 'Vietnamita'], ['qu', 'Quechua'], ['gn', 'Guaraní'], ['nah', 'Náhuatl'], ['yua', 'Maya yucateco']];
const PROVIDERS = [['google', 'Google Translate (gratuito, sin cuenta)'], ['mymemory', 'MyMemory (gratuito, ~5 000 caracteres al día)'], ['deepl', 'DeepL (requiere clave de API)'], ['libre', 'LibreTranslate (servidor propio o público)']];
const PREF = 'escribiente.traducir';
const loadPref = () => { try { return JSON.parse(localStorage.getItem(PREF) || '{}'); } catch { return {}; } };
const savePref = (p) => { try { localStorage.setItem(PREF, JSON.stringify(p)); } catch { /* nada */ } };

const BLOCKS = 'p,h1,h2,h3,h4,h5,h6,li,blockquote,td,th,figcaption,dt,dd';
// bloques que se traducen (los más internos: una celda con párrafos se traduce por párrafo)
function leafBlocks(root) { return [...root.querySelectorAll(BLOCKS)].filter((b) => !b.querySelector(BLOCKS) && !b.closest(SKIP) && b.textContent.trim()); }
function blockText(b) {
  let s = ''; const tw = document.createTreeWalker(b, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.parentElement.closest(SKIP) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) });
  while (tw.nextNode()) s += tw.currentNode.nodeValue;
  return s.replace(/​/g, '');
}
// pone el texto traducido en el primer texto del bloque (conserva estilo de párrafo, del primer
// tramo, imágenes, notas al pie y campos) y vacía los demás textos
function setBlockText(b, text) {
  const tw = document.createTreeWalker(b, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.parentElement.closest(SKIP) || !n.nodeValue.replace(/​/g, '').trim() ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) });
  const nodes = []; while (tw.nextNode()) nodes.push(tw.currentNode);
  if (!nodes.length) return;
  nodes[0].nodeValue = text; nodes.slice(1).forEach((n) => { const p = n.parentElement; n.remove(); if (p !== b && !p.textContent && !p.querySelector('img,br,sup') && p.matches('span,b,strong,i,em,u,s,font,mark,small,sub,sup')) p.remove(); });
}

async function callTranslate(texts, o) {
  if (!api?.translate) throw Object.assign(new Error(t('La traducción no está disponible.')), { soft: true });
  if (typeof navigator !== 'undefined' && navigator.onLine === false) throw Object.assign(new Error(t('No hay conexión a Internet. La traducción usa servicios en línea.')), { soft: true });
  const res = await api.translate({ ...o, texts });
  if (!res?.ok) throw Object.assign(new Error(res?.offline ? t('No hay conexión a Internet. La traducción usa servicios en línea.') : `${t('El servicio de traducción respondió con un error:')} ${res?.error || ''}`), { soft: true });
  return res;
}

// Revisar → Traducir
export async function translateDialog() {
  C.saveSelection?.();
  const s0 = getSelection();
  const hasSel = s0.rangeCount && !s0.isCollapsed && C.editor.contains(s0.anchorNode) && s0.toString().trim();
  const range = hasSel ? s0.getRangeAt(0).cloneRange() : null;
  const p = { provider: 'google', from: 'auto', to: 'en', ...loadPref() };
  const body = await tabbed('Traducir', [{ label: 'Traducir', body: h('div', { class: 'fx-grid' },
    field('Texto', sel('scope', hasSel ? 'sel' : 'doc', [...(hasSel ? [['sel', 'El texto seleccionado']] : []), ['doc', 'Todo el documento']])),
    field('Del idioma', sel('from', p.from, [['auto', 'Detectar idioma'], ...LANGS])),
    field('Al idioma', sel('to', p.to, LANGS)),
    field('Resultado', sel('mode', p.mode || 'replace', [['replace', 'Reemplazar el texto'], ['below', 'Insertar la traducción después'], ['copy', 'Mostrarla para copiarla']])),
    field('Servicio', sel('provider', p.provider, PROVIDERS)),
    field('Clave de API', h('input', { name: 'key', type: 'password', value: p.key || '', placeholder: t('(solo DeepL o LibreTranslate)'), autocomplete: 'off' })),
    field('Servidor', h('input', { name: 'url', type: 'url', value: p.url || '', placeholder: 'https://libretranslate.com' })),
    h('p', { class: 'fx-note', style: { gridColumn: '1 / -1' } }, t('Se necesita conexión a Internet. Solo se envía el texto que se va a traducir al servicio elegido.'))) }], { width: '520px' });
  if (!body) return null;
  const v = (n) => body.querySelector(`[name=${n}]`).value;
  const o = { provider: v('provider'), from: v('from'), to: v('to'), key: v('key').trim(), url: v('url').trim() };
  const mode = v('mode'); const scope = v('scope');
  savePref({ ...o, mode });
  if (o.from !== 'auto' && o.from === o.to) { toast('Elige dos idiomas distintos.', { error: true }); return null; }
  toast('Traduciendo…');
  try {
    if (scope === 'sel' && range) return await translateSelection(range, o, mode);
    return await translateDocument(o, mode);
  } catch (e) { toast(e.message || String(e), { error: true }); return null; }
}

async function translateSelection(range, o, mode) {
  const lines = range.toString().replace(/​/g, '').split(/\r?\n/);
  const res = await callTranslate(lines, o);
  const out = res.texts;
  if (mode === 'copy') return showResult(out.join('\n'));
  C.editor.focus();
  const s = getSelection(); s.removeAllRanges();
  if (mode === 'below') { const r = range.cloneRange(); r.collapse(false); s.addRange(r); } else s.addRange(range);
  if (out.length > 1 || mode === 'below') {
    const esc = (x) => x.replace(/&/g, '&amp;').replace(/</g, '&lt;');
    document.execCommand('insertHTML', false, mode === 'below' ? out.map((l) => `<p>${esc(l) || '<br>'}</p>`).join('') : out.map(esc).join('<br>'));
  } else document.execCommand('insertText', false, out[0]);
  C.shell.markDirty(); C.schedulePaginate();
  toast(tf('Traducido ({0})', langName(res.detected || o.from) + ' → ' + langName(o.to)));
  return out;
}
const langName = (c) => t((LANGS.find((l) => l[0] === c) || [c, c === 'auto' ? 'Detectar idioma' : String(c || '?')])[1]);

async function translateDocument(o, mode) {
  const blocks = leafBlocks(C.editor);
  if (!blocks.length) { toast('El documento no tiene texto.'); return null; }
  const texts = blocks.map(blockText);
  const res = await callTranslate(texts, o);
  if (mode === 'copy') return showResult(res.texts.join('\n\n'));
  const before = C.editor.innerHTML;
  if (mode === 'below') {
    blocks.forEach((b, i) => { const c = b.cloneNode(true); c.removeAttribute('id'); c.querySelectorAll('[id]').forEach((x) => x.removeAttribute('id')); setBlockText(c, res.texts[i]); c.classList.add('trad'); if (b.matches('td,th')) { b.append(h('br'), ...[...c.childNodes]); } else b.after(c); });
  } else blocks.forEach((b, i) => setBlockText(b, res.texts[i]));
  C.shell.markDirty(); C.schedulePaginate();
  toast(tf('Documento traducido ({0})', langName(res.detected || o.from) + ' → ' + langName(o.to)), { action: { label: t('Deshacer'), run: () => { C.editor.innerHTML = before; C.shell.markDirty(); C.schedulePaginate(); } } });
  return res.texts;
}

function showResult(text) {
  const ta = h('textarea', { rows: 12, style: { width: '100%', boxSizing: 'border-box', font: 'inherit' } }, text);
  modal({ title: t('Traducción'), body: h('div', {}, ta), width: '560px', buttons: [{ label: t('Copiar'), value: () => { navigator.clipboard?.writeText(ta.value).catch(() => {}); toast('Copiado'); return true; } }, { label: t('Cerrar'), value: null, primary: true }], onOpen: () => ta.select() });
  return [text];
}
