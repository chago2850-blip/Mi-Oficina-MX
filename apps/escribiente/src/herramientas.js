// Escribiente — Herramientas como en LibreOffice Writer: búsqueda con expresiones regulares y palabras
// completas, autocorrección al escribir (reemplazos, mayúscula al empezar oración, DOs MAyúsculas,
// comillas tipográficas, guiones, *negrita* /cursiva/ _subrayado_, vínculos), diccionario de
// sinónimos (RLA-ES), separación silábica (patrones de Liang de Javier Bezos) y ordenar párrafos. GPL-3.0
import { h, toast, form, t, tf, modal, esc } from '../../shared/office.js';
import HYPH from '../../shared/data/hyph-es.json';

let C = null; // { editor, state, shell, saveSelection, restoreSelection, schedulePaginate, currentBlock, selectedBlocks }
export function init(ctx) { C = ctx; C.editor.addEventListener('input', onInput); }
const LETTER = /[\p{L}\p{N}]/u;

// ================================================================ búsqueda
// devuelve los rangos que coinciden; admite expresión regular, palabras completas, acentos y mayúsculas
export function findRanges(root, query, { matchCase = false, accents = false, words = false, regex = false, style = '', inSel = null } = {}) {
  const out = []; if (!query) return out;
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.parentElement?.closest('.fld,.toc,sup.fn,sup.en,.ix,a.bm') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) });
  const nodes = []; let full = '';
  while (walker.nextNode()) { const n = walker.currentNode; const blk = n.parentElement.closest('p,h1,h2,h3,h4,h5,h6,li,td,th,blockquote,div'); const prev = nodes[nodes.length - 1]; if (prev && prev.blk !== blk) full += '\n'; nodes.push({ n, start: full.length, blk }); full += n.nodeValue; }
  // texto sin guiones opcionales, con un mapa a las posiciones reales
  const map = []; let hay = '';
  const prep = (ch) => { if (ch === '­') return ''; let x = ch; if (!accents && !regex) x = x.normalize('NFD').replace(/[̀-ͯ]/g, ''); if (!matchCase && !regex) x = x.toLowerCase(); return x; };
  for (let i = 0; i < full.length; i++) { const p = prep(full[i]); for (let k = 0; k < p.length; k++) map.push(i); hay += p; }
  map.push(full.length);
  let re;
  try {
    const src = regex ? query : (accents ? query : query.normalize('NFD').replace(/[̀-ͯ]/g, '')).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    re = new RegExp(words ? `(?<![\\p{L}\\p{N}_])(?:${src})(?![\\p{L}\\p{N}_])` : src, `gu${matchCase ? '' : 'i'}${regex ? 'm' : ''}`);
  } catch (e) { throw new Error(t('Expresión regular no válida')); }
  const locate = (pos) => { let lo = 0, hi = nodes.length - 1; while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (nodes[mid].start <= pos) lo = mid; else hi = mid - 1; } return lo; };
  let m;
  while ((m = re.exec(hay))) {
    if (!m[0].length) { re.lastIndex++; continue; }
    const s = map[m.index], e = map[m.index + m[0].length - 1] + 1;
    const a = locate(s), b = locate(e - 1);
    if (s - nodes[a].start > nodes[a].n.length || e - nodes[b].start > nodes[b].n.length) continue;
    const r = document.createRange(); r.setStart(nodes[a].n, s - nodes[a].start); r.setEnd(nodes[b].n, e - nodes[b].start);
    if (style) { const blk = nodes[a].blk; const st = blk?.dataset.style || (blk && /^H[1-6]$/.test(blk.tagName) ? `Heading${blk.tagName[1]}` : 'Normal'); if (st !== style) continue; }
    if (inSel && (inSel.compareBoundaryPoints(Range.START_TO_START, r) > 0 || inSel.compareBoundaryPoints(Range.END_TO_END, r) < 0)) continue;
    r.match = m; out.push(r);
  }
  return out;
}
// texto de reemplazo: con expresiones regulares admite $1, $2… y \n
export function replacementFor(r, rep, { regex = false, matchCase = false } = {}) {
  if (!regex || !r.match) return rep;
  const m = r.match; return rep.replace(/\$(\d+)|\$&|\\n|\\t/g, (x, d) => (x === '$&' ? m[0] : x === '\\n' ? '\n' : x === '\\t' ? '\t' : m[Number(d)] ?? '')) + (matchCase ? '' : '');
}

// ================================================================ autocorrección
const DEF_REPL = {
  '(c)': '©', '(r)': '®', '(tm)': '™', '-->': '→', '<--': '←', '->': '→', '<-': '←', '=>': '⇒', '<=>': '⇔', '+-': '±', '1/2': '½', '1/4': '¼', '3/4': '¾', '...': '…', '--': '—', 'x2': 'x2',
  ademas: 'además', tambien: 'también', despues: 'después', aveces: 'a veces', atravez: 'a través', atraves: 'a través', enserio: 'en serio', osea: 'o sea', porfavor: 'por favor', sinembargo: 'sin embargo',
  apesar: 'a pesar', deacuerdo: 'de acuerdo', nisiquiera: 'ni siquiera', quizas: 'quizás', facil: 'fácil', dificil: 'difícil', util: 'útil', arbol: 'árbol', lapiz: 'lápiz', razon: 'razón', corazon: 'corazón',
  cancion: 'canción', informacion: 'información', educacion: 'educación', tecnologia: 'tecnología', telefono: 'teléfono', numero: 'número', pagina: 'página', publico: 'público', ultimo: 'último', unico: 'único',
  rapido: 'rápido', musica: 'música', metodo: 'método', analisis: 'análisis', solucion: 'solución', atencion: 'atención', direccion: 'dirección', organizacion: 'organización', situacion: 'situación', comunicacion: 'comunicación',
  teh: 'the', adn: 'and', taht: 'that', recieve: 'receive', seperate: 'separate', definately: 'definitely', occured: 'occurred',
};
delete DEF_REPL.x2;
const ABBR = /^(etc|sr|sra|srta|dr|dra|lic|ing|ej|p|pág|págs|núm|tel|av|col|no|vol|cap|fig|art|aprox|ud|uds|vs|e\.g|i\.e|mr|mrs|ms)$/i;
const LS = 'oficina.autocorrect';
export function acOptions() {
  let o = {}; try { o = JSON.parse(localStorage.getItem(LS) || '{}'); } catch { /* */ }
  return { on: true, repl: true, capSentence: true, twoCaps: true, quotes: true, dashes: true, emphasis: true, urls: true, custom: {}, removed: [], ...o };
}
const replTable = (o) => { const tbl = { ...DEF_REPL, ...o.custom }; for (const k of o.removed || []) delete tbl[k]; return tbl; };
export async function autocorrectDialog() {
  const o = acOptions(); const tbl = replTable(o);
  const chk = (name, label, on) => h('label', { class: 'fx-check' }, h('input', { type: 'checkbox', name, checked: on ? true : null }), ` ${t(label)}`);
  const ta = h('textarea', { name: 'tbl', rows: 10, style: { width: '100%', fontFamily: 'monospace', fontSize: '12px' } }, Object.entries(tbl).map(([a, b]) => `${a} → ${b}`).join('\n'));
  const body = h('div', { class: 'fx-col', style: { width: '520px' } },
    chk('on', 'Corregir mientras se escribe', o.on), h('h4', {}, t('Opciones')),
    chk('repl', 'Usar la tabla de reemplazo', o.repl), chk('capSentence', 'Poner en mayúscula la primera letra de cada oración', o.capSentence), chk('twoCaps', 'Corregir DOs MAyúsculas INiciales', o.twoCaps),
    chk('quotes', 'Comillas tipográficas “ ” y apóstrofo ’', o.quotes), chk('dashes', 'Reemplazar guiones (– y —)', o.dashes), chk('emphasis', 'Aplicar *negrita*, /cursiva/, _subrayado_ y -tachado-', o.emphasis), chk('urls', 'Reconocer direcciones web', o.urls),
    h('h4', {}, t('Tabla de reemplazo (una por línea: texto → reemplazo)')), ta);
  const ok = await modal({ title: t('Opciones de autocorrección'), body, buttons: [{ label: 'Cancelar', value: false }, { label: 'Aceptar', value: true, primary: true }] });
  if (!ok) return;
  const v = (n) => body.querySelector(`[name="${n}"]`).checked;
  const now = {}; for (const ln of ta.value.split('\n')) { const m = /^(.+?)\s*(?:→|->|=>|\t)\s*(.*)$/.exec(ln.trim()); if (m && m[1] && m[1] !== m[2]) now[m[1]] = m[2]; }
  const custom = {}; for (const [k, val] of Object.entries(now)) if (DEF_REPL[k] !== val) custom[k] = val;
  const removed = Object.keys(DEF_REPL).filter((k) => !(k in now));
  const next = { on: v('on'), repl: v('repl'), capSentence: v('capSentence'), twoCaps: v('twoCaps'), quotes: v('quotes'), dashes: v('dashes'), emphasis: v('emphasis'), urls: v('urls'), custom, removed };
  try { localStorage.setItem(LS, JSON.stringify(next)); } catch { /* */ }
  toast('Opciones de autocorrección guardadas');
}
const TRIGGERS = /^[\s.,;:!?)\]}»”]$/;
// reemplaza el rango por texto (o HTML) con la pila de deshacer del navegador
function swap(r, text, html) {
  const s = getSelection(); const k = s.rangeCount ? { n: s.anchorNode, o: s.anchorOffset } : null;
  const same = k && k.n === r.endContainer; const tail = same ? Math.max(0, k.o - r.endOffset) : 0;
  s.removeAllRanges(); s.addRange(r);
  if (html) document.execCommand('insertHTML', false, html); else document.execCommand('insertText', false, text);
  // el cursor vuelve a quedar después del carácter que disparó la corrección
  if (same) { for (let i = 0; i < tail; i++) getSelection().modify('move', 'forward', 'character'); }
  else if (k?.n?.isConnected) { try { const r2 = document.createRange(); r2.setStart(k.n, Math.min(k.o, k.n.nodeType === 3 ? k.n.length : k.n.childNodes.length)); r2.collapse(true); const s2 = getSelection(); s2.removeAllRanges(); s2.addRange(r2); } catch { /* */ } }
}
// cambia el texto de la palabra y le aplica un formato (negrita, vínculo…) con los comandos del editor
function formatWord(r, text, cmd, arg) {
  const s = getSelection(); const k = s.rangeCount ? { n: s.anchorNode, o: s.anchorOffset } : null;
  const same = k && k.n === r.endContainer; const tail = same ? Math.max(0, k.o - r.endOffset) : 0;
  const node = r.startContainer; const start = r.startOffset;
  s.removeAllRanges(); s.addRange(r); document.execCommand('insertText', false, text);
  const sel = document.createRange(); sel.setStart(node, start); sel.setEnd(node, start + text.length); s.removeAllRanges(); s.addRange(sel);
  document.execCommand(cmd, false, arg ?? null);
  s.collapseToEnd(); if (cmd !== 'createLink') document.execCommand(cmd, false, arg ?? null); // lo que se escriba después ya no lleva el formato
  for (let i = 0; i < tail; i++) getSelection().modify('move', 'forward', 'character');
}
function onInput(e) {
  if (e.isComposing || C.state.track) return;
  const o = acOptions(); if (!o.on) return;
  const s = getSelection(); if (!s.rangeCount || !s.isCollapsed) return;
  const node = s.anchorNode; const off = s.anchorOffset;
  if (!node || node.nodeType !== 3 || node.parentElement?.closest('.fld,code,pre,a,.eq,.toc')) return;
  const text = node.nodeValue.slice(0, off);
  // comillas tipográficas al escribirlas
  if (e.inputType === 'insertText' && o.quotes && (e.data === '"' || e.data === "'")) {
    const prev = text.slice(-2, -1); const open = !prev || /[\s([{«¿¡—–-]/.test(prev);
    const ch = e.data === '"' ? (open ? '“' : '”') : (open ? '‘' : '’');
    const r = document.createRange(); r.setStart(node, off - 1); r.setEnd(node, off); swap(r, ch); return;
  }
  if (!(e.inputType === 'insertText' && TRIGGERS.test(e.data || '')) && e.inputType !== 'insertParagraph') return;
  let tnode = node; let end = off - (e.inputType === 'insertParagraph' ? 0 : 1); let before = tnode.nodeValue.slice(0, end);
  if (e.inputType === 'insertParagraph') { // la palabra quedó al final del párrafo anterior
    const blk = (node.parentElement || node).closest('p,li,h1,h2,h3,h4,h5,h6,blockquote,div'); const prevBlk = blk?.previousElementSibling; if (!prevBlk) return;
    const tw = document.createTreeWalker(prevBlk, NodeFilter.SHOW_TEXT); let last = null; while (tw.nextNode()) last = tw.currentNode; if (!last) return;
    tnode = last; end = last.length; before = last.nodeValue;
  }
  const wm = /(\S+)$/.exec(before); if (!wm) return;
  let word = wm[1]; const start = end - word.length;
  const rng = () => { const r = document.createRange(); r.setStart(tnode, start); r.setEnd(tnode, end); return r; };
  // vínculos
  if (o.urls && /^(https?:\/\/|www\.)[^\s]+\.[^\s]{2,}$/i.test(word) && !tnode.parentElement.closest('a')) { const href = /^www\./i.test(word) ? `https://${word}` : word; formatWord(rng(), word, 'createLink', href); return; }
  // *negrita* /cursiva/ _subrayado_ -tachado-
  const em = o.emphasis && /^([*/_-])([^*/_\s-][^*/_]*[^*/_\s-]|[^*/_\s-])\1$/.exec(word);
  if (em && !/^-\d/.test(word)) { formatWord(rng(), em[2], { '*': 'bold', '/': 'italic', _: 'underline', '-': 'strikeThrough' }[em[1]]); return; }
  // tabla de reemplazo (respeta la mayúscula inicial)
  if (o.repl) {
    const tbl = replTable(o); const whole = tbl[word] != null || tbl[word.toLowerCase()] != null;
    const core = whole ? [word, word, ''] : /^(.*?)([.,;:!?)\]}»”]*)$/.exec(word); const key = core[1];
    let rep = tbl[key] ?? tbl[key.toLowerCase()];
    if (rep != null && rep !== key) {
      if (key[0] && key[0] !== key[0].toLowerCase() && tbl[key] == null) rep = rep[0].toUpperCase() + rep.slice(1);
      const r = document.createRange(); r.setStart(tnode, start); r.setEnd(tnode, start + key.length); swap(r, rep); word = rep + core[2];
      if (e.inputType !== 'insertParagraph') return;
    }
  }
  // guiones: «palabra - palabra» → raya corta
  if (o.dashes && e.data === ' ' && word === '-' && /\S\s$/.test(before.slice(0, start))) { swap(rng(), '–'); return; }
  // DOs MAyúsculas
  if (o.twoCaps && /^\p{Lu}\p{Lu}\p{Ll}{2,}$/u.test(word)) { swap(rng(), word[0] + word.slice(1).toLowerCase()); return; }
  // mayúscula al empezar oración
  if (o.capSentence && /^\p{Ll}/u.test(word) && !/[@/:\\]/.test(word)) {
    const blk = tnode.parentElement.closest('p,li,h1,h2,h3,h4,h5,h6,blockquote,td,th');
    if (blk) {
      const pre = document.createRange(); pre.selectNodeContents(blk); pre.setEnd(tnode, start);
      const prevTxt = pre.toString().replace(/­/g, '');
      const lastWord = (/(\S+)\s*$/.exec(prevTxt.replace(/[.!?]\s*$/, '')) || [])[1] || '';
      const sentenceStart = !prevTxt.trim() || (/[.!?…]["”»)]*\s+$/.test(prevTxt) && !ABBR.test(lastWord.replace(/[.]$/, '')) && !/^\p{L}$/u.test(lastWord));
      if (sentenceStart && !/^[¿¡]/.test(word) ? true : sentenceStart && /^[¿¡]\p{Ll}/u.test(word)) {
        const i = /^[¿¡"“«(]*/.exec(word)[0].length; const r = document.createRange(); r.setStart(tnode, start + i); r.setEnd(tnode, start + i + 1); swap(r, word[i].toUpperCase());
      }
    }
  }
}

// ================================================================ sinónimos
let THES = null;
async function loadThes() {
  if (THES) return THES;
  const res = await fetch('../shared/data/thes-es.json.gz');
  const stream = res.body.pipeThrough(new DecompressionStream('gzip'));
  THES = await new Response(stream).json();
  return THES;
}
const strip = (s) => s.normalize('NFD').replace(/[̀-ͯ]/g, '');
export async function lookup(word) {
  const th = await loadThes(); const w = word.toLowerCase();
  const tries = [w, w.replace(/es$/, ''), w.replace(/s$/, ''), w.replace(/as$/, 'o'), w.replace(/a$/, 'o'), w.replace(/os$/, 'o')];
  for (const k of tries) if (th[k]) return { key: k, groups: th[k] };
  const sw = strip(w); const k2 = Object.keys(th).find((k) => strip(k) === sw);
  return k2 ? { key: k2, groups: th[k2] } : { key: w, groups: [] };
}
function wordAtCaret() {
  C.restoreSelection(); const s = getSelection(); if (!s.rangeCount) return null;
  let r = s.getRangeAt(0).cloneRange();
  if (!r.collapsed && r.toString().trim()) return { range: r, word: r.toString().replace(/­/g, '').trim() };
  const n = r.startContainer; if (n.nodeType !== 3) return null;
  const txt = n.nodeValue; let a = r.startOffset, b = r.startOffset;
  while (a > 0 && /[\p{L}­'’-]/u.test(txt[a - 1])) a--; while (b < txt.length && /[\p{L}­'’-]/u.test(txt[b])) b++;
  if (a === b) return null;
  r = document.createRange(); r.setStart(n, a); r.setEnd(n, b);
  return { range: r, word: txt.slice(a, b).replace(/­/g, '') };
}
export async function thesaurusDialog() {
  const w = wordAtCaret(); if (!w) { toast('Coloca el cursor en una palabra.'); return; }
  let res; try { res = await lookup(w.word); } catch { toast('No se pudo cargar el diccionario de sinónimos.', { error: true }); return; }
  const input = h('input', { name: 'rep', value: w.word, style: { width: '100%' } });
  const list = h('div', { class: 'th-list' });
  const pick = (s) => { input.value = /^\p{Lu}/u.test(w.word) ? s[0].toUpperCase() + s.slice(1) : s; };
  if (!res.groups.length) list.append(h('p', { class: 'fx-note' }, tf('No hay sinónimos para «{0}».', w.word)));
  res.groups.forEach((g, i) => list.append(h('div', { class: 'th-g' }, h('b', {}, `${i + 1}. `), ...g.map((s) => h('button', { type: 'button', class: 'th-w', onclick: () => pick(s), ondblclick: () => { pick(s); document.querySelector('.modal .btn.primary')?.click(); } }, s)))));
  const body = h('div', { class: 'fx-col', style: { width: '520px' } }, h('div', { class: 'fx-note' }, tf('Palabra: {0}', res.key)), list, h('label', { class: 'fx-field' }, h('span', {}, t('Reemplazar con')), input),
    h('p', { class: 'fx-note' }, t('Diccionario de sinónimos de RLA-ES (GPL/LGPL/MPL).')));
  const ok = await modal({ title: t('Sinónimos'), body, buttons: [{ label: 'Cancelar', value: false }, { label: 'Reemplazar', value: true, primary: true }] });
  if (!ok || !input.value || input.value === w.word) return;
  const s = getSelection(); s.removeAllRanges(); s.addRange(w.range); document.execCommand('insertText', false, input.value); C.saveSelection(); C.shell.markDirty();
}

// ================================================================ separación silábica (Liang)
let PAT = null;
function patterns() {
  if (PAT) return PAT;
  PAT = new Map();
  for (const p of HYPH.patterns) { const letters = p.replace(/\d/g, ''); const vals = []; let k = 0; for (const ch of p) { if (/\d/.test(ch)) vals[k] = Number(ch); else { vals[k] = vals[k] || 0; k++; } } vals.length = letters.length + 1; for (let i = 0; i <= letters.length; i++) vals[i] = vals[i] || 0; PAT.set(letters, vals); }
  return PAT;
}
export function hyphenate(word) {
  const P = patterns(); const w = `.${word.toLowerCase()}.`; const pts = new Array(w.length + 1).fill(0);
  for (let i = 0; i < w.length; i++) for (let j = i + 1; j <= Math.min(w.length, i + 10); j++) { const v = P.get(w.slice(i, j)); if (v) for (let k = 0; k < v.length; k++) pts[i + k] = Math.max(pts[i + k], v[k]); }
  const parts = []; let cur = '';
  for (let i = 0; i < word.length; i++) { cur += word[i]; if (pts[i + 2] % 2 === 1 && i + 1 >= HYPH.left && word.length - (i + 1) >= HYPH.right) { parts.push(cur); cur = ''; } }
  parts.push(cur); return parts;
}
// pone guiones opcionales (U+00AD) en las palabras largas; no toca el párrafo donde está el cursor
export function hyphenateDoc(root = C.editor, { skipCaret = true } = {}) {
  const sel = getSelection(); const caretBlk = skipCaret && sel.rangeCount ? (sel.anchorNode?.nodeType === 3 ? sel.anchorNode.parentElement : sel.anchorNode)?.closest?.('p,li,h1,h2,h3,h4,h5,h6,td,th,blockquote') : null;
  let n = 0;
  const tw = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, { acceptNode: (x) => (x.parentElement?.closest('.fld,.toc,code,pre,a,.eq,h1,h2,h3,h4,h5,h6') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) });
  const nodes = []; while (tw.nextNode()) nodes.push(tw.currentNode);
  for (const tn of nodes) {
    if (caretBlk && caretBlk.contains(tn)) continue;
    const v = tn.nodeValue; const nv = v.replace(/[\p{L}­]{7,}/gu, (wd) => { const clean = wd.replace(/­/g, ''); if (!/^\p{Ll}/u.test(clean) && !/^\p{Lu}\p{Ll}+$/u.test(clean)) return wd; return hyphenate(clean).join('­'); });
    if (nv !== v) { tn.nodeValue = nv; n++; }
  }
  return n;
}
export function removeHyphens(root = C.editor) { const tw = document.createTreeWalker(root, NodeFilter.SHOW_TEXT); while (tw.nextNode()) { const n = tw.currentNode; if (n.nodeValue.includes('­')) n.nodeValue = n.nodeValue.replace(/­/g, ''); } }
export async function hyphenationDialog() {
  const v = await form('Separación silábica', [{ name: 'auto', label: 'Separar automáticamente las palabras al final de las líneas', type: 'checkbox', value: !!C.state.hyphenate }], 'Aceptar');
  if (!v) return;
  C.state.hyphenate = v.auto;
  if (v.auto) { const n = hyphenateDoc(C.editor, { skipCaret: false }); toast(tf('Separación silábica activada ({0} párrafos)', n)); } else { removeHyphens(); toast('Separación silábica desactivada'); }
  C.shell.markDirty(); C.schedulePaginate();
}

// ================================================================ ordenar párrafos
export async function sortDialog() {
  C.restoreSelection();
  const blocks = C.selectedBlocks().filter((b) => /^(P|LI|H[1-6])$/.test(b.tagName));
  if (blocks.length < 2) { toast('Selecciona al menos dos párrafos para ordenar.'); return; }
  const v = await form('Ordenar', [
    { name: 'order', label: 'Orden', type: 'select', value: 'asc', options: [['asc', t('Ascendente')], ['desc', t('Descendente')]] },
    { name: 'type', label: 'Tipo de clave', type: 'select', value: 'text', options: [['text', t('Alfanumérico')], ['num', t('Numérico')]] },
    { name: 'col', label: 'Columna (separada por tabulación; 1 = primera)', type: 'number', value: 1, min: 1, max: 20 },
    { name: 'case', label: 'Distinguir mayúsculas y minúsculas', type: 'checkbox', value: false },
  ], 'Ordenar');
  if (!v) return;
  sortBlocks(blocks, v);
}
export function sortBlocks(blocks, { order = 'asc', type = 'text', col = 1, case: cs = false } = {}) {
  const key = (b) => { const c = b.cloneNode(true); c.querySelectorAll('.otab').forEach((x) => x.replaceWith('\t')); const parts = c.textContent.replace(/­/g, '').split('\t'); return (parts[Math.max(0, col - 1)] || '').trim(); };
  const coll = new Intl.Collator(document.documentElement.lang || 'es', { numeric: true, sensitivity: cs ? 'case' : 'base' });
  const num = (s) => { const m = /-?[\d.,]+/.exec(s); return m ? Number(m[0].replace(/,(?=\d{3}\b)/g, '').replace(',', '.')) : -Infinity; };
  const groups = new Map(); for (const b of blocks) { const p = b.parentElement; if (!groups.has(p)) groups.set(p, []); groups.get(p).push(b); }
  for (const list of groups.values()) {
    const anchor = list[0].previousSibling; const parent = list[0].parentElement;
    const sorted = [...list].sort((a, b) => { const r = type === 'num' ? num(key(a)) - num(key(b)) : coll.compare(key(a), key(b)); return order === 'desc' ? -r : r; });
    let ref = anchor; for (const b of sorted) { if (ref) ref.after(b); else parent.prepend(b); ref = b; }
  }
  C.shell.markDirty(); C.schedulePaginate();
}
