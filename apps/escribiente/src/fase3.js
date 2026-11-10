// Escribiente — funciones de la fase 3: ecuaciones, citas y bibliografía, comparar documentos,
// etiquetas y sobres, modos de enfoque y lectura, exportar Markdown y diagramas (SmartArt). GPL-3.0
import temml from 'temml';
import { diffWords, diffArrays } from 'diff';
import { h, $, $$, toast, modal, form, esc, api, t, tf, confirmBox } from '../../shared/office.js';
import { docxToHtml, textToHtml, sanitize } from './docx-in.js';
import { odtToHtml } from './odt.js';

let C = null;
export function init(ctx) {
  C = ctx;
  C.state.sources = C.state.sources || [];
  C.editor.addEventListener('click', (e) => {
    const eq = e.target.closest('.eq'); if (eq) { editEquation(eq); return; }
    const ci = e.target.closest('.cite'); if (ci) editSourceById(ci.dataset.src);
  });
  C.editor.addEventListener('dblclick', (e) => { const sa = e.target.closest('img.smartart'); if (sa) editSmartArt(sa); });
  addEventListener('keydown', (e) => { if (e.key === 'Escape' && (document.body.classList.contains('focus-mode') || document.body.classList.contains('reading-mode'))) exitModes(); }, true);
}
const CM = 96 / 2.54;
const author = () => C.shell.prefs?.author || t('Yo');

// ------------------------------------------------------------------ ecuaciones
const EQ_SAMPLES = [
  ['Fracción', '\\frac{a}{b}'], ['Potencia', 'x^{2}'], ['Subíndice', 'x_{i}'], ['Raíz', '\\sqrt{x}'], ['Suma', '\\sum_{i=1}^{n} i'],
  ['Integral', '\\int_{a}^{b} f(x)\\,dx'], ['Fórmula general', 'x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}'], ['Pitágoras', 'a^2 + b^2 = c^2'],
  ['Área del círculo', 'A = \\pi r^2'], ['Matriz', '\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}'],
];
export function renderTex(tex) {
  try { return temml.renderToString(tex, { annotate: true, throwOnError: true }); } catch (err) { return null; }
}
async function equationDialog(tex = '') {
  const ta = h('textarea', { rows: 3, class: 'eq-src', spellcheck: 'false', placeholder: 'Ej.: x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}' }, tex);
  const prev = h('div', { class: 'eq-preview' });
  const err = h('div', { class: 'eq-err' });
  const upd = () => { const m = renderTex(ta.value.trim()); prev.innerHTML = m || ''; err.textContent = ta.value.trim() && !m ? t('La ecuación tiene un error de escritura.') : ''; };
  ta.addEventListener('input', upd);
  const chips = h('div', { class: 'eq-chips' }, ...EQ_SAMPLES.map(([n, s]) => h('button', { class: 'btn ghost', type: 'button', title: s, onclick: () => { const a = ta.selectionStart; ta.value = ta.value.slice(0, a) + s + ta.value.slice(ta.selectionEnd); ta.focus(); upd(); } }, n)));
  const body = h('div', { class: 'eq-dialog' }, h('p', { class: 'muted' }, 'Escribe la ecuación con notación LaTeX o elige un modelo. Se guarda como ecuación editable de Word.'), chips, ta, err, h('div', { class: 'muted small' }, 'Vista previa'), prev);
  upd();
  return modal({ title: tex ? 'Editar ecuación' : 'Insertar ecuación', body, width: '620px', buttons: [{ label: 'Cancelar', value: null }, { label: tex ? 'Guardar' : 'Insertar', primary: true, value: () => ta.value.trim() }] });
}
export async function insertEquation() {
  C.saveSelection();
  const tex = await equationDialog();
  if (!tex) return;
  const m = renderTex(tex); if (!m) { toast('La ecuación tiene un error de escritura.', { error: true }); return; }
  C.restoreSelection(); normalizeCaret();
  insertAtom(`<span class="eq" data-tex="${esc(tex)}">∑</span>&nbsp;`, m);
  C.shell.markDirty(); C.schedulePaginate();
}
async function editEquation(el) {
  const tex = await equationDialog(el.dataset.tex || el.querySelector('annotation')?.textContent || '');
  if (tex == null) return;
  if (!tex) { el.remove(); } else { const m = renderTex(tex); if (!m) { toast('La ecuación tiene un error de escritura.', { error: true }); return; } el.dataset.tex = tex; el.innerHTML = m; }
  C.shell.markDirty(); C.schedulePaginate();
}

// ------------------------------------------------------------------ citas y bibliografía (estilo APA)
const uid = () => `s${Math.random().toString(36).slice(2, 9)}`;
const SRC_TYPES = [['libro', 'Libro'], ['articulo', 'Artículo de revista'], ['web', 'Página web'], ['otro', 'Otro']];
const lastName = (a) => String(a || '').split(';')[0].split(',')[0].trim() || t('Anónimo');
function citeText(s) {
  const authors = String(s.author || '').split(';').map((x) => x.trim()).filter(Boolean);
  const who = !authors.length ? (s.title || t('Anónimo')) : authors.length === 1 ? lastName(authors[0]) : authors.length === 2 ? `${lastName(authors[0])} ${t('y')} ${lastName(authors[1])}` : `${lastName(authors[0])} et al.`;
  return `(${who}, ${s.year || t('s. f.')})`;
}
function initials(name) { const parts = String(name).split(','); const ln = parts[0].trim(); const fn = (parts[1] || '').trim(); return fn ? `${ln}, ${fn.split(/\s+/).map((x) => `${x[0]}.`).join(' ')}` : ln; }
function biblioHtml(s) {
  const authors = String(s.author || '').split(';').map((x) => x.trim()).filter(Boolean).map(initials);
  const who = authors.length ? (authors.length > 1 ? `${authors.slice(0, -1).join(', ')} ${t('y')} ${authors[authors.length - 1]}` : authors[0]) : '';
  const head = who ? `${esc(who)} (${esc(s.year || t('s. f.'))}). ` : '';
  const title = s.type === 'articulo' ? `${esc(s.title || '')}. <i>${esc(s.publisher || '')}</i>` : `<i>${esc(s.title || '')}</i>${s.publisher ? `. ${esc(s.publisher)}` : ''}`;
  return `${head}${title}.${s.url ? ` ${esc(s.url)}` : ''}`;
}
async function sourceForm(s = {}) {
  return form(s.id ? 'Editar fuente' : 'Nueva fuente', [
    { name: 'type', label: 'Tipo de fuente', type: 'select', value: s.type || 'libro', options: SRC_TYPES },
    { name: 'author', label: 'Autor(es): Apellido, Nombre; separa varios con punto y coma', value: s.author || '' },
    { name: 'title', label: 'Título', value: s.title || '' },
    { name: 'year', label: 'Año', value: s.year || '' },
    { name: 'publisher', label: 'Editorial, revista o sitio', value: s.publisher || '' },
    { name: 'url', label: 'Dirección web (opcional)', value: s.url || '' },
  ], 'Guardar');
}
export async function newSource() {
  const v = await sourceForm(); if (!v || !(v.title || v.author)) return null;
  const s = { id: uid(), ...v }; C.state.sources.push(s); C.shell.markDirty(); return s;
}
async function editSourceById(id) {
  const s = C.state.sources.find((x) => x.id === id); if (!s) return;
  const v = await sourceForm(s); if (!v) return;
  Object.assign(s, v); refreshCitations(); C.shell.markDirty();
}
export function refreshCitations() {
  C.editor.querySelectorAll('.cite').forEach((el) => { const s = C.state.sources.find((x) => x.id === el.dataset.src); if (s) el.textContent = citeText(s); });
  if (C.editor.querySelector('[data-biblio]')) updateBibliography(true);
}
export function citeMenu() {
  return [
    ...C.state.sources.map((s) => ({ label: `${citeText(s)} ${s.title || ''}`.slice(0, 70), run: () => insertCitation(s) })),
    ...(C.state.sources.length ? ['-'] : []),
    { label: 'Agregar fuente nueva…', run: async () => { C.saveSelection(); const s = await newSource(); if (s) { C.restoreSelection(); insertCitation(s); } } },
    { label: 'Administrar fuentes…', run: () => manageSources() },
  ];
}
// Chrome saca del párrafo los elementos contenteditable=false insertados con insertHTML:
// se insertan editables y se bloquean después.
function insertAtom(html, inner) {
  const mark = `a${Math.random().toString(36).slice(2, 8)}`;
  C.insertHTMLAtCursor(html.replace(/^<span /, `<span data-atom="${mark}" `));
  // El contenido especial (MathML) se pone después: insertHTML lo desarma
  C.editor.querySelectorAll(`[data-atom="${mark}"]`).forEach((el) => { if (inner != null) el.innerHTML = inner; el.contentEditable = 'false'; el.removeAttribute('data-atom'); el.removeAttribute('style'); });
}
// Si el cursor quedó entre nodos (p. ej. al final de un párrafo), se mueve al final del texto anterior
function normalizeCaret() {
  const sel = getSelection(); if (!sel.rangeCount || !sel.isCollapsed) return;
  const n = sel.anchorNode, o = sel.anchorOffset;
  if (n?.nodeType === 1 && o > 0 && n.childNodes[o - 1]?.nodeType === 3) { const tn = n.childNodes[o - 1]; const r = document.createRange(); r.setStart(tn, tn.length); r.collapse(true); sel.removeAllRanges(); sel.addRange(r); C.saveSelection(); }
}
export function insertCitation(s) {
  if (!C.editor.contains(getSelection().anchorNode)) C.restoreSelection();
  normalizeCaret();
  const sel = getSelection(); const n = sel.anchorNode; const prev = n?.nodeType === 3 ? n.nodeValue.slice(0, sel.anchorOffset) : '';
  if (prev && !/\s$/.test(prev)) { document.execCommand('insertText', false, ' '); C.saveSelection(); }
  insertAtom(`<span class="cite" data-src="${s.id}">${esc(citeText(s))}</span>&nbsp;`);
  C.shell.markDirty(); C.schedulePaginate();
}
export async function manageSources() {
  const list = h('div', { class: 'src-list' });
  const draw = () => {
    list.replaceChildren(...(C.state.sources.length ? C.state.sources.map((s) => h('div', { class: 'src-item' },
      h('div', { html: biblioHtml(s) }),
      h('div', { class: 'row-btns' }, h('button', { class: 'btn ghost', onclick: async () => { await editSourceById(s.id); draw(); } }, 'Editar'),
        h('button', { class: 'btn ghost', onclick: () => { if (C.editor.querySelector(`.cite[data-src="${s.id}"]`)) { toast('Esa fuente está citada en el documento.'); return; } C.state.sources = C.state.sources.filter((x) => x !== s); C.shell.markDirty(); draw(); } }, 'Eliminar')))) : [h('p', { class: 'muted' }, 'Todavía no hay fuentes. Agrega libros, artículos o páginas web para citarlos.')]));
  };
  draw();
  const body = h('div', {}, list, h('button', { class: 'btn', onclick: async () => { await newSource(); draw(); } }, 'Agregar fuente nueva…'));
  await modal({ title: 'Administrar fuentes', body, width: '640px', buttons: [{ label: 'Cerrar', value: true, primary: true }] });
  refreshCitations();
}
export function updateBibliography(silent = false) {
  const used = [...new Set([...C.editor.querySelectorAll('.cite')].map((x) => x.dataset.src))];
  const srcs = (used.length ? C.state.sources.filter((s) => used.includes(s.id)) : C.state.sources).slice().sort((a, b) => lastName(a.author).localeCompare(lastName(b.author)));
  if (!srcs.length) { if (!silent) toast('Agrega fuentes y cítalas para crear la bibliografía.'); return; }
  const html = `<h1 data-biblio="1">${esc(t('Bibliografía'))}</h1>${srcs.map((s) => `<p data-biblio="1" class="biblio-e" style="margin-left:48px;text-indent:-48px">${biblioHtml(s)}</p>`).join('')}`;
  const old = [...C.editor.querySelectorAll('[data-biblio]')];
  if (old.length) { const tmp = h('div', { html }); old[0].before(...tmp.childNodes); old.forEach((x) => x.remove()); } else { C.restoreSelection(); C.insertHTMLAtCursor(html); }
  C.shell.markDirty(); C.schedulePaginate();
  if (!silent) toast('Bibliografía actualizada');
}

// ------------------------------------------------------------------ comparar documentos
function blocksOf(root) {
  const out = [];
  for (const el of root.children) {
    if (el.classList?.contains('pb')) continue;
    if (/^(UL|OL)$/.test(el.tagName)) { for (const li of el.querySelectorAll('li')) out.push({ tag: el.tagName === 'OL' ? 'oli' : 'uli', text: li.textContent }); continue; }
    if (el.tagName === 'TABLE') { for (const tr of el.rows) out.push({ tag: 'p', text: [...tr.cells].map((c) => c.textContent.trim()).join('\t') }); continue; }
    if (el.tagName === 'DIV' && el.children.length) { out.push(...blocksOf(el)); continue; }
    out.push({ tag: /^H[1-6]$/.test(el.tagName) ? el.tagName.toLowerCase() : el.classList?.contains('title') ? 'title' : 'p', text: el.textContent });
  }
  return out;
}
export function compareHtml(origRoot, revRoot, who = t('Comparación')) {
  const A = blocksOf(origRoot), B = blocksOf(revRoot);
  const date = new Date().toISOString();
  const mark = (kind, s) => (s ? `<${kind} data-author="${esc(who)}" data-date="${date}">${esc(s)}</${kind}>` : '');
  const wrap = (tag, inner) => (tag === 'title' ? `<p class="title">${inner || '<br>'}</p>` : tag === 'uli' || tag === 'oli' ? `<${tag === 'uli' ? 'ul' : 'ol'}><li>${inner || '<br>'}</li></${tag === 'uli' ? 'ul' : 'ol'}>` : `<${tag}>${inner || '<br>'}</${tag}>`);
  const parts = diffArrays(A.map((b) => `${b.tag}\u0001${b.text}`), B.map((b) => `${b.tag}\u0001${b.text}`));
  const out = []; let stats = { ins: 0, del: 0 };
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    const dec = (v) => { const [tag, text] = v.split('\u0001'); return { tag, text }; };
    if (!p.added && !p.removed) { p.value.forEach((v) => { const b = dec(v); out.push(wrap(b.tag, esc(b.text))); }); continue; }
    if (p.removed && parts[i + 1]?.added) {
      const rem = p.value.map(dec), add = parts[i + 1].value.map(dec); i++;
      const n = Math.max(rem.length, add.length);
      for (let k = 0; k < n; k++) {
        const a = rem[k], b = add[k];
        if (a && b) { out.push(wrap(b.tag, diffWords(a.text, b.text).map((d) => { if (d.added) stats.ins++; if (d.removed) stats.del++; return d.added ? mark('ins', d.value) : d.removed ? mark('del', d.value) : esc(d.value); }).join(''))); }
        else if (a) { stats.del++; out.push(wrap(a.tag, mark('del', a.text))); }
        else { stats.ins++; out.push(wrap(b.tag, mark('ins', b.text))); }
      }
      continue;
    }
    p.value.map(dec).forEach((b) => { if (p.added) stats.ins++; else stats.del++; out.push(wrap(b.tag, mark(p.added ? 'ins' : 'del', b.text))); });
  }
  return { html: out.join('').replace(/<\/ul><ul>|<\/ol><ol>/g, ''), stats };
}
async function readOther(f) {
  if (f.ext === 'docx') return (await docxToHtml(f.bytes)).html;
  if (f.ext === 'odt') return (await odtToHtml(f.bytes)).html;
  if (f.ext === 'doc' && f.doc) return textToHtml(f.doc.body || '');
  return textToHtml(new TextDecoder().decode(f.bytes || new Uint8Array()));
}
export async function compareDocuments() {
  const v = await modal({ title: 'Comparar documentos', width: '520px', body: h('div', {}, h('p', {}, 'Elige la versión revisada. El resultado muestra lo que se agregó y lo que se quitó como cambios marcados, que puedes aceptar o rechazar.'), h('p', { class: 'muted' }, 'El documento abierto se toma como el original.')), buttons: [{ label: 'Cancelar', value: null }, { label: 'Elegir versión revisada…', primary: true, value: true }] });
  if (!v) return;
  const f = await api.openDialog({ title: t('Elige la versión revisada'), filters: [{ name: t('Documentos'), extensions: ['docx', 'odt', 'doc', 'txt', 'md'] }] });
  if (!f || f.redirected) return;
  if (f.encrypted) { toast('Quita primero la contraseña de la versión revisada.', { error: true }); return; }
  const revRoot = h('div', { html: sanitize(await readOther(f)) });
  const { html, stats } = compareHtml(C.editor, revRoot);
  C.setContent(html);
  C.state.comments = []; C.state.track = true; C.editor.classList.add('tracking');
  C.shell.doc.path = null; C.shell.doc.name = `${t('Comparación')} - ${f.name.replace(/\.[^.]+$/, '')}.docx`; C.shell.doc.ext = 'docx'; C.shell.updateTitle(); C.shell.markDirty();
  toast(tf('Comparación lista: {0} inserciones y {1} eliminaciones. Revísalas en Revisar → Panel de revisiones.', stats.ins, stats.del), { ms: 7000 });
}

// ------------------------------------------------------------------ etiquetas y sobres
const LABELS = {
  5160: { name: 'Avery 5160 · 3 × 10 (6.7 × 2.5 cm)', cols: 3, rows: 10, w: 6.67, h: 2.54, page: [21.59, 27.94], mt: 1.27, ml: 0.48, gap: 0.32 },
  5163: { name: 'Avery 5163 · 2 × 5 (10.2 × 5.1 cm)', cols: 2, rows: 5, w: 10.16, h: 5.08, page: [21.59, 27.94], mt: 1.27, ml: 0.4, gap: 0.48 },
  L7163: { name: 'Avery L7163 · 2 × 7 (9.9 × 3.8 cm, A4)', cols: 2, rows: 7, w: 9.91, h: 3.81, page: [21.0, 29.7], mt: 1.51, ml: 0.47, gap: 0.25 },
};
async function replaceOk() {
  if (C.shell.adapter.isPristine()) return true;
  return confirmBox(t('Reemplazar el documento'), t('Se reemplazará el contenido del documento actual. ¿Continuar?'), t('Continuar'));
}
const fill = (tpl, row, fields) => tpl.replace(/«([^»]+)»/g, (_, f) => { const i = fields.indexOf(f); return i >= 0 ? (row[i] ?? '') : ''; });
export async function labels() {
  const m = C.state.merge;
  const v = await form('Etiquetas', [
    { name: 'preset', label: 'Tipo de etiqueta', type: 'select', value: '5160', options: Object.entries(LABELS).map(([k, x]) => [k, x.name]) },
    { name: 'text', label: m ? 'Texto (usa «Campo» para los datos de los destinatarios)' : 'Texto de cada etiqueta', type: 'textarea', rows: 4, value: m ? m.fields.slice(0, 3).map((x) => `«${x}»`).join('\n') : '' },
    ...(m ? [{ name: 'merge', label: tf('Una etiqueta por destinatario ({0})', m.rows.length), type: 'checkbox', value: true }] : []),
  ], 'Crear');
  if (!v || !(await replaceOk())) return;
  const L = LABELS[v.preset];
  const texts = v.merge && m ? m.rows.map((r) => fill(v.text, r, m.fields)) : Array(L.cols * L.rows).fill(v.text);
  const cell = (s) => `<td style="width:${(L.w * CM).toFixed(1)}px;height:${(L.h * CM).toFixed(1)}px;vertical-align:middle;padding:2px 10px">${esc(s || '').split('\n').map((x) => x || '&nbsp;').join('<br>')}</td>`;
  let rows = '';
  for (let i = 0; i < texts.length; i += L.cols) rows += `<tr style="height:${(L.h * CM).toFixed(1)}px">${Array.from({ length: L.cols }, (_, k) => cell(texts[i + k])).join('')}</tr>`;
  Object.assign(C.state.page, { size: 'custom', w: Math.round(L.page[0] * CM), h: Math.round(L.page[1] * CM), mt: L.mt * CM, mb: 0.5 * CM, ml: L.ml * CM, mr: 0.3 * CM, landscape: false });
  C.setContent(`<table class="labels">${rows}</table><p><br></p>`);
  C.shell.markDirty();
  toast(tf('{0} etiquetas listas para imprimir', texts.length));
}
const ENVELOPES = { dl: ['DL (22 × 11 cm)', 22, 11], c5: ['C5 (22.9 × 16.2 cm)', 22.9, 16.2], n10: ['Núm. 10 (24.1 × 10.5 cm)', 24.13, 10.48] };
export async function envelope() {
  const v = await form('Sobres', [
    { name: 'to', label: 'Destinatario', type: 'textarea', rows: 4, value: '' },
    { name: 'from', label: 'Remitente', type: 'textarea', rows: 3, value: C.shell.prefs?.author || '' },
    { name: 'size', label: 'Tamaño del sobre', type: 'select', value: 'dl', options: Object.entries(ENVELOPES).map(([k, x]) => [k, x[0]]) },
  ], 'Crear');
  if (!v || !(await replaceOk())) return;
  const [, w, hh] = ENVELOPES[v.size];
  Object.assign(C.state.page, { size: 'custom', w: Math.round(w * CM), h: Math.round(hh * CM), mt: 1 * CM, mb: 0.8 * CM, ml: 1.2 * CM, mr: 1.2 * CM, landscape: true });
  const lines = (s) => esc(s || '').split('\n').join('<br>');
  const gap = Math.max(0, Math.round((hh - 1.8) * CM * 0.38));
  C.setContent(`<p style="font-size:9pt">${lines(v.from) || '<br>'}</p><p style="margin-top:${gap}px;margin-left:${Math.round(w * CM * 0.42)}px;font-size:12pt">${lines(v.to) || '<br>'}</p>`);
  C.shell.markDirty();
}

// ------------------------------------------------------------------ modo enfoque y modo lectura
let exitBtn = null;
function showExit(label) { exitBtn?.remove(); exitBtn = h('button', { class: 'mode-exit', onclick: () => exitModes() }, label); document.body.append(exitBtn); }
export function focusMode() {
  exitModes(true);
  document.body.classList.add('focus-mode');
  showExit(t('Salir del modo enfoque (Esc)'));
  setTimeout(() => { dispatchEvent(new Event('resize')); C.editor.focus(); }, 50);
}
let prevZoom = 100;
export function readingMode() {
  exitModes(true);
  document.body.classList.add('reading-mode');
  C.editor.contentEditable = 'false';
  prevZoom = C.state.zoom;
  showExit(t('Salir del modo lectura (Esc)'));
  setTimeout(() => { dispatchEvent(new Event('resize')); C.zoomWidth?.(); }, 50);
}
export function exitModes(quiet) {
  const wasReading = document.body.classList.contains('reading-mode');
  document.body.classList.remove('focus-mode', 'reading-mode');
  C.editor.contentEditable = 'true';
  exitBtn?.remove(); exitBtn = null;
  if (wasReading) C.setZoom?.(prevZoom);
  if (!quiet) setTimeout(() => dispatchEvent(new Event('resize')), 30);
}

// ------------------------------------------------------------------ Markdown
export function htmlToMarkdown(root) {
  const inline = (n) => [...n.childNodes].map((c) => {
    if (c.nodeType === 3) return c.nodeValue.replace(/​/g, '').replace(/([*_`\\[\]])/g, '\\$1');
    if (c.nodeType !== 1) return '';
    const tag = c.tagName; const inner = inline(c);
    if (c.classList.contains('eq')) return `$${c.dataset.tex || c.textContent}$`;
    if (c.classList.contains('otab')) return '\t';
    if (tag === 'B' || tag === 'STRONG' || /bold|[6-9]00/.test(c.style?.fontWeight || '')) return inner.trim() ? `**${inner}**` : inner;
    if (tag === 'I' || tag === 'EM' || c.style?.fontStyle === 'italic') return inner.trim() ? `*${inner}*` : inner;
    if (tag === 'S' || tag === 'STRIKE') return `~~${inner}~~`;
    if (tag === 'DEL') return '';
    if (tag === 'CODE') return `\`${c.textContent}\``;
    if (tag === 'A') return `[${inner}](${c.getAttribute('href') || ''})`;
    if (tag === 'BR') return '  \n';
    if (tag === 'IMG') return c.classList.contains('smartart') ? '' : `![](${(c.getAttribute('src') || '').startsWith('data:') ? 'imagen' : c.getAttribute('src')})`;
    if (tag === 'SUP' && c.classList.contains('fn')) return `[^${c.textContent}]`;
    return inner;
  }).join('');
  const out = [];
  const list = (el, depth = 0) => {
    let i = 1;
    for (const li of el.children) {
      if (li.tagName !== 'LI') continue;
      const clone = li.cloneNode(true); clone.querySelectorAll('ul,ol').forEach((x) => x.remove());
      out.push(`${'  '.repeat(depth)}${el.tagName === 'OL' ? `${i++}.` : '-'} ${inline(clone).trim()}`);
      for (const sub of li.children) if (sub.tagName === 'UL' || sub.tagName === 'OL') list(sub, depth + 1);
    }
  };
  const block = (el) => {
    const tag = el.tagName;
    if (el.classList?.contains('pb')) { out.push('', '---', ''); return; }
    if (/^H[1-6]$/.test(tag)) { out.push(`${'#'.repeat(Number(tag[1]))} ${inline(el).trim()}`, ''); return; }
    if (el.classList?.contains('title')) { out.push(`# ${inline(el).trim()}`, ''); return; }
    if (el.classList?.contains('subtitle')) { out.push(`*${inline(el).trim()}*`, ''); return; }
    if (tag === 'UL' || tag === 'OL') { list(el); out.push(''); return; }
    if (tag === 'BLOCKQUOTE') { out.push(`> ${inline(el).trim()}`, ''); return; }
    if (tag === 'HR') { out.push('---', ''); return; }
    if (tag === 'TABLE') {
      const rows = [...el.rows].map((r) => [...r.cells].map((c) => inline(c).trim().replace(/\|/g, '\\|').replace(/\n/g, ' ')));
      if (rows.length) { out.push(`| ${rows[0].join(' | ')} |`, `| ${rows[0].map(() => '---').join(' | ')} |`, ...rows.slice(1).map((r) => `| ${r.join(' | ')} |`), ''); }
      return;
    }
    if (tag === 'DIV' && el.children.length && !el.classList.contains('eq')) { for (const c of el.children) block(c); return; }
    const txt = inline(el).trim();
    out.push(txt, '');
  };
  for (const el of root.children) block(el);
  const fns = Object.entries(C?.state?.footnotes || {});
  if (fns.length) { const order = [...root.querySelectorAll('sup.fn')].map((x) => x.dataset.fid); fns.sort((a, b) => order.indexOf(a[0]) - order.indexOf(b[0])).forEach(([fid, text], i) => out.push(`[^${i + 1}]: ${text}`)); }
  return `${out.join('\n').replace(/\n{3,}/g, '\n\n').trim()}\n`;
}

// ------------------------------------------------------------------ diagramas (SmartArt)
const SA_TYPES = [['lista', 'Lista de bloques'], ['proceso', 'Proceso (pasos)'], ['ciclo', 'Ciclo'], ['jerarquia', 'Jerarquía (organigrama)'], ['piramide', 'Pirámide']];
const SA_COLORS = [['#5B48E8', 'Índigo'], ['#0E9C74', 'Verde'], ['#DC3A63', 'Rosa'], ['#E07B2A', 'Naranja'], ['#1F8FD6', 'Azul'], ['#4A4F63', 'Carbón']];
const x = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
function shade(hex, f) { const n = parseInt(hex.slice(1), 16); const c = [n >> 16, (n >> 8) & 255, n & 255].map((v) => Math.round(f > 0 ? v + (255 - v) * f : v * (1 + f))); return `#${c.map((v) => v.toString(16).padStart(2, '0')).join('')}`; }
function wrapText(s, max) { const words = String(s).split(/\s+/); const lines = []; let cur = ''; for (const w of words) { if ((`${cur} ${w}`).trim().length > max && cur) { lines.push(cur); cur = w; } else cur = `${cur} ${w}`.trim(); } if (cur) lines.push(cur); return lines.slice(0, 4); }
function textBlock(cx, cy, s, max, size = 15, color = '#fff') {
  const lines = wrapText(s, max); const lh = size * 1.25; const y0 = cy - ((lines.length - 1) * lh) / 2;
  return lines.map((l, i) => `<text x="${cx}" y="${(y0 + i * lh).toFixed(1)}" font-size="${size}" fill="${color}" text-anchor="middle" dominant-baseline="central">${x(l)}</text>`).join('');
}
export function smartArtSvg({ type, items, color }) {
  const W = 640; let H = 300; let body = '';
  const its = items.filter((s) => s.text.trim()); const n = Math.max(1, its.length);
  const fillN = (i) => shade(color, (i / Math.max(1, n - 1)) * 0.45 - 0.05);
  if (type === 'proceso') {
    const gap = 26, bw = (W - 20 - gap * (n - 1)) / n, bh = 110; H = 160;
    its.forEach((it, i) => { const bx = 10 + i * (bw + gap); body += `<rect x="${bx}" y="25" width="${bw}" height="${bh}" rx="12" fill="${fillN(i)}"/>${textBlock(bx + bw / 2, 80, it.text, Math.max(6, Math.floor(bw / 9)))}`; if (i < n - 1) body += `<path d="M${bx + bw + 5} 70 l${gap - 12} 10 l-${gap - 12} 10 z" fill="${shade(color, 0.55)}"/>`; });
  } else if (type === 'ciclo') {
    H = 420; const cx = W / 2, cy = H / 2, R = 150, r = Math.min(62, 300 / n);
    body += `<circle cx="${cx}" cy="${cy}" r="${R}" fill="none" stroke="${shade(color, 0.6)}" stroke-width="10" stroke-dasharray="18 12"/>`;
    its.forEach((it, i) => { const a = -Math.PI / 2 + (i * 2 * Math.PI) / n; const px = cx + R * Math.cos(a), py = cy + R * Math.sin(a); body += `<circle cx="${px.toFixed(1)}" cy="${py.toFixed(1)}" r="${r}" fill="${fillN(i)}"/>${textBlock(px, py, it.text, Math.max(6, Math.floor(r / 5)), 13)}`; });
  } else if (type === 'jerarquia') {
    const roots = []; const stack = [];
    for (const it of its) { const node = { text: it.text, kids: [] }; while (stack.length && stack[stack.length - 1].lvl >= it.lvl) stack.pop(); if (stack.length) stack[stack.length - 1].node.kids.push(node); else roots.push(node); stack.push({ lvl: it.lvl, node }); }
    const levels = []; const walk = (nd, d) => { (levels[d] = levels[d] || []).push(nd); nd.kids.forEach((k) => walk(k, d + 1)); }; roots.forEach((r) => walk(r, 0));
    const bh = 54, vgap = 46; H = levels.length * (bh + vgap) + 10;
    levels.forEach((row, d) => { const bw = Math.min(170, (W - 20 - 14 * (row.length - 1)) / row.length); const total = row.length * bw + (row.length - 1) * 14; row.forEach((nd, i) => { nd.x = (W - total) / 2 + i * (bw + 14); nd.y = 10 + d * (bh + vgap); nd.w = bw; }); });
    levels.forEach((row, d) => row.forEach((nd) => { nd.kids.forEach((k) => { const x1 = nd.x + nd.w / 2, y1 = nd.y + bh, x2 = k.x + k.w / 2, y2 = k.y, ym = (y1 + y2) / 2; body += `<path d="M${x1} ${y1} V${ym} H${x2} V${y2}" fill="none" stroke="${shade(color, 0.4)}" stroke-width="2"/>`; }); }));
    levels.forEach((row, d) => row.forEach((nd) => { body += `<rect x="${nd.x.toFixed(1)}" y="${nd.y}" width="${nd.w.toFixed(1)}" height="${bh}" rx="8" fill="${shade(color, d * 0.18)}"/>${textBlock(nd.x + nd.w / 2, nd.y + bh / 2, nd.text, Math.max(6, Math.floor(nd.w / 8.5)), 13)}`; }));
  } else if (type === 'piramide') {
    const lh = 64; H = n * lh + 20;
    its.forEach((it, i) => { const top = 10 + i * lh; const w0 = ((i) / n) * (W - 40), w1 = ((i + 1) / n) * (W - 40); const cx = W / 2; body += `<path d="M${cx - w0 / 2} ${top} H${cx + w0 / 2} L${cx + w1 / 2} ${top + lh - 4} H${cx - w1 / 2} Z" fill="${fillN(i)}"/>${textBlock(cx, top + lh / 2 + 4, it.text, Math.max(10, Math.floor(((w0 + w1) / 2) / 9)), 14)}`; });
  } else {
    const cols = n <= 3 ? n : n <= 4 ? 2 : 3; const rowsN = Math.ceil(n / cols); const gap = 14; const bw = (W - 20 - gap * (cols - 1)) / cols; const bh = 92; H = rowsN * (bh + gap) + 6;
    its.forEach((it, i) => { const c = i % cols, r = Math.floor(i / cols); const bx = 10 + c * (bw + gap), by = 3 + r * (bh + gap); body += `<rect x="${bx}" y="${by}" width="${bw}" height="${bh}" rx="10" fill="${fillN(i)}"/>${textBlock(bx + bw / 2, by + bh / 2, it.text, Math.max(8, Math.floor(bw / 9)))}`; });
  }
  return { svg: `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="Inter, Carlito, Arial, sans-serif" font-weight="600">${body}</svg>`, w: W, h: H };
}
const parseItems = (text) => text.split('\n').map((l) => ({ lvl: (/^(\s*|-*)/.exec(l)[0].replace(/\t/g, '  ').length), text: l.replace(/^[\s-]+/, '') })).filter((i) => i.text.trim());
async function smartArtDialog(d = {}) {
  return form(d.items ? 'Editar diagrama' : 'Insertar diagrama (SmartArt)', [
    { name: 'type', label: 'Tipo de diagrama', type: 'select', value: d.type || 'proceso', options: SA_TYPES },
    { name: 'text', label: 'Elementos, uno por línea (en Jerarquía, sangra con espacios o guiones los niveles inferiores)', type: 'textarea', rows: 6, value: d.text || [t('Planear'), t('Hacer'), t('Revisar'), t('Mejorar')].join('\n') },
    { name: 'color', label: 'Color', type: 'select', value: d.color || '#5B48E8', options: SA_COLORS },
  ], d.items ? 'Guardar' : 'Insertar');
}
const svgUrl = (svg) => `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(svg)))}`;
export async function insertSmartArt() {
  C.saveSelection();
  const v = await smartArtDialog(); if (!v) return;
  const items = parseItems(v.text); if (!items.length) return;
  const { svg, w, h: hh } = smartArtSvg({ type: v.type, items, color: v.color });
  const width = Math.min(w, C.state.page.w - C.state.page.ml - C.state.page.mr);
  C.restoreSelection();
  C.insertHTMLAtCursor(`<p style="text-align:center"><img class="smartart" width="${Math.round(width)}" height="${Math.round((hh * width) / w)}" data-smartart="${esc(JSON.stringify({ type: v.type, text: v.text, color: v.color }))}" src="${svgUrl(svg)}"></p>`);
  C.shell.markDirty(); C.schedulePaginate();
}
async function editSmartArt(img) {
  let d = {}; try { d = JSON.parse(img.dataset.smartart); } catch { /* */ }
  const v = await smartArtDialog({ ...d, items: true }); if (!v) return;
  const items = parseItems(v.text); if (!items.length) return;
  const { svg, w, h: hh } = smartArtSvg({ type: v.type, items, color: v.color });
  const width = img.width || w;
  img.src = svgUrl(svg); img.height = Math.round((hh * width) / w); img.dataset.smartart = JSON.stringify({ type: v.type, text: v.text, color: v.color });
  C.shell.markDirty(); C.schedulePaginate();
}
