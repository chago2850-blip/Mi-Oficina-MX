// Escribiente — fase 2: navegación, brocha, comentarios, control de cambios, notas al pie,
// columnas, cuadros de texto, formas, combinar correspondencia y sangrías en la regla. GPL-3.0
import { h, $, $$, icon, toast, modal, form, dropdown, closeDropdown, esc, api, t } from '../../shared/office.js';
import { roman } from './odt.js';

const uid = () => Math.random().toString(36).slice(2, 10);
let C = null; // contexto que entrega main.js
export function initFeatures(ctx) { C = ctx; buildPanels(); bindTrack(); bindPainter(); bindComments(); bindFootnotes(); }
const author = () => C.shell.prefs?.author || 'Yo';
const nowIso = () => new Date().toISOString();
const fmtDate = (d) => new Date(d).toLocaleString('es-MX', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' });

// ------------------------------------------------------------------ paneles laterales
let navPanel, sidePanel;
function buildPanels() {
  navPanel = h('aside', { class: 'sidepanel left', hidden: true, 'aria-label': 'Navegación' });
  sidePanel = h('aside', { class: 'sidepanel right', hidden: true, 'aria-label': 'Comentarios y revisiones' });
  C.shell.workspace.append(navPanel, sidePanel);
}
function layoutPanels() {
  C.scroller.style.left = navPanel.hidden ? '0' : '250px';
  C.scroller.style.right = sidePanel.hidden ? '0' : '310px';
  dispatchEvent(new Event('resize'));
}

// Panel de navegación por títulos
export function toggleNav(force) {
  navPanel.hidden = force !== undefined ? !force : !navPanel.hidden;
  C.shell.ribbon.setOn('navPane', !navPanel.hidden);
  renderNav(); layoutPanels();
}
export function renderNav() {
  if (!navPanel || navPanel.hidden) return;
  const heads = [...C.editor.querySelectorAll('h1,h2,h3,h4,p.title')];
  navPanel.replaceChildren(h('div', { class: 'sp-head' }, h('b', {}, 'Navegación'), h('button', { class: 'rb sm', title: 'Cerrar', html: icon('close', 16), onclick: () => toggleNav(false) })));
  if (!heads.length) navPanel.append(h('p', { class: 'sp-empty' }, 'Usa los estilos Título 1, 2 o 3 para ver aquí la estructura del documento.'));
  const list = h('div', { class: 'nav-list' });
  heads.forEach((hd) => {
    const lvl = hd.tagName === 'P' ? 0 : Number(hd.tagName[1]) - 1;
    list.append(h('button', { class: `nav-item l${lvl}`, title: hd.textContent, onclick: () => { scrollToNode(hd); placeCaret(hd); } }, hd.textContent || '(sin texto)'));
  });
  navPanel.append(list);
}
function scrollToNode(n) {
  const z = C.state.zoom / 100;
  const top = (n.offsetTop + C.editor.offsetTop) * z;
  C.scroller.scrollTo({ top: Math.max(0, top - 60), behavior: 'smooth' });
}
function placeCaret(n, end = false) {
  C.editor.focus({ preventScroll: true });
  const r = document.createRange(); r.selectNodeContents(n); r.collapse(!end);
  const s = document.getSelection(); s.removeAllRanges(); s.addRange(r); C.saveSelection();
}

// ------------------------------------------------------------------ brocha (copiar formato)
let painter = null;
function bindPainter() {
  C.editor.addEventListener('mouseup', () => {
    if (!painter) return;
    const s = document.getSelection();
    if (s.isCollapsed) return;
    applyPainted(painter.fmt);
    if (!painter.sticky) stopPainter();
  });
  addEventListener('keydown', (e) => { if (e.key === 'Escape' && painter) stopPainter(); });
}
export function startPainter(sticky = false) {
  if (painter && Date.now() - painter.t < 450) { painter.sticky = true; toast('Brocha fija: selecciona textos para copiar el formato. Esc para terminar.'); return; }
  if (painter) { stopPainter(); return; }
  const s = document.getSelection();
  let n = s.anchorNode; if (!n || !C.editor.contains(n)) { toast('Coloca el cursor en el texto que tiene el formato.'); return; }
  if (n.nodeType === 3) n = n.parentElement;
  const cs = getComputedStyle(n);
  const block = n.closest('p,h1,h2,h3,h4,h5,h6,blockquote,li');
  painter = {
    sticky, t: Date.now(),
    fmt: {
      bold: Number(cs.fontWeight) >= 600, italic: cs.fontStyle === 'italic',
      underline: cs.textDecorationLine.includes('underline'), strike: cs.textDecorationLine.includes('line-through'),
      color: cs.color, size: `${Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2}pt`, family: n.closest('[style*="font-family"]')?.style.fontFamily || null,
      bg: n.closest('[style*="background-color"]')?.style.backgroundColor || null,
      align: block ? getComputedStyle(block).textAlign : null,
    },
  };
  C.editor.classList.add('painting');
  C.shell.ribbon.setOn('painter', true);
  toast(sticky ? 'Brocha fija: selecciona textos para copiar el formato. Esc para terminar.' : 'Selecciona el texto al que quieres copiar el formato.');
}
function stopPainter() { painter = null; C.editor.classList.remove('painting'); C.shell.ribbon.setOn('painter', false); }
function applyPainted(f) {
  C.saveSelection();
  const q = (c) => { try { return document.queryCommandState(c); } catch { return false; } };
  document.execCommand('styleWithCSS', false, true);
  for (const [k, cmd] of [['bold', 'bold'], ['italic', 'italic'], ['underline', 'underline'], ['strike', 'strikeThrough']]) if (q(cmd) !== f[k]) document.execCommand(cmd);
  document.execCommand('foreColor', false, f.color);
  if (f.bg) document.execCommand('hiliteColor', false, f.bg);
  if (f.family) document.execCommand('fontName', false, f.family);
  C.saveSelection();
  C.applySpanStyle('fontSize', f.size);
  if (f.align) C.selectedBlocks().forEach((b) => { b.style.textAlign = f.align === 'start' ? '' : f.align; });
  C.shell.markDirty(); C.schedulePaginate();
}

// ------------------------------------------------------------------ comentarios
export function comments() { return C.state.comments; }
function bindComments() {
  C.editor.addEventListener('click', (e) => {
    const cm = e.target.closest('.cm');
    if (cm) { openSide('comments'); focusComment(cm.dataset.cid); }
  });
}
export async function newComment() {
  C.restoreSelection();
  const s = document.getSelection();
  if (!s.rangeCount || s.isCollapsed) { toast('Selecciona el texto que quieres comentar.'); return; }
  const v = await form('Nuevo comentario', [{ name: 't', label: 'Comentario', type: 'textarea', rows: 4 }], 'Comentar');
  if (!v || !v.t.trim()) return;
  C.restoreSelection();
  const id = uid();
  wrapSelection((txt) => { const sp = document.createElement('span'); sp.className = 'cm'; sp.dataset.cid = id; return sp; });
  C.state.comments.push({ id, author: author(), date: nowIso(), text: v.t.trim(), replies: [], resolved: false });
  C.shell.markDirty(); openSide('comments'); focusComment(id);
}
function deleteComment(id) {
  C.editor.querySelectorAll(`.cm[data-cid="${id}"]`).forEach(unwrap);
  C.state.comments = C.state.comments.filter((c) => c.id !== id);
  C.shell.markDirty(); renderSide();
}
function focusComment(id) {
  $$('.cm.on', C.editor).forEach((x) => x.classList.remove('on'));
  C.editor.querySelectorAll(`.cm[data-cid="${id}"]`).forEach((x) => x.classList.add('on'));
  const card = sidePanel.querySelector(`[data-card="${id}"]`);
  card?.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  $$('.cmcard.on', sidePanel).forEach((x) => x.classList.remove('on')); card?.classList.add('on');
}

// ------------------------------------------------------------------ control de cambios
export function setTrack(on) {
  C.state.track = on;
  C.shell.ribbon.setOn('track', on);
  C.editor.classList.toggle('tracking', on);
  toast(on ? 'Control de cambios activado: tus ediciones quedan marcadas' : 'Control de cambios desactivado');
  C.shell.markDirty();
}
function revEl(kind) { const e = document.createElement(kind); e.dataset.author = author(); e.dataset.date = nowIso(); e.title = `${kind === 'ins' ? 'Insertado' : 'Eliminado'} por ${author()}`; return e; }
function textNodesIn(range) {
  const out = [];
  const w = document.createTreeWalker(range.commonAncestorContainer.nodeType === 3 ? range.commonAncestorContainer.parentNode : range.commonAncestorContainer, NodeFilter.SHOW_TEXT);
  while (w.nextNode()) if (range.intersectsNode(w.currentNode) && w.currentNode.nodeValue) out.push(w.currentNode);
  return out;
}
// Envuelve cada nodo de texto dentro del rango con el elemento que devuelve make()
function wrapRange(range, make, skip = () => false) {
  const nodes = textNodesIn(range);
  let last = null;
  for (let n of nodes) {
    if (skip(n)) continue;
    let start = n === range.startContainer ? range.startOffset : 0;
    let end = n === range.endContainer ? range.endOffset : n.nodeValue.length;
    if (end <= start) continue;
    if (end < n.nodeValue.length) n.splitText(end);
    if (start > 0) n = n.splitText(start);
    const w = make(); n.parentNode.insertBefore(w, n); w.append(n); last = w;
  }
  return last;
}
function wrapSelection(make) {
  const s = document.getSelection(); if (!s.rangeCount) return null;
  const r = s.getRangeAt(0);
  const last = wrapRange(r, make);
  s.removeAllRanges();
  return last;
}
function unwrap(el) { const p = el.parentNode; while (el.firstChild) p.insertBefore(el.firstChild, el); el.remove(); }

function bindTrack() {
  C.editor.addEventListener('beforeinput', (e) => {
    if (!C.state.track) return;
    const t = e.inputType;
    const s = document.getSelection();
    if (!s.rangeCount) return;
    const r = s.getRangeAt(0);
    const inOwnIns = (node) => { const el = (node.nodeType === 3 ? node.parentElement : node)?.closest('ins'); return el && el.dataset.author === author() ? el : null; };
    if (t === 'insertText' || t === 'insertReplacementText' || t === 'insertFromPaste' || t === 'insertFromDrop') {
      const text = t === 'insertText' ? e.data : (e.dataTransfer?.getData('text/plain') ?? e.data ?? '');
      if (text == null) return;
      e.preventDefault();
      if (!r.collapsed) markDeleted(r, 'forward');
      const sel = document.getSelection(); const rr = sel.getRangeAt(0);
      const own = inOwnIns(rr.startContainer);
      if (own) { const tn = document.createTextNode(text); rr.insertNode(tn); rr.setStartAfter(tn); rr.collapse(true); }
      else {
        // si el cursor quedó dentro de un <del>, se sale de él
        const del = (rr.startContainer.nodeType === 3 ? rr.startContainer.parentElement : rr.startContainer).closest('del');
        if (del) { rr.setStartAfter(del); rr.collapse(true); }
        const ins = revEl('ins'); ins.textContent = text; rr.insertNode(ins);
        rr.setStart(ins.firstChild, text.length); rr.collapse(true);
      }
      sel.removeAllRanges(); sel.addRange(rr);
      C.shell.markDirty(); C.schedulePaginate(); renderSideSoon();
      return;
    }
    if (t.startsWith('delete')) {
      if (t === 'deleteByDrag') return;
      e.preventDefault();
      const backward = t.includes('Backward');
      if (r.collapsed) {
        s.modify('extend', backward ? 'backward' : 'forward', t.includes('Word') ? 'word' : 'character');
      }
      const rng = s.getRangeAt(0);
      if (rng.collapsed) return;
      markDeleted(rng, backward ? 'backward' : 'forward');
      C.shell.markDirty(); C.schedulePaginate(); renderSideSoon();
    }
  });
}
function markDeleted(range, dir) {
  // Lo que el mismo autor acaba de insertar se borra de verdad; lo demás se marca como eliminado.
  const s = document.getSelection();
  const ownIns = (n) => { const el = n.parentElement?.closest('ins'); return el && el.dataset.author === author(); };
  const nodes = textNodesIn(range);
  const ownNodes = nodes.filter(ownIns);
  const endNode = range.endContainer, endOff = range.endOffset, startNode = range.startContainer, startOff = range.startOffset;
  const last = wrapRange(range, () => revEl('del'), (n) => ownIns(n) || !!n.parentElement?.closest('del'));
  // borrar lo propio
  for (const n of ownNodes) {
    const st = n === startNode ? startOff : 0, en = n === endNode ? endOff : n.nodeValue.length;
    n.nodeValue = n.nodeValue.slice(0, st) + n.nodeValue.slice(en);
    const ins = n.parentElement.closest('ins'); if (ins && !ins.textContent) ins.remove();
  }
  // cursor: antes (hacia atrás) o después (hacia adelante) de lo marcado
  const dels = C.editor.querySelectorAll('del');
  const r2 = document.createRange();
  const anchor = last || dels[dels.length - 1];
  try {
    if (dir === 'backward') { const first = range.startContainer.isConnected ? null : null; void first; if (anchor) { const firstDel = findFirstDelOfRun(anchor); r2.setStartBefore(firstDel); } }
    else if (anchor) r2.setStartAfter(anchor);
    r2.collapse(true); s.removeAllRanges(); s.addRange(r2);
  } catch { /* cursor sin cambio */ }
}
function findFirstDelOfRun(del) { let d = del; while (d.previousSibling && d.previousSibling.nodeName === 'DEL') d = d.previousSibling; return d; }

function revisions() { return [...C.editor.querySelectorAll('ins,del')]; }
export function reviewAction(kind) {
  const revs = revisions();
  if (!revs.length) { toast('No hay cambios para revisar.'); return; }
  const s = document.getSelection();
  const at = s.rangeCount ? s.getRangeAt(0) : null;
  let cur = revs.find((r) => at && r.contains(at.startContainer)) || null;
  if (kind === 'next' || kind === 'prev') {
    const pos = (el) => (at ? at.comparePoint(el, 0) : -1);
    cur = kind === 'next' ? revs.find((r) => pos(r) > 0 && !r.contains(at?.startContainer)) || revs[0] : [...revs].reverse().find((r) => pos(r) < 0 && !r.contains(at?.startContainer)) || revs[revs.length - 1];
    selectNode(cur); scrollToNode(cur.closest('p,h1,h2,h3,li,td') || cur);
    return;
  }
  const apply = (el, accept) => { if ((el.tagName === 'INS') === accept) unwrap(el); else el.remove(); };
  if (kind === 'acceptAll' || kind === 'rejectAll') { revs.forEach((el) => apply(el, kind === 'acceptAll')); toast(`${revs.length} cambios ${kind === 'acceptAll' ? 'aceptados' : 'rechazados'}`); }
  else {
    if (!cur) { reviewAction('next'); toast('Selecciona un cambio y vuelve a pulsar.'); return; }
    const next = revs[revs.indexOf(cur) + 1];
    apply(cur, kind === 'accept');
    if (next?.isConnected) selectNode(next);
  }
  C.shell.markDirty(); C.schedulePaginate(); renderSideSoon();
}
function selectNode(n) { C.editor.focus({ preventScroll: true }); const r = document.createRange(); r.selectNodeContents(n); const s = document.getSelection(); s.removeAllRanges(); s.addRange(r); C.saveSelection(); }

// ------------------------------------------------------------------ panel derecho (comentarios + cambios)
let sideMode = 'comments';
export function openSide(mode) { sideMode = mode || sideMode; sidePanel.hidden = false; renderSide(); layoutPanels(); }
export function toggleSide(mode) { if (!sidePanel.hidden && sideMode === mode) { sidePanel.hidden = true; layoutPanels(); return; } openSide(mode); }
const renderSideSoon = () => { clearTimeout(renderSideSoon.t); renderSideSoon.t = setTimeout(renderSide, 300); };
export function renderSide() {
  if (!sidePanel || sidePanel.hidden) return;
  const tabs = h('div', { class: 'sp-tabs' },
    h('button', { class: sideMode === 'comments' ? 'on' : '', onclick: () => openSide('comments') }, `Comentarios (${C.state.comments.length})`),
    h('button', { class: sideMode === 'changes' ? 'on' : '', onclick: () => openSide('changes') }, `Cambios (${revisions().length})`),
    h('button', { class: 'rb sm', title: 'Cerrar', html: icon('close', 16), onclick: () => { sidePanel.hidden = true; layoutPanels(); } }));
  const body = h('div', { class: 'sp-body' });
  if (sideMode === 'comments') {
    const order = [...C.editor.querySelectorAll('.cm')].map((x) => x.dataset.cid);
    const list = [...C.state.comments].sort((a, b) => order.indexOf(a.id) - order.indexOf(b.id));
    if (!list.length) body.append(h('p', { class: 'sp-empty' }, 'Selecciona texto y pulsa “Nuevo comentario” (Revisar).'));
    list.forEach((c) => {
      const replies = h('div', { class: 'cm-replies' }, ...c.replies.map((r) => h('div', { class: 'cm-reply' }, h('b', {}, r.author), h('small', {}, fmtDate(r.date)), h('p', {}, r.text))));
      const reply = h('input', { placeholder: 'Responder…' });
      reply.addEventListener('keydown', (e) => { if (e.key === 'Enter' && reply.value.trim()) { c.replies.push({ author: author(), date: nowIso(), text: reply.value.trim() }); C.shell.markDirty(); renderSide(); } });
      body.append(h('div', { class: `cmcard${c.resolved ? ' resolved' : ''}`, 'data-card': c.id, onclick: (e) => { if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'BUTTON') { focusComment(c.id); const el = C.editor.querySelector(`.cm[data-cid="${c.id}"]`); if (el) scrollToNode(el.closest('p,h1,h2,h3,li,td') || el); } } },
        h('div', { class: 'cm-top' }, h('span', { class: 'avatar' }, (c.author || '?')[0].toUpperCase()), h('div', {}, h('b', {}, c.author), h('small', {}, fmtDate(c.date)))),
        h('p', { class: 'cm-text' }, c.text), replies,
        c.resolved ? null : reply,
        h('div', { class: 'cm-actions' },
          h('button', { class: 'btn ghost', onclick: () => { c.resolved = !c.resolved; C.shell.markDirty(); renderSide(); } }, c.resolved ? 'Reabrir' : 'Resolver'),
          h('button', { class: 'btn ghost', onclick: () => deleteComment(c.id) }, 'Eliminar'))));
    });
  } else {
    const revs = revisions();
    if (!revs.length) body.append(h('p', { class: 'sp-empty' }, 'No hay cambios marcados. Activa “Control de cambios” para registrar tus ediciones.'));
    body.append(h('div', { class: 'row', style: { gap: '6px', marginBottom: '8px' } }, h('button', { class: 'btn', onclick: () => reviewAction('acceptAll') }, 'Aceptar todos'), h('button', { class: 'btn', onclick: () => reviewAction('rejectAll') }, 'Rechazar todos')));
    revs.slice(0, 300).forEach((el) => {
      body.append(h('div', { class: `revcard ${el.tagName.toLowerCase()}`, onclick: () => { selectNode(el); scrollToNode(el.closest('p,h1,h2,h3,li,td') || el); } },
        h('div', {}, h('b', {}, el.tagName === 'INS' ? 'Insertó' : 'Eliminó'), ' · ', el.dataset.author || '', h('small', {}, el.dataset.date ? fmtDate(el.dataset.date) : '')),
        h('p', {}, `“${el.textContent.slice(0, 120)}”`),
        h('div', { class: 'cm-actions' }, h('button', { class: 'btn ghost', onclick: (e) => { e.stopPropagation(); if (el.tagName === 'INS') unwrap(el); else el.remove(); C.shell.markDirty(); C.schedulePaginate(); renderSide(); } }, 'Aceptar'), h('button', { class: 'btn ghost', onclick: (e) => { e.stopPropagation(); if (el.tagName === 'INS') el.remove(); else unwrap(el); C.shell.markDirty(); C.schedulePaginate(); renderSide(); } }, 'Rechazar'))));
    });
  }
  sidePanel.replaceChildren(tabs, body);
}

// ------------------------------------------------------------------ notas al pie
function bindFootnotes() {
  C.editor.addEventListener('click', (e) => { const fn = e.target.closest('sup.fn'); if (fn) editFootnote(fn.dataset.fid); const en = e.target.closest('sup.en'); if (en) editEndnote(en.dataset.fid); });
}
// ------------------------------------------------------------------ notas al final
export async function insertEndnote() {
  C.restoreSelection();
  const v = await form('Insertar nota al final', [{ name: 't', label: 'Texto de la nota', type: 'textarea', rows: 3 }], 'Insertar');
  if (!v || !v.t.trim()) return;
  const fid = `e${uid()}`;
  C.state.endnotes[fid] = v.t.trim();
  C.insertHTMLAtCursor(`<sup class="en" data-fid="${fid}" contenteditable="false">#</sup>`);
  numberFootnotes(); C.shell.markDirty(); C.schedulePaginate();
}
async function editEndnote(fid) {
  const v = await form('Nota al final', [{ name: 't', label: 'Texto de la nota (déjalo vacío para eliminarla)', type: 'textarea', rows: 3, value: C.state.endnotes[fid] || '' }], 'Guardar');
  if (!v) return;
  if (!v.t.trim()) { C.editor.querySelectorAll(`sup.en[data-fid="${fid}"]`).forEach((x) => x.remove()); delete C.state.endnotes[fid]; }
  else C.state.endnotes[fid] = v.t.trim();
  numberFootnotes(); C.shell.markDirty(); C.schedulePaginate();
}
export function endnoteOrder() {
  return [...C.editor.querySelectorAll('sup.en')].map((x) => x.dataset.fid).filter((v, k, a) => a.indexOf(v) === k && C.state.endnotes[v] != null);
}
export function endnotesEl(style) {
  const order = endnoteOrder();
  return h('div', { class: 'enotes doc', style }, h('div', { class: 'enhead' }, t('Notas al final')), ...order.map((f, i) => h('div', { class: 'fnote' }, h('sup', {}, roman(i + 1)), ` ${C.state.endnotes[f]}`)));
}
export function endnotesHeight(width) {
  if (!meas) { meas = h('div', { class: 'doc fnmeasure' }); document.body.append(meas); }
  meas.style.width = `${width}px`;
  meas.replaceChildren(endnotesEl({ position: 'static', width: `${width}px` }));
  return meas.firstChild.offsetHeight;
}
export async function insertFootnote() {
  C.restoreSelection();
  const v = await form('Insertar nota al pie', [{ name: 't', label: 'Texto de la nota', type: 'textarea', rows: 3 }], 'Insertar');
  if (!v || !v.t.trim()) return;
  const fid = uid();
  C.state.footnotes[fid] = v.t.trim();
  C.insertHTMLAtCursor(`<sup class="fn" data-fid="${fid}" contenteditable="false">#</sup>`);
  numberFootnotes();
}
async function editFootnote(fid) {
  const v = await form('Nota al pie', [{ name: 't', label: 'Texto de la nota (déjalo vacío para eliminarla)', type: 'textarea', rows: 3, value: C.state.footnotes[fid] || '' }], 'Guardar');
  if (!v) return;
  if (!v.t.trim()) { C.editor.querySelectorAll(`sup.fn[data-fid="${fid}"]`).forEach((x) => x.remove()); delete C.state.footnotes[fid]; }
  else C.state.footnotes[fid] = v.t.trim();
  numberFootnotes(); C.shell.markDirty(); C.schedulePaginate();
}
export function numberFootnotes() {
  const sups = [...C.editor.querySelectorAll('sup.fn')];
  const order = [];
  sups.forEach((s) => { if (!order.includes(s.dataset.fid)) order.push(s.dataset.fid); s.textContent = String(order.indexOf(s.dataset.fid) + 1); });
  const eo = [];
  C.editor.querySelectorAll('sup.en').forEach((s) => { if (!eo.includes(s.dataset.fid)) eo.push(s.dataset.fid); s.textContent = roman(eo.indexOf(s.dataset.fid) + 1); });
  return order;
}
// Mide el alto de cada nota (en px) con el ancho del texto
let meas;
export function footnoteHeights(width) {
  if (!meas) { meas = h('div', { class: 'doc fnmeasure' }); document.body.append(meas); }
  meas.style.width = `${width}px`;
  const out = {};
  for (const [fid, text] of Object.entries(C.state.footnotes)) { meas.innerHTML = `<div class="fnote"><sup>0</sup> ${esc(text)}</div>`; out[fid] = meas.firstChild.offsetHeight; }
  return out;
}

// ------------------------------------------------------------------ columnas
export function setColumns(n) {
  C.restoreSelection();
  const blocks = C.selectedBlocks().map((b) => topBlock(b)).filter(Boolean);
  const uniq = [...new Set(blocks)];
  const inCols = uniq.find((b) => b.classList.contains('cols'));
  if (inCols) {
    if (n === 1) { unwrapBlock(inCols); } else { inCols.dataset.n = n; inCols.style.columnCount = n; }
  } else if (n > 1) {
    if (!uniq.length) return;
    const wrap = h('div', { class: 'cols', 'data-n': n, style: { columnCount: n } });
    uniq[0].before(wrap);
    uniq.forEach((b) => wrap.append(b));
  }
  C.shell.markDirty(); C.schedulePaginate();
}
function topBlock(n) { while (n && n.parentElement !== C.editor) n = n.parentElement; return n; }
function unwrapBlock(el) { while (el.firstChild) el.before(el.firstChild); el.remove(); }

// ------------------------------------------------------------------ cuadro de texto y formas
export function insertTextBox() {
  C.insertHTMLAtCursor('<div class="tbox" style="border-color:#5B48E8"><p>Escribe aquí el texto del cuadro.</p></div><p><br></p>');
}
const SHAPES = {
  rect: (c) => `<rect x="4" y="4" width="192" height="112" rx="2" fill="${c}"/>`,
  round: (c) => `<rect x="4" y="4" width="192" height="112" rx="22" fill="${c}"/>`,
  ellipse: (c) => `<ellipse cx="100" cy="60" rx="96" ry="56" fill="${c}"/>`,
  arrow: (c) => `<path d="M4 40 H130 V8 L196 60 L130 112 V80 H4 Z" fill="${c}"/>`,
  star: (c) => `<path d="M100 4 L124 44 L170 48 L136 78 L146 116 L100 96 L54 116 L64 78 L30 48 L76 44 Z" fill="${c}"/>`,
  triangle: (c) => `<path d="M100 6 L196 114 H4 Z" fill="${c}"/>`,
};
export function shapeMenu() {
  const names = { rect: 'Rectángulo', round: 'Rectángulo redondeado', ellipse: 'Elipse', arrow: 'Flecha', star: 'Estrella', triangle: 'Triángulo' };
  return Object.keys(SHAPES).map((k) => ({ label: names[k], icon: { round: 'rect', rect: 'rect' }[k] || k, run: () => insertShape(k) }));
}
function insertShape(kind, color = '#5B48E8') {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="200" height="120" viewBox="0 0 200 120">${SHAPES[kind](color)}</svg>`;
  C.insertHTMLAtCursor(`<img src="data:image/svg+xml;base64,${btoa(svg)}" width="200" alt="Forma">`);
}

// ------------------------------------------------------------------ combinar correspondencia
export async function chooseRecipients() {
  if (!api) return;
  const f = await api.openDialog({ title: 'Elegir destinatarios', filters: [{ name: 'Lista de datos', extensions: ['csv', 'xlsx', 'tsv'] }] });
  if (!f || f.redirected) return;
  let rows;
  if (f.ext === 'xlsx') {
    const ExcelJS = (await import('exceljs')).default;
    const x = new ExcelJS.Workbook(); await x.xlsx.load(f.bytes.buffer.slice(f.bytes.byteOffset, f.bytes.byteOffset + f.bytes.byteLength));
    const ws = x.worksheets[0]; rows = [];
    ws.eachRow({ includeEmpty: false }, (row) => { rows.push(row.values.slice(1).map((v) => (v && typeof v === 'object' ? (v.text ?? v.result ?? (v instanceof Date ? v.toLocaleDateString('es-MX') : '')) : v ?? '')).map(String)); });
  } else {
    const { parseCsv } = await import('../../calcula/src/csv.js');
    rows = parseCsv(new TextDecoder('utf-8').decode(f.bytes));
  }
  rows = rows.filter((r) => r.some((c) => String(c).trim()));
  if (rows.length < 2) { toast('La lista necesita una fila de encabezados y al menos un registro.', { error: true }); return; }
  const fields = rows[0].map((x, i) => String(x).trim() || `Campo${i + 1}`);
  C.state.merge = { name: f.name, fields, rows: rows.slice(1), preview: false, idx: 0 };
  toast(`${rows.length - 1} destinatarios de ${f.name}`);
  C.shell.ribbon.renderPanel();
}
export function fieldMenu() {
  if (!C.state.merge) return [{ label: 'Primero elige los destinatarios', icon: 'open', run: chooseRecipients }];
  return C.state.merge.fields.map((fl) => ({ label: fl, run: () => C.insertHTMLAtCursor(`<span class="mf" contenteditable="false" data-f="${esc(fl)}">«${esc(fl)}»</span>&nbsp;`) }));
}
export function previewMerge(step = 0) {
  const m = C.state.merge; if (!m) { toast('Primero elige los destinatarios.'); return; }
  if (step === 0) m.preview = !m.preview; else { m.preview = true; m.idx = Math.max(0, Math.min(m.rows.length - 1, m.idx + step)); }
  C.editor.querySelectorAll('.mf').forEach((el) => { const i = m.fields.indexOf(el.dataset.f); el.textContent = m.preview ? (m.rows[m.idx][i] ?? '') : `«${el.dataset.f}»`; el.classList.toggle('pv', m.preview); });
  C.shell.ribbon.setOn('mergePreview', m.preview);
  if (m.preview) toast(`Registro ${m.idx + 1} de ${m.rows.length}`);
  C.schedulePaginate();
}
export function mergedHtml() {
  const m = C.state.merge;
  const src = C.editor.cloneNode(true);
  src.querySelectorAll('[data-push]').forEach((b) => { b.style.marginTop = ''; b.removeAttribute('data-push'); });
  const parts = m.rows.map((row) => {
    const c = src.cloneNode(true);
    c.querySelectorAll('.mf').forEach((el) => { const i = m.fields.indexOf(el.dataset.f); el.replaceWith(document.createTextNode(row[i] ?? '')); });
    return c.innerHTML;
  });
  return parts.join('<div class="pb" contenteditable="false"></div>');
}
export async function finishMerge() {
  const m = C.state.merge; if (!m) { toast('Primero elige los destinatarios.'); return; }
  if (!C.editor.querySelector('.mf')) { toast('Inserta al menos un campo de combinación en el documento.'); return; }
  const v = await form('Finalizar y combinar', [{ name: 'out', label: `Combinar ${m.rows.length} registros en:`, type: 'select', value: 'docx', options: [['docx', 'Un documento de Word (.docx)'], ['pdf', 'Un PDF'], ['here', 'Este documento (reemplaza el contenido)']] }], 'Combinar');
  if (!v) return;
  if (v.out === 'here') { if (m.preview) previewMerge(0); const html = mergedHtml(); C.setContent(html); C.state.merge = null; C.shell.markDirty(); toast('Documento combinado'); return; }
  if (m.preview) previewMerge(0);
  const backup = C.editor.innerHTML;
  const html = mergedHtml();
  if (v.out === 'docx') {
    const div = h('div', { html });
    const bytes = await C.toDocx(div);
    const res = await api.save({ bytes, ask: true, ext: 'docx', suggestedName: `${C.shell.doc.name.replace(/\.[^.]+$/, '')} (combinado).docx` });
    if (res) toast(`Listo: ${res.name}`, { action: { label: 'Abrir', run: () => api.openInSuite(res.path) } });
  } else {
    C.editor.innerHTML = html; C.paginate();
    await C.shell.exportPdf();
    C.editor.innerHTML = backup; C.schedulePaginate();
  }
}

// ------------------------------------------------------------------ sangrías arrastrables en la regla
export function drawIndentMarkers(left, z, cw) {
  const b = C.currentBlock();
  if (!b) return;
  const ml = parseFloat(b.style.marginLeft) || 0, ti = parseFloat(b.style.textIndent) || 0;
  const mk = (cls, x, title, onDrag) => {
    const m = h('div', { class: `indent ${cls}`, title, style: { left: `${x}px` } });
    m.addEventListener('mousedown', (e) => {
      e.preventDefault();
      const x0 = e.clientX; const blocks = C.selectedBlocks();
      const start = blocks.map((bl) => ({ bl, ml: parseFloat(bl.style.marginLeft) || 0, ti: parseFloat(bl.style.textIndent) || 0 }));
      const move = (ev) => { const d = Math.round(((ev.clientX - x0) / z) / 6) * 6; start.forEach((s) => onDrag(s, d)); m.style.left = `${x + d * z}px`; };
      const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); C.shell.markDirty(); C.schedulePaginate(); };
      addEventListener('mousemove', move); addEventListener('mouseup', up);
    });
    C.ruler.append(m);
  };
  mk('first', left + (ml + ti) * z, 'Sangría de primera línea', (s, d) => { s.bl.style.textIndent = s.ti + d ? `${s.ti + d}px` : ''; });
  mk('left', left + ml * z, 'Sangría izquierda', (s, d) => { const v = Math.max(0, Math.min(cw - 40, s.ml + d)); s.bl.style.marginLeft = v ? `${v}px` : ''; });
}

// ------------------------------------------------------------------ tabulaciones
// Cada tabulación es <span class="otab">\t</span>; su ancho se calcula para llegar a la
// siguiente parada del párrafo (data-tabs="l:120,c:300,r:500", en px desde el margen izquierdo)
// o a la siguiente parada predeterminada (cada 1.25 cm, como Word).
const DEF_TAB = 48;
let tabKind = 'l';
export function parseTabs(v) {
  return String(v || '').split(',').map((t) => { const m = /^([lcr]):(\d+(?:\.\d+)?)$/.exec(t.trim()); return m ? { k: m[1], x: Number(m[2]) } : null; }).filter(Boolean).sort((a, b) => a.x - b.x);
}
const tabsStr = (list) => list.sort((a, b) => a.x - b.x).map((t) => `${t.k}:${Math.round(t.x)}`).join(',');
const TAB_BLOCK = 'p,h1,h2,h3,h4,h5,h6,li,blockquote,td,th';
export function insertTab() {
  if (C.state.track) { document.execCommand('insertText', false, '\t'); return; }
  document.execCommand('insertHTML', false, '<span class="otab" contenteditable="false">\t</span>\u200B');
}
export function layoutTabs() {
  // tabuladores sueltos (pegados o importados) -> span.otab
  const w = document.createTreeWalker(C.editor, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.nodeValue.includes('\t') && !n.parentElement.closest('.otab,pre,del') ? 1 : 3) });
  const raw = []; while (w.nextNode()) raw.push(w.currentNode);
  for (const t of raw) {
    const frag = document.createDocumentFragment();
    t.nodeValue.split('\t').forEach((part, i) => { if (i) frag.append(h('span', { class: 'otab', contenteditable: 'false' }, '\t')); if (part) frag.append(document.createTextNode(part)); });
    t.replaceWith(frag);
  }
  const spans = [...C.editor.querySelectorAll('.otab')];
  if (!spans.length) return;
  const z = C.state.zoom / 100;
  const edLeft = C.editor.getBoundingClientRect().left;
  spans.forEach((s) => { s.style.width = '1px'; });
  for (const s of spans) {
    const blk = s.parentElement.closest(TAB_BLOCK) || C.editor;
    const stops = parseTabs(blk.dataset?.tabs);
    const bl = blk === C.editor ? 0 : (blk.getBoundingClientRect().left - edLeft) / z;
    const x = (s.getBoundingClientRect().left - edLeft) / z;
    const seg = () => {
      const r = document.createRange(); r.setStartAfter(s);
      let end = s.nextSibling; while (end && !(end.nodeType === 1 && end.classList?.contains('otab'))) end = end.nextSibling;
      if (end) r.setEndBefore(end); else r.setEnd(s.parentElement, s.parentElement.childNodes.length);
      return r.getBoundingClientRect().width / z;
    };
    let wdt = 0;
    for (const st of stops) {
      if (st.x <= x + 1) continue;
      if (st.k === 'l') wdt = st.x - x;
      else { const sw = seg(); wdt = st.k === 'r' ? st.x - x - sw : st.x - x - sw / 2; }
      if (wdt > 1) break; wdt = 0;
    }
    if (!wdt) { const last = stops.length ? Math.max(stops[stops.length - 1].x, 0) : 0; const base = Math.max(last, bl); let nx = Math.floor(x / DEF_TAB + 1) * DEF_TAB; if (nx <= base) nx = Math.ceil((base + 1) / DEF_TAB) * DEF_TAB; wdt = Math.max(4, nx - x); }
    s.style.width = `${wdt}px`;
  }
}
export function drawTabMarkers(left, z, cw) {
  const sel = h('button', { class: 'tabkind', title: 'Tipo de tabulación (clic para cambiar): izquierda, centrada o derecha', onclick: () => { tabKind = { l: 'c', c: 'r', r: 'l' }[tabKind]; sel.dataset.k = tabKind; } });
  sel.dataset.k = tabKind;
  C.ruler.append(sel);
  const b = C.currentBlock();
  const blocks = () => C.selectedBlocks().map((x) => x.closest(TAB_BLOCK) || x);
  const track = C.ruler.querySelector('.track');
  if (track && !track.dataset.tabs) {
    track.dataset.tabs = '1';
    track.title = 'Haz clic para poner una tabulación';
    track.addEventListener('mousedown', (e) => {
      if (e.target !== track && !e.target.classList.contains('m')) return;
      const x = (e.clientX - C.ruler.getBoundingClientRect().left - left) / z;
      if (x <= 0 || x >= cw) return;
      e.preventDefault();
      for (const bl of blocks()) { const list = parseTabs(bl.dataset.tabs); list.push({ k: tabKind, x }); bl.dataset.tabs = tabsStr(list); }
      C.shell.markDirty(); C.schedulePaginate();
    });
  }
  if (!b) return;
  const blk = b.closest(TAB_BLOCK) || b;
  for (const t of parseTabs(blk.dataset.tabs)) {
    const m = h('div', { class: `tabstop k-${t.k}`, title: 'Arrastra para mover; arrástrala fuera de la regla para quitarla', style: { left: `${left + t.x * z}px` } });
    m.addEventListener('mousedown', (e) => {
      e.preventDefault(); e.stopPropagation();
      const x0 = e.clientX, y0 = e.clientY; const targets = blocks(); let nx = t.x, gone = false;
      const move = (ev) => { nx = Math.max(0, Math.min(cw, t.x + (ev.clientX - x0) / z)); gone = Math.abs(ev.clientY - y0) > 24; m.style.left = `${left + nx * z}px`; m.style.opacity = gone ? '0.3' : ''; };
      const up = () => {
        removeEventListener('mousemove', move); removeEventListener('mouseup', up);
        for (const bl of targets) { const list = parseTabs(bl.dataset.tabs).filter((s) => !(s.k === t.k && Math.abs(s.x - t.x) < 1)); if (!gone) list.push({ k: t.k, x: nx }); if (list.length) bl.dataset.tabs = tabsStr(list); else delete bl.dataset.tabs; }
        C.shell.markDirty(); C.schedulePaginate();
      };
      addEventListener('mousemove', move); addEventListener('mouseup', up);
    });
    C.ruler.append(m);
  }
}
