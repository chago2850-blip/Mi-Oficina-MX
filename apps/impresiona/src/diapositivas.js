// Impresiona — parte 2 (como LibreOffice Impress): presentaciones personalizadas, configuración de la
// presentación (intervalo, en bucle con pausa, avance manual), vista Esquema, diapositiva de resumen,
// expandir diapositiva, insertar diapositivas desde un archivo, duplicar objetos con desplazamiento
// (Mayús+F3), campos en el texto (fecha, hora, número, total, título, archivo, autor) e imprimir
// diapositivas, notas, documentos (folletos) o esquema. GPL-3.0
import { h, t, tf, toast, esc, api, modal } from '../../shared/office.js';
import { tabbed, field, numIn, sel, chk } from '../../shared/dialogs.js';

let C = null; // { getPres, setPres, cur, setCur, slide, commit, renderAll, renderSlide, uid, newSlide, stopEditing, selObjs, select, present, readPptx, readOdp, docName, docProps }
export function init(ctx) { C = ctx; }
const P = () => C.getPres();
const plain = (html) => { const d = document.createElement('div'); d.innerHTML = html || ''; return (d.innerText || d.textContent || '').trim(); };
const titleOf = (s) => plain(s.objects.find((o) => o.role === 'title')?.html || '');
const v = (body, n) => { const el = body.querySelector(`[name="${n}"]`); return el.type === 'checkbox' ? el.checked : el.value; };

// ================================================================ presentaciones personalizadas
export async function customShowsDialog() {
  P().customShows ||= [];
  const list = h('select', { size: 6, style: { width: '100%' }, 'aria-label': t('Presentaciones personalizadas') });
  const fill = () => list.replaceChildren(...(P().customShows || []).map((cs, i) => h('option', { value: String(i) }, `${cs.name} (${cs.ids.length})`)));
  fill();
  const pick = () => (P().customShows || [])[Number(list.value)];
  const btn = (label, fn) => h('button', { type: 'button', class: 'btn', onclick: fn }, t(label));
  const body = h('div', { class: 'cshow' }, list, h('div', { class: 'fx-col' },
    btn('Nueva…', async () => { const cs = await editCustomShow({ name: tf('Presentación {0}', P().customShows.length + 1), ids: [] }); if (cs) { C.commit(t('Presentación personalizada'), () => { (P().customShows ||= []).push(cs); }); fill(); } }),
    btn('Editar…', async () => { const i = Number(list.value); if (!pick()) return; const cs = await editCustomShow(pick()); if (cs) { C.commit(t('Presentación personalizada'), () => { P().customShows[i] = cs; }); fill(); } }),
    btn('Copiar', () => { const cs = pick(); if (!cs) return; C.commit(t('Presentación personalizada'), () => P().customShows.push({ name: tf('Copia de {0}', cs.name), ids: [...cs.ids] })); fill(); }),
    btn('Eliminar', () => { const i = Number(list.value); if (!pick()) return; C.commit(t('Presentación personalizada'), () => P().customShows.splice(i, 1)); fill(); })));
  const r = await modal({ title: t('Presentaciones personalizadas'), body, width: '520px', buttons: [{ label: t('Iniciar'), value: () => pick()?.name || null }, { label: t('Cerrar'), value: true, primary: true }] });
  if (typeof r === 'string') presentCustom(r);
}
export async function editCustomShow(cs) {
  const pres = P();
  const order = [...cs.ids.filter((id) => pres.slides.some((s) => s.id === id)), ...pres.slides.filter((s) => !cs.ids.includes(s.id)).map((s) => s.id)];
  const rows = h('div', { class: 'cs-rows' });
  const draw = () => rows.replaceChildren(...order.map((id, k) => {
    const s = pres.slides.find((x) => x.id === id); const n = pres.slides.indexOf(s) + 1;
    return h('div', { class: 'cs-row' }, h('label', { class: 'fx-check' }, h('input', { type: 'checkbox', 'data-id': id, checked: cs.ids.includes(id) ? true : null }), ` ${n}. ${titleOf(s) || t('(sin título)')}`),
      h('button', { type: 'button', class: 'btn sm', 'aria-label': t('Subir'), onclick: () => { if (k > 0) { sync(); [order[k - 1], order[k]] = [order[k], order[k - 1]]; draw(); } } }, '↑'),
      h('button', { type: 'button', class: 'btn sm', 'aria-label': t('Bajar'), onclick: () => { if (k < order.length - 1) { sync(); [order[k + 1], order[k]] = [order[k], order[k + 1]]; draw(); } } }, '↓'));
  }));
  const sync = () => { cs = { ...cs, ids: order.filter((id) => rows.querySelector(`input[data-id="${id}"]`)?.checked) }; };
  draw();
  const body = await tabbed('Definir presentación personalizada', [{ label: 'Diapositivas', body: h('div', { class: 'fx-col' }, field('Nombre', h('input', { name: 'name', value: cs.name })), h('p', { class: 'fx-note' }, t('Marca las diapositivas y ordénalas con las flechas.')), rows) }], { width: '520px' });
  if (!body) return null;
  sync();
  return { name: v(body, 'name').trim() || cs.name, ids: cs.ids };
}
export function presentCustom(name) {
  const cs = (P().customShows || []).find((x) => x.name === name); if (!cs) return;
  C.present(0, { ids: cs.ids });
}

// ================================================================ configuración de la presentación
export async function showSettingsDialog() {
  const pres = P(); const ss = pres.showSettings || {};
  const slides = pres.slides.map((s, i) => [String(i), `${i + 1}. ${titleOf(s).slice(0, 40) || t('(sin título)')}`]);
  const shows = (pres.customShows || []).map((c) => [c.name, c.name]);
  const body = await tabbed('Configuración de la presentación', [
    { label: 'Intervalo', body: h('div', { class: 'fx-grid' },
      field('Diapositivas', sel('range', ss.range || 'all', [['all', 'Todas las diapositivas'], ['from', 'Desde'], ...(shows.length ? [['custom', 'Presentación personalizada']] : [])])),
      field('Desde la diapositiva', sel('from', String(ss.from || 0), slides)),
      ...(shows.length ? [field('Presentación personalizada', sel('custom', ss.custom || shows[0][0], shows))] : [])) },
    { label: 'Modo y opciones', body: h('div', { class: 'fx-grid' },
      chk('loop', 'Repetir en bucle (como un quiosco) y volver a empezar tras la pausa', !!ss.loop), field('Pausa entre repeticiones', numIn('pause', ss.pause ?? 5, { step: 1, min: 0, max: 600, unit: 's' })),
      chk('manual', 'Cambiar de diapositiva manualmente (sin intervalos)', !!ss.manual), chk('clickBg', 'Cambiar de diapositiva al hacer clic en el fondo', ss.clickBg !== false),
      chk('pen', 'Empezar con el puntero como pluma', !!ss.pen), chk('laser', 'Empezar con el puntero láser', !!ss.laser)) },
  ], { width: '560px' });
  if (!body) return null;
  const next = { range: v(body, 'range'), from: Number(v(body, 'from')) || 0, custom: shows.length ? v(body, 'custom') : undefined, loop: v(body, 'loop'), pause: Number(v(body, 'pause')) || 0, manual: v(body, 'manual'), clickBg: v(body, 'clickBg'), pen: v(body, 'pen'), laser: v(body, 'laser') };
  C.commit(t('Configuración de la presentación'), () => { P().showSettings = next; });
  return next;
}
// diapositivas que muestra la presentación según la configuración (sin las ocultas)
export function showSlides(opts = {}) {
  const pres = P(); const ss = pres.showSettings || {};
  if (opts.ids) return opts.ids.map((id) => pres.slides.find((s) => s.id === id)).filter(Boolean);
  if (ss.range === 'custom' && ss.custom) { const cs = (pres.customShows || []).find((c) => c.name === ss.custom); if (cs) return cs.ids.map((id) => pres.slides.find((s) => s.id === id)).filter(Boolean); }
  const vis = pres.slides.filter((s) => !s.hidden);
  if (ss.range === 'from') return vis.filter((s) => pres.slides.indexOf(s) >= (ss.from || 0));
  return vis;
}

// ================================================================ vista Esquema
const bodyLines = (s) => {
  const out = []; const bodies = s.objects.filter((o) => o.role === 'body' && o.html); if (!bodies.length) return out;
  const d = document.createElement('div'); d.innerHTML = bodies.map((b) => b.html).join('');
  const walk = (el, depth) => {
    for (const n of el.children) {
      if (n.tagName === 'UL' || n.tagName === 'OL') walk(n, depth + (el.tagName === 'LI' || el === d ? 0 : 0) + (n.parentElement === d ? 0 : 1));
      else if (n.tagName === 'LI') { const own = [...n.childNodes].filter((k) => !(k.nodeType === 1 && /^(UL|OL)$/.test(k.tagName))).map((k) => k.textContent).join('').trim(); if (own) out.push({ depth, text: own }); walk(n, depth); }
      else if (/^(P|DIV|H\d)$/.test(n.tagName)) { if (n.textContent.trim()) out.push({ depth, text: n.textContent.trim() }); }
    }
  };
  walk(d, 0);
  if (!out.length) bodies.forEach((b) => plain(b.html).split('\n').filter(Boolean).forEach((x) => out.push({ depth: 0, text: x.trim() })));
  return out;
};
export function outlineText() {
  return P().slides.map((s) => [titleOf(s) || '', ...bodyLines(s).map((l) => `${'\t'.repeat(l.depth + 1)}${l.text}`)].join('\n')).join('\n');
}
function normalizeList(lines) {
  // construye <ul> anidadas a partir de la profundidad de cada línea
  let html = ''; const stack = [];
  lines.forEach((l, i) => {
    const d = Math.min(l.depth, stack.length);
    while (stack.length > d + 1) { html += '</li></ul>'; stack.pop(); }
    if (stack.length === d + 1) html += '</li>';
    if (stack.length < d + 1) { html += '<ul>'; stack.push(d); }
    html += `<li>${esc(l.text)}`;
    void i;
  });
  while (stack.length) { html += '</li></ul>'; stack.pop(); }
  return html;
}
export function applyOutline(text) {
  const entries = []; let curE = null;
  for (const raw of String(text).replace(/\r/g, '').split('\n')) {
    if (!raw.trim()) continue;
    const m = /^(\t+| {2,})?(.*)$/.exec(raw); const ind = m[1] ? (m[1].startsWith('\t') ? m[1].length : Math.floor(m[1].length / 2)) : 0;
    if (ind === 0 || !curE) { curE = { title: m[2].trim(), lines: [] }; entries.push(curE); if (ind > 0) curE.lines.push({ depth: ind - 1, text: m[2].trim() }); }
    else curE.lines.push({ depth: ind - 1, text: m[2].trim() });
  }
  C.commit(t('Esquema'), () => {
    const pres = P();
    entries.forEach((e, i) => {
      let s = pres.slides[i];
      if (!s) { s = C.newSlide('content'); pres.slides.push(s); }
      let tt = s.objects.find((o) => o.role === 'title');
      if (!tt) { const fresh = C.newSlide('titleOnly').objects.find((o) => o.role === 'title'); if (fresh) { s.objects.unshift(fresh); tt = fresh; } }
      if (tt && plain(tt.html) !== e.title) tt.html = esc(e.title);
      const old = bodyLines(s);
      if (JSON.stringify(old) !== JSON.stringify(e.lines)) {
        let b = s.objects.find((o) => o.role === 'body');
        if (!b && e.lines.length) { b = C.newSlide('content').objects.find((o) => o.role === 'body'); if (b) s.objects.push(b); }
        if (b) b.html = e.lines.length ? normalizeList(e.lines) : '';
        s.objects.filter((o) => o.role === 'body' && o !== b).forEach((o) => { o.html = ''; }); // todo queda en el primer cuadro
      }
    });
  });
  return entries.length;
}
export async function outlineView() {
  C.stopEditing();
  const ta = h('textarea', { class: 'outline-ta', name: 'outline', spellcheck: 'true', 'aria-label': t('Esquema') }, outlineText());
  // Tab y Mayús+Tab bajan o suben el nivel de la línea, como en la vista Esquema de LibreOffice
  ta.addEventListener('keydown', (e) => {
    if (e.key !== 'Tab') return; e.preventDefault();
    const val = ta.value; const a = val.lastIndexOf('\n', ta.selectionStart - 1) + 1; const pos = ta.selectionStart;
    if (e.shiftKey) { if (val[a] === '\t') { ta.value = val.slice(0, a) + val.slice(a + 1); ta.selectionStart = ta.selectionEnd = Math.max(a, pos - 1); } }
    else { ta.value = `${val.slice(0, a)}\t${val.slice(a)}`; ta.selectionStart = ta.selectionEnd = pos + 1; }
  });
  const body = await tabbed('Esquema', [{ label: 'Esquema', body: h('div', { class: 'fx-col' }, h('p', { class: 'fx-note' }, t('Cada línea sin sangría es el título de una diapositiva; con Tab, sus puntos (Mayús+Tab sube de nivel). Las líneas nuevas crean diapositivas.')), ta) }], { width: '760px' });
  if (!body) return null;
  return applyOutline(ta.value);
}

// ================================================================ resumen y expandir
export function summarySlide() {
  C.stopEditing();
  const pres = P(); const titles = pres.slides.filter((s) => !s.hidden).map(titleOf).filter(Boolean);
  if (!titles.length) { toast('Las diapositivas no tienen títulos.', { error: true }); return null; }
  const s = C.newSlide('content');
  s.objects.find((o) => o.role === 'title').html = esc(t('Resumen'));
  s.objects.find((o) => o.role === 'body').html = normalizeList(titles.map((x) => ({ depth: 0, text: x })));
  const at = C.cur() + 1;
  C.commit(t('Diapositiva de resumen'), () => { P().slides.splice(at, 0, s); C.setCur(at); });
  return titles.length;
}
export function expandSlide() {
  C.stopEditing();
  const pres = P(); const s0 = pres.slides[C.cur()]; const lines = bodyLines(s0);
  const groups = []; for (const l of lines) { if (l.depth === 0) groups.push({ title: l.text, lines: [] }); else if (groups.length) groups[groups.length - 1].lines.push({ depth: l.depth - 1, text: l.text }); }
  if (!groups.length) { toast('La diapositiva no tiene puntos para expandir.', { error: true }); return null; }
  const fresh = groups.map((g) => { const s = C.newSlide(g.lines.length ? 'content' : 'titleOnly'); s.objects.find((o) => o.role === 'title').html = esc(g.title); const b = s.objects.find((o) => o.role === 'body'); if (b) b.html = normalizeList(g.lines); return s; });
  const at = C.cur() + 1;
  C.commit(t('Expandir diapositiva'), () => { P().slides.splice(at, 0, ...fresh); C.setCur(at); });
  return fresh.length;
}

// ================================================================ insertar diapositivas desde un archivo
export async function insertFromFileDialog() {
  const f = await api?.openDialog({ title: t('Insertar diapositivas desde un archivo'), filters: [{ name: t('Presentaciones'), extensions: ['pptx', 'odp'] }] });
  if (!f?.bytes) return null;
  return insertFromFile(f, true);
}
export async function insertFromFile(f, ask = false) {
  const r = f.ext === 'odp' ? await C.readOdp(f.bytes, C.uid) : await C.readPptx(f.bytes, C.uid);
  let slides = r.pres.slides;
  if (ask) {
    const body = await tabbed('Insertar diapositivas', [{ label: 'Diapositivas', body: h('div', { class: 'fx-col cs-rows' }, h('p', { class: 'fx-note' }, tf('{0} diapositivas en «{1}»', slides.length, f.name)), ...slides.map((s, i) => chk(`s${i}`, `${i + 1}. ${titleOf(s) || t('(sin título)')}`, true))) }], { width: '480px' });
    if (!body) return null;
    slides = slides.filter((_, i) => v(body, `s${i}`));
  }
  if (!slides.length) return 0;
  // ids nuevos para no chocar con los de esta presentación
  const copy = JSON.parse(JSON.stringify(slides)).map((s) => ({ ...s, id: C.uid(), objects: s.objects.map((o) => ({ ...o, id: C.uid() })) }));
  // si la otra presentación es de otro tamaño, se escala
  const kx = P().size.w / (r.pres.size?.w || P().size.w); const ky = P().size.h / (r.pres.size?.h || P().size.h);
  if (Math.abs(kx - 1) > 0.01 || Math.abs(ky - 1) > 0.01) copy.forEach((s) => s.objects.forEach((o) => { o.x *= kx; o.y *= ky; o.w *= kx; o.h *= ky; }));
  const at = C.cur() + 1;
  C.commit(t('Insertar diapositivas'), () => { P().slides.splice(at, 0, ...copy); C.setCur(at); });
  toast(tf('Se insertaron {0} diapositivas', copy.length));
  return copy.length;
}

// ================================================================ duplicar (Mayús+F3)
export async function duplicateDialog() {
  const objs = C.selObjs(); if (!objs.length) { toast('Selecciona un objeto.', { error: true }); return null; }
  const o0 = objs[0];
  const body = await tabbed('Duplicar', [{ label: 'Duplicar', body: h('div', { class: 'fx-grid' },
    field('Número de copias', numIn('n', 1, { step: 1, min: 1, max: 100 })), h('span'),
    field('Desplazamiento X', numIn('dx', 20, { step: 1, min: -2000, max: 2000, unit: 'px' })), field('Desplazamiento Y', numIn('dy', 20, { step: 1, min: -2000, max: 2000, unit: 'px' })),
    field('Ángulo de giro', numIn('rot', 0, { step: 1, min: -360, max: 360, unit: '°' })), h('span'),
    field('Aumentar el ancho', numIn('dw', 0, { step: 1, min: -1000, max: 1000, unit: 'px' })), field('Aumentar el alto', numIn('dh', 0, { step: 1, min: -1000, max: 1000, unit: 'px' })),
    field('Color inicial', h('input', { type: 'color', name: 'c1', value: /^#[0-9a-f]{6}$/i.test(o0.style?.fill || '') ? o0.style.fill : '#dc3a63' })), field('Color final', h('input', { type: 'color', name: 'c2', value: /^#[0-9a-f]{6}$/i.test(o0.style?.fill || '') ? o0.style.fill : '#dc3a63' })),
    chk('colors', 'Cambiar el color de relleno del inicial al final', false)) }], { width: '520px' });
  if (!body) return null;
  return duplicateObjects({ n: Number(v(body, 'n')) || 1, dx: Number(v(body, 'dx')) || 0, dy: Number(v(body, 'dy')) || 0, rot: Number(v(body, 'rot')) || 0, dw: Number(v(body, 'dw')) || 0, dh: Number(v(body, 'dh')) || 0, colors: v(body, 'colors') ? [v(body, 'c1'), v(body, 'c2')] : null });
}
const mix = (a, b, k) => `#${[1, 3, 5].map((i) => Math.round(parseInt(a.slice(i, i + 2), 16) * (1 - k) + parseInt(b.slice(i, i + 2), 16) * k).toString(16).padStart(2, '0')).join('')}`;
export function duplicateObjects({ n = 1, dx = 20, dy = 20, rot = 0, dw = 0, dh = 0, colors = null }) {
  const objs = C.selObjs(); if (!objs.length) return null;
  const made = [];
  C.commit(t('Duplicar'), () => {
    const s = C.slide();
    for (const o of objs) {
      const live = s.objects.find((x) => x.id === o.id) || o;
      for (let k = 1; k <= n; k++) {
        const c = JSON.parse(JSON.stringify(live)); c.id = C.uid(); delete c.role;
        c.x = live.x + dx * k; c.y = live.y + dy * k; c.w = Math.max(4, live.w + dw * k); c.h = Math.max(4, live.h + dh * k); c.rot = ((live.rot || 0) + rot * k) % 360;
        if (colors) c.style = { ...(c.style || {}), fill: mix(colors[0], colors[1], k / n) };
        if (c.group) c.group = `${c.group}-d${k}`;
        s.objects.push(c); made.push(c.id);
      }
    }
  });
  C.select(made);
  return made.length;
}

// ================================================================ campos
export const FIELDS = [['date', 'Fecha'], ['time', 'Hora'], ['num', 'Número de diapositiva'], ['count', 'Número de diapositivas'], ['title', 'Título de la diapositiva'], ['file', 'Nombre del archivo'], ['author', 'Autor']];
export function fieldValue(kind, s) {
  const pres = P(); const i = pres.slides.indexOf(s);
  switch (kind) {
    case 'date': return new Date().toLocaleDateString(document.documentElement.lang === 'en' ? 'en-US' : 'es-MX');
    case 'time': return new Date().toLocaleTimeString(document.documentElement.lang === 'en' ? 'en-US' : 'es-MX', { hour: '2-digit', minute: '2-digit' });
    case 'num': return i >= 0 ? String(i + (pres.firstNum ?? 1)) : '#';
    case 'count': return String(pres.slides.length);
    case 'title': return (i >= 0 ? titleOf(s) : '') || t('Título');
    case 'file': return C.docName?.() || '';
    case 'author': return C.docProps?.().author || '';
    default: return '';
  }
}
// actualiza los campos de una diapositiva ya dibujada
export function updateFields(el, s) { el.querySelectorAll('.fld[data-fld]').forEach((f) => { if (f.dataset.fixed) return; f.textContent = fieldValue(f.dataset.fld, s); }); }
export function insertField(kind) {
  const ed = document.querySelector('.stage [contenteditable="true"], .slide [contenteditable="true"]');
  if (!ed) { toast('Haz doble clic en un cuadro de texto y coloca el cursor donde va el campo.', { error: true }); return false; }
  ed.focus();
  const span = document.createElement('span'); span.className = 'fld'; span.dataset.fld = kind; span.contentEditable = 'false'; span.textContent = fieldValue(kind, C.slide());
  const sel2 = getSelection(); let r = sel2.rangeCount ? sel2.getRangeAt(0) : null;
  if (!r || !ed.contains(r.startContainer)) { r = document.createRange(); r.selectNodeContents(ed); r.collapse(false); }
  r.deleteContents(); r.insertNode(span); r.setStartAfter(span); r.collapse(true); sel2.removeAllRanges(); sel2.addRange(r);
  ed.dispatchEvent(new Event('input', { bubbles: true }));
  return true;
}
export const fieldMenu = () => FIELDS.map(([k, l]) => ({ label: l, run: () => insertField(k) }));

// ================================================================ imprimir: diapositivas, notas, documento o esquema
export function printChoices() {
  const pp = P().printOpts || {};
  return [
    { name: 'content', label: 'Documento', value: pp.content || 'slides', options: [['slides', 'Diapositivas'], ['notes', 'Notas'], ['handouts', 'Documento (folletos)'], ['outline', 'Esquema']] },
    { name: 'per', label: 'Diapositivas por página', value: String(pp.per || 6), options: [['1', '1'], ['2', '2'], ['3', '3 (con líneas para notas)'], ['4', '4'], ['6', '6'], ['9', '9']] },
  ];
}
export function setPrintChoice(name, value) { const pres = P(); pres.printOpts = { ...(pres.printOpts || {}), [name]: name === 'per' ? Number(value) : value }; }
const LETTER = { w: 816, h: 1056 };
export function printPageSize() { const c = P().printOpts?.content || 'slides'; return c === 'slides' ? { w: P().size.w, h: P().size.h } : LETTER; }
// arma las hojas para imprimir o exportar a PDF
export function buildPrintPages(root, slides = P().slides) {
  const pres = P(); const pp = pres.printOpts || {}; const content = pp.content || 'slides';
  const { w: W, h: H } = pres.size;
  const mini = (s, width) => { const k = width / W; const el = C.renderSlide(s, 'print'); el.style.transform = `scale(${k})`; el.style.transformOrigin = '0 0'; return h('div', { class: 'pmini', style: { width: `${width}px`, height: `${H * k}px` } }, el); };
  const page = (...kids) => h('div', { class: 'ppage pdoc', style: { width: `${LETTER.w}px`, height: `${LETTER.h}px` } }, ...kids);
  const head = () => h('div', { class: 'phead' }, h('span', {}, C.docProps?.().title || C.docName?.() || ''), h('span', {}, new Date().toLocaleDateString('es-MX')));
  if (content === 'slides') { for (const s of slides) root.append(h('div', { class: 'ppage', style: { width: `${W}px`, height: `${H}px` } }, C.renderSlide(s, 'print'))); return; }
  if (content === 'notes') {
    slides.forEach((s) => { const i = pres.slides.indexOf(s); root.append(page(head(), h('div', { class: 'pnotes-slide' }, mini(s, 620)), h('div', { class: 'pnotes-text' }, ...(String(s.notes || '').split('\n').map((l) => h('p', {}, l)))), h('div', { class: 'pfoot' }, String(i + 1)))); });
    return;
  }
  if (content === 'outline') {
    const box = []; let pg = page(head()); root.append(pg); let used = 60;
    slides.forEach((s) => {
      const items = [h('h3', {}, `${pres.slides.indexOf(s) + 1}. ${titleOf(s) || t('(sin título)')}`), ...bodyLines(s).map((l) => h('p', { style: { marginLeft: `${24 + l.depth * 24}px` } }, `• ${l.text}`))];
      const need = 34 + (items.length - 1) * 22;
      if (used + need > LETTER.h - 80) { pg = page(head()); root.append(pg); used = 60; }
      pg.append(...items); used += need;
    });
    void box; return;
  }
  // folletos
  const per = [1, 2, 3, 4, 6, 9].includes(pp.per) ? pp.per : 6;
  const cols = { 1: 1, 2: 1, 3: 1, 4: 2, 6: 2, 9: 3 }[per]; const rowsN = Math.ceil(per / cols);
  const cw = per === 3 ? 330 : (LETTER.w - 120 - (cols - 1) * 30) / cols;
  const maxH = (LETTER.h - 140 - (rowsN - 1) * 24) / rowsN; const width = Math.min(cw, maxH * (W / H));
  for (let i = 0; i < slides.length; i += per) {
    const grid2 = h('div', { class: `phand per${per}`, style: { gridTemplateColumns: per === 3 ? `${width}px 1fr` : `repeat(${cols}, ${width}px)` } });
    slides.slice(i, i + per).forEach((s) => { grid2.append(mini(s, width)); if (per === 3) grid2.append(h('div', { class: 'plines' }, ...Array.from({ length: 7 }, () => h('i')))); });
    root.append(page(head(), grid2, h('div', { class: 'pfoot' }, String(i / per + 1))));
  }
}
