// Imprimir y exportar a PDF como en LibreOffice: diálogo de impresión con vista previa, impresora,
// copias, intervalo de páginas, páginas por hoja, color y doble cara; vista previa de impresión a
// pantalla completa; y el diálogo «Opciones de PDF» (intervalo, marcadores, PDF etiquetado, vista
// inicial, marca de agua, contraseñas y permisos). Compartido por las apps de la suite. GPL-3.0
import { h, t, tf, modal, toast } from './office.js';
import { tabbed, field, numIn, sel, chk } from './dialogs.js';

const api = window.oficina;
const LS = 'oficina.pdfopts';
const load = () => { try { return JSON.parse(localStorage.getItem(LS) || '{}'); } catch { return {}; } };
const store = (o) => { try { localStorage.setItem(LS, JSON.stringify(o)); } catch { /* sin almacenamiento */ } };

// «1-3, 5, 8-» → válido; devuelve la cadena normalizada o null si no es válida
export function parseRanges(s, max = Infinity) {
  const txt = String(s || '').trim(); if (!txt) return null;
  const out = [];
  for (const part of txt.split(/[,;]/)) {
    const m = /^\s*(\d+)\s*(?:(-)\s*(\d*))?\s*$/.exec(part); if (!m) return null;
    const a = Number(m[1]); const b = m[2] ? (m[3] ? Number(m[3]) : max) : a;
    if (a < 1 || b < a) return null;
    out.push([a, Math.min(b, max)]);
  }
  return out.filter(([a]) => a <= max);
}
const rangeStr = (r) => r.map(([a, b]) => (a === b ? `${a}` : `${a}-${b}`)).join(', ');

// ---------------------------------------------------------------- páginas como imágenes (pdf.js)
let pdfjs = null;
async function pdfLib() {
  if (pdfjs) return pdfjs;
  const base = new URL('../../node_modules/pdfjs-dist/build/', location.href).href;
  pdfjs = await import(/* @vite-ignore */ `${base}pdf.mjs`);
  pdfjs.GlobalWorkerOptions.workerSrc = `${base}pdf.worker.mjs`;
  return pdfjs;
}
async function renderedPages(shell) {
  const bytes = await shell.withPrintLayout(() => api.pdfBytes(shell.adapter.pdfOptions?.() || {}));
  const lib = await pdfLib();
  const doc = await lib.getDocument({ data: bytes, standardFontDataUrl: new URL('../../node_modules/pdfjs-dist/standard_fonts/', location.href).href }).promise;
  const cache = new Map();
  return {
    count: doc.numPages,
    async canvas(i, width) {
      const key = `${i}:${width}`; if (cache.has(key)) return cache.get(key);
      const pg = await doc.getPage(i + 1);
      const vp0 = pg.getViewport({ scale: 1 });
      const vp = pg.getViewport({ scale: (width * (window.devicePixelRatio || 1)) / vp0.width });
      const c = h('canvas', { width: Math.round(vp.width), height: Math.round(vp.height), class: 'pv-page', style: { width: `${width}px` } });
      await pg.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
      cache.set(key, c); return c;
    },
    destroy() { doc.destroy(); },
  };
}

// opciones propias de la app (por ejemplo, en Impresiona: diapositivas, notas, folletos o esquema)
function choiceFields(shell, onChange) {
  const list = shell.adapter.printChoices?.() || [];
  return list.map((c) => { const el = sel(`app_${c.name}`, c.value, c.options); el.addEventListener('change', () => { shell.adapter.setPrintChoice(c.name, el.value); onChange?.(); }); return field(c.label, el); });
}

// ---------------------------------------------------------------- diálogo Imprimir
export async function printDialog(shell) {
  if (!api?.pdfBytes) return shell.print();
  const prev = load().print || {};
  const printers = await api.printers().catch(() => []);
  let pages;
  try { pages = await renderedPages(shell); } catch (err) { console.error(err); return shell.print(); }
  const cur = Math.max(1, Math.min(pages.count, shell.adapter.currentPage?.() || 1));
  const pr = h('select', { name: 'printer' }, h('option', { value: '' }, t('Elegir en el cuadro del sistema…')), ...printers.map((p) => h('option', { value: p.name, selected: (prev.printer ? p.name === prev.printer : p.isDefault) ? true : null }, p.displayName)));
  const range = sel('range', 'all', [['all', 'Todas las páginas'], ['pages', 'Páginas'], ['current', 'Página actual']]);
  const rtxt = h('input', { name: 'pages', placeholder: '1-3, 5', value: `1-${pages.count}`, disabled: true });
  range.addEventListener('change', () => { rtxt.disabled = range.value !== 'pages'; if (!rtxt.disabled) rtxt.focus(); redraw(); });
  rtxt.addEventListener('input', () => redraw());
  const total = h('p', { class: 'fx-note', style: { gridColumn: '1 / -1' } }, tf(pages.count === 1 ? '{0} página en total' : '{0} páginas en total', pages.count));
  const reload = async () => { pages.destroy(); pages = await renderedPages(shell); total.textContent = tf(pages.count === 1 ? '{0} página en total' : '{0} páginas en total', pages.count); at = 0; redraw(); };
  const opts = h('div', { class: 'fx-grid pv-opts' },
    ...choiceFields(shell, reload),
    field('Impresora', pr, 'wide'),
    field('Copias', numIn('copies', 1, { step: 1, min: 1, max: 999 })), chk('collate', 'Intercalar', prev.collate !== false),
    field('Intervalo', range), field('Páginas', rtxt),
    field('Páginas por hoja', sel('pps', prev.pps || 1, [[1, '1'], [2, '2'], [4, '4'], [6, '6'], [9, '9'], [16, '16']])),
    field('Color', sel('color', prev.color || 'color', [['color', 'Color'], ['gray', 'Escala de grises']])),
    field('Doble cara', sel('duplex', prev.duplex || 'simplex', [['simplex', 'Una cara'], ['longEdge', 'Doble cara, borde largo'], ['shortEdge', 'Doble cara, borde corto']])),
    total);
  const sheet = h('div', { class: 'pv-sheet' });
  const info = h('span', { class: 'pv-info' });
  let at = 0;
  const chosen = () => {
    const v = range.value;
    if (v === 'current') return [cur - 1];
    if (v === 'pages') { const r = parseRanges(rtxt.value, pages.count); if (!r) return null; return r.flatMap(([a, b]) => Array.from({ length: b - a + 1 }, (_, k) => a - 1 + k)); }
    return Array.from({ length: pages.count }, (_, k) => k);
  };
  async function redraw() {
    const list = chosen(); rtxt.classList.toggle('bad', !list);
    if (!list?.length) { sheet.replaceChildren(h('p', { class: 'muted' }, t('Intervalo de páginas no válido'))); info.textContent = ''; return; }
    const n = Number(opts.querySelector('[name=pps]').value) || 1;
    const sheets = Math.ceil(list.length / n); at = Math.max(0, Math.min(at, sheets - 1));
    const cols = { 1: 1, 2: 2, 4: 2, 6: 3, 9: 3, 16: 4 }[n]; const w = Math.floor(260 / cols) - 4;
    const grid = h('div', { class: 'pv-grid', style: { gridTemplateColumns: `repeat(${cols}, 1fr)`, filter: opts.querySelector('[name=color]').value === 'gray' ? 'grayscale(1)' : '' } });
    sheet.replaceChildren(grid);
    for (const i of list.slice(at * n, at * n + n)) grid.append(copyCanvas(await pages.canvas(i, w)));
    info.textContent = tf('Hoja {0} de {1}', at + 1, sheets);
  }
  const nav = h('div', { class: 'pv-nav' },
    h('button', { class: 'btn', type: 'button', 'aria-label': t('Anterior'), onclick: () => { at -= 1; redraw(); } }, '‹'), info,
    h('button', { class: 'btn', type: 'button', 'aria-label': t('Siguiente'), onclick: () => { at += 1; redraw(); } }, '›'));
  opts.addEventListener('change', (e) => { if (!e.target.name?.startsWith('app_')) redraw(); });
  const body = h('div', { class: 'pv-dlg' }, h('div', { class: 'pv-left' }, sheet, nav), opts);
  const p = modal({ title: t('Imprimir'), body, width: '760px', buttons: [{ label: t('Cancelar'), value: null }, { label: t('Imprimir'), primary: true, value: true }] });
  redraw();
  const ok = await p; pages.destroy();
  if (!ok) return null;
  const list = chosen(); if (!list) { toast('Intervalo de páginas no válido', { error: true }); return null; }
  const v = (n) => opts.querySelector(`[name="${n}"]`);
  const o = { copies: Math.max(1, Number(v('copies').value) || 1), collate: v('collate').checked, pagesPerSheet: Number(v('pps').value) || 1, color: v('color').value !== 'gray', duplexMode: v('duplex').value };
  if (range.value !== 'all') o.pageRanges = rangesOf(list);
  if (pr.value) { o.deviceName = pr.value; o.silent = true; }
  store({ ...load(), print: { printer: pr.value, collate: o.collate, pps: o.pagesPerSheet, color: v('color').value, duplex: o.duplexMode } });
  return shell.print(o);
}
function copyCanvas(c) { const d = h('canvas', { width: c.width, height: c.height, class: 'pv-page', style: { width: c.style.width } }); d.getContext('2d').drawImage(c, 0, 0); return d; }
// [0,1,2,5] → [{from:0,to:2},{from:5,to:5}] (base 0, como pide Electron)
function rangesOf(list) {
  const out = [];
  for (const i of list) { const l = out[out.length - 1]; if (l && i === l.to + 1) l.to = i; else out.push({ from: i, to: i }); }
  return out;
}

// ---------------------------------------------------------------- vista previa de impresión
export async function printPreview(shell) {
  if (!api?.pdfBytes) return null;
  let pages;
  try { pages = await renderedPages(shell); } catch (err) { console.error(err); toast('No se pudo preparar la vista previa.', { error: true }); return null; }
  const prev = load().preview || {};
  let cols = prev.cols || 2; let zoom = prev.zoom || 100;
  const area = h('div', { class: 'pv-area' });
  const cnt = h('span', { class: 'pv-info' }, tf(pages.count === 1 ? '{0} página' : '{0} páginas', pages.count));
  const colSel = h('select', { 'aria-label': t('Páginas por fila') }, ...[1, 2, 3, 4, 6].map((n) => h('option', { value: n, selected: n === cols ? true : null }, tf('{0} por fila', n))));
  const zIn = h('input', { type: 'range', min: 40, max: 200, value: zoom, 'aria-label': t('Zoom') });
  const close = () => { back.remove(); document.removeEventListener('keydown', onKey, true); pages.destroy(); store({ ...load(), preview: { cols, zoom } }); };
  const bar = h('div', { class: 'pv-bar', role: 'toolbar', 'aria-label': t('Vista previa de impresión') },
    h('strong', {}, t('Vista previa de impresión')), cnt, h('span', { class: 'grow' }), ...choiceFields(shell, async () => { pages.destroy(); pages = await renderedPages(shell); cnt.textContent = tf(pages.count === 1 ? '{0} página' : '{0} páginas', pages.count); draw(); }), colSel, zIn,
    h('button', { class: 'btn', onclick: () => { close(); printDialog(shell); } }, t('Imprimir…')),
    h('button', { class: 'btn', onclick: () => { close(); pdfDialog(shell); } }, t('Exportar como PDF…')),
    h('button', { class: 'btn primary', onclick: () => close() }, t('Cerrar vista previa')));
  const back = h('div', { class: 'pv-full', role: 'dialog', 'aria-modal': 'true', 'aria-label': t('Vista previa de impresión') }, bar, area);
  async function draw() {
    const width = Math.round(((Math.min(window.innerWidth - 80, 1500) / cols) - 24) * (zoom / 100));
    const grid = h('div', { class: 'pv-grid big', style: { gridTemplateColumns: `repeat(${cols}, ${width}px)` } });
    area.replaceChildren(grid);
    for (let i = 0; i < pages.count; i += 1) grid.append(h('figure', {}, copyCanvas(await pages.canvas(i, width)), h('figcaption', {}, String(i + 1))));
  }
  colSel.addEventListener('change', () => { cols = Number(colSel.value); draw(); });
  zIn.addEventListener('change', () => { zoom = Number(zIn.value); draw(); });
  const onKey = (e) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); close(); } };
  document.addEventListener('keydown', onKey, true);
  document.body.append(back);
  await draw();
  bar.querySelector('.btn.primary').focus();
  return back;
}

// ---------------------------------------------------------------- Opciones de PDF
export async function pdfDialog(shell) {
  const L = load().pdf || {};
  const cur = shell.adapter.currentPage?.();
  const pw = (name, label) => field(label, h('input', { type: 'password', name, autocomplete: 'new-password' }));
  const body = await tabbed('Opciones de PDF', [
    { label: 'General', body: h('div', { class: 'fx-grid' }, ...choiceFields(shell),
      field('Intervalo', sel('range', 'all', [['all', 'Todo'], ['pages', 'Páginas'], ...(cur ? [['current', 'Página actual']] : [])])), field('Páginas', h('input', { name: 'pages', placeholder: '1-3, 5' })),
      chk('outline', 'Exportar los títulos como marcadores', L.outline !== false), chk('tagged', 'PDF etiquetado (accesible para lectores de pantalla)', !!L.tagged),
      chk('meta', 'Incluir las propiedades del documento (título, autor, asunto, palabras clave)', L.meta !== false), chk('open', 'Abrir en PapelPDF después de exportar', !!L.open)) },
    { label: 'Vista inicial', body: h('div', { class: 'fx-grid' },
      field('Panel', sel('initialView', L.initialView || 'none', [['none', 'Solo la página'], ['outline', 'Marcadores y página'], ['thumbs', 'Miniaturas y página'], ['full', 'Pantalla completa']])),
      field('Diseño de página', sel('layout', L.layout || '', [['', 'Predeterminado'], ['single', 'Página individual'], ['continuous', 'Continuo'], ['facing', 'Páginas enfrentadas'], ['book', 'Libro (la primera sola)']])),
      chk('fitWindow', 'Ajustar la ventana a la página', !!L.fitWindow), chk('displayTitle', 'Mostrar el título del documento en la ventana', L.displayTitle !== false)) },
    { label: 'Marca de agua', body: h('div', { class: 'fx-grid' },
      field('Texto', h('input', { name: 'wm', value: L.wm || '', placeholder: t('Por ejemplo: BORRADOR') }), 'wide'),
      field('Color', h('input', { type: 'color', name: 'wmColor', value: L.wmColor || '#c0c0c0' })), field('Opacidad', numIn('wmOpacity', L.wmOpacity ?? 35, { step: 5, min: 5, max: 100, unit: '%' }))) },
    { label: 'Seguridad', body: h('div', { class: 'fx-grid' },
      pw('user', 'Contraseña para abrir'), pw('user2', 'Confirmar contraseña'), pw('owner', 'Contraseña de permisos'), h('span'),
      chk('print', 'Permitir imprimir', true), chk('copy', 'Permitir copiar texto e imágenes', true), chk('modify', 'Permitir modificar', false),
      h('p', { class: 'fx-note', style: { gridColumn: '1 / -1' } }, t('Cifrado AES de 256 bits. Las contraseñas no se guardan.'))) },
  ], { width: '600px' });
  if (!body) return null;
  const v = (n) => body.querySelector(`[name="${n}"]`);
  const o = { outline: v('outline').checked, tagged: v('tagged').checked, open: v('open').checked, initialView: v('initialView').value !== 'none' ? v('initialView').value : undefined, layout: v('layout').value || undefined, fitWindow: v('fitWindow').checked || undefined, displayTitle: v('displayTitle').checked || undefined };
  if (v('range').value === 'pages') { const r = parseRanges(v('pages').value); if (!r) { toast('Intervalo de páginas no válido', { error: true }); return null; } o.pageRanges = rangeStr(r); }
  if (v('range').value === 'current') o.pageRanges = String(cur);
  if (v('meta').checked) { const p = shell.docProps(); o.meta = { title: p.title, author: p.author, subject: p.subject, keywords: p.keywords, creator: `${shell.appName} · Mi Oficina MX` }; }
  if (v('wm').value.trim()) o.watermark = { text: v('wm').value.trim(), color: v('wmColor').value, opacity: Math.max(0.05, Math.min(1, Number(v('wmOpacity').value) / 100)) };
  if (v('user').value || v('owner').value) {
    if (v('user').value !== v('user2').value) { toast('Las contraseñas no coinciden', { error: true }); return null; }
    o.password = { user: v('user').value, owner: v('owner').value || v('user').value, print: v('print').checked, copy: v('copy').checked, modify: v('modify').checked };
  }
  store({ ...load(), pdf: { outline: o.outline, tagged: o.tagged, open: o.open, meta: v('meta').checked, initialView: o.initialView, layout: o.layout, fitWindow: o.fitWindow, displayTitle: !!o.displayTitle, wm: v('wm').value.trim(), wmColor: v('wmColor').value, wmOpacity: Number(v('wmOpacity').value) } });
  return shell.exportPdf(o);
}
