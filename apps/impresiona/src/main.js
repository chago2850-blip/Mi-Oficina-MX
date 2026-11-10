// Impresiona — presentaciones de Mi Oficina MX. GPL-3.0-or-later
import { OfficeShell, h, $, $$, icon, toast, modal, form, dropdown, closeDropdown, esc, api, colorPicker, t, localizeHtml } from '../../shared/office.js';
import { THEMES, LAYOUTS, themeById, layoutObjects, PLACEHOLDER } from './themes.js';
import { readPptx, writePptx } from './pptx.js';
import { readCoreProps } from '../../shared/coreprops.js';
import { CHART_TYPES, parseChartData, chartDataText, chartSvg } from './charts.js';
import { runShow, runPresenter, runAudience, EXIT_TYPES, PATH_TYPES } from './show.js';
import * as F3 from './fase3.js';
import * as F4 from './fase4.js';
import { readOdp, writeOdp } from './odp.js';

const uid = () => Math.random().toString(36).slice(2, 10);
const clone = (o) => JSON.parse(JSON.stringify(o));
const FONTS = ['Inter', 'Calibri', 'Arial', 'Georgia', 'Verdana', 'Trebuchet MS', 'Times New Roman', 'Courier New'];
const SIZES = [10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 54, 60, 72, 96];
const TRANSITIONS = [['none', 'Ninguna'], ['fade', 'Desvanecer'], ['push', 'Empujar'], ['wipe', 'Barrido'], ['cover', 'Cubrir'], ['zoom', 'Zoom'], ['morph', 'Transformación']];
const ANIMS = [['none', 'Ninguna'], ['appear', 'Aparecer'], ['fade', 'Desvanecer'], ['fly', 'Entrar desde abajo'], ['zoom', 'Acercar'], ['wipe', 'Barrido'], ['fadeOut', 'Salida: desvanecer'], ['flyOut', 'Salida: hacia abajo'], ['zoomOut', 'Salida: alejar'], ['pathRight', 'Trayectoria: derecha'], ['pathLeft', 'Trayectoria: izquierda'], ['pathUp', 'Trayectoria: arriba'], ['pathDown', 'Trayectoria: abajo'], ['pathArc', 'Trayectoria: arco'], ['pathCircle', 'Trayectoria: círculo'], ['pathWave', 'Trayectoria: onda']];
const ANIM_NAMES = Object.fromEntries(ANIMS);

// ------------------------------------------------------------------ estado
let pres = null; // { size:{w,h}, theme, slides:[...] }
let cur = 0;
let selected = new Set();
let editing = null; // id del objeto en edición
let shell, stageWrap, stage, slideEl, selLayer, thumbs, notesEl, zoomCtl, statusSlide;
let scale = 1, fitMode = true;
const undoStack = [], redoStack = [];

let masterMode = false, sorterMode = false;
const slide = () => (masterMode ? (pres.master ||= { objects: [], footer: '', showNumber: false, showDate: false }) : pres.slides[cur]);
const objById = (id) => slide().objects.find((o) => o.id === id);
const theme = () => themeById(pres.theme);

function commit(label, fn) {
  undoStack.push(JSON.stringify({ pres, cur }));
  if (undoStack.length > 80) undoStack.shift();
  redoStack.length = 0;
  fn?.();
  shell.markDirty();
  renderAll();
  void label;
}
function undo() { if (editing) { document.execCommand('undo'); return; } const s = undoStack.pop(); if (!s) return; redoStack.push(JSON.stringify({ pres, cur })); const j = JSON.parse(s); pres = j.pres; cur = Math.min(j.cur, pres.slides.length - 1); selected.clear(); renderAll(); shell.markDirty(); }
function redo() { const s = redoStack.pop(); if (!s) return; undoStack.push(JSON.stringify({ pres, cur })); const j = JSON.parse(s); pres = j.pres; cur = Math.min(j.cur, pres.slides.length - 1); selected.clear(); renderAll(); shell.markDirty(); }

// ------------------------------------------------------------------ dibujar una diapositiva
function objStyle(o, th) {
  const st = o.style || {};
  const isTitle = o.role === 'title';
  const color = st.color || (o.type === 'shape' ? '#FFFFFF' : (isTitle ? th.title : th.text));
  const font = st.font || (isTitle || o.role === 'subtitle' ? th.titleFont : th.bodyFont);
  return { color, font, size: st.size || (o.type === 'shape' ? 18 : 20) };
}
function shapeSvg(o, th) {
  const st = o.style || {}, w = o.w, hh = o.h;
  const fill = st.fill === undefined ? th.accent : (st.fill || 'none');
  const stroke = st.stroke || 'none', sw = st.strokeW || (st.stroke ? 2 : 0);
  const i = sw / 2;
  let d;
  switch (o.shape) {
    case 'ellipse': return `<svg width="100%" height="100%" viewBox="0 0 ${w} ${hh}" preserveAspectRatio="none"><ellipse cx="${w / 2}" cy="${hh / 2}" rx="${Math.max(0, w / 2 - i)}" ry="${Math.max(0, hh / 2 - i)}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/></svg>`;
    case 'roundRect': return `<svg width="100%" height="100%" viewBox="0 0 ${w} ${hh}" preserveAspectRatio="none"><rect x="${i}" y="${i}" width="${Math.max(0, w - sw)}" height="${Math.max(0, hh - sw)}" rx="${Math.min(w, hh) * 0.15}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/></svg>`;
    case 'triangle': d = `M${w / 2},${i} L${w - i},${hh - i} L${i},${hh - i} Z`; break;
    case 'arrow': d = `M${i},${hh * 0.3} L${w * 0.62},${hh * 0.3} L${w * 0.62},${i} L${w - i},${hh / 2} L${w * 0.62},${hh - i} L${w * 0.62},${hh * 0.7} L${i},${hh * 0.7} Z`; break;
    case 'star': { const pts = []; for (let k = 0; k < 10; k++) { const a = -Math.PI / 2 + (k * Math.PI) / 5; const r = k % 2 ? 0.4 : 0.5; pts.push(`${w / 2 + Math.cos(a) * r * w},${hh / 2 + Math.sin(a) * r * hh}`); } d = `M${pts.join(' L')} Z`; break; }
    case 'line': return `<svg width="100%" height="100%" viewBox="0 0 ${w} ${hh}" preserveAspectRatio="none"><line x1="0" y1="${hh / 2}" x2="${w}" y2="${hh / 2}" stroke="${st.stroke || th.accent}" stroke-width="${st.strokeW || 4}"/></svg>`;
    default: return `<svg width="100%" height="100%" viewBox="0 0 ${w} ${hh}" preserveAspectRatio="none"><rect x="${i}" y="${i}" width="${Math.max(0, w - sw)}" height="${Math.max(0, hh - sw)}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/></svg>`;
  }
  return `<svg width="100%" height="100%" viewBox="0 0 ${w} ${hh}" preserveAspectRatio="none"><path d="${d}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/></svg>`;
}

// mode: 'edit' | 'thumb' | 'show' | 'print'
function renderSlide(s, mode = 'thumb') {
  const th = theme();
  const { w: W, h: H } = pres.size;
  const el = h('div', { class: `slide m-${mode}`, style: { width: `${W}px`, height: `${H}px`, background: s.bgImage ? `center/cover url("${s.bgImage}")` : (s.bg || th.band || th.bg) } });
  if (!s.bg && !s.bgImage && th.id === 'claro') el.append(h('div', { class: 'deco', style: { background: th.accent } }));
  const isMaster = s === pres.master;
  const master = !isMaster && pres.master ? pres.master : null;
  const list = master ? [...master.objects.map((o) => ({ ...o, __m: true })), ...s.objects] : s.objects;
  for (const o of list) {
    const st = o.style || {};
    const box = h('div', { class: `obj t-${o.type}${o.role ? ` r-${o.role}` : ''}${o.__m ? ' mobj' : ''}`, 'data-id': o.__m ? null : o.id, style: { left: `${o.x}px`, top: `${o.y}px`, width: `${o.w}px`, height: `${o.h}px`, transform: o.rot ? `rotate(${o.rot}deg)` : '' } });
    if (o.anim && o.anim.type !== 'none' && !EXIT_TYPES.includes(o.anim.type) && !PATH_TYPES.includes(o.anim.type) && mode === 'show' && !o.__m) box.classList.add('anim-wait');
    if (o.type === 'media') F3.mediaBox(o, mode, box);
    if (o.type === 'camera') F4.cameraBox(o, mode, box);
    if (mode === 'edit' && o.anim && !o.__m) F3.pathPreview(o, box);
    if (o.type === 'chart') box.innerHTML = chartSvg(o, th);
    if (o.anim && o.anim.type !== 'none' && mode === 'edit') box.append(h('span', { class: 'anim-badge', title: ANIM_NAMES[o.anim.type] }, String(o.anim.order || '')));
    if (o.type === 'image') box.append(h('img', { src: o.src, alt: '', draggable: 'false' }));
    if (o.type === 'shape') box.innerHTML = shapeSvg(o, th);
    if (o.type === 'table') {
      const t = h('table', { class: 'stable', style: { fontSize: `${st.size || 14}pt`, color: th.text } });
      o.rows.forEach((row, ri) => {
        const tr = h('tr');
        row.forEach((c) => { if (c.hidden) return; tr.append(h(ri === 0 ? 'th' : 'td', { rowspan: c.rs > 1 ? c.rs : null, colspan: c.cs > 1 ? c.cs : null, style: { background: c.bg || (ri === 0 ? th.accent : ''), color: c.color || (ri === 0 ? '#fff' : ''), fontWeight: c.bold || ri === 0 ? '600' : '' } }, c.text)); });
        t.append(tr);
      });
      box.append(t);
    }
    if (o.type === 'text' || (o.type === 'shape' && (o.html || mode === 'edit'))) {
      const os = objStyle(o, th);
      const txt = h('div', { class: `otxt v-${st.valign || 'top'}${st.list ? ' list' : ''}`, style: { color: os.color, fontFamily: `"${os.font}", Inter, Carlito, sans-serif`, fontSize: `${os.size}pt`, textAlign: st.align || (o.type === 'shape' ? 'center' : 'left'), fontWeight: st.bold ? '700' : '', fontStyle: st.italic ? 'italic' : '', background: o.type === 'text' && st.fill ? st.fill : '', border: o.type === 'text' && st.stroke ? `${st.strokeW || 1}px solid ${st.stroke}` : '' } });
      const inner = h('div', { class: 'tin', html: o.html || '' });
      if (o.role && !o.html) { inner.dataset.ph = t(PLACEHOLDER[o.role] || ''); if (mode !== 'edit') inner.classList.add('ph-hidden'); }
      txt.append(inner);
      box.append(txt);
    }
    el.append(box);
  }
  const mst = pres.master;
  if (mst && !isMaster) {
    const idx = pres.slides.indexOf(s);
    if (mst.footer) el.append(h('div', { class: 'master-footer', style: { color: th.text } }, mst.footer));
    if (mst.showNumber && idx >= 0) el.append(h('div', { class: 'master-num', style: { color: th.text } }, String(idx + 1)));
    if (mst.showDate) el.append(h('div', { class: 'master-date', style: { color: th.text } }, new Date().toLocaleDateString('es-MX')));
  }
  if (isMaster && mode === 'edit') el.append(h('div', { class: 'master-banner' }, 'Patrón: lo que pongas aquí aparece en todas las diapositivas'));
  return el;
}

// ------------------------------------------------------------------ espacio de trabajo
function buildWorkspace(ws) {
  ws.classList.add('pres');
  thumbs = h('aside', { class: 'thumbs', 'aria-label': 'Diapositivas' });
  stageWrap = h('div', { class: 'stagewrap' });
  stage = h('div', { class: 'stage' });
  selLayer = h('div', { class: 'sellayer' });
  stageWrap.append(stage);
  notesEl = h('textarea', { class: 'notes', placeholder: 'Haz clic para agregar notas del orador', 'aria-label': 'Notas del orador' });
  const center = h('div', { class: 'center' }, stageWrap, h('div', { class: 'notesbar' }, notesEl));
  ws.append(thumbs, center, h('div', { id: 'printRoot' }));
  notesEl.addEventListener('focus', () => { notesEl.dataset.before = slide().notes || ''; });
  notesEl.addEventListener('input', () => { slide().notes = notesEl.value; shell.markDirty(); });
  notesEl.addEventListener('blur', () => { if ((notesEl.dataset.before || '') !== notesEl.value) { const v = notesEl.value; slide().notes = notesEl.dataset.before; commit('Notas', () => { slide().notes = v; }); } });
  new ResizeObserver(() => fit()).observe(stageWrap);
  stageWrap.addEventListener('mousedown', onStageDown);
  stageWrap.addEventListener('dblclick', onStageDbl);
  stageWrap.addEventListener('contextmenu', onStageContext);
  addEventListener('keydown', onKey);
  document.addEventListener('copy', onCopy); document.addEventListener('cut', (e) => onCopy(e, true)); document.addEventListener('paste', onPaste);
}

function fit() {
  if (!pres) return;
  const { w: W, h: H } = pres.size;
  if (fitMode) {
    const avW = stageWrap.clientWidth - 64, avH = stageWrap.clientHeight - 48;
    scale = Math.max(0.1, Math.min(avW / W, avH / H));
  }
  const zl = zoomCtl?.querySelector('span'); if (zl) zl.textContent = `${Math.round(scale * 100)}%`;
  stage.style.width = `${W * scale}px`; stage.style.height = `${H * scale}px`;
  if (slideEl) slideEl.style.transform = `scale(${scale})`;
  drawSelection();
}

function renderAll() {
  if (!pres) return;
  cur = Math.max(0, Math.min(cur, pres.slides.length - 1));
  renderStage();
  renderThumbs();
  notesEl.value = slide()?.notes || '';
  statusSlide.textContent = masterMode ? 'Vista Patrón' : `Diapositiva ${cur + 1} de ${pres.slides.length}`;
  renderAnimPanel();
  reflect();
}
function renderStage() {
  stage.replaceChildren();
  if (!slide()) return;
  slideEl = renderSlide(slide(), 'edit');
  slideEl.style.transformOrigin = '0 0';
  stage.append(slideEl, selLayer);
  if (editing) { const o = objById(editing); if (o) startEditing(o.id, false); else editing = null; }
  fit();
}
function renderThumbs() {
  const scroll = thumbs.scrollTop;
  thumbs.replaceChildren();
  const tw = 172, ts = tw / pres.size.w;
  pres.slides.forEach((s, i) => {
    const mini = renderSlide(s, 'thumb');
    mini.style.transform = `scale(${ts})`; mini.style.transformOrigin = '0 0';
    const frame = h('div', { class: 'tframe', style: { width: `${tw}px`, height: `${pres.size.h * ts}px` } }, mini);
    if (s.section) thumbs.append(h('div', { class: 'thumb-sec', ondblclick: () => renameSection(i) }, s.section));
    const item = h('div', { class: `thumb${i === cur ? ' on' : ''}`, draggable: 'true', 'data-i': i, tabindex: 0 }, h('span', { class: 'n' }, String(i + 1), s.transition?.type && s.transition.type !== 'none' ? h('i', { class: 'tr', title: 'Tiene transición' }, '★') : null), frame);
    item.addEventListener('click', () => goto(i));
    item.addEventListener('contextmenu', (e) => { e.preventDefault(); goto(i); dropdown(item, [{ label: 'Nueva diapositiva', icon: 'slideAdd', run: () => addSlide() }, { label: 'Duplicar diapositiva', icon: 'slideDup', run: duplicateSlide }, { label: 'Eliminar diapositiva', icon: 'slideDel', run: deleteSlide }, '-', { label: 'Presentar desde aquí', icon: 'presentCur', run: () => present(i) }]); });
    item.addEventListener('dragstart', (e) => { e.dataTransfer.setData('text/x-slide', String(i)); item.classList.add('dragging'); });
    item.addEventListener('dragend', () => item.classList.remove('dragging'));
    item.addEventListener('dragover', (e) => { if ([...e.dataTransfer.types].includes('text/x-slide')) { e.preventDefault(); item.classList.add('over'); } });
    item.addEventListener('dragleave', () => item.classList.remove('over'));
    item.addEventListener('drop', (e) => { e.preventDefault(); e.stopPropagation(); item.classList.remove('over'); const from = Number(e.dataTransfer.getData('text/x-slide')); if (Number.isNaN(from) || from === i) return; commit('Mover diapositiva', () => { const [s2] = pres.slides.splice(from, 1); pres.slides.splice(i, 0, s2); cur = i; }); });
    thumbs.append(item);
  });
  thumbs.append(h('button', { class: 'addslide', html: `${icon('plus', 16)} Nueva diapositiva`, onclick: () => addSlide() }));
  thumbs.scrollTop = scroll;
  thumbs.querySelector('.thumb.on')?.scrollIntoView({ block: 'nearest' });
}
const thumbLater = () => { clearTimeout(thumbLater.t); thumbLater.t = setTimeout(renderThumbs, 250); };

function goto(i) {
  if (editing) stopEditing();
  if (i < 0 || i >= pres.slides.length) return;
  cur = i; selected.clear(); renderAll();
}

// ------------------------------------------------------------------ selección y manipulación
function stagePoint(e) { const r = stage.getBoundingClientRect(); return { x: (e.clientX - r.left) / scale, y: (e.clientY - r.top) / scale }; }
function drawSelection(guides = []) {
  selLayer.replaceChildren();
  selLayer.style.width = stage.style.width; selLayer.style.height = stage.style.height;
  for (const id of selected) {
    const o = objById(id); if (!o) continue;
    const box = h('div', { class: `selbox${editing === id ? ' editing' : ''}`, style: { left: `${o.x * scale}px`, top: `${o.y * scale}px`, width: `${o.w * scale}px`, height: `${o.h * scale}px`, transform: o.rot ? `rotate(${o.rot}deg)` : '' } });
    if (selected.size === 1 && editing !== id) {
      for (const hd of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) box.append(h('div', { class: `hd ${hd}`, 'data-h': hd }));
      box.append(h('div', { class: 'rot', 'data-h': 'rot', title: 'Girar' }));
    }
    selLayer.append(box);
  }
  for (const g of guides) selLayer.append(h('div', { class: `guide ${g.axis}`, style: g.axis === 'v' ? { left: `${g.v * scale}px` } : { top: `${g.v * scale}px` } }));
}
function select(ids, add = false) {
  if (!add) selected.clear();
  ids.forEach((id) => selected.add(id));
  drawSelection(); reflect();
}

function onStageDown(e) {
  if (e.button !== 0) return;
  closeDropdown();
  const handle = e.target.dataset?.h;
  const objEl = e.target.closest('.obj');
  if (editing && objEl && objEl.dataset.id === editing) return; // clic dentro del texto en edición
  if (editing) stopEditing();
  const p0 = stagePoint(e);
  if (handle) { startTransform(e, handle, p0); return; }
  if (objEl && slideEl.contains(objEl)) {
    const id = objEl.dataset.id;
    if (e.shiftKey || e.ctrlKey) { if (selected.has(id)) selected.delete(id); else selected.add(id); drawSelection(); reflect(); return; }
    if (!selected.has(id)) select([id]);
    startMove(e, p0);
    return;
  }
  // selección por recuadro
  select([]);
  const band = h('div', { class: 'band' }); selLayer.append(band);
  const move = (ev) => {
    const p = stagePoint(ev);
    const x = Math.min(p.x, p0.x), y = Math.min(p.y, p0.y), w = Math.abs(p.x - p0.x), hh = Math.abs(p.y - p0.y);
    Object.assign(band.style, { left: `${x * scale}px`, top: `${y * scale}px`, width: `${w * scale}px`, height: `${hh * scale}px` });
    band.rect = { x, y, w, h: hh };
  };
  const up = () => {
    removeEventListener('mousemove', move); removeEventListener('mouseup', up);
    const r = band.rect; band.remove();
    if (r && (r.w > 3 || r.h > 3)) select(slide().objects.filter((o) => o.x < r.x + r.w && o.x + o.w > r.x && o.y < r.y + r.h && o.y + o.h > r.y).map((o) => o.id));
  };
  addEventListener('mousemove', move); addEventListener('mouseup', up);
}

function snap(o, dx, dy, ids) {
  const { w: W, h: H } = pres.size, th = 6 / scale;
  const others = slide().objects.filter((x) => !ids.has(x.id));
  const xs = [0, W / 2, W, ...others.flatMap((x) => [x.x, x.x + x.w / 2, x.x + x.w])];
  const ys = [0, H / 2, H, ...others.flatMap((x) => [x.y, x.y + x.h / 2, x.y + x.h])];
  const guides = [];
  const best = (cands, vals) => { let b = null; for (const c of cands) for (const v of vals) { const d = c - v.p; if (Math.abs(d) < th && (!b || Math.abs(d) < Math.abs(b.d))) b = { d, c }; } return b; };
  const nx = o.x + dx, ny = o.y + dy;
  const bx = best(xs, [{ p: nx }, { p: nx + o.w / 2 }, { p: nx + o.w }]);
  const by = best(ys, [{ p: ny }, { p: ny + o.h / 2 }, { p: ny + o.h }]);
  if (bx) { dx += bx.d; guides.push({ axis: 'v', v: bx.c }); }
  if (by) { dy += by.d; guides.push({ axis: 'h', v: by.c }); }
  return { dx, dy, guides };
}
function startMove(e, p0) {
  const ids = new Set(selected);
  const start = [...ids].map((id) => ({ o: objById(id), x: objById(id).x, y: objById(id).y }));
  const snapshot = JSON.stringify({ pres, cur });
  let moved = false;
  const lead = start[0];
  const move = (ev) => {
    const p = stagePoint(ev);
    let dx = p.x - p0.x, dy = p.y - p0.y;
    if (!moved && Math.hypot(dx, dy) * scale < 3) return;
    moved = true;
    if (ev.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
    const sn = ev.altKey ? { dx, dy, guides: [] } : snap({ ...lead.o, x: lead.x, y: lead.y }, dx, dy, ids);
    start.forEach((s) => { s.o.x = Math.round(s.x + sn.dx); s.o.y = Math.round(s.y + sn.dy); const el = slideEl.querySelector(`[data-id="${s.o.id}"]`); if (el) { el.style.left = `${s.o.x}px`; el.style.top = `${s.o.y}px`; } });
    drawSelection(sn.guides);
  };
  const up = () => {
    removeEventListener('mousemove', move); removeEventListener('mouseup', up);
    if (moved) { undoStack.push(snapshot); redoStack.length = 0; shell.markDirty(); drawSelection(); thumbLater(); }
  };
  addEventListener('mousemove', move); addEventListener('mouseup', up);
}
function startTransform(e, handle, p0) {
  e.preventDefault();
  const o = objById([...selected][0]); if (!o) return;
  const s0 = { x: o.x, y: o.y, w: o.w, h: o.h, rot: o.rot || 0 };
  const snapshot = JSON.stringify({ pres, cur });
  const keepRatio = o.type === 'image';
  const move = (ev) => {
    const p = stagePoint(ev);
    if (handle === 'rot') {
      const cx = s0.x + s0.w / 2, cy = s0.y + s0.h / 2;
      let a = (Math.atan2(p.y - cy, p.x - cx) * 180) / Math.PI + 90;
      if (ev.shiftKey || Math.abs(a % 90) < 4 || Math.abs(a % 90) > 86) a = Math.round(a / (ev.shiftKey ? 15 : 90)) * (ev.shiftKey ? 15 : 90);
      o.rot = Math.round(((a % 360) + 360) % 360);
    } else {
      let { x, y, w, h: hh } = s0;
      const dx = p.x - p0.x, dy = p.y - p0.y;
      if (handle.includes('e')) w = Math.max(10, s0.w + dx);
      if (handle.includes('s')) hh = Math.max(10, s0.h + dy);
      if (handle.includes('w')) { w = Math.max(10, s0.w - dx); x = s0.x + s0.w - w; }
      if (handle.includes('n')) { hh = Math.max(10, s0.h - dy); y = s0.y + s0.h - hh; }
      if ((keepRatio || ev.shiftKey) && handle.length === 2) { const r = s0.h / s0.w; if (Math.abs(dx) > Math.abs(dy)) { hh = w * r; if (handle.includes('n')) y = s0.y + s0.h - hh; } else { w = hh / r; if (handle.includes('w')) x = s0.x + s0.w - w; } }
      Object.assign(o, { x: Math.round(x), y: Math.round(y), w: Math.round(w), h: Math.round(hh) });
    }
    const el = slideEl.querySelector(`[data-id="${o.id}"]`);
    if (el) { Object.assign(el.style, { left: `${o.x}px`, top: `${o.y}px`, width: `${o.w}px`, height: `${o.h}px`, transform: o.rot ? `rotate(${o.rot}deg)` : '' }); if (o.type === 'shape') { const n = renderSlide({ ...slide(), objects: [o] }, 'edit').querySelector('.obj'); el.innerHTML = n.innerHTML; } }
    drawSelection();
  };
  const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); undoStack.push(snapshot); redoStack.length = 0; shell.markDirty(); renderStage(); thumbLater(); };
  addEventListener('mousemove', move); addEventListener('mouseup', up);
}

// ------------------------------------------------------------------ edición de texto
function onStageDbl(e) {
  const objEl = e.target.closest('.obj'); if (!objEl) return;
  const o = objById(objEl.dataset.id);
  if (o?.type === 'chart') { editChart(o); return; }
  if (o && (o.type === 'text' || o.type === 'shape' || o.type === 'table')) startEditing(o.id, true, e);
}
function startEditing(id, placeCaret = true, ev = null) {
  const o = objById(id); if (!o) return;
  editing = id; selected = new Set([id]);
  let el = slideEl.querySelector(`[data-id="${id}"]`);
  if (o.type === 'table') {
    el.querySelectorAll('td,th').forEach((c) => { c.contentEditable = 'true'; });
    const first = ev ? ev.target.closest('td,th') : el.querySelector('td,th');
    first?.focus();
    el.addEventListener('focusout', () => setTimeout(() => { if (editing === id && !el.contains(document.activeElement)) stopEditing(); }, 0));
    drawSelection(); reflect(); return;
  }
  let tin = el.querySelector('.tin');
  if (!tin) { renderStage(); el = slideEl.querySelector(`[data-id="${id}"]`); tin = el.querySelector('.tin'); }
  if (!tin) return;
  tin.contentEditable = 'true'; tin.spellcheck = true;
  tin.dataset.before = o.html || '';
  tin.focus();
  if (o.style?.list && !o.html) { document.execCommand('insertUnorderedList'); }
  if (placeCaret) {
    const s = document.getSelection(); s.removeAllRanges();
    let r = null;
    if (ev && document.caretRangeFromPoint) r = document.caretRangeFromPoint(ev.clientX, ev.clientY);
    if (!r || !tin.contains(r.startContainer)) { r = document.createRange(); r.selectNodeContents(tin); r.collapse(false); }
    s.addRange(r);
  }
  tin.oninput = () => { o.html = cleanHtml(tin.innerHTML); shell.markDirty(); thumbLater(); };
  tin.onkeydown = (e) => { if (e.key === 'Escape') { e.preventDefault(); stopEditing(); } e.stopPropagation(); if ((e.ctrlKey || e.metaKey) && ['b', 'i', 'u'].includes(e.key.toLowerCase())) { e.preventDefault(); document.execCommand({ b: 'bold', i: 'italic', u: 'underline' }[e.key.toLowerCase()]); o.html = cleanHtml(tin.innerHTML); } };
  drawSelection(); reflect();
}
const cleanHtml = (s) => (s.replace(/<br>$/, '').trim() === '' || s === '<ul><li><br></li></ul>' ? '' : s);
function stopEditing() {
  const id = editing; if (!id) return;
  const o = objById(id); editing = null;
  const el = slideEl?.querySelector(`[data-id="${id}"]`);
  if (o && el) {
    if (o.type === 'table') {
      const before = JSON.stringify(o.rows);
      const cells = [...el.querySelectorAll('td,th')]; let k = 0;
      o.rows.forEach((row) => row.forEach((c) => { if (!c.hidden) c.text = cells[k++]?.innerText.trim() ?? c.text; }));
      if (JSON.stringify(o.rows) !== before) { const after = clone(o.rows); o.rows = JSON.parse(before); commit('Editar tabla', () => { objById(id).rows = after; }); return; }
    } else {
      const tin = el.querySelector('.tin');
      const after = cleanHtml(tin?.innerHTML || '');
      const before = tin?.dataset.before ?? '';
      if (after !== before) { o.html = before; commit('Editar texto', () => { objById(id).html = after; }); return; }
    }
  }
  renderStage(); renderThumbs();
}

// ------------------------------------------------------------------ teclado y portapapeles
let clipboard = null;
function onKey(e) {
  if (!pres || $('.backstage:not([hidden])') || $('.modal-back') || $('.show')) return;
  const tag = document.activeElement?.tagName;
  if (editing || tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || document.activeElement?.isContentEditable) return;
  const mod = e.ctrlKey || e.metaKey;
  if ((e.key === 'Delete' || e.key === 'Backspace') && selected.size) { e.preventDefault(); deleteSelected(); return; }
  if (e.key.startsWith('Arrow') && selected.size) {
    e.preventDefault();
    const d = e.shiftKey ? 10 : 1;
    const [dx, dy] = { ArrowLeft: [-d, 0], ArrowRight: [d, 0], ArrowUp: [0, -d], ArrowDown: [0, d] }[e.key];
    if (!onKey.nudging) { undoStack.push(JSON.stringify({ pres, cur })); redoStack.length = 0; onKey.nudging = true; setTimeout(() => { onKey.nudging = false; }, 800); }
    selected.forEach((id) => { const o = objById(id); o.x += dx; o.y += dy; });
    shell.markDirty(); renderStage(); thumbLater(); return;
  }
  if (!selected.size && (e.key === 'ArrowDown' || e.key === 'PageDown')) { e.preventDefault(); goto(cur + 1); return; }
  if (!selected.size && (e.key === 'ArrowUp' || e.key === 'PageUp')) { e.preventDefault(); goto(cur - 1); return; }
  if (e.key === 'Escape') { select([]); return; }
  if (e.key === 'Enter' && selected.size === 1) { e.preventDefault(); startEditing([...selected][0]); return; }
  if (mod && e.key.toLowerCase() === 'a') { e.preventDefault(); select(slide().objects.map((o) => o.id)); return; }
  if (mod && e.key.toLowerCase() === 'd' && selected.size) { e.preventDefault(); duplicateObjects(); return; }
  if (e.key === 'F5') { e.preventDefault(); present(e.shiftKey ? cur : 0); return; }
  if (!mod && !e.altKey && e.key.length === 1 && selected.size === 1) {
    const o = objById([...selected][0]);
    if (o.type === 'text' || o.type === 'shape') { startEditing(o.id); }
  }
}
function onCopy(e, cut = false) {
  if (editing || !pres || !selected.size || ['INPUT', 'TEXTAREA'].includes(document.activeElement?.tagName) || document.activeElement?.isContentEditable) return;
  clipboard = slide().objects.filter((o) => selected.has(o.id)).map(clone);
  e.clipboardData.setData('text/plain', clipboard.map((o) => (o.html ? new DOMParser().parseFromString(o.html, 'text/html').body.innerText : '')).join('\n'));
  e.clipboardData.setData('application/x-impresiona', JSON.stringify(clipboard));
  e.preventDefault();
  if (cut) deleteSelected();
}
function onPaste(e) {
  if (editing || !pres || ['INPUT', 'TEXTAREA'].includes(document.activeElement?.tagName) || document.activeElement?.isContentEditable) return;
  const data = e.clipboardData.getData('application/x-impresiona');
  if (data) { e.preventDefault(); pasteObjects(JSON.parse(data)); return; }
  const img = [...e.clipboardData.items].find((i) => i.type.startsWith('image/'));
  if (img) { e.preventDefault(); const rd = new FileReader(); rd.onload = () => addImage(rd.result); rd.readAsDataURL(img.getAsFile()); return; }
  const text = e.clipboardData.getData('text/plain');
  if (text) { e.preventDefault(); addTextBox(esc(text).replace(/\n/g, '<br>')); }
}
function pasteObjects(list) {
  const fresh = list.map((o) => ({ ...clone(o), id: uid(), x: o.x + 20, y: o.y + 20, role: undefined }));
  commit('Pegar', () => { slide().objects.push(...fresh); });
  select(fresh.map((o) => o.id));
}
function deleteSelected() { const ids = new Set(selected); commit('Eliminar', () => { slide().objects = slide().objects.filter((o) => !ids.has(o.id)); selected.clear(); }); }
function duplicateObjects() { pasteObjects(slide().objects.filter((o) => selected.has(o.id))); }

// ------------------------------------------------------------------ insertar
function center(w, hh) { return { x: Math.round((pres.size.w - w) / 2), y: Math.round((pres.size.h - hh) / 2), w, h: hh }; }
function addTextBox(html = '') {
  const o = { id: uid(), type: 'text', ...center(420, 70), rot: 0, html, style: { size: 24 } };
  commit('Cuadro de texto', () => slide().objects.push(o));
  select([o.id]);
  if (!html) setTimeout(() => startEditing(o.id), 30);
}
function addShape(shape) {
  const sz = shape === 'line' ? [300, 20] : [220, shape === 'arrow' ? 120 : 160];
  const o = { id: uid(), type: 'shape', shape, ...center(...sz), rot: 0, html: '', style: { valign: 'middle', align: 'center' } };
  commit('Forma', () => slide().objects.push(o));
  select([o.id]);
}
function addImage(src) {
  const img = new Image();
  img.onload = () => {
    const max = Math.min(pres.size.w * 0.7, img.naturalWidth), r = img.naturalHeight / img.naturalWidth;
    let w = max, hh = w * r; if (hh > pres.size.h * 0.8) { hh = pres.size.h * 0.8; w = hh / r; }
    const o = { id: uid(), type: 'image', src, ...center(Math.round(w), Math.round(hh)), rot: 0 };
    commit('Imagen', () => slide().objects.push(o)); select([o.id]);
  };
  img.src = src;
}
async function insertImage() { const f = await api?.openImage(); if (!f) return; const b64 = btoa(Array.from(f.bytes, (c) => String.fromCharCode(c)).join('')); addImage(`data:${f.mime};base64,${b64}`); }
async function insertTable() {
  const v = await form('Insertar tabla', [{ name: 'cols', label: 'Columnas', type: 'number', value: 3, min: 1, max: 12 }, { name: 'rows', label: 'Filas', type: 'number', value: 4, min: 1, max: 20 }], 'Insertar');
  if (!v) return;
  const rows = Array.from({ length: v.rows }, (_, r) => Array.from({ length: v.cols }, (_, c) => ({ text: r === 0 ? `Encabezado ${c + 1}` : '', rs: 1, cs: 1 })));
  const o = { id: uid(), type: 'table', ...center(Math.min(pres.size.w - 120, v.cols * 160), v.rows * 44), rot: 0, rows, style: { size: 16 } };
  commit('Tabla', () => slide().objects.push(o)); select([o.id]);
}

// ------------------------------------------------------------------ gráficos
function insertChart(opts) {
  const data = typeof opts.data === 'string' ? parseChartData(opts.data) : opts.data;
  const o = { id: uid(), type: 'chart', ...center(Math.min(640, pres.size.w - 120), Math.min(380, pres.size.h - 120)), rot: 0, chartType: opts.type || 'column', title: opts.title || '', data };
  commit('Gráfico', () => slide().objects.push(o)); select([o.id]);
  return o;
}
async function chartDialog(o) {
  const v = await form(o ? 'Editar gráfico' : 'Insertar gráfico', [
    { name: 'type', label: 'Tipo', type: 'select', value: o?.chartType || 'column', options: CHART_TYPES },
    { name: 'title', label: 'Título', value: o?.title || '' },
    { name: 'data', label: 'Datos (la primera fila son los nombres de las series; la primera columna, las categorías)', type: 'textarea', rows: 7, value: o ? chartDataText(o.data) : 'Categoría,Serie 1,Serie 2\nEne,10,14\nFeb,12,18\nMar,9,20\nAbr,15,22' },
  ], o ? 'Aplicar' : 'Insertar');
  return v;
}
async function addChartUI() { const v = await chartDialog(null); if (v) insertChart({ type: v.type, title: v.title, data: v.data }); }
async function editChart(o) { const v = await chartDialog(o); if (!v) return; commit('Editar gráfico', () => { const t = objById(o.id); t.chartType = v.type; t.title = v.title; t.data = parseChartData(v.data); }); }

// ------------------------------------------------------------------ patrón de diapositivas
function toggleMaster(on = !masterMode) {
  if (editing) stopEditing();
  masterMode = on; selected.clear();
  pres.master ||= { objects: [], footer: '', showNumber: false, showDate: false };
  shell.workspace.classList.toggle('master-mode', on);
  shell.ribbon.setOn('master', on);
  renderAll();
  if (on) toast('Vista Patrón: inserta logotipos, formas o texto que deban repetirse. Pulsa de nuevo para salir.');
}
async function footerDialog() {
  pres.master ||= { objects: [], footer: '', showNumber: false, showDate: false };
  const m = pres.master;
  const v = await form('Encabezado y pie de página', [{ name: 'footer', label: 'Pie de página (en todas las diapositivas)', value: m.footer }, { name: 'showNumber', label: 'Número de diapositiva', type: 'checkbox', value: m.showNumber }, { name: 'showDate', label: 'Fecha', type: 'checkbox', value: m.showDate }], 'Aplicar a todas');
  if (!v) return;
  commit('Pie de página', () => Object.assign(pres.master, v));
}

// ------------------------------------------------------------------ clasificador y secciones
function toggleSorter(on = !sorterMode) {
  if (editing) stopEditing();
  sorterMode = on;
  shell.workspace.classList.toggle('sorting', on);
  shell.ribbon.setOn('sorter', on);
  if (on) renderSorter(); else { sorterEl?.replaceChildren(); renderAll(); }
}
let sorterEl = null;
function renderSorter() {
  if (!sorterEl) { sorterEl = h('div', { class: 'sorter' }); shell.workspace.append(sorterEl); }
  sorterEl.replaceChildren();
  const tw = 220, ts = tw / pres.size.w;
  let grid = null;
  pres.slides.forEach((s, i) => {
    if (s.section || i === 0) {
      sorterEl.append(h('div', { class: 'sec-head' }, h('b', {}, s.section || 'Sección predeterminada'), h('small', {}, `${countSection(i)} diapositivas`), h('button', { class: 'rb sm', title: 'Cambiar nombre', html: icon('textbox', 15), onclick: () => renameSection(i) })));
      if (!s.section) sorterEl.lastChild.classList.add('default');
      grid = h('div', { class: 'sgrid' }); sorterEl.append(grid);
    }
    const mini = renderSlide(s, 'thumb'); mini.style.transform = `scale(${ts})`; mini.style.transformOrigin = '0 0';
    const item = h('div', { class: `sitem${i === cur ? ' on' : ''}`, draggable: 'true', tabindex: 0 }, h('div', { class: 'tframe', style: { width: `${tw}px`, height: `${pres.size.h * ts}px` } }, mini), h('span', {}, String(i + 1)));
    item.addEventListener('click', () => { cur = i; renderSorter(); });
    item.addEventListener('dblclick', () => { cur = i; toggleSorter(false); });
    item.addEventListener('dragstart', (e) => e.dataTransfer.setData('text/x-slide', String(i)));
    item.addEventListener('dragover', (e) => { e.preventDefault(); item.classList.add('over'); });
    item.addEventListener('dragleave', () => item.classList.remove('over'));
    item.addEventListener('drop', (e) => { e.preventDefault(); const from = Number(e.dataTransfer.getData('text/x-slide')); if (Number.isNaN(from) || from === i) return; commit('Mover diapositiva', () => { const [m] = pres.slides.splice(from, 1); pres.slides.splice(i, 0, m); cur = i; }); renderSorter(); });
    grid.append(item);
  });
}
function countSection(start) { let n = 1; for (let k = start + 1; k < pres.slides.length && !pres.slides[k].section; k++) n++; return n; }
async function addSection() {
  const v = await form('Agregar sección', [{ name: 'n', label: 'Nombre de la sección (empieza en la diapositiva actual)', value: 'Nueva sección' }], 'Agregar');
  if (v?.n?.trim()) { commit('Sección', () => { pres.slides[cur].section = v.n.trim(); }); if (sorterMode) renderSorter(); }
}
async function renameSection(i) {
  const v = await form('Cambiar nombre de sección', [{ name: 'n', label: 'Nombre (vacío para quitar la sección)', value: pres.slides[i].section || '' }], 'Aceptar');
  if (!v) return;
  commit('Sección', () => { if (v.n.trim()) pres.slides[i].section = v.n.trim(); else delete pres.slides[i].section; });
  if (sorterMode) renderSorter();
}

// ------------------------------------------------------------------ panel de animación
let animPanel = null;
function toggleAnimPanel(force) {
  if (!animPanel) { animPanel = h('aside', { class: 'animpanel', hidden: true }); stageWrap.parentElement.append(animPanel); }
  animPanel.hidden = force !== undefined ? !force : !animPanel.hidden;
  shell.ribbon.setOn('animPanel', !animPanel.hidden);
  renderAnimPanel();
}
function renderAnimPanel() {
  if (!animPanel || animPanel.hidden) return;
  const s = slide();
  const items = s.objects.filter((o) => o.anim && o.anim.type !== 'none').sort((a, b) => (a.anim.order || 0) - (b.anim.order || 0));
  const label = (o) => { const t = o.html ? new DOMParser().parseFromString(o.html, 'text/html').body.textContent.trim() : ''; return t ? t.slice(0, 26) : ({ image: 'Imagen', shape: 'Forma', chart: 'Gráfico', table: 'Tabla', text: 'Cuadro de texto' }[o.type]); };
  const head = h('div', { class: 'sp-head' }, h('b', {}, 'Panel de animación'), h('button', { class: 'rb sm', html: icon('close', 16), onclick: () => toggleAnimPanel(false) }));
  const body = h('div', { class: 'ap-body' });
  if (!items.length) body.append(h('p', { class: 'sp-empty' }, 'Selecciona un objeto y elige una animación en la pestaña Animaciones.'));
  items.forEach((o, k) => {
    const start = h('select', {}, ...[['click', 'Al hacer clic'], ['with', 'Con la anterior'], ['after', 'Después de la anterior']].map(([v, l]) => h('option', { value: v, selected: (o.anim.start || 'click') === v ? true : null }, l)));
    const dur = h('input', { type: 'number', step: 0.1, min: 0.1, max: 10, value: o.anim.dur || 0.6, title: 'Duración (s)' });
    const delay = h('input', { type: 'number', step: 0.1, min: 0, max: 30, value: o.anim.delay || 0, title: 'Retraso (s)' });
    const upd = (patch) => commit('Animación', () => { Object.assign(objById(o.id).anim, patch); });
    start.addEventListener('change', () => upd({ start: start.value }));
    dur.addEventListener('change', () => upd({ dur: Number(dur.value) }));
    delay.addEventListener('change', () => upd({ delay: Number(delay.value) }));
    const move = (d) => { const j = k + d; if (j < 0 || j >= items.length) return; commit('Orden', () => { const a = objById(items[k].id).anim, b = objById(items[j].id).anim; [a.order, b.order] = [b.order, a.order]; }); };
    body.append(h('div', { class: `ap-item${selected.has(o.id) ? ' on' : ''}`, onclick: (e) => { if (e.target.tagName === 'DIV' || e.target.tagName === 'B') select([o.id]); } },
      h('div', { class: 'ap-row' }, h('span', { class: 'anim-num' }, String(k + 1)), h('b', {}, label(o)), h('small', {}, ANIM_NAMES[o.anim.type] || o.anim.type)),
      h('div', { class: 'ap-row' }, start, h('label', {}, 'Dur. ', dur), h('label', {}, 'Ret. ', delay)),
      h('div', { class: 'ap-row end' }, h('button', { class: 'rb sm', title: 'Antes', html: icon('chevronDown', 15).replace('M6 9l6 6 6-6', 'M6 15l6-6 6 6'), onclick: () => move(-1) }), h('button', { class: 'rb sm', title: 'Después', html: icon('chevronDown', 15), onclick: () => move(1) }), h('button', { class: 'rb sm', title: 'Quitar', html: icon('trash', 15), onclick: () => commit('Quitar animación', () => { delete objById(o.id).anim; }) }))));
  });
  body.append(h('button', { class: 'btn', style: { marginTop: '8px' }, onclick: () => present(cur) }, 'Reproducir desde esta diapositiva'));
  animPanel.replaceChildren(head, body);
}

// ------------------------------------------------------------------ exportar imágenes
async function exportImages() {
  if (!api?.renderSlides) return;
  toast('Creando las imágenes…');
  const r = await api.renderSlides({ pres: JSON.parse(JSON.stringify(pres)), base: 'Diapositiva' });
  if (r) toast(`${r.files.length} imágenes guardadas`, { action: { label: 'Mostrar', run: () => api.showInFolder(r.files[0]) } });
}

// ------------------------------------------------------------------ diapositivas
function newSlide(layout = 'content') { return { id: uid(), layout, bg: null, notes: '', transition: { type: 'none', dur: 0.6 }, objects: layoutObjects(layout, pres.size.w, pres.size.h, uid) }; }
function addSlide(layout) {
  const lay = layout || (slide()?.layout === 'title' ? 'content' : slide()?.layout || 'content');
  commit('Nueva diapositiva', () => { pres.slides.splice(cur + 1, 0, newSlide(lay === 'blank' && !layout ? 'content' : lay)); cur += 1; selected.clear(); });
}
function duplicateSlide() { commit('Duplicar diapositiva', () => { const c = clone(slide()); c.id = uid(); c.objects.forEach((o) => { o.mk = o.mk || o.id; o.id = uid(); }); pres.slides.splice(cur + 1, 0, c); cur += 1; }); }
function deleteSlide() {
  if (pres.slides.length === 1) { commit('Eliminar diapositiva', () => { pres.slides[0] = newSlide('title'); }); return; }
  commit('Eliminar diapositiva', () => { pres.slides.splice(cur, 1); cur = Math.min(cur, pres.slides.length - 1); selected.clear(); });
}
function setLayout(layout) {
  commit('Diseño', () => {
    const s = slide();
    const old = s.objects.filter((o) => o.role);
    const fresh = layoutObjects(layout, pres.size.w, pres.size.h, uid);
    fresh.forEach((f) => { const prev = old.find((o) => o.role === f.role && !o.used); if (prev) { f.html = prev.html; prev.used = true; } });
    s.objects = [...fresh, ...s.objects.filter((o) => !o.role)];
    s.layout = layout;
  });
}

// ------------------------------------------------------------------ formato
const selObjs = () => slide().objects.filter((o) => selected.has(o.id));
function textCmd(cmd, val) {
  if (editing) {
    document.execCommand('styleWithCSS', false, true);
    document.execCommand(cmd, false, val);
    const o = objById(editing); const tin = slideEl.querySelector(`[data-id="${editing}"] .tin`); if (o && tin) { o.html = cleanHtml(tin.innerHTML); shell.markDirty(); }
    reflect(); return true;
  }
  return false;
}
function styleObjs(patch) {
  if (!selected.size) { toast('Selecciona primero un objeto.'); return; }
  commit('Formato', () => selObjs().forEach((o) => { o.style = { ...(o.style || {}), ...patch }; for (const k of Object.keys(patch)) if (patch[k] === undefined) delete o.style[k]; }));
}
function applyFontSize(size) {
  if (editing) {
    document.execCommand('styleWithCSS', false, true); document.execCommand('fontSize', false, '7');
    slideEl.querySelectorAll('[data-id] .tin span[style*="xxx-large"], [data-id] .tin font[size="7"]').forEach((el) => { const sp = document.createElement('span'); sp.style.fontSize = `${size}pt`; sp.innerHTML = el.innerHTML; el.replaceWith(sp); });
    const o = objById(editing); const tin = slideEl.querySelector(`[data-id="${editing}"] .tin`); if (o && tin) { o.html = cleanHtml(tin.innerHTML); shell.markDirty(); }
    return;
  }
  styleObjs({ size: Number(size) });
}
function arrange(kind) {
  commit('Organizar', () => {
    const objs = slide().objects; const sel = objs.filter((o) => selected.has(o.id)); const rest = objs.filter((o) => !selected.has(o.id));
    if (kind === 'front') slide().objects = [...rest, ...sel];
    if (kind === 'back') slide().objects = [...sel, ...rest];
    const { w: W, h: H } = pres.size;
    sel.forEach((o) => {
      if (kind === 'left') o.x = 0; if (kind === 'hcenter') o.x = Math.round((W - o.w) / 2); if (kind === 'right') o.x = W - o.w;
      if (kind === 'top') o.y = 0; if (kind === 'vcenter') o.y = Math.round((H - o.h) / 2); if (kind === 'bottom') o.y = H - o.h;
    });
  });
}
function reflect() {
  if (!shell || !pres) return;
  const rb = shell.ribbon;
  const one = selected.size === 1 ? objById([...selected][0]) : null;
  const q = (c) => { try { return editing && document.queryCommandState(c); } catch { return false; } };
  rb.setOn('bold', editing ? q('bold') : one?.style?.bold);
  rb.setOn('italic', editing ? q('italic') : one?.style?.italic);
  rb.setOn('underline', q('underline'));
  const al = one?.style?.align || 'left';
  rb.setOn('alignLeft', al === 'left'); rb.setOn('alignCenter', al === 'center'); rb.setOn('alignRight', al === 'right');
  if (one) { rb.setValue('fontSize', String(one.style?.size || 20)); rb.setValue('fontName', one.style?.font || (one.role === 'title' ? theme().titleFont : theme().bodyFont)); }
  rb.setContextual('format', selected.size > 0);
  rb.setContextual('media', selObjs().some((o) => o.type === 'media'));
  rb.setContextual('camera', selObjs().some((o) => o.type === 'camera'));
  rb.setOn('useNarration', pres.useNarration !== false && pres.slides.some((s) => s.narration));
  rb.setOn('useTimings', pres.useTimings !== false && pres.slides.some((s) => s.advanceAfter > 0));
  $$('.themegal .gi').forEach((g) => g.classList.toggle('on', g.dataset.theme === pres.theme));
  $$('.transgal .gi').forEach((g) => g.classList.toggle('on', g.dataset.t === (slide()?.transition?.type || 'none')));
  $$('.animgal .gi').forEach((g) => g.classList.toggle('on', g.dataset.a === (one?.anim?.type || 'none')));
}

function onStageContext(e) {
  e.preventDefault();
  const objEl = e.target.closest('.obj');
  if (objEl && !selected.has(objEl.dataset.id)) select([objEl.dataset.id]);
  const anchor = h('div', { style: { position: 'fixed', left: `${e.clientX}px`, top: `${e.clientY}px`, width: '1px', height: '1px' } }); document.body.append(anchor);
  const items = selected.size ? [
    { label: 'Cortar', icon: 'cut', run: () => document.execCommand('cut') }, { label: 'Copiar', icon: 'copy', run: () => document.execCommand('copy') }, { label: 'Duplicar', icon: 'slideDup', kbd: 'Ctrl+D', run: duplicateObjects },
    '-', { label: 'Traer al frente', icon: 'front', run: () => arrange('front') }, { label: 'Enviar al fondo', icon: 'back', run: () => arrange('back') },
    '-', { label: 'Eliminar', icon: 'trash', run: deleteSelected },
  ] : [
    { label: 'Pegar', icon: 'paste', run: () => { if (clipboard) pasteObjects(clipboard); } }, { label: 'Nueva diapositiva', icon: 'slideAdd', run: () => addSlide() }, { label: 'Fondo…', icon: 'background', run: () => colorPicker(anchor, (c) => commit('Fondo', () => { slide().bg = c === 'transparent' ? null : c; slide().bgImage = null; })) },
  ];
  dropdown(anchor, items); setTimeout(() => anchor.remove(), 0);
}

// ------------------------------------------------------------------ presentar
function prepareShow() { if (editing) stopEditing(); if (masterMode) toggleMaster(false); }
function present(from = 0) {
  prepareShow();
  runShow({ renderSlide, getPres: () => pres, from, onEnd: (i) => goto(Math.min(i, pres.slides.length - 1)) });
}
async function presenter(from = cur) {
  if (editing) stopEditing();
  if (!api?.openAudience) { present(from); return; }
  const r = await api.openAudience({ pres: JSON.parse(JSON.stringify(pres)), i: from, step: 0 });
  if (r?.sameScreen) toast('Solo hay una pantalla: la presentación se abrió en otra ventana. Conecta un proyector para verla ahí.');
  runPresenter({ renderSlide, getPres: () => pres, from, api, onEnd: (i) => goto(Math.min(i, pres.slides.length - 1)) });
}

// ------------------------------------------------------------------ PDF / imprimir
function buildPrint() {
  const root = $('#printRoot'); root.replaceChildren();
  for (const s of pres.slides) root.append(h('div', { class: 'ppage', style: { width: `${pres.size.w}px`, height: `${pres.size.h}px` } }, renderSlide(s, 'print')));
  let st = $('#pagestyle'); if (!st) { st = h('style', { id: 'pagestyle' }); document.head.append(st); }
  st.textContent = `@page { size: ${pres.size.w / 96}in ${pres.size.h / 96}in; margin: 0; }`;
}

// ------------------------------------------------------------------ plantillas
function makePres(tpl) {
  const p = { size: { w: 960, h: 540 }, theme: { negocio: 'noche', clase: 'jade', proyecto: 'indigo', blank: 'claro' }[tpl] || 'claro', slides: [] };
  pres = p;
  const S = (layout, texts = [], extra = {}) => { const s = newSlide(layout); const ph = s.objects.filter((o) => o.role); texts.forEach((t, k) => { if (ph[k]) ph[k].html = t; }); return Object.assign(s, extra); };
  const ul = (...items) => `<ul>${items.map((x) => `<li>${esc(x)}</li>`).join('')}</ul>`;
  if (tpl === 'negocio') {
    p.slides = [
      S('title', ['Plan de negocio 2027', 'Tu empresa · Acapulco, Guerrero'], { transition: { type: 'fade', dur: 0.8 } }),
      S('content', ['El problema', ul('Qué necesidad no está resuelta', 'A quién le afecta y cuánto le cuesta', 'Por qué ahora es el momento')], { transition: { type: 'push', dur: 0.6 } }),
      S('two', ['Nuestra propuesta', ul('Producto o servicio', 'Lo que nos hace diferentes'), ul('Precio y modelo de ingresos', 'Canales de venta')], { transition: { type: 'push', dur: 0.6 } }),
      S('section', ['Gracias', 'Contacto: tu correo · tu teléfono'], { transition: { type: 'zoom', dur: 0.7 } }),
    ];
  } else if (tpl === 'clase') {
    p.slides = [
      S('title', ['El ciclo del agua', 'Ciencias Naturales · 5.º grado']),
      S('content', ['¿Qué vamos a aprender?', ul('Las etapas del ciclo del agua', 'Por qué es importante para la vida', 'Cómo cuidar el agua en casa')]),
      S('content', ['Etapas', ul('Evaporación', 'Condensación', 'Precipitación', 'Escurrimiento e infiltración')]),
      S('titleOnly', ['Actividad en equipo']),
    ];
  } else if (tpl === 'proyecto') {
    p.slides = [
      S('title', ['Avance del proyecto', 'Reunión semanal']),
      S('content', ['Lo que logramos', ul('Entregable 1 terminado', 'Pruebas con 10 usuarios', 'Presupuesto dentro de lo planeado')]),
      S('content', ['Siguientes pasos', ul('Corregir los hallazgos de las pruebas', 'Preparar el lanzamiento', 'Definir responsables y fechas')]),
    ];
  } else p.slides = [S('title')];
  // plantillas en el idioma de la interfaz
  for (const s of p.slides) for (const o of s.objects) if (o.html) o.html = localizeHtml(o.html);
  return p;
}
const tplThumb = (bg, fg, acc) => `<div style="height:100%;background:${bg};display:flex;flex-direction:column;justify-content:center;align-items:center;gap:6px"><div style="width:62%;height:9px;border-radius:3px;background:${fg}"></div><div style="width:40%;height:5px;border-radius:3px;background:${acc}"></div></div>`;

// ------------------------------------------------------------------ adaptador
const adapter = {
  i18nSkip: '.stage .slide, .thumb .slide, .sorter .slide, .show-stage, .presenter .slide, .tin',
  exts: ['pptx', 'odp'],
  writable: ['pptx', 'odp'],
  saveExt: 'pptx',
  docIcon: 'slides',
  homeHint: 'Crea presentaciones con temas, transiciones y animaciones.',
  templates: [
    { id: 'blank', name: 'Presentación en blanco', blank: true, thumb: tplThumb('#fff', '#1D2030', '#DC3A63') },
    { id: 'negocio', name: 'Plan de negocio', thumb: tplThumb('linear-gradient(135deg,#1F2337,#151827)', '#fff', '#FFB547') },
    { id: 'clase', name: 'Clase', thumb: tplThumb('#F0FAF6', '#0A4A3A', '#0E9C74') },
    { id: 'proyecto', name: 'Avance de proyecto', thumb: tplThumb('linear-gradient(135deg,#5B48E8,#2A1C8F)', '#fff', '#FFD38A') },
  ],
  exportFormats: [{ ext: 'pptx', ic: 'slides', label: 'Presentación de PowerPoint (.pptx)', desc: 'Para abrir en otros programas de presentaciones.' }, { ext: 'odp', ic: 'slides', label: 'Presentación ODF (.odp)', desc: 'Formato abierto, compatible con LibreOffice Impress.' }, { ic: 'image', label: 'Imágenes PNG', desc: 'Una imagen por diapositiva, para redes sociales o documentos.', run: () => { shell.hideBackstage(); exportImages(); } }],
  async newDoc(tpl) { makePres(tpl); cur = 0; selected.clear(); editing = null; undoStack.length = 0; redoStack.length = 0; fitMode = true; renderAll(); },
  async load(f) {
    if (f.ext === 'odp') {
      const r = await readOdp(f.bytes, uid);
      pres = r.pres; if (!pres.slides.length) pres.slides = [newSlide('title')];
      cur = 0; selected.clear(); editing = null; masterMode = false; undoStack.length = 0; redoStack.length = 0; fitMode = true;
      renderAll(); return { warnings: r.warnings, props: r.props };
    }
    const r = await readPptx(f.bytes, uid);
    pres = r.pres; if (!pres.slides.length) pres.slides = [newSlide('title')];
    cur = 0; selected.clear(); editing = null; undoStack.length = 0; redoStack.length = 0; fitMode = true;
    renderAll();
    return { warnings: r.warnings, props: await readCoreProps(f.bytes) };
  },
  async serialize(ext) { if (editing) stopEditing(); if (ext === 'odp') return writeOdp(pres, shell.docProps()); return writePptx(pres, null, shell.docProps()); },
  snapshot() { return JSON.stringify(pres); },
  async restore(s) { pres = JSON.parse(s); cur = 0; selected.clear(); renderAll(); },
  isPristine() { return !shell.doc.dirty && pres.slides.length === 1 && pres.slides[0].objects.every((o) => !o.html); },
  pdfOptions() { return { pageSize: { width: pres.size.w / 96, height: pres.size.h / 96 } }; },
  beforePrint() { if (editing) stopEditing(); buildPrint(); },
  afterPrint() { $('#printRoot').replaceChildren(); },
  focus() {},
  stats() { return [['Diapositivas', pres.slides.length], ['Tamaño', pres.size.w / pres.size.h > 1.5 ? 'Panorámica 16:9' : 'Estándar 4:3'], ['Tema', theme().name]]; },
  onDropFile(file) { if (file.type.startsWith('image/')) { const rd = new FileReader(); rd.onload = () => addImage(rd.result); rd.readAsDataURL(file); return true; } return false; },
  onCommand,
};

async function onCommand(cmd, arg) {
  switch (cmd) {
    case 'undo': return undo();
    case 'redo': return redo();
    case 'selectAll': if (editing) { document.execCommand('selectAll'); return; } return select(slide().objects.map((o) => o.id));
    case 'cut': case 'copy': document.execCommand(cmd); return;
    case 'paste': if (clipboard && !editing) pasteObjects(clipboard); else toast('Usa Ctrl+V para pegar.'); return;
    case 'newSlide': return addSlide();
    case 'duplicateSlide': return duplicateSlide();
    case 'deleteSlide': return deleteSlide();
    case 'present': return present(0);
    case 'insertVideo': return F3.insertMedia('video');
    case 'insertCamera': return F4.insertCamera();
    case 'cameraOptions': return F4.cameraOptions(selObjs()[0]);
    case 'recordShow': return F4.recordShow();
    case 'useNarration': return F4.setUseNarration(pres.useNarration === false);
    case 'clearNarration': return F4.clearNarration(true);
    case 'insertAudio': return F3.insertMedia('audio');
    case 'captions': return F3.captionsDialog(selObjs()[0]);
    case 'mediaOptions': return F3.mediaOptions(selObjs()[0]);
    case 'pathDistance': return F3.pathDistance(selObjs());
    case 'rehearse': return F3.rehearse();
    case 'useTimings': return F3.setUseTimings(pres.useTimings === false);
    case 'clearTimings': return F3.clearTimings();
    case 'exportVideo': return F3.exportVideoDialog();
    case 'presentCurrent': return present(cur);
    case 'textbox': return addTextBox();
    case 'insertImage': return insertImage();
    case 'insertTable': return insertTable();
    case 'bold': if (!textCmd('bold')) styleObjs({ bold: !objById([...selected][0])?.style?.bold || undefined }); return;
    case 'italic': if (!textCmd('italic')) styleObjs({ italic: !objById([...selected][0])?.style?.italic || undefined }); return;
    case 'underline': textCmd('underline'); return;
    case 'strike': textCmd('strikeThrough'); return;
    case 'bullets': if (!textCmd('insertUnorderedList')) styleObjs({ list: !objById([...selected][0])?.style?.list || undefined }); return;
    case 'numbers': textCmd('insertOrderedList'); return;
    case 'alignLeft': if (!textCmd('justifyLeft')) styleObjs({ align: 'left' }); return;
    case 'alignCenter': if (!textCmd('justifyCenter')) styleObjs({ align: 'center' }); return;
    case 'alignRight': if (!textCmd('justifyRight')) styleObjs({ align: 'right' }); return;
    case 'fontName': if (!textCmd('fontName', arg)) styleObjs({ font: arg }); return;
    case 'fontSize': return applyFontSize(arg);
    case 'grow': return applyFontSize(SIZES.find((s) => s > (objById([...selected][0])?.style?.size || 20)) || 96);
    case 'shrink': return applyFontSize([...SIZES].reverse().find((s) => s < (objById([...selected][0])?.style?.size || 20)) || 10);
    case 'textColor': if (!textCmd('foreColor', arg)) styleObjs({ color: arg === 'transparent' ? undefined : arg }); return;
    case 'fill': return styleObjs({ fill: arg === 'transparent' ? null : arg });
    case 'stroke': return styleObjs({ stroke: arg === 'transparent' ? null : arg, strokeW: arg === 'transparent' ? 0 : (objById([...selected][0])?.style?.strokeW || 2) });
    case 'front': case 'back': case 'left': case 'hcenter': case 'right': case 'top': case 'vcenter': case 'bottom': return arrange(cmd);
    case 'delete': return deleteSelected();
    case 'background': return commit('Fondo', () => { slide().bg = arg === 'transparent' ? null : arg; slide().bgImage = null; });
    case 'bgAll': return commit('Fondo', () => { const bg = slide().bg; pres.slides.forEach((s) => { s.bg = bg; }); });
    case 'size169': return setSize(960, 540);
    case 'size43': return setSize(720, 540);
    case 'applyTransAll': return commit('Transición', () => { const t = clone(slide().transition || { type: 'none', dur: 0.6 }); pres.slides.forEach((s) => { s.transition = clone(t); }); toast('Transición aplicada a todas las diapositivas'); });
    case 'transDur': { const v = await form('Duración de la transición', [{ name: 'd', label: 'Segundos', type: 'number', step: 0.1, value: slide().transition?.dur || 0.6, min: 0.1, max: 5 }]); if (v) commit('Duración', () => { slide().transition = { ...(slide().transition || { type: 'fade' }), dur: v.d }; }); return; }
    case 'previewTrans': return present(cur);
    case 'chart': return addChartUI();
    case 'master': return toggleMaster();
    case 'footer': return footerDialog();
    case 'sorter': return toggleSorter();
    case 'addSection': return addSection();
    case 'renameSection': { let i = cur; while (i > 0 && !pres.slides[i].section) i--; return renameSection(i); }
    case 'animPanel': return toggleAnimPanel();
    case 'presenter': return presenter(cur);
    case 'exportImages': return exportImages();
    case 'exportOdp': return shell.saveAsFormat('odp', 'Presentación ODF');
    case 'zoomIn': fitMode = false; scale = Math.min(4, scale * 1.15); fit(); return;
    case 'zoomOut': fitMode = false; scale = Math.max(0.1, scale / 1.15); fit(); return;
    case 'zoomReset': fitMode = true; fit(); return;
    case 'find': case 'findText': return findText(arg);
    case 'replace': return findText(null, true);
    default: return undefined;
  }
}
function setSize(w, hh) {
  if (pres.size.w === w && pres.size.h === hh) return;
  commit('Tamaño de diapositiva', () => {
    const k = w / pres.size.w;
    pres.slides.forEach((s) => s.objects.forEach((o) => { o.x = Math.round(o.x * k); o.w = Math.round(o.w * k); }));
    pres.size = { w, h: hh };
  });
  fitMode = true; fit();
}
async function findText(q, replace = false) {
  const v = q ? { q } : await form(replace ? 'Reemplazar' : 'Buscar', [{ name: 'q', label: 'Buscar' }, ...(replace ? [{ name: 'r', label: 'Reemplazar con' }] : [])]);
  if (!v || !v.q) return;
  const n = (s) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
  if (replace) {
    let count = 0;
    const re = new RegExp(v.q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    commit('Reemplazar', () => pres.slides.forEach((s) => s.objects.forEach((o) => { if (o.html && re.test(o.html)) { re.lastIndex = 0; o.html = o.html.replace(/>([^<]*)</g, (m, t) => `>${t.replace(re, () => { count++; return esc(v.r); })}<`); } re.lastIndex = 0; })));
    toast(`${count} reemplazos`); return;
  }
  for (let k = 1; k <= pres.slides.length; k++) {
    const i = (cur + k) % pres.slides.length;
    const o = pres.slides[i].objects.find((x) => x.html && n(new DOMParser().parseFromString(x.html, 'text/html').body.textContent).includes(n(v.q)));
    if (o) { goto(i); select([o.id]); return; }
  }
  toast(`No se encontró "${v.q}"`);
}

// ------------------------------------------------------------------ cinta
const layoutMenu = (apply) => LAYOUTS.map((l) => ({ label: l.name, icon: 'layout', run: () => (apply ? setLayout(l.id) : addSlide(l.id)) }));
const shapeMenu = () => [['rect', 'Rectángulo', 'rect'], ['roundRect', 'Rectángulo redondeado', 'rect'], ['ellipse', 'Elipse', 'ellipse'], ['triangle', 'Triángulo', 'triangle'], ['arrow', 'Flecha', 'arrow'], ['star', 'Estrella', 'star'], ['line', 'Línea', 'line']].map(([k, l, ic]) => ({ label: l, icon: ic, run: () => addShape(k) }));
const arrangeMenu = () => [{ label: 'Traer al frente', icon: 'front', run: () => arrange('front') }, { label: 'Enviar al fondo', icon: 'back', run: () => arrange('back') }, '-', { label: 'Alinear a la izquierda', icon: 'alignLeft', run: () => arrange('left') }, { label: 'Centrar horizontalmente', icon: 'alignCenter', run: () => arrange('hcenter') }, { label: 'Alinear a la derecha', icon: 'alignRight', run: () => arrange('right') }, { label: 'Alinear arriba', icon: 'alignObj', run: () => arrange('top') }, { label: 'Centrar verticalmente', icon: 'alignCenter', run: () => arrange('vcenter') }, { label: 'Alinear abajo', icon: 'footer', run: () => arrange('bottom') }];

function themeGallery() {
  const g = h('div', { class: 'gallery themegal' });
  THEMES.forEach((t) => {
    const b = h('button', { class: 'gi', 'data-theme': t.id, title: t.name, style: { background: t.band || t.bg, width: '86px' }, html: `<b style="font:600 15px ${t.titleFont},sans-serif;color:${t.title}">Aa</b><i style="display:block;width:46px;height:4px;border-radius:2px;background:${t.accent}"></i><small style="color:${t.text}">${t.name}</small>` });
    b.addEventListener('click', () => commit('Tema', () => { pres.theme = t.id; }));
    g.append(b);
  });
  return g;
}
function transGallery() {
  const g = h('div', { class: 'gallery transgal' });
  TRANSITIONS.forEach(([k, l]) => { const b = h('button', { class: 'gi', 'data-t': k, html: `${icon(k === 'none' ? 'close' : 'transition', 22)}<small>${l}</small>`, style: { alignItems: 'center', minWidth: '74px' } }); b.addEventListener('click', () => { commit('Transición', () => { slide().transition = { type: k, dur: slide().transition?.dur || 0.6 }; }); }); g.append(b); });
  return g;
}
function animGallery() {
  const g = h('div', { class: 'gallery animgal' });
  ANIMS.forEach(([k, l]) => { const b = h('button', { class: 'gi', 'data-a': k, html: `${icon(k === 'none' ? 'close' : 'animate', 22)}<small>${l}</small>`, style: { alignItems: 'center', minWidth: '74px' } }); b.addEventListener('click', () => { if (!selected.size) { toast('Selecciona un objeto para animarlo.'); return; } commit('Animación', () => { let n = Math.max(0, ...slide().objects.map((o) => o.anim?.order || 0)); selObjs().forEach((o) => { o.anim = k === 'none' ? undefined : { start: 'click', dur: 0.6, delay: 0, ...(o.anim || {}), type: k, order: o.anim?.order || ++n }; }); }); }); g.append(b); });
  return g;
}

const TABS = [
  { id: 'home', label: 'Inicio', groups: [
    { label: 'Portapapeles', items: [{ cmd: 'paste', label: 'Pegar', icon: 'paste', size: 'lg' }, { stack: [{ cmd: 'cut', label: 'Cortar', icon: 'cut' }, { cmd: 'copy', label: 'Copiar', icon: 'copy' }] }] },
    { label: 'Diapositivas', items: [{ cmd: 'newSlide', label: 'Nueva diapositiva', icon: 'slideAdd', size: 'lg', kbd: 'Ctrl+M' }, { stack: [{ cmd: 'layout', label: 'Diseño', icon: 'layout', showLabel: true, menu: () => layoutMenu(true) }, { cmd: 'duplicateSlide', label: 'Duplicar', icon: 'slideDup', showLabel: true }, { cmd: 'deleteSlide', label: 'Eliminar', icon: 'slideDel', showLabel: true }] }] },
    { label: 'Fuente', items: [{ stack: [
      { row: [{ cmd: 'fontName', label: 'Fuente', select: { options: FONTS, width: '130px' } }, { cmd: 'fontSize', label: 'Tamaño', select: { options: SIZES.map(String), width: '58px' } }, { cmd: 'grow', label: 'Aumentar', icon: 'grow' }, { cmd: 'shrink', label: 'Reducir', icon: 'shrink' }] },
      { row: [{ cmd: 'bold', label: 'Negrita', icon: 'bold' }, { cmd: 'italic', label: 'Cursiva', icon: 'italic' }, { cmd: 'underline', label: 'Subrayado', icon: 'underline' }, { cmd: 'strike', label: 'Tachado', icon: 'strike' }, { cmd: 'textColor', label: 'Color de fuente', icon: 'textColor', color: true, colorValue: '#DC3A63', colorOpts: { allowNone: false } }] },
    ] }] },
    { label: 'Párrafo', items: [{ stack: [{ row: [{ cmd: 'bullets', label: 'Viñetas', icon: 'bullets' }, { cmd: 'numbers', label: 'Numeración', icon: 'numbers' }] }, { row: [{ cmd: 'alignLeft', label: 'Izquierda', icon: 'alignLeft' }, { cmd: 'alignCenter', label: 'Centrar', icon: 'alignCenter' }, { cmd: 'alignRight', label: 'Derecha', icon: 'alignRight' }] }] }] },
    { label: 'Dibujo', items: [{ cmd: 'shapes', label: 'Formas', icon: 'star', size: 'lg', menu: shapeMenu }, { cmd: 'arrange', label: 'Organizar', icon: 'front', size: 'lg', menu: arrangeMenu }, { stack: [{ cmd: 'fill', label: 'Relleno', icon: 'fill', color: true, colorValue: '#DC3A63' }, { cmd: 'stroke', label: 'Contorno', icon: 'rect', color: true, colorValue: '#1D2030' }] }] },
    { label: 'Edición', items: [{ stack: [{ cmd: 'find', label: 'Buscar', icon: 'find', showLabel: true }, { cmd: 'replace', label: 'Reemplazar', icon: 'replace', showLabel: true }] }] },
  ] },
  { id: 'insert', label: 'Insertar', groups: [
    { label: 'Diapositivas', items: [{ cmd: 'newSlideMenu', label: 'Nueva diapositiva', icon: 'slideAdd', size: 'lg', menu: () => layoutMenu(false) }] },
    { label: 'Tablas', items: [{ cmd: 'insertTable', label: 'Tabla', icon: 'table', size: 'lg' }] },
    { label: 'Imágenes', items: [{ cmd: 'insertImage', label: 'Imagen', icon: 'image', size: 'lg' }] },
    { label: 'Multimedia', items: [{ cmd: 'insertVideo', label: 'Video', icon: 'present', size: 'lg' }, { cmd: 'insertAudio', label: 'Audio', icon: 'notes', size: 'lg' }, { cmd: 'insertCamera', label: 'Cámara en vivo', icon: 'image', size: 'lg' }] },
    { label: 'Ilustraciones', items: [{ cmd: 'shapes', label: 'Formas', icon: 'star', size: 'lg', menu: shapeMenu }, { cmd: 'chart', label: 'Gráfico', icon: 'chart', size: 'lg' }] },
    { label: 'Encabezado y pie', items: [{ cmd: 'footer', label: 'Pie y número', icon: 'footer', size: 'lg' }] },
    { label: 'Texto', items: [{ cmd: 'textbox', label: 'Cuadro de texto', icon: 'textbox', size: 'lg' }] },
  ] },
  { id: 'design', label: 'Diseño', groups: [
    { label: 'Temas', items: [{ render: themeGallery }] },
    { label: 'Personalizar', items: [{ cmd: 'size', label: 'Tamaño de diapositiva', icon: 'pageSetup', size: 'lg', menu: () => [{ label: 'Panorámica (16:9)', run: () => onCommand('size169') }, { label: 'Estándar (4:3)', run: () => onCommand('size43') }] }, { cmd: 'background', label: 'Fondo', icon: 'background', size: 'lg', color: true, colorValue: '#FFFFFF', colorOpts: { noneLabel: 'Fondo del tema' } }, { cmd: 'bgAll', label: 'Fondo a todas', icon: 'slideDup', size: 'lg' }] },
  ] },
  { id: 'trans', label: 'Transiciones', groups: [
    { label: 'Vista previa', items: [{ cmd: 'previewTrans', label: 'Vista previa', icon: 'present', size: 'lg' }] },
    { label: 'Transición a esta diapositiva', items: [{ render: transGallery }] },
    { label: 'Intervalos', items: [{ cmd: 'transDur', label: 'Duración', icon: 'recover', size: 'lg' }, { cmd: 'applyTransAll', label: 'Aplicar a todas', icon: 'slideDup', size: 'lg' }] },
  ] },
  { id: 'anim', label: 'Animaciones', groups: [
    { label: 'Animación del objeto (entrada y salida)', items: [{ render: animGallery }] },
    { label: 'Avanzado', items: [{ cmd: 'animPanel', label: 'Panel de animación', icon: 'layout', size: 'lg' }, { cmd: 'pathDistance', label: 'Distancia de la trayectoria', icon: 'arrow', size: 'lg' }] },
  ] },
  { id: 'show', label: 'Presentación', groups: [
    { label: 'Iniciar', items: [{ cmd: 'present', label: 'Desde el principio', icon: 'present', size: 'lg', kbd: 'F5' }, { cmd: 'presentCurrent', label: 'Desde la actual', icon: 'presentCur', size: 'lg', kbd: 'Mayús+F5' }, { cmd: 'presenter', label: 'Vista del moderador', icon: 'notes', size: 'lg', kbd: 'Alt+F5' }] },
    { label: 'Grabar', items: [{ cmd: 'recordShow', label: 'Grabar con narración', icon: 'present', size: 'lg' }, { stack: [{ cmd: 'useNarration', label: 'Reproducir narración', icon: 'check', showLabel: true }, { cmd: 'clearNarration', label: 'Borrar narración', icon: 'trash', showLabel: true }] }] },
    { label: 'Configurar', items: [{ cmd: 'rehearse', label: 'Ensayar intervalos', icon: 'recover', size: 'lg' }, { stack: [{ cmd: 'useTimings', label: 'Usar intervalos', icon: 'check', showLabel: true }, { cmd: 'clearTimings', label: 'Borrar intervalos', icon: 'trash', showLabel: true }] }] },
    { label: 'Exportar', items: [{ cmd: 'exportImages', label: 'Imágenes PNG', icon: 'image', size: 'lg' }, { cmd: 'exportVideo', label: 'Video', icon: 'present', size: 'lg' }] },
  ] },
  { id: 'view', label: 'Vista', groups: [
    { label: 'Vistas', items: [{ cmd: 'sorter', label: 'Clasificador', icon: 'columns', size: 'lg' }, { cmd: 'master', label: 'Patrón de diapositivas', icon: 'layout', size: 'lg' }] },
    { label: 'Secciones', items: [{ stack: [{ cmd: 'addSection', label: 'Agregar sección', icon: 'plus', showLabel: true }, { cmd: 'renameSection', label: 'Cambiar nombre', icon: 'textbox', showLabel: true }] }] },
    { label: 'Zoom', items: [{ cmd: 'zoomReset', label: 'Ajustar', icon: 'zoomIn', size: 'lg' }, { cmd: 'zoomIn', label: 'Acercar', icon: 'zoomIn', size: 'lg' }, { cmd: 'zoomOut', label: 'Alejar', icon: 'zoomOut', size: 'lg' }] },
  ] },
  { id: 'camera', label: 'Cámara', contextual: true, groups: [
    { label: 'Cámara', items: [{ cmd: 'cameraOptions', label: 'Opciones de la cámara', icon: 'image', size: 'lg' }] },
  ] },
  { id: 'media', label: 'Reproducción', contextual: true, groups: [
    { label: 'Opciones', items: [{ cmd: 'mediaOptions', label: 'Opciones de reproducción', icon: 'present', size: 'lg' }, { cmd: 'captions', label: 'Subtítulos', icon: 'textbox', size: 'lg' }] },
  ] },
  { id: 'format', label: 'Formato de forma', contextual: true, groups: [
    { label: 'Estilos', items: [{ cmd: 'fill', label: 'Relleno', icon: 'fill', size: 'lg', color: true, colorValue: '#DC3A63' }, { cmd: 'stroke', label: 'Contorno', icon: 'rect', size: 'lg', color: true, colorValue: '#1D2030' }] },
    { label: 'Organizar', items: [{ cmd: 'front', label: 'Al frente', icon: 'front', size: 'lg' }, { cmd: 'back', label: 'Al fondo', icon: 'back', size: 'lg' }, { cmd: 'align', label: 'Alinear', icon: 'alignObj', size: 'lg', menu: arrangeMenu }, { cmd: 'delete', label: 'Eliminar', icon: 'trash', size: 'lg' }] },
  ] },
];

// ------------------------------------------------------------------ arranque
const MODE = new URLSearchParams(location.search);
if (MODE.get('audience') === '1' || MODE.get('render') === '1') {
  runAudience({ renderSlide, setPres: (p) => { pres = p; }, getPres: () => pres, api, render: MODE.get('render') === '1' });
} else {
F4.init({ getPres: () => pres, slide: () => slide(), commit, select, center, uid, runShow, renderSlide, prepareShow, goto: (i) => goto(Math.min(i, pres.slides.length - 1)) });
F3.init({ getPres: () => pres, slide: () => slide(), commit, select, center, uid, renderAll, runShow, renderSlide, prepareShow, goto: (i) => goto(Math.min(i, pres.slides.length - 1)), docName: () => shell.doc.name });
shell = new OfficeShell({ appId: 'impresiona', appName: 'Impresiona', logo: '../shared/img/impresiona.svg', tabs: TABS, adapter });
shell.ribbon.onRender = () => setTimeout(reflect, 0);
buildWorkspace(shell.workspace);
statusSlide = h('span', {}, '');
zoomCtl = h('div', { class: 'zoom' }, h('button', { title: 'Alejar', html: icon('minus', 14), onclick: () => onCommand('zoomOut') }), h('span', { style: { minWidth: '44px', textAlign: 'center' } }, '100%'), h('button', { title: 'Acercar', html: icon('plus', 14), onclick: () => onCommand('zoomIn') }), h('button', { class: 'sb-btn', title: 'Ajustar a la ventana', onclick: () => onCommand('zoomReset') }, 'Ajustar'));
zoomCtl.set = true;
shell.status.append(statusSlide, h('span', {}, 'Español (México)'), h('div', { class: 'grow' }), h('button', { class: 'sb-btn', html: `${icon('present', 14)} Presentar`, style: { display: 'flex', gap: '6px', alignItems: 'center' }, onclick: () => present(cur) }), zoomCtl);
const origFit = fit;
void origFit;
makePres('blank'); renderAll();
shell.start();
addEventListener('keydown', (e) => { if (e.altKey && e.key === 'F5') { e.preventDefault(); presenter(cur); } });
}
window.__impresiona = { get pres() { return pres; }, set pres(p) { pres = p; }, get cur() { return cur; }, F: { insertChart }, F3, F4, presenter, renderAll, present, onCommand, buildPrint, select, startEditing, stopEditing, goto };
