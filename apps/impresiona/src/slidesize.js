// Impresiona — tamaño de diapositiva: tamaños predefinidos (como PowerPoint), ancho y alto en cm,
// orientación, numerar desde y cómo escalar el contenido al cambiar de tamaño. GPL-3.0
import { h, modal, toast, t, tf } from '../../shared/office.js';

// 96 px = 1 pulgada = 2.54 cm
export const CM_PX = 96 / 2.54;
export const cmToPx = (cm) => Math.round(cm * CM_PX);
export const pxToCm = (px) => Math.round((px / CM_PX) * 100) / 100;
const MIN_CM = 2.54, MAX_CM = 142.24;

// tamaños en cm (horizontal); los mismos que ofrece PowerPoint
export const PRESETS = [
  { id: 'screen169', label: 'Presentación en pantalla (16:9)', w: 25.4, h: 14.29 },
  { id: 'wide', label: 'Panorámica (33.87 × 19.05 cm)', w: 33.867, h: 19.05 },
  { id: 'screen43', label: 'Presentación en pantalla (4:3)', w: 25.4, h: 19.05 },
  { id: 'screen1610', label: 'Presentación en pantalla (16:10)', w: 25.4, h: 15.875 },
  { id: 'letter', label: 'Papel carta (8.5 × 11 pulg.)', w: 25.4, h: 19.05 },
  { id: 'ledger', label: 'Papel de contabilidad (11 × 17 pulg.)', w: 33.831, h: 25.374 },
  { id: 'a3', label: 'Hoja A3 (297 × 420 mm)', w: 35.56, h: 26.67 },
  { id: 'a4', label: 'Hoja A4 (210 × 297 mm)', w: 27.517, h: 19.05 },
  { id: 'b4', label: 'Hoja B4 (ICO) (250 × 353 mm)', w: 30.074, h: 22.556 },
  { id: 'b5', label: 'Hoja B5 (ICO) (176 × 250 mm)', w: 21.004, h: 15.749 },
  { id: 'slide35', label: 'Diapositivas de 35 mm', w: 28.575, h: 19.05 },
  { id: 'overhead', label: 'Transparencia', w: 25.4, h: 19.05 },
  { id: 'banner', label: 'Banner', w: 20.32, h: 2.54 },
  { id: 'custom', label: 'Personalizado', w: null, h: null },
];
const same = (a, b) => Math.abs(a - b) <= 1;
// tamaño predefinido que corresponde a la presentación (en cualquier orientación)
export function currentPreset(pres) {
  const { w, h: hh } = pres.size;
  const fits = (p) => p.w && ((same(cmToPx(p.w), w) && same(cmToPx(p.h), hh)) || (same(cmToPx(p.w), hh) && same(cmToPx(p.h), w)));
  const named = PRESETS.find((p) => p.id === pres.sizeName && fits(p));
  return named || PRESETS.find(fits) || PRESETS.find((p) => p.id === 'custom');
}
export const sizeLabel = (pres) => `${t(currentPreset(pres).label)} · ${pxToCm(pres.size.w)} × ${pxToCm(pres.size.h)} cm`;

// Cambia el tamaño y escala el contenido: 'fit' = asegurar ajuste (todo cabe), 'max' = maximizar (llena la diapositiva)
export function resizePres(pres, w, hh, mode, objStyle, theme) {
  const W0 = pres.size.w, H0 = pres.size.h;
  if (W0 === w && H0 === hh) return;
  const kx = w / W0, ky = hh / H0;
  const k = mode === 'max' ? Math.max(kx, ky) : Math.min(kx, ky);
  const ox = (w - W0 * k) / 2, oy = (hh - H0 * k) / 2;
  const scaleHtml = (html) => String(html || '').replace(/font-size:\s*([\d.]+)(px|pt)/g, (m, n, u) => `font-size: ${Math.max(1, Math.round(Number(n) * k * 10) / 10)}${u}`);
  const scaleObj = (o) => {
    o.x = Math.round(o.x * k + ox); o.y = Math.round(o.y * k + oy); o.w = Math.max(1, Math.round(o.w * k)); o.h = Math.max(1, Math.round(o.h * k));
    if (o.type === 'text' || o.type === 'shape' || o.type === 'table') {
      o.style ||= {};
      const cur = o.style.size || objStyle(o, theme).size;
      if (cur) o.style.size = Math.max(1, Math.round(cur * k * 2) / 2);
      if (o.html) o.html = scaleHtml(o.html);
      if (o.rows) o.rows.forEach((r) => r.forEach((c) => { if (c && typeof c === 'object' && c.html) c.html = scaleHtml(c.html); }));
    }
    if (o.style?.strokeW) o.style.strokeW = Math.max(1, Math.round(o.style.strokeW * k * 10) / 10);
    if (o.anim?.dist) o.anim.dist = Math.round(Number(o.anim.dist) * k);
  };
  pres.slides.forEach((s) => s.objects.forEach(scaleObj));
  (pres.master?.objects || []).forEach(scaleObj);
  pres.size = { w, h: hh };
}

// Cuadro "Configurar página": devuelve { w, h, sizeName, firstNum, mode } o null
export async function slideSizeDialog(pres) {
  const cur = currentPreset(pres);
  const sel = h('select', { name: 'preset', 'aria-label': t('Tamaño de la diapositiva') }, ...PRESETS.map((p) => h('option', { value: p.id, selected: p.id === cur.id ? true : null }, t(p.label))));
  const num = (v) => h('input', { type: 'number', step: '0.01', min: MIN_CM, max: MAX_CM, value: v });
  const wIn = num(pxToCm(pres.size.w)); const hIn = num(pxToCm(pres.size.h));
  wIn.name = 'w'; hIn.name = 'h';
  const ori = h('select', { name: 'orient' }, h('option', { value: 'land', selected: pres.size.w >= pres.size.h ? true : null }, t('Horizontal')), h('option', { value: 'port', selected: pres.size.w < pres.size.h ? true : null }, t('Vertical')));
  const first = h('input', { type: 'number', name: 'first', min: 0, max: 9999, step: 1, value: pres.firstNum ?? 1 });
  const scale = h('select', { name: 'scale' }, h('option', { value: 'fit' }, t('Asegurar ajuste (todo el contenido cabe)')), h('option', { value: 'max' }, t('Maximizar (el contenido llena la diapositiva)')));
  const prev = h('div', { class: 'ss-prev', 'aria-hidden': 'true' }, h('div', { class: 'ss-page' }));
  const drawPrev = () => {
    const w = Number(wIn.value) || 1, hh = Number(hIn.value) || 1; const k = 90 / Math.max(w, hh);
    Object.assign(prev.firstChild.style, { width: `${Math.max(4, w * k)}px`, height: `${Math.max(4, hh * k)}px` });
  };
  const applyPreset = () => {
    const p = PRESETS.find((x) => x.id === sel.value); if (!p?.w) return;
    const port = ori.value === 'port';
    wIn.value = port ? p.h : p.w; hIn.value = port ? p.w : p.h; drawPrev();
  };
  sel.addEventListener('change', applyPreset);
  ori.addEventListener('change', () => {
    const w = Number(wIn.value), hh = Number(hIn.value);
    if ((ori.value === 'port') === (w > hh)) { wIn.value = hh; hIn.value = w; }
    drawPrev();
  });
  const typed = () => {
    const w = Number(wIn.value), hh = Number(hIn.value);
    const match = (x) => x.w && ((Math.abs(x.w - w) < 0.02 && Math.abs(x.h - hh) < 0.02) || (Math.abs(x.h - w) < 0.02 && Math.abs(x.w - hh) < 0.02));
    if (!match(PRESETS.find((x) => x.id === sel.value) || {})) sel.value = PRESETS.find(match)?.id || 'custom';
    ori.value = w < hh ? 'port' : 'land'; drawPrev();
  };
  wIn.addEventListener('input', typed); hIn.addEventListener('input', typed);
  const field = (label, el) => h('label', { class: 'field' }, h('span', {}, t(label)), el);
  const body = h('div', { class: 'ss' },
    h('div', { class: 'ss-main' },
      field('Tamaño de la diapositiva', sel),
      h('div', { class: 'ss-row' }, field('Ancho (cm)', wIn), field('Alto (cm)', hIn)),
      field('Orientación de la diapositiva', ori),
      field('Numerar las diapositivas desde', first),
      field('Al cambiar el tamaño, el contenido', scale)),
    prev);
  drawPrev();
  const ok = await modal({ title: t('Tamaño de diapositiva'), body, width: '470px', buttons: [{ label: 'Cancelar', value: null }, { label: 'Aceptar', primary: true, value: true }] });
  if (!ok) return null;
  const w = Number(wIn.value), hh = Number(hIn.value);
  if (!(w >= MIN_CM && w <= MAX_CM && hh >= MIN_CM && hh <= MAX_CM)) { toast(tf('El ancho y el alto deben estar entre {0} y {1} cm.', MIN_CM, MAX_CM), { error: true }); return null; }
  return { w: cmToPx(w), h: cmToPx(hh), sizeName: sel.value, firstNum: Math.max(0, Math.round(Number(first.value) || 0)), mode: scale.value };
}
