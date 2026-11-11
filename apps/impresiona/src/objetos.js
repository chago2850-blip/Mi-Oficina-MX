// Impresiona — herramientas de objetos como en LibreOffice Impress: Posición y tamaño (F4), Línea
// (estilo, grosor, puntas de flecha), Área (color, degradado, transparencia, sombra), Texto (márgenes,
// anclaje, ajuste, interlineado y espaciado), más formas, voltear y girar, distribuir, alinear a la
// selección, traer adelante/enviar atrás un nivel, agrupar, recortar imagen y filtros, hipervínculos
// y acciones al hacer clic, ocultar diapositiva, cuadrícula con ajuste y clonar formato. GPL-3.0
import { h, toast, form, t } from '../../shared/office.js';
import { tabbed, field, numIn, sel, chk } from '../../shared/dialogs.js';

let C = null; // { getPres, slide, commit, select, selected, selObjs, objById, renderAll, cur, uid }
export function init(ctx) { C = ctx; }
const PX_CM = 2.54 / 96;
const r2 = (n) => Math.round(n * 100) / 100;
const cm = (px) => r2(px * PX_CM); const px = (c) => Math.round(Number(c) / PX_CM);
const need = () => { const l = C.selObjs(); if (!l.length) toast('Selecciona primero un objeto.'); return l; };

// ================================================================ formas
export const SHAPES = [
  ['rect', 'Rectángulo'], ['roundRect', 'Rectángulo redondeado'], ['ellipse', 'Elipse'], ['triangle', 'Triángulo'], ['rtTriangle', 'Triángulo rectángulo'],
  ['diamond', 'Rombo'], ['parallelogram', 'Paralelogramo'], ['trapezoid', 'Trapecio'], ['pentagon', 'Pentágono'], ['hexagon', 'Hexágono'], ['octagon', 'Octágono'],
  ['plus', 'Cruz'], ['star', 'Estrella'], ['star4', 'Estrella de 4 puntas'], ['heart', 'Corazón'], ['callout', 'Globo de diálogo'], ['cloud', 'Nube'],
  ['arrow', 'Flecha derecha'], ['leftArrow', 'Flecha izquierda'], ['upArrow', 'Flecha arriba'], ['downArrow', 'Flecha abajo'], ['leftRightArrow', 'Flecha doble'], ['chevron', 'Cheurón'], ['homePlate', 'Pentágono (flecha)'],
  ['line', 'Línea'], ['lineArrow', 'Flecha (línea)'],
];
const poly = (pts) => `M${pts.map((p) => p.join(',')).join(' L')} Z`;
function shapePath(kind, w, hh, i) {
  const W = w - i, H = hh - i, cx = w / 2, cy = hh / 2;
  switch (kind) {
    case 'triangle': return poly([[cx, i], [W, H], [i, H]]);
    case 'rtTriangle': return poly([[i, i], [W, H], [i, H]]);
    case 'diamond': return poly([[cx, i], [W, cy], [cx, H], [i, cy]]);
    case 'parallelogram': return poly([[w * 0.25, i], [W, i], [w * 0.75, H], [i, H]]);
    case 'trapezoid': return poly([[w * 0.2, i], [w * 0.8, i], [W, H], [i, H]]);
    case 'pentagon': { const p = []; for (let k = 0; k < 5; k++) { const a = -Math.PI / 2 + (k * 2 * Math.PI) / 5; p.push([cx + Math.cos(a) * (cx - i), cy + Math.sin(a) * (cy - i) * 1.05 + hh * 0.025]); } return poly(p); }
    case 'hexagon': return poly([[w * 0.25, i], [w * 0.75, i], [W, cy], [w * 0.75, H], [w * 0.25, H], [i, cy]]);
    case 'octagon': { const a = Math.min(w, hh) * 0.29; return poly([[a, i], [w - a, i], [W, a], [W, hh - a], [w - a, H], [a, H], [i, hh - a], [i, a]]); }
    case 'plus': { const a = w * 0.33, b = hh * 0.33; return poly([[a, i], [w - a, i], [w - a, b], [W, b], [W, hh - b], [w - a, hh - b], [w - a, H], [a, H], [a, hh - b], [i, hh - b], [i, b], [a, b]]); }
    case 'star': { const p = []; for (let k = 0; k < 10; k++) { const a = -Math.PI / 2 + (k * Math.PI) / 5; const r = k % 2 ? 0.4 : 0.5; p.push([cx + Math.cos(a) * r * w, cy + Math.sin(a) * r * hh]); } return poly(p); }
    case 'star4': { const p = []; for (let k = 0; k < 8; k++) { const a = -Math.PI / 2 + (k * Math.PI) / 4; const r = k % 2 ? 0.18 : 0.5; p.push([cx + Math.cos(a) * r * w, cy + Math.sin(a) * r * hh]); } return poly(p); }
    case 'heart': return `M${cx},${hh * 0.3} C${cx},${hh * 0.05} ${i},${hh * 0.05} ${i},${hh * 0.32} C${i},${hh * 0.6} ${cx},${hh * 0.8} ${cx},${H} C${cx},${hh * 0.8} ${W},${hh * 0.6} ${W},${hh * 0.32} C${W},${hh * 0.05} ${cx},${hh * 0.05} ${cx},${hh * 0.3} Z`;
    case 'callout': return poly([[i, i], [W, i], [W, hh * 0.72], [w * 0.42, hh * 0.72], [w * 0.2, H], [w * 0.25, hh * 0.72], [i, hh * 0.72]]);
    case 'cloud': return `M${w * 0.25},${hh * 0.85} C${w * 0.05},${hh * 0.85} ${w * 0.02},${hh * 0.55} ${w * 0.2},${hh * 0.5} C${w * 0.15},${hh * 0.2} ${w * 0.45},${hh * 0.1} ${w * 0.52},${hh * 0.3} C${w * 0.6},${hh * 0.08} ${w * 0.9},${hh * 0.15} ${w * 0.82},${hh * 0.42} C${w * 1.0},${hh * 0.48} ${w * 0.98},${hh * 0.85} ${w * 0.75},${hh * 0.85} Z`;
    case 'arrow': return poly([[i, hh * 0.3], [w * 0.62, hh * 0.3], [w * 0.62, i], [W, cy], [w * 0.62, H], [w * 0.62, hh * 0.7], [i, hh * 0.7]]);
    case 'leftArrow': return poly([[W, hh * 0.3], [w * 0.38, hh * 0.3], [w * 0.38, i], [i, cy], [w * 0.38, H], [w * 0.38, hh * 0.7], [W, hh * 0.7]]);
    case 'upArrow': return poly([[w * 0.3, H], [w * 0.3, hh * 0.38], [i, hh * 0.38], [cx, i], [W, hh * 0.38], [w * 0.7, hh * 0.38], [w * 0.7, H]]);
    case 'downArrow': return poly([[w * 0.3, i], [w * 0.3, hh * 0.62], [i, hh * 0.62], [cx, H], [W, hh * 0.62], [w * 0.7, hh * 0.62], [w * 0.7, i]]);
    case 'leftRightArrow': return poly([[i, cy], [w * 0.22, i], [w * 0.22, hh * 0.3], [w * 0.78, hh * 0.3], [w * 0.78, i], [W, cy], [w * 0.78, H], [w * 0.78, hh * 0.7], [w * 0.22, hh * 0.7], [w * 0.22, H]]);
    case 'chevron': return poly([[i, i], [w * 0.7, i], [W, cy], [w * 0.7, H], [i, H], [w * 0.3, cy]]);
    case 'homePlate': return poly([[i, i], [w * 0.75, i], [W, cy], [w * 0.75, H], [i, H]]);
    default: return null;
  }
}
const DASH = { solid: '', dash: '8 5', dot: '2 4', dashDot: '8 4 2 4', longDash: '16 6' };
// relleno (color o degradado), línea (estilo, grosor, puntas) y sombra de una forma, en SVG
export function shapeSvg(o, th) {
  const st = o.style || {}, w = o.w, hh = o.h;
  const id = `g${o.id || 'x'}`;
  let fill = st.fill === undefined ? th.accent : (st.fill || 'none');
  let defs = '';
  if (st.grad && st.fill !== null) {
    const g = st.grad; const c1 = st.fill || th.accent; const c2 = g.c2 || '#ffffff';
    if (g.type === 'radial') defs += `<radialGradient id="${id}"><stop offset="0" stop-color="${c1}"/><stop offset="1" stop-color="${c2}"/></radialGradient>`;
    else { const a = ((g.angle ?? 90) * Math.PI) / 180; const x = Math.cos(a) / 2, y = Math.sin(a) / 2; defs += `<linearGradient id="${id}" x1="${0.5 - x}" y1="${0.5 - y}" x2="${0.5 + x}" y2="${0.5 + y}"><stop offset="0" stop-color="${c1}"/><stop offset="1" stop-color="${c2}"/></linearGradient>`; }
    fill = `url(#${id})`;
  }
  const isLine = o.shape === 'line' || o.shape === 'lineArrow';
  const stroke = isLine ? (st.stroke || th.accent) : (st.stroke || 'none');
  const sw = isLine ? (st.strokeW || 4) : (st.strokeW || (st.stroke ? 2 : 0));
  const dash = DASH[st.dash] ? ` stroke-dasharray="${DASH[st.dash].split(' ').map((n) => n * Math.max(1, sw / 2)).join(' ')}"` : '';
  const sop = st.strokeOpacity != null ? ` stroke-opacity="${st.strokeOpacity}"` : '';
  const fop = st.fillOpacity != null ? ` fill-opacity="${st.fillOpacity}"` : '';
  const i = sw / 2;
  const head = (k, end) => {
    if (!k || k === 'none') return '';
    const mid = `${id}${end ? 'e' : 's'}`;
    const shape = k === 'circle' ? '<circle cx="5" cy="5" r="4"/>' : k === 'diamond' ? '<path d="M5 0 L10 5 L5 10 L0 5 Z"/>' : k === 'open' ? '<path d="M0 0 L10 5 L0 10" fill="none" stroke="context-stroke" stroke-width="1.5"/>' : '<path d="M0 0 L10 5 L0 10 Z"/>';
    defs += `<marker id="${mid}" viewBox="0 0 10 10" refX="${k === 'circle' || k === 'diamond' ? 5 : 8}" refY="5" markerWidth="${Math.max(3, 9 - sw / 2)}" markerHeight="${Math.max(3, 9 - sw / 2)}" orient="${end ? 'auto' : 'auto-start-reverse'}" fill="${stroke}">${shape}</marker>`;
    return ` marker-${end ? 'end' : 'start'}="url(#${mid})"`;
  };
  const wrap = (body) => `<svg width="100%" height="100%" viewBox="0 0 ${w} ${hh}" preserveAspectRatio="none">${defs ? `<defs>${defs}</defs>` : ''}${body}</svg>`;
  const common = `fill="${fill}" stroke="${stroke}" stroke-width="${sw}"${dash}${sop}${fop}`;
  if (isLine) {
    const ms = head(st.arrowStart, false); const me = head(st.arrowEnd || (o.shape === 'lineArrow' ? 'triangle' : null), true);
    return wrap(`<line x1="${sw}" y1="${hh / 2}" x2="${w - sw}" y2="${hh / 2}" stroke="${stroke}" stroke-width="${sw}"${dash}${sop}${ms}${me}/>`);
  }
  if (o.shape === 'ellipse') return wrap(`<ellipse cx="${w / 2}" cy="${hh / 2}" rx="${Math.max(0, w / 2 - i)}" ry="${Math.max(0, hh / 2 - i)}" ${common}/>`);
  if (o.shape === 'roundRect') return wrap(`<rect x="${i}" y="${i}" width="${Math.max(0, w - sw)}" height="${Math.max(0, hh - sw)}" rx="${Math.min(w, hh) * 0.15}" ${common}/>`);
  const d = shapePath(o.shape, w, hh, i);
  if (d) return wrap(`<path d="${d}" ${common}/>`);
  return wrap(`<rect x="${i}" y="${i}" width="${Math.max(0, w - sw)}" height="${Math.max(0, hh - sw)}" ${common}/>`);
}

// efectos comunes a cualquier objeto: voltear, transparencia, sombra, filtros, recorte, vínculo, texto
export function decorate(o, box, mode) {
  const st = o.style || {};
  const tf = []; if (o.rot) tf.push(`rotate(${o.rot}deg)`); if (o.flipH) tf.push('scaleX(-1)'); if (o.flipV) tf.push('scaleY(-1)');
  box.style.transform = tf.join(' ');
  if (o.flipH || o.flipV) box.classList.add('flipped');
  if (st.opacity != null && st.opacity < 1) box.style.opacity = String(st.opacity);
  const filters = [];
  if (st.shadow) { const s = st.shadow; filters.push(`drop-shadow(${s.dx ?? 4}px ${s.dy ?? 4}px ${s.blur ?? 6}px ${hexA(s.color || '#000000', s.alpha ?? 0.45)})`); }
  if (o.type === 'image' && o.img) { const f = o.img; if (f.gray) filters.push('grayscale(1)'); if (f.sepia) filters.push('sepia(1)'); if (f.bright != null && f.bright !== 100) filters.push(`brightness(${f.bright / 100})`); if (f.contrast != null && f.contrast !== 100) filters.push(`contrast(${f.contrast / 100})`); if (f.invert) filters.push('invert(1)'); }
  if (filters.length) box.style.filter = filters.join(' ');
  if (o.type === 'image' && o.crop) {
    const { l = 0, t: tp = 0, r = 0, b = 0 } = o.crop; const img = box.querySelector('img');
    if (img && (l || tp || r || b)) { const fw = 100 - l - r, fh = 100 - tp - b; box.style.overflow = 'hidden'; Object.assign(img.style, { position: 'absolute', width: `${(100 / fw) * 100}%`, height: `${(100 / fh) * 100}%`, left: `${(-l / fw) * 100}%`, top: `${(-tp / fh) * 100}%`, maxWidth: 'none' }); }
  }
  if (o.link) { box.dataset.link = o.link; if (mode === 'show') box.classList.add('has-link'); }
  const txt = box.querySelector('.otxt');
  if (txt && (o.flipH || o.flipV)) txt.style.transform = `scale(${o.flipH ? -1 : 1}, ${o.flipV ? -1 : 1})`; // el texto no se ve al revés
  if (txt) {
    if (st.pad) txt.style.padding = `${st.pad.t ?? 6}px ${st.pad.r ?? 10}px ${st.pad.b ?? 6}px ${st.pad.l ?? 10}px`;
    if (st.lineH) txt.style.lineHeight = String(st.lineH);
    if (st.paraSp != null) txt.style.setProperty('--para-sp', `${st.paraSp}pt`);
    if (st.autofit === 'shrink' && mode !== 'thumb') requestAnimationFrame(() => shrinkToFit(txt));
    else if (st.autofit === 'shrink') shrinkToFit(txt);
    if (st.cols > 1) { const tin = txt.querySelector('.tin'); if (tin) { tin.style.columnCount = String(st.cols); tin.style.columnGap = '18px'; } }
  }
}
function shrinkToFit(txt) {
  const tin = txt.querySelector('.tin'); if (!tin || !txt.isConnected) return;
  let k = 1; tin.style.fontSize = '';
  for (let n = 0; n < 12 && tin.scrollHeight > txt.clientHeight + 1 && k > 0.3; n++) { k -= 0.07; tin.style.fontSize = `${k * 100}%`; }
}
const hexA = (hex, a) => { const n = parseInt(hex.slice(1), 16); return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`; };

// ================================================================ Posición y tamaño (F4)
export async function posSizeDialog() {
  const l = need(); if (!l.length) return;
  const o = l[0];
  const body = await tabbed('Posición y tamaño', [
    { label: 'Posición y tamaño', body: h('div', { class: 'fx-grid' },
      h('h4', {}, t('Posición')), field('Posición X', numIn('x', cm(o.x), { step: 0.1, min: -100, max: 200, unit: 'cm' })), field('Posición Y', numIn('y', cm(o.y), { step: 0.1, min: -100, max: 200, unit: 'cm' })),
      h('h4', {}, t('Tamaño')), field('Ancho', numIn('w', cm(o.w), { step: 0.1, min: 0.05, max: 200, unit: 'cm' })), field('Alto', numIn('hh', cm(o.h), { step: 0.1, min: 0.05, max: 200, unit: 'cm' })),
      chk('keep', 'Mantener proporción', o.type === 'image'), chk('lock', 'Proteger posición y tamaño', !!o.locked)) },
    { label: 'Giro', body: h('div', { class: 'fx-grid' }, field('Ángulo de giro', numIn('rot', o.rot || 0, { step: 1, min: -360, max: 360, unit: '°' })), h('span'), chk('flipH', 'Voltear horizontalmente', !!o.flipH), chk('flipV', 'Voltear verticalmente', !!o.flipV)) },
  ], { width: '520px' });
  if (!body) return;
  const v = (n) => body.querySelector(`[name="${n}"]`);
  const ratio = o.h / (o.w || 1);
  C.commit('Posición y tamaño', () => {
    const nw = px(v('w').value); let nh = px(v('hh').value);
    if (v('keep').checked && Math.abs(nw - o.w) > 1 && Math.abs(nh - o.h) <= 1) nh = Math.round(nw * ratio);
    const dx = px(v('x').value) - o.x, dy = px(v('y').value) - o.y;
    for (const x of l) { x.x += dx; x.y += dy; x.rot = ((Number(v('rot').value) % 360) + 360) % 360 || 0; x.flipH = v('flipH').checked || undefined; x.flipV = v('flipV').checked || undefined; x.locked = v('lock').checked || undefined; }
    if (l.length === 1) { o.w = Math.max(2, nw); o.h = Math.max(2, nh); }
  });
}
export function flip(axis) { const l = need(); if (!l.length) return; C.commit('Voltear', () => l.forEach((o) => { if (axis === 'h') o.flipH = !o.flipH || undefined; else o.flipV = !o.flipV || undefined; })); }
export function rotate(deg) { const l = need(); if (!l.length) return; C.commit('Girar', () => l.forEach((o) => { o.rot = (((o.rot || 0) + deg) % 360 + 360) % 360; })); }

// ================================================================ Línea
const DASHES = [['solid', 'Continua'], ['dash', 'Guiones'], ['dot', 'Puntos'], ['dashDot', 'Guion y punto'], ['longDash', 'Guiones largos']];
const HEADS = [['none', 'Ninguna'], ['triangle', 'Flecha'], ['open', 'Flecha abierta'], ['circle', 'Círculo'], ['diamond', 'Rombo']];
export async function lineDialog() {
  const l = need(); if (!l.length) return; const st = l[0].style || {};
  const isLine = l.some((o) => o.shape === 'line' || o.shape === 'lineArrow');
  const body = await tabbed('Línea', [
    { label: 'Línea', body: h('div', { class: 'fx-grid' },
      field('Estilo', sel('dash', st.stroke === null && !isLine ? 'none' : st.dash || 'solid', [['none', 'Sin línea'], ...DASHES])), field('Color', h('input', { type: 'color', name: 'stroke', value: st.stroke || '#1d2030' })),
      field('Grosor', numIn('sw', r2((st.strokeW || (isLine ? 4 : 2)) * 0.75), { step: 0.25, min: 0.25, max: 40, unit: 'pt' })), field('Transparencia', numIn('sop', Math.round((1 - (st.strokeOpacity ?? 1)) * 100), { step: 5, min: 0, max: 100, unit: '%' })),
      h('h4', {}, t('Puntas de flecha (líneas)')), field('Inicio', sel('a1', st.arrowStart || 'none', HEADS)), field('Fin', sel('a2', st.arrowEnd || (l[0].shape === 'lineArrow' ? 'triangle' : 'none'), HEADS))) },
  ], { width: '520px' });
  if (!body) return;
  const v = (n) => body.querySelector(`[name="${n}"]`);
  C.commit('Línea', () => l.forEach((o) => {
    const s = { ...(o.style || {}) };
    if (v('dash').value === 'none') { s.stroke = null; delete s.dash; } else { s.stroke = v('stroke').value; s.dash = v('dash').value === 'solid' ? undefined : v('dash').value; s.strokeW = Math.max(0.3, Number(v('sw').value) / 0.75); }
    const op = 1 - Number(v('sop').value) / 100; s.strokeOpacity = op < 1 ? r2(op) : undefined;
    if (o.shape === 'line' || o.shape === 'lineArrow') { s.arrowStart = v('a1').value === 'none' ? undefined : v('a1').value; s.arrowEnd = v('a2').value; }
    for (const k of Object.keys(s)) if (s[k] === undefined) delete s[k];
    o.style = s;
  }));
}

// ================================================================ Área (relleno, degradado, transparencia, sombra)
export async function areaDialog() {
  const l = need(); if (!l.length) return; const o = l[0]; const st = o.style || {};
  const kind = st.fill === null ? 'none' : st.grad ? 'grad' : 'color';
  const sh = st.shadow || {};
  const body = await tabbed('Área', [
    { label: 'Relleno', body: h('div', { class: 'fx-grid' },
      field('Tipo', sel('kind', kind, [['none', 'Ninguno'], ['color', 'Color'], ['grad', 'Degradado']])), field('Color', h('input', { type: 'color', name: 'c1', value: st.fill || '#5b48e8' })),
      h('h4', {}, t('Degradado')), field('Segundo color', h('input', { type: 'color', name: 'c2', value: st.grad?.c2 || '#ffffff' })), field('Tipo de degradado', sel('gtype', st.grad?.type || 'linear', [['linear', 'Lineal'], ['radial', 'Radial']])),
      field('Ángulo', numIn('angle', st.grad?.angle ?? 90, { step: 15, min: 0, max: 360, unit: '°' }))) },
    { label: 'Transparencia', body: h('div', { class: 'fx-grid' }, field('Transparencia del relleno', numIn('fop', Math.round((1 - (st.fillOpacity ?? 1)) * 100), { step: 5, min: 0, max: 100, unit: '%' })), field('Transparencia del objeto', numIn('op', Math.round((1 - (st.opacity ?? 1)) * 100), { step: 5, min: 0, max: 100, unit: '%' }))) },
    { label: 'Sombra', body: h('div', { class: 'fx-grid' }, chk('sh', 'Usar sombra', !!st.shadow), h('span'), field('Color', h('input', { type: 'color', name: 'shc', value: sh.color || '#000000' })), field('Distancia', numIn('shd', sh.dx ?? 4, { step: 1, min: 0, max: 60, unit: 'px' })), field('Desenfoque', numIn('shb', sh.blur ?? 6, { step: 1, min: 0, max: 60, unit: 'px' })), field('Transparencia', numIn('sha', Math.round((1 - (sh.alpha ?? 0.45)) * 100), { step: 5, min: 0, max: 100, unit: '%' }))) },
  ], { width: '540px' });
  if (!body) return;
  const v = (n) => body.querySelector(`[name="${n}"]`);
  C.commit('Área', () => l.forEach((x) => {
    const s = { ...(x.style || {}) };
    const k = v('kind').value;
    if (k === 'none') { s.fill = null; delete s.grad; } else { s.fill = v('c1').value; if (k === 'grad') s.grad = { c2: v('c2').value, type: v('gtype').value, angle: Number(v('angle').value) }; else delete s.grad; }
    const fo = 1 - Number(v('fop').value) / 100; if (fo < 1) s.fillOpacity = r2(fo); else delete s.fillOpacity;
    const op = 1 - Number(v('op').value) / 100; if (op < 1) s.opacity = r2(op); else delete s.opacity;
    if (v('sh').checked) { const d = Number(v('shd').value); s.shadow = { color: v('shc').value, dx: d, dy: d, blur: Number(v('shb').value), alpha: r2(1 - Number(v('sha').value) / 100) }; } else delete s.shadow;
    x.style = s;
  }));
}

// ================================================================ Texto (márgenes, anclaje, ajuste, columnas, interlineado)
export async function textDialog() {
  const l = need().filter((o) => o.type === 'text' || o.type === 'shape'); if (!l.length) return; const st = l[0].style || {}; const p = st.pad || {};
  const body = await tabbed('Texto', [
    { label: 'Texto', body: h('div', { class: 'fx-grid' },
      h('h4', {}, t('Espacio a los bordes')), field('Izquierda', numIn('pl', cm(p.l ?? 10), { step: 0.05, min: 0, max: 10, unit: 'cm' })), field('Derecha', numIn('pr', cm(p.r ?? 10), { step: 0.05, min: 0, max: 10, unit: 'cm' })),
      field('Arriba', numIn('pt', cm(p.t ?? 6), { step: 0.05, min: 0, max: 10, unit: 'cm' })), field('Abajo', numIn('pb', cm(p.b ?? 6), { step: 0.05, min: 0, max: 10, unit: 'cm' })),
      h('h4', {}, t('Anclaje y ajuste')), field('Anclaje del texto', sel('valign', st.valign || 'top', [['top', 'Arriba'], ['middle', 'Centro'], ['bottom', 'Abajo']])),
      field('Ajuste automático', sel('autofit', st.autofit || 'none', [['none', 'No ajustar'], ['shrink', 'Reducir el texto al desbordar']])), field('Columnas', numIn('cols', st.cols || 1, { step: 1, min: 1, max: 4 }))) },
    { label: 'Sangrías y espacios', body: h('div', { class: 'fx-grid' },
      field('Interlineado (múltiplo)', numIn('lineH', st.lineH || 1.2, { step: 0.05, min: 0.5, max: 4 })), field('Espacio entre párrafos', numIn('paraSp', st.paraSp ?? 3, { step: 1, min: 0, max: 72, unit: 'pt' }))) },
  ], { width: '540px' });
  if (!body) return;
  const v = (n) => Number(body.querySelector(`[name="${n}"]`).value); const vs = (n) => body.querySelector(`[name="${n}"]`).value;
  C.commit('Texto', () => l.forEach((o) => {
    const s = { ...(o.style || {}) };
    s.pad = { l: px(v('pl')), r: px(v('pr')), t: px(v('pt')), b: px(v('pb')) };
    s.valign = vs('valign'); s.autofit = vs('autofit') === 'none' ? undefined : vs('autofit'); s.cols = v('cols') > 1 ? v('cols') : undefined;
    s.lineH = v('lineH') && Math.abs(v('lineH') - 1.2) > 0.001 ? v('lineH') : undefined; s.paraSp = v('paraSp') !== 3 ? v('paraSp') : undefined;
    for (const k of Object.keys(s)) if (s[k] === undefined) delete s[k];
    o.style = s;
  }));
}

// ================================================================ organizar: distribuir, alinear a la selección, un nivel, agrupar
export function distribute(axis) {
  const l = need(); if (l.length < 3) { toast('Selecciona al menos tres objetos.'); return; }
  C.commit('Distribuir', () => {
    const k = axis === 'h' ? 'x' : 'y', d = axis === 'h' ? 'w' : 'h';
    const s = [...l].sort((a, b) => a[k] - b[k]);
    const first = s[0], last = s[s.length - 1];
    const total = s.reduce((a, o) => a + o[d], 0); const span = last[k] + last[d] - first[k];
    const gap = (span - total) / (s.length - 1);
    let pos = first[k];
    for (const o of s) { o[k] = Math.round(pos); pos += o[d] + gap; }
  });
}
export function alignSel(kind) {
  const l = need(); if (!l.length) return;
  if (l.length === 1) return C.arrange(kind);
  const x1 = Math.min(...l.map((o) => o.x)), y1 = Math.min(...l.map((o) => o.y)), x2 = Math.max(...l.map((o) => o.x + o.w)), y2 = Math.max(...l.map((o) => o.y + o.h));
  C.commit('Alinear', () => l.forEach((o) => {
    if (kind === 'left') o.x = x1; if (kind === 'hcenter') o.x = Math.round((x1 + x2 - o.w) / 2); if (kind === 'right') o.x = x2 - o.w;
    if (kind === 'top') o.y = y1; if (kind === 'vcenter') o.y = Math.round((y1 + y2 - o.h) / 2); if (kind === 'bottom') o.y = y2 - o.h;
  }));
}
export function zStep(dir) {
  const l = need(); if (!l.length) return;
  C.commit(dir > 0 ? 'Traer adelante' : 'Enviar atrás', () => {
    const arr = C.slide().objects; const ids = new Set(l.map((o) => o.id));
    const idx = arr.map((o, i) => (ids.has(o.id) ? i : -1)).filter((i) => i >= 0);
    if (dir > 0) for (const i of idx.reverse()) { if (i < arr.length - 1 && !ids.has(arr[i + 1].id)) [arr[i], arr[i + 1]] = [arr[i + 1], arr[i]]; }
    else for (const i of idx) { if (i > 0 && !ids.has(arr[i - 1].id)) [arr[i], arr[i - 1]] = [arr[i - 1], arr[i]]; }
  });
}
export function group() {
  const l = need(); if (l.length < 2) { toast('Selecciona al menos dos objetos.'); return; }
  const g = C.uid(); C.commit('Agrupar', () => l.forEach((o) => { o.group = g; }));
}
export function ungroup() { const l = need(); if (!l.length) return; C.commit('Desagrupar', () => l.forEach((o) => { delete o.group; })); }
// al seleccionar un objeto agrupado se selecciona todo el grupo
export function expandGroups(ids) {
  const objs = C.slide().objects; const gs = new Set(ids.map((id) => objs.find((o) => o.id === id)?.group).filter(Boolean));
  if (!gs.size) return ids;
  return [...new Set([...ids, ...objs.filter((o) => gs.has(o.group)).map((o) => o.id)])];
}

// ================================================================ imagen: recortar, filtros, tamaño original
export async function cropDialog() {
  const o = need().find((x) => x.type === 'image'); if (!o) { toast('Selecciona una imagen.'); return; }
  const c = o.crop || {};
  const v = await form('Recortar imagen', [
    { name: 'l', label: 'Izquierda (%)', type: 'number', value: c.l || 0, min: 0, max: 90 }, { name: 'r', label: 'Derecha (%)', type: 'number', value: c.r || 0, min: 0, max: 90 },
    { name: 't', label: 'Arriba (%)', type: 'number', value: c.t || 0, min: 0, max: 90 }, { name: 'b', label: 'Abajo (%)', type: 'number', value: c.b || 0, min: 0, max: 90 },
    { name: 'keep', label: 'Conservar el tamaño de la imagen (cambiar el marco)', type: 'checkbox', value: true },
  ], 'Recortar');
  if (!v) return;
  const n = { l: Math.max(0, Math.min(90, Number(v.l) || 0)), r: Math.max(0, Math.min(90, Number(v.r) || 0)), t: Math.max(0, Math.min(90, Number(v.t) || 0)), b: Math.max(0, Math.min(90, Number(v.b) || 0)) };
  if (n.l + n.r >= 95 || n.t + n.b >= 95) { toast('El recorte es demasiado grande.', { error: true }); return; }
  C.commit('Recortar', () => {
    if (v.keep) { // el marco se ajusta para que la imagen no se deforme
      const ow = o.w / ((100 - (c.l || 0) - (c.r || 0)) / 100), oh = o.h / ((100 - (c.t || 0) - (c.b || 0)) / 100);
      o.x += Math.round(ow * (n.l - (c.l || 0)) / 100); o.y += Math.round(oh * (n.t - (c.t || 0)) / 100);
      o.w = Math.round(ow * (100 - n.l - n.r) / 100); o.h = Math.round(oh * (100 - n.t - n.b) / 100);
    }
    o.crop = n.l || n.r || n.t || n.b ? n : undefined;
  });
}
export const IMG_FILTERS = [['gray', 'Escala de grises'], ['sepia', 'Sepia'], ['invert', 'Invertir colores']];
export async function imageDialog() {
  const o = need().find((x) => x.type === 'image'); if (!o) { toast('Selecciona una imagen.'); return; }
  const f = o.img || {};
  const v = await form('Ajustes de imagen', [
    { name: 'bright', label: 'Brillo (%)', type: 'number', value: f.bright ?? 100, min: 0, max: 300 }, { name: 'contrast', label: 'Contraste (%)', type: 'number', value: f.contrast ?? 100, min: 0, max: 300 },
    { name: 'mode', label: 'Modo de color', type: 'select', value: f.gray ? 'gray' : f.sepia ? 'sepia' : f.invert ? 'invert' : '', options: [['', t('Predeterminado')], ...IMG_FILTERS.map(([k, l]) => [k, t(l)])] },
    { name: 'op', label: 'Transparencia (%)', type: 'number', value: Math.round((1 - ((o.style || {}).opacity ?? 1)) * 100), min: 0, max: 100 },
  ], 'Aceptar');
  if (!v) return;
  C.commit('Ajustes de imagen', () => {
    const n = { bright: Number(v.bright), contrast: Number(v.contrast) }; if (v.mode) n[v.mode] = true;
    o.img = (n.bright !== 100 || n.contrast !== 100 || v.mode) ? n : undefined;
    const op = 1 - Number(v.op) / 100; o.style = { ...(o.style || {}) }; if (op < 1) o.style.opacity = r2(op); else delete o.style.opacity;
  });
}
export function originalSize() {
  const o = need().find((x) => x.type === 'image'); if (!o) return;
  const img = new Image(); img.onload = () => C.commit('Tamaño original', () => { o.w = img.naturalWidth; o.h = img.naturalHeight; o.crop = undefined; }); img.src = o.src;
}

// ================================================================ hipervínculo / acción al hacer clic
export async function linkDialog() {
  const l = need(); if (!l.length) return; const o = l[0];
  const slides = C.getPres().slides;
  const cur = o.link || '';
  const kind = cur.startsWith('#slide:') ? 'slide' : cur.startsWith('#') ? cur.slice(1) : cur ? 'url' : 'none';
  const v = await form('Acción al hacer clic', [
    { name: 'kind', label: 'Acción', type: 'select', value: kind, options: [['none', t('Ninguna')], ['url', t('Ir a una página web o correo')], ['slide', t('Ir a la diapositiva')], ['next', t('Diapositiva siguiente')], ['prev', t('Diapositiva anterior')], ['first', t('Primera diapositiva')], ['last', t('Última diapositiva')], ['end', t('Terminar presentación')]] },
    { name: 'url', label: 'Dirección', value: kind === 'url' ? cur : 'https://' },
    { name: 'slide', label: 'Diapositiva', type: 'select', value: kind === 'slide' ? cur.slice(7) : slides[0]?.id, options: slides.map((s, i) => [s.id, `${i + 1}. ${slideTitle(s) || t('Diapositiva')}`]) },
  ], 'Aceptar');
  if (!v) return;
  const link = v.kind === 'none' ? undefined : v.kind === 'url' ? (v.url.trim() === 'https://' ? undefined : v.url.trim()) : v.kind === 'slide' ? `#slide:${v.slide}` : `#${v.kind}`;
  C.commit('Hipervínculo', () => l.forEach((x) => { x.link = link; }));
}
export const slideTitle = (s) => { const o = s.objects.find((x) => x.role === 'title') || s.objects.find((x) => x.type === 'text'); return o ? (new DOMParser().parseFromString(o.html || '', 'text/html').body.textContent || '').trim().slice(0, 40) : ''; };

// ================================================================ ocultar diapositiva
export function toggleHidden(i = C.cur()) { const s = C.getPres().slides[i]; if (!s) return; C.commit('Ocultar diapositiva', () => { s.hidden = !s.hidden || undefined; }); }

// ================================================================ cuadrícula y ajuste
export function toggleGrid() { const p = C.getPres(); p.grid = { ...(p.grid || { size: 20 }), on: !p.grid?.on }; C.renderAll(); return p.grid.on; }
export function toggleSnapGrid() { const p = C.getPres(); p.grid = { ...(p.grid || { size: 20 }), snap: !p.grid?.snap }; return p.grid.snap; }
export async function gridDialog() {
  const p = C.getPres(); const g = p.grid || { size: 20 };
  const v = await form('Cuadrícula', [{ name: 'size', label: 'Separación (cm)', type: 'number', value: cm(g.size || 20), min: 0.1, max: 10, step: 0.1 }, { name: 'on', label: 'Mostrar cuadrícula', type: 'checkbox', value: !!g.on }, { name: 'snap', label: 'Ajustar a la cuadrícula', type: 'checkbox', value: !!g.snap }], 'Aceptar');
  if (!v) return; p.grid = { size: Math.max(4, px(v.size)), on: v.on, snap: v.snap }; C.renderAll();
}
export const snapToGrid = (v) => { const g = C.getPres().grid; return g?.snap ? Math.round(v / (g.size || 20)) * (g.size || 20) : v; };

// ================================================================ clonar formato
let painter = null;
export function startPainter() { const l = need(); if (!l.length) return false; const o = l[0]; painter = { style: JSON.parse(JSON.stringify(o.style || {})), img: o.img ? { ...o.img } : undefined }; toast('Haz clic en el objeto al que quieres dar el mismo formato.'); return true; }
export function painterActive() { return !!painter; }
export function applyPainter(id) {
  if (!painter) return false;
  const o = C.slide().objects.find((x) => x.id === id); const p = painter; painter = null; if (!o) return false;
  C.commit('Clonar formato', () => { const keep = { ...(o.style || {}) }; o.style = { ...p.style }; if (keep.raw) o.style.raw = keep.raw; if (o.type === 'image' && p.img) o.img = { ...p.img }; });
  return true;
}
