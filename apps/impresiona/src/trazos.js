// Impresiona — parte 4 (como LibreOffice Impress): curvas, polígonos, polilíneas y formas libres
// dibujadas con el ratón, edición de puntos, y conectores que siguen a las formas que unen (recto,
// en ángulo o curvo). Se guardan como geometría propia en .pptx (custGeom) y en .odp (draw:path y
// draw:connector, que LibreOffice reconoce como conectores de verdad). GPL-3.0
import { h, t, toast } from '../../shared/office.js';

let C = null; // { stage, stagePoint, slide, commit, select, selObjs, uid, renderAll, objById, theme }
export function init(ctx) { C = ctx; }

// ================================================================ geometría
// segmentos en píxeles de la caja: [{ op: 'M'|'L'|'C'|'Z', p, c1, c2 }]
export function segments(o, w = o.w, hh = o.h) {
  const P = (o.pts || []).map(([x, y]) => [x * w, y * hh]);
  if (!P.length) return [];
  const out = [{ op: 'M', p: P[0] }];
  if (o.bez && P.length === 4) { out.push({ op: 'C', c1: P[1], c2: P[2], p: P[3] }); return out; }
  if (!o.smooth || P.length < 3) { for (let i = 1; i < P.length; i++) out.push({ op: 'L', p: P[i] }); }
  else {
    const n = P.length; const closed = !!o.closed;
    const at = (i) => (closed ? P[(i + n) % n] : P[Math.max(0, Math.min(n - 1, i))]);
    const last = closed ? n : n - 1;
    for (let i = 0; i < last; i++) {
      const p0 = at(i - 1), p1 = at(i), p2 = at(i + 1), p3 = at(i + 2);
      out.push({ op: 'C', c1: [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6], c2: [p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6], p: p2 });
    }
  }
  if (o.closed) out.push({ op: 'Z' });
  return out;
}
const f = (n) => Math.round(n * 100) / 100;
export function svgPath(o, w = o.w, hh = o.h) {
  return segments(o, w, hh).map((s) => (s.op === 'Z' ? 'Z' : s.op === 'C' ? `C${f(s.c1[0])} ${f(s.c1[1])} ${f(s.c2[0])} ${f(s.c2[1])} ${f(s.p[0])} ${f(s.p[1])}` : `${s.op}${f(s.p[0])} ${f(s.p[1])}`)).join(' ');
}
// puntos absolutos (de la diapositiva) → caja y puntos relativos
export function fromAbsolute(abs) {
  const xs = abs.map((p) => p[0]); const ys = abs.map((p) => p[1]);
  const x = Math.min(...xs); const y = Math.min(...ys); const w = Math.max(1, Math.max(...xs) - x); const hh = Math.max(1, Math.max(...ys) - y);
  return { x: Math.round(x), y: Math.round(y), w: Math.round(w), h: Math.round(hh), pts: abs.map(([px, py]) => [Math.round(((px - x) / w) * 10000) / 10000, Math.round(((py - y) / hh) * 10000) / 10000]) };
}
export const absolute = (o) => (o.pts || []).map(([x, y]) => [o.x + x * o.w, o.y + y * o.h]);
// simplificación de Ramer–Douglas–Peucker (para la forma libre)
export function simplify(pts, tol = 1.5) {
  if (pts.length < 3) return pts;
  const d2 = (p, a, b) => { const dx = b[0] - a[0], dy = b[1] - a[1]; const l = dx * dx + dy * dy; let tt = l ? ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / l : 0; tt = Math.max(0, Math.min(1, tt)); const x = a[0] + tt * dx - p[0], y = a[1] + tt * dy - p[1]; return x * x + y * y; };
  const keep = new Array(pts.length).fill(false); keep[0] = keep[pts.length - 1] = true;
  const stack = [[0, pts.length - 1]];
  while (stack.length) {
    const [a, b] = stack.pop(); let m = -1; let md = tol * tol;
    for (let i = a + 1; i < b; i++) { const d = d2(pts[i], pts[a], pts[b]); if (d > md) { md = d; m = i; } }
    if (m > 0) { keep[m] = true; stack.push([a, m], [m, b]); }
  }
  return pts.filter((_, i) => keep[i]);
}
export function newPath(abs, { closed = false, smooth = false } = {}) {
  const b = fromAbsolute(abs);
  return { id: C.uid(), type: 'shape', shape: 'path', ...b, rot: 0, closed, smooth, html: '', style: closed ? { valign: 'middle', align: 'center' } : { stroke: undefined, strokeW: 3 } };
}

// ================================================================ dibujar con el ratón
let tool = null;
export const active = () => tool?.kind || null;
export function cancel() { tool?.layer.remove(); tool = null; }
export function startTool(kind) {
  cancel();
  const layer = h('div', { class: 'draw-layer', tabindex: -1 });
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); svg.setAttribute('class', 'draw-preview'); layer.append(svg);
  const pathEl = document.createElementNS('http://www.w3.org/2000/svg', 'path'); svg.append(pathEl);
  C.stage.append(layer);
  tool = { kind, layer, pts: [], hover: null, down: false };
  const redraw = () => {
    const pts = tool.hover && tool.kind !== 'freehand' ? [...tool.pts, tool.hover] : tool.pts;
    const { w: W, h: H } = C.size(); svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
    pathEl.setAttribute('d', pts.length ? `M${pts.map((p) => `${f(p[0])} ${f(p[1])}`).join(' L')}${tool.kind === 'polygon' && pts.length > 2 ? ' Z' : ''}` : '');
  };
  const finish = () => {
    const k = tool.kind; let pts = tool.pts;
    if (k === 'freehand') pts = simplify(pts, 2);
    cancel();
    if (pts.length < 2) return null;
    const o = newPath(pts, { closed: k === 'polygon' || (k === 'freehand' && Math.hypot(pts[0][0] - pts.at(-1)[0], pts[0][1] - pts.at(-1)[1]) < 12 && pts.length > 3), smooth: k === 'curve' || k === 'freehand' });
    C.commit(t('Dibujar'), () => C.slide().objects.push(o)); C.select([o.id]);
    return o;
  };
  tool.finish = finish;
  layer.addEventListener('pointerdown', (e) => {
    e.preventDefault(); e.stopPropagation(); layer.setPointerCapture?.(e.pointerId);
    const p = C.stagePoint(e); const pt = [p.x, p.y];
    if (kind === 'freehand') { tool.down = true; tool.pts = [pt]; }
    else if (kind === 'connector') { tool.from = hitObject(e); if (!tool.from) { toast('Empieza el conector sobre una forma.'); } tool.pts = [pt]; tool.down = true; }
    else { const l = tool.pts.at(-1); if (!l || Math.hypot(l[0] - pt[0], l[1] - pt[1]) > 2) tool.pts.push(pt); }
    redraw();
  });
  layer.addEventListener('pointermove', (e) => {
    const p = C.stagePoint(e); const pt = [p.x, p.y];
    if (kind === 'freehand' && tool?.down) { const l = tool.pts.at(-1); if (Math.hypot(l[0] - pt[0], l[1] - pt[1]) > 1.5) tool.pts.push(pt); }
    else if (kind === 'connector' && tool?.down) tool.hover = pt;
    else if (tool) tool.hover = pt;
    if (tool) redraw();
  });
  layer.addEventListener('pointerup', (e) => {
    if (!tool) return;
    if (kind === 'freehand' && tool.down) finish();
    if (kind === 'connector' && tool.down) {
      tool.down = false; const a = tool.from; const p = C.stagePoint(e);
      layer.style.pointerEvents = 'none'; const b = hitObject(e); layer.style.pointerEvents = '';
      const kindC = tool.connKind || 'elbow'; cancel();
      if (a && b && a.id !== b.id) return connect(a.id, b.id, { kind: kindC, from: [a.x, a.y], to: [p.x, p.y] });
    }
    return null;
  });
  layer.addEventListener('dblclick', (e) => { e.preventDefault(); e.stopPropagation(); if (tool && kind !== 'freehand' && kind !== 'connector') finish(); });
  layer.addEventListener('keydown', (e) => { if (e.key === 'Escape') { e.preventDefault(); cancel(); } if (e.key === 'Enter') { e.preventDefault(); finish(); } if (e.key === 'Backspace' && tool) { tool.pts.pop(); redraw(); } });
  layer.focus();
  return tool;
}
export function setConnectorKind(k) { if (tool) tool.connKind = k; }
function hitObject(e) {
  const els = document.elementsFromPoint(e.clientX, e.clientY);
  const el = els.find((x) => x.closest?.('.slide.m-edit .obj[data-id]'))?.closest('.obj[data-id]');
  if (!el) return null;
  const o = C.objById(el.dataset.id); if (!o || o.conn) return null;
  const p = C.stagePoint(e); return { id: o.id, x: p.x, y: p.y };
}

// ================================================================ conectores
export const SIDES = ['top', 'right', 'bottom', 'left'];
const DIR = { top: [0, -1], right: [1, 0], bottom: [0, 1], left: [-1, 0] };
export function gluePoint(o, side) {
  const cx = o.x + o.w / 2, cy = o.y + o.h / 2;
  const p = { top: [cx, o.y], right: [o.x + o.w, cy], bottom: [cx, o.y + o.h], left: [o.x, cy] }[side];
  if (!o.rot) return p;
  const a = (o.rot * Math.PI) / 180; const dx = p[0] - cx, dy = p[1] - cy;
  return [cx + dx * Math.cos(a) - dy * Math.sin(a), cy + dx * Math.sin(a) + dy * Math.cos(a)];
}
// lado más cercano a un punto (o el que mira hacia la otra forma)
export function nearestSide(o, pt) { let best = 'right'; let bd = Infinity; for (const s of SIDES) { const g = gluePoint(o, s); const d = Math.hypot(g[0] - pt[0], g[1] - pt[1]); if (d < bd) { bd = d; best = s; } } return best; }
export function connect(aId, bId, { kind = 'elbow', sideA, sideB, from, to } = {}) {
  const A = C.objById(aId); const B = C.objById(bId); if (!A || !B) return null;
  const ca = [B.x + B.w / 2, B.y + B.h / 2]; const cb = [A.x + A.w / 2, A.y + A.h / 2];
  const o = { id: C.uid(), type: 'shape', shape: 'path', x: 0, y: 0, w: 1, h: 1, rot: 0, pts: [], html: '', closed: false, conn: { a: { id: aId, side: sideA || nearestSide(A, from || ca) }, b: { id: bId, side: sideB || nearestSide(B, to || cb) }, kind }, style: { strokeW: 2, arrowEnd: 'triangle' } };
  route(o, C.slide());
  C.commit(t('Conector'), () => C.slide().objects.push(o)); C.select([o.id]);
  return o;
}
// recalcula el trazado de un conector a partir de las formas que une
export function route(o, s) {
  const A = s.objects.find((x) => x.id === o.conn.a.id); const B = s.objects.find((x) => x.id === o.conn.b.id);
  if (!A || !B) return false;
  const pa = gluePoint(A, o.conn.a.side); const pb = gluePoint(B, o.conn.b.side);
  const da = DIR[o.conn.a.side]; const db = DIR[o.conn.b.side]; const k = o.conn.kind; const gap = 20;
  let abs; o.bez = false; o.smooth = false;
  if (k === 'straight') abs = [pa, pb];
  else if (k === 'curved') { const L = Math.max(40, Math.hypot(pb[0] - pa[0], pb[1] - pa[1]) / 2.5); abs = [pa, [pa[0] + da[0] * L, pa[1] + da[1] * L], [pb[0] + db[0] * L, pb[1] + db[1] * L], pb]; o.bez = true; }
  else {
    const s1 = [pa[0] + da[0] * gap, pa[1] + da[1] * gap]; const e1 = [pb[0] + db[0] * gap, pb[1] + db[1] * gap];
    const hA = da[0] !== 0; const hB = db[0] !== 0;
    let mid;
    if (hA && hB) { const mx = (s1[0] + e1[0]) / 2; mid = [[mx, s1[1]], [mx, e1[1]]]; }
    else if (!hA && !hB) { const my = (s1[1] + e1[1]) / 2; mid = [[s1[0], my], [e1[0], my]]; }
    else if (hA) mid = [[e1[0], s1[1]]]; else mid = [[s1[0], e1[1]]];
    abs = [pa, s1, ...mid, e1, pb].filter((p, i, arr) => i === 0 || Math.hypot(p[0] - arr[i - 1][0], p[1] - arr[i - 1][1]) > 0.5);
  }
  const b = fromAbsolute(abs);
  // si es una línea recta horizontal o vertical, la caja tiene grosor 1: se ajusta para no dividir entre 0
  Object.assign(o, b);
  return true;
}
// antes de dibujar: los conectores siguen a sus formas
export function updateConnectors(s) { let n = 0; for (const o of s?.objects || []) if (o.conn && route(o, s)) n++; return n; }
export function setKind(kind) {
  const objs = C.selObjs().filter((o) => o.conn); if (!objs.length) { toast('Selecciona un conector.', { error: true }); return; }
  C.commit(t('Conector'), () => { for (const o of objs) { const x = C.slide().objects.find((q) => q.id === o.id); x.conn.kind = kind; route(x, C.slide()); } });
}

// ================================================================ editar puntos (F8)
let editLayer = null;
export function editPoints(on = !editLayer) {
  editLayer?.remove(); editLayer = null;
  if (!on) return false;
  const o = C.selObjs().find((x) => x.shape === 'path' && !x.conn);
  if (!o) { toast('Selecciona una curva, un polígono o una forma libre.', { error: true }); return false; }
  editLayer = h('div', { class: 'pts-layer' }); C.stage.append(editLayer);
  const draw = () => {
    const cur = C.objById(o.id); if (!cur) { editPoints(false); return; }
    editLayer.replaceChildren(...absolute(cur).map(([x, y], i) => {
      const hd = h('div', { class: 'pt-handle', 'data-i': i, title: t('Arrastra para mover; doble clic para quitar'), style: { left: `${x * C.scale()}px`, top: `${y * C.scale()}px` } });
      hd.addEventListener('pointerdown', (e) => {
        e.preventDefault(); e.stopPropagation(); hd.setPointerCapture?.(e.pointerId);
        const abs0 = absolute(C.objById(o.id));
        const move = (ev) => { const p = C.stagePoint(ev); hd.style.left = `${p.x * C.scale()}px`; hd.style.top = `${p.y * C.scale()}px`; hd.dataset.x = p.x; hd.dataset.y = p.y; };
        const up = () => { hd.removeEventListener('pointermove', move); hd.removeEventListener('pointerup', up); if (hd.dataset.x == null) return; const abs = abs0.map((q, k) => (k === i ? [Number(hd.dataset.x), Number(hd.dataset.y)] : q)); movePoints(o.id, abs); draw(); };
        hd.addEventListener('pointermove', move); hd.addEventListener('pointerup', up);
      });
      hd.addEventListener('dblclick', (e) => { e.stopPropagation(); const abs = absolute(C.objById(o.id)); if (abs.length <= 2) return; abs.splice(i, 1); movePoints(o.id, abs); draw(); });
      return hd;
    }));
  };
  draw();
  editLayer.redraw = draw;
  return true;
}
export const editing = () => !!editLayer;
export function movePoints(id, abs) {
  C.commit(t('Editar puntos'), () => { const x = C.slide().objects.find((q) => q.id === id); if (x.rot) { x.rot = 0; } Object.assign(x, fromAbsolute(abs)); });
}
export function insertPoint(id, index, pt) { const abs = absolute(C.objById(id)); abs.splice(index, 0, pt); movePoints(id, abs); editLayer?.redraw?.(); }
export function toggleClosed() { const o = C.selObjs().find((x) => x.shape === 'path' && !x.conn); if (!o) return; C.commit(t('Cerrar objeto'), () => { const x = C.slide().objects.find((q) => q.id === o.id); x.closed = !x.closed; if (x.closed) x.style = { ...(x.style || {}), arrowStart: undefined, arrowEnd: undefined }; }); }
export function toggleSmooth() { const o = C.selObjs().find((x) => x.shape === 'path' && !x.conn); if (!o) return; C.commit(t('Suavizar'), () => { const x = C.slide().objects.find((q) => q.id === o.id); x.smooth = !x.smooth; }); }

// ================================================================ .pptx (custGeom) y .odp (draw:path)
// puntos para pptxgenjs, en pulgadas relativas a la caja del objeto
export function pptxPoints(o, pxIn) {
  return segments(o).map((s) => {
    if (s.op === 'Z') return { close: true };
    if (s.op === 'M') return { x: s.p[0] * pxIn, y: s.p[1] * pxIn, moveTo: true };
    if (s.op === 'L') return { x: s.p[0] * pxIn, y: s.p[1] * pxIn };
    return { x: s.p[0] * pxIn, y: s.p[1] * pxIn, curve: { type: 'cubic', x1: s.c1[0] * pxIn, y1: s.c1[1] * pxIn, x2: s.c2[0] * pxIn, y2: s.c2[1] * pxIn } };
  });
}
// svg:d en unidades de 1/100 mm para ODF, con viewBox de la caja
export function odfPath(o) {
  const k = 2540 / 96; // px → 1/100 mm
  return { viewBox: `0 0 ${Math.round(o.w * k)} ${Math.round(o.h * k)}`, d: segments(o).map((s) => (s.op === 'Z' ? 'Z' : s.op === 'C' ? `C ${Math.round(s.c1[0] * k)} ${Math.round(s.c1[1] * k)} ${Math.round(s.c2[0] * k)} ${Math.round(s.c2[1] * k)} ${Math.round(s.p[0] * k)} ${Math.round(s.p[1] * k)}` : `${s.op} ${Math.round(s.p[0] * k)} ${Math.round(s.p[1] * k)}`)).join(' ') };
}
// svg:d de ODF/SVG (M, L, C, Q, H, V, Z, también relativos) → puntos absolutos en la caja (px)
export function parseSvgPath(d, vb, box) {
  const [vx, vy, vw, vh] = vb.split(/[\s,]+/).map(Number);
  const sx = box.w / (vw || 1), sy = box.h / (vh || 1);
  const toks = String(d).match(/[a-zA-Z]|-?[\d.]+(?:e-?\d+)?/g) || [];
  const pts = []; let i = 0; let cmd = ''; let cx = 0, cy = 0; let start = [0, 0]; let closed = false; let curves = false;
  const num = () => Number(toks[i++]);
  const push = (x, y) => pts.push([(x - vx) * sx, (y - vy) * sy]);
  while (i < toks.length) {
    if (/[a-zA-Z]/.test(toks[i])) cmd = toks[i++];
    const rel = cmd === cmd.toLowerCase(); const U = cmd.toUpperCase();
    if (U === 'Z') { closed = true; cx = start[0]; cy = start[1]; continue; }
    if (U === 'M' || U === 'L') { const x = num() + (rel ? cx : 0); const y = num() + (rel ? cy : 0); cx = x; cy = y; if (U === 'M') start = [x, y]; push(x, y); if (U === 'M') cmd = rel ? 'l' : 'L'; }
    else if (U === 'H') { cx = num() + (rel ? cx : 0); push(cx, cy); }
    else if (U === 'V') { cy = num() + (rel ? cy : 0); push(cx, cy); }
    else if (U === 'C') { num(); num(); num(); num(); const x = num() + (rel ? cx : 0); const y = num() + (rel ? cy : 0); cx = x; cy = y; push(x, y); curves = true; }
    else if (U === 'Q' || U === 'S') { num(); num(); const x = num() + (rel ? cx : 0); const y = num() + (rel ? cy : 0); cx = x; cy = y; push(x, y); curves = true; }
    else if (U === 'A') { for (let k = 0; k < 5; k++) num(); const x = num() + (rel ? cx : 0); const y = num() + (rel ? cy : 0); cx = x; cy = y; push(x, y); curves = true; }
    else i++;
    if (Number.isNaN(cx) || Number.isNaN(cy)) break;
  }
  if (closed && pts.length > 1 && Math.hypot(pts[0][0] - pts.at(-1)[0], pts[0][1] - pts.at(-1)[1]) < 0.5) pts.pop();
  return { pts, closed, curves };
}
