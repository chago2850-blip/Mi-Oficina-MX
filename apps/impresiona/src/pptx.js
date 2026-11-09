// Abrir y guardar .pptx para Impresiona (pptxtojson para leer, PptxGenJS para escribir). GPL-3.0
import { parse } from 'pptxtojson';
import PptxGenJS from 'pptxgenjs';
import { themeById } from './themes.js';

const PT_PX = 4 / 3;
const PX_IN = 1 / 96;

function sanitizeHtml(html) {
  const d = new DOMParser().parseFromString(`<body>${html || ''}</body>`, 'text/html');
  d.querySelectorAll('script,style,iframe,object,embed').forEach((n) => n.remove());
  d.querySelectorAll('*').forEach((el) => { for (const a of [...el.attributes]) if (/^on/i.test(a.name) || (a.name === 'href' && /^\s*javascript:/i.test(a.value))) el.removeAttribute(a.name); });
  return d.body.innerHTML.replace(/&nbsp;/g, (m, i, str) => (/\S/.test(str[i - 1] || '') && /\S/.test(str[i + 6] || '') ? ' ' : m));
}
const fillColor = (f) => (f && f.type === 'color' ? f.value : null);
const SHAPES = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', rightArrow: 'arrow', line: 'line', star5: 'star', star: 'star' };

export async function readPptx(bytes, uid) {
  const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const json = await parse(buf, { imageMode: 'base64', videoMode: 'none', audioMode: 'none' });
  const W = Math.round(json.size.width * PT_PX), H = Math.round(json.size.height * PT_PX);
  const warnings = new Set();
  const conv = (el, ox = 0, oy = 0) => {
    const base = { id: uid(), x: (el.left + ox) * PT_PX, y: (el.top + oy) * PT_PX, w: el.width * PT_PX, h: el.height * PT_PX, rot: el.rotate || 0 };
    if (el.type === 'text') return [{ ...base, type: 'text', html: sanitizeHtml(el.content), style: { valign: { mid: 'middle', ctr: 'middle', b: 'bottom', bottom: 'bottom' }[el.vAlign] || 'top', fill: fillColor(el.fill), stroke: el.borderWidth ? el.borderColor : null, strokeW: el.borderWidth || 0, raw: true } }];
    if (el.type === 'shape' && !fillColor(el.fill) && !el.borderWidth && el.content && (el.shapType === 'rect' || !el.shapType)) {
      // Cuadro de texto guardado como rectángulo sin relleno ni borde
      return [{ ...base, type: 'text', html: sanitizeHtml(el.content), style: { valign: { mid: 'middle', ctr: 'middle', b: 'bottom', bottom: 'bottom' }[el.vAlign] || 'top', raw: true } }];
    }
    if (el.type === 'shape') {
      const sh = SHAPES[el.shapType];
      if (!sh) warnings.add('algunas formas especiales se convirtieron en rectángulos');
      return [{ ...base, type: 'shape', shape: sh || 'rect', html: sanitizeHtml(el.content), style: { fill: fillColor(el.fill) || null, stroke: el.borderWidth ? el.borderColor : null, strokeW: el.borderWidth * PT_PX || 0, valign: { mid: 'middle', ctr: 'middle', b: 'bottom' }[el.vAlign] || 'middle', align: 'center', raw: true } }];
    }
    if (el.type === 'image' && el.base64) return [{ ...base, type: 'image', src: el.base64 }];
    if (el.type === 'table') return [{ ...base, type: 'table', rows: el.data.map((row) => row.map((c) => ({ text: c.text || '', bg: c.fillColor || null, color: c.fontColor || null, bold: !!c.fontBold, hidden: !!(c.vMerge || c.hMerge), rs: c.rowSpan || 1, cs: c.colSpan || 1 }))), style: {} }];
    if (el.type === 'group') return el.elements.flatMap((e) => conv(e, el.left + ox, el.top + oy));
    if (el.type === 'chart') warnings.add('los gráficos se muestran como un espacio vacío');
    else if (el.type === 'diagram') warnings.add('los diagramas SmartArt no se pudieron editar');
    else if (el.type === 'video' || el.type === 'audio') warnings.add('los videos y audios no se incluyen todavía');
    return [];
  };
  const slides = json.slides.map((s) => {
    const objects = [...(s.layoutElements || []).filter((e) => e.type === 'image' || e.type === 'shape'), ...s.elements].flatMap((e) => conv(e));
    const bg = fillColor(s.fill);
    const bgImage = s.fill?.type === 'image' ? s.fill.value?.picBase64 : null;
    return { id: uid(), layout: 'blank', bg: bg || null, bgImage: bgImage || null, notes: s.note || '', transition: s.transition ? { type: mapTransition(s.transition.type), dur: Math.max(0.3, (s.transition.duration || 700) / 1000) } : { type: 'none', dur: 0.6 }, objects };
  });
  return { pres: { size: { w: W, h: H }, theme: 'claro', slides: slides.length ? slides : [] }, warnings: [...warnings] };
}
const mapTransition = (t) => ({ fade: 'fade', push: 'push', wipe: 'wipe', cover: 'cover', zoom: 'zoom', split: 'zoom' }[t] || (t ? 'fade' : 'none'));

// HTML del cuadro de texto -> párrafos con fragmentos para PptxGenJS
function htmlToRuns(html, defaults) {
  const d = new DOMParser().parseFromString(`<body>${html || ''}</body>`, 'text/html');
  const out = [];
  const pushPara = (el, bullet, level = 0) => {
    const runs = [];
    const walk = (n, fmt) => {
      for (const c of n.childNodes) {
        if (c.nodeType === 3) { if (c.nodeValue) runs.push({ text: c.nodeValue, options: { ...fmt } }); continue; }
        if (c.nodeType !== 1) continue;
        if (c.tagName === 'BR') { runs.push({ text: '', options: { ...fmt, breakLine: true } }); continue; }
        if (c.tagName === 'UL' || c.tagName === 'OL') continue;
        const f = { ...fmt };
        if (['B', 'STRONG'].includes(c.tagName)) f.bold = true;
        if (['I', 'EM'].includes(c.tagName)) f.italic = true;
        if (c.tagName === 'U') f.underline = { style: 'sng' };
        if (['S', 'STRIKE'].includes(c.tagName)) f.strike = 'sngStrike';
        const st = c.style;
        if (st.fontWeight === 'bold' || Number(st.fontWeight) >= 600) f.bold = true;
        if (st.fontStyle === 'italic') f.italic = true;
        if (st.color) f.color = toHex(st.color);
        if (st.fontSize) f.fontSize = st.fontSize.endsWith('pt') ? parseFloat(st.fontSize) : parseFloat(st.fontSize) * 0.75;
        if (st.fontFamily) f.fontFace = st.fontFamily.split(',')[0].replace(/["']/g, '').trim();
        if (c.tagName === 'A' && c.getAttribute('href')) f.hyperlink = { url: c.getAttribute('href') };
        walk(c, f);
      }
    };
    walk(el, {});
    const align = el.style?.textAlign || defaults.align;
    if (!runs.length) runs.push({ text: '', options: {} });
    runs[0].options = { ...runs[0].options, ...(bullet ? { bullet: bullet === 'num' ? { type: 'number' } : true, indentLevel: level } : {}), align };
    runs[runs.length - 1].options.breakLine = true;
    out.push(...runs);
  };
  const walkBlocks = (n, level = 0) => {
    for (const c of n.children) {
      if (c.tagName === 'UL' || c.tagName === 'OL') { for (const li of c.children) { pushPara(li, c.tagName === 'OL' ? 'num' : 'bullet', level); for (const sub of li.children) if (sub.tagName === 'UL' || sub.tagName === 'OL') walkBlocks({ children: [sub] }, level + 1); } }
      else pushPara(c, null);
    }
  };
  if (!d.body.children.length && d.body.textContent) pushPara(d.body, null);
  else walkBlocks(d.body);
  if (out.length) delete out[out.length - 1].options.breakLine;
  return out;
}
function toHex(c) {
  if (!c) return undefined;
  if (c.startsWith('#')) return c.slice(1, 7).toUpperCase();
  const m = c.match(/rgba?\(([^)]+)\)/); if (!m) return undefined;
  return m[1].split(',').slice(0, 3).map((x) => Math.round(parseFloat(x)).toString(16).padStart(2, '0')).join('').toUpperCase();
}
const SHAPE_OUT = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', arrow: 'rightArrow', line: 'line', star: 'star5' };

export async function writePptx(pres, renderSlideImage) {
  const p = new PptxGenJS();
  const W = pres.size.w * PX_IN, H = pres.size.h * PX_IN;
  p.defineLayout({ name: 'MIOFICINA', width: W, height: H });
  p.layout = 'MIOFICINA';
  p.author = 'Impresiona · Mi Oficina MX';
  p.company = 'Mi Oficina MX';
  const th = themeById(pres.theme);
  for (const s of pres.slides) {
    const slide = p.addSlide();
    if (s.bgImage) slide.background = { data: s.bgImage };
    else slide.background = { color: (s.bg || th.bg).replace('#', '') };
    for (const o of s.objects) {
      const pos = { x: o.x * PX_IN, y: o.y * PX_IN, w: Math.max(0.05, o.w * PX_IN), h: Math.max(0.05, o.h * PX_IN), rotate: o.rot || 0 };
      const st = o.style || {};
      const defaults = { align: st.align || 'left' };
      if (o.type === 'text' || (o.type === 'shape' && o.html)) {
        const runs = htmlToRuns(o.html, defaults);
        if (!runs.some((r) => r.text) && o.type === 'text') continue; // marcador vacío
        const opts = {
          ...pos, valign: { middle: 'middle', bottom: 'bottom' }[st.valign] || 'top', margin: 4,
          fontFace: st.font || (o.role === 'title' ? th.titleFont : th.bodyFont), fontSize: st.size || 18, color: (st.color || (o.role === 'title' ? th.title : th.text)).replace('#', ''), bold: !!st.bold,
          fit: 'none',
        };
        if (o.type === 'shape') Object.assign(opts, { shape: SHAPE_OUT[o.shape] || 'rect', fill: st.fill ? { color: st.fill.replace('#', '') } : undefined, line: st.stroke ? { color: st.stroke.replace('#', ''), width: (st.strokeW || 1) * 0.75 } : undefined, rectRadius: o.shape === 'roundRect' ? 0.15 : undefined, color: (st.color || '#FFFFFF').replace('#', '') });
        else if (st.fill) opts.fill = { color: st.fill.replace('#', '') };
        slide.addText(runs, opts);
      } else if (o.type === 'shape') {
        slide.addShape(SHAPE_OUT[o.shape] || 'rect', { ...pos, fill: st.fill ? { color: st.fill.replace('#', '') } : { type: 'none' }, line: st.stroke ? { color: st.stroke.replace('#', ''), width: (st.strokeW || 1) * 0.75 } : { type: 'none' }, rectRadius: o.shape === 'roundRect' ? 0.15 : undefined });
      } else if (o.type === 'image' && o.src) {
        slide.addImage({ data: o.src, ...pos });
      } else if (o.type === 'table') {
        const rows = o.rows.map((r) => r.filter((c) => !c.hidden).map((c) => ({ text: c.text, options: { fill: c.bg ? { color: c.bg.replace('#', '') } : undefined, color: c.color ? c.color.replace('#', '') : undefined, bold: c.bold, colspan: c.cs > 1 ? c.cs : undefined, rowspan: c.rs > 1 ? c.rs : undefined } })));
        slide.addTable(rows, { ...pos, fontSize: st.size || 14, border: { type: 'solid', pt: 0.75, color: '9AA0B4' }, autoPage: false });
      }
    }
    if (s.notes) slide.addNotes(s.notes);
  }
  void renderSlideImage;
  const ab = await p.write({ outputType: 'arraybuffer' });
  return new Uint8Array(ab);
}
