// Abrir y guardar .pptx para Impresiona (pptxtojson para leer, PptxGenJS para escribir). GPL-3.0
import JSZip from 'jszip';
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
  const json = await parse(buf, { imageMode: 'base64', videoMode: 'blob', audioMode: 'blob' });
  const media = [];
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
    if (el.type === 'chart') {
      const ct = String(el.chartType || '');
      const type = /bar/i.test(ct) ? (el.barDir === 'bar' ? 'bar' : 'column') : /doughnut/i.test(ct) ? 'doughnut' : /pie/i.test(ct) ? 'pie' : /area/i.test(ct) ? 'area' : /scatter/i.test(ct) ? 'scatter' : 'line';
      const series = Array.isArray(el.data) ? el.data : [];
      if (type === 'scatter' && el.data?.length) {
        const xs = el.data[0] || [], ys = el.data[1] || [];
        return [{ ...base, type: 'chart', chartType: 'scatter', title: '', data: { cats: xs.map((_, i) => String(i + 1)), series: [{ name: 'X', values: xs }, { name: 'Y', values: ys }] } }];
      }
      const first = series[0] || { values: [], xlabels: {} };
      const cats = first.values.map((v, i) => String(first.xlabels?.[i] ?? first.xlabels?.[v.x] ?? v.x ?? i + 1));
      return [{ ...base, type: 'chart', chartType: type, title: '', data: { cats, series: series.map((sr) => ({ name: String(sr.key ?? 'Serie'), values: sr.values.map((v) => (typeof v.y === 'number' ? v.y : Number(v.y) || null)) })) } }];
    }
    else if (el.type === 'diagram') warnings.add('los diagramas SmartArt no se pudieron editar');
    else if ((el.type === 'video' || el.type === 'audio') && el.blob) { const o = { ...base, type: 'media', kind: el.type, src: el.blob, autoplay: false, loop: false, captions: [], style: {} }; media.push(o); return [o]; }
    else if (el.type === 'video' || el.type === 'audio') warnings.add('algunos videos o audios vinculados (no incrustados) no se incluyeron');
    return [];
  };
  const slides = json.slides.map((s) => {
    const objects = [...(s.layoutElements || []).filter((e) => e.type === 'image' || e.type === 'shape'), ...s.elements].flatMap((e) => conv(e));
    const bg = fillColor(s.fill);
    const bgImage = s.fill?.type === 'image' ? s.fill.value?.picBase64 : null;
    return { id: uid(), layout: 'blank', bg: bg || null, bgImage: bgImage || null, notes: noteText(s.note), transition: s.transition ? { type: mapTransition(s.transition.type), dur: Math.max(0.3, (s.transition.duration || 700) / 1000) } : { type: 'none', dur: 0.6 }, objects };
  });
  // los videos y audios incrustados llegan como blob: se guardan como datos dentro de la presentación
  for (const o of media) {
    try { const b = await (await fetch(o.src)).blob(); o.src = await new Promise((res) => { const rd = new FileReader(); rd.onload = () => res(rd.result); rd.readAsDataURL(b); }); } catch { o.src = ''; }
  }
  for (const s of slides) s.objects = s.objects.filter((o) => o.type !== 'media' || o.src);
  const imported = { size: { w: W, h: H }, theme: 'claro', slides: slides.length ? slides : [] };
  const merged = await applySuiteMeta(bytes, imported);
  return { pres: merged || imported, warnings: merged ? [] : [...warnings] };
}
function noteText(n) {
  if (!n) return '';
  if (!/[<>]/.test(n)) return n;
  const d = new DOMParser().parseFromString(`<body>${n}</body>`, 'text/html');
  d.querySelectorAll('p,br,li').forEach((el) => el.append('\n'));
  return d.body.textContent.replace(/\n{3,}/g, '\n\n').trim();
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

export async function writePptx(pres, renderSlideImage, props = {}) {
  const p = new PptxGenJS();
  const W = pres.size.w * PX_IN, H = pres.size.h * PX_IN;
  p.defineLayout({ name: 'MIOFICINA', width: W, height: H });
  p.layout = 'MIOFICINA';
  p.author = props.author || 'Impresiona · Mi Oficina MX';
  if (props.title) p.title = props.title;
  if (props.subject) p.subject = props.subject;
  p.company = 'Mi Oficina MX';
  const th = themeById(pres.theme);
  for (const s of pres.slides) {
    const slide = p.addSlide();
    if (s.bgImage) slide.background = { data: s.bgImage };
    else slide.background = { color: (s.bg || th.bg).replace('#', '') };
    const masterObjs = pres.master?.objects || [];
    for (const o of [...masterObjs, ...s.objects]) {
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
      } else if (o.type === 'chart') {
        const d = o.data || { cats: [], series: [] };
        const CT = { column: 'bar', bar: 'bar', line: 'line', area: 'area', pie: 'pie', doughnut: 'doughnut', scatter: 'scatter' };
        const data = o.chartType === 'scatter'
          ? [{ name: 'X', values: d.series[0]?.values || [] }, ...d.series.slice(1).map((sr) => ({ name: sr.name, values: sr.values }))]
          : d.series.map((sr) => ({ name: sr.name, labels: d.cats, values: sr.values.map((v) => v ?? 0) }));
        if (!data.length) continue;
        slide.addChart(CT[o.chartType] || 'bar', data, { ...pos, barDir: o.chartType === 'bar' ? 'bar' : 'col', title: o.title || undefined, showTitle: !!o.title, showLegend: d.series.length > 1 || o.chartType === 'pie' || o.chartType === 'doughnut', legendPos: 'b', chartColors: [th.accent, th.accent2, 'F08A24', '1F8FD6', '2FB4A8', '8E6CD9'].map((c) => c.replace('#', '')), catAxisLabelColor: th.text.replace('#', ''), valAxisLabelColor: th.text.replace('#', ''), titleColor: th.title.replace('#', ''), dataLabelColor: th.text.replace('#', '') });
      } else if (o.type === 'image' && o.src) {
        slide.addImage({ data: o.src, ...pos });
      } else if (o.type === 'media' && o.src?.startsWith('data:')) {
        slide.addMedia({ type: o.kind === 'audio' ? 'audio' : 'video', data: o.src.replace(/^data:/, ''), ...pos });
      } else if (o.type === 'table') {
        const rows = o.rows.map((r) => r.filter((c) => !c.hidden).map((c) => ({ text: c.text, options: { fill: c.bg ? { color: c.bg.replace('#', '') } : undefined, color: c.color ? c.color.replace('#', '') : undefined, bold: c.bold, colspan: c.cs > 1 ? c.cs : undefined, rowspan: c.rs > 1 ? c.rs : undefined } })));
        slide.addTable(rows, { ...pos, fontSize: st.size || 14, border: { type: 'solid', pt: 0.75, color: '9AA0B4' }, autoPage: false });
      }
    }
    const idx = pres.slides.indexOf(s);
    if (pres.master?.footer) slide.addText(pres.master.footer, { x: 0.4, y: H - 0.45, w: W * 0.6, h: 0.3, fontSize: 10, color: th.text.replace('#', ''), fontFace: th.bodyFont });
    if (pres.master?.showNumber) slide.addText(String(idx + 1), { x: W - 1.2, y: H - 0.45, w: 0.8, h: 0.3, fontSize: 10, align: 'right', color: th.text.replace('#', ''), fontFace: th.bodyFont });
    if (s.notes) slide.addNotes(s.notes);
  }
  void renderSlideImage;
  const ab = await p.write({ outputType: 'arraybuffer' });
  return addSuiteMeta(new Uint8Array(ab), pres);
}

// ------------------------------------------------------------------ datos propios de Impresiona dentro del .pptx
// PowerPoint no guarda todo lo que usa Impresiona (animaciones, trayectorias, Transformación, intervalos,
// patrón, secciones, subtítulos). Se guardan en docProps/impresiona.json junto con una huella de cada
// diapositiva; al abrir, solo se usan para las diapositivas que nadie cambió en otro programa.
const fnv = (str) => { let h = 0x811c9dc5; for (let i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 0x01000193); } return (h >>> 0).toString(16); };
const slideFiles = (zip) => Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)).sort((a, b) => Number(/(\d+)\.xml/.exec(a)[1]) - Number(/(\d+)\.xml/.exec(b)[1]));
async function addSuiteMeta(bytes, pres) {
  try {
    const zip = await JSZip.loadAsync(bytes);
    const hashes = await Promise.all(slideFiles(zip).map(async (f) => fnv(await zip.file(f).async('string'))));
    const strip = (o) => (o.src ? { ...o, src: '@' } : o);
    const meta = { app: 'Impresiona', v: 1, hashes, pres: { ...pres, master: pres.master ? { ...pres.master, objects: pres.master.objects.map(strip) } : pres.master, slides: pres.slides.map((s) => ({ ...s, bgImage: s.bgImage ? '@' : s.bgImage, objects: s.objects.map(strip) })) } };
    zip.file('docProps/impresiona.json', JSON.stringify(meta));
    let ct = await zip.file('[Content_Types].xml').async('string');
    if (!/Extension="json"/.test(ct)) ct = ct.replace(/(<Types[^>]*>)/, '$1<Default Extension="json" ContentType="application/json"/>');
    zip.file('[Content_Types].xml', ct);
    return await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  } catch (e) { console.warn('meta', e); return bytes; }
}
async function applySuiteMeta(bytes, imported) {
  try {
    const zip = await JSZip.loadAsync(bytes);
    const f = zip.file('docProps/impresiona.json'); if (!f) return null;
    const meta = JSON.parse(await f.async('string'));
    const hashes = await Promise.all(slideFiles(zip).map(async (x) => fnv(await zip.file(x).async('string'))));
    if (!meta?.pres?.slides || hashes.length !== imported.slides.length) return null;
    const mp = meta.pres; const mCount = (mp.master?.objects || []).filter((o) => o.src).length;
    const fill = (objs, srcs) => { let k = 0; for (const o of objs) if (o.src === '@') o.src = srcs[k++] || ''; };
    if (mp.master) fill(mp.master.objects, (imported.slides[0]?.objects || []).filter((o) => o.src).map((o) => o.src).slice(0, mCount));
    const out = [];
    for (let i = 0; i < hashes.length; i++) {
      const ms = mp.slides[i];
      if (ms && meta.hashes[i] === hashes[i]) {
        fill(ms.objects, imported.slides[i].objects.filter((o) => o.src).map((o) => o.src).slice(mCount));
        if (ms.bgImage === '@') ms.bgImage = imported.slides[i].bgImage || null;
        ms.objects = ms.objects.filter((o) => o.src !== '');
        out.push(ms);
      } else {
        // diapositiva cambiada fuera de Impresiona: se usa lo importado (sin el patrón horneado)
        const s = imported.slides[i]; if (mCount) s.objects = s.objects.slice((mp.master?.objects || []).length); out.push(s);
      }
    }
    return { ...mp, size: imported.size, slides: out };
  } catch (e) { console.warn('meta', e); return null; }
}
