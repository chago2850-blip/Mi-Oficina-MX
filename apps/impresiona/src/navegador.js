// Impresiona — parte 3 (como LibreOffice Impress): álbum fotográfico, navegador de diapositivas y
// objetos, nombre y texto alternativo de los objetos (accesibilidad, también en .pptx y .odp) y
// exportar la presentación como página web (HTML) que se navega con el teclado. GPL-3.0
import { h, t, tf, toast, modal, esc, api } from '../../shared/office.js';
import { tabbed, field, sel, chk } from '../../shared/dialogs.js';

let C = null; // { getPres, cur, setCur, goto, slide, commit, select, selObjs, newSlide, uid, buildSteps, docName, docProps }
export function init(ctx) { C = ctx; }
const P = () => C.getPres();
const plain = (html) => { const d = document.createElement('div'); d.innerHTML = html || ''; return (d.textContent || '').trim(); };
const KIND = { text: 'Texto', shape: 'Forma', image: 'Imagen', table: 'Tabla', chart: 'Gráfico', media: 'Multimedia', camera: 'Cámara' };
export const objName = (o, k) => o.name || `${t(KIND[o.type] || 'Objeto')} ${k + 1}${o.type === 'text' && plain(o.html) ? ` — ${plain(o.html).slice(0, 30)}` : ''}`;

// ================================================================ álbum fotográfico
const readImg = (src) => new Promise((res) => { const im = new Image(); im.onload = () => res({ w: im.naturalWidth, h: im.naturalHeight }); im.onerror = () => res({ w: 4, h: 3 }); im.src = src; });
export async function photoAlbumDialog() {
  const files = []; const list = h('div', { class: 'cs-rows album-list' });
  const draw = () => list.replaceChildren(...files.map((f, i) => h('div', { class: 'cs-row' }, h('span', {}, `${i + 1}. ${f.name}`),
    h('button', { type: 'button', class: 'btn sm', 'aria-label': t('Subir'), onclick: () => { if (i > 0) { [files[i - 1], files[i]] = [files[i], files[i - 1]]; draw(); } } }, '↑'),
    h('button', { type: 'button', class: 'btn sm', 'aria-label': t('Quitar'), onclick: () => { files.splice(i, 1); draw(); } }, '×'))));
  const input = h('input', { type: 'file', accept: 'image/*', multiple: true, hidden: true });
  input.addEventListener('change', async () => {
    for (const f of input.files) { const src = await new Promise((res) => { const rd = new FileReader(); rd.onload = () => res(rd.result); rd.readAsDataURL(f); }); files.push({ name: f.name.replace(/\.[^.]+$/, ''), src }); }
    input.value = ''; draw();
  });
  const body = await tabbed('Crear álbum fotográfico', [{ label: 'Álbum', body: h('div', { class: 'fx-col' },
    h('button', { type: 'button', class: 'btn', onclick: () => input.click() }, t('Agregar imágenes…')), input, list,
    field('Diapositiva', sel('per', '1', [['1', '1 imagen'], ['2', '2 imágenes'], ['4', '4 imágenes']])),
    chk('caption', 'Agregar el nombre del archivo como leyenda', true), chk('title', 'Agregar una diapositiva de título', true),
    field('Título del álbum', h('input', { name: 'albumTitle', value: t('Álbum fotográfico') }))) }], { width: '520px' });
  if (!body || !files.length) return null;
  const v = (n) => { const el = body.querySelector(`[name="${n}"]`); return el.type === 'checkbox' ? el.checked : el.value; };
  return photoAlbum(files, { per: Number(v('per')), caption: v('caption'), title: v('title') ? v('albumTitle') : '' });
}
export async function photoAlbum(images, { per = 1, caption = true, title = '' } = {}) {
  const { w: W, h: H } = P().size; const m = 30;
  const dims = await Promise.all(images.map((im) => readImg(im.src)));
  const slides = [];
  if (title) { const s = C.newSlide('title'); const tt = s.objects.find((o) => o.role === 'title'); if (tt) tt.html = esc(title); const sub = s.objects.find((o) => o.role === 'subtitle'); if (sub) sub.html = esc(tf('{0} fotos', images.length)); slides.push(s); }
  const cols = per === 1 ? 1 : 2; const rows = per === 4 ? 2 : 1; const capH = caption ? 34 : 0;
  for (let i = 0; i < images.length; i += per) {
    const s = C.newSlide('blank'); s.objects = [];
    const cw = (W - m * (cols + 1)) / cols; const ch = (H - m * (rows + 1)) / rows;
    images.slice(i, i + per).forEach((im, k) => {
      const d = dims[i + k]; const cx = m + (k % cols) * (cw + m); const cy = m + Math.floor(k / cols) * (ch + m);
      const sc = Math.min(cw / d.w, (ch - capH) / d.h); const iw = Math.round(d.w * sc); const ih = Math.round(d.h * sc);
      s.objects.push({ id: C.uid(), type: 'image', src: im.src, x: Math.round(cx + (cw - iw) / 2), y: Math.round(cy + (ch - capH - ih) / 2), w: iw, h: ih, rot: 0, name: im.name, alt: { desc: im.name } });
      if (caption) s.objects.push({ id: C.uid(), type: 'text', x: Math.round(cx), y: Math.round(cy + ch - capH), w: Math.round(cw), h: capH, rot: 0, html: `<p>${esc(im.name)}</p>`, style: { size: 14, align: 'center', valign: 'middle' } });
    });
    slides.push(s);
  }
  const at = C.cur() + 1;
  C.commit(t('Álbum fotográfico'), () => { P().slides.splice(at, 0, ...slides); C.setCur(at); });
  toast(tf('Se crearon {0} diapositivas', slides.length));
  return slides.length;
}

// ================================================================ navegador
export async function navigator() {
  const pres = P(); const q = h('input', { type: 'search', placeholder: t('Buscar un objeto o diapositiva'), 'aria-label': t('Buscar') });
  const tree = h('div', { class: 'nav-tree', role: 'tree' });
  let close = null;
  const draw = () => {
    const f = q.value.trim().toLowerCase(); tree.replaceChildren();
    pres.slides.forEach((s, i) => {
      const title = plain(s.objects.find((o) => o.role === 'title')?.html) || t('(sin título)');
      const objs = s.objects.map((o, k) => ({ o, k, name: objName(o, k) })).filter((x) => !f || x.name.toLowerCase().includes(f) || title.toLowerCase().includes(f));
      if (f && !objs.length && !title.toLowerCase().includes(f)) return;
      tree.append(h('button', { type: 'button', class: `nav-slide${i === C.cur() ? ' on' : ''}`, role: 'treeitem', onclick: () => { C.goto(i); close?.(); } }, `${i + 1}. ${title}${s.hidden ? ` (${t('oculta')})` : ''}`));
      for (const x of objs) tree.append(h('button', { type: 'button', class: 'nav-obj', role: 'treeitem', 'data-id': x.o.id, onclick: () => { C.goto(i); C.select([x.o.id]); close?.(); } }, x.name));
    });
  };
  q.addEventListener('input', draw); draw();
  const body = h('div', { class: 'fx-col' }, q, tree);
  const p = modal({ title: t('Navegador'), body, width: '440px', buttons: [{ label: t('Cerrar'), value: true, primary: true }], onOpen: () => q.focus() });
  close = () => document.querySelector('.modal-back:last-of-type footer .btn.primary')?.click();
  return p;
}

// ================================================================ nombre y texto alternativo
export async function nameDialog() {
  const o = C.selObjs()[0]; if (!o) { toast('Selecciona un objeto.', { error: true }); return null; }
  const body = await tabbed('Nombre y texto alternativo', [{ label: 'Objeto', body: h('div', { class: 'fx-col' },
    field('Nombre', h('input', { name: 'name', value: o.name || '' })),
    field('Título (texto alternativo)', h('input', { name: 'altTitle', value: o.alt?.title || '' })),
    field('Descripción', h('textarea', { name: 'altDesc', rows: 3, style: { width: '100%', fontFamily: 'inherit' } }, o.alt?.desc || '')),
    chk('decorative', 'Decorativo (los lectores de pantalla lo omiten)', !!o.alt?.decorative)) }], { width: '460px' });
  if (!body) return null;
  const v = (n) => { const el = body.querySelector(`[name="${n}"]`); return el.type === 'checkbox' ? el.checked : el.value.trim(); };
  const alt = {}; if (v('altTitle')) alt.title = v('altTitle'); if (v('altDesc')) alt.desc = v('altDesc'); if (v('decorative')) alt.decorative = true;
  const patch = { name: v('name') || undefined, alt: Object.keys(alt).length ? alt : undefined };
  C.commit(t('Nombre del objeto'), () => { const x = C.slide().objects.find((q) => q.id === o.id); Object.assign(x, patch); if (!patch.name) delete x.name; if (!patch.alt) delete x.alt; });
  return patch;
}
// imágenes sin texto alternativo (revisión de accesibilidad)
export function missingAlt() {
  const out = []; P().slides.forEach((s, i) => s.objects.forEach((o) => { if ((o.type === 'image' || o.type === 'chart') && !o.alt?.desc && !o.alt?.title && !o.alt?.decorative) out.push({ slide: i, id: o.id }); }));
  return out;
}

// ================================================================ exportar como página web
export async function exportHtml({ save = true } = {}) {
  if (!api?.renderFrames) return null;
  const pres = P(); const props = C.docProps?.() || {};
  toast('Creando la página web…');
  const states = pres.slides.map((s, i) => ({ i, step: C.buildSteps(s).groups.length }));
  const r = await api.renderFrames({ pres: JSON.parse(JSON.stringify(pres)), states, width: Math.min(1600, pres.size.w * 1.5) });
  const b64 = (u8) => { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s); };
  const slides = pres.slides.map((s, i) => ({ img: `data:image/png;base64,${b64(r.frames[i])}`, title: plain(s.objects.find((o) => o.role === 'title')?.html) || tf('Diapositiva {0}', i + 1), notes: s.notes || '', alt: s.objects.map((o) => (o.alt?.desc ? o.alt.desc : (o.type === 'text' ? plain(o.html) : ''))).filter(Boolean).join('. ') })).filter((_, i) => !pres.slides[i].hidden);
  const title = props.title || C.docName?.().replace(/\.[^.]+$/, '') || 'Presentación';
  const html = `<!doctype html>
<html lang="${document.documentElement.lang || 'es'}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(title)}</title>
<meta name="generator" content="Impresiona · Mi Oficina MX">${props.author ? `<meta name="author" content="${esc(props.author)}">` : ''}
<style>
:root{color-scheme:light dark}*{box-sizing:border-box}body{margin:0;font:15px/1.5 system-ui,sans-serif;background:#1b1d26;color:#eee;display:grid;grid-template-columns:240px 1fr;height:100vh}
nav{overflow:auto;background:#12131a;padding:12px}nav h1{font-size:15px;margin:4px 6px 12px}nav ol{margin:0;padding:0 0 0 22px}nav a{color:#ccd;text-decoration:none;display:block;padding:3px 4px;border-radius:4px}nav a[aria-current]{background:#3a3d52;color:#fff}
main{display:flex;flex-direction:column;min-width:0}figure{flex:1;margin:0;display:flex;align-items:center;justify-content:center;padding:16px;min-height:0}figure img{max-width:100%;max-height:100%;box-shadow:0 4px 18px #0008;background:#fff}
.bar{display:flex;gap:10px;align-items:center;justify-content:center;padding:8px}.bar button{font:inherit;padding:6px 14px;border-radius:6px;border:1px solid #555;background:#2a2c3a;color:#eee;cursor:pointer}
.notes{max-height:22vh;overflow:auto;padding:8px 20px;border-top:1px solid #333;color:#bbb;white-space:pre-wrap}.notes:empty{display:none}
@media (max-width:700px){body{grid-template-columns:1fr;grid-template-rows:auto 1fr}nav{max-height:30vh}}
</style></head><body>
<nav aria-label="Diapositivas"><h1>${esc(title)}</h1><ol>${slides.map((s, i) => `<li><a href="#${i + 1}">${esc(s.title)}</a></li>`).join('')}</ol></nav>
<main><figure><img id="img" alt=""></figure><div class="bar"><button id="prev" aria-label="Anterior">◀</button><span id="cnt"></span><button id="next" aria-label="Siguiente">▶</button></div><div class="notes" id="notes"></div></main>
<script type="application/json" id="data">${JSON.stringify(slides).replace(/</g, '\\u003c')}</script>
<script>
const S=JSON.parse(document.getElementById('data').textContent);let i=0;
function show(k){i=Math.max(0,Math.min(S.length-1,k));const s=S[i];const im=document.getElementById('img');im.src=s.img;im.alt=s.title+(s.alt?'. '+s.alt:'');
document.getElementById('cnt').textContent=(i+1)+' / '+S.length;document.getElementById('notes').textContent=s.notes;
document.querySelectorAll('nav a').forEach((a,k)=>{if(k===i)a.setAttribute('aria-current','page');else a.removeAttribute('aria-current');});if(location.hash!=='#'+(i+1))history.replaceState(null,'','#'+(i+1));}
document.getElementById('prev').onclick=()=>show(i-1);document.getElementById('next').onclick=()=>show(i+1);
addEventListener('keydown',(e)=>{if(['ArrowRight','PageDown',' '].includes(e.key)){e.preventDefault();show(i+1);}if(['ArrowLeft','PageUp'].includes(e.key)){e.preventDefault();show(i-1);}if(e.key==='Home')show(0);if(e.key==='End')show(S.length-1);});
addEventListener('hashchange',()=>show((parseInt(location.hash.slice(1))||1)-1));show((parseInt(location.hash.slice(1))||1)-1);
</script></body></html>`;
  if (!save) return html;
  const res = await api.save({ text: html, ask: true, ext: 'html', filterName: t('Página web'), suggestedName: `${title}.html` });
  if (res) toast(tf('Exportado: {0}', res.name), { action: { label: t('Mostrar'), run: () => api.showInFolder(res.path) } });
  return res;
}
