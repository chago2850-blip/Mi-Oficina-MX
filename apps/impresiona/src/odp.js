// Presentaciones ODF (.odp): abrir y guardar. GPL-3.0
import JSZip from 'jszip';
import { themeById } from './themes.js';
import { chartSvg } from './charts.js';

const NS = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0', style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0', table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0', fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  xlink: 'http://www.w3.org/1999/xlink', dc: 'http://purl.org/dc/elements/1.1/', meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0', presentation: 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0',
  smil: 'urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0', anim: 'urn:oasis:names:tc:opendocument:xmlns:animation:1.0',
};
const fnv = (str) => { let h = 0x811c9dc5; for (let i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 0x01000193); } return (h >>> 0).toString(16); };
NS.script = 'urn:oasis:names:tc:opendocument:xmlns:script:1.0';
// tamaño original de una imagen (fo:clip se mide en el tamaño original, a 96 ppp)
const natural = (src) => new Promise((res) => { try { const im = new Image(); im.onload = () => res({ w: im.naturalWidth, h: im.naturalHeight }); im.onerror = () => res(null); im.src = src; } catch { res(null); } });
const NSDECL = Object.entries(NS).map(([k, v]) => `xmlns:${k}="${v}"`).join(' ');
const x = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const cm = (px) => `${(px / 37.7953).toFixed(3)}cm`;
const A = (el, ns, name) => el.getAttributeNS(NS[ns], name) || el.getAttribute(`${ns}:${name}`);
const lenToPx = (v) => { if (!v) return 0; const n = parseFloat(v); if (v.endsWith('cm')) return n * 37.7953; if (v.endsWith('mm')) return n * 3.77953; if (v.endsWith('in')) return n * 96; if (v.endsWith('pt')) return n * 4 / 3; return n; };
const rgb = (c) => { if (!c) return null; if (c.startsWith('#')) return c.slice(0, 7); const m = c.match(/rgba?\(([^)]+)\)/); if (!m) return null; const p = m[1].split(',').map(parseFloat); return `#${p.slice(0, 3).map((n) => Math.round(n).toString(16).padStart(2, '0')).join('')}`; };
const SHAPE_ODF = { rect: 'rectangle', roundRect: 'round-rectangle', ellipse: 'ellipse', triangle: 'isosceles-triangle', rtTriangle: 'right-triangle', arrow: 'right-arrow', star: 'star5', star4: 'star4', diamond: 'diamond', parallelogram: 'parallelogram', trapezoid: 'trapezoid', pentagon: 'pentagon', hexagon: 'hexagon', octagon: 'octagon', plus: 'cross', heart: 'heart', callout: 'rectangular-callout', cloud: 'cloud', leftArrow: 'left-arrow', upArrow: 'up-arrow', downArrow: 'down-arrow', leftRightArrow: 'left-right-arrow', chevron: 'chevron', homePlate: 'pentagon-right' };
// guiones relativos al grosor de la línea, como los estilos de LibreOffice (así se exportan como guiones estándar)
const DASH_ODF = { dash: ['Dash', 'rect', 1, '400%', 0, '0', '300%'], dot: ['Dot', 'rect', 1, '100%', 0, '0', '300%'], dashDot: ['Dash Dot', 'rect', 1, '400%', 1, '100%', '300%'], longDash: ['Long Dash', 'rect', 1, '800%', 0, '0', '300%'] };
// puntas con los nombres de LibreOffice
const MARKER_ODF = { triangle: ['Arrow', '0 0 20 30', 'M10 0l-10 30h20z'], open: ['Line Arrow', '0 0 1122 2243', 'M0 2108v17 17l12 42 30 34 38 21 43 4 29-8 30-21 25-26 13-34 343-1532 339 1520 13 42 29 34 39 21 42 4 42-12 34-30 21-42v-39-12l-4 4-440-1998-9-42-25-39-38-25-43-8-42 8-38 25-26 39-8 42z'], circle: ['Circle', '0 0 1131 1131', 'M462 1118l-102-29-102-51-93-72-72-93-51-102-29-102-13-105 13-102 29-106 51-102 72-89 93-72 102-50 102-34 106-9 101 9 106 34 98 50 93 72 72 89 51 102 29 106 13 102-13 105-29 102-51 102-72 93-93 72-98 51-106 29-101 13z'], diamond: ['Square 45', '0 0 1131 1131', 'M566 0l565 566-565 565-566-565z'] };
const OOXML_IN = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', rtTriangle: 'rtTriangle', rightArrow: 'arrow', star5: 'star', star4: 'star4', diamond: 'diamond', parallelogram: 'parallelogram', trapezoid: 'trapezoid', pentagon: 'pentagon', hexagon: 'hexagon', octagon: 'octagon', plus: 'plus', heart: 'heart', wedgeRectCallout: 'callout', wedgeRoundRectCallout: 'callout', cloud: 'cloud', leftArrow: 'leftArrow', upArrow: 'upArrow', downArrow: 'downArrow', leftRightArrow: 'leftRightArrow', chevron: 'chevron', homePlate: 'homePlate' };
const MARKER_IN = { ...Object.fromEntries(Object.entries(MARKER_ODF).map(([k, v]) => [v[0], k])), Flecha: 'triangle', FlechaAbierta: 'open', Circulo: 'circle', Rombo: 'diamond' };
const DASH_IN = { ...Object.fromEntries(Object.entries(DASH_ODF).map(([k, v]) => [v[0], k])), Guiones: 'dash', Puntos: 'dot', GuionPunto: 'dashDot', GuionLargo: 'longDash' };
const SHAPE_IN = Object.fromEntries(Object.entries(SHAPE_ODF).map(([k, v]) => [v, k]));

export async function writeOdp(pres, props = {}) {
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.presentation', { compression: 'STORE' });
  const th = themeById(pres.theme);
  const naturals = new Map();
  for (const sl of [...pres.slides, ...(pres.master ? [pres.master] : [])]) for (const o of sl.objects) if (o.type === 'image' && o.crop && !naturals.has(o.src)) naturals.set(o.src, await natural(o.src));
  const autos = []; const keys = new Map(); let pic = 0; const pics = [];
  const auto = (family, inner, extra = '') => { const k = `${family}|${inner}|${extra}`; if (!keys.has(k)) { const n = `${family === 'paragraph' ? 'P' : family === 'text' ? 'T' : family === 'drawing-page' ? 'dp' : 'gr'}${keys.size + 1}`; keys.set(k, n); autos.push(`<style:style style:name="${n}" style:family="${family}"${extra}>${inner}</style:style>`); } return keys.get(k); };
  const runs = (node, f) => {
    let out = '';
    for (const n of node.childNodes) {
      if (n.nodeType === 3) { const t = n.nodeValue; if (!t) continue; const props2 = []; if (f.bold) props2.push('fo:font-weight="bold"'); if (f.italic) props2.push('fo:font-style="italic"'); if (f.underline) props2.push('style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"'); if (f.color) props2.push(`fo:color="${f.color}"`); if (f.size) props2.push(`fo:font-size="${f.size}pt"`); if (f.font) props2.push(`fo:font-family="'${x(f.font)}'"`); const body = x(t).replace(/ {2,}/g, (m) => ` <text:s text:c="${m.length - 1}"/>`); out += props2.length ? `<text:span text:style-name="${auto('text', `<style:text-properties ${props2.join(' ')}/>`)}">${body}</text:span>` : body; continue; }
      if (n.nodeType !== 1) continue;
      if (n.tagName === 'BR') { out += '<text:line-break/>'; continue; }
      if (n.tagName === 'UL' || n.tagName === 'OL') continue;
      const g = { ...f };
      if (['B', 'STRONG'].includes(n.tagName)) g.bold = true; if (['I', 'EM'].includes(n.tagName)) g.italic = true; if (n.tagName === 'U') g.underline = true;
      const st = n.style; if (st) { if (st.fontWeight === 'bold' || Number(st.fontWeight) >= 600) g.bold = true; if (st.fontStyle === 'italic') g.italic = true; if (st.color) g.color = rgb(st.color); if (st.fontSize) g.size = st.fontSize.endsWith('pt') ? parseFloat(st.fontSize) : parseFloat(st.fontSize) * 0.75; if (st.fontFamily) g.font = st.fontFamily.split(',')[0].replace(/["']/g, '').trim(); }
      out += runs(n, g);
    }
    return out;
  };
  const paras = (html, base, align, extra = '') => {
    const d = new DOMParser().parseFromString(`<body>${html || ''}</body>`, 'text/html');
    const pStyle = (el) => auto('paragraph', `<style:paragraph-properties fo:text-align="${{ center: 'center', right: 'end', justify: 'justify' }[el?.style?.textAlign || align] || 'start'}"${extra}/>`);
    let out = '';
    const walk = (parent) => {
      for (const el of parent.children) {
        if (el.tagName === 'UL' || el.tagName === 'OL') { out += `<text:list>${[...el.children].map((li) => `<text:list-item><text:p text:style-name="${pStyle(li)}">${runs(li, base)}</text:p></text:list-item>`).join('')}</text:list>`; }
        else out += `<text:p text:style-name="${pStyle(el)}">${runs(el, base)}</text:p>`;
      }
    };
    if (!d.body.children.length) out = d.body.textContent ? `<text:p text:style-name="${pStyle(null)}">${runs(d.body, base)}</text:p>` : '';
    else walk(d.body);
    return out;
  };
  // definiciones con nombre (degradados, guiones, puntas de flecha) que se ponen en styles.xml
  const defs = new Map();
  const defName = (kind, key, xml) => { const k = `${kind}|${key}`; if (!defs.has(k)) defs.set(k, xml(`${kind}${defs.size + 1}`)); return /draw:name="([^"]+)"/.exec(defs.get(k))[1]; };
  const graphicProps = (o, st, { fillDef, text = false } = {}) => {
    const p = [];
    const isLine = o.shape === 'line' || o.shape === 'lineArrow';
    const fill = isLine ? null : st.fill === undefined ? fillDef : st.fill;
    if (st.grad && fill) {
      const g = st.grad; const gn = defName('Degradado', JSON.stringify([fill, g]), (n) => `<draw:gradient draw:name="${n}" draw:style="${g.type === 'radial' ? 'radial' : 'linear'}" draw:start-color="${fill}" draw:end-color="${g.c2 || '#ffffff'}" draw:start-intensity="100%" draw:end-intensity="100%" draw:angle="${Math.round((((90 - (g.angle ?? 90)) % 360) + 360) % 360) * 10}" draw:border="0%"${g.type === 'radial' ? ' draw:cx="50%" draw:cy="50%"' : ''}/>`);
      p.push(`draw:fill="gradient" draw:fill-gradient-name="${gn}"`);
    } else p.push(fill ? `draw:fill="solid" draw:fill-color="${fill}"` : 'draw:fill="none"');
    if (st.fillOpacity != null || (st.opacity != null && o.type !== 'image')) p.push(`draw:opacity="${Math.round((st.fillOpacity ?? 1) * (st.opacity ?? 1) * 100)}%"`);
    const stroke = isLine ? (st.stroke || th.accent) : st.stroke;
    if (stroke) {
      const dn = DASH_ODF[st.dash] && defName('Guion', st.dash, () => { const d = DASH_ODF[st.dash]; return `<draw:stroke-dash draw:name="${d[0].replace(/ /g, '_20_')}" draw:display-name="${d[0]}" draw:style="${d[1]}" draw:dots1="${d[2]}" draw:dots1-length="${d[3]}"${d[4] ? ` draw:dots2="${d[4]}" draw:dots2-length="${d[5]}"` : ''} draw:distance="${d[6]}"/>`; });
      p.push(`draw:stroke="${dn ? 'dash' : 'solid'}"${dn ? ` draw:stroke-dash="${dn}"` : ''} svg:stroke-color="${stroke}" svg:stroke-width="${cm(st.strokeW || (isLine ? 4 : 2))}"`);
      if (st.strokeOpacity != null) p.push(`svg:stroke-opacity="${Math.round(st.strokeOpacity * 100)}%"`);
      if (isLine) for (const [end, kind] of [['start', st.arrowStart], ['end', st.arrowEnd || (o.shape === 'lineArrow' ? 'triangle' : null)]]) {
        const M = MARKER_ODF[kind]; if (!M) continue;
        const mn = defName('Punta', kind, () => `<draw:marker draw:name="${M[0].replace(/ /g, '_20_')}" draw:display-name="${M[0]}" svg:viewBox="${M[1]}" svg:d="${M[2]}"/>`);
        p.push(`draw:marker-${end}="${mn}" draw:marker-${end}-width="${cm(Math.max(8, (st.strokeW || 4) * 3))}"`);
      }
    } else p.push('draw:stroke="none"');
    if (st.shadow) { const sh = st.shadow; p.push(`draw:shadow="visible" draw:shadow-offset-x="${cm(sh.dx ?? 4)}" draw:shadow-offset-y="${cm(sh.dy ?? 4)}" draw:shadow-color="${sh.color || '#000000'}" draw:shadow-opacity="${Math.round((sh.alpha ?? 0.45) * 100)}%"`); }
    if (text || o.html) {
      p.push(`draw:textarea-vertical-align="${{ middle: 'middle', bottom: 'bottom' }[st.valign] || (o.type === 'shape' ? 'middle' : 'top')}"`);
      const pd = st.pad; p.push(pd ? `fo:padding-left="${cm(pd.l)}" fo:padding-right="${cm(pd.r)}" fo:padding-top="${cm(pd.t)}" fo:padding-bottom="${cm(pd.b)}"` : 'fo:padding="0.15cm"');
      if (st.autofit === 'shrink') p.push('draw:fit-to-size="shrink-to-fit" style:shrink-to-fit="true"');
    }
    if (o.type === 'image') {
      const f = o.img || {};
      if (f.gray) p.push('draw:color-mode="greyscale"'); if (f.bright != null && f.bright !== 100) p.push(`draw:luminance="${Math.round(f.bright - 100)}%"`); if (f.contrast != null && f.contrast !== 100) p.push(`draw:contrast="${Math.round(f.contrast - 100)}%"`);
      if (st.opacity != null) p.push(`draw:image-opacity="${Math.round(st.opacity * 100)}%"`);
      if (o.flipH || o.flipV) p.push(`style:mirror="${[o.flipH && 'horizontal', o.flipV && 'vertical'].filter(Boolean).join(' ')}"`);
      if (o.crop) { const c = o.crop; const nat = naturals.get(o.src); const W0 = nat ? nat.w : o.w / ((100 - c.l - c.r) / 100), H0 = nat ? nat.h : o.h / ((100 - c.t - c.b) / 100); p.push(`fo:clip="rect(${cm(H0 * c.t / 100)}, ${cm(W0 * c.r / 100)}, ${cm(H0 * c.b / 100)}, ${cm(W0 * c.l / 100)})"`); }
    }
    return p.join(' ');
  };
  const paraExtra = (st) => `${st.lineH ? ` fo:line-height="${Math.round((st.lineH / 1.2) * 100)}%"` : ''}${st.paraSp != null ? ` fo:margin-bottom="${(st.paraSp / 28.35).toFixed(3)}cm"` : ''}`;
  // acción al hacer clic
  const evt = (o, slidesArr) => {
    const l = o.link; if (!l) return '';
    let a;
    if (l.startsWith('#slide:')) { const i = slidesArr.findIndex((s) => s.id === l.slice(7)); a = `presentation:action="show" xlink:href="#Diapositiva${i + 1}" xlink:type="simple"`; }
    else if (l.startsWith('#')) a = `presentation:action="${{ next: 'next-page', prev: 'previous-page', first: 'first-page', last: 'last-page', end: 'stop' }[l.slice(1)] || 'none'}"`;
    else a = `presentation:action="show" xlink:href="${x(l)}" xlink:type="simple"`;
    return `<office:event-listeners><presentation:event-listener script:event-name="dom:click" ${a}/></office:event-listeners>`;
  };
  const image = (src, name) => { const m = /^data:image\/([a-z+]+);base64,(.*)$/i.exec(src || ''); if (!m) return null; const ext = m[1].toLowerCase().replace('jpeg', 'jpg').replace('svg+xml', 'svg'); const p = `Pictures/${name || `img${++pic}`}.${ext}`; zip.file(p, m[2], { base64: true }); pics.push(p); return p; };
  const objXml = (o, idx) => {
    const st = o.style || {};
    const pos = `svg:x="${cm(o.x)}" svg:y="${cm(o.y)}" svg:width="${cm(o.w)}" svg:height="${cm(o.h)}"${o.rot ? ` draw:transform="rotate(${(-o.rot * Math.PI / 180).toFixed(5)}) translate(${cm(o.x)} ${cm(o.y)})"` : ''}`;
    const posNoRot = o.rot ? `svg:width="${cm(o.w)}" svg:height="${cm(o.h)}" draw:transform="rotate(${(-o.rot * Math.PI / 180).toFixed(5)}) translate(${cm(o.x)} ${cm(o.y)})"` : pos;
    const isTitle = o.role === 'title';
    const base = { color: st.color || (o.type === 'shape' ? '#FFFFFF' : isTitle ? th.title : th.text), size: st.size || (o.type === 'shape' ? 18 : 20), bold: st.bold, italic: st.italic, font: st.font || (isTitle ? th.titleFont : th.bodyFont) };
    if (o.type === 'text') {
      const gr = auto('graphic', `<style:graphic-properties ${graphicProps(o, { ...st, fill: st.fill || null }, { text: true })}/>`);
      return `<draw:frame draw:style-name="${gr}" ${posNoRot}${o.role ? ` presentation:class="${o.role === 'body' ? 'outline' : o.role}"` : ''}><draw:text-box>${paras(o.html, base, st.align, paraExtra(st))}</draw:text-box>${evt(o, pres.slides)}</draw:frame>`;
    }
    if (o.type === 'shape' && (o.shape === 'line' || o.shape === 'lineArrow')) {
      const gr = auto('graphic', `<style:graphic-properties ${graphicProps(o, st)}/>`);
      const ang = ((o.rot || 0) * Math.PI) / 180; const cx = o.x + o.w / 2, cy = o.y + o.h / 2; const dx = (Math.cos(ang) * o.w) / 2, dy = (Math.sin(ang) * o.w) / 2;
      const [x1, x2] = o.flipH ? [cx + dx, cx - dx] : [cx - dx, cx + dx]; const [y1, y2] = o.flipH ? [cy + dy, cy - dy] : [cy - dy, cy + dy];
      return `<draw:line draw:style-name="${gr}" svg:x1="${cm(x1)}" svg:y1="${cm(y1)}" svg:x2="${cm(x2)}" svg:y2="${cm(y2)}">${evt(o, pres.slides)}</draw:line>`;
    }
    if (o.type === 'shape') {
      const gr = auto('graphic', `<style:graphic-properties ${graphicProps(o, st, { fillDef: th.accent })}/>`);
      return `<draw:custom-shape draw:style-name="${gr}" ${posNoRot}>${paras(o.html, base, st.align || 'center', paraExtra(st))}${evt(o, pres.slides)}<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="${SHAPE_ODF[o.shape] || 'rectangle'}"${o.flipH ? ' draw:mirror-horizontal="true"' : ''}${o.flipV ? ' draw:mirror-vertical="true"' : ''}/></draw:custom-shape>`;
    }
    if (o.type === 'image') { const p = image(o.src); const gr = auto('graphic', `<style:graphic-properties ${graphicProps(o, { ...(o.style || {}), fill: null })}/>`); return p ? `<draw:frame draw:style-name="${gr}" ${posNoRot}><draw:image xlink:href="${p}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>${evt(o, pres.slides)}</draw:frame>` : ''; }
    if (o.type === 'chart') { const svg = chartSvg(o, th); const p = `Pictures/grafico${idx}-${++pic}.svg`; zip.file(p, svg); pics.push(p); return `<draw:frame ${posNoRot} draw:name="Gráfico"><draw:image xlink:href="${p}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`; }
    if (o.type === 'table') {
      const cols = Math.max(...o.rows.map((r) => r.length));
      const rows = o.rows.map((r, ri) => `<table:table-row>${r.map((c) => (c.hidden ? '<table:covered-table-cell/>' : `<table:table-cell office:value-type="string"${c.cs > 1 ? ` table:number-columns-spanned="${c.cs}"` : ''}><text:p>${ri === 0 ? `<text:span text:style-name="${auto('text', '<style:text-properties fo:font-weight="bold"/>')}">${x(c.text)}</text:span>` : x(c.text)}</text:p></table:table-cell>`)).join('')}</table:table-row>`).join('');
      return `<draw:frame ${posNoRot}><table:table><table:table-column table:number-columns-repeated="${cols}"/>${rows}</table:table></draw:frame>`;
    }
    return '';
  };
  let pages = '';
  pres.slides.forEach((s, i) => {
    const bgImg = s.bgImage ? image(s.bgImage, `fondo${i + 1}`) : null;
    const dp = auto('drawing-page', bgImg ? `<style:drawing-page-properties draw:fill="bitmap" draw:fill-image-name="Fondo${i + 1}"${s.hidden ? ' presentation:visibility="hidden"' : ''}/>` : `<style:drawing-page-properties draw:fill="solid" draw:fill-color="${s.bg || th.bg}" presentation:transition-type="${s.transition?.type && s.transition.type !== 'none' ? 'automatic' : 'manual'}"${s.hidden ? ' presentation:visibility="hidden"' : ''}/>`);
    const master = (pres.master?.objects || []).map((o, k) => objXml(o, `m${k}`)).join('');
    // los objetos agrupados van dentro de un draw:g
    let objs = ''; const done = new Set();
    s.objects.forEach((o, k) => {
      if (done.has(o.id)) return;
      if (o.group) { const mem = s.objects.filter((m) => m.group === o.group); mem.forEach((m) => done.add(m.id)); objs += `<draw:g>${mem.map((m) => objXml(m, `${i}-${k}-${m.id}`)).join('')}</draw:g>`; }
      else objs += objXml(o, `${i}-${k}`);
    });
    const notes = s.notes ? `<presentation:notes><draw:frame presentation:class="notes" svg:x="2cm" svg:y="14cm" svg:width="17cm" svg:height="10cm"><draw:text-box>${s.notes.split('\n').map((l) => `<text:p>${x(l)}</text:p>`).join('')}</draw:text-box></draw:frame></presentation:notes>` : '';
    pages += `<draw:page draw:name="Diapositiva${i + 1}" draw:style-name="${dp}" draw:master-page-name="Default">${master}${objs}${notes}</draw:page>`;
  });
  const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NSDECL} office:version="1.3"><office:automatic-styles>${autos.join('')}</office:automatic-styles><office:body><office:presentation>${pages}</office:presentation></office:body></office:document-content>`;
  const styles = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NSDECL} office:version="1.3"><office:styles>${[...defs.values()].join('')}<style:default-style style:family="graphic"><style:text-properties fo:font-size="18pt" fo:language="es" fo:country="MX"/></style:default-style></office:styles><office:automatic-styles><style:page-layout style:name="PM1"><style:page-layout-properties fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:page-width="${cm(pres.size.w)}" fo:page-height="${cm(pres.size.h)}" style:print-orientation="landscape"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles></office:document-styles>`;
  const meta = `<?xml version="1.0" encoding="UTF-8"?><office:document-meta ${NSDECL} office:version="1.3"><office:meta><meta:generator>Impresiona · Mi Oficina MX</meta:generator><dc:title>${x(props.title)}</dc:title><dc:creator>${x(props.author)}</dc:creator><meta:user-defined meta:name="impresiona-tema">${x(pres.theme)}</meta:user-defined></office:meta></office:document-meta>`;
  zip.file('content.xml', content); zip.file('styles.xml', styles); zip.file('meta.xml', meta);
  // lo que el formato ODF no describe (animaciones, trayectorias, intervalos, secciones…): copia exacta para Impresiona, con la huella del contenido
  const strip = (o) => (o.src && o.type === 'image' ? { ...o, src: '@' } : o);
  zip.file('impresiona.json', JSON.stringify({ app: 'Impresiona', v: 1, hash: fnv(content), pres: { ...pres, master: pres.master ? { ...pres.master, objects: pres.master.objects.map(strip) } : pres.master, slides: pres.slides.map((sl) => ({ ...sl, objects: sl.objects.map(strip) })) } }));
  zip.file('META-INF/manifest.xml', `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="impresiona.json" manifest:media-type="application/json"/>${pics.map((p) => `<manifest:file-entry manifest:full-path="${p}" manifest:media-type="image/${p.split('.').pop().replace('jpg', 'jpeg').replace('svg', 'svg+xml')}"/>`).join('')}</manifest:manifest>`);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.oasis.opendocument.presentation' });
}

export async function readOdp(bytes, uid) {
  const zip = await JSZip.loadAsync(bytes);
  const t = await zip.file('content.xml')?.async('string'); if (!t) throw new Error('No es una presentación ODF válida');
  const doc = new DOMParser().parseFromString(t, 'application/xml');
  const sdoc = zip.file('styles.xml') ? new DOMParser().parseFromString(await zip.file('styles.xml').async('string'), 'application/xml') : null;
  const mdoc = zip.file('meta.xml') ? new DOMParser().parseFromString(await zip.file('meta.xml').async('string'), 'application/xml') : null;
  const smap = new Map(); const grads = new Map(), dashes = new Map(), markers = new Map();
  for (const d of [sdoc, doc].filter(Boolean)) {
    for (const g of d.getElementsByTagNameNS(NS.draw, 'gradient')) grads.set(A(g, 'draw', 'name'), { c1: A(g, 'draw', 'start-color'), c2: A(g, 'draw', 'end-color'), type: /radial|ellips/.test(A(g, 'draw', 'style') || '') ? 'radial' : 'linear', angle: ((90 - (parseFloat(A(g, 'draw', 'angle') || '0') / (/deg/.test(A(g, 'draw', 'angle') || '') ? 1 : 10))) % 360 + 360) % 360 });
    for (const g of d.getElementsByTagNameNS(NS.draw, 'stroke-dash')) {
      // por la forma del guion (los nombres cambian entre programas)
      const nm = A(g, 'draw', 'display-name') || A(g, 'draw', 'name'); const l1 = A(g, 'draw', 'dots1-length') || ''; const two = Number(A(g, 'draw', 'dots2') || 0) > 0;
      const len = l1.endsWith('%') ? parseFloat(l1) / 100 : lenToPx(l1) / Math.max(1, lenToPx(A(g, 'draw', 'distance')) / 3);
      dashes.set(A(g, 'draw', 'name'), two ? 'dashDot' : l1 ? (len <= 1.5 ? 'dot' : len >= 7 ? 'longDash' : 'dash') : DASH_IN[nm] || 'dash');
    }
    for (const g of d.getElementsByTagNameNS(NS.draw, 'marker')) { const nm = A(g, 'draw', 'display-name') || A(g, 'draw', 'name'); markers.set(A(g, 'draw', 'name'), MARKER_IN[nm] || (/circle|oval/i.test(nm) ? 'circle' : /square|diamond|rombo/i.test(nm) ? 'diamond' : /open|stealth/i.test(nm) ? 'open' : 'triangle')); }
  }
  for (const d of [sdoc, doc].filter(Boolean)) for (const st of d.getElementsByTagNameNS(NS.style, 'style')) {
    const o = {}; const name = A(st, 'style', 'name');
    const gp = st.getElementsByTagNameNS(NS.style, 'graphic-properties')[0], dp = st.getElementsByTagNameNS(NS.style, 'drawing-page-properties')[0], tp = st.getElementsByTagNameNS(NS.style, 'text-properties')[0], pp = st.getElementsByTagNameNS(NS.style, 'paragraph-properties')[0];
    if (gp) {
      const fk = A(gp, 'draw', 'fill');
      if (fk === 'solid') o.fill = A(gp, 'draw', 'fill-color'); if (fk === 'none') o.fill = null;
      if (fk === 'gradient') { const g = grads.get(A(gp, 'draw', 'fill-gradient-name')); if (g) { o.fill = g.c1; o.grad = { c2: g.c2, type: g.type, angle: g.angle }; } }
      const sk = A(gp, 'draw', 'stroke');
      if (sk === 'solid' || sk === 'dash') { o.stroke = A(gp, 'svg', 'stroke-color') || '#000000'; o.strokeW = lenToPx(A(gp, 'svg', 'stroke-width')) || 1; if (sk === 'dash') o.dash = dashes.get(A(gp, 'draw', 'stroke-dash')) || 'dash'; }
      if (sk === 'none') o.stroke = null;
      const so = A(gp, 'svg', 'stroke-opacity'); if (so && parseFloat(so) < 100) o.strokeOpacity = parseFloat(so) / 100;
      const op = A(gp, 'draw', 'opacity'); if (op && parseFloat(op) < 100) o.fillOpacity = parseFloat(op) / 100;
      for (const end of ['start', 'end']) { const m = A(gp, 'draw', `marker-${end}`); if (m) o[end === 'start' ? 'arrowStart' : 'arrowEnd'] = markers.get(m) || 'triangle'; }
      if (A(gp, 'draw', 'shadow') === 'visible') o.shadow = { color: A(gp, 'draw', 'shadow-color') || '#000000', dx: Math.round(lenToPx(A(gp, 'draw', 'shadow-offset-x')) || 4), dy: Math.round(lenToPx(A(gp, 'draw', 'shadow-offset-y')) || 4), blur: 6, alpha: A(gp, 'draw', 'shadow-opacity') ? parseFloat(A(gp, 'draw', 'shadow-opacity')) / 100 : 0.45 };
      const va = A(gp, 'draw', 'textarea-vertical-align'); if (va) o.valign = va;
      const pl = A(gp, 'fo', 'padding-left'); if (pl) o.pad = { l: Math.round(lenToPx(pl)), r: Math.round(lenToPx(A(gp, 'fo', 'padding-right'))), t: Math.round(lenToPx(A(gp, 'fo', 'padding-top'))), b: Math.round(lenToPx(A(gp, 'fo', 'padding-bottom'))) };
      if (A(gp, 'draw', 'fit-to-size') === 'shrink-to-fit' || A(gp, 'style', 'shrink-to-fit') === 'true') o.autofit = 'shrink';
      const img = {}; if (A(gp, 'draw', 'color-mode') === 'greyscale') img.gray = true; const lu = A(gp, 'draw', 'luminance'); if (lu && parseFloat(lu)) img.bright = 100 + parseFloat(lu); const co = A(gp, 'draw', 'contrast'); if (co && parseFloat(co)) img.contrast = 100 + parseFloat(co); if (Object.keys(img).length) o.img = img;
      const io = A(gp, 'draw', 'image-opacity'); if (io && parseFloat(io) < 100) o.opacity = parseFloat(io) / 100;
      const mir = A(gp, 'style', 'mirror'); if (mir && mir !== 'none') { if (/horizontal/.test(mir)) o.flipH = true; if (/vertical/.test(mir)) o.flipV = true; }
      const clip = A(gp, 'fo', 'clip'); if (clip) { const m = /rect\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/.exec(clip); if (m) o.clip = m.slice(1).map(lenToPx); }
    }
    if (dp && A(dp, 'presentation', 'visibility') === 'hidden') o.hidden = true;
    if (dp && (A(dp, 'draw', 'fill') === 'solid' || A(dp, 'draw', 'fill-color'))) o.bg = A(dp, 'draw', 'fill-color');
    if (tp) { if (A(tp, 'fo', 'font-weight') === 'bold') o.bold = true; if (A(tp, 'fo', 'font-style') === 'italic') o.italic = true; o.color = A(tp, 'fo', 'color') || undefined; const sz = A(tp, 'fo', 'font-size'); if (sz?.endsWith('pt')) o.size = parseFloat(sz); const ul = A(tp, 'style', 'text-underline-style'); if (ul && ul !== 'none') o.underline = true; }
    if (pp) { const al = A(pp, 'fo', 'text-align'); if (al) o.align = { center: 'center', end: 'right', right: 'right', justify: 'justify' }[al] || 'left'; const lh = A(pp, 'fo', 'line-height'); if (lh?.endsWith('%') && parseFloat(lh) !== 100) o.lineH = Math.round(parseFloat(lh) * 1.2) / 100; const mb = A(pp, 'fo', 'margin-bottom'); if (mb) o.paraSp = Math.round(lenToPx(mb) * 0.75); }
    smap.set(name, o);
  }
  let W = 960, H = 540;
  const pl = sdoc?.getElementsByTagNameNS(NS.style, 'page-layout-properties')[0];
  if (pl && A(pl, 'fo', 'page-width')) { W = Math.round(lenToPx(A(pl, 'fo', 'page-width'))); H = Math.round(lenToPx(A(pl, 'fo', 'page-height'))); }
  const k = 960 / W; // se normaliza a 960 px de ancho
  const e = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const inline = (node) => {
    let out = '';
    for (const n of node.childNodes) {
      if (n.nodeType === 3) { out += e(n.nodeValue); continue; }
      if (n.nodeType !== 1) continue;
      if (n.localName === 's') { out += ' '.repeat(Number(A(n, 'text', 'c') || 1)); continue; }
      if (n.localName === 'line-break') { out += '<br>'; continue; }
      if (n.localName === 'tab') { out += ' '; continue; }
      if (n.localName === 'span') { const st = smap.get(A(n, 'text', 'style-name')) || {}; const css = [st.bold && 'font-weight:bold', st.italic && 'font-style:italic', st.underline && 'text-decoration:underline', st.color && `color:${st.color}`, st.size && `font-size:${st.size}pt`].filter(Boolean).join(';'); out += css ? `<span style="${css}">${inline(n)}</span>` : inline(n); continue; }
      out += inline(n);
    }
    return out;
  };
  const textHtml = (el) => {
    let out = '';
    for (const c of el.children) {
      if (c.localName === 'p') { const st = smap.get(A(c, 'text', 'style-name')) || {}; out += `<p${st.align && st.align !== 'left' ? ` style="text-align:${st.align}"` : ''}>${inline(c) || '<br>'}</p>`; }
      else if (c.localName === 'list') out += `<ul>${[...c.getElementsByTagNameNS(NS.text, 'list-item')].map((li) => `<li>${[...li.children].filter((p) => p.localName === 'p').map(inline).join('<br>')}</li>`).join('')}</ul>`;
    }
    return out.replace(/^<p><br><\/p>$/, '');
  };
  const imgUrl = async (href) => { const f = zip.file(href); if (!f) return null; const ext = href.split('.').pop().toLowerCase(); const mime = { jpg: 'image/jpeg', jpeg: 'image/jpeg', svg: 'image/svg+xml', gif: 'image/gif' }[ext] || 'image/png'; return `data:${mime};base64,${await f.async('base64')}`; };
  const posOf = (el) => {
    let xx = lenToPx(A(el, 'svg', 'x')), yy = lenToPx(A(el, 'svg', 'y')); let rot = 0;
    const tr = A(el, 'draw', 'transform');
    if (tr) { const r = /rotate\s*\(([-\d.e]+)\)/.exec(tr); if (r) rot = Math.round((-Number(r[1]) * 180) / Math.PI); const tt = /translate\s*\(([-\d.]+[a-z]*)\s+([-\d.]+[a-z]*)\)/.exec(tr); if (tt) { xx = lenToPx(tt[1]); yy = lenToPx(tt[2]); } }
    return { x: Math.round(xx * k), y: Math.round(yy * k), w: Math.round(lenToPx(A(el, 'svg', 'width')) * k), h: Math.round(lenToPx(A(el, 'svg', 'height')) * k), rot: ((rot % 360) + 360) % 360 };
  };
  const slides = [];
  for (const page of doc.getElementsByTagNameNS(NS.draw, 'page')) {
    const pst = smap.get(A(page, 'draw', 'style-name')) || {};
    const s = { id: uid(), layout: 'blank', bg: pst.bg || null, notes: '', transition: { type: 'none', dur: 0.6 }, objects: [] };
    const evtLink = (el) => { const l = el.getElementsByTagNameNS(NS.presentation, 'event-listener')[0]; if (!l) return undefined; const act = A(l, 'presentation', 'action'); const href = A(l, 'xlink', 'href'); if (act === 'show' && href) { const m = /^#(.+)$/.exec(href); if (m) { const names = [...doc.getElementsByTagNameNS(NS.draw, 'page')].map((p) => A(p, 'draw', 'name')); const i = names.indexOf(m[1]); return i >= 0 ? `#slide:@${i}` : undefined; } return href; } return { 'next-page': '#next', 'previous-page': '#prev', 'first-page': '#first', 'last-page': '#last', stop: '#end' }[act]; };
    const styleOf = (st, extra = {}) => { const o = { ...extra }; for (const k of ['fill', 'stroke', 'strokeW', 'valign', 'grad', 'dash', 'strokeOpacity', 'fillOpacity', 'arrowStart', 'arrowEnd', 'shadow', 'pad', 'autofit', 'lineH', 'paraSp']) if (st[k] !== undefined) o[k] = st[k]; return o; };
    const walk = async (parent, group) => {
    for (const el of parent.children) {
      const ln = el.localName;
      if (ln === 'notes') { s.notes = [...el.getElementsByTagNameNS(NS.text, 'p')].map((p) => p.textContent).join('\n'); continue; }
      if (ln === 'g') { await walk(el, uid()); continue; }
      const st = { ...(smap.get(A(el, 'draw', 'style-name')) || {}) };
      const tbx = el.getElementsByTagNameNS(NS.draw, 'text-box')[0] || el;
      const ps = [...tbx.getElementsByTagNameNS(NS.text, 'p')][0]; if (ps) { const pst2 = smap.get(A(ps, 'text', 'style-name')) || {}; if (pst2.lineH) st.lineH = pst2.lineH; if (pst2.paraSp != null) st.paraSp = pst2.paraSp; }
      const before = s.objects.length;
      if (ln === 'frame') {
        const pos = posOf(el);
        const tb = el.getElementsByTagNameNS(NS.draw, 'text-box')[0], im = el.getElementsByTagNameNS(NS.draw, 'image')[0], tbl = el.getElementsByTagNameNS(NS.table, 'table')[0];
        const cls = A(el, 'presentation', 'class');
        if (tb) s.objects.push({ id: uid(), type: 'text', ...pos, role: cls === 'title' ? 'title' : cls === 'subtitle' ? 'subtitle' : cls === 'outline' ? 'body' : undefined, html: textHtml(tb), style: styleOf(st, { fill: st.fill || undefined, raw: true }) });
        else if (im) {
          const src = await imgUrl(A(im, 'xlink', 'href'));
          if (src) {
            const o = { id: uid(), type: 'image', ...pos, src };
            if (st.img) o.img = st.img; if (st.opacity != null) o.style = { opacity: st.opacity }; if (st.flipH) o.flipH = true; if (st.flipV) o.flipV = true;
            if (st.clip) { const [ct, cr, cb, cl] = st.clip; const nat = await natural(src); if (nat?.w && nat?.h) { const r = (v, T) => Math.round((v / T) * 1000) / 10; o.crop = { t: r(ct, nat.h), r: r(cr, nat.w), b: r(cb, nat.h), l: r(cl, nat.w) }; } }
            s.objects.push(o);
          }
        }
        else if (tbl) s.objects.push({ id: uid(), type: 'table', ...pos, rows: [...tbl.getElementsByTagNameNS(NS.table, 'table-row')].map((r) => [...r.children].map((c) => ({ text: c.textContent, hidden: c.localName === 'covered-table-cell', rs: 1, cs: Number(A(c, 'table', 'number-columns-spanned') || 1) }))), style: {} });
      } else if (ln === 'custom-shape' || ln === 'rect' || ln === 'ellipse') {
        const pos = posOf(el);
        const geo = el.getElementsByTagNameNS(NS.draw, 'enhanced-geometry')[0];
        const dt = A(geo || el, 'draw', 'type') || '';
        const type = ln === 'ellipse' ? 'ellipse' : ln === 'rect' ? 'rect' : SHAPE_IN[dt] || OOXML_IN[dt.replace(/^ooxml-/, '')] || 'rect';
        const o = { id: uid(), type: 'shape', shape: type, ...pos, html: textHtml(el), style: styleOf(st, { fill: st.fill === undefined ? '#729FCF' : st.fill, valign: 'middle', align: 'center', raw: true }) };
        if (geo && A(geo, 'draw', 'mirror-horizontal') === 'true') o.flipH = true; if (geo && A(geo, 'draw', 'mirror-vertical') === 'true') o.flipV = true;
        s.objects.push(o);
      } else if (ln === 'line') {
        const x1 = lenToPx(A(el, 'svg', 'x1')) * k, x2 = lenToPx(A(el, 'svg', 'x2')) * k, y1 = lenToPx(A(el, 'svg', 'y1')) * k, y2 = lenToPx(A(el, 'svg', 'y2')) * k;
        const len = Math.hypot(x2 - x1, y2 - y1); const ang = Math.round((Math.atan2(y2 - y1, x2 - x1) * 180) / Math.PI);
        const cx = (x1 + x2) / 2, cy = (y1 + y2) / 2;
        s.objects.push({ id: uid(), type: 'shape', shape: 'line', x: Math.round(cx - len / 2), y: Math.round(cy - 10), w: Math.round(len), h: 20, rot: ((ang % 360) + 360) % 360, html: '', style: styleOf(st, { stroke: st.stroke || '#000000', strokeW: st.strokeW || 2, fill: undefined }) });
      }
      if (s.objects.length > before) { const o = s.objects[s.objects.length - 1]; const l = evtLink(el); if (l) o.link = l; if (group) o.group = group; for (const key of Object.keys(o.style || {})) if (o.style[key] === undefined) delete o.style[key]; }
    }
    };
    await walk(page, null);
    if (pst.hidden) s.hidden = true;
    slides.push(s);
  }
  // vínculos a diapositivas: del número de página al id
  for (const sl of slides) for (const o of sl.objects) if (o.link?.startsWith('#slide:@')) { const t2 = slides[Number(o.link.slice(8))]; o.link = t2 ? `#slide:${t2.id}` : undefined; }
  // copia exacta de Impresiona si nadie cambió el contenido en otro programa
  const mf = zip.file('impresiona.json');
  if (mf) {
    try {
      const meta = JSON.parse(await mf.async('string'));
      if (meta?.pres?.slides?.length === slides.length && meta.hash === fnv(t)) {
        const mp = meta.pres; const mImgs = (mp.master?.objects || []).filter((o) => o.src === '@').length;
        mp.slides.forEach((ms, i) => {
          const srcs = slides[i].objects.filter((o) => o.type === 'image').map((o) => o.src);
          if (mp.master && i === 0) { let q = 0; for (const o of mp.master.objects) if (o.src === '@') o.src = srcs[q++] || ''; }
          let q = mImgs; for (const o of ms.objects) if (o.src === '@') o.src = srcs[q++] || '';
          ms.objects = ms.objects.filter((o) => o.src !== ''); if (ms.bgImage === '@') ms.bgImage = slides[i].bgImage || null;
        });
        const g2 = (n) => mdoc?.getElementsByTagNameNS(NS.dc, n)[0]?.textContent || '';
        return { pres: mp, props: { title: g2('title'), author: g2('creator') }, warnings: [] };
      }
    } catch (e) { console.warn('meta', e); }
  }
  const themeMeta = mdoc ? [...mdoc.getElementsByTagNameNS(NS.meta, 'user-defined')].find((u) => A(u, 'meta', 'name') === 'impresiona-tema')?.textContent : null;
  const g = (n) => mdoc?.getElementsByTagNameNS(NS.dc, n)[0]?.textContent || '';
  return { pres: { size: { w: 960, h: Math.round(H * k) }, theme: themeMeta || 'claro', slides }, props: { title: g('title'), author: g('creator') }, warnings: [] };
}
