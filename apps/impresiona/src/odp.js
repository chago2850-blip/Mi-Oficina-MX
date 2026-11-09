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
const NSDECL = Object.entries(NS).map(([k, v]) => `xmlns:${k}="${v}"`).join(' ');
const x = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const cm = (px) => `${(px / 37.7953).toFixed(3)}cm`;
const A = (el, ns, name) => el.getAttributeNS(NS[ns], name) || el.getAttribute(`${ns}:${name}`);
const lenToPx = (v) => { if (!v) return 0; const n = parseFloat(v); if (v.endsWith('cm')) return n * 37.7953; if (v.endsWith('mm')) return n * 3.77953; if (v.endsWith('in')) return n * 96; if (v.endsWith('pt')) return n * 4 / 3; return n; };
const rgb = (c) => { if (!c) return null; if (c.startsWith('#')) return c.slice(0, 7); const m = c.match(/rgba?\(([^)]+)\)/); if (!m) return null; const p = m[1].split(',').map(parseFloat); return `#${p.slice(0, 3).map((n) => Math.round(n).toString(16).padStart(2, '0')).join('')}`; };
const SHAPE_ODF = { rect: 'rectangle', roundRect: 'round-rectangle', ellipse: 'ellipse', triangle: 'isosceles-triangle', arrow: 'right-arrow', star: 'star5', line: 'rectangle' };
const SHAPE_IN = Object.fromEntries(Object.entries(SHAPE_ODF).map(([k, v]) => [v, k]));

export async function writeOdp(pres, props = {}) {
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.presentation', { compression: 'STORE' });
  const th = themeById(pres.theme);
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
  const paras = (html, base, align) => {
    const d = new DOMParser().parseFromString(`<body>${html || ''}</body>`, 'text/html');
    const pStyle = (el) => auto('paragraph', `<style:paragraph-properties fo:text-align="${{ center: 'center', right: 'end', justify: 'justify' }[el?.style?.textAlign || align] || 'start'}"/>`);
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
  const image = (src, name) => { const m = /^data:image\/([a-z+]+);base64,(.*)$/i.exec(src || ''); if (!m) return null; const ext = m[1].toLowerCase().replace('jpeg', 'jpg').replace('svg+xml', 'svg'); const p = `Pictures/${name || `img${++pic}`}.${ext}`; zip.file(p, m[2], { base64: true }); pics.push(p); return p; };
  const objXml = (o, idx) => {
    const st = o.style || {};
    const pos = `svg:x="${cm(o.x)}" svg:y="${cm(o.y)}" svg:width="${cm(o.w)}" svg:height="${cm(o.h)}"${o.rot ? ` draw:transform="rotate(${(-o.rot * Math.PI / 180).toFixed(5)}) translate(${cm(o.x)} ${cm(o.y)})"` : ''}`;
    const posNoRot = o.rot ? `svg:width="${cm(o.w)}" svg:height="${cm(o.h)}" draw:transform="rotate(${(-o.rot * Math.PI / 180).toFixed(5)}) translate(${cm(o.x)} ${cm(o.y)})"` : pos;
    const isTitle = o.role === 'title';
    const base = { color: st.color || (o.type === 'shape' ? '#FFFFFF' : isTitle ? th.title : th.text), size: st.size || (o.type === 'shape' ? 18 : 20), bold: st.bold, italic: st.italic, font: st.font || (isTitle ? th.titleFont : th.bodyFont) };
    if (o.type === 'text') {
      const gr = auto('graphic', `<style:graphic-properties draw:stroke="${st.stroke ? 'solid' : 'none'}"${st.stroke ? ` svg:stroke-color="${st.stroke}"` : ''} draw:fill="${st.fill ? 'solid' : 'none'}"${st.fill ? ` draw:fill-color="${st.fill}"` : ''} draw:textarea-vertical-align="${{ middle: 'middle', bottom: 'bottom' }[st.valign] || 'top'}" fo:padding="0.15cm"/>`);
      return `<draw:frame draw:style-name="${gr}" ${posNoRot}${o.role ? ` presentation:class="${o.role === 'body' ? 'outline' : o.role}"` : ''}><draw:text-box>${paras(o.html, base, st.align)}</draw:text-box></draw:frame>`;
    }
    if (o.type === 'shape') {
      const fill = st.fill === undefined ? th.accent : st.fill;
      const gr = auto('graphic', `<style:graphic-properties draw:stroke="${st.stroke ? 'solid' : 'none'}"${st.stroke ? ` svg:stroke-color="${st.stroke}" svg:stroke-width="${cm(st.strokeW || 2)}"` : ''} draw:fill="${fill ? 'solid' : 'none'}"${fill ? ` draw:fill-color="${fill}"` : ''} draw:textarea-vertical-align="middle"/>`);
      return `<draw:custom-shape draw:style-name="${gr}" ${posNoRot}>${paras(o.html, base, st.align || 'center')}<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="${SHAPE_ODF[o.shape] || 'rectangle'}"/></draw:custom-shape>`;
    }
    if (o.type === 'image') { const p = image(o.src); return p ? `<draw:frame ${posNoRot}><draw:image xlink:href="${p}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>` : ''; }
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
    const dp = auto('drawing-page', bgImg ? `<style:drawing-page-properties draw:fill="bitmap" draw:fill-image-name="Fondo${i + 1}"/>` : `<style:drawing-page-properties draw:fill="solid" draw:fill-color="${s.bg || th.bg}" presentation:transition-type="${s.transition?.type && s.transition.type !== 'none' ? 'automatic' : 'manual'}"/>`);
    const master = (pres.master?.objects || []).map((o, k) => objXml(o, `m${k}`)).join('');
    const objs = s.objects.map((o, k) => objXml(o, `${i}-${k}`)).join('');
    const notes = s.notes ? `<presentation:notes><draw:frame presentation:class="notes" svg:x="2cm" svg:y="14cm" svg:width="17cm" svg:height="10cm"><draw:text-box>${s.notes.split('\n').map((l) => `<text:p>${x(l)}</text:p>`).join('')}</draw:text-box></draw:frame></presentation:notes>` : '';
    pages += `<draw:page draw:name="Diapositiva${i + 1}" draw:style-name="${dp}" draw:master-page-name="Default">${master}${objs}${notes}</draw:page>`;
  });
  const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NSDECL} office:version="1.3"><office:automatic-styles>${autos.join('')}</office:automatic-styles><office:body><office:presentation>${pages}</office:presentation></office:body></office:document-content>`;
  const styles = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NSDECL} office:version="1.3"><office:styles><style:default-style style:family="graphic"><style:text-properties fo:font-size="18pt" fo:language="es" fo:country="MX"/></style:default-style></office:styles><office:automatic-styles><style:page-layout style:name="PM1"><style:page-layout-properties fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:page-width="${cm(pres.size.w)}" fo:page-height="${cm(pres.size.h)}" style:print-orientation="landscape"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles></office:document-styles>`;
  const meta = `<?xml version="1.0" encoding="UTF-8"?><office:document-meta ${NSDECL} office:version="1.3"><office:meta><meta:generator>Impresiona · Mi Oficina MX</meta:generator><dc:title>${x(props.title)}</dc:title><dc:creator>${x(props.author)}</dc:creator><meta:user-defined meta:name="impresiona-tema">${x(pres.theme)}</meta:user-defined></office:meta></office:document-meta>`;
  zip.file('content.xml', content); zip.file('styles.xml', styles); zip.file('meta.xml', meta);
  zip.file('META-INF/manifest.xml', `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>${pics.map((p) => `<manifest:file-entry manifest:full-path="${p}" manifest:media-type="image/${p.split('.').pop().replace('jpg', 'jpeg').replace('svg', 'svg+xml')}"/>`).join('')}</manifest:manifest>`);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.oasis.opendocument.presentation' });
}

export async function readOdp(bytes, uid) {
  const zip = await JSZip.loadAsync(bytes);
  const t = await zip.file('content.xml')?.async('string'); if (!t) throw new Error('No es una presentación ODF válida');
  const doc = new DOMParser().parseFromString(t, 'application/xml');
  const sdoc = zip.file('styles.xml') ? new DOMParser().parseFromString(await zip.file('styles.xml').async('string'), 'application/xml') : null;
  const mdoc = zip.file('meta.xml') ? new DOMParser().parseFromString(await zip.file('meta.xml').async('string'), 'application/xml') : null;
  const smap = new Map();
  for (const d of [sdoc, doc].filter(Boolean)) for (const st of d.getElementsByTagNameNS(NS.style, 'style')) {
    const o = {}; const name = A(st, 'style', 'name');
    const gp = st.getElementsByTagNameNS(NS.style, 'graphic-properties')[0], dp = st.getElementsByTagNameNS(NS.style, 'drawing-page-properties')[0], tp = st.getElementsByTagNameNS(NS.style, 'text-properties')[0], pp = st.getElementsByTagNameNS(NS.style, 'paragraph-properties')[0];
    if (gp) { if (A(gp, 'draw', 'fill') === 'solid') o.fill = A(gp, 'draw', 'fill-color'); if (A(gp, 'draw', 'fill') === 'none') o.fill = null; if (A(gp, 'draw', 'stroke') === 'solid') { o.stroke = A(gp, 'svg', 'stroke-color') || '#000000'; o.strokeW = lenToPx(A(gp, 'svg', 'stroke-width')) || 1; } const va = A(gp, 'draw', 'textarea-vertical-align'); if (va) o.valign = va; }
    if (dp && (A(dp, 'draw', 'fill') === 'solid' || A(dp, 'draw', 'fill-color'))) o.bg = A(dp, 'draw', 'fill-color');
    if (tp) { if (A(tp, 'fo', 'font-weight') === 'bold') o.bold = true; if (A(tp, 'fo', 'font-style') === 'italic') o.italic = true; o.color = A(tp, 'fo', 'color') || undefined; const sz = A(tp, 'fo', 'font-size'); if (sz?.endsWith('pt')) o.size = parseFloat(sz); const ul = A(tp, 'style', 'text-underline-style'); if (ul && ul !== 'none') o.underline = true; }
    if (pp) { const al = A(pp, 'fo', 'text-align'); if (al) o.align = { center: 'center', end: 'right', right: 'right', justify: 'justify' }[al] || 'left'; }
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
    for (const el of page.children) {
      const ln = el.localName;
      if (ln === 'notes') { s.notes = [...el.getElementsByTagNameNS(NS.text, 'p')].map((p) => p.textContent).join('\n'); continue; }
      const st = smap.get(A(el, 'draw', 'style-name')) || {};
      if (ln === 'frame') {
        const pos = posOf(el);
        const tb = el.getElementsByTagNameNS(NS.draw, 'text-box')[0], im = el.getElementsByTagNameNS(NS.draw, 'image')[0], tbl = el.getElementsByTagNameNS(NS.table, 'table')[0];
        const cls = A(el, 'presentation', 'class');
        if (tb) s.objects.push({ id: uid(), type: 'text', ...pos, role: cls === 'title' ? 'title' : cls === 'subtitle' ? 'subtitle' : cls === 'outline' ? 'body' : undefined, html: textHtml(tb), style: { fill: st.fill || undefined, stroke: st.stroke, strokeW: st.strokeW, valign: st.valign, raw: true } });
        else if (im) { const src = await imgUrl(A(im, 'xlink', 'href')); if (src) s.objects.push({ id: uid(), type: 'image', ...pos, src }); }
        else if (tbl) s.objects.push({ id: uid(), type: 'table', ...pos, rows: [...tbl.getElementsByTagNameNS(NS.table, 'table-row')].map((r) => [...r.children].map((c) => ({ text: c.textContent, hidden: c.localName === 'covered-table-cell', rs: 1, cs: Number(A(c, 'table', 'number-columns-spanned') || 1) }))), style: {} });
      } else if (ln === 'custom-shape' || ln === 'rect' || ln === 'ellipse') {
        const pos = posOf(el);
        const geo = el.getElementsByTagNameNS(NS.draw, 'enhanced-geometry')[0];
        const type = ln === 'ellipse' ? 'ellipse' : ln === 'rect' ? 'rect' : SHAPE_IN[A(geo || el, 'draw', 'type')] || 'rect';
        s.objects.push({ id: uid(), type: 'shape', shape: type, ...pos, html: textHtml(el), style: { fill: st.fill === undefined ? '#729FCF' : st.fill, stroke: st.stroke, strokeW: st.strokeW, valign: 'middle', align: 'center', raw: true } });
      } else if (ln === 'line') {
        const x1 = lenToPx(A(el, 'svg', 'x1')) * k, x2 = lenToPx(A(el, 'svg', 'x2')) * k, y1 = lenToPx(A(el, 'svg', 'y1')) * k;
        s.objects.push({ id: uid(), type: 'shape', shape: 'line', x: Math.min(x1, x2), y: y1 - 10, w: Math.abs(x2 - x1), h: 20, rot: 0, html: '', style: { stroke: st.stroke || '#000000', strokeW: st.strokeW || 2 } });
      }
    }
    slides.push(s);
  }
  const themeMeta = mdoc ? [...mdoc.getElementsByTagNameNS(NS.meta, 'user-defined')].find((u) => A(u, 'meta', 'name') === 'impresiona-tema')?.textContent : null;
  const g = (n) => mdoc?.getElementsByTagNameNS(NS.dc, n)[0]?.textContent || '';
  return { pres: { size: { w: 960, h: Math.round(H * k) }, theme: themeMeta || 'claro', slides }, props: { title: g('title'), author: g('creator') }, warnings: [] };
}
