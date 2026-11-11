// Documentos ODF de texto (.odt): abrir y guardar. GPL-3.0
import JSZip from 'jszip';
import { t } from '../../shared/i18n.js';
import { collectMeta, applyMeta, META_ODT } from './meta.js';

const NS = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0', style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0', table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0', fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  xlink: 'http://www.w3.org/1999/xlink', dc: 'http://purl.org/dc/elements/1.1/', meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
};
const NSDECL = Object.entries(NS).map(([k, v]) => `xmlns:${k}="${v}"`).join(' ');
const x = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const px2cm = (px) => `${(px / 37.7953).toFixed(3)}cm`;
export const roman = (n) => { let r = ''; for (const [v, s] of [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]) while (n >= v) { r += s; n -= v; } return r; };
export function lenToPx(v) {
  if (!v) return 0;
  const n = parseFloat(v);
  if (v.endsWith('cm')) return n * 37.7953; if (v.endsWith('mm')) return n * 3.77953; if (v.endsWith('in')) return n * 96;
  if (v.endsWith('pt')) return n * (4 / 3); if (v.endsWith('pc')) return n * 16; return n;
}
const rgb = (c) => { if (!c) return null; if (c.startsWith('#')) return c.length === 4 ? `#${[...c.slice(1)].map((y) => y + y).join('')}` : c.slice(0, 7); const m = c.match(/rgba?\(([^)]+)\)/); if (!m) return null; const p = m[1].split(',').map(parseFloat); if (p[3] === 0) return null; return `#${p.slice(0, 3).map((n) => Math.round(n).toString(16).padStart(2, '0')).join('')}`; };


// formato de texto a partir del estilo CSS de un elemento
function readFmt(st, g, el) {
  if (!st) return g;
  if (st.fontWeight === 'bold' || Number(st.fontWeight) >= 600) g.bold = true; else if (st.fontWeight === 'normal' || st.fontWeight === '400') g.bold = false;
  if (st.fontStyle === 'italic') g.italic = true; else if (st.fontStyle === 'normal') g.italic = false;
  const dec = st.textDecorationLine || st.textDecoration || ''; const dst = st.textDecorationStyle || (/(double|dotted|dashed|wavy)/.exec(st.textDecoration || '') || [])[1] || '';
  if (dec.includes('underline')) { g.underline = true; g.uStyle = dst; }
  if (dec.includes('line-through')) { if (dst === 'double') g.dstrike = true; else g.strike = true; }
  if (st.color && st.color !== 'transparent') g.color = rgb(st.color); if (st.backgroundColor) g.bg = rgb(st.backgroundColor);
  if (st.fontSize) { const v = st.fontSize; g.size = v.endsWith('pt') ? parseFloat(v) : v.endsWith('px') ? parseFloat(v) * 0.75 : g.size; }
  if (st.fontFamily) g.font = st.fontFamily.split(',')[0].replace(/["']/g, '').trim();
  if (st.fontVariant === 'small-caps' || st.fontVariantCaps === 'small-caps') g.smallCaps = true;
  if (st.textTransform === 'uppercase') g.caps = true;
  if (st.letterSpacing && st.letterSpacing !== 'normal') g.ls = lenToPx(st.letterSpacing) * 0.75;
  if (/(pt|px)$/.test(st.verticalAlign || '')) g.pos = lenToPx(st.verticalAlign) * 0.75;
  const sc = /scaleX\(([\d.]+)\)/.exec(st.transform || ''); if (sc) g.scale = Math.round(Number(sc[1]) * 100);
  if (st.textShadow && st.textShadow !== 'none') g.shadow = true;
  if (st.webkitTextStroke && parseFloat(st.webkitTextStrokeWidth || st.webkitTextStroke) > 0) { g.outline = true; const m = /(#[0-9a-f]{3,8}|rgba?\([^)]*\))\s*$/i.exec(st.webkitTextStroke); if (m) g.color = rgb(m[1]); }
  if (el?.classList?.contains('hidden-text')) g.hidden = true;
  return g;
}

// ================================================================ escribir
export async function htmlToOdt(editorEl, settings) {
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' });
  const autoStyles = new Map(); // clave -> nombre
  const styleXml = [];
  const pictures = [];
  let fnNum = 0, enNum = 0; let tableNo = 0;
  const pg = settings.page;
  const formulas = [];
  const textProps = (f) => {
    const props = [];
    if (f.bold) props.push('fo:font-weight="bold"'); else if (f.bold === false) props.push('fo:font-weight="normal"');
    if (f.italic) props.push('fo:font-style="italic"'); else if (f.italic === false) props.push('fo:font-style="normal"');
    if (f.underline) props.push(`style:text-underline-style="${{ dotted: 'dotted', dashed: 'dash', wavy: 'wave' }[f.uStyle] || 'solid'}" style:text-underline-width="auto" style:text-underline-color="font-color"${f.uStyle === 'double' ? ' style:text-underline-type="double"' : ''}`);
    if (f.strike || f.dstrike) props.push(`style:text-line-through-style="solid"${f.dstrike ? ' style:text-line-through-type="double"' : ''}`);
    if (f.color) props.push(`fo:color="${f.color}"`); if (f.bg) props.push(`fo:background-color="${f.bg}"`);
    if (f.size) props.push(`fo:font-size="${f.size}pt"`); if (f.font) props.push(`style:font-name="${x(f.font)}" fo:font-family="'${x(f.font)}'"`);
    if (f.sup) props.push('style:text-position="super 58%"'); else if (f.sub) props.push('style:text-position="sub 58%"'); else if (f.pos) props.push(`style:text-position="${Math.round((f.pos / (f.size || 11)) * 100)}% 100%"`);
    if (f.smallCaps) props.push('fo:font-variant="small-caps"');
    if (f.caps) props.push('fo:text-transform="uppercase"');
    if (f.ls) props.push(`fo:letter-spacing="${(f.ls).toFixed(2)}pt"`);
    if (f.shadow) props.push('fo:text-shadow="1pt 1pt"');
    if (f.outline) props.push('style:text-outline="true"');
    if (f.scale) props.push(`style:text-scale="${f.scale}%"`);
    if (f.hidden) props.push('text:display="none"');
    return props;
  };
  const textStyle = (f) => {
    const props = textProps(f);
    if (!props.length) return null;
    const key = `T|${props.join(' ')}`;
    if (!autoStyles.has(key)) { const name = `T${autoStyles.size + 1}`; autoStyles.set(key, name); styleXml.push(`<style:style style:name="${name}" style:family="text"><style:text-properties ${props.join(' ')}/></style:style>`); }
    return autoStyles.get(key);
  };
  const cm = (px) => px2cm(px);
  const borderOdf = (style, width, color) => (style && style !== 'none' && parseFloat(width) > 0 ? `${(lenToPx(width) * 0.75).toFixed(2)}pt ${{ double: 'double', dotted: 'dotted', dashed: 'dashed' }[style] || 'solid'} ${rgb(color) || '#000000'}` : null);
  const paraStyle = (el, parent) => {
    const st = el.style || {}; const ds = el.dataset || {};
    const props = [];
    const al = st.textAlign; if (al && al !== 'left' && al !== 'start') props.push(`fo:text-align="${al === 'justify' ? 'justify' : al === 'center' ? 'center' : 'end'}"`);
    const hasBL = st.borderLeftStyle && st.borderLeftStyle !== 'none';
    const ml = lenToPx(st.marginLeft) || (!hasBL ? lenToPx(st.paddingLeft) : 0); if (ml) props.push(`fo:margin-left="${cm(ml)}"`);
    const mr = lenToPx(st.marginRight); if (mr) props.push(`fo:margin-right="${cm(mr)}"`);
    const ti = lenToPx(st.textIndent); if (ti) props.push(`fo:text-indent="${cm(ti)}"`);
    const mt = ds.push ? ds.mt0 : st.marginTop; if (mt) props.push(`fo:margin-top="${cm(lenToPx(mt))}"`);
    if (st.marginBottom) props.push(`fo:margin-bottom="${cm(lenToPx(st.marginBottom))}"`);
    const lhs = String(st.lineHeight || '');
    if (/(pt|px)$/.test(lhs)) props.push(ds.lr === 'atLeast' ? `style:line-height-at-least="${(lenToPx(lhs) * 0.75).toFixed(2)}pt"` : `fo:line-height="${(lenToPx(lhs) * 0.75).toFixed(2)}pt"`);
    else if (parseFloat(lhs)) props.push(`fo:line-height="${Math.round(parseFloat(lhs) * 100)}%"`);
    for (const sd of ['top', 'right', 'bottom', 'left']) {
      const S = sd[0].toUpperCase() + sd.slice(1);
      const b = borderOdf(st[`border${S}Style`], st[`border${S}Width`], st[`border${S}Color`]);
      if (b) { props.push(`fo:border-${sd}="${b}"`); props.push(`fo:padding-${sd}="${(lenToPx(st[`padding${S}`]) * 0.75).toFixed(2)}pt"`); }
    }
    const bg = rgb(st.backgroundColor); if (bg) props.push(`fo:background-color="${bg}"`);
    if (ds.keepnext) props.push('fo:keep-with-next="always"');
    if (ds.keeplines) props.push('fo:keep-together="always"');
    if (ds.pbb) props.push('fo:break-before="page"');
    if (ds.widow === '0') props.push('fo:widows="0" fo:orphans="0"');
    const tabs = String(ds.tabs || '').split(',').map((t) => /^([lcrd]):(\d+(?:\.\d+)?)(?::(.))?/.exec(t)).filter(Boolean)
      .map((m) => `<style:tab-stop style:position="${px2cm(Number(m[2]) - ml)}"${m[1] === 'l' ? '' : m[1] === 'd' ? ' style:type="char" style:char=","' : ` style:type="${m[1] === 'c' ? 'center' : 'right'}"`}${m[3] ? ` style:leader-text="${x(m[3])}" style:leader-style="${m[3] === '_' ? 'solid' : m[3] === '-' ? 'dash' : 'dotted'}"` : ''}/>`).join('');
    // letra del propio párrafo
    const tprops = textProps(readFmt(st, {}, null));
    if (!props.length && !tabs && !tprops.length) return parent;
    const key = `P|${parent}|${props.join(' ')}|${tabs}|${tprops.join(' ')}`;
    if (!autoStyles.has(key)) { const name = `P${autoStyles.size + 1}`; autoStyles.set(key, name); styleXml.push(`<style:style style:name="${name}" style:family="paragraph" style:parent-style-name="${parent}"><style:paragraph-properties ${props.join(' ')}>${tabs ? `<style:tab-stops>${tabs}</style:tab-stops>` : ''}</style:paragraph-properties>${tprops.length ? `<style:text-properties ${tprops.join(' ')}/>` : ''}</style:style>`); }
    return autoStyles.get(key);
  };
  const inline = async (node, f = {}) => {
    let out = '';
    for (const n of node.childNodes) {
      if (n.nodeType === 3) {
        let t = n.nodeValue.replace(/​/g, ''); if (!t) continue;
        let body = x(t).replace(/\t/g, '<text:tab/>').replace(/ {2,}/g, (m) => ` <text:s text:c="${m.length - 1}"/>`);
        const sn = textStyle(f);
        out += sn ? `<text:span text:style-name="${sn}">${body}</text:span>` : body;
        continue;
      }
      if (n.nodeType !== 1) continue;
      const tag = n.tagName;
      if (tag === 'BR') { out += '<text:line-break/>'; continue; }
      if (tag === 'DEL') continue; // se guardan los cambios ya aceptados
      if (tag === 'IMG') { out += await image(n); continue; }
      if (n.classList?.contains('eq')) {
        // Ecuación: objeto de fórmula ODF incrustado (MathML)
        const m = n.querySelector('math'); if (!m) continue;
        const k = formulas.length + 1; const dir = `Object ${k}`;
        let xml = new XMLSerializer().serializeToString(m); if (!/xmlns=/.test(xml.slice(0, xml.indexOf('>')))) xml = xml.replace('<math', '<math xmlns="http://www.w3.org/1998/Math/MathML"');
        formulas.push(dir);
        zip.file(`${dir}/content.xml`, `<?xml version="1.0" encoding="UTF-8"?>${xml}`);
        const wcm = px2cm(Math.max(20, (n.offsetWidth || n.textContent.length * 9) + 6)), hcm = px2cm(Math.max(16, (n.offsetHeight || 22) + 4));
        out += `<draw:frame draw:name="Fórmula${k}" text:anchor-type="as-char" svg:width="${wcm}" svg:height="${hcm}"><draw:object xlink:href="./${dir}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
        continue;
      }
      if (tag === 'SUP' && n.classList.contains('en')) { const t = settings.endnotes?.[n.dataset.fid]; if (t != null) { enNum++; out += `<text:note text:id="edn${enNum}" text:note-class="endnote"><text:note-citation>${roman(enNum)}</text:note-citation><text:note-body><text:p text:style-name="Endnote">${x(t)}</text:p></text:note-body></text:note>`; } continue; }
      if (tag === 'SUP' && n.classList.contains('fn')) { const t = settings.footnotes?.[n.dataset.fid]; if (t != null) { fnNum++; out += `<text:note text:id="ftn${fnNum}" text:note-class="footnote"><text:note-citation>${fnNum}</text:note-citation><text:note-body><text:p text:style-name="Footnote">${x(t)}</text:p></text:note-body></text:note>`; } continue; }
      if (n.classList?.contains('mf')) { out += x(`«${n.dataset.f}»`); continue; }
      if (['UL', 'OL', 'TABLE'].includes(tag)) continue;
      const g = { ...f };
      if (tag === 'B' || tag === 'STRONG') g.bold = true; if (tag === 'I' || tag === 'EM') g.italic = true; if (tag === 'U') g.underline = true;
      if (tag === 'S' || tag === 'STRIKE') g.strike = true; if (tag === 'SUP') g.sup = true; if (tag === 'SUB') g.sub = true;
      readFmt(n.style, g, n);
      if (tag === 'A' && n.getAttribute('href')) { out += `<text:a xlink:type="simple" xlink:href="${x(n.getAttribute('href'))}">${await inline(n, g)}</text:a>`; continue; }
      if (n.classList?.contains('cm')) {
        const c = settings.comments?.find((cc) => cc.id === n.dataset.cid);
        if (c) out += `<office:annotation><dc:creator>${x(c.author)}</dc:creator><dc:date>${x(c.date)}</dc:date><text:p>${x(c.text)}</text:p>${(c.replies || []).map((r) => `<text:p>↪ ${x(r.author)}: ${x(r.text)}</text:p>`).join('')}</office:annotation>`;
      }
      out += await inline(n, g);
    }
    return out;
  };
  const image = async (img) => {
    const m = /^data:image\/([a-z+]+);base64,(.*)$/i.exec(img.getAttribute('src') || ''); if (!m) return '';
    let ext = m[1].toLowerCase().replace('jpeg', 'jpg').replace('svg+xml', 'svg');
    const name = `Pictures/img${pictures.length + 1}.${ext}`;
    pictures.push(name);
    zip.file(name, m[2], { base64: true });
    const w = img.width || img.naturalWidth || 300, hgt = img.height || Math.round(w * ((img.naturalHeight || 1) / (img.naturalWidth || 1)));
    return `<draw:frame draw:name="Imagen${pictures.length}" text:anchor-type="as-char" svg:width="${px2cm(w)}" svg:height="${px2cm(hgt)}"><draw:image xlink:href="${name}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
  };
  const list = async (el, depth = 0) => {
    const style = el.tagName === 'OL' ? 'LNum' : 'LBul';
    let out = `<text:list${depth === 0 ? ` text:style-name="${style}"` : ''}>`;
    for (const li of el.children) {
      if (li.tagName !== 'LI') continue;
      const clone = li.cloneNode(true); clone.querySelectorAll('ul,ol').forEach((y) => y.remove());
      const ps = [...clone.children].filter((c) => /^(P|H[1-6])$/.test(c.tagName));
      if (ps.length) { out += '<text:list-item>'; for (const pp of ps) out += `<text:p text:style-name="${paraStyle(pp, 'List_20_Paragraph')}">${await inline(pp)}</text:p>`; }
      else out += `<text:list-item><text:p text:style-name="${paraStyle(li, 'List_20_Paragraph')}">${await inline(clone)}</text:p>`;
      for (const sub of li.children) if (sub.tagName === 'UL' || sub.tagName === 'OL') out += await list(sub, depth + 1);
      out += '</text:list-item>';
    }
    return `${out}</text:list>`;
  };
  const block = async (el) => {
    const tag = el.tagName;
    if (!/^(P|H[1-6]|DIV|UL|OL|TABLE|BLOCKQUOTE|HR|LI|PRE)$/.test(tag)) return `<text:p text:style-name="Standard">${await inline({ childNodes: [el] })}</text:p>`;
    if (el.classList?.contains('pb')) return '<text:p text:style-name="PBreak"/>';
    if (el.classList?.contains('cols') || el.classList?.contains('tbox') || (tag === 'DIV' && el.querySelector(':scope > p, :scope > h1, :scope > ul, :scope > table'))) { let o = ''; for (const c of el.children) o += await block(c); return o; }
    if (el.classList?.contains('toc')) { let o = ''; for (const l of el.querySelectorAll('.toc-h, .toc-l')) o += `<text:p text:style-name="${l.classList.contains('toc-h') ? 'Heading_20_1' : 'Standard'}">${x(l.classList.contains('toc-h') ? l.textContent : `${l.children[0]?.textContent || ''}\t${l.children[2]?.textContent || ''}`).replace(/\t/g, '<text:tab/>')}</text:p>`; return o; }
    if (tag === 'UL' || tag === 'OL') return list(el);
    if (tag === 'HR') return '<text:p text:style-name="HRule"/>';
    if (tag === 'TABLE') {
      const rows = [...el.querySelectorAll(':scope > tbody > tr, :scope > thead > tr, :scope > tr')];
      // cuadrícula con celdas combinadas
      const occ = []; const placed = [];
      rows.forEach((tr, ri) => { occ[ri] ||= []; let ci = 0; for (const td of tr.children) { while (occ[ri][ci]) ci++; const cs = td.colSpan || 1; const rs = td.rowSpan || 1; placed.push({ td, ri, ci, cs, rs }); for (let r = ri; r < ri + rs; r++) { occ[r] ||= []; for (let c = ci; c < ci + cs; c++) occ[r][c] = r === ri && c === ci ? 'own' : 'cov'; } ci += cs; } });
      const cols = Math.max(1, ...occ.map((r) => r.length));
      const tblW = el.offsetWidth || (pg.w - pg.ml - pg.mr);
      let widths = [...el.querySelectorAll(':scope > colgroup > col')].map((c) => (/%$/.test(c.style.width) ? (parseFloat(c.style.width) / 100) * tblW : lenToPx(c.style.width) || 0));
      if (widths.length !== cols || widths.some((w) => !w)) { widths = Array(cols).fill(0); for (const pc of placed) if (pc.cs === 1 && !widths[pc.ci]) widths[pc.ci] = pc.td.offsetWidth || 0; const known = widths.filter(Boolean); const avg = known.length ? known.reduce((a, b) => a + b, 0) / known.length : (pg.w - pg.ml - pg.mr) / cols; widths = widths.map((w) => w || avg); }
      const tn = `Tabla${++tableNo}`;
      const tw = widths.reduce((a, b) => a + b, 0);
      const ts = el.style; const center = ts.marginLeft === 'auto' && ts.marginRight === 'auto'; const right = ts.marginLeft === 'auto' && !center;
      styleXml.push(`<style:style style:name="${tn}" style:family="table"><style:table-properties style:width="${px2cm(tw)}" table:align="${center ? 'center' : right ? 'right' : 'left'}"${!center && !right && lenToPx(ts.marginLeft) ? ` fo:margin-left="${px2cm(lenToPx(ts.marginLeft))}"` : ''}/></style:style>`);
      let o = `<table:table table:name="${tn}" table:style-name="${tn}">`;
      widths.forEach((w, i) => { styleXml.push(`<style:style style:name="${tn}.C${i + 1}" style:family="table-column"><style:table-column-properties style:column-width="${px2cm(w)}"/></style:style>`); o += `<table:table-column table:style-name="${tn}.C${i + 1}"/>`; });
      const live = el.isConnected;
      const cellStyle = (td) => {
        const cs = live ? getComputedStyle(td) : td.style; const pr = [];
        for (const sd of ['top', 'right', 'bottom', 'left']) { const S = sd[0].toUpperCase() + sd.slice(1); const b = borderOdf(cs[`border${S}Style`], cs[`border${S}Width`], cs[`border${S}Color`]); pr.push(`fo:border-${sd}="${b || 'none'}"`); pr.push(`fo:padding-${sd}="${(lenToPx(cs[`padding${S}`]) * 0.75).toFixed(2)}pt"`); }
        const bg = rgb(cs.backgroundColor); if (bg) pr.push(`fo:background-color="${bg}"`);
        if (cs.verticalAlign === 'middle' || cs.verticalAlign === 'bottom') pr.push(`style:vertical-align="${cs.verticalAlign}"`);
        const key = `C|${pr.join(' ')}`;
        if (!autoStyles.has(key)) { const name = `Ce${autoStyles.size + 1}`; autoStyles.set(key, name); styleXml.push(`<style:style style:name="${name}" style:family="table-cell"><style:table-cell-properties ${pr.join(' ')}/></style:style>`); }
        return autoStyles.get(key);
      };
      for (let ri = 0; ri < rows.length; ri++) {
        o += '<table:table-row>';
        for (let ci = 0; ci < cols; ci++) {
          const pc = placed.find((p) => p.ri === ri && p.ci === ci);
          if (!pc) { o += occ[ri]?.[ci] === 'cov' ? '<table:covered-table-cell/>' : '<table:table-cell><text:p/></table:table-cell>'; continue; }
          const { td, cs, rs } = pc;
          const kids = [...td.children].filter((c) => /^(P|H[1-6]|UL|OL|TABLE|BLOCKQUOTE|DIV)$/.test(c.tagName));
          const inner = kids.length ? (await Promise.all(kids.map(block))).join('') : `<text:p>${await inline(td, td.tagName === 'TH' ? { bold: true } : {})}</text:p>`;
          o += `<table:table-cell table:style-name="${cellStyle(td)}" office:value-type="string"${cs > 1 ? ` table:number-columns-spanned="${cs}"` : ''}${rs > 1 ? ` table:number-rows-spanned="${rs}"` : ''}>${inner}</table:table-cell>`;
        }
        o += '</table:table-row>';
      }
      return `${o}</table:table>`;
    }
    if (/^H[1-6]$/.test(tag)) { const lvl = Number(tag[1]); return `<text:h text:style-name="${paraStyle(el, `Heading_20_${lvl}`)}" text:outline-level="${lvl}">${await inline(el)}</text:h>`; }
    const base = el.classList?.contains('title') ? 'Title' : el.classList?.contains('subtitle') ? 'Subtitle' : tag === 'BLOCKQUOTE' ? 'Quotations' : 'Standard';
    return `<text:p text:style-name="${paraStyle(el, base)}">${await inline(el)}</text:p>`;
  };
  let body = '';
  for (const el of editorEl.children) body += await block(el);
  const content = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${NSDECL} office:version="1.3"><office:automatic-styles>
<style:style style:name="PBreak" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:break-after="page"/></style:style>
<style:style style:name="HRule" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:border-bottom="0.5pt solid #999999"/></style:style>
<style:style style:name="Tbl" style:family="table"><style:table-properties style:width="${px2cm(pg.w - pg.ml - pg.mr)}" table:align="margins"/></style:style>
<style:style style:name="Cell" style:family="table-cell"><style:table-cell-properties fo:border="0.5pt solid #9A9EAD" fo:padding="0.1cm"/></style:style>
<style:style style:name="CellH" style:family="table-cell"><style:table-cell-properties fo:border="0.5pt solid #9A9EAD" fo:padding="0.1cm" fo:background-color="#EEECFB"/></style:style>
<text:list-style style:name="LBul">${[...Array(6)].map((_, i) => `<text:list-level-style-bullet text:level="${i + 1}" text:bullet-char="${['•', '◦', '▪'][i % 3]}"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" fo:text-indent="-0.635cm" fo:margin-left="${(0.635 * (i + 1) + 0.6).toFixed(2)}cm"/></style:list-level-properties></text:list-level-style-bullet>`).join('')}</text:list-style>
<text:list-style style:name="LNum">${[...Array(6)].map((_, i) => `<text:list-level-style-number text:level="${i + 1}" style:num-suffix="." style:num-format="${['1', 'a', 'i'][i % 3]}"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" fo:text-indent="-0.635cm" fo:margin-left="${(0.635 * (i + 1) + 0.6).toFixed(2)}cm"/></style:list-level-properties></text:list-level-style-number>`).join('')}</text:list-style>
${styleXml.join('\n')}
</office:automatic-styles><office:body><office:text>${body || '<text:p/>'}</office:text></office:body></office:document-content>`;
  // {pagina} y {paginas} → campos de ODF; cada línea, un párrafo
  const hfOdt = (text, tail = '') => { const lines = String(text || '').split('\n'); return lines.map((l, i) => `<text:p>${x(l).replace(/\{pagina\}/gi, '<text:page-number text:select-page="current"/>').replace(/\{paginas\}/gi, '<text:page-count/>')}${i === lines.length - 1 ? tail : ''}</text:p>`).join(''); };
  const heading = (n, size, color) => `<style:style style:name="Heading_20_${n}" style:display-name="Heading ${n}" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Standard" style:default-outline-level="${n}"><style:paragraph-properties fo:margin-top="0.3cm" fo:margin-bottom="0.15cm" fo:keep-with-next="always"/><style:text-properties fo:font-size="${size}pt" fo:font-weight="bold" fo:color="${color}"/></style:style>`;
  const styles = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${NSDECL} office:version="1.3"><office:font-face-decls><style:font-face style:name="Calibri" svg:font-family="Calibri, Carlito"/></office:font-face-decls>
<office:styles>
<style:default-style style:family="paragraph"><style:paragraph-properties fo:margin-bottom="0.28cm" fo:line-height="115%" style:tab-stop-distance="${px2cm(settings.defTab || 47.24)}"/><style:text-properties style:font-name="Calibri" fo:font-size="11pt" fo:language="es" fo:country="MX"/></style:default-style>
<style:style style:name="Standard" style:family="paragraph"/>
${heading(1, 16, '#2F2496')}${heading(2, 13, '#3D30B5')}${heading(3, 12, '#3D30B5')}${heading(4, 11, '#3D30B5')}${heading(5, 11, '#3D30B5')}${heading(6, 11, '#3D30B5')}
<style:style style:name="Title" style:family="paragraph" style:parent-style-name="Standard"><style:text-properties fo:font-size="28pt"/></style:style>
<style:style style:name="Subtitle" style:family="paragraph" style:parent-style-name="Standard"><style:text-properties fo:font-size="14pt" fo:color="#6B6F80"/></style:style>
<style:style style:name="Quotations" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-left="1cm" fo:margin-right="1cm"/><style:text-properties fo:font-style="italic" fo:color="#444444"/></style:style>
<style:style style:name="List_20_Paragraph" style:display-name="List Paragraph" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-bottom="0.05cm"/></style:style>
<style:style style:name="Footnote" style:family="paragraph" style:parent-style-name="Standard"><style:text-properties fo:font-size="9pt"/></style:style>
<style:style style:name="Endnote" style:family="paragraph" style:parent-style-name="Standard"><style:text-properties fo:font-size="9pt"/></style:style>
<text:notes-configuration text:note-class="endnote" style:num-format="i" text:start-value="0"/>
</office:styles>
<office:automatic-styles><style:page-layout style:name="PL1"><style:page-layout-properties fo:page-width="${px2cm(pg.w)}" fo:page-height="${px2cm(pg.h)}" style:print-orientation="${pg.landscape ? 'landscape' : 'portrait'}" fo:margin-top="${px2cm(pg.mt)}" fo:margin-bottom="${px2cm(pg.mb)}" fo:margin-left="${px2cm(pg.ml)}" fo:margin-right="${px2cm(pg.mr)}"${settings.pageColor ? ` fo:background-color="${settings.pageColor}"` : ''}/></style:page-layout></office:automatic-styles>
<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="PL1">${settings.header ? `<style:header>${hfOdt(settings.header)}</style:header>` : ''}${settings.footer || (settings.pageNumbers && settings.pageNumbers !== 'none') ? `<style:footer>${hfOdt(settings.footer || '', settings.pageNumbers && settings.pageNumbers !== 'none' ? ` ${t('Página')} <text:page-number text:select-page="current"/>` : '')}</style:footer>` : ''}</style:master-page></office:master-styles></office:document-styles>`;
  const meta = `<?xml version="1.0" encoding="UTF-8"?><office:document-meta ${NSDECL} office:version="1.3"><office:meta><meta:generator>Escribiente · Mi Oficina MX</meta:generator><dc:title>${x(settings.title)}</dc:title><dc:creator>${x(settings.author)}</dc:creator><dc:subject>${x(settings.subject)}</dc:subject><meta:keyword>${x(settings.keywords)}</meta:keyword><dc:date>${new Date().toISOString()}</dc:date></office:meta></office:document-meta>`;
  zip.file('content.xml', content); zip.file('styles.xml', styles); zip.file('meta.xml', meta);
  zip.file('META-INF/manifest.xml', `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>${pictures.map((p) => `<manifest:file-entry manifest:full-path="${p}" manifest:media-type="image/${p.split('.').pop().replace('jpg', 'jpeg').replace('svg', 'svg+xml')}"/>`).join('')}${formulas.map((d) => `<manifest:file-entry manifest:full-path="${d}/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.formula"/><manifest:file-entry manifest:full-path="${d}/content.xml" manifest:media-type="text/xml"/>`).join('')}</manifest:manifest>`);
  const xmeta = collectMeta(editorEl, settings);
  if (xmeta) {
    zip.file(META_ODT, JSON.stringify(xmeta));
    const mf = await zip.file('META-INF/manifest.xml').async('string');
    zip.file('META-INF/manifest.xml', mf.replace('</manifest:manifest>', `<manifest:file-entry manifest:full-path="${META_ODT}" manifest:media-type="application/json"/></manifest:manifest>`));
  }
  return zip.generateAsync({ type: 'uint8array', mimeType: 'application/vnd.oasis.opendocument.text', compression: 'DEFLATE' });
}

// ================================================================ leer
const A = (el, ns, name) => el.getAttributeNS(NS[ns], name) || el.getAttribute(`${ns}:${name}`);
function collectStyles(doc, map) {
  for (const st of doc.getElementsByTagNameNS(NS.style, 'style')) {
    const name = A(st, 'style', 'name'); if (!name) continue;
    const tp = st.getElementsByTagNameNS(NS.style, 'text-properties')[0];
    const pp = st.getElementsByTagNameNS(NS.style, 'paragraph-properties')[0];
    const o = { parent: A(st, 'style', 'parent-style-name'), display: A(st, 'style', 'display-name') || name, family: A(st, 'style', 'family') };
    if (tp) {
      if (A(tp, 'fo', 'font-weight') === 'bold' || Number(A(tp, 'fo', 'font-weight')) >= 600) o.bold = true;
      if (A(tp, 'fo', 'font-style') === 'italic') o.italic = true;
      const ul = A(tp, 'style', 'text-underline-style'); if (ul && ul !== 'none') o.underline = true;
      const lt = A(tp, 'style', 'text-line-through-style'); if (lt && lt !== 'none') o.strike = true;
      o.color = A(tp, 'fo', 'color') || undefined; const bg = A(tp, 'fo', 'background-color'); if (bg && bg !== 'transparent') o.bg = bg;
      const sz = A(tp, 'fo', 'font-size'); if (sz && sz.endsWith('pt')) o.size = parseFloat(sz);
      o.font = A(tp, 'style', 'font-name') || A(tp, 'fo', 'font-family') || undefined;
      const pos = A(tp, 'style', 'text-position'); if (pos?.startsWith('super')) o.sup = true; if (pos?.startsWith('sub')) o.sub = true;
    }
    if (pp) { const al = A(pp, 'fo', 'text-align'); if (al) o.align = { center: 'center', end: 'right', right: 'right', justify: 'justify' }[al]; const ml = A(pp, 'fo', 'margin-left'); if (ml) o.ml = lenToPx(ml); const bb = A(pp, 'fo', 'break-before'); if (bb === 'page') o.breakBefore = true; const ba = A(pp, 'fo', 'break-after'); if (ba === 'page') o.breakAfter = true; }
    map.set(name, o);
  }
}
export async function odtToHtml(bytes) {
  // lector con formato completo; si falla, el lector sencillo de abajo
  try { const { readOdt } = await import('./odt-read.js'); return await readOdt(bytes); } catch (e) { console.warn('lector odt', e); }
  const zip = await JSZip.loadAsync(bytes);
  const parse = async (f) => { const t = await zip.file(f)?.async('string'); return t ? new DOMParser().parseFromString(t, 'application/xml') : null; };
  const content = await parse('content.xml'); if (!content) throw new Error('No es un documento ODF válido');
  const styles = await parse('styles.xml'); const meta = await parse('meta.xml');
  const smap = new Map(); if (styles) collectStyles(styles, smap); collectStyles(content, smap);
  const resolve = (name) => { const out = {}; const chain = []; let n = name, guard = 0; while (n && smap.has(n) && guard++ < 12) { chain.unshift(smap.get(n)); n = smap.get(n).parent; } chain.forEach((c) => Object.assign(out, Object.fromEntries(Object.entries(c).filter(([, v]) => v !== undefined)))); out.names = chain.map((c) => c.display); return out; };
  const listKinds = new Map();
  for (const ls of content.getElementsByTagNameNS(NS.text, 'list-style')) listKinds.set(A(ls, 'style', 'name'), ls.getElementsByTagNameNS(NS.text, 'list-level-style-number').length && !ls.firstElementChild?.localName?.includes('bullet') ? 'ol' : 'ul');
  if (styles) for (const ls of styles.getElementsByTagNameNS(NS.text, 'list-style')) if (!listKinds.has(A(ls, 'style', 'name'))) listKinds.set(A(ls, 'style', 'name'), ls.firstElementChild?.localName === 'list-level-style-number' ? 'ol' : 'ul');
  const footnotes = {}, endnotes = {};
  const imgs = {};
  const e = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const css = (o) => [o.bold && 'font-weight:bold', o.italic && 'font-style:italic', (o.underline || o.strike) && `text-decoration:${[o.underline && 'underline', o.strike && 'line-through'].filter(Boolean).join(' ')}`, o.color && `color:${o.color}`, o.bg && `background-color:${o.bg}`, o.size && `font-size:${o.size}pt`, o.font && `font-family:${o.font}`].filter(Boolean).join(';');
  const inline = async (node) => {
    let out = '';
    for (const n of node.childNodes) {
      if (n.nodeType === 3) { out += e(n.nodeValue); continue; }
      if (n.nodeType !== 1) continue;
      const ln = n.localName;
      if (ln === 's') { out += ' '.repeat(Number(A(n, 'text', 'c') || 1)); continue; }
      if (ln === 'tab') { out += '\t'; continue; }
      if (ln === 'line-break') { out += '<br>'; continue; }
      if (ln === 'span') { const st = resolve(A(n, 'text', 'style-name')); const c = css(st); let inner = await inline(n); if (st.sup) inner = `<sup>${inner}</sup>`; if (st.sub) inner = `<sub>${inner}</sub>`; out += c ? `<span style="${c}">${inner}</span>` : inner; continue; }
      if (ln === 'a') { out += `<a href="${e(A(n, 'xlink', 'href') || '')}">${await inline(n)}</a>`; continue; }
      if (ln === 'note') { const body = n.getElementsByTagNameNS(NS.text, 'note-body')[0]; const end = A(n, 'text', 'note-class') === 'endnote'; const fid = `${end ? 'e' : 'f'}${Math.random().toString(36).slice(2, 8)}`; (end ? endnotes : footnotes)[fid] = body?.textContent.trim() || ''; out += `<sup class="${end ? 'en' : 'fn'}" data-fid="${fid}" contenteditable="false">#</sup>`; continue; }
      if (ln === 'frame') {
        const im = n.getElementsByTagNameNS(NS.draw, 'image')[0]; const href = im && A(im, 'xlink', 'href');
        if (href && zip.file(href)) {
          if (!imgs[href]) { const ext = href.split('.').pop().toLowerCase(); const mime = { jpg: 'image/jpeg', jpeg: 'image/jpeg', svg: 'image/svg+xml', gif: 'image/gif' }[ext] || 'image/png'; imgs[href] = `data:${mime};base64,${await zip.file(href).async('base64')}`; }
          const w = Math.round(lenToPx(A(n, 'svg', 'width')) || 300); out += `<img src="${imgs[href]}" width="${w}" alt="">`;
        } else if (n.getElementsByTagNameNS(NS.draw, 'object')[0]) {
          // Fórmula incrustada -> ecuación editable
          const href = A(n.getElementsByTagNameNS(NS.draw, 'object')[0], 'xlink', 'href').replace(/^\.\//, '');
          const src = await zip.file(`${href}/content.xml`)?.async('string');
          const md = src && new DOMParser().parseFromString(src, 'application/xml');
          const math = md?.documentElement?.localName === 'math' ? md.documentElement : md?.getElementsByTagNameNS('http://www.w3.org/1998/Math/MathML', 'math')[0];
          if (math) { const tex = [...math.getElementsByTagName('*')].find((a) => a.localName === 'annotation')?.textContent || ''; out += `<span class="eq" contenteditable="false" data-tex="${e(tex)}">${new XMLSerializer().serializeToString(math).replace(/ xmlns(:\w+)?="[^"]*"/g, '').replace(/<(\/?)\w+:/g, '<$1')}</span>`; }
        } else { const tb = n.getElementsByTagNameNS(NS.draw, 'text-box')[0]; if (tb) for (const p of tb.children) out += await inline(p); }
        continue;
      }
      if (ln === 'annotation' || ln === 'annotation-end' || ln === 'bookmark' || ln === 'bookmark-start' || ln === 'bookmark-end' || ln === 'soft-page-break' || ln === 'tracked-changes') continue;
      out += await inline(n);
    }
    return out;
  };
  const pStyle = (st, el) => { const parts = []; if (st.align) parts.push(`text-align:${st.align}`); if (st.ml) parts.push(`margin-left:${Math.round(st.ml)}px`); const c = css({ ...st, size: undefined }); void el; return parts.join(';') + (c && (st.bold || st.italic || st.color) ? '' : ''); };
  const blocks = async (parent) => {
    let out = '';
    for (const n of parent.children) {
      const ln = n.localName;
      if (ln === 'h') { const lvl = Math.min(6, Number(A(n, 'text', 'outline-level') || 1)); const st = resolve(A(n, 'text', 'style-name')); out += `<h${lvl}${st.align ? ` style="text-align:${st.align}"` : ''}>${await inline(n) || '<br>'}</h${lvl}>`; continue; }
      if (ln === 'p') {
        const st = resolve(A(n, 'text', 'style-name'));
        const names = (st.names || []).join(' ').toLowerCase();
        if (st.breakBefore && out) out += '<div class="pb" contenteditable="false"></div>';
        const inner = (await inline(n)) || '<br>';
        const style = pStyle(st, n); const wrapCss = css({ bold: st.bold, italic: st.italic, color: st.color, size: st.size && st.size !== 11 && !/title/.test(names) ? st.size : undefined });
        const body = wrapCss ? `<span style="${wrapCss}">${inner}</span>` : inner;
        if (/^title| title/.test(names) && !/sub/.test(names)) out += `<p class="title"${style ? ` style="${style}"` : ''}>${inner}</p>`;
        else if (/subtitle/.test(names)) out += `<p class="subtitle"${style ? ` style="${style}"` : ''}>${inner}</p>`;
        else if (/quotation|quote|cita/.test(names)) out += `<blockquote>${inner}</blockquote>`;
        else out += `<p${style ? ` style="${style}"` : ''}>${body}</p>`;
        if (st.breakAfter) out += '<div class="pb" contenteditable="false"></div>';
        continue;
      }
      if (ln === 'list') {
        const kind = listKinds.get(A(n, 'text', 'style-name')) || 'ul';
        const walk = async (l) => { let o = `<${kind}>`; for (const it of l.children) { if (it.localName !== 'list-item' && it.localName !== 'list-header') continue; let li = ''; for (const c of it.children) { if (c.localName === 'list') li += await walk(c); else li += await inline(c); } o += `<li>${li || '<br>'}</li>`; } return `${o}</${kind}>`; };
        out += await walk(n); continue;
      }
      if (ln === 'table') {
        let o = '<table><tbody>';
        const rows = [...n.getElementsByTagNameNS(NS.table, 'table-row')];
        for (const r of rows) {
          o += '<tr>';
          for (const c of r.children) { if (c.localName !== 'table-cell') continue; const span = Number(A(c, 'table', 'number-columns-spanned') || 1); const ps = [...c.children]; const txt = (await Promise.all(ps.map((p) => inline(p)))).join('<br>'); o += `<td${span > 1 ? ` colspan="${span}"` : ''}>${txt || '<br>'}</td>`; }
          o += '</tr>';
        }
        out += `${o}</tbody></table>`; continue;
      }
      if (ln === 'section' || ln === 'index-body' || ln === 'table-of-content' || ln === 'alphabetical-index') { out += await blocks(ln === 'table-of-content' ? (n.getElementsByTagNameNS(NS.text, 'index-body')[0] || n) : n); continue; }
    }
    return out;
  };
  const body = content.getElementsByTagNameNS(NS.office, 'text')[0];
  const html = body ? await blocks(body) : '';
  let page = null;
  const pl = styles?.getElementsByTagNameNS(NS.style, 'page-layout-properties')[0];
  if (pl && A(pl, 'fo', 'page-width')) {
    const w = Math.round(lenToPx(A(pl, 'fo', 'page-width'))), hgt = Math.round(lenToPx(A(pl, 'fo', 'page-height')));
    page = { w, h: hgt, mt: lenToPx(A(pl, 'fo', 'margin-top') || '2cm'), mb: lenToPx(A(pl, 'fo', 'margin-bottom') || '2cm'), ml: lenToPx(A(pl, 'fo', 'margin-left') || '2cm'), mr: lenToPx(A(pl, 'fo', 'margin-right') || '2cm'), landscape: w > hgt };
  }
  const mg = (tag, ns = 'dc') => meta?.getElementsByTagNameNS(NS[ns], tag)[0]?.textContent || '';
  let html2 = html; let sources = [];
  try {
    const mt = await zip.file(META_ODT)?.async('string');
    if (mt) { const box = document.createElement('div'); box.innerHTML = html; ({ sources } = applyMeta(box, JSON.parse(mt), [], null)); html2 = box.innerHTML; }
  } catch { /* */ }
  return { html: html2 || '<p><br></p>', page, footnotes, endnotes, sources, props: { title: mg('title'), author: mg('creator') || mg('initial-creator', 'meta'), subject: mg('subject'), keywords: mg('keyword', 'meta') } };
}
