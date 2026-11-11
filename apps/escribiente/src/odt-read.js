// Lector de documentos ODF (.odt) con su formato: estilos con herencia, propiedades de párrafo y de
// texto (casi idénticas a CSS), listas con su numeración, tablas con celdas combinadas, bordes y
// fondos, imágenes, cuadros de texto, notas, comentarios, encabezado y pie, secciones. GPL-3.0
import JSZip from 'jszip';
import { applyMeta, META_ODT } from './meta.js';
import { sanitize } from './docx-in.js';

const NS = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0', style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0', table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0', fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  xlink: 'http://www.w3.org/1999/xlink', dc: 'http://purl.org/dc/elements/1.1/', meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
};
const A = (el, ns, name) => (el ? el.getAttributeNS(NS[ns], name) ?? el.getAttribute(`${ns}:${name}`) : null);
const kids = (el, ln) => (el ? [...el.children].filter((c) => !ln || c.localName === ln) : []);
const kid = (el, ln) => kids(el, ln)[0] || null;
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const FONT_ALT = { Calibri: 'Carlito', Cambria: 'Caladea', Arial: "'Liberation Sans'", 'Times New Roman': "'Liberation Serif'", 'Courier New': "'Liberation Mono'" };
const fontCss = (f) => { const n = String(f).replace(/['"]/g, '').split(',')[0].trim(); const alt = FONT_ALT[n]; return `'${n}'${alt ? `, ${alt}` : ''}`; };
export function lenToPx(v) {
  if (!v) return 0;
  const n = parseFloat(v);
  if (v.endsWith('cm')) return n * 37.7953; if (v.endsWith('mm')) return n * 3.77953; if (v.endsWith('in')) return n * 96;
  if (v.endsWith('pt')) return n * (4 / 3); if (v.endsWith('pc')) return n * 16; return n;
}
const ptOf = (v) => `${+(lenToPx(v) * 0.75).toFixed(2)}pt`;

// propiedades de un estilo (párrafo y texto) en forma CSS
function readProps(st) {
  const p = {}; const t = {}; const flags = {};
  const pp = kid(st, 'paragraph-properties'); const tp = kid(st, 'text-properties');
  if (pp) {
    const al = A(pp, 'fo', 'text-align'); if (al) p['text-align'] = { start: 'left', left: 'left', end: 'right', right: 'right', center: 'center', justify: 'justify' }[al] || 'left';
    for (const k of ['margin-top', 'margin-bottom', 'margin-left', 'margin-right', 'text-indent', 'padding-top', 'padding-bottom', 'padding-left', 'padding-right']) { const v = A(pp, 'fo', k); if (v != null) p[k] = /%$/.test(v) ? v : ptOf(v); }
    const pad = A(pp, 'fo', 'padding'); if (pad) for (const s of ['top', 'bottom', 'left', 'right']) if (p[`padding-${s}`] == null) p[`padding-${s}`] = ptOf(pad);
    const lh = A(pp, 'fo', 'line-height'); if (lh) p['line-height'] = /%$/.test(lh) ? String(+(parseFloat(lh) / 100).toFixed(3)) : lh === 'normal' ? '1' : ptOf(lh);
    const lal = A(pp, 'style', 'line-height-at-least'); if (lal) { p['line-height'] = ptOf(lal); flags.lr = 'atLeast'; }
    const bd = A(pp, 'fo', 'border'); const sides = ['top', 'bottom', 'left', 'right'];
    for (const s of sides) { const v = A(pp, 'fo', `border-${s}`) || bd; if (v && v !== 'none') p[`border-${s}`] = v.replace(/(\d*\.?\d+)(cm|mm|in)/, (m) => ptOf(m)); else if (v === 'none') p[`border-${s}`] = 'none'; }
    const bg = A(pp, 'fo', 'background-color'); if (bg && bg !== 'transparent') p['background-color'] = bg;
    if (A(pp, 'fo', 'keep-with-next') === 'always') flags.keepnext = true;
    if (A(pp, 'fo', 'keep-together') === 'always') flags.keeplines = true;
    if (A(pp, 'fo', 'break-before') === 'page') flags.pbb = true;
    if (A(pp, 'fo', 'break-after') === 'page') flags.pba = true;
    if (A(pp, 'fo', 'widows') === '0' || A(pp, 'fo', 'orphans') === '0') flags.widow = '0';
    const tabs = kid(pp, 'tab-stops');
    // en ODF la posición se mide desde la sangría izquierda del párrafo
    const mlPx = lenToPx(A(pp, 'fo', 'margin-left') || '0');
    if (tabs) flags.tabs = kids(tabs, 'tab-stop').map((ts) => { const lt = A(ts, 'style', 'leader-text'); const ls = A(ts, 'style', 'leader-style'); const f = lt && lt !== ' ' ? lt : ls && ls !== 'none' ? (ls === 'solid' ? '_' : ls === 'dash' ? '-' : '.') : ''; return `${{ center: 'c', right: 'r', char: 'd' }[A(ts, 'style', 'type')] || 'l'}:${Math.round((lenToPx(A(ts, 'style', 'position')) + mlPx) * 10) / 10}${f ? `:${f}` : ''}`; }).join(',');
  }
  if (tp) {
    const fn = A(tp, 'style', 'font-name') || A(tp, 'fo', 'font-family'); if (fn) t['font-family'] = fontCss(fn);
    const fs = A(tp, 'fo', 'font-size'); if (fs) t['font-size'] = /%$/.test(fs) ? fs : ptOf(fs);
    const c = A(tp, 'fo', 'color'); if (c) t.color = c;
    const fw = A(tp, 'fo', 'font-weight'); if (fw) t['font-weight'] = fw === 'bold' || Number(fw) >= 600 ? 'bold' : 'normal';
    const fst = A(tp, 'fo', 'font-style'); if (fst) t['font-style'] = fst === 'italic' || fst === 'oblique' ? 'italic' : 'normal';
    const ul = A(tp, 'style', 'text-underline-style'); const ult = A(tp, 'style', 'text-underline-type');
    const lt = A(tp, 'style', 'text-line-through-style'); const ltt = A(tp, 'style', 'text-line-through-type');
    const lines = []; let dstyle = '';
    if (ul && ul !== 'none') { lines.push('underline'); dstyle = ult === 'double' ? 'double' : { dotted: 'dotted', dash: 'dashed', 'long-dash': 'dashed', wave: 'wavy' }[ul] || ''; }
    if (lt && lt !== 'none') { lines.push('line-through'); if (ltt === 'double') dstyle = 'double'; }
    if (ul === 'none' || lt === 'none') t['text-decoration-line'] = lines.join(' ') || 'none';
    if (lines.length) { t['text-decoration-line'] = lines.join(' '); if (dstyle) t['text-decoration-style'] = dstyle; }
    const bg = A(tp, 'fo', 'background-color'); if (bg && bg !== 'transparent') t['background-color'] = bg;
    if (A(tp, 'fo', 'font-variant') === 'small-caps') t['font-variant'] = 'small-caps';
    const tt = A(tp, 'fo', 'text-transform'); if (tt) t['text-transform'] = tt;
    const ls = A(tp, 'fo', 'letter-spacing'); if (ls && ls !== 'normal') t['letter-spacing'] = ptOf(ls);
    const sh = A(tp, 'fo', 'text-shadow'); if (sh && sh !== 'none') t['text-shadow'] = '1px 1px 1px rgba(0,0,0,.45)';
    if (A(tp, 'style', 'text-outline') === 'true') { t['-webkit-text-stroke'] = '.6px currentColor'; flags.outline = true; }
    const pos = A(tp, 'style', 'text-position'); if (pos) { if (/^super/.test(pos)) flags.sup = true; else if (/^sub/.test(pos)) flags.sub = true; else if (parseFloat(pos)) t['vertical-align'] = `${parseFloat(pos) / 10}em`; }
    const sc = A(tp, 'style', 'text-scale'); if (sc && parseFloat(sc) !== 100) { t.display = 'inline-block'; t.transform = `scaleX(${parseFloat(sc) / 100})`; t['transform-origin'] = 'left'; }
    if (A(tp, 'text', 'display') === 'none') flags.hidden = true;
  }
  return { p, t, flags };
}

export async function readOdt(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const parse = async (f) => { const t = await zip.file(f)?.async('string'); return t ? new DOMParser().parseFromString(t, 'application/xml') : null; };
  const content = await parse('content.xml'); if (!content) throw new Error('No es un documento ODF válido');
  const stylesDoc = await parse('styles.xml'); const meta = await parse('meta.xml');
  // estilos: nombre → {family, parent, props}
  const styles = new Map(); const defaults = {};
  const listStyles = new Map();
  const collect = (doc) => {
    if (!doc) return;
    for (const d of doc.getElementsByTagNameNS(NS.style, 'default-style')) defaults[A(d, 'style', 'family')] = readProps(d);
    for (const s of doc.getElementsByTagNameNS(NS.style, 'style')) {
      styles.set(`${A(s, 'style', 'family')}:${A(s, 'style', 'name')}`, { name: A(s, 'style', 'name'), display: A(s, 'style', 'display-name') || A(s, 'style', 'name'), family: A(s, 'style', 'family'), parent: A(s, 'style', 'parent-style-name'), listStyle: A(s, 'style', 'list-style-name'), props: readProps(s), el: s, outline: A(s, 'style', 'default-outline-level'), auto: s.parentNode?.localName === 'automatic-styles', next: A(s, 'style', 'next-style-name') });
    }
    for (const ls of doc.getElementsByTagNameNS(NS.text, 'list-style')) {
      const levels = {};
      for (const l of ls.children) {
        const lv = Number(A(l, 'text', 'level')) || 1;
        const lp = kid(l, 'list-level-properties'); const la = lp && kid(lp, 'list-level-label-alignment');
        const ind = lenToPx(A(la, 'fo', 'margin-left') || A(lp, 'text', 'space-before') || '0') || 24 * lv;
        levels[lv] = l.localName === 'list-level-style-number' ? { num: true, fmt: A(l, 'style', 'num-format') || '1', start: Number(A(l, 'text', 'start-value') || 1), ind, display: Number(A(l, 'text', 'display-levels') || 1), suffix: A(l, 'style', 'num-suffix'), prefix: A(l, 'style', 'num-prefix') } : { num: false, char: A(l, 'text', 'bullet-char') || '•', ind };
      }
      listStyles.set(A(ls, 'style', 'name'), levels);
    }
  };
  collect(stylesDoc); collect(content);
  // intervalo de las paradas predeterminadas
  let defTab = null;
  for (const d of stylesDoc?.getElementsByTagNameNS(NS.style, 'default-style') || []) { if (A(d, 'style', 'family') !== 'paragraph') continue; const v = A(kid(d, 'paragraph-properties'), 'style', 'tab-stop-distance'); if (v) defTab = Math.round(lenToPx(v) * 100) / 100; }
  const chain = (family, name) => { const out = []; let n = name; let g = 0; while (n && g++ < 20) { const s = styles.get(`${family}:${n}`); if (!s) break; out.unshift(s); n = s.parent; } return out; };
  const resolve = (family, name) => {
    const ch = chain(family, name); const d = defaults[family] || { p: {}, t: {}, flags: {} };
    const r = { p: { ...d.p }, t: { ...d.t }, flags: { ...d.flags }, names: ch.map((s) => s.display.toLowerCase()), listStyle: null, outline: null };
    for (const s of ch) { Object.assign(r.p, s.props.p); Object.assign(r.t, s.props.t); Object.assign(r.flags, s.props.flags); if (s.listStyle) r.listStyle = s.listStyle; if (s.outline) r.outline = s.outline; }
    return r;
  };
  // estilos con nombre del documento (no los de LibreOffice): se conservan como estilos de Escribiente
  const LO_BUILTIN = /^(Standard|Text_20_|Default|Heading|Title|Subtitle|Quotations|List|Table|Header|Footer|Footnote|Endnote|Caption|Illustration|Index|Contents|Frame|Preformatted|Signature|Salutation|Horizontal|Hanging|First|Marginalia|Addressee|Sender|Drawing|Figure|Text$|Numbering|Appendix|Bibliography|User_20_Index|Internet|Visited|Placeholder|ListLabel|WW8Num|WWNum|Outline|Main_20_index|Page_20_Number|Line_20_numbering|Example|Teletype|User_20_Entry|Variable|Definition|Citation|Rubies|Vertical|Drop_20_Caps|Bullet_20_Symbols|Numbering_20_Symbols|Endnote_20_|Footnote_20_)/;
  const CHAR_NAMES = { emphasis: 'Emphasis', 'énfasis': 'Emphasis', 'strong emphasis': 'Strong', 'énfasis fuerte': 'Strong', strong: 'Strong', 'source text': 'SourceText', 'texto de código': 'SourceText', quotation: 'Quotation', 'cita textual': 'Quotation' };
  const CHAR_IN = new Proxy({}, { get: (_, k) => (typeof k === 'string' ? CHAR_NAMES[(styles.get(`text:${k}`)?.display || k.replace(/_20_/g, ' ')).toLowerCase()] : undefined) });
  const usedStyles = new Map();
  const PARENT_IN = { Standard: 'Normal', Text_20_body: 'Normal', Title: 'Title', Subtitle: 'Subtitle', Quotations: 'Quote', ...Object.fromEntries([1, 2, 3, 4, 5, 6].map((n) => [`Heading_20_${n}`, `Heading${n}`])) };
  const namedOf = (family, name) => { const ch = chain(family, name); for (let i = ch.length - 1; i >= 0; i--) if (!ch[i].auto) return ch[i]; return null; };
  const ownCss = (s) => Object.fromEntries(Object.entries({ ...s.props.p, ...s.props.t }).filter(([, v]) => v != null && v !== '').map(([k, v]) => [k, String(v).replace(/"/g, "'")]));
  const register = (s, kind) => {
    const id = kind === 'c' && CHAR_IN[s.name] ? CHAR_IN[s.name] : s.name;
    if (usedStyles.has(id)) return id;
    if (kind === 'c' && CHAR_IN[s.name]) { usedStyles.set(id, { id, name: id, kind: 'c', css: ownCss(s) }); return id; }
    const par = s.parent && styles.get(`paragraph:${s.parent}`);
    let parent = PARENT_IN[s.parent] || 'Normal';
    if (kind === 'p' && par && !LO_BUILTIN.test(par.name) && !par.auto) parent = register(par, 'p');
    usedStyles.set(id, { id, name: s.display, kind, parent, next: s.next && !LO_BUILTIN.test(s.next) ? s.next : id, css: ownCss(s) });
    return id;
  };
  const diffOf = (full, base) => { const o = {}; for (const [k, v] of Object.entries(full)) if (base[k] !== v) o[k] = v; return o; };
  const cssOf = (o) => Object.entries(o).filter(([, v]) => v != null && v !== '').map(([k, v]) => `${k}:${String(v).replace(/"/g, "'")}`).join(';');
  const footnotes = {}; const endnotes = {}; const comments = []; const imgs = {};
  const warn = new Set();
  const media = async (href) => {
    if (!href || !zip.file(href)) return null;
    if (!imgs[href]) { const ext = href.split('.').pop().toLowerCase(); const mime = { jpg: 'image/jpeg', jpeg: 'image/jpeg', svg: 'image/svg+xml', gif: 'image/gif', png: 'image/png', bmp: 'image/bmp', webp: 'image/webp' }[ext]; if (!mime) { warn.add(`imagen ${ext}`); return null; } imgs[href] = `data:${mime};base64,${await zip.file(href).async('base64')}`; }
    return imgs[href];
  };
  let openComments = [];

  // texto en línea; base = estilo de texto del párrafo (para no repetirlo)
  const inline = async (node, base) => {
    let out = '';
    for (const n of node.childNodes) {
      if (n.nodeType === 3) { out += esc(n.nodeValue.replace(/\s+/g, ' ')); continue; }
      if (n.nodeType !== 1) continue;
      const ln = n.localName;
      switch (ln) {
        case 's': out += '&nbsp;'.repeat(Number(A(n, 'text', 'c') || 1)); break;
        case 'tab': out += '\t'; break;
        case 'line-break': out += '<br>'; break;
        case 'soft-page-break': case 'bookmark-end': case 'reference-mark': case 'reference-mark-start': case 'reference-mark-end': case 'tracked-changes': case 'change': case 'change-start': case 'change-end': break;
        case 'bookmark': case 'bookmark-start': { const nm = A(n, 'text', 'name'); if (nm) out += `<a id="${esc(nm)}" class="bm"></a>`; break; }
        case 'span': {
          const st = resolve('text', A(n, 'text', 'style-name'));
          const named = namedOf('text', A(n, 'text', 'style-name'));
          const cid = named && (CHAR_IN[named.name] || !LO_BUILTIN.test(named.name)) ? register(named, 'c') : null;
          const nbase = cid ? { ...base, ...resolve('text', named.name).t } : base;
          const diff = {}; for (const [k, v] of Object.entries(st.t)) if (nbase[k] !== v) diff[k] = v;
          let inner = await inline(n, { ...base, ...st.t });
          if (st.flags.sup) inner = `<sup>${inner}</sup>`; if (st.flags.sub) inner = `<sub>${inner}</sub>`;
          if (st.flags.outline && diff.color) diff['-webkit-text-stroke'] = `.6px ${diff.color}`;
          const c = cssOf(diff);
          out += c || st.flags.hidden || cid ? `<span${st.flags.hidden ? ' class="hidden-text"' : ''}${cid ? ` data-cstyle="${esc(cid)}"` : ''}${c ? ` style="${c}"` : ''}>${inner}</span>` : inner;
          break;
        }
        case 'a': out += `<a href="${esc(A(n, 'xlink', 'href') || '')}">${await inline(n, base)}</a>`; break;
        case 'note': {
          const body = n.getElementsByTagNameNS(NS.text, 'note-body')[0]; const end = A(n, 'text', 'note-class') === 'endnote';
          const fid = `${end ? 'e' : 'f'}${Math.random().toString(36).slice(2, 8)}`;
          (end ? endnotes : footnotes)[fid] = [...(body?.children || [])].map((p) => p.textContent).join('\n').trim();
          out += `<sup class="${end ? 'en' : 'fn'}" data-fid="${fid}" contenteditable="false">#</sup>`;
          break;
        }
        case 'annotation': {
          const author = n.getElementsByTagNameNS(NS.dc, 'creator')[0]?.textContent || 'Autor';
          const date = n.getElementsByTagNameNS(NS.dc, 'date')[0]?.textContent || new Date().toISOString();
          const lines = kids(n, 'p').concat(kids(n, 'list')).map((p) => p.textContent);
          const replies = []; const main = [];
          for (const l of lines) { const m = /^↪\s*([^:]*):\s?(.*)$/.exec(l); if (m) replies.push({ author: m[1].trim(), date, text: m[2] }); else main.push(l); }
          const id = `c${Math.random().toString(36).slice(2, 9)}`;
          comments.push({ id, author, date, text: main.join('\n').trim(), replies, resolved: false });
          const name = A(n, 'office', 'name');
          if (name) { openComments.push({ name, id }); out += `\uE002${id}\uE003`; } else out += `<span class="cm" data-cid="${id}">​</span>`;
          break;
        }
        case 'annotation-end': { const name = A(n, 'office', 'name'); const i = openComments.findIndex((c) => c.name === name); if (i >= 0) { const { id } = openComments[i]; openComments.splice(i, 1); out = wrapComment(out, id); } break; }
        case 'frame': out += await frame(n); break;
        case 'sequence': case 'page-number': case 'page-count': case 'date': case 'time': case 'title': case 'chapter': case 'author-name': case 'file-name': case 'bibliography-mark': case 'variable-get': case 'user-defined': case 'meta': case 'meta-field': case 'ruby': case 'ruby-base': out += await inline(n, base); break;
        default: out += await inline(n, base);
      }
    }
    return out;
  };
  // envuelve en span.cm el texto desde la marca de inicio del comentario
  const markers = new Map();
  const wrapComment = (html, id) => { const k = html.lastIndexOf(`${id}`); if (k < 0) return `${html}<span class="cm" data-cid="${id}">​</span>`; return `${html.slice(0, k)}<span class="cm" data-cid="${id}">${html.slice(k + id.length + 2)}</span>`; };
  void markers;
  const frame = async (n) => {
    const w = Math.round(lenToPx(A(n, 'svg', 'width')) || 0); const hh = Math.round(lenToPx(A(n, 'svg', 'height')) || 0);
    const im = n.getElementsByTagNameNS(NS.draw, 'image')[0];
    const obj = n.getElementsByTagNameNS(NS.draw, 'object')[0];
    const tb = n.getElementsByTagNameNS(NS.draw, 'text-box')[0];
    if (obj) {
      const href = (A(obj, 'xlink', 'href') || '').replace(/^\.\//, '');
      const src = await zip.file(`${href}/content.xml`)?.async('string');
      const md = src && new DOMParser().parseFromString(src, 'application/xml');
      const math = md?.documentElement?.localName === 'math' ? md.documentElement : md?.getElementsByTagNameNS('http://www.w3.org/1998/Math/MathML', 'math')[0];
      if (math) { const tex = [...math.getElementsByTagName('*')].find((a) => a.localName === 'annotation')?.textContent || ''; return `<span class="eq" contenteditable="false" data-tex="${esc(tex)}">${new XMLSerializer().serializeToString(math).replace(/ xmlns(:\w+)?="[^"]*"/g, '')}</span>`; }
      if (im) { const src2 = await media(A(im, 'xlink', 'href')); if (src2) return `<img src="${src2}"${w ? ` width="${w}"` : ''} alt="">`; }
      warn.add('objetos incrustados'); return '';
    }
    if (tb) { let inner = ''; for (const c of tb.children) inner += await block(c); return `<div class="tbox"${w ? ` style="width:${w}px"` : ''}>${inner || '<p><br></p>'}</div>`; }
    if (im) {
      const src = await media(A(im, 'xlink', 'href')); if (!src) return '';
      const fst = resolve('graphic', A(n, 'draw', 'style-name'));
      const gs = styles.get(`graphic:${A(n, 'draw', 'style-name')}`)?.el; const gp = gs && kid(gs, 'graphic-properties');
      const wrap = A(gp, 'style', 'wrap'); const hp = A(gp, 'style', 'horizontal-pos');
      let style = '';
      if (A(n, 'text', 'anchor-type') !== 'as-char' && (wrap === 'right' || wrap === 'parallel' || wrap === 'dynamic' || wrap === 'left')) style = hp === 'right' ? 'float:right;margin:0 0 6pt 9pt' : 'float:left;margin:0 9pt 6pt 0';
      void fst;
      return `<img src="${src}"${w ? ` width="${w}"` : ''}${hh ? ` height="${hh}"` : ''} alt="${esc(n.getElementsByTagNameNS(NS.svg, 'desc')[0]?.textContent || '')}"${style ? ` style="${style}"` : ''}>`;
    }
    return '';
  };

  // numeración de listas: contador por estilo de lista y nivel
  const counters = new Map();
  let chapFromList = null;
  const para = async (n, { inList = false } = {}) => {
    const ln = n.localName;
    const st = resolve('paragraph', A(n, 'text', 'style-name'));
    const names = st.names.join('|');
    let tag = 'p'; let cls = '';
    const hn = /(?:^|\|)(?:heading|título|titulo|encabezado) (\d)(?:\||$)/.exec(names);
    if (ln === 'h') tag = `h${Math.min(6, Number(A(n, 'text', 'outline-level') || st.outline || 1))}`;
    else if (hn) tag = `h${Math.min(6, Number(hn[1]))}`;
    else if (/(^|\|)(title|título)(\||$)/.test(names)) cls = 'title';
    else if (/(^|\|)(subtitle|subtítulo)(\||$)/.test(names)) cls = 'subtitle';
    else if (/(^|\|)(quotations|quote|cita|citas)(\||$)/.test(names)) tag = 'blockquote';
    const pcss = { ...st.p }; if (inList) { delete pcss['margin-left']; delete pcss['text-indent']; }
    // valores normales explícitos para que el CSS del editor no los cambie
    if (pcss['margin-top'] == null) pcss['margin-top'] = '0pt'; if (pcss['margin-bottom'] == null) pcss['margin-bottom'] = '0pt';
    const named = namedOf('paragraph', A(n, 'text', 'style-name'));
    const custom = named && !LO_BUILTIN.test(named.name) && tag === 'p' && !cls ? register(named, 'p') : null;
    let full = { ...pcss, ...st.t };
    if (custom) { const ns = resolve('paragraph', named.name); full = diffOf(full, { ...ns.p, ...ns.t }); }
    const css = cssOf(full);
    const data = custom ? [`data-style="${esc(custom)}"`] : [];
    for (const k of ['keepnext', 'keeplines', 'pbb']) if (st.flags[k]) data.push(`data-${k}="1"`);
    if (st.flags.widow) data.push('data-widow="0"'); if (st.flags.lr) data.push(`data-lr="${st.flags.lr}"`); if (st.flags.tabs) data.push(`data-tabs="${st.flags.tabs}"`);
    let inner = await inline(n, st.t);
    // marcas de inicio de comentario abiertas en este párrafo
    let boxes = '';
    inner = inner.replace(/([\s\S]*?)/g, (m, b) => { boxes += b; return ''; });
    if (!inner.replace(/<a id="[^"]*" class="bm"><\/a>/g, '').trim()) inner += '<br>';
    let html = `<${tag}${cls ? ` class="${cls}"` : ''}${css ? ` style="${css}"` : ''}${data.length ? ` ${data.join(' ')}` : ''}>${inner}</${tag}>`;
    if (st.flags.pba) html += '<div class="pb" contenteditable="false"></div>';
    return { html: html + boxes, pbb: st.flags.pbb, listStyle: st.listStyle };
  };
  const listHtml = async (l, styleName, level) => {
    const sname = A(l, 'text', 'style-name') || styleName;
    const ls = listStyles.get(sname) || {};
    const lv = ls[level] || ls[1] || { num: false, ind: 24 * level };
    const tag = lv.num ? 'ol' : 'ul';
    const type = lv.num ? ({ 1: 'decimal', a: 'lower-alpha', A: 'upper-alpha', i: 'lower-roman', I: 'upper-roman' }[lv.fmt] || 'decimal') : /[○◦o]/.test(lv.char) ? 'circle' : /[■▪]/.test(lv.char) ? 'square' : 'disc';
    const key = `${sname}:${level}`;
    const cont = A(l, 'text', 'continue-numbering') === 'true' || A(l, 'text', 'continue-list');
    if (!cont || !counters.has(key)) counters.set(key, (lv.start || 1) - 1);
    const startAt = counters.get(key) + 1;
    const prevInd = level > 1 ? (ls[level - 1]?.ind || 24 * (level - 1)) : 0;
    const legal = level === 1 && tag === 'ol' && Object.values(ls).some((x) => x.display > 1);
    const BUL = { '◦': 'b-circle', '▪': 'b-square', '✓': 'b-check', '➢': 'b-arrow', '–': 'b-dash' };
    const ml = legal ? 'm-legal' : level === 1 && tag === 'ul' && BUL[lv.char] ? BUL[lv.char] : level === 1 && tag === 'ol' && lv.suffix === ')' ? 'n-1p' : '';
    let o = `<${tag}${ml ? ` data-ml="${ml}"` : ''} style="list-style-type:${type};padding-left:${+(Math.max(16, lv.ind - prevInd) * 0.75).toFixed(1)}pt;margin-top:0;margin-bottom:0"${tag === 'ol' && startAt !== 1 ? ` start="${startAt}"` : ''}>`;
    for (const it of l.children) {
      if (it.localName !== 'list-item' && it.localName !== 'list-header') continue;
      if (it.localName === 'list-item') { const sv = A(it, 'text', 'start-value'); if (sv) counters.set(key, Number(sv) - 1); counters.set(key, counters.get(key) + 1); }
      let li = '';
      for (const c of it.children) {
        if (c.localName === 'list') li += await listHtml(c, sname, level + 1);
        else if (c.localName === 'p' || c.localName === 'h') li += (await para(c, { inList: true })).html;
        else li += await block(c);
      }
      o += `<li${it.localName === 'list-header' ? ' style="list-style:none"' : ''}>${li || '<p><br></p>'}</li>`;
    }
    return `${o}</${tag}>`;
  };
  const table = async (n) => {
    const tst = styles.get(`table:${A(n, 'table', 'style-name')}`)?.el; const tpp = tst && kid(tst, 'table-properties');
    const cols = [];
    const colEls = [...kids(n, 'table-column'), ...kids(n, 'table-columns').flatMap((g) => kids(g, 'table-column')), ...kids(n, 'table-header-columns').flatMap((g) => kids(g, 'table-column'))];
    for (const c of colEls) { const cs = styles.get(`table-column:${A(c, 'table', 'style-name')}`)?.el; const w = lenToPx(A(kid(cs, 'table-column-properties'), 'style', 'column-width')); const rep = Number(A(c, 'table', 'number-columns-repeated') || 1); for (let i = 0; i < Math.min(rep, 64); i++) cols.push(w); }
    const tw = lenToPx(A(tpp, 'style', 'width'));
    const al = A(tpp, 'table', 'align');
    const tcss = ['border-collapse:collapse'];
    if (tw) tcss.push(`width:min(${Math.round(tw)}px, 100%)`);
    if (al === 'center') tcss.push('margin-left:auto;margin-right:auto'); else if (al === 'right') tcss.push('margin-left:auto;margin-right:0');
    else if (A(tpp, 'fo', 'margin-left') && al === 'left') tcss.push(`margin-left:${Math.round(lenToPx(A(tpp, 'fo', 'margin-left')))}px`);
    let o = `<table style="${tcss.join(';')}">`;
    const cw = cols.reduce((x, y) => x + y, 0);
    if (cols.length && cols.every(Boolean)) o += `<colgroup>${cols.map((w) => `<col style="width:${((w / cw) * 100).toFixed(2)}%">`).join('')}</colgroup>`;
    o += '<tbody>';
    const rows = [];
    const walkRows = (el) => { for (const c of el.children) { if (c.localName === 'table-row') rows.push(c); else if (/^table-(header-rows|rows|row-group)$/.test(c.localName)) walkRows(c); } };
    walkRows(n);
    for (const r of rows) {
      const header = r.parentElement?.localName === 'table-header-rows';
      const rsEl = styles.get(`table-row:${A(r, 'table', 'style-name')}`)?.el; const rpp = rsEl && kid(rsEl, 'table-row-properties');
      const mrh = rpp && (A(rpp, 'style', 'min-row-height') || A(rpp, 'style', 'row-height')); const keepR = rpp && A(rpp, 'fo', 'keep-together') === 'always';
      o += `<tr${keepR ? ' data-keep="1"' : ''}${mrh ? ` style="height:${(lenToPx(mrh) * 0.75).toFixed(2)}pt"` : ''}>`;
      for (const c of r.children) {
        if (c.localName !== 'table-cell') continue;
        const cs = styles.get(`table-cell:${A(c, 'table', 'style-name')}`)?.el; const cp = cs && kid(cs, 'table-cell-properties');
        const css = [];
        const bd = A(cp, 'fo', 'border');
        for (const s of ['top', 'right', 'bottom', 'left']) { const v = A(cp, 'fo', `border-${s}`) || bd; css.push(`border-${s}:${v && v !== 'none' ? v.replace(/(\d*\.?\d+)(cm|mm|in)/, (m) => ptOf(m)) : 'none'}`); }
        const bg = A(cp, 'fo', 'background-color'); css.push(`background-color:${bg && bg !== 'transparent' ? bg : 'transparent'}`);
        if (header) css.push('font-weight:inherit', 'text-align:inherit');
        const pad = A(cp, 'fo', 'padding'); if (pad) css.push(`padding:${ptOf(pad)}`);
        for (const s of ['top', 'bottom', 'left', 'right']) { const v = A(cp, 'fo', `padding-${s}`); if (v) css.push(`padding-${s}:${ptOf(v)}`); }
        const va = A(cp, 'style', 'vertical-align'); if (va) css.push(`vertical-align:${va === 'middle' ? 'middle' : va === 'bottom' ? 'bottom' : 'top'}`);
        const cs2 = Number(A(c, 'table', 'number-columns-spanned') || 1); const rs = Number(A(c, 'table', 'number-rows-spanned') || 1);
        let inner = ''; for (const k of c.children) inner += await block(k);
        const tg = header ? 'th' : 'td';
        o += `<${tg}${cs2 > 1 ? ` colspan="${cs2}"` : ''}${rs > 1 ? ` rowspan="${rs}"` : ''} style="${css.join(';')}">${inner || '<p><br></p>'}</${tg}>`;
      }
      o += '</tr>';
    }
    return `${o}</tbody></table>`;
  };
  const block = async (n) => {
    const ln = n.localName;
    if (ln === 'p' || ln === 'h') { const r = await para(n); return `${r.pbb ? '' : ''}${r.html}`; }
    if (ln === 'list') {
      // lista que sólo tiene títulos (así guarda LibreOffice la numeración de capítulos de Word): numeración de capítulos
      const items = [...n.getElementsByTagNameNS(NS.text, 'list-item')];
      const hs = items.flatMap((it) => [...it.children].filter((c) => c.localName === 'h' || c.localName === 'p'));
      if (hs.length && hs.every((c) => c.localName === 'h')) {
        const lv = (listStyles.get(A(n, 'text', 'style-name')) || {})[1];
        if (lv?.num && !chapFromList) chapFromList = lv.fmt === 'I' ? 'IA1' : lv.prefix?.trim() ? 'cap' : (lv.suffix || '').includes('.') ? '1.' : '1';
        let o = ''; for (const c of hs) o += (await para(c)).html; return o;
      }
      return listHtml(n, null, 1);
    }
    if (ln === 'table') return table(n);
    if (ln === 'section') {
      const sst = styles.get(`section:${A(n, 'text', 'style-name')}`)?.el; const cn = Number(A(sst?.getElementsByTagNameNS(NS.style, 'columns')[0], 'fo', 'column-count') || 1);
      let o = ''; for (const c of n.children) o += await block(c);
      return cn > 1 ? `<div class="cols" data-n="${cn}" style="column-count:${cn}">${o}</div>` : o;
    }
    if (ln === 'table-of-content' || ln === 'alphabetical-index' || ln === 'illustration-index' || ln === 'table-index' || ln === 'bibliography' || ln === 'user-index' || ln === 'object-index') { const b = n.getElementsByTagNameNS(NS.text, 'index-body')[0]; let o = ''; if (b) for (const c of b.children) o += await block(c); return o; }
    if (ln === 'index-title') { let o = ''; for (const c of n.children) o += await block(c); return o; }
    if (ln === 'frame') return frame(n);
    if (ln === 'soft-page-break' || ln === 'sequence-decls' || ln === 'variable-decls' || ln === 'user-field-decls' || ln === 'tracked-changes') return '';
    return '';
  };
  const body = content.getElementsByTagNameNS(NS.office, 'text')[0];
  let html = '';
  if (body) for (const c of body.children) html += await block(c);
  // salto antes de párrafo: se inserta el salto (salvo al principio)
  html = html.replace(/\uE002[^\uE003]*\uE003/g, '').replace(/^\uE0FF/, '').replace(//g, '<div class="pb" contenteditable="false"></div>');
  // página, encabezado y pie del estilo de página principal
  let page = null; let header = ''; let footer = '';
  const mp = stylesDoc?.getElementsByTagNameNS(NS.style, 'master-page')[0];
  const plName = A(mp, 'style', 'page-layout-name');
  const pl = [...(stylesDoc?.getElementsByTagNameNS(NS.style, 'page-layout') || [])].find((x) => A(x, 'style', 'name') === plName) || stylesDoc?.getElementsByTagNameNS(NS.style, 'page-layout')[0];
  const plp = pl && kid(pl, 'page-layout-properties');
  if (plp && A(plp, 'fo', 'page-width')) {
    const w = Math.round(lenToPx(A(plp, 'fo', 'page-width'))), hgt = Math.round(lenToPx(A(plp, 'fo', 'page-height')));
    const m = (k) => lenToPx(A(plp, 'fo', k) || A(plp, 'fo', 'margin') || '2cm');
    // en ODF el encabezado ocupa parte del margen: se suma su altura
    const hh = pl && kid(pl, 'header-style') && kid(kid(pl, 'header-style'), 'header-footer-properties');
    const fh = pl && kid(pl, 'footer-style') && kid(kid(pl, 'footer-style'), 'header-footer-properties');
    const extra = (x) => (x ? lenToPx(A(x, 'fo', 'min-height') || '0') + lenToPx(A(x, 'fo', 'margin-bottom') || A(x, 'fo', 'margin-top') || '0') : 0);
    page = { w, h: hgt, mt: Math.round(m('margin-top') + (mp && kid(mp, 'header') ? extra(hh) : 0)), mb: Math.round(m('margin-bottom') + (mp && kid(mp, 'footer') ? extra(fh) : 0)), ml: Math.round(m('margin-left')), mr: Math.round(m('margin-right')), landscape: w > hgt };
  }
  const hfText = (el) => {
    if (!el) return '';
    const lines = [];
    for (const p of el.getElementsByTagNameNS(NS.text, 'p')) {
      let s = '';
      const walk = (x) => { for (const c of x.childNodes) { if (c.nodeType === 3) s += c.nodeValue; else if (c.nodeType === 1) { if (c.localName === 'page-number') s += '{pagina}'; else if (c.localName === 'page-count') s += '{paginas}'; else if (c.localName === 'tab') s += '   '; else if (c.localName === 's') s += ' '; else walk(c); } } };
      walk(p);
      if (s.trim()) lines.push(s.trim());
    }
    return lines.join('\n');
  };
  if (mp) { header = hfText(kid(mp, 'header')); footer = hfText(kid(mp, 'footer')); }
  // numeración de capítulos (esquema de numeración de LibreOffice)
  let chapnum = null;
  const ol1 = [...(stylesDoc?.getElementsByTagNameNS(NS.text, 'outline-level-style') || [])].find((l) => A(l, 'text', 'level') === '1');
  const of = ol1 && A(ol1, 'style', 'num-format');
  if (!of) chapnum = chapFromList;
  if (of) chapnum = of === 'I' ? 'IA1' : (A(ol1, 'style', 'num-prefix') || '').trim() ? 'cap' : (A(ol1, 'style', 'num-suffix') || '').includes('.') ? '1.' : '1';
  const mg = (tag, ns = 'dc') => meta?.getElementsByTagNameNS(NS[ns], tag)[0]?.textContent || '';
  const box = document.createElement('div');
  box.innerHTML = sanitize(html || '<p><br></p>');
  let sources = [];
  try { const mt = await zip.file(META_ODT)?.async('string'); if (mt) ({ sources } = applyMeta(box, JSON.parse(mt), comments, null)); } catch { /* */ }
  return { html: box.innerHTML || '<p><br></p>', page, header, footer, defTab, chapnum, styles: Object.fromEntries(usedStyles), footnotes, endnotes, comments, sources, warnings: warn.size ? [`No se pudieron conservar del todo: ${[...warn].join(', ')}`] : [], props: { title: mg('title'), author: mg('creator') || mg('initial-creator', 'meta'), subject: mg('subject'), keywords: mg('keyword', 'meta') } };
}
