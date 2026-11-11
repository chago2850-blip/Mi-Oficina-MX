// Lector propio de .docx (Office Open XML) para Escribiente: conserva el formato de párrafo y de
// carácter (estilos, fuentes, tamaños, colores, espaciados, sangrías, bordes, sombreado), listas
// con su numeración, tablas con celdas combinadas y bordes, imágenes, vínculos, notas, comentarios,
// revisiones, ecuaciones, campos, saltos, encabezado y pie. GPL-3.0
import temml from 'temml';
import JSZip from 'jszip';
import { ommlToLatex, OMML_NS } from './omml.js';
import { META_DOCX, applyMeta } from './meta.js';
import { readComments, wrapMarks, sanitize, MK0, MK1, b64 } from './docx-in.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const TW_PT = 1 / 20; // twips → pt
const EMU_PX = 1 / 9525;
const TWIP_PX = 1 / 15;

// ------------------------------------------------------------------ utilidades XML
const kids = (el, name) => (el ? [...el.children].filter((c) => !name || c.localName === name) : []);
const kid = (el, name) => (el ? [...el.children].find((c) => c.localName === name) || null : null);
const av = (el, a) => {
  if (!el) return null;
  for (const x of el.attributes) if (x.localName === a && (x.namespaceURI === W || x.prefix === 'w' || !x.prefix)) return x.value;
  return null;
};
const avNs = (el, a) => { if (!el) return null; for (const x of el.attributes) if (x.localName === a) return x.value; return null; };
const num = (v) => (v == null || v === '' ? null : Number(v));
// Valor booleano de propiedades como <w:b/>, <w:b w:val="0"/>
const onOff = (el) => { if (!el) return undefined; const v = av(el, 'val'); return !(v === '0' || v === 'false' || v === 'off'); };
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const mk = (txt) => esc(MK0 + txt + MK1);

const HIGHLIGHT = { yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF', blue: '#0000FF', red: '#FF0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0', black: '#000000', white: '#FFFFFF' };
// fuentes con métricas compatibles que suelen estar en Linux
const FONT_ALT = { Calibri: 'Carlito', Cambria: 'Caladea', Arial: "'Liberation Sans'", 'Times New Roman': "'Liberation Serif'", 'Courier New': "'Liberation Mono'", Helvetica: "'Liberation Sans'", 'Calibri Light': 'Carlito', 'Segoe UI': "'Noto Sans'", Georgia: 'Gelasio', Verdana: "'DejaVu Sans'" };
// comillas simples: el CSS va dentro de style="…"
const fontCss = (f) => { if (!f) return null; const alt = FONT_ALT[f]; return `'${f.replace(/['"]/g, '')}'${alt ? `, ${alt}` : ''}`; };
const NUMFMT = { decimal: 'decimal', decimalZero: 'decimal-leading-zero', lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha', lowerRoman: 'lower-roman', upperRoman: 'upper-roman', ordinal: 'decimal', cardinalText: 'decimal', ordinalText: 'decimal', none: 'none' };

// ------------------------------------------------------------------ propiedades
function parseBorder(el) {
  if (!el) return undefined;
  const val = av(el, 'val');
  if (!val || val === 'nil' || val === 'none') return null;
  const sz = num(av(el, 'sz')) || 4; // octavos de punto
  const style = /double/i.test(val) ? 'double' : /dash/i.test(val) ? 'dashed' : /dot/i.test(val) ? 'dotted' : /inset/.test(val) ? 'inset' : /outset/.test(val) ? 'outset' : 'solid';
  const col = av(el, 'color');
  const w = Math.max(0.5, sz / 8) * (style === 'double' ? 3 : 1);
  return { css: `${+w.toFixed(2)}pt ${style} ${col && col !== 'auto' ? `#${col}` : '#000'}`, space: num(av(el, 'space')) || 0 };
}
function parseShd(el) {
  if (!el) return undefined;
  const fill = av(el, 'fill'); const val = av(el, 'val'); const color = av(el, 'color');
  if (fill && fill !== 'auto') return `#${fill}`;
  if (val === 'solid' && color && color !== 'auto') return `#${color}`;
  return null;
}
function parsePPr(pPr) {
  const o = {};
  if (!pPr) return o;
  for (const c of pPr.children) {
    switch (c.localName) {
      case 'pStyle': o.style = av(c, 'val'); break;
      case 'jc': o.jc = av(c, 'val'); break;
      case 'ind': {
        const l = av(c, 'left') ?? av(c, 'start'); const r = av(c, 'right') ?? av(c, 'end');
        if (l != null) o.indL = num(l); if (r != null) o.indR = num(r);
        if (av(c, 'firstLine') != null) { o.indFirst = num(av(c, 'firstLine')); o.indHang = 0; }
        if (av(c, 'hanging') != null) { o.indHang = num(av(c, 'hanging')); o.indFirst = 0; }
        break;
      }
      case 'spacing': {
        if (av(c, 'before') != null) o.before = num(av(c, 'before'));
        if (av(c, 'after') != null) o.after = num(av(c, 'after'));
        if (av(c, 'line') != null) { o.line = num(av(c, 'line')); o.lineRule = av(c, 'lineRule') || 'auto'; }
        if (onOff(kid(pPr, 'spacing')) && av(c, 'beforeAutospacing') === '1') o.before = 280;
        if (av(c, 'afterAutospacing') === '1') o.after = 280;
        break;
      }
      case 'keepNext': o.keepNext = onOff(c); break;
      case 'keepLines': o.keepLines = onOff(c); break;
      case 'pageBreakBefore': o.pageBreakBefore = onOff(c); break;
      case 'widowControl': o.widow = onOff(c); break;
      case 'contextualSpacing': o.contextual = onOff(c); break;
      case 'outlineLvl': o.outline = num(av(c, 'val')); break;
      case 'numPr': { const il = kid(c, 'ilvl'); const ni = kid(c, 'numId'); if (ni) o.numId = av(ni, 'val'); if (il) o.ilvl = num(av(il, 'val')); break; }
      case 'pBdr': { o.bdr = {}; for (const s of c.children) { const b = parseBorder(s); if (b !== undefined) o.bdr[s.localName === 'start' ? 'left' : s.localName === 'end' ? 'right' : s.localName] = b; } break; }
      case 'shd': o.shd = parseShd(c); break;
      case 'tabs': o.tabs = kids(c, 'tab').map((t) => ({ k: { left: 'l', start: 'l', center: 'c', right: 'r', end: 'r', decimal: 'r' }[av(t, 'val')], x: num(av(t, 'pos')) })).filter((t) => t.k && t.x != null); break;
      case 'framePr': o.frame = true; break;
      case 'rPr': o.markRPr = c; break;
      case 'sectPr': o.sectPr = c; break;
      default:
    }
  }
  return o;
}
function parseRPr(rPr, theme) {
  const o = {};
  if (!rPr) return o;
  for (const c of rPr.children) {
    switch (c.localName) {
      case 'rStyle': o.style = av(c, 'val'); break;
      case 'rFonts': {
        const f = av(c, 'ascii') || av(c, 'hAnsi') || av(c, 'cs');
        const th = av(c, 'asciiTheme') || av(c, 'hAnsiTheme');
        if (f) o.font = f; else if (th && theme) o.font = /major/i.test(th) ? theme.major : theme.minor;
        break;
      }
      case 'sz': o.size = num(av(c, 'val')) / 2; break;
      case 'b': o.b = onOff(c); break;
      case 'i': o.i = onOff(c); break;
      case 'u': { const v = av(c, 'val'); o.u = v && v !== 'none' ? v : false; const col = av(c, 'color'); if (col && col !== 'auto') o.uColor = `#${col}`; break; }
      case 'strike': o.strike = onOff(c); break;
      case 'dstrike': o.dstrike = onOff(c); break;
      case 'caps': o.caps = onOff(c); break;
      case 'smallCaps': o.smallCaps = onOff(c); break;
      case 'vanish': o.vanish = onOff(c); break;
      case 'shadow': o.shadow = onOff(c); break;
      case 'outline': o.outline = onOff(c); break;
      case 'emboss': o.emboss = onOff(c); break;
      case 'imprint': o.imprint = onOff(c); break;
      case 'color': { const v = av(c, 'val'); o.color = v && v !== 'auto' ? `#${v}` : null; break; }
      case 'highlight': { const v = av(c, 'val'); o.highlight = HIGHLIGHT[v] || null; break; }
      case 'shd': o.shd = parseShd(c); break;
      case 'vertAlign': o.va = av(c, 'val'); break;
      case 'spacing': o.spacing = num(av(c, 'val')); break; // twips
      case 'w': o.scale = num(av(c, 'val')); break;
      case 'position': o.position = num(av(c, 'val')) / 2; break; // medios puntos
      case 'bdr': o.bdr = parseBorder(c); break;
      case 'lang': o.lang = av(c, 'val'); break;
      default:
    }
  }
  return o;
}
const merge = (...xs) => Object.assign({}, ...xs.filter(Boolean));

// ------------------------------------------------------------------ estilos
function readStyles(doc, theme) {
  const styles = new Map(); const defaults = { p: {}, r: {} }; const byName = new Map();
  if (!doc) return { styles, defaults, byName, defaultPara: null, defaultTable: null };
  const dd = doc.getElementsByTagNameNS(W, 'docDefaults')[0];
  if (dd) {
    defaults.r = parseRPr(dd.getElementsByTagNameNS(W, 'rPr')[0], theme);
    defaults.p = parsePPr(dd.getElementsByTagNameNS(W, 'pPr')[0]);
  }
  let defaultPara = null; let defaultTable = null;
  for (const s of doc.getElementsByTagNameNS(W, 'style')) {
    const id = av(s, 'styleId'); const type = av(s, 'type');
    const name = av(kid(s, 'name'), 'val') || id;
    const tblPr = kid(s, 'tblPr');
    const st = { id, type, name, basedOn: av(kid(s, 'basedOn'), 'val'), p: parsePPr(kid(s, 'pPr')), r: parseRPr(kid(s, 'rPr'), theme), tbl: tblPr ? parseTblPr(tblPr) : null };
    styles.set(id, st); byName.set(name.toLowerCase(), st);
    if (av(s, 'default') === '1' || av(s, 'default') === 'true') { if (type === 'paragraph') defaultPara = id; if (type === 'table') defaultTable = id; }
  }
  return { styles, defaults, byName, defaultPara, defaultTable };
}
function styleChain(S, id) {
  const out = []; let cur = id; let guard = 0;
  while (cur && S.styles.has(cur) && guard++ < 20) { const s = S.styles.get(cur); out.unshift(s); cur = s.basedOn; }
  return out;
}
function parseTblPr(tblPr) {
  const o = {};
  const b = kid(tblPr, 'tblBorders');
  if (b) { o.bdr = {}; for (const s of b.children) { const v = parseBorder(s); if (v !== undefined) o.bdr[s.localName === 'start' ? 'left' : s.localName === 'end' ? 'right' : s.localName] = v; } }
  const w = kid(tblPr, 'tblW'); if (w) o.w = { v: num(av(w, 'w')), t: av(w, 'type') };
  const jc = kid(tblPr, 'jc'); if (jc) o.jc = av(jc, 'val');
  const shd = kid(tblPr, 'shd'); if (shd) o.shd = parseShd(shd);
  const mar = kid(tblPr, 'tblCellMar');
  if (mar) { o.mar = {}; for (const m of mar.children) o.mar[m.localName === 'start' ? 'left' : m.localName === 'end' ? 'right' : m.localName] = num(av(m, 'w')); }
  const ind = kid(tblPr, 'tblInd'); if (ind) o.ind = num(av(ind, 'w'));
  const lay = kid(tblPr, 'tblLayout'); if (lay) o.fixed = av(lay, 'type') === 'fixed';
  return o;
}

// ------------------------------------------------------------------ CSS
const pt = (v) => `${+(v).toFixed(2)}pt`;
function paraCss(p, r, { inList = false } = {}) {
  const css = [];
  if (p.jc) { const a = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify', left: 'left', start: 'left' }[p.jc]; if (a) css.push(`text-align:${a}`); }
  css.push(`margin-top:${pt((p.before || 0) * TW_PT)}`, `margin-bottom:${pt((p.after || 0) * TW_PT)}`);
  if (!inList) {
    const l = (p.indL || 0) * TW_PT; const first = p.indHang ? -p.indHang * TW_PT : (p.indFirst || 0) * TW_PT;
    if (l) css.push(`margin-left:${pt(l)}`);
    if (first) css.push(`text-indent:${pt(first)}`);
  }
  if (p.indR) css.push(`margin-right:${pt(p.indR * TW_PT)}`);
  if (p.line) {
    if (p.lineRule === 'exact' || p.lineRule === 'atLeast') css.push(`line-height:${pt(p.line * TW_PT)}`);
    else css.push(`line-height:${+(p.line / 240).toFixed(3)}`);
  }
  if (p.bdr) {
    for (const side of ['top', 'right', 'bottom', 'left']) {
      const b = p.bdr[side]; if (!b) continue;
      css.push(`border-${side}:${b.css}`, `padding-${side}:${pt(b.space || 1)}`);
    }
  }
  if (p.shd) css.push(`background-color:${p.shd}`);
  // letra base del párrafo
  if (r.font) css.push(`font-family:${fontCss(r.font)}`);
  if (r.size) css.push(`font-size:${pt(r.size)}`);
  if (r.color) css.push(`color:${r.color}`);
  if (r.b) css.push('font-weight:bold'); else if (r.b === false) css.push('font-weight:normal');
  if (r.i) css.push('font-style:italic'); else if (r.i === false) css.push('font-style:normal');
  if (r.caps) css.push('text-transform:uppercase');
  if (r.smallCaps) css.push('font-variant:small-caps');
  if (r.spacing) css.push(`letter-spacing:${pt(r.spacing * TW_PT)}`);
  return css.join(';');
}
function paraData(p) {
  const d = [];
  if (p.keepNext) d.push('data-keepnext="1"');
  if (p.keepLines) d.push('data-keeplines="1"');
  if (p.pageBreakBefore) d.push('data-pbb="1"');
  if (p.widow === false) d.push('data-widow="0"');
  if (p.line && p.lineRule === 'atLeast') d.push('data-lr="atLeast"');
  if (p.tabs?.length) d.push(`data-tabs="${p.tabs.map((t) => `${t.k}:${Math.round(t.x * TWIP_PX)}`).join(',')}"`);
  return d.length ? ` ${d.join(' ')}` : '';
}
// diferencias de formato del texto respecto del párrafo
function runWrap(r, base, html) {
  if (!html) return html;
  const css = [];
  if (r.font && r.font !== base.font) css.push(`font-family:${fontCss(r.font)}`);
  if (r.size && r.size !== base.size) css.push(`font-size:${pt(r.size)}`);
  if (r.color && r.color !== base.color) css.push(`color:${r.color}`);
  const bg = r.highlight || r.shd; if (bg) css.push(`background-color:${bg}`);
  if (r.caps && !base.caps) css.push('text-transform:uppercase');
  if (r.caps === false && base.caps) css.push('text-transform:none');
  if (r.smallCaps && !base.smallCaps) css.push('font-variant:small-caps');
  if (r.spacing && r.spacing !== base.spacing) css.push(`letter-spacing:${pt(r.spacing * TW_PT)}`);
  if (r.position) css.push(`vertical-align:${pt(r.position)}`);
  if (r.shadow) css.push('text-shadow:1px 1px 1px rgba(0,0,0,.45)');
  if (r.outline) css.push('-webkit-text-stroke:.6px currentColor;color:transparent');
  if (r.emboss || r.imprint) css.push(`text-shadow:${r.emboss ? '-1px -1px 0 #fff, 1px 1px 0 #999' : '1px 1px 0 #fff, -1px -1px 0 #999'}`);
  if (r.bdr) css.push(`border:${r.bdr.css};padding:0 1pt`);
  if (r.scale && r.scale !== 100) css.push(`display:inline-block;transform:scaleX(${r.scale / 100});transform-origin:left`);
  if (r.u && r.u !== 'single' && r.u !== 'words') css.push(`text-decoration:underline ${/double/i.test(r.u) ? 'double' : /dot/i.test(r.u) ? 'dotted' : /dash/i.test(r.u) ? 'dashed' : /wav/i.test(r.u) ? 'wavy' : 'solid'}${r.uColor ? ` ${r.uColor}` : ''}`);
  if (r.dstrike) css.push('text-decoration:line-through double');
  if (r.b === false && base.b) css.push('font-weight:normal');
  if (r.i === false && base.i) css.push('font-style:normal');
  let out = html;
  if (r.b && !base.b) out = `<b>${out}</b>`;
  if (r.i && !base.i) out = `<i>${out}</i>`;
  if (r.u && (r.u === 'single' || r.u === 'words')) out = `<u>${out}</u>`;
  if (r.strike) out = `<s>${out}</s>`;
  if (r.va === 'superscript') out = `<sup>${out}</sup>`;
  if (r.va === 'subscript') out = `<sub>${out}</sub>`;
  const cls = r.vanish ? ' class="hidden-text"' : '';
  if (css.length || cls) out = `<span${cls}${css.length ? ` style="${css.join(';')}"` : ''}>${out}</span>`;
  return out;
}

// ------------------------------------------------------------------ lector
export async function readDocx(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const xml = async (p) => { const t = await zip.file(p)?.async('string'); return t ? new DOMParser().parseFromString(t, 'application/xml') : null; };
  const docXml = await xml('word/document.xml');
  if (!docXml) throw new Error('No es un documento de Word válido');
  const warnings = new Set();
  // tema (fuentes mayor/menor)
  const theme = { minor: 'Calibri', major: 'Calibri Light' };
  const th = await xml('word/theme/theme1.xml');
  if (th) {
    const mf = th.getElementsByTagNameNS(A, 'minorFont')[0]; const Mf = th.getElementsByTagNameNS(A, 'majorFont')[0];
    const lat = (f) => avNs(f?.getElementsByTagNameNS(A, 'latin')[0], 'typeface');
    theme.minor = lat(mf) || theme.minor; theme.major = lat(Mf) || theme.major;
  }
  const S = readStyles(await xml('word/styles.xml'), theme);
  // relaciones
  const rels = async (p) => {
    const m = new Map(); const d = await xml(p); if (!d) return m;
    for (const r of d.getElementsByTagName('Relationship')) m.set(r.getAttribute('Id'), { target: r.getAttribute('Target'), type: r.getAttribute('Type'), mode: r.getAttribute('TargetMode') });
    return m;
  };
  const docRels = await rels('word/_rels/document.xml.rels');
  const mediaCache = new Map();
  const media = async (relMap, id, base = 'word/') => {
    const rel = relMap.get(id); if (!rel || rel.mode === 'External') return null;
    const path = rel.target.startsWith('/') ? rel.target.slice(1) : (base + rel.target).replace(/[^/]+\/\.\.\//g, '');
    if (mediaCache.has(path)) return mediaCache.get(path);
    const f = zip.file(path); if (!f) return null;
    const ext = path.split('.').pop().toLowerCase();
    const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff' }[ext];
    if (!mime) { warnings.add(`imagen ${ext.toUpperCase()}`); return null; }
    const url = `data:${mime};base64,${await f.async('base64')}`;
    mediaCache.set(path, url); return url;
  };
  // numeración
  const numbering = { nums: new Map(), abstracts: new Map() };
  const nd = await xml('word/numbering.xml');
  if (nd) {
    for (const a of nd.getElementsByTagNameNS(W, 'abstractNum')) {
      const lv = new Map();
      for (const l of kids(a, 'lvl')) lv.set(num(av(l, 'ilvl')), { start: num(av(kid(l, 'start'), 'val')) ?? 1, fmt: av(kid(l, 'numFmt'), 'val') || 'decimal', text: av(kid(l, 'lvlText'), 'val') ?? '', p: parsePPr(kid(l, 'pPr')), r: parseRPr(kid(l, 'rPr'), theme) });
      numbering.abstracts.set(av(a, 'abstractNumId'), { levels: lv, styleLink: av(kid(a, 'numStyleLink'), 'val') });
    }
    for (const n of nd.getElementsByTagNameNS(W, 'num')) {
      const over = new Map();
      for (const o of kids(n, 'lvlOverride')) { const so = kid(o, 'startOverride'); over.set(num(av(o, 'ilvl')), so ? num(av(so, 'val')) : null); }
      numbering.nums.set(av(n, 'numId'), { abs: av(kid(n, 'abstractNumId'), 'val'), over });
    }
  }
  const counters = new Map(); // abs -> [n por nivel]
  const levelOf = (numId, ilvl) => { const n = numbering.nums.get(numId); if (!n) return null; const a = numbering.abstracts.get(n.abs); return a ? a.levels.get(ilvl) || a.levels.get(0) : null; };
  // comentarios, notas
  const rc = await readComments(zip);
  const replyIds = rc?.replyIds || new Set();
  const footnotes = {}; const endnotes = {};
  const noteMaps = { footnote: new Map(), endnote: new Map() };
  for (const kind of ['footnote', 'endnote']) {
    const d = await xml(`word/${kind}s.xml`); if (!d) continue;
    for (const n of d.getElementsByTagNameNS(W, kind)) {
      const t = av(n, 'type'); if (t && t !== 'normal') continue;
      const txt = [...n.getElementsByTagNameNS(W, 'p')].map((p) => [...p.getElementsByTagNameNS(W, 't')].map((x) => x.textContent).join('')).join('\n').trim();
      noteMaps[kind].set(av(n, 'id'), txt);
    }
  }

  // ---------------------------------------------------------------- conversión
  const effPara = (pPr) => {
    const direct = parsePPr(pPr);
    const sid = direct.style || S.defaultPara;
    const chain = styleChain(S, sid);
    const sty = chain[chain.length - 1];
    // dentro de una tabla, el estilo de tabla manda sobre el estilo Normal (no sobre otros estilos)
    const tsc = ctx.tableStyles[ctx.tableStyles.length - 1] || [];
    const isDefault = !direct.style || direct.style === S.defaultPara;
    const tp = tsc.map((s) => s.p); const tr = tsc.map((s) => s.r);
    const p = isDefault ? merge(S.defaults.p, ...chain.map((s) => s.p), ...tp, direct) : merge(S.defaults.p, ...tp, ...chain.map((s) => s.p), direct); p.style = sid;
    const r = isDefault ? merge(S.defaults.r, ...chain.map((s) => s.r), ...tr) : merge(S.defaults.r, ...tr, ...chain.map((s) => s.r));
    if (!r.size) r.size = 10;
    // numeración heredada del estilo
    if (p.numId == null) for (const s of chain) if (s.p.numId != null) { p.numId = s.p.numId; if (p.ilvl == null) p.ilvl = s.p.ilvl; }
    return { p, r, sty, name: (sty?.name || '').toLowerCase() };
  };
  const effRun = (rPr, base) => {
    const direct = parseRPr(rPr, theme);
    const chain = direct.style ? styleChain(S, direct.style) : [];
    return merge(base, ...chain.map((s) => s.r), direct);
  };
  const ctx = { field: [], hyperlink: null, tableStyles: [] };

  async function drawing(el, relMap, base) {
    // imagen en línea o anclada; cuadros de texto
    const inline = el.getElementsByTagNameNS('*', 'inline')[0]; const anchor = el.getElementsByTagNameNS('*', 'anchor')[0];
    const holder = inline || anchor; if (!holder) return '';
    const ext = [...holder.children].find((c) => c.localName === 'extent');
    const w = Math.round(num(avNs(ext, 'cx')) * EMU_PX) || null; const hh = Math.round(num(avNs(ext, 'cy')) * EMU_PX) || null;
    const txbx = holder.getElementsByTagNameNS('*', 'txbxContent')[0];
    if (txbx) { const inner = await blocks(txbx, relMap, base); return { box: `<div class="tbox"${w ? ` style="width:${w}px"` : ''}>${inner || '<p><br></p>'}</div>` }; }
    const blip = holder.getElementsByTagNameNS(A, 'blip')[0];
    if (!blip) { if (holder.getElementsByTagNameNS('*', 'chart').length) warnings.add('gráficos'); else warnings.add('formas o diagramas'); return ''; }
    const src = await media(relMap, avNs(blip, 'embed'), base);
    if (!src) return '';
    let style = '';
    if (anchor) {
      const ph = [...anchor.children].find((c) => c.localName === 'positionH');
      const al = ph && [...ph.children].find((c) => c.localName === 'align')?.textContent;
      const behind = avNs(anchor, 'behindDoc') === '1';
      const wrapNone = [...anchor.children].some((c) => c.localName === 'wrapNone' || c.localName === 'wrapTopAndBottom');
      if (!wrapNone && !behind) style = al === 'right' ? 'float:right;margin:0 0 6pt 9pt' : al === 'center' ? '' : 'float:left;margin:0 9pt 6pt 0';
    }
    const descr = esc(avNs(holder.getElementsByTagNameNS('*', 'docPr')[0], 'descr') || '');
    return `<img src="${src}"${w ? ` width="${w}"` : ''}${hh ? ` height="${hh}"` : ''} alt="${descr}"${style ? ` style="${style}"` : ''}>`;
  }

  // contenido en línea de un párrafo (runs, vínculos, campos, revisiones…)
  async function inline(parent, base, relMap, b, out) {
    for (const n of parent.children) {
      const ln = n.localName;
      if (n.namespaceURI === OMML_NS && (ln === 'oMath' || ln === 'oMathPara')) { const tex = ommlToLatex(n); if (tex) out.html += mk(`q:${b64(tex)}`); continue; }
      switch (ln) {
        case 'r': await run(n, base, relMap, b, out); break;
        case 'hyperlink': {
          const id = avNs(n, 'id'); const anchorName = av(n, 'anchor');
          const href = id ? docRelsHref(relMap, id) : anchorName ? `#${anchorName}` : null;
          const sub = { html: '', boxes: out.boxes, breakAfter: out.breakAfter };
          await inline(n, base, relMap, b, sub);
          out.html += href ? `<a href="${esc(href)}">${sub.html}</a>` : sub.html;
          break;
        }
        case 'fldSimple': {
          const instr = (av(n, 'instr') || '').trim();
          const sub = { html: '', boxes: out.boxes };
          await inline(n, base, relMap, b, sub);
          const hm = /^HYPERLINK\s+"([^"]+)"/i.exec(instr);
          out.html += hm ? `<a href="${esc(hm[1])}">${sub.html}</a>` : sub.html;
          break;
        }
        case 'ins': case 'moveTo': out.html += mk(`i:${[av(n, 'author'), av(n, 'date')].map((x) => (x || '').replace(/[|]/g, '')).join('|')}`); await inline(n, base, relMap, b, out); out.html += mk('/i'); break;
        case 'del': case 'moveFrom': out.html += mk(`d:${[av(n, 'author'), av(n, 'date')].map((x) => (x || '').replace(/[|]/g, '')).join('|')}`); await inline(n, base, relMap, b, out); out.html += mk('/d'); break;
        case 'commentRangeStart': { const id = av(n, 'id'); if (!replyIds.has(id)) out.html += mk(`c:${id}`); break; }
        case 'commentRangeEnd': { const id = av(n, 'id'); if (!replyIds.has(id)) out.html += mk(`/c:${id}`); break; }
        case 'bookmarkStart': { const nm = av(n, 'name'); if (nm && !nm.startsWith('_GoBack')) out.html += `<a id="${esc(nm)}" class="bm"></a>`; break; }
        case 'smartTag': case 'customXml': case 'sdt': case 'sdtContent': await inline(ln === 'sdt' ? (kid(n, 'sdtContent') || n) : n, base, relMap, b, out); break;
        case 'AlternateContent': { const ch = kid(n, 'Choice') || kid(n, 'Fallback'); if (ch) await inline(ch, base, relMap, b, out); break; }
        default:
      }
    }
  }
  function docRelsHref(relMap, id) { const r = relMap.get(id); return r ? r.target : null; }

  async function run(rEl, base, relMap, b, out) {
    const r = effRun(kid(rEl, 'rPr'), base);
    let txt = '';
    const flush = () => { if (txt) { emitText(txt); txt = ''; } };
    const emitText = (t) => {
      if (ctx.field.length && ctx.field[ctx.field.length - 1].phase === 'instr') return;
      out.html += runWrap(r, base, esc(t).replace(/\n/g, '<br>'));
    };
    for (const c of rEl.children) {
      switch (c.localName) {
        case 't': case 'delText': txt += c.textContent; break;
        case 'tab': txt += '\t'; break;
        case 'br': {
          const type = av(c, 'type');
          if (type === 'page') { flush(); out.breakAfter = true; out.html += ''; } else if (type === 'column') txt += '\n'; else txt += '\n';
          break;
        }
        case 'cr': txt += '\n'; break;
        case 'noBreakHyphen': txt += '‑'; break;
        case 'softHyphen': txt += '­'; break;
        case 'sym': { const ch = parseInt(av(c, 'char') || '0', 16); txt += String.fromCodePoint(ch >= 0xF000 ? ch - 0xF000 : ch); break; }
        case 'fldChar': {
          flush();
          const t = av(c, 'fldCharType');
          if (t === 'begin') ctx.field.push({ phase: 'instr', instr: '', start: out.html.length, out });
          else if (t === 'separate') { const f = ctx.field[ctx.field.length - 1]; if (f) { f.phase = 'result'; f.start = out.html.length; f.out = out; } }
          else if (t === 'end') {
            const f = ctx.field.pop();
            if (f) {
              const hm = /HYPERLINK\s+(?:\\l\s+)?"([^"]+)"/i.exec(f.instr);
              if (hm && f.out === out) { const inner = out.html.slice(f.start); out.html = `${out.html.slice(0, f.start)}<a href="${esc(/\\l/.test(f.instr) ? `#${hm[1]}` : hm[1])}">${inner}</a>`; }
            }
          }
          break;
        }
        case 'instrText': { const f = ctx.field[ctx.field.length - 1]; if (f) f.instr += c.textContent; break; }
        case 'drawing': case 'pict': case 'object': {
          flush();
          if (c.localName === 'drawing') { const d = await drawing(c, relMap, b); if (typeof d === 'string') out.html += d; else if (d?.box) out.boxes.push(d.box); }
          else {
            const tb = c.getElementsByTagNameNS('*', 'txbxContent')[0];
            if (tb) out.boxes.push(`<div class="tbox">${await blocks(tb, relMap, b)}</div>`);
            const im = c.getElementsByTagNameNS('*', 'imagedata')[0];
            if (im) { const src = await media(relMap, avNs(im, 'id'), b); if (src) out.html += `<img src="${src}" alt="">`; }
            else if (!tb && c.localName === 'object') warnings.add('objetos incrustados');
          }
          break;
        }
        case 'AlternateContent': { flush(); const ch = kid(c, 'Choice') || kid(c, 'Fallback'); if (ch) for (const x of ch.children) { if (x.localName === 'drawing') { const d = await drawing(x, relMap, b); if (typeof d === 'string') out.html += d; else if (d?.box) out.boxes.push(d.box); } else if (x.localName === 'pict') { const tb = x.getElementsByTagNameNS('*', 'txbxContent')[0]; if (tb) out.boxes.push(`<div class="tbox">${await blocks(tb, relMap, b)}</div>`); } } break; }
        case 'footnoteReference': case 'endnoteReference': {
          flush();
          const kind = c.localName === 'footnoteReference' ? 'footnote' : 'endnote';
          const txtN = noteMaps[kind].get(av(c, 'id'));
          if (txtN != null) { const fid = `${kind[0]}${Math.random().toString(36).slice(2, 8)}`; (kind === 'footnote' ? footnotes : endnotes)[fid] = txtN; out.html += `<sup class="${kind === 'footnote' ? 'fn' : 'en'}" data-fid="${fid}" contenteditable="false">#</sup>`; }
          break;
        }
        case 'lastRenderedPageBreak': case 'rPr': case 'footnoteRef': case 'endnoteRef': case 'separator': case 'continuationSeparator': break;
        case 'commentReference': break;
        default:
      }
    }
    flush();
  }

  // listas: pila de listas abiertas
  function listInfo(p) {
    if (p.numId == null || p.numId === '0') return null;
    const lvl = levelOf(p.numId, p.ilvl || 0); if (!lvl) return null;
    if (lvl.fmt === 'none' && !lvl.text) return null;
    return { lvl, ilvl: p.ilvl || 0, numId: p.numId, abs: numbering.nums.get(p.numId)?.abs };
  }
  function nextNumber(li) {
    const key = li.abs;
    if (!counters.has(key)) counters.set(key, { vals: [], numIds: new Set() });
    const c = counters.get(key);
    const n = numbering.nums.get(li.numId);
    // la primera vez que aparece un numId con startOverride, se reinicia
    if (!c.numIds.has(li.numId)) { c.numIds.add(li.numId); if (n?.over.has(li.ilvl) && n.over.get(li.ilvl) != null) c.vals[li.ilvl] = n.over.get(li.ilvl) - 1; }
    if (c.vals[li.ilvl] == null) c.vals[li.ilvl] = li.lvl.start - 1;
    c.vals[li.ilvl] += 1;
    c.vals.length = li.ilvl + 1; // niveles inferiores se reinician
    return c.vals[li.ilvl];
  }
  const numberText = (li) => {
    const c = counters.get(li.abs); const a = numbering.abstracts.get(li.abs);
    return li.lvl.text.replace(/%(\d)/g, (m, k) => { const i = Number(k) - 1; const v = c?.vals[i] ?? 1; const f = a?.levels.get(i)?.fmt || 'decimal'; return fmtNum(v, f); });
  };

  async function paragraph(pEl, relMap, b, opts = {}) {
    const { p, r, name } = effPara(kid(pEl, 'pPr'));
    const out = { html: '', boxes: [], breakAfter: false };
    await inline(pEl, r, relMap, b, out);
    const li = opts.noLists ? null : listInfo(p);
    // tipo de bloque
    let tag = 'p'; let cls = '';
    const hm = /^heading (\d)$/.exec(name) || /^(?:título|titulo|encabezado) (\d)$/.exec(name);
    const outline = p.outline != null && p.outline < 6 ? p.outline + 1 : null;
    if (hm || (outline && /heading|título|titulo|encabezado/.test(name))) tag = `h${Math.min(6, Number(hm ? hm[1] : outline))}`;
    else if (name === 'title' || name === 'título' || name === 'titulo') cls = 'title';
    else if (name === 'subtitle' || name === 'subtítulo' || name === 'subtitulo') cls = 'subtitle';
    else if (/^(quote|intense quote|cita|cita destacada)$/.test(name)) tag = 'blockquote';
    let content = out.html;
    const hasPB = false;
    // número de títulos numerados: se escribe como texto (las listas de títulos llegan en la parte 2)
    if (li && tag !== 'p') { const n = nextNumber(li); void n; content = `${esc(numberText(li))}\t${content}`; }
    const style = paraCss(p, r, { inList: !!li && tag === 'p' });
    const open = `<${tag}${cls ? ` class="${cls}"` : ''} style="${style}"${paraData(p)}>`;
    // saltos de página dentro del párrafo: se parte en dos con el salto en medio
    const parts = content.split('\uE0FF');
    const PB = '<div class="pb" contenteditable="false"></div>';
    const empty = (x) => !x.replace(/<a id="[^"]*" class="bm"><\/a>/g, '').trim();
    // un párrafo que solo contiene el salto de página es solo el salto
    const html = parts.length > 1 && parts.every(empty) ? PB : parts.map((x, i) => (empty(x) && parts.length > 1 && (i === 0 || i === parts.length - 1) ? '' : `${open}${empty(x) ? `${x}<br>` : x}</${tag}>`)).join(PB);
    return { html, li: tag === 'p' ? li : null, p, r, boxes: out.boxes, pageBreak: hasPB, sect: p.sectPr, styleId: p.style };
  }

  async function table(tbl, relMap, b) {
    const tblPr = kid(tbl, 'tblPr');
    const direct = tblPr ? parseTblPr(tblPr) : {};
    const sid = av(kid(tblPr, 'tblStyle'), 'val') || S.defaultTable;
    const chain = styleChain(S, sid);
    const tp = merge(...chain.map((s) => s.tbl || {}), direct);
    const bdr = merge(...chain.map((s) => s.tbl?.bdr || {}), direct.bdr || {});
    const grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => Math.round(num(av(g, 'w')) * TWIP_PX));
    const rows = kids(tbl, 'tr');
    // celdas con su posición en la cuadrícula (para rowspan de vMerge)
    const matrix = [];
    for (const tr of rows) {
      const cells = []; let col = 0;
      const trPr = kid(tr, 'trPr');
      const gb = num(av(kid(trPr, 'gridBefore'), 'val')) || 0; col += gb;
      for (const tc of kids(tr, 'tc')) {
        const tcPr = kid(tc, 'tcPr');
        const span = num(av(kid(tcPr, 'gridSpan'), 'val')) || 1;
        const vm = kid(tcPr, 'vMerge'); const vmv = vm ? (av(vm, 'val') || 'continue') : null;
        cells.push({ tc, tcPr, col, span, vm: vmv, rowspan: 1 });
        col += span;
      }
      matrix.push({ tr, trPr, cells });
    }
    for (let i = 0; i < matrix.length; i++) for (const c of matrix[i].cells) {
      if (c.vm !== 'restart') continue;
      for (let j = i + 1; j < matrix.length; j++) { const below = matrix[j].cells.find((x) => x.col === c.col); if (below && below.vm === 'continue') { c.rowspan++; below.skip = true; } else break; }
    }
    const nCols = Math.max(grid.length, ...matrix.map((m) => m.cells.reduce((a, c) => Math.max(a, c.col + c.span), 0)));
    const totalW = grid.reduce((a, x) => a + x, 0);
    const tcss = ['border-collapse:collapse'];
    if (tp.w?.t === 'pct') tcss.push(`width:${(tp.w.v / 50).toFixed(1)}%`);
    else if (totalW) tcss.push(`width:min(${totalW}px, 100%)`);
    if (tp.jc === 'center') tcss.push('margin-left:auto;margin-right:auto'); else if (tp.jc === 'right' || tp.jc === 'end') tcss.push('margin-left:auto;margin-right:0');
    else if (tp.ind) tcss.push(`margin-left:${Math.round(tp.ind * TWIP_PX)}px`);
    const mar = merge({ left: 108, right: 108, top: 0, bottom: 0 }, tp.mar || {});
    ctx.tableStyles.push(chain.filter((x) => Object.keys(x.p).length || Object.keys(x.r).length));
    let html = `<table style="${tcss.join(';')}">`;
    // anchos de columna en %: la tabla se ajusta si es más ancha que la página
    if (grid.length && totalW) html += `<colgroup>${grid.map((g) => `<col style="width:${((g / totalW) * 100).toFixed(2)}%">`).join('')}</colgroup>`;
    html += '<tbody>';
    for (let i = 0; i < matrix.length; i++) {
      const { trPr, cells } = matrix[i];
      const hgt = kid(trPr, 'trHeight'); const isHeader = onOff(kid(trPr, 'tblHeader'));
      html += `<tr${hgt ? ` style="height:${Math.round(num(av(hgt, 'val')) * TWIP_PX)}px"` : ''}>`;
      for (const c of cells) {
        if (c.skip) continue;
        const tcp = c.tcPr;
        const own = {}; const tb = kid(tcp, 'tcBorders');
        if (tb) for (const s of tb.children) { const v = parseBorder(s); if (v !== undefined) own[s.localName === 'start' ? 'left' : s.localName === 'end' ? 'right' : s.localName] = v; }
        const lastRow = i + c.rowspan - 1 === matrix.length - 1; const lastCol = c.col + c.span >= nCols;
        const side = (s, outer, inner) => (own[s] !== undefined ? own[s] : (outer ? bdr[s] : bdr[inner]));
        const bs = { top: side('top', i === 0, 'insideH'), bottom: side('bottom', lastRow, 'insideH'), left: side('left', c.col === 0, 'insideV'), right: side('right', lastCol, 'insideV') };
        const css = Object.entries(bs).map(([s, v]) => `border-${s}:${v ? v.css : 'none'}`);
        const shd = parseShd(kid(tcp, 'shd')) || tp.shd; css.push(`background-color:${shd || 'transparent'}`);
        if (isHeader) css.push('font-weight:inherit', 'text-align:inherit');
        const va = av(kid(tcp, 'vAlign'), 'val'); if (va) css.push(`vertical-align:${va === 'center' ? 'middle' : va === 'bottom' ? 'bottom' : 'top'}`);
        const cm = kid(tcp, 'tcMar'); const cmar = { ...mar }; if (cm) for (const m of cm.children) cmar[m.localName === 'start' ? 'left' : m.localName === 'end' ? 'right' : m.localName] = num(av(m, 'w'));
        css.push(`padding:${Math.round(cmar.top * TWIP_PX)}px ${Math.round(cmar.right * TWIP_PX)}px ${Math.round(cmar.bottom * TWIP_PX)}px ${Math.round(cmar.left * TWIP_PX)}px`);
        const tag = isHeader ? 'th' : 'td';
        const inner = await blocks(c.tc, relMap, b);
        html += `<${tag}${c.span > 1 ? ` colspan="${c.span}"` : ''}${c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : ''} style="${css.join(';')}">${inner || '<p><br></p>'}</${tag}>`;
      }
      html += '</tr>';
    }
    ctx.tableStyles.pop();
    return `${html}</tbody></table>`;
  }

  // bloques (cuerpo, celdas, cuadros de texto)
  async function blocks(container, relMap, b, sections = null) {
    let html = '';
    const stack = []; // listas abiertas {tag, ilvl}
    const closeTo = (lvl) => { while (stack.length && stack[stack.length - 1].ilvl >= lvl) { html += `</li></${stack.pop().tag}>`; } };
    const closeAll = () => closeTo(-1);
    let prevStyle = null; let prevCtx = false;
    const nodes = [];
    const flat = (el) => { for (const n of el.children) { if (n.localName === 'sdt') { const c = kid(n, 'sdtContent'); if (c) flat(c); } else if (n.localName === 'customXml' || n.localName === 'smartTag') flat(n); else nodes.push(n); } };
    flat(container);
    for (const n of nodes) {
      if (n.localName === 'p') {
        const pr = await paragraph(n, relMap, b);
        let phtml = pr.html;
        // espaciado contextual: sin espacio entre párrafos del mismo estilo
        if (pr.p.contextual && prevCtx && pr.styleId === prevStyle) {
          phtml = phtml.replace(/margin-top:[^;"]*/, 'margin-top:0pt');
          html = html.replace(/margin-bottom:[^;"]*(;[^"]*"[^>]*>(?:(?!<\/?(?:p|h\d|blockquote)\b)[\s\S])*<\/(?:p|h\d|blockquote)>(?:<\/li>)?)$/, 'margin-bottom:0pt$1');
        }
        prevStyle = pr.styleId; prevCtx = !!pr.p.contextual;
        if (pr.li) {
          const lv = pr.li.ilvl; const fmt = pr.li.lvl.fmt;
          const tag = fmt === 'bullet' ? 'ul' : 'ol';
          const top = stack[stack.length - 1];
          if (top && top.ilvl > lv) closeTo(lv + 1);
          const cur = stack[stack.length - 1];
          const n2 = nextNumber(pr.li);
          if (cur && cur.ilvl === lv && cur.tag === tag && cur.numId === pr.li.numId) html += '</li>';
          else {
            if (cur && cur.ilvl === lv) { html += `</li></${stack.pop().tag}>`; }
            const bullet = pr.li.lvl.text;
            const lst = tag === 'ul' ? (/[○o◦]/.test(bullet) ? 'circle' : /[■▪§]/.test(bullet) ? 'square' : 'disc') : NUMFMT[fmt] || 'decimal';
            const ind = pr.li.lvl.p; const left = ((pr.p.indL ?? ind.indL) || 360 * (lv + 1)) * TW_PT;
            const prevLeft = stack.length ? stack[stack.length - 1].left : 0;
            html += `<${tag} style="list-style-type:${lst};padding-left:${pt(Math.max(12, left - prevLeft))};margin-top:0;margin-bottom:0"${tag === 'ol' && n2 !== 1 ? ` start="${n2}"` : ''}>`;
            stack.push({ tag, ilvl: lv, numId: pr.li.numId, left });
          }
          html += `<li>${phtml.replace(/^<p /, '<p ').replace(/margin-bottom:([^;"]*)/, 'margin-bottom:$1')}`;
        } else {
          closeAll();
          html += phtml;
        }
        if (pr.boxes.length) { closeAll(); html += pr.boxes.join(''); }
        if (pr.pageBreak) { closeAll(); html += '<div class="pb" contenteditable="false"></div>'; }
        if (pr.sect && sections) { closeAll(); sections.push(pr.sect); const type = av(kid(pr.sect, 'type'), 'val'); if (type !== 'continuous') html += '<div class="pb" contenteditable="false"></div>'; }
      } else if (n.localName === 'tbl') { closeAll(); prevStyle = null; html += await table(n, relMap, b); }
      else if (n.localName === 'sectPr') { if (sections) sections.push(n); }
      else if (n.namespaceURI === OMML_NS && n.localName === 'oMathPara') { closeAll(); const tex = ommlToLatex(n); if (tex) html += `<p style="text-align:center">${mk(`q:${b64(tex)}`)}</p>`; }
      else if (n.localName === 'AlternateContent') { const ch = kid(n, 'Choice') || kid(n, 'Fallback'); if (ch) html += await blocks(ch, relMap, b); }
    }
    closeAll();
    return html;
  }

  const body = docXml.getElementsByTagNameNS(W, 'body')[0];
  const sections = [];
  let html = await blocks(body, docRels, 'word/', sections);
  // la última página no necesita salto
  html = html.replace(/(<div class="pb" contenteditable="false"><\/div>)+$/, '');

  // página (última sección)
  const sect = sections[sections.length - 1];
  let page = null; let header = ''; let footer = ''; let columns = 1;
  if (sect) {
    const sz = kid(sect, 'pgSz'); const mar = kid(sect, 'pgMar');
    const w = num(av(sz, 'w')), hh = num(av(sz, 'h'));
    if (w && hh) page = { w: Math.round(w * TWIP_PX), h: Math.round(hh * TWIP_PX), mt: Math.round(Math.abs(num(av(mar, 'top')) ?? 1440) * TWIP_PX), mb: Math.round(Math.abs(num(av(mar, 'bottom')) ?? 1440) * TWIP_PX), ml: Math.round((num(av(mar, 'left')) ?? 1440) * TWIP_PX), mr: Math.round((num(av(mar, 'right')) ?? 1440) * TWIP_PX), landscape: w > hh };
    columns = num(av(kid(sect, 'cols'), 'num')) || 1;
    // encabezado y pie predeterminados como texto, con {pagina} y {paginas}
    const hfText = async (kind) => {
      const ref = kids(sect, `${kind}Reference`).find((x) => av(x, 'type') === 'default') || kids(sect, `${kind}Reference`)[0];
      if (!ref) return '';
      const rel = docRels.get(avNs(ref, 'id')); if (!rel) return '';
      const d = await xml(`word/${rel.target}`); if (!d) return '';
      const lines = [];
      for (const p of d.getElementsByTagNameNS(W, 'p')) {
        let s = ''; let instr = ''; let state = 0; // 0 texto, 1 instrucción, 2 resultado reemplazado
        const walk = (el) => {
          for (const c of el.children) {
            const ln = c.localName;
            if (ln === 'fldSimple') { const i = av(c, 'instr') || ''; if (/\bNUMPAGES\b|\bSECTIONPAGES\b/.test(i)) s += '{paginas}'; else if (/\bPAGE\b/.test(i)) s += '{pagina}'; else walk(c); continue; }
            if (ln === 'fldChar') { const t = av(c, 'fldCharType'); if (t === 'begin') { state = 1; instr = ''; } else if (t === 'separate') { if (/\bNUMPAGES\b|\bSECTIONPAGES\b/.test(instr)) { s += '{paginas}'; state = 2; } else if (/\bPAGE\b/.test(instr)) { s += '{pagina}'; state = 2; } else state = 0; } else if (t === 'end') { if (state === 1 && /\bNUMPAGES\b/.test(instr)) s += '{paginas}'; else if (state === 1 && /\bPAGE\b/.test(instr)) s += '{pagina}'; state = 0; } continue; }
            if (ln === 'instrText') { instr += c.textContent; continue; }
            if (ln === 't') { if (state === 0) s += c.textContent; continue; }
            if (ln === 'tab' && el.localName === 'r') { if (state === 0) s += '   '; continue; }
            if (ln === 'pPr' || ln === 'rPr') continue;
            walk(c);
          }
        };
        walk(p);
        if (s.trim()) lines.push(s.trim());
      }
      return lines.join('\n');
    };
    header = await hfText('header'); footer = await hfText('footer');
  }
  // revisiones, comentarios y ecuaciones (marcas privadas → elementos del editor)
  const box = document.createElement('div');
  box.innerHTML = sanitize(html || '<p><br></p>');
  const comments = wrapMarks(box, rc?.list || []);
  box.querySelectorAll('p:empty,h1:empty,h2:empty,h3:empty,li:empty').forEach((p) => p.append(document.createElement('br')));
  let meta = null; try { const mt = await zip.file(META_DOCX)?.async('string'); if (mt) meta = JSON.parse(mt); } catch { /* */ }
  const { sources } = applyMeta(box, meta, comments, (tex) => { try { return temml.renderToString(tex, { annotate: true }); } catch { return null; } });
  return { html: box.innerHTML || '<p><br></p>', page, footnotes, endnotes, comments, sources, header, footer, columns, warnings: warnings.size ? [`No se pudieron conservar del todo: ${[...warnings].join(', ')}`] : [] };
}

function fmtNum(v, f) {
  if (f === 'lowerLetter' || f === 'upperLetter') { let s = ''; let n = v; while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26); } return f === 'upperLetter' ? s.toUpperCase() : s; }
  if (f === 'lowerRoman' || f === 'upperRoman') { const m = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]; let s = ''; let n = v; for (const [k, r] of m) while (n >= k) { s += r; n -= k; } return f === 'upperRoman' ? s.toUpperCase() : s; }
  if (f === 'bullet' || f === 'none') return '';
  if (f === 'decimalZero') return String(v).padStart(2, '0');
  return String(v);
}
