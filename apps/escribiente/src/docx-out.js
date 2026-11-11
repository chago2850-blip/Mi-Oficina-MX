// Convierte el documento de Escribiente (HTML del editor) a .docx (Office Open XML). GPL-3.0
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, ExternalHyperlink, ImageRun,
  Table, TableRow, TableCell, WidthType, ShadingType, PageBreak, Header, Footer, PageNumber,
  LevelFormat, PageOrientation, BorderStyle, TabStopType, LineRuleType, VerticalMergeType, VerticalAlignTable, TableLayoutType, AlignmentType as AT,
  CommentRangeStart, CommentRangeEnd, CommentReference, InsertedTextRun, DeletedTextRun, FootnoteReferenceRun, SectionType, Tab,
  SimpleField, Bookmark, ColumnBreak,
  Math as OMath, MathRun, MathFraction, MathSuperScript, MathSubScript, MathSubSuperScript, MathRadical, MathSum, MathIntegral,
} from 'docx';
import JSZip from 'jszip';
import { t } from '../../shared/i18n.js';
import { collectMeta, META_DOCX } from './meta.js';
import { LIST_PRESETS, CHAPTER_PRESETS, BUILTIN, BUILTIN_CHAR } from './estilos.js';

const PX_TWIP = 15; // 1 px (96 ppp) = 15 twips
const ptFromCss = (v) => {
  if (!v) return null;
  const n = parseFloat(v);
  if (Number.isNaN(n)) return null;
  if (v.endsWith('pt')) return n;
  if (v.endsWith('px')) return n * 0.75;
  if (v.endsWith('em')) return n * 11;
  const kw = { 'x-small': 7.5, small: 10, medium: 12, large: 13.5, 'x-large': 18, 'xx-large': 24, 'xxx-large': 36 }[v];
  return kw || null;
};
const hex = (c) => {
  if (!c || c === 'transparent' || c === 'inherit' || c === 'initial') return null;
  if (c.startsWith('#')) { let x = c.slice(1); if (x.length === 3) x = x.split('').map((y) => y + y).join(''); return x.slice(0, 6).toUpperCase(); }
  const m = c.match(/rgba?\(([^)]+)\)/);
  if (m) {
    const [r, g, b, a] = m[1].split(',').map((s) => parseFloat(s));
    if (a === 0) return null;
    return [r, g, b].map((n) => Math.round(n).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  return null;
};
// longitud CSS → px (96 ppp)
const lenPx = (v) => { if (!v) return 0; const n = parseFloat(v); if (Number.isNaN(n)) return 0; if (/pt$/.test(v)) return n / 0.75; if (/cm$/.test(v)) return (n * 96) / 2.54; if (/mm$/.test(v)) return (n * 96) / 25.4; if (/in$/.test(v)) return n * 96; return n; };
const BSTYLE = { solid: BorderStyle.SINGLE, double: BorderStyle.DOUBLE, dotted: BorderStyle.DOTTED, dashed: BorderStyle.DASHED, inset: BorderStyle.INSET, outset: BorderStyle.OUTSET || BorderStyle.SINGLE, groove: BorderStyle.SINGLE, ridge: BorderStyle.SINGLE };
// borde CSS (estilo, ancho y color calculados) → borde de Word
function cssBorder(style, width, color, spacePx = 0) {
  if (!style || style === 'none' || style === 'hidden' || !(parseFloat(width) > 0)) return null;
  const wpt = lenPx(width) * 0.75;
  return { style: BSTYLE[style] || BorderStyle.SINGLE, size: Math.max(2, Math.round((style === 'double' ? wpt / 3 : wpt) * 8)), color: hex(color) || '000000', space: Math.round(spacePx * 0.75) };
}
// formato de texto a partir de un estilo CSS (de un <span> o del propio párrafo)
function fmtFromStyle(st, f, el) {
  if (!st) return f;
  if (st.fontWeight === 'bold' || Number(st.fontWeight) >= 600) f.bold = true;
  if (st.fontWeight === 'normal' || st.fontWeight === '400') f.bold = false;
  if (st.fontStyle === 'italic') f.italics = true;
  if (st.fontStyle === 'normal') f.italics = false;
  const line = st.textDecorationLine || st.textDecoration || '';
  const dstyle = st.textDecorationStyle || (/(double|dotted|dashed|wavy)/.exec(st.textDecoration || '') || [])[1] || '';
  if (line.includes('underline')) { f.underline = true; f.underlineType = { double: 'double', dotted: 'dotted', dashed: 'dash', wavy: 'wave' }[dstyle] || f.underlineType; }
  if (line.includes('line-through')) { if (dstyle === 'double') f.dstrike = true; else f.strike = true; }
  if (line === 'none') { f.underline = false; f.strike = false; f.dstrike = false; }
  if (st.color && st.color !== 'transparent') f.color = hex(st.color);
  if (st.backgroundColor) f.shade = hex(st.backgroundColor);
  if (st.fontSize) f.size = ptFromCss(st.fontSize);
  if (st.fontFamily) f.font = firstFont(st.fontFamily);
  if (st.verticalAlign === 'super') f.superScript = true;
  if (st.verticalAlign === 'sub') f.subScript = true;
  if (/(pt|px)$/.test(st.verticalAlign || '')) f.position = Math.round(lenPx(st.verticalAlign) * 0.75 * 2);
  if (st.fontVariant === 'small-caps' || st.fontVariantCaps === 'small-caps') f.smallCaps = true;
  if (st.textTransform === 'uppercase') f.allCaps = true;
  if (st.textTransform === 'none') f.allCaps = false;
  if (st.letterSpacing && st.letterSpacing !== 'normal') f.spacing = Math.round(lenPx(st.letterSpacing) * 15);
  const sc = /scaleX\(([\d.]+)\)/.exec(st.transform || ''); if (sc) f.scale = Math.round(Number(sc[1]) * 100);
  if (st.textShadow && st.textShadow !== 'none') { if (/#fff/i.test(st.textShadow)) f[/^-1px/.test(st.textShadow) ? 'emboss' : 'imprint'] = true; else f.shadow = true; }
  if (st.webkitTextStroke && parseFloat(st.webkitTextStrokeWidth || st.webkitTextStroke) > 0) { f.outline = true; const m = /(#[0-9a-f]{3,8}|rgba?\([^)]*\))\s*$/i.exec(st.webkitTextStroke); if (m) f.color = hex(m[1]); }
  if (el?.classList?.contains('hidden-text')) f.vanish = true;
  return f;
}
const firstFont = (f) => (f ? f.split(',')[0].trim().replace(/^["']|["']$/g, '') : null);
const ALIGN = { center: AlignmentType.CENTER, right: AlignmentType.RIGHT, justify: AlignmentType.JUSTIFIED, left: AlignmentType.LEFT, start: AlignmentType.LEFT, end: AlignmentType.RIGHT };

function dataUrlToImage(src) {
  const m = /^data:image\/([a-z+]+);base64,(.*)$/i.exec(src || '');
  if (!m) return null;
  let type = m[1].toLowerCase();
  if (type === 'jpeg') type = 'jpg';
  const bytes = Uint8Array.from(atob(m[2]), (c) => c.charCodeAt(0));
  return { type, bytes };
}
async function rasterize(img) {
  // SVG/WebP no son compatibles con Word: se convierten a PNG.
  const c = document.createElement('canvas');
  c.width = img.naturalWidth || img.width; c.height = img.naturalHeight || img.height;
  c.getContext('2d').drawImage(img, 0, 0);
  const blob = await new Promise((r) => c.toBlob(r, 'image/png'));
  return { type: 'png', bytes: new Uint8Array(await blob.arrayBuffer()) };
}

class Builder {
  constructor(settings = {}) {
    this.olInstance = 0; this.settings = settings; this.listDefs = new Map();
    this.commentIds = new Map(); this.footIds = new Map(); this.endIds = new Map(); this.revId = 1;
  }
  numFor(map, key) { if (!map.has(key)) map.set(key, map.size + 1); return map.get(key); }

  // Recorre nodos en línea y produce TextRun / ImageRun / ExternalHyperlink
  async runs(node, fmt = {}, out = []) {
    for (const n of node.childNodes) {
      if (n.nodeType === 3) {
        if (this._skip?.has(n)) continue;
        const text = n.nodeValue.replace(/\u200B/g, '');
        if (!text) continue;
        if (fmt.rev) {
          const o = { text, id: this.revId++, author: fmt.rev.author || 'Escribiente', date: fmt.rev.date || new Date().toISOString(), ...this.runProps(fmt) };
          out.push(fmt.rev.kind === 'del' ? new DeletedTextRun(o) : new InsertedTextRun(o));
        } else if (text.includes('\t')) {
          text.split('\t').forEach((part, i) => { if (i) out.push(new TextRun({ children: [new Tab()], ...this.runProps(fmt) })); if (part) out.push(new TextRun({ text: part, ...this.runProps(fmt) })); });
        } else out.push(new TextRun({ text, ...this.runProps(fmt) }));
        continue;
      }
      if (n.nodeType !== 1) continue;
      if (this._skip?.has(n)) continue;
      const tag = n.tagName;
      if (tag === 'BR') { out.push(new TextRun({ break: 1, ...this.runProps(fmt) })); continue; }
      if (tag === 'IMG') { const r = await this.image(n); if (r) out.push(r); continue; }
      if (['UL', 'OL', 'TABLE'].includes(tag)) continue; // se manejan como bloques
      if (n.classList?.contains('eq')) { const m = n.querySelector('math'); if (m) out.push(new OMath({ children: mathKids(m) })); else out.push(new TextRun({ text: n.dataset.tex || '' })); continue; }
      if (tag === 'SUP' && n.classList.contains('en')) {
        const fid = n.dataset.fid;
        if (this.settings.endnotes?.[fid] != null) out.push(new TextRun({ text: `\uE010${this.numFor(this.endIds, fid)}\uE011` }));
        continue;
      }
      if (tag === 'SUP' && n.classList.contains('fn')) {
        const fid = n.dataset.fid;
        if (this.settings.footnotes?.[fid] != null) out.push(new FootnoteReferenceRun(this.numFor(this.footIds, fid)));
        continue;
      }
      // campos (fecha, página…), marcadores, entradas de índice y salto de columna
      if (n.classList?.contains('fld')) { out.push(...fieldRuns(n, this.runProps(fmt))); continue; }
      if (tag === 'A' && n.classList.contains('bm') && n.id) {
        // en un título (Figura 1…) el marcador abarca la categoría y el número, como en Word
        const cap = n.parentElement?.classList.contains('caption') && n.parentElement.querySelector('.fld[data-fld="seq"]');
        if (cap && !this._capDone?.has(cap)) {
          (this._capDone ||= new Set()).add(cap);
          const tmp = document.createElement('span'); let k = n.nextSibling; const take = [];
          while (k) { take.push(k); if (k === cap) break; k = k.nextSibling; }
          if (take.includes(cap)) { (this._skip ||= new Set()); take.forEach((x) => { this._skip.add(x); tmp.append(x.cloneNode(true)); }); out.push(new Bookmark({ id: n.id, children: await this.runs(tmp, fmt, []) })); continue; }
        }
        out.push(new Bookmark({ id: n.id, children: [] })); continue;
      }
      if (n.classList?.contains('ix')) { out.push(new SimpleField(`XE "${(n.dataset.ix || '').replace(/"/g, '')}"`, '')); continue; }
      if (n.classList?.contains('cbk')) { out.push(new ColumnBreak()); continue; }
      if (n.classList?.contains('mf')) { out.push(new TextRun({ text: `«${n.dataset.f}»`, ...this.runProps(fmt) })); continue; }
      const f = { ...fmt };
      if (tag === 'INS' || tag === 'DEL') f.rev = { kind: tag.toLowerCase(), author: n.dataset.author, date: n.dataset.date };
      if (n.classList?.contains('cm') && n.dataset.cid && this.settings.comments?.some((c) => c.id === n.dataset.cid)) {
        const id = this.numFor(this.commentIds, n.dataset.cid);
        out.push(new CommentRangeStart(id));
        await this.runs(n, f, out);
        out.push(new CommentRangeEnd(id), new TextRun({ children: [new CommentReference(id)] }));
        continue;
      }
      if (tag === 'B' || tag === 'STRONG') f.bold = true;
      if (tag === 'I' || tag === 'EM') f.italics = true;
      if (tag === 'U') f.underline = true;
      if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') f.strike = true;
      if (tag === 'SUP') f.superScript = true;
      if (tag === 'SUB') f.subScript = true;
      if (tag === 'FONT') { if (n.color) f.color = hex(n.color); if (n.face) f.font = firstFont(n.face); }
      fmtFromStyle(n.style, f, n);
      if (n.dataset?.cstyle) f.cstyle = n.dataset.cstyle;
      if (tag === 'A' && n.getAttribute('href')) {
        const inner = await this.runs(n, { ...f, link: true }, []);
        out.push(new ExternalHyperlink({ link: n.getAttribute('href'), children: inner }));
        continue;
      }
      await this.runs(n, f, out);
    }
    return out;
  }

  runProps(f) {
    const p = {};
    if (f.bold) p.bold = true;
    if (f.italics) p.italics = true;
    if (f.underline || f.link) p.underline = {};
    if (f.strike) p.strike = true;
    if (f.superScript) p.superScript = true;
    if (f.subScript) p.subScript = true;
    if (f.link) p.style = 'Hyperlink'; else if (f.cstyle) p.style = CHAR_ID[f.cstyle] || f.cstyle;
    if (f.color) p.color = f.color;
    if (f.size) p.size = Math.round(f.size * 2);
    if (f.font) p.font = f.font;
    if (f.shade) p.shading = { type: ShadingType.CLEAR, color: 'auto', fill: f.shade };
    if (f.underlineType && p.underline) p.underline = { type: f.underlineType };
    if (f.dstrike) { p.doubleStrike = true; delete p.strike; }
    if (f.smallCaps) p.smallCaps = true;
    if (f.allCaps) p.allCaps = true;
    if (f.vanish) p.vanish = true;
    if (f.spacing) p.characterSpacing = f.spacing;
    if (f.scale && f.scale !== 100) p.scale = f.scale;
    if (f.position) p.position = `${f.position / 2}pt`;
    if (f.emboss) p.emboss = true;
    if (f.imprint) p.imprint = true;
    return p;
  }

  async image(img) {
    let d = dataUrlToImage(img.getAttribute('src'));
    if (!d) return null;
    if (!['png', 'jpg', 'gif', 'bmp'].includes(d.type)) {
      const tmp = new Image(); tmp.src = img.src; await tmp.decode().catch(() => {});
      d = await rasterize(tmp);
    }
    const w = img.width || parseFloat(img.style.width) || img.naturalWidth || 300;
    const ratio = (img.naturalHeight || 1) / (img.naturalWidth || 1);
    const hgt = img.height || Math.round(w * ratio);
    return new ImageRun({ type: d.type, data: d.bytes, transformation: { width: Math.round(w), height: Math.round(hgt) } });
  }

  paraProps(el, extra = {}) {
    const st = el.style || {};
    const ds = el.dataset || {};
    const p = { ...extra };
    const al = st.textAlign || el.getAttribute?.('align');
    if (al && ALIGN[al]) p.alignment = ALIGN[al];
    const hasBL = st.borderLeftStyle && st.borderLeftStyle !== 'none';
    const ml = lenPx(st.marginLeft) || (!hasBL ? lenPx(st.paddingLeft) : 0);
    const mr = lenPx(st.marginRight);
    const ti = lenPx(st.textIndent);
    if (ml || ti || mr) p.indent = { ...(ml ? { left: Math.round(ml * PX_TWIP) } : {}), ...(mr ? { right: Math.round(mr * PX_TWIP) } : {}), ...(ti > 0 ? { firstLine: Math.round(ti * PX_TWIP) } : ti < 0 ? { hanging: Math.round(-ti * PX_TWIP) } : {}) };
    const tabs = parseTabs(ds.tabs);
    if (tabs.length) p.tabStops = tabs.map((t) => ({ type: { l: TabStopType.LEFT, c: TabStopType.CENTER, r: TabStopType.RIGHT, d: TabStopType.DECIMAL }[t.k], position: Math.round(t.x * PX_TWIP), ...(t.f ? { leader: { '.': 'dot', '-': 'hyphen', _: 'underscore', '·': 'middleDot' }[t.f] } : {}) }));
    const spacing = {};
    const mt = ds.push ? ds.mt0 : st.marginTop;
    if (mt) spacing.before = Math.round(lenPx(mt) * PX_TWIP);
    if (st.marginBottom) spacing.after = Math.round(lenPx(st.marginBottom) * PX_TWIP);
    const lhs = String(st.lineHeight || '');
    if (/(pt|px)$/.test(lhs)) { spacing.line = Math.round(lenPx(lhs) * PX_TWIP); spacing.lineRule = ds.lr === 'atLeast' ? LineRuleType.AT_LEAST : LineRuleType.EXACT; }
    else if (parseFloat(lhs)) { spacing.line = Math.round(240 * parseFloat(lhs)); spacing.lineRule = LineRuleType.AUTO; }
    if (Object.keys(spacing).length) p.spacing = spacing;
    if (ds.keepnext) p.keepNext = true;
    if (ds.keeplines) p.keepLines = true;
    if (ds.pbb) p.pageBreakBefore = true;
    if (ds.widow === '0') p.widowControl = false;
    const border = {};
    for (const s of ['top', 'right', 'bottom', 'left']) {
      const S = s[0].toUpperCase() + s.slice(1);
      const b = cssBorder(st[`border${S}Style`], st[`border${S}Width`], st[`border${S}Color`], lenPx(st[`padding${S}`]));
      if (b) border[s] = b;
    }
    if (Object.keys(border).length) p.border = border;
    const bg = hex(st.backgroundColor); if (bg) p.shading = { type: ShadingType.CLEAR, color: 'auto', fill: bg };
    return p;
  }

  async paragraph(el, extra = {}) {
    const tag = el.tagName;
    const props = this.paraProps(el, extra);
    if (/^H[1-6]$/.test(tag)) { props.heading = HeadingLevel[`HEADING_${tag[1]}`]; props.outlineLevel = Number(tag[1]) - 1; }
    if (/^H[1-3]$/.test(tag) && this.settings.chapnum && CHAPTER_PRESETS[this.settings.chapnum]?.fmts && !props.numbering) props.numbering = { reference: 'chapters', level: Number(tag[1]) - 1 };
    if (el.dataset?.style && this.settings.styles?.[el.dataset.style]) props.style = el.dataset.style;
    if (tag === 'P' && el.classList.contains('title')) props.heading = HeadingLevel.TITLE;
    if (tag === 'P' && el.classList.contains('caption') && !props.style) props.style = 'Caption';
    if (tag === 'P' && el.classList.contains('subtitle')) props.style = 'Subtitle';
    if (tag === 'BLOCKQUOTE') props.style = 'Quote';
    // la letra del propio párrafo (fuente, tamaño, color…) pasa a cada fragmento de texto
    const base = fmtFromStyle(el.style, el.classList?.contains('subtitle') ? { color: '6B6F80', size: 14 } : {});
    delete base.shade;
    const children = await this.runs(el, base);
    return new Paragraph({ ...props, children });
  }

  // formato de cada nivel de una lista (biblioteca de listas, o el tipo de la propia lista)
  listCtx(top) {
    const lst = (top.style.listStyleType || '').replace(/["']/g, '').trim();
    let key = top.dataset.ml && LIST_PRESETS[top.dataset.ml] ? top.dataset.ml : top.tagName === 'OL' ? `n:${lst || 'decimal'}` : `b:${lst && !/^(disc|circle|square)$/.test(lst) ? lst : { circle: '◦', square: '▪' }[lst] || '•'}`;
    const start = top.tagName === 'OL' ? Number(top.getAttribute('start')) || 1 : 1;
    const ref = `L${[...`${key}|${start}`].reduce((a, c) => (a * 31 + c.charCodeAt(0)) >>> 0, 7).toString(36)}`;
    if (!this.listDefs.has(ref)) this.listDefs.set(ref, levelsFor(key, start));
    return { ref, instance: top.tagName === 'OL' ? ++this.olInstance : 0, kind: top.tagName };
  }

  async list(el, level = 0, out = [], ctx = null) {
    if (!ctx) ctx = this.listCtx(el);
    for (const li of el.children) {
      // sublista suelta dentro de la lista (así la crea «Aumentar sangría»)
      if (li.tagName === 'OL' || li.tagName === 'UL') { await this.list(li, Math.min(level + 1, 8), out, li.tagName === ctx.kind ? ctx : this.listCtx(li)); continue; }
      if (li.tagName !== 'LI') continue;
      const clone = li.cloneNode(true);
      clone.querySelectorAll('ul,ol').forEach((x) => x.remove());
      const blocks = [...clone.children].filter((c) => /^(P|H[1-6]|DIV)$/.test(c.tagName));
      const numbering = { reference: ctx.ref, level, ...(ctx.instance ? { instance: ctx.instance } : {}) };
      if (blocks.length) for (const b of blocks) out.push(await this.paragraph(b, { numbering }));
      else out.push(new Paragraph({ ...this.paraProps(li), numbering, children: await this.runs(clone) }));
      for (const sub of li.children) if (sub.tagName === 'UL' || sub.tagName === 'OL') await this.list(sub, Math.min(level + 1, 8), out, sub.tagName === ctx.kind ? ctx : this.listCtx(sub));
    }
    return out;
  }

  async table(el, contentWidthPx) {
    const trs = [...el.querySelectorAll(':scope > tbody > tr, :scope > thead > tr, :scope > tr')].filter((r) => !/tpush|trep/.test(r.className));
    // cuadrícula: posición de cada celda contando las combinadas (colspan / rowspan)
    const occ = []; const placed = [];
    trs.forEach((tr, ri) => {
      occ[ri] ||= [];
      let ci = 0;
      for (const td of tr.children) {
        while (occ[ri][ci]) ci++;
        const cs = td.colSpan || 1; const rs = td.rowSpan || 1;
        placed.push({ td, ri, ci, cs, rs });
        for (let r = ri; r < ri + rs; r++) { occ[r] ||= []; for (let c = ci; c < ci + cs; c++) occ[r][c] = r === ri ? 'own' : { cont: td, first: c === ci, cs }; }
        ci += cs;
      }
    });
    const cols = Math.max(1, ...occ.map((r) => r.length));
    // anchos de columna: <col> o lo que mide en pantalla, escalado al ancho disponible
    const tblW = el.offsetWidth || contentWidthPx;
    let widths = [...el.querySelectorAll(':scope > colgroup > col')].map((c) => (/%$/.test(c.style.width) ? (parseFloat(c.style.width) / 100) * tblW : lenPx(c.style.width) || Number(c.getAttribute('width')) || 0));
    if (widths.length !== cols || widths.some((w) => !w)) {
      widths = Array(cols).fill(0);
      const r0 = placed.filter((x) => x.cs === 1);
      for (const x of r0) { const w = x.td.offsetWidth || 0; if (w && !widths[x.ci]) widths[x.ci] = w; }
      const known = widths.filter(Boolean); const avg = known.length ? known.reduce((a, b) => a + b, 0) / known.length : contentWidthPx / cols;
      widths = widths.map((w) => w || avg);
    }
    const total = widths.reduce((a, b) => a + b, 0) || contentWidthPx;
    const pctW = /%$/.test(el.style.width || '') ? parseFloat(el.style.width) / 100 : null;
    const target = pctW ? contentWidthPx * pctW : Math.min(total, contentWidthPx * 1.2);
    const colTw = widths.map((w) => Math.max(200, Math.round((w / total) * target * PX_TWIP)));
    const live = el.isConnected;
    const cellOpts = (td) => {
      const cs = live ? getComputedStyle(td) : td.style;
      const borders = {};
      for (const s of ['top', 'right', 'bottom', 'left']) {
        const S = s[0].toUpperCase() + s.slice(1);
        const b = cssBorder(cs[`border${S}Style`], cs[`border${S}Width`], cs[`border${S}Color`]);
        borders[s] = b || { style: BorderStyle.NIL, size: 0, color: 'auto' };
      }
      const bg = hex(cs.backgroundColor);
      const va = cs.verticalAlign === 'middle' ? VerticalAlignTable.CENTER : cs.verticalAlign === 'bottom' ? VerticalAlignTable.BOTTOM : undefined;
      const pad = (k) => Math.round(lenPx(cs[`padding${k}`]) * PX_TWIP);
      return { borders, shading: bg ? { type: ShadingType.CLEAR, color: 'auto', fill: bg } : undefined, verticalAlign: va, margins: { top: pad('Top'), bottom: pad('Bottom'), left: pad('Left'), right: pad('Right') } };
    };
    const rows = [];
    for (let ri = 0; ri < trs.length; ri++) {
      const tr = trs[ri];
      const cells = [];
      for (let ci = 0; ci < cols;) {
        const o = occ[ri]?.[ci];
        if (o && typeof o === 'object') {
          // celda que continúa una combinación vertical
          if (o.first) cells.push(new TableCell({ children: [new Paragraph({})], verticalMerge: VerticalMergeType.CONTINUE, columnSpan: o.cs > 1 ? o.cs : undefined, width: { size: colTw.slice(ci, ci + o.cs).reduce((a, b) => a + b, 0), type: WidthType.DXA }, ...cellOpts(o.cont) }));
          ci += o.cs; continue;
        }
        const pc = placed.find((x) => x.ri === ri && x.ci === ci);
        if (!pc) { cells.push(new TableCell({ children: [new Paragraph({})], width: { size: colTw[ci], type: WidthType.DXA } })); ci++; continue; }
        const { td, cs, rs } = pc;
        const kids = [];
        const blocks = [...td.children].filter((c) => /^(P|H[1-6]|DIV|UL|OL|BLOCKQUOTE|TABLE)$/.test(c.tagName));
        const wPx = widths.slice(ci, ci + cs).reduce((a, b) => a + b, 0);
        if (blocks.length) { for (const b of blocks) kids.push(...(await this.block(b, wPx))); }
        else kids.push(new Paragraph({ ...this.paraProps(td), children: await this.runs(td, td.tagName === 'TH' ? { bold: true } : {}) }));
        if (!kids.length || kids[kids.length - 1] instanceof Table) kids.push(new Paragraph({}));
        cells.push(new TableCell({
          children: kids,
          columnSpan: cs > 1 ? cs : undefined,
          verticalMerge: rs > 1 ? VerticalMergeType.RESTART : undefined,
          width: { size: colTw.slice(ci, ci + cs).reduce((a, b) => a + b, 0), type: WidthType.DXA },
          ...cellOpts(td),
        }));
        ci += cs;
      }
      const hgt = lenPx(tr.style.height);
      rows.push(new TableRow({ children: cells, tableHeader: tr.parentElement.tagName === 'THEAD' || [...tr.children].every((c) => c.tagName === 'TH') || undefined, height: hgt ? { value: Math.round(hgt * PX_TWIP), rule: 'atLeast' } : undefined, cantSplit: tr.dataset.keep === '1' || undefined }));
    }
    const ts = el.style;
    const center = ts.marginLeft === 'auto' && ts.marginRight === 'auto';
    const right = ts.marginLeft === 'auto' && !center;
    const none = { style: BorderStyle.NIL, size: 0, color: 'auto' };
    return new Table({
      rows, width: pctW ? { size: Math.round(pctW * 5000), type: WidthType.PERCENTAGE } : { size: colTw.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: colTw,
      layout: TableLayoutType.FIXED,
      alignment: center ? AT.CENTER : right ? AT.RIGHT : undefined,
      indent: !center && !right && lenPx(ts.marginLeft) ? { size: Math.round(lenPx(ts.marginLeft) * PX_TWIP), type: WidthType.DXA } : undefined,
      borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
    });
  }

  async block(el, contentWidthPx) {
    const tag = el.tagName;
    if (!/^(P|H[1-6]|DIV|UL|OL|TABLE|BLOCKQUOTE|HR|LI|PRE)$/.test(tag)) { const wrap = document.createElement('p'); wrap.append(el.cloneNode(true)); return [await this.paragraph(wrap)]; }
    if (el.classList?.contains('pb')) return [new Paragraph({ children: [new PageBreak()] })];
    if (el.classList?.contains('toc')) {
      const out = [];
      for (const l of el.querySelectorAll('.toc-h, .toc-l')) {
        if (l.classList.contains('toc-h')) out.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(l.textContent)] }));
        else {
          const [t, , pg] = l.children;
          const lvl = l.classList.contains('l3') ? 2 : l.classList.contains('l2') ? 1 : 0;
          out.push(new Paragraph({
            indent: { left: lvl * 280 },
            tabStops: [{ type: TabStopType.RIGHT, position: Math.round(contentWidthPx * PX_TWIP), leader: 'dot' }],
            children: [new TextRun(t?.textContent || ''), new TextRun({ text: `\t${pg?.textContent || ''}` })],
          }));
        }
      }
      return out;
    }
    if (tag === 'UL' || tag === 'OL') return this.list(el);
    if (el.classList?.contains('tbox')) {
      const bc = hex(el.style.borderColor) || '444444';
      const border = { style: BorderStyle.SINGLE, size: 6, color: bc, space: 4 };
      const bg = hex(el.style.backgroundColor);
      const out = [];
      const kids = [...el.children].filter((c) => /^(P|H[1-6]|DIV|BLOCKQUOTE)$/.test(c.tagName));
      for (const k of (kids.length ? kids : [el])) out.push(new Paragraph({ ...this.paraProps(k), border: { top: border, bottom: border, left: border, right: border }, shading: bg ? { type: ShadingType.CLEAR, color: 'auto', fill: bg } : undefined, children: await this.runs(k) }));
      return out;
    }
    if (tag === 'TABLE') return [await this.table(el, contentWidthPx)];
    if (tag === 'HR') return [new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '999999', space: 1 } }, children: [] })];
    if (tag === 'DIV' && [...el.children].some((c) => /^(P|H[1-6]|UL|OL|TABLE|DIV|BLOCKQUOTE)$/.test(c.tagName))) {
      const out = [];
      for (const c of el.children) out.push(...(await this.block(c, contentWidthPx)));
      return out;
    }
    return [await this.paragraph(el)];
  }
}

// settings: { page:{w,h,mt,mb,ml,mr,landscape}, header, footer, pageNumbers, title, author }
// campo de Escribiente → campo de Word, con el último resultado
const WFMT = (f = '') => f.replace(/a\b/, 'am/pm');
function fieldRuns(n, props) {
  const d = n.dataset; const txt = n.textContent;
  switch (d.fld) {
    case 'page': return [new TextRun({ children: [PageNumber.CURRENT], ...props })];
    case 'pages': return [new TextRun({ children: [PageNumber.TOTAL_PAGES], ...props })];
    case 'date': case 'time': return d.fixed ? [new TextRun({ text: txt, ...props })] : [new SimpleField(`${d.fld === 'date' ? 'DATE' : 'TIME'} \\@ "${WFMT(d.fmt || '')}"`, txt)];
    case 'title': return [new SimpleField('TITLE', txt)];
    case 'author': return [new SimpleField('AUTHOR', txt)];
    case 'file': return [new SimpleField('FILENAME', txt)];
    case 'words': return [new SimpleField('NUMWORDS', txt)];
    case 'seq': return [new SimpleField(`SEQ ${d.seq || 'Figura'} \\* ARABIC`, txt)];
    case 'ref': return [new SimpleField(`${d.show === 'page' ? 'PAGEREF' : 'REF'} ${d.ref}${d.show === 'number' ? ' \\n' : d.show === 'above' ? ' \\p' : ''} \\h`, txt)];
    default: return [new TextRun({ text: txt, ...props })];
  }
}
export async function htmlToDocx(editorEl, settings) {
  const b = new Builder(settings);
  const pg = settings.page;
  const contentW = pg.w - pg.ml - pg.mr;
  // Cada bloque en columnas (div.cols) se vuelve una sección continua con su número de columnas.
  const groups = [{ cols: 1, children: [] }];
  for (const el of editorEl.children) {
    if (el.nodeType !== 1) continue;
    if (el.classList.contains('cols')) {
      const n = Number(el.dataset.n) || 2;
      const kids = [];
      for (const c of el.children) kids.push(...(await b.block(c, contentW / n)));
      groups.push({ cols: n, children: kids.length ? kids : [new Paragraph({})] }, { cols: 1, children: [] });
    } else groups[groups.length - 1].children.push(...(await b.block(el, contentW)));
  }
  const sectionsData = groups.filter((g, i) => g.children.length || i === 0);
  if (!sectionsData[0].children.length) sectionsData[0].children.push(new Paragraph({}));

  // texto del encabezado/pie: {pagina} y {paginas} se vuelven campos de Word; cada línea, un párrafo
  const hs = settings.hfStyle || {};
  const hrun = { size: Math.round((hs.size || 10) * 2), color: hex(hs.color || '#555555'), bold: hs.bold || undefined, italics: hs.italic || undefined };
  const tokRuns = (line) => {
    const kids = [];
    for (const part of line.split(/(\{pagina\}|\{paginas\}|\{fecha\}|\{titulo\}|\{autor\}|\{archivo\}|\t)/i)) {
      if (!part) continue;
      if (/^\{pagina\}$/i.test(part)) kids.push(new TextRun({ children: [PageNumber.CURRENT], ...hrun }));
      else if (/^\{paginas\}$/i.test(part)) kids.push(new TextRun({ children: [PageNumber.TOTAL_PAGES], ...hrun }));
      else if (part === '\t') kids.push(new TextRun({ children: [new Tab()], ...hrun }));
      else if (/^\{(fecha|titulo|autor|archivo)\}$/i.test(part)) { const k = part.slice(1, -1).toLowerCase(); kids.push(new SimpleField({ fecha: 'DATE \\@ "dd/MM/yyyy"', titulo: 'TITLE', autor: 'AUTHOR', archivo: 'FILENAME' }[k], { fecha: new Date().toLocaleDateString('es-MX'), titulo: settings.title || '', autor: settings.author || '', archivo: '' }[k])); }
      else kids.push(new TextRun({ text: part, ...hrun }));
    }
    return kids;
  };
  let isHeader = true;
  const hfParas = (text, numPos, header = true) => {
    isHeader = header;
    const lines = text ? String(text).split('\n') : [];
    const out = lines.map((l, i) => {
      const kids = tokRuns(l);
      if (numPos && i === lines.length - 1) { kids.push(new TextRun({ text: '   ', size: 20 }), new TextRun({ children: [`${t('Página')} `, PageNumber.CURRENT, ` ${t('de')} `, PageNumber.TOTAL_PAGES], size: 20, color: '555555' })); }
      return new Paragraph({ alignment: numPos === 'right' ? AlignmentType.RIGHT : numPos === 'center' ? AlignmentType.CENTER : hs.align === 'center' ? AlignmentType.CENTER : hs.align === 'right' ? AlignmentType.RIGHT : AlignmentType.LEFT, children: kids, tabStops: [{ type: TabStopType.CENTER, position: Math.round(((pg.w - pg.ml - pg.mr) / 2) * PX_TWIP) }, { type: TabStopType.RIGHT, position: Math.round((pg.w - pg.ml - pg.mr) * PX_TWIP) }], ...(hs.line ? { border: { [isHeader ? 'bottom' : 'top']: { style: BorderStyle.SINGLE, size: 4, color: '555555', space: 1 } } } : {}) });
    });
    if (!lines.length && numPos) out.push(new Paragraph({ alignment: numPos === 'right' ? AlignmentType.RIGHT : numPos === 'center' ? AlignmentType.CENTER : AlignmentType.LEFT, children: [new TextRun({ children: [`${t('Página')} `, PageNumber.CURRENT, ` ${t('de')} `, PageNumber.TOTAL_PAGES], size: 20, color: '555555' })] }));
    return out;
  };
  const pn = settings.pageNumbers || 'none';
  const hNum = pn === 'header-right' ? 'right' : null; const fNum = pn === 'footer-center' ? 'center' : pn === 'footer-right' ? 'right' : null;
  const headerChildren = (settings.header || pn.startsWith('header')) ? hfParas(settings.header, hNum, true) : null;
  const footerChildren = (settings.footer || pn.startsWith('footer')) ? hfParas(settings.footer, fNum, false) : null;
  // primera página y páginas pares distintas
  const hdrs = {}; const ftrs = {};
  if (headerChildren) hdrs.default = new Header({ children: headerChildren }); if (footerChildren) ftrs.default = new Footer({ children: footerChildren });
  if (settings.hfFirst) { hdrs.first = new Header({ children: hfParas(settings.headerFirst || '', null, true) }); ftrs.first = new Footer({ children: hfParas(settings.footerFirst || '', null, false) }); hdrs.default ||= new Header({ children: [new Paragraph({})] }); ftrs.default ||= new Footer({ children: [new Paragraph({})] }); }
  if (settings.hfOddEven) { hdrs.even = new Header({ children: hfParas(settings.headerEven ?? settings.header ?? '', hNum, true) }); ftrs.even = new Footer({ children: hfParas(settings.footerEven ?? settings.footer ?? '', fNum, false) }); }
  const pbd = settings.pageBorder; const BS = { solid: BorderStyle.SINGLE, double: BorderStyle.DOUBLE, dashed: BorderStyle.DASHED, dotted: BorderStyle.DOTTED };
  const pageBorderSide = pbd ? { style: BS[pbd.style] || BorderStyle.SINGLE, size: Math.round(pbd.width * 8), color: hex(pbd.color), space: Math.min(31, Math.round((pbd.space ?? 18) * 0.75)) } : null;
  const ln = settings.lineNum;

  const bulletLevels = Array.from({ length: 9 }, (_, l) => ({
    level: l, format: LevelFormat.BULLET, text: ['•', '◦', '▪'][l % 3], alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 360 * (l + 1) + 360, hanging: 360 } } },
  }));
  const numLevels = Array.from({ length: 9 }, (_, l) => ({
    level: l, format: [LevelFormat.DECIMAL, LevelFormat.LOWER_LETTER, LevelFormat.LOWER_ROMAN][l % 3], text: `%${l + 1}.`, alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 360 * (l + 1) + 360, hanging: 360 } } },
  }));

  const comments = (settings.comments || []).filter((c) => b.commentIds.has(c.id)).map((c) => ({
    id: b.commentIds.get(c.id), author: c.author || 'Escribiente', date: new Date(c.date || Date.now()),
    children: [new Paragraph({ children: [new TextRun(c.text || '')] }), ...(c.replies || []).map((r) => new Paragraph({ children: [new TextRun({ text: `↪ ${r.author || ''}: `, bold: true }), new TextRun(r.text || '')] }))],
  }));
  const footnotes = {};
  for (const [fid, n] of b.footIds) footnotes[n] = { children: [new Paragraph({ children: [new TextRun({ text: settings.footnotes[fid] || '', size: 20 })] })] };
  const pageProps = (i) => ({
    page: {
      size: { width: Math.round(Math.min(pg.w, pg.h) * PX_TWIP), height: Math.round(Math.max(pg.w, pg.h) * PX_TWIP), orientation: pg.landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT },
      margin: { top: Math.round(pg.mt * PX_TWIP), bottom: Math.round(pg.mb * PX_TWIP), left: Math.round(pg.ml * PX_TWIP), right: Math.round(pg.mr * PX_TWIP), header: 708, footer: 708 },
      ...(pageBorderSide && i === 0 ? { borders: { pageBorderTop: pageBorderSide, pageBorderBottom: pageBorderSide, pageBorderLeft: pageBorderSide, pageBorderRight: pageBorderSide, pageBorders: { display: 'allPages', offsetFrom: 'text', zOrder: 'front' } } } : {}),
      ...(i === 0 && (settings.firstPageNum || 1) !== 1 ? { pageNumbers: { start: settings.firstPageNum } } : {}),
    },
    ...(ln?.on ? { lineNumbers: { countBy: ln.every || 1, restart: ln.restart === 'page' ? 'newPage' : 'continuous', distance: Math.round((ln.dist ?? 18) * PX_TWIP) } } : {}),
    ...(i === 0 && settings.hfFirst ? { titlePage: true } : {}),
    ...(i > 0 ? { type: SectionType.CONTINUOUS } : {}),
  });
  const doc = new Document({
    defaultTabStop: Math.round((settings.defTab || 47.24) * PX_TWIP),
    creator: settings.author || 'Escribiente',
    title: settings.title || '',
    subject: settings.subject || undefined,
    keywords: settings.keywords || undefined,
    lastModifiedBy: settings.author || undefined,
    description: 'Creado con Escribiente · Mi Oficina MX',
    styles: {
      default: { document: { run: { font: 'Calibri', size: 22 }, paragraph: { spacing: { after: 160, line: 276, lineRule: 'auto' } } } },
      paragraphStyles: [
        { id: 'Subtitle', name: 'Subtitle', basedOn: 'Normal', next: 'Normal', run: { size: 28, color: '6B6F80' } },
        { id: 'Quote', name: 'Quote', basedOn: 'Normal', next: 'Normal', run: { italics: true, color: '444444' }, paragraph: { indent: { left: 720, right: 720 } } },
        ...docStyles(b, settings, 'p'),
      ],
      characterStyles: docStyles(b, settings, 'c'),
    },
    numbering: { config: [{ reference: 'bullets', levels: bulletLevels }, { reference: 'numbers', levels: numLevels }, ...[...b.listDefs].map(([reference, levels]) => ({ reference, levels })), ...(settings.chapnum && CHAPTER_PRESETS[settings.chapnum]?.fmts ? [{ reference: 'chapters', levels: chapterLevels(settings.chapnum) }] : [])] },
    comments: comments.length ? { children: comments } : undefined,
    footnotes: Object.keys(footnotes).length ? footnotes : undefined,
    background: settings.pageColor ? { color: hex(settings.pageColor) } : undefined,
    evenAndOddHeaderAndFooters: settings.hfOddEven || undefined,
    features: settings.track ? { trackRevisions: true } : undefined,
    sections: sectionsData.map((g, i) => ({
      properties: { ...pageProps(i), ...(g.cols > 1 ? { column: { count: g.cols, space: 708 } } : {}) },
      headers: i === 0 && Object.keys(hdrs).length ? hdrs : undefined,
      footers: i === 0 && Object.keys(ftrs).length ? ftrs : undefined,
      children: g.children,
    })),
  });
  const blob = await Packer.toBlob(doc);
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const withNotes = b.endIds.size ? await addEndnotes(bytes, [...b.endIds].map(([fid, k]) => [k, settings.endnotes[fid] || ''])) : bytes;
  const meta = collectMeta(editorEl, settings);
  if (!meta) return withNotes;
  const zip = await JSZip.loadAsync(withNotes);
  zip.file(META_DOCX, JSON.stringify(meta));
  let ct = await zip.file('[Content_Types].xml').async('string');
  if (!/Extension="json"/.test(ct)) ct = ct.replace(/(<Types[^>]*>)/, '$1<Default Extension="json" ContentType="application/json"/>');
  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

// docx (la biblioteca) no genera notas al final: se agregan al paquete después.
const xmlEsc = (t) => String(t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
async function addEndnotes(bytes, notes) {
  const zip = await JSZip.loadAsync(bytes);
  let doc = await zip.file('word/document.xml').async('string');
  const ref = (k) => `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="${k}"/></w:r>`;
  doc = doc.replace(/<w:r>(?:(?!<\/w:r>).)*?\uE010(\d+)\uE011(?:(?!<\/w:r>).)*?<\/w:r>/gs, (_, k) => ref(k));
  zip.file('word/document.xml', doc);
  const NSD = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const sep = '<w:endnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>';
  const body = notes.map(([k, t]) => `<w:endnote w:id="${k}"><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:vertAlign w:val="superscript"/><w:sz w:val="20"/></w:rPr><w:endnoteRef/></w:r><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve"> ${xmlEsc(t)}</w:t></w:r></w:p></w:endnote>`).join('');
  zip.file('word/endnotes.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes ${NSD}>${sep}${body}</w:endnotes>`);
  let rels = await zip.file('word/_rels/document.xml.rels').async('string');
  rels = rels.replace('</Relationships>', '<Relationship Id="rIdEndnotes1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', rels);
  let ct = await zip.file('[Content_Types].xml').async('string');
  if (!ct.includes('/word/endnotes.xml')) ct = ct.replace('</Types>', '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>');
  zip.file('[Content_Types].xml', ct);
  const set = zip.file('word/settings.xml');
  if (set) { let x = await set.async('string'); if (!x.includes('w:endnotePr')) x = x.replace(/(<w:settings[^>]*>)/, '$1<w:endnotePr><w:numFmt w:val="lowerRoman"/></w:endnotePr>'); zip.file('word/settings.xml', x); }
  return zip.generateAsync({ type: 'uint8array' });
}

// MathML (de temml) -> ecuación nativa de Word (OMML) con los componentes de la biblioteca docx
const NARY = /^[∑∏∐⋃⋂]$/;
function mathKids(el) {
  const out = [];
  const kids = [...el.children].filter((k) => !['annotation', 'annotation-xml', 'mspace', 'none'].includes(k.localName));
  for (let i = 0; i < kids.length; i++) {
    const k = kids[i];
    const base = k.children[0], a = k.children[1], b = k.children[2];
    switch (k.localName) {
      case 'semantics': case 'mrow': case 'mstyle': case 'mpadded': case 'mphantom': case 'menclose': case 'math': out.push(...mathKids(k)); break;
      case 'mfrac': out.push(new MathFraction({ numerator: mathKids(wrap(base)), denominator: mathKids(wrap(a)) })); break;
      case 'msqrt': out.push(new MathRadical({ children: mathKids(k) })); break;
      case 'mroot': out.push(new MathRadical({ children: mathKids(wrap(base)), degree: mathKids(wrap(a)) })); break;
      case 'msup': case 'mover': out.push(new MathSuperScript({ children: mathKids(wrap(base)), superScript: mathKids(wrap(a)) })); break;
      case 'msub': case 'munder': {
        if (NARY.test(base?.textContent || '') || /^∫/.test(base?.textContent || '')) { out.push(...nary(base, a, null, kids, i)); i = kids.length; break; }
        out.push(new MathSubScript({ children: mathKids(wrap(base)), subScript: mathKids(wrap(a)) })); break;
      }
      case 'msubsup': case 'munderover': {
        if (NARY.test(base?.textContent || '') || /^∫/.test(base?.textContent || '')) { out.push(...nary(base, a, b, kids, i)); i = kids.length; break; }
        out.push(new MathSubSuperScript({ children: mathKids(wrap(base)), subScript: mathKids(wrap(a)), superScript: mathKids(wrap(b)) })); break;
      }
      case 'mtable': [...k.querySelectorAll('mtr')].forEach((tr, ri) => { if (ri) out.push(new MathRun('; ')); [...tr.children].forEach((td, ci) => { if (ci) out.push(new MathRun(', ')); out.push(...mathKids(td)); }); }); break;
      default: { const tx = k.textContent; if (tx) out.push(new MathRun(tx)); }
    }
  }
  return out;
  function wrap(n) { if (!n) { const e = document.createElementNS('http://www.w3.org/1998/Math/MathML', 'mrow'); return e; } if (n.localName === 'mrow') return n; const r = document.createElementNS('http://www.w3.org/1998/Math/MathML', 'mrow'); r.append(n.cloneNode(true)); return r; }
  // Suma o integral con límites: el resto de la expresión es su argumento
  function nary(sym, lo, hi, all, i) {
    const rest = document.createElementNS('http://www.w3.org/1998/Math/MathML', 'mrow');
    all.slice(i + 1).forEach((x) => rest.append(x.cloneNode(true)));
    const Cls = /^∫/.test(sym.textContent) ? MathIntegral : MathSum;
    return [new Cls({ children: mathKids(rest), subScript: lo ? mathKids(wrap(lo)) : undefined, superScript: hi ? mathKids(wrap(hi)) : undefined })];
  }
}

export function parseTabs(v) {
  return String(v || '').split(',').map((t) => { const m = /^([lcrd]):(\d+(?:\.\d+)?)(?::(.))?$/.exec(t.trim()); return m ? { k: m[1], x: Number(m[2]), f: m[3] || '' } : null; }).filter(Boolean).sort((a, b) => a.x - b.x);
}

// ---------------------------------------------------------------- listas, capítulos y estilos para Word
const FMT = { decimal: LevelFormat.DECIMAL, 'lower-alpha': LevelFormat.LOWER_LETTER, 'upper-alpha': LevelFormat.UPPER_LETTER, 'lower-roman': LevelFormat.LOWER_ROMAN, 'upper-roman': LevelFormat.UPPER_ROMAN };
function levelsFor(key, start = 1) {
  const P = LIST_PRESETS[key];
  const ind = (l) => ({ paragraph: { indent: { left: 360 * (l + 1) + 360, hanging: 360 } } });
  if ((P && P.bullet) || key.startsWith('b:')) {
    const ch = P?.bullet || key.slice(2) || '•';
    return Array.from({ length: 9 }, (_, l) => ({ level: l, format: LevelFormat.BULLET, text: l === 0 ? ch : ['◦', '▪', '•'][(l - 1) % 3], alignment: AlignmentType.LEFT, style: ind(l) }));
  }
  const fmts = P ? P.levels : [key.slice(2), 'lower-alpha', 'lower-roman'];
  return Array.from({ length: 9 }, (_, l) => {
    const f = fmts[Math.min(l, fmts.length - 1)] || 'decimal';
    const text = P?.legal ? `${Array.from({ length: l + 1 }, (_, k) => `%${k + 1}`).join('.')}.` : `%${l + 1}${P?.suffix || '.'}`;
    return { level: l, format: FMT[f] || LevelFormat.DECIMAL, text, start: l === 0 ? start : 1, alignment: AlignmentType.LEFT, style: ind(l) };
  });
}
function chapterLevels(key) {
  const P = CHAPTER_PRESETS[key];
  return [0, 1, 2].map((l) => {
    const body = P.own ? `%${l + 1}` : Array.from({ length: l + 1 }, (_, k) => `%${k + 1}`).join(P.sep || '.');
    const text = `${l === 0 && P.prefix1 ? t(P.prefix1) : ''}${body}${P.after.trim()}`;
    return { level: l, format: FMT[P.fmts[l]] || LevelFormat.DECIMAL, text, alignment: AlignmentType.LEFT, suffix: 'space', style: { paragraph: { indent: { left: 0, hanging: 0 } } } };
  });
}
// estilos del documento (propios y cambios a los predeterminados) como estilos de Word
// la biblioteca docx ya trae un estilo de párrafo «Strong»: el de carácter va con otro identificador
const CHAR_ID = { Strong: 'StrongEmphasis' };
function docStyles(b, settings, kind) {
  const out = [];
  const st = settings.styles || {};
  const cssText = (o) => Object.entries(o || {}).map(([k, v]) => `${k}:${v}`).join(';');
  const builtinIds = kind === 'p' ? BUILTIN.map((x) => x.id) : BUILTIN_CHAR.map((x) => x.id);
  for (const s of Object.values(st)) {
    if ((s.kind || 'p') !== kind) continue;
    const bc = kind === 'c' && BUILTIN_CHAR.find((x) => x.id === s.id);
    const el = document.createElement(kind === 'p' ? 'p' : 'span'); el.style.cssText = cssText(bc ? { ...bc.css, ...s.css } : s.css);
    const run = b.runProps(fmtFromStyle(el.style, {}, el));
    const def = { id: kind === 'c' ? CHAR_ID[s.id] || s.id : s.id, name: s.name || s.id, run };
    if (kind === 'p') {
      const par = b.paraProps(el); delete par.style;
      Object.assign(def, { basedOn: builtinIds.includes(s.id) ? undefined : (s.parent || 'Normal'), next: s.next || 'Normal', paragraph: par, quickFormat: true });
      if (/^Heading[1-6]$/.test(s.id)) def.basedOn = 'Normal';
    } else if (s.parent) def.basedOn = s.parent;
    out.push(def);
  }
  // estilos de carácter predeterminados que se usan en el documento
  if (kind === 'p' && !out.some((d) => d.id === 'Caption')) out.push({ id: 'Caption', name: 'caption', basedOn: 'Normal', next: 'Normal', run: { italics: true, size: 20, color: '333333' }, paragraph: { spacing: { before: 60, after: 120 } } });
  if (kind === 'c') for (const bc of BUILTIN_CHAR) if (!st[bc.id]) { const el = document.createElement('span'); el.style.cssText = cssText(bc.css); out.push({ id: CHAR_ID[bc.id] || bc.id, name: bc.name, run: b.runProps(fmtFromStyle(el.style, {}, el)) }); }
  return out;
}
