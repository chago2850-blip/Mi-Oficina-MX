// Convierte el documento de Escribiente (HTML del editor) a .docx (Office Open XML). GPL-3.0
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, ExternalHyperlink, ImageRun,
  Table, TableRow, TableCell, WidthType, ShadingType, PageBreak, Header, Footer, PageNumber,
  LevelFormat, PageOrientation, BorderStyle, TabStopType,
} from 'docx';

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
  constructor() { this.olInstance = 0; }

  // Recorre nodos en línea y produce TextRun / ImageRun / ExternalHyperlink
  async runs(node, fmt = {}, out = []) {
    for (const n of node.childNodes) {
      if (n.nodeType === 3) {
        const text = n.nodeValue.replace(/​/g, '');
        if (text) out.push(new TextRun({ text, ...this.runProps(fmt) }));
        continue;
      }
      if (n.nodeType !== 1) continue;
      const tag = n.tagName;
      if (tag === 'BR') { out.push(new TextRun({ break: 1, ...this.runProps(fmt) })); continue; }
      if (tag === 'IMG') { const r = await this.image(n); if (r) out.push(r); continue; }
      if (['UL', 'OL', 'TABLE'].includes(tag)) continue; // se manejan como bloques
      const f = { ...fmt };
      if (tag === 'B' || tag === 'STRONG') f.bold = true;
      if (tag === 'I' || tag === 'EM') f.italics = true;
      if (tag === 'U') f.underline = true;
      if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') f.strike = true;
      if (tag === 'SUP') f.superScript = true;
      if (tag === 'SUB') f.subScript = true;
      if (tag === 'FONT') { if (n.color) f.color = hex(n.color); if (n.face) f.font = firstFont(n.face); }
      const st = n.style;
      if (st) {
        if (st.fontWeight === 'bold' || Number(st.fontWeight) >= 600) f.bold = true;
        if (st.fontWeight === 'normal' || st.fontWeight === '400') f.bold = false;
        if (st.fontStyle === 'italic') f.italics = true;
        if ((st.textDecorationLine || st.textDecoration || '').includes('underline')) f.underline = true;
        if ((st.textDecorationLine || st.textDecoration || '').includes('line-through')) f.strike = true;
        if (st.color) f.color = hex(st.color);
        if (st.backgroundColor) f.shade = hex(st.backgroundColor);
        if (st.fontSize) f.size = ptFromCss(st.fontSize);
        if (st.fontFamily) f.font = firstFont(st.fontFamily);
        if (st.verticalAlign === 'super') f.superScript = true;
        if (st.verticalAlign === 'sub') f.subScript = true;
      }
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
    if (f.link) p.style = 'Hyperlink';
    if (f.color) p.color = f.color;
    if (f.size) p.size = Math.round(f.size * 2);
    if (f.font) p.font = f.font;
    if (f.shade) p.shading = { type: ShadingType.CLEAR, color: 'auto', fill: f.shade };
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
    const p = { ...extra };
    const al = st.textAlign || el.getAttribute?.('align');
    if (al && ALIGN[al]) p.alignment = ALIGN[al];
    const ml = parseFloat(st.marginLeft) || parseFloat(st.paddingLeft) || 0;
    if (ml) p.indent = { left: Math.round(ml * PX_TWIP) };
    const lh = parseFloat(st.lineHeight);
    const spacing = {};
    if (lh && !String(st.lineHeight).match(/px|pt/)) spacing.line = Math.round(240 * lh);
    if (Object.keys(spacing).length) p.spacing = spacing;
    return p;
  }

  async paragraph(el, extra = {}) {
    const tag = el.tagName;
    const props = this.paraProps(el, extra);
    if (/^H[1-6]$/.test(tag)) props.heading = HeadingLevel[`HEADING_${tag[1]}`];
    if (tag === 'P' && el.classList.contains('title')) props.heading = HeadingLevel.TITLE;
    if (tag === 'P' && el.classList.contains('subtitle')) props.style = 'Subtitle';
    if (tag === 'BLOCKQUOTE') props.style = 'Quote';
    const children = await this.runs(el, el.classList?.contains('subtitle') ? { color: '6B6F80', size: 14 } : {});
    return new Paragraph({ ...props, children });
  }

  async list(el, level = 0, out = [], instance = null) {
    const ordered = el.tagName === 'OL';
    if (ordered && instance == null) instance = ++this.olInstance;
    for (const li of el.children) {
      if (li.tagName !== 'LI') continue;
      // Contenido directo del <li> (sin sublistas)
      const clone = li.cloneNode(true);
      clone.querySelectorAll('ul,ol').forEach((x) => x.remove());
      const blocks = [...clone.children].filter((c) => /^(P|H[1-6]|DIV)$/.test(c.tagName));
      const numbering = ordered ? { reference: 'numbers', level, instance } : { reference: 'bullets', level };
      if (blocks.length) for (const b of blocks) out.push(await this.paragraph(b, { numbering }));
      else out.push(new Paragraph({ ...this.paraProps(li), numbering, children: await this.runs(clone) }));
      for (const sub of li.children) if (sub.tagName === 'UL' || sub.tagName === 'OL') await this.list(sub, Math.min(level + 1, 8), out, sub.tagName === 'OL' ? (ordered ? instance : null) : null);
    }
    return out;
  }

  async table(el, contentWidthPx) {
    const rows = [];
    const trs = [...el.querySelectorAll(':scope > tbody > tr, :scope > thead > tr, :scope > tr')];
    const cols = Math.max(1, ...trs.map((tr) => [...tr.children].reduce((n, c) => n + (c.colSpan || 1), 0)));
    const colW = Math.round((contentWidthPx * PX_TWIP) / cols);
    for (const tr of trs) {
      const cells = [];
      for (const td of tr.children) {
        const kids = [];
        const blocks = [...td.children].filter((c) => /^(P|H[1-6]|DIV|UL|OL|BLOCKQUOTE)$/.test(c.tagName));
        if (blocks.length) { for (const b of blocks) kids.push(...(await this.block(b, contentWidthPx / cols))); }
        else kids.push(new Paragraph({ ...this.paraProps(td), children: await this.runs(td, td.tagName === 'TH' ? { bold: true } : {}) }));
        if (!kids.length) kids.push(new Paragraph({}));
        const bg = hex(td.style.backgroundColor) || (td.tagName === 'TH' ? 'EEECFB' : null);
        cells.push(new TableCell({
          children: kids,
          columnSpan: td.colSpan > 1 ? td.colSpan : undefined,
          width: { size: colW * (td.colSpan || 1), type: WidthType.DXA },
          shading: bg ? { type: ShadingType.CLEAR, color: 'auto', fill: bg } : undefined,
          margins: { top: 40, bottom: 40, left: 80, right: 80 },
        }));
      }
      rows.push(new TableRow({ children: cells, tableHeader: tr.parentElement.tagName === 'THEAD' || undefined }));
    }
    const border = { style: BorderStyle.SINGLE, size: 4, color: '9A9EAD' };
    return new Table({
      rows, width: { size: cols * colW, type: WidthType.DXA }, columnWidths: Array(cols).fill(colW),
      borders: { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border },
    });
  }

  async block(el, contentWidthPx) {
    const tag = el.tagName;
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
export async function htmlToDocx(editorEl, settings) {
  const b = new Builder();
  const pg = settings.page;
  const contentW = pg.w - pg.ml - pg.mr;
  const children = [];
  for (const el of editorEl.children) {
    if (el.nodeType !== 1) continue;
    children.push(...(await b.block(el, contentW)));
  }
  if (!children.length) children.push(new Paragraph({}));

  const hfPara = (text, numPos) => {
    const kids = [];
    if (text) kids.push(new TextRun({ text, size: 20, color: '555555' }));
    if (numPos) { if (text) kids.push(new TextRun({ text: '   ', size: 20 })); kids.push(new TextRun({ children: ['Página ', PageNumber.CURRENT, ' de ', PageNumber.TOTAL_PAGES], size: 20, color: '555555' })); }
    return new Paragraph({ alignment: numPos === 'right' ? AlignmentType.RIGHT : numPos === 'center' ? AlignmentType.CENTER : AlignmentType.LEFT, children: kids });
  };
  const pn = settings.pageNumbers || 'none';
  const headerChildren = (settings.header || pn.startsWith('header')) ? [hfPara(settings.header, pn === 'header-right' ? 'right' : null)] : null;
  const footerChildren = (settings.footer || pn.startsWith('footer')) ? [hfPara(settings.footer, pn === 'footer-center' ? 'center' : pn === 'footer-right' ? 'right' : null)] : null;

  const bulletLevels = Array.from({ length: 9 }, (_, l) => ({
    level: l, format: LevelFormat.BULLET, text: ['•', '◦', '▪'][l % 3], alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 360 * (l + 1) + 360, hanging: 360 } } },
  }));
  const numLevels = Array.from({ length: 9 }, (_, l) => ({
    level: l, format: [LevelFormat.DECIMAL, LevelFormat.LOWER_LETTER, LevelFormat.LOWER_ROMAN][l % 3], text: `%${l + 1}.`, alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 360 * (l + 1) + 360, hanging: 360 } } },
  }));

  const doc = new Document({
    creator: settings.author || 'Escribiente',
    title: settings.title || '',
    description: 'Creado con Escribiente · Mi Oficina MX',
    styles: {
      default: { document: { run: { font: 'Calibri', size: 22 }, paragraph: { spacing: { after: 160, line: 276 } } } },
      paragraphStyles: [
        { id: 'Subtitle', name: 'Subtitle', basedOn: 'Normal', next: 'Normal', run: { size: 28, color: '6B6F80' } },
        { id: 'Quote', name: 'Quote', basedOn: 'Normal', next: 'Normal', run: { italics: true, color: '444444' }, paragraph: { indent: { left: 720, right: 720 } } },
      ],
    },
    numbering: { config: [{ reference: 'bullets', levels: bulletLevels }, { reference: 'numbers', levels: numLevels }] },
    sections: [{
      properties: {
        page: {
          size: { width: Math.round(Math.min(pg.w, pg.h) * PX_TWIP), height: Math.round(Math.max(pg.w, pg.h) * PX_TWIP), orientation: pg.landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT },
          margin: { top: Math.round(pg.mt * PX_TWIP), bottom: Math.round(pg.mb * PX_TWIP), left: Math.round(pg.ml * PX_TWIP), right: Math.round(pg.mr * PX_TWIP), header: 708, footer: 708 },
        },
      },
      headers: headerChildren ? { default: new Header({ children: headerChildren }) } : undefined,
      footers: footerChildren ? { default: new Footer({ children: footerChildren }) } : undefined,
      children,
    }],
  });
  const blob = await Packer.toBlob(doc);
  return new Uint8Array(await blob.arrayBuffer());
}
