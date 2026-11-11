// Escribiente — Formato de texto enriquecido (.rtf): lectura y escritura, como en LibreOffice Writer.
// Párrafos con estilo (títulos, cita), alineación y sangrías, negrita, cursiva, subrayado, tachado,
// superíndice y subíndice, fuente, tamaño, color y resaltado, vínculos, campos, listas, tablas
// (con celdas combinadas), imágenes PNG y JPEG, notas al pie y al final, saltos de página y de
// columna, encabezado y pie, tamaño y márgenes de página y propiedades del documento. GPL-3.0
import { grid } from './tablas.js';

const TW = 15; // twips por píxel CSS (1440 / 96)
const rgbHex = (c) => {
  if (!c) return null;
  if (/^#[0-9a-f]{6}$/i.test(c)) return c.toLowerCase();
  if (/^#[0-9a-f]{3}$/i.test(c)) return `#${c.slice(1).split('').map((d) => d + d).join('')}`.toLowerCase();
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?/.exec(c); if (!m || m[4] === '0') return null;
  return `#${[m[1], m[2], m[3]].map((n) => Number(n).toString(16).padStart(2, '0')).join('')}`;
};

// ================================================================ escritura
export function htmlToRtf(root, settings = {}) {
  const fonts = ['Calibri', 'Symbol']; const colors = [];
  const fontIx = (name) => { const n = String(name || '').split(',')[0].replace(/["']/g, '').trim(); if (!n) return 0; let i = fonts.indexOf(n); if (i < 0) { fonts.push(n); i = fonts.length - 1; } return i; };
  const colorIx = (c) => { const hx = rgbHex(c); if (!hx) return 0; let i = colors.indexOf(hx); if (i < 0) { colors.push(hx); i = colors.length - 1; } return i + 1; };
  const esc = (s) => {
    let o = '';
    for (const ch of String(s)) {
      const cp = ch.codePointAt(0);
      if (ch === '\\' || ch === '{' || ch === '}') o += `\\${ch}`;
      else if (ch === '\t') o += '\\tab ';
      else if (ch === '\u00A0') o += '\\~';
      else if (ch === '\u00AD') o += '\\-';
      else if (ch === '\n') o += '\\line ';
      else if (cp < 128) o += ch;
      else if (cp > 0xFFFF) { const s2 = ch; for (let k = 0; k < s2.length; k++) { const u = s2.charCodeAt(k); o += `\\u${u > 32767 ? u - 65536 : u}?`; } }
      else o += `\\u${cp > 32767 ? cp - 65536 : cp}?`;
    }
    return o;
  };
  const fn = settings.footnotes || {}; const en = settings.endnotes || {};
  const field = (instr, result) => `{\\field{\\*\\fldinst ${instr}}{\\fldrslt ${result}}}`;
  const FLD = { page: 'PAGE', pages: 'NUMPAGES', title: 'TITLE', author: 'AUTHOR', file: 'FILENAME', words: 'NUMWORDS', date: 'DATE \\\\@ "dd/MM/yyyy"', time: 'TIME \\\\@ "HH:mm"' };
  const img = (el) => {
    const m = /^data:image\/(png|jpeg|jpg);base64,(.+)$/.exec(el.getAttribute('src') || '');
    let type; let b64;
    if (m) { type = m[1] === 'png' ? 'pngblip' : 'jpegblip'; b64 = m[2]; } else {
      try { const c = document.createElement('canvas'); c.width = el.naturalWidth || el.width; c.height = el.naturalHeight || el.height; c.getContext('2d').drawImage(el, 0, 0); type = 'pngblip'; b64 = c.toDataURL('image/png').split(',')[1]; } catch { return ''; }
    }
    const bin = atob(b64); let hex = '';
    for (let i = 0; i < bin.length; i++) { hex += bin.charCodeAt(i).toString(16).padStart(2, '0'); if (i % 64 === 63) hex += '\n'; }
    const w = Math.round((el.width || el.naturalWidth || 100) * TW); const hh = Math.round((el.height || el.naturalHeight || 100) * TW);
    return `{\\pict\\${type}\\picw${el.naturalWidth || el.width}\\pich${el.naturalHeight || el.height}\\picwgoal${w}\\pichgoal${hh}\n${hex}}`;
  };
  // texto con formato de un nodo
  const inline = (node) => {
    let o = '';
    for (const n of node.childNodes) {
      if (n.nodeType === 3) { o += esc(n.data); continue; }
      if (n.nodeType !== 1) continue;
      const tag = n.tagName;
      if (tag === 'BR') { o += '\\line '; continue; }
      if (tag === 'IMG') { o += img(n); continue; }
      if (n.classList.contains('cbk')) { o += '\\column '; continue; }
      if (tag === 'SUP' && n.classList.contains('fn')) { o += `{\\super\\chftn}{\\footnote\\pard\\plain\\fs20{\\super\\chftn} ${esc(fn[n.dataset.fid] || '')}}`; continue; }
      if (tag === 'SUP' && n.classList.contains('en')) { o += `{\\super\\chftn}{\\footnote\\ftnalt\\pard\\plain\\fs20{\\super\\chftn} ${esc(en[n.dataset.fid] || '')}}`; continue; }
      if (n.classList.contains('fld') && FLD[n.dataset.fld]) { o += field(FLD[n.dataset.fld], esc(n.textContent)); continue; }
      if (n.classList.contains('bm')) { if (n.id) o += `{\\*\\bkmkstart ${esc(n.id)}}{\\*\\bkmkend ${esc(n.id)}}`; o += inline(n); continue; }
      if (tag === 'A' && n.getAttribute('href')) { o += field(`HYPERLINK "${esc(n.getAttribute('href'))}"`, `{\\ul\\cf${colorIx('#0563c1')} ${inline(n)}}`); continue; }
      let pre = '';
      const st = n.style;
      if (tag === 'B' || tag === 'STRONG' || /^(bold|[6-9]00)$/.test(st.fontWeight)) pre += '\\b';
      if (tag === 'I' || tag === 'EM' || st.fontStyle === 'italic') pre += '\\i';
      const td = `${st.textDecoration} ${st.textDecorationLine}`;
      if (tag === 'U' || /underline/.test(td)) pre += '\\ul';
      if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL' || /line-through/.test(td)) pre += '\\strike';
      if (tag === 'SUP' || st.verticalAlign === 'super') pre += '\\super';
      if (tag === 'SUB' || st.verticalAlign === 'sub') pre += '\\sub';
      if (tag === 'CODE') pre += `\\f${fontIx('Courier New')}`;
      if (st.fontFamily) pre += `\\f${fontIx(st.fontFamily)}`;
      if (tag === 'FONT' && n.getAttribute('face')) pre += `\\f${fontIx(n.getAttribute('face'))}`;
      const fsz = st.fontSize; if (fsz) { const pt = /pt$/.test(fsz) ? parseFloat(fsz) : parseFloat(fsz) * 0.75; if (pt) pre += `\\fs${Math.round(pt * 2)}`; }
      const col = st.color || (tag === 'FONT' ? n.getAttribute('color') : ''); if (col && colorIx(col)) pre += `\\cf${colorIx(col)}`;
      const bg = st.backgroundColor || (tag === 'MARK' ? '#ffff00' : ''); if (bg && colorIx(bg)) pre += `\\highlight${colorIx(bg)}\\chcbpat${colorIx(bg)}`;
      o += pre ? `{${pre} ${inline(n)}}` : inline(n);
    }
    return o;
  };
  const paraProps = (el, { list, level = 0 } = {}) => {
    let p = '\\pard\\plain';
    const tag = el.tagName; const st = el.style;
    if (/^H[1-6]$/.test(tag)) { const l = Number(tag[1]); p += `\\s${l}\\outlinelevel${l - 1}\\b\\fs${[0, 32, 26, 24, 22, 22, 22][l]}\\sb240\\sa120\\keepn`; if (l <= 2) p += `\\cf${colorIx('#2f2496')}`; }
    else if (el.classList?.contains('title')) p += `\\s10\\fs56\\sa120`;
    else if (el.classList?.contains('subtitle')) p += `\\s11\\fs30\\cf${colorIx('#6b6f80')}\\sa120`;
    else if (tag === 'BLOCKQUOTE') p += `\\s12\\li720\\ri720\\i`;
    else if (tag === 'PRE') p += `\\f${fontIx('Courier New')}`;
    else p += '\\s0\\sa160';
    const al = st?.textAlign || el.getAttribute?.('align');
    p += { center: '\\qc', right: '\\qr', justify: '\\qj' }[al] || '';
    const px = (v) => (/pt$/.test(v) ? parseFloat(v) / 0.75 : parseFloat(v)) || 0;
    if (list) p += `\\fi-360\\li${720 + level * 360}`;
    else {
      if (st?.marginLeft || st?.paddingLeft) p += `\\li${Math.round((px(st.marginLeft) + px(st.paddingLeft)) * TW)}`;
      if (st?.textIndent) p += `\\fi${Math.round(px(st.textIndent) * TW)}`;
    }
    if (st?.marginTop) p += `\\sb${Math.round(px(st.marginTop) * TW)}`;
    if (st?.marginBottom) p += `\\sa${Math.round(px(st.marginBottom) * TW)}`;
    if (st?.lineHeight && /^[\d.]+$/.test(st.lineHeight)) p += `\\sl${Math.round(Number(st.lineHeight) * 240)}\\slmult1`;
    if (el.dataset?.pbb) p += '\\pagebb';
    if (st?.backgroundColor && colorIx(st.backgroundColor)) p += `\\cbpat${colorIx(st.backgroundColor)}`;
    return `${p} `;
  };
  const table = (tbl) => {
    const { cells, m } = gridOf(tbl);
    const ncols = Math.max(0, ...m.map((r) => r.length));
    const colW = Array.from({ length: ncols }, (_, c) => { const info = m.find((r) => r[c] && r[c].cs === 1)?.[c]; return Math.max(300, Math.round(((info?.td.getBoundingClientRect().width || info?.td.offsetWidth || 120) * TW))); });
    let o = '';
    m.forEach((row, ri) => {
      o += `\\trowd\\trgaph108\\trleft0${ri === 0 && tbl.querySelector('tr')?.querySelector('th') ? '\\trhdr' : ''}`;
      let x = 0; const seen = new Set(); const parts = [];
      for (let c = 0; c < ncols; c++) {
        const info = row[c]; if (!info) { x += colW[c]; o += `\\cellx${x}`; parts.push('\\pard\\intbl\\cell '); continue; }
        if (seen.has(info) ) continue; seen.add(info);
        const w = colW.slice(c, c + info.cs).reduce((a, b) => a + b, 0); x += w;
        const bd = '\\brdrs\\brdrw10';
        const bg = rgbHex(info.td.style.backgroundColor);
        o += `${info.rs > 1 && info.r === ri ? '\\clvmgf' : info.r !== ri ? '\\clvmrg' : ''}\\clvertalt\\clbrdrt${bd}\\clbrdrl${bd}\\clbrdrb${bd}\\clbrdrr${bd}${bg ? `\\clcbpat${colorIx(bg)}` : ''}\\cellx${x}`;
        if (info.r !== ri) { parts.push('\\pard\\intbl\\cell '); continue; }
        const blocks = [...info.td.childNodes].some((k) => k.nodeType === 1 && /^(P|H[1-6]|UL|OL|DIV|BLOCKQUOTE)$/.test(k.tagName));
        let content;
        if (blocks) content = [...info.td.children].map((b, k, all) => `${paraProps(b).replace('\\pard', '\\pard\\intbl')}${info.td.tagName === 'TH' ? '{\\b ' : '{'}${inline(b)}}${k < all.length - 1 ? '\\par ' : ''}`).join('');
        else content = `\\pard\\intbl\\plain ${info.td.tagName === 'TH' ? '{\\b ' : '{'}${inline(info.td)}}`;
        parts.push(`${content}\\cell `);
      }
      o += `\n${parts.join('')}\\row\n`;
    });
    return `${o}\\pard\\plain `;
  };
  const list = (el, level = 0) => {
    let o = ''; let n = Number(el.getAttribute('start')) || 1;
    for (const li of el.children) {
      if (li.tagName !== 'LI') continue;
      const sub = [...li.children].filter((k) => k.tagName === 'UL' || k.tagName === 'OL');
      const clone = li.cloneNode(true); clone.querySelectorAll(':scope > ul, :scope > ol').forEach((k) => k.remove());
      const marker = el.tagName === 'OL' ? `{\\pntext ${n}.\\tab}{\\*\\pn\\pnlvlbody\\pndec\\pnstart${n}\\pnindent360{\\pntxta .}}` : `{\\pntext\\f1\\'b7\\tab}{\\*\\pn\\pnlvlblt\\pnf1\\pnindent360{\\pntxtb\\'b7}}`;
      o += `${paraProps(li, { list: true, level })}${marker}${inline(clone)}\\par\n`;
      n += 1;
      sub.forEach((s) => { o += list(s, level + 1); });
    }
    return o;
  };
  const block = (el) => {
    if (el.nodeType === 3) return el.data.trim() ? `\\pard\\plain ${esc(el.data)}\\par\n` : '';
    if (el.nodeType !== 1) return '';
    const tag = el.tagName;
    if (el.classList.contains('pb')) return '\\page\n';
    if (tag === 'TABLE') return table(el);
    if (tag === 'UL' || tag === 'OL') return list(el);
    if (tag === 'HR') return `\\pard\\plain\\brdrb\\brdrs\\brdrw10\\brsp20 \\par\n`;
    if (tag === 'DIV' && (el.classList.contains('cols') || el.classList.contains('toc') || !el.className)) return [...el.childNodes].map(block).join('');
    if (tag === 'FIGURE') return [...el.childNodes].map(block).join('');
    if (tag === 'IMG') return `\\pard\\plain\\qc ${img(el)}\\par\n`;
    return `${paraProps(el)}${inline(el)}\\par\n`;
  };
  const body = [...root.childNodes].map(block).join('');
  const hf = (txt) => String(txt || '').split('\n').map((l) => l.replace(/\{pagina\}/gi, '\u0001').replace(/\{paginas\}/gi, '\u0002')).map((l) => {
    const parts = l.split('\t');
    const conv = (s) => esc(s).replace(/\u0001/g, field('PAGE', '1')).replace(/\u0002/g, field('NUMPAGES', '1'));
    return parts.length > 1 ? `\\pard\\plain\\tqc\\tx4680\\tqr\\tx9360 ${parts.map(conv).join('\\tab ')}\\par` : `\\pard\\plain ${conv(l)}\\par`;
  }).join('');
  let header = settings.header || ''; let footer = settings.footer || '';
  if (!footer && /footer/.test(settings.pageNumbers || '')) footer = settings.pageNumbers === 'footer-right' ? '\t\t{pagina}' : '\t{pagina}\t';
  if (!header && settings.pageNumbers === 'header-right') header = '\t\t{pagina}';
  const pg = settings.page || {};
  const info = `{\\info${settings.title ? `{\\title ${esc(settings.title)}}` : ''}${settings.author ? `{\\author ${esc(settings.author)}}` : ''}${settings.subject ? `{\\subject ${esc(settings.subject)}}` : ''}${settings.keywords ? `{\\keywords ${esc(settings.keywords)}}` : ''}${settings.description ? `{\\doccomm ${esc(settings.description)}}` : ''}}`;
  const doc = [
    `\\paperw${Math.round((pg.w || 816) * TW)}\\paperh${Math.round((pg.h || 1056) * TW)}\\margl${Math.round((pg.ml || 96) * TW)}\\margr${Math.round((pg.mr || 96) * TW)}\\margt${Math.round((pg.mt || 96) * TW)}\\margb${Math.round((pg.mb || 96) * TW)}${pg.landscape ? '\\landscape' : ''}`,
    Object.keys(en).length ? '\\fet2\\enddoc\\aendnotes' : '\\fet0',
    `\\sectd${settings.hfFirst ? '\\titlepg' : ''}${settings.firstPageNum > 1 ? `\\pgnstarts${settings.firstPageNum}\\pgnrestart` : ''}`,
    header ? `{\\header ${hf(header)}}` : '', footer ? `{\\footer ${hf(footer)}}` : '',
    settings.hfFirst ? `{\\headerf ${hf(settings.headerFirst)}}{\\footerf ${hf(settings.footerFirst)}}` : '',
  ].join('\n');
  const fonttbl = `{\\fonttbl${fonts.map((f, i) => `{\\f${i}\\f${f === 'Symbol' ? 'tech\\fcharset2' : 'swiss\\fcharset0'} ${esc(f)};}`).join('')}}`;
  const colortbl = `{\\colortbl;${colors.map((c) => `\\red${parseInt(c.slice(1, 3), 16)}\\green${parseInt(c.slice(3, 5), 16)}\\blue${parseInt(c.slice(5, 7), 16)};`).join('')}}`;
  const styles = `{\\stylesheet{\\s0\\sa160 Normal;}${[1, 2, 3, 4, 5, 6].map((l) => `{\\s${l}\\outlinelevel${l - 1}\\b\\sbasedon0\\snext0 heading ${l};}`).join('')}{\\s10\\fs56\\sbasedon0\\snext0 Title;}{\\s11\\fs30\\sbasedon0\\snext0 Subtitle;}{\\s12\\li720\\ri720\\i\\sbasedon0\\snext0 Quote;}}`;
  return `{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1\\deflang3082\n${fonttbl}\n${colortbl}\n${styles}\n${info}\n{\\*\\generator Escribiente - Mi Oficina MX;}\n${doc}\n${body}}`;
}
function gridOf(tbl) {
  try { return grid(tbl); } catch {
    const m = []; const cells = [];
    [...tbl.rows].forEach((tr, r) => { m[r] = []; [...tr.cells].forEach((td, c) => { const info = { td, r, c, cs: 1, rs: 1 }; cells.push(info); m[r][c] = info; }); });
    return { m, cells };
  }
}

// ================================================================ lectura
const SKIP = new Set(['fonttbl', 'colortbl', 'stylesheet', 'listtable', 'listoverridetable', 'rsidtbl', 'generator', 'xmlnstbl', 'themedata', 'colorschememapping',
  'latentstyles', 'datastore', 'mmathPr', 'pgdsctbl', 'revtbl', 'filetbl', 'fldinst', 'nonshppict', 'shprslt', 'sn', 'bkmkend', 'pntext', 'listtext', 'pn', 'pgdsc',
  'docvar', 'wgrffmtfilter', 'ftnsep', 'ftnsepc', 'aftnsep', 'aftnsepc', 'objdata', 'blipuid', 'passwordhash', 'template', 'company', 'operator', 'hlinkbase', 'xe', 'tc', 'falt', 'panose', 'fontemb', 'fontfile', 'userprops', 'propname', 'staticval', 'atnid', 'atnauthor', 'annotation', 'atrfstart', 'atrfend']);
const TRANSPARENT = new Set(['shp', 'shpinst', 'sp', 'sv', 'shppict', 'fldrslt', 'field', 'upr', 'ud', 'shptxt']);
const INFO = new Set(['info', 'title', 'author', 'subject', 'keywords', 'doccomm', 'category', 'creatim', 'revtim', 'printim', 'buptim', 'version', 'vern', 'edmins', 'nofpages', 'nofwords', 'nofchars', 'nofcharsws', 'id', 'manager', 'comment', 'hlinkbase']);

export function rtfToHtml(input) {
  const src = typeof input === 'string' ? input : new TextDecoder('latin1').decode(input);
  let decoder = new TextDecoder('windows-1252');
  const fonts = {}; const colors = []; const styles = {};
  const props = {}; const page = {}; const out = { header: '', footer: '', headerFirst: '', footerFirst: '', footnotes: {}, endnotes: {} };
  const escH = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  // ---------------------------------------------------------------- análisis léxico
  let i = 0; const n = src.length;
  const tokens = [];
  while (i < n) {
    const c = src[i];
    if (c === '{') { tokens.push({ t: '{' }); i++; continue; }
    if (c === '}') { tokens.push({ t: '}' }); i++; continue; }
    if (c === '\\') {
      const d = src[i + 1];
      if (/[a-zA-Z]/.test(d)) {
        let j = i + 1; while (j < n && /[a-zA-Z]/.test(src[j])) j++;
        const word = src.slice(i + 1, j);
        let k = j; while (k < n && /[-0-9]/.test(src[k]) && (k === j ? true : src[k] !== '-')) k++;
        const num = k > j ? Number(src.slice(j, k)) : null;
        if (src[k] === ' ') k++;
        tokens.push({ t: 'w', w: word, n: num }); i = k;
        if (word === 'bin' && num > 0) { tokens.push({ t: 'bin', v: src.slice(i, i + num) }); i += num; }
        continue;
      }
      if (d === "'") { tokens.push({ t: 'hex', v: parseInt(src.substr(i + 2, 2), 16) }); i += 4; continue; }
      if (d === '\r' || d === '\n') { tokens.push({ t: 'w', w: 'par', n: null }); i += 2; continue; }
      tokens.push({ t: 'sym', v: d }); i += 2; continue;
    }
    if (c === '\r' || c === '\n') { i++; continue; }
    let j = i; while (j < n && !'{}\\\r\n'.includes(src[j])) j++;
    tokens.push({ t: 'txt', v: src.slice(i, j) }); i = j;
  }
  // ---------------------------------------------------------------- interpretación
  const blocks = []; // { type:'p'|'pb'|'table', ... }
  const defChar = () => ({ b: 0, i: 0, ul: 0, strike: 0, sup: 0, sub: 0, f: null, fs: 0, cf: 0, hl: 0, uc: 1, link: null });
  const defPara = () => ({ s: 0, al: '', li: 0, fi: 0, ri: 0, intbl: 0, pagebb: 0, list: null, outline: null });
  const root = { ch: defChar(), dest: 'body', skip: 0, target: null };
  let st = root; const stack = [];
  let para = defPara();
  let runs = [];
  let row = null; let rowDefs = []; let cellDefs = []; let table = null; let cellParas = []; let pend = {};
  let ucSkip = 0;
  const sink = () => (st.target ? st.target.runs : runs);
  const pushText = (txt) => { if (txt) sink().push({ text: txt, ch: { ...st.ch } }); };
  const pushHtml = (html) => { sink().push({ html, ch: { ...st.ch } }); };
  const runHtml = (rs) => {
    let o = '';
    for (const r of rs) {
      const c = r.ch; let s = r.html ?? escH(r.text);
      const css = [];
      if (c.f != null && fonts[c.f] && !/^(Calibri|Carlito|Times New Roman|Liberation Serif|Liberation Sans|Arial|Symbol)$/.test(fonts[c.f])) css.push(`font-family:'${fonts[c.f].replace(/'/g, '')}'`);
      if (c.fs && c.fs !== 24 && c.fs !== 22) css.push(`font-size:${c.fs / 2}pt`);
      if (c.cf && colors[c.cf] && colors[c.cf] !== '#000000' && !(c.link && colors[c.cf] === '#0563c1')) css.push(`color:${colors[c.cf]}`);
      if (c.hl && colors[c.hl] && colors[c.hl] !== '#ffffff') css.push(`background-color:${colors[c.hl]}`);
      if (css.length) s = `<span style="${css.join(';')}">${s}</span>`;
      if (c.sub) s = `<sub>${s}</sub>`; if (c.sup) s = `<sup>${s}</sup>`;
      if (c.strike) s = `<s>${s}</s>`; if (c.ul && !c.link) s = `<u>${s}</u>`; if (c.i) s = `<i>${s}</i>`; if (c.b) s = `<b>${s}</b>`;
      if (c.link) s = `<a href="${escH(c.link)}">${s}</a>`;
      o += s;
    }
    return o.replace(/<\/a><a href="[^"]*">/g, '');
  };
  const tagFor = (p) => {
    const nm = (styles[p.s] || '').toLowerCase();
    let lvl = p.outline != null && p.outline < 6 ? p.outline + 1 : null;
    const hm = /^(?:heading|t[ií]tulo|encabezado)\s*(\d)$/.exec(nm); if (hm) lvl = Number(hm[1]);
    if (lvl) return `h${Math.min(6, lvl)}`;
    if (/^(title|t[ií]tulo)$/.test(nm)) return 'p class="title"';
    if (/^(subtitle|subt[ií]tulo)$/.test(nm)) return 'p class="subtitle"';
    if (/^(quote|quotations|cita|citas)$/.test(nm)) return 'blockquote';
    return 'p';
  };
  const paraCss = (p) => {
    const css = [];
    if (p.al) css.push(`text-align:${p.al}`);
    if (!p.list && p.li && !/^(quote|quotations|cita|citas)$/i.test(styles[p.s] || '')) css.push(`margin-left:${Math.round(p.li / TW)}px`);
    if (!p.list && p.fi) css.push(`text-indent:${Math.round(p.fi / TW)}px`);
    return css.length ? ` style="${css.join(';')}"` : '';
  };
  const endPara = () => {
    const html = runHtml(runs); runs = [];
    if (para.intbl) { cellParas.push({ ...para, html }); return; }
    table = null;
    blocks.push({ type: 'p', ...para, html });
  };
  const endCell = () => {
    if (runs.length) { const html = runHtml(runs); runs = []; cellParas.push({ ...para, html }); }
    (row ||= []).push(cellParas); cellParas = [];
  };
  const endRow = () => {
    if (runs.length || cellParas.length) endCell();
    const defs = cellDefs.length ? cellDefs : rowDefs; rowDefs = defs;
    if (!table) { table = { type: 'table', rows: [] }; blocks.push(table); }
    table.rows.push({ cells: row || [], defs: [...defs], hdr: pend.hdr });
    row = null; cellDefs = []; pend = {};
  };
  const STAR_OK = new Set(['fldinst', 'shppict', 'shpinst', 'bkmkstart', 'footnote', 'pn']);
  const groupClosed = (child) => {
    if (!child.opened) return;
    switch (child.dest) {
      case 'fldinst': if (st.isField) { const tx = child.fldText || ''; const m = /HYPERLINK\s+(\\l\s+)?"([^"]+)"/i.exec(tx); if (m) st.fieldLink = m[1] ? `#${m[2]}` : m[2]; else st.fieldKind = (/^\s*([A-Z]+)/i.exec(tx) || [])[1]?.toUpperCase(); } break;
      case 'pict': {
        const pc = child.pict; if (!pc.type || !pc.hex) break;
        const hx = pc.hex.replace(/[^0-9a-f]/gi, ''); const u8 = new Uint8Array(hx.length >> 1);
        for (let q = 0; q < u8.length; q++) u8[q] = parseInt(hx.substr(q * 2, 2), 16);
        let bin = ''; for (let q = 0; q < u8.length; q += 0x8000) bin += String.fromCharCode.apply(null, u8.subarray(q, q + 0x8000));
        const w = pc.wg ? Math.round((pc.wg * (pc.sx || 100)) / 100 / TW) : null; const hh = pc.hg ? Math.round((pc.hg * (pc.sy || 100)) / 100 / TW) : null;
        pushHtml(`<img src="data:image/${pc.type};base64,${btoa(bin)}"${w ? ` width="${w}"` : ''}${hh ? ` height="${hh}"` : ''}>`);
        break;
      }
      case 'note': {
        const nt = child.note; if (nt.runs.length) nt.paras.push(runHtml(nt.runs));
        const txt = nt.paras.map((x) => x.replace(/<[^>]+>/g, '')).join(' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').trim();
        const fid = `r${Math.random().toString(36).slice(2, 9)}`;
        (nt.endnote ? out.endnotes : out.footnotes)[fid] = txt;
        pushHtml(`<sup class="${nt.endnote ? 'en' : 'fn'}" data-fid="${fid}" contenteditable="false">#</sup>`);
        break;
      }
      case 'hf': {
        const hf = child.hf; if (hf.runs.length) hf.paras.push(runHtml(hf.runs));
        out[hf.kind] = hf.paras.map((x) => x.replace(/<[^>]+>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')).join('\n').replace(/^\n+|\n+$/g, '');
        break;
      }
      case 'bkmk': if (child.cap?.trim()) pushHtml(`<a class="bm" id="${escH(child.cap.trim())}"></a>`); break;
      case 'listtext': { const tx = (child.cap || '').trim(); if (para.list?.bullet == null) para.list = { bullet: !/^[0-9]+[.)]?$|^[a-zA-Z]{1,4}[.)]$/.test(tx) }; break; }
      default:
    }
  };
  const SPECIAL = { emdash: '—', endash: '–', bullet: '•', lquote: '‘', rquote: '’', ldblquote: '“', rdblquote: '”', emspace: ' ', enspace: ' ', qmspace: ' ' };
  for (let t = 0; t < tokens.length; t++) {
    const tk = tokens[t];
    if (tk.t === '{') { stack.push(st); st = { ...st, ch: { ...st.ch }, star: false, isField: false, fieldLink: null, opened: false }; continue; }
    if (tk.t === '}') { const child = st; st = stack.pop() || root; groupClosed(child); continue; }
    if (tk.t === 'w' && st.star) {
      st.star = false;
      if (!STAR_OK.has(tk.w)) { st.skip = 1; st.dest = tk.w; continue; }
    }
    if (st.skip) {
      if (st.dest === 'fldinst') { if (tk.t === 'txt') st.fldText += tk.v; else if (tk.t === 'w') st.fldText += `\\${tk.w} `; }
      continue;
    }
    if (tk.t === 'w') {
      const { w, n: num } = tk; const on = num == null ? 1 : num ? 1 : 0;
      if (w === 'fonttbl' || w === 'colortbl' || w === 'stylesheet') { st.dest = w; st.col = [0, 0, 0, false]; continue; }
      if (st.dest === 'fonttbl') { if (w === 'f') st.fontRef = num; continue; }
      if (st.dest === 'colortbl') { if (w === 'red') st.col[0] = num; if (w === 'green') st.col[1] = num; if (w === 'blue') st.col[2] = num; if (/^(red|green|blue)$/.test(w)) st.col[3] = true; continue; }
      if (st.dest === 'stylesheet') { if (w === 's') st.styleRef = num; if (/^(cs|ds|ts)$/.test(w)) st.styleRef = -1; continue; }
      if (w === 'info') { st.dest = 'info'; continue; }
      if (st.dest === 'info' || st.dest === 'info-item') { if (INFO.has(w)) { st.dest = 'info-item'; st.infoKey = w; } continue; }
      if (w === 'field') { st.isField = true; continue; }
      if (w === 'fldinst') { st.dest = 'fldinst'; st.skip = 1; st.fldText = ''; st.opened = true; continue; }
      if (w === 'fldrslt') { const f = stack[stack.length - 1]; if (f?.fieldLink) st.ch.link = f.fieldLink; else if (f?.fieldKind) st.fieldKind = f.fieldKind; continue; }
      if (w === 'pict') { st.dest = 'pict'; st.pict = { hex: '' }; st.opened = true; continue; }
      if (st.dest === 'pict') {
        const pc = st.pict;
        if (w === 'pngblip') pc.type = 'png'; else if (w === 'jpegblip') pc.type = 'jpeg'; else if (/^(wmetafile|emfblip|macpict|dibitmap|wbitmap|pmmetafile)$/.test(w)) pc.type = null;
        if (w === 'picwgoal') pc.wg = num; if (w === 'pichgoal') pc.hg = num; if (w === 'picscalex') pc.sx = num; if (w === 'picscaley') pc.sy = num;
        continue;
      }
      if (w === 'footnote') { st.dest = 'note'; st.note = { runs: [], paras: [], endnote: false }; st.target = st.note; st.opened = true; continue; }
      if (w === 'ftnalt' && st.dest === 'note') { st.note.endnote = true; continue; }
      if (/^(header|footer)[flr]?$/.test(w)) { const m = /^(header|footer)([flr]?)$/.exec(w); const kind = m[2] === 'f' ? `${m[1]}First` : m[2] === 'l' ? `${m[1]}Even` : m[1]; st.dest = 'hf'; st.hf = { kind, runs: [], paras: [] }; st.target = st.hf; st.opened = true; continue; }
      if (w === 'bkmkstart') { st.dest = 'bkmk'; st.capture = true; st.cap = ''; st.opened = true; continue; }
      if (w === 'listtext' || w === 'pntext') { st.dest = 'listtext'; st.capture = true; st.cap = ''; st.opened = true; continue; }
      if (w === 'shp' || w === 'shpinst') { st.noText = true; continue; }
      if (SKIP.has(w)) { st.skip = 1; st.dest = w; continue; }
      if (TRANSPARENT.has(w)) continue;
      if (w === 'ansicpg' && num) { try { decoder = new TextDecoder(`windows-${num}`); } catch { /* sin cambio */ } continue; }
      const pageKw = { paperw: 'w', paperh: 'h', margl: 'ml', margr: 'mr', margt: 'mt', margb: 'mb' }[w];
      if (pageKw) { page[pageKw] = Math.round(num / TW); continue; }
      if (w === 'landscape') { page.landscape = true; continue; }
      if (w === 'titlepg') { out.hfFirst = true; continue; }
      if (w === 'pgnstarts') { out.firstPageNum = num; continue; }
      if (st.capture) continue;
      if (w === 'par' || w === 'sect') { if (st.target) { st.target.paras.push(runHtml(st.target.runs)); st.target.runs = []; } else endPara(); continue; }
      if (w === 'line') { pushHtml('<br>'); continue; }
      if (w === 'tab') { pushText('\t'); continue; }
      if (w === 'page') { if (runs.length) endPara(); table = null; blocks.push({ type: 'pb' }); continue; }
      if (w === 'column') { pushHtml('<span class="cbk" contenteditable="false"></span>'); continue; }
      if (w === 'u') { pushText(String.fromCharCode(num < 0 ? num + 65536 : num)); ucSkip = st.ch.uc; continue; }
      if (w === 'uc') { st.ch.uc = num; continue; }
      if (SPECIAL[w]) { pushText(SPECIAL[w]); continue; }
      if (w === 'pard') { para = defPara(); continue; }
      if (w === 's') { para.s = num; continue; }
      const al = { ql: '', qc: 'center', qr: 'right', qj: 'justify' }[w]; if (al !== undefined) { para.al = al; continue; }
      if (w === 'li' || w === 'lin') { para.li = num; continue; } if (w === 'fi') { para.fi = num; continue; } if (w === 'ri' || w === 'rin') { para.ri = num; continue; }
      if (w === 'intbl') { para.intbl = 1; continue; } if (w === 'outlinelevel') { para.outline = num; continue; } if (w === 'pagebb') { para.pagebb = 1; continue; }
      if (w === 'ls') { para.list ||= {}; continue; }
      if (w === 'pnlvlblt') { para.list = { bullet: true }; continue; } if (w === 'pnlvlbody' || w === 'pndec') { para.list = { bullet: false }; continue; }
      if (w === 'trowd') { cellDefs = []; pend = {}; continue; }
      if (w === 'trhdr') { pend.hdr = true; continue; }
      if (w === 'clvmgf') { pend.vm = 'first'; continue; } if (w === 'clvmrg') { pend.vm = 'cont'; continue; }
      if (w === 'clmgf') { pend.hm = 'first'; continue; } if (w === 'clmrg') { pend.hm = 'cont'; continue; }
      if (w === 'clcbpat') { pend.bg = num; continue; }
      if (w === 'cellx') { cellDefs.push({ x: num, vm: pend.vm, hm: pend.hm, bg: pend.bg }); delete pend.vm; delete pend.hm; delete pend.bg; continue; }
      if (w === 'cell') { endCell(); continue; }
      if (w === 'row') { endRow(); continue; }
      const c = st.ch;
      if (w === 'plain') { st.ch = { ...defChar(), uc: c.uc, link: c.link }; continue; }
      if (w === 'b') c.b = on; else if (w === 'i') c.i = on; else if (/^ul(d|db|w|dash|th|wave)?$/.test(w)) c.ul = on; else if (w === 'ulnone') c.ul = 0;
      else if (w === 'strike' || w === 'striked') c.strike = on; else if (w === 'super') { c.sup = on; c.sub = 0; } else if (w === 'sub') { c.sub = on; c.sup = 0; } else if (w === 'nosupersub') { c.sup = 0; c.sub = 0; }
      else if (w === 'f') c.f = num; else if (w === 'fs') c.fs = num; else if (w === 'cf') c.cf = num; else if (w === 'highlight' || w === 'chcbpat' || w === 'cb') c.hl = num;
      continue;
    }
    if (tk.t === 'sym') {
      const v = tk.v;
      if (v === '*') { st.star = true; continue; }
      if (st.dest === 'fonttbl' || st.dest === 'colortbl' || st.dest === 'stylesheet' || st.noText) continue;
      const ch = { '~': ' ', '-': '­', _: '‑', '\\': '\\', '{': '{', '}': '}' }[v];
      if (ch) { if (st.capture) st.cap += ch; else pushText(ch); }
      continue;
    }
    if (tk.t === 'hex' || tk.t === 'txt') {
      let v = tk.t === 'hex' ? null : tk.v;
      if (tk.t === 'hex') { if (ucSkip > 0) { ucSkip--; continue; } v = decoder.decode(new Uint8Array([tk.v])); }
      else if (ucSkip > 0) { const k = Math.min(ucSkip, v.length); v = v.slice(k); ucSkip -= k; }
      if (!v) continue;
      if (st.dest === 'pict') { if (tk.t === 'txt') st.pict.hex += v; continue; }
      if (st.dest === 'fonttbl') { const full = (st.fontName || '') + v; if (full.includes(';')) { if (st.fontRef != null) fonts[st.fontRef] = full.split(';')[0].trim(); st.fontName = ''; } else st.fontName = full; continue; }
      if (st.dest === 'colortbl') { for (const ch of v) if (ch === ';') { colors.push(st.col[3] ? `#${st.col.slice(0, 3).map((q) => q.toString(16).padStart(2, '0')).join('')}` : null); st.col = [0, 0, 0, false]; } continue; }
      if (st.dest === 'stylesheet') { const nm = v.replace(/;\s*$/, '').trim(); if (nm && st.styleRef !== -1) styles[st.styleRef ?? 0] = (styles[st.styleRef ?? 0] || '') + nm; continue; }
      if (st.dest === 'info-item') { props[st.infoKey] = (props[st.infoKey] || '') + v; continue; }
      if (st.dest === 'info' || st.noText) continue;
      if (st.capture) { st.cap += v; continue; }
      if (st.fieldKind) {
        const k = { PAGE: 'page', NUMPAGES: 'pages', TITLE: 'title', AUTHOR: 'author', DATE: 'date', TIME: 'time', FILENAME: 'file', NUMWORDS: 'words' }[st.fieldKind];
        if (k && st.target?.kind) { if (!st.fieldDone) pushText({ page: '{pagina}', pages: '{paginas}', title: '{titulo}', author: '{autor}', date: '{fecha}', file: '{archivo}' }[k] || v); st.fieldDone = true; continue; }
        if (k) { pushHtml(`<span class="fld" data-fld="${k}" contenteditable="false">${escH(v)}</span>`); continue; }
      }
      pushText(v);
    }
  }
  if (runs.length) endPara();
  // ---------------------------------------------------------------- HTML
  let html = ''; let listOpen = null;
  const closeList = () => { if (listOpen) { html += listOpen.close; listOpen = null; } };
  const isBulletText = (p) => p.list && !p.intbl;
  for (const b of blocks) {
    if (b.type === 'pb') { closeList(); html += '<div class="pb" contenteditable="false"></div>'; continue; }
    if (b.type === 'table') {
      closeList();
      const cols = []; let prevX = 0;
      const all = b.rows.flatMap((r) => r.defs.map((d) => d.x)); const xs = [...new Set(all)].sort((a, q) => a - q);
      xs.forEach((xv) => { cols.push(xv - prevX); prevX = xv; });
      const owner = []; // [r][c] → celda que la cubre
      const cellsHtml = b.rows.map((r, ri) => {
        owner[ri] = [];
        let prevx = 0; const out2 = [];
        r.defs.forEach((d, ci) => {
          const c0 = xs.findIndex((xv) => xv > prevx); const c1 = xs.indexOf(d.x); prevx = d.x;
          const cs = Math.max(1, c1 - c0 + 1);
          const content = (r.cells[ci] || []).map((p) => (p.html ? `<p${paraCss(p)}>${p.html}</p>` : '<p><br></p>')).join('') || '<p><br></p>';
          if (d.vm === 'cont') { const up = owner[ri - 1]?.[c0]; if (up) { up.rs += 1; for (let k = c0; k < c0 + cs; k++) owner[ri][k] = up; return; } }
          if (d.hm === 'cont') { const left = out2[out2.length - 1]; if (left) { left.cs += cs; return; } }
          const cell = { cs, rs: 1, content, tag: r.hdr ? 'th' : 'td', bg: d.bg && colors[d.bg] };
          for (let k = c0; k < c0 + cs; k++) owner[ri][k] = cell;
          out2.push(cell);
        });
        return out2;
      });
      html += `<table class="tbl"><colgroup>${cols.map((w) => `<col style="width:${Math.round(w / TW)}px">`).join('')}</colgroup><tbody>${cellsHtml.map((r) => `<tr>${r.map((c) => `<${c.tag}${c.cs > 1 ? ` colspan="${c.cs}"` : ''}${c.rs > 1 ? ` rowspan="${c.rs}"` : ''}${c.bg ? ` style="background-color:${c.bg}"` : ''}>${c.content}</${c.tag}>`).join('')}</tr>`).join('')}</tbody></table>`;
      continue;
    }
    if (isBulletText(b)) {
      const want = b.list.bullet === false ? 'ol' : 'ul';
      if (!listOpen || listOpen.tag !== want) { closeList(); html += `<${want}>`; listOpen = { tag: want, close: `</${want}>` }; }
      html += `<li>${b.html || '<br>'}</li>`; continue;
    }
    closeList();
    const tg = tagFor(b); const tn = tg.split(' ')[0];
    let inner = b.html || '<br>';
    // el estilo del título ya pone la negrita, el tamaño y el color
    if (tg !== 'p' && tn !== 'blockquote') inner = inner.replace(/<\/?b>/g, '').replace(/<span style="([^"]*)">/g, (m, css) => { const keep = css.split(';').filter((d) => d && !/^(font-size|color|font-family)/.test(d)); return `<span style="${keep.join(';')}">`; }).replace(/<span style="">([\s\S]*?)<\/span>/g, '$1');
    if (tn === 'blockquote') inner = inner.replace(/<\/?i>/g, '');
    html += `<${tg}${tg.includes('class') ? '' : paraCss(b)}${b.pagebb ? ' data-pbb="1"' : ''}>${inner}</${tn}>`;
  }
  closeList();
  const pg = Object.keys(page).length ? page : null;
  if (pg && pg.w && pg.h && pg.landscape == null) pg.landscape = pg.w > pg.h;
  return {
    html, page: pg, header: out.header, footer: out.footer, headerFirst: out.headerFirst, footerFirst: out.footerFirst, hfFirst: !!out.hfFirst, firstPageNum: out.firstPageNum || 1,
    footnotes: out.footnotes, endnotes: out.endnotes,
    props: { title: (props.title || '').trim(), author: (props.author || '').trim(), subject: (props.subject || '').trim(), keywords: (props.keywords || '').trim(), description: (props.doccomm || '').trim() },
  };
}
