// Hojas de cálculo ODF (.ods): abrir y guardar, con fórmulas OpenFormula. GPL-3.0
import JSZip from 'jszip';
import { colName, colIndex, DEF_W, DEF_H } from './workbook.js';
import { dateToSerial } from './format.js';

const NS = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0', style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0', table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', number: 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
  of: 'urn:oasis:names:tc:opendocument:xmlns:of:1.2', dc: 'http://purl.org/dc/elements/1.1/', meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0', xlink: 'http://www.w3.org/1999/xlink', tableooo: 'http://openoffice.org/2009/table',
};
const NSDECL = Object.entries(NS).map(([k, v]) => `xmlns:${k}="${v}"`).join(' ');
const x = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const A = (el, ns, name) => el.getAttributeNS(NS[ns], name) || el.getAttribute(`${ns}:${name}`);
const splitStrings = (f, fn) => f.split(/("(?:[^"]|"")*")/).map((p, i) => (i % 2 ? p : fn(p))).join('');

// ---------------- fórmulas: A1 (inglés) <-> OpenFormula
const REF = /((?:'[^']+'|[A-Za-z_][\w.]*)!)?(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?(?![\w(])/g;
export function toOpenFormula(f) {
  return `of:=${splitStrings(f.replace(/^=/, ''), (p) => p.replace(REF, (m, sh, a, b, off, str) => {
    if (/[A-Za-z_.]/.test(str[off - 1] || '')) return m; // parte de un nombre
    const sheet = sh ? sh.slice(0, -1) : '';
    const pre = sheet ? `${sheet.startsWith("'") ? `$${sheet}` : `$${sheet}`}.` : '.';
    return `[${pre}${a}${b ? `:.${b}` : ''}]`;
  }).replace(/,/g, ';'))}`;
}
export function fromOpenFormula(f) {
  let s = f.replace(/^(of:|oooc:|msoxl:)?=/, '');
  s = splitStrings(s, (p) => p.replace(/\[\$?('[^']+'|[^.\]:]+)?\.(\$?[A-Z]{1,3}\$?\d+)(?::\$?('[^']+'|[^.\]:]+)?\.(\$?[A-Z]{1,3}\$?\d+))?\]/g, (m, sh, a, sh2, b) => `${sh ? `${sh}!` : ''}${a}${b ? `:${b}` : ''}`).replace(/;/g, ','));
  return `=${s}`;
}

// ---------------- escribir
export async function writeOds(wb, props = {}) {
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet', { compression: 'STORE' });
  const autos = new Map(); const styleXml = [];
  const numStyles = new Map(); const numXml = [];
  const numberStyle = (fmt) => {
    if (!fmt || fmt === 'General' || fmt === '@') return null;
    if (numStyles.has(fmt)) return numStyles.get(fmt);
    // negativos en rojo: «pos;[Red]-pos» → estilo para positivos + mapa (como lo guarda LibreOffice)
    const secs = fmt.split(';');
    if (secs.length === 2 && /\[(Red|Rojo)\]/i.test(secs[1])) {
      const pos = numberStyle(secs[0]); const name = `N${numStyles.size + 1}`;
      const inner = numXml[numXml.length - 1].replace(/^<number:([\w-]+) style:name="[^"]+">/, '').replace(/<\/number:[\w-]+>$/, '');
      const kind = /^<number:([\w-]+)/.exec(numXml[numXml.length - 1])[1];
      numXml.push(`<number:${kind} style:name="${name}"><style:text-properties fo:color="#ff0000"/><number:text>-</number:text>${inner}<style:map style:condition="value()&gt;=0" style:apply-style-name="${pos}"/></number:${kind}>`);
      numStyles.set(fmt, name); return name;
    }
    const name = `N${numStyles.size + 1}`;
    const dec = ((/\.(0+)/.exec(fmt) || [, ''])[1]).length;
    const grouping = fmt.includes(',') ? ' number:grouping="true"' : '';
    let xml;
    if (/%/.test(fmt)) xml = `<number:percentage-style style:name="${name}"><number:number number:decimal-places="${dec}" number:min-integer-digits="1"/><number:text>%</number:text></number:percentage-style>`;
    else if (/\$/.test(fmt)) xml = `<number:currency-style style:name="${name}"><number:currency-symbol>$</number:currency-symbol><number:number number:decimal-places="${dec}" number:min-integer-digits="1"${grouping}/></number:currency-style>`;
    else if (/[dmy]/i.test(fmt.replace(/"[^"]*"/g, '')) && !/[0#]/.test(fmt)) xml = `<number:date-style style:name="${name}"><number:day number:style="long"/><number:text>/</number:text><number:month number:style="long"/><number:text>/</number:text><number:year number:style="long"/>${/h/.test(fmt) ? '<number:text> </number:text><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/>' : ''}</number:date-style>`;
    else xml = `<number:number-style style:name="${name}"><number:number number:decimal-places="${dec}" number:min-integer-digits="1"${grouping}/></number:number-style>`;
    numXml.push(xml); numStyles.set(fmt, name); return name;
  };
  const cellStyle = (st) => {
    if (!st) return null;
    const tp = [], cp = [], pp = [];
    if (st.b) tp.push('fo:font-weight="bold"'); if (st.i) tp.push('fo:font-style="italic"');
    if (st.u) tp.push('style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"');
    if (st.color) tp.push(`fo:color="${st.color}"`); if (st.size) tp.push(`fo:font-size="${st.size}pt"`);
    if (st.font) tp.push(`style:font-name="${x(st.font)}" fo:font-family="'${x(st.font)}'"`);
    if (st.s) tp.push('style:text-line-through-style="solid"');
    if (st.rot) cp.push(st.rot === 'v' ? 'style:direction="ttb"' : `style:rotation-angle="${(Number(st.rot) + 360) % 360}"`);
    if (st.shrink) cp.push('style:shrink-to-fit="true"');
    if (st.unlocked || st.hidef) cp.push(`style:cell-protect="${st.unlocked ? (st.hidef ? 'formula-hidden' : 'none') : 'protected formula-hidden'}"`);
    if (st.indent) pp.push(`fo:margin-left="${(st.indent * 0.3).toFixed(2)}cm"`);
    if (st.bg) cp.push(`fo:background-color="${st.bg}"`); if (st.wrap) cp.push('fo:wrap-option="wrap"');
    if (st.valign) cp.push(`style:vertical-align="${st.valign}"`);
    const BW = { thin: '0.74pt solid', medium: '1.76pt solid', thick: '2.49pt solid', dashed: '0.74pt dashed', dotted: '0.74pt dotted', double: '1.76pt double', hair: '0.05pt solid' };
    if (st.border) for (const [k, side] of [['t', 'top'], ['b', 'bottom'], ['l', 'left'], ['r', 'right']]) if (st.border[k]) { const [col, sty = 'thin'] = String(st.border[k]).split(' '); cp.push(`fo:border-${side}="${BW[sty] || BW.thin} ${col}"`); if (sty === 'double') cp.push(`style:border-line-width-${side}="0.0071in 0.0102in 0.0071in"`); }
    if (st.align) pp.push(`fo:text-align="${st.align === 'left' ? 'start' : st.align === 'right' ? 'end' : st.align === 'justify' ? 'justify' : 'center'}"`);
    if (st.align === 'fill') cp.push('style:repeat-content="true"');
    const ns = numberStyle(st.fmt);
    if (!tp.length && !cp.length && !pp.length && !ns) return null;
    const key = [tp, cp, pp, ns].join('|');
    if (!autos.has(key)) {
      const name = `ce${autos.size + 1}`; autos.set(key, name);
      styleXml.push(`<style:style style:name="${name}" style:family="table-cell" style:parent-style-name="Default"${ns ? ` style:data-style-name="${ns}"` : ''}>${cp.length ? `<style:table-cell-properties ${cp.join(' ')}/>` : ''}${pp.length ? `<style:paragraph-properties ${pp.join(' ')}/>` : ''}${tp.length ? `<style:text-properties ${tp.join(' ')}/>` : ''}</style:style>`);
    }
    return autos.get(key);
  };
  const rowStyles = new Map();
  const rowStyle = (hh) => { const k = Math.round(hh); if (!rowStyles.has(k)) { const n = `ro${rowStyles.size + 1}`; rowStyles.set(k, n); styleXml.push(`<style:style style:name="${n}" style:family="table-row"><style:table-row-properties style:row-height="${(k / 37.7953).toFixed(3)}cm" style:use-optimal-row-height="false"/></style:style>`); } return rowStyles.get(k); };
  const colStyles = new Map();
  const colStyle = (w) => { const k = Math.round(w); if (!colStyles.has(k)) { const n = `co${colStyles.size + 1}`; colStyles.set(k, n); styleXml.push(`<style:style style:name="${n}" style:family="table-column"><style:table-column-properties style:column-width="${(k / 37.7953).toFixed(3)}cm"/></style:style>`); } return colStyles.get(k); };
  let tables = '';
  for (const [si, s] of wb.sheets.entries()) {
    const u = wb.usedRange(s);
    let tstyle = '';
    if (s.x3?.hidden || s.x3?.tabColor) { tstyle = `ta${si + 1}`; styleXml.push(`<style:style style:name="${tstyle}" style:family="table"><style:table-properties table:display="${s.x3?.hidden ? 'false' : 'true'}"${s.x3?.tabColor ? ` tableooo:tab-color="${s.x3.tabColor}"` : ''}/></style:style>`); }
    let t = `<table:table table:name="${x(s.name)}"${tstyle ? ` table:style-name="${tstyle}"` : ''}>`;
    for (let c = 0; c <= u.c2; c++) t += `<table:table-column table:style-name="${colStyle(s.cw(c) || DEF_W)}" table:default-cell-style-name="Default"/>`;
    const covered = new Set();
    for (let r = 0; r <= u.r2; r++) {
      t += `<table:table-row${s.rowH.has(r) ? ` table:style-name="${rowStyle(s.rowH.get(r))}"` : ''}${s.hiddenRows.has(r) ? ' table:visibility="collapse"' : ''}>`;
      for (let c = 0; c <= u.c2; c++) {
        if (covered.has(`${r},${c}`)) { t += '<table:covered-table-cell/>'; continue; }
        const raw = wb.raw(r, c, s), st = s.style(r, c), val = wb.value(r, c, s);
        const m = s.merges.find((mm) => mm.r1 === r && mm.c1 === c);
        if (m) for (let rr = m.r1; rr <= m.r2; rr++) for (let cc = m.c1; cc <= m.c2; cc++) if (rr !== r || cc !== c) covered.add(`${rr},${cc}`);
        const sn = cellStyle(st);
        const attrs = [sn ? `table:style-name="${sn}"` : '', m ? `table:number-columns-spanned="${m.c2 - m.c1 + 1}" table:number-rows-spanned="${m.r2 - m.r1 + 1}"` : ''];
        const isF = typeof raw === 'string' && raw.startsWith('=');
        if (isF) attrs.push(`table:formula="${x(toOpenFormula(raw))}"`);
        const shown = wb.display(r, c, s).text;
        const note = s.notes?.get(`${r},${c}`);
        const link = s.x3?.links?.[`${r},${c}`]; const P = (txt) => (link ? `<text:p><text:a xlink:type="simple" xlink:href="${x(link)}">${x(txt)}</text:a></text:p>` : `<text:p>${x(txt)}</text:p>`);
        const ann = note ? `<office:annotation><text:p>${x(note)}</text:p></office:annotation>` : '';
        if (typeof val === 'number') { const pct = st?.fmt && /%/.test(st.fmt); const cur = st?.fmt && /\$/.test(st.fmt); t += `<table:table-cell ${attrs.join(' ')} office:value-type="${pct ? 'percentage' : cur ? 'currency' : 'float'}"${cur ? ' office:currency="MXN"' : ''} office:value="${val}">${ann}<text:p>${x(shown)}</text:p></table:table-cell>`; }
        else if (typeof val === 'boolean') t += `<table:table-cell ${attrs.join(' ')} office:value-type="boolean" office:boolean-value="${val}">${ann}<text:p>${val ? 'VERDADERO' : 'FALSO'}</text:p></table:table-cell>`;
        else if (val != null && val !== '' && !(typeof val === 'object')) t += `<table:table-cell ${attrs.join(' ')} office:value-type="string">${ann}${String(val).split('\n').map((l) => P(l)).join('')}</table:table-cell>`;
        else if (isF) t += `<table:table-cell ${attrs.join(' ')} office:value-type="string">${ann}<text:p>${x(shown)}</text:p></table:table-cell>`;
        else t += `<table:table-cell ${attrs.filter(Boolean).join(' ')}>${ann}</table:table-cell>`;
      }
      t += '</table:table-row>';
    }
    tables += `${t}</table:table>`;
  }
  const names = (wb.names || []).map((n) => { const ref = n.expr.replace(/^=/, ''); const of = toOpenFormula(`=${ref}`).replace(/^of:=/, ''); return `<table:named-range table:name="${x(n.name)}" table:base-cell-address="$${x((wb.sheets[0]?.name) || 'Hoja1')}.$A$1" table:cell-range-address="${x(of.replace(/^\[|\]$/g, ''))}"/>`; }).join('');
  const content = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${NSDECL} office:version="1.3"><office:automatic-styles>${numXml.join('')}${styleXml.join('')}</office:automatic-styles>
<office:body><office:spreadsheet>${tables}${names ? `<table:named-expressions>${names}</table:named-expressions>` : ''}</office:spreadsheet></office:body></office:document-content>`;
  const styles = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NSDECL} office:version="1.3"><office:styles><style:default-style style:family="table-cell"><style:text-properties style:font-name="Calibri" fo:font-size="11pt" fo:language="es" fo:country="MX"/></style:default-style><style:style style:name="Default" style:family="table-cell"/></office:styles></office:document-styles>`;
  const meta = `<?xml version="1.0" encoding="UTF-8"?><office:document-meta ${NSDECL} office:version="1.3"><office:meta><meta:generator>Calcula Datos · Mi Oficina MX</meta:generator><dc:title>${x(props.title)}</dc:title><dc:creator>${x(props.author)}</dc:creator></office:meta></office:document-meta>`;
  zip.file('content.xml', content); zip.file('styles.xml', styles); zip.file('meta.xml', meta);
  // datos propios de Calcula Datos que ODF no guarda tal cual (escenarios)
  const extra = {}; for (const s of wb.sheets) if (s.x3?.scenarios?.length) extra[s.name] = { scenarios: s.x3.scenarios };
  if (Object.keys(extra).length) zip.file('calcula.json', JSON.stringify({ app: 'Calcula', v: 1, sheets: extra }));
  zip.file('META-INF/manifest.xml', '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/></manifest:manifest>');
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.oasis.opendocument.spreadsheet' });
}

// ---------------- leer
function readStyles(doc, map, nums) {
  const pendingMaps = [];
  for (const ns of ['number-style', 'percentage-style', 'currency-style', 'date-style', 'time-style']) {
    for (const el of doc.getElementsByTagNameNS(NS.number, ns)) {
      const name = A(el, 'style', 'name'); const n = el.getElementsByTagNameNS(NS.number, 'number')[0];
      const dec = n ? Number(A(n, 'number', 'decimal-places') || 0) : 2; const grp = n && A(n, 'number', 'grouping') === 'true';
      const d = dec ? `.${'0'.repeat(dec)}` : '';
      const map = [...el.children].find((c) => c.localName === 'map' && /value\(\)\s*>=\s*0/.test(A(c, 'style', 'condition') || ''));
      const red = [...el.children].some((c) => c.localName === 'text-properties' && /^#(ff0000|c9211e|c00000|e00000)$/i.test(A(c, 'fo', 'color') || ''));
      if (map && red) { pendingMaps.push([name, A(map, 'style', 'apply-style-name')]); }
      nums.set(name, ns === 'percentage-style' ? `0${d}%` : ns === 'currency-style' ? `"$"#,##0${d}` : ns === 'date-style' ? 'dd/mm/yyyy' : ns === 'time-style' ? 'hh:mm' : `${grp ? '#,##0' : '0'}${d}`);
    }
  }
  for (const [n, pos] of pendingMaps) if (nums.has(pos)) nums.set(n, `${nums.get(pos)};[Red]-${nums.get(pos)}`);
  for (const st of doc.getElementsByTagNameNS(NS.style, 'style')) {
    const name = A(st, 'style', 'name'); const o = {};
    const tp = st.getElementsByTagNameNS(NS.style, 'text-properties')[0], cp = st.getElementsByTagNameNS(NS.style, 'table-cell-properties')[0], pp = st.getElementsByTagNameNS(NS.style, 'paragraph-properties')[0];
    const colp = st.getElementsByTagNameNS(NS.style, 'table-column-properties')[0];
    if (tp) { if (A(tp, 'style', 'text-line-through-style') && A(tp, 'style', 'text-line-through-style') !== 'none') o.s = true; const fn = A(tp, 'style', 'font-name') || A(tp, 'fo', 'font-family'); if (fn && !/^'?(Calibri|Liberation Sans|Arial)'?$/i.test(fn)) o.font = fn.replace(/'/g, ''); }
    if (tp) { if (A(tp, 'fo', 'font-weight') === 'bold') o.b = true; if (A(tp, 'fo', 'font-style') === 'italic') o.i = true; const ul = A(tp, 'style', 'text-underline-style'); if (ul && ul !== 'none') o.u = true; const col = A(tp, 'fo', 'color'); if (col && col.toLowerCase() !== '#000000') o.color = col; const sz = A(tp, 'fo', 'font-size'); if (sz?.endsWith('pt') && parseFloat(sz) !== 11 && parseFloat(sz) !== 10) o.size = parseFloat(sz); }
    if (cp) {
      const BS = (v) => { if (!v || v === 'none') return null; const m = /([\d.]+)(pt|cm|mm|in)?\s+([\w-]+)\s+(#[0-9a-f]{6})/i.exec(v); if (!m) return null; const pt = parseFloat(m[1]) * ({ cm: 28.35, mm: 2.835, in: 72 }[m[2]] || 1); const sty = /^double/.test(m[3]) ? ' double' : /dash/.test(m[3]) ? ' dashed' : /dot/.test(m[3]) ? ' dotted' : pt >= 2.3 ? ' thick' : pt >= 1.4 ? ' medium' : ''; return `${m[4].toLowerCase()}${sty}`; };
      const all = A(cp, 'fo', 'border'); const bd = {};
      for (const [k, side] of [['t', 'top'], ['b', 'bottom'], ['l', 'left'], ['r', 'right']]) { const v = BS(A(cp, 'fo', `border-${side}`) || all); if (v) bd[k] = v; }
      if (Object.keys(bd).length) o.border = bd;
      const ra = Number(A(cp, 'style', 'rotation-angle') || 0); if (ra) o.rot = ra > 180 ? ra - 360 : ra; if (A(cp, 'style', 'direction') === 'ttb') o.rot = 'v';
      if (A(cp, 'style', 'shrink-to-fit') === 'true') o.shrink = true; if (A(cp, 'style', 'repeat-content') === 'true') o.align = 'fill';
      const pr = A(cp, 'style', 'cell-protect'); if (pr) { if (/none|^formula-hidden$/.test(pr)) o.unlocked = true; if (/formula-hidden|hidden-and-protected/.test(pr)) o.hidef = true; }
    }
    if (cp) { const bg = A(cp, 'fo', 'background-color'); if (bg && bg !== 'transparent') o.bg = bg; if (A(cp, 'fo', 'wrap-option') === 'wrap') o.wrap = true; const va = A(cp, 'style', 'vertical-align'); if (va === 'top' || va === 'middle') o.valign = va; }
    if (pp) { const al = A(pp, 'fo', 'text-align'); if (al && o.align !== 'fill') o.align = al === 'center' ? 'center' : (al === 'end' || al === 'right') ? 'right' : al === 'justify' ? 'justify' : 'left'; const ml = A(pp, 'fo', 'margin-left'); if (ml && parseFloat(ml) > 0) o.indent = Math.max(1, Math.round(parseFloat(ml) * (ml.endsWith('in') ? 2.54 : ml.endsWith('mm') ? 0.1 : ml.endsWith('pt') ? 0.0353 : 1) / 0.3)); }
    const rowp = st.getElementsByTagNameNS(NS.style, 'table-row-properties')[0];
    if (rowp && A(rowp, 'style', 'use-optimal-row-height') !== 'true') { const rh = A(rowp, 'style', 'row-height'); if (rh) o.rowH = Math.round(parseFloat(rh) * (rh.endsWith('in') ? 96 : rh.endsWith('mm') ? 3.77953 : rh.endsWith('pt') ? 1.3333 : 37.7953)); }
    const ds = A(st, 'style', 'data-style-name'); if (ds && nums.has(ds)) o.fmt = nums.get(ds);
    const tblp = st.getElementsByTagNameNS(NS.style, 'table-properties')[0];
    if (tblp) { if (A(tblp, 'table', 'display') === 'false') o.hiddenSheet = true; const tc = A(tblp, 'tableooo', 'tab-color'); if (tc) o.tabColor = tc.toLowerCase(); }
    if (colp) o.colW = Math.round(parseFloat(A(colp, 'style', 'column-width')) * (A(colp, 'style', 'column-width')?.endsWith('in') ? 96 : 37.7953));
    map.set(name, o);
  }
}
export async function readOds(wb, bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const txt = await zip.file('content.xml')?.async('string'); if (!txt) throw new Error('No es una hoja de cálculo ODF válida');
  const doc = new DOMParser().parseFromString(txt, 'application/xml');
  const sdoc = zip.file('styles.xml') ? new DOMParser().parseFromString(await zip.file('styles.xml').async('string'), 'application/xml') : null;
  const map = new Map(), nums = new Map();
  if (sdoc) readStyles(sdoc, map, nums); readStyles(doc, map, nums);
  wb.reset();
  const MAXR = 100000, MAXC = 1024;
  for (const tbl of doc.getElementsByTagNameNS(NS.table, 'table')) {
    if (tbl.parentElement?.localName !== 'spreadsheet') continue;
    const s = wb.addSheet(A(tbl, 'table', 'name') || undefined);
    const ts = map.get(A(tbl, 'table', 'style-name')); if (ts?.hiddenSheet || ts?.tabColor) { s.x3 ||= {}; if (ts.hiddenSheet) s.x3.hidden = true; if (ts.tabColor) s.x3.tabColor = ts.tabColor; }
    const data = [];
    let c0 = 0;
    for (const col of tbl.getElementsByTagNameNS(NS.table, 'table-column')) { const rep = Math.min(Number(A(col, 'table', 'number-columns-repeated') || 1), 200); const cs = map.get(A(col, 'table', 'style-name')); for (let k = 0; k < rep; k++, c0++) if (cs?.colW && Math.abs(cs.colW - DEF_W) > 4 && c0 < MAXC) s.colW.set(c0, cs.colW); }
    let r = 0;
    for (const row of tbl.getElementsByTagNameNS(NS.table, 'table-row')) {
      const rrep = Number(A(row, 'table', 'number-rows-repeated') || 1);
      const cells = [...row.children].filter((e) => e.localName === 'table-cell' || e.localName === 'covered-table-cell');
      const hasContent = cells.some((e) => e.childElementCount || A(e, 'office', 'value-type') || A(e, 'table', 'formula'));
      if (!hasContent) { r += rrep; if (r > MAXR) break; continue; }
      const rst = map.get(A(row, 'table', 'style-name'));
      for (let k = 0; k < Math.min(rrep, 1000); k++, r++) {
        if (rst?.rowH && Math.abs(rst.rowH - DEF_H) > 2) s.rowH.set(r, rst.rowH);
        if (A(row, 'table', 'visibility') === 'collapse' || A(row, 'table', 'visibility') === 'filter') s.hiddenRows.add(r);
        let c = 0; const rowArr = [];
        for (const cell of cells) {
          const crep = Math.min(Number(A(cell, 'table', 'number-columns-repeated') || 1), MAXC);
          for (let j = 0; j < crep; j++, c++) {
            if (c >= MAXC || cell.localName === 'covered-table-cell') continue;
            const f = A(cell, 'table', 'formula'); const vt = A(cell, 'office', 'value-type');
            let v = null;
            if (f) v = fromOpenFormula(f);
            else if (vt === 'float' || vt === 'percentage' || vt === 'currency') v = Number(A(cell, 'office', 'value'));
            else if (vt === 'boolean') v = A(cell, 'office', 'boolean-value') === 'true';
            else if (vt === 'date') { const d = new Date(`${A(cell, 'office', 'date-value')}${A(cell, 'office', 'date-value')?.length <= 10 ? 'T00:00:00Z' : 'Z'}`); v = dateToSerial(d); }
            else if (vt === 'time') { const m = /PT(\d+)H(\d+)M(\d+)/.exec(A(cell, 'office', 'time-value') || ''); v = m ? (Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3])) / 86400 : null; }
            else if (vt === 'string' || cell.getElementsByTagNameNS(NS.text, 'p').length) v = [...cell.getElementsByTagNameNS(NS.text, 'p')].filter((p) => p.parentElement === cell).map((p) => p.textContent).join('\n') || null;
            rowArr[c] = v;
            const st = { ...(map.get(A(cell, 'table', 'style-name')) || {}) }; delete st.colW; delete st.rowH;
            const la = cell.getElementsByTagNameNS(NS.text, 'a')[0]; const href = la && (la.getAttributeNS(NS.xlink, 'href') || la.getAttribute('xlink:href'));
            if (href) { s.x3 ||= {}; s.x3.links = { ...(s.x3.links || {}), [`${r},${c}`]: href }; }
            if (vt === 'date' && !st.fmt) st.fmt = 'dd/mm/yyyy';
            if (Object.keys(st).length) s.styles.set(`${r},${c}`, Object.freeze(st));
            const cs = Number(A(cell, 'table', 'number-columns-spanned') || 1), rs = Number(A(cell, 'table', 'number-rows-spanned') || 1);
            if (cs > 1 || rs > 1) s.merges.push({ r1: r, c1: c, r2: r + rs - 1, c2: c + cs - 1 });
            const ann = cell.getElementsByTagNameNS(NS.office, 'annotation')[0];
            if (ann) s.notes.set(`${r},${c}`, [...ann.getElementsByTagNameNS(NS.text, 'p')].map((p) => p.textContent).join('\n'));
          }
        }
        data[r] = rowArr;
      }
      if (r > MAXR) break;
    }
    wb.hf.setSheetContent(s.id, Array.from({ length: data.length }, (_, i) => data[i] || []));
    const u = wb.usedRange(s); s.nRows = Math.max(1000, u.r2 + 200); s.nCols = Math.max(52, u.c2 + 20);
  }
  for (const nr of doc.getElementsByTagNameNS(NS.table, 'named-range')) {
    const name = A(nr, 'table', 'name'); const addr = A(nr, 'table', 'cell-range-address');
    if (!name || !addr) continue;
    const ref = fromOpenFormula(`of:=[${addr.replace(/^\$?([^.]+)\./, '$$$1.')}]`);
    try { wb.hf.addNamedExpression(name, ref); wb.names.push({ name, expr: ref }); } catch { /* */ }
  }
  if (!wb.sheets.length) wb.addSheet('Hoja1');
  try { const j = JSON.parse(await zip.file('calcula.json')?.async('string') || '{}'); for (const [name, x] of Object.entries(j.sheets || {})) { const sh = wb.sheets.find((q) => q.name === name); if (sh && x.scenarios) { sh.x3 ||= {}; sh.x3.scenarios = x.scenarios; } } } catch { /* sin datos propios */ }
  wb.active = 0; wb.hf.clearUndoStack?.(); wb.undoStack = []; wb.redoStack = [];
  return { warnings: [] };
}
export { colName, colIndex, DEF_H };
