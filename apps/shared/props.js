// Propiedades del documento como en LibreOffice (Archivo ▸ Propiedades): descripción, comentarios,
// propiedades personalizadas y estadísticas. Se guardan en docProps/core.xml y docProps/custom.xml
// (Office) o en meta.xml (ODF) y se leen de vuelta, para las tres apps. GPL-3.0
import JSZip from 'jszip';
import { h, t, toast } from './office.js';
import { tabbed, field } from './dialogs.js';

const OOXML = ['docx', 'xlsx', 'pptx'];
const ODF = ['odt', 'ods', 'odp'];
const NS = {
  cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', dc: 'http://purl.org/dc/elements/1.1/', dcterms: 'http://purl.org/dc/terms/',
  cu: 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties', vt: 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0', meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0', xsi: 'http://www.w3.org/2001/XMLSchema-instance',
};
const x = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const parse = (s) => new DOMParser().parseFromString(s, 'application/xml');
const ser = (d) => new XMLSerializer().serializeToString(d);
const VT = { text: 'lpwstr', number: 'r8', date: 'filetime', bool: 'bool' };
const OT = { text: 'string', number: 'float', date: 'date', bool: 'boolean' };
const hasExtra = (p) => !!(p?.description || p?.custom?.length || p?.category);

// ---------------------------------------------------------------- escribir
export async function withProps(ext, data, props) {
  if (!(data instanceof Uint8Array) || !hasExtra(props)) return data;
  try {
    if (OOXML.includes(ext)) return await ooxml(data, props);
    if (ODF.includes(ext)) return await odf(data, props);
  } catch (err) { console.error('propiedades', err); }
  return data;
}
async function ooxml(bytes, p) {
  const zip = await JSZip.loadAsync(bytes);
  let core = await zip.file('docProps/core.xml')?.async('string');
  if (core) {
    const d = parse(core); const root = d.documentElement;
    const set = (ns, tag, v) => { let el = d.getElementsByTagNameNS(ns, tag.split(':')[1])[0]; if (!v) { el?.remove(); return; } if (!el) { el = d.createElementNS(ns, tag); root.append(el); } el.textContent = v; };
    set(NS.dc, 'dc:description', p.description); set(NS.cp, 'cp:category', p.category);
    core = ser(d); zip.file('docProps/core.xml', core);
  }
  const custom = (p.custom || []).filter((c) => c.name);
  if (custom.length) {
    const val = (c) => { const ty = VT[c.type] || 'lpwstr'; const v = c.type === 'bool' ? (c.value === true || c.value === 'true' ? 'true' : 'false') : c.type === 'date' ? `${String(c.value).slice(0, 10)}T00:00:00Z` : c.value; return `<vt:${ty}>${x(v)}</vt:${ty}>`; };
    zip.file('docProps/custom.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="${NS.cu}" xmlns:vt="${NS.vt}">${custom.map((c, i) => `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${i + 2}" name="${x(c.name)}">${val(c)}</property>`).join('')}</Properties>`);
    let rels = await zip.file('_rels/.rels').async('string');
    if (!rels.includes('docProps/custom.xml')) rels = rels.replace('</Relationships>', '<Relationship Id="rIdCustomProps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>');
    zip.file('_rels/.rels', rels);
    let ct = await zip.file('[Content_Types].xml').async('string');
    if (!ct.includes('/docProps/custom.xml')) ct = ct.replace('</Types>', '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>');
    zip.file('[Content_Types].xml', ct);
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
async function odf(bytes, p) {
  const zip = await JSZip.loadAsync(bytes);
  const mt = await zip.file('mimetype')?.async('string');
  let meta = await zip.file('meta.xml')?.async('string');
  if (!meta) return bytes;
  const d = parse(meta);
  const om = d.getElementsByTagNameNS(NS.office, 'meta')[0]; if (!om) return bytes;
  const set = (ns, tag, v) => { let el = om.getElementsByTagNameNS(ns, tag.split(':')[1])[0]; if (!v) { el?.remove(); return; } if (!el) { el = d.createElementNS(ns, tag); om.append(el); } el.textContent = v; };
  set(NS.dc, 'dc:description', p.description);
  [...om.getElementsByTagNameNS(NS.meta, 'user-defined')].forEach((el) => el.remove());
  for (const c of (p.custom || []).filter((q) => q.name)) {
    const el = d.createElementNS(NS.meta, 'meta:user-defined');
    el.setAttributeNS(NS.meta, 'meta:name', c.name); el.setAttributeNS(NS.meta, 'meta:value-type', OT[c.type] || 'string');
    el.textContent = c.type === 'bool' ? String(c.value === true || c.value === 'true') : String(c.value ?? '');
    om.append(el);
  }
  if (p.category) { const kw = d.createElementNS(NS.meta, 'meta:user-defined'); kw.setAttributeNS(NS.meta, 'meta:name', 'Category'); kw.setAttributeNS(NS.meta, 'meta:value-type', 'string'); kw.textContent = p.category; om.append(kw); }
  meta = ser(d);
  // el tipo MIME va primero y sin comprimir
  const out = new JSZip();
  if (mt) out.file('mimetype', mt, { compression: 'STORE' });
  for (const [name, f] of Object.entries(zip.files)) {
    if (name === 'mimetype' || f.dir) continue;
    out.file(name, name === 'meta.xml' ? meta : await f.async('uint8array'));
  }
  return out.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

// ---------------------------------------------------------------- leer
export async function readExtraProps(f) {
  try {
    if (!(f?.bytes instanceof Uint8Array) && !f?.bytes?.length) return {};
    const ext = String(f.ext || '').toLowerCase();
    if (!OOXML.includes(ext) && !ODF.includes(ext)) return {};
    const zip = await JSZip.loadAsync(f.bytes);
    const out = {};
    if (OOXML.includes(ext)) {
      const core = await zip.file('docProps/core.xml')?.async('string');
      if (core) {
        const d = parse(core); const g = (ns, tag) => d.getElementsByTagNameNS(ns, tag)[0]?.textContent?.trim() || '';
        Object.assign(out, { description: g(NS.dc, 'description'), category: g(NS.cp, 'category'), created: g(NS.dcterms, 'created'), modified: g(NS.dcterms, 'modified'), revision: g(NS.cp, 'revision') });
      }
      const cx = await zip.file('docProps/custom.xml')?.async('string');
      if (cx) {
        out.custom = [...parse(cx).getElementsByTagNameNS(NS.cu, 'property')].map((el) => {
          const v = el.firstElementChild; const ln = v?.localName || 'lpwstr';
          const type = { r8: 'number', r4: 'number', i4: 'number', i8: 'number', int: 'number', filetime: 'date', bool: 'bool' }[ln] || 'text';
          let value = v?.textContent || '';
          if (type === 'date') value = value.slice(0, 10); if (type === 'bool') value = value === 'true' || value === '1';
          return { name: el.getAttribute('name'), type, value };
        });
      }
    } else {
      const meta = await zip.file('meta.xml')?.async('string');
      if (meta) {
        const d = parse(meta); const g = (ns, tag) => d.getElementsByTagNameNS(ns, tag)[0]?.textContent?.trim() || '';
        Object.assign(out, { description: g(NS.dc, 'description'), created: g(NS.meta, 'creation-date'), modified: g(NS.dc, 'date'), revision: g(NS.meta, 'editing-cycles') });
        const ud = [...d.getElementsByTagNameNS(NS.meta, 'user-defined')].map((el) => {
          const vt = el.getAttributeNS(NS.meta, 'value-type') || el.getAttribute('meta:value-type') || 'string';
          const type = { float: 'number', date: 'date', boolean: 'bool' }[vt] || 'text';
          let value = el.textContent || ''; if (type === 'date') value = value.slice(0, 10); if (type === 'bool') value = value === 'true';
          return { name: el.getAttributeNS(NS.meta, 'name') || el.getAttribute('meta:name'), type, value };
        });
        const cat = ud.find((c) => c.name === 'Category'); if (cat) out.category = String(cat.value);
        // LibreOffice escribe «Info 1…4» vacías en algunos documentos; se omiten
        out.custom = ud.filter((c) => c.name !== 'Category' && !(/^Info \d$/.test(c.name) && !c.value));
      }
    }
    Object.keys(out).forEach((k) => { if (out[k] === '' || (Array.isArray(out[k]) && !out[k].length)) delete out[k]; });
    return out;
  } catch { return {}; }
}

// ---------------------------------------------------------------- diálogo
const fmtWhen = (s) => { if (!s) return '—'; const d = new Date(s); return Number.isNaN(d.getTime()) ? s : d.toLocaleString(); };
export async function propertiesDialog(shell, start = 0) {
  const p = shell.docProps(); const pr = shell.props;
  const row = (k, v) => h('tr', {}, h('th', {}, t(k)), h('td', {}, String(v)));
  const general = h('table', { class: 'props-tbl' },
    row('Nombre', shell.doc.name), row('Tipo', `.${shell.doc.ext}`), row('Ubicación', shell.doc.path || t('Sin guardar')),
    row('Creado', fmtWhen(pr.created)), row('Modificado', fmtWhen(pr.modified)), ...(pr.revision ? [row('Número de revisión', pr.revision)] : []));
  const inp = (name, v) => h('input', { name, value: v || '' });
  const desc = h('div', { class: 'fx-col' },
    field('Título', inp('title', pr.title || '')), field('Autor', inp('author', pr.author || p.author)), field('Asunto', inp('subject', p.subject)),
    field('Palabras clave', inp('keywords', p.keywords)), field('Categoría', inp('category', pr.category)),
    field('Comentarios', h('textarea', { name: 'description', rows: 4, style: { width: '100%', resize: 'vertical', fontFamily: 'inherit' } }, pr.description || '')));
  const list = h('div', { class: 'cprops' });
  const addRow = (c = { name: '', type: 'text', value: '' }) => {
    const ty = h('select', { class: 'cp-type', 'aria-label': t('Tipo') }, ...[['text', 'Texto'], ['number', 'Número'], ['date', 'Fecha'], ['bool', 'Sí o no']].map(([v, l]) => h('option', { value: v, selected: v === c.type ? true : null }, t(l))));
    const valBox = h('span', { class: 'cp-valbox' });
    const mkVal = (type, v) => valBox.replaceChildren(type === 'bool' ? h('input', { type: 'checkbox', class: 'cp-val', checked: v === true || v === 'true' ? true : null, 'aria-label': t('Valor') }) : h('input', { class: 'cp-val', type: type === 'number' ? 'number' : type === 'date' ? 'date' : 'text', value: v ?? '', 'aria-label': t('Valor') }));
    mkVal(c.type, c.value);
    ty.addEventListener('change', () => mkVal(ty.value, ''));
    const r = h('div', { class: 'cp-row' }, h('input', { class: 'cp-name', value: c.name, placeholder: t('Nombre'), 'aria-label': t('Nombre') }), ty, valBox,
      h('button', { type: 'button', class: 'btn', title: t('Quitar'), 'aria-label': t('Quitar'), onclick: () => r.remove() }, '×'));
    list.append(r);
  };
  (pr.custom || []).forEach(addRow);
  const custom = h('div', { class: 'fx-col' }, list, h('button', { type: 'button', class: 'btn', style: { alignSelf: 'flex-start' }, onclick: () => addRow() }, t('Agregar propiedad')));
  const stats = h('table', { class: 'props-tbl' }, ...(shell.adapter.stats?.() || []).map(([k, v]) => row(k, v)));
  const body = await tabbed('Propiedades del documento', [
    { label: 'General', body: general }, { label: 'Descripción', body: desc },
    { label: 'Propiedades personalizadas', body: custom }, { label: 'Estadísticas', body: stats },
  ], { width: '620px', start });
  if (!body) return null;
  const v = (n) => body.querySelector(`[name="${n}"]`).value.trim();
  Object.assign(shell.props, { title: v('title'), author: v('author'), subject: v('subject'), keywords: v('keywords'), category: v('category'), description: body.querySelector('[name=description]').value.trim() });
  shell.props.custom = [...list.querySelectorAll('.cp-row')].map((r) => {
    const type = r.querySelector('.cp-type').value; const val = r.querySelector('.cp-val');
    return { name: r.querySelector('.cp-name').value.trim(), type, value: type === 'bool' ? val.checked : val.value };
  }).filter((c) => c.name);
  shell.markDirty();
  toast('Propiedades actualizadas (se guardan con el archivo)');
  return shell.props;
}
