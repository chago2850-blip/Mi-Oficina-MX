// Abre .docx (vía mammoth) y otros formatos de texto, y los convierte al HTML del editor. GPL-3.0
import temml from 'temml';
import { ommlToLatex, OMML_NS } from './omml.js';
import { META_DOCX, applyMeta } from './meta.js';
import mammoth from 'mammoth';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const TWIP_PX = 1 / 15;

export function sanitize(html) {
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  doc.querySelectorAll('script,style,iframe,object,embed,link,meta,form,input,button,textarea,select').forEach((n) => n.remove());
  doc.querySelectorAll('*').forEach((el) => {
    for (const a of [...el.attributes]) {
      if (/^on/i.test(a.name)) el.removeAttribute(a.name);
      if (['href', 'src'].includes(a.name) && /^\s*javascript:/i.test(a.value)) el.removeAttribute(a.name);
      if (a.name === 'src' && el.tagName === 'IMG' && !/^data:image\//i.test(a.value)) el.removeAttribute('src');
    }
  });
  return doc.body.innerHTML;
}

async function docxExtras(bytes) {
  // Datos que mammoth no conserva: tamaño de página, márgenes y alineación de párrafos.
  const out = { page: null, aligns: [] };
  try {
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml')?.async('string');
    if (!xml) return out;
    const d = new DOMParser().parseFromString(xml, 'application/xml');
    const body = d.getElementsByTagNameNS(W, 'body')[0];
    const sect = body ? [...body.getElementsByTagNameNS(W, 'sectPr')].pop() : null;
    if (sect) {
      const sz = sect.getElementsByTagNameNS(W, 'pgSz')[0];
      const mar = sect.getElementsByTagNameNS(W, 'pgMar')[0];
      const g = (el, a) => (el ? Number(el.getAttributeNS(W, a) || el.getAttribute(`w:${a}`)) : NaN);
      const w = g(sz, 'w'), hh = g(sz, 'h');
      if (w && hh) {
        out.page = {
          w: Math.round(w * TWIP_PX), h: Math.round(hh * TWIP_PX),
          mt: Math.round((g(mar, 'top') || 1440) * TWIP_PX), mb: Math.round((g(mar, 'bottom') || 1440) * TWIP_PX),
          ml: Math.round((g(mar, 'left') || 1440) * TWIP_PX), mr: Math.round((g(mar, 'right') || 1440) * TWIP_PX),
          landscape: w > hh,
        };
      }
    }
    // Alineación de cada párrafo/tabla de nivel superior, en orden
    const walk = (parent) => {
      for (const n of parent.children) {
        if (n.localName === 'p') {
          const jc = n.getElementsByTagNameNS(W, 'pPr')[0]?.getElementsByTagNameNS(W, 'jc')[0];
          const v = jc ? (jc.getAttributeNS(W, 'val') || jc.getAttribute('w:val')) : null;
          const tabs = [...(n.getElementsByTagNameNS(W, 'pPr')[0]?.getElementsByTagNameNS(W, 'tab') || [])]
            .map((t) => { const k = { left: 'l', start: 'l', center: 'c', right: 'r', end: 'r', decimal: 'r' }[t.getAttributeNS(W, 'val') || t.getAttribute('w:val')]; const pos = Number(t.getAttributeNS(W, 'pos') || t.getAttribute('w:pos')); return k && pos > 0 ? `${k}:${Math.round(pos * TWIP_PX)}` : null; }).filter(Boolean).join(',');
          out.aligns.push({ kind: 'p', align: { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[v] || null, tabs });
        } else if (n.localName === 'tbl') out.aligns.push({ kind: 'tbl' });
        else if (n.localName === 'sdt') { const c = n.getElementsByTagNameNS(W, 'sdtContent')[0]; if (c) walk(c); }
      }
    };
    if (body) walk(body);
  } catch { /* opcional */ }
  return out;
}

function applyAligns(container, aligns) {
  let i = 0;
  const nextP = () => { while (i < aligns.length && aligns[i].kind !== 'p') i++; return aligns[i++]; };
  for (const el of container.children) {
    if (/^(P|H[1-6]|BLOCKQUOTE)$/.test(el.tagName)) { const a = nextP(); if (a?.align) el.style.textAlign = a.align; if (a?.tabs) el.dataset.tabs = a.tabs; }
    else if (el.tagName === 'UL' || el.tagName === 'OL') { el.querySelectorAll('li').forEach(() => nextP()); }
    else if (el.tagName === 'TABLE') { while (i < aligns.length && aligns[i].kind !== 'tbl') i++; i++; }
  }
}

// ------------------------------------------------------------------ comentarios y revisiones
// mammoth no conserva comentarios ni texto eliminado: antes de convertir, se marcan con
// caracteres privados (U+E000…U+E001) dentro de document.xml y después se envuelven en el HTML.
export const MK0 = '\uE000', MK1 = '\uE001';
export const b64 = (s) => btoa(unescape(encodeURIComponent(s)));
const unb64 = (s) => decodeURIComponent(escape(atob(s)));
const attr = (el, a) => el.getAttributeNS(W, a) || el.getAttribute(`w:${a}`) || '';
const textOf = (el) => [...el.getElementsByTagNameNS(W, 'p')].map((p) => [...p.getElementsByTagNameNS(W, 't')].map((t) => t.textContent).join('')).join('\n').trim();
// Comentarios de comments.xml (+ respuestas de commentsExtended.xml)
export async function readComments(zip) {
  const cxml = await zip.file('word/comments.xml')?.async('string');
  if (!cxml) return null;
  const cd = new DOMParser().parseFromString(cxml, 'application/xml');
  const byPara = new Map(); const list = [];
  for (const c of cd.getElementsByTagNameNS(W, 'comment')) {
    const item = { wid: attr(c, 'id'), author: attr(c, 'author') || 'Autor', date: attr(c, 'date') || new Date().toISOString(), text: textOf(c), replies: [], resolved: false };
    const lines = item.text.split('\n'); const main = [];
    for (const l of lines) { const m = /^↪\s*([^:]*):\s?(.*)$/.exec(l); if (m) item.replies.push({ author: m[1].trim() || 'Autor', date: item.date, text: m[2] }); else main.push(l); }
    item.text = main.join('\n').trim();
    list.push(item);
    const ps = c.getElementsByTagNameNS(W, 'p'); const last = ps[ps.length - 1];
    const pid = last && (last.getAttribute('w14:paraId') || [...last.attributes].find((a) => a.localName === 'paraId')?.value);
    if (pid) byPara.set(pid, item);
  }
  const ext = await zip.file('word/commentsExtended.xml')?.async('string');
  if (ext) {
    const ed = new DOMParser().parseFromString(ext, 'application/xml');
    for (const e of ed.getElementsByTagName('*')) {
      if (e.localName !== 'commentEx') continue;
      const a = (n) => [...e.attributes].find((x) => x.localName === n)?.value;
      const me = byPara.get(a('paraId')); const parent = byPara.get(a('paraIdParent'));
      if (me && a('done') === '1') me.resolved = true;
      if (me && parent && me !== parent) { parent.replies.push({ author: me.author, date: me.date, text: me.text }); me.reply = true; }
    }
  }
  return { list, replyIds: new Set(list.filter((c) => c.reply).map((c) => c.wid)) };
}
async function markRevisions(bytes) {
  const out = { bytes, comments: [], endnoteIds: new Set() };
  try {
    const zip = await JSZip.loadAsync(bytes);
    try { const mt = await zip.file(META_DOCX)?.async('string'); if (mt) out.meta = JSON.parse(mt); } catch { /* */ }
    const xml = await zip.file('word/document.xml')?.async('string');
    if (!xml) return out;
    const d = new DOMParser().parseFromString(xml, 'application/xml');
    const mkRun = (txt) => { const r = d.createElementNS(W, 'w:r'); const t = d.createElementNS(W, 'w:t'); t.setAttribute('xml:space', 'preserve'); t.textContent = MK0 + txt + MK1; r.append(t); return r; };
    let changed = false;
    // Comentarios
    const rc = await readComments(zip);
    if (rc) {
      const { list } = rc;
      out.comments = list.filter((c) => !c.reply);
      const replyIds = rc.replyIds;
      for (const tag of ['commentRangeStart', 'commentRangeEnd']) {
        for (const el of [...d.getElementsByTagNameNS(W, tag)]) {
          const id = attr(el, 'id');
          if (replyIds.has(id)) { el.remove(); continue; }
          el.replaceWith(mkRun(`${tag === 'commentRangeEnd' ? '/' : ''}c:${id}`)); changed = true;
        }
      }
    }
    // Ecuaciones de Word (OMML) -> LaTeX editable
    for (const el of [...d.getElementsByTagNameNS(OMML_NS, 'oMathPara'), ...d.getElementsByTagNameNS(OMML_NS, 'oMath')]) {
      if (!d.documentElement.contains(el)) continue;
      if (el.localName === 'oMath' && el.parentNode?.localName === 'oMathPara') continue;
      const tex = ommlToLatex(el);
      if (!tex) continue;
      const run = mkRun(`q:${b64(tex)}`);
      if (el.parentNode?.localName === 'body') { const p = d.createElementNS(W, 'w:p'); p.append(run); el.replaceWith(p); } else el.replaceWith(run);
      changed = true;
    }
    // Inserciones y eliminaciones (solo las que contienen texto, no las marcas de párrafo)
    for (const kind of ['ins', 'del']) {
      for (const el of [...d.getElementsByTagNameNS(W, kind)]) {
        if (el.parentNode?.localName === 'rPr' || !el.getElementsByTagNameNS(W, 'r').length) continue;
        const meta = [attr(el, 'author'), attr(el, 'date')].map((x) => x.replace(/[|\uE000\uE001]/g, '')).join('|');
        el.parentNode.insertBefore(mkRun(`${kind[0]}:${meta}`), el);
        if (kind === 'del') for (const dt of [...el.getElementsByTagNameNS(W, 'delText')]) { const t = d.createElementNS(W, 'w:t'); t.setAttribute('xml:space', 'preserve'); t.textContent = dt.textContent; dt.replaceWith(t); }
        while (el.firstChild) el.parentNode.insertBefore(el.firstChild, el);
        el.replaceWith(mkRun(`/${kind[0]}`));
        changed = true;
      }
    }
    // Notas al final: se reconocen para mostrarlas como tales
    for (const r of d.getElementsByTagNameNS(W, 'endnoteReference')) out.endnoteIds.add(attr(r, 'id'));
    if (!changed) return out;
    zip.file('word/document.xml', new XMLSerializer().serializeToString(d));
    out.bytes = await zip.generateAsync({ type: 'uint8array' });
  } catch (e) { console.warn('revisiones', e); }
  return out;
}
export function wrapMarks(box, comments) {
  const re = /\uE000([^\uE001]*)\uE001/g;
  const ids = new Map(); const used = new Set();
  const cid = (w) => { if (!ids.has(w)) ids.set(w, `c${Math.random().toString(36).slice(2, 9)}`); return ids.get(w); };
  const open = []; let rev = null;
  const walker = document.createTreeWalker(box, NodeFilter.SHOW_TEXT);
  const nodes = []; while (walker.nextNode()) nodes.push(walker.currentNode);
  for (const t of nodes) {
    const v = t.nodeValue;
    if (!v.includes(MK0) && !open.length && !rev) continue;
    const frag = document.createDocumentFragment();
    let last = 0; let m;
    const emit = (txt) => {
      if (!txt) return;
      let node = document.createTextNode(txt);
      if (rev) { const e = document.createElement(rev.kind); if (rev.author) e.dataset.author = rev.author; if (rev.date) e.dataset.date = rev.date; e.title = `${rev.kind === 'ins' ? 'Insertado' : 'Eliminado'} por ${rev.author || 'Autor'}`; e.append(node); node = e; }
      for (const c of [...open].reverse()) { const sp = document.createElement('span'); sp.className = 'cm'; sp.dataset.cid = cid(c); used.add(c); sp.append(node); node = sp; }
      frag.append(node);
    };
    re.lastIndex = 0;
    while ((m = re.exec(v))) {
      emit(v.slice(last, m.index)); last = re.lastIndex;
      const k = m[1];
      if (k.startsWith('c:')) open.push(k.slice(2));
      else if (k.startsWith('/c:')) { const i = open.lastIndexOf(k.slice(3)); if (i >= 0) open.splice(i, 1); }
      else if (k.startsWith('i:') || k.startsWith('d:')) { const [author, date] = k.slice(2).split('|'); rev = { kind: k[0] === 'i' ? 'ins' : 'del', author, date }; }
      else if (k === '/i' || k === '/d') rev = null;
      else if (k.startsWith('q:')) { const tex = unb64(k.slice(2)); const sp = document.createElement('span'); sp.className = 'eq'; sp.contentEditable = 'false'; sp.dataset.tex = tex; try { sp.innerHTML = temml.renderToString(tex, { annotate: true }); } catch { sp.textContent = tex; } frag.append(sp); }
    }
    emit(v.slice(last));
    t.replaceWith(frag);
  }
  return comments.filter((c) => used.has(c.wid)).map(({ wid, reply, ...c }) => ({ ...c, id: ids.get(wid) }));
}

export async function docxToHtml(bytes) {
  // lector propio (conserva el formato); si falla, se usa mammoth
  try { const { readDocx } = await import('./docx-read.js'); return await readDocx(bytes); } catch (e) { console.warn('lector docx propio', e); }
  const marked = await markRevisions(bytes);
  const src = marked.bytes;
  const buf = src.buffer.slice(src.byteOffset, src.byteOffset + src.byteLength);
  const res = await mammoth.convertToHtml({ arrayBuffer: buf }, {
    ignoreEmptyParagraphs: false,
    styleMap: [
      "p[style-name='Title'] => p.title:fresh",
      "p[style-name='Título'] => p.title:fresh",
      "p[style-name='Subtitle'] => p.subtitle:fresh",
      "p[style-name='Subtítulo'] => p.subtitle:fresh",
      "p[style-name='Quote'] => blockquote:fresh",
      "p[style-name='Cita'] => blockquote:fresh",
      "p[style-name='Intense Quote'] => blockquote:fresh",
      'u => u',
      'strike => s',
    ],
  });
  const extras = await docxExtras(bytes);
  const box = document.createElement('div');
  box.innerHTML = sanitize(res.value || '<p><br></p>');
  applyAligns(box, extras.aligns);
  const comments = wrapMarks(box, marked.comments);
  // Notas al pie y al final de mammoth -> sup.fn / sup.en + listas de notas del documento
  const footnotes = {}, endnotes = {};
  box.querySelectorAll('li[id^="footnote-"], li[id^="endnote-"]').forEach((li) => {
    const key = li.id; const isEnd = key.startsWith('endnote-');
    li.querySelectorAll('a[href^="#footnote-ref-"], a[href^="#endnote-ref-"]').forEach((a) => a.remove());
    const fid = `${isEnd ? 'e' : 'f'}${Math.random().toString(36).slice(2, 8)}`;
    (isEnd ? endnotes : footnotes)[fid] = li.textContent.replace(/[\uE000-\uE001]/g, '').trim();
    box.querySelectorAll(`a[href="#${key}"]`).forEach((a) => {
      const sup = document.createElement('sup');
      sup.className = isEnd ? 'en' : 'fn'; sup.dataset.fid = fid; sup.contentEditable = 'false'; sup.textContent = '#';
      (a.parentElement?.tagName === 'SUP' ? a.parentElement : a).replaceWith(sup);
    });
    const ol = li.parentElement; li.remove();
    if (ol && !ol.children.length) ol.remove();
  });
  // párrafos vacíos -> <p><br></p> para que se puedan editar
  box.querySelectorAll('p:empty').forEach((p) => p.append(document.createElement('br')));
  const warnings = (res.messages || []).filter((m) => m.type === 'warning').map((m) => m.message);
  const { sources } = applyMeta(box, marked.meta, comments, (tex) => { try { return temml.renderToString(tex, { annotate: true }); } catch { return null; } });
  return { html: box.innerHTML || '<p><br></p>', page: extras.page, footnotes, endnotes, comments, sources, warnings: warnings.length ? [`${warnings.length} elementos no se pudieron conservar del todo`] : [] };
}

const escHtml = (s) => s.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));

export function textToHtml(text) {
  return text.replace(/\r\n?/g, '\n').split('\n').map((l) => `<p>${l ? escHtml(l) : '<br>'}</p>`).join('');
}

export function markdownToHtml(md) {
  const inline = (s) => escHtml(s)
    .replace(/\*\*(.+?)\*\*/g, '<b>$1</b>').replace(/__(.+?)__/g, '<b>$1</b>')
    .replace(/(^|[^*])\*(?!\s)(.+?)\*/g, '$1<i>$2</i>').replace(/`(.+?)`/g, '<code>$1</code>')
    .replace(/\[(.+?)\]\((https?:[^)]+)\)/g, '<a href="$2">$1</a>');
  const out = [];
  let list = null;
  for (const raw of md.replace(/\r\n?/g, '\n').split('\n')) {
    const l = raw.trimEnd();
    const hm = /^(#{1,6})\s+(.*)$/.exec(l);
    const ul = /^\s*[-*+]\s+(.*)$/.exec(l);
    const ol = /^\s*\d+[.)]\s+(.*)$/.exec(l);
    if ((ul && list !== 'ul') || (ol && list !== 'ol') || (!ul && !ol && list)) { if (list) out.push(`</${list}>`); list = null; }
    if (ul || ol) { if (!list) { list = ul ? 'ul' : 'ol'; out.push(`<${list}>`); } out.push(`<li>${inline((ul || ol)[1])}</li>`); continue; }
    if (hm) out.push(`<h${hm[1].length}>${inline(hm[2])}</h${hm[1].length}>`);
    else if (/^>\s?/.test(l)) out.push(`<blockquote>${inline(l.replace(/^>\s?/, ''))}</blockquote>`);
    else if (/^(-{3,}|\*{3,})$/.test(l)) out.push('<hr>');
    else out.push(l ? `<p>${inline(l)}</p>` : '<p><br></p>');
  }
  if (list) out.push(`</${list}>`);
  return out.join('');
}

export function htmlToText(el) {
  const lines = [];
  for (const b of el.children) {
    if (b.tagName === 'UL' || b.tagName === 'OL') [...b.querySelectorAll('li')].forEach((li, i) => lines.push(`${b.tagName === 'OL' ? `${i + 1}.` : '•'} ${li.innerText}`));
    else if (b.tagName === 'TABLE') [...b.rows].forEach((r) => lines.push([...r.cells].map((c) => c.innerText).join('\t')));
    else if (!b.classList.contains('pb')) lines.push(b.innerText);
  }
  return lines.join('\n');
}
