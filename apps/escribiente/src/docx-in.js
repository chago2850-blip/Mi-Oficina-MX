// Abre .docx (vía mammoth) y otros formatos de texto, y los convierte al HTML del editor. GPL-3.0
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
          out.aligns.push({ kind: 'p', align: { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[v] || null });
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
    if (/^(P|H[1-6]|BLOCKQUOTE)$/.test(el.tagName)) { const a = nextP(); if (a?.align) el.style.textAlign = a.align; }
    else if (el.tagName === 'UL' || el.tagName === 'OL') { el.querySelectorAll('li').forEach(() => nextP()); }
    else if (el.tagName === 'TABLE') { while (i < aligns.length && aligns[i].kind !== 'tbl') i++; i++; }
  }
}

export async function docxToHtml(bytes) {
  const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
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
  // párrafos vacíos -> <p><br></p> para que se puedan editar
  box.querySelectorAll('p:empty').forEach((p) => p.append(document.createElement('br')));
  const warnings = (res.messages || []).filter((m) => m.type === 'warning').map((m) => m.message);
  return { html: box.innerHTML || '<p><br></p>', page: extras.page, warnings: warnings.length ? [`${warnings.length} elementos no se pudieron conservar del todo`] : [] };
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
