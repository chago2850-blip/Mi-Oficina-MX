// Escribiente — Exportar como libro electrónico EPUB 3 (como «Exportar como EPUB» de LibreOffice):
// capítulos en cada título de nivel 1 (o en cada salto de página), índice de navegación con los
// títulos, imágenes, notas al pie al final de cada capítulo, portada opcional y metadatos. GPL-3.0
import JSZip from 'jszip';

const x = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const CSS = `body{font-family:serif;line-height:1.5;margin:0 5%}h1,h2,h3,h4,h5,h6{font-family:sans-serif;line-height:1.2;page-break-after:avoid}
h1{font-size:1.6em;margin:1.2em 0 .6em}h2{font-size:1.3em}h3{font-size:1.1em}p{margin:0 0 .6em}.title{font-size:2em;text-align:center;margin:2em 0 .4em}
.subtitle{font-size:1.2em;text-align:center;color:#555}blockquote{margin:1em 2em;font-style:italic}table{border-collapse:collapse;margin:1em 0}
td,th{border:1px solid #888;padding:.2em .4em;vertical-align:top}img{max-width:100%;height:auto}.caption{font-style:italic;text-align:center;font-size:.9em}
.notes{border-top:1px solid #aaa;margin-top:2em;font-size:.9em}.notes li{margin-bottom:.4em}sup a{text-decoration:none}
.cover{text-align:center;margin-top:30%}.cover h1{font-size:2.2em}.cover p{font-size:1.2em;color:#555}`;

// options: { title, author, subject, keywords, description, lang, split: 'h1'|'pb'|'none', cover: bool, footnotes, endnotes }
export async function htmlToEpub(root, o = {}) {
  const zip = new JSZip();
  zip.file('mimetype', 'application/epub+zip', { compression: 'STORE' });
  zip.file('META-INF/container.xml', '<?xml version="1.0" encoding="UTF-8"?>\n<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>');
  zip.file('OEBPS/style.css', CSS);
  // documento aparte: las imágenes con rutas relativas no se cargan
  const clone = document.implementation.createHTMLDocument('').importNode(root, true);
  clone.querySelectorAll('tr.tpush, tr.trep, .sel').forEach((n) => (n.classList.contains('sel') ? n.classList.remove('sel') : n.remove()));
  clone.querySelectorAll('[contenteditable]').forEach((n) => n.removeAttribute('contenteditable'));
  clone.querySelectorAll('[data-push]').forEach((n) => { n.style.marginTop = ''; n.removeAttribute('data-push'); });
  clone.querySelectorAll('.cbk').forEach((n) => n.remove());
  clone.querySelectorAll('span.ix').forEach((n) => n.remove());
  // imágenes incrustadas → archivos
  const images = []; let imgN = 0;
  clone.querySelectorAll('img').forEach((img) => {
    const m = /^data:image\/(png|jpeg|jpg|gif|svg\+xml|webp);base64,(.+)$/.exec(img.getAttribute('src') || '');
    if (!m) { img.remove(); return; }
    const ext = { jpeg: 'jpg', 'svg+xml': 'svg' }[m[1]] || m[1];
    const name = `img${++imgN}.${ext}`;
    zip.file(`OEBPS/images/${name}`, m[2], { base64: true });
    images.push({ name, type: `image/${m[1] === 'jpg' ? 'jpeg' : m[1]}` });
    img.setAttribute('src', `images/${name}`); if (!img.getAttribute('alt')) img.setAttribute('alt', '');
    img.removeAttribute('style');
  });
  // capítulos
  const chapters = []; let cur = [];
  const flush = () => { if (cur.some((n) => (n.textContent || '').trim() || n.querySelector?.('img,table'))) chapters.push(cur); cur = []; };
  for (const n of [...clone.childNodes]) {
    if (n.nodeType === 1 && n.classList.contains('pb')) { if (o.split === 'pb') flush(); continue; }
    if (n.nodeType === 1 && n.tagName === 'H1' && o.split !== 'none' && o.split !== 'pb') flush();
    cur.push(n);
  }
  flush();
  if (!chapters.length) chapters.push([clone.ownerDocument.createElement('p')]);
  // documentos XHTML
  const ser = new XMLSerializer();
  const xhtml = (title, body) => `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE html>\n<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="${x(o.lang || 'es')}" lang="${x(o.lang || 'es')}"><head><meta charset="utf-8"/><title>${x(title)}</title><link rel="stylesheet" type="text/css" href="style.css"/></head><body>${body}</body></html>`;
  const toc = []; const idMap = {};
  const files = chapters.map((nodes, ci) => {
    const file = `chap${String(ci + 1).padStart(3, '0')}.xhtml`;
    const wrap = clone.ownerDocument.createElement('div');
    nodes.forEach((n) => wrap.append(n));
    // notas al pie al final del capítulo
    const notes = [];
    wrap.querySelectorAll('sup.fn, sup.en').forEach((s) => {
      const txt = (s.classList.contains('fn') ? o.footnotes : o.endnotes)?.[s.dataset.fid]; if (txt == null) { s.remove(); return; }
      const k = notes.length + 1; const id = `n${ci + 1}-${k}`;
      notes.push({ id, txt });
      s.removeAttribute('data-fid'); s.className = 'noteref'; s.innerHTML = `<a epub:type="noteref" href="#${id}" id="r${id}">${k}</a>`;
    });
    wrap.querySelectorAll('h1,h2,h3').forEach((hd, k) => {
      if (!hd.id) hd.id = `h${ci + 1}-${k + 1}`;
      toc.push({ level: Number(hd.tagName[1]), text: hd.textContent.trim(), href: `${file}#${hd.id}` });
    });
    wrap.querySelectorAll('[id]').forEach((el) => { idMap[el.id] = file; });
    let body = '';
    for (const n of wrap.childNodes) body += n.nodeType === 1 ? ser.serializeToString(n).replace(/ xmlns="http:\/\/www\.w3\.org\/1999\/xhtml"/g, '') : x(n.textContent);
    if (notes.length) body += `<aside class="notes" epub:type="footnotes"><ol>${notes.map((nt) => `<li id="${nt.id}" epub:type="footnote">${x(nt.txt)} <a href="#r${nt.id}">↩</a></li>`).join('')}</ol></aside>`;
    const title = wrap.querySelector('h1,h2,.title')?.textContent.trim() || o.title || '';
    return { file, html: body, title };
  });
  // vínculos internos a marcadores de otros capítulos
  files.forEach((f) => { f.html = f.html.replace(/href="#([^"]+)"/g, (m, id) => (idMap[id] && idMap[id] !== f.file ? `href="${idMap[id]}#${id}"` : m)); zip.file(`OEBPS/${f.file}`, xhtml(f.title, f.html)); });
  if (o.cover) zip.file('OEBPS/cover.xhtml', xhtml(o.title || '', `<section class="cover" epub:type="cover"><h1>${x(o.title)}</h1>${o.author ? `<p>${x(o.author)}</p>` : ''}</section>`));
  // índice (nav.xhtml) anidado por niveles y toc.ncx para lectores antiguos
  const items = toc.length ? toc : files.map((f, i) => ({ level: 1, text: f.title || `${i + 1}`, href: f.file }));
  let nav = '<ol>'; let lvl = items[0]?.level || 1; const base = lvl; let open = false;
  for (const it of items) {
    const L = Math.max(base, it.level);
    if (L > lvl) { nav += '<ol>'.repeat(L - lvl); lvl = L; open = false; } else if (L < lvl) { nav += '</li></ol>'.repeat(lvl - L) + '</li>'; lvl = L; open = false; } else if (open) nav += '</li>';
    nav += `<li><a href="${x(it.href)}">${x(it.text || '—')}</a>`; open = true;
  }
  nav += `</li>${'</ol></li>'.repeat(lvl - base)}</ol>`;
  zip.file('OEBPS/nav.xhtml', xhtml('Índice', `<nav epub:type="toc" id="toc"><h1>${x(o.tocTitle || 'Índice')}</h1>${nav}</nav>`));
  const uid = `urn:uuid:${crypto.randomUUID ? crypto.randomUUID() : Date.now()}`;
  zip.file('OEBPS/toc.ncx', `<?xml version="1.0" encoding="UTF-8"?>\n<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1"><head><meta name="dtb:uid" content="${uid}"/></head><docTitle><text>${x(o.title)}</text></docTitle><navMap>${items.map((it, i) => `<navPoint id="np${i + 1}" playOrder="${i + 1}"><navLabel><text>${x(it.text)}</text></navLabel><content src="${x(it.href)}"/></navPoint>`).join('')}</navMap></ncx>`);
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const manifest = [
    '<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>', '<item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>', '<item id="css" href="style.css" media-type="text/css"/>',
    ...(o.cover ? ['<item id="cover" href="cover.xhtml" media-type="application/xhtml+xml"/>'] : []),
    ...files.map((f, i) => `<item id="c${i + 1}" href="${f.file}" media-type="application/xhtml+xml"/>`),
    ...images.map((im, i) => `<item id="i${i + 1}" href="images/${im.name}" media-type="${im.type}"/>`),
  ].join('');
  const spine = [...(o.cover ? ['<itemref idref="cover"/>'] : []), ...files.map((f, i) => `<itemref idref="c${i + 1}"/>`)].join('');
  zip.file('OEBPS/content.opf', `<?xml version="1.0" encoding="UTF-8"?>\n<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="uid" xml:lang="${x(o.lang || 'es')}"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:identifier id="uid">${uid}</dc:identifier><dc:title>${x(o.title || 'Documento')}</dc:title><dc:language>${x(o.lang || 'es')}</dc:language>${o.author ? `<dc:creator>${x(o.author)}</dc:creator>` : ''}${o.subject ? `<dc:subject>${x(o.subject)}</dc:subject>` : ''}${o.description ? `<dc:description>${x(o.description)}</dc:description>` : ''}<meta property="dcterms:modified">${now}</meta><meta name="generator" content="Escribiente · Mi Oficina MX"/></metadata><manifest>${manifest}</manifest><spine toc="ncx">${spine}</spine></package>`);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/epub+zip' });
}
