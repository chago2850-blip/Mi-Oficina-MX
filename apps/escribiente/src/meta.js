// Datos propios de Escribiente que Word y LibreOffice no guardan (reacciones a comentarios,
// diagramas editables, fuentes y citas, LaTeX de las ecuaciones). Se guardan como un archivo
// JSON dentro del .docx/.odt y al abrir se vuelven a aplicar si el documento no cambió de forma. GPL-3.0
export const META_DOCX = 'docProps/escribiente.json';
export const META_ODT = 'escribiente.json';

export function collectMeta(editorEl, settings) {
  const imgs = [...editorEl.querySelectorAll('img')];
  const meta = {
    app: 'Escribiente', v: 1,
    reactions: (settings.comments || []).filter((c) => c.reactions || c.replies?.some((r) => r.reactions)).map((c) => ({ author: c.author, text: c.text, reactions: c.reactions || null, replies: (c.replies || []).map((r) => ({ text: r.text, reactions: r.reactions || null })) })),
    imgCount: imgs.length,
    smartart: imgs.map((im, i) => (im.classList.contains('smartart') ? { i, data: im.dataset.smartart } : null)).filter(Boolean),
    sources: settings.sources || [],
    cites: [...editorEl.querySelectorAll('.cite')].map((c) => ({ src: c.dataset.src, text: c.textContent })),
    eqs: [...editorEl.querySelectorAll('.eq')].map((e) => e.dataset.tex || ''),
  };
  const empty = !meta.reactions.length && !meta.smartart.length && !meta.sources.length && !meta.cites.length && !meta.eqs.some(Boolean);
  return empty ? null : meta;
}

// Aplica los datos al HTML importado (box) y a la lista de comentarios
export function applyMeta(box, meta, comments, renderTex) {
  if (!meta || meta.app !== 'Escribiente') return { sources: [] };
  // reacciones: por autor y texto del comentario
  for (const m of meta.reactions || []) {
    const c = comments.find((x) => x.author === m.author && x.text.trim() === String(m.text || '').trim());
    if (!c) continue;
    if (m.reactions) c.reactions = m.reactions;
    (m.replies || []).forEach((r, k) => { const rc = c.replies.find((x) => x.text.trim() === String(r.text || '').trim()) || c.replies[k]; if (rc && r.reactions) rc.reactions = r.reactions; });
  }
  // diagramas: por posición entre las imágenes
  const imgs = [...box.querySelectorAll('img')];
  if (imgs.length === meta.imgCount) for (const s of meta.smartart || []) { const im = imgs[s.i]; if (im) { im.classList.add('smartart'); im.dataset.smartart = s.data; } }
  // ecuaciones: el LaTeX original
  const eqs = [...box.querySelectorAll('.eq')];
  if (eqs.length === (meta.eqs || []).length) eqs.forEach((e, k) => { const tex = meta.eqs[k]; if (tex && tex !== e.dataset.tex) { const m = renderTex?.(tex); if (m) { e.dataset.tex = tex; e.innerHTML = m; } } });
  // citas: se vuelven a enlazar a sus fuentes buscando su texto en orden
  const walker = document.createTreeWalker(box, NodeFilter.SHOW_TEXT);
  let node = walker.nextNode(); let off = 0;
  for (const c of meta.cites || []) {
    if (!c.text) continue;
    while (node) {
      const i = node.nodeValue.indexOf(c.text, off);
      if (i >= 0 && !node.parentElement.closest('.cite')) {
        const after = node.splitText(i); const rest = after.splitText(c.text.length);
        const sp = document.createElement('span'); sp.className = 'cite'; sp.contentEditable = 'false'; sp.dataset.src = c.src; sp.textContent = c.text;
        after.replaceWith(sp);
        node = rest; off = 0; walker.currentNode = rest;
        break;
      }
      node = walker.nextNode(); off = 0;
    }
  }
  return { sources: meta.sources || [] };
}
