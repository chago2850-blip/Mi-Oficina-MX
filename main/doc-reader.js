// Lector de documentos de Word 97-2003 (.doc) con formato básico. GPL-3.0
// Usa word-extractor para abrir el contenedor OLE y la tabla de piezas, y lee aquí
// las propiedades de párrafo (estilos de título, alineación, tablas, listas) y de carácter
// (negrita, cursiva, subrayado, tachado, tamaño, color, superíndice) de las páginas FKP.
const WordOleExtractor = require('word-extractor/lib/word-ole-extractor');
const BufferReader = require('word-extractor/lib/buffer-reader');

function* sprms(buf, start, end) {
  let p = start;
  while (p + 2 <= end) {
    const sprm = buf.readUInt16LE(p); p += 2;
    const spra = sprm >> 13;
    let len;
    if (spra === 0 || spra === 1) len = 1; else if (spra === 2 || spra === 4 || spra === 5) len = 2; else if (spra === 3) len = 4; else if (spra === 7) len = 3;
    else if (sprm === 0xD608 || sprm === 0xD606) { len = buf.readUInt16LE(p) + 2 - 1; p += 1; } else { len = buf[p] + 1; }
    if (p + len > end + 1) return;
    yield { sprm, p, len };
    p += len;
  }
}
// Lista de tramos [fcInicio, fcFin, datos] de una tabla PlcfBte (PAPX o CHPX)
function readFkps(word, table, fc, lcb, kind) {
  const runs = [];
  if (!lcb) return runs;
  const n = (lcb - 4) / 8;
  for (let k = 0; k < n; k++) {
    const pn = table.readUInt32LE(fc + 4 * (n + 1) + 4 * k) & 0x3FFFFF;
    const base = pn * 512; if (base + 512 > word.length) continue;
    const page = word.subarray(base, base + 512);
    const crun = page[511];
    for (let i = 0; i < crun; i++) {
      const a = page.readUInt32LE(4 * i), b = page.readUInt32LE(4 * (i + 1));
      const props = {};
      if (kind === 'pap') {
        const bOff = page[4 * (crun + 1) + 13 * i];
        props.istd = 0;
        if (bOff) {
          let off = bOff * 2; let cb = page[off]; let len; let st;
          if (cb === 0) { cb = page[off + 1]; len = 2 * cb; st = off + 2; } else { len = 2 * cb - 1; st = off + 1; }
          if (st + 2 <= 512) {
            props.istd = page.readUInt16LE(st);
            for (const s of sprms(page, st + 2, Math.min(512, st + len))) {
              const v = page[s.p];
              if (s.sprm === 0x2403 || s.sprm === 0x2461) props.jc = v;
              else if (s.sprm === 0x2416) props.inTable = v === 1;
              else if (s.sprm === 0x2417) props.ttp = v === 1;
              else if (s.sprm === 0x460B) props.ilfo = page.readUInt16LE(s.p);
              else if (s.sprm === 0x260A) props.ilvl = v;
            }
          }
        }
      } else {
        const bOff = page[4 * (crun + 1) + i];
        if (bOff) {
          const off = bOff * 2; const cb = page[off];
          for (const s of sprms(page, off + 1, Math.min(512, off + 1 + cb))) {
            const v = page[s.p]; const on = v === 1 || v === 0x81;
            if (s.sprm === 0x0835) props.b = on; else if (s.sprm === 0x0836) props.i = on; else if (s.sprm === 0x0837) props.s = on;
            else if (s.sprm === 0x2A3E) props.u = v !== 0; else if (s.sprm === 0x4A43) props.hps = page.readUInt16LE(s.p);
            else if (s.sprm === 0x2A48) props.iss = v; else if (s.sprm === 0x6870) { const r = page[s.p], g = page[s.p + 1], bl = page[s.p + 2], f = page[s.p + 3]; if (f === 0) props.color = `#${[r, g, bl].map((x) => x.toString(16).padStart(2, '0')).join('')}`; }
          }
        }
      }
      runs.push([a, b, props]);
    }
  }
  return runs.sort((x, y) => x[0] - y[0]);
}
function lookup(runs, fc) {
  let lo = 0, hi = runs.length - 1;
  while (lo <= hi) { const m = (lo + hi) >> 1; if (fc < runs[m][0]) hi = m - 1; else if (fc >= runs[m][1]) lo = m + 1; else return runs[m][2]; }
  return null;
}
function readStyles(table, fc, lcb) {
  const out = [];
  if (!lcb) return out;
  const st = table.subarray(fc, fc + lcb);
  const cbStshi = st.readUInt16LE(0); const cstd = st.readUInt16LE(2); const cbBase = st.readUInt16LE(4);
  let pos = 2 + cbStshi;
  for (let i = 0; i < cstd && pos + 2 <= st.length; i++) {
    const cb = st.readUInt16LE(pos);
    if (cb > 0) {
      const std = st.subarray(pos + 2, pos + 2 + cb);
      const sti = std.readUInt16LE(0) & 0x0FFF;
      let name = '';
      try { const cch = std.readUInt16LE(cbBase); name = std.subarray(cbBase + 2, cbBase + 2 + cch * 2).toString('ucs2'); } catch { /* */ }
      out[i] = { sti, name };
    }
    pos += 2 + cb;
  }
  return out;
}
const esc = (s) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

async function readDocRich(buf) {
  const ex = new WordOleExtractor();
  let cap = null;
  const orig = ex.writePieces.bind(ex);
  ex.writePieces = (b, t) => { orig(b, t); cap = { word: b, table: t, pieces: ex._pieces.map((p) => ({ ...p })) }; };
  const doc = await ex.extract(new BufferReader(buf));
  const plain = { body: doc.getBody(), headers: doc.getHeaders({ includeFooters: false }), footers: doc.getFooters(), footnotes: doc.getFootnotes(), endnotes: doc.getEndnotes() };
  if (!cap) return plain;
  try {
    const { word, table, pieces } = cap;
    const ccpText = word.readUInt32LE(0x004C);
    const styles = readStyles(table, word.readUInt32LE(0x00A2), word.readUInt32LE(0x00A6));
    const chpx = readFkps(word, table, word.readUInt32LE(0x00FA), word.readUInt32LE(0x00FE), 'chp');
    const papx = readFkps(word, table, word.readUInt32LE(0x0102), word.readUInt32LE(0x0106), 'pap');
    // caracteres del texto principal con su posición en el archivo
    const chars = [];
    for (const p of pieces) {
      for (let k = 0; k < p.text.length && p.startCp + k < ccpText; k++) chars.push([p.text[k], p.startFilePos + k * p.bpc]);
    }
    // párrafos
    const paras = []; let cur = [];
    for (const c of chars) { if (c[0] === '\r' || c[0] === '\x07') { paras.push({ end: c, chars: cur }); cur = []; } else cur.push(c); }
    if (cur.length) paras.push({ end: null, chars: cur });
    const html = []; let row = null; let rowsHtml = []; let list = null;
    const closeList = () => { if (list) { html.push(`<ul>${list.join('')}</ul>`); list = null; } };
    const flushTable = () => { if (rowsHtml.length) { html.push(`<table><tbody>${rowsHtml.join('')}</tbody></table>`); rowsHtml = []; } };
    for (const para of paras) {
      const pp = (para.end && lookup(papx, para.end[1])) || {};
      // texto con formato (los campos se reducen a su resultado)
      let out = ''; let field = 0; let showRes = [];
      let run = null, runText = '';
      const flush = () => {
        if (!runText) return;
        let t = esc(runText).replace(/\x0B/g, '<br>');
        if (run) {
          const st = [];
          if (run.hps && run.hps !== 24 && run.hps !== 22 && run.hps !== 20) st.push(`font-size:${run.hps / 2}pt`);
          if (run.color && run.color !== '#000000') st.push(`color:${run.color}`);
          if (st.length) t = `<span style="${st.join(';')}">${t}</span>`;
          if (run.iss === 1) t = `<sup>${t}</sup>`; else if (run.iss === 2) t = `<sub>${t}</sub>`;
          if (run.s) t = `<s>${t}</s>`; if (run.u) t = `<u>${t}</u>`; if (run.i) t = `<i>${t}</i>`; if (run.b) t = `<b>${t}</b>`;
        }
        out += t; runText = '';
      };
      for (const [ch, fc] of para.chars) {
        if (ch === '\x13') { field++; showRes.push(false); continue; }
        if (ch === '\x14') { showRes[showRes.length - 1] = true; continue; }
        if (ch === '\x15') { field = Math.max(0, field - 1); showRes.pop(); continue; }
        if (field && !showRes[showRes.length - 1]) continue;
        if (ch === '\x0C') { flush(); out += '\u0000PB\u0000'; continue; }
        if (/[\x00-\x08\x0E-\x1D\x1F]/.test(ch)) continue;
        const props = lookup(chpx, fc) || {};
        const key = JSON.stringify(props);
        if (!run || key !== run.key) { flush(); run = { ...props, key }; }
        runText += ch === '\x1E' ? '-' : ch;
      }
      flush();
      const parts = out.split('\u0000PB\u0000');
      for (let k = 1; k < parts.length; k++) { closeList(); flushTable(); html.push('<div class="pb" contenteditable="false"></div>'); }
      out = parts[parts.length - 1];
      // tablas: cada celda termina con \x07; la marca de fin de fila tiene ttp
      if (pp.inTable || para.end?.[0] === '\x07') {
        closeList();
        if (pp.ttp) { if (row) rowsHtml.push(`<tr>${row.join('')}</tr>`); row = null; continue; }
        row = row || []; row.push(`<td>${out || '<br>'}</td>`);
        continue;
      }
      if (row) { rowsHtml.push(`<tr>${row.join('')}</tr>`); row = null; }
      flushTable();
      const st = styles[pp.istd] || { sti: 0, name: '' };
      const lvl = st.sti >= 1 && st.sti <= 9 ? st.sti : (/^(heading|título|titulo|encabezado)\s*(\d)$/i.exec(st.name)?.[2] | 0);
      const align = { 1: 'center', 2: 'right', 3: 'justify' }[pp.jc];
      const sty = align ? ` style="text-align:${align}"` : '';
      if (pp.ilfo) { (list = list || []).push(`<li${sty}>${out || '<br>'}</li>`); continue; }
      closeList();
      if (lvl) html.push(`<h${Math.min(lvl, 6)}${sty}>${out.replace(/<\/?b>/g, '') || '<br>'}</h${Math.min(lvl, 6)}>`);
      else if (st.sti === 62 || /^(title|título|titulo)$/i.test(st.name)) html.push(`<p class="title"${sty}>${out || '<br>'}</p>`);
      else if (st.sti === 74 || /^(subtitle|subtítulo|subtitulo)$/i.test(st.name)) html.push(`<p class="subtitle"${sty}>${out || '<br>'}</p>`);
      else html.push(`<p${sty}>${out || '<br>'}</p>`);
    }
    if (row) rowsHtml.push(`<tr>${row.join('')}</tr>`);
    closeList(); flushTable();
    while (html.length && html[html.length - 1] === '<p><br></p>') html.pop();
    return { ...plain, html: html.join('') };
  } catch (e) {
    console.warn('doc-reader', e.message);
    return plain;
  }
}
module.exports = { readDocRich };
