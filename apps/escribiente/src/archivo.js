// Escribiente — Archivo como en LibreOffice Writer: exportar como EPUB (con opciones), como texto
// codificado (codificación y fin de línea) y como RTF; página actual para imprimir o exportar. GPL-3.0
import { h, t } from '../../shared/office.js';
import { tabbed, field, sel, chk } from '../../shared/dialogs.js';
import { htmlToEpub } from './epub.js';
import { htmlToRtf } from './rtf.js';
import { htmlToText } from './docx-in.js';

let C = null; // { state, editor, shell, settings: () => {} }
export function init(ctx) { C = ctx; }
let pending = null; // opciones del diálogo para el siguiente serialize()

export async function epubDialog() {
  const p = C.shell.docProps();
  const body = await tabbed('Exportar como EPUB', [{ label: 'General', body: h('div', { class: 'fx-grid' },
    field('Título', h('input', { name: 'title', value: p.title || '' }), 'wide'), field('Autor', h('input', { name: 'author', value: p.author || '' })),
    field('Idioma', sel('lang', document.documentElement.lang === 'en' ? 'en' : 'es', [['es', 'Español'], ['es-MX', 'Español (México)'], ['en', 'Inglés']])),
    field('Dividir en capítulos', sel('split', 'h1', [['h1', 'En cada título de nivel 1'], ['pb', 'En cada salto de página'], ['none', 'Un solo capítulo']])),
    chk('cover', 'Agregar una portada con el título y el autor', true)) }], { width: '520px' });
  if (!body) return null;
  const v = (n) => body.querySelector(`[name="${n}"]`);
  pending = { epub: { title: v('title').value.trim(), author: v('author').value.trim(), lang: v('lang').value, split: v('split').value, cover: v('cover').checked } };
  return C.shell.saveAsFormat('epub', 'Libro electrónico EPUB');
}

export async function txtDialog() {
  const body = await tabbed('Exportar como texto codificado', [{ label: 'Opciones', body: h('div', { class: 'fx-grid' },
    field('Codificación', sel('enc', 'utf-8', [['utf-8', 'Unicode (UTF-8)'], ['utf-8-bom', 'Unicode (UTF-8 con BOM)'], ['windows-1252', 'Europa occidental (Windows-1252)'], ['utf-16le', 'Unicode (UTF-16)']])),
    field('Fin de línea', sel('eol', /Win/.test(navigator.platform) ? 'crlf' : 'lf', [['crlf', 'Windows (CR+LF)'], ['lf', 'Linux y macOS (LF)'], ['cr', 'Mac clásico (CR)']]))) }], { width: '460px' });
  if (!body) return null;
  pending = { txt: { enc: body.querySelector('[name=enc]').value, eol: body.querySelector('[name=eol]').value } };
  return C.shell.saveAsFormat('txt', 'Texto codificado');
}

// Windows-1252: los 32 caracteres del intervalo 0x80–0x9F
const CP1252 = '€\u0081‚ƒ„…†‡ˆ‰Š‹Œ\u008DŽ\u008F\u0090‘’“”•–—˜™š›œ\u009DžŸ';
export function encodeText(text, enc = 'utf-8', eol = 'lf') {
  const s = text.replace(/\r\n?/g, '\n').replace(/\n/g, { crlf: '\r\n', lf: '\n', cr: '\r' }[eol] || '\n').replace(/­/g, '');
  if (enc === 'utf-8') return new TextEncoder().encode(s);
  if (enc === 'utf-8-bom') { const b = new TextEncoder().encode(s); const o = new Uint8Array(b.length + 3); o.set([0xEF, 0xBB, 0xBF]); o.set(b, 3); return o; }
  if (enc === 'utf-16le') { const o = new Uint8Array(2 + s.length * 2); o[0] = 0xFF; o[1] = 0xFE; for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); o[2 + i * 2] = c & 255; o[3 + i * 2] = c >> 8; } return o; }
  const o = new Uint8Array(s.length); let k = 0;
  for (const ch of s) {
    const c = ch.codePointAt(0);
    if (c < 0x80 || (c >= 0xA0 && c <= 0xFF)) o[k++] = c;
    else { const i = CP1252.indexOf(ch); o[k++] = i >= 0 ? 0x80 + i : 0x3F; }
  }
  return o.slice(0, k);
}

// para serialize(): EPUB, RTF y texto con las opciones elegidas
export async function serializeExtra(ext) {
  const settings = { ...C.settings(), ...C.shell.docProps() };
  const opt = pending; pending = null;
  if (ext === 'rtf') return htmlToRtf(C.editor, settings);
  if (ext === 'epub') return htmlToEpub(C.editor, { ...settings, ...(opt?.epub || { split: 'h1', cover: false }), tocTitle: t('Índice') });
  if (ext === 'txt' && opt?.txt) return encodeText(htmlToText(C.editor), opt.txt.enc, opt.txt.eol);
  return null;
}

// número de página (desde 1) donde está el cursor
export function currentPage() {
  const s = getSelection(); let n = s?.rangeCount ? s.getRangeAt(0).startContainer : null;
  if (!n || !C.editor.contains(n)) return 1;
  while (n && n.parentNode !== C.editor) n = n.parentNode;
  const i = [...C.editor.children].indexOf(n);
  return (C.state.pageOf?.[i] ?? 0) + 1;
}
