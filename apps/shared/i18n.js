// Interfaz en español e inglés para toda la suite. GPL-3.0
//
// La interfaz se escribe en español. En inglés, un traductor recorre el DOM y cambia los textos
// (y los atributos title, placeholder, aria-label) usando el diccionario de i18n-en.js. El contenido
// del documento (páginas, diapositivas, celdas) nunca se traduce. Cambiar de idioma no recarga la
// ventana: cada nodo recuerda su texto original.
import EN from './i18n-en.js';

const DICT = EN.dict;
const RULES = EN.rules.map(([re, out]) => [new RegExp(re), out]);

let lang = (() => {
  try { const q = new URLSearchParams(location.search).get('lang'); if (q === 'en' || q === 'es') return q; } catch { /* */ }
  return 'es';
})();
const listeners = new Set();
export const getLang = () => lang;
export const onLang = (fn) => { listeners.add(fn); return () => listeners.delete(fn); };

// Traduce una cadena de la interfaz (si el idioma es español, la devuelve igual)
export function t(s) {
  if (lang === 'es' || typeof s !== 'string' || !s) return s;
  return tr(s) ?? s;
}
// Igual que t(), pero con valores: tf('Página {0} de {1}', 2, 5)
export function tf(s, ...args) { return t(s).replace(/\{(\d+)\}/g, (_, i) => String(args[Number(i)] ?? '')); }

const missing = new Set();
const keys = (k) => k.replace(/Mayús/g, 'Shift').replace(/\bSupr\b/g, 'Del').replace(/\bInicio\b/g, 'Home').replace(/\bFin\b/g, 'End');
function tr(s) {
  const hit = DICT[s];
  if (hit !== undefined) return hit;
  const m = /^(\s*)([\s\S]*?)(\s*)$/.exec(s);
  if (m[2] !== s) { const core = tr(m[2]); return core == null ? null : m[1] + core + m[3]; }
  // "Texto…" y "Texto:" se traducen con la forma base
  const tail = /^(.*?)(…|:|\.\.\.)$/.exec(s);
  if (tail && DICT[tail[1]] !== undefined) return DICT[tail[1]] + tail[2];
  for (const [re, out] of RULES) {
    const r = re.exec(s);
    if (r) return out.replace(/\$(\d)/g, (_, i) => { const v = r[Number(i)] ?? ''; return tr(v) ?? v; });
  }
  // "Negrita (Ctrl+B)" -> "Bold (Ctrl+B)"
  const par = /^(.*?) \(([^()]*)\)$/.exec(s);
  if (par) { const a = tr(par[1]); if (a != null) return `${a} (${tr(par[2]) ?? keys(par[2])})`; }
  if (/[a-záéíóúñ]{2}/i.test(s) && /[áéíóúñ¿¡]|\b(de|la|el|los|las|una?|para|con|sin|del|que|por|en|se|es|al|y)\b/i.test(s)) missing.add(s);
  return null;
}
if (typeof window !== 'undefined') window.__i18nMissing = missing;

// Traduce los textos de un fragmento HTML (para las plantillas)
export function localizeHtml(html) {
  if (lang === 'es' || !html) return html;
  const tpl = document.createElement('template'); tpl.innerHTML = html;
  const w = document.createTreeWalker(tpl.content, NodeFilter.SHOW_TEXT);
  while (w.nextNode()) { const n = w.currentNode; const v = tr(n.nodeValue); if (v != null) n.nodeValue = v; }
  return tpl.innerHTML;
}
// Fecha larga en el idioma de la interfaz
export const longDate = (d = new Date()) => d.toLocaleDateString(lang === 'en' ? 'en-US' : 'es-MX', { day: 'numeric', month: 'long', year: 'numeric' });

// ------------------------------------------------------------------ traductor del DOM
const ATTRS = ['title', 'placeholder', 'aria-label', 'data-tip', 'alt'];
let skipSel = '';
const textOrig = new WeakMap(); // nodo de texto -> { es, shown }
const attrOrig = new WeakMap(); // elemento -> { attr: { es, shown } }
let observer = null;
let busy = false;

function skipped(el) { return !el || (skipSel && el.closest(skipSel)) || el.closest('script,style,textarea,[contenteditable="true"],[contenteditable=""],.no-i18n'); }
// Los atributos (title, placeholder…) de campos de texto sí se traducen
function attrSkipped(el) { return !el || (skipSel && el.closest(skipSel)) || el.closest('script,style,.no-i18n') || el.parentElement?.closest('[contenteditable="true"],[contenteditable=""]'); }

function doText(n) {
  const v = n.nodeValue;
  if (!v || !/[A-Za-zÁÉÍÓÚáéíóúÑñ¿¡]/.test(v)) return;
  if (skipped(n.parentElement)) return;
  let rec = textOrig.get(n);
  if (!rec || v !== rec.shown) { rec = { es: v, shown: v }; textOrig.set(n, rec); }
  const want = lang === 'en' ? (tr(rec.es) ?? rec.es) : rec.es;
  if (want !== v) { rec.shown = want; n.nodeValue = want; } else rec.shown = v;
}
function doAttrs(el) {
  for (const a of ATTRS) {
    if (!el.hasAttribute?.(a)) continue;
    const v = el.getAttribute(a);
    if (!v || !/[A-Za-zÁÉÍÓÚáéíóú]/.test(v)) continue;
    let map = attrOrig.get(el); if (!map) { map = {}; attrOrig.set(el, map); }
    let rec = map[a];
    if (!rec || v !== rec.shown) { rec = { es: v, shown: v }; map[a] = rec; }
    const want = lang === 'en' ? (tr(rec.es) ?? rec.es) : rec.es;
    rec.shown = want;
    if (want !== v) el.setAttribute(a, want);
  }
}
function walk(root) {
  if (root.nodeType === 3) { doText(root); return; }
  if (root.nodeType !== 1 && root.nodeType !== 9 && root.nodeType !== 11) return;
  if (root.nodeType === 1) { if (attrSkipped(root)) return; doAttrs(root); }
  const w = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, {
    acceptNode: (n) => (n.nodeType === 1 && attrSkipped(n) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  while (w.nextNode()) { const n = w.currentNode; if (n.nodeType === 3) doText(n); else doAttrs(n); }
}
function translateTitle() {
  const el = document.querySelector('title'); if (!el) return;
  const v = document.title; let rec = textOrig.get(el);
  if (!rec || v !== rec.shown) { rec = { es: v, shown: v }; textOrig.set(el, rec); }
  const want = lang === 'en' ? (tr(rec.es) ?? rec.es) : rec.es;
  rec.shown = want; if (want !== v) document.title = want;
}

// skip: selector de zonas que son contenido del usuario (páginas, diapositivas…)
export function installTranslator({ skip = '' } = {}) {
  skipSel = skip;
  document.documentElement.lang = lang;
  if (observer) return;
  observer = new MutationObserver((muts) => {
    if (busy) return;
    busy = true;
    try {
      for (const m of muts) {
        if (m.type === 'characterData') doText(m.target);
        else if (m.type === 'attributes') { if (!attrSkipped(m.target)) doAttrs(m.target); }
        else for (const n of m.addedNodes) walk(n);
      }
      translateTitle();
    } finally { busy = false; }
  });
  observer.observe(document.documentElement, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ATTRS });
  if (lang !== 'es') { busy = true; try { walk(document.documentElement); translateTitle(); } finally { busy = false; } }
}

export function setLang(l) {
  l = l === 'en' ? 'en' : 'es';
  if (l === lang) return;
  lang = l;
  document.documentElement.lang = lang;
  if (observer) { busy = true; try { walk(document.documentElement); translateTitle(); } finally { busy = false; } observer.takeRecords(); }
  for (const fn of listeners) { try { fn(lang); } catch (e) { console.error(e); } }
}
