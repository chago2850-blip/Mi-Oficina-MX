// Fuentes de la suite: las de Office (con su equivalente libre de LibreOffice, para que un .docx o
// .xlsx se vea igual aunque el equipo no tenga Calibri o Arial) y las fuentes libres incluidas en
// apps/shared/fonts (ver scripts/fonts/build.py). GPL-3.0
import BUNDLED from './fonts.json';
import { t } from './i18n.js';

// letras de Microsoft Office → la equivalente libre (mismas medidas cuando existe)
export const OFFICE = [
  ['Calibri', 'Carlito', 'sans-serif'], ['Calibri Light', 'Carlito', 'sans-serif'], ['Cambria', 'Caladea', 'serif'],
  ['Arial', "'Liberation Sans'", 'sans-serif'], ['Helvetica', "'Liberation Sans'", 'sans-serif'], ['Times New Roman', "'Liberation Serif'", 'serif'],
  ['Courier New', "'Liberation Mono'", 'monospace'], ['Georgia', 'Gelasio', 'serif'], ['Garamond', "'EB Garamond'", 'serif'],
  ['Verdana', "'DejaVu Sans'", 'sans-serif'], ['Tahoma', "'DejaVu Sans'", 'sans-serif'], ['Trebuchet MS', "'Source Sans Pro'", 'sans-serif'],
  ['Segoe UI', "'Noto Sans'", 'sans-serif'], ['Book Antiqua', "'Linux Libertine G'", 'serif'], ['Palatino Linotype', "'Linux Libertine G'", 'serif'],
  ['Century Gothic', 'Montserrat', 'sans-serif'], ['Consolas', "'Source Code Pro'", 'monospace'], ['Comic Sans MS', "'Comic Neue'", 'sans-serif'],
];
const bundledNames = new Set(BUNDLED.map((f) => f.name));
// pila CSS para un nombre de letra (lo que se escribe en font-family)
export function fontStack(name) {
  if (!name) return "Carlito, 'Liberation Sans', sans-serif";
  const n = String(name).replace(/^["']|["']$/g, '');
  const o = OFFICE.find((f) => f[0].toLowerCase() === n.toLowerCase());
  if (o) return `'${o[0]}', ${o[1]}, ${o[2]}`;
  const b = BUNDLED.find((f) => f.name.toLowerCase() === n.toLowerCase());
  if (b) return b.stack;
  return `'${n}', Carlito, 'Liberation Sans', sans-serif`;
}
// lista para los menús: [nombre, pila CSS], primero las de Office más usadas y luego todas las libres
const TOP_OFFICE = ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Garamond', 'Courier New', 'Segoe UI', 'Century Gothic', 'Consolas', 'Comic Sans MS'];
export const FONT_LIST = [...TOP_OFFICE.map((n) => [n, fontStack(n)]), ...BUNDLED.filter((f) => !TOP_OFFICE.includes(f.name)).map((f) => [f.name, f.stack])];
export const FONT_NAMES = FONT_LIST.map((f) => f[0]);
export const isBundled = (n) => bundledNames.has(n);
// cuando terminan de cargar las letras, las apps que dibujan en lienzo vuelven a dibujar
export function onFontsReady(fn) { try { document.fonts.addEventListener('loadingdone', () => fn()); document.fonts.ready.then(() => fn()); } catch { /* sin FontFaceSet */ } }
// precarga las letras que aparecen en un documento (para el lienzo de Calcula y para imprimir)
export function preload(names) { try { for (const n of new Set(names)) if (n) { document.fonts.load(`16px ${fontStack(n)}`); document.fonts.load(`bold 16px ${fontStack(n)}`); } } catch { /* nada */ } }

// ---------------------------------------------------------------- letras del sistema operativo
// Se piden al proceso principal (fontconfig, NSFontManager o el registro de Windows) y se agregan
// al final de FONT_LIST / FONT_NAMES (las mismas listas, para que los menús ya creados las vean).
export const SYSTEM_FONTS = [];
let sysPromise = null;
const sysListeners = [];
export function onSystemFonts(fn) { sysListeners.push(fn); if (SYSTEM_FONTS.length) fn(SYSTEM_FONTS.slice()); loadSystemFonts(); }
export function loadSystemFonts() {
  if (sysPromise) return sysPromise;
  const api = typeof window !== 'undefined' ? window.oficina : null;
  if (!api?.systemFonts) { sysPromise = Promise.resolve([]); return sysPromise; }
  sysPromise = api.systemFonts().then((list) => {
    const have = new Set(FONT_NAMES.map((n) => n.toLowerCase()));
    const added = [];
    for (const n of list || []) {
      if (typeof n !== 'string' || have.has(n.toLowerCase())) continue;
      have.add(n.toLowerCase()); added.push(n);
      SYSTEM_FONTS.push(n); FONT_LIST.push([n, fontStack(n)]); FONT_NAMES.push(n);
    }
    if (added.length) for (const fn of sysListeners) { try { fn(added); } catch { /* nada */ } }
    return added;
  }).catch(() => []);
  return sysPromise;
}
export const isSystemFont = (n) => SYSTEM_FONTS.some((x) => x.toLowerCase() === String(n || '').toLowerCase());
// agrega opciones nuevas a un <select> de letras ya dibujado
export function addFontOptions(select, names) {
  if (!select) return;
  const have = new Set([...select.options].map((o) => o.value.toLowerCase()));
  let grp = select.querySelector('optgroup[data-sys]');
  for (const n of names) {
    if (have.has(n.toLowerCase())) continue;
    if (!grp) { grp = document.createElement('optgroup'); grp.dataset.sys = '1'; grp.label = t('Letras del sistema'); select.append(grp); }
    const o = document.createElement('option'); o.value = n; o.textContent = n; grp.append(o);
  }
}
