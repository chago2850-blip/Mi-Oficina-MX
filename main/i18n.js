// Traducción de menús y diálogos del proceso principal (mismo diccionario que la interfaz). GPL-3.0
const fs = require('fs');
const path = require('path');

let EN = { dict: {}, rules: [] };
try {
  const src = fs.readFileSync(path.join(__dirname, '..', 'apps', 'shared', 'i18n-en.js'), 'utf8');
  EN = JSON.parse(src.replace(/^[\s\S]*?export default\s*/, '').replace(/;\s*$/, ''));
} catch (e) { console.error('i18n', e.message); }
const RULES = EN.rules.map(([re, out]) => [new RegExp(re), out]);
let lang = 'es';

function tr(s) {
  if (EN.dict[s] !== undefined) return EN.dict[s];
  const tail = /^(.*?)(…|:)$/.exec(s);
  if (tail && EN.dict[tail[1]] !== undefined) return EN.dict[tail[1]] + tail[2];
  for (const [re, out] of RULES) { const r = re.exec(s); if (r) return out.replace(/\$(\d)/g, (_, i) => { const v = r[Number(i)] ?? ''; return EN.dict[v] ?? v; }); }
  return s;
}
const t = (s) => (lang === 'en' && typeof s === 'string' && s ? tr(s) : s);
// Traduce las etiquetas de una plantilla de menú (recursivo)
function tMenu(items) { return items.map((it) => ({ ...it, ...(it.label ? { label: t(it.label) } : {}), ...(it.submenu && Array.isArray(it.submenu) ? { submenu: tMenu(it.submenu) } : {}) })); }

module.exports = {
  t, tMenu,
  setLang: (l) => { lang = l === 'en' ? 'en' : 'es'; },
  getLang: () => lang,
  // 'auto' sigue el idioma del sistema: español si el sistema está en español; si no, inglés
  resolve: (pref, locale) => (pref === 'es' || pref === 'en' ? pref : (/^es\b/i.test(locale || '') ? 'es' : 'en')),
};

// Traduce automáticamente menús y diálogos nativos (así no hay que tocar cada llamada)
module.exports.patchElectron = function patchElectron({ Menu, dialog }) {
  const build = Menu.buildFromTemplate.bind(Menu);
  Menu.buildFromTemplate = (tpl) => build(tMenu(tpl));
  const tOpts = (o) => {
    if (!o || typeof o !== 'object') return o;
    const r = { ...o };
    for (const k of ['title', 'message', 'detail', 'buttonLabel', 'checkboxLabel']) if (typeof r[k] === 'string') r[k] = t(r[k]);
    if (Array.isArray(r.buttons)) r.buttons = r.buttons.map(t);
    if (Array.isArray(r.filters)) r.filters = r.filters.map((f) => ({ ...f, name: t(f.name) }));
    return r;
  };
  for (const k of ['showMessageBox', 'showMessageBoxSync', 'showOpenDialog', 'showOpenDialogSync', 'showSaveDialog', 'showSaveDialogSync']) {
    const orig = dialog[k].bind(dialog);
    dialog[k] = (...args) => orig(...args.map((a) => (a && typeof a === 'object' && !a.webContents ? tOpts(a) : a)));
  }
};
// Agrega ?lang= a una URL app://
module.exports.withLang = (url) => `${url}${url.includes('?') ? '&' : '?'}lang=${lang}`;
