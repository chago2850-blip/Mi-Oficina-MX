// Activa la interfaz en inglés en las páginas que no usan el armazón de la suite (dona y PapelPDF). GPL-3.0
import { installTranslator, setLang } from './i18n.js';

installTranslator({ skip: document.documentElement.dataset.i18nSkip || '' });
const bridge = window.launcher || window.papel;
bridge?.prefs?.().then((p) => { if (p?.uiLang) setLang(p.uiLang); }).catch(() => {});
bridge?.onPrefsChanged?.((p) => { if (p?.uiLang) setLang(p.uiLang); });
