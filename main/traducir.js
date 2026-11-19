// Traducción con servicios en línea (cuando el equipo tiene Internet). Solo se envía el texto a
// traducir y los idiomas; nunca datos de la persona (ni correo ni nombre). GPL-3.0
//  - google: el servicio gratuito de Google Translate que usan translate-shell o Crow Translate
//  - mymemory: MyMemory (gratuito, hasta ~5 000 caracteres al día por equipo)
//  - deepl: DeepL API (requiere una clave; la gratuita sirve)
//  - libre: LibreTranslate (un servidor propio o público, con clave opcional)
const LIMITS = { google: 1800, mymemory: 450, deepl: 25000, libre: 4000 };

const doFetch = (url, opts) => {
  try { const { net } = require('electron'); if (net?.fetch) return net.fetch(url, opts); } catch { /* fuera de Electron */ }
  return fetch(url, opts);
};
async function json(url, opts = {}) {
  const ctl = new AbortController(); const tm = setTimeout(() => ctl.abort(), 30000);
  try {
    const r = await doFetch(url, { ...opts, signal: ctl.signal });
    const txt = await r.text();
    if (!r.ok) { const err = new Error(`HTTP ${r.status}: ${txt.slice(0, 200)}`); err.status = r.status; throw err; }
    return JSON.parse(txt);
  } finally { clearTimeout(tm); }
}

// idioma probable de un texto (para MyMemory, que no detecta el idioma de origen)
const HINTS = {
  es: /\b(el|la|los|las|de|que|y|en|un|una|por|con|para|es|del|se|no|al|lo|como|más|pero|sus|le|ya|o|este|sí|porque|esta|entre|cuando|muy|sin|sobre)\b/gi,
  en: /\b(the|and|of|to|in|is|that|it|for|was|on|are|with|as|this|be|at|by|have|from|or|not|but|which|you|they|we|an|were)\b/gi,
  pt: /\b(o|os|as|de|que|e|em|um|uma|para|com|não|na|no|do|da|dos|das|se|mais|por|como|mas|foi|ao|ele|isso|está|você)\b/gi,
  fr: /\b(le|la|les|de|des|et|est|en|un|une|du|que|qui|dans|pour|pas|sur|au|avec|il|elle|ce|ne|plus|par|sont|nous|vous)\b/gi,
  it: /\b(il|la|di|che|e|un|una|per|non|in|del|della|sono|con|si|da|lo|gli|le|questo|ma|come|anche|più|nel|alla)\b/gi,
  de: /\b(der|die|das|und|ist|nicht|ein|eine|zu|den|von|mit|sich|auf|für|im|dem|des|auch|es|an|als|wie|aber|wir|ich|sie)\b/gi,
};
function guessLang(text) {
  let best = 'es'; let n = 0;
  for (const [l, re] of Object.entries(HINTS)) { const c = (text.match(re) || []).length; if (c > n) { n = c; best = l; } }
  return best;
}

async function google(text, from, to) {
  const u = `https://translate.googleapis.com/translate_a/single?client=gtx&dt=t&sl=${encodeURIComponent(from || 'auto')}&tl=${encodeURIComponent(to)}&q=${encodeURIComponent(text)}`;
  const j = await json(u);
  return { text: (j[0] || []).map((s) => s[0] || '').join(''), detected: j[2] || null };
}
async function mymemory(text, from, to) {
  const src = !from || from === 'auto' ? guessLang(text) : from;
  if (src === to) return { text, detected: src };
  const j = await json(`https://api.mymemory.translated.net/get?q=${encodeURIComponent(text)}&langpair=${encodeURIComponent(`${src}|${to}`)}`);
  if (Number(j.responseStatus) !== 200) throw new Error(j.responseDetails || `MyMemory ${j.responseStatus}`);
  return { text: j.responseData?.translatedText || '', detected: src };
}
async function deepl(text, from, to, o) {
  if (!o.key) throw new Error('DeepL necesita una clave de API.');
  const host = /:fx$/.test(o.key) ? 'https://api-free.deepl.com' : 'https://api.deepl.com';
  const body = { text: [text], target_lang: to.toUpperCase() === 'EN' ? 'EN-US' : to.toUpperCase() === 'PT' ? 'PT-BR' : to.toUpperCase() };
  if (from && from !== 'auto') body.source_lang = from.toUpperCase();
  const j = await json(`${host}/v2/translate`, { method: 'POST', headers: { Authorization: `DeepL-Auth-Key ${o.key}`, 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  const t = j.translations?.[0] || {};
  return { text: t.text || '', detected: (t.detected_source_language || '').toLowerCase() || null };
}
async function libre(text, from, to, o) {
  const base = String(o.url || 'https://libretranslate.com').replace(/\/+$/, '');
  if (!/^https?:\/\//i.test(base)) throw new Error('Dirección de LibreTranslate no válida.');
  const j = await json(`${base}/translate`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ q: text, source: from || 'auto', target: to, format: 'text', ...(o.key ? { api_key: o.key } : {}) }) });
  if (j.error) throw new Error(j.error);
  return { text: j.translatedText || '', detected: j.detectedLanguage?.language || null };
}
const PROVIDERS = { google, mymemory, deepl, libre };

// parte una lista de textos en paquetes unidos por saltos de línea, sin pasar del límite
function pack(texts, limit) {
  const out = []; let cur = []; let len = 0;
  for (let i = 0; i < texts.length; i++) {
    const t = texts[i];
    if (cur.length && len + t.length + 1 > limit) { out.push(cur); cur = []; len = 0; }
    cur.push(i); len += t.length + 1;
  }
  if (cur.length) out.push(cur);
  return out;
}
// un texto más largo que el límite se parte por oraciones
function splitLong(t, limit) {
  if (t.length <= limit) return [t];
  const parts = t.match(/[^.!?¡¿…]+[.!?…]*\s*/g) || [t]; const out = []; let cur = '';
  for (const p of parts) {
    if ((cur + p).length > limit && cur) { out.push(cur); cur = ''; }
    if (p.length > limit) { for (let i = 0; i < p.length; i += limit) out.push(p.slice(i, i + limit)); } else cur += p;
  }
  if (cur) out.push(cur);
  return out;
}

// texts: lista de párrafos; devuelve { texts, detected }
async function translate({ provider = 'google', texts = [], from = 'auto', to = 'en', key = '', url = '' } = {}) {
  const fn = PROVIDERS[provider]; if (!fn) throw new Error(`Servicio desconocido: ${provider}`);
  if (!to) throw new Error('Falta el idioma de destino.');
  if (process.env.OFICINA_FAKE_TRANSLATE) return { texts: texts.map((t) => (t.trim() ? `[${to}] ${t}` : t)), detected: from === 'auto' ? guessLang(texts.join(' ')) : from };
  const limit = LIMITS[provider] || 1500;
  const clean = texts.map((t) => String(t ?? '').replace(/[\r\n]+/g, ' '));
  const result = clean.slice(); let detected = null;
  const one = async (t) => {
    if (!t.trim()) return t;
    const pieces = splitLong(t, limit); const outs = [];
    for (const p of pieces) { const r = await fn(p, from, to, { key, url }); detected = detected || r.detected; outs.push(r.text); }
    return outs.join(pieces.length > 1 ? ' ' : '').replace(/\s+$/, '') + (t.match(/\s+$/)?.[0] || '');
  };
  for (const group of pack(clean, limit)) {
    const items = group.filter((i) => clean[i].trim());
    if (!items.length) continue;
    if (items.length === 1 || provider === 'mymemory') { for (const i of items) result[i] = await one(clean[i]); continue; }
    const joined = items.map((i) => clean[i]).join('\n');
    const r = await fn(joined, from, to, { key, url }); detected = detected || r.detected;
    const lines = r.text.split(/\r?\n/);
    if (lines.length === items.length) items.forEach((i, k) => { result[i] = lines[k]; });
    else for (const i of items) result[i] = await one(clean[i]); // el servicio unió o partió líneas
  }
  return { texts: result, detected };
}

module.exports = { translate, guessLang, pack, splitLong };
