// Formato de números estilo hoja de cálculo y reconocimiento de lo que se escribe. GPL-3.0
import { getLang } from '../../shared/i18n.js';
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const DIAS = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
const EPOCH = Date.UTC(1899, 11, 30);

export const serialToDate = (n) => new Date(EPOCH + Math.round(n * 86400000));
export const dateToSerial = (d) => (Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds()) - EPOCH) / 86400000;

export const FORMATS = [
  ['General', 'General'],
  ['0.00', 'Número'],
  ['#,##0', 'Número con separador de miles'],
  ['#,##0.00', 'Número con miles y 2 decimales'],
  ['"$"#,##0.00', 'Moneda ($)'],
  ['"$"#,##0.00;[Red]-"$"#,##0.00', 'Moneda, negativos en rojo'],
  ['0%', 'Porcentaje'],
  ['0.00%', 'Porcentaje con 2 decimales'],
  ['dd/mm/yyyy', 'Fecha corta (31/12/2026)'],
  ['d "de" mmmm "de" yyyy', 'Fecha larga (31 de diciembre de 2026)'],
  ['hh:mm', 'Hora (13:30)'],
  ['dd/mm/yyyy hh:mm', 'Fecha y hora'],
  ['0.00E+00', 'Científico'],
  ['# ?/?', 'Fracción'],
  ['@', 'Texto'],
];

function groupThousands(intStr) { return intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ','); }

function formatNumberSection(v, sec) {
  // Quita colores/condiciones
  let color = null;
  sec = sec.replace(/\[(Red|Rojo|Blue|Azul|Green|Verde|Black|Negro)\]/gi, (_, c) => { color = { red: '#C62828', rojo: '#C62828', blue: '#1565C0', azul: '#1565C0', green: '#2E7D32', verde: '#2E7D32', black: '#000', negro: '#000' }[c.toLowerCase()]; return ''; }).replace(/\[[^\]]*\]/g, '');
  // texto literal
  const lit = [];
  sec = sec.replace(/"([^"]*)"|\\(.)/g, (_, a, b) => { lit.push(a ?? b); return `\u0001${String.fromCharCode(0xE000 + lit.length - 1)}\u0001`; });
  const pct = (sec.match(/%/g) || []).length;
  if (pct) v *= 100 ** pct;
  if (/E\+0+/i.test(sec)) {
    const dec = (/\.(0+)E/i.exec(sec) || [, ''])[1].length;
    const s = Math.abs(v).toExponential(dec).toUpperCase().replace(/E([+-])(\d)$/, 'E$10$2');
    return { text: sec.replace(/[#0.,]+E\+0+/i, s).replace(/\u0001([\uE000-\uE0FF])\u0001/g, (_, i) => lit[i.charCodeAt(0) - 0xE000]), color };
  }
  if (/\?\/\?/.test(sec)) {
    const whole = Math.trunc(Math.abs(v)); const frac = Math.abs(v) - whole;
    let best = [0, 1, 1];
    for (let d = 1; d <= 9; d++) { const n = Math.round(frac * d); const err = Math.abs(frac - n / d); if (err < best[2]) best = [n, d, err]; }
    const t = best[0] ? `${whole ? `${whole} ` : ''}${best[0]}/${best[1]}` : String(whole);
    return { text: t, color };
  }
  const m = /([#0,]*)(\.([0#]*))?/.exec(sec.replace(/\u0001[\uE000-\uE0FF]\u0001/g, ''));
  const numPat = /[#0][#0,]*(\.[0#]*)?|\.[0#]+/.exec(sec);
  if (!numPat) return { text: sec.replace(/\u0001([\uE000-\uE0FF])\u0001/g, (_, i) => lit[i.charCodeAt(0) - 0xE000]), color };
  const decPart = numPat[1] ? numPat[1].slice(1) : (numPat[0].startsWith('.') ? numPat[0].slice(1) : '');
  const minDec = (decPart.match(/0/g) || []).length, maxDec = decPart.length;
  const thousands = numPat[0].includes(',');
  const intZeros = ((numPat[0].split('.')[0]).match(/0/g) || []).length;
  let s = Math.abs(v).toFixed(maxDec);
  if (maxDec > minDec) s = s.replace(/0+$/, (z) => z.slice(0, Math.max(0, z.length - (maxDec - minDec)))).replace(/\.$/, minDec ? '.' : '');
  let [ip, dp] = s.split('.');
  if (ip === '0' && intZeros === 0) ip = '';
  ip = ip.padStart(intZeros, '0');
  if (thousands) ip = groupThousands(ip);
  const num = dp ? `${ip}.${dp}` : ip;
  void m;
  return { text: sec.replace(numPat[0], num).replace(/\u0001([\uE000-\uE0FF])\u0001/g, (_, i) => lit[i.charCodeAt(0) - 0xE000]), color };
}

function formatDate(v, fmt) {
  const d = serialToDate(v);
  const Y = d.getUTCFullYear(), M = d.getUTCMonth(), D = d.getUTCDate(), W = d.getUTCDay();
  const hh = d.getUTCHours(), mi = d.getUTCMinutes(), ss = d.getUTCSeconds();
  const ampm = /AM\/PM|a\.m\.\/p\.m\./i.test(fmt);
  const lit = [];
  let f = fmt.replace(/\[[^\]]*\]/g, '').replace(/"([^"]*)"|\\(.)/g, (_, a, b) => { lit.push(a ?? b); return `\u0001${String.fromCharCode(0xE000 + lit.length - 1)}\u0001`; });
  // "m" es minuto si va después de h o antes de s
  f = f.replace(/(h+[^a-z]*)(mm?)/gi, (_, a, b) => a + (b.length === 2 ? '\u0002MM' : '\u0002M')).replace(/(?<![\u0002m])(mm?)([^a-z]*s)/gi, (_, a, b) => (a.length === 2 ? '\u0002MM' : '\u0002M') + b);
  const h12 = ((hh + 11) % 12) + 1;
  const tok = {
    yyyy: String(Y), yy: String(Y).slice(-2), mmmmm: MESES[M][0].toUpperCase(), mmmm: MESES[M], mmm: MESES[M].slice(0, 3), mm: String(M + 1).padStart(2, '0'), m: String(M + 1),
    dddd: DIAS[W], ddd: DIAS[W].slice(0, 3), dd: String(D).padStart(2, '0'), d: String(D),
    hh: String(ampm ? h12 : hh).padStart(2, '0'), h: String(ampm ? h12 : hh), ss: String(ss).padStart(2, '0'), s: String(ss),
  };
  let out = f.replace(/\u0002MM/g, String(mi).padStart(2, '0')).replace(/\u0002M/g, String(mi))
    .replace(/AM\/PM/gi, hh < 12 ? 'a. m.' : 'p. m.').replace(/a\.m\.\/p\.m\./gi, hh < 12 ? 'a. m.' : 'p. m.')
    .replace(/yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s/gi, (t) => tok[t.toLowerCase()] ?? t);
  out = out.replace(/\u0001([\uE000-\uE0FF])\u0001/g, (_, i) => lit[i.charCodeAt(0) - 0xE000]);
  return out;
}

export const isDateFormat = (fmt) => !!fmt && /(^|[^"\\])[dmyhs]/i.test(fmt.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '')) && !/^General$/i.test(fmt) && !/[0#]/.test(fmt.replace(/"[^"]*"/g, ''));

// Devuelve { text, color, align } para mostrar el valor
export function formatValue(v, fmt) {
  if (v === null || v === undefined || v === '') return { text: '' };
  if (typeof v === 'boolean') return { text: getLang() === 'en' ? (v ? 'TRUE' : 'FALSE') : (v ? 'VERDADERO' : 'FALSO'), align: 'center' };
  if (typeof v === 'string') return { text: v, align: 'left' };
  if (typeof v !== 'number') return { text: String(v) };
  if (!fmt || fmt === 'General') {
    if (!Number.isFinite(v)) return { text: '#¡NUM!', align: 'center' };
    let s;
    const a = Math.abs(v);
    if (a !== 0 && (a >= 1e11 || a < 1e-9)) s = v.toExponential(5).replace(/\.?0+e/, 'E').replace(/E([+-])(\d)$/, 'E$10$2').toUpperCase();
    else s = String(Math.round(v * 1e10) / 1e10);
    if (s.length > 11 && !s.includes('E')) s = String(Number(v.toPrecision(10)));
    return { text: s, align: 'right' };
  }
  if (fmt === '@') return { text: String(v), align: 'left' };
  const secs = fmt.split(/;(?=(?:[^"]*"[^"]*")*[^"]*$)/);
  if (isDateFormat(secs[0])) return { text: formatDate(v, secs[0]), align: 'right' };
  let sec = secs[0];
  let val = v;
  if (v < 0 && secs[1] !== undefined) { sec = secs[1]; val = -v; }
  else if (v === 0 && secs[2] !== undefined) sec = secs[2];
  const r = formatNumberSection(val, sec);
  if (v < 0 && secs[1] === undefined) r.text = `-${r.text}`;
  return { text: r.text, color: r.color, align: 'right' };
}

// Interpreta lo que el usuario escribe: devuelve { value, fmt? }
export function parseInput(raw) {
  const s = String(raw);
  const t = s.trim();
  if (t === '') return { value: null };
  if (t.startsWith('=')) return { value: t };
  if (t.startsWith("'")) return { value: t.slice(1), fmt: '@', text: true };
  const up = t.toUpperCase();
  if (up === 'VERDADERO' || up === 'TRUE') return { value: true };
  if (up === 'FALSO' || up === 'FALSE') return { value: false };
  let m;
  if ((m = /^(-)?\$\s?(-)?([\d,]*\.?\d+)$/.exec(t))) { const n = Number(m[3].replace(/,/g, '')); if (!Number.isNaN(n)) return { value: (m[1] || m[2] ? -1 : 1) * n, fmt: '"$"#,##0.00' }; }
  if ((m = /^(-?[\d,]*\.?\d+)\s?%$/.exec(t))) { const n = Number(m[1].replace(/,/g, '')); if (!Number.isNaN(n)) return { value: n / 100, fmt: Number.isInteger(n) ? '0%' : '0.00%' }; }
  if (/^-?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$/.test(t)) { const n = Number(t.replace(/,/g, '')); return { value: n, fmt: t.includes(',') ? (t.includes('.') ? '#,##0.00' : '#,##0') : undefined }; }
  if (/^-?\d*\.?\d+(e[+-]?\d+)?$/i.test(t)) return { value: Number(t) };
  if ((m = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t))) {
    let y = Number(m[3]); if (y < 100) y += y < 50 ? 2000 : 1900;
    const D = Number(m[1]), M = Number(m[2]);
    if (M >= 1 && M <= 12 && D >= 1 && D <= 31) {
      const ser = (Date.UTC(y, M - 1, D, Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0)) - EPOCH) / 86400000;
      return { value: ser, fmt: m[4] ? 'dd/mm/yyyy hh:mm' : 'dd/mm/yyyy' };
    }
  }
  if ((m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t))) return { value: (Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3] || 0)) / 86400, fmt: m[3] ? 'hh:mm:ss' : 'hh:mm' };
  return { value: s };
}

// Cambia la cantidad de decimales de un formato
export function changeDecimals(fmt, delta, sample) {
  if (!fmt || fmt === 'General') {
    const dec = sample != null && Number.isFinite(sample) ? Math.min(10, (String(sample).split('.')[1] || '').length) : 0;
    fmt = dec ? `0.${'0'.repeat(dec)}` : '0';
  }
  if (isDateFormat(fmt)) return fmt;
  return fmt.split(';').map((sec) => sec.replace(/([0#])(\.0*)?(?=[^0#.]*$|%|E|\s|"|;|$)/, (all, d, dec) => {
    const n = Math.max(0, (dec ? dec.length - 1 : 0) + delta);
    return d + (n ? `.${'0'.repeat(n)}` : '');
  })).join(';');
}
