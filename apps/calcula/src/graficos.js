// Calcula Datos — gráficos como en LibreOffice Calc (parte 4): más tipos (apilados y 100 %, anillo,
// radar, burbujas, combinado con eje secundario), cuadro de propiedades (títulos, leyenda, etiquetas,
// ejes, colores, series en filas o columnas) y líneas de tendencia con ecuación y R². Se guardan como
// gráficos de verdad en .xlsx (DrawingML) y en .ods (objetos de gráfico de ODF), y se leen de vuelta
// desde archivos de Excel y de LibreOffice. GPL-3.0
import JSZip from 'jszip';
import { h, t, tf, toast } from '../../shared/office.js';
import { tabbed, field, numIn, sel, chk } from '../../shared/dialogs.js';
import { colName, colIndex, parseRange } from './workbook.js';

export const CHART_TYPES = [['column', 'Columnas'], ['bar', 'Barras'], ['line', 'Líneas'], ['area', 'Área'], ['pie', 'Circular'], ['doughnut', 'Anillo'], ['scatter', 'Dispersión (XY)'], ['bubble', 'Burbujas'], ['radar', 'Red (radar)'], ['combo', 'Columnas y líneas']];
export const PALETTE = ['#0E9C74', '#5B48E8', '#DC3A63', '#F08A24', '#1F8FD6', '#8E6CD9', '#2FB4A8', '#E0B000'];
const STACKABLE = ['column', 'bar', 'line', 'area', 'combo'];

// ================================================================ datos
// Analiza el intervalo: fila de encabezados, columna de categorías y series (en columnas o en filas)
export function chartLayout(wb, ch, s, g) {
  if (!g) return null;
  const val = (r, c) => wb.value(r, c, s);
  const byRows = !!ch.byRows;
  // en filas se trasponen los papeles de filas y columnas
  const R1 = byRows ? g.c1 : g.r1, R2 = byRows ? g.c2 : g.r2, C1 = byRows ? g.r1 : g.c1, C2 = byRows ? g.r2 : g.c2;
  const at = (i, j) => (byRows ? val(j, i) : val(i, j)); // i = «fila» lógica, j = «columna» lógica
  const addr = (i, j) => (byRows ? { r: j, c: i } : { r: i, c: j });
  const headerRow = ch.headers != null ? !!ch.headers : ([...Array(C2 - C1 + 1).keys()].some((k) => typeof at(R1, C1 + k) === 'string') && R2 > R1);
  const catCol = ch.catCol != null ? !!ch.catCol : ([...Array(R2 - R1 + 1).keys()].slice(headerRow ? 1 : 0).some((k) => typeof at(R1 + k, C1) === 'string') || (C2 > C1 && headerRow && at(R1, C1) == null));
  const i0 = headerRow ? R1 + 1 : R1, j0 = catCol ? C1 + 1 : C1;
  const rows = []; for (let i = i0; i <= R2; i++) if (byRows || !s.hiddenRows.has(i)) rows.push(i);
  const cats = rows.map((i) => (catCol ? wb.display(addr(i, C1).r, addr(i, C1).c, s).text : String(i - i0 + 1)));
  const series = [];
  for (let j = j0; j <= C2; j++) {
    const data = rows.map((i) => { const v = at(i, j); return typeof v === 'number' ? v : null; });
    const nameAt = headerRow ? addr(R1, j) : null;
    series.push({ name: headerRow ? wb.display(nameAt.r, nameAt.c, s).text : tf('Serie {0}', j - j0 + 1), data, nameAt, cells: rows.map((i) => addr(i, j)) });
  }
  return { cats, series, catCells: catCol ? rows.map((i) => addr(i, C1)) : null, byRows, headerRow, catCol };
}

// ================================================================ líneas de tendencia
export function trendFit(xs, ys, kind = 'linear', period = 2) {
  const P = xs.map((x, i) => [x, ys[i]]).filter(([x, y]) => typeof x === 'number' && typeof y === 'number' && (kind !== 'exp' || y > 0) && (kind !== 'log' || x > 0) && (kind !== 'power' || (x > 0 && y > 0)));
  if (P.length < 2) return null;
  if (kind === 'movavg') { const out = ys.map((_, i) => { if (i < period - 1) return null; const w = ys.slice(i - period + 1, i + 1); return w.some((v) => typeof v !== 'number') ? null : w.reduce((a, b) => a + b, 0) / period; }); return { kind, f: null, points: out, eq: tf('Media móvil ({0} períodos)', period) }; }
  const tx = (x) => (kind === 'log' || kind === 'power' ? Math.log(x) : x); const ty = (y) => (kind === 'exp' || kind === 'power' ? Math.log(y) : y);
  let f; let eq;
  if (kind === 'poly2') {
    // mínimos cuadrados de grado 2
    const S = (p, q) => P.reduce((a, [x, y]) => a + x ** p * y ** q, 0);
    const n = P.length; const M = [[n, S(1, 0), S(2, 0)], [S(1, 0), S(2, 0), S(3, 0)], [S(2, 0), S(3, 0), S(4, 0)]]; const B = [S(0, 1), S(1, 1), S(2, 1)];
    const det = (m) => m[0][0] * (m[1][1] * m[2][2] - m[1][2] * m[2][1]) - m[0][1] * (m[1][0] * m[2][2] - m[1][2] * m[2][0]) + m[0][2] * (m[1][0] * m[2][1] - m[1][1] * m[2][0]);
    const D = det(M); if (!D) return null;
    const co = [0, 1, 2].map((k) => det(M.map((row, r) => row.map((v, c) => (c === k ? B[r] : v)))) / D);
    f = (x) => co[0] + co[1] * x + co[2] * x * x; eq = `y = ${fmt(co[2])}x² ${sg(co[1])}x ${sg(co[0])}`;
  } else {
    const n = P.length; const X = P.map(([x]) => tx(x)); const Y = P.map(([, y]) => ty(y));
    const mx = X.reduce((a, b) => a + b, 0) / n; const my = Y.reduce((a, b) => a + b, 0) / n;
    const sxx = X.reduce((a, x) => a + (x - mx) ** 2, 0); if (!sxx) return null;
    const b = X.reduce((a, x, i) => a + (x - mx) * (Y[i] - my), 0) / sxx; const a0 = my - b * mx;
    if (kind === 'exp') { f = (x) => Math.exp(a0) * Math.exp(b * x); eq = `y = ${fmt(Math.exp(a0))}·e^(${fmt(b)}x)`; }
    else if (kind === 'log') { f = (x) => a0 + b * Math.log(x); eq = `y = ${fmt(b)}·ln(x) ${sg(a0)}`; }
    else if (kind === 'power') { f = (x) => Math.exp(a0) * x ** b; eq = `y = ${fmt(Math.exp(a0))}·x^${fmt(b)}`; }
    else { f = (x) => a0 + b * x; eq = `y = ${fmt(b)}x ${sg(a0)}`; }
  }
  const my = P.reduce((a, [, y]) => a + y, 0) / P.length;
  const ssr = P.reduce((a, [x, y]) => a + (y - f(x)) ** 2, 0); const sst = P.reduce((a, [, y]) => a + (y - my) ** 2, 0);
  const r2 = sst ? 1 - ssr / sst : 1;
  return { kind, f, eq, r2, points: xs.map((x) => (typeof x === 'number' && (kind !== 'log' && kind !== 'power' || x > 0) ? f(x) : null)) };
}
const fmt = (n) => (Math.abs(n) >= 1000 || Math.abs(n) < 0.01 ? n.toPrecision(4) : String(Math.round(n * 10000) / 10000));
const sg = (n) => (n < 0 ? `− ${fmt(-n)}` : `+ ${fmt(n)}`);

// ================================================================ opción de ECharts
export function buildOption(ch, d, { printing = false } = {}) {
  const type = ch.type || 'column';
  const pal = PALETTE.map((c, i) => ch.colors?.[i] || c);
  const legendPos = ch.legend || (d.series.length > 1 || ['pie', 'doughnut'].includes(type) ? 'bottom' : 'none');
  const legend = legendPos === 'none' ? undefined : { type: 'scroll', ...(legendPos === 'top' ? { top: ch.title ? 34 : 4 } : legendPos === 'left' ? { left: 4, top: 'middle', orient: 'vertical' } : legendPos === 'right' ? { right: 4, top: 'middle', orient: 'vertical' } : { bottom: 4 }) };
  const titleH = ch.title ? (ch.subtitle ? 52 : 36) : 0;
  const base = {
    color: pal, animation: !printing,
    title: ch.title ? { text: ch.title, subtext: ch.subtitle || '', left: 'center', top: 6, textStyle: { fontSize: 15, fontWeight: 600, color: '#1D2030', fontFamily: 'Inter, sans-serif' } } : undefined,
    tooltip: { trigger: ['pie', 'doughnut', 'scatter', 'bubble', 'radar'].includes(type) ? 'item' : 'axis' }, legend,
    grid: { left: legendPos === 'left' ? 110 : 50, right: legendPos === 'right' ? 110 : (ch.secondary?.length ? 50 : 20), top: titleH + 14 + (legendPos === 'top' ? 24 : 0), bottom: (legendPos === 'bottom' ? 40 : 18) + (ch.xTitle ? 18 : 0), containLabel: true },
  };
  const lbl = (pos) => (ch.labels && ch.labels !== 'none' ? { show: true, position: pos, formatter: ch.labels === 'percent' ? '{d}%' : ch.labels === 'category' ? '{b}' : ch.labels === 'both' ? '{b}: {c}' : undefined } : undefined);
  if (type === 'pie' || type === 'doughnut') {
    return { ...base, series: [{ type: 'pie', radius: type === 'doughnut' ? ['38%', '62%'] : ['0%', '62%'], center: ['50%', `${52 + titleH / 10}%`], data: d.cats.map((c, i) => ({ name: c, value: d.series[0]?.data[i] })), label: ch.labels === 'none' ? { show: false } : { formatter: { value: '{b}: {c}', category: '{b}', both: '{b}: {c} ({d}%)' }[ch.labels] || '{b}: {d}%' } }] };
  }
  if (type === 'radar') {
    const max = Math.max(1, ...d.series.flatMap((s) => s.data.filter((v) => typeof v === 'number')));
    return { ...base, radar: { indicator: d.cats.map((c) => ({ name: c, max: ch.yMax ?? max * 1.1 })), center: ['50%', '55%'], radius: '62%' }, series: [{ type: 'radar', data: d.series.map((s) => ({ name: s.name, value: s.data.map((v) => v ?? 0), areaStyle: ch.fillRadar ? { opacity: 0.2 } : undefined, label: lbl('top') })) }] };
  }
  const valAx = (extra = {}) => ({ type: ch.logScale ? 'log' : 'value', min: ch.yMin ?? undefined, max: ch.yMax ?? undefined, name: ch.yTitle || '', nameLocation: 'middle', nameGap: 42, axisLabel: { color: '#4B5066', formatter: ch.stacked === 'percent' ? '{value}%' : undefined }, splitLine: { show: ch.grid !== false, lineStyle: { color: '#ECEDF2' } }, minorSplitLine: { show: !!ch.minorGrid }, ...extra });
  const xTitle = { name: ch.xTitle || '', nameLocation: 'middle', nameGap: 28 };
  if (type === 'scatter' || type === 'bubble') {
    const xs = d.series[0]?.data || [];
    const ys = type === 'bubble' ? [d.series[1]].filter(Boolean) : d.series.slice(1);
    const size = d.series[2]?.data || [];
    const smax = Math.max(1, ...size.filter((v) => typeof v === 'number'));
    const series = ys.map((sr, k) => ({ type: 'scatter', name: sr.name, data: sr.data.map((v, i) => (type === 'bubble' ? [xs[i], v, size[i]] : [xs[i], v])), symbolSize: type === 'bubble' ? (p) => 8 + 42 * Math.sqrt((p[2] || 0) / smax) : 9, label: lbl('top') }));
    if (ch.scatterLines && type === 'scatter') series.forEach((s) => { s.type = 'line'; s.smooth = !!ch.smooth; s.data = [...s.data].sort((a, b) => a[0] - b[0]); });
    addTrend(ch, series, xs, ys, true);
    return { ...base, xAxis: { type: 'value', ...xTitle, axisLabel: { color: '#4B5066' }, splitLine: { show: false } }, yAxis: valAx(), series };
  }
  const horizontal = type === 'bar';
  const cat = { type: 'category', data: d.cats, axisLabel: { color: '#4B5066' }, boundaryGap: type !== 'area' && type !== 'line', ...xTitle };
  const stack = STACKABLE.includes(type) && ch.stacked && ch.stacked !== 'none' ? 'total' : undefined;
  let data = d.series.map((s) => s.data);
  if (ch.stacked === 'percent') { const tot = d.cats.map((_, i) => data.reduce((a, s) => a + Math.abs(s[i] || 0), 0)); data = data.map((s) => s.map((v, i) => (v == null || !tot[i] ? v : Math.round((v / tot[i]) * 10000) / 100))); }
  const sec = new Set(ch.secondary || []);
  const series = d.series.map((sr, k) => {
    const asLine = type === 'line' || type === 'area' || (type === 'combo' && (ch.lineSeries ? ch.lineSeries.includes(k) : k > 0));
    return { type: asLine ? 'line' : 'bar', name: sr.name, data: data[k], stack: asLine && type === 'combo' ? undefined : stack, smooth: !!ch.smooth, symbol: ch.markers === false ? 'none' : 'circle', showSymbol: ch.markers !== false, areaStyle: type === 'area' ? { opacity: stack ? 0.6 : 0.25 } : undefined, barMaxWidth: 48, yAxisIndex: sec.has(k) && !horizontal ? 1 : 0, label: lbl(asLine ? 'top' : stack ? 'inside' : (horizontal ? 'right' : 'top')) };
  });
  addTrend(ch, series, d.cats.map((_, i) => i + 1), d.series, false);
  const yAxis = sec.size && !horizontal ? [valAx(), valAx({ name: ch.y2Title || '', splitLine: { show: false }, min: undefined, max: undefined })] : valAx();
  return { ...base, xAxis: horizontal ? valAx() : cat, yAxis: horizontal ? { ...cat, boundaryGap: true } : yAxis, series };
}
function addTrend(ch, series, xs, ys, xy) {
  const tr = ch.trend; if (!tr || !tr.kind || tr.kind === 'none') return;
  const sr = ys[tr.series || 0]; if (!sr) return;
  const fit = trendFit(xs, sr.data, tr.kind, tr.period || 2); if (!fit) return;
  const label = [t('Tendencia'), tr.showEq ? fit.eq : '', tr.showR2 && fit.r2 != null ? `R² = ${fmt(fit.r2)}` : ''].filter(Boolean).join(' · ');
  const data = xy ? xs.map((x, i) => [x, fit.points[i]]).filter((p) => p[1] != null && typeof p[0] === 'number').sort((a, b) => a[0] - b[0]) : fit.points.map((v) => (v == null ? null : Math.round(v * 1e6) / 1e6));
  series.push({ type: 'line', name: label, data, smooth: tr.kind !== 'linear', symbol: 'none', lineStyle: { type: 'dashed', width: 2 }, z: 5, tooltip: { show: false } });
  ch.__fit = fit;
}

// ================================================================ cuadro de propiedades
export async function chartDialog(ch, title, { resolve, seriesNames = [] } = {}) {
  const k = (n, v) => (v === undefined ? ch[n] : v);
  const tabs = [
    { label: 'Tipo', body: h('div', { class: 'fx-grid' },
      field('Tipo de gráfico', sel('type', ch.type || 'column', CHART_TYPES)),
      field('Apilado', sel('stacked', ch.stacked || 'none', [['none', 'Normal'], ['stacked', 'Apilado'], ['percent', 'Apilado al 100 %']])),
      chk('markers', 'Mostrar marcadores en las líneas', ch.markers !== false), chk('smooth', 'Líneas suavizadas', !!ch.smooth),
      chk('scatterLines', 'Dispersión con líneas', !!ch.scatterLines), chk('fillRadar', 'Radar relleno', !!ch.fillRadar)) },
    { label: 'Datos', body: h('div', { class: 'fx-grid' },
      field('Intervalo de datos', h('input', { name: 'range', value: ch.range || '' }), 'wide'),
      field('Series de datos en', sel('byRows', ch.byRows ? '1' : '0', [['0', 'Columnas'], ['1', 'Filas']])), h('span'),
      field('Primera fila como etiqueta', sel('headers', ch.headers == null ? 'auto' : ch.headers ? '1' : '0', [['auto', 'Automático'], ['1', 'Sí'], ['0', 'No']])),
      field('Primera columna como etiqueta', sel('catCol', ch.catCol == null ? 'auto' : ch.catCol ? '1' : '0', [['auto', 'Automático'], ['1', 'Sí'], ['0', 'No']]))) },
    { label: 'Títulos', body: h('div', { class: 'fx-col' },
      field('Título', h('input', { name: 'title', value: ch.title || '' })), field('Subtítulo', h('input', { name: 'subtitle', value: ch.subtitle || '' })),
      field('Eje X', h('input', { name: 'xTitle', value: ch.xTitle || '' })), field('Eje Y', h('input', { name: 'yTitle', value: ch.yTitle || '' })), field('Eje Y secundario', h('input', { name: 'y2Title', value: ch.y2Title || '' }))) },
    { label: 'Leyenda y etiquetas', body: h('div', { class: 'fx-grid' },
      field('Leyenda', sel('legend', ch.legend || 'auto', [['auto', 'Automática'], ['bottom', 'Abajo'], ['top', 'Arriba'], ['right', 'Derecha'], ['left', 'Izquierda'], ['none', 'Sin leyenda']])),
      field('Etiquetas de datos', sel('labels', ch.labels || 'none', [['none', 'Ninguna'], ['value', 'Valor'], ['percent', 'Porcentaje (circular)'], ['category', 'Categoría'], ['both', 'Categoría y valor']]))) },
    { label: 'Ejes', body: h('div', { class: 'fx-grid' },
      field('Mínimo del eje Y', h('input', { name: 'yMin', type: 'number', value: ch.yMin ?? '', placeholder: t('Automático') })), field('Máximo del eje Y', h('input', { name: 'yMax', type: 'number', value: ch.yMax ?? '', placeholder: t('Automático') })),
      chk('logScale', 'Escala logarítmica', !!ch.logScale), chk('grid', 'Líneas de cuadrícula principales', ch.grid !== false), chk('minorGrid', 'Líneas de cuadrícula secundarias', !!ch.minorGrid)) },
    { label: 'Series', body: h('div', { class: 'fx-col' },
      h('p', { class: 'fx-note' }, t('Color de cada serie; en «Columnas y líneas», marca qué series van como línea y cuáles en el eje secundario.')),
      ...seriesNames.map((nm, i) => h('div', { class: 'ser-row' }, h('span', {}, nm), h('input', { type: 'color', name: `color${i}`, value: ch.colors?.[i] || PALETTE[i % PALETTE.length], 'aria-label': tf('Color de {0}', nm) }),
        h('label', { class: 'fx-check' }, h('input', { type: 'checkbox', name: `line${i}`, checked: (ch.lineSeries ? ch.lineSeries.includes(i) : i > 0) ? true : null }), ` ${t('Línea')}`),
        h('label', { class: 'fx-check' }, h('input', { type: 'checkbox', name: `sec${i}`, checked: ch.secondary?.includes(i) ? true : null }), ` ${t('Eje secundario')}`)))) },
    { label: 'Tendencia', body: h('div', { class: 'fx-grid' },
      field('Línea de tendencia', sel('trend', ch.trend?.kind || 'none', [['none', 'Ninguna'], ['linear', 'Lineal'], ['exp', 'Exponencial'], ['log', 'Logarítmica'], ['power', 'Potencial'], ['poly2', 'Polinómica (grado 2)'], ['movavg', 'Media móvil']])),
      field('Serie', sel('trendSeries', String(ch.trend?.series || 0), seriesNames.map((n, i) => [String(i), n]))),
      field('Períodos (media móvil)', numIn('period', ch.trend?.period || 2, { step: 1, min: 2, max: 100 })), h('span'),
      chk('showEq', 'Mostrar la ecuación', !!ch.trend?.showEq), chk('showR2', 'Mostrar el coeficiente de determinación (R²)', !!ch.trend?.showR2)) },
  ];
  const body = await tabbed(title, tabs, { width: '600px' });
  if (!body) return null;
  const v = (n) => { const el = body.querySelector(`[name="${n}"]`); return !el ? undefined : el.type === 'checkbox' ? el.checked : el.value; };
  const num = (n) => (v(n) === '' || v(n) == null ? undefined : Number(v(n)));
  const tri = (n) => (v(n) === 'auto' ? undefined : v(n) === '1');
  if (resolve && !resolve(v('range'))) { toast('Rango no válido', { error: true }); return null; }
  const out = {
    type: v('type'), stacked: v('stacked') === 'none' ? undefined : v('stacked'), markers: v('markers') ? undefined : false, smooth: v('smooth') || undefined, scatterLines: v('scatterLines') || undefined, fillRadar: v('fillRadar') || undefined,
    range: v('range'), byRows: v('byRows') === '1' || undefined, headers: tri('headers'), catCol: tri('catCol'),
    title: v('title'), subtitle: v('subtitle') || undefined, xTitle: v('xTitle') || undefined, yTitle: v('yTitle') || undefined, y2Title: v('y2Title') || undefined,
    legend: v('legend') === 'auto' ? undefined : v('legend'), labels: v('labels') === 'none' ? undefined : v('labels'),
    yMin: num('yMin'), yMax: num('yMax'), logScale: v('logScale') || undefined, grid: v('grid') ? undefined : false, minorGrid: v('minorGrid') || undefined,
    colors: seriesNames.map((_, i) => v(`color${i}`)).some((c, i) => c && c.toLowerCase() !== PALETTE[i % PALETTE.length].toLowerCase()) ? seriesNames.map((_, i) => v(`color${i}`)) : undefined,
    lineSeries: v('type') === 'combo' ? seriesNames.map((_, i) => i).filter((i) => v(`line${i}`)) : undefined,
    secondary: seriesNames.map((_, i) => i).filter((i) => v(`sec${i}`)),
    trend: v('trend') === 'none' ? undefined : { kind: v('trend'), series: Number(v('trendSeries')) || 0, period: num('period') || 2, showEq: v('showEq') || undefined, showR2: v('showR2') || undefined },
  };
  if (!out.secondary.length) out.secondary = undefined;
  return out;
}

// ================================================================ gráficos nativos en .xlsx (DrawingML)
const X = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const qs = (name) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);
const ref = (sheet, cells) => { if (!cells?.length) return ''; const a = cells[0], b = cells[cells.length - 1]; return `${qs(sheet)}!$${colName(a.c)}$${a.r + 1}${a.r === b.r && a.c === b.c ? '' : `:$${colName(b.c)}$${b.r + 1}`}`; };
const C_NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const richTitle = (text, size = 1400) => `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${size}" b="1"/></a:pPr><a:r><a:rPr lang="es-MX" sz="${size}" b="1"/><a:t>${X(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
const solid = (hex) => `<c:spPr><a:solidFill><a:srgbClr val="${hex.replace('#', '').toUpperCase()}"/></a:solidFill></c:spPr>`;
const lineSp = (hex, w = 28575) => `<c:spPr><a:ln w="${w}" cap="rnd"><a:solidFill><a:srgbClr val="${hex.replace('#', '').toUpperCase()}"/></a:solidFill><a:round/></a:ln></c:spPr>`;
const TREND_X = { linear: 'linear', exp: 'exp', log: 'log', power: 'power', poly2: 'poly', movavg: 'movingAvg' };
export function chartSpaceXml(ch, L, sheetName) {
  const type = ch.type || 'column'; const pal = PALETTE.map((c, i) => ch.colors?.[i] || c);
  const strCache = (vals) => `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${X(v)}</c:v></c:pt>`).join('')}</c:strCache>`;
  const numCache = (vals) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (v == null ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join('')}</c:numCache>`;
  const tx = (s) => (s.nameAt ? `<c:tx><c:strRef><c:f>${ref(sheetName, [s.nameAt])}</c:f>${strCache([s.name])}</c:strRef></c:tx>` : `<c:tx><c:v>${X(s.name)}</c:v></c:tx>`);
  const cat = L.catCells ? `<c:cat><c:strRef><c:f>${ref(sheetName, L.catCells)}</c:f>${strCache(L.cats)}</c:strRef></c:cat>` : '';
  const val = (s, tag = 'val') => `<c:${tag}><c:numRef><c:f>${ref(sheetName, s.cells)}</c:f>${numCache(s.data)}</c:numRef></c:${tag}>`;
  const dLbls = ch.labels && ch.labels !== 'none' ? `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="${ch.labels === 'value' || ch.labels === 'both' ? 1 : 0}"/><c:showCatName val="${ch.labels === 'category' || ch.labels === 'both' ? 1 : 0}"/><c:showSerName val="0"/><c:showPercent val="${ch.labels === 'percent' ? 1 : 0}"/><c:showBubbleSize val="0"/></c:dLbls>` : '';
  const trend = (k) => (ch.trend && ch.trend.kind && (ch.trend.series || 0) === k ? `<c:trendline><c:spPr><a:ln w="19050"><a:prstDash val="dash"/></a:ln></c:spPr><c:trendlineType val="${TREND_X[ch.trend.kind] || 'linear'}"/>${ch.trend.kind === 'poly2' ? '<c:order val="2"/>' : ''}${ch.trend.kind === 'movavg' ? `<c:period val="${ch.trend.period || 2}"/>` : ''}<c:dispRSqr val="${ch.trend.showR2 ? 1 : 0}"/><c:dispEq val="${ch.trend.showEq ? 1 : 0}"/></c:trendline>` : '');
  const grouping = ch.stacked === 'percent' ? 'percentStacked' : ch.stacked === 'stacked' ? 'stacked' : null;
  const axIds = (a, b) => `<c:axId val="${a}"/><c:axId val="${b}"/>`;
  const sec = new Set(ch.secondary || []);
  let plot = ''; let axes = '';
  const catAx = (id, cross, pos = 'b', del = 0, title = ch.xTitle) => `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${del}"/><c:axPos val="${pos}"/>${title ? richTitle(title, 1000) : ''}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  const valAx = (id, cross, pos = 'l', { title = ch.yTitle, grid = ch.grid !== false, crosses = 'autoZero', del = 0, min = ch.yMin, max = ch.yMax, log = ch.logScale } = {}) => `<c:valAx><c:axId val="${id}"/><c:scaling>${log ? '<c:logBase val="10"/>' : ''}<c:orientation val="minMax"/>${max != null ? `<c:max val="${max}"/>` : ''}${min != null ? `<c:min val="${min}"/>` : ''}</c:scaling><c:delete val="${del}"/><c:axPos val="${pos}"/>${grid ? '<c:majorGridlines/>' : ''}${ch.minorGrid ? '<c:minorGridlines/>' : ''}${title ? richTitle(title, 1000) : ''}<c:numFmt formatCode="${ch.stacked === 'percent' ? '0%' : 'General'}" sourceLinked="${ch.stacked === 'percent' ? 0 : 1}"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="${crosses}"/><c:crossBetween val="between"/></c:valAx>`;
  const sers = L.series;
  const serXml = (s, k, kind) => {
    const color = pal[k % pal.length];
    if (kind === 'bar') return `<c:ser><c:idx val="${k}"/><c:order val="${k}"/>${tx(s)}${solid(color)}<c:invertIfNegative val="0"/>${dLbls}${trend(k)}${cat}${val(s)}</c:ser>`;
    if (kind === 'line') return `<c:ser><c:idx val="${k}"/><c:order val="${k}"/>${tx(s)}${lineSp(color)}<c:marker><c:symbol val="${ch.markers === false ? 'none' : 'circle'}"/><c:size val="5"/>${solid(color)}</c:marker>${dLbls}${trend(k)}${cat}${val(s)}<c:smooth val="${ch.smooth ? 1 : 0}"/></c:ser>`;
    if (kind === 'area') return `<c:ser><c:idx val="${k}"/><c:order val="${k}"/>${tx(s)}${solid(color)}${dLbls}${trend(k)}${cat}${val(s)}</c:ser>`;
    if (kind === 'radar') return `<c:ser><c:idx val="${k}"/><c:order val="${k}"/>${tx(s)}${lineSp(color)}<c:marker><c:symbol val="circle"/></c:marker>${dLbls}${cat}${val(s)}</c:ser>`;
    return '';
  };
  if (type === 'pie' || type === 'doughnut') {
    const s = sers[0]; if (!s) return null;
    const pts = L.cats.map((_, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/>${solid(pal[i % pal.length])}</c:dPt>`).join('');
    plot = `<c:${type === 'pie' ? 'pieChart' : 'doughnutChart'}><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>${tx(s)}${pts}${dLbls || (ch.labels === 'none' ? '' : '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/></c:dLbls>')}${cat}${val(s)}</c:ser><c:firstSliceAng val="0"/>${type === 'doughnut' ? '<c:holeSize val="55"/>' : ''}</c:${type === 'pie' ? 'pieChart' : 'doughnutChart'}>`;
  } else if (type === 'scatter' || type === 'bubble') {
    const xs = sers[0]; const ys = type === 'bubble' ? sers.slice(1, 2) : sers.slice(1);
    plot = type === 'bubble'
      ? `<c:bubbleChart><c:varyColors val="0"/>${ys.map((s, k) => `<c:ser><c:idx val="${k}"/><c:order val="${k}"/>${tx(s)}${solid(pal[k % pal.length])}<c:invertIfNegative val="0"/>${dLbls}${val(xs, 'xVal')}${val(s, 'yVal')}${sers[2] ? val(sers[2], 'bubbleSize') : ''}<c:bubble3D val="0"/></c:ser>`).join('')}<c:bubbleScale val="100"/><c:showNegBubbles val="0"/>${axIds(500, 501)}</c:bubbleChart>`
      : `<c:scatterChart><c:scatterStyle val="${ch.scatterLines ? (ch.smooth ? 'smoothMarker' : 'lineMarker') : 'lineMarker'}"/><c:varyColors val="0"/>${ys.map((s, k) => `<c:ser><c:idx val="${k}"/><c:order val="${k}"/>${tx(s)}${ch.scatterLines ? lineSp(pal[k % pal.length]) : '<c:spPr><a:ln w="25400"><a:noFill/></a:ln></c:spPr>'}<c:marker><c:symbol val="circle"/><c:size val="6"/>${solid(pal[k % pal.length])}</c:marker>${dLbls}${trend(k)}${val(xs, 'xVal')}${val(s, 'yVal')}<c:smooth val="${ch.smooth ? 1 : 0}"/></c:ser>`).join('')}${axIds(500, 501)}</c:scatterChart>`;
    axes = valAx(500, 501, 'b', { title: ch.xTitle, grid: false, min: undefined, max: undefined, log: false }) + valAx(501, 500, 'l');
  } else if (type === 'radar') {
    plot = `<c:radarChart><c:radarStyle val="${ch.fillRadar ? 'filled' : 'marker'}"/><c:varyColors val="0"/>${sers.map((s, k) => serXml(s, k, 'radar')).join('')}${axIds(500, 501)}</c:radarChart>`;
    axes = catAx(500, 501) + valAx(501, 500);
  } else {
    const isBar = type === 'column' || type === 'bar' || type === 'combo';
    const lineIdx = type === 'combo' ? (ch.lineSeries || sers.map((_, i) => i).filter((i) => i > 0)) : [];
    const groups = { primary: { bar: [], line: [] }, secondary: { bar: [], line: [] } };
    sers.forEach((s, k) => { const kind = type === 'line' ? 'line' : type === 'area' ? 'area' : (lineIdx.includes(k) ? 'line' : 'bar'); const g = sec.has(k) && type !== 'bar' ? groups.secondary : groups.primary; (g[kind] ||= []).push([s, k]); });
    const chartOf = (kind, list, ids) => {
      if (!list?.length) return '';
      const body = list.map(([s, k]) => serXml(s, k, kind)).join('');
      if (kind === 'bar') return `<c:barChart><c:barDir val="${type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${grouping || 'clustered'}"/><c:varyColors val="0"/>${body}<c:gapWidth val="80"/>${grouping ? '<c:overlap val="100"/>' : ''}${axIds(...ids)}</c:barChart>`;
      if (kind === 'line') return `<c:lineChart><c:grouping val="${type === 'combo' ? 'standard' : grouping || 'standard'}"/><c:varyColors val="0"/>${body}<c:marker val="1"/>${axIds(...ids)}</c:lineChart>`;
      return `<c:areaChart><c:grouping val="${grouping || 'standard'}"/><c:varyColors val="0"/>${body}${axIds(...ids)}</c:areaChart>`;
    };
    plot = ['bar', 'line', 'area'].map((kd) => chartOf(kd, groups.primary[kd], [500, 501])).join('') + ['bar', 'line', 'area'].map((kd) => chartOf(kd, groups.secondary[kd], [502, 503])).join('');
    axes = catAx(500, 501, type === 'bar' ? 'l' : 'b') + valAx(501, 500, type === 'bar' ? 'b' : 'l');
    if (sec.size && type !== 'bar') axes += catAx(502, 503, 'b', 1, '') + valAx(503, 502, 'r', { title: ch.y2Title, grid: false, crosses: 'max', min: undefined, max: undefined, log: false });
    void isBar;
  }
  const legend = (ch.legend || (sers.length > 1 || type === 'pie' || type === 'doughnut' ? 'bottom' : 'none'));
  const lpos = { bottom: 'b', top: 't', right: 'r', left: 'l' }[legend];
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace ${C_NS}><c:roundedCorners val="0"/><c:chart>${ch.title ? richTitle(ch.title) : ''}<c:autoTitleDeleted val="${ch.title ? 0 : 1}"/><c:plotArea><c:layout/>${plot}${axes}</c:plotArea>${lpos ? `<c:legend><c:legendPos val="${lpos}"/><c:overlay val="0"/></c:legend>` : ''}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:chartSpace>`;
}
// celda (fila/col y desplazamiento en EMU) para una posición en píxeles
function cellAt(sheet, px, py) {
  const EMU = 9525; let x = 0; let c = 0; while (c < 16000 && x + sheet.cw(c) <= px) { x += sheet.cw(c); c++; }
  let y = 0; let r = 0; while (r < 1048000 && y + sheet.rh(r) <= py) { y += sheet.rh(r); r++; }
  return { c, r, dx: Math.round((px - x) * EMU), dy: Math.round((py - y) * EMU) };
}
const anchorXml = (sheet, ch, rid, n) => {
  const a = cellAt(sheet, ch.x, ch.y); const b = cellAt(sheet, ch.x + ch.w, ch.y + ch.h);
  return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${a.c}</xdr:col><xdr:colOff>${a.dx}</xdr:colOff><xdr:row>${a.r}</xdr:row><xdr:rowOff>${a.dy}</xdr:rowOff></xdr:from><xdr:to><xdr:col>${b.c}</xdr:col><xdr:colOff>${b.dx}</xdr:colOff><xdr:row>${b.r}</xdr:row><xdr:rowOff>${b.dy}</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${n + 1}" name="${X(ch.title || tf('Gráfico {0}', n))}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
};
// agrega los gráficos al .xlsx escrito por ExcelJS
export async function addXlsxCharts(bytes, wb, layoutOf) {
  const sheetsWithCharts = wb.sheets.map((s, i) => [s, i]).filter(([s]) => s.charts?.length);
  if (!sheetsWithCharts.length) return bytes;
  const zip = await JSZip.loadAsync(bytes);
  const wbx = await zip.file('xl/workbook.xml').async('string'); const wbr = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const rels = Object.fromEntries([...wbr.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [m[1], m[2]]));
  const sheetFile = Object.fromEntries([...wbx.matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)].map((m) => [m[1].replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>'), `xl/${rels[m[2]].replace(/^\/?xl\//, '')}`]));
  let ct = await zip.file('[Content_Types].xml').async('string');
  let chartN = Object.keys(zip.files).filter((f) => /^xl\/charts\/chart\d+\.xml$/.test(f)).length; let drawN = Object.keys(zip.files).filter((f) => /^xl\/drawings\/drawing\d+\.xml$/.test(f)).length;
  for (const [s] of sheetsWithCharts) {
    const sf = sheetFile[s.name]; if (!sf || !zip.file(sf)) continue;
    let sx = await zip.file(sf).async('string'); if (/<drawing\b/.test(sx)) continue;
    drawN += 1; const dPath = `xl/drawings/drawing${drawN}.xml`; const anchors = []; const dRels = [];
    s.charts.forEach((ch, k) => {
      const L = layoutOf(ch, s); if (!L || !L.series.length) return;
      const xml = chartSpaceXml(ch, L, L.sheetName || s.name); if (!xml) return;
      chartN += 1; zip.file(`xl/charts/chart${chartN}.xml`, xml);
      ct = ct.replace('</Types>', `<Override PartName="/xl/charts/chart${chartN}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`);
      dRels.push(`<Relationship Id="rIdC${k + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartN}.xml"/>`);
      anchors.push(anchorXml(s, ch, `rIdC${k + 1}`, k + 1));
    });
    if (!anchors.length) continue;
    zip.file(dPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchors.join('')}</xdr:wsDr>`);
    zip.file(`xl/drawings/_rels/drawing${drawN}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${dRels.join('')}</Relationships>`);
    ct = ct.replace('</Types>', `<Override PartName="/${dPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`);
    const relPath = sf.replace(/([^/]+)$/, '_rels/$1.rels');
    let sr = zip.file(relPath) ? await zip.file(relPath).async('string') : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    sr = sr.replace('</Relationships>', `<Relationship Id="rIdDrw1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawN}.xml"/></Relationships>`);
    zip.file(relPath, sr);
    if (!/xmlns:r=/.test(sx.slice(0, 600))) sx = sx.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
    const before = ['<legacyDrawing', '<legacyDrawingHF', '<picture', '<oleObjects', '<controls', '<webPublishItems', '<tableParts', '<extLst', '</worksheet>'].map((tag) => sx.indexOf(tag)).filter((i) => i >= 0);
    const at = Math.min(...before);
    sx = `${sx.slice(0, at)}<drawing r:id="rIdDrw1"/>${sx.slice(at)}`;
    zip.file(sf, sx);
  }
  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

// gráficos de un .xlsx de Excel o LibreOffice → modelo de Calcula
const TYPE_FROM_X = (plot, el) => {
  const ln = el.localName;
  if (ln === 'barChart') { const dir = el.getElementsByTagNameNS(CNS, 'barDir')[0]?.getAttribute('val'); return dir === 'bar' ? 'bar' : 'column'; }
  return { lineChart: 'line', line3DChart: 'line', areaChart: 'area', pieChart: 'pie', pie3DChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter', bubbleChart: 'bubble', radarChart: 'radar', bar3DChart: 'column' }[ln] || null;
};
const CNS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const parseRef = (f) => { const m = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(String(f || '').trim()); if (!m) return null; const sheet = (m[1] || m[2]).replace(/''/g, "'"); const r1 = Number(m[4]) - 1, c1 = colIndex(m[3]); return { sheet, r1, c1, r2: m[6] ? Number(m[6]) - 1 : r1, c2: m[5] ? colIndex(m[5]) : c1 }; };
function chartFromXml(xml) {
  const d = new DOMParser().parseFromString(xml, 'application/xml');
  const plotArea = d.getElementsByTagNameNS(CNS, 'plotArea')[0]; if (!plotArea) return null;
  const plots = [...plotArea.children].filter((e) => /Chart$/.test(e.localName));
  if (!plots.length) return null;
  let type = TYPE_FROM_X(plotArea, plots[0]); if (!type) return null;
  const kinds = plots.map((p) => TYPE_FROM_X(plotArea, p));
  const ch = { type };
  if (kinds.includes('column') && kinds.includes('line')) ch.type = type = 'combo';
  const grouping = plots[0].getElementsByTagNameNS(CNS, 'grouping')[0]?.getAttribute('val');
  if (grouping === 'stacked') ch.stacked = 'stacked'; if (grouping === 'percentStacked') ch.stacked = 'percent';
  const titleEl = [...(d.getElementsByTagNameNS(CNS, 'chart')[0]?.children || [])].find((e) => e.localName === 'title');
  if (titleEl) ch.title = [...titleEl.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 't')].map((e) => e.textContent).join('');
  const lp = d.getElementsByTagNameNS(CNS, 'legendPos')[0]?.getAttribute('val'); ch.legend = lp ? { b: 'bottom', t: 'top', r: 'right', l: 'left', tr: 'right' }[lp] : 'none';
  const refs = []; let serIdx = 0; const lineSeries = [];
  for (const p of plots) for (const ser of p.getElementsByTagNameNS(CNS, 'ser')) {
    const f = (tag) => { const e = [...ser.children].find((x) => x.localName === tag); return e?.getElementsByTagNameNS(CNS, 'f')[0]?.textContent; };
    for (const tag of ['tx', 'cat', 'val', 'xVal', 'yVal', 'bubbleSize']) { const r = parseRef(f(tag)); if (r) refs.push(r); }
    if (p.localName === 'lineChart' && type === 'combo') lineSeries.push(serIdx);
    const tl = ser.getElementsByTagNameNS(CNS, 'trendline')[0];
    if (tl && !ch.trend) { const tt = tl.getElementsByTagNameNS(CNS, 'trendlineType')[0]?.getAttribute('val'); ch.trend = { kind: { linear: 'linear', exp: 'exp', log: 'log', power: 'power', poly: 'poly2', movingAvg: 'movavg' }[tt] || 'linear', series: serIdx, showEq: tl.getElementsByTagNameNS(CNS, 'dispEq')[0]?.getAttribute('val') === '1' || undefined, showR2: tl.getElementsByTagNameNS(CNS, 'dispRSqr')[0]?.getAttribute('val') === '1' || undefined }; }
    serIdx++;
  }
  if (type === 'combo') ch.lineSeries = lineSeries;
  if (!refs.length) return null;
  const sheet = refs[0].sheet; const same = refs.filter((r) => r.sheet === sheet);
  const g = { r1: Math.min(...same.map((r) => r.r1)), c1: Math.min(...same.map((r) => r.c1)), r2: Math.max(...same.map((r) => r.r2)), c2: Math.max(...same.map((r) => r.c2)) };
  // series en filas si los valores van a lo ancho
  const v0 = same.find((r) => r.r1 !== r.r2 || r.c1 !== r.c2); if (v0 && v0.r1 === v0.r2 && v0.c2 > v0.c1) ch.byRows = true;
  ch.range = `${colName(g.c1)}${g.r1 + 1}:${colName(g.c2)}${g.r2 + 1}`; ch.sourceSheet = sheet;
  if (type === 'scatter' || type === 'bubble') { ch.headers = refs.some((r) => r.r1 === g.r1 && r.r2 === g.r1 && r.c1 === r.c2) ? true : undefined; }
  if (ch.legend === 'none' && serIdx > 1) ch.legend = 'none';
  return ch;
}
export async function readXlsxCharts(bytes, wb) {
  try {
    const zip = await JSZip.loadAsync(bytes);
    const wbx = await zip.file('xl/workbook.xml')?.async('string'); const wbr = await zip.file('xl/_rels/workbook.xml.rels')?.async('string'); if (!wbx || !wbr) return 0;
    const rels = Object.fromEntries([...wbr.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [m[1], m[2]]));
    const resolveT = (base, target) => { if (target.startsWith('/')) return target.slice(1); const parts = base.split('/'); parts.pop(); for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.') parts.push(seg); } return parts.join('/'); };
    let n = 0;
    for (const m of wbx.matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)) {
      const name = m[1].replace(/&amp;/g, '&'); const s = wb.sheets.find((q) => q.name === name); if (!s || s.charts.length) continue;
      const sf = resolveT('xl/workbook.xml', rels[m[2]]); const relF = sf.replace(/([^/]+)$/, '_rels/$1.rels');
      const sx = await zip.file(sf)?.async('string'); const sr = await zip.file(relF)?.async('string'); if (!sx || !sr) continue;
      const did = /<drawing\b[^>]*r:id="([^"]+)"/.exec(sx)?.[1]; if (!did) continue;
      const dTarget = new RegExp(`Id="${did}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${did}"`).exec(sr); if (!dTarget) continue;
      const dPath = resolveT(sf, dTarget[1] || dTarget[2]); const dx = await zip.file(dPath)?.async('string'); if (!dx) continue;
      const dRelsTxt = await zip.file(dPath.replace(/([^/]+)$/, '_rels/$1.rels'))?.async('string') || '';
      const dRels = Object.fromEntries([...dRelsTxt.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((q) => [q[1], q[2]]));
      const doc = new DOMParser().parseFromString(dx, 'application/xml');
      const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
      for (const an of [...doc.getElementsByTagNameNS(XDR, 'twoCellAnchor'), ...doc.getElementsByTagNameNS(XDR, 'oneCellAnchor')]) {
        const cref = an.getElementsByTagNameNS(CNS, 'chart')[0]; if (!cref) continue;
        const rid = cref.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || cref.getAttribute('r:id');
        const cPath = resolveT(dPath, dRels[rid] || ''); const cx = await zip.file(cPath)?.async('string'); if (!cx) continue;
        const ch = chartFromXml(cx); if (!ch) continue;
        const num = (el, tag) => Number(el?.getElementsByTagNameNS(XDR, tag)[0]?.textContent || 0);
        const from = an.getElementsByTagNameNS(XDR, 'from')[0]; const to = an.getElementsByTagNameNS(XDR, 'to')[0]; const ext = an.getElementsByTagNameNS(XDR, 'ext')[0];
        const px = (col, off) => { let x = 0; for (let c = 0; c < col; c++) x += s.cw(c); return x + off / 9525; };
        const py = (row, off) => { let y = 0; for (let r = 0; r < row; r++) y += s.rh(r); return y + off / 9525; };
        const x = px(num(from, 'col'), num(from, 'colOff')); const y = py(num(from, 'row'), num(from, 'rowOff'));
        const w = to ? px(num(to, 'col'), num(to, 'colOff')) - x : Number(ext?.getAttribute('cx') || 4572000) / 9525; const hh = to ? py(num(to, 'row'), num(to, 'rowOff')) - y : Number(ext?.getAttribute('cy') || 2743200) / 9525;
        if (ch.sourceSheet && ch.sourceSheet !== s.name) ch.range = `${qs(ch.sourceSheet)}!${ch.range}`;
        delete ch.sourceSheet;
        s.charts.push({ id: Math.random().toString(36).slice(2, 10), ...ch, x: Math.round(x), y: Math.round(y), w: Math.max(220, Math.round(w)), h: Math.max(160, Math.round(hh)) });
        n++;
      }
    }
    return n;
  } catch (e) { console.warn('gráficos xlsx', e); return 0; }
}

// ================================================================ gráficos nativos en .ods (objetos de ODF)
const ODF_CLASS = { column: 'chart:bar', bar: 'chart:bar', combo: 'chart:bar', line: 'chart:line', area: 'chart:area', pie: 'chart:circle', doughnut: 'chart:ring', scatter: 'chart:scatter', bubble: 'chart:bubble', radar: 'chart:radar' };
const CLASS_IN = { 'chart:bar': 'column', 'chart:line': 'line', 'chart:area': 'area', 'chart:circle': 'pie', 'chart:ring': 'doughnut', 'chart:scatter': 'scatter', 'chart:bubble': 'bubble', 'chart:radar': 'radar', 'chart:filled-radar': 'radar' };
const odsRef = (sheet, cells) => { if (!cells?.length) return ''; const a = cells[0], b = cells[cells.length - 1]; const q = `$${/^[A-Za-z0-9_]+$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`}`; return `${q}.$${colName(a.c)}$${a.r + 1}${a.r === b.r && a.c === b.c ? '' : `:${q}.$${colName(b.c)}$${b.r + 1}`}`; };
export function odsChartContent(ch, L, sheetName, wpx, hpx) {
  const type = ch.type || 'column'; const pal = PALETTE.map((c, i) => ch.colors?.[i] || c);
  const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"';
  const cm = (px) => `${(px * 2.54 / 96).toFixed(3)}cm`;
  const styles = []; const st = (name, inner) => { styles.push(`<style:style style:name="${name}" style:family="chart">${inner}</style:style>`); return name; };
  const plotProps = [type === 'bar' ? 'chart:vertical="true"' : '', ch.stacked === 'stacked' ? 'chart:stacked="true"' : '', ch.stacked === 'percent' ? 'chart:percentage="true"' : '', ch.smooth && (type === 'line' || type === 'scatter') ? 'chart:interpolation="cubic-spline"' : '', ch.markers === false && type === 'line' ? 'chart:symbol-type="none"' : (type === 'line' || type === 'scatter' || type === 'radar' ? 'chart:symbol-type="automatic"' : ''), ch.labels === 'value' || ch.labels === 'both' ? 'chart:data-label-number="value"' : ch.labels === 'percent' ? 'chart:data-label-number="percentage"' : '', ch.labels === 'category' || ch.labels === 'both' ? 'chart:data-label-text="true"' : '', type === 'scatter' && !ch.scatterLines ? 'chart:lines="false"' : '', type === 'radar' && ch.fillRadar ? '' : ''].filter(Boolean).join(' ');
  const pa = st('Plot', `<style:chart-properties ${plotProps} chart:include-hidden-cells="false"/>`);
  const trendSt = ch.trend?.kind ? st('Trend', `<style:chart-properties chart:regression-type="${{ linear: 'linear', exp: 'exponential', log: 'logarithmic', power: 'power', poly2: 'polynomial', movavg: 'moving-average' }[ch.trend.kind] || 'linear'}"${ch.trend.kind === 'poly2' ? ' loext:regression-max-degree="2"' : ''}${ch.trend.kind === 'movavg' ? ` loext:regression-period="${ch.trend.period || 2}"` : ''}/>`) : null;
  const eqSt = ch.trend?.showEq || ch.trend?.showR2 ? st('Eq', `<style:chart-properties chart:display-equation="${ch.trend.showEq ? 'true' : 'false'}" chart:display-r-square="${ch.trend.showR2 ? 'true' : 'false'}"/>`) : null;
  const sec = new Set(ch.secondary || []);
  const lineIdx = type === 'combo' ? (ch.lineSeries || L.series.map((_, i) => i).filter((i) => i > 0)) : [];
  const series = (type === 'scatter' || type === 'bubble' ? (type === 'bubble' ? L.series.slice(1, 2) : L.series.slice(1)) : (type === 'pie' || type === 'doughnut' ? L.series.slice(0, 1) : L.series)).map((s, k0) => {
    const k = type === 'scatter' || type === 'bubble' ? k0 + 1 : k0;
    const color = pal[k0 % pal.length];
    const sst = st(`S${k}`, `<style:graphic-properties svg:stroke-color="${color}" draw:fill-color="${color}" svg:stroke-width="0.07cm"/>`);
    const cls = type === 'combo' ? (lineIdx.includes(k) ? 'chart:line' : 'chart:bar') : '';
    const dom = type === 'scatter' ? `<chart:domain table:cell-range-address="${odsRef(sheetName, L.series[0].cells)}"/>` : type === 'bubble' ? `<chart:domain table:cell-range-address="${odsRef(sheetName, L.series[0].cells)}"/>` : '';
    const reg = trendSt && (ch.trend.series || 0) === k ? `<chart:regression-curve chart:style-name="${trendSt}">${eqSt ? `<chart:equation chart:style-name="${eqSt}" chart:display-equation="${ch.trend.showEq ? 'true' : 'false'}" chart:display-r-square="${ch.trend.showR2 ? 'true' : 'false'}"/>` : ''}</chart:regression-curve>` : '';
    const pts = type === 'pie' || type === 'doughnut' ? L.cats.map((_, i) => `<chart:data-point chart:style-name="${st(`P${i}`, `<style:graphic-properties draw:fill-color="${pal[i % pal.length]}"/>`)}"/>`).join('') : `<chart:data-point chart:repeated="${Math.max(1, L.cats.length)}"/>`;
    const values = type === 'bubble' && L.series[2] ? odsRef(sheetName, L.series[2].cells) : odsRef(sheetName, s.cells);
    if (type === 'bubble') return `<chart:series chart:style-name="${sst}" chart:values-cell-range-address="${values}"${s.nameAt ? ` chart:label-cell-address="${odsRef(sheetName, [s.nameAt])}"` : ''} chart:class="chart:bubble"><chart:domain table:cell-range-address="${odsRef(sheetName, s.cells)}"/>${dom}${pts}</chart:series>`;
    return `<chart:series chart:style-name="${sst}" chart:values-cell-range-address="${values}"${s.nameAt ? ` chart:label-cell-address="${odsRef(sheetName, [s.nameAt])}"` : ''}${cls ? ` chart:class="${cls}"` : ''}${sec.has(k) ? ' chart:attached-axis="secondary-y"' : ''}>${dom}${reg}${pts}</chart:series>`;
  }).join('');
  const axSt = st('Ax', `<style:chart-properties chart:display-label="true"${ch.logScale ? ' chart:logarithmic="true"' : ''}${ch.yMin != null ? ` chart:minimum="${ch.yMin}"` : ''}${ch.yMax != null ? ` chart:maximum="${ch.yMax}"` : ''}/>`);
  const tt = (tag, text) => (text ? `<chart:${tag}><text:p>${X(text)}</text:p></chart:${tag}>` : '');
  const axes = type === 'pie' || type === 'doughnut' ? '' : `<chart:axis chart:dimension="x" chart:name="primary-x">${tt('title', ch.xTitle)}${L.catCells && type !== 'scatter' && type !== 'bubble' ? `<chart:categories table:cell-range-address="${odsRef(sheetName, L.catCells)}"/>` : ''}</chart:axis><chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="${axSt}">${tt('title', ch.yTitle)}${ch.grid !== false ? '<chart:grid chart:class="major"/>' : ''}${ch.minorGrid ? '<chart:grid chart:class="minor"/>' : ''}</chart:axis>${sec.size ? `<chart:axis chart:dimension="y" chart:name="secondary-y">${tt('title', ch.y2Title)}</chart:axis>` : ''}`;
  const legend = ch.legend || (L.series.length > 1 || type === 'pie' || type === 'doughnut' ? 'bottom' : 'none');
  const cls = type === 'radar' && ch.fillRadar ? 'chart:filled-radar' : ODF_CLASS[type];
  const allCells = [...(L.catCells || []), ...L.series.flatMap((s) => [...(s.nameAt ? [s.nameAt] : []), ...s.cells])];
  const g = { r1: Math.min(...allCells.map((c) => c.r)), c1: Math.min(...allCells.map((c) => c.c)), r2: Math.max(...allCells.map((c) => c.r)), c2: Math.max(...allCells.map((c) => c.c)) };
  const whole = odsRef(sheetName, [{ r: g.r1, c: g.c1 }, { r: g.r2, c: g.c2 }]);
  const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3"><office:automatic-styles>${styles.join('')}</office:automatic-styles><office:body><office:chart><chart:chart svg:width="${cm(wpx)}" svg:height="${cm(hpx)}" chart:class="${cls}">${tt('title', ch.title)}${tt('subtitle', ch.subtitle)}${legend !== 'none' ? `<chart:legend chart:legend-position="${{ bottom: 'bottom', top: 'top', right: 'end', left: 'start' }[legend]}"/>` : ''}<chart:plot-area chart:style-name="${pa}" table:cell-range-address="${whole}" chart:data-source-has-labels="${L.headerRow && L.catCol ? 'both' : L.headerRow ? 'row' : L.catCol ? 'column' : 'none'}"${L.byRows ? ' chart:series-source="rows"' : ''}>${axes}${series}</chart:plot-area></chart:chart></office:chart></office:body></office:document-content>`;
  return content;
}
// agrega los gráficos a un .ods ya escrito (marcos en table:shapes + subdocumentos «Object N»)
export async function addOdsCharts(bytes, wb, layoutOf) {
  if (!wb.sheets.some((s) => s.charts?.length)) return bytes;
  const zip = await JSZip.loadAsync(bytes);
  let content = await zip.file('content.xml').async('string'); let man = await zip.file('META-INF/manifest.xml').async('string');
  let n = 0;
  for (const s of wb.sheets) {
    if (!s.charts?.length) continue;
    const frames = [];
    for (const ch of s.charts) {
      const L = layoutOf(ch, s); if (!L || !L.series.length) continue;
      n += 1; const dir = `Object ${n}`;
      zip.file(`${dir}/content.xml`, odsChartContent(ch, L, L.sheetName || s.name, ch.w, ch.h));
      zip.file(`${dir}/styles.xml`, '<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"/>');
      man = man.replace('</manifest:manifest>', `<manifest:file-entry manifest:full-path="${dir}/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.chart"/><manifest:file-entry manifest:full-path="${dir}/content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="${dir}/styles.xml" manifest:media-type="text/xml"/></manifest:manifest>`);
      const cm = (px) => `${(px * 2.54 / 96).toFixed(3)}cm`;
      frames.push(`<draw:frame draw:z-index="${n}" draw:name="${X(ch.title || tf('Gráfico {0}', n))}" svg:width="${cm(ch.w)}" svg:height="${cm(ch.h)}" svg:x="${cm(ch.x)}" svg:y="${cm(ch.y)}"><draw:object xlink:href="./${dir}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`);
    }
    if (!frames.length) continue;
    const esc = s.name.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    const re = new RegExp(`(<table:table table:name="${esc.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*>)`);
    content = content.replace(re, `$1<table:shapes>${frames.join('')}</table:shapes>`);
  }
  if (!/xmlns:draw=/.test(content.slice(0, 2000))) content = content.replace('<office:document-content ', '<office:document-content xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" ');
  if (!/xmlns:xlink=/.test(content.slice(0, 2000))) content = content.replace('<office:document-content ', '<office:document-content xmlns:xlink="http://www.w3.org/1999/xlink" ');
  if (!/xmlns:svg=/.test(content.slice(0, 2000))) content = content.replace('<office:document-content ', '<office:document-content xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ');
  zip.file('content.xml', content); zip.file('META-INF/manifest.xml', man);
  // el tipo MIME primero y sin comprimir
  const out = new JSZip(); const mt = await zip.file('mimetype').async('string'); out.file('mimetype', mt, { compression: 'STORE' });
  for (const [name, f] of Object.entries(zip.files)) if (name !== 'mimetype' && !f.dir) out.file(name, await f.async('uint8array'));
  return out.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
// gráficos de un .ods (LibreOffice) → modelo de Calcula
const parseOdsRange = (s) => { const parts = String(s || '').trim().split(/\s+/)[0].split(':'); const one = (p) => { const m = /^\$?'?((?:[^'.]|'')+)'?\.\$?([A-Z]+)\$?(\d+)$/.exec(p); return m ? { sheet: m[1].replace(/''/g, "'"), r: Number(m[3]) - 1, c: colIndex(m[2]) } : null; }; const a = one(parts[0]); if (!a) return null; const b = parts[1] ? (one(parts[1]) || one(`${a.sheet}.${parts[1].replace(/^\$?[^.]*\./, '')}`)) : a; return b ? { sheet: a.sheet, r1: a.r, c1: a.c, r2: b.r, c2: b.c } : null; };
export async function readOdsCharts(bytes, wb) {
  try {
    const zip = await JSZip.loadAsync(bytes);
    const content = await zip.file('content.xml')?.async('string'); if (!content) return 0;
    const doc = new DOMParser().parseFromString(content, 'application/xml');
    const TNS = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0', DNS = 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0', SVG = 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0', XL = 'http://www.w3.org/1999/xlink', CH = 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0', TX = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0';
    const len = (v) => { const m = /^([\d.]+)(cm|mm|in|pt)?$/.exec(v || ''); if (!m) return 0; return Number(m[1]) * ({ cm: 96 / 2.54, mm: 96 / 25.4, in: 96, pt: 96 / 72 }[m[2] || 'cm']); };
    let n = 0;
    for (const tbl of doc.getElementsByTagNameNS(TNS, 'table')) {
      const s = wb.sheets.find((q) => q.name === tbl.getAttributeNS(TNS, 'name')); if (!s || s.charts.length) continue;
      for (const obj of tbl.getElementsByTagNameNS(DNS, 'object')) {
        const fr = obj.parentElement; const href = (obj.getAttributeNS(XL, 'href') || '').replace(/^\.\//, '');
        const cx = await zip.file(`${href}/content.xml`)?.async('string'); if (!cx) continue;
        const cd = new DOMParser().parseFromString(cx, 'application/xml');
        const chart = cd.getElementsByTagNameNS(CH, 'chart')[0]; if (!chart) continue;
        let type = CLASS_IN[chart.getAttributeNS(CH, 'class')]; if (!type) continue;
        const ch = { type };
        const autos = new Map(); for (const st of cd.getElementsByTagNameNS('urn:oasis:names:tc:opendocument:xmlns:style:1.0', 'style')) autos.set(st.getAttributeNS('urn:oasis:names:tc:opendocument:xmlns:style:1.0', 'name'), st.getElementsByTagNameNS('urn:oasis:names:tc:opendocument:xmlns:style:1.0', 'chart-properties')[0]);
        const plot = chart.getElementsByTagNameNS(CH, 'plot-area')[0];
        const pp = autos.get(plot?.getAttributeNS(CH, 'style-name'));
        if (pp?.getAttributeNS(CH, 'vertical') === 'true' && type === 'column') ch.type = type = 'bar';
        if (pp?.getAttributeNS(CH, 'stacked') === 'true') ch.stacked = 'stacked'; if (pp?.getAttributeNS(CH, 'percentage') === 'true') ch.stacked = 'percent';
        if (chart.getAttributeNS(CH, 'class') === 'chart:filled-radar') ch.fillRadar = true;
        const title = [...chart.children].find((e) => e.localName === 'title'); if (title) ch.title = [...title.getElementsByTagNameNS(TX, 'p')].map((p) => p.textContent).join(' ');
        const sub = [...chart.children].find((e) => e.localName === 'subtitle'); if (sub) ch.subtitle = sub.textContent.trim();
        const lg = [...chart.children].find((e) => e.localName === 'legend'); ch.legend = lg ? ({ bottom: 'bottom', top: 'top', end: 'right', start: 'left' }[lg.getAttributeNS(CH, 'legend-position')] || 'right') : 'none';
        const refs = []; const lineSeries = []; let k = 0;
        const sers = [...(plot?.getElementsByTagNameNS(CH, 'series') || [])];
        for (const se of sers) {
          for (const a of ['values-cell-range-address', 'label-cell-address']) { const r = parseOdsRange(se.getAttributeNS(CH, a)); if (r) refs.push(r); }
          for (const dm of se.getElementsByTagNameNS(CH, 'domain')) { const r = parseOdsRange(dm.getAttributeNS(TNS, 'cell-range-address')); if (r) refs.push(r); }
          if (se.getAttributeNS(CH, 'class') === 'chart:line' && type === 'column') lineSeries.push(k);
          const reg = se.getElementsByTagNameNS(CH, 'regression-curve')[0];
          if (reg && !ch.trend) { const rp = autos.get(reg.getAttributeNS(CH, 'style-name')); const rt = rp?.getAttributeNS(CH, 'regression-type'); const eq = reg.getElementsByTagNameNS(CH, 'equation')[0]; ch.trend = { kind: { linear: 'linear', exponential: 'exp', logarithmic: 'log', power: 'power', polynomial: 'poly2', 'moving-average': 'movavg' }[rt] || 'linear', series: k + (type === 'scatter' ? 1 : 0), showEq: eq?.getAttributeNS(CH, 'display-equation') === 'true' || undefined, showR2: eq?.getAttributeNS(CH, 'display-r-square') === 'true' || undefined }; }
          k++;
        }
        if (lineSeries.length) { ch.type = 'combo'; ch.lineSeries = lineSeries; }
        for (const cat of plot?.getElementsByTagNameNS(CH, 'categories') || []) { const r = parseOdsRange(cat.getAttributeNS(TNS, 'cell-range-address')); if (r) refs.push(r); }
        if (!refs.length) { const r = parseOdsRange(plot?.getAttributeNS(TNS, 'cell-range-address')); if (r) refs.push(r); }
        if (!refs.length) continue;
        const sheet = refs[0].sheet; const same = refs.filter((r) => r.sheet === sheet);
        const g = { r1: Math.min(...same.map((r) => r.r1)), c1: Math.min(...same.map((r) => r.c1)), r2: Math.max(...same.map((r) => r.r2)), c2: Math.max(...same.map((r) => r.c2)) };
        ch.range = `${sheet !== s.name ? `${qs(sheet)}!` : ''}${colName(g.c1)}${g.r1 + 1}:${colName(g.c2)}${g.r2 + 1}`;
        if (plot?.getAttributeNS(CH, 'series-source') === 'rows') ch.byRows = true;
        s.charts.push({ id: Math.random().toString(36).slice(2, 10), ...ch, x: Math.round(len(fr.getAttributeNS(SVG, 'x'))), y: Math.round(len(fr.getAttributeNS(SVG, 'y'))), w: Math.max(220, Math.round(len(fr.getAttributeNS(SVG, 'width')))), h: Math.max(160, Math.round(len(fr.getAttributeNS(SVG, 'height')))) });
        n++;
      }
    }
    return n;
  } catch (e) { console.warn('gráficos ods', e); return 0; }
}
void parseRange;
