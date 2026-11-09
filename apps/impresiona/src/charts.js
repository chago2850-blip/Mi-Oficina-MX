// Gráficos para Impresiona: se dibujan como SVG estático (ECharts en modo servidor). GPL-3.0
import * as echarts from 'echarts/dist/echarts.esm.min.mjs';

export const CHART_TYPES = [['column', 'Columnas'], ['bar', 'Barras'], ['line', 'Líneas'], ['area', 'Área'], ['pie', 'Circular'], ['doughnut', 'Anillo'], ['scatter', 'Dispersión']];

// "Mes,2025,2026\nEne,10,14" -> { cats, series:[{name, values}] }
export function parseChartData(text) {
  const rows = String(text || '').trim().split(/\r?\n/).map((l) => l.split(/[,;\t]/).map((x) => x.trim()));
  if (rows.length < 2) return { cats: [], series: [] };
  const head = rows[0];
  const cats = rows.slice(1).map((r) => r[0]);
  const series = head.slice(1).map((name, i) => ({ name: name || `Serie ${i + 1}`, values: rows.slice(1).map((r) => { const n = Number(String(r[i + 1] ?? '').replace(/[$,%]/g, '')); return Number.isFinite(n) ? n : null; }) }));
  return { cats, series };
}
export function chartDataText(d) {
  return [['', ...d.series.map((s) => s.name)].join(','), ...d.cats.map((c, i) => [c, ...d.series.map((s) => s.values[i] ?? '')].join(','))].join('\n');
}

const cache = new Map();
export function chartSvg(o, theme, w = o.w, h = o.h) {
  const key = JSON.stringify([o.chartType, o.title, o.data, theme.id, Math.round(w), Math.round(h)]);
  if (cache.has(key)) return cache.get(key);
  const d = o.data || { cats: [], series: [] };
  const palette = [theme.accent, theme.accent2, '#F08A24', '#1F8FD6', '#2FB4A8', '#8E6CD9', '#E0B000', '#DC3A63'];
  const txt = theme.text, axis = { axisLabel: { color: txt, fontSize: 13 }, axisLine: { lineStyle: { color: `${txt}66` } }, splitLine: { lineStyle: { color: `${txt}22` } } };
  const base = {
    animation: false, color: palette, backgroundColor: 'transparent', textStyle: { fontFamily: 'Inter, Carlito, sans-serif' },
    title: o.title ? { text: o.title, left: 'center', top: 6, textStyle: { color: theme.title, fontSize: 18, fontWeight: 600 } } : undefined,
    legend: d.series.length > 1 || o.chartType === 'pie' || o.chartType === 'doughnut' ? { bottom: 4, textStyle: { color: txt, fontSize: 13 } } : undefined,
    grid: { left: 16, right: 20, top: o.title ? 50 : 20, bottom: d.series.length > 1 ? 40 : 16, containLabel: true },
  };
  let opt;
  if (o.chartType === 'pie' || o.chartType === 'doughnut') opt = { ...base, series: [{ type: 'pie', radius: o.chartType === 'doughnut' ? ['40%', '66%'] : ['0%', '66%'], center: ['50%', '52%'], label: { color: txt, formatter: '{b}: {d}%' }, data: d.cats.map((c, i) => ({ name: c, value: d.series[0]?.values[i] })) }] };
  else if (o.chartType === 'scatter') opt = { ...base, xAxis: { type: 'value', ...axis }, yAxis: { type: 'value', ...axis }, series: d.series.slice(1).map((s) => ({ type: 'scatter', name: s.name, symbolSize: 10, data: s.values.map((v, i) => [d.series[0].values[i], v]) })) };
  else {
    const horiz = o.chartType === 'bar';
    const cat = { type: 'category', data: d.cats, ...axis }, val = { type: 'value', ...axis };
    opt = { ...base, xAxis: horiz ? val : cat, yAxis: horiz ? cat : val, series: d.series.map((s) => ({ type: o.chartType === 'column' || o.chartType === 'bar' ? 'bar' : 'line', name: s.name, data: s.values, barMaxWidth: 56, smooth: false, symbolSize: 7, areaStyle: o.chartType === 'area' ? { opacity: 0.3 } : undefined })) };
  }
  const chart = echarts.init(null, null, { renderer: 'svg', ssr: true, width: Math.max(80, Math.round(w)), height: Math.max(60, Math.round(h)) });
  chart.setOption(opt);
  const svg = chart.renderToSVGString();
  chart.dispose();
  if (cache.size > 200) cache.clear();
  cache.set(key, svg);
  return svg;
}
export const svgDataUrl = (svg) => `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(svg)))}`;
