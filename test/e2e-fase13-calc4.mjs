// Pruebas de Calcula Datos, parte 4: gráficos (tipos, propiedades, tendencias) nativos en .xlsx y .ods.


import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase13');
fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(out, { recursive: true });
const only = process.argv[2];
const EN = process.env.OFICINA_LANG === 'en';
const UNTITLED = EN ? 'Untitled' : 'Sin título';
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const log = (m) => console.log(`✔ ${m}`);
const exe = process.env.OFICINA_EXE;
// cámara y micrófono de prueba (Chromium genera video y un tono)
const FAKE = ['--use-fake-device-for-media-stream', '--use-fake-ui-for-media-stream'];
const app = await electron.launch({ ...(exe ? { executablePath: exe, args: [`--user-data-dir=${path.join(out, 'perfil')}`, '--no-sandbox', ...FAKE] } : { args: [root, `--user-data-dir=${path.join(out, 'perfil')}`, '--no-sandbox', ...FAKE] }), cwd: root, env: { ...process.env, OFICINA_TEST_DIR: out, OFICINA_LANG: process.env.OFICINA_LANG || 'es' } });
// errores de JavaScript en cualquier ventana: la prueba falla si aparece alguno
const pageErrors = [];
const watch = (w) => { w.on('pageerror', (e) => pageErrors.push(`${w.url().split('/').slice(-2).join('/')}: ${e.message}`)); w.on('console', (m) => { if (m.type() === 'error' && !/favicon|Autofill|dictionary|Dictionar|getUserMedia|Requested device not found/i.test(m.text())) pageErrors.push(`${w.url().split('/').slice(-2).join('/')} console: ${m.text().slice(0, 300)}`); }); };
app.on('window', watch); app.windows().forEach(watch);
async function waitWindow(pred, ms = 15000) {
  const t0 = Date.now();
  while (Date.now() - t0 < ms) { for (const w of app.windows()) { try { if (pred(w.url())) return w; } catch { /* */ } } await wait(150); }
  throw new Error('No apareció la ventana');
}
async function openApp(id) {
  const l = await waitWindow((u) => u.includes('/launcher/'));
  await l.waitForLoadState('domcontentloaded'); await wait(500);
  await l.evaluate((x) => window.__launcher.launch(x), id);
  const w = await waitWindow((u) => u.includes(`/${id}/`));
  await w.waitForLoadState('domcontentloaded'); await wait(900);
  return w;
}
const closeApp = (id) => app.evaluate(({ BrowserWindow }, x) => BrowserWindow.getAllWindows().find((w) => w.webContents.getURL().includes(`/${x}/`))?.close(), id);
const reopen = (w, p) => w.evaluate(async (f) => { const d = await window.oficina.readPath(f); await window.__shell.loadFile(d); }, p);
// Abre un formulario sin esperar a que se cierre
const startDialog = (w, cmd) => w.evaluate(([c]) => { setTimeout(() => window.__shell.command(c)); }, [cmd]);
async function fillForm(w, values, submit = true) {
  await w.waitForSelector('.modal', { timeout: 5000 }); await wait(200); // el cuadro pone el foco en su primer campo justo después de abrirse
  for (const [name, v] of Object.entries(values)) {
    const sel = `.modal [name="${name}"]`;
    const tag = await w.evaluate((s) => document.querySelector(s)?.tagName, sel);
    if (tag === 'SELECT') await w.selectOption(sel, v); else if (typeof v === 'boolean') { if (v !== await w.isChecked(sel)) await w.click(sel); } else await w.fill(sel, v);
  }
  if (submit) await w.click('.modal .btn.primary');
  await wait(300);
}
const py = (code) => execFileSync('python3', ['-c', code], { encoding: 'utf8' }).trim();

const hasSoffice = (() => { try { execFileSync('soffice', ['--version'], { stdio: 'ignore' }); return true; } catch { return false; } })();


// pone valores en el cuadro abierto (también en pestañas ocultas) y acepta
async function setDialog(w, values, submit = true) {
  await w.waitForSelector('.modal', { timeout: 5000 }); await wait(200); // el cuadro pone el foco en su primer campo justo después de abrirse
  await w.evaluate((vals) => {
    const m = [...document.querySelectorAll('.modal')].pop();
    for (const [k, v] of Object.entries(vals)) { const el = m.querySelector(`[name="${k}"]`); if (!el) throw new Error(`sin campo ${k}`); if (el.type === 'checkbox') el.checked = !!v; else el.value = String(v); el.dispatchEvent(new Event('change', { bubbles: true })); el.dispatchEvent(new Event('input', { bubbles: true })); }
  }, values);
  if (submit) { await w.click('.modal .btn.primary'); await wait(300); }
}
const set = (w, cells) => w.evaluate((cs) => { const { wb } = window.__calcula; wb.action('datos', () => cs.forEach(([r, c, v]) => wb.setInput(r, c, String(v))), { batch: true }); window.__calcula.refresh(); }, cells);
const val = (w, r, c) => w.evaluate(([a, b]) => window.__calcula.wb.value(a, b), [r, c]);
const raw = (w, r, c) => w.evaluate(([a, b]) => window.__calcula.wb.raw(a, b), [r, c]);
const style = (w, r, c) => w.evaluate(([a, b]) => window.__calcula.wb.sheet.style(a, b), [r, c]);
const selR = (w, r1, c1, r2 = r1, c2 = c1) => w.evaluate(([a, b, c, d]) => { const { grid } = window.__calcula; grid.select(a, b); if (c !== a || d !== b) grid.select(c, d, true); }, [r1, c1, r2, c2]);
const cmd = (w, c) => w.evaluate((x) => window.__calcula.onCommand(x), c);

const near = (a, b, eps = 1e-6) => assert.ok(Math.abs(a - b) < eps, `${a} ≈ ${b}`);
const col = (w, c, r1, r2) => w.evaluate(([cc, a, b]) => { const o = []; for (let r = a; r <= b; r++) o.push(window.__calcula.wb.value(r, cc)); return o; }, [c, r1, r2]);
const opt = (w, i = 0) => w.evaluate((k) => { const { wb, chartOption } = window.__calcula; const o = chartOption(wb.sheet.charts[k]); return JSON.parse(JSON.stringify(o, (key, v) => (typeof v === 'function' ? 'fn' : v))); }, i);
const chart = (w, i = 0) => w.evaluate((k) => JSON.parse(JSON.stringify(window.__calcula.wb.sheet.charts[k] || null)), i);
const zipText = async (file, name) => { const z = await JSZip.loadAsync(fs.readFileSync(file)); return z.file(name) ? z.file(name).async('string') : null; };
try {
  const w = await openApp('calcula');
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  const rows = [['Mes', 'Ventas', 'Costos', 'Margen'], ['Ene', 100, 60, 0.4], ['Feb', 120, 70, 0.42], ['Mar', 150, 80, 0.47], ['Abr', 160, 85, 0.47], ['May', 190, 95, 0.5], ['Jun', 210, 100, 0.52]];
  await set(w, rows.flatMap((row, r) => row.map((v, c) => [r, c, v])));
  // 1) insertar con el cuadro de propiedades: columnas y líneas, eje secundario, títulos, leyenda, etiquetas y tendencia
  await selR(w, 0, 0, 6, 3); await startDialog(w, 'insertChart'); await w.waitForSelector('.modal [name="color2"]', { state: 'attached' });
  await setDialog(w, { type: 'combo', title: 'Ventas 2026', subtitle: 'Primer semestre', xTitle: 'Mes', yTitle: 'Pesos', y2Title: 'Margen', legend: 'top', labels: 'value', line0: false, line1: true, line2: true, sec2: true, color0: '#1f77b4', trend: 'linear', trendSeries: '0', showEq: true, showR2: true });
  let ch = await chart(w);
  assert.deepEqual([ch.type, ch.title, ch.subtitle, ch.legend, ch.labels, ch.lineSeries, ch.secondary, ch.trend.kind, ch.colors[0]], ['combo', 'Ventas 2026', 'Primer semestre', 'top', 'value', [1, 2], [2], 'linear', '#1f77b4']);
  let o = await opt(w);
  assert.deepEqual(o.series.slice(0, 3).map((x) => [x.type, x.yAxisIndex]), [['bar', 0], ['line', 0], ['line', 1]]);
  assert.equal(o.yAxis.length, 2); assert.equal(o.title.subtext, 'Primer semestre'); assert.equal(o.legend.top, 34);
  assert.match(o.series[3].name, /y = 22x \+ 78 · R² = 0\.99/);
  log('Gráfico combinado (columnas y líneas) con eje secundario, títulos, leyenda, etiquetas y línea de tendencia con ecuación y R²');
  // 2) apilado al 100 % y otros tipos
  await w.evaluate(() => { const c = window.__calcula.wb.sheet.charts[0]; Object.assign(c, { type: 'column', stacked: 'percent', range: 'A1:C7', trend: undefined, secondary: undefined }); });
  o = await opt(w); assert.equal(o.series[0].stack, 'total'); assert.equal(Math.round(o.series[0].data[0] + o.series[1].data[0]), 100);
  for (const type of ['bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'bubble', 'radar']) {
    await w.evaluate((t) => { Object.assign(window.__calcula.wb.sheet.charts[0], { type: t, stacked: undefined, range: 'A1:D7' }); }, type);
    o = await opt(w); assert.ok(o.series.length >= 1, type);
    if (type === 'doughnut') assert.equal(o.series[0].radius[0], '38%');
    if (type === 'radar') assert.equal(o.radar.indicator.length, 6);
    if (type === 'bubble') assert.equal(o.series[0].data[0].length, 3);
  }
  await w.evaluate(() => { window.__calcula.wb.sheet.charts[0].type = 'line'; window.__calcula.wb.sheet.charts[0].byRows = true; window.__calcula.wb.sheet.charts[0].range = 'A1:D3'; });
  o = await opt(w); assert.deepEqual(o.series.map((x) => x.name), ['Ene', 'Feb']); assert.deepEqual(o.xAxis.data, ['Ventas', 'Costos', 'Margen']);
  log('Apilado al 100 %, barras, líneas, área, circular, anillo, dispersión, burbujas, radar y series en filas');
  // 3) gráficos de verdad en .xlsx y .ods, que LibreOffice lee
  await w.evaluate(() => { const s = window.__calcula.wb.sheet; s.charts = [{ id: 'c1', type: 'combo', range: 'A1:D7', title: 'Ventas 2026', x: 420, y: 20, w: 480, h: 300, lineSeries: [2], secondary: [2], legend: 'bottom', trend: { kind: 'linear', series: 0, showEq: true } }, { id: 'c2', type: 'pie', range: 'A1:B7', title: 'Reparto', x: 420, y: 340, w: 360, h: 260 }]; window.__calcula.afterLoad(); });
  await w.evaluate(() => { window.__shell.doc.name = 'graf.xlsx'; });
  for (const e of ['xlsx', 'ods']) await w.evaluate((x) => window.__shell.saveAsFormat(x, x), e);
  const c1 = await zipText(path.join(out, 'graf.xlsx'), 'xl/charts/chart1.xml');
  assert.match(c1, /<c:barChart><c:barDir val="col"\/>/); assert.match(c1, /<c:lineChart>/); assert.match(c1, /<c:f>Hoja1!\$B\$2:\$B\$7<\/c:f>|<c:f>[^<]*!\$B\$2:\$B\$7<\/c:f>/); assert.match(c1, /<c:trendlineType val="linear"\/>/); assert.match(c1, /<c:axPos val="r"\/>/); assert.match(c1, /Ventas 2026/);
  assert.match(await zipText(path.join(out, 'graf.xlsx'), 'xl/charts/chart2.xml'), /<c:pieChart>/);
  const oc = await zipText(path.join(out, 'graf.ods'), 'Object 1/content.xml');
  assert.match(oc, /chart:class="chart:bar"/); assert.match(oc, /chart:class="chart:line"/); assert.match(oc, /chart:attached-axis="secondary-y"/); assert.match(oc, /<chart:regression-curve/);
  assert.match(await zipText(path.join(out, 'graf.ods'), 'content.xml'), /<table:shapes><draw:frame[^>]*><draw:object xlink:href="\.\/Object 1"/);
  if (hasSoffice) {
    fs.mkdirSync(path.join(out, 'lo'), { recursive: true });
    execFileSync('soffice', ['--headless', '--convert-to', 'ods', '--outdir', path.join(out, 'lo'), path.join(out, 'graf.xlsx')], { stdio: 'ignore' });
    execFileSync('soffice', ['--headless', '--convert-to', 'xlsx', '--outdir', path.join(out, 'lo'), path.join(out, 'graf.ods')], { stdio: 'ignore' });
    const loXlsx = await JSZip.loadAsync(fs.readFileSync(path.join(out, 'lo', 'graf.xlsx')));
    const chartFiles = Object.keys(loXlsx.files).filter((f) => /xl\/charts\/chart\d+\.xml$/.test(f));
    assert.equal(chartFiles.length, 2, 'LibreOffice lee los gráficos del .ods');
    const loOds = await JSZip.loadAsync(fs.readFileSync(path.join(out, 'lo', 'graf.ods')));
    assert.ok(Object.keys(loOds.files).some((f) => /^Object \d+\/content\.xml$/.test(f)), 'LibreOffice lee los gráficos del .xlsx');
    for (const f of ['graf.ods', 'graf.xlsx']) {
      await reopen(w, path.join(out, 'lo', f)); await wait(600);
      const got = await w.evaluate(() => window.__calcula.wb.sheets[0].charts.map((c) => [c.type, c.title, c.range, c.trend?.kind || null]).sort((a, b) => String(a[1]).localeCompare(String(b[1]))));
      assert.deepEqual(got.map((g) => [g[0], g[1]]), [['pie', 'Reparto'], ['combo', 'Ventas 2026']], f);
      assert.ok(got[1][2].startsWith('A1:D7') || /A1:D7$/.test(got[1][2]), got[1][2]);
      assert.equal(got[1][3], 'linear', f);
      assert.ok(await w.evaluate(() => window.__calcula.charts().size === 2));
    }
    log('LibreOffice abre los gráficos de .xlsx y .ods de Calcula, y Calcula abre los que guarda LibreOffice');
  }
  log('Gráficos de verdad en .xlsx (DrawingML) y .ods (objetos de gráfico de ODF)');
  await closeApp('calcula'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de Calcula, parte 4, pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
