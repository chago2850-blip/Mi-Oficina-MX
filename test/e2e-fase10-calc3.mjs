// Pruebas de Calcula Datos, parte 3: escenarios, operaciones múltiples, listas de ordenamiento, detective y fecha/hora.


import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase10');
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
try {
  const w = await openApp('calcula');
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  // 1) escenarios: préstamo con tasa y plazo
  await set(w, [[0, 0, 'Monto'], [0, 1, 100000], [1, 0, 'Tasa'], [1, 1, 0.1], [2, 0, 'Años'], [2, 1, 2], [4, 0, 'Total'], [4, 1, '=B1*(1+B2)^B3']]);
  await selR(w, 1, 1, 2, 1); await startDialog(w, 'scenarioNew'); await setDialog(w, { name: 'Base', comment: 'Actual' });
  await set(w, [[1, 1, 0.2], [2, 1, 3]]);
  await w.evaluate(() => window.__calcula.AN.createScenario('Caro', 'B2:B3', 'Tasa alta'));
  near(await val(w, 4, 1), 172800, 1e-6);
  await w.evaluate(() => window.__calcula.AN.showScenario('Base'));
  near(await val(w, 4, 1), 121000, 1e-6);
  const cols = await w.evaluate(() => window.__calcula.AN.scenarioSummary({ r1: 4, c1: 1, r2: 4, c2: 1 }).map((c) => c.outs[0]));
  near(cols[0], 121000); near(cols[1], 172800);
  const sum = await w.evaluate(() => { const { wb } = window.__calcula; const o = []; for (let r = 2; r <= 7; r++) o.push([wb.value(r, 0), wb.value(r, 1), wb.value(r, 2)]); return [wb.sheet.name, o]; });
  assert.match(sum[0], /Resumen de escenarios|Scenario summary/); assert.deepEqual(sum[1][0].slice(1), ['Base', 'Caro']); assert.deepEqual(sum[1][1], ['B2', 0.1, 0.2]); assert.deepEqual(sum[1][5], ['B5', 121000, 172800]);
  await w.evaluate(() => window.__calcula.onCommand && window.__calcula.wb && (() => {})());
  // el libro original no cambió y los escenarios se guardan
  await w.evaluate(() => { const { wb } = window.__calcula; window.__shellSwitch = wb.sheets[0]; });
  for (const ext of ['xlsx', 'ods']) {
    await w.evaluate(() => { window.__shell.doc.name = 'esc.xlsx'; }); await w.evaluate((e) => window.__shell.saveAsFormat(e, e), ext);
    await reopen(w, path.join(out, `esc.${ext}`)); await wait(400);
    assert.deepEqual(await w.evaluate(() => (window.__calcula.wb.sheets[0].x3?.scenarios || []).map((s) => [s.name, s.range, s.values])), [['Base', 'B2:B3', [[0.1], [2]]], ['Caro', 'B2:B3', [[0.2], [3]]]], ext);
  }
  if (hasSoffice) { fs.mkdirSync(path.join(out, 'lo'), { recursive: true }); execFileSync('soffice', ['--headless', '--convert-to', 'xlsx', '--outdir', path.join(out, 'lo'), path.join(out, 'esc.ods')], { stdio: 'ignore' }); assert.ok(fs.existsSync(path.join(out, 'lo', 'esc.xlsx')), 'LibreOffice abre el .ods con escenarios'); }
  log('Escenarios: crear, mostrar y resumen en una hoja nueva (se guardan en .xlsx y .ods)');
  // 2) operaciones múltiples de una y dos variables
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  await set(w, [[0, 0, 5], [0, 1, 3], [0, 2, '=A1*B1'], [0, 3, '=A1+B1'], [3, 0, 1], [4, 0, 2], [5, 0, 3]]);
  await selR(w, 3, 0, 5, 2); await startDialog(w, 'multipleOps'); await setDialog(w, { range: 'A4:C6', formulas: 'C1:D1', colCell: 'A1', rowCell: '' });
  assert.deepEqual(await w.evaluate(() => { const { wb } = window.__calcula; return [3, 4, 5].map((r) => [wb.value(r, 1), wb.value(r, 2)]); }), [[3, 4], [6, 5], [9, 6]]);
  assert.equal(await val(w, 0, 2), 15);
  await set(w, [[8, 1, 10], [8, 2, 20], [9, 0, 1], [10, 0, 2]]);
  await w.evaluate(() => window.__calcula.AN.multipleOps({ range: 'A9:C11', formulas: 'C1', rowCell: 'B1', colCell: 'A1' }));
  assert.deepEqual(await w.evaluate(() => { const { wb } = window.__calcula; return [9, 10].map((r) => [wb.value(r, 1), wb.value(r, 2)]); }), [[10, 20], [20, 40]]);
  log('Operaciones múltiples (tabla de datos) de una variable con varias fórmulas y de dos variables');
  // 3) listas de ordenamiento: autorrellenar y ordenar
  await startDialog(w, 'sortLists'); await setDialog(w, { lists: 'Bajo, Medio, Alto, Urgente' });
  await set(w, [[20, 0, 'Bajo']]); await selR(w, 20, 0, 23, 0); await startDialog(w, 'fillSeries'); await setDialog(w, { type: 'auto' });
  assert.deepEqual(await col(w, 0, 20, 23), ['Bajo', 'Medio', 'Alto', 'Urgente']);
  await set(w, [[30, 0, 'Prioridad'], [31, 0, 'Alto'], [32, 0, 'Bajo'], [33, 0, 'Urgente'], [34, 0, 'Medio']]);
  await selR(w, 31, 0); await startDialog(w, 'sortDialog'); await setDialog(w, { k1: '0', o1: 'asc', header: true, list: '0' });
  assert.deepEqual(await col(w, 0, 30, 34), ['Prioridad', 'Bajo', 'Medio', 'Alto', 'Urgente']);
  log('Listas de ordenamiento propias: autorrellenar y ordenar con orden personalizado');
  // 4) detective: marcar datos no válidos y rastrear error
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  await set(w, [[0, 0, 5], [1, 0, 50], [2, 0, 7], [3, 0, 99]]);
  await w.evaluate(() => window.__calcula.F.addValidation({ range: 'A1:A4', type: 'whole', min: 1, max: 10 }));
  await cmd(w, 'markInvalid');
  assert.deepEqual(await w.evaluate(() => window.__calcula.AN.markCount()), { invalid: 2, errors: 0 });
  await set(w, [[0, 2, '=1/0'], [0, 3, '=C1+1'], [0, 4, '=D1*2']]);
  await selR(w, 0, 4); await cmd(w, 'traceError');
  assert.equal((await w.evaluate(() => window.__calcula.AN.markCount())).errors, 2);
  await cmd(w, 'clearMarks'); assert.deepEqual(await w.evaluate(() => window.__calcula.AN.markCount()), { invalid: 0, errors: 0 });
  log('Detective: marcar datos no válidos y rastrear el origen de un error');
  // 5) fecha y hora con el teclado
  await selR(w, 5, 0); await w.focus('canvas'); await w.keyboard.press('Control+;');
  const today = new Date(); const ds = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${today.getFullYear()}`;
  assert.equal(await w.evaluate(() => window.__calcula.wb.display(5, 0).text), ds);
  await selR(w, 6, 0); await cmd(w, 'insertTime');
  assert.match(await w.evaluate(() => window.__calcula.wb.display(6, 0).text), /^\d\d:\d\d:\d\d$/);
  log('Insertar la fecha de hoy (Ctrl+;) y la hora actual');
  await closeApp('calcula'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de Calcula, parte 3, pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
