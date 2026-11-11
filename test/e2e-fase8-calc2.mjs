// Pruebas de Calcula Datos, parte 2 (como LibreOffice Calc): ordenar, rellenar, aleatorios, estadísticas,
// consolidar, duplicados, estilos de celda, mayúsculas y hojas ocultas, movidas y copiadas.


import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase8');
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
  // 1) ordenar con varios criterios
  await set(w, [[0, 0, 'Nombre'], [0, 1, 'Depto'], [0, 2, 'Sueldo'], [1, 0, 'Ana'], [1, 1, 'Ventas'], [1, 2, 100], [2, 0, 'Beto'], [2, 1, 'Admin'], [2, 2, 300], [3, 0, 'Caro'], [3, 1, 'Ventas'], [3, 2, 250], [4, 0, 'item10'], [4, 1, 'Admin'], [4, 2, 50], [5, 0, 'item2'], [5, 1, 'Admin'], [5, 2, 50]]);
  await selR(w, 1, 1); await startDialog(w, 'sortDialog');
  await setDialog(w, { k1: '1', o1: 'asc', k2: '2', o2: 'desc', k3: '0', o3: 'asc' });
  assert.deepEqual(await col(w, 0, 0, 5), ['Nombre', 'Beto', 'item2', 'item10', 'Caro', 'Ana']);
  await w.evaluate(() => window.__calcula.DT.sortMulti({ r1: 0, c1: 0, r2: 5, c2: 2 }, [{ c: 0, asc: false }], { header: true, natural: true, dest: { r: 0, c: 5 } }));
  assert.deepEqual(await col(w, 5, 0, 5), ['Nombre', 'item10', 'item2', 'Caro', 'Beto', 'Ana']);
  log('Ordenar con tres criterios, orden natural y copiar el resultado a otro lugar');
  // 2) rellenar series y rellenar hacia abajo
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  await selR(w, 0, 0, 4, 0); await startDialog(w, 'fillSeries'); await setDialog(w, { dir: 'down', type: 'linear', start: '2', step: '3' });
  assert.deepEqual(await col(w, 0, 0, 4), [2, 5, 8, 11, 14]);
  await selR(w, 0, 1, 3, 1); await startDialog(w, 'fillSeries'); await setDialog(w, { type: 'date', unit: 'month', start: '2026-01-31', step: '1' });
  const d = await w.evaluate(() => [0, 1, 2, 3].map((r) => window.__calcula.wb.display(r, 1).text));
  assert.deepEqual(d, ['31/01/2026', '28/02/2026', '31/03/2026', '30/04/2026']);
  await selR(w, 0, 2); await startDialog(w, 'fillSeries'); await setDialog(w, { type: 'growth', start: '1', step: '2', end: '100' });
  assert.deepEqual(await col(w, 2, 0, 7), [1, 2, 4, 8, 16, 32, 64, null]);
  await set(w, [[0, 3, 'lunes'], [0, 4, 'Pieza 007']]);
  await selR(w, 0, 3, 2, 3); await startDialog(w, 'fillSeries'); await setDialog(w, { type: 'auto' });
  await selR(w, 0, 4, 2, 4); await startDialog(w, 'fillSeries'); await setDialog(w, { type: 'auto' });
  assert.deepEqual(await col(w, 3, 0, 2), ['lunes', 'martes', 'miércoles']); assert.deepEqual(await col(w, 4, 0, 2), ['Pieza 007', 'Pieza 008', 'Pieza 009']);
  await set(w, [[0, 5, '=A1*10']]); await selR(w, 0, 5, 4, 5); await cmd(w, 'fillDown');
  assert.deepEqual(await col(w, 5, 0, 4), [20, 50, 80, 110, 140]); assert.equal(await raw(w, 4, 5), '=A5*10');
  log('Rellenar serie (lineal, fecha por mes, crecimiento hasta un valor final, autorrelleno) y rellenar hacia abajo con fórmulas');
  // 3) números aleatorios con semilla
  await selR(w, 10, 0, 29, 0); await startDialog(w, 'fillRandom'); await setDialog(w, { dist: 'int', a: '1', b: '6', seed: 'dado' });
  const r1 = await col(w, 0, 10, 29);
  assert.ok(r1.every((x) => Number.isInteger(x) && x >= 1 && x <= 6)); assert.ok(new Set(r1).size > 2);
  await startDialog(w, 'fillRandom'); await setDialog(w, { dist: 'int', a: '1', b: '6', seed: 'dado' });
  assert.deepEqual(await col(w, 0, 10, 29), r1);
  log('Números aleatorios (uniforme entera) con semilla reproducible');
  // 4) estadísticas
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  const data = [[0, 0, 'X'], [0, 1, 'Y']]; for (let i = 1; i <= 10; i++) data.push([i, 0, i], [i, 1, 2 * i + 1 + (i === 4 ? 0 : 0)]);
  await set(w, data);
  await selR(w, 0, 0, 10, 1); await startDialog(w, 'stats'); await setDialog(w, { kind: 'descriptive', input: 'A1:B11', output: 'D1', by: 'cols', labels: true });
  const ROWS = ['Media', 'Error estándar', 'Moda', 'Mediana', 'Primer cuartil', 'Tercer cuartil', 'Varianza', 'Desviación estándar', 'Curtosis', 'Asimetría', 'Rango', 'Mínimo', 'Máximo', 'Suma', 'Cuenta'];
  const st = await w.evaluate((R) => { const { wb } = window.__calcula; const o = {}; for (let r = 1; r <= 15; r++) o[R[r - 1]] = wb.value(r, 4); return o; }, ROWS);
  near(st.Media, 5.5); near(st.Mediana, 5.5); near(st.Curtosis, -1.2); near(st.Asimetría, 0); near(st.Varianza, 9.166666666); near(st.Cuenta, 10); near(st['Primer cuartil'], 3.25);
  assert.equal(await raw(w, 9, 4), '=KURT($A$2:$A$11)');
  await startDialog(w, 'stats'); await setDialog(w, { kind: 'regression', input: 'A1:B11', output: 'D20', labels: true });
  near(await val(w, 20, 4), 2); near(await val(w, 21, 4), 1); near(await val(w, 22, 4), 1);
  await startDialog(w, 'stats'); await setDialog(w, { kind: 'moving', input: 'A1:A11', output: 'H1', labels: true, n: 3 });
  near(await val(w, 3, 7), 2); near(await val(w, 10, 7), 9);
  await startDialog(w, 'stats'); await setDialog(w, { kind: 'correlation', input: 'A1:B11', output: 'J1', labels: true });
  near(await val(w, 2, 10), 1);
  await set(w, [[30, 0, 1], [31, 0, 2], [32, 0, 2], [33, 0, 3], [30, 1, '=MODA(A31:A34)']]);
  assert.equal(await val(w, 30, 1), 2);
  log('Estadísticas con fórmulas: descriptiva (CURTOSIS, MODA, cuartiles), regresión, media móvil y correlación');
  if (hasSoffice) {
    await set(w, [[40, 0, '=INTERCEPT(B2:B11,A2:A11)'], [40, 1, '=FORECAST(20,B2:B11,A2:A11)']]);
    await w.evaluate(() => { window.__shell.doc.name = 'estad.xlsx'; }); await w.evaluate(() => window.__shell.saveAsFormat('xlsx', 'Excel'));
    execFileSync('soffice', ['--headless', '--convert-to', 'csv', '--outdir', out, path.join(out, 'estad.xlsx')], { stdio: 'ignore' });
    const csv = fs.readFileSync(path.join(out, 'estad.csv'), 'utf8').split(/\r?\n/).map((l) => l.split(','));
    near(Number(csv[9][4]), -1.2, 1e-3); near(Number(csv[40][0]), 1, 1e-9); near(Number(csv[40][1]), 41, 1e-9); near(Number(csv[30][1]), 2, 1e-9);
    log('LibreOffice calcula igual las fórmulas estadísticas (CURTOSIS, INTERSECCION.EJE, PRONOSTICO, MODA)');
  }
  // 5) consolidar dos hojas
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  await set(w, [[0, 0, 'Producto'], [0, 1, 'Ene'], [1, 0, 'Pan'], [1, 1, 10], [2, 0, 'Leche'], [2, 1, 5]]);
  await w.evaluate(() => { const { wb } = window.__calcula; const s = wb.addSheet('Sucursal 2'); wb.action('d', () => { [['Producto', 'Ene', 'Feb'], ['Leche', 7, 1], ['Queso', 3, 2]].forEach((row, r) => row.forEach((v, c) => wb.setRaw(r, c, v, s))); }); });
  const s1 = await w.evaluate(() => window.__calcula.wb.sheets[0].name);
  await selR(w, 0, 0); await startDialog(w, 'consolidate');
  await setDialog(w, { fn: 'SUM', sources: `${s1}!A1:B3\n'Sucursal 2'!A1:C3`, dest: 'E1', rowLabels: true, colLabels: true });
  const cons = await w.evaluate(() => { const { wb } = window.__calcula; const o = []; for (let r = 0; r < 4; r++) { const row = []; for (let c = 4; c < 7; c++) row.push(wb.value(r, c)); o.push(row); } return o; });
  assert.deepEqual(cons, [[null, 'Ene', 'Feb'], ['Pan', 10, null], ['Leche', 12, 1], ['Queso', 3, 2]]);
  log('Consolidar por etiquetas de fila y columna de dos hojas');
  // 6) duplicados, estilos de celda y mayúsculas
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  await set(w, [[0, 0, 'Nombre'], [0, 1, 'Ciudad'], [1, 0, 'ana'], [1, 1, 'puebla'], [2, 0, 'ANA'], [2, 1, 'Puebla'], [3, 0, 'luis'], [3, 1, 'león'], [4, 0, 'ana'], [4, 1, 'Mérida']]);
  await selR(w, 1, 0); await startDialog(w, 'duplicates'); await setDialog(w, { header: true, cs: false, action: 'remove' });
  assert.deepEqual(await col(w, 0, 0, 4), ['Nombre', 'ana', 'luis', 'ana', null]);
  await selR(w, 1, 0, 3, 1); await cmd(w, 'caseTitle');
  assert.deepEqual(await col(w, 1, 1, 3), ['Puebla', 'León', 'Mérida']); assert.equal(await val(w, 1, 0), 'Ana');
  await selR(w, 0, 0, 0, 1); await w.evaluate(() => window.__calcula.onCommand('cellStyle', 'Encabezado'));
  await selR(w, 1, 0); await w.evaluate(() => window.__calcula.onCommand('cellStyle', 'Bueno'));
  assert.equal((await style(w, 0, 0)).bg, '#D9E1F2'); assert.ok((await style(w, 0, 1)).b); assert.equal((await style(w, 1, 0)).bg, '#C6EFCE');
  log('Quitar duplicados, Tipo Título y estilos de celda (Encabezado, Bueno)');
  // 7) hojas: mover, copiar, ocultar y mostrar (se guardan en .xlsx y .ods)
  await w.evaluate(() => { const { wb } = window.__calcula; wb.addSheet('B'); wb.addSheet('C'); });
  const names0 = await w.evaluate(() => window.__calcula.wb.sheets.map((s) => s.name));
  await w.evaluate(() => window.__calcula.DT.moveCopySheet(2, 0, false, ''));
  assert.deepEqual(await w.evaluate(() => window.__calcula.wb.sheets.map((s) => s.name)), [names0[2], names0[0], names0[1]]);
  await w.evaluate(() => window.__calcula.DT.moveCopySheet(0, 3, true, 'Copia C'));
  assert.deepEqual(await w.evaluate(() => window.__calcula.wb.sheets.map((s) => s.name)), [names0[2], names0[0], names0[1], 'Copia C']);
  await w.evaluate(() => window.__calcula.DT.hideSheet(1)); await wait(200);
  assert.equal(await w.evaluate(() => document.querySelectorAll('.stab').length), 3);
  for (const ext of ['xlsx', 'ods']) {
    await w.evaluate(() => { window.__shell.doc.name = 'hojas.xlsx'; });
    await w.evaluate((e) => window.__shell.saveAsFormat(e, e), ext);
    await reopen(w, path.join(out, `hojas.${ext}`)); await wait(500);
    assert.deepEqual(await w.evaluate(() => window.__calcula.wb.sheets.map((s) => [s.name, !!s.x3?.hidden])), [[names0[2], false], [names0[0], true], [names0[1], false], ['Copia C', false]], ext);
  }
  if (hasSoffice) {
    fs.mkdirSync(path.join(out, 'lo'), { recursive: true });
    execFileSync('soffice', ['--headless', '--convert-to', 'xlsx', '--outdir', path.join(out, 'lo'), path.join(out, 'hojas.ods')], { stdio: 'ignore' });
    const z = await JSZip.loadAsync(fs.readFileSync(path.join(out, 'lo', 'hojas.xlsx')));
    assert.match(await z.file('xl/workbook.xml').async('string'), new RegExp(`name="${names0[0]}"[^>]*state="hidden"|state="hidden"[^>]*name="${names0[0]}"`));
  }
  await startDialog(w, 'showSheets'); await setDialog(w, {});
  assert.equal(await w.evaluate(() => document.querySelectorAll('.stab').length), 4);
  log('Mover, copiar, ocultar y mostrar hojas (se conservan en .xlsx y .ods, también con LibreOffice)');
  await closeApp('calcula'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de Calcula, parte 2, pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
