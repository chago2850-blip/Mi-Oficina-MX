// Pruebas de la fase 6 (Calcula Datos como LibreOffice Calc, parte 1): Formato de celdas completo, pegado especial,
// texto en columnas, insertar/eliminar celdas, mostrar fórmulas, color de pestaña, hipervínculos, rastrear, subtotales, filtro estándar, autocompletar.

import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase6');
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
  await w.waitForSelector('.modal', { timeout: 5000 });
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
  await w.waitForSelector('.modal', { timeout: 5000 });
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

try {
  const w = await openApp('calcula');
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  // 1) cuadro Formato de celdas: números (opciones → código), fuente, alineación con sangría y giro, bordes, fondo, protección
  await set(w, [[0, 0, '-1234.5'], [0, 1, 'Girado'], [0, 2, 'Sangría']]);
  await selR(w, 0, 0); await startDialog(w, 'formatCells');
  await w.waitForSelector('.modal'); await w.selectOption('.modal [name="cat"]', 'number'); await wait(100);
  await setDialog(w, { dec: 1, thou: true, neg: true }, false);
  assert.equal(await w.inputValue('.modal [name="code"]'), '#,##0.0;[Red]-#,##0.0');
  assert.equal(await w.evaluate(() => document.querySelector('.modal .cf-prev').textContent), '-1,234.5');
  await w.evaluate(() => document.querySelectorAll('.modal .fx-tab')[1].click());
  await setDialog(w, { fstyle: 'bi', size: 14, s: true }, false);
  await w.evaluate(() => document.querySelectorAll('.modal .fx-tab')[3].click());
  await setDialog(w, { bp: 'outer', bstyle: 'double', bcolor: '#c00000' }, false);
  await w.evaluate(() => document.querySelectorAll('.modal .fx-tab')[4].click());
  await setDialog(w, { nobg: false, bg: '#ffeeaa' }, false);
  await w.evaluate(() => document.querySelectorAll('.modal .fx-tab')[5].click());
  await setDialog(w, { locked: false, hidef: true });
  let st = await style(w, 0, 0);
  assert.equal(st.fmt, '#,##0.0;[Red]-#,##0.0'); assert.ok(st.b && st.i && st.s); assert.equal(st.size, 14); assert.equal(st.bg, '#ffeeaa');
  assert.equal(st.border.t, '#c00000 double'); assert.ok(st.unlocked && st.hidef);
  { const d = await w.evaluate(() => window.__calcula.wb.display(0, 0)); assert.equal(d.text, '-1,234.5'); assert.equal(d.color, '#C62828'); }
  await selR(w, 0, 1); await startDialog(w, 'formatCells'); await w.waitForSelector('.modal'); await w.evaluate(() => document.querySelectorAll('.modal .fx-tab')[2].click());
  await setDialog(w, { rot: 45, valign: 'middle', shrink: true });
  await selR(w, 0, 2); await startDialog(w, 'formatCells'); await w.waitForSelector('.modal'); await w.evaluate(() => document.querySelectorAll('.modal .fx-tab')[2].click());
  await setDialog(w, { align: 'left', indent: 2 });
  assert.equal((await style(w, 0, 1)).rot, 45); assert.ok((await style(w, 0, 1)).shrink); assert.equal((await style(w, 0, 2)).indent, 2);
  log('Formato de celdas: números con opciones, fuente, giro, sangría, bordes dobles, fondo y protección');
  // 2) pegado especial: valores, operación multiplicar, transponer
  await set(w, [[3, 0, '1'], [3, 1, '2'], [3, 2, '=A4+B4'], [5, 0, '10'], [5, 1, '10'], [5, 2, '10']]);
  await selR(w, 3, 0, 3, 2); await w.evaluate(() => window.__calcula.grid.copy());
  await selR(w, 5, 0); await startDialog(w, 'pasteSpecial'); await fillForm(w, { what: 'values', op: 'mul' });
  assert.deepEqual([await val(w, 5, 0), await val(w, 5, 1), await val(w, 5, 2)], [10, 20, 30]);
  await selR(w, 7, 0); await w.evaluate(() => window.__calcula.CE.doPasteSpecial({ what: 'all', transpose: true }));
  assert.deepEqual([await val(w, 7, 0), await val(w, 8, 0), await val(w, 9, 0)], [1, 2, 3]);
  log('Pegado especial: solo valores con operación (multiplicar) y transponer');
  // 3) texto en columnas
  await set(w, [[12, 0, 'Ana;García;30'], [13, 0, '"Luis; Jr";Pérez;41']]);
  await selR(w, 12, 0, 13, 0); await startDialog(w, 'textToColumns'); await fillForm(w, { sep: ';' });
  assert.deepEqual([await val(w, 12, 0), await val(w, 12, 1), await val(w, 12, 2), await val(w, 13, 0), await val(w, 13, 2)], ['Ana', 'García', 30, 'Luis; Jr', 41]);
  log('Texto en columnas (con comillas y números reconocidos)');
  // 4) insertar y eliminar celdas desplazando (las fórmulas se ajustan)
  await set(w, [[20, 0, 'a'], [21, 0, 'b'], [22, 0, 'c'], [20, 3, '=A22']]);
  await selR(w, 20, 0); await w.evaluate(() => window.__calcula.CE.shiftCells('down', true));
  assert.deepEqual([await val(w, 20, 0), await val(w, 21, 0), await val(w, 23, 0)], [null, 'a', 'c']);
  assert.equal(await raw(w, 20, 3), '=A23'); assert.equal(await val(w, 20, 3), 'b');
  await selR(w, 20, 0); await w.evaluate(() => window.__calcula.CE.shiftCells('up', false));
  assert.deepEqual([await val(w, 20, 0), await val(w, 22, 0)], ['a', 'c']); assert.equal(await raw(w, 20, 3), '=A22');
  log('Insertar y eliminar celdas desplazando hacia abajo/arriba, con referencias ajustadas');
  // 5) mostrar fórmulas, color de pestaña, hipervínculo
  await cmd(w, 'showFormulas'); assert.match((await w.evaluate(() => window.__calcula.wb.display(3, 2).text)), /^=A4\+B4$/); await cmd(w, 'showFormulas');
  assert.equal((await w.evaluate(() => window.__calcula.wb.display(3, 2).text)), '3');
  await w.evaluate(() => window.__calcula.CE.setTabColor(0, '#2E7D32')); await wait(100);
  assert.match(await w.evaluate(() => document.querySelector('.stab.on').getAttribute('style') || ''), /--tabc: ?#2E7D32/i);
  await selR(w, 25, 0); await startDialog(w, 'hyperlink'); await fillForm(w, { url: 'https://example.org', text: 'Sitio' });
  assert.equal(await val(w, 25, 0), 'Sitio'); assert.equal(await w.evaluate(() => window.__calcula.wb.sheet.x3.links['25,0']), 'https://example.org');
  log('Mostrar fórmulas, color de pestaña e hipervínculos');
  // 6) rastrear precedentes / dependientes
  await selR(w, 3, 2); await cmd(w, 'tracePre'); assert.equal(await w.evaluate(() => window.__calcula.CE.arrowCount()), 2);
  await selR(w, 3, 0); await cmd(w, 'traceDep'); assert.ok(await w.evaluate(() => window.__calcula.CE.arrowCount()) >= 3);
  await cmd(w, 'clearArrows'); assert.equal(await w.evaluate(() => window.__calcula.CE.arrowCount()), 0);
  log('Rastrear precedentes y dependientes (flechas) y quitarlas');
  // 7) subtotales y filtro estándar en una hoja nueva
  await cmd(w, 'newSheet'); await wait(200);
  await set(w, [[0, 0, 'Zona'], [0, 1, 'Ventas'], [1, 0, 'Norte'], [1, 1, '10'], [2, 0, 'Sur'], [2, 1, '5'], [3, 0, 'Norte'], [3, 1, '7'], [4, 0, 'Sur'], [4, 1, '3']]);
  await selR(w, 0, 0, 4, 1); await startDialog(w, 'subtotals'); await fillForm(w, { by: '0', fn: '9', cols: 'B' });
  const tbl = await w.evaluate(() => { const { wb } = window.__calcula; const o = []; for (let r = 0; r < 8; r++) o.push(`${wb.display(r, 0).text}:${wb.display(r, 1).text}`); return o.join('|'); });
  assert.equal(tbl, EN ? 'Zona:Ventas|Norte:10|Norte:7|Total Norte:17|Sur:5|Sur:3|Total Sur:8|Grand total:25' : 'Zona:Ventas|Norte:10|Norte:7|Total Norte:17|Sur:5|Sur:3|Total Sur:8|Total general:25');
  await selR(w, 0, 0, 7, 1); await startDialog(w, 'stdFilter'); await fillForm(w, { c1: '1', o1: '>', v1: '8' });
  assert.deepEqual(await w.evaluate(() => [...window.__calcula.wb.sheet.hiddenRows].sort((a, b) => a - b)), [2, 4, 5, 6]);
  await w.evaluate(() => window.__calcula.CE.clearStdFilter({ r1: 0, c1: 0, r2: 7, c2: 1 }));
  assert.equal(await w.evaluate(() => window.__calcula.wb.sheet.hiddenRows.size), 0);
  log('Subtotales (agrupar, SUBTOTAL y total general) y filtro estándar');
  // 8) autocompletar al escribir
  await selR(w, 9, 0); await w.keyboard.type('No'); await wait(100);
  assert.equal(await w.evaluate(() => window.__calcula.grid.editor.value), 'Norte');
  await w.keyboard.press('Enter'); assert.equal(await val(w, 9, 0), 'Norte');
  log('Autocompletar con lo que ya está en la columna');
  // 9) guardar en .xlsx y .ods y volver a abrir
  await w.evaluate(() => window.__calcula.switchSheet?.(0) ?? window.__calcula.onCommand('noop'));
  await w.evaluate(() => { window.__calcula.wb.active = 0; window.__calcula.grid.invalidate(); window.__calcula.refresh(); });
  for (const [ext, lab] of [['xlsx', 'Excel'], ['ods', 'ODS']]) {
    await w.evaluate(async ([e, l]) => window.__shell.saveAsFormat(e, l), [ext, lab]); await wait(900);
    const file = path.join(out, `${UNTITLED}.${ext}`);
    await reopen(w, file); await wait(900);
    const r = await w.evaluate(() => { const { wb } = window.__calcula; const s = wb.sheets[0]; const st = s.style(0, 0) || {}; return { fmt: st.fmt, b: st.b, i: st.i, s: st.s, size: st.size, bt: st.border?.t, bg: st.bg, unlocked: st.unlocked, hidef: st.hidef, rot: s.style(0, 1)?.rot, shrink: s.style(0, 1)?.shrink, indent: s.style(0, 2)?.indent, link: s.x3?.links?.['25,0'], tab: s.x3?.tabColor, sub: wb.raw(3, 1, wb.sheets[1]) }; });
    assert.equal(r.fmt, '#,##0.0;[Red]-#,##0.0', `${ext} ${JSON.stringify(r)}`); assert.ok(r.b && r.i && r.s, ext); assert.equal(r.size, 14, ext);
    assert.match(r.bt || '', /^#c00000 double$/i, `${ext}: ${r.bt}`); assert.equal(String(r.bg).toLowerCase(), '#ffeeaa', ext);
    assert.ok(r.unlocked && r.hidef, `${ext} protección ${JSON.stringify(r)}`); assert.equal(r.rot, 45, ext); assert.ok(r.shrink, ext); assert.equal(r.indent, 2, ext);
    assert.equal(r.link, 'https://example.org', ext); if (ext === 'xlsx') assert.equal(String(r.tab).toUpperCase(), '#2E7D32');
    assert.match(String(r.sub), /^=SUBTOTAL\(9,\s*B2:B3\)$/i, ext);
    if (hasSoffice) {
      const lo = path.join(out, `lo-${ext}`); fs.mkdirSync(lo, { recursive: true }); const to = ext === 'xlsx' ? 'ods' : 'xlsx';
      execFileSync('soffice', ['--headless', '--convert-to', to, '--outdir', lo, file], { stdio: 'ignore', timeout: 120000 });
      await reopen(w, path.join(lo, `${UNTITLED}.${to}`)); await wait(900);
      const q = await w.evaluate(() => { const s = window.__calcula.wb.sheets[0]; const st = s.style(0, 0) || {}; return { fmt: st.fmt, bt: st.border?.t, rot: s.style(0, 1)?.rot, indent: s.style(0, 2)?.indent, link: s.x3?.links?.['25,0'], b: st.b, s: st.s }; });
      assert.ok(q.b && q.s, `LO ${ext}→${to} ${JSON.stringify(q)}`); assert.match(q.bt || '', /double/, `LO ${ext}→${to} ${JSON.stringify(q)}`); assert.equal(q.rot, 45, `LO ${ext}→${to}`); assert.equal(q.indent, 2, `LO ${ext}→${to} ${JSON.stringify(q)}`); assert.match(q.link || '', /^https:\/\/example.org\/?$/, `LO ${ext}→${to}`);
    }
  }
  log(`Formato, protección, vínculos y subtotales se guardan en .xlsx y .ods${hasSoffice ? ' y sobreviven a LibreOffice' : ''}`);
  await closeApp('calcula'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de la fase 6 (Calcula, parte 1) pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
