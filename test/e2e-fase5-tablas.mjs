// Pruebas de la fase 5, parte 3 (Escribiente como LibreOffice Writer): tablas — selección y combinación de
// celdas, filas y columnas con celdas combinadas, ordenar, fórmulas, estilos, propiedades, texto ↔ tabla.
// Linux sin pantalla:  xvfb-run -a node test/e2e-fase5.mjs
import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase5c');
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
const BASE = '<table><tbody><tr><th>Producto</th><th>Cantidad</th><th>Precio</th></tr><tr><td>Peras</td><td>3</td><td>10.5</td></tr><tr><td>Manzanas</td><td>12</td><td>4</td></tr><tr><td>Uvas</td><td>7</td><td>25</td></tr></tbody></table><p>fin</p>';
// la cuadrícula de cada tabla: cada fila cubre todas las columnas exactamente una vez
const shape = (w, i = 0) => w.evaluate((k) => { const TB = window.__escribiente.TB; const tb = document.querySelectorAll('#editor table')[k]; const G = TB.grid(tb); const ok = G.m.every((r) => r.length === G.cols && r.every(Boolean)); return { ok, rows: G.rows.length, cols: G.cols, cells: G.cells.map((x) => `${x.td.textContent.trim() || '·'}${x.cs > 1 ? `/c${x.cs}` : ''}${x.rs > 1 ? `/r${x.rs}` : ''}`).join('|') }; }, i);
const cellRect = (w, r, c) => w.evaluate(([R, Cc]) => { const tb = document.querySelector('#editor table'); const G = window.__escribiente.TB.grid(tb); const b = G.m[R][Cc].td.getBoundingClientRect(); return { x: b.left + b.width / 2, y: b.top + b.height / 2, l: b.left, r: b.right }; }, [r, c]);
const caretIn = (w, r, c) => w.evaluate(([R, Cc]) => { const tb = document.querySelector('#editor table'); const G = window.__escribiente.TB.grid(tb); const td = G.m[R][Cc].td; const rg = document.createRange(); rg.selectNodeContents(td); rg.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(rg); window.__escribiente.saveSelection(); }, [r, c]);
const cmd = (w, c) => w.evaluate((x) => window.__shell.command(x), c);

try {
  const w = await openApp('escribiente');
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  await w.evaluate((h) => window.__escribiente.setContent(h), BASE); await wait(300);
  // 1) seleccionar celdas arrastrando y combinarlas
  const a = await cellRect(w, 1, 1); const b = await cellRect(w, 2, 2);
  await w.mouse.move(a.x, a.y); await w.mouse.down(); await w.mouse.move((a.x + b.x) / 2, (a.y + b.y) / 2, { steps: 4 }); await w.mouse.move(b.x, b.y, { steps: 4 }); await w.mouse.up();
  assert.equal(await w.evaluate(() => document.querySelectorAll('#editor .csel').length), 4);
  await cmd(w, 'mergeCells'); await wait(200);
  let s = await shape(w); assert.ok(s.ok, JSON.stringify(s));
  assert.match(s.cells, /^Producto\|Cantidad\|Precio\|Peras\|310\.5124\/c2\/r2\|Manzanas\|Uvas\|7\|25$/, s.cells);
  // 2) dividir la celda combinada: en 2 columnas y luego en 2 filas
  await caretIn(w, 1, 1); await w.evaluate(() => { const td = document.querySelector('#editor td[colspan]'); window.__escribiente.TB.splitCell(td, 2, 'v'); });
  s = await shape(w); assert.ok(s.ok, JSON.stringify(s)); assert.match(s.cells, /\/r2\|·\/r2\|/, s.cells);
  await w.evaluate(() => { const td = document.querySelector('#editor td[rowspan]'); window.__escribiente.TB.splitCell(td, 2, 'h'); });
  s = await shape(w); assert.ok(s.ok, JSON.stringify(s)); assert.equal(s.rows, 4);
  // dividir una celda normal en 3 columnas ensancha las demás filas con colspan
  await w.evaluate(() => { const td = document.querySelectorAll('#editor tr')[3].children[0]; window.__escribiente.TB.splitCell(td, 3, 'v'); });
  s = await shape(w); assert.ok(s.ok, JSON.stringify(s)); assert.equal(s.cols, 5);
  // otra dirección: una celda normal en 2 filas
  await w.evaluate(() => { const td = document.querySelectorAll('#editor tr')[0].children[0]; window.__escribiente.TB.splitCell(td, 2, 'h'); });
  s = await shape(w); assert.ok(s.ok, JSON.stringify(s)); assert.equal(s.rows, 5);
  // una celda de 2 filas dividida en 3
  await w.evaluate((h) => window.__escribiente.setContent(h), BASE); await wait(100);
  await w.evaluate(() => { const G = window.__escribiente.TB.grid(document.querySelector('#editor table')); G.m[1][2].td.classList.add('csel'); G.m[2][2].td.classList.add('csel'); }); await cmd(w, 'mergeCells');
  await w.evaluate(() => window.__escribiente.TB.splitCell(document.querySelector('#editor td[rowspan]'), 3, 'h'));
  s = await shape(w); assert.ok(s.ok && s.rows === 5, JSON.stringify(s));
  await w.evaluate(() => { const G = window.__escribiente.TB.grid(document.querySelector('#editor table')); G.m[3][0].td.classList.add('csel'); G.m[3][1].td.classList.add('csel'); }); await cmd(w, 'mergeCells');
  await w.evaluate(() => window.__escribiente.TB.splitCell(document.querySelector('#editor td[colspan]'), 3, 'v'));
  s = await shape(w); assert.ok(s.ok && s.cols === 4, JSON.stringify(s));
  log('Tablas: seleccionar celdas arrastrando, combinar, dividir en filas y columnas sin romper la cuadrícula');
  // 3) filas y columnas cuando hay celdas combinadas
  await w.evaluate((h) => window.__escribiente.setContent(h), BASE); await wait(200);
  await w.evaluate(() => { const G = window.__escribiente.TB.grid(document.querySelector('#editor table')); G.m[1][1].td.classList.add('csel'); G.m[2][1].td.classList.add('csel'); }); await cmd(w, 'mergeCells');
  await caretIn(w, 1, 0); await cmd(w, 'rowBelow'); s = await shape(w); assert.ok(s.ok, JSON.stringify(s)); assert.match(s.cells, /\/r3/, s.cells);
  await caretIn(w, 1, 1); await cmd(w, 'colLeft'); s = await shape(w); assert.ok(s.ok && s.cols === 4, JSON.stringify(s));
  await caretIn(w, 1, 0); await cmd(w, 'delRow'); s = await shape(w); assert.ok(s.ok && s.rows === 4, JSON.stringify(s)); assert.match(s.cells, /312\/r2/, s.cells);
  await caretIn(w, 1, 2); await cmd(w, 'delCol'); s = await shape(w); assert.ok(s.ok && s.cols === 3, JSON.stringify(s));
  await caretIn(w, 0, 2); await cmd(w, 'colRight'); await caretIn(w, 0, 0); await cmd(w, 'delCol'); s = await shape(w); assert.ok(s.ok && s.cols === 3, JSON.stringify(s));
  log('Tablas: insertar y eliminar filas y columnas respetando las celdas combinadas');
  // 4) ordenar, fórmulas
  await w.evaluate((h) => window.__escribiente.setContent(h), BASE); await wait(200);
  await caretIn(w, 1, 1); await startDialog(w, 'sortTable'); await fillForm(w, { order: 'desc', head: true }, true);
  assert.equal(await w.evaluate(() => [...document.querySelectorAll('#editor tr')].map((r) => r.children[0].textContent).join(',')), 'Producto,Manzanas,Uvas,Peras');
  await caretIn(w, 3, 2); await cmd(w, 'rowBelow');
  await caretIn(w, 4, 1); await startDialog(w, 'tableFormula'); await fillForm(w, { f: '=SUM(ABOVE)' });
  await caretIn(w, 4, 2); await startDialog(w, 'tableFormula'); await fillForm(w, { f: '=B2*C2+B3*C3+B4*C4' });
  const vals = () => w.evaluate(() => { const G = window.__escribiente.TB.grid(document.querySelector('#editor table')); return [G.m[4][1].td.textContent, G.m[4][2].td.textContent]; });
  assert.deepEqual(await vals(), ['22', '254.5']);
  await w.evaluate(() => { const G = window.__escribiente.TB.grid(document.querySelector('#editor table')); G.m[1][1].td.textContent = '20'; G.m[1][1].td.dispatchEvent(new InputEvent('input', { bubbles: true })); });
  await wait(500); assert.deepEqual(await vals(), ['30', '286.5']);
  log('Tablas: ordenar (numérico, descendente, con encabezado) y fórmulas que se recalculan (=SUM(ABOVE), =B2*C2+…)');
  // 5) estilo de tabla, bordes y propiedades
  await caretIn(w, 1, 0); await w.evaluate(() => window.__escribiente.TB.applyTableStyle('blue'));
  const st = await w.evaluate(() => { const th = document.querySelector('#editor th'); const td = document.querySelectorAll('#editor tr')[2].children[0]; return [getComputedStyle(th).backgroundColor, getComputedStyle(th).color, getComputedStyle(td).backgroundColor]; });
  assert.deepEqual(st, ['rgb(47, 84, 150)', 'rgb(255, 255, 255)', 'rgb(222, 234, 246)']);
  await caretIn(w, 1, 0); await startDialog(w, 'tableProps');
  await setDialog(w, { tw: 80, talign: 'center', va: 'middle', bp: 'outer', bw: 1.5, bc: '#c00000', bto: 'table' });
  const pr = await w.evaluate(() => { const tb = document.querySelector('#editor table'); const c = getComputedStyle(document.querySelectorAll('#editor tr')[1].children[0]); const c2 = getComputedStyle(document.querySelectorAll('#editor tr')[1].children[1]); return { w: tb.style.width, ml: tb.style.marginLeft, va: c.verticalAlign, bl: c.borderLeftStyle + ' ' + c.borderLeftColor + ' ' + c.borderLeftWidth, inner: c2.borderLeftStyle }; });
  assert.equal(pr.w, '80%'); assert.equal(pr.ml, 'auto'); assert.equal(pr.va, 'middle'); assert.equal(pr.bl, 'solid rgb(192, 0, 0) 2px'); assert.equal(pr.inner, 'none');
  // 6) ancho de columna arrastrando el borde
  const before = await w.evaluate(() => document.querySelector('#editor tr').children[0].getBoundingClientRect().width);
  const e1 = await cellRect(w, 1, 0); await w.mouse.move(e1.r, e1.y); await w.mouse.down(); await w.mouse.move(e1.r + 60, e1.y, { steps: 5 }); await w.mouse.up();
  const after = await w.evaluate(() => document.querySelector('#editor tr').children[0].getBoundingClientRect().width);
  assert.ok(Math.abs(after - before - 60) < 6, `${before} → ${after}`);
  log('Tablas: estilos de tabla, cuadro Propiedades (ancho, alineación, celda, bordes) y ancho de columna arrastrando el borde');
  // 7) guardar en .docx y .odt y volver a abrir (combinadas, fórmulas, colores, anchos)
  await w.evaluate(() => { const G = window.__escribiente.TB.grid(document.querySelector('#editor table')); G.m[2][0].td.classList.add('csel'); G.m[2][1].td.classList.add('csel'); }); await cmd(w, 'mergeCells');
  const exp = await shape(w);
  const ratio0 = await w.evaluate(() => { const tb = document.querySelector('#editor table'); const G = window.__escribiente.TB.grid(tb); return Math.round(G.m[1][0].td.getBoundingClientRect().width / tb.getBoundingClientRect().width * 100); });
  const check = async (lab) => {
    const s2 = await shape(w); assert.ok(s2.ok, `${lab}: ${JSON.stringify(s2)}`); assert.equal(s2.cells.replace(/\s/g, ''), exp.cells.replace(/\s/g, ''), lab);
    const r = await w.evaluate(() => { const tb = document.querySelector('#editor table'); const G = window.__escribiente.TB.grid(tb); return { f: G.m[4][1].td.dataset.formula, th: getComputedStyle(tb.querySelector('tr').children[0]).backgroundColor, w0: Math.round(G.m[1][0].td.getBoundingClientRect().width / tb.getBoundingClientRect().width * 100) }; });
    assert.equal(r.th, 'rgb(47, 84, 150)', lab); assert.ok(Math.abs(r.w0 - ratio0) <= 2, `${lab} ancho ${r.w0} ≠ ${ratio0}`);
    return r;
  };
  for (const [ext, lab] of [['docx', 'Word'], ['odt', 'ODT']]) {
    await w.evaluate(async ([e, l]) => window.__shell.saveAsFormat(e, l), [ext, lab]); await wait(800);
    const file = path.join(out, `${UNTITLED}.${ext}`);
    await reopen(w, file); await wait(800);
    const r = await check(ext); assert.equal(r.f, '=SUM(ABOVE)', ext);
    if (hasSoffice) {
      const lo = path.join(out, `lo-${ext}`); fs.mkdirSync(lo, { recursive: true }); const to = ext === 'docx' ? 'odt' : 'docx';
      execFileSync('soffice', ['--headless', '--convert-to', to, '--outdir', lo, file], { stdio: 'ignore', timeout: 120000 });
      await reopen(w, path.join(lo, `${UNTITLED}.${to}`)); await wait(900);
      await check(`LibreOffice ${ext}→${to}`);
      await reopen(w, file); await wait(600);
    }
  }
  log(`Tablas: celdas combinadas, fórmulas, colores y anchos se guardan en .docx y .odt${hasSoffice ? ' y sobreviven a LibreOffice' : ''}`);
  // 8) texto → tabla y tabla → texto
  await w.evaluate(() => window.__escribiente.setContent('<p>a;b;c</p><p>1;2;3</p><p>fin</p>')); await wait(200);
  await w.evaluate(() => { const ps = document.querySelectorAll('#editor p'); const r = document.createRange(); r.setStart(ps[0].firstChild, 0); r.setEnd(ps[1].firstChild, 3); const s3 = getSelection(); s3.removeAllRanges(); s3.addRange(r); window.__escribiente.saveSelection(); });
  await startDialog(w, 'textToTable'); await fillForm(w, { sep: ';' });
  s = await shape(w); assert.equal(s.cells, 'a|b|c|1|2|3'); assert.equal(await w.evaluate(() => document.querySelectorAll('#editor th').length), 3);
  await caretIn(w, 1, 1); await startDialog(w, 'tableToText'); await fillForm(w, { sep: ';' });
  assert.equal(await w.evaluate(() => [...document.querySelectorAll('#editor p')].map((p) => p.textContent).join('/')), 'a; b; c/1; 2; 3/fin');
  log('Tablas: convertir texto en tabla y tabla en texto');
  // 9) tabla larga: se parte por filas entre páginas y repite la fila de encabezado
  await w.evaluate(() => { let h = '<p>Intro</p>'.repeat(20) + '<table><tbody><tr><th>N</th><th>Descripción</th></tr>'; for (let i = 1; i <= 60; i++) h += `<tr><td>${i}</td><td>Fila ${i}</td></tr>`; window.__escribiente.setContent(`${h}</tbody></table><p>Después</p>`); window.__escribiente.paginate(); });
  await wait(500);
  const pg = await w.evaluate(() => { const S = window.__escribiente.state; const rows = [...document.querySelectorAll('#editor tr:not(.tpush):not(.trep)')]; return { push: document.querySelectorAll('#editor tr.tpush').length, rep: document.querySelectorAll('#editor tr.trep').length, rows: rows.length, cells: window.__escribiente.TB.grid(document.querySelector('#editor table')).rows.length }; });
  assert.ok(pg.push >= 1 && pg.rep === pg.push, JSON.stringify(pg)); assert.equal(pg.cells, 61);
  await w.evaluate(() => { const td = document.querySelectorAll('#editor tr')[30].children[0]; const r = document.createRange(); r.selectNodeContents(td); getSelection().removeAllRanges(); getSelection().addRange(r); window.__escribiente.saveSelection(); });
  await w.evaluate(async () => window.__shell.saveAsFormat('docx', 'Word')); await wait(800);
  const z9 = await JSZip.loadAsync(fs.readFileSync(path.join(out, `${UNTITLED}.docx`))); const d9 = await z9.file('word/document.xml').async('string');
  assert.equal((d9.match(/<w:tr[ >]/g) || []).length, 61, 'el .docx no lleva filas de relleno'); assert.match(d9, /<w:tblHeader\/>/);
  log('Tablas: una tabla larga se parte por filas entre páginas y repite el encabezado (sin filas de más al guardar)');
  await closeApp('escribiente'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de la fase 5 (parte 3) pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
