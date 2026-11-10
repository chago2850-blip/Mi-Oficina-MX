// Prueba de punta a punta de Mi Oficina MX con Playwright + Electron.
// Linux sin pantalla:  xvfb-run -a node test/e2e.mjs
import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { PDFDocument } from 'pdf-lib';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'e2e');
fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(out, { recursive: true });
const userData = path.join(out, 'perfil');
const UNTITLED = process.env.OFICINA_LANG === 'en' ? 'Untitled' : 'Sin título';
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const log = (m) => console.log(`✔ ${m}`);

// OFICINA_EXE=ruta/al/ejecutable prueba la app ya empaquetada (instalador)
const exe = process.env.OFICINA_EXE;
const app = await electron.launch({ ...(exe ? { executablePath: exe, args: [`--user-data-dir=${userData}`, "--no-sandbox"] } : { args: [root, `--user-data-dir=${userData}`, "--no-sandbox"] }), cwd: root, env: { ...process.env, OFICINA_TEST_DIR: out, OFICINA_LANG: process.env.OFICINA_LANG || 'es' } });
const main = (fn, arg) => app.evaluate(fn, arg);
// errores de JavaScript en cualquier ventana: la prueba falla si aparece alguno
const pageErrors = [];
const watch = (w) => { w.on('pageerror', (e) => pageErrors.push(`${w.url().split('/').slice(-2).join('/')}: ${e.message}`)); w.on('console', (m) => { if (m.type() === 'error' && !/favicon|Autofill|dictionary|Dictionar|getUserMedia|Requested device not found/i.test(m.text())) pageErrors.push(`${w.url().split('/').slice(-2).join('/')} console: ${m.text().slice(0, 300)}`); }); };
app.on('window', watch); app.windows().forEach(watch);
async function waitWindow(pred, ms = 15000) {
  const t0 = Date.now();
  while (Date.now() - t0 < ms) {
    for (const w of app.windows()) { try { if (pred(w.url())) return w; } catch { /* cerrada */ } }
    await wait(150);
  }
  throw new Error('No apareció la ventana esperada');
}
const launcherState = () => main(({ BrowserWindow }) => {
  const w = BrowserWindow.getAllWindows().find((x) => x.webContents.getURL().includes('/launcher/'));
  return w ? { visible: w.isVisible(), minimized: w.isMinimized() } : null;
});

try {
  // ---------------------------------------------------------------- Launcher
  const launcher = await waitWindow((u) => u.includes('/launcher/'));
  await launcher.waitForLoadState('domcontentloaded');
  await wait(600);
  assert.equal(await launcher.evaluate(() => document.getElementById('stage').classList.contains('closed')), true);
  await launcher.click('#core');
  await wait(700);
  assert.equal(await launcher.evaluate(() => document.getElementById('stage').classList.contains('open')), true);
  assert.equal(await launcher.locator('.sector').count(), 4);
  await launcher.screenshot({ path: path.join(out, 'launcher.png') });
  log('la dona se enciende y muestra las 4 apps');

  // ---------------------------------------------------------------- Escribiente
  await launcher.click('[data-app="escribiente"]', { force: true });
  const esc = await waitWindow((u) => u.includes('/escribiente/'));
  await esc.waitForLoadState('domcontentloaded');
  await wait(800);
  let ls = await launcherState();
  assert.ok(ls.minimized || !ls.visible, 'la dona debe minimizarse al abrir una app');
  log('la dona se minimiza al abrir Escribiente');

  await esc.click('.tpl[data-tpl="blank"]');
  await wait(300);
  await esc.click('#editor');
  await esc.keyboard.type('Hola Mi Oficina MX');
  await esc.keyboard.press('Control+a');
  await esc.evaluate(() => window.__escribiente.onCommand('bold'));
  await esc.keyboard.press('End');
  await esc.keyboard.press('Enter');
  // texto largo para forzar varias páginas
  await esc.evaluate(() => {
    const ed = window.__escribiente.editor;
    for (let i = 0; i < 90; i++) ed.insertAdjacentHTML('beforeend', `<p>Párrafo de prueba número ${i + 1}. Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.</p>`);
    ed.insertAdjacentHTML('beforeend', '<table><tbody><tr><th>A</th><th>B</th></tr><tr><td>1</td><td>2</td></tr></tbody></table><ul><li>uno</li><li>dos</li></ul>');
    ed.dispatchEvent(new Event('input'));
    window.__escribiente.paginate();
  });
  const pages = await esc.evaluate(() => window.__escribiente.state.pages);
  assert.ok(pages >= 3, `deben ser 3+ páginas, fueron ${pages}`);
  assert.ok(await esc.evaluate(() => !!document.querySelector('#editor b, #editor strong, #editor span[style*="bold"]')), 'negrita aplicada');
  await esc.screenshot({ path: path.join(out, 'escribiente.png') });
  log(`Escribiente: texto, negritas y paginación (${pages} páginas)`);

  await esc.evaluate(() => window.__shell.save());
  const docxPath = path.join(out, `${UNTITLED}.docx`);
  assert.ok(fs.existsSync(docxPath), 'se guardó el .docx');
  const docx = fs.readFileSync(docxPath);
  assert.equal(docx.subarray(0, 2).toString(), 'PK');
  await esc.evaluate(() => window.__shell.exportPdf());
  const pdfPath = path.join(out, `${UNTITLED}.pdf`);
  assert.ok(fs.existsSync(pdfPath)); assert.equal(fs.readFileSync(pdfPath).subarray(0, 5).toString(), '%PDF-');
  const pdfPages = (await PDFDocument.load(fs.readFileSync(pdfPath))).getPageCount();
  assert.equal(pdfPages, pages, `el PDF debe tener las mismas páginas que la pantalla (${pdfPages} vs ${pages})`);
  log(`Escribiente: guarda .docx y exporta PDF con ${pdfPages} páginas`);

  // reabrir el .docx
  await esc.evaluate(async (p) => { const f = await window.oficina.readPath(p); await window.__shell.loadFile(f); }, docxPath);
  await wait(500);
  const reopened = await esc.evaluate(() => ({ text: document.getElementById('editor').innerText, tables: document.querySelectorAll('#editor table').length, lists: document.querySelectorAll('#editor ul li').length, bold: !!document.querySelector('#editor strong, #editor b') }));
  assert.ok(reopened.text.includes('Hola Mi Oficina MX'));
  assert.ok(reopened.text.includes('Párrafo de prueba número 90'));
  assert.equal(reopened.tables, 1); assert.equal(reopened.lists, 2); assert.ok(reopened.bold);
  log('Escribiente: el .docx se vuelve a abrir con texto, negritas, tabla y lista');

  // ---------------------------------------------------------------- la dona vuelve al cerrar
  await main(({ BrowserWindow }) => BrowserWindow.getAllWindows().find((w) => w.webContents.getURL().includes('/escribiente/')).close());
  await wait(900);
  ls = await launcherState();
  assert.ok(ls.visible && !ls.minimized, 'la dona debe volver al cerrar la última app');
  log('la dona se restaura al cerrar la app');

  // ---------------------------------------------------------------- Calcula Datos
  await launcher.evaluate(() => window.__launcher.launch('calcula'));
  const calc = await waitWindow((u) => u.includes('/calcula/'));
  await calc.waitForLoadState('domcontentloaded');
  await wait(800);
  await calc.click('.tpl[data-tpl="blank"]');
  await wait(300);
  await calc.focus('.grid-canvas');
  for (const v of ['10', '20', '30.5']) { await calc.keyboard.type(v); await calc.keyboard.press('Enter'); }
  await calc.keyboard.type('=SUMA(A1:A3)'); await calc.keyboard.press('Enter');
  await calc.keyboard.type('=SI(A4>50,"alto","bajo")'); await calc.keyboard.press('Enter');
  await calc.keyboard.type('$1,500.00'); await calc.keyboard.press('Enter');
  const vals = await calc.evaluate(() => { const { wb } = window.__calcula; return [wb.value(3, 0), wb.value(4, 0), wb.raw(3, 0), wb.editText(3, 0), wb.display(5, 0).text]; });
  assert.equal(vals[0], 60.5); assert.equal(vals[1], 'alto'); assert.equal(vals[2], '=SUM(A1:A3)'); assert.equal(vals[3], process.env.OFICINA_LANG === 'en' ? '=SUM(A1:A3)' : '=SUMA(A1:A3)'); assert.equal(vals[4], '$1,500.00');
  log('Calcula Datos: fórmulas en español (SUMA, SI) y formato de moneda');
  // deshacer / rehacer
  await calc.evaluate(() => window.__calcula.onCommand('undo'));
  assert.equal(await calc.evaluate(() => window.__calcula.wb.raw(5, 0)), null);
  await calc.evaluate(() => window.__calcula.onCommand('redo'));
  assert.equal(await calc.evaluate(() => window.__calcula.wb.raw(5, 0)), 1500);
  // relleno de serie
  await calc.evaluate(() => { const { wb, grid } = window.__calcula; wb.action('x', () => { wb.setInput(0, 2, 'enero'); wb.setInput(0, 3, '1'); wb.setInput(1, 3, '2'); }); grid.fill({ r1: 0, c1: 2, r2: 0, c2: 2 }, { r1: 0, c1: 2, r2: 3, c2: 2 }); grid.fill({ r1: 0, c1: 3, r2: 1, c2: 3 }, { r1: 0, c1: 3, r2: 4, c2: 3 }); });
  const series = await calc.evaluate(() => { const { wb } = window.__calcula; return [wb.value(3, 2), wb.value(4, 3)]; });
  assert.deepEqual(series, ['abril', 5]);
  log('Calcula Datos: deshacer/rehacer y relleno de series (enero→abril, 1,2→5)');
  await calc.evaluate(() => window.__calcula.onCommand('insertChart', 'column'));
  await wait(400);
  await calc.screenshot({ path: path.join(out, 'calcula.png') });
  await calc.evaluate(() => window.__shell.save());
  const xlsxPath = path.join(out, `${UNTITLED}.xlsx`);
  assert.ok(fs.existsSync(xlsxPath));
  await calc.evaluate(async (p) => { const f = await window.oficina.readPath(p); await window.__shell.loadFile(f); }, xlsxPath);
  await wait(400);
  const re = await calc.evaluate(() => { const { wb } = window.__calcula; return [wb.raw(3, 0), wb.value(3, 0), wb.display(5, 0).text, wb.sheet.charts.length]; });
  assert.equal(re[0], '=SUM(A1:A3)'); assert.equal(re[1], 60.5); assert.equal(re[2], '$1,500.00'); assert.equal(re[3], 1);
  log('Calcula Datos: guarda y reabre .xlsx con fórmulas, formato y gráfico');
  await calc.evaluate(() => window.__shell.exportPdf());
  assert.ok(fs.existsSync(path.join(out, `${UNTITLED}.pdf`)));
  log('Calcula Datos: exporta PDF');
  await main(({ BrowserWindow }) => BrowserWindow.getAllWindows().find((w) => w.webContents.getURL().includes('/calcula/')).close());
  await wait(600);

  // ---------------------------------------------------------------- Impresiona
  await launcher.evaluate(() => window.__launcher.launch('impresiona'));
  const imp = await waitWindow((u) => u.includes('/impresiona/'));
  await imp.waitForLoadState('domcontentloaded');
  await wait(800);
  await imp.click('.tpl[data-tpl="negocio"]');
  await wait(400);
  assert.equal(await imp.locator('.thumbs .thumb').count(), 4);
  await imp.evaluate(() => window.__impresiona.onCommand('newSlide'));
  await imp.evaluate(() => window.__impresiona.onCommand('textbox'));
  await wait(200);
  await imp.keyboard.type('Cuadro de prueba');
  await imp.keyboard.press('Escape');
  await wait(200);
  const n = await imp.evaluate(() => window.__impresiona.pres.slides.length);
  assert.equal(n, 5);
  assert.ok(await imp.evaluate(() => window.__impresiona.pres.slides[1].objects.some((o) => (o.html || '').includes('Cuadro de prueba'))));
  await imp.screenshot({ path: path.join(out, 'impresiona.png') });
  log('Impresiona: plantilla, nueva diapositiva y cuadro de texto');
  // presentación
  await imp.evaluate(() => window.__impresiona.present(0));
  await wait(500);
  assert.equal(await imp.locator('.show').count(), 1);
  await imp.evaluate(() => { window.__show.next(); window.__show.next(); });
  await wait(900);
  await imp.screenshot({ path: path.join(out, 'impresiona-presentacion.png') });
  await imp.keyboard.press('Escape');
  await wait(300);
  assert.equal(await imp.locator('.show').count(), 0);
  log('Impresiona: presentación a pantalla completa con transiciones');
  await imp.evaluate(() => window.__shell.save());
  const pptxPath = path.join(out, `${UNTITLED}.pptx`);
  assert.ok(fs.existsSync(pptxPath));
  await imp.evaluate(async (p) => { const f = await window.oficina.readPath(p); await window.__shell.loadFile(f); }, pptxPath);
  await wait(600);
  const back = await imp.evaluate(() => ({ n: window.__impresiona.pres.slides.length, text: window.__impresiona.pres.slides.map((s) => s.objects.map((o) => o.html || '').join(' ')).join(' ') }));
  assert.equal(back.n, 5);
  assert.ok(back.text.includes(process.env.OFICINA_LANG === 'en' ? 'Business plan 2027' : 'Plan de negocio 2027') && back.text.includes('Cuadro de prueba'));
  log('Impresiona: guarda y reabre .pptx con sus 5 diapositivas');
  await imp.evaluate(() => window.__shell.exportPdf());
  assert.ok(fs.existsSync(path.join(out, `${UNTITLED}.pdf`)));
  log('Impresiona: exporta PDF');
  await main(({ BrowserWindow }) => BrowserWindow.getAllWindows().find((w) => w.webContents.getURL().includes('/impresiona/')).close());
  await wait(600);

  // ---------------------------------------------------------------- PapelPDF
  await launcher.evaluate(() => window.__launcher.launch('papelpdf'));
  const pdf = await waitWindow((u) => u.includes('/papelpdf/'));
  await pdf.waitForLoadState('domcontentloaded');
  await wait(1200);
  await pdf.evaluate(() => {});
  await main(({ BrowserWindow }, p) => {
    const w = BrowserWindow.getAllWindows().find((x) => x.webContents.getURL().includes('/papelpdf/'));
    w.webContents.send('open-path', p);
  }, pdfPath);
  await wait(2500);
  await pdf.screenshot({ path: path.join(out, 'papelpdf.png') });
  const pdfOk = await pdf.evaluate(() => document.querySelectorAll('canvas').length > 0);
  assert.ok(pdfOk, 'PapelPDF muestra el PDF exportado');
  log('PapelPDF: se abre desde la dona y muestra un PDF');
  const comp = await pdf.evaluate(async (bytes) => { const r = await window.papel.qpdf({ op: 'compress', bytes: new Uint8Array(bytes) }); return { head: String.fromCharCode(...r.bytes.slice(0, 5)), n: r.bytes.length }; }, [...fs.readFileSync(pdfPath)]);
  assert.equal(comp.head, '%PDF-');
  const ocrOk = await pdf.evaluate(async () => (await fetch('../../vendor/ocr/tessdata/spa.traineddata.gz', { method: 'HEAD' })).ok);
  assert.ok(ocrOk, 'modelos de OCR incluidos');
  log('PapelPDF: compresión (qpdf) y modelos de OCR disponibles');
  await main(({ BrowserWindow }) => BrowserWindow.getAllWindows().find((w) => w.webContents.getURL().includes('/papelpdf/')).close());
  await wait(800);
  ls = await launcherState();
  assert.ok(ls.visible && !ls.minimized);
  log('la dona vuelve después de cerrar PapelPDF');
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de punta a punta pasaron.');
} catch (err) {
  console.error('✘', err);
  for (const w of app.windows()) { try { await w.screenshot({ path: path.join(out, `error-${path.basename(new URL(w.url()).pathname.replace(/\/index\.html$/, ''))}.png`) }); } catch { /* */ } }
  process.exitCode = 1;
} finally {
  await app.close().catch(() => {});
}
