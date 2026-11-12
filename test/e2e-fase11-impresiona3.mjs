// Pruebas de Impresiona, parte 3: diseño de tablas, álbum fotográfico, nombre y texto alternativo, navegador y página web.

import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase11');
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
const I = (w, fn, arg) => w.evaluate(fn, arg);
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
const tbl = (w) => I(w, () => { const o = window.__impresiona.pres.slides[window.__impresiona.cur].objects.find((x) => x.id === 'tb'); return o.rows.map((r) => r.map((c) => (c.hidden ? '·' : `${c.text}${c.rs > 1 || c.cs > 1 ? `[${c.rs}x${c.cs}]` : ''}`))); });
try {
  const w = await openApp('impresiona');
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(400);
  // 1) diseño de tabla
  await I(w, () => { const P = window.__impresiona; const cell = (t) => ({ text: t, rs: 1, cs: 1 }); P.pres.slides[0].objects.push({ id: 'tb', type: 'table', x: 100, y: 100, w: 480, h: 180, rot: 0, style: { size: 14 }, rows: [['A', 'B', 'C'].map(cell), ['1', '2', '3'].map(cell), ['4', '5', '6'].map(cell)] }); P.renderAll(); P.select(['tb']); P.TBL.setCurrent('tb', 1, 1); });
  assert.ok(await I(w, () => !document.querySelector('[data-tab="tableDesign"]').hidden));
  await I(w, () => window.__impresiona.onCommand('tblRowBelow'));
  await I(w, () => window.__impresiona.onCommand('tblColRight'));
  assert.deepEqual(await tbl(w), [['A', 'B', '', 'C'], ['1', '2', '', '3'], ['', '', '', ''], ['4', '5', '', '6']]);
  await I(w, () => { window.__impresiona.TBL.setCurrent('tb', 0, 0); window.__impresiona.TBL.mergeCells(2, 2); });
  assert.deepEqual(await tbl(w), [['A B 1 2[2x2]', '·', '', 'C'], ['·', '·', '', '3'], ['', '', '', ''], ['4', '5', '', '6']]);
  await I(w, () => { window.__impresiona.TBL.setCurrent('tb', 2, 0); window.__impresiona.onCommand('tblDelRow'); window.__impresiona.TBL.setCurrent('tb', 0, 2); window.__impresiona.onCommand('tblDelCol'); });
  assert.deepEqual(await tbl(w), [['A B 1 2[2x2]', '·', 'C'], ['·', '·', '3'], ['4', '5', '6']]);
  await I(w, () => { window.__impresiona.TBL.setCurrent('tb', 0, 0); window.__impresiona.onCommand('tblSplit'); });
  assert.deepEqual(await tbl(w), [['A B 1 2', '', 'C'], ['', '', '3'], ['4', '5', '6']]);
  await I(w, () => { const T = window.__impresiona.TBL; T.setOption('preset', 'azul'); T.setOption('banded', true); T.setOption('lastRow', true); T.setCurrent('tb', 2, 2); window.__impresiona.onCommand('tblFill', '#ffff00'); window.__impresiona.onCommand('tblRight'); });
  const look = await I(w, () => [...document.querySelectorAll('.slide.m-edit [data-id="tb"] tr')].map((tr) => [...tr.children].map((c) => [getComputedStyle(c).backgroundColor, getComputedStyle(c).fontWeight, c.style.textAlign])));
  assert.equal(look[0][0][0], 'rgb(47, 85, 151)'); assert.equal(look[2][0][0], 'rgb(218, 227, 243)'); assert.equal(look[1][0][0], 'rgba(0, 0, 0, 0)');
  assert.deepEqual(look[2][2], ['rgb(255, 255, 0)', '600', 'right']);
  await I(w, () => { window.__shell.doc.name = 'tabla.pptx'; });
  for (const e of ['pptx', 'odp']) await I(w, (x) => window.__shell.saveAsFormat(x, x), e);
  const px = await (await JSZip.loadAsync(fs.readFileSync(path.join(out, 'tabla.pptx')))).file('ppt/slides/slide1.xml').async('string');
  assert.match(px, /2F5597/i); assert.match(px, /DAE3F3/i); assert.match(px, /FFFF00/i); assert.match(px, /algn="r"/);
  const ox = await (await JSZip.loadAsync(fs.readFileSync(path.join(out, 'tabla.odp')))).file('content.xml').async('string');
  assert.match(ox, /draw:fill-color="#2F5597"/i); assert.match(ox, /draw:fill-color="#ffff00"/i);
  if (hasSoffice) {
    fs.mkdirSync(path.join(out, 'lo'), { recursive: true });
    execFileSync('soffice', ['--headless', '--convert-to', 'odp', '--outdir', path.join(out, 'lo'), path.join(out, 'tabla.pptx')], { stdio: 'ignore' });
    const lc = await (await JSZip.loadAsync(fs.readFileSync(path.join(out, 'lo', 'tabla.odp')))).file('content.xml').async('string');
    assert.match(lc, /#2f5597/i); assert.match(lc, /#ffff00/i);
  }
  log('Diseño de tabla: insertar y eliminar filas y columnas, combinar y dividir, estilos, bandas y formato de celda (en .pptx y .odp)');
  // 2) álbum fotográfico
  const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAADwyuo0AAAAEElEQVR4nGP4z8AARwzIHABvqgf5gNwAKAAAAABJRU5ErkJggg==';
  const n0 = await I(w, () => window.__impresiona.pres.slides.length);
  const made = await I(w, (src) => window.__impresiona.NV.photoAlbum(['playa', 'cerro', 'río', 'mar', 'sol'].map((name) => ({ name, src })), { per: 2, caption: true, title: 'Viaje' }), PNG);
  assert.equal(made, 4); assert.equal(await I(w, () => window.__impresiona.pres.slides.length), n0 + 4);
  const al = await I(w, (k) => { const s = window.__impresiona.pres.slides[k]; return s.objects.map((o) => [o.type, o.name || '', o.type === 'text' ? o.html : '', o.w / o.h]); }, n0 + 1);
  assert.deepEqual(al.map((x) => x.slice(0, 3)), [['image', 'playa', ''], ['text', '', '<p>playa</p>'], ['image', 'cerro', ''], ['text', '', '<p>cerro</p>']]);
  assert.ok(Math.abs(al[0][3] - 2) < 0.05);
  log('Álbum fotográfico: diapositiva de título, 2 fotos por diapositiva con leyenda y proporción respetada');
  // 3) nombre y texto alternativo
  await I(w, (k) => { const P = window.__impresiona; P.goto(k); P.select([P.pres.slides[k].objects[0].id]); }, n0 + 1);
  await startDialog(w, 'objName'); await w.waitForSelector('.modal');
  await setDialog(w, { name: 'Foto playa', altTitle: 'Playa', altDesc: 'Atardecer en la playa' });
  assert.deepEqual(await I(w, (k) => { const o = window.__impresiona.pres.slides[k].objects[0]; return [o.name, o.alt]; }, n0 + 1), ['Foto playa', { title: 'Playa', desc: 'Atardecer en la playa' }]);
  for (const e of ['pptx', 'odp']) await I(w, (x) => window.__shell.saveAsFormat(x, x), e);
  const pz = await JSZip.loadAsync(fs.readFileSync(path.join(out, 'tabla.pptx')));
  const sx = await pz.file(`ppt/slides/slide${n0 + 2}.xml`).async('string');
  assert.match(sx, /name="Foto playa"[^>]*descr="Atardecer en la playa"|descr="Atardecer en la playa"[^>]*name="Foto playa"/);
  const oc = await (await JSZip.loadAsync(fs.readFileSync(path.join(out, 'tabla.odp')))).file('content.xml').async('string');
  assert.match(oc, /<draw:frame draw:name="Foto playa"[^>]*>[\s\S]*?<svg:title>Playa<\/svg:title><svg:desc>Atardecer en la playa<\/svg:desc><\/draw:frame>/);
  if (hasSoffice) {
    fs.mkdirSync(path.join(out, 'lo2'), { recursive: true });
    execFileSync('soffice', ['--headless', '--convert-to', 'pptx', '--outdir', path.join(out, 'lo2'), path.join(out, 'tabla.odp')], { stdio: 'ignore' });
    const lz = await JSZip.loadAsync(fs.readFileSync(path.join(out, 'lo2', 'tabla.pptx')));
    const all = (await Promise.all(Object.keys(lz.files).filter((f) => /slides\/slide\d+\.xml$/.test(f)).map((f) => lz.file(f).async('string')))).join('');
    assert.match(all, /descr="Atardecer en la playa"/); assert.match(all, /name="Foto playa"/);
  }
  log('Nombre y texto alternativo de los objetos (en .pptx, .odp y para LibreOffice)');
  // 4) navegador
  await startDialog(w, 'navigator'); await w.waitForSelector('.modal .nav-tree');
  assert.ok(await I(w, () => [...document.querySelectorAll('.modal .nav-obj')].some((b) => b.textContent === 'Foto playa')));
  await w.fill('.modal input[type=search]', 'cerro'); await wait(200);
  assert.deepEqual(await I(w, () => [...document.querySelectorAll('.modal .nav-obj')].map((b) => b.textContent)), ['cerro', `${EN ? 'Text' : 'Texto'} 4 — cerro`]);
  await w.click('.modal .nav-obj'); await wait(300);
  assert.deepEqual(await I(w, () => [window.__impresiona.cur, window.__impresiona.selected().length, document.querySelectorAll('.modal').length]), [n0 + 1, 1, 0]);
  log('Navegador: diapositivas y objetos por nombre, búsqueda y saltar al objeto');
  // 5) página web
  const html = await I(w, () => window.__impresiona.NV.exportHtml({ save: false }));
  const total = await I(w, () => window.__impresiona.pres.slides.length);
  assert.equal((html.match(/data:image\/png;base64,/g) || []).length, total); assert.match(html, /<nav aria-label/); assert.match(html, /Atardecer en la playa/);
  await I(w, () => window.__impresiona.onCommand('exportHtml')); await wait(2500);
  assert.ok(fs.existsSync(path.join(out, 'tabla.html')));
  log('Exportar como página web: un solo .html con índice, notas, texto alternativo y navegación con el teclado');
  await closeApp('impresiona'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de Impresiona, parte 3, pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
