// Pruebas de Impresiona, parte 2 (como LibreOffice Impress): presentaciones personalizadas, configuración
// (bucle), esquema, resumen y expandir, insertar desde archivo, duplicar, campos e imprimir notas y folletos.

import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase9');
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
const I = (w, fn, arg) => w.evaluate(fn, arg);
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAAASFvFNAAAAFklEQVR4nGP4z8DAwMDAxMDAwMDAAAANHQEDK+mmyAAAAABJRU5ErkJggg==';
try {
  const w = await openApp('impresiona');
  await w.evaluate(() => window.__shell.newDoc('negocio')); await wait(500);
  const ids = await I(w, () => window.__impresiona.pres.slides.map((s) => s.id));
  const N = ids.length; assert.ok(N >= 4);
  // 1) presentación personalizada: definir con el cuadro (marcar y ordenar)
  await I(w, () => { window.__csP = window.__impresiona.DS.editCustomShow({ name: 'x', ids: [] }); });
  await w.waitForSelector('.modal .cs-row'); await wait(200);
  await I(w, ([a, b]) => { const m = document.querySelector('.modal'); m.querySelector('[name=name]').value = 'Corta'; m.querySelector(`input[data-id="${a}"]`).checked = true; m.querySelector(`input[data-id="${b}"]`).checked = true; }, [ids[0], ids[2]]);
  await w.click('.modal .cs-row:nth-child(3) .btn:first-of-type'); await w.click('.modal .cs-row:nth-child(2) .btn:first-of-type');
  await w.click('.modal .btn.primary');
  const cs = await I(w, () => window.__csP);
  assert.deepEqual(cs, { name: 'Corta', ids: [ids[2], ids[0]] });
  await I(w, (c) => { const P = window.__impresiona.pres; P.customShows = [c]; }, cs);
  await I(w, () => window.__impresiona.DS.presentCustom('Corta')); await wait(600);
  assert.deepEqual(await I(w, () => window.__show.engine.pres.slides.map((s) => s.id)), [ids[2], ids[0]]);
  await I(w, () => window.__show.end()); await wait(300);
  log('Presentación personalizada: elegir diapositivas, ordenarlas y presentarla');
  // 2) configuración: en bucle con pausa
  await startDialog(w, 'showSettings'); await w.waitForSelector('.modal');
  await setDialog(w, { range: 'custom', custom: 'Corta', loop: true, pause: 1, manual: true });
  assert.deepEqual(await I(w, () => window.__impresiona.pres.showSettings), { range: 'custom', from: 0, custom: 'Corta', loop: true, pause: 1, manual: true, clickBg: true, pen: false, laser: false });
  await I(w, () => window.__impresiona.onCommand('present')); await wait(600);
  await I(w, () => { const e = window.__show.engine; e.mount(1, { animate: false }); for (let k = 0; k < 20; k++) window.__show.next(); });
  assert.ok(await I(w, () => document.querySelector('.show').classList.contains('pausing')));
  await wait(2200);
  assert.deepEqual(await I(w, () => [document.querySelector('.show').classList.contains('pausing'), document.querySelector('.show').classList.contains('ended'), window.__show.engine.i]), [false, false, 0]);
  await I(w, () => window.__show.end()); await wait(300);
  log('Configuración de la presentación: presentación personalizada, en bucle con pausa y avance manual');
  // 3) se guardan en .odp y .pptx (propio y para LibreOffice)
  await I(w, () => { window.__shell.doc.name = 'pres2.pptx'; });
  for (const ext of ['odp', 'pptx']) await I(w, (e) => window.__shell.saveAsFormat(e, e), ext);
  for (const ext of ['odp', 'pptx']) {
    await reopen(w, path.join(out, `pres2.${ext}`)); await wait(500);
    const r = await I(w, () => { const P = window.__impresiona.pres; return [P.customShows?.[0]?.name, P.customShows?.[0]?.ids.length, P.showSettings?.loop, P.showSettings?.custom]; });
    assert.deepEqual(r, ['Corta', 2, true, 'Corta'], ext);
  }
  const pz = await JSZip.loadAsync(fs.readFileSync(path.join(out, 'pres2.pptx')));
  assert.match(await pz.file('ppt/presentation.xml').async('string'), /<p:custShow name="Corta" id="0"><p:sldLst><p:sld r:id="[^"]+"\/><p:sld r:id="[^"]+"\/><\/p:sldLst>/);
  assert.match(await pz.file('ppt/presProps.xml').async('string'), /<p:showPr loop="1"[^>]*><p:present\/><p:custShow id="0"\/>/);
  const oz = await JSZip.loadAsync(fs.readFileSync(path.join(out, 'pres2.odp')));
  assert.match(await oz.file('content.xml').async('string'), /<presentation:settings presentation:endless="true" presentation:pause="PT1S"[^>]*><presentation:show presentation:name="Corta" presentation:pages="Diapositiva3,Diapositiva1"\/>/);
  if (hasSoffice) {
    fs.mkdirSync(path.join(out, 'lo'), { recursive: true });
    execFileSync('soffice', ['--headless', '--convert-to', 'odp', '--outdir', path.join(out, 'lo'), path.join(out, 'pres2.pptx')], { stdio: 'ignore' });
    const lc = await (await JSZip.loadAsync(fs.readFileSync(path.join(out, 'lo', 'pres2.odp')))).file('content.xml').async('string');
    assert.match(lc, /presentation:show presentation:name="Corta"/); assert.match(lc, /presentation:endless="true"/);
    await reopen(w, path.join(out, 'lo', 'pres2.odp')); await wait(600);
    const r = await I(w, () => { const P = window.__impresiona.pres; return [P.customShows?.[0]?.name, P.customShows?.[0]?.ids.length, P.showSettings?.loop]; });
    assert.deepEqual(r, ['Corta', 2, true]);
    log('LibreOffice lee la presentación personalizada y el bucle del .pptx, e Impresiona los lee del .odp de LibreOffice');
  }
  // 4) esquema, resumen y expandir
  await w.evaluate(() => window.__shell.newDoc('negocio')); await wait(500);
  const titles = await I(w, () => window.__impresiona.pres.slides.map((s) => (s.objects.find((o) => o.role === 'title')?.html || '').replace(/<[^>]+>/g, '')));
  await startDialog(w, 'outline'); await w.waitForSelector('.modal .outline-ta');
  const txt = await w.inputValue('.modal .outline-ta');
  assert.ok(txt.startsWith(titles[0])); assert.match(txt, /\n\t\S/);
  await w.fill('.modal .outline-ta', `${txt}\nNueva idea\n\tPunto A\n\t\tDetalle A1\n\tPunto B`);
  await w.click('.modal .btn.primary'); await wait(400);
  const last = await I(w, () => { const s = window.__impresiona.pres.slides.at(-1); return [s.objects.find((o) => o.role === 'title').html, s.objects.find((o) => o.role === 'body').html]; });
  assert.deepEqual(last, ['Nueva idea', '<ul><li>Punto A<ul><li>Detalle A1</li></ul></li><li>Punto B</li></ul>']);
  const n1 = await I(w, () => window.__impresiona.pres.slides.length); assert.equal(n1, titles.length + 1);
  await I(w, () => window.__impresiona.goto(window.__impresiona.pres.slides.length - 1));
  await I(w, () => window.__impresiona.onCommand('expandSlide')); await wait(300);
  const ex = await I(w, () => window.__impresiona.pres.slides.slice(-2).map((s) => [s.objects.find((o) => o.role === 'title').html, s.objects.find((o) => o.role === 'body')?.html || '']));
  assert.deepEqual(ex, [['Punto A', '<ul><li>Detalle A1</li></ul>'], ['Punto B', '']]);
  await I(w, () => window.__impresiona.goto(0)); await I(w, () => window.__impresiona.onCommand('summarySlide')); await wait(300);
  const sum = await I(w, () => { const s = window.__impresiona.pres.slides[1]; return [s.objects.find((o) => o.role === 'title').html, (s.objects.find((o) => o.role === 'body').html.match(/<li>/g) || []).length]; });
  assert.deepEqual(sum, [EN ? 'Summary' : 'Resumen', n1 + 2]);
  log('Vista Esquema (títulos y puntos con Tab), Expandir diapositiva y Diapositiva de resumen');
  // 5) insertar diapositivas desde un archivo
  const before = await I(w, () => window.__impresiona.pres.slides.length);
  const added = await I(w, async (p) => { const f = await window.oficina.readPath(p); return window.__impresiona.DS.insertFromFile(f); }, path.join(out, 'pres2.pptx'));
  assert.equal(added, N); assert.equal(await I(w, () => window.__impresiona.pres.slides.length), before + N);
  assert.equal(new Set(await I(w, () => window.__impresiona.pres.slides.map((s) => s.id))).size, before + N);
  log('Insertar diapositivas desde otro archivo (.pptx), con identificadores nuevos');
  // 6) duplicar con desplazamiento, giro y colores
  await I(w, () => { const P = window.__impresiona; P.goto(0); P.pres.slides[0].objects.push({ id: 'sq', type: 'shape', shape: 'rect', x: 50, y: 50, w: 80, h: 60, rot: 0, html: '', style: { fill: '#ff0000' } }); P.renderAll(); P.select(['sq']); });
  await startDialog(w, 'duplicateObj'); await w.waitForSelector('.modal');
  await setDialog(w, { n: 3, dx: 30, dy: 10, rot: 15, dw: 10, colors: true, c1: '#ff0000', c2: '#0000ff' });
  const dup = await I(w, () => window.__impresiona.pres.slides[0].objects.filter((o) => o.shape === 'rect').map((o) => [o.x, o.y, o.w, o.rot, o.style.fill]));
  assert.deepEqual(dup, [[50, 50, 80, 0, '#ff0000'], [80, 60, 90, 15, '#aa0055'], [110, 70, 100, 30, '#5500aa'], [140, 80, 110, 45, '#0000ff']]);
  log('Duplicar (Mayús+F3): copias con desplazamiento, giro, tamaño y colores intermedios');
  // 7) campos
  await I(w, () => { const P = window.__impresiona; P.pres.slides[2].objects.push({ id: 'fd', type: 'text', x: 600, y: 480, w: 300, h: 40, rot: 0, html: '<p>Diapositiva <span class="fld" data-fld="num">#</span> de <span class="fld" data-fld="count">#</span></p>', style: { size: 14 } }); P.goto(2); P.renderAll(); });
  const total = await I(w, () => window.__impresiona.pres.slides.length);
  assert.equal(await I(w, () => document.querySelector('.slide.m-edit [data-id="fd"]').textContent.trim()), `Diapositiva 3 de ${total}`);
  await I(w, () => { const P = window.__impresiona; const s = P.pres.slides.splice(2, 1)[0]; P.pres.slides.push(s); P.goto(P.pres.slides.length - 1); P.renderAll(); });
  assert.equal(await I(w, () => document.querySelector('.slide.m-edit [data-id="fd"]').textContent.trim()), `Diapositiva ${total} de ${total}`);
  log('Campos en el texto: número de diapositiva y total, se actualizan solos');
  // 8) imprimir notas y folletos (y en PDF)
  await I(w, () => { const P = window.__impresiona.pres; P.slides.forEach((s, i) => { s.notes = `Nota ${i + 1}`; }); P.printOpts = { content: 'handouts', per: 6 }; window.__impresiona.buildPrint(); });
  assert.equal(await I(w, () => document.querySelectorAll('#printRoot .ppage').length), Math.ceil(total / 6));
  assert.equal(await I(w, () => document.querySelectorAll('#printRoot .ppage:first-child .pmini').length), 6);
  await I(w, () => { window.__impresiona.pres.printOpts = { content: 'notes' }; window.__impresiona.buildPrint(); });
  assert.equal(await I(w, () => document.querySelectorAll('#printRoot .ppage').length), total);
  assert.match(await I(w, () => document.querySelector('#printRoot .ppage:nth-child(2) .pnotes-text').textContent), /Nota 2/);
  await I(w, () => { window.__shell.doc.name = 'notas.pptx'; window.__impresiona.pres.printOpts = { content: 'handouts', per: 4 }; });
  await I(w, () => window.__shell.exportPdf()); await wait(1500);
  assert.equal(py(`import pypdf\nr=pypdf.PdfReader('${path.join(out, 'notas.pdf')}')\nb=r.pages[0].mediabox\nprint(len(r.pages), round(float(b.width)), round(float(b.height)))`), `${Math.ceil(total / 4)} 612 792`);
  await startDialog(w, 'print'); await w.waitForSelector('.modal [name="app_content"]', { timeout: 15000 }); await wait(600);
  await setDialog(w, { app_content: 'slides' }, false); await wait(1500);
  assert.match(await I(w, () => document.querySelector('.modal .pv-opts .fx-note').textContent), new RegExp(`${total} (páginas|pages)`));
  await w.keyboard.press('Escape'); await wait(300);
  log('Imprimir y exportar a PDF: notas, documento (folletos de 4 y 6 por página, tamaño carta) y diapositivas');
  await closeApp('impresiona'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de Impresiona, parte 2, pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
