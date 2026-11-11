// Pruebas de la fase 7 (Impresiona como LibreOffice Impress, parte 1): Posición y tamaño, línea, área, texto, formas,
// voltear, distribuir, agrupar, recortar e imagen, acciones al hacer clic, ocultar diapositiva, cuadrícula; .pptx y .odp.

import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase7');
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
const I = (w, fn, arg) => w.evaluate(fn, arg);
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAAASFvFNAAAAFklEQVR4nGP4z8DAwMDAxMDAwMDAAAANHQEDK+mmyAAAAABJRU5ErkJggg==';

try {
  const w = await openApp('impresiona');
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(400);
  // objetos de prueba: hexágono, flecha (línea), imagen, cuadro de texto
  await I(w, (png) => {
    const P = window.__impresiona; const s = P.pres.slides[P.cur];
    s.objects.push({ id: 'hx', type: 'shape', shape: 'hexagon', x: 100, y: 100, w: 200, h: 150, rot: 0, html: '', style: { valign: 'middle', align: 'center' } });
    s.objects.push({ id: 'ln', type: 'shape', shape: 'lineArrow', x: 400, y: 100, w: 240, h: 20, rot: 0, html: '', style: {} });
    s.objects.push({ id: 'im', type: 'image', src: png, x: 100, y: 300, w: 200, h: 100, rot: 0 });
    s.objects.push({ id: 'tx', type: 'text', x: 400, y: 300, w: 300, h: 120, rot: 0, html: '<p>Primer párrafo</p><p>Segundo párrafo</p>', style: { size: 20 } });
    P.renderAll();
  }, PNG);
  // 1) Posición y tamaño (F4)
  await I(w, () => window.__impresiona.select(['hx']));
  await startDialog(w, 'posSize'); await w.waitForSelector('.modal');
  await setDialog(w, { x: 5, y: 4, w: 6, hh: 3 }, false);
  await w.evaluate(() => document.querySelectorAll('.modal .fx-tab')[1].click());
  await setDialog(w, { rot: 30, flipH: true });
  let o = await I(w, () => window.__impresiona.pres.slides[0].objects.find((x) => x.id === 'hx'));
  assert.deepEqual([o.x, o.y, o.w, o.h, o.rot, o.flipH], [189, 151, 227, 113, 30, true]);
  assert.match(await I(w, () => document.querySelector('.slide.m-edit [data-id="hx"]').style.transform), /rotate\(30deg\) scaleX\(-1\)/);
  // 2) Área: degradado, transparencia y sombra
  await startDialog(w, 'areaDlg'); await w.waitForSelector('.modal');
  await setDialog(w, { kind: 'grad', c1: '#ff0000', c2: '#0000ff', angle: 0 }, false);
  await w.evaluate(() => document.querySelectorAll('.modal .fx-tab')[1].click()); await setDialog(w, { fop: 20 }, false);
  await w.evaluate(() => document.querySelectorAll('.modal .fx-tab')[2].click()); await setDialog(w, { sh: true, shd: 6 });
  o = await I(w, () => window.__impresiona.pres.slides[0].objects.find((x) => x.id === 'hx'));
  assert.deepEqual(o.style.grad, { c2: '#0000ff', type: 'linear', angle: 0 }); assert.equal(o.style.fill, '#ff0000'); assert.equal(o.style.fillOpacity, 0.8); assert.equal(o.style.shadow.dx, 6);
  assert.ok(await I(w, () => !!document.querySelector('.slide.m-edit [data-id="hx"] linearGradient')));
  assert.match(await I(w, () => document.querySelector('.slide.m-edit [data-id="hx"]').style.filter), /drop-shadow/);
  // 3) Línea: guiones, grosor y puntas
  await I(w, () => window.__impresiona.select(['ln']));
  await startDialog(w, 'lineDlg'); await setDialog(w, { dash: 'dash', sw: 3, a1: 'circle', a2: 'triangle', stroke: '#008000' });
  o = await I(w, () => window.__impresiona.pres.slides[0].objects.find((x) => x.id === 'ln'));
  assert.deepEqual([o.style.dash, o.style.arrowStart, o.style.arrowEnd, o.style.stroke, o.style.strokeW], ['dash', 'circle', 'triangle', '#008000', 4]);
  assert.equal(await I(w, () => document.querySelectorAll('.slide.m-edit [data-id="ln"] marker').length), 2);
  // 4) Texto: márgenes, interlineado, espacio entre párrafos
  await I(w, () => window.__impresiona.select(['tx']));
  await startDialog(w, 'textDlg'); await w.waitForSelector('.modal'); await setDialog(w, { pl: 0.5, valign: 'middle' }, false);
  await w.evaluate(() => document.querySelectorAll('.modal .fx-tab')[1].click()); await setDialog(w, { lineH: 1.5, paraSp: 12 });
  const tx = await I(w, () => { const el = document.querySelector('.slide.m-edit [data-id="tx"] .otxt'); return { lh: el.style.lineHeight, pl: el.style.paddingLeft, mb: getComputedStyle(el.querySelector('p')).marginBottom }; });
  assert.deepEqual(tx, { lh: '1.5', pl: '19px', mb: '16px' });
  log('Posición y tamaño (F4), Área (degradado, transparencia, sombra), Línea (guiones, puntas) y Texto (márgenes, interlineado)');
  // 5) imagen: recortar y ajustes
  await I(w, () => window.__impresiona.select(['im']));
  await startDialog(w, 'crop'); await fillForm(w, { l: '25', r: '25' });
  o = await I(w, () => window.__impresiona.pres.slides[0].objects.find((x) => x.id === 'im'));
  assert.deepEqual([o.crop, o.x, o.w], [{ l: 25, r: 25, t: 0, b: 0 }, 150, 100]);
  await startDialog(w, 'imgAdjust'); await fillForm(w, { mode: 'gray', bright: '120' });
  assert.match(await I(w, () => document.querySelector('.slide.m-edit [data-id="im"]').style.filter), /grayscale\(1\) brightness\(1\.2\)/);
  log('Imagen: recortar conservando el tamaño y ajustes (escala de grises, brillo)');
  // 6) organizar: agrupar, distribuir, un nivel
  await I(w, () => window.__impresiona.select(['hx', 'tx']));
  await w.evaluate(() => window.__impresiona.onCommand('group'));
  await I(w, () => window.__impresiona.select([]));
  const box = await I(w, () => { const r = document.querySelector('.slide.m-edit [data-id="tx"]').getBoundingClientRect(); return { x: r.x + 20, y: r.y + 20 }; });
  await w.mouse.click(box.x, box.y); await wait(100);
  assert.deepEqual((await I(w, () => window.__impresiona.selected())).sort(), ['hx', 'tx']);
  await I(w, () => window.__impresiona.select(['hx', 'ln', 'im']));
  await I(w, () => window.__impresiona.OB.distribute('h'));
  const xs = await I(w, () => window.__impresiona.pres.slides[0].objects.filter((x) => ['hx', 'ln', 'im'].includes(x.id)).map((x) => [x.id, x.x, x.w]));
  const sorted = [...xs].sort((a, b) => a[1] - b[1]); const g1 = sorted[1][1] - (sorted[0][1] + sorted[0][2]), g2 = sorted[2][1] - (sorted[1][1] + sorted[1][2]);
  assert.ok(Math.abs(g1 - g2) <= 1, JSON.stringify(xs));
  await I(w, () => window.__impresiona.select(['hx'])); await I(w, () => window.__impresiona.OB.zStep(1));
  assert.deepEqual(await I(w, () => window.__impresiona.pres.slides[0].objects.map((x) => x.id).filter((id) => ['ln', 'hx', 'im', 'tx'].includes(id))), ['ln', 'hx', 'im', 'tx']);
  log('Agrupar (un clic selecciona el grupo), distribuir y traer adelante un nivel');
  // 7) más formas, acción al hacer clic, ocultar diapositiva
  await I(w, () => { const P = window.__impresiona; ['diamond', 'heart', 'callout', 'cloud', 'chevron', 'upArrow'].forEach((k, i) => P.pres.slides[0].objects.push({ id: `s${i}`, type: 'shape', shape: k, x: 20 + i * 90, y: 450, w: 80, h: 60, rot: 0, html: '', style: {} })); P.renderAll(); });
  assert.equal(await I(w, () => [...document.querySelectorAll('.slide.m-edit [data-id^="s"] path')].length), 6);
  await w.evaluate(() => window.__impresiona.onCommand('newSlide')); await w.evaluate(() => window.__impresiona.onCommand('newSlide')); await wait(200);
  await w.evaluate(() => window.__impresiona.goto(0)); await I(w, () => window.__impresiona.select(['tx']));
  await startDialog(w, 'objLink'); await fillForm(w, { kind: 'slide', slide: await I(w, () => window.__impresiona.pres.slides[2].id) });
  assert.equal(await I(w, () => window.__impresiona.pres.slides[0].objects.find((x) => x.id === 'tx').link), `#slide:${await I(w, () => window.__impresiona.pres.slides[2].id)}`);
  await w.evaluate(() => { window.__impresiona.goto(1); window.__impresiona.onCommand('hideSlide'); }); await wait(100);
  assert.ok(await I(w, () => window.__impresiona.pres.slides[1].hidden));
  assert.ok(await I(w, () => document.querySelectorAll('.thumb')[1].classList.contains('hidden-slide')));
  // presentación: la diapositiva oculta no se cuenta y el vínculo salta a la 3
  await w.evaluate(() => window.__impresiona.present(0)); await wait(600);
  assert.equal(await w.evaluate(() => document.querySelector('.show .cnt').textContent), '1 / 2');
  await w.evaluate(() => document.querySelector('.show .obj[data-link]').click()); await wait(500);
  assert.equal(await w.evaluate(() => document.querySelector('.show .cnt').textContent), '2 / 2');
  await w.keyboard.press('Escape'); await wait(400);
  log('Formas nuevas, acción al hacer clic (ir a diapositiva) y diapositiva oculta en la presentación');
  // 8) cuadrícula con ajuste
  await w.evaluate(() => window.__impresiona.goto(0));
  await w.evaluate(() => window.__impresiona.onCommand('grid')); await w.evaluate(() => window.__impresiona.onCommand('snapGrid'));
  assert.ok(await I(w, () => document.querySelector('.slide.m-edit').classList.contains('show-grid')));
  await I(w, () => window.__impresiona.select(['ln']));
  const lb = await I(w, () => { const r = document.querySelector('.slide.m-edit [data-id="ln"]').getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2 }; });
  await w.mouse.move(lb.x, lb.y); await w.mouse.down(); await w.mouse.move(lb.x + 37, lb.y + 23, { steps: 5 }); await w.mouse.up();
  const ln = await I(w, () => window.__impresiona.pres.slides[0].objects.find((x) => x.id === 'ln'));
  assert.ok(ln.x % 20 === 0 && ln.y % 20 === 0, `${ln.x},${ln.y}`);
  await w.evaluate(() => { window.__impresiona.onCommand('grid'); window.__impresiona.onCommand('snapGrid'); });
  log('Cuadrícula visible y ajuste a la cuadrícula al mover');
  // 9) guardar en .pptx y .odp y volver a abrir
  const snap = () => I(w, () => { const s = window.__impresiona.pres.slides; const f = (id) => s[0].objects.find((x) => x.id === id || (x.type === 'image' && id === 'im') || (x.shape === 'hexagon' && id === 'hx') || (x.shape?.startsWith('line') && id === 'ln') || (x.type === 'text' && id === 'tx')); const hx = f('hx'), ln2 = f('ln'), im = f('im'), tx2 = f('tx'); return { n: s.length, hidden: !!s[1].hidden, hx: hx && { shape: hx.shape, rot: hx.rot, flipH: !!hx.flipH, fill: hx.style?.fill, grad: hx.style?.grad?.c2, op: hx.style?.fillOpacity, shadow: !!hx.style?.shadow }, ln: ln2 && { dash: ln2.style?.dash, a1: ln2.style?.arrowStart, a2: ln2.style?.arrowEnd }, im: im && { crop: im.crop && Math.round(im.crop.l), gray: !!im.img?.gray }, link: tx2?.link === `#slide:${s[2].id}`, lineH: tx2?.style?.lineH, shapes: s[0].objects.filter((x) => /^(diamond|heart|callout|cloud|chevron|upArrow)$/.test(x.shape)).length }; });
  const expect = await snap();
  for (const [ext, lab] of [['pptx', 'PowerPoint'], ['odp', 'ODP']]) {
    await w.evaluate(async ([e, l]) => window.__shell.saveAsFormat(e, l), [ext, lab]); await wait(1200);
    const file = path.join(out, `${UNTITLED}.${ext}`);
    await reopen(w, file); await wait(1000);
    assert.deepEqual(await snap(), expect, ext);
    if (hasSoffice) {
      const lo = path.join(out, `lo-${ext}`); fs.mkdirSync(lo, { recursive: true }); const to = ext === 'pptx' ? 'odp' : 'pptx';
      execFileSync('soffice', ['--headless', '--convert-to', to, '--outdir', lo, file], { stdio: 'ignore', timeout: 120000 });
      await reopen(w, path.join(lo, `${UNTITLED}.${to}`)); await wait(1200);
      const r = await snap();
      assert.equal(r.n, 3, `LO ${ext}→${to}`); assert.ok(r.hidden, `LO ${ext}→${to} oculta`);
      assert.equal(r.hx?.shape, 'hexagon', `LO ${ext}→${to} ${JSON.stringify(r)}`); assert.ok(r.hx.flipH, `LO ${ext}→${to} volteo ${JSON.stringify(r.hx)}`);
      assert.equal(r.ln?.dash, 'dash', `LO ${ext}→${to} ${JSON.stringify(r.ln)}`); assert.equal(r.ln?.a2, 'triangle', `LO ${ext}→${to} ${JSON.stringify(r.ln)}`);
      assert.ok(r.shapes >= (ext === 'pptx' ? 5 : 3), `LO ${ext}→${to} formas ${r.shapes}`);
      console.log(`  LibreOffice ${ext}→${to}: ${JSON.stringify(r)}`);
      await reopen(w, file); await wait(800);
    }
  }
  log(`Todo se guarda en .pptx y .odp${hasSoffice ? ' y lo principal sobrevive a LibreOffice' : ''}`);
  await closeApp('impresiona'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de la fase 7 (Impresiona, parte 1) pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
