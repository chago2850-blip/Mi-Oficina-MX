// Pruebas de Impresiona, parte 4: curvas, polígonos, formas libres, edición de puntos y conectores.

import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase12');
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
// de coordenadas de la diapositiva a la pantalla
const scr = (w, x, y) => I(w, ([a, b]) => { const r = document.querySelector('.stage').getBoundingClientRect(); const sl = document.querySelector('.slide.m-edit'); const k = sl.getBoundingClientRect().width / window.__impresiona.pres.size.w; return [r.left + a * k, r.top + b * k]; }, [x, y]);
const last = (w) => I(w, () => { const s = window.__impresiona.pres.slides[window.__impresiona.cur]; return JSON.parse(JSON.stringify(s.objects.at(-1))); });
try {
  const w = await openApp('impresiona');
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(500);
  await I(w, () => { const P = window.__impresiona; P.pres.slides[0].objects = []; P.renderAll(); });
  // 1) polígono con clics y doble clic
  await I(w, () => window.__impresiona.onCommand('drawPolygon'));
  for (const [x, y] of [[100, 100], [300, 120], [200, 300]]) { const [sx, sy] = await scr(w, x, y); await w.mouse.click(sx, sy); await wait(60); }
  { const [sx, sy] = await scr(w, 200, 300); await w.mouse.dblclick(sx, sy); } await wait(300);
  let o = await last(w);
  assert.equal(o.shape, 'path'); assert.equal(o.closed, true); assert.ok(o.pts.length >= 3);
  assert.ok(Math.abs(o.x - 100) <= 2 && Math.abs(o.y - 100) <= 2 && Math.abs(o.w - 200) <= 3 && Math.abs(o.h - 200) <= 3, JSON.stringify(o));
  assert.ok(await I(w, (id) => !!document.querySelector(`.slide.m-edit [data-id="${id}"] svg path[d^="M"]`), o.id));
  // 2) curva (clics + Intro) y polilínea
  await I(w, () => window.__impresiona.onCommand('drawCurve'));
  for (const [x, y] of [[400, 100], [500, 200], [600, 100], [700, 200]]) { const [sx, sy] = await scr(w, x, y); await w.mouse.click(sx, sy); await wait(60); }
  await w.keyboard.press('Enter'); await wait(300);
  o = await last(w); assert.deepEqual([o.closed, o.smooth, o.pts.length], [false, true, 4]);
  assert.match(await I(w, (id) => document.querySelector(`.slide.m-edit [data-id="${id}"] svg path`).getAttribute('d'), o.id), /C/);
  // 3) forma libre (arrastrar)
  await I(w, () => window.__impresiona.onCommand('drawFreehand'));
  { const pts = Array.from({ length: 40 }, (_, i) => [150 + 80 * Math.cos(i / 6), 420 + 50 * Math.sin(i / 6)]); const [x0, y0] = await scr(w, ...pts[0]); await w.mouse.move(x0, y0); await w.mouse.down(); for (const p of pts.slice(1)) { const [sx, sy] = await scr(w, ...p); await w.mouse.move(sx, sy); } await w.mouse.up(); }
  await wait(300); o = await last(w); assert.equal(o.smooth, true); assert.ok(o.pts.length >= 6 && o.pts.length < 40, String(o.pts.length));
  log('Polígono, curva y forma libre dibujados con el ratón');
  // 4) conector entre dos formas que las sigue al moverlas
  await I(w, () => { const P = window.__impresiona; const s = P.pres.slides[0]; s.objects.push({ id: 'A', type: 'shape', shape: 'rect', x: 380, y: 300, w: 120, h: 80, rot: 0, html: '', style: {} }, { id: 'B', type: 'shape', shape: 'ellipse', x: 700, y: 400, w: 120, h: 80, rot: 0, html: '', style: {} }); P.renderAll(); P.onCommand('connector', 'elbow'); });
  { const [ax, ay] = await scr(w, 495, 340); const [bx, by] = await scr(w, 705, 440); await w.mouse.move(ax, ay); await w.mouse.down(); await w.mouse.move(bx - 30, by); await w.mouse.move(bx, by); await w.mouse.up(); }
  await wait(300); o = await last(w);
  assert.deepEqual(o.conn, { a: { id: 'A', side: 'right' }, b: { id: 'B', side: 'left' }, kind: 'elbow' });
  const ends = (q) => [[q.x + q.pts[0][0] * q.w, q.y + q.pts[0][1] * q.h], [q.x + q.pts.at(-1)[0] * q.w, q.y + q.pts.at(-1)[1] * q.h]].map((p) => p.map(Math.round));
  assert.deepEqual(ends(o), [[500, 340], [700, 440]]);
  await I(w, () => { const P = window.__impresiona; const b = P.pres.slides[0].objects.find((x) => x.id === 'B'); b.y = 100; P.renderAll(); });
  o = await last(w); assert.deepEqual(ends(o), [[500, 340], [700, 140]]);
  await I(w, (id) => { window.__impresiona.select([id]); window.__impresiona.onCommand('connCurved'); }, o.id);
  o = await last(w); assert.equal(o.conn.kind, 'curved'); assert.equal(o.bez, true);
  log('Conectores (en ángulo y curvo) que siguen a las formas al moverlas');
  // 5) editar puntos: mover un punto con el ratón
  const poly = await I(w, () => window.__impresiona.pres.slides[0].objects[0].id);
  await I(w, (id) => { window.__impresiona.select([id]); window.__impresiona.onCommand('editPoints'); }, poly);
  assert.equal(await I(w, () => document.querySelectorAll('.pt-handle').length), 3);
  { const box = await w.locator('.pt-handle').first().boundingBox(); const [tx, ty] = await scr(w, 60, 80); await w.mouse.move(box.x + 5, box.y + 5); await w.mouse.down(); await w.mouse.move(tx, ty, { steps: 4 }); await w.mouse.up(); }
  await wait(300);
  o = await I(w, (id) => window.__impresiona.pres.slides[0].objects.find((x) => x.id === id), poly);
  assert.ok(Math.abs(o.x - 60) <= 3 && Math.abs(o.y - 80) <= 3, `${o.x},${o.y}`);
  await I(w, () => window.__impresiona.onCommand('editPoints')); assert.equal(await I(w, () => document.querySelectorAll('.pt-handle').length), 0);
  log('Editar puntos (F8): arrastrar un punto cambia la forma');
  // 6) .pptx (custGeom) y .odp (draw:path y draw:connector), también con LibreOffice
  await I(w, () => { window.__shell.doc.name = 'trazos.pptx'; });
  for (const e of ['pptx', 'odp']) await I(w, (x) => window.__shell.saveAsFormat(x, x), e);
  const sx1 = await (await JSZip.loadAsync(fs.readFileSync(path.join(out, 'trazos.pptx')))).file('ppt/slides/slide1.xml').async('string');
  assert.ok((sx1.match(/<a:custGeom>/g) || []).length >= 4); assert.match(sx1, /<a:cubicBezTo>/); assert.match(sx1, /<a:close \/>/);
  const ox = await (await JSZip.loadAsync(fs.readFileSync(path.join(out, 'trazos.odp')))).file('content.xml').async('string');
  assert.equal((ox.match(/<draw:path /g) || []).length, 3); assert.match(ox, /<draw:connector [^>]*draw:type="curve" draw:start-shape="shA" draw:start-glue-point="1" draw:end-shape="shB" draw:end-glue-point="3"/); assert.match(ox, /draw:id="shA"/);
  for (const e of ['odp', 'pptx']) {
    await reopen(w, path.join(out, `trazos.${e}`)); await wait(500);
    const r = await I(w, () => window.__impresiona.pres.slides[0].objects.filter((x) => x.shape === 'path').map((x) => [x.closed, !!x.conn]));
    assert.deepEqual(r, [[true, false], [false, false], [r[2][0], false], [false, true]], e);
  }
  if (hasSoffice) {
    fs.mkdirSync(path.join(out, 'lo'), { recursive: true });
    execFileSync('soffice', ['--headless', '--convert-to', 'odp', '--outdir', path.join(out, 'lo'), path.join(out, 'trazos.odp')], { stdio: 'ignore' });
    const lc = await (await JSZip.loadAsync(fs.readFileSync(path.join(out, 'lo', 'trazos.odp')))).file('content.xml').async('string');
    assert.match(lc, /<draw:connector[^>]*draw:start-shape="[^"]+"[^>]*draw:end-shape="[^"]+"/);
    await reopen(w, path.join(out, 'lo', 'trazos.odp')); await wait(600);
    const r = await I(w, () => { const s = window.__impresiona.pres.slides[0]; const c = s.objects.find((x) => x.conn); return [s.objects.filter((x) => x.shape === 'path' && !x.conn).length, c ? [s.objects.find((x) => x.id === c.conn.a.id)?.shape, s.objects.find((x) => x.id === c.conn.b.id)?.shape, c.conn.a.side, c.conn.b.side] : null]; });
    assert.deepEqual(r, [3, ['rect', 'ellipse', 'right', 'left']]);
    execFileSync('soffice', ['--headless', '--convert-to', 'pptx', '--outdir', path.join(out, 'lo'), path.join(out, 'trazos.odp')], { stdio: 'ignore' });
    await reopen(w, path.join(out, 'lo', 'trazos.pptx')); await wait(600);
    assert.ok(await I(w, () => window.__impresiona.pres.slides[0].objects.filter((x) => x.shape === 'path').length) >= 3);
    log('LibreOffice conserva los trazos y los conectores unidos a sus formas, e Impresiona los lee de vuelta');
  }
  log('Trazos en .pptx (custGeom) y .odp (draw:path y draw:connector)');
  await closeApp('impresiona'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de Impresiona, parte 4, pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
