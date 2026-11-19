// Pruebas de Escribiente 0.17.1: opciones de pegado (Ctrl) — mantener el formato original, combinar
// formato y mantener solo texto, con el ratón, con el teclado y desde la cinta.

import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase15');
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
const app = await electron.launch({ ...(exe ? { executablePath: exe, args: [`--user-data-dir=${path.join(out, 'perfil')}`, '--no-sandbox', ...FAKE] } : { args: [root, `--user-data-dir=${path.join(out, 'perfil')}`, '--no-sandbox', ...FAKE] }), cwd: root, env: { ...process.env, OFICINA_TEST_DIR: out, OFICINA_LANG: process.env.OFICINA_LANG || 'es', OFICINA_FAKE_TRANSLATE: '1' } });
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

// pone valores en el cuadro abierto (también en pestañas ocultas) y acepta
async function setDialog(w, values, submit = true) {
  await w.waitForSelector('.modal', { timeout: 5000 }); await wait(200); // el cuadro pone el foco en su primer campo justo después de abrirse
  await w.evaluate((vals) => {
    const m = [...document.querySelectorAll('.modal')].pop();
    for (const [k, v] of Object.entries(vals)) { const el = m.querySelector(`[name="${k}"]`); if (!el) throw new Error(`sin campo ${k}`); if (el.type === 'checkbox') el.checked = !!v; else el.value = String(v); el.dispatchEvent(new Event('change', { bubbles: true })); el.dispatchEvent(new Event('input', { bubbles: true })); }
  }, values);
  if (submit) { await w.click('.modal .btn.primary'); await wait(300); }
}
const E = (w, fn, a) => w.evaluate(fn, a);
const CM = 96 / 2.54;
const box = (w, s) => E(w, (q) => { const r = document.querySelector(q)?.getBoundingClientRect(); return r ? { x: r.left + r.width / 2, y: r.top + r.height / 2, l: r.left, r: r.right } : null; }, s);
async function drag(w, sel, dx) {
  const b = await box(w, sel); assert.ok(b, `no existe ${sel}`);
  await w.mouse.move(b.x, b.y); await w.mouse.down(); await w.mouse.move(b.x + dx / 2, b.y, { steps: 4 }); await w.mouse.move(b.x + dx, b.y, { steps: 4 }); await w.mouse.up(); await wait(400);
}
const caretIn = (w) => E(w, () => { const s = getSelection(); const n = s.anchorNode; const el = n?.nodeType === 3 ? n.parentElement : n; return { inEditor: document.getElementById('editor').contains(n), tag: el?.closest('p,h1,h2,h3,li,td')?.tagName || el?.tagName, idx: [...document.getElementById('editor').children].indexOf(el?.closest('#editor > *')) }; });

const CLIP_HTML = '<p style="font-family:Courier New;font-size:20pt;color:#c00000">Hola <span style="font-weight:bold">mundo</span> <i>bonito</i></p><p style="text-align:center;font-family:Georgia">Segunda línea</p>';
const CLIP_TEXT = 'Hola mundo bonito\nSegunda línea';
const setClip = () => app.evaluate(({ clipboard }, [h, t]) => clipboard.write({ html: h, text: t }), [CLIP_HTML, CLIP_TEXT]);
const ed = (w) => E(w, () => document.getElementById('editor').innerHTML);
try {
  const w = await openApp('escribiente');
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  const fresh = async () => { await E(w, () => window.__escribiente.setContent('<p id="z">Inicio: </p>')); await wait(300); await E(w, () => { const p = document.getElementById('z'); const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); document.getElementById('editor').focus(); }); };
  await setClip(); await fresh();
  await w.keyboard.press('Control+V'); await wait(400);
  let h = await ed(w);
  assert.match(h, /Courier New/); assert.match(h, /mundo/);
  assert.ok(await w.isVisible('.paste-ctl'), 'aparece el botón (Ctrl)');
  // el botón queda junto al final de lo pegado
  const pos = await E(w, () => { const b = document.querySelector('.paste-ctl').getBoundingClientRect(); const ps = [...document.querySelectorAll('#editor p')]; const last = ps[ps.length - 1].getBoundingClientRect(); return { dy: b.top - last.bottom, dx: b.left - last.left }; });
  assert.ok(pos.dy > -30 && pos.dy < 30, `posición ${JSON.stringify(pos)}`);
  // ratón: abrir el menú y combinar formato
  await w.click('.paste-ctl .pc-btn'); await wait(150);
  await w.screenshot({ path: path.join(out, 'pegar-menu.png') });
  assert.equal(await E(w, () => document.querySelectorAll('.paste-ctl .pc-opt').length), 3);
  assert.equal(await E(w, () => document.querySelector('.paste-ctl .pc-opt.on').dataset.mode), 'keep');
  await w.click('.paste-ctl .pc-opt[data-mode="merge"]'); await wait(400);
  h = await ed(w);
  assert.doesNotMatch(h, /Courier|Georgia|20pt|c00000|rgb\(192/i, h);
  assert.match(h, /<b[^>]*>mundo<\/b>/); assert.match(h, /<i[^>]*>bonito<\/i>/); assert.match(h, /text-align: center/);
  assert.match(h, /Inicio:/);
  assert.equal(await E(w, () => window.__escribiente.PG2.lastMode()), 'merge');
  // solo texto
  await w.click('.paste-ctl .pc-btn'); await wait(150);
  await w.click('.paste-ctl .pc-opt[data-mode="text"]'); await wait(400);
  h = await ed(w);
  assert.doesNotMatch(h, /<b[ >]|<i[ >]|Courier/, h);
  assert.match(await E(w, () => document.getElementById('editor').innerText), /Inicio: ?Hola mundo bonito\s*\n+Segunda línea/);
  // de vuelta al formato original
  await w.click('.paste-ctl .pc-btn'); await wait(150);
  await w.click('.paste-ctl .pc-opt[data-mode="keep"]'); await wait(400);
  h = await ed(w);
  assert.match(h, /Courier New/);
  assert.equal((h.match(/Segunda línea/g) || []).length, 1, 'no se duplica');
  // deshacer quita el pegado completo
  await E(w, () => document.getElementById('editor').focus());
  await w.keyboard.press('Control+z'); await wait(300);
  assert.doesNotMatch(await ed(w), /mundo/);
  log('Opciones de pegado con el ratón: original, combinar formato y solo texto (con deshacer)');

  // teclado: Ctrl solo abre el menú y T / M / K eligen
  await fresh(); await w.keyboard.press('Control+V'); await wait(400);
  await w.keyboard.press('Control'); await wait(150);
  assert.ok(await E(w, () => window.__escribiente.PG2.menuOpen()), 'Ctrl abre el menú');
  await w.keyboard.press('t'); await wait(400);
  assert.doesNotMatch(await ed(w), /Courier/);
  await w.keyboard.press('Control'); await wait(150);
  await w.keyboard.press('m'); await wait(400);
  assert.match(await ed(w), /<b[^>]*>mundo<\/b>/);
  // Esc oculta el control y escribir también
  await w.keyboard.press('Escape'); await wait(150);
  assert.equal(await w.isVisible('.paste-ctl'), false);
  await fresh(); await w.keyboard.press('Control+V'); await wait(400);
  await w.keyboard.type('x'); await wait(200);
  assert.equal(await w.isVisible('.paste-ctl'), false, 'escribir lo oculta');
  log('Opciones de pegado con el teclado: Ctrl abre el menú; K, M y T eligen; Esc o escribir lo cierran');

  // Ctrl+Mayús+V pega solo el texto
  await fresh(); await w.keyboard.press('Control+Shift+V'); await wait(400);
  h = await ed(w);
  assert.doesNotMatch(h, /Courier|<b[ >]/); assert.match(h, /Hola mundo bonito/);
  // desde la cinta: flecha de Pegar → Combinar formato
  await fresh();
  await E(w, () => window.__escribiente.saveSelection?.());
  await w.evaluate(() => window.__shell.command('paste', 'merge')); await wait(500);
  h = await ed(w);
  assert.match(h, /<b[^>]*>mundo<\/b>/); assert.doesNotMatch(h, /Courier/);
  assert.ok(await w.isVisible('.paste-ctl'));
  // un texto sin formato se pega normalmente, sin control
  await app.evaluate(({ clipboard }) => clipboard.writeText('solo texto'));
  await fresh(); await w.keyboard.press('Control+V'); await wait(300);
  assert.match(await ed(w), /Inicio: ?solo texto/);
  log('Ctrl+Mayús+V pega solo el texto; la flecha de Pegar en la cinta ofrece las tres formas');

  await closeApp('escribiente'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de opciones de pegado pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
