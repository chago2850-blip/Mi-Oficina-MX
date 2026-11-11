// Pruebas de la fase 5, parte 6 (Escribiente como LibreOffice Writer): autocorrección, buscar con expresiones regulares,
// sinónimos, separación silábica y ordenar párrafos.

import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase5f');
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
const E = (w, fn, a) => w.evaluate(fn, a);
const text = (w) => E(w, () => document.getElementById('editor').innerText.replace(/­/g, ''));
try {
  const w = await openApp('escribiente');
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  await E(w, () => { try { localStorage.removeItem('oficina.autocorrect'); } catch { /* */ } });
  // 1) autocorrección al escribir
  await w.click('#editor'); await w.keyboard.press('Control+End');
  await w.keyboard.type('hola mundo. esto es (c) una prueba --> ademas ', { delay: 5 });
  assert.equal((await text(w)).trim(), 'Hola mundo. Esto es © una prueba → además');
  await w.keyboard.type('dijo "cita" y MUndo palabras *fuerte* fin. ', { delay: 5 });
  const html = await E(w, () => document.getElementById('editor').innerHTML);
  assert.match(html, /dijo “cita” y Mundo palabras <b>fuerte<\/b>(&nbsp;| )fin\./, html);
  await w.keyboard.type('visita www.ejemplo.com hoy', { delay: 5 });
  assert.ok(await E(w, () => !!document.querySelector('#editor a[href="https://www.ejemplo.com"]')));
  log('Autocorrección: mayúscula al empezar oración, tabla de reemplazo, comillas, DOs MAyúsculas, *negrita* y vínculos');
  // 2) buscar con expresiones regulares y palabras completas
  await E(w, () => window.__escribiente.setContent('<p>Tel 555-1234 y 555-9876.</p><p>casa casamiento Casa</p>')); await wait(200);
  await w.evaluate(() => window.__shell.command('replace')); await wait(200);
  await w.check('.findbar input[type=checkbox] >> nth=3'); // expresiones regulares
  await w.fill(`.findbar input[aria-label="${EN ? "Find" : "Buscar"}"]`, '(\\d{3})-(\\d{4})'); await wait(200);
  assert.match(await E(w, () => document.querySelector('.findbar .count').textContent), /1 de 2|1 of 2/);
  await w.fill(`.findbar input[aria-label="${EN ? "Replace with" : "Reemplazar con"}"]`, '$2/$1');
  await w.click('.findbar .row:nth-child(2) .btn:nth-child(3)'); await wait(200);
  assert.match(await text(w), /Tel 1234\/555 y 9876\/555\./);
  await w.uncheck('.findbar input[type=checkbox] >> nth=3'); await w.check('.findbar input[type=checkbox] >> nth=2');
  await w.fill(`.findbar input[aria-label="${EN ? "Find" : "Buscar"}"]`, 'casa'); await wait(200);
  assert.match(await E(w, () => document.querySelector('.findbar .count').textContent), /1 de 2|1 of 2/);
  await w.keyboard.press('Escape');
  log('Buscar y reemplazar con expresiones regulares ($1, $2) y palabras completas');
  // 3) sinónimos
  await E(w, () => window.__escribiente.setContent('<p id="s1">Estoy muy feliz hoy.</p>')); await wait(200);
  await E(w, () => { const t = document.querySelector('#s1').firstChild; const r = document.createRange(); r.setStart(t, 11); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); window.__escribiente.saveSelection(); setTimeout(() => window.__escribiente.onCommand('thesaurus')); });
  await w.waitForSelector('.modal .th-w', { timeout: 8000 });
  const syn = await E(w, () => document.querySelector('.modal .th-w').textContent);
  await w.click('.modal .th-w'); await w.click('.modal .btn.primary'); await wait(200);
  assert.equal(await text(w), `Estoy muy ${syn} hoy.`);
  log(`Sinónimos (feliz → ${syn})`);
  // 4) separación silábica
  const parts = await E(w, () => window.__escribiente.HT.hyphenate('extraordinariamente'));
  assert.deepEqual(parts, ['ex', 'tra', 'or', 'di', 'na', 'ria', 'men', 'te']);
  assert.deepEqual(await E(w, () => window.__escribiente.HT.hyphenate('responsabilidades')), ['res', 'pon', 'sa', 'bi', 'li', 'da', 'des']);
  await E(w, () => window.__escribiente.setContent('<p id="h1">Las responsabilidades administrativas extraordinariamente complicadas.</p>')); await wait(200);
  await startDialog(w, 'hyphenation'); await fillForm(w, { auto: true });
  assert.ok(await E(w, () => document.querySelector('#h1').textContent.includes('­')));
  assert.equal(await E(w, () => window.__escribiente.HT.findRanges(document.getElementById('editor'), 'extraordinariamente').length), 1);
  assert.match(await E(w, () => document.querySelector('.status, footer')?.textContent || document.body.innerText), /5 palabras|5 words/);
  log('Separación silábica (patrones de Liang) con guiones opcionales; buscar y contar palabras los ignoran');
  // 5) ordenar párrafos
  await E(w, () => window.__escribiente.setContent('<p>pera</p><p>Manzana</p><p>10 uvas</p><p>2 kiwis</p>')); await wait(200);
  await E(w, () => { const ps = document.querySelectorAll('#editor p'); const r = document.createRange(); r.setStart(ps[0].firstChild, 0); r.setEnd(ps[3].firstChild, 3); getSelection().removeAllRanges(); getSelection().addRange(r); window.__escribiente.saveSelection(); });
  await startDialog(w, 'sortParas'); await fillForm(w, { order: 'asc' });
  assert.deepEqual(await E(w, () => [...document.querySelectorAll('#editor p')].map((p) => p.textContent)), ['2 kiwis', '10 uvas', 'Manzana', 'pera']);
  log('Ordenar párrafos (alfanumérico, números en orden natural)');
  await closeApp('escribiente'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de la fase 5 (parte 6) pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
