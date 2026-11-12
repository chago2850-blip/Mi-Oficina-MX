// Pruebas de las fuentes libres incluidas (las de LibreOffice y otras) en las tres apps.

import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fuentes');
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
  // Escribiente: el menú tiene todas las letras y cada una se carga de la suite (no del sistema)
  const w = await openApp('escribiente');
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  const names = await w.evaluate(() => [...document.querySelector('select[aria-label="Fuente"], select[aria-label="Font"]')?.options || []].map((o) => o.value));
  for (const f of ['Calibri', 'Liberation Sans', 'Liberation Serif', 'Liberation Mono', 'Carlito', 'Caladea', 'DejaVu Sans', 'DejaVu Serif', 'DejaVu Sans Mono', 'Linux Libertine G', 'Linux Biolinum G', 'Gentium Book Basic', 'Noto Sans', 'Noto Serif', 'Source Sans Pro', 'Source Serif Pro', 'Source Code Pro', 'Open Sans', 'EB Garamond', 'Amiri', 'Alef']) assert.ok(names.includes(f), `falta ${f}`);
  assert.ok(names.length >= 45, String(names.length));
  const loaded = await w.evaluate(async () => { const out = {}; for (const f of ['Carlito', 'Liberation Serif', 'Linux Libertine G', 'DejaVu Sans Mono', 'Amiri']) { await document.fonts.load(`16px "${f}"`); out[f] = [...document.fonts].some((ff) => ff.family.replace(/["']/g, '') === f && ff.status === 'loaded'); } return out; });
  assert.deepEqual(loaded, { Carlito: true, 'Liberation Serif': true, 'Linux Libertine G': true, 'DejaVu Sans Mono': true, Amiri: true });
  // el texto con Calibri se dibuja con Carlito (mismas medidas) aunque el equipo no tenga Calibri
  await w.evaluate(() => window.__escribiente.setContent('<p><span style="font-family: Calibri, Carlito, sans-serif">Hola mundo con Calibri</span></p><p><span style="font-family: &quot;Linux Libertine G&quot;, serif">Texto en Linux Libertine G, ñandú.</span></p>'));
  await wait(800);
  await w.evaluate(() => { window.__shell.doc.name = 'fuentes.docx'; }); await w.evaluate(() => window.__shell.exportPdf()); await wait(1500);
  const fonts = execFileSync('python3', ['-c', `import pypdf\nr=pypdf.PdfReader('${path.join(out, 'fuentes.pdf')}')\ns=set()\nfor p in r.pages:\n  for k,v in (p['/Resources'].get('/Font') or {}).items(): s.add(str(v.get_object()['/BaseFont']))\nprint(' '.join(sorted(s)))`], { encoding: 'utf8' });
  assert.match(fonts, /Carlito/); assert.match(fonts, /Libertinus|Libertine/);
  log(`Escribiente: ${names.length} letras en el menú; las libres se cargan de la suite y van incrustadas en el PDF`);
  await closeApp('escribiente'); await wait(400);
  // Calcula e Impresiona
  for (const id of ['calcula', 'impresiona']) {
    const a = await openApp(id); await a.evaluate(() => window.__shell.newDoc()); await wait(500);
    const n = await a.evaluate(() => [...document.querySelector('select[aria-label="Fuente"], select[aria-label="Font"]')?.options || []].map((o) => o.value));
    assert.ok(n.includes('Liberation Sans') && n.includes('Linux Libertine G') && n.length >= 45, `${id}: ${n.length}`);
    await closeApp(id); await wait(400);
  }
  log('Calcula Datos e Impresiona tienen la misma lista de letras');
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de las fuentes pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
