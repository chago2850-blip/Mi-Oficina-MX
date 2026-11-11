// Pruebas de la fase 5, parte 5 (Escribiente como LibreOffice Writer): encabezado y pie con primera página y pares distintos,
// en tres partes y con formato; borde de página y numeración de líneas; .docx y .odt (y LibreOffice).

import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase5e');
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
try {
  const w = await openApp('escribiente');
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  await E(w, () => { let h = ''; for (let i = 0; i < 90; i++) h += `<p>Párrafo ${i + 1}: lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.</p>`; window.__escribiente.setContent(h); window.__escribiente.paginate(); });
  await wait(300);
  // 1) encabezado y pie: primera página distinta, pares distintos, tres partes
  await startDialog(w, 'headerFooter'); await w.waitForSelector('.modal');
  await w.fill('.modal [name="header"]', 'Empresa\\tInforme\\t{pagina}');
  await w.fill('.modal [name="headerFirst"]', 'PORTADA');
  await w.fill('.modal [name="headerEven"]', 'Par {pagina}');
  await w.evaluate(() => document.querySelectorAll('.modal .fx-tab')[1].click());
  await w.fill('.modal [name="footer"]', 'Pie {pagina} de {paginas}');
  await w.evaluate(() => document.querySelectorAll('.modal .fx-tab')[2].click());
  await setDialog(w, { hfFirst: true, hfOddEven: true, bold: true, line: true, size: 9 });
  await wait(400);
  const hf = await E(w, () => [...document.querySelectorAll('.sheet')].slice(0, 3).map((sh) => ({ top: [...sh.querySelectorAll('.hf.top .hf-l')].map((l) => (l.classList.contains('hf-3') ? [...l.children].map((c) => c.textContent).join('|') : l.textContent)).join('/'), bottom: sh.querySelector('.hf.bottom')?.textContent || '', w: getComputedStyle(sh.querySelector('.hf.top') || document.body).fontWeight })));
  const pages = await E(w, () => window.__escribiente.state.pages);
  assert.ok(pages >= 3, `páginas ${pages}`);
  assert.deepEqual(hf.map((x) => x.top), ['PORTADA', 'Par 2', 'Empresa|Informe|3']);
  assert.equal(hf[0].bottom, ''); assert.equal(hf[2].bottom, `Pie 3 de ${pages}`); assert.equal(hf[1].w, '700');
  log('Encabezado y pie: primera página distinta, pares distintos, tres partes (izq./centro/der.) y formato');
  // 2) borde de página y numeración de líneas
  await startDialog(w, 'pageBorder'); await w.waitForSelector('.modal'); await setDialog(w, { style: 'double', width: 1, color: '#1f4e79', space: 0.5 });
  await startDialog(w, 'lineNum'); await w.waitForSelector('.modal'); await setDialog(w, { on: true, every: 5, restart: 'page' });
  await wait(500);
  const pbl = await E(w, () => { const b = document.querySelector('.sheet .pborder'); return b && getComputedStyle(b).borderTopStyle + ' ' + getComputedStyle(b).borderTopColor; });
  assert.equal(pbl, 'double rgb(31, 78, 121)');
  const nums = await E(w, () => [...document.querySelectorAll('.lnum')].map((x) => Number(x.textContent)));
  assert.ok(nums.length > 5 && nums.every((n) => n % 5 === 0), JSON.stringify(nums.slice(0, 12)));
  assert.ok(nums.indexOf(5, 1) > 0, `reinicia en cada página ${JSON.stringify(nums)} ${JSON.stringify(await E(w, () => window.__escribiente.state.lineNum))}`);
  log('Borde de página (doble, color) y números de línea cada 5, reiniciando en cada página');
  // 3) guardar en .docx y .odt y volver a abrir
  const snap = () => E(w, () => { const s = window.__escribiente.state; return { header: s.header, headerFirst: s.headerFirst, headerEven: s.headerEven, footer: s.footer, hfFirst: !!s.hfFirst, hfOddEven: !!s.hfOddEven, border: s.pageBorder && `${s.pageBorder.style} ${s.pageBorder.color}`, ln: s.lineNum && `${s.lineNum.every} ${s.lineNum.restart}` }; });
  const exp = await snap();
  for (const [ext, lab] of [['docx', 'Word'], ['odt', 'ODT']]) {
    await w.evaluate(async ([e, l]) => window.__shell.saveAsFormat(e, l), [ext, lab]); await wait(900);
    const file = path.join(out, `${UNTITLED}.${ext}`);
    await reopen(w, file); await wait(900);
    assert.deepEqual(await snap(), exp, ext);
    if (hasSoffice) {
      const lo = path.join(out, `lo-${ext}`); fs.mkdirSync(lo, { recursive: true }); const to = ext === 'docx' ? 'odt' : 'docx';
      execFileSync('soffice', ['--headless', '--convert-to', to, '--outdir', lo, file], { stdio: 'ignore', timeout: 120000 });
      await reopen(w, path.join(lo, `${UNTITLED}.${to}`)); await wait(1000);
      const q = await snap(); console.log(`  LibreOffice ${ext}→${to}: ${JSON.stringify(q)}`);
      assert.equal(q.headerFirst, 'PORTADA', `LO ${ext}→${to}`); assert.equal(q.headerEven, 'Par {pagina}', `LO ${ext}→${to}`); assert.equal(q.header, 'Empresa\tInforme\t{pagina}', `LO ${ext}→${to}`);
      assert.ok(q.hfFirst && q.hfOddEven, `LO ${ext}→${to}`); assert.match(q.border || '', /^double #1f4e79$/i, `LO ${ext}→${to}`); assert.equal(q.ln, '5 page', `LO ${ext}→${to}`);
      await reopen(w, file); await wait(700);
    }
  }
  log(`Encabezados, borde y números de línea se guardan en .docx y .odt${hasSoffice ? ' y sobreviven a LibreOffice' : ''}`);
  await closeApp('escribiente'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de la fase 5 (parte 5) pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
