// Pruebas de la fase 5, parte 2 (Escribiente como LibreOffice Writer): estilos de párrafo y de carácter,
// estilo siguiente, listas multinivel y numeración de capítulos; se conservan en .docx y .odt.
// Linux sin pantalla:  xvfb-run -a node test/e2e-fase5.mjs
import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase5b');
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

try {
  const w = await openApp('escribiente');
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(400);
  // 1) estilo predeterminado y estilo siguiente: después de un Título 1 viene Normal
  await w.click('#editor'); await w.keyboard.type('Introducción');
  await w.evaluate(() => { window.__escribiente.saveSelection(); window.__escribiente.ES.applyPara('Heading1'); });
  await w.keyboard.press('End'); await w.keyboard.press('Enter'); await w.keyboard.type('Texto con estilo propio.');
  const s1 = await w.evaluate(() => [...document.querySelectorAll('#editor > *')].map((e) => e.tagName + (e.dataset.style ? `[${e.dataset.style}]` : '')));
  assert.deepEqual(s1, ['H1', 'P'], JSON.stringify(s1));
  // 2) nuevo estilo a partir de la selección (formato directo → estilo)
  await w.evaluate(() => { const p = document.querySelector('#editor p'); p.style.color = 'rgb(192, 0, 0)'; p.style.marginLeft = '1cm'; p.style.fontSize = '14pt'; });
  await w.evaluate(() => { const t = document.querySelector('#editor p').firstChild; const r = document.createRange(); r.setStart(t, 3); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); window.__escribiente.saveSelection(); setTimeout(() => window.__escribiente.ES.newFromSelection()); });
  await fillForm(w, { name: 'Destacado' }); await wait(300);
  const s2 = await w.evaluate(() => { const p = document.querySelector('#editor p'); const cs = getComputedStyle(p); return { ds: p.dataset.style, inline: p.getAttribute('style') || '', color: cs.color, ml: Math.round(parseFloat(cs.marginLeft)), st: window.__escribiente.state.styles[p.dataset.style] }; });
  assert.ok(s2.ds); assert.ok(!/color|margin-left/.test(s2.inline), s2.inline); assert.equal(s2.color, 'rgb(192, 0, 0)'); assert.equal(s2.ml, 38); assert.equal(s2.st.name, 'Destacado');
  // al cambiar el estilo cambian todos los párrafos que lo usan
  await w.keyboard.press('End'); await w.keyboard.press('Enter'); await w.keyboard.type('Otro párrafo destacado.');
  await w.evaluate(() => { const E = window.__escribiente; E.state.styles[document.querySelector('#editor p').dataset.style].css.color = '#0000aa'; E.ES.render(); });
  const s3 = await w.evaluate(() => [...document.querySelectorAll('#editor p')].map((p) => [p.dataset.style, getComputedStyle(p).color]));
  assert.deepEqual(s3, [[s2.ds, 'rgb(0, 0, 170)'], [s2.ds, 'rgb(0, 0, 170)']], JSON.stringify(s3));
  // 3) estilo de carácter
  await w.evaluate(() => { const t = document.querySelectorAll('#editor p')[1].firstChild; const r = document.createRange(); r.setStart(t, 0); r.setEnd(t, 4); const s = getSelection(); s.removeAllRanges(); s.addRange(r); window.__escribiente.saveSelection(); window.__escribiente.ES.applyChar('Strong'); });
  const s4 = await w.evaluate(() => { const sp = document.querySelector('#editor span[data-cstyle="Strong"]'); return sp && [sp.textContent, getComputedStyle(sp).fontWeight]; });
  assert.deepEqual(s4, ['Otro', '700']);
  // 4) lista multinivel 1. / 1.1. y numeración de capítulos
  await w.evaluate(() => { const p = document.querySelectorAll('#editor p')[1]; const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); });
  await w.keyboard.press('Enter'); await w.keyboard.type('Uno');
  await w.evaluate(() => { window.__escribiente.saveSelection(); window.__escribiente.ES.applyListPreset('m-legal'); });
  if (process.env.DUMP) console.log(await w.evaluate(() => { const s = getSelection(); return document.getElementById('editor').innerHTML + ' || ' + s.anchorNode?.nodeName + ':' + s.anchorNode?.textContent + '@' + s.anchorOffset; }));
  await w.keyboard.press('Enter'); await w.keyboard.type('Uno punto uno');
  await w.keyboard.press('Home'); await w.keyboard.press('Tab'); await wait(100); await w.keyboard.press('End');
  await w.keyboard.press('Enter'); await w.keyboard.press('Shift+Tab'); await w.keyboard.type('Dos');
  await w.evaluate(() => window.__escribiente.ES.setChapterNumbering('1.'));
  await w.evaluate(() => { const h = document.createElement('h1'); h.textContent = 'Desarrollo'; document.getElementById('editor').append(h); });
  await wait(300);
  // panel de estilos (doble clic aplica) y cuadro Modificar estilo
  await w.evaluate(() => window.__shell.command('stylesPane')); await wait(300);
  const pane = await w.evaluate(() => [...document.querySelectorAll('.st-item')].map((x) => x.dataset.id));
  assert.ok(pane.includes('Normal') && pane.includes('Heading1') && pane.includes(s2.ds), JSON.stringify(pane));
  await w.evaluate((id) => { setTimeout(() => window.__escribiente.ES.modifyStyle(id)); }, await w.evaluate(() => document.querySelector('#editor p').dataset.style));
  await w.waitForSelector('.modal'); await w.click('.modal .btn.primary'); await wait(300);
  await w.evaluate(() => window.__shell.command('stylesPane')); await wait(200);
  if (process.env.DUMP) console.log('FIN', await w.evaluate(() => document.getElementById('editor').innerHTML));
  const s5 = await w.evaluate(() => ({ txt: [...document.querySelectorAll('#editor li')].map((l) => l.firstChild?.textContent).join('|'), ml: document.querySelector('#editor ol')?.dataset.ml, nested: document.querySelectorAll('#editor ol ol li').length, top: document.querySelectorAll('#editor ol[data-ml] > li').length, h: [...document.querySelectorAll('#editor h1')].map((h) => getComputedStyle(h, '::before').content) }));
  assert.equal(s5.txt, 'Uno|Uno punto uno|Dos', JSON.stringify(s5)); assert.equal(s5.ml, 'm-legal'); assert.equal(s5.nested, 1); assert.equal(s5.top, 2, JSON.stringify(s5)); assert.match(s5.h[0], /counter\(ch1|^"1\. "|1/); assert.equal(s5.h.length, 2);
  if (process.env.DUMP) console.log(await w.evaluate(() => document.getElementById('editor').innerHTML));
  log('Estilos: estilo siguiente, nuevo estilo desde la selección, modificar el estilo cambia todo, estilo de carácter, lista 1./1.1. y capítulos');
  // 5) se guardan y se recuperan en .docx y .odt
  const check = () => w.evaluate(() => { const E = window.__escribiente; const S = Object.values(E.state.styles).find((x) => x.name === 'Destacado'); return { st: S && { name: S.name }, paras: S ? document.querySelectorAll(`#editor p[data-style="${S.id}"]`).length : 0, cs: document.querySelector('#editor span[data-cstyle="Strong"]')?.textContent, ch: E.state.chapnum, ml: document.querySelector('#editor ol')?.dataset.ml, nested: document.querySelectorAll('#editor ol ol li').length, colorP: getComputedStyle(document.querySelector('#editor p[data-style]') || document.body).color }; });
  const expectBack = (r, lab) => {
    assert.equal(r.st?.name, 'Destacado', `${lab}: ${JSON.stringify(r)}`); assert.ok(r.paras >= 2, `${lab}: ${JSON.stringify(r)}`); assert.equal(r.colorP, 'rgb(0, 0, 170)', lab);
    assert.equal(r.cs, 'Otro', `${lab}: ${JSON.stringify(r)}`); assert.equal(r.ch, '1.', `${lab}: ${JSON.stringify(r)}`); assert.equal(r.ml, 'm-legal', `${lab}: ${JSON.stringify(r)}`); assert.equal(r.nested, 1, lab);
  };
  for (const [ext, lab] of [['docx', 'Word'], ['odt', 'ODT']]) {
    await w.evaluate(async ([e, l]) => window.__shell.saveAsFormat(e, l), [ext, lab]); await wait(800);
    const file = path.join(out, `${UNTITLED}.${ext}`);
    const z = await JSZip.loadAsync(fs.readFileSync(file));
    if (ext === 'docx') {
      const sx = await z.file('word/styles.xml').async('string'); const nx = await z.file('word/numbering.xml').async('string');
      assert.match(sx, /<w:name w:val="Destacado"\/>/); assert.match(sx, /w:type="character" w:styleId="StrongEmphasis"/); assert.match(nx, /w:lvlText w:val="%1\.%2\."/);
    } else {
      const sx = await z.file('styles.xml').async('string');
      assert.match(sx, /style:display-name="Destacado"/); assert.match(sx, /<text:outline-level-style text:level="1" style:num-format="1"/);
    }
    await reopen(w, file); await wait(800);
    expectBack(await check(), ext);
    if (hasSoffice) {
      const lo = path.join(out, `lo-${ext}`); fs.mkdirSync(lo, { recursive: true });
      const to = ext === 'docx' ? 'odt' : 'docx';
      execFileSync('soffice', ['--headless', '--convert-to', to, '--outdir', lo, file], { stdio: 'ignore', timeout: 120000 });
      await reopen(w, path.join(lo, `${UNTITLED}.${to}`)); await wait(900);
      expectBack(await check(), `LibreOffice ${ext}→${to}`);
    }
  }
  log(`Estilos, listas y capítulos se guardan en .docx y .odt${hasSoffice ? ' y sobreviven al pasar por LibreOffice' : ''}`);
  await closeApp('escribiente'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de la fase 5 (parte 2) pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
