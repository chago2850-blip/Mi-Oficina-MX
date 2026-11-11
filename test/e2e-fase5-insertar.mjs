// Pruebas de la fase 5, parte 4 (Escribiente como LibreOffice Writer): campos, marcadores, referencias cruzadas,
// títulos de objetos, índice de ilustraciones, índice alfabético y salto de columna; .docx y .odt (y LibreOffice).

import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase5d');
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
const caretEnd = (w, sel) => E(w, (s) => { document.getElementById('editor').focus(); const el = document.querySelector(s); const r = document.createRange(); r.selectNodeContents(el); r.collapse(false); const g = getSelection(); g.removeAllRanges(); g.addRange(r); window.__escribiente.saveSelection(); }, sel);

try {
  const w = await openApp('escribiente');
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAAASFvFNAAAAFklEQVR4nGP4z8DAwMDAxMDAwMDAAAANHQEDK+mmyAAAAABJRU5ErkJggg==';
  await E(w, (png) => { let h = '<h1 id="h1a">Introducción</h1><p id="p1">Fecha: </p><p id="p2">Página </p>'; for (let i = 0; i < 40; i++) h += `<p>Relleno ${i} lorem ipsum dolor sit amet, consectetur adipiscing elit.</p>`; h += `<h1>Resultados</h1><p id="pimg"><img src="${png}" style="width:200px;height:100px"></p><table id="t1"><tbody><tr><th>A</th><th>B</th></tr><tr><td>1</td><td>2</td></tr></tbody></table><p id="p3">Ver </p><p id="p4">Fin</p>`; window.__escribiente.setContent(h); window.__escribiente.paginate(); }, PNG);
  await wait(300);
  // 1) campos: fecha con formato, número de página y total de páginas
  await caretEnd(w, '#p1'); await E(w, () => window.__escribiente.IN.insertField('date', { fmt: 'yyyy-MM-dd' }));
  await caretEnd(w, '#p2'); await E(w, () => window.__escribiente.IN.insertField('page'));
  await E(w, () => { document.getElementById('editor').focus(); const p = document.querySelector('#p2'); p.append(' de '); const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); window.__escribiente.saveSelection(); window.__escribiente.IN.insertField('pages'); });
  await caretEnd(w, '#p4'); await E(w, () => window.__escribiente.IN.insertField('page'));
  await wait(400); await E(w, () => window.__escribiente.paginate()); await wait(200);
  if (process.env.DUMP) console.log(await E(w, () => ['#p1', '#p2', '#p4'].map((s) => document.querySelector(s)?.outerHTML).join('\n')));
  const today = new Date(); const iso = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
  const f1 = await E(w, () => ({ date: document.querySelector('#p1 .fld').textContent, p2: document.querySelector('#p2').textContent, p4: document.querySelector('#p4 .fld').textContent, pages: window.__escribiente.state.pages }));
  assert.equal(f1.date, iso); assert.equal(f1.p2, `Página 1 de ${f1.pages}`); assert.ok(f1.pages >= 2); assert.equal(f1.p4, String(f1.pages), JSON.stringify(f1));
  log('Campos: fecha con formato, número de página y total de páginas (se actualizan al paginar)');
  // 2) títulos de objetos (Figura 1, Tabla 1), marcador y referencias cruzadas
  await caretEnd(w, '#pimg'); await E(w, () => window.__escribiente.IN.insertCaption('Figura', 'Logotipo', { near: document.querySelector('#pimg'), pos: 'below' }));
  await E(w, () => window.__escribiente.IN.insertCaption('Tabla', 'Datos', { near: document.querySelector('#t1'), pos: 'above' }));
  await E(w, () => window.__escribiente.IN.insertCaption('Figura', 'Otra', { near: document.querySelector('#p4'), pos: 'above' }));
  const caps = await E(w, () => [...document.querySelectorAll('p.caption')].map((p) => p.textContent));
  assert.deepEqual(caps, EN ? ['Figure 1: Logotipo', 'Table 1: Datos', 'Figure 2: Otra'] : ['Figura 1: Logotipo', 'Tabla 1: Datos', 'Figura 2: Otra']);
  await caretEnd(w, '#h1a'); await E(w, () => window.__escribiente.IN.insertBookmark('intro'));
  await caretEnd(w, '#p3');
  const capId = await E(w, () => document.querySelectorAll('p.caption')[0].querySelector('a.bm').id);
  await E(w, (id) => { window.__escribiente.saveSelection(); window.__escribiente.IN.insertInline(`<span class="fld" data-fld="ref" data-ref="${id}" data-show="label" contenteditable="false">…</span> (p. <span class="fld" data-fld="ref" data-ref="${id}" data-show="page" contenteditable="false">…</span>) y <span class="fld" data-fld="ref" data-ref="intro" data-show="text" contenteditable="false">…</span>`); window.__escribiente.IN.updateFields(); }, capId);
  const pimgPage = await E(w, () => window.__escribiente.IN.pageOfNode(document.querySelector('#pimg')));
  assert.equal(await E(w, () => document.querySelector('#p3').textContent), `Ver ${EN ? 'Figure' : 'Figura'} 1 (p. ${pimgPage}) y Introducción`);
  // al insertar una figura antes, la numeración y la referencia cambian
  await E(w, () => { window.__escribiente.IN.insertCaption('Figura', 'Nueva', { near: document.querySelector('#p1'), pos: 'below' }); window.__escribiente.IN.updateFields(); });
  assert.match(await E(w, () => document.querySelector('#p3').textContent), EN ? /^Ver Figure 2 / : /^Ver Figura 2 /);
  log('Títulos (Figura/Tabla con numeración automática), marcadores y referencias cruzadas que se actualizan');
  // 3) índice de ilustraciones y alfabético
  await caretEnd(w, '#p4');
  await E(w, () => { const p = document.querySelector('#p1'); const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); window.__escribiente.saveSelection(); window.__escribiente.IN.insertInline('<span class="ix" data-ix="Fecha" contenteditable="false"></span>'); });
  await E(w, () => { const p = document.querySelector('#p4'); const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); window.__escribiente.IN.insertInline('<span class="ix" data-ix="Apéndice" contenteditable="false"></span>'); window.__escribiente.saveSelection(); });
  await E(w, () => window.__escribiente.IN.insertIndex('fig', 'Figura')); await wait(300);
  await caretEnd(w, '#p4'); await E(w, () => window.__escribiente.IN.insertIndex('alpha')); await wait(300);
  const idx = await E(w, () => ({ fig: [...document.querySelectorAll('.toc[data-kind="fig"] .toc-l')].map((l) => l.children[0].textContent + '|' + l.children[2]?.textContent), alpha: [...document.querySelectorAll('.toc[data-kind="alpha"] .toc-l')].map((l) => l.textContent) }));
  assert.equal(idx.fig.length, 3); assert.match(idx.fig[0], /1: Nueva\|1$/); assert.match(idx.fig[1], /2: Logotipo\|\d+$/);
  assert.equal(idx.alpha[0], 'A'); assert.match(idx.alpha[1], /^Apéndice\d+$/); assert.equal(idx.alpha[2], 'F');
  log('Índice de ilustraciones e índice alfabético (entradas marcadas, agrupadas por letra, con páginas)');
  // 4) salto de columna en una sección de dos columnas
  await E(w, () => { const ed = window.__escribiente.editor; const d = document.createElement('div'); d.className = 'cols'; d.dataset.n = '2'; d.style.columnCount = '2'; d.innerHTML = '<p id="c1">Columna uno</p><p id="c2">Columna dos</p>'; ed.append(d); });
  await caretEnd(w, '#c1'); await E(w, () => window.__escribiente.IN.insertColumnBreak());
  const cb = await E(w, () => { const a = document.querySelector('#c1').getBoundingClientRect(), b = document.querySelector('#c2').getBoundingClientRect(); return b.left - a.left; });
  assert.ok(cb > 100, `la segunda columna debe empezar a la derecha (${cb})`);
  log('Salto de columna');
  // 5) guardar en .docx y .odt y volver a abrir
  const snap = () => E(w, () => { const ed = window.__escribiente.editor; const F = (k) => [...ed.querySelectorAll(`.fld[data-fld="${k}"]`)]; return { date: F('date')[0]?.dataset.fmt, page: F('page').length, pages: F('pages').length, seq: F('seq').map((s) => `${s.dataset.seq}${s.textContent}`).join(','), caps: ed.querySelectorAll('p.caption').length, refs: F('ref').map((r) => r.dataset.show).join(','), refsText: document.querySelector('#p3, p:has(> .fld[data-fld="ref"])')?.textContent, bm: !!ed.querySelector('a.bm#intro'), ix: [...ed.querySelectorAll('.ix')].map((x) => x.dataset.ix).sort().join(','), cbk: ed.querySelectorAll('.cbk').length }; });
  const exp = await snap();
  for (const [ext, lab] of [['docx', 'Word'], ['odt', 'ODT']]) {
    await w.evaluate(async ([e, l]) => window.__shell.saveAsFormat(e, l), [ext, lab]); await wait(900);
    const file = path.join(out, `${UNTITLED}.${ext}`);
    const z = await JSZip.loadAsync(fs.readFileSync(file));
    if (ext === 'docx') { const d = await z.file('word/document.xml').async('string'); assert.match(d, /SEQ Figura/); assert.match(d, /DATE \\@ &quot;yyyy-MM-dd&quot;|DATE \\@ "yyyy-MM-dd"/); assert.match(d, /w:bookmarkStart[^>]*w:name="intro"/); assert.match(d, /REF _Ref/); assert.match(d, /w:br w:type="column"/); }
    else { const d = await z.file('content.xml').async('string'); assert.match(d, /<text:sequence text:name="Figura"/); assert.match(d, /<text:bookmark text:name="intro"\/>/); assert.match(d, /<text:sequence-ref|<text:bookmark-ref/); assert.match(d, /fo:break-after="column"/); assert.match(d, /<text:date style:data-style-name/); }
    await reopen(w, file); await wait(900);
    const r = await snap(); assert.deepEqual({ ...r, refsText: undefined }, { ...exp, refsText: undefined }, `${ext}: ${JSON.stringify(r)}`);
    if (hasSoffice) {
      const lo = path.join(out, `lo-${ext}`); fs.mkdirSync(lo, { recursive: true }); const to = ext === 'docx' ? 'odt' : 'docx';
      execFileSync('soffice', ['--headless', '--convert-to', to, '--outdir', lo, file], { stdio: 'ignore', timeout: 120000 });
      await reopen(w, path.join(lo, `${UNTITLED}.${to}`)); await wait(1000);
      const q = await snap();
      console.log(`  LibreOffice ${ext}→${to}: ${JSON.stringify(q)}`);
      assert.ok(q.page >= 1 && q.pages >= 1, `LO ${ext}→${to} páginas`); assert.equal(q.caps, exp.caps, `LO ${ext}→${to} títulos`); assert.match(q.seq, /Figura1/, `LO ${ext}→${to} seq`); assert.ok(q.bm, `LO ${ext}→${to} marcador`); assert.ok(q.refs.split(',').length >= 2, `LO ${ext}→${to} refs`); assert.equal(q.ix, 'Apéndice,Fecha', `LO ${ext}→${to} índice`); assert.equal(q.cbk, 1, `LO ${ext}→${to} salto de columna`); assert.doesNotMatch(q.refsText, /undefined/);
      await reopen(w, file); await wait(700);
    }
  }
  log(`Campos, títulos, marcadores, referencias e índices se guardan en .docx y .odt${hasSoffice ? ' y sobreviven a LibreOffice' : ''}`);
  await closeApp('escribiente'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de la fase 5 (parte 4) pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
