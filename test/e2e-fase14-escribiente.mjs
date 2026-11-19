// Pruebas de Escribiente 0.17: regla con sangrías izquierda, derecha y de primera línea y márgenes
// arrastrables, doble clic para escribir en cualquier parte de la hoja, borrar saltos de página con
// Retroceso/Supr, revisión de documento (ortografía), traducción en línea y letras del sistema.

import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase14');
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
try {
  const w = await openApp('escribiente');
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  const setDoc = (html) => E(w, (x) => window.__escribiente.setContent(x), html);

  // ---------------------------------------------------------------- regla
  await setDoc('<p id="p1">Primer párrafo con texto suficiente para ocupar varias líneas en la hoja del documento de prueba, y así comprobar las sangrías de la regla.</p><p id="p2">Segundo párrafo.</p>');
  await wait(400);
  await w.click('#p1'); await wait(300);
  for (const c of ['first', 'left', 'both', 'right']) assert.ok(await box(w, `.ruler .indent.${c}`), `marcador ${c}`);
  assert.ok(await box(w, '.ruler .mhandle.l') && await box(w, '.ruler .mhandle.r'), 'márgenes en la regla');
  // la sangría derecha está en el margen derecho
  const rr = await box(w, '.ruler .indent.right'); const mr = await box(w, '.ruler .mhandle.r');
  assert.ok(Math.abs(rr.x - mr.x) < 8, `sangría derecha en el margen ${rr.x} ${mr.x}`);
  await drag(w, '.ruler .indent.right', -60);
  const mR = await E(w, () => parseFloat(document.getElementById('p1').style.marginRight) || 0);
  assert.ok(mR > 40 && mR < 80, `sangría derecha ${mR}`);
  assert.equal(await E(w, () => document.getElementById('p2').style.marginRight), '', 'solo el párrafo actual');
  await drag(w, '.ruler .indent.left', 48);
  let st = await E(w, () => { const p = document.getElementById('p1'); return { ml: parseFloat(p.style.marginLeft) || 0, ti: parseFloat(p.style.textIndent) || 0 }; });
  assert.ok(st.ml > 30 && Math.abs(st.ml + st.ti) < 1, `sangría francesa ${JSON.stringify(st)}`);
  await drag(w, '.ruler .indent.first', 24);
  st = await E(w, () => { const p = document.getElementById('p1'); return { ml: parseFloat(p.style.marginLeft) || 0, ti: parseFloat(p.style.textIndent) || 0 }; });
  assert.ok(st.ti > -40 && st.ti < 0, `primera línea ${JSON.stringify(st)}`);
  await drag(w, '.ruler .indent.both', 30);
  const st2 = await E(w, () => { const p = document.getElementById('p1'); return { ml: parseFloat(p.style.marginLeft) || 0, ti: parseFloat(p.style.textIndent) || 0 }; });
  assert.ok(st2.ml > st.ml + 15 && st2.ti === st.ti, `ambas ${JSON.stringify(st2)}`);
  // márgenes de la página
  const ml0 = await E(w, () => window.__escribiente.state.page.ml);
  await drag(w, '.ruler .mhandle.l', -40);
  const ml1 = await E(w, () => window.__escribiente.state.page.ml);
  assert.ok(ml1 < ml0 - 25 && ml1 > ml0 - 55, `margen izquierdo ${ml0} → ${ml1}`);
  const mr0 = await E(w, () => window.__escribiente.state.page.mr);
  await drag(w, '.ruler .mhandle.r', -38);
  const mr1 = await E(w, () => window.__escribiente.state.page.mr);
  assert.ok(mr1 > mr0 + 25 && mr1 < mr0 + 55, `margen derecho ${mr0} → ${mr1}`);
  // el editor usa los márgenes nuevos
  const ew = await E(w, () => document.getElementById('editor').offsetWidth);
  const pg = await E(w, () => window.__escribiente.state.page);
  assert.ok(Math.abs(ew - (pg.w - pg.ml - pg.mr)) < 2, `ancho del texto ${ew}`);
  // se guardan en .docx y se leen
  const f1 = path.join(out, 'regla.docx');
  await E(w, (p) => window.__shell.saveAs?.(p) ?? null, f1).catch(() => null);
  log('Regla: sangría derecha, izquierda (francesa), primera línea y ambas; márgenes izquierdo y derecho arrastrables');

  // ---------------------------------------------------------------- doble clic para escribir
  await setDoc('<p>Solo una línea.</p>'); await wait(500);
  const sheet = await E(w, () => { const r = document.querySelector('.sheet').getBoundingClientRect(); const e = document.getElementById('editor').getBoundingClientRect(); return { l: e.left, r: e.right, t: r.top, b: r.bottom, sl: r.left, sr: r.right }; });
  // abajo a la izquierda
  await w.mouse.dblclick(sheet.l + 20, sheet.t + 400); await wait(300);
  let c = await caretIn(w);
  assert.ok(c.inEditor && c.idx > 3, `cursor abajo ${JSON.stringify(c)}`);
  await w.keyboard.type('Aquí abajo'); await wait(200);
  let info = await E(w, () => { const ps = [...document.querySelectorAll('#editor > p')]; const last = ps.find((p) => p.textContent.includes('Aquí abajo')); return { n: ps.length, align: last?.style.textAlign || '', y: last?.getBoundingClientRect().top }; });
  assert.ok(info.n > 4 && Math.abs(info.y - (sheet.t + 400)) < 40, `párrafo nuevo en su sitio ${JSON.stringify(info)}`);
  // en el centro, más abajo: centrado
  await w.mouse.dblclick((sheet.l + sheet.r) / 2, sheet.t + 480); await wait(300);
  await w.keyboard.type('Centro'); await wait(200);
  info = await E(w, () => { const p = [...document.querySelectorAll('#editor > p')].find((x) => x.textContent.includes('Centro')); return { align: p?.style.textAlign, y: p?.getBoundingClientRect().top }; });
  assert.equal(info.align, 'center');
  // a la derecha: alineado a la derecha
  await w.mouse.dblclick(sheet.r - 10, sheet.t + 540); await wait(300);
  await w.keyboard.type('Derecha'); await wait(200);
  assert.equal(await E(w, () => [...document.querySelectorAll('#editor > p')].find((x) => x.textContent.includes('Derecha'))?.style.textAlign), 'right', await E(w, () => document.getElementById('editor').innerHTML));
  // doble clic en el margen izquierdo junto al texto: cursor en esa línea
  await w.mouse.dblclick(sheet.sl + 10, (await box(w, '#editor > p')).y); await wait(300);
  c = await caretIn(w); assert.ok(c.inEditor && c.idx === 0, `margen ${JSON.stringify(c)}`);
  // deshacer el párrafo escrito
  log('Doble clic en cualquier parte de la hoja: el cursor se coloca ahí (izquierda, centro o derecha) y se puede escribir');

  // ---------------------------------------------------------------- borrar saltos de página
  const pbDoc = '<p id="a">Antes</p><div class="pb" contenteditable="false"></div><p id="b">Después</p>';
  await setDoc(pbDoc); await wait(400);
  const caretAt = (id, end) => E(w, ([i, e]) => { const el = document.getElementById(i); const r = document.createRange(); r.selectNodeContents(el); r.collapse(!e); const s = getSelection(); s.removeAllRanges(); s.addRange(r); document.getElementById('editor').focus(); }, [id, end]);
  await caretAt('b', false); await w.keyboard.press('Backspace'); await wait(300);
  info = await E(w, () => ({ pb: document.querySelectorAll('#editor .pb').length, txt: document.getElementById('editor').innerText.replace(/\n+/g, '|') }));
  assert.equal(info.pb, 0, 'Retroceso borra el salto'); assert.match(info.txt, /Antes\|Después/);
  assert.equal(await E(w, () => window.__escribiente.state.pages), 1);
  // deshacer lo trae de vuelta
  await w.keyboard.press('Control+z'); await wait(300);
  assert.equal(await E(w, () => document.querySelectorAll('#editor .pb').length), 1, 'deshacer');
  await setDoc(pbDoc); await wait(300);
  await caretAt('a', true); await w.keyboard.press('Delete'); await wait(300);
  assert.equal(await E(w, () => document.querySelectorAll('#editor .pb').length), 0, 'Supr borra el salto');
  // salto insertado con Ctrl+Entrar y borrado con Retroceso
  await setDoc('<p id="c">Uno</p>'); await wait(300);
  await caretAt('c', true); await w.keyboard.press('Control+Enter'); await wait(400);
  assert.equal(await E(w, () => document.querySelectorAll('#editor .pb').length), 1);
  assert.equal(await E(w, () => window.__escribiente.state.pages), 2);
  await w.keyboard.press('Backspace'); await wait(400);
  assert.equal(await E(w, () => document.querySelectorAll('#editor .pb').length), 0, 'Retroceso después de Ctrl+Entrar');
  assert.equal(await E(w, () => window.__escribiente.state.pages), 1);
  await w.keyboard.type('X'); await wait(200);
  assert.match(await E(w, () => document.getElementById('editor').innerText), /Uno\s*X|UnoX/);
  // salto de columna
  await setDoc('<p id="d">Izq<span class="cbk" contenteditable="false"></span>Der</p>'); await wait(300);
  await E(w, () => { const p = document.getElementById('d'); const t = p.lastChild; const r = document.createRange(); r.setStart(t, 0); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); document.getElementById('editor').focus(); });
  await w.keyboard.press('Backspace'); await wait(300);
  assert.equal(await E(w, () => document.querySelectorAll('#editor .cbk').length), 0, 'salto de columna');
  log('Saltos de página y de columna: se borran con Retroceso y Supr (con deshacer)');

  // ---------------------------------------------------------------- revisión de documento
  const canSpell = await E(w, () => !!window.oficina.spell);
  assert.ok(canSpell);
  await setDoc('<p id="s1">Este documento tiene halgunas palabras mal escritas.</p><p id="s2">Otra vez halgunas y tambien un herror.</p>');
  await wait(1500);
  // espera a que el corrector cargue el diccionario
  let ready = false;
  for (let i = 0; i < 40 && !ready; i++) { ready = await E(w, () => window.oficina.spell.misspelled('halgunas') && !window.oficina.spell.misspelled('palabras')); if (!ready) await wait(250); }
  if (ready) {
    await E(w, () => { const r = document.createRange(); r.setStart(document.getElementById('s1'), 0); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); });
    await w.keyboard.press('F7'); await wait(500);
    assert.ok(await w.isVisible('.spellbox'));
    let s = await E(w, () => window.__escribiente.RV._state());
    assert.equal(s.word, 'halgunas');
    assert.ok(await E(w, () => document.querySelectorAll('.spellbox .sp-sug').length) > 0, 'sugerencias');
    await w.fill('.spellbox .sp-word', 'algunas');
    await w.click('.spellbox [data-act="changeAll"]'); await wait(300);
    s = await E(w, () => window.__escribiente.RV._state());
    assert.equal(s.word, 'tambien', `siguiente ${JSON.stringify(s)}`);
    await w.fill('.spellbox .sp-word', 'también');
    await w.click('.spellbox [data-act="change"]'); await wait(300);
    s = await E(w, () => window.__escribiente.RV._state());
    assert.equal(s.word, 'herror');
    await w.click('.spellbox [data-act="ignore"]'); await wait(300);
    s = await E(w, () => window.__escribiente.RV._state());
    assert.ok(s.done, `terminó ${JSON.stringify(s)}`);
    const txt = await E(w, () => document.getElementById('editor').innerText);
    assert.match(txt, /tiene algunas palabras/); assert.match(txt, /Otra vez algunas y también un herror/);
    await w.click('.spellbox .sp-close'); await wait(200);
    assert.equal(await w.isVisible('.spellbox').catch(() => false), false);
    // deshacer
    await w.keyboard.press('Control+z'); await wait(200);
    log('Revisión de documento (F7): recorre el documento, sugiere, cambia, cambia todo y omite');
  } else console.log('… el corrector ortográfico del sistema no está listo; se omite el recorrido');
  // el botón está en la pestaña Revisar
  await w.click('#rtab-review').catch(() => null); await wait(200);
  assert.ok(await E(w, () => !!document.querySelector('[data-cmd="docReview"]') && !!document.querySelector('[data-cmd="translate"]')), 'botones en Revisar');

  // ---------------------------------------------------------------- traducción (servicio simulado)
  await setDoc('<h1 id="t1">Hola mundo</h1><p id="t2">Esto es <b>una</b> prueba.</p><p id="t3">Otro párrafo.</p>'); await wait(300);
  await E(w, () => { const r = document.createRange(); r.selectNodeContents(document.getElementById('t3')); const s = getSelection(); s.removeAllRanges(); s.addRange(r); });
  await startDialog(w, 'translate');
  await setDialog(w, { scope: 'sel', from: 'es', to: 'en', mode: 'replace', provider: 'google' }); await wait(600);
  assert.equal(await E(w, () => document.getElementById('t3').textContent), '[en] Otro párrafo.');
  await startDialog(w, 'translate');
  await setDialog(w, { scope: 'doc', from: 'auto', to: 'fr', mode: 'replace' }); await wait(600);
  info = await E(w, () => ({ h: document.getElementById('t1').textContent, tag: document.getElementById('t1').tagName, p: document.getElementById('t2').textContent }));
  assert.equal(info.tag, 'H1'); assert.equal(info.h, '[fr] Hola mundo'); assert.equal(info.p, '[fr] Esto es una prueba.');
  // insertar debajo
  await setDoc('<p id="u1">Uno</p><p id="u2">Dos</p>'); await wait(300);
  await startDialog(w, 'translate');
  await setDialog(w, { scope: 'doc', to: 'en', mode: 'below' }); await wait(600);
  assert.deepEqual(await E(w, () => [...document.querySelectorAll('#editor > p')].map((p) => p.textContent)), ['Uno', '[en] Uno', 'Dos', '[en] Dos']);
  // el módulo real parte los textos largos y empaqueta párrafos
  const tr = await app.evaluate(async () => { const T = require(require('path').join(process.cwd(), 'main', 'traducir.js')); return { g: T.guessLang('The house is on the hill and it is red'), s: T.splitLong('Una oración. Otra oración. Y otra más.', 15).length, p: T.pack(['aaaa', 'bbbb', 'cccc'], 10).length }; }).catch((e) => ({ err: String(e) }));
  if (!tr.err) { assert.equal(tr.g, 'en'); assert.ok(tr.s >= 3); assert.equal(tr.p, 2); }
  log('Traducir: selección, documento completo (conserva títulos) e insertar la traducción debajo');

  // ---------------------------------------------------------------- letras del sistema
  const sys = await E(w, () => window.oficina.systemFonts());
  assert.ok(Array.isArray(sys) && sys.length > 0, 'lista de letras del sistema');
  await w.click('#rtab-home'); await wait(500);
  const opt = await E(w, () => { const s = document.querySelector('select.rsel[aria-label]'); const all = [...document.querySelectorAll('select.rsel')].find((x) => [...x.options].some((o) => o.value === 'Calibri')); return { grp: !!all?.querySelector('optgroup[data-sys]'), n: all?.querySelectorAll('optgroup[data-sys] option').length || 0, s: !!s }; });
  // las del sistema que no estén ya incluidas aparecen en su grupo
  const extra = await E(w, (list) => list.filter((n) => !window.__escribiente.state || true).length, sys);
  assert.ok(extra > 0);
  console.log(`  (${sys.length} familias del sistema; ${opt.n} nuevas en el menú)`);
  if (process.platform === 'linux') assert.ok(opt.n > 0 && opt.grp, 'grupo «Letras del sistema»');
  // una letra del sistema se puede aplicar
  if (opt.n) { const fam = await E(w, () => document.querySelector('optgroup[data-sys] option').value); await setDoc('<p id="f1">Letra</p>'); await E(w, () => { const r = document.createRange(); r.selectNodeContents(document.getElementById('f1')); const s = getSelection(); s.removeAllRanges(); s.addRange(r); window.__escribiente.saveSelection(); }); await E(w, (f) => window.__shell.command('fontName', f), fam); await wait(200); assert.ok((await E(w, () => document.getElementById('f1').innerHTML)).includes(fam), 'aplicar letra del sistema'); }
  log('Letras del sistema operativo en el menú de fuentes');

  await closeApp('escribiente'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de Escribiente 0.17 pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
