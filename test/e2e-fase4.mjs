// Pruebas de las funciones de la fase 4.
// Linux sin pantalla:  xvfb-run -a node test/e2e-fase4.mjs [escribiente|calcula|impresiona|suite]
import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase4');
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
  // ============================================================ Escribiente: lo que quedó parcial
  if (!only || only === 'escribiente') {
    const w = await openApp('escribiente');
    // .doc de Word 97-2003 con formato
    fs.copyFileSync(path.join(root, 'test', 'fixtures', 'antiguo.doc'), path.join(out, 'antiguo.doc'));
    await reopen(w, path.join(out, 'antiguo.doc')); await wait(500);
    const doc = await w.evaluate(() => { const ed = window.__escribiente.editor; return { h1: ed.querySelector('h1')?.textContent, h2: ed.querySelector('h2')?.textContent, b: ed.querySelector('b')?.textContent, i: ed.querySelector('i')?.textContent, td: [...ed.querySelectorAll('td')].map((x) => x.textContent), li: ed.querySelectorAll('li').length, center: ed.querySelector('p[style*="center"]')?.textContent, red: !!ed.querySelector('span[style*="color"]') }; });
    assert.equal(doc.h1, 'Documento antiguo'); assert.equal(doc.h2, 'Datos'); assert.equal(doc.b, 'negritas'); assert.equal(doc.i, 'cursivas');
    assert.deepEqual(doc.td, ['Producto', 'Precio', 'Café', '$45']); assert.equal(doc.li, 2); assert.ok(doc.center?.includes('canción') && doc.red, JSON.stringify(doc));
    log('Escribiente: abre .doc de Word 97-2003 con títulos, negritas, tablas, listas, alineación y color');
    // documento con ecuación, diagrama, citas y comentario con reacciones
    await w.evaluate(() => window.__shell.newDoc('blank'));
    await w.evaluate(() => {
      const E = window.__escribiente; const { editor, state, F3 } = E;
      state.sources = [{ id: 's1', type: 'libro', author: 'Pérez, Ana', title: 'Agua', year: '2020', publisher: 'UNAM', url: '' }];
      const svg = F3.smartArtSvg({ type: 'proceso', items: [{ lvl: 0, text: 'Uno' }, { lvl: 0, text: 'Dos' }], color: '#5B48E8' }).svg;
      editor.innerHTML = `<h1>Informe</h1><p><span class="cm" data-cid="c1">Texto comentado</span> con una cita <span class="cite" contenteditable="false" data-src="s1">(Pérez, 2020)</span> y la ecuación <span class="eq" contenteditable="false" data-tex="\\frac{a}{b}+\\sqrt{x}">${F3.renderTex('\\frac{a}{b}+\\sqrt{x}')}</span>.</p><p><img class="smartart" width="320" height="80" data-smartart='{"type":"proceso","text":"Uno\\nDos","color":"#5B48E8"}' src="data:image/svg+xml;base64,${btoa(svg)}"></p>`;
      state.comments = [{ id: 'c1', author: 'Ana', date: new Date().toISOString(), text: 'Revisar', replies: [{ author: 'Luis', date: new Date().toISOString(), text: 'Listo', reactions: { '🎉': ['Ana'] } }], resolved: false, reactions: { '👍': ['Luis', 'Ana'] } }];
      E.paginate();
    });
    await w.evaluate(() => window.__shell.save());
    const dx = path.join(out, `${UNTITLED}.docx`);
    await reopen(w, dx); await wait(500);
    const back = await w.evaluate(() => { const { editor, state } = window.__escribiente; const c = state.comments[0]; return { react: c?.reactions?.['👍']?.length, reply: c?.replies?.[0]?.reactions?.['🎉']?.length, sa: editor.querySelector('img.smartart')?.dataset.smartart, cite: editor.querySelector('.cite')?.dataset.src, sources: state.sources.length, tex: editor.querySelector('.eq')?.dataset.tex }; });
    assert.deepEqual(back, { react: 2, reply: 1, sa: '{"type":"proceso","text":"Uno\\nDos","color":"#5B48E8"}', cite: 's1', sources: 1, tex: '\\frac{a}{b}+\\sqrt{x}' });
    log('Escribiente: el .docx conserva reacciones, diagramas editables, citas con sus fuentes y el LaTeX de las ecuaciones');
    // sin los datos de Escribiente, la ecuación se lee del OMML de Word
    const z = await JSZip.loadAsync(fs.readFileSync(dx)); z.remove('docProps/escribiente.json');
    const plainDocx = path.join(out, 'solo-word.docx'); fs.writeFileSync(plainDocx, await z.generateAsync({ type: 'nodebuffer' }));
    await reopen(w, plainDocx); await wait(500);
    const tex = await w.evaluate(() => document.querySelector('#editor .eq')?.dataset.tex);
    assert.ok(/\\frac\{a\}\{b\}/.test(tex) && /\\sqrt\{x\}/.test(tex), `OMML → LaTeX: ${tex}`);
    log(`Escribiente: lee ecuaciones de Word (OMML → ${tex})`);
    if (hasSoffice) {
      // un .docx hecho por otro programa (LibreOffice escribe la ecuación como OMML)
      await w.evaluate(() => window.__shell.saveAsFormat('odt', 'ODF'));
      execFileSync('soffice', ['--headless', '--convert-to', 'docx', '--outdir', path.join(out, 'lo'), path.join(out, 'solo-word.odt')], { stdio: 'ignore', timeout: 120000 });
      await reopen(w, path.join(out, 'lo', 'solo-word.docx')); await wait(600);
      const t2 = await w.evaluate(() => document.querySelector('#editor .eq')?.dataset.tex || '');
      assert.ok(/frac/.test(t2) && /sqrt/.test(t2), `ecuación de LibreOffice: ${t2}`);
      log(`Escribiente: lee la ecuación de un .docx hecho en LibreOffice (${t2})`);
    }
    // escribir y borrar con el teclado, con y sin control de cambios; ortografía en ambos casos
    await w.evaluate(() => window.__shell.newDoc('blank')); await wait(500);
    await app.evaluate(({ BrowserWindow, Menu }) => { Menu.prototype.popup = function () { globalThis.__menu = this.items.map((i) => i.label); }; const bw = BrowserWindow.getAllWindows().find((b) => b.webContents.getURL().includes('/escribiente/')); bw.webContents.on('context-menu', (_e, p) => { globalThis.__cm = p.misspelledWord; }); });
    const dict = fs.existsSync(path.join(root, 'build', 'dict')) && fs.readdirSync(path.join(root, 'build', 'dict')).some((f) => f.startsWith('es'));
    const rightClickOn = async (word) => {
      await app.evaluate(() => { globalThis.__cm = null; });
      const b = await w.evaluate((wd) => { const tw = document.createTreeWalker(document.getElementById('editor'), NodeFilter.SHOW_TEXT); let n; while ((n = tw.nextNode())) { const i = n.nodeValue.indexOf(wd); if (i >= 0) { n.parentElement.scrollIntoView({ block: 'center' }); const r = document.createRange(); r.setStart(n, i + 2); r.setEnd(n, i + 3); const q = r.getBoundingClientRect(); return { x: q.x + q.width / 2, y: q.y + q.height / 2 }; } } return null; }, word);
      await w.mouse.click(b.x, b.y, { button: 'right' }); await wait(500);
      return app.evaluate(() => globalThis.__cm);
    };
    const edText = () => w.evaluate(() => { const c = document.getElementById('editor').cloneNode(true); c.querySelectorAll('del').forEach((d) => d.remove()); return c.innerText.replace(/\u00a0/g, ' ').trim(); });
    await w.click('#editor'); await w.keyboard.type('Hola mundo prueva ', { delay: 15 });
    for (let i = 0; i < 3; i++) await w.keyboard.press('Backspace');
    await w.keyboard.type('a final', { delay: 15 });
    assert.equal(await edText(), 'Hola mundo pruea final');
    if (dict) { await wait(1200); assert.equal(await rightClickOn('pruea'), 'pruea'); await w.keyboard.press('Escape'); }
    // con control de cambios: lo propio se borra de verdad y el cursor sigue en su lugar
    await w.evaluate(() => window.__shell.command('track')); await wait(200);
    assert.ok(/activado|on$/.test(await w.textContent('.sb-track')) && await w.evaluate(() => document.querySelector('.sb-track').classList.contains('on')));
    await w.click('#editor'); await w.keyboard.press('Control+End');
    await w.keyboard.type(' nuevoo textp', { delay: 15 });
    await w.keyboard.press('Backspace'); await w.keyboard.type('o y mas', { delay: 15 });
    assert.equal(await edText(), 'Hola mundo pruea final nuevoo texto y mas');
    assert.equal(await w.evaluate(() => document.querySelectorAll('#editor del').length), 0, 'lo propio no se marca como eliminado');
    assert.ok(await w.evaluate(() => [...document.querySelectorAll('#editor ins')].map((i) => i.textContent).join('').replace(/\u00a0/g, ' ') === ' nuevoo texto y mas'));
    if (dict) { await wait(1200); assert.equal(await rightClickOn('nuevoo'), 'nuevoo', 'ortografía dentro del texto insertado con control de cambios'); await w.keyboard.press('Escape'); }
    // borrar texto que ya estaba: queda tachado y el cursor queda antes de lo tachado
    await w.click('#editor'); await w.keyboard.press('Control+Home');
    for (let i = 0; i < 4; i++) await w.keyboard.press('ArrowRight');
    await w.keyboard.press('Backspace'); await w.keyboard.press('Backspace'); await w.keyboard.type('Z');
    const html = await w.evaluate(() => document.querySelector('#editor p').innerHTML.replace(/ data-[a-z]+="[^"]*"| title="[^"]*"/g, ''));
    assert.ok(/^Ho<ins>Z<\/ins><del>la<\/del>|^Ho<del>la<\/del>/.test(html) && /<ins>Z<\/ins>/.test(html), html);
    await w.evaluate(() => window.__shell.command('track')); await wait(200);
    assert.equal(await w.evaluate(() => document.querySelector('.sb-track').classList.contains('on')), false);
    log('Escribiente: retroceso y escritura normales; con control de cambios el cursor no se pierde y la ortografía sigue marcando');
    await closeApp('escribiente'); await wait(600);
  }
  // ============================================================ Impresiona: cámara y narración
  if (!only || only === 'impresiona') {
    const w = await openApp('impresiona');
    await w.evaluate(() => window.__shell.newDoc('negocio')); await wait(400);
    await w.evaluate(() => window.__impresiona.onCommand('insertCamera'));
    assert.ok(await w.evaluate(() => !!document.querySelector('.stage .obj.t-camera .cam-ph')));
    await w.evaluate(() => window.__impresiona.present(0)); await wait(1500);
    const cam = await w.evaluate(() => { const v = document.querySelector('.show .obj.t-camera video'); return { live: v?.dataset.live, w: v?.videoWidth || 0 }; });
    assert.ok(cam.live === '1' && cam.w > 0, `cámara: ${JSON.stringify(cam)}`);
    await w.evaluate(() => window.__show.end()); await wait(300);
    log('Impresiona: cámara en vivo dentro de la diapositiva');
    // grabar con narración
    await w.evaluate(() => { setTimeout(() => window.__impresiona.onCommand('recordShow')); });
    await w.waitForSelector('.modal'); await w.click('.modal .btn.primary');
    await w.waitForSelector('.show .rec-dot'); await wait(1300);
    await w.evaluate(() => window.__show.next()); await wait(1100);
    await w.evaluate(() => window.__show.end());
    await w.waitForSelector('.modal', { timeout: 15000 }); await w.click('.modal .btn.primary'); await wait(400);
    const nar = await w.evaluate(() => window.__impresiona.pres.slides.slice(0, 2).map((s) => ({ src: (s.narration?.src || '').slice(0, 16), dur: s.narration?.dur, adv: s.advanceAfter })));
    assert.ok(nar[0].src.startsWith('data:audio/') && nar[0].dur >= 1 && nar[1].dur >= 0.8 && nar[0].adv === nar[0].dur, JSON.stringify(nar));
    log('Impresiona: grabar presentación con narración por diapositiva');
    // video con la narración
    const vbytes = await w.evaluate(async () => { const b = await window.__impresiona.F3.exportVideo({ width: 640, secs: 1, useTimings: true, save: false }); const r = await window.oficina.save({ bytes: b, ask: true, ext: 'webm', suggestedName: 'narrado.webm' }); return b.length; });
    assert.ok(vbytes > 5000);
    try {
      const streams = execFileSync('ffprobe', ['-v', 'error', '-show_entries', 'stream=codec_type,codec_name', '-of', 'csv=p=0', path.join(out, 'narrado.webm')], { encoding: 'utf8' });
      assert.ok(/opus,audio/.test(streams) && /video/.test(streams), streams);
    } catch (e) { if (!/ENOENT/.test(String(e))) throw e; }
    log('Impresiona: el video exportado lleva la narración (pista de audio Opus)');
    // se conservan en el .pptx
    await w.evaluate(() => window.__shell.save());
    await reopen(w, path.join(out, `${UNTITLED}.pptx`)); await wait(800);
    const kept = await w.evaluate(() => { const p = window.__impresiona.pres; return { cam: p.slides.some((s) => s.objects.some((o) => o.type === 'camera')), nar: !!p.slides[0].narration?.src }; });
    assert.deepEqual(kept, { cam: true, nar: true });
    log('Impresiona: la cámara y la narración se conservan en el .pptx');
    // Diseño → Tamaño de diapositiva: menú con la marca y ajustes avanzados
    await w.evaluate(() => window.__shell.ribbon.show('design'));
    await w.click('.ribbon [data-cmd="size"]'); await wait(200);
    const menu = await w.evaluate(() => [...document.querySelectorAll('.dropdown .mi')].map((b) => [b.textContent.trim(), b.querySelector('path')?.getAttribute('d') === 'M4 12l5 5L20 6']));
    assert.equal(menu.length, 3, JSON.stringify(menu)); assert.ok(menu[1][1] && !menu[0][1], JSON.stringify(menu));
    await w.locator('.dropdown .mi').nth(2).click(); await wait(300);
    const before = await w.evaluate(() => { const p = window.__impresiona.pres; return { ...p.size, n: p.slides.length }; });
    await w.selectOption('.modal [name="preset"]', 'a4');
    await w.selectOption('.modal [name="orient"]', 'port');
    assert.equal(await w.inputValue('.modal [name="w"]'), '19.05'); assert.equal(await w.inputValue('.modal [name="h"]'), '27.517');
    await w.fill('.modal [name="first"]', '5');
    await w.click('.modal .btn.primary'); await wait(500);
    const after = await w.evaluate(() => { const p = window.__impresiona.pres; const out = p.slides.flatMap((s) => s.objects).filter((o) => o.x < -1 || o.y < -1 || o.x + o.w > p.size.w + 1 || o.y + o.h > p.size.h + 1).length; return { size: p.size, name: p.sizeName, first: p.firstNum, out }; });
    assert.deepEqual(after, { size: { w: 720, h: 1040 }, name: 'a4', first: 5, out: 0 }, JSON.stringify([before, after]));
    // ancho a mano → Personalizado
    await startDialog(w, 'slideSize'); await w.waitForSelector('.modal [name="w"]');
    await w.fill('.modal [name="w"]', '20'); await w.dispatchEvent('.modal [name="w"]', 'input');
    assert.equal(await w.inputValue('.modal [name="preset"]'), 'custom');
    await w.click('.modal .btn:not(.primary)'); await wait(200);
    // se guarda en el .pptx (tamaño y "numerar desde" también para PowerPoint)
    await w.evaluate(() => window.__shell.save()); await wait(500);
    const zp = await JSZip.loadAsync(fs.readFileSync(path.join(out, `${UNTITLED}.pptx`)));
    const px = await zp.file('ppt/presentation.xml').async('string');
    assert.ok(/firstSlideNum="5"/.test(px), px.slice(0, 400));
    const sz = /<p:sldSz cx="(\d+)" cy="(\d+)"/.exec(px); assert.ok(sz && Math.abs(Number(sz[1]) / 360000 - 19.05) < 0.05 && Math.abs(Number(sz[2]) / 360000 - 27.517) < 0.05, sz?.[0]);
    await reopen(w, path.join(out, `${UNTITLED}.pptx`)); await wait(800);
    assert.deepEqual(await w.evaluate(() => { const p = window.__impresiona.pres; return [p.size.w, p.size.h, p.firstNum, p.sizeName]; }), [720, 1040, 5, 'a4']);
    log('Impresiona: tamaño de diapositiva con ajustes avanzados (A4 vertical, numerar desde 5) y se guarda en el .pptx');
    await closeApp('impresiona'); await wait(600);
  }
  if (!only || only === 'calcula') {
    const w = await openApp('calcula');
    const top = () => w.locator('.modal').last();
    async function topForm(values) {
      for (const [name, v] of Object.entries(values)) { const el = top().locator(`[name="${name}"]`); if (await el.evaluate((e) => e.tagName) === 'SELECT') await el.selectOption(v); else await el.fill(String(v)); }
      await top().locator('.btn.primary').click(); await wait(250);
    }
    const colMenu = async (col, idx) => { await top().locator('.qe-grid th', { hasText: col }).first().click(); await w.locator('.dropdown .mi').nth(idx).click(); await wait(250); };
    const csv = path.join(out, 'ventas.csv');
    const write = (rows) => fs.writeFileSync(csv, rows.map((r) => r.join(',')).join('\n'));
    write([['Fecha', 'Region', 'Vendedor', 'Monto'], ['05/01/2026', 'Norte', ' ana ', '100'], ['06/01/2026', 'Sur', 'luis', '250'], ['07/01/2026', 'Norte', 'ana', '50'], ['07/01/2026', 'Norte', 'ana', '50'], ['08/01/2026', 'Centro', 'eva', '80'], ['09/01/2026', 'Sur', 'luis', '20']]);
    await w.evaluate((p) => { setTimeout(() => window.__calcula.Q.getData('csv', p)); }, csv);
    await w.waitForSelector('.qe-grid table', { timeout: 10000 });
    assert.equal(await top().locator('.qe-grid tbody tr').count(), 6);
    // quitar duplicados (botón), recortar Vendedor + Nombre propio (menú de columna), tipo fecha
    await top().locator('.qe-tools .btn').first().click(); await wait(250);
    assert.equal(await top().locator('.qe-grid tbody tr').count(), 5);
    const items = await (async () => { await top().locator('.qe-grid th', { hasText: 'Vendedor' }).click(); const l = await w.locator('.dropdown .mi').allTextContents(); await w.keyboard.press('Escape'); return l; })();
    const at = (es, en) => items.indexOf(EN ? en : es);
    assert.ok(at('Recortar espacios', 'Trim spaces') >= 0, items.join('|'));
    await colMenu('Vendedor', at('Recortar espacios', 'Trim spaces'));
    await colMenu('Vendedor', at('Nombre Propio', 'Proper Case'));
    await colMenu('Fecha', at('Tipo: fecha', 'Type: date'));
    // filtro Monto > 30
    await colMenu('Monto', at('Filtrar…', 'Filter…'));
    await topForm({ cmp: '>', value: '30' });
    assert.equal(await top().locator('.qe-grid tbody tr').count(), 4);
    // columna personalizada e IVA
    await top().locator('.qe-tools .btn').nth(4).click(); await wait(200);
    await topForm({ name: 'IVA', expr: '[Monto] * 0.16' });
    // agrupar por Region: suma de Monto
    await top().locator('.qe-tools .btn').nth(5).click(); await wait(200);
    await topForm({ by: 'Region', fn: 'sum', col: 'Monto', as: 'Total' });
    const steps = await top().locator('.qe-step').count();
    assert.equal(steps, 8, `pasos ${steps}`);
    // ordenar desc por Total y cargar
    await colMenu('Total', at('Ordenar descendente', 'Sort descending'));
    await top().locator('input.qe-name').fill('Ventas');
    await top().locator('.btn.primary').click(); await wait(600);
    const res = await w.evaluate(() => { const c = window.__calcula; const s = c.wb.sheet; return { name: s.name, rows: c.wb.hf.getSheetValues(s.id), q: c.wb.queries.map((q) => q.name) }; });
    assert.equal(res.name, 'Ventas'); assert.deepEqual(res.q, ['Ventas']);
    assert.deepEqual(res.rows, [['Region', 'Total'], ['Sur', 250], ['Norte', 150], ['Centro', 80]], JSON.stringify(res.rows));
    log('Calcula: obtener datos de CSV y transformarlos (duplicados, recortar, tipo, filtro, columna, agrupar, ordenar)');
    // el origen cambia → Actualizar todo
    write([['Fecha', 'Region', 'Vendedor', 'Monto'], ['05/01/2026', 'Norte', 'ana', '100'], ['06/01/2026', 'Oeste', 'zoe', '900']]);
    await w.evaluate(() => window.__shell.command('refreshAll')); await wait(800);
    const res2 = await w.evaluate(() => { const c = window.__calcula; const s = c.wb.sheets.find((x) => x.name === 'Ventas'); return c.wb.hf.getSheetValues(s.id); });
    assert.deepEqual(res2, [['Region', 'Total'], ['Oeste', 900], ['Norte', 100]], JSON.stringify(res2));
    log('Calcula: “Actualizar todo” vuelve a leer el origen y repite los pasos');
    // segunda consulta desde JSON + combinar
    const js = path.join(out, 'regiones.json');
    fs.writeFileSync(js, JSON.stringify({ data: [{ region: 'Norte', jefe: { nombre: 'Ana' } }, { region: 'Oeste', jefe: { nombre: 'Zoe' } }] }));
    const merged = await w.evaluate(async (p) => {
      const c = window.__calcula; const Q = c.Q;
      c.wb.queries.push({ id: 'qjs', name: 'Regiones', source: { kind: 'json', path: p, headers: true }, steps: [] });
      const v = c.wb.queries[0];
      v.steps.push({ op: 'merge', query: 'Regiones', left: 'Region', right: 'region', kind: 'left' });
      const tb = await Q.runQuery(v); return tb;
    }, js);
    assert.deepEqual(merged.cols, ['Region', 'Total', 'jefe.nombre']);
    assert.deepEqual(merged.rows.map((r) => r[2]), ['Zoe', 'Ana']);
    // expresión segura: no ejecuta código
    const bad = await w.evaluate(() => { try { window.__calcula.Q.compileExpr('alert(1)', ['A']); return 'ok'; } catch (e) { return 'err'; } });
    assert.equal(bad, 'err');
    log('Calcula: consulta desde JSON (aplanado) y combinar consultas');
    // persiste en .xlsx
    await w.evaluate(() => window.__shell.save()); await wait(500);
    await reopen(w, path.join(out, `${UNTITLED}.xlsx`)); await wait(800);
    const kept = await w.evaluate(() => window.__calcula.wb.queries.map((q) => [q.name, q.steps.length, q.source.kind]));
    assert.deepEqual(kept, [['Ventas', 9, 'csv'], ['Regiones', 0, 'json']]);
    write([['Fecha', 'Region', 'Vendedor', 'Monto'], ['05/01/2026', 'Sur', 'ana', '70']]);
    await w.evaluate(() => window.__shell.command('refreshAll')); await wait(800);
    const res3 = await w.evaluate(() => { const c = window.__calcula; const s = c.wb.sheets.find((x) => x.name === 'Ventas'); return c.wb.hf.getSheetValues(s.id); });
    assert.deepEqual(res3, [['Region', 'Total', 'jefe.nombre'], ['Sur', 70]], JSON.stringify(res3));
    assert.equal(await w.evaluate(() => window.__calcula.wb.sheets.length), 3);
    log('Calcula: las consultas se guardan en el .xlsx y se actualizan al reabrir');
    await closeApp('calcula'); await wait(600);
  }
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  log('Sin errores de JavaScript en las ventanas');
  console.log('\nTodas las pruebas de la fase 4 pasaron.');
} catch (err) {
  console.error('✘', err);
  for (const w of app.windows()) { try { await w.screenshot({ path: path.join(out, `error-${Date.now()}.png`) }); } catch { /* */ } }
  process.exitCode = 1;
} finally {
  await app.close();
}
