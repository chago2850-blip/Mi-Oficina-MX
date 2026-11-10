// Pruebas de las funciones de la fase 3.
// Linux sin pantalla:  xvfb-run -a node test/e2e-fase3.mjs [escribiente|calcula|impresiona|suite]
import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase3');
fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(out, { recursive: true });
const only = process.argv[2];
const EN = process.env.OFICINA_LANG === 'en';
const UNTITLED = EN ? 'Untitled' : 'Sin título';
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const log = (m) => console.log(`✔ ${m}`);
const exe = process.env.OFICINA_EXE;
const app = await electron.launch({ ...(exe ? { executablePath: exe, args: [`--user-data-dir=${path.join(out, 'perfil')}`, '--no-sandbox'] } : { args: [root, `--user-data-dir=${path.join(out, 'perfil')}`, '--no-sandbox'] }), cwd: root, env: { ...process.env, OFICINA_TEST_DIR: out, OFICINA_LANG: process.env.OFICINA_LANG || 'es' } });
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

try {
  // ============================================================ Escribiente
  if (!only || only === 'escribiente') {
    const w = await openApp('escribiente');
    await w.evaluate(() => window.__shell.newDoc('blank'));
    await w.evaluate(() => { const { editor } = window.__escribiente; editor.innerHTML = '<h1>Informe</h1><p id="p1">La fórmula es </p><p id="p2">Según el estudio</p>'; editor.dispatchEvent(new Event('input')); });
    // ecuación
    await w.evaluate(() => { const p = document.querySelector('#p1'); const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); window.__escribiente.editor.focus(); });
    await startDialog(w, 'equation');
    await w.waitForSelector('.modal .eq-src');
    await w.fill('.modal .eq-src', 'x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}');
    await wait(200);
    assert.ok(await w.evaluate(() => !!document.querySelector('.modal .eq-preview math mfrac')), 'vista previa de la ecuación');
    await w.click('.modal .btn.primary'); await wait(300);
    assert.ok(await w.evaluate(() => { const e = document.querySelector('#editor .eq math'); return !!(e && e.querySelector('mfrac') && e.textContent.startsWith('x=')); }));
    log('Escribiente: inserta una ecuación con vista previa (LaTeX → MathML)');
    // cita y bibliografía
    await w.evaluate(() => { const p = document.querySelector('#p2'); const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); window.__escribiente.editor.focus(); window.__escribiente.state.sources.push({ id: 's1', type: 'libro', author: 'García Márquez, Gabriel', title: 'Cien años de soledad', year: '1967', publisher: 'Sudamericana', url: '' }, { id: 's2', type: 'web', author: 'Pérez, Ana; López, Luis', title: 'El agua en México', year: '2024', publisher: 'Conagua', url: 'https://example.org' }); });
    await w.evaluate(() => { const { F3 } = window.__escribiente; });
    await w.evaluate(() => { const s = window.__escribiente.state.sources; window.__escribiente.insertCite(s[0]); window.__escribiente.insertCite(s[1]); });
    const cites = await w.evaluate(() => [...document.querySelectorAll('#editor .cite')].map((c) => c.textContent));
    assert.deepEqual(cites, ['(García Márquez, 1967)', EN ? '(Pérez and López, 2024)' : '(Pérez y López, 2024)'].map((x) => x));
    await w.evaluate(() => { const ed = window.__escribiente.editor; const p = document.createElement('p'); p.innerHTML = '<br>'; ed.append(p); const r = document.createRange(); r.setStart(p, 0); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); window.__escribiente.saveSelection(); window.__escribiente.onCommand('bibliography'); });
    await wait(300);
    const bib = await w.evaluate(() => [...document.querySelectorAll('#editor [data-biblio]')].map((x) => x.textContent));
    assert.equal(bib.length, 3); assert.ok(bib[1].startsWith('García Márquez, G. (1967). Cien años de soledad'), bib[1]);
    log('Escribiente: citas (APA) y bibliografía ordenada');
    // diagrama
    await w.evaluate(() => { const ed = window.__escribiente.editor; const p = document.createElement('p'); p.innerHTML = '<br>'; ed.append(p); const r = document.createRange(); r.setStart(p, 0); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); window.__escribiente.saveSelection(); });
    await startDialog(w, 'smartart');
    await fillForm(w, { type: 'ciclo', text: 'Planear\nHacer\nVerificar\nActuar' });
    assert.ok(await w.evaluate(() => { const i = document.querySelector('#editor img.smartart'); return i && i.src.startsWith('data:image/svg+xml') && atob(i.src.split(',')[1]).includes('Verificar'); }));
    log('Escribiente: diagrama SmartArt (ciclo) editable');
    // docx con ecuación nativa
    await w.evaluate(() => window.__shell.save());
    const dx = await JSZip.loadAsync(fs.readFileSync(path.join(out, `${UNTITLED}.docx`)));
    const dxml = await dx.file('word/document.xml').async('string');
    assert.ok(/<m:oMath><m:r><m:t>x<\/m:t>/.test(dxml) && /<m:f>/.test(dxml) && /<m:rad>/.test(dxml), 'OMML en el docx');
    assert.ok(dxml.includes('García Márquez, G. (1967)'), 'bibliografía en el docx');
    assert.ok(Object.keys(dx.files).some((f) => f.startsWith('word/media/')), 'diagrama como imagen');
    log('Escribiente: el .docx lleva la ecuación como ecuación de Word (OMML), las citas y el diagrama');
    // odt con fórmula
    await w.evaluate(() => window.__shell.saveAsFormat('odt', 'ODF'));
    const oz = await JSZip.loadAsync(fs.readFileSync(path.join(out, `${UNTITLED}.odt`)));
    assert.ok(oz.file('Object 1/content.xml'), 'objeto de fórmula en el odt');
    await reopen(w, path.join(out, `${UNTITLED}.odt`)); await wait(400);
    assert.ok(await w.evaluate(() => { const e = document.querySelector('#editor .eq'); return !!(e && e.querySelector('mfrac') && e.dataset.tex.includes('frac')); }));
    log('Escribiente: la ecuación se guarda en .odt como fórmula y se vuelve a editar');
    // markdown
    await w.evaluate(() => window.__shell.saveAsFormat('md', 'Markdown'));
    const md = fs.readFileSync(path.join(out, `${UNTITLED}.md`), 'utf8');
    assert.ok(md.startsWith('# Informe') && md.includes('$x = \\frac') && md.includes('(García Márquez, 1967)'), md);
    log('Escribiente: exporta a Markdown');
    // comparar documentos
    await w.evaluate(() => window.__shell.newDoc('blank'));
    await w.evaluate(() => { const { editor } = window.__escribiente; editor.innerHTML = '<h1>Contrato</h1><p>El pago será de 100 pesos al mes.</p><p>Vigencia de un año.</p><p>Párrafo que se quita.</p>'; });
    fs.writeFileSync(path.join(out, 'revisado.txt'), 'Contrato\nEl pago será de 120 pesos al mes.\nVigencia de un año.\nCláusula nueva.\n');
    const cmp = await w.evaluate(async (f) => {
      const { compareHtml } = window.__escribiente.F3;
      const rev = document.createElement('div'); rev.innerHTML = (await window.oficina.readPath(f)) && '<p>Contrato</p><p>El pago será de 120 pesos al mes.</p><p>Vigencia de un año.</p><p>Cláusula nueva.</p>';
      const r = compareHtml(window.__escribiente.editor, rev, 'Prueba');
      window.__escribiente.setContent(r.html);
      const ed = window.__escribiente.editor;
      return { ins: [...ed.querySelectorAll('ins')].map((x) => x.textContent), del: [...ed.querySelectorAll('del')].map((x) => x.textContent) };
    }, path.join(out, 'revisado.txt'));
    assert.ok(cmp.ins.includes('120') && cmp.del.includes('100') && cmp.ins.some((x) => x.includes('Cláusula nueva')) && cmp.del.some((x) => x.includes('Párrafo que se quita')), JSON.stringify(cmp));
    log('Escribiente: comparar documentos marca inserciones y eliminaciones palabra por palabra');
    // etiquetas y sobres
    await startDialog(w, 'labels');
    await fillForm(w, { preset: '5160', text: 'Ana López\nCalle 1\nAcapulco' });
    if (await w.$('.modal')) { await w.click('.modal .btn.primary'); await wait(300); }
    const lab = await w.evaluate(() => ({ cells: document.querySelectorAll('#editor table.labels td').length, txt: document.querySelector('#editor table.labels td')?.innerText }));
    assert.equal(lab.cells, 30); assert.ok(lab.txt.includes('Acapulco'));
    await startDialog(w, 'envelope');
    await fillForm(w, { to: 'Lic. Bruno Díaz\nAv. Costera 100', from: 'Mi empresa', size: 'dl' });
    if (await w.$('.modal')) { await w.click('.modal .btn.primary'); await wait(300); }
    const env = await w.evaluate(() => ({ w: window.__escribiente.state.page.w, h: window.__escribiente.state.page.h, txt: window.__escribiente.editor.innerText }));
    assert.ok(Math.abs(env.w - 831) < 3 && Math.abs(env.h - 416) < 3 && env.txt.includes('Bruno Díaz'), JSON.stringify(env));
    log('Escribiente: etiquetas (30 por hoja) y sobre DL');
    // modos
    await w.evaluate(() => window.__escribiente.onCommand('focusMode')); await wait(200);
    assert.ok(await w.evaluate(() => getComputedStyle(document.querySelector('nav.ribbon')).display === 'none' && !!document.querySelector('.mode-exit')));
    await w.keyboard.press('Escape'); await wait(200);
    await w.evaluate(() => window.__escribiente.onCommand('readingMode')); await wait(200);
    assert.equal(await w.evaluate(() => window.__escribiente.editor.contentEditable), 'false');
    await w.screenshot({ path: path.join(out, 'escribiente-lectura.png') });
    await w.keyboard.press('Escape'); await wait(200);
    assert.equal(await w.evaluate(() => window.__escribiente.editor.contentEditable), 'true');
    log('Escribiente: modo enfoque y modo lectura (Esc para salir)');
    // .doc binario antiguo
    const docFixture = path.join(root, 'test', 'fixtures', 'antiguo.doc');
    if (fs.existsSync(docFixture)) {
      fs.copyFileSync(docFixture, path.join(out, 'antiguo.doc'));
      await reopen(w, path.join(out, 'antiguo.doc')); await wait(500);
      assert.ok((await w.evaluate(() => window.__escribiente.editor.innerText)).includes('Mi Oficina'), 'texto del .doc');
      assert.equal(await w.evaluate(() => window.__shell.doc.ext), 'docx');
      log('Escribiente: abre un .doc de Word 97-2003 (modo compatibilidad)');
    }
    // reacciones en comentarios
    await w.evaluate(() => window.__shell.newDoc('blank'));
    await w.evaluate(() => { const { editor, state } = window.__escribiente; editor.innerHTML = '<p><span class="cm" data-cid="c1">Hola</span> mundo</p>'; state.comments = [{ id: 'c1', author: 'Ana', date: new Date().toISOString(), text: 'Revisar', replies: [], resolved: false }]; window.__escribiente.F.toggleSide('comments'); });
    await wait(300);
    await w.click('.cmcard .react.add'); await wait(150);
    await w.click('.dropdown button >> nth=0').catch(async () => { await w.evaluate(() => document.querySelector('.dropdown [role="menuitem"], .dropdown button, .dropdown .item')?.click()); });
    await wait(200);
    const re = await w.evaluate(() => window.__escribiente.state.comments[0].reactions);
    assert.ok(re && re['👍']?.length === 1, JSON.stringify(re));
    log('Escribiente: reacciones en comentarios');
    await closeApp('escribiente'); await wait(600);
  }
  // ============================================================ Calcula Datos
  if (!only || only === 'calcula') {
    const w = await openApp('calcula');
    await w.evaluate(() => window.__shell.newDoc('blank'));
    const set = (cells) => w.evaluate((cs) => { const { wb } = window.__calcula; wb.action('datos', () => cs.forEach(([r, c, v]) => wb.setInput(r, c, String(v))), { batch: true }); window.__calcula.refresh(); }, cells);
    const val = (r, c) => w.evaluate(([a, b]) => { const v = window.__calcula.wb.value(a, b); return v && typeof v === 'object' ? (v.value ?? String(v)) : v; }, [r, c]);
    // LAMBDA, MAP, REDUCE, BYROW, LET, TOMAR, APILARV, TEXTOANTES
    await set([[0, 0, 1], [0, 1, 2], [0, 2, 3], [1, 0, 4], [1, 1, 5], [1, 2, 6],
      [3, 0, EN ? '=MAP(A1:C1, LAMBDA(x, x*10))' : '=MAP(A1:C1, LAMBDA(x, x*10))'],
      [4, 0, '=REDUCE(0, A1:C2, LAMBDA(a, v, a+v))'],
      [5, 0, EN ? '=BYROW(A1:C2, LAMBDA(r, SUM(r)))' : '=BYROW(A1:C2, LAMBDA(r, SUMA(r)))'],
      [7, 0, '=LET(n, 5, m, n*2, n+m)'],
      [8, 0, EN ? '=TAKE(A1:C2, 1, 2)' : '=TOMAR(A1:C2, 1, 2)'],
      [9, 0, EN ? '=VSTACK(A1:B1, A2:B2)' : '=APILARV(A1:B1, A2:B2)'],
      [11, 0, EN ? '=TEXTBEFORE("hola-mundo", "-")' : '=TEXTOANTES("hola-mundo", "-")'],
      [12, 0, '=SCAN(0, A1:C1, LAMBDA(a, v, a+v))'],
    ]);
    assert.deepEqual([await val(3, 0), await val(3, 1), await val(3, 2)], [10, 20, 30]);
    assert.equal(await val(4, 0), 21); assert.deepEqual([await val(5, 0), await val(6, 0)], [6, 15]);
    assert.equal(await val(7, 0), 15); assert.deepEqual([await val(8, 0), await val(8, 1)], [1, 2]);
    assert.deepEqual([await val(9, 0), await val(10, 1)], [1, 5]); assert.equal(await val(11, 0), 'hola'); assert.equal(await val(12, 2), 6);
    if (!EN) assert.equal(await w.evaluate(() => window.__calcula.wb.editText(8, 0)), '=TOMAR(A1:C2, 1, 2)');
    log('Calcula Datos: LAMBDA, MAP, REDUCE, SCAN, BYROW, LET, TOMAR, APILARV y TEXTOANTES');
    // buscar objetivo: préstamo — ¿qué tasa da un pago de 1000?
    await set([[14, 0, 'Monto'], [14, 1, 20000], [15, 0, 'Pago'], [15, 1, '=B15*0.05+B17'], [16, 0, 'Cargo'], [16, 1, 0]]);
    await startDialog(w, 'goalSeek');
    await fillForm(w, { target: 'B16', want: '1500', change: 'B17' });
    assert.ok(Math.abs((await val(16, 1)) - 500) < 1e-6, `buscar objetivo: ${await val(16, 1)}`);
    await w.evaluate(() => window.__calcula.onCommand('undo')); await wait(150);
    assert.equal(await val(16, 1), 0, 'deshacer buscar objetivo');
    log('Calcula Datos: buscar objetivo (y se deshace en un paso)');
    // Solver: maximizar 3x + 2y con x + y <= 4, x + 3y <= 6, x,y >= 0  (óptimo x=4, y=0 -> 12)
    await set([[19, 0, 'x'], [19, 1, 0], [20, 0, 'y'], [20, 1, 0], [21, 0, 'z'], [21, 1, '=3*B20+2*B21'], [22, 1, '=B20+B21'], [23, 1, '=B20+3*B21']]);
    await startDialog(w, 'solver');
    const tS = Date.now();
    await fillForm(w, { target: 'B22', goal: 'max', vars: 'B20:B21', cons: 'B23 <= 4\nB24 <= 6\nB20:B21 >= 0' });
    await w.waitForSelector('.modal', { timeout: 60000 }); console.log(`  (Solver: ${Date.now() - tS} ms)`); await w.click('.modal .btn.primary'); await wait(300);
    assert.ok(Math.abs((await val(21, 1)) - 12) < 1e-3, `solver: ${await val(21, 1)}`);
    const more = await w.evaluate(() => {
      const { wb, F3 } = window.__calcula; const A = (r, c) => ({ r, c });
      wb.action('m', () => { wb.setInput(30, 1, '0'); wb.setInput(31, 1, '0'); wb.setInput(32, 1, '=5*B31+4*B32'); wb.setInput(33, 1, '=6*B31+4*B32'); wb.setInput(34, 1, '=B31+2*B32'); wb.setInput(36, 1, '0'); wb.setInput(37, 1, '0'); wb.setInput(38, 1, '=(B37-3)^2+(B38+1)^2'); }, { batch: true });
      const ip = F3.solverCore({ target: A(32, 1), goal: 'max', vars: [A(30, 1), A(31, 1)], constraints: [{ kind: 'cmp', r: 33, c: 1, op: '<=', rhs: 24 }, { kind: 'cmp', r: 34, c: 1, op: '<=', rhs: 6 }, { kind: 'cmp', r: 30, c: 1, op: '>=', rhs: 0 }, { kind: 'cmp', r: 31, c: 1, op: '>=', rhs: 0 }, { kind: 'int', r: 30, c: 1 }, { kind: 'int', r: 31, c: 1 }] });
      const nl = F3.solverCore({ target: A(38, 1), goal: 'min', vars: [A(36, 1), A(37, 1)], constraints: [] });
      return { ip: [ip.obj, ...ip.xs].map((v) => v + 0), nl: nl.xs.map((x) => Math.round(x * 1000) / 1000 + 0) };
    });
    assert.deepEqual(more.ip, [20, 4, 0], `entero: ${more.ip}`); assert.deepEqual(more.nl, [3, -1], `no lineal: ${more.nl}`);
    log('Calcula Datos: Solver maximiza con restricciones (Simplex, enteros y no lineal)');
    // minigráficos
    await w.evaluate(() => { const { wb, grid } = window.__calcula; grid.select(0, 0); grid.select(1, 2, true); });
    await startDialog(w, 'sparkline');
    await fillForm(w, { data: 'A1:C2', loc: 'E1:E2', type: 'line' });
    assert.equal(await w.evaluate(() => window.__calcula.wb.sheet.x3.sparklines.length), 2);
    // agrupar filas 15-18 y plegar
    await w.evaluate(() => { const { grid } = window.__calcula; grid.select(14, 0); grid.select(17, 0, true); window.__calcula.onCommand('group'); window.__calcula.onCommand('hideDetail'); });
    assert.ok(await w.evaluate(() => [14, 15, 16, 17].every((r) => window.__calcula.wb.sheet.hiddenRows.has(r))));
    await w.screenshot({ path: path.join(out, 'calcula-fase3.png') });
    await w.evaluate(() => window.__calcula.onCommand('showDetail'));
    assert.ok(await w.evaluate(() => !window.__calcula.wb.sheet.hiddenRows.has(15)));
    log('Calcula Datos: minigráficos y agrupar/plegar filas');
    // IMAGEN
    await set([[25, 0, `=IMAGE("data:image/svg+xml;base64,${Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10" fill="red"/></svg>').toString('base64')}")`]]);
    assert.ok(String(await val(25, 0)).startsWith('data:image/svg+xml'));
    if (!EN) assert.ok((await w.evaluate(() => window.__calcula.wb.editText(25, 0))).startsWith('=IMAGEN('));
    log('Calcula Datos: función IMAGEN');
    // gráfico con rango de matriz dinámica (A4#)
    await w.evaluate(() => { window.__calcula.wb.action('g', () => window.__calcula.wb.sheet.charts.push({ id: 'spill', type: 'column', range: 'A4#', title: '', x: 600, y: 10, w: 300, h: 200 }), { hf: false }); });
    const sp = await w.evaluate(() => { const g = window.__calcula.F3.resolveRange('A4#'); return [g.r1, g.c1, g.r2, g.c2]; });
    assert.deepEqual(sp, [3, 0, 3, 2]);
    await set([[0, 3, 7]]);
    log('Calcula Datos: los gráficos siguen matrices dinámicas (A4#)');
    // segmentación
    await w.evaluate(() => window.__shell.newDoc('blank'));
    await set([[0, 0, 'Zona'], [0, 1, 'Ventas'], [1, 0, 'Norte'], [1, 1, 10], [2, 0, 'Sur'], [2, 1, 20], [3, 0, 'Norte'], [3, 1, 30], [4, 0, 'Centro'], [4, 1, 40]]);
    await w.evaluate(() => window.__calcula.grid.select(0, 0));
    await startDialog(w, 'slicer');
    await fillForm(w, { col: '0' });
    await wait(200);
    await w.click('.slicer .sl-item:has-text("Norte")'); await wait(200);
    const hid = await w.evaluate(() => [...window.__calcula.wb.sheet.hiddenRows].sort());
    assert.deepEqual(hid, [2, 4]);
    log('Calcula Datos: segmentación de datos filtra la tabla');
    // se guarda en xlsx: esquema nativo + datos de fase 3
    await w.evaluate(() => { const { grid } = window.__calcula; grid.select(1, 0); grid.select(3, 0, true); window.__calcula.onCommand('group'); });
    await w.evaluate(() => window.__shell.save());
    const xp = path.join(out, `${UNTITLED}.xlsx`);
    const xz = await JSZip.loadAsync(fs.readFileSync(xp));
    assert.ok(/outlineLevel="1"/.test(await xz.file('xl/worksheets/sheet1.xml').async('string')), 'esquema en el xlsx');
    await reopen(w, xp); await wait(400);
    const back = await w.evaluate(() => ({ sl: window.__calcula.wb.sheet.x3.slicers.length, g: window.__calcula.wb.sheet.x3.rowGroups.length }));
    assert.deepEqual(back, { sl: 1, g: 1 });
    log('Calcula Datos: el .xlsx conserva el esquema y la segmentación');
    await closeApp('calcula'); await wait(600);
  }
  // ============================================================ Impresiona
  if (!only || only === 'impresiona') {
    const w = await openApp('impresiona');
    await w.evaluate(() => window.__shell.newDoc('negocio')); await wait(400);
    // exportar a video
    const vid = await w.evaluate(async () => { const b = await window.__impresiona.F3.exportVideo({ width: 640, secs: 1, save: false }); return b ? Array.from(b.slice(0, 4)).concat([b.length]) : null; });
    assert.ok(vid && vid[0] === 0x1a && vid[1] === 0x45 && vid[2] === 0xdf && vid[3] === 0xa3 && vid[4] > 5000, `webm: ${vid}`);
    await w.evaluate(async () => { const b = await window.__impresiona.F3.exportVideo({ width: 640, secs: 1, save: false }); const r = await window.oficina.save({ bytes: b, ask: true, ext: 'webm', suggestedName: 'presentacion.webm' }); window.__vidPath = r.path; });
    const webm = path.join(out, 'presentacion.webm');
    try { const d = Number(execFileSync('ffprobe', ['-v', 'error', '-show_entries', 'format=duration', '-of', 'csv=p=0', webm], { encoding: 'utf8' })); assert.ok(d > 3 && d < 6, `duración ${d}`); } catch (e) { if (!/ENOENT/.test(String(e))) throw e; }
    log('Impresiona: exporta la presentación a video WebM con transiciones');
    // video con subtítulos
    const media = await w.evaluate(async (file) => {
      const f = await window.oficina.readPath(file); let s = ''; for (let i = 0; i < f.bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, f.bytes.subarray(i, i + 0x8000));
      const { F3 } = window.__impresiona; const o = F3.addMedia('video', `data:video/webm;base64,${btoa(s)}`, 'presentacion.webm');
      o.captions = [{ lang: 'es', label: 'Español', vtt: F3.simpleToVtt('0:00 Hola a todos\n0:02 Este es el plan') }];
      window.__impresiona.renderAll();
      return { id: o.id, vtt: o.captions[0].vtt };
    }, webm);
    assert.ok(media.vtt.startsWith('WEBVTT') && media.vtt.includes('00:00:02.000 --> '), media.vtt);
    assert.ok(await w.evaluate(() => !!document.querySelector('.stage .obj.t-media video')));
    log('Impresiona: inserta video con subtítulos (WebVTT)');
    // trayectoria y transformación
    await w.evaluate((mid) => {
      const I = window.__impresiona; const p = I.pres; const s0 = p.slides[I.cur];
      const title = s0.objects.find((o) => o.type === 'text');
      title.anim = { type: 'pathRight', start: 'click', dur: 0.4, delay: 0, order: 1, dist: 150 };
      I.onCommand('duplicateSlide');
      const s1 = p.slides[I.cur]; const t1 = s1.objects.find((o) => o.mk === title.id); t1.x += 120; t1.y += 60; t1.w *= 0.8; delete t1.anim;
      s1.transition = { type: 'morph', dur: 0.8 };
      s1.objects = s1.objects.filter((o) => o.type !== 'media');
      I.renderAll();
    }, media.id);
    await w.evaluate(() => window.__impresiona.present(window.__impresiona.cur - 1)); await wait(500);
    await w.evaluate(() => window.__show.next());
    // espera a que termine la animación (en el CI el equipo es más lento)
    const tr = await w.evaluate(async () => {
      const t = window.__impresiona.pres.slides.flatMap((x) => x.objects).find((o) => o.anim?.type === 'pathRight');
      const el = document.querySelector(`.show [data-id="${t.id}"]`);
      await Promise.all(el.getAnimations().map((a) => a.finished.catch(() => {})));
      for (let k = 0; k < 40 && !/150, 0\)$/.test(getComputedStyle(el).transform); k++) await new Promise((r) => setTimeout(r, 100));
      return getComputedStyle(el).transform;
    });
    assert.ok(/matrix\(1, 0, 0, 1, 150, 0\)/.test(tr), `trayectoria: ${tr}`);
    const morphing = await w.evaluate(() => { window.__show.next(); return [...document.querySelectorAll('.show .show-layer:last-child .obj')].some((el) => el.getAnimations().length > 0); });
    assert.ok(morphing, 'la transición Transformación anima los objetos');
    await wait(900);
    await w.evaluate(() => window.__show.end()); await wait(300);
    log('Impresiona: trayectoria de animación y transición Transformación (Morph)');
    // ensayar intervalos y avance automático
    await w.evaluate(() => window.__impresiona.onCommand('rehearse')); await wait(1300);
    await w.evaluate(() => window.__show.next()); await wait(700);
    await w.evaluate(() => window.__show.end()); await wait(400);
    await w.waitForSelector('.modal'); await w.click('.modal .btn.primary'); await wait(300);
    const timings = await w.evaluate(() => window.__impresiona.pres.slides.map((s) => s.advanceAfter || 0));
    assert.ok(timings[0] >= 1 && timings[0] < 3, `intervalos: ${timings}`);
    await w.evaluate(() => { const p = window.__impresiona.pres; p.slides.forEach((s) => { s.advanceAfter = 0.6; }); window.__impresiona.present(0); }); await wait(1600);
    const auto = await w.evaluate(() => window.__show.engine.i);
    assert.ok(auto >= 1, `avance automático: ${auto}`);
    await w.evaluate(() => window.__show.end()); await wait(300);
    log('Impresiona: ensayar intervalos y avance automático');
    // pptx con video
    await w.evaluate(() => window.__shell.save());
    const pz = await JSZip.loadAsync(fs.readFileSync(path.join(out, `${UNTITLED}.pptx`)));
    assert.ok(Object.keys(pz.files).some((f) => /^ppt\/media\/.*\.webm$/.test(f)), 'video dentro del pptx');
    await reopen(w, path.join(out, `${UNTITLED}.pptx`)); await wait(800);
    assert.ok(await w.evaluate(() => window.__impresiona.pres.slides.some((s) => s.objects.some((o) => o.type === 'media' && o.src.startsWith('data:video')))));
    const kept = await w.evaluate(() => { const p = window.__impresiona.pres; const all = p.slides.flatMap((s) => s.objects); return { cap: all.some((o) => o.captions?.[0]?.vtt?.includes('Este es el plan')), path: all.some((o) => o.anim?.type === 'pathRight' && o.anim.dist === 150), morph: p.slides.some((s) => s.transition?.type === 'morph'), timing: p.slides.every((s) => s.advanceAfter === 0.6), useTimings: p.useTimings }; });
    assert.deepEqual(kept, { cap: true, path: true, morph: true, timing: true, useTimings: true });
    log('Impresiona: el video se guarda en el .pptx y al abrirlo se recuperan subtítulos, trayectorias, Transformación e intervalos');
    await closeApp('impresiona'); await wait(600);
  }
  // ============================================================ Suite
  if (!only || only === 'suite') {
    const w = await openApp('escribiente');
    await w.evaluate(() => window.__shell.newDoc('blank'));
    await w.evaluate(() => { window.__escribiente.editor.innerHTML = '<p>Documento confidencial de prueba.</p>'; });
    // barra de acceso rápido
    await w.click('.qat .qa-more'); await wait(200);
    await w.evaluate(() => { const it = [...document.querySelectorAll('.dropdown button, .dropdown .item, .dropdown [role="menuitem"]')].find((b) => /Imprimir|Print/.test(b.textContent)); it?.click(); });
    await wait(300);
    assert.ok(await w.evaluate(() => !!document.querySelector('.qat [data-qat="print"]')), 'Imprimir en la barra de acceso rápido');
    const qat = await w.evaluate(() => window.oficina.prefs().then((p) => p.qat?.escribiente));
    assert.deepEqual(qat, ['save', 'undo', 'redo', 'print']);
    log('Suite: barra de acceso rápido personalizable (se guarda en Preferencias)');
    // cifrar con contraseña
    await startDialog(w, 'protectFile');
    await fillForm(w, { p: 'Secreta123', p2: 'Secreta123' });
    await w.evaluate(() => window.__shell.save());
    const enc = path.join(out, `${UNTITLED}.docx`);
    assert.equal(fs.readFileSync(enc).subarray(0, 4).toString('hex'), 'd0cf11e0', 'archivo cifrado (contenedor OLE)');
    let pyOk = 'True';
    try { pyOk = py(`import msoffcrypto,io\nf=msoffcrypto.OfficeFile(open(${JSON.stringify(enc)},'rb'))\nf.load_key(password='Secreta123')\no=io.BytesIO()\nf.decrypt(o,verify_integrity=True)\nimport zipfile\nz=zipfile.ZipFile(o)\nprint('confidencial' in z.read('word/document.xml').decode())`); } catch (e) { if (!/No module named/.test(String(e.stderr || e))) throw e; console.log('  (msoffcrypto no está instalado: se omite la verificación externa)'); }
    assert.equal(pyOk, 'True', 'otra implementación (msoffcrypto) lo descifra');
    // abrir: pide la contraseña
    const opening = w.evaluate(async (f) => { const d = await window.oficina.readPath(f); await window.__shell.loadFile(d); }, enc);
    await fillForm(w, { p: 'mala' });
    await w.waitForSelector('.modal'); assert.ok(/incorrecta|Incorrect/.test(await w.textContent('.modal')));
    await fillForm(w, { p: 'Secreta123' });
    await opening; await wait(300);
    assert.ok((await w.evaluate(() => window.__escribiente.editor.innerText)).includes('confidencial'));
    assert.equal(await w.evaluate(() => window.__shell.doc.password), 'Secreta123', 'se conserva la contraseña al volver a guardar');
    log('Suite: cifrar con contraseña (AES-256, compatible con Office) y abrir pidiendo la contraseña');
    await closeApp('escribiente'); await wait(600);
  }
  console.log('\nTodas las pruebas de la fase 3 pasaron.');
} catch (err) {
  console.error('✘', err);
  for (const w of app.windows()) { try { await w.screenshot({ path: path.join(out, `error-${Date.now()}.png`) }); } catch { /* */ } }
  process.exitCode = 1;
} finally {
  await app.close();
}
