// Pruebas de las funciones de la fase 2 (Escribiente, Calcula Datos, Impresiona).
// Linux sin pantalla:  xvfb-run -a node test/e2e-fase2.mjs [escribiente|calcula|impresiona]
import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase2');
fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(out, { recursive: true });
const only = process.argv[2];
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const log = (m) => console.log(`✔ ${m}`);
const exe = process.env.OFICINA_EXE;
const app = await electron.launch({ ...(exe ? { executablePath: exe, args: [`--user-data-dir=${path.join(out, 'perfil')}`, '--no-sandbox'] } : { args: [root, `--user-data-dir=${path.join(out, 'perfil')}`, '--no-sandbox'] }), cwd: root, env: { ...process.env, OFICINA_TEST_DIR: out } });
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

try {
  // ============================================================ Escribiente
  if (!only || only === 'escribiente') {
    const w = await openApp('escribiente');
    await w.evaluate(() => window.__shell.newDoc('blank'));
    await w.evaluate(() => {
      const { editor } = window.__escribiente;
      editor.innerHTML = '<h1>Introducción</h1><p id="p1">Texto original para comentar.</p><h2>Detalle</h2><p id="p2">Párrafo con nota.</p><p id="p3">Columna uno. Columna dos. Columna tres.</p><p>Estimado <b>cliente</b>:</p>';
      editor.dispatchEvent(new Event('input'));
    });
    // comentario
    await w.evaluate(() => { const t = document.querySelector('#p1').firstChild; const r = document.createRange(); r.setStart(t, 0); r.setEnd(t, 5); const s = getSelection(); s.removeAllRanges(); s.addRange(r); window.__escribiente.editor.focus(); });
    await w.evaluate(() => { const { F, state } = window.__escribiente; document.getElementById('editor').dispatchEvent(new Event('selectionchange')); });
    const cmP = w.evaluate(() => window.__escribiente.onCommand('comment'));
    await wait(300); await w.fill('.modal textarea', 'Revisar esta palabra'); await w.click('.modal .btn.primary'); await cmP;
    assert.equal(await w.evaluate(() => document.querySelectorAll('#editor .cm').length), 1);
    // respuesta
    await w.fill('.cmcard input', 'De acuerdo'); await w.press('.cmcard input', 'Enter');
    assert.equal(await w.evaluate(() => window.__escribiente.state.comments[0].replies.length), 1);
    log('Escribiente: comentario con respuesta');
    // control de cambios
    await w.evaluate(() => window.__escribiente.F.setTrack(true));
    await w.evaluate(() => { const t = document.querySelector('#p2').firstChild; const r = document.createRange(); r.setStart(t, t.length); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); window.__escribiente.editor.focus(); });
    await w.keyboard.type(' agregado');
    await w.keyboard.press('Home');
    await w.keyboard.press('Shift+End');
    await w.evaluate(() => { const t = document.querySelector('#p2').firstChild; const r = document.createRange(); r.setStart(t, 0); r.setEnd(t, 8); const s = getSelection(); s.removeAllRanges(); s.addRange(r); });
    await w.keyboard.press('Delete');
    const rev = await w.evaluate(() => ({ ins: document.querySelectorAll('#editor ins').length, del: document.querySelectorAll('#editor del').length, text: document.querySelector('#p2').textContent }));
    assert.ok(rev.ins >= 1 && rev.del >= 1, `marcas de cambios ${JSON.stringify(rev)}`);
    assert.ok(rev.text.includes('Párrafo') && rev.text.includes('agregado'), 'el texto borrado sigue visible tachado');
    await w.evaluate(() => window.__escribiente.F.setTrack(false));
    log('Escribiente: control de cambios marca inserciones y eliminaciones');
    // nota al pie
    await w.evaluate(() => { const t = document.querySelector('#p2').lastChild; const r = document.createRange(); r.selectNodeContents(document.querySelector('#p2')); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); window.__escribiente.editor.focus(); window.__escribiente.onCommand('selectAll'); r.collapse(false); s.removeAllRanges(); s.addRange(r); });
    const fnP = w.evaluate(() => window.__escribiente.onCommand('footnote'));
    await wait(300); await w.fill('.modal textarea', 'Fuente: INEGI 2025.'); await w.click('.modal .btn.primary'); await fnP;
    await w.evaluate(() => window.__escribiente.paginate());
    const fnDbg = await w.evaluate(() => ({ sups: document.querySelectorAll('#editor sup.fn').length, fns: Object.keys(window.__escribiente.state.footnotes).length, pageFns: JSON.stringify(window.__escribiente.state.pageFns), notes: document.querySelectorAll('.sheet .fnote').length }));
    assert.equal(fnDbg.notes, 1, JSON.stringify(fnDbg));
    log('Escribiente: nota al pie al final de la página');
    // columnas, cuadro de texto, marca de agua
    await w.evaluate(() => { const r = document.createRange(); r.selectNodeContents(document.querySelector('#p3')); const s = getSelection(); s.removeAllRanges(); s.addRange(r); window.__escribiente.editor.focus(); });
    await wait(100); await w.evaluate(() => window.__escribiente.F.setColumns(2));
    assert.equal(await w.evaluate(() => document.querySelector('#editor .cols')?.dataset.n), '2');
    await w.evaluate(() => { window.__escribiente.state.watermark = { text: 'BORRADOR', color: '#C9CBD6' }; window.__escribiente.paginate(); });
    assert.equal(await w.evaluate(() => document.querySelectorAll('.sheet .wm').length) > 0, true);
    await w.evaluate(() => window.__escribiente.F.toggleNav(true));
    assert.equal(await w.evaluate(() => document.querySelectorAll('.nav-item').length), 2);
    await w.screenshot({ path: path.join(out, 'escribiente-fase2.png') });
    log('Escribiente: columnas, marca de agua y panel de navegación');
    // guardar docx y revisar el XML
    await w.evaluate(() => window.__shell.save());
    const docx = await JSZip.loadAsync(fs.readFileSync(path.join(out, 'Sin título.docx')));
    const docXml = await docx.file('word/document.xml').async('string');
    assert.ok(docx.file('word/comments.xml'), 'comments.xml'); assert.ok(docx.file('word/footnotes.xml'), 'footnotes.xml');
    assert.ok(/<w:ins /.test(docXml) && /<w:del /.test(docXml), 'revisiones en el docx');
    assert.ok(/<w:cols [^>]*w:num="2"/.test(docXml), 'columnas en el docx');
    log('Escribiente: el .docx incluye comentarios, notas al pie, revisiones y columnas');
    // ODT ida y vuelta
    await w.evaluate(() => window.__shell.saveAsFormat('odt', 'ODF'));
    const odtPath = path.join(out, 'Sin título.odt');
    assert.ok(fs.existsSync(odtPath));
    const odt = await JSZip.loadAsync(fs.readFileSync(odtPath));
    assert.equal(await odt.file('mimetype').async('string'), 'application/vnd.oasis.opendocument.text');
    await reopen(w, odtPath); await wait(400);
    const back = await w.evaluate(() => ({ h1: document.querySelector('#editor h1')?.textContent, fn: document.querySelectorAll('#editor sup.fn').length, b: !!document.querySelector('#editor [style*="bold"], #editor b, #editor strong') }));
    assert.equal(back.h1, 'Introducción'); assert.equal(back.fn, 1); assert.ok(back.b);
    log('Escribiente: guarda y reabre .odt con títulos, negritas y nota al pie');
    // combinar correspondencia
    fs.writeFileSync(path.join(out, 'clientes.csv'), 'Nombre,Ciudad\nAna López,Acapulco\nBruno Díaz,Chilpancingo\nCarla Ruiz,Taxco\n');
    await w.evaluate(() => window.__shell.newDoc('blank'));
    await w.evaluate(async (csvPath) => {
      const { F, state } = window.__escribiente;
      const f = await window.oficina.readPath(csvPath);
      const { parseCsv } = { parseCsv: (t) => t.trim().split('\n').map((l) => l.split(',')) };
      const rows = parseCsv(new TextDecoder().decode(f.bytes));
      state.merge = { name: 'clientes.csv', fields: rows[0], rows: rows.slice(1), preview: false, idx: 0 };
      const ed = window.__escribiente.editor;
      ed.innerHTML = '<p>Hola <span class="mf" contenteditable="false" data-f="Nombre">«Nombre»</span>, de <span class="mf" contenteditable="false" data-f="Ciudad">«Ciudad»</span>.</p>';
      F.previewMerge(1);
    }, path.join(out, 'clientes.csv'));
    assert.ok((await w.evaluate(() => document.querySelector('#editor p').textContent)).includes('Bruno Díaz'));
    const merged = await w.evaluate(() => { window.__escribiente.F.previewMerge(0); return window.__escribiente.F.mergedHtml(); });
    assert.ok(merged.includes('Ana López') && merged.includes('Carla Ruiz') && (merged.match(/class="pb"/g) || []).length === 2);
    log('Escribiente: combinar correspondencia con vista previa y 3 registros');
    await closeApp('escribiente'); await wait(600);
  }

  // ============================================================ Calcula Datos
  if (!only || only === 'calcula') {
    const w = await openApp('calcula');
    await w.evaluate(() => window.__shell.newDoc('blank'));
    await w.evaluate(() => {
      const { wb } = window.__calcula;
      wb.action('datos', () => { [['Producto', 'Zona', 'Ventas'], ['Café', 'Norte', 120], ['Té', 'Sur', 80], ['Café', 'Sur', 150], ['Té', 'Norte', 60], ['Café', 'Norte', 90]].forEach((r, i) => r.forEach((v, j) => wb.setInput(i, j, String(v)))); });
      window.__calcula.afterLoad();
    });
    const C = (fn, arg) => w.evaluate(([f, a]) => window.__calcula.F[f](a), [fn, arg]);
    // formato condicional
    await w.evaluate(() => { const { wb, F } = window.__calcula; F.addRule({ range: 'C2:C6', type: 'gt', value: 100, style: { bg: '#FFCDD2', color: '#B71C1C' } }); F.addRule({ range: 'C2:C6', type: 'bar', color: '#5B9BD5' }); });
    const cf = await w.evaluate(() => window.__calcula.F.cfStyle(3, 2));
    assert.equal(cf.bg, '#FFCDD2');
    log('Calcula Datos: formato condicional (regla y barras de datos)');
    // validación
    await w.evaluate(() => window.__calcula.F.addValidation({ range: 'B2:B20', type: 'list', list: ['Norte', 'Sur', 'Centro'] }));
    const bad = await w.evaluate(() => window.__calcula.F.validate(1, 1, 'Oeste'));
    assert.equal(bad.ok, false);
    log('Calcula Datos: validación con lista');
    // nombres
    await w.evaluate(() => { const { wb, F } = window.__calcula; F.addName('Ventas', 'Hoja1!$C$2:$C$6'); wb.action('f', () => wb.setInput(7, 2, '=SUMA(Ventas)')); });
    assert.equal(await w.evaluate(() => window.__calcula.wb.value(7, 2)), 500);
    log('Calcula Datos: nombres definidos en fórmulas');
    // notas
    await w.evaluate(() => window.__calcula.F.setNote(1, 0, 'Producto estrella'));
    // tabla dinámica
    await w.evaluate(() => window.__calcula.F.buildPivot({ source: 'Hoja1!A1:C6', rows: 'Producto', cols: 'Zona', value: 'Ventas', agg: 'sum' }));
    const piv = await w.evaluate(() => { const { wb } = window.__calcula; const s = wb.sheets.find((x) => x.name.startsWith('Dinámica')); const ix = wb.sheets.indexOf(s); const g = (r, c) => wb.value(r, c, s); return { name: s?.name, total: g(4, 3), cafeNorte: g(2, 1) }; });
    assert.equal(piv.total, 500); assert.equal(piv.cafeNorte, 210);
    log('Calcula Datos: tabla dinámica (filas, columnas y totales)');
    await w.evaluate(() => { window.__calcula.wb.active = 0; window.__calcula.afterLoad(); window.__calcula.F.formatAsTable({ r1: 0, c1: 0, r2: 5, c2: 2 }, 'verde'); window.__calcula.F.protectSheet(true); });
    assert.equal(await w.evaluate(() => window.__calcula.F.canEdit(1, 1)), false);
    await w.evaluate(() => window.__calcula.F.protectSheet(false));
    await w.screenshot({ path: path.join(out, 'calcula-fase2.png') });
    log('Calcula Datos: formato de tabla y proteger hoja');
    await w.evaluate(() => window.__shell.save());
    const xlsxPath = path.join(out, 'Sin título.xlsx');
    await reopen(w, xlsxPath); await wait(400);
    const re = await w.evaluate(() => { const { wb } = window.__calcula; const s = wb.sheets[0]; return { cf: s.cf.length, val: s.validations.length, note: !!s.notes?.get?.('1,0'), name: wb.value(7, 2), sheets: wb.sheets.length }; });
    assert.ok(re.cf >= 1 && re.val === 1 && re.note && re.name === 500 && re.sheets === 2, JSON.stringify(re));
    log('Calcula Datos: el .xlsx conserva formato condicional, validación, notas y nombres');
    await w.evaluate(() => window.__shell.saveAsFormat('ods', 'ODS'));
    const odsPath = path.join(out, 'Sin título.ods');
    await reopen(w, odsPath); await wait(400);
    const ods = await w.evaluate(() => { const { wb } = window.__calcula; return { a1: wb.value(0, 0), f: wb.raw(7, 2), v: wb.value(7, 2), c3: wb.value(3, 2) }; });
    assert.equal(ods.a1, 'Producto'); assert.equal(ods.c3, 150); assert.equal(ods.v, 500);
    log('Calcula Datos: guarda y reabre .ods con valores y fórmulas');
    await closeApp('calcula'); await wait(600);
  }

  // ============================================================ Impresiona
  if (!only || only === 'impresiona') {
    const w = await openApp('impresiona');
    await w.evaluate(() => window.__shell.newDoc('negocio'));
    await wait(300);
    // gráfico
    await w.evaluate(() => window.__impresiona.F.insertChart({ type: 'column', title: 'Ventas', data: 'Mes,2025,2026\nEne,10,14\nFeb,12,18\nMar,9,20' }));
    assert.ok(await w.evaluate(() => window.__impresiona.pres.slides[window.__impresiona.cur].objects.some((o) => o.type === 'chart')));
    // animaciones de salida + orden
    await w.evaluate(() => { const I = window.__impresiona; const s = I.pres.slides[1]; s.objects[1].anim = { type: 'fadeOut', order: 1, start: 'click' }; I.renderAll(); });
    // patrón
    await w.evaluate(() => { const I = window.__impresiona; I.pres.master = { objects: [], footer: 'Mi empresa', showNumber: true }; I.renderAll(); });
    assert.ok(await w.evaluate(() => document.querySelector('.stage .master-footer')?.textContent.includes('Mi empresa')));
    // secciones y clasificador
    await w.evaluate(() => { const I = window.__impresiona; I.pres.slides[0].section = 'Inicio'; I.pres.slides[2].section = 'Propuesta'; I.onCommand('sorter'); });
    assert.equal(await w.evaluate(() => document.querySelectorAll('.sorter .sec-head').length), 2);
    await w.screenshot({ path: path.join(out, 'impresiona-clasificador.png') });
    await w.evaluate(() => window.__impresiona.onCommand('sorter'));
    log('Impresiona: gráfico, animación de salida, patrón y secciones');
    // pluma en presentación
    await w.evaluate(() => window.__impresiona.present(1));
    await wait(400);
    await w.keyboard.press('p');
    await w.mouse.move(400, 300); await w.mouse.down(); await w.mouse.move(600, 380, { steps: 5 }); await w.mouse.up();
    assert.ok(await w.evaluate(() => window.__show.inkCount() > 0));
    await w.keyboard.press('Escape'); await w.keyboard.press('Escape'); await wait(300);
    log('Impresiona: pluma durante la presentación');
    // exportar imágenes
    await w.evaluate(() => window.__impresiona.onCommand('exportImages'));
    await wait(3000);
    const pngs = fs.readdirSync(out).filter((f) => /^Diapositiva-\d+\.png$/.test(f));
    assert.equal(pngs.length, 4, `imágenes: ${pngs.join(',')}`);
    log('Impresiona: exporta cada diapositiva como PNG');
    await w.evaluate(() => window.__shell.save());
    const pptxPath = path.join(out, 'Sin título.pptx');
    const pz = await JSZip.loadAsync(fs.readFileSync(pptxPath));
    assert.ok(Object.keys(pz.files).some((f) => f.startsWith('ppt/charts/')), 'gráfico nativo en el pptx');
    await reopen(w, pptxPath); await wait(500);
    assert.ok(await w.evaluate(() => window.__impresiona.pres.slides.some((s) => s.objects.some((o) => o.type === 'chart'))));
    log('Impresiona: el gráfico se guarda como gráfico nativo del .pptx y se recupera');
    await w.evaluate(() => window.__shell.saveAsFormat('odp', 'ODP'));
    const odpPath = path.join(out, 'Sin título.odp');
    await reopen(w, odpPath); await wait(500);
    const odp = await w.evaluate(() => ({ n: window.__impresiona.pres.slides.length, t: window.__impresiona.pres.slides[0].objects.map((o) => o.html || '').join(' ') }));
    assert.equal(odp.n, 4); assert.ok(odp.t.includes('Plan de negocio 2027'));
    log('Impresiona: guarda y reabre .odp');
    // vista del moderador
    await w.evaluate(() => window.__impresiona.presenter(0));
    const aud = await waitWindow((u) => u.includes('audience=1'), 8000);
    await aud.waitForLoadState('domcontentloaded'); await wait(800);
    assert.ok(await w.evaluate(() => !!document.querySelector('.presenter .pv-next')));
    await w.evaluate(() => window.__presenter.next());
    await wait(500);
    assert.equal(await aud.evaluate(() => window.__audience.index()), 1);
    await w.screenshot({ path: path.join(out, 'impresiona-moderador.png') });
    await w.evaluate(() => window.__presenter.end()); await wait(500);
    log('Impresiona: vista del moderador con ventana para el público');
    await closeApp('impresiona'); await wait(500);
  }
  console.log('\nTodas las pruebas de la fase 2 pasaron.');
} catch (err) {
  console.error('✘', err);
  for (const w of app.windows()) { try { await w.screenshot({ path: path.join(out, `error-${Date.now()}.png`) }); } catch { /* */ } }
  process.exitCode = 1;
} finally { await app.close().catch(() => {}); }
