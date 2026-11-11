// Pruebas de la fase 5, parte 7 (Escribiente como LibreOffice Writer): Archivo — RTF, EPUB, texto codificado,
// propiedades del documento, opciones de PDF, diálogo de impresión y vista previa.

import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase5g');
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
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFWEXHbQSACj/P8Fu7N9hAAAAAElFTkSuQmCC';
const DOC = `<h1>Capítulo uno</h1><p>Texto con <b>negrita</b>, <i>cursiva</i>, <u>subrayado</u>, <s>tachado</s>, <span style="color:#c00000">rojo</span> y un <a href="https://ejemplo.mx/">vínculo</a>. Año €ñ.<sup class="fn" data-fid="f1" contenteditable="false">1</sup></p>
<p style="text-align:center">Centrado</p><ul><li>uno</li><li>dos</li></ul><ol><li>primero</li><li>segundo</li></ol>
<table class="tbl"><tbody><tr><td colspan="2">combinada</td><td>c</td></tr><tr><td>a1</td><td>b1</td><td>c1</td></tr></tbody></table>
<p><img src="${PNG}" width="40" height="40"></p><div class="pb" contenteditable="false"></div><h1>Capítulo dos</h1><h2>Sección 2.1</h2><p>Final del documento.</p>`;
const setDoc = (w) => E(w, ([html]) => { const st = window.__escribiente.state; st.footnotes = { f1: 'Nota de prueba' }; st.endnotes = {}; st.header = 'Mi encabezado\t\tPágina {pagina}'; window.__escribiente.setContent(html); }, [DOC]);
try {
  const w = await openApp('escribiente');
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  // 1) RTF: escribir y volver a leer
  await setDoc(w); await wait(300);
  await E(w, () => { window.__shell.doc.name = 'prueba.docx'; });
  await E(w, () => window.__shell.saveAsFormat('rtf', 'RTF'));
  const rtfPath = path.join(out, 'prueba.rtf');
  const rtf = fs.readFileSync(rtfPath, 'latin1');
  assert.match(rtf, /^\{\\rtf1\\ansi/); assert.match(rtf, /\\pngblip/); assert.match(rtf, /HYPERLINK "https:\/\/ejemplo.mx\/"/); assert.match(rtf, /\\footnote/);
  await E(w, () => window.__escribiente.setContent('<p>x</p>'));
  await reopen(w, rtfPath); await wait(500);
  const r1 = await E(w, () => { const ed = document.getElementById('editor'); return { h1: [...ed.querySelectorAll('h1')].map((x) => x.textContent), h2: ed.querySelector('h2')?.textContent, b: ed.querySelector('b')?.textContent, i: ed.querySelector('i')?.textContent, u: ed.querySelector('u')?.textContent, s: ed.querySelector('s')?.textContent, red: !!ed.querySelector('span[style*="rgb(192, 0, 0)"], span[style*="#c00000"]'), a: ed.querySelector('a[href]')?.getAttribute('href'), ul: ed.querySelectorAll('ul li').length, ol: ed.querySelectorAll('ol li').length, cs: ed.querySelector('td[colspan="2"]')?.textContent, cells: ed.querySelectorAll('td').length, img: !!ed.querySelector('img[src^="data:image/png"]'), pb: ed.querySelectorAll('.pb').length, fn: Object.values(window.__escribiente.state.footnotes), center: ed.querySelector('p[style*="center"]')?.textContent, txt: ed.innerText, header: window.__escribiente.state.header }; });
  assert.deepEqual(r1.h1, ['Capítulo uno', 'Capítulo dos']); assert.equal(r1.h2, 'Sección 2.1');
  assert.equal(r1.b, 'negrita'); assert.equal(r1.i, 'cursiva'); assert.equal(r1.u, 'subrayado'); assert.equal(r1.s, 'tachado'); assert.ok(r1.red);
  assert.equal(r1.a, 'https://ejemplo.mx/'); assert.equal(r1.ul, 2); assert.equal(r1.ol, 2); assert.equal(r1.cs, 'combinada'); assert.equal(r1.cells, 5);
  assert.ok(r1.img); assert.equal(r1.pb, 1); assert.deepEqual(r1.fn, ['Nota de prueba']); assert.equal(r1.center, 'Centrado'); assert.match(r1.txt, /Año €ñ/);
  assert.equal(r1.header, 'Mi encabezado\t\tPágina {pagina}');
  log('RTF: títulos, formato de carácter, color, vínculo, listas, tabla con celdas combinadas, imagen, nota al pie, salto de página y encabezado');
  if (hasSoffice) {
    execFileSync('soffice', ['--headless', '--convert-to', 'odt', '--outdir', out, rtfPath], { stdio: 'ignore' });
    const z = await JSZip.loadAsync(fs.readFileSync(path.join(out, 'prueba.odt')));
    const cx = await z.file('content.xml').async('string');
    assert.match(cx, /Capítulo uno/); assert.match(cx, /text:note[^>]*text:note-class="footnote"/); assert.match(cx, /Nota de prueba/); assert.match(cx, /table:number-columns-spanned="2"/); assert.match(cx, /xlink:href="https:\/\/ejemplo.mx\/"/);
    assert.match(cx, /<text:h[^>]*text:outline-level="1"/);
    // LibreOffice → RTF → Escribiente
    fs.mkdirSync(path.join(out, 'lo'), { recursive: true });
    execFileSync('soffice', ['--headless', '--convert-to', 'rtf', '--outdir', path.join(out, 'lo'), path.join(out, 'prueba.odt')], { stdio: 'ignore' });
    await reopen(w, path.join(out, 'lo', 'prueba.rtf')); await wait(500);
    const r2 = await E(w, () => { const ed = document.getElementById('editor'); return { h1: [...ed.querySelectorAll('h1')].map((x) => x.textContent.trim()), b: ed.querySelector('b')?.textContent, a: ed.querySelector('a[href]')?.getAttribute('href'), cs: ed.querySelector('td[colspan="2"]')?.textContent.trim(), img: !!ed.querySelector('img[src^="data:image/"]'), fn: Object.values(window.__escribiente.state.footnotes), li: ed.querySelectorAll('li').length, txt: ed.innerText }; });
    assert.deepEqual(r2.h1, ['Capítulo uno', 'Capítulo dos']); assert.equal(r2.b, 'negrita'); assert.equal(r2.a, 'https://ejemplo.mx/'); assert.equal(r2.cs, 'combinada');
    assert.ok(r2.img, 'imagen desde LibreOffice'); assert.deepEqual(r2.fn, ['Nota de prueba']); assert.equal(r2.li, 4); assert.match(r2.txt, /Año €ñ/); assert.doesNotMatch(r2.txt, /HYPERLINK|Heading|Calibri/);
    log('RTF con LibreOffice: lo abre con todo y Escribiente lee el RTF que escribe LibreOffice');
  }
  // 2) EPUB
  await setDoc(w); await wait(300);
  await startDialog(w, 'exportEpub'); await setDialog(w, { split: 'h1', cover: true, title: 'Mi libro', author: 'Ana' });
  await wait(500);
  const ez = await JSZip.loadAsync(fs.readFileSync(path.join(out, 'prueba.epub')));
  const names = Object.keys(ez.files);
  assert.equal(names[0], 'mimetype'); assert.equal(await ez.file('mimetype').async('string'), 'application/epub+zip');
  assert.ok(names.includes('OEBPS/chap001.xhtml') && names.includes('OEBPS/chap002.xhtml') && !names.includes('OEBPS/chap003.xhtml'));
  assert.ok(names.some((n) => /OEBPS\/images\/img1\.png/.test(n)));
  const nav = await ez.file('OEBPS/nav.xhtml').async('string'); assert.match(nav, /Capítulo uno[\s\S]*Capítulo dos[\s\S]*Sección 2\.1/);
  const opf = await ez.file('OEBPS/content.opf').async('string'); assert.match(opf, /<dc:title>Mi libro<\/dc:title>/); assert.match(opf, /<dc:creator>Ana<\/dc:creator>/); assert.match(opf, /cover\.xhtml/);
  const ch1 = await ez.file('OEBPS/chap001.xhtml').async('string'); assert.match(ch1, /epub:type="footnote"[^>]*>Nota de prueba/);
  for (const n of names.filter((q) => /\.(xhtml|opf|ncx)$/.test(q))) fs.writeFileSync(path.join(out, path.basename(n)), await ez.file(n).async('string'));
  py(`import xml.dom.minidom,glob\nfor f in glob.glob('${out}/*.xhtml')+glob.glob('${out}/*.opf')+glob.glob('${out}/*.ncx'): xml.dom.minidom.parse(f)`);
  log('EPUB 3: capítulos por título, índice anidado, portada, imágenes, notas y metadatos (XHTML bien formado)');
  // 3) texto codificado
  await startDialog(w, 'exportTxt'); await setDialog(w, { enc: 'windows-1252', eol: 'crlf' }); await wait(300);
  const tb = fs.readFileSync(path.join(out, 'prueba.txt'));
  assert.ok(tb.includes(Buffer.from([0x41, 0xF1, 0x6F, 0x20, 0x80, 0xF1]))); assert.ok(tb.includes(Buffer.from('\r\n')));
  await reopen(w, path.join(out, 'prueba.txt')); await wait(300);
  assert.match(await text(w), /Año €ñ/);
  log('Texto codificado (Windows-1252 y CR+LF) y lectura de texto que no es UTF-8');
  // 4) propiedades del documento
  await setDoc(w); await wait(200);
  await startDialog(w, 'docProperties');
  await w.waitForSelector('.modal .fx-tab'); await wait(200);
  await E(w, () => { const m = document.querySelector('.modal'); m.querySelector('[name=title]').value = 'Informe anual'; m.querySelector('[name=description]').value = 'Comentario del documento'; m.querySelector('[name=category]').value = 'Reportes'; });
  await w.click('.modal .fx-tab >> nth=2');
  for (let k = 0; k < 4; k++) await w.click('.modal .fx-pane:not([hidden]) > .fx-col > .btn');
  await E(w, () => {
    const rows = document.querySelectorAll('.modal .cp-row');
    const set = (r, name, type, val) => { r.querySelector('.cp-name').value = name; const s = r.querySelector('.cp-type'); s.value = type; s.dispatchEvent(new Event('change')); const v = r.querySelector('.cp-val'); if (type === 'bool') v.checked = val; else v.value = val; };
    set(rows[0], 'Cliente', 'text', 'ACME'); set(rows[1], 'Monto', 'number', '1500.5'); set(rows[2], 'Entrega', 'date', '2026-12-01'); set(rows[3], 'Aprobado', 'bool', true);
  });
  await w.click('.modal .btn.primary'); await wait(300);
  await E(w, () => window.__shell.saveAsFormat('docx', 'Word'));
  await E(w, () => window.__shell.saveAsFormat('odt', 'ODF'));
  const dz = await JSZip.loadAsync(fs.readFileSync(path.join(out, 'prueba.docx')));
  assert.match(await dz.file('docProps/core.xml').async('string'), /<dc:description>Comentario del documento<\/dc:description>/);
  const cu = await dz.file('docProps/custom.xml').async('string');
  assert.match(cu, /name="Cliente"><vt:lpwstr>ACME/); assert.match(cu, /name="Monto"><vt:r8>1500.5/); assert.match(cu, /name="Aprobado"><vt:bool>true/);
  assert.match(await dz.file('[Content_Types].xml').async('string'), /custom-properties/);
  for (const f of ['prueba.docx', 'prueba.odt']) {
    await E(w, () => { window.__shell.props = {}; });
    await reopen(w, path.join(out, f)); await wait(400);
    const pr = await E(w, () => window.__shell.props);
    assert.equal(pr.title, 'Informe anual', f); assert.equal(pr.description, 'Comentario del documento', f); assert.equal(pr.category, 'Reportes', f);
    assert.deepEqual(pr.custom.map((c) => [c.name, c.type, String(c.value)]), [['Cliente', 'text', 'ACME'], ['Monto', 'number', '1500.5'], ['Entrega', 'date', '2026-12-01'], ['Aprobado', 'bool', 'true']], f);
  }
  if (hasSoffice) {
    fs.mkdirSync(path.join(out, 'lo2'), { recursive: true });
    execFileSync('soffice', ['--headless', '--convert-to', 'odt', '--outdir', path.join(out, 'lo2'), path.join(out, 'prueba.docx')], { stdio: 'ignore' });
    const mz = await (await JSZip.loadAsync(fs.readFileSync(path.join(out, 'lo2', 'prueba.odt')))).file('meta.xml').async('string');
    assert.match(mz, /Comentario del documento/); assert.match(mz, /meta:name="Cliente"[^>]*>ACME/); assert.match(mz, /meta:name="Monto"[^>]*meta:value-type="float"|meta:value-type="float"[^>]*meta:name="Monto"/);
  }
  log('Propiedades del documento: comentarios, categoría y propiedades personalizadas (texto, número, fecha, sí/no) en .docx y .odt');
  // 5) opciones de PDF
  await reopen(w, path.join(out, 'prueba.docx')); await wait(500);
  await startDialog(w, 'exportPdf');
  await setDialog(w, { range: 'pages', pages: '2', outline: true, wm: 'BORRADOR', initialView: 'outline', layout: 'continuous', displayTitle: true, open: false });
  await wait(1500);
  const pdf1 = path.join(out, 'prueba.pdf');
  const info = py(`import pypdf\nr=pypdf.PdfReader('${pdf1}')\nc=r.trailer['/Root']\nprint(len(r.pages), r.metadata.title, c.get('/PageMode'), c.get('/PageLayout'), 'BORRADOR' in r.pages[0].extract_text(), len(r.outline)>0)`);
  assert.match(info, /^1 Informe anual \/UseOutlines \/OneColumn True True$/);
  fs.renameSync(pdf1, path.join(out, 'sin-clave.pdf'));
  await startDialog(w, 'exportPdf');
  await setDialog(w, { range: 'all', wm: '', user: 'abc', user2: 'abc', owner: 'dueño', copy: false });
  await wait(1500);
  assert.match(py(`import pypdf\nr=pypdf.PdfReader('${pdf1}')\nprint(r.is_encrypted, r.decrypt('abc')>0, len(r.pages))`), /^True True [2-9]$/);
  log('Opciones de PDF: intervalo, marcadores, marca de agua, vista inicial, propiedades y contraseña con permisos');
  // 6) diálogo Imprimir con vista previa y vista previa de impresión
  await startDialog(w, 'print');
  await w.waitForSelector('.modal .pv-sheet canvas.pv-page', { timeout: 15000 });
  await setDialog(w, { range: 'pages', pages: '1-2', pps: '2', color: 'gray', duplex: 'longEdge' }, false);
  await wait(600);
  assert.equal(await E(w, () => document.querySelectorAll('.modal .pv-sheet canvas').length), 2);
  await w.click('.modal .btn.primary'); await wait(800);
  const pj = JSON.parse(fs.readFileSync(path.join(out, 'impresion.json'), 'utf8'));
  assert.deepEqual(pj.pageRanges, [{ from: 0, to: 1 }]); assert.equal(pj.pagesPerSheet, 2); assert.equal(pj.color, false); assert.equal(pj.duplexMode, 'longEdge');
  await startDialog(w, 'printPreview');
  await w.waitForSelector('.pv-full figure canvas', { timeout: 15000 }); await wait(800);
  const pages = await E(w, () => window.__escribiente.state.pages);
  assert.equal(await E(w, () => document.querySelectorAll('.pv-full figure').length), pages);
  await w.keyboard.press('Escape'); await wait(200);
  assert.equal(await E(w, () => !!document.querySelector('.pv-full')), false);
  log('Imprimir: vista previa, intervalo, páginas por hoja, escala de grises y doble cara; vista previa de impresión');
  await closeApp('escribiente'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de la fase 5 (parte 7) pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
