// Pruebas de la fase 5 (Escribiente como LibreOffice Writer), parte 1: abrir y guardar sin perder
// formato, cuadros Párrafo y Fuente, mayúsculas/minúsculas, marcas de formato.
// Linux sin pantalla:  xvfb-run -a node test/e2e-fase5.mjs
import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'fase5');
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

const hasSoffice = (() => { try { execFileSync('soffice', ['--version'], { stdio: 'ignore' }); return true; } catch { return false; } })();

// pone valores en el cuadro abierto (también en pestañas ocultas) y acepta
async function setDialog(w, values, submit = true) {
  await w.waitForSelector('.modal', { timeout: 5000 });
  await w.evaluate((vals) => {
    const m = [...document.querySelectorAll('.modal')].pop();
    for (const [k, v] of Object.entries(vals)) { const el = m.querySelector(`[name="${k}"]`); if (!el) throw new Error(`sin campo ${k}`); if (el.type === 'checkbox') el.checked = !!v; else el.value = String(v); el.dispatchEvent(new Event('change', { bubbles: true })); el.dispatchEvent(new Event('input', { bubbles: true })); }
  }, values);
  if (submit) { await w.click('.modal .btn.primary'); await wait(300); }
}
const fx = (p) => path.join(root, 'test', 'fixtures', p);
// lo que muestra el editor tras abrir el documento de prueba
const probe = (w) => w.evaluate(() => {
  const ed = document.getElementById('editor'); const cs = (el) => (el ? getComputedStyle(el) : {});
  const byText = (t, sel = 'p,h1,h2,h3,li') => [...ed.querySelectorAll(sel)].find((p) => p.textContent.includes(t));
  const span = (t) => { const tw = document.createTreeWalker(ed, NodeFilter.SHOW_TEXT); let n; while ((n = tw.nextNode())) if (n.nodeValue.includes(t)) return n.parentElement; return null; };
  const p1 = byText('Párrafo justificado'); const p2 = byText('sangría izquierda y derecha'); const p3 = byText('borde y sombreado');
  const tbl = ed.querySelector('table');
  const st = window.__escribiente.state;
  return {
    title: !!ed.querySelector('p.title'), h1: ed.querySelector('h1')?.textContent,
    p1: { align: cs(p1).textAlign, mt: Math.round(parseFloat(cs(p1).marginTop)), lh: (parseFloat(cs(p1).lineHeight) / parseFloat(cs(p1).fontSize)).toFixed(1), ti: Math.round(parseFloat(cs(p1).textIndent)) },
    p2: { ml: Math.round(parseFloat(cs(p2).marginLeft)), mr: Math.round(parseFloat(cs(p2).marginRight)), lh: Math.round(parseFloat(cs(p2).lineHeight)), keep: p2?.dataset.keepnext },
    p3: { border: cs(p3).borderTopStyle, bcolor: cs(p3).borderTopColor, bg: cs(p3).backgroundColor },
    red: cs(span('Rojo Georgia')).color, georgia: cs(span('Rojo Georgia')).fontFamily.split(',')[0], size16: cs(span('Rojo Georgia')).fontSize,
    hl: cs(span('resaltado')).backgroundColor, smallcaps: cs(span('versalitas')).fontVariantCaps, caps: cs(span('mayúsculas')).textTransform,
    dstrike: cs(span('doble tachado')).textDecorationStyle, ls: Math.round(parseFloat(cs(span('espaciado')).letterSpacing)), hidden: cs(span('oculto')).display,
    link: ed.querySelector('a[href^="https://www.libreoffice"]')?.textContent,
    ul: ed.querySelectorAll('ul > li').length, ol: ed.querySelectorAll('ol > li').length,
    colspan: tbl?.querySelector('[colspan="2"]')?.textContent.trim(), rowspan: tbl?.querySelector('[rowspan="2"]')?.textContent.trim(),
    cellBg: cs([...(tbl?.querySelectorAll('td') || [])].find((c) => c.textContent.trim() === 'C10')).backgroundColor,
    img: !!ed.querySelector('img[src^="data:image/png"]'), pb: ed.querySelectorAll('.pb').length,
    header: st.header, footer: st.footer, pages: st.pages,
  };
});
const expectFormato = (r, label) => {
  assert.ok(r.title, `${label}: título`); assert.equal(r.h1, 'Introducción');
  assert.equal(r.p1.align, 'justify', label); assert.equal(r.p1.mt, 24, `${label} espacio antes`); assert.equal(r.p1.lh, '2.0', `${label} interlineado doble`); assert.ok(r.p1.ti >= 44 && r.p1.ti <= 49, `${label} sangría ${r.p1.ti}`);
  assert.ok(Math.abs(r.p2.ml - 76) <= 2 && Math.abs(r.p2.mr - 76) <= 2, `${label} sangrías ${JSON.stringify(r.p2)}`); assert.equal(r.p2.lh, 27, `${label} interlineado exacto 20 pt`);
  assert.equal(r.p3.border, 'solid'); assert.equal(r.p3.bcolor, 'rgb(192, 0, 0)', label); assert.equal(r.p3.bg, 'rgb(255, 242, 204)', label);
  assert.equal(r.red, 'rgb(192, 0, 0)'); assert.equal(r.georgia.replace(/["']/g, ''), 'Georgia'); assert.equal(r.size16, '21.3333px');
  assert.equal(r.hl, 'rgb(255, 255, 0)', label); assert.equal(r.smallcaps, 'small-caps', label); assert.equal(r.caps, 'uppercase', label);
  assert.equal(r.dstrike, 'double', label); assert.equal(r.ls, 4, `${label} espaciado`); assert.equal(r.hidden, 'none', `${label} oculto`);
  assert.equal(r.link, 'LibreOffice'); assert.equal(r.ul, 3, label); assert.equal(r.ol, 3, label);
  assert.equal(r.colspan, 'Combinada horizontal', label); assert.equal(r.rowspan, 'Combinada vertical', label); assert.equal(r.cellBg, 'rgb(217, 226, 243)', label);
  assert.ok(r.img, `${label}: imagen`); assert.equal(r.header, 'Informe anual — Mi Empresa', label); assert.equal(r.footer, 'Página {pagina} de {paginas}', label);
  assert.equal(r.pages, 3, `${label}: páginas`);
};

try {
  const w = await openApp('escribiente');
  // 1) abrir documentos con formato: hecho en Word (python-docx), guardado por LibreOffice como .docx y como .odt
  for (const f of ['formato.docx', 'formato-lo.docx', 'formato-lo.odt']) {
    await reopen(w, fx(f)); await wait(900);
    expectFormato(await probe(w), f);
    log(`Escribiente abre ${f} con su formato (estilos, sangrías, espaciado, bordes, efectos de letra, listas, tabla combinada, encabezado y pie)`);
  }
  // 2) guardar en .docx y .odt y volver a abrir: el formato se conserva
  for (const [ext, lab] of [['docx', 'Word'], ['odt', 'ODT']]) {
    await reopen(w, fx('formato.docx')); await wait(800);
    await w.evaluate(async ([e, l]) => window.__shell.saveAsFormat(e, l), [ext, lab]); await wait(800);
    const saved = path.join(out, `formato.${ext}`);
    assert.ok(fs.existsSync(saved), saved);
    await reopen(w, saved); await wait(900);
    expectFormato(await probe(w), `ida y vuelta .${ext}`);
    if (hasSoffice) {
      // LibreOffice lo abre y lo vuelve a guardar; Escribiente lo sigue leyendo igual
      const lo = path.join(out, `lo-${ext}`); fs.mkdirSync(lo, { recursive: true });
      execFileSync('soffice', ['--headless', '--convert-to', ext === 'docx' ? 'odt' : 'docx', '--outdir', lo, saved], { stdio: 'ignore' });
      const back = path.join(lo, `formato.${ext === 'docx' ? 'odt' : 'docx'}`);
      await reopen(w, back); await wait(900);
      expectFormato(await probe(w), `LibreOffice ${ext}→${ext === 'docx' ? 'odt' : 'docx'}`);
    }
    log(`Escribiente guarda .${ext} sin perder el formato${hasSoffice ? ' (comprobado también al pasar por LibreOffice)' : ''}`);
  }
  // 3) cuadro Párrafo
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(400);
  await w.click('#editor'); await w.keyboard.type('Primer párrafo de prueba.'); await w.keyboard.press('Enter'); await w.keyboard.type('Segundo párrafo.');
  await w.keyboard.press('ArrowUp');
  await startDialog(w, 'paraDialog');
  await setDialog(w, { ml: 1.5, mr: 1, special: 'first', ti: 0.75, mt: 12, mb: 6, lh: 'exact', lhv: 18, align: 'justify', keepnext: true, pbb: true, border: 'bottom', bstyle: 'double', bw: 1.5, bcolor: '#1f6feb', bpad: 3, shadeOn: true, shade: '#eeeeee' });
  const pr = await w.evaluate(() => { const p = document.querySelector('#editor p'); const c = getComputedStyle(p); return { ml: Math.round(parseFloat(c.marginLeft)), mr: Math.round(parseFloat(c.marginRight)), ti: Math.round(parseFloat(c.textIndent)), mt: p.dataset.mt0 || p.style.marginTop, mb: c.marginBottom, lh: c.lineHeight, al: c.textAlign, keep: p.dataset.keepnext, pbb: p.dataset.pbb, bb: `${c.borderBottomStyle} ${c.borderBottomColor}`, bt: c.borderTopStyle, bg: c.backgroundColor }; });
  assert.deepEqual(pr, { ml: 57, mr: 38, ti: 28, mt: '12pt', mb: '8px', lh: '24px', al: 'justify', keep: '1', pbb: '1', bb: 'double rgb(31, 111, 235)', bt: 'none', bg: 'rgb(238, 238, 238)' }, JSON.stringify(pr));
  log('Formato → Párrafo: sangrías, espaciado, interlineado exacto, alineación, flujo del texto, bordes y sombreado');
  // 4) cuadro Fuente sobre una palabra
  await w.evaluate(() => { const t = document.querySelector('#editor p').firstChild; const r = document.createRange(); r.setStart(t, 0); r.setEnd(t, 6); const s = getSelection(); s.removeAllRanges(); s.addRange(r); window.__escribiente.saveSelection(); });
  await startDialog(w, 'fontDialog');
  await setDialog(w, { font: 'Georgia', fstyle: 'b', size: 14, color: '#aa0000', under: 'double', smallcaps: true, dstrike: false, ls: 2, scale: 100, pos: 0, shadow: true });
  const fr = await w.evaluate(() => { const tw = document.createTreeWalker(document.querySelector('#editor p'), NodeFilter.SHOW_TEXT); tw.nextNode(); const el = tw.currentNode.parentElement; const c = getComputedStyle(el); let deco = el; while (deco && !/underline/.test(getComputedStyle(deco).textDecorationLine)) deco = deco.parentElement; return { text: tw.currentNode.nodeValue, ff: c.fontFamily.split(',')[0].replace(/["']/g, ''), fw: c.fontWeight, fs: c.fontSize, color: c.color, sc: c.fontVariantCaps, ls: c.letterSpacing, u: deco ? getComputedStyle(deco).textDecorationStyle : null, sh: c.textShadow !== 'none' }; });
  if (fr.ff !== 'Georgia') console.log('HTML', await w.evaluate(() => document.querySelector('#editor').innerHTML));
  assert.equal(fr.text, 'Primer'); assert.equal(fr.ff, 'Georgia'); assert.equal(fr.fw, '700'); assert.equal(fr.fs, '18.6667px'); assert.equal(fr.color, 'rgb(170, 0, 0)'); assert.equal(fr.sc, 'small-caps'); assert.equal(fr.ls, '2.66667px'); assert.equal(fr.u, 'double'); assert.ok(fr.sh);
  log('Formato → Fuente: familia, estilo, tamaño, color, subrayado doble, versalitas, sombra, espaciado entre caracteres');
  // se guardan en .docx con las propiedades de Word
  await w.evaluate(async () => window.__shell.saveAsFormat('docx', 'Word')); await wait(700);
  const dz = await JSZip.loadAsync(fs.readFileSync(path.join(out, `${UNTITLED}.docx`)));
  const dx = await dz.file('word/document.xml').async('string');
  for (const re of [/<w:ind w:left="85\d" w:right="5[67]\d" w:firstLine="4\d\d"\/>/, /<w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="exact"\/>/, /<w:keepNext\/>/, /<w:pageBreakBefore\/>/, /<w:pBdr><w:bottom w:val="double"/, /w:fill="EEEEEE"/, /<w:smallCaps\/>/, /<w:spacing w:val="40"\/>/, /<w:u w:val="double"\/>/, /<w:rFonts w:ascii="Georgia"/]) assert.match(dx, re);
  await reopen(w, path.join(out, `${UNTITLED}.docx`)); await wait(700);
  const back = await w.evaluate(() => { const p = document.querySelector('#editor p'); const c = getComputedStyle(p); return { ml: Math.round(parseFloat(c.marginLeft)), lh: c.lineHeight, keep: p.dataset.keepnext, bb: c.borderBottomStyle, sc: !!p.querySelector('[style*="small-caps"]') }; });
  assert.deepEqual(back, { ml: 57, lh: '24px', keep: '1', bb: 'double', sc: true });
  log('El formato de párrafo y de letra se guarda en .docx como propiedades de Word y se recupera al abrirlo');
  // 5) mayúsculas y minúsculas, marcas de formato
  await w.evaluate(() => window.__shell.newDoc('blank')); await wait(300);
  await w.click('#editor'); await w.keyboard.type('hola mundo. otra frase aquí');
  await w.keyboard.press('Control+a');
  const cases = [];
  for (const m of ['upper', 'lower', 'title', 'sentence', 'toggle']) { await w.evaluate((mm) => window.__escribiente.F_X.changeCase(mm), m); cases.push(await w.evaluate(() => document.querySelector('#editor p').textContent)); }
  assert.deepEqual(cases, ['HOLA MUNDO. OTRA FRASE AQUÍ', 'hola mundo. otra frase aquí', 'Hola Mundo. Otra Frase Aquí', 'Hola mundo. Otra frase aquí', 'hOLA MUNDO. oTRA FRASE AQUÍ']);
  await w.evaluate(() => window.__shell.command('marks')); await wait(100);
  assert.ok(await w.evaluate(() => document.getElementById('editor').classList.contains('marks') && getComputedStyle(document.querySelector('#editor p'), '::after').content.includes('¶')));
  await w.evaluate(() => window.__shell.command('marks'));
  log('Cambiar mayúsculas y minúsculas (5 modos) y mostrar marcas de formato ¶');
  // 6) conservar con el siguiente: el título no se queda solo al final de la página
  await w.evaluate(() => { const ed = document.getElementById('editor'); let h = ''; for (let i = 0; i < 37; i++) h += `<p>Línea de relleno ${i}</p>`; h += '<h2 data-keepnext="1">Título que no debe quedar solo</h2><p>Texto que sigue al título, bastante largo para no caber al final de la página actual de ningún modo posible aquí.</p>'; window.__escribiente.setContent(h); });
  await wait(600);
  const kn = await w.evaluate(() => { const st = window.__escribiente.state; const ed = document.getElementById('editor'); const kids = [...ed.children]; const hi = kids.findIndex((k) => k.tagName === 'H2'); return [st.pageOf[hi], st.pageOf[hi + 1]]; });
  assert.equal(kn[0], kn[1], `título y párrafo en la misma página: ${kn}`);
  log('“Conservar con el siguiente” lleva el título a la página del párrafo que lo sigue');
  await closeApp('escribiente'); await wait(600);
  assert.deepEqual(pageErrors, [], `errores en las ventanas:\n${pageErrors.join('\n')}`);
  console.log('\nTodas las pruebas de la fase 5 (parte 1) pasaron.');
} catch (e) {
  console.error('✘', e);
  process.exitCode = 1;
} finally {
  await app.close();
}
