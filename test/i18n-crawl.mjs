// Recorre la interfaz en inglés y lista los textos que siguen en español (para completar el diccionario).
import { _electron as electron } from 'playwright';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'i18n');
fs.rmSync(out, { recursive: true, force: true }); fs.mkdirSync(out, { recursive: true });
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const app = await electron.launch({ args: [root, `--user-data-dir=${path.join(out, 'perfil')}`, '--no-sandbox'], cwd: root, env: { ...process.env, OFICINA_TEST_DIR: out, OFICINA_LANG: 'en' } });
const SKIP = new Set(['present', 'presentCurrent', 'print', 'exportPdf', 'close', 'save', 'saveAs', 'exportImages', 'presenter', 'open', 'insertImage', 'chooseRecipients', 'master']);
async function win(pred) { for (let i = 0; i < 100; i++) { for (const w of app.windows()) if (pred(w.url())) return w; await wait(150); } throw new Error('ventana'); }
const all = new Set();
const SPANISH = /[áéíóúñ¿¡]|\b(de|la|el|los|las|del|para|con|sin|una?|que|por|en|se|al|y|o)\b/i;
async function harvest(w, tag) {
  const found = await w.evaluate((skipSel) => {
    const out = new Set();
    const skip = (el) => !el || el.closest('script,style,textarea,[contenteditable="true"],.no-i18n' + (skipSel ? `,${skipSel}` : ''));
    const tw = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    while (tw.nextNode()) { const n = tw.currentNode; const v = n.nodeValue.trim(); if (v && /[a-z]{2}/i.test(v) && !skip(n.parentElement) && n.parentElement.offsetParent !== null) out.add(v); }
    document.querySelectorAll('[title],[placeholder],[aria-label]').forEach((el) => { if (skip(el)) return; for (const a of ['title', 'placeholder', 'aria-label']) { const v = el.getAttribute(a); if (v && /[a-z]{2}/i.test(v)) out.add(v); } });
    for (const m of window.__i18nMissing || []) out.add(m);
    return [...out];
  }, await w.evaluate(() => window.__shell?.adapter?.i18nSkip || document.documentElement.dataset.i18nSkip || ''));
  for (const s of found) if (SPANISH.test(s)) all.add(s);
}
const l = await win((u) => u.includes('/launcher/'));
await wait(1200);
await l.evaluate(() => window.__launcher?.setOpen?.(true)); await wait(600);
await harvest(l, 'launcher');
await l.evaluate(() => document.getElementById('prefs').hidden = false); await harvest(l, 'prefs');
await l.evaluate(() => document.getElementById('about').hidden = false); await harvest(l, 'about');
for (const id of ['escribiente', 'calcula', 'impresiona']) {
  await l.evaluate((x) => window.__launcher.launch(x), id);
  const w = await win((u) => u.includes(`/${id}/`));
  await w.waitForLoadState('domcontentloaded'); await wait(1500);
  w.on('dialog', (d) => d.dismiss().catch(() => {}));
  for (const page of ['new', 'info', 'export', 'open', 'save']) { await w.evaluate((p) => window.__shell.showBackstage(p), page).catch(() => {}); await wait(250); await harvest(w, page); }
  await w.evaluate(() => window.__shell.hideBackstage?.() ?? window.__shell.closeBackstage?.()).catch(() => {});
  await w.keyboard.press('Escape');
  await w.evaluate(() => window.__shell.newDoc?.('blank')).catch(() => {}); await wait(400);
  const tabs = await w.evaluate(() => [...document.querySelectorAll('.ribbon .tabs .tab[data-tab]')].map((t) => t.dataset.tab));
  for (const t of tabs) {
    await w.evaluate((x) => window.__shell.ribbon.show(x), t); await wait(200);
    await harvest(w, t);
    const cmds = await w.evaluate(() => [...document.querySelectorAll('.ribbon .panel [data-cmd]')].map((b) => b.dataset.cmd));
    for (const c of cmds) {
      if (SKIP.has(c)) continue;
      try {
        await w.evaluate((x) => { const b = document.querySelector(`.ribbon .panel [data-cmd="${x}"]`); b?.click(); }, c);
        await wait(220);
        await harvest(w, `${t}/${c}`);
        await w.keyboard.press('Escape'); await wait(60); await w.keyboard.press('Escape');
        await w.evaluate(() => { document.querySelectorAll('.modal-back, .dropdown, .dd').forEach((n) => n.remove()); });
        await w.evaluate((x) => window.__shell.ribbon.show(x), t).catch(() => {});
      } catch (e) { console.error(id, c, e.message); }
    }
  }
  await app.evaluate(({ BrowserWindow }, x) => BrowserWindow.getAllWindows().find((bw) => bw.webContents.getURL().includes(`/${x}/`))?.destroy(), id);
  await wait(800);
}
// PapelPDF
await l.evaluate(() => window.__launcher.launch('papelpdf'));
const p = await win((u) => u.includes('/papelpdf/')); await wait(2500); await harvest(p, 'papelpdf');
const pdfFile = path.join(out, 'Untitled.pdf');
if (!fs.existsSync(pdfFile)) { const { PDFDocument } = await import('pdf-lib'); const d = await PDFDocument.create(); d.addPage([612, 792]).drawText('Hola mundo', { x: 72, y: 700 }); fs.writeFileSync(pdfFile, await d.save()); }
await app.evaluate(({ BrowserWindow }, f) => BrowserWindow.getAllWindows().find((x) => x.webContents.getURL().includes('/papelpdf/')).webContents.send('open-path', f), pdfFile);
await wait(2500); await harvest(p, 'papel-doc');
for (const sel of await p.evaluate(() => [...document.querySelectorAll('[data-tab]')].map((b) => `[data-tab="${b.dataset.tab}"]`))) {
  await p.click(sel).catch(() => {}); await wait(250); await harvest(p, sel);
  for (const tool of await p.evaluate(() => [...document.querySelectorAll('[data-tool]')].filter((b) => b.offsetParent).map((b) => `[data-tool="${b.dataset.tool}"]`))) { await p.click(tool).catch(() => {}); await wait(150); await harvest(p, tool); }
}
for (const cmd of ['ocr', 'protect', 'compress', 'watermark', 'pageNumbers', 'shortcuts', 'about', 'toImages']) {
  await app.evaluate(({ BrowserWindow }, c) => BrowserWindow.getAllWindows().find((x) => x.webContents.getURL().includes('/papelpdf/')).webContents.send('menu', c), cmd);
  await wait(400); await harvest(p, cmd); await p.keyboard.press('Escape'); await wait(150);
}
fs.writeFileSync(path.join(out, 'faltan.txt'), [...all].sort().join('\n'));
console.log(`${all.size} textos sin traducir -> test/out/i18n/faltan.txt`);
await app.close();
