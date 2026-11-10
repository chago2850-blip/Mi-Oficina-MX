// Capturas de la interfaz en inglés
import { _electron as electron } from 'playwright';
import fs from 'node:fs'; import path from 'node:path'; import { fileURLToPath } from 'node:url';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'en'); fs.rmSync(out, { recursive: true, force: true }); fs.mkdirSync(out, { recursive: true });
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const app = await electron.launch({ args: [root, `--user-data-dir=${path.join(out, 'perfil')}`, '--no-sandbox'], cwd: root, env: { ...process.env, OFICINA_TEST_DIR: out, OFICINA_LANG: 'en' } });
async function win(pred) { for (let i = 0; i < 100; i++) { for (const w of app.windows()) if (pred(w.url())) return w; await wait(150); } throw new Error('ventana'); }
const l = await win((u) => u.includes('/launcher/')); await wait(1200);
await l.evaluate(() => window.__launcher.setOpen(true)); await wait(900); await l.screenshot({ path: path.join(out, 'launcher.png') });
const shots = { escribiente: ["window.__shell.newDoc('reporte')", 'insert'], calcula: ["window.__shell.newDoc('presupuesto')", 'formulas'], impresiona: ["window.__shell.newDoc('negocio')", 'home'] };
for (const [id, [js, tab]] of Object.entries(shots)) {
  await l.evaluate((x) => window.__launcher.launch(x), id);
  const w = await win((u) => u.includes(`/${id}/`)); await w.waitForLoadState('domcontentloaded'); await wait(1500);
  await w.evaluate(js).catch((e) => console.error(e.message)); await wait(800);
  await w.evaluate((t) => window.__shell.ribbon.show(t), tab).catch(() => {}); await wait(300);
  await w.screenshot({ path: path.join(out, `${id}.png`) });
  if (id === 'calcula') { await w.evaluate(() => { setTimeout(() => window.__calcula.onCommand('insertFunction')); }); await wait(500); await w.screenshot({ path: path.join(out, 'calcula-fn.png') }); await w.keyboard.press('Escape'); }
  await w.evaluate(() => window.__shell.showBackstage('new')); await wait(400); await w.screenshot({ path: path.join(out, `${id}-home.png`) });
  await app.evaluate(({ BrowserWindow }, x) => BrowserWindow.getAllWindows().find((b) => b.webContents.getURL().includes(`/${x}/`))?.destroy(), id); await wait(800);
}
await l.evaluate(() => window.__launcher.launch('papelpdf')); const p = await win((u) => u.includes('/papelpdf/')); await wait(2500); await p.screenshot({ path: path.join(out, 'papelpdf.png') });
await app.close();
