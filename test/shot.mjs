// Uso: node test/shot.mjs <escena> — captura pantallas para revisar el diseño.
import { _electron as electron } from 'playwright';
import fs from 'node:fs';
const out = new URL('./out/', import.meta.url).pathname;
fs.mkdirSync(out, { recursive: true });
const testDir = out + 'files/'; fs.mkdirSync(testDir, { recursive: true });
const app = await electron.launch({ args: ['.'], env: { ...process.env, OFICINA_TEST_DIR: testDir, ELECTRON_DISABLE_GPU: '1' } });
const scene = process.argv[2] || 'launcher';
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const win = await app.firstWindow();
await wait(1200);
const cap = async (page, name) => { const p = out + name + '.png'; await page.screenshot({ path: p }); console.log(p); };
if (scene === 'launcher') {
  await cap(win, 'launcher-cerrado');
  await win.evaluate(() => window.__launcher.setOpen(true));
  await wait(900);
  await cap(win, 'launcher-abierto');
  await win.hover('[data-app="calcula"]');
  await wait(500);
  await cap(win, 'launcher-hover');
} else {
  await win.evaluate((id) => window.__launcher.launch(id), scene);
  const w = await app.waitForEvent('window');
  await w.waitForLoadState('domcontentloaded');
  await wait(1500);
  if (process.argv[3]) await w.evaluate(process.argv[3]);
  await wait(800);
  await cap(w, scene);
}
await app.close();
