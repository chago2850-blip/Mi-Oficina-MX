// Auditoría de accesibilidad con axe-core en cada app (falla si hay violaciones serias o críticas)
import { _electron as electron } from 'playwright';
import path from 'node:path'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'a11y'); fs.mkdirSync(out, { recursive: true });
const axe = fs.readFileSync(path.join(root, 'node_modules', 'axe-core', 'axe.min.js'), 'utf8');
const app = await electron.launch({ args: [root, `--user-data-dir=${path.join(out, 'perfil')}`, '--no-sandbox'], cwd: root, env: { ...process.env, OFICINA_TEST_DIR: out, OFICINA_LANG: process.env.OFICINA_LANG || 'es' } });
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
async function win(pred) { for (let i = 0; i < 150; i++) { const w = app.windows().find((x) => pred(x.url())); if (w) { await w.waitForLoadState('domcontentloaded'); return w; } await wait(100); } throw new Error('ventana'); }
const verbose = process.argv.includes('-v');
let bad = 0;
async function audit(w, label) {
  // axe se inyecta por CDP (la CSP de las apps no permite scripts en línea)
  const cdp = await w.context().newCDPSession(w);
  await cdp.send('Runtime.evaluate', { expression: axe });
  const { result } = await cdp.send('Runtime.evaluate', { expression: `axe.run(document, { resultTypes: ['violations'], rules: { 'color-contrast': { enabled: ${!process.env.NOCONTRAST} } } }).then((r) => JSON.stringify(r.violations.map((v) => ({ id: v.id, impact: v.impact, help: v.help, n: v.nodes.length, nodes: v.nodes.slice(0, 30).map((n) => n.target.join(" ") + " " + (n.any?.[0]?.message || "").slice(0, 110)) }))))`, awaitPromise: true });
  const vs = JSON.parse(result.value);
  const serious = vs.filter((v) => v.impact === 'serious' || v.impact === 'critical');
  console.log(`${serious.length ? '✘' : '✔'} ${label}: ${vs.length} reglas con avisos, ${serious.length} serias`);
  for (const v of verbose ? vs : serious) console.log(`   [${v.impact}] ${v.id} ×${v.n}: ${v.help}\n      ${v.nodes.join('\n      ')}`);
  bad += serious.length;
}
try {
  const l = await win((u) => u.includes('launcher'));
  await audit(l, 'Lanzador');
  for (const id of ['escribiente', 'calcula', 'impresiona']) {
    await l.evaluate((x) => window.__launcher.launch(x), id);
    const w = await win((u) => u.includes(`/${id}/`)); await wait(1500);
    await audit(w, `${id} (inicio)`);
    await w.evaluate(() => window.__shell.newDoc?.() ?? window.__shell.hideBackstage?.()); await wait(800);
    await audit(w, `${id} (documento)`);
  }
  await l.evaluate(() => window.__launcher.launch('papelpdf'));
  const p = await win((u) => u.includes('papelpdf')); await wait(1200);
  await audit(p, 'PapelPDF');
} finally { await app.close(); }
if (bad) { console.log(`\n${bad} problemas serios de accesibilidad`); process.exitCode = 1; }
