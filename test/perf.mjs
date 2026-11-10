// Mide tiempos de carga y dibujo con archivos grandes (no es parte de CI)
import { _electron as electron } from 'playwright';
import path from 'node:path'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test', 'out', 'perf');
const app = await electron.launch({ args: [root, `--user-data-dir=${path.join(out, 'perfil')}`, '--no-sandbox'], cwd: root, env: { ...process.env, OFICINA_TEST_DIR: out } });
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
async function win(pred) { for (let i = 0; i < 150; i++) { const w = app.windows().find((x) => pred(x.url())); if (w) { await w.waitForLoadState('domcontentloaded'); return w; } await wait(100); } throw new Error('ventana'); }
const l = await win((u) => u.includes('launcher'));
const which = process.argv[2] || 'calcula';
try {
  if (which === 'calcula') {
    await l.evaluate(() => window.__launcher.launch('calcula'));
    const w = await win((u) => u.includes('/calcula/')); await w.waitForFunction(() => window.__calcula);
    const csv = path.join(out, 'big.csv');
    const t = await w.evaluate(async (p) => { const t0 = performance.now(); const d = await window.oficina.readPath(p); await window.__shell.loadFile(d); const t1 = performance.now(); return { load: t1 - t0 }; }, csv);
    console.log('CSV 100k load ms', Math.round(t.load));
    const d = await w.evaluate(async () => { const g = window.__calcula.grid; const t0 = performance.now(); for (let i = 0; i < 30; i++) { g.scroller.scrollTop += 2000; g.draw(); } return (performance.now() - t0) / 30; });
    console.log('draw ms', d.toFixed(1));
    if (process.env.PROFILE) {
      const cdp = await w.context().newCDPSession(w); await cdp.send('Profiler.enable'); await cdp.send('Profiler.start');
      await w.evaluate(() => { const g = window.__calcula.grid; for (let i = 0; i < 30; i++) { g.scroller.scrollTop += 2000; g.draw(); } });
      const { profile } = await cdp.send('Profiler.stop');
      const self = new Map(); const dt = profile.timeDeltas; const byId = new Map(profile.nodes.map((n) => [n.id, n]));
      profile.samples.forEach((id, i) => { const n = byId.get(id); const k = `${n.callFrame.functionName || '(anon)'} ${n.callFrame.url.split('/').pop()}:${n.callFrame.lineNumber}`; self.set(k, (self.get(k) || 0) + (dt[i] || 0)); });
      console.log([...self].sort((a, b) => b[1] - a[1]).slice(0, 15).map(([k, v]) => `${(v / 1000).toFixed(0)}ms ${k}`).join('\n'));
      process.exit(0);
    }
    const sv = await w.evaluate(async () => { const t0 = performance.now(); await window.__shell.saveAsFormat('xlsx', 'Excel'); return performance.now() - t0; });
    console.log('save xlsx ms', Math.round(sv), fs.readdirSync(out).join(' '));
    const x = path.join(out, fs.readdirSync(out).find((f) => f.endsWith('.xlsx')));
    const ro = await w.evaluate(async (p) => { const t0 = performance.now(); const d = await window.oficina.readPath(p); await window.__shell.loadFile(d); return performance.now() - t0; }, x);
    console.log('open xlsx ms', Math.round(ro));
    const st = await w.evaluate(async () => { const g = window.__calcula.grid; g.select(0, 3, false, { scroll: false }); g.select(99999, 3, true, { scroll: false }); const t0 = performance.now(); window.__calcula.refresh(); return performance.now() - t0; });
    console.log('refresh (status sum) ms', Math.round(st));
    const so = await w.evaluate(async () => { const t0 = performance.now(); window.__calcula.onCommand('sortDesc'); return performance.now() - t0; });
    console.log('sort ms', Math.round(so));
  } else {
    await l.evaluate(() => window.__launcher.launch('escribiente'));
    const w = await win((u) => u.includes('/escribiente/')); await w.waitForFunction(() => window.__shell);
    const r = await w.evaluate(async () => {
      const ed = document.querySelector('.editor, [contenteditable="true"]');
      let html = ''; for (let i = 0; i < 6000; i++) html += (i % 40 === 0 ? `<h2>Sección ${i}</h2>` : '') + `<p>Párrafo ${i}: Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.</p>`;
      const t0 = performance.now(); ed.innerHTML = html; ed.dispatchEvent(new Event('input', { bubbles: true })); await new Promise((r) => setTimeout(r, 50)); await new Promise((r) => requestAnimationFrame(() => r()));
      const t1 = performance.now();
      return { set: t1 - t0, pages: document.querySelectorAll('.page, .page-sep, .pg').length };
    });
    console.log('escribiente', r);
    await wait(3000);
    if (process.env.PROFILE) {
      const cdp = await w.context().newCDPSession(w); await cdp.send('Profiler.enable'); await cdp.send('Profiler.start');
      await w.evaluate(async () => { const t0 = performance.now(); const ed = document.querySelector('.editor, [contenteditable="true"]'); ed.focus(); for (let i = 0; i < 20; i++) { document.execCommand('insertText', false, 'x'); await new Promise((r) => requestAnimationFrame(r)); } await new Promise((r) => setTimeout(r, 1500)); });
      const { profile } = await cdp.send('Profiler.stop');
      const self = new Map(); const dt = profile.timeDeltas; const byId = new Map(profile.nodes.map((n) => [n.id, n]));
      profile.samples.forEach((id, i) => { const n = byId.get(id); const k = `${n.callFrame.functionName || '(anon)'} ${n.callFrame.url.split('/').pop()}:${n.callFrame.lineNumber}`; self.set(k, (self.get(k) || 0) + (dt[i] || 0)); });
      console.log([...self].sort((a, b) => b[1] - a[1]).slice(0, 15).map(([k, v]) => `${(v / 1000).toFixed(0)}ms ${k}`).join('\n'));
      process.exit(0);
    }
    const typ = await w.evaluate(async () => { const t0 = performance.now(); for (let i = 0; i < 20; i++) { document.execCommand('insertText', false, 'x'); await new Promise((r) => requestAnimationFrame(r)); } return (performance.now() - t0) / 20; });
    console.log('type ms/char', typ.toFixed(1));
    const sv = await w.evaluate(async () => { const t0 = performance.now(); await window.__shell.save(); return performance.now() - t0; });
    console.log('save docx ms', Math.round(sv));
  }
} finally { await app.close(); }
