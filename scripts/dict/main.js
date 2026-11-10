// Descarga los diccionarios de ortografía (Hunspell .bdic de Chromium) para incluirlos en el instalador
// y que la revisión ortográfica funcione sin internet. Se ejecuta en CI: electron scripts/dict
// Salida: build/dict/*.bdic. GPL-3.0
const { app, BrowserWindow, session } = require('electron');
const fs = require('fs'); const path = require('path'); const os = require('os');
const LANGS = ['es-MX', 'es', 'en-US'];
const outDir = path.resolve(__dirname, '..', '..', 'build', 'dict');
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'oficina-dict-'));
app.setPath('userData', tmp);
app.commandLine.appendSwitch('no-sandbox');
app.whenReady().then(async () => {
  const ses = session.defaultSession;
  const langs = LANGS.filter((l) => ses.availableSpellCheckerLanguages.includes(l));
  const pending = new Set(langs); let failed = false;
  const done = () => {
    const dir = path.join(tmp, 'Dictionaries');
    fs.mkdirSync(outDir, { recursive: true });
    const files = fs.existsSync(dir) ? fs.readdirSync(dir).filter((f) => f.endsWith('.bdic')) : [];
    for (const f of files) fs.copyFileSync(path.join(dir, f), path.join(outDir, f));
    console.log(`[dict] ${files.length ? files.join(', ') : 'ninguno'} → ${outDir}`);
    app.exit(failed || !files.length ? 1 : 0);
  };
  ses.on('spellcheck-dictionary-download-success', (_e, l) => { console.log(`[dict] ${l} listo`); pending.delete(l); if (!pending.size) setTimeout(done, 1000); });
  ses.on('spellcheck-dictionary-download-failure', (_e, l) => { console.log(`[dict] ${l} falló`); failed = true; pending.delete(l); if (!pending.size) setTimeout(done, 500); });
  ses.on('spellcheck-dictionary-initialized', (_e, l) => { pending.delete(l); if (!pending.size) setTimeout(done, 1000); });
  ses.setSpellCheckerLanguages(langs);
  const w = new BrowserWindow({ show: false, webPreferences: { spellcheck: true } });
  await w.loadURL('data:text/html,<textarea autofocus>holaa mundoo helo wrld</textarea>');
  setTimeout(() => { console.log('[dict] tiempo agotado'); failed = true; done(); }, 60000);
});
