// Familias de letras instaladas en el sistema operativo (macOS, Windows y Linux), para que las
// apps de la suite las ofrezcan junto a las letras libres incluidas. Se guardan en caché en la
// carpeta de datos y se actualizan en segundo plano. GPL-3.0
const { execFile } = require('child_process');
const path = require('path');
const fs = require('fs/promises');

const run = (cmd, args, timeout = 15000) => new Promise((resolve) => {
  execFile(cmd, args, { timeout, maxBuffer: 16 * 1024 * 1024, windowsHide: true, encoding: 'utf8' }, (err, out) => resolve(err ? '' : String(out || '')));
});

const clean = (list) => {
  const seen = new Map();
  for (let n of list) {
    n = String(n || '').trim().replace(/^["']|["']$/g, '');
    // se ocultan las familias internas del sistema (.SF NS, .AppleSystemUIFont, etc.)
    if (!n || n.startsWith('.') || n.startsWith('@') || n.length > 64 || /[\u0000-\u001f]/.test(n)) continue;
    const k = n.toLowerCase(); if (!seen.has(k)) seen.set(k, n);
  }
  return [...seen.values()].sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));
};

async function listNow() {
  if (process.platform === 'darwin') {
    // NSFontManager devuelve los nombres de familia tal como los muestra el Selector de tipos de letra
    const js = "ObjC.import('AppKit'); JSON.stringify(ObjC.deepUnwrap($.NSFontManager.sharedFontManager.availableFontFamilies))";
    const out = await run('/usr/bin/osascript', ['-l', 'JavaScript', '-e', js]);
    try { const a = JSON.parse(out.trim()); if (Array.isArray(a) && a.length) return clean(a); } catch { /* sigue */ }
    const sp = await run('/usr/sbin/system_profiler', ['SPFontsDataType', '-json'], 60000);
    try {
      const j = JSON.parse(sp); const fam = [];
      for (const f of j.SPFontsDataType || []) for (const t of f.typefaces || []) if (t.family) fam.push(t.family);
      return clean(fam);
    } catch { return []; }
  }
  if (process.platform === 'win32') {
    const ps = '[Console]::OutputEncoding=[Text.Encoding]::UTF8; Add-Type -AssemblyName System.Drawing; (New-Object System.Drawing.Text.InstalledFontCollection).Families | ForEach-Object { $_.Name }';
    const out = await run('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', ps]);
    if (out.trim()) return clean(out.split(/\r?\n/));
    // respaldo: el registro (nombres como «Arial Bold (TrueType)»)
    const reg = (await run('reg', ['query', 'HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts'])) + (await run('reg', ['query', 'HKCU\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts']));
    return clean(reg.split(/\r?\n/).map((l) => /^\s{4}(.+?)\s+\(.*?\)\s+REG_SZ/.exec(l)?.[1]).filter(Boolean)
      .map((n) => n.split(' & ')[0].replace(/\s+(Bold|Italic|Light|Semibold|SemiBold|Black|Regular|Oblique|Medium|Thin|Condensed|Narrow)(\s.*)?$/i, '')));
  }
  // Linux y otros: fontconfig
  const out = await run('fc-list', ['--format', '%{family[0]}\\n']);
  return clean(out.split('\n'));
}

let cache = null; let pending = null;
const file = (dir) => path.join(dir, 'fuentes-sistema.json');

// devuelve la lista (la de caché al instante si existe y la actualiza en segundo plano)
async function systemFonts(dataDir) {
  if (process.env.OFICINA_NO_SYSFONTS) return [];
  if (cache) return cache;
  let disk = null;
  try { const j = JSON.parse(await fs.readFile(file(dataDir), 'utf8')); if (Array.isArray(j.fonts)) disk = j; } catch { /* sin caché */ }
  const refresh = () => {
    if (!pending) {
      pending = listNow().then(async (fonts) => {
        if (fonts.length) { cache = fonts; try { await fs.writeFile(file(dataDir), JSON.stringify({ at: Date.now(), fonts })); } catch { /* nada */ } }
        return cache || fonts;
      }).catch(() => cache || []).finally(() => { pending = null; });
    }
    return pending;
  };
  if (disk) { cache = disk.fonts; if (Date.now() - (disk.at || 0) > 6 * 3600e3) refresh(); return cache; }
  return refresh();
}

module.exports = { systemFonts, listNow };
