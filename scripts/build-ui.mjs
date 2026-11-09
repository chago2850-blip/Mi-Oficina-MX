// Empaqueta el código de interfaz de cada app (apps/<app>/src/main.js -> apps/<app>/bundle.js)
// con esbuild, para que las bibliotecas de formatos funcionen sin Node dentro de la ventana.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
let esbuild;
try { esbuild = await import('esbuild'); } catch { console.log('[ui] esbuild no está instalado (instalación de producción): se omite.'); process.exit(0); }

const apps = fs.readdirSync(path.join(root, 'apps')).filter((a) => fs.existsSync(path.join(root, 'apps', a, 'src', 'main.js')));
await Promise.all(apps.map((a) => esbuild.build({
  entryPoints: [path.join(root, 'apps', a, 'src', 'main.js')],
  outfile: path.join(root, 'apps', a, 'bundle.js'),
  bundle: true,
  format: 'iife',
  platform: 'browser',
  target: 'chrome128',
  minify: process.env.NODE_ENV === 'production' || process.argv.includes('--min'),
  sourcemap: false,
  legalComments: 'eof',
  define: { 'process.env.NODE_ENV': '"production"', global: 'globalThis' },
  logLevel: 'warning',
})));
console.log('[ui] listo:', apps.join(', ') || '(ninguna)');
