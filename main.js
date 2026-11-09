// Mi Oficina MX — proceso principal de Electron
// Copyright (C) 2026 Santiago Miguel <chago2850@gmail.com>
//
// Este programa es software libre: puedes redistribuirlo y/o modificarlo bajo los
// términos de la Licencia Pública General GNU publicada por la Free Software
// Foundation, ya sea la versión 3 de la Licencia o (a tu elección) cualquier
// versión posterior. Se distribuye SIN NINGUNA GARANTÍA. Ver el archivo LICENSE.
//
// Sin telemetría, sin cuentas, sin red.

const { app, BrowserWindow, ipcMain, dialog, Menu, shell, protocol, Tray, nativeImage, screen } = require('electron');
const path = require('path');
const fs = require('fs/promises');
const fss = require('fs');

const isMac = process.platform === 'darwin';
const ROOT = __dirname;
const TEST_DIR = process.env.OFICINA_TEST_DIR || null;
if (TEST_DIR) process.env.PAPELPDF_TEST_DIR = TEST_DIR;

const papel = require('./main/papelpdf');

// ---------------------------------------------------------------- apps
const APPS = {
  escribiente: {
    name: 'Escribiente', color: '#5B48E8', bg: '#F3F1FB',
    exts: ['docx', 'odt', 'txt', 'html', 'htm', 'md'],
    saveExt: 'docx', saveName: 'Documento de texto',
  },
  calcula: {
    name: 'Calcula Datos', color: '#0E9C74', bg: '#F0F8F5',
    exts: ['xlsx', 'xlsm', 'ods', 'csv', 'tsv'],
    saveExt: 'xlsx', saveName: 'Libro de hojas de cálculo',
  },
  impresiona: {
    name: 'Impresiona', color: '#DC3A63', bg: '#FBF1F3',
    exts: ['pptx', 'odp'],
    saveExt: 'pptx', saveName: 'Presentación',
  },
  papelpdf: { name: 'PapelPDF', exts: ['pdf'] },
};
const appForFile = (p) => {
  const ext = path.extname(p).slice(1).toLowerCase();
  return Object.keys(APPS).find((k) => APPS[k].exts.includes(ext)) || null;
};

// ---------------------------------------------------------------- protocolo app://
protocol.registerSchemesAsPrivileged([
  { scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true, stream: true } },
]);
const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.mjs': 'text/javascript', '.css': 'text/css',
  '.json': 'application/json', '.wasm': 'application/wasm', '.png': 'image/png', '.svg': 'image/svg+xml',
  '.gz': 'application/gzip', '.bcmap': 'application/octet-stream', '.pfb': 'application/octet-stream',
  '.ttf': 'font/ttf', '.woff2': 'font/woff2', '.jpg': 'image/jpeg',
};
function registerAppProtocol() {
  protocol.handle('app', async (req) => {
    const url = new URL(req.url);
    const rel = decodeURIComponent(url.pathname).replace(/^\/+/, '');
    const file = path.normalize(path.join(ROOT, rel));
    if (!file.startsWith(ROOT + path.sep)) return new Response('Prohibido', { status: 403 });
    try {
      const data = await fs.readFile(file);
      return new Response(data, { headers: { 'content-type': MIME[path.extname(file).toLowerCase()] || 'application/octet-stream' } });
    } catch {
      return new Response('No encontrado', { status: 404 });
    }
  });
}

// ---------------------------------------------------------------- archivos pasados por el sistema
const fileArgs = (argv) => argv.slice(app.isPackaged ? 1 : 2).filter((a) => !a.startsWith('-') && appForFile(a) && fss.existsSync(a));
let pendingFiles = fileArgs(process.argv);
let ready = false;

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', (_e, argv) => {
    const files = fileArgs(argv);
    if (files.length) files.forEach((f) => openApp(appForFile(f), f));
    else showLauncher();
  });
}
app.on('open-file', (e, file) => {
  e.preventDefault();
  if (ready) openApp(appForFile(file), file); else pendingFiles.push(file);
});

// ---------------------------------------------------------------- Launcher (dona)
let launcher = null;
let tray = null;
let quitting = false;

function createLauncher({ minimized = false } = {}) {
  const size = 440;
  const { workArea } = screen.getPrimaryDisplay();
  launcher = new BrowserWindow({
    width: size, height: size + 60,
    x: Math.round(workArea.x + (workArea.width - size) / 2),
    y: Math.round(workArea.y + (workArea.height - size) / 2),
    frame: false, transparent: true, resizable: false, maximizable: false, fullscreenable: false,
    hasShadow: false, backgroundColor: '#00000000', show: false, skipTaskbar: false,
    title: 'Mi Oficina MX',
    icon: path.join(ROOT, 'build', 'icon.png'),
    webPreferences: {
      preload: path.join(ROOT, 'preload', 'launcher.js'),
      contextIsolation: true, nodeIntegration: false, sandbox: true,
    },
  });
  launcher.loadURL('app://suite/apps/launcher/index.html');
  launcher.once('ready-to-show', () => { if (minimized) launcher.minimize(); else launcher.show(); });
  launcher.on('close', async (e) => {
    if (quitting) return;
    // Cerrar la dona cierra la suite: primero las apps (cada una pregunta si hay cambios sin guardar).
    e.preventDefault();
    await quitSuite();
  });
  launcher.on('closed', () => { launcher = null; });
}

function showLauncher() {
  if (!launcher) { createLauncher(); return; }
  if (launcher.isMinimized()) launcher.restore();
  launcher.show();
  launcher.focus();
  launcher.webContents.send('launcher:restored');
}

function minimizeLauncher() {
  if (!launcher || launcher.isDestroyed()) return;
  launcher.webContents.send('launcher:minimizing');
  // En Linux algunos gestores de ventanas no minimizan ventanas sin marco: se oculta y queda la bandeja.
  launcher.minimize();
  // Sin gestor de ventanas que minimice (algunos escritorios Linux), se oculta y la bandeja o el menú la vuelven a mostrar.
  setTimeout(() => { if (launcher && !launcher.isDestroyed() && !launcher.isMinimized() && appWindows().length) launcher.hide(); }, 400);
}

function appWindows() {
  return BrowserWindow.getAllWindows().filter((w) => w.__app && !w.isDestroyed());
}

function onAppWindowClosed() {
  // La dona vuelve cuando se cierra la última app.
  setTimeout(() => {
    if (quitting) return;
    if (appWindows().length === 0) showLauncher();
    broadcastRunning();
  }, 50);
}

function broadcastRunning() {
  if (!launcher || launcher.isDestroyed()) return;
  const running = [...new Set(appWindows().map((w) => w.__app))];
  launcher.webContents.send('launcher:running', running);
}

async function quitSuite() {
  const wins = appWindows();
  for (const w of wins) {
    if (w.isDestroyed()) continue;
    const closed = new Promise((res) => w.once('closed', () => res(true)));
    w.close();
    // Si el usuario cancela en el diálogo de cambios sin guardar, la ventana sigue viva.
    const ok = await Promise.race([closed, new Promise((res) => setTimeout(() => res(false), 60000))]);
    if (!ok || !w.isDestroyed()) return false;
  }
  quitting = true;
  app.quit();
  return true;
}

function createTray() {
  try {
    const img = nativeImage.createFromPath(path.join(ROOT, 'build', 'icons', isMac ? '16x16.png' : '32x32.png'));
    tray = new Tray(img);
    tray.setToolTip('Mi Oficina MX');
    const menu = Menu.buildFromTemplate([
      { label: 'Mostrar Mi Oficina MX', click: () => showLauncher() },
      { type: 'separator' },
      { label: 'Escribiente', click: () => openApp('escribiente') },
      { label: 'Calcula Datos', click: () => openApp('calcula') },
      { label: 'Impresiona', click: () => openApp('impresiona') },
      { label: 'PapelPDF', click: () => openApp('papelpdf') },
      { type: 'separator' },
      { label: 'Salir', click: () => quitSuite() },
    ]);
    tray.setContextMenu(menu);
    tray.on('click', () => showLauncher());
  } catch { tray = null; }
}

// ---------------------------------------------------------------- ventanas de las apps
function openApp(id, filePath) {
  if (!id || !APPS[id]) return null;
  let win;
  if (id === 'papelpdf') {
    win = papel.openPath(filePath, { onClosed: onAppWindowClosed });
  } else if (!filePath) {
    // Desde la dona: si la app ya está abierta, se enfoca en vez de abrir otra.
    const existing = appWindows().find((w) => w.__app === id);
    if (existing) {
      if (existing.isMinimized()) existing.restore();
      existing.focus();
      win = existing;
    } else win = createOfficeWindow(id, null);
  } else {
    win = createOfficeWindow(id, filePath);
  }
  minimizeLauncher();
  setTimeout(broadcastRunning, 200);
  return win;
}

function createOfficeWindow(id, filePath) {
  const meta = APPS[id];
  const { workArea } = screen.getPrimaryDisplay();
  const win = new BrowserWindow({
    width: Math.min(1400, workArea.width), height: Math.min(900, workArea.height),
    minWidth: 860, minHeight: 560,
    title: meta.name, backgroundColor: meta.bg, show: false,
    icon: path.join(ROOT, 'build', 'apps', `${id}.png`),
    webPreferences: {
      preload: path.join(ROOT, 'preload', 'office.js'),
      contextIsolation: true, nodeIntegration: false, sandbox: true, spellcheck: true,
    },
  });
  win.__app = id;
  win.__dirty = false;
  win.__filePath = filePath || null;
  const menu = buildOfficeMenu(win, id);
  if (isMac) { win.on('focus', () => Menu.setApplicationMenu(menu)); Menu.setApplicationMenu(menu); } else win.setMenu(menu);
  win.webContents.session.setSpellCheckerLanguages?.(['es-MX', 'es', 'en-US'].filter((l) => win.webContents.session.availableSpellCheckerLanguages?.includes(l)));
  win.loadURL(`app://suite/apps/${id}/index.html`);
  win.once('ready-to-show', () => { win.show(); if (TEST_DIR) win.blur(); });
  win.webContents.on('did-finish-load', () => { if (filePath) win.webContents.send('open-path', filePath); });
  win.webContents.setWindowOpenHandler(({ url }) => { if (/^https?:|^mailto:/.test(url)) shell.openExternal(url); return { action: 'deny' }; });
  win.webContents.on('will-navigate', (e, url) => { if (!url.startsWith('app://')) e.preventDefault(); });
  win.webContents.on('context-menu', (_e, params) => editContextMenu(win, params));
  win.on('close', async (e) => {
    if (win.__forceClose) return;
    if (!win.__dirty || TEST_DIR) { cleanupRecovery(win); return; }
    e.preventDefault();
    const { response } = await dialog.showMessageBox(win, {
      type: 'warning', buttons: ['Guardar', 'No guardar', 'Cancelar'], defaultId: 0, cancelId: 2,
      message: `¿Guardar los cambios en "${win.__docName || 'Sin título'}"?`,
      detail: 'Si no los guardas, se perderán.',
    });
    if (response === 2) return;
    if (response === 0) {
      const saved = await win.webContents.executeJavaScript('window.__oficinaSave ? window.__oficinaSave() : false').catch(() => false);
      if (!saved) return;
    }
    cleanupRecovery(win);
    win.__forceClose = true;
    win.close();
  });
  win.on('closed', onAppWindowClosed);
  return win;
}

function editContextMenu(win, p) {
  const items = [];
  if (p.misspelledWord) {
    p.dictionarySuggestions.slice(0, 6).forEach((s) => items.push({ label: s, click: () => win.webContents.replaceMisspelling(s) }));
    if (!p.dictionarySuggestions.length) items.push({ label: 'Sin sugerencias', enabled: false });
    items.push({ label: 'Agregar al diccionario', click: () => win.webContents.session.addWordToSpellCheckerDictionary(p.misspelledWord) });
    items.push({ type: 'separator' });
  }
  if (p.isEditable || p.selectionText) {
    items.push(
      { role: 'cut', label: 'Cortar', enabled: p.editFlags.canCut },
      { role: 'copy', label: 'Copiar', enabled: p.editFlags.canCopy },
      { role: 'paste', label: 'Pegar', enabled: p.editFlags.canPaste },
      { role: 'pasteAndMatchStyle', label: 'Pegar solo texto', enabled: p.editFlags.canPaste },
    );
  }
  if (items.length) Menu.buildFromTemplate(items).popup({ window: win });
}

function buildOfficeMenu(win, id) {
  const send = (cmd) => { if (!win.isDestroyed()) win.webContents.send('menu', cmd); };
  const meta = APPS[id];
  const extra = {
    escribiente: [
      { label: 'Insertar', submenu: [
        { label: 'Tabla…', click: () => send('insertTable') },
        { label: 'Imagen…', click: () => send('insertImage') },
        { label: 'Vínculo…', accelerator: 'CmdOrCtrl+K', click: () => send('insertLink') },
        { label: 'Salto de página', accelerator: 'CmdOrCtrl+Return', click: () => send('pageBreak') },
      ] },
      { label: 'Formato', submenu: [
        { label: 'Negrita', accelerator: 'CmdOrCtrl+B', click: () => send('bold') },
        { label: 'Cursiva', accelerator: 'CmdOrCtrl+I', click: () => send('italic') },
        { label: 'Subrayado', accelerator: 'CmdOrCtrl+U', click: () => send('underline') },
        { type: 'separator' },
        { label: 'Configurar página…', click: () => send('pageSetup') },
      ] },
    ],
    calcula: [
      { label: 'Insertar', submenu: [
        { label: 'Gráfico…', click: () => send('insertChart') },
        { label: 'Función…', accelerator: 'Shift+F3', click: () => send('insertFunction') },
        { label: 'Hoja nueva', accelerator: 'Shift+F11', click: () => send('newSheet') },
      ] },
      { label: 'Datos', submenu: [
        { label: 'Ordenar A → Z', click: () => send('sortAsc') },
        { label: 'Ordenar Z → A', click: () => send('sortDesc') },
        { label: 'Filtro', accelerator: 'CmdOrCtrl+Shift+L', click: () => send('filter') },
      ] },
    ],
    impresiona: [
      { label: 'Diapositiva', submenu: [
        { label: 'Nueva diapositiva', accelerator: 'CmdOrCtrl+M', click: () => send('newSlide') },
        { label: 'Duplicar diapositiva', accelerator: 'CmdOrCtrl+D', click: () => send('duplicateSlide') },
        { type: 'separator' },
        { label: 'Presentar desde el principio', accelerator: 'F5', click: () => send('present') },
        { label: 'Presentar desde la actual', accelerator: 'Shift+F5', click: () => send('presentCurrent') },
      ] },
    ],
  }[id] || [];
  return Menu.buildFromTemplate([
    ...(isMac ? [{ role: 'appMenu', label: 'Mi Oficina MX' }] : []),
    { label: 'Archivo', submenu: [
      { label: 'Nuevo', accelerator: 'CmdOrCtrl+N', click: () => send('new') },
      { label: 'Abrir…', accelerator: 'CmdOrCtrl+O', click: () => send('open') },
      { type: 'separator' },
      { label: 'Guardar', accelerator: 'CmdOrCtrl+S', click: () => send('save') },
      { label: 'Guardar como…', accelerator: 'CmdOrCtrl+Shift+S', click: () => send('saveAs') },
      { label: 'Exportar a PDF…', accelerator: 'CmdOrCtrl+Shift+E', click: () => send('exportPdf') },
      { label: 'Imprimir…', accelerator: 'CmdOrCtrl+P', click: () => send('print') },
      { type: 'separator' },
      { label: 'Propiedades…', click: () => send('properties') },
      { label: 'Opciones…', click: () => send('options') },
      { type: 'separator' },
      { label: 'Mostrar Mi Oficina MX', click: () => showLauncher() },
      { label: 'Cerrar ventana', accelerator: 'CmdOrCtrl+W', click: () => win.close() },
    ] },
    { label: 'Editar', submenu: [
      { label: 'Deshacer', accelerator: 'CmdOrCtrl+Z', click: () => send('undo') },
      { label: 'Rehacer', accelerator: isMac ? 'Cmd+Shift+Z' : 'Ctrl+Y', click: () => send('redo') },
      { type: 'separator' },
      { role: 'cut', label: 'Cortar' }, { role: 'copy', label: 'Copiar' }, { role: 'paste', label: 'Pegar' },
      { label: 'Seleccionar todo', accelerator: 'CmdOrCtrl+A', click: () => send('selectAll') },
      { type: 'separator' },
      { label: 'Buscar…', accelerator: 'CmdOrCtrl+F', click: () => send('find') },
      { label: 'Reemplazar…', accelerator: 'CmdOrCtrl+H', click: () => send('replace') },
    ] },
    ...extra,
    { label: 'Ver', submenu: [
      { label: 'Acercar', accelerator: 'CmdOrCtrl+=', click: () => send('zoomIn') },
      { label: 'Alejar', accelerator: 'CmdOrCtrl+-', click: () => send('zoomOut') },
      { label: 'Tamaño real', accelerator: 'CmdOrCtrl+0', click: () => send('zoomReset') },
      { type: 'separator' },
      { role: 'togglefullscreen', label: 'Pantalla completa' },
      { role: 'toggleDevTools', label: 'Herramientas de desarrollo' },
    ] },
    { label: 'Ayuda', submenu: [
      { label: `Acerca de ${meta.name}`, click: () => send('about') },
      { label: 'Código fuente (GitHub)', click: () => shell.openExternal('https://github.com/chago2850-blip/Mi-Oficina-MX') },
    ] },
  ]);
}

// ---------------------------------------------------------------- preferencias de la suite
const DEFAULT_PREFS = { theme: 'auto', author: '', autosaveSec: 20, defaultFormat: 'office', spellcheck: true };
let prefs = { ...DEFAULT_PREFS };
const prefsFile = () => path.join(app.getPath('userData'), 'preferencias.json');
function loadPrefs() { try { prefs = { ...DEFAULT_PREFS, ...JSON.parse(fss.readFileSync(prefsFile(), 'utf8')) }; } catch { prefs = { ...DEFAULT_PREFS }; } }
function savePrefs(p) {
  prefs = { ...prefs, ...p };
  try { fss.mkdirSync(path.dirname(prefsFile()), { recursive: true }); fss.writeFileSync(prefsFile(), JSON.stringify(prefs, null, 2)); } catch { /* sin permiso */ }
  const { nativeTheme } = require('electron');
  nativeTheme.themeSource = prefs.theme === 'auto' ? 'system' : prefs.theme;
  for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed()) w.webContents.send('prefs', prefs);
}
ipcMain.handle('of:prefs', () => prefs);
ipcMain.handle('of:set-prefs', (_e, p) => { savePrefs(p); return prefs; });

// ---------------------------------------------------------------- recientes y recuperación
const dataDir = () => app.getPath('userData');
const recentsFile = (id) => path.join(dataDir(), `recientes-${id}.json`);
async function getRecents(id) {
  try {
    const list = JSON.parse(await fs.readFile(recentsFile(id), 'utf8'));
    return list.filter((r) => fss.existsSync(r.path));
  } catch { return []; }
}
async function addRecent(id, p) {
  if (!p || !APPS[id]) return;
  const list = (await getRecents(id)).filter((r) => r.path !== p);
  list.unshift({ path: p, name: path.basename(p), openedAt: Date.now() });
  await fs.mkdir(dataDir(), { recursive: true });
  await fs.writeFile(recentsFile(id), JSON.stringify(list.slice(0, 15), null, 2));
  if (!process.env.OFICINA_TEST_DIR && (isMac || process.platform === 'win32')) app.addRecentDocument(p);
}
const recoveryDir = () => path.join(dataDir(), 'recuperacion');
function cleanupRecovery(win) {
  if (win.__recoveryId) fs.rm(path.join(recoveryDir(), `${win.__recoveryId}.json`), { force: true }).catch(() => {});
}

// ---------------------------------------------------------------- IPC del Launcher
ipcMain.handle('launcher:open', (_e, id) => { openApp(id); return true; });
ipcMain.handle('launcher:quit', () => quitSuite());
ipcMain.handle('launcher:hide', () => { minimizeLauncher(); return true; });
ipcMain.handle('launcher:info', () => ({ version: app.getVersion(), platform: process.platform }));
ipcMain.handle('launcher:menu', (e) => {
  const w = BrowserWindow.fromWebContents(e.sender);
  Menu.buildFromTemplate([
    { label: 'Preferencias…', click: () => e.sender.send('launcher:prefs') },
    { label: 'Acerca de Mi Oficina MX', click: () => e.sender.send('launcher:about') },
    { label: 'Código fuente (GitHub)', click: () => shell.openExternal('https://github.com/chago2850-blip/Mi-Oficina-MX') },
    { label: 'Licencia GPL-3.0', click: () => shell.openExternal('https://www.gnu.org/licenses/gpl-3.0.html') },
    { type: 'separator' },
    { label: 'Minimizar', click: () => minimizeLauncher() },
    { label: 'Salir', click: () => quitSuite() },
  ]).popup({ window: w });
  return true;
});

// ---------------------------------------------------------------- IPC de las apps de oficina
const winOf = (e) => BrowserWindow.fromWebContents(e.sender);

async function readDoc(id, p) {
  const buf = await fs.readFile(p);
  await addRecent(id, p);
  return { path: p, name: path.basename(p), ext: path.extname(p).slice(1).toLowerCase(), bytes: new Uint8Array(buf) };
}

ipcMain.handle('of:open-dialog', async (e, { filters, title } = {}) => {
  const w = winOf(e);
  const id = w.__app;
  const r = await dialog.showOpenDialog(w, {
    title: title || `Abrir — ${APPS[id].name}`,
    properties: ['openFile'],
    filters: filters || [{ name: APPS[id].saveName, extensions: APPS[id].exts }, { name: 'Todos los archivos', extensions: ['*'] }],
  });
  if (r.canceled || !r.filePaths[0]) return null;
  const p = r.filePaths[0];
  const owner = appForFile(p);
  if (owner && owner !== id) { openApp(owner, p); return { redirected: owner }; }
  return readDoc(id, p);
});

ipcMain.handle('of:open-image', async (e) => {
  const r = await dialog.showOpenDialog(winOf(e), {
    title: 'Insertar imagen', properties: ['openFile'],
    filters: [{ name: 'Imágenes', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg'] }],
  });
  if (r.canceled || !r.filePaths[0]) return null;
  const p = r.filePaths[0];
  const ext = path.extname(p).slice(1).toLowerCase();
  const mime = { jpg: 'image/jpeg', jpeg: 'image/jpeg', svg: 'image/svg+xml' }[ext] || `image/${ext}`;
  return { name: path.basename(p), mime, bytes: new Uint8Array(await fs.readFile(p)) };
});

ipcMain.handle('of:read-path', async (e, p) => readDoc(winOf(e).__app, p));

ipcMain.handle('of:save', async (e, { bytes, text, path: target, ask, suggestedName, ext, filterName }) => {
  const w = winOf(e);
  const id = w.__app;
  let out = target;
  const useExt = ext || APPS[id].saveExt;
  if (TEST_DIR && (ask || !out)) out = path.join(TEST_DIR, suggestedName || `salida.${useExt}`);
  else if (ask || !out) {
    const r = await dialog.showSaveDialog(w, {
      title: 'Guardar como',
      defaultPath: suggestedName || `Sin título.${useExt}`,
      filters: [{ name: filterName || APPS[id].saveName, extensions: [useExt] }],
    });
    if (r.canceled || !r.filePath) return null;
    out = r.filePath;
    if (!path.extname(out)) out += `.${useExt}`;
  }
  await fs.writeFile(out, text != null ? text : Buffer.from(bytes));
  if (APPS[id].exts.includes(path.extname(out).slice(1).toLowerCase())) await addRecent(id, out);
  return { path: out, name: path.basename(out) };
});

ipcMain.handle('of:export-pdf', async (e, { suggestedName, landscape = false, pageSize = 'Letter', open = false } = {}) => {
  const w = winOf(e);
  let out;
  if (TEST_DIR) out = path.join(TEST_DIR, suggestedName || 'salida.pdf');
  else {
    const r = await dialog.showSaveDialog(w, {
      title: 'Exportar a PDF', defaultPath: suggestedName || 'documento.pdf',
      filters: [{ name: 'Documento PDF', extensions: ['pdf'] }],
    });
    if (r.canceled || !r.filePath) return null;
    out = r.filePath;
  }
  const data = await e.sender.printToPDF({
    printBackground: true, landscape, pageSize, preferCSSPageSize: true,
    margins: { marginType: 'none' },
  });
  await fs.writeFile(out, data);
  if (open) openApp('papelpdf', out);
  return { path: out, name: path.basename(out) };
});

ipcMain.handle('of:print', (e) => new Promise((resolve) => {
  e.sender.print({ printBackground: true }, (ok, reason) => resolve({ ok, reason }));
}));

ipcMain.handle('of:recents', async (e) => getRecents(winOf(e).__app));
ipcMain.handle('of:clear-recents', async (e) => { await fs.rm(recentsFile(winOf(e).__app), { force: true }); return []; });

ipcMain.handle('of:confirm', async (e, { message, detail, buttons }) => {
  if (TEST_DIR) return 0;
  const { response } = await dialog.showMessageBox(winOf(e), {
    type: 'question', message, detail, buttons, defaultId: 0, cancelId: buttons.length - 1,
  });
  return response;
});

ipcMain.handle('of:state', (e, { dirty, name, path: p }) => {
  const w = winOf(e);
  if (dirty !== undefined) { w.__dirty = !!dirty; if (isMac) w.setDocumentEdited(!!dirty); }
  if (name !== undefined) w.__docName = name;
  if (p !== undefined) { w.__filePath = p; if (isMac && p) w.setRepresentedFilename(p); }
  const title = `${w.__dirty ? '● ' : ''}${w.__docName || 'Sin título'} — ${APPS[w.__app].name}`;
  w.setTitle(title);
  return true;
});

ipcMain.handle('of:new-window', (e, { app: id, path: p } = {}) => {
  const target = id || winOf(e).__app;
  if (p) openApp(appForFile(p) || target, p); else createOfficeWindow(target, null);
  return true;
});

ipcMain.handle('of:open-in-suite', (_e, p) => { openApp(appForFile(p), p); return true; });
ipcMain.handle('of:show-in-folder', (_e, p) => { shell.showItemInFolder(p); return true; });
ipcMain.handle('of:show-launcher', () => { showLauncher(); return true; });
ipcMain.handle('of:app-info', (e) => ({
  version: app.getVersion(), platform: process.platform, arch: process.arch, app: winOf(e).__app,
  electron: process.versions.electron,
}));

// Recuperación automática
ipcMain.handle('of:autosave', async (e, { id, name, data }) => {
  const w = winOf(e);
  w.__recoveryId = id;
  await fs.mkdir(recoveryDir(), { recursive: true });
  await fs.writeFile(path.join(recoveryDir(), `${id}.json`), JSON.stringify({ app: w.__app, name, savedAt: Date.now(), path: w.__filePath, data }));
  return true;
});
ipcMain.handle('of:recovered', async (e) => {
  const id = winOf(e).__app;
  const open = new Set(BrowserWindow.getAllWindows().map((w) => w.__recoveryId).filter(Boolean));
  try {
    const files = await fs.readdir(recoveryDir());
    const out = [];
    for (const f of files) {
      if (!f.endsWith('.json') || open.has(f.slice(0, -5))) continue;
      try {
        const j = JSON.parse(await fs.readFile(path.join(recoveryDir(), f), 'utf8'));
        if (j.app === id) out.push({ id: f.slice(0, -5), name: j.name, savedAt: j.savedAt, path: j.path });
      } catch { /* dañado */ }
    }
    return out.sort((a, b) => b.savedAt - a.savedAt);
  } catch { return []; }
});
ipcMain.handle('of:recover', async (e, rid) => {
  const j = JSON.parse(await fs.readFile(path.join(recoveryDir(), `${path.basename(rid)}.json`), 'utf8'));
  winOf(e).__recoveryId = path.basename(rid);
  return j;
});
ipcMain.handle('of:discard-recovered', async (_e, rid) => {
  await fs.rm(path.join(recoveryDir(), `${path.basename(rid)}.json`), { force: true });
  return true;
});

// ---------------------------------------------------------------- Impresiona: público y exportar imágenes
const audiences = new Map(); // webContents del moderador -> ventana del público
ipcMain.handle('of:audience-open', (e, { pres, i = 0, step = 0 }) => {
  const presenter = BrowserWindow.fromWebContents(e.sender);
  const displays = screen.getAllDisplays();
  const here = screen.getDisplayMatching(presenter.getBounds());
  const target = displays.find((d) => d.id !== here.id) || here;
  const aud = new BrowserWindow({
    x: target.bounds.x, y: target.bounds.y, width: Math.min(1280, target.bounds.width), height: Math.min(800, target.bounds.height),
    fullscreen: !TEST_DIR && target.id !== here.id, backgroundColor: '#000000', title: 'Presentación — Impresiona', autoHideMenuBar: true,
    icon: path.join(ROOT, 'build', 'apps', 'impresiona.png'),
    webPreferences: { preload: path.join(ROOT, 'preload', 'office.js'), contextIsolation: true, nodeIntegration: false, sandbox: true },
  });
  aud.setMenu(null);
  aud.loadURL('app://suite/apps/impresiona/index.html?audience=1');
  aud.webContents.on('did-finish-load', () => aud.webContents.send('audience-msg', { type: 'load', pres, i, step }));
  aud.on('closed', () => { audiences.delete(e.sender.id); if (!e.sender.isDestroyed()) e.sender.send('audience-closed'); });
  audiences.set(e.sender.id, aud);
  return { sameScreen: target.id === here.id };
});
ipcMain.handle('of:audience-send', (e, msg) => { const a = audiences.get(e.sender.id); if (a && !a.isDestroyed()) a.webContents.send('audience-msg', msg); return true; });
ipcMain.handle('of:audience-close', (e) => { const a = audiences.get(e.sender.id); if (a && !a.isDestroyed()) a.close(); audiences.delete(e.sender.id); return true; });

const renderWaiters = new Map();
ipcMain.handle('of:render-ready', (e, i) => { const w = renderWaiters.get(e.sender.id); if (w) w(i); return true; });
ipcMain.handle('of:render-slides', async (e, { pres, base = 'Diapositiva' }) => {
  let folder = TEST_DIR;
  if (!folder) {
    const r = await dialog.showOpenDialog(BrowserWindow.fromWebContents(e.sender), { title: 'Elige la carpeta para las imágenes', properties: ['openDirectory', 'createDirectory'] });
    if (r.canceled || !r.filePaths[0]) return null;
    folder = r.filePaths[0];
  }
  const { w, h } = pres.size;
  const win = new BrowserWindow({ width: w, height: h, show: false, useContentSize: true, frame: false, webPreferences: { preload: path.join(ROOT, 'preload', 'office.js'), contextIsolation: true, sandbox: true, offscreen: true } });
  await win.loadURL('app://suite/apps/impresiona/index.html?render=1');
  win.webContents.send('audience-msg', { type: 'load', pres });
  const files = [];
  try {
    for (let i = 0; i < pres.slides.length; i++) {
      const ready = new Promise((res) => renderWaiters.set(win.webContents.id, res));
      win.webContents.send('audience-msg', { type: 'render', i });
      await Promise.race([ready, new Promise((res) => setTimeout(res, 4000))]);
      const img = await win.webContents.capturePage({ x: 0, y: 0, width: w, height: h });
      const out = path.join(folder, `${base}-${i + 1}.png`);
      await fs.writeFile(out, img.toPNG());
      files.push(out);
    }
  } finally { renderWaiters.delete(win.webContents.id); win.destroy(); }
  return { folder, files };
});

// ---------------------------------------------------------------- arranque
app.whenReady().then(() => {
  loadPrefs();
  require('electron').nativeTheme.themeSource = prefs.theme === 'auto' ? 'system' : prefs.theme;
  registerAppProtocol();
  if (!isMac) Menu.setApplicationMenu(null);
  else Menu.setApplicationMenu(Menu.buildFromTemplate([{ role: 'appMenu', label: 'Mi Oficina MX' }, { role: 'editMenu', label: 'Editar' }, { role: 'windowMenu', label: 'Ventana' }]));
  if (!TEST_DIR) createTray();
  const files = pendingFiles;
  pendingFiles = [];
  createLauncher({ minimized: files.length > 0 });
  ready = true;
  files.forEach((f) => openApp(appForFile(f), f));
  if (process.env.OFICINA_OPEN) openApp(process.env.OFICINA_OPEN);
  app.on('activate', () => showLauncher());
});

app.on('before-quit', () => { quitting = true; });
app.on('window-all-closed', () => { if (!isMac || quitting) app.quit(); });

module.exports = { APPS, appForFile };
