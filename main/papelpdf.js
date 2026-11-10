// PapelPDF dentro de Mi Oficina MX — proceso principal.
// Código original de PapelPDF (MIT, Santiago Miguel) adaptado para varias ventanas.
const { app, BrowserWindow, ipcMain, dialog, Menu, shell } = require('electron');
const path = require('path');
const fs = require('fs/promises');
const os = require('os');

const isMac = process.platform === 'darwin';
const ROOT = path.join(__dirname, '..');
const windows = new Set();
const handle = (ch, fn) => ipcMain.handle(ch, fn);
const winOf = (e) => BrowserWindow.fromWebContents(e.sender);

// --- recientes ---
const recentsFile = () => path.join(app.getPath('userData'), 'recientes.json');
async function getRecents() {
  try {
    const list = JSON.parse(await fs.readFile(recentsFile(), 'utf8'));
    const out = [];
    for (const r of list) {
      try { await fs.access(r.path); out.push(r); } catch { /* ya no existe */ }
    }
    return out;
  } catch { return []; }
}
async function addRecent(p) {
  if (!p) return;
  const list = (await getRecents()).filter((r) => r.path !== p);
  list.unshift({ path: p, name: path.basename(p), openedAt: Date.now() });
  await fs.mkdir(path.dirname(recentsFile()), { recursive: true });
  await fs.writeFile(recentsFile(), JSON.stringify(list.slice(0, 12), null, 2));
  if (isMac || process.platform === 'win32') app.addRecentDocument(p);
}

async function readPdf(p) {
  const buf = await fs.readFile(p);
  await addRecent(p);
  return { path: p, name: path.basename(p), bytes: new Uint8Array(buf) };
}

// En pruebas automáticas se evitan los diálogos nativos.
const TEST_DIR = process.env.PAPELPDF_TEST_DIR || null;

// --- qpdf (WebAssembly): cifrar, descifrar, comprimir ---
const QPDF_WASM = path.join(ROOT, 'node_modules', '@neslinesli93', 'qpdf-wasm', 'dist', 'qpdf.wasm');
async function runQpdf(args, input) {
  const createQpdf = require('@neslinesli93/qpdf-wasm');
  const errors = [];
  const q = await createQpdf({ locateFile: () => QPDF_WASM, print: () => {}, printErr: (t) => errors.push(t) });
  q.FS.writeFile('/in.pdf', input);
  const rc = q.callMain(args);
  let out = null;
  try { out = q.FS.readFile('/out.pdf'); } catch { /* sin salida */ }
  return { rc, out, errors };
}

handle('qpdf', async (_e, opts) => {
  const bytes = Buffer.from(opts.bytes);
  if (opts.op === 'decrypt') {
    const pw = opts.password ? [`--password=${opts.password}`] : [];
    // --is-encrypted también necesita la contraseña para poder abrir el archivo
    const check = await runQpdf([...pw, '--is-encrypted', '/in.pdf'], bytes);
    if (check.rc === 2) return { encrypted: false }; // 2 = no está cifrado
    const r = await runQpdf([...pw, '--decrypt', '/in.pdf', '/out.pdf'], bytes);
    if (!r.out || (r.rc !== 0 && r.rc !== 3)) throw new Error(r.errors.join(' ') || 'No se pudo descifrar el PDF.');
    return { encrypted: true, bytes: r.out };
  }
  if (opts.op === 'encrypt') {
    const yn = (v) => (v ? 'y' : 'n');
    const args = ['/in.pdf', '--encrypt', opts.user || '', opts.owner || opts.user || '', '256',
      `--print=${opts.print ? 'full' : 'none'}`, `--extract=${yn(opts.copy)}`, `--modify=${opts.modify ? 'all' : 'none'}`,
      '--', '/out.pdf'];
    const r = await runQpdf(args, bytes);
    if (!r.out || (r.rc !== 0 && r.rc !== 3)) throw new Error(r.errors.join(' ') || 'No se pudo cifrar el PDF.');
    return { bytes: r.out };
  }
  if (opts.op === 'compress') {
    const r = await runQpdf(['--object-streams=generate', '--compress-streams=y', '--recompress-flate', '--compression-level=9', '--remove-unreferenced-resources=yes', '/in.pdf', '/out.pdf'], bytes);
    if (!r.out || (r.rc !== 0 && r.rc !== 3)) throw new Error(r.errors.join(' ') || 'No se pudo comprimir el PDF.');
    return { bytes: r.out };
  }
  throw new Error('Operación desconocida');
});

// --- Word (.docx) ---
handle('make-docx', async (_e, { title, pages }) => {
  const { Document, Packer, Paragraph, TextRun, HeadingLevel } = require('docx');
  const children = [];
  pages.forEach((lines, pi) => {
    if (!lines.length) {
      children.push(new Paragraph({ pageBreakBefore: pi > 0, children: [new TextRun('')] }));
      return;
    }
    lines.forEach((ln, li) => {
      const big = ln.size >= 16;
      children.push(new Paragraph({
        pageBreakBefore: pi > 0 && li === 0,
        heading: big ? (ln.size >= 22 ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2) : undefined,
        spacing: { after: ln.gapAfter ? 160 : 40 },
        children: [new TextRun({ text: ln.text, size: Math.round(Math.min(Math.max(ln.size, 8), 40) * 2), bold: ln.bold || undefined })],
      }));
    });
  });
  const docx = new Document({ creator: 'PapelPDF', title, sections: [{ children }] });
  return new Uint8Array(await Packer.toBuffer(docx));
});

// --- diálogos y archivos ---
handle('open-dialog', async (_e, { multi } = {}) => {
  const r = await dialog.showOpenDialog(winOf(_e), {
    title: multi ? 'Selecciona los PDFs' : 'Abrir PDF',
    properties: multi ? ['openFile', 'multiSelections'] : ['openFile'],
    filters: [{ name: 'Documentos PDF', extensions: ['pdf'] }],
  });
  if (r.canceled) return [];
  return Promise.all(r.filePaths.map(readPdf));
});

handle('open-images', async () => {
  const r = await dialog.showOpenDialog(winOf(_e), {
    title: 'Selecciona las imágenes',
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'Imágenes', extensions: ['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp'] }],
  });
  if (r.canceled) return [];
  return Promise.all(r.filePaths.map(async (p) => ({ name: path.basename(p), bytes: new Uint8Array(await fs.readFile(p)) })));
});

handle('choose-folder', async (_e, { title }) => {
  if (TEST_DIR) return TEST_DIR;
  const r = await dialog.showOpenDialog(winOf(_e), { title, properties: ['openDirectory', 'createDirectory'] });
  return r.canceled ? null : r.filePaths[0];
});

handle('write-into', async (_e, { folder, name, bytes }) => {
  const safe = path.basename(name);
  const out = path.join(folder, safe);
  await fs.writeFile(out, Buffer.from(bytes));
  return out;
});

handle('read-path', async (_e, p) => readPdf(p));
handle('recents', async () => getRecents());
handle('clear-recents', async () => { await fs.rm(recentsFile(), { force: true }); return []; });

const KINDS = {
  pdf: { name: 'Documento PDF', ext: 'pdf', title: 'Guardar PDF' },
  png: { name: 'Imagen PNG', ext: 'png', title: 'Exportar imagen' },
  csv: { name: 'CSV', ext: 'csv', title: 'Exportar datos' },
  txt: { name: 'Texto', ext: 'txt', title: 'Exportar texto' },
  docx: { name: 'Documento de Word', ext: 'docx', title: 'Exportar a Word' },
};
handle('save-file', async (_e, { bytes, path: target, suggestedName, ask, kind = 'pdf' }) => {
  const k = KINDS[kind] || KINDS.pdf;
  let out = target;
  if (TEST_DIR && (ask || !out)) {
    out = path.join(TEST_DIR, suggestedName || `salida.${k.ext}`);
  } else if (ask || !out) {
    const r = await dialog.showSaveDialog(winOf(_e), {
      title: k.title,
      defaultPath: suggestedName || `documento.${k.ext}`,
      filters: [{ name: k.name, extensions: [k.ext] }],
    });
    if (r.canceled || !r.filePath) return null;
    out = r.filePath;
  }
  await fs.writeFile(out, Buffer.from(bytes));
  if (k.ext === 'pdf') await addRecent(out);
  return { path: out, name: path.basename(out) };
});

handle('print-pdf', async (_e, bytes) => {
  // Abre el PDF final en el visor de Chromium, que trae su propio botón de imprimir.
  const tmp = path.join(os.tmpdir(), `papelpdf-imprimir-${Date.now()}.pdf`);
  await fs.writeFile(tmp, Buffer.from(bytes));
  const pw = new BrowserWindow({
    width: 1000, height: 800, title: 'Imprimir — PapelPDF', parent: winOf(_e),
    webPreferences: { plugins: true, contextIsolation: true, sandbox: true },
  });
  pw.loadURL('file://' + tmp);
  return true;
});

handle('show-in-folder', (_e, p) => { shell.showItemInFolder(p); return true; });

handle('confirm', async (_e, { message, detail, buttons }) => {
  if (TEST_DIR) return 0;
  const { response } = await dialog.showMessageBox(winOf(_e), {
    type: 'question', message, detail, buttons, defaultId: 0, cancelId: buttons.length - 1,
  });
  return response;
});

handle('app-info', () => ({ version: app.getVersion(), platform: process.platform, arch: process.arch }));


function buildMenu(win) {
  const send = (cmd) => { if (!win.isDestroyed()) win.webContents.send('menu', cmd); };
  return Menu.buildFromTemplate([
    ...(isMac ? [{ role: 'appMenu', label: 'Mi Oficina MX' }] : []),
    { label: 'Archivo', submenu: [
      { label: 'Abrir…', accelerator: 'CmdOrCtrl+O', click: () => send('open') },
      { label: 'Combinar PDFs…', click: () => send('merge') },
      { label: 'Crear PDF desde imágenes…', click: () => send('imagesToPdf') },
      { type: 'separator' },
      { label: 'Guardar', accelerator: 'CmdOrCtrl+S', click: () => send('save') },
      { label: 'Guardar como…', accelerator: 'CmdOrCtrl+Shift+S', click: () => send('saveAs') },
      { label: 'Exportar', submenu: [
        { label: 'Word (.docx)…', click: () => send('toDocx') },
        { label: 'Texto (.txt)…', click: () => send('toTxt') },
        { label: 'Imágenes (PNG/JPG)…', click: () => send('toImages') },
      ] },
      { label: 'Imprimir…', accelerator: 'CmdOrCtrl+P', click: () => send('print') },
      { type: 'separator' },
      { label: 'Cerrar documento', accelerator: 'CmdOrCtrl+W', click: () => send('close') },
      { label: 'Cerrar ventana', accelerator: 'CmdOrCtrl+Shift+W', click: () => win.close() },
    ] },
    { label: 'Editar', submenu: [
      { label: 'Deshacer', accelerator: 'CmdOrCtrl+Z', click: () => send('undo') },
      { label: 'Rehacer', accelerator: isMac ? 'Cmd+Shift+Z' : 'Ctrl+Y', click: () => send('redo') },
      { type: 'separator' },
      { role: 'cut', label: 'Cortar' }, { role: 'copy', label: 'Copiar' }, { role: 'paste', label: 'Pegar' },
      { role: 'selectAll', label: 'Seleccionar todo' },
      { type: 'separator' },
      { label: 'Buscar…', accelerator: 'CmdOrCtrl+F', click: () => send('find') },
    ] },
    { label: 'Herramientas', submenu: [
      { label: 'Reconocer texto (OCR)…', click: () => send('ocr') },
      { label: 'Proteger con contraseña…', click: () => send('protect') },
      { label: 'Comprimir…', click: () => send('compress') },
      { label: 'Marca de agua…', click: () => send('watermark') },
      { label: 'Numerar páginas…', click: () => send('pageNumbers') },
    ] },
    { label: 'Ver', submenu: [
      { label: 'Acercar', accelerator: 'CmdOrCtrl+=', click: () => send('zoomIn') },
      { label: 'Alejar', accelerator: 'CmdOrCtrl+-', click: () => send('zoomOut') },
      { label: 'Ajustar al ancho', accelerator: 'CmdOrCtrl+0', click: () => send('zoomFit') },
      { type: 'separator' },
      { role: 'togglefullscreen', label: 'Pantalla completa' },
      { role: 'toggleDevTools', label: 'Herramientas de desarrollo' },
    ] },
    { label: 'Ayuda', submenu: [
      { label: 'Atajos de teclado', click: () => send('shortcuts') },
      { label: 'Acerca de PapelPDF', click: () => send('about') },
    ] },
  ]);
}

function openWindow(filePath, { onClosed } = {}) {
  const win = new BrowserWindow({
    width: 1360, height: 880, minWidth: 900, minHeight: 600,
    title: 'PapelPDF', backgroundColor: '#f4efe6', show: false,
    icon: path.join(ROOT, 'build', 'apps', 'papelpdf.png'),
    webPreferences: {
      preload: path.join(ROOT, 'preload', 'papelpdf.js'),
      contextIsolation: true, nodeIntegration: false, sandbox: true, spellcheck: true,
    },
  });
  win.__app = 'papelpdf';
  windows.add(win);
  win.__menu = buildMenu(win);
  if (isMac) win.on('focus', () => Menu.setApplicationMenu(win.__menu)); else win.setMenu(win.__menu);
  if (isMac) Menu.setApplicationMenu(win.__menu);
  win.loadURL(require('./i18n').withLang('app://suite/apps/papelpdf/index.html'));
  win.once('ready-to-show', () => win.show());
  win.webContents.on('did-finish-load', () => { if (filePath) win.webContents.send('open-path', filePath); });
  win.webContents.setWindowOpenHandler(({ url }) => { if (/^https?:/.test(url)) shell.openExternal(url); return { action: 'deny' }; });
  win.webContents.on('will-navigate', (e, url) => { if (!url.startsWith('app://')) e.preventDefault(); });
  win.on('close', async (e) => {
    if (win.__forceClose) return;
    e.preventDefault();
    const dirty = await win.webContents.executeJavaScript('window.__papelDirty ? window.__papelDirty() : false').catch(() => false);
    if (!dirty || TEST_DIR) { win.__forceClose = true; win.close(); return; }
    const { response } = await dialog.showMessageBox(win, {
      type: 'warning', buttons: ['Descartar cambios', 'Cancelar'], defaultId: 1, cancelId: 1,
      message: 'Hay cambios sin guardar.', detail: '¿Quieres cerrar PapelPDF sin guardar?',
    });
    if (response === 0) { win.__forceClose = true; win.close(); }
  });
  win.on('closed', () => { windows.delete(win); if (onClosed) onClosed(win); });
  return win;
}

function openPath(filePath, opts) {
  // Si hay una ventana de PapelPDF abierta, se abre ahí (PapelPDF maneja varios documentos por ventana).
  const existing = [...windows][0];
  if (existing && !existing.isDestroyed()) {
    if (existing.isMinimized()) existing.restore();
    existing.focus();
    if (filePath) existing.webContents.send('open-path', filePath);
    return existing;
  }
  return openWindow(filePath, opts);
}

module.exports = {
  rebuildMenu: (win) => { win.__menu = buildMenu(win); if (isMac) { if (win.isFocused()) Menu.setApplicationMenu(win.__menu); } else win.setMenu(win.__menu); }, openWindow, openPath, windows };
