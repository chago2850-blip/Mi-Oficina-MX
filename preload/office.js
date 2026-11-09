// Puente seguro entre la interfaz de Escribiente / Calcula Datos / Impresiona y el sistema.
const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('oficina', {
  openDialog: (opts) => ipcRenderer.invoke('of:open-dialog', opts || {}),
  openImage: () => ipcRenderer.invoke('of:open-image'),
  readPath: (p) => ipcRenderer.invoke('of:read-path', p),
  pathForFile: (file) => { try { return webUtils.getPathForFile(file); } catch { return null; } },
  save: (opts) => ipcRenderer.invoke('of:save', opts),
  exportPdf: (opts) => ipcRenderer.invoke('of:export-pdf', opts || {}),
  print: () => ipcRenderer.invoke('of:print'),
  recents: () => ipcRenderer.invoke('of:recents'),
  clearRecents: () => ipcRenderer.invoke('of:clear-recents'),
  confirm: (opts) => ipcRenderer.invoke('of:confirm', opts),
  setState: (s) => ipcRenderer.invoke('of:state', s),
  newWindow: (opts) => ipcRenderer.invoke('of:new-window', opts || {}),
  openInSuite: (p) => ipcRenderer.invoke('of:open-in-suite', p),
  showInFolder: (p) => ipcRenderer.invoke('of:show-in-folder', p),
  showLauncher: () => ipcRenderer.invoke('of:show-launcher'),
  appInfo: () => ipcRenderer.invoke('of:app-info'),
  autosave: (s) => ipcRenderer.invoke('of:autosave', s),
  recovered: () => ipcRenderer.invoke('of:recovered'),
  recover: (id) => ipcRenderer.invoke('of:recover', id),
  discardRecovered: (id) => ipcRenderer.invoke('of:discard-recovered', id),
  onMenu: (fn) => ipcRenderer.on('menu', (_e, cmd) => fn(cmd)),
  onOpenPath: (fn) => ipcRenderer.on('open-path', (_e, p) => fn(p)),
});
