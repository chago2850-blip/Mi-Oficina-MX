// Puente del Launcher (dona).
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('launcher', {
  open: (id) => ipcRenderer.invoke('launcher:open', id),
  quit: () => ipcRenderer.invoke('launcher:quit'),
  hide: () => ipcRenderer.invoke('launcher:hide'),
  menu: () => ipcRenderer.invoke('launcher:menu'),
  info: () => ipcRenderer.invoke('launcher:info'),
  onRunning: (fn) => ipcRenderer.on('launcher:running', (_e, list) => fn(list)),
  onRestored: (fn) => ipcRenderer.on('launcher:restored', () => fn()),
  onMinimizing: (fn) => ipcRenderer.on('launcher:minimizing', () => fn()),
  onAbout: (fn) => ipcRenderer.on('launcher:about', () => fn()),
  prefs: () => ipcRenderer.invoke('of:prefs'),
  setPrefs: (p) => ipcRenderer.invoke('of:set-prefs', p),
  onPrefs: (fn) => ipcRenderer.on('launcher:prefs', () => fn()),
});
