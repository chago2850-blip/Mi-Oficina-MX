// Mi Oficina MX — armazón compartido de Escribiente, Calcula Datos e Impresiona:
// barra superior, cinta, inicio (backstage), diálogos, avisos y ciclo de vida del archivo.
// GPL-3.0-or-later
import { icon } from './icons.js';

export { icon };
export const $ = (s, r = document) => r.querySelector(s);
export const $$ = (s, r = document) => [...r.querySelectorAll(s)];
export function h(tag, attrs = {}, ...kids) {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === 'class') n.className = v;
    else if (k === 'html') n.innerHTML = v;
    else if (k === 'style' && typeof v === 'object') Object.assign(n.style, v);
    else if (k.startsWith('on') && typeof v === 'function') n.addEventListener(k.slice(2), v);
    else n.setAttribute(k, v === true ? '' : v);
  }
  for (const k of kids.flat()) if (k != null && k !== false) n.append(k.nodeType ? k : document.createTextNode(String(k)));
  return n;
}
export const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
export const api = window.oficina || null;

// ------------------------------------------------------------------ avisos
let toastBox;
export function toast(msg, { error = false, action, ms = 3500 } = {}) {
  if (!toastBox) { toastBox = h('div', { class: 'toasts', role: 'status', 'aria-live': 'polite' }); document.body.append(toastBox); }
  const t = h('div', { class: `toast${error ? ' err' : ''}` }, h('span', {}, msg));
  if (action) t.append(h('button', { onclick: () => { action.run(); t.remove(); } }, action.label));
  toastBox.append(t);
  setTimeout(() => t.remove(), action ? ms * 2 : ms);
}

// ------------------------------------------------------------------ diálogos
export function modal({ title, body, buttons = [{ label: 'Aceptar', value: true, primary: true }], width, onOpen }) {
  return new Promise((resolve) => {
    const prev = document.activeElement;
    const content = typeof body === 'string' ? h('div', { html: body }) : body;
    const footer = h('footer');
    const box = h('div', { class: 'modal', role: 'dialog', 'aria-modal': 'true', style: width ? { width } : null },
      h('header', {}, title), h('div', { class: 'body' }, content), footer);
    const back = h('div', { class: 'modal-back' }, box);
    const close = (v) => { back.remove(); document.removeEventListener('keydown', onKey, true); prev?.focus?.(); resolve(v); };
    buttons.forEach((b) => footer.append(h('button', { class: `btn${b.primary ? ' primary' : ''}`, onclick: () => close(typeof b.value === 'function' ? b.value(content) : b.value) }, b.label)));
    const onKey = (e) => {
      if (e.key === 'Escape') { e.stopPropagation(); close(null); }
      if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA' && !e.target.matches('button')) {
        const p = buttons.find((b) => b.primary);
        if (p) { e.preventDefault(); close(typeof p.value === 'function' ? p.value(content) : p.value); }
      }
    };
    document.addEventListener('keydown', onKey, true);
    back.addEventListener('mousedown', (e) => { if (e.target === back) close(null); });
    document.body.append(back);
    onOpen?.(content, close);
    setTimeout(() => (content.querySelector('input,select,textarea') || footer.querySelector('.primary'))?.focus(), 30);
  });
}

// Formulario rápido: fields = [{ name, label, type, value, options, min, max, step }]
export async function form(title, fields, okLabel = 'Aceptar') {
  const body = h('div', {});
  for (const f of fields) {
    if (f.type === 'checkbox') {
      body.append(h('label', { class: 'check' }, h('input', { type: 'checkbox', name: f.name, checked: f.value ? true : null }), f.label));
    } else if (f.type === 'select') {
      const s = h('select', { name: f.name });
      f.options.forEach((o) => { const [v, l] = Array.isArray(o) ? o : [o, o]; s.append(h('option', { value: v, selected: String(v) === String(f.value) ? true : null }, l)); });
      body.append(h('label', { class: 'field' }, h('span', {}, f.label), s));
    } else if (f.type === 'textarea') {
      body.append(h('label', { class: 'field' }, h('span', {}, f.label), h('textarea', { name: f.name, rows: f.rows || 4 }, f.value || '')));
    } else if (f.type === 'note') {
      body.append(h('p', { style: { color: 'var(--ink-3)', margin: '6px 0' } }, f.label));
    } else {
      body.append(h('label', { class: 'field' }, h('span', {}, f.label), h('input', { name: f.name, type: f.type || 'text', value: f.value ?? '', min: f.min, max: f.max, step: f.step, placeholder: f.placeholder })));
    }
  }
  return modal({
    title, body,
    buttons: [
      { label: 'Cancelar', value: null },
      { label: okLabel, primary: true, value: (c) => {
        const out = {};
        c.querySelectorAll('[name]').forEach((el) => { out[el.name] = el.type === 'checkbox' ? el.checked : (el.type === 'number' ? Number(el.value) : el.value); });
        return out;
      } },
    ],
  });
}
export const alertBox = (title, text) => modal({ title, body: h('p', {}, text) });
export const confirmBox = (title, text, ok = 'Aceptar') => modal({ title, body: h('p', {}, text), buttons: [{ label: 'Cancelar', value: false }, { label: ok, value: true, primary: true }] });

// ------------------------------------------------------------------ menús desplegables
let openDrop = null;
export function closeDropdown() { openDrop?.remove(); openDrop = null; }
export function dropdown(anchor, content) {
  closeDropdown();
  const d = h('div', { class: 'dropdown' });
  if (Array.isArray(content)) {
    content.forEach((it) => {
      if (it === '-') d.append(h('hr'));
      else if (it.header) d.append(h('div', { class: 'hdr' }, it.header));
      else d.append(h('button', { class: 'mi', onclick: () => { closeDropdown(); it.run(); }, html: `${it.icon ? icon(it.icon, 18) : ''}<span>${esc(it.label)}</span>${it.kbd ? `<kbd>${esc(it.kbd)}</kbd>` : ''}`, style: it.style || null }));
    });
  } else d.append(content);
  document.body.append(d);
  const r = anchor.getBoundingClientRect();
  const dw = d.offsetWidth, dh = d.offsetHeight;
  d.style.left = `${Math.max(4, Math.min(r.left, innerWidth - dw - 6))}px`;
  d.style.top = `${r.bottom + dh + 4 > innerHeight ? Math.max(4, r.top - dh - 4) : r.bottom + 4}px`;
  openDrop = d;
  setTimeout(() => d.querySelector('button')?.focus({ preventScroll: true }), 0);
  return d;
}
document.addEventListener('mousedown', (e) => { if (openDrop && !openDrop.contains(e.target)) closeDropdown(); }, true);
document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && openDrop) { closeDropdown(); e.stopPropagation(); } }, true);
addEventListener('blur', closeDropdown);

export const PALETTE = [
  '#000000', '#434343', '#666666', '#999999', '#B7B7B7', '#D9D9D9', '#EFEFEF', '#F3F3F3', '#FFFFFF', 'transparent',
  '#C62828', '#E64A19', '#F9A825', '#2E7D32', '#00897B', '#0277BD', '#1565C0', '#4527A0', '#6A1B9A', '#AD1457',
  '#EF9A9A', '#FFAB91', '#FFF59D', '#A5D6A7', '#80CBC4', '#81D4FA', '#90CAF9', '#B39DDB', '#CE93D8', '#F48FB1',
  '#FFEBEE', '#FBE9E7', '#FFFDE7', '#E8F5E9', '#E0F2F1', '#E1F5FE', '#E3F2FD', '#EDE7F6', '#F3E5F5', '#FCE4EC',
];
export function colorPicker(anchor, onPick, { allowNone = true, noneLabel = 'Sin color' } = {}) {
  const box = h('div', {});
  const sw = h('div', { class: 'swatches' });
  PALETTE.forEach((c) => {
    if (c === 'transparent' && !allowNone) return;
    sw.append(h('button', { title: c === 'transparent' ? noneLabel : c, style: { background: c === 'transparent' ? 'linear-gradient(135deg,#fff 45%,#d33 45%,#d33 55%,#fff 55%)' : c }, onclick: () => { closeDropdown(); onPick(c); } }));
  });
  box.append(sw);
  const custom = h('input', { type: 'color', value: '#5B48E8', style: { width: '100%', height: '28px', border: 0, background: 'none', cursor: 'pointer' } });
  custom.addEventListener('change', () => { closeDropdown(); onPick(custom.value); });
  box.append(h('div', { class: 'hdr' }, 'Personalizado'), custom);
  dropdown(anchor, box);
}

// ------------------------------------------------------------------ cinta
// tabs: [{ id, label, contextual, groups: [{ label, items: [...] }] }]
// item: { cmd, label, icon, size: 'lg'|'sm', toggle, title, kbd, menu: [...] , color, select: {options, width}, html, render }
export class Ribbon {
  constructor(el, { tabs, onCommand, onFile }) {
    this.el = el; this.tabs = tabs; this.onCommand = onCommand; this.onFile = onFile;
    this.active = tabs.find((t) => !t.contextual)?.id;
    this.visibleContextual = new Set();
    this.controls = new Map();
    this.render();
  }
  render() {
    this.el.innerHTML = '';
    this.el.classList.add('ribbon');
    const tabsRow = h('div', { class: 'tabs', role: 'tablist' });
    tabsRow.append(h('button', { class: 'tab file', onclick: () => this.onFile() }, 'Archivo'));
    this.tabEls = {};
    for (const t of this.tabs) {
      const b = h('button', { class: `tab${t.contextual ? ' contextual' : ''}`, role: 'tab', 'aria-selected': String(t.id === this.active), 'data-tab': t.id, onclick: () => this.show(t.id), hidden: t.contextual && !this.visibleContextual.has(t.id) ? true : null }, t.label);
      this.tabEls[t.id] = b;
      tabsRow.append(b);
    }
    tabsRow.append(h('div', { class: 'spacer' }), h('button', { class: 'collapse', title: 'Contraer la cinta (Ctrl+F1)', html: icon('chevronDown', 16), onclick: () => this.toggleCollapse() }));
    this.panel = h('div', { class: 'panel', role: 'tabpanel' });
    this.el.append(tabsRow, this.panel);
    this.renderPanel();
  }
  toggleCollapse() { this.el.classList.toggle('collapsed'); dispatchEvent(new Event('resize')); }
  show(id) {
    this.active = id;
    if (this.el.classList.contains('collapsed')) this.toggleCollapse();
    Object.entries(this.tabEls).forEach(([k, b]) => b.setAttribute('aria-selected', String(k === id)));
    this.renderPanel();
  }
  setContextual(id, visible) {
    if (visible === this.visibleContextual.has(id)) return;
    if (visible) this.visibleContextual.add(id); else this.visibleContextual.delete(id);
    if (this.tabEls[id]) this.tabEls[id].hidden = !visible;
    if (!visible && this.active === id) this.show(this.tabs[0].id);
  }
  renderPanel() {
    this.panel.innerHTML = '';
    this.controls.clear();
    const tab = this.tabs.find((t) => t.id === this.active);
    for (const g of tab.groups) {
      const items = h('div', { class: 'items' });
      for (const it of g.items) items.append(this.renderItem(it));
      this.panel.append(h('div', { class: 'group' }, items, h('div', { class: 'label' }, g.label)));
    }
    this.onRender?.();
  }
  renderItem(it) {
    if (it.stack) { const s = h('div', { class: 'stack' }); it.stack.forEach((x) => s.append(this.renderItem(x))); return s; }
    if (it.row) { const r = h('div', { class: 'row' }); it.row.forEach((x) => r.append(this.renderItem(x))); return r; }
    if (it.render) { const n = it.render(); if (it.id) this.controls.set(it.id, n); return n; }
    if (it.select) {
      const s = h('select', { class: 'rsel', title: it.title || it.label, style: { width: it.select.width || 'auto' } });
      it.select.options.forEach((o) => { const [v, l] = Array.isArray(o) ? o : [o, o]; s.append(h('option', { value: v }, l)); });
      s.addEventListener('mousedown', () => this.onBeforeSelect?.());
      s.addEventListener('change', () => this.onCommand(it.cmd, s.value));
      this.controls.set(it.cmd, s);
      return s;
    }
    const size = it.size || 'sm';
    const b = h('button', { class: `rb ${size}${it.color ? ' color' : ''}`, title: `${it.title || it.label || ''}${it.kbd ? ` (${it.kbd})` : ''}`, 'aria-label': it.label || it.title, 'data-cmd': it.cmd });
    b.innerHTML = `${icon(it.icon, size === 'lg' ? 26 : 18)}${it.label && (size === 'lg' || it.showLabel) ? `<span class="txt">${esc(it.label)}</span>` : ''}${it.menu || it.color ? `<svg class="caret" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M6 9l6 6 6-6"/></svg>` : ''}${it.color ? '<span class="bar"></span>' : ''}`;
    if (it.color) b.style.setProperty('--c', it.colorValue || '#E53935');
    b.addEventListener('mousedown', (e) => e.preventDefault()); // no robar la selección del editor
    b.addEventListener('click', () => {
      if (it.menu) dropdown(b, typeof it.menu === 'function' ? it.menu() : it.menu);
      else if (it.color) colorPicker(b, (c) => { it.colorValue = c; b.style.setProperty('--c', c); this.onCommand(it.cmd, c); }, it.colorOpts);
      else this.onCommand(it.cmd, b);
    });
    this.controls.set(it.cmd, b);
    return b;
  }
  setOn(cmd, on) { const c = this.controls.get(cmd); if (c && c.classList) c.classList.toggle('on', !!on); }
  setValue(cmd, v) { const c = this.controls.get(cmd); if (c && c.tagName === 'SELECT' && document.activeElement !== c) { if (![...c.options].some((o) => o.value === String(v)) && v !== '') c.append(h('option', { value: v }, v)); c.value = v; } }
  setEnabled(cmd, en) { const c = this.controls.get(cmd); if (c) c.disabled = !en; }
}

// ------------------------------------------------------------------ armazón y ciclo de vida
// adapter = { exts, saveExt, saveFilter, newDoc(tpl), load(file), serialize(ext) -> Uint8Array|string,
//             snapshot() -> string, restore(str), isPristine(), templates: [{id,name,thumb}], pdfOptions(),
//             beforePrint(), afterPrint(), onCommand(cmd, arg), onReady() }
export class OfficeShell {
  constructor({ appId, appName, logo, tabs, adapter }) {
    this.appId = appId; this.appName = appName; this.logo = logo; this.adapter = adapter;
    this.doc = { name: 'Sin título', path: null, ext: adapter.saveExt, dirty: false };
    this.recoveryId = `${appId}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
    document.body.dataset.app = appId;
    this.top = h('header', { class: 'topbar' });
    this.ribbonEl = h('nav', {});
    this.workspace = h('main', { class: 'workspace', id: 'workspace' });
    this.status = h('footer', { class: 'statusbar' });
    document.body.prepend(this.top, this.ribbonEl, this.workspace, this.status);
    this.buildTop();
    this.ribbon = new Ribbon(this.ribbonEl, { tabs, onCommand: (c, a) => this.command(c, a), onFile: () => this.showBackstage('info') });
    this.bindGlobal();
    window.__oficinaSave = () => this.save();
    window.__shell = this;
    this.prefs = { theme: 'auto', author: '', autosaveSec: 20, spellcheck: true };
    this.props = {};
    this.loadPrefs();
  }

  // -------------------------------------------------------------- preferencias y tema
  async loadPrefs() {
    if (api?.prefs) this.prefs = { ...this.prefs, ...(await api.prefs()) };
    this.applyPrefs();
    api?.onPrefs?.((p) => { this.prefs = { ...this.prefs, ...p }; this.applyPrefs(); });
    matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => this.applyPrefs());
  }
  applyPrefs() {
    const dark = this.prefs.theme === 'dark' || (this.prefs.theme === 'auto' && matchMedia('(prefers-color-scheme: dark)').matches);
    document.body.dataset.theme = dark ? 'dark' : 'light';
    this.adapter.onPrefs?.(this.prefs);
  }
  async optionsDialog() {
    const v = await form('Opciones de Mi Oficina MX', [
      { name: 'author', label: 'Tu nombre (autor de documentos, comentarios y cambios)', value: this.prefs.author },
      { name: 'theme', label: 'Tema de la interfaz', type: 'select', value: this.prefs.theme, options: [['auto', 'Igual que el sistema'], ['light', 'Claro'], ['dark', 'Oscuro']] },
      { name: 'autosaveSec', label: 'Guardar copia de recuperación cada (segundos)', type: 'number', value: this.prefs.autosaveSec, min: 5, max: 600 },
      { name: 'spellcheck', label: 'Revisar la ortografía mientras escribes', type: 'checkbox', value: this.prefs.spellcheck !== false },
      { type: 'note', label: 'Las opciones se aplican a todas las apps de la suite.' },
    ], 'Guardar');
    if (!v) return;
    v.autosaveSec = Math.max(5, Math.min(600, Number(v.autosaveSec) || 20));
    this.prefs = { ...this.prefs, ...v };
    if (api?.setPrefs) await api.setPrefs(v); else this.applyPrefs();
    toast('Opciones guardadas');
  }
  docProps() { return { title: this.props.title || this.doc.name.replace(/\.[^.]+$/, ''), author: this.props.author || this.prefs.author || '', subject: this.props.subject || '', keywords: this.props.keywords || '' }; }

  buildTop() {
    this.top.append(
      h('div', { class: 'brand', title: 'Mostrar Mi Oficina MX', onclick: () => api?.showLauncher() }, h('img', { src: this.logo, alt: '' }), this.appName),
      h('button', { class: 'qa', title: 'Guardar (Ctrl+S)', html: icon('save', 18), onclick: () => this.command('save') }),
      h('button', { class: 'qa', title: 'Deshacer (Ctrl+Z)', html: icon('undo', 18), onclick: () => this.command('undo') }),
      h('button', { class: 'qa', title: 'Rehacer (Ctrl+Y)', html: icon('redo', 18), onclick: () => this.command('redo') }),
    );
    this.docNameEl = h('div', { class: 'docname' });
    this.top.append(this.docNameEl);
    this.searchBox = h('input', { class: 'search', placeholder: 'Buscar (Ctrl+F)', 'aria-label': 'Buscar en el documento' });
    this.searchBox.addEventListener('keydown', (e) => { if (e.key === 'Enter') this.adapter.onCommand('findText', this.searchBox.value); if (e.key === 'Escape') this.searchBox.blur(); });
    this.top.append(this.searchBox);
    this.updateTitle();
  }

  updateTitle() {
    this.docNameEl.innerHTML = `${esc(this.doc.name)}${this.doc.dirty ? '<span class="dirty">• sin guardar</span>' : ''}`;
    document.title = `${this.doc.dirty ? '● ' : ''}${this.doc.name} — ${this.appName}`;
    api?.setState({ dirty: this.doc.dirty, name: this.doc.name, path: this.doc.path });
  }
  markDirty() {
    if (!this.doc.dirty) { this.doc.dirty = true; this.updateTitle(); }
    clearTimeout(this.autosaveT);
    this.autosaveT = setTimeout(() => this.autosave(), (this.prefs?.autosaveSec || 20) * 1000);
  }
  markClean() { this.doc.dirty = false; this.updateTitle(); }
  async autosave() {
    if (!api || !this.doc.dirty) return;
    try { await api.autosave({ id: this.recoveryId, name: this.doc.name, data: this.adapter.snapshot() }); } catch { /* sin espacio */ }
  }

  bindGlobal() {
    api?.onMenu((cmd) => this.command(cmd));
    api?.onOpenPath(async (p) => { this.hideBackstage(); const f = await api.readPath(p); await this.loadFile(f); });
    // arrastrar y soltar
    let hint = null;
    addEventListener('dragover', (e) => { if ([...e.dataTransfer.types].includes('Files')) { e.preventDefault(); if (!hint) { hint = h('div', { class: 'drop-hint' }, 'Suelta el archivo para abrirlo'); document.body.append(hint); } } });
    addEventListener('dragleave', (e) => { if (e.relatedTarget == null) { hint?.remove(); hint = null; } });
    addEventListener('drop', async (e) => {
      hint?.remove(); hint = null;
      const file = e.dataTransfer.files[0];
      if (!file) return;
      const ext = file.name.split('.').pop().toLowerCase();
      if (this.adapter.onDropFile && this.adapter.onDropFile(file, e)) { e.preventDefault(); return; }
      e.preventDefault();
      const p = api?.pathForFile(file);
      if (!this.adapter.exts.includes(ext)) { if (p) api.openInSuite(p); else toast('Ese tipo de archivo no se abre aquí.', { error: true }); return; }
      if (p) { const f = await api.readPath(p); await this.openFileHere(f); }
      else await this.openFileHere({ name: file.name, ext, bytes: new Uint8Array(await file.arrayBuffer()), path: null });
    });
    addEventListener('keydown', (e) => {
      const mod = e.ctrlKey || e.metaKey;
      if (mod && e.key === 'F1') { e.preventDefault(); this.ribbon.toggleCollapse(); }
      if (e.key === 'Escape' && !this.backstage?.hidden && this.backstage && this.doc.started) this.hideBackstage();
    });
    addEventListener('beforeunload', () => { this.autosave(); });
  }

  async command(cmd, arg) {
    try {
      switch (cmd) {
        case 'new': return this.newFromMenu();
        case 'open': return this.open();
        case 'save': return this.save();
        case 'saveAs': return this.save({ ask: true });
        case 'exportPdf': return this.exportPdf();
        case 'print': return this.print();
        case 'about': return this.about();
        case 'options': return this.optionsDialog();
        case 'properties': return this.showBackstage('info');
        case 'find': this.searchBox.focus(); this.searchBox.select(); return this.adapter.onCommand('find', arg);
        default: return this.adapter.onCommand(cmd, arg);
      }
    } catch (err) {
      console.error(err);
      toast(`Ocurrió un error: ${err.message || err}`, { error: true, ms: 6000 });
    }
  }

  async newFromMenu() {
    if (!this.adapter.isPristine() || this.doc.path) { api?.newWindow(); return; }
    this.showBackstage('new');
  }

  async newDoc(tpl) {
    this.props = {};
    await this.adapter.newDoc(tpl);
    this.doc = { name: 'Sin título', path: null, ext: this.adapter.saveExt, dirty: false, started: true };
    this.updateTitle();
    this.hideBackstage();
  }

  async open() {
    if (!api) return;
    const f = await api.openDialog();
    if (!f || f.redirected) return;
    await this.openFileHere(f);
  }

  async openFileHere(f) {
    // Si ya hay un documento con trabajo, se abre en otra ventana.
    if (this.doc.started && (!this.adapter.isPristine() || this.doc.path) && f.path) { api.newWindow({ path: f.path }); return; }
    await this.loadFile(f);
  }

  async loadFile(f) {
    const veil = h('div', { class: 'modal-back' }, h('div', { class: 'modal', style: { padding: '18px 24px', minWidth: '240px' } }, `Abriendo ${f.name}…`));
    document.body.append(veil);
    try {
      this.props = {};
      const info = await this.adapter.load(f);
      if (info?.props) this.props = { ...info.props };
      const nativeSave = this.adapter.writable ? this.adapter.writable.includes(f.ext) : f.ext === this.adapter.saveExt;
      this.doc = { name: f.name, path: nativeSave ? f.path : null, ext: nativeSave ? f.ext : this.adapter.saveExt, dirty: false, started: true };
      if (!nativeSave) this.doc.name = f.name.replace(/\.[^.]+$/, '') + '.' + this.adapter.saveExt;
      this.updateTitle();
      this.hideBackstage();
      if (info?.warnings?.length) toast(`Se abrió con avisos: ${info.warnings.slice(0, 2).join(' · ')}`, { ms: 7000 });
      if (!nativeSave) toast(`Se guardará como .${this.adapter.saveExt}`);
    } catch (err) {
      console.error(err);
      toast(`No se pudo abrir "${f.name}": ${err.message || err}`, { error: true, ms: 7000 });
    } finally { veil.remove(); }
  }

  async save({ ask = false } = {}) {
    if (!api) return false;
    const ext = this.doc.path ? this.doc.ext : this.adapter.saveExt;
    const data = await this.adapter.serialize(ext);
    const res = await api.save({
      ...(typeof data === 'string' ? { text: data } : { bytes: data }),
      path: ask ? null : this.doc.path, ask: ask || !this.doc.path,
      suggestedName: this.doc.name.replace(/\.[^.]+$/, '') + '.' + ext, ext,
    });
    if (!res) return false;
    this.doc.path = res.path; this.doc.name = res.name; this.doc.ext = res.name.split('.').pop().toLowerCase();
    this.markClean();
    api.discardRecovered(this.recoveryId);
    toast(`Guardado: ${res.name}`);
    return true;
  }

  async saveAsFormat(ext, filterName) {
    const data = await this.adapter.serialize(ext);
    const res = await api.save({ ...(typeof data === 'string' ? { text: data } : { bytes: data }), ask: true, ext, filterName, suggestedName: this.doc.name.replace(/\.[^.]+$/, '') + '.' + ext });
    if (res) toast(`Exportado: ${res.name}`, { action: { label: 'Mostrar', run: () => api.showInFolder(res.path) } });
    return res;
  }

  async withPrintLayout(fn) {
    await this.adapter.beforePrint?.(); // se prepara con la vista normal (medidas reales)
    document.body.classList.add('printing');
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    try { return await fn(); } finally {
      document.body.classList.remove('printing');
      await this.adapter.afterPrint?.();
    }
  }

  async exportPdf() {
    if (!api) return;
    const opts = this.adapter.pdfOptions?.() || {};
    const res = await this.withPrintLayout(() => api.exportPdf({ suggestedName: this.doc.name.replace(/\.[^.]+$/, '') + '.pdf', ...opts }));
    if (res) toast(`PDF listo: ${res.name}`, { action: { label: 'Abrir en PapelPDF', run: () => api.openInSuite(res.path) } });
  }

  async print() {
    if (!api) { window.print(); return; }
    await this.withPrintLayout(() => api.print());
  }

  async about() {
    const info = api ? await api.appInfo() : { version: '0.1.0', platform: '', electron: '' };
    modal({
      title: `Acerca de ${this.appName}`,
      body: h('div', { style: { display: 'flex', gap: '16px', alignItems: 'flex-start', maxWidth: '460px' } },
        h('img', { src: this.logo, width: 72, height: 72, alt: '' }),
        h('div', {},
          h('p', { style: { margin: '0 0 6px', fontWeight: 600 } }, `${this.appName} · Mi Oficina MX v${info.version}`),
          h('p', { style: { margin: '0 0 8px' } }, 'Software libre bajo la Licencia Pública General GNU v3 (GPL-3.0). Puedes usarlo, estudiarlo, compartirlo y modificarlo.'),
          h('p', { style: { margin: '0 0 8px', color: 'var(--ink-3)', fontSize: '12px' } }, 'Desarrollado por Santiago Miguel · 744 354 8848 · chago2850@gmail.com'),
          h('p', { style: { margin: 0, color: 'var(--ink-3)', fontSize: '11.5px' } }, `Electron ${info.electron || ''} · ${info.platform || ''} · Sin anuncios, sin cuentas, sin telemetría.`))),
    });
  }

  // -------------------------------------------------------------- inicio / backstage
  async showBackstage(section = 'new') {
    if (!this.backstage) {
      this.backstage = h('section', { class: 'backstage', role: 'dialog', 'aria-label': 'Archivo' });
      document.body.append(this.backstage);
    }
    this.backstage.hidden = false;
    const started = !!this.doc.started;
    const side = h('div', { class: 'side' });
    if (started) side.append(h('button', { class: 'back', html: `${icon('arrow', 18).replace('M3 12h15 M13 6l6 6-6 6', 'M21 12H6 M11 6l-6 6 6 6')}<span>Volver</span>`, onclick: () => this.hideBackstage() }));
    side.append(h('div', { class: 'app' }, h('img', { src: this.logo, alt: '' }), this.appName));
    const item = (id, ic, label, run, disabled) => h('button', { class: section === id ? 'on' : '', html: `${icon(ic, 18)}<span>${label}</span>`, onclick: run, disabled: disabled ? true : null });
    side.append(
      item('new', 'new', 'Inicio', () => this.showBackstage('new')),
      item('open', 'open', 'Abrir', () => this.open()),
      item('save', 'save', 'Guardar', () => this.save().then((ok) => ok && this.hideBackstage()), !started),
      item('saveAs', 'saveAs', 'Guardar como', () => this.save({ ask: true }).then((ok) => ok && this.hideBackstage()), !started),
      item('export', 'pdf', 'Exportar', () => this.showBackstage('export'), !started),
      item('print', 'print', 'Imprimir', () => { this.hideBackstage(); this.print(); }, !started),
      item('info', 'info', 'Información', () => this.showBackstage('info'), !started),
      h('div', { class: 'grow' }),
      item('options', 'brush', 'Opciones', () => this.optionsDialog()),
      item('suite', 'launcher', 'Mi Oficina MX', () => api?.showLauncher()),
      item('about', 'info', 'Acerca de', () => this.about()),
    );
    const main = h('div', { class: 'main' });
    this.backstage.replaceChildren(side, main);
    if (section === 'info' && !started) section = 'new';
    if (section === 'new') await this.renderHome(main);
    if (section === 'export') this.renderExport(main);
    if (section === 'info') this.renderInfo(main);
  }
  hideBackstage() { if (this.backstage) this.backstage.hidden = true; this.adapter.focus?.(); }

  async renderHome(main) {
    const hour = new Date().getHours();
    main.append(h('h1', {}, hour < 12 ? 'Buenos días' : hour < 19 ? 'Buenas tardes' : 'Buenas noches'), h('p', { class: 'hello' }, this.adapter.homeHint || 'Crea algo nuevo o sigue donde te quedaste.'));
    // recuperados
    const rec = api ? await api.recovered() : [];
    if (rec.length) {
      const box = h('div', { class: 'recovered' }, h('h3', { html: `${icon('recover', 18)} Documentos recuperados` }), h('p', { style: { margin: '0 0 6px', fontSize: '12px' } }, 'Se cerró sin guardar. Puedes recuperar estas versiones:'));
      rec.slice(0, 5).forEach((r) => {
        const row = h('div', { class: 'r' }, h('span', {}, `${r.name} · ${new Date(r.savedAt).toLocaleString('es-MX')}`));
        row.append(
          h('button', { class: 'btn primary', onclick: async () => {
            const j = await api.recover(r.id);
            this.recoveryId = r.id;
            await this.adapter.restore(j.data);
            this.doc = { name: j.name || 'Recuperado', path: j.path || null, ext: this.adapter.saveExt, dirty: true, started: true };
            this.updateTitle(); this.hideBackstage();
            toast('Documento recuperado. Guárdalo para conservarlo.');
          } }, 'Recuperar'),
          h('button', { class: 'btn ghost', onclick: async () => { await api.discardRecovered(r.id); row.remove(); } }, 'Descartar'));
        box.append(row);
      });
      main.append(box);
    }
    main.append(h('h2', {}, 'Nuevo'));
    const tpls = h('div', { class: 'tpls' });
    for (const t of this.adapter.templates) {
      const thumb = h('div', { class: `thumb${t.blank ? ' blank' : ''}`, html: t.thumb || '' });
      tpls.append(h('button', { class: 'tpl', 'data-tpl': t.id, onclick: () => this.newDoc(t.id) }, thumb, h('span', { class: 'name' }, t.name)));
    }
    main.append(tpls);
    main.append(h('h2', {}, 'Recientes'));
    const list = h('div', { class: 'recents' });
    const recents = api ? await api.recents() : [];
    if (!recents.length) list.append(h('div', { class: 'empty' }, 'Todavía no hay archivos recientes. Abre uno o arrástralo a esta ventana.'));
    recents.forEach((r) => {
      list.append(h('div', { class: 'recent', tabindex: 0, onclick: async () => { const f = await api.readPath(r.path).catch(() => null); if (f) this.openFileHere(f); else toast('El archivo ya no existe.', { error: true }); } },
        h('span', { html: icon(this.adapter.docIcon || 'doc', 22) }),
        h('div', {}, h('div', {}, r.name), h('div', { class: 'p' }, r.path)),
        h('time', {}, new Date(r.openedAt).toLocaleDateString('es-MX', { day: 'numeric', month: 'short' }))));
    });
    main.append(list);
  }

  renderExport(main) {
    main.append(h('h1', {}, 'Exportar'));
    const opts = [{ id: 'pdf', ic: 'pdf', label: 'PDF', desc: 'Documento fijo para compartir o imprimir. Se puede abrir en PapelPDF.', run: () => { this.hideBackstage(); this.exportPdf(); } }, ...(this.adapter.exportFormats || [])];
    const box = h('div', { style: { display: 'grid', gap: '10px', maxWidth: '560px', marginTop: '18px' } });
    opts.forEach((o) => box.append(h('button', { class: 'recent', style: { gridTemplateColumns: '32px 1fr', textAlign: 'left', border: '1px solid var(--line)', background: '#fff' }, onclick: () => { if (o.run) o.run(); else this.saveAsFormat(o.ext, o.label); } },
      h('span', { html: icon(o.ic || 'doc', 22) }), h('div', {}, h('div', { style: { fontWeight: 600 } }, o.label), h('div', { class: 'p', style: { whiteSpace: 'normal' } }, o.desc)))));
    main.append(box);
  }

  renderInfo(main) {
    main.append(h('h1', {}, 'Información'));
    const grid = h('div', { style: { display: 'grid', gridTemplateColumns: 'minmax(260px, 420px) 1fr', gap: '40px', marginTop: '16px', alignItems: 'start' } });
    const stats = this.adapter.stats?.() || [];
    const tbl = h('table', { style: { borderCollapse: 'collapse' } });
    [['Nombre', this.doc.name], ['Ubicación', this.doc.path || 'Sin guardar'], ['Formato', `.${this.doc.ext}`], ...stats].forEach(([k, v]) => tbl.append(h('tr', {}, h('td', { style: { color: 'var(--ink-3)', padding: '6px 24px 6px 0', verticalAlign: 'top' } }, k), h('td', { style: { padding: '6px 0', wordBreak: 'break-all' } }, String(v)))));
    const left = h('div', {}, h('h2', { style: { marginTop: 0 } }, 'Propiedades'));
    const pr = this.docProps();
    const fields = [['title', 'Título'], ['author', 'Autor'], ['subject', 'Asunto'], ['keywords', 'Palabras clave']];
    const inputs = {};
    fields.forEach(([k, l]) => { inputs[k] = h('input', { value: pr[k] || '', 'data-prop': k }); left.append(h('label', { class: 'field' }, h('span', {}, l), inputs[k])); });
    left.append(h('button', { class: 'btn primary', onclick: () => { fields.forEach(([k]) => { this.props[k] = inputs[k].value.trim(); }); this.markDirty(); toast('Propiedades actualizadas (se guardan con el archivo)'); } }, 'Aplicar'));
    const right = h('div', {}, h('h2', { style: { marginTop: 0 } }, 'Datos'), tbl);
    if (this.doc.path) right.append(h('p', {}, h('button', { class: 'btn', onclick: () => api.showInFolder(this.doc.path) }, 'Mostrar en la carpeta')));
    grid.append(left, right);
    main.append(grid);
  }

  // -------------------------------------------------------------- barra de estado con zoom
  zoomControl({ value = 100, min = 25, max = 300, onChange }) {
    const lbl = h('span', { style: { minWidth: '40px', textAlign: 'right' } }, `${value}%`);
    const range = h('input', { type: 'range', min, max, value, 'aria-label': 'Zoom' });
    const set = (v) => { v = Math.max(min, Math.min(max, Math.round(v))); range.value = v; lbl.textContent = `${v}%`; onChange(v); };
    range.addEventListener('input', () => set(Number(range.value)));
    const box = h('div', { class: 'zoom' },
      h('button', { title: 'Alejar', html: icon('minus', 14), onclick: () => set(Number(range.value) - 10) }), range,
      h('button', { title: 'Acercar', html: icon('plus', 14), onclick: () => set(Number(range.value) + 10) }), lbl);
    box.set = set;
    box.get = () => Number(range.value);
    return box;
  }

  async start() {
    if (this.adapter.onReady) await this.adapter.onReady();
    // Si la ventana se abrió para un archivo, el proceso principal lo enviará; si no, se muestra el inicio.
    setTimeout(() => { if (!this.doc.started) this.showBackstage('new'); }, 60);
  }
}

// utilidades
export const u8ToB64 = (u8) => {
  let s = '';
  for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
  return btoa(s);
};
export const b64ToU8 = (b64) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
export const download = (bytes, name) => { const a = h('a', { href: URL.createObjectURL(new Blob([bytes])), download: name }); a.click(); };
