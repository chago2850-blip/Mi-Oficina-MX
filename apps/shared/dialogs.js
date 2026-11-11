// Cuadros de diálogo con pestañas como los de LibreOffice (Formato de celdas, Párrafo…), compartidos por las apps. GPL-3.0
import { h, modal, t } from './office.js';

export function tabbed(title, tabs, { preview, width = '560px', start = 0 } = {}) {
  const heads = h('div', { class: 'fx-tabs', role: 'tablist' });
  const panes = h('div', { class: 'fx-panes' });
  tabs.forEach((tb, i) => {
    const id = `fx${Math.random().toString(36).slice(2, 7)}`;
    const b = h('button', { type: 'button', class: `fx-tab${i === start ? ' on' : ''}`, role: 'tab', 'aria-selected': String(i === start), 'aria-controls': id, onclick: () => { heads.querySelectorAll('.fx-tab').forEach((x, k) => { x.classList.toggle('on', k === i); x.setAttribute('aria-selected', String(k === i)); }); panes.querySelectorAll('.fx-pane').forEach((x, k) => { x.hidden = k !== i; }); } }, t(tb.label));
    heads.append(b);
    panes.append(h('div', { class: 'fx-pane', id, role: 'tabpanel', hidden: i !== start ? true : null }, tb.body));
  });
  const body = h('div', { class: 'fx' }, heads, panes, preview || '');
  return modal({ title: t(title), body, width, buttons: [{ label: 'Cancelar', value: null }, { label: 'Aceptar', primary: true, value: true }] }).then((ok) => (ok ? body : null));
}
export const field = (label, el, extra = '') => h('label', { class: `fx-field ${extra}` }, h('span', {}, t(label)), el);
export const numIn = (name, value, { step = 0.1, min = -100, max = 1000, unit = '' } = {}) => h('span', { class: 'fx-num' }, h('input', { type: 'number', name, value, step, min, max }), unit ? h('small', {}, unit) : '');
export const sel = (name, value, options) => h('select', { name }, ...options.map(([v, l]) => h('option', { value: v, selected: String(v) === String(value) ? true : null }, t(l))));
export const chk = (name, label, on) => h('label', { class: 'fx-check' }, h('input', { type: 'checkbox', name, checked: on ? true : null }), ` ${t(label)}`);
