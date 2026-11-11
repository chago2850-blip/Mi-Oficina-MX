// Escribiente — cuadros de Formato como en LibreOffice Writer: Párrafo (sangrías y espaciado, flujo
// del texto, bordes y sombreado), Fuente (efectos, posición y espaciado), cambiar mayúsculas y
// minúsculas, y marcas de formato. GPL-3.0
import { h, modal, toast, t } from '../../shared/office.js';

let C = null; // { editor, selectedBlocks, restoreSelection, saveSelection, markDirty, schedulePaginate, reflectState, FONTS, applySpanStyles, exec }
export function init(ctx) { C = ctx; }

const PX_PT = 0.75; const PX_CM = 2.54 / 96;
const r2 = (n) => Math.round(n * 100) / 100;
const pxToPt = (px) => r2(px * PX_PT);
const pxToCm = (px) => r2(px * PX_CM);
const lenPx = (v) => { if (!v) return 0; const n = parseFloat(v); if (Number.isNaN(n)) return 0; if (/pt$/.test(v)) return n / PX_PT; if (/cm$/.test(v)) return n / PX_CM; if (/mm$/.test(v)) return n / PX_CM / 10; if (/in$/.test(v)) return n * 96; return n; };

// ------------------------------------------------------------------ cuadro con pestañas
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
const val = (box, name) => { const el = box.querySelector(`[name="${name}"]`); if (!el) return undefined; if (el.type === 'checkbox') return el.checked; if (el.type === 'number') return Number(el.value); return el.value; };

// ------------------------------------------------------------------ Párrafo
const BORDER_STYLES = [['solid', 'Sólido'], ['double', 'Doble'], ['dotted', 'Punteado'], ['dashed', 'Discontinuo']];
export async function paragraphDialog(tab) {
  C.restoreSelection();
  const blocks = C.selectedBlocks().filter((b) => /^(P|H[1-6]|BLOCKQUOTE|LI|DIV)$/.test(b.tagName) && !b.classList.contains('pb'));
  if (!blocks.length) { toast('Coloca el cursor en un párrafo.'); return; }
  const b0 = blocks[0]; const cs = getComputedStyle(b0);
  const mt = b0.dataset.push ? lenPx(b0.dataset.mt0) || 0 : parseFloat(cs.marginTop) || 0;
  const ti = parseFloat(cs.textIndent) || 0;
  const lhInline = b0.style.lineHeight;
  let lhKind = 'mult'; let lhVal = 1.15;
  if (lhInline) {
    if (/pt|px/.test(lhInline)) { lhKind = b0.dataset.lr === 'atLeast' ? 'atLeast' : 'exact'; lhVal = pxToPt(lenPx(lhInline)); }
    else lhVal = Number(lhInline) || 1.15;
  } else { const n = parseFloat(cs.lineHeight); const fs = parseFloat(cs.fontSize) || 14.67; if (n) lhVal = r2(n / fs); }
  const lhPreset = lhKind !== 'mult' ? lhKind : ({ 1: 'single', 1.15: '1.15', 1.5: '1.5', 2: 'double' }[lhVal] || 'mult');
  const bt = b0.style;
  const anyBorder = ['Top', 'Right', 'Bottom', 'Left'].filter((s) => bt[`border${s}Style`] && bt[`border${s}Style`] !== 'none');
  const borderPreset = !anyBorder.length ? 'none' : anyBorder.length === 4 ? 'box' : anyBorder.join(',') === 'Bottom' ? 'bottom' : anyBorder.join(',') === 'Top' ? 'top' : anyBorder.join(',') === 'Top,Bottom' ? 'topbottom' : anyBorder.join(',') === 'Left' ? 'left' : 'box';
  const bs = anyBorder[0] ? bt[`border${anyBorder[0]}Style`] : 'solid';
  const bw = anyBorder[0] ? pxToPt(lenPx(bt[`border${anyBorder[0]}Width`])) : 1;
  const bc = anyBorder[0] ? toHex(bt[`border${anyBorder[0]}Color`]) : '#000000';
  const pad = anyBorder[0] ? pxToPt(lenPx(bt[`padding${anyBorder[0]}`])) : 4;
  const shade = bt.backgroundColor ? toHex(bt.backgroundColor) : '';
  const align = cs.textAlign === 'start' ? 'left' : cs.textAlign === 'end' ? 'right' : cs.textAlign;
  const lhIn = numIn('lhv', lhVal, { step: lhKind === 'mult' ? 0.05 : 0.5, min: 0, max: 400 });
  const lhSel = sel('lh', lhPreset, [['single', 'Sencillo'], ['1.15', '1.15 líneas'], ['1.5', '1.5 líneas'], ['double', 'Doble'], ['mult', 'Múltiple'], ['atLeast', 'Al menos'], ['exact', 'Exacto']]);
  const syncLh = () => { const v = lhSel.value; const inp = lhIn.querySelector('input'); inp.disabled = !['mult', 'atLeast', 'exact'].includes(v); lhIn.querySelector('small')?.remove(); if (v === 'atLeast' || v === 'exact') { lhIn.append(h('small', {}, 'pt')); if (Number(inp.value) < 4) inp.value = 14; inp.step = 0.5; } else if (v === 'mult') { if (Number(inp.value) > 10) inp.value = 1.15; inp.step = 0.05; } };
  lhSel.addEventListener('change', syncLh);
  const special = ti > 0 ? 'first' : ti < 0 ? 'hang' : 'none';
  const indentTab = h('div', { class: 'fx-grid' },
    h('h4', {}, t('Sangría')),
    field('Antes del texto', numIn('ml', pxToCm(parseFloat(cs.marginLeft) || 0), { unit: 'cm', step: 0.1 })),
    field('Después del texto', numIn('mr', pxToCm(parseFloat(cs.marginRight) || 0), { unit: 'cm', step: 0.1 })),
    field('Especial', sel('special', special, [['none', '(ninguna)'], ['first', 'Primera línea'], ['hang', 'Francesa']])),
    field('En', numIn('ti', pxToCm(Math.abs(ti)), { unit: 'cm', step: 0.1, min: 0 })),
    h('h4', {}, t('Espaciado')),
    field('Encima del párrafo', numIn('mt', pxToPt(mt), { unit: 'pt', step: 1, min: 0 })),
    field('Debajo del párrafo', numIn('mb', pxToPt(parseFloat(cs.marginBottom) || 0), { unit: 'pt', step: 1, min: 0 })),
    h('h4', {}, t('Interlineado')),
    field('Interlineado', lhSel), field('De', lhIn),
    h('h4', {}, t('Alineación')),
    field('Alineación', sel('align', align || 'left', [['left', 'Izquierda'], ['center', 'Centrado'], ['right', 'Derecha'], ['justify', 'Justificado']]), 'wide'));
  const flowTab = h('div', { class: 'fx-col' },
    h('h4', {}, t('Saltos')),
    chk('pbb', 'Salto de página antes del párrafo', !!b0.dataset.pbb),
    h('h4', {}, t('Opciones')),
    chk('keepnext', 'Conservar con el párrafo siguiente', !!b0.dataset.keepnext),
    chk('keeplines', 'No dividir el párrafo', !!b0.dataset.keeplines),
    chk('widow', 'Control de viudas y huérfanas', b0.dataset.widow !== '0'));
  const shadeOn = chk('shadeOn', 'Color de fondo', !!shade);
  const borderTab = h('div', { class: 'fx-grid' },
    h('h4', {}, t('Bordes')),
    field('Lados con borde', sel('border', borderPreset, [['none', 'Sin borde'], ['box', 'Cuadro'], ['top', 'Solo arriba'], ['bottom', 'Solo abajo'], ['topbottom', 'Arriba y abajo'], ['left', 'Solo a la izquierda']])),
    field('Estilo', sel('bstyle', bs || 'solid', BORDER_STYLES)),
    field('Ancho', numIn('bw', bw || 1, { unit: 'pt', step: 0.25, min: 0.25, max: 12 })),
    field('Color', h('input', { type: 'color', name: 'bcolor', value: bc || '#000000' })),
    field('Espacio al texto', numIn('bpad', pad, { unit: 'pt', step: 1, min: 0, max: 60 })),
    h('h4', {}, t('Sombreado')),
    shadeOn, field('Color', h('input', { type: 'color', name: 'shade', value: shade || '#FFF2CC' })));
  syncLh();
  const box = await tabbed('Párrafo', [{ label: 'Sangría y espaciado', body: indentTab }, { label: 'Flujo del texto', body: flowTab }, { label: 'Bordes y sombreado', body: borderTab }], { start: tab === 'borders' ? 2 : tab === 'flow' ? 1 : 0 });
  if (!box) return;
  C.restoreSelection();
  const sp = val(box, 'special'); const tiCm = val(box, 'ti');
  const lh = val(box, 'lh'); const lhv = val(box, 'lhv');
  for (const b of blocks) {
    const st = b.style;
    const pt = (n) => `${r2(n)}pt`; const cm = (n) => `${r2(n)}cm`;
    st.marginLeft = val(box, 'ml') ? cm(val(box, 'ml')) : '';
    st.marginRight = val(box, 'mr') ? cm(val(box, 'mr')) : '';
    st.textIndent = sp === 'none' || !tiCm ? '' : cm(sp === 'hang' ? -tiCm : tiCm);
    if (sp === 'hang' && tiCm && !val(box, 'ml')) st.marginLeft = cm(tiCm);
    if (b.dataset.push) b.dataset.mt0 = pt(val(box, 'mt')); else st.marginTop = pt(val(box, 'mt'));
    st.marginBottom = pt(val(box, 'mb'));
    delete b.dataset.lr;
    if (lh === 'single') st.lineHeight = '1'; else if (lh === '1.15') st.lineHeight = '1.15'; else if (lh === '1.5') st.lineHeight = '1.5'; else if (lh === 'double') st.lineHeight = '2';
    else if (lh === 'mult') st.lineHeight = String(Math.max(0.5, lhv || 1));
    else { st.lineHeight = pt(Math.max(1, lhv || 12)); if (lh === 'atLeast') b.dataset.lr = 'atLeast'; }
    st.textAlign = val(box, 'align') === 'left' && !b.style.textAlign ? '' : val(box, 'align');
    for (const k of ['pbb', 'keepnext', 'keeplines']) { if (val(box, k)) b.dataset[k] = '1'; else delete b.dataset[k]; }
    if (val(box, 'widow')) delete b.dataset.widow; else b.dataset.widow = '0';
    const sides = { none: [], box: ['Top', 'Right', 'Bottom', 'Left'], top: ['Top'], bottom: ['Bottom'], topbottom: ['Top', 'Bottom'], left: ['Left'] }[val(box, 'border')];
    for (const s of ['Top', 'Right', 'Bottom', 'Left']) {
      if (sides.includes(s)) { st[`border${s}`] = `${r2(val(box, 'bw'))}pt ${val(box, 'bstyle')} ${val(box, 'bcolor')}`; st[`padding${s}`] = pt(val(box, 'bpad')); }
      else { st[`border${s}`] = ''; st[`padding${s}`] = ''; }
    }
    st.backgroundColor = val(box, 'shadeOn') ? val(box, 'shade') : '';
  }
  C.markDirty(); C.schedulePaginate();
}

// ------------------------------------------------------------------ Fuente
export async function fontDialog() {
  C.restoreSelection();
  const s = document.getSelection();
  if (!s.rangeCount) return;
  const node = s.anchorNode?.nodeType === 3 ? s.anchorNode.parentElement : s.anchorNode;
  if (!node || !C.editor.contains(node)) { toast('Selecciona texto en el documento.'); return; }
  const cs = getComputedStyle(node);
  const fam = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
  const inDeco = (el, re) => { let e = el; while (e && e !== C.editor) { const d = getComputedStyle(e); if (re.test(`${d.textDecorationLine} ${d.textDecorationStyle}`)) return d; e = e.parentElement; } return null; };
  const uDeco = inDeco(node, /underline/);
  const sDeco = inDeco(node, /line-through/);
  const under = uDeco ? ({ double: 'double', dotted: 'dotted', dashed: 'dashed', wavy: 'wavy' }[uDeco.textDecorationStyle] || 'single') : 'none';
  const fonts = [...new Set([fam, ...C.FONTS.map((f) => f[0])])];
  const weight = Number(cs.fontWeight) >= 600; const italic = cs.fontStyle === 'italic';
  const preview = h('div', { class: 'fx-preview' }, h('span', {}, s.toString().trim().slice(0, 40) || 'AaBbYyZz'));
  const tab1 = h('div', { class: 'fx-grid' },
    field('Familia', sel('font', fam, fonts.map((f) => [f, f])), 'wide'),
    field('Estilo', sel('fstyle', weight && italic ? 'bi' : weight ? 'b' : italic ? 'i' : 'n', [['n', 'Normal'], ['b', 'Negrita'], ['i', 'Cursiva'], ['bi', 'Negrita cursiva']])),
    field('Tamaño', numIn('size', pxToPt(parseFloat(cs.fontSize)), { unit: 'pt', step: 0.5, min: 1, max: 400 })),
    field('Color de fuente', h('input', { type: 'color', name: 'color', value: toHex(cs.color) || '#000000' })),
    field('Subrayado', sel('under', under, [['none', '(ninguno)'], ['single', 'Sencillo'], ['double', 'Doble'], ['dotted', 'Punteado'], ['dashed', 'Discontinuo'], ['wavy', 'Ondulado']])));
  const tab2 = h('div', { class: 'fx-col fx-two' },
    chk('strike', 'Tachado', !!sDeco && sDeco.textDecorationStyle !== 'double'),
    chk('dstrike', 'Doble tachado', !!sDeco && sDeco.textDecorationStyle === 'double'),
    chk('sup', 'Superíndice', !!node.closest('sup')),
    chk('sub', 'Subíndice', !!node.closest('sub')),
    chk('smallcaps', 'Versalitas', cs.fontVariantCaps === 'small-caps' || cs.fontVariant === 'small-caps'),
    chk('caps', 'Mayúsculas', cs.textTransform === 'uppercase'),
    chk('hidden', 'Oculto', !!node.closest('.hidden-text')),
    chk('shadow', 'Sombra', cs.textShadow && cs.textShadow !== 'none'),
    chk('outline', 'Contorno', !!(cs.webkitTextStroke && parseFloat(cs.webkitTextStrokeWidth) > 0)));
  const ls = parseFloat(cs.letterSpacing) || 0;
  const posEl = node.closest('span[style*="vertical-align"]');
  const pos = posEl && /pt|px/.test(posEl.style.verticalAlign) ? pxToPt(lenPx(posEl.style.verticalAlign)) : 0;
  const scEl = node.closest('span[style*="scaleX"]'); const sc = scEl ? Math.round(parseFloat(/scaleX\(([\d.]+)\)/.exec(scEl.style.transform)?.[1] || 1) * 100) : 100;
  const tab3 = h('div', { class: 'fx-grid' },
    field('Espaciado entre caracteres', numIn('ls', pxToPt(ls), { unit: 'pt', step: 0.1, min: -10, max: 100 })),
    field('Escala de ancho', numIn('scale', sc, { unit: '%', step: 1, min: 1, max: 600 })),
    field('Posición (elevar o bajar)', numIn('pos', pos, { unit: 'pt', step: 0.5, min: -100, max: 100 })));
  const all = h('div'); // para leer los valores de las tres pestañas
  const v = (n) => { for (const tb of [tab1, tab2, tab3]) { const x = val(tb, n); if (x !== undefined) return x; } return undefined; };
  const upd = () => {
    const sp = preview.firstChild; const st = sp.style; const fsv = v('fstyle'); const uv = v('under');
    st.fontFamily = `'${v('font')}'`; st.fontSize = `${Math.min(36, v('size') || 11)}pt`; st.color = v('outline') ? 'transparent' : v('color');
    st.fontWeight = fsv === 'b' || fsv === 'bi' ? 'bold' : 'normal'; st.fontStyle = fsv === 'i' || fsv === 'bi' ? 'italic' : 'normal';
    const lines = []; if (uv !== 'none') lines.push('underline'); if (v('strike') || v('dstrike')) lines.push('line-through');
    st.textDecorationLine = lines.join(' ') || 'none'; st.textDecorationStyle = v('dstrike') ? 'double' : uv === 'none' || uv === 'single' ? 'solid' : uv;
    st.fontVariant = v('smallcaps') ? 'small-caps' : ''; st.textTransform = v('caps') ? 'uppercase' : '';
    st.textShadow = v('shadow') ? '1px 1px 1px rgba(0,0,0,.45)' : ''; st.webkitTextStroke = v('outline') ? `.6px ${v('color')}` : '';
    st.letterSpacing = v('ls') ? `${v('ls')}pt` : ''; st.verticalAlign = v('pos') ? `${v('pos')}pt` : '';
    st.display = 'inline-block'; st.transform = v('scale') !== 100 ? `scaleX(${v('scale') / 100})` : ''; st.opacity = v('hidden') ? '.45' : '';
  };
  for (const tb of [tab1, tab2, tab3]) { tb.addEventListener('input', upd); tb.addEventListener('change', upd); }
  upd(); void all;
  const box = await tabbed('Fuente', [{ label: 'Fuente', body: tab1 }, { label: 'Efectos tipográficos', body: tab2 }, { label: 'Posición y espaciado', body: tab3 }], { preview });
  if (!box) return;
  C.restoreSelection();
  const fs = val(box, 'fstyle'); const u = val(box, 'under');
  // negrita, cursiva, subrayado sencillo y tachado: con los comandos del editor (se guardan como <b>, <i>, <u>, <s>)
  const want = { bold: fs === 'b' || fs === 'bi', italic: fs === 'i' || fs === 'bi', underline: u === 'single', strikeThrough: !!val(box, 'strike') && !val(box, 'dstrike'), superscript: !!val(box, 'sup'), subscript: !!val(box, 'sub') && !val(box, 'sup') };
  for (const [cmd, on] of Object.entries(want)) if (document.queryCommandState(cmd) !== on) C.exec(cmd);
  const deco2 = [];
  if (u !== 'none' && u !== 'single') deco2.push('underline');
  if (val(box, 'dstrike')) deco2.push('line-through');
  const styles = {
    fontFamily: (C.FONTS.find((f) => f[0] === val(box, 'font')) || [0, `'${val(box, 'font')}'`])[1],
    fontSize: `${val(box, 'size')}pt`,
    color: val(box, 'color'),
    textDecorationLine: deco2.join(' '),
    textDecorationStyle: deco2.length ? (val(box, 'dstrike') ? 'double' : u) : '',
    fontVariant: val(box, 'smallcaps') ? 'small-caps' : '',
    textTransform: val(box, 'caps') ? 'uppercase' : '',
    textShadow: val(box, 'shadow') ? '1px 1px 1px rgba(0,0,0,.45)' : '',
    webkitTextStroke: val(box, 'outline') ? '.6px currentColor' : '',
    letterSpacing: val(box, 'ls') ? `${val(box, 'ls')}pt` : '',
    verticalAlign: val(box, 'pos') ? `${val(box, 'pos')}pt` : '',
    display: val(box, 'scale') !== 100 ? 'inline-block' : '',
    transform: val(box, 'scale') !== 100 ? `scaleX(${val(box, 'scale') / 100})` : '',
    transformOrigin: val(box, 'scale') !== 100 ? 'left' : '',
  };
  if (val(box, 'outline')) { styles.webkitTextStroke = `.6px ${val(box, 'color')}`; styles.color = 'transparent'; }
  C.applySpanStyles(styles, { hidden: val(box, 'hidden') });
  C.markDirty(); C.schedulePaginate(); C.reflectState?.();
}

// ------------------------------------------------------------------ mayúsculas y minúsculas
export function changeCase(mode) {
  C.restoreSelection();
  const s = document.getSelection();
  if (!s.rangeCount || s.isCollapsed) { toast('Selecciona el texto que quieres cambiar.'); return; }
  const r = s.getRangeAt(0);
  const w = document.createTreeWalker(r.commonAncestorContainer.nodeType === 3 ? r.commonAncestorContainer.parentNode : r.commonAncestorContainer, NodeFilter.SHOW_TEXT);
  const nodes = []; while (w.nextNode()) if (r.intersectsNode(w.currentNode)) nodes.push(w.currentNode);
  let atSentence = true; let first = true;
  // el inicio de oración se decide con el texto anterior a la selección
  if (mode === 'sentence') { const pre = document.createRange(); pre.setStart(C.editor, 0); pre.setEnd(r.startContainer, r.startOffset); const tx = pre.toString().trimEnd(); atSentence = !tx || /[.!?¿¡:]\s*$|\n$/.test(tx) || r.startOffset === 0; }
  const startN = r.startContainer; const startO = r.startOffset; const endN = r.endContainer; const endO = r.endOffset;
  for (const n of nodes) {
    const a = n === startN ? startO : 0; const b = n === endN ? endO : n.nodeValue.length;
    const seg = n.nodeValue.slice(a, b);
    let out = seg;
    if (mode === 'upper') out = seg.toLocaleUpperCase('es');
    else if (mode === 'lower') out = seg.toLocaleLowerCase('es');
    else if (mode === 'toggle') out = [...seg].map((c) => (c === c.toLocaleUpperCase('es') ? c.toLocaleLowerCase('es') : c.toLocaleUpperCase('es'))).join('');
    else if (mode === 'title') { out = seg.toLocaleLowerCase('es').replace(/(^|[\s\-–—(«"'¿¡])(\p{L})/gu, (m, p, c) => p + c.toLocaleUpperCase('es')); if (!first && a === 0 && /\p{L}/u.test(seg[0] || '') && !/\s$/.test(n.previousSibling?.textContent || ' ')) out = (seg[0] || '').toLocaleLowerCase('es') + out.slice(1); }
    else if (mode === 'sentence') { out = ''; for (const c of seg.toLocaleLowerCase('es')) { if (atSentence && /\p{L}/u.test(c)) { out += c.toLocaleUpperCase('es'); atSentence = false; } else { out += c; if (/[.!?]/.test(c)) atSentence = true; } } }
    if (out !== seg) n.replaceData(a, b - a, out);
    first = false;
  }
  const nr = document.createRange(); nr.setStart(startN, Math.min(startO, startN.nodeValue?.length ?? startO)); nr.setEnd(endN, Math.min(endO, endN.nodeValue?.length ?? endO));
  s.removeAllRanges(); s.addRange(nr);
  C.saveSelection(); C.markDirty();
}
export function caseMenu() {
  return [
    { label: 'Tipo oración.', run: () => changeCase('sentence') },
    { label: 'minúsculas', run: () => changeCase('lower') },
    { label: 'MAYÚSCULAS', run: () => changeCase('upper') },
    { label: 'Poner En Mayúsculas Cada Palabra', run: () => changeCase('title') },
    { label: 'aLTERNAR mAY/min', run: () => changeCase('toggle') },
  ];
}

// ------------------------------------------------------------------ marcas de formato
export function toggleMarks(on) {
  const v = on ?? !C.editor.classList.contains('marks');
  C.editor.classList.toggle('marks', v);
  return v;
}

function toHex(c) {
  if (!c) return '';
  if (c.startsWith('#')) return c.length === 4 ? `#${[...c.slice(1)].map((x) => x + x).join('')}` : c.slice(0, 7);
  const m = /rgba?\(([^)]+)\)/.exec(c); if (!m) return '';
  const [r, g, b, a] = m[1].split(',').map((x) => parseFloat(x));
  if (a === 0) return '';
  return `#${[r, g, b].map((n) => Math.round(n).toString(16).padStart(2, '0')).join('')}`;
}
