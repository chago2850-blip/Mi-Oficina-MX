// Leer y escribir .xlsx (ExcelJS) y .csv para Calcula Datos. GPL-3.0
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { dateToSerial, isDateFormat } from './format.js';
import { DEF_W, DEF_H, parseRange, colName } from './workbook.js';
import { errorEnglish } from './i18n.js';

const META_SHEET = '_calcula_datos';
const argbToHex = (argb) => (argb && /^[0-9A-F]{8}$/i.test(argb) ? `#${argb.slice(2)}` : (argb && /^[0-9A-F]{6}$/i.test(argb) ? `#${argb}` : null));
const hexToArgb = (hex) => (hex && hex.startsWith('#') ? `FF${hex.slice(1).toUpperCase().padEnd(6, '0').slice(0, 6)}` : undefined);
const THEME = ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
const colorOf = (c) => (!c ? null : c.argb ? argbToHex(c.argb) : (c.theme !== undefined ? THEME[c.theme] || null : null));

export async function readXlsx(wb, bytes) {
  const x = new ExcelJS.Workbook();
  await x.xlsx.load(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
  wb.reset();
  const warnings = [];
  let meta = null;
  const metaWs = x.getWorksheet(META_SHEET);
  if (metaWs) { try { meta = JSON.parse(metaWs.getCell('A1').value); } catch { meta = null; } }
  for (const ws of x.worksheets) {
    if (ws.name === META_SHEET) continue;
    const s = wb.addSheet(ws.name);
    const rows = [];
    ws.eachRow({ includeEmpty: false }, (row, rn) => {
      row.eachCell({ includeEmpty: false }, (cell, cn) => {
        const r = rn - 1, c = cn - 1;
        let v = cell.value;
        let fmt = cell.numFmt && cell.numFmt !== 'General' ? cell.numFmt : null;
        if (v && typeof v === 'object') {
          if (v instanceof Date) { v = dateToSerial(v); if (!fmt || !isDateFormat(fmt)) fmt = 'dd/mm/yyyy'; }
          else if ('formula' in v || 'sharedFormula' in v) { const f = cell.formula; v = f ? `=${f}` : (v.result ?? null); }
          else if ('richText' in v) v = v.richText.map((t) => t.text).join('');
          else if ('text' in v && 'hyperlink' in v) v = v.text;
          else if ('error' in v) v = v.error;
          else v = String(v.result ?? '');
        }
        (rows[r] ||= [])[c] = v;
        const st = {};
        const f = cell.font || {};
        if (f.bold) st.b = true; if (f.italic) st.i = true; if (f.underline) st.u = true; if (f.strike) st.s = true;
        if (f.size && f.size !== 11) st.size = f.size;
        if (f.name && !/^Calibri$/i.test(f.name)) st.font = f.name;
        const fc = colorOf(f.color); if (fc && fc !== '#000000') st.color = fc;
        const fill = cell.fill;
        if (fill && fill.type === 'pattern' && fill.pattern === 'solid') { const bg = colorOf(fill.fgColor); if (bg) st.bg = bg; }
        const al = cell.alignment || {};
        if (al.horizontal && ['left', 'center', 'right'].includes(al.horizontal)) st.align = al.horizontal;
        if (al.vertical && ['top', 'middle', 'bottom'].includes(al.vertical)) st.valign = al.vertical;
        if (al.wrapText) st.wrap = true;
        if (fmt) st.fmt = fmt;
        const bd = cell.border || {};
        const b = {};
        for (const side of ['top', 'bottom', 'left', 'right']) if (bd[side]?.style) b[side[0]] = colorOf(bd[side].color) || '#000000';
        if (Object.keys(b).length) st.border = b;
        if (cell.protection && cell.protection.locked === false) st.unlocked = true;
        if (Object.keys(st).length) s.styles.set(`${r},${c}`, Object.freeze(st));
        if (cell.note) { const t = typeof cell.note === 'string' ? cell.note : (cell.note.texts || []).map((x) => x.text).join(''); if (t) s.notes.set(`${r},${c}`, t); }
      });
    });
    // Rellena huecos para HyperFormula
    const data = Array.from({ length: rows.length }, (_, i) => rows[i] || []);
    wb.hf.setSheetContent(s.id, data);
    // tamaños
    (ws.columns || []).forEach((col, i) => { if (col && col.width && Math.abs(col.width - 8.43) > 0.2) s.colW.set(i, Math.round(col.width * 7 + 5)); if (col?.hidden) s.colW.set(i, 0); });
    ws.eachRow({ includeEmpty: true }, (row, rn) => { if (row.height && Math.abs(row.height - 15) > 0.5) s.rowH.set(rn - 1, Math.round(row.height * 4 / 3)); if (row.hidden) s.hiddenRows.add(rn - 1); });
    // combinadas
    const merges = ws.model?.merges || [];
    s.merges = merges.map((m) => parseRange(m)).filter(Boolean);
    // inmovilizar
    const v = (ws.views || [])[0];
    if (v && v.state === 'frozen') s.freeze = { rows: v.ySplit || 0, cols: v.xSplit || 0 };
    const u = wb.usedRange(s); s.nRows = Math.max(1000, u.r2 + 200); s.nCols = Math.max(52, u.c2 + 20);
    if (meta?.sheets?.[ws.name]) { const ms = meta.sheets[ws.name]; s.charts = ms.charts || []; if (ms.pivot) s.pivot = ms.pivot; if (ms.cf) s.cf = ms.cf; }
    // formato condicional nativo (si no venía de Calcula Datos)
    if (!s.cf.length) s.cf = readCf(ws);
    // validaciones
    s.validations = readValidations(ws);
    // protección e impresión
    if (ws.sheetProtection?.sheet) s.protected = { password: '' };
    const pg = ws.pageSetup || {};
    s.pageSetup = { orientation: pg.orientation || 'auto', fit: pg.fitToPage ? (pg.fitToHeight === 0 ? 'width' : 'page') : 'none', titleRows: pg.printTitlesRow || '', gridlines: !!pg.showGridLines || true };
    if (pg.printArea) s.printArea = String(pg.printArea).split(',')[0].replace(/\$/g, '');
    if ((ws.getImages?.() || []).length) warnings.push('las imágenes de la hoja no se muestran todavía');
  }
  if (!wb.sheets.length) wb.addSheet('Hoja1');
  // nombres definidos
  for (const dn of x.definedNames?.model || []) {
    if (!dn.name || dn.name.startsWith('_xlnm')) continue;
    const ref = (dn.ranges || [])[0]; if (!ref) continue;
    try { wb.hf.addNamedExpression(dn.name, `=${ref}`); wb.names.push({ name: dn.name, expr: `=${ref}` }); } catch { /* */ }
  }
  if (meta?.protectedStructure) wb.protectedStructure = true;
  // macros: se guardan tal cual para no perderlas al guardar como .xlsm
  try { const z = await JSZip.loadAsync(bytes); const vba = z.file('xl/vbaProject.bin'); if (vba) { wb.vba = await vba.async('uint8array'); warnings.push('las macros se conservan, pero no se ejecutan'); } } catch { /* */ }
  wb.active = 0;
  wb.hf.clearUndoStack?.(); wb.undoStack = []; wb.redoStack = [];
  return { warnings, props: { title: x.title || '', author: x.creator || '', subject: x.subject || '', keywords: x.keywords || '' } };
}

export async function writeXlsx(wb, props = {}) {
  const x = new ExcelJS.Workbook();
  x.creator = props.author || 'Calcula Datos · Mi Oficina MX';
  x.lastModifiedBy = props.author || '';
  if (props.title) x.title = props.title;
  if (props.subject) x.subject = props.subject;
  if (props.keywords) x.keywords = props.keywords;
  x.created = new Date();
  const meta = { app: 'Calcula Datos', v: 1, sheets: {} };
  for (const s of wb.sheets) {
    const ws = x.addWorksheet(s.name, { views: s.freeze.rows || s.freeze.cols ? [{ state: 'frozen', xSplit: s.freeze.cols, ySplit: s.freeze.rows }] : [] });
    const dims = wb.hf.getSheetDimensions(s.id);
    const used = wb.usedRange(s);
    for (let r = 0; r <= Math.max(dims.height - 1, used.r2); r++) {
      for (let c = 0; c <= Math.max(dims.width - 1, used.c2); c++) {
        const raw = wb.raw(r, c, s);
        const st = s.style(r, c);
        if (raw == null && !st) continue;
        const cell = ws.getCell(r + 1, c + 1);
        const note = s.notes.get(`${r},${c}`); if (note) cell.note = note;
        if (st?.unlocked) cell.protection = { locked: false };
        if (typeof raw === 'string' && raw.startsWith('=')) {
          const res = wb.value(r, c, s);
          cell.value = { formula: raw.slice(1), result: wb.isError(res) ? { error: errorEnglish(res) } : (res ?? undefined) };
        } else if (raw != null) cell.value = raw;
        if (st) {
          const font = {};
          if (st.b) font.bold = true; if (st.i) font.italic = true; if (st.u) font.underline = true; if (st.s) font.strike = true;
          if (st.size) font.size = st.size; if (st.font) font.name = st.font; if (st.color) font.color = { argb: hexToArgb(st.color) };
          if (Object.keys(font).length) cell.font = { name: 'Calibri', size: 11, ...font };
          if (st.bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(st.bg) } };
          const al = {};
          if (st.align) al.horizontal = st.align; if (st.valign) al.vertical = st.valign; if (st.wrap) al.wrapText = true;
          if (Object.keys(al).length) cell.alignment = al;
          if (st.fmt) cell.numFmt = st.fmt;
          if (st.border) {
            const b = {};
            for (const [k, side] of [['t', 'top'], ['b', 'bottom'], ['l', 'left'], ['r', 'right']]) if (st.border[k]) b[side] = { style: 'thin', color: { argb: hexToArgb(st.border[k]) } };
            cell.border = b;
          }
        }
      }
    }
    for (const [c, w] of s.colW) { const col = ws.getColumn(c + 1); if (w === 0) col.hidden = true; else col.width = Math.max(1, (w - 5) / 7); }
    for (const [r, hgt] of s.rowH) ws.getRow(r + 1).height = hgt * 0.75;
    for (const r of s.hiddenRows) ws.getRow(r + 1).hidden = true;
    for (const m of s.merges) ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1);
    if (s.filter) ws.autoFilter = `${colName(s.filter.range.c1)}${s.filter.range.r1 + 1}:${colName(s.filter.range.c2)}${s.filter.range.r2 + 1}`;
    for (const [k, note] of s.notes) { const [r, c] = k.split(',').map(Number); const cell = ws.getCell(r + 1, c + 1); if (!cell.note) cell.note = note; }
    writeCf(ws, s.cf);
    for (const v of s.validations) writeValidation(ws, v);
    const ps = s.pageSetup || {};
    ws.pageSetup = { paperSize: 1, orientation: ps.orientation === 'landscape' ? 'landscape' : 'portrait', fitToPage: ps.fit === 'width' || ps.fit === 'page', fitToWidth: 1, fitToHeight: ps.fit === 'page' ? 1 : 0, showGridLines: ps.gridlines !== false, ...(s.printArea ? { printArea: s.printArea } : {}), ...(ps.titleRows ? { printTitlesRow: ps.titleRows } : {}) };
    if (s.protected) await ws.protect(s.protected.password || '', { selectLockedCells: true, selectUnlockedCells: true });
    if (s.charts.length || s.pivot || s.cf.some((r) => r.type === 'dup' || r.type === 'text')) meta.sheets[s.name] = { charts: s.charts, pivot: s.pivot || undefined, cf: s.cf };
  }
  for (const n of wb.names || []) { try { x.definedNames.add(n.expr.replace(/^=/, ''), n.name); } catch { /* */ } }
  if (wb.protectedStructure) meta.protectedStructure = true;
  if (Object.keys(meta.sheets).length || meta.protectedStructure) {
    // Los gráficos se guardan en una hoja oculta para que Calcula Datos los recupere al abrir el archivo.
    const ms = x.addWorksheet(META_SHEET, { state: 'veryHidden' });
    ms.getCell('A1').value = JSON.stringify(meta);
  }
  if (wb.active < x.worksheets.length) x.views = [{ activeTab: wb.active }];
  const buf = await x.xlsx.writeBuffer();
  return new Uint8Array(buf);
}

// Inserta las macros originales en un .xlsx para guardarlo como .xlsm
export async function injectVba(xlsxBytes, vba) {
  const z = await JSZip.loadAsync(xlsxBytes);
  z.file('xl/vbaProject.bin', vba);
  let ct = await z.file('[Content_Types].xml').async('string');
  if (!ct.includes('Extension="bin"')) ct = ct.replace('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>');
  ct = ct.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml');
  z.file('[Content_Types].xml', ct);
  let rels = await z.file('xl/_rels/workbook.xml.rels').async('string');
  if (!rels.includes('vbaProject')) rels = rels.replace('</Relationships>', '<Relationship Id="rIdVba1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>');
  z.file('xl/_rels/workbook.xml.rels', rels);
  return z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

// ---------------- CSV
import { parseCsv } from './csv.js';
export { parseCsv };
export function toCsv(wb, s = wb.sheet, delim = ',') {
  const u = wb.usedRange(s);
  const out = [];
  for (let r = 0; r <= u.r2; r++) {
    const row = [];
    for (let c = 0; c <= u.c2; c++) {
      const d = wb.display(r, c, s).text;
      row.push(/[",\n;]/.test(d) ? `"${d.replace(/"/g, '""')}"` : d);
    }
    out.push(row.join(delim));
  }
  return `﻿${out.join('\r\n')}`;
}
export { DEF_W, DEF_H };

// ---------------- formato condicional y validación (ExcelJS)
const argb = (hex) => (hex ? `FF${hex.replace('#', '').toUpperCase().slice(0, 6)}` : undefined);
const fromArgb = (c) => (c?.argb ? `#${c.argb.slice(-6)}` : null);
function writeCf(ws, rules) {
  for (const r of rules || []) {
    const style = r.style ? { ...(r.style.bg ? { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: argb(r.style.bg) } } } : {}), ...(r.style.color ? { font: { color: { argb: argb(r.style.color) } } } : {}) } : undefined;
    let rule = null;
    const op = { gt: 'greaterThan', lt: 'lessThan', eq: 'equal', between: 'between' }[r.type];
    if (op) rule = { type: 'cellIs', operator: op, formulae: r.type === 'between' ? [Number(r.value) || r.value, Number(r.value2) || r.value2] : [isNaN(Number(r.value)) ? `"${r.value}"` : Number(r.value)], style };
    else if (r.type === 'text') rule = { type: 'containsText', operator: 'containsText', text: r.text || '', style };
    else if (r.type === 'dup') rule = { type: 'duplicateValues', style };
    else if (r.type === 'top' || r.type === 'bottom') rule = { type: 'top10', rank: Number(r.n) || 10, bottom: r.type === 'bottom', style };
    else if (r.type === 'above' || r.type === 'below') rule = { type: 'aboveAverage', aboveAverage: r.type === 'above', style };
    else if (r.type === 'bar') rule = { type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: argb(r.color || '#5B9BD5') } };
    else if (r.type === 'scale2' || r.type === 'scale3') { const cols = r.colors || (r.type === 'scale3' ? ['#F8696B', '#FFEB84', '#63BE7B'] : ['#FFFFFF', '#63BE7B']); rule = { type: 'colorScale', cfvo: r.type === 'scale3' ? [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }] : [{ type: 'min' }, { type: 'max' }], color: cols.map((c) => ({ argb: argb(c) })) }; }
    else if (r.type === 'icons') rule = { type: 'iconSet', iconSet: r.set === 'lights' ? '3TrafficLights1' : '3Arrows', cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }] };
    if (rule) { try { ws.addConditionalFormatting({ ref: r.range, rules: [{ priority: 1, ...rule }] }); } catch { /* */ } }
  }
}
function readCf(ws) {
  const out = [];
  for (const cf of ws.conditionalFormattings || []) {
    for (const rule of cf.rules || []) {
      const style = rule.style ? { bg: fromArgb(rule.style.fill?.bgColor) || fromArgb(rule.style.fill?.fgColor) || undefined, color: fromArgb(rule.style.font?.color) || undefined } : undefined;
      const base = { id: Math.random().toString(36).slice(2, 9), range: cf.ref.split(' ')[0] };
      const op = { greaterThan: 'gt', lessThan: 'lt', equal: 'eq', between: 'between' }[rule.operator];
      if (rule.type === 'cellIs' && op) out.push({ ...base, type: op, value: String(rule.formulae?.[0] ?? '').replace(/^"|"$/g, ''), value2: rule.formulae?.[1], style });
      else if (rule.type === 'containsText') out.push({ ...base, type: 'text', text: rule.text, style });
      else if (rule.type === 'duplicateValues') out.push({ ...base, type: 'dup', style });
      else if (rule.type === 'top10') out.push({ ...base, type: rule.bottom ? 'bottom' : 'top', n: rule.rank, style });
      else if (rule.type === 'aboveAverage') out.push({ ...base, type: rule.aboveAverage === false ? 'below' : 'above', style });
      else if (rule.type === 'dataBar') out.push({ ...base, type: 'bar', color: fromArgb(rule.color) || '#5B9BD5' });
      else if (rule.type === 'colorScale') out.push({ ...base, type: (rule.color || []).length === 3 ? 'scale3' : 'scale2', colors: (rule.color || []).map(fromArgb).filter(Boolean) });
      else if (rule.type === 'iconSet') out.push({ ...base, type: 'icons', set: /Traffic/i.test(rule.iconSet || '') ? 'lights' : 'arrows' });
    }
  }
  return out;
}
function writeValidation(ws, v) {
  const g = parseRange(v.range); if (!g) return;
  let dv;
  if (v.type === 'list') dv = { type: 'list', allowBlank: true, formulae: [`"${v.list.join(',').replace(/"/g, '')}"`] };
  else if (v.type === 'length') dv = { type: 'textLength', operator: 'between', allowBlank: true, formulae: [Number(v.min) || 0, Number(v.max) || 32767] };
  else dv = { type: v.type === 'whole' ? 'whole' : v.type === 'date' ? 'date' : 'decimal', operator: 'between', allowBlank: true, formulae: [v.min === '' || v.min == null ? -1e15 : v.min, v.max === '' || v.max == null ? 1e15 : v.max] };
  Object.assign(dv, { showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Valor no válido', error: v.message || 'El valor no cumple la validación de datos.' });
  try { ws.dataValidations.add(v.range, dv); } catch { /* */ }
}
function readValidations(ws) {
  const model = ws.dataValidations?.model || {};
  const groups = new Map();
  for (const [ref, dv] of Object.entries(model)) {
    if (!dv || !dv.type) continue;
    const key = JSON.stringify(dv);
    const rg = ref.replace(/^range:/, '');
    if (!groups.has(key)) groups.set(key, { dv, refs: [] });
    groups.get(key).refs.push(rg);
  }
  const out = [];
  for (const { dv, refs } of groups.values()) {
    // une celdas sueltas en un rango que las contenga
    const gs = refs.map((r) => parseRange(r)).filter(Boolean);
    if (!gs.length) continue;
    const g = { r1: Math.min(...gs.map((x) => x.r1)), c1: Math.min(...gs.map((x) => x.c1)), r2: Math.max(...gs.map((x) => x.r2)), c2: Math.max(...gs.map((x) => x.c2)) };
    const range = `${colName(g.c1)}${g.r1 + 1}:${colName(g.c2)}${g.r2 + 1}`;
    const base = { id: Math.random().toString(36).slice(2, 9), range, message: dv.error && dv.error !== 'El valor no cumple la validación de datos.' ? dv.error : undefined };
    if (dv.type === 'list') { const f = String(dv.formulae?.[0] ?? ''); out.push({ ...base, type: 'list', list: f.replace(/^"|"$/g, '').split(',').map((x) => x.trim()).filter(Boolean) }); }
    else if (dv.type === 'textLength') out.push({ ...base, type: 'length', min: dv.formulae?.[0], max: dv.formulae?.[1] });
    else out.push({ ...base, type: dv.type === 'whole' ? 'whole' : dv.type === 'date' ? 'date' : 'decimal', min: Number(dv.formulae?.[0]) <= -1e14 ? '' : dv.formulae?.[0], max: Number(dv.formulae?.[1]) >= 1e14 ? '' : dv.formulae?.[1] });
  }
  return out;
}
