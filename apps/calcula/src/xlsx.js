// Leer y escribir .xlsx (ExcelJS) y .csv para Calcula Datos. GPL-3.0
import ExcelJS from 'exceljs';
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
        if (Object.keys(st).length) s.styles.set(`${r},${c}`, Object.freeze(st));
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
    if (meta?.sheets?.[ws.name]) s.charts = meta.sheets[ws.name].charts || [];
    if ((ws.getImages?.() || []).length) warnings.push('las imágenes de la hoja no se muestran todavía');
  }
  if (!wb.sheets.length) wb.addSheet('Hoja1');
  wb.active = 0;
  wb.hf.clearUndoStack?.(); wb.undoStack = []; wb.redoStack = [];
  return { warnings };
}

export async function writeXlsx(wb) {
  const x = new ExcelJS.Workbook();
  x.creator = 'Calcula Datos · Mi Oficina MX';
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
    if (s.charts.length) meta.sheets[s.name] = { charts: s.charts };
  }
  if (Object.keys(meta.sheets).length) {
    // Los gráficos se guardan en una hoja oculta para que Calcula Datos los recupere al abrir el archivo.
    const ms = x.addWorksheet(META_SHEET, { state: 'veryHidden' });
    ms.getCell('A1').value = JSON.stringify(meta);
  }
  if (wb.active < x.worksheets.length) x.views = [{ activeTab: wb.active }];
  const buf = await x.xlsx.writeBuffer();
  return new Uint8Array(buf);
}

// ---------------- CSV
export function parseCsv(text) {
  text = text.replace(/^﻿/, '');
  const first = text.split(/\r?\n/, 1)[0] || '';
  const delim = [',', ';', '\t', '|'].map((d) => [d, first.split(d).length]).sort((a, b) => b[1] - a[1])[0][0];
  const rows = []; let row = []; let cur = ''; let q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; }
      else cur += ch;
    } else if (ch === '"') q = true;
    else if (ch === delim) { row.push(cur); cur = ''; }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; }
    else cur += ch;
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  return rows;
}
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
