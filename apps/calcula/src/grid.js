// Cuadrícula de Calcula Datos dibujada en <canvas>: rápida con muchas filas. GPL-3.0
import { h, dropdown, closeDropdown, icon, esc, toast } from '../../shared/office.js';
import { colName, addr, rangeStr, parseRange, parseAddr } from './workbook.js';
import { toSpanish, toEnglish } from './i18n.js';
import { parseInput } from './format.js';

const HDR_W = 46, HDR_H = 24;
const GRID = '#E2E4EA', HDR_BG = '#F4F5F8', HDR_SEL = '#DCEFE8', HDR_TXT = '#5E6378';
const FONT_DEF = 'Calibri, Carlito, "Liberation Sans", Arial, sans-serif';
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const DIAS = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo'];

export class Grid {
  constructor(host, wb, { onSelect, onEdit, onChange, onContext, accent = '#0E9C74' }) {
    this.wb = wb; this.accent = accent;
    this.onSelect = onSelect; this.onEdit = onEdit; this.onChange = onChange; this.onContext = onContext;
    this.sel = { r: 0, c: 0 }; this.range = { r1: 0, c1: 0, r2: 0, c2: 0 }; this.anchor = { r: 0, c: 0 };
    this.scrollX = 0; this.scrollY = 0; this.zoom = 1;
    this.clip = null; // { sheet, range, cut, text }
    this.editing = null;
    this.host = host;
    this.canvas = h('canvas', { class: 'grid-canvas', tabindex: '0', 'aria-label': 'Hoja de cálculo' });
    this.scroller = h('div', { class: 'grid-scroll' }, this.spacer = h('div', { class: 'grid-spacer' }));
    this.editor = h('textarea', { class: 'cell-editor', spellcheck: 'false', hidden: true, 'aria-label': 'Editar celda' });
    this.chartLayer = h('div', { class: 'chart-layer' });
    host.append(this.scroller, this.canvas, this.chartLayer, this.editor);
    this.ctx = this.canvas.getContext('2d');
    this.measureCtx = document.createElement('canvas').getContext('2d');
    this.bind();
    new ResizeObserver(() => this.resize()).observe(host);
  }

  get s() { return this.wb.sheet; }

  // ------------------------------------------------------------------ geometría
  invalidate() { this._prefix = null; }
  prefix() {
    const s = this.s;
    if (this._prefix && this._prefix.sheet === s && this._prefix.nR === s.nRows && this._prefix.nC === s.nCols && this._prefix.ver === this.metaVer) return this._prefix;
    const cx = new Float64Array(s.nCols + 1), ry = new Float64Array(s.nRows + 1);
    for (let c = 0; c < s.nCols; c++) cx[c + 1] = cx[c] + s.cw(c);
    for (let r = 0; r < s.nRows; r++) ry[r + 1] = ry[r] + s.rh(r);
    this._prefix = { sheet: s, nR: s.nRows, nC: s.nCols, cx, ry, ver: this.metaVer };
    return this._prefix;
  }
  touchMeta() { this.metaVer = (this.metaVer || 0) + 1; }
  frozen() { const p = this.prefix(), f = this.s.freeze; return { fr: f.rows, fc: f.cols, FW: p.cx[f.cols], FH: p.ry[f.rows] }; }
  colLeft(c) { const p = this.prefix(), { fc, FW } = this.frozen(); return c < fc ? HDR_W + p.cx[c] : HDR_W + FW + p.cx[c] - p.cx[fc] - this.scrollX; }
  rowTop(r) { const p = this.prefix(), { fr, FH } = this.frozen(); return r < fr ? HDR_H + p.ry[r] : HDR_H + FH + p.ry[r] - p.ry[fr] - this.scrollY; }
  cellRect(r, c) {
    const m = this.s.mergeAt(r, c);
    const [r1, c1, r2, c2] = m ? [m.r1, m.c1, m.r2, m.c2] : [r, c, r, c];
    const x = this.colLeft(c1), y = this.rowTop(r1);
    const p = this.prefix();
    return { x, y, w: p.cx[c2 + 1] - p.cx[c1], h: p.ry[r2 + 1] - p.ry[r1] };
  }
  bsearch(arr, v) { let lo = 0, hi = arr.length - 2; while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (arr[mid] <= v) lo = mid; else hi = mid - 1; } return lo; }
  hit(x, y) {
    x /= this.zoom; y /= this.zoom;
    const p = this.prefix(), { fr, fc, FW, FH } = this.frozen();
    let c, r;
    if (x < HDR_W) c = -1; else if (x < HDR_W + FW) c = this.bsearch(p.cx, x - HDR_W); else c = this.bsearch(p.cx, p.cx[fc] + this.scrollX + x - HDR_W - FW);
    if (y < HDR_H) r = -1; else if (y < HDR_H + FH) r = this.bsearch(p.ry, y - HDR_H); else r = this.bsearch(p.ry, p.ry[fr] + this.scrollY + y - HDR_H - FH);
    return { r: Math.min(r, this.s.nRows - 1), c: Math.min(c, this.s.nCols - 1) };
  }

  resize() {
    const dpr = window.devicePixelRatio || 1;
    const w = this.host.clientWidth, hgt = this.host.clientHeight;
    this.W = w; this.H = hgt;
    this.canvas.width = Math.round(w * dpr); this.canvas.height = Math.round(hgt * dpr);
    this.canvas.style.width = `${w}px`; this.canvas.style.height = `${hgt}px`;
    this.dpr = dpr;
    this.updateSpacer();
    this.draw();
  }
  updateSpacer() {
    const p = this.prefix(), { fc, fr } = this.frozen();
    this.spacer.style.width = `${(p.cx[this.s.nCols] - p.cx[fc]) * this.zoom + HDR_W + 200}px`;
    this.spacer.style.height = `${(p.ry[this.s.nRows] - p.ry[fr]) * this.zoom + HDR_H + 200}px`;
  }

  // ------------------------------------------------------------------ dibujo
  requestDraw() { if (this._raf) return; this._raf = requestAnimationFrame(() => { this._raf = 0; this.draw(); }); }
  font(st) { return `${st?.i ? 'italic ' : ''}${st?.b ? 'bold ' : ''}${((st?.size || 11) * 4) / 3}px ${st?.font ? `"${st.font}", ${FONT_DEF}` : FONT_DEF}`; }

  draw() {
    const ctx = this.ctx, s = this.s, wb = this.wb;
    if (!this.W) return;
    ctx.setTransform(this.dpr * this.zoom, 0, 0, this.dpr * this.zoom, 0, 0);
    const W = this.W / this.zoom, H = this.H / this.zoom;
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
    const p = this.prefix(), { fr, fc, FW, FH } = this.frozen();
    // rangos visibles por cuadrante
    const visCols = (frozenPart) => {
      if (frozenPart) return [0, fc - 1];
      const a = this.bsearch(p.cx, p.cx[fc] + this.scrollX), b = this.bsearch(p.cx, p.cx[fc] + this.scrollX + (W - HDR_W - FW));
      return [Math.max(a, fc), Math.min(b, s.nCols - 1)];
    };
    const visRows = (frozenPart) => {
      if (frozenPart) return [0, fr - 1];
      const a = this.bsearch(p.ry, p.ry[fr] + this.scrollY), b = this.bsearch(p.ry, p.ry[fr] + this.scrollY + (H - HDR_H - FH));
      return [Math.max(a, fr), Math.min(b, s.nRows - 1)];
    };
    const quads = [[false, false], [true, false], [false, true], [true, true]];
    for (const [frR, frC] of quads) {
      if ((frR && !fr) || (frC && !fc)) continue;
      const [r1, r2] = visRows(frR), [c1, c2] = visCols(frC);
      if (r2 < r1 || c2 < c1) continue;
      const qx = frC ? HDR_W : HDR_W + FW, qy = frR ? HDR_H : HDR_H + FH;
      const qw = frC ? FW : W - qx, qh = frR ? FH : H - qy;
      ctx.save(); ctx.beginPath(); ctx.rect(qx, qy, qw, qh); ctx.clip();
      this.drawCells(r1, r2, c1, c2);
      ctx.restore();
    }
    this.drawSelection(W, H);
    this.drawHeaders(W, H);
    // líneas de inmovilizado
    ctx.strokeStyle = '#A9AEBF'; ctx.lineWidth = 1.5;
    if (fr) { ctx.beginPath(); ctx.moveTo(0, HDR_H + FH); ctx.lineTo(W, HDR_H + FH); ctx.stroke(); }
    if (fc) { ctx.beginPath(); ctx.moveTo(HDR_W + FW, 0); ctx.lineTo(HDR_W + FW, H); ctx.stroke(); }
    this.positionCharts();
    if (this.editing) this.positionEditor();
  }

  drawCells(r1, r2, c1, c2) {
    const ctx = this.ctx, s = this.s, wb = this.wb;
    // fondos
    for (let r = r1; r <= r2; r++) {
      if (!s.rh(r)) continue;
      const y = this.rowTop(r), hh = s.rh(r);
      for (let c = c1; c <= c2; c++) {
        const st = this.eff(r, c);
        if (st?.bg) { ctx.fillStyle = st.bg; ctx.fillRect(this.colLeft(c), y, s.cw(c), hh); }
        if (st?.bar) { const w = Math.max(0, Math.min(1, st.bar.pct)) * (s.cw(c) - 4); const grd = ctx.createLinearGradient(this.colLeft(c), 0, this.colLeft(c) + w, 0); grd.addColorStop(0, st.bar.color); grd.addColorStop(1, `${st.bar.color}55`); ctx.fillStyle = grd; ctx.fillRect(this.colLeft(c) + 2, y + 3, w, hh - 6); }
      }
    }
    // cuadrícula
    ctx.strokeStyle = GRID; ctx.lineWidth = 1; ctx.beginPath();
    const yTop = this.rowTop(r1), yBot = this.rowTop(r2) + s.rh(r2), xL = this.colLeft(c1), xR = this.colLeft(c2) + s.cw(c2);
    for (let c = c1; c <= c2 + 1; c++) { const x = Math.round(this.colLeft(c)) + 0.5; ctx.moveTo(x, yTop); ctx.lineTo(x, yBot); }
    for (let r = r1; r <= r2 + 1; r++) { if (r <= r2 && !s.rh(r)) continue; const y = Math.round(this.rowTop(r)) + 0.5; ctx.moveTo(xL, y); ctx.lineTo(xR, y); }
    ctx.stroke();
    // celdas combinadas visibles: tapan la cuadrícula interior
    const merges = s.merges.filter((m) => m.r2 >= r1 && m.r1 <= r2 && m.c2 >= c1 && m.c1 <= c2);
    for (const m of merges) {
      const rc = this.cellRect(m.r1, m.c1); const st = s.style(m.r1, m.c1);
      ctx.fillStyle = st?.bg || '#fff'; ctx.fillRect(rc.x + 1, rc.y + 1, rc.w - 1, rc.h - 1);
    }
    // texto
    ctx.textBaseline = 'middle';
    for (let r = r1; r <= r2; r++) {
      const hh = s.rh(r); if (!hh) continue;
      const y = this.rowTop(r);
      for (let c = c1; c <= c2; c++) {
        const m = s.mergeAt(r, c);
        if (m && (m.r1 !== r || m.c1 !== c)) continue;
        if (this.editing && this.editing.r === r && this.editing.c === c) continue;
        if (this.skipText?.(r, c)) continue;
        const d = wb.display(r, c);
        if (!d.text) continue;
        const st = this.eff(r, c);
        const rect = m ? this.cellRect(r, c) : { x: this.colLeft(c), y, w: s.cw(c), h: hh };
        if (st?.icon != null) { this.drawIcon(st.icon, st.iconSet, rect); rect.x += 16; rect.w -= 16; }
        this.drawText(d, st, rect, r, c, c2);
      }
    }
    // texto de combinadas cuyo origen está fuera de la vista
    for (const m of merges) if (m.r1 < r1 || m.c1 < c1) { const d = wb.display(m.r1, m.c1); if (d.text) this.drawText(d, s.style(m.r1, m.c1), this.cellRect(m.r1, m.c1), m.r1, m.c1, c2); }
    // bordes
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
      const st = s.style(r, c); if (!st?.border) continue;
      const rc = this.cellRect(r, c); const b = st.border;
      ctx.lineWidth = 1;
      // cada lado: "#color" o "#color estilo" (thin, medium, thick, dashed, dotted, double, hair)
      const line = (x1, y1, x2, y2, spec) => {
        const [col, sty = 'thin'] = String(spec).split(' ');
        ctx.strokeStyle = col; ctx.lineWidth = sty === 'thick' ? 3 : sty === 'medium' ? 2 : 1;
        ctx.setLineDash(sty === 'dashed' ? [4, 2] : sty === 'dotted' || sty === 'hair' ? [1, 2] : []);
        const draw = (o) => { ctx.beginPath(); if (y1 === y2) { ctx.moveTo(x1, y1 + o); ctx.lineTo(x2, y2 + o); } else { ctx.moveTo(x1 + o, y1); ctx.lineTo(x2 + o, y2); } ctx.stroke(); };
        if (sty === 'double') { draw(-1); draw(1); } else draw(0);
        ctx.setLineDash([]); ctx.lineWidth = 1;
      };
      const X = Math.round(rc.x) + 0.5, Y = Math.round(rc.y) + 0.5, X2 = Math.round(rc.x + rc.w) + 0.5, Y2 = Math.round(rc.y + rc.h) + 0.5;
      if (b.t) line(X, Y, X2, Y, b.t); if (b.b) line(X, Y2, X2, Y2, b.b); if (b.l) line(X, Y, X, Y2, b.l); if (b.r) line(X2, Y, X2, Y2, b.r);
    }
    // minigráficos, imágenes y otros dibujos sobre las celdas
    this.overlay?.(ctx, r1, r2, c1, c2);
    // indicador de notas
    if (s.notes?.size) for (const k of s.notes.keys()) { const [r, c] = k.split(',').map(Number); if (r < r1 || r > r2 || c < c1 || c > c2) continue; const rc = this.cellRect(r, c); ctx.fillStyle = '#D93025'; ctx.beginPath(); ctx.moveTo(rc.x + rc.w - 7, rc.y + 1); ctx.lineTo(rc.x + rc.w, rc.y + 1); ctx.lineTo(rc.x + rc.w, rc.y + 8); ctx.fill(); }
    // botones de filtro
    const f = s.filter;
    if (f && f.range.r1 >= r1 && f.range.r1 <= r2) {
      for (let c = Math.max(c1, f.range.c1); c <= Math.min(c2, f.range.c2); c++) {
        const rc = this.cellRect(f.range.r1, c);
        const bx = rc.x + rc.w - 18, by = rc.y + (rc.h - 16) / 2;
        const active = f.hidden.get(c)?.size;
        ctx.fillStyle = active ? this.accent : '#fff'; ctx.strokeStyle = '#9AA0B4';
        ctx.fillRect(bx, by, 16, 16); ctx.strokeRect(bx + 0.5, by + 0.5, 15, 15);
        ctx.fillStyle = active ? '#fff' : '#4B5066'; ctx.beginPath(); ctx.moveTo(bx + 4, by + 6); ctx.lineTo(bx + 12, by + 6); ctx.lineTo(bx + 8, by + 11); ctx.fill();
      }
    }
  }

  eff(r, c) { const st = this.s.style(r, c); const cf = this.cfStyle ? this.cfStyle(r, c) : null; return cf ? { ...(st || {}), ...cf } : st; }
  drawIcon(level, set, rect) {
    const ctx = this.ctx, x = rect.x + 4, y = rect.y + rect.h / 2;
    const colors = ['#D93025', '#F2A600', '#1E8E3E'];
    ctx.fillStyle = colors[level];
    ctx.beginPath();
    if (set === 'lights') ctx.arc(x + 5, y, 5, 0, Math.PI * 2);
    else if (level === 2) { ctx.moveTo(x, y + 4); ctx.lineTo(x + 5, y - 5); ctx.lineTo(x + 10, y + 4); }
    else if (level === 0) { ctx.moveTo(x, y - 4); ctx.lineTo(x + 5, y + 5); ctx.lineTo(x + 10, y - 4); }
    else { ctx.moveTo(x, y - 5); ctx.lineTo(x + 9, y); ctx.lineTo(x, y + 5); }
    ctx.fill();
  }
  drawText(d, st, rect, r, c, cMax) {
    const ctx = this.ctx, s = this.s;
    ctx.font = this.font(st);
    const link = this.linkAt?.(r, c);
    ctx.fillStyle = d.error ? '#C62828' : (d.color || st?.color || (link ? '#1565C0' : '#111'));
    if (link && !st?.u) st = { ...(st || {}), u: true };
    // texto girado (grados) o apilado en vertical (rot = 'v')
    if (st?.rot) return this.drawRotated(d, st, rect);
    // reducir para ajustar al tamaño de la celda
    if (st?.shrink && !st.wrap) { const tw0 = ctx.measureText(d.text).width; if (tw0 > rect.w - 8) { const k = Math.max(0.3, (rect.w - 8) / tw0); st = { ...st, size: (st.size || 11) * k }; ctx.font = this.font(st); } }
    const pad = 4 + (st?.indent ? st.indent * 9 : 0);
    let align = st?.align || d.align || 'left';
    if (align === 'justify') align = 'left';
    if (align === 'fill') { const one = ctx.measureText(d.text).width || 1; d = { ...d, text: d.text.repeat(Math.max(1, Math.floor((rect.w - 8) / one))) }; align = 'left'; }
    const valign = st?.valign || 'bottom';
    let clipX = rect.x, clipW = rect.w;
    const lines = st?.wrap ? this.wrapLines(d.text, rect.w - pad * 2) : [d.text];
    const tw = st?.wrap ? 0 : ctx.measureText(d.text).width;
    // Desborde hacia las celdas vecinas vacías (texto alineado a la izquierda)
    if (!st?.wrap && tw > rect.w - pad * 2 && typeof this.wb.value(r, c) === 'string' && align === 'left' && !s.mergeAt(r, c)) {
      let cc = c + 1, extra = 0;
      while (cc < s.nCols && extra + rect.w < tw + pad * 2 && this.wb.raw(r, cc) == null && cc <= cMax + 5) { extra += s.cw(cc); cc++; }
      clipW += extra;
      if (extra) { ctx.fillStyle = st?.bg || '#fff'; ctx.fillRect(rect.x + rect.w + 1, rect.y + 1, extra - 2, rect.h - 1); ctx.fillStyle = d.color || st?.color || '#111'; }
    }
    if (!st?.wrap && tw > rect.w - pad * 2 && typeof this.wb.value(r, c) === 'number') { lines[0] = '#'.repeat(Math.max(1, Math.floor((rect.w - pad * 2) / ctx.measureText('#').width))); }
    ctx.save(); ctx.beginPath(); ctx.rect(clipX + 1, rect.y, clipW - 1, rect.h); ctx.clip();
    const lh = ((st?.size || 11) * 4) / 3 * 1.2;
    const total = lh * lines.length;
    let y0 = valign === 'top' ? rect.y + 2 + lh / 2 : valign === 'middle' ? rect.y + rect.h / 2 - total / 2 + lh / 2 : rect.y + rect.h - total + lh / 2 - 2;
    lines.forEach((ln, i) => {
      const w = ctx.measureText(ln).width;
      const x = align === 'center' ? rect.x + rect.w / 2 - w / 2 : align === 'right' ? rect.x + rect.w - pad - w : rect.x + pad;
      const y = y0 + i * lh;
      ctx.fillText(ln, x, y);
      if (st?.u || st?.s) {
        ctx.fillRect(x, st?.u ? y + lh * 0.38 : y, w, 1);
      }
    });
    ctx.restore();
  }
  drawRotated(d, st, rect) {
    const ctx = this.ctx; const lh = ((st.size || 11) * 4) / 3 * 1.2;
    ctx.save(); ctx.beginPath(); ctx.rect(rect.x + 1, rect.y, rect.w - 1, rect.h); ctx.clip();
    const align = st.align || d.align || 'left'; const valign = st.valign || 'bottom';
    if (st.rot === 'v') {
      const chars = [...String(d.text)]; const total = chars.length * lh;
      const x = align === 'center' ? rect.x + rect.w / 2 : align === 'right' ? rect.x + rect.w - 4 - lh / 2 : rect.x + 4 + lh / 2;
      let y = valign === 'top' ? rect.y + 2 + lh / 2 : valign === 'middle' ? rect.y + rect.h / 2 - total / 2 + lh / 2 : rect.y + rect.h - total + lh / 2 - 2;
      ctx.textAlign = 'center'; for (const ch of chars) { ctx.fillText(ch, x, y); y += lh; } ctx.textAlign = 'left';
    } else {
      const a = (Number(st.rot) * Math.PI) / 180; const w = ctx.measureText(d.text).width;
      const bw = Math.abs(w * Math.cos(a)) + Math.abs(lh * Math.sin(a)); const bh = Math.abs(w * Math.sin(a)) + Math.abs(lh * Math.cos(a));
      const cx = align === 'center' ? rect.x + rect.w / 2 : align === 'right' ? rect.x + rect.w - 4 - bw / 2 : rect.x + 4 + bw / 2;
      const cy = valign === 'top' ? rect.y + 2 + bh / 2 : valign === 'middle' ? rect.y + rect.h / 2 : rect.y + rect.h - 2 - bh / 2;
      ctx.translate(cx, cy); ctx.rotate(-a); ctx.textAlign = 'center'; ctx.fillText(d.text, 0, 0); ctx.textAlign = 'left';
    }
    ctx.restore();
  }
  wrapLines(text, maxW) {
    const out = [];
    for (const para of String(text).split('\n')) {
      let line = '';
      for (const word of para.split(/(\s+)/)) {
        const t = line + word;
        if (this.ctx.measureText(t).width > maxW && line.trim()) { out.push(line.trimEnd()); line = word.trimStart(); } else line = t;
      }
      out.push(line);
    }
    return out;
  }

  drawSelection(W, H) {
    const ctx = this.ctx, g = this.range;
    const a = this.cellRect(g.r1, g.c1), b = this.cellRect(g.r2, g.c2);
    const x = a.x, y = a.y, w = b.x + b.w - a.x, hh = b.y + b.h - a.y;
    ctx.save(); ctx.beginPath(); ctx.rect(HDR_W, HDR_H, W - HDR_W, H - HDR_H); ctx.clip();
    if (g.r1 !== g.r2 || g.c1 !== g.c2) { ctx.fillStyle = 'rgba(14,156,116,.10)'; ctx.fillRect(x, y, w, hh); const ac = this.cellRect(this.sel.r, this.sel.c); ctx.fillStyle = '#fff'; ctx.fillRect(ac.x + 1, ac.y + 1, ac.w - 1, ac.h - 1); const d = this.wb.display(this.sel.r, this.sel.c); if (d.text && !this.editing) this.drawText(d, this.s.style(this.sel.r, this.sel.c), ac, this.sel.r, this.sel.c, this.sel.c); }
    ctx.strokeStyle = this.accent; ctx.lineWidth = 2;
    ctx.strokeRect(Math.round(x) + 0.5, Math.round(y) + 0.5, w - 1, hh - 1);
    // flecha de lista de validación
    const vl = this.validationAt?.(this.sel.r, this.sel.c);
    if (vl?.type === 'list' && !this.editing) { const ac = this.cellRect(this.sel.r, this.sel.c); const bx = ac.x + ac.w + 2, by = ac.y + (ac.h - 18) / 2; ctx.fillStyle = '#fff'; ctx.strokeStyle = '#9AA0B4'; ctx.fillRect(bx, by, 18, 18); ctx.strokeRect(bx + 0.5, by + 0.5, 17, 17); ctx.fillStyle = '#4B5066'; ctx.beginPath(); ctx.moveTo(bx + 5, by + 7); ctx.lineTo(bx + 13, by + 7); ctx.lineTo(bx + 9, by + 12); ctx.fill(); this.listBtn = { x: bx, y: by, w: 18, h: 18 }; } else this.listBtn = null;
    // controlador de relleno
    if (!this.editing) { ctx.fillStyle = this.accent; ctx.fillRect(x + w - 4, y + hh - 4, 7, 7); ctx.fillStyle = '#fff'; ctx.fillRect(x + w - 4, y + hh - 4, 7, 1); ctx.fillRect(x + w - 4, y + hh - 4, 1, 7); }
    // vista previa de relleno
    if (this.fillDrag?.target) { const t = this.fillDrag.target; const ta = this.cellRect(t.r1, t.c1), tb = this.cellRect(t.r2, t.c2); ctx.setLineDash([4, 3]); ctx.strokeStyle = '#555'; ctx.lineWidth = 1; ctx.strokeRect(ta.x + 0.5, ta.y + 0.5, tb.x + tb.w - ta.x - 1, tb.y + tb.h - ta.y - 1); ctx.setLineDash([]); }
    // marco animado de copiar
    if (this.clip && this.clip.sheet === this.s) {
      const cg = this.clip.range; const ca = this.cellRect(cg.r1, cg.c1), cb = this.cellRect(cg.r2, cg.c2);
      ctx.setLineDash([5, 4]); ctx.lineDashOffset = -(performance.now() / 60) % 9; ctx.strokeStyle = this.accent; ctx.lineWidth = 2;
      ctx.strokeRect(ca.x + 1, ca.y + 1, cb.x + cb.w - ca.x - 2, cb.y + cb.h - ca.y - 2); ctx.setLineDash([]);
      this.requestDrawSoon();
    }
    // referencias al editar una fórmula
    if (this.pickRange) { const pr = this.pickRange; const pa = this.cellRect(pr.r1, pr.c1), pb = this.cellRect(pr.r2, pr.c2); ctx.setLineDash([3, 3]); ctx.strokeStyle = '#1565C0'; ctx.lineWidth = 2; ctx.strokeRect(pa.x + 1, pa.y + 1, pb.x + pb.w - pa.x - 2, pb.y + pb.h - pa.y - 2); ctx.setLineDash([]); }
    ctx.restore();
  }
  requestDrawSoon() { clearTimeout(this._ant); this._ant = setTimeout(() => this.requestDraw(), 90); }

  drawHeaders(W, H) {
    const ctx = this.ctx, s = this.s, p = this.prefix(), { fr, fc, FW, FH } = this.frozen();
    const g = this.range;
    ctx.fillStyle = HDR_BG; ctx.fillRect(0, 0, W, HDR_H); ctx.fillRect(0, 0, HDR_W, H);
    ctx.font = `12px ${FONT_DEF}`; ctx.textBaseline = 'middle'; ctx.textAlign = 'center';
    const colsVisible = [];
    for (let c = 0; c < fc; c++) colsVisible.push(c);
    for (let c = Math.max(fc, this.bsearch(p.cx, p.cx[fc] + this.scrollX)); c < s.nCols && this.colLeft(c) < W; c++) colsVisible.push(c);
    for (const c of colsVisible) {
      const x = this.colLeft(c), w = s.cw(c); if (!w) continue;
      const on = c >= g.c1 && c <= g.c2;
      if (on) { ctx.fillStyle = HDR_SEL; ctx.fillRect(x, 0, w, HDR_H); ctx.fillStyle = this.accent; ctx.fillRect(x, HDR_H - 2, w, 2); }
      ctx.fillStyle = on ? '#0A6B50' : HDR_TXT; ctx.fillText(colName(c), x + w / 2, HDR_H / 2 + 1);
      ctx.fillStyle = GRID; ctx.fillRect(Math.round(x + w), 0, 1, HDR_H);
    }
    const rowsVisible = [];
    for (let r = 0; r < fr; r++) rowsVisible.push(r);
    for (let r = Math.max(fr, this.bsearch(p.ry, p.ry[fr] + this.scrollY)); r < s.nRows && this.rowTop(r) < H; r++) rowsVisible.push(r);
    for (const r of rowsVisible) {
      const y = this.rowTop(r), hh = s.rh(r); if (!hh) continue;
      const on = r >= g.r1 && r <= g.r2;
      if (on) { ctx.fillStyle = HDR_SEL; ctx.fillRect(0, y, HDR_W, hh); ctx.fillStyle = this.accent; ctx.fillRect(HDR_W - 2, y, 2, hh); }
      ctx.fillStyle = on ? '#0A6B50' : (s.hiddenRows.has(r + 1) ? '#1565C0' : HDR_TXT); ctx.fillText(String(r + 1), HDR_W / 2, y + hh / 2 + 1);
      ctx.fillStyle = GRID; ctx.fillRect(0, Math.round(y + hh), HDR_W, 1);
    }
    ctx.fillStyle = '#E9EBF0'; ctx.fillRect(0, 0, HDR_W, HDR_H);
    ctx.fillStyle = '#C3C7D3'; ctx.beginPath(); ctx.moveTo(HDR_W - 4, HDR_H - 4); ctx.lineTo(HDR_W - 4, 8); ctx.lineTo(HDR_W - 16, HDR_H - 4); ctx.fill();
    ctx.fillStyle = '#CDD0DA'; ctx.fillRect(0, HDR_H - 1, W, 1); ctx.fillRect(HDR_W - 1, 0, 1, H);
    ctx.textAlign = 'start';
    this.headerOverlay?.(ctx, W, H);
  }

  // ------------------------------------------------------------------ selección
  select(r, c, extend = false, { scroll = true } = {}) {
    const s = this.s;
    r = Math.max(0, Math.min(s.nRows - 1, r)); c = Math.max(0, Math.min(s.nCols - 1, c));
    if (extend) { this.range = { r1: Math.min(this.anchor.r, r), c1: Math.min(this.anchor.c, c), r2: Math.max(this.anchor.r, r), c2: Math.max(this.anchor.c, c) }; this.cursorEnd = { r, c }; }
    else { this.sel = { r, c }; this.anchor = { r, c }; this.cursorEnd = { r, c }; const m = s.mergeAt(r, c); this.range = m ? { ...m } : { r1: r, c1: c, r2: r, c2: c }; }
    // expandir a combinadas
    for (const m of s.merges) if (m.r1 <= this.range.r2 && m.r2 >= this.range.r1 && m.c1 <= this.range.c2 && m.c2 >= this.range.c1) { this.range.r1 = Math.min(this.range.r1, m.r1); this.range.c1 = Math.min(this.range.c1, m.c1); this.range.r2 = Math.max(this.range.r2, m.r2); this.range.c2 = Math.max(this.range.c2, m.c2); }
    if (scroll) this.scrollTo(extend ? this.cursorEnd.r : r, extend ? this.cursorEnd.c : c);
    this.draw();
    this.onSelect?.();
  }
  selectRange(g) { this.sel = { r: g.r1, c: g.c1 }; this.anchor = { r: g.r1, c: g.c1 }; this.range = { ...g }; this.scrollTo(g.r1, g.c1); this.draw(); this.onSelect?.(); }
  scrollTo(r, c) {
    const p = this.prefix(), { fr, fc, FW, FH } = this.frozen();
    const viewW = this.W / this.zoom - HDR_W - FW, viewH = this.H / this.zoom - HDR_H - FH;
    if (c >= fc) {
      const left = p.cx[c] - p.cx[fc], right = p.cx[c + 1] - p.cx[fc];
      if (left < this.scrollX) this.scrollX = left; else if (right > this.scrollX + viewW) this.scrollX = Math.max(0, right - viewW);
    }
    if (r >= fr) {
      const top = p.ry[r] - p.ry[fr], bot = p.ry[r + 1] - p.ry[fr];
      if (top < this.scrollY) this.scrollY = top; else if (bot > this.scrollY + viewH) this.scrollY = Math.max(0, bot - viewH);
    }
    // crecer la hoja al llegar al final
    if (r > this.s.nRows - 20) { this.s.nRows += 500; this.invalidate(); this.updateSpacer(); }
    if (c > this.s.nCols - 5) { this.s.nCols += 20; this.invalidate(); this.updateSpacer(); }
    this.syncScroll();
  }
  syncScroll() { this._syncing = true; this.scroller.scrollLeft = this.scrollX * this.zoom; this.scroller.scrollTop = this.scrollY * this.zoom; requestAnimationFrame(() => { this._syncing = false; }); }

  // Salto con Ctrl+flecha
  jump(dr, dc) {
    const s = this.s, wb = this.wb;
    let { r, c } = this.cursorEnd || this.sel;
    const filled = (rr, cc) => rr >= 0 && cc >= 0 && rr < s.nRows && cc < s.nCols && wb.raw(rr, cc) != null;
    const dims = wb.hf.getSheetDimensions(s.id);
    const limR = Math.max(dims.height, 1), limC = Math.max(dims.width, 1);
    const cur = filled(r, c), next = filled(r + dr, c + dc);
    if (cur && next) { while (filled(r + dr, c + dc)) { r += dr; c += dc; } }
    else { r += dr; c += dc; while (r >= 0 && c >= 0 && r < limR && c < limC && !filled(r, c)) { r += dr; c += dc; } if (r >= limR || c >= limC) { r = dr ? (dr > 0 ? limR - 1 : 0) : r; c = dc ? (dc > 0 ? limC - 1 : 0) : c; } }
    return { r: Math.max(0, r), c: Math.max(0, c) };
  }

  // ------------------------------------------------------------------ edición
  startEdit(initial = null, { caretEnd = true } = {}) {
    if (this.editing) return;
    const { r, c } = this.sel;
    if (this.canEdit && !this.canEdit(r, c)) { toast('La hoja está protegida. Desprotégela en Revisar → Proteger hoja.', { error: true }); return; }
    this.editing = { r, c, typed: initial != null };
    const ed = this.editor;
    ed.value = initial != null ? initial : this.wb.editText(r, c);
    ed.hidden = false;
    this.positionEditor();
    ed.focus();
    if (caretEnd) ed.setSelectionRange(ed.value.length, ed.value.length);
    this.onEdit?.(ed.value);
    this.draw();
  }
  positionEditor() {
    const { r, c } = this.editing; const rc = this.cellRect(r, c); const st = this.s.style(r, c);
    const ed = this.editor, z = this.zoom;
    ed.style.left = `${rc.x * z}px`; ed.style.top = `${rc.y * z}px`;
    ed.style.minWidth = `${rc.w * z + 1}px`; ed.style.height = `${Math.max(rc.h, 22) * z + 1}px`;
    ed.style.font = this.font(st); ed.style.fontSize = `${((st?.size || 11) * 4 / 3) * z}px`;
    ed.style.textAlign = st?.align || 'left';
    ed.style.width = `${Math.max(rc.w * z + 1, Math.min(600, this.measure(ed.value, st) * z + 24))}px`;
  }
  measure(text, st) { this.measureCtx.font = this.font(st); return Math.max(...String(text).split('\n').map((l) => this.measureCtx.measureText(l).width)); }
  commitEdit(move = [1, 0]) {
    if (!this.editing) return;
    const { r, c } = this.editing;
    let v = this.editor.value;
    // Autocierra paréntesis de fórmulas
    if (v.startsWith('=')) { const open = (v.match(/\(/g) || []).length - (v.match(/\)/g) || []).length; if (open > 0) v += ')'.repeat(open); }
    const chk = this.validate ? this.validate(r, c, v) : { ok: true };
    if (!chk.ok) { toast(chk.msg, { error: true }); this.editor.focus(); return; }
    this.editing = null; this.pickRange = null;
    this.editor.hidden = true;
    const before = this.wb.editText(r, c);
    if (v !== before) {
      try { this.wb.action('Escribir', () => this.wb.setInput(r, c, v)); }
      catch (err) { toast(`Fórmula no válida: ${err.message}`, { error: true }); }
      this.onChange?.();
    }
    this.canvas.focus();
    if (move) this.select(r + move[0], c + move[1]);
    else this.draw();
  }
  cancelEdit() { if (!this.editing) return; this.editing = null; this.pickRange = null; this.editor.hidden = true; this.canvas.focus(); this.draw(); this.onSelect?.(); }

  // ¿El cursor del editor está en un punto donde cabe una referencia? (=, (, coma, operador)
  canPick() {
    if (!this.editing) return false;
    const v = this.editor.value; if (!v.startsWith('=')) return false;
    const pos = this.editor.selectionStart;
    const before = v.slice(0, pos).replace(/\$?[A-Z]{1,3}\$?\d+(:\$?[A-Z]{1,3}\$?\d+)?$/i, '');
    return /[=(,;+\-*/^&<>:]\s*$/.test(before) || this._picking;
  }
  insertRef(g) {
    const ed = this.editor, v = ed.value, pos = ed.selectionStart;
    const before = v.slice(0, pos).replace(/\$?[A-Z]{1,3}\$?\d+(:\$?[A-Z]{1,3}\$?\d+)?$/i, '');
    const ref = rangeStr(g);
    ed.value = before + ref + v.slice(pos);
    const np = before.length + ref.length;
    ed.setSelectionRange(np, np);
    this.pickRange = g;
    this.positionEditor(); this.onEdit?.(ed.value); this.draw();
  }

  // ------------------------------------------------------------------ portapapeles
  rangeText(g) {
    const rows = [];
    for (let r = g.r1; r <= g.r2; r++) { if (this.s.hiddenRows.has(r)) continue; const row = []; for (let c = g.c1; c <= g.c2; c++) row.push(this.wb.display(r, c).text); rows.push(row.join('\t')); }
    return rows.join('\n');
  }
  rangeHtml(g) {
    let out = '<table>';
    for (let r = g.r1; r <= g.r2; r++) { out += '<tr>'; for (let c = g.c1; c <= g.c2; c++) { const st = this.s.style(r, c) || {}; const css = [st.b && 'font-weight:bold', st.i && 'font-style:italic', st.color && `color:${st.color}`, st.bg && `background:${st.bg}`, st.align && `text-align:${st.align}`].filter(Boolean).join(';'); out += `<td style="${css}">${esc(this.wb.display(r, c).text)}</td>`; } out += '</tr>'; }
    return `${out}</table>`;
  }
  copy(e, cut = false) {
    const g = { ...this.range };
    const text = this.rangeText(g);
    this.clip = { sheet: this.s, range: g, cut, text };
    if (e) { e.clipboardData.setData('text/plain', text); e.clipboardData.setData('text/html', this.rangeHtml(g)); e.preventDefault(); }
    this.draw();
  }
  paste(text, html) {
    const wb = this.wb, s = this.s;
    const dst = { r: this.range.r1, c: this.range.c1 };
    if (this.clip && text != null && text.replace(/\r/g, '') === this.clip.text) {
      // Pegado interno: fórmulas con referencias ajustadas y formato
      const src = this.clip.range, srcSheet = this.clip.sheet;
      const H = src.r2 - src.r1 + 1, W = src.c2 - src.c1 + 1;
      // Si se seleccionó un área múltiplo del origen, se repite (como en otras hojas de cálculo)
      const reps = { r: Math.max(1, Math.floor((this.range.r2 - this.range.r1 + 1) / H)), c: Math.max(1, Math.floor((this.range.c2 - this.range.c1 + 1) / W)) };
      const styles = []; for (let r = 0; r < H; r++) { styles.push([]); for (let c = 0; c < W; c++) styles[r].push(srcSheet.style(src.r1 + r, src.c1 + c)); }
      const cut = this.clip.cut;
      wb.action(cut ? 'Cortar y pegar' : 'Pegar', () => {
        if (cut) {
          wb.hf.moveCells({ start: { sheet: srcSheet.id, row: src.r1, col: src.c1 }, end: { sheet: srcSheet.id, row: src.r2, col: src.c2 } }, { sheet: s.id, row: dst.r, col: dst.c });
          for (let r = 0; r < H; r++) for (let c = 0; c < W; c++) srcSheet.styles.delete(`${src.r1 + r},${src.c1 + c}`);
        } else {
          wb.hf.copy({ start: { sheet: srcSheet.id, row: src.r1, col: src.c1 }, end: { sheet: srcSheet.id, row: src.r2, col: src.c2 } });
          for (let i = 0; i < reps.r; i++) for (let j = 0; j < reps.c; j++) wb.hf.paste({ sheet: s.id, row: dst.r + i * H, col: dst.c + j * W });
        }
        for (let i = 0; i < reps.r; i++) for (let j = 0; j < reps.c; j++) for (let r = 0; r < H; r++) for (let c = 0; c < W; c++) {
          const st = styles[r][c]; const k = `${dst.r + i * H + r},${dst.c + j * W + c}`;
          if (st) s.styles.set(k, st); else s.styles.delete(k);
        }
        wb.grow(dst.r + H * reps.r, dst.c + W * reps.c);
      }, { sheets: srcSheet === s ? [s] : [s, srcSheet] });
      if (cut) this.clip = null;
      this.selectRange({ r1: dst.r, c1: dst.c, r2: dst.r + H * reps.r - 1, c2: dst.c + W * reps.c - 1 });
      this.onChange?.();
      return;
    }
    if (text == null) return;
    let rows = text.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map((l) => l.split('\t'));
    if (html && /<table/i.test(html) && rows.length <= 1) {
      const d = new DOMParser().parseFromString(html, 'text/html');
      const tr = [...d.querySelectorAll('tr')];
      if (tr.length) rows = tr.map((t) => [...t.querySelectorAll('td,th')].map((c) => c.textContent.trim()));
    }
    wb.action('Pegar', () => {
      rows.forEach((row, i) => row.forEach((v, j) => wb.setInput(dst.r + i, dst.c + j, v)));
    }, { batch: true });
    this.selectRange({ r1: dst.r, c1: dst.c, r2: dst.r + rows.length - 1, c2: dst.c + Math.max(...rows.map((r) => r.length)) - 1 });
    this.onChange?.();
  }

  // ------------------------------------------------------------------ relleno (arrastrar el cuadrito)
  fill(src, target) {
    const wb = this.wb, s = this.s;
    const vertical = target.c1 === src.c1 && target.c2 === src.c2;
    const forward = vertical ? target.r2 > src.r2 : target.c2 > src.c2;
    wb.action('Rellenar', () => {
      const lines = vertical ? range(src.c1, src.c2) : range(src.r1, src.r2);
      for (const ln of lines) {
        const srcCells = vertical ? range(src.r1, src.r2).map((r) => [r, ln]) : range(src.c1, src.c2).map((c) => [ln, c]);
        const dstCells = vertical
          ? (forward ? range(src.r2 + 1, target.r2) : range(target.r1, src.r1 - 1).reverse()).map((r) => [r, ln])
          : (forward ? range(src.c2 + 1, target.c2) : range(target.c1, src.c1 - 1).reverse()).map((c) => [ln, c]);
        const seq = forward ? srcCells : [...srcCells].reverse();
        const raws = seq.map(([r, c]) => wb.raw(r, c));
        const series = detectSeries(raws, !!this.fillCtrl);
        dstCells.forEach(([r, c], i) => {
          const k = i % seq.length;
          const [sr, sc] = seq[k];
          const raw = raws[k];
          if (typeof raw === 'string' && raw.startsWith('=')) {
            wb.hf.copy({ start: { sheet: s.id, row: sr, col: sc }, end: { sheet: s.id, row: sr, col: sc } });
            wb.hf.paste({ sheet: s.id, row: r, col: c });
          } else wb.setRaw(r, c, series ? series(i) : raw);
          const st = s.style(sr, sc);
          if (st) s.styles.set(`${r},${c}`, st); else s.styles.delete(`${r},${c}`);
        });
      }
    });
    this.selectRange({ r1: Math.min(src.r1, target.r1), c1: Math.min(src.c1, target.c1), r2: Math.max(src.r2, target.r2), c2: Math.max(src.c2, target.c2) });
    this.onChange?.();
  }

  // ------------------------------------------------------------------ gráficos (posición)
  positionCharts() {
    const z = this.zoom;
    const p = this.prefix(), { fr, fc, FW, FH } = this.frozen();
    for (const el of this.chartLayer.children) {
      const ch = this.s.charts.find((x) => x.id === el.dataset.id);
      if (!ch) continue;
      const x = HDR_W + FW + ch.x - p.cx[fc] - this.scrollX, y = HDR_H + FH + ch.y - p.ry[fr] - this.scrollY;
      el.style.transform = `translate(${x * z}px, ${y * z}px) scale(${z})`;
    }
    this.chartLayer.style.clipPath = `inset(${(HDR_H + FH) * z}px 0 0 ${(HDR_W + FW) * z}px)`;
  }

  // ------------------------------------------------------------------ eventos
  bind() {
    const cv = this.canvas;
    this.scroller.addEventListener('scroll', () => {
      if (this._syncing) return;
      this.scrollX = this.scroller.scrollLeft / this.zoom; this.scrollY = this.scroller.scrollTop / this.zoom;
      const p = this.prefix(), { fr } = this.frozen();
      if (this.scrollY + this.H / this.zoom > p.ry[this.s.nRows] - p.ry[fr] - 100) { this.s.nRows += 500; this.invalidate(); this.updateSpacer(); }
      this.requestDraw();
    });
    cv.addEventListener('wheel', (e) => {
      if (e.ctrlKey) { e.preventDefault(); this.onZoomWheel?.(e.deltaY < 0 ? 10 : -10); return; }
      e.preventDefault();
      this.scroller.scrollTop += e.shiftKey ? 0 : e.deltaY;
      this.scroller.scrollLeft += e.shiftKey ? e.deltaY : e.deltaX;
    }, { passive: false });
    cv.addEventListener('mousedown', (e) => this.onMouseDown(e));
    cv.addEventListener('mousemove', (e) => this.onHover(e));
    cv.addEventListener('dblclick', (e) => this.onDblClick(e));
    cv.addEventListener('contextmenu', (e) => { e.preventDefault(); const { r, c } = this.hit(e.offsetX, e.offsetY); if (r >= 0 && c >= 0 && !(r >= this.range.r1 && r <= this.range.r2 && c >= this.range.c1 && c <= this.range.c2)) this.select(r, c); this.onContext?.(e, r < 0 ? 'col' : c < 0 ? 'row' : 'cell'); });
    cv.addEventListener('keydown', (e) => this.onKey(e));
    // copiar/pegar con el portapapeles del sistema
    document.addEventListener('copy', (e) => { if (document.activeElement === cv) this.copy(e); });
    document.addEventListener('cut', (e) => { if (document.activeElement === cv) this.copy(e, true); });
    document.addEventListener('paste', (e) => { if (document.activeElement === cv) { e.preventDefault(); this.paste(e.clipboardData.getData('text/plain'), e.clipboardData.getData('text/html')); } });
    // editor de celda
    const ed = this.editor;
    ed.addEventListener('input', () => { this._picking = false; this.positionEditor(); this.onEdit?.(ed.value); });
    ed.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); this.commitEdit(e.shiftKey ? [-1, 0] : [1, 0]); }
      else if (e.key === 'Enter' && e.altKey) { e.preventDefault(); const p = ed.selectionStart; ed.value = `${ed.value.slice(0, p)}\n${ed.value.slice(ed.selectionEnd)}`; ed.setSelectionRange(p + 1, p + 1); }
      else if (e.key === 'Tab') { e.preventDefault(); this.commitEdit(e.shiftKey ? [0, -1] : [0, 1]); }
      else if (e.key === 'Escape') { e.preventDefault(); this.cancelEdit(); }
      else if (this.editing?.typed && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
        if (this.canPick()) {
          // Flechas mientras se escribe una fórmula: eligen una referencia
          e.preventDefault();
          const base = this.pickRange ? { r: this.pickRange.r1, c: this.pickRange.c1 } : { ...this.editing };
          const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
          const r = Math.max(0, base.r + d[0]), c = Math.max(0, base.c + d[1]);
          this._picking = true; this.insertRef({ r1: r, c1: c, r2: r, c2: c });
        } else { e.preventDefault(); const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key]; this.commitEdit(d); }
      } else if (e.key === 'F4') {
        // Alterna $ en la referencia bajo el cursor
        e.preventDefault();
        const v = ed.value, pos = ed.selectionStart;
        const m = /(\$?)([A-Z]{1,3})(\$?)(\d+)$/i.exec(v.slice(0, pos));
        if (m) {
          const states = [['', ''], ['$', '$'], ['', '$'], ['$', '']];
          const i = states.findIndex(([a, b]) => a === m[1] && b === m[3]);
          const [a, b] = states[(i + 1) % 4];
          const rep = `${a}${m[2]}${b}${m[4]}`;
          ed.value = v.slice(0, pos - m[0].length) + rep + v.slice(pos);
          const np = pos - m[0].length + rep.length; ed.setSelectionRange(np, np); this.onEdit?.(ed.value);
        }
      }
    });
  }

  onHover(e) {
    if (this.drag) return;
    const x = e.offsetX / this.zoom, y = e.offsetY / this.zoom;
    this.noteHover(e);
    const rz = this.resizeZone(x, y);
    if (rz) { this.canvas.style.cursor = rz.axis === 'col' ? 'col-resize' : 'row-resize'; return; }
    const g = this.range, a = this.cellRect(g.r1, g.c1), b = this.cellRect(g.r2, g.c2);
    const hx = b.x + b.w, hy = b.y + b.h;
    if (Math.abs(x - hx) < 6 && Math.abs(y - hy) < 6) { this.canvas.style.cursor = 'crosshair'; return; }
    this.canvas.style.cursor = y < HDR_H && x > HDR_W ? 's-resize' : x < HDR_W && y > HDR_H ? 'e-resize' : 'cell';
    void a;
  }
  noteHover(e) {
    const s = this.s;
    if (!s.notes?.size) { this.hideNote(); return; }
    const { r, c } = this.hit(e.offsetX, e.offsetY);
    const txt = r >= 0 && c >= 0 ? s.notes.get(`${r},${c}`) : null;
    if (!txt) { this.hideNote(); return; }
    if (!this.noteTip) { this.noteTip = h('div', { class: 'note-tip' }); this.host.append(this.noteTip); }
    const rc = this.cellRect(r, c);
    this.noteTip.textContent = txt; this.noteTip.hidden = false;
    this.noteTip.style.left = `${(rc.x + rc.w + 6) * this.zoom}px`; this.noteTip.style.top = `${rc.y * this.zoom}px`;
  }
  hideNote() { if (this.noteTip) this.noteTip.hidden = true; }
  resizeZone(x, y) {
    const s = this.s;
    if (y < HDR_H && x > HDR_W) { const { c } = this.hit(x * this.zoom, (HDR_H + 1) * this.zoom); for (const cc of [c - 1, c]) { if (cc < 0) continue; const edge = this.colLeft(cc) + s.cw(cc); if (Math.abs(x - edge) < 4) return { axis: 'col', i: cc, start: edge }; } }
    if (x < HDR_W && y > HDR_H) { const { r } = this.hit((HDR_W + 1) * this.zoom, y * this.zoom); for (const rr of [r - 1, r]) { if (rr < 0) continue; const edge = this.rowTop(rr) + s.rh(rr); if (Math.abs(y - edge) < 4) return { axis: 'row', i: rr, start: edge }; } }
    return null;
  }

  onMouseDown(e) {
    if (e.button !== 0) return;
    closeDropdown();
    const x = e.offsetX / this.zoom, y = e.offsetY / this.zoom;
    if (this.overlayClick?.(x, y, e)) { e.preventDefault(); return; }
    if (this.listBtn && x >= this.listBtn.x && x <= this.listBtn.x + this.listBtn.w && y >= this.listBtn.y && y <= this.listBtn.y + this.listBtn.h) { e.preventDefault(); this.onListButton?.(this.sel.r, this.sel.c, { left: e.clientX, top: e.clientY }); return; }
    const s = this.s;
    // ¿filtro?
    if (s.filter) {
      const f = s.filter;
      for (let c = f.range.c1; c <= f.range.c2; c++) {
        const rc = this.cellRect(f.range.r1, c);
        if (x >= rc.x + rc.w - 18 && x <= rc.x + rc.w - 2 && y >= rc.y + (rc.h - 16) / 2 && y <= rc.y + (rc.h + 16) / 2) { e.preventDefault(); this.onFilterButton?.(c, { left: e.clientX, top: e.clientY }); return; }
      }
    }
    // cambiar tamaño de fila/columna
    const rz = this.resizeZone(x, y);
    if (rz) {
      e.preventDefault();
      const size0 = rz.axis === 'col' ? s.cw(rz.i) : s.rh(rz.i);
      const before = s.meta();
      const p0 = rz.axis === 'col' ? e.clientX : e.clientY;
      const cols = rz.axis === 'col' && this.range.c1 <= rz.i && this.range.c2 >= rz.i && this.range.r1 === 0 && this.range.r2 >= s.nRows - 1 ? range(this.range.c1, this.range.c2) : [rz.i];
      const move = (ev) => {
        const d = ((rz.axis === 'col' ? ev.clientX : ev.clientY) - p0) / this.zoom;
        const v = Math.max(rz.axis === 'col' ? 8 : 10, Math.round(size0 + d));
        if (rz.axis === 'col') cols.forEach((c) => s.colW.set(c, v)); else s.rowH.set(rz.i, v);
        this.touchMeta(); this.invalidate(); this.updateSpacer(); this.draw();
      };
      const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); this.wb.undoStack.push({ label: 'Tamaño', hf: false, before: [[s, before]], after: [[s, s.meta()]] }); this.wb.redoStack = []; this.onChange?.(); };
      addEventListener('mousemove', move); addEventListener('mouseup', up);
      return;
    }
    const { r, c } = this.hit(e.offsetX, e.offsetY);
    // elegir referencia para una fórmula
    if (this.editing && this.canPick() && r >= 0 && c >= 0) {
      e.preventDefault();
      this._picking = true;
      const start = { r, c };
      this.insertRef({ r1: r, c1: c, r2: r, c2: c });
      const move = (ev) => { const p = this.hit(ev.clientX - this.canvas.getBoundingClientRect().left, ev.clientY - this.canvas.getBoundingClientRect().top); if (p.r < 0 || p.c < 0) return; this.insertRef({ r1: Math.min(start.r, p.r), c1: Math.min(start.c, p.c), r2: Math.max(start.r, p.r), c2: Math.max(start.c, p.c) }); };
      const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); this.editor.focus(); };
      addEventListener('mousemove', move); addEventListener('mouseup', up);
      return;
    }
    if (this.editing) this.commitEdit(null);
    this.canvas.focus();
    // controlador de relleno
    const g = this.range, b = this.cellRect(g.r2, g.c2);
    if (Math.abs(x - (b.x + b.w)) < 6 && Math.abs(y - (b.y + b.h)) < 6) {
      e.preventDefault();
      this.fillDrag = { src: { ...g }, target: null }; this.fillCtrl = e.ctrlKey || e.altKey;
      const move = (ev) => {
        const rect = this.canvas.getBoundingClientRect();
        const p = this.hit(ev.clientX - rect.left, ev.clientY - rect.top);
        const src = this.fillDrag.src;
        let t = null;
        const dr = p.r > src.r2 ? p.r - src.r2 : p.r < src.r1 ? p.r - src.r1 : 0;
        const dc = p.c > src.c2 ? p.c - src.c2 : p.c < src.c1 ? p.c - src.c1 : 0;
        if (Math.abs(dr) >= Math.abs(dc) && dr) t = dr > 0 ? { ...src, r2: p.r } : { ...src, r1: Math.max(0, p.r) };
        else if (dc) t = dc > 0 ? { ...src, c2: p.c } : { ...src, c1: Math.max(0, p.c) };
        this.fillDrag.target = t;
        this.autoScroll(ev);
        this.draw();
      };
      const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); this.stopAutoScroll(); const fd = this.fillDrag; this.fillDrag = null; if (fd.target) this.fill(fd.src, fd.target); else this.draw(); };
      addEventListener('mousemove', move); addEventListener('mouseup', up);
      return;
    }
    // encabezados: seleccionar fila/columna completa
    if (r < 0 && c < 0) { this.selectRange({ r1: 0, c1: 0, r2: s.nRows - 1, c2: s.nCols - 1 }); return; }
    if (r < 0) {
      if (e.shiftKey) this.range = { r1: 0, r2: s.nRows - 1, c1: Math.min(this.anchor.c, c), c2: Math.max(this.anchor.c, c) };
      else { this.sel = { r: this.rowTopVisible(), c }; this.anchor = { r: 0, c }; this.range = { r1: 0, c1: c, r2: s.nRows - 1, c2: c }; this.sel.r = 0; }
      this.draw(); this.onSelect?.();
      this.dragHeader('col', e);
      return;
    }
    if (c < 0) {
      if (e.shiftKey) this.range = { c1: 0, c2: s.nCols - 1, r1: Math.min(this.anchor.r, r), r2: Math.max(this.anchor.r, r) };
      else { this.sel = { r, c: 0 }; this.anchor = { r, c: 0 }; this.range = { r1: r, c1: 0, r2: r, c2: s.nCols - 1 }; }
      this.draw(); this.onSelect?.();
      this.dragHeader('row', e);
      return;
    }
    this.select(r, c, e.shiftKey, { scroll: false });
    this.drag = true;
    const move = (ev) => {
      const rect = this.canvas.getBoundingClientRect();
      const p = this.hit(ev.clientX - rect.left, ev.clientY - rect.top);
      if (p.r < 0 || p.c < 0) { this.autoScroll(ev); return; }
      this.select(p.r, p.c, true, { scroll: false });
      this.autoScroll(ev);
    };
    const up = () => { this.drag = false; this.stopAutoScroll(); removeEventListener('mousemove', move); removeEventListener('mouseup', up); };
    addEventListener('mousemove', move); addEventListener('mouseup', up);
  }
  rowTopVisible() { return 0; }
  dragHeader(axis, e0) {
    const s = this.s;
    const move = (ev) => {
      const rect = this.canvas.getBoundingClientRect();
      const p = this.hit(ev.clientX - rect.left, ev.clientY - rect.top);
      if (axis === 'col' && p.c >= 0) this.range = { r1: 0, r2: s.nRows - 1, c1: Math.min(this.anchor.c, p.c), c2: Math.max(this.anchor.c, p.c) };
      if (axis === 'row' && p.r >= 0) this.range = { c1: 0, c2: s.nCols - 1, r1: Math.min(this.anchor.r, p.r), r2: Math.max(this.anchor.r, p.r) };
      this.draw(); this.onSelect?.();
    };
    const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); };
    addEventListener('mousemove', move); addEventListener('mouseup', up);
    void e0;
  }
  autoScroll(ev) {
    const rect = this.canvas.getBoundingClientRect();
    const dx = ev.clientX < rect.left + HDR_W ? -20 : ev.clientX > rect.right - 10 ? 20 : 0;
    const dy = ev.clientY < rect.top + HDR_H ? -20 : ev.clientY > rect.bottom - 10 ? 20 : 0;
    this.stopAutoScroll();
    if (dx || dy) this._as = setInterval(() => { this.scroller.scrollLeft += dx; this.scroller.scrollTop += dy; }, 30);
  }
  stopAutoScroll() { clearInterval(this._as); }

  onDblClick(e) {
    const x = e.offsetX / this.zoom, y = e.offsetY / this.zoom;
    const rz = this.resizeZone(x, y);
    if (rz && rz.axis === 'col') { this.autofit(rz.i); return; }
    if (rz && rz.axis === 'row') { const s = this.s; const before = s.meta(); s.rowH.delete(rz.i); this.wb.undoStack.push({ label: 'Alto', hf: false, before: [[s, before]], after: [[s, s.meta()]] }); this.touchMeta(); this.invalidate(); this.updateSpacer(); this.draw(); return; }
    const { r, c } = this.hit(e.offsetX, e.offsetY);
    if (r >= 0 && c >= 0) { this.select(r, c); this.startEdit(); }
  }
  autofit(col) {
    const s = this.s, wb = this.wb;
    const cols = this.range.c1 <= col && this.range.c2 >= col && this.range.r2 >= s.nRows - 1 ? range(this.range.c1, this.range.c2) : [col];
    const before = s.meta();
    const dims = wb.hf.getSheetDimensions(s.id);
    for (const c of cols) {
      let max = 30;
      for (let r = 0; r < Math.min(dims.height, 5000); r++) { const d = wb.display(r, c); if (d.text && !s.mergeAt(r, c)) max = Math.max(max, this.measure(d.text, s.style(r, c)) + 12); }
      s.colW.set(c, Math.ceil(Math.min(max, 600)));
    }
    wb.undoStack.push({ label: 'Autoajustar', hf: false, before: [[s, before]], after: [[s, s.meta()]] }); wb.redoStack = [];
    this.touchMeta(); this.invalidate(); this.updateSpacer(); this.draw(); this.onChange?.();
  }

  onKey(e) {
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key;
    const move = (dr, dc) => {
      e.preventDefault();
      const base = e.shiftKey ? this.cursorEnd || this.sel : this.sel;
      let target;
      if (mod) { this.cursorEnd = base; target = this.jump(dr, dc); }
      else {
        // saltar filas ocultas y combinadas
        let r = base.r + dr, c = base.c + dc;
        while (dr && r > 0 && r < this.s.nRows && this.s.hiddenRows.has(r)) r += dr;
        const m = this.s.mergeAt(base.r, base.c);
        if (m && !e.shiftKey) { if (dr > 0) r = m.r2 + 1; if (dc > 0) c = m.c2 + 1; }
        target = { r, c };
      }
      this.select(target.r, target.c, e.shiftKey);
    };
    if (k === 'ArrowDown') return move(1, 0);
    if (k === 'ArrowUp') return move(-1, 0);
    if (k === 'ArrowRight') return move(0, 1);
    if (k === 'ArrowLeft') return move(0, -1);
    if (k === 'PageDown') { e.preventDefault(); const n = Math.floor(this.H / this.zoom / 24) - 2; this.select(this.sel.r + n, this.sel.c, e.shiftKey); return; }
    if (k === 'PageUp') { e.preventDefault(); const n = Math.floor(this.H / this.zoom / 24) - 2; this.select(this.sel.r - n, this.sel.c, e.shiftKey); return; }
    if (k === 'Home') { e.preventDefault(); this.select(mod ? 0 : this.sel.r, 0, e.shiftKey); return; }
    if (k === 'End' && mod) { e.preventDefault(); const u = this.wb.usedRange(); this.select(u.r2, u.c2, e.shiftKey); return; }
    if (k === 'Tab') { e.preventDefault(); this.select(this.sel.r, this.sel.c + (e.shiftKey ? -1 : 1)); return; }
    if (k === 'Enter') {
      e.preventDefault();
      const g = this.range;
      if (g.r1 !== g.r2 || g.c1 !== g.c2) {
        // moverse dentro de la selección
        let { r, c } = this.sel; r += e.shiftKey ? -1 : 1;
        if (r > g.r2) { r = g.r1; c = c + 1 > g.c2 ? g.c1 : c + 1; } if (r < g.r1) { r = g.r2; c = c - 1 < g.c1 ? g.c2 : c - 1; }
        this.sel = { r, c }; this.draw(); this.onSelect?.(); return;
      }
      this.select(this.sel.r + (e.shiftKey ? -1 : 1), this.sel.c); return;
    }
    if (k === 'F2') { e.preventDefault(); this.startEdit(); return; }
    if (k === 'Escape') { if (this.clip) { this.clip = null; this.draw(); } return; }
    if ((k === 'Delete' || k === 'Backspace') && !mod) { e.preventDefault(); this.onClear?.(); return; }
    if (mod && k.toLowerCase() === 'a') { e.preventDefault(); const reg = this.wb.region(this.sel.r, this.sel.c); const same = reg.r1 === this.range.r1 && reg.r2 === this.range.r2 && reg.c1 === this.range.c1 && reg.c2 === this.range.c2; if (same || (reg.r1 === reg.r2 && reg.c1 === reg.c2)) this.selectRange({ r1: 0, c1: 0, r2: this.s.nRows - 1, c2: this.s.nCols - 1 }); else { this.range = reg; this.draw(); this.onSelect?.(); } return; }
    if (mod && k === ';') { e.preventDefault(); const d = new Date(); this.wb.action('Fecha', () => this.wb.setInput(this.sel.r, this.sel.c, `${d.getDate()}/${d.getMonth() + 1}/${d.getFullYear()}`)); this.draw(); this.onChange?.(); return; }
    if (mod && (k === 'd' || k === 'D')) { e.preventDefault(); const g = this.range; if (g.r2 > g.r1) this.fill({ ...g, r2: g.r1 }, g); return; }
    if (mod && (k === 'r' || k === 'R')) { e.preventDefault(); const g = this.range; if (g.c2 > g.c1) this.fill({ ...g, c2: g.c1 }, g); return; }
    if (mod || e.altKey) return;
    if (k.length === 1) { e.preventDefault(); this.startEdit(k); }
  }
}

function range(a, b) { const out = []; for (let i = a; i <= b; i++) out.push(i); return out; }

// Detecta una serie (1,2,3… / enero, febrero… / "Producto 1"…). Devuelve fn(i) -> valor o null para copiar.
function detectSeries(raws, forceStep) {
  const n = raws.length;
  if (raws.every((v) => typeof v === 'number')) {
    if (n === 1) return forceStep ? (i) => raws[0] + i + 1 : null;
    const step = raws[1] - raws[0];
    for (let i = 2; i < n; i++) if (Math.abs(raws[i] - raws[i - 1] - step) > 1e-9) return null;
    return (i) => raws[n - 1] + step * (i + 1);
  }
  if (raws.every((v) => typeof v === 'string' && !v.startsWith('='))) {
    const lists = [MESES, MESES.map((m) => m.slice(0, 3)), DIAS, DIAS.map((d) => d.slice(0, 3))];
    for (const L of lists) {
      const idx = raws.map((v) => L.indexOf(v.toLowerCase()));
      if (idx.every((x) => x >= 0)) {
        const step = n > 1 ? idx[1] - idx[0] : 1;
        const cap = (s, ref) => (ref[0] === ref[0].toUpperCase() ? (ref === ref.toUpperCase() ? s.toUpperCase() : s[0].toUpperCase() + s.slice(1)) : s);
        return (i) => cap(L[(((idx[n - 1] + step * (i + 1)) % L.length) + L.length) % L.length], raws[0]);
      }
    }
    const m = raws.map((v) => /^(.*?)(\d+)$/.exec(v));
    if (m.every(Boolean) && m.every((x) => x[1] === m[0][1])) {
      const nums = m.map((x) => Number(x[2]));
      const step = n > 1 ? nums[1] - nums[0] : 1;
      return (i) => `${m[0][1]}${nums[n - 1] + step * (i + 1)}`;
    }
  }
  return null;
}
export { HDR_W, HDR_H, parseRange, parseAddr, addr, toSpanish, toEnglish, parseInput, icon, h, dropdown };
