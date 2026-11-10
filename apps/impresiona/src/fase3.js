// Impresiona — fase 3: video y audio con subtítulos, ensayar intervalos, trayectorias de animación
// y exportar la presentación a video (WebM). GPL-3.0
import { Muxer, ArrayBufferTarget } from 'webm-muxer';
import { h, toast, modal, form, api, t, tf, esc } from '../../shared/office.js';
import { buildSteps, pathPoints, PATH_TYPES } from './show.js';

let C = null; // { getPres, slide, commit, select, center, uid, renderAll, present, runShow, renderSlide }
export function init(ctx) { C = ctx; }

// ================================================================ video y audio
const MIME = { mp4: 'video/mp4', m4v: 'video/mp4', webm: 'video/webm', ogv: 'video/ogg', mov: 'video/quicktime', mp3: 'audio/mpeg', m4a: 'audio/mp4', wav: 'audio/wav', ogg: 'audio/ogg', oga: 'audio/ogg', aac: 'audio/aac', flac: 'audio/flac' };
const blobs = new Map();
function dataToBlobUrl(data) {
  if (!data || !data.startsWith('data:')) return data;
  if (blobs.has(data)) return blobs.get(data);
  const [head, b64] = data.split(','); const mime = /data:([^;]+)/.exec(head)?.[1] || 'application/octet-stream';
  const bin = atob(b64); const u8 = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
  const url = URL.createObjectURL(new Blob([u8], { type: mime })); blobs.set(data, url); return url;
}
const vttUrls = new Map();
function vttUrl(text) { if (vttUrls.has(text)) return vttUrls.get(text); const u = URL.createObjectURL(new Blob([text], { type: 'text/vtt' })); vttUrls.set(text, u); return u; }
function bytesToDataUrl(bytes, mime) {
  let s = ''; const CH = 0x8000; for (let i = 0; i < bytes.length; i += CH) s += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
  return `data:${mime};base64,${btoa(s)}`;
}
export async function insertMedia(kind) {
  const exts = kind === 'video' ? ['mp4', 'webm', 'm4v', 'ogv', 'mov'] : ['mp3', 'm4a', 'wav', 'ogg', 'oga', 'aac', 'flac'];
  const f = await api.openDialog({ title: kind === 'video' ? t('Insertar video') : t('Insertar audio'), filters: [{ name: kind === 'video' ? t('Videos') : t('Audio'), extensions: exts }] });
  if (!f || f.redirected || !f.bytes) return;
  if (f.bytes.length > 200 * 1024 * 1024) { toast('El archivo es muy grande (máximo 200 MB).', { error: true }); return; }
  const src = bytesToDataUrl(f.bytes, MIME[f.ext] || `${kind}/${f.ext}`);
  addMedia(kind, src, f.name);
}
export function addMedia(kind, src, name = '') {
  const pres = C.getPres();
  const [w, hh] = kind === 'video' ? [Math.min(640, pres.size.w * 0.66), Math.min(360, pres.size.h * 0.66)] : [96, 96];
  const o = { id: C.uid(), type: 'media', kind, src, name, ...C.center(w, hh), rot: 0, autoplay: kind === 'audio', loop: false, captions: [], style: {} };
  C.commit(kind === 'video' ? t('Insertar video') : t('Insertar audio'), () => C.slide().objects.push(o));
  C.select([o.id]);
  return o;
}
// Dibuja el objeto multimedia dentro de la diapositiva
export function mediaBox(o, mode, box) {
  const url = dataToBlobUrl(o.src);
  if (o.kind === 'audio') {
    box.classList.add('t-audio');
    box.append(h('div', { class: 'audio-ico', html: '<svg viewBox="0 0 24 24" width="60%" height="60%" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 9v6h4l5 4V5L8 9H4z"/><path d="M16 9a4 4 0 010 6M18.5 6.5a8 8 0 010 11"/></svg>' }));
    if (mode === 'show') { const a = h('audio', { src: url, controls: true, preload: 'auto' }); if (o.autoplay) a.dataset.autoplay = '1'; a.loop = !!o.loop; box.append(a); }
    return;
  }
  const v = h('video', { src: url, preload: 'metadata', playsinline: true });
  v.muted = mode !== 'show'; v.loop = !!o.loop;
  if (mode === 'show') { v.controls = true; if (o.autoplay) v.dataset.autoplay = '1'; }
  for (const [k, c] of (o.captions || []).entries()) v.append(h('track', { kind: 'subtitles', src: vttUrl(c.vtt), srclang: c.lang || 'es', label: c.label || c.lang || 'Subtítulos', default: k === 0 ? true : null }));
  v.addEventListener('loadedmetadata', () => { if (mode !== 'show' && v.duration) v.currentTime = Math.min(0.5, v.duration / 3); });
  box.append(v);
  if (mode === 'edit') box.append(h('div', { class: 'media-play', title: t('Reproducir'), html: '▶', onmousedown: (e) => { e.stopPropagation(); }, onclick: (e) => { e.stopPropagation(); if (v.paused) { v.muted = false; v.play(); e.currentTarget.textContent = '❚❚'; } else { v.pause(); e.currentTarget.textContent = '▶'; } } }));
}
// ---- subtítulos
export function srtToVtt(text) {
  const s = String(text || '').replace(/\r/g, '').trim();
  if (/^WEBVTT/.test(s)) return `${s}\n`;
  return `WEBVTT\n\n${s.replace(/^\d+\n(?=\d\d:)/gm, '').replace(/(\d\d:\d\d:\d\d),(\d\d\d)/g, '$1.$2')}\n`;
}
export function parseCues(vtt) {
  return [...String(vtt).matchAll(/(\d\d:)?(\d\d):(\d\d)[.,](\d\d\d)\s+-->\s+(\d\d:)?(\d\d):(\d\d)[.,](\d\d\d)[^\n]*\n([\s\S]*?)(?=\n\n|\n*$)/g)].map((m) => ({ a: (Number((m[1] || '0:').slice(0, -1)) * 3600) + Number(m[2]) * 60 + Number(m[3]) + Number(m[4]) / 1000, b: (Number((m[5] || '0:').slice(0, -1)) * 3600) + Number(m[6]) * 60 + Number(m[7]) + Number(m[8]) / 1000, text: m[9].trim() }));
}
const ts = (x) => { const hh = Math.floor(x / 3600), mm = Math.floor((x % 3600) / 60), ss = Math.floor(x % 60), ms = Math.round((x % 1) * 1000); return `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}:${String(ss).padStart(2, '0')}.${String(ms).padStart(3, '0')}`; };
// Texto simple "0:05 Hola" por línea -> WebVTT
export function simpleToVtt(text, total = 0) {
  const lines = String(text || '').split('\n').map((l) => /^\s*(?:(\d+):)?(\d+):(\d\d(?:\.\d+)?)\s+(.+)$/.exec(l)).filter(Boolean).map((m) => ({ at: (Number(m[1] || 0) * 3600) + Number(m[2]) * 60 + Number(m[3]), text: m[4] }));
  return `WEBVTT\n\n${lines.map((c, i) => `${ts(c.at)} --> ${ts(lines[i + 1]?.at ?? Math.max(c.at + 4, total || c.at + 4))}\n${c.text}`).join('\n\n')}\n`;
}
export async function captionsDialog(o) {
  if (!o || o.type !== 'media') { toast('Selecciona primero un video o un audio.'); return; }
  const cur = (o.captions || [])[0];
  const simple = cur ? parseCues(cur.vtt).map((c) => `${Math.floor(c.a / 60)}:${String(Math.floor(c.a % 60)).padStart(2, '0')} ${c.text}`).join('\n') : '';
  const area = h('textarea', { rows: 8, name: 'txt', placeholder: '0:00 Bienvenidos\n0:04 Hoy veremos…' }, simple);
  const lang = h('select', { name: 'lang' }, ...[['es', 'Español'], ['en', 'English']].map(([v, l]) => h('option', { value: v, selected: (cur?.lang || 'es') === v ? true : null }, l)));
  const body = h('div', {}, h('p', { class: 'muted' }, t('Escribe un subtítulo por línea con el minuto y segundo en que aparece, o carga un archivo .vtt o .srt.')),
    h('label', { class: 'field' }, h('span', {}, t('Idioma')), lang), h('label', { class: 'field' }, h('span', {}, t('Subtítulos')), area),
    h('button', { class: 'btn', type: 'button', onclick: async () => { const f = await api.openDialog({ title: t('Subtítulos'), filters: [{ name: 'WebVTT / SRT', extensions: ['vtt', 'srt'] }] }); if (!f?.bytes) return; const vtt = srtToVtt(new TextDecoder().decode(f.bytes)); area.value = parseCues(vtt).map((c) => `${Math.floor(c.a / 60)}:${String((c.a % 60).toFixed(0)).padStart(2, '0')} ${c.text.replace(/\n/g, ' ')}`).join('\n'); area.dataset.vtt = vtt; } }, t('Cargar .vtt o .srt…')));
  const v = await modal({ title: t('Subtítulos'), body, width: '520px', buttons: [{ label: 'Quitar subtítulos', value: 'remove' }, { label: 'Cancelar', value: null }, { label: 'Guardar', primary: true, value: () => ({ lang: lang.value, txt: area.value, vtt: area.dataset.vtt }) }] });
  if (!v) return;
  C.commit(t('Subtítulos'), () => {
    if (v === 'remove' || !v.txt.trim()) o.captions = [];
    else { const vtt = v.vtt && v.txt.trim() === area.value.trim() ? v.vtt : simpleToVtt(v.txt); o.captions = [{ lang: v.lang, label: v.lang === 'en' ? 'English' : 'Español', vtt }]; }
  });
}
export async function mediaOptions(o) {
  if (!o || o.type !== 'media') { toast('Selecciona primero un video o un audio.'); return; }
  const v = await form(o.kind === 'video' ? 'Opciones del video' : 'Opciones del audio', [
    { name: 'autoplay', label: 'Reproducir automáticamente al llegar a la diapositiva', type: 'checkbox', value: !!o.autoplay },
    { name: 'loop', label: 'Repetir hasta que se detenga', type: 'checkbox', value: !!o.loop },
  ], 'Aceptar');
  if (!v) return;
  C.commit(t('Opciones de reproducción'), () => { o.autoplay = v.autoplay; o.loop = v.loop; });
}

// ================================================================ trayectorias (vista previa en edición)
export function pathPreview(o, box) {
  if (!PATH_TYPES.includes(o.anim?.type)) return;
  const pts = pathPoints(o.anim.type, Number(o.anim.dist) || 200);
  const xs = pts.map((p) => p[0]), ys = pts.map((p) => p[1]);
  const minX = Math.min(0, ...xs), minY = Math.min(0, ...ys), maxX = Math.max(0, ...xs), maxY = Math.max(0, ...ys);
  const cx = o.w / 2, cy = o.h / 2;
  const svg = `<svg class="path-prev" style="left:${minX + cx - 6}px;top:${minY + cy - 6}px" width="${maxX - minX + 12}" height="${maxY - minY + 12}"><polyline points="${pts.map(([x, y]) => `${x - minX + 6},${y - minY + 6}`).join(' ')}" fill="none" stroke="#2BA84A" stroke-width="2" stroke-dasharray="5 4"/><circle cx="${-minX + 6}" cy="${-minY + 6}" r="4" fill="#2BA84A"/><circle cx="${xs[xs.length - 1] - minX + 6}" cy="${ys[ys.length - 1] - minY + 6}" r="4" fill="#D93025"/></svg>`;
  box.insertAdjacentHTML('beforeend', svg);
}
export async function pathDistance(objs) {
  if (!objs.length || !objs.every((o) => PATH_TYPES.includes(o.anim?.type))) { toast('Selecciona un objeto con una trayectoria.'); return; }
  const v = await form('Trayectoria', [{ name: 'dist', label: 'Distancia del recorrido (px)', type: 'number', value: objs[0].anim.dist || 200, min: 20, max: 2000 }], 'Aceptar');
  if (!v) return;
  C.commit(t('Trayectoria'), () => objs.forEach((o) => { o.anim.dist = v.dist; }));
}

// ================================================================ ensayar intervalos
export function rehearse() {
  C.prepareShow();
  C.runShow({ renderSlide: C.renderSlide, getPres: C.getPres, from: 0, rehearse: true, onEnd: (i) => C.goto(i), onRehearsed: async (times, total) => {
    const pres = C.getPres();
    const keep = await modal({ title: t('Ensayar intervalos'), body: h('p', {}, tf('La presentación duró {0}. ¿Quieres guardar los intervalos para que avance sola?', fmt(total))), buttons: [{ label: 'No', value: false }, { label: 'Sí, guardar', primary: true, value: true }] });
    if (!keep) return;
    C.commit(t('Ensayar intervalos'), () => { pres.slides.forEach((s, i) => { if (times[i]) s.advanceAfter = times[i]; }); pres.useTimings = true; });
    toast('Intervalos guardados. La presentación avanzará sola.');
  } });
}
const fmt = (ms) => { const sec = Math.round(ms / 1000); return `${Math.floor(sec / 60)}:${String(sec % 60).padStart(2, '0')}`; };
export function setUseTimings(on) { C.commit(t('Usar intervalos'), () => { C.getPres().useTimings = on; }); }
export function clearTimings() { C.commit(t('Borrar intervalos'), () => { C.getPres().slides.forEach((s) => { delete s.advanceAfter; }); }); }

// ================================================================ exportar a video (WebM)
async function pickCodec(width, height, fps) {
  for (const [codec, mux] of [['vp09.00.10.08', 'V_VP9'], ['vp8', 'V_VP8']]) {
    try { const r = await VideoEncoder.isConfigSupported({ codec, width, height, bitrate: 6e6, framerate: fps }); if (r.supported) return { codec, mux }; } catch { /* */ }
  }
  return null;
}
// Dibuja un cuadro intermedio de una transición entre dos imágenes
function drawTransition(ctx, a, b, type, p, W, H) {
  const e = p < 0.5 ? 2 * p * p : 1 - (-2 * p + 2) ** 2 / 2;
  ctx.globalAlpha = 1;
  if (type === 'push') { ctx.drawImage(a, -W * e, 0, W, H); ctx.drawImage(b, W - W * e, 0, W, H); return; }
  if (type === 'cover') { ctx.drawImage(a, 0, 0, W, H); ctx.drawImage(b, 0, H - H * e, W, H); return; }
  if (type === 'wipe') { ctx.drawImage(a, 0, 0, W, H); ctx.save(); ctx.beginPath(); ctx.rect(0, 0, W * e, H); ctx.clip(); ctx.drawImage(b, 0, 0, W, H); ctx.restore(); return; }
  if (type === 'zoom') { ctx.drawImage(a, 0, 0, W, H); const s = 0.6 + 0.4 * e; ctx.globalAlpha = e; ctx.drawImage(b, (W - W * s) / 2, (H - H * s) / 2, W * s, H * s); ctx.globalAlpha = 1; return; }
  ctx.drawImage(a, 0, 0, W, H); ctx.globalAlpha = e; ctx.drawImage(b, 0, 0, W, H); ctx.globalAlpha = 1;
}
export async function exportVideoDialog() {
  const pres = C.getPres();
  const timed = pres.slides.some((s) => s.advanceAfter > 0);
  const v = await form('Exportar a video', [
    { name: 'res', label: 'Calidad', type: 'select', value: '1280', options: [['1920', 'Full HD (1080p)'], ['1280', 'HD (720p)'], ['854', 'Estándar (480p)']] },
    { name: 'secs', label: 'Segundos en cada diapositiva', type: 'number', value: 5, min: 1, max: 120 },
    ...(timed ? [{ name: 'useTimings', label: 'Usar los intervalos ensayados', type: 'checkbox', value: true }] : []),
    { type: 'note', label: 'Se incluyen las transiciones y cada paso de las animaciones. El video se guarda en formato WebM, que se reproduce en navegadores, VLC y la mayoría de los reproductores.' },
  ], 'Exportar');
  if (!v) return null;
  return exportVideo({ width: Number(v.res), secs: Number(v.secs) || 5, useTimings: !!v.useTimings });
}
export async function exportVideo({ width = 1280, secs = 5, useTimings = false, fps = 30, save = true } = {}) {
  const pres = JSON.parse(JSON.stringify(C.getPres()));
  if (typeof VideoEncoder === 'undefined') { toast('Este equipo no puede codificar video.', { error: true }); return null; }
  const states = []; const plan = [];
  pres.slides.forEach((s, i) => {
    const g = buildSteps(s).groups.length; const total = (useTimings && s.advanceAfter > 0 ? s.advanceAfter : secs);
    for (let k = 0; k <= g; k++) { states.push({ i, step: k }); plan.push({ i, k, hold: total / (g + 1), trans: k === 0 && i > 0 ? s.transition : (k > 0 ? { type: 'fade', dur: 0.35 } : null) }); }
  });
  const veil = h('div', { class: 'modal-back' }, h('div', { class: 'modal', style: { padding: '18px 24px', minWidth: '300px' } }, h('div', { class: 'vx-msg' }, t('Preparando las diapositivas…')), h('progress', { max: 100, value: 0, style: { width: '100%' } })));
  document.body.append(veil);
  const msg = veil.querySelector('.vx-msg'), bar = veil.querySelector('progress');
  try {
    const r = await api.renderFrames({ pres, states, width });
    const W = r.width - (r.width % 2), H = r.height - (r.height % 2);
    const imgs = await Promise.all(r.frames.map((b) => createImageBitmap(new Blob([b], { type: 'image/png' }))));
    const cfg = await pickCodec(W, H, fps); if (!cfg) throw new Error(t('Este equipo no puede codificar video.'));
    const muxer = new Muxer({ target: new ArrayBufferTarget(), video: { codec: cfg.mux, width: W, height: H, frameRate: fps } });
    let encErr = null;
    const enc = new VideoEncoder({ output: (chunk, meta) => muxer.addVideoChunk(chunk, meta), error: (e) => { encErr = e; } });
    enc.configure({ codec: cfg.codec, width: W, height: H, bitrate: W >= 1900 ? 8e6 : W >= 1200 ? 5e6 : 2.5e6, framerate: fps });
    const cv = new OffscreenCanvas(W, H); const ctx = cv.getContext('2d');
    const totalFrames = Math.round(plan.reduce((a, p) => a + p.hold, 0) * fps);
    let n = 0;
    const emit = async () => {
      const vf = new VideoFrame(cv, { timestamp: Math.round((n * 1e6) / fps), duration: Math.round(1e6 / fps) });
      enc.encode(vf, { keyFrame: n % (fps * 2) === 0 }); vf.close(); n++;
      if (enc.encodeQueueSize > 8) await new Promise((res) => setTimeout(res, 0));
      if (n % 15 === 0) { bar.value = Math.min(99, (n / totalFrames) * 100); msg.textContent = tf('Creando el video… {0}%', Math.round(bar.value)); await new Promise((res) => setTimeout(res, 0)); }
      if (encErr) throw encErr;
    };
    for (let p = 0; p < plan.length; p++) {
      const cur = imgs[p], prev = p > 0 ? imgs[p - 1] : null;
      const frames = Math.max(1, Math.round(plan[p].hold * fps));
      const tr = plan[p].trans; const tType = tr && tr.type !== 'none' ? (tr.type === 'morph' ? 'fade' : tr.type) : null;
      const tFrames = prev && tType ? Math.min(frames, Math.max(1, Math.round((Number(tr.dur) || 0.6) * fps))) : 0;
      for (let f = 0; f < frames; f++) {
        if (f < tFrames) drawTransition(ctx, prev, cur, tType, (f + 1) / tFrames, W, H);
        else { ctx.globalAlpha = 1; ctx.drawImage(cur, 0, 0, W, H); }
        await emit();
      }
    }
    await enc.flush(); muxer.finalize();
    const bytes = new Uint8Array(muxer.target.buffer);
    bar.value = 100;
    if (!save) return bytes;
    const name = `${C.docName().replace(/\.[^.]+$/, '')}.webm`;
    const res = await api.save({ bytes, ask: true, ext: 'webm', filterName: t('Video WebM'), suggestedName: name });
    if (res) toast(tf('Video guardado: {0}', res.name), { action: { label: 'Mostrar', run: () => api.showInFolder(res.path) } });
    return res;
  } catch (err) {
    console.error(err); toast(tf('No se pudo crear el video: {0}', err.message || err), { error: true, ms: 7000 }); return null;
  } finally { veil.remove(); }
}
export { esc };
