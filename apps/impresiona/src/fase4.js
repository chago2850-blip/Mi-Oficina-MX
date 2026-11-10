// Impresiona — fase 4: cámara en vivo en la diapositiva (tipo Cameo) y grabar la presentación
// con narración por diapositiva; la narración se reproduce en la presentación y entra en el video. GPL-3.0
import { h, toast, modal, form, t, tf } from '../../shared/office.js';

let C = null; // { getPres, slide, commit, select, center, uid, runShow, renderSlide, prepareShow, goto }
export function init(ctx) { C = ctx; }

// ================================================================ cámara en vivo
let camStream = null; let camUsers = 0;
async function getCam() {
  if (!camStream) camStream = await navigator.mediaDevices.getUserMedia({ video: { width: 1280, height: 720 }, audio: false });
  camUsers++;
  return camStream;
}
export function releaseCam() { camUsers = Math.max(0, camUsers - 1); if (!camUsers && camStream) { camStream.getTracks().forEach((tr) => tr.stop()); camStream = null; } }
export function stopAllCams() { camUsers = 0; if (camStream) { camStream.getTracks().forEach((tr) => tr.stop()); camStream = null; } }
export function insertCamera() {
  const pres = C.getPres();
  const s = Math.round(Math.min(pres.size.w, pres.size.h) * 0.36);
  const o = { id: C.uid(), type: 'camera', shape: 'circle', mirror: true, border: '#FFFFFF', ...C.center(s, s), rot: 0, style: {} };
  o.x = pres.size.w - s - 40; o.y = pres.size.h - s - 40;
  C.commit(t('Cámara en vivo'), () => C.slide().objects.push(o));
  C.select([o.id]);
}
export async function cameraOptions(o) {
  if (!o || o.type !== 'camera') { toast('Selecciona primero una cámara.'); return; }
  const v = await form('Cámara en vivo', [
    { name: 'shape', label: 'Forma', type: 'select', value: o.shape || 'circle', options: [['circle', 'Círculo'], ['rounded', 'Rectángulo redondeado'], ['rect', 'Rectángulo']] },
    { name: 'mirror', label: 'Mostrar como espejo', type: 'checkbox', value: o.mirror !== false },
    { name: 'border', label: 'Borde', type: 'select', value: o.border || '#FFFFFF', options: [['#FFFFFF', 'Blanco'], ['none', 'Sin borde'], ['#5B48E8', 'Índigo'], ['#0E9C74', 'Verde'], ['#111111', 'Negro']] },
  ], 'Aceptar');
  if (!v) return;
  C.commit(t('Cámara en vivo'), () => Object.assign(o, v));
}
// En edición se ve una vista previa fija; en la presentación, la cámara en vivo
export function cameraBox(o, mode, box) {
  box.classList.add(`cam-${o.shape || 'circle'}`);
  if (o.border && o.border !== 'none') box.style.boxShadow = `0 0 0 4px ${o.border}`;
  if (mode !== 'show') {
    box.append(h('div', { class: 'cam-ph', html: '<svg viewBox="0 0 24 24" width="38%" height="38%" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M3 7h11a2 2 0 012 2v6a2 2 0 01-2 2H3a1 1 0 01-1-1V8a1 1 0 011-1z"/><path d="M16 11l5-3v8l-5-3"/></svg>' }, h('span', {}, t('Cámara en vivo'))));
    return;
  }
  const v = h('video', { autoplay: true, muted: true, playsinline: true, class: o.mirror !== false ? 'mirror' : '' });
  v.muted = true;
  box.append(v);
  getCam().then((st) => { if (!box.isConnected && !document.body.contains(box)) { releaseCam(); return; } v.srcObject = st; v.play().catch(() => {}); v.dataset.live = '1'; }).catch((e) => { box.append(h('div', { class: 'cam-err' }, t('No se pudo abrir la cámara'))); console.warn(e); });
  // cuando la capa se quita de la presentación, se suelta la cámara
  const mo = new MutationObserver(() => { if (!box.isConnected) { mo.disconnect(); if (v.srcObject) releaseCam(); v.srcObject = null; } });
  requestAnimationFrame(() => { const root = box.closest('.show, .show-stage') || document.body; mo.observe(root, { childList: true, subtree: true }); });
}

// ================================================================ grabar con narración
function blobToDataUrl(b) { return new Promise((res) => { const r = new FileReader(); r.onload = () => res(r.result); r.readAsDataURL(b); }); }
export async function recordShow() {
  let mic;
  try { mic = await navigator.mediaDevices.getUserMedia({ audio: { echoCancellation: true, noiseSuppression: true }, video: false }); } catch { toast('No se pudo usar el micrófono.', { error: true }); return; }
  const ok = await modal({ title: t('Grabar presentación'), body: h('div', {}, h('p', {}, t('Se grabará tu voz en cada diapositiva y el tiempo que pases en ella. Avanza como en una presentación normal; al terminar, pulsa Esc.')), h('p', { class: 'muted' }, t('Las diapositivas que ya tengan narración se volverán a grabar.'))), buttons: [{ label: 'Cancelar', value: false }, { label: 'Comenzar a grabar', primary: true, value: true }] });
  if (!ok) { mic.getTracks().forEach((tr) => tr.stop()); return; }
  const mime = ['audio/webm;codecs=opus', 'audio/webm', 'audio/ogg;codecs=opus'].find((m) => MediaRecorder.isTypeSupported(m)) || '';
  const takes = {}; let rec = null; let curI = null; let t0 = 0;
  const startRec = (i) => {
    curI = i; t0 = performance.now(); const chunks = [];
    rec = new MediaRecorder(mic, mime ? { mimeType: mime } : undefined);
    rec.ondataavailable = (e) => { if (e.data.size) chunks.push(e.data); };
    const my = rec; const myI = i;
    my.onstop = () => { takes[myI] = { blob: new Blob(chunks, { type: my.mimeType || 'audio/webm' }), dur: Math.max(0.5, (performance.now() - t0) / 1000) }; };
    rec.start(250);
  };
  const stopRec = () => new Promise((res) => { if (!rec || rec.state === 'inactive') { res(); return; } const r = rec; const prev = r.onstop; r.onstop = (e) => { prev(e); res(); }; r.stop(); });
  let pending = Promise.resolve();
  C.prepareShow();
  C.runShow({
    renderSlide: C.renderSlide, getPres: C.getPres, from: 0, recording: true,
    onSlide: (i) => { pending = pending.then(async () => { await stopRec(); startRec(i); }); },
    onEnd: async (i) => {
      await pending; await stopRec(); mic.getTracks().forEach((tr) => tr.stop());
      C.goto(i);
      const idx = Object.keys(takes).map(Number);
      if (!idx.length) return;
      const total = idx.reduce((a, k) => a + takes[k].dur, 0);
      const keep = await modal({ title: t('Grabar presentación'), body: h('p', {}, tf('Se grabó la narración de {0} diapositivas ({1}). ¿Guardarla en la presentación?', idx.length, `${Math.floor(total / 60)}:${String(Math.round(total % 60)).padStart(2, '0')}`)), buttons: [{ label: 'Descartar', value: false }, { label: 'Guardar', primary: true, value: true }] });
      if (!keep) return;
      const urls = {}; for (const k of idx) urls[k] = await blobToDataUrl(takes[k].blob);
      const pres = C.getPres();
      C.commit(t('Grabar presentación'), () => { for (const k of idx) { const s = pres.slides[k]; if (!s) continue; s.narration = { src: urls[k], dur: Math.round(takes[k].dur * 10) / 10 }; s.advanceAfter = Math.round(takes[k].dur * 10) / 10; } pres.useTimings = true; pres.useNarration = true; });
      toast('Narración guardada. La presentación avanzará sola con tu voz.');
    },
  });
}
export function clearNarration(all) {
  const pres = C.getPres();
  C.commit(t('Borrar narración'), () => { (all ? pres.slides : [C.slide()]).forEach((s) => { delete s.narration; }); });
}
export function setUseNarration(on) { C.commit(t('Reproducir narración'), () => { C.getPres().useNarration = on; }); }

// Audio de la narración para el video: un solo canal mezclado a 48 kHz
export async function narrationTrack(pres, plan, fps) {
  const ctx = new OfflineAudioContext(1, 48000, 48000);
  const starts = []; let tcur = 0; const done = new Set();
  for (const p of plan) { if (!done.has(p.i)) { starts[p.i] = tcur; done.add(p.i); } tcur += Math.max(1, Math.round(p.hold * fps)) / fps; }
  const total = tcur;
  const out = new Float32Array(Math.ceil(total * 48000));
  let any = false;
  for (const [i, s] of pres.slides.entries()) {
    if (!s.narration?.src || starts[i] == null) continue;
    try {
      const bytes = await (await fetch(s.narration.src)).arrayBuffer();
      const buf = await ctx.decodeAudioData(bytes);
      const ch = buf.getChannelData(0); const off = Math.round(starts[i] * 48000);
      const src = buf.sampleRate === 48000 ? ch : resample(ch, buf.sampleRate, 48000);
      for (let k = 0; k < src.length && off + k < out.length; k++) out[off + k] += src[k];
      any = true;
    } catch (e) { console.warn('narración', e); }
  }
  return any ? out : null;
}
function resample(data, from, to) {
  const n = Math.round(data.length * to / from); const o = new Float32Array(n);
  for (let i = 0; i < n; i++) { const x = i * from / to; const a = Math.floor(x); const f = x - a; o[i] = (data[a] || 0) * (1 - f) + (data[a + 1] || 0) * f; }
  return o;
}
