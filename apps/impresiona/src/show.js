// Motor de presentación de Impresiona: transiciones, animaciones por pasos, pluma y vista del moderador. GPL-3.0
import { h, icon, t } from '../../shared/office.js';
import { stopAllCams } from './fase4.js';

export const EXIT_TYPES = ['fadeOut', 'flyOut', 'zoomOut'];
// Trayectorias de animación: el objeto se ve desde el inicio y se mueve por el camino
export const PATH_TYPES = ['pathRight', 'pathLeft', 'pathUp', 'pathDown', 'pathArc', 'pathCircle', 'pathWave'];
export function pathPoints(type, d = 200) {
  const n = 24; const pts = [];
  for (let k = 0; k <= n; k++) {
    const u = k / n;
    if (type === 'pathRight') pts.push([d * u, 0]); else if (type === 'pathLeft') pts.push([-d * u, 0]);
    else if (type === 'pathUp') pts.push([0, -d * u]); else if (type === 'pathDown') pts.push([0, d * u]);
    else if (type === 'pathArc') pts.push([d * u, -Math.sin(Math.PI * u) * d * 0.45]);
    else if (type === 'pathCircle') pts.push([Math.sin(2 * Math.PI * u) * d * 0.4, (1 - Math.cos(2 * Math.PI * u)) * d * 0.4 * -1]);
    else if (type === 'pathWave') pts.push([d * u, -Math.sin(4 * Math.PI * u) * d * 0.15]);
  }
  return pts;
}
const isPath = (o) => PATH_TYPES.includes(o.anim?.type);
const isExit = (o) => EXIT_TYPES.includes(o.anim?.type);

// Agrupa las animaciones de una diapositiva en pasos de clic
export function buildSteps(slide) {
  const items = slide.objects.filter((o) => o.anim && o.anim.type && o.anim.type !== 'none').sort((a, b) => (a.anim.order || 0) - (b.anim.order || 0));
  const auto = []; const groups = [];
  for (const o of items) {
    const start = o.anim.start || 'click';
    if (start === 'click' || (!groups.length && start === 'click')) groups.push([o]);
    else if (groups.length) groups[groups.length - 1].push(o);
    else auto.push(o);
  }
  return { auto, groups };
}

export class ShowEngine {
  constructor(container, { renderSlide, getPres, onChange, mode = 'show' }) {
    this.container = container; this.renderSlide = renderSlide; this.getPres = getPres; this.onChange = onChange; this.mode = mode;
    this.i = 0; this.step = 0; this.ink = new Map(); this.pen = false;
  }
  get pres() { return this.getPres(); }
  scale() { const { w, h: hh } = this.pres.size; return Math.min(this.container.clientWidth / w, this.container.clientHeight / hh) || 1; }
  layer() { return this.container.querySelector('.show-layer:last-child'); }
  mount(i, { dir = 1, animate = true, step = 0 } = {}) {
    const pres = this.pres;
    clearTimeout(this.autoT);
    const prevI = this.i;
    this.i = Math.max(0, Math.min(pres.slides.length - 1, i));
    const s = pres.slides[this.i];
    const { w: W, h: H } = pres.size; const sc = this.scale();
    const el = this.renderSlide(s, 'show');
    el.style.transform = `scale(${sc})`; el.style.transformOrigin = '0 0';
    const layer = h('div', { class: 'show-layer', style: { width: `${W * sc}px`, height: `${H * sc}px` } }, el);
    const t = s.transition?.type || 'none';
    const old = [...this.container.querySelectorAll('.show-layer')];
    if (animate && t !== 'none' && t !== 'morph' && old.length) { layer.classList.add(`tr-${t}`, dir < 0 ? 'rev' : 'fwd'); layer.style.animationDuration = `${s.transition.dur || 0.6}s`; }
    this.container.append(layer);
    if (animate && t === 'morph' && old.length && pres.slides[prevI] && prevI !== this.i) this.morph(pres.slides[prevI], s, old, layer, (s.transition.dur || 1) * 1000);
    setTimeout(() => old.forEach((o) => o.remove()), animate && t !== 'none' ? (s.transition.dur || 0.6) * 1000 : 0);
    this.steps = buildSteps(s);
    this.step = 0;
    // las de inicio automático se reproducen al entrar
    this.play(this.steps.auto, animate);
    for (let k = 0; k < step && k < this.steps.groups.length; k++) this.play(this.steps.groups[k], false);
    this.step = Math.min(step, this.steps.groups.length);
    this.drawInk();
    if (this.mode === 'show') layer.querySelectorAll('video[data-autoplay], audio[data-autoplay]').forEach((m) => { m.currentTime = 0; m.play().catch(() => {}); });
    // narración grabada de la diapositiva
    this.narr?.pause(); this.narr = null;
    if (this.mode === 'show' && !this.noNarration && pres.useNarration !== false && s.narration?.src) { this.narr = new Audio(s.narration.src); this.narr.play().catch(() => {}); }
    // intervalos ensayados: avance automático
    if (this.mode === 'show' && this.useTimings && s.advanceAfter > 0) this.autoT = setTimeout(() => this.onAuto?.(), s.advanceAfter * 1000);
    this.onChange?.(this.state());
  }
  // Transformación (Morph): los objetos que coinciden entre diapositivas se mueven y cambian de tamaño
  morph(from, to, oldLayers, layer, ms) {
    const key = (o) => o.mk || o.id;
    const sig = (o) => `${o.type}|${o.shape || ''}|${(o.html || o.src || o.title || '').slice(0, 200)}`;
    const used = new Set();
    for (const o of to.objects) {
      const el = layer.querySelector(`[data-id="${o.id}"]`); if (!el) continue;
      let p = from.objects.find((q) => !used.has(q) && key(q) === key(o)) || from.objects.find((q) => !used.has(q) && sig(q) === sig(o) && q.type === o.type);
      if (!p) { el.animate([{ opacity: 0 }, { opacity: 1 }], { duration: ms, easing: 'ease', fill: 'backwards' }); continue; }
      used.add(p);
      const sx = (p.w || 1) / (o.w || 1), sy = (p.h || 1) / (o.h || 1);
      const fromT = `translate(${p.x - o.x}px, ${p.y - o.y}px) scale(${sx}, ${sy}) rotate(${p.rot || 0}deg)`;
      el.style.transformOrigin = '0 0';
      el.animate([{ transform: fromT, opacity: 1 }, { transform: `rotate(${o.rot || 0}deg)`, opacity: 1 }], { duration: ms, easing: 'cubic-bezier(.45,.05,.25,1)' }).finished.then(() => { el.style.transformOrigin = ''; }).catch(() => {});
    }
    // lo que solo estaba en la diapositiva anterior se desvanece; el fondo se funde
    layer.animate([{ opacity: 0.0001 }, { opacity: 1 }], { duration: Math.min(ms, 350), fill: 'backwards' });
    for (const ol of oldLayers) { ol.style.zIndex = '2'; ol.style.pointerEvents = 'none'; ol.animate([{ opacity: 1 }, { opacity: 0 }], { duration: ms * 0.6, fill: 'forwards' }); }
    setTimeout(() => oldLayers.forEach((x) => x.remove()), ms * 0.6 + 30);
  }
  play(group, animate = true) {
    const layer = this.layer(); if (!layer) return;
    let t = 0;
    for (const o of group) {
      const el = layer.querySelector(`[data-id="${o.id}"]`); if (!el) continue;
      const dur = Number(o.anim.dur) || 0.6, delay = Number(o.anim.delay) || 0;
      if ((o.anim.start || 'click') === 'after') t += delay; else t = Math.max(0, t - (group.indexOf(o) ? dur : 0)) + delay;
      const startAt = (o.anim.start || 'click') === 'after' ? t : delay;
      if (isPath(o)) {
        const pts = pathPoints(o.anim.type, Number(o.anim.dist) || 200); const [ex, ey] = pts[pts.length - 1];
        const base = o.rot ? ` rotate(${o.rot}deg)` : '';
        if (!animate) { el.style.transform = `translate(${ex}px, ${ey}px)${base}`; continue; }
        el.animate(pts.map(([x, y]) => ({ transform: `translate(${x}px, ${y}px)${base}` })), { duration: dur * 1000, delay: startAt * 1000, easing: 'ease-in-out', fill: 'forwards' });
        if ((o.anim.start || 'click') === 'after') t += dur;
        continue;
      }
      if (!animate) { if (isExit(o)) el.classList.add('gone'); else el.classList.remove('anim-wait'); continue; }
      el.style.animationDuration = `${dur}s`; el.style.animationDelay = `${startAt}s`;
      if (isExit(o)) { el.classList.add(`an-${o.anim.type}`); setTimeout(() => el.classList.add('gone'), (startAt + dur) * 1000); }
      else { el.classList.remove('anim-wait'); if (startAt) { el.style.visibility = 'hidden'; setTimeout(() => { el.style.visibility = ''; }, startAt * 1000); } el.classList.add(`an-${o.anim.type}`); }
      if ((o.anim.start || 'click') === 'after') t += dur;
    }
  }
  next() {
    if (this.step < this.steps.groups.length) { this.play(this.steps.groups[this.step]); this.step++; this.onChange?.(this.state()); return true; }
    if (this.i < this.pres.slides.length - 1) { this.mount(this.i + 1, { dir: 1 }); return true; }
    return false;
  }
  prev() {
    if (this.step > 0) { this.mount(this.i, { animate: false, step: this.step - 1 }); return true; }
    if (this.i > 0) { const prevSteps = buildSteps(this.pres.slides[this.i - 1]).groups.length; this.mount(this.i - 1, { dir: -1, step: prevSteps }); return true; }
    return false;
  }
  setState(i, step, animate = false) {
    if (i === this.i && step === this.step + 1) { this.next(); return; }
    this.mount(i, { animate: animate && i !== this.i, step, dir: i < this.i ? -1 : 1 });
  }
  state() { return { i: this.i, step: this.step, steps: this.steps?.groups.length || 0, total: this.pres.slides.length }; }
  // ---- pluma
  setPen(on) { this.pen = on; this.container.classList.toggle('pen', on); this.drawInk(); }
  clearInk() { this.ink.delete(this.i); this.drawInk(); }
  inkCount() { let n = 0; for (const v of this.ink.values()) n += v.length; return n; }
  drawInk() {
    const layer = this.layer(); if (!layer) return;
    let cv = layer.querySelector('canvas.ink');
    if (!cv) { cv = h('canvas', { class: 'ink' }); layer.append(cv); this.bindInk(cv); }
    const r = layer.getBoundingClientRect(); cv.width = r.width; cv.height = r.height;
    const ctx = cv.getContext('2d'); ctx.clearRect(0, 0, cv.width, cv.height);
    ctx.lineCap = 'round'; ctx.lineJoin = 'round'; ctx.strokeStyle = '#FF2D55'; ctx.lineWidth = 4;
    for (const stroke of this.ink.get(this.i) || []) { ctx.beginPath(); stroke.forEach(([x, y], k) => (k ? ctx.lineTo(x * cv.width, y * cv.height) : ctx.moveTo(x * cv.width, y * cv.height))); ctx.stroke(); }
    cv.style.pointerEvents = this.pen ? 'auto' : 'none';
  }
  bindInk(cv) {
    cv.addEventListener('mousedown', (e) => {
      if (!this.pen) return;
      e.stopPropagation(); e.preventDefault();
      const r = cv.getBoundingClientRect();
      const stroke = [[(e.clientX - r.left) / r.width, (e.clientY - r.top) / r.height]];
      const list = this.ink.get(this.i) || []; list.push(stroke); this.ink.set(this.i, list);
      const move = (ev) => { stroke.push([(ev.clientX - r.left) / r.width, (ev.clientY - r.top) / r.height]); this.drawInk(); };
      const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); };
      addEventListener('mousemove', move); addEventListener('mouseup', up);
    });
    cv.addEventListener('click', (e) => { if (this.pen) e.stopPropagation(); });
  }
}

// ---------------------------------------------------------------- presentación en la misma ventana
export function runShow({ renderSlide, getPres, from = 0, onEnd, rehearse = false, onRehearsed, recording = false, onSlide, loop = false, pauseSec = 0, manual = false, clickBg = true, pen = false, laserOn = false }) {
  const show = h('div', { class: 'show', tabindex: 0 });
  const stageS = h('div', { class: 'show-stage' });
  const cnt = h('span', { class: 'cnt' });
  const btn = (title, ic, fn) => h('button', { title, html: icon(ic, 18), onclick: (e) => { e.stopPropagation(); fn(); } });
  const laser = h('div', { class: 'laser-dot' });
  show.append(stageS, laser);
  document.body.append(show);
  document.documentElement.requestFullscreen?.().catch(() => {});
  // Ensayar intervalos: se mide cuánto tiempo pasa en cada diapositiva
  const times = []; let tSlide = Date.now(), tTotal = Date.now(), lastI = from;
  const rehBox = rehearse ? h('div', { class: 'rehearse' }, h('span', { class: 'rh-slide' }, '0:00'), h('span', { class: 'rh-total' }, '0:00')) : null;
  const fmt = (ms) => { const sec = Math.floor(ms / 1000); return `${Math.floor(sec / 60)}:${String(sec % 60).padStart(2, '0')}`; };
  const rehTick = rehearse ? setInterval(() => { rehBox.firstChild.textContent = fmt(Date.now() - tSlide); rehBox.lastChild.textContent = fmt(Date.now() - tTotal); }, 250) : 0;
  if (rehBox) show.append(rehBox);
  if (recording) show.append(h('div', { class: 'rec-dot' }, h('i'), 'REC'));
  let slideSeen = null;
  const eng = new ShowEngine(stageS, { renderSlide, getPres, onChange: (st) => {
    cnt.textContent = `${st.i + 1} / ${st.total}`;
    if (rehearse && st.i !== lastI) { times[lastI] = (times[lastI] || 0) + (Date.now() - tSlide); tSlide = Date.now(); lastI = st.i; }
    if (st.i !== slideSeen) { slideSeen = st.i; onSlide?.(st.i); }
  } });
  eng.useTimings = !rehearse && !recording && !manual && getPres().useTimings !== false;
  eng.noNarration = recording;
  // con intervalos, al terminar el tiempo se pasa a la siguiente diapositiva (con sus pasos ya incluidos)
  eng.onAuto = () => { if (ended()) return; if (eng.i < getPres().slides.length - 1) eng.mount(eng.i + 1, { dir: 1 }); else next(); };
  // acciones al hacer clic en un objeto (hipervínculos y saltos)
  stageS.addEventListener('click', (e) => {
    const a = e.target.closest?.('[data-link]'); if (!a) return;
    e.stopPropagation(); e.preventDefault();
    const l = a.dataset.link; const P = getPres(); const n = P.slides.length;
    if (l.startsWith('#slide:')) { const i = P.slides.findIndex((s) => s.id === l.slice(7)); if (i >= 0) eng.mount(i, { dir: i < eng.i ? -1 : 1 }); }
    else if (l === '#next') next(); else if (l === '#prev') prev(); else if (l === '#first') eng.mount(0, { dir: -1 }); else if (l === '#last') eng.mount(n - 1); else if (l === '#end') end();
    else if (/^(https?:|mailto:)/i.test(l)) { if (window.oficina?.openExternal) window.oficina.openExternal(l); else window.open(l, '_blank'); }
  }, true);
  const ended = () => show.classList.contains('ended');
  let loopT = 0;
  const next = () => {
    if (ended()) { end(); return; }
    if (show.classList.contains('pausing')) return;
    if (eng.next()) return;
    clearTimeout(eng.autoT);
    // en bucle: pausa (con cuenta atrás) y se vuelve a empezar, como en un quiosco
    if (loop) {
      show.classList.add('pausing');
      const box = h('div', { class: 'show-pause' }, h('b', {}, String(Math.ceil(pauseSec))), h('span', {}, t('La presentación volverá a empezar')));
      stageS.append(box); let left = pauseSec;
      const tick = () => { if (!show.isConnected) return; if (left <= 0) { box.remove(); show.classList.remove('pausing'); eng.mount(0, { animate: false }); return; } box.firstChild.textContent = String(Math.ceil(left)); left -= 1; loopT = setTimeout(tick, 1000); };
      tick(); return;
    }
    show.classList.add('ended'); stageS.append(h('div', { class: 'show-end' }, t('Fin de la presentación. Haz clic para salir.')));
  };
  const prev = () => { if (ended()) { show.classList.remove('ended'); stageS.querySelector('.show-end')?.remove(); return; } eng.prev(); };
  const end = () => {
    clearTimeout(eng.autoT); clearInterval(rehTick); clearTimeout(loopT);
    show.querySelectorAll('video, audio').forEach((m) => m.pause());
    eng.narr?.pause(); stopAllCams();
    show.remove(); if (document.fullscreenElement) document.exitFullscreen?.(); removeEventListener('resize', onResize);
    if (rehearse) { times[lastI] = (times[lastI] || 0) + (Date.now() - tSlide); onRehearsed?.(times.map((x) => (x ? Math.max(1, Math.round(x / 100) / 10) : 0)), Date.now() - tTotal); }
    onEnd?.(eng.i);
  };
  const ctrls = h('div', { class: 'show-ctrls' }, btn('Anterior', 'undo', prev), cnt, btn('Siguiente', 'redo', next), btn('Puntero láser (L)', 'star', () => show.classList.toggle('laser')), btn('Pluma (P)', 'brush', () => eng.setPen(!eng.pen)), btn('Borrar tinta (E)', 'clear', () => eng.clearInk()), btn('Terminar (Esc)', 'close', end));
  show.append(ctrls);
  const onResize = () => { if (!ended()) eng.mount(eng.i, { animate: false, step: eng.step }); };
  addEventListener('resize', onResize);
  show.addEventListener('click', (e) => { if (eng.pen) return; if (!clickBg && !e.target.closest?.('.show-ctrls')) return; next(); });
  show.addEventListener('contextmenu', (e) => { e.preventDefault(); prev(); });
  show.addEventListener('mousemove', (e) => { show.classList.add('moving'); clearTimeout(show.t); show.t = setTimeout(() => show.classList.remove('moving'), 1800); laser.style.transform = `translate(${e.clientX}px, ${e.clientY}px)`; });
  let typed = '';
  show.addEventListener('keydown', (e) => {
    e.stopPropagation();
    const k = e.key;
    if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'n'].includes(k)) { e.preventDefault(); if (k === 'Enter' && typed) { const n = Number(typed) - 1; typed = ''; if (n >= 0 && n < getPres().slides.length) { show.classList.remove('ended'); eng.mount(n, { animate: false }); } return; } next(); }
    else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'].includes(k)) { e.preventDefault(); prev(); }
    else if (k === 'Escape') { e.preventDefault(); if (eng.pen) eng.setPen(false); else end(); }
    else if (k === 'Home') eng.mount(0, { animate: false });
    else if (k === 'End') eng.mount(getPres().slides.length - 1, { animate: false });
    else if (k.toLowerCase() === 'b') show.classList.toggle('black');
    else if (k.toLowerCase() === 'w') show.classList.toggle('white');
    else if (k.toLowerCase() === 'l') show.classList.toggle('laser');
    else if (k.toLowerCase() === 'p') eng.setPen(!eng.pen);
    else if (k.toLowerCase() === 'e') eng.clearInk();
    else if (/^\d$/.test(k)) typed += k;
  });
  show.focus();
  eng.mount(from, { animate: false });
  if (pen) eng.setPen(true); if (laserOn) show.classList.add('laser');
  window.__show = { next, prev, end, inkCount: () => eng.inkCount(), engine: eng, times: () => times };
  return eng;
}

// ---------------------------------------------------------------- vista del moderador
export function runPresenter({ renderSlide, getPres, from = 0, api, onEnd }) {
  const pres = getPres();
  const root = h('div', { class: 'presenter', tabindex: 0 });
  const curBox = h('div', { class: 'pv-cur' }), nextBox = h('div', { class: 'pv-next' });
  const notes = h('div', { class: 'pv-notes' });
  const timer = h('div', { class: 'pv-timer' }, '00:00'), clock = h('div', { class: 'pv-clock' });
  const cnt = h('div', { class: 'pv-cnt' });
  const t0 = Date.now();
  const tick = setInterval(() => { const s = Math.floor((Date.now() - t0) / 1000); timer.textContent = `${String(Math.floor(s / 60)).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`; clock.textContent = new Date().toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' }); }, 500);
  const b = (label, ic, fn) => h('button', { class: 'btn', html: `${icon(ic, 16)} ${label}`, onclick: fn });
  let black = false;
  root.append(
    h('div', { class: 'pv-top' }, h('b', {}, 'Vista del moderador'), timer, clock, h('div', { class: 'grow' }), b('Pantalla negra', 'background', () => { black = !black; api?.sendAudience({ type: 'black', on: black }); }), b('Terminar', 'close', () => end())),
    h('div', { class: 'pv-main' }, h('div', { class: 'pv-left' }, curBox, h('div', { class: 'pv-ctrl' }, b('Anterior', 'undo', () => prev()), cnt, b('Siguiente', 'redo', () => next()))), h('div', { class: 'pv-right' }, h('small', {}, 'Siguiente'), nextBox, h('small', {}, 'Notas del orador'), notes)),
  );
  document.body.append(root);
  const renderNext = (st) => {
    nextBox.replaceChildren();
    const ni = st.step < st.steps ? st.i : st.i + 1;
    if (ni < pres.slides.length) { const el = renderSlide(pres.slides[ni], 'thumb'); const sc = nextBox.clientWidth / pres.size.w; el.style.transform = `scale(${sc})`; el.style.transformOrigin = '0 0'; nextBox.style.height = `${pres.size.h * sc}px`; nextBox.append(el); }
    else nextBox.append(h('div', { class: 'pv-endnote' }, 'Fin de la presentación'));
    notes.textContent = pres.slides[st.i].notes || '(sin notas)';
    cnt.textContent = `Diapositiva ${st.i + 1} de ${st.total}${st.steps ? ` · paso ${st.step}/${st.steps}` : ''}`;
  };
  const eng = new ShowEngine(curBox, { renderSlide, getPres, onChange: (st) => { renderNext(st); api?.sendAudience({ type: 'state', i: st.i, step: st.step }); } });
  const next = () => eng.next();
  const prev = () => eng.prev();
  let closed = false;
  const end = () => { if (closed) return; closed = true; clearInterval(tick); root.remove(); api?.closeAudience(); onEnd?.(eng.i); };
  root.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter'].includes(e.key)) { e.preventDefault(); next(); }
    else if (['ArrowLeft', 'ArrowUp', 'PageUp'].includes(e.key)) { e.preventDefault(); prev(); }
    else if (e.key === 'Escape') end();
    else if (e.key.toLowerCase() === 'b') { black = !black; api?.sendAudience({ type: 'black', on: black }); }
  });
  api?.onAudienceClosed?.(() => end());
  root.focus();
  requestAnimationFrame(() => eng.mount(from, { animate: false }));
  window.__presenter = { next, prev, end, engine: eng };
  return eng;
}

// ---------------------------------------------------------------- ventana del público / render a imagen
export function runAudience({ renderSlide, setPres, getPres, api, render = false }) {
  document.body.className = render ? 'render-mode' : 'audience-mode';
  document.body.replaceChildren();
  const stage = h('div', { class: 'show-stage aud' });
  document.body.append(stage);
  let eng = null;
  api.onAudience(async (msg) => {
    if (msg.type === 'load') { setPres(msg.pres); eng = new ShowEngine(stage, { renderSlide, getPres }); if (!render) eng.mount(msg.i || 0, { animate: false, step: msg.step || 0 }); }
    if (msg.type === 'state' && eng) eng.setState(msg.i, msg.step, true);
    if (msg.type === 'black') stage.style.visibility = msg.on ? 'hidden' : '';
    if (msg.type === 'render') {
      const s = getPres().slides[msg.i];
      if (msg.step != null) {
        // un estado intermedio de las animaciones (para exportar a video)
        stage.replaceChildren();
        const e2 = new ShowEngine(stage, { renderSlide, getPres, mode: 'render' });
        e2.mount(msg.i, { animate: false, step: msg.step });
      } else {
        const el = renderSlide(s, 'print');
        stage.replaceChildren(el);
      }
      await document.fonts?.ready;
      await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
      api.renderReady(msg.i);
    }
  });
  addEventListener('keydown', (e) => { if (e.key === 'Escape' && !render) window.close(); });
  window.__audience = { index: () => eng?.i ?? -1 };
}
