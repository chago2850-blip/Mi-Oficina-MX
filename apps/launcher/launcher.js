// Launcher tipo dona de Mi Oficina MX.
// Cerrada: dona oscura con el botón "Inicio". Al pulsarlo se enciende y se divide en 4 sectores.
'use strict';

const APPS = [
  { id: 'escribiente', name: 'Escribiente', hint: 'Procesador de textos' },
  { id: 'calcula', name: 'Calcula Datos', hint: 'Hoja de cálculo' },
  { id: 'impresiona', name: 'Impresiona', hint: 'Presentaciones' },
  { id: 'papelpdf', name: 'PapelPDF', hint: 'Editor de PDF' },
];
const C = 200, R1 = 104, R2 = 186, GAP = 5; // grados de separación a cada lado
const NS = 'http://www.w3.org/2000/svg';

const stage = document.getElementById('stage');
const sectorsG = document.getElementById('sectors');
const core = document.getElementById('core');
const coreLabel = document.getElementById('core-label');

const pt = (r, deg) => {
  const a = (deg * Math.PI) / 180;
  return [C + r * Math.cos(a), C + r * Math.sin(a)];
};
function arcPath(a0, a1) {
  const [x0, y0] = pt(R2, a0), [x1, y1] = pt(R2, a1), [x2, y2] = pt(R1, a1), [x3, y3] = pt(R1, a0);
  return `M${x0},${y0} A${R2},${R2} 0 0 1 ${x1},${y1} L${x2},${y2} A${R1},${R1} 0 0 0 ${x3},${y3} Z`;
}
function el(tag, attrs, parent) {
  const n = document.createElementNS(NS, tag);
  for (const [k, v] of Object.entries(attrs)) n.setAttribute(k, v);
  if (parent) parent.appendChild(n);
  return n;
}

// Sectores: Escribiente arriba-derecha, Calcula abajo-derecha, Impresiona abajo-izquierda, PapelPDF arriba-izquierda
APPS.forEach((app, i) => {
  const a0 = -90 + i * 90 + GAP, a1 = -90 + (i + 1) * 90 - GAP, mid = (a0 + a1) / 2;
  const [dx, dy] = [Math.cos((mid * Math.PI) / 180) * 5, Math.sin((mid * Math.PI) / 180) * 5];
  const g = el('g', { class: 'sector', tabindex: '-1', role: 'button', 'data-app': app.id, 'aria-label': `${app.name}: ${app.hint}` }, sectorsG);
  g.style.setProperty('--dx', `${dx}px`);
  g.style.setProperty('--dy', `${dy}px`);
  el('path', { class: 'arc', d: arcPath(a0, a1), fill: `url(#g-${app.id})` }, g);
  const [ix, iy] = pt((R1 + R2) / 2, mid);
  el('image', { href: `../shared/img/${app.id}.svg`, x: ix - 26, y: iy - 34, width: 52, height: 52 }, g);
  const t = el('text', { class: 'label', x: ix, y: iy + 34, 'text-anchor': 'middle' }, g);
  t.textContent = app.name;
  const [rx, ry] = pt(R2 - 10, mid);
  el('circle', { class: 'running', cx: rx, cy: ry, r: 3.5 }, g);
  const title = el('title', {}, g);
  title.textContent = `${app.name} — ${app.hint}`;
  g.addEventListener('click', () => launch(app.id, g));
  g.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); launch(app.id, g); }
  });
});
const sectors = [...sectorsG.querySelectorAll('.sector')];

function setOpen(open) {
  stage.classList.toggle('open', open);
  stage.classList.toggle('closed', !open);
  coreLabel.textContent = open ? 'Cerrar' : 'Inicio';
  core.setAttribute('aria-label', open ? 'Cerrar' : 'Inicio');
  core.setAttribute('aria-expanded', String(open));
  sectors.forEach((s) => s.setAttribute('tabindex', open ? '0' : '-1'));
  if (open) setTimeout(() => sectors[0].focus({ preventScroll: true }), 350);
}
const isOpen = () => stage.classList.contains('open');

core.addEventListener('click', () => setOpen(!isOpen()));
core.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setOpen(!isOpen()); }
});

let launching = null;
async function launch(id, g) {
  if (launching) return;
  launching = id;
  g.classList.add('launching');
  try { await window.launcher.open(id); } finally {
    setTimeout(() => { g.classList.remove('launching'); launching = null; }, 600);
  }
}

// Teclado: flechas recorren los sectores, Esc cierra la dona.
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') { if (!prefsCard.hidden) prefsCard.hidden = true; else if (!document.getElementById('about').hidden) closeAbout(); else if (isOpen()) { setOpen(false); core.focus(); } }
  if (!isOpen()) return;
  const idx = sectors.indexOf(document.activeElement);
  if (['ArrowRight', 'ArrowDown'].includes(e.key)) { e.preventDefault(); sectors[(idx + 1 + 4) % 4].focus(); }
  if (['ArrowLeft', 'ArrowUp'].includes(e.key)) { e.preventDefault(); sectors[(idx - 1 + 4) % 4].focus(); }
  const n = Number(e.key);
  if (n >= 1 && n <= 4) launch(APPS[n - 1].id, sectors[n - 1]);
});

// Barra
document.getElementById('btn-menu').addEventListener('click', () => window.launcher.menu());
document.getElementById('btn-min').addEventListener('click', () => window.launcher.hide());
document.getElementById('btn-quit').addEventListener('click', () => window.launcher.quit());
document.getElementById('dona').addEventListener('contextmenu', (e) => { e.preventDefault(); window.launcher.menu(); });

// Acerca de
const about = document.getElementById('about');
function closeAbout() { about.hidden = true; }
document.getElementById('about-ok').addEventListener('click', closeAbout);
window.launcher.onAbout(async () => {
  const info = await window.launcher.info();
  document.getElementById('ver').textContent = `v${info.version}`;
  about.hidden = false;
});

// Preferencias
const prefsCard = document.getElementById('prefs');
window.launcher.onPrefs(async () => {
  const p = await window.launcher.prefs();
  document.getElementById('pf-author').value = p.author || '';
  document.getElementById('pf-theme').value = p.theme || 'auto';
  document.getElementById('pf-lang').value = p.lang || 'auto';
  const sel = document.getElementById('pf-auto'); sel.value = String(p.autosaveSec || 20); if (!sel.value) sel.value = '20';
  prefsCard.hidden = false;
  document.getElementById('pf-author').focus();
});
document.getElementById('prefs-cancel').addEventListener('click', () => { prefsCard.hidden = true; });
document.getElementById('prefs-ok').addEventListener('click', async () => {
  await window.launcher.setPrefs({ author: document.getElementById('pf-author').value.trim(), theme: document.getElementById('pf-theme').value, lang: document.getElementById('pf-lang').value, autosaveSec: Number(document.getElementById('pf-auto').value) });
  prefsCard.hidden = true;
});

// Estado que manda el proceso principal
window.launcher.onRunning((list) => sectors.forEach((s) => s.classList.toggle('is-running', list.includes(s.dataset.app))));
window.launcher.onRestored(() => {
  // Al volver, la dona aparece encendida para elegir otra app al instante.
  setOpen(true);
});
window.launcher.onMinimizing(() => {});

// Para pruebas automáticas
window.__launcher = { setOpen, isOpen, launch: (id) => launch(id, sectors[APPS.findIndex((a) => a.id === id)]) };
