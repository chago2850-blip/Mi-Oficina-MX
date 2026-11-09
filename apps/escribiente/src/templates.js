// Plantillas originales de Escribiente (escritas desde cero). GPL-3.0
const today = () => new Date().toLocaleDateString('es-MX', { day: 'numeric', month: 'long', year: 'numeric' });

export const TEMPLATES = [
  { id: 'blank', name: 'Documento en blanco', blank: true, html: () => '<p><br></p>' },
  {
    id: 'carta', name: 'Carta formal',
    html: () => `
<p style="text-align:right">Acapulco, Guerrero, a ${today()}</p>
<p><br></p>
<p><b>Lic. Nombre Apellido</b><br>Puesto<br>Empresa o institución<br>P r e s e n t e</p>
<p><br></p>
<p>Estimado(a) Lic. Apellido:</p>
<p style="text-align:justify">Por medio de la presente me dirijo a usted para [motivo de la carta]. Escribe aquí el primer párrafo, donde explicas de forma breve y clara el propósito.</p>
<p style="text-align:justify">En el segundo párrafo puedes dar los detalles, los antecedentes o la información de apoyo que la persona necesita para responder.</p>
<p style="text-align:justify">Agradezco de antemano su atención y quedo a sus órdenes para cualquier aclaración.</p>
<p><br></p>
<p style="text-align:center">A t e n t a m e n t e</p>
<p><br></p><p><br></p>
<p style="text-align:center">______________________________<br><b>Tu nombre</b><br>Tu cargo · tu teléfono · tu correo</p>`,
  },
  {
    id: 'reporte', name: 'Reporte',
    html: () => `
<p class="title">Título del reporte</p>
<p class="subtitle">Subtítulo o área · ${today()}</p>
<h1>Resumen</h1>
<p style="text-align:justify">Escribe en dos o tres frases lo más importante: qué se hizo, qué se encontró y qué se recomienda.</p>
<h1>Antecedentes</h1>
<p style="text-align:justify">Explica el contexto y por qué se hizo este trabajo.</p>
<h1>Resultados</h1>
<h2>Hallazgo principal</h2>
<p>Describe el resultado más importante.</p>
<table><tbody><tr><th>Indicador</th><th>Antes</th><th>Después</th></tr><tr><td>Ventas</td><td>$120,000</td><td>$148,500</td></tr><tr><td>Clientes</td><td>310</td><td>362</td></tr></tbody></table>
<h2>Otros hallazgos</h2>
<ul><li>Primer punto.</li><li>Segundo punto.</li><li>Tercer punto.</li></ul>
<h1>Recomendaciones</h1>
<ol><li>Primera recomendación.</li><li>Segunda recomendación.</li></ol>`,
  },
  {
    id: 'cv', name: 'Currículum',
    html: () => `
<p class="title">Nombre Apellido</p>
<p class="subtitle">Profesión · Ciudad · 744 000 0000 · correo@ejemplo.com</p>
<h1>Perfil</h1>
<p style="text-align:justify">Dos o tres líneas sobre quién eres, en qué eres bueno y qué buscas.</p>
<h1>Experiencia</h1>
<h2>Puesto — Empresa</h2>
<p><i>Enero 2023 – actual</i></p>
<ul><li>Logro medible número uno.</li><li>Logro medible número dos.</li></ul>
<h2>Puesto anterior — Empresa</h2>
<p><i>2020 – 2022</i></p>
<ul><li>Responsabilidad o logro.</li></ul>
<h1>Educación</h1>
<p><b>Licenciatura en …</b> — Universidad, 2019</p>
<h1>Habilidades</h1>
<p>Habilidad · Habilidad · Habilidad · Idioma (nivel)</p>`,
  },
];

// Miniatura simple dibujada con HTML (no imágenes externas).
export function thumb(t) {
  if (t.blank) return '<svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.4"><path d="M12 5v14 M5 12h14"/></svg>';
  const lines = { carta: [70, 0, 50, 40, 45, 0, 92, 96, 90, 94, 60, 0, 40], reporte: [80, 50, 0, 30, 95, 90, 30, 70, 0, 96, 92], cv: [70, 85, 0, 30, 95, 80, 30, 60, 90, 85] }[t.id] || [];
  return `<div style="padding:14px 14px;display:grid;gap:5px">${lines.map((w, i) => `<div style="height:${i === 0 && t.id !== 'carta' ? 8 : 4}px;width:${w}%;background:${w === 30 ? 'var(--accent)' : '#D5D7E0'};border-radius:2px;${w === 0 ? 'background:none' : ''};${t.id === 'carta' && i === 0 ? 'margin-left:auto' : ''}"></div>`).join('')}</div>`;
}
