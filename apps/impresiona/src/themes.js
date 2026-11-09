// Temas y diseños originales de Impresiona. GPL-3.0
export const THEMES = [
  { id: 'claro', name: 'Claro', bg: '#FFFFFF', title: '#1D2030', text: '#3A3F52', accent: '#DC3A63', accent2: '#5B48E8', titleFont: 'Inter', bodyFont: 'Inter', band: null },
  { id: 'noche', name: 'Noche', bg: '#151827', title: '#FFFFFF', text: '#C9CCDA', accent: '#FFB547', accent2: '#8E7FFF', titleFont: 'Inter', bodyFont: 'Inter', band: 'linear-gradient(135deg,#1F2337 0%,#151827 60%)' },
  { id: 'coral', name: 'Coral', bg: '#FFF4F1', title: '#7A1F33', text: '#5B3A42', accent: '#E8456B', accent2: '#FF8A5C', titleFont: 'Georgia', bodyFont: 'Inter', band: null },
  { id: 'jade', name: 'Jade', bg: '#F0FAF6', title: '#0A4A3A', text: '#2F4B44', accent: '#0E9C74', accent2: '#2FB4A8', titleFont: 'Inter', bodyFont: 'Inter', band: null },
  { id: 'indigo', name: 'Índigo', bg: '#3F2BC4', title: '#FFFFFF', text: '#E3DEFF', accent: '#FFD38A', accent2: '#8E7FFF', titleFont: 'Inter', bodyFont: 'Inter', band: 'linear-gradient(135deg,#5B48E8 0%,#3F2BC4 55%,#2A1C8F 100%)' },
  { id: 'papel', name: 'Papel', bg: '#F7F1E6', title: '#3B2A1A', text: '#5A4632', accent: '#B9481A', accent2: '#2F6F5E', titleFont: 'Georgia', bodyFont: 'Georgia', band: null },
  { id: 'cielo', name: 'Cielo', bg: '#EEF6FD', title: '#0D3B66', text: '#2E4A62', accent: '#1F8FD6', accent2: '#F08A24', titleFont: 'Inter', bodyFont: 'Inter', band: null },
  { id: 'carbon', name: 'Carbón', bg: '#22252B', title: '#F5F5F5', text: '#BFC3CC', accent: '#2ED3A0', accent2: '#FF6B8B', titleFont: 'Inter', bodyFont: 'Inter', band: null },
];
export const themeById = (id) => THEMES.find((t) => t.id === id) || THEMES[0];

export const LAYOUTS = [
  { id: 'title', name: 'Diapositiva de título' },
  { id: 'content', name: 'Título y contenido' },
  { id: 'two', name: 'Dos contenidos' },
  { id: 'section', name: 'Encabezado de sección' },
  { id: 'titleOnly', name: 'Solo el título' },
  { id: 'blank', name: 'En blanco' },
];

// Objetos de marcador de posición para cada diseño, en una diapositiva de W×H
export function layoutObjects(layout, W, H, uid) {
  const m = 60;
  const T = (o) => ({ id: uid(), type: 'text', rot: 0, ...o });
  switch (layout) {
    case 'title': return [
      T({ role: 'title', x: m, y: H * 0.30, w: W - 2 * m, h: 120, html: '', style: { size: 44, align: 'center', valign: 'bottom', bold: true } }),
      T({ role: 'subtitle', x: m, y: H * 0.30 + 136, w: W - 2 * m, h: 70, html: '', style: { size: 22, align: 'center', valign: 'top' } }),
    ];
    case 'content': return [
      T({ role: 'title', x: m, y: 36, w: W - 2 * m, h: 80, html: '', style: { size: 34, align: 'left', valign: 'middle', bold: true } }),
      T({ role: 'body', x: m, y: 132, w: W - 2 * m, h: H - 132 - 40, html: '', style: { size: 22, align: 'left', valign: 'top', list: true } }),
    ];
    case 'two': return [
      T({ role: 'title', x: m, y: 36, w: W - 2 * m, h: 80, html: '', style: { size: 34, align: 'left', valign: 'middle', bold: true } }),
      T({ role: 'body', x: m, y: 132, w: (W - 2 * m - 30) / 2, h: H - 172, html: '', style: { size: 20, align: 'left', valign: 'top', list: true } }),
      T({ role: 'body', x: m + (W - 2 * m - 30) / 2 + 30, y: 132, w: (W - 2 * m - 30) / 2, h: H - 172, html: '', style: { size: 20, align: 'left', valign: 'top', list: true } }),
    ];
    case 'section': return [
      T({ role: 'title', x: m, y: H * 0.38, w: W - 2 * m, h: 100, html: '', style: { size: 40, align: 'left', valign: 'bottom', bold: true } }),
      T({ role: 'subtitle', x: m, y: H * 0.38 + 110, w: W - 2 * m, h: 60, html: '', style: { size: 20, align: 'left', valign: 'top' } }),
    ];
    case 'titleOnly': return [T({ role: 'title', x: m, y: 36, w: W - 2 * m, h: 80, html: '', style: { size: 34, align: 'left', valign: 'middle', bold: true } })];
    default: return [];
  }
}
export const PLACEHOLDER = { title: 'Haz clic para agregar un título', subtitle: 'Haz clic para agregar un subtítulo', body: 'Haz clic para agregar texto' };
