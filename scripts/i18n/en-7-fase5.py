# Fase 5, parte 1 (Escribiente como LibreOffice Writer): Párrafo, Fuente, mayúsculas, marcas, lectores
D = {
'Párrafo…': 'Paragraph…', 'Fuente… (Ctrl+D)': 'Font… (Ctrl+D)', 'Marcas de formato (Ctrl+F10)': 'Formatting marks (Ctrl+F10)', 'Borrar formato': 'Clear formatting',
'Cambiar mayúsculas y minúsculas': 'Change case', 'Tipo oración.': 'Sentence case.', 'Poner En Mayúsculas Cada Palabra': 'Capitalize Each Word', 'aLTERNAR mAY/min': 'tOGGLE cASE',
'Selecciona el texto que quieres cambiar.': 'Select the text you want to change.', 'Selecciona texto en el documento.': 'Select text in the document.', 'Coloca el cursor en un párrafo.': 'Place the cursor in a paragraph.',
'Sangría y espaciado': 'Indents & spacing', 'Flujo del texto': 'Text flow', 'Bordes y sombreado': 'Borders & shading',
'Sangría': 'Indent', 'Antes del texto': 'Before text', 'Después del texto': 'After text', 'Especial': 'Special', '(ninguna)': '(none)', 'Primera línea': 'First line', 'Francesa': 'Hanging', 'En': 'By',
'Espaciado': 'Spacing', 'Encima del párrafo': 'Above paragraph', 'Debajo del párrafo': 'Below paragraph', 'Interlineado': 'Line spacing', 'De': 'Of',
'Sencillo': 'Single', '1.15 líneas': '1.15 lines', '1.5 líneas': '1.5 lines', 'Doble': 'Double', 'Múltiple': 'Multiple', 'Al menos': 'At least', 'Exacto': 'Exactly',
'Alineación': 'Alignment', 'Izquierda': 'Left', 'Centrado': 'Centered', 'Derecha': 'Right', 'Justificado': 'Justified',
'Saltos': 'Breaks', 'Salto de página antes del párrafo': 'Page break before the paragraph', 'Opciones': 'Options', 'Conservar con el párrafo siguiente': 'Keep with next paragraph', 'No dividir el párrafo': 'Do not split paragraph', 'Control de viudas y huérfanas': 'Widow and orphan control',
'Bordes': 'Borders', 'Lados con borde': 'Bordered sides', 'Sin borde': 'No border', 'Cuadro': 'Box', 'Solo arriba': 'Top only', 'Solo abajo': 'Bottom only', 'Arriba y abajo': 'Top and bottom', 'Solo a la izquierda': 'Left only',
'Estilo': 'Style', 'Sólido': 'Solid', 'Punteado': 'Dotted', 'Discontinuo': 'Dashed', 'Ancho': 'Width', 'Color': 'Color', 'Espacio al texto': 'Spacing to text', 'Sombreado': 'Shading', 'Color de fondo': 'Background color',
'Familia': 'Family', 'Normal': 'Normal', 'Negrita': 'Bold', 'Cursiva': 'Italic', 'Negrita cursiva': 'Bold italic', 'Tamaño': 'Size', 'Color de fuente': 'Font color', 'Subrayado': 'Underline', '(ninguno)': '(none)', 'Ondulado': 'Wave',
'Efectos tipográficos': 'Font effects', 'Tachado': 'Strikethrough', 'Doble tachado': 'Double strikethrough', 'Superíndice': 'Superscript', 'Subíndice': 'Subscript', 'Versalitas': 'Small capitals', 'Oculto': 'Hidden', 'Sombra': 'Shadow', 'Contorno': 'Outline',
'Posición y espaciado': 'Position & spacing', 'Espaciado entre caracteres': 'Character spacing', 'Escala de ancho': 'Width scale', 'Posición (elevar o bajar)': 'Position (raise or lower)',
'No es un documento de Word válido': 'Not a valid Word document', 'gráficos': 'charts', 'formas o diagramas': 'shapes or diagrams', 'objetos incrustados': 'embedded objects',
}
RULES = [
    [r'^Se abrió con avisos: No se pudieron conservar del todo: (.+)$', 'Opened with warnings: could not be fully preserved: $1'],
    [r'^No se pudieron conservar del todo: (.+)$', 'Could not be fully preserved: $1'],
]
