# Fase 5, parte 2 (Escribiente como LibreOffice Writer): estilos, listas multinivel, numeración de capítulos
D = {
'Carácter': 'Character', 'Título 5': 'Heading 5', 'Título 6': 'Heading 6',
'Panel de estilos': 'Styles pane', 'Panel de estilos (Ctrl+F11)': 'Styles pane (Ctrl+F11)',
'Nuevo a partir de la selección': 'New from selection', 'Nuevo estilo a partir de la selección': 'New style from selection', 'Nuevo estilo a partir de la selección…': 'New style from selection…',
'Actualizar con la selección': 'Update to match selection', 'Actualizar estilo': 'Update style', 'Modificar…': 'Modify…', 'Buscar estilo…': 'Search style…', 'Doble clic para aplicar': 'Double-click to apply',
'Nombre del estilo': 'Style name', 'Ya existe un estilo llamado {0}.': 'A style named {0} already exists.', 'Estilo {0} creado': 'Style {0} created', 'Estilo {0} actualizado': 'Style {0} updated', 'Estilo {0} eliminado': 'Style {0} deleted',
'Modificar estilo: {0}': 'Modify style: {0}', 'Organizador': 'Organizer', 'Heredado de': 'Inherit from', 'Estilo siguiente': 'Next style',
'Estilo predeterminado: se puede modificar, no cambiar de nombre ni eliminar.': 'Built-in style: it can be modified, not renamed or deleted.',
'Antes del texto (cm)': 'Before text (cm)', 'Después del texto (cm)': 'After text (cm)', 'Primera línea (cm)': 'First line (cm)', 'Interlineado (múltiplo)': 'Line spacing (multiple)',
'Espaciado entre caracteres (pt)': 'Character spacing (pt)', 'Color del borde': 'Border color',
'Énfasis': 'Emphasis', 'Énfasis fuerte': 'Strong Emphasis', 'Texto de código': 'Source Text', 'Cita textual': 'Quotation', '(Quitar estilo de carácter)': '(Remove character style)',
'Viñeta': 'Bullet', 'Reiniciar numeración': 'Restart numbering', 'Continuar numeración': 'Continue numbering',
'Establecer valor de numeración': 'Set numbering value', 'Establecer valor de numeración…': 'Set numbering value…', 'Coloca el cursor en una lista numerada.': 'Place the cursor in a numbered list.',
'Numeración de capítulos': 'Chapter numbering', 'Capítulo 1 / 1.1 / 1.1.1': 'Chapter 1 / 1.1 / 1.1.1', 'Capítulo ': 'Chapter ',
}
RULES = [
    [r'^(.) {2}Viñeta$', '$1  Bullet'],
]
