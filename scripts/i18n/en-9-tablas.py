# Fase 5, parte 3 (Escribiente como LibreOffice Writer): tablas
D = {
'Seleccionar': 'Select', 'Seleccionar celda': 'Select cell', 'Seleccionar fila': 'Select row', 'Seleccionar columna': 'Select column', 'Seleccionar tabla': 'Select table',
'Insertar filas': 'Insert rows', 'Insertar columnas': 'Insert columns', 'Insertar filas…': 'Insert rows…', 'Insertar columnas…': 'Insert columns…', 'Cantidad': 'Number', 'Posición': 'Position', 'Antes': 'Before', 'Después': 'After',
'Combinar': 'Merge', 'Combinar celdas': 'Merge cells', 'Dividir celdas': 'Split cells', 'Dividir celdas…': 'Split cells…', 'Dividir tabla': 'Split table', 'Dividir en': 'Split into', 'Dirección': 'Direction',
'Verticalmente (columnas)': 'Vertically (columns)', 'Horizontalmente (filas)': 'Horizontally (rows)',
'Distribuir filas': 'Distribute rows', 'Distribuir columnas': 'Distribute columns', 'Ancho óptimo': 'Optimal width', 'Arriba': 'Top', 'Centro': 'Center', 'Abajo': 'Bottom',
'Estilos de tabla': 'Table styles', 'Predeterminado': 'Default', 'Rejilla': 'Grid', 'Sin bordes': 'No borders', 'Encabezado azul': 'Blue header', 'Verde': 'Green', 'Naranja': 'Orange', 'Gris': 'Gray', 'Elegante': 'Elegant', 'Filas alternas': 'Banded rows',
'Todos los bordes': 'All borders', 'Solo contorno': 'Outer border only', 'Solo bordes interiores': 'Inner borders only', 'Contorno y líneas horizontales': 'Outer and horizontal lines', 'Solo líneas horizontales': 'Horizontal lines only',
'Datos': 'Data', 'Ordenar': 'Sort', 'Ordenar…': 'Sort…', 'Ordenar por': 'Sort by', 'Después por': 'Then by', 'Tipo de clave': 'Key type', 'Automático': 'Automatic', 'Numérico': 'Numeric', 'Alfanumérico': 'Alphanumeric', 'Orden': 'Order', 'Ascendente': 'Ascending', 'Descendente': 'Descending',
'La primera fila es encabezado': 'First row is a header', 'Columna': 'Column',
'Fórmula': 'Formula', 'Fórmula…': 'Formula…', 'Fórmula no válida': 'Invalid formula', 'Actualizar fórmulas': 'Update formulas',
'Funciones: SUM, AVERAGE, MIN, MAX, COUNT, PRODUCT, ROUND, ABS. Referencias: A1, A1:B3, ABOVE (arriba), LEFT (izquierda), BELOW, RIGHT. Operadores: + − * / ^ %.': 'Functions: SUM, AVERAGE, MIN, MAX, COUNT, PRODUCT, ROUND, ABS. References: A1, A1:B3, ABOVE, LEFT, BELOW, RIGHT. Operators: + − * / ^ %.',
'Convertir en texto…': 'Convert to text…', 'Convertir tabla en texto': 'Convert table to text', 'Convertir texto en tabla': 'Convert text to table', 'Texto en tabla…': 'Text to table…', 'Convertir': 'Convert',
'Separar el texto en': 'Separate text at', 'Separar las celdas con': 'Separate cells with', 'Tabulaciones': 'Tabs', 'Punto y coma': 'Semicolons', 'Comas': 'Commas', 'Párrafos': 'Paragraphs', 'Párrafos (una columna)': 'Paragraphs (one column)', 'Otro carácter': 'Other character',
'Propiedades': 'Properties', 'Propiedades de tabla': 'Table properties', 'Propiedades de tabla…': 'Table properties…', 'Tabla': 'Table', 'Columnas': 'Columns', 'Fila': 'Row', 'Celda': 'Cell',
'Ancho (% de la página)': 'Width (% of page)', 'Sangría izquierda': 'Left indent', 'Encabezado': 'Heading', 'Repetir la primera fila como encabezado': 'Repeat first row as heading', 'No dividir la tabla entre páginas': 'Do not split table across pages',
'También puedes arrastrar los bordes de las columnas en el documento.': 'You can also drag the column borders in the document.',
'Alto mínimo de la fila': 'Minimum row height', 'No dividir la fila entre páginas': 'Do not split row across pages',
'Alineación vertical': 'Vertical alignment', 'Relleno (márgenes de celda)': 'Padding (cell margins)', 'Sin fondo': 'No fill',
'Disposición': 'Arrangement', '(sin cambios)': '(no change)', 'Aplicar a': 'Apply to', 'Toda la tabla': 'Whole table', 'Celdas seleccionadas': 'Selected cells',
'Coloca el cursor en una tabla.': 'Place the cursor in a table.', 'Coloca el cursor en la celda donde va el resultado.': 'Place the cursor in the cell for the result.', 'Coloca el cursor en una fila que no sea la primera.': 'Place the cursor in a row other than the first.',
'Selecciona los párrafos que quieres convertir.': 'Select the paragraphs you want to convert.', 'Selecciona varias celdas arrastrando el ratón (o con Mayús+clic).': 'Select several cells by dragging the mouse (or with Shift+click).',
'Las celdas seleccionadas deben formar un rectángulo.': 'The selected cells must form a rectangle.', 'No se puede dividir por una celda combinada verticalmente.': 'Cannot split at a vertically merged cell.', 'No se puede ordenar una tabla con celdas combinadas verticalmente.': 'Cannot sort a table with vertically merged cells.',
}
RULES = [
    [r'^Columna (\d+)( \(.*\))?$', 'Column $1$2'],
    [r'^Fórmula — ([A-Z]+\d+)$', 'Formula — $1'],
]
