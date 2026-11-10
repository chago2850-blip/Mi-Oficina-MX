# Contenido de las plantillas (solo se usa al crear un documento nuevo con la interfaz en inglés)
D = {
# Escribiente: carta
'Lic. Nombre Apellido': 'Mr./Ms. First Last', 'Puesto': 'Position', 'Empresa o institución': 'Company or institution', 'P r e s e n t e': '',
'Estimado(a) Lic. Apellido:': 'Dear Mr./Ms. Last:',
'Por medio de la presente me dirijo a usted para [motivo de la carta]. Escribe aquí el primer párrafo, donde explicas de forma breve y clara el propósito.': 'I am writing to you regarding [purpose of the letter]. Write the first paragraph here, briefly and clearly explaining the purpose.',
'En el segundo párrafo puedes dar los detalles, los antecedentes o la información de apoyo que la persona necesita para responder.': 'In the second paragraph you can give the details, background or supporting information the reader needs to respond.',
'Agradezco de antemano su atención y quedo a sus órdenes para cualquier aclaración.': 'Thank you in advance for your attention. I remain at your disposal for any questions.',
'A t e n t a m e n t e': 'Sincerely', 'Tu cargo · tu teléfono · tu correo': 'Your title · your phone · your email',
# Escribiente: reporte
'Título del reporte': 'Report title', 'Resumen': 'Summary', 'Escribe en dos o tres frases lo más importante: qué se hizo, qué se encontró y qué se recomienda.': 'In two or three sentences, write what matters most: what was done, what was found and what is recommended.',
'Antecedentes': 'Background', 'Explica el contexto y por qué se hizo este trabajo.': 'Explain the context and why this work was done.', 'Resultados': 'Results', 'Hallazgo principal': 'Main finding',
'Describe el resultado más importante.': 'Describe the most important result.', 'Indicador': 'Metric', 'Clientes': 'Customers', 'Otros hallazgos': 'Other findings',
'Primer punto.': 'First point.', 'Segundo punto.': 'Second point.', 'Tercer punto.': 'Third point.', 'Recomendaciones': 'Recommendations', 'Primera recomendación.': 'First recommendation.', 'Segunda recomendación.': 'Second recommendation.',
# Escribiente: currículum
'Nombre Apellido': 'First Last', 'Profesión · Ciudad · 744 000 0000 · correo@ejemplo.com': 'Profession · City · 555 000 0000 · email@example.com', 'Perfil': 'Profile',
'Dos o tres líneas sobre quién eres, en qué eres bueno y qué buscas.': 'Two or three lines about who you are, what you are good at and what you are looking for.',
'Experiencia': 'Experience', 'Puesto — Empresa': 'Position — Company', 'Enero 2023 – actual': 'January 2023 – present', 'Logro medible número uno.': 'Measurable achievement number one.',
'Logro medible número dos.': 'Measurable achievement number two.', 'Puesto anterior — Empresa': 'Previous position — Company', 'Responsabilidad o logro.': 'Responsibility or achievement.',
'Educación': 'Education', 'Licenciatura en …': "Bachelor's degree in …", '— Universidad, 2019': '— University, 2019', 'Habilidades': 'Skills', 'Habilidad · Habilidad · Habilidad · Idioma (nivel)': 'Skill · Skill · Skill · Language (level)',
# Calcula Datos
'Presupuesto mensual': 'Monthly budget', 'Concepto': 'Item', 'Presupuesto': 'Budget', 'Real': 'Actual', 'Diferencia': 'Difference', 'Renta': 'Rent', 'Luz y agua': 'Utilities',
'Internet y teléfono': 'Internet & phone', 'Despensa': 'Groceries', 'Transporte': 'Transport', 'Ahorro': 'Savings', 'Mes': 'Month', 'Costos': 'Costs', 'Utilidad': 'Profit', 'Margen': 'Margin',
'Abril': 'April', 'Asistencias': 'Attendance', 'Lista de asistencia': 'Attendance list',
# Impresiona
'Plan de negocio 2027': 'Business plan 2027', 'Tu empresa · Acapulco, Guerrero': 'Your company · City, State', 'El problema': 'The problem',
'Qué necesidad no está resuelta': 'What need is not being met', 'A quién le afecta y cuánto le cuesta': 'Who it affects and what it costs them', 'Por qué ahora es el momento': 'Why now is the time',
'Nuestra propuesta': 'Our proposal', 'Producto o servicio': 'Product or service', 'Lo que nos hace diferentes': 'What makes us different', 'Precio y modelo de ingresos': 'Price and revenue model',
'Canales de venta': 'Sales channels', 'Gracias': 'Thank you', 'Contacto: tu correo · tu teléfono': 'Contact: your email · your phone',
'El ciclo del agua': 'The water cycle', 'Ciencias Naturales · 5.º grado': 'Science · 5th grade', '¿Qué vamos a aprender?': 'What will we learn?',
'Las etapas del ciclo del agua': 'The stages of the water cycle', 'Por qué es importante para la vida': 'Why it matters for life', 'Cómo cuidar el agua en casa': 'How to save water at home',
'Etapas': 'Stages', 'Evaporación': 'Evaporation', 'Condensación': 'Condensation', 'Precipitación': 'Precipitation', 'Escurrimiento e infiltración': 'Runoff and infiltration', 'Actividad en equipo': 'Team activity',
'Avance del proyecto': 'Project update', 'Reunión semanal': 'Weekly meeting', 'Lo que logramos': 'What we achieved', 'Entregable 1 terminado': 'Deliverable 1 completed',
'Pruebas con 10 usuarios': 'Tests with 10 users', 'Presupuesto dentro de lo planeado': 'Budget on track', 'Siguientes pasos': 'Next steps',
'Corregir los hallazgos de las pruebas': 'Fix the issues found in testing', 'Preparar el lanzamiento': 'Prepare the launch', 'Definir responsables y fechas': 'Assign owners and dates',
}
RULES = [
    [r'^Acapulco, Guerrero, a (.*)$', 'City, State, $1'],
    [r'^Subtítulo o área · (.*)$', 'Subtitle or department · $1'],
]
