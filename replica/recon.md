# Recon: Microsoft Office (Word, Excel, PowerPoint) → suite libre "Mi Oficina MX"

Fecha: 2026-11-09 · Método: clean-room, solo documentación pública de Microsoft
y fichas de licencia de bibliotecas libres. No se ejecutó, descompiló ni
inspeccionó ningún binario de Microsoft Office. Sin capturas: `replica/screens/`
queda vacío; la estructura de pantallas sale de los artículos de soporte.

## 1. Alcance

| pregunta | respuesta |
| --- | --- |
| App / plataforma | Microsoft Word, Excel y PowerPoint de escritorio. Clon: escritorio **Windows, macOS y Linux** |
| Licencia del clon | **GPL-3.0-or-later** para toda la suite |
| Componentes | **Escribiente** (procesador de textos), **Calcula Datos** (hoja de cálculo), **Impresiona** (presentaciones), **PapelPDF** (editor PDF, repo existente v1.2.0) + **Launcher dona** |
| Rebanada (v1.0) | El ciclo central de cada app: **crear/abrir → editar con formato → guardar en el formato de Office (.docx/.xlsx/.pptx) y en ODF → exportar PDF/imprimir**. Colaboración en la nube, macros VBA y Copilot quedan fuera |
| Para quién | Personas, escuelas y negocios pequeños en México que hoy pagan Microsoft 365 o usan copias sin licencia; firma: Santiago Miguel |

**Nota honesta sobre "las mismas funciones que Office":** Office tiene más de
30 años de desarrollo y miles de funciones. LibreOffice, con cientos de
colaboradores, todavía no alcanza paridad completa. Este mapa separa lo que
entra en la v1.0 (must), lo que sigue después (should/could) y lo que no se
puede clonar (skip). Ver §9.

## 2. Fuentes

| # | fuente | URL | qué dio |
| --- | --- | --- | --- |
| 1 | Funciones de Excel por categoría (Microsoft Support) | https://support.microsoft.com/en-au/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb | Categorías y conteo de funciones; funciones nuevas (FILTER, XLOOKUP, LET, UNIQUE, SORT, LAMBDA, MAP, VSTACK…) |
| 2 | Novedades de Office 2024 / LTSC 2024 | https://support.microsoft.com/topic/baadecbe-a8fc-4435-8cb8-70e93fd08b56 | Lo reciente por app: ODF 1.4, recuperación automática de todos los documentos, reacciones en comentarios, gráficos sobre matrices dinámicas, función IMAGE, Cameo, Recording Studio, subtítulos |
| 3 | Novedades de Word 2024 | https://support.microsoft.com/en-gb/office/what-s-new-in-word-2024-for-windows-and-mac-9fe36ecd-820b-4f42-a9a2-0861f7615954 | Confirmación de lo anterior para Word |
| 4 | Transición Transformación (Morph) | https://support.microsoft.com/en-us/office/8dd1c7b2-b935-44f5-a74c-741d8d9244ea | Cómo funciona la transición estrella de PowerPoint |
| 5 | Diferencia entre animaciones y transiciones | https://support.microsoft.com/en-us/office/the-difference-between-animations-and-transitions-bae174a4-dad3-4268-bf9c-c201a70995f1 | Modelo de transiciones (entre diapositivas) vs animaciones (por objeto) |
| 6 | Cinta de Word (Inicio, Insertar, Diseño, Disposición, Referencias, Correspondencia, Revisar, Vista) | https://fluent.apryse.com/documentation/designer-guide/Reference/OfficeSupport/Word/WordFeatureSupport | Lista de funciones de Word por pestaña; referencia de qué partes del formato .docx importan |
| 7 | Términos MS Office vs LibreOffice | https://help.libreoffice.org/latest/en-US/text/shared/guide/microsoft_terms.html | Equivalencias de nombres (útil para traducir funciones a español neutro sin copiar textos) |
| 8 | HyperFormula (licencia) | https://cdn.jsdelivr.net/npm/hyperformula@3.2.0/README.md | Motor de fórmulas: **GPLv3** o comercial, ~400 funciones, expresiones con nombre, deshacer/rehacer |
| 9 | Univer (GitHub) | https://github.com/dream-num/univer | Apache-2.0, pero **importar/exportar .xlsx/.docx, gráficos, tablas dinámicas e impresión son de pago (Pro)** → descartado como base |
| 10 | FortuneSheet (GitHub) | https://www.github.com/ruilisi/fortune-sheet | MIT; fórmulas, formato condicional, inmovilizar, combinar; .xlsx vía plugin; **gráficos y tablas dinámicas sin terminar**; API inestable antes de 1.0 |
| 11 | Tiptap Pages | https://tiptap.dev/docs/pages/getting-started/overview.md | **Paginación y exportar DOCX de Tiptap son de pago (Pro)** → la paginación hay que escribirla nosotros |
| 12 | OnlyOffice (Wikipedia) | https://en.wikipedia.org/wiki/OnlyOffice | AGPL-3.0; editores en C++/JS con historial de cláusulas adicionales de marca → no se usa como base |
| 13 | pptxtojson | https://github.imc.re/pipipi-pikachu/pptxtojson | MIT; lee texto, imágenes, formas, tablas, gráficos, video/audio y SmartArt de un .pptx |
| 14 | PapelPDF (repo propio) | https://github.com/chago2850-blip/PapelPDF | Electron 33, pdf.js, pdf-lib, qpdf-wasm, tesseract.js; CI que ya produce .exe/.dmg/.deb/.rpm/.AppImage; **licencia MIT** |

## 3. Ciclo central

> Abrir o crear un documento, darle formato profesional sin pensar en el
> archivo, y entregarlo en un formato que el destinatario abra en Office o
> en PDF.

## 4. Inventario de pantallas

### Suite y Launcher

| ID | pantalla | cómo se llega | propósito | componentes clave | estados |
| --- | --- | --- | --- | --- | --- |
| S01 | Launcher cerrado | abrir "Mi Oficina MX" | punto de entrada de la suite | ventana sin marco, transparente, dona con botón central **Inicio** | reposo, hover (brillo), arrastrando la ventana |
| S02 | Launcher abierto | clic en Inicio | elegir app | dona encendida dividida en 4 sectores (Escribiente, Calcula Datos, Impresiona, PapelPDF) con logo y nombre; centro cambia a "Cerrar" | animación de apertura, sector en hover, sector con foco de teclado, app ya abierta (punto indicador) |
| S03 | Launcher minimizado | se abre una app | apartarse mientras se trabaja | ícono en bandeja/dock; restaura sola al cerrar la última app | minimizado, restaurando |
| S04 | Menú del Launcher | clic derecho en la dona o en la bandeja | Acerca de, Preferencias, Salir | menú contextual | — |
| S05 | Acerca de / Licencia | S04 → Acerca de | versión, GPL-3.0, autor y contacto, créditos de bibliotecas | diálogo, lista de licencias de terceros | — |
| S06 | Preferencias de la suite | S04 → Preferencias | idioma, tema claro/oscuro, autoguardado, carpeta por defecto, formato por defecto (Office u ODF) | formulario | — |

### Compartidas por las tres apps nuevas (backstage)

| ID | pantalla | cómo se llega | propósito | componentes clave | estados |
| --- | --- | --- | --- | --- | --- |
| S07 | Inicio de la app | abrir la app sin archivo | nuevo en blanco, plantillas, recientes, abrir | tarjetas de plantilla, lista de recientes, zona de arrastre | sin recientes, con recientes, archivo reciente ya no existe |
| S08 | Archivo (backstage) | pestaña Archivo | Nuevo, Abrir, Guardar, Guardar como, Exportar PDF, Imprimir, Propiedades, Cerrar | panel lateral de pantalla completa | — |
| S09 | Imprimir con vista previa | Archivo → Imprimir / Ctrl+P | elegir impresora, páginas, copias, orientación | vista previa paginada, selector de impresora | sin impresoras, imprimiendo, error |
| S10 | Recuperación de documentos | abrir tras un cierre inesperado | recuperar lo no guardado | panel con versiones recuperadas | hay recuperados, vacío |
| S11 | Buscar y reemplazar | Ctrl+F / Ctrl+H | buscar, reemplazar, opciones (mayúsculas, palabra completa, sin acentos, regex) | barra o diálogo | sin resultados, n de m |
| S12 | Diálogo de color / fuente | desde la cinta | elegir fuente, tamaño, color | selector de color con temas y recientes | — |

### Escribiente (procesador de textos, equivalente a Word)

| ID | pantalla | cómo se llega | propósito | componentes clave | estados |
| --- | --- | --- | --- | --- | --- |
| S20 | Editor | abrir/crear documento | escribir y dar formato | cinta (Archivo, Inicio, Insertar, Diseño, Disposición, Referencias, Correspondencia, Revisar, Vista), regla, páginas, barra de estado (páginas, palabras, idioma, zoom) | documento vacío, cargando, solo lectura, modo compatibilidad (.doc), modo enfoque |
| S21 | Panel de estilos | Inicio → Estilos | aplicar y modificar estilos | galería de estilos, diálogo de modificar estilo | — |
| S22 | Configurar página | Disposición → Márgenes/Tamaño/Orientación/Columnas | página, márgenes, columnas, saltos de sección | diálogo con pestañas | — |
| S23 | Encabezado y pie | doble clic en el margen | encabezados, pies, numeración de página | zona editable, pestaña contextual | primera página distinta, pares/impares |
| S24 | Tabla | Insertar → Tabla | insertar y editar tablas | cuadrícula de inserción, pestañas contextuales Diseño de tabla / Disposición | celda seleccionada, combinadas |
| S25 | Imagen y forma | Insertar → Imágenes/Formas | insertar y ajustar | controles de tamaño, ajuste de texto, pestaña Formato de imagen | seleccionada, recortando |
| S26 | Tabla de contenido / notas / citas | Referencias | índice automático, notas al pie, bibliografía, títulos | diálogos | índice desactualizado |
| S27 | Combinar correspondencia | Correspondencia | cartas/etiquetas desde CSV/XLSX | asistente por pasos, vista previa de registros | sin datos, campos sin asignar |
| S28 | Revisión | Revisar | ortografía, comentarios, control de cambios, comparar | panel de comentarios, globos, barra de aceptar/rechazar | sin cambios, con cambios |
| S29 | Navegación | Vista → Panel de navegación | títulos, páginas, resultados | árbol de títulos, miniaturas | documento sin títulos |
| S30 | Contar palabras / estadísticas | clic en la barra de estado | palabras, caracteres, párrafos | diálogo | — |

### Calcula Datos (hoja de cálculo, equivalente a Excel)

| ID | pantalla | cómo se llega | propósito | componentes clave | estados |
| --- | --- | --- | --- | --- | --- |
| S40 | Cuadrícula | abrir/crear libro | capturar y calcular | cinta (Archivo, Inicio, Insertar, Disposición de página, Fórmulas, Datos, Revisar, Vista), cuadro de nombres, barra de fórmulas, cuadrícula, pestañas de hojas, barra de estado (suma, promedio, recuento) | vacío, editando celda, rango seleccionado, error de fórmula (#¡DIV/0!, #N/D, #¡REF!…), referencia circular, cálculo manual |
| S41 | Formato de celdas | Ctrl+1 | número, alineación, fuente, borde, relleno, protección | diálogo con pestañas | — |
| S42 | Insertar función | fx | buscar y armar funciones | buscador por categoría, argumentos con ayuda | — |
| S43 | Ordenar y filtrar | Datos → Ordenar/Filtro | ordenar por varias columnas, autofiltro | menús desplegables en encabezados | filtro activo |
| S44 | Formato condicional | Inicio → Formato condicional | reglas, barras de datos, escalas de color, íconos | administrador de reglas | sin reglas |
| S45 | Gráfico | Insertar → Gráfico | crear y editar gráficos | selector de tipo, panel de formato, pestaña contextual | sin datos, seleccionado |
| S46 | Tabla dinámica | Insertar → Tabla dinámica | resumir datos | panel de campos (Filtros, Columnas, Filas, Valores) | origen inválido, actualizada |
| S47 | Validación de datos | Datos → Validación | listas, rangos, mensajes | diálogo | — |
| S48 | Administrador de nombres | Fórmulas → Administrador de nombres | rangos con nombre | tabla | — |
| S49 | Configurar página / área de impresión | Disposición de página | márgenes, ajustar a página, títulos que se repiten, vista de salto de página | diálogo, vista especial | — |
| S50 | Importar CSV / texto | Datos → Desde texto | delimitador, codificación, tipos | vista previa | codificación errónea |

### Impresiona (presentaciones, equivalente a PowerPoint)

| ID | pantalla | cómo se llega | propósito | componentes clave | estados |
| --- | --- | --- | --- | --- | --- |
| S60 | Vista Normal | abrir/crear presentación | editar diapositivas | cinta (Archivo, Inicio, Insertar, Diseño, Transiciones, Animaciones, Presentación, Revisar, Vista), panel de miniaturas, lienzo, panel de notas | presentación vacía, objeto seleccionado, varios seleccionados, guías inteligentes visibles |
| S61 | Clasificador de diapositivas | Vista → Clasificador | reordenar, secciones | cuadrícula de miniaturas | — |
| S62 | Patrón de diapositivas | Vista → Patrón | editar diseños y tema | árbol de diseños, marcadores de posición | — |
| S63 | Galería de temas y diseños | Diseño | tema, variantes, tamaño de diapositiva, fondo | galería | — |
| S64 | Panel de animaciones | Animaciones → Panel | orden, inicio, duración, retraso | lista con línea de tiempo | sin animaciones |
| S65 | Transiciones | Transiciones | elegir transición entre diapositivas, incluida Transformación | galería, opciones de efecto, duración | — |
| S66 | Presentación a pantalla completa | F5 / Mayús+F5 | exponer | diapositiva, controles al pasar el mouse, puntero láser, pluma | primera, última, pantalla negra/blanca |
| S67 | Vista del moderador | presentación con 2 pantallas | exponer con notas | diapositiva actual, siguiente, notas, temporizador, reloj | una sola pantalla |
| S68 | Insertar tabla / gráfico / SmartArt simple | Insertar | contenido estructurado | diálogos, edición de datos en Calcula Datos embebido | — |
| S69 | Exportar | Archivo → Exportar | PDF, imágenes PNG/JPG, video (después) | diálogo | exportando |

### PapelPDF (existente)

| ID | pantalla | nota |
| --- | --- | --- |
| S80 | PapelPDF completo | Ya construido (v1.2.0). Sus pantallas S01–S09 están en `PapelPDF/replica/recon.md`. Aquí solo se integra: mismo instalador, se lanza desde la dona, comparte tema e ícono de familia |

## 5. Flujos

```
F01 Abrir una app desde la suite
    S01 -> clic Inicio -> S02 -> clic sector "Escribiente" -> S03 (dona se minimiza) + S07
    clics en el camino feliz: 2
    borde: la app ya está abierta (enfocarla, no abrir otra), arranque lento (mostrar progreso en el sector),
           error al lanzar (mensaje y la dona vuelve a S02)

F02 La dona vuelve al cerrar la app
    S20/S40/S60/S80 -> cerrar la última ventana de la suite -> S01 se restaura y se enfoca
    borde: quedan otras apps abiertas (no restaurar), cambios sin guardar (preguntar antes),
           el usuario cerró la dona a propósito (respetarlo)

F03 Escribir una carta y enviarla en PDF
    S07 -> plantilla "Carta" -> S20 escribir, negritas, estilo Título -> S08 Exportar PDF -> guardar
    clics en el camino feliz: ~6
    borde: fuente no instalada, imagen muy grande, nombre de archivo con acentos

F04 Abrir un .docx recibido, corregir y devolverlo como .docx
    doble clic en el archivo (asociación) -> S20 -> S28 control de cambios -> guardar (Ctrl+S, mismo formato)
    borde: elementos no soportados (avisar sin perderlos: conservar XML desconocido), archivo protegido,
           .doc binario antiguo (abrir en modo compatibilidad, guardar como .docx)

F05 Hacer un reporte con índice
    S20 -> estilos Título 1/2 -> S26 insertar tabla de contenido -> S23 numerar páginas -> S09 imprimir
    borde: actualizar índice tras cambios, primera página sin número

F06 Combinar correspondencia
    S20 -> S27 elegir CSV/XLSX -> insertar campos -> vista previa -> generar documento o PDFs
    borde: columnas con acentos, registros vacíos, más de 1,000 registros

F07 Capturar ventas y sacar totales
    S07 -> libro en blanco -> S40 capturar -> =SUMA(B2:B30) -> formato moneda (S41) -> guardar .xlsx
    clics en el camino feliz: ~5
    borde: separador decimal coma/punto según idioma, fechas dd/mm/aaaa, nombres de funciones en español
           (SUMA ↔ SUM: guardar siempre en inglés en el archivo, mostrar en español)

F08 Abrir un .xlsx con fórmulas y gráficos
    S40 -> recalcular -> editar -> guardar
    borde: función no soportada (conservar fórmula y último valor calculado, marcar celda),
           macros .xlsm (abrir sin ejecutar, avisar), vínculos externos

F09 Resumir con tabla dinámica y graficar
    S40 -> seleccionar datos -> S46 -> arrastrar campos -> S45 gráfico de la tabla
    borde: datos sin encabezados, celdas combinadas en el origen

F10 Armar una presentación y exponer
    S07 -> tema -> S60 añadir diapositivas, texto, imágenes -> S65 transición -> S64 animación -> F5 -> S66/S67
    clics en el camino feliz: ~10
    borde: un solo monitor, video que no reproduce, fuentes faltantes en otra computadora

F11 Abrir un .pptx recibido y exportar a PDF
    S60 -> S69 PDF
    borde: SmartArt complejo (mostrar como imagen de respaldo si existe), animaciones no soportadas

F12 Abrir un PDF desde la suite
    S02 -> sector PapelPDF -> S80
```

## 6. Componentes

| componente | variantes | estados | se usa en |
| --- | --- | --- | --- |
| Dona del Launcher | cerrada, abierta (4 sectores) | reposo, hover, foco, animando, app abierta | S01–S03 |
| Cinta (ribbon) | pestañas fijas + contextuales; modo completo, contraído, simplificado | pestaña activa, grupo colapsado al estrechar la ventana | S20, S40, S60 |
| Botón de cinta | grande, pequeño, con menú (split), alternable | normal, hover, presionado, activo, deshabilitado | todas |
| Galería | en línea, desplegable | elemento en hover con vista previa en vivo | estilos, temas, transiciones, formato condicional |
| Selector de fuente / tamaño | combo con búsqueda | fuente faltante (cursiva + aviso) | S20, S40, S60 |
| Selector de color | tema, estándar, recientes, personalizado | — | todas |
| Barra de herramientas de acceso rápido | — | personalizada | todas |
| Backstage | panel completo | — | S08 |
| Diálogo modal | simple, con pestañas, asistente por pasos | — | S22, S27, S41… |
| Panel lateral | estilos, navegación, comentarios, animaciones, formato de objeto | abierto, cerrado, acoplado | varias |
| Regla | horizontal, vertical | tabulaciones, sangrías arrastrables | S20, S60 |
| Página (lienzo paginado) | vertical/horizontal, columnas | cursor, selección, saltos | S20 |
| Cuadrícula de celdas | — | celda activa, rango, copia con marco animado, relleno con controlador, inmovilizar | S40 |
| Barra de fórmulas + cuadro de nombres | — | editando, error | S40 |
| Pestañas de hojas | — | renombrando, color, oculta | S40 |
| Miniaturas de diapositivas | — | seleccionada, arrastrando, sección | S60, S61 |
| Objeto en lienzo (forma, imagen, cuadro de texto) | — | seleccionado, rotando, recortando, agrupado; guías inteligentes | S25, S60 |
| Barra de estado | por app | — | todas |
| Aviso / toast | info, éxito, error | — | todas |
| Panel de comentarios | hilo, respuesta, resuelto | — | S28 y equivalentes |

## 7. Modelo de datos inferido

No hay base de datos: los documentos son archivos. El "modelo" es el formato
interno en memoria y su correspondencia con OOXML (ECMA-376 / ISO 29500) y ODF.

```
Documento (Escribiente)   secciones[], estilos{}, numeraciones{}, encabezados/pies{}, notas[], comentarios[],
                          cambios_registrados[], propiedades (autor, título), partes_desconocidas (XML a conservar)
                          Sección: tamaño, márgenes, orientación, columnas, bloques[]
                          Bloque: Párrafo(estilo, runs[]) | Tabla(filas[celdas[bloques[]]]) | SaltoDePágina | ÍndiceTOC
                          Run: texto, marcas (negrita, cursiva, subrayado, color, fuente, tamaño, vínculo), objetos en línea
            evidencia: cinta de Word (fuente 6), estructura de .docx (ECMA-376)
            confianza: alta

Libro (Calcula Datos)     hojas[], nombres_definidos{}, estilos_celda[], tablas[], tablas_dinámicas[], gráficos[]
                          Hoja: nombre, celdas{ref -> Celda}, filas/columnas (alto, ancho, ocultas), combinadas[],
                                inmovilizado, formato_condicional[], validaciones[], filtro, área_impresión
                          Celda: valor (número|texto|booleano|error|fecha), fórmula (siempre en inglés en el archivo),
                                 estilo_id, comentario
            evidencia: fuente 1 (categorías de funciones), estructura de .xlsx
            confianza: alta

Presentación (Impresiona) tamaño_diapositiva, tema (colores, fuentes), patrones[] -> diseños[], diapositivas[]
                          Diapositiva: diseño_id, objetos[], notas, transición (tipo, duración, opciones), animaciones[]
                          Objeto: Forma | CuadroTexto | Imagen | Tabla | Gráfico | Video | Grupo; posición, tamaño,
                                  rotación, relleno, línea, texto enriquecido
                          Animación: objeto_id, tipo (entrada|énfasis|salida|trayectoria), inicio (clic|con|después),
                                     duración, retraso
            evidencia: fuentes 4 y 5, estructura de .pptx; pptxtojson (fuente 13) confirma los tipos de objeto
            confianza: alta

Ajustes de la suite       idioma, tema, autoguardado_min, carpeta_defecto, formato_defecto, recientes[] por app,
                          posición_dona
            evidencia: S06; supuesto de diseño
            confianza: guess
```

Relaciones: Libro 1-n Hoja 1-n Celda; Presentación 1-n Patrón 1-n Diseño 1-n
Diapositiva 1-n Objeto; Impresiona y Escribiente embeben gráficos con datos de
Calcula Datos (mismo motor de gráficos).

## 8. Matriz de funciones

Ver `features.csv`. 145 filas. Must: 68, should: 42, could: 25, skip: 10.
Por área: Calcula Datos 42, Escribiente 34, Impresiona 33, compartido 22, suite 14.

## 9. Fuera de alcance (no se puede o no se debe clonar)

- **Marcas, nombres e íconos de Microsoft** (Word, Excel, PowerPoint, Office, Microsoft 365, los íconos de letra en cuadro). Los nombres propios ya están definidos y los logos serán originales.
- **Textos de ayuda, plantillas, imágenes prediseñadas e íconos de la cinta de Office**: se escriben y dibujan desde cero.
- **Fuentes con licencia de Microsoft** (Calibri, Aptos, Cambria, Segoe UI): se usan sustitutas libres métricamente compatibles (Carlito ↔ Calibri, Caladea ↔ Cambria, Liberation Sans/Serif ↔ Arial/Times New Roman).
- **OneDrive, coautoría en tiempo real, Microsoft Stream, Teams**: servicios en la nube de Microsoft.
- **Copilot / Designer / ideas con IA**: servicios de pago en la nube.
- **Macros VBA**: ejecutar VBA implica reimplementar un lenguaje y un modelo de objetos enormes y es un riesgo de seguridad. Se abren .docm/.xlsm/.pptm sin ejecutar y conservando las macros.
- **Biblioteca de imágenes de archivo e íconos de Office (stock)**: contenido con licencia.
- **Funciones de cubo (CUBEVALUE…) y conectores Power Query/Power Pivot**: dependen de servidores OLAP y del motor de Microsoft.
- **Cameo (cámara en vivo) y Recording Studio**: posibles después, no en v1.0.

### Base técnica: qué NO usar y por qué

| opción | por qué no |
| --- | --- |
| Univer | Apache-2.0, pero importar/exportar .xlsx/.docx, gráficos, tablas dinámicas e impresión son de pago |
| Tiptap Pages | Paginación y exportar DOCX de pago |
| OnlyOffice | AGPL, C++ pesado, historial de cláusulas de marca; difícil de reempaquetar como suite propia |
| PPTist | Revisar licencia antes de reusar código (se usa solo pptxtojson, MIT) |

### Base técnica sugerida (para /replica-architect)

| pieza | biblioteca | licencia | compatible con GPL-3.0 |
| --- | --- | --- | --- |
| Contenedor de escritorio | Electron (igual que PapelPDF) | MIT | sí |
| Motor de fórmulas | HyperFormula | GPLv3 | sí (por eso la suite es GPL-3.0) |
| Editor de texto enriquecido | ProseMirror (+ paginación propia) | MIT | sí |
| Leer .docx | mammoth (texto/estilos) + lector OOXML propio | BSD-2 | sí |
| Escribir .docx | docx (ya en PapelPDF) | MIT | sí |
| Leer/escribir .xlsx | ExcelJS | MIT | sí |
| Leer .pptx | pptxtojson | MIT | sí |
| Escribir .pptx | PptxGenJS | MIT | sí |
| Gráficos | Apache ECharts | Apache-2.0 | sí (con GPLv3) |
| PDF | pdf.js, pdf-lib (ya en PapelPDF) | Apache-2.0 / MIT | sí |
| PapelPDF | código propio | MIT | sí (puede relicenciarse a GPL-3.0 por su autor) |

Todas las licencias se verifican de nuevo en `/replica-architect` contra la
versión exacta que se instale.

## 10. Tamaño

Pantallas: 6 de suite + 6 compartidas + 11 Escribiente + 11 Calcula Datos +
10 Impresiona + PapelPDF ya hecho = **44 nuevas**. Flujos: **12**. Entidades
de formato: **4**.

Partes difíciles:

1. **Paginación del procesador de textos** (dividir el texto en páginas reales mientras se escribe, con tablas que cruzan páginas, notas al pie y encabezados). Es lo más difícil de toda la suite.
2. **Fidelidad al abrir y guardar .docx/.xlsx/.pptx** sin destruir lo que no se entiende (conservar partes XML desconocidas).
3. **Hoja de cálculo rápida con 100,000+ filas** (cuadrícula dibujada en canvas, recálculo incremental) y fórmulas localizadas en español.
4. Tablas dinámicas y gráficos compartidos entre las tres apps.
5. Transiciones tipo Transformación y vista del moderador con dos pantallas.
6. Launcher dona: ventana transparente sin marco que se comporta igual en Windows, macOS (sin transparencia de clics) y Linux (Wayland/X11, compositores distintos).

Tamaño de "todas las funciones de Office": **XL → reescalar**. Propuesta por fases:

| fase | contenido | tamaño |
| --- | --- | --- |
| v0.1 "Primera prueba" | Launcher dona + logos + PapelPDF integrado + las 3 apps con lo básico (texto con formato, hoja con fórmulas, diapositivas con texto/imagen), guardar .docx/.xlsx/.pptx, exportar PDF, instaladores para los 3 sistemas | M |
| v0.5 | Todas las filas **must** | L |
| v1.0 | Todas las **should** | L |
| v2.x | **could**: combinar correspondencia avanzada, Transformación, gráficos avanzados, video | L |
