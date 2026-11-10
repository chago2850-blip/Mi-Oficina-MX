<p align="center"><img src="docs/logo.png" width="120" alt="Mi Oficina MX"></p>

# Mi Oficina MX

Suite ofimática **libre, gratuita y de código abierto** (GPL-3.0) para **Windows, macOS y Linux**.
Abre y guarda documentos de Word, Excel y PowerPoint (`.docx`, `.xlsx`, `.pptx`), exporta a PDF y funciona sin conexión.

**Sin anuncios, sin cuentas, sin telemetría, sin límites.**

Desarrollado por Santiago Miguel · 744 354 8848 · chago2850@gmail.com

| | App | Para qué sirve | Formatos |
| --- | --- | --- | --- |
| <img src="docs/logo-escribiente.png" width="40"> | **Escribiente** | Procesador de textos: cartas, reportes, currículum | `.docx` (abrir y guardar), `.txt`, `.html`, `.md`, PDF |
| <img src="docs/logo-calcula.png" width="40"> | **Calcula Datos** | Hoja de cálculo con fórmulas en español | `.xlsx` (abrir y guardar), `.csv`, PDF |
| <img src="docs/logo-impresiona.png" width="40"> | **Impresiona** | Presentaciones con temas, transiciones y animaciones | `.pptx` (abrir y guardar), PDF |
| <img src="docs/logo-papelpdf.png" width="40"> | **PapelPDF** | Editor de PDF: anotar, firmar, formularios, OCR | `.pdf` |

## El Launcher

Al abrir Mi Oficina MX aparece una **dona**. Pulsa **Inicio**: se enciende y se divide en cuatro partes, una por app.
Al abrir una app la dona se minimiza, y cuando cierras la última app vuelve a aparecer.

<p align="center"><img src="docs/launcher.png" width="300" alt="Launcher tipo dona"></p>

Atajos de la dona: **1–4** abren cada app, las **flechas** recorren los sectores, **Esc** la cierra. Arrastra la barra "Mi Oficina MX" para moverla.

## Capturas

![Escribiente](docs/escribiente.png)
![Calcula Datos](docs/calcula.png)
![Impresiona](docs/impresiona.png)

## Funciones (versión 0.4)

**Común a las tres apps:** cinta con pestañas, inicio con plantillas y recientes, guardar / guardar como, exportar a PDF, imprimir, deshacer y rehacer, buscar y reemplazar, arrastrar y soltar archivos, **recuperación automática** si la app se cierra sin guardar, revisión ortográfica en español e inglés, y el botón para volver a la dona.

**Nuevo en la 0.2 para toda la suite:** abrir y guardar **ODF** (`.odt`, `.ods`, `.odp`), **modo oscuro** (sigue al sistema o se elige), **Preferencias** (idioma, nombre del autor, tema e intervalo de autoguardado) y **Propiedades del documento** (título, autor, asunto, palabras clave).

**Interfaz en español e inglés (0.2.1):** el idioma se elige en *Preferencias* (en la dona) o en *Archivo → Opciones*, y cambia al momento, sin reiniciar: cinta, menús, diálogos, avisos, plantillas nuevas y PapelPDF. Con *Igual que el sistema* se usa español si el sistema está en español y si no, inglés. En inglés, Calcula Datos muestra las fórmulas con los nombres en inglés (`=SUM()`, `=IF()`), como Excel en inglés; el archivo guardado es el mismo en ambos idiomas.

**Nuevo en la 0.3 para toda la suite:** **cifrar con contraseña** (`.docx`, `.xlsx` y `.pptx` con el cifrado estándar de Office, AES-256; también abre los archivos cifrados por Office) y **barra de acceso rápido** personalizable en cada app.

**Nuevo en la 0.4 para toda la suite:** **accesibilidad** (lector de pantalla y teclado: flechas para moverse por la cinta, **F6** para pasar entre la cinta, el documento y la barra de estado, foco visible, el foco no se escapa de los diálogos y Calcula Datos anuncia la celda activa; contraste AA revisado con axe-core), **revisión ortográfica sin internet** (los diccionarios vienen en el instalador) y archivos grandes más rápidos (un CSV de 100 000 filas abre en unos 2 s; un documento de 300 páginas, en menos de 1 s).

**Escribiente**
- Páginas reales (Carta, Oficio México, Legal, A4, A5), orientación y márgenes; la vista en pantalla y el PDF tienen las mismas páginas
- Fuentes, tamaño, negrita, cursiva, subrayado, tachado, sub/superíndice, color y resaltado
- Estilos (Título, Subtítulo, Título 1–4, Cita), alineación, interlineado, sangrías, viñetas y numeración
- Tablas (insertar/eliminar filas y columnas, sombreado), imágenes con cambio de tamaño, vínculos, saltos de página, símbolos
- Encabezado, pie y números de página; índice automático; contador de palabras
- Plantillas: carta formal, reporte y currículum
- **0.2:** comentarios con respuestas, **control de cambios** (aceptar/rechazar), **notas al pie y al final**, **tabulaciones** (izquierda, centro, derecha) en la regla, columnas, cuadros de texto y formas, marca de agua y color de página, panel de navegación por títulos, brocha para copiar formato, sangrías arrastrables en la regla y **combinar correspondencia** desde CSV o XLSX. Todo se guarda en el `.docx` de forma que Word lo entiende
- **0.3:** **ecuaciones** (escritas en LaTeX con vista previa; en el `.docx` quedan como ecuaciones de Word y en el `.odt` como fórmulas), **citas y bibliografía** (APA), **comparar documentos** (el resultado queda como control de cambios), **etiquetas y sobres** (también con los destinatarios de la combinación), **diagramas** tipo SmartArt (lista, proceso, ciclo, jerarquía y pirámide), **modo lectura** y **modo enfoque**, reacciones en comentarios, exportar a **Markdown** y abrir documentos `.doc` de Word 97-2003
- **0.4.2:** el **control de cambios** se ve en la barra de estado (*Control de cambios: activado/desactivado*, con un clic se apaga); con el control de cambios activo, lo que tú escribiste se borra normalmente con Retroceso, el cursor ya no salta y la ortografía sigue marcando palabras y dando sugerencias con el clic derecho
- **0.4:** abre las **ecuaciones de Word y LibreOffice** como ecuaciones editables, abre los `.doc` antiguos **con su formato** (títulos, negritas, colores, tamaños, tablas, listas y alineación), y guarda dentro del `.docx`/`.odt` las **reacciones** a comentarios, los diagramas tipo SmartArt editables, las fuentes de las citas y el LaTeX de las ecuaciones

**Calcula Datos**
- Cuadrícula rápida (dibujada en canvas), varias hojas, inmovilizar paneles, cambiar ancho/alto, autoajuste
- **Unas 400 funciones** (motor HyperFormula) escritas **en español** — `=SUMA()`, `=SI()`, `=BUSCARV()`, `=BUSCARX()`, `=CONTAR.SI()`… — y guardadas en inglés dentro del `.xlsx` para que Excel las entienda
- Formato de número (moneda, porcentaje, fechas dd/mm/aaaa, miles, decimales), fuente, relleno, bordes, alineación, ajustar texto, combinar y centrar
- Controlador de relleno con series (1, 2, 3… / enero, febrero… / Producto 1, 2…), copiar y pegar con referencias relativas
- Ordenar, filtros por columna, autosuma, insertar función con ayuda en español
- Gráficos de columnas, barras, líneas, área, circular y dispersión
- Plantillas: presupuesto mensual, ventas por mes y lista de asistencia
- **0.2:** **formato condicional** (reglas, barras de datos, escalas de color, íconos), **validación** con listas desplegables, administrador de nombres, notas en celdas, tablas con formato, **tablas dinámicas**, proteger hoja y libro, área de impresión y ajustar a página, y abrir `.xlsm` sin ejecutar macros (se conservan al guardar)
- **0.3:** funciones **LAMBDA**, LET, MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY, APILARV/APILARH, TOMAR, EXCLUIR, ENCOL, ENFILA, ELEGIRCOLS/ELEGIRFILAS, AJUSTARFILAS, TEXTOANTES/TEXTODESPUES e **IMAGEN**; **buscar objetivo** y **Solver** (Simplex con restricciones, enteros y binarios; búsqueda directa en modelos no lineales); **minigráficos** (línea, columna y ganancia/pérdida); **agrupar y esquema** de filas y columnas (esquema nativo de Excel); **segmentación de datos**; y gráficos que siguen matrices dinámicas (rango `A4#`)
- **0.4:** **Obtener y transformar datos** (como Power Query): trae datos de **CSV/TXT, un libro de Excel, JSON** o un rango, y los limpia con **pasos aplicados** que se pueden revisar y quitar — quitar o renombrar columnas, filtrar, ordenar, quitar duplicados, cambiar tipo, dividir, reemplazar, recortar, mayúsculas, rellenar hacia abajo, **columna personalizada** (`[Precio] * [Cantidad]`), **agrupar** (suma, promedio, contar…), **anexar** y **combinar** consultas. *Cerrar y cargar* deja el resultado en una hoja; **Actualizar todo** vuelve a leer los archivos de origen. Las consultas se guardan en el `.xlsx`

**Impresiona**
- Diapositivas con diseños (título, contenido, dos contenidos, sección…), 8 temas originales, fondo, tamaño 16:9 o 4:3
- Cuadros de texto, formas, imágenes y tablas; mover, cambiar tamaño, girar, guías inteligentes, alinear y ordenar
- Transiciones (desvanecer, empujar, barrido, cubrir, zoom) y animaciones de entrada
- Presentación a pantalla completa con puntero láser, pantalla negra/blanca (B/W) y saltar a una diapositiva
- Notas del orador, reordenar arrastrando las miniaturas
- Plantillas: plan de negocio, clase y avance de proyecto
- **0.2:** **gráficos** (se guardan como gráficos nativos del `.pptx`), animaciones de salida y **panel de animaciones** (al hacer clic / con la anterior / después), **patrón de diapositivas** con pie, número y fecha, clasificador con secciones, **vista del moderador** con ventana para el público en la segunda pantalla, pluma durante la presentación y exportar cada diapositiva como PNG
- **0.3:** **video y audio** en las diapositivas (se guardan dentro del `.pptx`) con **subtítulos** WebVTT o SRT, transición **Transformación** (Morph), **trayectorias de animación**, **ensayar intervalos** con avance automático y **exportar a video** (WebM) con transiciones y pasos de animación. Las animaciones, trayectorias, intervalos, patrón, secciones y subtítulos se conservan al guardar y abrir el `.pptx`
- **0.4:** **cámara en vivo** dentro de la diapositiva (círculo, rectángulo redondeado o rectángulo, con espejo y borde, como Cameo) y **grabar la presentación con narración**: tu voz y el tiempo de cada diapositiva quedan guardados, la presentación avanza sola con la narración y el video exportado lleva el audio
- **0.4.1:** **Diseño → Tamaño de diapositiva → Ajustes avanzados**: tamaños predefinidos (pantalla 16:9, 16:10 y 4:3, panorámica, carta, contabilidad, A3, A4, B4, B5, 35 mm, transparencia, banner o personalizado), ancho y alto en cm, orientación horizontal o vertical, numerar las diapositivas desde cualquier número y elegir si el contenido se ajusta o se maximiza. Se guarda en el `.pptx` para PowerPoint

**PapelPDF** — la app completa del repositorio [PapelPDF](https://github.com/chago2850-blip/PapelPDF): anotar, editar, formularios, firma, organizar páginas, convertir, contraseña, redacción y OCR sin conexión.

## Instalar

**[⬇ Descargar la última versión](https://github.com/chago2850-blip/Mi-Oficina-MX/releases/latest)**: elige el archivo para tu sistema.

| Sistema | Archivo | Cómo |
| --- | --- | --- |
| Windows 10/11 | `Mi.Oficina.MX.Setup.0.4.2.exe` | Ábrelo y sigue el asistente. Como no está firmado, SmartScreen puede avisar: **Más información → Ejecutar de todas formas**. |
| macOS Apple Silicon (M1–M4) | `MiOficinaMX-0.4.2-arm64.dmg` | Abre el `.dmg` y arrastra Mi Oficina MX a Aplicaciones. Ve abajo, *Primera vez en macOS*. |
| macOS Intel | `MiOficinaMX-0.4.2-x64.dmg` | Igual que el anterior. |
| Ubuntu, Debian, Mint | `mi-oficina-mx_0.4.2_amd64.deb` | `sudo apt install ./mi-oficina-mx_0.4.2_amd64.deb` |
| Fedora, RHEL, openSUSE | `mi-oficina-mx-0.4.2.x86_64.rpm` | `sudo dnf install ./mi-oficina-mx-0.4.2.x86_64.rpm` |
| Cualquier Linux | `Mi.Oficina.MX-0.4.2.AppImage` | `chmod +x Mi.Oficina.MX-0.4.2.AppImage` y ábrelo |

El instalador asocia los archivos `.docx`, `.xlsx`, `.csv`, `.pptx` y `.pdf`: con doble clic se abren en la app correcta.

### Primera vez en macOS

La app tiene firma *ad-hoc*, pero no está notarizada por Apple (eso requiere una cuenta de Apple Developer de pago). La primera vez:

1. Abre **Aplicaciones**, haz **clic derecho en Mi Oficina MX → Abrir** y confirma **Abrir**.
2. Si macOS dice que la app "está dañada", ejecuta esto en Terminal y vuelve a abrirla:
   ```bash
   xattr -cr "/Applications/Mi Oficina MX.app"
   ```

## Desarrollar

```bash
npm install                   # dependencias, modelos de OCR y empaquetado de la interfaz
npm start                     # abre la suite (la dona)
node test/formats.test.mjs    # pruebas de formatos de número y fórmulas en español
npm i --no-save playwright@1.49.1 && xvfb-run -a node test/e2e.mjs   # punta a punta (Linux sin pantalla)
xvfb-run -a node test/e2e-fase2.mjs              # funciones de la fase 2 (OFICINA_LANG=en para probar en inglés)
xvfb-run -a node test/e2e-fase3.mjs              # funciones de la fase 3
xvfb-run -a node test/e2e-fase4.mjs              # funciones de la fase 4
xvfb-run -a node test/a11y.mjs                   # auditoría de accesibilidad (axe-core)
xvfb-run -a node test/perf.mjs calcula           # tiempos con archivos grandes (o: escribiente)
xvfb-run -a npx electron --no-sandbox scripts/dict   # descarga los diccionarios de ortografía a build/dict
python3 scripts/i18n/merge.py                     # regenera apps/shared/i18n-en.js desde scripts/i18n/en-*.py
xvfb-run -a node test/i18n-crawl.mjs              # recorre la interfaz en inglés y lista textos sin traducir
npm run dist:linux            # AppImage + deb + rpm
npm run dist:win              # instalador NSIS (en Windows)
npm run dist:mac              # en una Mac: dmg + zip
npm run icons                 # regenera los íconos desde build/brand/*.svg
```

Cada vez que se sube código a `main`, `.github/workflows/build.yml` ejecuta las pruebas, compila el `.exe` en Windows y el `.deb`, `.rpm`, `.AppImage` y `.dmg` en Linux, y los publica en el *Release* `v<versión>` según el campo `version` de `package.json`.

## Arquitectura

```
main.js                 proceso principal: protocolo app://, dona, ventanas, archivos, PDF, recuperación
main/papelpdf.js        PapelPDF (proceso principal) adaptado a varias ventanas
preload/                puentes seguros (contextIsolation + sandbox)
apps/launcher/          la dona (SVG + CSS)
apps/shared/            cinta, inicio, diálogos e íconos compartidos
apps/escribiente/src/   editor paginado, docx ⇄ HTML
apps/calcula/src/       cuadrícula en canvas, HyperFormula, xlsx/csv, gráficos
apps/impresiona/src/    editor de diapositivas, presentación, pptx
apps/papelpdf/          interfaz de PapelPDF
build/brand/            logotipos originales (SVG)
replica/                mapa de funciones y pantallas del proyecto
```

## Lo que falta (siguientes versiones)

Revisa [`replica/features.csv`](replica/features.csv) para la lista completa. Las fases 1 a 4 están completas: todas las funciones *must*, *should* y *could* del mapa; solo el SmartArt queda con alcance parcial (ver abajo). Lo que queda fuera a propósito (*skip*) son las funciones en la nube: coautoría en tiempo real, Copilot, OneDrive y cuentas.

**Limitaciones conocidas:** al abrir un `.docx` se conservan texto, estilos de título, negritas, listas, tablas, imágenes, notas al pie y al final, tabulaciones, comentarios con sus respuestas, control de cambios y alineación, pero no siempre los colores ni las fuentes de cada palabra; los diagramas tipo SmartArt se ven en Word como imagen (en Escribiente siguen editables); los `.doc` antiguos se abren sin imágenes; PowerPoint ve el patrón de Impresiona aplanado en cada diapositiva y la transición Transformación como un cambio sin efecto; los gráficos de Calcula Datos se recuperan en Calcula Datos, pero Excel no los muestra; Excel no ve las consultas de *Obtener y transformar* (solo la hoja con el resultado); la cámara en vivo se guarda en el `.pptx` pero PowerPoint no la reconoce; la protección de hoja no usa contraseña cifrada. Las macros VBA nunca se ejecutan.

## Licencia

[GPL-3.0-or-later](LICENSE). Puedes usar, estudiar, compartir y modificar Mi Oficina MX; si distribuyes una versión modificada, también debe ser libre. Créditos de las bibliotecas en [THIRD_PARTY.md](THIRD_PARTY.md).
