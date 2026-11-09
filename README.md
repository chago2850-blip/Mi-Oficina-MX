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

## Funciones (versión 0.1)

**Común a las tres apps:** cinta con pestañas, inicio con plantillas y recientes, guardar / guardar como, exportar a PDF, imprimir, deshacer y rehacer, buscar y reemplazar, arrastrar y soltar archivos, **recuperación automática** si la app se cierra sin guardar, revisión ortográfica en español e inglés, y el botón para volver a la dona.

**Escribiente**
- Páginas reales (Carta, Oficio México, Legal, A4, A5), orientación y márgenes; la vista en pantalla y el PDF tienen las mismas páginas
- Fuentes, tamaño, negrita, cursiva, subrayado, tachado, sub/superíndice, color y resaltado
- Estilos (Título, Subtítulo, Título 1–4, Cita), alineación, interlineado, sangrías, viñetas y numeración
- Tablas (insertar/eliminar filas y columnas, sombreado), imágenes con cambio de tamaño, vínculos, saltos de página, símbolos
- Encabezado, pie y números de página; índice automático; contador de palabras
- Plantillas: carta formal, reporte y currículum

**Calcula Datos**
- Cuadrícula rápida (dibujada en canvas), varias hojas, inmovilizar paneles, cambiar ancho/alto, autoajuste
- **Unas 400 funciones** (motor HyperFormula) escritas **en español** — `=SUMA()`, `=SI()`, `=BUSCARV()`, `=BUSCARX()`, `=CONTAR.SI()`… — y guardadas en inglés dentro del `.xlsx` para que Excel las entienda
- Formato de número (moneda, porcentaje, fechas dd/mm/aaaa, miles, decimales), fuente, relleno, bordes, alineación, ajustar texto, combinar y centrar
- Controlador de relleno con series (1, 2, 3… / enero, febrero… / Producto 1, 2…), copiar y pegar con referencias relativas
- Ordenar, filtros por columna, autosuma, insertar función con ayuda en español
- Gráficos de columnas, barras, líneas, área, circular y dispersión
- Plantillas: presupuesto mensual, ventas por mes y lista de asistencia

**Impresiona**
- Diapositivas con diseños (título, contenido, dos contenidos, sección…), 8 temas originales, fondo, tamaño 16:9 o 4:3
- Cuadros de texto, formas, imágenes y tablas; mover, cambiar tamaño, girar, guías inteligentes, alinear y ordenar
- Transiciones (desvanecer, empujar, barrido, cubrir, zoom) y animaciones de entrada
- Presentación a pantalla completa con puntero láser, pantalla negra/blanca (B/W) y saltar a una diapositiva
- Notas del orador, reordenar arrastrando las miniaturas
- Plantillas: plan de negocio, clase y avance de proyecto

**PapelPDF** — la app completa del repositorio [PapelPDF](https://github.com/chago2850-blip/PapelPDF): anotar, editar, formularios, firma, organizar páginas, convertir, contraseña, redacción y OCR sin conexión.

## Instalar

**[⬇ Descargar la última versión](https://github.com/chago2850-blip/Mi-Oficina-MX/releases/latest)**: elige el archivo para tu sistema.

| Sistema | Archivo | Cómo |
| --- | --- | --- |
| Windows 10/11 | `Mi Oficina MX Setup 0.1.0.exe` | Ábrelo y sigue el asistente. Como no está firmado, SmartScreen puede avisar: **Más información → Ejecutar de todas formas**. |
| macOS Apple Silicon (M1–M4) | `MiOficinaMX-0.1.0-arm64.dmg` | Abre el `.dmg` y arrastra Mi Oficina MX a Aplicaciones. Ve abajo, *Primera vez en macOS*. |
| macOS Intel | `MiOficinaMX-0.1.0-x64.dmg` | Igual que el anterior. |
| Ubuntu, Debian, Mint | `mi-oficina-mx_0.1.0_amd64.deb` | `sudo apt install ./mi-oficina-mx_0.1.0_amd64.deb` |
| Fedora, RHEL, openSUSE | `mi-oficina-mx-0.1.0.x86_64.rpm` | `sudo dnf install ./mi-oficina-mx-0.1.0.x86_64.rpm` |
| Cualquier Linux | `Mi Oficina MX-0.1.0.AppImage` | `chmod +x "Mi Oficina MX-0.1.0.AppImage"` y ábrelo |

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

Revisa [`replica/features.csv`](replica/features.csv) para la lista completa. Lo más importante que sigue: control de cambios y comentarios en Escribiente, tablas dinámicas y formato condicional en Calcula Datos, vista del moderador con dos pantallas en Impresiona, y abrir/guardar ODF (`.odt`, `.ods`, `.odp`).

**Limitaciones conocidas de la v0.1:** al abrir un `.docx` se conservan texto, estilos de título, negritas, listas, tablas, imágenes y alineación, pero no siempre los colores ni las fuentes de cada palabra; los gráficos de Calcula Datos se guardan en el `.xlsx` de forma que Calcula Datos los recupera, pero Excel no los muestra; las transiciones de Impresiona no se guardan en el `.pptx`. Las macros VBA no se ejecutan.

## Licencia

[GPL-3.0-or-later](LICENSE). Puedes usar, estudiar, compartir y modificar Mi Oficina MX; si distribuyes una versión modificada, también debe ser libre. Créditos de las bibliotecas en [THIRD_PARTY.md](THIRD_PARTY.md).
