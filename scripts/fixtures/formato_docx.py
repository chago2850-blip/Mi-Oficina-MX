# Genera test/fixtures/formato.docx: documento con formato variado para probar el lector .docx
# (se ejecuta una vez; el archivo generado se guarda en el repositorio). Requiere python-docx y Pillow.
import io, sys
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_COLOR_INDEX, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image, ImageDraw

out = sys.argv[1]
d = Document()
sec = d.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2)
sec.top_margin = sec.bottom_margin = Cm(2.5)
normal = d.styles['Normal']; normal.font.name = 'Arial'; normal.font.size = Pt(12)

def field(par, instr):
    r = par.add_run(); f = OxmlElement('w:fldChar'); f.set(qn('w:fldCharType'), 'begin'); r._r.append(f)
    r = par.add_run(); t = OxmlElement('w:instrText'); t.set(qn('xml:space'), 'preserve'); t.text = f' {instr} '; r._r.append(t)
    r = par.add_run(); f = OxmlElement('w:fldChar'); f.set(qn('w:fldCharType'), 'separate'); r._r.append(f)
    par.add_run('1')
    r = par.add_run(); f = OxmlElement('w:fldChar'); f.set(qn('w:fldCharType'), 'end'); r._r.append(f)

hp = sec.header.paragraphs[0]; hp.text = 'Informe anual — Mi Empresa'
fp = sec.footer.paragraphs[0]; fp.add_run('Página '); field(fp, 'PAGE'); fp.add_run(' de '); field(fp, 'NUMPAGES')

d.add_paragraph('Título del documento', style='Title')
d.add_heading('Introducción', level=1)

p = d.add_paragraph('Párrafo justificado con sangría de primera línea, espacio antes de 18 pt, después de 6 pt e interlineado doble.')
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
pf = p.paragraph_format; pf.first_line_indent = Cm(1.25); pf.space_before = Pt(18); pf.space_after = Pt(6); pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE

p = d.add_paragraph('Párrafo con sangría izquierda y derecha de 2 cm e interlineado exacto de 20 pt.')
pf = p.paragraph_format; pf.left_indent = Cm(2); pf.right_indent = Cm(2); pf.line_spacing = Pt(20); pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY; pf.keep_with_next = True

p = d.add_paragraph('Párrafo con borde y sombreado.')
pPr = p._p.get_or_add_pPr()
bdr = OxmlElement('w:pBdr')
for side in ('top', 'left', 'bottom', 'right'):
    e = OxmlElement(f'w:{side}'); e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '12'); e.set(qn('w:space'), '4'); e.set(qn('w:color'), 'C00000'); bdr.append(e)
pPr.append(bdr)
shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), 'FFF2CC'); pPr.append(shd)

p = d.add_paragraph()
r = p.add_run('Rojo Georgia 16 '); r.font.color.rgb = RGBColor(0xC0, 0, 0); r.font.name = 'Georgia'; r.font.size = Pt(16)
r = p.add_run('resaltado '); r.font.highlight_color = WD_COLOR_INDEX.YELLOW
r = p.add_run('versalitas '); r.font.small_caps = True
r = p.add_run('mayúsculas '); r.font.all_caps = True
r = p.add_run('doble tachado '); r.font.double_strike = True
r = p.add_run('negrita cursiva subrayada'); r.bold = True; r.italic = True; r.underline = True
r = p.add_run(' espaciado'); rPr = r._r.get_or_add_rPr(); sp = OxmlElement('w:spacing'); sp.set(qn('w:val'), '60'); rPr.append(sp)
r = p.add_run(' oculto'); r.font.hidden = True
r = p.add_run(' x'); r = p.add_run('2'); r.font.superscript = True

# vínculo
p = d.add_paragraph('Visita ')
rid = d.part.relate_to('https://www.libreoffice.org/', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
hl = OxmlElement('w:hyperlink'); hl.set(qn('r:id'), rid)
r = OxmlElement('w:r'); t = OxmlElement('w:t'); t.text = 'LibreOffice'; r.append(t); hl.append(r); p._p.append(hl)

d.add_heading('Listas', level=2)
d.add_paragraph('Viñeta uno', style='List Bullet')
d.add_paragraph('Viñeta dos', style='List Bullet')
d.add_paragraph('Viñeta anidada', style='List Bullet 2')
d.add_paragraph('Paso uno', style='List Number')
d.add_paragraph('Paso dos', style='List Number')
d.add_paragraph('Paso tres', style='List Number')

d.add_heading('Tabla', level=2)
t = d.add_table(rows=3, cols=3); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, row in enumerate(t.rows):
    for j, c in enumerate(row.cells): c.text = f'C{i}{j}'
a = t.cell(0, 0).merge(t.cell(0, 1)); a.text = 'Combinada horizontal'
b = t.cell(1, 2).merge(t.cell(2, 2)); b.text = 'Combinada vertical'
tcPr = t.cell(1, 0)._tc.get_or_add_tcPr(); s = OxmlElement('w:shd'); s.set(qn('w:val'), 'clear'); s.set(qn('w:fill'), 'D9E2F3'); tcPr.append(s)

img = Image.new('RGB', (200, 100), '#5B48E8'); ImageDraw.Draw(img).ellipse((60, 10, 140, 90), fill='#FFD54A')
bio = io.BytesIO(); img.save(bio, 'PNG'); bio.seek(0)
d.add_paragraph().add_run().add_picture(bio, width=Cm(5))

p = d.add_paragraph('Este párrafo empieza en una página nueva.'); p.paragraph_format.page_break_before = True
p = d.add_paragraph('Antes del salto'); p.add_run().add_break(WD_BREAK.PAGE); p.add_run('Después del salto')
d.save(out)
print('ok', out)
