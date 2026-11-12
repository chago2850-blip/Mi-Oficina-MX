#!/usr/bin/env python3
"""Copia las fuentes libres que trae LibreOffice (y algunas más muy usadas) a apps/shared/fonts/
y escribe apps/shared/fonts.css con sus @font-face, para que la suite tenga las mismas letras en
Windows, macOS y Linux sin depender de lo que tenga instalado cada equipo.

Origen: paquetes @fontsource/* (Google Fonts, SIL OFL/Apache) y dejavu-fonts-ttf (licencia de
Bitstream Vera/DejaVu). Uso: python3 scripts/fonts/build.py   (después de npm install)
"""
import json, os, re, shutil, subprocess, sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
NM = os.path.join(ROOT, 'node_modules')
OUT = os.path.join(ROOT, 'apps', 'shared', 'fonts')
CSS = os.path.join(ROOT, 'apps', 'shared', 'fonts.css')
LIST = os.path.join(ROOT, 'apps', 'shared', 'fonts.json')

# nombre en el menú (como en LibreOffice), otros nombres, paquete, tipo, alfabetos extra, licencia
FAMILIES = [
  # equivalentes métricos de las letras de Office (las que LibreOffice usa para abrir .docx sin que se muevan las líneas)
  ('Liberation Sans', ['Arimo'], 'arimo', 'sans', [], 'Apache-2.0'),
  ('Liberation Serif', ['Tinos'], 'tinos', 'serif', [], 'Apache-2.0'),
  ('Liberation Mono', ['Cousine'], 'cousine', 'mono', [], 'Apache-2.0'),
  ('Carlito', [], 'carlito', 'sans', [], 'OFL-1.1'),
  ('Caladea', [], 'caladea', 'serif', [], 'OFL-1.1'),
  ('Gelasio', [], 'gelasio', 'serif', [], 'OFL-1.1'),
  # las demás que trae LibreOffice
  ('DejaVu Sans', [], 'dejavu-sans', 'sans', ['greek', 'cyrillic'], 'Bitstream-Vera'),
  ('DejaVu Serif', [], 'dejavu-serif', 'serif', ['greek', 'cyrillic'], 'Bitstream-Vera'),
  ('Linux Libertine G', ['Libertinus Serif', 'Linux Libertine O'], 'libertinus-serif', 'serif', ['greek', 'cyrillic'], 'OFL-1.1'),
  ('Linux Biolinum G', ['Libertinus Sans', 'Linux Biolinum O'], 'libertinus-sans', 'sans', ['greek', 'cyrillic'], 'OFL-1.1'),
  ('Gentium Book Basic', ['Gentium Book Plus', 'Gentium Basic'], 'gentium-book-plus', 'serif', ['greek', 'cyrillic'], 'OFL-1.1'),
  ('Noto Sans', [], 'noto-sans', 'sans', ['greek', 'cyrillic'], 'OFL-1.1'),
  ('Noto Serif', [], 'noto-serif', 'serif', ['greek', 'cyrillic'], 'OFL-1.1'),
  ('Noto Sans Mono', [], 'noto-sans-mono', 'mono', [], 'OFL-1.1'),
  ('Source Sans Pro', ['Source Sans 3'], 'source-sans-3', 'sans', [], 'OFL-1.1'),
  ('Source Serif Pro', ['Source Serif 4'], 'source-serif-4', 'serif', [], 'OFL-1.1'),
  ('Source Code Pro', [], 'source-code-pro', 'mono', [], 'OFL-1.1'),
  ('Open Sans', [], 'open-sans', 'sans', [], 'OFL-1.1'),
  ('Karla', [], 'karla', 'sans', [], 'OFL-1.1'),
  ('Rubik', [], 'rubik', 'sans', ['hebrew'], 'OFL-1.1'),
  ('EB Garamond', [], 'eb-garamond', 'serif', [], 'OFL-1.1'),
  ('Alef', [], 'alef', 'sans', ['hebrew'], 'OFL-1.1'),
  ('Frank Ruhl Libre', ['Frank Ruehl CLM'], 'frank-ruhl-libre', 'serif', ['hebrew'], 'OFL-1.1'),
  ('David Libre', ['David CLM'], 'david-libre', 'serif', ['hebrew'], 'OFL-1.1'),
  ('Miriam Libre', ['Miriam CLM'], 'miriam-libre', 'sans', ['hebrew'], 'OFL-1.1'),
  ('Noto Sans Hebrew', [], 'noto-sans-hebrew', 'sans', ['hebrew'], 'OFL-1.1'),
  ('Amiri', [], 'amiri', 'serif', ['arabic'], 'OFL-1.1'),
  ('Reem Kufi', [], 'reem-kufi', 'sans', ['arabic'], 'OFL-1.1'),
  ('Scheherazade New', ['Scheherazade'], 'scheherazade-new', 'serif', ['arabic'], 'OFL-1.1'),
  ('Noto Sans Arabic', [], 'noto-sans-arabic', 'sans', ['arabic'], 'OFL-1.1'),
  # libres y muy usadas
  ('Inter', [], 'inter', 'sans', [], 'OFL-1.1'),
  ('Lato', [], 'lato', 'sans', [], 'OFL-1.1'),
  ('Roboto', [], 'roboto', 'sans', [], 'OFL-1.1'),
  ('Montserrat', [], 'montserrat', 'sans', [], 'OFL-1.1'),
  ('Comic Neue', [], 'comic-neue', 'sans', [], 'OFL-1.1'),
]
BASE_SUBSETS = ['latin', 'latin-ext']
GENERIC = {'sans': 'sans-serif', 'serif': 'serif', 'mono': 'monospace'}

def faces(pkg, subset):
  d = os.path.join(NM, '@fontsource', pkg)
  out = []
  for w in (400, 700):
    for style in ('normal', 'italic'):
      css = os.path.join(d, f"{subset}-{w}{'-italic' if style == 'italic' else ''}.css")
      if not os.path.exists(css):
        continue
      txt = open(css, encoding='utf-8').read()
      m = re.search(r"url\(\./files/([^)]+\.woff2)\)", txt)
      ur = re.search(r"unicode-range:\s*([^;]+);", txt)
      if m:
        out.append((w, style, m.group(1), ur.group(1).strip() if ur else None))
  return out

def main():
  if os.path.isdir(OUT):
    shutil.rmtree(OUT)
  os.makedirs(OUT)
  rules = ['/* Generado por scripts/fonts/build.py: fuentes libres incluidas en la suite (las de LibreOffice y otras). */']
  menu = []
  total = 0
  for name, aliases, pkg, kind, extra, lic in FAMILIES:
    fdir = os.path.join(OUT, pkg)
    os.makedirs(fdir, exist_ok=True)
    got = False
    for subset in BASE_SUBSETS + extra:
      for w, style, fname, ur in faces(pkg, subset):
        src = os.path.join(NM, '@fontsource', pkg, 'files', fname)
        shutil.copy(src, os.path.join(fdir, fname)); total += os.path.getsize(src); got = True
        for fam in [name] + aliases:
          rules.append(f"@font-face{{font-family:'{fam}';font-style:{style};font-weight:{w};font-display:swap;src:url(fonts/{pkg}/{fname}) format('woff2');{f'unicode-range:{ur};' if ur else ''}}}")
    if not got:
      print('sin archivos para', pkg, file=sys.stderr); continue
    menu.append({'name': name, 'kind': kind, 'stack': f"'{name}', {GENERIC[kind]}", 'license': lic})
  # DejaVu Sans Mono (no está en Fontsource): se convierte a woff2 con fontTools
  dj = os.path.join(NM, 'dejavu-fonts-ttf', 'ttf')
  if os.path.isdir(dj):
    fdir = os.path.join(OUT, 'dejavu-sans-mono'); os.makedirs(fdir, exist_ok=True)
    for f, w, style in [('DejaVuSansMono.ttf', 400, 'normal'), ('DejaVuSansMono-Bold.ttf', 700, 'normal'), ('DejaVuSansMono-Oblique.ttf', 400, 'italic'), ('DejaVuSansMono-BoldOblique.ttf', 700, 'italic')]:
      dst = os.path.join(fdir, f.replace('.ttf', '.woff2'))
      subprocess.run([sys.executable, '-m', 'fontTools.subset', os.path.join(dj, f), '--unicodes=U+0000-024F,U+0370-03FF,U+0400-04FF,U+1E00-1EFF,U+2000-206F,U+20A0-20CF,U+2100-214F,U+2190-21FF,U+2200-22FF,U+2500-259F,U+25A0-25FF',
                      '--flavor=woff2', f'--output-file={dst}', '--layout-features=*'], check=True)
      total += os.path.getsize(dst)
      rules.append(f"@font-face{{font-family:'DejaVu Sans Mono';font-style:{style};font-weight:{w};font-display:swap;src:url(fonts/dejavu-sans-mono/{os.path.basename(dst)}) format('woff2');}}")
    menu.append({'name': 'DejaVu Sans Mono', 'kind': 'mono', 'stack': "'DejaVu Sans Mono', monospace", 'license': 'Bitstream-Vera'})
  open(CSS, 'w', encoding='utf-8').write('\n'.join(rules) + '\n')
  menu.sort(key=lambda m: m['name'].lower())
  json.dump(menu, open(LIST, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
  print(f'{len(menu)} familias, {total / 1048576:.1f} MB')

main()
