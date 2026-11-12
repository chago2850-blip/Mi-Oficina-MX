#!/usr/bin/env bash
# Genera todos los íconos (PNG, ICO, ICNS) a partir de los SVG de build/brand.
# Requiere rsvg-convert (librsvg), ImageMagick (convert) y python3.
set -euo pipefail
cd "$(dirname "$0")/../build"
B=brand

png() { rsvg-convert -w "$2" -h "$2" "$B/$1.svg" -o "$3"; }

icns() { # $1 = nombre del svg, $2 = salida .icns (PNG dentro de ICNS: ic07..ic10)
  local tmp; tmp=$(mktemp -d)
  png "$1" 128 "$tmp/128.png"; png "$1" 256 "$tmp/256.png"; png "$1" 512 "$tmp/512.png"; png "$1" 1024 "$tmp/1024.png"
  python3 - "$tmp" "$2" <<'PY'
import struct, sys
d, out = sys.argv[1], sys.argv[2]
chunks = b''
for t, n in [(b'ic07', 128), (b'ic08', 256), (b'ic09', 512), (b'ic10', 1024)]:
    data = open(f'{d}/{n}.png', 'rb').read()
    chunks += t + struct.pack('>I', len(data) + 8) + data
open(out, 'wb').write(b'icns' + struct.pack('>I', len(chunks) + 8) + chunks)
PY
  rm -rf "$tmp"
}

ico() { # $1 = svg, $2 = salida .ico
  local tmp; tmp=$(mktemp -d)
  for s in 16 24 32 48 64 128 256; do png "$1" $s "$tmp/$s.png"; done
  convert "$tmp"/16.png "$tmp"/24.png "$tmp"/32.png "$tmp"/48.png "$tmp"/64.png "$tmp"/128.png "$tmp"/256.png "$2"
  rm -rf "$tmp"
}

# Suite
png suite 1024 icon.png
ico suite icon.ico
icns suite icon.icns
mkdir -p icons
for s in 16 24 32 48 64 128 256 512 1024; do png suite $s "icons/${s}x${s}.png"; done

# Apps (asociaciones de archivo, ventanas y lanzador)
mkdir -p apps
for a in escribiente calcula impresiona papelpdf; do
  png "$a" 512 "apps/$a.png"
  ico "$a" "apps/$a.ico"
  icns "$a" "apps/$a.icns"
done
# El .csv usa el ícono de Calcula Datos, pero macOS exige un archivo distinto por asociación
for e in png ico icns; do cp "apps/calcula.$e" "apps/csv.$e"; cp "apps/escribiente.$e" "apps/odt.$e"; cp "apps/escribiente.$e" "apps/doc.$e"; cp "apps/escribiente.$e" "apps/rtf.$e"; cp "apps/calcula.$e" "apps/ods.$e"; cp "apps/impresiona.$e" "apps/odp.$e"; done
echo "[íconos] listos"
