#!/usr/bin/env python3
"""Genera los datos lingüísticos que usa Escribiente sin conexión a internet:
- apps/shared/data/thes-es.json.gz: diccionario de sinónimos del español (th_es_v2 de RLA-ES,
  GPL-3.0+/LGPL-3.0+/MPL-1.1, https://github.com/sbosio/rla-es) como {palabra: [[sinónimos], ...]}
- apps/shared/data/hyph-es.json: patrones de separación silábica (hyph_es.dic, de Javier Bezos,
  MIT / GPL-3.0+) para el algoritmo de Liang.
Uso: python3 scripts/data/build.py  (descarga de github.com/LibreOffice/dictionaries)."""
import gzip, json, os, re, sys, urllib.request
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
OUT = os.path.join(ROOT, 'apps', 'shared', 'data')
BASE = 'https://raw.githubusercontent.com/LibreOffice/dictionaries/master/es/'
def get(name):
    with urllib.request.urlopen(BASE + name, timeout=120) as r: return r.read()

def thesaurus(raw):
    lines = raw.decode('latin-1').split('\n')
    enc = lines[0].strip()
    if enc.upper() != 'ISO8859-1': lines = raw.decode(enc).split('\n')
    out = {}; i = 1
    while i < len(lines):
        head = lines[i].strip(); i += 1
        if not head or '|' not in head: continue
        word, n = head.rsplit('|', 1)
        try: n = int(n)
        except ValueError: continue
        groups = []
        for _ in range(n):
            if i >= len(lines): break
            parts = [p.strip() for p in lines[i].split('|')]; i += 1
            pos = parts[0].strip('()') if parts and parts[0] not in ('-', '') else ''
            syn = [p for p in parts[1:] if p and p.lower() != word.lower()]
            if syn: groups.append([pos] + syn if pos else syn)
        if groups: out[word.lower()] = groups
    return out

def hyph(raw):
    txt = raw.decode('utf-8', 'replace').split('\n')
    pats = []; mins = {'l': 2, 'r': 2}
    for ln in txt[1:]:
        ln = ln.strip()
        if not ln or ln.startswith('%'): continue
        if ln.startswith('LEFTHYPHENMIN'): mins['l'] = int(ln.split()[1]); continue
        if ln.startswith('RIGHTHYPHENMIN'): mins['r'] = int(ln.split()[1]); continue
        if re.match(r'^[A-Z]', ln) and ' ' in ln: continue
        pats.extend(ln.split())
    return {'left': mins['l'], 'right': mins['r'], 'patterns': pats}

os.makedirs(OUT, exist_ok=True)
th = thesaurus(get('th_es_v2.dat'))
with gzip.open(os.path.join(OUT, 'thes-es.json.gz'), 'wt', encoding='utf-8', compresslevel=9) as f: json.dump(th, f, ensure_ascii=False, separators=(',', ':'))
hy = hyph(get('hyph_es.dic'))
with open(os.path.join(OUT, 'hyph-es.json'), 'w', encoding='utf-8') as f: json.dump(hy, f, ensure_ascii=False, separators=(',', ':'))
print(f'sinónimos: {len(th)} entradas; separación: {len(hy["patterns"])} patrones')
