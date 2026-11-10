# Une los diccionarios parciales (scripts/i18n/en-*.py) en apps/shared/i18n-en.js
import json, glob, os, runpy
root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
d = {}; rules = []
for f in sorted(glob.glob(os.path.join(root, 'scripts', 'i18n', 'en-*.py'))):
    g = runpy.run_path(f)
    d.update(g.get('D', {})); rules += g.get('RULES', [])
d = {k: v for k, v in d.items() if k != v}
src = 'export default ' + json.dumps({'dict': dict(sorted(d.items())), 'rules': rules}, ensure_ascii=False, indent=0) + ';\n'
open(os.path.join(root, 'apps', 'shared', 'i18n-en.js'), 'w', encoding='utf-8').write(src)
print(len(d), 'cadenas,', len(rules), 'reglas')
