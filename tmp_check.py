# -*- coding: utf-8 -*-
import json, pathlib
p = pathlib.Path('data/opciones_co2_reglas.json')
data = json.loads(p.read_text(encoding='utf-8'))
bad = []
for b in data.get('bloques', []):
    for r in b.get('reglas', []):
        if not (r.get('item') or '').strip() or not (r.get('pregunta') or '').strip():
            bad.append((b.get('id'), r.get('pregunta'), r.get('item')))
print('count', len(bad))
for e in bad[:10]:
    print(e)
