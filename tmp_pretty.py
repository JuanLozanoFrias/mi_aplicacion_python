# -*- coding: utf-8 -*-
import json, pathlib
p = pathlib.Path('data/opciones_co2_reglas.json')
data = json.loads(p.read_text(encoding='utf-8'))
p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
print('restored pretty indent')
