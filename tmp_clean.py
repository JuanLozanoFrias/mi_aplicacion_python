# -*- coding: utf-8 -*-
import json, pathlib
p = pathlib.Path('data/opciones_co2_reglas.json')
data = json.loads(p.read_text(encoding='utf-8'))
ids_bad = {
    'SISTEMA RILINE RITTAL',
    'ROUTER TENDA',
    'SV RILINE60 SOPORTE DE BARRAS',
    'SV RILINE60 BANDEJA BASE SECCION',
    'SV RILINE60 CUBIERTA FINA 9340070',
    'SV ADAPTADOR CONEX 800A-690 9342280',
    'BARRA COBRE P/TAB BARRA NAL001 30X10XMTR',
}
orig = len(data.get('bloques', []))
data['bloques'] = [b for b in data.get('bloques', []) if b.get('id') not in ids_bad]
removed = orig - len(data['bloques'])
p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
print('removed', removed, 'bloques temporales')
