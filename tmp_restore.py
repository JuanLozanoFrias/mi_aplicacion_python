# -*- coding: utf-8 -*-
import json, pathlib, pandas as pd, re
ids_target = {
    'SISTEMA RILINE RITTAL',
    'ROUTER TENDA',
    'SV RILINE60 SOPORTE DE BARRAS',
    'SV RILINE60 BANDEJA BASE SECCION',
    'SV RILINE60 CUBIERTA FINA 9340070',
    'SV ADAPTADOR CONEX 800A-690 9342280',
    'BARRA COBRE P/TAB BARRA NAL001 30X10XMTR',
}
book = pathlib.Path('data/basedatos.xlsx')
df = pd.read_excel(book, sheet_name='OPCIONES CO2', header=None, dtype=str).fillna('')
blocks = []
for i in range(df.shape[0]):
    nombre = str(df.iat[i,29]) if df.shape[1]>29 else ''
    if nombre.strip() not in ids_target:
        continue
    pregunta = str(df.iat[i,0]).strip() or 'SIEMPRE'
    formula = str(df.iat[i,28]).strip() if df.shape[1]>28 else ''
    item = str(df.iat[i,25]).strip() if df.shape[1]>25 else ''
    alt = {
        'tokens': [],
        'permitidos': [],
    }
    if formula:
        m = re.search(r"\*\s*(#|\d+)", formula)
        if m:
            alt['mult'] = '#' if m.group(1)=='#' else m.group(1)
            formula = re.sub(r"\*\s*(#|\d+)", "", formula)
        formula = re.sub(r"^\s*MARCA\s+DE\s+ELEMENTOS\s*:\s*", "", formula, flags=re.I)
        parts = formula.split('=',1)
        conds = parts[0].strip()
        perms = [p.strip() for p in parts[1].split(',') if p.strip()] if len(parts)>1 else []
        alt['permitidos'] = perms
        for c,v in re.findall(r"([A-Z]+)\(\s*(.*?)\s*\)", conds):
            if (v.startswith('"') and v.endswith('"')) or (v.startswith("'") and v.endswith("'")):
                v=v[1:-1]
            alt.setdefault('tokens',[]).append({'op':c.strip(),'val':v})
    regla = {
        'pregunta': pregunta,
        'accion': 'pone',
        'item': item,
        'accion_no': 'borra',
        'item_no': '',
        'alternativas': [alt],
        'allow_multi': False,
    }
    if nombre:
        regla['nombre']=nombre
    blocks.append({'id': nombre or pregunta or f'REGLA_{i}', 'reglas':[regla]})

json_path = pathlib.Path('data/opciones_co2_reglas.json')
data = json.loads(json_path.read_text(encoding='utf-8'))
# remove any existing with those ids
ids_set = {b['id'] for b in blocks}
data['bloques'] = [b for b in data.get('bloques',[]) if b.get('id') not in ids_set]
data['bloques'].extend(blocks)
json_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
print('readded', len(blocks), 'bloques')
