import json, re, pathlib, pandas as pd
book = pathlib.Path('data/basedatos.xlsx')
json_path = pathlib.Path('data/opciones_co2_reglas.json')
start,end = 59,113
# load excel
if not book.exists():
    raise SystemExit('no basedatos.xlsx')
df = pd.read_excel(book, sheet_name='OPCIONES CO2', header=None, dtype=str).fillna('')
new_blocks = []
for i in range(start, end):
    pregunta_ad = str(df.iat[i, 29]).strip() if df.shape[1] > 29 else ''
    pregunta_col0 = str(df.iat[i, 0]).strip()
    formula = str(df.iat[i, 28]).strip() if df.shape[1] > 28 else ''
    item = str(df.iat[i, 25]).strip() if df.shape[1] > 25 else ''
    nombre = str(df.iat[i, 27]).strip() if df.shape[1] > 27 else ''
    if not formula and not item:
        continue
    pregunta = pregunta_ad or pregunta_col0 or 'SIEMPRE'
    mult = None
    m = re.search(r"\*\s*(#|\d+)", formula)
    if m:
        mult = m.group(1)
        formula = re.sub(r"\*\s*(#|\d+)", "", formula)
    formula = re.sub(r"^\s*MARCA\s+DE\s+ELEMENTOS\s*:\s*", "", formula, flags=re.I)
    parts = formula.split('=', 1)
    conds = parts[0].strip()
    permisos = [p.strip() for p in parts[1].split(',') if p.strip()] if len(parts) > 1 else []
    tokens = []
    for c, v in re.findall(r"([A-Z]+)\(\s*(.*?)\s*\)", conds):
        if (v.startswith('"') and v.endswith('"')) or (v.startswith("'") and v.endswith("'")):
            v = v[1:-1]
        tokens.append({'op': c.strip(), 'val': v})
    alt = {'tokens': tokens, 'permitidos': permisos}
    if mult:
        alt['mult'] = '#' if mult == '#' else mult
    regla = {
        'pregunta': pregunta,
        'accion': 'pone',
        'item': item,
        'accion_no': 'borra',
        'item_no': '',
        'alternativas': [alt],
    }
    if nombre:
        regla['nombre'] = nombre
    new_blocks.append({'id': nombre or pregunta or f'REGLA_{i}', 'reglas': [regla]})

# load existing json
data = {'version':1, 'bloques':[]}
try:
    data = json.loads(json_path.read_text(encoding='utf-8'))
    if 'bloques' not in data:
        data['bloques'] = []
except Exception:
    pass
# remove blocks with same id as new ones
ids_new = {b['id'] for b in new_blocks}
data['bloques'] = [b for b in data.get('bloques', []) if b.get('id') not in ids_new]
data['bloques'].extend(new_blocks)
json_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
print('replaced', len(ids_new), 'bloques')
