import json
from pathlib import Path
path = Path('data/opciones_co2_reglas.json')
if not path.exists():
    raise SystemExit('no existe opciones_co2_reglas.json')
data = json.loads(path.read_text(encoding='utf-8'))

def dump(obj, indent=0):
    sp = ' ' * indent
    if isinstance(obj, dict):
        items = []
        for k, v in obj.items():
            items.append(f"{sp}  \"{k}\": {dump(v, indent+2)}")
        return '{\n' + ',\n'.join(items) + f'\n{sp}}}'
    if isinstance(obj, list):
        if all(isinstance(x, str) for x in obj) and len(obj) <= 12:
            inner = ', '.join(json.dumps(x, ensure_ascii=False) for x in obj)
            return '[' + inner + ']'
        return '[\n' + ',\n'.join(sp + '  ' + dump(x, indent+2) for x in obj) + '\n' + sp + ']'
    return json.dumps(obj, ensure_ascii=False)

new_text = dump(data, 0) + '\n'
path.write_text(new_text, encoding='utf-8')
print('reformatted', len(new_text))
