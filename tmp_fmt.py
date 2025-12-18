import json
from pathlib import Path
p=Path('data/opciones_co2_reglas.json')
data=json.loads(p.read_text(encoding='utf-8'))

def fmt_general(obj, indent=2):
    sp=' '*indent
    if isinstance(obj, dict):
        inline_keys = ("tokens","permitidos","bornera")
        parts=[]
        for k,v in obj.items():
            if k=="tokens" and isinstance(v,list):
                tok='[' + ', '.join(json.dumps(t, ensure_ascii=False, separators=(', ', ': ')) for t in v) + ']'
                parts.append(f"{sp}\"tokens\": {tok}")
            elif k=="permitidos" and isinstance(v,list) and all(isinstance(x,str) for x in v):
                perm='[' + ', '.join(json.dumps(x, ensure_ascii=False) for x in v) + ']'
                parts.append(f"{sp}\"permitidos\": {perm}")
            elif k=="bornera" and isinstance(v,dict):
                parts.append(f"{sp}\"bornera\": " + json.dumps(v, ensure_ascii=False, separators=(', ', ': ')))
            else:
                parts.append(f"{sp}\"{k}\": {fmt_general(v, indent+2)}")
        return '{\n' + ',\n'.join(parts) + '\n' + sp[:-2] + '}'
    if isinstance(obj, list):
        if all(isinstance(x,str) for x in obj):
            return '[' + ', '.join(json.dumps(x, ensure_ascii=False) for x in obj) + ']'
        items=[fmt_general(el, indent+2) for el in obj]
        return '[\n' + ',\n'.join(items) + '\n' + sp[:-2] + ']'
    return json.dumps(obj, ensure_ascii=False)

text = fmt_general(data,2) + '\n'
p.write_text(text, encoding='utf-8')
print('done')
