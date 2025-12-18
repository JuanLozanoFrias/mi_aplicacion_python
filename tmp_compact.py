# -*- coding: utf-8 -*-
import json, pathlib
p=pathlib.Path('data/opciones_co2_reglas.json')
obj=json.loads(p.read_text(encoding='utf-8'))

def fmt(obj, indent=0):
    sp=' ' * indent
    if isinstance(obj, dict):
        parts=[]
        for k,v in obj.items():
            parts.append(f'{sp}"{k}": {fmt(v, indent+2)}')
        return '{\n' + ',\n'.join(parts) + '\n' + sp + '}'
    if isinstance(obj, list):
        # inline simple lists/dicts
        if all(isinstance(x,(str,int,float,bool)) or x is None for x in obj):
            return '[' + ','.join(json.dumps(x, ensure_ascii=False) for x in obj) + ']'
        if all(isinstance(x, dict) and all(isinstance(v,(str,int,float,bool)) or v is None for v in x.values()) for x in obj):
            inner = ','.join('{' + ','.join(f'"{k}":'+json.dumps(v, ensure_ascii=False) for k,v in x.items()) + '}' for x in obj)
            return '[' + inner + ']'
        items = [fmt(x, indent+2) for x in obj]
        return '[\n' + ',\n'.join(items) + '\n' + sp + ']'
    return json.dumps(obj, ensure_ascii=False)

p.write_text(fmt(obj,0), encoding='utf-8')
print('reformatted compact lists')
