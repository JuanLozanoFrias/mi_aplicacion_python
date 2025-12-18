import json, pathlib
p=pathlib.Path('data/opciones_co2_reglas.json')
obj=json.loads(p.read_text(encoding='utf-8'))

def ser_inline(v):
    import json as _j
    if isinstance(v, (str,int,float)) or v is None:
        return _j.dumps(v, ensure_ascii=False)
    if isinstance(v, bool):
        return 'true' if v else 'false'
    return None

def ser_dict_inline(d):
    parts=[]
    for k,v in d.items():
        sv=ser_inline(v)
        if sv is None:
            sv=ser_any(v, inline_ok=True)
        parts.append(f'"{k}":{sv}')
    return '{' + ','.join(parts) + '}'

def ser_list_inline(lst):
    if all(isinstance(x,(str,int,float,bool)) or x is None for x in lst):
        return '[' + ','.join(ser_inline(x) for x in lst) + ']'
    if all(isinstance(x, dict) and all(isinstance(v,(str,int,float,bool)) or v is None for v in x.values()) for x in lst):
        return '[' + ','.join(ser_dict_inline(x) for x in lst) + ']'
    return None

def ser_any(v, indent=0, inline_ok=False):
    sp=' ' * indent
    if isinstance(v, dict):
        if inline_ok and all(isinstance(val,(str,int,float,bool)) or val is None for val in v.values()):
            return ser_dict_inline(v)
        lines=[]
        for k,val in v.items():
            lines.append(f'{sp}  "{k}": {ser_any(val, indent+2, inline_ok=True)}')
        return '{\n' + ',\n'.join(lines) + '\n'+sp+'}'
    if isinstance(v, list):
        inline=ser_list_inline(v)
        if inline is not None:
            return inline
        lines=[]
        for item in v:
            lines.append(ser_any(item, indent+2, inline_ok=True))
        return '[' + ',\n'.join('\n'+(' '*(indent+2))+line if '\n' in line else '\n'+(' '*(indent+2))+line for line in lines) + '\n'+sp+']'
    si=ser_inline(v)
    if si is not None:
        return si
    return 'null'

formatted=ser_any(obj, indent=0, inline_ok=False)
p.write_text(formatted, encoding='utf-8')
print('compact format written')
