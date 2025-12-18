import json
from pathlib import Path
p=Path('data/opciones_co2_reglas.json')
data=json.loads(p.read_text(encoding='utf-8'))
counts={}
locs={}
for b in data.get('bloques', []):
    bid=b.get('id','?')
    for r in b.get('reglas', []):
        rq = r.get('pregunta','?')
        item=r.get('item') or ''
        if item:
            counts[item]=counts.get(item,0)+1
            locs.setdefault(item, []).append(f"{bid}:{rq}")
        for alt in r.get('alternativas',[]):
            it=alt.get('item')
            if it:
                counts[it]=counts.get(it,0)+1
                locs.setdefault(it, []).append(f"{bid}:{rq} (alt)")
dups={k:v for k,v in counts.items() if v>1}
for item,count in sorted(dups.items()):
    print(f"{item}: {count} -> {locs[item]}")
