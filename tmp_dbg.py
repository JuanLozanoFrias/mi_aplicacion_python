from pathlib import Path
import sys
sys.path.append('.')
from logic.tableros.step4_elementos_fijos import cargar_elementos_fijos
from types import SimpleNamespace
base = Path('data/basedatos.xlsx')
print('base exists', base.exists())
if not base.exists():
    raise SystemExit('base not found')
ctx = SimpleNamespace(
    norma_ap='IEC',
    marca_elementos='ABB',
    t_ctl='220',
    t_alim='460',
    step3_state={'UL': {'value': False}},
)
out, tot, dbg = cargar_elementos_fijos(base, ctx)
mask = dbg['NOMBRE'].str.contains('PROTECTOR DE FASE', case=False, na=False)
print(dbg[mask][['ROW','NOMBRE','ESTADO','DETALLE']].head(10).to_string(index=False))
