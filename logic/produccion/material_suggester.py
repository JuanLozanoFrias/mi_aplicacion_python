from __future__ import annotations

from typing import List, Dict

from .models import MaterialItem


_TEMPLATES: Dict[str, List[MaterialItem]] = {
    "CONTROL": [
        MaterialItem("CTRL-1", "11001", "BORNERA", "BORNERA 2.5MM", 6, "CONTROL"),
        MaterialItem("CTRL-2", "11002", "CANALETA", "CANALETA 40X40", 3, "CONTROL"),
        MaterialItem("CTRL-3", "11003", "RIEL DIN", "RIEL DIN 35MM", 2, "CONTROL"),
        MaterialItem("CTRL-4", "11004", "FUENTE", "FUENTE 24VDC", 1, "CONTROL"),
    ],
    "RESISTENCIAS": [
        MaterialItem("RES-1", "12001", "CONTACTOR", "CONTACTOR 3P", 2, "RESISTENCIAS"),
        MaterialItem("RES-2", "12002", "BREAKER", "BREAKER 3P", 2, "RESISTENCIAS"),
        MaterialItem("RES-3", "12003", "CABLE POTENCIA", "CABLE POTENCIA 3X10", 20, "RESISTENCIAS"),
    ],
    "ILUMINACION": [
        MaterialItem("ILU-1", "13001", "LÁMPARA", "LÁMPARA LED", 6, "ILUMINACION"),
        MaterialItem("ILU-2", "13002", "DRIVER", "DRIVER LED", 2, "ILUMINACION"),
        MaterialItem("ILU-3", "13003", "BREAKER", "BREAKER 1P", 2, "ILUMINACION"),
    ],
    "VARIADOR": [
        MaterialItem("VFD-1", "14001", "VFD", "VARIADOR FRECUENCIA", 1, "VARIADOR"),
        MaterialItem("VFD-2", "14002", "BREAKER", "BREAKER 3P", 1, "VARIADOR"),
        MaterialItem("VFD-3", "14003", "CABLE APANTALLADO", "CABLE APANTALLADO", 15, "VARIADOR"),
    ],
    "COMPRESOR": [
        MaterialItem("CMP-1", "15001", "BORNERA", "BORNERA POTENCIA", 8, "COMPRESOR"),
        MaterialItem("CMP-2", "15002", "GUARDAMOTOR", "GUARDAMOTOR 3P", 2, "COMPRESOR"),
    ],
}


def suggest_materials(categories: str) -> List[MaterialItem]:
    if not categories:
        return []
    items: List[MaterialItem] = []
    seen = {}
    for raw in categories.split(","):
        cat = raw.strip().upper()
        if not cat:
            continue
        tpl = _TEMPLATES.get(cat, [])
        for it in tpl:
            key = (it.code, it.name)
            if key in seen:
                seen[key].qty += it.qty
            else:
                copy = MaterialItem(
                    it.item, it.code, it.name, it.description, it.qty, it.category
                )
                seen[key] = copy
    items.extend(seen.values())
    return items
