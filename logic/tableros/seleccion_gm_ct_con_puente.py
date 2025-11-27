# -*- coding: utf-8 -*-
# logic/seleccion_gm_ct_con_puente.py
"""
Selecciona *en conjunto* GUARDAMOTOR y CONTACTOR para ABB
exigiendo que exista PUENTE válido: Q(GM) == Q(CT) y Q != "".
El código interno del puente (R) solo se devuelve para el Paso 4.

Depende de:
- logic.guardamotor.listar_guardamotores_abb
- logic.contactor.listar_contactores_abb
"""

from __future__ import annotations
from typing import Any, Dict

from logic.guardamotor import listar_guardamotores_abb
from logic.contactor import listar_contactores_abb


def seleccionar_par_gm_ct_con_puente_abb(
    excel_path: str,
    corriente_a: float,
    tipo_arranque: str,
) -> Dict[str, Any]:
    """
    Recorre candidatos por prioridad eléctrica (como vienen ordenados)
    y devuelve el primer par (GM, CT) con Q coincidente.

    return {
      "aplica": True/False,
      "guardamotor": {"modelo":..., "cantidad":..., "puente_modelo":..., "puente_codigo":...},
      "contactor":   {"modelo":..., "cantidad":..., "puente_modelo":..., "puente_codigo":...},
      "puente":      {"modelo": <Q>, "codigo": <R|''>}
    }
    """
    gms = listar_guardamotores_abb(excel_path, corriente_a, tipo_arranque)
    cts = listar_contactores_abb(excel_path, corriente_a, tipo_arranque)

    for gm in gms:
        qg = (gm.get("puente_modelo") or "").strip()
        if not qg:
            continue
        for ct in cts:
            qc = (ct.get("puente_modelo") or "").strip()
            if not qc:
                continue
            if qg == qc:
                return {
                    "aplica": True,
                    "guardamotor": gm,
                    "contactor": ct,
                    "puente": {
                        "modelo": qg,
                        "codigo": (gm.get("puente_codigo") or ct.get("puente_codigo") or ""),
                    },
                }

    return {"aplica": False, "motivo": "Sin pareja GM–CT con puente (Q) coincidente."}

