# logic/step4_types.py
from __future__ import annotations
from dataclasses import dataclass
from typing import Optional

@dataclass
class ContextoStep4:
    """
    Contexto del Paso 4 para decisiones de marca, norma y tensiones.
    """
    marca_elementos: str            # p.ej. "ABB" | "SCHNEIDER" | "GENERICO"
    norma_ap: str                   # "UL" o "IEC"
    tension_control: str            # "120" o "220" (solo dígitos como texto)
    tension_alimentacion: str       # opcional para filtros adicionales
    ul_bool: bool                   # True si norma_ap = UL
    tipo_arranque: str              # 'V', 'P', 'D' (mapeo tolerante)
    backup: bool                    # si aplica condición BACK UP / RESERVA

    @staticmethod
    def normalizar_arranque(valor: str) -> str:
        v = (valor or "").strip().upper()
        mapa = {
            "V": "V", "VARIADOR": "V", "VFD": "V", "INVERTER": "V",
            "P": "P", "PARTIDO": "P", "PART-WINDING": "P", "ESTRELLA-TRIANGULO": "P",
            "D": "D", "DIRECTO": "D", "ARRANQUE DIRECTO": "D"
        }
        return mapa.get(v, v[:1])
