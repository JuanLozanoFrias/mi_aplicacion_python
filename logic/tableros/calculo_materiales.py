# logic/calculo_materiales.py
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List, Dict, Any, Optional

import pandas as pd


@dataclass
class ParamsMateriales:
    proyecto: str
    ciudad: str
    responsable: str

    tension_alimentacion: str        # "", "220", "460"
    refrigerante: str                # "", "R744", "R290", "R507", "R404"
    tension_control: str             # "", "120", "220"
    norma_aplicable: str             # "", "UL", "IEC"
    tipo_compresores: str            # "", "BITZER", "COPELAND", "TECUMSEH", "FRASCOLD", "DORIN", "OTRO"
    n_comp_media: int
    n_comp_paralelo: int
    n_comp_baja: int
    tension_gascooler: str           # "", "220", "460", "120"
    backup: bool
    programador_aks800: bool
    medidores_energia: bool
    marca_elementos: str             # "", "ABB", "SIEMENS", "SCHNEIDER", "LS", "WEG", "CHINT", "DELIXY"
    marca_variadores: Optional[str]  # "", "NO", "ABB", "DANFOSS", "SCHNEIDER", "YASKAWA", "DELTA", "CHINT", "DELIXY"
    desuperheater: bool


def catalogo_materiales(path: Path | None = None, sheet: str = "MATERIALES") -> List[str]:
    """Lista de descripciones para combos (col A). Si falla, usa un catálogo mínimo."""
    book = path or Path(__file__).resolve().parents[2] / "data" / "tableros_electricos" / "basedatos.xlsx"
    try:
        df = pd.read_excel(book, sheet_name=sheet, usecols="A", header=None)
        items = df.iloc[:, 0].dropna().astype(str).str.strip().unique().tolist()
        if items:
            return items
    except Exception:
        pass
    return [
        "Gabinete metálico 600x800x300 NEMA 1",
        "Gabinete acero inox 600x800x300 NEMA 4X",
        "Riel DIN 35 mm x 2 m",
        "Canaleta 40x60 mm (2 m)",
        "Bornera paso 10 mm²",
        "Etiquetas para bornes (tirilla)",
        "Interruptor principal 125 A 3P",
        "Interruptor principal 250 A 3P",
        "Seccionador 3P con manija externa",
        "Contactor 18 A bobina 220V",
        "Relé térmico 18 A",
        "Guardamotor 32 A curva C",
        "Barraje cobre 20x3 mm (1 m)",
        "Ventilador 120 mm 220V",
        "Regleta servicios 110V",
        "Luz interior LED 110-230V",
        "Programador AKS800",
        "Medidor de energía trifásico",
        "Variador de velocidad 7.5 kW",
        "Desuperheater (kit)",
    ]


def _kit_base() -> List[Dict[str, Any]]:
    return [
        {"ITEM": 1, "DESCRIPCION": "GABINETE METÁLICO 600X800X300 NEMA 1", "CANTIDAD": 1, "UND": "UND", "NOTAS": ""},
        {"ITEM": 2, "DESCRIPCION": "RIEL DIN 35 MM X 2 M",                 "CANTIDAD": 2, "UND": "UND", "NOTAS": ""},
        {"ITEM": 3, "DESCRIPCION": "CANALETA 40X60 MM (2 M)",              "CANTIDAD": 3, "UND": "UND", "NOTAS": ""},
        {"ITEM": 4, "DESCRIPCION": "LUZ INTERIOR LED 110-230V",            "CANTIDAD": 1, "UND": "UND", "NOTAS": ""},
        {"ITEM": 5, "DESCRIPCION": "REGLETA SERVICIOS 110V",               "CANTIDAD": 1, "UND": "UND", "NOTAS": ""},
        {"ITEM": 6, "DESCRIPCION": "PRENSAESTOPA M25",                     "CANTIDAD": 6, "UND": "UND", "NOTAS": ""},
    ]


def _principal_estimada(n_total_comp: int) -> Dict[str, Any]:
    """Estimación grosera: ~25A por compresor para escoger principal."""
    amps = max(1, n_total_comp) * 25
    if amps <= 125:
        desc = "INTERRUPTOR PRINCIPAL 125 A 3P"
    elif amps <= 250:
        desc = "INTERRUPTOR PRINCIPAL 250 A 3P"
    else:
        desc = "SECCIONADOR 3P CON MANIJA EXTERNA"
    return {"ITEM": 7, "DESCRIPCION": desc, "CANTIDAD": 1, "UND": "UND", "NOTAS": ""}


def _por_alimentador(control_v: str, marca_elem: str) -> List[Dict[str, Any]]:
    bobina = "".join(ch for ch in control_v if ch.isdigit()) or control_v or "220"
    marca = f" ({marca_elem})" if marca_elem else ""
    return [
        {"DESCRIPCION": f"BORNERA PASO 10 MM²{marca}",             "UND": "UND"},
        {"DESCRIPCION": "ETIQUETAS PARA BORNES (TIRILLA)",         "UND": "UND"},
        {"DESCRIPCION": f"CONTACTOR 18 A BOBINA {bobina}V{marca}", "UND": "UND"},
        {"DESCRIPCION": f"RELÉ TÉRMICO 18 A{marca}",               "UND": "UND"},
        {"DESCRIPCION": f"GUARDAMOTOR 32 A CURVA C{marca}",        "UND": "UND"},
    ]


def calcular_bom(p: ParamsMateriales) -> List[Dict[str, Any]]:
    bom: List[Dict[str, Any]] = []
    bom += _kit_base()

    total_comp = p.n_comp_media + p.n_comp_baja + p.n_comp_paralelo
    bom.append(_principal_estimada(total_comp))

    repetitivos = _por_alimentador(p.tension_control, p.marca_elementos)
    base_item = 100
    n_items = max(total_comp, 1)
    for i in range(1, n_items + 1):
        for j, r in enumerate(repetitivos, start=0):
            bom.append({
                "ITEM": base_item + (i - 1) * len(repetitivos) + j + 1,
                "DESCRIPCION": r["DESCRIPCION"],
                "CANTIDAD": 1,
                "UND": r["UND"],
                "NOTAS": f"ALIM {i}",
            })

    if p.programador_aks800:
        bom.append({"ITEM": 800, "DESCRIPCION": "PROGRAMADOR AKS800", "CANTIDAD": 1, "UND": "UND", "NOTAS": ""})
    if p.medidores_energia:
        bom.append({"ITEM": 810, "DESCRIPCION": "MEDIDOR DE ENERGÍA TRIFÁSICO", "CANTIDAD": 1, "UND": "UND", "NOTAS": ""})
    if p.desuperheater:
        bom.append({"ITEM": 820, "DESCRIPCION": "DESUPERHEATER (KIT)", "CANTIDAD": 1, "UND": "UND", "NOTAS": ""})
    if p.marca_variadores and p.marca_variadores not in ("", "NO"):
        bom.append({"ITEM": 830, "DESCRIPCION": f"VARIADOR DE VELOCIDAD 7.5 KW ({p.marca_variadores})",
                    "CANTIDAD": max(total_comp, 1), "UND": "UND", "NOTAS": "SUGERIDO"})
    if (p.norma_aplicable or "").upper().startswith("UL"):
        bom.append({"ITEM": 900, "DESCRIPCION": "BARRAJE COBRE 20X3 MM (1 M)", "CANTIDAD": 2, "UND": "UND", "NOTAS": "RECOMENDACIÓN UL"})
        bom.append({"ITEM": 901, "DESCRIPCION": "VENTILADOR 120 MM 220V", "CANTIDAD": 2, "UND": "UND", "NOTAS": "FLUJO FORZADO"})

    bom.append({"ITEM": 990, "DESCRIPCION": f"REFRIGERANTE: {p.refrigerante or '-'}", "CANTIDAD": 0, "UND": "", "NOTAS": ""})
    bom.append({"ITEM": 991, "DESCRIPCION": f"COMPRESORES M/P/B: {p.n_comp_media}/{p.n_comp_paralelo}/{p.n_comp_baja} — TIPO: {p.tipo_compresores or '-'}", "CANTIDAD": 0, "UND": "", "NOTAS": ""})
    if p.backup:
        bom.append({"ITEM": 992, "DESCRIPCION": "PREVISIÓN CIRCUITO BACK UP", "CANTIDAD": 0, "UND": "", "NOTAS": ""})

    bom.sort(key=lambda x: x.get("ITEM", 99999))
    return bom


