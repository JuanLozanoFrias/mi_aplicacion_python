from __future__ import annotations

from pathlib import Path
from typing import List, Dict
from datetime import datetime

import pandas as pd

from .models import Technician, ProductionOrder


def _find_data_dir() -> Path:
    for p in Path(__file__).resolve().parents:
        cand = p / "data" / "produccion"
        if cand.exists():
            return cand
    return Path("data/produccion")


def _find_excel_by_keyword(keyword: str) -> Path | None:
    data_dir = _find_data_dir()
    for f in data_dir.glob("*.xlsx"):
        if keyword.upper() in f.name.upper():
            return f
    return None


def _normalize_cols(df: pd.DataFrame) -> Dict[str, str]:
    cols = {}
    for c in df.columns:
        key = str(c).strip().upper().replace(" ", "").replace(".", "")
        key = key.translate(str.maketrans("ÁÉÍÓÚÑ", "AEIOUN"))
        cols[key] = c
    return cols


def load_personal_excel() -> List[Technician]:
    path = _find_excel_by_keyword("CONTROL PERSONAL")
    if not path or not path.exists():
        return []
    try:
        df_raw = pd.read_excel(path, sheet_name="20251031", header=None)
    except Exception:
        try:
            df_raw = pd.read_excel(path, sheet_name=0, header=None)
        except Exception:
            return []

    header_row = None
    for idx, row in df_raw.iterrows():
        for cell in row.tolist():
            if str(cell).strip().upper() == "NUMERO":
                header_row = idx
                break
        if header_row is not None:
            break
    if header_row is None:
        return []

    df = pd.read_excel(path, sheet_name="20251031", header=header_row)
    cols = _normalize_cols(df)
    def col(name: str) -> str | None:
        return cols.get(name)

    c_num = col("NUMERO")
    c_name = col("ELECTRICISTA")
    c_ced = col("CEDULA")
    c_turn = col("TURNO")
    c_area = col("AREA")
    c_op = col("OP")
    c_act = col("ACTIVIDAD")
    c_obs = col("OBSERVACIONES")

    techs: List[Technician] = []
    for _, r in df.iterrows():
        num = str(r.get(c_num, "")).strip()
        if not num:
            break
        techs.append(
            Technician(
                number=num,
                name=str(r.get(c_name, "")).strip(),
                cedula=str(r.get(c_ced, "")).strip(),
                shift=str(r.get(c_turn, "")).strip(),
                area=str(r.get(c_area, "")).strip(),
                op=str(r.get(c_op, "")).strip(),
                activity=str(r.get(c_act, "")).strip(),
                notes=str(r.get(c_obs, "")).strip(),
                status="DISPONIBLE",
                productivity=0.0,
            )
        )
    return techs


def load_orders_excel() -> List[ProductionOrder]:
    path = _find_excel_by_keyword("ORDENES_PRODUCCION")
    if not path:
        # fallback: first xlsx not personal
        data_dir = _find_data_dir()
        for f in data_dir.glob("*.xlsx"):
            if "CONTROL PERSONAL" not in f.name.upper():
                path = f
                break
    if not path or not path.exists():
        return []
    try:
        df = pd.read_excel(path, sheet_name="Sheet1")
    except Exception:
        try:
            df = pd.read_excel(path, sheet_name=0)
        except Exception:
            return []

    cols = _normalize_cols(df)
    def c(name: str, *alts: str) -> str | None:
        if name in cols:
            return cols[name]
        for a in alts:
            if a in cols:
                return cols[a]
        return None

    c_op = c("OPNUMERO", "OPNUMERO", "OPNUM", "OPNÚMERO", "OPNÚMERO", "OPNUMERO")
    c_fecha = c("FECHADOCTO", "FECHADOCTO", "FECHADOC", "FECHADCTO")
    c_estado = c("ESTADO")
    c_ref1 = c("OPREFERENCIA1", "OPREFERENCIA1", "REFERENCIA1")
    c_ref2 = c("OPREFERENCIA2", "OPREFERENCIA2", "REFERENCIA2")
    c_notas = c("NOTAS")
    c_cats = c("CATEGORIASELECTRICAS", "CATEGORIASELECTRICAS", "CATEGORIAS_ELECTRICAS")

    orders: List[ProductionOrder] = []
    for _, r in df.iterrows():
        op = str(r.get(c_op, "")).strip()
        if not op:
            continue
        fecha = r.get(c_fecha, "")
        if isinstance(fecha, datetime):
            fecha = fecha.strftime("%Y-%m-%d")
        orders.append(
            ProductionOrder(
                op_number=op,
                date=str(fecha).strip(),
                state=str(r.get(c_estado, "")).strip() or "ABIERTA",
                ref1=str(r.get(c_ref1, "")).strip(),
                ref2=str(r.get(c_ref2, "")).strip(),
                notes=str(r.get(c_notas, "")).strip(),
                categories=str(r.get(c_cats, "")).strip(),
                avance_pct=0.0,
                priority="MEDIA",
            )
        )
    return orders
