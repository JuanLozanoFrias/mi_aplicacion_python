# export_siesa_to_package.py
from __future__ import annotations

import json
import os
import hashlib
import datetime
from pathlib import Path
from typing import Any

from siesa_uno_client import SiesaConfig, SiesaUNOClient, CIA_MAP, SQL_FN_INVENTARIO

REQUIRED_COLUMNS = [
    "Item",
    "Referencia",
    "Desc_corta_item",
    "Desc_item",
    "Unidad_inventario",
    "Unidad_orden",
    "Tipo_inv_serv",
    "Cant_existencia",
    "Cant_requerida",
    "Cant_OC_OP",
    "Fecha_creacion",
    "Estado",
    "Notas",
]

_STRIP_COLUMNS = [
    "Referencia",
    "Unidad_inventario",
    "Unidad_orden",
    "Tipo_inv_serv",
    "Estado",
]

_ORDER_STRIP_COLUMNS = [
    "OpNumero",
    "Estado",
    "OpReferencia1",
    "OpReferencia2",
    "Notas",
    "CategoriasElectricas",
]


def _sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _dump_json(path: str, obj: Any) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2, default=str)


def _env(name: str, default: str | None = None) -> str | None:
    val = os.getenv(name)
    if val is None or val.strip() == "":
        return default
    return val


def _env_bool(name: str, default: bool) -> bool:
    val = _env(name)
    if val is None:
        return default
    return val.strip().lower() in {"1", "true", "yes", "y"}


def _normalize_col(name: str) -> str:
    return (
        str(name).strip().lower()
        .replace(".", "")
        .replace("/", "_")
        .replace("__", "_")
        .replace(" ", "_")
    )


def _rename_columns(df) -> None:
    mapping = {}
    for col in df.columns:
        mapping[_normalize_col(col)] = col

    def pick(*keys: str) -> str | None:
        for k in keys:
            if k in mapping:
                return mapping[k]
        return None

    rename_map = {}
    rename_map[pick("item")] = "Item"
    rename_map[pick("referencia", "ref", "codigo")] = "Referencia"
    rename_map[pick("desc_corta_item", "desc_corta", "desc_corta_item")] = "Desc_corta_item"
    rename_map[pick("desc_item", "descripcion", "descripcion_item")] = "Desc_item"
    rename_map[pick("unidad_inventario", "unidad_inv")] = "Unidad_inventario"
    rename_map[pick("unidad_orden", "unidad_ord")] = "Unidad_orden"
    rename_map[pick("tipo_inv_serv", "tipo_inv", "tipo_servicio")] = "Tipo_inv_serv"
    rename_map[pick("cant_existencia", "existencia")] = "Cant_existencia"
    rename_map[pick("cant_requerida", "requerida")] = "Cant_requerida"
    rename_map[pick("cant_oc_op", "cant_oc", "cant_op")] = "Cant_OC_OP"
    rename_map[pick("fecha_creacion", "fecha")] = "Fecha_creacion"
    rename_map[pick("estado")] = "Estado"
    rename_map[pick("notas", "nota")] = "Notas"

    rename_map = {k: v for k, v in rename_map.items() if k}
    if rename_map:
        df.rename(columns=rename_map, inplace=True)


def _clean_strings(df) -> None:
    for col in _STRIP_COLUMNS:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()


def _clean_order_strings(df) -> None:
    for col in _ORDER_STRIP_COLUMNS:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()


def _read_sql_file(path: Path) -> str | None:
    if not path.exists():
        return None
    try:
        raw = path.read_text(encoding="utf-8")
    except Exception:
        return None
    lines = []
    for line in raw.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("--"):
            continue
        lines.append(line)
    sql = "\n".join(lines).strip()
    return sql or None


def _get_orders_sql() -> str | None:
    sql_env = _env("SIESA_ORDERS_SQL")
    if sql_env:
        return sql_env
    sql_path = _env("SIESA_ORDERS_SQL_PATH")
    if sql_path:
        return _read_sql_file(Path(sql_path))
    default_path = Path(__file__).resolve().parent / "sql" / "ordenes_produccion.sql"
    return _read_sql_file(default_path)


def _rename_orders_columns(df) -> None:
    mapping = {}
    for col in df.columns:
        mapping[_normalize_col(col)] = col

    def pick(*keys: str) -> str | None:
        for k in keys:
            if k in mapping:
                return mapping[k]
        return None

    rename_map = {}
    rename_map[pick("opnumero", "op_numero", "opnum", "op")] = "OpNumero"
    rename_map[pick("fechadocto", "fecha_docto", "fechadoc", "fecha")] = "FechaDocto"
    rename_map[pick("estado", "status")] = "Estado"
    rename_map[pick("opreferencia1", "referencia1", "ref1")] = "OpReferencia1"
    rename_map[pick("opreferencia2", "referencia2", "ref2")] = "OpReferencia2"
    rename_map[pick("notas", "nota", "observaciones", "obs")] = "Notas"
    rename_map[pick("categoriaselectricas", "categorias_electricas", "categorias")] = "CategoriasElectricas"

    rename_map = {k: v for k, v in rename_map.items() if k}
    if rename_map:
        df.rename(columns=rename_map, inplace=True)


def _find_orders_excel(base_dir: Path) -> Path | None:
    if not base_dir.exists():
        return None
    for f in base_dir.glob("*.xlsx"):
        if "ORDENES_PRODUCCION" in f.name.upper():
            return f
    for f in base_dir.glob("*.xlsx"):
        if "CONTROL PERSONAL" not in f.name.upper():
            return f
    return None


def _load_orders_excel_df():
    base_dir = Path(__file__).resolve().parents[2] / "data" / "produccion"
    path = _find_orders_excel(base_dir)
    if not path:
        return None, None
    try:
        import pandas as pd
    except Exception:
        return None, None
    try:
        df = pd.read_excel(path, sheet_name="Sheet1")
    except Exception:
        try:
            df = pd.read_excel(path, sheet_name=0)
        except Exception:
            return None, None
    return df, str(path)


def export_inventory_snapshot(
    out_company_data_dir: str,
    company: str,
    client: SiesaUNOClient,
    sql: str = SQL_FN_INVENTARIO,
) -> str:
    id_cia = CIA_MAP[company]
    df = client.fetch_df(sql, params=[id_cia])
    _rename_columns(df)
    _clean_strings(df)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        print(f"WARNING: Faltan columnas requeridas: {missing}")
        raise SystemExit(1)

    rows = df[REQUIRED_COLUMNS].to_dict(orient="records")
    row_count = len(rows)
    if company.lower() == "weston" and row_count < 5000:
        print(f"WARNING: row_count={row_count} parece bajo para Weston.")

    payload = {
        "meta": {
            "company": company,
            "id_cia": id_cia,
            "source": {"server": client.cfg.server, "database": client.cfg.database},
            "generated_at": datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-5))).isoformat(),
            "row_count": row_count,
            "columns": REQUIRED_COLUMNS,
        },
        "rows": rows,
    }

    rel = f"snapshots/inventory_{company}.json"
    out_path = os.path.join(out_company_data_dir, rel)
    _dump_json(out_path, payload)
    return rel


def export_production_orders_snapshot(
    out_company_data_dir: str,
    company: str,
    client: SiesaUNOClient | None = None,
    sql: str | None = None,
    allow_excel_fallback: bool = True,
) -> str | None:
    id_cia = CIA_MAP.get(company)
    if id_cia is None:
        raise KeyError(f"Company no registrada en CIA_MAP: {company}")

    df = None
    source = {}

    sql_error = None
    if sql:
        if client is None:
            raise ValueError("client requerido para exportar ordenes via SQL.")
        params = [id_cia] if "?" in sql else []
        try:
            df = client.fetch_df(sql, params=params)
            source = {
                "type": "sql",
                "server": client.cfg.server,
                "database": client.cfg.database,
            }
        except Exception as e:
            sql_error = e

    if df is None and allow_excel_fallback:
        df, path = _load_orders_excel_df()
        if df is not None and path:
            source = {"type": "excel", "path": path}
            if sql_error:
                print(f"WARNING: fallo SQL en ordenes, usando Excel -> {sql_error}")

    if df is None:
        if sql_error:
            raise sql_error
        return None

    _rename_orders_columns(df)
    _clean_order_strings(df)

    rows = df.to_dict(orient="records")
    payload = {
        "meta": {
            "company": company,
            "id_cia": id_cia,
            "source": source,
            "generated_at": datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-5))).isoformat(),
            "row_count": len(rows),
            "columns": list(df.columns),
        },
        "rows": rows,
    }

    rel = f"snapshots/production_orders_{company}.json"
    out_path = os.path.join(out_company_data_dir, rel)
    _dump_json(out_path, payload)
    return rel


def build_manifest(out_company_data_dir: str) -> None:
    manifest = {
        "package_name": "weston-company-data",
        "schema_version": "1.0",
        "generated_at": datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-5))).isoformat(),
        "files": [],
    }

    roots = ["master_data", "snapshots", "rules", "audit"]
    for root in roots:
        folder = os.path.join(out_company_data_dir, root)
        if not os.path.isdir(folder):
            continue
        for dirpath, _, filenames in os.walk(folder):
            for fn in filenames:
                if not (fn.endswith(".json") or fn.endswith(".jsonl")):
                    continue
                full = os.path.join(dirpath, fn)
                rel = os.path.relpath(full, out_company_data_dir).replace("\\", "/")
                kind = "master" if rel.startswith("master_data/") else (
                    "snapshot" if rel.startswith("snapshots/") else (
                        "rules" if rel.startswith("rules/") else "audit"
                    )
                )
                manifest["files"].append({
                    "path": rel,
                    "sha256": _sha256_file(full),
                    "kind": kind,
                })

    out_path = os.path.join(out_company_data_dir, "manifests", "manifest.json")
    _dump_json(out_path, manifest)


def main():
    out_dir = os.path.abspath("company_data")
    os.makedirs(os.path.join(out_dir, "manifests"), exist_ok=True)

    auth = (_env("SIESA_AUTH", "credman") or "credman").lower()
    if auth not in {"windows", "sql", "credman"}:
        auth = "windows"
    cfg = SiesaConfig(
        server=_env("SIESA_SERVER", "192.168.155.93") or "192.168.155.93",
        database=_env("SIESA_DATABASE", "UNOEE") or "UNOEE",
        auth=auth,
        user=_env("SIESA_USER"),
        password=_env("SIESA_PASSWORD"),
        cred_target=_env("SIESA_CRED_TARGET", "CalvoSiesaUNOEE"),
        cred_user=_env("SIESA_CRED_USER", "sa"),
        driver=_env("SIESA_DRIVER", "ODBC Driver 18 for SQL Server") or "ODBC Driver 18 for SQL Server",
        encrypt=_env_bool("SIESA_ENCRYPT", False),
        trust_server_certificate=_env_bool("SIESA_TRUST", True),
    )
    client = SiesaUNOClient(cfg)

    try:
        export_inventory_snapshot(out_dir, "Weston", client)
    except SystemExit as e:
        print(f"WARNING: inventario no exportado -> {e}")
    except Exception as e:
        print(f"WARNING: inventario no exportado -> {e}")

    orders_sql = _get_orders_sql()
    try:
        rel_orders = export_production_orders_snapshot(
            out_dir,
            "Weston",
            client=client,
            sql=orders_sql,
            allow_excel_fallback=True,
        )
        if not rel_orders:
            print("INFO: no se encontraron ordenes de produccion para exportar.")
    except Exception as e:
        print(f"WARNING: ordenes de produccion no exportadas -> {e}")
    # export_inventory_snapshot(out_dir, "WBR", client)
    # export_inventory_snapshot(out_dir, "TEKOAM", client)

    build_manifest(out_dir)
    print("OK. Paquete generado en:", out_dir)


if __name__ == "__main__":
    main()
