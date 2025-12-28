from __future__ import annotations

from pathlib import Path
from typing import Dict, List

from .models import LegendConfig, PlantillaBOMItem
from .folder_loader import (
    load_config,
    load_equipos,
    load_plantillas,
    load_usos,
    load_variadores,
    load_wcr,
)


ROOT_DIR = Path(__file__).resolve().parents[2]
LEGEND_DIR = ROOT_DIR / "data" / "LEGEND"


class LegendJDService:
    def __init__(self, base_dir: Path | None = None):
        self.base_dir = base_dir or LEGEND_DIR

    def resolve_data_dir(self) -> Path:
        return self.base_dir

    def load_all(self) -> Dict[str, object]:
        base = self.resolve_data_dir()
        files_found: List[str] = []

        cfg, cfg_found = load_config(base)
        if cfg_found:
            files_found.append("config.json")

        equipos, eq_found = load_equipos(base)
        if eq_found:
            files_found.append("equipos.json")

        usos, usos_found = load_usos(base)
        if usos_found:
            files_found.append("usos.json")

        variadores, var_found = load_variadores(base)
        if var_found:
            files_found.append("variadores.json")

        wcr, wcr_found = load_wcr(base)
        if wcr_found:
            files_found.append("wcr.json")

        plantillas, pla_found = load_plantillas(base)
        if pla_found:
            files_found.append("plantillas.json")

        return {
            "config": cfg,
            "equipos": equipos,
            "usos": usos,
            "variadores": variadores,
            "wcr": wcr,
            "plantillas": plantillas,
            "sources": {
                "folder": str(base),
                "files_found": files_found,
            },
        }


if __name__ == "__main__":
    svc = LegendJDService()
    data = svc.load_all()
    cfg = data["config"]
    if not data["sources"]["files_found"]:
        print("OK (LEGEND) dataset vacío")
    else:
        print(f"Proyecto: {getattr(cfg, 'proyecto', None)} | Ciudad: {getattr(cfg, 'ciudad', None)} | Refrigerante: {getattr(cfg, 'refrigerante', None)} | Tcond: {getattr(cfg, 'tcond', None)}")
    print(f"Equipos: {len(data['equipos'])} | Usos BT: {len(data['usos'].get('BT', []))} | Usos MT: {len(data['usos'].get('MT', []))}")
    print(f"Variadores: {len(data['variadores'])} | WCR: {len(data['wcr'])}")
    print({k: len(v) for k, v in data['plantillas'].items()})
    print(f"Sources: {data['sources']}")
