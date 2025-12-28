from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List

from .models import LegendConfig, PlantillaBOMItem
from .reader import (
    _load_wb,
    read_equipos,
    read_info_config,
    read_plantillas_bom,
    read_usos,
    read_variadores_fc102,
    read_wcr,
)


ROOT_DIR = Path(__file__).resolve().parents[2]
LEGEND_DIR = ROOT_DIR / "data" / "LEGEND"
DEFAULT_FILE = "LEGEND jd.xlsm"


class LegendJDService:
    def __init__(self, base_dir: Path | None = None):
        self.base_dir = base_dir or LEGEND_DIR

    def _resolve_excel(self) -> Path:
        preferred = self.base_dir / DEFAULT_FILE
        if preferred.exists():
            return preferred
        # fallback: primer .xlsm
        candidates = list(self.base_dir.glob("*.xlsm"))
        if candidates:
            return candidates[0]
        raise FileNotFoundError("No se encontró Excel en data/LEGEND/")

    def load_all(self) -> Dict[str, object]:
        excel_path = self._resolve_excel()
        wb = _load_wb(excel_path)
        cfg: LegendConfig = read_info_config(wb)
        equipos = read_equipos(wb)
        usos = read_usos(wb)
        variadores = read_variadores_fc102(wb)
        wcr = read_wcr(wb)
        plantillas: Dict[str, List[PlantillaBOMItem]] = read_plantillas_bom(wb)

        return {
            "config": cfg,
            "equipos": equipos,
            "usos": usos,
            "variadores": variadores,
            "wcr": wcr,
            "plantillas": plantillas,
        }


if __name__ == "__main__":
    svc = LegendJDService()
    data = svc.load_all()
    cfg = data["config"]
    print(f"Proyecto: {cfg.proyecto} | Ciudad: {cfg.ciudad} | Refrigerante: {cfg.refrigerante} | Tcond: {cfg.tcond}")
    print(f"Equipos: {len(data['equipos'])} | Usos BT: {len(data['usos'].get('BT', []))} | Usos MT: {len(data['usos'].get('MT', []))}")
    print(f"Variadores: {len(data['variadores'])} | WCR: {len(data['wcr'])}")
    print({k: len(v) for k, v in data["plantillas"].items()})
