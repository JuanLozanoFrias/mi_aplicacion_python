# datahub.py
from __future__ import annotations

import json
import hashlib
import os
from dataclasses import dataclass
from typing import Any, Dict, List, Optional


class PackageError(RuntimeError):
    pass


def _sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _read_json(path: str) -> Any:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


@dataclass(frozen=True)
class ManifestFile:
    path: str
    sha256: str
    kind: str  # master | snapshot | rules | audit
    company: Optional[str] = None


@dataclass(frozen=True)
class Manifest:
    package_name: str
    schema_version: str
    generated_at: str
    files: List[ManifestFile]


class CompanyDataHub:
    """
    Carga y valida un 'paquete JSON' (company_data/) y expone getters simples para Calvo.

    Estructura esperada:
      company_data/
        manifests/manifest.json
        master_data/*.json
        snapshots/*.json
        rules/*.json
        audit/*.jsonl (opcional)

    - Verifica hashes SHA256 para detectar archivos corruptos o modificados.
    - No requiere librerías externas.
    """

    def __init__(self, root_dir: str):
        self.root_dir = os.path.abspath(root_dir)
        self.company_dir = self.root_dir
        self._manifest: Optional[Manifest] = None
        self._cache: Dict[str, Any] = {}

    def load(self, verify_hashes: bool = True) -> None:
        man_path = os.path.join(self.company_dir, "manifests", "manifest.json")
        if not os.path.exists(man_path):
            raise PackageError(f"No existe manifest.json en: {man_path}")

        raw = _read_json(man_path)
        files: List[ManifestFile] = []
        for f in raw.get("files", []):
            files.append(ManifestFile(
                path=str(f["path"]),
                sha256=str(f["sha256"]),
                kind=str(f.get("kind", "unknown")),
                company=f.get("company")
            ))

        self._manifest = Manifest(
            package_name=str(raw.get("package_name", "unknown")),
            schema_version=str(raw.get("schema_version", "unknown")),
            generated_at=str(raw.get("generated_at", "")),
            files=files
        )

        if verify_hashes:
            self.verify_hashes()

        self._cache.clear()

    @property
    def manifest(self) -> Manifest:
        if self._manifest is None:
            raise PackageError("Manifest no cargado. Llama hub.load() primero.")
        return self._manifest

    def verify_hashes(self) -> None:
        man = self.manifest
        for f in man.files:
            fp = os.path.join(self.company_dir, f.path)
            if not os.path.exists(fp):
                raise PackageError(f"Falta archivo declarado en manifest: {f.path}")
            got = _sha256_file(fp)
            if got.lower() != f.sha256.lower():
                raise PackageError(f"Hash inválido para {f.path}. Esperado={f.sha256} Got={got}")

    def _get(self, rel_path: str) -> Any:
        if rel_path in self._cache:
            return self._cache[rel_path]
        fp = os.path.join(self.company_dir, rel_path)
        data = _read_json(fp)
        self._cache[rel_path] = data
        return data

    # ----------- Getters comunes -----------
    def get_assets(self) -> List[Dict[str, Any]]:
        return self._get("master_data/assets.json")

    def get_locations(self) -> List[Dict[str, Any]]:
        return self._get("master_data/locations.json")

    def get_rules(self) -> Dict[str, Any]:
        return self._get("rules/materiales_rules.json")

    def get_inventory(self, company: str = "Weston") -> List[Dict[str, Any]]:
        filename = f"snapshots/inventory_{company}.json"
        data = self._get(filename)
        if isinstance(data, dict) and 'rows' in data:
            return data.get('rows', [])
        return data

    def get_production_orders(self, company: str = "Weston") -> List[Dict[str, Any]]:
        filename = f"snapshots/production_orders_{company}.json"
        data = self._get(filename)
        if isinstance(data, dict) and "rows" in data:
            return data.get("rows", [])
        return data

    # ----------- Utilidades -----------
    def find_asset(self, asset_id: str) -> Optional[Dict[str, Any]]:
        for a in self.get_assets():
            if str(a.get("asset_id")) == asset_id:
                return a
        return None

    def inventory_index(self, company: str = "Weston") -> Dict[str, Dict[str, Any]]:
        """
        Índice por Referencia para búsquedas rápidas.
        """
        out: Dict[str, Dict[str, Any]] = {}
        for row in self.get_inventory(company):
            ref = str(row.get("Referencia", "")).strip()
            if ref:
                out[ref] = row
        return out
