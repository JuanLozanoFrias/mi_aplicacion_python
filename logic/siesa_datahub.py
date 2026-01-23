from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

from data.siesa.datahub import CompanyDataHub, PackageError

_DEFAULT_DIR = Path(__file__).resolve().parents[1] / "data" / "siesa" / "company_data"
_current_dir: Path = _DEFAULT_DIR
_hub: CompanyDataHub | None = None


def set_company_data_dir(path: Path) -> CompanyDataHub:
    global _current_dir, _hub
    _current_dir = Path(path)
    _hub = CompanyDataHub(str(_current_dir))
    return _hub


def get_company_data_dir() -> Path:
    return _current_dir


def get_hub() -> CompanyDataHub:
    global _hub
    if _hub is None:
        _hub = CompanyDataHub(str(_current_dir))
    return _hub


def load_hub(verify_hashes: bool = True) -> CompanyDataHub:
    hub = get_hub()
    hub.load(verify_hashes=verify_hashes)
    return hub


def get_assets() -> List[Dict[str, Any]]:
    return get_hub().get_assets()


def get_locations() -> List[Dict[str, Any]]:
    return get_hub().get_locations()


def get_rules() -> Dict[str, Any]:
    return get_hub().get_rules()


def get_inventory(company: str = "Weston") -> List[Dict[str, Any]]:
    return get_hub().get_inventory(company)


def get_production_orders(company: str = "Weston") -> List[Dict[str, Any]]:
    return get_hub().get_production_orders(company)


def inventory_index(company: str = "Weston") -> Dict[str, Dict[str, Any]]:
    return get_hub().inventory_index(company)
