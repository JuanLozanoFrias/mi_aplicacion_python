from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional


@dataclass
class LegendConfig:
    proyecto: Optional[str] = None
    ciudad: Optional[str] = None
    tipo_sistema: Optional[str] = None
    refrigerante: Optional[str] = None
    tcond: Optional[float] = None
    tevap_bt: Optional[float] = None
    tevap_mt: Optional[float] = None
    marca_compresores: Optional[str] = None
    tipo_instalacion: Optional[str] = None
    factor_seguridad: Optional[float] = None
    extras: Dict[str, Any] = field(default_factory=dict)


@dataclass
class EquipoCatalogItem:
    equipo: str
    btu_hr_ft: float


@dataclass
class VariadorItem:
    modelo: str
    potencia: float


@dataclass
class WcrItem:
    modelo: str
    capacidad: float


@dataclass
class PlantillaBOMItem:
    qty: Optional[float]
    descripcion: str
