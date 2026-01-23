from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import List


@dataclass
class StageLog:
    timestamp: datetime
    stage_name: str
    note: str = ""


@dataclass
class MaterialItem:
    item: str
    code: str
    name: str
    description: str
    qty: float
    category: str
    status: str = "PENDIENTE"


@dataclass
class Task:
    title: str
    op_number: str
    status: str = "PENDIENTE"
    priority: str = "MEDIA"
    assigned_to: str = ""


@dataclass
class Technician:
    number: str
    name: str
    cedula: str
    shift: str
    area: str
    op: str
    activity: str
    notes: str
    status: str = "DISPONIBLE"
    productivity: float = 0.0


@dataclass
class ProductionOrder:
    op_number: str
    date: str
    state: str
    ref1: str
    ref2: str
    notes: str
    categories: str
    avance_pct: float = 0.0
    priority: str = "MEDIA"
    technicians: List[str] = field(default_factory=list)
    stage_index: int = 0
    materials: List[MaterialItem] = field(default_factory=list)
    material_files: List[str] = field(default_factory=list)
    task_checks: List[str] = field(default_factory=list)
    task_assignments: dict[str, str] = field(default_factory=dict)
    logs: List[StageLog] = field(default_factory=list)
