from __future__ import annotations

from datetime import datetime
from typing import List

from .models import ProductionOrder, Technician, Task, StageLog, MaterialItem
from .material_suggester import suggest_materials


STAGES = [
    "Recepcion OP",
    "Personal",
    "Tareas",
]


class DemoStore:
    def __init__(self, orders: List[ProductionOrder], techs: List[Technician]):
        self.orders = orders
        self.techs = techs
        self.tasks: List[Task] = []

    def get_order(self, op_number: str) -> ProductionOrder | None:
        for o in self.orders:
            if o.op_number == op_number:
                return o
        return None

    def add_order(self, order: ProductionOrder) -> None:
        self.orders.append(order)

    def assign_techs(self, order: ProductionOrder, tech_names: List[str]) -> None:
        order.technicians = tech_names
        for name in tech_names:
            for t in self.techs:
                if t.name == name:
                    t.status = "ASIGNADO"
                    t.op = order.op_number
        self._create_tasks_for_order(order)

    def advance_stage(self, order: ProductionOrder, stage_index: int) -> None:
        if stage_index < 0 or stage_index >= len(STAGES):
            return
        order.stage_index = stage_index
        order.avance_pct = (stage_index / (len(STAGES) - 1)) * 100.0
        order.logs.append(StageLog(datetime.now(), STAGES[stage_index]))

    def generate_materials(self, order: ProductionOrder) -> List[MaterialItem]:
        order.materials = suggest_materials(order.categories)
        self.advance_stage(order, max(order.stage_index, 1))
        return order.materials

    def _create_tasks_for_order(self, order: ProductionOrder) -> None:
        cats = [c.strip().upper() for c in order.categories.split(",") if c.strip()]
        for cat in cats:
            self.tasks.append(Task(
                title=f"TAREA {cat}",
                op_number=order.op_number,
                status="PENDIENTE",
                priority=order.priority,
                assigned_to="",
            ))
