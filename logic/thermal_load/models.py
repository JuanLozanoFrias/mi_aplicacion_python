from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class ValidationIssue:
    level: str  # warning | error
    message: str

