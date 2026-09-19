from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path


class ValidationError(ValueError):
    """An actionable error in a workbook, batch, or destination."""


@dataclass(frozen=True)
class MaterialRow:
    row_number: int
    material_id: str
    description: str
    unit: str
    opening: Decimal
    received: Decimal
    used: Decimal
    closing: Decimal


@dataclass(frozen=True)
class Report:
    source: Path
    month: date
    location: str
    rows: tuple[MaterialRow, ...]
