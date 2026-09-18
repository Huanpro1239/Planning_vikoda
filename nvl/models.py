"""NVL data models."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

@dataclass
class NVLConfig:
    source_name: str
    source_path: str
    source_sheet: str
    source_code_col: int
    source_code_col_letter: str
    source_value_col: int
    source_value_col_letter: str
    source_start_row: int
    source_sourcedoc: str

    target_name: str
    target_path: str
    target_sheet: str
    target_code_col: int
    target_code_col_letter: str
    target_value_col: int
    target_value_col_letter: str
    target_start_row: int
    target_sourcedoc: str

    direct_copy: bool = True
    allow_zero: bool = True
    allow_negative: bool = True
    preserve_missing_in_source: bool = True
    reject_duplicates: bool = True
    reject_formula_in_target_cell: bool = True
    number_convention: str = "strict"


@dataclass
@dataclass
class NVLReconcileResult:
    changes: list[dict[str, Any]] = field(default_factory=list)
    unchanged: list[dict[str, Any]] = field(default_factory=list)
    missing_in_source: list[dict[str, Any]] = field(default_factory=list)
    source_only: list[str] = field(default_factory=list)
    target_codes: dict[str, int] = field(default_factory=dict)
    source_codes: dict[str, int] = field(default_factory=dict)
    status: str = "success"
    message: str = ""

