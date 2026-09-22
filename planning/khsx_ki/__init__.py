"""Canonical KHSX_ki API.

Implementation is split by responsibility while preserving the stable API used
by Planning pipeline, weekly verification and regression tests.
"""

from .calendar import compute_month_weeks, compute_standard_calendar_weeks
from .layout import (
    _detect_current_layout,
    get_layout_spec,
    parse_week_columns,
)
from .verification import verify_khsx_ki
from .workbook import (
    EPS,
    KHSX_KI_SHEET,
    PLANNING_SHEET,
    has_khsx_ki_sheet,
    patch_khsx_ki_workbook,
    reconcile_skus,
)

__all__ = [
    "EPS",
    "KHSX_KI_SHEET",
    "PLANNING_SHEET",
    "_detect_current_layout",
    "compute_month_weeks",
    "compute_standard_calendar_weeks",
    "get_layout_spec",
    "has_khsx_ki_sheet",
    "parse_week_columns",
    "patch_khsx_ki_workbook",
    "reconcile_skus",
    "verify_khsx_ki",
]
